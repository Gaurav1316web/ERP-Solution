Imports common
Public Class clsUserMgtCode
    Public Const frmNIRQCvsWetQc As String = "NIRQ_VS_W"
    Public Const frmUnionWiseDockEntryReport As String = "UNN-DE-RPT"
    Public Const FrmInActiveDCSReport As String = "UNN-IA-DC"

    Public Const frmSubRouteMaster As String = "SUB-RT-MSTR"
    Public Const frmAccountSalesReport As String = "ACC-SLS-RPT"
    Public Const frmPIPaymentStatusReport As String = "PIP-STS-RPT"
    Public Const frmEmployeeOTEntry As String = "EMP-OT"
    Public Const rptDCSTruckSheetRegister As String = "DCS_REG"
    Public Const frmTruckWiseRejectionReport As String = "TRW-RJN-RPT"
    Public Const frmBatchManufacturingMaster As String = "Bat-MAN-MA"
    Public Const rptDbtStatusRCDF As String = "DBT-ST-RCD"
    Public Const rptMachineSurveyRegister As String = "MH-SRV-RGSR"
    Public Const frmDCSDeductionReport As String = "DCS-DED-RPT"
    Public Const rptTankerMilkQcVsERPEntry As String = "TAN-ERP"
    Public Const rptUnionWiseMilkTankerCollectionDetail As String = "UNI-MI-TANK"
    Public Const rptDBTStatusReport As String = "RPT-DBT-STS"
    Public Const frmRALQCReport As String = "RAL-QC-RPT"
    Public Const frmUnionWiseAppUserReport As String = "UNN-APP-RPT"
    Public Const FrmCmSambalYojanaSubsidy As String = "CM-SAM-YO"
    Public Const rptDBTBankResponse As String = "DBT-BAN-RE"
    Public Const ZoneWiseReport As String = "ZONE_Q_R"
    Public Const frmRouteWiseSaleTargetReport As String = "RSL-TGT-RPT"
    Public Const frmRouteWiseSaleTarget As String = "RTWS-SL-TGT"
    Public Const frmPDAccountSummariesReport As String = "PD-ACC-SRPT"
    Public Const frmCommissionTPTCalculation As String = "CMN-TPT-CLN"
    Public Const frmSendBilltoDistributor As String = "SND-BLL-DST"
    Public Const rptSchemeSaleReport As String = "SHM-SLE-RPT"
    Public Const frmIBPReport As String = "IBP-RPT"
    Public Const frmMonthWiseBeneficieriesReport As String = "MNTH-BNFRCS"
    Public Const rptBMCGazeCapacityUtilizationReport As String = "BMCGAZE-UT"
    Public Const frmTTSavingReport As String = "SVNG-BLNCE"
    Public Const rptBmcTankerProfitLossReport As String = "BMC-TAN-G-L"
    Public Const frmDCSWiseHoldPayments As String = "DCS-WS-PYNT"
    Public Const rptBoothWiseBillReport As String = "BOOTH-WI-BI"
    Public Const rptFarmerSaleRegister As String = "FAR-SAL-REG"
    ' Public Const DCSDeductionBalanceReport As String = "DCS_D_B_R"

    Public Const frmGatepassDetailReport As String = "GTP-DTL-RPT"
    Public Const frmChangePasswordScreen As String = "CHNG-PASS"
    Public Const rptEmployeeDeductionMasterReport As String = "EMP-DED-RPT"
    Public Const rptTentativeReport As String = "TEN-SAL-RPT"
    Public Const rptEmployeeOTReport As String = "EMP-OT-RPT"
    Public Const frmRCDFRateControl As String = "RCDF-RT-CNT"
    Public Const frmOfficeOrder As String = "OFFCE-ORD"
    Public Const frmEmployeeDeductionMaster As String = "EMP-DED-MST"
    'Public Const FrmReceiptInvoiceMapping As String = "REC-INV-MAP"
    Public Const frmActiveInactiveEmpRegister As String = "ACT-INACT"
    Public Const frmPOBulkProc As String = "PO-BULK-P"
    Public Const frmDBTNEFTUnionReport As String = "UNION-RPT"
    Public Const DBTPDAccountReport As String = "DBT-PDA-RPT"
    Public Const rptDBTNEFTPaymentDetailReport As String = "DBT-PAY-DET"
    Public Const rptExeVersionReport As String = "EXE-VER"
    Public Const rptProductionReport As String = "PRO-RPT"
    Public Const rptProductionStoreReport As String = "PRO-SAL-RPT"
    Public Const frmDCSSavingLedger As String = "DCS-LGR-RPT"
    Public Const frmDBTStatusAndLastDPTStatus As String = "DBT-STATUS"
    Public Const frmDBTApprovalStatus As String = "DBT-APPRO"
    Public Const frmGateEntryReturnCS As String = "GERETURN-CS"
    Public Const frmGateEntryReturnTransfer As String = "GERET-TRANSF"
    Public Const frmGateEntryReturnPS As String = "GERETURN-PS"
    Public Const ExpertERP As String = "ExpertERP"
    Public Const frmRawMilkConsumtion As String = "RMConsump"
    Public Const FrmJobWorkInventory As String = "JobWOInvent"
    Public Const FrmShortCloseDOPS As String = "SHORT-CDO-PS"
    Public Const FrmShortCloseDOCS As String = "SHORT-CDO-CS"
    Public Const ModuleFavourite As String = "MFavourite"
    Public Const ModuleSystemAdmin As String = "MSysAdmin"
    Public Const ModuleCommonServices As String = "MCommSer"
    Public Const ModuleReceivable As String = "MReceivable"
    Public Const MPMilkPrice As String = "PCU-MCC"
    Public Const frmUnpostBmcDcs As String = "Unpo-BM-DC"
    Public Const frmDCSTransportationCharges As String = "DCS-TPT-CH"
    Public Const frmManualIoTFarmerCollection As String = "MN-IoT-F-CL"
    Public Const FrmBMCCategory As String = "BMC-CATY"
    Public Const ModulePayable As String = "MPayable"
    Public Const ModuleGL As String = "MGenLedger"
    Public Const ModuleSales As String = "MSales"
    Public Const ModuleSalesNew As String = "MSalesNew"
    Public Const ModuleMaterial As String = "MMaterial"
    Public Const ModulePurchase As String = "MPurchase"
    Public Const ModuleTDS As String = "MTDS"
    Public Const ModuleFixedAsset As String = "MFixedAsset"
    Public Const frmComplaintEntry As String = "Compl_Ety"
    Public Const ModuleHR As String = "MHR"
    Public Const ModuleReco As String = "MReco"
    Public Const SubModuleRecoReport As String = "SMRecoRpt"
    Public Const ModuleAccSet As String = "MAccSet"
    Public Const SubModuleAccSetReport As String = "SMAccSetRpt"
    Public Const FrmBulkRoutMaster As String = "BROUT_MAST"
    Public Const TankerCleaningItem As String = "TAN-CLN-ITM"
    Public Const frmShiftMasterBulk As String = "SH-MST-BULK"
    Public Const ModuleParavet As String = "MPVS"
    Public Const ModuleFarmerPayment As String = "MFPayment"
    Public Const frmProductionCancelledTransation As String = "PRO-CAN-RPT"
    '' add sub modules Master,Transaction and Report in Farmer Payment

    Public Const SubModuleFarmerPaymentMaster As String = "MFPayMaster"

    'Public Const RptRouteWiseTrendBar As String = "RPT_Route"
    Public Const rptMccCollectionDetails As String = "RPT_MCC_COL"
    Public Const rptMccMasterDetail As String = "RPT_MCC_DET"
    Public Const rptListofCowDCS As String = "COW-LST-RPT"
    Public Const rptMccProcurementUploader As String = "RPT_MCC_UP"
    Public Const rptMCCWiseAbstractReport As String = "RPT_MCC_ABST"
    Public Const rptCollectionDataChangeReport As String = "RPT_MCC_CDC"
    'Public Const rptFarmerPaymentApprovalReport As String = "RPT_MCC_FPA"
    Public Const rptMCCPaymentSummary As String = "RPT_MCC_PS"
    Public Const FrmItemCostUOM As String = "Cost-UOM"


    '' add masters
    Public Const FrmFarmerPaymMCCFarmerMapping As String = "FPMCCFMapp"

    Public Const SubModuleFarmerPaymentTransaction As String = "MFPayTrans"
    Public Const SubModuleFarmerPaymentReport As String = "MFPayReport"


    '' add trasaction
    Public Const frmFarmerPaymentAdjustment As String = "F-Pay-Adj"
    Public Const frmMCCFarmerMappingFP As String = "F-MCC-FMap"
    Public Const MPBillGeneration As String = "MCC-MP-B"
    Public Const frmLockMPCollectionPC As String = "MPC-LOCK"
    Public Const frmFarmerPaymentEntry As String = "F-PAY-ENT"
    Public Const frmFarmerMilkPurchaseInvoice As String = "F-M-P-INV"
    Public Const frmItemTaxRate As String = "ITR"

    '' add Setup JobWork outword 
    Public Const frmJobWorkoutwordMaster As String = "JWOutward"
    Public Const frmVendorItemChargeMaster As String = "JWVITEMC"

    Public Const JWOParameterMaster As String = "JWO-PMM"
    Public Const JWOParameterFormula As String = "JWO-PFM"
    Public Const JWOFormulaVendorMapping As String = "JWO-FVM"



    Public Const RptSubsidyCredit As String = "SubSidy_Rpt"
    Public Const RptSchemeDetail As String = "Scheme_D_Rpt"
    Public Const rptLeakageDetail As String = "Leakage_Rpt"

    ''Transactions JobWork Outward
    Public Const frmMilkJobWorkTransfer As String = "MILK-TRA-JB"
    Public Const frmMilkJobWorkTransferReturn As String = "JWO-TRAM-RET"
    Public Const frmMilkJobWorkTransferOther As String = "MILK-TRA-JBO"
    Public Const frmMilkJobWorkTransferOtherReturn As String = "JWO-TRAO-RET"
    Public Const frmGateEntry_JWO As String = "GATE-EN-JWO"
    Public Const JWO_Item_Acceptance As String = "JWO-ITM-ACC"
    Public Const frmWeighment_JWO As String = "WEIGH-JWO"
    Public Const frmQC_JWO As String = "QC-JWO"
    Public Const frmUnloading_JWO As String = "UNLOAD-JWO"
    Public Const JWO_Estimation As String = "JWO-EST"
    Public Const JWO_SRN As String = "JWO-SRN"

    Public Const JWO_SRN_Return As String = "JWO-SRN-RET"


    ' add Setup Jobwork InWard 
    Public Const frmJWParameterMaster As String = "JWINPM"
    Public Const frmJWFormulaMaster As String = "JWInWFRMst"
    Public Const frmJWVendorFormula As String = "JWVenFom"
    Public Const JWIItemPriceMaster As String = "JWI-ITM-PRI"

    Public Const FrmSRNJobWorkEstimate As String = "JWSRNESTM"
    Public Const frmJobWorkBillig As String = "JobWorkBill"
    Public Const rptJobworkMilkReceipt As String = "RPT-M-RCPT"
    Public Const rptJobworkBilling As String = "RPT-JW-BILL"
    Public Const rptJobworkProductionReport As String = "RPT-JW-PROD"
    Public Const rptJWTankerStatusReport As String = "RPT-JW-TAKS"
    Public Const rptJWDispatchReport As String = "RPT-JW-DISP"
    '' add reports in farmer Payment Report
    Public Const frmFarmerLedgerReport As String = "Ledger-F"
    Public Const rptVspFarmerPaymentDetail As String = "VSP-FRM-PAY"
    Public Const rptFarmerNotMappedWithVSP As String = "FNM-WITH-VSP"

    Public Const SubModuleParavetMaster As String = "MPVSMaster"
    Public Const SubModuleParavetReport As String = "MPVSReport"
    Public Const SubModuleParavetTrans As String = "MPVSTrans"
    Public Const frmCattleType As String = "CattleType"
    Public Const frmBredType As String = "BredType"
    Public Const frmCattleColor As String = "CattleColor"
    Public Const frmCattleMaster As String = "CattleMaster"
    Public Const FrmServiceGroup As String = "SERVGROP"
    Public Const FrmServiceName As String = "SERVNAME"
    Public Const FrmServiceMaster As String = "SERVMST"
    Public Const FrmNDDBMaster As String = "NDDBMST"
    Public Const FrmBullMaster As String = "BULLMST"
    Public Const FrmFarmerServiceOrderWithRate As String = "FARSRVTRN"
    Public Const rptSMSDetails As String = "RPTSMSDET"
    Public Const rptActiveUsers As String = "Active_User"

    'Public Const ModuleProduction As String = "MProduction"
    Public Const ModuleProductionSTD As String = "MProdSTD"
    Public Const ModuleProductionDairy As String = "MProdDairy"
    Public Const ModuleProductionPepsi As String = "MProdPepsi"
    Public Const ModuleMISReports As String = "MMISReports"
    Public Const SubModuleMISReports As String = "SMISReports"
    Public Const ModuleVPF As String = "MVPF"
    Public Const FrmRptAPInvoiceDetailsReport As String = "APINDetRPT"
    Public Const ModuleBI As String = "ModuleBI"
    Public Const ModuleUtility As String = "MUtility"
    Public Const ModuleProjectManagement As String = "MPROJMGT"
    Public Const ModuleService As String = "MService"
    Public Const ModuleMCCMilkProcurement As String = "MMMProc"
    'Public Const ModuleBulkMilkProcurement As String = "MMBProc"
    Public Const ModuleMilkProcurement As String = "MMProc"
    Public Const ModulesalePurchaseSecurity As String = "MMSPS"
    Public Const ModuleJobWorkOutWard As String = "MJWOut"
    Public Const ModuleJobWorkInWard As String = "MJWIn"
    '
    Public Const ModuleElectrical As String = "MELCL"
    Public Const SubModuleSetupElectrical As String = "SMSMELCL"
    Public Const SubModuleTransactionElectrical As String = "SMTMELCL"
    Public Const SubModuleReportElectrical As String = "SMRMELCL"

    Public Const ModuleEngPlantMgt As String = "MEPM"
    Public Const SubModuleSetupEngPlantMgt As String = "SMSEPM"
    Public Const SubModuleTransactionEngPlantMgt As String = "SMTEPM"
    Public Const SubModuleReportEngPlantMgt As String = "SMREPM"

    Public Const ModuleComplaint As String = "MCOMP"
    Public Const SubModuleSetupComplaint As String = "SMSCOMP"
    Public Const SubModuleTransactionComplaint As String = "SMTCOMP"
    ''=================== Dairy Sales start here

    Public Const ModuleSaleDairy As String = "MSaleDairy"
    Public Const ModuleProductSaleDairy As String = "MPSaleDairy"
    Public Const ModuleAPSSaleDairy As String = "MASaleDairy"
    Public Const SubModuleSaleDairySetUp As String = "SMSDairyM"
    Public Const SubModuleProductSaleDairySetUp As String = "SMSDairyPM"
    Public Const SubModuleAPSSaleDairySetUp As String = "SMSDairyAM"
    Public Const SubModuleProductSaleDairyTransaction As String = "SMSDairyPT"
    Public Const SubModuleAPSSaleDairyTransaction As String = "SMSDairyAT"
    Public Const SubModuleProductSaleDairyReport As String = "SMSDairyPR"
    Public Const SubModuleAPSSaleDairyReport As String = "SMSDairyAR"


    ' ADDED BY KUNAL TICKET : BM00000009674
    Public Const FrmSaleSettingFreshDS As String = "SSFMst"
    Public Const FrmZoneMasterDS As String = "ZMMst"
    Public Const frmRouteMasterDS As String = "RMMst"
    Public Const FrmFreshTransactionApprovalDS As String = "FTAMst"
    Public Const FrmSchemeMasterDairyDS As String = "SMDMst"
    Public Const frmRouteFreightDetailsDS As String = "SFDMst"
    Public Const SaleIncentiveMaster As String = "SaleINCMst"
    Public Const CustomerDeduction As String = "CUST-DED"

    'Public Const frmCustCategoryWiseDefaultItemUomMaster As String = "CUST-CAT-UOM"
    'Public Const frmTranspoterDeductionMaster As String = "TRAN-DED_MST"
    'Public Const frmItemSublocationMapping As String = "ITM_SLOC_MP"

    '--Dairy Sales Trnasaction
    Public Const RptRouteWiseSaleRegister As String = "RPT_RTE_SLE"
    Public Const SaleEinvoiceReport As String = "SA_EIV_RPT"
    Public Const rptSaleInvoiceStatusReport As String = "SA_INV_RPT"
    Public Const SubModuleSaleDairyTransaction As String = "SMSDairyT"
    Public Const frmbookingdairy As String = "BOOK-DS"
    'Public Const frmbookingdairyFreshSale As String = "BOOK-DS_FSH"
    Public Const frmGenerateBookingFreshSale As String = "GEN_BOO_FS"
    Public Const frmCustomerTender As String = "CUST-TEND"
    Public Const FrmSalesOrder As String = "SALES-ORD"
    Public Const FrmSalesOrderDispatch As String = "SAL-ORD_DIS"
    Public Const FrmSalesOrderGatePass As String = "SAL-ORD-GP"
    Public Const FrmAPSSaleReport As String = "APS-SAL-RPT"
    'Public Const frmAdvanceForCD As String = "ADV_For_CD"
    'Public Const frmAcknowledgeMentOfGRN As String = "ACK_OF_GRN"
    Public Const frmDairyBookingUploader As String = "DA_BOK_UPL"
    Public Const frmCashIndentBookingMobApp As String = "MA_CIBS"
    Public Const POSDairySale As String = "POS-DS"
    Public Const frmSaleOrderDairy As String = "SALEORD-DS"
    Public Const frmDeliveryOrderDairy As String = "DELORD-DS"
    Public Const frmSaleDispatchDairy As String = "DISPATCH-DS"
    Public Const frmSaleInvoicedairy As String = "INVOICE-DS"
    Public Const frmSaleReturndairy As String = "RETRUN-DS"
    Public Const frmDispatchReturn As String = "DIS-RETURN"
    Public Const frmGatePassDairy As String = "GATEPASS-DS"
    Public Const frmCrateReceviedDairySale As String = "CRT-REC-DS"
    Public Const frmCrateReceviedBox As String = "CRT-REC-BOX"
    Public Const frmBookingDairyMultipleDistributor As String = "BOOK-DS-DIS"
    Public Const frmDairyBookingCustomer As String = "BOOK-DS-CU"
    Public Const frmPerformaInvoiceDairy As String = "Prof_Inv-D"
    Public Const frmDairyFreshDispatchMultiple As String = "FS_DIS-M"
    Public Const frmMultipleInvoice As String = "FS_Inv-M"
    Public Const frmCustomerPenalty As String = "CUST-PNLTY"
    'Public Const frmItemProjection As String = "ITEM_PROJTN"
    '===========================Added by preeti Gupta[14/03/2018]==================
    Public Const frmsaleReturnGateEntryFS As String = "SR-GE-FS"
    Public Const frmsaleReturnGateEntryPS As String = "SR-GE-PS"
    Public Const frmsaleReturnGateEntryMISSAle As String = "SR-GE-MISS"
    'Public Const frmsaleReturnGateEntryMCCSAle As String = "SR-GE-MCCS"
    Public Const frmsaleReturnGateEntryBulkSAle As String = "SR-GE-BULS"
    Public Const frmsaleReturnGateEntryExportSAle As String = "SR-GE-EXPS"
    Public Const frmsaleReturnGateEntryCSATransfer As String = "SR-GE-CSAT"
    '================================================================================
    '===Dairy Sales reports
    Public Const SubModuleSaleDairyReport As String = "SMSDairyR"
    Public Const RptEffectiveRateReport1 As String = "EFF_RATE"
    Public Const rptPlantCustomerDemand As String = "PLT_CUS_DEMD"
    Public Const rptCustomerEffective_ItemRate As String = "RPT-CUS-RATE"
    Public Const RptDeliveryOrderReport1 As String = "RPT-Del-Ord"
    Public Const rptDemandForsingleBranch As String = "DMD_SNG_BRA"
    Public Const rptCrateJalliReportForTransfer As String = "TRA_CRT_JALL"
    Public Const FrmPrintDistributerInvoiceStatement As String = "DSTR_INV_SMT"
    Public Const frmDistributerLedgerReport As String = "DSTR_Ldgr"
    Public Const RptDairyBookingDistributorReport As String = "RPT_DBDR"
    Public Const rptMatrixFreshSalesReportSaleDairy As String = "RPT_MAT_SAL"
    Public Const rptDairyTruckSheetReport As String = "RPT_TRK_SHT"
    Public Const rptDayWiseBoothDemand As String = "DYA_WIS_RPT"
    Public Const rptSaleRegisterDetail As String = "RPT_SR_DT"
    Public Const rptSaleRegisterUOMReport As String = "RPT_UOM_SR"
    Public Const RptCommissionTPTReport As String = "RPT_Com_TPT"
    Public Const rptDairySaleRegisterReport As String = "RPT_SRR_DT"
    Public Const InvoicePriceWiseReport As String = "RPT_INV_P"
    Public Const rptInvoiceDetailReport As String = "RPT_INV_D"
    Public Const rptDailyReceiptReport As String = "RPT_REC_D"
    Public Const rptTransporterWiseReport As String = "RPT_TRAN_D"
    Public Const rptDailyLeakageReplacementReport As String = "RPT_DLE_REP"
    Public Const FrmZoneWiseSKUReport As String = "ZW_SKU_REP"
    Public Const FrmDairySaleSchemeReport As String = "DS_SCH_REP"
    Public Const rptBookingWiseRegister As String = "DS_BW_RPT"
    Public Const rptDemandNill As String = "DM_NL_RPT"
    Public Const rptMonthlyBillSummaryReport As String = "MON_BLL_RPT"
    Public Const frmDemandHistory As String = "DS_DEM_HIS"
    Public Const frmEwaybill As String = "E_WAY_BILL"
    Public Const rptItemWiseBillReport As String = "IT_WI_BI"
    Public Const rptDistributorRouteMarginStatusReport As String = "Dis_R_ST"
    Public Const rptPartWiseItemReport As String = "PA_WI_ITE"

    Public Const rptBankDetailsWithVendorMargin As String = "RPT_BDWVM"
    Public Const rptAbsentBooth As String = "ABS_BOOTH"
    Public Const rptUnpostedDocumentReport As String = "UNPOST_DOC"
    Public Const rptCustomerWiseSalesReport As String = "CUST_SAL_RPT"
    Public Const rptNewTenderPartyListReport As String = "NEW_TPT_LST"
    Public Const rptBookingReport As String = "BOOKING_RPT"
    Public Const rptBookingQtyAmtReport As String = "DEMD_BK_RPT"
    Public Const rptSalesLedgerReport As String = "SAL_LED_RPT"
    Public Const rptSalesStock As String = "SAL_STK_RPT"
    Public Const rptSalesComparisonReport As String = "SAL_CMP_RPT"
    Public Const FrmBoothDemandReport As String = "BO_DE_RPT"
    Public Const FrmBoothCountReport As String = "BO_CU_RPT"
    Public Const rptAvgSaleDetailReport As String = "AVG_SAL_RPT"
    Public Const rptdairydispatch As String = "DRY_DIS_RPT"
    Public Const rptRouteSaleRegister As String = "ROT_SAL_REG"
    Public Const RptSalesSummaryReport As String = "SAL_SUM_RPT"
    Public Const rptGatePassReport As String = "GAT_PAS_RPT"
    Public Const rptNewSalesReport As String = "NE_SALE_RPT"
    Public Const rptDemandReturnReport As String = "DEM_RET_RPT"
    Public Const RouteWiseDemandReport As String = "RO-WI-DE-REP"
    Public Const rptDailyStatementReport As String = "DlY_SALE_RPT"
    Public Const FrmPrintMultipleGatepass As String = "MUL_PNT_GP"
    Public Const rptBoothTruckSheet As String = "Booth-T-S"
    Public Const rptDemandReport As String = "DEMAND-RE"

    Public Const rptCreditCustomerReport As String = "CRD_CST_RPT"
    Public Const rptD1D2Report As String = "D1_D2_RPT"
    Public Const rptZoneWiseReport1 As String = "ZW_RPT"
    Public Const rptMonthWiseDMRReport As String = "MW_DMR_RPT"

    '===Dairy Visual Sales reports  
    Public Const SubModuleVisualSaleReport As String = "SMVSaleR"
    Public Const rptZoneWiseMSVisual As String = "ZW_VSaleR"
    Public Const VisualReportSale As String = "VIS-RPT-SAL"
    ''===================
    Public Const ModuleQualityControl As String = "MMQC"
    Public Const SubModuleQCSetup As String = "SMQCS"
    Public Const frmPPLogSheetMaster_QC As String = "QC_PARAM_MST"
    Public Const frmVendorItemQCMapping As String = "VEN_ITM_MAP"
    Public Const frmQualityModuleParameterRangeMaster As String = "QM-P-RNG"
    Public Const FrmChildParameterRangeMasterForQC1 As String = "QM-P-RNG"
    '==trnasaction==
    Public Const SubModuleQCTransaction As String = "SMQCTRANS"
    Public Const SubModuleQCReport As String = "SMQCRPT"
    Public Const VisualRandomQC As String = "Vs-Rn-QC"
    Public Const frmQualityCheckForSRN As String = "QC_CHK_SRN"
    Public Const frmOutgoingQC As String = "OUT-QC-CHK"
    Public Const frmQualityCheckApprovalForSRN As String = "QC_APRV_SRN"
    Public Const rptPendingQCReport As String = "PND_QC_RPT"
    Public Const QualitySummaryReport As String = "QTY_SUMY_RPT"
    Public Const rptQCAnalysisReport As String = "QC_ANLS_RPT"
    Public Const rptNIRQC As String = "NIR_QC_RPT"

    'frmQualityCheckForSRN
    ''==================

    ''===============ShareModule==============
    Public Const ModuleShare As String = "MShare"
    Public Const SubModuleShareSetup As String = "SMShareS"
    Public Const ShareMaster As String = "SHA_MSTR"

    ''===========ShareTransaction=================
    Public Const SubModuleShareTransaction As String = "SMShareT"
    Public Const frmMultipleShareAllotment As String = "MU-SHA-ALT"
    Public Const frmShareAllotment As String = "SHA-ALT"

    ''============ShareReport=======================
    Public Const SubModuleShareReport As String = "SMShareR"
    Public Const rptShareRegister As String = "SHA-REG"

    ''==========BullManagementModule======
    Public Const ModuleBullManagement As String = "MSBullMang"
    Public Const SubModuleBullManagementSetup As String = "MSBullMangS"
    ''===========BullManagementTransaction=================
    Public Const SubModuleBullManagementTransaction As String = "MSBullMangT"
    Public Const frmBullBreedMaster As String = "BULL-BR-M"
    Public Const frmBullTestParameter As String = "BULL-TE-P"
    Public Const frmBullParameterGroup As String = "BULL-PAR-G"
    Public Const frmBullShedMaster As String = "BULL-SHE-M"
    Public Const frmBullShedParameterMaster As String = "BULL-S-PA-M"
    Public Const frmBullShedParameterGroup As String = "Bull-Sh-P-G"
    Public Const frmBullMovementType As String = "BULL-MO-TY"
    Public Const frmBullSpeciesMaster As String = "BULL-SPEC-M"
    Public Const frmBullCategoryMaster As String = "BULL-CATE-M"
    Public Const frmBullSubCategoryMaster As String = "BULL-SUB-CAT"
    Public Const frmBullSSCentreMaster As String = "BULL-SS-CE"
    Public Const frmBullPenIDMaster As String = "BULL-PAN-M"
    Public Const frmBullMasters As String = "BULL-Maste"
    Public Const frmBullMovement As String = "BULL-MVMNT"
    Public Const frmBullCurlingEntry As String = "BULL-CUR-ENT"
    Public Const frmBullStatus As String = "BULL-Status"
    Public Const frmBullSubStatus As String = "BULL-SUB-S"
    Public Const frmBullSourcePainting As String = "BULL-SOUR"
    Public Const frmBullRating As String = "BULL-RATING"
    Public Const frmBullSourceName As String = "BULL-SOURCE"
    Public Const frmBullPurchaseNo As String = "BUL-PURCHASE"
    Public Const frmBullInsuranceMaster As String = "BULL-INSR-M"
    Public Const frmBullInsuranceType As String = "BULL-INSR-TY"
    Public Const frmBullInsurance As String = "BULL-INSRCE"
    'Public Const frmBullSourceName As String = "BULL-SOURCE"
    'Public Const frmBullPurchaseNo As String = "BUL-PURCHASE"
    Public Const frmDiseaseMaster As String = "DIZIZ-MST"
    Public Const frmBullTestParameterEntry As String = "BLL-TST-MTR"
    Public Const frmBullVaccinationEntry As String = "BULL-VAC-EN"

    Public Const frmInsuranceTagAllocation As String = "INS-TG-ALOC"

    Public Const frmBullCMUGrouping As String = "BLL-CMU-GRP"

    Public Const frmCMUChecklistEntry As String = "BLL-CMU-CHKL"

    ''============BullManagementReport=======================
    Public Const SubModuleBullManagementReport As String = "MSBullMangR"



    ''richa 03/11/2014
    Public Const ModuleFreshSale As String = "MFreshSale"
    Public Const SubModuleFSSetup As String = "SMFreshSale"
    Public Const FrmSaleSettingFresh As String = "FS-SETT"

    Public Const ModuleBulkSale As String = "MBulkSale"
    Public Const SubModuleBSSetup As String = "SMBulkSale"
    Public Const FrmSaleSettingBulk As String = "BS-SETT"

    Public Const ModuleMerchantTradeSale As String = "MMTSale"
    Public Const SubModuleMTSetup As String = "SMMTSale"
    Public Const FrmSaleSettingMerchant As String = "MT-SETT"

    Public Const ModuleExportSale As String = "MExportSale"
    Public Const SubModuleExportSetup As String = "SMExportSale"
    Public Const FrmSaleSettingExport As String = "Export-SETT"

    Public Const ModuleCSASale As String = "MCSASale"
    Public Const SubModuleCSASetup As String = "SMCSASale"
    Public Const FrmSaleSettingCSA As String = "CSA-SETT"

    Public Const ModuleProductSale As String = "MProductSale"
    Public Const SubModuleProductSetup As String = "SMProductSle"
    Public Const FrmSaleSettingProduct As String = "Product-SETT"
    Public Const FrmTragetMaster As String = "Prd_Trg"
    ''=================
    '' Anubhooti 05-Aug-2014
    Public Const ModuleHumanResource As String = "MHumRes"
    Public Const SubModuleHumResSetup As String = "SMHumRes"
    Public Const SubModuleHumResTrans As String = "SMHumResT"
    Public Const SubModuleHumResTraining As String = "SMHRTrain"
    Public Const SubModuleHumResPerformance As String = "SMHumResPerf"
    Public Const SubModuleHumResReimburse As String = "SMHRRm"
    Public Const SubModuleHumResGrievancce As String = "SMHRGR"
    Public Const SubModuleHumResEMployeeEquipment As String = "SMHRET"
    Public Const SubModuleHumResExitManagment As String = "SMHRExitMng"

    '==shivani
    Public Const SubModuleHumResReport As String = "SMHRRPT"

    Public Const ModuleServiceAndWarranty As String = "MSerWarr"
    Public Const SubModuleServiceAndWarranty As String = "SMSerWarr"
    Public Const SubModuleTransServiceAndWarranty As String = "SMSWTrans"

    Public Const SubModuleSystemAdminSetup As String = "SMSysAdmin"
    Public Const SubModuleSystemAdminTransation As String = "SMSysTrans"
    Public Const SubModuleSystemAdminSetupReport As String = "SMSysReport"

    Public Const SubModuleCommonServicesSetup As String = "SMCSSetup"
    Public Const SubModuleCommonServicesTransaction As String = "SMCSTrans"
    Public Const SubModuleCommonServicesReport As String = "SMCSReport"

    Public Const SubModuleReceivableSetup As String = "SMRecSetup"
    Public Const SubModuleReceivableTransaction As String = "SMRecTrans"
    Public Const SubModuleReceivableReport As String = "SMRecReport"

    Public Const SubModulePayableSetup As String = "SMPaySetup"
    Public Const SubModulePayableTransaction As String = "SMPayTrans"
    Public Const SubModulePayableReport As String = "SMPayReport"

    Public Const SubModuleGLSetup As String = "SMGLSetup"
    Public Const SubModuleGLTransaction As String = "SMGLTrans"
    Public Const SubModuleGLReport As String = "SMGLReport"


    Public Const SubModuleSaleSetup As String = "SMSaleSetup"
    Public Const SubModuleSaleTransaction As String = "SMSaleTrans"
    Public Const SubModuleSaleReport As String = "SMSaleReport"
    Public Const SubModuleSaleShippingReport As String = "SMSaleShiRep"

    Public Const SubModuleSaleNewSetup As String = "SMSaleNSetup"
    Public Const SubModuleSaleNewTransaction As String = "SMSaleNTrans"
    Public Const SubModuleSaleNewReport As String = "SMSaNRept"

    Public Const SubModuleMaterialSetup As String = "SMMatSetup"
    Public Const SubModuleMaterialTransaction As String = "SMMatTrans"
    Public Const SubModuleMaterialReport As String = "SMMatReport"


    Public Const SubModulePurchaseSetup As String = "SMPurSetup"
    Public Const SubModulePurchaseTransaction As String = "SMPurTrans"
    Public Const SubModulePurchaseReport As String = "SMPurReport"
    Public Const SubModulePurchaseTransactionReport As String = "SMPurTransRe"

    Public Const SubModuleTDSSetup As String = "SMTDSSetup"
    Public Const SubModuleTDSTransaction As String = "SMTDSTrans"
    Public Const SubModuleTDSReport As String = "SMTDSReport"

    Public Const SubModuleFixedAssetSetup As String = "SMFASetup"
    Public Const SubModuleFixedAssetTransaction As String = "SMFATrans"
    Public Const SubModuleFixedAssetReport As String = "SMFAReport"

    Public Const SubModuleHRSetup As String = "SMHRSetup"
    Public Const SubModuleHRTransaction As String = "SMHRTrans"
    Public Const SubModuleHRReport As String = "SMHRReport"
    Public Const SubModuleHRMonthlyReport As String = "SMHRMonRpt"
    Public Const SubModuleHRStatutoryReport As String = "SMHRStaRpt"

    'Public Const SubModuleProductionSetup As String = "SMProSetup"
    'Public Const SubModuleProductionTransaction As String = "SMProTrans"
    'Public Const SubModuleProductionReport As String = "SMProReport"
    'Public Const SubModuleProductionSetupReport As String = "SMProSetRpt"

    Public Const SubModuleProductionSetupSTD As String = "SMPDSetupSTD"
    Public Const SubModuleProductionTransactionSTD As String = "SMPDTransSTD"
    Public Const SubModuleProductionReportSTD As String = "SMPDRptSTD"
    Public Const SubModuleProductionSetupReportSTD As String = "SMPDSPRptSTD"

    Public Const SubModuleProductionSetupDairy As String = "SMPDSetupD"
    Public Const SubModuleProductionTransactionDairy As String = "SMPDTransD"
    Public Const SubModuleProductionReportDairy As String = "SMPDReportD"

    Public Const SubModuleProductionSetupPepsi As String = "SMPDSetupPEP"
    Public Const SubModuleProductionTransactionPepsi As String = "SMPDTransPEP"
    Public Const SubModuleProductionReportPepsi As String = "SMPDRptPEP"

    Public Const SubModuleSetupJobWorkOurWard As String = "SMJWOut"
    Public Const SubModuleTransactionJobWorkOurWard As String = "SMTJWOut"
    Public Const SubModuleReportJobWorkOurWard As String = "SMRJWOut"

    Public Const SubModuleSetupJobWorkInWard As String = "SMJWIn"
    Public Const SubModuleTransactionJobWorkInWard As String = "SMTJWIn"
    Public Const SubModuleReportJobWorkInWard As String = "SMRJWIn"

    Public Const SubModuleBISetup As String = "SMBISetup"

    Public Const SubModuleBISale As String = "SMBISale"
    Public Const SubModuleBIPurchase As String = "SMBIPur"
    Public Const SubModuleBIFinance As String = "SMBIFin"
    Public Const SubModuleBIAsseet As String = "SMBIAsset"
    Public Const SubModuleBIProduction As String = "SMBIPro"
    Public Const SubModuleBIUserDefined As String = "SMBIUD"




    Public Const SubModuleUtilitySetup As String = "SMUtilityS"

    Public Const SubModuleProjectManagementSetup As String = "SMPMSETUP"
    Public Const SubModuleProjectManagementTransaction As String = "SMPMTRANS"
    Public Const SubModuleProjectManagementReport As String = "SMPMRPT"


    Public Const SubModuleServiceSetup As String = "SMSSetup"
    Public Const SubModuleServiceTransaction As String = "SMSTrans"
    Public Const SubModuleServiceReport As String = "SMSReport"

    Public Const SubModuleMCCMilkProcurementSetup As String = "SMMPROCSetup"
    Public Const SubModuleMCCMilkProcurementTransaction As String = "SMCPROCTRAN"
    Public Const SubModuleMCCMilkProcurementReport As String = "SMMPROCRPT"
    Public Const SubModuleMCCMilkProcurementBMCReport As String = "SMMPROCBCRPT"


    'Public Const SubModuleBULKMilkProcurementSetup As String = "SMBPROCSetup"
    'Public Const SubModuleMilkProcurementTransaction As String = "SMMPROCTRANS"
    'Public Const SubModuleBulkMilkProcurementReport As String = "SMBPROCRPT"

    'Public Const SubModuleShareSetup As String = "SMShareSetup"
    'Public Const SubModuleShareTransaction As String = "SMShareTRANS"
    'Public Const SubModuleShareReport As String = "SMShareRPT"

    'Public Const SubModuleSaleDairySetup As String = "SMSALEDSetup"
    Public Const SubModuleFreshSaleTransaction As String = "SMFRESHSTRAN"
    Public Const SubModuleBulkSaleTransaction As String = "SMBULKSTRAN"
    Public Const SubModuleMerchantTradeTransaction As String = "SMMRCHNTTRDE"
    Public Const SubModuleProductSaleTransaction As String = "SMPRODSTRAN"
    Public Const SubModuleExportSaleTransaction As String = "SMEXPORTTRAN"
    Public Const SubModuleCSASaleTransaction As String = "SMCSATRAN"
    'Public Const ItemLocationMapping As String = "ILMAPPING"



    Public Const SubModuleFreshSaleReport As String = "SMFRESHSRPT"
    Public Const SubModuleBulkSaleReport As String = "SMBULKSRPT"
    Public Const SubModuleMerchantReport As String = "SMMRCHNTRPT"
    Public Const SubModuleProductSaleReport As String = "SMPRODSRPT"
    Public Const SubModuleExportSaleReport As String = "SMEXPORTRPT"
    Public Const SubModuleCSAReport As String = "SMCSARPT"



    Public Const SubSalePurchaseSecuritySetup As String = "SPSecurity"
    Public Const SubSalePurchaseSecurityReport As String = "SPSecRpt"
    ' ------- Visual Process Flow --------- '
    Public Const ModuleVPFlow As String = "MVPFlow"
    Public Const SubModuleVPFlowSetup As String = "SMVPF"
    Public Const FrmVPFActiveReport As String = "VPF_A_R"
    Public Const FrmVPFSettings As String = "VPF_Sett"
    Public Const FrmCommonCycle As String = "Comm_VPF"
    Public Const FrmGLCycle As String = "GL_VPF"
    Public Const FrmAPCycle As String = "AP_VPF"
    Public Const FrmARCycle As String = "AR_VPF"
    Public Const FrmBulkSaleCycle As String = "BulkS_VPF"
    Public Const FrmFreshSaleCycle As String = "FreshS_VPF"
    Public Const FrmProductSaleCycle As String = "ProtS_VPF"
    Public Const FrmCSASaleCycle As String = "CSAS_VPF"
    Public Const FrmMMCycle As String = "MM_VPF"
    Public Const FrmMCCProcurementCycle As String = "MCCP_VPF"
    Public Const FrmBulkProcurementCycle As String = "BulkP_VPF"
    Public Const FrmPurchaseCycle As String = "VPF_Pur"
    Public Const FrmDProductionCycle As String = "VPF_DProd"
    Public Const FrmSupplierMaster As String = "SUPPLIER-M"
    Public Const frmDivertedContractor As String = "DIVERT-CON-M"
    Public Const frmMilkTypeMast As String = "MILK-TYPE-M"
    Public Const frmMilkGradeMaster As String = "MILK-GRADE-M"
    Public Const frmIntimation As String = "INTIMATION"
    '------------Administrative Services-----------------------
    Public Const EmployeeMaster As String = "EMPLOYEE-M"
    Public Const DesignationMaster As String = "DESIGN-M"
    Public Const DesignationMasterHierarchy As String = "DESIGN-M-H"
    Public Const CostCenter As String = "COST-CENTER"
    Public Const CostFACenter As String = "CFA-CENTER"
    Public Const userMaster As String = "USER-M"
    Public Const userGroupMaster As String = "USERGRP-MAST"
    Public Const userGroupMapping As String = "USER-GRP-MAP"
    Public Const groupProgramMapping As String = "GRP-PROG-MAP"
    Public Const frmScheduling As String = "SCN-NTF-SCH"
    Public Const frmApprovalLevelScreen As String = "APP-LVL-SCR"
    Public Const frmLocationSetting As String = "LOC-SET"
    'Public Const frmSynchronization As String = "On-Sync"
    Public Const FisaclYearEndProcess As String = "FY-END"
    Public Const frmAppIntegrator As String = "App-Inte"
    Public Const frmApprovalAlert_Child As String = "APP-CHD-SCR"
    Public Const frmApprovalAlertSumm As String = "APP-SUM-SCR"
    Public Const frmNotification As String = "NOT-FI-ION"
    Public Const FrmOfficeOrderTemplate As String = "OFI-ORD-TEM"
    Public Const DistrictMaster As String = "DIST-MST"
    'Public Const TimeTable As String = "TIME-TBL"
    Public Const Security_Matr As String = "Secu_Matr"
    Public Const frmDocumentVersionReport As String = "DOC_VER_RPT"
    Public Const FrmSystemAdministratorTool As String = "SYS_ADM_TOL"
    'Public Const FrmSendSMSMultipleUser As String = "SNDSMS"
    'Public Const UtilityImportExport As String = "UTLT_IM_EX"
    Public Const frmSendSMSEmailSetting As String = "EML_SMS_SET"
    Public Const ShiftEndForAllMcc As String = "SET_ALL_MCC"
    '----------------------------------Common Services------------------------------------
    '======================Setup============================
    Public Const cityMaster As String = "CITY-M"
    Public Const taxAuthority As String = "TAX-AUTH-M"
    Public Const taxRate As String = "TAX-RATE-M"
    Public Const taxGroup As String = "TAX-GRP-M"
    Public Const paymentTerms As String = "PAYM-TERM-M"
    Public Const paymentCodes As String = "PAYM-CODE"
    Public Const FrmAbateMentMaster As String = "ABATMENT"
    Public Const PrefixGeneration As String = "PRIFIX"
    Public Const FormMaster As String = "FRM-MAST"
    Public Const ChangePwd As String = "CHG-PWD"
    Public Const FrmAdditionalCharges As String = "ADSNL-CHRG"
    Public Const FrmCompanyMaster As String = "COMPANY-M"
    'Public Const frmCurrencyConversion As String = "Curr_Conv_Rt"
    'Public Const CustomFieldMaster As String = "CST_FIE_M"
    'Public Const CustomFieldMapping As String = "CST_FIE_MP"
    ' Public Const frmModuleCurrencyMapping As String = "Mod_Curr_Map"
    Public Const CommonServicesSetting As String = "COM-SER-SET"
    Public Const frmRegionMaster As String = "RGN-MST"
    Public Const frmBlockMaster As String = "BLOK-MST"
    Public Const frmCountryMaster1 As String = "CONT-MST"
    Public Const frmStateMaster1 As String = "STAT-MST"
    Public Const frmClientFormLableDetails As String = "FRM_LABEL"
    Public Const frmModuleScreen As String = "MODULE_SCREEN"
    Public Const FrmCalScreen As String = "CalSCRN"
    Public Const frmNotificationScreen As String = "NOTFSCRN"
    Public Const FrmFormSerialNoMaster As String = "FRMSRNO-MST"
    '' Anubhooti 24-Sep-2014
    'Public Const FrmBranchAccountMapping As String = "B_ACC_MAP"
    Public Const RptBranchAccountMapping As String = "RPT_B_ACC"
    Public Const FrmMCCDiscountMaster As String = "MCC_Dis"
    Public Const FrmLockTransactionReport As String = "RPT_LOCK_TRN"
    Public Const frmLocationDistanceMapping As String = "LOC-DIST-M"
    Public Const frmDisplaySquenece As String = "DIS-SEQ"
    Public Const frmRevenueVillageMaster As String = "REV_VIL_MST"
    Public Const frmGrampanchayatMaster As String = "GRM_PANCH_M"
    Public Const frmPanchayatSamitiMaster As String = "PANCH_SMI_M"
    Public Const frmVidhanSabhaMaster As String = "VIDN_SABA_M"
    Public Const frmRequestMaster As String = "USR_REQ_MST"
    '===================Transactions========================
    Public Const frmFormCollection As String = "FRM-COLL"
    Public Const FormIssue As String = "FRM-IS"
    Public Const FrmCFormEntry As String = "CFORM-ENT"
    Public Const FrmBankGuaranteeMaster As String = "BNK-GURNTE"
    Public Const BankOpeningReco As String = "BNK-OPR"
    'Public Const RevaluationEntry As String = "REVO-ENT"
    '===================Reports========================
    Public Const frmOpeningBalance As String = "OPNG-BALANCE"
    Public Const frmCFormReport As String = "C-FORM-RPT"
    Public Const RptBankWiseChequeIssue As String = "BANK_CHEQ"

    Public Const RptInventoryMovement As String = "RptInv-Mvt"
    Public Const rptItemWiseTaxMasterReport As String = "ITEM-TAX-RPT"
    Public Const frmDasboard As String = "RPT_DAS_BRD"
    Public Const frmDasboardCombine As String = "DAS_BRD_COM"
    'Public Const rptDataEntryTracingReport As String = "RPT_DATA_TRA"
    Public Const frmDocumentCancelledReport As String = "RPT_DOC_CANC"
    Public Const DashboardMilkProcurement As String = "DAS_MK_PROC"
    Public Const JanaadharStatusReport As String = "JAN_STS_RPT"
    '--------------------------------Receivable---------------------------
    '===================Setup====================

    Public Const ShiptoLocation As String = "SHIP-LOC-M"
    Public Const CustomerType As String = "CUST-TYPE"
    Public Const CustomeCategory As String = "CUST-CAT-M"
    Public Const mbtnCustomerInfo As String = "CUST-INFORM"
    Public Const CustomerMaster As String = "CUSTOMER-M"
    Public Const CustomerGroup As String = "CUST-GRP-M"
    Public Const CustomerAccountSet As String = "CUST-ACC-SET"
    Public Const BankGroupMaster As String = "BNK-GRP-MTR"
    Public Const bankMaster As String = "BANK-M"
    Public Const bankBranchMaster As String = "BNK-BCH-MST"
    Public Const frmShipToLocationDetails As String = "SP-LOC-D"
    Public Const frmCustomerCategoryLevel As String = "CSTCATLVEL"
    Public Const frmCustomerCategoryStructure As String = "CSTCATSTRUCT"
    ''richa Ticket no.BM00000003438 on 20/08/2014
    Public Const FrmReceivablePaymentTerms As String = "R-PAY-TERMS"
    Public Const rptMCCVLCVariationReportPourersNo As String = "VLC-VAR-PUR"  ' 15-12-2016
    Public Const FrmReceivableSettings As String = "REC-SET"
    '' Prabhakar Ticket : BM00000009801
    Public Const SecondaryCustomerMaster As String = "CUSTOMER-SM"
    '==================Transactions=====================
    Public Const ReceiptEntry As String = "RECEIPT-E"
    Public Const ReceiptAdjustmentEntry As String = "FINC-ADJ"
    Public Const PaymentAdjustmentEntry As String = "PA-ADJ-ENTRY"
    Public Const mbtnARInvoiceEntry As String = "AR-INVOICE"
    Public Const frmQuickBook As String = "QUICKBOOK"
    Public Const bankTransfer As String = "BANK-TRANSF"
    Public Const reverseTransaction As String = "REVERSE-TRAN"
    Public Const FrmBankReco As String = "BANK-RECOO"
    Public Const FrmCustomerInquiry As String = "CUST-INQ"
    Public Const FrmCustomersSetOff As String = "CUS_SET_OFF"
    Public Const FrmCustomersOutstanding As String = "CUS_OUTSTND"
    Public Const FrmBankUpdateUploader As String = "BNK_UPD_UPL"
    Public Const FrmMakePayment As String = "MakePayment"
    '=================Reports=====================
    Public Const rptCustomerAdvanceReg As String = "Rec_Adv_Rpt"
    Public Const RptCustomersSetOff As String = "CUS_SET_RPT"
    Public Const CustomerGroupReport As String = "CUST-GRP-RPT"
    Public Const CustomerDetails As String = "CUST-DTL-RPT"
    Public Const mbtnCustomerLedger As String = "CUST-LDG-RPT"
    Public Const mbtnCustomerLedgerZoneAreaWise As String = "CUST-LDG-AZ"

    Public Const frmRoute_CustomerOutstanding As String = "RT-CST-OSTDG"
    Public Const receiptreport As String = "RECEIPT-RPT"
    Public Const receiptFillreport As String = "RECEIPT-Fill"
    Public Const receiptWOTreport As String = "RECEIPT-WOT"
    Public Const CustomersListReport As String = "CUST-LST-RPT"
    Public Const frmCustomerAgeing As String = "CUST-AGE-RPT"
    Public Const CustomerRouteHistoryReport As String = "CUST-RHIS"
    Public Const cash_Register_Details4 As String = "CSH-RGR-RPT"
    Public Const mbtnCustomerEmptyTrial As String = "CUST-E-TR"
    Public Const FrmSecurityDeposit1 As String = "SEURTY-RPT"
    Public Const frmBankBook As String = "BANK-BK-RPT"
    Public Const frmReconcilationRpt As String = "RECON-RPT"
    Public Const frmBankBookRecoReport As String = "BBR-RPT"
    Public Const frmBankBookLocationDetail As String = "BLW-BK-RPT"
    Public Const frmCashVoucher As String = "CSH-VCHR"
    Public Const fundprovision As String = "FND-PRN"
    Public Const RptBankReconcilliation As String = "BNK-REC"
    Public Const FrmBankReverse As String = "BNK-Rev1"
    Public Const FrmCustomerOutstanding As String = "CUS-OUTS-RPT"
    Public Const FrmCustomerAgingSummary As String = "BI-CST-SMRY"
    Public Const frmBankBookChart As String = "BI-BNK-BK-1"
    Public Const frmBankBookClosing As String = "BI-BNK-BK-2"
    Public Const RptQualityStatus As String = "Qua_Sta"
    Public Const rptSecurityLevel As String = "Rpt-Sec_levl"
    Public Const rptVendorSecurity As String = "Rpt-Ven_Secu"
    Public Const FrmRptCustomerTransList As String = "CUST-T-LRPT"
    Public Const FrmRptCustomerTransHistory As String = "CUST-T-HIS"
    '====================Tax Report===================
    Public Const DVAT30 As String = "DVAT-30"
    Public Const DVAT31 As String = "DVAT-31"
    Public Const TaxTracking As String = "TAXTRACK"
    Public Const FrmDetailsOfForm2B As String = "FORM2B-RPT"
    '------------------------------Payables------------------------------------
    '=================Setup=======================


    Public Const vendoraccountset As String = "VEN-ACCT-SET"
    Public Const Farmeraccountset As String = "FAR-ACCT-SET"
    Public Const vendorgroup As String = "VEN-GRP"
    Public Const vendorSubgroup As String = "VEN-SGRP"
    Public Const vendortype As String = "VEN-TYPE"
    Public Const vendormaster As String = "VENDOR-M"
    Public Const INven As String = "INvenM"
    Public Const frmVendorAgingSummary As String = "BI-VND-SMRY"
    Public Const frmVendorCategoryLevel As String = "VNDCATLVEL"
    Public Const frmVendorCategoryStructure As String = "VNDCATSTRCT"
    Public Const FrmVendorListRPT As String = "VND-LIST-RPT"
    Public Const FrmPayableSettings As String = "PAY-SET"
    Public Const RptMultipleRTGS As String = "MUL_RTGS"
    Public Const RptMultiplePaymentAdvice1 As String = "MUL_PAY_AD"
    '' Pankaj Jha 23/09/2014
    Public Const VendorBankMaster As String = "VEN-B-M"
    Public Const VendorCustomerLedgerReport As String = "VENCUS-LR"
    'stuti
    Public Const capexmaster As String = "CAPEX-M"
    Public Const capexbudget As String = "CAPEX-BUDGET"
    Public Const EmployeeBandMaster As String = "BAND-MASTER"
    Public Const VendorRegistration As String = "VEN-REG"
    Public Const DispatchChecklist As String = "D-CHKLST"
    Public Const TankerMasterSale As String = "TNK-MST-SALE"
    Public Const BulkSaleFreightMaster As String = "BK-S-FR_MST"
    Public Const rptDailyWiseMilkCost As String = "MILK-COST"
    Public Const rptCapexRegister As String = "CAP-REG"
    'Public Const SRNReturnListCancellation As String = "SRN-RT-CAN"
    Public Const RptDayWisePurchasePriceReport As String = "DAY_PCH_PRC"

    Public Const RptPurchasePlanReport As String = "PUR_PLN_RPT"
    Public Const RptPurchaseMaterialRegister As String = "PUR_MTR_REG"
    Public Const RptTransporterProvisionReport As String = "RPT-TRA-PRO"
    Public Const rptBatchItemReport1 As String = "BTH_ITM_RPT"
    ''---
    '' Anubhooti 05-Sep-2014 BM00000003755
    Public Const FrmPaymentUploader As String = "PAY-UPL"
    Public Const FrmHirerachyLevelMaster As String = "HRC-LVL"
    '-----------------------------------------------Sales And Distribution---------------------------------------------
    '================Transaction====================

    Public Const saleOrders As String = "SALE-ORDER"
    Public Const LoadOut As String = "LOAD-OUT"
    Public Const saleReturn As String = "SALE-RET"
    Public Const frmQuickSettlement As String = "QUICK-SET"
    Public Const LoadOutStatus As String = "LOUT-STATUS"
    Public Const FrmSettlementEntry As String = "SET-ETY"
    Public Const FrmCompleteTransfer As String = "TRA-COM"
    Public Const FrmTransferIncompleteRemarks1 As String = "TRANS-INCOMP"
    Public Const SaleReturnInterCompany As String = "SAL-RET-INT"
    Public Const ChangeInvoiceSalesman As String = "CH-INV-SM"
    Public Const ReverseEntry As String = "REV-ENT"
    Public Const FrmCheckSlipEntry As String = "CHECK-SLIP"
    Public Const frmSaleHistory As String = "INQUIRY"
    '================Reports====================
    Public Const frmCarteJaliRpt As String = "CRT-JALI"

    '-------------------------------------------------Paybles--------------------------------------------------
    '==================Transactions==================

    Public Const PaymentEntryNew As String = "PYMT-NEW"
    Public Const mbtnAPInvoiceEntry As String = "AP-INVOICE"
    Public Const FrmVendorService As String = "VEN-SER-CHG"
    Public Const FrmVendorInquiry As String = "VEN-INQ"
    Public Const FrmSupplierReg As String = "SUPP-REG"
    Public Const FrmApprovedSuppliers As String = "APP-SUPP"
    Public Const FrmVendorSetOff As String = "VEN_SET_OFF"
    Public Const FrmVSPCSASetOff As String = "VSP_SET_OFF"
    Public Const FrmEmployeeSetOff As String = "EMP_SET_OFF"
    'Public Const bankTransfer As String = "PRIFIX"
    'Public Const reverseTransaction As String = "PRIFIX"
    '=================Reports=======================
    Public Const frmPaymentEntry As String = "PAY-ENT-RPT"
    Public Const mbtnAPInvoiceReport As String = "AP-INV-RPT"
    Public Const VendorLedgerReport As String = "VEN-LDG-RPT"
    Public Const frmAgingPayble As String = "AP-AGE-RPT"
    Public Const frmRptVendorList As String = "VEN-LST-RPT"
    Public Const FrmProvSaleExcel As String = "PRO-SL-EXL"
    Public Const frmAgingDrillDown As String = "AP-AGE-DRL"
    Public Const FrmRptVendorTransList As String = "VEN-T-LRPT"
    Public Const FrmRptVendorAgeingDetails As String = "VEN-AGE-D"
    'Public Const frmCashVoucher As String = "PRIFIX"
    Public Const frmAdvancePaymentRegister As String = "ADV-PAY-REG"
    Public Const rptAPReport As String = "AP-RPT"
    Public Const rptGSTRReport As String = "GSTR-RPT"
    Public Const rptARReport As String = "AR-RPT"
    Public Const RptPayableClearing As String = "PYBL-RPT"
    Public Const RptShippmentClearing As String = "SHMT-RPT"
    Public Const rptVendorAccountSetReport As String = "VAS-RPT"
    Public Const RptCustomerAccountSet As String = "CST-ACC-RPT"
    Public Const rptPaymentPayable As String = "PMT_PYB_RPT"
    '' Anubhooti 05-Aug-2014
    '------------------------------------------------------------Human Resource---------------------------------------
    '================================================================HR Settings============================================
    Public Const FrmHRSettings As String = "HR-Set"
    '================================================================Setup============================================
    Public Const JobTitle As String = "Job-T"
    Public Const frmHRParameterMaster As String = "HR_P"
    Public Const frmQualificationMaster As String = "Q_M"
    Public Const frmProfileMaster As String = "P_M"
    Public Const frmchkList As String = "Chk_LIS"
    Public Const frmRoundMaster As String = "R_M"
    Public Const EmployeeTypeMaster As String = "ETYP-M"
    Public Const frmRelationMaster As String = "Rel_M"
    Public Const frmSourceTypeMaster As String = "Src_T_M"
    Public Const frmSourceTypeDetail As String = "Src_T_D"
    Public Const HRBudgeting As String = "HR_Budg"
    Public Const RequesitionEntry As String = "REQ_ENT"
    Public Const RequesitionApproval As String = "REQ_Apr"
    Public Const RequesitionClose As String = "REQ_ClS"
    Public Const AgencyMaster As String = "AGENCY_M"
    Public Const HRIndustryType As String = "HR_I_T"
    Public Const HRVerticalMaster As String = "HR_V_M"
    Public Const TrainingMaster As String = "TRA_M"
    Public Const TrainingResourceMaster As String = "TRARES_M"
    Public Const TrainingInstituteMaster As String = "TRAINS_M"
    Public Const TrainingRequestMaster As String = "TRAREQ_M"
    Public Const OfferChkList As String = "Off_Chk"
    Public Const TRAINER_MASTER As String = "TRAINER_M"
    Public Const Schedule_Training As String = "Sch_TRAINING"
    Public Const Training_Attendance As String = "T_Attendance"
    Public Const FrmDamageMaster As String = "DAMAGE_M"
    ''
    '================================================================Transaction============================================
    Public Const frmApplicantEntry As String = "A_E"
    Public Const frmShortlist As String = "Short_L"
    Public Const frmInterviewSchedule As String = "Intr_S"
    Public Const frmInterviewFeedback As String = "Intr_F"
    Public Const frmReferenceCheck As String = "Ref_Chk"
    Public Const frmSalaryFitment As String = "Sal_Fit"
    Public Const JOININGCHECKLIST As String = "JOIN_CHK"
    Public Const frmOfferLetterHR As String = "Off_HR"
    Public Const frmAppointmentLetterHR As String = "App_HR"
    Public Const frmHireEmployee As String = "Hire_Emp"
    Public Const frmHrTraineeFeedBack As String = "Tree_Feed"
    Public Const frmHrTrainerFeedBack As String = "Trer_Feed"
    Public Const frmDamageCaused As String = "DAMAG_CAUS"
    '================================================================Performance Evaluation============================================
    Public Const frmHRPerformanceCategoryMaster As String = "Perf_E"
    Public Const frmHRPerformanceMaster As String = "Perf_Cat"
    Public Const frmHRPerformanceGroup As String = "Perf_Grp"
    Public Const frmHRPerformanceGroupMapping As String = "P_Grp_M"
    Public Const frmHRPerformanceRating As String = "Perf_Rate"
    Public Const frmHRPerformanceRatingRpt As String = "P_Rate_Rpt"
    '================================================================Reimbursement============================================
    Public Const frmHRReimbursementTypeMaster As String = "HR_RM"
    Public Const FrmHRTravelPurposeMaster As String = "HR_TP"
    Public Const FrmHRTravelBookedForMaster As String = "HR_TBM"
    Public Const FrmHRTravelCategoryMaster As String = "HR_TC"
    Public Const FrmHRTravelModeTypeMaster As String = "HR_TMT"
    Public Const FrmHRTravelCityMaster As String = "HR_TCM"
    Public Const FrmHRTravelClassTypeMaster As String = "HR_TCTM"
    Public Const FrmHRHotelRatingMaster As String = "HR_HRM"
    Public Const FrmHRTravelRoomTypeMaster As String = "HR_HRT"
    Public Const FrmHRTravelCarTypeMaster As String = "HR_TCAR"
    Public Const FrmHRRaiseTravelRequisition As String = "HR_RTR"
    Public Const FrmHRTravelReqApproval As String = "HR_TRQA"
    Public Const FrmHRTravelReimbursementExpense As String = "HR_TRE"
    Public Const FrmHRApprovarCreationMaster As String = "HR_ACM"
    Public Const FrmHRTravelExpenseApproval As String = "HR_TEA"
    '================================================================End===========================================
    '======Reports==========================================='
    Public Const RptRegisterOfDeduction As String = "REG_DED"
    Public Const RptCapxRevHis As String = "Capx_REV_Rpt"
    '================================================================Reimbursement============================================

    '================================================================Service And Warranty============================================
    Public Const FrmFaultCategory As String = "SW_Fau_Cat"
    Public Const FrmFaultMaster As String = "SW_Fau_Mas"
    Public Const FrmServiceChargeMaster As String = "SW_Ser_Char"
    Public Const FrmProblemType As String = "SW_Prob_T"
    Public Const FrmCallType As String = "SW_Call_T"
    Public Const FrmActivityType As String = "SW_Act_T"
    Public Const FrmSolutionKnowledgeBase As String = "SW_Sol_KB"
    Public Const FrmServiceEnquiry As String = "SW_Ser_Enq"
    Public Const FrmServiceAllocation As String = "SW_Ser_Allc"
    Public Const FrmServiceVisitDetails As String = "SW_Ser_VD"
    Public Const FrmServiceCall As String = "SW_Ser_CL"
    '================================================================End Service And Warranty============================================
    ''-----------------------------------------------------------General Ledger-----------------------------------------------------------------------
    '-------------------Setup-----------------------------
    Public Const glOptions As String = "GL-OPTION"
    Public Const segmentCode As String = "SEG-CODE-M"
    Public Const accountStructure As String = "ACCT-STRUC-M"
    Public Const accountGroup As String = "ACCT-GROUP"
    Public Const glAccount As String = "GL-ACCT"
    Public Const createAccounts As String = "CREATE-ACCT"
    Public Const sourceCode As String = "SOURCE-CODE"
    Public Const glsecurity As String = "GL-SECURITY"
    'Public Const FrmGL_account_excluded As String = "GL-EXCLUDED"
    Public Const frmProfitAndLossPerforma As String = "PRF-LOS-PRFM"
    Public Const frmBalanceSheetPerforma As String = "BAL-SHT-PRFM"
    'Public Const CashFlowPerforma As String = "CAS-FLO-PRFM"
    Public Const frm_Account_Mapping As String = "Act_Mapping"
    Public Const frmMapLedgerAccToTally As String = "MapGL_Tally"
    Public Const frmPostAllGLToTally As String = "PostGL_Tally"
    Public Const FiscalYear As String = "FIS_YEA"
    Public Const CostCentreFinancial As String = "CC_Fin"
    Public Const AccountMainGroup As String = "ACC-M-GRP"
    Public Const AccountSubGroup As String = "ACC-S-GRP"
    Public Const AccountGLMainAccount As String = "ACC-GL-MAC"
    Public Const FrmGLControlAccountMapping As String = "GL-CtlAccMap"
    '------------------Transactions----------------------------
    Public Const journalEntry As String = "JRN-ENTRY"
    Public Const mbtnVCGLEntry As String = "VCGL"
    Public Const frmJEReverse As String = "RVRS-JE"
    Public Const ReversejournalEntry As String = "RJRN-ENTRY"

    '------------------Reports--------------------------
    Public Const frmGLTransReport As String = "GL-TRAN-RPT"
    Public Const JECheckSystem As String = "JE-ENT-CHK"
    Public Const FrmCostCenterAnalysisRpt As String = "CstCntr-RPT"
    Public Const frmJrnlVoucher As String = "JRN-VOUCHER"
    Public Const rptTrialBalance As String = "TRL-BAL-RPT"
    Public Const rptTrialBalanceCV As String = "TRL-CV-RPT"
    Public Const mbtnJournalBook As String = "GL-JB-RPT"
    Public Const rptProfitAndLoss As String = "PRF-LOS-RPT"
    Public Const rptBalanceSheet As String = "BAL-SHT-RPT"
    Public Const CashFlow As String = "CAS-FLO-RPT"
    Public Const frmViewTDS As String = "VIEW-TDS"
    Public Const frmBankBookDayWise As String = "BNK-BK-DW"
    Public Const frmUnpostedJV As String = "UnPost-JV"
    Public Const rptChartOfAccount As String = "CHRT_OF_ACC"
    ''----------------------------------------------Sales & Distribution--------------------------------------------------------------
    '-------------------Setup-----------------------------

    Public Const groupMasterRoute As String = "GRP-ROUTE-M"
    Public Const routeMaster As String = "ROUTE-M"
    Public Const transportMaster As String = "TRANSP-M"
    Public Const transportMasterVendor As String = "TRANSP-MV"
    Public Const vhicleMaster As String = "VEHICLE-M"
    Public Const channelCategory As String = "CHANNEL-CAT"
    Public Const channelMaster As String = "CHANNEL-M"
    Public Const visiMaster As String = "VISI-M"
    Public Const transportType As String = "TRANS-TYPE"
    Public Const routetypemaster As String = "ROUTE-TYPE"
    Public Const frmSalesManHierarchy As String = "SalesMan-H"
    Public Const frmSettlementMaster As String = "SETTLEMENT-M"
    Public Const frmCommissionMaster As String = "COMM-M"
    Public Const CustomerVendorMapping As String = "CST-VNDR-MAP"
    Public Const CustomerVendorMappingVendor As String = "CST-VNDR-MAV"
    Public Const frmCustomerEmployeeMapping As String = "CST_EMO_MAP"

    Public Const rptDispatchChallanReport As String = "SMDptRpt"
    Public Const rptMatrixFreshReport As String = "SMMRXRpt"
    Public Const rptZoneWiseReport As String = "SMZNRpt"

    Public Const frmRouteShifting As String = "OTLT-SIFT"
    Public Const frmDiscountMaster As String = "DISC-MAST"
    Public Const frmDCSforSale As String = "DCS-SALE"
    Public Const frmFrieghtRateMaster As String = "FRI-RAT-MST"
    Public Const Sampling_Master As String = "SMPL-MAST"
    Public Const frmDiscountCategoryMaster As String = "DIS-CAT-MST"
    Public Const mbtnTargetMaster As String = "TRGT-MSTR"
    Public Const mbtnTmplateCreation As String = "TMP-CR"
    Public Const FrmViewPunchingInvoice As String = "PUNCH-ETY"
    Public Const FrmGatePassENtry1 As String = "GP-ENTRY"
    Public Const frmCustomerTargetFixing As String = "CUST-TARGET"
    Public Const frmClaimMaster As String = "CLAIM_MAST"
    Public Const frmClaimReportNew As String = "CLAIM_RPT_N"
    Public Const FrmBatchReceiptSTD As String = "BAT_REC_STD"
    Public Const FrmBatchReceiptPepsi As String = "BAT_REC_PEP"
    Public Const frmCheckPrinting As String = "Check_Print"
    '------------------------Sales Reports---------------------------
    Public Const crptLoadOut As String = "SAL-INV-RPT"
    Public Const nrptSales As String = "SAL-REG-RPT"
    Public Const SaleReport As String = "SL-VOL-RPT"
    Public Const rptCreditSaleReport As String = "CREDIT-SALE"
    Public Const NoSaleReport As String = "NO-SALE"
    Public Const OtherPartySale As String = "OT-PRT-SL"
    Public Const PendingSaleOrderReport As String = "PD-SLO-RPT"
    Public Const mbtnNetSaleReport As String = "NET-SL-RPT"
    Public Const rptNetSaleDetailReport As String = "NET-DTSALE"
    Public Const ItemCommissionSummary As String = "ITM-CM-RPT"
    Public Const CEAllocationReport As String = "CE-ALOC-RPT"
    Public Const DailySettlement As String = "DAILYSET-RPT"
    Public Const SalesCollection As String = "SALVsCOLL"
    Public Const ProvisionalSaleReport As String = "PRO-SAL-RPT"
    Public Const ProvSaleDetail As String = "PRV-SL-DT"
    Public Const reportQuickSettlement As String = "ITM-DIS-RPT"
    Public Const mbtnCustomerRanking As String = "CUS-RAK-RPT"
    Public Const rptPenetration As String = "SAL-PENET"
    Public Const ItemDiscountReport As String = "DISC-REPO"
    Public Const DistributedDiscountReport As String = "DSTR-D-RPT"
    Public Const frmSaleDiscount1 As String = "SALE-DIS-RPT"
    Public Const frmMCDiscountReport As String = "MRG-CNT"
    Public Const SaleAccountBreakDetail As String = "SL-AC-BRKG"
    Public Const SaleAccountSummary As String = "SL-AC-SUMM"
    Public Const mbtnCashDiscountReport As String = "CSH-DIS-RPT"
    Public Const PrimarySales As String = "PRIM-SAL-RPT"
    Public Const SecondarySales As String = "SEC-SALE-RPT"
    Public Const mbtnMismatchReport As String = "MIS-RPT"
    Public Const mbtnRouteSale As String = "RUT_SAL"
    Public Const Settlement As String = "SETL-RPT"
    Public Const mbtnRptSalesManSalesReport As String = "SALM-SAL-RPT"
    Public Const RptTransfer_IncompleteReport As String = "TRANS-INRPT"
    Public Const FrmRGP_Register_NRGP As String = "RGP-RPT"

    'KUNAL > TICKET : BM00000010298 > CLIENT : UDL > DATE : 28-NOV-2016
    Public Const FrmRpt_OutStnd_Items_RGP = "OUTSTND-RGP"

    Public Const FrmRoutewiseSaleReport As String = "RTW-SALE-RPT"
    Public Const FrmKPIReport As String = "KPI-RPT"
    Public Const RptPendingSettlement As String = "PEN-SET-RPT"
    Public Const CashDiscountReport As String = "SL-CA-DI"
    Public Const CashDiscount As String = "CASH-DISC"
    Public Const TransferRegister As String = "TRANS-REG"
    Public Const SaleAccountBreakage As String = "SA-ACC-BRK"

    Public Const FrmLoadReport As String = "LOAD-RPT"
    Public Const frmLoadOutInvoiceRecoReport As String = "Lod-Rec-RPT"
    Public Const FrmTargetReport1 As String = "TARGET-RPT"
    Public Const FrmDailySettlementActualAndProvisionalReport As String = "DSET-A/P-RPT"
    Public Const frmReverseSettlementDetail As String = "RV-STMT-DTL"
    Public Const frmMismatchCashMemo As String = "MM-CSH-MEM"
    Public Const FrmSettlement_CashMemoStatus As String = "SSH-MEMO-ST"
    Public Const frmCanceledSaleInvoice As String = "CNC-SI-RPT"
    Public Const frmSaleVolumeTracker As String = "SLE-VOL-TRK"
    Public Const FrmEmptyTransactionReport As String = "ET-RPT"
    Public Const frmpendingLoadin As String = "PNDLODIN-RPT"
    Public Const FrmMismatchSettlement As String = "MMSETTLE-RPT"
    Public Const FrmRptSalesReturn As String = "SALERETN-RPT"
    Public Const FrmQuickSettlementHead As String = "SET-HED-RPT"
    Public Const FrmSettlementSheetReconcilationeport As String = "SET-SRCO-RPT"
    Public Const FrmVendorsOutstandings As String = "VEN-OUTS-RPT"
    Public Const FrmSaleOrderSummary As String = "PNDORDER-RPT"
    Public Const FrmDistrbutorSaleTarget As String = "DIS-TARG-RPT"
    Public Const FrmDayReportDirectSale As String = "DAY-DIRECT-S"
    Public Const KeyValue As String = "KEY-VALUE"
    Public Const FrmLeakageBreakage As String = "Leak-Break"
    Public Const RptInvoiceAgainstInward As String = "In-Agnst-In"
    Public Const FrmDiscountAnalysis As String = "DIS-ANAl-RPT"
    Public Const frmCustomerTargetReport As String = "CUST-TAR-RPT"
    Public Const rptClaimMaster As String = "RPT_CLAIM_MS"
    Public Const frmTDMReport As String = "TDM-RPT"
    Public Const SalesmanSalesOrderReport As String = "SM-SOR"
    Public Const frmSaleAnalysisReport As String = "SL-ANLS-RP"
    Public Const rptSalesAnalysis As String = "SALE-AnlsRPT"
    Public Const frmCustomerBillWiseDetail As String = "CUST-BILL"
    Public Const frmRptFormOfGuarntee As String = "FM-GUARNTEE"
    Public Const RptCustomerWiseMonthlySalesAnalysis As String = "SALE-ANLYSIS"
    Public Const frmCustomerBillWiseDuesSummary As String = "CST-DU-SMRY"
    Public Const frmVendorGroupWiseSaleReport As String = "VNDR-GRP-WIS"
    Public Const frmPaySlipReport As String = "PAY-SLIP"
    Public Const frmDVAT32 As String = "SALES-DVAT32"
    Public Const frmDVAT33 As String = "SALES-DVAT33"
    Public Const FrmVehiclePendingStatusRpt As String = "VHL-PND-RPT"
    Public Const frmSaleSummaryAgainstPO As String = "SALE-SUMM"
    Public Const frmYearMonthWiseSaleComparison As String = "YR-MNTH-CMP"
    Public Const frmCompanyMonthWiseSaleComparison As String = "CP-MNTH-CMP"
    Public Const rptDistributerPerformance As String = "dist-Perform"
    '--------------Shipping Reports--------------------------------


    Public Const LoadOutReport1 As String = "LD-RPT"
    Public Const FilloutwardRegisterReport1 As String = "FL-OTWD-RPT"
    Public Const EmptyInwardSaleRegister1 As String = "INWD-RG"
    Public Const vehicle_Details_Report1 As String = "VEH-DT-RPT"
    Public Const LO_vs_Vechile As String = "LO-VEH-RPT"
    Public Const GatePass_Vs_actual As String = "GA-VEH-ACT"

    Public Const visiDetail1 As String = "VISI-DT"
    Public Const OutletEmpty1 As String = "OUT-EMT-RPT"
    Public Const VehiclewiseSale1 As String = "VHL-SALE-RPT"
    Public Const mbtnVisiVPO1 As String = "VISI-VPO-RPT"
    Public Const Channelwisecustomer1 As String = "CHL-CUS-RPT"
    Public Const RouteListReport11 As String = "RT-LIST-RPT"
    Public Const LoadOutStatusreport1 As String = "LOUT-STS-RPT"
    Public Const frmTransitBreakageReport1 As String = "TRN-BRK-RPT"
    Public Const FrmFilledOutWard As String = "FIL-OUT-RPT"
    Public Const EmptyReportDetail As String = "EMP-RPT_DTL"
    '------------------Excise Reports-------------------------------

    Public Const frmExciseChapterWise As String = "EX-CH-RPT"
    'Public Const ExciseSummary1 As String = "EX-SUMM"
    Public Const ExciseSummaryNew As String = "EX-SUMM-N"
    Public Const CrptRG1Detail1 As String = "RG-RPT"
    Public Const FrmDetailOfForm2A As String = "FORM2A-RPT"
    'Public Const RptServiceTaxDetail As String = "SERV_TAX"
    ''-------------------------------------------------Material Management-----------------------------------------------------------------
    '-------------------Setup-------------------------------
    Public Const PricePlan As String = "PRICE-P"
    Public Const PriceMaster As String = "PRICE-M"
    Public Const PriceComponentMapping As String = "PRIC-COM-MAP"
    Public Const frmPriceGroupMapping As String = "PRIC-GRP-MAP"
    Public Const PriceComponentMasters As String = "PRICE-COMP-M"
    Public Const SchemeMaster As String = "SCHEME-M"
    Public Const inventorySetting As String = "INV-SETT-M"
    Public Const itemStructure As String = "ITEM-STRUCT"
    Public Const itemGroups As String = "ITEM-GROUP"
    Public Const itemMaster As String = "ITEM-MASTER"
    Public Const locationMaster As String = "LOCATION-M"
    Public Const CustomerLocationMapping As String = "CUS-LOC-MAP"
    Public Const itemPurchaseAccount As String = "ITEM-PUR-ACC"
    Public Const itemSaleAccount As String = "ITEM-SAL-ACC"
    Public Const GSTunitMeasure As String = "GST-U-M"
    Public Const unitMaster As String = "UNIT-M"
    Public Const packType As String = "PACKTYPE-M"
    Public Const ItemLocationDetails As String = "ITEM-LOC"
    Public Const mbtnItemCategory As String = "ITEM-CAT"
    Public Const mbtnItemSubCategory As String = "ITM-SUB-CAT"
    Public Const ItemReorderLevel As String = "ITM-REOD-M"
    Public Const chapterhead As String = "CHAP-M"
    Public Const ItemExciseMapping As String = "ITEX-MAP-M"
    Public Const ItemBasicPrice As String = "ITEM-BASIC-P"
    Public Const mbtnBreakageHead1 As String = "BREAKAGE"
    Public Const FrmItemMasterRMOther As String = "RM-OTHERS"
    Public Const TypeOfItemMaster As String = "TYP-OF-ITEM"
    Public Const frmStandardscheme As String = "STD_SCHEME"
    Public Const frmStandardRateItem As String = "STD_RAT_ITM"
    Public Const frmItemCategoryLevel As String = "ITMCATLVEL"
    Public Const frmItemCategoryStructure As String = "ITMCATSTRUCT"
    Public Const frmBarCodeGenerator As String = "BR-CD-GNTR"
    Public Const WarrantyMaster As String = "WAR-MAST"
    Public Const frmSchemeMasterNew As String = "SCHM-MSTR-N"
    Public Const frmWeightConversion As String = "WEIGHT-CONV"
    'Public Const SublocationMaster As String = "S_LOCATION-M"
    Public Const frmLocationCategoryLevel As String = "LOCCATLVEL"
    Public Const frmLocationCategoryStructure As String = "LOCCATSTRUCT"
    ''richa 21/08/2014
    Public Const FrmCatalogMaster As String = "CAT-MASTER"
    Public Const frmPartNoMaster As String = "PART-NO-MSTR"
    Public Const InvetorySourceCode As String = "INV-SOR"

    Public Const FrmItemTypeMaster As String = "ITM-TYP-M"
    'Public Const ItemRackBinMapping As String = "ITEM-R_B_M"
    'Public Const LoanEntry As String = "LOAN-ENTRY"
    'Public Const LoanInstallment As String = "LOAN-INST"
    Public Const frmRequestAproval As String = "REQ_APROVL"
    '------------------Transactions-----------------------------


    Public Const Transfer As String = "STO-TRANSFER"
    Public Const CreateTransfer As String = "STO-CRETRA"
    Public Const adjust As String = "ADJUST-ENTRY"
    Public Const mbtnEmptyTrans As String = "EMPTY-TRANS"
    Public Const mbtnProductionEntry As String = "PROD-ENTRY"
    Public Const mbtnStoreAdjustment As String = "STORE-ADJ"
    Public Const frmStartBatchEntry As String = "S-BATCH-ENT"
    Public Const FrmItemMcMapping As String = "ITEM-MC-MAP"
    Public Const frmWarehouseBreakage As String = "WH-BRKG"
    Public Const frmPhysicalStock As String = "PHY-STOCK"
    Public Const frmPhysicalStockMultipleLocation As String = "PHY-STK-MLO"
    Public Const ChangeItemSerialNumber As String = "CHG-ISN"
    Public Const Indent As String = "TRN-IND"
    Public Const FrmExpiryDateEntry As String = "EXPIRY-ENT"
    Public Const frmRiceBOM As String = "RICE-BOM"
    Public Const frmRiceMixingEntry As String = "RICE-MIX"
    Public Const frmRiceProcessingEntry As String = "RICE-PROC"
    Public Const TransferReturn As String = "STO-TRANS-R"
    Public Const GatePassTransfer As String = "GP-TRANSFR"
    Public Const TransferCrateReceived As String = "TRANSFR_CR"
    Public Const frmGeneralWeighment As String = "G-WEIGHT"
    Public Const frmRMProcessLoss As String = "RMI-PRS-LSS"
    '---------Reports--------------------------

    Public Const mbtnItemMovement As String = "ITM-MOV-RPT"
    Public Const ItemLocationDetailsReport As String = "ITM-LOCD-RPT"
    Public Const ItemPrice As String = "ITM-PRIC-RPT"
    Public Const StockRecoReport As String = "STOCK-RPT"
    Public Const FrmStockDispatchReport As String = "STOCK-DIS-RP"
    Public Const ReportTransfer As String = "TRN-RPT"
    Public Const mbtnStockAdjustmentReport As String = "ST-ADJ-RPT"
    Public Const FrmAdjustmentStatusReport1 As String = "ADJ-ST-RPT"
    Public Const mbtnBreakageReport As String = "BRK-RPT"
    Public Const BreakageReportSummary As String = "BRK-RPT-SUM"
    Public Const RoutewiseBreakageReport As String = "RT-BRK-RPT"
    Public Const SchemeReport As String = "SCH-RPT"
    Public Const StockReportForFinishedGoods As String = "STK-FIN-RPT"
    Public Const FrmAdjustmentReport As String = "ADJ-RPT"
    Public Const rptVehicleWiseLoadout As String = "VLO-RPT"
    Public Const FrmPendingIndentTransferReport As String = "PEN-IND-RPT"
    Public Const FrmExpiredItemDetails As String = "EXP-ITM-RPT"
    Public Const stockRecoNew As String = "STO-REC-RPT"
    Public Const UnionStockReport As String = "UNN-STK-RPT"
    Public Const FATSNFGainLoss As String = "FAT-SNF-GAN"
    Public Const MeterialstockReco As String = "MTRL-STO-REC"
    Public Const stockRecoBatch As String = "STO-RECB-RPT"
    Public Const frmBarCodeGenerator1 As String = "BR-CD-GNTR1"
    Public Const FrmItemSerialTrackingReport As String = "SERIAL-TRA-R"
    Public Const frmInventoryAgeingReport As String = "INV-AGE-RPT"
    Public Const FrmItemListRpt As String = "ITEM-LST-RPT"
    Public Const frmTransferRegister As String = "TRANS-REG"
    Public Const frmFatSnfStockReport As String = "FAT-Rpt"
    Public Const stockRecoNewJR As String = "STO-RJR-RPT" '' Stock Reco Job work
    Public Const frmDatewiseQtyFatSnfStockReport As String = "Q-FAT-Rpt"
    Public Const frmMCCMilkLossGain As String = "Q-LG-Rpt"
    Public Const RptUnpostingTransItemQty As String = "RPT_UPOS_QTY"
    Public Const RptTankerInTransit As String = "RPT-TAK-TRAN"
    Public Const RptItemPurchaseSet As String = "ITM-PUR-SET"
    Public Const rptSaleAccountSetList As String = "ITM-SAL-SET"
    Public Const RptItemSalePurchaseSet As String = "ITM-PUR-SAL"
    Public Const rptCustomerWiseStockReco As String = "CUST-REC-RPT"
    Public Const ItemStockReport As String = "ITEM_STO_RPT"
    Public Const ItemRateStatusReport As String = "ITEM_RAT_STA"
    Public Const rptDailySTAFGSStatement As String = "DLY_STA_FGS"

    'Public Const rptStockReport As String = "STOCK-RPT"
    ''-----------------------------------------------------Purchase----------------------------------------------------------------------
    '-------------------Setup--------------------------
    Public Const frmPurchaseSetting As String = "PUR-SETT"
    Public Const VendorItemDetails As String = "VEN-ITEM-DET"
    Public Const TDMTARGET As String = "TDM-EMP-DET"
    'Public Const frmRequisitionApproval As String = "REQ-APRVL"
    Public Const RequisitSubTypeMaster As String = "REQ-SUB"
    Public Const FrmFormIssueReceiptEntry As String = "FRM_ISS_RCV"
    ''richa Ticket no BM00000003507
    Public Const FrmCostCentreGroupStores As String = "COST-CENTRE"
    Public Const frmDeliveryTermsMaster As String = "DEL_TRM_MST"
    '-------------Transactions--------------------------
    Public Const RFQ As String = "PO-RFQ"
    Public Const frmStoreRequistion As String = "STORE-REQ"
    Public Const mbtnPurchaseRequistion As String = "PO-REQ"
    Public Const RMDemanApproval As String = "RM-DMD-APP"
    Public Const mbtnPurchaseOrder As String = "PO-ODR"
    'Public Const SetPOSchedule As String = "SET-PO-SCH"
    Public Const mbtnGRN As String = "PO-GRN"
    'Public Const VisualRandomQC As String = "Vs-Rn-QC"
    Public Const POWeighment As String = "PO-WHT"
    Public Const POUnloading As String = "PO-UND"
    Public Const mbtnMRN As String = "PO-MRN"
    Public Const NIRQC As String = "PO-NIR-QC"
    Public Const mbtnSRN As String = "PO-SRN"
    Public Const PurchaseGateOut As String = "GTE-OUT"
    Public Const SRNReturn As String = "PO-SRN-RET"
    Public Const mbtnPurchaseInvoice As String = "PO-INV"
    Public Const mbtnPurchaseReturn As String = "PO-RET"
    Public Const mbtnPurchaseJobWork As String = "PO-JWK"

    Public Const frmTender As String = "Pur-Tender"
    Public Const TenderShortPenalty As String = "TND-SHT-PNL"
    Public Const RALNOC As String = "RAL-NOC"
    Public Const frmCorrectionforWrongEntry As String = "CF-WRNG-ENT"
    Public Const frmDeletionForEntry As String = "DEL-FOR-ENT"
    Public Const frmShortSupplyPenalty As String = "SHT-SLY-PLT"

    'Sanjay BHA/09/05/18-000014
    Public Const frmMaterialQuotation As String = "SCRAP-QU"
    Public Const frmMaterialQuotationOrder As String = "SCRAP-QU-OR"
    Public Const frmMaterialQuotationComparison As String = "SCRAP-QU-CMP"

    'KUNAL > CLIENT : UDIL > TICKET : BM00000010226 
    Public Const mbtnNRGP As String = "PO-NRGP"

    Public Const mbtnGatePass As String = "PO-GP"
    Public Const mbtnIssueReturn As String = "PO-IRT"
    Public Const ScrapSale As String = "SCRAP-SALE"
    Public Const ScrapSaleRetrun As String = "SCRAP-S-R"
    'Public Const JobWorkDispatch As String = "JOBW-DISP"
    Public Const JobWorkDispatchProduction As String = "JOBW-DISP_PR"
    Public Const FrmScrapSaleGateOut As String = "SCRP-GATOUT"
    Public Const mbtnPendingApprovalOfReq As String = "PO-PRE"
    Public Const VendorQuotation As String = "PO-VQU"
    Public Const VendorComparison As String = "VEN-COM"
    Public Const VendorComparisonApproval As String = "VEN-COMA"
    Public Const FrmPurchaseHistory As String = "PRC_HIS_REP"
    'Public Const FrmItemConversion As String = "Item_Conv"
    Public Const frmPurchaseSchedule As String = "PO-SCHD"
    Public Const frmNEFTUploader As String = "NEFT-UP"
    Public Const frmNEFTUploaderFarmer As String = "NEFT-UP_FAR"
    Public Const frmQuickPaymentBySingleCheque As String = "QCK-PNT-CHQ"
    ''richa
    'Public Const frmItemQuantityInformation As String = "PO-IQI"

    '------------------Reports----------------------------------

    Public Const mbtnGRNReport As String = "GRN-RPT"
    Public Const frmPo_action As String = "PO-TRA-RPT"
    Public Const FrmPendingRequisitionQty As String = "RQ-RPT"
    Public Const FrmPendingGrn_Qty As String = "PEN-GRN"
    Public Const FrmPendingMrn_Qty As String = "PEN-MRN"
    Public Const FrmPendingSrn_Qty As String = "PEN-SRN"
    Public Const Parti_VS_Rejected As String = "PARTY-REJ"
    Public Const Store_Receipt_Note As String = "DLY-REC-N"
    Public Const FrmIssueOrReturnItemWiseSummary As String = "NET-ISS-RPT"
    Public Const Vendor_Rating_Rejection As String = "VEN-RET-REJ"
    Public Const FrmPurchasebookReport1 As String = "PUR-BOOK"
    Public Const frmStockAnalysis As String = "STK-AYS-RPT"
    Public Const frmMorningReport As String = "MOR-RPT"
    Public Const frmVendorWiseReturnableGoodBalance As String = "VWBRGB"
    Public Const mbtnStoresLedger As String = "STR-LGD-RPT"

    Public Const mbtnStoreLedger As String = "STR-LGR-RPT"
    Public Const RM_Consumption_Detail As String = "RM-CONS-RPT"
    Public Const AddCharge As String = "ADD-CHARGE"
    Public Const frmConsumptionReport1 As String = "CONSM-RPT"
    Public Const mbtnDailyRcptNoteSummary As String = "DRN-SUM-RPT"
    Public Const DetailofWtdPriceofRawMaterial As String = "DET-WTD-RAW"
    Public Const FrmIndentReport As String = "PEN-IND-RPT"
    Public Const StockStatement As String = "STOCK-ST"
    Public Const FrmPurchaseOrderRegister As String = "Pur-Ord-Rag"
    Public Const frmItemWiseDispatchLedger3 As String = "ITM-DISP-LED"
    Public Const FrmFreightCosting As String = "FRGHT-RPT"
    Public Const FrmRptPurchaseReturnBook As String = "PUR-RET-BOK"
    Public Const frmPurchaseOrderList As String = "LIST-PO-REP"
    Public Const frmPurchaseOrderAmd As String = "LIST-PO-AMD"
    Public Const rptMaterialSendforJobWork As String = "MAT-SND-JW"
    Public Const rptMaterialReceivedAfterJobWork As String = "MAT-REC-JW"
    Public Const rptBalanceStockForJobWork As String = "BLN_STK-JW"
    Public Const rptRGPWiseJobWork As String = "RGP_WIS-JW"
    Public Const rptSparePartStatus As String = "SPR_PAR-STA"
    Public Const rptUnpostedPO As String = "PO-UNP"
    '-------------------------------Transaction Report''--------------------------------------------

    Public Const mbtnRGP_NRGP_Rpt As String = "RGPNRGP-rpt"
    Public Const MRDAReport As String = "MRDA-RPT"
    Public Const DebitAdviseReport As String = "DEBIT-RPT"
    Public Const frmScrapSaleInvoice As String = "SRP-IN-RPT"
    Public Const frmSrnReport As String = "SRN-RPT"
    Public Const PJVReport As String = "PJV-RPT"
    Public Const FrmPurchaseOrderReport As String = "PO-RPT"

    'DATE - 14 -02 -2017  - UDL
    Public Const frmMonthlyConsumptionReport As String = "MCR-RPT"
    Public Const frmHSNMaster As String = "HSN-MASTER"

    ''----------------------------------------------------Tax Deducted At Source------------------------------------------------------------------------
    '---------------------Setup----------------------------

    Public Const FinancialYear As String = "FIN_YEAR"
    Public Const ResponsiblePerson As String = "RESP-PERSON"
    Public Const BranchDetails As String = "BRANCH-DET"
    Public Const NatureOfDeduction As String = "DED-NATURE"
    Public Const PartyDetails As String = "PARTY-DET"
    Public Const TDSSection As String = "TDS-SECTION"
    Public Const StateCode As String = "STATE-M"

    '-----------------Transactions---------------------

    Public Const mbtnCreateRemittance As String = "CREAT_REMIT"
    Public Const remittanceentry As String = "REMIT-ENTRY"
    Public Const mbtnAPInvoiceEntryTDS As String = "AP-INVTDS"
    Public Const TDSPAYMENT As String = "TDS-PY"
    '--------------------Reports------------------------------

    Public Const frmrptTDSLedger As String = "TDS-LDG-RPT"
    Public Const TDSForm26Q As String = "FRM26/27-RPT"
    Public Const TDSSectionSummaryReport As String = "TDS-SEC-RPT"
    Public Const Form16AReport As String = "FRM16-RPT"

    '-------------------------------------------------------------Fixed Assets-----------------------------------------------------------------
    '---------------------Setup----------------------------

    Public Const AssetSegment As String = "ASST-SGMT"
    Public Const DepAccSets As String = "Dep-AcSets"
    Public Const DepPeriod As String = "Dep-Period"
    Public Const Categories As String = "ASST-CAT"
    Public Const FrmDepreciationField As String = "DEP-FIELD"
    Public Const frmAssetGroups As String = "ASST-GRPS"
    Public Const FrmAMAcquisitionCode As String = "ACQUI-CODE"
    Public Const frmDepreciationMethod As String = "DEP-MTD"
    Public Const frmSecondaryCustomer As String = "SEC-CST-MST"
    Public Const FrmConsumerDetailsForm As String = "CON-DET-FRM"
    Public Const fixedsetting As String = "ASST_SETTING"
    'Public Const frmCompetitorMaster As String = "COMPET-MST"
    '-----------------Transactions---------------------
    Public Const FAAcquisitionEntry As String = "FA-ACQE"
    Public Const FAAssetDepreciation As String = "FA-ACDE"
    Public Const FADisposalEntry As String = "FA-DISE"
    Public Const frmVisi_Install_Pullout As String = "VISI-IN_OUT"
    Public Const frmAsset_Issue_Return As String = "ASST-ISS-RET"
    Public Const frmAssetRequisition As String = "ASSET-REQ"
    Public Const frmAssetStoreRequistion As String = "FA-REQUI"
    Public Const frmSecondaryCustomerSale As String = "SC-CST-SALE"
    Public Const FAAssetWork As String = "FA-WORK"
    Public Const frmIssueItemsToAsset As String = "FA-IssueItem"
    Public Const FAMergeAcquisitionEntry As String = "FA-MRG-ACE"
    Public Const FAAssembleAsset As String = "FA-ASSMBL"
    Public Const frmAssetAccountChange As String = "FA-AAC"
    Public Const frmAssetDispatchRetailer As String = "FA-DIS-RET"
    '--------------------Reports------------------------------
    Public Const FrmAssetRegister As String = "FA-AR"
    Public Const FrmAssetDetail As String = "FA-DE"
    Public Const FrmDisposalDetail As String = "FA-DD"
    Public Const frmVisi_Install_Pullout_Report As String = "VISI-IO-RPT"
    Public Const frmAsset_Issue_Return_Report As String = "ASST-IR-RPT"
    Public Const frmDistributor_VS_SecondaryCustomer_Sale As String = "D_VS_S_CUS"
    Public Const rptFARReport As String = "FA_FAR_RPT"
    Public Const rptCWIPReport As String = "CWIP_RPT"
    Public Const rptALCRReport As String = "ALCR_RPT"
    Public Const rptCapexConsumptionRpt As String = "C_Cons_RPT"
    '-------------------------------------------------------------Code Ends Here----------------------------------------------------------------
    ''Business Inteligence
    ''Setup 
    Public Const BICreateReport As String = "BI-CR"
    Public Const BICreateFilter As String = "BI-CF"
    Public Const BICreateDashBoard As String = "CR-DB"
    ''Sale
    Public Const BIMonthWiseSale As String = "BI-MS"
    Public Const BITopCustomer As String = "BI-TC"
    Public Const BITopItemCategory As String = "BI-TIC"
    Public Const BIDashBoadr As String = "BI-DB"
    ''Purchase
    Public Const BIMonthWisePurchase As String = "BI-MP"
    Public Const BITopVendor As String = "BI-TV"
    Public Const BITopItemCategoryPurchase As String = "BI-TIV"

    ''Asset
    Public Const BIMonthWiseAssset As String = "BI-MAS"

    ''Finance
    Public Const BITopExpence As String = "BI-TE"
    ''--------------------------------------------------Utility--------------------------------------------------------------------------------------------
    Public Const Calculator As String = "calc"
    Public Const mbtnTakeBackup As String = "BACKUP"
    Public Const mbtnRestoreDB As String = "RESTORE-DB"
    Public Const mbtnCreateReceiptAgainstInvoice As String = "CRE-REM"
    Public Const mbtnCreateReceiptAgainstSale As String = "DAY-ETY"
    Public Const mbtnPendingApproval1 As String = "PENAPP"
    Public Const frmBulkPostingNew As String = "NEW-BULK-PST"
    Public Const lockTransaction As String = "LOCK-TRANS"
    Public Const FrmDayWiseLoadOutEntered As String = "DAY-WISE-LOD"
    Public Const frmUserPerformanceDetail As String = "USR-PRF-DTL"
    Public Const frmStockReport As String = "STCK-RPT"
    Public Const FrmReconciliationSetting As String = "RECO-SETT"
    Public Const FrmUtilityForm As String = "Util-Frm"

    ''-----------------------------------------------------------HR and Payroll-----------------------------------------------------------------------
    '-------------------Setup-----------------------------
    Public Const frmSalaryAccountSetting As String = "SAL-ACC"
    Public Const frmDepartmentMaster As String = "DEP-MASTER"
    Public Const frmSkillMaster As String = "SKIL-MASTER"
    Public Const frmLanguageMaster As String = "LANG-MASTER"
    Public Const frmCourseMaster As String = "COUR-MASTER"
    Public Const frmShiftMaster As String = "SHIF-MASTER"
    Public Const frmDocumentMaster As String = "DOC-MASTER"
    Public Const frmPayPeriodMaster As String = "PAY-MASTER"
    Public Const frmCountryMaster As String = "COUN-MASTER"
    Public Const frmReligionMaster As String = "RELI-MASTER"
    'Public Const frmCurrencyMaster As String = "CURR-MASTER"
    Public Const frmCastCategoryMaster As String = "CAST-MASTER"
    Public Const frmGradeMaster As String = "GARD-MASTER"
    Public Const frmStateMaster As String = "STAT-MASTER"
    Public Const frmDevisionMaster As String = "DEVI-MASTER"
    Public Const frmOccupationMaster As String = "OCCU-MASTER"
    Public Const frmPFRulesMaster As String = "PF-RULES-MAS"
    Public Const frmESIRulesMaster As String = "ESI-RULES-MA"
    Public Const frmOTMaster As String = "OT-MASTER"
    Public Const frmAttendanceMaster As String = "ATTE-MASTER"
    Public Const frmLICPolicyMaster As String = "LIC-MASTER"
    Public Const frmBranchMaster As String = "BRANCH-MAST"
    Public Const frmBonusMaster As String = "BONUS-MASTE"
    Public Const frmLeaveMaster As String = "LEAVE-MAST"
    Public Const frmOTSlab As String = "OT-Slab"
    Public Const frmPTSlab As String = "PT-Slab"
    Public Const frmHRRSlab As String = "HRR-Slab"
    Public Const frmSavingsMaster As String = "SAVING_MAST"
    Public Const frmSectionAllowanceMaster As String = "SECALW_MAST"
    Public Const frmEmployeeSavingsMapping As String = "EMP_SAV_MAP"

    Public Const IncomeTaxTDSCalculation As String = "INC-TAX-CAL"

    Public Const frmConveyanceRateMaster As String = "Conv-Rate"
    Public Const frmODMaster As String = "OD-Master"
    Public Const frmSubDepartmentMaster As String = "SUB-DEP-MST"
    Public Const frmPayrollDesignationMaster As String = "PRoll-Desg"
    Public Const frmPaymentMode As String = "PROLL_PAY"
    '------------------Transactions----------------------------
    Public Const frmOTSheet As String = "OT-SHEET"
    Public Const frmLeaveSetting As String = "LEAVE-SETT"
    Public Const frmGeneralHolidays As String = "GEN-HOLIDAY"
    Public Const frmLeaveOpeningBalance As String = "LEAVE-OB"
    Public Const frmLeaveStartingDateSetting As String = "LV-ST-DT-SET"
    Public Const frmLeaveAllotment As String = "LEAVE_ALLOT"
    Public Const frmPayHeadDefinitions As String = "PAY_HEAD_DEF"
    Public Const frmSalaryStructure As String = "SAL_STRUCT"
    Public Const frmMapPayHeadsToSalaStructure As String = "MAP_PH_SS"
    Public Const frmWeeklyHolidays As String = "WEEKLY_HOL"
    Public Const frmLeaveApplication As String = "LEAVE_APPLI"
    Public Const frmMonthlyAttendance As String = "MONTH_ATTEND"
    Public Const frmLeaveAdjustment As String = "LEAVE_ADJUS"
    Public Const frmGenerateBonus As String = "GEN_BONUS"
    Public Const frmDailyAttendance As String = "DAILY_ATTEND"
    Public Const frmHourlyAttendance As String = "HOURL_ATTEND"
    Public Const frmAdjustmentVoucher As String = "EMP_ADJUST"
    Public Const frmReimbursementDetails As String = "EMP_REIMBURS"
    Public Const frmApplyLoan As String = "APPLY_LOAN"
    Public Const frmArrear As String = "Arrear"
    Public Const frmDA_Arrear As String = "DA_Arrear"
    Public Const frmLeaveEncashment As String = "L_Encashment"
    Public Const frmLICEmployeeTagging As String = "LIC-EMP-TAG"
    Public Const frmEmployeeOTCalculation As String = "EMP-OT-CAL"
    Public Const frmAllowanceDetails As String = "ALLOWANCE"
    Public Const frmDeductionDetails As String = "DEDUCTION"
    Public Const frmEmployeeStatus As String = "EMP_Status"
    Public Const frmEmpSalary As String = "EMP_Salary"
    Public Const frmLoanAdjustment As String = "LOAN_ADJUST"
    Public Const frmEmployee_Master As String = "EMP_MAST"
    Public Const frmLoanGeneration As String = "LOAN_GNTR"
    Public Const frmEmployeeIncrement As String = "EMP_INCR"
    Public Const frmSalaryGeneration As String = "SAL_GEN"
    Public Const frmEmployeeGratuity As String = "EMP_GRATITY"
    Public Const frmLTAClaim As String = "LTA_CLAIM"
    Public Const frmMediclaimEntry As String = "MED-CLM"
    Public Const frmFullAndFinalSettlement As String = "FF-Settle"
    Public Const frmEmployeeShiftChange As String = "Shift-Change"
    Public Const frmODSheet As String = "OD-SHEET"
    Public Const frmConveyanceClaim As String = "Conv-Claim"
    Public Const frmPayrollSetting As String = "PAY-SETT"
    Public Const FrmEmployeeTransfer As String = "EMP-TRAN"
    Public Const frmEPFund As String = "EMP-EPF"
    Public Const FrmEmpIncrement As String = "EMP-INCR"
    Public Const FrmSentSalarySlip As String = "SENT_SAL"
    Public Const FrmAllotmentOfLeaves As String = "ALLOT_LEAV"
    'Public Const frmEmployee_Salary As String = "EMP_GRATITY"
    '------------------Transaction End--------------------------

    '------------------Reports--------------------------
    Public Const frmSalaryGenerationRegister As String = "SAL_GEN_REG"
    Public Const frmSalaryGenerationRegisterArrear As String = "SAL_GEN_REGA"
    Public Const frmAllownceRegister As String = "ALLOWN_REG"
    Public Const frmDeductionRegister As String = "DEDUCT_REG"
    Public Const frmReimbursementRegister As String = "REIMBU_REG"
    Public Const frmAdjustmentRegister As String = "ADJUST_REG"
    Public Const frmAttendanceRegister As String = "ATTEND_REG"
    Public Const frmLeaveRegisterReport As String = "LEAVE_REG_RE"
    Public Const frmEmployeeRegister As String = "EMP_REG"
    Public Const frmPF_ESI_Reports As String = "PF_ESI_REP"
    Public Const frmPF_Covering_Letter As String = "PF_COVER_LTR"
    Public Const frmMonthlyESI_Report As String = "MNTH_ESI_RPT"
    Public Const RptEmployeeBday6 As String = "EMP_BDAY"
    Public Const rptEmployeeAdvanceLedger As String = "EMP-ADV-LED"
    Public Const rptLoanStatement As String = "RPT-LON-STM"
    '----------------------------Monthly------------------------------
    Public Const frmPaySlip_Reports As String = "PAYSLIP_REPO"
    Public Const frmSalarySheet_Reports As String = "SALARY_SHEET"
    Public Const frmSalaryAbstractReport As String = "SALaBSTRACT"
    Public Const frmAnnualCensusReport As String = "ANN_CENS_RPT"
    Public Const frmAttendedDaysReport As String = "ATTD_DAYS"
    Public Const frmSalaryVoucher_Reports As String = "SAL_VOUCH"
    Public Const frmBankStatement_Reports As String = "BANK_STATE"
    Public Const frmOT_Reports As String = "OT_REPORT"
    Public Const frmVarianceReport As String = "Variance"
    Public Const frmSalaryComponentDetails As String = "SALCOMPONENT"
    Public Const frmAditionalEarning_DeductionReport As String = "ADD_ER_DU"
    Public Const frmMonthlyPF_Report As String = "MONTHLY_PF"
    Public Const frmSalaryCertificate As String = "SAL_CERT"
    Public Const frmSalaryIncrementReport As String = "SAL_INCR_REP"
    Public Const frmBankSummary_Report As String = "BANK_SUMMARY"
    Public Const frmDeductionDetailsReport As String = "DED_DETAIL"
    Public Const frmForm5_PF As String = "FORM5_PF"
    Public Const frmVoucherPaymentsRegister As String = "VOU_PAY_REG"
    Public Const frmEmp_Id As String = "EMP_ID"
    Public Const frmFormA3_Report As String = "FORMA3_RPT"
    Public Const frmFormA6_Report As String = "FORMA6_RPT"
    Public Const frmForm9A As String = "Form9A"
    Public Const frmLabelPrinting As String = "LABEL_PRINT"
    Public Const frmESIOnline As String = "ESI_ONLINE"
    Public Const frmForm_T As String = "Form_T"
    Public Const FrmSalarySlipRpt As String = "Sal_Slip_rpt"
    Public Const FrmSalarySummaryRpt As String = "Sal_Smry_rpt"
    Public Const FrmEmployeePFRpt As String = "Emp_PF_rpt"
    Public Const frmESICRpt As String = "ESIC_Rpt"
    Public Const frmNewSalCertificate As String = "New_SAL_CERT"
    Public Const RptESIHalfYearly As String = "ESI_HALF_YR"
    Public Const RptESICStatement As String = "ESIC_STAT"
    Public Const RptESICForm6 As String = "ESIC_FRM6"
    Public Const RptESICChallan As String = "ESIC_CHALLAN"
    Public Const RptESICDeclarationForm As String = "ESIC_DECLA"
    Public Const RptPFForm2 As String = "PF_FRM2"
    Public Const RptPFForm3A As String = "PF_FRM3"
    Public Const RptPFForm5 As String = "PF_FRM5"
    Public Const RptPFForm6 As String = "PF_FRM6"
    Public Const RptPFStatement As String = "PF_STAT"
    Public Const RptPFForm10 As String = "PF_FRM10"
    Public Const RPtPFForm11_Revised_ As String = "PF_FRM11"
    Public Const RptPFForm12A_revised_ As String = "PF_FRM12"
    Public Const RptPFChallanStatement As String = "PF_CHALL"
    Public Const RptKDILSalarySlip As String = "SAL_SLIP"
    Public Const RptBOILetterReport As String = "BOI_LETT"
    Public Const RptForm22 As String = "FRM_22"
    Public Const RptBonusStatement As String = "BONU_STAT_R"
    Public Const RptForm34 As String = "BONU_STAT"
    Public Const RptPerformaForContributiondetail As String = "Per_Con_Det"
    Public Const RptBankTransferDetail As String = "BNK_TRN_DET"
    Public Const FrmDepartmentwiseSalarySheetRpt As String = "Dept_Sal_Sht"
    Public Const rptDetailOfWelfareFundAmount As String = "WEL_FUD_AMT"
    Public Const rptFromNO21 As String = "FRM_21"
    Public Const RptActurialValuation As String = "ACTU_VALUA"
    Public Const RptAttendanceReport As String = "ATT_RPT"
    Public Const rptBiometricAttendanceReport As String = "BIO_ATT_RPT"
    Public Const rptEmployeeStatusReport As String = "EMP_STS_RPT"
    '------------------Reports End--------------------------
    '------------------HR AND PAYROLL End--------------------------

    '------------------PRODUCTION----------------------------------'
    '------------------Master----------------------------------------'
    Public Const ACCSETMFGSTD As String = "ACCSETMFGSTD"
    Public Const ACCSETMFGDairy As String = "ACCSETMFGD"
    Public Const ACCSETMFGPepsi As String = "ACCSETMFGPEP"
    Public Const Template As String = "Template"

    Public Const ITEMCATEGORY As String = "ITEMCATEGORY"
    Public Const TOOL As String = "TOOL"
    Public Const SETTSTD As String = "SETTSTD"
    Public Const SETTPep As String = "SETTPEP"
    Public Const COSTMAINTAIN As String = "COSTMAINTAIN"
    Public Const frmWorkCenterMaster As String = "WRK_CEN_MAST"
    Public Const frmResourceMaster As String = "RESOURC_MAST"
    Public Const frmBillOfMaterialPepsi As String = "BOM_Pep"
    Public Const frmBillOfMaterialDairy As String = "BOM_Dairy"
    Public Const frmProfitCenter As String = "PRO-CENTER"
    Public Const frmOperationMaster As String = "OPER_MAST"
    Public Const frmBOMImport As String = "BOM_IMPORT"
    Public Const TOOLTYPE As String = "TOOLTYPE"
    Public Const EXPENSE As String = "EXPENSE"
    Public Const PRO As String = "PRO"
    Public Const ALTER As String = "ALTER"
    Public Const FrmProcessMaster1 As String = "PROCESS-M"
    Public Const FrmSectionMaster As String = "SEC-M"
    Public Const FrmStageMaster As String = "STG-M"
    Public Const FrmSectionStageMapping As String = "S-S-M"
    Public Const frmProcessProductionLogSheet As String = "PP_LOGSHEET"
    Public Const frmPPStageProcessQCLogSheet As String = "PP_STG_QC"
    Public Const frmPPLogSheetMaster As String = "PP_LOG_MST"
    Public Const FrmProductionLines As String = "PRO_LINE"
    Public Const frmBreakDownMaster As String = "PRO_BD_MST"
    '------------------Transaction----------------------------------------'
    Public Const frmBreakDownEntry As String = "PRO_BD_ENT"
    Public Const frmStanderdProductionEntry As String = "STD_PRO_ENT"
    Public Const FrmBagReceipt As String = "BAG_RCPT"
    Public Const frmBillOfMaterialCosting As String = "BOMCOSTING"
    Public Const frmProductionPlanningSTD As String = "PRODPLAN_STD"
    Public Const frmProductionPlanningDairy As String = "PRODPLAN_D"
    Public Const frmProductionPlanningPepsi As String = "PRODPLAN_PEP"

    Public Const frmProductionRequisition As String = "PROD_REQUIS"
    Public Const frmManufacturingOrder As String = "ProdMO"
    Public Const frmBatchOrderSTD As String = "Prod_BO_STD"
    Public Const frmBatchOrderDairy As String = "Prod_BO_D"
    Public Const frmProductionStoreRequest As String = "PROD_STOR_RQ"
    Public Const frmBatchOrderPepsi As String = "Prod_BO_Pep"

    Public Const frmStoreIssueSTD As String = "PD_ISSUE_STD"
    Public Const frmStoreIssuePepsi As String = "PD_ISSUE_Pep"
    Public Const frmProductionReturnSTD As String = "PD_RTN_STD"
    Public Const frmProductionReturnPep As String = "PD_RTN_PEP"
    Public Const frmProductionReceiptSTD As String = "PROD_REC_STD"
    Public Const frmProductionReceiptPepsi As String = "PROD_REC_PEP"

    Public Const frmLabourWorkingSheet As String = "LBR-WRK-SHT"
    Public Const frmDemoProductionPlanning As String = "PRODUCT_PLAN"
    Public Const frmProcessProductionIssueEntry As String = "PP_ISSUE"
    Public Const frmProcessProductionStandardization As String = "PP_STDN"
    Public Const ProcessProductionStandardizationFinalQC As String = "PP_STD-FQC"
    Public Const frmProductionReceiptDemo As String = "PROD_RCPT"
    Public Const frmProductionItemSerialReplace As String = "PRD_SR_RPLC"
    Public Const frmProcessProductionStageProcess As String = "PRD_STG_PROC"
    Public Const frmProductionEntry As String = "PROD_ENTRY"
    Public Const frmProductionEntryFinalQC As String = "PRO-ENT-FQC"
    Public Const DariyProductionUploader As String = "DRY-PRO-UPL"
    Public Const ProductionShiftMgmt As String = "PRO-SFT-MGM"
    Public Const frmWreckageBooking As String = "PROD_WR"
    Public Const frmProductionEntryWithoutBatch As String = "PROD_ENTR_WB"
    Public Const frmProcessProdReturn As String = "PP_PROD_RET"
    Public Const frmSiloMilkTransfer As String = "PP_SILO_MTR"
    Public Const frmSiloMilkTransferUploader As String = "UP_SILO_MTR"
    'Public Const frmSiloMilkTransfer_JOBWORK As String = "JW_SILO_MTR"
    Public Const CustomerIncentiveEntry As String = "CUST_INC_ENT"
    'Public Const frmMRP As String = "MRP"
    Public Const frmMRPAutoMobile As String = "MRP_Auto_Ind"
    Public Const rptproductionEntryReport As String = "Pro_ENT_RPtT"
    Public Const rptProductionIssueStatusReport As String = "PRO-ISS-STS"
    Public Const rptItemConsumptionReport As String = "ITM-CON-RPT"
    Public Const rptWTPReport As String = "PRO-WIP-RPT"
    Public Const rptBatchStatus As String = "PRO-BAT-STA"
    Public Const rptSectionWiseStockReport As String = "Pro_SEC_STK"
    Public Const rptDairyProductionWreckageReport As String = "Pro_WRE_RPT"
    'Public Const rptJobWorkProduction As String = "PROD_JW_RPT"
    Public Const rptAvailableQtyForProduction As String = "PROD_QTY_RPT"
    Public Const RptStandardQCReport As String = "STD_QC_RPT"
    Public Const rptStandardActualConsumption As String = "STD_AC_RPT"
    Public Const RptProductionQCReport As String = "PRD_QC_RPT"
    Public Const rptProductionWiseStockReco As String = "PRD_STK_RCO"
    Public Const rptIssueWIPConsumptionReport As String = "ISU_WIP_CON"
    Public Const rptProcessGainLossReport As String = "RPR_GAN_LOC"
    '------------------End of transaction----------------------------------------'

    '------------------Report----------------------------------------'
    Public Const Resource As String = "Resource"
    Public Const LTOOL As String = "L-TOOL"
    Public Const LACCt As String = "LACCt"
    Public Const PRODREPORT As String = "PRODREPORT"
    Public Const LToolT As String = "LToolT"
    Public Const LALT As String = "LALT"
    Public Const LOPER As String = "LOPER"
    Public Const LWC As String = "LWC"
    Public Const LOIC As String = "LOIC"
    Public Const frmListOfBOM As String = "ListOfBOM"
    Public Const frmIssueReturnItemWiseReportSTD As String = "IRItmRptSTD"
    Public Const frmIssueReturnItemWiseReportPepsi As String = "IRItmRpPEP"
    Public Const frmDatewiseProduction As String = "DWProdrpt"
    Public Const frmProductionPlanReportSTD As String = "PPRPT_STD"
    Public Const frmProductionPlanReportPepsi As String = "PPRPT_Pep"
    Public Const frmLineProductivity As String = "lineProdRPT"
    Public Const frmBatchOrderReportSTD As String = "BORPT_STD"
    Public Const frmBatchOrderReportPepsi As String = "BORPT_PEP"
    Public Const frmGraphicalBatchOrder As String = "GraphBORpt"
    Public Const frmGraphicalCategorywiseProduction As String = "GraphCatProd"
    Public Const frmListofRequisition As String = "Lst_Requis"
    Public Const frmProductionVarianceSTD As String = "PD_Vari_STD"
    Public Const frmProductionVariancePepsi As String = "PD_Vari_Pep"
    Public Const frmOperaterEfficiencyReport As String = "OPR-IFF-RPT"
    Public Const frmProductionSerializedReport As String = "PRD_SR-RPT"


    '------------------End of Report----------------------------------------'
    '-----------------END MASTER-------------------------------------'

    '-------------------END PRODUCTION------------------------------'



    ''--------------------------------------------------Sales New--------------------------------------------------------------------------------------------
    '------------------Masters-------------------------
    Public Const frm_User_Customer_Rate_Settings As String = "UC-RT-STNG"

    Public Const customerItemDetails As String = "CST-ITM-DTL"
    Public Const CustomerItemDetailApproval As String = "CST-ITM-DEA"
    Public Const frmSalesmanTarget As String = "SLSMN-TRGT"
    Public Const frmItemPriceListLevel3 As String = "ITM-PRC-L3"
    Public Const frmPendingQuotationApproval As String = "PND-QT-APRVL"
    Public Const FrmApprovalSetting As String = "APPROVAL-SET"
    Public Const FrmSaleSetting As String = "SALE-SETT"
    Public Const customerItemMapping As String = "CST-ITM-MAP"
    Public Const frmSNShipmentImportExport As String = "SALN-SP_IE"
    Public Const FrmRemarMaster As String = "REMARK-M"

    '-----------------Transactions---------------------
    Public Const frmSaleQuotation As String = "SAL-QUO"
    Public Const frmSNSalesOrder As String = "SALN-SO"
    Public Const frmSNShipment As String = "SALN-SP"
    Public Const frmSNSaleInvoice As String = "SALN-SI"
    Public Const frmSNServiceInvoice As String = "SERVICE-SI"
    Public Const frmSNSaleReturn As String = "SALN-SR"
    Public Const frmSNPOS As String = "SALN-OPS"
    Public Const frmSNReceiptChallan As String = "DL-Receipt"
    Public Const FrmAutoSTN As String = "AUTO-STN"
    Public Const frmCrptFooter As String = "CRPT_FOOT"
    Public Const FrmProspect As String = "Prospect"


    Public Const rptSalesmanTarget As String = "SAL-TGT-RPT"
    Public Const frmSaleOrderDetail As String = "SAL-ORD-DET"
    Public Const FrmSaleInvoiceSummary As String = "SAL-INV-SUMM"
    Public Const FrmSaleInvoiceDetail As String = "SAL-INV-DET"
    Public Const FrmShipmentSummary As String = "SAL-SHP-SUMM"
    Public Const FrmShipmentDetail As String = "SAL-SHP-DET"
    Public Const FrmSaleRegisterDemo As String = "SAL-REGD-RPT"
    Public Const ORDNEW As String = "ORDNEW"
    Public Const frmReceiptChallanReport As String = "REC-CHA-REP"
    'Public Const rptDetailedCardReport As String = "DET-CAR-REP"
    Public Const frmPendingSaleInvoiceforChilpPO As String = "PEN-INV-DEMO"
    Public Const rptSaleRegisterForAdv As String = "SAL-REG-Adv"
    Public Const rptSalesCancelledRegister2 As String = "SAL-CNL-REG"
    '------------------Fresh Sale----------------------------------'
    '------------------Master----------------------------------------'
    Public Const FrmZoneMaster As String = "ZONE-M"
    Public Const FrmCreditLimitApproval As String = "CRD-LIM_APPR"
    Public Const FrmFreshCreditLimitApproval As String = "FCRD-LIM_APP"
    Public Const FrmExpiryDate As String = "EXP_D"
    Public Const FrmTransactionApproval As String = "TRANS_APP"
    Public Const frmRouteMaster As String = "ROUTE-MAS"
    Public Const FrmPrintFreshInvoice As String = "PRT_FRESH"
    Public Const FrmFreshTransactionApproval As String = "FTRAN_APP"

    Public Const frmSalesHierarchy As String = "SALES-HIER"
    Public Const frmSalesLevelHierarchy As String = "SLS_Lvl_HIER"
    Public Const frmSalesHierarchyMapping As String = "SLS_HIER_MAP"

    '------------------Transaction----------------------------------------'
    Public Const FrmBookingEntry As String = "BOOK-ENTRY"
    Public Const frmDeliveryNoteFreshSale As String = "DEL-NOTE-FS"
    Public Const FrmDispatchFreshSale As String = "DISPATCH-FS"
    Public Const frmDispatchMultipleFreshSale As String = "DISP-MUL-FS"

    Public Const frmInvoiceFreshSale As String = "INVOICE-FS"
    Public Const frmCreateReceived As String = "CRATE-REC-FS"
    Public Const frmSaleReturnFreshSale As String = "SALE-RET-FS"
    Public Const FrmGatePassFS As String = "GATE-PASS-FS"
    Public Const FrmShortCloseDO As String = "SH-CL-DO"
    Public Const FrmInvoiceCrateLinerDetail As String = "INV-C-R-DT"
    '---------------Report-------------------------------------------------'
    Public Const RptFreshSaleRegister1 As String = "F_SALE_REG"
    Public Const rptCrateAccounting As String = "CRATE_ACC_FS"
    Public Const rptCrateAccountingReport As String = "CRTE_ACC_RP"
    Public Const RptVehicleCapacityFreshSaleReport As String = "VEH_CAP_FRS"
    Public Const RptVehicleWiseReport As String = "VEH_WISE"
    Public Const RptFreshBookingStatus As String = "FRS_BOK_STS"
    Public Const rptSaleReturnGateEntryReport As String = "RPT_SAL_GRE"
    Public Const frmSubsidyCreditNote As String = "SUBSIDY-FS"

    ''Prabhakar
    Public Const rptZoneWiseFreshSaleReport As String = "FRS_ZON_WIS"
    Public Const rptDispatchChallanReportFresh As String = "FRS_DIS_CHN"
    Public Const rptMatrixFreshSalesReport As String = "FRS_MTX_RPT"
    Public Const rptPriceChartFreshSalesReport As String = "FRS_PRC_RPT"
    Public Const EmployeeWiseReport As String = "EMP_WSE_RPT"
    Public Const rptCrateLinerReport As String = "CRT_LIN_RPT"

    '------------------Bulk Sale----------------------------------'
    '------------------Master----------------------------------------'
    Public Const FrmBulkCreditLimitApproval As String = "BCRD-LIM_APP"
    Public Const FrmBulkTransactionApproval As String = "BTRAN_APP"
    '------------------Transaction----------------------------------------'
    Public Const FrmGateEntrySale As String = "GATE-ENTRY-S"
    Public Const FrmSalesOrderBS As String = "SO_BS_PAVTRA"
    '' Richa 
    Public Const FrmBulkSalePriceChart As String = "BULK-SALE-PC"
    Public Const FrmWeighmentEntry As String = "WEIGH-ENTRY"
    Public Const FrmLoadingTanker As String = "LOAD-TANKER"
    Public Const FrmQualityCheckBulkSale As String = "QUA-CHK-BS"
    Public Const FrmDispatchBulkSale As String = "DISPATCH-BS"
    Public Const frmBulkSaleAcknowledgement As String = "BLK-SL-ACK"
    Public Const frmBulkSaleAcknowledgementUploader As String = "BLkSL-AC-UP"
    Public Const frmBulkSaleFreightCalculation As String = "BkSL-FR-CL"
    Public Const FrmBulkSaleSettings As String = "SALE-SET-NEW"
    Public Const FrmInvoiceBulkSale As String = "INVOICE-BS"
    Public Const FrmCreateAutoInvoiceBS As String = "AUTO-IN-BS"
    Public Const FrmTankerOut As String = "TANK-OUT-BS"
    'Public Const FrmDispatchBulkSaleTrade As String = "DIS-TRADE-BS"
    'Public Const FrmDispatchBulkSaleTradeReturn As String = "DS-TRD-BS_RE"
    Public Const FrmBulkSaleReturn As String = "BULK-SALE-RE"
    Public Const FrmCanSaleUploader As String = "CAN-SALE-UP"
    Public Const FrmCanSale As String = "CAN-SALE"
    Public Const FrmCanReceived As String = "CAN-RCVD"
    Public Const FrmBulkDispatchReturnSale As String = "BULK-DIS-RE"
    '========================Report=======================
    Public Const RptBulkSaleRegister As String = "BUL-SL-REG"

    '------------------Merchant Trade----------------------------------'
    '------------------Master----------------------------------------'
    Public Const FrmFixedDeposit As String = "FXD-DEPOSIT"
    Public Const FrmLCRequest As String = "LC-REQUEST"
    Public Const FrmLCCreation As String = "LC-CREATION"
    Public Const FrmMerchantPaymentTerms As String = "MT-PAY-TERMS"
    Public Const FrmMerchantPaymentTermsGroup As String = "MT-PAY-TGRP"
    Public Const FrmMTReportContextFormat As String = "MT-RPT-CON"
    '------------------Transaction----------------------------------------'
    Public Const FrmPurchaseOrderMT As String = "PO-ODR-MT"
    Public Const frmProformaInvoiceMT As String = "PE_IN_MT"
    Public Const frmSalesOrderMT As String = "SALE_ORDR_MT"
    Public Const frmCommercialInvoiceMT As String = "COMM_IN_MT"
    Public Const frmSalesInvoiceMT As String = "SALE_IN_MT"
    Public Const frmSalesReturnMT As String = "SALE_RE_MT"
    Public Const FrmDocumentAcceptance As String = "DOC-ACCEP-MT"
    Public Const FrmSRNMT As String = "MT_PO_SRN"
    'Public Const FrmBulkCloser As String = "BULK_CLOSER"
    '========================Reports========================================'
    Public Const RptCashAgainstDocs As String = "CASH-DOC"
    Public Const frmMTSalesCancelledTransation As String = "MT_CAN_REP"

    '------------------Merchant Trade----------------------------------'
    '------------------Master----------------------------------------'
    Public Const frmBookingProductSale As String = "BOOK-PS"
    Public Const frmDeliveryPrderProductSale As String = "DEL-ORD-PS"
    Public Const frmdispatchAdviceProductSale As String = "DIS-ADV-PS"
    Public Const frmSalesOrderProductSale As String = "SALE-ORD-PS"
    Public Const frmShipmentProductSale As String = "SHIPMENT-PS"
    Public Const frmSaleInvoiceProductSale As String = "INVOICE-PS"
    Public Const frmSaleReturnProductSale As String = "SALE-RET-PS"
    Public Const FrmGatePassPS As String = "GATE-PASS-PS"
    Public Const frmRouteFreightDetails As String = "R-F-D"
    Public Const FrmSchemeMasterDairy As String = "SCH-MAS-DAI"
    Public Const FrmDistanceMappingMaster As String = "DIS_MAP_M"
    Public Const FrmPrintProductInvoiceStatement As String = "PRT_PRODC"
    Public Const FrmPrintBulkInvoiceStatement As String = "PRT_BULK"
    Public Const RptProductSaleRegister1 As String = "P_SALE_REG"
    Public Const FrmProductDispatchGateOut As String = "PRD-DSP-GO"
    '====Parteek Added 10-01-2017
    Public Const frmProductBooking As String = "PR-BOOK"
    Public Const frmMultCustBookingDisp As String = "MutCustDis"

    '------------------Transaction----------------------------------------'
    '==========Report===================
    Public Const RptProductBookingStatus As String = "P_BO_STUS"
    Public Const RptProductDispatchStatus As String = "P_DI_STUS"
    Public Const RptProductDOStatus As String = "P_DO_STUS"
    Public Const RptProductSaleOrderStatus As String = "P_SO_STUS"
    Public Const RptBulkMultipleDispatch As String = "DIS_BULK_R"


    '=================Export Sale========================================
    '===========Master
    Public Const frmEnquiryMaster As String = "ENQ-MST"
    Public Const frmNotifiedPartyMaster As String = "NOT-PAT-MST"
    Public Const frmCHAChargeMaster As String = "CHA-CHRG-MST"
    Public Const frmExIncentiveMaster As String = "Ex-INCTV-MST"
    '===================transaction====================
    Public Const frmEXSalesQuotation As String = "EX_SALE_QUTN"
    Public Const frmEXSalesOrder As String = "EX_SALE_ORDR"
    Public Const frmEXSalesOrderR As String = "EX_S_ORDR_R"
    Public Const frmEXPorformaInvoice As String = "EX_PI"
    Public Const frmEXCommercialInvoice As String = "EX_COM_IN"
    Public Const frmEXSalesInvoice As String = "EX_SALE_IN"
    Public Const frmEXSalesReturn As String = "EX_SALE_RETN"
    '=========================Report===========================================
    Public Const RptExportSaleRegister As String = "EX_SALE_REG"
    Public Const RptMTPIInHead As String = "MT-PI-HEAD"
    Public Const RptEXProductWiseDetail As String = "EX-PW-DET"
    Public Const RptdateOfArrivalofCons As String = "EX-DAT_ARR"
    Public Const RptPaymentRelization As String = "EX-Pay_Rel"
    Public Const frmEXSalesCancelledTransation As String = "EX_CAN_REP"
    '==============================================================

    '=====CSA Sale===============================================
    '-----------Master----------------------------
    Public Const frmCSAPriceMaster As String = "CSA-PRC-MST"
    Public Const frmCSAAccountSet As String = "CSA-AC-SET"
    Public Const frmCSACommissionItemWise As String = "CSA-CMSN-MST"
    '-------------Transaction------------------------------
    Public Const frmCSADeliveryOrder As String = "CSA-DO-TRN"
    Public Const frmCSASaleInvoice As String = "CSA-INV-TRN"
    Public Const frmCSATransferReturn As String = "CSA-TRN-RTN"
    Public Const frmCSASalePattiReturn As String = "CSA-INV-RTN"
    '=================Report====================
    Public Const frmCSADOReport As String = "CSA-DO_RPT"
    Public Const frmCSATransferReport As String = "CSA-Trns-RPT"
    Public Const RptCSASaleRegister As String = "CSA-SALE-REG"
    Public Const RptCSACustomerLedger As String = "CSA-CUST-LED"
    Public Const RptCSAmonthlywisereport As String = "CSA-MW-ARPT"
    '-===================================================

    '------------------PROJECT MANAGEMENT----------------------------------'
    '------------------Master----------------------------------------'
    Public Const frmPJCSettings As String = "PJCSETT"
    Public Const frmCostTypes As String = "CostTypes"
    Public Const frmPJCAccountSets As String = "PJCACCSet"
    Public Const frmJobMaster As String = "PJCJOB"
    Public Const frmTaskMaster As String = "PJCTask"
    Public Const frmPJCEmployeeMaster As String = "PJCEMPLOYEE"
    Public Const FrmUserApproval As String = "USER-APPR"
    Public Const ProjectMaster As String = "PJ-MAS"
    Public Const FrmBudgetMaintenance As String = "BUDGET-MAIN"
    Public Const FrmExpenseType As String = "EXPENSE-TYPE"
    Public Const FrmProjectStatus As String = "PROJ-STATUS"
    '------------------Transaction----------------------------------------'
    Public Const frmTimeSheet As String = "PJCTIMESHEET"
    Public Const frmUserLog As String = "PJC-USR-LOG"
    Public Const frmAssemblies As String = "PJC-Assembly"
    Public Const frmAssembDis As String = "Assembly"
    Public Const FrmPJCExpense As String = "PJC-EXPENSE"
    '------------------Report----------------------------------------'
    Public Const frmProjectListReport As String = "PJC-PRJ-LST"
    Public Const frmProjectDetails As String = "PJC-PRJ-DTL"
    Public Const FrmProjectProfitReport As String = "PJC-PROFIT-R"
    '-------------------END PROJECT MANAGEMENT------------------------------'


    Public Const frmDealerManagementReport As String = "DEALER_RPT"
    Public Const frmProspectDetailReport As String = "ProspectRpt"
    '-------------------SERVICE-------------------------------------------'


    '--------------------------SETUP-------------------------------------'
    Public Const frmComplaintGroupMaster As String = "COMP-GR-MSTR"
    Public Const frmPrimaryReasonMaster As String = "PRM-REA-MSTR"
    Public Const frmComplaintMaster As String = "COMP-MASTER"
    Public Const frmPendingReasonMaster As String = "PDNG-RSN-MST"
    Public Const frmItemChargeCategoryMaster As String = "ITFR-CHR-MST"
    Public Const frmItemChargeFranchiseMappingMaster As String = "ITFR-MAP-MST"
    Public Const frmAssetServiceMaster As String = "ASS-SER-MST"

    '--------------------------TRANSACTION----------------------------------------'
    Public Const frmAssetAgreement As String = "ASSET-AGR"
    Public Const frmAssetInstallPullOut As String = "ASSET-IPOUT"
    Public Const frmComplaintDetailEntry As String = "COM-DET-TRN"
    Public Const frmAssetDistatch As String = "DISPATCH"
    Public Const frmCartMaintenanceEntry As String = "CRT-MNT-ENT"
    Public Const FrmPendingComplaintDetail As String = "PEN-COM-DET"
    Public Const frmQuickComplaintDetailEntry As String = "Q-COM-DET"
    '--------------------Report------------------------------------------------------'
    Public Const frmFranchiseChargesReport As String = "FRN-CHR-RPT"
    Public Const frmCustomersListReport As String = "FRN-CUST-RPT"
    Public Const frmPullOutRedeployReport As String = "PL-RDPL-RPT"
    Public Const frmVendorBillDetails As String = "VNDR-BL-DTL"
    Public Const frmDaywisePendingComplaint As String = "PND-COM-RPT"
    Public Const frmClaimReport As String = "CLAIM-RPT"
    'Public Const frmspareStockReport2 As String = "SPA-STO"
    Public Const frmAssetDetailReport As String = "AST-DET-RPT"

    '----------------SERVICE END-----------------------------------------------------------'

    '-------------------MILK PROCUREMENT-------------------------------------------'

    '--------------------------SETUP-------------------------------------'
    Public Const frmMilkTypeMaster As String = "MILK-TYPE"
    Public Const frmVehicleTypeMaster As String = "VEHICLE-TYPE"
    Public Const frmMilkCollectionLevels As String = "MC-LEVEL"
    Public Const frmMilkCollectionArea As String = "M-COLL-AREA"
    Public Const frmMilkRouteMaster As String = "MILK-ROUTE"
    Public Const FrmMilkRecurringScheduler As String = "REC_SCHEDULER"
    Public Const FrmMCCMilkTransPortorInvoice As String = "REC_INVOICE"
    Public Const frmMilkVehicleTypeMaster As String = "MILK-VH-TYPE"
    Public Const frmMilkVehicleMaster As String = "MILK-VEHICLE"
    Public Const frmMilkTransportRateMaster As String = "MILK-TR-RATE"
    Public Const frmMilkComponentMaster As String = "MILK-COMP"
    Public Const frmMilkComponentRateList As String = "M-COMP-RATE"
    Public Const frmMilkAdvanceMaster As String = "MILK-ADV"
    Public Const frmMilkRateTypeMaster As String = "M-RATE-TYPE"
    Public Const frmUOMMaster As String = "M-UOM"
    Public Const frmSeasonMaster As String = "MILK-SEASON"
    Public Const frmMilkShiftMaster As String = "MILK-SHIFT"
    Public Const frmVillageMaster As String = "VILL-MST"
    Public Const frmVSPMaster As String = "VSP-MST"
    Public Const frmPrimaryTransporterMaster As String = "PTM-MST"
    Public Const frmTankerTransporterMaster As String = "TTM-MST"
    Public Const frmMCCMaster As String = "MCC-MST"
    Public Const frmTankerMaster As String = "TNK-MST"
    Public Const frmVLCMaster As String = "VLC-MST"
    Public Const frmVLCMasterTarget As String = "VLC-MST-TGT"
    Public Const FrmPriceChartUploader As String = "PCU-MST"
    Public Const MilkPricePlanning As String = "PCM-PLN"
    Public Const GazeReading As String = "GAZ-RDN"
    Public Const frmHeadLoadMaster As String = "HED_LOD-MST"
    Public Const frmDeductionTypeMaster As String = "DED_TYP-MST"
    Public Const frmPriceChartMaster As String = "PCM-MST"
    Public Const frmVLCUploader As String = "VLC-UPL"
    Public Const frmMPMaster As String = "MP-MST"
    Public Const VerifyMPIFSC As String = "VFY-MP-IFSC"
    Public Const frmParameterMaster As String = "PM-MST"
    Public Const FrmContractTanker As String = "CONT-TANK-M"
    Public Const frmParameterRangeMaster As String = "PRM-MST"
    Public Const frmVLCRouteShiftMaster As String = "VRS-MST"
    Public Const frmCustomerRouteShiftMaster As String = "CRS-MST"
    Public Const frmPrimaryTransporterVehicalMaster As String = "PTV-MST"
    Public Const frmOpenMCCShift As String = "MCC-SHF"
    Public Const frmOpenMCCShiftManual As String = "MCC_MANUAL"
    Public Const frmPriceChartMaster_Bulk As String = "PCM_BLK-MST"
    Public Const frmParameterRangeMasterForQC As String = "P-RNG-QC"
    Public Const FrmPortSettings As String = "PORT-SET"
    Public Const frmPriceChartBulkProc As String = "P-CHRT-BLK"
    Public Const MCCSetting As String = "MCC-SETT-M"
    'Public Const MilkTruckSheet As String = "MCC-TRUK-SHT"
    'Public Const ApplyCapping As String = "MCC-APY-CAP"
    Public Const MilkVSPPayment As String = "MCC-VSP-P"
    'Public Const MilkMPPayment As String = "MCC-MP-P"
    Public Const MilkVSPIssuePayment As String = "MCC-VSP-IP"
    Public Const MilkRecurringScheduler As String = "Milk-RecShed"
    Public Const FrmVSPIncentiveTagging As String = "VSP-TAG"
    Public Const frmVSP_VLCMaster As String = "VSP_VLC-MST"
    '' Pankaj jha
    Public Const ParameterValueMaster As String = "PARM-VAL-M"
    Public Const LostDefectSealNo As String = "LST-DFCT-SN"
    Public Const LocationDistanceMapping As String = "L-DIS-MAP"
    Public Const MccMilkTransferPrice As String = "MCC-M-T-P"
    'Public Const SecondarySettingForQC As String = "SS-QC-BS"
    Public Const GenratePaymentCycle As String = "GEN-PAY-CYL"
    Public Const frmPaymentCycleMaster As String = "PY-CYCLE"
    Public Const frmIncentiveMaster As String = "Incentive"

    Public Const frmdeductionGroup As String = "DED_GRP"
    Public Const frmGroupOfDeduction As String = "GRP_O_DED"
    Public Const frmDeductionMaster As String = "DED_M"
    Public Const DeductionMapping As String = "DED_MAP"
    Public Const CanMaster As String = "CAN_MAS"
    Public Const DockMaster As String = "DOCK_MAS"
    Public Const FrmTankerDispatchPriceMaster As String = "TAK_DIS_PMST"
    Public Const MilkRejectType As String = "MKL-REJ-TYP"
    Public Const DCSAdditionDeduction As String = "DCS-ADD-DED"

    Public Const FrmMCCTransactionApproval As String = "MTRAN_APP"
    Public Const FrmMCCFarmerMapping As String = "MCC_FARMR"
    Public Const VLCMappingForMPAmount As String = "VLC_MPA"
    Public Const VLCMappingForMP_PP As String = "VLC_MPA_PP"
    Public Const VSP_Commsission As String = "VSP_COMM"
    Public Const VSP_Deduction As String = "VSP_DED"
    Public Const VSP_DayWiseIncetive As String = "VSP_DWIN"
    Public Const VSP_Mapping As String = "VSP_MAP"
    Public Const FarmerProMaster As String = "FRM_PRO_MST"
    Public Const CappingMaster As String = "CAP-Master"
    Public Const MPIncetiveSlab As String = "MP-INC-SLP"
    Public Const ChillingChargesSlab As String = "CHL-CHR-SLB"
    Public Const FrmOwnBmcExpanse As String = "OWN-BMC-EXP"
    Public Const OwnBMCGainLossRate As String = "BMC-RATE"
    Public Const FreightChargesMaster As String = "FRE_CHG"
    Public Const DCSFinancialHead As String = "DCS-FIN-HED"
    Public Const DCSFinancialEntry As String = "DCS-FIN-ENT"
    Public Const DCSSupervisorTagging As String = "DCS-SUP-TAG"
    ''''------------
    '--------------------------TRANSACTION-------------------------------------'
    Public Const frmMilkCollectionCenters As String = "MILK-MCC"
    Public Const frmMilkSuppliers As String = "MILK-SUPPL"
    Public Const frmMCCRouteMapping As String = "MCC-RT-MAP"
    Public Const frmMCCSuperwiserMapping As String = "MCC-SVR-MAP"
    Public Const frmMCCSupplierMapping As String = "MCC-SUPP-MAP"
    Public Const frmMilkCollection As String = "MILK-COLLECT"
    Public Const frmMilkQualityCheck As String = "M-QUALITY"
    Public Const frmMilkRateProcessingScheme As String = "M-R-PROC-SCH"
    Public Const frmVehicleMovement As String = "M-VH-MOVE"
    Public Const frmMilkBillGeneration As String = "M-BILL-GEN"
    Public Const MilkGateEntryIn As String = "M-G-IN"
    ' Public Const MilkRetesting As String = "M-R-Tst"
    Public Const MilkGateEntryWeightment As String = "M-G-WH"
    Public Const MilkGateEntryOut As String = "M-G-OUT"
    Public Const frmMilkReceipt As String = "M-RECEIPT"
    Public Const MilkReject As String = "M-REJECT"

    Public Const frmGateEntry As String = "M-GT-ENTR"
    'Public Const frmMCCGateEntry As String = "MCC-GE"
    'Public Const frmMCCWeighment As String = "MCC-WT"
    Public Const frmMCCDispatch As String = "MCC-DISP"
    Public Const frmAcknowledgementEntry As String = "ACK-ENTRY"
    'Public Const frmMCCTankerDispatchReturn As String = "TNK-DISP-RET"
    'Public Const frmMCCTankerGateOut As String = "MCC_TGT_OUT"
    'Public Const frmMCCTankerGateOutSecurity As String = "MCC_TGT_OTS"
    Public Const FrmLineMaster As String = "LINE_MST"
    'Public Const MCCDispatchReturn As String = "MCC-DISPR"
    Public Const frmWeighment As String = "M-WEIGHT"
    Public Const frmQualityCheck As String = "M-QC"
    Public Const frmMilkSample As String = "M-SAMPLE"
    Public Const frmMilkSRN As String = "M-SRN"
    Public Const frmMilkPurchaseInvoice As String = "M-PURINVOICE"
    Public Const FrmMilkPurchaseReturn As String = "M-PURRETURN"
    Public Const frmMilkShiftEndMCC As String = "M-Shift_End"
    Public Const frmMilkReasonMaster As String = "M-Reason"
    Public Const frmMCCMaterial As String = "M-Material"
    Public Const frmDCSSaleEntry As String = "DCS-SAL-ENT"
    Public Const frmMobileDCSDemand As String = "MOB-DCS-DMD"
    Public Const frmCancelDCSSale As String = "CNL-DCS-SAL"
    Public Const frmMCCMaterialSaleReturn As String = "M-Material-R"
    Public Const frmMCCMaterialFarmer As String = "M-MaterialF"
    Public Const frmMCCMaterialFarmerUploader As String = "M-MatFUP"
    Public Const frmMCCMaterialSaleReturnFarmer As String = "M-MatrialF-R"
    Public Const frmMCCMaterialSalePriceChart As String = "MCC_M_PRICE"
    Public Const frmMCCSMSSettiing As String = "MCC-SMS"
    Public Const frmCancelAfterPosting As String = "C-Posting"
    Public Const FrmVSPWiseDataReport As String = "VSP-Data-Rpt"
    Public Const rptMemberPaymentSlip As String = "Mem-Pay-slp"
    Public Const rptMPWiseMilkCollectionATPoolingPoint As String = "MP-Poll-PNT"
    Public Const rptMPWiseMilkCollection As String = "MP-WIS-CLL"
    Public Const rptPrimaryTransporter As String = "PRM-TRN_MCC"
    'Public Const rptMCCMilkRegisterDripSaver As String = "MCC-DRP_SAV"
    Public Const rptMCCVLCVarationReport As String = "MCC_VLC_VAR"
    Public Const rptVSPOrVLCVarationRpt As String = "MCC-VSP_VLC"
    Public Const RptVSPAssetIssue1 As String = "VSP_ASSET_I"
    'Public Const RptPriceRateDifferenceReport As String = "PRICE_DIFF"
    Public Const RptMCCMilkStatus As String = "MCC_MLK_ST"
    Public Const rptVSPIncentiveRegister As String = "VSP_INC_REG"
    Public Const rptRoutewiseTPTimeTable As String = "RPT-RT-TM-TL"
    Public Const rptVLCwiseTPTimeTable As String = "RPT-VLC-TMT"
    '================Grievance Module===================
    Public Const frmGrievanceTypeMaster As String = "H_Gri_Type"
    Public Const frmGrievanceLogging As String = "H_Gri_Logg"
    Public Const frmGrievanceAllocation As String = "H_Gri_Alloc"
    Public Const frmGrievanceResolution As String = "H_Gri_Resol"
    '===================================================
    '================Employee Equipment Tracking Module===================
    Public Const frmAssetTypeMaster As String = "H_Asset_Type"
    Public Const frmAssetCategoryMaster As String = "H_Cat_Mas"
    Public Const frmAssetSubCategoryMaster As String = "H_SCat_Mas"
    Public Const frmAssetMaster As String = "H_Asset_Mas"
    Public Const frmAssetIssueReturn As String = "H_Asset_IR"
    '===================================================
    Public Const frmUnloading As String = "M-UNLD"
    Public Const frmMilkTransferIn As String = "M-TR-IN"
    Public Const frmMilkTransferInReturn As String = "M-TR-INR"
    Public Const frmQCSeparation As String = "QC_SEPARATE"
    Public Const frmTankerProvision As String = "M-TANK-PROV"
    Public Const MilkCollectionGenerate As String = "MLK-COL-GEN"
    Public Const MilkCollectionMCCMultipleDays As String = "MLK-COL-MLD"
    Public Const MilkCollectionDCSMultipleDays As String = "MLK-DCS-MLD"
    Public Const AMCUMaster As String = "AMCU-MAS"
    Public Const WEIGHINGMaster As String = "WEI-Mas"
    Public Const frmGateEntryTransaction As String = "GT-ENT-TRNS"
    Public Const MilkCollectionDCSMultipleDaysMerge As String = "MLK-MLD-MRG"
    Public Const DCSMilkCollectionSetting As String = "MLK-CLN-STN"
    Public Const MilkCollectionMCC As String = "MLK-COL-MCC"
    Public Const MilkCollectionMCCGateEntry As String = "MLK-COL-MCCG"
    Public Const MilkCollectionMCCSample As String = "MLK-COL-MCCS"
    Public Const MilkCollectionDCS As String = "MLK-COL-DCS"
    Public Const frmBulkMilkSRN As String = "M-SRN-B"
    Public Const BulkProcurementUploader As String = "B-PRC-U"
    Public Const frmBulkMilkPurchaseInvoice As String = "M-PURINV-B"
    Public Const BulkMilkPurchaseInvoiceMultiple As String = "BM-PINV-M"
    Public Const frmCleaning As String = "M-CLEAN"
    Public Const frmGateOut As String = "M-GT-OUT"
    Public Const FrmTransferGateOut As String = "TRN-GT-OUT"
    Public Const frmBulkPurchaseUploader As String = "B-P-UP"
    Public Const rptTemporaryPaymentDeductionSummary As String = "TPDS_R"
    Public Const rptAutoMultipleAdditionDeduction As String = "AMAD-R"
    Public Const rptBmcCollection As String = "BMC-COL-RPT"
    Public Const CompareVoucherReport As String = "CMP-VCH-RPT"
    Public Const rptBankSavingAdvice As String = "BA-SAV-ADV"
    Public Const rptDBTMilkPayment As String = "DBT-MPAYM"
    Public Const rptDBTSummaryMonthlyWise As String = "DBT-Monthly"
    Public Const rptBMCTankerTestingReport As String = "BMCTan_Tes_R"
    Public Const rptMilkPaymentSummary As String = "Milk-PS"
    Public Const rptMilkUnion2 As String = "UN_MLK_RPT2"
    Public Const rptMilkUnion As String = "UNN_MLK_RPT"
    Public Const rptUnionMilkPurchase As String = "MLK_PURU_RPT"
    Public Const latestPaymentProcess As String = "LA_PA_PRS"
    Public Const rptCattleFeedSaleReport As String = "Cattle-F-S"
    Public Const rptTransferToSavingReport As String = "Tran-Sv-Rt"
    Public Const rptDcsPaymentReport As String = "Dcs-Pyt-Rpt"
    Public Const rptYearlyMonthlyDcsCollectionReport As String = "Yr-Mont-Dc"
    Public Const DcsMonthWiseMilkCollectionReport As String = "D_Mo_Wi_M_C"
    Public Const rptMilkProcurementReport As String = "MLK_PRC_RPT"
    Public Const frmBulkMilkSRNReturn As String = "M-BMSRN-R"
    'Public Const frmTranReverse As String = "M-TRN-R"
    'Public Const rptMDConversion As String = "MD_CONVERSIN"
    Public Const rptTruckSheetDailySummaryReport As String = "TRSH_DS_R"
    Public Const rptTankerStatusReport As String = "TNK_ST_RPT"
    Public Const rptTruckSheetReport As String = "TRU_SH_RPT"
    Public Const rptUnionBookingReport As String = "RPT_UNI_BOO"
    Public Const rptDailyQtyReport As String = "DAILY_QTY_R"
    Public Const rptBMCDCSAbsent As String = "BMC_DCS_RPT"
    Public Const rptTankerUtilization As String = "TANKER_ULT_R"
    Public Const rptBMCCapacityUtilizationReport As String = "BMC_CP_UTLZ"
    Public Const rptBMCPerodicalReport As String = "BMC_PERO_R"
    Public Const rptTankerProfitLossReport As String = "TNR_P_L_R"
    Public Const rptBmcTankerInOutReport As String = "BM_T_IO_R"
    Public Const rptPerodicalDispatchReport As String = "PERO_DISP_R"
    Public Const rptDBTCappingReport As String = "DBT_CAP_R"
    Public Const rptHeadLoadReport As String = "HEAD-LOAD-R"
    Public Const rptDcsSecurity As String = "DCS_SEC_RPT"
    Public Const rptDayWiseCol As String = "DAY_WSE_RPT"
    Public Const rptMarginAmt As String = "MRG_AMT_RPT"
    Public Const rptDCSCredit As String = "DCS_CRD_RPT"
    Public Const rptPaymentCycleWiseReport As String = "PYMT_CYCL_R"
    Public Const rptTempTruckSheetCollectionReport As String = "TEMP_TRCK_R"
    Public Const rptMobileAppMilkCollection As String = "MOB-MLK-R"
    Public Const rptBMCMobileHistory As String = "BMC-MOB-HIS"
    Public Const YearlyBillReport As String = "DCS-INC"
    Public Const FrmLastDCSCollectionItem As String = "Las-DCS-ITE"

    Public Const rptGheeAndCattleFeedDeductionStatementReport As String = "GH-CF-D-SMT"
    Public Const FrmMilkProcurementSocietySummary As String = "MILK-SUM"
    Public Const frmDBTRecoVsIncentiveReport As String = "DBTREC_INC_R"
    Public Const frmAutoAdditionDeductionReport As String = "AUTO_AD_R"
    'Public Const rptTankerStatusReport As String = "R-T-STAT"
    Public Const frmVendorPriceChartMapping As String = "V-PC-MPNG"
    Public Const ItemStockConversion As String = "I-S-CNV"
    Public Const frmVlcdataUploadar As String = "VLC-D-UPL"
    ''richa Against Ticket no BM00000003791 on 05/08/2014
    Public Const FrmVLCDataUploaderManual As String = "VLC-D-UPL-MA"
    Public Const MPIncentiveEntry As String = "MP-INC-ENT"
    Public Const DCSMPIncentiveReco As String = "DCS-MP-RECO"
    Public Const DBTMonthlyFarmerMilk As String = "DBT-MNT-FMLK"
    Public Const DBTCappingCheck As String = "DBT-CAP-CHK"
    Public Const DBTCappingIncrease As String = "DBT-CAP-INC"
    Public Const DBTNEFTUploader As String = "DBT-NEFT-UPL"
    Public Const DBTPayment As String = "DBT-PAY-MNT"
    Public Const SendDBTToJanaadhar As String = "SND-DBT-JAD"
    Public Const FrmCreateBMCDCSbyMobile As String = "BMC-DCS-MOB"
    Public Const DBTNEFTReject As String = "DBT-NEFT-REJ"
    ''=============================================
    '' Anubhooti 16-Sep-2014 
    Public Const frmVSPAssetIssue As String = "VSP-Issue"
    ''
    '==========Rohit=========
    'Public Const frmVSPItemIssue As String = "VSP-Item"
    Public Const frmProvisionEntry As String = "Prov-Entry"
    Public Const IncentiveEntry As String = "INC-ENT"
    Public Const frmPaymentProcess As String = "Pay-Pro"
    Public Const frmBankAdvise As String = "BNK-ADVS"
    Public Const frmTransferToSaving As String = "TRA-SAV"

    Public Const frmMakeSavingPayment As String = "MAK-SAV-PAY"
    Public Const frmTDSReport As String = "TDS-RPT"
    Public Const frmSendBillToDCS As String = "PMT-SND-DCS"
    Public Const frmSendSMSToDCS As String = "SMS-SND-DCS"
    Public Const frmReviseMilkBill As String = "REV-MLK-BLL"
    'Public Const frmDispatchTransfer As String = "M-DIS-TRAN"
    Public Const frmPaymentProcessFarmer As String = "Pay-Pro-Fa"
    Public Const BMC_Transporter_Bill As String = "BMC-TRN-BLL"
    Public Const frmTransporterPaymentProcess As String = "TRN-PAY-PRC"
    Public Const frmCattleFeedSalePurchaseUploader As String = "CF_S_PU_UPL"
    '===========================
    '--------------------------Reports-------------------------------------'
    Public Const FATSNFDiffReport As String = "FAT-SNF-DIFF"
    Public Const VLCProgressReport As String = "VLC-PRO-RPT"
    Public Const MPIncentiveEntryReport As String = "R-MP-INC-ENT"
    'Public Const FailedBDF As String = "Failed BDF"
    Public Const rptCollectionLevelChart As String = "M-COLL-CHART"
    Public Const rptCollectionCenterChart As String = "M-COLC-CHART"
    Public Const rptMCCShiftReportRouteWise As String = "M-SFT_RUTW"
    Public Const rptvillageslip As String = "M-VLG_SLP"
    Public Const rptMilkBillMCC As String = "M-BIL_MCC"
    Public Const rptMilkBillRouteWise As String = "M-Mlk_RUT"
    Public Const rptMilkPaymentRegister As String = "M-PAY_REG"
    Public Const rptMCCMilkBillSummary As String = "M-BILL_SUMM"
    Public Const rptLowProcurement As String = "M-LOW_PURMT"
    'Public Const rptSecondaryQuality As String = "M-SEC_QUAL"
    Public Const rptDailyProgressReport As String = "M-DAI_PROG"
    Public Const AreaMaster As String = "AREA-MST"
    Public Const rptMonthlyVLCProcurement As String = "M-MON_VLC"
    Public Const rptDailyDifferentReport As String = "M-DAI_DIFF"
    Public Const rptCDA As String = "CDA_RPT"
    Public Const rptMillPurchaseBill As String = "M_PURC_BILL"
    'Public Const RptGainSheetPeriod As String = "GAIN_SHEET"
    Public Const RptWeighment As String = "WEI_RPT"
    Public Const RptVLCVehicleWeigmentRegister As String = "VLC_WEG_REG"
    Public Const RptMilkReceiptImproperWeight As String = "MR-IMP-WET"
    'Public Const RptImproperMilkSample As String = "MS-IMP-MSP"
    Public Const rptMCCRouteTimeTable As String = "MCC-Rut-Time"
    Public Const RpttankerReport As String = "MCC-TAN-REP"

    'Public Const RptTankerVariation As String = "TANK_VARN"
    Public Const RptDailyGainDay As String = "DLY_GAN"
    Public Const RptPendingMilkSRN As String = "PEN_M_SRN"
    Public Const RptTankerSummaryReport As String = "TANK_SUM_RPT"
    Public Const rptShiftCodeWise As String = "RPT-Cod-Ws"
    Public Const rptShifReportZeroAmtSample As String = "Rpt-Z-A-S"
    Public Const FrmPendingProvisionReport As String = "PEN_PROV"
    Public Const rptMilkStockLedgerSummary As String = "MLK-STK-LED"
    Public Const RptMilkWeigmentRegister As String = "MLK-WGT_RGT"

    'Public Const rptPTReport As String = "MLK-PT-Rpt"
    Public Const MCCMilkRegister As String = "MCC_MLK_REG"
    Public Const MilkProcurementVisualReport As String = "PRO-VIS-RPT"
    'Public Const rptVSPItemIssue As String = "MCC_VSP_ITM"
    Public Const RptMilkRouteVehicleReport As String = "M_ROU_VEH"
    Public Const RptMccSaleRegister As String = "M_SAL_Reg"
    Public Const RptBulkMilkRegister As String = "M_BUL_Reg"
    Public Const RptTotalMilkProcurement As String = "TO_MILK_Proc"
    Public Const RptDispatchOfMilkTransfer As String = "DIS_MLK_TRN"
    'Public Const RptVLCTragetMasterReport As String = "VLC_TRG_RPT"
    'Public Const rptMCCVLCTragetMonthWiseReport As String = "MCC_TRG_RPT"
    Public Const RptBulkMilkMultiplePurchaseInvoice As String = "MUL_BUL_INV"
    Public Const MccSummaryReport As String = "Mcc_Sum_RPT"
    Public Const RptVillageDiffReport As String = "VILL_DIFF"
    Public Const RptVillageDiffReportParas As String = "VILL_DIFFPA"
    Public Const RptMPIDReport As String = "MP_ID_RPT"
    Public Const RptCollectionAnalysis As String = "COLL_ANALY"
    'Public Const frmMPIssueIDCard As String = "MP_ISS_ID"
    'Public Const RptMPIssueIDCard As String = "MP_Iss_Rpt"
    Public Const MCCProvisonReport As String = "MCC_PROR"
    'Public Const rptTankerDispatchWidthDeduction As String = "TD_DDED"
    Public Const RptDailyStanderdMilkQtyMCCWise As String = "DSM-QMW-RPT"
    'Public Const RptDailyLandedCost As String = "DLCost"
    Public Const RptWeighmentRegister As String = "WGH_REG_RPT"
    Public Const MonthlyProgressReport As String = "Mth-Prog"
    Public Const rptSuspanceDCSReport As String = "SUSP-DCS"
    Public Const rptTankerGainLossReport As String = "TAN-GAIN-L"

    Public Const RptSecondaryTransporterReport As String = "SEC-TR-RPT"
    Public Const rptMCCBillGenrationStatus As String = "MCC-BIL_GEN"
    Public Const rptMCCDataEntrySummaryReport As String = "MCC_DE_SUM"
    Public Const rptTankerAllocationReport As String = "MCC_TAN_ALL"
    Public Const rptMilkCollectionShiftWiseReport As String = "M_COL_SHWISE"
    Public Const rptMccDeductionReport As String = "MCC_DED_RPT"
    'Public Const rptShedIOAbstract As String = "SHD_IO_ABS"
    'Public Const rptWholeMilkAccount As String = "WHL-MLK-ACT"
    'Public Const rptPTPBill As String = "PTP-BIL"
    ' Public Const rptVLCTransportExpense As String = "VLC_TR_EXP"
    Public Const RptTankerDispatchvsAckn As String = "TAN_DIS_ACK"
    Public Const rptPaymentProcessReport As String = "RPT_PAY_PRO"
    Public Const rptMultipleDeductionReport As String = "RPT_MUL_DED"
    Public Const rptLiabilityReport As String = "MCC_LIA_RPT"
    Public Const rptVSPMilkNotsold As String = "Mlk_Not_Sold"
    Public Const rptMilkBillProcurementSummary As String = "MCC_MB_PS"
    Public Const frmVendorBankAdvice As String = "MCC_VB_ADV"
    Public Const rptPaymentProcessRouteReport As String = "ROUT_PR_RPT"
    Public Const BmcLabReport As String = "BMC_LB_RPT"
    Public Const CanSummaryReport As String = "CN_SM_RPT"
    Public Const YearlyDBTSummaryReport As String = "YE_DBT_SU"
    Public Const FrmTotalDeductionReport As String = "TO_DED_RPT"
    Public Const rptBMCDailyQtyReport As String = "BMC_D_Q_RPT"
    Public Const rptDBTMonthWiseFarmerDetail As String = "DBT-MO-W-FAR"

    Public Const FrmDayWiseMilkCollectionVsDispatch As String = "DY_MKCO_DS"
    Public Const rptMilkAnalysis As String = "MCC_MLK_ANA"
    Public Const rptSocietyLedgerReport As String = "MCC_SOC_LED"
    Public Const rptMilkCostReport As String = "MCC_MLK_COST"
    Public Const FrmDcscollectionHistory As String = "DCS_COL_HIS"
    Public Const FrmHeadLoadHistory As String = "HED_LOD_HIS"
    Public Const FrmOneDayStop As String = "One-Day-S"
    Public Const FrmLocalMilkSaleRegister As String = "Lo_M_Sale"

    Public Const rptPaymentProcessReportBMCSocietyWise As String = "PR_BSW_RPT"
    Public Const rptDCSFinancial As String = "DCS-FIN-RPT"
    '-----------------------Sale Purchase Security(Master)--------------
    Public Const FrmBankPermission As String = "BNK_PER"
    Public Const frmCustomerPermission As String = "CUST_PER"
    Public Const frmVendorPermission As String = "VENDOR_PER"
    Public Const frmCustomerPermissionReport As String = "CUST_PER_RPT"
    ''--------------------CSA SALE----------------------------------
    Public Const frmCSABooking As String = "CSA_BOOKING"
    Public Const frmCSARequest As String = "CSA_REQUEST"
    Public Const frmCSATransfer As String = "CSA_Transfer"
    Public Const rptsaleRegisterReport As String = "SLE_Register"
    Public Const rptPurchaseRegisterReport As String = "PUR_Register"
    Public Const rptSalesHierarchyReport As String = "SaleHier-rpt"
    Public Const FrmCSATransferGateOut As String = "CSA_TRSF_GO"
    Public Const RptPartyWiseSale As String = "PartyWiseS"
    Public Const RptGPDetail As String = "RPT_GP_DET"

    '======Sanjeet===================
    Public Const FrmDailySaleReport As String = "DAILY_SLERPT"
    Public Const FrmMonthlySaleReport As String = "MTHLY_SLERPT"
    Public Const FrmCustomerGroupOutstanding As String = "CUST_OUTSRPT"
    Public Const FrmStockAgeingAnalysisReport As String = "STAG_ANLSRPT"
    Public Const FrmCategoryAnalysisReport As String = "CAT_ANARPT"
    Public Const RptFlavouredMilk As String = "FLV_MLK_RPT"
    Public Const RptMonthWiseSaleAnalysis As String = "MNTH_ANA_RPT"
    Public Const RptJobWorkDebitNoteReport As String = "JW_DBT_NOTE"
    Public Const FrmSAC As String = "SAC_MASTER"
    Public Const FrmSACWiseTax As String = "SAC_WIS_TAX"
    Public Const FrmItemWiseTax As String = "ITEM_WIS_TAX"

    Public Const frmOverheadCostMaster As String = "OVER_COST_MT"
    Public Const FrmOverheadCostGroup As String = "OVER_COST_GP"
    Public Const FrmItemCostMapping As String = "ITM_COST_MAP"
    Public Const frmWeightUomMaster As String = "WGT_UOM_MST"
    Public Const frmJobWorkConsumption As String = "JWO-CONSU"

    Public Const FrmRackBinMaster As String = "RackBin_M"
    'Public Const frmDeptHeadCustomerMapping As String = "DPTH_CM"

    '===========================Added by Preeti Gupta=================  
    '======================Exit Managment=============================
    Public Const frmResignationLetter As String = "HR_REG_LTR"
    Public Const frmTerminationLetter As String = "HR_TER_LTR"
    Public Const frmResignationAcceptanceOrRejection As String = "HR_REG_Acc"
    Public Const frmHREMInterviewQuestion As String = "HR_INT_QUS"
    Public Const frmHREMExitInterview As String = "HR_EXT_MGM"
    '================================End Of Exit Managment

    '--------------------MIS Reports-------------------------------------------------
    Public Const MISDebtorReport As String = "MIS-DBTR-RPT"
    Public Const MISCreditorReport As String = "MIS-CDTR-RPT"
    Public Const MISSaleRegister As String = "MIS-Sale-RPT"
    Public Const rptTCSLedger As String = "TCS-LEDGER"
    Public Const RMStockConsumption As String = "STK-CNSP"
    Public Const MSIProductionSaleReport As String = "MIS-PSRPT"
    Public Const rptDCSSaleRegister As String = "DCS-SAL-REG"
    Public Const saleconsignee As String = "SAL_CON-DET"
    Public Const RoutePartyWisePriceList As String = "RT-PW-PL"
    Public Const MISSaleRegisterWithCSASalePatti As String = "MSCSAPATIREG"
    Public Const MISSaleRegisterWithCSASalePattiProductLocationWise As String = "MISPrdLocWis"
    Public Const MISSaleRegisterWithCSASalePattiProductPackWise As String = "MISPrdPckWis"
    Public Const MISStockReco As String = "MIS-SRec-RPT"
    Public Const MISStockLedgerReport As String = "MIS-STLE-RPT"
    Public Const FrmItemReloadReport As String = "ITM-RLD-RPT"
    Public Const rptSKUWiseSale As String = "SKU-SAL-RPT"
    Public Const rptPromptMsgPendindDoc As String = "PROMPTPENDOC"
    Public Const MISMassBalanceReport As String = "MAS-BAL-RPT"
    Public Const SundryDebtorsReport As String = "SUN-DBTR-RPT"
    Public Const AccountStatementReport As String = "ACC-SMT-RPT"
    '==========================Milk Job Work===================
    Public Const MilkJobWork As String = "MJW"
    Public Const SubModuleMJWTransaction As String = "SMMJWTRANS"
    Public Const SubModuleMJWReport As String = "SMMJWREPT"
    Public Const FrmMilkJobWork As String = "MI-J-W"
    Public Const FrmMilkGateEntry As String = "MI-GAT-ENT"
    Public Const FrmMilkWeighment As String = "MI-WEIGHM"
    Public Const FrmJobMilkQualityCheck As String = "MI-QUAL_CH"
    Public Const FrmMilkUnloading As String = "MI_UNLOAD"
    'Public Const FrmMilkCleaning As String = "MI_CLEANI"
    Public Const FrmJobMilkSRN As String = "JOB_MI_SRN"
    Public Const RptJobWorkStatus As String = "JOB_WRK_ST"
    '==SANJEET(07/09/2017)============
    Public Const RptJobWorkRegister As String = "RPT_JW_REG"
    Public Const RptJobWorOutwardPurchasekRegister As String = "RPT_JO_PREG"
    Public Const frmJobworkTransfer As String = "RPT_JW_TRSF"
    Public Const frmJobworkChargesReport As String = "RPT_JW_CHRG"
    Public Const frmJobworkSRNReceiptReport As String = "RPT_JW_SRNR"
    Public Const frmJobworkItemInStatusReport As String = "RPT_JW_ISR"
    '--------------------MIS Reports-------------------------------------ENDS HERE---

    '====================TDS MODULE PAYROLL===========Added By Preeti Gupta=======================
    Public Const ModuleTDSPayroll As String = "MTDSPay"
    Public Const SubModuleTDSPayrollSetup As String = "SMTDSPayS"
    Public Const SubModuleTDSPayrollTrans As String = "SMTDSPayT"
    Public Const SubModuleTDSPayrollReport As String = "SMTDSPayR"


    Public Const frmEmployeeTDSRpt As String = "EMPTDS_Rpt"
    Public Const frmIncomeTaxSlab As String = "TDS_ITS"
    Public Const frmITSection As String = "TDS_IT_Sec"
    Public Const frmInvestmentType As String = "TDS_INV_TYP"
    Public Const frmHouseRentDeclaration As String = "TDS_HUS_DEC"
    Public Const frmInvestmentDeclaration As String = "TDS_Inv_DEC"
    Public Const frmHRAExemptionRule As String = "TDS_ExE_RUL"
    '===== Added By Kunal : Silage Production Masters =======================
    Public Const ModuleSilageProduction As String = "MSILPRD"
    Public Const SubModuleSilagePrdSetup As String = "SMSILPRDSTP"
    Public Const SubModuleSilagePrdTrans As String = "SMSILPRDTRN"
    Public Const SubModuleSilagePrdRpt As String = "SMSILPRDRPT"

    Public Const frmSilageCriteriaMaster As String = "SlgCtraMst"
    Public Const frmSilageAreaMaster As String = "SlgAraMst"
    Public Const frmSilageProductionApplication As String = "SlgAppMst"
    Public Const frmSilageEnterPrenur As String = "SlgEntrMst"
    Public Const frmSilageFormerProduction As String = "SlgFrmrMst"

    Public Const frmSilageTankerTransporterMaster As String = "SlgTTM-MST"
    Public Const frmSilageTankerMaster As String = "SlgTNK-MST"
    Public Const frmSilageParameterMaster As String = "SlgPM-MST"
    Public Const frmSialgeParameterRangeMaster As String = "SlgPRM-MST"
    Public Const frmSialgeParameterRangeMasterForQC As String = "SlgP-RNG-QC"
    Public Const frmSialgeParameterValueMaster As String = "SlgPRM-VAL-M"
    Public Const frmSilagePriceChartBulkProc As String = "Slg-CHRT-BLK"
    Public Const frmSilageVendorPriceChartMapping As String = "SlgV-PC-MPNG"
    Public Const FrmSialgeSupplierMaster As String = "SlgSUPPLR-M"
    Public Const frmSilageDivertedContractor As String = "SlgDIV-CON-M"
    Public Const frmSilageTypeMast As String = "Slg-TYPE-M"
    Public Const frmSilageGradeMaster As String = "Slg-GRADE-M"

    '"""""""""""""""""""""""Transaction""""""""""""""""""""""""
    Public Const frmSialgeIntimation As String = "SLG-Intimat"
    Public Const frmSilageGateEntry As String = "Slg-GT-ENTR"
    Public Const frmSilageWeighment As String = "Slg-WEIGHT"
    Public Const frmSilageQualityCheck As String = "Slg-QC"
    Public Const frmSilageUnloading As String = "Slg-UNLD"
    Public Const frmSialgeCleaning As String = "Slg-CLEAN"
    Public Const frmSilageGateOut As String = "Slg-GT-OUT"
    Public Const frmSilageBulkSRN As String = "Slg-SRN"
    Public Const frmSilageBulkSRNReturn As String = "Slg-SRN-Ret"
    '======KUNAL  ===============================================
    Public Const frmLoadedWtEntry As String = "LDD-WT"
    Public Const frmEmptyWtEntry As String = "EMT-WT"
    '=========================================Preeti===========
    Public Const frmLocationItemMapping As String = "Loc_ITM_MAP"
    Public Const frmDistributorRouteTagging As String = "DIS-R-T"
    Public Const frmDailyDemand As String = "DAI-DEM"
    Public Const frmDBTUnionPayment As String = "DBT-UPAY"
    Public Const frmUnionLastGstInvoice As String = "DBT-GST_INV"
    Public Const rptCattleFeedProduction As String = "CF-PROD"
    Public Const frmUpdateAppLocation As String = "Upd_App_L"

    '=========Sanjeet(21/112016)=====================
    ' Public Const FrmTruckSheetRouteWiseRpt As String = "TSHT_RTWise"
    '=========Sanjeet(11/01/2018)=====================
    Public Const FrmSaleVsReceipReport As String = "SALE_VS_RCPT"
    Public Const rptCustomerIncentiveEntry As String = "CUS_INC_ENT"
    Public Const FrmCostCentreConsumptionRpt As String = "CST_CNT_CUN"
    Public Const FrmTenderTrackingReport As String = "PO_TTRRPT"
    Public Const FrmERPStatusTrackingReport As String = "PO_ERPSTRPT"
    Public Const rptrlPenaltyRegister As String = "RL-PEN-REG"
    Public Const rptPerformanceReport As String = "PERF-RPT"
    Public Const VehicleUnloadingReport As String = "VEH-UNL-RPT"
    Public Const rptTransactionWiseStock As String = "TRAN-STCK"
    Public Const Rptralwisematerialreceivedandpayment As String = "RAL-REC-RPT"
    Public Const RmSecurityDeduction As String = "RM-SRT-DED"
    Public Const WeightmentAutoAndManualReport As String = "WE-AU-MAN"
    Public Const rptRejectionReport As String = "REJ-RPT"
    Public Const CancelPurchaseInvoiceReport As String = "CN-P-I-RPT"
    Public Const rptRMUnloading As String = "RM-UNL-RPT"
    Public Const RptPOAgainstDocument As String = "PO_AGT"
    Public Const frmBillChecklist As String = "PO_BCL"
    Public Const frmCancelledTransactions_Purchase As String = "PO_CanRPT"
    Public Const frmDairyGatePass As String = "DS_GATE_PASS"
    Public Const frmUnitMaster As String = "CST_UNT_MSTR"
    Public Const frmCostCenterTypeMaster As String = "CST_TYP_MSTR"
    Public Const frmMaterialSalePriceChart As String = "Mat_Prc-Chrt"
    Public Const RptPendingDocumentList As String = "PND_DOC_LIST"
    Public Const RptApprovalReport As String = "Apprl_Rpt"
    Public Const rptUserScreenRightsReport As String = "USR_SEN_RGH"
    Public Const rptSMSDetailsReport As String = "SMS_DET_RPT"
    Public Const rptFarmerAbstractReport As String = "PP_ABST_RPT"
    Public Const rptScreenSettingReport As String = "RPT_SCN_RPT"
    Public Const rptProvisionChart As String = "RPT_PROV_CT"
    Public Const rptSalesVehicleReport As String = "RPT_SAL_VEH"


    Public Const frmMccScrapGatePass As String = "MS_GATE_PASS"
    Public Const frmMccMatSaleUploader As String = "MCC_MAT_UPL"
    Public Const frmProcurementDeduction As String = "MCC_PRO_DED"
    Public Const frmMultipleProcDeduction As String = "MCC_MUL_DED"
    '22/11/2016
    'Public Const FrmMccWeightDifferenceRpt As String = "Wgt_Dif_Rpt"


    '====================
    Public Const MilkProcurementUploader As String = "MLK-UPL-PRO"
    Public Const MilkShiftUploader As String = "MLK-PRO-UPL"
    Public Const MilkProcurementCorrection As String = "MLK-PRO-COR"
    Public Const MilkRetesting As String = "MLK-RE-TST"
    Public Const PrimaryTransportProvisionCorrection As String = "PTP-PRO-COR"
    Public Const frmOutputEntry As String = "OUTPUT-ENT"
    '=====(09/01/2017)==================
    'Public Const RptAClassMilkRate As String = "ACls-MlkRt"
    '=====(11/01/2017)==================
    Public Const RptPendingPO As String = "Rpt_Pend_PO"

    '=====(12/01/2017)==================
    Public Const RptIssueReturnHirerachyWise As String = "RptIssRetHw"

    '=====(17/01/2017)==================
    Public Const RptMccSaleAdjustment As String = "RptMccSleAdj"

    '=================Ravi====================
    Public Const FrmPOSGRoupMaster As String = "POS-GRP-M"
    Public Const frmPOSBookingDairyMultipleDistributor As String = "POS-BOOK-D"

    Public Const rptHierarchyWiseReport As String = "HIER_WIS_RPT"
    ''=========Parteek (30-01-2017)
    Public Const RptMccBulkmilkRegister As String = "Rpt-MccBulk"
    Public Const frmPromptMsgRelatedtopending As String = "PROMT_MSG"
    'Public Const frmjobWorkDebitNote As String = "JW-D-Note"

    Public Const frmChildRouteFreight As String = "Route-Frg"
    'Public Const frmConfigureSynchronization As String = "Conf-Sync"
    Public Const rptSaleReco As String = "rptSaleReco"
    Public Const rptPurReco As String = "rptPurReco"
    Public Const rptVendorReco As String = "rptVendReco"
    Public Const rptCustomerReco As String = "rptCustReco"
    Public Const RptInvReco As String = "RptInvReco"
    Public Const RptVSPCustomerReco As String = "VSP_Cus_Reco"
    Public Const rptSRNReco As String = "rptSRNReco"
    Public Const rptTankerDispatchGainLossReco As String = "rptTDGLReco"
    Public Const frmSettingDetails As String = "Setting-Det"
    Public Const frmAssetBookMaster As String = "Asset-Book"
    Public Const frmPendingBooking As String = "Pend-BOK"
    Public Const frmCustomerZone As String = "Cust-ZNE"
    Public Const FrmQuickDemandReport As String = "Quk-DEM-RPT"
    Public Const frmCrateCanreceiptReport As String = "CraCan-Rpt"
    Public Const rptCrateRegister As String = "rptCRAREG"
    'Public Const ApproveFailedSample As String = "APP-FAIL-SAM"

    ' Electrical
    Public Const frmSlotMaster As String = "SLOT_MST"
    Public Const frmDGMaster As String = "DG_MST"
    Public Const frmDailyElectricalEntry As String = "DLY_ELC_ETY"
    Public Const rptDailyElectricalEntryReport As String = "DLY_ELC_RPT"
    Public Const AuditTrailReceivable As String = "ADIT_REC"
    Public Const CustomerLedgerVsAgeing As String = "CL_VS_AG"
    Public Const AuditTrailPayables As String = "ADIT_PAY"
    Public Const VendorPaymentDetails As String = "VPD_REP"
    Public Const VendorLedgerVsAgeing As String = "VL_VS_AG"
    Public Const AuditTrailDairySale As String = "ADIT_DRYSAL"
    Public Const CancelledTransactionReportDS As String = "DS_CanRPT"
    Public Const AuditTrailSystemAdmin As String = "ADIT_SYSADM"
    Public Const AuditTrailCommonServices As String = "ADIT_COMSRV"
    Public Const AuditTrailGeneralLedger As String = "ADIT_GENLEG"
    Public Const AuditTrailSaleAndDistribution As String = "ADIT_SALDBU"
    Public Const rptSalesReport As String = "SALE-RPT"

    Public Const rptDealerSalesReport As String = "DEA_SA_RE"
    Public Const rptPlantWiseAvailableStockQtyReport As String = "P_AV_STC_QT"

    Public Const rptSaleReportCustomerWise As String = "SALE-CW-RPT"
    Public Const AuditTrailMaterialManagement As String = "ADIT_MATMGT"
    Public Const AuditTrailPurchase As String = "ADIT_PURCH"
    Public Const AuditTrailTaxDeductedAtSource As String = "ADIT_TAXSC"
    Public Const AuditTrailFixedAssets As String = "ADIT_FIXAST"
    Public Const RptAssetDispatchRetailer As String = "ASR_RPT"
    Public Const AuditTrailHRAndPayroll As String = "ADIT_HRPYRL"
    Public Const AuditTrailStandardProduction As String = "ADIT_STDPRO"
    Public Const rptSPItemConsumptionReport As String = "MATCONS_RPT"
    Public Const ProductionReport As String = "PRODUC_RPT"
    'Public Const QualitySummaryReport As String = "QTY_SUMY_RPT"
    Public Const FrmProductionAndSaleReport As String = "PANDS_RPT"
    Public Const AuditTrailDairyProduction As String = "ADIT_DARYPD"
    Public Const AuditTrailMilkProcurementMCC As String = "ADIT_MPMCC"
    Public Const AuditTrailMilkProcurementBulk As String = "ADIT_MPBULK"
    Public Const AuditTrailHumanResource As String = "ADIT_HUMREC"
    Public Const AuditTrailBulkSale As String = "ADIT_BLKSAL"
    Public Const BulkSaleAcknowledgementUploaderReport As String = "BKSAL_UPRPT"
    Public Const AuditTrailCSASale As String = "ADIT_CSASAL"
    Public Const AuditTrailFreshSale As String = "ADIT_FRSALE"
    Public Const AuditTrailProductSale As String = "ADIT_PRDSAL"
    Public Const AuditTrailExportSale As String = "ADIT_EXPSAL"
    Public Const AuditTrailMerchantTrade As String = "ADIT_MCHSAL"
    Public Const AuditTrailMilkJobWork As String = "ADIT_MILJW"
    Public Const AuditTrailTDSPayroll As String = "ADIT_TDSPRL"
    Public Const AuditTrailFarmerPayment As String = "ADIT_FARPYM"
    Public Const AuditTrailJobWorkOutward As String = "ADIT_JWOUT"
    Public Const AuditTrailJobWorkInward As String = "ADIT_JWINW"
    Public Const AuditTrailElectrical As String = "ADIT_ELTICL"
    Public Const rptProductionStatusReport As String = "PRO_STS_RPT"
    Public Const rptProductionStandardization As String = "PRO_SDZ_RPT"
    Public Const ProductionEntryReport As String = "PRO_ENT_RPT"
    Public Const frmProductionUtilityCost As String = "PRO_UC_RPT"
    Public Const VehicleMasterForDairySale As String = "VM_DS"
    Public Const frmDistributorCommission As String = "Dis-COM-MST"
    Public Const FrmBoothCommissionMaster As String = "BTH-COM-MST"
    Public Const FrmItemCapacityLimt As String = "ITM-CAP-LMT"
    Public Const FrmBoothRouteMapping As String = "BTH-ROT-MAP"
    Public Const frmDayBook As String = "DAY-BOOK"
    Public Const VehicleMasterForProductSale As String = "VM_PS"
    Public Const frmMRPForProduction As String = "PP_MRP"
    Public Const frmPriceMasterPS As String = "PRC-MST-PS"
    Public Const rptDeleteHistoryReport As String = "DEL-HIST"
    Public Const RCDFDashboard As String = "RCDF-DSB"
    Public Const RCDFUnionDashboard As String = "RCDF-UN-DSB"
    Public Const Complainfeedback As String = "COMP-FEED"
    Public Const DashboardMilkUnion As String = "UN-DSB"
    Public Const DairySaleDashboard As String = "DAR-DSB"
    Public Const SalesMarketingDashboard As String = "SAL-MAR-DAS"
    Public Const MostUserScreen As String = "MO_US_SC"
    Public Const rptDBTDashboard As String = "DBT-DSB"
    'Public Const CardSale As String = "CRD_SALE"

    'Eng. And Plant Management
    Public Const frmSectionMasterEng As String = "SECE-M"
    Public Const frmConsumptionTypeMaster As String = "CTE-M"
    Public Const frmSectionConsumptionMapping As String = "SCME-M"
    Public Const frmParameterMasterEng As String = "PME-M"
    Public Const frmSectionParameterMapping As String = "SPM-M"
    Public Const frmLogSheetEng As String = "LSE-T"
    Public Const WorkRequisitionEng As String = "WRE-T"
    Public Const frmWorkEstimationEng As String = "WEE-T"
    Public Const WorkOrderEng As String = "WOE-T"
    Public Const frmWorkOrderStatusEng As String = "WSE-T"

    Public Const frmCustomerComplaintMaster As String = "CUST_CMP_MST"
    Public Const frmCustomerComplaint As String = "CUST_CMP_TRN"
    Public Const frmLeakageReplacementUploader As String = "LEA_REP_TRN"
    Public Const frmAdjProductionEntry As String = "ADJ_PRO_ENY"
    'Public Const frmAdjProductionEntryQC As String = "ADJ_PE_QC"
    'Public Const frmAdjProductionStoreEntry As String = "ADJ_PE_SE"
    Public Const rptProductReceivingReport As String = "PRO_REC_RPT"
    Public Const frmTranspoterDeduction As String = "TRAN_DED_ENT"
    Public Const frmNotepadFileMatching As String = "NOT-MIS-MAT"
    Public Const frmDemandBooking As String = "DEM_BOO_TRN"
    Public Const FrmProductDemandBooking As String = "PRD_DEM_BOO"
    Public Const FrmProductDispatch As String = "PRD_DIS_TRN"
    Public Const frmDemandAdjustment As String = "DEM-ADJ-TRN"
    Public Const frmDemandApproval As String = "DEM-APR-TRN"
    Public Const frmDemand_Sheet As String = "DEM-DEM_SHE"
    Public Const frmQuickDemand As String = "QUICK-DEM"
    Public Const frmProductQuickDemand As String = "PRD-QUK-DEM"
    Public Const FrmBoothCommission As String = "BOOTH-COMM"
    Public Const frmLeakedSaleReturn As String = "LEK-SAL-RTN"
    Public Const frmDCSDEmandBooking As String = "DCS-DEM-TRN"
    Public Const frmBoothDispatch As String = "BOO-TH-DIS"
    Public Const frmDemandUploader As String = "DEM_UPL_TRN"
    Public Const frmDailyDemandReport As String = "DAI-DE-RE"
    Public Const rptHSNWiseSaleReport As String = "HSN_SAL_RPT"
    Public Const rptCustItemWiseSaleReport As String = "CUS_ITM_RPT"
    Public Const rptItemAndShiftWiseSaleSummaryReport As String = "ITM_SFT_RPT"
    Public Const rptCostCenterReport As String = "RPT_COST_CN"
    Public Const FrmperdayDetailRpt As String = "PER_DAY_RPT"
    Public Const rptBookingVsDispatchReport As String = "BOOK-VS-DI"
    Public Const rptMilkSupplySaleReport As String = "MLK_SUP_RPT"
    Public Const ModuleXpertAPI As String = "API"
    Public Const SubModuleXpertAPISetup As String = "MXAPIMaster"
    Public Const SubModuleXpertAPITrans As String = "MAPITrans"
    Public Const SubModuleXpertAPIReport As String = "MAPIReport"


    Public Const XpertAPIWeighment As String = "API-CTF-WGT"
    Public Const XpertAPILoadinSlip As String = "API-CTF-LIS"
    Public Const XpertAPIMilkReceipt As String = "API-MLK-REC"
    Public Const XpertAPIMilkSample As String = "API-MLK-SAM"
    Public Const XpertAPIMilkReject As String = "API-MLK-REJ"
    Public Const XpertAPIMilkEmptySample As String = "API-MLK-ETS"
    Public Const XpertAPIPlantWeighment As String = "API-PLT-WGT"
    Public Const XpertAPIQualityCheck As String = "API-PLT-QC"
    Public Const XpertAPIParlourShiftOpen As String = "API-PLR-SFO"
    Public Const XpertAPIParlourPOS As String = "API-PLR-POS"
    Public Const XpertAPIParlourShiftClose As String = "API-PLR-SFC"

    Public Const XpertAPIMilkMilkRegister As String = "API-MLK-REG"
    Public Const XpertAPIDairyGatePass As String = "API-GT-PASS"
    Public Const XpertAPITruckSheet As String = "API-TS-RPT"
    Public Const XpertAPIBMCDCSSamplingReceiving As String = "BMC-DCS-SR"
    Public Const XpertAPIBMCDCSSampleReceivedRegister As String = "API-DCS-SMR"
    Public Const XpertAPITankerCollectionReport As String = "API-TNK-CLN"
    Public Const XpertAPIRCDFChemicalQCRegister As String = "API-RCF-QCR"
    Public Const XpertAPIRCDFUnionMilkQCRegister As String = "API-RUM-QCR"
    Public Const XpertAPIGateEntry As String = "API-GT-ENT"
    Public Const XpertAPIBMCMilkQC As String = "API-BM-MK-QC"
    Public Const XpertAPIDEMANDBOOKING As String = "API-DEM-BOO"
    Public Const XpertAPISaleDispatchDairy As String = "API-DIS-DS"

    Public Const XpertAPIRCDFChemicalQC As String = "API-RCDF-QC"
    Public Const XpertAPIRCDFUnionMilkQC As String = "API-RUMK-QC"
    ''=================== MIS start here============

    Public Const ModuleMIS As String = "MMIS"
    Public Const SubModuleMISSetUp As String = "SMMIS"
    Public Const MISitemGroups As String = "MISITEM-GRP"
    Public Const MISitemMaster As String = "MISITEM-MST"
    Public Const frmProductionTransactionType As String = "PRO-TRAN-TYP"

    ''=================Transaction==========
    Public Const SubModuleMISTransaction As String = "SMMIST"
    Public Const frmDailyMilkProducts As String = "DLY-MLK-PROD"
    Public Const frmDailySMPProduction As String = "DLY-SMP-PROD"

    ''==================MILK UNION===================
    Public Const ModuleMilkUnion As String = "MLK_UNION"
    Public Const SubModuleMilkUnion As String = "RPT_MLK_UNN"


    ''==================Seed Module start here ===============
    Public Const ModuleSheed As String = "MSHEED"
    Public Const SubModuleSheedSetUp As String = "SMSHEED"
    Public Const frmSeedGrowerMaster As String = "SEED-GRWR-M"


    ''=========== seed transactions=====
    Public Const SubModuleSheedTransaction As String = "SMSHEED-T"
    Public Const frmSeedSelectionEntry As String = "SED-SEL-ENT"


    ''============seed MODULE  Report=======================
    Public Const SubModuleSheedReport As String = "SMSHEED-R"


    Private Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strLevel1 As String, ByVal strLevel2 As String, ByVal CurrentUserCode As String, ByVal CurrentCompanyCode As String) As Boolean
        Dim qry As String = "select PROGRAM_NAME from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Program_Name", strProgramName)
        clsCommon.AddColumnsForChange(coll, "Modify_By", CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy"))
        clsCommon.AddColumnsForChange(coll, "Comp_Code", CurrentCompanyCode)
        clsCommon.AddColumnsForChange(coll, "Level_1", strLevel1)
        clsCommon.AddColumnsForChange(coll, "Level_2", strLevel2)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            clsCommon.AddColumnsForChange(coll, "Program_Code", strProgramCode)
            clsCommon.AddColumnsForChange(coll, "Created_By", CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy"))
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PROGRAM_MASTER", OMInsertOrUpdate.Insert, "")
        Else ''If Not clsCommon.CompairString(strProgramName, clsCommon.myCstr(dt.Rows(0)("PROGRAM_NAME"))) = CompairStringResult.Equal Then
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PROGRAM_MASTER", OMInsertOrUpdate.Update, "Program_Code='" + strProgramCode + "'")
        End If
        Return True
    End Function

    ''  This function created by Pradeep Sharma on 14/06/13 11.34 AM
    ''  to insert default values in table TSPL_FIXED_PARAMETER
    ''  ONE TIME ONLY
    Public Shared Sub InsertDefaultValueIn_TSPL_FIXED_PARAMETER()
        Dim coll As New Hashtable()
        Dim qry As String = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'CboRound' "
        Dim check As Double = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "CboRound")
            clsCommon.AddColumnsForChange(coll, "Description", "Lower")
            clsCommon.AddColumnsForChange(coll, "Code", "L")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "CboRound")
            clsCommon.AddColumnsForChange(coll, "Description", "Upper")
            clsCommon.AddColumnsForChange(coll, "Code", "U")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "CboRound")
            clsCommon.AddColumnsForChange(coll, "Description", "Round")
            clsCommon.AddColumnsForChange(coll, "Code", "R")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Monthly")
            clsCommon.AddColumnsForChange(coll, "Code", "M")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'cboPeriodicity' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Monthly")
            clsCommon.AddColumnsForChange(coll, "Code", "M")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Quarterly")
            clsCommon.AddColumnsForChange(coll, "Code", "Q")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Half Yearly")
            clsCommon.AddColumnsForChange(coll, "Code", "HY")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Annually")
            clsCommon.AddColumnsForChange(coll, "Code", "A")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "cboPeriodicity")
            clsCommon.AddColumnsForChange(coll, "Description", "Other")
            clsCommon.AddColumnsForChange(coll, "Code", "OTH")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

        End If
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'PayHeadSubHead' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Employee PF")
            clsCommon.AddColumnsForChange(coll, "Code", "EPF")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Company PF")
            clsCommon.AddColumnsForChange(coll, "Code", "COPF")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Company EPS")
            clsCommon.AddColumnsForChange(coll, "Code", "COEPS")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Company ESI")
            clsCommon.AddColumnsForChange(coll, "Code", "COESI")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Employee ESI")
            clsCommon.AddColumnsForChange(coll, "Code", "EMPESI")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "TDS")
            clsCommon.AddColumnsForChange(coll, "Code", "TDS")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Reimbursement")
            clsCommon.AddColumnsForChange(coll, "Code", "RMBT")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Allowance")
            clsCommon.AddColumnsForChange(coll, "Code", "ALLOW")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Deduction")
            clsCommon.AddColumnsForChange(coll, "Code", "DEDUCT")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Loan")
            clsCommon.AddColumnsForChange(coll, "Code", "LOAN")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Bonus")
            clsCommon.AddColumnsForChange(coll, "Code", "BONUS")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "OT")
            clsCommon.AddColumnsForChange(coll, "Code", "OT")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Other")
            clsCommon.AddColumnsForChange(coll, "Code", "OTHER")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Basic")
            clsCommon.AddColumnsForChange(coll, "Code", "BASIC")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "HRA")
            clsCommon.AddColumnsForChange(coll, "Code", "HRA")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "TA")
            clsCommon.AddColumnsForChange(coll, "Code", "TA")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "DA")
            clsCommon.AddColumnsForChange(coll, "Code", "DA")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PayHeadSubHead")
            clsCommon.AddColumnsForChange(coll, "Description", "Professional Tax")
            clsCommon.AddColumnsForChange(coll, "Code", "PT")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")

        End If

        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'PromptForTally' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "PromptForTally")
            clsCommon.AddColumnsForChange(coll, "Description", "0")
            clsCommon.AddColumnsForChange(coll, "Code", "PromptForTally")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If
        '' create po without requisition
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'CreatePOWithRequisition' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "CreatePOWithRequisition")
            clsCommon.AddColumnsForChange(coll, "Description", "0")
            clsCommon.AddColumnsForChange(coll, "Code", "POWITHREQ")
            clsCommon.AddColumnsForChange(coll, "Specification", "0-OFF;1-On: if off then user can create PO without selecting Req No else REq No is mandatory to create PO")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

        '' Popup enabler for Reorder level of items.
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'EnablePopupItemReorderLevel' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "EnablePopupItemReorderLevel")
            clsCommon.AddColumnsForChange(coll, "Description", "0")
            clsCommon.AddColumnsForChange(coll, "Code", "POPUPITEMREORDERLEVEL")
            clsCommon.AddColumnsForChange(coll, "Specification", "0-OFF;1-On: if off then Popup for item reorder level will not displayed  else displayed during login of erp")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

        '' Ask for reason on delete of records.
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'DisplayReasonOnDelete' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "DisplayReasonOnDelete")
            clsCommon.AddColumnsForChange(coll, "Description", "0")
            clsCommon.AddColumnsForChange(coll, "Code", "DisplayReasonOnDelete")
            clsCommon.AddColumnsForChange(coll, "Specification", "0-OFF;1-On: if on then Reason will be asked during delete of records else will not asked.")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

        '' "DISCRETE" AND "PROCESS" Manufacturing.
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'ManufacturingType' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "ManufacturingType")
            clsCommon.AddColumnsForChange(coll, "Description", "D")
            clsCommon.AddColumnsForChange(coll, "Code", "ManufacturingType")
            clsCommon.AddColumnsForChange(coll, "Specification", "D-DISCRETE Manufacturing; P-PROCESS Manufacturing")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

        '' ENABLE MILK PROCUREMENT MODULE.
        qry = "SELECT Count(*) FROM TSPL_FIXED_PARAMETER where Type= 'MilkProc' "
        check = clsDBFuncationality.getSingleValue(qry)
        If check = 0 Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Type", "MilkProc")
            clsCommon.AddColumnsForChange(coll, "Description", "0")
            clsCommon.AddColumnsForChange(coll, "Code", "EnableMilkProc")
            clsCommon.AddColumnsForChange(coll, "Specification", "1-Enable Milk Procurement Module; 0-Disable Milk Procurement Module")
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If

    End Sub

    Public Shared Function GetUserLevels() As DataTable
        Dim dt As New DataTable
        Try
            dt.Columns.Add("Code", GetType(Integer))
            dt.Columns.Add("Description", GetType(String))

            dt.Rows.Add(0, "Select")
            dt.Rows.Add(1, "Budgetary")
            dt.Rows.Add(2, "Finance")
            dt.Rows.Add(3, "Level1")
            dt.Rows.Add(4, "Level2")
            dt.Rows.Add(5, "Level3")
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function GetName(ByVal strProgramCode As String) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when LEN(ISNULL(Re_Name,''))>0 then Re_Name else Program_Name end as Program_Name from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "'"))
    End Function
End Class

Public Class clsDashBoard
    Public Const BankCashBook As String = "BankCashBook"
    Public Const VehicleUtili As String = "VehicleUtili"
    Public Const ProcMilkPur As String = "ProcMilkPur"
    Public Const MilkRecAtFac As String = "MilkRecAtFac"
    Public Const MilkSale As String = "MilkSale"
    Public Const ProductSale As String = "ProductSale"
    Public Const Transportcos As String = "Transportcos"
    Public Const StoreReport As String = "StoreReport"

    Public Shared Function InsertDefaultValueDashBoard(ByVal strProgramCode As String, ByVal strProgramName As String) As Boolean
        Dim qry As String = "select NAME from TSPL_DASHBOARD_REPORT where Code='" + strProgramCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Name", strProgramName)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            clsCommon.AddColumnsForChange(coll, "Code", strProgramCode)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DASHBOARD_REPORT", OMInsertOrUpdate.Insert, "")
        Else
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DASHBOARD_REPORT", OMInsertOrUpdate.Update, "Code='" + strProgramCode + "'")
        End If
        Return True
    End Function
End Class
