'===========BM00000007858==============
Imports common
Imports System.Data.SqlClient
Public Class clsFixedParameterType
    Public Const SkadaIntegration As String = "Skada Integration"
    Public Const SkipDebitNoteforRejectInBMCTruckSheet As String = "Skip Debit Note for Reject In BMC Truck Sheet"
    Public Const GatepassForTaxableandNonTaxableItems As String = "Gatepass For Taxable and NonTaxable Items"
    Public Const PickDataFromMasterOrTransaction As String = "Pick Data From Master Or Transaction"
    Public Const ExportTruckSheet As String = "Export Truck Sheet"
    Public Const ThirtPartyFarmerCollectionIntegration As String = "Thirt Party Farmer Collection Integration"
    Public Const AndroidFarmerCollection As String = "Android Farmer Collection"
    Public Const ShowRateAndAmountOnFarmerCollection As String = "Show Rate And Amount On Farmer Collection"
    Public Const RefreshTimeForUnionWiseMilkCollectionInms As String = "Refresh Time For Union Wise MilkCollection In MS"
    Public Const SuspenceAdjusmtmentFATSNFTolerancePercentage = "Suspence Adjusmtment FAT SNF Tolerance %"
    Public Const TransactionModify As String = "Transaction Modify"
    Public Const TransactionValidity As String = "Transaction Validity"
    Public Const AutoGenerateFarmerPriceFromDCSPrice As String = "Auto Generate Farmer Price From DCS Price"
    Public Const AutoGenrateBatchInventoryOnPost As String = "Auto Genrate Batch Inventory On Post"

    Public Const PickcompanyBankCodeFromPaymentProcess As String = "pick company BankCode From Payment Process"
    Public Const isThereOnlyOneRowOfOwnDCS As String = "is There Only One Row Of Own DCS"
    Public Const MaxFATPerLimitforReport As String = "Max FAT Per Limit for Report"
    Public Const PartialSRNAmount As String = "Partial SRN Amount"
    Public Const MaxSNFPerLimitforReport As String = "Max SNF Per Limit for Report"
    Public Const PrintOnlyPostedDocument As String = "Print Only Posted Document"
    Public Const DefaultFilesPath As String = "Default File Path"
    Public Const B2CDocumentAmountRangeOtherState As String = "B2C Document Amount Range Other State"
    Public Const B2CDocumentAmountRange As String = "B2C Document Amount Range"
    Public Const ApplySendApprovalSetting As String = "Apply Send Approval Setting"
    Public Const ImportorExport As String = "Password for Import/Export"
    Public Const DCSWiseFilterEnableOnSavingCheck As String = "DCS Wise Filter Enable On Saving Check"
    Public Const NoOfDCSToLoadDeductionData As String = "No Of DCS To Load Deduction Data"
    Public Const PenaltyAfterDays As String = "Define Penalty Days to Apply"
    Public Const PenaltyCost As String = "Define Penalty Cost Value for Per Unit"
    Public Const CalculateProRateAuto As String = "Calculate Pro Rate Auto"
    Public Const ConvertVlcCodeUploaderToInt As String = "Convert VlcCodeUploader To Int"
    Public Const PickOnlyOWNBMCDCS As String = "Pick Only OWN BMC DCS"
    Public Const PrintPouchCrateQtyOnPrint As String = "Print PouchCrate Qty On Print"
    Public Const PickHeadLoadRateFromSecretaryMaster As String = "Pick HeadLoadRate From Secretary Master"
    Public Const Transactionupdate As String = "Transtion Update"
    Public Const CreateEInvoicePassword As String = "Create EInvoice Password"
    Public Const RCDFRateControl As String = "RCDF Rate Control"
    Public Const ImportExportDrive As String = "Import/Export-Drive"
    Public Const RCDFControl As String = "RCDF Control"
    Public Const EmployeePFRetirementAge As String = "Employee PF Retirement Age"
    Public Const EmployeeRetirementAge As String = "Employee Retirement Age"
    Public Const DBTMilkQtyCapping As String = "DBT Milk Qty Capping"
    Public Const DBTByDailyFarmerMilkCollection As String = "DBT By Daily Farmer Milk Collection"
    Public Const PDAccountPaymanager As String = "PD Account Paymanager"
    Public Const BankAdviseRequired As String = "Bank Advice Start Date"
    Public Const SeprateBankForDCSBankAdvice As String = "Seprate Bank For DCS Bank Advice"
    Public Const ApplyPartialBankAdvise As String = "Apply Partial Bank Advise"
    Public Const ApplyMergeForDCSMultipleDays As String = "Apply Merge For DCS Multiple Days"
    Public Const CalculateFATSNFLossCycleWise As String = "Calculate FAT SNF Loss Cycle Wise"
    Public Const PurchaseSlab As String = "PurchaseSlab"
    Public Const RefreshDBTReco As String = "Refresh DBT Reco"
    Public Const DistributorWiseBilling As String = "Distributor Wise Billing"
    Public Const ApplyBoothSecurity As String = "Apply Booth Security"
    Public Const AndroidSaleOrder As String = "Android Sale Order"
    Public Const RoundOffBankAdvice As String = "RoundOff Bank Advice"
    Public Const NewDCSScreen As String = "New DCS Screen"
    Public Const SameBankCodeForCurrentAndSaving As String = "Same Bank Code For Current And Saving"
    Public Const MinimumQtyForHeadLoad As String = "Minimum Qty For Head Load"
    Public Const ApplyHeadLoadDayAndCycleWiseForAJM As String = "Apply Head Load Day And Cycle Wise For AJM"
    Public Const StopSetting As String = "Stop Setting"
    Public Const PickBulkRoute As String = "Pick Bulk Route"
    Public Const ShowMultipleLegers As String = "Show Multiple Legers"
    Public Const LoadLedgerMixedMilk As String = "Load Ledger Mixed Milk"
    Public Const HeadLoadRODecimalPlace As String = "Head Load RO Decimal Place"
    Public Const HeadLoadROIncreaseAfter As String = "Head Load RO Increase After"
    Public Const ApplyUnpaidBank As String = "Apply Unpaid Bank"
    Public Const AllowZeroFATSNF As String = "Allow Zero FAT SNF"
    Public Const XpertAPI As String = "Xpert API"
    Public Const MaxRowsExcelDBTNEFTUploader As String = "Max Rows Excel DBT NEFT Uploader"
    Public Const ShowSampleNoOnBMC As String = "Show Sample No On BMC"
    Public Const DBTRevisePayment As String = "DBT Revise Payment"
    Public Const CycleWiseDBT As String = "Cycle Wise DBT"
    Public Const ShowTempratureOnBMC As String = "Show Temprature On BMC"
    Public Const PickMilkTransferInData As String = "Pick Milk Transfer In Data"
    Public Const FillRouteTankerNo As String = "Fill Route Tanker No"
    Public Const PickMilkPurchaseInvoiceQtyOrRecoQty As String = "Pick Milk Purchase Invoice Qty Or Reco Qty"
    Public Const AllowMPIncetiveQtyAboveBilledQty As String = "Allow MP Incentive Qty Above Billed Qty"
    Public Const RepeatBMCSampleNo As String = "Repeat BMC Sample No"
    Public Const HeaderFATSNFKGDecimalPlaces = "Header FAT SNF KG Decimal Places"
    Public Const SNFDecimalPlaces = "SNF Decimal Places"
    Public Const BMCDecimalPlaces = "BMC Decimal Places"
    Public Const BMCDCSApplyShift = "BMC DCS Apply Shift"
    Public Const AdjustFATSNFINOwnVSP = "Adjust FAT SNF IN Own VSP"
    Public Const ApplySameDayShift = "Apply Same Day Shift"
    Public Const HideShiftCollection As String = "Hide Shift Collection"
    Public Const MilkCollectionPickBulkRoute As String = "Milk Collection Pick Bulk Route"
    Public Const DailyQtyReport As String = "Daily Qty Report"
    Public Const ShowOwnBMCDCS As String = "Show Own BMC DCS"
    Public Const OwnBMCCreateDRCRNote As String = "Own BMC Create DR CR Note"
    Public Const TDSOpeningBalanceConsiderCreditNote As String = "TDS Opening Balance Consider Credit Note"
    Public Const OwnBMCApplicationFATRatio As String = "Own BMC Application FAT Ratio"
    Public Const OwnBMCApplicationSNFRatio As String = "Own BMC Application SNF Ratio"
    Public Const FATSNFNoDecimalMCC As String = "FAT SNF No Decimal MCC"
    Public Const FATSNFNoDecimalDCS As String = "FAT SNF No Decimal DCS"
    Public Const ShowAllMCC As String = "Show All MCC"
    Public Const ShowDCSDetMerge As String = "Show DCS Detail in Merge"
    Public Const ShowAllDCS As String = "Show All DCS"
    Public Const ApplyGaze As String = "Apply Gaze"
    Public Const MilkCollectionFATSNFType As String = "Milk Collection FAT SNF Type"
    Public Const MilkCollectionFATSNFTypeHeader As String = "Milk Collection FAT ShowDCSDetMerge Type Header"
    Public Const JanAadharNoMandatory As String = "Jan Aadhar No Mandatory"
    Public Const StopUpdateForWeigingMilkReceipt As String = "Stop Update For Weiging Milk Receipt"
    Public Const CreatePOFromMultipleLocation As String = "Create PO From Multiple Location Indent"
    Public Const ApplyPashuAaharAndMineralMixture As String = "Apply Pashu Aahar And Mineral Mixture"
    Public Const DCSRecoCondition As String = "DCS Reco Condition"
    Public Const BulkProcRunOneTypeGateEntry As String = "Bulk Proc Run One Type Gate Entry"
    Public Const ApplyZoneInDBT As String = "Apply Zone In DBT"
    Public Const BankIFSCCodeValidateByService As String = "Bank IFSC Code Validate By Service"
    Public Const AndroidAPP As String = "Android APP"
    Public Const AndroidAPPVersion As String = "Android APP Version"
    Public Const AndroidMPMaster As String = "Android MP Master"
    Public Const AndroidMPIncetiveEntry As String = "Android MP Incetive Entry"
    Public Const AndroidMilkCollectionBMCDCS As String = "Android Milk Collection BMC DCS"
    Public Const AndroidDemandBooking As String = "Android Demand Booking"
    Public Const AndroidAMCUDetails As String = "Android AMCU Details"

    Public Const MPIncentiveEntryApplyMonthly As String = "MP Incentive Entry Apply Monthly"
    Public Const DCSQtyDecimalPlaces As String = "DCS Qty Decimal Places"
    Public Const MPIncentiveEntryCycleWiseButNEFTMonthly As String = "MP Incentive Entry Cycle Wise But NEFT Monthly"
    Public Const MPIncentiveEntryMaxMilkLimit As String = "MP Incentive Entry Max Milk Limit"
    Public Const MPIncentiveEntryIncentiveRate As String = "MP Incentive Entry Incentive Rate"
    'Public Const TrendDiffValueForColor As String = "Trend Diff Value For Color"
    Public Const AllowPurReturnEvenIfPaymentDone As String = "Allow Pur Return Even If Complete Payment Done"
    Public Const TransporterCollection = "Transporter Collection"
    Public Const ProcurmentShiftUploaderNo = "Procurment Shift Uploader No"
    'Public Const FillBatchNoAsBatch = "Bhole Baba Payment Process Pro Print Start Date"
    Public Const BholeBabaPaymentProcessProPrintStartDate = "Bhole Baba Payment Process Pro Print Start Date"
    Public Const BankCodeForApplyDocumentPaymentOFAssetLost = "Bank Code For Apply Document PaymentOF Asset Lost"
    Public Const SeprateDistanceMorningEveningShift = "Seprate Distance Morning/Evening Shift"
    Public Const TankerDispatchProvisionLocationSegment = "Tanker Dispatch Provision Location Segment"
    Public Const IncludeInceAndDedInFATSNFRate = "Include Ince And Ded In FAT SNF Rate"
    Public Const CanSaleAvgFATSNFPer = "Can Sale Avg FAT SNF Per"
    Public Const FATPerShouldBeMultipleOf5 = "FAT Per Should Be Multiple Of 5"
    Public Const PickFATSNFKgFromInventory = "Pick FAT SNF Kg From Inventory For Milk Register"
    Public Const ALLOWBOOKINGSHITWISE = "ALLOW BOOKING SHIT WISE"
    Public Const FindReasonWhyInvoiceIssueOccursOnErode = "FindReasonWhyInvoiceIssueOccursOnErode"
    Public Const ApplyGovtRulesInTDS = "Apply Govt Rules In TDS"
    Public Const AllowtoChangeFatANdSNFPerforHighClassVendorinGE = "AllowtoChangeFatANdSNFPerforHighClassVendorinGE"
    Public Const BookingMobileAppChangeorderByBookingQty = "Booking Mobile App Change order By Booking Qty"
    Public Const BookingMobileAppSetNxtDateofBookingOrder = "Booking Mobile App Set Nxt Date of Booking Order"
    Public Const MCCBioSyncDate = "MCC Bio Sync Date"
    Public Const EnableTDSforServiceVendorSeparately = "Enable TDS for Service Vendor Separately"
    Public Const VSPBillDocumentToBeAddedInMilkCost = "VSP Bill Document To Be Added In Milk Cost"
    Public Const DCSAddDedRODecimalPlace = "DCS Add Ded RO Decimal Place"
    Public Const DCSAddDedROIncreaseAfter = "DCS Add Ded RO Increase After"
    Public Const DCSAddDedROHeaderLevel = "DCS Add Ded RO Header Level"
    Public Const ApplyTDSValidationFrom = "Apply TDS Validation From"
    Public Const ConsiderPreviousCurrentFYForTCSTaxVendOutstanding = "ConsiderPreviousCurrentFYForTCSTaxVendOutstanding"
    Public Const AmountToCheckVendorOutstandingForTCSTax = "AmountToCheckVendorOutstandingForTCSTax"
    Public Const AllowtoChangeTCSBaseAmountPurchase = "AllowtoChangeTCSBaseAmountPurchase"
    Public Const AllowFatPerInanynumberofMultipesonBulkQC = "AllowFatPerInanynumberofMultipesonBulkQC"
    Public Const CreateSeparateInvoiceforeachrowinCansale = "CreateSeparateInvoiceforeachrowinCansale"
    Public Const AllowCanInformationintoGridForTankerDispatch = "AllowCanInformationintoGridForTankerDispatch"
    Public Const TCSRateforVendorWithPanNo = "TCSRateforVendorWithPanNo"
    Public Const ShowAvailableQtyOnDairyBooking = "ShowAvailableQtyOnDairyBooking"
    Public Const AllowTransferVSPAmtToFarmerinVCGL = "AllowTransferVSPAmtToFarmerinVCGL"
    Public Const TCSRateforVendorWithoutPanNo = "TCSRateforVendorWithoutPanNo"
    Public Const IncetiveEntryApplyArrear As String = "Incetive Entry Apply Arrear"
    Public Const VSPDayWiseIncentiveAtSRN As String = "VSP Day Wise Incentive At SRN"
    Public Const IncentiveProcessPaymentStartDate As String = "Incentive Process Payment Start Date"
    Public Const CreateJVOFRGPNRGPAndItsRetrun As String = "Create JV Of RGP NRGP And Its Retrun"
    Public Const FixedAssetAcquisitionEntryHitInventoryMovement As String = "Acquisition Entry Hit Inventory Movement"
    Public Const CalculateDeductionByStdQtyinBulkMilkSRN = "Calculate Deduction By Std Qty in Bulk Milk SRN"
    Public Const CreateAPinvoiceofsalaryemployeewiseduringsalarygen = "CreateAPinvoiceofsalaryemployeewiseduringsalarygen"
    Public Const ApplyTaxInBulkMilkPurchaseInvoice = "Apply Tax In Bulk Milk Purchase Invoice"
    Public Const MilkSRNFATSNFDecimalPlaces = "Milk SRN FAT SNF Decimal Places"
    Public Const TankerDispatchIntermittentSingleGateIn = "Tanker Dispatch Intermittent Single Gate In"
    Public Const JobWorkOutwardComsumeItemAccordingToBOM = "Job Work Outward Comsume Item According to BOM"
    Public Const JWOutWardForRCDF As String = "Job Work Out Ward For RCDF"
    Public Const CreateARAdjAPDebitnoteforEmployeesinMCCMS = "CreateARAdjAPDebitnoteforEmployeesinMCCMS"
    Public Const ApplyCardSaleInvoiceOnlyWithCardSaleAdvance = "ApplyCardSaleInvoiceOnlyWithCardSaleAdvance"
    Public Const StartDateforDispatchSchedular = "StartDateforDispatchSchedular"
    Public Const StockRecoRateTunning = "Stock Reco Rate Tunning"
    Public Const BankRecoCheckFutureDocuments = "Bank Reco Check Future Documents"
    Public Const FATSNFRate = "Production Utility FAT/SNF Rate"
    Public Const ProductionRemoveFATSNFKgTollerance = "Production Remove FAT SNF Kg Tollerance"
    Public Const SubLocationForTaxableItemDairyDispatch = "SubLocationForTaxableItemDairyDispatch"
    Public Const SubLocationForNonTaxableItemDairyDispatch = "SubLocationForNonTaxableItemDairyDispatch"
    Public Const CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo = "CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo"
    Public Const EnableTCSRateValidityFrom01July2021 = "EnableTCSRateValidityFrom01July2021"
    Public Const ProductionCheckFATKg = "Production Check FAT Kg"
    Public Const ProductionCheckSNFKg = "Production Check SNF Kg"
    Public Const ProductionOnlyOneIssueAgainstBatch = "Production Only One Issue Against Batch"
    Public Const MilkShiftEndAutoAdjInItemCode = "Milk Shift End Auto Adjustment In Item Code"
    Public Const StockCheckOnPostForDairyDispatchMultiple = "StockCheckOnPostForDairyDispatchMultiple"
    Public Const BatchFileCounter = "BatchFileCounter"
    Public Const checkstockMRPwise = "checkstockMRPwise"
    Public Const TCSRateforCustomerWithPanNo = "TCSRateforCustomerWithPanNo"
    Public Const ScrapTCSRateforCustomerWithPanNo = "ScrapTCSRateforCustomerWithPanNo"
    Public Const TCSRateforCustomerWithoutPanNo = "TCSRateforCustomerWithoutPanNo"
    Public Const ScrapTCSRateforCustomerWithoutPanNo = "ScrapTCSRateforCustomerWithoutPanNo"
    Public Const TCSTaxApplicableOnbulkSale = "TCSTaxApplicableOnbulkSale"
    Public Const TCSTaxApplicableOnCanSale = "TCSTaxApplicableOnCanSale"
    Public Const AllowtoChangeTCSBaseAmount = "AllowtoChangeTCSBaseAmount"
    Public Const EInvoiceVendor = "EInvoiceVendor"
    Public Const TokenTimeReGenForEinvoice = "TokenTimeReGenForEinvoice"
    Public Const TokenTimeReGenForEWayBill = "TokenTimeReGenForEWayBill"
    Public Const ItemCostZeroOnStoreAdjForTypeFlushing = "ItemCostZeroOnStoreAdjForTypeFlushing"
    Public Const AutoCreateSaleInvoice = "Auto Create Sale Invoice"
    Public Const AmountToCheckCustomerOutstandingForTCSTax = "AmountToCheckCustomerOutstandingForTCSTax"
    Public Const DonotConsiderDirectARInvoiceinCustomerOutTCS = "DonotConsiderDirectARInvoiceinCustomerOutTCS"
    Public Const ConsiderPreviousCurrentFYForTCSTaxCustOutstanding = "ConsiderPreviousCurrentFYForTCSTaxCustOutstanding"
    Public Const AllowJEofDifferentLocationOnJournalEntry = "AllowJEofDifferentLocationOnJournalEntry"
    Public Const MilkProcurementBatchPosting = "Milk Procurement Batch Posting"
    Public Const popupcustomernamewhileupdating = "popupcustomernamewhileupdating"
    Public Const ApplyFEFO = "ApplyFEFO"
    Public Const EnterLocationForJWEStimationOutPackingMaterial = "EnterLocationForJWEStimationOutPackingMaterial"
    Public Const CreateJVofPackingMaterialofJWInwardinJWEstimate = "CreateJVofPackingMaterialofJWInwardinJWEstimate"
    Public Const AllowtoenterrateIntoJobWorkDispatch = "AllowtoenterrateIntoJobWorkDispatch"
    Public Const ConsiderUnpostedDocForBalance = "Consider Unposted Doc For Balance"
    Public Const TankerDispatchAvgFATSNFPer = "Tanker Dispatch Avg FAT SNF Per"
    Public Const DisableToPickMainLocationType = "Disable To Pick Main Location Type"
    Public Const MultipleMCCFinder = "Multiple MCC Finder In VSP Bill Generation"
    Public Const MilkProcurementSNF2DecimalPlaces = "Milk Procurement SNF 2 Decimal Places"
    Public Const SelectMilkRejectDefaulterManually = "Select Milk Reject Defaulter Manually"
    Public Const DayWiseCustomerIncentiveCalculation = "Day Wise Customer Incentive Calculation"
    Public Const CustomerIncetiveAutoSecuity = "Create Auto Security on Customer Incentive"
    Public Const CustomerIncetiveBankForSecuity = "Customer Incetive Bank For Secuity"
    Public Const CustomerIncetivePaymentModeForSecuity = "Customer Incetive Payment Mode For Secuity"
    Public Const RepeatDeductionAndVendor = "Repeat Deduction And Vendor"
    Public Const LastMilkReceiptQtyTollerance = "Last Milk Receipt Qty Tollerance"
    Public Const ShowMCCFinderInPaymentProcess = "Show MCC Finder In Payment Process"
    Public Const UseDescInsteadOFCodeOnMCCMAterialSale = "UseDescInsteadOFCodeOnMCCMAterialSale"
    Public Const AllowShelfLifeMandatoryOnFG = "AllowShelfLifeMandatoryOnFG"
    Public Const RefundknockoffwithCreditNote = "Refund knockoff with Credit Note"
    Public Const CreateInvoiceAutomaticallyOnbulkDispatch = "CreateInvoiceAutomaticallyOnbulkDispatch"
    Public Const CrateReceivingWithMultipleRoute = "CrateReceivingWithMultipleRoute"
    Public Const ShowMulMRPOfSameItemOnDairyBookingCustomer = "ShowMulMRPOfSameItemOnDairyBookingCustomer"
    Public Const AllowZeroQtyOnDairyBooking = "AllowZeroQtyOnDairyBooking"
    Public Const UseCutOffTimeonRouteForERP = "UseCutOffTimeonRouteForERP"
    Public Const AllowZeroQtyOnDairyBookingUploader = "AllowZeroQtyOnDairyBookingUploader"
    Public Const AllowtoPostNoOFDocofDOatatime = "AllowtoPostNoOFDocofDOatatime"
    Public Const DonotIncludeSecurityInCustomerOutstanding = "DonotIncludeSecurityInCustomerOutstanding"
    Public Const IsManualTCS = "Manual TCS"
    Public Const DefaultLocationForCardSaleIntegration = "DefaultLocationForCardSaleIntegration"
    Public Const ApplyNoGSTCreditIndependentlyOnVendorServiceCharge = "ApplyNoGSTCreditIndependentlyOnVendorServiceCharge"
    Public Const CheckNoOfDaysforCardSaleBooking = "CheckNoOfDaysforCardSaleBooking"
    Public Const MaxNoOfBookingAllowedThroughBookingApp = "MaxNoOfBookingAllowedThroughBookingApp"
    Public Const ApplyBothtsrateAndFatRateinBulkProcurement = "ApplyBothtsrateAndFatRateinBulkProcurement"
    Public Const PickTCAForStockTransferAndTankerDispatch = "PickTCAForStockTransferAndTankerDispatch"
    Public Const ShowBookingTypeDropDownonDairyBookingCustomer = "ShowBookingTypeDropDownonDairyBookingCustomer"
    'Public Const AutoCreateGateEntryTillMilkTransferInForIntermittent = "AutoCreateGETillMilkTransferInForIntermittent"
    Public Const AllowItemConversionAutomation = "AllowItemConversionAutomation"
    Public Const AutoMilkTransferInDateSameasWeighmentDate = "AutoMilkTransferInDateSameasWeighmentDate"
    Public Const EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt = "EnableGoButtonofRcptEntryWithoutEnteringReceiptAmt"
    Public Const ShowOutstandingAmtofCustomerOnQuickBookEntry = "ShowOutstandingAmtofCustomerOnQuickBookEntry"
    Public Const DoNotCreateJournalVoucheronJobWorkDispatch = "DoNotCreateJournalVoucheronJobWorkDispatch"
    Public Const PickCostOFMaterialSaleFromPriceMaster = "PickCostOFMaterialSaleFromPriceMaster"
    Public Const ForceToSelectIteminGateEntry = "ForceToSelectIteminGateEntry"
    Public Const CreateProvisionforBulkContractorInGateIn = "CreateProvisionforBulkContractorInGateIn"
    Public Const DoNotCreateAdjustmentonMilkTransferInGL = "DoNotCreateAdjustmentonMilkTransferInGL"
    Public Const Donotshowtrasnfertransactionsbydefault = "Do not show trasnfer transactions by default"
    Public Const AllowManualItemPriceOnMCCSale = "AllowManualItemPriceOnMCCSale"
    Public Const ExcludeKKFAndMandiForDCS = "Exclude KKF And Mandi For DCS"
    Public Const DonotAllowtoChangeUOMinDairyBookingCustomer = "DonotAllowtoChangeUOMinDairyBookingCustomer"
    Public Const AutoPopulateItemCodeOnDairyBooking = "AutoPopulateItemCodeOnDairyBooking"
    Public Const ItemwiseCorrectionFactoronQC = "ItemwiseCorrectionFactoronQC"
    Public Const Allow0FatPerOnBulkSaleQualityCheckScreen = "Allow0FatPerOnBulkSaleQualityCheckScreen"
    Public Const CreateMultipleDispatchWithoutSelectingVehicle = "CreateMultipleDispatchWithoutSelectingVehicle"
    Public Const RejectiononQCforSeparationofBulkProcurementMCC = "RejectiononQCforSeparationofBulkProcurementMCC"
    Public Const ParameterForSNFatQC = "ParameterForSNFatQC"
    Public Const TreatCLR50As52 = "Treat CLR 50 As 52"
    Public Const AllowmultipleconsumptionLocation = "Allow Multiple Consumption Location"
    Public Const PickCostFromItemMaster As String = "Pick Cost From Item Master(FG)"
    Public Const EditItemCost As String = "Edit Item Cost"
    Public Const AllowBatchQtyin3DecimalPlaces As String = "Allow Batch Qty in 3 Decimal Places"
    Public Const DisplayAverageFatSNFMPWise As String = "DisplayAverageFatSNFMPWise"
    Public Const SentschemecogsinAnotherAccount As String = "SentschemecogsinAnotherAccount"
    Public Const OPkmMandatoryonDS As String = "OP km mandatory on Dairy dispatch sale"
    Public Const ApplyBankChargesasperSlabonBankMaster As String = "ApplyBankChargesasperSlabonBankMaster"
    Public Const UseKGLitreConversionInBulkSaleAsperCLRCalculation As String = "UseKGLitreConversionInBulkSaleAsperCLRCalculation"
    Public Const ShowLastUnitCostZeroForNonInventoryItemOnPO As String = "ShowLastUnitCostZeroForNonInventoryItemOnPO"
    Public Const ManualBatchNoMandatoryOnBatchOrderScreen As String = "ManualBatchNoMandatoryOnBatchOrderScreen"
    Public Const ShowSiloLocationItemLocationwise = "ShowSiloLocationItemLocationwise"
    Public Const GateEntryChamberwisewithManualTankerEntry = "GateEntryChamberwisewithManualTankerEntry"
    Public Const AllowPriceMappingOnBulkSRNinChamberBulkProc = "AllowPriceMappingOnBulkSRNinChamberBulkProc"
    Public Const EnableManualCrateonTaxableDairyDispatch = "EnableManualCrateonTaxableDairyDispatch"
    Public Const CreateOpeningEntryAutomatically = "Create Opening Entry Automatically"
    Public Const AllowCratePhysicalStock = "AllowCratePhysicalStock"
    Public Const AllowSiloMilkTransfertoMainLocation = "AllowSiloMilkTransfertoMainLocation"
    Public Const EnableAutoDocNoShipToLocation = "Enable Auto DocNo ShipToLocation"
    Public Const ChangeFATCLRafterspecialApprovalonQC = "Change FAT CLR after special approval on QC"
    Public Const CreateCommonSeriesLocationwiseForAllSale = "Create Common Series Location Wise For All Sale"
    Public Const EnableCustomerPODetailonDairyBooking = "EnableCustomerPODetailonDairyBooking"
    Public Const CreateCommonDairyDispatchforFreshAmbient = "CreateCommonDairyDispatchforFreshAmbient"
    Public Const CreateSeperateTaxInvForFOCIteminNonTaxdispatch = "CreateSeperateTaxInvForFOCIteminNonTaxdispatch"
    Public Const CreateSeperateSeriesforRefDocARinvforCreditdebit = "CreateSeperateSeriesforRefDocARinvforCreditdebit"
    Public Const CreateSeperateSeriesforRefDocAPinvforCreditdebit = "CreateSeperateSeriesforRefDocAPinvforCreditdebit"
    Public Const AllowUniqueNoOnMilkTransferInandTankDis = "Allow UniqueNo On Milk TransferIn and Tanker Dispatch"
    Public Const DoNotConsiderCustomerCreditLimit = "Do not Consider Customer Credit Limit"
    Public Const AllowVSPMasterAutoPrefix = "Allow VSP Master Auto Prefix"
    Public Const EnableGSTRelatedfields = "EnableGSTRelatedfields"
    Public Const AllowPostGSTPayment = "AllowPostGSTPayment"
    Public Const ConsiderSiloCapicityForStockIn = "ConsiderSiloCapicityForStockIn"
    Public Const AllowBulkProcTransDateSameasGateEntryDate = "Allow BulkProc TransDate Same as GateEntry Date"
    Public Const AllowDifferentStateofChildCustomerOnPS = "Allow Different State of Child Customer On PS"
    Public Const AllowProvisionknokoffOnAPInvoice = "Allow Provision knokoff On APInvoice"
    Public Const AllowTransactionFiltersOnCustomerlegder = "AllowTransactionFiltersOnCustomerlegder"
    Public Const GroupCustomerlegderZoneWiseAreaWise = "GroupCustomerlegderZoneWiseAreaWise"
    Public Const CustomerDashboardWithOpeningAndClosing = "CustomerDashboardWithOpeningAndClosing"
    Public Const AllowtoSHOWParentChildCustomer = "AllowtoSHOWParentChildCustomer"
    Public Const AllowtoMakeApplyDocOnbyDefault = "AllowtoMakeApplyDocOnbyDefault"
    Public Const PenaltyPercentage = "Penalty%"
    Public Const CreateProvisionJournalEntryForSale = "Create Provision JournalEntry Sale"
    Public Const CreateProvisionJournalEntryForTankerDispatch = "Create Provision JournalEntry Tanker Dispatch"
    Public Const ShowDairySaleModuleOnBulkPosting = "Show DairySale Module On BulkPosting"
    Public Const ShowGateEntryTypeonGateEntryBulkProc = "Show GateEntryType on GateEntry BulkProc"
    Public Const AllowAutoBulkMilkSRNonWeighmentBulkProc = "Allow Auto BulkMilkSRN on Weighment BulkProc"
    Public Const PickCorrectionFactorProcurementTypewise = "Pick CorrectionFactor ProcurementType wise"
    Public Const CheckParameterRangerProcurementTypewise = "Check ParameterRanger ProcurementType wise"
    Public Const CalculateTaxRatefromItemwsieTaxOnSale = "CalculateTaxRatefromItemwsieTaxOnSale"
    Public Const AllowBulkProcMCCwithoutTankerDispatch = "Allow Bulk Proc MCC without Tanker Dispatch"
    Public Const AllowJobWorkonGateEntryBulkProc = "Allow Job Work ON Gate Enty BulkProc"
    Public Const ApplyMaTransRateOnMultChamberTankerDis = "ApplyMaTransferRateOnMultilpeChamberTankerDispatch"
    Public Const AllowManualPriceONBulkPO = "Allow Manual Price ON Bulk PO"
    Public Const AllowReverseUnpost = "Allow Reverse and Unpost"
    Public Const AllowtoSetNoOfTransactionsforSetOff = "AllowtoSetNoOfTransactionsforSetOff"
    Public Const AllowtoUnlockTransactionsforSetOff = "AllowtoUnlockTransactionsforSetOff"
    Public Const AllowtoSetReceiptAmountForCashTransaction = "AllowtoSetReceiptAmountForCashTransaction"
    Public Const AllowtoSetPaymentAmountForCashTransaction = "AllowtoSetPaymentAmountForCashTransaction"
    Public Const AllowtoShowCreditBalanceonCustomerAgeing = "Allow to Show Credit Balance"
    Public Const AllowtoShowDebitBalanceonVendorAgeing = "Allow to Show Debit Balance"
    Public Const ConsiderOpeningDocintoBucketsInAgeing = "ConsiderOpeningDocintoBucketsInAgeing"
    Public Const ConsiderOpeningDocintoBucketsonCustomerAgeing = "ConsiderOpeningDocintoBucketsonCustomerAgeing"
    Public Const AllowtoNegativeStockInventoryAtTankerDispatch = "AllowtoNegativeStockInventoryAtTankerDispatch"
    Public Const AllowtoNegativeFATSNFKgAtTankerDispatch = "AllowtoNegativeFATSNFKgAtTankerDispatch"
    Public Const AllowtoSkipfunctionalityafterSRNOnBulkProcurement = "AllowtoSkipfunctionalityafterSRNOnBulkProcurement"
    Public Const AllowtoSetoffDocDateWise = "AllowtoSetoffDocDateWise"
    Public Const AllowtoSkipJournalEntryofPaymentandReceiptforAD = "AllowtoSkipJournalEntryofPaymentandReceiptforAD"
    Public Const AllowtoEmployeeSalaryIntegration = "AllowtoEmployeeSalaryIntegration"
    Public Const AllowtoFutureDateTransForPDCCheque = "AllowtoFutureDateTransForPDCCheque"
    Public Const WeightOfCanForCanSale = "Weight Of Can For Can Sale"
    Public Const RunBulkProcWithoutMilkGrade As String = "Run Bulk Proc without Milk Grade"
    Public Const DaysToStartAutoLock As String = "Days ToStart AutoLock"
    Public Const AllowRandomOnlyOneSecondaryQC As String = "Allow Random OnlyOne SecondaryQC"
    Public Const AllowGateEntryAgainstPO As String = "Allow GateEntry Against PO"
    Public Const SeparateDairyDispatchTaxableNonTaxable As String = "Separate DairyDispatch TaxableNon Taxable"
    Public Const RunBatchFifowise As String = "Run Batch Fifowise"
    Public Const RunBatchFifowisewithModifyfunctionality As String = "Run Batch Fifowise with Modify functionality"
    Public Const PromptTimeToPostTransactions As String = "Prompt Time ToPost Transactions"
    Public Const CreateFreshInvoiceOnDispatchSave As String = "Create FreshInvoice OnDispatch Save"
    Public Const AllowLockTransactionUserwise As String = "AllowLockTransactionUserwise"
    Public Const AllowAutoLockTransaction As String = "AllowAutoLockTransaction"
    Public Const AllowDefaultBankCodeforCreditNote As String = "AllowDefaultBankCodeforCreditNote"
    Public Const AllowUseApplyDocSeriesForReceipt As String = "AllowUseApplyDocSeriesForReceipt"
    Public Const AllowUseApplyDocSeriesForPayment As String = "AllowUseApplyDocSeriesForPayment"
    Public Const AllowCreditNoteWithoutReference As String = "AllowCreditNoteWithoutReference"
    Public Const AllowBranchAcconReceiptPrint As String = "AllowBranchAcconReceiptPrint"
    Public Const AllowCreditNoteWithoutReferenceonAP As String = "AllowCreditNoteWithoutReferenceonAP"
    Public Const SecurityDocumentKnockOffonReceipt As String = "SecurityDocumentKnockOffonReceipt"
    Public Const AllowFreshInvoiceAutoPost As String = "AllowFreshInvoiceAutoPost"
    Public Const AllowReceiptThroughSO As String = "AllowReceiptThroughSO"
    Public Const AllowSetOffUntilTransactionsnotend As String = "AllowSetOffUntilTransactionsnotend"
    Public Const AllowGateReturn As String = "AllowGateReturn"
    Public Const CalculateCommOnCSATransWOConversion As String = "CalculateCommOnCSATransWOConversion"
    Public Const AllowSRNWithoutShortageRejection As String = "AllowSRNWithoutShortageRejection"
    Public Const AllowPurchaseModulewithUniqueItem As String = "AllowPurchaseModulewithUniqueItem"
    Public Const GrossWeightUnit As String = "GrossWeightUnit"
    Public Const ExpiryDaysBulkProcurementPriceChart As String = "ExpiryDays BulkProcurement PriceChart"
    Public Const ShowSchemeItemRateonDairyDispatch As String = "Show Scheme Item Rate on Dairy Dispatch"
    Public Const ShowSchemeItemRateonDairyDispatchTaxable As String = "Show Scheme Item Rate on Dairy Dispatch Taxable"
    Public Const AutoCalculateCrateOnDairyDispatch As String = "Auto Calculate Crate on Dairy Dispatch"
    Public Const AutoCalculateCANOnDairyDispatch As String = "Auto Calculate Can on Dairy Dispatch"
    Public Const GrossWtFromItemMasterONProductSale As String = "Gross Wt. from item master on Product Sale"
    Public Const CreateVatSeriesForProductExciseinvoice = "Create Vat Series for PS Excise invoice"
    Public Const AllowFreshPriceChartOnProductSale = "Allow Fresh Price Chart on Product Sale"
    Public Const AllowFreshPriceChartOnBookingProductSale = "Allow Fresh Price Chart on Booking PS"
    Public Const ShowUnloadingandWeighmentSequencewise = "Show Unloading Weighment sequence wise"
    Public Const ShowBothTankertypeOnCleaning = "Show Both Tanker Type on Cleaning"
    Public Const isCleaningMandatoryBeforeGateout = "Is Cleaning Mandatory before Gate Out"
    Public Const AllowBulkProcurementSequencewise = "Allow Bulk Procurement Sequence wise"
    Public Const ShowItemLocationWiseonDairyBooking = "Show Item Location wise on Dairy Booking"
    Public Const SeprateDemandForMorningEveningShift = "Seprate Demand For Morning Evening Shift"
    Public Const CheckOutstandingCreditLimitOnBooking = "Check Customer Outstanding on Booking"
    Public Const AllowBulkPriceChartMultiplepriceToMultipleVendor = "Allow BulkPrice Multiple Price to Mult Vendor"
    Public Const ShowOptionOnItemMasterChangeItemRate = "Show option on Item Change Rate on DDispatch"
    Public Const SHowOptionOnLocationForDairyDispatchfromDOorGatepass = "Show option on loc For DDispatch from DO/GP"
    Public Const showPostrequiredforBulkSale = "showPostrequiredforBulkSale"
    Public Const ApplyDocumentDate = "Apply Document Date"
    Public Const AllowStockCheckatDOLevel = "AllowStockCheckatDOLevel"
    Public Const AllowAdditionalWeightinPercentage = "AllowAdditionalWeightinPercentage"
    Public Const EnterAdditionalWeight = "EnterAdditionalWeight"
    Public Const AllowTankerBasedonVendorofGE = "Allow Tanker Based on Vendor of GE"
    Public Const AllowUseBoilingParameteronParameterMaster = "Allow Use Boiling Parameter on Parameter Master"
    Public Const DairyDispatchFromDeliveryNote = "Dairy Dispatch From Delivery Note"
    Public Const ItemTypeForDairyBooking = "Item Type Fresh or Ambient For Dairy Booking"
    Public Const AllowStockToleranceNegative = "Allow Stock Tolerance Negative"
    Public Const StockToleranceLimit = "Stock Tolerance Limit"
    Public Const ShowAllMenu As String = "Show All Menu"
    Public Const POSeriesWithoutItemTypewise As String = "POSeriesWithoutItemTypeWise"
    Public Const WorkApprovalFlowInERP As String = "Work Approval Flow in ERP"
    Public Const OpenAvailorEmptyStckLocationOn_Standardization As String = "Open Avail./Empty Location on Standardization"
    Public Const BOM_Amend_Pswd As String = "Amendment Password for BOM"
    Public Const ProductionFATSNF_KG_Unit As String = "ProductionFATSNF_KG_Unit"
    Public Const RemoveGridFatSnfValidation As String = "Remove Grid Fat Snf Validation"
    Public Const ChangeRateAT_CSA_Return As String = "Rate Change At CSA Transfer Return"
    Public Const VehicleCapacityUnit As String = "VehicleCapacityUnit"
    Public Const StopGLEntryForConsignmentAtCSATransfer As String = "Stop GL Consignment at CSA Transfer"
    Public Const CSATransfer_SalePatti_All_Tax_Open As String = "Open All Tax Mapped With Location In CSA"
    Public Const CreateGLForTransfer As String = "CreateGLForTransfer"
    Public Const Active As String = "Active"
    Public Const Agency As String = "Agency"
    Public Const Category1 As String = "Category1"
    Public Const Category2 As String = "Category2"
    Public Const Category3 As String = "Category3"
    Public Const cboPeriodicity As String = "cboPeriodicity"
    Public Const CboRound As String = "CboRound"
    Public Const QuickExport As String = "Q_EX"
    Public Const UploaderPassword As String = "Uploader Password"
    Public Const BankUploaderPassword As String = "Bank Uploader Password"
    Public Const AllowDesignAtRunTime As String = "AllowDesignRunTime"
    Public Const SkipDiffGLOnPI As String = "SkipDiffGLOnPI"
    Public Const SkipCogsEntry As String = "SkipCogsEntry"
    Public Const ShowPurchaseControlAc As String = "ShowPurchaseControlAc"
    Public Const CreatePOWithRequisition As String = "CreatePOWithRequisition"
    Public Const DisplayReasonOnDelete As String = "DisplayReasonOnDelete"
    Public Const DisplayReasonOnUpdateAfterPost As String = "DisplayReasonOnUpdateAfterPost"
    Public Const Importbulkdatafromexcelsheet As String = "Importbulkdatafromexcelsheet"
    Public Const Distributor As String = "Distributor"
    Public Const Helper As String = "Helper"
    Public Const Driver As String = "Driver"
    Public Const ZM As String = "ZM"
    Public Const TSM As String = "TSM"
    Public Const ASM As String = "ASM"
    Public Const PROVISIONENTRYONSTOCKTRANSFER As String = "ProvisionOnStockTransfer"
    Public Const Inactive As String = "Inactive"
    Public Const ALLOWANYBO As String = "Allow Any Type of BO"
    Public Const ALLOWCBOSBO As String = "Allow Child and SubChild BO"
    Public Const INDUSTRYTYPE As String = "Industry Type"
    Public Const Others As String = "Others"
    Public Const PayHeadSubHead As String = "PayHeadSubHead"
    Public Const PZ As String = "PZ"
    Public Const ROUTE As String = "ROUTE"
    Public Const Route1 As String = "Route"
    Public Const Salesman As String = "Salesman"
    Public Const SIR As String = "SIRC"
    Public Const SIRC As String = "SIRC"
    Public Const PP_MRP As String = "PP_MRP"
    Public Const MulProcDedReversAndCreate As String = "MulProcDedReversAndCreate"
    Public Const WEUpdateAfterPost As String = "WEUpdateAfterPost"
    Public Const GEUpdateAfterPost As String = "GEUpdateAfterPost"
    Public Const PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE As String = "PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE"
    Public Const GEUpdatePriceChart As String = "GEUpdatePriceChart"
    Public Const SetCSATransferwithZeroOnSalePatti As String = "SetCSATransferwithZeroOnSalePatti"
    Public Const POAmendmentType As String = "POAmendmentType"
    Public Const BulkInvoiceDeleteType As String = "BulkInvoiceDeleteType"
    Public Const BulkSaleSequence As String = "BulkSaleSequence"
    Public Const BulkQCTableHavingUniqueKey As String = "BulkQCTableHavingUniqueKey"
    Public Const SrPath As String = "SrPath"
    Public Const TempProvisional As String = "TempProvisional"
    Public Const TF As String = "TF"
    Public Const Z As String = "Z"
    Public Const ZA As String = "ZA"
    Public Const ZB As String = "ZB"
    Public Const ZC As String = "ZC"
    Public Const ZD As String = "ZD"
    Public Const ZE As String = "ZE"
    Public Const ZF As String = "ZF"
    Public Const MilkProc As String = "MilkProc"
    Public Const EnablePopupItemReorderLevel As String = "EnablePopupItemReorderLevel"
    Public Const PrintVerify As String = "Print Verify"
    Public Const LOReceiptDefaultBankForSettlement As String = "Default Bank For Settlement"
    Public Const LOReceiptPaymentTypeForSettlement As String = "Default Payment Type For Settlement"
    Public Const DefaultValue As String = "DefaultValue"
    Public Const SalesmanPhysicalLocation As String = "SPL"
    Public Const IndentTolerence As String = "IndentTolerence"
    Public Const AskForDate As String = "AskForDate"
    Public Const PickMachineDateForTran As String = "PickMachineDateForTran"
    Public Const ReqLimitOnSRN As String = "ReqLimitOnSRN"
    Public Const AutoLoadinFromLocation As String = "AutoLoadinFromLocation"
    Public Const CrreateTransferShipmentJE As String = "CrreateTransferShipmentJE"
    Public Const IsNotIncludeWasteQtyInCal As String = "IsNotIncludeWasteQtyInCal"
    Public Const IsConsiderOutTypeDocForBalance As String = "IsConsiderOutTypeDocForBalance"
    Public Const BankTransferRunPaymentCounter As String = "BankTransferRunPaymentCounter"
    Public Const PaymentReceiptTypeRunReceiptCounter As String = "PaymentReceiptTypeRunReceiptCounter"
    Public Const CounterFinancialYearStyle As String = "CounterFinancialYearStyle"
    Public Const LinkFinancialYearStyleWithGSTDate As String = "LinkFinancialYearStyleWithGSTDate"
    Public Const ApplyLocationWisePrefix As String = "Apply Location Wise Prefix"
    Public Const CashDiscountFromClaimMaster As String = "CashDiscountFromClaimMaster"
    Public Const TransferTransTypeRouteHide As String = "TransferTransTypeRouteHide"
    Public Const AllowNegtiveOfSaleInvoiceBlanceAmt As String = "AllowNegtiveOfSaleInvoiceBlanceAmt"
    Public Const SalesRateEditable As String = "Sales Rate Editable"
    Public Const RunDemoERP As String = "RunDemoERP"
    Public Const IsKDIL As String = "IsKDIL"
    Public Const SendToTally As String = "SendToTally"
    Public Const PromptForTally As String = "PromptForTally"
    Public Const CurrentMaufacturingType As String = "ManufacturingType"
    Public Const TallyCompany As String = "TallyCompany"
    Public Const TallyIP As String = "TallyIP"
    Public Const TallyPort As String = "TallyPort"
    Public Const TaxRoundOffToZeroDecimalPlace As String = "TaxRoundOffToZeroDecimalPlace"
    Public Const BalanceSheetProftAndLossGroupCode As String = "BalanceSheetProftAndLossGroupCode"
    Public Const BalanceSheetProftAndLossGroupDesc As String = "BalanceSheetProftAndLossGroupDesc"
    Public Const ApplyCostingOnPostedDate As String = "ApplyCostingOnPostedDate"
    Public Const isBatchApplyOnInventoryMovement As String = "isBatchApplyOnInventoryMovement"
    Public Const BlankDatabase As String = "BlankDatabase"
    Public Const ServiceDealer As String = "Service Dealer"
    Public Const TDM As String = "TDM"
    Public Const MAILOFF As String = "MAILOFF"
    ''added by shivani
    Public Const AllowToSaveTimeWithDocumentDate As String = "Allow To Save Time With Document Date"
    Public Const AllowToPrintTimeWithDocumentDate As String = "Allow To Print Time With Document Date"
    Public Const AllLevelApprovalIsMandatory As String = "All Level Approval Is Mandatory"
    Public Const AssetGroupPrefix As String = "AssetGroupPrefix"
    Public Const DepreciationCalculationMethod As String = "Depreciation Calculation Method"
    Public Const STDPURRATE As String = "STDPURRATE"
    Public Const AutoPOAtSRN As String = "AUTOPOATSRN"
    Public Const DisableShipToLocation As String = "Disable Ship_To_Location For (PO,PI,SRN)"
    Public Const AllowLargerItemCostThenVendorItemCost As String = "Allow Larger Item Cost Then Vendor Item Cost"
    Public Const AcceptQtyWithoutConsiderRejectQty As String = "Accept Qty Without Consider Reject Qty"
    Public Const PurchasePickItemFromVendorItemDetails As String = "PurchasePickItemFromVendorItemDetails"
    Public Const PurchaseOneItemOneVendor As String = "PurchaseOneItemOneVendor"
    Public Const PostShipmentonAutoSTN As String = "PostShipmentonAutoSTN"
    Public Const IsRemarksMandatoryOnCloseSaleOrder As String = "IsRemarksMandatoryOnCloseSaleOrder"
    Public Const LCCancellationPwd As String = "LCCancellationPwd"
    Public Const ShowQtySum_in_GRN_MRN_SRN As String = "ShowQtySum_in_GRN_MRN_SRN"
    Public Const CreateInvoicewithShipmentonAutoSTN As String = "CreateInvoicewithShipmentonAutoSTN"
    Public Const AllowSingleInvoiceAgainstSingleOrder As String = "AllowSingleInvoiceAgainstSingleOrder"
    Public Const WorkingHours As String = "WorkingHours"
    Public Const TreatExcessLeaveAbsent As String = "TreatExcessLeaveAbsent"
    Public Const VehicleInsuranceAlert As String = "VehicleInsuranceAlert"
    Public Const IsItemRateEditableOnTransfer As String = "IsItemRateEditableOnTransfer"
    Public Const GLACAccordingToTaxRate As String = "GLACAccordingToTaxRate"
    Public Const AutoSchemeOn As String = "AutoSchemeOn"
    Public Const MultiplySubsidyWithQuantity As String = "Multiply Subsidy With Quantity"
    Public Const CreateAutoReceiptEntryDCSSale As String = "CreateAutoReceiptEntryDCSSale"
    Public Const GeneratePDFForMobile As String = "Generate PDF For Mobile"
    Public Const IsTransferQtyEditableOnAutoSTN As String = "IsTransferQtyEditableOnAutoSTN"
    Public Const IsItemRateEditableOnSales As String = "IsItemRateEditableOnSales"
    Public Const IsItemMRPEditableOnSales As String = "IsItemMRPEditableOnSales"
    Public Const IsItemRateEditableOnSalesForAprilOnly As String = "ForAprilOnly"
    Public Const PWD As String = "PWD"
    Public Const AllowMilkReceiptAfterSettingsisOn As String = "AllowMilkReceiptAfterSettingsisOn"
    Public Const BMCMilkQCManualPassword As String = "BMCMilkQCManualPassword"
    Public Const MilkReceiptTolerancePwd As String = "MilkReceiptTolerancePwd"
    Public Const MCC_DLTDATA_PWD As String = "MCC_DLTDATA_PWD"
    Public Const Allow_Excel_Code_on_Mcc_Master As String = "AllowExCodeONMcc"
    Public Const is_Allow_cancel_Transaction As String = "is_Allow_cancel_Transaction"
    Public Const is_Allow_cancel_Posted_Transaction As String = "is_Allow_cancel_Posted_Transaction"
    Public Const ShiftTiming As String = "ShiftTiming"
    '=shivani
    Public Const MulticurrencyDecimalPlaces As String = "MulticurrencyDecimalPlaces"
    Public Const SMS_User_Name As String = "SMS_User_Name"
    Public Const SMS_User_PWD As String = "SMS_User_PWD"
    Public Const SMS_Sendor_ID As String = "SMS_Sendor_ID"
    Public Const SMS_Provider As String = "SMS_Provider"
    Public Const MCCDefaultMilkItem As String = "MCCDefaultMilkItem"
    Public Const MCCDefaultMilkItemCow As String = "MCC Default Milk Item Cow"
    Public Const MCCDefaultMilkItemBuffalo As String = "MCC Default Milk Item Buffalo"
    Public Const BulkSaleDefaultMilkItem As String = "BulkSaleDefaultMilkItem"
    Public Const BSDefaultMilkItem As String = "BSDefaultMilkItem"
    Public Const DefaultRoundOffGLAccount As String = "DefaultRoundOffGLAccount"
    Public Const MCCSampleRange As String = "MCCSampleRange"
    Public Const MCCReceiptRange As String = "MCCReceiptRange"
    Public Const MCCMinKmRange As String = "MCCMinKmRange"
    Public Const Milk_Can_Weight_Ratio As String = "Can_Weight_Ratio"
    Public Const Milk_Can_Weight_Tolerance_Neg As String = "Can_Weight_Tolerance_Neg"
    Public Const Milk_Can_Weight_Tolerance_Positive As String = "Can_Weight_Tolerance_Positive"
    Public Const MCCFSSAI_DAYS As String = "MCCFSSAI_DAYS"
    Public Const MCCDisplay_All_Parameter As String = "MCCDis_P"
    '==========Rohit on 29,Oct 2014================
    Public Const MCCInvoiceScheduleDate As String = "MCCInvoiceScheduleDate"
    Public Const MCCInvoiceScheduleTime As String = "MCCInvoiceScheduleTime"
    Public Const MCCInvoiceScheduleInterval As String = "MCCInvoiceScheduleInterval"
    Public Const MCCMilkSRNRepost As String = "MCCMilkSRNRepost"
    '==============================================
    '===========Rohit on Jan 31,2015=====
    Public Const Is_Send_Sms As String = "Is_Send_Sms"
    Public Const Send_Sms_Time As String = "Send_Sms_Time"
    Public Const Is_Send_Sms_ForVSP As String = "Is_Send_Sms_ForVSP"
    Public Const Is_Pick_No_from_Mail_Setting As String = "Is_Pick_No_from_Mail_Setting"
    '=====================================
    '------Pankaj Jha
    Public Const ControlSampleMandatory As String = "ControlSampleMandatory"
    Public Const defaultCorrectionFactor As String = "defaultCorrectionFactor"
    '------End 
    Public Const ShowTaxRateColumnOnTransaction As String = "ShowTaxRateColumnOnTransaction"
    Public Const ShowGRN As String = "ShowGRN"
    Public Const SkipMRNGRNinCaseofMT As String = "SkipMRNGRNinCaseofMT"
    Public Const AutoGenerateMRN As String = "AutoGenerateMRN"
    Public Const ShowMRN As String = "ShowMRN"
    Public Const LicenceExpiryDate As String = "IsApplyCommonService1" 'A B
    Public Const LicenceNoOfExeConnection As String = "IsApplyCommonService2" 'C
    Public Const LicenceNoOfJournalEntry As String = "IsApplyCommonService3" 'D
    Public Const LicenceNoOfUser As String = "IsApplyCommonService4" 'E
    Public Const EnableProjectFinder As String = "EnableProjectFinder"
    'richa 
    Public Const InvoiceManualNoWithPrefix As String = "InvoiceManualNoWithPrefix"
    Public Const AutoBackUp As String = "AutoBackUp"
    Public Const MCCPurchase As String = "MCCPurchase"
    'richa Ticket No BM00000003045 09/07/2014
    Public Const NotificationSettingforReOrderInPO As String = "NotificationSettingforReOrderInPO"
    'richa Ticket No BM00000003042 09/07/2014
    Public Const NotificationSettingforReOrderInPurchaseRequisition As String = "NotificationSettingforReOrderInPurchaseRequisition"
    Public Const PurchaseOrderAutomaticallyItemQtyBelowReorderLevel As String = "PurchaseOrderAutomaticallyItemQtyBelowReorderLevel"
    Public Const NLevelAtVendor As String = "NLevel_Vendor"
    Public Const NLevelAtCustomer As String = "NLevel_Customer"
    Public Const NLevelAtLocation As String = "NLevel_Location"
    Public Const AutoItemNLevel As String = "NLevel_ItemCode"
    Public Const Princi_Bom As String = "Principle_BOM"
    Public Const AP_INV_COMMSN As String = "AP_INV_COMMSN"
    Public Const Principal_Vendor As String = "Principal_Vendor"
    Public Const Principal_Vendor_Database As String = "Principal_Vendor_Database"
    Public Const Principal_Customer As String = "Principal_Customer"
    Public Const CalculateLTAOnHoliday As String = "CalculateLTAOnHoliday"
    Public Const CalculateLTAOnWeekend As String = "CalculateLTAOnWeekend"
    Public Const CalculateMediclaimOnHoliday As String = "CalculateMediclaimOnHoliday"
    Public Const CalculateMediclaimOnWeekend As String = "CalculateMediclaimOnWeekend"
    Public Const DiscountCodeForArAdj As String = "DiscountCodeForArAdj"
    Public Const DiscountCodeForMPAdj As String = "DiscountCodeForMPAdj"
    Public Const CreateAutoRecieptForManualCustomer As String = "CreateAutoRecieptForManualCustomer"
    Public Const AutoRecieptBankCode As String = "AutoRecieptBankCode"
    Public Const AutoRecieptPaymentMode As String = "AutoRecieptPaymentMode"
    Public Const Is_Purchaseable_Item As String = "Is_Purchaseable_Item"
    Public Const Is_AbemdmentForDemo As String = "Is_AbemdmentForDemo"
    Public Const Is_FinishedGoods As String = "Is_FinishedGoods"
    Public Const ShowStatusForPurchase As String = "ShowStatusForPurchase"
    Public Const ShowStatusForSales As String = "ShowStatusForSales"
    Public Const ShowSerialNoForSales As String = "ShowSerialNoForSales"
    Public Const AutoGeneratedVendorCode As String = "AutoGeneratedVendorCode"
    Public Const AutoGeneratedVendorCodeForAllCompany As String = "AutoGeneratedVendorCodeForAllCompany"
    Public Const AutoGeneratedCustomerCode As String = "AutoGeneratedCustomerCode"
    Public Const AutoGeneratedCustomerCodeForAllCompany As String = "AutoGeneratedCustomerCodeeForAllCompany"
    Public Const AllowToUseSubAccount As String = "AllowToUseSubAccount"
    Public Const InTransitFeatureIsRequired As String = "InTransitFeatureIsRequired"
    Public Const PermissionSettingForTransactionWithBank As String = "Permission_Setting_For_Trans_With_Bank"
    Public Const ApplyBrachAccounting As String = "ApplyBrachAccounting"
    Public Const AllowToEnterMRPManually As String = "AllowToEnterMRPManually"
    Public Const AllowFieldsToBeManadatory As String = "AllowFieldsToBeManadatory"
    Public Const AutoGeneratedDigitsForVendor As String = "AutoGeneratedDigitsForVendor"
    Public Const AutoGeneratedDigitsForCustomer As String = "AutoGeneratedDigitsForCustomer"
    Public Const IsRateEditableOnSRN As String = "IsRateEditableOnSRN"
    Public Const CreateGLAccToItem As String = "CreateGLAccToItem"
    Public Const IsCostEditableOnIssueReturnTransfer As String = "IsCostEditableOnIssueReturnTransfer"
    Public Const UpdateCrateLinerQty As String = "UpdateCrateLinerQty"
    'Richa Agarwal 05/08/2014 Against Ticket No BM00000003248
    Public Const AllowDispatchOutstandingBS As String = "AllowDispatchOutstandingBS"
    Public Const AllowDispatchOutstandingFS As String = "AllowDispatchOutstandingFS"
    Public Const AllowDispatchOutstandingPS As String = "AllowDispatchOutstandingPS"
    Public Const IsVolumeSchemeBydefault As String = "IsVolumeSchemeBydefault"
    Public Const AllowtoSelDateandBankforPayEntryOnSalaryGeneration As String = "AllowtoSelDateandBankforPayEntryOnSalaryGeneration"
    Public Const AllowDeliveryOrderIncaseAmountIncreases As String = "AllowDeliveryOrderIncaseAmountIncreases"
    Public Const AllowAutoMRNGRNonDocumentAcceptance As String = "AllowAutoMRNGRNonDocumentAcceptance"
    Public Const AllowToShowSaleTypeinPaymentTermsReceivable As String = "AllowToShowSaleTypeinPaymentTermsReceivable"
    Public Const AllowToShowMilkTypeinAdjustmentEntry As String = "AllowToShowMilkTypeinAdjustmentEntry"
    Public Const GatePassAfterTransfer As String = "GatePassAfterTransfer"
    Public Const CreateTransferFromBooking As String = "CreateTransferFromBooking"
    Public Const PickRateFromPRICEChrtMasterFORUMang As String = "PickRateFromPRICEChrtMasterFORUMang"
    Public Const StockTranferFromTransferPriceAndInvJVWithAvgCost As String = "StockTranfer - TransferPriceAndInvJVWithAvgCost"
    Public Const IGnoreGITAccount As String = "Ignore GIT Account in Financial Entry"
    Public Const AllowToEditCategoryCodeinItemMaster As String = "AllowToEditCategoryCodeinItemMaster"
    Public Const CreditLimitApproval As String = "CustomerCreditLimit"
    Public Const ViewTDSPwd As String = "ViewTDSPwd"
    Public Const ShowVendorLedgerasPerRightsForLocation As String = "ShowVendorLedgerasPerRightsForLocation"
    Public Const ShowCustomerLedgerasPerRightsForLocation As String = "ShowCustomerLedgerasPerRightsForLocation"
    Public Const EnableDynamicQRCodeForB2CInvoice As String = "EnableDynamicQRCodeForB2CInvoice"
    Public Const InvoiceBasedPO As String = "InvoiceBasedPO"
    Public Const AdvanceAgainstSO As String = "AdvanceAgainstSO"
    Public Const Purchase_SMSATPOST As String = "SMSATPOST_PUR"
    Public Const Sale_SMSATPOST As String = "SMSATPOST_SALE"
    Public Const AmountLimitForInvoiceBulkSale As String = "AmountLimitForInvoiceBulkSale"
    Public Const ShowSaleInvoiceNoInPOfinderInSRN As String = "ShowSaleInvoiceNoInPOfinderInSRN"
    Public Const CrateValue As String = "CrateValue"
    Public Const CommitedDefaultQty As String = "CommitedDefaultQty"
    Public Const ShowBinMapping As String = "ShowBinMapping"
    Public Const ShowPrintChallanInDairyDispatch As String = "ShowPrintChallanInDairyDispatch"
    Public Const ShowCrateJaaliBoxIntransfer As String = "Show Crate Jaali & Box In transfer"
    Public Const DefaultCorrectionFactorForBulkSale As String = "DefaultCorrectionFactorForBulkSale"
    Public Const MCCdefaultCorrectionFactorBS As String = "MCCdefaultCorrectionFactorBS"
    Public Const JOBdefaultCorrectionFactorBS As String = "JOBdefaultCorrectionFactorBS"
    Public Const PurchasedefaultCorrectionFactorBS As String = "PurchasedefaultCorrectionFactorBS"
    Public Const AllowDeliveryQtygreaterthanBookingQtyPS As String = "AllowDeliveryQtygreaterthanBookingQtyPS"
    Public Const IsPickServerDateForMultipleDispatchInvoice As String = "IsPickServerDateForMultipleDispatchInvoice"
    Public Const TabOrder As String = "TabOrder"
    Public Const LoadLoginScreen As String = "LoadLoginScreen"
    Public Const IsItemWithDifferntUnitConsiderAsOtherItem As String = "ItemWithDifferntUnitConsiderAsOtherItem"
    Public Const IsMRPWiseBalance As String = "IsMRPWiseBalance"
    Public Const showRFQ As String = "showRFQ"
    Public Const CreateDbitNoteForShortPI As String = "CreateDbitNoteForLeakAndShortPI"
    Public Const CreateDbitNoteForRejectPI As String = "CreateDbitNoteForRejectPI"
    Public Const CreateDebitNoteForUnitCost As String = "CreateDebitNoteForUnitCost"
    Public Const TransferWithProductionSale_Retail_Series As String = "CreateTransferWithProductionSale_Retail_Series"
    Public Const TransferLocalInterState As String = "Stock/CSA_Transfer_With_Local/InterState_Series"
    Public Const ProductionQtyDecimalPoint As String = "ProductionQtyDecimalPoint"
    Public Const ConsumptionQtyTollerance As String = "Consumption Qty Tollerance"
    Public Const GunnyBagTollerance As String = "Gunny Bag Tollerance"
    Public Const CheckBalanceOnSave As String = "Check Balance On Save"
    Public Const ProductionFATSNFPerDecimalPoint As String = "ProductionFATSNFPerDecimalPoint"
    Public Const ManualySelectBOMForChildBatch As String = "ManualySelectBOMForChildBatch"
    Public Const CSATransferWithProductionSale_Retail_Series As String = "CreateCSATransferWithProductionSale_Retail_Series"
    Public Const TransferJEForLocationMapping As String = "TransferJEForLocationMapping"
    Public Const AllowToDispalyAlertForBDayAnniversary As String = "AllowToDispalyAlertForBDayAnniversary"
    Public Const AllowToSendEmailForBDayAnniversary As String = "AllowToSendEmailForBDayAnniversary"
    Public Const ItemDescForTankerdispatchPrint As String = "ItemDescForTankerDispatchPrint"
    Public Const AllowPOScheduling As String = "Allow PO Scheduling"
    Public Const ERPStartDate As String = "ERPStartDate"
    Public Const ItemBatchWiseStartDate As String = "ItemBatchWiseStartDate"
    Public Const CreateJEForTransfer As String = "CreateJEForTransfer"
    Public Const AllowToSkipStageQLLogSheetInProd As String = "AllowToSkipStageQLLogSheetInProd"
    Public Const IsRemarkReasonMandatoryOnPO As String = "IsRemarkReasonMandatoryOnPO"
    Public Const ShowCostCenterAndHierarchyLevelInPurchaseModule As String = "ShowCostCenterAndHierarchyLevelInPurchaseModule"
    Public Const IsQCColumnRequiredonMRN As String = "IsQCColumnRequiredonMRN"
    Public Const ShowItemAllStructureWise As String = "ShowItemAllStructureWise"
    Public Const ShowUpto6DecimalPlaces As String = "ShowUpto6DecimalPlaces"
    Public Const EnableOnPrivateChkbox As String = "EnableOnPrivateChkbox"
    Public Const PurchaseModule As String = "Purchase Module"
    Public Const IsRGPAfterPurchaseOrder As String = "Do RGP After Purchase Order"
    Public Const AllowQualityModuleInERP As String = "On Quality Module"
    Public Const SRNReportQuantityWise As String = "SRNReportQuantityWise"
    Public Const IsCustomerGroupFieldsMandatory As String = "IsCustomerGroupFieldsMandatory"
    Public Const IsVendorGroupFieldsMandatory As String = "IsVendorGroupFieldsMandatory"
    Public Const AllowAutoNoForBackLogEntry As String = "AllowAutoNoForBackLogEntry"
    Public Const AllowDiffentSeriesExemptedItemONPS As String = "AllowDiffentSeriesExemptedItemONPS"
    Public Const DisplayFranchiseeinCustomer As String = "DisplayFranchiseeinCustomer"
    Public Const Idle As String = "Idle"
    Public Const AddressOnPaymentVoucher As String = "AddressOnPaymentVoucher"
    Public Const AllowBankDetailsManualinVM As String = "AllowBankDetailsManualinVM"
    Public Const AllowToGenerateSaleInvoiceSeriesTaxTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesTaxTypeatPS"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesRetailTypeatPS"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesExciseTypeatPS"
    Public Const AllowToGenerateSaleInvoiceSeriesTaxatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesTaxatMCCSale"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesRetailatMCCSale"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesExciseatMCCSale"
    Public Const AllowToGenerateSaleInvoiceSeriesTaxatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesTaxatMiscSale"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesRetailatMiscSale"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesExciseatMiscSale"
    Public Const ShowHierarchyAndCostCenterInAPInvoiceEntry As String = "ShowHierarchyAndCostCenterInAP"
    Public Const WeighmentNotMandatoryInMCC As String = "WeighmentNotMandatoryInMCC"
    Public Const ShowHierarchyAndCostCenterInARInvoiceEntry As String = "ShowHierarchyAndCostCenterInAR"
    Public Const PartialFADepDays As String = "PartialFADepDays"
    Public Const RateMultPartialFADepDays As String = "RateMultPartialFADepDays"
    Public Const AllowNegativeStock As String = "AllowNegativeStock"
    Public Const SendSalarySlipMailToEmployee As String = "SendSalarySlipMailToEmployee"
    Public Const DoNotMergeAPARAccount As String = "DoNotMergeAPARAccount"
    Public Const ShowVisiDetail As String = "ShowVisiDetail"
    Public Const CustomerNameUniqueOnCM As String = "CustomerNameUniqueOnCM"
    Public Const IsShortageIncludeInLandedCost As String = "IsShortageIncludeInLandedCost"
    Public Const AlowwdateChangeinPaymentEntry As String = "AlowwdateChangeinPaymentEntry"
    Public Const CreateAutoMilkRGPinBulkSRN As String = "CreateAutoMilkRGPinBulkSRN"
    Public Const DisplayAllParameterinQualityCheck As String = "DisplayAllParameterinQualityCheck"
    Public Const DisplayTypeInMilkReceipt As String = "DisplayTypeInMilkReceipt"
    Public Const AddValidationofMilkTypeinsample As String = "AddValidationofMilkTypeinsample"
    Public Const MilkRateRoundOffType As String = "Milk Rate Round Off Type"
    Public Const SubStdFatCow As String = "Sub Std FAT Per Cow"
    Public Const SubStdSNFCow As String = "Sub Std SNF Per Cow"
    Public Const SubStdFatBuff As String = "Sub Std FAT Per Buff"
    Public Const SubStdSNFBuff As String = "Sub Std SNF Per Buff"
    Public Const SubStdFatMix As String = "Sub Std FAT Per Mix"
    Public Const SubStdSNFMix As String = "Sub Std SNF Per Mix"
    Public Const FatMinCow As String = "FatMinCow"
    Public Const FatMaxCow As String = "FatMaxCow"
    Public Const SNFMinCow As String = "SNFMinCow"
    Public Const SNFMaxCow As String = "SNFMaxCow"
    Public Const FatMinBuff As String = "FatMinBuff"
    Public Const FatMaxBuff As String = "FatMaxBuff"
    Public Const SNFMinBuff As String = "SNFMinBuff"
    Public Const SNFMaxBuff As String = "SNFMaxBuff"
    Public Const FatMinMix As String = "FatMinMix"
    Public Const FatMaxMix As String = "FatMaxMix"
    Public Const SNFMinMix As String = "SNFMinMix"
    Public Const SNFMaxMix As String = "SNFMaxMix"
    Public Const AddIncentiveDeductioninMilkSample As String = "AddIncentiveDeductioninMilkSample"
    Public Const AllowManualEnterinWeighment As String = "AllowManualEnterinWeighment"
    Public Const SettlementBankOnlyPWD As String = "SettlementBankOnlyPWD"
    Public Const DocumentSequence As String = "DocumentSequence"
    Public Const BOOKINGFINDER_ON_CSASALEPATTI As String = "CSA Sale Patti With Booking Knock-off"
    Public Const AllowPurchaseAccounting As String = "AllowPurchaseAccounting"
    Public Const SHowBulkMilkWeighment As String = "SHowBulkMilkWeighment"
    Public Const StoreADJExportImportAfterPost As String = "StoreADJExportImportAfterPost"
    Public Const FatSNFControlOnProductionConsumption As String = "FatSNFControlOnProductionConsumption"
    Public Const QuantityControlToleranceOnProductionConsumption As String = "QuantityControlToleranceOnProductionConsumption"
    Public Const LeaveBalanceAlertTypeOnAttendance As String = "LeaveBalanceAlertTypeOnAttendance"
    Public Const StopNegativeBankBalance As String = "StopNegativeBankBalance"
    Public Const ConsumptionType As String = "ConsumptionType"
    Public Const ValidateFatSnfOnProduction As String = "ValidateFatSnfOnProduction"
    Public Const ShowOverheadCostOnProductionEntry As String = "ShowOverheadCostOnProductionEntry"
    Public Const ActivateProductionWithoutBatch As String = "ActivateProductionWithoutBatch"
    Public Const CreateJEOnProduction As String = "CreateJEOnProduction"
    ''===================================Setting for Payroll=========================================
    Public Const AllowToSaveMultipleEmployeeStatus As String = "AllowToSaveMultipleEmployeeStatus"
    Public Const CreateJEForProvisionEntry As String = "CreateJEForProvisionEntry"
    Public Const DoubleClickOnVC As String = "Double Click On VC"
    Public Const PickManual_CSATransfer_OnTRansferReturn As String = "CSA Transfer Effect on Return is Manual"
    Public Const PickManual_CSATransfer_OnCSASalePatti As String = "CSA Transfer Effect on Sale Patti is Manual"
    Public Const AllowDistributorSaleAtCSA_SaleInvoice As String = "Allow Distributor Sale at CSA Sale Patti"
    Public Const AllowItemWiseCSAAccountingON_CSASale As String = "CSA Account set pick Item-wise"
    Public Const IsAutoTankerWeightment As String = "Auto Tanker Weightment"
    Public Const IsAutoTankerWeighmentForBulkSale As String = "Auto Tanker Weighment for Bulk Sale"
    Public Const IsAdditionalInformationOnVillageMaster As String = "Show Village Add Info"
    Public Const CheckLiveStockInProductionDuringTrans As String = "CheckLiveStockInProductionDuringTrans"
    Public Const VLCTimeTableColumnShow As String = "VLCTimeTableColumnShow"
    Public Const VLCTimeTableColumnMandatory As String = "VLCTimeTableColumnMandatory"
    Public Const ApplyEffectiveStartDate As String = "Apply Effective Start Date"
    Public Const isOneMCCOnePrimaryTranporter As String = "One MCC One Primary Tranporter"
    Public Const isIntimationRequired As String = "Show Intimation Screen"
    Public Const QualityThenWeighmentinBulkProcurement As String = "First QC then Weighment"
    Public Const GateEntryTankerFromTankerMaster As String = "Gate Entry tanker From Master"
    Public Const isItemMilkType As String = "Is Item Milk Type"
    Public Const isPriceChartGradeWise As String = "Is Price Chart Grade Wise"
    Public Const isCreateBulkProcPriceChartItemWise As String = "Create Bulk Procurement price chart-Itemwise"
    Public Const isFarmerPaymentCycle As String = "is Farmer Payment Cycle"
    Public Const CowFATPer As String = "Cow FAT Per"
    Public Const MixFATPer As String = "Mix FAT Per"
    Public Const CmtSetting As String = "Comment Setting"
    Public Const AreaWiseBilling As String = "Area Wise Billing"
    Public Const Hidedetaildate As String = "Hide Detail Date"
    Public Const DeleteTempData As String = "Delete Temp Data"
    Public Const DeleteTempDataFromTime As String = "Delete Temp Data From Time"
    Public Const DeleteTempDataToTime As String = "Delete Temp Data To Time"
    Public Const SkipDatePayPeriodMaster As String = "Skip Date Pay Period Master"


    Public Const ShowNewFormatofPDF As String = "Show New Format of PDF"
    Public Const PaymentProcessPrintInHindi As String = "Payment Process Print In Hindi"
    Public Const VendorBankAdviceForSWM As String = "Vendor Bank Advice For SWM"
    Public Const MilkSamplShowOddEvenTwoGrid As String = "Show Odd and Even Two Grid"
    Public Const OpenODDEvenForm As String = "Open Odd-Even Form"
    Public Const Open4AnalyzerForm As String = "Open 4 Milk Analyzer Form"
    Public Const IsApplyEMIOnAssetValue As String = "Is Apply EMI On Asset Value"
    Public Const AllowFutureDateTransaction As String = "AllowFutureDateTransaction"
    Public Const FindNRGP_Request As String = "Show_NRGP_RequestNo"
    Public Const AllowCSAPriceMasterPostedData As String = "Allow CSAPriceMaster Posted Data"
    Public Const AllowItemMasterPostedData As String = "Allow Item Master Posted Data"
    Public Const AllowMilkItemMasterPostedData As String = "Allow Milk Item Master Posted Data"
    Public Const AllowBulkProcItemPostedData As String = "Allow Bulk Proc Milk Item Posted Data"
    Public Const AllowPriceListMasterPostedData As String = "Allow Price List Item Posted Data"
    Public Const ItemCrateWtinKg As String = "Item Default Crate Wt.(Kg.)"
    Public Const ItemJaaliWtinKg As String = "Item Default Jaali Wt.(Kg.)"
    Public Const ItemBoxWtinKg As String = "Item Default Box Wt.(Kg.)"
    Public Const ItemCrateRate As String = "Item Default Crate Rate"
    Public Const ItemJaaliRate As String = "Item Default Jaali Rate"
    Public Const ItemBoxRate As String = "Item Default Box Rate"
    Public Const ItemCanRate As String = "Item Default Can Rate"
    Public Const CustomerMasterFinderOnLocationwiseARReceipt As String = "Customer master finder location-wise on AR Receipt"
    Public Const SameuserCanNotloginmultipletimes As String = "Same user cannot login multiple times"
    Public Const ShowCancelButtonPO As String = "Show cancel button on purchase order"
    Public Const ShowOptionforSelectingCapex As String = "Show option for selecting capex code/subcode on PO"
    Public Const AutoClosePO As String = "Auto close PO when all qty. received."
    Public Const POCancel As String = "PO Cancel"
    'Public Const CreateJVForAllCasesinRGP = "Crate JV for all cases in RGP"
    Public Const StoreRequisitionMandatoryforstorerequest = "Store Requisition mandatory for store request"
    Public Const MandatoryEmployeeOnVehicleMaster = "Make employee no mandatory"
    Public Const PlantDepotMappingMandatory = "Map location of plant with depot is mandatory"
    Public Const AllowThreeFormatByDefaultForPrint = "Allow printing 3 formats by default"
    Public Const MTCapacityRequired = "MT Capacity Required"
    Public Const AllowBackDateEntry As String = "Allow back date entry for given days"
    Public Const BackDateEntryPwd As String = "BackDateEntryPwd"
    Public Const RevisedBudget As String = "Revised Budget"
    Public Const DipMarkingMendatory As String = "Make dip marking mendatory."
    Public Const AllowDispatchChecklistOnProductDispatch As String = "Allow dispatch checklist on product dispatch"
    Public Const ShowIndentBasedOnCreatedUser As String = "Show indent based on created user"
    Public Const ShowSystemStockinOpenMCC As String = "Show system stock in open MCC shift"
    Public Const Tankerfromtankersalemasteringateentry As String = "Tanker from tanker sale master in gate entry"
    Public Const ApplyMultiChamberInBulkWeighmentEntry As String = "Apply multi-chamber in bulk weighment entry"
    Public Const DefaultItemUOMForBulkSale As String = "Default item uom for bulk sale"
    Public Const InsuranceNoAndSealNoInBulkDispatch As String = "Show option for entering Insurance No and Seal No"
    Public Const ValidateFatSNFOnJobMilkSRN As String = "Validate FAT KG & SNF KG on Job Milk SRN"
    Public Const CancelDocDueToSRNReturn As String = "Cancel document due to SRN Return"
    Public Const AmountInLacsOnMisSaleRegister As String = "Allow amount in lacs on MIS SALE REGISTER"
    Public Const ShortCloseItemWiseOnPO As String = "Allow short close item wise on PO"
    Public Const MakeClosingofPOreadonlyforuser As String = "Make closing of PO read only"
    Public Const MakeClosingofCustomerTenderreadonlyforuser As String = "Make closing of customer tender read only"
    Public Const MakeClosingofCustomerTenderOrderreadonlyforuser As String = "Make closing of customer tender order read only"
    Public Const AllowModificationOnApprovalByApprovalUser As String = "Allow Modification On Approval By Approval User"
    Public Const AllowAutoCalculateADDREMOVEQty As String = "Auto Calculate Qty of Add/Remove Item"
    Public Const FATDeductionPercent As String = "FAT Deduction Percent"
    Public Const SNFDeductionPercent As String = "SNF Deduction Percent"
    Public Const RejectionReturnPenaltyPerUnit As String = "Rejection Return Penalty Per Unit"
    Public Const RejectionDrainPenaltyPerUnit As String = "Rejection Drain Penalty Per Unit"
    Public Const RejectionCOBPenaltyPerUnit As String = "Rejection COB Penalty Per Unit"
    Public Const GraceTimeForTransporter As String = "Grace Time For Transporter"
    Public Const GraceTimeFromGateEntryToDocWeighing As String = "Grace Time From Gate Entry To Dock Weighing"
    ''===========CSA Sale Settings====================================================================
    Public Const ShowCSAReturnTypeOnScreen As String = "Show CSA Return Type on screen"
    Public Const ShowCSARequestScreen As String = "Enable CSA Request Instead of Booking"
    Public Const AllowSchemeOnCSADeliveryOrder As String = "Allow Scheme at CSA DO Entry"
    Public Const AllowOtherItemOnCSAPriceMaster As String = "Allow Other Items On CSA Price Master"
    Public Const AllowRoundOff_OnCSASalePatti As String = "Inv. Amount Round-off on All Sale Invoice"
    Public Const AllowRoundOffAmountOnDCSSale As String = "Allow Round Off Amount On DCSSale"
    Public Const FreightChargeOnCSASaleInvoice As String = "Comm./Freight itemwise on CSA Sale Invoice"
    Public Const AllowDisabledCommissionOnCSATransfer As String = "Commission disabled on CSA Transfer"
    Public Const DoReadonly_UnitRate_AtCSASale As String = "Allow Rate readonly on CSA Sale"
    Public Const Allow_SaleMfgACONCSAPatti As String = "Allow Sale mfg. A/c on CSA Sale Patti"
    Public Const AllowSchemeItemCondONSchemeMaster As String = "Allow Scheme type item on Scheme Master"
    Public Const ForUDLOnly As String = "CSA Sale changes For UDL only"
    Public Const CheckCreditLimitonCSADO As String = "Check Credit Limit on CSA DO"
    Public Const GrossWtFromItemMasterONCSATransfer As String = "Gross Wt. from item master on CSA Transfer"
    Public Const DisplayPriceCodeForBKN As String = "Display Price Code For BKN"
    Public Const EnableExciseONCSASalePatti As String = "Enable Excise entry on CSA Sale Patti"
    Public Const BatchSkipCSAReturn As String = "Batch Skip at CSA sale patti/Return"
    ''===========End Here====================================================================
    Public Const IsChamberWiseTanker As String = "Chamber wise Tanker"
    Public Const AllowLoginTypeCNFdistributerRetailer As String = "Allow Login Type CNF , Distributer, Retailer"
    Public Const AllowSchemeItemQty As String = "Allow Scheme Item in Materix Report"
    Public Const AllowDairyDeliveryOrderPrint As String = "Allow Print Button for Delivery Order "
    Public Const ShowSealNumberForTunkerOut As String = "Show Seal Number for Tunker Out"
    Public Const HideRateDispatchCentreCode As String = "Hide Rate and Dispatch Centre Code"
    Public Const AllowPromptPendingDocs As String = "Allow Prompt Pending Docs"
    Public Const AllowAutoGenerateDocNoInMaster As String = "Allow Auto Generate Doc No In Master Screen"
    Public Const ShowDocsStatusFilters As String = "Show Documents Declaration Status Filters"
    Public Const AutoDepartmentMendatroryFieldOnPurcahseCycle As String = "Allow Department Mandatory On Purchase Cycle"
    Public Const AllowVehicleGateOutValidationScrapSale As String = "Allow Vehicle Gate Out Validation For Scrap Sale"
    Public Const AllowVehicleGateOutValidationCSATransfer As String = "Allow Vehicle Gate Out Validation For CSA Transfer"
    Public Const AllowVehicleGateOutValidationSPSale As String = "Allow Vehicle Gate Out Validation For SP Sale"
    Public Const AllowVehicleGateOutValidationTransfer As String = "Allow Vehicle Gate Out Validation For Transfer"
    Public Const AllowWithoutUnitCostIssueReturnEntry As String = "Allow without amount save Issue/Return Entry"
    Public Const ZeroCostForReprocess As String = "Zero Cost For Reprocess"
    Public Const IsAutoReceiptPayment As String = "Auto Receipt Payment"
    Public Const TransferEntryOnInvCtrlAccount As String = "Transfer Entry On Inventory Control Account"
    Public Const AutoUpdateVLCUploaderCodeInVLCMaster As String = "AutoUpdateVLCUploaderCodeInVLCMaster"
    Public Const StandardInterfaceForMilkShiftEnd As String = "StandardInterfaceForMilkShiftEnd"
    Public Const ShiftEndAllowManualEntryOfDeduction As String = "Allow Manual Entry Of Deduction"
    Public Const PTMRatePerLtrKGOnStdQty As String = "Rate Ltr/KG On Std Qty"
    'for mobile app cash payment
    Public Const DefaultBank = "Default Bank for Cash Payment"
    Public Const DefaultLocation = "Default Location for Cash Payment"
    Public Const ShowParticluarColumnInSalesRegisterForGopalJee As String = "Show Column in sale register report for GopalJee"
    Public Const ShowPrintDiscountInDairyDispatchForGopaljee As String = "Show print discount in Dairy Dispatch"
    Public Const MilkReceiptRequiredApproval As String = "Milk Receipt Required Approval"
    Public Const LinkDepartmentBetweenIndentAndIssue As String = "Link Department Between Indent And Issue"
    Public Const CombineExportImportOnSchemeMaster As String = "Combined Export/Import on Scheme Master Dairy"
    Public Const OpenPOforRejectShortageQty As String = "Open PO for Reject/Shortage Qty"
    Public Const AutoSelectMCCRouteVLC As String = "Auto Select MCC Route VLC"
    Public Const PickServerDateWithNoChange As String = "Pick server Date With No Change"
    Public Const PickFinishedItemasBatchItem As String = "Finish Item as BatchItem default on Item Master"
    Public Const ToleranceFixFor_RM_OT_TRADE As String = "Tol.% mandatory for RM,Other,Trade on Item Master"
    Public Const ConsiderAdvancePayment As String = "Consider Advance Payment"
    Public Const PayableAmountZeroForMCCSale As String = "Payable Amount Zero For MCC Sale"
    Public Const Allow_AmountTruncate_BulkMilkSRN As String = "Allow truncate amount on Bulk Milk SRN"
    Public Const AutoPurchaseReturnFromIssueReturn As String = "Auto Purchase Return from Issue/Return screen."
    Public Const ShowAlternateVechileforFreshSale As String = "Gate pass with alternate vechile for fresh sale"
    Public Const CreateProvisionOfTransporterInDairyDispatch As String = "Create provision of transporter in Dairy Dispatch"
    Public Const IncludeRatePerHoursIn As String = "Include Rate Per Hours In"
    Public Const SinglePrintCopyDairyInvoice As String = "Print Out Single Copy for Dairy Invoice"
    Public Const GSTApplicable As String = "Allow GST Applicable"
    Public Const GSTApplicableDate As String = "Allow GST Applicable Date"
    Public Const AllowPanNoValidation As String = "Allow PAN No Validation"
    Public Const GSTActiveTaxesRatesGroup As String = "Show only Active Taxes/Rates/Groups for GST"
    Public Const AllowManualRejectionOfTanker As String = "Allow Manual Rejection Of Tanker"
    Public Const RunBulkProcOnAdjustedFATCLR As String = "Run Bulk Proc on adjusted FAT and CLR"
    Public Const BulkProcNetWeightCalculationWithVendorWeight As String = "Bulk Proc NetWeight Calculation by Vendor Weight"
    Public Const BulkProcPriceChartStandardRateWithZero As String = "Bulk Proc Price Chart standard rate with zero"
    Public Const RemoveForceAapprovalofBulkSRN As String = "Remove Force Approval of Bulk SRN"
    Public Const Allow_Plant_Depot_MCC_typeLocation As String = "Allow Plant Depot MCC type Location"
    Public Const ApplySaleRouteInDCSSale As String = "Apply Sale Route In DCS Sale"
    Public Const ValidateCustomerPANwithName As String = "Allow Validate Customer PAN with Name"
    Public Const ValidateTaxGroupForTransaction As String = "Allow Validate Tax Group Should Not Blank"
    Public Const AllowSeprateSchemeItemPrintDairySaleInvoice As String = "Allow Seprate Scheme Item Print DairySaleInvoice"
    Public Const EnableHirerachyCostCentre As String = "Enable Hirerachy Level Cost Centre"
    Public Const EnableStoreCostCentre As String = "Enable Store Cost Centre"
    Public Const EnableCostingMethod As String = "Enable Costing Method"
    Public Const CalculateItemCostonAvgForAssembly As String = "CalculateItemCostonAvgForAssembly"
    Public Const ShowAllCustomerOnMccMaterialSale As String = "Show All Customer On MCC Material Sale"
    Public Const ShowDefaultUser As String = "Show Default User"
    Public Const ShowVatSeriesNoSeprately As String = "Allow Tax Tracking to Show Vat series No Seperatly"
    Public Const AllowToGenerate_NEFTUPLOADER As String = "Allow Generate New NEFT UPLOADER File"
    Public Const DebitBankSelectWithNewFormateInNFTUploader As String = "Debit Bank Select With New Formate In NFT Uploader"
    Public Const AllowBulkPostingofAllDocuments As String = "Allow Bulk Posting of All Documents"
    Public Const AllowSameaAdditionalChargesMultiTime As String = "Allow Same Additaional Charges Multiple time"
    Public Const AllowToSaveAndUpdatePasswordBased As String = "Allow Masters To Save and Update Pasword Based"
    Public Const AllowMasterModificationWithSecurity As String = "Allow Master Modification With Security"
    Public Const ApplyRTGSAmtMoreThanGiven As String = "Apply RTGS Amount More Than Given"
    Public Const GenerateSecondryCode As String = "Excise Secondary Series on Transfer"
    Public Const POWeighmentManual As String = "Manual Weighment"
    Public Const AddTypeForUserMaster As String = "Add Type(Super User, Driver) in UserMaster"
    Public Const AddParavetEmployeeType As String = "Add Type Paravet in Employee Type"
    Public Const CalculateFIFOAndLIFOCosting As String = "Calculate FIFO And LIFO Costing"
    Public Const AllowDeductionPercentOnIncoming As String = "Allow Deduction(%) on Incoming Quality"
    Public Const AllowLoginType As String = "Allow POS Functionality in ERP"
    Public Const MilkProcurementUploader As String = "Milk Procurement Uploader"
    Public Const TankerDispatchBulkUploader As String = "Bulk Tanker Uploader"
    Public Const EmptyCanWeight As String = "Empty Can Weight"
    Public Const MinuteInLastVehicleForGateEntry As String = "Minute Last Vehicle For Gate Entry"
    Public Const MinuteGateEntryToGrossWeight As String = "Minute Gate Entry To Gross Weight"
    Public Const MinuteGrossWeightToTareWeight As String = "Minute Gross Weight To Tare Weight"
    Public Const NoOfDaysForMultiInceForSameVSPForSamePayCycle As String = "NoOfDaysForMultiInceForSameVSPForSamePayCycle"
    Public Const NoOfDaysForShowFOSSDocument As String = "NoOfDaysForShowFOSSDocument"
    Public Const PurchaseCounterOnTransactionType As String = "Purchase Counter On Transaction Type"
    Public Const BulkProcurementCounterOnEntryType As String = "Bulk Procurement Counter On Entry Type"
    Public Const StopForRepeatedFATSNF As String = "Stop Repeat FAT SNF"
    Public Const SampleFONTSize As String = "Font Size"
    Public Const SMSPrefix As String = "SMS Prefix"
    Public Const EmailPrefix As String = "Email Prefix"
    Public Const PickPendingMilkSRNinNextPaymentCycle As String = "Pick Pending Milk-SRN in Next Payment Cycle"
    Public Const TreatChequeClearDateAsRecoDate As String = "TreatChequeClearDateAsRecoDate"
    Public Const BookWreckageFromSublocationOrSection As String = "BookWreckageFromSublocationOrSection"
    Public Const StopVSPBillIfSomethingWrong As String = "Stop VSP Bill If Something Wrong"
    Public Const PDCSetting As String = "PDC Setting"
    Public Const AllowRoadPermitNo As String = "AllowRoadPermitNo"
    Public Const ShowMessgForTDS As String = "ShowMessgForTDS"
    Public Const IsShowTreeView As String = "IsShowTreeView"
    Public Const ShowVLCUploaderData As String = "Show VLC Uploader Data"
    Public Const FatSnfWhenMilktypeSelect As String = "Fat Snf persentage allow When Milk Type Select"
    Public Const DairyFreshTaxableandNonTaxable As String = "Taxable and Non-Taxable Item"
    Public Const SMSEMailPassword As String = "SMS EMail Password"
    Public Const CreateNewDocumentOnUploader As String = "Create New Document On Uploader"
    Public Const PopupJE As String = "Popup JE"
    Public Const ShowAliasNames As String = "ShowAliasNames"
    Public Const ShowFatAndSnfPercentageFields As String = "ShowFatNSNFPerc"
    Public Const VehicleFitnessAndInsuranceFields As String = "VehicleFitnessFields"
    Public Const DocumentCancel As String = "Document Cancelation"
    Public Const PICancelUserPwd As String = "PI Cancel"
    Public Const DocumentCancelReturn As String = "Document Cancelation Return"
    Public Const CSADocumentCancel As String = "CSA Transfer Cancelation"
    Public Const FixVSPEMP As String = "Fix VSP EMP"
    Public Const FatSNFStockControl As String = "FatSNFStockControl"
    Public Const CheckBalanceFromInvMoveSummry As String = "CheckBalanceFromInvMoveSummry"
    Public Const ItemwiseFatSNFStockControl As String = "ItemwiseFatSNFStockControl"
    Public Const SepratePriceChartForCowMilk As String = "Seprate Price Chart For Cow Milk"
    Public Const SepratePriceChartForCamelMilk As String = "Seprate Price Chart For Camel Milk"
    Public Const ApplyStdFATSNFRate As String = "Apply Standard FAT SNF Rate"
    Public Const AllowRoundInFixedAsset As String = "Allow Round In Fixed Asset"
    Public Const AllowDecimalInFixedAsset As String = "Allow Decimal In Fixed Asset"
    Public Const OpenPriceChartPlanningScreenOnTotalSolid As String = "Open Price Chart Planning on Total Solid"
    Public Const MaxFATSNFPerForRate As String = "Max FAT SNF Per For Rate "
    Public Const AllowZeroQtyFATSNFInOpenMCCShift As String = "Allow Zero Qty FAT SNF In Open MCC Shift"
    Public Const AllowZeroQtyFATSNFInCloseMCCShift As String = "Allow Zero Qty FAT SNF In Close MCC Shift"
    Public Const POLimit As String = "POLimit"
    Public Const RequiredPOLimit As String = "RequiredPOLimit"
    Public Const UnitCostIncreasePurchaseInvoice As String = "UnitCostIncreasePurchaseInvoice"
    Public Const PromptMsgForPendingDocIntervel As String = "Prompt Messg For Pending Doc Intervel"
    Public Const UDLPurchaseOrderthroughAP As String = "UDL Purchase Order through AP invoice"
    Public Const UpdateInventorySummaryTable As String = "UpdateInventorySummaryTable"
    Public Const CreateConsumeEntry As String = "Create Consume Entry"
    Public Const ShowOptionforSelectingCapexForFA As String = "Showoptionforselectingcapexcode/subcodeonFA"
    Public Const UDLCapexAcquisionEntry As String = "UDL Capex for Acquision Entry"
    Public Const UDLRGPWiseDocument As String = "UDL RGP Wise Document Created"
    Public Const AllowAssetItemOnMiscSale As String = "Allow Asset Item on Misc. Sale"
    Public Const TriggerOfGLEntryForWinTable As String = "Trigger Of GL Entry For Win Table"
    Public Const ShowRouteWiseAndVLCWiseReport As String = "ReportOfRouteAndVLCWise"
    Public Const UOMAtDiarySaleReturn As String = "Opt to change UOM At Diary SR on Actual UOM & Qty"
    Public Const PayableAmountZeroForFarmerPayment As String = "PayableAmountZeroForFarmerPayment"
    Public Const CheckDocAmountInAPInvoiceEntry As String = "Check Doc Amount For AP Invoice Entry"
    Public Const ApplyTSPriceAtBulkSale As String = "Apply TS Price At Bulk Sale"
    Public Const ShowPendingDocumentsListScreenOverDeclaredDocumentList As String = "Show Pending Documents Screen"
    Public Const MannualySetMPUploaderData As String = "MannualySetMPUploaderData"
    Public Const AllowSNFNotManditoryInBulkSale As String = "Allow SNF Not Manditory in Bulk Sale"
    Public Const VSPMPDiffrenceOnTSBasis As String = "VSP MP Diffrence On TS Basis"
    Public Const MilkProcuremntPickCLRInsteadOfSNF As String = "Milk Procuremnt Pick CLR Instead Of SNF"
    Public Const DBF As String = "DBF"
    Public Const chkGSTTaxGroupValidity As String = "check GST Tax Group Validity"
    Public Const ShowShipToPartyInDairyDispatch As String = "Show Ship To Party In Dairy Dispatch"
    Public Const BulkQCWithoutCLR As String = "Bulk Quality Check Without CLR"
    Public Const DOTaggingForDairySaleModule As String = "DO Tagging For Dairy Sale Module"
    Public Const AllowFractionInMCCTankerDispatchGrossQty As String = "Allow Fraction In MCC Tanker Dispatch Gross Qty"
    Public Const PurchaseModulePickFixTaxRate As String = "Purchase Module Pick Fix Tax Rate"
    Public Const TankerDispatchFinancialImpactInTransferIn As String = "Tanker Dispatch Financial Impact In Transfer In"
    Public Const ConvertQtyIntoKG = "Convert Qty into KG Bulk Sale Dispatch"
    Public Const GSTExemptedAmountForNonRegisteredVendor As String = "GST Exempted Amount For Non Registered Vendor"
    Public Const IncreaseCrateQtyOnFiftyPercent As String = "Increase Crate Qty On Fifty Percent"
    Public Const FATSNFDeductionMixMilkFATMinValue As String = "FAT SNF Deduction Mix Milk FAT Min Value"
    Public Const FATSNFDeductionMixMilkFATMaxValue As String = "FAT SNF Deduction Mix Milk FAT Max Value"
    Public Const FATSNFDeductionMixMilkSNFMinValue As String = "FAT SNF Deduction Mix Milk SNF Min Value"
    Public Const FATSNFDeductionMixMilkSNFMaxValue As String = "FAT SNF Deduction Mix Milk SNF Max Value"
    Public Const FATSNFDeductionMixMilkDeductionPer As String = "FAT SNF Deduction Mix Milk Deduction Per"
    Public Const VSPHoldPaymentNotCompanyBank As String = "VSP Hold Payment Not Company Bank"
    Public Const TDSRoundOffAmount As String = "TDS Round Off Amount"
    Public Const RoundOffPaiseAmount As String = "Round Off Paise Amount"
    Public Const RemoveSavingDocumentWhenPayableAmtZero As String = "Remove Saving Document When Payable Amt Zero"
    Public Const RecalculatePaymentProcessOnSave As String = "Recalculate Payment Process On Save"
    Public Const EnableInternalTransfer As String = "Enable Internal Transfer for UDL"
    Public Const JobWorkTransfer As String = "Job Work Transfer"
    Public Const FreightProvisionAccount As String = "Freight Provision Account"
    Public Const FreightProvisionAccountInward As String = "Freight Provision Account Inward"
    Public Const TreatUnregisteredVendorAsRegisteredVendor As String = "Treat Unregistered Vendor As Registered Vendor"
    Public Const RecreateConsumptionEntry As String = "RecreateConsumptionEntry"
    Public Const SaleRemarksForTruckSheetReport As String = "SaleRemarksForTruckSheetReport"
    Public Const BankRecoHidePWD As String = "Bank Reco Hide PWD"
    Public Const EnableItemGroupGLMapping As String = "Enable Item Group GL Mapping"
    Public Const EnableRackBin As String = "Enable Rack Bin Item"
    Public Const ChangeVehicleOnDairySaleBooking = "Change Vehicle On Dairy Sale Booking"
    Public Const VendorSetOffDayWise = "Vendor Set Off Day Wise"
    Public Const ReadOnlyTemplateFieldsOnAcqusition As String = "ReadOnlyTemplateFieldsOnAcqusition"
    Public Const IsAutoStartReading As String = "IsAutoStartReading"
    Public Const AddHighSecurityOnWeighingIntegratedScreen As String = "Add High Security On Weighing Integrated Screen"
    Public Const HighSecurityStableSeconds As String = "High Security Stable Seconds"
    Public Const HighSecurityWeightTolerance As String = "High Security Weight Tolerance"
    Public Const AllowManualvehicleOnDairyBooking As String = "AllowManualvehicleOnDairyBoking"
    Public Const FreeIndentQtyAfterPOClose As String = "Free Indent Qty After PO Close"
    Public Const ShowFATSNFinPaymentProcess As String = "Show FAT SNF in Payment Process"
    Public Const FileUpload As String = "File Upload"
    Public Const MaxRowsInCSVExport As String = "MaxRowsInCSVExport"
    Public Const MaxRowsInExcelExport As String = "MaxRowsInExcelExport"
    Public Const BigValidity As String = "Big Validity"
    Public Const AllowAssetBookChangeInTemplate As String = "AllowAssetBookChangeInTemplate"
    Public Const AllowSMSSendtoSalePerson As String = "Allow SMS Send to Sale Person"
    Public Const AllowSMSwhenCustomerCreditLimit As String = "SMS when Customer Credit limit reaches on DO."
    Public Const EnableScreenSelection As String = "Enable Screen Selection"
    Public Const SkipJobWorkSRNInPI As String = "Skip JobWork SRN in PI"
    Public Const ShowFatSnfAfterApproval As String = "Show Fat/Snf After Approval"
    Public Const ApplyTotalSolidPriceChart As String = "Apply Total Solid Price Chart"
    Public Const RequiredMgmtApprovalForRateIncrease As String = "Required Mgmt Approval For Rate Increase"
    Public Const AutoRoundOffSeprateAccountOnVendorTransaction As String = "Auto Round Off Seprate Account on Vendor Trans"
    Public Const TreatCRATEAsItems As String = "Treat CRATE as Item"
    Public Const TreatCANAsItems As String = "Treat CAN as Item"
    Public Const DoNotShowDairyTypeItems As String = "Do not show dairy type items"
    Public Const PasswordRules As String = "Password Rules"
    Public Const AlwaysVSPDefaulter As String = "Always VSP Defaulter"
    Public Const RejectedMilkSendToRejectLocation As String = "Rejected Milk Send To Reject Location"
    Public Const NoOfPreNxtDayToPickAvgFATSNF As String = "No Of Pre Nxt Day To Pick Avg FAT SNF %"
    Public Const AutoGeneratePrefix As String = "Auto Generate Prefix"
    Public Const SingleUserParticularDairyBookingEdit = "SingleUserCanEditOneDairyBookingDocumentAtaTime"
    Public Const DBUnlock As String = "DBUnlock"
    Public Const MaxReceiveSNFPer As String = "Max Receive SNF %"
    Public Const MailForAdvancePaymentTerm As String = "send mail for advance payment term"
    Public Const DonotCheckAvgQtyOnDairyBooking As String = "Check Average Qty On Dairy Booking"
    Public Const leakagededuction_freshsale As String = "Leakage Deduction Freshsale (%) "
    Public Const FirstGateOutProcessForMCCBulkProcument As String = "First Gate Out Process For MCC Bulk Procument"
    Public Const MCCBulkProcumentSecurityGateOut As String = "MCC Bulk Procument Security Gate Out"
    Public Const EnableDistributorSubsidy As String = "Enable Distributor Subsidy"
    Public Const DoNotAllowSavePOWhenQtyNRateZero As String = "Do Not Allow Save PO When Qty and Rate Zero"
    Public Const ActivateSFGProduction As String = "ActivateSFGProduction"
    Public Const ShowOnlyProductionItemInAdRemove As String = "ShowOnlyProductionItemInAdRemove"
    Public Const PeriodofSubsidyCreditNote As String = "Subsidy Credit Note Period"
    Public Const CHADetailsMandatoryOnExportSale As String = "CHA Details Mandatory On ExportSale"
    Public Const KnockOffFATSNFKG As String = "Knock Off FAT SNF KG"
    Public Const CreateLoadINSlipVehicleWise As String = "Create Load IN Slip Vehicle Wise"
    Public Const RouteCodeNotMandatoryOnLoadINSlip As String = "Route Code Not Mandatory On Loadin Slip"
    Public Const CrateReceiveddairyRouteWise As String = "Crate Received dairy Route Wise"
    Public Const RequiredFinalQCofstandardization As String = "Required Final QC of standardization"
    Public Const PickOnAccountPaymentForAdvanceKnockOff As String = "Pick OnAccount Payment For Advance Knock Off"
    Public Const PickProductCostFromItemUOMDetail As String = "Pick Product Cost From Item UOM Detail"
    Public Const AllowNegativeStockInDairyProduction As String = "Allow Negative Stock In Dairy Production" ''BHA/20/08/18-000463 by balwinder on 21/08/2018
    Public Const SkipLockLocation As String = "Skip Lock Location"
    Public Const ItemStructureMandatoryOnWeightConversion As String = "Item Structure Mandatory On Weight Conversion"
    Public Const VillageDiffrenceReportValueWithTwoDecimalPoint As String = "Village Diff Report Value With Two Decimal Point"
    Public Const RequiredFinalQCofProductionEntry As String = "Required Final QC of ProductionEntry" ''BHA/24/08/18-000479 by balwinder on 30/08/2018
    Public Const ShowAllPendingDOIrrespectiveOfDeliveryDate As String = "Show All Pendening DO Irrespective Of DeliveryDate" '' work to be done agaist ticket no.TEC/30/08/18-000317 on 31/08/2018
    Public Const PickFATSNFPerFromStock As String = "Pick FAT SNF % From Stock"
    Public Const ShowBulkDispatchQtyInLtr As String = "Show Bulk Dispatch Qty In Ltr"
    Public Const DoNotIncludeIncentiveInMilkPurchaseInvoice As String = "Do Not Include Incentive In Milk Purchase Invoice"
    Public Const DoNotConsiderTheFutureDateOfAdvancePayment As String = "Do Not Consider The Future Date Of Advance Payment"
    Public Const CreateEmpCodeAsPerEmployeeBasisType As String = "Create Emp Code As Per Employee Basis Type"
    Public Const FATSNFRateMandatory As String = "FAT SNF Rate Mandatory"
    Public Const ProductionOrStandAccordingToItemType As String = "Production Or Stand. According To Item Type"
    Public Const PwdOpenOnMainGLAccountAfterSave As String = "PwdOpenOnMainGLAccountAfterSave"
    Public Const UseProductFATSNFKgForEstimationCost As String = "Use Product FAT SNF Kg For Estimation Cost"
    Public Const ShowItemInCaseofNonInventory As String = "Show Item In Case of Non Inventory"
    Public Const ProductionFATRateTollerance As String = "Production FAT Rate Tollerance"
    Public Const ProductionSNFRateTollerance As String = "Production SNF Rate Tollerance"
    Public Const MaxFATPerLimit As String = "Max FAT Per"
    Public Const MaxSNFPerLimit As String = "Max SNF Per"
    Public Const MinFATPerLimit As String = "Min FAT Per"
    Public Const MinSNFPerLimit As String = "Min SNF Per"
    Public Const LockDate As String = "Lock Date"
    Public Const ApplyTransFATSNFRateForCalculateFATSNFRate As String = "Apply Trans FAT SNF Rate For Calculate FATSNF Rate"
    Public Const GrossWeightChangePWD As String = "Gross Weight Change PWD"
    Public Const MachineIntegrationInGeneralWeighment As String = "Machine Integrate In General Weighment"
    Public Const FillGeneralWeighmentDetailsByJobworkTypeGateInNo As String = "Fill General Weighment By Jobwork Type Gate In No"
    Public Const GateOutTankerNoAfterGeneralWeighment As String = "Gate Out Tanker No After General Weighment"
    Public Const NoOFCustomerForImportExport As String = "No OF Customer For Import Export"
    Public Const NoOFIncentiveForMPImportExport As String = "No OF Incentive For MP Import Export"
    Public Const NoOFSlabForImportExport As String = "No OF Slab For Import Export "
    Public Const NoOFItemStructureForImportExport As String = "No OF Item Structure For Import Export"
    Public Const CheckUnpostedPaymentProcess As String = "Check Unposted Payment Process"
    Public Const NoOFSavingCodeForImportExport As String = "No OF Saving Code For Import Export "
    Public Const NoOFCustomerForImportExportOnRouteMaster As String = "No OF Customer For Import Export On Route Master"
    Public Const CrateToLTR = "CRATE To LTR"
    Public Const CanToLTR = "CAN To LTR"
    Public Const ProdcutionDoNotCheckForwardDocuments = "Production Do Not Check Forward Documents"
    Public Const PurchaseDoNotCheckForwardDocuments = "Purchase Do Not Check Forward Documents"
    Public Const DoNotStopOnItemBalanceExceptionStoreAdjustment = "Do Not Stop on Balance Exception Store Adjustment"
    Public Const ProductionIssueQtyTollerance = "Production Issue Qty Tollerance"
    Public Const NopreviousDaysInSaleVSReceipt = "No of previous days in Sale vs Receipt"
    Public Const SelectGLInProftAndLossPerforma = "Select GL In Proft And Loss Performa"
    Public Const SelectGLInBalanceSheetPerforma = "Select GL In Balance Sheet Performa"
    Public Const BalanceSheetPerformaWithFormula = "Balance Sheet Performa With Formula"
    Public Const SelectGLInCashFlowPerforma = "Select GL In Cash Flow Performa"
    Public Const BulkProcurementApplyTotalSoidRate = "Bulk Procurement Apply Total Soid Rate"
    Public Const CalculateLtrQtyFromKGQtyByCLR = "Calculate Ltr Qty From KG Qty By CLR"
    Public Const CalculateProvisionOnGateePass = "Calculate Provision On Gatee Pass"
    Public Const CalculateSNFFromCLRForMCCMilk = "Calculate SNF From CLR For MCC Milk"
    Public Const TagExemptedtaxgroupincaseofBankChargesinPaymentEntry = "Exempted group incase of Bnk Charges in Paymt Enty"
    Public Const SNFFromCLRAndCorrectionFactorInJWIEst = "SNF From CLR And Correction Factor in JWI Est"
    Public Const AutoCalculateProduceQty = "Auto Calculate Produce Qty"
    Public Const SyncedMccToServerStartDateForEmailSms = "Synced Mcc To Server Start Date For Email Sms"
    Public Const CrateReceiveddairyCustomerWise As String = "Crate Received dairy Customer Wise"
    Public Const MaintainLogForImporperSample = "Maintain Log For Imporper Sample"
    Public Const PFCalculationOnFormulaHead As String = "PF Calculation On Head Value"
    Public Const MODValueForFAT As String = "MOD Value For FAT"
    Public Const AutoFillSameLocationInGrid As String = "Auto Fill Same Location In Grid"
    Public Const ApplyCalculateWeightInLtr = "Apply Calculate Weight In Ltr"
    Public Const ApplyTransportChargeAddInActualAmount = "Apply Transport Charge Add In Actual Amount"
    Public Const CheckUniqueDocumentCode = "Check Unique Document Code"
    Public Const JWIRateofFGasPerRM = "JWI Rate of FG as per RM"
    Public Const TagMultipleRouteWithCustomer = "Tag Multiple Route With Customer"
    Public Const ApplyNotShowJobWorkTypeTanker = "Apply Not Show JobWork Type Tanker"
    Public Const ExportToDefineLocation As String = "Export To Define Location"
    Public Const ImportMultipleAssetAssembled As String = "ImportMultipleAssetAssembled"
    Public Const ShowEmpCurrentSalaryOnEmployeeSatatusReport As String = "Show Emp Current Salary On Employee Satatus Report"
    Public Const AllowOneFormatByDefaultForPrint = "Allow printing One formats by default"
    Public Const CreateProvisionOnOpeningAndClosingKM = "Create Provision On Opening And Closing KM"
    Public Const FromLocationStockNotCheckConsumptionLocation = "From Location Stock not check Consumption Location"
    Public Const MilkIncetiveByMilkSRN = "Milk Incetive By Milk SRN"
    Public Const AllowShowCoumnInCrateReceivedDairy = "Show Coumn In Crate Received Dairy"
    Public Const CalculateLeakageAmount As String = "Calculate Leakage Amount"
    Public Const ShowFATSNFPerOnBulkProcInGateIN As String = "Show Fat% and SNF% On Bulk Proc In Gate IN Screen"
    Public Const AllowItemCostMandatoryForStockingUnit As String = "Allow Item Cost Mandatory for Stocking Unit"
    Public Const UseControlMForHelp As String = "Use Control M For Help"
    Public Const DairyBookingTolleranceQty As String = "Dairy Booking Tollerance Qty"
    Public Const VehicleCodeNotMandatoryOnLoadINSlip As String = "Vehicle Code Not Mandatory On LoadIN Slip"
    Public Const ApplyLatestPriceChartWhilecreatingNewVSP As String = "Apply Latest Price Chart While creating New VSP"
    Public Const AllowToPrintInvoiceAfterPosting As String = "Allow To Print Invoice After Posting"
    Public Const EnableBankFromMaster As String = "Enable Bank From Master"
    Public Const Autoadjustdebitcreditnote As String = "Auto Adjust Debit Credit Note"
    Public Const UpdateItemMasterWithoutTransactionValidation As String = "Update Item Master Without Transaction Validation"
    Public Const AddItemAliasInSMS As String = "Add Item Alias In SMS"
    Public Const ItemWiseQualityCheckInGeneralPurchase As String = "Item Wise Quality Check In General Purchase"
    Public Const SendInternalSMSEmailinPurchaseModule As String = "Send Internal SMS/Email in Purchase Module"
    Public Const AllowtoEnterNetWeightManuallyinPOWeighmentScreen As String = "AllowtoEnterNetWeightManuallyinPOWeighmentScreen"
    Public Const ShowNotificationInMDI As String = "Show Notification In MDI"
    Public Const TIPRateMix As String = "TIP Rate Mix"
    Public Const TIPRateCow As String = "TIP Rate Cow"
    Public Const TIPRateBuffalo As String = "TIP Rate Buffalo"
    Public Const DefaultCustomerGroupCodeforVSP As String = "Default Customer Group Code for VSP"
    Public Const DefaultVendorGroupCodeforVSP As String = "Default Vendor Group Code for VSP"
    Public Const AllowSameTankerNoforPrimarySecondaryTransporter As String = "AllowSameTankerNoforPrimarySecondaryTransporter"
    Public Const AllocateToMandatoryonGateOut As String = "AllocateToMandatoryonGateOut"
    Public Const PrintTruckSheetAfterGenerate As String = "PrintTruckSheetAfterGenerate"
    Public Const AskForPwdForOutAdjustmentOnPost As String = "AskForPwdForOutAdjustmentOnPost"
    Public Const checkStockOfItemTillTransactionDateOnly As String = "check stock of item till transaction date only"
    Public Const UseProductionPlaningDateForWholeProductionCycle As String = "UseProductionPlaningDateForWholeProductionCycle"
    Public Const allowMilkJWOutowordWithAvgFatSNFRate As String = "allowMilkJWOutowordWithAvgFatSNFRate"
    Public Const allowMilkJWOutowordWithAvgFatSNFPerAtInventory As String = "allowMilkJWOutowordWithAvgFatSNFPerAtInventory"
    Public Const CreateProvisionOfTankerDispatchWithClosingKM As String = "CreateProvisionOfTankerDispatchWithClosingKM"
    Public Const MaterialSaleInvoiceEnablePrintOnPost = "MaterialSaleInvoiceEnablePrintOnPost"
    Public Const DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn = "DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn"
    Public Const UpdateItemMasterConversationWithoutValidation = "UpdateItemMasterConversationWithoutValidation"
    Public Const TollTaxMaster = "Toll Tax Master"
    Public Const InDocMandatoryOnInternalTransfer = "InDocMandatoryOnInternalTransfer"
    Public Const AllowTransferInAfterGatePassOnly = "Allow Transfer In After Gate Pass Only"
    Public Const EnableItemShortDescriptionInBooking = "Enable Item Short Description In Booking"
    Public Const AllowDuplicateItemShortDescriptionInItemMaster = "AllowDuplicateItemShortDescriptionInItemMaster"
    Public Const DeleteMccMilkShiftPassword As String = "Delete Mcc Milk Shift Password"
    Public Const DoNotCreatePaymentWhileSalaryGeneration As String = "DoNotCreatePaymentWhileSalaryGeneration"
    Public Const MccPlantSelectionOptionInMccTankerGateOut As String = "MccPlantSelectionOptionInMccTankerGateOut"
    Public Const EnableTankerNoInMccTankerDispWithMccTankerGateOut As String = "EnableTankerNoInMccTankerDispWithMccTankerGateOut"
    Public Const AllowToEnterSnfAtPlantInMccTankerDispatch As String = "AllowToEnterSnfAtPlantInMccTankerDispatch"
    Public Const DateOfEInvoiceImplementation As String = "DateOfEInvoiceImplementation"
    Public Const BennyImportAutoCreateMP As String = "Benny Import Auto Create MP"
    Public Const BennyImportPickRateFromPrice As String = "Benny Import Pick Rate From Price"
    Public Const AllowOnlyOneIssueAgainstStoreRequisition As String = "AllowOnlyOneIssueAgainstStoreRequisition"
    Public Const UseVLCUploaderCodeMPUploaderCodeInMCCProcurement As String = "UseVLCUploaderCodeMPUploaderCodeInMCCProcurement"
    Public Const ShowNotificationWithoutSMSAPP As String = "ShowNotificationWithoutSMSAPP"
    Public Const SetNotificationRefreshTimeInMinutes As String = "SetNotificationRefreshTimeInMinutes"
    Public Const GenerateEWayBillWithEInvoice As String = "GenerateEWayBillWithEInvoice"
    Public Const VillageDataReverse As String = "VillageDataReverse"
    Public Const ShowTankerWithoutCheckingAnyValidation As String = "ShowTankerWithoutCheckingAnyValidation"
    'Public Const AllocatedTankerGateOut As String = "AllocatedTankerGateOut"
    Public Const AllowToCreateNoOfBookingPerDay As String = "AllowToCreateNoOfBookingPerDay"
    Public Const EnterWeightManuallyOnWeighmentInGrid As String = "EnterWeightManuallyOnWeighmentInGrid"
    Public Const AllowSameChequeNoForMultiplePaymentEntry As String = "AllowSameChequeNoForMultiplePaymentEntry"
    Public Const AllowGenerateReferenceNoForBulkGateEntry As String = "Allow Generate Reference No For Bulk Gate Entry"
    Public Const CreateNeftuploaderPlantWise As String = "Create Neftuploader Plant Wise"
    Public Const AllowBankTransferAgainstMilkBill As String = "Allow Bank Transfer Against Milk Bill"
    Public Const ShowTCSAmountOnBookingForOtherCustomer As String = "ShowTCSAmountOnBookingForOtherCustomer"
    Public Const ShowCheckExcludeProvisionBank As String = "Show Check Exclude Provision Bank"
    Public Const ChangeLeaveDescriptionOnSalarySlip As String = "ChangeLeaveDescriptionOnSalarySlip"
    Public Const ApplyDepartmentWiseDataVisibleInDepartmentIndent As String = "ApplyDepartmentWiseDataVisibleInDepartmentIndent"
    Public Const UpdateMapPayHeadsToSalaStructurePassword As String = "UpdateMapPayHeadsToSalaStructurePassword"
    Public Const ApplyFinancialCostCenter As String = "Apply Financial Cost Center"
    Public Const SalarySlipLeaveStatusOnTheBasisOfCalendarYear As String = "SalarySlipLeaveStatusOnTheBasisOfCalendarYear"
    Public Const ApplyCalculationOnRouteLenth As String = "ApplyCalculationOnRouteLenth"
    Public Const EnableExportExcelOnIncentiveEntry As String = "EnableExportExcelOnIncentiveEntry"
    Public Const AllowBankSectionEnableOnMCCMaster As String = "AllowBankSectionEnableOnMCCMaster"
    Public Const DateOfDynamicQRCodeForB2CInvoiceImplementation As String = "DateOfDynamicQRCodeForB2CInvoiceImplementation"
    Public Const ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster As String = "ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster"
    Public Const AllowRoundDownAmtForMCCDateWiseChilling As String = "AllowRoundDownAmtForMCCDateWiseChilling"
    Public Const CtreateJEOfVspAssetIssueAndReturn As String = "CtreateJEOfVspAssetIssueAndReturn"
    Public Const createDebitNoteOnAssetIssue As String = "createDebitNoteOnAssetIssue"
    Public Const PickAvgCostonAssetissue As String = "PickAvgCostonAssetissue"
    Public Const ApplyTCSAmtOnAbstractReportDotMatrix As String = "ApplyTCSAmtOnAbstractReportDotMatrix"
    Public Const NotAllowDuplicatePANOnPrimaryTransporter As String = "NotAllowDuplicatePANOnPrimaryTransporter"
    Public Const ApplyIncludeTCSAmountInRouteTotalOnTruckSheet As String = "ApplyIncludeTCSAmountInRouteTotalOnTruckSheet"
    Public Const ShowEarlyRouteOnTruckSheet As String = "ShowEarlyRouteOnTruckSheet"
    Public Const AllowOutEntryOnCrateReceivedDairyForAdjustment As String = "AllowOutEntryOnCrateReceivedDairyForAdjustment"
    Public Const BulkMilkFATSNFKGDecimalPlaces As String = "Bulk Milk FAT SNF KG Decimal Places"
    Public Const BulkMilkFATSNFAmtDecimalPlaces As String = "Bulk Milk FAT SNF Amt Decimal Places"
    Public Const BulkMilkConsiderAllParametersForIncetive As String = "Bulk Milk Consider All Parameters For Incetive"
    Public Const MCCMaterialSaleFarmerReverse As String = "MCCMaterialSaleFarmerReverse"
    Public Const LocalSaleAllowedPer As String = "Local Sale Allowed Per"
    Public Const LocalSaleAllowedRate As String = "Local Sale Allowed Rate"
    Public Const MultiDairyGatePassReversePWD As String = "MultiDairyGatePassReversePWD"
    Public Const CostCenterAndHirerachyCodeUpdateAfterPost As String = "CostCenter And Hirerachy Code Update After Post"
    Public Const ShowStatusItemWiseInPendingRequisitionRpt As String = "Show Status Item Wise In Pending Requisition Rpt"
    Public Const DoNotCheckAnyValidationOnVendorInactive As String = "Do Not Check Any Validation On Vendor Inactive"
    Public Const SupportHindiFont As String = "Support Hindi Font"
    Public Const ConvertHindiFontLastDateTime As String = "Convert Hindi Font Last Date Time"
    Public Const ApplyMilkPouchPrint As String = "ApplyMilkPouchPrint"
    Public Const UserWiseRouteMapping As String = "UserWiseRouteMapping"
    Public Const CreateMCCTankerGateOutBasedOnBulkRouteMaster As String = "CreateMCCTankerGateOutBasedOnBulkRouteMaster"
    Public Const ApplyAutoLoginIdCreateOfCustomer As String = "ApplyAutoLoginIdCreateOfCustomer"
    Public Const DefaultLocationOfUserMaster As String = "DefaultLocationOfUserMaster"
    Public Const CreateGatePassFromDemand As String = "CreateGatePassFromDemand"
    Public Const PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister As String = "PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister"
    Public Const AadharNoMandatoryOnEmpMaster As String = "Aadhar No Mandatory On Emp Master"
    Public Const SuperUserCustomer As String = "SuperUserCustomer"
    Public Const UploadMultipleMasterPwd As String = "UploadMultipleMasterPwd"
    Public Const ApplyDefaultsInMaster As String = "ApplyDefaultsInMaster"
    Public Const IncentiveAccNoMandatoryInMPMaster As String = "IncentiveAccNoMandatoryInMPMaster"
    Public Const Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster As String = "Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster"
    Public Const MultipleFinderFillAuto As String = "MultipleFinderFillAuto"
    Public Const MandatoryLineNoMaxMinQtyForProductionPlan As String = "MandatoryLineNoMaxMinQtyForProductionPlan"
    Public Const RunProductionBaseOnPercentage As String = "RunProductionBaseOnPercentage"
    Public Const RCDFCFP As String = "RCDFCFP"
    Public Const ApplyLocationFilterBasedOnPermission As String = "ApplyLocationFilterBasedOnPermission"
    Public Const MandatoryPDFFileMilkPricePlan As String = "Mandatory PDF File In Milk Price Plan"
    Public Const AutoClosePOBasedOnSRNQtyWithTolerance As String = "AutoClosePOBasedOnSRNQtyWithTolerance"
    Public Const ApplyMilkTypeBuffaloCowOnPrint As String = "ApplyMilkTypeBuffaloCowOnPrint"
    Public Const ApplyZoneWiseVSP As String = "ApplyZoneWiseVSP"
    Public Const ApplyStandardProductionVariance As String = "ApplyStandardProductionVariance"
    Public Const ItemCostTolerancePercentage = "ItemCostTolerancePercentage"
    Public Const HeadLoadDescriptionInPaymentProcessPrint = "HeadLoadDescriptionInPaymentProcessPrint"
    Public Const PrefixForUserMaster = "Prefix For User Master"
    Public Const ApplyDemandApproval = "ApplyDemandApproval"
    Public Const ApplyDemandAll = "ApplyDemandAll"
    Public Const ApplyDemandCustomerWise = "ApplyDemandCustomerWise"
    Public Const CheckCreditLimit = "CheckCreditLimit"
    Public Const ApplyTolerance = "ApplyTolerance"
    Public Const ApplyOrderByNumeric = "ApplyOrderByNumeric"
    Public Const SetShiftTimeOut = "SetShiftTimeOut"
    Public Const AllowToSelectShift = "Allow To Select Shift"
    Public Const ApplyRoundOffZero = "ApplyRoundOffZero"
    Public Const EnableLocation = "EnableLocation"
    Public Const EnableResetDemand = "EnableResetDemand"
    Public Const IsLoadingSlipMandatory = "Is Loading Slip Mandatory"
    Public Const CrateCapacityCheck = "Crate Capacity Check"
    Public Const PickAllBMC = "VSP Milk Not Sold"
    Public Const ApplyMonthlyYearly = "Apply Monthly Yearly"
    Public Const SeparateDemandMilkandProduct = "Separate Demand Milk & Product"
    Public Const ShowDemandDoc = "Show Demand Document In Customer Booking"
    Public Const ApplyCommission = "Apply Distributor/Transpoter Commission"
    Public Const PricePlanRoundOffTruncate = "Price Plan RoundOff/Truncate"
    Public Const AllowIncreaseDispatchQty = "Allow Increase Dispatch Qty"
    Public Const FORPRICE = "FOR Price"
    Public Const AllowGatePassDemandTripWise = "Allow Gatepass Demand Trip Wise"
    Public Const ConvertPouchtoCrate = "Convert 1 Pouch to Crate"
    Public Const DontCreateForPouch = "do not calculate Crate for pouch"
    Public Const ChangeInventroyMovemnet = "Inventroy Movement from Dispath/Gatepass"
    Public Const ApplyCommissionRateWithTax = "Apply Commission Rate With TAX"
    Public Const DisableRouteandVehicle = "Disable Route and Vehicle No."
    Public Const UpdateDemandBeforePost = "Update Demand Before Post"
    'Public Const SuperuserViewOnly = "Super User View Only"
    'Public Const DistributorViewOnly = "Distributor View Only"
    Public Const ApplyPricePlanDigitRule = "Apply Price Plan Digit Rule"
    Public Const PricePlanRoundOffDigit = "Enter RoundOff Digit"
    Public Const DemandIncreaseDecreaseQty = "Demand Increase Decrease Qty"
    Public Const DemandIncreaseDecreasePer = "Demand Increase Decrease Percentage"
    Public Const CreateAutoMCCPriceChat = "Create Auto MCC Price Chat"
    Public Const VehicleNofromDispatch = "VehicleNo From Dispatch"
    Public Const DispatchPriceCodeForCreditCustomer = "Dispatch PriceCode For Credit Customer"
    Public Const ApplyCommissionTPTForCreditCustomer = "Apply Commission/TPT For Credit Customer"
    Public Const ConvertPouchtoCrateonDispatch = "Convert Pouch to Crate on Dispatch"
    Public Const PickRatefromMaster = "Pick Rate from Master"
    Public Const EnableProductSaleForJPR = "Enable Product Sale For JPR"
    Public Const EnableVehicleType = "Enable Vehicle Type"
    Public Const PickDataFromRetestingTable = "Pick Data From Retesting Table"
    Public Const ShowAllCustomerItemWiseSaleReportOptions = "Show All Customer Item Wise Sale Report Options"
    Public Const ApplyTPT = "Apply Transpoter Commission"
    Public Const AllowMultipleUOMForProduct = "Allow Multiple UOM For Product"
    Public Const AllowManualCrateForDispatch = "Allow Manual Crate For Dispatch"
    Public Const SetDefaultShiftTime = "Set Default Shift Time"
    Public Const ApplyManualScheme = "Apply Manual Scheme"
    Public Const AllowResetCustomerDemandOnRouteChange = "Allow Reset Customer Demand On Route Change"
    Public Const AutoSchemeOnTotalDispatchQty = "Auto Scheme On Total Dispatch Qty"
    Public Const ApplyDefaultTCSIsChecked = "Apply Default TCS IsChecked"
    Public Const ApplyDefaultTPTIsChecked = "Apply Default TPT IsChecked"
    Public Const ApplyBoothRouteMapping = "Apply Booth Route Mapping"
    Public Const APIMilkSample = "API Milk Sample"
    Public Const ApplyDepartmentRoute = "Apply Department Route"
    Public Const QuantityTolerance = "Quantity Tolerance"
    Public Const AmountTolerance = "Amount Tolerance"
    Public Const ApplyItemCapacityLimit = "Apply Item Capacity Limit"
    Public Const SeprateMorningEveningSequence = "Seprate Morning Evening Sequence"
    Public Const FinancialImpactForTPT = "Financial Impact For TPT"
    Public Const FinancialImpactForDistributor = "Financial Impact For Distributor"
    Public Const FinancialImpactForSecurity = "Financial Impact For Security"
    Public Const ShowAndSaveCrystalReportActionType = "Show And Save Crystal Report Action Type"
    Public Const QuickDemandUOMCrate = "Quick Demand UOM Crate"
    Public Const QuickDemandUOMPouch = "Quick Demand UOM Pouch"
    Public Const QuickDemandUOMLtr = "Quick Demand UOM Ltr"
    Public Const ApplyBoothWiseScheme = "Apply Booth Wise Scheme"
    Public Const OfflineDCSMilkCollectionExeVersion = "Offline DCS Milk Collection Exe Version"
    Public Const ApplyMonthEndDispatch = "Apply Month End Dispatch"
    Public Const ApplyItemUOMOnDemand = "Apply Item UOM On Demand"
    Public Const MergeTCAmtofCreditCust = "Merge Transpoter Commission Amt of CreditCust"
    Public Const ConvertIntoBulkUOM = "Convert into Bulk UOM"
    Public Const AddDistributorCommissioninItemPrice = "Add Distributor Commission In Item Price"
    Public Const SubstractDistributorCommissioninItemPrice = "Substract Distributor Commission In Item Price"
    Public Const ConvertTOBillingUOM = "Convert into Billing UOM"
    Public Const DifferentCrateTypeForFGItem = "Different Crate Type For FG Item"
    Public Const StopMultipleLoginInAPP = "Stop Multiple Login In APP"
    Public Const DispatchCommissionDecimalPlaces = "Dispatch Commission Decimal Places"
    Public Const ApplyPricePlanOnDocumentDate = "Apply Price Plan On Document Date"
    Public Const DisplayTransporterAddress = "Display Transporter Address"
    Public Const AllowRouteWiseDemandEntryInDecimal = "Allow Route Wise Demand Entry In Decimal"
    Public Const AllowToCheckEntryUOMForProduct = "Allow To Check Entry UOM For Product"
    Public Const AllowToCheckEntryUOMForGNG = "Allow To Check Entry UOM For GNG"
    Public Const AllowToCheckZeroQtyonDispatch = "Allow To Check Zero Qty on Dispatch"
    Public Const DefaultEnableEWayBill = "Default Enable EWay Bill"
    Public Const CreateAutoGatePass = "Create Auto Gate Pass"
    Public Const ApplyEWBThresholdLimit = "Apply EWB Threshold Limit"
    Public Const DefaultEnableNoTransporter = "Default Enable No Transporter"
    Public Const ManualBatchOnCustomerBooking = "Manual Batch On Customer Booking"
    Public Const GenerateCustomerWiseGatePass = "Generate Customer Wise GatePass"
    Public Const DeductTPTFromDocAmt = "Deduct TPT From Document Amt"
    Public Const CheckCustomeroutStandingAmt = "Check Customer OutStanding Amt"
    Public Const ServerDateTimeForTaxableInvoice = "Server DateTime For Taxable Invoice"
End Class
Public Class clsFixedParameterCode
    Public Const SkadaDefaultSiloLocation As String = "Default Silo Location"

    Public Const SkadaLastDocument As String = "Last Document"
    Public Const SkadaDefaultRoute As String = "Default Route"
    Public Const SkipDebitNoteforRejectInBMCTruckSheet As String = "Skip Debit Note for Reject In BMC Truck Sheet"
    Public Const GatepassForTaxableandNonTaxableItems As String = "Gatepass For Taxable and NonTaxable Items"
    Public Const PickDataFromMasterOrTransaction As String = "Pick Data From Master Or Transaction"
    Public Const ExportTruckSheet As String = "Export Truck Sheet"
    Public Const ThirtPartyFarmerCollectionIntegration As String = "Thirt Party Farmer Collection Integration"
    Public Const SuspenceAdjusmtmentFATSNFTolerancePercentage = "Suspence Adjusmtment FAT SNF Tolerance %"
    Public Const TransactionModify As String = "Transaction Modify"
    Public Const TransactionValidity As String = "Transaction Validity"
    Public Const AutoGenerateFarmerPriceFromDCSPrice As String = "Auto Generate Farmer Price From DCS Price"
    Public Const AutoGenrateBatchInventoryOnPost As String = "Auto Genrate Batch Inventory On Post"
    Public Const PickcompanyBankCodeFromPaymentProcess As String = "Pick company BankCode From Payment Process"
    Public Const RefreshTimeForUnionWiseMilkCollectionInms As String = "Refresh Time For Union Wise MilkCollection In MS"
    Public Const isThereOnlyOneRowOfOwnDCS As String = "is There Only One Row Of Own DCS"
    Public Const StopYesBankVirtualAccount As String = "Stop Yes Bank virtual Account"
    Public Const MaxFATPerLimitforReport As String = "Max FAT Per Limit for Report"
    Public Const MaxSNFPerLimitforReport As String = "Max SNF Per Limit for Report"
    Public Const PartialSRNAmount As String = "Partial SRN Amount"
    Public Const PrintOnlyPostedDocument As String = "Print Only Posted Document"
    Public Const HideShiftFilterOnDemandStatus As String = "Hide Shift Filter On Demand Status"
    Public Const DefaultFilesPath As String = "Default File Path"
    Public Const B2CDocumentAmountRangeOtherState As String = "B2C Document Amount Range Other State"
    Public Const B2CDocumentAmountRange As String = "B2C Document Amount Range"
    Public Const TuneTime As String = "Tune Time"

    Public Const BMCStartTime As String = "BMC Start Time"
    Public Const BMCEndHours As String = "BMC End hours"
    Public Const DCSEndHours As String = "DCS End hours"

    Public Const ApplySendApprovalSetting As String = "Apply Send Approval Setting"
    Public Const ViewDCSMilkPurchaseRegister As String = "View DCS Milk Purchase Register"
    Public Const NoOfDCSToLoadDeductionData As String = "No Of DCS To Load Deduction Data"
    Public Const DCSWiseFilterEnableOnSavingCheck As String = "DCS Wise Filter Enable On Saving Check"
    Public Const PenaltyAfterDays As String = "Define Penalty Days to Apply"
    Public Const PenaltyCost As String = "Define Penalty Cost Value for Per Unit"
    'Public Const SecretorydcsMaster As String = "Secretory/DCS Master"
    Public Const CalculateProRateAuto As String = "Calculate Pro Rate Auto"
    Public Const VendorMaster As String = "Vendor Master"
    Public Const VendorIntegratedMaster As String = "Vendor Integarted Master Password"

    Public Const MaterialSales As String = "Material Sales"
    Public Const UserMaster As String = "User Master"
    Public Const ItemMaster As String = "Item Master"
    Public Const DemandUnpost As String = "Demand Unpost"
    Public Const DispatchCancel As String = "Dispatch Cancel"
    Public Const BookingCancel As String = "Booking Cancel"
    Public Const EWayBillUpdate As String = "Update E-Way Bill"
    Public Const GatePassCancel As String = "GatePass Cancel"
    Public Const CustomerMaster As String = "Customer Master"
    Public Const FarmerMaster As String = "Customer Master"
    Public Const ConvertVlcCodeUploaderToInt As String = "Convert VlcCodeUploader To Int"
    Public Const PickOnlyOWNBMCDCS As String = "Pick Only OWN BMC DCS"
    Public Const PrintPouchCrateQtyOnPrint As String = "Print PouchCrate Qty On Print"
    Public Const PickHeadLoadRateFromSecretaryMaster As String = "Pick HeadLoadRate From Secretary Master"
    Public Const RCDFRateControl As String = "RCDF Rate Control"
    Public Const ImportExportDrive As String = "Import/Export-Drive"

    Public Const EmployeePFRetirementAge As String = "Employee PF Retirement Age"
    Public Const EmployeeRetirementAge As String = "Employee Retirement Age"
    Public Const NoOfRecordForLotNo As String = "No Of Record For Lot No"
    Public Const BankAdviseRequired As String = "Bank Advice Start Date"
    Public Const SeprateBankForDCSBankAdvice As String = "Seprate Bank For DCS Bank Advice"
    Public Const ApplyPartialBankAdvise As String = "Apply Partial Bank Advise"
    Public Const DBTMilkQtyCapping As String = "DBT Milk Qty Capping"
    Public Const DBTByDailyFarmerMilkCollection As String = "DBT By Daily Farmer Milk Collection"
    Public Const ApplyMergeForDCSMultipleDays As String = "Apply Merge For DCS Multiple Days"
    Public Const CalculateFATSNFLossCycleWise As String = "Calculate FAT SNF Loss Cycle Wise"
    Public Const ApplyRange As String = "Apply Range"
    Public Const RangeNotApplicable As String = "RangeNotApplicable"
    Public Const RangePO As String = "Range PO"
    Public Const RangeRAL As String = "Range RAL"
    Public Const RefreshDBTReco As String = "Refresh DBT Reco"
    Public Const DistributorWiseBilling As String = "Distributor Wise Billing"
    Public Const ApplyBoothSecurity As String = "Apply Booth Security"
    Public Const SetDCSAvgFATSNF As String = "Set DCS Avg FAT SNF of ALL Samples And Merge"
    Public Const BackDays As String = "Back Days"
    Public Const MaximumBackDays As String = "Maximum Back Days"
    Public Const ViewItemImage As String = "View Item Image"
    Public Const RoundOffBankAdvice As String = "RoundOff Bank Advice"
    Public Const NewDCSScreen As String = "New DCS Screen"
    Public Const SameBankCodeForCurrentAndSaving As String = "Same Bank Code For Current And Saving"
    Public Const MinimumQtyForHeadLoad As String = "Minimum Qty For Head Load"
    Public Const ApplyHeadLoadDayAndCycleWiseForAJM As String = "Apply Head Load Day And Cycle Wise For AJM"
    Public Const JournalEntry As String = "Journal Entry"
    Public Const Inventory As String = "Inventory"
    Public Const InventoryNew As String = "Inventory New"
    Public Const PickBulkRoute As String = "Pick Bulk Route"
    Public Const ShowMultipleLegers As String = "Show Multiple Legers"
    Public Const LoadLedgerMixedMilk As String = "Load Ledger Mixed Milk"
    Public Const HeadLoadRODecimalPlace As String = "Head Load RO Decimal Place"
    Public Const HeadLoadROIncreaseAfter As String = "Head Load RO Increase After"
    Public Const ApplyUnpaidBank As String = "Apply Unpaid Bank"
    Public Const AllowZeroFATSNF As String = "Allow Zero FAT SNF"
    Public Const MilkRateRoundOffType As String = "Milk Rate Round Off Type"
    Public Const WeighingRoundSetting As String = "Weighing Round Setting"
    Public Const AverageDaysForFATSNF As String = "Average Days For FAT SNF"
    Public Const MarqueText As String = "Marque Text"
    Public Const CustomerVirtualACCounter As String = "Customer Virtual AC Counter"
    Public Const MaxRowsExcelDBTNEFTUploader As String = "Max Rows Excel DBT NEFT Uploader"
    Public Const ShowSampleNoOnBMC As String = "Fill Route Tanker No"
    Public Const DBTRevisePayment As String = "DBT Revise Payment"
    Public Const CycleWiseDBT As String = "Cycle Wise DBT"
    Public Const ShowTempratureOnBMC As String = "Fill Route Tanker No"
    Public Const PickMilkTransferInData As String = "Pick Milk Transfer In Data"
    Public Const FillRouteTankerNo As String = "Fill Route Tanker No"
    Public Const PickMilkPurchaseInvoiceQtyOrRecoQty As String = "Pick Milk Purchase Invoice Qty Or Reco Qty"
    Public Const AllowMPIncetiveQtyAboveBilledQty As String = "Allow MP Incentive Qty Above Billed Qty"
    Public Const RepeatBMCSampleNo As String = "Repeat BMC Sample No"
    Public Const HeaderFATSNFKGDecimalPlaces = "Header FAT SNF KG Decimal Places"
    Public Const SNFDecimalPlaces = "SNF Decimal Places"
    Public Const BMCSNFDecimalPlaces = "BMC SNF Decimal Places"
    Public Const BMCFATDecimalPlaces = "BMC FAT Decimal Places"
    Public Const BMCDCSApplyShift = "BMC DCS Apply Shift"
    Public Const AdjustFATSNFINOwnVSP = "Adjust FAT SNF IN Own VSP"
    Public Const AdjustQtyINOwnVSP = "Adjust Qty IN Own VSP"
    Public Const ApplySameDayShift = "Apply Same Day Shift"
    Public Const HideShiftCollection As String = "Hide Shift Collection"
    Public Const MilkCollectionPickBulkRoute As String = "Milk Collection Pick Bulk Route"
    Public Const OwnBMCCreateDRCRNote As String = "Own BMC Create DR CR Note"
    Public Const TDSOpeningBalanceConsiderCreditNote As String = "TDS Opening Balance Consider Credit Note"
    Public Const OwnBMCApplicationFATRatio As String = "Own BMC Application FAT Ratio"
    Public Const OwnBMCApplicationSNFRatio As String = "Own BMC Application SNF Ratio"
    Public Const FATSNFNoDecimalMCC As String = "FAT SNF No Decimal MCC"
    Public Const FATSNFNoDecimalDCS As String = "FAT SNF No Decimal DCS"
    Public Const ShowAllMCC As String = "Show All MCC"
    Public Const ShowDCSDetMerge As String = "Show DCS Detail in Merge"
    Public Const ShowAllDCS As String = "Show All DCS"
    Public Const ApplyGaze As String = "Apply Gaze"
    Public Const MilkCollectionFATSNFType As String = "Milk Collection FAT SNF Type"
    Public Const MilkCollectionFATSNFTypeHeader As String = "Milk Collection FAT SNF Type Header"
    Public Const JanAadharNoMandatory As String = "Jan Aadhar No Mandatory"
    Public Const StopUpdateForWeigingMilkReceipt As String = "Stop Update For Weiging Milk Receipt"
    Public Const CreatePOFromMultipleLocation As String = "Create PO From Multiple Location Indent"
    Public Const ApplyPashuAaharAndMineralMixture As String = "Apply Pashu Aahar And Mineral Mixture"
    Public Const BulkProcRunOneTypeGateEntry As String = "Bulk Proc Run One Type Gate Entry"
    Public Const ApplyZoneInDBT As String = "Apply Zone In DBT"
    Public Const MandatoryDCSMPIncetiveReco As String = "Mandatory DCS MP Incetive Reco"
    Public Const MatchingUOM As String = "MP Matching UOM"
    Public Const AllowMPQtyGreaterThanDCSQty As String = "Allow MP Qty > DCS Qty "
    Public Const AllowMPQtyLessThanDCSQty As String = "Allow MP Qty < DCS Qty "
    Public Const AllowMPQtyEqualToDCSQty As String = "Allow MP Qty = DCS Qty"
    Public Const BankIFSCCodeValidateByService As String = "Bank IFSC Code Validate By Service"
    Public Const ShowForgetPwd As String = "Show Forget Pwd"
    Public Const StopDBTBefore As String = "Stop DBT Before"
    Public Const StopResetPassword As String = "Stop Reset Password"
    Public Const OneDBTOneDoc As String = "One DBT One Doc"
    Public Const StopMPUpdate As String = "Stop MP Update"
    Public Const DisableUploaderNo As String = "Disable Uploader No"
    Public Const JPRDairyMandatoryColumn As String = "JPR Dairy Mandatory Column"
    Public Const VerifiedJanAadharNo As String = "Verified Jan Aadhar No"
    Public Const FarmerLevelCollection As String = "Farmer Level Collection"
    Public Const MultipleEntryScreen As String = "Multiple Entry Screen"
    Public Const MultipleEntryScreenAdmin As String = "Multiple Entry Screen Admin"
    Public Const QtyDecimalPlaces As String = "Qty Decimal Places"
    Public Const SkipPaymentCycle As String = "Skip Payment Cycle"
    Public Const TolleranceQty As String = "Tollerance Qty"
    Public Const TolleranceFAT As String = "Tollerance FAT"
    Public Const TolleranceSNF As String = "Tollerance SNF"
    Public Const IncludeOwnBMC As String = "Include Own BMC"
    Public Const TransactionForCurrentDateOnly As String = "Transaction For Current Date Only"
    Public Const AddPostFunctionality As String = "Add Post Functionality"
    Public Const DashboardDays As String = "Dashboard Days"
    Public Const UOM As String = "UOM"
    Public Const Shift As String = "Shift"
    Public Const AllowFutureDateBooking As String = "Allow Future Date Booking"
    Public Const ApplyOrderConfirmation As String = "Apply Order Confirmation"
    Public Const SelectDemandUOMAtRuntime = "Select Demand UOM at Runtime"
    Public Const ShowCatalogPrice = "Show Catalog Price"
    Public Const StopDemandUpdateForCNF = "Stop Demand Update for CNF"
    Public Const comtecxpertappsaras As String = "com.tecxpert.app.saras"
    Public Const comTecxpertSarasPro As String = "com.tecxpert.saras_pro"
    Public Const comAnchal_ucdfErp As String = "com.anchal_ucdf.erp"
    Public Const comtecxpertsoftposerode As String = "com.tecxpertsoft.amirthaa.erode"
    Public Const XpertMilkCollection As String = "XpertMilkCollection"
    Public Const MPIncentiveEntryApplyMonthly As String = "MP Incentive Entry Apply Monthly"
    Public Const DCSQtyDecimalPlaces As String = "DCS Qty Decimal Places"
    Public Const MPIncentiveEntryCycleWiseButNEFTMonthly As String = "MP Incentive Entry Cycle Wise But NEFT Monthly"
    Public Const MPIncentiveEntryMaxMilkLimit As String = "MP Incentive Entry Max Milk Limit"
    Public Const MPIncentiveEntryIncentiveRate As String = "MP Incentive Entry Incentive Rate"
    'Public Const TrendDiffValueForColor As String = "Trend Diff Value For Color"
    Public Const AllowPurReturnEvenIfPaymentDone As String = "Allow Pur Return Even If Complete Payment Done"
    Public Const TransporterCollection = "Transporter Collection"
    Public Const ProcurmentShiftUploaderNo = "Procurment Shift Uploader No"
    Public Const BholeBabaPaymentProcessProPrintStartDate = "Bhole Baba Payment Process Pro Print Start Date"
    Public Const BankCodeForApplyDocumentPaymentOFAssetLost = "Bank Code For Apply Document PaymentOF Asset Lost"
    Public Const SeprateDistanceMorningEveningShift = "Seprate Distance Morning/Evening Shift"
    Public Const TankerDispatchProvisionLocationSegment = "Tanker Dispatch Provision Location Segment"
    Public Const IncludeInceAndDedInFATSNFRate = "Include Ince And Ded In FAT SNF Rate"
    Public Const CanSaleAvgFATSNFPer = "Can Sale Avg FAT SNF Per"
    Public Const FATPerShouldBeMultipleOf5 = "FAT Per Should Be Multiple Of 5"
    Public Const PickFATSNFKgFromInventory = "Pick FAT SNF Kg From Inventory For Milk Register"
    Public Const ALLOWBOOKINGSHITWISE = "ALLOW BOOKING SHIT WISE"
    Public Const FindReasonWhyInvoiceIssueOccursOnErode = "FindReasonWhyInvoiceIssueOccursOnErode"
    Public Const ApplyGovtRulesInTDS = "Apply Govt Rules In TDS"
    Public Const AllowtoChangeFatANdSNFPerforHighClassVendorinGE = "AllowtoChangeFatANdSNFPerforHighClassVendorinGE"
    Public Const BookingMobileAppChangeorderByBookingQty = "Booking Mobile App Change order By Booking Qty"
    Public Const BookingMobileAppSetNxtDateofBookingOrder = "Booking Mobile App Set Nxt Date of Booking Order"
    Public Const MCCBioSyncDate = "MCC Bio Sync Date"
    Public Const EnableTDSforServiceVendorSeparately = "Enable TDS for Service Vendor Separately"
    Public Const VSPBillDocumentToBeAddedInMilkCost = "VSP Bill Document To Be Added In Milk Cost"
    Public Const DCSAddDedRODecimalPlace = "DCS Add Ded RO Decimal Place"
    Public Const DCSAddDedROIncreaseAfter = "DCS Add Ded RO Increase After"
    Public Const DCSAddDedROHeaderLevel = "DCS Add Ded RO Header Level"
    Public Const ApplyTDSValidationFrom = "Apply TDS Validation From"
    Public Const ConsiderPreviousCurrentFYForTCSTaxVendOutstanding = "ConsiderPreviousCurrentFYForTCSTaxVendOutstanding"
    Public Const AmountToCheckVendorOutstandingForTCSTax = "AmountToCheckVendorOutstandingForTCSTax"
    Public Const AllowtoChangeTCSBaseAmountPurchase = "AllowtoChangeTCSBaseAmountPurchase"
    Public Const AllowFatPerInanynumberofMultipesonBulkQC = "AllowFatPerInanynumberofMultipesonBulkQC"
    Public Const CreateSeparateInvoiceforeachrowinCansale = "CreateSeparateInvoiceforeachrowinCansale"
    Public Const AllowCanInformationintoGridForTankerDispatch = "AllowCanInformationintoGridForTankerDispatch"
    Public Const TCSRateforVendorWithPanNo = "TCSRateforVendorWithPanNo"
    Public Const ShowAvailableQtyOnDairyBooking = "ShowAvailableQtyOnDairyBooking"
    Public Const AllowTransferVSPAmtToFarmerinVCGL = "AllowTransferVSPAmtToFarmerinVCGL"
    Public Const TCSRateforVendorWithoutPanNo = "TCSRateforVendorWithoutPanNo"
    Public Const IncetiveEntryApplyArrear As String = "Incetive Entry Apply Arrear"
    Public Const VSPDayWiseIncentiveAtSRN As String = "VSP Day Wise Incentive At SRN"
    Public Const IncentiveProcessPaymentStartDate As String = "Incentive Process Payment Start Date"
    Public Const CreateJVOFRGPNRGPAndItsRetrun As String = "Create JV Of RGP NRGP And Its Retrun"
    Public Const FixedAssetAcquisitionEntryHitInventoryMovement As String = "Acquisition Entry Hit Inventory Movement"
    Public Const CalculateDeductionByStdQtyinBulkMilkSRN = "Calculate Deduction By Std Qty in Bulk Milk SRN"
    Public Const CreateAPinvoiceofsalaryemployeewiseduringsalarygen = "CreateAPinvoiceofsalaryemployeewiseduringsalarygen"
    Public Const ApplyTaxInBulkMilkPurchaseInvoice = "Apply Tax In Bulk Milk Purchase Invoice"
    Public Const TankerDispatchIntermittentSingleGateIn = "Tanker Dispatch Intermittent Single Gate In"
    Public Const MilkSRNFATSNFDecimalPlaces = "Milk SRN FAT SNF Decimal Places"
    Public Const JobWorkOutwardComsumeItemAccordingToBOM = "Job Work Outward Comsume Item According to BOM"
    Public Const JWOutWardForRCDF As String = "Job Work Out Ward For RCDF"
    Public Const CreateARAdjAPDebitnoteforEmployeesinMCCMS = "CreateARAdjAPDebitnoteforEmployeesinMCCMS"
    Public Const ApplyCardSaleInvoiceOnlyWithCardSaleAdvance = "ApplyCardSaleInvoiceOnlyWithCardSaleAdvance"
    Public Const StartDateforDispatchSchedular = "StartDateforDispatchSchedular"
    Public Const StockRecoRateTunning = "Stock Reco Rate Tunning"
    Public Const BankRecoCheckFutureDocuments = "Bank Reco Check Future Documents"
    Public Const FATSNFRate = "Production Utility FAT/SNF Rate"
    Public Const ProductionRemoveFATSNFKgTollerance = "Production Remove FAT SNF Kg Tollerance"
    Public Const SubLocationForTaxableItemDairyDispatch = "SubLocationForTaxableItemDairyDispatch"
    Public Const SubLocationForNonTaxableItemDairyDispatch = "SubLocationForNonTaxableItemDairyDispatch"
    Public Const CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo = "CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo"
    Public Const EnableTCSRateValidityFrom01July2021 = "EnableTCSRateValidityFrom01July2021"
    Public Const ProductionCheckFATKg = "Production Check FAT Kg"
    Public Const ProductionCheckSNFKg = "Production Check SNF Kg"
    Public Const ProductionOnlyOneIssueAgainstBatch = "Production Only One Issue Against Batch"
    Public Const MilkShiftEndAutoAdjInItemCode = "Milk Shift End Auto Adjustment In Item Code"
    Public Const StockCheckOnPostForDairyDispatchMultiple = "StockCheckOnPostForDairyDispatchMultiple"
    Public Const BatchFileCounter = "BatchFileCounter"
    Public Const checkstockMRPwise = "checkstockMRPwise"
    Public Const TCSRateforCustomerWithPanNo = "TCSRateforCustomerWithPanNo"
    Public Const TCSRateforCustomerWithoutPanNo = "TCSRateforCustomerWithoutPanNo"
    Public Const ScrapTCSRateforCustomerWithPanNo = "ScrapTCSRateforCustomerWithPanNo"
    Public Const ScrapTCSRateforCustomerWithoutPanNo = "ScrapTCSRateforCustomerWithoutPanNo"
    Public Const TCSTaxApplicableOnbulkSale = "TCSTaxApplicableOnbulkSale"
    Public Const TCSTaxApplicableOnCanSale = "TCSTaxApplicableOnCanSale"
    Public Const AllowtoChangeTCSBaseAmount = "AllowtoChangeTCSBaseAmount"
    Public Const EInvoiceVendor = "EInvoiceVendor"
    Public Const TokenTimeReGenForEinvoice = "TokenTimeReGenForEinvoice"
    Public Const TokenTimeReGenForEWayBill = "TokenTimeReGenForEWayBill"
    Public Const ItemCostZeroOnStoreAdjForTypeFlushing = "ItemCostZeroOnStoreAdjForTypeFlushing"
    Public Const AutoCreateSaleInvoice = "Auto Create Sale Invoice"
    Public Const AmountToCheckCustomerOutstandingForTCSTax = "AmountToCheckCustomerOutstandingForTCSTax"
    Public Const DonotConsiderDirectARInvoiceinCustomerOutTCS = "DonotConsiderDirectARInvoiceinCustomerOutTCS"
    Public Const ConsiderPreviousCurrentFYForTCSTaxCustOutstanding = "ConsiderPreviousCurrentFYForTCSTaxCustOutstanding"
    Public Const AllowJEofDifferentLocationOnJournalEntry = "AllowJEofDifferentLocationOnJournalEntry"
    Public Const MilkProcurementBatchPosting = "Milk Procurement Batch Posting"
    Public Const popupcustomernamewhileupdating = "popupcustomernamewhileupdating"
    Public Const ApplyFEFO = "ApplyFEFO"
    Public Const EnterLocationForJWEStimationOutPackingMaterial = "EnterLocationForJWEStimationOutPackingMaterial"
    Public Const CreateJVofPackingMaterialofJWInwardinJWEstimate = "CreateJVofPackingMaterialofJWInwardinJWEstimate"
    Public Const AllowtoenterrateIntoJobWorkDispatch = "AllowtoenterrateIntoJobWorkDispatch"
    Public Const ConsiderUnpostedDocForBalance = "Consider Unposted Doc For Balance"
    Public Const TankerDispatchAvgFATSNFPer = "Tanker Dispatch Avg FAT SNF Per"
    Public Const DisableToPickMainLocationType = "Disable To Pick Main Location Type"
    Public Const MultipleMCCFinder = "Multiple MCC Finder In VSP Bill Generation"
    Public Const MilkProcurementSNF2DecimalPlaces = "Milk Procurement SNF 2 Decimal Places"
    Public Const SelectMilkRejectDefaulterManually = "Select Milk Reject Defaulter Manually"
    Public Const DayWiseCustomerIncentiveCalculation = "Day Wise Customer Incentive Calculation"
    Public Const CustomerIncetiveAutoSecuity = "Create Auto Security on Customer Incentive"
    Public Const CustomerIncetiveBankForSecuity = "Customer Incetive Bank For Secuity"
    Public Const CustomerIncetivePaymentModeForSecuity = "Customer Incetive Payment Mode For Secuity"
    Public Const RepeatDeductionAndVendor = "Repeat Deduction And Vendor"
    Public Const LastMilkReceiptQtyTollerance = "Last Milk Receipt Qty Tollerance"
    Public Const ShowMCCFinderInPaymentProcess = "Show MCC Finder In Payment Process"
    Public Const UseDescInsteadOFCodeOnMCCMAterialSale = "UseDescInsteadOFCodeOnMCCMAterialSale"
    Public Const AllowShelfLifeMandatoryOnFG = "AllowShelfLifeMandatoryOnFG"
    Public Const RefundknockoffwithCreditNote = "Refund knockoff with Credit Note"
    Public Const CreateInvoiceAutomaticallyOnbulkDispatch = "CreateInvoiceAutomaticallyOnbulkDispatch"
    Public Const CrateReceivingWithMultipleRoute = "CrateReceivingWithMultipleRoute"
    Public Const ShowMulMRPOfSameItemOnDairyBookingCustomer = "ShowMulMRPOfSameItemOnDairyBookingCustomer"
    Public Const AllowZeroQtyOnDairyBooking = "AllowZeroQtyOnDairyBooking"
    Public Const UseCutOffTimeonRouteForERP = "UseCutOffTimeonRouteForERP"
    Public Const AllowZeroQtyOnDairyBookingUploader = "AllowZeroQtyOnDairyBookingUploader"
    Public Const AllowtoPostNoOFDocofDOatatime = "AllowtoPostNoOFDocofDOatatime"
    Public Const DonotIncludeSecurityInCustomerOutstanding = "DonotIncludeSecurityInCustomerOutstanding"
    Public Const IsManualTCS = "Manual TCS"
    Public Const DefaultLocationForCardSaleIntegration = "DefaultLocationForCardSaleIntegration"
    Public Const ApplyNoGSTCreditIndependentlyOnVendorServiceCharge = "ApplyNoGSTCreditIndependentlyOnVendorServiceCharge"
    Public Const CheckNoOfDaysforCardSaleBooking = "CheckNoOfDaysforCardSaleBooking"
    Public Const MaxNoOfBookingAllowedThroughBookingApp = "MaxNoOfBookingAllowedThroughBookingApp"
    Public Const ApplyBothtsrateAndFatRateinBulkProcurement = "ApplyBothtsrateAndFatRateinBulkProcurement"
    Public Const PickTCAForStockTransferAndTankerDispatch = "PickTCAForStockTransferAndTankerDispatch"
    Public Const ShowBookingTypeDropDownonDairyBookingCustomer = "ShowBookingTypeDropDownonDairyBookingCustomer"
    'Public Const AutoCreateGateEntryTillMilkTransferInForIntermittent = "AutoCreateGETillMilkTransferInForIntermittent"
    Public Const AllowItemConversionAutomation = "AllowItemConversionAutomation"
    Public Const AutoMilkTransferInDateSameasWeighmentDate = "AutoMilkTransferInDateSameasWeighmentDate"
    Public Const EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt = "EnableGoButtonofRcptEntryWithoutEnteringReceiptAmt"
    Public Const ShowOutstandingAmtofCustomerOnQuickBookEntry = "ShowOutstandingAmtofCustomerOnQuickBookEntry"
    Public Const DoNotCreateJournalVoucheronJobWorkDispatch = "DoNotCreateJournalVoucheronJobWorkDispatch"
    Public Const PickCostOFMaterialSaleFromPriceMaster = "PickCostOFMaterialSaleFromPriceMaster"
    Public Const ForceToSelectIteminGateEntry = "ForceToSelectIteminGateEntry"
    Public Const CreateProvisionforBulkContractorInGateIn = "CreateProvisionforBulkContractorInGateIn"
    Public Const DoNotCreateAdjustmentonMilkTransferInGL = "DoNotCreateAdjustmentonMilkTransferInGL"
    Public Const Donotshowtrasnfertransactionsbydefault = "Do not show trasnfer transactions by default"
    Public Const AllowManualItemPriceOnMCCSale = "AllowManualItemPriceOnMCCSale"
    Public Const ExcludeKKFAndMandiForDCS = "Exclude KKF And Mandi For DCS"
    Public Const DonotAllowtoChangeUOMinDairyBookingCustomer = "DonotAllowtoChangeUOMinDairyBookingCustomer"
    Public Const AutoPopulateItemCodeOnDairyBooking = "AutoPopulateItemCodeOnDairyBooking"
    Public Const ItemwiseCorrectionFactoronQC = "ItemwiseCorrectionFactoronQC"
    Public Const Allow0FatPerOnBulkSaleQualityCheckScreen = "Allow0FatPerOnBulkSaleQualityCheckScreen"
    Public Const CreateMultipleDispatchWithoutSelectingVehicle = "CreateMultipleDispatchWithoutSelectingVehicle"
    Public Const RejectiononQCforSeparationofBulkProcurementMCC = "RejectiononQCforSeparationofBulkProcurementMCC"
    Public Const ParameterForSNFatQC = "ParameterForSNFatQC"
    Public Const TreatCLR50As52 = "Treat CLR 50 As 52"
    Public Const AllowmultipleconsumptionLocation = "Allow Multiple Consumption Location"
    Public Const PickCostFromItemMaster As String = "Pick Cost From Item Master(FG)"
    Public Const EditItemCost As String = "Edit Item Cost"
    Public Const AllowBatchQtyin3DecimalPlaces As String = "Allow Batch Qty in 3 Decimal Places"
    Public Const DisplayAverageFatSNFMPWise As String = "DisplayAverageFatSNFMPWise"
    Public Const SentschemecogsinAnotherAccount As String = "SentschemecogsinAnotherAccount"
    Public Const OPkmMandatoryonDS As String = "OP km mandatory on Dairy dispatch sale"
    Public Const ApplyBankChargesasperSlabonBankMaster As String = "ApplyBankChargesasperSlabonBankMaster"
    Public Const UseKGLitreConversionInBulkSaleAsperCLRCalculation As String = "UseKGLitreConversionInBulkSaleAsperCLRCalculation"
    Public Const ShowLastUnitCostZeroForNonInventoryItemOnPO As String = "ShowLastUnitCostZeroForNonInventoryItemOnPO"
    Public Const ManualBatchNoMandatoryOnBatchOrderScreen As String = "ManualBatchNoMandatoryOnBatchOrderScreen"
    Public Const ShowSiloLocationItemLocationwise = "ShowSiloLocationItemLocationwise"
    Public Const GateEntryChamberwisewithManualTankerEntry = "GateEntryChamberwisewithManualTankerEntry"
    Public Const AllowPriceMappingOnBulkSRNinChamberBulkProc = "AllowPriceMappingOnBulkSRNinChamberBulkProc"
    Public Const EnableManualCrateonTaxableDairyDispatch = "EnableManualCrateonTaxableDairyDispatch"
    Public Const CreateOpeningEntryAutomatically = "Create Opening Entry Automatically"
    Public Const AllowCratePhysicalStock = "AllowCratePhysicalStock"
    Public Const AllowSiloMilkTransfertoMainLocation = "AllowSiloMilkTransfertoMainLocation"
    Public Const EnableAutoDocNoShipToLocation = "Enable Auto DocNo ShipToLocation"
    Public Const ChangeFATCLRafterspecialApprovalonQC = "Change FAT CLR after special approval on QC"
    Public Const CreateCommonSeriesLocationwiseForAllSale = "Create Common Series Location Wise For All Sale"
    Public Const EnableCustomerPODetailonDairyBooking = "EnableCustomerPODetailonDairyBooking"
    Public Const CreateCommonDairyDispatchforFreshAmbient = "CreateCommonDairyDispatchforFreshAmbient"
    Public Const CreateSeperateTaxInvForFOCIteminNonTaxdispatch = "CreateSeperateTaxInvForFOCIteminNonTaxdispatch"
    Public Const CreateSeperateSeriesforRefDocARinvforCreditdebit = "CreateSeperateSeriesforRefDocARinvforCreditdebit"
    Public Const CreateSeperateSeriesforRefDocAPinvforCreditdebit = "CreateSeperateSeriesforRefDocAPinvforCreditdebit"
    Public Const AllowUniqueNoOnMilkTransferInandTankDis = "Allow UniqueNo On Milk TransferIn and Tanker Dispatch"
    Public Const DoNotConsiderCustomerCreditLimit = "Do not Consider Customer Credit Limit"
    Public Const AllowVSPMasterAutoPrefix = "Allow VSP Master Auto Prefix"
    Public Const EnableGSTRelatedfields = "EnableGSTRelatedfields"
    Public Const AllowPostGSTPayment = "AllowPostGSTPayment"
    Public Const ConsiderSiloCapicityForStockIn = "ConsiderSiloCapicityForStockIn"
    Public Const AllowBulkProcTransDateSameasGateEntryDate = "Allow BulkProc TransDate Same as GateEntry Date"
    Public Const AllowDifferentStateofChildCustomerOnPS = "Allow Different State of Child Customer On PS"
    Public Const AllowProvisionknokoffOnAPInvoice = "Allow Provision knokoff On APInvoice"
    Public Const AllowTransactionFiltersOnCustomerlegder = "AllowTransactionFiltersOnCustomerlegder"
    Public Const GroupCustomerlegderZoneWiseAreaWise = "GroupCustomerlegderZoneWiseAreaWise"
    Public Const CustomerDashboardWithOpeningAndClosing = "CustomerDashboardWithOpeningAndClosing"
    Public Const AllowtoSHOWParentChildCustomer = "AllowtoSHOWParentChildCustomer"
    Public Const AllowtoMakeApplyDocOnbyDefault = "AllowtoMakeApplyDocOnbyDefault"
    Public Const PenaltyPercentage = "Penalty%"
    Public Const CreateProvisionJournalEntryForSale = "Create Provision JournalEntry Sale"
    Public Const CreateProvisionJournalEntryForTankerDispatch = "Create Provision JournalEntry Tanker Dispatch"
    Public Const ShowDairySaleModuleOnBulkPosting = "Show DairySale Module On BulkPosting"
    Public Const ShowGateEntryTypeonGateEntryBulkProc = "Show GateEntryType on GateEntry BulkProc"
    Public Const AllowAutoBulkMilkSRNonWeighmentBulkProc = "Allow Auto BulkMilkSRN on Weighment BulkProc"
    Public Const CheckParameterRangerProcurementTypewise = "Check ParameterRanger ProcurementType wise"
    Public Const PickCorrectionFactorProcurementTypewise = "Pick CorrectionFactor ProcurementType wise"
    Public Const CalculateTaxRatefromItemwsieTaxOnSale = "CalculateTaxRatefromItemwsieTaxOnSale"
    Public Const AllowBulkProcMCCwithoutTankerDispatch = "Allow Bulk Proc MCC without Tanker Dispatch"
    Public Const AllowJobWorkonGateEntryBulkProc = "Allow Job Work ON Gate Enty BulkProc"
    Public Const ApplyMaTransRateOnMultChamberTankerDis = "ApplyMaTransferRateOnMultilpeChamberTankerDispatch"
    Public Const AllowManualPriceONBulkPO = "Allow Manual Price ON Bulk PO"
    Public Const AllowReverseUnpost = "Allow Reverse and Unpost"
    Public Const AllowtoSetNoOfTransactionsforSetOff = "AllowtoSetNoOfTransactionsforSetOff"
    Public Const AllowtoUnlockTransactionsforSetOff = "AllowtoUnlockTransactionsforSetOff"
    Public Const AllowtoSetReceiptAmountForCashTransaction = "AllowtoSetReceiptAmountForCashTransaction"
    Public Const AllowtoSetPaymentAmountForCashTransaction = "AllowtoSetPaymentAmountForCashTransaction"
    Public Const AllowtoShowCreditBalanceonCustomerAgeing = "Allow to Show Credit Balance"
    Public Const AllowtoShowDebitBalanceonVendorAgeing = "Allow to Show Debit Balance"
    Public Const ConsiderOpeningDocintoBucketsInAgeing = "ConsiderOpeningDocintoBucketsInAgeing"
    Public Const ConsiderOpeningDocintoBucketsonCustomerAgeing = "ConsiderOpeningDocintoBucketsonCustomerAgeing"
    Public Const AllowtoSetoffDocDateWise = "AllowtoSetoffDocDateWise"
    Public Const AllowtoSkipJournalEntryofPaymentandReceiptforAD = "AllowtoSkipJournalEntryofPaymentandReceiptforAD"
    Public Const AllowtoSkipfunctionalityafterSRNOnBulkProcurement = "AllowtoSkipfunctionalityafterSRNOnBulkProcurement"
    Public Const AllowtoNegativeStockInventoryAtTankerDispatch = "AllowtoNegativeStockInventoryAtTankerDispatch"
    Public Const AllowtoNegativeFATSNFKgAtTankerDispatch = "AllowtoNegativeFATSNFKgAtTankerDispatch"
    Public Const AllowtoEmployeeSalaryIntegration = "AllowtoEmployeeSalaryIntegration"
    Public Const AllowtoFutureDateTransForPDCCheque = "AllowtoFutureDateTransForPDCCheque"
    Public Const WeightOfCanForCanSale = "Weight Of Can For Can Sale"
    Public Const RunBulkProcWithoutMilkGrade As String = "Run Bulk Proc without Milk Grade"
    Public Const DaysToStartAutoLock As String = "Days ToStart AutoLock"
    Public Const AllowRandomOnlyOneSecondaryQC As String = "Allow Random OnlyOne SecondaryQC"
    Public Const AllowGateEntryAgainstPO As String = "Allow GateEntry Against PO"
    Public Const SeparateDairyDispatchTaxableNonTaxable As String = "Separate DairyDispatch TaxableNon Taxable"
    Public Const RunBatchFifowise As String = "Run Batch Fifowise"
    Public Const RunBatchFifowisewithModifyfunctionality As String = "Run Batch Fifowise with Modify functionality"
    Public Const PromptTimeToPostTransactions As String = "Prompt Time ToPost Transactions"
    Public Const CreateFreshInvoiceOnDispatchSave As String = "Create FreshInvoice OnDispatch Save"
    Public Const AllowLockTransactionUserwise As String = "AllowLockTransactionUserwise"
    Public Const AllowAutoLockTransaction As String = "AllowAutoLockTransaction"
    Public Const AllowDefaultBankCodeforCreditNote As String = "AllowDefaultBankCodeforCreditNote"
    Public Const AllowCreditNoteWithoutReference As String = "AllowCreditNoteWithoutReference"
    Public Const AllowUseApplyDocSeriesForReceipt As String = "AllowUseApplyDocSeriesForReceipt"
    Public Const AllowUseApplyDocSeriesForPayment As String = "AllowUseApplyDocSeriesForPayment"
    Public Const AllowCreditNoteWithoutReferenceonAP As String = "AllowCreditNoteWithoutReferenceonAP"
    Public Const AllowBranchAcconReceiptPrint As String = "AllowBranchAcconReceiptPrint"
    Public Const SecurityDocumentKnockOffonReceipt As String = "SecurityDocumentKnockOffonReceipt"
    Public Const AllowFreshInvoiceAutoPost As String = "AllowFreshInvoiceAutoPost"
    Public Const AllowReceiptThroughSO As String = "AllowReceiptThroughSO"
    Public Const AllowSetOffUntilTransactionsnotend As String = "AllowSetOffUntilTransactionsnotend"
    Public Const AllowGateReturn As String = "AllowGateReturn"
    Public Const CalculateCommOnCSATransWOConversion As String = "CalculateCommOnCSATransWOConversion"
    Public Const AllowSRNWithoutShortageRejection As String = "AllowSRNWithoutShortageRejection"
    Public Const AllowPurchaseModulewithUniqueItem As String = "AllowPurchaseModulewithUniqueItem"
    Public Const GrossWeightUnit As String = "GrossWeightUnit"
    Public Const ExpiryDaysBulkProcurementPriceChart As String = "ExpiryDays BulkProcurement PriceChart"
    Public Const ShowSchemeItemRateonDairyDispatch As String = "Show Scheme Item Rate on Dairy Dispatch"
    Public Const ShowSchemeItemRateonDairyDispatchTaxable As String = "Show Scheme Item Rate on Dairy Dispatch Taxable"
    Public Const AutoCalculateCrateOnDairyDispatch As String = "Auto Calculate Crate on Dairy Dispatch"
    Public Const AutoCalculateCANOnDairyDispatch As String = "Auto Calculate Can on Dairy Dispatch"
    Public Const GrossWtFromItemMasterONProductSale As String = "Gross Wt. from item master on Product Sale"
    Public Const AllowFreshPriceChartOnBookingProductSale = "Allow Fresh Price Chart on Booking PS"
    Public Const CreateVatSeriesForProductExciseinvoice = "Create Vat Series for PS Excise invoice"
    Public Const AllowFreshPriceChartOnProductSale = "Allow Fresh Price Chart on Product Sale"
    Public Const ShowUnloadingandWeighmentSequencewise = "Show Unloading Weighment sequence wise"
    Public Const ShowBothTankertypeOnCleaning = "Show Both Tanker Type on Cleaning"
    Public Const isCleaningMandatoryBeforeGateout = "Is Cleaning Mandatory before Gate Out"
    Public Const AllowBulkProcurementSequencewise = "Allow Bulk Procurement Sequence wise"
    Public Const ShowItemLocationWiseonDairyBooking = "Show Item Location wise on Dairy Booking"
    Public Const SeprateDemandForMorningEveningShift = "Seprate Demand For Morning Evening Shift"
    Public Const CheckOutstandingCreditLimitOnBooking = "Check Customer Outstanding on Booking"
    Public Const AllowBulkPriceChartMultiplepriceToMultipleVendor = "Allow BulkPrice Multiple Price to Mult Vendor"
    Public Const isItemMilkType As String = "Is Item Milk Type"
    Public Const isPriceChartGradeWise As String = "Is Price Chart Grade Wise"
    Public Const isCreateBulkProcPriceChartItemWise As String = "Create Bulk Procurement price chart-Itemwise"
    Public Const isFarmerPaymentCycle As String = "is Farmer Payment Cycle"
    Public Const CowFATPer As String = "Cow FAT Per"
    Public Const MixFATPer As String = "Mix FAT Per"
    Public Const CmtSetting As String = "Comment Setting"
    Public Const AreaWiseBilling As String = "Area Wise Billing"
    Public Const Hidedetaildate As String = "Hide Detail Date"
    Public Const DeleteTempData As String = "Delete Temp Data"
    Public Const DeleteTempDataFromTime As String = "Delete Temp Data From Time"
    Public Const DeleteTempDataToTime As String = "Delete Temp Data To Time"
    Public Const SkipDatePayPeriodMaster As String = "Skip Date Pay Period Master"

    Public Const ShowNewFormatofPDF As String = "Show New Format of PDF"
    Public Const PaymentProcessPrintInHindi As String = "Payment Process Print In Hindi"
    Public Const VendorBankAdviceForSWM As String = "Vendor Bank Advice For SWM"
    Public Const GateEntryTankerFromTankerMaster As String = "Gate Entry tanker From Master"
    Public Const QualityThenWeighmentinBulkProcurement As String = "First QC then Weighment"
    Public Const isIntimationRequired As String = "Show Intimation Screen"
    Public Const ShowOptionOnItemMasterChangeItemRate = "Show option on Item Change Rate on DDispatch"
    Public Const SHowOptionOnLocationForDairyDispatchfromDOorGatepass = "Show option on loc For DDispatch from DO/GP"
    Public Const showPostrequiredforBulkSale = "showPostrequiredforBulkSale"
    Public Const ApplyDocumentDate = "Apply Document Date"
    Public Const AllowStockCheckatDOLevel = "AllowStockCheckatDOLevel"
    Public Const AllowAdditionalWeightinPercentage = "AllowAdditionalWeightinPercentage"
    Public Const EnterAdditionalWeight = "EnterAdditionalWeight"
    Public Const AllowTankerBasedonVendorofGE = "Allow Tanker Based on Vendor of GE"
    Public Const AllowUseBoilingParameteronParameterMaster = "Allow Use Boiling Parameter on Parameter Master"
    Public Const DairyDispatchFromDeliveryNote = "Dairy Dispatch From Delivery Note"
    Public Const ItemTypeForDairyBooking = "Item Type Fresh or Ambient For Dairy Booking"
    Public Const AllowStockToleranceNegative = "Allow Stock Tolerance Negative"
    Public Const StockToleranceLimit = "Stock Tolerance Limit"
    Public Const ShowAllMenu As String = "Show All Menu"
    Public Const POSeriesWithoutItemTypewise As String = "POSeriesWithoutItemTypeWise"
    Public Const WorkApprovalFlowInERP As String = "Work Approval Flow in ERP"
    Public Const BOOKINGFINDER_ON_CSASALEPATTI As String = "CSA Sale Patti With Booking Knock-off"
    Public Const OpenAvailorEmptyStckLocationOn_Standardization As String = "Open Avail./Empty Location on Standardization"
    Public Const BOM_Amend_Pswd As String = "Amendment Password for BOM"
    Public Const ProductionFATSNF_KG_Unit As String = "ProductionFATSNF_KG_Unit"
    Public Const RemoveGridFatSnfValidation As String = "Remove Grid Fat Snf Validation"
    Public Const ChangeRateAT_CSA_Return As String = "Rate Change At CSA Transfer Return"
    Public Const VehicleCapacityUnit As String = "VehicleCapacityUnit"
    Public Const StopGLEntryForConsignmentAtCSATransfer As String = "Stop GL Consignment at CSA Transfer"
    Public Const CSATransfer_SalePatti_All_Tax_Open As String = "Open All Tax Mapped With Location In CSA"
    Public Const Emps1 As String = "Emps1"
    Public Const Agency As String = "Agency"
    Public Const UploaderPassword As String = "Uploader Password"
    Public Const BankUploaderPassword As String = "Bank Uploader Password"
    Public Const CreateGLForTransfer As String = "CreateGLForTransfer"
    Public Const Category1 As String = "Category1"
    Public Const AllowDesignAtRunTime As String = "AllowDesignRunTime"
    Public Const Category2 As String = "Category2"
    Public Const Category3 As String = "Category3"
    Public Const SkipDiffGLOnPI As String = "SkipDiffGLOnPI"
    Public Const A As String = "A"
    Public Const SkipCogsEntry As String = "SkipCogsEntry"
    Public Const HY As String = "HY"
    Public Const M As String = "M"
    Public Const Q As String = "Q"
    Public Const L As String = "L"
    Public Const R As String = "R"
    Public Const U As String = "U"
    Public Const POWITHREQ As String = "POWITHREQ"
    Public Const DisplayReasonOnDelete As String = "DisplayReasonOnDelete"
    Public Const DisplayReasonOnUpdateAfterPost As String = "DisplayReasonOnUpdateAfterPost"
    Public Const Importbulkdatafromexcelsheet As String = "Importbulkdatafromexcelsheet"
    Public Const Distributor As String = "Distributor"
    Public Const POPUPITEMREORDERLEVEL As String = "POPUPITEMREORDERLEVEL"
    Public Const EMP02 As String = "EMP02"
    Public Const Driver As String = "Driver"
    Public Const ZM As String = "ZM"
    Public Const TSM As String = "TSM"
    Public Const ASM As String = "ASM"
    Public Const Emps2 As String = "Emps2"
    Public Const EMP3 As String = "EMP3"
    Public Const ALLOW As String = "ALLOW"
    Public Const BASIC As String = "BASIC"
    Public Const BONUS As String = "BONUS"
    Public Const COEPS As String = "COEPS"
    Public Const COESI As String = "COESI"
    Public Const COPF As String = "COPF"
    Public Const DA As String = "DA"
    Public Const DEDUCT As String = "DEDUCT"
    Public Const EMPESI As String = "EMPESI"
    Public Const MaxRowsForQuickExport As String = "Q-EXP-MX-RW"
    Public Const ShowPurchaseControlAc As String = "ShowPurchaseControlAc"
    Public Const CreateTransferInGL As String = "CreateTransferInGL"
    Public Const CreateTankerDispatchGL As String = "CreateTankerDispatchGL"
    Public Const PostTankerDispatchWithZeroAvgCost As String = "PostTankerDispatchWithZeroAvgCost"
    Public Const EPF As String = "EPF"
    Public Const HRA As String = "HRA"
    Public Const HRR As String = "HRR"
    Public Const LOAN As String = "LOAN"
    Public Const OT As String = "OT"
    Public Const OTHER As String = "OTHER"
    Public Const RMBT As String = "RMBT"
    Public Const TA As String = "TA"
    Public Const TDS As String = "TDS"
    Public Const Conveyance As String = "Conveyance"
    Public Const LaborUnFairFund As String = "LUF"
    Public Const LaborWelFairFund As String = "LWF"
    Public Const AllowSkippingPrevDocumentsOnPaymentProcess As String = "AllowSkippingPrevDocumentsOnPaymentProcess"
    Public Const PrefixGenerator As String = "PrefixGenerator"
    Public Const DuplicateRoute As String = "Duplicate Route"
    Public Const RJ As String = "RJ"
    Public Const EMP01 As String = "EMP01"
    Public Const SIRevers As String = "SIRevers"
    Public Const SIReversAndCreate As String = "SIReversAndCreate"
    Public Const CustomerTenderAmendment As String = "CustomerTenderAmendment"
    Public Const AllowBoothWiseReturn As String = "AllowBoothWiseReturn"
    Public Const SICrateEntry As String = "SICrateEntry"
    Public Const ShuffleDemand As String = "ShuffleDemand"
    Public Const ItemAllowDecimal As String = "ItemAllowDecimal"
    Public Const UpdatePassword As String = "Update Password"
    Public Const GatePassPasswod As String = "GatePass Password"
    Public Const MulProcDedReversAndCreate As String = "MulProcDedReversAndCreate"
    Public Const PPMrpReversAndCreate As String = "PPMrpReversAndCreate"
    Public Const WEUpdateAfterPost As String = "WEUpdateAfterPost"
    Public Const GEUpdateAfterPost As String = "GEUpdateAfterPost"
    Public Const PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE As String = "PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE"
    Public Const GEUpdatePriceChart As String = "GEUpdatePriceChart"
    Public Const SetCSATransferwithZeroOnSalePatti As String = "SetCSATransferwithZeroOnSalePatti"
    Public Const POAmendment As String = "POAmendment"
    Public Const BulkInvoiceDelete As String = "BulkInvoiceDelete"
    Public Const BulkSaleSequence As String = "BulkSaleSequence"
    Public Const BulkQCTableHavingUniqueKey As String = "BulkQCTableHavingUniqueKey"
    Public Const SrPath As String = "SrPath"
    Public Const TempProvisional As String = "TempProvisional"
    Public Const LoadInRollback As String = "LoadInRollback"
    Public Const Sunday As String = "Sunday"
    Public Const Monday As String = "Monday"
    Public Const Tuesday As String = "Tuesday"
    Public Const Wednesday As String = "Wednesday"
    Public Const Thursday As String = "Thursday"
    Public Const Friday As String = "Friday"
    Public Const Saturday As String = "Saturday"
    Public Const o As String = "0"
    Public Const EnableMilkProc As String = "EnableMilkProc"
    Public Const LCCreationPwd As String = "LCCreationPwd "
    Public Const ShowQtySum_in_GRN_MRN_SRN As String = "ShowQtySum_in_GRN_MRN_SRN"
    Public Const SalesInvoice As String = "Sales Invoice"
    Public Const LOReceiptDefaultBankForSettlement As String = "Default Bank For Settlement"
    Public Const LOReceiptPaymentTypeForSettlement As String = "Default Payment Type For Settlement"
    Public Const ALLOWANYBO As String = "Allow Any Type of BO"
    Public Const ALLOWCBOSBO As String = "Allow Child and SubChild BO"
    Public Const PROVISIONENTRYONSTOCKTRANSFER As String = "ProvisionOnStockTransfer"
    Public Const INDUSTRYTYPE As String = "Industry Type"
    Public Const Transfer As String = "Transfer"
    Public Const DefaultTypeFC As String = "FC"
    Public Const DefaultTypeFB As String = "FB"
    Public Const DefaultTypeSH As String = "SH"
    Public Const SalesmanPhysicalLocation As String = "SPL"
    Public Const IndentTolerence As String = "IndentTolerence"
    Public Const AskForDate As String = "AskForDate"
    Public Const PickMachineDateForTran As String = "PickMachineDateForTran"
    Public Const ReqLimitOnSRN As String = "ReqLimitOnSRN"
    Public Const AutoLoadinFromLocation As String = "AutoLoadinFromLocation"
    Public Const CrreateTransferShipmentJE As String = "CrreateTransferShipmentJE"
    Public Const IsNotIncludeWasteQtyInCal As String = "IsNotIncludeWasteQtyInCal"
    Public Const IsConsiderOutTypeDocForBalance As String = "IsConsiderOutTypeDocForBalance"
    Public Const BankTransferRunPaymentCounter As String = "BankTransferRunPaymentCounter"
    Public Const PaymentReceiptTypeRunReceiptCounter As String = "PaymentReceiptTypeRunReceiptCounter"
    Public Const CounterFinancialYearStyle As String = "CounterFinancialYearStyle"
    Public Const LinkFinancialYearStyleWithGSTDate As String = "LinkFinancialYearStyleWithGSTDate"
    Public Const ApplyLocationWisePrefix As String = "Apply Location Wise Prefix"
    Public Const CashDiscountFromClaimMaster As String = "CashDiscountFromClaimMaster"
    Public Const TransferTransTypeRouteHide As String = "TransferTransTypeRouteHide"
    Public Const AllowNegtiveOfSaleInvoiceBlanceAmt As String = "AllowNegtiveOfSaleInvoiceBlanceAmt"
    Public Const SalesRateEditable As String = "Sales Rate Editable"
    Public Const RunDemoERP As String = "RunDemoERP"
    Public Const IsKDIL As String = "IsKDIL"
    Public Const SendToTally As String = "SendToTally"
    Public Const PromptForTally As String = "PromptForTally"
    Public Const CurrentMaufacturingType As String = "ManufacturingType"
    Public Const TallyCompany As String = "TallyCompany"
    Public Const TallyIP As String = "TallyIP"
    Public Const TallyPort As String = "TallyPort"
    Public Const TaxRoundOffToZeroDecimalPlace As String = "TaxRoundOffToZeroDecimalPlace"
    Public Const BalanceSheetProftAndLossGroupCode As String = "BalanceSheetProftAndLossGroupCode"
    Public Const BalanceSheetProftAndLossGroupDesc As String = "BalanceSheetProftAndLossGroupDesc"
    Public Const ApplyCostingOnPostedDate As String = "ApplyCostingOnPostedDate"
    Public Const isBatchApplyOnInventoryMovement As String = "isBatchApplyOnInventoryMovement"
    Public Const BlankDatabase As String = "BlankDatabase"
    Public Const ServiceDealer As String = "Service Dealer"
    Public Const TDM As String = "TDM"
    Public Const MAILOFF As String = "MAILOFF"
    Public Const AllowToSaveTimeWithDocumentDate As String = "Allow To Save Time With Document Date"
    Public Const AllowToPrintTimeWithDocumentDate As String = "Allow To Print Time With Document Date"
    Public Const AllLevelApprovalIsMandatory As String = "All Level Approval Is Mandatory"
    Public Const AssetGroupPrefix As String = "AssetGroupPrefix"
    Public Const DepreciationCalculationMethod As String = "Depreciation Calculation Method"
    Public Const STDPURRATE As String = "STDPURRATE"
    Public Const AutoPOAtSRN As String = "AUTOPOATSRN"
    Public Const DisableShipToLocation As String = "Disable Ship_To_Location For (PO,PI,SRN)"
    Public Const PurchasePickItemFromVendorItemDetails As String = "PurchasePickItemFromVendorItemDetails"
    Public Const PurchaseOneItemOneVendor As String = "PurchaseOneItemOneVendor"
    Public Const AllowLargerItemCostThenVendorItemCost As String = "AllowLargerItemCostThenVendorItemCost"
    Public Const AcceptQtyWithoutConsiderRejectQty As String = "Accept Qty Without Consider Reject Qty"
    Public Const ShowGRN As String = "ShowGRN"
    Public Const SkipMRNGRNinCaseofMT As String = "SkipMRNGRNinCaseofMT"
    Public Const AutoGenerateMRN As String = "AutoGenerateMRN"
    Public Const ShowMRN As String = "ShowMRN"
    Public Const EnableProjectFinder As String = "EnableProjectFinder"
    Public Const PostShipmentonAutoSTN As String = "PostShipmentonAutoSTN"
    Public Const IsRemarksMandatoryOnCloseSaleOrder As String = "IsRemarksMandatoryOnCloseSaleOrder"
    Public Const CreateInvoicewithShipmentonAutoSTN As String = "CreateInvoicewithShipmentonAutoSTN"
    Public Const AllowSingleInvoiceAgainstSingleOrder As String = "AllowSingleInvoiceAgainstSingleOrder"
    Public Const WorkingHours As String = "WorkingHours"
    Public Const TreatExcessLeaveAbsent As String = "TreatExcessLeaveAbsent"
    Public Const VehicleInsuranceAlert As String = "VehicleInsuranceAlert"
    Public Const IsItemRateEditableOnTransfer As String = "IsItemRateEditableOnTransfer"
    Public Const GLACAccordingToTaxRate As String = "GLACAccordingToTaxRate"
    Public Const AutoSchemeOn As String = "AutoSchemeOn"
    Public Const MultiplySubsidyWithQuantity As String = "Multiply Subsidy With Quantity"
    Public Const CreateAutoReceiptEntryDCSSale As String = "CreateAutoReceiptEntryDCSSale"
    Public Const GeneratePDFForMobile As String = "Generate PDF For Mobile"
    Public Const IsTransferQtyEditableOnAutoSTN As String = "IsTransferQtyEditableOnAutoSTN"
    Public Const IsItemRateEditableOnSales As String = "IsItemRateEditableOnSales"
    Public Const IsItemMRPEditableOnSales As String = "IsItemMRPEditableOnSales"
    Public Const IsItemRateEditableOnSalesForAprilOnly As String = "ForAprilOnly"
    Public Const Mediclaim As String = "Mediclaim"
    Public Const LTA As String = "LTA"
    Public Const Gratuity As String = "Gratuity"
    Public Const LeaveEnchashed As String = "LeaveEnchashed"
    Public Const UnpaidAmount As String = "UnpaidAmount"
    Public Const Arrear As String = "Arrear"
    Public Const PT As String = "PT"
    Public Const UserPWD As String = "UserPWD"
    Public Const AllowMilkReceiptAfterSettingsisOn As String = "AllowMilkReceiptAfterSettingsisOn"
    Public Const BMCMilkQCManualPassword As String = "BMCMilkQCManualPassword"
    Public Const MilkReceiptTolerancePwd As String = "MilkReceiptTolerancePwd"
    Public Const MCCDLTPWD As String = "MCCDLTPWD"
    Public Const Allow_ExcelCode_On_Mcc As String = "Allow_ExcelCode_On_Mcc"
    Public Const Is_Allow_Cancel_Transaction As String = "Is_Allow_Cancel_Transaction"
    Public Const is_Allow_cancel_Posted_Transaction As String = "is_Allow_cancel_Posted_Transaction"
    Public Const ShiftTiming As String = "ShiftTiming"
    Public Const GetMulitcurrencyDecimalPlaces As String = "GetMulitcurrencyDecimalPlaces"
    Public Const MilkSetting As String = "MilkSetting"
    Public Const ShowTaxRateColumnOnTransaction As String = "ShowTaxRateColumnOnTransaction"
    Public Const LicenceExpiryDate As String = "IsApplyCommonService1" 'A B
    Public Const LicenceNoOfExeConnection As String = "IsApplyCommonService2" 'C
    Public Const LicenceNoOfJournalEntry As String = "IsApplyCommonService3" 'D
    Public Const LicenceNoOfUser As String = "IsApplyCommonService4" 'E
    Public Const InvoiceManualNoWithPrefix As String = "InvoiceManualNoWithPrefix"
    Public Const AutoBackUp As String = "AutoBackUp"
    Public Const MCCPurchase As String = "MCCPurchase"
    Public Const BulkSaleDefaultMilkItem As String = "BulkSaleDefaultMilkItem"
    Public Const BSDefaultMilkItem As String = "BSDefaultMilkItem"
    Public Const DefaultRoundOffGLAccount As String = "DefaultRoundOffGLAccount"
    Public Const NotificationSettingforReOrderInPO As String = "NotificationSettingforReOrderInPO"
    Public Const NotificationSettingforReOrderInPurchaseRequisition As String = "NotificationSettingforReOrderInPurchaseRequisition"
    Public Const PurchaseOrderAutomaticallyItemQtyBelowReorderLevel As String = "PurchaseOrderAutomaticallyItemQtyBelowReorderLevel"
    Public Const NLevelAtVendor As String = "NLevel_Vendor"
    Public Const NLevelAtCustomer As String = "NLevel_Customer"
    Public Const NLevelAtLocation As String = "NLevel_Location"
    Public Const AutoItemNLevel As String = "NLevel_ItemCode"
    Public Const CounterRawMaterial As String = "R"
    Public Const CounterFinishGood As String = "F"
    Public Const CounterSemiFinishGood As String = "S"
    Public Const CounterTradingGood As String = "T"
    Public Const CounterAsset As String = "A"
    Public Const CounterOther As String = "O"
    Public Const Princi_Bom As String = "Principle_BOM"
    Public Const AP_INV_COMMSN As String = "AP_INV_COMMSN"
    Public Const Principal_Vendor As String = "Principal_Vendor"
    Public Const Principal_Vendor_Database As String = "Principal_Vendor_Database"
    Public Const Principal_Customer As String = "Principal_Customer"
    'Public Const ExeExpiredDate As String = "ExpiredDate"
    Public Const CalculateLTAOnHoliday As String = "CalculateLTAOnHoliday"
    Public Const CalculateLTAOnWeekend As String = "CalculateLTAOnWeekend"
    Public Const CalculateMediclaimOnHoliday As String = "CalculateMediclaimOnHoliday"
    Public Const CalculateMediclaimOnWeekend As String = "CalculateMediclaimOnWeekend"
    Public Const Is_Purchaseable_Item As String = "Is_Purchaseable_Item"
    Public Const Is_AbemdmentForDemo As String = "Is_AbemdmentForDemo"
    Public Const Is_FinishedGoods As String = "Is_FinishedGoods"
    Public Const ShowStatusForPurchase As String = "ShowStatusForPurchase"
    Public Const ShowStatusForSales As String = "ShowStatusForSales"
    Public Const ShowSerialNoForSales As String = "ShowSerialNoForSales"
    Public Const AutoGeneratedVendorCode As String = "AutoGeneratedVendorCode"
    Public Const AutoGeneratedVendorCodeForAllCompany As String = "AutoGeneratedVendorCodeForAllCompany"
    Public Const AutoGeneratedCustomerCode As String = "AutoGeneratedCustomerCode"
    Public Const AutoGeneratedCustomerCodeForAllCompany As String = "AutoGeneratedCustomerCodeeForAllCompany"
    Public Const ApplyBrachAccounting As String = "ApplyBrachAccounting"
    Public Const AllowToUseSubAccount As String = "AllowToUseSubAccount"
    Public Const InTransitFeatureIsRequired As String = "InTransitFeatureIsRequired"
    Public Const PermissionSettingForTransactionWithBank As String = "Permission_Setting_For_Trans_With_Bank"
    Public Const AllowToEnterMRPManually As String = "AllowToEnterMRPManually"
    Public Const AllowFieldsToBeManadatory As String = "AllowFieldsToBeManadatory"
    Public Const AutoGeneratedDigitsForVendor As String = "AutoGeneratedDigitsForVendor"
    Public Const AutoGeneratedDigitsForCustomer As String = "AutoGeneratedDigitsForCustomer"
    Public Const IsRateEditableOnSRN As String = "IsRateEditableOnSRN"
    Public Const DisAllowIntermittentTankerForPlantDispatch As String = "DisAllowIntermittentTankerForPlantDispatch"
    Public Const CreateGLAccToItem As String = "CreateGLAccToItem"
    Public Const IsCostEditableOnIssueReturnTransfer As String = "IsCostEditableOnIssueReturnTransfer"
    Public Const UpdateCrateLinerQty As String = "UpdateCrateLinerQty"
    Public Const AllowDispatchOutstandingBS As String = "AllowDispatchOutstandingBS"
    Public Const AllowDispatchOutstandingFS As String = "AllowDispatchOutstandingFS"
    Public Const AllowDispatchOutstandingPS As String = "AllowDispatchOutstandingPS"
    Public Const IsVolumeSchemeBydefault As String = "IsVolumeSchemeBydefault"
    Public Const DiscountCodeForArAdj As String = "DiscountCodeForArAdj"
    Public Const DiscountCodeForMPAdj As String = "DiscountCodeForMPAdj"
    Public Const CreateAutoRecieptForManualCustomer As String = "CreateAutoRecieptForManualCustomer"
    Public Const AutoRecieptBankCode As String = "AutoRecieptBankCode"
    Public Const AutoRecieptPaymentMode As String = "AutoRecieptPaymentMode"
    Public Const AllowtoSelDateandBankforPayEntryOnSalaryGeneration As String = "AllowtoSelDateandBankforPayEntryOnSalaryGeneration"
    Public Const AllowDeliveryOrderIncaseAmountIncreases As String = "AllowDeliveryOrderIncaseAmountIncreases"
    Public Const AllowAutoMRNGRNonDocumentAcceptance As String = "AllowAutoMRNGRNonDocumentAcceptance"
    Public Const AllowToShowSaleTypeinPaymentTermsReceivable As String = "AllowToShowSaleTypeinPaymentTermsReceivable"
    Public Const AllowToShowMilkTypeinAdjustmentEntry As String = "AllowToShowMilkTypeinAdjustmentEntry"
    Public Const GatePassAfterTransfer As String = "GatePassAfterTransfer"
    Public Const CreateTransferFromBooking As String = "CreateTransferFromBooking"
    Public Const PickRateFromPRICEChrtMasterFORUMang As String = "PickRateFromPRICEChrtMasterFORUMang"
    Public Const StockTranferFromTransferPriceAndInvJVWithAvgCost As String = "StockTranfer - TransferPriceAndInvJVWithAvgCost"
    Public Const IGnoreGITAccount As String = "Ignore GIT Account in Financial Entry"
    Public Const AllowToEditCategoryCodeinItemMaster As String = "AllowToEditCategoryCodeinItemMaster"
    Public Const CreditLimitApproval As String = "CustomerCreditLimit"
    Public Const ViewTDSPwd As String = "ViewTDSPwd"
    Public Const ShowVendorLedgerasPerRightsForLocation As String = "ShowVendorLedgerasPerRightsForLocation"
    Public Const ShowCustomerLedgerasPerRightsForLocation As String = "ShowCustomerLedgerasPerRightsForLocation"
    Public Const EnableDynamicQRCodeForB2CInvoice As String = "EnableDynamicQRCodeForB2CInvoice"
    Public Const InvoiceBasedPO As String = "InvoiceBasedPO"
    Public Const AdvanceAgainstSO As String = "AdvanceAgainstSO"
    Public Const Purchase_SMSATPOST As String = "SMSATPOST_PUR"
    Public Const Sale_SMSATPOST As String = "SMSATPOST_SALE"
    Public Const showRFQ As String = "showRFQ"
    Public Const AmountLimitForInvoiceBulkSale As String = "AmountLimitForInvoiceBulkSale"
    Public Const ShowSaleInvoiceNoInPOfinderInSRN As String = "ShowSaleInvoiceNoInPOfinderInSRN"
    Public Const CrateValue As String = "CrateValue"
    Public Const CommitedDefaultQty As String = "CommitedDefaultQty"
    Public Const ShowBinMapping As String = "ShowBinMapping"
    Public Const ShowPrintChallanInDairyDispatch As String = "ShowPrintChallanInDairyDispatch"
    Public Const FATKGSNFKGRoundOff As String = "FAT KG & SNF KG Round Off"
    Public Const ShowOwnBMCDCS As String = "Show Own BMC DCS"
    Public Const ShowCrateJaaliBoxIntransfer As String = "Show Crate Jaali & Box In transfer"
    Public Const DefaultCorrectionFactorForBulkSale As String = "DefaultCorrectionFactorForBulkSale"
    Public Const MCCdefaultCorrectionFactorBS As String = "MCCdefaultCorrectionFactorBS"
    Public Const JOBdefaultCorrectionFactorBS As String = "JOBdefaultCorrectionFactorBS"
    Public Const PurchasedefaultCorrectionFactorBS As String = "PurchasedefaultCorrectionFactorBS"
    Public Const AllowDeliveryQtygreaterthanBookingQtyPS As String = "AllowDeliveryQtygreaterthanBookingQtyPS"
    Public Const IsPickServerDateForMultipleDispatchInvoice As String = "IsPickServerDateForMultipleDispatchInvoice"
    Public Const AutoTabOrdering As String = "AutoTabOrdering"
    Public Const AutoTabOrderingPattern As String = "AutoTabOrderingPattern"
    Public Const IsItemEditableOnMCCDispatch As String = "IsItemEditableOnMCCDispatch"
    Public Const IsUOMSelectableOnMCCDispatch As String = "IsUOMSelectableOnMCCDispatch"
    Public Const LoadLoginScreenDirectlyAfterStartup As String = "LoadLoginScreenDirectlyAfterStartup"
    Public Const IsItemWithDifferntUnitConsiderAsOtherItem As String = "ItemWithDifferntUnitConsiderAsOtherItem"
    Public Const AutoSetTabStopFalseForReadonlyControls As String = "AutoSetTabStopFalseForReadonlyControls"
    Public Const AutoRestoreGridLayout As String = "AutoRestoreGridLayout"
    Public Const IsMRPWiseBalance As String = "IsMRPWiseBalance"
    Public Const CreateDbitNoteForShortPI As String = "CreateDbitNoteForLeakAndShortPI"
    Public Const CreateDbitNoteForRejectPI As String = "CreateDbitNoteForRejectPI"
    Public Const CreateDebitNoteForUnitCost As String = "CreateDebitNoteForUnitCost"
    Public Const TransferJEForLocationMapping As String = "TransferJEForLocationMapping"
    Public Const TransferWithProductionSale_Retail_Series As String = "CreateTransferWithProductionSale_Retail_Series"
    Public Const ProductionQtyDecimalPoint As String = "ProductionQtyDecimalPoint"
    Public Const ConsumptionQtyTollerance As String = "Consumption Qty Tollerance"
    Public Const GunnyBagTollerance As String = "Gunny Bag Tollerance"
    Public Const RCDFCFProductionEntry As String = "RCDFCF Production Entry"
    Public Const RCDFCFDispatch As String = "RCDFCF Dispatch"
    Public Const ProductionFATSNFPerDecimalPoint As String = "ProductionFATSNFPerDecimalPoint"
    Public Const ManualySelectBOMForChildBatch As String = "ManualySelectBOMForChildBatch"
    Public Const AllowToDispalyAlertForBDayAnniversary As String = "AllowToDispalyAlertForBDayAnniversary"
    Public Const AllowToSendEmailForBDayAnniversary As String = "AllowToSendEmailForBDayAnniversary"
    Public Const ItemDescForTankerDispatchPrint As String = "ItemDescForTankerDispatchPrint"
    Public Const AllowPOScheduling As String = "Allow PO Scheduling"
    Public Const AllowGateEntryInPrevDate As String = "AllowGateEntryInPrevDate"
    Public Const ERPStartDate As String = "ERPStartDate"
    Public Const ItemBatchWiseStartDate As String = "ItemBatchWiseStartDate"
    Public Const AllowQcDateBeforeGateEntryDateTime As String = "AllowQcDateBeforeGateEntryDateTime"
    Public Const CreateJEForTransfer As String = "CreateJEForTransfer"
    Public Const AllowToSkipStageQLLogSheetInProd As String = "AllowToSkipStageQLLogSheetInProd"
    Public Const IsRemarkReasonMandatoryOnPO As String = "IsRemarkReasonMandatoryOnPO"
    Public Const ShowCostCenterAndHierarchyLevelInPurchaseModule As String = "ShowCostCenterAndHierarchyLevelInPurchaseModule"
    Public Const IsQCColumnRequiredonMRN As String = "IsQCColumnRequiredonMRN"
    Public Const ShowItemAllStructureWise As String = "ShowItemAllStructureWise"
    Public Const ShowUpto6DecimalPlaces As String = "ShowUpto6DecimalPlaces"
    Public Const EnableOnPrivateChkbox As String = "EnableOnPrivateChkbox"
    Public Const RateDecimalPlaces As String = "Rate Decimal Places"
    Public Const AllowQcDateAfterCurrentDate As String = "AllowQcDateAfterCurrentDate"
    Public Const AllowWeighmentDateAfterCurrentDate As String = "AllowWeighmentDateAfterCurrentDate"
    Public Const AllowUnloadingDateAfterCurrentDate As String = "AllowUnloadingDateAfterCurrentDate"
    Public Const AllowcleaningDateAfterCurrentDate As String = "AllowcleaningDateAfterCurrentDate"
    Public Const AllowGateOutDateAfterCurrentDate As String = "AllowGateOutDateAfterCurrentDate"
    Public Const AllowSRNDateAfterCurrentDate As String = "AllowSRNDateAfterCurrentDate"
    Public Const IsRGPAfterPurchaseOrder As String = "Do RGP After Purchase Order"
    Public Const AllowQualityModuleInERP As String = "On Quality Module"
    Public Const SRNReportQuantityWise As String = "SRNReportQuantityWise"
    Public Const IsCustomerGroupFieldsMandatory As String = "IsCustomerGroupFieldsMandatory"
    Public Const IsVendorGroupFieldsMandatory As String = "IsVendorGroupFieldsMandatory"
    Public Const AllowAutoNoForBackLogEntry As String = "AllowAutoNoForBackLogEntry"
    Public Const AllowDiffentSeriesExemptedItemONPS As String = "AllowDiffentSeriesExemptedItemONPS"
    Public Const DisplayFranchiseeinCustomer As String = "DisplayFranchiseeinCustomer"
    Public Const isIdleTimerOn As String = "isIdleTimerOn"
    Public Const idleTime As String = "idleTime"
    Public Const AddressOnPaymentVoucherOnBankBasis As String = "AddressOnPaymentVoucherOnBankBasis"
    'richa agarwal 17/03/2015 against ticket no BM00000005874
    Public Const AllowBankDetailsManualinVM As String = "AllowBankDetailsManualinVM"
    ''--------------------------------
    ''RICHA AGARWAL 17/03/2015
    Public Const AllowToGenerateSaleInvoiceSeriesTaxTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesTaxTypeatPS"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesRetailTypeatPS"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseTypeatPS As String = "AllowToGenerateSaleInvoiceSeriesExciseTypeatPS"
    ''-------------------------
    ''RICHA AGARWAL 17/03/2015 MCC Sale
    Public Const AllowToGenerateSaleInvoiceSeriesTaxatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesTaxatMCCSale"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesRetailatMCCSale"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseatMCCSale As String = "AllowToGenerateSaleInvoiceSeriesExciseatMCCSale"
    ''-------------------------
    ''RICHA AGARWAL 17/03/2015 Misc Sale
    Public Const AllowToGenerateSaleInvoiceSeriesTaxatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesTaxatMiscSale"
    Public Const AllowToGenerateSaleInvoiceSeriesRetailatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesRetailatMiscSale"
    Public Const AllowToGenerateSaleInvoiceSeriesExciseatMiscSale As String = "AllowToGenerateSaleInvoiceSeriesExciseatMiscSale"
    ''-------------------------
    '=========================Preeti Gupta===========================
    Public Const ShowHierarchyAndCostCenterInAPInvoiceEntry As String = "ShowHierarchyAndCostCenterInAP"
    Public Const WeighmentNotMandatoryInMCC As String = "WeighmentNotMandatoryInMCC"
    '=================================================================
    Public Const ShowHierarchyAndCostCenterInARInvoiceEntry As String = "ShowHierarchyAndCostCenterInAR"
    Public Const PartialFADepDays As String = "PartialFADepDays"
    Public Const RateMultPartialFADepDays As String = "RateMultPartialFADepDays"
    Public Const AllowNegativeStock As String = "AllowNegativeStock"
    Public Const SendSalarySlipMailToEmployee As String = "SendSalarySlipMailToEmployee"
    Public Const DoNotMergeAPARAccount As String = "DoNotMergeAPARAccount"
    Public Const ShowVisiDetail As String = "ShowVisiDetail"
    Public Const CustomerNameUniqueOnCM As String = "CustomerNameUniqueOnCM"
    Public Const IsShortageIncludeInLandedCost As String = "IsShortageIncludeInLandedCost"
    Public Const AlowwdateChangeinPaymentEntry As String = "AlowwdateChangeinPaymentEntry"
    Public Const CreateAutoMilkRGPinBulkSRN As String = "CreateAutoMilkRGPinBulkSRN"
    Public Const DisplayAllParameterinQualityCheck As String = "DisplayAllParameterinQualityCheck"
    Public Const DisplayTypeInMilkReceipt As String = "DisplayTypeInMilkReceipt"
    '============Added by Rohit on Aug 03,2015 For Milk Type Validation in Milk sample.============
    Public Const AddValidationofMilkTypeinsample As String = "AddValidationofMilkTypeinsample"
    Public Const SubStdFatCow As String = "Sub Std FAT Per Cow"
    Public Const SubStdSNFCow As String = "Sub Std SNF Per Cow"
    Public Const SubStdFatBuff As String = "Sub Std FAT Per Buff"
    Public Const SubStdSNFBuff As String = "Sub Std SNF Per Buff"
    Public Const SubStdFatMix As String = "Sub Std FAT Per Mix"
    Public Const SubStdSNFMix As String = "Sub Std SNF Per Mix"
    Public Const FatMinCow As String = "FatMinCow"
    Public Const FatMaxCow As String = "FatMaxCow"
    Public Const SNFMinCow As String = "SNFMinCow"
    Public Const SNFMaxCow As String = "SNFMaxCow"
    Public Const FatMinBuff As String = "FatMinBuff"
    Public Const FatMaxBuff As String = "FatMaxBuff"
    Public Const SNFMinBuff As String = "SNFMinBuff"
    Public Const SNFMaxBuff As String = "SNFMaxBuff"
    Public Const FatMinMix As String = "FatMinMix"
    Public Const FatMaxMix As String = "FatMaxMix"
    Public Const SNFMinMix As String = "SNFMinMix"
    Public Const SNFMaxMix As String = "SNFMaxMix"
    '================================================================================================
    Public Const AddIncentiveDeductioninMilkSample As String = "AddIncentiveDeductioninMilkSample"
    Public Const AllowManualEnterinWeighment As String = "AllowManualEnterinWeighment"
    Public Const SettlementBankOnlyPWD As String = "SettlementBankOnlyPWD"
    Public Const DocumentSequence As String = "DocumentSequence"
    Public Const AllowPurchaseAccounting As String = "AllowPurchaseAccounting"
    Public Const SHowBulkMilkWeighment As String = "SHowBulkMilkWeighment"
    Public Const StoreADJExportImportAfterPost As String = "StoreADJExportImportAfterPost"
    Public Const FatSNFControlOnProductionConsumption As String = "FatSNFControlOnProductionConsumption"
    Public Const QuantityControlToleranceOnProductionConsumption As String = "QuantityControlToleranceOnProductionConsumption"
    Public Const LeaveBalanceAlertTypeOnAttendance As String = "LeaveBalanceAlertTypeOnAttendance"
    Public Const StopNegativeBankBalance As String = "StopNegativeBankBalance"
    Public Const ConsumptionTypeMilk As String = "ConsumptionTypeMilk"
    Public Const ConsumptionTypeMilkStandardization As String = "ConsumptionTypeMilkSTD"
    Public Const ConsumptionTypeMilkProduct As String = "ConsumptionTypeMilkProduct"
    Public Const ConsumptionTypeOther As String = "ConsumptionTypeOther"
    Public Const ValidateFatSnfOnProduction As String = "ValidateFatSnfOnProduction"
    Public Const ShowOverheadCostOnProductionEntry As String = "ShowOverheadCostOnProductionEntry"
    Public Const ActivateProductionWithoutBatch As String = "ActivateProductionWithoutBatch"
    Public Const CreateJEOnProduction As String = "CreateJEOnProduction"
    Public Const AllowToSaveMultipleEmployeeStatus As String = "AllowToSaveMultipleEmployeeStatus"
    Public Const CreateJEForProvisionEntrySecondaryTransporter As String = "Secondary Transporter"
    Public Const CreateJEForProvisionEntryMCCLeaseVendor As String = "MCC Lease Vendor"
    Public Const CreateJEForProvisionEntryTransporterForFreshSale As String = "Transporter For Fresh Sale"
    Public Const CreateJEForProvisionEntryTransporterForProductSale As String = "Transporter For Product Sale"
    Public Const CreateJEForProvisionEntryTransporterForBulkSale As String = "Transporter For Bulk Sale"
    Public Const CreateJEForProvisionEntryOthers As String = "Others"
    Public Const CreateJEForProvisionEntryPrimaryTransporter As String = "Primary Transporter"
    Public Const CreateJEForProvisionEntryTransporterForTransfer As String = "Transporter For Transfer"
    Public Const CreateJEForProvisionEntryTransporterForGateentry As String = "Transporter For Gate Entry"
    Public Const CreateJEForProvisionEntryTransporterForDairySale As String = "Transpoter"
    Public Const CreateJEForProvisionEntryTransporterForCSATransfer As String = "Transporter For CSA Transfer"
    Public Const DoubleClickOnVC As String = "Double Click On VC"
    Public Const PickManual_CSATransfer_OnTRansferReturn As String = "CSA Transfer Effect on Return is Manual"
    Public Const PickManual_CSATransfer_OnCSASalePatti As String = "CSA Transfer Effect on Sale Patti is Manual"
    Public Const AllowDistributorSaleAtCSA_SaleInvoice As String = "Allow Distributor Sale at CSA Sale Patti"
    Public Const AllowItemWiseCSAAccountingON_CSASale As String = "CSA Account set pick Item-wise"
    Public Const IsAutoTankerWeightment As String = "Auto Tanker Weightment"
    Public Const IsAutoTankerWeighmentForBulkSale As String = "Auto Tanker Weighment for Bulk Sale"
    Public Const IsAdditionalInformationOnVillageMaster As String = "Show Village Add Info"
    Public Const CheckLiveStockInProductionDuringTrans As String = "CheckLiveStockInProductionDuringTrans"
    Public Const VLCTimeTableColumnShow As String = "VLCTimeTableColumnShow"
    Public Const VLCTimeTableColumnMandatory As String = "VLCTimeTableColumnMandatory"
    Public Const ApplyEffectiveStartDate As String = "Apply Effective Start Date"
    Public Const isOneMCCOnePrimaryTranporter As String = "One MCC One Primary Tranporter"
    Public Const MilkSamplShowOddEvenTwoGrid As String = "Show Odd and Even Two Grid"
    Public Const OpenODDEvenForm As String = "Open Odd-Even Form"
    Public Const Open4AnalyzerForm As String = "Open 4 Milk Analyzer Form"
    Public Const IsApplyEMIOnAssetValue As String = "Is Apply EMI On Asset Value"
    ' KUNAL 6-SEP-2016 ======================================================================
    Public Const AllowFutureDateTransaction As String = "AllowFutureDateTransaction"
    'KUNAL > UDIL > DATE : 16-NOV-2016
    Public Const FindNRGP_Request As String = "Show_NRGP_RequestNo"
    '========================================================================================
    Public Const AllowCSAPriceMasterPostedData As String = "Allow CSAPriceMaster Posted Data"
    Public Const AllowItemMasterPostedData As String = "Allow Item Master Posted Data"
    Public Const AllowMilkItemMasterPostedData As String = "Allow Milk Item Master Posted Data"
    Public Const AllowBulkProcItemPostedData As String = "Allow Bulk Proc Milk Item Posted Data"
    Public Const AllowPriceListMasterPostedData As String = "Allow Price List Item Posted Data"
    'Stuti
    Public Const ItemCrateWtinKg As String = "Item Default Crate Wt.(Kg.)"
    Public Const ItemJaaliWtinKg As String = "Item Default Jaali Wt.(Kg.)"
    Public Const ItemBoxWtinKg As String = "Item Default Box Wt.(Kg.)"
    Public Const ItemCrateRate As String = "Item Default Crate Rate"
    Public Const ItemJaaliRate As String = "Item Default Jaali Rate"
    Public Const ItemBoxRate As String = "Item Default Box Rate"
    Public Const ItemCanRate As String = "Item Default Can Rate"
    Public Const CustomerMasterFinderOnLocationwiseARReceipt As String = "Customer master finder location-wise on AR Receipt"
    Public Const SameuserCanNotloginmultipletimes As String = "Same user can-not login multiple times"
    Public Const MandatoryEmployeeOnVehicleMaster = "Make employee no mandatory"
    Public Const PlantDepotMappingMandatory = "Map location of plant with depot is mandatory"
    Public Const ShowCancelButtonPO As String = "Show cancel button on purchase order"
    Public Const ShowOptionforSelectingCapex As String = "Show option for selecting capex code/subcode on PO"
    Public Const AutoClosePO As String = "Auto close PO when all qty. received."
    Public Const POCancel As String = "PO Cancel"
    'Public Const CreateJVForAllCasesinRGP = "Crate JV for all cases in RGP"
    Public Const StoreRequisitionMandatoryforstorerequest = "Store Requisition mandatory for store request"
    Public Const AllowThreeFormatByDefaultForPrint = "Allow printing 3 formats by default"
    Public Const MTCapacityRequired = "MT Capacity Required"
    Public Const AllowBackDateEntry As String = "Allow back date entry for given days"
    Public Const BackDateEntryPwd As String = "BackDateEntryPwd"
    Public Const RevisedBudget As String = "Revised Budget"
    Public Const DipMarkingMendatory As String = "Make dip marking mendatory."
    Public Const AllowDispatchChecklistOnProductDispatch As String = "Allow dispatch checklist on product dispatch"
    Public Const ShowIndentBasedOnCreatedUser As String = "Show indent based on created user"
    Public Const ShowSystemStockinOpenMCC As String = "Show system stock in open MCC shift"
    Public Const Tankerfromtankersalemasteringateentry As String = "Tanker from tanker sale master in gate entry"
    Public Const ApplyMultiChamberInBulkWeighmentEntry As String = "Apply multi-chamber in bulk weighment entry"
    Public Const DefaultItemUOMForBulkSale As String = "Default item uom for bulk sale"
    Public Const InsuranceNoAndSealNoInBulkDispatch As String = "Show option for entering Insurance No and Seal No"
    Public Const ValidateFatSNFOnJobMilkSRN As String = "Validate FAT KG & SNF KG on Job Milk SRN"
    Public Const CancelDocDueToSRNReturn As String = "Cancel document due to SRN Return"
    Public Const AmountInLacsOnMisSaleRegister As String = "Allow amount in lacs on MIS SALE REGISTER"
    Public Const ShortCloseItemWiseOnPO As String = "Allow short close item wise on PO"
    Public Const MakeClosingofPOreadonlyforuser As String = "Make closing of PO read only"
    Public Const MakeClosingofCustomerTenderreadonlyforuser As String = "Make closing of customer tender read only"
    Public Const MakeClosingofCustomerTenderOrderreadonlyforuser As String = "Make closing of customer tender order read only"
    Public Const AllowModificationOnApprovalByApprovalUser As String = "Allow Modification On Approval By Approval User"
    Public Const AllowAutoCalculateADDREMOVEQty As String = "Auto Calculate Qty of Add/Remove Item"
    '-----------------end here---------------'
    Public Const FATDeductionPercent As String = "FAT Deduction Percent"
    Public Const SNFDeductionPercent As String = "SNF Deduction Percent"
    Public Const RejectionReturnPaneltyPerUnit As String = "Rejection Return Penelty Per Unit"
    Public Const RejectionDrainPenaltyPerUnit As String = "Rejection Drain Penelty Per Unit"
    Public Const RejectionCOBPenaltyPerUnit As String = "Rejection COB Penalty Per Unit"
    Public Const GraceTimeForTransporter As String = "Grace Time For Transporter"
    Public Const GraceTimeFromGateEntryToDocWeighing As String = "Grace Time From Gate Entry To Dock Weighing"
    ''==============end here================
    ''============CSA Sale Setting=====================================================================
    Public Const ShowCSAReturnTypeOnScreen As String = "Show CSA Return Type on screen"
    Public Const ShowCSARequestScreen As String = "Enable CSA Request Instead of Booking"
    Public Const AllowSchemeOnCSADeliveryOrder As String = "Allow Scheme at CSA DO Entry"
    Public Const AllowOtherItemOnCSAPriceMaster As String = "Allow Other Items On CSA Price Master"
    Public Const AllowRoundOff_OnCSASalePatti As String = "Inv. Amount Round-off on All Sale Invoice"
    Public Const AllowRoundOffAmountOnDCSSale As String = "Allow Round Off Amount On DCSSale"
    Public Const FreightChargeOnCSASaleInvoice As String = "Comm./Freight itemwise on CSA Sale Invoice"
    Public Const AllowDisabledCommissionOnCSATransfer As String = "Commission disabled on CSA Transfer"
    Public Const DoReadonly_UnitRate_AtCSASale As String = "Allow Rate readonly on CSA Sale"
    Public Const Allow_SaleMfgACONCSAPatti As String = "Allow Sale mfg. A/c on CSA Sale Patti"
    Public Const AllowSchemeItemCondONSchemeMaster As String = "Allow Scheme type item on Scheme Master"
    Public Const ForUDLOnly As String = "CSA Sale changes For UDL only"
    Public Const CheckCreditLimitonCSADO As String = "Check Credit Limit on CSA DO"
    Public Const GrossWtFromItemMasterONCSATransfer As String = "Gross Wt. from item master on CSA Transfer"
    Public Const DisplayPriceCodeForBKN As String = "Display Price Code For BKN"
    Public Const EnableExciseONCSASalePatti As String = "Enable Excise entry on CSA Sale Patti"
    Public Const BatchSkipCSAReturn As String = "Batch Skip at CSA sale patti/Return"
    ''====================end here=====================================================
    Public Const IsChamberWiseTanker As String = "Chamber wise Tanker"
    'Prabhakar'
    Public Const AllowLoginTypeCNFdistributerRetailer As String = "Allow Login Type CNF , Distributer, Retailer"
    Public Const AllowSchemeItemQty As String = "Allow Scheme Item in Materix Report"
    Public Const AllowDairyDeliveryOrderPrint As String = "Allow Print Button for Delivery Order "
    Public Const ShowSealNumberForTunkerOut As String = "Show Seal Number for Tunker Out"
    Public Const HideRateDispatchCentreCode As String = "Hide Rate and Dispatch Centre Code"
    Public Const AllowPromptPendingDocs As String = "Allow Prompt Pending Docs"
    Public Const AllowAutoGenerateDocNoInMaster As String = "Allow Auto Generate Doc No In Master Screen"
    'kunal
    Public Const ShowDocsStatusFilters As String = "Show Documents Declaration Status Filters"
    Public Const AutoDepartmentMendatroryFieldOnPurcahseCycle As String = "Allow Department Mandatory On Purchase Cycle"
    Public Const AllowVehicleGateOutValidationScrapSale As String = "Allow Vehicle Gate Out Validation For Scrap Sale "
    Public Const AllowVehicleGateOutValidationCSATransfer As String = "Allow Vehicle Gate Out Validation For CSA Transfer"
    Public Const AllowVehicleGateOutValidationSPSale As String = "Allow Vehicle Gate Out Validation For SP Sale"
    Public Const AllowVehicleGateOutValidationTransfer As String = "Allow Vehicle Gate Out Validation For Transfer"
    Public Const AllowWithoutUnitCostIssueReturnEntry As String = "Allow without amount save Issue/Return Entry"
    Public Const ZeroCostForReprocess As String = "Zero Cost For Reprocess"
    Public Const IsAutoReceiptPayment As String = "IsAutoReceiptPayment"
    Public Const TransferEntryOnInvCtrlAccount As String = "Transfer Entry On Inventory Control Account"
    Public Const AutoUpdateVLCUploaderCodeInVLCMaster As String = "AutoUpdateVLCUploaderCodeInVLCMaster"
    Public Const StandardInterfaceForMilkShiftEnd As String = "StandardInterfaceForMilkShiftEnd"
    Public Const ShiftEndAllowManualEntryOfDeduction As String = "Allow Manual Entry Of Deduction"
    Public Const PTMRatePerLtrKGOnStdQty As String = "Rate Ltr/KG On Std Qty"
    'default bank payment
    Public Const DefaultBank = "Default Bank for Cash Payment"
    Public Const DefaultLocation = "Default Location for Cash Payment"
    'added by preeti gupta 03/10/2016==============
    Public Const ShowParticluarColumnInSalesRegisterForGopalJee As String = "Show Column in sale register report for GopalJee"
    ''Added by Nazia
    Public Const ShowPrintDiscountInDairyDispatchForGopaljee As String = "Show print discount in Dairy Dispatch"
    Public Const MilkReceiptRequiredApproval As String = "Milk Receipt Required Approval"
    Public Const LinkDepartmentBetweenIndentAndIssue As String = "Link Department Between Indent And Issue"
    Public Const CombineExportImportOnSchemeMaster As String = "Combined Export/Import on Scheme Master Dairy"
    Public Const OpenPOforRejectShortageQty As String = "Open PO for Reject/Shortage Qty"
    Public Const AutoSelectMCCRouteVLC As String = "Auto Select MCC Route VLC"
    Public Const PickServerDateWithNoChange As String = "Pick server Date With No Change"
    Public Const PickFinishedItemasBatchItem As String = "Finish Item as BatchItem default on Item Master"
    Public Const ToleranceFixFor_RM_OT_TRADE As String = "Tol.% mandatory for RM,Other,Trade on Item Master"
    Public Const ConsiderAdvancePayment As String = "Consider Advance Payment"
    Public Const PayableAmountZeroForMCCSale As String = "Payable Amount Zero For MCC Sale"
    Public Const Allow_AmountTruncate_BulkMilkSRN As String = "Allow truncate amount on Bulk Milk SRN"
    Public Const AutoPurchaseReturnFromIssueReturn As String = "Auto Purchase Return from Issue/Return screen."
    ''===Sanjeet====
    Public Const ShowAlternateVechileforFreshSale As String = "Gate pass with alternate vechile for fresh sale"
    Public Const CreateProvisionOfTransporterInDairyDispatch As String = "Create provision of transporter in Dairy Dispatch"
    Public Const SinglePrintCopyDairyInvoice As String = "Print Out Single Copy for Dairy Invoice"
    Public Const IncludeRatePerHoursIn As String = "Include Rate Per Hours In"
    Public Const GSTApplicable As String = "Allow GST Applicable"
    Public Const GSTApplicableDate As String = "Allow GST Applicable Date"
    Public Const AllowPanNoValidation As String = "Allow PAN No Validation"
    Public Const GSTActiveTaxesRatesGroup As String = "Show only Active Taxes/Rates/Groups for GST"
    Public Const AllowManualRejectionOfTanker As String = "Allow Manual Rejection Of Tanker"
    Public Const RunBulkProcOnAdjustedFATCLR As String = "Run Bulk Proc on adjusted FAT and CLR"
    Public Const BulkProcNetWeightCalculationWithVendorWeight As String = "Bulk Proc NetWeight Calculation by Vendor Weight"
    Public Const BulkProcPriceChartStandardRateWithZero As String = "Bulk Proc Price Chart standard rate with zero"
    Public Const RemoveForceAapprovalofBulkSRN As String = "Remove Force Approval of Bulk SRN"
    Public Const Allow_Plant_Depot_MCC_typeLocation As String = "Allow Plant Depot MCC type Location"
    Public Const ApplySaleRouteInDCSSale As String = "Apply Sale Route In DCS Sale"
    Public Const ValidateCustomerPANwithName As String = "Allow Validate Customer PAN with Name"
    Public Const ValidateTaxGroupForTransaction As String = "Allow Validate Tax Group Should Not Blank"
    Public Const AllowSeprateSchemeItemPrintDairySaleInvoice As String = "Allow Seprate Scheme Item Print DairySaleInvoice"
    Public Const EnableHirerachyCostCentre As String = "Enable Hirerachy Level Cost Centre"
    Public Const EnableStoreCostCentre As String = "Enable Store Cost Centre"
    Public Const EnableCostingMethod As String = "Enable Costing Method"
    Public Const CalculateItemCostonAvgForAssembly As String = "CalculateItemCostonAvgForAssembly"
    Public Const ShowAllCustomerOnMccMaterialSale As String = "Show All Customer On MCC Material Sale"
    Public Const ShowDefaultUser As String = "Show Default User"
    '(UDL)17/11/2016========
    Public Const ShowVatSeriesNoSeprately As String = "Allow Tax Tracking to Show Vat series No Seperatly"
    '(UDL)21/12/2016========
    Public Const AllowToGenerate_NEFTUPLOADER As String = "Allow Generate New NEFT UPLOADER File"
    Public Const DebitBankSelectWithNewFormateInNFTUploader As String = "Debit Bank Select With New Formate In NFT Uploader"
    '(UDL)05/01/2017
    Public Const AllowBulkPostingofAllDocuments As String = "Allow Bulk Posting of All Documents"
    '(UDL)10/01/2017
    Public Const AllowSameaAdditionalChargesMultiTime As String = "Allow Same Additaional Charges Multiple time"
    '(01/02/2017)
    Public Const AllowToSaveAndUpdatePasswordBased As String = "Allow Masters To Save and Update Pasword Based"
    Public Const AllowMasterModificationWithSecurity As String = "Allow Master Modification With Security"
    '(02/02/2017)
    Public Const ApplyRTGSAmtMoreThanGiven As String = "Apply RTGS Amount More Than Given"
    '====================
    Public Const GenerateSecondryCode As String = "Excise Secondary Series on Transfer"
    ''=====
    Public Const POWeighmentManual As String = "Mannual Weighment"
    ''======Ravi============
    Public Const AddTypeForUserMaster As String = "Add Type(Super User, Driver) in UserMaster"
    Public Const AddParavetEmployeeType As String = "Add Type Paravet in Employee Type"
    Public Const CalculateFIFOAndLIFOCosting As String = "Calculate FIFO And LIFO Costing"
    Public Const AllowDeductionPercentOnIncoming As String = "Allow Deduction(%) on incoming Quality"
    Public Const AllowLoginType As String = "Allow POS Functionality in ERP"
    Public Const MilkProcurementUploader As String = "Milk Procurement Uploader"
    Public Const TankerDispatchBulkUploader As String = "Bulk Tanker Uploader"
    Public Const EmptyCanWeight As String = "Empty Can Weight"
    Public Const MinuteInLastVehicleForGateEntry As String = "Minute Last Vehicle For Gate Entry"
    Public Const MinuteGateEntryToGrossWeight As String = "Minute Gate Entry To Gross Weight"
    Public Const MinuteGrossWeightToTareWeight As String = "Minute Gross Weight To Tare Weight"
    Public Const NoOfDaysForMultiInceForSameVSPForSamePayCycle As String = "NoOfDaysForMultiInceForSameVSPForSamePayCycle"
    Public Const NoOfDaysForShowFOSSDocument As String = "NoOfDaysForShowFOSSDocument"
    Public Const PurchaseCounterOnTransactionType As String = "Purchase Counter On Transaction Type"
    Public Const BulkProcurementCounterOnEntryType As String = "Bulk Procurement Counter On Entry Type"
    Public Const StopForRepeatedFATSNF As String = "Stop Repeat FAT SNF"
    Public Const SampleFONTSize As String = "Font Size"
    Public Const SMSPrefix As String = "SMS Prefix"
    Public Const InstantSendTheSMS As String = "Instant Send The SMS"
    Public Const InstantSendTheEmail As String = "Instant Send The Email"
    Public Const PickPendingMilkSRNinNextPaymentCycle As String = "Pick Pending Milk-SRN in Next Payment Cycle"
    '======================Preeti Gupta[29/12/2016]===========================
    Public Const TreatChequeClearDateAsRecoDate As String = "TreatChequeClearDateAsRecoDate"
    '========================END========================================
    '======================Preeti Gupta[29/12/2016]===========================
    Public Const BookWreckageFromSublocationOrSection As String = "BookWreckageFromSublocationOrSection"
    '========================END========================================
    Public Const StopVSPBillIfSomethingWrong As String = "Stop VSP Bill If Something Wrong"
    Public Const PDCSetting As String = "PDC Setting"
    Public Const AllowRoadPermitNo As String = "AllowRoadPermitNo"
    Public Const ShowMessgForTDS As String = "ShowMessgForTDS"
    Public Const IsShowTreeView As String = "IsShowTreeView"
    Public Const ShowVLCUploaderData As String = "Show VLC Uploader Data"
    '========================added parteek 09/01/2017
    Public Const FatSnfWhenMilktypeSelect As String = "Fat Snf persentage allow When Milk Type Select"
    Public Const DairyFreshTaxableandNonTaxable As String = "Taxable and Non-Taxable Item"
    Public Const SMSEMailPassword As String = "SMS EMail Password"
    Public Const CreateNewDocumentOnUploader As String = "Create New Document On Uploader"
    Public Const PopupJE As String = "Popup JE"
    'KUNAL > DATE : 23-01-2017 > CLIENT : Sahayog Dairy
    Public Const ShowAliasNames As String = "ShowAliasNames"
    'KUNAL > DATE : 23-01-2017 > CLIENT : Sahayog Dairy
    Public Const ShowFatAndSnfPercentageFields As String = "ShowFatNSNFPerc"
    'KUNAL > DATE : 23-01-2017 > CLIENT : Sahayog Dairy
    Public Const VehicleFitnessAndInsuranceFields As String = "VehicleFitnessFields"
    Public Const DocumentCancel As String = "Document Cancelation"
    Public Const DocumentCancelReturn As String = "Document Cancelation Return"
    Public Const PICancelUserPwd As String = "PI Cancel"
    Public Const CSADocumentCancel As String = "CSA Transfer Cancelation"
    Public Const FixVSPEMP As String = "Fix VSP EMP"
    Public Const FatSNFStockControl As String = "FatSNFStockControl"
    Public Const CheckBalanceFromInvMoveSummry As String = "CheckBalanceFromInvMoveSummry"
    Public Const ItemwiseFatSNFStockControl As String = "ItemwiseFatSNFStockControl"
    Public Const SepratePriceChartForCowMilk As String = "Seprate Price Chart For Cow Milk"
    Public Const SepratePriceChartForCamelMilk As String = "Seprate Price Chart For Camel Milk"
    '=======================Added by preeti gupta [20/02/207]=====================================
    Public Const AllowRoundInFixedAsset As String = "Allow Round In Fixed Asset"
    Public Const AllowDecimalInFixedAsset As String = "Allow Decimal In Fixed Asset"
    '==============================================================================================
    Public Const ApplyStdFATSNFRate As String = "Apply Standard FAT SNF Rate"
    Public Const OpenPriceChartPlanningScreenOnTotalSolid As String = "Open Price Chart Planning on Total Solid"
    Public Const AllowZeroQtyFATSNFInOpenMCCShift As String = "Allow Zero Qty FAT SNF In Open MCC Shift"
    Public Const AllowZeroQtyFATSNFInCloseMCCShift As String = "Allow Zero Qty FAT SNF In Close MCC Shift"
    ''============Parteek Added setting 03-03-2017
    Public Const POLimit As String = "POLimit"
    Public Const RequiredPOLimit As String = "RequiredPOLimit"
    Public Const UnitCostIncreasePurchaseInvoice As String = "UnitCostIncreasePurchaseInvoice"
    Public Const PromptMsgForPendingDocIntervel As String = "Prompt Messg For Pending Doc Intervel"
    Public Const UDLPurchaseOrderthroughAP As String = "UDL Purchase Order through AP invoice"
    Public Const UpdateInventorySummaryTable As String = "UpdateInventorySummaryTable"
    Public Const CreateConsumeEntry As String = "Create Consume Entry"
    Public Const ShowOptionforSelectingCapexForFA As String = "Showoptionforselectingcapexcode/subcodeonFA"
    Public Const UDLCapexAcquisionEntry As String = "UDL Capex for Acquision Entry"
    Public Const UDLRGPWiseDocument As String = "UDL RGP Wise Document Created"
    Public Const AllowAssetItemOnMiscSale As String = "Allow Asset Item on Misc. Sale"
    Public Const TriggerOfGLEntryForWinTable As String = "Trigger Of GL Entry For Win Table"
    'UDL DATE : 21-04-2017
    Public Const ShowRouteWiseAndVLCWiseReport As String = "ReportOfRouteAndVLCWise"
    Public Const UOMAtDiarySaleReturn As String = "Opt to change UOM At Diary SR on Actual UOM & Qty"
    Public Const PayableAmountZeroForFarmerPayment As String = "PayableAmountZeroForFarmerPayment"
    Public Const CheckDocAmountInAPInvoiceEntry As String = "Check Doc Amount For AP Invoice Entry"
    Public Const ApplyTSPriceAtBulkSale As String = "Apply TS Price At Bulk Sale"
    'UDL > DATE : 3-MAY-2017 : CHANGING DECLARED DOCUMENT LIST TO PENDING DOCUMENT LIST OR VICE VERSA
    Public Const ShowPendingDocumentsListScreenOverDeclaredDocumentList As String = "Show Pending Documents Screen"
    Public Const MannualySetMPUploaderData As String = "MannualySetMPUploaderData"
    Public Const AllowSNFNotManditoryInBulkSale As String = "Allow SNF Not Manditory in Bulk Sale"
    Public Const VSPMPDiffrenceOnTSBasis As String = "VSP MP Diffrence On TS Basis"
    Public Const MilkProcuremntPickCLRInsteadOfSNF As String = "Milk Procuremnt Pick CLR Instead Of SNF"
    Public Const FATDivideBy As String = "FAT Divide By"
    Public Const SNFDivideBy As String = "SNF Divide By"
    Public Const PickPriceFromFATAndSNF As String = "Pick Price From FAT And SNF"
    Public Const chkGSTTaxGroupValidity As String = "check GST Tax Group Validity"
    'GHO- Date : 29-Aug-2017
    Public Const ShowShipToPartyInDairyDispatch As String = "Show Ship To Party In Dairy Dispatch"
    Public Const BulkQCWithoutCLR As String = "Bulk Quality Check Without CLR"
    Public Const DOTaggingForDairySaleModule As String = "DO Tagging For Dairy Sale Module"
    Public Const AllowFractionInMCCTankerDispatchGrossQty As String = "Allow Fraction In MCC Tanker Dispatch Gross Qty"
    Public Const PurchaseModulePickFixTaxRate As String = "Purchase Module Pick Fix Tax Rate"
    Public Const TankerDispatchFinancialImpactInTransferIn As String = "Tanker Dispatch Financial Impact In Transfer In"
    Public Const ConvertQtyIntoKG = "Convert Qty into KG Bulk Dispatch"
    Public Const GSTExemptedAmountForNonRegisteredVendor As String = "GST Exempted Amount For Non Registered Vendor"
    Public Const IncreaseCrateQtyOnFiftyPercent As String = "Increase Crate Qty On Fifty Percent"
    Public Const FATSNFDeductionMixMilkFATMinValue As String = "FAT SNF Deduction Mix Milk FAT Min Value"
    Public Const FATSNFDeductionMixMilkFATMaxValue As String = "FAT SNF Deduction Mix Milk FAT Max Value"
    Public Const FATSNFDeductionMixMilkSNFMinValue As String = "FAT SNF Deduction Mix Milk SNF Min Value"
    Public Const FATSNFDeductionMixMilkSNFMaxValue As String = "FAT SNF Deduction Mix Milk SNF Max Value"
    Public Const FATSNFDeductionMixMilkDeductionPer As String = "FAT SNF Deduction Mix Milk Deduction Per"
    Public Const VSPHoldPaymentNotCompanyBank As String = "VSP Hold Payment Not Company Bank"
    Public Const TDSRoundOffAmount As String = "TDS Round Off Amount"
    Public Const RoundOffPaiseAmount As String = "Round Off Paise Amount"
    Public Const RemoveSavingDocumentWhenPayableAmtZero As String = "Remove Saving Document When Payable Amt Zero"
    Public Const RecalculatePaymentProcessOnSave As String = "Recalculate Payment Process On Save"
    Public Const EnableInternalTransfer As String = "Enable Internal Transfer for UDL"
    Public Const JobWorkTransfer As String = "Job Work Transfer"
    Public Const FreightProvisionAccount As String = "Freight Provision Account"
    Public Const FreightProvisionAccountInward As String = "Freight Provision Account Inward"
    Public Const TreatUnregisteredVendorAsRegisteredVendor As String = "Treat Unregistered Vendor As Registered Vendor"
    Public Const RecreateConsumptionEntry As String = "RecreateConsumptionEntry"
    Public Const SaleRemarksForTruckSheetReport As String = "SaleRemarksForTruckSheetReport"
    Public Const BankRecoHidePWD As String = "Bank Reco Hide PWD"
    Public Const EnableItemGroupGLMapping As String = "Enable Item Group GL Mapping"
    Public Const EnableRackBin As String = "Enable Rack Bin Item"
    Public Const ChangeVehicleOnDairySaleBooking = "Change Vehicle On Dairy Sale Booking"
    Public Const VendorSetOffDayWise = "Vendor Set Off Day Wise"
    Public Const ReadOnlyTemplateFieldsOnAcqusition As String = "ReadOnlyTemplateFieldsOnAcqusition"
    Public Const IsAutoStartReading As String = "IsAutoStartReading"
    Public Const AddHighSecurityOnWeighingIntegratedScreen As String = "Add High Security On Weighing Integrated Screen"
    Public Const HighSecurityStableSeconds As String = "High Security Stable Seconds"
    Public Const HighSecurityWeightTolerance As String = "High Security Weight Tolerance"
    Public Const AllowManualvehicleOnDairyBooking As String = "AllowManualvehicleOnDairyBoking"
    Public Const FreeIndentQtyAfterPOClose As String = "Free Indent Qty After PO Close"
    Public Const ShowFATSNFinPaymentProcess As String = "Show FAT SNF in Payment Process"
    Public Const MaxRowsInCSVExport As String = "MaxRowsInCSVExport"
    Public Const MaxRowsInExcelExport As String = "MaxRowsInExcelExport"
    Public Const BigValidity As String = "Big Validity"
    Public Const AllowAssetBookChangeInTemplate As String = "AllowAssetBookChangeInTemplate"
    Public Const AllowSMSSendtoSalePerson As String = "Allow SMS Send to Sale Person"
    Public Const AllowSMSwhenCustomerCreditLimit As String = "SMS when Customer Credit limit reaches on DO."
    Public Const EnableScreenSelection As String = "Enable Screen Selection"
    Public Const SkipJobWorkSRNInPI = "Skip JobWork SRN in PI"
    Public Const ShowFatSnfAfterApproval = "Show Fat/Snf After Approval"
    Public Const ApplyTotalSolidPriceChart As String = "Apply Total Solid Price Chart"
    Public Const RequiredMgmtApprovalForRateIncrease As String = "Required Mgmt Approval For Rate Increase"
    Public Const AutoRoundOffSeprateAccountOnVendorTransaction As String = "Auto Round Off Seprate Account on Vendor Trans"
    '=================Added by preeti Gupta Against Ticket No[ADV/17/05/18-000032]===================================
    Public Const TreatCRATEAsItems As String = "Treat CRATE as Item"
    Public Const TreatCANAsItems As String = "Treat CAN as Item"
    Public Const DoNotShowDairyTypeItems As String = "Do not show dairy type items"
    '=============================================================================================
    Public Const PasswordRules As String = "Password Rules"
    Public Const AlwaysVSPDefaulter As String = "Always VSP Defaulter"
    Public Const RejectedMilkSendToRejectLocation As String = "Rejected Milk Send To Reject Location"
    Public Const NoOfPreNxtDayToPickAvgFATSNF As String = "No Of Pre Nxt Day To Pick Avg FAT SNF %"
    Public Const AutoGeneratePrefix As String = "Auto Generate Prefix"
    Public Const SingleUserParticularDairyBookingEdit = "SingleUserCanEditOneDairyBookingDocumentAtaTime"
    Public Const DairyBookingUnlock As String = "DairyBookingUnlock"
    Public Const MaxReceiveSNFPer As String = "Max Receive SNF %"
    Public Const MailForAdvancePaymentTerm As String = "send mail for advance payment term"
    Public Const DonotCheckAvgQtyOnDairyBooking As String = "Check Average Qty On Dairy Booking"
    Public Const leakagededuction_freshsale As String = "Leakage Deduction Freshsale (%) "
    Public Const FirstGateOutProcessForMCCBulkProcument As String = "First Gate Out Process For MCC Bulk Procument"
    Public Const MCCBulkProcumentSecurityGateOut As String = "MCC Bulk Procument Security Gate Out"
    Public Const EnableDistributorSubsidy As String = "Enable Distributor Subsidy"
    Public Const DoNotAllowSavePOWhenQtyNRateZero As String = "Do Not Allow Save PO When Qty and Rate Zero"
    Public Const ActivateSFGProduction As String = "ActivateSFGProduction"
    Public Const ShowOnlyProductionItemInAdRemove As String = "ShowOnlyProductionItemInAdRemove"
    Public Const PeriodofSubsidyCreditNote As String = "Subsidy Credit Note Period"
    Public Const CHADetailsMandatoryOnExportSale As String = "CHA Details Mandatory On ExportSale"
    Public Const KnockOffFATSNFKG As String = "Knock Off FAT SNF KG"
    Public Const CreateLoadINSlipVehicleWise As String = "Create Load IN Slip Vehicle Wise"
    Public Const RouteCodeNotMandatoryOnLoadINSlip As String = "Route Code Not Mandatory On Loadin Slip"
    Public Const CrateReceiveddairyRouteWise As String = "Crate Received dairy Route Wise"
    Public Const RequiredFinalQCofstandardization As String = "Required Final QC of standardization"
    Public Const PickOnAccountPaymentForAdvanceKnockOff As String = "Pick OnAccount Payment For Advance Knock Off"
    Public Const PickProductCostFromItemUOMDetail As String = "Pick Product Cost From Item UOM Detail"
    Public Const AllowNegativeStockInDairyProduction As String = "Allow Negative Stock In Dairy Production"
    Public Const SkipLockLocation As String = "Skip Lock Location"
    Public Const ItemStructureMandatoryOnWeightConversion As String = "Item Structure Mandatory On Weight Conversion"
    Public Const VillageDiffrenceReportValueWithTwoDecimalPoint As String = "Village Diff Report Value With Two Decimal Point"
    Public Const RequiredFinalQCofProductionEntry As String = "Required Final QC of ProductionEntry"
    Public Const ShowAllPendingDOIrrespectiveOfDeliveryDate As String = "Show All Pendening DO Irrespective Of DeliveryDate"
    Public Const PickFATSNFPerFromStock As String = "Pick FAT SNF % From Stock"
    Public Const ShowBulkDispatchQtyInLtr As String = "Show Bulk Dispatch Qty In Ltr"
    Public Const DoNotIncludeIncentiveInMilkPurchaseInvoice As String = "Do Not Include Incentive In Milk Purchase Invoice"
    Public Const DoNotConsiderTheFutureDateOfAdvancePayment As String = "Do Not Consider The Future Date Of Advance Payment"
    Public Const CreateEmpCodeAsPerEmployeeBasisType As String = "Create Emp Code As Per Employee Basis Type"
    Public Const FATSNFRateMandatory As String = "FAT SNF Rate Mandatory"
    Public Const ProductionOrStandAccordingToItemType As String = "Production Or Stand. According To Item Type"
    Public Const PwdOpenOnMainGLAccountAfterSave As String = "PwdOpenOnMainGLAccountAfterSave"
    Public Const UseProductFATSNFKgForEstimationCost As String = "Use Product FAT SNF Kg For Estimation Cost"
    Public Const ShowItemInCaseofNonInventory As String = "Show Item In Case of Non Inventory"
    Public Const ProductionFATRateTollerance As String = "Production FAT Rate Tollerance"
    Public Const ProductionSNFRateTollerance As String = "Production SNF Rate Tollerance"
    Public Const MaxFATPerLimitOwnDCS As String = "Max FAT Per Limit Own DCS"
    Public Const MaxSNFPerLimitOwnDCS As String = "Max SNF Per Limit Own DCS"
    Public Const MaxFATPerLimit As String = "Max FAT Per Limit"
    Public Const MaxSNFPerLimit As String = "Max SNF Per Limit"
    Public Const MinFATPerLimit As String = "Min FAT Per Limit"
    Public Const MinSNFPerLimit As String = "Min SNF Per Limit"
    Public Const ShowRateAndAmountOnFarmerCollection As String = "Show Rate And Amount On Farmer Collection"
    Public Const MorningShiftStartTime As String = "Morning Shift Start Time"
    Public Const MorningShiftEndHours As String = "Morning Shift End Hours"
    Public Const EveningShiftStartTime As String = "Evening Shift Start Time"
    Public Const EveningShiftEndHours As String = "Evening Shift End Hours"
    Public Const LockDate As String = "Lock Date"
    Public Const ApplyTransFATSNFRateForCalculateFATSNFRate As String = "Apply Trans FAT SNF Rate For Calculate FATSNF Rate"
    Public Const GrossWeightChangePWD As String = "Gross Weight Change PWD"
    Public Const MachineIntegrationInGeneralWeighment As String = "Machine Integrate In General Weighment"
    Public Const FillGeneralWeighmentDetailsByJobworkTypeGateInNo As String = "Fill General Weighment By Jobwork Type Gate In No"
    Public Const GateOutTankerNoAfterGeneralWeighment As String = "Gate Out Tanker No After General Weighment"
    Public Const NoOFCustomerForImportExport As String = "No Of Customer For Import Export"
    Public Const NoOFIncentiveForMPImportExport As String = "No OF Incentive For MP Import Export"
    Public Const NoOFSlabForImportExport As String = "No OF Slab For Import Export "
    Public Const NoOFItemStructureForImportExport As String = "No OF Item Structure For Import Export"
    Public Const CheckUnpostedPaymentProcess As String = "Check Unposted Payment Process"
    Public Const NoOFSavingCodeForImportExport As String = "No OF Saving Code For Import Export "
    Public Const NoOFCustomerForImportExportOnRouteMaster As String = "No OF Customer For Import Export On Route Master"
    Public Const CrateToLTR = "CRATE To LTR"
    Public Const CanToLTR = "CAN To LTR"
    Public Const ProdcutionDoNotCheckForwardDocuments = "Production Do Not Check Forward Documents"
    Public Const PurchaseDoNotCheckForwardDocuments = "Purchase Do Not Check Forward Documents"
    Public Const DoNotStopOnItemBalanceExceptionStoreAdjustment = "Do Not Stop on Balance Exception Store Adjustment"
    Public Const ProductionIssueQtyTollerance = "Production Issue Qty Tollerance"
    Public Const NopreviousDaysInSaleVSReceipt = "No of previous days in Sale vs Receipt"
    Public Const SelectGLInProftAndLossPerforma = "Select GL In Proft And Loss Performa"
    Public Const SelectGLInBalanceSheetPerforma = "Select GL In Balance Sheet Performa"
    Public Const BalanceSheetPerformaWithFormula = "Balance Sheet Performa With Formula"
    Public Const SelectGLInCashFlowPerforma = "Select GL In Cash Flow Performa"
    Public Const BulkProcurementApplyTotalSoidRate = "Bulk Procurement Apply Total Soid Rate"
    Public Const CalculateLtrQtyFromKGQtyByCLR = "Calculate Ltr Qty From KG Qty By CLR"
    Public Const CalculateProvisionOnGateePass = "Calculate Provision On Gatee Pass"
    Public Const CalculateSNFFromCLRForMCCMilk = "Calculate SNF From CLR For MCC Milk"
    Public Const TagExemptedtaxgroupincaseofBankChargesinPaymentEntry = "Exempted group incase of Bnk Charges in Paymt Enty"
    Public Const SNFFromCLRAndCorrectionFactorInJWIEst = "SNF From CLR And Correction Factor in JWI Est"
    Public Const AutoCalculateProduceQty = "Auto Calculate Produce Qty"
    Public Const SyncedMccToServerStartDateForEmailSms = "Synced Mcc To Server Start Date For Email Sms"
    Public Const CrateReceiveddairyCustomerWise As String = "Crate Received dairy Customer Wise"
    Public Const MaintainLogForImporperSample = "Maintain Log For Imporper Sample"
    Public Const PFCalculationOnFormulaHead As String = "PF Calculation On Head Value"
    Public Const MODValueForFAT As String = "MOD Value For FAT" ''ERO/25/02/19-000497 by balwinder on 22/Feb/2019
    Public Const AutoFillSameLocationInGrid As String = "Auto Fill Same Location In Grid"
    Public Const ApplyCalculateWeightInLtr = "Apply Calculate Weight In Ltr"
    Public Const ApplyTransportChargeAddInActualAmount = "Apply Transport Charge Add In Actual Amount"
    Public Const CheckUniqueDocumentCode = "Check Unique Document Code"
    Public Const JWIRateofFGasPerRM = "JWI Rate of FG as per RM"
    Public Const TagMultipleRouteWithCustomer = "Tag Multiple Route With Customer"
    Public Const ApplyNotShowJobWorkTypeTanker = "Apply Not Show JobWork Type Tanker"
    Public Const ExportToDefineLocation As String = "Export To Define Location"
    Public Const ImportMultipleAssetAssembled As String = "ImportMultipleAssetAssembled"
    Public Const ShowEmpCurrentSalaryOnEmployeeSatatusReport As String = "Show Emp Current Salary On Employee Satatus Report"
    Public Const AllowOneFormatByDefaultForPrint = "Allow printing One formats by default"
    Public Const CreateProvisionOnOpeningAndClosingKM = "Create Provision On Opening And Closing KM"
    Public Const FromLocationStockNotCheckConsumptionLocation = "From Location Stock not check Consumption Location"
    Public Const MilkIncetiveByMilkSRN = "Milk Incetive By Milk SRN"
    Public Const AllowShowCoumnInCrateReceivedDairy = "Show Coumns In Crate Received Dairy"
    Public Const CalculateLeakageAmount As String = "Calculate Leakage Amount"
    Public Const ShowFATSNFPerOnBulkProcInGateIN As String = "Show Fat% and SNF% On Bulk Proc In Gate IN Screen"
    Public Const AllowItemCostMandatoryForStockingUnit As String = "Allow Item Cost Mandatory for Stocking Unit"
    Public Const UseControlMForHelp As String = "Use Control M For Help"
    Public Const DairyBookingTolleranceQty As String = "Dairy Booking Tollerance Qty"
    Public Const VehicleCodeNotMandatoryOnLoadINSlip As String = "Vehicle Code Not Mandatory On LoadIN Slip"
    Public Const ApplyLatestPriceChartWhilecreatingNewVSP As String = "Apply Latest Price Chart While creating New VSP"
    Public Const AllowToPrintInvoiceAfterPosting As String = "Allow To Print Invoice After Posting"
    Public Const EnableBankFromMaster As String = "Enable Bank From Master"
    Public Const Autoadjustdebitcreditnote As String = "Auto Adjust Debit Credit Note"
    Public Const UpdateItemMasterWithoutTransactionValidation As String = "Update Item Master Without Transaction Validation"
    Public Const AddItemAliasInSMS As String = "Add Item Alias In SMS"
    Public Const ItemWiseQualityCheckInGeneralPurchase As String = "Item Wise Quality Check In General Purchase"
    Public Const SendInternalSMSEmailinPurchaseModule As String = "Send Internal SMS/Email in Purchase Module"
    Public Const AllowtoEnterNetWeightManuallyinPOWeighmentScreen As String = "AllowtoEnterNetWeightManuallyinPOWeighmentScreen"
    Public Const ShowNotificationInMDI As String = "Show Notification In MDI"
    Public Const TIPRateMix As String = "TIP Rate Mix"
    Public Const TIPRateCow As String = "TIP Rate Cow"
    Public Const TIPRateBuffalo As String = "TIP Rate Buffalo"
    Public Const DefaultCustomerGroupCodeforVSP As String = "Default Customer Group Code for VSP"
    Public Const DefaultVendorGroupCodeforVSP As String = "Default Vendor Group Code for VSP"
    Public Const AllowSameTankerNoforPrimarySecondaryTransporter As String = "AllowSameTankerNoforPrimarySecondaryTransporter"
    Public Const AllocateToMandatoryonGateOut As String = "AllocateToMandatoryonGateOut"
    Public Const PrintTruckSheetAfterGenerate As String = "PrintTruckSheetAfterGenerate"
    Public Const AskForPwdForOutAdjustmentOnPost As String = "AskForPwdForOutAdjustmentOnPost"
    Public Const checkStockOfItemTillTransactionDateOnly As String = "check stock of item till transaction date only"
    Public Const UseProductionPlaningDateForWholeProductionCycle As String = "UseProductionPlaningDateForWholeProductionCycle"
    Public Const allowMilkJWOutowordWithAvgFatSNFRate As String = "allowMilkJWOutowordWithAvgFatSNFRate"
    Public Const allowMilkJWOutowordWithAvgFatSNFPerAtInventory As String = "allowMilkJWOutowordWithAvgFatSNFPerAtInventory"
    Public Const CreateProvisionOfTankerDispatchWithClosingKM As String = "CreateProvisionOfTankerDispatchWithClosingKM"
    Public Const MaterialSaleInvoiceEnablePrintOnPost = "MaterialSaleInvoiceEnablePrintOnPost"
    Public Const DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn = "DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn"
    Public Const UpdateItemMasterConversationWithoutValidation = "UpdateItemMasterConversationWithoutValidation"
    Public Const TollTaxMaster = "Toll Tax Master"
    Public Const InDocMandatoryOnInternalTransfer = "InDocMandatoryOnInternalTransfer"
    Public Const AllowTransferInAfterGatePassOnly = "Allow Transfer In After Gate Pass Only"
    Public Const EnableItemShortDescriptionInBooking = "Enable Item Short Description In Booking"
    Public Const AllowDuplicateItemShortDescriptionInItemMaster = "AllowDuplicateItemShortDescriptionInItemMaster"
    Public Const DeleteMccMilkShiftPassword As String = "Delete Mcc Milk Shift Password"
    Public Const DoNotCreatePaymentWhileSalaryGeneration As String = "DoNotCreatePaymentWhileSalaryGeneration"
    Public Const MccPlantSelectionOptionInMccTankerGateOut As String = "MccPlantSelectionOptionInMccTankerGateOut"
    Public Const EnableTankerNoInMccTankerDispWithMccTankerGateOut As String = "EnableTankerNoInMccTankerDispWithMccTankerGateOut"
    Public Const AllowToEnterSnfAtPlantInMccTankerDispatch As String = "AllowToEnterSnfAtPlantInMccTankerDispatch"
    Public Const DateOfEInvoiceImplementation As String = "DateOfEInvoiceImplementation"
    Public Const BennyImportAutoCreateMP As String = "Benny Import Auto Create MP"
    Public Const BennyImportPickRateFromPrice As String = "Benny Import Pick Rate From Price"
    Public Const AllowOnlyOneIssueAgainstStoreRequisition As String = "AllowOnlyOneIssueAgainstStoreRequisition"
    Public Const UseVLCUploaderCodeMPUploaderCodeInMCCProcurement As String = "UseVLCUploaderCodeMPUploaderCodeInMCCProcurement"
    Public Const ShowNotificationWithoutSMSAPP As String = "ShowNotificationWithoutSMSAPP"
    Public Const SetNotificationRefreshTimeInMinutes As String = "SetNotificationRefreshTimeInMinutes"
    Public Const GenerateEWayBillWithEInvoice As String = "GenerateEWayBillWithEInvoice"
    Public Const VillageDataReverse As String = "VillageDataReverse"
    Public Const ShowTankerWithoutCheckingAnyValidation As String = "ShowTankerWithoutCheckingAnyValidation"
    'Public Const AllocatedTankerGateOut As String = "AllocatedTankerGateOut"
    Public Const AllowToCreateNoOfBookingPerDay As String = "AllowToCreateNoOfBookingPerDay"
    Public Const EnterWeightManuallyOnWeighmentInGrid As String = "EnterWeightManuallyOnWeighmentInGrid"
    Public Const AllowSameChequeNoForMultiplePaymentEntry As String = "AllowSameChequeNoForMultiplePaymentEntry"
    Public Const AllowGenerateReferenceNoForBulkGateEntry As String = "Allow Generate Reference No For Bulk Gate Entry"
    Public Const CreateNeftuploaderPlantWise As String = "Create Neftuploader Plant Wise"
    Public Const AllowBankTransferAgainstMilkBill As String = "Allow Bank Transfer Against Milk Bill"
    Public Const ShowTCSAmountOnBookingForOtherCustomer As String = "ShowTCSAmountOnBookingForOtherCustomer"
    Public Const ShowCheckExcludeProvisionBank As String = "Show Check Exclude Provision Bank"
    Public Const ChangeLeaveDescriptionOnSalarySlip As String = "ChangeLeaveDescriptionOnSalarySlip"
    Public Const ApplyDepartmentWiseDataVisibleInDepartmentIndent As String = "ApplyDepartmentWiseDataVisibleInDepartmentIndent"
    Public Const UpdateMapPayHeadsToSalaStructurePassword As String = "UpdateMapPayHeadsToSalaStructurePassword"
    Public Const ApplyFinancialCostCenter As String = "Apply Financial Cost Center"
    Public Const SalarySlipLeaveStatusOnTheBasisOfCalendarYear As String = "SalarySlipLeaveStatusOnTheBasisOfCalendarYear"
    Public Const ApplyCalculationOnRouteLenth As String = "ApplyCalculationOnRouteLenth"
    Public Const EnableExportExcelOnIncentiveEntry As String = "EnableExportExcelOnIncentiveEntry"
    Public Const AllowBankSectionEnableOnMCCMaster As String = "AllowBankSectionEnableOnMCCMaster"
    Public Const DateOfDynamicQRCodeForB2CInvoiceImplementation As String = "DateOfDynamicQRCodeForB2CInvoiceImplementation"
    Public Const ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster As String = "ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster"
    Public Const AllowRoundDownAmtForMCCDateWiseChilling As String = "AllowRoundDownAmtForMCCDateWiseChilling"
    Public Const CtreateJEOfVspAssetIssueAndReturn As String = "CtreateJEOfVspAssetIssueAndReturn"
    Public Const createDebitNoteOnAssetIssue As String = "createDebitNoteOnAssetIssue"
    Public Const PickAvgCostonAssetissue As String = "PickAvgCostonAssetissue"
    Public Const ApplyTCSAmtOnAbstractReportDotMatrix As String = "ApplyTCSAmtOnAbstractReportDotMatrix"
    Public Const NotAllowDuplicatePANOnPrimaryTransporter As String = "NotAllowDuplicatePANOnPrimaryTransporter"
    Public Const ApplyIncludeTCSAmountInRouteTotalOnTruckSheet As String = "ApplyIncludeTCSAmountInRouteTotalOnTruckSheet"
    Public Const ShowEarlyRouteOnTruckSheet As String = "ShowEarlyRouteOnTruckSheet"
    Public Const AllowOutEntryOnCrateReceivedDairyForAdjustment As String = "AllowOutEntryOnCrateReceivedDairyForAdjustment"
    Public Const BulkMilkFATSNFKGDecimalPlaces As String = "Bulk Milk FAT SNF KG Decimal Places"
    Public Const BulkMilkFATSNFAmtDecimalPlaces As String = "Bulk Milk FAT SNF Amt Decimal Places"
    Public Const BulkMilkConsiderAllParametersForIncetive As String = "Bulk Milk Consider All Parameters For Incetive"
    Public Const MCCMaterialSaleFarmerReverse As String = "MCCMaterialSaleFarmerReverse"
    Public Const LocalSaleAllowedPer As String = "Local Sale Allowed Per"
    Public Const LocalSaleAllowedRate As String = "Local Sale Allowed Rate"
    Public Const MultiDairyGatePassReversePWD As String = "MultiDairyGatePassReversePWD"
    Public Const CostCenterAndHirerachyCodeUpdateAfterPost As String = "CostCenter And Hirerachy Code Update After Post"
    Public Const ShowStatusItemWiseInPendingRequisitionRpt As String = "Show Status Item Wise In Pending Requisition Rpt"
    Public Const DoNotCheckAnyValidationOnVendorInactive As String = "Do Not Check Any Validation On Vendor Inactive"
    Public Const SupportHindiFont As String = "Support Hindi Font"
    Public Const ConvertHindiFontLastDateTime As String = "Convert Hindi Font Last Date Time"
    Public Const ApplyMilkPouchPrint As String = "ApplyMilkPouchPrint"
    Public Const UserWiseRouteMapping As String = "UserWiseRouteMapping"
    Public Const CreateMCCTankerGateOutBasedOnBulkRouteMaster As String = "CreateMCCTankerGateOutBasedOnBulkRouteMaster"
    Public Const ApplyAutoLoginIdCreateOfCustomer As String = "ApplyAutoLoginIdCreateOfCustomer"
    Public Const DefaultLocationOfUserMaster As String = "DefaultLocationOfUserMaster"
    Public Const CreateGatePassFromDemand As String = "CreateGatePassFromDemand"
    Public Const PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister As String = "PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister"
    Public Const AadharNoMandatoryOnEmpMaster As String = "Aadhar No Mandatory On Emp Master"
    Public Const SuperUserCustomer As String = "SuperUserCustomer"
    Public Const UploadMultipleMasterPwd As String = "UploadMultipleMasterPwd"
    Public Const ApplyDefaultsInMaster As String = "ApplyDefaultsInMaster"
    Public Const IncentiveAccNoMandatoryInMPMaster As String = "IncentiveAccNoMandatoryInMPMaster"
    Public Const Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster As String = "Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster"
    Public Const MultipleFinderFillAuto = "MultipleFinderFillAuto"
    Public Const MandatoryLineNoMaxMinQtyForProductionPlan As String = "MandatoryLineNoMaxMinQtyForProductionPlan"
    Public Const RunProductionBaseOnPercentage As String = "RunProductionBaseOnPercentage"
    Public Const RCDFCFP As String = "RCDFCFP"
    Public Const ApplyLocationFilterBasedOnPermission As String = "ApplyLocationFilterBasedOnPermission"
    Public Const MandatoryPDFFileMilkPricePlan As String = "Mandatory PDF File In Milk Price Plan"
    Public Const AutoClosePOBasedOnSRNQtyWithTolerance As String = "AutoClosePOBasedOnSRNQtyWithTolerance"
    Public Const ApplyMilkTypeBuffaloCowOnPrint As String = "ApplyMilkTypeBuffaloCowOnPrint"
    Public Const ApplyZoneWiseVSP As String = "ApplyZoneWiseVSP"
    Public Const ApplyStandardProductionVariance As String = "ApplyStandardProductionVariance"
    Public Const ItemCostTolerancePercentage = "ItemCostTolerancePercentage"
    Public Const HeadLoadDescriptionInPaymentProcessPrint = "HeadLoadDescriptionInPaymentProcessPrint"
    Public Const PrefixForUserMaster = "Prefix For User Master"
    Public Const ApplyDemandApproval = "ApplyDemandApproval"
    Public Const ApplyDemandAll = "ApplyDemandAll"
    Public Const ApplyDemandCustomerWise = "ApplyDemandCustomerWise"
    Public Const CheckCreditLimit = "CheckCreditLimit"
    Public Const ApplyTolerance = "ApplyTolerance"
    Public Const ApplyOrderByNumeric = "ApplyOrderByNumeric"
    Public Const SetShiftTimeOut = "SetShiftTimeOut"
    Public Const AllowToSelectShift = "Allow To Select Shift"
    Public Const ApplyRoundOffZero = "ApplyRoundOffZero"
    Public Const EnableLocation = "EnableLocation"
    Public Const EnableResetDemand = "EnableResetDemand"
    Public Const IsLoadingSlipMandatory = "Is Loading Slip Mandatory"
    Public Const CrateCapacityCheck = "Crate Capacity Check"
    Public Const PickAllBMC = "Pick All BMC"
    Public Const ApplyMonthlyYearly = "Apply Monthly Yearly"
    Public Const SeparateDemandMilkandProduct = "Separate Demand Milk & Product"
    Public Const ShowDemandDoc = "Show Demand Document In Customer Booking"
    Public Const ApplyCommission = "Apply Distributor/Transpoter Commission"
    Public Const PricePlanRoundOffTruncate = "Price Plan RoundOff/Truncate"
    Public Const AllowIncreaseDispatchQty = "Allow Increase Dispatch Qty"
    Public Const FORPRICE = "FOR Price"
    Public Const AllowGatePassDemandTripWise = "Allow Gatepass Demand Trip Wise"
    Public Const ConvertPouchtoCrate = "Convert 1 Pouch to Crate"
    Public Const DontCreateForPouch = "do not calculate Crate for pouch"
    Public Const ChangeInventroyMovemnet = "Inventroy Movement from Dispath/Gatepass"
    Public Const ApplyCommissionRateWithTax = "Apply Commission Rate With TAX"
    Public Const DisableRouteandVehicle = "Disable Route and Vehicle No."
    Public Const UpdateDemandBeforePost = "Update Demand Before Post"
    Public Const SuperuserViewOnly = "Super User View Only"
    Public Const DistributorViewOnly = "Distributor View Only"
    Public Const ShowDemandvsDispatch = "Show Demand vs Dispatch"
    Public Const ShowFreshAmbientItems = "Show Fresh Ambient Items"
    Public Const CustomerReorder = "Customer Reorder"
    Public Const AcknowledgeDispatchQty = "CNF User Acknowledge Dispatch Qty"
    Public Const ApplyPricePlanDigitRule = "Apply Price Plan Digit Rule"
    Public Const PricePlanRoundOffDigit = "Enter RoundOff Digit"
    Public Const DemandIncreaseDecreaseQty = "Demand Increase Decrease Qty"
    Public Const DemandIncreaseDecreasePer = "Demand Increase Decrease Percentage"
    Public Const CreateAutoMCCPriceChat = "Create Auto MCC Price Chat"
    Public Const VehicleNofromDispatch = "VehicleNo From Dispatch"
    Public Const HideQtyCrateRegister = "Hide Qty Crate Register"
    Public Const DispatchPriceCodeForCreditCustomer = "Dispatch PriceCode For Credit Customer"
    Public Const ApplyCommissionTPTForCreditCustomer = "Apply Commission/TPT For Credit Customer"
    Public Const ConvertPouchtoCrateonDispatch = "Convert Pouch to Crate on Dispatch"
    Public Const PickRatefromMaster = "Pick Rate from Master"
    Public Const EnableProductSaleForJPR = "Enable Product Sale For JPR"
    Public Const ShowTodayDemandAsCurrentandUpcoming = "Show Today Demand As Current And Upcoming"
    Public Const EnableVehicleType = "Enable Vehicle Type"
    Public Const PickDataFromRetestingTable = "Pick Data From Retesting Table"
    Public Const ShowAllCustomerItemWiseSaleReportOptions = "Show All Customer Item Wise Sale Report Options"
    Public Const ApplyTPT = "Apply Transpoter Commission"
    Public Const AllowMultipleUOMForProduct = "Allow Multiple UOM For Product"
    Public Const AllowManualCrateForDispatch = "Allow Manual Crate For Dispatch"
    Public Const SetDefaultShiftTime = "Set Default Shift Time"
    Public Const ApplyManualScheme = "Apply Manual Scheme"
    Public Const AllowResetCustomerDemandOnRouteChange = "Allow Reset Customer Demand On Route Change"
    Public Const AutoSchemeOnTotalDispatchQty = "Auto Scheme On Total Dispatch Qty"
    Public Const ApplyDefaultTPTIsChecked = "Apply Default TPT IsChecked"
    Public Const ApplyDefaultTCSIsChecked = "Apply Default TCS IsChecked"
    Public Const ApplyBoothRouteMapping = "Apply Booth Route Mapping"
    Public Const DrippingQuantity = "Dripping Quantity"
    Public Const ApplyDepartmentRoute = "Apply Department Route"
    Public Const QuantityTolerance = "Quantity Tolerance"
    Public Const AmountTolerance = "Amount Tolerance"
    Public Const ApplyItemCapacityLimit = "Apply Item Capacity Limit"
    Public Const SeprateMorningEveningSequence = "Seprate Morning Evening Sequence"
    Public Const FinancialImpactForTPT = "Financial Impact For TPT"
    Public Const FinancialImpactForDistributor = "Financial Impact For Distributor"
    Public Const FinancialImpactForSecurity = "Financial Impact For Security"
    Public Const ShowAndSaveCrystalReportActionType = "Show And Save Crystal Report Action Type"
    Public Const QuickDemandUOMCrate = "Quick Demand UOM Crate"
    Public Const QuickDemandUOMPouch = "Quick Demand UOM Pouch"
    Public Const QuickDemandUOMLtr = "Quick Demand UOM Ltr"
    Public Const ApplyBoothWiseScheme = "Apply Booth Wise Scheme"
    Public Const OfflineDCSMilkCollectionExeVersion = "Offline DCS Milk Collection Exe Version"
    Public Const ApplyMonthEndDispatch = "Apply Month End Dispatch"
    Public Const ApplyItemUOMOnDemand = "Apply Item UOM On Demand"
    Public Const MergeTCAmtofCreditCust = "Merge Transpoter Commission Amt of CreditCust"
    Public Const ConvertIntoBulkUOM = "Convert into Bulk UOM"
    Public Const AddDistributorCommissioninItemPrice = "Add Distributor Commission In Item Price"
    Public Const SubstractDistributorCommissioninItemPrice = "Substract Distributor Commission In Item Price"
    Public Const ConvertTOBillingUOM = "Convert into Billing UOM"
    Public Const DifferentCrateTypeForFGItem = "Different Crate Type For FG Item"
    Public Const StopMultipleLoginInAPP = "Stop Multiple Login In APP"
    Public Const AMCUDetailsHeading1 As String = "AMCU Details Heading 1"
    Public Const AMCUDetailsHeading2 As String = "AMCU Details Heading 2"
    Public Const AMCUDetailsHeading3 As String = "AMCU Details Heading 3"
    Public Const DispatchCommissionDecimalPlaces = "Dispatch Commission Decimal Places"
    Public Const ApplyPricePlanOnDocumentDate = "Apply Price Plan On Document Date"
    Public Const DisplayTransporterAddress = "Display Transporter Address"
    Public Const AllowRouteWiseDemandEntryInDecimal = "Allow Route Wise Demand Entry In Decimal"
    Public Const AllowToCheckEntryUOMForProduct = "Allow To Check Entry UOM For Product"
    Public Const AllowToCheckEntryUOMForGNG = "Allow To Check Entry UOM For GNG"
    Public Const AllowToCheckZeroQtyonDispatch = "Allow To Check Zero Qty on Dispatch"
    Public Const DefaultEnableEWayBill = "Default Enable EWay Bill"
    Public Const CreateAutoGatePass = "Create Auto Gate Pass"
    Public Const ApplyEWBThresholdLimit = "Apply EWB Threshold Limit"
    Public Const EWBThresholdLimitForIntraCity = "EWB Threshold Limit For Intra-City"
    Public Const EWBThresholdLimitForIntraState = "EWB Threshold Limit For Intra-State"
    Public Const EWBThresholdLimitForInterState = "EWB Threshold Limit For Inter-State"
    Public Const DefaultEnableNoTransporter = "Default Enable No Transporter"
    Public Const ManualBatchOnCustomerBooking = "Manual Batch On Customer Booking"
    Public Const GenerateCustomerWiseGatePass = "Generate Customer Wise GatePass"
    Public Const DeductTPTFromDocAmt = "Deduct TPT From Document Amt"
    Public Const CheckCustomeroutStandingAmt = "Check Customer OutStanding Amt"
    Public Const ServerDateTimeForTaxableInvoice = "Server DateTime For Taxable Invoice"
End Class
Public Class clsFixedParameter
#Region "Variables"
    Public Type As String = ""
    Public Code As String = ""
    Public Description As String = ""
    Public Specification As String = ""
#End Region
    Public Shared Function GetData(ByVal strType As String, ByVal strCode As String, ByVal trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where TYPE='" + strType + "' and Code='" + strCode + "'", trans))
    End Function
    Public Shared Function GetSpecification(ByVal strType As String, ByVal strCode As String, ByVal trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Specification from TSPL_FIXED_PARAMETER where TYPE='" + strType + "' and Code='" + strCode + "'", trans))
    End Function
    Public Shared Function GetUserCode()
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select (case when ISNUMERIC(Code) = 1 then '' else Code end) as Code from (
  Select top 1  substring(user_code,1,2) As Code from tspl_user_master where user_code Like '%1%' and User_APP_Type='V' order by user_code desc)x"))
    End Function
    ''Created by Pradeep Sharma on 14/06/13 TO Get Combobox datatable
    Public Shared Function GetCboDataTable(ByVal strType As String, ByVal trans As SqlTransaction) As DataTable
        Dim qry As String = "SELECT Code,DESCRIPTION FROM TSPL_FIXED_PARAMETER where Type= '" + strType + "' "
        Dim dt_Cbo As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Return dt_Cbo
    End Function
    Public Shared Sub UpdateData(ByVal Type As String, ByVal Code As String, ByVal Description As String, ByVal trans As SqlTransaction)
        Dim qry As String = "Update TSPL_FIXED_PARAMETER Set Description='" + Description + "' where TYPE='" + Type + "'"
        If clsCommon.myLen(Code) > 0 Then
            qry += " and Code='" + Code + "'"
        End If
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = String.Empty
    End Sub
    Public Shared Function UpdateFixedParameter(ByVal strType As String, ByVal strCode As String, ByVal strDescription As String, ByVal strSpecification As String, ByVal tran As SqlTransaction) As Boolean
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Description", strDescription)
        clsCommon.AddColumnsForChange(coll, "Specification", strSpecification, True)
        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Update, "Type= '" + strType + "' and Code = '" + strCode + "'", tran)
        Return True
    End Function
    Public Shared Function UpdateFixedParameter(ByVal obj As clsFixedParameter, ByVal trans As SqlTransaction, ByVal isNewEntry As Boolean) As Boolean
        Try
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Specification", obj.Specification)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Type", obj.Type)
                clsCommon.AddColumnsForChange(coll, "Code", obj.Code)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Update, "Type='" + obj.Type + "' AND Code='" + obj.Code + "'", trans)
            End If
            clsCommonFunctionality.SaveHistoryData(False, objCommonVar.CurrentUserCode, obj.Type, "TSPL_FIXED_PARAMETER", "Type", "", "", "", "", "", "", "", "", "Code='" + obj.Code + "", "", "", "", trans)

            '' update program code of the template, asset category and asset group
            If clsCommon.CompairString(obj.Code, "ReadOnlyTemplateFieldsOnAcqusition") = CompairStringResult.Equal Then
                ProgramCodeNew.InsertDefaultValue(clsUserMgtCode.Categories, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, trans)) = "1", "Asset Group", "Asset Category") & "", "1.10.01.03", clsUserMgtCode.SubModuleFixedAssetSetup, 27, 27, trans)
                ProgramCodeNew.InsertDefaultValue(clsUserMgtCode.frmAssetGroups, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, trans)) = "1", "Asset Sub Group", "Asset Group") & "", "1.10.01.06", clsUserMgtCode.SubModuleFixedAssetSetup, 27, 27, trans)
                ProgramCodeNew.InsertDefaultValue(clsUserMgtCode.Template, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, trans)) = "1", "Asset Category Master", "Asset Template") & "", "1.10.01.09", clsUserMgtCode.SubModuleFixedAssetSetup, 27, 27, trans)
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function GetFixedParameter(ByVal trans As SqlTransaction) As DataTable
        Return GetFixedParameter(trans)
    End Function
    Public Shared Function GetFixedParameter(ByVal trans As SqlTransaction, ByVal strWhr As String) As DataTable
        Try
            Dim Qry As String = "select Type, Code, Description, Specification  from TSPL_FIXED_PARAMETER where 2=2 "
            If clsCommon.myLen(strWhr) > 0 Then
                Qry += " and " + strWhr
            End If
            Return clsDBFuncationality.GetDataTable(Qry)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function FixedParameterValues() As Boolean
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkadaIntegration, clsFixedParameterCode.SkadaDefaultSiloLocation, "", "Default silo location")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkadaIntegration, clsFixedParameterCode.SkadaLastDocument, "", "Last Synced Document")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkadaIntegration, clsFixedParameterCode.SkadaDefaultRoute, "", "Default Bulk Route")

        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipDebitNoteforRejectInBMCTruckSheet, clsFixedParameterCode.SkipDebitNoteforRejectInBMCTruckSheet, "0", "0:Off;1-ON Skip to Generate Debit Note for Reject In BMC Truck Sheet")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GatepassForTaxableandNonTaxableItems, clsFixedParameterCode.GatepassForTaxableandNonTaxableItems, "0", "0:Off;1-ON When ON can choose taxable or nontaxable option")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickDataFromMasterOrTransaction, clsFixedParameterCode.PickDataFromMasterOrTransaction, "0", "0:Off;1-ON When ON From Transaction else Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ExportTruckSheet, clsFixedParameterCode.ExportTruckSheet, "0", "0:Off;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ThirtPartyFarmerCollectionIntegration, clsFixedParameterCode.ThirtPartyFarmerCollectionIntegration, "0", "0:Off;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MaxFATPerLimit, "15", "Android Farmer Collection Max FAT Per")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MaxSNFPerLimit, "15", "Android Farmer Collection Max SNF Per")

        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MorningShiftStartTime, "06:00:00AM", "Farmer Milk Collection Morning Shift Start Time")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MorningShiftEndHours, "4", "Morning Shift End Time =Morning Shift Start Time + Hours")

        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.EveningShiftStartTime, "06:00:00PM", "Farmer Milk Collection Evening Shift Start Time")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.EveningShiftEndHours, "4", "Evening Shift End Time =Evening Shift Start Time + Hours")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowRateAndAmountOnFarmerCollection, clsFixedParameterCode.ShowRateAndAmountOnFarmerCollection, "0", "0:Off;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SuspenceAdjusmtmentFATSNFTolerancePercentage, clsFixedParameterCode.SuspenceAdjusmtmentFATSNFTolerancePercentage, "1.5", "Suspence Adjusmtment FAT SNF Tolerance Percentage")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransactionModify, clsFixedParameterCode.TransactionModify, "Trans@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransactionValidity, clsFixedParameterCode.TransactionValidity, "30", "Set Transaction Validity Day")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGenerateFarmerPriceFromDCSPrice, clsFixedParameterCode.AutoGenerateFarmerPriceFromDCSPrice, "1", "0:OFF,1:ON;if on then Auto Genrate Farmer Price on Post dcs Price of Transaction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGenrateBatchInventoryOnPost, clsFixedParameterCode.AutoGenrateBatchInventoryOnPost, "0", "0:OFF,1:ON;if on then Auto Genrate Batch Inventory At Post of Transaction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAMCUDetails, clsFixedParameterCode.AMCUDetailsHeading1, "Heading 1", "Dynamic Heading 1 for AMCU Detail")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAMCUDetails, clsFixedParameterCode.AMCUDetailsHeading2, "Heading 2", "Dynamic Heading 2 for AMCU Detail")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAMCUDetails, clsFixedParameterCode.AMCUDetailsHeading3, "Heading 3", "Dynamic Heading 3 for AMCU Detail")

        InsertDefaultValueFixedParameter(clsFixedParameterType.PickcompanyBankCodeFromPaymentProcess, clsFixedParameterCode.PickcompanyBankCodeFromPaymentProcess, "0", "0:OFF,1:ON;Pick company BankCode From Payment Process")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RefreshTimeForUnionWiseMilkCollectionInms, clsFixedParameterCode.RefreshTimeForUnionWiseMilkCollectionInms, "60000", "Insert time for refreshing ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isThereOnlyOneRowOfOwnDCS, clsFixedParameterCode.isThereOnlyOneRowOfOwnDCS, "0", "0:OFF,1:ON;Adjust qty of own DCS when there is only one row")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrintOnlyPostedDocument, clsFixedParameterCode.PrintOnlyPostedDocument, "0", "0:OFF,1:ON; Print Only Posted Document")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxFATPerLimitforReport, clsFixedParameterCode.MaxFATPerLimitforReport, "3.2", "MaxFatPerLimitForReport")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PartialSRNAmount, clsFixedParameterCode.PartialSRNAmount, "0", "PartialSRNAmount")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxSNFPerLimitforReport, clsFixedParameterCode.MaxSNFPerLimitforReport, "7.2", "MaxSnfPerLimitForReport")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultFilesPath, clsFixedParameterCode.DefaultFilesPath, "E:\XpertServices\XpertFileUpload\Upload", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.B2CDocumentAmountRangeOtherState, clsFixedParameterCode.B2CDocumentAmountRangeOtherState, "100000", "Define Document Amount Range For Other State B2C Large/Small")
        InsertDefaultValueFixedParameter(clsFixedParameterType.B2CDocumentAmountRange, clsFixedParameterCode.B2CDocumentAmountRange, "250000", "Define Document Amount Range For Same State B2C Large/Small")

        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOfDCSToLoadDeductionData, clsFixedParameterCode.NoOfDCSToLoadDeductionData, "50", "It Should above 50")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSWiseFilterEnableOnSavingCheck, clsFixedParameterCode.DCSWiseFilterEnableOnSavingCheck, "0", "0:OFF,1:ON;Show DCS Filter")
        InsertDefaultValueFixedParameter(clsFixedParameterType.XpertAPI, clsFixedParameterCode.SetDCSAvgFATSNF, "0", "0:OFF;1:ON")

        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBoothSecurity, clsFixedParameterCode.ApplyBoothSecurity, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FileUpload, clsUserMgtCode.frmSendBillToDCS, "0", "0:OFF;1:ON Upload File")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FileUpload, clsUserMgtCode.frmDairyGatePass, "0", "0:OFF;1:ON Upload File")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PenaltyAfterDays, clsFixedParameterCode.PenaltyAfterDays, "30", "Define Penalty Days to Apply")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteTempDataFromTime, clsFixedParameterCode.DeleteTempDataFromTime, "", "Delete Temp Data From Time ")

        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteTempDataToTime, clsFixedParameterCode.DeleteTempDataToTime, "", "Delete Temp Data To Time Apply")

        InsertDefaultValueFixedParameter(clsFixedParameterType.PenaltyCost, clsFixedParameterCode.PenaltyCost, "0.40", "Define Penalty Cost Value for Per Unit")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateProRateAuto, clsFixedParameterCode.CalculateProRateAuto, "0", "0:OFF;1:ON to Calculate Pro Rate Auto")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertVlcCodeUploaderToInt, clsFixedParameterCode.ConvertVlcCodeUploaderToInt, "1", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickOnlyOWNBMCDCS, clsFixedParameterCode.PickOnlyOWNBMCDCS, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrintPouchCrateQtyOnPrint, clsFixedParameterCode.PrintPouchCrateQtyOnPrint, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickHeadLoadRateFromSecretaryMaster, clsFixedParameterCode.PickHeadLoadRateFromSecretaryMaster, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RCDFRateControl, clsFixedParameterCode.RCDFRateControl, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ImportExportDrive, clsFixedParameterCode.ImportExportDrive, "E", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RCDFControl, clsFixedParameterCode.MaxFATPerLimit, "0", "0-OFF.Max FAT % limit on Milk")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RCDFControl, clsFixedParameterCode.MaxSNFPerLimit, "0", "0-OFF.Max SNF % limit on Milk")

        InsertDefaultValueFixedParameter(clsFixedParameterType.EmployeePFRetirementAge, clsFixedParameterCode.EmployeePFRetirementAge, "58", "Define PF Retirement Age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EmployeeRetirementAge, clsFixedParameterCode.EmployeeRetirementAge, "60", "Define Retirement Age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PDAccountPaymanager, clsFixedParameterCode.NoOfRecordForLotNo, "4500", "No of Records to Set Lot No")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankAdviseRequired, clsFixedParameterCode.BankAdviseRequired, "", "Bank Advise Start Date [dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeprateBankForDCSBankAdvice, clsFixedParameterCode.SeprateBankForDCSBankAdvice, "", "Seprate Bank For DCS Bank Advice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyPartialBankAdvise, clsFixedParameterCode.ApplyPartialBankAdvise, "0", "0:OFF;1:ON")

        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMergeForDCSMultipleDays, clsFixedParameterCode.ApplyMergeForDCSMultipleDays, "1", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateFATSNFLossCycleWise, clsFixedParameterCode.CalculateFATSNFLossCycleWise, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseSlab, clsFixedParameterCode.ApplyRange, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseSlab, clsFixedParameterCode.RangeNotApplicable, "0.01-10000", "Doument Not Required")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseSlab, clsFixedParameterCode.RangePO, "10000.01-100000", "PO Mandatory")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseSlab, clsFixedParameterCode.RangeRAL, "100000.01-999999999999", "RAL Mandatory")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RefreshDBTReco, clsFixedParameterCode.RefreshDBTReco, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBTMilkQtyCapping, clsFixedParameterCode.DBTMilkQtyCapping, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBTByDailyFarmerMilkCollection, clsFixedParameterCode.DBTByDailyFarmerMilkCollection, "1", "0:OFF;1:ON")

        InsertDefaultValueFixedParameter(clsFixedParameterType.DistributorWiseBilling, clsFixedParameterCode.DistributorWiseBilling, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidSaleOrder, clsFixedParameterCode.BackDays, "3", "Back Days of From Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidSaleOrder, clsFixedParameterCode.MaximumBackDays, "60", "Back Days of From Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidSaleOrder, clsFixedParameterCode.ViewItemImage, "0", "0:OFF,1:ON;View Item Image")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, "0", "0:OFF;1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NewDCSScreen, clsFixedParameterCode.NewDCSScreen, "0", "0:Off, 1:On;New DCS Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SameBankCodeForCurrentAndSaving, clsFixedParameterCode.SameBankCodeForCurrentAndSaving, "0", "0:Off;1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinimumQtyForHeadLoad, clsFixedParameterCode.MinimumQtyForHeadLoad, "0", "Minimum Qty To Apply For Head Load")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyHeadLoadDayAndCycleWiseForAJM, clsFixedParameterCode.ApplyHeadLoadDayAndCycleWiseForAJM, "0", "0:Off-Head Load Calculation Milk SRN wise;1:On-Head Load Calculation Day wise;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopSetting, clsFixedParameterCode.JournalEntry, "0", "0:OFF:1 Stop Jouranl Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopSetting, clsFixedParameterCode.Inventory, "0", "0:OFF:1 Stop Inventory Movement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopSetting, clsFixedParameterCode.InventoryNew, "0", "0:OFF:1 Stop Inventory Movement New(Milk)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickBulkRoute, clsFixedParameterCode.PickBulkRoute, "1", "0:VLC Master Route;1:Bulk Route Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowMultipleLegers, clsFixedParameterCode.ShowMultipleLegers, "1", "0:Payment Cycle;1:Show Multiple Ledgers")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LoadLedgerMixedMilk, clsFixedParameterCode.LoadLedgerMixedMilk, "0", "0:OFF:1 Show Mixed Milk Only")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HeadLoadRODecimalPlace, clsFixedParameterCode.HeadLoadRODecimalPlace, "2", "Head Load Round Off Decimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HeadLoadROIncreaseAfter, clsFixedParameterCode.HeadLoadROIncreaseAfter, "5", "Head Load Round off Increase Value After")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyUnpaidBank, clsFixedParameterCode.ApplyUnpaidBank, "1", "0:OFF;1-Apply Unpaid Bank")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowZeroFATSNF, clsFixedParameterCode.AllowZeroFATSNF, "0", "0-OFF;1-ON Allow Zero FAT/SNF in Milk Collection")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AdjustFATSNFINOwnVSP, clsFixedParameterCode.AdjustFATSNFINOwnVSP, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AdjustFATSNFINOwnVSP, clsFixedParameterCode.AdjustQtyINOwnVSP, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplySameDayShift, clsFixedParameterCode.ApplySameDayShift, "0", "0-OFF;1-ON Apply same day shift")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickMilkPurchaseInvoiceQtyOrRecoQty, clsFixedParameterCode.PickMilkPurchaseInvoiceQtyOrRecoQty, "0", "0-Milk Purchase Qty;1-Reco Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowMPIncetiveQtyAboveBilledQty, clsFixedParameterCode.AllowMPIncetiveQtyAboveBilledQty, "1", "1-Allow;0-Now Allowed;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HeaderFATSNFKGDecimalPlaces, clsFixedParameterCode.HeaderFATSNFKGDecimalPlaces, "3", "Header FAT/SNF Decaimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFDecimalPlaces, clsFixedParameterCode.SNFDecimalPlaces, "2", "SNF Decaimal Places")

        InsertDefaultValueFixedParameter(clsFixedParameterType.BMCDecimalPlaces, clsFixedParameterCode.BMCFATDecimalPlaces, "2", "FAT Decaimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BMCDecimalPlaces, clsFixedParameterCode.BMCSNFDecimalPlaces, "2", "SNF Decaimal Places")

        InsertDefaultValueFixedParameter(clsFixedParameterType.BMCDCSApplyShift, clsFixedParameterCode.BMCDCSApplyShift, "0", "0-OFF;1-ON BMC/DCS Milk collection Apply shift on header")

        InsertDefaultValueFixedParameter(clsFixedParameterType.HideShiftCollection, clsFixedParameterCode.HideShiftCollection, "0", "0-Show Both Shift;1-Hide Evening Shift;2-Hide Morning Shift")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkCollectionPickBulkRoute, clsFixedParameterCode.MilkCollectionPickBulkRoute, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOwnBMCDCS, clsFixedParameterCode.ShowOwnBMCDCS, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DailyQtyReport, clsFixedParameterCode.FATKGSNFKGRoundOff, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OwnBMCApplicationFATRatio, clsFixedParameterCode.OwnBMCApplicationFATRatio, "100", "Applicable Ratio")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OwnBMCApplicationSNFRatio, clsFixedParameterCode.OwnBMCApplicationSNFRatio, "100", "Applicable Ratio")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OwnBMCCreateDRCRNote, clsFixedParameterCode.OwnBMCCreateDRCRNote, "1", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TDSOpeningBalanceConsiderCreditNote, clsFixedParameterCode.TDSOpeningBalanceConsiderCreditNote, "1", "0-OFF;1-ON Include AP Invoice Credit Note for TDS Opening balance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxRowsExcelDBTNEFTUploader, clsFixedParameterCode.MaxRowsExcelDBTNEFTUploader, "0", "0-All;No. of Rows To Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.XpertAPI, clsFixedParameterCode.WeighingRoundSetting, "NA", "NA;+1(Round Up Decimal Places);-1(Round Down Decimal Places)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.XpertAPI, clsFixedParameterCode.AverageDaysForFATSNF, "10", "No of Days to Get Average")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkRateRoundOffType, clsFixedParameterCode.MilkRateRoundOffType, "0", "0: MidpointRounding.ToEven 39.825=39.82; 1:MidpointRounding.AwayFromZero 39.825=39.83")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSampleNoOnBMC, clsFixedParameterCode.ShowSampleNoOnBMC, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBTRevisePayment, clsFixedParameterCode.DBTRevisePayment, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CycleWiseDBT, clsFixedParameterCode.CycleWiseDBT, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowTempratureOnBMC, clsFixedParameterCode.ShowTempratureOnBMC, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickMilkTransferInData, clsFixedParameterCode.PickMilkTransferInData, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FillRouteTankerNo, clsFixedParameterCode.FillRouteTankerNo, "1", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RepeatBMCSampleNo, clsFixedParameterCode.RepeatBMCSampleNo, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFNoDecimalMCC, clsFixedParameterCode.FATSNFNoDecimalMCC, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFNoDecimalDCS, clsFixedParameterCode.FATSNFNoDecimalDCS, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllMCC, clsFixedParameterCode.ShowAllMCC, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowDCSDetMerge, clsFixedParameterCode.ShowDCSDetMerge, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllDCS, clsFixedParameterCode.ShowAllDCS, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyGaze, clsFixedParameterCode.ApplyGaze, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkCollectionFATSNFType, clsFixedParameterCode.MilkCollectionFATSNFType, "0", "0-Percentage;1-KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkCollectionFATSNFTypeHeader, clsFixedParameterCode.MilkCollectionFATSNFTypeHeader, "1", "0-Percentage;1-KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.JanAadharNoMandatory, clsFixedParameterCode.JanAadharNoMandatory, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopUpdateForWeigingMilkReceipt, clsFixedParameterCode.StopUpdateForWeigingMilkReceipt, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreatePOFromMultipleLocation, clsFixedParameterCode.CreatePOFromMultipleLocation, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyPashuAaharAndMineralMixture, clsFixedParameterCode.ApplyPashuAaharAndMineralMixture, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkProcRunOneTypeGateEntry, clsFixedParameterCode.BulkProcRunOneTypeGateEntry, "0", "0-Both,1-Contractor only,2-Plant/MCC Transfer in 0nly")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyZoneInDBT, clsFixedParameterCode.ApplyZoneInDBT, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MandatoryDCSMPIncetiveReco, "0", "0-OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MatchingUOM, "", "Blank-No Convertion,LTR/KG convert into given UOM")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyGreaterThanDCSQty, "0", "0-OFF,1-ON Valid if MP Qty > DCS Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyLessThanDCSQty, "0", "0-OFF,1-ON Valid if MP Qty < DCS Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyEqualToDCSQty, "0", "0-OFF,1--ON Valid if MP Qty = DCS Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankIFSCCodeValidateByService, clsFixedParameterCode.BankIFSCCodeValidateByService, "0", "0-OFF,1-Service Only,2-ERP Only,3-Service And ERP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.StopDBTBefore, "", "Stop DBT Entry Before date[dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.StopResetPassword, "0", "Message To Prompt")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.OneDBTOneDoc, "", "Set BMC Code For One DBT One Doc")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.StopMPUpdate, "0", "0-OFF, 1:On; Stop To Update Farmer Data By Mobile APP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.ShowForgetPwd, "0", "0-Hide, 1:Show; Show/Hide Forget Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.MarqueText, "Tecxpert Software Pvt Ltd", "Welcome to Saras Pro App")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.CustomerVirtualACCounter, "00001", "Counter of Customer Virtual AC")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPMaster, clsFixedParameterCode.DisableUploaderNo, "0", "MP Regisration Disable(1)/Enable(0) Uploader no")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPMaster, clsFixedParameterCode.JPRDairyMandatoryColumn, "0", "Mandatory Column Account No,IFSC")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPMaster, clsFixedParameterCode.VerifiedJanAadharNo, "0", "0:OFF,1:ON;Pick Farmer only Verified Janaadhar No")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPMaster, clsFixedParameterCode.FarmerLevelCollection, "0", "0:OFF,1:ON;Show Famer menu in Pro")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPIncetiveEntry, clsFixedParameterCode.MultipleEntryScreen, "0", "0-Sinle Entry Screen;1-Multiple MP Entry Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPIncetiveEntry, clsFixedParameterCode.MultipleEntryScreenAdmin, "1", "0-Sinle Entry Screen;1-Multiple MP Entry Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPIncetiveEntry, clsFixedParameterCode.QtyDecimalPlaces, "0", "[0,1,2] Qty Decimal Places In MP Incentive Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMPIncetiveEntry, clsFixedParameterCode.SkipPaymentCycle, "1", "Skip to set to date accroding to payment cycle wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.DashboardDays, "3", "No of Days To Show Dashboard")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.UOM, "0", "0-Both;1-Crate;2-Pouch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.Shift, "0", "0-Both;1-Morning;2-Evening")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.AllowFutureDateBooking, "0", "1-True;0-False")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ApplyOrderConfirmation, "0", "1-True;0-False")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.SelectDemandUOMAtRuntime, "0", "1-True;0-False")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ShowCatalogPrice, "0", "0:Selling,1:Basic")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.StopDemandUpdateForCNF, "0", "0:OFF,1:ON;CNF Type User will not update Demand")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.SuperuserViewOnly, "0", "0:OFF,1:ON;Super User Type View Report Only")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.DistributorViewOnly, "0", "0:OFF,1:ON;Distributor Type Use View Report Only")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ShowDemandvsDispatch, "0", "0:OFF,1:ON;Show Demand VS Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ShowFreshAmbientItems, "0", "0:OFF,1:ON;Show Fresh Ambient Items")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.CustomerReorder, "0", "0:OFF,1:ON;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.AcknowledgeDispatchQty, "0", "0:OFF,1:ON;CNF User Acknowledge Dispatch Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.HideQtyCrateRegister, "0", "0:OFF,1:ON;Demand register Hide Qty crate column")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ShowTodayDemandAsCurrentandUpcoming, "0", "0:OFF,1:ON;Vehicle Type on Distributor Commission")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.StopYesBankVirtualAccount, "1", "0:OFF,1:ON;Stop Yes Bank virtual Account from menu in order app")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.HideShiftFilterOnDemandStatus, "0", "0:OFF,1:ON;Hide Shift Filter On Demand Status report")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableVehicleType, clsFixedParameterCode.EnableVehicleType, "0", "0:OFF,1:ON;Show Today Demand As Currentand Upcoming")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.TolleranceQty, "100", "Tollerance % of Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.TolleranceFAT, "100", "Tollerance % of FAT")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.TolleranceSNF, "100", "Tollerance % of SNF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.IncludeOwnBMC, "1", "1:On,0:Off; Include Own BMC in BMC List")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.TransactionForCurrentDateOnly, "1", "1:On;0:Off Disable Date selection by user")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.AddPostFunctionality, "0", "0:OFF,1:ON;Add Post functionality in BMC DCS APK")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.TuneTime, "NA", "NA [27-Mar-2025 12:00:00AM to 27-Mar-2025 11:59:59PM]; 06:00:00AM [27-Mar-2025 06:00:00AM to 28-Mar-2025 05:59:59AM]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.ViewDCSMilkPurchaseRegister, "0", "0:OFF,1:ON;Show DCS Milk Purchase register on APP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.BMCStartTime, "02:00:00AM", "BMC Milk Collection Start Time")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.BMCEndHours, "12", "BMC End Time = BMC Start Time + Hours")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.DCSEndHours, "6", "DCS End Time = BMC Start Time + Hours")

        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPPVersion, clsFixedParameterCode.comAnchal_ucdfErp, "", "Version of Aanchal Pro APP [0 Skip]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPPVersion, clsFixedParameterCode.comtecxpertappsaras, "", "Version of SARAS Sales APP [0 Skip]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPPVersion, clsFixedParameterCode.comTecxpertSarasPro, "", "Version of SARAS Pro APP [0 Skip]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPPVersion, clsFixedParameterCode.comtecxpertsoftposerode, "", "Version of Amirtha Sales APP [0 Skip]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AndroidAPPVersion, clsFixedParameterCode.XpertMilkCollection, "1.0.0.0", "Xpert Milk Collection Version Control[0.0.0.0 Skip]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, "0", "1:On Apply Monthly,0:Off Apply Payment Cycle Wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSQtyDecimalPlaces, clsFixedParameterCode.DCSQtyDecimalPlaces, "2", "Value should be 0 to 3")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MPIncentiveEntryCycleWiseButNEFTMonthly, clsFixedParameterCode.MPIncentiveEntryCycleWiseButNEFTMonthly, "0", "1:On,0:Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MPIncentiveEntryMaxMilkLimit, clsFixedParameterCode.MPIncentiveEntryMaxMilkLimit, "10000", "Max Farmer milk Limit in a Cycle")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MPIncentiveEntryIncentiveRate, clsFixedParameterCode.MPIncentiveEntryIncentiveRate, "0", "Mp Incentive Rate")
        'InsertDefaultValueFixedParameter(clsFixedParameterType.TrendDiffValueForColor, clsFixedParameterCode.TrendDiffValueForColor, "30", "Diff Qty With Previous Day To Change Color of cell")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPurReturnEvenIfPaymentDone, clsFixedParameterCode.AllowPurReturnEvenIfPaymentDone, "0", "1:On,0:Off Allow Purchase Return Even If Complete Payment Done")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransporterCollection, clsFixedParameterCode.TransporterCollection, "0", "1:On,0:Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProcurmentShiftUploaderNo, clsFixedParameterCode.ProcurmentShiftUploaderNo, "1", "1:Grid,2:Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BholeBabaPaymentProcessProPrintStartDate, clsFixedParameterCode.BholeBabaPaymentProcessProPrintStartDate, "", "Bhole Baba Payment Process New Format of Pro Print Start Date[dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankCodeForApplyDocumentPaymentOFAssetLost, clsFixedParameterCode.BankCodeForApplyDocumentPaymentOFAssetLost, "", "Bank Code For Apply Document Payment Entry of Asset Lost")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeprateDistanceMorningEveningShift, clsFixedParameterCode.SeprateDistanceMorningEveningShift, "0", "0-OFF;1-ON Milk Route Show Morning Evening Distance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TankerDispatchProvisionLocationSegment, clsFixedParameterCode.TankerDispatchProvisionLocationSegment, "", "location Segment For Provision JE of Tanker Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncludeInceAndDedInFATSNFRate, clsFixedParameterCode.IncludeInceAndDedInFATSNFRate, "0", "0-OFF;1-ON Bulk Milk Add Incentive and Deduction in FAT SNF Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATPerShouldBeMultipleOf5, clsFixedParameterCode.FATPerShouldBeMultipleOf5, "1", "0-OFF;1-ON Tankder Dispatch FAT % decimal part multiple of 5")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ALLOWBOOKINGSHITWISE, clsFixedParameterCode.ALLOWBOOKINGSHITWISE, "0", "ALLOW BOOKING SHIT WISE MORNING/EVENING")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FindReasonWhyInvoiceIssueOccursOnErode, clsFixedParameterCode.FindReasonWhyInvoiceIssueOccursOnErode, "0", "Find Reason Why Invoice Issue Occurs On Erode")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickFATSNFKgFromInventory, clsFixedParameterCode.PickFATSNFKgFromInventory, "0", "0-OFF(From Milk SRN);1-ON(Pick From Inventory)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyGovtRulesInTDS, clsFixedParameterCode.ApplyGovtRulesInTDS, "0", "0-OFF;1-ON Apply Govt Rules In TDS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoChangeFatANdSNFPerforHighClassVendorinGE, clsFixedParameterCode.AllowtoChangeFatANdSNFPerforHighClassVendorinGE, "0", "Allow to Change Fat & SNF Per for High Class Vendor in GE in case of Contractor after posting")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncetiveEntryApplyArrear, clsFixedParameterCode.IncetiveEntryApplyArrear, "0", "0-OFF;1-ON Apply Incentive and Deduction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VSPDayWiseIncentiveAtSRN, clsFixedParameterCode.VSPDayWiseIncentiveAtSRN, "0", "0-OFF;1-ON Calculate VSP Day Wise Incentive At Milk SRN")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncentiveProcessPaymentStartDate, clsFixedParameterCode.IncentiveProcessPaymentStartDate, "", "Start Date[dd/MMM/yyyy] for make Payment from incentive entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJVOFRGPNRGPAndItsRetrun, clsFixedParameterCode.CreateJVOFRGPNRGPAndItsRetrun, "0", "0-OFF;1-ON Create JV OF RGP/NRGP And Its Return")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BennyImportAutoCreateMP, clsFixedParameterCode.BennyImportAutoCreateMP, "0", "0-OFF;1-ON Auto Create MP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BennyImportPickRateFromPrice, clsFixedParameterCode.BennyImportPickRateFromPrice, "0", "0-OFF;1-ON Pick Rate From FAT/SNF Uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FixedAssetAcquisitionEntryHitInventoryMovement, clsFixedParameterCode.FixedAssetAcquisitionEntryHitInventoryMovement, "0", "0-OFF;1-ON Hit Inventory Movement In Assets Acquisition")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateDeductionByStdQtyinBulkMilkSRN, clsFixedParameterCode.CalculateDeductionByStdQtyinBulkMilkSRN, "1", "0-OFF;1-ON Calculate Deduction By [1] Std Qty [0] Net Weight")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAPinvoiceofsalaryemployeewiseduringsalarygen, clsFixedParameterCode.CreateAPinvoiceofsalaryemployeewiseduringsalarygen, "0", "Create AP invoice of salary employee wise during salary generation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTaxInBulkMilkPurchaseInvoice, clsFixedParameterCode.ApplyTaxInBulkMilkPurchaseInvoice, "0", "0-OFF;1-ON Apply Tax in Bulk Milk Purchase Invoice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkSRNFATSNFDecimalPlaces, clsFixedParameterCode.MilkSRNFATSNFDecimalPlaces, "3", "FAT SNF KG Decimal Places [1-3]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TankerDispatchIntermittentSingleGateIn, clsFixedParameterCode.TankerDispatchIntermittentSingleGateIn, "0", "0-OFF;1-ON Make Single Gate In for Intermittent type Tanker Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.JobWorkOutwardComsumeItemAccordingToBOM, clsFixedParameterCode.JobWorkOutwardComsumeItemAccordingToBOM, "1", "0-OFF;1-ON Consume Items accoding to BOM")
        InsertDefaultValueFixedParameter(clsFixedParameterType.JWOutWardForRCDF, clsFixedParameterCode.JWOutWardForRCDF, "0", "0-OFF;1-ON ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StartDateforDispatchSchedular, clsFixedParameterCode.StartDateforDispatchSchedular, clsCommon.GetPrintDate(clsCommon.GETSERVERDATE().AddDays(-1), "dd/MMM/yyyy"), "StartDateforDispatchSchedular format should be dd/MMM/yyyy")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StockRecoRateTunning, clsFixedParameterCode.StockRecoRateTunning, "0.11", "Stock Reco Rate Tunning")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankRecoCheckFutureDocuments, clsFixedParameterCode.BankRecoCheckFutureDocuments, "1", "0-OFF;1-ON On Reverse and Unpost Check future Transaction exists")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFRate, clsFixedParameterCode.FATSNFRate, "0", "0")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateARAdjAPDebitnoteforEmployeesinMCCMS, clsFixedParameterCode.CreateARAdjAPDebitnoteforEmployeesinMCCMS, "0", "Create AP Debit Note & AR Adjustment for Employees,Chilling Vendor & PTM in MCC Material Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCardSaleInvoiceOnlyWithCardSaleAdvance, clsFixedParameterCode.ApplyCardSaleInvoiceOnlyWithCardSaleAdvance, "0", "ApplyCardSaleInvoiceOnlyWithCardSaleAdvance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionRemoveFATSNFKgTollerance, clsFixedParameterCode.ProductionRemoveFATSNFKgTollerance, "0", "0-OFF;Tollerance % age of Remove FAT/SNF KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubLocationForTaxableItemDairyDispatch, clsFixedParameterCode.SubLocationForTaxableItemDairyDispatch, "", "SubLocationForTaxableItemDairyDispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubLocationForNonTaxableItemDairyDispatch, clsFixedParameterCode.SubLocationForNonTaxableItemDairyDispatch, "", "SubLocationForNonTaxableItemDairyDispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionCheckFATKg, clsFixedParameterCode.ProductionCheckFATKg, "0", "0-OFF;Check Net FAT KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionCheckSNFKg, clsFixedParameterCode.ProductionCheckSNFKg, "0", "0-OFF;Check Net SNF KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo, clsFixedParameterCode.CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo, "0", "CalculateTurnOverForTCS_CustomerBasedOnCommonPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableTCSRateValidityFrom01July2021, clsFixedParameterCode.EnableTCSRateValidityFrom01July2021, "0", "EnableTCSRateValidityFrom01July2021")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionOnlyOneIssueAgainstBatch, clsFixedParameterCode.ProductionOnlyOneIssueAgainstBatch, "0", "0-OFF;1-ON Dairy Production Apply Only One Issue Against Batch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkShiftEndAutoAdjInItemCode, clsFixedParameterCode.MilkShiftEndAutoAdjInItemCode, "", "Milk Shift End Auto Adjustment In Item Code")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StockCheckOnPostForDairyDispatchMultiple, clsFixedParameterCode.StockCheckOnPostForDairyDispatchMultiple, "0", "StockCheckOnPostForDairyDispatchMultiple")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BatchFileCounter, clsFixedParameterCode.BatchFileCounter, clsCommon.EncryptString("3"), "BatchFileCounter")
        InsertDefaultValueFixedParameter(clsFixedParameterType.checkstockMRPwise, clsFixedParameterCode.checkstockMRPwise, "0", "checkstockMRPwise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, "0", "TCSRateforCustomerWithPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, "0", "TCSRateforCustomerWithoutPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ScrapTCSRateforCustomerWithoutPanNo, clsFixedParameterCode.ScrapTCSRateforCustomerWithoutPanNo, "0", "ScrapTCSRateforCustomerWithoutPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, "0", "TCSRateforCustomerWithoutPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ScrapTCSRateforCustomerWithPanNo, clsFixedParameterCode.ScrapTCSRateforCustomerWithPanNo, "0", "ScrapTCSRateforCustomerWithPanNo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSTaxApplicableOnbulkSale, clsFixedParameterCode.TCSTaxApplicableOnbulkSale, "0", "TCSTaxApplicableOnbulkSale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSTaxApplicableOnCanSale, clsFixedParameterCode.TCSTaxApplicableOnCanSale, "0", "TCSTaxApplicableOnCanSale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoChangeTCSBaseAmount, clsFixedParameterCode.AllowtoChangeTCSBaseAmount, "0", "AllowtoChangeTCSBaseAmount")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, "mastergst", "EInvoiceVendor")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TokenTimeReGenForEinvoice, clsFixedParameterCode.TokenTimeReGenForEinvoice, "360", "Token Time ReGen For Einvoice (in mins)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TokenTimeReGenForEWayBill, clsFixedParameterCode.TokenTimeReGenForEWayBill, "360", "Token Time ReGen For Einvoice (in mins)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemCostZeroOnStoreAdjForTypeFlushing, clsFixedParameterCode.ItemCostZeroOnStoreAdjForTypeFlushing, "0", "ItemCostZeroOnStoreAdjForTypeFlushing")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AmountToCheckCustomerOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckCustomerOutstandingForTCSTax, "0", "AmountToCheckCustomerOutstandingForTCSTax")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoCreateSaleInvoice, clsFixedParameterCode.AutoCreateSaleInvoice, "1", "Auto Create Sale Invoice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DonotConsiderDirectARInvoiceinCustomerOutTCS, clsFixedParameterCode.DonotConsiderDirectARInvoiceinCustomerOutTCS, "0", "Do not Consider Direct AR Invoice in Customer Outstanding for TCS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderPreviousCurrentFYForTCSTaxCustOutstanding, clsFixedParameterCode.ConsiderPreviousCurrentFYForTCSTaxCustOutstanding, "0", "ConsiderPreviousCurrentFYForTCSTaxCustOutstanding")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderUnpostedDocForBalance, clsFixedParameterCode.ConsiderUnpostedDocForBalance, "1", "0-OFF;1-ON In Balance Consider Unposted Documents")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowJEofDifferentLocationOnJournalEntry, clsFixedParameterCode.AllowJEofDifferentLocationOnJournalEntry, "0", "Allow JE of Different Location On Journal Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TankerDispatchAvgFATSNFPer, clsFixedParameterCode.TankerDispatchAvgFATSNFPer, "0", "0-OFF;1-ON Pick Average FAT/SNF % in Tanker Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisableToPickMainLocationType, clsFixedParameterCode.DisableToPickMainLocationType, "0", "0-OFF;1-Disable To Pick Main Location Type")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MultipleMCCFinder, clsFixedParameterCode.MultipleMCCFinder, "0", "0-OFF;1-ON Multiple MCC Finder In VSP Bill Generation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProcurementSNF2DecimalPlaces, clsFixedParameterCode.MilkProcurementSNF2DecimalPlaces, "0", "0-OFF;1-ON SNF Calculation in 2 Decimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SelectMilkRejectDefaulterManually, clsFixedParameterCode.SelectMilkRejectDefaulterManually, "0", "0-OFF;1-ON Milk Reject Manually Select Defaulter(ONLY FOR UPLOADER)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DayWiseCustomerIncentiveCalculation, clsFixedParameterCode.DayWiseCustomerIncentiveCalculation, "0", "0-OFF;1-ON Day wise calcualte Customer Incentive")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerIncetiveAutoSecuity, clsFixedParameterCode.CustomerIncetiveAutoSecuity, "0", "0-OFF;1-ON Auto Secuity")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerIncetiveBankForSecuity, clsFixedParameterCode.CustomerIncetiveBankForSecuity, "", "0-OFF;1-ON Auto Secuity Receipt Bank")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerIncetivePaymentModeForSecuity, clsFixedParameterCode.CustomerIncetivePaymentModeForSecuity, "", "0-OFF;1-ON Auto Secuity Receipt Payment Mode")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RepeatDeductionAndVendor, clsFixedParameterCode.RepeatDeductionAndVendor, "0", "0-OFF;1-ON Repeat Same Vendor and Deduction ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LastMilkReceiptQtyTollerance, clsFixedParameterCode.LastMilkReceiptQtyTollerance, "0", "0-OFF;Above Will +/- Tollerance of Last Milk Receipt(ONLY FOR UPLOADER)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowMCCFinderInPaymentProcess, clsFixedParameterCode.ShowMCCFinderInPaymentProcess, "0", "0-OFF;1-ON Show MCC Finder In Payment Process")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseDescInsteadOFCodeOnMCCMAterialSale, clsFixedParameterCode.UseDescInsteadOFCodeOnMCCMAterialSale, "0", "0-OFF;1-ON UseDescInsteadOFCodeOnMCCMAterialSale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowItemConversionAutomation, clsFixedParameterCode.AllowItemConversionAutomation, "0", "0-OFF;1-COrrection factow will come from item master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowShelfLifeMandatoryOnFG, clsFixedParameterCode.AllowShelfLifeMandatoryOnFG, "0", "0-OFF;1-Self life Manadatory on Item Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RefundknockoffwithCreditNote, clsFixedParameterCode.RefundknockoffwithCreditNote, "0", "0-OFF;1-Refund knockoff with Credit Note")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateInvoiceAutomaticallyOnbulkDispatch, clsFixedParameterCode.CreateInvoiceAutomaticallyOnbulkDispatch, "0", "0-OFF;1-CreateInvoiceAutomaticallyOnbulkDispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateReceivingWithMultipleRoute, clsFixedParameterCode.CrateReceivingWithMultipleRoute, "0", "0-OFF;1-CrateReceivingWithMultipleRoute")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowMulMRPOfSameItemOnDairyBookingCustomer, clsFixedParameterCode.ShowMulMRPOfSameItemOnDairyBookingCustomer, "0", "0-OFF;1-Show Multiple MRP Of Same Item On Dairy Booking Customer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowZeroQtyOnDairyBooking, clsFixedParameterCode.AllowZeroQtyOnDairyBooking, "0", "Allow Zero Qty On Dairy Booking Customer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseCutOffTimeonRouteForERP, clsFixedParameterCode.UseCutOffTimeonRouteForERP, "0", "Use Cut Off Time on Route For ERP on Demand Booking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowZeroQtyOnDairyBookingUploader, clsFixedParameterCode.AllowZeroQtyOnDairyBookingUploader, "0", "Allow Zero Qty On Dairy Booking Uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoPostNoOFDocofDOatatime, clsFixedParameterCode.AllowtoPostNoOFDocofDOatatime, "0", "AllowtoPostNoOFDocofDOatatime on Dispatch Multiple screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DonotIncludeSecurityInCustomerOutstanding, clsFixedParameterCode.DonotIncludeSecurityInCustomerOutstanding, "0", "0-OFF;1-Do not Include Security In Customer Outstanding On Dairy Booking Customer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsManualTCS, clsFixedParameterCode.IsManualTCS, "0", "0-OFF;1-Apply Manual TCS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultLocationForCardSaleIntegration, clsFixedParameterCode.DefaultLocationForCardSaleIntegration, "", "Default Location For Card Sale Integration")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, clsFixedParameterCode.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, "0", "0-OFF;1-ApplyNoGSTCreditIndependentlyOnVendorServiceCharge")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckNoOfDaysforCardSaleBooking, clsFixedParameterCode.CheckNoOfDaysforCardSaleBooking, "1", "CheckNoOfDaysforCardSaleBooking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxNoOfBookingAllowedThroughBookingApp, clsFixedParameterCode.MaxNoOfBookingAllowedThroughBookingApp, "2", "it should be a number MaxNoOfBookingAllowedThroughBookingApp")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBothtsrateAndFatRateinBulkProcurement, clsFixedParameterCode.ApplyBothtsrateAndFatRateinBulkProcurement, "0", "it should be a number ApplyBothtsrateAndFatRateinBulkProcurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, "0", "Pick Transfer CLearing A/c For Stock Transfer And Tanker Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowBookingTypeDropDownonDairyBookingCustomer, clsFixedParameterCode.ShowBookingTypeDropDownonDairyBookingCustomer, "0", "0-OFF;1- Booking Type Drop Down (Cash/CR/Festive Offer/SO)")
        'InsertDefaultValueFixedParameter(clsFixedParameterType.AutoCreateGateEntryTillMilkTransferInForIntermittent, clsFixedParameterCode.AutoCreateGateEntryTillMilkTransferInForIntermittent, "0", "AutoCreateGateEntryTillMilkTransferInForIntermittent")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoMilkTransferInDateSameasWeighmentDate, clsFixedParameterCode.AutoMilkTransferInDateSameasWeighmentDate, "0", "0-OFF;1-COrrection factow will come from item master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt, clsFixedParameterCode.EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt, "0", "0-OFF;1-EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOutstandingAmtofCustomerOnQuickBookEntry, clsFixedParameterCode.ShowOutstandingAmtofCustomerOnQuickBookEntry, "0", "0-OFF;1-ShowOutstandingAmtofCustomerOnQuickBookEntry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotCreateJournalVoucheronJobWorkDispatch, clsFixedParameterCode.DoNotCreateJournalVoucheronJobWorkDispatch, "0", "0-OFF;1-COrrection factow will come from item master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickCostOFMaterialSaleFromPriceMaster, clsFixedParameterCode.PickCostOFMaterialSaleFromPriceMaster, "0", "0-OFF;1-PickCostOFMaterialSaleFromPriceMaster.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemwiseCorrectionFactoronQC, clsFixedParameterCode.ItemwiseCorrectionFactoronQC, "0", "0-OFF;1-COrrection factow will come from item master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow0FatPerOnBulkSaleQualityCheckScreen, clsFixedParameterCode.Allow0FatPerOnBulkSaleQualityCheckScreen, "0", "Allow 0 Fat Per On Bulk Sale Quality Check Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateMultipleDispatchWithoutSelectingVehicle, clsFixedParameterCode.CreateMultipleDispatchWithoutSelectingVehicle, "0", "CreateMultipleDispatchWithoutSelectingVehicle")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RejectiononQCforSeparationofBulkProcurementMCC, clsFixedParameterCode.RejectiononQCforSeparationofBulkProcurementMCC, "0", "RejectiononQCforSeparationofBulkProcurementMCC")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ParameterForSNFatQC, clsFixedParameterCode.ParameterForSNFatQC, "0", "0-OFF;1-Parameter used in calculation of SNF%")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatCLR50As52, clsFixedParameterCode.TreatCLR50As52, "0", "0-OFF;1-When convert CLR To SNF than take .50 as .52")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowmultipleconsumptionLocation, clsFixedParameterCode.AllowmultipleconsumptionLocation, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickCostFromItemMaster, clsFixedParameterCode.PickCostFromItemMaster, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EditItemCost, clsFixedParameterCode.EditItemCost, "1", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBatchQtyin3DecimalPlaces, clsFixedParameterCode.AllowBatchQtyin3DecimalPlaces, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayAverageFatSNFMPWise, clsFixedParameterCode.DisplayAverageFatSNFMPWise, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SentschemecogsinAnotherAccount, clsFixedParameterCode.SentschemecogsinAnotherAccount, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OPkmMandatoryonDS, clsFixedParameterCode.OPkmMandatoryonDS, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBankChargesasperSlabonBankMaster, clsFixedParameterCode.ApplyBankChargesasperSlabonBankMaster, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseKGLitreConversionInBulkSaleAsperCLRCalculation, clsFixedParameterCode.UseKGLitreConversionInBulkSaleAsperCLRCalculation, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowLastUnitCostZeroForNonInventoryItemOnPO, clsFixedParameterCode.ShowLastUnitCostZeroForNonInventoryItemOnPO, "0", "1-On;0-off Show Last Unit Cost Zero(last Vendor or Same Vendor) For Non Inventory Item On PO")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ManualBatchNoMandatoryOnBatchOrderScreen, clsFixedParameterCode.ManualBatchNoMandatoryOnBatchOrderScreen, "0", "1-On;0-off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ForceToSelectIteminGateEntry, clsFixedParameterCode.ForceToSelectIteminGateEntry, "0", "0-OFF;1-Adjustment not created on Milk transfer In posting.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionforBulkContractorInGateIn, clsFixedParameterCode.CreateProvisionforBulkContractorInGateIn, "0", "0-OFF;1-CreateProvisionforBulkContractorInGateIn.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSiloLocationItemLocationwise, clsFixedParameterCode.ShowSiloLocationItemLocationwise, "0", "0-OFF;1-Silo location mapped in item location mappig screen in finder")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotCreateAdjustmentonMilkTransferInGL, clsFixedParameterCode.DoNotCreateAdjustmentonMilkTransferInGL, "0", "0-OFF;1-Adjustment not created on Milk transfer In posting.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Donotshowtrasnfertransactionsbydefault, clsFixedParameterCode.Donotshowtrasnfertransactionsbydefault, "1", "0-OFF;1-Do not show transfer transactions by default on Mis Sale Register/ Sale Register Detail")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualItemPriceOnMCCSale, clsFixedParameterCode.AllowManualItemPriceOnMCCSale, "0", "0-OFF;1-Allow Manual Item Price On MCC Material Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ExcludeKKFAndMandiForDCS, clsFixedParameterCode.ExcludeKKFAndMandiForDCS, "0", "0-OFF;1-Exclude KKF And Mandi For DCS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DonotAllowtoChangeUOMinDairyBookingCustomer, clsFixedParameterCode.DonotAllowtoChangeUOMinDairyBookingCustomer, "0", "0-OFF;1-DonotAllowtoChangeUOMinDairyBookingCustomer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoPopulateItemCodeOnDairyBooking, clsFixedParameterCode.AutoPopulateItemCodeOnDairyBooking, "0", "0-OFF;1-AutoPopulateItemCodeOnDairyBooking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GateEntryChamberwisewithManualTankerEntry, clsFixedParameterCode.GateEntryChamberwisewithManualTankerEntry, "0", "0-OFF;1-On Option will appear on Dairy Dispatch.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPriceMappingOnBulkSRNinChamberBulkProc, clsFixedParameterCode.AllowPriceMappingOnBulkSRNinChamberBulkProc, "0", "0-OFF;1-On Option will appear on Dairy Dispatch.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableManualCrateonTaxableDairyDispatch, clsFixedParameterCode.EnableManualCrateonTaxableDairyDispatch, "0", "0-OFF;1-On Option will appear on Dairy Dispatch.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoCalculateCANOnDairyDispatch, clsFixedParameterCode.AutoCalculateCANOnDairyDispatch, "0", "0-OFF;1-On Option will appear on Dairy Dispatch.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSiloMilkTransfertoMainLocation, clsFixedParameterCode.AllowSiloMilkTransfertoMainLocation, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableAutoDocNoShipToLocation, clsFixedParameterCode.EnableAutoDocNoShipToLocation, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ChangeFATCLRafterspecialApprovalonQC, clsFixedParameterCode.ChangeFATCLRafterspecialApprovalonQC, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableCustomerPODetailonDairyBooking, clsFixedParameterCode.EnableCustomerPODetailonDairyBooking, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateCommonDairyDispatchforFreshAmbient, clsFixedParameterCode.CreateCommonDairyDispatchforFreshAmbient, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateSeperateTaxInvForFOCIteminNonTaxdispatch, clsFixedParameterCode.CreateSeperateTaxInvForFOCIteminNonTaxdispatch, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateSeperateSeriesforRefDocARinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocARinvforCreditdebit, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateSeperateSeriesforRefDocAPinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocAPinvforCreditdebit, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotConsiderCustomerCreditLimit, clsFixedParameterCode.DoNotConsiderCustomerCreditLimit, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowVSPMasterAutoPrefix, clsFixedParameterCode.AllowVSPMasterAutoPrefix, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableGSTRelatedfields, clsFixedParameterCode.EnableGSTRelatedfields, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPostGSTPayment, clsFixedParameterCode.AllowPostGSTPayment, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderSiloCapicityForStockIn, clsFixedParameterCode.ConsiderSiloCapicityForStockIn, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDifferentStateofChildCustomerOnPS, clsFixedParameterCode.AllowDifferentStateofChildCustomerOnPS, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowProvisionknokoffOnAPInvoice, clsFixedParameterCode.AllowProvisionknokoffOnAPInvoice, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowTransactionFiltersOnCustomerlegder, clsFixedParameterCode.AllowTransactionFiltersOnCustomerlegder, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GroupCustomerlegderZoneWiseAreaWise, clsFixedParameterCode.GroupCustomerlegderZoneWiseAreaWise, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerDashboardWithOpeningAndClosing, clsFixedParameterCode.CustomerDashboardWithOpeningAndClosing, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSHOWParentChildCustomer, clsFixedParameterCode.AllowtoSHOWParentChildCustomer, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoMakeApplyDocOnbyDefault, clsFixedParameterCode.AllowtoMakeApplyDocOnbyDefault, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PenaltyPercentage, clsFixedParameterCode.PenaltyPercentage, "0", "Enter % for Penalty calculation ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionJournalEntryForSale, clsFixedParameterCode.CreateProvisionJournalEntryForSale, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionJournalEntryForTankerDispatch, clsFixedParameterCode.CreateProvisionJournalEntryForTankerDispatch, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowDairySaleModuleOnBulkPosting, clsFixedParameterCode.ShowDairySaleModuleOnBulkPosting, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoBulkMilkSRNonWeighmentBulkProc, clsFixedParameterCode.AllowAutoBulkMilkSRNonWeighmentBulkProc, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowGateEntryTypeonGateEntryBulkProc, clsFixedParameterCode.ShowGateEntryTypeonGateEntryBulkProc, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickCorrectionFactorProcurementTypewise, clsFixedParameterCode.PickCorrectionFactorProcurementTypewise, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckParameterRangerProcurementTypewise, clsFixedParameterCode.CheckParameterRangerProcurementTypewise, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkProcMCCwithoutTankerDispatch, clsFixedParameterCode.AllowBulkProcMCCwithoutTankerDispatch, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowReverseUnpost, clsFixedParameterCode.AllowReverseUnpost, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoShowCreditBalanceonCustomerAgeing, clsFixedParameterCode.AllowtoShowCreditBalanceonCustomerAgeing, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoShowDebitBalanceonVendorAgeing, clsFixedParameterCode.AllowtoShowDebitBalanceonVendorAgeing, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderOpeningDocintoBucketsInAgeing, clsFixedParameterCode.ConsiderOpeningDocintoBucketsInAgeing, "0", "ConsiderOpeningDocintoBucketsInAgeing")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderOpeningDocintoBucketsonCustomerAgeing, clsFixedParameterCode.ConsiderOpeningDocintoBucketsonCustomerAgeing, "0", "ConsiderOpeningDocintoBucketsonCustomerAgeing")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSetNoOfTransactionsforSetOff, clsFixedParameterCode.AllowtoSetNoOfTransactionsforSetOff, "100", "maximum No of transactions required for set off at a time")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoUnlockTransactionsforSetOff, clsFixedParameterCode.AllowtoUnlockTransactionsforSetOff, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSetReceiptAmountForCashTransaction, clsFixedParameterCode.AllowtoSetReceiptAmountForCashTransaction, "0", "max Receipt amount for cash transactions")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSetPaymentAmountForCashTransaction, clsFixedParameterCode.AllowtoSetPaymentAmountForCashTransaction, "0", "max Payment amount for cash transactions")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSetoffDocDateWise, clsFixedParameterCode.AllowtoSetoffDocDateWise, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSkipfunctionalityafterSRNOnBulkProcurement, clsFixedParameterCode.AllowtoSkipfunctionalityafterSRNOnBulkProcurement, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoNegativeStockInventoryAtTankerDispatch, clsFixedParameterCode.AllowtoNegativeStockInventoryAtTankerDispatch, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoNegativeFATSNFKgAtTankerDispatch, clsFixedParameterCode.AllowtoNegativeFATSNFKgAtTankerDispatch, "0", "0-OFF;1-On Option Check Negative FAT/SNF Kg ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoEmployeeSalaryIntegration, clsFixedParameterCode.AllowtoEmployeeSalaryIntegration, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoFutureDateTransForPDCCheque, clsFixedParameterCode.AllowtoFutureDateTransForPDCCheque, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.WeightOfCanForCanSale, clsFixedParameterCode.WeightOfCanForCanSale, "", "enter weight of a can it is considered as in KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualPriceONBulkPO, clsFixedParameterCode.AllowManualPriceONBulkPO, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunBulkProcWithoutMilkGrade, clsFixedParameterCode.RunBulkProcWithoutMilkGrade, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DaysToStartAutoLock, clsFixedParameterCode.DaysToStartAutoLock, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRandomOnlyOneSecondaryQC, clsFixedParameterCode.AllowRandomOnlyOneSecondaryQC, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowGateEntryAgainstPO, clsFixedParameterCode.AllowGateEntryAgainstPO, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeparateDairyDispatchTaxableNonTaxable, clsFixedParameterCode.SeparateDairyDispatchTaxableNonTaxable, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PromptTimeToPostTransactions, clsFixedParameterCode.PromptTimeToPostTransactions, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoLockTransaction, clsFixedParameterCode.AllowAutoLockTransaction, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, "", "bank code used on Receipt and Payment Entry Screen for Apply Document.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowCreditNoteWithoutReference, clsFixedParameterCode.AllowCreditNoteWithoutReference, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowUseApplyDocSeriesForReceipt, clsFixedParameterCode.AllowUseApplyDocSeriesForReceipt, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowUseApplyDocSeriesForPayment, clsFixedParameterCode.AllowUseApplyDocSeriesForPayment, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowCreditNoteWithoutReferenceonAP, clsFixedParameterCode.AllowCreditNoteWithoutReferenceonAP, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBranchAcconReceiptPrint, clsFixedParameterCode.AllowBranchAcconReceiptPrint, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SecurityDocumentKnockOffonReceipt, clsFixedParameterCode.SecurityDocumentKnockOffonReceipt, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowLockTransactionUserwise, clsFixedParameterCode.AllowLockTransactionUserwise, "0", "0-OFF;1-On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFreshInvoiceAutoPost, clsFixedParameterCode.AllowFreshInvoiceAutoPost, "0", "0-OFF;1-On Option will appear on Frsh Dispatch Multiple.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowReceiptThroughSO, clsFixedParameterCode.AllowReceiptThroughSO, "1", "0-finder open of DO of Product Sale;1-finder open of SO of Product Sale ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSetOffUntilTransactionsnotend, clsFixedParameterCode.AllowSetOffUntilTransactionsnotend, "0", "0-OFF;1-On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowGateReturn, clsFixedParameterCode.AllowGateReturn, "0", "0-OFF;1-On Option will appear on Sale and Transfer.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateCommOnCSATransWOConversion, clsFixedParameterCode.CalculateCommOnCSATransWOConversion, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPurchaseModulewithUniqueItem, clsFixedParameterCode.AllowPurchaseModulewithUniqueItem, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GrossWeightUnit, clsFixedParameterCode.GrossWeightUnit, "Kg", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ExpiryDaysBulkProcurementPriceChart, clsFixedParameterCode.ExpiryDaysBulkProcurementPriceChart, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSchemeItemRateonDairyDispatch, clsFixedParameterCode.ShowSchemeItemRateonDairyDispatch, "0", "0-OFF;1-On Option will appear on Prouct Booking.[Non Taxable]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSchemeItemRateonDairyDispatchTaxable, clsFixedParameterCode.ShowSchemeItemRateonDairyDispatchTaxable, "0", "0-OFF;1-On Option will appear on Prouct Booking.[Taxable]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoCalculateCrateOnDairyDispatch, clsFixedParameterCode.AutoCalculateCrateOnDairyDispatch, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFreshPriceChartOnBookingProductSale, clsFixedParameterCode.AllowFreshPriceChartOnBookingProductSale, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GrossWtFromItemMasterONProductSale, clsFixedParameterCode.GrossWtFromItemMasterONProductSale, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, "0", "0-OFF;1-On Option will appear on Prouct Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowUnloadingandWeighmentSequencewise, clsFixedParameterCode.ShowUnloadingandWeighmentSequencewise, "0", "0-OFF;1-On Option will appear on Unloading.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowBothTankertypeOnCleaning, clsFixedParameterCode.ShowBothTankertypeOnCleaning, "0", "0-OFF;1-Both Option will appear on Cleanig.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isCleaningMandatoryBeforeGateout, clsFixedParameterCode.isCleaningMandatoryBeforeGateout, "0", "0-OFF;1-All Option will appear on Item Master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkProcurementSequencewise, clsFixedParameterCode.AllowBulkProcurementSequencewise, "0", "0-OFF;1-All Option will appear on Item Master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowItemLocationWiseonDairyBooking, clsFixedParameterCode.ShowItemLocationWiseonDairyBooking, "0", "0-OFF;1-All Option will appear on Item Master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeprateDemandForMorningEveningShift, clsFixedParameterCode.SeprateDemandForMorningEveningShift, "0", "0-OFF;1-Create Seprate Demand For Morning Evening Shift")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckOutstandingCreditLimitOnBooking, clsFixedParameterCode.CheckOutstandingCreditLimitOnBooking, "0", "0-OFF;1-Check Customer Outstanding on Booking.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkPriceChartMultiplepriceToMultipleVendor, clsFixedParameterCode.AllowBulkPriceChartMultiplepriceToMultipleVendor, "0", "0-OFF;1-All Option will appear on Item Master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOptionOnItemMasterChangeItemRate, clsFixedParameterCode.ShowOptionOnItemMasterChangeItemRate, "0", "0-OFF;1-All Option will appear on Item Master.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SHowOptionOnLocationForDairyDispatchfromDOorGatepass, clsFixedParameterCode.SHowOptionOnLocationForDairyDispatchfromDOorGatepass, "0", "0-OFF;1-All Option will appear on Location screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, "1", "1-ON;0-OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDocumentDate, clsFixedParameterCode.ApplyDocumentDate, "0", "1-ON;0-OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowStockCheckatDOLevel, clsFixedParameterCode.AllowStockCheckatDOLevel, "0", "1-ON;0-OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAdditionalWeightinPercentage, clsFixedParameterCode.AllowAdditionalWeightinPercentage, "0", "ON- additional weight considered as in % ;OFF - additional weight considered as value")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnterAdditionalWeight, clsFixedParameterCode.EnterAdditionalWeight, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowTankerBasedonVendorofGE, clsFixedParameterCode.AllowTankerBasedonVendorofGE, "0", "1-ON;0-OFF Applicable for Route Wise Vendor in Bulk Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowUseBoilingParameteronParameterMaster, clsFixedParameterCode.AllowUseBoilingParameteronParameterMaster, "0", "1-ON;0-OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DairyDispatchFromDeliveryNote, clsFixedParameterCode.DairyDispatchFromDeliveryNote, "0", "0-OFF;1-All Menu will appear on screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemTypeForDairyBooking, clsFixedParameterCode.ItemTypeForDairyBooking, "B", "B-Both;F-Fresh A-Ambient Menu will appear on screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, "0", "0-OFF;1-All Menu will appear on screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StockToleranceLimit, clsFixedParameterCode.StockToleranceLimit, "0", "Enter Tolerance limit")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllMenu, clsFixedParameterCode.ShowAllMenu, "0", "0-OFF;1-All Menu will appear on screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OpenAvailorEmptyStckLocationOn_Standardization, clsFixedParameterCode.OpenAvailorEmptyStckLocationOn_Standardization, "0", "0-OFF;1-On then selected item available or empty location is open in finder on Production Standardization screen.") 'BM00000008148
        InsertDefaultValueFixedParameter(clsFixedParameterType.WorkApprovalFlowInERP, clsFixedParameterCode.WorkApprovalFlowInERP, "0", "0-OFF;1-On then Work approval flow start in ERP.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BOM_Amend_Pswd, clsFixedParameterCode.BOM_Amend_Pswd, "Tecxpert2015", "Set password for BOM amendment.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionFATSNF_KG_Unit, clsFixedParameterCode.ProductionFATSNF_KG_Unit, "Kg", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RemoveGridFatSnfValidation, clsFixedParameterCode.RemoveGridFatSnfValidation, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ChangeRateAT_CSA_Return, clsFixedParameterCode.ChangeRateAT_CSA_Return, "0", "0-OFF;1-On for allow rate change at csa sale return screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VehicleCapacityUnit, clsFixedParameterCode.VehicleCapacityUnit, "MT", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopGLEntryForConsignmentAtCSATransfer, clsFixedParameterCode.StopGLEntryForConsignmentAtCSATransfer, "0", "0-Off;1-On for stop debit consignment entry at CSA Transfer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CSATransfer_SalePatti_All_Tax_Open, clsFixedParameterCode.CSATransfer_SalePatti_All_Tax_Open, "1", "0-OFF;1-On for Open All Tax Mapped With Location In CSA Without State Condition.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Active, clsFixedParameterCode.Emps1, "employee status", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Agency, clsFixedParameterCode.Agency, "Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowNegtiveOfSaleInvoiceBlanceAmt, clsFixedParameterCode.AllowNegtiveOfSaleInvoiceBlanceAmt, "1", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BalanceSheetProftAndLossGroupCode, clsFixedParameterCode.BalanceSheetProftAndLossGroupCode, "999", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BalanceSheetProftAndLossGroupDesc, clsFixedParameterCode.BalanceSheetProftAndLossGroupDesc, "Profit And Loss", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BlankDatabase, clsFixedParameterCode.BlankDatabase, "NAPOLEON", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipDiffGLOnPI, clsFixedParameterCode.SkipDiffGLOnPI, "0", "0: OFF, 1: ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDesignAtRunTime, clsFixedParameterCode.AllowDesignAtRunTime, "0", "0: OFF, 1: ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Category1, clsFixedParameterCode.Category1, "Category", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Category2, clsFixedParameterCode.Category2, "Category", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Category3, clsFixedParameterCode.Category3, "Category", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.cboPeriodicity, clsFixedParameterCode.A, "Annually", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.cboPeriodicity, clsFixedParameterCode.HY, "Half Yearly", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.cboPeriodicity, clsFixedParameterCode.M, "Monthly", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.cboPeriodicity, clsFixedParameterCode.Q, "Quarterly", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CboRound, clsFixedParameterCode.L, "Lower", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CboRound, clsFixedParameterCode.R, "Round", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CboRound, clsFixedParameterCode.U, "Upper", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreatePOWithRequisition, clsFixedParameterCode.POWITHREQ, "0", "0-OFF;1-On: if off then user can create PO without selecting Req No else REq No is mandatory to create PO")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LOReceiptDefaultBankForSettlement, clsFixedParameterCode.LOReceiptDefaultBankForSettlement, "SETTLEMENT", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LOReceiptPaymentTypeForSettlement, clsFixedParameterCode.LOReceiptPaymentTypeForSettlement, "SETTLEMENT", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ALLOWANYBO, clsFixedParameterCode.ALLOWANYBO, "0", "0-OFF;1-On: allow to access BO in any sequence.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ALLOWCBOSBO, clsFixedParameterCode.ALLOWCBOSBO, "0", "0-OFF;1-On: allow to access child and sub-child BO in any sequence,after that Main BO access.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.INDUSTRYTYPE, clsFixedParameterCode.INDUSTRYTYPE, "D", "D-Dairy; A-Auto Mobile; M-Manufacturing; etc. are industries type.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultValue, clsFixedParameterCode.DefaultTypeFB, "5", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultValue, clsFixedParameterCode.DefaultTypeFC, "220", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultValue, clsFixedParameterCode.DefaultTypeSH, "100", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PROVISIONENTRYONSTOCKTRANSFER, clsFixedParameterCode.PROVISIONENTRYONSTOCKTRANSFER, "0", "0-OFF;1-On: allow Provision entry on Stock transfer at post.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayReasonOnUpdateAfterPost, clsFixedParameterCode.DisplayReasonOnUpdateAfterPost, "1", "0-OFF;1-On: if on then Reason will be asked during Update of records else will not asked.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayReasonOnDelete, clsFixedParameterCode.DisplayReasonOnDelete, "0", "0-OFF;1-On: if on then Reason will be asked during delete of records else will not asked.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllLevelApprovalIsMandatory, clsFixedParameterCode.AllLevelApprovalIsMandatory, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Distributor, clsFixedParameterCode.Distributor, "Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnablePopupItemReorderLevel, clsFixedParameterCode.POPUPITEMREORDERLEVEL, "0", "0-OFF;1-On: if off then Popup for item reorder level will not displayed  else displayed during login of erp")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Helper, clsFixedParameterCode.EMP02, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Driver, clsFixedParameterCode.Driver, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZM, clsFixedParameterCode.ZM, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TSM, clsFixedParameterCode.TSM, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ASM, clsFixedParameterCode.ASM, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Inactive, clsFixedParameterCode.Emps2, "Employee Status", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IndentTolerence, clsFixedParameterCode.IndentTolerence, "20", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isBatchApplyOnInventoryMovement, clsFixedParameterCode.isBatchApplyOnInventoryMovement, "0", "0-OFF   1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsConsiderOutTypeDocForBalance, clsFixedParameterCode.IsConsiderOutTypeDocForBalance, "1", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsNotIncludeWasteQtyInCal, clsFixedParameterCode.IsNotIncludeWasteQtyInCal, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Others, clsFixedParameterCode.EMP3, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.ALLOW, "Allowance", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.BASIC, "Basic", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.BONUS, "Bonus", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.COEPS, "Company EPS", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.COESI, "Company ESICompany ESI", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.COPF, "Company PF", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.DA, "DA", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.DEDUCT, "Deduction", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.EMPESI, "Employee ESI", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.EPF, "Employee PF", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.HRA, "HRA", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.HRR, "HRR", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.LOAN, "LOAN", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.OT, "OT", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.OTHER, "Other", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.RMBT, "Reimbursement", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.TA, "TA", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.TDS, "TDS", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.Mediclaim, "Mediclaim", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.LTA, "LTA", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.Gratuity, "Gratuity", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.LeaveEnchashed, "LeaveEnchashed", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.UnpaidAmount, "UnpaidAmount", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.Arrear, "Arrear", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.PT, "Proffesional Tax", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.Conveyance, "Conveyance", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.LaborUnFairFund, "Labor Unfair Fund", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayHeadSubHead, clsFixedParameterCode.LaborWelFairFund, "Labor Welfair Fund", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickMachineDateForTran, clsFixedParameterCode.PickMachineDateForTran, "1", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrintVerify, clsFixedParameterCode.SalesInvoice, "N", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrintVerify, clsFixedParameterCode.Transfer, "n", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PromptForTally, clsFixedParameterCode.o, "PromptForTally", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseOneItemOneVendor, clsFixedParameterCode.PurchaseOneItemOneVendor, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PostShipmentonAutoSTN, clsFixedParameterCode.PostShipmentonAutoSTN, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateInvoicewithShipmentonAutoSTN, clsFixedParameterCode.CreateInvoicewithShipmentonAutoSTN, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsItemRateEditableOnTransfer, clsFixedParameterCode.IsItemRateEditableOnTransfer, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MultiplySubsidyWithQuantity, clsFixedParameterCode.MultiplySubsidyWithQuantity, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAutoReceiptEntryDCSSale, clsFixedParameterCode.CreateAutoReceiptEntryDCSSale, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GeneratePDFForMobile, clsFixedParameterCode.GeneratePDFForMobile, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsTransferQtyEditableOnAutoSTN, clsFixedParameterCode.IsTransferQtyEditableOnAutoSTN, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsItemRateEditableOnSales, clsFixedParameterCode.IsItemRateEditableOnSales, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsItemMRPEditableOnSales, clsFixedParameterCode.IsItemMRPEditableOnSales, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsItemRateEditableOnSalesForAprilOnly, clsFixedParameterCode.IsItemRateEditableOnSalesForAprilOnly, "1", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsRemarksMandatoryOnCloseSaleOrder, clsFixedParameterCode.IsRemarksMandatoryOnCloseSaleOrder, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchasePickItemFromVendorItemDetails, clsFixedParameterCode.PurchasePickItemFromVendorItemDetails, "0", "0-Off 1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PZ, clsFixedParameterCode.PrefixGenerator, "b12sec2", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LCCancellationPwd, clsFixedParameterCode.LCCreationPwd, "xpert", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowQtySum_in_GRN_MRN_SRN, clsFixedParameterCode.ShowQtySum_in_GRN_MRN_SRN, "xpert", "1")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ReqLimitOnSRN, clsFixedParameterCode.ReqLimitOnSRN, "5000", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ROUTE, clsFixedParameterCode.DuplicateRoute, "ADMIN", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Route1, clsFixedParameterCode.RJ, "admin", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunDemoERP, clsFixedParameterCode.RunDemoERP, "1", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsKDIL, clsFixedParameterCode.IsKDIL, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Salesman, clsFixedParameterCode.EMP01, "Employee Type", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SendToTally, clsFixedParameterCode.SendToTally, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ServiceDealer, clsFixedParameterCode.ServiceDealer, "Employee Type", "Employee Type")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIR, clsFixedParameterCode.SIRevers, "salereverse", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.SIReversAndCreate, "b12sec2", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.CustomerTenderAmendment, "Amendment@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.AllowBoothWiseReturn, "BoothWiseReturn@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.SICrateEntry, "CrateEntry@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.ShuffleDemand, "ShuffleDemand@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.ItemAllowDecimal, "AllowDecimal@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.UpdatePassword, "c1032floor", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SIRC, clsFixedParameterCode.GatePassPasswod, "gate123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MulProcDedReversAndCreate, clsFixedParameterCode.MulProcDedReversAndCreate, "ProDeduction", "Password for reverse of Multiple Procurement Deduction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PP_MRP, clsFixedParameterCode.PPMrpReversAndCreate, "b12sec2", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.WEUpdateAfterPost, clsFixedParameterCode.WEUpdateAfterPost, "WEUPDATE", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GEUpdateAfterPost, clsFixedParameterCode.GEUpdateAfterPost, "GEUPDATE", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE, clsFixedParameterCode.PwdAllowtoChangeFatANdSNFPerforHighClassVendorinGE, "admin@11", " password to update fat and snf in case of contractor for gate entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GEUpdatePriceChart, clsFixedParameterCode.GEUpdatePriceChart, "UpdatePrice", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SetCSATransferwithZeroOnSalePatti, clsFixedParameterCode.SetCSATransferwithZeroOnSalePatti, "SETZERO", "")
        ''richa agarwal 09/04/2015 create password for PO Amendment
        'InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.SecretorydcsMasterUpdate, "Secretorydcs@123", "Secretory DCS Master Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ImportorExport, clsFixedParameterCode.FarmerMaster, "Import@123", "Farmer Master Import/Export Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.VendorMaster, "Vendor@123", "Vendor Master Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.VendorIntegratedMaster, "Inte@123", "Vendor Integarted Master Password")

        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.ItemMaster, "Item@123", "Item Master Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.UserMaster, "User@123", "User Master Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.DemandUnpost, "Demand@123", "Demand Password for Unpost")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.EWayBillUpdate, "EwayBill@123", "Dipatch Password for Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.DispatchCancel, "Dispatch@123", "Dipatch Password for Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.BookingCancel, "Booking@123", "Booking Password for Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.GatePassCancel, "GatePass@123", "GatePass Password for Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Transactionupdate, clsFixedParameterCode.CustomerMaster, "Customer@123", "Customer Master Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.POAmendmentType, clsFixedParameterCode.POAmendment, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkInvoiceDeleteType, clsFixedParameterCode.BulkInvoiceDelete, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Importbulkdatafromexcelsheet, clsFixedParameterCode.Importbulkdatafromexcelsheet, "RakeshSharma", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateEInvoicePassword, clsFixedParameterCode.MaterialSales, "EInvoice@123", "Create EInvoice Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkSaleSequence, clsFixedParameterCode.BulkSaleSequence, "0", "0-KDIL; 1-OTHERS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkQCTableHavingUniqueKey, clsFixedParameterCode.BulkQCTableHavingUniqueKey, "0", "0-Pavitra; 1-OTHERS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SrPath, clsFixedParameterCode.SrPath, "\\192.168.100.3\KIWI ERP", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TaxRoundOffToZeroDecimalPlace, clsFixedParameterCode.TaxRoundOffToZeroDecimalPlace, "1", "1-On,0-Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TDM, clsFixedParameterCode.TDM, "Employee Type", "Employee Type")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TempProvisional, clsFixedParameterCode.TempProvisional, "tp", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TF, clsFixedParameterCode.LoadInRollback, "tecx2008", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Z, clsFixedParameterCode.Sunday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZA, clsFixedParameterCode.Monday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZB, clsFixedParameterCode.Tuesday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZC, clsFixedParameterCode.Wednesday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZD, clsFixedParameterCode.Thursday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZE, clsFixedParameterCode.Friday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZF, clsFixedParameterCode.Saturday, "Route Off Day", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.EnableMilkProc, "0", "1-Enable Milk Procurement Module; 0-Disable Milk Procurement Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowSkippingPrevDocumentsOnPaymentProcess, "0", "1-ON; 0-Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateGLForTransfer, clsFixedParameterCode.CreateGLForTransfer, "0", "1-ON; 0-Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SalesmanPhysicalLocation, clsFixedParameterCode.SalesmanPhysicalLocation, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AskForDate, clsFixedParameterCode.AskForDate, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoLoadinFromLocation, clsFixedParameterCode.AutoLoadinFromLocation, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrreateTransferShipmentJE, clsFixedParameterCode.CrreateTransferShipmentJE, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankTransferRunPaymentCounter, clsFixedParameterCode.BankTransferRunPaymentCounter, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PaymentReceiptTypeRunReceiptCounter, clsFixedParameterCode.PaymentReceiptTypeRunReceiptCounter, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CounterFinancialYearStyle, clsFixedParameterCode.CounterFinancialYearStyle, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LinkFinancialYearStyleWithGSTDate, clsFixedParameterCode.LinkFinancialYearStyleWithGSTDate, "", "OFF: [FY Format 17-18].ON: [FY Format 17] Applicable After GST Start Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyLocationWisePrefix, clsFixedParameterCode.ApplyLocationWisePrefix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CashDiscountFromClaimMaster, clsFixedParameterCode.CashDiscountFromClaimMaster, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransferTransTypeRouteHide, clsFixedParameterCode.TransferTransTypeRouteHide, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CurrentMaufacturingType, clsFixedParameterCode.CurrentMaufacturingType, "", "D-DESCRETE; P-PROCESS; R-RICE")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TallyCompany, clsFixedParameterCode.TallyCompany, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TallyIP, clsFixedParameterCode.TallyIP, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TallyPort, clsFixedParameterCode.TallyPort, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MAILOFF, clsFixedParameterCode.MAILOFF, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DepreciationCalculationMethod, clsFixedParameterCode.DepreciationCalculationMethod, "MT", "DL-Daily;MT-Monthly;QTL-Quarterly;YRY-Yearly")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AssetGroupPrefix, clsFixedParameterCode.AssetGroupPrefix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.STDPURRATE, clsFixedParameterCode.STDPURRATE, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoPOAtSRN, clsFixedParameterCode.AutoPOAtSRN, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisableShipToLocation, clsFixedParameterCode.DisableShipToLocation, "", "If Setting Is On Then Ship To Location Will be Inactive for Purchase Order, Purchase Invoice, SRN Otherwise Active")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSingleInvoiceAgainstSingleOrder, clsFixedParameterCode.AllowSingleInvoiceAgainstSingleOrder, "0", "If This Setting Is On Then Only Single Sale Invoice can be Created against Single Sale Order")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowLargerItemCostThenVendorItemCost, clsFixedParameterCode.AllowLargerItemCostThenVendorItemCost, "", "If Setting Is On Then Item Price Will be accepted larger then vendor Item Price for  Purchase Invoice and SRN .It will only work when vendor item detail price is ON ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AcceptQtyWithoutConsiderRejectQty, clsFixedParameterCode.AcceptQtyWithoutConsiderRejectQty, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowGRN, clsFixedParameterCode.ShowGRN, "", "If Setting Is On Then GRN (Gate Receipt Note) Screen will be shown in Purchase Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowMRN, clsFixedParameterCode.ShowMRN, "", "If Setting Is On Then MRN (Material Receipt Note) Screen will be shown in Purchase Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipMRNGRNinCaseofMT, clsFixedParameterCode.SkipMRNGRNinCaseofMT, "1", "0-OFF;1-On setting will work only in case of Merchant trade")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGenerateMRN, clsFixedParameterCode.AutoGenerateMRN, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.WorkingHours, clsFixedParameterCode.WorkingHours, "8", "8: Total no of working hours in company")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatExcessLeaveAbsent, clsFixedParameterCode.TreatExcessLeaveAbsent, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VehicleInsuranceAlert, clsFixedParameterCode.VehicleInsuranceAlert, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GLACAccordingToTaxRate, clsFixedParameterCode.GLACAccordingToTaxRate, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PWD, clsFixedParameterCode.UserPWD, "dont type password", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowMilkReceiptAfterSettingsisOn, clsFixedParameterCode.AllowMilkReceiptAfterSettingsisOn, "kdil123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BMCMilkQCManualPassword, clsFixedParameterCode.BMCMilkQCManualPassword, "c106", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkReceiptTolerancePwd, clsFixedParameterCode.MilkReceiptTolerancePwd, "udl@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCC_DLTDATA_PWD, clsFixedParameterCode.MCCDLTPWD, "TecXpertb12sec2", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow_Excel_Code_on_Mcc_Master, clsFixedParameterCode.Allow_ExcelCode_On_Mcc, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AlowwdateChangeinPaymentEntry, clsFixedParameterCode.AlowwdateChangeinPaymentEntry, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.is_Allow_cancel_Transaction, clsFixedParameterCode.Is_Allow_Cancel_Transaction, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.is_Allow_cancel_Posted_Transaction, clsFixedParameterCode.is_Allow_cancel_Posted_Transaction, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShiftTiming, clsFixedParameterCode.ShiftTiming, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MulticurrencyDecimalPlaces, clsFixedParameterCode.GetMulitcurrencyDecimalPlaces, "4", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMS_User_Name, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMS_User_PWD, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMS_Sendor_ID, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMS_Provider, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCDefaultMilkItemCow, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCDefaultMilkItemBuffalo, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_Send_Sms, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_Send_Sms_ForVSP, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_Pick_No_from_Mail_Setting, clsFixedParameterCode.MilkSetting, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Send_Sms_Time, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCSampleRange, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCMilkSRNRepost, clsFixedParameterCode.MilkSetting, "tecxpertsrn", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCReceiptRange, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCMinKmRange, clsFixedParameterCode.MilkSetting, "1", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCFSSAI_DAYS, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCDisplay_All_Parameter, clsFixedParameterCode.MilkSetting, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, "0", "0- Off,1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableProjectFinder, clsFixedParameterCode.EnableProjectFinder, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Milk_Can_Weight_Ratio, clsFixedParameterCode.MilkSetting, "40", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Milk_Can_Weight_Tolerance_Neg, clsFixedParameterCode.MilkSetting, "10", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Milk_Can_Weight_Tolerance_Positive, clsFixedParameterCode.MilkSetting, "10", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCInvoiceScheduleDate, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCInvoiceScheduleTime, clsFixedParameterCode.MilkSetting, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCInvoiceScheduleInterval, clsFixedParameterCode.MilkSetting, "", "")
        'richa
        InsertDefaultValueFixedParameter(clsFixedParameterType.InvoiceManualNoWithPrefix, clsFixedParameterCode.InvoiceManualNoWithPrefix, "1", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoBackUp, clsFixedParameterCode.AutoBackUp, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCPurchase, clsFixedParameterCode.MCCPurchase, "0", "0-OFF;1-On")
        '
        'richa Ticket No BM00000003045 09/07/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.NotificationSettingforReOrderInPO, clsFixedParameterCode.NotificationSettingforReOrderInPO, "0", "0-None;1-Stop;2-Warning")
        ''
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkSaleDefaultMilkItem, clsFixedParameterCode.BulkSaleDefaultMilkItem, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BSDefaultMilkItem, clsFixedParameterCode.BSDefaultMilkItem, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, "", "")
        'richa Ticket No BM00000003042 09/07/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.NotificationSettingforReOrderInPurchaseRequisition, clsFixedParameterCode.NotificationSettingforReOrderInPurchaseRequisition, "0", "0-None;1-Warning")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseOrderAutomaticallyItemQtyBelowReorderLevel, clsFixedParameterCode.PurchaseOrderAutomaticallyItemQtyBelowReorderLevel, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NLevelAtVendor, clsFixedParameterCode.NLevelAtVendor, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NLevelAtCustomer, clsFixedParameterCode.NLevelAtCustomer, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NLevelAtLocation, clsFixedParameterCode.NLevelAtLocation, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.AutoItemNLevel, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterAsset, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterFinishGood, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterOther, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterRawMaterial, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterSemiFinishGood, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoItemNLevel, clsFixedParameterCode.CounterTradingGood, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowPurchaseControlAc, clsFixedParameterCode.ShowPurchaseControlAc, "0", "1:On, 0:Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Princi_Bom, clsFixedParameterCode.Princi_Bom, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AP_INV_COMMSN, clsFixedParameterCode.AP_INV_COMMSN, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Principal_Vendor, clsFixedParameterCode.Principal_Vendor, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Principal_Vendor_Database, clsFixedParameterCode.Principal_Vendor_Database, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Principal_Customer, clsFixedParameterCode.Principal_Customer, "", "")
        ''
        'InsertDefaultValueFixedParameter(clsFixedParameterType.ExeExpiredDate, clsFixedParameterCode.ExeExpiredDate, "", "")
        '' Anubhooti 10-July-2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateLTAOnHoliday, clsFixedParameterCode.CalculateLTAOnHoliday, "1", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateLTAOnWeekend, clsFixedParameterCode.CalculateLTAOnWeekend, "1", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateMediclaimOnHoliday, clsFixedParameterCode.CalculateMediclaimOnHoliday, "1", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateMediclaimOnWeekend, clsFixedParameterCode.CalculateMediclaimOnWeekend, "1", "0-OFF;1-On")
        '' Anubhooti 21-Aug-2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_Purchaseable_Item, clsFixedParameterCode.Is_Purchaseable_Item, "0", "0-OFF;1-On")
        '' Anubhooti 22-Aug-2014 (Setting For Demo)
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_AbemdmentForDemo, clsFixedParameterCode.Is_AbemdmentForDemo, "0", "0-OFF;1-On")
        '' Anubhooti 26-Aug-2014 (Setting For Item Is_FinishedGoods)
        InsertDefaultValueFixedParameter(clsFixedParameterType.Is_FinishedGoods, clsFixedParameterCode.Is_FinishedGoods, "0", "0-OFF;1-On")
        '' Anubhooti 28-Aug-2014 (Setting For Demo Print:Purchase Module)
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowStatusForPurchase, clsFixedParameterCode.ShowStatusForPurchase, "0", "0-OFF;1-On")
        '' Anubhooti 28-Aug-2014 (Setting For Demo Print:Sales Module)
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowStatusForSales, clsFixedParameterCode.ShowStatusForSales, "0", "0-OFF;1-On")
        '' Anubhooti 28-Aug-2014 (Setting For Demo:Sales Module)
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSerialNoForSales, clsFixedParameterCode.ShowSerialNoForSales, "0", "0-OFF;1-On")
        '' Anubhooti 02-Sep-2014 (Setting For Vendor Master)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedVendorCode, clsFixedParameterCode.AutoGeneratedVendorCode, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedVendorCodeForAllCompany, clsFixedParameterCode.AutoGeneratedVendorCodeForAllCompany, "0", "0-OFF;1-On")
        '' Anubhooti 02-Sep-2014 (Setting For Customer Master)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedCustomerCode, clsFixedParameterCode.AutoGeneratedCustomerCode, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedCustomerCodeForAllCompany, clsFixedParameterCode.AutoGeneratedCustomerCodeForAllCompany, "0", "0-OFF;1-On")
        '' Anubhooti 03-Sep-2014 BM00000003437 (Setting For Sub Account in Bank Master)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.InTransitFeatureIsRequired, clsFixedParameterCode.InTransitFeatureIsRequired, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PermissionSettingForTransactionWithBank, clsFixedParameterCode.PermissionSettingForTransactionWithBank, "0", "0-None; 1-GL Security; 2-Bank Permission")
        '' Anubhooti 12-Sep-2014 BM00000003890 (Setting For Fresh Sale)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToEnterMRPManually, clsFixedParameterCode.AllowToEnterMRPManually, "0", "0-OFF;1-On")
        '' Anubhooti 24-Sep-2014 BM00000003940 (Setting For Vehicle Master)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFieldsToBeManadatory, clsFixedParameterCode.AllowFieldsToBeManadatory, "0", "0-OFF;1-On")
        '' Anubhooti 08-Oct-2014 (Setting For Auto Generated Digits For Vendor)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedDigitsForVendor, clsFixedParameterCode.AutoGeneratedDigitsForVendor, "0", "")
        '' Anubhooti 08-Oct-2014 (Setting For Auto Generated Digits For Customer)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratedDigitsForCustomer, clsFixedParameterCode.AutoGeneratedDigitsForCustomer, "0", "")
        '' Anubhooti 08-Oct-2014 (Setting For Unit Cost Editable/Non-Editable On SRN)
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsRateEditableOnSRN, clsFixedParameterCode.IsRateEditableOnSRN, "0", "0-OFF;1-On")
        '' Anubhooti 23-Jan-2015 (Setting For Creation of GL Acc To Item GL Account(Issue/Return/Transfer))
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateGLAccToItem, clsFixedParameterCode.CreateGLAccToItem, "0", "0-OFF;1-On")
        '' Anubhooti 29-Jan-2015 (Setting For Cost Edit/Non-Edit On(Issue/Return/Transfer))
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsCostEditableOnIssueReturnTransfer, clsFixedParameterCode.IsCostEditableOnIssueReturnTransfer, "0", "0-OFF;1-On")
        'Richa Agarwal 05/08/2014 Against Ticket No BM00000003248
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsVolumeSchemeBydefault, clsFixedParameterCode.IsVolumeSchemeBydefault, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, "", "Have to Select discount Code From discount Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DiscountCodeForMPAdj, clsFixedParameterCode.DiscountCodeForMPAdj, "", "Have to Select discount Code From discount Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoSelDateandBankforPayEntryOnSalaryGeneration, clsFixedParameterCode.AllowtoSelDateandBankforPayEntryOnSalaryGeneration, "0", "0-OFF;1-On Allow to Select Date and Bank for PaymentEntry on Salary Generation screen ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAutoRecieptForManualCustomer, clsFixedParameterCode.CreateAutoRecieptForManualCustomer, "1", "Create Auto Reciept For Manual Customer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoRecieptBankCode, clsFixedParameterCode.AutoRecieptBankCode, "", "Have to Select Bank Code for auto rec")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoRecieptPaymentMode, clsFixedParameterCode.AutoRecieptPaymentMode, "", "Have to Select Payment Mode for auto rec")
        ''Pankaj Jha on 15/07/2015 For Skipping COGS GL Entry
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, "0", "0: Do Not Skip, 1: Skip")
        ''  End
        'Richa Agarwal 19/08/2014 Against Ticket No BM00000003110
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, "0", "0-OFF;1-On")
        '--------Richa Agarwal 21/08/2014 Against Ticket No BM00000003438
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoMRNGRNonDocumentAcceptance, clsFixedParameterCode.AllowAutoMRNGRNonDocumentAcceptance, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CommitedDefaultQty, clsFixedParameterCode.CommitedDefaultQty, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowPrintChallanInDairyDispatch, clsFixedParameterCode.ShowPrintChallanInDairyDispatch, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowBinMapping, clsFixedParameterCode.ShowBinMapping, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToShowSaleTypeinPaymentTermsReceivable, clsFixedParameterCode.AllowToShowSaleTypeinPaymentTermsReceivable, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToShowMilkTypeinAdjustmentEntry, clsFixedParameterCode.AllowToShowMilkTypeinAdjustmentEntry, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GatePassAfterTransfer, clsFixedParameterCode.GatePassAfterTransfer, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickRateFromPRICEChrtMasterFORUMang, clsFixedParameterCode.PickRateFromPRICEChrtMasterFORUMang, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StockTranferFromTransferPriceAndInvJVWithAvgCost, clsFixedParameterCode.StockTranferFromTransferPriceAndInvJVWithAvgCost, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToEditCategoryCodeinItemMaster, clsFixedParameterCode.AllowToEditCategoryCodeinItemMaster, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreditLimitApproval, clsFixedParameterCode.CreditLimitApproval, "yes", "This Password is used while approving Credit Days, Or Credit Limit.")
        '--------Richa Agarwal 28/08/2014  Against Ticket No .BM00000003667
        InsertDefaultValueFixedParameter(clsFixedParameterType.InvoiceBasedPO, clsFixedParameterCode.InvoiceBasedPO, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AdvanceAgainstSO, clsFixedParameterCode.AdvanceAgainstSO, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Purchase_SMSATPOST, clsFixedParameterCode.Purchase_SMSATPOST, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.showRFQ, clsFixedParameterCode.showRFQ, "0", "0-OFF;1-On")
        ''richA 02/09/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSaleInvoiceNoInPOfinderInSRN, clsFixedParameterCode.ShowSaleInvoiceNoInPOfinderInSRN, "0", "")
        ''-------------Preeti Gupta
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateValue, clsFixedParameterCode.CrateValue, "1", "")
        ' -----------end----------------
        '--------Pankaj jha 08/09/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.ControlSampleMandatory, clsFixedParameterCode.MilkSetting, "0", "When Set to 1 then control sample filed become mandatory at MCC dispatch Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.defaultCorrectionFactor, clsFixedParameterCode.MilkSetting, "0.14", "set the default correction factor value for different screens")
        ''richa 09/09/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCdefaultCorrectionFactorBS, clsFixedParameterCode.MCCdefaultCorrectionFactorBS, "0.14", "default correction factor value for MCC procument on Bulk Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchasedefaultCorrectionFactorBS, clsFixedParameterCode.PurchasedefaultCorrectionFactorBS, "0.14", "default correction factor value for Gate entry Purchase Type bulk Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.JOBdefaultCorrectionFactorBS, clsFixedParameterCode.JOBdefaultCorrectionFactorBS, "0.14", "default correction factor value for Gate entry Job Type on bulk Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultCorrectionFactorForBulkSale, clsFixedParameterCode.DefaultCorrectionFactorForBulkSale, "0.14", "default correction factor value for different screens on bulk Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, "0.14", "default correction factor value for different screens on bulk Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TabOrder, clsFixedParameterCode.AutoTabOrdering, "1", " When 1, Then Auto Tab Ordering Is ON, When 0, Auto Tab Ordering is OFF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TabOrder, clsFixedParameterCode.AutoTabOrderingPattern, "1", " When 1, Then Tab Ordering Pattern is Horizontal, When 2,  Tab Ordering Pattern is Vertical")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.IsItemEditableOnMCCDispatch, "0", " When 1, Then Item Finder Editable, When 0,  Item Finder Non-Editable")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LoadLoginScreen, clsFixedParameterCode.LoadLoginScreenDirectlyAfterStartup, "0", " 0: OFF, 1: ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.IsUOMSelectableOnMCCDispatch, "0", " When 1, Then UOM Finder Editable, When 0,  UOM Finder Non-Editable")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsItemWithDifferntUnitConsiderAsOtherItem, clsFixedParameterCode.IsItemWithDifferntUnitConsiderAsOtherItem, "0", " 0-Off,1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ViewTDSPwd, clsFixedParameterCode.ViewTDSPwd, "vtds", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowVendorLedgerasPerRightsForLocation, clsFixedParameterCode.ShowVendorLedgerasPerRightsForLocation, "0", "ShowVendorLedgerasPerRightsForLocation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCustomerLedgerasPerRightsForLocation, clsFixedParameterCode.ShowCustomerLedgerasPerRightsForLocation, "0", "ShowCustomerLedgerasPerRightsForLocation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableDynamicQRCodeForB2CInvoice, clsFixedParameterCode.EnableDynamicQRCodeForB2CInvoice, "0", "EnableDynamicQRCodeForB2CInvoice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AutoSetTabStopFalseForReadonlyControls, "0", " 1: On, 0:Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AutoRestoreGridLayout, "1", " 1: On, 0:Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsMRPWiseBalance, clsFixedParameterCode.IsMRPWiseBalance, "1", " 0-Off,1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateDbitNoteForShortPI, clsFixedParameterCode.CreateDbitNoteForShortPI, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateDbitNoteForRejectPI, clsFixedParameterCode.CreateDbitNoteForRejectPI, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateDebitNoteForUnitCost, clsFixedParameterCode.CreateDebitNoteForUnitCost, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, "0", "Fill integer for showing no. of digits after decimal in production.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsumptionQtyTollerance, clsFixedParameterCode.ConsumptionQtyTollerance, "0.6", "Tollerance % of ConsumptionQty.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GunnyBagTollerance, clsFixedParameterCode.GunnyBagTollerance, "10", "Tollerance % of Gunny Bag.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckBalanceOnSave, clsFixedParameterCode.RCDFCFProductionEntry, "1", "0-OFF;1-On;Check Balance on save")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckBalanceOnSave, clsFixedParameterCode.RCDFCFDispatch, "1", "0-OFF;1-On;Check Balance on save")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransferJEForLocationMapping, clsFixedParameterCode.TransferJEForLocationMapping, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, "0", "Fill integer for showing no. of digits after decimal in production for Fat and SNF Percentage.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ManualySelectBOMForChildBatch, clsFixedParameterCode.ManualySelectBOMForChildBatch, "0", "0:Off,1:On")
        '' setting for emails/ sms/alert reminder for bday/anniversary of employee(Panch Raj) on 05/12/2014
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToDispalyAlertForBDayAnniversary, clsFixedParameterCode.AllowToDispalyAlertForBDayAnniversary, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSendEmailForBDayAnniversary, clsFixedParameterCode.AllowToSendEmailForBDayAnniversary, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemDescForTankerdispatchPrint, clsFixedParameterCode.ItemDescForTankerDispatchPrint, "Mixed Chilled Milk", " Desc of Milk")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateEntryInPrevDate, "0", " 0: Not Allow, 1: Allow")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, "", "Xpert ERP Start Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemBatchWiseStartDate, clsFixedParameterCode.ItemBatchWiseStartDate, "", "Xpert ERP Item Batch Wise Start Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForTransfer, clsFixedParameterCode.CreateJEForTransfer, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSkipStageQLLogSheetInProd, clsFixedParameterCode.AllowToSkipStageQLLogSheetInProd, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsRemarkReasonMandatoryOnPO, clsFixedParameterCode.IsRemarkReasonMandatoryOnPO, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCostCenterAndHierarchyLevelInPurchaseModule, clsFixedParameterCode.ShowCostCenterAndHierarchyLevelInPurchaseModule, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsQCColumnRequiredonMRN, clsFixedParameterCode.IsQCColumnRequiredonMRN, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowUpto6DecimalPlaces, clsFixedParameterCode.ShowUpto6DecimalPlaces, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableOnPrivateChkbox, clsFixedParameterCode.EnableOnPrivateChkbox, "0", "0 :Off, 1: On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseModule, clsFixedParameterCode.RateDecimalPlaces, "3", "Allow user to input rate decimal places [0-10]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPOScheduling, clsFixedParameterCode.AllowPOScheduling, "0", "0:Off, 1:On for allow PO Scheduling in system.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowQcDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowWeighmentDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowQcDateBeforeGateEntryDateTime, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowUnloadingDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowcleaningDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateOutDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowSRNDateAfterCurrentDate, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsRGPAfterPurchaseOrder, clsFixedParameterCode.IsRGPAfterPurchaseOrder, "0", "0:Off,1:On if setting is On then PO finder is seen on RGP screen, otherwise RGP finder on PO screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowQualityModuleInERP, clsFixedParameterCode.AllowQualityModuleInERP, "0", "0:Off,1:On if setting is On then QC module is seen and SRN made way a QC Entry of item that mapped for QC.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SRNReportQuantityWise, clsFixedParameterCode.SRNReportQuantityWise, "0", "0:Off,1:On if setting is On then SRN Report Qty wise is seen on SRN screen, otherwise Value wise.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsCustomerGroupFieldsMandatory, clsFixedParameterCode.IsCustomerGroupFieldsMandatory, "0", "0:Off,1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsVendorGroupFieldsMandatory, clsFixedParameterCode.IsVendorGroupFieldsMandatory, "0", "0:Off,1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsPickServerDateForMultipleDispatchInvoice, clsFixedParameterCode.IsPickServerDateForMultipleDispatchInvoice, "0", "0-DO Date,1-Server Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayFranchiseeinCustomer, clsFixedParameterCode.DisplayFranchiseeinCustomer, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Idle, clsFixedParameterCode.isIdleTimerOn, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Idle, clsFixedParameterCode.idleTime, "5", "Time to be specified in minute")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddressOnPaymentVoucher, clsFixedParameterCode.AddressOnPaymentVoucherOnBankBasis, "1", "1: ON, 0:OFF Address on Payment Voucher Print On Bank Basis")
        'richa agarwal 17/03/2015 against ticket no BM00000005874
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBankDetailsManualinVM, clsFixedParameterCode.AllowBankDetailsManualinVM, "0", "0:Off, 1:On ")
        ''--------------------------------
        '============================Preeti Gupta Against Ticket no[BM00000009038]==========================================
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowHierarchyAndCostCenterInAPInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInAPInvoiceEntry, "0", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.WeighmentNotMandatoryInMCC, clsFixedParameterCode.WeighmentNotMandatoryInMCC, "0", "0:Off, 1:On ")
        '=======================================================End==============================================
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowHierarchyAndCostCenterInARInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInARInvoiceEntry, "0", "0:Off, 1:On ")
        ''RICHA AGARWAL 17/03/2015 Product Sale
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesTaxTypeatPS, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesTaxTypeatPS, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesRetailTypeatPS, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesRetailTypeatPS, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesExciseTypeatPS, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesExciseTypeatPS, "1", "0:Off, 1:On ")
        ''-------------------------
        ''RICHA AGARWAL 17/03/2015 MCC Sale
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesTaxatMCCSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesTaxatMCCSale, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesRetailatMCCSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesRetailatMCCSale, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesExciseatMCCSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesExciseatMCCSale, "1", "0:Off, 1:On ")
        ''-------------------------
        ''RICHA AGARWAL 17/03/2015 Misc Sale
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesTaxatMiscSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesTaxatMiscSale, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesRetailatMiscSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesRetailatMiscSale, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesExciseatMiscSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesExciseatMiscSale, "1", "0:Off, 1:On ")
        ''-------------------------
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.DisAllowIntermittentTankerForPlantDispatch, "1", "0:Off, 1:On ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, "0", "100000")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTankerDispatchGL, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTransferInGL, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProc, clsFixedParameterCode.PostTankerDispatchWithZeroAvgCost, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SendSalarySlipMailToEmployee, clsFixedParameterCode.SendSalarySlipMailToEmployee, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuickExport, clsFixedParameterCode.MaxRowsForQuickExport, "25000", "Number of rows to be Saved at once when using Quick Export")
        Dim dt As Date = clsCommon.GETSERVERDATE()
        dt = dt.AddDays(30)
        InsertDefaultValueFixedParameter(clsFixedParameterType.LicenceExpiryDate, clsFixedParameterCode.LicenceExpiryDate, clsCommon.EncryptString(clsCommon.GetPrintDate(dt, "dd/MMM/yyyy"), objCommonVar.CurrentCompanyCode + "A"), clsCommon.EncryptString("0", objCommonVar.CurrentCompanyCode + "B"))
        'InsertDefaultValueFixedParameter(clsFixedParameterType.LicenceExpiryDate, clsFixedParameterCode.LicenceExpiryDate, clsCommon.EncryptString(objCommonVar.CurrentCompanyCode, objCommonVar.CurrentCompanyCode + "A"), clsCommon.EncryptString(clsCommon.GetPrintDate(dt, "dd/MMM/yyyy"), objCommonVar.CurrentCompanyCode + "B"))
        InsertDefaultValueFixedParameter(clsFixedParameterType.LicenceNoOfExeConnection, clsFixedParameterCode.LicenceNoOfExeConnection, clsCommon.EncryptString("2", objCommonVar.CurrentCompanyCode + "C"), "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LicenceNoOfJournalEntry, clsFixedParameterCode.LicenceNoOfJournalEntry, clsCommon.EncryptString("100", objCommonVar.CurrentCompanyCode + "D"), "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LicenceNoOfUser, clsFixedParameterCode.LicenceNoOfUser, clsCommon.EncryptString("2", objCommonVar.CurrentCompanyCode + "E"), "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotMergeAPARAccount, clsFixedParameterCode.DoNotMergeAPARAccount, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowVisiDetail, clsFixedParameterCode.ShowVisiDetail, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerNameUniqueOnCM, clsFixedParameterCode.CustomerNameUniqueOnCM, "0", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsShortageIncludeInLandedCost, clsFixedParameterCode.IsShortageIncludeInLandedCost, "1", "0:OFF, 1:ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAutoMilkRGPinBulkSRN, clsFixedParameterCode.CreateAutoMilkRGPinBulkSRN, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayAllParameterinQualityCheck, clsFixedParameterCode.DisplayAllParameterinQualityCheck, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayTypeInMilkReceipt, clsFixedParameterCode.DisplayTypeInMilkReceipt, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddValidationofMilkTypeinsample, clsFixedParameterCode.AddValidationofMilkTypeinsample, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdFatCow, clsFixedParameterCode.SubStdFatCow, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdSNFCow, clsFixedParameterCode.SubStdSNFCow, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdFatBuff, clsFixedParameterCode.SubStdFatBuff, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdSNFBuff, clsFixedParameterCode.SubStdSNFBuff, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdFatMix, clsFixedParameterCode.SubStdFatMix, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubStdSNFMix, clsFixedParameterCode.SubStdSNFMix, "0", "Premium will Not applicable Below This % age")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMinCow, clsFixedParameterCode.FatMinCow, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMaxCow, clsFixedParameterCode.FatMaxCow, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMinCow, clsFixedParameterCode.SNFMinCow, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMaxCow, clsFixedParameterCode.SNFMaxCow, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMinBuff, clsFixedParameterCode.FatMinBuff, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMaxBuff, clsFixedParameterCode.FatMaxBuff, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMinBuff, clsFixedParameterCode.SNFMinBuff, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMaxBuff, clsFixedParameterCode.SNFMaxBuff, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMinMix, clsFixedParameterCode.FatMinMix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatMaxMix, clsFixedParameterCode.FatMaxMix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMinMix, clsFixedParameterCode.SNFMinMix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFMaxMix, clsFixedParameterCode.SNFMaxMix, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddIncentiveDeductioninMilkSample, clsFixedParameterCode.AddIncentiveDeductioninMilkSample, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualEnterinWeighment, clsFixedParameterCode.AllowManualEnterinWeighment, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SettlementBankOnlyPWD, clsFixedParameterCode.SettlementBankOnlyPWD, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UploaderPassword, clsFixedParameterCode.UploaderPassword, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankUploaderPassword, clsFixedParameterCode.BankUploaderPassword, "Bankupdate@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DocumentSequence, clsFixedParameterCode.DocumentSequence, "tecxpert@321", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, "1", "0-Off;1-On It will Effect All Purchase Transaction.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SHowBulkMilkWeighment, clsFixedParameterCode.SHowBulkMilkWeighment, "1", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BOOKINGFINDER_ON_CSASALEPATTI, clsFixedParameterCode.BOOKINGFINDER_ON_CSASALEPATTI, "1", "0-Off;1-On for filling manually booking rate on sale patti set 0, without selection of booking no.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StoreADJExportImportAfterPost, clsFixedParameterCode.StoreADJExportImportAfterPost, "tecxpert@importafterpost", "Password to import posted store adjustments.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatSNFControlOnProductionConsumption, clsFixedParameterCode.FatSNFControlOnProductionConsumption, "0", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuantityControlToleranceOnProductionConsumption, clsFixedParameterCode.QuantityControlToleranceOnProductionConsumption, "50", "Tolerance percentage for quantity in production and Issue/consumption")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LeaveBalanceAlertTypeOnAttendance, clsFixedParameterCode.LeaveBalanceAlertTypeOnAttendance, "0", "Set alert type on attendance transactions.0: None 1. On Saving Attendance 2. On Entering leave data 3. Both")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, "0", "0: None 1. Cash/Petty Cash 2. Bank 3. Stop All Bank Types: Stop Negative Bank Balance for Bank Type(Cash,Petty Cash,Bank, All) in Payment Entry,Bank Transfer and Refund transactions.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilk, "0", "0: Issue Basis 1. BOM Basis .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilkStandardization, "0", "0: Issue Basis 1. BOM Basis .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilkProduct, "0", "0: Issue Basis 1. BOM Basis .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeOther, "1", "0: Issue Basis 1. BOM Basis .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ValidateFatSnfOnProduction, clsFixedParameterCode.ValidateFatSnfOnProduction, "1", "0: Off 1. On System will not allow to produce more fat/snf than consumption .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOverheadCostOnProductionEntry, clsFixedParameterCode.ShowOverheadCostOnProductionEntry, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ActivateProductionWithoutBatch, clsFixedParameterCode.ActivateProductionWithoutBatch, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEOnProduction, clsFixedParameterCode.CreateJEOnProduction, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSaveMultipleEmployeeStatus, clsFixedParameterCode.AllowToSaveMultipleEmployeeStatus, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntrySecondaryTransporter, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryMCCLeaseVendor, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForFreshSale, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForProductSale, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForBulkSale, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryOthers, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryPrimaryTransporter, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForCSATransfer, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForTransfer, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoubleClickOnVC, clsFixedParameterCode.DoubleClickOnVC, "1", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForGateentry, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryTransporterForDairySale, "0", "0: Off 1. On .")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickManual_CSATransfer_OnTRansferReturn, clsFixedParameterCode.PickManual_CSATransfer_OnTRansferReturn, "1", "0:Off, 1:On; When setting is OFF then transfer is knock-off auto on FIFO based and Rate is manual or auto on setting based, But when Manual setting is On then transfer selected by user manually and Rate not changed in any condition.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickManual_CSATransfer_OnCSASalePatti, clsFixedParameterCode.PickManual_CSATransfer_OnCSASalePatti, "0", "0:Off, 1:On; When setting is OFF then transfer is knock-off auto on FIFO based, But when Manual setting is On then transfer selected by user manually on CSA Sale Patti.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDistributorSaleAtCSA_SaleInvoice, clsFixedParameterCode.AllowDistributorSaleAtCSA_SaleInvoice, "0", "0:Off, 1:On; When setting is OFF then CSA Sale Patti done for consignee,otherwise done for Distributor.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowItemWiseCSAAccountingON_CSASale, clsFixedParameterCode.AllowItemWiseCSAAccountingON_CSASale, "0", "0:Off, 1:On; When setting is OFF then CSA accounting done in normal way,otherwise item-wise CSA accounting done.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, "0", "0:Off, 1:On; Auto Integrate from weighing scale.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsAutoTankerWeighmentForBulkSale, clsFixedParameterCode.IsAutoTankerWeighmentForBulkSale, "0", "0:Off, 1:On; Auto Integrate weighing scale for bulk sale.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsAdditionalInformationOnVillageMaster, clsFixedParameterCode.IsAdditionalInformationOnVillageMaster, "0", "0:Off, 1:On; Additional Information On Village Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckLiveStockInProductionDuringTrans, clsFixedParameterCode.CheckLiveStockInProductionDuringTrans, "1", "0:Off, 1:On; all validation regarding the Milk Item stock will appled of Server date if setting :1 otherwise on Transaction Date.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VLCTimeTableColumnShow, clsFixedParameterCode.VLCTimeTableColumnShow, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyEffectiveStartDate, clsFixedParameterCode.ApplyEffectiveStartDate, "0", "0:Off, 1:On; Apply Effective Start Date in Milk Route Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VLCTimeTableColumnMandatory, clsFixedParameterCode.VLCTimeTableColumnMandatory, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isOneMCCOnePrimaryTranporter, clsFixedParameterCode.isOneMCCOnePrimaryTranporter, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isIntimationRequired, clsFixedParameterCode.isIntimationRequired, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QualityThenWeighmentinBulkProcurement, clsFixedParameterCode.QualityThenWeighmentinBulkProcurement, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, "0", "0:Off, 1:On; Bulk Procurement Chamber wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isPriceChartGradeWise, clsFixedParameterCode.isPriceChartGradeWise, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isCreateBulkProcPriceChartItemWise, clsFixedParameterCode.isCreateBulkProcPriceChartItemWise, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CowFATPer, clsFixedParameterCode.CowFATPer, "5", "Default Cow FAT Percentage")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MixFATPer, clsFixedParameterCode.MixFATPer, "0", "Mix FAT Percentage")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CmtSetting, clsFixedParameterCode.CmtSetting, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Hidedetaildate, clsFixedParameterCode.Hidedetaildate, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteTempData, clsFixedParameterCode.DeleteTempData, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteTempDataFromTime, clsFixedParameterCode.DeleteTempDataFromTime, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteTempDataFromTime, clsFixedParameterCode.DeleteTempDataFromTime, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipDatePayPeriodMaster, clsFixedParameterCode.SkipDatePayPeriodMaster, "0", "0:Off, 1:On;")

        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowNewFormatofPDF, clsFixedParameterCode.ShowNewFormatofPDF, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PaymentProcessPrintInHindi, clsFixedParameterCode.PaymentProcessPrintInHindi, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VendorBankAdviceForSWM, clsFixedParameterCode.VendorBankAdviceForSWM, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkSamplShowOddEvenTwoGrid, clsFixedParameterCode.MilkSamplShowOddEvenTwoGrid, "0", "0:Off, 1:On; Show Tow Grid Odd Even")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OpenODDEvenForm, clsFixedParameterCode.OpenODDEvenForm, "1", "0:Off, 1:On; Open Odd Even Form")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Open4AnalyzerForm, clsFixedParameterCode.Open4AnalyzerForm, "0", "0:Off, 1:On; Open 4 Analyzer Form")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsApplyEMIOnAssetValue, clsFixedParameterCode.IsApplyEMIOnAssetValue, "0", "0:Off, 1:On; Add Asset Value")
        '= KUNAL '========================================================================================
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, "0", "0:Off, 1:On; Allow Future Date Transaction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowCSAPriceMasterPostedData, clsFixedParameterCode.AllowCSAPriceMasterPostedData, "0", "1:On, 0:Off; when 1 then post button is visible on CSA Price Master and effect seen accordingly.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowItemMasterPostedData, clsFixedParameterCode.AllowItemMasterPostedData, "1", "1:On, 0:Off; when 1 then Item Master item post button is visible")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowMilkItemMasterPostedData, clsFixedParameterCode.AllowMilkItemMasterPostedData, "1", "1:On, 0:Off; when 1 then Milk Item Master item post button is visible")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkProcItemPostedData, clsFixedParameterCode.AllowBulkProcItemPostedData, "1", "1:On, 0:Off; when 1 then Bulk Proc Milk Item post button is visible")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPriceListMasterPostedData, clsFixedParameterCode.AllowPriceListMasterPostedData, "0", "0:Off, 1:On; Show NRGP Request Document Numbers")
        'KUNAL > UDIL > DATE : 16-NOV-2016
        InsertDefaultValueFixedParameter(clsFixedParameterType.FindNRGP_Request, clsFixedParameterCode.FindNRGP_Request, "0", "1:On, 0:Off;  Show on Checked and Hide on Unchecked NRGP Request Numbers")
        'stuti
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemCrateWtinKg, clsFixedParameterCode.ItemCrateWtinKg, "0", "Item Default Crate Wt.(Kg.)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemJaaliWtinKg, clsFixedParameterCode.ItemJaaliWtinKg, "0", "Item Default Jaali Wt.(Kg.)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemBoxWtinKg, clsFixedParameterCode.ItemBoxWtinKg, "0", "Item Default Box Wt.(Kg.)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemCrateRate, clsFixedParameterCode.ItemCrateRate, "0", "Item Default Crate Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemJaaliRate, clsFixedParameterCode.ItemJaaliRate, "0", "Item Default Jaali Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemBoxRate, clsFixedParameterCode.ItemBoxRate, "0", "Item Default Box Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemCanRate, clsFixedParameterCode.ItemCanRate, "0", "Item Default Can Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CustomerMasterFinderOnLocationwiseARReceipt, clsFixedParameterCode.CustomerMasterFinderOnLocationwiseARReceipt, "0", "0:Off, 1:On; When setting is on then customer master finder work on location basis otherwise independent")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SameuserCanNotloginmultipletimes, clsFixedParameterCode.SameuserCanNotloginmultipletimes, "0", "0:Off, 1:On; When setting is on then same user can-not login multiple times")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCancelButtonPO, clsFixedParameterCode.ShowCancelButtonPO, "0", "0:Off, 1:On; Show cancel button on purchase order")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplySendApprovalSetting, clsFixedParameterCode.ApplySendApprovalSetting, "0", "0:Off, 1:On; Send GRN for Approval if GRN QTY>PO QTY")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, "0", "0:Off, 1:On; Show option for selecting capex code and capex sub codes on purchase order")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoClosePO, clsFixedParameterCode.AutoClosePO, "0", "0:Off, 1:On; Auto close PO when all qty received.")
        'InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJVForAllCasesinRGP, clsFixedParameterCode.CreateJVForAllCasesinRGP, "0", "0:Off, 1:On; Create JV for all cases in RGP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StoreRequisitionMandatoryforstorerequest, clsFixedParameterCode.StoreRequisitionMandatoryforstorerequest, "0", "0:Off, 1:On; Store requisition mandatory for store request")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MandatoryEmployeeOnVehicleMaster, clsFixedParameterCode.MandatoryEmployeeOnVehicleMaster, "0", "0:Off, 1:On; Employee no mandatory for vehicle master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MTCapacityRequired, clsFixedParameterCode.MTCapacityRequired, "0", "0:Off, 1:On; MT Capacity Required")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PlantDepotMappingMandatory, clsFixedParameterCode.PlantDepotMappingMandatory, "0", "0:Off, 1:On; Mapping location of plant type with depot type is mandatory")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowThreeFormatByDefaultForPrint, clsFixedParameterCode.AllowThreeFormatByDefaultForPrint, "0", "0:Off, 1:On; Allow printing 3 formats by default rather than 4.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.POCancel, clsFixedParameterCode.POCancel, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BackDateEntryPwd, clsFixedParameterCode.BackDateEntryPwd, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBackDateEntry, clsFixedParameterCode.AllowBackDateEntry, "365", "Allow back date entry for given days")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RevisedBudget, clsFixedParameterCode.RevisedBudget, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DipMarkingMendatory, clsFixedParameterCode.DipMarkingMendatory, "0", "0:Off, 1:On; Make dip marking mendatory")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDispatchChecklistOnProductDispatch, clsFixedParameterCode.AllowDispatchChecklistOnProductDispatch, "0", "0:Off, 1:On; Allow dispatch checklist on product dispatch.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowIndentBasedOnCreatedUser, clsFixedParameterCode.ShowIndentBasedOnCreatedUser, "0", "0:Off, 1:On; Show indent based on created user.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSystemStockinOpenMCC, clsFixedParameterCode.ShowSystemStockinOpenMCC, "0", "0:Off, 1:On; Show system stock on open MCC shift.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Tankerfromtankersalemasteringateentry, clsFixedParameterCode.Tankerfromtankersalemasteringateentry, "0", "0:Off, 1:On; Select tanker from tanker sale master in gate entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, "0", "0:Off, 1:On; Apply multi-chamber in bulk Sale weighment entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultItemUOMForBulkSale, clsFixedParameterCode.DefaultItemUOMForBulkSale, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.InsuranceNoAndSealNoInBulkDispatch, clsFixedParameterCode.InsuranceNoAndSealNoInBulkDispatch, "0", "0:Off, 1:On; Option for entering Insurance No and Seal No.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ValidateFatSNFOnJobMilkSRN, clsFixedParameterCode.ValidateFatSNFOnJobMilkSRN, "0", "0:Off, 1:On; Validate FAT KG & SNF KG on Job Milk SRN")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CancelDocDueToSRNReturn, clsFixedParameterCode.CancelDocDueToSRNReturn, "0", "0:Off, 1:On; Cancel document due to SRN Return")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AmountInLacsOnMisSaleRegister, clsFixedParameterCode.AmountInLacsOnMisSaleRegister, "0", "0:Off, 1:On; Amount in lacs on MIS SALE REGISTER")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MakeClosingofPOreadonlyforuser, clsFixedParameterCode.MakeClosingofPOreadonlyforuser, "0", "0:Off, 1:On; Make closing of PO read only.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MakeClosingofCustomerTenderreadonlyforuser, clsFixedParameterCode.MakeClosingofCustomerTenderreadonlyforuser, "0", "0:Off, 1:On; Make closing of Customer Tender read only.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MakeClosingofCustomerTenderOrderreadonlyforuser, clsFixedParameterCode.MakeClosingofCustomerTenderOrderreadonlyforuser, "0", "0:Off, 1:On; Make closing of Customer Tender order read only.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShortCloseItemWiseOnPO, clsFixedParameterCode.ShortCloseItemWiseOnPO, "0", "0:Off, 1:On; Allow short close item-wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowModificationOnApprovalByApprovalUser, clsFixedParameterCode.AllowModificationOnApprovalByApprovalUser, "0", "0:Off, 1:On; Allow Modification On Approval By Approval User")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoCalculateADDREMOVEQty, clsFixedParameterCode.AllowAutoCalculateADDREMOVEQty, "0", "0:Off, 1:On; Auto Calculate Qty of Add/Remove Item")
        '---end here----------
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATDeductionPercent, clsFixedParameterCode.FATDeductionPercent, "65", "Deduction FAT % (1-100)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFDeductionPercent, clsFixedParameterCode.SNFDeductionPercent, "100", "Deduction SNF % (1-100)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RejectionReturnPenaltyPerUnit, clsFixedParameterCode.RejectionReturnPaneltyPerUnit, "1.50", "Return Penalty/Unit")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RejectionDrainPenaltyPerUnit, clsFixedParameterCode.RejectionDrainPenaltyPerUnit, "15", "Drain Penalty/Unit ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RejectionCOBPenaltyPerUnit, clsFixedParameterCode.RejectionCOBPenaltyPerUnit, "0", "COB Penalty/Unit ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GraceTimeForTransporter, clsFixedParameterCode.GraceTimeForTransporter, "40", "Tranporter Grace Time in Minutes")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GraceTimeFromGateEntryToDocWeighing, clsFixedParameterCode.GraceTimeFromGateEntryToDocWeighing, "40", "Time From Gate Entry To Dock Weighing")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCSAReturnTypeOnScreen, clsFixedParameterCode.ShowCSAReturnTypeOnScreen, "0", "0:Off,1:On; When setting ON then return type drop-down seen on screen and relative effects working,Otherwise no effect worked.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCSARequestScreen, clsFixedParameterCode.ShowCSARequestScreen, "0", "0:Off,1:On; When setting ON then CSA Request screen seen and used on CSA DO,Otherwise CSA Booking screen worked and effect seen on Sale Patti.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSchemeOnCSADeliveryOrder, clsFixedParameterCode.AllowSchemeOnCSADeliveryOrder, "0", "0:Off,1:On; When setting ON then scheme is applicable at CSA DO,Otherwise no scheme effect.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowOtherItemOnCSAPriceMaster, clsFixedParameterCode.AllowOtherItemOnCSAPriceMaster, "0", "0:Off,1:On; When setting ON then other than CPD items seen on CSA Price Master, otherwise only CPD items seen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FreightChargeOnCSASaleInvoice, clsFixedParameterCode.FreightChargeOnCSASaleInvoice, "0", "0:Off,1:On; When setting ON then Freight column seen and commission/freight cal. itemwise on CSA Sale Invoice/Patti, otherwise no freight calculation.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, "0", "0:Off,1:On; When setting ON then Round-off field seen on CSA Sale Invoice/Patti,CSA Transfer, otherwise document total amount is not rounded.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRoundOffAmountOnDCSSale, clsFixedParameterCode.AllowRoundOffAmountOnDCSSale, "0", "0:Off,1:On; When setting ON then Round-off field seen on DCS Sale and DCS Sale Entry, otherwise document total amount is not rounded.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDisabledCommissionOnCSATransfer, clsFixedParameterCode.AllowDisabledCommissionOnCSATransfer, "0", "0:Off,1:On; When setting ON then no commission seen on CSA Transfer, otherwise commission calculation done.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoReadonly_UnitRate_AtCSASale, clsFixedParameterCode.DoReadonly_UnitRate_AtCSASale, "0", "0:Off,1:On; When setting ON then unit rate column is not editable on CSA DO/Transfer and Sale Patti, otherwise unit rate is editable.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow_SaleMfgACONCSAPatti, clsFixedParameterCode.Allow_SaleMfgACONCSAPatti, "0", "0:Off,1:On; When setting ON then sale manufacturing account go for GL entry,otherwise not seen on GL for CSA Sale Patti.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateCrateLinerQty, clsFixedParameterCode.UpdateCrateLinerQty, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSchemeItemCondONSchemeMaster, clsFixedParameterCode.AllowSchemeItemCondONSchemeMaster, "0", "0:Off,1:On; When setting ON then item set as scheme item open on Scheme Master Dairy,otherwise all finder of item open.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ForUDLOnly, clsFixedParameterCode.ForUDLOnly, "0", "0:Off,1:On; When setting ON then CSA Sale effect for UDL,otherwise standard cycle run.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckCreditLimitonCSADO, clsFixedParameterCode.CheckCreditLimitonCSADO, "0", "0:Off,1:On; When setting ON then Credit limit outstanding check on CSA DO,otherwise no check.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, "0", "0:Off,1:On; When setting ON then Gross wt. calculated from item master set values and readonly on CSA Transfer,otherwise manual entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayPriceCodeForBKN, clsFixedParameterCode.DisplayPriceCodeForBKN, "0", "0:Off,1:On; When setting ON then Price Code Finder Will Show.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableExciseONCSASalePatti, clsFixedParameterCode.EnableExciseONCSASalePatti, "0", "0:Off,1:On; When setting ON then Excise entry made on CSA Sale Patti,otherwise no excise entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BatchSkipCSAReturn, clsFixedParameterCode.BatchSkipCSAReturn, "1", "0:Off,1:On; When setting ON then batch entry not mandatory on CSA Sale Patti/Return,otherwise mandatory entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, "0", "0:Off, 1:On; Apply tanker chamber wise for MCC")
        'Prabhakar'
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowLoginTypeCNFdistributerRetailer, clsFixedParameterCode.AllowLoginTypeCNFdistributerRetailer, "0", "0:Off,1:On; When settinng ON Login Type - CNF , Distributer , Retailer Panel visible.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSchemeItemQty, clsFixedParameterCode.AllowSchemeItemQty, "0", "0:Off,1:On; When settinng ON Allow Scheme Item in Materix Report.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDairyDeliveryOrderPrint, clsFixedParameterCode.AllowDairyDeliveryOrderPrint, "0", "0:Off,1:On; When settinng ON Allow Print Button for Delivery Order.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowSealNumberForTunkerOut, clsFixedParameterCode.ShowSealNumberForTunkerOut, "0", "0:Off,1:On; When settinng ON Show Seal Number for Tunker Out visible.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HideRateDispatchCentreCode, clsFixedParameterCode.HideRateDispatchCentreCode, "0", "0:Off,1:On; When settinng ON Hide Rate and Dispatch Centre Code.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPromptPendingDocs, clsFixedParameterCode.AllowPromptPendingDocs, "0", "0:Off,1:On; When settinng ON Allow Prompt Pending Docs")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, "0", "0:Off,1:On; When settinng ON Allow Auto Generate Document Number In Master Screen.")
        'kunal
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowDocsStatusFilters, clsFixedParameterCode.ShowDocsStatusFilters, "0", "0:Off,1:On; Show Or Hide Documents Declaration's Status Filters")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, "0", "0:Off,1:On; When settinng ON Department is used Mendatrory Field for Puchase Cycle.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowVehicleGateOutValidationScrapSale, clsFixedParameterCode.AllowVehicleGateOutValidationScrapSale, "0", "0:Off,1:On; When settinng ON Not Allow Vehicle if Vehicle Code not Gate Out For Scrap Sale.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowVehicleGateOutValidationCSATransfer, clsFixedParameterCode.AllowVehicleGateOutValidationCSATransfer, "0", "0:Off,1:On; When settinng ON Not Allow Vehicle if Vehicle Code not Gate Out For CSA Transfer.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowVehicleGateOutValidationSPSale, clsFixedParameterCode.AllowVehicleGateOutValidationSPSale, "0", "0:Off,1:On; When settinng ON Not Allow Vehicle if Vehicle Code not Gate Out For Shipment Product Sale.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowVehicleGateOutValidationTransfer, clsFixedParameterCode.AllowVehicleGateOutValidationTransfer, "0", "0:Off,1:On; When settinng ON Not Allow Vehicle if Vehicle Code not Gate Out For Transfer.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowWithoutUnitCostIssueReturnEntry, clsFixedParameterCode.AllowWithoutUnitCostIssueReturnEntry, "0", "0:Off,1:On; When settinng ON  Allow Without amount save issue/return entry.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ZeroCostForReprocess, clsFixedParameterCode.ZeroCostForReprocess, "0", "0:Off,1:On; When settinng ON  Zero Cost For Reprocess.")
        'AllowWithoutUnitCostIssueReturnEntry
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsAutoReceiptPayment, clsFixedParameterCode.IsAutoReceiptPayment, "0", "0:Off,1:On; When Auto Receipt Payment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TransferEntryOnInvCtrlAccount, clsFixedParameterCode.TransferEntryOnInvCtrlAccount, "0", "0:Off,1:On; When off GIT concept.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoUpdateVLCUploaderCodeInVLCMaster, clsFixedParameterCode.AutoUpdateVLCUploaderCodeInVLCMaster, "0", "0:Off,1:On.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StandardInterfaceForMilkShiftEnd, clsFixedParameterCode.StandardInterfaceForMilkShiftEnd, "0", "0:Off,1:On.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShiftEndAllowManualEntryOfDeduction, clsFixedParameterCode.ShiftEndAllowManualEntryOfDeduction, "1", "0:Off,1:On.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PTMRatePerLtrKGOnStdQty, clsFixedParameterCode.PTMRatePerLtrKGOnStdQty, "0", "0:Off,1:On. calualte Tranporter charges on Standard Qty")
        ' for mobile app
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultBank, clsFixedParameterCode.DefaultBank, "", "Default bank for mobile application.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultLocation, clsFixedParameterCode.DefaultLocation, "", "Default Location for mobile application.")
        '====================Added by preeti===============
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowParticluarColumnInSalesRegisterForGopalJee, clsFixedParameterCode.ShowParticluarColumnInSalesRegisterForGopalJee, "0", "0:Off,1:On")
        'Added by Nazia
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowPrintDiscountInDairyDispatchForGopaljee, clsFixedParameterCode.ShowPrintDiscountInDairyDispatchForGopaljee, "0", "0-OFF;1-On: Show print discount.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkReceiptRequiredApproval, clsFixedParameterCode.MilkReceiptRequiredApproval, "1", "0-OFF;1-On: Milk Receipt Required Approval collect from other VLC")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LinkDepartmentBetweenIndentAndIssue, clsFixedParameterCode.LinkDepartmentBetweenIndentAndIssue, "0", "0-OFF;1-On: check between Department Indednt Qty by issue Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CombineExportImportOnSchemeMaster, clsFixedParameterCode.CombineExportImportOnSchemeMaster, "0", "0-OFF;1-On: when ON then combined sheet export/import menu seen,otherwise not seen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OpenPOforRejectShortageQty, clsFixedParameterCode.OpenPOforRejectShortageQty, "0", "0-OFF;1-On: Open PO for Reject/Shortage Qty.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoSelectMCCRouteVLC, clsFixedParameterCode.AutoSelectMCCRouteVLC, "0", "0-OFF;1-On: Auto Select Route when MCC select and VLC When Route select")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickServerDateWithNoChange, clsFixedParameterCode.PickServerDateWithNoChange, "1", "0-OFF;1-On: Can not change the document Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickFinishedItemasBatchItem, clsFixedParameterCode.PickFinishedItemasBatchItem, "0", "0-OFF;1-On: when ON then on new creation of item as finish good bydefault BatchItem is checked on Item Master,otherwise no effect user select manually.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ToleranceFixFor_RM_OT_TRADE, clsFixedParameterCode.ToleranceFixFor_RM_OT_TRADE, "0", "0-OFF;1-On:when ON then tolerance% is mandatory for RM,Other and Trade items on Item Masters,otherwise non mandatory.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderAdvancePayment, clsFixedParameterCode.ConsiderAdvancePayment, "0", "0-OFF;1-On,When On Remaining Adavance/On Account payment Tab is visible and deduct in payable amount")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayableAmountZeroForMCCSale, clsFixedParameterCode.PayableAmountZeroForMCCSale, "0", "0-OFF;1-On,Payable Amount Zero For MCC Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow_AmountTruncate_BulkMilkSRN, clsFixedParameterCode.Allow_AmountTruncate_BulkMilkSRN, "0", "0-OFF;1-On,When On truncate Amount on Bulk Milk SRN[Eg:- 2.1=2 and 2.9=2],Otherwise no truncation.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoPurchaseReturnFromIssueReturn, clsFixedParameterCode.AutoPurchaseReturnFromIssueReturn, "0", "0-OFF;1-On,When On then auto purchase return from issue/return for Rejected type,Otherwise no purchase return.")
        'Sanjeet =========
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAlternateVechileforFreshSale, clsFixedParameterCode.ShowAlternateVechileforFreshSale, "0", "0-OFF;1-On: Show Alternate Vechile")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionOfTransporterInDairyDispatch, clsFixedParameterCode.CreateProvisionOfTransporterInDairyDispatch, "0", "0-OFF;1-On: Create provision of transporter")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SinglePrintCopyDairyInvoice, clsFixedParameterCode.SinglePrintCopyDairyInvoice, "0", "0-OFF;1-On:Print Single copy for dairy Invoice")
        'CreateProvisionOfTransporterInDairyDispatch
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncludeRatePerHoursIn, clsFixedParameterCode.IncludeRatePerHoursIn, "0", "0-OFF;1-On: Show Include Rate Per Hours In")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, "1", "0-OFF;1-ON: Allow GST Applicable")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GSTApplicableDate, clsFixedParameterCode.GSTApplicableDate, "01/JUN/2017", "Enter Valid GST Applicable Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowPanNoValidation, clsFixedParameterCode.AllowPanNoValidation, "1", "0-OFF;1-ON: Allow PAN No Validation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GSTActiveTaxesRatesGroup, clsFixedParameterCode.GSTActiveTaxesRatesGroup, "0", "0-OFF;1-ON: Show  only Active Taxes/Rates/Group")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualRejectionOfTanker, clsFixedParameterCode.AllowManualRejectionOfTanker, "0", "0-OFF;1-ON: Allow Manual Rejection of Tanker")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunBulkProcOnAdjustedFATCLR, clsFixedParameterCode.RunBulkProcOnAdjustedFATCLR, "0", "0-OFF;1-ON: Run Bulk Proc On Adjusted FAT and CLR")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkProcNetWeightCalculationWithVendorWeight, clsFixedParameterCode.BulkProcNetWeightCalculationWithVendorWeight, "0", "0-OFF;1-ON: Allow Bulk Proc Net Weight Calculation With Vendor Weight")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkProcPriceChartStandardRateWithZero, clsFixedParameterCode.BulkProcPriceChartStandardRateWithZero, "0", "0-OFF;1-ON: Allow Bulk Proc Price Chart standard rate with zero")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RemoveForceAapprovalofBulkSRN, clsFixedParameterCode.RemoveForceAapprovalofBulkSRN, "0", "0-OFF;1-ON: Remove Force Approval of Bulk SRN")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, "0", "0-OFF;1-ON: Allow Plant/Depot/MCC type Location On MCC Item Price Chart/Mcc sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplySaleRouteInDCSSale, clsFixedParameterCode.ApplySaleRouteInDCSSale, "0", "0-OFF;1-ON: Apply sale Route in DCS Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ValidateCustomerPANwithName, clsFixedParameterCode.ValidateCustomerPANwithName, "1", "0-OFF;1-ON: Validate Customer PAN with Name")
        InsertDefaultValueFixedParameter(clsFixedParameterType.POWeighmentManual, clsFixedParameterCode.POWeighmentManual, "0", "0-OFF;1-On:0-OFF, Manulay enter weighment.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ValidateTaxGroupForTransaction, clsFixedParameterCode.ValidateTaxGroupForTransaction, "0", "0-OFF;1-On:0-OFF, Validate Tax Group Should not Blank.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSeprateSchemeItemPrintDairySaleInvoice, clsFixedParameterCode.AllowSeprateSchemeItemPrintDairySaleInvoice, "0", "0-OFF;1-On:0-OFF, Allow Seprate Scheme Item Print DairySaleInvoice.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, "0", "0-OFF;1-On:0-OFF, Enable Hirerachy Level Cost Centre.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableStoreCostCentre, clsFixedParameterCode.EnableStoreCostCentre, "0", "0-OFF;1-On:0-OFF, Enable Store Cost Centre.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableCostingMethod, clsFixedParameterCode.EnableCostingMethod, "1", "1-Average,2-FIFO,3-LIFO; Enable Costin Method.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateItemCostonAvgForAssembly, clsFixedParameterCode.CalculateItemCostonAvgForAssembly, "0", "Calculate Item Cost on Avg For Assembly")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllCustomerOnMccMaterialSale, clsFixedParameterCode.ShowAllCustomerOnMccMaterialSale, "0", "0-OFF;1-On:0-OFF, Show All Customer On Mcc Material Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowDefaultUser, clsFixedParameterCode.ShowDefaultUser, "0", "0-OFF;1-On:0-OFF, Show Default User")
        ' (UDL) 21/12/2016
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToGenerate_NEFTUPLOADER, clsFixedParameterCode.AllowToGenerate_NEFTUPLOADER, "0", "0-OFF;1-On:0-OFF, Generate New NEFT UPLOADER")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DebitBankSelectWithNewFormateInNFTUploader, clsFixedParameterCode.DebitBankSelectWithNewFormateInNFTUploader, "0", "0-OFF;1-On:0-OFF, DebitBankSelectWithNewFormateInNFTUploader ")
        '======(UDL)17/11/2016======
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowVatSeriesNoSeprately, clsFixedParameterCode.ShowVatSeriesNoSeprately, "0", "0-OFF;1-On: Show Vat Series No. Seperatly")
        '======(UDL)05/01/2017======
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBulkPostingofAllDocuments, clsFixedParameterCode.AllowBulkPostingofAllDocuments, "0", "0-OFF;1-On: Allow Bulk Posting of All Documents")
        '======(UDL)10/01/2017======
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSameaAdditionalChargesMultiTime, clsFixedParameterCode.AllowSameaAdditionalChargesMultiTime, "1", "0-OFF;1-On: Allow Sam Additional Charges Multiple Time (Vendor Service Charges)")
        '(01/02/2017)
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSaveAndUpdatePasswordBased, clsFixedParameterCode.AllowToSaveAndUpdatePasswordBased, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowMasterModificationWithSecurity, clsFixedParameterCode.AllowMasterModificationWithSecurity, "0", "0-OFF;1-On: Allow Master Modification With Security")
        '(01/02/2017)
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyRTGSAmtMoreThanGiven, clsFixedParameterCode.ApplyRTGSAmtMoreThanGiven, "200", "RTGS Amount")
        '=============================
        InsertDefaultValueFixedParameter(clsFixedParameterType.GenerateSecondryCode, clsFixedParameterCode.GenerateSecondryCode, "0", "0-OFF;1-On: Excise Secondary Series on Transfer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateFIFOAndLIFOCosting, clsFixedParameterCode.CalculateFIFOAndLIFOCosting, "1", "0-OFF;1-On: Calculate FIFO And LIFO Costing")
        '===
        '============end here===========================
        '    Dim qry As String = " select distinct OtherAssemblyFilePathAndName  from TSPL_PROGRAM_MASTER  where isnull(IsLoadFromOtherAssembly ,0)=1"
        '    Dim dtt As DataTable = clsDBFuncationality.GetDataTable(qry)
        '    If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
        '        For i As Integer = 0 To dtt.Rows.Count - 1
        '            Dim AsmName As String = clsCommon.myCstr(dtt.Rows(i)("OtherAssemblyFilePathAndName"))
        '            clsCreateAllTables.InvokeMethodSlow(AsmName, "clsFixedParameterCustom", "FixedParameterValues", Nothing)
        '        Next
        '    End If
        'Catch ex As Exception
        'End Try
        ''======Ravi============
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddTypeForUserMaster, clsFixedParameterCode.AddTypeForUserMaster, "0", "0-OFF;1-On: Add Type(Super User, Driver) in UserMaster")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddParavetEmployeeType, clsFixedParameterCode.AddParavetEmployeeType, "0", "0-OFF;1-On: Add Type Paravet in Employee Type")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDeductionPercentOnIncoming, clsFixedParameterCode.AllowDeductionPercentOnIncoming, "0", "0-OFF;1-On: Allow Deduction(%) on incoming Quality")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowLoginType, clsFixedParameterCode.AllowLoginType, "0", "0-OFF;1-On: Allow POS Functionality in ERP")
        ''==============================
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProcurementUploader, clsFixedParameterCode.MilkProcurementUploader, "b12sec2", "Password for milk procurement uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TankerDispatchBulkUploader, clsFixedParameterCode.TankerDispatchBulkUploader, "b12sec2", "Password for Bulk procurement and Tanker Dispatch uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EmptyCanWeight, clsFixedParameterCode.EmptyCanWeight, "12", "Average Weight of Empty milk Can")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinuteInLastVehicleForGateEntry, clsFixedParameterCode.MinuteInLastVehicleForGateEntry, "10", "Average minute for gate entry in")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinuteGateEntryToGrossWeight, clsFixedParameterCode.MinuteGateEntryToGrossWeight, "10", "Average Minute Gate Entry To Gross Weight")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinuteGrossWeightToTareWeight, clsFixedParameterCode.MinuteGrossWeightToTareWeight, "10", "Average Minute Gross Weight To Tare Weight")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOfDaysForMultiInceForSameVSPForSamePayCycle, clsFixedParameterCode.NoOfDaysForMultiInceForSameVSPForSamePayCycle, "0", "0:No of days in Payment Cycle,1:No of Collection days for Incentive applied")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOfDaysForShowFOSSDocument, clsFixedParameterCode.NoOfDaysForShowFOSSDocument, "5", "No of days For Show FOSS Document")

        ''----Balwinder on UDL Plant
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseCounterOnTransactionType, clsFixedParameterCode.PurchaseCounterOnTransactionType, "0", "1:ON;0 OFF Form Prefixe Generation of PO")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopForRepeatedFATSNF, clsFixedParameterCode.StopForRepeatedFATSNF, "1", "Do not pick sample if Previous and current FAT and SNF is same")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SampleFONTSize, clsFixedParameterCode.SampleFONTSize, "15", "Milk sample of grid Font size")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, "0", "1:ON;0 OFF Prefixe Generation of Bulk Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMSPrefix, clsFixedParameterCode.SMSPrefix, "SMS", "SMS Prefix")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMSPrefix, clsFixedParameterCode.InstantSendTheSMS, "0", "1:ON;0:OFF Send SMS While saving in SMS table")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EmailPrefix, clsFixedParameterCode.InstantSendTheEmail, "0", "1:ON;0:OFF Send Email While saving in Email table")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, "0", "1:ON;0 OFF In Milk Purchase Invoice Pick Pending SRN")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatChequeClearDateAsRecoDate, clsFixedParameterCode.TreatChequeClearDateAsRecoDate, "0", "1:ON;0 OFF Treat Cheque Clear Date As Reco Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BookWreckageFromSublocationOrSection, clsFixedParameterCode.BookWreckageFromSublocationOrSection, "1", "1:ON;0 OFF Book Wreckage From Sublocation/Section")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopVSPBillIfSomethingWrong, clsFixedParameterCode.StopVSPBillIfSomethingWrong, "0", "1:ON;0 OFF Stop VSP Bill And Incentive Generation If Something found Wrong")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PDCSetting, clsFixedParameterCode.PDCSetting, "0", "1:ON;0 OFF Doc Date PDC Setting")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRoadPermitNo, clsFixedParameterCode.AllowRoadPermitNo, "0", "1:ON;0 OFF Allow Road Permit No")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowMessgForTDS, clsFixedParameterCode.ShowMessgForTDS, "0", "1:ON;0 OFF Show Message For TDS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsShowTreeView, clsFixedParameterCode.IsShowTreeView, "1", "1:ON;0 Show Tree view control")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowVLCUploaderData, clsFixedParameterCode.ShowVLCUploaderData, "1", "1:ON;0 Show VLC Uploader Data")
        '============added parteek 09/01/2017
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatSnfWhenMilktypeSelect, clsFixedParameterCode.FatSnfWhenMilktypeSelect, "0", "1:ON;0:Off Fat Snf % Allow when milk type selct")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DairyFreshTaxableandNonTaxable, clsFixedParameterCode.DairyFreshTaxableandNonTaxable, "0", "1:ON;0:Off Taxable and Non-Taxable Item")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SMSEMailPassword, clsFixedParameterCode.SMSEMailPassword, "tecxpertsms", "SMS And Email Screen setting Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateNewDocumentOnUploader, clsFixedParameterCode.CreateNewDocumentOnUploader, "0", "Create new document on MCC Procurement uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProcurementBatchPosting, clsFixedParameterCode.MilkProcurementBatchPosting, "0", "Apply Batch Posting of Milk Procuremnt/Shift Uploader")
        InsertDefaultValueFixedParameter(clsFixedParameterType.popupcustomernamewhileupdating, clsFixedParameterCode.popupcustomernamewhileupdating, "0", "pop-up on customer name while updating")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyFEFO, clsFixedParameterCode.ApplyFEFO, "0", "Apply FEFO on transfer in case of Job Work")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnterLocationForJWEStimationOutPackingMaterial, clsFixedParameterCode.EnterLocationForJWEStimationOutPackingMaterial, "001", "EnterLocationForJWEStimationOutPackingMaterial")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateJVofPackingMaterialofJWInwardinJWEstimate, clsFixedParameterCode.CreateJVofPackingMaterialofJWInwardinJWEstimate, "0", "CreateJVofPackingMaterialofJWInwardinJWEstimate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoenterrateIntoJobWorkDispatch, clsFixedParameterCode.AllowtoenterrateIntoJobWorkDispatch, "0", "Allow to enter rate Into Job Work Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PopupJE, clsFixedParameterCode.PopupJE, "0", "Popup Journal Entry")
        'KUNAL > DATE : 23-01-2017 > CLIENT : Sahayog Dairy > ASSIGNED VIA : EMAIL > REQUEST NO : SCMPLREQ000002
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAliasNames, clsFixedParameterCode.ShowAliasNames, "0", "1:ON;0:Off Show Or Hide Dropdown Items's Names Aliases")
        'KUNAL > DATE : 24-01-2017 > CLIENT : Sahayog Dairy > ASSIGNED VIA : EMAIL > REQUEST NO : SCMPLREQ000002
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowFatAndSnfPercentageFields, clsFixedParameterCode.ShowFatAndSnfPercentageFields, "0", "1:ON;0:Off  Show Or Hide Fat And Snf Percentage Fields")
        'KUNAL > DATE : 24-01-2017 > CLIENT : Sahayog Dairy > ASSIGNED VIA : EMAIL > REQUEST NO : SCMPLREQ000002
        InsertDefaultValueFixedParameter(clsFixedParameterType.VehicleFitnessAndInsuranceFields, clsFixedParameterCode.VehicleFitnessAndInsuranceFields, "0", "1:ON;0:Off  Show Vehicle's Fitness and Insurance Related Fields")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DocumentCancel, clsFixedParameterCode.DocumentCancel, "0", "1:ON;0:Off Document Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PICancelUserPwd, clsFixedParameterCode.PICancelUserPwd, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DocumentCancelReturn, clsFixedParameterCode.DocumentCancelReturn, "0", "1:ON;0:Off Document Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CSADocumentCancel, clsFixedParameterCode.CSADocumentCancel, "0", "1:ON;0:Off Document Cancel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FixVSPEMP, clsFixedParameterCode.FixVSPEMP, "0", "If +ve and Non Zero EMP % Will be Fixed and Non changable")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FatSNFStockControl, clsFixedParameterCode.FatSNFStockControl, "0", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckBalanceFromInvMoveSummry, clsFixedParameterCode.CheckBalanceFromInvMoveSummry, "0", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemwiseFatSNFStockControl, clsFixedParameterCode.ItemwiseFatSNFStockControl, "0", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SepratePriceChartForCowMilk, clsFixedParameterCode.SepratePriceChartForCowMilk, "0", "0-Off;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SepratePriceChartForCamelMilk, clsFixedParameterCode.SepratePriceChartForCamelMilk, "0", "0-Off;1-On")
        '================Added by preeti gupta[20/02/2017]==========================
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRoundInFixedAsset, clsFixedParameterCode.AllowRoundInFixedAsset, "Round Type", "Apply Round In Fixed Asset Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDecimalInFixedAsset, clsFixedParameterCode.AllowDecimalInFixedAsset, "2", "Apply Decimal In Fixed Asset Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyStdFATSNFRate, clsFixedParameterCode.ApplyStdFATSNFRate, "1", "Apply FAT/SNF Rate = (Std Rate*FAT/SNFWeightage)/(FAT/SNF Ratio)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OpenPriceChartPlanningScreenOnTotalSolid, clsFixedParameterCode.OpenPriceChartPlanningScreenOnTotalSolid, "1", " 1)Total Solid 2)Standard Price With Deduction")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxFATSNFPerForRate, clsFixedParameterCode.MaxFATPerLimit, "10", "For Rate if FAT % above this limit than It will apply")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxFATSNFPerForRate, clsFixedParameterCode.MaxSNFPerLimit, "10", "For Rate if SNF % above this limit than It will apply")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowZeroQtyFATSNFInOpenMCCShift, clsFixedParameterCode.AllowZeroQtyFATSNFInOpenMCCShift, "1", "If Off then mendatory to fill Qty,FAT% and SNF% ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowZeroQtyFATSNFInCloseMCCShift, clsFixedParameterCode.AllowZeroQtyFATSNFInCloseMCCShift, "1", "If Off then mendatory to fill Qty,FAT% and SNF%")
        InsertDefaultValueFixedParameter(clsFixedParameterType.POLimit, clsFixedParameterCode.POLimit, "5000", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RequiredPOLimit, clsFixedParameterCode.RequiredPOLimit, "0", "If O then off to PO limit Qty functionality")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UnitCostIncreasePurchaseInvoice, clsFixedParameterCode.UnitCostIncreasePurchaseInvoice, "0", "If O then off setting to Increase Purchase Invoice Unit Cost functionality")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PromptMsgForPendingDocIntervel, clsFixedParameterCode.PromptMsgForPendingDocIntervel, "30", "Allow Prompt Message for pending Documnet intervel")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UDLPurchaseOrderthroughAP, clsFixedParameterCode.UDLPurchaseOrderthroughAP, "0", "If O then off UDL Purchase Order through AP invoice when PO type is Services")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateInventorySummaryTable, clsFixedParameterCode.UpdateInventorySummaryTable, "0", "0:off;1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateConsumeEntry, clsFixedParameterCode.CreateConsumeEntry, "0", "0:off;1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOptionforSelectingCapexForFA, clsFixedParameterCode.ShowOptionforSelectingCapexForFA, "0", "0:Off, 1:On; Show option for selecting capex code and capex sub codes on Fixed Asset")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UDLCapexAcquisionEntry, clsFixedParameterCode.UDLCapexAcquisionEntry, "0", "0:Off, 1:On; Show Capex for Acquision Entry in Fixed Asset")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UDLRGPWiseDocument, clsFixedParameterCode.UDLRGPWiseDocument, "0", "If O then off UDL RGP Wise Document Created")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAssetItemOnMiscSale, clsFixedParameterCode.AllowAssetItemOnMiscSale, "1", "0:Off,1:On; when on then asset items show in Misc. Sale,otherwise not seen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TriggerOfGLEntryForWinTable, clsFixedParameterCode.TriggerOfGLEntryForWinTable, "0", "0:Off,1:On; Trigger GL Entry win table.")
        'UDL DATE : 21-04-2017
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowRouteWiseAndVLCWiseReport, clsFixedParameterCode.ShowRouteWiseAndVLCWiseReport, "0", "0:Off,1:On; Setting to show and hide Routewise and VLC wise reports in same screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UOMAtDiarySaleReturn, clsFixedParameterCode.UOMAtDiarySaleReturn, "0", "0:Off,1:On; Setting to UOM At Diary Sale Return ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PayableAmountZeroForFarmerPayment, clsFixedParameterCode.PayableAmountZeroForFarmerPayment, "1", "0:Off,1:On; System will allow to process farmer payment in case of VSP payable amount is less than or equal 0 ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckDocAmountInAPInvoiceEntry, clsFixedParameterCode.CheckDocAmountInAPInvoiceEntry, "0", "0-OFF;1-On: Check Doc Amount For AP Invoice Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, "0", "0-OFF;1-On: Apply TS Price At Bulk Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, "0", "0-OFF;1-On: Allow SNF Not Manditory in Bulk Sale")
        'UDL > DATE : 3-MAY-2017 : CHANGING DECLARED DOCUMENT LIST TO PENDING DOCUMENT LIST OR VICE VERSA
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowPendingDocumentsListScreenOverDeclaredDocumentList, clsFixedParameterCode.ShowPendingDocumentsListScreenOverDeclaredDocumentList, "0", "0-OFF;1-On: Show and Hide Pending Documents List Over Declared DOcuments List")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MannualySetMPUploaderData, clsFixedParameterCode.MannualySetMPUploaderData, "0", "0-OFF;1-On: Mannualy Set MP Uploader Data")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VSPMPDiffrenceOnTSBasis, clsFixedParameterCode.VSPMPDiffrenceOnTSBasis, "0", "0-OFF;1-On: MP And VSP Collection Diffence On Total Solid Basis")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, "0", "0-OFF;1-On: Pick CLR Value Instead Of SNF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.PickPriceFromFATAndSNF, "0", "0-OFF;1-On: Pick Price From FAT And SNF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBF, clsFixedParameterCode.FATDivideBy, "10", "Value to be divide by for getting FAT")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBF, clsFixedParameterCode.SNFDivideBy, "10", "Value to be divide by for getting SNF")
        InsertDefaultValueFixedParameter(clsFixedParameterType.chkGSTTaxGroupValidity, clsFixedParameterCode.chkGSTTaxGroupValidity, "1", "0-OFF;1-On: check GST Tax Group Validity on all Purchase and sale transactions")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowShipToPartyInDairyDispatch, clsFixedParameterCode.ShowShipToPartyInDairyDispatch, "0", "0-OFF;1-On: Show Ship To Party In Dairy Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkQCWithoutCLR, clsFixedParameterCode.BulkQCWithoutCLR, "0", "0-OFF;1-On: Bulk Quality Check Without CLR")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DOTaggingForDairySaleModule, clsFixedParameterCode.DOTaggingForDairySaleModule, "0", "0-OFF;1-On: DO Tagging For Dairy Sale Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFractionInMCCTankerDispatchGrossQty, clsFixedParameterCode.AllowFractionInMCCTankerDispatchGrossQty, "0", "0-OFF;1-On: Allow Fraction In MCC Tanker Dispatch Gross Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseModulePickFixTaxRate, clsFixedParameterCode.PurchaseModulePickFixTaxRate, "0", "0-OFF;1-On: In Purchase Module Pick Fixed Item Wise Tax Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GSTExemptedAmountForNonRegisteredVendor, clsFixedParameterCode.GSTExemptedAmountForNonRegisteredVendor, "5000", "GST Exempted Amount For Non Registered Vendor")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TankerDispatchFinancialImpactInTransferIn, clsFixedParameterCode.TankerDispatchFinancialImpactInTransferIn, "0", "ON-Tanker dispatch GL Entry/Inventory in Transferin;OFF:Tanker dispatch GL Entry/Inventory in Tanker dispatch [Without chamber]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertQtyIntoKG, clsFixedParameterCode.ConvertQtyIntoKG, "0", "0-OFF;1-On: Convert Qty into KG Bulk Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncreaseCrateQtyOnFiftyPercent, clsFixedParameterCode.IncreaseCrateQtyOnFiftyPercent, "0", "0-OFF;1-On Increase Crate Qty On Fifty Percent")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFDeductionMixMilkFATMinValue, clsFixedParameterCode.FATSNFDeductionMixMilkFATMinValue, "3.0", "FAT SNF Deduction Mix Milk FAT Min Value (Should be One Decimal)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFDeductionMixMilkFATMaxValue, clsFixedParameterCode.FATSNFDeductionMixMilkFATMaxValue, "4.5", "FAT SNF Deduction Mix Milk FAT Max Value (Should be One Decimal)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFDeductionMixMilkSNFMinValue, clsFixedParameterCode.FATSNFDeductionMixMilkSNFMinValue, "7.0", "FAT SNF Deduction Mix Milk SNF Min Value (Should be One Decimal)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFDeductionMixMilkSNFMaxValue, clsFixedParameterCode.FATSNFDeductionMixMilkSNFMaxValue, "7.2", "FAT SNF Deduction Mix Milk SNF Max Value (Should be One Decimal)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFDeductionMixMilkDeductionPer, clsFixedParameterCode.FATSNFDeductionMixMilkDeductionPer, "20", "FAT SNF Deduction Mix Milk Deduction Per")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VSPHoldPaymentNotCompanyBank, clsFixedParameterCode.VSPHoldPaymentNotCompanyBank, "0", "0:Off,1:ON,If VSP Company Bank Not Defined Then Hold")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TDSRoundOffAmount, clsFixedParameterCode.TDSRoundOffAmount, "1", "0:OFF(Amount in Ruppes and paisa),1:ON(Amount in Ruppes)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RoundOffPaiseAmount, clsFixedParameterCode.RoundOffPaiseAmount, "0", "0:Off,1:ON,Remove paise and add in Round off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RemoveSavingDocumentWhenPayableAmtZero, clsFixedParameterCode.RemoveSavingDocumentWhenPayableAmtZero, "0", "0:Off,1:ON,Remove Saving Docuemnts (AP invoice-Credit Note) On post when payabale amount is zero")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RecalculatePaymentProcessOnSave, clsFixedParameterCode.RecalculatePaymentProcessOnSave, "0", "0:Off,1:ON,Re-calculate Payment Process On Save")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableInternalTransfer, clsFixedParameterCode.EnableInternalTransfer, "0", "0:Off,1:ON,Enable Internal Transfer for UDL")
        InsertDefaultValueFixedParameter(clsFixedParameterType.JobWorkTransfer, clsFixedParameterCode.JobWorkTransfer, "0", "0:Off,1:ON,Enable Job Work Transfer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FreightProvisionAccount, clsFixedParameterCode.FreightProvisionAccount, "", "Set Freight Provision Account From GL Account")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FreightProvisionAccountInward, clsFixedParameterCode.FreightProvisionAccountInward, "", "Set Freight Inward Provision Account From GL Account")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatUnregisteredVendorAsRegisteredVendor, clsFixedParameterCode.TreatUnregisteredVendorAsRegisteredVendor, "1", "0:Off,1:ON,Treat Unregistered Vendor Same As Registered Vendor")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RecreateConsumptionEntry, clsFixedParameterCode.RecreateConsumptionEntry, "consumption", "Consumption entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SaleRemarksForTruckSheetReport, clsFixedParameterCode.SaleRemarksForTruckSheetReport, "", "Sale Remarks For Truck Sheet Report")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BankRecoHidePWD, clsFixedParameterCode.BankRecoHidePWD, "hide", "Bank Reconsilation hide password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableItemGroupGLMapping, clsFixedParameterCode.EnableItemGroupGLMapping, "0", "0:Off,1:ON,Enable Item Group GL Mapping")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableRackBin, clsFixedParameterCode.EnableRackBin, "0", "0:Off,1:ON,Enable Rack Bin Item")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PartialFADepDays, clsFixedParameterCode.PartialFADepDays, "0", "Enter no of days for which depreciation of the Asset is partially calculated.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RateMultPartialFADepDays, clsFixedParameterCode.RateMultPartialFADepDays, "0", "Enter value up to to decimal places less than 1.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ChangeVehicleOnDairySaleBooking, clsFixedParameterCode.ChangeVehicleOnDairySaleBooking, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VendorSetOffDayWise, clsFixedParameterCode.VendorSetOffDayWise, "0", "0-OFF;1-On At setoff Make payment one paymentt in one transaction date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsAutoStartReading, clsFixedParameterCode.IsAutoStartReading, IIf(clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal, "0", "1"), "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddHighSecurityOnWeighingIntegratedScreen, clsFixedParameterCode.AddHighSecurityOnWeighingIntegratedScreen, "0", "0-OFF;1-On [Open only single Weighment screen + validate zero + stable weighment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HighSecurityStableSeconds, clsFixedParameterCode.HighSecurityStableSeconds, 5, "Weight should be Stable for given seconds")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HighSecurityWeightTolerance, clsFixedParameterCode.HighSecurityWeightTolerance, 10, "Weight tolerance for next reading")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualvehicleOnDairyBooking, clsFixedParameterCode.AllowManualvehicleOnDairyBooking, "0", "0:Off,1:ON,Allow Manual vehicle On Dairy Booking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FreeIndentQtyAfterPOClose, clsFixedParameterCode.FreeIndentQtyAfterPOClose, "0", "0:Off,1:ON,PO Indent Qty Free after PO Closeed")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowFATSNFinPaymentProcess, clsFixedParameterCode.ShowFATSNFinPaymentProcess, "0", "0:Off,1:ON,Show FAT%,FATKG,SNF%,SNFKG in Payment Process")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxRowsInCSVExport, clsFixedParameterCode.MaxRowsInCSVExport, "200000", "Maximum rows to be exported for csv per file")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxRowsInExcelExport, clsFixedParameterCode.MaxRowsInExcelExport, "500000", "Maximum rows to be exported for excel per file")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, clsCommon.EncryptString(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE.AddMonths(4), "dd/MMM/yyyy"), objCommonVar.CurrentCompanyCode), "1")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowAssetBookChangeInTemplate, clsFixedParameterCode.AllowAssetBookChangeInTemplate, "0", "0:off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSMSSendtoSalePerson, clsFixedParameterCode.AllowSMSSendtoSalePerson, "0", "0:off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSMSwhenCustomerCreditLimit, clsFixedParameterCode.AllowSMSwhenCustomerCreditLimit, "0", "0:off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableScreenSelection, clsFixedParameterCode.EnableScreenSelection, "0", "0:off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipJobWorkSRNInPI, clsFixedParameterCode.SkipJobWorkSRNInPI, "0", "0:off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowFatSnfAfterApproval, clsFixedParameterCode.ShowFatSnfAfterApproval, "0", "0:off, 1:On")
        '======Added by preeti Gupta[16/04/2018]==
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMaTransRateOnMultChamberTankerDis, clsFixedParameterCode.ApplyMaTransRateOnMultChamberTankerDis, "0", "Option to enter Manual Transfer Rate On Tanker Dispatch Chamber wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTotalSolidPriceChart, clsFixedParameterCode.ApplyTotalSolidPriceChart, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RequiredMgmtApprovalForRateIncrease, clsFixedParameterCode.RequiredMgmtApprovalForRateIncrease, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoRoundOffSeprateAccountOnVendorTransaction, clsFixedParameterCode.AutoRoundOffSeprateAccountOnVendorTransaction, "0", "0-OFF;1-On Option ")
        '===================Added by preeti Gupta Against Ticket No[]=============================
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatCRATEAsItems, clsFixedParameterCode.TreatCRATEAsItems, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TreatCANAsItems, clsFixedParameterCode.TreatCANAsItems, "0", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotShowDairyTypeItems, clsFixedParameterCode.DoNotShowDairyTypeItems, "0", "0-OFF;1-On Option ")
        '========================================================================================================================================================
        InsertDefaultValueFixedParameter(clsFixedParameterType.PasswordRules, clsFixedParameterCode.PasswordRules, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AlwaysVSPDefaulter, clsFixedParameterCode.AlwaysVSPDefaulter, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RejectedMilkSendToRejectLocation, clsFixedParameterCode.RejectedMilkSendToRejectLocation, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOfPreNxtDayToPickAvgFATSNF, clsFixedParameterCode.NoOfPreNxtDayToPickAvgFATSNF, "0", "To Get Avg FAT/SNF % First Previous then Next ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoGeneratePrefix, clsFixedParameterCode.AutoGeneratePrefix, "0", "0-OFF;1-On 1.Pick Last Year Prefix 2.Auto Make")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SingleUserParticularDairyBookingEdit, clsFixedParameterCode.SingleUserParticularDairyBookingEdit, "0", "0-OFF;1-On Option ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DBUnlock, clsFixedParameterCode.DairyBookingUnlock, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxReceiveSNFPer, clsFixedParameterCode.MaxReceiveSNFPer, "0", "Maximum receive SNF% from Milk Sample, if 0 it will not work")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MailForAdvancePaymentTerm, clsFixedParameterCode.MailForAdvancePaymentTerm, "0", "0-OFF;1-Mail for advance payment terms")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DonotCheckAvgQtyOnDairyBooking, clsFixedParameterCode.DonotCheckAvgQtyOnDairyBooking, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.leakagededuction_freshsale, clsFixedParameterCode.leakagededuction_freshsale, "0", "Leakage Deduction Fresh Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FirstGateOutProcessForMCCBulkProcument, clsFixedParameterCode.FirstGateOutProcessForMCCBulkProcument, "0", "0-OFF;1-On Option")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCBulkProcumentSecurityGateOut, clsFixedParameterCode.MCCBulkProcumentSecurityGateOut, "0", "0-OFF;1-On Option For Security Gate Out")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableDistributorSubsidy, clsFixedParameterCode.EnableDistributorSubsidy, "0", "0-OFF;1-On Enable Distributor Subsidy")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotAllowSavePOWhenQtyNRateZero, clsFixedParameterCode.DoNotAllowSavePOWhenQtyNRateZero, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowOnlyProductionItemInAdRemove, clsFixedParameterCode.ShowOnlyProductionItemInAdRemove, "0", "0-OFF;1-On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PeriodofSubsidyCreditNote, clsFixedParameterCode.PeriodofSubsidyCreditNote, "0", "Subsidy Credit Note Period")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CHADetailsMandatoryOnExportSale, clsFixedParameterCode.CHADetailsMandatoryOnExportSale, "1", "1-On;0-off CHA Details Mandatory On ExportSale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.KnockOffFATSNFKG, clsFixedParameterCode.KnockOffFATSNFKG, "0", "1-On;0-off Knock Off FAT SNF KG")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateLoadINSlipVehicleWise, clsFixedParameterCode.CreateLoadINSlipVehicleWise, "0", "1-On;0-off Create LoadIN Slip Vehicle Wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RouteCodeNotMandatoryOnLoadINSlip, clsFixedParameterCode.RouteCodeNotMandatoryOnLoadINSlip, "0", "1-On;0-off Route Code Not Mandatory On Loadin Slip")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateReceiveddairyRouteWise, clsFixedParameterCode.CrateReceiveddairyRouteWise, "0", "1-On;0-off Crate Received dairy Route Wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RequiredFinalQCofstandardization, clsFixedParameterCode.RequiredFinalQCofstandardization, "0", "1-On;0-off Quality Check After Standardization")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickOnAccountPaymentForAdvanceKnockOff, clsFixedParameterCode.PickOnAccountPaymentForAdvanceKnockOff, "1", "1-On;0-off Pick OnAccount Payment For Advance Knock Off")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickProductCostFromItemUOMDetail, clsFixedParameterCode.PickProductCostFromItemUOMDetail, "0", "1-On;0-off Pick Product Cost From Item UOM Detail")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, "0", "1-On;0-off Allow Negative Stock In Dairy Production")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SkipLockLocation, clsFixedParameterCode.SkipLockLocation, "0", "1-On;0-off Skip Lock Location")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemStructureMandatoryOnWeightConversion, clsFixedParameterCode.ItemStructureMandatoryOnWeightConversion, "0", "1-On;0-off Item StructurevMandatory On Weight Conversion")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VillageDiffrenceReportValueWithTwoDecimalPoint, clsFixedParameterCode.VillageDiffrenceReportValueWithTwoDecimalPoint, "0", "1-On;0-off Village Diff Report Value With Two Decimal Point")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RequiredFinalQCofProductionEntry, clsFixedParameterCode.RequiredFinalQCofProductionEntry, "0", "1-On;0-off Quality Check After Production Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllPendingDOIrrespectiveOfDeliveryDate, clsFixedParameterCode.ShowAllPendingDOIrrespectiveOfDeliveryDate, "0", "1-On;0-off Show All Pendening DO Irrespective Of DeliveryDate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickFATSNFPerFromStock, clsFixedParameterCode.PickFATSNFPerFromStock, "0", "1-On;0-off Pick FAT SNF Per From Stock")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowBulkDispatchQtyInLtr, clsFixedParameterCode.ShowBulkDispatchQtyInLtr, "0", "1-On;0-off Show Bulk Dispatch Qty In Ltr")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotIncludeIncentiveInMilkPurchaseInvoice, clsFixedParameterCode.DoNotIncludeIncentiveInMilkPurchaseInvoice, "0", "1-On;0-off If On Incetive amount Not calculate in Milk Puchase Invoice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotConsiderTheFutureDateOfAdvancePayment, clsFixedParameterCode.DoNotConsiderTheFutureDateOfAdvancePayment, "0", "1-On;0-off if On Then Advance payment Date is less then To date of payment cycle.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateEmpCodeAsPerEmployeeBasisType, clsFixedParameterCode.CreateEmpCodeAsPerEmployeeBasisType, "0", "1-On;0-off Create Emp Code As Per Employee Basis Type")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FATSNFRateMandatory, clsFixedParameterCode.FATSNFRateMandatory, "1", "1-On;0-off FAT SNF Rate Mandatory")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionOrStandAccordingToItemType, clsFixedParameterCode.ProductionOrStandAccordingToItemType, "0", "1-On;0-off Item Type[F] Production Entry,Item Type[S] Standardization")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PwdOpenOnMainGLAccountAfterSave, clsFixedParameterCode.PwdOpenOnMainGLAccountAfterSave, "tecxpert@aftersave", "Please enter password to update Main GL Account Desc or Sub Group as it has already used in Create Account")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseProductFATSNFKgForEstimationCost, clsFixedParameterCode.UseProductFATSNFKgForEstimationCost, "0", "0-Off-Issued FAT SNF KG;1-ON-Product FAT SNF KG For Estimation Formula")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowItemInCaseofNonInventory, clsFixedParameterCode.ShowItemInCaseofNonInventory, "0", "0-Off-Show Item In Case of Non Inventory(Work Order)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionFATRateTollerance, clsFixedParameterCode.ProductionFATRateTollerance, "0", "Dairy Production FAT Rate Tollerance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionSNFRateTollerance, clsFixedParameterCode.ProductionSNFRateTollerance, "0", "Dairy Production SNF Rate Tollerance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxFATPerLimit, clsFixedParameterCode.MaxFATPerLimitOwnDCS, "0", "0-OFF.Max FAT % limit on Milk Sampling OwnDCS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxSNFPerLimit, clsFixedParameterCode.MaxSNFPerLimitOwnDCS, "0", "0-OFF.Max SNF % limit on Milk Sampling OwnDCS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxFATPerLimit, clsFixedParameterCode.MaxFATPerLimit, "0", "0-OFF.Max FAT % limit on Milk Sampling")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaxSNFPerLimit, clsFixedParameterCode.MaxSNFPerLimit, "0", "0-OFF.Max SNF % limit on Milk Sampling")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinFATPerLimit, clsFixedParameterCode.MinFATPerLimit, "3.5", "0-OFF.Min FAT % limit on Milk Sampling")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MinSNFPerLimit, clsFixedParameterCode.MinSNFPerLimit, "8.1", "0-OFF.Min SNF % limit on Milk Sampling")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LockDate, clsFixedParameterCode.LockDate, "", "Journal Entry will stop before this date [dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTransFATSNFRateForCalculateFATSNFRate, clsFixedParameterCode.ApplyTransFATSNFRateForCalculateFATSNFRate, "1", "0-OFF;1-ON Apply Trans FAT SNF Rate For Calculate FAT SNF Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GrossWeightChangePWD, clsFixedParameterCode.GrossWeightChangePWD, "mdlf@123", "Pwd To Change Gross Weighment PWD")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MachineIntegrationInGeneralWeighment, clsFixedParameterCode.MachineIntegrationInGeneralWeighment, "0", "Weighing machine integrate in General Weighment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FillGeneralWeighmentDetailsByJobworkTypeGateInNo, clsFixedParameterCode.FillGeneralWeighmentDetailsByJobworkTypeGateInNo, "0", "Fill General Weighment By Jobwork Type Gate In No")
        ' Public Const GateOutTankerNoAfterGeneralWeighment As String = "Gate Out Tanker No After General Weighment"
        InsertDefaultValueFixedParameter(clsFixedParameterType.GateOutTankerNoAfterGeneralWeighment, clsFixedParameterCode.GateOutTankerNoAfterGeneralWeighment, "0", "Tanker of JobWork (Bulk Procurement) will reflect on Gate Out After General Weighment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFCustomerForImportExport, clsFixedParameterCode.NoOFCustomerForImportExport, "10", "No Of Customer For Import Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFIncentiveForMPImportExport, clsFixedParameterCode.NoOFIncentiveForMPImportExport, "5", "No Of Incentive For Import Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFSlabForImportExport, clsFixedParameterCode.NoOFSlabForImportExport, "10", "No Of Slab For Import Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFItemStructureForImportExport, clsFixedParameterCode.NoOFItemStructureForImportExport, "10", "No Of Item Structure For Import Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckUnpostedPaymentProcess, clsFixedParameterCode.CheckUnpostedPaymentProcess, "1", "Check Unposted payment process before save")
        ' Public Const NoOFSavingCodeForImportExport As String = "No OF Saving Code For Import Export "
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFSavingCodeForImportExport, clsFixedParameterCode.NoOFSavingCodeForImportExport, "10", "No Of Saving Code For Import Export")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateToLTR, clsFixedParameterCode.CrateToLTR, "10", "Crate To LTR In Report")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CanToLTR, clsFixedParameterCode.CanToLTR, "40", "Can To LTR In Report")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProdcutionDoNotCheckForwardDocuments, clsFixedParameterCode.ProdcutionDoNotCheckForwardDocuments, "0", "On Reverse Do Not Check Forward Document Created")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PurchaseDoNotCheckForwardDocuments, clsFixedParameterCode.PurchaseDoNotCheckForwardDocuments, "0", "On Reverse Do Not Check Forward Document Created")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotStopOnItemBalanceExceptionStoreAdjustment, clsFixedParameterCode.DoNotStopOnItemBalanceExceptionStoreAdjustment, "0", "Do Not Stop On Item Balance Exception For Store Adjustment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ProductionIssueQtyTollerance, clsFixedParameterCode.ProductionIssueQtyTollerance, "0", "[0:Off else ON]Tollerance of Issue Qty vs Required Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NopreviousDaysInSaleVSReceipt, clsFixedParameterCode.NopreviousDaysInSaleVSReceipt, "2", "No of previous days in Sale vs Receipt")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SelectGLInProftAndLossPerforma, clsFixedParameterCode.SelectGLInProftAndLossPerforma, "0", "0:Off;1:On,select GL In Perfoma")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SelectGLInBalanceSheetPerforma, clsFixedParameterCode.SelectGLInBalanceSheetPerforma, "0", "0:Off;1:On,select GL In Perfoma")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BalanceSheetPerformaWithFormula, clsFixedParameterCode.BalanceSheetPerformaWithFormula, "0", "0:Off;1:On,select Formula balance sheet performa")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SelectGLInCashFlowPerforma, clsFixedParameterCode.SelectGLInCashFlowPerforma, "0", "0:Off;1:On,select GL In Perfoma")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkProcurementApplyTotalSoidRate, clsFixedParameterCode.BulkProcurementApplyTotalSoidRate, "0", "0:Off;1:On,Bulk Procurement Apply Total Solid Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateLtrQtyFromKGQtyByCLR, clsFixedParameterCode.CalculateLtrQtyFromKGQtyByCLR, "0", "0:Off;1:On,Calculate Ltr Qty From KG Qty By CLR [Not From Item Master UOM]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCalculateWeightInLtr, clsFixedParameterCode.ApplyCalculateWeightInLtr, "0", "0:Off;1:On,Price to be apply on Ltr Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTransportChargeAddInActualAmount, clsFixedParameterCode.ApplyTransportChargeAddInActualAmount, "0", "0:Off;1:On,Transport Charge Add In Actual Amount")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateProvisionOnGateePass, clsFixedParameterCode.CalculateProvisionOnGateePass, "0", "0:Off;1:On,Calculate Provision On GateePass")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateSNFFromCLRForMCCMilk, clsFixedParameterCode.CalculateSNFFromCLRForMCCMilk, "0", "0:Off;1:On,Calculate SNF From CLR For MCC Milk as calculate for contractor milk")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, clsFixedParameterCode.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, "0", "0:Off;1:On,Tag Exempted tax group in case of Bank Charges in Payment Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SNFFromCLRAndCorrectionFactorInJWIEst, clsFixedParameterCode.SNFFromCLRAndCorrectionFactorInJWIEst, "0", "0:Off;1:On,Calculate SNF From CLR And correction Factor in Job work Estimation screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoCalculateProduceQty, clsFixedParameterCode.AutoCalculateProduceQty, "0", "0:Off;1:On,Auto Calculate Produce Qty in Job work Estimation screen.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NoOFCustomerForImportExportOnRouteMaster, clsFixedParameterCode.NoOFCustomerForImportExportOnRouteMaster, "10", "No OF Customer For Import Export On Route Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SyncedMccToServerStartDateForEmailSms, clsFixedParameterCode.SyncedMccToServerStartDateForEmailSms, "", " Synced Mcc To Server Start Date For Email/SMS (dd/MM/yyyy)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateReceiveddairyCustomerWise, clsFixedParameterCode.CrateReceiveddairyCustomerWise, "0", "1-On;0-off Crate Received dairy Customer Wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaintainLogForImporperSample, clsFixedParameterCode.MaintainLogForImporperSample, "0", "1-On;0-off Maintain Log For Imporper Sample")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PFCalculationOnFormulaHead, clsFixedParameterCode.PFCalculationOnFormulaHead, "0", "1-On;0-off If On Then PF compair with FORMULA HEAD else HEAD VALUE ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MODValueForFAT, clsFixedParameterCode.MODValueForFAT, "5", "Check value of FAT % After MOD operation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoFillSameLocationInGrid, clsFixedParameterCode.AutoFillSameLocationInGrid, "1", "Auto Fill Same Location In Grid")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyNotShowJobWorkTypeTanker, clsFixedParameterCode.ApplyNotShowJobWorkTypeTanker, "0", "Do Not Show JobWork Type Tanker in Tanker status Report")
        'sanjay
        InsertDefaultValueFixedParameter(clsFixedParameterType.ExportToDefineLocation, clsFixedParameterCode.ExportToDefineLocation, "0", "0-OFF;1-ON: Export To Define Location")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowOneFormatByDefaultForPrint, clsFixedParameterCode.AllowOneFormatByDefaultForPrint, "0", "0:Off, 1:On; Allow printing 1 formats by default ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CostCenterAndHirerachyCodeUpdateAfterPost, clsFixedParameterCode.CostCenterAndHirerachyCodeUpdateAfterPost, "0", "0:Off, 1:On; Cost Center And Hirerachy Code Update After Post")
        'sanjay
        Dim DescriptionForCheckUniqueDocumentCode As String
        Dim qry1 As String = "select COMP_CODE from TSPL_COMPANY_MASTER where Is_Main_Company=1"
        Dim Comp_Code As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry1))
        If ((clsCommon.CompairString(Comp_Code, "001") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "ADVANTEK") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "ALPHA") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "BHAD") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "DCC") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "GDFPL") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "GHO") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "GK") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "GMD") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "Jakson") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "KL") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "MPD") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "PSFI") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "SAHAYOG") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "SPMMD") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "UDL") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "Viney") = CompairStringResult.Equal) Or (clsCommon.CompairString(Comp_Code, "WHOLLY") = CompairStringResult.Equal)) Then
            DescriptionForCheckUniqueDocumentCode = "0"
        Else
            DescriptionForCheckUniqueDocumentCode = "1"
        End If
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckUniqueDocumentCode, clsFixedParameterCode.CheckUniqueDocumentCode, DescriptionForCheckUniqueDocumentCode, "Check Unique Document Code")
        'sanjay
        InsertDefaultValueFixedParameter(clsFixedParameterType.JWIRateofFGasPerRM, clsFixedParameterCode.JWIRateofFGasPerRM, "0", "0:Off[Pick Rate From item UOM],1:On[Pick Rate from JWI item price master]")
        ''Should be call at last
        InsertDefaultValueFixedParameter(clsFixedParameterType.TagMultipleRouteWithCustomer, clsFixedParameterCode.TagMultipleRouteWithCustomer, "0", "0:Off[Tag Single Route With Customer],1:On[Tag Multiple Route With Customer]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionOnOpeningAndClosingKM, clsFixedParameterCode.CreateProvisionOnOpeningAndClosingKM, "0", "0:Off[Km From Route Master],1:On[Km From Opening and Closing]")
        If ((clsCommon.CompairString(Comp_Code, "KL") = CompairStringResult.Equal)) Then
            DescriptionForCheckUniqueDocumentCode = "1"
        Else
            DescriptionForCheckUniqueDocumentCode = "0"
        End If
        InsertDefaultValueFixedParameter(clsFixedParameterType.FromLocationStockNotCheckConsumptionLocation, clsFixedParameterCode.FromLocationStockNotCheckConsumptionLocation, DescriptionForCheckUniqueDocumentCode, "0:[Check Consumption Location],1:On[Not Check Consumption Location]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MilkIncetiveByMilkSRN, clsFixedParameterCode.MilkIncetiveByMilkSRN, 1, "0:[Incetive Entry Payment Days By Milk Purchase Invoice],1:On[Incetive Entry Monthly By Milk SRN]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowShowCoumnInCrateReceivedDairy, clsFixedParameterCode.AllowShowCoumnInCrateReceivedDairy, 0, "0:[All Column shown],1:On[Some Column shown]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CalculateLeakageAmount, clsFixedParameterCode.CalculateLeakageAmount, 0, "0:[Not Calculate Leakage Amount],1:On[Calculate Leakage Amount]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowFATSNFPerOnBulkProcInGateIN, clsFixedParameterCode.ShowFATSNFPerOnBulkProcInGateIN, 0, "0:OFF[Not Show FAT%,SNF% column in Gate In Screen In case of Bulk Proc.],1:On[Show FAT%,SNF% column in Gate In Screen In case of Bulk Proc.]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowItemCostMandatoryForStockingUnit, clsFixedParameterCode.AllowItemCostMandatoryForStockingUnit, 0, "1-On;0-off Allow Item Cost Mandatory for Stocking Unit")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ImportMultipleAssetAssembled, clsFixedParameterCode.ImportMultipleAssetAssembled, "0", "0-OFF;1-ON: Import Multiple Asset Assembled")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowEmpCurrentSalaryOnEmployeeSatatusReport, clsFixedParameterCode.ShowEmpCurrentSalaryOnEmployeeSatatusReport, "0", "0-OFF;1-ON:Show Employee Current Salary")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseControlMForHelp, clsFixedParameterCode.UseControlMForHelp, "0", "0-OFF;1-ON:Open Help- This should be on in case of ERP working on Web")
        ''Endo Of Should be call at last
        InsertDefaultValueFixedParameter(clsFixedParameterType.DairyBookingTolleranceQty, clsFixedParameterCode.DairyBookingTolleranceQty, "5", "Dairy Booking Tollerance(+/-) Qty")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VehicleCodeNotMandatoryOnLoadINSlip, clsFixedParameterCode.VehicleCodeNotMandatoryOnLoadINSlip, "0", "0-OFF;1-ON: Vehicle Code Not Mandatory On LoadIN Slip")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyLatestPriceChartWhilecreatingNewVSP, clsFixedParameterCode.ApplyLatestPriceChartWhilecreatingNewVSP, "1", "0-OFF;1-ON: Apply Latest Price Chart While creating New VSP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToPrintInvoiceAfterPosting, clsFixedParameterCode.AllowToPrintInvoiceAfterPosting, "0", "0-OFF;1-Allow To Print Invoice After Posting")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableBankFromMaster, clsFixedParameterCode.EnableBankFromMaster, "0", "0-OFF;1-Enable Bank From Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Autoadjustdebitcreditnote, clsFixedParameterCode.Autoadjustdebitcreditnote, "0", "0-OFF;1-Auto Adjust Debit Credit Note")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateItemMasterWithoutTransactionValidation, clsFixedParameterCode.UpdateItemMasterWithoutTransactionValidation, "0", "0-OFF;1-On Update Item Master Without Transaction Validation[Implement Import/Export] ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddItemAliasInSMS, clsFixedParameterCode.AddItemAliasInSMS, "0", "0-OFF;1-On Add Item Alias In SMS")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemWiseQualityCheckInGeneralPurchase, clsFixedParameterCode.ItemWiseQualityCheckInGeneralPurchase, "0", "0-OFF[Apply Vendor Wise QC Parameter];1-On[Apply Item Wise QC Parameter] ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SendInternalSMSEmailinPurchaseModule, clsFixedParameterCode.SendInternalSMSEmailinPurchaseModule, "0", "0-OFF;1-On Send Internal Mail/SMS to reporting officer in Purchase Module")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoEnterNetWeightManuallyinPOWeighmentScreen, clsFixedParameterCode.AllowtoEnterNetWeightManuallyinPOWeighmentScreen, "0", "0-OFF;1-On Allow to Enter Net Weight Manually in PO Weighment Screen")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowNotificationInMDI, clsFixedParameterCode.ShowNotificationInMDI, "0", "0-OFF;Show Notification In MDI(Enter Interval In Minute)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TIPRateMix, clsFixedParameterCode.TIPRateMix, "0", "Enter Default TIP Rate for Mix")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TIPRateCow, clsFixedParameterCode.TIPRateCow, "0", "Enter Default TIP Rate for Cow")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TIPRateBuffalo, clsFixedParameterCode.TIPRateBuffalo, "0", "Enter Default TIP Rate for Buffalo")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultCustomerGroupCodeforVSP, clsFixedParameterCode.DefaultCustomerGroupCodeforVSP, "", "Default Customer Group Code for VSP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultVendorGroupCodeforVSP, clsFixedParameterCode.DefaultVendorGroupCodeforVSP, "", "Default Vendor Group Code for  VSP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSameTankerNoforPrimarySecondaryTransporter, clsFixedParameterCode.AllowSameTankerNoforPrimarySecondaryTransporter, "0", "0-OFF;1-Allow Same Tanker No for Primary and Secondary Transporter Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllocateToMandatoryonGateOut, clsFixedParameterCode.AllocateToMandatoryonGateOut, "0", "0-OFF;1-Allocate To Mandatory on Gate Out(MCC)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrintTruckSheetAfterGenerate, clsFixedParameterCode.PrintTruckSheetAfterGenerate, "0", "0-OFF;1-Print Truck Sheet After Generate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AskForPwdForOutAdjustmentOnPost, clsFixedParameterCode.AskForPwdForOutAdjustmentOnPost, "", "Set password for Out Adjustment On Post.")
        InsertDefaultValueFixedParameter(clsFixedParameterType.checkStockOfItemTillTransactionDateOnly, clsFixedParameterCode.checkStockOfItemTillTransactionDateOnly, "0", "0-OFF;1-check stock of item till transaction date only")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseProductionPlaningDateForWholeProductionCycle, clsFixedParameterCode.UseProductionPlaningDateForWholeProductionCycle, "0", "0-OFF;1-Use Production Planing Date For Whole Production Cycle")
        InsertDefaultValueFixedParameter(clsFixedParameterType.allowMilkJWOutowordWithAvgFatSNFRate, clsFixedParameterCode.allowMilkJWOutowordWithAvgFatSNFRate, "0", "0-OFF;1-allowMilkJWOutowordWithAvgFatSNFRate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.allowMilkJWOutowordWithAvgFatSNFPerAtInventory, clsFixedParameterCode.allowMilkJWOutowordWithAvgFatSNFPerAtInventory, "0", "0-OFF;1-allowMilkJWOutdWithAvgFatSNFPerAtInv")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateProvisionOfTankerDispatchWithClosingKM, clsFixedParameterCode.CreateProvisionOfTankerDispatchWithClosingKM, "0", "0-OFF;1-On: Create Provision Of Tanker Dispatch With Closing KM")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MaterialSaleInvoiceEnablePrintOnPost, clsFixedParameterCode.MaterialSaleInvoiceEnablePrintOnPost, "0", "0-OFF;1-On: Material Sale Invoice Enable Print On Post")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn, clsFixedParameterCode.DoNotCreateJVOnSameLocationSegmentInTanDisAndMTIn, "0", "0-OFF;1-On: Do Not Create JV On Same Location Segment In Tanker Dispatch And Milk Transfer In")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateItemMasterConversationWithoutValidation, clsFixedParameterCode.UpdateItemMasterConversationWithoutValidation, "0", "0-OFF;1-On: Update Item Master Conversation Without Validation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TollTaxMaster, clsFixedParameterCode.TollTaxMaster, "1", "0-OFF;1-On: Allow Toll Tax Value By Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.InDocMandatoryOnInternalTransfer, clsFixedParameterCode.InDocMandatoryOnInternalTransfer, "0", "0-OFF;1-On: In Doc Mandatory On Internal Transfer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowTransferInAfterGatePassOnly, clsFixedParameterCode.AllowTransferInAfterGatePassOnly, "0", "0-OFF;1-On: Allow Transfer In After Gate Pass Only")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableItemShortDescriptionInBooking, clsFixedParameterCode.EnableItemShortDescriptionInBooking, "0", "0-OFF;1-On: Enable Item Short Description In Booking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowDuplicateItemShortDescriptionInItemMaster, clsFixedParameterCode.AllowDuplicateItemShortDescriptionInItemMaster, "0", "0-OFF;1-On: Allow Duplicate Item Short Description In Item Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeleteMccMilkShiftPassword, clsFixedParameterCode.DeleteMccMilkShiftPassword, "admin@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotCreatePaymentWhileSalaryGeneration, clsFixedParameterCode.DoNotCreatePaymentWhileSalaryGeneration, "0", "0-OFF;1-On: Do Not Create Payment While Salary Generation ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MccPlantSelectionOptionInMccTankerGateOut, clsFixedParameterCode.MccPlantSelectionOptionInMccTankerGateOut, "0", "0-OFF;1-On: Mcc Plant Selection Option In Mcc Tanker Gate Out")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableTankerNoInMccTankerDispWithMccTankerGateOut, clsFixedParameterCode.EnableTankerNoInMccTankerDispWithMccTankerGateOut, "0", "0-OFF;1-On: Enable Tanker No In Mcc Tanker Dispatch along With Mcc Tanker Gate Out")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToEnterSnfAtPlantInMccTankerDispatch, clsFixedParameterCode.AllowToEnterSnfAtPlantInMccTankerDispatch, "0", "0-OFF;1-On: Allow To Enter Snf At Plant In Mcc Tanker Dispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DateOfEInvoiceImplementation, clsFixedParameterCode.DateOfEInvoiceImplementation, "", "Date Of E-Invoice Implementation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowOnlyOneIssueAgainstStoreRequisition, clsFixedParameterCode.AllowOnlyOneIssueAgainstStoreRequisition, "0", "0-OFF;1-ON Allow Only One Issue Against Store Requisition")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, clsFixedParameterCode.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, "0", "0-OFF;1-ON Use VLC Uploader Code MP Uploader Code In MCC Procurement")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowNotificationWithoutSMSAPP, clsFixedParameterCode.ShowNotificationWithoutSMSAPP, "0", "0-OFF;1-ON Show Notification Without SMS APP")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SetNotificationRefreshTimeInMinutes, clsFixedParameterCode.SetNotificationRefreshTimeInMinutes, "15", "Set Notification Refresh Time In Minutes")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GenerateEWayBillWithEInvoice, clsFixedParameterCode.GenerateEWayBillWithEInvoice, "0", "0-OFF;1-ON Generate EWay Bill With EInvoice")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VillageDataReverse, clsFixedParameterCode.VillageDataReverse, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowTankerWithoutCheckingAnyValidation, clsFixedParameterCode.ShowTankerWithoutCheckingAnyValidation, "0", "0-OFF;1-ON Show Tanker Without Checking Any Validation")
        'InsertDefaultValueFixedParameter(clsFixedParameterType.AllocatedTankerGateOut, clsFixedParameterCode.AllocatedTankerGateOut, "0", "0-OFF;1-ON Allocated Tanker Gate Out")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToCreateNoOfBookingPerDay, clsFixedParameterCode.AllowToCreateNoOfBookingPerDay, "0", "0-OFF;Enter No - Allow To Create No Of Booking Per Day")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnterWeightManuallyOnWeighmentInGrid, clsFixedParameterCode.EnterWeightManuallyOnWeighmentInGrid, "0", "0-OFF;1-ON Enter Weight Manually On Weighment In Grid")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowSameChequeNoForMultiplePaymentEntry, clsFixedParameterCode.AllowSameChequeNoForMultiplePaymentEntry, "0", "0-OFF;1-ON Allow Same Cheque No For Multiple Payment Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowGenerateReferenceNoForBulkGateEntry, clsFixedParameterCode.AllowGenerateReferenceNoForBulkGateEntry, "0", "0-OFF;1-ON Allow Generate Reference No For Bulk Gate Entry")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateNeftuploaderPlantWise, clsFixedParameterCode.CreateNeftuploaderPlantWise, "0", "0-OFF;1-ON Create Neftuploader Plant Wise")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBankTransferAgainstMilkBill, clsFixedParameterCode.AllowBankTransferAgainstMilkBill, "0", "0-OFF;1-ON Allow Bank Transfer Against Milk Bill")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowTCSAmountOnBookingForOtherCustomer, clsFixedParameterCode.ShowTCSAmountOnBookingForOtherCustomer, "0", "0-OFF;1-ON ShowTCSAmountOnBookingForOtherCustomer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowCheckExcludeProvisionBank, clsFixedParameterCode.ShowCheckExcludeProvisionBank, "0", "0-OFF;1-ON Show Check Exclude Provision Bank")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ChangeLeaveDescriptionOnSalarySlip, clsFixedParameterCode.ChangeLeaveDescriptionOnSalarySlip, "0", "0-OFF;1-Change Leave Description On Salary Slip")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDepartmentWiseDataVisibleInDepartmentIndent, clsFixedParameterCode.ApplyDepartmentWiseDataVisibleInDepartmentIndent, "0", "0-OFF;1-ApplyDepartmentWiseDataVisibleInDepartmentIndent")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateMapPayHeadsToSalaStructurePassword, clsFixedParameterCode.UpdateMapPayHeadsToSalaStructurePassword, "", "Update Map Pay Heads To Sala Structure Password")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyFinancialCostCenter, clsFixedParameterCode.ApplyFinancialCostCenter, "", "Apply Financial Cost Center")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SalarySlipLeaveStatusOnTheBasisOfCalendarYear, clsFixedParameterCode.SalarySlipLeaveStatusOnTheBasisOfCalendarYear, "0", "0-OFF;1-Salary Slip Leave Status On The Basis Of Calendar Year")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCalculationOnRouteLenth, clsFixedParameterCode.ApplyCalculationOnRouteLenth, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableExportExcelOnIncentiveEntry, clsFixedParameterCode.EnableExportExcelOnIncentiveEntry, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowBankSectionEnableOnMCCMaster, clsFixedParameterCode.AllowBankSectionEnableOnMCCMaster, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConsiderPreviousCurrentFYForTCSTaxVendOutstanding, clsFixedParameterCode.ConsiderPreviousCurrentFYForTCSTaxVendOutstanding, "0", "[TCS Purchase] Consider Previous Current FY For TCS Tax Vendor Outstanding")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AmountToCheckVendorOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckVendorOutstandingForTCSTax, "0", "[TCS Purchase] Amount To Check Vendor Outstanding For TCS Tax")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowtoChangeTCSBaseAmountPurchase, clsFixedParameterCode.AllowtoChangeTCSBaseAmountPurchase, "0", "[TCS Purchase]Allow to Change TCS Base Amount In Purchase")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowFatPerInanynumberofMultipesonBulkQC, clsFixedParameterCode.AllowFatPerInanynumberofMultipesonBulkQC, "0", "Allow Fat Per In any number of Multipes on Bulk QC")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowCanInformationintoGridForTankerDispatch, clsFixedParameterCode.AllowCanInformationintoGridForTankerDispatch, "0", "AllowCanInformationintoGridForTankerDispatch")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateSeparateInvoiceforeachrowinCansale, clsFixedParameterCode.CreateSeparateInvoiceforeachrowinCansale, "0", "Create Separate Invoice for each row in Can sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSRateforVendorWithPanNo, clsFixedParameterCode.TCSRateforVendorWithPanNo, "0", "[TCS Purchase] TCS Rate for Vendor With PAN No")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAvailableQtyOnDairyBooking, clsFixedParameterCode.ShowAvailableQtyOnDairyBooking, "0", "ShowAvailableQtyOnDairyBooking")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowTransferVSPAmtToFarmerinVCGL, clsFixedParameterCode.AllowTransferVSPAmtToFarmerinVCGL, "0", "Allow to Transfer VSP Amount To Farmer in VCGL")
        InsertDefaultValueFixedParameter(clsFixedParameterType.TCSRateforVendorWithoutPanNo, clsFixedParameterCode.TCSRateforVendorWithoutPanNo, "0", "[TCS Purchase] TCS Rate for Vendor Without PAN No")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTDSValidationFrom, clsFixedParameterCode.ApplyTDSValidationFrom, "", "Apply TDS Validations From [dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DateOfDynamicQRCodeForB2CInvoiceImplementation, clsFixedParameterCode.DateOfDynamicQRCodeForB2CInvoiceImplementation, "", "DateOfDynamicQRCodeForB2CInvoiceImplementation")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSAddDedRODecimalPlace, clsFixedParameterCode.DCSAddDedRODecimalPlace, "2", "Decimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSAddDedROIncreaseAfter, clsFixedParameterCode.DCSAddDedROIncreaseAfter, "5", "Increase when Value above or Equals")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DCSAddDedROHeaderLevel, clsFixedParameterCode.DCSAddDedROHeaderLevel, "5", "RO At Amount At Header level")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VSPBillDocumentToBeAddedInMilkCost, clsFixedParameterCode.VSPBillDocumentToBeAddedInMilkCost, "0", "0:OFF;1:ON VSP Bill Document To Be Added In Milk Cost")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, clsFixedParameterCode.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, "0", "0:OFF;1:ON Show Vehicle No Separately In Primary Transporter Vehicle Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRoundDownAmtForMCCDateWiseChilling, clsFixedParameterCode.AllowRoundDownAmtForMCCDateWiseChilling, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableTDSforServiceVendorSeparately, clsFixedParameterCode.EnableTDSforServiceVendorSeparately, "0", "0:OFF;1:ON Enable TDS for Service Vendor Separately In Vendor Master")
        ' ' 
        InsertDefaultValueFixedParameter(clsFixedParameterType.CtreateJEOfVspAssetIssueAndReturn, clsFixedParameterCode.CtreateJEOfVspAssetIssueAndReturn, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.createDebitNoteOnAssetIssue, clsFixedParameterCode.createDebitNoteOnAssetIssue, "0", "create AP Debit Note On Asset Issue")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickAvgCostonAssetissue, clsFixedParameterCode.PickAvgCostonAssetissue, "0", "Pick Avg Cost on Asset issue")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCBioSyncDate, clsFixedParameterCode.MCCBioSyncDate, "", "MCC Bio Sync Date [dd/MMM/yyyy]")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BookingMobileAppChangeorderByBookingQty, clsFixedParameterCode.BookingMobileAppChangeorderByBookingQty, "0", "0:1,1-> PrevBookingQty desc")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BookingMobileAppSetNxtDateofBookingOrder, clsFixedParameterCode.BookingMobileAppSetNxtDateofBookingOrder, "0", "0:1,1-> Pick Nxt Date of Booking Order")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTCSAmtOnAbstractReportDotMatrix, clsFixedParameterCode.ApplyTCSAmtOnAbstractReportDotMatrix, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.NotAllowDuplicatePANOnPrimaryTransporter, clsFixedParameterCode.NotAllowDuplicatePANOnPrimaryTransporter, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, clsFixedParameterCode.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, "0", "0-OFF;1-ON")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowEarlyRouteOnTruckSheet, clsFixedParameterCode.ShowEarlyRouteOnTruckSheet, "0", "0-OFF;1-ON Show Early Route On Truck Sheet")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowOutEntryOnCrateReceivedDairyForAdjustment, clsFixedParameterCode.AllowOutEntryOnCrateReceivedDairyForAdjustment, "0", "0-OFF;1-ON Allow Out Entry On Crate Received Dairy For Adjustment")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CanSaleAvgFATSNFPer, clsFixedParameterCode.CanSaleAvgFATSNFPer, "0", "0-OFF;1-ON Pick Average FAT/SNF % in Can Sale")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkMilkFATSNFKGDecimalPlaces, clsFixedParameterCode.BulkMilkFATSNFKGDecimalPlaces, "3", "Bulk Milk SRN FAT/SNF Kg Decimal places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkMilkFATSNFAmtDecimalPlaces, clsFixedParameterCode.BulkMilkFATSNFAmtDecimalPlaces, "2", "Bulk Milk SRN FAT/SNF Amt Decimal places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.BulkMilkConsiderAllParametersForIncetive, clsFixedParameterCode.BulkMilkConsiderAllParametersForIncetive, "0", "Bulk Milk All QC Parameter Satisfy Range For Incentive")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MCCMaterialSaleFarmerReverse, clsFixedParameterCode.MCCMaterialSaleFarmerReverse, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LocalSaleAllowedPer, clsFixedParameterCode.LocalSaleAllowedPer, "3", "Allowed Local Sale %")
        InsertDefaultValueFixedParameter(clsFixedParameterType.LocalSaleAllowedRate, clsFixedParameterCode.LocalSaleAllowedRate, "3", "Local Sale Rate")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MultiDairyGatePassReversePWD, clsFixedParameterCode.MultiDairyGatePassReversePWD, "tecxpert@123", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowStatusItemWiseInPendingRequisitionRpt, clsFixedParameterCode.ShowStatusItemWiseInPendingRequisitionRpt, "0", "0:Off, 1:On; Show Status Item Wise In Pending Requisition Report")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DoNotCheckAnyValidationOnVendorInactive, clsFixedParameterCode.DoNotCheckAnyValidationOnVendorInactive, "0", "0:Off, 1:On; Do Not Check Any Validation On Vendor Inactive")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SupportHindiFont, clsFixedParameterCode.SupportHindiFont, "0", "0:Off, 1:On; Support Hindi Font")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertHindiFontLastDateTime, clsFixedParameterCode.ConvertHindiFontLastDateTime, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMilkPouchPrint, clsFixedParameterCode.ApplyMilkPouchPrint, "0", "0:Off, 1:On; ApplyMilkPouchPrint")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UserWiseRouteMapping, clsFixedParameterCode.UserWiseRouteMapping, "0", "0:Off, 1:On; User Wise Route Mapping")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateMCCTankerGateOutBasedOnBulkRouteMaster, clsFixedParameterCode.CreateMCCTankerGateOutBasedOnBulkRouteMaster, "0", "0:Off, 1:On; Create MCC Tanker Gate Out Based On Bulk Route Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyAutoLoginIdCreateOfCustomer, clsFixedParameterCode.ApplyAutoLoginIdCreateOfCustomer, "0", "0:Off, 1:On; Apply Auto Login Id Create Of Customer")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultLocationOfUserMaster, clsFixedParameterCode.DefaultLocationOfUserMaster, " ", "DefaultLocationOfUserMaster")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateGatePassFromDemand, clsFixedParameterCode.CreateGatePassFromDemand, "0", "0:Off, 1:On; Create Gate Pass From Demand")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister, clsFixedParameterCode.PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister, "1", "0:Off, 1:On; Pick Fat/Snf KG From Bulk Milk SRN In Bulk Milk Register")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AadharNoMandatoryOnEmpMaster, clsFixedParameterCode.AadharNoMandatoryOnEmpMaster, "0", "0-OFF, 1:On; Aadhar No Mandatory On Employee Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SuperUserCustomer, clsFixedParameterCode.SuperUserCustomer, "1", "0-OFF, 1:On; Remove InActive Customer from User Master (SuperUser)")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UploadMultipleMasterPwd, clsFixedParameterCode.UploadMultipleMasterPwd, "xpert", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDefaultsInMaster, clsFixedParameterCode.ApplyDefaultsInMaster, "0", "0:Off, 1:On; Apply Defaults In Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IncentiveAccNoMandatoryInMPMaster, clsFixedParameterCode.IncentiveAccNoMandatoryInMPMaster, "1", "0:Off, 1:On; Incentive Account No Mandatory In MP Master")
        InsertDefaultValueFixedParameter(clsFixedParameterType.Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster, clsFixedParameterCode.Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster, "0", "0:Off, 1:On;Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MultipleFinderFillAuto, clsFixedParameterCode.MultipleFinderFillAuto, "0", "0:Off, 1:On;MultipleFinderFillAuto")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MandatoryLineNoMaxMinQtyForProductionPlan, clsFixedParameterCode.MandatoryLineNoMaxMinQtyForProductionPlan, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RunProductionBaseOnPercentage, clsFixedParameterCode.RunProductionBaseOnPercentage, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.RCDFCFP, clsFixedParameterCode.RCDFCFP, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyLocationFilterBasedOnPermission, clsFixedParameterCode.ApplyLocationFilterBasedOnPermission, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MandatoryPDFFileMilkPricePlan, clsFixedParameterCode.MandatoryPDFFileMilkPricePlan, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoClosePOBasedOnSRNQtyWithTolerance, clsFixedParameterCode.AutoClosePOBasedOnSRNQtyWithTolerance, "0", "0:Off, 1:On; Auto Close PO Based On SRN Qty With Tolerance")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMilkTypeBuffaloCowOnPrint, clsFixedParameterCode.ApplyMilkTypeBuffaloCowOnPrint, "1", "0:Off, 1:On; Off:Mixed ; On: Buffalo if FAT > 5 else Cow")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyZoneWiseVSP, clsFixedParameterCode.ApplyZoneWiseVSP, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyStandardProductionVariance, clsFixedParameterCode.ApplyStandardProductionVariance, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ItemCostTolerancePercentage, clsFixedParameterCode.ItemCostTolerancePercentage, "0", "Enter Item Cost Tolerance In Percentage")
        InsertDefaultValueFixedParameter(clsFixedParameterType.HeadLoadDescriptionInPaymentProcessPrint, clsFixedParameterCode.HeadLoadDescriptionInPaymentProcessPrint, "Head Load", "Head Load Description In Payment Process Print")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PrefixForUserMaster, clsFixedParameterCode.PrefixForUserMaster, clsFixedParameter.GetUserCode(), "Prefix For User Master Code")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDemandApproval, clsFixedParameterCode.ApplyDemandApproval, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckCreditLimit, clsFixedParameterCode.CheckCreditLimit, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTolerance, clsFixedParameterCode.ApplyTolerance, "0", "Enter Tolerance Percentage")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SetShiftTimeOut, clsFixedParameterCode.SetShiftTimeOut, "10:05 AM", "Enter Shift Time Out")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SetDefaultShiftTime, clsFixedParameterCode.SetDefaultShiftTime, "04:15 PM", "Enter Default Shift Time")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToSelectShift, clsFixedParameterCode.AllowToSelectShift, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyOrderByNumeric, clsFixedParameterCode.ApplyOrderByNumeric, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyRoundOffZero, clsFixedParameterCode.ApplyRoundOffZero, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableLocation, clsFixedParameterCode.EnableLocation, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableResetDemand, clsFixedParameterCode.EnableResetDemand, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.IsLoadingSlipMandatory, clsFixedParameterCode.IsLoadingSlipMandatory, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CrateCapacityCheck, clsFixedParameterCode.CrateCapacityCheck, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickAllBMC, clsFixedParameterCode.PickAllBMC, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMonthlyYearly, clsFixedParameterCode.ApplyMonthlyYearly, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeparateDemandMilkandProduct, clsFixedParameterCode.SeparateDemandMilkandProduct, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowDemandDoc, clsFixedParameterCode.ShowDemandDoc, "0", "0:Don't Show, 1:Show;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCommission, clsFixedParameterCode.ApplyCommission, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCommissionRateWithTax, clsFixedParameterCode.ApplyCommissionRateWithTax, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisableRouteandVehicle, clsFixedParameterCode.DisableRouteandVehicle, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.UpdateDemandBeforePost, clsFixedParameterCode.UpdateDemandBeforePost, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PricePlanRoundOffTruncate, clsFixedParameterCode.PricePlanRoundOffTruncate, "0", "0:RoundOff, 1:Truncate;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyPricePlanDigitRule, clsFixedParameterCode.ApplyPricePlanDigitRule, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PricePlanRoundOffDigit, clsFixedParameterCode.PricePlanRoundOffDigit, "2", "Enter Roundoff Digit;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DemandIncreaseDecreaseQty, clsFixedParameterCode.DemandIncreaseDecreaseQty, "10", "Limit Increase/Decrease Qty;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DemandIncreaseDecreasePer, clsFixedParameterCode.DemandIncreaseDecreasePer, "10", "Limit Increase/Decrease Per;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowIncreaseDispatchQty, clsFixedParameterCode.AllowIncreaseDispatchQty, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FORPRICE, clsFixedParameterCode.FORPRICE, "0", "Enter Weight in KG ")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowGatePassDemandTripWise, clsFixedParameterCode.AllowGatePassDemandTripWise, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertPouchtoCrate, clsFixedParameterCode.ConvertPouchtoCrate, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DontCreateForPouch, clsFixedParameterCode.DontCreateForPouch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAutoMCCPriceChat, clsFixedParameterCode.CreateAutoMCCPriceChat, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.VehicleNofromDispatch, clsFixedParameterCode.VehicleNofromDispatch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DispatchPriceCodeForCreditCustomer, clsFixedParameterCode.DispatchPriceCodeForCreditCustomer, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyCommissionTPTForCreditCustomer, clsFixedParameterCode.ApplyCommissionTPTForCreditCustomer, "0", "0:Off,1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertPouchtoCrateonDispatch, clsFixedParameterCode.ConvertPouchtoCrateonDispatch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ChangeInventroyMovemnet, clsFixedParameterCode.ChangeInventroyMovemnet, "0", "0:Dispatch, 1:Gatepass;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickRatefromMaster, clsFixedParameterCode.PickRatefromMaster, "0", "0:Pick Rate Manually, 1:Pick Rate from Master;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyTPT, clsFixedParameterCode.ApplyTPT, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowMultipleUOMForProduct, clsFixedParameterCode.AllowMultipleUOMForProduct, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowManualCrateForDispatch, clsFixedParameterCode.AllowManualCrateForDispatch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyManualScheme, clsFixedParameterCode.ApplyManualScheme, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowResetCustomerDemandOnRouteChange, clsFixedParameterCode.AllowResetCustomerDemandOnRouteChange, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AutoSchemeOnTotalDispatchQty, clsFixedParameterCode.AutoSchemeOnTotalDispatchQty, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDefaultTCSIsChecked, clsFixedParameterCode.ApplyDefaultTCSIsChecked, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDefaultTPTIsChecked, clsFixedParameterCode.ApplyDefaultTPTIsChecked, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBoothRouteMapping, clsFixedParameterCode.ApplyBoothRouteMapping, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.PickDataFromRetestingTable, clsFixedParameterCode.PickDataFromRetestingTable, "0", "0:Pick Data From Milk Procurement Uploader History Table, 1:Pick Data From Retesting Table;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAllCustomerItemWiseSaleReportOptions, clsFixedParameterCode.ShowAllCustomerItemWiseSaleReportOptions, "0", "0:Off, 1:Show All Customer Item Wise Sale Report Options;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.APIMilkSample, clsFixedParameterCode.DrippingQuantity, "0", "Dripping Quantity;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyDepartmentRoute, clsFixedParameterCode.ApplyDepartmentRoute, "0", "Apply Department Route;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuantityTolerance, clsFixedParameterCode.QuantityTolerance, "0", "Quantity Tolerance;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AmountTolerance, clsFixedParameterCode.AmountTolerance, "0", "Amount Tolerance;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyItemCapacityLimit, clsFixedParameterCode.ApplyItemCapacityLimit, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SeprateMorningEveningSequence, clsFixedParameterCode.SeprateMorningEveningSequence, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FinancialImpactForTPT, clsFixedParameterCode.FinancialImpactForTPT, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FinancialImpactForDistributor, clsFixedParameterCode.FinancialImpactForDistributor, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.FinancialImpactForSecurity, clsFixedParameterCode.FinancialImpactForSecurity, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuickDemandUOMCrate, clsFixedParameterCode.QuickDemandUOMCrate, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuickDemandUOMPouch, clsFixedParameterCode.QuickDemandUOMPouch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.QuickDemandUOMLtr, clsFixedParameterCode.QuickDemandUOMLtr, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyBoothWiseScheme, clsFixedParameterCode.ApplyBoothWiseScheme, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ShowAndSaveCrystalReportActionType, clsFixedParameterCode.ShowAndSaveCrystalReportActionType, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.OfflineDCSMilkCollectionExeVersion, clsFixedParameterCode.OfflineDCSMilkCollectionExeVersion, "", "")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyMonthEndDispatch, clsFixedParameterCode.ApplyMonthEndDispatch, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyItemUOMOnDemand, clsFixedParameterCode.ApplyItemUOMOnDemand, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.MergeTCAmtofCreditCust, clsFixedParameterCode.MergeTCAmtofCreditCust, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertIntoBulkUOM, clsFixedParameterCode.ConvertIntoBulkUOM, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AddDistributorCommissioninItemPrice, clsFixedParameterCode.AddDistributorCommissioninItemPrice, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.SubstractDistributorCommissioninItemPrice, clsFixedParameterCode.SubstractDistributorCommissioninItemPrice, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ConvertTOBillingUOM, clsFixedParameterCode.ConvertTOBillingUOM, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DifferentCrateTypeForFGItem, clsFixedParameterCode.DifferentCrateTypeForFGItem, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.StopMultipleLoginInAPP, clsFixedParameterCode.StopMultipleLoginInAPP, "0", "0:Off, 1:On;")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DispatchCommissionDecimalPlaces, clsFixedParameterCode.DispatchCommissionDecimalPlaces, "4", "Enter Comission Decimal Places")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyPricePlanOnDocumentDate, clsFixedParameterCode.ApplyPricePlanOnDocumentDate, "0", "0:Supply Date, 1:Document Date")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DisplayTransporterAddress, clsFixedParameterCode.DisplayTransporterAddress, "0", "0:Company Address, 1:Transporter Address")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowRouteWiseDemandEntryInDecimal, clsFixedParameterCode.AllowRouteWiseDemandEntryInDecimal, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToCheckEntryUOMForProduct, clsFixedParameterCode.AllowToCheckEntryUOMForProduct, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToCheckEntryUOMForGNG, clsFixedParameterCode.AllowToCheckEntryUOMForGNG, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.AllowToCheckZeroQtyonDispatch, clsFixedParameterCode.AllowToCheckZeroQtyonDispatch, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultEnableEWayBill, clsFixedParameterCode.DefaultEnableEWayBill, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CreateAutoGatePass, clsFixedParameterCode.CreateAutoGatePass, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.ApplyEWBThresholdLimit, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraCity, "0", "Enter Threshold Limit For Intra-City")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraState, "0", "Enter Threshold Limit For Intra-State")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForInterState, "0", "Enter Threshold Limit For Inter-State")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DefaultEnableNoTransporter, clsFixedParameterCode.DefaultEnableNoTransporter, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ManualBatchOnCustomerBooking, clsFixedParameterCode.ManualBatchOnCustomerBooking, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.GenerateCustomerWiseGatePass, clsFixedParameterCode.GenerateCustomerWiseGatePass, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.DeductTPTFromDocAmt, clsFixedParameterCode.DeductTPTFromDocAmt, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.CheckCustomeroutStandingAmt, clsFixedParameterCode.CheckCustomeroutStandingAmt, "0", "0:Off, 1:On")
        InsertDefaultValueFixedParameter(clsFixedParameterType.ServerDateTimeForTaxableInvoice, clsFixedParameterCode.ServerDateTimeForTaxableInvoice, "0", "0:Off, 1:On")
        clsFixedParameterProgramMapping.SetDefaultValues()
        Return True
    End Function
    Public Shared Function InsertDefaultValueFixedParameter(ByVal strType As String, ByVal strCode As String, ByVal strDescription As String, ByVal strSpecification As String) As Boolean
        Dim qry As String = "select Type from TSPL_FIXED_PARAMETER where Code='" + strCode + "' and Type ='" + strType + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Type", strType)
        clsCommon.AddColumnsForChange(coll, "Code", strCode)
        clsCommon.AddColumnsForChange(coll, "Description", strDescription)
        clsCommon.AddColumnsForChange(coll, "Specification", strSpecification, True)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER", OMInsertOrUpdate.Insert, "")
        End If
        Return True
    End Function
End Class
Public Class clsFixedParameterProgramMapping
    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strType As String, ByVal strCode As String, ByVal ControlType As EnumControlType) As Boolean
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Program_Code", strProgramCode)
        clsCommon.AddColumnsForChange(coll, "FP_Type", strType)
        clsCommon.AddColumnsForChange(coll, "FP_Code", strCode)
        clsCommon.AddColumnsForChange(coll, "Control_Type", clsCommon.myCdbl(ControlType))
        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FIXED_PARAMETER_PROGRAM_MAPPING", OMInsertOrUpdate.Insert, "")
        Return True
    End Function
    Public Shared Sub SetDefaultValues()
        clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_FIXED_PARAMETER_PROGRAM_MAPPING")
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.SkipDebitNoteforRejectInBMCTruckSheet, clsFixedParameterCode.SkipDebitNoteforRejectInBMCTruckSheet, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.GatepassForTaxableandNonTaxableItems, clsFixedParameterCode.GatepassForTaxableandNonTaxableItems, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.PickDataFromMasterOrTransaction, clsFixedParameterCode.PickDataFromMasterOrTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.ExportTruckSheet, clsFixedParameterCode.ExportTruckSheet, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MaxFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.AndroidFarmerCollection, clsFixedParameterCode.MaxSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.ShowRateAndAmountOnFarmerCollection, clsFixedParameterCode.ShowRateAndAmountOnFarmerCollection, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.PickcompanyBankCodeFromPaymentProcess, clsFixedParameterCode.PickcompanyBankCodeFromPaymentProcess, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.isThereOnlyOneRowOfOwnDCS, clsFixedParameterCode.isThereOnlyOneRowOfOwnDCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.PrintOnlyPostedDocument, clsFixedParameterCode.PrintOnlyPostedDocument, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmMilkProcurementSocietySummary, clsFixedParameterType.MaxFATPerLimitforReport, clsFixedParameterCode.MaxFATPerLimitforReport, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmMilkProcurementSocietySummary, clsFixedParameterType.MaxSNFPerLimitforReport, clsFixedParameterCode.MaxSNFPerLimitforReport, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.PartialSRNAmount, clsFixedParameterCode.PartialSRNAmount, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmSendSMSToDCS, clsFixedParameterType.DefaultFilesPath, clsFixedParameterCode.DefaultFilesPath, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.rptGSTRReport, clsFixedParameterType.B2CDocumentAmountRangeOtherState, clsFixedParameterCode.B2CDocumentAmountRangeOtherState, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.rptGSTRReport, clsFixedParameterType.B2CDocumentAmountRange, clsFixedParameterCode.B2CDocumentAmountRange, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.NoOfDCSToLoadDeductionData, clsFixedParameterCode.NoOfDCSToLoadDeductionData, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.DCSWiseFilterEnableOnSavingCheck, clsFixedParameterCode.DCSWiseFilterEnableOnSavingCheck, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyBoothSecurity, clsFixedParameterCode.ApplyBoothSecurity, EnumControlType.CheckBox)

        InsertDefaultValue(clsUserMgtCode.frmSendBillToDCS, clsFixedParameterType.FileUpload, clsUserMgtCode.frmSendBillToDCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.FileUpload, clsUserMgtCode.frmDairyGatePass, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmUtilityForm, clsFixedParameterType.DeleteTempDataFromTime, clsFixedParameterCode.DeleteTempDataFromTime, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmUtilityForm, clsFixedParameterType.DeleteTempDataToTime, clsFixedParameterCode.DeleteTempDataToTime, EnumControlType.TextBox)

        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.PenaltyAfterDays, clsFixedParameterCode.PenaltyAfterDays, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.PenaltyCost, clsFixedParameterCode.PenaltyCost, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.BulkSaleFreightMaster, clsFixedParameterType.CalculateProRateAuto, clsFixedParameterCode.CalculateProRateAuto, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.ConvertVlcCodeUploaderToInt, clsFixedParameterCode.ConvertVlcCodeUploaderToInt, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.PickOnlyOWNBMCDCS, clsFixedParameterCode.PickOnlyOWNBMCDCS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.rptBoothTruckSheet, clsFixedParameterType.PrintPouchCrateQtyOnPrint, clsFixedParameterCode.PrintPouchCrateQtyOnPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.PickHeadLoadRateFromSecretaryMaster, clsFixedParameterCode.PickHeadLoadRateFromSecretaryMaster, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmRCDFRateControl, clsFixedParameterType.RCDFRateControl, clsFixedParameterCode.RCDFRateControl, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.RCDFControl, clsFixedParameterCode.MaxFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.RCDFControl, clsFixedParameterCode.MaxSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmEmployee_Master, clsFixedParameterType.EmployeePFRetirementAge, clsFixedParameterCode.EmployeePFRetirementAge, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmEmployee_Master, clsFixedParameterType.EmployeeRetirementAge, clsFixedParameterCode.EmployeeRetirementAge, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, clsFixedParameterType.PDAccountPaymanager, clsFixedParameterCode.NoOfRecordForLotNo, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBankAdvise, clsFixedParameterType.BankAdviseRequired, clsFixedParameterCode.BankAdviseRequired, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmBankAdvise, clsFixedParameterType.SeprateBankForDCSBankAdvice, clsFixedParameterCode.SeprateBankForDCSBankAdvice, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmBankAdvise, clsFixedParameterType.ApplyPartialBankAdvise, clsFixedParameterCode.ApplyPartialBankAdvise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.SeprateBankForDCSBankAdvice, clsFixedParameterCode.SeprateBankForDCSBankAdvice, EnumControlType.TextBox)

        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCSMultipleDays, clsFixedParameterType.ApplyMergeForDCSMultipleDays, clsFixedParameterCode.ApplyMergeForDCSMultipleDays, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.CalculateFATSNFLossCycleWise, clsFixedParameterCode.CalculateFATSNFLossCycleWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FATSNFDiffReport, clsFixedParameterType.CalculateFATSNFLossCycleWise, clsFixedParameterCode.CalculateFATSNFLossCycleWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DBTByDailyFarmerMilkCollection, clsFixedParameterCode.DBTByDailyFarmerMilkCollection, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DBTMilkQtyCapping, clsFixedParameterCode.DBTMilkQtyCapping, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.RefreshDBTReco, clsFixedParameterCode.RefreshDBTReco, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DistributorWiseBilling, clsFixedParameterCode.DistributorWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.MinimumQtyForHeadLoad, clsFixedParameterCode.MinimumQtyForHeadLoad, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmHeadLoadMaster, clsFixedParameterType.ApplyHeadLoadDayAndCycleWiseForAJM, clsFixedParameterCode.ApplyHeadLoadDayAndCycleWiseForAJM, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmNewDCSScreen, clsFixedParameterType.StopSetting, clsFixedParameterCode.JournalEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptTrialBalance, clsFixedParameterType.StopSetting, clsFixedParameterCode.JournalEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.stockRecoNew, clsFixedParameterType.StopSetting, clsFixedParameterCode.Inventory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.stockRecoNew, clsFixedParameterType.StopSetting, clsFixedParameterCode.InventoryNew, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessRouteReport, clsFixedParameterType.PickBulkRoute, clsFixedParameterCode.PickBulkRoute, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessRouteReport, clsFixedParameterType.ShowMultipleLegers, clsFixedParameterCode.ShowMultipleLegers, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessRouteReport, clsFixedParameterType.LoadLedgerMixedMilk, clsFixedParameterCode.LoadLedgerMixedMilk, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessRouteReport, clsFixedParameterType.ShowNewFormatofPDF, clsFixedParameterCode.ShowNewFormatofPDF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.HeadLoadRODecimalPlace, clsFixedParameterCode.HeadLoadRODecimalPlace, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.HeadLoadROIncreaseAfter, clsFixedParameterCode.HeadLoadROIncreaseAfter, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.ApplyUnpaidBank, clsFixedParameterCode.ApplyUnpaidBank, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.AllowZeroFATSNF, clsFixedParameterCode.AllowZeroFATSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.AdjustFATSNFINOwnVSP, clsFixedParameterCode.AdjustFATSNFINOwnVSP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.AdjustFATSNFINOwnVSP, clsFixedParameterCode.AdjustQtyINOwnVSP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.ApplySameDayShift, clsFixedParameterCode.ApplySameDayShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkPricePlanning, clsFixedParameterType.MandatoryPDFFileMilkPricePlan, clsFixedParameterCode.MandatoryPDFFileMilkPricePlan, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.AllowMPIncetiveQtyAboveBilledQty, clsFixedParameterCode.AllowMPIncetiveQtyAboveBilledQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.PickMilkPurchaseInvoiceQtyOrRecoQty, clsFixedParameterCode.PickMilkPurchaseInvoiceQtyOrRecoQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.HeaderFATSNFKGDecimalPlaces, clsFixedParameterCode.HeaderFATSNFKGDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.SNFDecimalPlaces, clsFixedParameterCode.SNFDecimalPlaces, EnumControlType.NumericBox)

        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.BMCDecimalPlaces, clsFixedParameterCode.BMCFATDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.BMCDecimalPlaces, clsFixedParameterCode.BMCSNFDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.BMCDCSApplyShift, clsFixedParameterCode.BMCDCSApplyShift, EnumControlType.CheckBox)

        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.HideShiftCollection, clsFixedParameterCode.HideShiftCollection, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MilkCollectionPickBulkRoute, clsFixedParameterCode.MilkCollectionPickBulkRoute, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptDailyQtyReport, clsFixedParameterType.DailyQtyReport, clsFixedParameterCode.FATKGSNFKGRoundOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmHeadLoadMaster, clsFixedParameterType.ShowOwnBMCDCS, clsFixedParameterCode.ShowOwnBMCDCS, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.OwnBMCApplicationFATRatio, clsFixedParameterCode.OwnBMCApplicationFATRatio, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.OwnBMCApplicationSNFRatio, clsFixedParameterCode.OwnBMCApplicationSNFRatio, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.OwnBMCCreateDRCRNote, clsFixedParameterCode.OwnBMCCreateDRCRNote, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.TDSOpeningBalanceConsiderCreditNote, clsFixedParameterCode.TDSOpeningBalanceConsiderCreditNote, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.FATSNFNoDecimalDCS, clsFixedParameterCode.FATSNFNoDecimalDCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.ShowAllDCS, clsFixedParameterCode.ShowAllDCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, clsFixedParameterType.MaxRowsExcelDBTNEFTUploader, clsFixedParameterCode.MaxRowsExcelDBTNEFTUploader, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.rptMccProcurementUploader, clsFixedParameterType.XpertAPI, clsFixedParameterCode.WeighingRoundSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.rptMccProcurementUploader, clsFixedParameterType.XpertAPI, clsFixedParameterCode.SetDCSAvgFATSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.MilkRateRoundOffType, clsFixedParameterCode.MilkRateRoundOffType, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, clsFixedParameterType.DBTRevisePayment, clsFixedParameterCode.DBTRevisePayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, clsFixedParameterType.CycleWiseDBT, clsFixedParameterCode.CycleWiseDBT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.ShowSampleNoOnBMC, clsFixedParameterCode.ShowSampleNoOnBMC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.ShowTempratureOnBMC, clsFixedParameterCode.ShowTempratureOnBMC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.PickMilkTransferInData, clsFixedParameterCode.PickMilkTransferInData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.FillRouteTankerNo, clsFixedParameterCode.FillRouteTankerNo, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.RepeatBMCSampleNo, clsFixedParameterCode.RepeatBMCSampleNo, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.ShowAllMCC, clsFixedParameterCode.ShowAllMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCSMultipleDaysMerge, clsFixedParameterType.ShowDCSDetMerge, clsFixedParameterCode.ShowDCSDetMerge, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.FATSNFNoDecimalMCC, clsFixedParameterCode.FATSNFNoDecimalMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.ApplyGaze, clsFixedParameterCode.ApplyGaze, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.MilkCollectionFATSNFType, clsFixedParameterCode.MilkCollectionFATSNFType, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.MilkCollectionFATSNFTypeHeader, clsFixedParameterCode.MilkCollectionFATSNFTypeHeader, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, clsFixedParameterType.ShowAllMCC, clsFixedParameterCode.ShowAllMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, clsFixedParameterType.FATSNFNoDecimalMCC, clsFixedParameterCode.FATSNFNoDecimalMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, clsFixedParameterType.ApplyGaze, clsFixedParameterCode.ApplyGaze, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, clsFixedParameterType.MilkCollectionFATSNFType, clsFixedParameterCode.MilkCollectionFATSNFType, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, clsFixedParameterType.MilkCollectionFATSNFTypeHeader, clsFixedParameterCode.MilkCollectionFATSNFTypeHeader, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMPMaster, clsFixedParameterType.IncentiveAccNoMandatoryInMPMaster, clsFixedParameterCode.IncentiveAccNoMandatoryInMPMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.AndroidAPP, clsFixedParameterCode.StopDBTBefore, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmMPMaster, clsFixedParameterType.JanAadharNoMandatory, clsFixedParameterCode.JanAadharNoMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.StopUpdateForWeigingMilkReceipt, clsFixedParameterCode.StopUpdateForWeigingMilkReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.CreatePOFromMultipleLocation, clsFixedParameterCode.CreatePOFromMultipleLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.BulkProcRunOneTypeGateEntry, clsFixedParameterCode.BulkProcRunOneTypeGateEntry, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.ApplyZoneInDBT, clsFixedParameterCode.ApplyZoneInDBT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.ApplyPashuAaharAndMineralMixture, clsFixedParameterCode.ApplyPashuAaharAndMineralMixture, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MandatoryDCSMPIncetiveReco, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MatchingUOM, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyGreaterThanDCSQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyLessThanDCSQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyEqualToDCSQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMPMaster, clsFixedParameterType.BankIFSCCodeValidateByService, clsFixedParameterCode.BankIFSCCodeValidateByService, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, clsFixedParameterType.MPIncentiveEntryCycleWiseButNEFTMonthly, clsFixedParameterCode.MPIncentiveEntryCycleWiseButNEFTMonthly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, clsFixedParameterType.DCSQtyDecimalPlaces, clsFixedParameterCode.DCSQtyDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.MPIncentiveEntryIncentiveRate, clsFixedParameterCode.MPIncentiveEntryIncentiveRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, clsFixedParameterType.MPIncentiveEntryMaxMilkLimit, clsFixedParameterCode.MPIncentiveEntryMaxMilkLimit, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.VLCProgressReport, clsFixedParameterType.TrendDiffValueForColor, clsFixedParameterCode.TrendDiffValueForColor, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseReturn, clsFixedParameterType.AllowPurReturnEvenIfPaymentDone, clsFixedParameterCode.AllowPurReturnEvenIfPaymentDone, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmMakePayment, clsFixedParameterType.TransporterCollection, clsFixedParameterCode.TransporterCollection, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.ProcurmentShiftUploaderNo, clsFixedParameterCode.ProcurmentShiftUploaderNo, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.BholeBabaPaymentProcessProPrintStartDate, clsFixedParameterCode.BholeBabaPaymentProcessProPrintStartDate, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.IncludeInceAndDedInFATSNFRate, clsFixedParameterCode.IncludeInceAndDedInFATSNFRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.FATPerShouldBeMultipleOf5, clsFixedParameterCode.FATPerShouldBeMultipleOf5, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowtoChangeFatANdSNFPerforHighClassVendorinGE, clsFixedParameterCode.AllowtoChangeFatANdSNFPerforHighClassVendorinGE, EnumControlType.CheckBox)
        ''Add Here to map screen and setting
        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.PickFATSNFKgFromInventory, clsFixedParameterCode.PickFATSNFKgFromInventory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.EnableTDSforServiceVendorSeparately, clsFixedParameterCode.EnableTDSforServiceVendorSeparately, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.VSPBillDocumentToBeAddedInMilkCost, clsFixedParameterCode.VSPBillDocumentToBeAddedInMilkCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.DCSAddDedRODecimalPlace, clsFixedParameterCode.DCSAddDedRODecimalPlace, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.DCSAddDedROIncreaseAfter, clsFixedParameterCode.DCSAddDedROIncreaseAfter, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.DCSAddDedROHeaderLevel, clsFixedParameterCode.DCSAddDedROHeaderLevel, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.IncentiveEntry, clsFixedParameterType.IncetiveEntryApplyArrear, clsFixedParameterCode.IncetiveEntryApplyArrear, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, clsFixedParameterType.AllowFatPerInanynumberofMultipesonBulkQC, clsFixedParameterCode.AllowFatPerInanynumberofMultipesonBulkQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanSaleUploader, clsFixedParameterType.CreateSeparateInvoiceforeachrowinCansale, clsFixedParameterCode.CreateSeparateInvoiceforeachrowinCansale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowCanInformationintoGridForTankerDispatch, clsFixedParameterCode.AllowCanInformationintoGridForTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.VSPDayWiseIncentiveAtSRN, clsFixedParameterCode.VSPDayWiseIncentiveAtSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.IncentiveEntry, clsFixedParameterType.IncentiveProcessPaymentStartDate, clsFixedParameterCode.IncentiveProcessPaymentStartDate, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.CreateJVOFRGPNRGPAndItsRetrun, clsFixedParameterCode.CreateJVOFRGPNRGPAndItsRetrun, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.FixedAssetAcquisitionEntryHitInventoryMovement, clsFixedParameterCode.FixedAssetAcquisitionEntryHitInventoryMovement, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.CalculateDeductionByStdQtyinBulkMilkSRN, clsFixedParameterCode.CalculateDeductionByStdQtyinBulkMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, clsFixedParameterType.CreateAPinvoiceofsalaryemployeewiseduringsalarygen, clsFixedParameterCode.CreateAPinvoiceofsalaryemployeewiseduringsalarygen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, clsFixedParameterType.ApplyTaxInBulkMilkPurchaseInvoice, clsFixedParameterCode.ApplyTaxInBulkMilkPurchaseInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.MilkSRNFATSNFDecimalPlaces, clsFixedParameterCode.MilkSRNFATSNFDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TankerDispatchIntermittentSingleGateIn, clsFixedParameterCode.TankerDispatchIntermittentSingleGateIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransferOther, clsFixedParameterType.JobWorkOutwardComsumeItemAccordingToBOM, clsFixedParameterCode.JobWorkOutwardComsumeItemAccordingToBOM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.JWOutWardForRCDF, clsFixedParameterCode.JWOutWardForRCDF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionRemoveFATSNFKgTollerance, clsFixedParameterCode.ProductionRemoveFATSNFKgTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionCheckFATKg, clsFixedParameterCode.ProductionCheckFATKg, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionCheckSNFKg, clsFixedParameterCode.ProductionCheckSNFKg, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ProductionOnlyOneIssueAgainstBatch, clsFixedParameterCode.ProductionOnlyOneIssueAgainstBatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.MilkShiftEndAutoAdjInItemCode, clsFixedParameterCode.MilkShiftEndAutoAdjInItemCode, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmVillageMaster, clsFixedParameterType.IsAdditionalInformationOnVillageMaster, clsFixedParameterCode.IsAdditionalInformationOnVillageMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyFreshDispatchMultiple, clsFixedParameterType.StockCheckOnPostForDairyDispatchMultiple, clsFixedParameterCode.StockCheckOnPostForDairyDispatchMultiple, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.locationMaster, clsFixedParameterType.SHowOptionOnLocationForDairyDispatchfromDOorGatepass, clsFixedParameterCode.SHowOptionOnLocationForDairyDispatchfromDOorGatepass, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkRouteMaster, clsFixedParameterType.VLCTimeTableColumnShow, clsFixedParameterCode.VLCTimeTableColumnShow, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkRouteMaster, clsFixedParameterType.ApplyEffectiveStartDate, clsFixedParameterCode.ApplyEffectiveStartDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkRouteMaster, clsFixedParameterType.VLCTimeTableColumnMandatory, clsFixedParameterCode.VLCTimeTableColumnMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterMaster, clsFixedParameterType.isOneMCCOnePrimaryTranporter, clsFixedParameterCode.isOneMCCOnePrimaryTranporter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.isIntimationRequired, clsFixedParameterCode.isIntimationRequired, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowBulkProcurementSequencewise, clsFixedParameterCode.AllowBulkProcurementSequencewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowBulkProcurementSequencewise, clsFixedParameterCode.AllowBulkProcurementSequencewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.QualityThenWeighmentinBulkProcurement, clsFixedParameterCode.QualityThenWeighmentinBulkProcurement, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.QualityThenWeighmentinBulkProcurement, clsFixedParameterCode.QualityThenWeighmentinBulkProcurement, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorPriceChartMapping, clsFixedParameterType.AllowBulkPriceChartMultiplepriceToMultipleVendor, clsFixedParameterCode.AllowBulkPriceChartMultiplepriceToMultipleVendor, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.OpenODDEvenForm, clsFixedParameterCode.OpenODDEvenForm, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.Open4AnalyzerForm, clsFixedParameterCode.Open4AnalyzerForm, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MilkSamplShowOddEvenTwoGrid, clsFixedParameterCode.MilkSamplShowOddEvenTwoGrid, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.CheckOutstandingCreditLimitOnBooking, clsFixedParameterCode.CheckOutstandingCreditLimitOnBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.ShowItemLocationWiseonDairyBooking, clsFixedParameterCode.ShowItemLocationWiseonDairyBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.SeprateDemandForMorningEveningShift, clsFixedParameterCode.SeprateDemandForMorningEveningShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPAssetIssue, clsFixedParameterType.IsApplyEMIOnAssetValue, clsFixedParameterCode.IsApplyEMIOnAssetValue, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.isCleaningMandatoryBeforeGateout, clsFixedParameterCode.isCleaningMandatoryBeforeGateout, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCleaning, clsFixedParameterType.ShowBothTankertypeOnCleaning, clsFixedParameterCode.ShowBothTankertypeOnCleaning, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.ShowUnloadingandWeighmentSequencewise, clsFixedParameterCode.ShowUnloadingandWeighmentSequencewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.ShowUnloadingandWeighmentSequencewise, clsFixedParameterCode.ShowUnloadingandWeighmentSequencewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmScrapSaleInvoice, clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.AllowFreshPriceChartOnBookingProductSale, clsFixedParameterCode.AllowFreshPriceChartOnBookingProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.GrossWtFromItemMasterONProductSale, clsFixedParameterCode.GrossWtFromItemMasterONProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AutoCalculateCrateOnDairyDispatch, clsFixedParameterCode.AutoCalculateCrateOnDairyDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ShowSchemeItemRateonDairyDispatch, clsFixedParameterCode.ShowSchemeItemRateonDairyDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ShowSchemeItemRateonDairyDispatchTaxable, clsFixedParameterCode.ShowSchemeItemRateonDairyDispatchTaxable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.CalculateCommOnCSATransWOConversion, clsFixedParameterCode.CalculateCommOnCSATransWOConversion, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AutoCalculateCANOnDairyDispatch, clsFixedParameterCode.AutoCalculateCANOnDairyDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.ConsiderSiloCapicityForStockIn, clsFixedParameterCode.ConsiderSiloCapicityForStockIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowBulkProcMCCwithoutTankerDispatch, clsFixedParameterCode.AllowBulkProcMCCwithoutTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowRandomOnlyOneSecondaryQC, clsFixedParameterCode.AllowRandomOnlyOneSecondaryQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowFractionInMCCTankerDispatchGrossQty, clsFixedParameterCode.AllowFractionInMCCTankerDispatchGrossQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowAdditionalWeightinPercentage, clsFixedParameterCode.AllowAdditionalWeightinPercentage, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.EnterAdditionalWeight, clsFixedParameterCode.EnterAdditionalWeight, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.FirstGateOutProcessForMCCBulkProcument, clsFixedParameterCode.FirstGateOutProcessForMCCBulkProcument, EnumControlType.CheckBox)
        ' == KUNAL > TICKET :  BM00000009575 ===
        'InsertDefaultValue(clsUserMgtCode.MilkTruckSheet, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShift, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRNReturn, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmDispatchTransfer, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCGateEntry, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MccMilkTransferPrice, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCWeighment, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmMCCMilkTransPortorInvoice, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVLCDataUploaderManual, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmVSPItemIssue, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        ' == KUNAL > TICKET : BM00000009580 ===
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.AllowPurchaseModulewithUniqueItem, clsFixedParameterCode.AllowPurchaseModulewithUniqueItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.AllowPurchaseModulewithUniqueItem, clsFixedParameterCode.AllowPurchaseModulewithUniqueItem, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AllowSRNWithoutShortageRejection, clsFixedParameterCode.AllowSRNWithoutShortageRejection, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseReturn, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RFQ, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SRNReturn, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.VendorQuotation, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowItemMasterPostedData, clsFixedParameterCode.AllowItemMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegisterWithCSASalePatti, clsFixedParameterType.EnableGSTRelatedfields, clsFixedParameterCode.EnableGSTRelatedfields, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowPostGSTPayment, clsFixedParameterCode.AllowPostGSTPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.ConsiderSiloCapicityForStockIn, clsFixedParameterCode.ConsiderSiloCapicityForStockIn, EnumControlType.CheckBox)
        'KUNAL > TICKET : BM00000009609  =======
        InsertDefaultValue(clsUserMgtCode.FrmBookingEntry, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmGateEntrySale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.TransferCrateReceived, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCreateReceived, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryNoteFreshSale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDispatchMultipleFreshSale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchFreshSale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmInvoiceFreshSale, clsFixedParameterType.AllowFutureDateTransaction, clsFixedParameterCode.AllowFutureDateTransaction, EnumControlType.CheckBox)
        '= PARTEEK ===
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, clsFixedParameterType.AllowCSAPriceMasterPostedData, clsFixedParameterCode.AllowCSAPriceMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.AllowCSAPriceMasterPostedData, clsFixedParameterCode.AllowCSAPriceMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.AllowCSAPriceMasterPostedData, clsFixedParameterCode.AllowCSAPriceMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartMaster, clsFixedParameterType.AllowMilkItemMasterPostedData, clsFixedParameterCode.AllowMilkItemMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, clsFixedParameterType.AllowMilkItemMasterPostedData, clsFixedParameterCode.AllowMilkItemMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, clsFixedParameterType.AutoGenerateFarmerPriceFromDCSPrice, clsFixedParameterCode.AutoGenerateFarmerPriceFromDCSPrice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.AllowBulkProcItemPostedData, clsFixedParameterCode.AllowBulkProcItemPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowBulkProcItemPostedData, clsFixedParameterCode.AllowBulkProcItemPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PriceMaster, clsFixedParameterType.AllowPriceListMasterPostedData, clsFixedParameterCode.AllowPriceListMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBookingEntry, clsFixedParameterType.AllowPriceListMasterPostedData, clsFixedParameterCode.AllowPriceListMasterPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDispatchMultipleFreshSale, clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDispatchMultipleFreshSale, clsFixedParameterType.AllowFreshInvoiceAutoPost, clsFixedParameterCode.AllowFreshInvoiceAutoPost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.SeparateDairyDispatchTaxableNonTaxable, clsFixedParameterCode.SeparateDairyDispatchTaxableNonTaxable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowGateEntryAgainstPO, clsFixedParameterCode.AllowGateEntryAgainstPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowRandomOnlyOneSecondaryQC, clsFixedParameterCode.AllowRandomOnlyOneSecondaryQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.lockTransaction, clsFixedParameterType.DaysToStartAutoLock, clsFixedParameterCode.DaysToStartAutoLock, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.RunBulkProcWithoutMilkGrade, clsFixedParameterCode.RunBulkProcWithoutMilkGrade, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, clsFixedParameterType.AllowManualPriceONBulkPO, clsFixedParameterCode.AllowManualPriceONBulkPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateEntryInPrevDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmIntimation, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmIntimation, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateEntryInPrevDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnVCGLEntry, clsFixedParameterType.AllowTransferVSPAmtToFarmerinVCGL, clsFixedParameterCode.AllowTransferVSPAmtToFarmerinVCGL, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowManualPriceONBulkPO, clsFixedParameterCode.AllowManualPriceONBulkPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateEntryInPrevDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowBulkProcMCCwithoutTankerDispatch, clsFixedParameterCode.AllowBulkProcMCCwithoutTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.CheckParameterRangerProcurementTypewise, clsFixedParameterCode.CheckParameterRangerProcurementTypewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmParameterRangeMaster, clsFixedParameterType.CheckParameterRangerProcurementTypewise, clsFixedParameterCode.CheckParameterRangerProcurementTypewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.PickCorrectionFactorProcurementTypewise, clsFixedParameterCode.PickCorrectionFactorProcurementTypewise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowAutoBulkMilkSRNonWeighmentBulkProc, clsFixedParameterCode.AllowAutoBulkMilkSRNonWeighmentBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.ShowGateEntryTypeonGateEntryBulkProc, clsFixedParameterCode.ShowGateEntryTypeonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowAutoBulkMilkSRNonWeighmentBulkProc, clsFixedParameterCode.AllowAutoBulkMilkSRNonWeighmentBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPendingApproval1, clsFixedParameterType.ShowDairySaleModuleOnBulkPosting, clsFixedParameterCode.ShowDairySaleModuleOnBulkPosting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.CreateProvisionJournalEntryForSale, clsFixedParameterCode.CreateProvisionJournalEntryForSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.CreateProvisionJournalEntryForTankerDispatch, clsFixedParameterCode.CreateProvisionJournalEntryForTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AllowProvisionknokoffOnAPInvoice, clsFixedParameterCode.AllowProvisionknokoffOnAPInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.AllowDifferentStateofChildCustomerOnPS, clsFixedParameterCode.AllowDifferentStateofChildCustomerOnPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.AllowVSPMasterAutoPrefix, clsFixedParameterCode.AllowVSPMasterAutoPrefix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.DoNotConsiderCustomerCreditLimit, clsFixedParameterCode.DoNotConsiderCustomerCreditLimit, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.AllowUniqueNoOnMilkTransferInandTankDis, clsFixedParameterCode.AllowUniqueNoOnMilkTransferInandTankDis, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowUniqueNoOnMilkTransferInandTankDis, clsFixedParameterCode.AllowUniqueNoOnMilkTransferInandTankDis, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.CreateSeperateSeriesforRefDocAPinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocAPinvforCreditdebit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.CreateSeperateSeriesforRefDocARinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocARinvforCreditdebit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.CreateSeperateTaxInvForFOCIteminNonTaxdispatch, clsFixedParameterCode.CreateSeperateTaxInvForFOCIteminNonTaxdispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.CreateCommonDairyDispatchforFreshAmbient, clsFixedParameterCode.CreateCommonDairyDispatchforFreshAmbient, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCreateReceived, clsFixedParameterType.CreateCommonDairyDispatchforFreshAmbient, clsFixedParameterCode.CreateCommonDairyDispatchforFreshAmbient, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.EnableCustomerPODetailonDairyBooking, clsFixedParameterCode.EnableCustomerPODetailonDairyBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.ChangeFATCLRafterspecialApprovalonQC, clsFixedParameterCode.ChangeFATCLRafterspecialApprovalonQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipToLocationDetails, clsFixedParameterType.EnableAutoDocNoShipToLocation, clsFixedParameterCode.EnableAutoDocNoShipToLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipToLocationDetails, clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCustomerCategoryLevel, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCustomerCategoryLevel, clsFixedParameterType.ShowBinMapping, clsFixedParameterCode.ShowBinMapping, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCustomerCategoryStructure, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmReceivablePaymentTerms, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmReceivablePaymentTerms, clsFixedParameterType.AllowToShowSaleTypeinPaymentTermsReceivable, clsFixedParameterCode.AllowToShowSaleTypeinPaymentTermsReceivable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer, clsFixedParameterType.AllowSiloMilkTransfertoMainLocation, clsFixedParameterCode.AllowSiloMilkTransfertoMainLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer, clsFixedParameterType.ConsiderSiloCapicityForStockIn, clsFixedParameterCode.ConsiderSiloCapicityForStockIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.EnableManualCrateonTaxableDairyDispatch, clsFixedParameterCode.EnableManualCrateonTaxableDairyDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowPriceMappingOnBulkSRNinChamberBulkProc, clsFixedParameterCode.AllowPriceMappingOnBulkSRNinChamberBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.GateEntryChamberwisewithManualTankerEntry, clsFixedParameterCode.GateEntryChamberwisewithManualTankerEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowPriceMappingOnBulkSRNinChamberBulkProc, clsFixedParameterCode.AllowPriceMappingOnBulkSRNinChamberBulkProc, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.CreateBulkMilkSRNItemwise, clsFixedParameterCode.CreateBulkMilkSRNItemwise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.ForceToSelectIteminGateEntry, clsFixedParameterCode.ForceToSelectIteminGateEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.CreateProvisionforBulkContractorInGateIn, clsFixedParameterCode.CreateProvisionforBulkContractorInGateIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.ApplyBankChargesasperSlabonBankMaster, clsFixedParameterCode.ApplyBankChargesasperSlabonBankMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptMPWiseMilkCollectionATPoolingPoint, clsFixedParameterType.DisplayAverageFatSNFMPWise, clsFixedParameterCode.DisplayAverageFatSNFMPWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.UseKGLitreConversionInBulkSaleAsperCLRCalculation, clsFixedParameterCode.UseKGLitreConversionInBulkSaleAsperCLRCalculation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShowLastUnitCostZeroForNonInventoryItemOnPO, clsFixedParameterCode.ShowLastUnitCostZeroForNonInventoryItemOnPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.SentschemecogsinAnotherAccount, clsFixedParameterCode.SentschemecogsinAnotherAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.OPkmMandatoryonDS, clsFixedParameterCode.OPkmMandatoryonDS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderDairy, clsFixedParameterType.ManualBatchNoMandatoryOnBatchOrderScreen, clsFixedParameterCode.ManualBatchNoMandatoryOnBatchOrderScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderDairy, clsFixedParameterType.ManualySelectBOMForChildBatch, clsFixedParameterCode.ManualySelectBOMForChildBatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderDairy, clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderDairy, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmStanderdProductionEntry, clsFixedParameterType.ConsumptionQtyTollerance, clsFixedParameterCode.ConsumptionQtyTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmStanderdProductionEntry, clsFixedParameterType.GunnyBagTollerance, clsFixedParameterCode.GunnyBagTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmStanderdProductionEntry, clsFixedParameterType.CheckBalanceOnSave, clsFixedParameterCode.RCDFCFProductionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.CheckBalanceOnSave, clsFixedParameterCode.RCDFCFDispatch, EnumControlType.CheckBox)
        ''stuti
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemCrateWtinKg, clsFixedParameterCode.ItemCrateWtinKg, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemJaaliWtinKg, clsFixedParameterCode.ItemJaaliWtinKg, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemBoxWtinKg, clsFixedParameterCode.ItemBoxWtinKg, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemCrateRate, clsFixedParameterCode.ItemCrateRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanSale, clsFixedParameterType.ItemCanRate, clsFixedParameterCode.ItemCanRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemJaaliRate, clsFixedParameterCode.ItemJaaliRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.itemMaster, clsFixedParameterType.ItemBoxRate, clsFixedParameterCode.ItemBoxRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.ShowHierarchyAndCostCenterInARInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInARInvoiceEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.DOTaggingForDairySaleModule, clsFixedParameterCode.DOTaggingForDairySaleModule, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowtoUnlockTransactionsforSetOff, clsFixedParameterCode.AllowtoUnlockTransactionsforSetOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowUseApplyDocSeriesForReceipt, clsFixedParameterCode.AllowUseApplyDocSeriesForReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShowCancelButtonPO, clsFixedParameterCode.ShowCancelButtonPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.AutoClosePO, clsFixedParameterCode.AutoClosePO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.ApplySendApprovalSetting, clsFixedParameterCode.ApplySendApprovalSetting, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.CreateJVForAllCasesinRGP, clsFixedParameterCode.CreateJVForAllCasesinRGP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.StoreRequisitionMandatoryforstorerequest, clsFixedParameterCode.StoreRequisitionMandatoryforstorerequest, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vhicleMaster, clsFixedParameterType.MandatoryEmployeeOnVehicleMaster, clsFixedParameterCode.MandatoryEmployeeOnVehicleMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vhicleMaster, clsFixedParameterType.MTCapacityRequired, clsFixedParameterCode.MTCapacityRequired, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.MTCapacityRequired, clsFixedParameterCode.MTCapacityRequired, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.locationMaster, clsFixedParameterType.PlantDepotMappingMandatory, clsFixedParameterCode.PlantDepotMappingMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.AllowThreeFormatByDefaultForPrint, clsFixedParameterCode.AllowThreeFormatByDefaultForPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.AllowOneFormatByDefaultForPrint, clsFixedParameterCode.AllowOneFormatByDefaultForPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.JobWorkTransfer, clsFixedParameterCode.JobWorkTransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoicedairy, clsFixedParameterType.AllowThreeFormatByDefaultForPrint, clsFixedParameterCode.AllowThreeFormatByDefaultForPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.AllowBackDateEntry, clsFixedParameterCode.AllowBackDateEntry, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.DipMarkingMendatory, clsFixedParameterCode.DipMarkingMendatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegister, clsFixedParameterType.AmountInLacsOnMisSaleRegister, clsFixedParameterCode.AmountInLacsOnMisSaleRegister, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegisterWithCSASalePatti, clsFixedParameterType.AmountInLacsOnMisSaleRegister, clsFixedParameterCode.AmountInLacsOnMisSaleRegister, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShortCloseItemWiseOnPO, clsFixedParameterCode.ShortCloseItemWiseOnPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.MakeClosingofPOreadonlyforuser, clsFixedParameterCode.MakeClosingofPOreadonlyforuser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.MakeClosingofCustomerTenderreadonlyforuser, clsFixedParameterCode.MakeClosingofCustomerTenderreadonlyforuser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.MakeClosingofCustomerTenderOrderreadonlyforuser, clsFixedParameterCode.MakeClosingofCustomerTenderOrderreadonlyforuser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmApprovalLevelScreen, clsFixedParameterType.AllowModificationOnApprovalByApprovalUser, clsFixedParameterCode.AllowModificationOnApprovalByApprovalUser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.AllowAutoCalculateADDREMOVEQty, clsFixedParameterCode.AllowAutoCalculateADDREMOVEQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegister, clsFixedParameterType.Donotshowtrasnfertransactionsbydefault, clsFixedParameterCode.Donotshowtrasnfertransactionsbydefault, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptSaleRegisterDetail, clsFixedParameterType.Donotshowtrasnfertransactionsbydefault, clsFixedParameterCode.Donotshowtrasnfertransactionsbydefault, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.AllowManualItemPriceOnMCCSale, clsFixedParameterCode.AllowManualItemPriceOnMCCSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.ExcludeKKFAndMandiForDCS, clsFixedParameterCode.ExcludeKKFAndMandiForDCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.ExcludeKKFAndMandiForDCS, clsFixedParameterCode.ExcludeKKFAndMandiForDCS, EnumControlType.CheckBox)
        'KUNAL > DATE : 11-JAN-2016 > MPD > REQ No. : MPDREQ000017 
        InsertDefaultValue(clsUserMgtCode.FrmDispatchFreshSale, clsFixedParameterType.AllowDispatchChecklistOnProductDispatch, clsFixedParameterCode.AllowDispatchChecklistOnProductDispatch, EnumControlType.CheckBox)
        'STUTI > MPD 
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowDispatchChecklistOnProductDispatch, clsFixedParameterCode.AllowDispatchChecklistOnProductDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.ShowIndentBasedOnCreatedUser, clsFixedParameterCode.ShowIndentBasedOnCreatedUser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShift, clsFixedParameterType.ShowSystemStockinOpenMCC, clsFixedParameterCode.ShowSystemStockinOpenMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmGateEntrySale, clsFixedParameterType.Tankerfromtankersalemasteringateentry, clsFixedParameterCode.Tankerfromtankersalemasteringateentry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.InsuranceNoAndSealNoInBulkDispatch, clsFixedParameterCode.InsuranceNoAndSealNoInBulkDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmJobMilkSRN, clsFixedParameterType.ValidateFatSNFOnJobMilkSRN, clsFixedParameterCode.ValidateFatSNFOnJobMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SRNReturn, clsFixedParameterType.CancelDocDueToSRNReturn, clsFixedParameterCode.CancelDocDueToSRNReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.lockTransaction, clsFixedParameterType.AllowAutoLockTransaction, clsFixedParameterCode.AllowAutoLockTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.lockTransaction, clsFixedParameterType.AllowLockTransactionUserwise, clsFixedParameterCode.AllowLockTransactionUserwise, EnumControlType.CheckBox)
        '--------------------end here---------------
        ''richa 
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCreateAutoInvoiceBS, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSalePriceChart, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.PickRateFromPRICEChrtMasterFORUMang, clsFixedParameterCode.PickRateFromPRICEChrtMasterFORUMang, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.StockTranferFromTransferPriceAndInvJVWithAvgCost, clsFixedParameterCode.StockTranferFromTransferPriceAndInvJVWithAvgCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.ApplyDocumentDate, clsFixedParameterCode.ApplyDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanSale, clsFixedParameterType.WeightOfCanForCanSale, clsFixedParameterCode.WeightOfCanForCanSale, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.AllowStockCheckatDOLevel, clsFixedParameterCode.AllowStockCheckatDOLevel, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowTankerBasedonVendorofGE, clsFixedParameterCode.AllowTankerBasedonVendorofGE, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmParameterMaster, clsFixedParameterType.AllowUseBoilingParameteronParameterMaster, clsFixedParameterCode.AllowUseBoilingParameteronParameterMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowUseBoilingParameteronParameterMaster, clsFixedParameterCode.AllowUseBoilingParameteronParameterMaster, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.SecondarySettingForQC, clsFixedParameterType.AllowAdditionalWeightinPercentage, clsFixedParameterCode.AllowAdditionalWeightinPercentage, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.SecondarySettingForQC, clsFixedParameterType.EnterAdditionalWeight, clsFixedParameterCode.EnterAdditionalWeight, EnumControlType.TextBox)
        'InsertDefaultValue(clsUserMgtCode.SecondarySettingForQC, clsFixedParameterType.AllowRandomOnlyOneSecondaryQC, clsFixedParameterCode.AllowRandomOnlyOneSecondaryQC, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATDeductionPercent, clsFixedParameterCode.FATDeductionPercent, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.SNFDeductionPercent, clsFixedParameterCode.SNFDeductionPercent, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.GraceTimeForTransporter, clsFixedParameterCode.GraceTimeForTransporter, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.RejectionReturnPenaltyPerUnit, clsFixedParameterCode.RejectionReturnPaneltyPerUnit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.RejectionDrainPenaltyPerUnit, clsFixedParameterCode.RejectionDrainPenaltyPerUnit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.RejectionCOBPenaltyPerUnit, clsFixedParameterCode.RejectionCOBPenaltyPerUnit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.GraceTimeFromGateEntryToDocWeighing, clsFixedParameterCode.GraceTimeFromGateEntryToDocWeighing, EnumControlType.NumericBox)
        'KUNAL > DATE : 23-01-2017 > CLIENT : Sahayog Dairy
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.ShowAliasNames, clsFixedParameterType.ShowAliasNames, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryIn, clsFixedParameterType.ShowFatAndSnfPercentageFields, clsFixedParameterType.ShowFatAndSnfPercentageFields, EnumControlType.CheckBox)
        ''======================CSA Sale Setting================================================================
        InsertDefaultValue(clsUserMgtCode.frmCSASalePattiReturn, clsFixedParameterType.ShowCSAReturnTypeOnScreen, clsFixedParameterCode.ShowCSAReturnTypeOnScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransferReturn, clsFixedParameterType.ShowCSAReturnTypeOnScreen, clsFixedParameterCode.ShowCSAReturnTypeOnScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSABooking, clsFixedParameterType.ShowCSARequestScreen, clsFixedParameterCode.ShowCSARequestScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSARequest, clsFixedParameterType.ShowCSARequestScreen, clsFixedParameterCode.ShowCSARequestScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.ShowCSARequestScreen, clsFixedParameterCode.ShowCSARequestScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.AllowSchemeOnCSADeliveryOrder, clsFixedParameterCode.AllowSchemeOnCSADeliveryOrder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, clsFixedParameterType.AllowOtherItemOnCSAPriceMaster, clsFixedParameterCode.AllowOtherItemOnCSAPriceMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.AllowOtherItemOnCSAPriceMaster, clsFixedParameterCode.AllowOtherItemOnCSAPriceMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.AllowOtherItemOnCSAPriceMaster, clsFixedParameterCode.AllowOtherItemOnCSAPriceMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.AllowOtherItemOnCSAPriceMaster, clsFixedParameterCode.AllowOtherItemOnCSAPriceMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.FreightChargeOnCSASaleInvoice, clsFixedParameterCode.FreightChargeOnCSASaleInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.AllowDisabledCommissionOnCSATransfer, clsFixedParameterCode.AllowDisabledCommissionOnCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSARequest, clsFixedParameterType.DoReadonly_UnitRate_AtCSASale, clsFixedParameterCode.DoReadonly_UnitRate_AtCSASale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.DoReadonly_UnitRate_AtCSASale, clsFixedParameterCode.DoReadonly_UnitRate_AtCSASale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.DoReadonly_UnitRate_AtCSASale, clsFixedParameterCode.DoReadonly_UnitRate_AtCSASale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.DoReadonly_UnitRate_AtCSASale, clsFixedParameterCode.DoReadonly_UnitRate_AtCSASale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.Allow_SaleMfgACONCSAPatti, clsFixedParameterCode.Allow_SaleMfgACONCSAPatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSchemeMasterDairy, clsFixedParameterType.AllowSchemeItemCondONSchemeMaster, clsFixedParameterCode.AllowSchemeItemCondONSchemeMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.CheckCreditLimitonCSADO, clsFixedParameterCode.CheckCreditLimitonCSADO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.EnableExciseONCSASalePatti, clsFixedParameterCode.EnableExciseONCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, clsFixedParameterType.EnableExciseONCSASalePatti, clsFixedParameterCode.EnableExciseONCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, clsFixedParameterType.ForUDLOnly, clsFixedParameterCode.ForUDLOnly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.ForUDLOnly, clsFixedParameterCode.ForUDLOnly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.ForUDLOnly, clsFixedParameterCode.ForUDLOnly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.ForUDLOnly, clsFixedParameterCode.ForUDLOnly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.BatchSkipCSAReturn, clsFixedParameterCode.BatchSkipCSAReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASalePattiReturn, clsFixedParameterType.BatchSkipCSAReturn, clsFixedParameterCode.BatchSkipCSAReturn, EnumControlType.CheckBox)
        ''=================end here-==========================================================
        InsertDefaultValue(clsUserMgtCode.frmTankerMaster, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AllowLoginTypeCNFdistributerRetailer, clsFixedParameterCode.AllowLoginTypeCNFdistributerRetailer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptMatrixFreshSalesReportSaleDairy, clsFixedParameterType.AllowSchemeItemQty, clsFixedParameterCode.AllowSchemeItemQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryOrderDairy, clsFixedParameterType.AllowDairyDeliveryOrderPrint, clsFixedParameterCode.AllowDairyDeliveryOrderPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmTankerOut, clsFixedParameterType.ShowSealNumberForTunkerOut, clsFixedParameterCode.ShowSealNumberForTunkerOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.HideRateDispatchCentreCode, clsFixedParameterCode.HideRateDispatchCentreCode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.AllowPanNoValidation, clsFixedParameterCode.AllowPanNoValidation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.AutoGeneratedVendorCodeForAllCompany, clsFixedParameterCode.AutoGeneratedVendorCodeForAllCompany, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.AutoGeneratedVendorCode, clsFixedParameterCode.AutoGeneratedVendorCode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.INDUSTRYTYPE, clsFixedParameterCode.INDUSTRYTYPE, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorCategoryLevel, clsFixedParameterType.ShowBinMapping, clsFixedParameterCode.ShowBinMapping, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorCategoryLevel, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorCategoryStructure, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.paymentTerms, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.transportMasterVendor, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmHirerachyLevelMaster, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmHirerachyLevelMaster, clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, EnumControlType.CheckBox)
        'Pravet
        'InsertDefaultValue(clsUserMgtCode.EmployeeMaster, clsFixedParameterType.AddParavetEmployeeType, clsFixedParameterCode.AddParavetEmployeeType, EnumControlType.CheckBox)
        ' =================
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        ' =================
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.AllowVehicleGateOutValidationScrapSale, clsFixedParameterCode.AllowVehicleGateOutValidationScrapSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.AllowVehicleGateOutValidationCSATransfer, clsFixedParameterCode.AllowVehicleGateOutValidationCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowVehicleGateOutValidationSPSale, clsFixedParameterCode.AllowVehicleGateOutValidationSPSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.AllowVehicleGateOutValidationTransfer, clsFixedParameterCode.AllowVehicleGateOutValidationTransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.AllowWithoutUnitCostIssueReturnEntry, clsFixedParameterCode.AllowWithoutUnitCostIssueReturnEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.ZeroCostForReprocess, clsFixedParameterCode.ZeroCostForReprocess, EnumControlType.CheckBox)
        '====================Added by preeti gupta=====================
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.IsAutoReceiptPayment, clsFixedParameterCode.IsAutoReceiptPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TransferEntryOnInvCtrlAccount, clsFixedParameterCode.TransferEntryOnInvCtrlAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTankerDispatchGL, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVLCMaster, clsFixedParameterType.AutoUpdateVLCUploaderCodeInVLCMaster, clsFixedParameterCode.AutoUpdateVLCUploaderCodeInVLCMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.StandardInterfaceForMilkShiftEnd, clsFixedParameterCode.StandardInterfaceForMilkShiftEnd, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.ShiftEndAllowManualEntryOfDeduction, clsFixedParameterCode.ShiftEndAllowManualEntryOfDeduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterVehicalMaster, clsFixedParameterType.PTMRatePerLtrKGOnStdQty, clsFixedParameterCode.PTMRatePerLtrKGOnStdQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedger, clsFixedParameterType.AllowTransactionFiltersOnCustomerlegder, clsFixedParameterCode.AllowTransactionFiltersOnCustomerlegder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedger, clsFixedParameterType.GroupCustomerlegderZoneWiseAreaWise, clsFixedParameterCode.GroupCustomerlegderZoneWiseAreaWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedger, clsFixedParameterType.AllowtoSHOWParentChildCustomer, clsFixedParameterCode.AllowtoSHOWParentChildCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedger, clsFixedParameterType.AllowtoMakeApplyDocOnbyDefault, clsFixedParameterCode.AllowtoMakeApplyDocOnbyDefault, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.VendorLedgerReport, clsFixedParameterType.AllowtoMakeApplyDocOnbyDefault, clsFixedParameterCode.AllowtoMakeApplyDocOnbyDefault, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomerOutstanding, clsFixedParameterType.PenaltyPercentage, clsFixedParameterCode.PenaltyPercentage, EnumControlType.NumericBox)
        'KUNAL > DATE : 24-01-2017 > CLIENT : Sahayog Dairy > ASSIGNED VIA : EMAIL > REQUEST NO : SCMPLREQ000002
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterVehicalMaster, clsFixedParameterType.VehicleFitnessAndInsuranceFields, clsFixedParameterType.VehicleFitnessAndInsuranceFields, EnumControlType.CheckBox)
        ' for mobile app
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.DefaultBank, clsFixedParameterCode.DefaultBank, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.DefaultLocation, clsFixedParameterCode.DefaultLocation, EnumControlType.TextBox)
        '=============added by preeti gupta============03/10/2016
        InsertDefaultValue(clsUserMgtCode.rptsaleRegisterReport, clsFixedParameterType.ShowParticluarColumnInSalesRegisterForGopalJee, clsFixedParameterCode.ShowParticluarColumnInSalesRegisterForGopalJee, EnumControlType.CheckBox)
        'Added by Nazia
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ShowPrintDiscountInDairyDispatchForGopaljee, clsFixedParameterCode.ShowPrintDiscountInDairyDispatchForGopaljee, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.MilkReceiptRequiredApproval, clsFixedParameterCode.MilkReceiptRequiredApproval, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.LinkDepartmentBetweenIndentAndIssue, clsFixedParameterCode.LinkDepartmentBetweenIndentAndIssue, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSchemeMasterNew, clsFixedParameterType.CombineExportImportOnSchemeMaster, clsFixedParameterCode.CombineExportImportOnSchemeMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.OpenPOforRejectShortageQty, clsFixedParameterCode.OpenPOforRejectShortageQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, clsFixedParameterType.AutoSelectMCCRouteVLC, clsFixedParameterCode.AutoSelectMCCRouteVLC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryIn, clsFixedParameterType.PickServerDateWithNoChange, clsFixedParameterCode.PickServerDateWithNoChange, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryOut, clsFixedParameterType.PickServerDateWithNoChange, clsFixedParameterCode.PickServerDateWithNoChange, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryWeightment, clsFixedParameterType.PickServerDateWithNoChange, clsFixedParameterCode.PickServerDateWithNoChange, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.PickFinishedItemasBatchItem, clsFixedParameterCode.PickFinishedItemasBatchItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.ToleranceFixFor_RM_OT_TRADE, clsFixedParameterCode.ToleranceFixFor_RM_OT_TRADE, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.ConsiderAdvancePayment, clsFixedParameterCode.ConsiderAdvancePayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.PayableAmountZeroForMCCSale, clsFixedParameterCode.PayableAmountZeroForMCCSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.Allow_AmountTruncate_BulkMilkSRN, clsFixedParameterCode.Allow_AmountTruncate_BulkMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.Allow_AmountTruncate_BulkMilkSRN, clsFixedParameterCode.Allow_AmountTruncate_BulkMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.AutoPurchaseReturnFromIssueReturn, clsFixedParameterCode.AutoPurchaseReturnFromIssueReturn, EnumControlType.CheckBox)
        '=====Sanjeet=======
        InsertDefaultValue(clsUserMgtCode.FrmGatePassFS, clsFixedParameterType.ShowAlternateVechileforFreshSale, clsFixedParameterCode.ShowAlternateVechileforFreshSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.CreateProvisionOfTransporterInDairyDispatch, clsFixedParameterCode.CreateProvisionOfTransporterInDairyDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmOverheadCostMaster, clsFixedParameterType.IncludeRatePerHoursIn, clsFixedParameterCode.IncludeRatePerHoursIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.SinglePrintCopyDairyInvoice, clsFixedParameterCode.SinglePrintCopyDairyInvoice, EnumControlType.CheckBox)
        '===Sanjeet(UDL)-17/11/2016====
        InsertDefaultValue(clsUserMgtCode.TaxTracking, clsFixedParameterType.ShowVatSeriesNoSeprately, clsFixedParameterCode.ShowVatSeriesNoSeprately, EnumControlType.CheckBox)
        '===Sanjeet(UDL)-21/12/2016====
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploader, clsFixedParameterType.AllowToGenerate_NEFTUPLOADER, clsFixedParameterCode.AllowToGenerate_NEFTUPLOADER, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploader, clsFixedParameterType.DebitBankSelectWithNewFormateInNFTUploader, clsFixedParameterCode.DebitBankSelectWithNewFormateInNFTUploader, EnumControlType.CheckBox)
        '===Sanjeet(UDL)-05/01/2017====
        InsertDefaultValue(clsUserMgtCode.mbtnPendingApproval1, clsFixedParameterType.AllowBulkPostingofAllDocuments, clsFixedParameterCode.AllowBulkPostingofAllDocuments, EnumControlType.CheckBox)
        '===Sanjeet(UDL)-10/01/2017====
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AllowSameaAdditionalChargesMultiTime, clsFixedParameterCode.AllowSameaAdditionalChargesMultiTime, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCompanyMaster, clsFixedParameterType.EnableCostingMethod, clsFixedParameterCode.EnableCostingMethod, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.CalculateItemCostonAvgForAssembly, clsFixedParameterCode.CalculateItemCostonAvgForAssembly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.ShowAllCustomerOnMccMaterialSale, clsFixedParameterCode.ShowAllCustomerOnMccMaterialSale, EnumControlType.CheckBox)
        '================================
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.GenerateSecondryCode, clsFixedParameterCode.GenerateSecondryCode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.BulkProcNetWeightCalculationWithVendorWeight, clsFixedParameterCode.BulkProcNetWeightCalculationWithVendorWeight, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowManualRejectionOfTanker, clsFixedParameterCode.AllowManualRejectionOfTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.RunBulkProcOnAdjustedFATCLR, clsFixedParameterCode.RunBulkProcOnAdjustedFATCLR, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.BulkProcPriceChartStandardRateWithZero, clsFixedParameterCode.BulkProcPriceChartStandardRateWithZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.RemoveForceAapprovalofBulkSRN, clsFixedParameterCode.RemoveForceAapprovalofBulkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSalePriceChart, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSalePriceChart, clsFixedParameterType.ApplySaleRouteInDCSSale, clsFixedParameterCode.ApplySaleRouteInDCSSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.ValidateCustomerPANwithName, clsFixedParameterCode.ValidateCustomerPANwithName, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.AutoGeneratedCustomerCode, clsFixedParameterCode.AutoGeneratedCustomerCode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.DisplayFranchiseeinCustomer, clsFixedParameterCode.DisplayFranchiseeinCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.GSTActiveTaxesRatesGroup, clsFixedParameterCode.GSTActiveTaxesRatesGroup, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.EnableDistributorSubsidy, clsFixedParameterCode.EnableDistributorSubsidy, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.AutoGeneratedCustomerCodeForAllCompany, clsFixedParameterCode.AutoGeneratedCustomerCodeForAllCompany, EnumControlType.CheckBox)
        '=================for show all structure wise item 
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.ValidateTaxGroupForTransaction, clsFixedParameterCode.ValidateTaxGroupForTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoicedairy, clsFixedParameterType.AllowSeprateSchemeItemPrintDairySaleInvoice, clsFixedParameterCode.AllowSeprateSchemeItemPrintDairySaleInvoice, EnumControlType.CheckBox)
        '===================
        InsertDefaultValue(clsUserMgtCode.POWeighment, clsFixedParameterType.POWeighmentManual, clsFixedParameterCode.POWeighmentManual, EnumControlType.CheckBox)
        'KUNAL > UDIL > DATE : 16-NOV-2016
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.FindNRGP_Request, clsFixedParameterCode.FindNRGP_Request, EnumControlType.CheckBox)
        ''======Ravi============
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AddTypeForUserMaster, clsFixedParameterCode.AddTypeForUserMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.EmployeeMaster, clsFixedParameterType.AddParavetEmployeeType, clsFixedParameterCode.AddParavetEmployeeType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityModuleParameterRangeMaster, clsFixedParameterType.AllowDeductionPercentOnIncoming, clsFixedParameterCode.AllowDeductionPercentOnIncoming, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SecondaryCustomerMaster, clsFixedParameterType.AllowLoginType, clsFixedParameterCode.AllowLoginType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SecondaryCustomerMaster, clsFixedParameterType.AutoGeneratedDigitsForCustomer, clsFixedParameterCode.AutoGeneratedDigitsForCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SecondaryCustomerMaster, clsFixedParameterType.AutoGeneratedDigitsForVendor, clsFixedParameterCode.AutoGeneratedDigitsForVendor, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.SecondaryCustomerMaster, clsFixedParameterType.AutoGeneratedCustomerCode, clsFixedParameterCode.AutoGeneratedCustomerCode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PriceMaster, clsFixedParameterType.AllowLoginType, clsFixedParameterCode.AllowLoginType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AllowLoginType, clsFixedParameterCode.AllowLoginType, EnumControlType.CheckBox)
        ''==============================
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.CalculateFIFOAndLIFOCosting, clsFixedParameterCode.CalculateFIFOAndLIFOCosting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.EmptyCanWeight, clsFixedParameterCode.EmptyCanWeight, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.MinuteInLastVehicleForGateEntry, clsFixedParameterCode.MinuteInLastVehicleForGateEntry, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.MinuteGateEntryToGrossWeight, clsFixedParameterCode.MinuteGateEntryToGrossWeight, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.MinuteGrossWeightToTareWeight, clsFixedParameterCode.MinuteGrossWeightToTareWeight, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.NoOfDaysForMultiInceForSameVSPForSamePayCycle, clsFixedParameterCode.NoOfDaysForMultiInceForSameVSPForSamePayCycle, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.NIRQC, clsFixedParameterType.NoOfDaysForShowFOSSDocument, clsFixedParameterCode.NoOfDaysForShowFOSSDocument, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.PurchaseCounterOnTransactionType, clsFixedParameterCode.PurchaseCounterOnTransactionType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.StopForRepeatedFATSNF, clsFixedParameterCode.StopForRepeatedFATSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SampleFONTSize, clsFixedParameterCode.SampleFONTSize, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.SampleFONTSize, clsFixedParameterCode.SampleFONTSize, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBankReco, clsFixedParameterType.TreatChequeClearDateAsRecoDate, clsFixedParameterCode.TreatChequeClearDateAsRecoDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, clsFixedParameterType.BookWreckageFromSublocationOrSection, clsFixedParameterCode.BookWreckageFromSublocationOrSection, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, clsFixedParameterType.CreateJEOnProduction, clsFixedParameterCode.CreateJEOnProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, clsFixedParameterType.PickProductCostFromItemUOMDetail, clsFixedParameterCode.PickProductCostFromItemUOMDetail, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.StopVSPBillIfSomethingWrong, clsFixedParameterCode.StopVSPBillIfSomethingWrong, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.PDCSetting, clsFixedParameterCode.PDCSetting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AllowRoadPermitNo, clsFixedParameterCode.AllowRoadPermitNo, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.ShowMessgForTDS, clsFixedParameterCode.ShowMessgForTDS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.IsShowTreeView, clsFixedParameterCode.IsShowTreeView, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.ShowVLCUploaderData, clsFixedParameterCode.ShowVLCUploaderData, EnumControlType.CheckBox)
        '==============Parteek added 09/01/2017
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.FatSnfWhenMilktypeSelect, clsFixedParameterCode.FatSnfWhenMilktypeSelect, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.DairyFreshTaxableandNonTaxable, clsFixedParameterCode.DairyFreshTaxableandNonTaxable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.CreateNewDocumentOnUploader, clsFixedParameterCode.CreateNewDocumentOnUploader, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.MilkProcurementBatchPosting, clsFixedParameterCode.MilkProcurementBatchPosting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.popupcustomernamewhileupdating, clsFixedParameterCode.popupcustomernamewhileupdating, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.ApplyFEFO, clsFixedParameterCode.ApplyFEFO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNJobWorkEstimate, clsFixedParameterType.CreateJVofPackingMaterialofJWInwardinJWEstimate, clsFixedParameterCode.CreateJVofPackingMaterialofJWInwardinJWEstimate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.journalEntry, clsFixedParameterType.AllowJEofDifferentLocationOnJournalEntry, clsFixedParameterCode.AllowJEofDifferentLocationOnJournalEntry, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.JobWorkDispatch, clsFixedParameterType.AllowtoenterrateIntoJobWorkDispatch, clsFixedParameterCode.AllowtoenterrateIntoJobWorkDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.TCSTaxApplicableOnbulkSale, clsFixedParameterCode.TCSTaxApplicableOnbulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanSale, clsFixedParameterType.TCSTaxApplicableOnCanSale, clsFixedParameterCode.TCSTaxApplicableOnCanSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.ItemCostZeroOnStoreAdjForTypeFlushing, clsFixedParameterCode.ItemCostZeroOnStoreAdjForTypeFlushing, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.journalEntry, clsFixedParameterType.PopupJE, clsFixedParameterCode.PopupJE, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.DocumentCancel, clsFixedParameterCode.DocumentCancel, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, clsFixedParameterType.DocumentCancelReturn, clsFixedParameterCode.DocumentCancelReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.CSADocumentCancel, clsFixedParameterCode.CSADocumentCancel, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.FixVSPEMP, clsFixedParameterCode.FixVSPEMP, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploader, clsFixedParameterType.ApplyRTGSAmtMoreThanGiven, clsFixedParameterCode.ApplyRTGSAmtMoreThanGiven, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmCompanyMaster, clsFixedParameterType.AllowToSaveAndUpdatePasswordBased, clsFixedParameterCode.AllowToSaveAndUpdatePasswordBased, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploaderFarmer, clsFixedParameterType.ApplyRTGSAmtMoreThanGiven, clsFixedParameterCode.ApplyRTGSAmtMoreThanGiven, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploaderFarmer, clsFixedParameterType.AllowToGenerate_NEFTUPLOADER, clsFixedParameterCode.AllowToGenerate_NEFTUPLOADER, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.AllowRoundOffAmountOnDCSSale, clsFixedParameterCode.AllowRoundOffAmountOnDCSSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.AllowRoundOffAmountOnDCSSale, clsFixedParameterCode.AllowRoundOffAmountOnDCSSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, clsFixedParameterType.SepratePriceChartForCowMilk, clsFixedParameterCode.SepratePriceChartForCowMilk, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, clsFixedParameterType.SepratePriceChartForCamelMilk, clsFixedParameterCode.SepratePriceChartForCamelMilk, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkPricePlanning, clsFixedParameterType.OpenPriceChartPlanningScreenOnTotalSolid, clsFixedParameterCode.OpenPriceChartPlanningScreenOnTotalSolid, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkPricePlanning, clsFixedParameterType.MaxFATSNFPerForRate, clsFixedParameterCode.MaxFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkPricePlanning, clsFixedParameterType.MaxFATSNFPerForRate, clsFixedParameterCode.MaxSNFPerLimit, EnumControlType.NumericBox)

        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShiftManual, clsFixedParameterType.AllowZeroQtyFATSNFInOpenMCCShift, clsFixedParameterCode.AllowZeroQtyFATSNFInOpenMCCShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.AllowZeroQtyFATSNFInCloseMCCShift, clsFixedParameterCode.AllowZeroQtyFATSNFInCloseMCCShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.RequiredPOLimit, clsFixedParameterCode.RequiredPOLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.UnitCostIncreasePurchaseInvoice, clsFixedParameterCode.UnitCostIncreasePurchaseInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.UDLPurchaseOrderthroughAP, clsFixedParameterCode.UDLPurchaseOrderthroughAP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.CreateConsumeEntry, clsFixedParameterCode.CreateConsumeEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.CreateConsumeEntry, clsFixedParameterCode.CreateConsumeEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.CreateConsumeEntry, clsFixedParameterCode.CreateConsumeEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.DoNotCreateAdjustmentonMilkTransferInGL, clsFixedParameterCode.DoNotCreateAdjustmentonMilkTransferInGL, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmIssueItemsToAsset, clsFixedParameterType.ShowOptionforSelectingCapexForFA, clsFixedParameterCode.ShowOptionforSelectingCapexForFA, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.UDLCapexAcquisionEntry, clsFixedParameterCode.UDLCapexAcquisionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.UDLRGPWiseDocument, clsFixedParameterCode.UDLRGPWiseDocument, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.AllowAssetItemOnMiscSale, clsFixedParameterCode.AllowAssetItemOnMiscSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.journalEntry, clsFixedParameterType.TriggerOfGLEntryForWinTable, clsFixedParameterCode.TriggerOfGLEntryForWinTable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturndairy, clsFixedParameterType.UOMAtDiarySaleReturn, clsFixedParameterCode.UOMAtDiarySaleReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcessFarmer, clsFixedParameterType.PayableAmountZeroForFarmerPayment, clsFixedParameterCode.PayableAmountZeroForFarmerPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.MannualySetMPUploaderData, clsFixedParameterCode.MannualySetMPUploaderData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptRoutewiseTPTimeTable, clsFixedParameterType.ShowRouteWiseAndVLCWiseReport, clsFixedParameterCode.ShowRouteWiseAndVLCWiseReport, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.VSPMPDiffrenceOnTSBasis, clsFixedParameterCode.VSPMPDiffrenceOnTSBasis, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartMaster, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.DBF, clsFixedParameterCode.FATDivideBy, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.DBF, clsFixedParameterCode.SNFDivideBy, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.PrefixGeneration, clsFixedParameterType.LinkFinancialYearStyleWithGSTDate, clsFixedParameterCode.LinkFinancialYearStyleWithGSTDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PrefixGeneration, clsFixedParameterType.ApplyLocationWisePrefix, clsFixedParameterCode.ApplyLocationWisePrefix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.AllowCreditNoteWithoutReference, clsFixedParameterCode.AllowCreditNoteWithoutReference, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.AllowUseApplyDocSeriesForReceipt, clsFixedParameterCode.AllowUseApplyDocSeriesForReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.AllowUseApplyDocSeriesForPayment, clsFixedParameterCode.AllowUseApplyDocSeriesForPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AllowCreditNoteWithoutReferenceonAP, clsFixedParameterCode.AllowCreditNoteWithoutReferenceonAP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.SecurityDocumentKnockOffonReceipt, clsFixedParameterCode.SecurityDocumentKnockOffonReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowBranchAcconReceiptPrint, clsFixedParameterCode.AllowBranchAcconReceiptPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.WeighmentNotMandatoryInMCC, clsFixedParameterCode.WeighmentNotMandatoryInMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.PurchaseModulePickFixTaxRate, clsFixedParameterCode.PurchaseModulePickFixTaxRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.GSTExemptedAmountForNonRegisteredVendor, clsFixedParameterCode.GSTExemptedAmountForNonRegisteredVendor, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TankerDispatchFinancialImpactInTransferIn, clsFixedParameterCode.TankerDispatchFinancialImpactInTransferIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.ConvertQtyIntoKG, clsFixedParameterCode.ConvertQtyIntoKG, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.AllowtoSetNoOfTransactionsforSetOff, clsFixedParameterCode.AllowtoSetNoOfTransactionsforSetOff, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.AllowtoUnlockTransactionsforSetOff, clsFixedParameterCode.AllowtoUnlockTransactionsforSetOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowtoUnlockTransactionsforSetOff, clsFixedParameterCode.AllowtoUnlockTransactionsforSetOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCustomerAgeing, clsFixedParameterType.AllowtoShowCreditBalanceonCustomerAgeing, clsFixedParameterCode.AllowtoShowCreditBalanceonCustomerAgeing, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAgingPayble, clsFixedParameterType.AllowtoShowDebitBalanceonVendorAgeing, clsFixedParameterCode.AllowtoShowDebitBalanceonVendorAgeing, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAgingPayble, clsFixedParameterType.ConsiderOpeningDocintoBucketsInAgeing, clsFixedParameterCode.ConsiderOpeningDocintoBucketsInAgeing, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCustomerAgeing, clsFixedParameterType.ConsiderOpeningDocintoBucketsonCustomerAgeing, clsFixedParameterCode.ConsiderOpeningDocintoBucketsonCustomerAgeing, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowtoNegativeStockInventoryAtTankerDispatch, clsFixedParameterCode.AllowtoNegativeStockInventoryAtTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowtoNegativeFATSNFKgAtTankerDispatch, clsFixedParameterCode.AllowtoNegativeFATSNFKgAtTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowtoSetoffDocDateWise, clsFixedParameterCode.AllowtoSetoffDocDateWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowSetOffUntilTransactionsnotend, clsFixedParameterCode.AllowSetOffUntilTransactionsnotend, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AllowtoEmployeeSalaryIntegration, clsFixedParameterCode.AllowtoEmployeeSalaryIntegration, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowtoEmployeeSalaryIntegration, clsFixedParameterCode.AllowtoEmployeeSalaryIntegration, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowtoFutureDateTransForPDCCheque, clsFixedParameterCode.AllowtoFutureDateTransForPDCCheque, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowtoSetReceiptAmountForCashTransaction, clsFixedParameterCode.AllowtoSetReceiptAmountForCashTransaction, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowtoSetPaymentAmountForCashTransaction, clsFixedParameterCode.AllowtoSetPaymentAmountForCashTransaction, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATSNFDeductionMixMilkFATMinValue, clsFixedParameterCode.FATSNFDeductionMixMilkFATMinValue, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATSNFDeductionMixMilkFATMaxValue, clsFixedParameterCode.FATSNFDeductionMixMilkFATMaxValue, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATSNFDeductionMixMilkSNFMinValue, clsFixedParameterCode.FATSNFDeductionMixMilkSNFMinValue, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATSNFDeductionMixMilkSNFMaxValue, clsFixedParameterCode.FATSNFDeductionMixMilkSNFMaxValue, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.FATSNFDeductionMixMilkDeductionPer, clsFixedParameterCode.FATSNFDeductionMixMilkDeductionPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.RoundOffPaiseAmount, clsFixedParameterCode.RoundOffPaiseAmount, EnumControlType.CheckBox)
        '' production
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ShowOverheadCostOnProductionEntry, clsFixedParameterCode.ShowOverheadCostOnProductionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ShowOverheadCostOnProductionEntry, clsFixedParameterCode.ShowOverheadCostOnProductionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ActivateProductionWithoutBatch, clsFixedParameterCode.ActivateProductionWithoutBatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ActivateProductionWithoutBatch, clsFixedParameterCode.ActivateProductionWithoutBatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.CreateJEOnProduction, clsFixedParameterCode.CreateJEOnProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.CreateJEOnProduction, clsFixedParameterCode.CreateJEOnProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilk, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilkProduct, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeOther, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, clsFixedParameterType.ConsumptionType, clsFixedParameterCode.ConsumptionTypeMilkStandardization, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.FreightProvisionAccount, clsFixedParameterCode.FreightProvisionAccount, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.FreightProvisionAccountInward, clsFixedParameterCode.FreightProvisionAccountInward, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.ChangeVehicleOnDairySaleBooking, clsFixedParameterCode.ChangeVehicleOnDairySaleBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.VendorSetOffDayWise, clsFixedParameterCode.VendorSetOffDayWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.PermissionSettingForTransactionWithBank, clsFixedParameterCode.PermissionSettingForTransactionWithBank, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.AllowUseApplyDocSeriesForPayment, clsFixedParameterCode.AllowUseApplyDocSeriesForPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.PaymentReceiptTypeRunReceiptCounter, clsFixedParameterCode.PaymentReceiptTypeRunReceiptCounter, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmVSPCSASetOff, clsFixedParameterType.AllowtoSetNoOfTransactionsforSetOff, clsFixedParameterCode.AllowtoSetNoOfTransactionsforSetOff, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmVSPCSASetOff, clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.PaymentAdjustmentEntry, clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptTruckSheetDailySummaryReport, clsFixedParameterType.CowFATPer, clsFixedParameterCode.CowFATPer, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.rptTruckSheetDailySummaryReport, clsFixedParameterType.MixFATPer, clsFixedParameterCode.MixFATPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.CowFATPer, clsFixedParameterCode.CowFATPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.CowFATPer, clsFixedParameterCode.CowFATPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.RemoveSavingDocumentWhenPayableAmtZero, clsFixedParameterCode.RemoveSavingDocumentWhenPayableAmtZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.RecalculatePaymentProcessOnSave, clsFixedParameterCode.RecalculatePaymentProcessOnSave, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.CmtSetting, clsFixedParameterCode.CmtSetting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.PaymentProcessPrintInHindi, clsFixedParameterCode.PaymentProcessPrintInHindi, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.VendorBankAdviceForSWM, clsFixedParameterCode.VendorBankAdviceForSWM, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementUploader, clsFixedParameterType.Hidedetaildate, clsFixedParameterCode.Hidedetaildate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.Hidedetaildate, clsFixedParameterCode.Hidedetaildate, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.LoadLoginScreen, clsFixedParameterType.DeleteTempData, clsFixedParameterCode.DeleteTempData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPayPeriodMaster, clsFixedParameterType.SkipDatePayPeriodMaster, clsFixedParameterCode.SkipDatePayPeriodMaster, EnumControlType.CheckBox)

        ' InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, clsFixedParameterType.MixFATPer, clsFixedParameterCode.MixFATPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.POWeighment, clsFixedParameterType.AddHighSecurityOnWeighingIntegratedScreen, clsFixedParameterCode.AddHighSecurityOnWeighingIntegratedScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.POWeighment, clsFixedParameterType.HighSecurityStableSeconds, clsFixedParameterCode.HighSecurityStableSeconds, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.POWeighment, clsFixedParameterType.HighSecurityWeightTolerance, clsFixedParameterCode.HighSecurityWeightTolerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.FreeIndentQtyAfterPOClose, clsFixedParameterCode.FreeIndentQtyAfterPOClose, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.ShowFATSNFinPaymentProcess, clsFixedParameterCode.ShowFATSNFinPaymentProcess, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.ApplyMaTransRateOnMultChamberTankerDis, clsFixedParameterCode.ApplyMaTransRateOnMultChamberTankerDis, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.SkipJobWorkSRNInPI, clsFixedParameterCode.SkipJobWorkSRNInPI, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptBulkMilkRegister, clsFixedParameterType.ShowFatSnfAfterApproval, clsFixedParameterCode.ShowFatSnfAfterApproval, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.ApplyTotalSolidPriceChart, clsFixedParameterCode.ApplyTotalSolidPriceChart, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, clsFixedParameterType.RequiredMgmtApprovalForRateIncrease, clsFixedParameterCode.RequiredMgmtApprovalForRateIncrease, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AutoRoundOffSeprateAccountOnVendorTransaction, clsFixedParameterCode.AutoRoundOffSeprateAccountOnVendorTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.PickOnAccountPaymentForAdvanceKnockOff, clsFixedParameterCode.PickOnAccountPaymentForAdvanceKnockOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.DoNotConsiderTheFutureDateOfAdvancePayment, clsFixedParameterCode.DoNotConsiderTheFutureDateOfAdvancePayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryOthers, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.UDLPurchaseOrderthroughAP, clsFixedParameterCode.UDLPurchaseOrderthroughAP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.ShowHierarchyAndCostCenterInAPInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInAPInvoiceEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.AutoRoundOffSeprateAccountOnVendorTransaction, clsFixedParameterCode.AutoRoundOffSeprateAccountOnVendorTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.GSTExemptedAmountForNonRegisteredVendor, clsFixedParameterCode.GSTExemptedAmountForNonRegisteredVendor, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.AllowSameaAdditionalChargesMultiTime, clsFixedParameterCode.AllowSameaAdditionalChargesMultiTime, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.CreateSeperateSeriesforRefDocAPinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocAPinvforCreditdebit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.DoNotConsiderTheFutureDateOfAdvancePayment, clsFixedParameterCode.DoNotConsiderTheFutureDateOfAdvancePayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.PickOnAccountPaymentForAdvanceKnockOff, clsFixedParameterCode.PickOnAccountPaymentForAdvanceKnockOff, EnumControlType.CheckBox)
        '=====================Added by preeti Gupta against Ticket No[ADV/17/05/18-000032]
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.TreatCRATEAsItems, clsFixedParameterCode.TreatCRATEAsItems, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.TreatCANAsItems, clsFixedParameterCode.TreatCANAsItems, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.DoNotShowDairyTypeItems, clsFixedParameterCode.DoNotShowDairyTypeItems, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.AlwaysVSPDefaulter, clsFixedParameterCode.AlwaysVSPDefaulter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.RejectedMilkSendToRejectLocation, clsFixedParameterCode.RejectedMilkSendToRejectLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.NoOfPreNxtDayToPickAvgFATSNF, clsFixedParameterCode.NoOfPreNxtDayToPickAvgFATSNF, EnumControlType.NumericBox)
        '================================= Start ======================================================================================================================= 
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.ApplyDocumentDate, clsFixedParameterCode.ApplyDocumentDate, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, EnumControlType.TextBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.IsAutoReceiptPayment, clsFixedParameterCode.IsAutoReceiptPayment, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.DefaultBank, clsFixedParameterCode.DefaultBank, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.DefaultLocation, clsFixedParameterCode.DefaultLocation, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.SecurityDocumentKnockOffonReceipt, clsFixedParameterCode.SecurityDocumentKnockOffonReceipt, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowBranchAcconReceiptPrint, clsFixedParameterCode.AllowBranchAcconReceiptPrint, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.CustomerMasterFinderOnLocationwiseARReceipt, clsFixedParameterCode.CustomerMasterFinderOnLocationwiseARReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.ShowHierarchyAndCostCenterInARInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInARInvoiceEntry, EnumControlType.CheckBox)
        ' InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.EnableDistributorSubsidy, clsFixedParameterCode.EnableDistributorSubsidy, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.AllowUseApplyDocSeriesForReceipt, clsFixedParameterCode.AllowUseApplyDocSeriesForReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.CustomerMasterFinderOnLocationwiseARReceipt, clsFixedParameterCode.CustomerMasterFinderOnLocationwiseARReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.AutoRecieptPaymentMode, clsFixedParameterCode.AutoRecieptPaymentMode, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowUseApplyDocSeriesForReceipt, clsFixedParameterCode.AllowUseApplyDocSeriesForReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.SecurityDocumentKnockOffonReceipt, clsFixedParameterCode.SecurityDocumentKnockOffonReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.AllowtoSetNoOfTransactionsforSetOff, clsFixedParameterCode.AllowtoSetNoOfTransactionsforSetOff, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.CustomerMasterFinderOnLocationwiseARReceipt, clsFixedParameterCode.CustomerMasterFinderOnLocationwiseARReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPayableSettings, clsFixedParameterType.AutoGeneratedDigitsForVendor, clsFixedParameterCode.AutoGeneratedDigitsForVendor, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendoraccountset, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendorgroup, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendorgroup, clsFixedParameterType.IsVendorGroupFieldsMandatory, clsFixedParameterCode.IsVendorGroupFieldsMandatory, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.FrmReceiptInvoiceMapping, clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PrefixGeneration, clsFixedParameterType.AutoGeneratePrefix, clsFixedParameterCode.AutoGeneratePrefix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersOutstanding, clsFixedParameterType.PenaltyPercentage, clsFixedParameterCode.PenaltyPercentage, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersOutstanding, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmReceivableSettings, clsFixedParameterType.AutoGeneratedDigitsForCustomer, clsFixedParameterCode.AutoGeneratedDigitsForCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerType, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerAccountSet, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerAccountSet, clsFixedParameterType.EnableDistributorSubsidy, clsFixedParameterCode.EnableDistributorSubsidy, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerGroup, clsFixedParameterType.IsCustomerGroupFieldsMandatory, clsFixedParameterCode.IsCustomerGroupFieldsMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomeCategory, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBankUpdateUploader, clsFixedParameterType.SkipLockLocation, clsFixedParameterCode.SkipLockLocation, EnumControlType.CheckBox)
        ' 04 June 2018 --------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShift, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShift, clsFixedParameterType.AllowZeroQtyFATSNFInOpenMCCShift, clsFixedParameterCode.AllowZeroQtyFATSNFInOpenMCCShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryWeightment, clsFixedParameterType.WeighmentNotMandatoryInMCC, clsFixedParameterCode.WeighmentNotMandatoryInMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkReject, clsFixedParameterType.WeighmentNotMandatoryInMCC, clsFixedParameterCode.WeighmentNotMandatoryInMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.RoundOffPaiseAmount, clsFixedParameterCode.RoundOffPaiseAmount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.AddValidationofMilkTypeinsample, clsFixedParameterCode.AddValidationofMilkTypeinsample, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryOut, clsFixedParameterType.WeighmentNotMandatoryInMCC, clsFixedParameterCode.WeighmentNotMandatoryInMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.PickPriceFromFATAndSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVLCDataUploaderManual, clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.PTMRatePerLtrKGOnStdQty, clsFixedParameterCode.PTMRatePerLtrKGOnStdQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.defaultCorrectionFactor, clsFixedParameterCode.MilkSetting, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCGateEntry, clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCWeighment, clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowFractionInMCCTankerDispatchGrossQty, clsFixedParameterCode.AllowFractionInMCCTankerDispatchGrossQty, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.IsUOMSelectableOnMCCDispatch, clsFixedParameterCode.IsUOMSelectableOnMCCDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.UpdateInventorySummaryTable, clsFixedParameterCode.UpdateInventorySummaryTable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        ' InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.DisAllowIntermittentTankerForPlantDispatch, clsFixedParameterCode.DisAllowIntermittentTankerForPlantDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.ItemDescForTankerdispatchPrint, clsFixedParameterCode.ItemDescForTankerDispatchPrint, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MCCDispatchReturn, clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTankerDispatchGL, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MCCDispatchReturn, clsFixedParameterType.TransferEntryOnInvCtrlAccount, clsFixedParameterCode.TransferEntryOnInvCtrlAccount, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MCCDispatchReturn, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MCCDispatchReturn, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, clsFixedParameterType.AutoRoundOffSeprateAccountOnVendorTransaction, clsFixedParameterCode.AutoRoundOffSeprateAccountOnVendorTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, clsFixedParameterType.NoOfDaysForMultiInceForSameVSPForSamePayCycle, clsFixedParameterCode.NoOfDaysForMultiInceForSameVSPForSamePayCycle, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, clsFixedParameterType.VSPMPDiffrenceOnTSBasis, clsFixedParameterCode.VSPMPDiffrenceOnTSBasis, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmVSPItemIssue, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPIssuePayment, clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPIssuePayment, clsFixedParameterType.StopVSPBillIfSomethingWrong, clsFixedParameterCode.StopVSPBillIfSomethingWrong, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPIssuePayment, clsFixedParameterType.RoundOffPaiseAmount, clsFixedParameterCode.RoundOffPaiseAmount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPIssuePayment, clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowSkippingPrevDocumentsOnPaymentProcess, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.VSPHoldPaymentNotCompanyBank, clsFixedParameterCode.VSPHoldPaymentNotCompanyBank, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.TDSRoundOffAmount, clsFixedParameterCode.TDSRoundOffAmount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmMCCMilkTransPortorInvoice, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.RoundOffPaiseAmount, clsFixedParameterCode.RoundOffPaiseAmount, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.DipMarkingMendatory, clsFixedParameterCode.DipMarkingMendatory, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryMCCSAle, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        ' Payable 
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.ShowHierarchyAndCostCenterInAPInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInAPInvoiceEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AlowwdateChangeinPaymentEntry, clsFixedParameterCode.AlowwdateChangeinPaymentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowUseApplyDocSeriesForPayment, clsFixedParameterCode.AllowUseApplyDocSeriesForPayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowtoSkipJournalEntryofPaymentandReceiptforAD, clsFixedParameterCode.AllowtoSkipJournalEntryofPaymentandReceiptforAD, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.PermissionSettingForTransactionWithBank, clsFixedParameterCode.PermissionSettingForTransactionWithBank, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.ShowHierarchyAndCostCenterInAPInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInAPInvoiceEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.CheckDocAmountInAPInvoiceEntry, clsFixedParameterCode.CheckDocAmountInAPInvoiceEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, clsFixedParameterType.AllowSetOffUntilTransactionsnotend, clsFixedParameterCode.AllowSetOffUntilTransactionsnotend, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.SingleUserParticularDairyBookingEdit, clsFixedParameterCode.SingleUserParticularDairyBookingEdit, EnumControlType.CheckBox)
        '================================= End   =======================================================================================================================
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MaxReceiveSNFPer, clsFixedParameterCode.MaxReceiveSNFPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.isCreateBulkProcPriceChartItemWise, clsFixedParameterCode.isCreateBulkProcPriceChartItemWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.DonotCheckAvgQtyOnDairyBooking, clsFixedParameterCode.DonotCheckAvgQtyOnDairyBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.FirstGateOutProcessForMCCBulkProcument, clsFixedParameterCode.FirstGateOutProcessForMCCBulkProcument, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.MCCBulkProcumentSecurityGateOut, clsFixedParameterCode.MCCBulkProcumentSecurityGateOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.DoNotAllowSavePOWhenQtyNRateZero, clsFixedParameterCode.DoNotAllowSavePOWhenQtyNRateZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ShowOnlyProductionItemInAdRemove, clsFixedParameterCode.ShowOnlyProductionItemInAdRemove, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.CheckLiveStockInProductionDuringTrans, clsFixedParameterCode.CheckLiveStockInProductionDuringTrans, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.OpenAvailorEmptyStckLocationOn_Standardization, clsFixedParameterCode.OpenAvailorEmptyStckLocationOn_Standardization, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ShowOnlyProductionItemInAdRemove, clsFixedParameterCode.ShowOnlyProductionItemInAdRemove, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.ShowOnlyProductionItemInAdRemove, clsFixedParameterCode.ShowOnlyProductionItemInAdRemove, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.AllowAutoCalculateADDREMOVEQty, clsFixedParameterCode.AllowAutoCalculateADDREMOVEQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.ShowOverheadCostOnProductionEntry, clsFixedParameterCode.ShowOverheadCostOnProductionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.KnockOffFATSNFKG, clsFixedParameterCode.KnockOffFATSNFKG, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, clsFixedParameterType.CreateLoadINSlipVehicleWise, clsFixedParameterCode.CreateLoadINSlipVehicleWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, clsFixedParameterType.RouteCodeNotMandatoryOnLoadINSlip, clsFixedParameterCode.RouteCodeNotMandatoryOnLoadINSlip, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ConsiderSiloCapicityForStockIn, clsFixedParameterCode.ConsiderSiloCapicityForStockIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.locationMaster, clsFixedParameterType.AllowmultipleconsumptionLocation, clsFixedParameterCode.AllowmultipleconsumptionLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.ItemwiseCorrectionFactoronQC, clsFixedParameterCode.ItemwiseCorrectionFactoronQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.ParameterForSNFatQC, clsFixedParameterCode.ParameterForSNFatQC, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCSMultipleDays, clsFixedParameterType.TreatCLR50As52, clsFixedParameterCode.TreatCLR50As52, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AutoMilkTransferInDateSameasWeighmentDate, clsFixedParameterCode.AutoMilkTransferInDateSameasWeighmentDate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt, clsFixedParameterCode.EnableGoButtonofReceiptEntryWithoutEnteringReceiptAmt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, clsFixedParameterType.ShowOutstandingAmtofCustomerOnQuickBookEntry, clsFixedParameterCode.ShowOutstandingAmtofCustomerOnQuickBookEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, clsFixedParameterType.Allow0FatPerOnBulkSaleQualityCheckScreen, clsFixedParameterCode.Allow0FatPerOnBulkSaleQualityCheckScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.CrateReceiveddairyRouteWise, clsFixedParameterCode.CrateReceiveddairyRouteWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyFreshDispatchMultiple, clsFixedParameterType.CreateMultipleDispatchWithoutSelectingVehicle, clsFixedParameterCode.CreateMultipleDispatchWithoutSelectingVehicle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.RejectiononQCforSeparationofBulkProcurementMCC, clsFixedParameterCode.RejectiononQCforSeparationofBulkProcurementMCC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.PickCostFromItemMaster, clsFixedParameterCode.PickCostFromItemMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.EditItemCost, clsFixedParameterCode.EditItemCost, EnumControlType.CheckBox)
        'Sanjay TEC/04/06/18-000273
        InsertDefaultValue(clsUserMgtCode.VendorItemDetails, clsFixedParameterType.STDPURRATE, clsFixedParameterCode.STDPURRATE, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, clsFixedParameterType.ShowStatusForPurchase, clsFixedParameterCode.ShowStatusForPurchase, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.itemPurchaseAccount, clsFixedParameterType.ShowPurchaseControlAc, clsFixedParameterCode.ShowPurchaseControlAc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.ShowCostCenterAndHierarchyLevelInPurchaseModule, clsFixedParameterCode.ShowCostCenterAndHierarchyLevelInPurchaseModule, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, clsFixedParameterType.ShowCostCenterAndHierarchyLevelInPurchaseModule, clsFixedParameterCode.ShowCostCenterAndHierarchyLevelInPurchaseModule, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmTender, clsFixedParameterType.ShowUpto6DecimalPlaces, clsFixedParameterCode.ShowUpto6DecimalPlaces, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.BMC_Transporter_Bill, clsFixedParameterType.EnableOnPrivateChkbox, clsFixedParameterCode.EnableOnPrivateChkbox, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.PurchaseModule, clsFixedParameterCode.RateDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.IsRGPAfterPurchaseOrder, clsFixedParameterCode.IsRGPAfterPurchaseOrder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.IsRemarkReasonMandatoryOnPO, clsFixedParameterCode.IsRemarkReasonMandatoryOnPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.PurchasePickItemFromVendorItemDetails, clsFixedParameterCode.PurchasePickItemFromVendorItemDetails, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, clsFixedParameterType.IsQCColumnRequiredonMRN, clsFixedParameterCode.IsQCColumnRequiredonMRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AllowLargerItemCostThenVendorItemCost, clsFixedParameterCode.AllowLargerItemCostThenVendorItemCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AcceptQtyWithoutConsiderRejectQty, clsFixedParameterCode.AcceptQtyWithoutConsiderRejectQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.IsRateEditableOnSRN, clsFixedParameterCode.IsRateEditableOnSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.AutoPOAtSRN, clsFixedParameterCode.AutoPOAtSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.EnableRackBin, clsFixedParameterCode.EnableRackBin, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.EnableStoreCostCentre, clsFixedParameterCode.EnableStoreCostCentre, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, clsFixedParameterType.IsCostEditableOnIssueReturnTransfer, clsFixedParameterCode.IsCostEditableOnIssueReturnTransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMaterialQuotationComparison, clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, clsFixedParameterType.EnableStoreCostCentre, clsFixedParameterCode.EnableStoreCostCentre, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.JobWorkDispatch, clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryMISSAle, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        'Sanjay TEC/04/06/18-000273
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.ShowSiloLocationItemLocationwise, clsFixedParameterCode.ShowSiloLocationItemLocationwise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.ItemCrateRate, clsFixedParameterCode.ItemCrateRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.ItemCanRate, clsFixedParameterCode.ItemCanRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmGeneralWeighment, clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmJobWorkInventory, clsFixedParameterType.AllowToShowMilkTypeinAdjustmentEntry, clsFixedParameterCode.AllowToShowMilkTypeinAdjustmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmJobWorkInventory, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmRawMilkConsumtion, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.PROVISIONENTRYONSTOCKTRANSFER, clsFixedParameterCode.PROVISIONENTRYONSTOCKTRANSFER, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.DisplayPriceCodeForBKN, clsFixedParameterCode.DisplayPriceCodeForBKN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.TransferReturn, clsFixedParameterType.AllowGateReturn, clsFixedParameterCode.AllowGateReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.RequiredFinalQCofstandardization, clsFixedParameterCode.RequiredFinalQCofstandardization, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.CheckLiveStockInProductionDuringTrans, clsFixedParameterCode.CheckLiveStockInProductionDuringTrans, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionFATSNFPerDecimalPoint, clsFixedParameterCode.ProductionFATSNFPerDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.AllowNegativeStockInDairyProduction, clsFixedParameterCode.AllowNegativeStockInDairyProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.PickOnAccountPaymentForAdvanceKnockOff, clsFixedParameterCode.PickOnAccountPaymentForAdvanceKnockOff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.PickProductCostFromItemUOMDetail, clsFixedParameterCode.PickProductCostFromItemUOMDetail, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeightConversion, clsFixedParameterType.ItemStructureMandatoryOnWeightConversion, clsFixedParameterCode.ItemStructureMandatoryOnWeightConversion, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.RequiredFinalQCofProductionEntry, clsFixedParameterCode.RequiredFinalQCofProductionEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.PickFATSNFPerFromStock, clsFixedParameterCode.PickFATSNFPerFromStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.ShowBulkDispatchQtyInLtr, clsFixedParameterCode.ShowBulkDispatchQtyInLtr, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.DoNotIncludeIncentiveInMilkPurchaseInvoice, clsFixedParameterCode.DoNotIncludeIncentiveInMilkPurchaseInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.DoNotConsiderTheFutureDateOfAdvancePayment, clsFixedParameterCode.DoNotConsiderTheFutureDateOfAdvancePayment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEmployee_Master, clsFixedParameterType.CreateEmpCodeAsPerEmployeeBasisType, clsFixedParameterCode.CreateEmpCodeAsPerEmployeeBasisType, EnumControlType.CheckBox)
        ' By Prabhakar==================================================================================================================
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.PickPendingMilkSRNinNextPaymentCycle, clsFixedParameterCode.PickPendingMilkSRNinNextPaymentCycle, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.StopVSPBillIfSomethingWrong, clsFixedParameterCode.StopVSPBillIfSomethingWrong, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.DoNotIncludeIncentiveInMilkPurchaseInvoice, clsFixedParameterCode.DoNotIncludeIncentiveInMilkPurchaseInvoice, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.MilkMPPayment, clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.ControlSampleMandatory, clsFixedParameterCode.MilkSetting, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.MilkProc, clsFixedParameterCode.IsItemEditableOnMCCDispatch, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.MilkProc, clsFixedParameterCode.IsUOMSelectableOnMCCDispatch, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.MilkProc, clsFixedParameterCode.DisAllowIntermittentTankerForPlantDispatch, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTankerDispatchGL, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.TransferEntryOnInvCtrlAccount, clsFixedParameterCode.TransferEntryOnInvCtrlAccount, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerDispatchReturn, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.FATSNFRateMandatory, clsFixedParameterCode.FATSNFRateMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.PickCostFromItemMaster, clsFixedParameterCode.PickCostFromItemMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.CreateJEOnProduction, clsFixedParameterCode.CreateJEOnProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.PickProductCostFromItemUOMDetail, clsFixedParameterCode.PickProductCostFromItemUOMDetail, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.AllowVehicleGateOutValidationScrapSale, clsFixedParameterCode.AllowVehicleGateOutValidationScrapSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesExciseTypeatPS, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesExciseTypeatPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesExciseatMiscSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesExciseatMiscSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesTaxatMiscSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesTaxatMiscSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSectionMaster, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmStageMaster, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSectionStageMapping, clsFixedParameterType.AllowToSkipStageQLLogSheetInProd, clsFixedParameterCode.AllowToSkipStageQLLogSheetInProd, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialDairy, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialDairy, clsFixedParameterType.ProductionFATSNF_KG_Unit, clsFixedParameterCode.ProductionFATSNF_KG_Unit, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialDairy, clsFixedParameterType.RemoveGridFatSnfValidation, clsFixedParameterCode.RemoveGridFatSnfValidation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningDairy, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningDairy, clsFixedParameterType.IsConsiderOutTypeDocForBalance, clsFixedParameterCode.IsConsiderOutTypeDocForBalance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ProductionQtyDecimalPoint, clsFixedParameterCode.ProductionQtyDecimalPoint, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ALLOWCBOSBO, clsFixedParameterCode.ALLOWCBOSBO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ActivateSFGProduction, clsFixedParameterCode.ActivateSFGProduction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ALLOWANYBO, clsFixedParameterCode.ALLOWANYBO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.MCCDisplay_All_Parameter, clsFixedParameterCode.MilkSetting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionOrStandAccordingToItemType, clsFixedParameterCode.ProductionOrStandAccordingToItemType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.ProductionOrStandAccordingToItemType, clsFixedParameterCode.ProductionOrStandAccordingToItemType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ProductionOrStandAccordingToItemType, clsFixedParameterCode.ProductionOrStandAccordingToItemType, EnumControlType.CheckBox)
        'Fixed Asset
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowItemConversionAutomation, clsFixedParameterCode.AllowItemConversionAutomation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowItemWiseCSAAccountingON_CSASale, clsFixedParameterCode.AllowItemWiseCSAAccountingON_CSASale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.INDUSTRYTYPE, clsFixedParameterCode.INDUSTRYTYPE, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.CreateGLAccToItem, clsFixedParameterCode.CreateGLAccToItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.ShowOptionOnItemMasterChangeItemRate, clsFixedParameterCode.ShowOptionOnItemMasterChangeItemRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PricePlan, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PriceMaster, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.locationMaster, clsFixedParameterType.AP_INV_COMMSN, clsFixedParameterCode.AP_INV_COMMSN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmOverheadCostGroup, clsFixedParameterType.IncludeRatePerHoursIn, clsFixedParameterCode.IncludeRatePerHoursIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemCostMapping, clsFixedParameterType.IncludeRatePerHoursIn, clsFixedParameterCode.IncludeRatePerHoursIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Template, clsFixedParameterType.AllowAssetBookChangeInTemplate, clsFixedParameterCode.AllowAssetBookChangeInTemplate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.ShowOptionforSelectingCapexForFA, clsFixedParameterCode.ShowOptionforSelectingCapexForFA, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.AllowRoundInFixedAsset, clsFixedParameterCode.AllowRoundInFixedAsset, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.AllowDecimalInFixedAsset, clsFixedParameterCode.AllowDecimalInFixedAsset, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.PartialFADepDays, clsFixedParameterCode.PartialFADepDays, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.RateMultPartialFADepDays, clsFixedParameterCode.RateMultPartialFADepDays, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, clsFixedParameterType.DepreciationCalculationMethod, clsFixedParameterCode.DepreciationCalculationMethod, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmAssetStoreRequistion, clsFixedParameterType.NotificationSettingforReOrderInPurchaseRequisition, clsFixedParameterCode.NotificationSettingforReOrderInPurchaseRequisition, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssetStoreRequistion, clsFixedParameterType.EnableProjectFinder, clsFixedParameterCode.EnableProjectFinder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmAssetStoreRequistion, clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetWork, clsFixedParameterType.ShowOptionforSelectingCapexForFA, clsFixedParameterCode.ShowOptionforSelectingCapexForFA, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAssetWork, clsFixedParameterType.UDLCapexAcquisionEntry, clsFixedParameterCode.UDLCapexAcquisionEntry, EnumControlType.CheckBox)
        '===============================================================================================================================
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.defaultCorrectionFactor, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.isPriceChartGradeWise, clsFixedParameterCode.isPriceChartGradeWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.FirstGateOutProcessForMCCBulkProcument, clsFixedParameterCode.FirstGateOutProcessForMCCBulkProcument, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.JOBdefaultCorrectionFactorBS, clsFixedParameterCode.JOBdefaultCorrectionFactorBS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.PurchasedefaultCorrectionFactorBS, clsFixedParameterCode.PurchasedefaultCorrectionFactorBS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.MCCdefaultCorrectionFactorBS, clsFixedParameterCode.MCCdefaultCorrectionFactorBS, EnumControlType.NumericBox)
        '' Unloading
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.isCreateBulkProcPriceChartItemWise, clsFixedParameterCode.isCreateBulkProcPriceChartItemWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, clsFixedParameterType.ShowSiloLocationItemLocationwise, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        '' Cleaning
        InsertDefaultValue(clsUserMgtCode.frmCleaning, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCleaning, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        '' Gateout
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.ShowBothTankertypeOnCleaning, clsFixedParameterCode.ShowBothTankertypeOnCleaning, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        '' Bulk Milk SRN 
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.BulkProcurementCounterOnEntryType, clsFixedParameterCode.BulkProcurementCounterOnEntryType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.CreateAutoMilkRGPinBulkSRN, clsFixedParameterCode.CreateAutoMilkRGPinBulkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.RunBulkProcOnAdjustedFATCLR, clsFixedParameterCode.RunBulkProcOnAdjustedFATCLR, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.isPriceChartGradeWise, clsFixedParameterCode.isPriceChartGradeWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowRandomOnlyOneSecondaryQC, clsFixedParameterCode.AllowRandomOnlyOneSecondaryQC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowManualPriceONBulkPO, clsFixedParameterCode.AllowManualPriceONBulkPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.BulkProcPriceChartStandardRateWithZero, clsFixedParameterCode.BulkProcPriceChartStandardRateWithZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowBulkProcTransDateSameasGateEntryDate, clsFixedParameterCode.AllowBulkProcTransDateSameasGateEntryDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.isCreateBulkProcPriceChartItemWise, clsFixedParameterCode.isCreateBulkProcPriceChartItemWise, EnumControlType.CheckBox)
        '' Bulk Milk Purchase Invoice Multiple
        InsertDefaultValue(clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple, clsFixedParameterType.BulkProcPriceChartStandardRateWithZero, clsFixedParameterCode.BulkProcPriceChartStandardRateWithZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple, clsFixedParameterType.Allow_AmountTruncate_BulkMilkSRN, clsFixedParameterCode.Allow_AmountTruncate_BulkMilkSRN, EnumControlType.CheckBox)
        '' Bulk Milk Purchase Invoice
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.BulkProcPriceChartStandardRateWithZero, clsFixedParameterCode.BulkProcPriceChartStandardRateWithZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.RunBulkProcOnAdjustedFATCLR, clsFixedParameterCode.RunBulkProcOnAdjustedFATCLR, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        ''Bulk Purchase\Sale Uploader
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.StockToleranceLimit, clsFixedParameterCode.StockToleranceLimit, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.AllowtoSkipfunctionalityafterSRNOnBulkProcurement, clsFixedParameterCode.AllowtoSkipfunctionalityafterSRNOnBulkProcurement, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        '' Milk Transfer In
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.isCreateBulkProcPriceChartItemWise, clsFixedParameterCode.isCreateBulkProcPriceChartItemWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.AllowBulkProcMCCwithoutTankerDispatch, clsFixedParameterCode.AllowBulkProcMCCwithoutTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.AllowReverseUnpost, clsFixedParameterCode.AllowReverseUnpost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.TankerDispatchFinancialImpactInTransferIn, clsFixedParameterCode.TankerDispatchFinancialImpactInTransferIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.MilkProc, clsFixedParameterCode.CreateTransferInGL, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, EnumControlType.CheckBox)
        '' Provision Entry
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.FreightProvisionAccount, clsFixedParameterCode.FreightProvisionAccount, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.FreightProvisionAccountInward, clsFixedParameterCode.FreightProvisionAccountInward, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.CreateJEForProvisionEntry, clsFixedParameterCode.CreateJEForProvisionEntryOthers, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        '' Bulk Milk SRN Return
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRNReturn, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRNReturn, clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRNReturn, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        ''MCC Milk Transfer Price
        InsertDefaultValue(clsUserMgtCode.MccMilkTransferPrice, clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, EnumControlType.CheckBox)
        ''Bulk Milk Purchase Return
        InsertDefaultValue(clsUserMgtCode.FrmMilkPurchaseReturn, clsFixedParameterType.AllowJobWorkonGateEntryBulkProc, clsFixedParameterCode.AllowJobWorkonGateEntryBulkProc, EnumControlType.CheckBox)
        '' Milk Transfer In Return
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.AllowBulkProcMCCwithoutTankerDispatch, clsFixedParameterCode.AllowBulkProcMCCwithoutTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.AllowReverseUnpost, clsFixedParameterCode.AllowReverseUnpost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.TankerDispatchFinancialImpactInTransferIn, clsFixedParameterCode.TankerDispatchFinancialImpactInTransferIn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.UseProductFATSNFKgForEstimationCost, clsFixedParameterCode.UseProductFATSNFKgForEstimationCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.UseProductFATSNFKgForEstimationCost, clsFixedParameterCode.UseProductFATSNFKgForEstimationCost, EnumControlType.CheckBox)
        '' Parameter Range Master
        InsertDefaultValue(clsUserMgtCode.frmParameterRangeMaster, clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, EnumControlType.CheckBox)
        '' Paramter Range Master for QC
        InsertDefaultValue(clsUserMgtCode.frmParameterRangeMasterForQC, clsFixedParameterType.AllowDeductionPercentOnIncoming, clsFixedParameterCode.AllowDeductionPercentOnIncoming, EnumControlType.CheckBox)
        '' Price Chart Master (Bulk Procurement)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.ExpiryDaysBulkProcurementPriceChart, clsFixedParameterCode.ExpiryDaysBulkProcurementPriceChart, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.isPriceChartGradeWise, clsFixedParameterCode.isPriceChartGradeWise, EnumControlType.CheckBox)
        '' Vendor Price Chart Mapping
        InsertDefaultValue(clsUserMgtCode.frmVendorPriceChartMapping, clsFixedParameterType.isPriceChartGradeWise, clsFixedParameterCode.isPriceChartGradeWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVendorPriceChartMapping, clsFixedParameterType.isItemMilkType, clsFixedParameterCode.isItemMilkType, EnumControlType.CheckBox)
        '' Supplier Master
        InsertDefaultValue(clsUserMgtCode.FrmSupplierMaster, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        '' Diverted Contractor
        InsertDefaultValue(clsUserMgtCode.frmDivertedContractor, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        '' Milk Type Master
        InsertDefaultValue(clsUserMgtCode.frmMilkTypeMast, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        '' Milk Grade Master
        InsertDefaultValue(clsUserMgtCode.frmMilkGradeMaster, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        '' Bulk Sale Order
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrderBS, clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrderBS, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        '' Bulk Gate Entry
        InsertDefaultValue(clsUserMgtCode.FrmGateEntrySale, clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, EnumControlType.CheckBox)
        '' Bulk Weighment
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, clsFixedParameterType.IsAutoTankerWeightment, clsFixedParameterCode.IsAutoTankerWeightment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, clsFixedParameterType.IsAutoTankerWeighmentForBulkSale, clsFixedParameterCode.IsAutoTankerWeighmentForBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        '' Bulk Loading Tanker
        InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, clsFixedParameterType.BSDefaultMilkItem, clsFixedParameterCode.BSDefaultMilkItem, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, clsFixedParameterType.StockToleranceLimit, clsFixedParameterCode.StockToleranceLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, clsFixedParameterType.IsConsiderOutTypeDocForBalance, clsFixedParameterCode.IsConsiderOutTypeDocForBalance, EnumControlType.CheckBox)
        '' Bulk Quality Check
        InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, clsFixedParameterType.BulkQCWithoutCLR, clsFixedParameterCode.BulkQCWithoutCLR, EnumControlType.CheckBox)
        '' Bulk Dispatch
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        '' Bulk Invoice
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.ApplyTSPriceAtBulkSale, clsFixedParameterCode.ApplyTSPriceAtBulkSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.ShowBulkDispatchQtyInLtr, clsFixedParameterCode.ShowBulkDispatchQtyInLtr, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.ApplyMultiChamberInBulkWeighmentEntry, clsFixedParameterCode.ApplyMultiChamberInBulkWeighmentEntry, EnumControlType.CheckBox)
        '' Bulk Dispatch Trade
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, clsFixedParameterType.BulkSaleDefaultMilkItem, clsFixedParameterCode.BulkSaleDefaultMilkItem, EnumControlType.TextBox)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        '' Bulk Dispatah Return
        InsertDefaultValue(clsUserMgtCode.FrmBulkDispatchReturnSale, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        '' Can Sale Uploader
        InsertDefaultValue(clsUserMgtCode.FrmCanSaleUploader, clsFixedParameterType.ItemCanRate, clsFixedParameterCode.ItemCanRate, EnumControlType.NumericBox)
        '' Can Sale
        InsertDefaultValue(clsUserMgtCode.FrmCanSale, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        '' Bulk Dispatch Trade Return
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTradeReturn, clsFixedParameterType.showPostrequiredforBulkSale, clsFixedParameterCode.showPostrequiredforBulkSale, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTradeReturn, clsFixedParameterType.AllowSNFNotManditoryInBulkSale, clsFixedParameterCode.AllowSNFNotManditoryInBulkSale, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTradeReturn, clsFixedParameterType.BulkSaleDefaultMilkItem, clsFixedParameterCode.BulkSaleDefaultMilkItem, EnumControlType.TextBox)
        '' Sales Setting
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.BulkSaleDefaultMilkItem, clsFixedParameterCode.BulkSaleDefaultMilkItem, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.BSDefaultMilkItem, clsFixedParameterCode.BSDefaultMilkItem, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.ItemTypeForDairyBooking, clsFixedParameterCode.ItemTypeForDairyBooking, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AutoRecieptBankCode, clsFixedParameterCode.AutoRecieptBankCode, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AutoRecieptPaymentMode, clsFixedParameterCode.AutoRecieptPaymentMode, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.DefaultItemUOMForBulkSale, clsFixedParameterCode.DefaultItemUOMForBulkSale, EnumControlType.CheckBox)
        '' Tanker Master Sale
        InsertDefaultValue(clsUserMgtCode.TankerMasterSale, clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, EnumControlType.CheckBox)
        'Export Sale
        InsertDefaultValue(clsUserMgtCode.frmEXSalesQuotation, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesOrder, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXPorformaInvoice, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXCommercialInvoice, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXCommercialInvoice, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesInvoice, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesInvoice, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesInvoice, clsFixedParameterType.EnableGSTRelatedfields, clsFixedParameterCode.EnableGSTRelatedfields, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesReturn, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesReturn, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryExportSAle, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        '' Product Booking Entry
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        ''Production Advice
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, clsFixedParameterType.AllowFreshPriceChartOnBookingProductSale, clsFixedParameterCode.AllowFreshPriceChartOnBookingProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        '' Product Delivery Order
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.AdvanceAgainstSO, clsFixedParameterCode.AdvanceAgainstSO, EnumControlType.CheckBox)
        '' Product Sale Order
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.DoNotConsiderCustomerCreditLimit, clsFixedParameterCode.DoNotConsiderCustomerCreditLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AllowDifferentStateofChildCustomerOnPS, clsFixedParameterCode.AllowDifferentStateofChildCustomerOnPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AllowFreshPriceChartOnBookingProductSale, clsFixedParameterCode.AllowFreshPriceChartOnBookingProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AllowStockCheckatDOLevel, clsFixedParameterCode.AllowStockCheckatDOLevel, EnumControlType.CheckBox)
        '' Product Dispatch
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowDifferentStateofChildCustomerOnPS, clsFixedParameterCode.AllowDifferentStateofChildCustomerOnPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.GrossWeightUnit, clsFixedParameterCode.GrossWeightUnit, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.VehicleCapacityUnit, clsFixedParameterCode.VehicleCapacityUnit, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowStockCheckatDOLevel, clsFixedParameterCode.AllowStockCheckatDOLevel, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.InvoiceManualNoWithPrefix, clsFixedParameterCode.InvoiceManualNoWithPrefix, EnumControlType.CheckBox)
        '' Product Invoice
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesRetailTypeatPS, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesRetailTypeatPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.SinglePrintCopyDairyInvoice, clsFixedParameterCode.SinglePrintCopyDairyInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.InvoiceManualNoWithPrefix, clsFixedParameterCode.InvoiceManualNoWithPrefix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.AllowManualvehicleOnDairyBooking, clsFixedParameterCode.AllowManualvehicleOnDairyBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.AllowSeprateSchemeItemPrintDairySaleInvoice, clsFixedParameterCode.AllowSeprateSchemeItemPrintDairySaleInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, clsFixedParameterType.ShowShipToPartyInDairyDispatch, clsFixedParameterCode.ShowShipToPartyInDairyDispatch, EnumControlType.CheckBox)
        '' Product Sale Return
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, clsFixedParameterType.AllowGateReturn, clsFixedParameterCode.AllowGateReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        '' Sale Return Gate Entry Product Sale
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryPS, clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, EnumControlType.CheckBox)
        '' Purchase Order
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, clsFixedParameterType.ShowItemInCaseofNonInventory, clsFixedParameterCode.ShowItemInCaseofNonInventory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionFATRateTollerance, clsFixedParameterCode.ProductionFATRateTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProductionSNFRateTollerance, clsFixedParameterCode.ProductionSNFRateTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ProductionFATRateTollerance, clsFixedParameterCode.ProductionFATRateTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ProductionSNFRateTollerance, clsFixedParameterCode.ProductionSNFRateTollerance, EnumControlType.NumericBox)
        'Merchant Trade
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.AllowPurchaseModulewithUniqueItem, clsFixedParameterCode.AllowPurchaseModulewithUniqueItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.ShowCancelButtonPO, clsFixedParameterCode.ShowCancelButtonPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.AllowBackDateEntry, clsFixedParameterCode.AllowBackDateEntry, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.ShortCloseItemWiseOnPO, clsFixedParameterCode.ShortCloseItemWiseOnPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.MakeClosingofPOreadonlyforuser, clsFixedParameterCode.MakeClosingofPOreadonlyforuser, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.PurchaseCounterOnTransactionType, clsFixedParameterCode.PurchaseCounterOnTransactionType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.RequiredPOLimit, clsFixedParameterCode.RequiredPOLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.UDLPurchaseOrderthroughAP, clsFixedParameterCode.UDLPurchaseOrderthroughAP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.PurchaseModulePickFixTaxRate, clsFixedParameterCode.PurchaseModulePickFixTaxRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.GSTExemptedAmountForNonRegisteredVendor, clsFixedParameterCode.GSTExemptedAmountForNonRegisteredVendor, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.FreeIndentQtyAfterPOClose, clsFixedParameterCode.FreeIndentQtyAfterPOClose, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.DoNotAllowSavePOWhenQtyNRateZero, clsFixedParameterCode.DoNotAllowSavePOWhenQtyNRateZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.IsRGPAfterPurchaseOrder, clsFixedParameterCode.IsRGPAfterPurchaseOrder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.IsRemarkReasonMandatoryOnPO, clsFixedParameterCode.IsRemarkReasonMandatoryOnPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.PurchasePickItemFromVendorItemDetails, clsFixedParameterCode.PurchasePickItemFromVendorItemDetails, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, clsFixedParameterType.ShowItemInCaseofNonInventory, clsFixedParameterCode.ShowItemInCaseofNonInventory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderMT, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProformaInvoiceMT, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDocumentAcceptance, clsFixedParameterType.AllowAutoMRNGRNonDocumentAcceptance, clsFixedParameterCode.AllowAutoMRNGRNonDocumentAcceptance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDocumentAcceptance, clsFixedParameterType.SkipMRNGRNinCaseofMT, clsFixedParameterCode.SkipMRNGRNinCaseofMT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDocumentAcceptance, clsFixedParameterType.ShowMRN, clsFixedParameterCode.ShowMRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDocumentAcceptance, clsFixedParameterType.ShowGRN, clsFixedParameterCode.ShowGRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmFixedDeposit, clsFixedParameterType.AllowToUseSubAccount, clsFixedParameterCode.AllowToUseSubAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCommercialInvoiceMT, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCommercialInvoiceMT, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesReturnMT, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesReturnMT, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, clsFixedParameterType.CHADetailsMandatoryOnExportSale, clsFixedParameterCode.CHADetailsMandatoryOnExportSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, clsFixedParameterType.EnableGSTRelatedfields, clsFixedParameterCode.EnableGSTRelatedfields, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.SkipMRNGRNinCaseofMT, clsFixedParameterCode.SkipMRNGRNinCaseofMT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.AutoGenerateMRN, clsFixedParameterCode.AutoGenerateMRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.ShowMRN, clsFixedParameterCode.ShowMRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.ShowGRN, clsFixedParameterCode.ShowGRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AutoDepartmentMendatroryFieldOnPurcahseCycle, clsFixedParameterCode.AutoDepartmentMendatroryFieldOnPurcahseCycle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AllowRoadPermitNo, clsFixedParameterCode.AllowRoadPermitNo, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AllowLargerItemCostThenVendorItemCost, clsFixedParameterCode.AllowLargerItemCostThenVendorItemCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.IsRateEditableOnSRN, clsFixedParameterCode.IsRateEditableOnSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AutoPOAtSRN, clsFixedParameterCode.AutoPOAtSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.EnableRackBin, clsFixedParameterCode.EnableRackBin, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.PurchaseCounterOnTransactionType, clsFixedParameterCode.PurchaseCounterOnTransactionType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.OpenPOforRejectShortageQty, clsFixedParameterCode.OpenPOforRejectShortageQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.PurchaseModulePickFixTaxRate, clsFixedParameterCode.PurchaseModulePickFixTaxRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.IsQCColumnRequiredonMRN, clsFixedParameterCode.IsQCColumnRequiredonMRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AllowPurchaseModulewithUniqueItem, clsFixedParameterCode.AllowPurchaseModulewithUniqueItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.ShowOptionforSelectingCapex, clsFixedParameterCode.ShowOptionforSelectingCapex, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.PurchasePickItemFromVendorItemDetails, clsFixedParameterCode.PurchasePickItemFromVendorItemDetails, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.IsRGPAfterPurchaseOrder, clsFixedParameterCode.IsRGPAfterPurchaseOrder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AllowPOScheduling, clsFixedParameterCode.AllowPOScheduling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.IsNotIncludeWasteQtyInCal, clsFixedParameterCode.IsNotIncludeWasteQtyInCal, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.IsShortageIncludeInLandedCost, clsFixedParameterCode.IsShortageIncludeInLandedCost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.DisableShipToLocation, clsFixedParameterCode.DisableShipToLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.AutoClosePO, clsFixedParameterCode.AutoClosePO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.MCCPurchase, clsFixedParameterCode.MCCPurchase, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.InvoiceBasedPO, clsFixedParameterCode.InvoiceBasedPO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.ShowQtySum_in_GRN_MRN_SRN, clsFixedParameterCode.ShowQtySum_in_GRN_MRN_SRN, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, clsFixedParameterType.SRNReportQuantityWise, clsFixedParameterCode.SRNReportQuantityWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MaxFATPerLimit, clsFixedParameterCode.MaxFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MaxSNFPerLimit, clsFixedParameterCode.MaxSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MinFATPerLimit, clsFixedParameterCode.MinFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MinSNFPerLimit, clsFixedParameterCode.MinSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MaxFATPerLimit, clsFixedParameterCode.MaxFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MaxSNFPerLimit, clsFixedParameterCode.MaxSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MinFATPerLimit, clsFixedParameterCode.MinFATPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MinSNFPerLimit, clsFixedParameterCode.MinSNFPerLimit, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MaxFATPerLimit, clsFixedParameterCode.MaxFATPerLimitOwnDCS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, clsFixedParameterType.MaxSNFPerLimit, clsFixedParameterCode.MaxSNFPerLimitOwnDCS, EnumControlType.NumericBox)
        'Jobwork Inward
        InsertDefaultValue(clsUserMgtCode.frmJobWorkBillig, clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkBillig, clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGeneralWeighment, clsFixedParameterType.GrossWeightChangePWD, clsFixedParameterCode.GrossWeightChangePWD, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmGeneralWeighment, clsFixedParameterType.MachineIntegrationInGeneralWeighment, clsFixedParameterCode.MachineIntegrationInGeneralWeighment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGeneralWeighment, clsFixedParameterType.FillGeneralWeighmentDetailsByJobworkTypeGateInNo, clsFixedParameterCode.FillGeneralWeighmentDetailsByJobworkTypeGateInNo, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.GateOutTankerNoAfterGeneralWeighment, clsFixedParameterCode.GateOutTankerNoAfterGeneralWeighment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.GateOutTankerNoAfterGeneralWeighment, clsFixedParameterCode.GateOutTankerNoAfterGeneralWeighment, EnumControlType.CheckBox)
        'JobWork OutWard
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.Allow_AmountTruncate_BulkMilkSRN, clsFixedParameterCode.Allow_AmountTruncate_BulkMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.AllowBulkProcItemPostedData, clsFixedParameterCode.AllowBulkProcItemPostedData, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry_JWO, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowGateEntryInPrevDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry_JWO, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment_JWO, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowWeighmentDateAfterCurrentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment_JWO, clsFixedParameterType.AllowToSaveTimeWithDocumentDate, clsFixedParameterCode.AllowToSaveTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQC_JWO, clsFixedParameterType.defaultCorrectionFactor, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmQC_JWO, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowQcDateAfterCurrentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQC_JWO, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowQcDateBeforeGateEntryDateTime, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmUnloading_JWO, clsFixedParameterType.MilkProc, clsFixedParameterCode.AllowUnloadingDateAfterCurrentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.JWO_SRN, clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.JWO_SRN_Return, clsFixedParameterType.CancelDocDueToSRNReturn, clsFixedParameterCode.CancelDocDueToSRNReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.AllowToShowMilkTypeinAdjustmentEntry, clsFixedParameterCode.AllowToShowMilkTypeinAdjustmentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.FatSnfWhenMilktypeSelect, clsFixedParameterCode.FatSnfWhenMilktypeSelect, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.AllowStockToleranceNegative, clsFixedParameterCode.AllowStockToleranceNegative, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.GateEntryTankerFromTankerMaster, clsFixedParameterCode.GateEntryTankerFromTankerMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerDeduction, clsFixedParameterType.NoOFCustomerForImportExport, clsFixedParameterCode.NoOFCustomerForImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMPMaster, clsFixedParameterType.NoOFIncentiveForMPImportExport, clsFixedParameterCode.NoOFIncentiveForMPImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.SaleIncentiveMaster, clsFixedParameterType.NoOFSlabForImportExport, clsFixedParameterCode.NoOFSlabForImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmIncomeTaxSlab, clsFixedParameterType.NoOFSlabForImportExport, clsFixedParameterCode.NoOFSlabForImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.SaleIncentiveMaster, clsFixedParameterType.NoOFItemStructureForImportExport, clsFixedParameterCode.NoOFItemStructureForImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.SaleIncentiveMaster, clsFixedParameterType.NoOFCustomerForImportExport, clsFixedParameterCode.NoOFCustomerForImportExport, EnumControlType.NumericBox)
        'NoOFSavingCodeForImportExport
        InsertDefaultValue(clsUserMgtCode.routeMaster, clsFixedParameterType.NoOFCustomerForImportExportOnRouteMaster, clsFixedParameterCode.NoOFCustomerForImportExportOnRouteMaster, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmRouteMasterDS, clsFixedParameterType.NoOFCustomerForImportExportOnRouteMaster, clsFixedParameterCode.NoOFCustomerForImportExportOnRouteMaster, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmRouteMaster, clsFixedParameterType.NoOFCustomerForImportExportOnRouteMaster, clsFixedParameterCode.NoOFCustomerForImportExportOnRouteMaster, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeSavingsMapping, clsFixedParameterType.NoOFSavingCodeForImportExport, clsFixedParameterCode.NoOFSavingCodeForImportExport, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.CheckUnpostedPaymentProcess, clsFixedParameterCode.CheckUnpostedPaymentProcess, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptRouteWiseSaleRegister, clsFixedParameterType.CrateToLTR, clsFixedParameterCode.CrateToLTR, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.RptRouteWiseSaleRegister, clsFixedParameterType.CanToLTR, clsFixedParameterCode.CanToLTR, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ProdcutionDoNotCheckForwardDocuments, clsFixedParameterCode.ProdcutionDoNotCheckForwardDocuments, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.DoNotStopOnItemBalanceExceptionStoreAdjustment, clsFixedParameterCode.DoNotStopOnItemBalanceExceptionStoreAdjustment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.ProductionIssueQtyTollerance, clsFixedParameterCode.ProductionIssueQtyTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleVsReceipReport, clsFixedParameterType.NopreviousDaysInSaleVSReceipt, clsFixedParameterCode.NopreviousDaysInSaleVSReceipt, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProfitAndLossPerforma, clsFixedParameterType.SelectGLInProftAndLossPerforma, clsFixedParameterCode.SelectGLInProftAndLossPerforma, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBalanceSheetPerforma, clsFixedParameterType.SelectGLInBalanceSheetPerforma, clsFixedParameterCode.SelectGLInBalanceSheetPerforma, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBalanceSheetPerforma, clsFixedParameterType.BalanceSheetPerformaWithFormula, clsFixedParameterCode.BalanceSheetPerformaWithFormula, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.CashFlowPerforma, clsFixedParameterType.SelectGLInCashFlowPerforma, clsFixedParameterCode.SelectGLInCashFlowPerforma, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.BulkProcurementApplyTotalSoidRate, clsFixedParameterCode.BulkProcurementApplyTotalSoidRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.CalculateLtrQtyFromKGQtyByCLR, clsFixedParameterCode.CalculateLtrQtyFromKGQtyByCLR, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.ApplyCalculateWeightInLtr, clsFixedParameterCode.ApplyCalculateWeightInLtr, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.ApplyTransportChargeAddInActualAmount, clsFixedParameterCode.ApplyTransportChargeAddInActualAmount, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmGatePassPS, clsFixedParameterType.CalculateProvisionOnGateePass, clsFixedParameterCode.CalculateProvisionOnGateePass, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.CalculateProvisionOnGateePass, clsFixedParameterCode.CalculateProvisionOnGateePass, EnumControlType.CheckBox)
        '****************************************************************************************************************
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.AllowDistributorSaleAtCSA_SaleInvoice, clsFixedParameterCode.AllowDistributorSaleAtCSA_SaleInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentEntry, clsFixedParameterType.AddressOnPaymentVoucher, clsFixedParameterCode.AddressOnPaymentVoucherOnBankBasis, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCompanyMaster, clsFixedParameterType.AllowMasterModificationWithSecurity, clsFixedParameterCode.AllowMasterModificationWithSecurity, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AllowSMSSendtoSalePerson, clsFixedParameterCode.AllowSMSSendtoSalePerson, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSilageQualityCheck, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSilageWeighment, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleReturn, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.AllowToPrintTimeWithDocumentDate, clsFixedParameterCode.AllowToPrintTimeWithDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowAutoNoForBackLogEntry, clsFixedParameterCode.AllowAutoNoForBackLogEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPayableSettings, clsFixedParameterType.AllowBankDetailsManualinVM, clsFixedParameterCode.AllowBankDetailsManualinVM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleSettings, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowDeliveryOrderIncaseAmountIncreases, clsFixedParameterCode.AllowDeliveryOrderIncaseAmountIncreases, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryNoteFreshSale, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryOrderDairy, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowDeliveryQtygreaterthanBookingQtyPS, clsFixedParameterCode.AllowDeliveryQtygreaterthanBookingQtyPS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowDiffentSeriesExemptedItemONPS, clsFixedParameterCode.AllowDiffentSeriesExemptedItemONPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleSettings, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowDispatchOutstandingBS, clsFixedParameterCode.AllowDispatchOutstandingBS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleSettings, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanReceived, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBookingEntry, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCreateReceived, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryNoteFreshSale, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingDairyMultipleDistributor, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.TransferCrateReceived, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryOrderDairy, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPOSBookingDairyMultipleDistributor, clsFixedParameterType.AllowDispatchOutstandingFS, clsFixedParameterCode.AllowDispatchOutstandingFS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AllowDispatchOutstandingPS, clsFixedParameterCode.AllowDispatchOutstandingPS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmPayableSettings, clsFixedParameterType.AllowFieldsToBeManadatory, clsFixedParameterCode.AllowFieldsToBeManadatory, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.FrmPayableSettings, clsFixedParameterType.AllowFieldsToBeManadatory, clsFixedParameterCode.AllowFieldsToBeManadatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vhicleMaster, clsFixedParameterType.AllowFieldsToBeManadatory, clsFixedParameterCode.AllowFieldsToBeManadatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.inventorySetting, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.AllowNegativeStock, clsFixedParameterCode.AllowNegativeStock, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.saleReturn, clsFixedParameterType.AllowNegtiveOfSaleInvoiceBlanceAmt, clsFixedParameterCode.AllowNegtiveOfSaleInvoiceBlanceAmt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.AllowReceiptThroughSO, clsFixedParameterCode.AllowReceiptThroughSO, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPendingBooking, clsFixedParameterType.AllowSingleInvoiceAgainstSingleOrder, clsFixedParameterCode.AllowSingleInvoiceAgainstSingleOrder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleSettings, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmCreateAutoInvoiceBS, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, clsFixedParameterType.AmountLimitForInvoiceBulkSale, clsFixedParameterCode.AmountLimitForInvoiceBulkSale, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBalanceSheetPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupCode, clsFixedParameterCode.BalanceSheetProftAndLossGroupCode, EnumControlType.NumericBox)
        'InsertDefaultValue(clsUserMgtCode.CashFlowPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupCode, clsFixedParameterCode.BalanceSheetProftAndLossGroupCode, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmProfitAndLossPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupCode, clsFixedParameterCode.BalanceSheetProftAndLossGroupCode, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.rptBalanceSheet, clsFixedParameterType.BalanceSheetProftAndLossGroupCode, clsFixedParameterCode.BalanceSheetProftAndLossGroupCode, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBalanceSheetPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupDesc, clsFixedParameterCode.BalanceSheetProftAndLossGroupDesc, EnumControlType.TextBox)
        'InsertDefaultValue(clsUserMgtCode.CashFlowPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupDesc, clsFixedParameterCode.BalanceSheetProftAndLossGroupDesc, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmProfitAndLossPerforma, clsFixedParameterType.BalanceSheetProftAndLossGroupDesc, clsFixedParameterCode.BalanceSheetProftAndLossGroupDesc, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.rptBalanceSheet, clsFixedParameterType.BalanceSheetProftAndLossGroupDesc, clsFixedParameterCode.BalanceSheetProftAndLossGroupDesc, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.MultiplySubsidyWithQuantity, clsFixedParameterCode.MultiplySubsidyWithQuantity, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.CreateAutoReceiptEntryDCSSale, clsFixedParameterCode.CreateAutoReceiptEntryDCSSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.GeneratePDFForMobile, clsFixedParameterCode.GeneratePDFForMobile, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.CreateAutoReceiptEntryDCSSale, clsFixedParameterCode.CreateAutoReceiptEntryDCSSale, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.GeneratePDFForMobile, clsFixedParameterCode.GeneratePDFForMobile, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderMT, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDispatchMultipleFreshSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchFreshSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductBooking, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNSalesOrder, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialFarmer, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyFreshDispatchMultiple, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AutoSchemeOn, clsFixedParameterCode.AutoSchemeOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.AssetGroupPrefix, clsFixedParameterCode.AssetGroupPrefix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer, clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, clsFixedParameterType.ApplyCostingOnPostedDate, clsFixedParameterCode.ApplyCostingOnPostedDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBankReco, clsFixedParameterType.BankRecoHidePWD, clsFixedParameterCode.BankRecoHidePWD, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, clsFixedParameterType.BulkInvoiceDeleteType, clsFixedParameterCode.BulkInvoiceDelete, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmApprovalLevelScreen, clsFixedParameterType.WorkApprovalFlowInERP, clsFixedParameterCode.WorkApprovalFlowInERP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCreateReceived, clsFixedParameterType.UpdateCrateLinerQty, clsFixedParameterCode.UpdateCrateLinerQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, clsFixedParameterType.TreatExcessLeaveAbsent, clsFixedParameterCode.TreatExcessLeaveAbsent, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickSettlement, clsFixedParameterType.TransferTransTypeRouteHide, clsFixedParameterCode.TransferTransTypeRouteHide, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSubsidyCreditNote, clsFixedParameterType.PeriodofSubsidyCreditNote, clsFixedParameterCode.PeriodofSubsidyCreditNote, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.StopGLEntryForConsignmentAtCSATransfer, clsFixedParameterCode.StopGLEntryForConsignmentAtCSATransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMinMix, clsFixedParameterCode.SNFMinMix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMinMix, clsFixedParameterCode.SNFMinMix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMinCow, clsFixedParameterCode.SNFMinCow, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMinCow, clsFixedParameterCode.SNFMinCow, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMinBuff, clsFixedParameterCode.SNFMinBuff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMinBuff, clsFixedParameterCode.SNFMinBuff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMaxMix, clsFixedParameterCode.SNFMaxMix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMaxMix, clsFixedParameterCode.SNFMaxMix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMaxCow, clsFixedParameterCode.SNFMaxCow, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMaxCow, clsFixedParameterCode.SNFMaxCow, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, clsFixedParameterType.SNFMaxBuff, clsFixedParameterCode.SNFMaxBuff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.SNFMaxBuff, clsFixedParameterCode.SNFMaxBuff, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCommercialInvoiceMT, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXPorformaInvoice, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesOrder, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesQuotation, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesReturn, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnFreshSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNSaleInvoice, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNSaleReturn, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNSalesOrder, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleQuotation, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNServiceInvoice, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoicedairy, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialFarmer, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturnFarmer, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturndairy, clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.CalculateSNFFromCLRForMCCMilk, clsFixedParameterCode.CalculateSNFFromCLRForMCCMilk, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, clsFixedParameterCode.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNJobWorkEstimate, clsFixedParameterType.SNFFromCLRAndCorrectionFactorInJWIEst, clsFixedParameterCode.SNFFromCLRAndCorrectionFactorInJWIEst, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNJobWorkEstimate, clsFixedParameterType.JOBdefaultCorrectionFactorBS, clsFixedParameterCode.JOBdefaultCorrectionFactorBS, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.FrmSRNJobWorkEstimate, clsFixedParameterType.AutoCalculateProduceQty, clsFixedParameterCode.AutoCalculateProduceQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, clsFixedParameterType.SyncedMccToServerStartDateForEmailSms, clsFixedParameterCode.SyncedMccToServerStartDateForEmailSms, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.CrateReceiveddairyCustomerWise, clsFixedParameterCode.CrateReceiveddairyCustomerWise, EnumControlType.CheckBox)
        'clsUserMgtCode.     Sale_SMSATPOST
        '****************************************************************************************************************
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, clsFixedParameterType.MaintainLogForImporperSample, clsFixedParameterCode.MaintainLogForImporperSample, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, clsFixedParameterType.PFCalculationOnFormulaHead, clsFixedParameterCode.PFCalculationOnFormulaHead, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, clsFixedParameterType.MODValueForFAT, clsFixedParameterCode.MODValueForFAT, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, clsFixedParameterType.AutoFillSameLocationInGrid, clsFixedParameterCode.AutoFillSameLocationInGrid, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkBillig, clsFixedParameterType.JWIRateofFGasPerRM, clsFixedParameterCode.JWIRateofFGasPerRM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.TagMultipleRouteWithCustomer, clsFixedParameterCode.TagMultipleRouteWithCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptTankerStatusReport, clsFixedParameterType.ApplyNotShowJobWorkTypeTanker, clsFixedParameterCode.ApplyNotShowJobWorkTypeTanker, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.ImportMultipleAssetAssembled, clsFixedParameterCode.ImportMultipleAssetAssembled, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnSRN, clsFixedParameterType.PurchaseDoNotCheckForwardDocuments, clsFixedParameterCode.PurchaseDoNotCheckForwardDocuments, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.CreateProvisionOnOpeningAndClosingKM, clsFixedParameterCode.CreateProvisionOnOpeningAndClosingKM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.FromLocationStockNotCheckConsumptionLocation, clsFixedParameterCode.FromLocationStockNotCheckConsumptionLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.IncentiveEntry, clsFixedParameterType.MilkIncetiveByMilkSRN, clsFixedParameterCode.MilkIncetiveByMilkSRN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.AllowShowCoumnInCrateReceivedDairy, clsFixedParameterCode.AllowShowCoumnInCrateReceivedDairy, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowShelfLifeMandatoryOnFG, clsFixedParameterCode.AllowShelfLifeMandatoryOnFG, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, clsFixedParameterType.RefundknockoffwithCreditNote, clsFixedParameterCode.RefundknockoffwithCreditNote, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, clsFixedParameterType.CreateInvoiceAutomaticallyOnbulkDispatch, clsFixedParameterCode.CreateInvoiceAutomaticallyOnbulkDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.CrateReceivingWithMultipleRoute, clsFixedParameterCode.CrateReceivingWithMultipleRoute, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ShowMulMRPOfSameItemOnDairyBookingCustomer, clsFixedParameterCode.ShowMulMRPOfSameItemOnDairyBookingCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.AllowZeroQtyOnDairyBooking, clsFixedParameterCode.AllowZeroQtyOnDairyBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.UseCutOffTimeonRouteForERP, clsFixedParameterCode.UseCutOffTimeonRouteForERP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingUploader, clsFixedParameterType.AllowZeroQtyOnDairyBookingUploader, clsFixedParameterCode.AllowZeroQtyOnDairyBookingUploader, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.DonotIncludeSecurityInCustomerOutstanding, clsFixedParameterCode.DonotIncludeSecurityInCustomerOutstanding, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.IsManualTCS, clsFixedParameterCode.IsManualTCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.IsManualTCS, clsFixedParameterCode.IsManualTCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSNSaleReturn, clsFixedParameterType.IsManualTCS, clsFixedParameterCode.IsManualTCS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, clsFixedParameterType.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, clsFixedParameterCode.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmbookingdairyFreshSale, clsFixedParameterType.CheckNoOfDaysforCardSaleBooking, clsFixedParameterCode.CheckNoOfDaysforCardSaleBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ShowBookingTypeDropDownonDairyBookingCustomer, clsFixedParameterCode.ShowBookingTypeDropDownonDairyBookingCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.ShowFATSNFPerOnBulkProcInGateIN, clsFixedParameterCode.ShowFATSNFPerOnBulkProcInGateIN, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowItemCostMandatoryForStockingUnit, clsFixedParameterCode.AllowItemCostMandatoryForStockingUnit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.PickCostOFMaterialSaleFromPriceMaster, clsFixedParameterCode.PickCostOFMaterialSaleFromPriceMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptActiveUsers, clsFixedParameterType.Idle, clsFixedParameterCode.isIdleTimerOn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptActiveUsers, clsFixedParameterType.Idle, clsFixedParameterCode.idleTime, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.rptActiveUsers, clsFixedParameterType.SameuserCanNotloginmultipletimes, clsFixedParameterCode.SameuserCanNotloginmultipletimes, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.DairyBookingTolleranceQty, clsFixedParameterCode.DairyBookingTolleranceQty, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, clsFixedParameterType.VehicleCodeNotMandatoryOnLoadINSlip, clsFixedParameterCode.VehicleCodeNotMandatoryOnLoadINSlip, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVLCMaster, clsFixedParameterType.ApplyLatestPriceChartWhilecreatingNewVSP, clsFixedParameterCode.ApplyLatestPriceChartWhilecreatingNewVSP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.DisplayTypeInMilkReceipt, clsFixedParameterCode.DisplayTypeInMilkReceipt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.MCCDefaultMilkItemCow, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, clsFixedParameterType.MCCDefaultMilkItemBuffalo, clsFixedParameterCode.MilkSetting, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.EnableBankFromMaster, clsFixedParameterCode.EnableBankFromMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.Autoadjustdebitcreditnote, clsFixedParameterCode.Autoadjustdebitcreditnote, EnumControlType.CheckBox)

        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.UpdateItemMasterWithoutTransactionValidation, clsFixedParameterCode.UpdateItemMasterWithoutTransactionValidation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoicedairy, clsFixedParameterType.AddItemAliasInSMS, clsFixedParameterCode.AddItemAliasInSMS, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.ItemWiseQualityCheckInGeneralPurchase, clsFixedParameterCode.ItemWiseQualityCheckInGeneralPurchase, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.ShowMCCFinderInPaymentProcess, clsFixedParameterCode.ShowMCCFinderInPaymentProcess, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.LastMilkReceiptQtyTollerance, clsFixedParameterCode.LastMilkReceiptQtyTollerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, clsFixedParameterType.DayWiseCustomerIncentiveCalculation, clsFixedParameterCode.DayWiseCustomerIncentiveCalculation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMultipleProcDeduction, clsFixedParameterType.RepeatDeductionAndVendor, clsFixedParameterCode.RepeatDeductionAndVendor, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, clsFixedParameterType.CustomerIncetiveAutoSecuity, clsFixedParameterCode.CustomerIncetiveAutoSecuity, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, clsFixedParameterType.CustomerIncetiveBankForSecuity, clsFixedParameterCode.CustomerIncetiveBankForSecuity, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, clsFixedParameterType.CustomerIncetivePaymentModeForSecuity, clsFixedParameterCode.CustomerIncetivePaymentModeForSecuity, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.SelectMilkRejectDefaulterManually, clsFixedParameterCode.SelectMilkRejectDefaulterManually, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, clsFixedParameterType.MilkProcurementSNF2DecimalPlaces, clsFixedParameterCode.MilkProcurementSNF2DecimalPlaces, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.MultipleMCCFinder, clsFixedParameterCode.MultipleMCCFinder, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.POWeighment, clsFixedParameterType.AllowtoEnterNetWeightManuallyinPOWeighmentScreen, clsFixedParameterCode.AllowtoEnterNetWeightManuallyinPOWeighmentScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.TIPRateMix, clsFixedParameterCode.TIPRateMix, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.TIPRateCow, clsFixedParameterCode.TIPRateCow, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.TIPRateBuffalo, clsFixedParameterCode.TIPRateBuffalo, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.DefaultCustomerGroupCodeforVSP, clsFixedParameterCode.DefaultCustomerGroupCodeforVSP, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, clsFixedParameterType.DefaultVendorGroupCodeforVSP, clsFixedParameterCode.DefaultVendorGroupCodeforVSP, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmTankerMaster, clsFixedParameterType.AllowSameTankerNoforPrimarySecondaryTransporter, clsFixedParameterCode.AllowSameTankerNoforPrimarySecondaryTransporter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TankerDispatchAvgFATSNFPer, clsFixedParameterCode.TankerDispatchAvgFATSNFPer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.TankerDispatchAvgFATSNFPer, clsFixedParameterCode.TankerDispatchAvgFATSNFPer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, clsFixedParameterType.DisableToPickMainLocationType, clsFixedParameterCode.DisableToPickMainLocationType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.AllocateToMandatoryonGateOut, clsFixedParameterCode.AllocateToMandatoryonGateOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptDairyTruckSheetReport, clsFixedParameterType.PrintTruckSheetAfterGenerate, clsFixedParameterCode.PrintTruckSheetAfterGenerate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptVillageDiffReport, clsFixedParameterType.VillageDiffrenceReportValueWithTwoDecimalPoint, clsFixedParameterCode.VillageDiffrenceReportValueWithTwoDecimalPoint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, clsFixedParameterType.AskForPwdForOutAdjustmentOnPost, clsFixedParameterCode.AskForPwdForOutAdjustmentOnPost, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmBankReco, clsFixedParameterType.BankRecoCheckFutureDocuments, clsFixedParameterCode.BankRecoCheckFutureDocuments, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.StockRecoReport, clsFixedParameterType.StockRecoRateTunning, clsFixedParameterCode.StockRecoRateTunning, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.allowMilkJWOutowordWithAvgFatSNFRate, clsFixedParameterCode.allowMilkJWOutowordWithAvgFatSNFRate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, clsFixedParameterType.allowMilkJWOutowordWithAvgFatSNFPerAtInventory, clsFixedParameterCode.allowMilkJWOutowordWithAvgFatSNFPerAtInventory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.CreateProvisionOfTankerDispatchWithClosingKM, clsFixedParameterCode.CreateProvisionOfTankerDispatchWithClosingKM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, clsFixedParameterType.MaterialSaleInvoiceEnablePrintOnPost, clsFixedParameterCode.MaterialSaleInvoiceEnablePrintOnPost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TollTaxMaster, clsFixedParameterCode.TollTaxMaster, EnumControlType.CheckBox)
        ' TollTaxMaster
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.InDocMandatoryOnInternalTransfer, clsFixedParameterCode.InDocMandatoryOnInternalTransfer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.Transfer, clsFixedParameterType.AllowTransferInAfterGatePassOnly, clsFixedParameterCode.AllowTransferInAfterGatePassOnly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.EnableItemShortDescriptionInBooking, clsFixedParameterCode.EnableItemShortDescriptionInBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.AllowDuplicateItemShortDescriptionInItemMaster, clsFixedParameterCode.AllowDuplicateItemShortDescriptionInItemMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, clsFixedParameterType.DoNotCreatePaymentWhileSalaryGeneration, clsFixedParameterCode.DoNotCreatePaymentWhileSalaryGeneration, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerGateOut, clsFixedParameterType.MccPlantSelectionOptionInMccTankerGateOut, clsFixedParameterCode.MccPlantSelectionOptionInMccTankerGateOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.EnableTankerNoInMccTankerDispWithMccTankerGateOut, clsFixedParameterCode.EnableTankerNoInMccTankerDispWithMccTankerGateOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.AllowToEnterSnfAtPlantInMccTankerDispatch, clsFixedParameterCode.AllowToEnterSnfAtPlantInMccTankerDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, clsFixedParameterType.AllowtoSelDateandBankforPayEntryOnSalaryGeneration, clsFixedParameterCode.AllowtoSelDateandBankforPayEntryOnSalaryGeneration, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.BennyImportAutoCreateMP, clsFixedParameterCode.BennyImportAutoCreateMP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, clsFixedParameterType.BennyImportPickRateFromPrice, clsFixedParameterCode.BennyImportPickRateFromPrice, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmGateOut, clsFixedParameterType.AllocatedTankerGateOut, clsFixedParameterCode.AllocatedTankerGateOut, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.AllowToCreateNoOfBookingPerDay, clsFixedParameterCode.AllowToCreateNoOfBookingPerDay, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, clsFixedParameterType.EnterWeightManuallyOnWeighmentInGrid, clsFixedParameterCode.EnterWeightManuallyOnWeighmentInGrid, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, clsFixedParameterType.AllowSameChequeNoForMultiplePaymentEntry, clsFixedParameterCode.AllowGenerateReferenceNoForBulkGateEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, clsFixedParameterType.AllowGenerateReferenceNoForBulkGateEntry, clsFixedParameterCode.AllowGenerateReferenceNoForBulkGateEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.bankTransfer, clsFixedParameterType.AllowBankTransferAgainstMilkBill, clsFixedParameterCode.AllowBankTransferAgainstMilkBill, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptKDILSalarySlip, clsFixedParameterType.ChangeLeaveDescriptionOnSalarySlip, clsFixedParameterCode.ChangeLeaveDescriptionOnSalarySlip, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, clsFixedParameterType.ApplyFinancialCostCenter, clsFixedParameterCode.ApplyFinancialCostCenter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptKDILSalarySlip, clsFixedParameterType.SalarySlipLeaveStatusOnTheBasisOfCalendarYear, clsFixedParameterCode.SalarySlipLeaveStatusOnTheBasisOfCalendarYear, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptPrimaryTransporter, clsFixedParameterType.ApplyCalculationOnRouteLenth, clsFixedParameterCode.ApplyCalculationOnRouteLenth, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, clsFixedParameterType.EnableExportExcelOnIncentiveEntry, clsFixedParameterCode.EnableExportExcelOnIncentiveEntry, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaster, clsFixedParameterType.AllowBankSectionEnableOnMCCMaster, clsFixedParameterCode.AllowBankSectionEnableOnMCCMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MCCProvisonReport, clsFixedParameterType.AllowRoundDownAmtForMCCDateWiseChilling, clsFixedParameterCode.AllowRoundDownAmtForMCCDateWiseChilling, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPAssetIssue, clsFixedParameterType.CtreateJEOfVspAssetIssueAndReturn, clsFixedParameterCode.CtreateJEOfVspAssetIssueAndReturn, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPAssetIssue, clsFixedParameterType.createDebitNoteOnAssetIssue, clsFixedParameterCode.createDebitNoteOnAssetIssue, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSPAssetIssue, clsFixedParameterType.PickAvgCostonAssetissue, clsFixedParameterCode.PickAvgCostonAssetissue, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptDairyTruckSheetReport, clsFixedParameterType.ApplyTCSAmtOnAbstractReportDotMatrix, clsFixedParameterCode.ApplyTCSAmtOnAbstractReportDotMatrix, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterMaster, clsFixedParameterType.NotAllowDuplicatePANOnPrimaryTransporter, clsFixedParameterCode.NotAllowDuplicatePANOnPrimaryTransporter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptDairyTruckSheetReport, clsFixedParameterType.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, clsFixedParameterCode.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptDairyTruckSheetReport, clsFixedParameterType.ShowEarlyRouteOnTruckSheet, clsFixedParameterCode.ShowEarlyRouteOnTruckSheet, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, clsFixedParameterType.AllowOutEntryOnCrateReceivedDairyForAdjustment, clsFixedParameterCode.AllowOutEntryOnCrateReceivedDairyForAdjustment, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmCanSaleUploader, clsFixedParameterType.CanSaleAvgFATSNFPer, clsFixedParameterCode.CanSaleAvgFATSNFPer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.BulkMilkFATSNFKGDecimalPlaces, clsFixedParameterCode.BulkMilkFATSNFKGDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.BulkMilkFATSNFAmtDecimalPlaces, clsFixedParameterCode.BulkMilkFATSNFAmtDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, clsFixedParameterType.BulkMilkConsiderAllParametersForIncetive, clsFixedParameterCode.BulkMilkConsiderAllParametersForIncetive, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, clsFixedParameterType.TankerDispatchProvisionLocationSegment, clsFixedParameterCode.TankerDispatchProvisionLocationSegment, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.LocalSaleAllowedPer, clsFixedParameterCode.LocalSaleAllowedPer, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.LocalSaleAllowedRate, clsFixedParameterCode.LocalSaleAllowedRate, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmMilkRouteMaster, clsFixedParameterType.SeprateDistanceMorningEveningShift, clsFixedParameterCode.SeprateDistanceMorningEveningShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.BankCodeForApplyDocumentPaymentOFAssetLost, clsFixedParameterCode.BankCodeForApplyDocumentPaymentOFAssetLost, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.FrmPendingRequisitionQty, clsFixedParameterType.ShowStatusItemWiseInPendingRequisitionRpt, clsFixedParameterCode.ShowStatusItemWiseInPendingRequisitionRpt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.vendormaster, clsFixedParameterType.DoNotCheckAnyValidationOnVendorInactive, clsFixedParameterCode.DoNotCheckAnyValidationOnVendorInactive, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.UserWiseRouteMapping, clsFixedParameterCode.UserWiseRouteMapping, EnumControlType.CheckBox)
        'InsertDefaultValue(clsUserMgtCode.frmMCCTankerGateOut, clsFixedParameterType.CreateMCCTankerGateOutBasedOnBulkRouteMaster, clsFixedParameterCode.CreateMCCTankerGateOutBasedOnBulkRouteMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.CreateGatePassFromDemand, clsFixedParameterCode.CreateGatePassFromDemand, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.RptBulkMilkRegister, clsFixedParameterType.PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister, clsFixedParameterCode.PickFatSnfKGFromBulkMilkSRNInBulkMilkRegister, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmEmployee_Master, clsFixedParameterType.AadharNoMandatoryOnEmpMaster, clsFixedParameterCode.AadharNoMandatoryOnEmpMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.SuperUserCustomer, clsFixedParameterCode.SuperUserCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSP_VLCMaster, clsFixedParameterType.ApplyDefaultsInMaster, clsFixedParameterCode.ApplyDefaultsInMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMapPayHeadsToSalaStructure, clsFixedParameterType.UpdateMapPayHeadsToSalaStructurePassword, clsFixedParameterCode.UpdateMapPayHeadsToSalaStructurePassword, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmVSP_VLCMaster, clsFixedParameterType.Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster, clsFixedParameterCode.Allow_Reg_PDCS_CLUSTER_2ndBank_MCC_VLCVSPMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSP_VLCMaster, clsFixedParameterType.NewDCSScreen, clsFixedParameterCode.NewDCSScreen, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmVSP_VLCMaster, clsFixedParameterType.SameBankCodeForCurrentAndSaving, clsFixedParameterCode.SameBankCodeForCurrentAndSaving, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, clsFixedParameterType.MultipleFinderFillAuto, clsFixedParameterCode.MultipleFinderFillAuto, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningSTD, clsFixedParameterType.MandatoryLineNoMaxMinQtyForProductionPlan, clsFixedParameterCode.MandatoryLineNoMaxMinQtyForProductionPlan, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialCosting, clsFixedParameterType.RunProductionBaseOnPercentage, clsFixedParameterCode.RunProductionBaseOnPercentage, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, clsFixedParameterType.AutoClosePOBasedOnSRNQtyWithTolerance, clsFixedParameterCode.AutoClosePOBasedOnSRNQtyWithTolerance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.ApplyMilkTypeBuffaloCowOnPrint, clsFixedParameterCode.ApplyMilkTypeBuffaloCowOnPrint, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.ApplyZoneWiseVSP, clsFixedParameterCode.ApplyZoneWiseVSP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmProductionVarianceSTD, clsFixedParameterType.ApplyStandardProductionVariance, clsFixedParameterCode.ApplyStandardProductionVariance, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.ItemCostTolerancePercentage, clsFixedParameterCode.ItemCostTolerancePercentage, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, clsFixedParameterType.HeadLoadDescriptionInPaymentProcessPrint, clsFixedParameterCode.HeadLoadDescriptionInPaymentProcessPrint, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, clsFixedParameterType.AutoCreateSaleInvoice, clsFixedParameterCode.AutoCreateSaleInvoice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.PrefixForUserMaster, clsFixedParameterCode.PrefixForUserMaster, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandApproval, clsFixedParameterType.ApplyDemandApproval, clsFixedParameterCode.ApplyDemandApproval, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.CheckCreditLimit, clsFixedParameterCode.CheckCreditLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, clsFixedParameterType.ApplyTolerance, clsFixedParameterCode.ApplyTolerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmDemand_Sheet, clsFixedParameterType.SetShiftTimeOut, clsFixedParameterCode.SetShiftTimeOut, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.SetDefaultShiftTime, clsFixedParameterCode.SetDefaultShiftTime, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemand_Sheet, clsFixedParameterType.AllowToSelectShift, clsFixedParameterCode.AllowToSelectShift, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomersListReport, clsFixedParameterType.ApplyOrderByNumeric, clsFixedParameterCode.ApplyOrderByNumeric, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyRoundOffZero, clsFixedParameterCode.ApplyRoundOffZero, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmRouteMasterDS, clsFixedParameterType.EnableLocation, clsFixedParameterCode.EnableLocation, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.EnableResetDemand, clsFixedParameterCode.EnableResetDemand, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.IsLoadingSlipMandatory, clsFixedParameterCode.IsLoadingSlipMandatory, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptVSPMilkNotsold, clsFixedParameterType.PickAllBMC, clsFixedParameterCode.PickAllBMC, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptVSPMilkNotsold, clsFixedParameterType.ApplyMonthlyYearly, clsFixedParameterCode.ApplyMonthlyYearly, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDistributorRouteTagging, clsFixedParameterType.SeparateDemandMilkandProduct, clsFixedParameterCode.SeparateDemandMilkandProduct, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ShowDemandDoc, clsFixedParameterCode.ShowDemandDoc, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyCommission, clsFixedParameterCode.ApplyCommission, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyCommissionRateWithTax, clsFixedParameterCode.ApplyCommissionRateWithTax, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.DisableRouteandVehicle, clsFixedParameterCode.DisableRouteandVehicle, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.UpdateDemandBeforePost, clsFixedParameterCode.UpdateDemandBeforePost, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PricePlan, clsFixedParameterType.PricePlanRoundOffTruncate, clsFixedParameterCode.PricePlanRoundOffTruncate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PricePlan, clsFixedParameterType.ApplyPricePlanDigitRule, clsFixedParameterCode.ApplyPricePlanDigitRule, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.PricePlan, clsFixedParameterType.PricePlanRoundOffDigit, clsFixedParameterCode.PricePlanRoundOffDigit, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandAdjustment, clsFixedParameterType.DemandIncreaseDecreasePer, clsFixedParameterCode.DemandIncreaseDecreasePer, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandAdjustment, clsFixedParameterType.DemandIncreaseDecreaseQty, clsFixedParameterCode.DemandIncreaseDecreaseQty, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.AllowIncreaseDispatchQty, clsFixedParameterCode.AllowIncreaseDispatchQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.FORPRICE, clsFixedParameterCode.FORPRICE, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.AddPostFunctionality, EnumControlType.CheckBox)

        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.BMCStartTime, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.BMCEndHours, EnumControlType.TextBox)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, clsFixedParameterType.AndroidMilkCollectionBMCDCS, clsFixedParameterCode.DCSEndHours, EnumControlType.TextBox)


        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.AllowGatePassDemandTripWise, clsFixedParameterCode.AllowGatePassDemandTripWise, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.ConvertPouchtoCrate, clsFixedParameterCode.ConvertPouchtoCrate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.DontCreateForPouch, clsFixedParameterCode.DontCreateForPouch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ChangeInventroyMovemnet, clsFixedParameterCode.ChangeInventroyMovemnet, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, clsFixedParameterType.CreateAutoMCCPriceChat, clsFixedParameterCode.CreateAutoMCCPriceChat, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, clsFixedParameterType.VehicleNofromDispatch, clsFixedParameterCode.VehicleNofromDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DispatchPriceCodeForCreditCustomer, clsFixedParameterCode.DispatchPriceCodeForCreditCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyCommissionTPTForCreditCustomer, clsFixedParameterCode.ApplyCommissionTPTForCreditCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ConvertPouchtoCrateonDispatch, clsFixedParameterCode.ConvertPouchtoCrateonDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyCommissionTPTForCreditCustomer, clsFixedParameterCode.ApplyCommissionTPTForCreditCustomer, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.ItemBatchWiseStartDate, clsFixedParameterCode.ItemBatchWiseStartDate, EnumControlType.TextBox)

        InsertDefaultValue(clsUserMgtCode.rptMonthWiseDMRReport, clsFixedParameterType.PickRatefromMaster, clsFixedParameterCode.PickRatefromMaster, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyTPT, clsFixedParameterCode.ApplyTPT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.AllowMultipleUOMForProduct, clsFixedParameterCode.AllowMultipleUOMForProduct, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AllowManualCrateForDispatch, clsFixedParameterCode.AllowManualCrateForDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyManualScheme, clsFixedParameterCode.ApplyManualScheme, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, clsFixedParameterType.AllowResetCustomerDemandOnRouteChange, clsFixedParameterCode.AllowResetCustomerDemandOnRouteChange, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.AutoSchemeOnTotalDispatchQty, clsFixedParameterCode.AutoSchemeOnTotalDispatchQty, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyDefaultTCSIsChecked, clsFixedParameterCode.ApplyDefaultTCSIsChecked, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.ApplyDefaultTPTIsChecked, clsFixedParameterCode.ApplyDefaultTPTIsChecked, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmBoothRouteMapping, clsFixedParameterType.ApplyBoothRouteMapping, clsFixedParameterCode.ApplyBoothRouteMapping, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDistributorCommission, clsFixedParameterType.EnableVehicleType, clsFixedParameterCode.EnableVehicleType, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptCollectionDataChangeReport, clsFixedParameterType.PickDataFromRetestingTable, clsFixedParameterCode.PickDataFromRetestingTable, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.rptCustItemWiseSaleReport, clsFixedParameterType.ShowAllCustomerItemWiseSaleReportOptions, clsFixedParameterCode.ShowAllCustomerItemWiseSaleReportOptions, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmRoundMaster, clsFixedParameterType.ApplyDepartmentRoute, clsFixedParameterCode.ApplyDepartmentRoute, EnumControlType.CheckBox)
        InsertDefaultValue("ASDS", clsFixedParameterType.QuantityTolerance, clsFixedParameterCode.QuantityTolerance, EnumControlType.NumericBox)
        InsertDefaultValue("ASDS", clsFixedParameterType.AmountTolerance, clsFixedParameterCode.AmountTolerance, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.ApplyItemCapacityLimit, clsFixedParameterCode.ApplyItemCapacityLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.SeprateMorningEveningSequence, clsFixedParameterCode.SeprateMorningEveningSequence, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.FinancialImpactForTPT, clsFixedParameterCode.FinancialImpactForTPT, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.FinancialImpactForDistributor, clsFixedParameterCode.FinancialImpactForDistributor, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.FinancialImpactForSecurity, clsFixedParameterCode.FinancialImpactForSecurity, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickDemand, clsFixedParameterType.QuickDemandUOMCrate, clsFixedParameterCode.QuickDemandUOMCrate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickDemand, clsFixedParameterType.QuickDemandUOMPouch, clsFixedParameterCode.QuickDemandUOMPouch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmQuickDemand, clsFixedParameterType.QuickDemandUOMLtr, clsFixedParameterCode.QuickDemandUOMLtr, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyMonthEndDispatch, clsFixedParameterCode.ApplyMonthEndDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, clsFixedParameterType.ApplyItemUOMOnDemand, clsFixedParameterCode.ApplyItemUOMOnDemand, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.MergeTCAmtofCreditCust, clsFixedParameterCode.MergeTCAmtofCreditCust, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmProductDispatch, clsFixedParameterType.ConvertIntoBulkUOM, clsFixedParameterCode.ConvertIntoBulkUOM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyBoothWiseScheme, clsFixedParameterCode.ApplyBoothWiseScheme, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AddDistributorCommissioninItemPrice, clsFixedParameterCode.AddDistributorCommissioninItemPrice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.SubstractDistributorCommissioninItemPrice, clsFixedParameterCode.SubstractDistributorCommissioninItemPrice, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ConvertTOBillingUOM, clsFixedParameterCode.ConvertTOBillingUOM, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, clsFixedParameterType.DifferentCrateTypeForFGItem, clsFixedParameterCode.DifferentCrateTypeForFGItem, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.StopMultipleLoginInAPP, clsFixedParameterCode.StopMultipleLoginInAPP, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DispatchCommissionDecimalPlaces, clsFixedParameterCode.DispatchCommissionDecimalPlaces, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyPricePlanOnDocumentDate, clsFixedParameterCode.ApplyPricePlanOnDocumentDate, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.FrmUtilityForm, clsFixedParameterType.DisplayTransporterAddress, clsFixedParameterCode.DisplayTransporterAddress, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmRouteMaster, clsFixedParameterType.AllowRouteWiseDemandEntryInDecimal, clsFixedParameterCode.AllowRouteWiseDemandEntryInDecimal, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AllowToCheckEntryUOMForProduct, clsFixedParameterCode.AllowToCheckEntryUOMForProduct, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.userMaster, clsFixedParameterType.AllowToCheckEntryUOMForGNG, clsFixedParameterCode.AllowToCheckEntryUOMForGNG, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.AllowToCheckZeroQtyonDispatch, clsFixedParameterCode.AllowToCheckZeroQtyonDispatch, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DefaultEnableEWayBill, clsFixedParameterCode.DefaultEnableEWayBill, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.CreateAutoGatePass, clsFixedParameterCode.CreateAutoGatePass, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementCorrection, clsFixedParameterType.SuspenceAdjusmtmentFATSNFTolerancePercentage, clsFixedParameterCode.SuspenceAdjusmtmentFATSNFTolerancePercentage, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.ApplyEWBThresholdLimit, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraCity, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraState, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForInterState, EnumControlType.NumericBox)
        InsertDefaultValue(clsUserMgtCode.frmSendBillToDCS, clsFixedParameterType.ThirtPartyFarmerCollectionIntegration, clsFixedParameterCode.ThirtPartyFarmerCollectionIntegration, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DefaultEnableNoTransporter, clsFixedParameterCode.DefaultEnableNoTransporter, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ManualBatchOnCustomerBooking, clsFixedParameterCode.ManualBatchOnCustomerBooking, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, clsFixedParameterType.GenerateCustomerWiseGatePass, clsFixedParameterCode.GenerateCustomerWiseGatePass, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.DeductTPTFromDocAmt, clsFixedParameterCode.DeductTPTFromDocAmt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.CheckCustomeroutStandingAmt, clsFixedParameterCode.CheckCustomeroutStandingAmt, EnumControlType.CheckBox)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, clsFixedParameterType.ServerDateTimeForTaxableInvoice, clsFixedParameterCode.ServerDateTimeForTaxableInvoice, EnumControlType.CheckBox)
    End Sub
End Class