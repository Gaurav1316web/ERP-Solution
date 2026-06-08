Imports common
Imports System.Reflection
Imports System.Windows.Forms
Public Class clsDocType
    Public Const RCDFChecmicalQC As String = "RCDF Checmical QC"
    Public Const RCDFUnionQC As String = "RCDF Union QC"
    Public Const ItemBatch As String = "Item Batch"
    Public Const RouteWiseSaleTarget As String = "Route Wise Sale Target"
    Public Const BMC_Transporter_Bill As String = "BMC Transporter Bill"
    Public Const frmCommissionTPTCalculation As String = "Commission/TPT Calculation"
    Public Const DCSDispatch As String = "DCS Dispatch"
    Public Const DCSInvoice As String = "DCS Invoice"
    Public Const DCSSaleEntry As String = "DCS Sale Entry"
    Public Const ShortSupplyPenalty As String = "Short Supply Penalty"
    Public Const frmRCDFRateControl As String = "RCDF Rate Control"
    Public Const frmEmployeeDeductionMaster As String = "Employee Deduction Master"
    Public Const TenderPenalty As String = "Tender Penalty"
    Public Const RALNoc As String = "RAL NOC"
    Public Const BreakDownEntry As String = "Break Down Entry"
    Public Const OutputEntry As String = "Output Entry"
    Public Const ShipToLocation = "Ship To Location"
    Public Const CommonSaleSeries = "Common Sale Series"
    Public Const ReceiptInvoiceMapping = "Receipt Invoice Mapping"
    Public Const VendorPriceMapDocwise = "Vendor Price Mapping"
    Public Const GateReturnProductSale = "Gate Return Product Sale"
    Public Const GateReturnCSASale = "Gate Return CSA Sale"
    Public Const GateReturnTransfer = "Gate Return Transfer"
    Public Const IntimationScreen As String = "Intimation Screen"
    Public Const ContractTanker = "Contract Tanker"
    Public Const APInvoice As String = "AP Invoice"
    Public Const MultipleProcDed As String = "Multiple Proc Ded"
    Public Const TransferToSaving As String = "Transfer To Saving"
    Public Const BMCTransporterBill As String = "BMC Transporter Bill"
    Public Const EmployeeOTEntry As String = "Employee OT Entry"
    Public Const TransporterPaymentProcess As String = "Transporter Payment Process"
    Public Const EmployeeOTEntryCalculation As String = "Employee OT Entry Calculation"
    Public Const CHACHARGEMASTER As String = "CHA Charge Master"
    Public Const DebitNote As String = "AP Debit Note"
    Public Const CreditNote As String = "AP Credit Note"
    Public Const BullMasters As String = "Bull Masters"
    ''richa job work module
    Public Const GateEntryJWO As String = "JWO Gate Entry"
    Public Const WeighmentJWO As String = "JWO Weighment"
    Public Const QCJWO As String = "JWO Quality Check"
    Public Const UnloadingJWO As String = "JWO Unloading"
    ''--------
    ''richa can sale
    Public Const CanSale As String = "Can Sale"
    Public Const CanSaleDispatch As String = "Can Sale Dispatch"
    Public Const CanSaleInvoice As String = "Can Sale Invoice"
    ''---------

    Public Const ARInvoice As String = "AR Invoice"
    Public Const ARDebitNote As String = "AR Debit Note"
    Public Const ARCreditNote As String = "AR Credit Note"

    Public Const SupplementaryARInvoice As String = "Supplementary AR Invoice"
    Public Const SupplementaryARCreditNote As String = "Supplementary AR Credit Note"

    Public Const ARServiceInvoice As String = "AR Service Invoice"
    Public Const ARServiceCreditNote As String = "AR Service Credit Note"
    Public Const ARServiceDebitNote As String = "AR Service Debit Note"

    ''Parteek
    Public Const BankRevseEnty As String = "Bank Reverse Entry"
    Public Const BankReco As String = "Bank Reco"
    Public Const JobWkDebit As String = "JobWork Debit Note"
    Public Const RackMaster As String = "Rack Master"
    Public Const BinMaster As String = "Bin Master"

    ''End

    ''Rakesh
    Public Const ConsumerDetails As String = "Consumer Details"
    Public Const AdditionalCharges As String = "Additional Charges"
    Public Const BankMaster As String = "Bank Master"
    Public Const AreaMaster As String = "Area Master"
    Public Const CompanyMaster As String = "Company Master"
    Public Const CityMaster As String = "City Master"
    Public Const DistrictMaster As String = "District Master"
    Public Const RegionMaster As String = "Region Master"
    Public Const CurrencyConversion As String = "Currency Conversion"
    Public Const FormMaster As String = "Form Master"
    Public Const PaymentCode As String = "Payment Code"
    Public Const PaymentTerms As String = "Payment Terms"
    Public Const TaxAuthority As String = "Tax Authority"
    Public Const TaxGroup As String = "Tax Group"
    Public Const TaxRates As String = "Tax Rates"
    Public Const TimeTable As String = "Time Table"
    Public Const DesignationMaster As String = "Designation Master"
    Public Const AcquisitionCodes As String = "Acquisition Codes"
    Public Const HirerachyLevelMaster As String = "Hirerachy Level Master"
    Public Const VendorAccountSet As String = "Vendor Account Set"
    Public Const VendorGroup As String = "Vendor Group"
    Public Const VendorType As String = "Vendor Type"
    Public Const CostCenter As String = "Cost Center"
    Public Const TankerMasterSale As String = "Tanker Master Sale"
    Public Const DispatchCheckList As String = "Dispatch Check List"
    Public Const CustomerType As String = "Customer Type"
    Public Const CustomerAccountSet As String = "Customer Account Set"
    Public Const RouteTypeMaster As String = "Route Type Master"
    Public Const VehicleMaster As String = "Vehicle Master"
    Public Const TransportType As String = "Transport Type"
    Public Const ChannelCategory As String = "Channel Category"
    Public Const ChannelMaster As String = "Channel Master"
    Public Const SalesManHierarchy As String = "Sales Man Hierarchy"
    Public Const SettlementMaster As String = "Settlement Master"
    Public Const RouteGroupMaster As String = "Route Group Master"
    Public Const DiscountCategoryMaster As String = "Discount Category Master"
    Public Const SamplingMaster As String = "Sampling Master"
    Public Const TransportMaster As String = "Transport Master"
    Public Const RemarkMaster As String = "Remark Master"
    Public Const DiscountMaster As String = "Discount Master"
    Public Const ZoneMaster As String = "Zone Master"
    Public Const RouteMaster As String = "Route Master"
    Public Const SalesLevelHierarchy As String = "Sales Level Hierarchy"
    Public Const MerchantPaymentTerms As String = "Merchant Payment Terms"
    Public Const RequisitSubTypeMaster As String = "Requisit Sub Type Master"
    Public Const CustomerCategory As String = "Customer Category"
    Public Const ItemCategoryLevel As String = "Item Category Level"
    Public Const ItemCategoryStructure As String = "Item Category Structure"
    Public Const ReceivablePaymentTerms As String = "Receivable Payment Terms"
    Public Const ShipToLocationDetails As String = "Ship To Location Details"
    Public Const SecondaryCustomerMaster As String = "Secondary Customer Master"
    Public Const AccountMainGroup As String = "Account Main Group"
    Public Const AccountGroup As String = "Account Group"
    Public Const AccountSubGroup As String = "Account Sub Group"
    Public Const MainGLAccount As String = "Main GLA Account"
    Public Const GLAccount As String = "GLA Account"
    Public Const CostCentreFinancial As String = "Cost Centre Financial"
    Public Const UserMaster As String = "User Master"
    Public Const UserGroupMaster As String = "User Group Master"
    Public Const ItemStructure As String = "Item Structure"
    Public Const PurchaseAccountSetCode As String = "Item Structure"
    Public Const ChapterHead As String = "Chapter Head"
    Public Const ItemCategory As String = "Item Category"
    Public Const ItemSubCategory As String = "Item Sub Category"
    Public Const PriceComponantMaster As String = "Price Componant Master"
    Public Const PriceComponantMapping As String = "Price Componant Mapping"
    Public Const BreakageHead As String = "Breakage Head"
    Public Const Warranty As String = "Warranty"
    Public Const InventorySourceCode As String = "Inventory Source Code"
    Public Const ItemType As String = "Item Type"
    Public Const ResponsiblePerson As String = "Responsible Person"
    Public Const BranchDetails As String = "Branch Details"
    Public Const NatureOfDeduction As String = "Nature Of Deduction"
    Public Const PartyDetails As String = "Party Details"
    Public Const TDSSection As String = "TDS Section"
    Public Const StateCode As String = "State Code"
    Public Const Category As String = "Category"
    Public Const DeprAccountSet As String = "Depr Account Set"
    Public Const FACostCenter As String = "FA Cost Center"
    Public Const AssetGroups As String = "Asset Groups"
    Public Const AssetBook As String = "Asset Book"
    Public Const DepreciationMethod As String = "Depreciation Method"
    Public Const JWInwardFormula As String = "JW Formula"
    Public Const JWOFormula As String = "OJW Formula"
    Public Const JWOVendorFormula As String = "OJW Vendor Formula Mapping"
    Public Const CusomerComplain As String = "Cusomer Complain"
    Public Const JWEstimate As String = "JW Estimate"
    Public Const JWOEstimate As String = "OJW Estimate"
    Public Const JWVendorFormula As String = "Vendor Formula Mapping"
    Public Const DepreciationPeriods As String = "Depreciation Periods"
    Public Const TemplateMaster As String = "Template Master"
    Public Const SalaryAccountSetting As String = "Salary Account Setting"
    Public Const CountryMaster As String = "Country Master"
    Public Const StateMaster As String = "State Master"
    Public Const BranchMaster As String = "Branch Master"
    Public Const OTMaster As String = "OT Master"
    Public Const AttendanceMaster As String = "Attendance Master"
    Public Const BonusMaster As String = "Attendance Master"
    Public Const PayPeriodMaster As String = "Attendance Master"
    Public Const PFRulesMaster As String = "PF Rules Master"
    Public Const ESIRulesMaster As String = "ESI Rules Master"
    Public Const LeaveMaster As String = "Leave Master"
    Public Const ShiftMaster As String = "Shift Master"
    Public Const DepartmentMaster As String = "Department Master"
    Public Const DevisionMaster As String = "Devision Master"
    Public Const CastCategory As String = "Cast Category"
    Public Const DocumentMaster As String = "Document Master"
    Public Const CourseMaster As String = "Course Master"
    Public Const GradeMaster As String = "Grade Master"
    Public Const LanguageMaster As String = "Language Master"
    Public Const OccupationMaster As String = "Occupation Master"
    Public Const ReligionMaster As String = "Grade Master"
    Public Const SkillMaster As String = "Grade Master"
    Public Const OTSLab As String = "OT Slab"
    Public Const PTSLab As String = "PT Slab"
    Public Const HRRSLab As String = "HRR Slab"
    Public Const ConveyanceRateMaster As String = "Conveyance Rate Master"
    Public Const ODMaster As String = "OD Master"
    Public Const SubDepartmentMaster As String = "Sub Department Master"
    Public Const PayrollPaymentMode As String = "Payroll Payment Mode"
    Public Const MFAccountSet As String = "MF Account Set"
    Public Const ExpenseHead As String = "Expense Head"
    Public Const MFToolType As String = "MF Tool Type"
    Public Const WorkCenterMaster As String = "WorkCenterMaster"
    Public Const ResourceMaster As String = "Resource Master"
    Public Const MFToolMaster As String = "MF Tool Master"
    Public Const OperationMaster As String = "Operation Master"
    Public Const SectionMaster As String = "Section Master"
    Public Const StageMaster As String = "StageMaster"
    Public Const ItemCategoryForProduction As String = "Item Category For Production"
    Public Const ProductionLine As String = "Production Line"
    Public Const PJCSetting As String = "PJC Setting"
    Public Const CostType As String = "Cost Type"
    Public Const PJCAccountSet As String = "PJC Account Set"
    Public Const JobMaster As String = "Job Master"
    Public Const TaskMaster As String = "Task Master"
    Public Const ComplaintGroup As String = "Complaint Group"
    Public Const PrimaryReasonMaster As String = "Primary Reason Master"
    Public Const DeductionMaster As String = "Deduction Master"
    Public Const DeductionGroup As String = "Deduction Group"
    Public Const MilkReasonMaster As String = "Milk Reason Master"
    Public Const PaymentCycleMaster As String = "Payment Cycle Master"
    Public Const ItemChargeCategoryMaster As String = "Item Charge Category Master"
    Public Const SupplierMaster As String = "Supplier Master"
    Public Const DivertedContractor As String = "Diverted Contractor"
    Public Const MilkGradeMaster As String = "Milk Grade Master"
    Public Const MilkTypeMaster As String = "Milk Type Master"
    ''
    ''Public Const AdjustmentEntry As String = "Adjustment Entry"
    Public Const TransferDCC As String = "Transfer"
    Public Const LoadOut As String = "Load Out"
    'Public Const GPEntry As String = "GPEntry"
    Public Const LoadIn As String = "Load In"
    Public Const Payment As String = "Payment"
    Public Const Receipt As String = "Receipt"
    Public Const RICEMIXING As String = "Mixing Entry(Rice)"
    Public Const RICEPROCESSING As String = "Processing Entry(Rice)"
    Public Const JournalEntry As String = "Journal Entry"
    Public Const JournalEntryOP As String = "Journal Entry OP"
    Public Const ReverseJournalEntry As String = "Reverse Journal Entry"
    Public Const SaleOrder As String = "Sale Order"
    Public Const SaleInvoice As String = "Sale Invoice"
    Public Const Shipment As String = "Shipment"
    Public Const SaleReturn As String = "Sale Return"
    Public Const PurchaserRegusitsion As String = "Purchase Requistion"
    Public Const frmStoreRequistion As String = "Store Requistion"
    Public Const frmProductionStoreRequest As String = "Production Store Request"
    Public Const PurchaserOrder As String = "Purchase Order"
    Public Const PurchaserOrderOutward As String = "Purchase Order Outward"
    Public Const DistributeCode As String = "Distribute Code"

    Public Const AdjustmentEntry As String = "Adjustment Entry (Finance)"
    Public Const PaymentAdjustmentEntry As String = "Adjustment Entry (Payment)"
    Public Const QuickSettlement As String = "Quick SettleMent"
    Public Const GRN As String = "Gate Receipt Note"

    Public Const POWeightment As String = "PO Weightment"
    Public Const POWeightmentOG As String = "PO Weightment OG"
    Public Const TenderShortPenalty As String = "Tender Short Penalty"
    Public Const MRN As String = "Material Receipt Note"
    Public Const SRN As String = "Store Receipt Note"
    Public Const GTOut As String = "Gate Out"
    Public Const OutgoingProduction As String = "Outgoing Production"
    Public Const RMProcessLoss As String = "RM Process Loss"
    Public Const DAArrear As String = "DA Arrear"
    Public Const DAArrearNew As String = "DA Arrear"
    Public Const LeaveEncashment As String = "Leave Encashment"
    Public Const EPF As String = "EP Fund"
    Public Const SRNReturn As String = "Store Receipt Note Return"
    Public Const NIRQC As String = "NIR QC"
    Public Const TransferReturn As String = "Transfer Return"
    Public Const GatePasstransfer As String = "Gate Pass Transfer"
    Public Const MTSRN As String = "Merchant SRN"
    Public Const POInvoice As String = "PO Invoice"
    Public Const PurchaseTaxInvoice As String = "Purchase Tax Invoice"
    Public Const PurchaseBillOfSupply As String = "Purchase Bill Of Supply"
    Public Const MT_POInvoice As String = "Merchant Purchase Invoice"
    Public Const Mcc_TransporterInvoice As String = "Mcc_TransporterInvoice"
    Public Const RMDA As String = "Rejected Materail Disp. Advice"
    Public Const RGP As String = "Returnable Gate Pass"
    Public Const RGPR As String = "RGP Return"
    Public Const NRGP As String = "Non Returnable Gate Pass"
    Public Const NRGPR As String = "NRGP Return"
    Public Const PJV As String = "Purchase Journal Voucher"
    Public Const PurchaseReturn As String = "Purchase Return"
    Public Const IssueReturn As String = "Issue/Return/Transfer"
    Public Const SalesReturn As String = "Sales Return"
    Public Const Scrap As String = "Scrap"
    Public Const ScrapReturn As String = "ScrapRetrun"
    Public Const ScrapSaleGateOut As String = "Misc. Sale Gate Out"
    Public Const DisposalEntry As String = "Disposal Entry"
    Public Const ScrapInvoice As String = "Scrap Invoice"
    Public Const VCGLEntry As String = "Vendor/Customer/GL"
    Public Const CreateRemmitance As String = "Create Remitance"
    Public Const FormIssue As String = "Form Issue"
    'Public Const EmptyTransaction As String = "Empty Transaction"
    'Public Const ProductionEntry As String = "Production Entry"
    Public Const warehouseBreakage As String = "Warehouse Breakage"
    Public Const StoreAdjustment As String = "Store Adjustment"
    Public Const StoreAdjustmentProductionEntry As String = "Production Entry"
    Public Const StoreAdjustmentProductionEntryQC As String = "Production Entry QC"
    Public Const StoreAdjustmentProductionStoreEntry As String = "Production Store Entry"
    Public Const JobWorkInventory As String = "JobWork Inventory"
    Public Const RawMilkConsumtion As String = "RawMilk Consumption"
    Public Const SalesReturnInterCompany As String = "Sales Return (Inter Company)"
    Public Const ReveseNo As String = "Reverse Entry"
    Public Const OTSheet As String = "OT Sheet"
    Public Const GeneralHolidays As String = "General Holidays"
    Public Const Indent As String = "Indent"
    Public Const LeaveAllotment As String = "Leave Allotment"
    Public Const LeaveAdjustment As String = "Leave Adjustment"
    Public Const LeaveApplication As String = "Leave Application"
    Public Const LoanAdjustment As String = "Loan Adjustment"
    Public Const LoanGeneration As String = "Loan Generation"
    Public Const ClaimMaster As String = "Claim Master"
    Public Const Employee_Master As String = "Employee Master"
    Public Const Enquiry_Master As String = "Enquiry Master"
    Public Const CSAPRICEMAASTER As String = "CSA Price Master"
    Public Const CSADELIVERYORDER As String = "CSA Delivery Order"
    Public Const PayHeadDefinitions As String = "Pay Head Definitions"
    Public Const MonthlyAttendance As String = "Monthly Attendance"
    Public Const ReimbursementDetails As String = "Reimbursement Details"
    Public Const WeeklyHoliday As String = "Weekly Holiday"
    Public Const SalaryGeneration As String = "Salary Generation"
    Public Const SentMail As String = "send Mail"
    Public Const SalaryAdjustment As String = "Salary Adjustment"
    Public Const AllowanceDetails As String = "Allowance Details"
    Public Const ApplyLoan As String = "Apply Loan"
    Public Const LTAClaim As String = "LTA Claim"
    Public Const StandardRateItem As String = "Standard Rate of Item"
    Public Const DeductionDetails As String = "Deduction Details"
    Public Const EmployeeSalary As String = "Employee Salary"
    Public Const EmployeeStatus As String = "Employee Status"
    Public Const EmployeeIncrement As String = "Employee Increment"
    Public Const DailyAttendance As String = "Daily Attendance"
    Public Const HourlyAttendance As String = "Hourly Attendance"
    Public Const EmpIncrement As String = "EmpIncrement"
    Public Const GenerateBonus As String = "Generate Bonus"
    Public Const SmsEmailScheduler As String = "Sms Email Scheduler"
    Public Const StandardScheme As String = "Standard Scheme"
    Public Const TransferGateOut As String = "Transfer Gate Out"
    Public Const ProductDispatchGateOut As String = "Product Dispatch Gate Out"
    Public Const CSATransferGateOut As String = "CSA Transfer Gate Out"
    Public Const MCCTankerGateOut As String = "MCC Tanker Gate Out"
    Public Const MCCTankerGateOutSecurity As String = "MCC Tanker Gate Out Security"
    Public Const ItemProjection As String = "Item Projection"
    Public Const SNPOS As String = "POS"
    Public Const SNSalesQuotation As String = "Sales Quotation"
    Public Const SNSalesOrder As String = "Sales Order"
    Public Const SNShipment As String = "Sales Shipment"
    Public Const SNSaleInvoice As String = "Sales Invoice"
    Public Const EXSALESQUOTATION As String = "Export Sales Quotation"
    Public Const EXSALESORDER As String = "Export Sales Order"
    Public Const EXSALESINVOICE As String = "Export Sales Invoice"
    Public Const EXPORFORMAINVOICE As String = "Export Proforma Invoice"
    Public Const EXCOMMERSIALINVOICE As String = "Export Commercial Invoice"
    Public Const EXSALESRETURN As String = "Export Sale Return"

    Public Const MTSALESQUOTATION As String = "Merchant Sales Quotation"
    Public Const MTSALESORDER As String = "Merchant Sales Order"
    Public Const MTPORFORMAINVOICE As String = "Merchant Proforma Invoice"
    Public Const MTCOMMERSIALINVOICE As String = "Merchant Commercial Invoice"
    Public Const MTSALESINVOICE As String = "Merchant Sales Invoice"
    Public Const MTSALESRETURN As String = "Merchant Sale Return"

    Public Const CSASaleInvoice As String = "CSA Sale Invoice"
    Public Const SNSaleReturn As String = "Sales Invoice Return"
    Public Const CSASALERETURN As String = "CSA Sale Return"
    Public Const frmSaleReturnFreshSale As String = "Sale Return FreshSale"
    Public Const frmSaleReturnProductSale As String = "Sale Return ProductSale"

    Public Const DairyProductionUploader As String = "Dairy Production Uploader"
    Public Const ShiftMgmt As String = "Shift Management"
    Public Const ShiftMgmtSFG As String = "Shift Management SFG"
    Public Const BarCode As String = "Bar Code"
    Public Const ShipToMaster As String = "Ship To Master"
    Public Const BOM As String = "BOM"
    Public Const RFQ As String = "Request For Quotation"
    Public Const RMDemandApproval As String = "RM Demand Approval"
    Public Const VendorQuotation As String = "Vendor Quotation"
    Public Const SetPOSchedule As String = "Set PO Schedule"
    Public Const BATCHORDER As String = "BATCH ORDER"
    Public Const NOTIFYPARTY As String = "Notify Party Master"
    Public Const PPISSUEENTRY As String = "Production Issue Entry"
    Public Const PPSTANDARDIZATION As String = "Production Standardization"
    Public Const PPSTDFinalQC As String = "Prod Stand Final QC"
    Public Const PPStageProcess As String = "Stage Process"
    Public Const ppProductionEntry As String = "Production Entry"
    Public Const ppSTProductionEntry As String = "ST Production Entry"
    Public Const ProductionEntryFinalQC As String = "Production Entry Final QC"
    Public Const PPProductionReturn As String = "Production Return"
    Public Const WreckageBooking As String = "Wreckage Booking"
    Public Const PPLOGSHEET As String = "QC Log Sheet"
    Public Const MO As String = "Manufacturing Order"
    Public Const ProductionRequisition As String = "Production Requisition"
    Public Const ProductionIssue As String = "Production Issue"
    Public Const ProductionMapping As String = "Production Mapping"
    Public Const ProductionReturn As String = "Production Return"
    Public Const ProductionReceipt As String = "Production Receipt"
    Public Const StandardProductionPlanning As String = "Standard Production Planning"
    Public Const ProductionPlanning As String = "Production Planning"
    Public Const SerializedReplace As String = "Replace Serialized Entry"
    Public Const AcquisitionEntry As String = "Acquisition Entry"
    Public Const FAAssembleAsset As String = "Assemble Asset"
    Public Const AssetAccountChange As String = "Asset Account Change"
    Public Const QCVendorItemMapping As String = "QC Vendor-Item Mapping"
    Public Const QualityCheckForSRN As String = "Quality Check Entry"
    Public Const frmOutgoingQCEntry As String = "  QC Production "
    Public Const AssetDepreciation As String = "Asset Depreciation"
    Public Const AssetRequisition As String = "Asset Requisition"
    Public Const ExpiryTransaction As String = "Expiry Transaction"
    Public Const POSCHEDULE As String = "Purchase Schedule"
    'Public Const CFormEntry As String = "CForm Entry"
    Public Const Assemblies As String = "Assemblies"
    Public Const PROD_Assemblies As String = "Production Assemblies"
    Public Const CheckSlip As String = "CheckSlip"
    Public Const AssetAgreement As String = "Asset Agreement"
    Public Const AssetDispatch As String = "Asset Dispatch"
    Public Const complaintdetail As String = "Complaint Detail"
    Public Const ComplaintEntry As String = "Complaint Entry"
    Public Const CartMaintenance As String = "Cart Maintenance Entry"
    Public Const CollectionCenter As String = "Collection Center"
    Public Const MilkCollectionMCC As String = "Milk Collection MCC"
    Public Const MilkCollectionMCCMuliple As String = "Milk Collection MCC Multiple"
    Public Const MilkCollectionDCS As String = "Milk Collection DCS"
    Public Const MilkShiftUploader As String = "Milk Shift Uploader"
    Public Const MilkProcurementUploader As String = "Milk Procurement Uploader"
    Public Const MilkCollectionDCSMuliple As String = "Milk Collection DCS Multiple"
    Public Const MilkCollectionDCSMulipleMerge As String = "Milk Coll DCS Multiple Merge"
    Public Const HeadLoadDCS As String = "Head Load DCS"
    Public Const MultipleShareAllotment As String = "Multiple Share Allotment"
    Public Const CustomerPenalty As String = "Customer Penalty"
    Public Const StartBatchEntry As String = "Start Batch Entry"
    Public Const CFSalePurchaseUploader As String = "CF Sale Purchase Uploader"
    Public Const QuickPaymentBySingleCheque As String = "Quick Payment By Single Cheque"
    Public Const BulkSaleFreightMaster As String = "Bulk Sale Freight Master"
    Public Const BullVaccinationEntry As String = "Bull Vaccination Entry"
    Public Const BullInsurance As String = "Bull Insurance"
    Public Const InsuranceTagAllocation As String = "Insurance Tag Allocation"
    Public Const SeedSelectionEntry As String = "Seed Selection Entry"
    Public Const VSPMASTER As String = "VSP Master"
    Public Const PTMMASTER As String = "PTM Master"
    Public Const TTMMASTER As String = "TTM Master"
    Public Const VLCMASTER As String = "VLC Master"
    Public Const MILKPRCMASTER As String = "Milk Price Master"
    Public Const TankerDispatchPriceMaster As String = "Tanker Dispatch Price Master"
    Public Const ListofPendingDocuments As String = "List of Pending Documents"
    Public Const MCCMaster As String = "MCC Master"
    Public Const MPMaster As String = "MP Master"
    Public Const MilkReceipt As String = "MilkReceipt"
    Public Const MilkReject As String = "Milk Reject"
    Public Const FrmGrievanceLogging As String = "FrmGrievanceLogging"
    Public Const VLCTarget As String = "VLCMasterTGT"
    Public Const MilkTransportorInvoice As String = "MilkTransportor"
    Public Const MilkSample As String = "MilkSample"
    Public Const MilkSRN As String = "MilkSRN"
    Public Const MilkPO As String = "Milk Purchase Order"
    Public Const MilkPurInvoice As String = "MilkPurInvoice"
    Public Const MilkPurInvoiceProvsion As String = "Prov Milk Pur Invoice"
    Public Const VSPIncentiveTagging As String = "VSPIncentiveTagging"
    Public Const Sch_Training As String = "Sch_Training"
    Public Const MilkShift_End As String = "MilkShift_End"
    Public Const VLCROUTESHIFT As String = "VLC Route Shift"
    Public Const CustomerRouteShift As String = "Customer Route Shift"
    Public Const MCCMaterialSale As String = "MCC Material Sale"
    Public Const MCCMaterialSaleReturn As String = "MCC Material Sale Return"
    Public Const MILKROUTEMASTER As String = "Milk Route Master"
    Public Const VILLAGEMASTER As String = "Village Master"
    Public Const MccDispatchChallan As String = "MCC Dispatch Challan"
    Public Const AcknowledgementEntry As String = "Acknowledgement Entry"
    Public Const MccDispatchChallanReturn As String = "MCC Dispatch Challan Return"
    Public Const MccTnkDispChallanReturn As String = "Tanker Dispatch Challan Return"
    Public Const LostDefectSealNo As String = "Lost Defect Seal No"
    Public Const ParamMaster As String = "Parameter Master"
    Public Const QCLOGSHEETMST As String = "QC Log Sheet Master"
    Public Const GateEntry As String = "Gate Entry"
    Public Const POBulkP As String = "PO Bulk Proc"
    Public Const TDSPayment As String = "TDS Payment"
    Public Const Weighment As String = "Weighment"
    'Public Const FreshSale As String = "Fresh Sale"
    Public Const BookingEntry As String = "Booking Entry"
    Public Const DeliveryNoteFreshSale As String = "Fresh Delivery"
    Public Const DispatchNoteFreshSale As String = "Fresh Dispatch"
    Public Const frmInvoiceFreshSale As String = "Fresh Invoice"
    Public Const frmSaleReturnGateEntry As String = "Sale Return Gate Entry"
    Public Const frmCreateReceived As String = "Fresh Crate Received"
    Public Const frmCreateReceivedAPI As String = "Fresh Crate Received API"
    Public Const frmCreateOut As String = "Fresh Crate Out"
    Public Const frmCanReceived As String = "Can Received"
    Public Const frmCanReceivedAPI As String = "Can Received API"
    Public Const FrmInvoiceCrateLinerDetail As String = "Invoice Crate Liner Detail"

    Public Const MilkGateEntry As String = "Milk Gate Entry"
    Public Const MilkWeighment As String = "Milk Weighment"
    Public Const JobMilkQualityCheck As String = "Job Milk Quality Check"
    Public Const MilkUnloading As String = "Job Milk Unloading"
    Public Const MilkRGP As String = "Milk RGP"
    ''richa agarwal changes against ticket no BM00000005253
    Public Const GatePassEntrySale As String = "Bulk Gate Entry"

    Public Const MCCGateEntry As String = "MCC Gate Entry"
    Public Const MCCWeighment As String = "MCC Weighment"
    ''----------------------------------------
    '--------priti 10/09/2014 ------------------
    Public Const frmBookingProductSale As String = "Booking Product Sale"
    Public Const frmDispatchAdviceProductSale As String = "Dispatch Advice Product Sale"
    Public Const frmDeliveryPrderProductSale As String = "Product Delivery Order"
    Public Const frmSalesOrderPSForExempted As String = "PS Sales Order Exempted"
    Public Const frmDeliveryOrderPSForExempted As String = "PS Delivery Order Exempted"
    Public Const frmSalesOrderProductSale As String = "Sale Order Product Sale"
    Public Const frmShipmentProductSale As String = "Shipment Product Sale"
    Public Const frmCustomerTender As String = "Customer Tender"
    Public Const frmCustomerTenderOrder As String = "Customer Tender Order"
    Public Const frmCustomerTenderDispatch As String = "Customer Tender Dispatch"
    Public Const frmCustomerTenderInvoice As String = "Customer Tender Invoice"
    Public Const frmCustomerTenderGatePass As String = "Customer Tender GatePass"
    Public Const frmLeakedSaleReturn As String = "Leaked Sale Return"

    Public Const frmSaleInvoiceProductSale As String = "Product Invoice"
    Public Const frmTransferKDIl As String = "Transfer KDIL"
    Public Const frmTransferGST As String = "Transfer GST"
    Public Const FrmGatePassFS As String = "Fresh Gate Pass Entry"
    Public Const FrmGatePassPS As String = "Product Gate Pass Entry"

    '--------priti 13/07/2016 ------------------
    Public Const frmDairySaleDeliveryOrder As String = "Dairy Sale Delivery Order"
    Public Const frmDemandBooking As String = "Demand Booking"
    Public Const frmProductDemandBooking As String = "Product Demand Booking"
    Public Const frmDemandBookingUploader As String = "Demand Booking Uploader"
    Public Const frmDemandAdjustment As String = "Demand Adjustment"
    Public Const frmDemandUploader As String = "Demand Uploader"
    Public Const FrmMCCMaterialSalePriceChart As String = "MCC Material Sale Price Chart"
    Public Const frmDCSDemandBooking As String = "DCS Demand Booking"
    Public Const frmDairySaleBooking As String = "Dairy Sale Booking"
    Public Const frmDairySaleBookingUploader As String = "Dairy Sale Booking Uploader"
    Public Const frmDairySaleGatePass As String = "Dairy Sale GatePass"
    Public Const frmDairySaleShipment As String = "Dairy Sale Shipment"
    Public Const frmDairySaleInvoice As String = "Dairy Sale Invoice"
    Public Const frmDairySaleReturn As String = "Dairy Sale Return"
    Public Const frmDairyShipmentReturn As String = "Dairy Shipment Return"
    Public Const FrmDairyGatePass As String = "Dairy GatePass"
    Public Const frmCustomerComplain As String = "Customer Complain"
    Public Const frmTranspoterDeductionEntry As String = "TranspoterDedEntry"
    Public Const FrmMCCScrapGatePass As String = "Mcc/Scrap GatePass"
    '====================added by preeti Gupta 06/10/2016=================
    Public Const frmPerformaInvoiceBooking As String = "Performa Invoice Booking"
    '=====================================================================
    Public Const rptMonthlySalesInvoice As String = "Monthly Sales Invoice"
    '--------Richa 21/07/2014 Against Ticket No BM00000003245
    Public Const BulkSalePriceChart As String = "Bulk Sale Price Chart"
    '-----------------------------------------------------
    '--------Richa 28/07/2014 Against Ticket No BM00000003157
    Public Const WeighmentEntryBulkSale As String = "Bulk Weighment Entry"
    Public Const SalesOrderBulkSale As String = "Bulk Sales Order"
    '-----------------------------------------------------
    '--------Richa 28/07/2014 Against Ticket No BM00000003158
    Public Const LoadingTankerBulkSale As String = "Bulk Loading Tanker"
    '-----------------------------------------------------
    Public Const QualityCheckBulkSale As String = "Quality Check Bulk Sale"
    Public Const DispatchBulkSale As String = "Bulk Dispatch"
    Public Const BulkSaleAcknowledgement As String = "Bulk Sale Acknowledgement"
    Public Const AcknowledgementGRN As String = "Acknowledgement of GRN"
    Public Const InvoiceBulkSale As String = "Bulk Invoice"
    Public Const ProformaInvoiceBulkSale As String = "Bulk Proforma Invoice"
    Public Const BulkSaleReturn As String = "Bulk Sale Return"
    Public Const MerchantTradePO As String = "Merchant Trade PO"
    Public Const CustomerOutstanding As String = "Customer Outstanding"
    Public Const SiloMilkTransfer As String = "Silo Milk Transfer"
    ''richa 01/04/2015
    Public Const BulkSaleReturnDispatch As String = "Bulk Dispatch Sale Return"
    ''-------------------------
    Public Const QualityCheck As String = "Quality Check"
    Public Const SecondarySettingForQC As String = "Secondary Setting For QC"
    Public Const FixedDeposit As String = "Fixed Deposit"
    Public Const PriceChartMasterBulk As String = "Price Chart Master (Bulk Proc)"
    Public Const Unloading As String = "Unloading"
    Public Const MilkTransferIn As String = "Milk Transfer In"
    Public Const MilkTransferInReturn As String = "Milk Transfer In Return"
    Public Const BulkMilkSRN As String = "Bulk Milk SRN"
    Public Const frmMilkJobWorkTransfer As String = "MILK JOB WORK TRANSFER"
    Public Const JobMilkSRN As String = "Job Milk SRN"
    Public Const JWOSRN As String = "JWO SRN"
    Public Const JWOSRNReturn As String = "JWO SRN Return"
    Public Const DailyElectricalEntry As String = "Daily Electrical Entry"
    Public Const CustomerIncentiveMaster As String = "Customer Incentive Master"
    Public Const TransporterDeductionMaster As String = "Transporter Deduction Master"
    Public Const InvestmentDeclarationMaster As String = "Investment Declaration Master"
    Public Const BulkMilkSRNReturn As String = "Bulk Milk SRN Return"
    Public Const BulkMilkPO As String = "Bulk Milk Purchase Order"
    Public Const BulkMilkPurchaseInvoice As String = "Bulk Milk Purchase Invoice"
    Public Const BulkMilkPurchaseInvoiceTrade As String = "Bulk Milk Purchase Inv Trade"
    ' Add By : Prabhakar 
    Public Const BulkMilkPurchaseReturnTrade As String = "Bulk Milk Purchase Rtn Trade"
    Public Const BulkMilkPurchaseReturn As String = "Bulk Milk Purchase Return"
    Public Const FarmerServiceOrder As String = "Farmer Service Order With Rate"
    Public Const DCSSale As String = "DCS Sale"
    Public Const DistributorCommission As String = "Distributor Commission"
    Public Const BoothCommission As String = "Booth Commission"
    Public Const BoothCalculation As String = "Booth Calculation"
    Public Const DCSTransportationCharges As String = "DCS Transportation Charges"
    Public Const DistributorRouteTagging As String = "Distributor Route Tagging"
    Public Const DailyDemand As String = "DailyDemand"
    Public Const ItemCapacityLimit As String = "ItemCapacityLimit"
    Public Const BoothRouteMapping As String = "BoothRouteMapping"

    Public Const frmofficeordertemplate As String = "office order template screen"
    Public Const frmofficeorder As String = "office order screen"
    Public Const frmNotification As String = "Notifications screen"
    Public Const ShareMaster As String = "Share Master"
    Public Const ShareAllotment As String = "Share Allotment"
    Public Const CostCenterGroupStore As String = "Cost Center"
    Public Const frmBullBreedMaster As String = "Bull Breed Master "
    Public Const frmBullTestParameter As String = "Bull Test Parameter "
    Public Const frmBullParameterGroup As String = "Bull Parameter Group"
    Public Const frmBullCurlingEntry As String = "Bull Curling Entry "
    Public Const frmBullCMUGrouping As String = "Bull CMU Grouping "

    Public Const Cleaning As String = "Cleaning"
    Public Const MilkCleaning As String = "MilkCleaning"
    Public Const GateOut As String = "Gate Out"
    Public Const ItemStockConversion As String = "Item Stock Conversion"
    Public Const TankerOut As String = "Tanker Out"
    Public Const MccMilkTransferPrice As String = "MCC Milk Transfer Price"
    Public Const DispatchBulkSaleTrade As String = "Bulk Dispatch Trade"
    Public Const DispatchBulkSaleTradeReturn As String = "Bulk Dispatch Trade Return"
    Public Const BulkMilkSRNTrade As String = "Bulk Milk SRN Trade"
    Public Const BulkMilkSRNTradeReturn As String = "Bulk Milk SRN Trade Return"
    Public Const LCRequest As String = "LC Request"
    Public Const LCCreation As String = "LC Creation"
    Public Const DocumentAcceptance As String = "Document Acceptance"
    ''richa 21/08/2014
    Public Const CatalogMaster As String = "Catalog Master"
    ''---------------------------------
    '' Anubhooti 07-Aug-2014 BM00000003405
    'Public Const OpenMCCShift As String = "Open MCC Shift"
    '' Anubhooti 01-Sep-2014 (HR:Applicant Entry)
    Public Const HRApplicantEntry As String = "HR Applicant Entry"
    '' Anubhooti 01-Sep-2014 (Vendor Master)''Balwinder not using these counter
    Public Const VendorMaster As String = "Vendor Master"
    Public Const CustomerMaster As String = "Customer Master"
    '' Anubhooti 16-Sep-2014 (VSP Asset Issue)
    Public Const VSPAssetIssue As String = "VSP Asset Issue"
    '' Anubhooti 16-Sep-2014 (VSP Asset Issue)
    Public Const VSPItemIssue As String = "VSP Item Issue"
    Public Const SupplierRegistration As String = "Supplier Registration"
    '--------Richa 08/09/2014 Against Ticket No BM00000003791
    Public Const VLCDataUploaderManual As String = "VLC Data Uploader Manual"
    Public Const MPIncentiveEntry As String = "MP Incentive Entrty"
    Public Const DBTMonthlyFarmerMilk As String = "DBT Monthly Farmer Milk"
    Public Const CattelFeedGRNQC As String = "Cattel Feed GRN QC"
    Public Const MPDCSIncentiveReco As String = "DCS MP Incentive Reco"
    Public Const DBTCaping As String = "DBT Caping"
    Public Const DBTNEFT As String = "DBT NEFT"
    Public Const DBTNEFTReject As String = "DBT NEFT Reject"
    ''--------------------------------------
    Public Const RequestForTraining As String = "Request For Training"

    'Public Const Trainer_Master As String = "Trainer Master"
    Public Const FixedAsset As String = "Fixed Asset"
    Public Const ProspectDetail As String = "Prospect"
    Public Const vlcDataUploader As String = "VLC Data Uploader"
    '========shivani
    Public Const EmployeeTransfer As String = "Employee Transfer"
    Public Const DamageDetail As String = "Damage/Fine Detail"
    '===========
    Public Const SWServiceCall As String = "Service Call"
    Public Const SWServiceEnquiry As String = "Service Enquiry"
    Public Const SWServiceAllocation As String = "Service Allocation"
    Public Const SWServiceVisitDetails As String = "Service Visit Details"
    '---preeti gupta

    Public Const HRRequisitionEntry As String = "HR REQUISITION ENTRY"
    Public Const HREMResignationLetter As String = "HR Resignation Letter "
    Public Const HREMTerminationLetter As String = "HR Termination Letter "
    Public Const Item_Conversion As String = "Item Conversion"
    Public Const HRTraineeFeedBack As String = "HR Trainee Feedback"

    Public Const HRTrainerFeedBack As String = "Hr Trainer Feedback"
    Public Const MRP As String = "MRP"
    Public Const CSABooking As String = "CSA Booking"
    Public Const CSARequest As String = "CSA Request"
    Public Const CSATransfer As String = "CSA Transfer"
    Public Const MILKTruckSheet As String = "Milk Truck Sheet"
    Public Const VLC_Target As String = "VLC Target"
    Public Const MilkRecurringSchedule As String = "Milk Recurring Schedule"
    Public Const VendorBankMaster As String = "Vendor Bank Master"
    Public Const IncentiveMaster As String = "Incentive Master"
    Public Const BankGuaranteeMaster As String = "Bank Guarantee Master"
    Public Const EmpMediclaim As String = "EMPMEDICM"
    Public Const EmployeeShiftChange As String = "Shift Change"
    Public Const PayrollSetting As String = "Payroll Setting"
    Public Const ExportIncentiveMaster As String = "EX INCENTIVE MASTER"
    Public Const ProvisionEntry As String = "Provision Entry"
    Public Const ProvisionEntryMilk As String = "Provision Entry Milk"
    Public Const DispatchTransfer As String = "Dispatch Transfer"
    Public Const ConveyanceClaim As String = "ConveyanceClaim"
    Public Const PaymentProcess As String = "Payment Process"
    Public Const ReviseMilkBill As String = "Revise Milk Bill"
    Public Const BankAdvise As String = "Bank Advise"
    Public Const PaymentProcessWithoutLoc As String = "Payment Process without Loc"
    Public Const AssetWork As String = "Asset Work"
    Public Const ItemIssueReturnAsset As String = "ItemIssueAsset"
    Public Const OpeningBankReco As String = "Opening Bank Reco"
    Public Const AssetDispatchRetailer As String = "Asset Dispatch Retailer"
    '=======Added by Preeti Gupta=========
    Public Const HRExitInterview As String = "HR Exit Interview"
    '====================End=====================
    Public Const UnclearedDoc As String = "Uncleared Document"
    Public Const RevaluationEntry As String = "Revaluation Entry"
    '===============Added by preeti Gupta============================
    Public Const frmCreateReceivedTransfer As String = "Transfer Crate Received"
    '==Added by kunal ===============================================================================
    Public Const EmptyWtEntry As String = "Empty Weight Entry"
    Public Const LoadedWtEntry As String = "Loaded Weight Entry"
    Public Const frmSilageProdApplicationForm As String = "Silage Application Form"
    Public Const frmSilageEnterPrenur As String = "Silage Entre Form"
    Public Const frmSilageFarmerSelection As String = "Silage Farmer Select"
    '===================================================================================================
    Public Const MilkGateEntryIn As String = "Milk Gate Entry In"
    Public Const MilkGateEntryOut As String = "Milk Gate Entry Out"
    Public Const MilkGateEntryWeighment As String = "Milk Gate Entry Weighment"
    Public Const PlantWeighment As String = "Plant Weighment"

    'stuti
    Public Const CAPEXMASTER As String = "CAPEX MASTER"
    Public Const CAPEXBUDGET As String = "CAPEX BUDGET"
    Public Const EMPLOYEEBANDMASTER As String = "EMPLOYEE BAND MASTER"
    Public Const VENREG As String = "VENDOR REGISTRATION"
    '====end here======

    'KUNAL > CLIENT : UDIL > TICKET : BM00000010226 
    Public Const NRGPREQUEST As String = "NRGP REQUEST"

    Public Const JobWorkDispatch As String = "Job Work Dispatch"
    Public Const JobWorkBilling As String = "Job Work Billing"
    Public Const JobWorkInvoice As String = "Job Work Invoice"
    ''Farmer PaymentStart Here
    Public Const frmMCCMaterialSaleFarmer As String = "Farmer Material Sale"
    Public Const frmMCCMaterialSaleInvoiceFarmer As String = "MCCMaterialSaleInvoiceFarmer"
    Public Const ARInvoiceFarmer As String = "AR Invoice Farmer"
    Public Const ARDebitNoteFarmer As String = "AR Debite Farmer"
    Public Const ARCreditNoteFarmer As String = "AR Credit Farmer"
    Public Const frmMCCMaterialSaleReturnFarmer As String = "Farmer Material Sale Return"
    Public Const FarmerPayment As String = "FarmerPayment"
    Public Const PaymentProcessFarmer As String = "PaymentProcessFarmer"
    Public Const FrmItemMasterRMOther As String = "Item Master"
    Public Const ContraVoucher As String = "Contra Voucher"
    Public Const PricePlan As String = "Price Plan"
    Public Const ItemWiseTax As String = "Item Wise Tax"
    Public Const SACWiseTax As String = "SAC Wise Tax"
    Public Const ItemCostMapping As String = "Item Cost Mapping"
    Public Const JobWorkTransferOther As String = "Job Work Transfer Other"
    Public Const JobWorkTransferOtherReturn As String = "Job Work Transfer Other Return"
    Public Const JobWorkTransferMilkReturn As String = "Job Work Transfer Milk Return"
    Public Const JWPriceCode As String = "Job Work Outward Price"
    Public Const DeductionMapping As String = "Deduction Mapping"
    Public Const JWIPriceCode As String = "Job Work Inward Price"
    Public Const QuickBookEntry As String = "Quick Book Entry"
    Public Const SubSidyNote As String = "SubSidyNote"

    'sanjay BHA/09/05/18-000014
    Public Const MaterialQuotation As String = "Material Quotation"
    Public Const MaterialQuotationOrder As String = "Material Quotation Order"
    Public Const MaterialQuotationComparison As String = "Material Quotation Comparison"

    Public Const LoanEntry As String = "Loan Entry"
    Public Const LoanInstallmentEntry As String = "Loan Installment"
    Public Const GeneralWeighment As String = "General Weighment"
    Public Const IncentiveEntry As String = "Incentive Entry"
    Public Const MakeSavingPayment As String = "Make Saving Payment"
    Public Const CustomerIncentiveEntry As String = "Customer Incentive Entry"
    Public Const Detail As String = "Detail"
    Public Const DetailSale As String = "Detail Sale"
    Public Const CustomerDeduction As String = "Customer Deduction"
    Public Const IncomeTaxSlab As String = "Income Tax Slab"
    Public Const IncomeTaxCalculation As String = "Income Tax Calculation"
    Public Const TankerCleaningItems As String = "Tanker Cleaning Items"
    Public Const CardSale As String = "Card Sale"
    'Eng. And Plant Management
    Public Const SectionMasterEng As String = "Section Master Eng"
    Public Const ConsumptionTypeMaster As String = "Consumption Type Master"
    Public Const LogSheetEng As String = "Log Sheet Eng"
    Public Const WorkRequisitionEng As String = "Work Requisition"
    Public Const WorkEstimationEng As String = "Work Estimation"
    Public Const WorkOrderEng As String = "Work Order"
    Public Const WorkOrderStatusEng As String = "Work Order Status"
    Public Const ThirdPartyReceipt As String = "Third Party Receipt"
    Public Const ThirdPartyCardSale As String = "Third Party Card Sale"
    Public Const ThirdPartyCSRecordHistory As String = "Third Party CS Record History"
    Public Const VSPMapping As String = "Mapping VSP"
    Public Const FarmerProMaster As String = "Farmer PRO Master"
    Public Const DCSAdditionDeduction As String = "DCS Addition Deduction"
    Public Const MPIncetiveSlab As String = "MP Incentive Slab"
    Public Const ChillingChargesSlab As String = "Chilling Charges Slab"
    Public Const OwnBMCGainLossRate As String = "Own BMC Gain Loss Rate"
    Public Const CappingMaster As String = "Capping Master"
    Public Const MatrixPriceChart As String = "Matrix Price Chart"
    Public Const MatrixPricePlan As String = "Matrix Price Plan"
    Public Const DSCFinancialHead As String = "DSC Financial Head"
    Public Const DSCFinancialEntry As String = "DSC Financial Entry"
    Public Const GazeReading As String = "Gaze Reading"
    Public Const Android As String = "Android"
    Public Const BlockMaster As String = "Block Master"
    Public Const RevenueVillageMaster As String = "Revenue Village Master"
    Public Const GrampanchayatMaster As String = "Grampanchayat Master"
    Public Const PanchayatSamitiMaster As String = "Panchayat Samiti Master"
    Public Const VidhanSabhaMaster As String = "Vidhan Sabha Master"
    Public Const UserRequestMaster As String = "User Request Master"
    Public Const MISItemMaster As String = "MIS Item Master"
    Public Const MISItemGroupMaster As String = "MIS Item Group Master"
    Public Const frmDailyMilkProducts As String = "Daily Milk Products"
    Public Const frmDailySMPProduction As String = "Daily SMP Production"
    Public Const frmBullMovement As String = "Bull Movement"
    Public Const CattleFeedPlantDaily As String = "Cattle Feed Plant"
    Public Shared Function SetDefaultValues() As Boolean
        Try



            Dim IsPOSeriesWithoutItemwise As Boolean = False
            IsPOSeriesWithoutItemwise = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.POSeriesWithoutItemTypewise, clsFixedParameterCode.POSeriesWithoutItemTypewise, Nothing)) = "1", True, False))
            Dim CreatVatSeriesOnExciseInvoice As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateVatSeriesForProductExciseinvoice, clsFixedParameterCode.CreateVatSeriesForProductExciseinvoice, Nothing))

            Dim qry As String = "delete from TSPL_DOCUMENT_TYPE"
            clsDBFuncationality.ExecuteNonQuery(qry)

            InsertDefaultValue(clsDocType.RCDFChecmicalQC, "", False, False)
            InsertDefaultValue(clsDocType.RCDFUnionQC, "", False, False)

            InsertDefaultValue(clsDocType.TenderPenalty, "", False, True)
            InsertDefaultValue(clsDocType.RALNoc, "", False, True)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.RequestService, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.MCCSaleRequest, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.MPSale, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.ComplainFeedback, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.DairyGateout, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.DairyDelivery, False, False)
            InsertDefaultValue(clsDocType.Android, clsDocTransactionType.DairyInvoiceAcknowledgment, False, False)


            InsertDefaultValue(clsDocType.CheckSlip, "", False, False)
            'InsertDefaultValue(clsDocType.CFormEntry, "", False)
            InsertDefaultValue(clsDocType.ExpiryTransaction, "", False, True)
            InsertDefaultValue(clsDocType.CHACHARGEMASTER, "", False, False)

            InsertDefaultValue(clsDocType.ARDebitNote, clsDocTransactionType.DebitRefDoc, False, True)
            InsertDefaultValue(clsDocType.ARCreditNote, clsDocTransactionType.CreditRefDoc, False, True)

            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.DebitRefDoc, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.CreditRefDoc, False, True)

            InsertDefaultValue(clsDocType.APInvoice, clsDocTransactionType.DirectAP, False, True)
            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.DirectAP, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.DirectAP, False, True)

            InsertDefaultValue(clsDocType.APInvoice, clsDocTransactionType.ServiceCharge, False, True)
            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.ServiceCharge, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.ServiceCharge, False, True)


            InsertDefaultValue(clsDocType.APInvoice, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.MccProc, False, True)

            InsertDefaultValue(clsDocType.APInvoice, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.BulkProc, False, True)

            InsertDefaultValue(clsDocType.APInvoice, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.DebitNote, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.CreditNote, clsDocTransactionType.GeneralPurchase, False, True)

            InsertDefaultValue(clsDocType.MultipleProcDed, "", False, True)
            InsertDefaultValue(clsDocType.TransferToSaving, "", False, True)
            InsertDefaultValue(clsDocType.BMCTransporterBill, "", False, True)
            InsertDefaultValue(clsDocType.ARInvoice, "", False, True)
            InsertDefaultValue(clsDocType.ARDebitNote, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.ARCreditNote, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.SupplementaryARInvoice, "", False, True)
            InsertDefaultValue(clsDocType.SupplementaryARCreditNote, "", False, True)
            InsertDefaultValue(clsDocType.ARServiceInvoice, "", False, True)
            InsertDefaultValue(clsDocType.ARServiceCreditNote, "", False, True)
            InsertDefaultValue(clsDocType.ARServiceDebitNote, "", False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.TransferOut, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.TransferLocalOut, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.TransferInterStateOut, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.TransferIn, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.TransferReject, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.InternalTransfer, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.InternalTransferOut, False, True)

            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, False, True)
            InsertDefaultValue(clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, False, True)
            InsertDefaultValue(clsDocType.LoadOut, clsDocTransactionType.TranferEmpty, False, True)
            InsertDefaultValue(clsDocType.LoadOut, clsDocTransactionType.TransferFull, False, True)
            InsertDefaultValue(clsDocType.LoadOut, clsDocTransactionType.TransferAutoRoute, False, True)
            InsertDefaultValue(clsDocType.LoadIn, clsDocTransactionType.TranferEmpty, False, True)
            InsertDefaultValue(clsDocType.LoadIn, clsDocTransactionType.TransferFull, False, True)
            InsertDefaultValue(clsDocType.PurchaserRegusitsion, clsDocTransactionType.OtherExternal, False, True)
            InsertDefaultValue(clsDocType.PurchaserRegusitsion, clsDocTransactionType.FinishedGoodExternal, False, True)
            InsertDefaultValue(clsDocType.frmStoreRequistion, clsDocTransactionType.FinishedGoodInternal, False, True)
            InsertDefaultValue(clsDocType.PurchaserRegusitsion, clsDocTransactionType.SemiFinishedGoodExternal, False, True)
            InsertDefaultValue(clsDocType.frmStoreRequistion, clsDocTransactionType.SemiFinishedGoodInternal, False, True)
            InsertDefaultValue(clsDocType.frmStoreRequistion, clsDocTransactionType.OtherInternal, False, True)
            InsertDefaultValue(clsDocType.PurchaserOrderOutward, clsDocTransactionType.POJobWorkOutward, False, True)
            'InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.frmProductionStoreRequest, clsDocTransactionType.FinishedGoodInternal, False, True)
            InsertDefaultValue(clsDocType.frmProductionStoreRequest, clsDocTransactionType.SemiFinishedGoodInternal, False, True)
            InsertDefaultValue(clsDocType.frmProductionStoreRequest, clsDocTransactionType.OtherInternal, False, True)
            If IsPOSeriesWithoutItemwise Then
                InsertDefaultValue(clsDocType.PurchaserOrder, "", False, True)
            Else
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POImport, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.PODomestic, False, True)

                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.DirectAP, False, True)


                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POJobWork, False, True)
                'InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POJobWorkOutward, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POFinishedGoods, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POSemiFinishedGoods, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POOther, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.PORawMaterial, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POAsset, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.PONonInvntry, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.POTrading, False, True)
                InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.All, False, True)
            End If
            InsertDefaultValue(clsDocType.JobWkDebit, "", False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.PODomestic, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POImport, False, True)

            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.RGPWise, False, True)

            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.GRN, clsDocTransactionType.All, False, True)

            'clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='" + clsDocTransactionType.NA + "'  where Doc_Type='" + clsDocType.POWeightment + "' and len(Doc_Trans_Type)<=0")
            InsertDefaultValue(clsDocType.POWeightment, clsDocTransactionType.NA, False, True)

            InsertDefaultValue(clsDocType.POWeightment, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.POWeightment, clsDocTransactionType.PODomestic, False, True)
            InsertDefaultValue(clsDocType.POWeightment, clsDocTransactionType.POImport, False, True)

            InsertDefaultValue(clsDocType.POWeightmentOG, clsDocTransactionType.POOutgoing, False, False)
            InsertDefaultValue(clsDocType.POWeightmentOG, clsDocTransactionType.RCDFLoadin, False, False)

            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.PODomestic, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POImport, False, True)

            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.RGPWise, False, True)

            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.MRN, clsDocTransactionType.All, False, True)

            InsertDefaultValue(clsDocType.TenderShortPenalty, "", False, True)

            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.PODomestic, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POImport, False, True)

            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.RGPWise, False, True)

            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.SRNRGP, False, True)
            InsertDefaultValue(clsDocType.POSCHEDULE, "", False, False)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.SRN, clsDocTransactionType.All, False, True)
            InsertDefaultValue(clsDocType.EPF, "", False, False)
            InsertDefaultValue(clsDocType.OutgoingProduction, "", False, False)
            InsertDefaultValue(clsDocType.RMProcessLoss, "", False, True)
            InsertDefaultValue(clsDocType.DAArrear, "", False, True)
            InsertDefaultValue(clsDocType.DAArrearNew, "", False, True)
            InsertDefaultValue(clsDocType.LeaveEncashment, "", False, True)
            InsertDefaultValue(clsDocType.SRNReturn, "", False, True)
            InsertDefaultValue(clsDocType.TransferReturn, "", False, True)
            InsertDefaultValue(clsDocType.GatePasstransfer, "", False, True)
            InsertDefaultValue(clsDocType.NIRQC, "", False, False)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.SRNRGP, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.MTSRN, clsDocTransactionType.POTrading, False, True)

            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.Mcc_TransporterInvoice, "", False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.POJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.POInvoice, clsDocTransactionType.All, False, True)

            InsertDefaultValue(clsDocType.PurchaseTaxInvoice, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.PurchaseTaxInvoice, clsDocTransactionType.MilkPurchase, False, True)
            InsertDefaultValue(clsDocType.PurchaseTaxInvoice, clsDocTransactionType.BulkMilkTaxInvoice, False, True)
            InsertDefaultValue(clsDocType.PurchaseTaxInvoice, clsDocTransactionType.DirectPurchaseFromAP, False, True)

            InsertDefaultValue(clsDocType.PurchaseBillOfSupply, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.PurchaseBillOfSupply, clsDocTransactionType.MilkPurchase, False, True)
            InsertDefaultValue(clsDocType.PurchaseBillOfSupply, clsDocTransactionType.BulkMilkTaxInvoice, False, True)
            InsertDefaultValue(clsDocType.PurchaseBillOfSupply, clsDocTransactionType.BulkMilkTaxInvoiceTrade, False, True)
            InsertDefaultValue(clsDocType.PurchaseBillOfSupply, clsDocTransactionType.DirectPurchaseFromAP, False, True)

            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.MT_POInvoice, clsDocTransactionType.All, False, True)

            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.POFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.POOther, False, True)
            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.POTrading, False, True)
            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.POAsset, False, True)
            InsertDefaultValue(clsDocType.PurchaseReturn, clsDocTransactionType.All, False, True)
            'InsertDefaultValue(clsDocType.GPEntry, "", False)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.Bank, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.Cash, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.PettyCash, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.Others, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.ApplyDocument, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.ApplyDoc, False, True)
            InsertDefaultValue(clsDocType.Payment, clsDocTransactionType.Tax, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Bank, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Tax, False, True)
            InsertDefaultValue(clsDocType.RICEMIXING, "", False, True)
            InsertDefaultValue(clsDocType.RICEPROCESSING, "", False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Cash, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.PettyCash, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Others, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.ApplyDocument, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Settlement, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.Normal, False, True)
            InsertDefaultValue(clsDocType.Receipt, clsDocTransactionType.ApplyDoc, False, True)
            InsertDefaultValue(clsDocType.RMDA, "", False, True)
            InsertDefaultValue(clsDocType.RGP, "", False, True)
            InsertDefaultValue(clsDocType.RGPR, "", False, True)
            InsertDefaultValue(clsDocType.NRGP, "", False, True)
            InsertDefaultValue(clsDocType.NRGPR, "", False, True)
            InsertDefaultValue(clsDocType.PJV, "", False, True)
            InsertDefaultValue(clsDocType.IssueReturn, clsDocTransactionType.ItemIssue, False, True)
            InsertDefaultValue(clsDocType.IssueReturn, clsDocTransactionType.ItemReturn, False, True)
            InsertDefaultValue(clsDocType.IssueReturn, clsDocTransactionType.ItemTransfer, False, True)
            InsertDefaultValue(clsDocType.IssueReturn, clsDocTransactionType.TransferCapex, False, True)
            InsertDefaultValue(clsDocType.IssueReturn, clsDocTransactionType.ItemIssue_Reprocess, False, True)
            InsertDefaultValue(clsDocType.SalesReturn, "", False, True)
            InsertDefaultValue(clsDocType.Scrap, "", False, True)
            ' InsertDefaultValue(clsDocType.ScrapReturn, "", False, True)
            'InsertDefaultValue(clsDocType.ScrapInvoice, "", False)
            InsertDefaultValue(clsDocType.DisposalEntry, "", False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.ScrapInvoiceRetail, False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.ScrapInvoiceTax, False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.SaleInvoiceExcise, False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.TaxExempted_ProductInvoice, False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.CashSale, False, True)

            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.GST_Interstate, False, True)


            If CreatVatSeriesOnExciseInvoice = 1 Then
                InsertDefaultValue(clsDocType.ScrapInvoice, clsDocTransactionType.SaleInvoiceInterstate, False, True)
            End If


            InsertDefaultValue(clsDocType.VCGLEntry, "", False, True)
            InsertDefaultValue(clsDocType.CreateRemmitance, "", False, False)
            InsertDefaultValue(clsDocType.frmDailyMilkProducts, "", False, False)
            InsertDefaultValue(clsDocType.frmDailySMPProduction, "", False, False)
            InsertDefaultValue(clsDocType.frmBullMovement, "", False, False)
            InsertDefaultValue(clsDocType.CattleFeedPlantDaily, "", False, True)
            'InsertDefaultValue(clsDocType.EmptyTransaction, clsDocTransactionType.EmptyTransactionOut, False)
            'InsertDefaultValue(clsDocType.EmptyTransaction, clsDocTransactionType.EmptyTransactionRouteIn, False)
            'InsertDefaultValue(clsDocType.EmptyTransaction, clsDocTransactionType.EmptyTransactionDepotIn, False)
            'InsertDefaultValue(clsDocType.EmptyTransaction, clsDocTransactionType.EmptyTransactionBySaleInvoice, False)
            'InsertDefaultValue(clsDocType.ProductionEntry, clsDocTransactionType.ProductionEntryFGTrading, False)
            'InsertDefaultValue(clsDocType.ProductionEntry, clsDocTransactionType.ProductionEntryFGManufacturing, False)
            InsertDefaultValue(clsDocType.StoreAdjustment, clsDocTransactionType.StoreAdjustmentAdjustment, False, True)
            InsertDefaultValue(clsDocType.StoreAdjustmentProductionEntry, clsDocTransactionType.StoreAdjustmentAdjustmentProductionEntry, False, True)
            InsertDefaultValue(clsDocType.StoreAdjustmentProductionEntryQC, clsDocTransactionType.StoreAdjustmentAdjustmentProductionEntryQC, False, True)
            InsertDefaultValue(clsDocType.StoreAdjustmentProductionStoreEntry, clsDocTransactionType.StoreAdjustmentAdjustmentProductionStoreEntry, False, True)
            InsertDefaultValue(clsDocType.RawMilkConsumtion, "", False, True)
            InsertDefaultValue(clsDocType.JobWorkInventory, "", False, True)

            InsertDefaultValue(clsDocType.warehouseBreakage, "", False, True)
            '-------------------------------------------------------------------------------------------------
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.JournalEntryJLJE, False, True)
            InsertDefaultValue(clsDocType.ReverseJournalEntry, clsDocTransactionType.JournalEntryJLJEReverseGeneral, False, True)
            InsertDefaultValue(clsDocType.ReverseJournalEntry, clsDocTransactionType.JournalEntryJLJEReverseMonthly, False, True)

            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.JournalEntryOther, False, True)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.JournalEntryPaymentProcess, False, True)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.JournalEntryMilkSRN, False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.JournalEntry, clsDocTransactionType.DirectAP, False, True)
            '---------------------------------------------------------------------------------------------
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.JournalEntryJLJE, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.JournalEntryOther, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.JournalEntryPaymentProcess, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.JournalEntryMilkSRN, False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.GeneralPurchase, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.JournalEntryOP, clsDocTransactionType.DirectAP, False, True)
            '------------------------------------------------------------------------------------------------


            InsertDefaultValue(clsDocType.ReveseNo, "", False, False)
            InsertDefaultValue(clsDocType.OTSheet, "", False, False)
            InsertDefaultValue(clsDocType.GeneralHolidays, "", False, True)
            InsertDefaultValue(clsDocType.Indent, "", False, False)
            InsertDefaultValue(clsDocType.LeaveAllotment, "", False, True)
            InsertDefaultValue(clsDocType.LeaveAdjustment, "", False, False)
            InsertDefaultValue(clsDocType.LeaveApplication, "", False, False)
            InsertDefaultValue(clsDocType.LoanAdjustment, "", False, False)
            InsertDefaultValue(clsDocType.LoanGeneration, "", False, False)
            InsertDefaultValue(clsDocType.ClaimMaster, "", False, False)

            InsertDefaultValue(clsDocType.Employee_Master, clsDocTransactionType.All, False, False)
            InsertDefaultValue(clsDocType.Employee_Master, clsDocTransactionType.PermanentBasis, False, False)
            InsertDefaultValue(clsDocType.Employee_Master, clsDocTransactionType.ContractBasis, False, False)
            InsertDefaultValue(clsDocType.Employee_Master, clsDocTransactionType.DailyBasis, False, False)

            InsertDefaultValue(clsDocType.Enquiry_Master, "", False, False)
            InsertDefaultValue(clsDocType.CSAPRICEMAASTER, "", False, False)
            InsertDefaultValue(clsDocType.CSADELIVERYORDER, "", False, True)
            InsertDefaultValue(clsDocType.PayHeadDefinitions, "", False, False)
            InsertDefaultValue(clsDocType.MonthlyAttendance, "", False, False)
            InsertDefaultValue(clsDocType.ReimbursementDetails, "", False, False)
            InsertDefaultValue(clsDocType.WeeklyHoliday, "", False, True)
            InsertDefaultValue(clsDocType.SalaryGeneration, "", False, False)
            InsertDefaultValue(clsDocType.SentMail, "", False, False)
            InsertDefaultValue(clsDocType.SalaryAdjustment, "", False, False)
            InsertDefaultValue(clsDocType.AllowanceDetails, "", False, False)
            InsertDefaultValue(clsDocType.ApplyLoan, "", False, False)
            InsertDefaultValue(clsDocType.LTAClaim, "", False, False)
            InsertDefaultValue(clsDocType.StandardRateItem, "", False, False)
            InsertDefaultValue(clsDocType.DeductionDetails, "", False, False)
            InsertDefaultValue(clsDocType.DailyAttendance, "", False, False)
            InsertDefaultValue(clsDocType.HourlyAttendance, "", False, False)
            InsertDefaultValue(clsDocType.GenerateBonus, "", False, False)
            InsertDefaultValue(clsDocType.SmsEmailScheduler, "", False, False)
            InsertDefaultValue(clsDocType.StandardScheme, "", False, False)
            InsertDefaultValue(clsDocType.TransferGateOut, "", False, False)
            InsertDefaultValue(clsDocType.ProductDispatchGateOut, "", False, False)
            InsertDefaultValue(clsDocType.MCCTankerGateOut, "", False, True)
            InsertDefaultValue(clsDocType.MCCTankerGateOutSecurity, "", False, True)
            InsertDefaultValue(clsDocType.ItemProjection, "", False, False)
            InsertDefaultValue(clsDocType.CSATransferGateOut, "", False, False)
            InsertDefaultValue(clsDocType.ScrapSaleGateOut, "", False, False)
            InsertDefaultValue(clsDocType.EmployeeStatus, "", False, False)
            InsertDefaultValue(clsDocType.EmployeeSalary, "", False, False)
            InsertDefaultValue(clsDocType.EmployeeIncrement, "", False, False)
            InsertDefaultValue(clsDocType.EmpIncrement, "", False, False)
            'InsertDefaultValue(clsDocType.FreshSale, "", False)
            InsertDefaultValue(clsDocType.BookingEntry, "", False, False)
            InsertDefaultValue(clsDocType.ContractTanker, "", False, False)
            InsertDefaultValue(clsDocType.GateReturnProductSale, "", False, False)
            InsertDefaultValue(clsDocType.GateReturnCSASale, "", False, False)
            InsertDefaultValue(clsDocType.GateReturnTransfer, "", False, False)
            InsertDefaultValue(clsDocType.VendorPriceMapDocwise, "", False, False)
            InsertDefaultValue(clsDocType.ReceiptInvoiceMapping, "", False, False)
            InsertDefaultValue(clsDocType.DeliveryNoteFreshSale, "", False, True)
            InsertDefaultValue(clsDocType.DispatchNoteFreshSale, "", False, True)
            InsertDefaultValue(clsDocType.frmCreateReceived, "", False, True)
            InsertDefaultValue(clsDocType.frmCreateReceivedAPI, "", False, True)
            InsertDefaultValue(clsDocType.frmCreateOut, "", False, True)
            InsertDefaultValue(clsDocType.frmCanReceived, "", False, True)
            InsertDefaultValue(clsDocType.frmCanReceivedAPI, "", False, True)
            InsertDefaultValue(clsDocType.GatePassEntrySale, "", False, True)
            InsertDefaultValue(clsDocType.MCCGateEntry, "", False, True)
            InsertDefaultValue(clsDocType.MCCWeighment, "", False, True)
            InsertDefaultValue(clsDocType.TDSPayment, "", False, True)
            InsertDefaultValue(clsDocType.frmInvoiceFreshSale, "", False, True)
            InsertDefaultValue(clsDocType.frmSaleReturnGateEntry, "", False, True)
            'InsertDefaultValue(clsDocType.frmInvoiceFreshSale, clsDocTransactionType.SaleInvoiceRetail, False, True)
            'InsertDefaultValue(clsDocType.frmInvoiceFreshSale, clsDocTransactionType.SaleInvoiceTax, False, True)
            ' InsertDefaultValue(clsDocType.frmSaleReturnFreshSale, "", False, True)
            clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Location_Code=''  where Doc_Type='" + clsDocType.FrmInvoiceCrateLinerDetail + "'")
            InsertDefaultValue(clsDocType.FrmInvoiceCrateLinerDetail, "", False, False)
            InsertDefaultValue(clsDocType.frmBookingProductSale, "", False, False)
            InsertDefaultValue(clsDocType.frmDispatchAdviceProductSale, "", False, True)
            InsertDefaultValue(clsDocType.frmSalesOrderProductSale, "", False, True)
            InsertDefaultValue(clsDocType.frmSalesOrderPSForExempted, "", False, True)
            InsertDefaultValue(clsDocType.frmDeliveryPrderProductSale, "", False, True)
            InsertDefaultValue(clsDocType.frmDeliveryOrderPSForExempted, "", False, True)

            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.Other, False, True)
            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.TaxExempted_ProductInvoice, False, True)

            'InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTNonTaxablePNR, False, True)
            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTInterstate, False, True)

            'InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.GSTNonTaxablePNR, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.GSTInterstate, False, True)

            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryLocal, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryInterstate, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryNonTaxable, False, True)

            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryCNoteLocal, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryCNoteInterstate, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SupplementaryCNoteNonTaxable, False, True)

            InsertDefaultValue(clsDocType.DCSDispatch, clsDocTransactionType.Other, False, True)
            InsertDefaultValue(clsDocType.DCSDispatch, clsDocTransactionType.TaxExempted_ProductInvoice, False, True)
            InsertDefaultValue(clsDocType.DCSInvoice, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.DCSInvoice, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.DCSInvoice, clsDocTransactionType.GSTInterstate, False, True)

            'InsertDefaultValue(clsDocType.frmTransferKDIl, clsDocTransactionType.GSTTaxable, False, True)
            'InsertDefaultValue(clsDocType.frmTransferKDIl, clsDocTransactionType.GSTNonTaxable, False, True)
            'InsertDefaultValue(clsDocType.frmTransferKDIl, clsDocTransactionType.GSTMandiTax, False, True)

            InsertDefaultValue(clsDocType.frmTransferGST, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.frmTransferGST, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.frmTransferGST, clsDocTransactionType.GSTBillofSupply, False, True)
            InsertDefaultValue(clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, False, True)
            InsertDefaultValue(clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, False, True) 'preeti gupta

            InsertDefaultValue(clsDocType.FrmGatePassFS, "", False, False)
            InsertDefaultValue(clsDocType.FrmGatePassPS, "", False, False)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceRetail, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceTax, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceExcise, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.TaxExempted_ProductInvoice, False, True)

            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceRetail, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceTax, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceExcise, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.TaxExempted_ProductInvoice, False, True)
            InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.CashSale, False, True)
            If CreatVatSeriesOnExciseInvoice = 1 Then
                InsertDefaultValue(clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceInterstate, False, True)
            End If
            InsertDefaultValue(clsDocType.frmDairySaleBooking, "", False, True)
            InsertDefaultValue(clsDocType.frmDairySaleBookingUploader, "", False, True)
            'added by preeti gupta===06/10/2016
            InsertDefaultValue(clsDocType.frmPerformaInvoiceBooking, "", False, True)
            InsertDefaultValue(clsDocType.frmDairySaleDeliveryOrder, "", False, True)
            InsertDefaultValue(clsDocType.frmDemandBooking, "", False, False, False, False, True)
            InsertDefaultValue(clsDocType.frmProductDemandBooking, "", False, False, False, False, True)
            InsertDefaultValue(clsDocType.frmDemandBookingUploader, "", False, False, False, False, True)
            InsertDefaultValue(clsDocType.frmDemandAdjustment, "", False, False)
            InsertDefaultValue(clsDocType.frmDemandUploader, "", False, False)
            InsertDefaultValue(clsDocType.FrmMCCMaterialSalePriceChart, "", False, False)
            InsertDefaultValue(clsDocType.frmDCSDemandBooking, "", False, True)
            InsertDefaultValue(clsDocType.frmDairySaleGatePass, "", False, True)
            InsertDefaultValue(clsDocType.FrmDairyGatePass, "", False, False)
            InsertDefaultValue(clsDocType.frmCustomerComplain, "", False, False)
            InsertDefaultValue(clsDocType.frmTranspoterDeductionEntry, "", False, False)
            InsertDefaultValue(clsDocType.frmLeakedSaleReturn, "", False, True)


            InsertDefaultValue(clsDocType.rptMonthlySalesInvoice, "", False, False)
            '--------Richa 21/07/2014 Against Ticket No BM00000003245
            InsertDefaultValue(clsDocType.BulkSalePriceChart, "", False, False)
            '-----------------------------------------------------
            '--------Richa 28/07/2014 Against Ticket No BM00000003157
            InsertDefaultValue(clsDocType.WeighmentEntryBulkSale, "", False, True)
            '-----------------------------------------------------
            '--------Richa 28/07/2014 Against Ticket No BM00000003158
            InsertDefaultValue(clsDocType.LoadingTankerBulkSale, "", False, True)
            '-----------------------------------------------------
            InsertDefaultValue(clsDocType.QualityCheckBulkSale, "", False, True)
            InsertDefaultValue(clsDocType.DispatchBulkSale, "", False, True)
            InsertDefaultValue(clsDocType.AcknowledgementGRN, "", False, True)
            InsertDefaultValue(clsDocType.CustomerOutstanding, "", False, True)
            InsertDefaultValue(clsDocType.SiloMilkTransfer, "", False, True)
            InsertDefaultValue(clsDocType.InvoiceBulkSale, clsDocTransactionType.BULKMilkSale, False, True)
            InsertDefaultValue(clsDocType.InvoiceBulkSale, clsDocTransactionType.BULKMilkSaleTrade, False, True)
            InsertDefaultValue(clsDocType.InvoiceBulkSale, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.ProformaInvoiceBulkSale, "", False, True)
            'InsertDefaultValue(clsDocType.BulkSaleReturn, "", False, True)
            InsertDefaultValue(clsDocType.MerchantTradePO, "", False, True)
            InsertDefaultValue(clsDocType.SalesOrderBulkSale, "", False, True)
            ''richa 01/04/2015
            InsertDefaultValue(clsDocType.BulkSaleReturnDispatch, "", False, True)
            ''-------------
            ''richa 21/08/2014
            InsertDefaultValue(clsDocType.CatalogMaster, "", False, False)
            '--------Created By Richa 08/09/2014 Against Ticket No BM00000003791
            InsertDefaultValue(clsDocType.VLCDataUploaderManual, "", False, True)
            InsertDefaultValue(clsDocType.MPIncentiveEntry, "", False, True)
            InsertDefaultValue(clsDocType.DBTMonthlyFarmerMilk, "", False, False)
            InsertDefaultValue(clsDocType.CattelFeedGRNQC, "", False, False)
            InsertDefaultValue(clsDocType.DBTCaping, "", False, False)
            InsertDefaultValue(clsDocType.MPDCSIncentiveReco, "", False, False)
            InsertDefaultValue(clsDocType.DBTNEFT, "", False, False)
            InsertDefaultValue(clsDocType.DBTNEFTReject, "", False, False)
            ''=========================
            InsertDefaultValue(clsDocType.FixedDeposit, "", False, False)
            InsertDefaultValue(clsDocType.LCRequest, "", False, True)
            InsertDefaultValue(clsDocType.LCCreation, "", False, True)
            InsertDefaultValue(clsDocType.DocumentAcceptance, "", False, True)
            InsertDefaultValue(clsDocType.DispatchBulkSaleTrade, "", False, True)
            InsertDefaultValue(clsDocType.DispatchBulkSaleTradeReturn, "", False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRNTrade, "", False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRNTradeReturn, "", False, True)
            InsertDefaultValue(clsDocType.SNSalesQuotation, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.SNSalesQuotation, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.SNSalesQuotation, "", False, True)
            'InsertDefaultValue(clsDocType.SNSalesOrder, "", False, True)
            InsertDefaultValue(clsDocType.SNSalesOrder, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.SNSalesOrder, clsDocTransactionType.SNQuotationOther, False, True)

            InsertDefaultValue(clsDocType.SNShipment, "", False, True)

            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceRetail, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceTax, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceService, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceExcise, False, True)
            If CreatVatSeriesOnExciseInvoice = 1 Then
                InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceInterstate, False, True)
            End If

            'InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.GSTInterstate, False, True)
            InsertDefaultValue(clsDocType.SNSaleInvoice, clsDocTransactionType.GST_Interstate, False, True)

            InsertDefaultValue(clsDocType.JobWorkDispatch, "", False, True)
            InsertDefaultValue(clsDocType.JobWorkInvoice, "", False, True)
            InsertDefaultValue(clsDocType.JobWorkBilling, "", False, True)
            '=====================export======================================
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXSALESQUOTATION, "", False, True)

            InsertDefaultValue(clsDocType.EXSALESORDER, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXSALESORDER, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESORDER, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXSALESORDER, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESORDER, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXSALESORDER, "", False, True)

            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXPORFORMAINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXCOMMERSIALINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.EXSALESINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXSALESINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXSALESINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXSALESINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.EXSALESRETURN, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.EXSALESRETURN, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESRETURN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.EXSALESRETURN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.EXSALESRETURN, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.EXSALESRETURN, "", False, True)
            '================end here==============================================================

            '==========================Merchant Trade=====================================================
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, clsDocTransactionType.SNOrderMerchant, False, True)
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTSALESQUOTATION, "", False, True)

            InsertDefaultValue(clsDocType.MTSALESORDER, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.MTSALESORDER, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESORDER, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTSALESORDER, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESORDER, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTSALESORDER, "", False, True)

            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTPORFORMAINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTCOMMERSIALINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.MTSALESINVOICE, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.MTSALESINVOICE, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESINVOICE, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTSALESINVOICE, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESINVOICE, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTSALESINVOICE, "", False, True)

            InsertDefaultValue(clsDocType.MTSALESRETURN, clsDocTransactionType.SNOrderExport, False, True)
            InsertDefaultValue(clsDocType.MTSALESRETURN, clsDocTransactionType.SNQuotationFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESRETURN, clsDocTransactionType.PORawMaterial, False, True)
            InsertDefaultValue(clsDocType.MTSALESRETURN, clsDocTransactionType.POSemiFinishedGoods, False, True)
            InsertDefaultValue(clsDocType.MTSALESRETURN, clsDocTransactionType.SNQuotationOther, False, True)
            InsertDefaultValue(clsDocType.MTSALESRETURN, "", False, True)
            '===========================end here=================================================

            InsertDefaultValue(clsDocType.CSASaleInvoice, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.CSASaleInvoice, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.CSASaleInvoice, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.SNSaleReturn, "", False, True)
            ' InsertDefaultValue(clsDocType.CSASALERETURN, "", False, True)
            InsertDefaultValue(clsDocType.SNPOS, "", False, True)
            InsertDefaultValue(clsDocType.BarCode, "", False, False)
            InsertDefaultValue(clsDocType.ShipToMaster, "", False, False)
            'InsertDefaultValue(clsDocType.BOM, "", False, True)
            InsertDefaultValue(clsDocType.BOM, clsDocTransactionType.SNQuotationOther, False, False)
            InsertDefaultValue(clsDocType.BOM, clsDocTransactionType.BOMOSPTYPE, False, False)
            InsertDefaultValue(clsDocType.BOM, clsDocTransactionType.NA, False, False)
            InsertDefaultValue(clsDocType.DairyProductionUploader, "", False, True)
            InsertDefaultValue(clsDocType.ShiftMgmt, "", False, True)
            InsertDefaultValue(clsDocType.ShiftMgmtSFG, "", False, True)
            InsertDefaultValue(clsDocType.VendorQuotation, "", False, False)
            InsertDefaultValue(clsDocType.SetPOSchedule, "", False, False)
            InsertDefaultValue(clsDocType.RFQ, "", False, False)
            InsertDefaultValue(clsDocType.RMDemandApproval, "", False, False)
            InsertDefaultValue(clsDocType.BATCHORDER, "", False, True)
            InsertDefaultValue(clsDocType.NOTIFYPARTY, "", False, False)
            InsertDefaultValue(clsDocType.PPISSUEENTRY, "", False, True)
            InsertDefaultValue(clsDocType.PPSTANDARDIZATION, "", False, True)
            InsertDefaultValue(clsDocType.PPSTDFinalQC, "", False, True)
            InsertDefaultValue(clsDocType.PPStageProcess, "", False, True)
            InsertDefaultValue(clsDocType.ppProductionEntry, "", False, True)
            InsertDefaultValue(clsDocType.ppSTProductionEntry, "", False, True)
            InsertDefaultValue(clsDocType.ProductionEntryFinalQC, "", False, True)
            InsertDefaultValue(clsDocType.PPProductionReturn, clsDocTransactionType.PPReturnProductionEntry, False, True)
            InsertDefaultValue(clsDocType.PPProductionReturn, clsDocTransactionType.PPReturnProductionStandardization, False, True)
            InsertDefaultValue(clsDocType.WreckageBooking, "", False, True)
            InsertDefaultValue(clsDocType.PPLOGSHEET, "", False, False)
            InsertDefaultValue(clsDocType.MO, "", False, False)
            InsertDefaultValue(clsDocType.ProductionRequisition, "", False, False)
            InsertDefaultValue(clsDocType.ProductionIssue, "", False, False)
            InsertDefaultValue(clsDocType.ProductionMapping, "", False, False)
            InsertDefaultValue(clsDocType.ProductionReturn, "", False, False)
            InsertDefaultValue(clsDocType.ProductionReceipt, "", False, False)
            InsertDefaultValue(clsDocType.ProductionPlanning, "", False, True)
            InsertDefaultValue(clsDocType.StandardProductionPlanning, "", False, False)
            InsertDefaultValue(clsDocType.SerializedReplace, "", False, False)
            InsertDefaultValue(clsDocType.AcquisitionEntry, "", False, True)
            InsertDefaultValue(clsDocType.OutputEntry, "", False, True)
            InsertDefaultValue(clsDocType.FAAssembleAsset, "", False, True)
            InsertDefaultValue(clsDocType.AssetAccountChange, "", False, True)
            InsertDefaultValue(clsDocType.QCVendorItemMapping, "", False, False)
            InsertDefaultValue(clsDocType.frmOutgoingQCEntry, clsDocTransactionType.ProductionQC, False, True)
            InsertDefaultValue(clsDocType.QualityCheckForSRN, clsDocTransactionType.IncomingQualityCheck, False, True)
            InsertDefaultValue(clsDocType.QualityCheckForSRN, clsDocTransactionType.InprocessQualityCheck, False, True)
            InsertDefaultValue(clsDocType.QualityCheckForSRN, clsDocTransactionType.OutgoingQualityCheck, False, True)
            InsertDefaultValue(clsDocType.AssetDepreciation, "", False, False)
            InsertDefaultValue(clsDocType.AssetRequisition, "", False, True)
            InsertDefaultValue(clsDocType.Assemblies, "", False, False)
            InsertDefaultValue(clsDocType.AssetAgreement, "", False, False)
            InsertDefaultValue(clsDocType.AssetDispatch, "", False, True)
            InsertDefaultValue(clsDocType.complaintdetail, "", False, False)
            InsertDefaultValue(clsDocType.ComplaintEntry, "", False, True)
            InsertDefaultValue(clsDocType.CartMaintenance, "", False, False)
            InsertDefaultValue(clsDocType.CollectionCenter, "", False, False)
            'InsertDefaultValue(clsDocType.VSPMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.MilkCollectionMCCMuliple, "", False, False, False)
            InsertDefaultValue(clsDocType.MilkCollectionMCC, "", False, False, False)
            InsertDefaultValue(clsDocType.MilkShiftUploader, "", False, True, False)
            InsertDefaultValue(clsDocType.MilkCollectionDCSMuliple, "", False, False, False)
            InsertDefaultValue(clsDocType.MilkCollectionDCSMulipleMerge, "", False, False, False)
            InsertDefaultValue(clsDocType.MilkCollectionDCS, "", False, False, False)
            InsertDefaultValue(clsDocType.CustomerPenalty, "", False, False, False)
            InsertDefaultValue(clsDocType.CFSalePurchaseUploader, "", False, False, False)
            InsertDefaultValue(clsDocType.VSPMASTER, clsDocTransactionType.PDCS, False, False, True)
            InsertDefaultValue(clsDocType.VSPMASTER, clsDocTransactionType.Registered, False, False, True)
            InsertDefaultValue(clsDocType.PTMMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.TTMMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.VLCMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.MILKPRCMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.TankerDispatchPriceMaster, "", False, False, True)
            InsertDefaultValue(clsDocType.ListofPendingDocuments, "", False, False)
            InsertDefaultValue(clsDocType.MCCMaster, clsDocTransactionType.BMCU, True, True, True)
            InsertDefaultValue(clsDocType.MCCMaster, clsDocTransactionType.MCC, True, True, True)

            InsertDefaultValue(clsDocType.MilkReceipt, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkReject, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkSample, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkSRN, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkPO, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkPurInvoice, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkShift_End, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)
            InsertDefaultValue(clsDocType.MilkPurInvoiceProvsion, "", False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)


            InsertDefaultValue(clsDocType.FrmGrievanceLogging, "", False, False)
            InsertDefaultValue(clsDocType.VLCTarget, "", False, True)
            InsertDefaultValue(clsDocType.MilkTransportorInvoice, "", False, True)

            InsertDefaultValue(clsDocType.VSPIncentiveTagging, "", False, True)
            InsertDefaultValue(clsDocType.MCCMaterialSale, "", False, True)
            ' InsertDefaultValue(clsDocType.MCCMaterialSaleReturn, "", False, True)
            InsertDefaultValue(clsDocType.Sch_Training, "", False, False)

            InsertDefaultValue(clsDocType.MPMaster, "", False, False, True)
            InsertDefaultValue(clsDocType.VLCROUTESHIFT, "", False, False)
            InsertDefaultValue(clsDocType.CustomerRouteShift, "", False, False)
            InsertDefaultValue(clsDocType.MILKROUTEMASTER, "", False, False, True)
            InsertDefaultValue(clsDocType.VILLAGEMASTER, "", False, False, True)

            clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='" + clsDocTransactionType.GSTNA + "' where Doc_Type='" + clsDocType.MccDispatchChallan + "' and Doc_Trans_Type=''")
            InsertDefaultValue(clsDocType.MccDispatchChallan, clsDocTransactionType.GSTNA, False, True)
            InsertDefaultValue(clsDocType.MccDispatchChallan, clsDocTransactionType.GSTLocal, False, True)
            InsertDefaultValue(clsDocType.MccDispatchChallan, clsDocTransactionType.GSTInterstate, False, True)


            InsertDefaultValue(clsDocType.MccDispatchChallanReturn, "", False, True)
            InsertDefaultValue(clsDocType.MccTnkDispChallanReturn, "", False, True)
            InsertDefaultValue(clsDocType.ParamMaster, "", False, False)
            InsertDefaultValue(clsDocType.QCLOGSHEETMST, "", False, False)
            InsertDefaultValue(clsDocType.AcknowledgementEntry, "", False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.MCCProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.GateEntry, clsDocTransactionType.ReferenceNo, False, False)

            InsertDefaultValue(clsDocType.POBulkP, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.POBulkP, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.POBulkP, clsDocTransactionType.BulkProc, False, True)

            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.Weighment, clsDocTransactionType.MCCProcJobWorkOutward, False, True)

            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.MccProc, False, True)
            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.QualityCheck, clsDocTransactionType.MCCProcJobWorkOutward, False, True)

            InsertDefaultValue(clsDocType.SecondarySettingForQC, clsDocTransactionType.BulkProc, False, True)
            InsertDefaultValue(clsDocType.IntimationScreen, "", False, True)
            InsertDefaultValue(clsDocType.PriceChartMasterBulk, "", False, False)
            InsertDefaultValue(clsDocType.AdjustmentEntry, "", False, False)
            InsertDefaultValue(clsDocType.PaymentAdjustmentEntry, "", False, False)

            clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='" + clsDocTransactionType.NA + "'  where Doc_Type='" + clsDocType.Unloading + "' and len(Doc_Trans_Type)<=0")
            InsertDefaultValue(clsDocType.Unloading, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.Unloading, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.Unloading, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.Unloading, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.Unloading, clsDocTransactionType.MCCProcJobWorkOutward, False, True)


            InsertDefaultValue(clsDocType.MilkTransferIn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.MilkTransferIn, clsDocTransactionType.MCCProcJobWorkOutward, False, True)


            InsertDefaultValue(clsDocType.MilkTransferInReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.MilkTransferInReturn, clsDocTransactionType.MCCProcJobWorkOutward, False, True)

            clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='" + clsDocTransactionType.NA + "'  where Doc_Type='" + clsDocType.BulkMilkSRN + "' and len(Doc_Trans_Type)<=0")
            InsertDefaultValue(clsDocType.BulkMilkSRN, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRN, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRN, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRN, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.frmMilkJobWorkTransfer, "", False, True)


            'InsertDefaultValue(clsDocType.JobMilkSRN, "", False, True)
            InsertDefaultValue(clsDocType.BulkMilkSRNReturn, "", False, True)
            InsertDefaultValue(clsDocType.BulkMilkPO, "", False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseInvoice, clsDocTransactionType.WithoutVendorInvoiceNo, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseInvoice, clsDocTransactionType.WithVendorInvoiceNo, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseInvoice, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseInvoiceTrade, clsDocTransactionType.WithoutVendorInvoiceNo, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseInvoiceTrade, clsDocTransactionType.WithVendorInvoiceNo, False, True)

            ' Add By Prabhakar 
            InsertDefaultValue(clsDocType.BulkMilkPurchaseReturnTrade, clsDocTransactionType.WithoutVendorInvoiceNo, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseReturnTrade, clsDocTransactionType.WithVendorInvoiceNo, False, True)
            'InsertDefaultValue(clsDocType.BulkMilkPurchaseReturn, clsDocTransactionType.WithoutVendorInvoiceNo, False, True)
            'InsertDefaultValue(clsDocType.BulkMilkPurchaseReturn, clsDocTransactionType.WithVendorInvoiceNo, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.BulkMilkPurchaseReturn, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            'FarmerServiceOrder
            InsertDefaultValue(clsDocType.FarmerServiceOrder, "", False, False)

            InsertDefaultValue(clsDocType.DCSSale, "", False, True, True)
            InsertDefaultValue(clsDocType.DistributorCommission, "", False, False, True)
            InsertDefaultValue(clsDocType.BoothCommission, "", False, False, True)
            InsertDefaultValue(clsDocType.BoothCalculation, "", False, False, True)
            InsertDefaultValue(clsDocType.frmCustomerTender, "", False, False, True)
            InsertDefaultValue(clsDocType.frmCustomerTenderOrder, "", False, False, True)
            InsertDefaultValue(clsDocType.frmCustomerTenderDispatch, "", False, False, True)
            InsertDefaultValue(clsDocType.frmCustomerTenderInvoice, "", False, False, True)
            InsertDefaultValue(clsDocType.frmCustomerTenderGatePass, "", False, False, True)
            InsertDefaultValue(clsDocType.ItemCapacityLimit, "", False, False, True)
            InsertDefaultValue(clsDocType.BoothRouteMapping, "", False, False, True)
            InsertDefaultValue(clsDocType.DistributorRouteTagging, "", False, False, True)
            InsertDefaultValue(clsDocType.frmNotification, "", False, True, True)

            'sanjay BHA/09/05/18-000014 
            InsertDefaultValue(clsDocType.MaterialQuotation, "", False, True)
            InsertDefaultValue(clsDocType.MaterialQuotationOrder, "", False, True)
            InsertDefaultValue(clsDocType.MaterialQuotationComparison, "", False, True)
            'sanjay BHA/09/05/18-000014 

            InsertDefaultValue(clsDocType.LoanEntry, clsDocTransactionType.LoanReceipt, False, False)
            InsertDefaultValue(clsDocType.LoanEntry, clsDocTransactionType.LoanPayment, False, False)

            InsertDefaultValue(clsDocType.LoanInstallmentEntry, clsDocTransactionType.LoanReceipt, False, False)
            InsertDefaultValue(clsDocType.LoanInstallmentEntry, clsDocTransactionType.LoanPayment, False, False)

            InsertDefaultValue(clsDocType.ThirdPartyReceipt, "", False, False)
            InsertDefaultValue(clsDocType.ThirdPartyCardSale, "", False, False)
            InsertDefaultValue(clsDocType.ThirdPartyCSRecordHistory, "", False, False)

            '--------Sanjay  Against Ticket No BHA/08/08/18-000397 
            InsertDefaultValue(clsDocType.GeneralWeighment, "", False, True)
            '-----------------------------------------------------
            InsertDefaultValue(clsDocType.JWEstimate, "", False, True)
            InsertDefaultValue(clsDocType.JWVendorFormula, "", False, False, True)
            InsertDefaultValue(clsDocType.JWInwardFormula, "", False, False, True)

            InsertDefaultValue(clsDocType.JWOFormula, "", False, False, True)
            InsertDefaultValue(clsDocType.JWOVendorFormula, "", False, False, True)
            InsertDefaultValue(clsDocType.JWOEstimate, "", False, True)

            If (IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowAutoGenerateDocNoInMaster, clsFixedParameterCode.AllowAutoGenerateDocNoInMaster, Nothing)) = "1", True, False) = True) Then
                InsertDefaultValue(clsDocType.CostCenterGroupStore, "", False, False)
                InsertDefaultValue(clsDocType.AdditionalCharges, "", False, False)
                InsertDefaultValue(clsDocType.AreaMaster, "", False, False)
                InsertDefaultValue(clsDocType.BankMaster, "", False, False)
                InsertDefaultValue(clsDocType.CityMaster, "", False, False)
                InsertDefaultValue(clsDocType.DistrictMaster, "", False, False)
                InsertDefaultValue(clsDocType.RegionMaster, "", False, False)
                InsertDefaultValue(clsDocType.BlockMaster, "", False, False)
                InsertDefaultValue(clsDocType.RevenueVillageMaster, "", False, False)
                InsertDefaultValue(clsDocType.GrampanchayatMaster, "", False, False)
                InsertDefaultValue(clsDocType.VidhanSabhaMaster, "", False, False)
                InsertDefaultValue(clsDocType.CompanyMaster, "", False, False)
                InsertDefaultValue(clsDocType.CurrencyConversion, "", False, False)
                InsertDefaultValue(clsDocType.DesignationMaster, "", False, False)
                InsertDefaultValue(clsDocType.FormMaster, "", False, False)
                InsertDefaultValue(clsDocType.PaymentCode, "", False, False)
                InsertDefaultValue(clsDocType.PaymentTerms, "", False, False)
                InsertDefaultValue(clsDocType.TaxAuthority, "", False, False)
                InsertDefaultValue(clsDocType.TaxGroup, "", False, False)
                InsertDefaultValue(clsDocType.TaxRates, "", False, False)
                InsertDefaultValue(clsDocType.TimeTable, "", False, False)
                InsertDefaultValue(clsDocType.AcquisitionCodes, "", False, False)
                InsertDefaultValue(clsDocType.HirerachyLevelMaster, "", False, False)
                InsertDefaultValue(clsDocType.VendorAccountSet, "", False, False)
                InsertDefaultValue(clsDocType.VendorGroup, "", False, False)
                InsertDefaultValue(clsDocType.VendorType, "", False, False)
                InsertDefaultValue(clsDocType.CostCenter, "", False, False)
                InsertDefaultValue(clsDocType.TankerMasterSale, "", False, False)
                InsertDefaultValue(clsDocType.DispatchCheckList, "", False, False)
                InsertDefaultValue(clsDocType.CustomerType, "", False, False)
                InsertDefaultValue(clsDocType.CustomerAccountSet, "", False, False)
                InsertDefaultValue(clsDocType.RouteTypeMaster, "", False, False)
                InsertDefaultValue(clsDocType.VehicleMaster, "", False, False)
                InsertDefaultValue(clsDocType.TransportType, "", False, False)
                InsertDefaultValue(clsDocType.ChannelCategory, "", False, False)
                InsertDefaultValue(clsDocType.ChannelMaster, "", False, False)
                InsertDefaultValue(clsDocType.SalesManHierarchy, "", False, False)
                InsertDefaultValue(clsDocType.SettlementMaster, "", False, False)
                InsertDefaultValue(clsDocType.RouteGroupMaster, "", False, False)
                InsertDefaultValue(clsDocType.DiscountCategoryMaster, "", False, False)
                InsertDefaultValue(clsDocType.SamplingMaster, "", False, False)
                InsertDefaultValue(clsDocType.TransportMaster, "", False, False)
                InsertDefaultValue(clsDocType.RemarkMaster, "", False, False)
                InsertDefaultValue(clsDocType.DiscountMaster, "", False, False)
                InsertDefaultValue(clsDocType.BulkSalePriceChart, "", False, False)
                InsertDefaultValue(clsDocType.ZoneMaster, "", False, False)
                InsertDefaultValue(clsDocType.RouteMaster, "", False, False)
                InsertDefaultValue(clsDocType.SalesLevelHierarchy, "", False, False)
                InsertDefaultValue(clsDocType.MerchantPaymentTerms, "", False, False)
                InsertDefaultValue(clsDocType.RequisitSubTypeMaster, "", False, False)
                InsertDefaultValue(clsDocType.CustomerCategory, "", False, False)
                InsertDefaultValue(clsDocType.ItemCategoryLevel, "", False, False)
                InsertDefaultValue(clsDocType.ItemCategoryStructure, "", False, False)
                InsertDefaultValue(clsDocType.ReceivablePaymentTerms, "", False, False)
                InsertDefaultValue(clsDocType.ShipToLocationDetails, "", False, False)
                InsertDefaultValue(clsDocType.SecondaryCustomerMaster, "", False, False)
                InsertDefaultValue(clsDocType.AccountMainGroup, "", False, False)
                InsertDefaultValue(clsDocType.AccountGroup, "", False, False)
                InsertDefaultValue(clsDocType.AccountSubGroup, "", False, False)
                InsertDefaultValue(clsDocType.MainGLAccount, "", False, False)
                InsertDefaultValue(clsDocType.GLAccount, "", False, False)
                InsertDefaultValue(clsDocType.CostCentreFinancial, "", False, False)
                InsertDefaultValue(clsDocType.UserMaster, "", False, False)
                InsertDefaultValue(clsDocType.UserGroupMaster, "", False, False)
                InsertDefaultValue(clsDocType.ItemStructure, "", False, False)
                InsertDefaultValue(clsDocType.PurchaseAccountSetCode, "", False, False)
                InsertDefaultValue(clsDocType.ChapterHead, "", False, False)
                InsertDefaultValue(clsDocType.ItemCategory, "", False, False)
                InsertDefaultValue(clsDocType.ItemSubCategory, "", False, False)
                InsertDefaultValue(clsDocType.PriceComponantMaster, "", False, False)
                InsertDefaultValue(clsDocType.PriceComponantMapping, "", False, False)
                InsertDefaultValue(clsDocType.BreakageHead, "", False, False)
                InsertDefaultValue(clsDocType.Warranty, "", False, False)
                InsertDefaultValue(clsDocType.InventorySourceCode, "", False, False)
                InsertDefaultValue(clsDocType.ItemType, "", False, False)
                InsertDefaultValue(clsDocType.ResponsiblePerson, "", False, False)
                InsertDefaultValue(clsDocType.BranchDetails, "", False, False)
                InsertDefaultValue(clsDocType.NatureOfDeduction, "", False, False)
                InsertDefaultValue(clsDocType.PartyDetails, "", False, False)
                InsertDefaultValue(clsDocType.TDSSection, "", False, False)
                InsertDefaultValue(clsDocType.StateCode, "", False, False)
                InsertDefaultValue(clsDocType.Category, "", False, False)
                InsertDefaultValue(clsDocType.DeprAccountSet, "", False, False)
                InsertDefaultValue(clsDocType.FACostCenter, "", False, False)
                InsertDefaultValue(clsDocType.AssetGroups, "", False, False)
                InsertDefaultValue(clsDocType.AssetBook, "", False, False)
                InsertDefaultValue(clsDocType.DepreciationMethod, "", False, False)
                'InsertDefaultValue(clsDocType.JWInwardFormula, "", False, False)

                InsertDefaultValue(clsDocType.DepreciationPeriods, "", False, False)
                InsertDefaultValue(clsDocType.TemplateMaster, "", False, False)
                InsertDefaultValue(clsDocType.SalaryAccountSetting, "", False, False)
                InsertDefaultValue(clsDocType.CountryMaster, "", False, False)
                InsertDefaultValue(clsDocType.StateMaster, "", False, False)
                InsertDefaultValue(clsDocType.BranchMaster, "", False, False)
                InsertDefaultValue(clsDocType.OTMaster, "", False, False)
                InsertDefaultValue(clsDocType.AttendanceMaster, "", False, False)
                InsertDefaultValue(clsDocType.BonusMaster, "", False, False)
                InsertDefaultValue(clsDocType.PayPeriodMaster, "", False, False)
                InsertDefaultValue(clsDocType.PFRulesMaster, "", False, False)
                InsertDefaultValue(clsDocType.ESIRulesMaster, "", False, False)
                InsertDefaultValue(clsDocType.LeaveMaster, "", False, False)
                InsertDefaultValue(clsDocType.ShiftMaster, "", False, False)
                InsertDefaultValue(clsDocType.DepartmentMaster, "", False, False)
                InsertDefaultValue(clsDocType.DevisionMaster, "", False, False)
                InsertDefaultValue(clsDocType.CastCategory, "", False, False)
                InsertDefaultValue(clsDocType.DocumentMaster, "", False, False)
                InsertDefaultValue(clsDocType.CourseMaster, "", False, False)
                InsertDefaultValue(clsDocType.GradeMaster, "", False, False)
                InsertDefaultValue(clsDocType.LanguageMaster, "", False, False)
                InsertDefaultValue(clsDocType.OccupationMaster, "", False, False)
                InsertDefaultValue(clsDocType.ReligionMaster, "", False, False)
                InsertDefaultValue(clsDocType.SkillMaster, "", False, False)
                InsertDefaultValue(clsDocType.OTSLab, "", False, False)
                InsertDefaultValue(clsDocType.PTSLab, "", False, False)
                InsertDefaultValue(clsDocType.HRRSLab, "", False, False)
                InsertDefaultValue(clsDocType.ConveyanceRateMaster, "", False, False)
                InsertDefaultValue(clsDocType.ODMaster, "", False, False)
                InsertDefaultValue(clsDocType.SubDepartmentMaster, "", False, False)
                InsertDefaultValue(clsDocType.PayrollPaymentMode, "", False, False)
                InsertDefaultValue(clsDocType.MFAccountSet, "", False, False)
                InsertDefaultValue(clsDocType.ExpenseHead, "", False, False)
                InsertDefaultValue(clsDocType.MFToolType, "", False, False)
                InsertDefaultValue(clsDocType.WorkCenterMaster, "", False, False)
                InsertDefaultValue(clsDocType.ResourceMaster, "", False, False)
                InsertDefaultValue(clsDocType.MFToolMaster, "", False, False)
                InsertDefaultValue(clsDocType.OperationMaster, "", False, False)
                InsertDefaultValue(clsDocType.SectionMaster, "", False, False)
                InsertDefaultValue(clsDocType.StageMaster, "", False, False)
                InsertDefaultValue(clsDocType.ItemCategoryForProduction, "", False, False)
                InsertDefaultValue(clsDocType.ProductionLine, "", False, False)
                InsertDefaultValue(clsDocType.PJCSetting, "", False, False)
                InsertDefaultValue(clsDocType.CostType, "", False, False)
                InsertDefaultValue(clsDocType.PJCAccountSet, "", False, False)
                InsertDefaultValue(clsDocType.JobMaster, "", False, False)
                InsertDefaultValue(clsDocType.TaskMaster, "", False, False)
                InsertDefaultValue(clsDocType.ComplaintGroup, "", False, False)
                InsertDefaultValue(clsDocType.PrimaryReasonMaster, "", False, False)
                InsertDefaultValue(clsDocType.DeductionMaster, "", False, False)
                InsertDefaultValue(clsDocType.DeductionGroup, "", False, False)
                InsertDefaultValue(clsDocType.MilkReasonMaster, "", False, False)
                InsertDefaultValue(clsDocType.PaymentCycleMaster, "", False, False)
                InsertDefaultValue(clsDocType.LostDefectSealNo, "", False, False)
                InsertDefaultValue(clsDocType.ItemChargeCategoryMaster, "", False, False)
                InsertDefaultValue(clsDocType.SupplierMaster, "", False, False)
                InsertDefaultValue(clsDocType.DivertedContractor, "", False, False)
                InsertDefaultValue(clsDocType.MilkGradeMaster, "", False, False)
                InsertDefaultValue(clsDocType.MilkTypeMaster, "", False, False)
            End If

            '=====sanjeet(PS RETURN Cancel)=====
            InsertDefaultValue(clsDocType.frmDairySaleReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.frmDairySaleReturn, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.frmSaleReturnProductSale, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.frmSaleReturnProductSale, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.frmSaleReturnFreshSale, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.frmSaleReturnFreshSale, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.BulkSaleReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.BulkSaleReturn, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.MCCMaterialSaleReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.MCCMaterialSaleReturn, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.ScrapReturn, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.ScrapReturn, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.CSASALERETURN, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.CSASALERETURN, clsDocTransactionType.SaleReturnCancel, False, True)
            InsertDefaultValue(clsDocType.frmDairyShipmentReturn, clsDocTransactionType.NA, False, True)

            InsertDefaultValue(clsDocType.FrmMCCScrapGatePass, clsDocTransactionType.frmMccScrapgatePass, False, True)

            '=====================
            '
            InsertDefaultValue(clsDocType.Cleaning, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.Cleaning, clsDocTransactionType.MCCProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.MilkCleaning, "", False, True)

            clsDBFuncationality.ExecuteNonQuery("update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='" + clsDocTransactionType.NA + "'  where Doc_Type='" + clsDocType.GateOut + "' and len(Doc_Trans_Type)<=0")
            InsertDefaultValue(clsDocType.GateOut, clsDocTransactionType.NA, False, True)
            InsertDefaultValue(clsDocType.GateOut, clsDocTransactionType.BulkProcPurchase, False, True)
            InsertDefaultValue(clsDocType.GateOut, clsDocTransactionType.BulkProcJobWork, False, True)
            InsertDefaultValue(clsDocType.GateOut, clsDocTransactionType.BulkProcJobWorkOutward, False, True)
            InsertDefaultValue(clsDocType.GateOut, clsDocTransactionType.MCCProcJobWorkOutward, False, True)

            InsertDefaultValue(clsDocType.LostDefectSealNo, "", False, True)

            InsertDefaultValue(clsDocType.MccMilkTransferPrice, "", False, True)
            ''richa 10/09/2014
            InsertDefaultValue(clsDocType.TankerOut, "", False, True)
            InsertDefaultValue(clsDocType.VSPAssetIssue, clsDocTransactionType.ItemIssue, False, True)
            InsertDefaultValue(clsDocType.VSPAssetIssue, clsDocTransactionType.ItemReturn, False, True)
            InsertDefaultValue(clsDocType.VSPItemIssue, clsDocTransactionType.ItemIssue, False, True)
            InsertDefaultValue(clsDocType.VSPItemIssue, clsDocTransactionType.ItemReturn, False, True)
            InsertDefaultValue(clsDocType.RequestForTraining, "", False, False)
            InsertDefaultValue(clsDocType.ProspectDetail, clsDocTransactionType.Prospect, False, True)
            '--Preeti Gupta
            InsertDefaultValue(clsDocType.HRApplicantEntry, "", False, False)
            InsertDefaultValue(clsDocType.HRRequisitionEntry, "", False, False)
            InsertDefaultValue(clsDocType.Item_Conversion, "", False, False)
            InsertDefaultValue(clsDocType.HRTrainerFeedBack, "", False, False)
            InsertDefaultValue(clsDocType.HRTraineeFeedBack, "", False, False)
            InsertDefaultValue(clsDocType.MRP, "", False, True)
            InsertDefaultValue(clsDocType.SupplierRegistration, "", False, False)

            '' Service And Warranty 
            InsertDefaultValue(clsDocType.SWServiceCall, "", False, False)
            InsertDefaultValue(clsDocType.SWServiceEnquiry, "", False, False)
            InsertDefaultValue(clsDocType.SWServiceAllocation, "", False, False)
            InsertDefaultValue(clsDocType.SWServiceVisitDetails, "", False, False)
            '==========> shivani (Milk job Work)
            InsertDefaultValue(clsDocType.MilkGateEntry, clsDocTransactionType.Tanker, False, True)
            InsertDefaultValue(clsDocType.MilkGateEntry, clsDocTransactionType.Sku_Receipt, False, True)
            InsertDefaultValue(clsDocType.JobMilkSRN, clsDocTransactionType.Tanker, False, True)
            InsertDefaultValue(clsDocType.JobMilkSRN, clsDocTransactionType.Sku_Receipt, False, True)
            InsertDefaultValue(clsDocType.JobMilkSRN, clsDocTransactionType.Manual, False, True)

            InsertDefaultValue(clsDocType.JWOSRN, clsDocTransactionType.Manual, False, True)
            InsertDefaultValue(clsDocType.JWOSRN, clsDocTransactionType.Tanker, False, True)
            InsertDefaultValue(clsDocType.JWOSRN, clsDocTransactionType.Sku_Receipt, False, True)

            InsertDefaultValue(clsDocType.JWOSRNReturn, clsDocTransactionType.Manual, False, True)
            InsertDefaultValue(clsDocType.JWOSRNReturn, clsDocTransactionType.Tanker, False, True)
            InsertDefaultValue(clsDocType.JWOSRNReturn, clsDocTransactionType.Sku_Receipt, False, True)

            InsertDefaultValue(clsDocType.DailyElectricalEntry, "", False, False)
            InsertDefaultValue(clsDocType.CustomerIncentiveMaster, "", False, False)

            InsertDefaultValue(clsDocType.TransporterDeductionMaster, "", False, False)
            InsertDefaultValue(clsDocType.InvestmentDeclarationMaster, "", False, False)
            InsertDefaultValue(clsDocType.CardSale, "", False, False)


            InsertDefaultValue(clsDocType.MilkWeighment, "", False, True)
            'InsertDefaultValue(clsDocType.MilkWeighment, clsDocTransactionType.Sku_Receipt, False, True)
            InsertDefaultValue(clsDocType.JobMilkQualityCheck, "", False, True)
            'InsertDefaultValue(clsDocType.JobMilkQualityCheck, clsDocTransactionType.Sku_Receipt, False, True)
            InsertDefaultValue(clsDocType.MilkUnloading, "", False, True)

            '======================================================
            '=======ROhit=================
            InsertDefaultValue(clsDocType.MilkRGP, "", False, True)
            '==============================================
            ''==payroll
            InsertDefaultValue(clsDocType.EmployeeTransfer, "", False, False)
            '==HR
            '=======================Added by preeti gupta=======================
            InsertDefaultValue(clsDocType.DamageDetail, "", False, False)
            InsertDefaultValue(clsDocType.HREMResignationLetter, "", False, False)
            InsertDefaultValue(clsDocType.HREMTerminationLetter, "", False, False)
            InsertDefaultValue(clsDocType.HRExitInterview, "", False, False)
            '========================
            '' csa booking 
            InsertDefaultValue(clsDocType.CSABooking, "", False, False)
            InsertDefaultValue(clsDocType.CSARequest, "", False, True)
            InsertDefaultValue(clsDocType.ProvisionEntry, "", False, True)
            InsertDefaultValue(clsDocType.ProvisionEntryMilk, clsDocTransactionType.MccProc, False, True, False, objCommonVar.ShowMCCFinderInPaymentProcess)

            InsertDefaultValue(clsDocType.PaymentProcess, "", False, True)
            InsertDefaultValue(clsDocType.PaymentProcessWithoutLoc, "", False, False)
            InsertDefaultValue(clsDocType.DispatchTransfer, "", False, True)

            '' job work outward
            InsertDefaultValue(clsDocType.GateEntryJWO, "", False, True)
            InsertDefaultValue(clsDocType.WeighmentJWO, "", False, True)
            InsertDefaultValue(clsDocType.QCJWO, "", False, True)
            InsertDefaultValue(clsDocType.UnloadingJWO, "", False, True)

            ''can sale
            InsertDefaultValue(clsDocType.CanSale, "", False, True)
            InsertDefaultValue(clsDocType.CanSaleDispatch, "", False, True)
            InsertDefaultValue(clsDocType.CanSaleInvoice, "", False, True)

            ''Common sale
            InsertDefaultValue(clsDocType.CommonSaleSeries, clsDocTransactionType.GSTTaxable, False, True)
            'InsertDefaultValue(clsDocType.CommonSaleSeries, clsDocTransactionType.GSTNonTaxable, False, True)
            InsertDefaultValue(clsDocType.CommonSaleSeries, clsDocTransactionType.GSTBillofSupply, False, True)
            InsertDefaultValue(clsDocType.CommonSaleSeries, clsDocTransactionType.GSTDeliveryChallan, False, True)

            InsertDefaultValue(clsDocType.ShipToLocation, "", False, True)

            InsertDefaultValue(clsDocType.MILKTruckSheet, "", False, True)
            InsertDefaultValue(clsDocType.VLC_Target, "", False, True)
            InsertDefaultValue(clsDocType.MilkRecurringSchedule, "", False, False)
            InsertDefaultValue(clsDocType.CSATransfer, clsDocTransactionType.NA, False, True)
            'InsertDefaultValue(clsDocType.CSATransfer, clsDocTransactionType.CSATransferLocal, False, True)
            'InsertDefaultValue(clsDocType.CSATransfer, clsDocTransactionType.CSATransferInterState, False, True)
            InsertDefaultValue(clsDocType.CSATransfer, clsDocTransactionType.GSTTaxable, False, True)
            InsertDefaultValue(clsDocType.CSATransfer, clsDocTransactionType.GSTNonTaxable, False, True)

            InsertDefaultValue(clsDocType.VendorBankMaster, "", False, False)
            InsertDefaultValue(clsDocType.IncentiveMaster, "", False, False)
            InsertDefaultValue(clsDocType.vlcDataUploader, "", False, True)
            InsertDefaultValue(clsDocType.BankGuaranteeMaster, "", False, False)
            InsertDefaultValue(clsDocType.ItemStockConversion, "", False, True)
            InsertDefaultValue(clsDocType.EmpMediclaim, "", False, False)
            InsertDefaultValue(clsDocType.EmployeeShiftChange, "", False, False)
            InsertDefaultValue(clsDocType.PayrollSetting, "", False, False)
            InsertDefaultValue(clsDocType.ExportIncentiveMaster, "", False, False)
            InsertDefaultValue(clsDocType.ConveyanceClaim, "", False, False)

            InsertDefaultValue(clsDocType.AssetWork, "", False, True)
            InsertDefaultValue(clsDocType.ItemIssueReturnAsset, clsDocTransactionType.ItemIssueAsset, False, True)
            InsertDefaultValue(clsDocType.ItemIssueReturnAsset, clsDocTransactionType.ItemReturnAsset, False, True)
            InsertDefaultValue(clsDocType.ItemIssueReturnAsset, clsDocTransactionType.TransferCapex, False, True)
            InsertDefaultValue(clsDocType.OpeningBankReco, "", False, True)
            InsertDefaultValue(clsDocType.UnclearedDoc, "", False, False)
            InsertDefaultValue(clsDocType.RevaluationEntry, "", False, False)
            InsertDefaultValue(clsDocType.AssetDispatchRetailer, clsDocTransactionType.ItemIssue, False, True)
            InsertDefaultValue(clsDocType.AssetDispatchRetailer, clsDocTransactionType.ItemReturn, False, True)
            '=================Added by preeti Gupta=======================
            InsertDefaultValue(clsDocType.frmCreateReceivedTransfer, "", False, True)
            InsertDefaultValue(clsDocType.CustomerMaster, "", False, False)
            InsertDefaultValue(clsDocType.VendorMaster, "", False, False)
            '================added by kunal======================
            InsertDefaultValue(clsDocType.EmptyWtEntry, "", False, True)
            InsertDefaultValue(clsDocType.LoadedWtEntry, "", False, True)
            InsertDefaultValue(clsDocType.frmSilageProdApplicationForm, "", False, False)
            InsertDefaultValue(clsDocType.frmSilageEnterPrenur, "", False, False)
            InsertDefaultValue(clsDocType.frmSilageFarmerSelection, "", False, False)
            '=================================================================================
            InsertDefaultValue(clsDocType.MilkGateEntryIn, "", False, True)
            InsertDefaultValue(clsDocType.MilkGateEntryWeighment, "", False, True)
            InsertDefaultValue(clsDocType.MilkGateEntryOut, "", False, True)
            InsertDefaultValue(clsDocType.PlantWeighment, "", False, True)

            '=====================================================
            'STUTI'
            InsertDefaultValue(clsDocType.CAPEXMASTER, "", False, False)
            InsertDefaultValue(clsDocType.CAPEXBUDGET, "", False, False)
            InsertDefaultValue(clsDocType.EMPLOYEEBANDMASTER, "", False, False)
            InsertDefaultValue(clsDocType.VENREG, "", False, False)
            '=======END HERE=========
            ''=======Parteek
            InsertDefaultValue(clsDocType.BankReco, clsDocTransactionType.Bank_Reco, False, True)
            InsertDefaultValue(clsDocType.BankRevseEnty, clsDocTransactionType.Account_Payable, False, True)
            InsertDefaultValue(clsDocType.BankRevseEnty, clsDocTransactionType.Account_Receivable, False, True)
            ''End=============
            ''Parteek============
            InsertDefaultValue(clsDocType.PROD_Assemblies, clsDocTransactionType.Assembles, False, True)
            InsertDefaultValue(clsDocType.PROD_Assemblies, clsDocTransactionType.Deassmbles, False, True)
            InsertDefaultValue(clsDocType.RackMaster, "", False, True)
            InsertDefaultValue(clsDocType.BinMaster, "", False, True)

            ''End============
            InsertDefaultValue(clsDocType.ConsumerDetails, "", False, False)
            'rranjan''====end here=============

            'KUNAL > CLIENT : UDIL > TICKET : BM00000010226 
            InsertDefaultValue(clsDocType.NRGPREQUEST, "", False, True)

            '-- added by parteek 23/03/2017 client Adinath
            ' InsertDefaultValue(clsDocType.PurchaserOrder, clsDocTransactionType.DirectAP, False, True)

            '' Farmer Payment Start Here
            InsertDefaultValue(clsDocType.frmMCCMaterialSaleFarmer, "", False, True)
            InsertDefaultValue(clsDocType.frmMCCMaterialSaleInvoiceFarmer, "", False, True)
            InsertDefaultValue(clsDocType.ARInvoiceFarmer, "", False, True)
            InsertDefaultValue(clsDocType.ARDebitNoteFarmer, "", False, True)
            InsertDefaultValue(clsDocType.ARCreditNoteFarmer, "", False, True)
            InsertDefaultValue(clsDocType.frmMCCMaterialSaleReturnFarmer, "", False, True)
            InsertDefaultValue(clsDocType.FarmerPayment, clsDocTransactionType.Bank, False, True)
            InsertDefaultValue(clsDocType.FarmerPayment, clsDocTransactionType.Cash, False, True)
            InsertDefaultValue(clsDocType.FarmerPayment, clsDocTransactionType.PettyCash, False, True)
            InsertDefaultValue(clsDocType.FarmerPayment, clsDocTransactionType.Others, False, True)
            InsertDefaultValue(clsDocType.FarmerPayment, clsDocTransactionType.ApplyDocument, False, True)
            InsertDefaultValue(clsDocType.PaymentProcessFarmer, "", False, True)
            InsertDefaultValue(clsDocType.FrmItemMasterRMOther, "", False, False)

            InsertDefaultValue(clsDocType.JWPriceCode, "", False, False)
            InsertDefaultValue(clsDocType.JWIPriceCode, "", False, False)
            InsertDefaultValue(clsDocType.DeductionMapping, "", False, False)
            '' Contra Voucher Createdby Parteek 18/05/2017
            InsertDefaultValue(clsDocType.ContraVoucher, "", False, True)
            InsertDefaultValue(clsDocType.VSPMapping, "", False, False)
            InsertDefaultValue(clsDocType.FarmerProMaster, "", False, False)
            InsertDefaultValue(clsDocType.UserRequestMaster, "", False, False)
            InsertDefaultValue(clsDocType.DCSAdditionDeduction, "", False, False)
            InsertDefaultValue(clsDocType.MPIncetiveSlab, "", False, False)
            InsertDefaultValue(clsDocType.ChillingChargesSlab, "", False, False)
            InsertDefaultValue(clsDocType.CappingMaster, "", False, False)
            InsertDefaultValue(clsDocType.MatrixPriceChart, "", False, False, True, False)
            InsertDefaultValue(clsDocType.MatrixPricePlan, "", False, False, True, False)
            InsertDefaultValue(clsDocType.GazeReading, "", False, False, True, False)
            InsertDefaultValue(clsDocType.DSCFinancialHead, "", False, False, True, False)
            InsertDefaultValue(clsDocType.DSCFinancialEntry, "", False, False, True, False)
            '' -----End
            InsertDefaultValue(clsDocType.PricePlan, "", False, True)
            'InsertDefaultValue(clsDocType.ItemWiseTax, "", False, False)

            InsertDefaultValue(clsDocType.ItemWiseTax, clsDocTransactionType.ItemPurchaseTax, False, False)
            InsertDefaultValue(clsDocType.ItemWiseTax, clsDocTransactionType.ItemSaleTax, False, False)
            InsertDefaultValue(clsDocType.ItemWiseTax, clsDocTransactionType.ItemTransferTax, False, False)

            InsertDefaultValue(clsDocType.SACWiseTax, clsDocTransactionType.SACWiseTaxMaster, False, False)

            InsertDefaultValue(clsDocType.JobWorkTransferOther, "", False, True)
            InsertDefaultValue(clsDocType.JobWorkTransferOtherReturn, "", False, True)
            InsertDefaultValue(clsDocType.JobWorkTransferMilkReturn, "", False, True)
            InsertDefaultValue(clsDocType.ItemCostMapping, "", False, False)
            InsertDefaultValue(clsDocType.QuickBookEntry, "", False, False)
            InsertDefaultValue(clsDocType.MakeSavingPayment, clsDocTransactionType.Transaction, False, False)
            InsertDefaultValue(clsDocType.IncentiveEntry, clsDocTransactionType.Transaction, False, True)
            InsertDefaultValue(clsDocType.Detail, clsDocTransactionType.Detail, False, False)

            InsertDefaultValue(clsDocType.DetailSale, clsDocTransactionType.Detail, False, False, False, False, True)
            InsertDefaultValue(clsDocType.DetailSale, clsDocTransactionType.Uploader, False, False, False, False, True)
            InsertDefaultValue(clsDocType.DetailSale, clsDocTransactionType.ProductDetail, False, False, False, False, True)
            InsertDefaultValue(clsDocType.CustomerDeduction, "", False, False)
            InsertDefaultValue(clsDocType.CustomerIncentiveEntry, clsDocTransactionType.Transaction, False, True)
            InsertDefaultValue(clsDocType.IncomeTaxSlab, "", False, False)
            InsertDefaultValue(clsDocType.IncomeTaxCalculation, "", False, False)
            InsertDefaultValue(clsDocType.TankerCleaningItems, "", False, True)

            'Eng. And Plant Management
            InsertDefaultValue(clsDocType.SectionMasterEng, "", False, False)
            InsertDefaultValue(clsDocType.ConsumptionTypeMaster, "", False, False)
            InsertDefaultValue(clsDocType.LogSheetEng, "", False, True)
            InsertDefaultValue(clsDocType.WorkRequisitionEng, clsDocTransactionType.OtherExternal, False, True)
            InsertDefaultValue(clsDocType.WorkEstimationEng, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.WorkOrderEng, clsDocTransactionType.POJobWork, False, True)
            InsertDefaultValue(clsDocType.WorkOrderStatusEng, "", False, True)

            InsertDefaultValue(clsDocType.FixedAsset, "", False, False)

            InsertDefaultValue(clsDocType.BreakDownEntry, "", False, True)

            qry = "select ITEM_TYPE_CODE from TSPL_ITEM_TYPE_MASTER"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    InsertDefaultValue(clsDocType.ItemBatch, clsCommon.myCstr(dr("ITEM_TYPE_CODE")), False, True)
                Next
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Try
            Dim sQuery As String = " select distinct OtherAssemblyFilePathAndName  from TSPL_PROGRAM_MASTER  where isnull(IsLoadFromOtherAssembly ,0)=1"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(sQuery)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim AsmName As String = clsCommon.myCstr(dt.Rows(i)("OtherAssemblyFilePathAndName"))
                    InvokeMethodSlow(AsmName, "clsDocTypeCustom", "SetDefaultValues", Nothing)
                Next
            End If
        Catch ex As Exception
        End Try
        Return True
    End Function

    Public Shared Sub InvokeMethodSlow(AssemblyName As String, ClassName As String, MethodName As String, args As Object())
        Try
            Dim ass As Assembly = Assembly.LoadFrom(Application.StartupPath & "\" & AssemblyName)
            Dim FileAtt As String = IO.Path.GetFileNameWithoutExtension(AssemblyName)
            Dim factory As Object = ass.CreateInstance(FileAtt & "." & ClassName, True)
            Dim t As Type = factory.GetType
            Dim method As MethodInfo = t.GetMethod(MethodName)
            Dim obj As Object = method.Invoke(factory, args)
        Catch ex As Exception
        End Try

    End Sub

    Private Shared Function InsertDefaultValue(ByVal strDocType As String, ByVal strDocTransactinType As String, ByVal IsStateWise As Boolean, ByVal IsLocationWise As Boolean) As Boolean
        Return InsertDefaultValue(strDocType, strDocTransactinType, IsStateWise, IsLocationWise, False)
    End Function
    Private Shared Function InsertDefaultValue(ByVal strDocType As String, ByVal strDocTransactinType As String, ByVal IsStateWise As Boolean, ByVal IsLocationWise As Boolean, ByVal Master_Prefix As Boolean) As Boolean
        Return InsertDefaultValue(strDocType, strDocTransactinType, IsStateWise, IsLocationWise, Master_Prefix, False)
    End Function
    Private Shared Function InsertDefaultValue(ByVal strDocType As String, ByVal strDocTransactinType As String, ByVal IsStateWise As Boolean, ByVal IsLocationWise As Boolean, ByVal Master_Prefix As Boolean, ByVal IsMCCWise As Boolean) As Boolean
        Return InsertDefaultValue(strDocType, strDocTransactinType, IsStateWise, IsLocationWise, Master_Prefix, IsMCCWise, False)
    End Function
    Private Shared Function InsertDefaultValue(ByVal strDocType As String, ByVal strDocTransactinType As String, ByVal IsStateWise As Boolean, ByVal IsLocationWise As Boolean, ByVal Master_Prefix As Boolean, ByVal IsMCCWise As Boolean, ByVal IsRouteWise As Boolean) As Boolean
        'If Master_Prefix = True Then
        '    Dim qry As String = ""
        '    qry = "update TSPL_DOCPREFIX_MASTER set Next_Number=(select max(Next_Number) FROM TSPL_DOCPREFIX_MASTER WHERE Doc_Type='" + strDocType + "') where Doc_Type='" + strDocType + "' AND PK_ID IN (SELECT MAX(PK_ID) FROM TSPL_DOCPREFIX_MASTER WHERE Doc_Type='" + strDocType + "')"
        '    clsDBFuncationality.ExecuteNonQuery(qry)
        '    'qry = "DELETE FROM TSPL_DOCPREFIX_MASTER WHERE Doc_Type='" + strDocType + "' AND PK_ID NOT IN (SELECT MAX(PK_ID) FROM TSPL_DOCPREFIX_MASTER WHERE Doc_Type='" + strDocType + "')"
        '    'clsDBFuncationality.ExecuteNonQuery(qry)
        'End If
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Doc_Type", strDocType)
        clsCommon.AddColumnsForChange(coll, "Doc_Trans_Type", strDocTransactinType)
        clsCommon.AddColumnsForChange(coll, "Is_State_Wise", IIf(IsStateWise, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Is_Location_Wise", IIf(IsLocationWise, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Master_Prefix", IIf(Master_Prefix, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Is_MCC_Wise", IIf(IsMCCWise, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Is_Route_Wise", IIf(IsRouteWise, 1, 0))
        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DOCUMENT_TYPE", OMInsertOrUpdate.Insert, "")
        Return True
    End Function
End Class

Public Class clsDocTransactionType
    Public Const RequestService As String = "Request Service"
    Public Const MCCSaleRequest As String = "MCC Sale Request"
    Public Const MPSale As String = "MP Sale"
    Public Const ComplainFeedback As String = "Complain Feedback"
    Public Const DairyGateout As String = "Dairy Gateout"
    Public Const DairyDelivery As String = "Dairy Delivery"
    Public Const DairyInvoiceAcknowledgment As String = "Dairy Invoice Acknowledgment"
    Public Const BulkProcJobWorkOutward As String = "BulkProc JobWork OutWard"
    Public Const MCCProcJobWorkOutward As String = "MCCProc JobWork OutWard"
    Public Const Normal As String = "Normal"
    Public Const ApplyDoc As String = "Apply Doc"
    Public Const Bank As String = "Bank"
    Public Const Cash As String = "Cash"
    Public Const PettyCash As String = "Petty Cash"
    Public Const Others As String = "Others"
    Public Const Settlement As String = "Settlement"
    Public Const Tax As String = "Tax"
    Public Const ApplyDocument As String = "Apply Document"
    Public Const SaleInvoiceExcise As String = "Excise"
    Public Const SaleInvoiceRetail As String = "Retail"
    Public Const SaleInvoiceTax As String = "Tax"
    Public Const SaleInvoiceInterstate As String = "Interstate"
    Public Const SaleInvoiceService As String = "Service"
    Public Const SaleReturnExcise As String = "Excise"
    Public Const SaleReturnRetail As String = "Retail"
    Public Const SaleReturnTax As String = "Tax"
    Public Const GSTTaxable As String = "Taxable"
    Public Const GSTNonTaxable As String = "Non Taxable"
    Public Const GSTNonTaxablePNR As String = "Non Taxable PNR"
    Public Const GSTBillofSupply As String = "Bill of Supply"
    Public Const GSTIn As String = "GST In"
    Public Const GSTReturn As String = "GST Return"
    Public Const GSTLocal As String = "Local"
    Public Const GSTInterstate As String = "Interstate"
    Public Const GST_Interstate As String = "GST Interstate"
    Public Const GSTDeliveryChallan As String = "Delivery Challan"
    Public Const GSTNA As String = "GST NA"
    Public Const SupplementaryLocal As String = "SupplementaryLocal"
    Public Const SupplementaryInterstate As String = "SupplementaryInterstate"
    Public Const SupplementaryNonTaxable As String = "SupplementaryNonTaxable"
    Public Const SupplementaryCNoteLocal As String = "SupplementaryCNoteLocal"
    Public Const SupplementaryCNoteInterstate As String = "SupplementaryCNoteInterstate"
    Public Const SupplementaryCNoteNonTaxable As String = "SupplementaryCNoteNonTaxable"
    Public Const CreditRefDoc As String = "CreditRefDoc"
    Public Const DebitRefDoc As String = "DebitRefDoc"
    ''Parteek
    Public Const Bank_Reco As String = "Bank Reco"
    Public Const Account_Payable As String = "Account Payable"
    Public Const Account_Receivable As String = "Account Receivable"
    Public Const Assembles As String = "Assembly"
    Public Const Deassmbles As String = "DisAssembly"
    ''End
    ''Public Const POLocal As String = "Local"
    ''Public Const POImport As String = "Import"
    ''Public Const POJobWork As String = "Job Work"
    ''Public Const POOpen As String = "Open"
    ''Public Const POSpecific As String = "Specific"

    ''Public Const PORawMaterial As String = "Raw Material"
    Public Const POFinishedGoods As String = "Finished Goods"
    Public Const POSemiFinishedGoods As String = "SemiFinished Goods"
    Public Const POJobWork As String = "Job Work"
    Public Const POJobWorkOutward As String = "Job Work Outward"
    Public Const DistributRateTag As String = "Distribute Rate Tag"

    Public Const PODomestic As String = "Domestic"
    Public Const POImport As String = "Import"
    Public Const POOutgoing As String = "Outgoing"
    Public Const RCDFLoadin As String = "RCDF Loadin"
    Public Const RGPWise As String = "RGP Wise"

    Public Const POOther As String = "Other"
    Public Const PORawMaterial As String = "Raw Material"
    Public Const POAsset As String = "Asset"
    Public Const POTrading As String = "Trading"
    Public Const PONonInvntry As String = "Non-Inventory"
    Public Const SRNRGP As String = "Returnable Gate Pass"

    Public Const MilkPurchase As String = "Milk Purchase"

    'Public Const BulkMilkTaxInvoice As String = "Bulk Milk Tax Invoice"
    'Public Const BulkMilkTaxInvoiceTrade As String = "Bulk Milk Tax Invoice Trade"

    Public Const BulkMilkTaxInvoice As String = "Bulk Milk Bill of Supply"
    Public Const BulkMilkTaxInvoiceTrade As String = "Bulk Milk Bill of Supply Trade"

    Public Const DirectPurchaseFromAP As String = "Direct Purchase From AP"

    Public Const ItemIssue As String = "Issue"
    Public Const ItemReturn As String = "Return"
    Public Const ItemTransfer As String = "Transfer"
    Public Const TransferCapex As String = "Transfer Capex"
    Public Const ItemIssue_Reprocess As String = "Issue for Reprocess"


    Public Const TransferOut As String = "Transfer Out"
    Public Const BULKMilkSale As String = "Sale"
    Public Const BULKMilkSaleTrade As String = "Trade"
    Public Const TransferLocalOut As String = "Transfer Out [Local]"
    Public Const TransferInterStateOut As String = "Transfer Out [InterState]"
    Public Const TransferIn As String = "Transfer In"
    Public Const TransferReject As String = "Transfer Reject"

    Public Const InternalTransfer As String = "Internal Transfer"
    Public Const InternalTransferOut As String = "Internal Transfer Out"


    Public Const JWTransfer As String = "Job Work Transfer"
    Public Const JWTransferOut As String = "Job Work Transfer Out"

    Public Const TransferFull As String = "Full"
    Public Const TranferEmpty As String = "Empty"
    Public Const TransferAutoRoute As String = "Auto Route"

    Public Const CSATransferLocal As String = "CSA Transfer [Local]"
    Public Const CSATransferInterState As String = "CSA Transfer [InterState]"
    ''Public Const AdjEmpty As String = "Empty"
    ''Public Const AdjRaw As String = "Raw Material"
    ''Public Const AdjOther As String = "Adjustment"
    ''Public Const AdjTrading As String = "FG Trading"
    ''Public Const AdjMfg As String = "FG Manufacturing"

    Public Const ScrapInvoiceRetail As String = "Retail"
    Public Const ScrapInvoiceTax As String = "Tax"
    Public Const ScrapInvoiceScrapSale As String = "ScrapI"
    Public Const ScrapInvoiceScrapSaleR As String = "ScrapR"
    Public Const ScrapInvoiceScrapSaleT As String = "ScrapT"
    Public Const ScrapInvoiceScrapSaleE As String = "ScrapE"

    Public Const EmptyTransactionBySaleInvoice As String = "By Sale Invoice"
    Public Const EmptyTransactionOut As String = "Out"
    Public Const EmptyTransactionRouteIn As String = "Route In"
    Public Const EmptyTransactionDepotIn As String = "Depot In"

    'Public Const ProductionEntryFGTrading As String = "FG Trading"
    'Public Const ProductionEntryFGManufacturing As String = "FG Manufacturing"

    Public Const StoreAdjustmentAdjustment As String = "Adjustment"
    Public Const StoreAdjustmentAdjustmentProductionEntry As String = "Production Entry"
    Public Const StoreAdjustmentAdjustmentProductionEntryQC As String = "Production Entry QC"
    Public Const StoreAdjustmentAdjustmentProductionStoreEntry As String = "Production Store Entry"

    Public Const RouteSeriesTax As String = "Route Series Tax"
    Public Const RouteSeriesRetail As String = "Route Series Retail"

    Public Const JournalEntryJLJE As String = "JL-JE"
    Public Const JournalEntryJLJEReverseGeneral As String = "JL-JE-RG"
    Public Const JournalEntryJLJEReverseMonthly As String = "JL-JE-RM"

    Public Const JournalEntryOther As String = "Others"
    Public Const JournalEntryMilkSRN As String = "MilkSRN"
    Public Const JournalEntryPaymentProcess As String = "Payment Process"
    Public Const FinishedGoodInternal As String = "Finished Good Internal"
    Public Const FinishedGoodExternal As String = "Finished Good External"
    Public Const SemiFinishedGoodInternal As String = "SemiFinished Good Internal"
    Public Const SemiFinishedGoodExternal As String = "SemiFinished Good External"
    Public Const OtherInternal As String = "Other Internal"
    Public Const OtherExternal As String = "Other External"

    Public Const IncomingQualityCheck As String = "Incoming QC"
    Public Const InprocessQualityCheck As String = "InProcess QC"
    Public Const OutgoingQualityCheck As String = "Outgoing QC"
    Public Const ProductionQC As String = "Out ProductionQC"

    Public Const BOMOSPTYPE As String = "Job-Work"
    Public Const PPReturnProductionEntry As String = "PE"
    Public Const PPReturnProductionStandardization As String = "STD"

    Public Const SNQuotationFinishedGoods As String = "Finished Goods"
    Public Const SNQuotationOther As String = "Other"
    Public Const SNOrderExport As String = "Export"
    Public Const SNOrderMerchant As String = "Merchant Trade"
    Public Const BulkProc As String = "Bulk Procurement"
    Public Const BulkProcPurchase As String = "Bulk Proc Purchase"
    Public Const BulkProcJobWork As String = "Bulk Proc Job Work"
    Public Const NA As String = "NA"
    Public Const GeneralPurchase As String = "General Purchase"
    Public Const MccProc As String = "MccProc"
    Public Const ServiceCharge As String = "Service Charge"
    Public Const Prospect As String = "Prospect"
    Public Const CSA_CPD_DESI_GHEE As String = "CPD Desi Ghee"
    Public Const CSA_BULK_DESI_GHEE As String = "Bulk Desi Ghee"
    Public Const CSA_CPD_OTHER As String = "CPD Other"
    Public Const CSA_BULK_OTHER As String = "Bulk Other"
    Public Const WithVendorInvoiceNo As String = "With Vendor Invoice No"
    Public Const WithoutVendorInvoiceNo As String = "Without Vendor Invoice No"
    Public Const ItemIssueAsset As String = "IssueAsset"
    Public Const ItemReturnAsset As String = "ReturnAsset"
    Public Const TaxExempted_ProductInvoice As String = "Tax Exempted"
    Public Const CashSale As String = "Cash Sale"
    Public Const Other As String = "Other"
    Public Const Tanker As String = "Tanker"
    Public Const Sku_Receipt As String = "Sku Recpt"
    Public Const Manual As String = "Manual"
    Public Const DirectAP As String = "Direct AP"
    Public Const ItemPurchaseTax As String = "Item Purchase Tax"
    Public Const ItemSaleTax As String = "Item Sale Tax"
    Public Const ItemTransferTax As String = "Item Transfer Tax"
    Public Const SACWiseTaxMaster As String = "Sac Wise tax Master"
    '======Sanjeet(PS RETURN)=====
    Public Const SaleReturnCancel As String = "Sale Return Cancel"

    Public Const frmMccScrapgatePass As String = "MCC/Scrap GatePass"
    Public Const LoanPayment As String = "Payment"
    Public Const LoanReceipt As String = "Receipt"
    Public Const Transaction As String = "Transaction"
    Public Const Detail As String = "Detail"
    Public Const ProductDetail As String = "ProductDetail"
    Public Const Uploader As String = "Uploader"
    Public Const ExcludeShipmentTRCode As String = "ExcludeShipmentTRCode"

    'Sanjay
    Public Const All As String = "All"
    Public Const PermanentBasis As String = "Permanent Basis"
    Public Const ContractBasis As String = "Contract Basis"
    Public Const DailyBasis As String = "Daily Basis"
    Public Const ReferenceNo As String = "Reference No"
    Public Const BMCU As String = "BMCU"
    Public Const MCC As String = "MCC"
    Public Const PDCS As String = "PDCS"
    Public Const Registered As String = "Registered"

End Class