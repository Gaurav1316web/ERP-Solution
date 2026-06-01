'--27/08/2012--Updation By-[Pankaj Kuamr]--Applied GL Security While navigating Document Finder----Fwdg By--Ranjana Mam
'--Updation By [Pankaj Kumar Chaudhary]--Against Ticket No--[BM00000001936 26/06/2014]
'================BM00000003058,Add code to save Rejected Qty in rejected Store====================.
'======BM00000007816
Imports common
Imports System.Data.SqlClient
Public Class clsSRNHead
#Region "Variables"
    Public isExemptSecurityDedution As Integer = 0
    Public ShowItemAllStructureWise As Boolean = False
    Public isJobWorkOutward As Integer = 0
    Public RGP_Type As String = Nothing
    Public Against_QC_Code As String = Nothing
    Public Against_QC_Date As String = Nothing
    Public Against_Schedule_Code As String = Nothing
    Public PurchaseOrder_Type As String = Nothing
    Public against_roadpermit As String = Nothing
    Public agnst_cform As String = Nothing
    Public Arr_Road As List(Of clsSRNRoadPermitDetail) = Nothing
    Public Arr_CFORM As List(Of clsSRNCFORMDetail) = Nothing
    Public Document_Type As String = Nothing
    Public autosrnfromrgp As String = Nothing
    Public isreadytopost As String = Nothing
    Public PROJECT_ID As String = Nothing
    Public SRN_No As String = Nothing
    Public SRN_Date As DateTime
    Public Vendor_Code As String = Nothing
    Public Vendor_Name As String = Nothing
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public On_Hold As Boolean = Nothing
    Public Ref_No As String = Nothing
    Public Challan_Date As String = Nothing
    Public Inv_Date As String = Nothing
    Public Description As String = Nothing
    Public Remarks As String = Nothing
    Public Inv_No As String = Nothing
    Public Bill_To_Location As String = Nothing
    Public BillToLocationName As String = Nothing
    Public Ship_To_Location As String = Nothing
    Public ShipToLocationName As String = Nothing
    Public Tax_Group As String = Nothing
    Public TaxGroupName As String = Nothing 'Not a table field
    Public TAX1 As String = Nothing
    Public TAX1_Rate As Double = 0
    Public TAX1_Base_Amt As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Rate As Double = 0
    Public TAX2_Base_Amt As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Rate As Double = 0
    Public TAX3_Base_Amt As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Rate As Double = 0
    Public TAX4_Base_Amt As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Rate As Double = 0
    Public TAX5_Base_Amt As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Rate As Double = 0
    Public TAX6_Base_Amt As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Rate As Double = 0
    Public TAX7_Base_Amt As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Rate As Double = 0
    Public TAX8_Base_Amt As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Rate As Double = 0
    Public TAX9_Base_Amt As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Rate As Double = 0
    Public TAX10_Base_Amt As Double = 0
    Public TAX10_Amt As Double = 0
    Public Total_Tax_Amt As Double = 0
    Public Discount_Base As Double = 0
    Public Discount_Amt As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public Total_Taxable_Amount As Decimal = 0
    Public SRN_Total_Amt As Double = 0
    Public TotalUnit_Cost_Tax As Double = 0
    Public Comments As String = Nothing
    Public Comp_Code As String = Nothing
    Public Terms_Code As String = Nothing
    Public TermsName As String = Nothing
    Public Due_Date As String = Nothing
    Public Posting_Date As String = Nothing
    Public Carrier As String = Nothing
    Public VehicleNo As String = Nothing
    Public GRNo As String = Nothing
    Public GENo As String = Nothing
    Public GEDate As Date? = Nothing
    Public RMDA_No As String = Nothing
    Public RMDA_Date As String = Nothing
    Public Dept As String = Nothing
    Public Dept_Desc As String = Nothing
    Public Item_Type As String = Nothing
    Public Landed_Add_Cost As Double = 0
    Public Total_Landed_Cost As Double = 0

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

    Public Against_Requisition As String = Nothing
    Public Against_PO As String = Nothing
    Public Against_GRN As String = Nothing
    Public Against_MRN As String = Nothing
    Public Against_RGP As String = Nothing
    Public Form_38 As String = Nothing
    Public Is_Internal As Boolean = False
    Public is_Excise_On_Qty As Boolean = False
    Public is_RGP_Non_Inventory As Boolean = False
    Public is_QCAccepted As Boolean = False
    Public AssessableAmt As Decimal = 0
    Public Inter_unit_purchase As Integer = 0

    Public Arr As List(Of clsSRNDetail) = Nothing
    Public Tax_Calculation_Type As EnumTaxCalucationType

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public IsAbatementPO As Integer = 0
    Public is_Srn_rejQty_goes_in_Rejstore As Boolean = False
    '===========================Rohit====================
    Public MRN_ID As String = Nothing
    '=============================

    Public Total_Accepted_Amount As Double = 0
    Public Total_Rejected_Amount As Double = 0
    Public Total_Shortage_Amount As Double = 0
    Public Total_Leak_Amount As Double = 0
    Public Total_Burst_Amount As Double = 0
    Public Is_Shortage_Include_In_Landed_Cost As Boolean = False
    'stuti
    Public RoadPermit_No As String = Nothing
    Public RoadPermit_Date As Date? = Nothing
    Public Sublocation_Code As String = String.Empty
    Public SubLocationName As String = String.Empty
    '===end here===
    Public Confirmatory_PO As Boolean

    Public Total_Item_Insurance_Amt As Decimal = 0
    Public Total_Add_Charge_Insurance As Decimal = 0
    Public Retention As Decimal = 0
    Public Arr_ACInsurance As List(Of clsSRNAdditionChargeInsurance) = Nothing
    Public TenderNo As String = Nothing
    Public GRN_Date As Date? = Nothing
    Public NIRQC As Double
    Public QC_Status As String = Nothing
    Public Against_CF_Sale_Purchase_No As String = Nothing
#End Region

    Public Shared Function funSRNPrint(ByVal Form_ID As String, ByVal isCancel As Boolean, ByVal TXTDATE As DateTime, ByVal FromDate As Date?, ByVal ToDate As Date?, ByVal IsDocTypeFinsihGoods As Boolean, ByVal ArrSrnNo As ArrayList, ByVal ArrVendor As ArrayList, ByVal ArrLocation As ArrayList)
        Dim qry As String
        Dim TSPL_SRN_HEAD As String = Nothing
        Dim TSPL_SRN_DETAIL As String = Nothing
        If isCancel Then
            TSPL_SRN_HEAD = " TSPL_SRN_HEAD_Cancel_Data"
            TSPL_SRN_DETAIL = " TSPL_SRN_DETAIL_Cancel_Data"
        Else
            TSPL_SRN_HEAD = " TSPL_SRN_HEAD"
            TSPL_SRN_DETAIL = " TSPL_SRN_DETAIL "
        End If
        Try

            '' Anubhooti 28-Aug-2014 (Demo Setting For Status) BM00000003672
            '=============Added by preeti Gupta Against ticket No[BHA/21/05/18-000028]
            Dim frmCRV As New frmCrystalReportViewer()
            Dim QryShowStatus As String = ""
            Dim IsSRNReportQtyWise As Boolean = False
            IsSRNReportQtyWise = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SRNReportQuantityWise, clsFixedParameterCode.SRNReportQuantityWise, Nothing)) = 1, True, False)
            Dim ShowStatusForPurchase As Double = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowStatusForPurchase' And Type ='ShowStatusForPurchase'")
            If clsCommon.CompairString(clsCommon.myCstr(ShowStatusForPurchase), "1") = CompairStringResult.Equal Then
                QryShowStatus = " ,(case when " + TSPL_SRN_HEAD + ".status =1 then 'Approved' else 'Pending' end) as Status "
            Else
                QryShowStatus = ""
            End If

            If IsDocTypeFinsihGoods Then

                qry = "Select '" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd-MM-yyyy") + "' as RunDate, XXXX.*, TSPL_COMPANY_MASTER.Comp_Name, TSPL_COMPANY_MASTER.Tin_No, TSPL_COMPANY_MASTER.CST_LST, TSPL_COMPANY_MASTER.Ecc_No, "
                qry += " (Case When ISNULL(TSPL_COMPANY_MASTER.Add1,'')='' Then '' Else TSPL_COMPANY_MASTER.Add1 + case When ISNULL(TSPL_COMPANY_MASTER.Add1,'')='' Then '' Else ', '+ TSPL_COMPANY_MASTER.Add2 + Case When ISNULL(TSPL_COMPANY_MASTER.Add3,'')='' Then '' Else TSPL_COMPANY_MASTER.Add3 + Case When ISNULL(TSPL_LOCATION_MASTER.add4,'')='' Then '' Else TSPL_LOCATION_MASTER.add4 End End End End) AS CompAddress, "
                qry += " TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.Tin_No as Vendor_Tin_No,TSPL_VENDOR_MASTER.Phone1 as Vendor_Contact, "
                qry += " (Case When ISNULL(TSPL_VENDOR_MASTER.Add1,'')='' Then '' Else TSPL_VENDOR_MASTER.Add1 + case When ISNULL(TSPL_VENDOR_MASTER.Add1,'')='' Then '' Else ', '+ TSPL_VENDOR_MASTER.Add2 + Case When ISNULL(TSPL_VENDOR_MASTER.Add3,'')='' Then '' Else ', '+TSPL_VENDOR_MASTER.Add3 End End End) AS VendorAddress from ("

                qry += " select max(Description) as Description ,max(Comments) as Comments,  max(Created_By) as Created_By,max(Modify_By) as Modify_By,  SRN_No,MAX(ItemType )as ItemType, MAX(Inv_No) as Inv_No, MAX(MRN_Date) as SRN_Date, MAX(SRNTime) as SrnTime,  MAX(Vendor_Code) as Vendor_Code,MAX(GRNo) as GRNo,"
                qry += " MAX(GENo) as GENo,MAX(GEDate) as GEDate,Item_Code,MAX(Item_Desc) as Item_Desc, MAX(MRP) as MRP,MAX(VehicleNo) as VehicleNo, "
                qry += " SUM(ISNULL( FCS,0)) as FCS, SUM(isnull(FBS,0))as FBS, SUM(ISNULL( FSH,0)) as FSH, SUM(ISNULL( ECS,0)) as ECS, SUM(ISNULL( EBS,0)) as EBS, "
                qry += " SUM(Leak_Qty) as HF,SUM(Burst_Qty) as Burst,SUM(Short_Qty) as Short,MAX(Remarks) as Remarks,max(Ref_No)as Ref_No,"
                qry += " max(Against_PO) as Against_PO, MAX(Against_RGP) as GPNo, Max(Comp_Code) as Comp_Code,MAX(Carrier) as Carrier from( "

                qry += " select Description,Comments, TSPL_SRN_HEAD.Created_By,TSPL_SRN_HEAD .Modify_By, TSPL_SRN_HEAD.SRN_No, TSPL_SRN_HEAD.Item_Type as ItemType,TSPL_SRN_HEAD.Against_PO,TSPL_SRN_HEAD.Carrier , TSPL_SRN_HEAD.Inv_No ,"
                qry += " (replace( CONVERT(varchar(11), TSPL_SRN_HEAD.SRN_Date,104),'.','/')+' '+CONVERT(varchar(100),TSPL_SRN_HEAD.SRN_Date,108) )as MRN_Date, CONVERT(VARCHAR,SRN_Date,108) as SRNTime,"
                qry += " TSPL_SRN_HEAD.Vendor_Code ,TSPL_SRN_HEAD.GRNo,TSPL_SRN_HEAD.GENo,(case when LEN(TSPL_SRN_HEAD.GEDate)>0  then REPLACE( CONVERT(varchar(11), TSPL_SRN_HEAD.GEDate,104),'.','/') else '' end) as GEDate,"
                qry += " TSPL_SRN_HEAD.VehicleNo,TSPL_SRN_HEAD.Remarks ,TSPL_SRN_HEAD.Ref_No, TSPL_SRN_HEAD.Against_RGP, TSPL_SRN_DETAIL.Item_Code, "
                qry += " TSPL_SRN_DETAIL.MRP, TSPL_SRN_DETAIL.Item_Desc,TSPL_SRN_DETAIL.Unit_code,case when Unit_code='FC' then SRN_Qty + ISNULL( Free_Qty,0) end as FCS,"
                qry += " case when Unit_code='FB' then SRN_Qty + ISNULL( Free_Qty,0) end as FBS, case when Unit_code='SH' then SRN_Qty + ISNULL( Free_Qty,0) end as FSH,"
                qry += " case when Unit_code='EC' then SRN_Qty + ISNULL( Free_Qty,0) end as ECS,case when Unit_code='EB' then SRN_Qty + ISNULL( Free_Qty,0) end as EBS, TSPL_SRN_DETAIL.Leak_Qty,TSPL_SRN_DETAIL.Burst_Qty,TSPL_SRN_DETAIL.Short_Qty, TSPL_SRN_HEAD.Comp_Code "
                qry += " from TSPL_SRN_DETAIL "
                qry += " left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No= TSPL_SRN_DETAIL.SRN_No  "
                qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_SRN_HEAD.Bill_To_Location"
                qry += " where Item_Type ='F' order by TSPL_SRN_DETAIL.line_no"

                If FromDate.HasValue AndAlso ToDate.HasValue Then
                    qry += " and Convert(date,TSPL_SRN_HEAD.SRN_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date,TSPL_SRN_HEAD.SRN_Date,103)<=Convert(date,'" + ToDate + "',103) "
                End If

                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
                    qry += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
                End If
                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
                    qry += " and TSPL_SRN_HEAD.SRN_No in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
                End If
                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
                    qry += " and TSPL_SRN_HEAD.Vendor_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")" 'ADDED BY ABHISHEK AS ON 30 AUG 2012
                End If
                qry += " )xxx group by SRN_No,Item_Code "
                qry += " ) XXXX LEFT OUTER JOIN TSPL_COMPANY_MASTER ON XXXX.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code"
                qry += " Left Outer Join TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=XXXX.Vendor_Code order by Item_Desc"


                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    common.clsCommon.MyMessageBoxShow("No Record Found")
                Else
                    'PurchaseOrderViewer.funreport(dt, EnumTecxpertPaperSize.PaperSize10x6, "rptSRNCustomReport", "SRN Report")
                    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "Vizag") = CompairStringResult.Equal Then
                        frmCRV.funreport(Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptSRN_4_FG_Vizag", "SRN Report")
                    Else
                        frmCRV.funreport(Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptSRN_4_FG_Guntur", "SRN Report")
                    End If
                End If
            Else ''For RM Other Print out
                Dim strquery As String = "SELECT TSPL_GRN_HEAD.Ref_No, Comp_Code1, "
                If isCancel Then
                    strquery += " 'Cancelled' As Report_Status, "
                Else
                    strquery += " '' As Report_Status, "
                End If
                strquery += " convert(Varchar,TSPL_TENDER_HEADER.DocumentDate,103) AS TendorDocumentDate,"

                strquery += " TSPL_GRN_HEAD.GRN_Date, " + TSPL_SRN_HEAD + ".Against_QC_Code," + TSPL_SRN_HEAD + ".Against_QC_Date,isnull(TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No,'') as PurchaseOrder_No,isnull(convert(varchar,TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date,103),'') as PurchaseOrder_Date,TSPL_ITEM_MASTER.HSN_Code,tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo ,TSPL_VENDOR_MASTER.GSTFinalNo AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode, " + TSPL_SRN_HEAD + ".Against_MRN," + TSPL_SRN_DETAIL + ".short_Qty, Location_Code,convert(varchar," + TSPL_SRN_HEAD + ".Challan_Date,103) as CHA_Date," + TSPL_SRN_DETAIL + ".Specification as Det_Specification," + TSPL_SRN_DETAIL + ".Remarks  as DetRemarks ,TSPL_VENDOR_MASTER.CST as vndr_cst,TSPL_VENDOR_MASTER.Tin_No as vndr_tin,TSPL_LOCATION_MASTER.CST_No as location_cst," + TSPL_SRN_HEAD + ".Ship_To_Location,TSPL_SHIP_TO_LOCATION.Ship_To_Desc,(TSPL_SHIP_TO_LOCATION.Add1+' '+TSPL_SHIP_TO_LOCATION.add2+' '+TSPL_SHIP_TO_LOCATION.add3) as ship_addr,TSPL_SHIP_TO_LOCATION.City_Code as ship_city,TSPL_SHIP_TO_LOCATION.State as ship_state," + TSPL_SRN_DETAIL + ".MRP,Location_Desc  as Location_Company ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end +case when LEN(isnull(TSPL_LOCATION_MASTER.add4,''))>0 then '- '+isnull(TSPL_LOCATION_MASTER.add4,'') else ' ' end   as company_address,TSPL_LOCATION_MASTER.City_Code as TSPL_LOCATION_MASTER_City_Code ,TSPL_LOCATION_MASTER.State as TSPL_LOCATION_MASTER_state ,TSPL_LOCATION_MASTER.Country as TSPL_LOCATION_MASTER_country, TSPL_VENDOR_MASTER.add1 +case when len(TSPL_VENDOR_MASTER.add2)>0 then ', '+TSPL_VENDOR_MASTER.add2 else '' end +case when LEN(isnull(TSPL_VENDOR_MASTER.Add3,''))>0 then ', '+isnull(TSPL_VENDOR_MASTER.Add3,'') else ' ' end   as vendor_address, TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end+case when LEN(isnull(TSPL_LOCATION_MASTER.Add4,''))>0 then '- '+isnull(TSPL_LOCATION_MASTER.Add4,'') else ' ' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Location_Desc,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Location_Desc,'') else ' ' end    as address1, isnull(TSPL_LOCATION_MASTER.Location_Desc,'') as Location_Desc_Heading,TSPL_LOCATION_MASTER.add1 as Location_Add1_heading,  TSPL_LOCATION_MASTER.add2 as Location_Add2_heading, TSPL_LOCATION_MASTER.Add3  as Location_Add3_heading,TSPL_LOCATION_MASTER.Add4  as Location_Add4_heading , TSPL_LOCATION_MASTER.IsMainPlant ,TSPL_LOCATION_MASTER .TIN_No ,( case when TSPL_LOCATION_MASTER.Phone2 <> '' then TSPL_LOCATION_MASTER.Phone1 +','+TSPL_LOCATION_MASTER.Phone2 else TSPL_LOCATION_MASTER.Phone1 end) as Location_Phn, " + TSPL_SRN_HEAD + ".Description," + TSPL_SRN_HEAD + ".Comments, user_master1.User_Name as Created_By,user_master2.User_Name as Modify_By, " + TSPL_SRN_HEAD + ".SRN_No, " + TSPL_SRN_HEAD + ".SRN_Date,TSPL_VENDOR_MASTER.Vendor_Name, TSPL_VENDOR_MASTER.Tin_No as Vendor_Tin_No, TSPL_VENDOR_MASTER.Phone1 as Vendor_Contact, (case when len(against_mrn)>0 then (select MRN_Date  from tspl_mrn_head where tspl_mrn_head.MRN_No =against_mrn) else SRN_Date end ) as Challan_Date, " + TSPL_SRN_HEAD + ".Ref_No  " &
                      "as Challan_No, " + TSPL_SRN_HEAD + ".Inv_No, convert(varchar," + TSPL_SRN_HEAD + ".Inv_Date,103) as Inv_Date, " + TSPL_SRN_HEAD + ".GRNo," + TSPL_SRN_HEAD + ".Amount_Less_Discount ," + TSPL_SRN_HEAD + ".GENo," + TSPL_SRN_HEAD + ".SRN_Total_Amt, " &
                      "convert(varchar," + TSPL_SRN_HEAD + ".GEDate,103) as GEDate, " + TSPL_SRN_HEAD + ".VehicleNo," + TSPL_SRN_HEAD + ".SRN_No, " + TSPL_SRN_HEAD + ".Carrier," + TSPL_SRN_HEAD + ".Remarks," + TSPL_SRN_DETAIL + ".Landed_Cost_Rate," + TSPL_SRN_DETAIL + ".Landed_Cost_Amount , " + TSPL_SRN_DETAIL + ".Item_Code," + TSPL_SRN_DETAIL + ".UOM_WEIGHT," + TSPL_SRN_DETAIL + ".UOM_WEIGHT_VALUE," + TSPL_SRN_DETAIL + ".Row_Type," + TSPL_SRN_DETAIL + ".Amt_Less_Discount," &
"" + TSPL_SRN_DETAIL + ".Item_Cost as basicRate," + TSPL_SRN_DETAIL + ".Item_Net_Amt as BasicTotal," + TSPL_SRN_DETAIL + ".Unit_Cost_Tax_Rate as UCTR," &
"" + TSPL_SRN_DETAIL + ".Unit_Cost_Tax as uctax," + TSPL_SRN_DETAIL + ".Item_Desc," + TSPL_SRN_DETAIL + ".Unit_code," + TSPL_SRN_DETAIL + ".SRN_Qty," + TSPL_SRN_DETAIL + ".Rejected_Qty," + TSPL_SRN_DETAIL + ".Short_Qty," + TSPL_SRN_DETAIL + ".GRN_Qty," + TSPL_SRN_HEAD + ".Vendor_Code," + TSPL_SRN_HEAD + ".SRN_Total_Amt," + TSPL_SRN_DETAIL + ".ITEM_COST," &
 "TSPL_VENDOR_MASTER.Add1 as venAdd1, TSPL_VENDOR_MASTER.Add2 as vanadd2, TSPL_VENDOR_MASTER.Add3 as venadd3, " &
"tax1.Tax_Code_Desc as tax1name,isnull (" + TSPL_SRN_HEAD + ".tax1_amt,0) as txt1amt,tax2.Tax_Code_Desc as tax2name," &
"isnull (" + TSPL_SRN_HEAD + ".tax2_amt,0) as txt2amt,tax3.Tax_Code_Desc as tax3name,isnull (" + TSPL_SRN_HEAD + ".tax3_amt,0) as txt3amt," &
"tax4.Tax_Code_Desc as tax4name,isnull (" + TSPL_SRN_HEAD + ".tax4_amt,0) as txt4amt,tax5.Tax_Code_Desc as tax5name," &
"isnull (" + TSPL_SRN_HEAD + ".tax5_amt,0) as txt5amt,tax6.Tax_Code_Desc as tax6name,isnull (" + TSPL_SRN_HEAD + ".tax6_amt,0) as txt6amt " &
",tax7.Tax_Code_Desc as tax7name,isnull (" + TSPL_SRN_HEAD + ".tax7_amt,0) as txt7amt,tax8.Tax_Code_Desc as tax8name," &
"isnull (" + TSPL_SRN_HEAD + ".tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,isnull (" + TSPL_SRN_HEAD + ".tax9_amt,0) as txt9amt," &
"tax10.Tax_Code_Desc as tax10name,isnull (" + TSPL_SRN_HEAD + ".tax10_amt,0) as txt10amt, TSPL_COMPANY_MASTER.Comp_Name as compname,'" & objCommonVar.CurrentUser & "' as User_Name, " &
"TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2," + TSPL_SRN_DETAIL + ".SRN_Qty," &
"case when tax1.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax1_amt else null end as Tax1Recoverable," &
"case when tax2.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".TAX2_Amt else null end as Tax2Recoverable, " &
"case when tax3.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax3_amt else null end as Tax3Recoverable, " &
"case when tax4.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax4_amt else null end as Tax4Recoverable, " &
"case when tax5.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax5_amt else null end as Tax5Recoverable, " &
"case when tax6.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax6_amt else null end as Tax6Recoverable," &
"case when tax7.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax7_amt else null end as Tax7Recoverable, " &
"case when tax8.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax8_amt else null end as Tax8Recoverable, " &
"case when tax9.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax9_amt else null end as Tax9Recoverable," &
"case when tax10.Tax_Recoverable='Y' then " + TSPL_SRN_HEAD + ".tax10_amt else null end as Tax10Recoverable, " &
"" + TSPL_SRN_HEAD + ".TAX1," + TSPL_SRN_HEAD + ".TAX2," + TSPL_SRN_HEAD + ".TAX3," + TSPL_SRN_HEAD + ".TAX4," + TSPL_SRN_HEAD + ".TAX5," + TSPL_SRN_HEAD + ".tax6," &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX1_Rate ,0),103)+'%' as txt1Rate," &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX2_Rate   ,0),103)+'%' as txt2Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX3_Rate  ,0),103)+'%' as txt3Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX4_Rate  ,0),103)+'%' as txt4Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX5_Rate  ,0),103)+'%' as txt5Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX6_Rate  ,0),103)+'%' as txt6Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX7_Rate  ,0),103)+'%' as txt7Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX8_Rate  ,0),103)+'%' as txt8Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX9_Rate  ,0),103)+'%' as txt9Rate, " &
"convert(varchar,isnull (" + TSPL_SRN_HEAD + ".TAX10_Rate  ,0),103)+'%' as txt10Rate," &
"" + TSPL_SRN_DETAIL + ".Amt_Less_Discount as Value,(select SUM(rejected_qty) from " + TSPL_SRN_DETAIL + " where srn_no=" + TSPL_SRN_HEAD + ".SRN_No) as Rej_qty, (select SUM(TSPL_MRN_DETAIL.MRN_Qty) from TSPL_SRN_DETAIl left outer join TSPL_MRN_DETAIL on TSPL_MRN_DETAIL .MRN_No=" + TSPL_SRN_DETAIL + ".MRN_Id and TSPL_MRN_DETAIL.Item_Code=" + TSPL_SRN_DETAIL + ".Item_Code left outer join TSPL_MRN_HEAD on TSPL_MRN_HEAD.MRN_No=TSPL_MRN_DETAIL.MRN_No where SRN_No =" + TSPL_SRN_HEAD + ".SRN_No and TSPL_MRN_HEAD.IsCancel=0 )as MrnTotQty, (select SUM(SRN_qty) from tspl_srn_detail where srn_no=" + TSPL_SRN_HEAD + ".SRN_No) as SRNQtyTotal, (select case when COUNT(xxx.PI_No)>1 then Min(xxx.PI_No)+ ' *' else Min(xxx.PI_No)end as PINO from" &
" ( select TSPL_PI_DETAIL.PI_No from TSPL_PI_DETAIL  where  TSPL_PI_DETAIL.SRN_Id= " + TSPL_SRN_HEAD + ".SRN_No " &
" GROUP by TSPL_PI_DETAIL.PI_No)xxx) as PInvNo  ,    " &
       " " + TSPL_SRN_HEAD + ".Add_Charge_Name1 as Add1Name, " &
     " " + TSPL_SRN_HEAD + ".Add_Charge_Amt1 as Add1 , " &
     "     " + TSPL_SRN_HEAD + ".Add_Charge_Name2 as Add2Name, " &
     "   " + TSPL_SRN_HEAD + ".Add_Charge_Amt2 as Add2 , " &
     "    " + TSPL_SRN_HEAD + ".Add_Charge_Name3 as Add3Name, " &
     "   " + TSPL_SRN_HEAD + ".Add_Charge_Amt3 as Add3 , " &
     "    " + TSPL_SRN_HEAD + ".Add_Charge_Name4 as Add4Name, " &
     "    " + TSPL_SRN_HEAD + ".Add_Charge_Amt4 as Add4 , " &
     "     " + TSPL_SRN_HEAD + ".Add_Charge_Name5 as Add5Name, " &
      "     " + TSPL_SRN_HEAD + ".Add_Charge_Amt5 as Add5 , " &
      "     " + TSPL_SRN_HEAD + ".Add_Charge_Name6 as Add6Name, " &
      "    " + TSPL_SRN_HEAD + ".Add_Charge_Amt6 as Add6 , " &
      "    " + TSPL_SRN_HEAD + ".Add_Charge_Name7 as Add7Name, " &
      "     " + TSPL_SRN_HEAD + ".Add_Charge_Amt7 as Add7 , " &
      "       " + TSPL_SRN_HEAD + ".Add_Charge_Name8 as Add8Name, " &
      "      " + TSPL_SRN_HEAD + ".Add_Charge_Amt8 as Add8 , " &
       "      " + TSPL_SRN_HEAD + ".Add_Charge_Name9 as Add9Name, " &
       "      " + TSPL_SRN_HEAD + ".Add_Charge_Amt9 as Add9 , " &
       "      " + TSPL_SRN_HEAD + ".Add_Charge_Name10 as Add10Name, " &
       "     " + TSPL_SRN_HEAD + ".Add_Charge_Amt10 as Add10," + TSPL_SRN_HEAD + ".Against_RGP," + TSPL_SRN_DETAIL + " .Specification ," + TSPL_SRN_HEAD + ".Against_Requisition ," + TSPL_SRN_DETAIL + ".PO_Qty," + TSPL_SRN_DETAIL + ".GRN_Qty," + TSPL_SRN_DETAIL + ".MRN_Qty," + TSPL_SRN_DETAIL + ".PO_id," + TSPL_SRN_DETAIL + ".Req_No," + TSPL_SRN_HEAD + ".Against_GRN," + TSPL_SRN_HEAD + ".Form_38, "

                If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "Vizag") = CompairStringResult.Equal Then
                    strquery += " (Case when len(" + TSPL_SRN_HEAD + ".Against_PO)>0 then " + TSPL_SRN_HEAD + ".Against_PO else " + TSPL_SRN_HEAD + ".Against_Requisition end) as Against_PO "
                Else
                    strquery += " " + TSPL_SRN_HEAD + ".Against_PO "
                End If
                strquery += " " & QryShowStatus & " "
                strquery += ", case when tax1.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax1_amt else null end as Tax1NonRecoverable," &
                " case when tax2.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".TAX2_Amt else null end as Tax2NonRecoverable, " &
                " case when tax3.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax3_amt else null end as Tax3NonRecoverable, " &
                " case when tax4.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax4_amt else null end as Tax4NonRecoverable, " &
                " case when tax5.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax5_amt else null end as Tax5NonRecoverable, " &
                " case when tax6.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax6_amt else null end as Tax6NonRecoverable," &
                " case when tax7.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax7_amt else null end as Tax7NonRecoverable, " &
                " case when tax8.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax8_amt else null end as Tax8NonRecoverable, " &
                " case when tax9.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax9_amt else null end as Tax9NonRecoverable," &
                " case when tax10.Tax_Recoverable='N' then " + TSPL_SRN_HEAD + ".tax10_amt else null end as Tax10NonRecoverable"
                strquery += " ," + TSPL_SRN_HEAD + ".Dept_desc
,CASE WHEN " + TSPL_SRN_DETAIL + ".disc_amt>0 THEN CONVERT(DECIMAL(18,3), (case when srn_qty>0 then ((" + TSPL_SRN_DETAIL + ".amount-" + TSPL_SRN_DETAIL + ".disc_amt)/srn_qty) else 0 end)) ELSE " + TSPL_SRN_DETAIL + ".ITEM_COST END as ITEM_COST_AFTER_DISC " + Environment.NewLine +
"FROM  " + TSPL_SRN_DETAIL + " INNER JOIN " + TSPL_SRN_HEAD + " ON " + TSPL_SRN_DETAIL + ".SRN_No = " + TSPL_SRN_HEAD + ".SRN_No " &
    "INNER JOIN TSPL_COMPANY_MASTER ON " + TSPL_SRN_HEAD + ".Comp_Code = TSPL_COMPANY_MASTER.Comp_Code  " &
    "INNER JOIN TSPL_VENDOR_MASTER ON " + TSPL_SRN_HEAD + ".Vendor_Code = TSPL_VENDOR_MASTER.Vendor_Code " &
    "left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =" + TSPL_SRN_HEAD + ".tax1  " &
    "left outer join tspl_tax_master as tax2 on tax2.tax_code = " + TSPL_SRN_HEAD + ".tax2 " &
    "left outer join tspl_tax_master as tax3 on tax3.Tax_Code=" + TSPL_SRN_HEAD + " .TAX3 " &
    "left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= " + TSPL_SRN_HEAD + " .tax4 " &
    "left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=" + TSPL_SRN_HEAD + " .tax5 " &
    "left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =" + TSPL_SRN_HEAD + " .TAX6  " &
    "left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =" + TSPL_SRN_HEAD + " .TAX7  " &
    "left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =" + TSPL_SRN_HEAD + " .TAX8 " &
    "left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =" + TSPL_SRN_HEAD + " .TAX9 " &
    " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =" + TSPL_SRN_HEAD + " .TAX10  " &
    "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=" + TSPL_SRN_HEAD + ".Bill_To_Location  " &
    "left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=" + TSPL_SRN_HEAD + ".Ship_To_Location " &
    " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code= " + TSPL_SRN_DETAIL + ".Item_Code " &
    " left outer join tspl_state_master as tspl_state_master_for_location_state on tspl_state_master_for_location_state.state_code=tspl_location_master.state  " &
    "  left outer join TSPL_STATE_MASTER on TSPL_VENDOR_MASTER.State_Code= TSPL_STATE_MASTER.State_Code " &
    " left outer join tspl_user_master as user_master1 on user_master1.user_code=" + TSPL_SRN_HEAD + ".Created_By  " &
    " left join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=" + TSPL_SRN_HEAD + ".Against_GRN  " &
    " left outer join tspl_user_master as user_master2 on user_master2.user_code=" + TSPL_SRN_HEAD + ".Modify_By " &
    " left join TSPL_PURCHASE_ORDER_HEAD on isnull(TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No,'') =isnull(" + TSPL_SRN_DETAIL + ".PO_ID,'') "

                strquery += " Left outer join TSPL_TENDER_HEADER on TSPL_TENDER_HEADER.DocumentCode=TSPL_PURCHASE_ORDER_HEAD.RefTendorNo"
                strquery += " where 2=2"

                If FromDate.HasValue AndAlso ToDate.HasValue Then
                    strquery += " And Convert(date," + TSPL_SRN_HEAD + ".SRN_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date," + TSPL_SRN_HEAD + ".SRN_Date,103)<=Convert(date,'" + ToDate + "',103) "

                End If
                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
                    strquery += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
                End If
                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
                    strquery += " and " + TSPL_SRN_HEAD + ".SRN_No in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
                End If
                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
                    strquery += " and " + TSPL_SRN_HEAD + ".Vendor_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")  "

                End If

                strquery = strquery + " order by " + TSPL_SRN_DETAIL + ".line_no"

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strquery)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    common.clsCommon.MyMessageBoxShow("No Record Found")
                Else
                    If clsCommon.CompairString(clsCommon.myCstr(objCommonVar.CurrentCompanyCode), "KL") <> CompairStringResult.Equal Then
                        clsSRNHead.SetItemWiseTax(dt, ArrSrnNo(0))
                    End If

                    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "Vizag") = CompairStringResult.Equal Then
                        frmCRV.funreport(Form_ID, CrystalReportFolder.PurchaseOrder, dt, "SRNReportThroughReport", "Store Receipt Report")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
                        If IsSRNReportQtyWise Then
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.PurchaseOrder, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "SRNReportThroughReportQtyWise", "Store Receipt Report", clsCommon.myCDate(TXTDATE), "rptCompanyAddress.rpt")

                        Else
                            frmCRV.funreport(Form_ID, CrystalReportFolder.PurchaseOrder, dt, "SRNReportThroughReport-G", "Store Receipt Report", clsCommon.myCDate(TXTDATE))
                        End If
                    Else
                        If IsSRNReportQtyWise Then
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.PurchaseOrder, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "SRNReportThroughReportQtyWise", "Store Receipt Report", clsCommon.myCDate(TXTDATE), "rptCompanyAddress.rpt")
                        ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.PurchaseOrder, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "SRNReportThroughReport-G-SKR", "Store Receipt Report", clsCommon.myCDate(clsCommon.GETSERVERDATE()), "rptCompanyAddress.rpt", "SubRptCmpnyMasterForERODE.rpt", clsERPFuncationality.CompanyAddresShowinHeaderPartForERODE()) 'update by preeti gupta Against Ticket No[ADV/27/07/18-000036]

                            ' frmCRV.funreport(CrystalReportFolder.PurchaseOrder, dt, "SRNReportThroughReport-G-SKR", "Store Receipt Report", clsCommon.myCDate(txtDate.Value))
                        Else
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.PurchaseOrder, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "SRNReportThroughReport-G", "Store Receipt Report", clsCommon.myCDate(clsCommon.GETSERVERDATE()), "rptCompanyAddress.rpt", "SubRptCmpnyMasterForERODE.rpt", clsERPFuncationality.CompanyAddresShowinHeaderPartForERODE()) 'update by preeti gupta Against Ticket No[ADV/27/07/18-000036]


                        End If


                        'SRNReportThroughReportQtyWise
                        'PurchaseOrderViewer.funreport(dt, "SRNReportThroughReport-G", "Store Receipt Report")
                    End If
                End If
            End If
            frmCRV = Nothing
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    '==============added by preeti gupta against ticket no[TEC/23/05/18-000253]
    Public Shared Function HistoryUpdate(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_SRN_HEAD", "SRN_No", "TSPL_SRN_DETAIL", "SRN_No", trans)
        Return True
    End Function
    Public Shared Function CheckRalPenalty(ByVal strcodeNo As String, ByVal trans As SqlTransaction) As Boolean
        'Dim qry As String = "select sum(fin.[cnt]) from (SELECT 1 as [cnt] from TSPL_SRN_DETAIL where TSPL_SRN_DETAIL.PO_ID ='" + clsCommon.myCstr(strPONo) + "' union all SELECT 1 as [cnt] from TSPL_GRN_DETAIL where TSPL_GRN_DETAIL.PO_ID ='" + clsCommon.myCstr(strPONo) + "')fin"
        '        Dim qry As String = "select sum(fin.[cnt]) from ( select  1 as [cnt]  from TSPL_TENDER_PENALTY_DETAIL 
        'inner join TSPL_PI_DETAIL on TSPL_PI_DETAIL.SRN_Id=TSPL_TENDER_PENALTY_DETAIL.SRN_No  
        'where TSPL_TENDER_PENALTY_DETAIL.Document_No='" + clsCommon.myCstr(strcodeNo) + "')fin"
        Dim qry As String = "select sum(fin.[cnt]) from ( select  1 as [cnt]  from TSPL_TENDER_PENALTY_DETAIL 
where TSPL_TENDER_PENALTY_DETAIL.SRN_No='" + clsCommon.myCstr(strcodeNo) + "')fin
"
        Dim count As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))

        If count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function CancelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal Document_Type As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            CancelData(Form_Id, Doc_No, Document_Type, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function CancelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal Document_Type As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        Try
            Dim Doc_Type As String = Nothing

            Dim obj As clsSRNHead = clsSRNHead.GetData(Doc_No, NavigatorType.Current, trans, Document_Type)
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePurchase, clsUserMgtCode.mbtnSRN, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location), obj.SRN_Date, trans)
            If obj Is Nothing OrElse clsCommon.myLen(obj.SRN_No) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If
            clsItemLocationDetails.CheckCancelInventoryBalance(Form_Id, Doc_No, trans)

            '' transfer data into cancel table

            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_SRN_HEAD", "SRN_No", "TSPL_SRN_DETAIL", "SRN_No", trans)
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL", "SRN_No", trans)
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_CFORM_ISSUE_RECEIVE_DETAIL", "SRN_No", trans)


            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No='" & Doc_No & "'"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then

                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            clsBatchInventory.DeleteData("SRN", Doc_No, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)

            qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_RGP_BOM_DETAIL where SRN_No='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_CFORM_ISSUE_RECEIVE_DETAIL where SRN_No='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL where SRN_No='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)


            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_SRN_DEDUCTION_SECURITY where SRN_No='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_QC_CHECK_MRN_DETAIL where SRN_Id='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_SRN_TENDER_CALC where SRN_No='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_QC_CHECK_APPROVAL_ENTRY where SRN_id='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_SRN_DEDUCTION where SRN_No='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_SRN_DETAIL where SRN_No='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            If clsCommon.myLen(Doc_No) > 0 Then
                qry = "delete from TSPL_SRN_Head where SRN_No='" & Doc_No & "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

            '' release objects 
            obj = Nothing
            qry = Nothing

        Catch ex As Exception

            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    '=======================================


    Public Function SaveData(ByVal obj As clsSRNHead, ByVal isNewEntry As Boolean) As Boolean
        ShowItemAllStructureWise = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ShowItemAllStructureWise, clsFixedParameterCode.ShowItemAllStructureWise, Nothing)) = "1", True, False))

        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(obj, isNewEntry, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Function SaveData(ByVal obj As clsSRNHead, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean

        Dim isSaved As Boolean = True
        ' Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()

        Try
            clsSRNAdditionChargeInsurance.DeleteData(obj.SRN_No, trans)
            clsSerializeInvenotry.DeleteData("SRN", obj.SRN_No, trans)
            clsBatchInventory.DeleteData("SRN", obj.SRN_No, trans)

            'If Not isNewEntry Then
            '    HistoryUpdate(obj.SRN_No, trans)
            'End If
            Dim qry As String = "delete from TSPL_SRN_DETAIL where SRN_No='" + obj.SRN_No + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim strDocNo As String = ""
            If isNewEntry Then
                If clsCommon.CompairString(obj.Document_Type, "SRN") = CompairStringResult.Equal Then
                    If obj.isJobWorkOutward = 1 Then
                        obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.SRN, clsDocTransactionType.POJobWorkOutward, obj.Sublocation_Code)
                    Else
                        Dim isPODocumentTypeWise As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PurchaseCounterOnTransactionType, clsFixedParameterCode.PurchaseCounterOnTransactionType, trans)) = 0, False, True) ''Make Setting Balwinder
                        clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePurchase, clsUserMgtCode.mbtnSRN, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location), obj.SRN_Date, trans)

                        If isPODocumentTypeWise Then
                            If clsCommon.CompairString(obj.PurchaseOrder_Type, "J") = CompairStringResult.Equal Then
                                obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.SRN, clsDocTransactionType.POJobWork, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location))
                            ElseIf clsCommon.CompairString(obj.PurchaseOrder_Type, "I") = CompairStringResult.Equal Then
                                obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.SRN, clsDocTransactionType.POImport, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location))
                            ElseIf clsCommon.CompairString(obj.PurchaseOrder_Type, "L") = CompairStringResult.Equal Then
                                obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.SRN, clsDocTransactionType.PODomestic, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location))
                            Else
                                Throw New Exception("Type is Not Correct To Generate the Transaction Code")
                            End If
                        Else
                            If clsCommon.myLen(obj.Against_RGP) > 0 Then
                                obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.SRN, clsDocTransactionType.SRNRGP, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location))
                            Else
                                If ShowItemAllStructureWise Then
                                    obj.SRN_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.SRN_Date), clsDocType.SRN, clsDocTransactionType.All, obj.Bill_To_Location)
                                Else
                                    Dim TransType = clsDBFuncationality.getSingleValue("SELECT PREFIX_CODE FROM TSPL_ITEM_TYPE_MASTER WHERE ITEM_TYPE_CODE='" + obj.Item_Type + "'", trans)
                                    obj.SRN_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.SRN_Date), clsDocType.SRN, TransType, obj.Bill_To_Location)
                                End If

                                If clsCommon.CompairString(obj.SRN_No, String.Empty) = CompairStringResult.Equal Then
                                    Throw New Exception("Item Type is Not Correct To Generate the Transaction Code")
                                End If
                            End If
                        End If
                    End If
                ElseIf clsCommon.CompairString(obj.Document_Type, "MT") = CompairStringResult.Equal Then
                    If clsCommon.myLen(obj.Against_RGP) > 0 Then
                        obj.SRN_No = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.MTSRN, clsDocTransactionType.SRNRGP, obj.Bill_To_Location)
                    Else
                        Dim TransType = clsDBFuncationality.getSingleValue("SELECT PREFIX_CODE FROM TSPL_ITEM_TYPE_MASTER WHERE ITEM_TYPE_CODE='" + obj.Item_Type + "'", trans)
                        obj.SRN_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.SRN_Date), clsDocType.MTSRN, TransType, obj.Bill_To_Location)
                        If clsCommon.CompairString(obj.SRN_No, String.Empty) = CompairStringResult.Equal Then
                            Throw New Exception("Item Type is Not Correct To Generate the Transaction Code")
                        End If
                    End If
                End If
            End If
            If (clsCommon.myLen(obj.SRN_No) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "AutoSRN", obj.autosrnfromrgp)
            Dim ServerDate As DateTime = Nothing
            If isNewEntry Then
                ServerDate = clsCommon.GETSERVERDATE(trans)
                obj.SRN_Date = New DateTime(obj.SRN_Date.Year, obj.SRN_Date.Month, obj.SRN_Date.Day, ServerDate.Hour, ServerDate.Minute, ServerDate.Second)
            End If
            clsCommon.AddColumnsForChange(coll, "SRN_Date", clsCommon.GetPrintDate(obj.SRN_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Vendor_Code", obj.Vendor_Code)
            clsCommon.AddColumnsForChange(coll, "Vendor_Name", obj.Vendor_Name)
            clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Internal", IIf(obj.Is_Internal, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Inter_unit_purchase", IIf(obj.Inter_unit_purchase, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Ref_No", obj.Ref_No)
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            clsCommon.AddColumnsForChange(coll, "Inv_No", obj.Inv_No)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Bill_To_Location", obj.Bill_To_Location)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Location", obj.Ship_To_Location)
            clsCommon.AddColumnsForChange(coll, "Sublocation_Code", obj.Sublocation_Code)
            clsCommon.AddColumnsForChange(coll, "Against_CF_Sale_Purchase_No", obj.Against_CF_Sale_Purchase_No, True)
            clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
            clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
            clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
            clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
            clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
            clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
            clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
            clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
            clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
            clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
            clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
            clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
            clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", obj.Total_Tax_Amt)
            clsCommon.AddColumnsForChange(coll, "Discount_Base", obj.Discount_Base)
            clsCommon.AddColumnsForChange(coll, "Discount_Amt", obj.Discount_Amt)
            clsCommon.AddColumnsForChange(coll, "Amount_Less_Discount", obj.Amount_Less_Discount)
            clsCommon.AddColumnsForChange(coll, "Total_Taxable_Amount", obj.Total_Taxable_Amount)
            clsCommon.AddColumnsForChange(coll, "SRN_Total_Amt", obj.SRN_Total_Amt)
            ''clsCommon.AddColumnsForChange(coll, "TotalUnit_Cost_Tax", obj.TotalUnit_Cost_Tax)
            clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)

            If clsCommon.myLen(obj.Due_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Due_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "Confirmatory_PO", IIf(obj.Confirmatory_PO, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Rgp_Type", obj.RGP_Type)
            clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))

            clsCommon.AddColumnsForChange(coll, "Carrier", obj.Carrier)
            clsCommon.AddColumnsForChange(coll, "VehicleNo", obj.VehicleNo)
            clsCommon.AddColumnsForChange(coll, "GRNo", obj.GRNo)
            clsCommon.AddColumnsForChange(coll, "GENo", obj.GENo)
            clsCommon.AddColumnsForChange(coll, "Dept", obj.Dept)
            clsCommon.AddColumnsForChange(coll, "Dept_Desc", obj.Dept_Desc)
            clsCommon.AddColumnsForChange(coll, "Item_Type", obj.Item_Type)

            clsCommon.AddColumnsForChange(coll, "Against_Requisition", obj.Against_Requisition, True)
            clsCommon.AddColumnsForChange(coll, "Against_PO", obj.Against_PO, True)
            clsCommon.AddColumnsForChange(coll, "Against_GRN", obj.Against_GRN, True)
            clsCommon.AddColumnsForChange(coll, "Against_MRN", obj.Against_MRN, True)
            clsCommon.AddColumnsForChange(coll, "Against_RGP", obj.Against_RGP, True)
            clsCommon.AddColumnsForChange(coll, "Against_Schedule_Code", obj.Against_Schedule_Code, True)
            clsCommon.AddColumnsForChange(coll, "PurchaseOrder_Type", obj.PurchaseOrder_Type)
            'clsCommon.AddColumnsForChange(coll, "Landed_Add_Cost", obj.Landed_Add_Cost)
            clsCommon.AddColumnsForChange(coll, "TotalUnit_Cost_Tax", obj.Total_Landed_Cost)
            clsCommon.AddColumnsForChange(coll, "Document_Type", obj.Document_Type)
            If obj.GEDate.HasValue Then
                clsCommon.AddColumnsForChange(coll, "GEDate", clsCommon.GetPrintDate(obj.GEDate, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "GEDate", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "PROJECT_ID", obj.PROJECT_ID, True)

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
            clsCommon.AddColumnsForChange(coll, "Form_38", obj.Form_38)

            clsCommon.AddColumnsForChange(coll, "Total_Add_Charge", obj.Total_Add_Charge)
            clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
            If clsCommon.myLen(obj.Challan_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Challan_Date", clsCommon.GetPrintDate(obj.Challan_Date, "dd/MMM/yyyy"))
            End If
            If clsCommon.myLen(obj.Inv_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Inv_Date", clsCommon.GetPrintDate(obj.Inv_Date, "dd/MMM/yyyy"))
            End If

            clsCommon.AddColumnsForChange(coll, "is_Excise_On_Qty", IIf(obj.is_Excise_On_Qty, 1, 0))
            clsCommon.AddColumnsForChange(coll, "is_RGP_Non_Inventory", IIf(obj.is_RGP_Non_Inventory, 1, 0))
            clsCommon.AddColumnsForChange(coll, "is_QCAccepted", IIf(obj.is_QCAccepted, 1, 0))

            clsCommon.AddColumnsForChange(coll, "AssessableAmt", obj.AssessableAmt)

            '' currencyconversion
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            ''richa agarwal against ticket no. BM00000006368 on 28/04/2015
            If obj.ApplicableFrom.HasValue Then
                clsCommon.AddColumnsForChange(coll, "ApplicableFrom", clsCommon.GetPrintDate(obj.ApplicableFrom, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "ApplicableFrom", Nothing, True)
            End If
            ''-----------------------------------
            '' End currencyconversion

            clsCommon.AddColumnsForChange(coll, "IsAbatementPO", obj.IsAbatementPO)
            'clsCommon.AddColumnsForChange(coll, "Issue_Road_Permit", obj.against_roadpermit)
            clsCommon.AddColumnsForChange(coll, "Issue_C_Form", obj.agnst_cform)
            clsCommon.AddColumnsForChange(coll, "Retention", obj.Retention)

            clsCommon.AddColumnsForChange(coll, "Against_QC_Code", obj.Against_QC_Code)
            clsCommon.AddColumnsForChange(coll, "Against_QC_Date", obj.Against_QC_Date)

            clsCommon.AddColumnsForChange(coll, "Total_Accepted_Amount", obj.Total_Accepted_Amount)
            clsCommon.AddColumnsForChange(coll, "Total_Rejected_Amount", obj.Total_Rejected_Amount)
            clsCommon.AddColumnsForChange(coll, "Total_Shortage_Amount", obj.Total_Shortage_Amount)
            clsCommon.AddColumnsForChange(coll, "Total_Leak_Amount", obj.Total_Leak_Amount)
            clsCommon.AddColumnsForChange(coll, "Total_Burst_Amount", obj.Total_Burst_Amount)
            clsCommon.AddColumnsForChange(coll, "isJobWorkOutward", obj.isJobWorkOutward)
            clsCommon.AddColumnsForChange(coll, "Is_Shortage_Include_In_Landed_Cost", IIf(obj.Is_Shortage_Include_In_Landed_Cost, 1, 0))
            clsCommon.AddColumnsForChange(coll, "isExemptSecurityDedution", obj.isExemptSecurityDedution)
            'stuti
            If obj.RoadPermit_Date IsNot Nothing AndAlso clsCommon.myLen(obj.RoadPermit_Date) > 0 AndAlso IsDate(obj.RoadPermit_Date) Then
                clsCommon.AddColumnsForChange(coll, "RoadPermit_Date", clsCommon.GetPrintDate(obj.RoadPermit_Date, "dd/MMM/yyyy"))
            End If
            clsCommon.AddColumnsForChange(coll, "RoadPermit_No", obj.RoadPermit_No)
            '=======end here===========
            clsCommon.AddColumnsForChange(coll, "Total_Add_Charge_Insurance", obj.Total_Add_Charge_Insurance)
            clsCommon.AddColumnsForChange(coll, "Total_Item_Insurance_Amt", obj.Total_Item_Insurance_Amt)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "SRN_No", obj.SRN_No)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_HEAD", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_HEAD", OMInsertOrUpdate.Update, "TSPL_SRN_HEAD.SRN_No='" + obj.SRN_No + "'", trans)
            End If
            isSaved = isSaved AndAlso clsSRNDetail.SaveData(obj.SRN_No, obj.SRN_Date, obj, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.SRN_No, "TSPL_SRN_HEAD", "SRN_No", "TSPL_SRN_DETAIL", "SRN_No", trans)

            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.SRN_No, obj.arrCustomFields, trans)
            isSaved = isSaved AndAlso clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "SRN_No", obj.SRN_No, "TSPL_SRN_HEAD", trans)
            isSaved = isSaved AndAlso clsSRNRoadPermitDetail.SaveData_RoadPermit(obj.SRN_No, obj.Against_PO, obj.Arr_Road, trans)
            isSaved = isSaved AndAlso clsSRNCFORMDetail.SaveData_CFORM(obj.SRN_No, obj.Against_PO, obj.Arr_CFORM, trans)
            isSaved = isSaved AndAlso clsSRNAdditionChargeInsurance.SaveData(obj.SRN_No, obj.SRN_Date, obj.Arr_ACInsurance, trans)
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function IsPOQtyRecv(ByVal PONo As String, ByVal trans As SqlTransaction) As Boolean
        Dim check As Boolean
        check = True
        Try
            Dim qry As String = Nothing
            qry = "select sum(case when (TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty)<=(b.qty) then 0 else 1 end) as [check]" &
                  " from (select sum(TSPL_SRN_DETAIL.SRN_Qty) as qty,TSPL_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.PO_Id ,max(ISNULL(TSPL_ITEM_MASTER.tolerence ,0)) AS TOLERENCE" &
                  " from TSPL_SRN_DETAIL left join tspl_item_master on tspl_item_master.Item_Code = TSPL_SRN_DETAIL.Item_Code" &
                  " where TSPL_SRN_DETAIL.PO_Id ='" + clsCommon.myCstr(PONo) + "'   group by TSPL_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.PO_Id union all " &
                    " select 0 as qty,TSPL_PURCHASE_ORDER_DETAIL.Item_Code,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ,max(ISNULL(TSPL_ITEM_MASTER.tolerence ,0)) AS TOLERENCE from TSPL_PURCHASE_ORDER_DETAIL left join tspl_item_master on tspl_item_master.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code " &
                    " left join TSPL_SRN_DETAIL on TSPL_SRN_DETAIL.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code and TSPL_SRN_DETAIL.po_id=TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No where TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ='" + clsCommon.myCstr(PONo) + "'" &
                    " and isnull(TSPL_SRN_DETAIL.Item_Code,'')='' group by TSPL_PURCHASE_ORDER_DETAIL.Item_Code,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No) as b" &
                  " left outer join 
                  (SELECT TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No
				   ,TSPL_PURCHASE_ORDER_DETAIL.Item_Code
				   ,SUM(TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty) AS PurchaseOrder_Qty
				   FROM TSPL_PURCHASE_ORDER_DETAIL
				  WHERE TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ='" + clsCommon.myCstr(PONo) + "'
				  GROUP BY TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No
				   ,TSPL_PURCHASE_ORDER_DETAIL.Item_Code)
                  TSPL_PURCHASE_ORDER_DETAIL on b.PO_Id = TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No " &
                  " and b.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code " &
                  " where TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No='" + clsCommon.myCstr(PONo) + "' group by b.Item_Code"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If clsCommon.myCdbl(dr("check")) > 0 Then
                        check = check AndAlso False
                    Else
                        check = check AndAlso True
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return check
    End Function

    Public Shared Function IsPOQtyReceivedWithTolerance(ByVal PONo As String, ByVal trans As SqlTransaction) As Boolean
        Dim check As Boolean
        check = True
        Try
            Dim qry As String = Nothing
            qry = "select sum(case when (TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty)<=(b.qty) then 0 else 1 end) as [check]" &
                  " from (select sum(TSPL_SRN_DETAIL.SRN_Qty) as qty,TSPL_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.PO_Id ,max(ISNULL(TSPL_ITEM_MASTER.tolerence ,0)) AS TOLERENCE" &
                  " from TSPL_SRN_DETAIL left join tspl_item_master on tspl_item_master.Item_Code = TSPL_SRN_DETAIL.Item_Code" &
                  " where TSPL_SRN_DETAIL.PO_Id ='" + clsCommon.myCstr(PONo) + "'   group by TSPL_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.PO_Id union all " &
                    " select 0 as qty,TSPL_PURCHASE_ORDER_DETAIL.Item_Code,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ,max(ISNULL(TSPL_ITEM_MASTER.tolerence ,0)) AS TOLERENCE from TSPL_PURCHASE_ORDER_DETAIL left join tspl_item_master on tspl_item_master.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code " &
                    " left join TSPL_SRN_DETAIL on TSPL_SRN_DETAIL.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code and TSPL_SRN_DETAIL.po_id=TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No where TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ='" + clsCommon.myCstr(PONo) + "'" &
                    " and isnull(TSPL_SRN_DETAIL.Item_Code,'')='' group by TSPL_PURCHASE_ORDER_DETAIL.Item_Code,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No) as b" &
                  " left outer join 
                  (SELECT TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No
				   ,TSPL_PURCHASE_ORDER_DETAIL.Item_Code
				   ,SUM(TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty) + ((SUM(TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty) * max(ISNULL(TSPL_ITEM_MASTER.tolerence ,0)))/100) AS PurchaseOrder_Qty
				   FROM TSPL_PURCHASE_ORDER_DETAIL left join tspl_item_master on tspl_item_master.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code
				  WHERE TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No ='" + clsCommon.myCstr(PONo) + "'
				  GROUP BY TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No
				   ,TSPL_PURCHASE_ORDER_DETAIL.Item_Code)
                  TSPL_PURCHASE_ORDER_DETAIL on b.PO_Id = TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No " &
                  " and b.Item_Code = TSPL_PURCHASE_ORDER_DETAIL.Item_Code " &
                  " where TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No='" + clsCommon.myCstr(PONo) + "' group by b.Item_Code"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If clsCommon.myCdbl(dr("check")) > 0 Then
                        check = check AndAlso False
                    Else
                        check = check AndAlso True
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return check
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType, Optional ByVal IsMerchantTrade As String = "") As clsSRNHead
        Return GetData(strDocumentNo, NavType, Nothing, IsMerchantTrade)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction, Optional ByVal IsMerchantTrade As String = "") As clsSRNHead
        ' Dim obj As clsSRNHead = Nothing
        Dim obj As New clsSRNHead
        Dim qry As String = "SELECT TSPL_SRN_HEAD.*,TSPL_LOCATION_MASTER.Location_Desc as BillToLocationName,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as ShipToLocationName, " &
        " TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as TaxGroupName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_LOCATION_MASTER_SubLocation.Location_Desc as SubLocationName,TSPL_PURCHASE_ORDER_HEAD.RefTendorNo as TenderNo,TSPL_GRN_HEAD.GRN_Date ,IsNull(TSPL_MRN_HEAD.NIR_QC,0)NIR_QC,TSPL_QC_CHECK_HEAD.QC_Status,TSPL_SRN_HEAD.Against_CF_Sale_Purchase_No  " &
        " FROM TSPL_SRN_HEAD
        left outer join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=TSPL_SRN_HEAD.Against_GRN left outer join TSPL_MRN_HEAD On TSPL_MRN_HEAD.Against_GRN=TSPL_GRN_HEAD.GRN_No " &
        "Left Outer Join TSPL_QC_CHECK_HEAD On TSPL_QC_CHECK_HEAD.Gate_Entry_No=TSPL_MRN_HEAD.Against_GRN left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SRN_HEAD.Bill_To_Location " &
        "left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No = TSPL_SRN_HEAD.Against_PO " &
        " left outer join TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_SubLocation  on TSPL_LOCATION_MASTER_SubLocation.Location_Code=TSPL_SRN_HEAD.Sublocation_Code " &
        " left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SRN_HEAD.Ship_To_Location " &
        " left outer join  TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code= TSPL_SRN_HEAD.Tax_Group " &
        " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_SRN_HEAD.Terms_Code where 2=2"
        Dim whrCls As String = ""
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrCls = " AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        End If
        If clsCommon.CompairString(IsMerchantTrade, "MT") = CompairStringResult.Equal Then
            whrCls += " and TSPL_SRN_HEAD.Against_PO in ( Select TSPL_SRN_HEAD.Against_PO  from TSPL_SRN_HEAD left Outer Join TSPL_PURCHASE_ORDER_HEAD on TSPL_SRN_HEAD.Against_PO =TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No where TSPL_PURCHASE_ORDER_HEAD.MT_Is_Merchant_Trade =1)  "
        ElseIf clsCommon.CompairString(IsMerchantTrade, "SRN") = CompairStringResult.Equal Then
            whrCls += " and TSPL_SRN_HEAD.Against_PO in ( Select TSPL_SRN_HEAD.Against_PO  from TSPL_SRN_HEAD left Outer Join TSPL_PURCHASE_ORDER_HEAD on TSPL_SRN_HEAD.Against_PO =TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No where TSPL_PURCHASE_ORDER_HEAD.MT_Is_Merchant_Trade =0) "
        End If
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_SRN_HEAD.SRN_No = (select MIN(SRN_No) from TSPL_SRN_HEAD WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Last
                qry += " and TSPL_SRN_HEAD.SRN_No = (select Max(SRN_No) from TSPL_SRN_HEAD WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Current
                qry += " and TSPL_SRN_HEAD.SRN_No = '" + strPONo + "'"
            Case NavigatorType.Next
                qry += " and TSPL_SRN_HEAD.SRN_No = (select Min(SRN_No) from TSPL_SRN_HEAD where SRN_No>'" + strPONo + "' " + whrCls + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_SRN_HEAD.SRN_No = (select Max(SRN_No) from TSPL_SRN_HEAD where SRN_No<'" + strPONo + "' " + whrCls + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            'obj = New clsSRNHead()
            obj.isExemptSecurityDedution = CInt(dt.Rows(0)("isExemptSecurityDedution"))
            obj.isJobWorkOutward = clsCommon.myCstr(dt.Rows(0)("isJobWorkOutward"))
            obj.against_roadpermit = clsCommon.myCstr(dt.Rows(0)("Issue_Road_Permit"))
            obj.agnst_cform = clsCommon.myCstr(dt.Rows(0)("Issue_C_Form"))
            obj.autosrnfromrgp = clsCommon.myCstr(dt.Rows(0)("autosrn"))
            'obj.isreadytopost = clsCommon.myCstr(dt.Rows(0)("IsReadyToPost"))
            obj.Document_Type = clsCommon.myCstr(dt.Rows(0)("Document_Type"))
            obj.Against_QC_Code = clsCommon.myCstr(dt.Rows(0)("Against_QC_Code"))
            obj.Against_QC_Date = clsCommon.myCstr(dt.Rows(0)("Against_QC_Date"))
            '-------------------------------------------------------------------------
            obj.RGP_Type = clsCommon.myCstr(dt.Rows(0)("RGP_Type"))
            obj.SRN_No = clsCommon.myCstr(dt.Rows(0)("SRN_No"))
            obj.SRN_Date = clsCommon.myCDate(dt.Rows(0)("SRN_Date"))
            obj.Vendor_Code = clsCommon.myCstr(dt.Rows(0)("Vendor_Code"))
            obj.Vendor_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1 AndAlso clsCommon.myCdbl(dt.Rows(0)("iscancel")) <> 1, ERPTransactionStatus.Approved, IIf(clsCommon.myCdbl(dt.Rows(0)("iscancel")) = 1, ERPTransactionStatus.Cancel, ERPTransactionStatus.Pending))
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Is_Internal = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Internal")) = 1, True, False)
            obj.is_RGP_Non_Inventory = IIf(clsCommon.myCdbl(dt.Rows(0)("is_RGP_Non_Inventory")) = 1, True, False)
            obj.Confirmatory_PO = (clsCommon.myCdbl(dt.Rows(0)("Confirmatory_PO")) = 1)
            obj.is_QCAccepted = IIf(clsCommon.myCdbl(dt.Rows(0)("is_QCAccepted")) = 1, True, False)
            obj.Ref_No = clsCommon.myCstr(dt.Rows(0)("Ref_No"))
            obj.TenderNo = clsCommon.myCstr(dt.Rows(0)("TenderNo"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Bill_To_Location = clsCommon.myCstr(dt.Rows(0)("Bill_To_Location"))
            obj.Ship_To_Location = clsCommon.myCstr(dt.Rows(0)("Ship_To_Location"))
            obj.Sublocation_Code = clsCommon.myCstr(dt.Rows(0)("Sublocation_Code"))
            obj.Against_CF_Sale_Purchase_No = clsCommon.myCstr(dt.Rows(0)("Against_CF_Sale_Purchase_No"))
            obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
            obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
            obj.TAX1_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
            obj.TAX1_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Base_Amt"))
            obj.TAX1_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            obj.TAX2 = clsCommon.myCstr(dt.Rows(0)("TAX2"))
            obj.TAX2_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
            obj.TAX2_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Base_Amt"))
            obj.TAX2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            obj.TAX3 = clsCommon.myCstr(dt.Rows(0)("TAX3"))
            obj.TAX3_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Base_Amt"))
            obj.TAX3_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
            obj.TAX3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            obj.TAX4 = clsCommon.myCstr(dt.Rows(0)("TAX4"))
            obj.TAX4_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
            obj.TAX4_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Base_Amt"))
            obj.TAX4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            obj.TAX5 = clsCommon.myCstr(dt.Rows(0)("TAX5"))
            obj.TAX5_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
            obj.TAX5_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Base_Amt"))
            obj.TAX5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
            obj.TAX6 = clsCommon.myCstr(dt.Rows(0)("TAX6"))
            obj.TAX6_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
            obj.TAX6_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Base_Amt"))
            obj.TAX6_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
            obj.TAX7 = clsCommon.myCstr(dt.Rows(0)("TAX7"))
            obj.TAX7_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX7_Rate"))
            obj.TAX7_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Base_Amt"))
            obj.TAX7_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
            obj.TAX8 = clsCommon.myCstr(dt.Rows(0)("TAX8"))
            obj.TAX8_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX8_Rate"))
            obj.TAX8_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Base_Amt"))
            obj.TAX8_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
            obj.TAX9 = clsCommon.myCstr(dt.Rows(0)("TAX9"))
            obj.TAX9_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX9_Rate"))
            obj.TAX9_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Base_Amt"))
            obj.TAX9_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))
            obj.TAX10 = clsCommon.myCstr(dt.Rows(0)("TAX10"))
            obj.TAX10_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX10_Rate"))
            obj.TAX10_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Base_Amt"))
            obj.TAX10_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
            obj.Total_Tax_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
            obj.Discount_Base = clsCommon.myCdbl(dt.Rows(0)("Discount_Base"))
            obj.Discount_Amt = clsCommon.myCdbl(dt.Rows(0)("Discount_Amt"))
            obj.Amount_Less_Discount = clsCommon.myCdbl(dt.Rows(0)("Amount_Less_Discount"))
            obj.Total_Taxable_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Taxable_Amount"))
            obj.SRN_Total_Amt = clsCommon.myCdbl(dt.Rows(0)("SRN_Total_Amt"))
            obj.Comments = clsCommon.myCstr(dt.Rows(0)("Comments"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            obj.BillToLocationName = clsCommon.myCstr(dt.Rows(0)("BillToLocationName"))
            obj.ShipToLocationName = clsCommon.myCstr(dt.Rows(0)("ShipToLocationName"))
            obj.SubLocationName = clsCommon.myCstr(dt.Rows(0)("SubLocationName"))
            obj.TaxGroupName = clsCommon.myCstr(dt.Rows(0)("TaxGroupName"))
            obj.TermsName = clsCommon.myCstr(dt.Rows(0)("TermsName"))
            obj.Retention = clsCommon.myCdbl(dt.Rows(0)("Retention"))
            If clsCommon.myLen(dt.Rows(0)("Posting_Date")) > 0 Then
                obj.Posting_Date = clsCommon.myCstr(dt.Rows(0)("Posting_Date"))
            End If
            If dt.Rows(0)("GRN_Date") IsNot DBNull.Value Then
                obj.GRN_Date = clsCommon.myCDate(dt.Rows(0)("GRN_Date"))
            End If
            obj.TotalUnit_Cost_Tax = clsCommon.myCdbl(dt.Rows(0)("TotalUnit_Cost_Tax"))
            obj.Carrier = clsCommon.myCstr(dt.Rows(0)("Carrier"))
            obj.VehicleNo = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
            obj.GRNo = clsCommon.myCstr(dt.Rows(0)("GRNo"))
            obj.GENo = clsCommon.myCstr(dt.Rows(0)("GENo"))
            If dt.Rows(0)("GEDate") IsNot DBNull.Value Then
                obj.GEDate = clsCommon.myCDate(dt.Rows(0)("GEDate"))
            End If

            'stuti
            obj.RoadPermit_No = clsCommon.myCstr(dt.Rows(0)("RoadPermit_No"))
            If dt.Rows(0)("RoadPermit_Date") IsNot DBNull.Value Then
                obj.RoadPermit_Date = clsCommon.myCDate(dt.Rows(0)("RoadPermit_Date"))
            End If
            '====end here===

            obj.RMDA_No = clsCommon.myCstr(dt.Rows(0)("RMDA_No"))
            obj.RMDA_Date = clsCommon.myCstr(dt.Rows(0)("RMDA_Date"))
            ' ''obj.Landed_Add_Cost = clsCommon.myCdbl(dt.Rows(0)("Landed_Add_Cost"))
            obj.Total_Landed_Cost = clsCommon.myCdbl(dt.Rows(0)("TotalUnit_Cost_Tax"))

            obj.Dept = clsCommon.myCstr(dt.Rows(0)("Dept"))
            obj.Dept_Desc = clsCommon.myCstr(dt.Rows(0)("Dept_Desc"))
            obj.Item_Type = clsCommon.myCstr(dt.Rows(0)("Item_Type"))
            obj.PROJECT_ID = clsCommon.myCstr(dt.Rows(0)("PROJECT_ID"))

            obj.Against_Requisition = clsCommon.myCstr(dt.Rows(0)("Against_Requisition"))
            obj.Against_PO = clsCommon.myCstr(dt.Rows(0)("Against_PO"))
            obj.Against_GRN = clsCommon.myCstr(dt.Rows(0)("Against_GRN"))
            obj.Against_MRN = clsCommon.myCstr(dt.Rows(0)("Against_MRN"))
            obj.Against_RGP = clsCommon.myCstr(dt.Rows(0)("Against_RGP"))
            obj.PurchaseOrder_Type = clsCommon.myCstr(dt.Rows(0)("PurchaseOrder_Type"))
            obj.Against_Schedule_Code = clsCommon.myCstr(dt.Rows(0)("Against_Schedule_Code"))

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

            obj.Total_Add_Charge = clsCommon.myCdbl(dt.Rows(0)("Total_Add_Charge"))
            obj.Inv_No = clsCommon.myCstr(dt.Rows(0)("Inv_No"))
            If clsCommon.myLen((dt.Rows(0)("Challan_Date"))) <= 0 Then
                obj.Challan_Date = ""
            Else
                obj.Challan_Date = clsCommon.GetPrintDate((dt.Rows(0)("Challan_Date")), "dd/MMM/yyyy")
            End If

            If clsCommon.myLen((dt.Rows(0)("Inv_Date"))) <= 0 Then
                obj.Inv_Date = ""
            Else
                obj.Inv_Date = clsCommon.GetPrintDate((dt.Rows(0)("Inv_Date")), "dd/MMM/yyyy")
            End If

            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)
            obj.Form_38 = clsCommon.myCstr(dt.Rows(0)("Form_38"))
            obj.is_Excise_On_Qty = IIf(clsCommon.myCdbl(dt.Rows(0)("is_Excise_On_Qty")) = 1, True, False)
            obj.AssessableAmt = clsCommon.myCdbl(dt.Rows(0)("AssessableAmt"))

            '' CURRENCYCONVERSION 
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            '' END CURRENCYCONVERSION 
            obj.IsAbatementPO = dt.Rows(0)("IsAbatementPO")

            obj.Total_Accepted_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Accepted_Amount"))
            obj.Total_Rejected_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Rejected_Amount"))
            obj.Total_Shortage_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Shortage_Amount"))
            obj.Total_Leak_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Leak_Amount"))
            obj.Total_Burst_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Burst_Amount"))
            obj.Is_Shortage_Include_In_Landed_Cost = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Shortage_Include_In_Landed_Cost")) = 1, True, False)
            obj.Total_Add_Charge_Insurance = clsCommon.myCdbl(dt.Rows(0)("Total_Add_Charge_Insurance"))
            obj.Total_Item_Insurance_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Item_Insurance_Amt"))
            obj.NIRQC = clsCommon.myCdbl(dt.Rows(0)("NIR_QC"))
            obj.QC_Status = clsCommon.myCstr(dt.Rows(0)("QC_Status"))
            qry = "SELECT TSPL_SRN_DETAIL.*,TSPL_LOCATION_MASTER.Location_Desc as LocationName,TSPL_QC_CHECK_HEAD.QC_Status,Case When IsNull(TSPL_SRN_DETAIL.Rejected_Qty,0)>0 Then IsNull(TSPL_SRN_DETAIL.Rejected_Qty,0) Else IsNull(TSPL_QC_CHECK_DETAIL.Reject_Qty,0) End As RjtQty
                   FROM TSPL_SRN_DETAIL left Outer Join TSPL_SRN_HEAD On TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No Left Outer Join TSPL_QC_CHECK_HEAD On TSPL_QC_CHECK_HEAD.Gate_Entry_No=TSPL_SRN_HEAD.Against_GRN left outer Join TSPL_QC_CHECK_DETAIL On TSPL_QC_CHECK_DETAIL.Document_Code=TSPL_QC_CHECK_HEAD.Document_Code And TSPL_QC_CHECK_DETAIL.Item_Code=TSPL_SRN_DETAIL.Item_Code left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SRN_DETAIL.Location where TSPL_SRN_DETAIL.SRN_No='" + obj.SRN_No + "' ORDER BY TSPL_SRN_DETAIL.Line_No"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsSRNDetail)
                Dim objTr As clsSRNDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsSRNDetail
                    objTr.SRN_No = clsCommon.myCstr(dr("SRN_No"))
                    objTr.Row_Type = clsCommon.myCstr(dr("Row_Type"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Status = Convert.ToInt32(clsCommon.myCdbl(dr("Status")))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.UOMWeight = clsCommon.myCstr(dr("UOM_WEIGHT"))
                    objTr.UOMWeightValue = clsCommon.myCdbl(dr("UOM_WEIGHT_VALUE"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.SRN_Qty = clsCommon.myCdbl(dr("SRN_Qty"))
                    objTr.Leak_Qty = clsCommon.myCdbl(dr("Leak_Qty"))
                    objTr.Burst_Qty = clsCommon.myCdbl(dr("Burst_Qty"))
                    objTr.Short_Qty = clsCommon.myCdbl(dr("Short_Qty"))
                    objTr.Rejected_Qty = clsCommon.myCdbl(dr("RjtQty"))
                    objTr.Freeqty = clsCommon.myCdbl(dr("Free_Qty"))
                    objTr.MRN_Id = clsCommon.myCstr(dr("MRN_Id"))
                    objTr.GRN_ID = clsCommon.myCstr(dr("GRN_ID"))
                    objTr.RGP_Id = clsCommon.myCstr(dr("RGP_Id"))
                    objTr.Req_No = clsCommon.myCstr(dr("Req_No"))
                    objTr.Balance_Qty = clsCommon.myCdbl(dr("Balance_Qty"))
                    objTr.Unit_code = clsCommon.myCstr(dr("Unit_code"))
                    objTr.Location = clsCommon.myCstr(dr("Location"))
                    objTr.LocationName = clsCommon.myCstr(dr("LocationName"))
                    objTr.Item_Cost = clsCommon.myCdbl(dr("Item_Cost"))
                    objTr.TAX1 = clsCommon.myCstr(dr("TAX1"))
                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(dr("TAX1_Base_Amt"))
                    objTr.TAX1_Rate = clsCommon.myCdbl(dr("TAX1_Rate"))
                    objTr.TAX1_Amt = clsCommon.myCdbl(dr("TAX1_Amt"))
                    objTr.TAX2 = clsCommon.myCstr(dr("TAX2"))
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(dr("TAX2_Base_Amt"))
                    objTr.TAX2_Rate = clsCommon.myCdbl(dr("TAX2_Rate"))
                    objTr.TAX2_Amt = clsCommon.myCdbl(dr("TAX2_Amt"))
                    objTr.TAX3 = clsCommon.myCstr(dr("TAX3"))
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(dr("TAX3_Base_Amt"))
                    objTr.TAX3_Rate = clsCommon.myCdbl(dr("TAX3_Rate"))
                    objTr.TAX3_Amt = clsCommon.myCdbl(dr("TAX3_Amt"))
                    objTr.TAX4 = clsCommon.myCstr(dr("TAX4"))
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(dr("TAX4_Base_Amt"))
                    objTr.TAX4_Rate = clsCommon.myCdbl(dr("TAX4_Rate"))
                    objTr.TAX4_Amt = clsCommon.myCdbl(dr("TAX4_Amt"))
                    objTr.TAX5 = clsCommon.myCstr(dr("TAX5"))
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(dr("TAX5_Base_Amt"))
                    objTr.TAX5_Rate = clsCommon.myCdbl(dr("TAX5_Rate"))
                    objTr.TAX5_Amt = clsCommon.myCdbl(dr("TAX5_Amt"))
                    objTr.TAX6 = clsCommon.myCstr(dr("TAX6"))
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(dr("TAX6_Base_Amt"))
                    objTr.TAX6_Rate = clsCommon.myCdbl(dr("TAX6_Rate"))
                    objTr.TAX6_Amt = clsCommon.myCdbl(dr("TAX6_Amt"))
                    objTr.TAX7 = clsCommon.myCstr(dr("TAX7"))
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(dr("TAX7_Base_Amt"))
                    objTr.TAX7_Rate = clsCommon.myCdbl(dr("TAX7_Rate"))
                    objTr.TAX7_Amt = clsCommon.myCdbl(dr("TAX7_Amt"))
                    objTr.TAX8 = clsCommon.myCstr(dr("TAX8"))
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(dr("TAX8_Base_Amt"))
                    objTr.TAX8_Rate = clsCommon.myCdbl(dr("TAX8_Rate"))
                    objTr.TAX8_Amt = clsCommon.myCdbl(dr("TAX8_Amt"))
                    objTr.TAX9 = clsCommon.myCstr(dr("TAX9"))
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(dr("TAX9_Base_Amt"))
                    objTr.TAX9_Rate = clsCommon.myCdbl(dr("TAX9_Rate"))
                    objTr.TAX9_Amt = clsCommon.myCdbl(dr("TAX9_Amt"))
                    objTr.TAX10 = clsCommon.myCstr(dr("TAX10"))
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(dr("TAX10_Base_Amt"))
                    objTr.TAX10_Rate = clsCommon.myCdbl(dr("TAX10_Rate"))
                    objTr.TAX10_Amt = clsCommon.myCdbl(dr("TAX10_Amt"))
                    objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                    objTr.Disc_Type = clsCommon.myCdbl(dr("Disc_Type"))
                    objTr.Header_Discount_Per = clsCommon.myCdbl(dr("Header_Discount_Per"))
                    objTr.Header_Discount_Amount = clsCommon.myCdbl(dr("Header_Discount_Amount"))
                    objTr.Disc_Per = clsCommon.myCdbl(dr("Disc_Per"))
                    objTr.Detail_Discount_Amount = clsCommon.myCdbl(dr("Detail_Discount_Amount"))

                    objTr.Disc_Per_Unit = clsCommon.myCdbl(dr("Disc_Per_Unit"))
                    objTr.Disc_Amt_Per_Unit = clsCommon.myCdbl(dr("Disc_Amt_Per_Unit"))

                    objTr.Disc_Amt = clsCommon.myCdbl(dr("Disc_Amt"))
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(dr("Amt_Less_Discount"))
                    objTr.Taxable_Amount_Per = clsCommon.myCdbl(dr("Taxable_Amount_Per"))
                    objTr.Taxable_Amount = clsCommon.myCdbl(dr("Taxable_Amount"))
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(dr("Total_Tax_Amt"))
                    objTr.Item_Net_Amt = clsCommon.myCdbl(dr("Item_Net_Amt"))

                    objTr.Landed_Cost_Rate = clsCommon.myCdbl(dr("Landed_Cost_Rate"))
                    objTr.Landed_Cost_Amount = clsCommon.myCdbl(dr("Landed_Cost_Amount"))
                    objTr.Is_Mannual_Amt = clsCommon.myCdbl(dr("Is_Mannual_Amt"))

                    ' ''objTr.Landed_Cost_Rate = clsCommon.myCdbl(dr("Landed_Cost_Rate"))
                    ' ''objTr.Landed_Cost_Amount = clsCommon.myCdbl(dr("Landed_Cost_Amount"))

                    objTr.MRP = clsCommon.myCdbl(dr("MRP"))
                    objTr.Assessable = clsCommon.myCdbl(dr("Assessable"))
                    objTr.AssessableAmt = clsCommon.myCdbl(dr("AssessableAmt"))
                    objTr.Batch_No = clsCommon.myCstr(dr("Batch_No"))
                    objTr.Bin_No = clsCommon.myCstr(dr("Bin_No"))
                    If dr("MFG_Date") IsNot DBNull.Value Then
                        objTr.MFG_Date = clsCommon.myCDate(dr("MFG_Date"))
                    End If
                    If dr("Expiry_Date") IsNot DBNull.Value Then
                        objTr.Expiry_Date = clsCommon.myCDate(dr("Expiry_Date"))
                    End If
                    objTr.Specification = clsCommon.myCstr(dr("Specification"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))


                    objTr.Fater_Code = clsCommon.myCstr(dr("Fater_Code"))
                    objTr.Fater_Rate = clsCommon.myCdbl(dr("Fater_Rate"))
                    objTr.Fater_Amt = clsCommon.myCdbl(dr("Fater_Amt"))

                    objTr.PO_Qty = clsCommon.myCdbl(dr("PO_Qty"))
                    objTr.GRN_Qty = clsCommon.myCdbl(dr("GRN_Qty"))
                    objTr.MRN_Qty = clsCommon.myCdbl(dr("MRN_Qty"))

                    objTr.Total_NonRecTax_PerUnit = clsCommon.myCdbl(dr("Total_NonRecTax_PerUnit"))
                    objTr.Total_RecTax_PerUnit = clsCommon.myCdbl(dr("Total_RecTax_PerUnit"))
                    objTr.Total_AddtionalCost_PerUnit = clsCommon.myCdbl(dr("Total_AddtionalCost_PerUnit"))
                    objTr.AssessableAmt = clsCommon.myCdbl(dr("AssessableAmt"))
                    ''If clsCommon.myLen(objTr.MRN_Id) > 0 Then
                    ''    dt = GetOriginalQty(objTr.MRN_Id, objTr.Item_Code, objTr.Unit_code, objTr.Assessable, objTr.MRP, trans)
                    ''    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    ''        objTr.OrgPOQty = clsCommon.myCdbl(dt.Rows(0)("PurchaseOrder_Qty"))
                    ''        objTr.OrgGRNQty = clsCommon.myCdbl(dt.Rows(0)("GRN_Qty"))
                    ''        objTr.OrgMRNQty = clsCommon.myCdbl(dt.Rows(0)("MRN_Qty"))
                    ''    End If
                    ''End If
                    objTr.PO_ID = clsCommon.myCstr(dr("PO_ID"))
                    objTr.MRN_Id = clsCommon.myCstr(dr("MRN_ID"))
                    objTr.Bar_Code = clsCommon.myCstr(dr("Bar_Code"))
                    objTr.arrSrItem = clsSerializeInvenotry.GetData("SRN", objTr.SRN_No, objTr.Item_Code, objTr.Line_No, trans)
                    '' for abatement srn
                    objTr.AbatementRate = clsCommon.myCdbl(dr("AbatementRate"))
                    objTr.AssessableMRP = clsCommon.myCdbl(dr("AssessableMRP"))
                    objTr.TotalAssessableMRP = clsCommon.myCdbl(dr("TotalAssessableMRP"))

                    objTr.Against_Schedule_Code = clsCommon.myCstr(dr("Against_Schedule_Code"))
                    objTr.RGP_Qty = clsCommon.myCdbl(dr("RGP_Qty"))
                    objTr.Schedule_Qty = clsCommon.myCdbl(dr("Schedule_Qty"))


                    objTr.Accepted_Amount = clsCommon.myCdbl(dr("Accepted_Amount"))
                    objTr.Rejected_Amount = clsCommon.myCdbl(dr("Rejected_Amount"))
                    objTr.Shortage_Amount = clsCommon.myCdbl(dr("Shortage_Amount"))
                    objTr.Leak_Amount = clsCommon.myCdbl(dr("Leak_Amount"))
                    objTr.Burst_Amount = clsCommon.myCdbl(dr("Burst_Amount"))
                    objTr.Amt_Less_Discount_Without_Shortage = clsCommon.myCdbl(dr("Amt_Less_Discount_Without_Shortage"))


                    ''------------------19/10/2016=======================
                    objTr.ItemAdd_Charge_Code1 = clsCommon.myCstr(dr("ItemAdd_Charge_Code1"))
                    objTr.ItemAdd_Charge_Code2 = clsCommon.myCstr(dr("ItemAdd_Charge_Code2"))
                    objTr.ItemAdd_Charge_Code3 = clsCommon.myCstr(dr("ItemAdd_Charge_Code3"))
                    objTr.ItemAdd_Charge_Code4 = clsCommon.myCstr(dr("ItemAdd_Charge_Code4"))
                    objTr.ItemAdd_Charge_Code5 = clsCommon.myCstr(dr("ItemAdd_Charge_Code5"))
                    objTr.ItemAdd_Charge_Code6 = clsCommon.myCstr(dr("ItemAdd_Charge_Code6"))
                    objTr.ItemAdd_Charge_Code7 = clsCommon.myCstr(dr("ItemAdd_Charge_Code7"))
                    objTr.ItemAdd_Charge_Code8 = clsCommon.myCstr(dr("ItemAdd_Charge_Code8"))
                    objTr.ItemAdd_Charge_Code9 = clsCommon.myCstr(dr("ItemAdd_Charge_Code9"))
                    objTr.ItemAdd_Charge_Code10 = clsCommon.myCstr(dr("ItemAdd_Charge_Code10"))
                    objTr.ItemAdd_Org_Charge_Amt1 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt1"))
                    objTr.ItemAdd_Org_Charge_Amt2 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt2"))
                    objTr.ItemAdd_Org_Charge_Amt3 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt3"))
                    objTr.ItemAdd_Org_Charge_Amt4 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt4"))
                    objTr.ItemAdd_Org_Charge_Amt5 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt5"))
                    objTr.ItemAdd_Org_Charge_Amt6 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt6"))
                    objTr.ItemAdd_Org_Charge_Amt7 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt7"))
                    objTr.ItemAdd_Org_Charge_Amt8 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt8"))
                    objTr.ItemAdd_Org_Charge_Amt9 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt9"))
                    objTr.ItemAdd_Org_Charge_Amt10 = clsCommon.myCdbl(dr("ItemAdd_Org_Charge_Amt10"))
                    objTr.ItemAdd_Calc_Charge_Amt1 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt1"))
                    objTr.ItemAdd_Calc_Charge_Amt2 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt2"))
                    objTr.ItemAdd_Calc_Charge_Amt3 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt3"))
                    objTr.ItemAdd_Calc_Charge_Amt4 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt4"))
                    objTr.ItemAdd_Calc_Charge_Amt5 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt5"))
                    objTr.ItemAdd_Calc_Charge_Amt6 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt6"))
                    objTr.ItemAdd_Calc_Charge_Amt7 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt7"))
                    objTr.ItemAdd_Calc_Charge_Amt8 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt8"))
                    objTr.ItemAdd_Calc_Charge_Amt9 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt9"))
                    objTr.ItemAdd_Calc_Charge_Amt10 = clsCommon.myCdbl(dr("ItemAdd_Calc_Charge_Amt10"))
                    objTr.Total_ItemAdd_Charge = clsCommon.myCdbl(dr("Total_ItemAdd_Charge"))
                    objTr.Category = clsCommon.myCstr(dr("Category"))
                    objTr.Emergency = IIf(clsCommon.myCdbl(dr("Emergency")) = 1, True, False)
                    objTr.Capex_Code = clsCommon.myCstr(dr("Capex_Code"))
                    objTr.Capex_SubCode = clsCommon.myCstr(dr("Capex_SubCode"))
                    objTr.Against_Item_Wise_Tax_Rate = clsCommon.myCstr(dr("Against_Item_Wise_Tax_Rate"))
                    objTr.Insurance_Per = clsCommon.myCdbl(dr("Insurance_Per"))
                    objTr.Insurance_Base_Amt = clsCommon.myCdbl(dr("Insurance_Base_Amt"))

                    objTr.Item_Insurance_Base_Amt = clsCommon.myCdbl(dr("Item_Insurance_Base_Amt"))
                    objTr.Item_Insurance_Apply_On = clsCommon.myCstr(dr("Item_Insurance_Apply_On"))
                    objTr.Item_Insurance_Rate = clsCommon.myCdbl(dr("Item_Insurance_Rate"))
                    objTr.Item_Insurance_Amt = clsCommon.myCdbl(dr("Item_Insurance_Amt"))
                    objTr.Item_Amt_After_Insurance = clsCommon.myCdbl(dr("Item_Amt_After_Insurance"))

                    objTr.arrBatchItem = clsBatchInventory.GetData("SRN", objTr.SRN_No, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr.Add(objTr)
                Next
            End If

            '----------------------roadpermit detail---------------------------------------------
            obj.Arr_Road = New List(Of clsSRNRoadPermitDetail)
            obj.Arr_CFORM = New List(Of clsSRNCFORMDetail)

            If obj.against_roadpermit = "1" Then
                qry = "select TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.form_code,TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.issue_no,TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.srn_no from TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL where TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.purchaseorder_NO='" + obj.Against_PO + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.vendor_code='" + obj.Vendor_Code + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.srn_no='" + obj.SRN_No + "'"
                dt = New DataTable()
                dt = clsDBFuncationality.GetDataTable(qry, trans)

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                    Dim objtr1 As New clsSRNRoadPermitDetail()
                    For Each dr As DataRow In dt.Rows
                        objtr1 = New clsSRNRoadPermitDetail()
                        objtr1.roadpono = clsCommon.myCstr(obj.Against_PO)
                        objtr1.roadvendor = clsCommon.myCstr(obj.Vendor_Code)
                        objtr1.roadcode = clsCommon.myCstr(dr("form_code"))
                        objtr1.roadissue_no = clsCommon.myCstr(dr("issue_no"))
                        objtr1.RoadpermitSRNNO = clsCommon.myCstr(dr("srn_no"))

                        obj.Arr_Road.Add(objtr1)
                    Next
                End If
            End If

            If obj.agnst_cform = "1" Then
                qry = "select TSPL_CFORM_ISSUE_RECEIVE_DETAIL.form_code,TSPL_CFORM_ISSUE_RECEIVE_DETAIL.issue_no,TSPL_CFORM_ISSUE_RECEIVE_DETAIL.srn_no from TSPL_CFORM_ISSUE_RECEIVE_DETAIL where TSPL_CFORM_ISSUE_RECEIVE_DETAIL.purchaseorder_no='" + obj.Against_PO + "' and TSPL_CFORM_ISSUE_RECEIVE_DETAIL.vendor_code='" + obj.Vendor_Code + "' and TSPL_CFORM_ISSUE_RECEIVE_DETAIL.srn_no='" + obj.SRN_No + "'"
                dt = New DataTable()
                dt = clsDBFuncationality.GetDataTable(qry, trans)

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                    Dim objtr1 As New clsSRNCFORMDetail()
                    For Each dr As DataRow In dt.Rows
                        objtr1 = New clsSRNCFORMDetail()
                        objtr1.cformpono = clsCommon.myCstr(obj.Against_PO)
                        objtr1.cformvendor = clsCommon.myCstr(obj.Vendor_Code)
                        objtr1.cformcode = clsCommon.myCstr(dr("form_code"))
                        objtr1.cformissue_no = clsCommon.myCstr(dr("issue_no"))
                        objtr1.cformSRNNO = clsCommon.myCstr(dr("srn_no"))

                        obj.Arr_CFORM.Add(objtr1)
                    Next
                End If
            End If
            '---------------------------------------------------------------------------------------
            obj.Arr_ACInsurance = clsSRNAdditionChargeInsurance.GetData(obj.SRN_No, trans)
        End If

        Return obj
    End Function

    Public Shared Function GetOriginalQty(ByVal strGRN As String, ByVal strMrnNo As String, ByVal strPOCode As String, ByVal strICode As String, ByVal strUOM As String, ByVal dblAssessable As Double, ByVal dblMRP As Double, ByVal trans As SqlTransaction) As DataTable
        Dim qry As String = "Select TSPL_MRN_DETAIL.accept_qty,TSPL_MRN_DETAIL.Reject_qty,TSPL_MRN_DETAIL.Short_Qty,TSPL_MRN_DETAIL.MRN_No,(TSPL_MRN_DETAIL.MRN_Qty ) as MRN_Qty,TSPL_GRN_DETAIL.GRN_No,(TSPL_GRN_DETAIL.GRN_Qty ) as GRN_Qty, TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No,case when TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty<=0 then podr.PurchaseOrder_Qty else TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty end as PurchaseOrder_Qty from TSPL_MRN_DETAIL left outer join TSPL_GRN_DETAIL on TSPL_GRN_DETAIL.GRN_No=TSPL_MRN_DETAIL.GRN_Id and isnull(tspl_grn_detail.po_id,'')=isnull(tspl_mrn_detail.po_id,'') and TSPL_GRN_DETAIL.Item_Code=TSPL_MRN_DETAIL.Item_Code and TSPL_GRN_DETAIL.Unit_code=TSPL_MRN_DETAIL.Unit_code and isnull(TSPL_GRN_DETAIL.Assessable,0)=isnull(TSPL_MRN_DETAIL.Assessable,0) and isnull(TSPL_GRN_DETAIL.Item_Code,0)=isnull(TSPL_MRN_DETAIL.Item_Code ,0) left outer join TSPL_PURCHASE_ORDER_DETAIL on TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No= TSPL_GRN_DETAIL.PO_Id and TSPL_PURCHASE_ORDER_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code and TSPL_PURCHASE_ORDER_DETAIL.Unit_code=  TSPL_GRN_DETAIL.Unit_code and isnull(TSPL_PURCHASE_ORDER_DETAIL.Assessable,0)=  isnull(TSPL_GRN_DETAIL.Assessable,0) and isnull(TSPL_PURCHASE_ORDER_DETAIL.MRP,0)=  isnull(TSPL_GRN_DETAIL.MRP,0)  left join TSPL_RGP_HEAD on TSPL_RGP_HEAD.RGP_No=TSPL_GRN_DETAIL.Against_RGP_No  left join TSPL_PURCHASE_ORDER_DETAIL podr on TSPL_RGP_HEAD.PO_ID=podr.PurchaseOrder_No and podr.Item_Code=TSPL_MRN_DETAIL.Item_Code where TSPL_MRN_DETAIL.MRN_No='" + strMrnNo + "' and TSPL_MRN_DETAIL.Item_Code='" + strICode + "' and TSPL_MRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_MRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_MRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' and isnull(TSPL_MRN_DETAIL.PO_ID,'')='" + strPOCode + "'"
        If clsCommon.myLen(strGRN) > 0 Then
            qry += " and TSPL_GRN_DETAIL.GRN_No in ('" + strGRN + "') "
        End If

        qry += " Union " _
        & " Select TSPL_MRN_DETAIL.accept_qty,TSPL_MRN_DETAIL.Reject_qty,TSPL_MRN_DETAIL.Short_Qty,TSPL_MRN_DETAIL.MRN_No,(TSPL_MRN_DETAIL.MRN_Qty ) as MRN_Qty,TSPL_GRN_DETAIL.GRN_No,(TSPL_GRN_DETAIL.GRN_Qty ) as GRN_Qty, TSPL_RGP_DETAIL.RGP_No,case when coalesce(TSPL_RGP_DETAIL.RGP_Qty,0)<=0 then podr.PurchaseOrder_Qty else TSPL_RGP_DETAIL.RGP_Qty end as PurchaseOrder_Qty from TSPL_MRN_DETAIL left outer join TSPL_GRN_DETAIL on TSPL_GRN_DETAIL.GRN_No=TSPL_MRN_DETAIL.GRN_Id and isnull(tspl_grn_detail.Against_RGP_No,'')=isnull(tspl_mrn_detail.RGP_NO,'') and TSPL_GRN_DETAIL.Item_Code=TSPL_MRN_DETAIL.Item_Code and TSPL_GRN_DETAIL.Unit_code=TSPL_MRN_DETAIL.Unit_code and isnull(TSPL_GRN_DETAIL.Assessable,0)=isnull(TSPL_MRN_DETAIL.Assessable,0) and isnull(TSPL_GRN_DETAIL.Item_Code,0)=isnull(TSPL_MRN_DETAIL.Item_Code ,0) left outer join TSPL_RGP_DETAIL on TSPL_RGP_DETAIL.RGP_No= TSPL_GRN_DETAIL.Against_RGP_No and TSPL_RGP_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code and TSPL_RGP_DETAIL.Unit_code=  TSPL_GRN_DETAIL.Unit_code and isnull(TSPL_RGP_DETAIL.MRP,0)=  isnull(TSPL_GRN_DETAIL.MRP,0)  left join TSPL_RGP_HEAD on TSPL_RGP_HEAD.RGP_No=TSPL_GRN_DETAIL.Against_RGP_No  left join TSPL_PURCHASE_ORDER_DETAIL podr on TSPL_RGP_HEAD.PO_ID=podr.PurchaseOrder_No and podr.Item_Code=TSPL_MRN_DETAIL.Item_Code where TSPL_MRN_DETAIL.MRN_No='" + strMrnNo + "' and TSPL_MRN_DETAIL.Item_Code='" + strICode + "' and TSPL_MRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_MRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "'  and isnull(TSPL_MRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' and isnull(TSPL_MRN_DETAIL.RGP_NO,'')='" + strPOCode + "'"
        If clsCommon.myLen(strGRN) > 0 Then
            qry += " and TSPL_GRN_DETAIL.GRN_No in ('" + strGRN + "') "
        End If
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function
    Public Shared Function CreateAutoJobWorkTransferOut(ByVal trans As SqlTransaction, ByVal objSRN As clsSRNHead) As Boolean
        Dim obj = New clsJWOTransferOtherHead()
        obj.TRANSFER_DATE = objSRN.SRN_Date
        obj.From_Locaction = objSRN.Bill_To_Location
        obj.To_Locaction = objSRN.Sublocation_Code
        obj.Vendor_Code = objSRN.Vendor_Code
        obj.Remarks = "Against SRN NO  - " + objSRN.SRN_No + "'   ' " + objSRN.Remarks
        obj.Vehicle_Code = objSRN.Vendor_Code
        obj.Vehicle_No = objSRN.VehicleNo
        obj.AgainstSRN_No = objSRN.SRN_No
        obj.Arr = New List(Of clsJWOTransferOtherDetail)
        Dim rowNo As Integer = 0
        For Each objSRNDetail As clsSRNDetail In objSRN.Arr
            Dim objTr As New clsJWOTransferOtherDetail()
            If clsCommon.CompairString(objSRNDetail.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                rowNo += 1
                objTr.line_No = rowNo
                objTr.Item_Code = objSRNDetail.Item_Code
                objTr.UOM = objSRNDetail.Unit_code
                objTr.Qty = objSRNDetail.SRN_Qty
                objTr.Rate = objSRNDetail.Item_Cost
                objTr.Amount = objSRNDetail.Amount
                If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                    obj.Arr.Add(objTr)
                End If
            End If
        Next
        clsJWOTransferOtherHead.SaveData(obj, True, trans)
        clsJWOTransferOtherHead.PostData(obj.TRANSFER_NO, trans)
        Return True
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, Optional ByVal IsRejectQty As Boolean = False) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            PostData(FormId, strDocNo, trans, True, IsRejectQty)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal trans As SqlTransaction, ByVal IsapprovalRequired As Boolean, ByVal WithoutRejection As Boolean, Optional ByVal strVoucherNoRecreateOnly As String = Nothing) As Boolean
        ' Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim isSaved As Boolean = True
            Dim OpenPOforRejectShortageQty As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.OpenPOforRejectShortageQty, clsFixedParameterCode.OpenPOforRejectShortageQty, trans)) = "1", True, False))
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("SRN No not found to Post")
            End If
            Dim obj As clsSRNHead = clsSRNHead.GetData(strDocNo, NavigatorType.Current, trans)
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase Order", "Store Receipt Note", IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location), obj.SRN_Date, trans)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.SRN_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePurchase, clsUserMgtCode.mbtnSRN, obj.Ship_To_Location, obj.SRN_Date, trans)

            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase", "Store receipt Note", obj.Bill_To_Location, obj.SRN_Date, trans)
            If (obj.Status = 1) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            If (obj.On_Hold) Then
                Throw New Exception("SRN No " + obj.SRN_No + " Is On Hold.Can't Post it")
            End If

            obj.is_Srn_rejQty_goes_in_Rejstore = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select SRN_Rejected_Store from TSPL_PURCHASE_SETTINGS", trans)) = 0, False, True)
            ''richa agarwal
            If IsapprovalRequired = True Then
                Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_SRN_HEAD", "SRN_No", obj.SRN_No, trans)
                If isResult = False Then
                    trans.Commit()
                    Return False
                End If
            End If
            ''-----------------------
            Dim dtPostDate As DateTime = clsCommon.GETSERVERDATE(trans)
            Dim qry As String = ""

            If objCommonVar.AutoGenrateBatchInventory Then
                For Each objtr As clsSRNDetail In obj.Arr
                    If clsItemMaster.IsBatchItem(objtr.Item_Code, trans) Then
                        Dim ArrBatchItem As New List(Of clsBatchInventory)
                        Dim objBatch As New clsBatchInventory
                        objBatch.Parent_Line_No = objtr.Line_No
                        objBatch.Line_No = 1
                        objBatch.Batch_No = clsERPFuncationality.GetNextCode(trans, clsCommon.GetPrintDate(obj.SRN_Date, "dd/MMM/yyyy"), clsDocType.ItemBatch, clsItemMaster.GetItemType(objtr.Item_Code, trans), obj.Bill_To_Location, False, True, False, False, False, False, "", clsCommon.GetPrintDate(dtPostDate, "ddMMyyHHmm"), "")
                        objBatch.Manual_BatchNo = objBatch.Batch_No
                        objBatch.Manufacture_Date = dtPostDate
                        objBatch.Expiry_Date = dtPostDate.AddDays(clsItemMaster.GetSelfLife(objtr.Item_Code, trans))
                        objBatch.Qty = objtr.SRN_Qty
                        ArrBatchItem.Add(objBatch)
                        clsBatchInventory.SaveData("SRN", strDocNo, obj.SRN_Date, "I", objtr.Item_Code, obj.Bill_To_Location, objtr.Line_No, 0, objtr.Unit_code, ArrBatchItem, trans)
                    End If
                Next
                obj = clsSRNHead.GetData(strDocNo, NavigatorType.Current, trans)
            End If


            Dim IsRejectedItemFound As Boolean = False
            '
            'Dim isSerialiedItem As Integer = clsDBFuncationality.getSingleValue("select  is_serial_item from tspl_item_master where item_code='" & strFirstItemCodeNonItemRowType & "'", trans)


            If obj.Confirmatory_PO Then
                CreateConfirmatoryPO(trans, obj)
            End If

            HitInventroyMovement(obj, trans, WithoutRejection, True)


            For Each objTr As clsSRNDetail In obj.Arr
                If objTr.Rejected_Qty > 0 Then
                    IsRejectedItemFound = True

                End If
                If OpenPOforRejectShortageQty Then
                    If objTr.Rejected_Qty > 0 OrElse objTr.Short_Qty > 0 OrElse objTr.Leak_Qty > 0 OrElse objTr.Burst_Qty > 0 Then
                        Dim qr As String = Nothing
                        qr = "Update TSPL_PURCHASE_ORDER_HEAD SET close_yn='N' where TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No='" + objTr.PO_ID + "'"
                        clsDBFuncationality.ExecuteNonQuery(qr, trans)
                    End If
                End If
            Next
            Dim strRMDANo As String = ""
            If IsRejectedItemFound Then 'Comment , Asked by Ranjana mam-  obj.is_Srn_rejQty_goes_in_Rejstore AndAlso
                strRMDANo = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.RMDA, "", obj.Bill_To_Location)
                If clsCommon.myLen(strRMDANo) < 0 Then
                    Throw New Exception("Error in Code Generation for Rejected Material Dispatch Advice")
                End If
            End If
            '' Anubhooti 12-Nov-2014/10-Dec-2014 BM00000003662 (Job Work GL Entries)
            Dim PurSetJobWork As String = ""
            Dim IsJobWork As String = ""
            IsJobWork = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ISNULL(Against_JobWork ,'') AS Against_JobWork From TSPL_RGP_HEAD Where RGP_No ='" & obj.Against_RGP & "'", trans))
            PurSetJobWork = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ISNULL(Job_Work_Account,'') AS Job_Work_Account From TSPL_PURCHASE_SETTINGS", trans))
            If clsCommon.myLen(obj.Against_RGP) > 0 AndAlso clsCommon.CompairString(IsJobWork, "1") = CompairStringResult.Equal Then
                '============================if job-work type transaction and GRN-MRN is not made then for against bom vendor stock fill-------
                If clsCommon.CompairString(obj.RGP_Type, "AB") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Against_MRN) <= 0 AndAlso clsCommon.myLen(obj.Against_GRN) <= 0 Then
                    SaveRGPBOMDetail(obj, trans)
                End If
                '=========end here================================================================================================================
            End If

            CreateJournalEntry(obj, strVoucherNoRecreateOnly, trans, WithoutRejection)

            If obj.isJobWorkOutward = 1 Then
                CreateAutoJobWorkTransferOut(trans, obj)
            End If

            qry = "Update TSPL_SRN_HEAD set Status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.SRN_Date, "dd/MMM/yyyy hh:mm tt") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
            If IsRejectedItemFound Then
                qry += ",RMDA_No='" + strRMDANo + "',RMDA_Date='" + clsCommon.GetPrintDate(obj.SRN_Date, "dd/MM/yyyy") + "'"
            End If
            qry += " where SRN_No='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.SRN_No, "TSPL_SRN_HEAD", "SRN_No", trans)

            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strDocNo), "TSPL_SRN_HEAD", "SRN_No", trans)
            If objCommonVar.InternalSMSEmailinPurchaseModule = True Then
                CreateInternalEmailSMS(obj, trans)
            End If

        Catch ex As Exception

            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function HitInventroyMovement(obj As clsSRNHead, trans As SqlTransaction, WithoutRejection As Boolean, updateMrnBalance As Boolean) As Boolean
        Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
        Dim intCounter As Integer = 0
        Dim strFirstItemCodeNonItemRowType As String = GetFirstItemCode(obj.Arr)
        If Not (obj.is_RGP_Non_Inventory AndAlso clsCommon.myLen(obj.Against_RGP) > 0) Then
            For Each objTr As clsSRNDetail In obj.Arr
                intCounter = intCounter + 1
                If clsCommon.CompairString(objTr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                    If clsCommon.myLen(objTr.MRN_Id) > 0 AndAlso updateMrnBalance Then
                        Dim qry1 As String = "update TSPL_MRN_DETAIL set Balance_Qty=Balance_Qty - " + clsCommon.myCstr(clsCommon.myCdbl(objTr.SRN_Qty + objTr.Rejected_Qty + objTr.Leak_Qty + objTr.Burst_Qty + objTr.Short_Qty)) + " where MRN_No='" + objTr.MRN_Id + "' and Item_Code='" + objTr.Item_Code + "' and Unit_code='" + objTr.Unit_code + "' and isnull(MRP,0)='" + clsCommon.myCstr(objTr.MRP) + "' and isnull(Assessable,0)='" + clsCommon.myCstr(objTr.Assessable) + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry1, trans)
                    End If
                Else
                    objTr.Item_Code = strFirstItemCodeNonItemRowType
                End If
                Dim strItemType As String = clsItemMaster.GetItemType(objTr.Item_Code, trans)
                If clsCommon.CompairString(objTr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                    Dim strItemTypeToSave As String = ""
                    If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                        strItemTypeToSave = "RM"
                    ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                        strItemTypeToSave = "OT"
                    ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                        strItemTypeToSave = "FT"

                    ElseIf clsCommon.CompairString(strItemType, "A") = CompairStringResult.Equal Then
                        strItemTypeToSave = "A"
                    Else
                        strItemTypeToSave = strItemType
                    End If



                    Dim objInventoryMovemnt As New clsInventoryMovement()
                    objInventoryMovemnt.InOut = "I"
                    If clsCommon.myLen(obj.Sublocation_Code) > 0 Then
                        objInventoryMovemnt.Location_Code = obj.Sublocation_Code
                    ElseIf clsCommon.myLen(obj.Ship_To_Location) > 0 Then
                        objInventoryMovemnt.Location_Code = obj.Ship_To_Location
                    Else
                        objInventoryMovemnt.Location_Code = obj.Bill_To_Location
                    End If

                    ''richa agarwal 4 Dec,2020
                    'If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(obj.Bill_To_Location) & "'", trans)), "Y") = CompairStringResult.Equal Then
                    '    If obj.isJobWorkOutward = 0 Then
                    '        objInventoryMovemnt.Location_Code = obj.Sublocation_Code
                    '    End If
                    'End If

                    objInventoryMovemnt.Vendor_Code = obj.Vendor_Code
                    objInventoryMovemnt.Vendor_Name = obj.Vendor_Name

                    objInventoryMovemnt.Item_Code = objTr.Item_Code
                    objInventoryMovemnt.Item_Desc = objTr.Item_Desc
                    objInventoryMovemnt.Qty = objTr.SRN_Qty + objTr.Freeqty '- IIf(obj.is_Srn_rejQty_goes_in_Rejstore = True And isSerialiedItem <> 1, objTr.Rejected_Qty, 0)
                    objInventoryMovemnt.UOM = objTr.Unit_code
                    objInventoryMovemnt.MRP = objTr.MRP
                    objInventoryMovemnt.Add_Cost = objTr.Total_Tax_Amt * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                    objInventoryMovemnt.Net_Cost = objTr.Landed_Cost_Amount * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                    If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                        objInventoryMovemnt.ItemType = "RM"
                    ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                        objInventoryMovemnt.ItemType = "OT"
                    ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                        objInventoryMovemnt.ItemType = "FT"
                    End If
                    objInventoryMovemnt.ItemType = strItemTypeToSave

                    objInventoryMovemnt.Basic_Cost = objTr.Landed_Cost_Rate * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                    objInventoryMovemnt.Batch_No = objTr.Batch_No
                    objInventoryMovemnt.MFG_Date = objTr.MFG_Date
                    objInventoryMovemnt.Expiry_Date = objTr.Expiry_Date

                    ArrInventoryMovement.Add(objInventoryMovemnt)
                    If WithoutRejection = False Then
                        If objTr.Rejected_Qty > 0 AndAlso Not obj.is_Srn_rejQty_goes_in_Rejstore Then
                            Throw New Exception("Please apply rejection setting on SRN")
                        End If
                    End If

                    If objTr.Rejected_Qty > 0 AndAlso obj.is_Srn_rejQty_goes_in_Rejstore Then
                        'Comment by balwider on 02/06/2021 becuase if location is not defined it return 1 and Condition will always false.
                        'Dim RejLoc As Integer = clsDBFuncationality.getSingleValue("select count(Rejected_Location) from TSPL_LOCATION_MASTER where Location_Code='" & objTr.Location & "'", trans)
                        'If RejLoc <= 0 Then
                        '    Throw New Exception("Rejected Location Not filled of [" + objTr.Location + "]")
                        'End If

                        Dim objInventoryMovemntForRej As New clsInventoryMovement()
                        objInventoryMovemntForRej.InOut = "I"
                        objInventoryMovemntForRej.Location_Code = clsDBFuncationality.getSingleValue("select Rejected_Location from TSPL_LOCATION_MASTER where Location_Code='" & objTr.Location & "'", trans)
                        If clsCommon.myLen(objInventoryMovemntForRej.Location_Code) <= 0 Then
                            Throw New Exception("Rejected Location Not filled of [" + objTr.Location + "]")
                        End If
                        ''richa agarwal 4 Dec,2020
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(objTr.Location) & "'", trans)), "Y") = CompairStringResult.Equal Then
                            If obj.isJobWorkOutward = 0 Then
                                objInventoryMovemnt.Location_Code = obj.Sublocation_Code
                            End If
                        End If

                        objInventoryMovemnt.Vendor_Code = obj.Vendor_Code
                        objInventoryMovemnt.Vendor_Name = obj.Vendor_Name

                        objInventoryMovemntForRej.Item_Code = objTr.Item_Code
                        objInventoryMovemntForRej.Item_Desc = objTr.Item_Desc
                        objInventoryMovemntForRej.Qty = objTr.Rejected_Qty
                        objInventoryMovemntForRej.UOM = objTr.Unit_code
                        objInventoryMovemntForRej.MRP = objTr.MRP
                        ''BM00000007656 fo handle multicurrancy .Balwinder on 26.08.2015
                        objInventoryMovemntForRej.Add_Cost = objTr.Total_Tax_Amt * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                        objInventoryMovemntForRej.Net_Cost = objTr.Landed_Cost_Amount * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                        If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                            objInventoryMovemntForRej.ItemType = "RM"
                        ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                            objInventoryMovemntForRej.ItemType = "OT"
                        ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                            objInventoryMovemntForRej.ItemType = "FT"
                        End If
                        objInventoryMovemntForRej.ItemType = strItemTypeToSave

                        objInventoryMovemntForRej.Basic_Cost = objTr.Landed_Cost_Rate * IIf(obj.ConvRate = 0, 1, obj.ConvRate)
                        objInventoryMovemntForRej.Batch_No = objTr.Batch_No
                        objInventoryMovemntForRej.MFG_Date = objTr.MFG_Date
                        objInventoryMovemntForRej.Expiry_Date = objTr.Expiry_Date
                        ArrInventoryMovement.Add(objInventoryMovemntForRej)
                    End If
                End If
            Next
            clsInventoryMovement.SaveData("SRN", obj.SRN_No, obj.SRN_Date, clsCommon.GetPrintDate(obj.SRN_Date, "dd/MM/yyyy"), ArrInventoryMovement, trans)
        End If
        Return True
    End Function

    Private Shared Sub CreateInternalEmailSMS(ByVal obj As clsSRNHead, ByVal trans As SqlTransaction)
        Dim itemName As String = ""
        Dim UOM As String = ""
        Dim qty As String = ""
        Dim ItemDetail As String = ""
        Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.mbtnSRN + "2" + "'", trans)

        Dim qry As String = "select TSPL_USER_MASTER.User_Code from TSPL_USER_MASTER " '& _
        If clsCommon.myLen(obj.Against_Requisition) > 0 Then
            qry += " left join TSPL_REQUISITION_HEAD on TSPL_REQUISITION_HEAD.Created_By=TSPL_USER_MASTER.User_Code "
            qry += " left join TSPL_SRN_HEAD on TSPL_SRN_HEAD.Against_Requisition=TSPL_REQUISITION_HEAD.Requisition_Id "
        ElseIf clsCommon.myLen(obj.Against_PO) > 0 Then
            qry += " left join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.Created_By=TSPL_USER_MASTER.User_Code "
            qry += " left join TSPL_SRN_HEAD on TSPL_SRN_HEAD.Against_PO=TSPL_PURCHASE_ORDER_HEAD.PURCHASEORDER_no "
        Else
            qry += " left join TSPL_SRN_HEAD on TSPL_SRN_HEAD.Created_By=TSPL_USER_MASTER.User_Code "
        End If

        qry += " where TSPL_SRN_HEAD.SRN_No='" + obj.SRN_No + "'"
        Dim StrUserCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        Dim arrMobileNo As New List(Of String)
        Dim arrMailID As List(Of String) = clsERPFuncationality.ReportingMailIdandPhone(StrUserCode, arrMobileNo, trans)

        If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 AndAlso ((arrMailID IsNot Nothing AndAlso arrMailID.Count > 0) Or (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0)) Then


            If (obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0) Then
                For Each objdetail As clsSRNDetail In obj.Arr

                    itemName = clsCommon.myCstr(objdetail.Item_Desc)
                    UOM = clsCommon.myCstr(objdetail.Unit_code)
                    qty = clsCommon.myCstr(objdetail.SRN_Qty)

                    ItemDetail += itemName + " " + UOM + "-" + qty + Environment.NewLine

                Next
            End If


            If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 AndAlso (arrMailID IsNot Nothing AndAlso arrMailID.Count > 0) Then
                Dim objEmailH As New clsEMailHead()
                objEmailH.arrEMail = New List(Of String)()
                objEmailH.arrEMail = arrMailID

                objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))

                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.DOC_NO, obj.SRN_No)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.SRN_Date, "dd/MMM/yyyy"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myCstr(obj.SRN_Total_Amt))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.VendorNo, clsCommon.myCstr(obj.Vendor_Code))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.VendorName, clsCommon.myCstr(obj.Vendor_Name))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)

                objEmailH.SaveData(clsUserMgtCode.mbtnSRN, objEmailH, trans)
                objEmailH = Nothing

            End If

            If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 AndAlso (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0) Then
                Dim objSMSH As New clsSMSHead()
                objSMSH.arrMobilNo = New List(Of String)()
                objSMSH.arrMobilNo = arrMobileNo

                objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))

                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_NO, obj.SRN_No)
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.SRN_Date, "dd/MMM/yyyy"))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myCstr(obj.SRN_Total_Amt))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.VendorNo, clsCommon.myCstr(obj.Vendor_Code))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.VendorName, clsCommon.myCstr(obj.Vendor_Name))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)

                objSMSH.SaveData(clsUserMgtCode.mbtnSRN, objSMSH, trans)
                objSMSH = Nothing
            End If
        End If


    End Sub


    Public Shared Function CreateConfirmatoryPO(ByVal trans As SqlTransaction, ByVal objSRN As clsSRNHead) As Boolean
        Dim objPO As New clsPurchaseOrderHead
        'objPO.PurchaseOrder_No = objSRN.PurchaseOrder_No
        objPO.PurchaseOrder_Date = objSRN.SRN_Date
        objPO.PurchaseOrder_Type = objSRN.PurchaseOrder_Type
        objPO.Vendor_Code = objSRN.Vendor_Code
        objPO.Vendor_Name = objSRN.Vendor_Name
        objPO.Status = objSRN.Status
        objPO.On_Hold = objSRN.On_Hold
        objPO.Ref_No = objSRN.Ref_No
        objPO.Description = objSRN.Description
        objPO.Remarks = objSRN.Remarks
        objPO.Tax_Group = objSRN.Tax_Group
        objPO.Bill_To_Location = objSRN.Bill_To_Location
        objPO.Ship_To_Location = objSRN.Ship_To_Location
        objPO.TAX1 = objSRN.TAX1
        objPO.TAX1_Rate = objSRN.TAX1_Rate
        objPO.TAX1_Amt = objSRN.TAX1_Amt
        objPO.TAX1_Base_Amt = objSRN.TAX1_Base_Amt
        objPO.TAX2 = objSRN.TAX2
        objPO.TAX2_Rate = objSRN.TAX2_Rate
        objPO.TAX2_Amt = objSRN.TAX2_Amt
        objPO.TAX2_Base_Amt = objSRN.TAX2_Base_Amt
        objPO.TAX3 = objSRN.TAX3
        objPO.TAX3_Rate = objSRN.TAX3_Rate
        objPO.TAX3_Amt = objSRN.TAX3_Amt
        objPO.TAX3_Base_Amt = objSRN.TAX3_Base_Amt
        objPO.TAX4 = objSRN.TAX4
        objPO.TAX4_Rate = objSRN.TAX4_Rate
        objPO.TAX4_Amt = objSRN.TAX4_Amt
        objPO.TAX4_Base_Amt = objSRN.TAX4_Base_Amt
        objPO.TAX5 = objSRN.TAX5
        objPO.TAX5_Rate = objSRN.TAX5_Rate
        objPO.TAX5_Amt = objSRN.TAX5_Amt
        objPO.TAX5_Base_Amt = objSRN.TAX5_Base_Amt
        objPO.TAX6 = objSRN.TAX6
        objPO.TAX6_Rate = objSRN.TAX6_Rate
        objPO.TAX6_Amt = objSRN.TAX6_Amt
        objPO.TAX6_Base_Amt = objSRN.TAX6_Base_Amt
        objPO.TAX7 = objSRN.TAX7
        objPO.TAX7_Rate = objSRN.TAX7_Rate
        objPO.TAX7_Amt = objSRN.TAX7_Amt
        objPO.TAX7_Base_Amt = objSRN.TAX7_Base_Amt
        objPO.TAX8 = objSRN.TAX8
        objPO.TAX8_Rate = objSRN.TAX8_Rate
        objPO.TAX8_Amt = objSRN.TAX8_Amt
        objPO.TAX8_Base_Amt = objSRN.TAX8_Base_Amt
        objPO.TAX9 = objSRN.TAX9
        objPO.TAX9_Rate = objSRN.TAX9_Rate
        objPO.TAX9_Amt = objSRN.TAX9_Amt
        objPO.TAX9_Base_Amt = objSRN.TAX9_Base_Amt
        objPO.TAX10 = objSRN.TAX10
        objPO.TAX10_Rate = objSRN.TAX10_Rate
        objPO.TAX10_Amt = objSRN.TAX10_Amt
        objPO.TAX10_Base_Amt = objSRN.TAX10_Base_Amt
        objPO.Discount_Base = objSRN.Discount_Base
        objPO.Discount_Amt = objSRN.Discount_Amt
        objPO.Amount_Less_Discount = objSRN.Amount_Less_Discount
        objPO.Total_Tax_Amt = objSRN.Total_Tax_Amt
        objPO.PO_Total_Amt = objSRN.SRN_Total_Amt
        objPO.Comments = objSRN.Comments
        objPO.Comp_Code = objSRN.Comp_Code
        objPO.Terms_Code = objSRN.Terms_Code
        objPO.Due_Date = objSRN.Due_Date
        'objPO.Mode_Of_Transport = objSRN.Mode_Of_Transport
        'objPO.Created_By = objSRN.Created_By
        'objPO.Created_Date = objSRN.Created_Date
        'objPO.Modify_By = objSRN.Modify_By
        'objPO.Modify_Date = objSRN.Modify_Date
        'objPO.Posting_Date = objSRN.Posting_Date
        objPO.Delivery_date = objSRN.SRN_Date
        objPO.Dept = objSRN.Dept
        objPO.Dept_Desc = objSRN.Dept_Desc
        objPO.Item_Type = objSRN.Item_Type
        'objPO.Abandonment_No = objSRN.Abandonment_No
        objPO.Against_Requisition = objSRN.Against_Requisition
        objPO.Add_Charge_Code1 = objSRN.Add_Charge_Code1
        objPO.Add_Charge_Name1 = objSRN.Add_Charge_Name1
        objPO.Add_Charge_Amt1 = objSRN.Add_Charge_Amt1
        objPO.Add_Charge_Code2 = objSRN.Add_Charge_Code2
        objPO.Add_Charge_Name2 = objSRN.Add_Charge_Name2
        objPO.Add_Charge_Amt2 = objSRN.Add_Charge_Amt2
        objPO.Add_Charge_Code3 = objSRN.Add_Charge_Code3
        objPO.Add_Charge_Name3 = objSRN.Add_Charge_Name3
        objPO.Add_Charge_Amt3 = objSRN.Add_Charge_Amt3
        objPO.Add_Charge_Code4 = objSRN.Add_Charge_Code4
        objPO.Add_Charge_Name4 = objSRN.Add_Charge_Name4
        objPO.Add_Charge_Amt4 = objSRN.Add_Charge_Amt4
        objPO.Add_Charge_Code5 = objSRN.Add_Charge_Code5
        objPO.Add_Charge_Name5 = objSRN.Add_Charge_Name5
        objPO.Add_Charge_Amt5 = objSRN.Add_Charge_Amt5
        objPO.Add_Charge_Code6 = objSRN.Add_Charge_Code6
        objPO.Add_Charge_Name6 = objSRN.Add_Charge_Name6
        objPO.Add_Charge_Amt6 = objSRN.Add_Charge_Amt6
        objPO.Add_Charge_Code7 = objSRN.Add_Charge_Code7
        objPO.Add_Charge_Name7 = objSRN.Add_Charge_Name7
        objPO.Add_Charge_Amt7 = objSRN.Add_Charge_Amt7
        objPO.Add_Charge_Code8 = objSRN.Add_Charge_Code8
        objPO.Add_Charge_Name8 = objSRN.Add_Charge_Name8
        objPO.Add_Charge_Amt8 = objSRN.Add_Charge_Amt8
        objPO.Add_Charge_Code9 = objSRN.Add_Charge_Code9
        objPO.Add_Charge_Name9 = objSRN.Add_Charge_Name9
        objPO.Add_Charge_Amt9 = objSRN.Add_Charge_Amt9
        objPO.Add_Charge_Code10 = objSRN.Add_Charge_Code10
        objPO.Add_Charge_Name10 = objSRN.Add_Charge_Name10
        objPO.Add_Charge_Amt10 = objSRN.Add_Charge_Amt10
        objPO.Total_Add_Charge = objSRN.Total_Add_Charge
        'objPO.Against_RGP = objSRN.Against_RGP
        objPO.Tax_Calculation_Type = objSRN.Tax_Calculation_Type
        'objPO.Against_RGP_NO = objSRN.Against_RGP_NO
        'objPO.Quotation_No = objSRN.Quotation_No
        'objPO.Quotation_Date = objSRN.Quotation_Date
        objPO.is_Excise_On_Qty = objSRN.is_Excise_On_Qty
        objPO.AssessableAmt = objSRN.AssessableAmt
        'objPO.Against_Vendor_Quotation = objSRN.Against_Vendor_Quotation
        'objPO.Terms_Remark = objSRN.Terms_Remark
        objPO.CURRENCY_CODE = objSRN.CURRENCY_CODE
        objPO.ConvRate = objSRN.ConvRate
        objPO.ApplicableFrom = objSRN.ApplicableFrom
        'objPO.Against_C_Form = objSRN.Against_C_Form
        'objPO.CFormRecd = objSRN.CFormRecd
        'objPO.CFormApplied = objSRN.CFormApplied
        'objPO.Delivery_Duration = objSRN.Delivery_Duration
        objPO.PROJECT_ID = objSRN.PROJECT_ID
        'objPO.close_yn = objSRN.close_yn
        objPO.IsAbatementPO = objSRN.IsAbatementPO
        'objPO.Bin_No = objSRN.Bin_No
        'objPO.Expiry_Date = objSRN.Expiry_Date
        'objPO.Approval_Level = objSRN.Approval_Level
        'objPO.Level1_User = objSRN.Level1_User
        'objPO.Level2_User = objSRN.Level2_User
        'objPO.Level3_User = objSRN.Level3_User
        'objPO.MCC_Purchase = objSRN.MCC_Purchase
        'objPO.State_Code = objSRN.State_Code
        'objPO.PO_Amount = objSRN.PO_AMOUNT
        'objPO.isBlanket = objSRN.isBlanket
        'objPO.Auto_Purchase = objSRN.Auto_Purchase
        'objPO.Issue_Road_Permit = objSRN.Issue_Road_Permit
        'objPO.Issue_C_Form = objSRN.Issue_C_Form
        'objPO.Is_Approved = objSRN.Is_Approved
        'objPO.SaleInvoiceNo = objSRN.SaleInvoiceNo
        'objPO.SaleInvoice_Type = objSRN.SaleInvoice_Type
        'objPO.MT_Is_Merchant_Trade = objSRN.MT_Is_Merchant_Trade
        'objPO.MT_PI_No = objSRN.MT_PI_No
        'objPO.MT_PI_Status = objSRN.MT_PI_Status
        'objPO.MT_PI_Status_Date = objSRN.MT_PI_Status_Date
        'objPO.MT_Payment_Terms_Group_Code = objSRN.MT_Payment_Terms_Group_Code
        'objPO.MT_Is_AmountinRs = objSRN.MT_Is_AmountinRs
        'objPO.MT_LC = objSRN.MT_LC
        'objPO.MT_CAD = objSRN.MT_CAD
        'objPO.MT_RETAINED = objSRN.MT_RETAINED
        'objPO.MT_Balance_Payment = objSRN.MT_Balance_Payment
        'objPO.MT_On_Account = objSRN.MT_On_Account
        'objPO.MT_Advance = objSRN.MT_Advance
        'objPO.MT_Beneficiary_Code = objSRN.MT_Beneficiary_Code
        'objPO.MT_INCOTERMS = objSRN.MT_INCOTERMS
        'objPO.Is_Open_PO = objSRN.Is_Open_PO
        'objPO.Against_PO = objSRN.Against_PO
        'objPO.Renewal_Date = objSRN.Renewal_Date
        'objPO.Payment_Terms = objSRN.Payment_Terms
        'objPO.Insurance_Terms = objSRN.Insurance_Terms
        'objPO.Delivery_Terms_Code = objSRN.Delivery_Terms_Code
        'objPO.Auto_Calculate = objSRN.Auto_Calculate
        'objPO.Subject = objSRN.Subject
        'objPO.Content_Subject = objSRN.Content_Subject
        'objPO.Kind_Attentation = objSRN.Kind_Attentation
        'objPO.MT_HS_Classification_No = objSRN.MT_HS_Classification_No
        'objPO.IsPO = objSRN.IsPO
        'objPO.IsContent = objSRN.IsContent
        'objPO.MT_Pre_Carriage_By = objSRN.MT_Pre_Carriage_By
        'objPO.MT_Carrier = objSRN.MT_Carrier
        'objPO.MT_Discharge_Port = objSRN.MT_Discharge_Port
        'objPO.MT_Final_Destination = objSRN.MT_Final_Destination
        'objPO.MT_Origin_Country = objSRN.MT_Origin_Country
        'objPO.MT_Final_Destination_Country = objSRN.MT_Final_Destination_Country
        'objPO.MT_CreditTerms_Code = objSRN.MT_CreditTerms_Code
        'objPO.MT_PI_Due_Date = objSRN.MT_PI_Due_Date
        'objPO.MT_Stuffing_Status = objSRN.MT_Stuffing_Status
        'objPO.MT_Payment_Terms = objSRN.MT_Payment_Terms
        'objPO.MT_EX_Term_Code = objSRN.MT_EX_Term_Code
        'objPO.MT_is_Accepted = objSRN.MT_is_Accepted
        'objPO.MT_Accepted_Date = objSRN.MT_Accepted_Date
        'objPO.MT_is_Partshipment = objSRN.MT_is_Partshipment
        'objPO.MT_is_Transshipment = objSRN.MT_is_Transshipment
        'objPO.MT_CreditTermsName = objSRN.MT_CreditTermsName
        'objPO.MT_CIF = objSRN.MT_CIF
        'objPO.MT_Advance_Type = objSRN.MT_Advance_Type
        'objPO.MT_PT_Advance_Amount = objSRN.MT_PT_Advance_Amount
        'objPO.MT_is_Partpayment = objSRN.MT_is_Partpayment
        'objPO.MT_Buyer_PO_No = objSRN.MT_Buyer_PO_No
        'objPO.MT_Buyer_PO_Date = objSRN.MT_Buyer_PO_Date
        'objPO.Capex_Code = objSRN.Capex_Code
        'objPO.Capex_SubCode = objSRN.Capex_SubCode
        'objPO.IsCancel = objSRN.IsCancel
        'objPO.Category = objSRN.Category
        'objPO.Capacity = objSRN.Capacity
        'objPO.Make = objSRN.Make
        'objPO.Model = objSRN.Model
        objPO.Amt_After_Tax = objSRN.SRN_Total_Amt
        'objPO.Add_Charge_Apply_On1 = objSRN.Add_Charge_Apply_On1
        'objPO.Add_Charge_Per1 = objSRN.Add_Charge_Per1
        'objPO.Add_Charge_Apply_On2 = objSRN.Add_Charge_Apply_On2
        'objPO.Add_Charge_Per2 = objSRN.Add_Charge_Per2
        'objPO.Add_Charge_Apply_On3 = objSRN.Add_Charge_Apply_On3
        'objPO.Add_Charge_Per3 = objSRN.Add_Charge_Per3
        'objPO.Add_Charge_Apply_On4 = objSRN.Add_Charge_Apply_On4
        'objPO.Add_Charge_Per4 = objSRN.Add_Charge_Per4
        'objPO.Add_Charge_Apply_On5 = objSRN.Add_Charge_Apply_On5
        'objPO.Add_Charge_Per5 = objSRN.Add_Charge_Per5
        'objPO.Add_Charge_Apply_On6 = objSRN.Add_Charge_Apply_On6
        'objPO.Add_Charge_Per6 = objSRN.Add_Charge_Per6
        'objPO.Add_Charge_Apply_On7 = objSRN.Add_Charge_Apply_On7
        'objPO.Add_Charge_Per7 = objSRN.Add_Charge_Per7
        'objPO.Add_Charge_Apply_On8 = objSRN.Add_Charge_Apply_On8
        'objPO.Add_Charge_Per8 = objSRN.Add_Charge_Per8
        'objPO.Add_Charge_Apply_On9 = objSRN.Add_Charge_Apply_On9
        'objPO.Add_Charge_Per9 = objSRN.Add_Charge_Per9
        'objPO.Add_Charge_Apply_On10 = objSRN.Add_Charge_Apply_On10
        'objPO.Add_Charge_Per10 = objSRN.Add_Charge_Per10
        'objPO.Emergency = objSRN.Emergency
        'objPO.Apply_Receive_Control = objSRN.Apply_Receive_Control
        'objPO.Delivery_days = objSRN.Delivery_days
        'objPO.Is_Auto_Weighment_Type = objSRN.Is_Auto_Weighment_Type
        'objPO.Amendment_Code = objSRN.Amendment_Code
        'objPO.Amendment_By = objSRN.Amendment_By
        'objPO.Amendment_Date = objSRN.Amendment_Date
        'objPO.ServiceBill_No = objSRN.ServiceBill_No
        'objPO.ServiceBill_Date = objSRN.ServiceBill_Date
        'objPO.Closed_By = objSRN.Closed_By
        'objPO.Closed_Date = objSRN.Closed_Date
        'objPO.Posted_By = objSRN.Posted_By
        'objPO.Bank_Code = objSRN.Bank_Code
        'objPO.Payment_Code = objSRN.Payment_Code
        'objPO.GSTRegistered = objSRN.GSTRegistered
        objPO.Sublocation_Code = objSRN.Sublocation_Code
        objPO.isJobWorkOutward = objSRN.isJobWorkOutward
        objPO.Total_Taxable_Amount = objSRN.Total_Taxable_Amount
        'objPO.WorkOrder_Vendor = objSRN.WorkOrder_Vendor
        'objPO.WorkOrder_Vendor_Phn = objSRN.WorkOrder_Vendor_Phn
        'objPO.WorkOrder_Vendor_Add = objSRN.WorkOrder_Vendor_Add
        'objPO.close_remarks = objSRN.close_remarks
        'objPO.Is_Repair = objSRN.Is_Repair
        'objPO.Freight = objSRN.Freight
        'objPO.Packing_Forward = objSRN.Packing_Forward
        'objPO.Insurance = objSRN.Insurance
        'objPO.Header_Discount_Amount = objSRN.Header_Discount_Amount
        'objPO.ReferencePO = objSRN.ReferencePO
        'objPO.Against_Tender = objSRN.Against_Tender
        objPO.Confirmatory_PO_SRN_No = objSRN.SRN_No
        objPO.Arr = New List(Of clsPurchaseOrderDetail)
        For Each objSRNDetail As clsSRNDetail In objSRN.Arr
            Dim objPODetail As New clsPurchaseOrderDetail()
            'objPODetail.PurchaseOrder_No = objSRNDetail.PurchaseOrder_No
            objPODetail.Line_No = objSRNDetail.Line_No
            objPODetail.Item_Code = objSRNDetail.Item_Code
            objPODetail.Item_Desc = objSRNDetail.Item_Desc
            objPODetail.PurchaseOrder_Qty = objSRNDetail.SRN_Qty
            'objPODetail.Requisition_Id = objSRNDetail.Requisition_Id
            objPODetail.Balance_Qty = objSRNDetail.Balance_Qty
            objPODetail.Unit_code = objSRNDetail.Unit_code
            objPODetail.Location = objSRNDetail.Location
            objPODetail.Item_Cost = objSRNDetail.Item_Cost
            objPODetail.TAX1 = objSRNDetail.TAX1
            objPODetail.TAX1_Base_Amt = objSRNDetail.TAX1_Base_Amt
            objPODetail.TAX1_Rate = objSRNDetail.TAX1_Rate
            objPODetail.TAX1_Amt = objSRNDetail.TAX1_Amt
            objPODetail.TAX2 = objSRNDetail.TAX2
            objPODetail.TAX2_Base_Amt = objSRNDetail.TAX2_Base_Amt
            objPODetail.TAX2_Rate = objSRNDetail.TAX2_Rate
            objPODetail.TAX2_Amt = objSRNDetail.TAX2_Amt
            objPODetail.TAX3 = objSRNDetail.TAX3
            objPODetail.TAX3_Base_Amt = objSRNDetail.TAX3_Base_Amt
            objPODetail.TAX3_Rate = objSRNDetail.TAX3_Rate
            objPODetail.TAX3_Amt = objSRNDetail.TAX3_Amt
            objPODetail.TAX4 = objSRNDetail.TAX4
            objPODetail.TAX4_Base_Amt = objSRNDetail.TAX4_Base_Amt
            objPODetail.TAX4_Rate = objSRNDetail.TAX4_Rate
            objPODetail.TAX4_Amt = objSRNDetail.TAX4_Amt
            objPODetail.TAX5 = objSRNDetail.TAX5
            objPODetail.TAX5_Base_Amt = objSRNDetail.TAX5_Base_Amt
            objPODetail.TAX5_Rate = objSRNDetail.TAX5_Rate
            objPODetail.TAX5_Amt = objSRNDetail.TAX5_Amt
            objPODetail.TAX6 = objSRNDetail.TAX6
            objPODetail.TAX6_Base_Amt = objSRNDetail.TAX6_Base_Amt
            objPODetail.TAX6_Rate = objSRNDetail.TAX6_Rate
            objPODetail.TAX6_Amt = objSRNDetail.TAX6_Amt
            objPODetail.TAX7 = objSRNDetail.TAX7
            objPODetail.TAX7_Base_Amt = objSRNDetail.TAX7_Base_Amt
            objPODetail.TAX7_Rate = objSRNDetail.TAX7_Rate
            objPODetail.TAX7_Amt = objSRNDetail.TAX7_Amt
            objPODetail.TAX8 = objSRNDetail.TAX8
            objPODetail.TAX8_Base_Amt = objSRNDetail.TAX8_Base_Amt
            objPODetail.TAX8_Rate = objSRNDetail.TAX8_Rate
            objPODetail.TAX8_Amt = objSRNDetail.TAX8_Amt
            objPODetail.TAX9 = objSRNDetail.TAX9
            objPODetail.TAX9_Base_Amt = objSRNDetail.TAX9_Base_Amt
            objPODetail.TAX9_Rate = objSRNDetail.TAX9_Rate
            objPODetail.TAX9_Amt = objSRNDetail.TAX9_Amt
            objPODetail.TAX10 = objSRNDetail.TAX10
            objPODetail.TAX10_Base_Amt = objSRNDetail.TAX10_Base_Amt
            objPODetail.TAX10_Rate = objSRNDetail.TAX10_Rate
            objPODetail.TAX10_Amt = objSRNDetail.TAX10_Amt
            objPODetail.Amount = objSRNDetail.Amount
            objPODetail.Disc_Per = objSRNDetail.Disc_Per
            objPODetail.Disc_Amt = objSRNDetail.Disc_Amt
            objPODetail.Amt_Less_Discount = objSRNDetail.Amt_Less_Discount
            objPODetail.Total_Tax_Amt = objSRNDetail.Total_Tax_Amt
            objPODetail.Item_Net_Amt = objSRNDetail.Item_Net_Amt
            objPODetail.Status = objSRNDetail.Status
            objPODetail.Specification = objSRNDetail.Specification
            objPODetail.Remarks = objSRNDetail.Remarks
            objPODetail.MRP = objSRNDetail.MRP
            objPODetail.Assessable = objSRNDetail.Assessable
            objPODetail.AssessableAmt = objSRNDetail.AssessableAmt
            objPODetail.Row_Type = objSRNDetail.Row_Type
            objPODetail.AbatementRate = objSRNDetail.AbatementRate
            objPODetail.AssessableMRP = objSRNDetail.AssessableMRP
            objPODetail.TotalAssessableMRP = objSRNDetail.TotalAssessableMRP
            objPODetail.Bin_No = objSRNDetail.Bin_No
            'objPODetail.Last_Same_Vendor_Rate = objSRNDetail.Last_Same_Vendor_Rate
            'objPODetail.Last_Other_Vendor_Rate = objSRNDetail.Last_Other_Vendor_Rate
            'objPODetail.Qty_Desc = objSRNDetail.Qty_Desc
            'objPODetail.Rate_Desc = objSRNDetail.Rate_Desc
            'objPODetail.Amount_Desc = objSRNDetail.Amount_Desc
            'objPODetail.FatPer_MT = objSRNDetail.FatPer_MT
            'objPODetail.SNFPer_MT = objSRNDetail.SNFPer_MT
            'objPODetail.FatKG_MT = objSRNDetail.FatKG_MT
            'objPODetail.SNFKG_MT = objSRNDetail.SNFKG_MT
            'objPODetail.Item_Weight_MT = objSRNDetail.Item_Weight_MT
            'objPODetail.Weight_UOM_MT = objSRNDetail.Weight_UOM_MT
            objPODetail.ItemAdd_Charge_Code1 = objSRNDetail.ItemAdd_Charge_Code1
            objPODetail.ItemAdd_Org_Charge_Amt1 = objSRNDetail.ItemAdd_Org_Charge_Amt1
            objPODetail.ItemAdd_Calc_Charge_Amt1 = objSRNDetail.ItemAdd_Calc_Charge_Amt1
            objPODetail.ItemAdd_Charge_Code2 = objSRNDetail.ItemAdd_Charge_Code2
            objPODetail.ItemAdd_Org_Charge_Amt2 = objSRNDetail.ItemAdd_Org_Charge_Amt2
            objPODetail.ItemAdd_Calc_Charge_Amt2 = objSRNDetail.ItemAdd_Calc_Charge_Amt2
            objPODetail.ItemAdd_Charge_Code3 = objSRNDetail.ItemAdd_Charge_Code3
            objPODetail.ItemAdd_Org_Charge_Amt3 = objSRNDetail.ItemAdd_Org_Charge_Amt3
            objPODetail.ItemAdd_Calc_Charge_Amt3 = objSRNDetail.ItemAdd_Calc_Charge_Amt3
            objPODetail.ItemAdd_Charge_Code4 = objSRNDetail.ItemAdd_Charge_Code4
            objPODetail.ItemAdd_Org_Charge_Amt4 = objSRNDetail.ItemAdd_Org_Charge_Amt4
            objPODetail.ItemAdd_Calc_Charge_Amt4 = objSRNDetail.ItemAdd_Calc_Charge_Amt4
            objPODetail.ItemAdd_Charge_Code5 = objSRNDetail.ItemAdd_Charge_Code5
            objPODetail.ItemAdd_Org_Charge_Amt5 = objSRNDetail.ItemAdd_Org_Charge_Amt5
            objPODetail.ItemAdd_Calc_Charge_Amt5 = objSRNDetail.ItemAdd_Calc_Charge_Amt5
            objPODetail.ItemAdd_Charge_Code6 = objSRNDetail.ItemAdd_Charge_Code6
            objPODetail.ItemAdd_Org_Charge_Amt6 = objSRNDetail.ItemAdd_Org_Charge_Amt6
            objPODetail.ItemAdd_Calc_Charge_Amt6 = objSRNDetail.ItemAdd_Calc_Charge_Amt6
            objPODetail.ItemAdd_Charge_Code7 = objSRNDetail.ItemAdd_Charge_Code7
            objPODetail.ItemAdd_Org_Charge_Amt7 = objSRNDetail.ItemAdd_Org_Charge_Amt7
            objPODetail.ItemAdd_Calc_Charge_Amt7 = objSRNDetail.ItemAdd_Calc_Charge_Amt7
            objPODetail.ItemAdd_Charge_Code8 = objSRNDetail.ItemAdd_Charge_Code8
            objPODetail.ItemAdd_Org_Charge_Amt8 = objSRNDetail.ItemAdd_Org_Charge_Amt8
            objPODetail.ItemAdd_Calc_Charge_Amt8 = objSRNDetail.ItemAdd_Calc_Charge_Amt8
            objPODetail.ItemAdd_Charge_Code9 = objSRNDetail.ItemAdd_Charge_Code9
            objPODetail.ItemAdd_Org_Charge_Amt9 = objSRNDetail.ItemAdd_Org_Charge_Amt9
            objPODetail.ItemAdd_Calc_Charge_Amt9 = objSRNDetail.ItemAdd_Calc_Charge_Amt9
            objPODetail.ItemAdd_Charge_Code10 = objSRNDetail.ItemAdd_Charge_Code10
            objPODetail.ItemAdd_Org_Charge_Amt10 = objSRNDetail.ItemAdd_Org_Charge_Amt10
            objPODetail.ItemAdd_Calc_Charge_Amt10 = objSRNDetail.ItemAdd_Calc_Charge_Amt10
            objPODetail.Total_ItemAdd_Charge = objSRNDetail.Total_ItemAdd_Charge
            'objPODetail.Capacity = objSRNDetail.Capacity
            'objPODetail.Make = objSRNDetail.Make
            'objPODetail.Model = objSRNDetail.Model
            objPODetail.Against_Item_Wise_Tax_Rate = objSRNDetail.Against_Item_Wise_Tax_Rate
            objPODetail.Taxable_Amount_Per = objSRNDetail.Taxable_Amount_Per
            objPODetail.Taxable_Amount = objSRNDetail.Taxable_Amount
            objPODetail.Insurance_Base_Amt = objSRNDetail.Insurance_Base_Amt
            objPODetail.Insurance_Per = objSRNDetail.Insurance_Per
            objPODetail.Header_Discount_Per = objSRNDetail.Header_Discount_Per
            objPODetail.Header_Discount_Amount = objSRNDetail.Header_Discount_Amount
            objPODetail.Detail_Discount_Amount = objSRNDetail.Detail_Discount_Amount
            objPO.Arr.Add(objPODetail)
        Next
        objPO.SaveData(objPO, True, False, trans)
        objPO.PostData(clsUserMgtCode.mbtnPurchaseOrder, objPO.PurchaseOrder_No, False, False, False, trans, Nothing, Nothing)

        Return True
    End Function

    Public Shared Function CreateJournalEntry(ByVal obj As clsSRNHead, ByVal strVoucherNoRecreateOnly As String, ByVal trans As SqlTransaction, Optional ByVal WithoutRejection As Boolean = False) As Boolean
        Try
            Dim qry As String = ""
            Dim ArryLstGLAC As ArrayList = New ArrayList()
            Dim strFirstItemCodeNonItemRowType As String = GetFirstItemCode(obj.Arr)
            Dim strRgpNo As String = Nothing
            Dim intCounter As Integer = 0
            Dim isApplyPurchaseAccounting As Boolean = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 1)
            If Not isApplyPurchaseAccounting Then
                If Not (obj.is_RGP_Non_Inventory AndAlso clsCommon.myLen(obj.Against_RGP) > 0) Then
                    Dim strBrachAC As String = ""
                    If clsCommon.myLen(obj.Ship_To_Location) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(obj.Ship_To_Location), clsCommon.myCstr(obj.Bill_To_Location)) <> CompairStringResult.Equal Then
                            qry = "select Branch_Account from TSPL_BRANCH_ACCOUNT_MAPPING where From_Location='" + clsLocation.GetSegmentCode(obj.Bill_To_Location, trans) + "' and To_Location='" + clsLocation.GetSegmentCode(obj.Ship_To_Location, trans) + "'"
                            strBrachAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                            If clsCommon.myLen(strBrachAC) <= 0 Then
                                Throw New Exception("Plase set Brach account with From_Location=" + clsLocation.GetSegmentCode(obj.Bill_To_Location, trans) + " and To_Location=" + clsLocation.GetSegmentCode(obj.Ship_To_Location, trans) + "")
                            End If
                        End If
                    End If
                    Dim isAgainstTender As Boolean = clsPurchaseOrderHead.AgainstTender(obj.Against_GRN, 1, trans)
                    For Each objTr As clsSRNDetail In obj.Arr
                        intCounter = intCounter + 1
                        If clsCommon.CompairString(objTr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                            If clsCommon.myLen(objTr.MRN_Id) > 0 Then
                                Dim qry1 As String = "update TSPL_MRN_DETAIL set Balance_Qty=Balance_Qty - " + clsCommon.myCstr(clsCommon.myCdbl(objTr.SRN_Qty + objTr.Rejected_Qty + objTr.Leak_Qty + objTr.Burst_Qty + objTr.Short_Qty)) + " where MRN_No='" + objTr.MRN_Id + "' and Item_Code='" + objTr.Item_Code + "' and Unit_code='" + objTr.Unit_code + "' and isnull(MRP,0)='" + clsCommon.myCstr(objTr.MRP) + "' and isnull(Assessable,0)='" + clsCommon.myCstr(objTr.Assessable) + "'"
                                clsDBFuncationality.ExecuteNonQuery(qry1, trans)
                            End If
                        Else
                            objTr.Item_Code = strFirstItemCodeNonItemRowType
                        End If

                        If clsCommon.myLen(objTr.RGP_Id) > 0 Then
                            strRgpNo = objTr.RGP_Id
                        End If


                        qry = "select TSPL_ITEM_MASTER.Purchase_Class_Code,TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_PURCHASE_ACCOUNTS.Inv_Payable_Clearing,TSPL_PURCHASE_ACCOUNTS.Assembly_Cost_Credit,TSPL_PURCHASE_ACCOUNTS.Breakage_Gl_Account,TSPL_PURCHASE_ACCOUNTS.RM_Consumption,TSPL_PURCHASE_ACCOUNTS.Other_1 ,TSPL_PURCHASE_ACCOUNTS.Other_2,TSPL_PURCHASE_ACCOUNTS.FA_CLEARING_AC  from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objTr.Item_Code + "'"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Please set Purchase Account set for item " + objTr.Item_Code + "(" + objTr.Item_Desc + ")")
                        End If
                        ''1)
                        Dim strInvCtrlAC As String = Nothing
                        If clsCommon.CompairString(clsCommon.myCstr(objTr.Category), "Capex") = CompairStringResult.Equal Then
                            strInvCtrlAC = clsCommon.myCstr(dt.Rows(0)("FA_CLEARING_AC"))
                            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                                Throw New Exception("Please set FA Clearing Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                            End If
                        Else
                            strInvCtrlAC = clsCommon.myCstr(dt.Rows(0)("Inv_Control_Account"))
                            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                                Throw New Exception("Please set Inventory Control Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                            End If
                        End If
                        strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, IIf(clsCommon.myLen(obj.Ship_To_Location) > 0, obj.Ship_To_Location, obj.Bill_To_Location), trans)
                        ' ''richa agarwal 4 Jan,2019 BHA/27/11/18-000720
                        If clsCommon.CompairString(clsCommon.myCstr(objTr.Category), "Capex") = CompairStringResult.Equal Then
                            Dim AccDr() As String = {strInvCtrlAC, Math.Round(((objTr.Accepted_Amount)), 2, MidpointRounding.ToEven)}
                            ArryLstGLAC.Add(AccDr)
                        Else
                            Dim AccDr() As String = {strInvCtrlAC, Math.Round(((objTr.Accepted_Amount)), 2, MidpointRounding.ToEven), "", "", "", "", "", "", "I"}
                            ArryLstGLAC.Add(AccDr)

                            ' ''richa agarwal 4 Jan,2019 BHA/27/11/18-000720 inevntory movement work
                            clsInventoryMovement.UpdateInvControlAccount(obj.SRN_No, "SRN", objTr.Item_Code, strInvCtrlAC, "", "", trans)
                            ''------------------
                        End If



                        ''2
                        If (objTr.Leak_Amount + objTr.Burst_Amount) > 0 Then
                            Dim strBrakageCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Breakage_Gl_Account"))
                            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                                Throw New Exception("Please set Breakage Control Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                            End If
                            strBrakageCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strBrakageCtrlAC, IIf(clsCommon.myLen(obj.Ship_To_Location) > 0, obj.Ship_To_Location, obj.Bill_To_Location), trans)
                            Dim AccDr1() As String = {strBrakageCtrlAC, Math.Round((objTr.Leak_Amount + objTr.Burst_Amount), 2, MidpointRounding.ToEven)}
                            ArryLstGLAC.Add(AccDr1)
                        End If


                        ''3
                        If (objTr.Shortage_Amount) > 0 Then
                            If Not isAgainstTender Then
                                Dim strShortageCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Other_2"))
                                If clsCommon.myLen(strShortageCtrlAC) <= 0 Then
                                    Throw New Exception("Please set FG Shortage Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                                End If
                                strShortageCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShortageCtrlAC, IIf(clsCommon.myLen(obj.Ship_To_Location) > 0, obj.Ship_To_Location, obj.Bill_To_Location), trans)
                                Dim AccDr1() As String = {strShortageCtrlAC, Math.Round(objTr.Shortage_Amount, 2, MidpointRounding.ToEven)}
                                ArryLstGLAC.Add(AccDr1)
                            End If
                        End If


                        ''4
                        If WithoutRejection = False Then
                            If objTr.Rejected_Amount > 0 Then
                                Dim strRejectedAcc As String = clsCommon.myCstr(dt.Rows(0)("Other_1"))
                                If clsCommon.myLen(strRejectedAcc) <= 0 Then
                                    Throw New Exception("Please set [Rejected Inventory Account] for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                                End If
                                strRejectedAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strRejectedAcc, IIf(clsCommon.myLen(obj.Ship_To_Location) > 0, obj.Ship_To_Location, obj.Bill_To_Location), trans)
                                Dim AccDrRej() As String = {strRejectedAcc, Math.Round(objTr.Rejected_Amount, 2, MidpointRounding.ToEven)}
                                ArryLstGLAC.Add(AccDrRej)
                            End If
                        End If

                        ''5
                        Dim AccCr() As String = Nothing
                        Dim strPaybleCleanigCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Inv_Payable_Clearing"))
                        If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                            Throw New Exception("Please set Payable Clearing Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                        End If
                        strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strPaybleCleanigCtrlAC, IIf(clsCommon.myLen(obj.Ship_To_Location) > 0, obj.Ship_To_Location, obj.Bill_To_Location), trans)
                        If WithoutRejection = False Then
                            AccCr = {strPaybleCleanigCtrlAC, -1 * Math.Round(((objTr.Accepted_Amount + objTr.Leak_Amount + objTr.Burst_Amount + IIf(isAgainstTender, 0, objTr.Shortage_Amount) + objTr.Rejected_Amount)), 2, MidpointRounding.ToEven), "", "", "", "", "", "", "Y"}
                        Else
                            AccCr = {strPaybleCleanigCtrlAC, -1 * Math.Round(((objTr.Accepted_Amount + objTr.Leak_Amount + objTr.Burst_Amount + IIf(isAgainstTender, 0, objTr.Shortage_Amount))), 2, MidpointRounding.ToEven), "", "", "", "", "", "", "Y"}

                        End If
                        ArryLstGLAC.Add(AccCr)
                        ''6
                        If clsCommon.myLen(obj.Ship_To_Location) > 0 Then
                            If clsCommon.CompairString(clsCommon.myCstr(obj.Ship_To_Location), clsCommon.myCstr(obj.Bill_To_Location)) <> CompairStringResult.Equal Then
                                If clsCommon.CompairString(clsCommon.myCstr(obj.Ship_To_Location), clsCommon.myCstr(obj.Bill_To_Location)) = CompairStringResult.Equal Then
                                    strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strPaybleCleanigCtrlAC, obj.Bill_To_Location, trans)
                                Else
                                    strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountWithOutLOcSegment(strPaybleCleanigCtrlAC.Substring(0, strPaybleCleanigCtrlAC.Length - 4), obj.Bill_To_Location, False, trans)
                                End If
                                Dim AccCr2() As String = {strPaybleCleanigCtrlAC, -1 * Math.Round(objTr.Accepted_Amount, 2, MidpointRounding.ToEven), "", "", "", "", "", "", "Y"}
                                ArryLstGLAC.Add(AccCr2)

                                Dim AccCr3() As String = {strBrachAC, Math.Round(objTr.Accepted_Amount, 2, MidpointRounding.ToEven)}
                                ArryLstGLAC.Add(AccCr3)
                            End If
                        End If
                    Next
                    If clsCommon.myLen(strRgpNo) <= 0 Then
                        If strVoucherNoRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoRecreateOnly) > 0 Then
                            clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, strVoucherNoRecreateOnly, trans, obj.SRN_Date, "Against Store Received Note " + obj.SRN_No, "PO-RC", "Store Received Note", obj.SRN_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstGLAC)
                        Else
                            clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, trans, obj.SRN_Date, "Against Store Received Note " + obj.SRN_No, "PO-RC", "Store Received Note", obj.SRN_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstGLAC)
                        End If
                    End If
                End If


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '' Anubhooti 12-Nov-2014/10-Dec-2014 BM00000003662 (Job Work GL Entries)
                Dim ArryLstRGP As ArrayList = New ArrayList()
                Dim PurAccQry As String = ""
                Dim PurSetJobWork As String = ""
                Dim IsJobWork As String = ""
                Dim TotalAmt As Double = 0
                Dim isReparirJobWork As Boolean = False
                Dim dtte As DataTable = clsDBFuncationality.GetDataTable("select ISNULL(Against_JobWork ,'') AS Against_JobWork,ISNULL(Is_Repair,0) as Is_Repair From TSPL_RGP_HEAD Where RGP_No ='" & obj.Against_RGP & "'", trans)
                If dtte IsNot Nothing AndAlso dtte.Rows.Count > 0 Then
                    IsJobWork = clsCommon.myCstr(dtte.Rows(0)("Against_JobWork"))
                    isReparirJobWork = IIf(clsCommon.myCdbl(dtte.Rows(0)("Is_Repair")) > 0, True, False)
                End If

                PurSetJobWork = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ISNULL(Job_Work_Account,'') AS Job_Work_Account From TSPL_PURCHASE_SETTINGS", trans))
                If clsCommon.myLen(obj.Against_RGP) > 0 AndAlso clsCommon.CompairString(IsJobWork, "1") = CompairStringResult.Equal Then
                    For Each objTr As clsSRNDetail In obj.Arr
                        PurAccQry = "select TSPL_ITEM_MASTER.Purchase_Class_Code,TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_PURCHASE_ACCOUNTS.Inv_Payable_Clearing,TSPL_PURCHASE_ACCOUNTS.Assembly_Cost_Credit,TSPL_PURCHASE_ACCOUNTS.Breakage_Gl_Account,TSPL_PURCHASE_ACCOUNTS.RM_Consumption,TSPL_PURCHASE_ACCOUNTS.Reserve_Stock  from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objTr.Item_Code + "'"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(PurAccQry, trans)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Please set Purchase Account set for item " + objTr.Item_Code + "(" + objTr.Item_Desc + ")")
                        End If
                        ''1)
                        Dim strInvCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Inv_Control_Account"))
                        If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                            Throw New Exception("Please set Inventory Control Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
                        End If
                        strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, obj.Bill_To_Location, trans)
                        Dim AccDr() As String = {strInvCtrlAC, Math.Round(((objTr.Amount + objTr.Total_Tax_Amt)), 2, MidpointRounding.ToEven), "", "", "", "", "", "", "I"}
                        ArryLstRGP.Add(AccDr)

                        ' ''richa agarwal 4 Jan,2019 BHA/27/11/18-000720 inevntory movement work
                        clsInventoryMovement.UpdateInvControlAccount(obj.SRN_No, "SRN", objTr.Item_Code, strInvCtrlAC, "", "", trans)
                        ''------------------

                        If isReparirJobWork Then
                            PurSetJobWork = clsCommon.myCstr(dt.Rows(0)("Reserve_Stock"))
                            If clsCommon.myLen(PurSetJobWork) <= 0 Then
                                Throw New Exception("Please set RGP clearing account of purchase account set " + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")))
                            End If
                            PurSetJobWork = clsERPFuncationality.ChangeGLAccountLocationSegment(PurSetJobWork, obj.Bill_To_Location, trans)
                            Dim AccDr1() As String = {PurSetJobWork, -1 * Math.Round(((objTr.Amount + objTr.Total_Tax_Amt)), 2, MidpointRounding.ToEven)}
                            ArryLstRGP.Add(AccDr1)
                        End If
                    Next
                    If Not isReparirJobWork Then
                        If clsCommon.myLen(PurSetJobWork) > 0 Then
                            PurSetJobWork = clsERPFuncationality.ChangeGLAccountLocationSegment(PurSetJobWork, obj.Bill_To_Location, trans)
                            Dim AccCr() As String = {PurSetJobWork, -1 * Math.Round((obj.SRN_Total_Amt), 2, MidpointRounding.ToEven)}
                            ArryLstRGP.Add(AccCr)
                        Else
                            Throw New Exception("Please set job work account in purchase settings")
                        End If
                    End If
                    If strVoucherNoRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoRecreateOnly) > 0 Then
                        clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, strVoucherNoRecreateOnly, trans, obj.SRN_Date, "SRN Against RGP-" & obj.Against_RGP & "", "SR-RG", "RGP Job Work", obj.SRN_No, obj.Remarks, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstRGP)
                    Else
                        clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, trans, obj.SRN_Date, "SRN Against RGP-" & obj.Against_RGP & "", "SR-RG", "RGP Job Work", obj.SRN_No, obj.Remarks, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstRGP)
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function SaveRGPBOMDetail(ByVal obj As clsSRNHead, ByVal trans As SqlTransaction) As Boolean
        Dim coll As New Hashtable()
        Dim dt As New DataTable()
        Try
            If obj IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                For Each objtr As clsSRNDetail In obj.Arr
                    coll = New Hashtable()
                    dt = New DataTable()

                    dt = clsRGPHead.GetRGPTypeBOMItemsDetail(obj.RGP_Type, objtr.SRN_Qty, obj.Vendor_Code, Nothing, objtr.Item_Code, objtr.Unit_code, trans, obj.SRN_No, True)

                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        For Each dr As DataRow In dt.Rows
                            coll = New Hashtable()

                            clsCommon.AddColumnsForChange(coll, "RGP_No", Nothing)
                            clsCommon.AddColumnsForChange(coll, "Item_Code", clsCommon.myCstr(dr("item_code")))
                            clsCommon.AddColumnsForChange(coll, "Unit_Code", clsCommon.myCstr(dr("unit_code")))
                            clsCommon.AddColumnsForChange(coll, "Qty", clsCommon.myCstr(dr("qty")))
                            clsCommon.AddColumnsForChange(coll, "Vendor_Code", obj.Vendor_Code)
                            clsCommon.AddColumnsForChange(coll, "InOut", "I")
                            clsCommon.AddColumnsForChange(coll, "GRN_No", Nothing)
                            clsCommon.AddColumnsForChange(coll, "SRN_No", obj.SRN_No)

                            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_RGP_BOM_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                        Next
                    End If ''dt cond.
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            coll = Nothing
            dt = Nothing
        End Try

        Return True
    End Function

    ''Public Shared Function PostData(ByVal strDocNo As String) As Boolean
    ''    Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
    ''    Try
    ''        Dim isSaved As Boolean = True
    ''        If (clsCommon.myLen(strDocNo) <= 0) Then
    ''            Throw New Exception("SRN No not found to Post")
    ''        End If
    ''        Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy")

    ''        Dim obj As clsSRNHead = clsSRNHead.GetData(strDocNo, NavigatorType.Current, trans)

    ''        If (obj Is Nothing OrElse clsCommon.myLen(obj.SRN_No) <= 0) Then
    ''            Throw New Exception("No Data found to Post")
    ''        End If
    ''        If (obj.Status = 1) Then
    ''            Throw New Exception("Already Post on :" + obj.Posting_Date)
    ''        End If
    ''        If (obj.On_Hold) Then
    ''            Throw New Exception("SRN No " + obj.SRN_No + " Is On Hold.Can't Post it")
    ''        End If
    ''        Dim qry As String = ""
    ''        Dim ArryLstGLAC As ArrayList = New ArrayList()
    ''        Dim ArrLocationDetails As List(Of clsItemLocationDetails) = New List(Of clsItemLocationDetails)()
    ''        Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
    ''        Dim IsRejectedItemFound As Boolean = False
    ''        Dim strFirstItemCodeNonItemRowType As String = GetFirstItemCode(obj.Arr)
    ''        Dim strRgpNo As String = Nothing
    ''        Dim intCounter As Integer = 0
    ''        For Each objTr As clsSRNDetail In obj.Arr
    ''            intCounter = intCounter + 1
    ''            If clsCommon.CompairString(objTr.Row_Type, frmSRN.RowTypeItem) = CompairStringResult.Equal Then
    ''                If clsCommon.myLen(objTr.MRN_Id) > 0 Then
    ''                    Dim qry1 As String = "update TSPL_MRN_DETAIL set Balance_Qty=Balance_Qty - " + clsCommon.myCstr(clsCommon.myCdbl(objTr.SRN_Qty + objTr.Rejected_Qty + objTr.Leak_Qty + objTr.Burst_Qty + objTr.Short_Qty)) + " where MRN_No='" + objTr.MRN_Id + "' and Item_Code='" + objTr.Item_Code + "' and Unit_code='" + objTr.Unit_code + "' and isnull(MRP,0)='" + clsCommon.myCstr(objTr.MRP) + "' and isnull(Assessable,0)='" + clsCommon.myCstr(objTr.Assessable) + "'"
    ''                    clsDBFuncationality.ExecuteNonQuery(qry1, trans)
    ''                End If
    ''            ElseIf clsCommon.myLen(objTr.RGP_Id) > 0 Then
    ''                strRgpNo = objTr.RGP_Id
    ''            Else
    ''                objTr.Item_Code = strFirstItemCodeNonItemRowType
    ''            End If

    ''            If objTr.Rejected_Qty > 0 Then
    ''                IsRejectedItemFound = True
    ''            End If

    ''            qry = "select TSPL_ITEM_MASTER.Purchase_Class_Code,TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_PURCHASE_ACCOUNTS.Inv_Payable_Clearing from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objTr.Item_Code + "'"
    ''            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
    ''            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    ''                Throw New Exception("Please set Purchase Account set for item " + objTr.Item_Code + "(" + objTr.Item_Desc + ")")
    ''            End If



    ''            '***************************** Manoj:07112011 ************
    ''            Dim chkQry As String
    ''            Dim chkRecov As String
    ''            Dim TxAccQ As String
    ''            Dim TxRecovAcc As String
    ''            Dim TxAmt As Decimal = 0
    ''            If objTr.TAX1 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX1 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX1 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = objTr.TAX1_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX2 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX2 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX2 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX2_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX3 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX3 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX3 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX3_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX4 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX4 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX4 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX4_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX5 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX5 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX5 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX5_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX6 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX6 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX6 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX6_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX7 <> "" Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX7 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX7 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX7_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX8 <> "" Or objTr.TAX8 = Nothing Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX8 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX8 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX8_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX9 <> "" Or objTr.TAX9 = Nothing Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX9 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX9 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX9_Amt
    ''                End If
    ''            End If

    ''            If objTr.TAX10 <> "" Or objTr.TAX10 = Nothing Then
    ''                chkQry = "select Tax_Recoverable  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX10 + "'"
    ''                chkRecov = connectSql.RunScalar(chkQry)
    ''                If chkRecov = "N" Then
    ''                    TxAccQ = "select Tax_Recoverable_Account  from TSPL_TAX_MASTER  where Tax_Code = '" + objTr.TAX10 + "'"
    ''                    TxRecovAcc = connectSql.RunScalar(TxAccQ)
    ''                    TxAmt = TxAmt + objTr.TAX10_Amt
    ''                End If
    ''            End If

    ''            '*****************************************************

    ''            Dim strInvCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Inv_Control_Account"))
    ''            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
    ''                Throw New Exception("Please set Inventory Control Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
    ''            End If
    ''            strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, obj.Bill_To_Location, trans)
    ''            Dim AccDr() As String = {strInvCtrlAC, objTr.Amount - objTr.Fater_Amt + TxAmt + IIf(intCounter = 1, obj.Total_Add_Charge, 0)}
    ''            ArryLstGLAC.Add(AccDr)

    ''            Dim strPaybleCleanigCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Inv_Payable_Clearing"))
    ''            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
    ''                Throw New Exception("Please set Payable Clearing Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Item_Code + "(" + objTr.Item_Desc + ") ")
    ''            End If
    ''            strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strPaybleCleanigCtrlAC, obj.Bill_To_Location, trans)
    ''            Dim AccCr() As String = {strPaybleCleanigCtrlAC, -1 * (objTr.Amount - objTr.Fater_Amt + TxAmt + IIf(intCounter = 1, obj.Total_Add_Charge, 0))}
    ''            ArryLstGLAC.Add(AccCr)


    ''            If clsCommon.CompairString(objTr.Row_Type, frmSRN.RowTypeItem) = CompairStringResult.Equal Then
    ''                Dim strItemTypeToSave As String = ""
    ''                If clsCommon.CompairString(obj.Item_Type, "R") = CompairStringResult.Equal Then
    ''                    strItemTypeToSave = "RM"
    ''                ElseIf clsCommon.CompairString(obj.Item_Type, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Item_Type, "O") = CompairStringResult.Equal Then
    ''                    strItemTypeToSave = "OT"
    ''                ElseIf clsCommon.CompairString(obj.Item_Type, "F") = CompairStringResult.Equal Then
    ''                    strItemTypeToSave = "FT"
    ''                End If

    ''                Dim objLocationDetails As New clsItemLocationDetails()
    ''                objLocationDetails.Item_Code = objTr.Item_Code
    ''                objLocationDetails.Item_Desc = objTr.Item_Desc
    ''                objLocationDetails.Location_Code = objTr.Location
    ''                objLocationDetails.Location_Desc = objTr.LocationName
    ''                objLocationDetails.Item_Qty = objTr.SRN_Qty
    ''                objLocationDetails.Amount = objTr.Amount
    ''                objLocationDetails.MRP = objTr.MRP
    ''                If objTr.MFG_Date.HasValue Then
    ''                    objLocationDetails.MFG_Date = objTr.MFG_Date
    ''                End If
    ''                objLocationDetails.Batch_No = objTr.Batch_No
    ''                If objTr.Expiry_Date.HasValue Then
    ''                    objLocationDetails.Expiry_Date = objTr.Expiry_Date
    ''                End If
    ''                objLocationDetails.ItemType = strItemTypeToSave
    ''                ArrLocationDetails.Add(objLocationDetails)

    ''                Dim objInventoryMovemnt As New clsInventoryMovement()
    ''                objInventoryMovemnt.InOut = "I"
    ''                objInventoryMovemnt.Location_Code = objTr.Location
    ''                objInventoryMovemnt.Item_Code = objTr.Item_Code
    ''                objInventoryMovemnt.Item_Desc = objTr.Item_Desc
    ''                objInventoryMovemnt.Qty = objTr.SRN_Qty
    ''                objInventoryMovemnt.UOM = objTr.Unit_code
    ''                objInventoryMovemnt.Basic_Cost = objTr.Amt_Less_Discount
    ''                ''''objInventoryMovemnt.Rec_Cost= objTr.MRP
    ''                objInventoryMovemnt.Add_Cost = objTr.Total_Tax_Amt
    ''                objInventoryMovemnt.Net_Cost = objTr.Item_Net_Amt
    ''                If clsCommon.CompairString(obj.Item_Type, "R") = CompairStringResult.Equal Then
    ''                    objInventoryMovemnt.ItemType = "RM"
    ''                ElseIf clsCommon.CompairString(obj.Item_Type, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Item_Type, "O") = CompairStringResult.Equal Then
    ''                    objInventoryMovemnt.ItemType = "OT"
    ''                ElseIf clsCommon.CompairString(obj.Item_Type, "F") = CompairStringResult.Equal Then
    ''                    objInventoryMovemnt.ItemType = "FT"
    ''                End If
    ''                objInventoryMovemnt.ItemType = strItemTypeToSave
    ''                ArrInventoryMovement.Add(objInventoryMovemnt)

    ''                ''Handling Leak Qty by Cahnge it's code in Father Code
    ''                If objTr.Leak_Qty > 0 And clsCommon.myLen(objTr.Fater_Code) > 0 Then
    ''                    qry = "select TSPL_ITEM_MASTER.Purchase_Class_Code,TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_PURCHASE_ACCOUNTS.Inv_Payable_Clearing from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objTr.Fater_Code + "'"
    ''                    Dim dtLeak As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
    ''                    If dtLeak Is Nothing OrElse dtLeak.Rows.Count <= 0 Then
    ''                        Throw New Exception("Please set Purchase Account set for item " + objTr.Fater_Code)
    ''                    End If

    ''                    strInvCtrlAC = clsCommon.myCstr(dtLeak.Rows(0)("Inv_Control_Account"))
    ''                    If clsCommon.myLen(strInvCtrlAC) <= 0 Then
    ''                        Throw New Exception("Please set Inventory Control Account for Purchase Account set Code :" + clsCommon.myCstr(dtLeak.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Fater_Code)
    ''                    End If
    ''                    strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, obj.Bill_To_Location, trans)
    ''                    Dim AccDrt() As String = {strInvCtrlAC, objTr.Fater_Amt}
    ''                    ArryLstGLAC.Add(AccDrt)

    ''                    strPaybleCleanigCtrlAC = clsCommon.myCstr(dtLeak.Rows(0)("Inv_Payable_Clearing"))
    ''                    If clsCommon.myLen(strInvCtrlAC) <= 0 Then
    ''                        Throw New Exception("Please set Payable Clearing Account for Purchase Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Purchase_Class_Code")) + " and Item: " + objTr.Fater_Code)
    ''                    End If
    ''                    strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strPaybleCleanigCtrlAC, obj.Bill_To_Location, trans)
    ''                    Dim AccCrt() As String = {strPaybleCleanigCtrlAC, -1 * (objTr.Fater_Amt)}
    ''                    ArryLstGLAC.Add(AccCrt)



    ''                    Dim strFatherName = clsItemMaster.GetItemName(objTr.Fater_Code, trans)
    ''                    objLocationDetails = New clsItemLocationDetails()
    ''                    objLocationDetails.Item_Code = objTr.Fater_Code
    ''                    objLocationDetails.Item_Desc = strFatherName
    ''                    objLocationDetails.Location_Code = objTr.Location
    ''                    objLocationDetails.Location_Desc = objTr.LocationName
    ''                    objLocationDetails.Item_Qty = objTr.Leak_Qty
    ''                    objLocationDetails.Amount = objTr.Item_Cost * objTr.Leak_Qty
    ''                    objLocationDetails.MRP = objTr.Fater_Rate
    ''                    objLocationDetails.ItemType = "E"
    ''                    ArrLocationDetails.Add(objLocationDetails)

    ''                    objInventoryMovemnt = New clsInventoryMovement()
    ''                    objInventoryMovemnt.InOut = "I"
    ''                    objInventoryMovemnt.Location_Code = objTr.Location
    ''                    objInventoryMovemnt.Item_Code = objTr.Fater_Code
    ''                    objInventoryMovemnt.Item_Desc = strFatherName
    ''                    objInventoryMovemnt.Qty = objTr.Leak_Qty
    ''                    objInventoryMovemnt.UOM = objTr.Unit_code
    ''                    objInventoryMovemnt.Basic_Cost = objTr.Item_Cost * objTr.Leak_Qty
    ''                    objInventoryMovemnt.Add_Cost = 0
    ''                    objInventoryMovemnt.Net_Cost = objTr.Item_Cost * objTr.Leak_Qty
    ''                    objInventoryMovemnt.ItemType = "E"
    ''                    ArrInventoryMovement.Add(objInventoryMovemnt)
    ''                End If
    ''            End If
    ''        Next

    ''        If clsCommon.myLen(strRgpNo) <= 0 Or strRgpNo = Nothing Then
    ''            clsJournalMaster.FunGrnlEntryWithTrans(trans, strPostDate, "Against Store Received Note " + obj.SRN_No, "PO-RC", "Store Received Note", obj.SRN_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstGLAC)
    ''        End If
    ''        isSaved = isSaved AndAlso clsItemLocationDetails.SaveData(strPostDate, ArrLocationDetails, trans)
    ''        isSaved = isSaved AndAlso clsInventoryMovement.SaveData("SRN", obj.SRN_No, obj.SRN_Date, strPostDate, ArrInventoryMovement, trans)

    ''        Dim strRMDANo As String = ""
    ''        If IsRejectedItemFound Then
    ''            strRMDANo = clsERPFuncationality.GetNextCode(trans, obj.SRN_Date, clsDocType.RMDA, "", obj.Bill_To_Location)
    ''            If clsCommon.myLen(strRMDANo) < 0 Then
    ''                Throw New Exception("Error in Code Generation for Rejected Material Dispatch Advice")
    ''            End If
    ''        End If
    ''        qry = "Update TSPL_SRN_HEAD set Status=1, Posting_Date='" + strPostDate + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"

    ''        If IsRejectedItemFound Then
    ''            qry += ",RMDA_No='" + strRMDANo + "',RMDA_Date='" + strPostDate + "'"
    ''        End If
    ''        qry += " where SRN_No='" + strDocNo + "'"
    ''        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

    ''        If isSaved Then
    ''            trans.Commit()
    ''        Else
    ''            trans.Rollback()
    ''        End If
    ''    Catch ex As Exception
    ''        trans.Rollback()
    ''        Throw New Exception(ex.Message)
    ''    End Try
    ''    Return True
    ''End Function

    Private Shared Function GetFirstItemCode(ByVal Arr As List(Of clsSRNDetail)) As String
        For Each objtr As clsSRNDetail In Arr
            If clsCommon.CompairString(objtr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                Return objtr.Item_Code
            End If
        Next
        Return ""
    End Function
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
    Public Shared Function DeleteData(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        If (clsCommon.myLen(strCode) <= 0) Then
            Throw New Exception("Purchase Order No not found to Delete")
        End If
        Dim obj As clsSRNHead = clsSRNHead.GetData(strCode, NavigatorType.Current, trans)

        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.SRN_No) > 0) Then
            Try
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePurchase, clsUserMgtCode.mbtnSRN, IIf(clsCommon.myLen(obj.Ship_To_Location) <= 0, obj.Bill_To_Location, obj.Ship_To_Location), obj.SRN_Date, trans)
                'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase", "Store receipt Note", obj.Bill_To_Location, obj.SRN_Date, trans)
                If (obj.Status = 1) Then
                    Throw New Exception("Already Posted on :" + obj.Posting_Date)
                End If
                clsSRNAdditionChargeInsurance.DeleteData(strCode, trans)
                clsSerializeInvenotry.DeleteData("SRN", strCode, trans)
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strCode, "TSPL_SRN_HEAD", "SRN_No", "TSPL_SRN_DETAIL", "SRN_No", trans)

                HistoryUpdate(strCode, trans)
                Dim qry As String = "delete from TSPL_SRN_DETAIL where SRN_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                clsCustomFieldValues.DeleteData(obj.Form_ID, strCode, trans)

                qry = "update TSPL_CFORM_ISSUE_RECEIVE_DETAIL set srn_no='' where srn_no='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "update TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL set srn_no='' where srn_no='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_RGP_BOM_DETAIL where srn_no='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "update TSPL_QC_CHECK_MRN_DETAIL set TSPL_QC_CHECK_MRN_DETAIL.SRN_ID=NULL where document_code='" + obj.Against_QC_Code + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_QC_CHECK_APPROVAL_ENTRY where document_code='" + obj.Against_QC_Code + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "update TSPL_QC_CHECK_HEAD set TSPL_QC_CHECK_HEAD.approved_for_srn=0 where document_code='" + obj.Against_QC_Code + "'" ''done by monika 31/01/2017
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SRN_DEDUCTION_SECURITY where SRN_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SRN_DEDUCTION where SRN_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SRN_TENDER_CALC where SRN_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SRN_HEAD where SRN_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)


            Catch ex As Exception

                Throw New Exception(ex.Message)
            End Try
        End If
        Return True
    End Function

    Public Shared Function IsValidVendorForSRN(ByVal Arr As List(Of String), ByVal strVendorCode As String) As Boolean
        If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
            Dim qry As String = "select SRN_No,Vendor_Code,Vendor_Name from TSPL_SRN_HEAD where SRN_No in (" + clsCommon.GetMulcallString(Arr) + ") and Vendor_Code not in ('" + strVendorCode + "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim msg As String = "Error. "
                For Each dr As DataRow In dt.Rows
                    msg += Environment.NewLine + "SRN No:" + clsCommon.myCstr(dr("SRN_No")) + " Is For Vendor Code: " + clsCommon.myCstr(dr("Vendor_Code")) + " Vendor Name:" + clsCommon.myCstr(dr("Vendor_Name"))
                Next
                Throw New Exception(msg)
            End If
        End If
        Return True
    End Function

    Public Shared Function SetItemWiseTax(ByVal dtAfterModify As DataTable, ByVal strSRNNo As String) As DataTable
        dtAfterModify.Columns.Add("TAX1_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX2_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX3_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX4_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX5_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt3", GetType(Double))




        Dim qry As String = "select Tax,Rate,SUM(Amt) as TaxAmt"
        qry += " from ("
        qry += " select TAX1 as Tax,TAX1_Rate as Rate,TAX1_Amt as Amt"
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "' "
        qry += " union all "
        qry += " select TAX2 as Tax,TAX2_Rate as Rate,TAX2_Amt as Amt "
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "'  "
        qry += " union all "
        qry += " select TAX3 as Tax,TAX3_Rate as Rate,TAX3_Amt as Amt "
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "'  "
        qry += " union all "
        qry += " select TAX4 as Tax,TAX4_Rate as Rate,TAX4_Amt as Amt "
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "'  "
        qry += " union all "
        qry += " select TAX5 as Tax,TAX5_Rate as Rate,TAX5_Amt as Amt "
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "'   "
        qry += " union all "
        qry += " select TAX6 as Tax,TAX6_Rate as Rate,TAX6_Amt as Amt "
        qry += " from TSPL_SRN_DETAIL where SRN_No='" + strSRNNo + "'   "
        qry += " )xxx "
        qry += " group by Tax,Rate   having SUM(Amt)>0   "


        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                For ii As Integer = 1 To 5
                    Dim strCol As String = "TAX" + clsCommon.myCstr(ii) + ""
                    If clsCommon.CompairString(clsCommon.myCstr(dtAfterModify.Rows(0)(strCol)), clsCommon.myCstr(dr("Tax"))) = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt1")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate1") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt1") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt2")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate2") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt2") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt3")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate3") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt3") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        Else
                            Throw New Exception("Printing Support only 3 Diffent Rates")
                        End If

                    End If
                Next
            Next
        End If
        Return dtAfterModify
    End Function
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
    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim obj As clsSRNHead = clsSRNHead.GetData(strCode, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.SRN_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If Not (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If
            For Each objtr As clsSRNDetail In obj.Arr
                ''Batch Item
                If objtr.arrBatchItem IsNot Nothing AndAlso objtr.arrBatchItem.Count > 0 Then
                    For Each objBatch As clsBatchInventory In objtr.arrBatchItem
                        Dim dblBalance As Decimal = clsBatchInventory.GetBatchBalance(objBatch.Item_Code, objBatch.Location_Code, objBatch.Batch_No, objBatch.MRP, objBatch.UOM, objBatch.Document_Code, objBatch.Document_Type, trans)
                        If dblBalance < objBatch.Qty Then
                            Throw New Exception("Balance will be going to -ve.Balance Qty : " + clsCommon.myCstr(dblBalance) + " and Entered Qty : " + clsCommon.myCstr(objBatch.Qty) + Environment.NewLine + "Item : " + objBatch.Item_Code + Environment.NewLine + "Batch : " + objBatch.Batch_No + Environment.NewLine + "MRP : " + clsCommon.myCstr(objBatch.MRP) + Environment.NewLine + "Unit : " + objBatch.UOM)
                        End If
                    Next
                End If
            Next

            Dim dts As DataTable = clsDBFuncationality.GetDataTable(" select TSPL_SRN_HEAD.Bill_To_Location, TSPL_SRN_HEAD.SRN_No, TSPL_SRN_HEAD.SRN_Date from TSPL_SRN_HEAD where SRN_No= '" + strCode + "' ", trans)
            If dts.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePurchase, clsUserMgtCode.mbtnSRN, clsCommon.myCstr(dts.Rows(0)("Bill_To_Location")), clsCommon.myCstr(dts.Rows(0)("SRN_Date")), trans)
            End If

            Dim Qry As String = "select Status,Confirmatory_PO from TSPL_SRN_HEAD where SRN_No='" + strCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("Document No [" + strCode + "] not found for reverse and unpost")
            End If

            If Not clsCommon.myCdbl(dt.Rows(0)("Status")) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If
            Dim strReturnExist As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from TSPL_SRN_RETURN where SRN_No = '" + strCode + "' ", trans))
            If clsCommon.myLen(strReturnExist) > 0 Then
                Throw New Exception("SRN Return Document  (" + strReturnExist + ") Exist Against this SRN. reverse and unpost Not Possible")
            End If
            If clsCommon.myCdbl(dt.Rows(0)("Confirmatory_PO")) = 1 Then
                Qry = "select PurchaseOrder_No from TSPL_PURCHASE_ORDER_HEAD  where Confirmatory_PO_SRN_No='" + strCode + "'"
                dt = clsDBFuncationality.GetDataTable(Qry, trans)

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    clsPurchaseOrderHead.ReverseAndUnpost(clsCommon.myCstr(dt.Rows(0)("PurchaseOrder_No")), clsUserMgtCode.mbtnPurchaseOrder, trans)
                    clsPurchaseOrderHead.DeleteData(clsCommon.myCstr(dt.Rows(0)("PurchaseOrder_No")), False, trans)
                End If
            End If


            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PurchaseDoNotCheckForwardDocuments, clsFixedParameterCode.PurchaseDoNotCheckForwardDocuments, trans)) <= 0 Then
                Qry = "select distinct PI_No from TSPL_PI_DETAIL where SRN_Id='" + strCode + "'"
                dt = clsDBFuncationality.GetDataTable(Qry, trans)

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Qry = "CURRENT SRN IS USED IN FOLLOWING PURCHASE INVOICE -"
                    For Each DR As DataRow In dt.Rows
                        Qry += Environment.NewLine + clsCommon.myCstr(DR("PI_NO"))
                    Next
                    Throw New Exception(Qry)
                End If
            End If

            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='PO-RC' and Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            Qry = "select Document_No from TSPL_TENDER_PENALTY_DETAIL where SRN_No='" + strCode + "'"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("CURRENT SRN IS USED IN FOLLOWING Tender Dedcution -" + clsCommon.myCstr(dt.Rows(0)("Document_No")))
            End If


            Qry = "update TSPL_SERIAL_ITEM set Against_Inv_Movement_Trans_Id=null where Document_Code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)


            If objCommonVar.AutoGenrateBatchInventory Then
                clsBatchInventory.DeleteData("SRN", strCode, trans)
            Else
                clsBatchInventory.ReverseAndUnpost("SRN", strCode, trans)
            End If

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)
            Qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" + strCode + "' and Trans_Type='SRN'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)




            Qry = "Update TSPL_SRN_HEAD set Status = 0 where SRN_No='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_SRN_HEAD", "SRN_No", "TSPL_SRN_detail", "SRN_No", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GenerateSRNDeduction(ByVal strSRNNo As String, ByVal strICode As String, ByVal ApplyQCDed As Boolean, ByVal ApplySecDed As Boolean, ByVal ApplyRALPenalty As Boolean, ByVal trans As SqlTransaction) As Boolean
        Return GenerateSRNDeduction(strSRNNo, strICode, ApplyQCDed, ApplySecDed, ApplyRALPenalty, trans, 0, 0)
    End Function


    Public Shared Function GenerateSRNDeduction(ByVal strSRNNo As String, ByVal strICode As String, ByVal ApplyQCDed As Boolean, ByVal ApplySecDed As Boolean, ByVal ApplyRALPenalty As Boolean, ByVal trans As SqlTransaction, ByVal isSkipSecurityDed As Integer, ByVal isSkipPenaltyDed As Integer) As Boolean
        Dim qry1 As String = Nothing
        Dim qry As String = "select TSPL_QC_CHECK_HEAD.Document_Code as Against_QC_Code,TSPL_SRN_DETAIL.PO_ID,TSPL_SRN_DETAIL.Row_Type,TSPL_SRN_DETAIL.SRN_Qty,TSPL_SRN_DETAIL.Leak_Qty,TSPL_SRN_DETAIL.Burst_Qty,TSPL_SRN_DETAIL.Short_Qty,TSPL_SRN_HEAD.Vendor_Code,TSPL_SRN_HEAD.isExemptSecurityDedution ,TSPL_SRN_DETAIL.GRN_ID,TSPL_SRN_DETAIL.Item_Net_Amt,TSPL_TENDER_HEADER.Tender_Type 
from TSPL_SRN_DETAIL
left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No
left outer join TSPL_QC_CHECK_HEAD on TSPL_QC_CHECK_HEAD.Gate_Entry_No=TSPL_SRN_HEAD.Against_GRN	
left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_SRN_DETAIL.PO_ID
left outer join TSPL_TENDER_HEADER on TSPL_TENDER_HEADER.DocumentCode=TSPL_PURCHASE_ORDER_HEAD.RefTendorNo
where TSPL_SRN_HEAD.SRN_No='" + strSRNNo + "' and TSPL_SRN_DETAIL.Item_Code='" + strICode + "'"
        Dim dtSRN As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dtSRN IsNot Nothing AndAlso dtSRN.Rows.Count > 0 Then
#Region "SRN Dedution"
            If ApplyQCDed Then
                If clsCommon.myLen(dtSRN.Rows(0)("Against_QC_Code")) > 0 Then
                    qry = "insert into TSPL_SRN_DEDUCTION (SRN_No,Item_Code,Amt,Ded_Per,Ded_Amt)
select SRN_No,Item_Code,Amount,InputDataDeductionPer,(Amount*InputDataDeductionPer/100) as DedAmt  from (
select SRN_No,Item_Code,max(Amount) as Amount,sum(isnull(InputDataDeductionPer,0)) as InputDataDeductionPer from (
select TSPL_SRN_HEAD.SRN_No,TSPL_QC_CHECK_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.Item_Net_Amt as Amount,TSPL_QC_CHECK_SRN_DETAIL.InputDataDeductionPer 
from TSPL_SRN_DETAIL 
left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No
left outer join TSPL_QC_CHECK_HEAD on TSPL_QC_CHECK_HEAD.Gate_Entry_No=TSPL_SRN_HEAD.Against_GRN	
left outer join TSPL_QC_CHECK_SRN_DETAIL on TSPL_QC_CHECK_SRN_DETAIL.document_code=TSPL_QC_CHECK_HEAD.document_code and TSPL_QC_CHECK_SRN_DETAIL.Item_Code=TSPL_SRN_DETAIL.Item_Code
where TSPL_SRN_HEAD.SRN_No='" + strSRNNo + "' and TSPL_SRN_DETAIL.Item_Code='" + strICode + "'
)x group by SRN_No,Item_Code
)xx where (Amount*InputDataDeductionPer/100)>0 "
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If
            End If
#End Region

#Region "Security Dedution"
            If isSkipSecurityDed <= 0 Then
                If ApplySecDed Then
                    If clsCommon.myCDecimal(dtSRN.Rows(0)("isExemptSecurityDedution")) = 0 Then
                        qry = "insert into TSPL_SRN_DEDUCTION_SECURITY (SRN_No,Item_Code,Amt,Ded_Per,Ded_Amt)
select SRN_No,Item_Code,Amount,Security_Deduction,round((Amount*Security_Deduction/100),0) as DedAmt  from (
select SRN_No,Item_Code,max(Amount) as Amount,sum(isnull(Security_Deduction,0)) as Security_Deduction from (
select TSPL_SRN_DETAIL.SRN_No,TSPL_SRN_DETAIL.Item_Code,TSPL_SRN_DETAIL.Item_Net_Amt as Amount,isnull(TSPL_ITEM_MASTER.Security_Deduction,0) as Security_Deduction from TSPL_SRN_DETAIL 
inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.item_Code=TSPL_SRN_DETAIL.Item_Code
where TSPL_SRN_DETAIL.SRN_No='" + strSRNNo + "' and TSPL_SRN_DETAIL.Item_Code='" + strICode + "' and isnull(TSPL_ITEM_MASTER.Security_Deduction,0)>0
)x group by SRN_No,Item_Code
)xx where (Amount*Security_Deduction/100)>0 "
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    End If
                End If
            End If
#End Region

#Region "Apply Tender Penalty"
            If isSkipPenaltyDed <= 0 Then
                If ApplyRALPenalty Then
                    If clsCommon.myLen(clsCommon.myCstr(dtSRN.Rows(0)("PO_ID"))) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(dtSRN.Rows(0)("Row_Type")), clsItemRowType.RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCDecimal(dtSRN.Rows(0)("SRN_Qty")) > 0 Then
                        qry = "select GRN_Date from TSPL_GRN_HEAD where GRN_No='" + clsCommon.myCstr(dtSRN.Rows(0)("GRN_ID")) + "'"
                        Dim GRNDate As Date = clsCommon.myCDate(clsDBFuncationality.getSingleValue(qry, trans))
                        Dim dclSRNQty As Decimal = clsCommon.myCDecimal(dtSRN.Rows(0)("SRN_Qty"))

                        If clsCommon.myCDecimal(dtSRN.Rows(0)("Tender_Type")) = 2 OrElse clsCommon.myCDecimal(dtSRN.Rows(0)("Tender_Type")) = 3 Then
                            qry = "select DocumentCode,PK_Id,max(To_Date) as To_Date,Item_Code,sum(Qty*RI) as Qty,max(schedule_no) as  schedule_no  from (
select TSPL_TENDER_SCHEDULE_PO.DocumentCode,TSPL_TENDER_SCHEDULE_PO.PK_Id
,DATEADD(day,isnull(TSPL_TENDER_SCHEDULE_PO.Extension_Days,0),TSPL_TENDER_SCHEDULE_PO.To_Date) as To_Date
,TSPL_TENDER_SCHEDULE_PO.Item_Code,TSPL_TENDER_SCHEDULE_PO.Schedule_Qty as Qty,1 AS RI ,1 as Chk ,TSPL_TENDER_SCHEDULE_PO.schedule_no
from TSPL_PURCHASE_ORDER_HEAD 
inner join TSPL_TENDER_SCHEDULE_PO on TSPL_TENDER_SCHEDULE_PO.DocumentCode=TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No  
where  TSPL_PURCHASE_ORDER_HEAD.Against_Tender='Y' and TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No='" + clsCommon.myCstr(dtSRN.Rows(0)("PO_ID")) + "' and TSPL_TENDER_SCHEDULE_PO.Item_Code='" + strICode + "'
union all
select Against_PO as DocumentCode,Against_Tender_Schedule_PK_Id_PO as PK_Id,null as To_Date,Item_Code ,Qty,-1 as RI,0 as chk, 0 as schedule_no from TSPL_SRN_TENDER_CALC
)xx group by DocumentCode,PK_Id,Item_Code having sum(Chk)>0 order by PK_Id"
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                For kk As Integer = 0 To dt.Rows.Count - 1
                                    Dim dclBalanceQty As Decimal = clsCommon.myCDecimal(dt.Rows(kk)("Qty"))
                                    Dim isLastSchedule As Boolean = False
                                    If kk + 1 = dt.Rows.Count Then
                                        isLastSchedule = True
                                    End If
                                    Dim coll As New Hashtable()
                                    clsCommon.AddColumnsForChange(coll, "Against_PO", clsCommon.myCstr(dt.Rows(kk)("DocumentCode")))
                                    clsCommon.AddColumnsForChange(coll, "Against_Tender_Schedule_PK_Id_PO", clsCommon.myCDecimal(dt.Rows(kk)("PK_Id")))
                                    clsCommon.AddColumnsForChange(coll, "SRN_No", strSRNNo)
                                    clsCommon.AddColumnsForChange(coll, "Item_Code", strICode)
                                    Dim dclApplyQty As Decimal = 0
                                    If isLastSchedule Then
                                        If dclSRNQty <= dclBalanceQty Then
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        Else
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        End If
                                    Else
                                        If dclBalanceQty <= 0 Then
                                            Continue For
                                        End If
                                        If dclSRNQty <= dclBalanceQty Then
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        Else
                                            dclApplyQty = dclBalanceQty
                                            dclSRNQty = dclSRNQty - dclApplyQty
                                        End If
                                    End If
                                    clsCommon.AddColumnsForChange(coll, "Qty", dclApplyQty)
                                    If clsCommon.GetDateWithStartTime(GRNDate) > clsCommon.GetDateWithStartTime(clsCommon.myCDate(dt.Rows(kk)("To_Date"))) Then
                                        Dim isPenaltyApply As Boolean = False
                                        Dim ArrPenalty As List(Of clsTenderSchedulePeneltyPO) = clsTenderSchedulePeneltyPO.GetData(clsCommon.myCDecimal(dt.Rows(kk)("PK_Id")), True, trans)
                                        If ArrPenalty IsNot Nothing AndAlso ArrPenalty.Count > 0 Then
                                            For ll As Integer = 0 To ArrPenalty.Count - 1
                                                If ll = ArrPenalty.Count - 1 OrElse
                                                   clsCommon.GetDateWithStartTime(GRNDate) <= clsCommon.GetDateWithStartTime(ArrPenalty(ll).Penalty_Date) Then
                                                    isPenaltyApply = True
                                                    clsCommon.AddColumnsForChange(coll, "Against_Tender_Schedule_Penalty_PK_Id_PO", ArrPenalty(ll).PK_Id)
                                                    clsCommon.AddColumnsForChange(coll, "Penalty", Math.Round(((dclApplyQty * clsCommon.myCDecimal(clsCommon.myCDivide(clsCommon.myCDecimal(dtSRN.Rows(0)("Item_Net_Amt")), (clsCommon.myCDecimal(dtSRN.Rows(0)("SRN_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Leak_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Burst_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Short_Qty"))))) * ArrPenalty(ll).Penalty) / 100), 2, MidpointRounding.AwayFromZero))
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        If Not isPenaltyApply Then
                                            If kk = dt.Rows.Count - 1 Then
                                                Throw New Exception("Tender [" + clsCommon.myCstr(dt.Rows(kk)("DocumentCode")) + "] PO [" + clsCommon.myCstr(dtSRN.Rows(0)("PO_ID")) + "] Item [" + strICode + "] Exeed the Last Date.Can't Accept it")
                                            End If
                                        End If
                                    End If
                                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_TENDER_CALC", OMInsertOrUpdate.Insert, "", trans)
                                    If dclSRNQty <= 0 Then
                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            qry = "select DocumentCode,PK_Id,max(To_Date) as To_Date,Item_Code,sum(Qty*RI) as Qty,MAX(Schedule_No) AS Schedule_No  from (
select TSPL_TENDER_SCHEDULE.DocumentCode,TSPL_TENDER_SCHEDULE.PK_Id
,DATEADD(day,isnull(TSPL_TENDER_SCHEDULE.Extension_Days,0),TSPL_TENDER_SCHEDULE.To_Date) as To_Date
,TSPL_TENDER_DETAIL.Item_Code,TSPL_TENDER_SCHEDULE.Schedule_Qty as Qty,1 AS RI ,1 as Chk ,TSPL_TENDER_SCHEDULE.Schedule_No as Schedule_No
from TSPL_PURCHASE_ORDER_HEAD 
inner join TSPL_TENDER_DETAIL on TSPL_TENDER_DETAIL.DocumentCode=TSPL_PURCHASE_ORDER_HEAD.RefTendorNo and TSPL_TENDER_DETAIL.Vendor_Code=TSPL_PURCHASE_ORDER_HEAD.Vendor_Code and TSPL_TENDER_DETAIL.Location=TSPL_PURCHASE_ORDER_HEAD.Bill_To_Location
inner join TSPL_TENDER_SCHEDULE on TSPL_TENDER_SCHEDULE.DocumentCode=TSPL_TENDER_DETAIL.DocumentCode and TSPL_TENDER_DETAIL.Line_No=TSPL_TENDER_SCHEDULE.PSNo
where  TSPL_PURCHASE_ORDER_HEAD.Against_Tender='Y' and TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No='" + clsCommon.myCstr(dtSRN.Rows(0)("PO_ID")) + "' and TSPL_TENDER_DETAIL.Vendor_Code='" + clsCommon.myCstr(dtSRN.Rows(0)("Vendor_Code")) + "' and TSPL_TENDER_DETAIL.Item_Code='" + strICode + "'
union all
select Against_TenderNo as DocumentCode,Against_Tender_Schedule_PK_Id as PK_Id,null as To_Date,Item_Code ,Qty,-1 as RI,0 as chk,0 as line_no from TSPL_SRN_TENDER_CALC
)xx group by DocumentCode,PK_Id,Item_Code having sum(Chk)>0 order by PK_Id"
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                For kk As Integer = 0 To dt.Rows.Count - 1
                                    Dim dclBalanceQty As Decimal = clsCommon.myCDecimal(dt.Rows(kk)("Qty"))
                                    Dim isLastSchedule As Boolean = False
                                    If kk + 1 = dt.Rows.Count Then
                                        isLastSchedule = True
                                    End If

                                    Dim coll As New Hashtable()
                                    clsCommon.AddColumnsForChange(coll, "Against_TenderNo", clsCommon.myCstr(dt.Rows(kk)("DocumentCode")))
                                    clsCommon.AddColumnsForChange(coll, "Against_Tender_Schedule_PK_Id", clsCommon.myCDecimal(dt.Rows(kk)("PK_Id")))
                                    clsCommon.AddColumnsForChange(coll, "SRN_No", strSRNNo)
                                    clsCommon.AddColumnsForChange(coll, "Item_Code", strICode)
                                    Dim dclApplyQty As Decimal = 0
                                    If isLastSchedule Then
                                        If dclSRNQty <= dclBalanceQty Then
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        Else
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        End If
                                    Else
                                        If dclBalanceQty <= 0 Then
                                            Continue For
                                        End If
                                        If dclSRNQty <= dclBalanceQty Then
                                            dclApplyQty = dclSRNQty
                                            dclSRNQty = 0
                                        Else
                                            dclApplyQty = dclBalanceQty
                                            dclSRNQty = dclSRNQty - dclApplyQty
                                        End If
                                    End If
                                    clsCommon.AddColumnsForChange(coll, "Qty", dclApplyQty)
                                    If clsCommon.GetDateWithStartTime(GRNDate) > clsCommon.GetDateWithStartTime(clsCommon.myCDate(dt.Rows(kk)("To_Date"))) Then
                                        Dim isPenaltyApply As Boolean = False
                                        Dim ArrPenalty As List(Of clsTenderSchedulePenelty) = clsTenderSchedulePenelty.GetData(clsCommon.myCDecimal(dt.Rows(kk)("PK_Id")), True, trans)
                                        If ArrPenalty IsNot Nothing AndAlso ArrPenalty.Count > 0 Then
                                            For ll As Integer = 0 To ArrPenalty.Count - 1
                                                If ll = ArrPenalty.Count - 1 OrElse
                                                   clsCommon.GetDateWithStartTime(GRNDate) <= clsCommon.GetDateWithStartTime(ArrPenalty(ll).Penalty_Date) Then
                                                    isPenaltyApply = True
                                                    clsCommon.AddColumnsForChange(coll, "Against_Tender_Schedule_Penalty_PK_Id", ArrPenalty(ll).PK_Id)
                                                    clsCommon.AddColumnsForChange(coll, "Penalty", Math.Round(((dclApplyQty * clsCommon.myCDecimal(clsCommon.myCDivide(clsCommon.myCDecimal(dtSRN.Rows(0)("Item_Net_Amt")), (clsCommon.myCDecimal(dtSRN.Rows(0)("SRN_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Leak_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Burst_Qty")) + clsCommon.myCDecimal(dtSRN.Rows(0)("Short_Qty"))))) * ArrPenalty(ll).Penalty) / 100), 2, MidpointRounding.AwayFromZero))
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        If Not isPenaltyApply Then
                                            If kk = dt.Rows.Count - 1 Then
                                                Throw New Exception("Tender [" + clsCommon.myCstr(dt.Rows(kk)("DocumentCode")) + "] Item [" + strICode + "] Exeed the Last Date.Can't Accept it")
                                            End If
                                        End If
                                    End If
                                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_TENDER_CALC", OMInsertOrUpdate.Insert, "", trans)
                                    If dclSRNQty <= 0 Then
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
#End Region
        End If
            Return True
    End Function

    Public Shared Function GetBaseQeryTenderPenalty(ByVal strTender As String, ByVal strItem As String, ByVal strVendor As String, ByVal strLocation As String, ByVal UserStaus As String, ByVal WhrCls As String) As String
        Return GetBaseQeryTenderPenalty(strTender, strItem, strVendor, strLocation, UserStaus, WhrCls, 0, 0)
    End Function

    Public Shared Function GetBaseQeryTenderPenalty(ByVal strTender As String, ByVal strItem As String, ByVal strVendor As String, ByVal strLocation As String, ByVal UserStaus As String, ByVal WhrCls As String, ByVal isSkipSecurityDed As Integer, ByVal isSkipPenaltyDed As Integer) As String
        Dim qry As String = "with CTERawData as (
select  cast(" + UserStaus + " as bit) as UserStatus, TSPL_GRN_HEAD.GRN_No,convert(varchar, TSPL_GRN_HEAD.GRN_Date,103) as GRN_Date,TSPL_GRN_HEAD.VehicleNo,isnull(TSPL_GRN_HEAD.Status,0) as GRNStatus,TSPL_SRN_HEAD.SRN_No,convert(varchar,TSPL_SRN_HEAD.SRN_Date,103) as  SRN_Date,isnull(TSPL_SRN_HEAD.Status,0) as SRNStatus, TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code,convert(varchar,TSPL_PO_WEIGHTMENT_HEAD.Weighment_Date,103) as Weighment_Date,TSPL_PO_WEIGHTMENT_DETAIL.Gross_Weight,TSPL_PO_WEIGHTMENT_DETAIL.Tare_Weight,TSPL_PO_WEIGHTMENT_DETAIL.Extra_Weight,TSPL_PO_WEIGHTMENT_DETAIL.UOM,TSPL_PO_WEIGHTMENT_DETAIL.Net_Weight,isnull(TSPL_PO_WEIGHTMENT_HEAD.Status,0) as WeightmentStatus,TSPL_SRN_DETAIL.SRN_Qty , TSPL_SRN_DEDUCTION.Ded_Per As QualityDeductionPer, TSPL_SRN_DEDUCTION.Ded_Amt As QualityDeductionAmt,"
        If isSkipSecurityDed > 0 Then
            qry += " 0 As SecurityDeductionAmt,"
        Else
            qry += "TSPL_SRN_DEDUCTION_SECURITY.Ded_Amt as SecurityDeductionAmt,"
        End If
        If isSkipPenaltyDed > 0 Then
            qry += " 0 As LatePenaltyQty,0 As LatePenaltyPer, 0 As LatePenaltyAmt,"
        Else
            qry += "case when isnull(TSPL_SRN_TENDER_CALC.Penalty, 0)=0 then null else TSPL_SRN_TENDER_CALC.Qty end as LatePenaltyQty,case when isnull(TSPL_SRN_TENDER_CALC.Penalty,0)=0 then null else TSPL_TENDER_SCHEDULE_PENALTY.Penalty end as LatePenaltyPer,TSPL_SRN_TENDER_CALC.Penalty as LatePenaltyAmt,"
        End If
        qry += "(case when isnull(TSPL_MRN_HEAD.NIR_QC,0)=0 then 1 else (case when (isnull(TSPL_NIR_QC.QC_Status,0)=1 And isnull(TSPL_NIR_QC.Status,0)=1 And TSPL_QC_CHECK_HEAD.Posted=1) then 1 else 0 end) end) as NIRQCStatus
," + UserStaus + " as FinalStatus, TSPL_GRN_HEAD.GRN_Date as FilterGRN_Date,isnull(TSPL_SRN_HEAD.Status,0) as FilterSRNStatus
from TSPL_GRN_DETAIL
left outer join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=TSPL_GRN_DETAIL.GRN_No
left outer join TSPL_MRN_HEAD on TSPL_MRN_HEAD.Against_GRN=TSPL_GRN_DETAIL.GRN_No
left outer join TSPL_NIR_QC on TSPL_NIR_QC.MRN_No=TSPL_MRN_HEAD.MRN_No
left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_GRN_DETAIL.PO_Id
left outer join TSPL_SRN_DETAIL on TSPL_SRN_DETAIL.GRN_ID=TSPL_GRN_HEAD.GRN_No and TSPL_SRN_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code
left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No
left outer join TSPL_PO_WEIGHTMENT_HEAD on TSPL_PO_WEIGHTMENT_HEAD.Against_GRN_No=TSPL_GRN_HEAD.GRN_No
left outer join TSPL_PO_WEIGHTMENT_DETAIL on TSPL_PO_WEIGHTMENT_DETAIL.Weighment_Code= TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code and  TSPL_PO_WEIGHTMENT_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code
left outer join TSPL_SRN_DEDUCTION_SECURITY on TSPL_SRN_DEDUCTION_SECURITY.SRN_No=TSPL_SRN_HEAD.SRN_No and TSPL_SRN_DEDUCTION_SECURITY.Item_Code=TSPL_SRN_DETAIL.Item_Code
left outer join TSPL_SRN_DEDUCTION on TSPL_SRN_DEDUCTION.SRN_No=TSPL_SRN_HEAD.SRN_No and TSPL_SRN_DEDUCTION.Item_Code=TSPL_SRN_DETAIL.Item_Code
left outer join TSPL_SRN_TENDER_CALC on TSPL_SRN_TENDER_CALC.SRN_No=TSPL_SRN_HEAD.SRN_No and TSPL_SRN_TENDER_CALC.Item_Code=TSPL_SRN_DETAIL.Item_Code and isnull(TSPL_SRN_TENDER_CALC.Penalty,0)>0
left outer join TSPL_TENDER_SCHEDULE_PENALTY on  TSPL_TENDER_SCHEDULE_PENALTY.PK_Id=TSPL_SRN_TENDER_CALC.Against_Tender_Schedule_Penalty_PK_Id
left outer join TSPL_TENDER_HEADER on TSPL_TENDER_HEADER.DocumentCode=TSPL_PURCHASE_ORDER_HEAD.RefTendorNo
left outer join TSPL_QC_CHECK_HEAD on TSPL_QC_CHECK_HEAD.Gate_Entry_No=TSPL_GRN_HEAD.GRN_No
where TSPL_PURCHASE_ORDER_HEAD.Against_Tender='Y' and TSPL_PURCHASE_ORDER_HEAD.RefTendorNo='" + strTender + "' and  isnull(TSPL_QC_CHECK_HEAD.QC_Status,'')<>'Rejected'  and TSPL_GRN_DETAIL.Item_Code='" + strItem + "' and TSPL_GRN_HEAD.Vendor_Code='" + strVendor + "' and TSPL_GRN_HEAD.Bill_To_Location='" + strLocation + "' and ISNULL( TSPL_GRN_HEAD.IsCancel,0)=0  
and 2= (case when isnull(TSPL_MRN_HEAD.NIR_QC,0)=1 then (case when isnull(TSPL_NIR_QC.QC_Status,0)=1 then 2 else 3 end) else 2 end) " + WhrCls + "
)
select CTERawData.* from CTERawData Order by FilterGRN_Date,  FilterSRNStatus desc"
        Return qry
    End Function
End Class

Public Class clsSRNDetail
#Region "Variables"
    Public Against_Schedule_Code As String = Nothing
    Public RGP_Qty As Double = Nothing
    Public Schedule_Qty As Double = Nothing

    Public SRN_No As String = Nothing
    Public Line_No As Integer = 0
    Public Status As Integer = 0
    Public Row_Type As String = Nothing
    Public Item_Code As String = Nothing
    Public Bar_Code As String = Nothing
    Public Item_Desc As String = Nothing
    Public SRN_Qty As Double = 0
    Public Leak_Qty As Double = 0
    Public Burst_Qty As Double = 0
    Public Short_Qty As Double = 0
    Public Rejected_Qty As Double = 0
    Public Freeqty As Double = 0
    Public MRN_Id As String = Nothing
    Public GRN_ID As String = Nothing
    Public UOMWeightValue As Double = 0
    Public UOMWeight As String = Nothing
    Public RGP_Id As String = Nothing
    Public PO_ID As String = Nothing
    Public OrgPOQty As Double = 0 'Not a Table Field
    Public OrgGRNQty As Double = 0 'Not a Table Field
    Public OrgMRNQty As Double = 0 'Not a Table Field

    Public Balance_Qty As Double = 0 '
    Public Unit_code As String = Nothing '
    Public Location As String = Nothing '
    Public LocationName As String = Nothing 'Not a Table Field
    Public SRNTax_Group As String = Nothing 'Not a Table Field
    Public Item_Cost As Double = 0 '
    Public TAX1 As String = Nothing
    Public TAX1_Base_Amt As Double = 0
    Public TAX1_Rate As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Base_Amt As Double = 0
    Public TAX2_Rate As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Base_Amt As Double = 0
    Public TAX3_Rate As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Base_Amt As Double = 0
    Public TAX4_Rate As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Base_Amt As Double = 0
    Public TAX5_Rate As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Base_Amt As Double = 0
    Public TAX6_Rate As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Base_Amt As Double = 0
    Public TAX7_Rate As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Base_Amt As Double = 0
    Public TAX8_Rate As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Base_Amt As Double = 0
    Public TAX9_Rate As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Base_Amt As Double = 0
    Public TAX10_Rate As Double = 0
    Public TAX10_Amt As Double = 0
    Public Amount As Double = 0
    Public ItemCostAmt As Double = 0
    Public Disc_Type As Integer = 0
    Public Header_Discount_Per As Decimal = 0
    Public Header_Discount_Amount As Decimal = 0
    Public Disc_Per As Double = 0
    Public Detail_Discount_Amount As Decimal = 0
    Public Disc_Per_Unit As Decimal = 0
    Public Disc_Amt_Per_Unit As Decimal = 0
    Public Disc_Amt As Double = 0
    Public Amt_Less_Discount As Double = 0
    Public Taxable_Amount_Per As Decimal = 0
    Public Taxable_Amount As Decimal = 0
    Public Total_Tax_Amt As Double = 0
    Public Item_Net_Amt As Double = 0
    Public Unit_Cost_Tax_Rate As Double = 0
    Public Unit_Cost_Tax As Double = 0
    Public MRP As Double = 0
    Public MFG_Date As Date? = Nothing
    Public Batch_No As String = Nothing
    Public Expiry_Date As Date? = Nothing

    Public Landed_Cost_Rate As Double = 0
    Public Landed_Cost_Amount As Double = 0

    Public Specification As String = Nothing
    Public Remarks As String = Nothing
    Public Is_Mannual_Amt As Integer = Nothing
    Public Assessable As Double = 0
    Public AssessableAmt As Double = 0

    Public Fater_Code As String = Nothing
    Public Fater_Rate As Double = 0
    Public Fater_Amt As Double = 0

    Public PO_Qty As Double = 0
    Public GRN_Qty As Double = 0
    Public MRN_Qty As Double = 0

    Public Total_RecTax_PerUnit As Double = 0
    Public Total_NonRecTax_PerUnit As Double = 0
    Public Total_AddtionalCost_PerUnit As Double = 0
    Public Req_No As String = Nothing
    '' for Abatement SRN
    Public AbatementRate As Decimal = 0
    Public AssessableMRP As Decimal = 0
    Public TotalAssessableMRP As Decimal = 0
    Public Bin_No As String = Nothing

    Public Accepted_Amount As Double = 0
    Public Rejected_Amount As Double = 0
    Public Shortage_Amount As Double = 0
    Public Leak_Amount As Double = 0
    Public Burst_Amount As Double = 0
    Public Amt_Less_Discount_Without_Shortage As Double = 0

    ''====================19/10/2016====================================
    Public ItemAdd_Charge_Code1 As String = Nothing
    Public ItemAdd_Org_Charge_Amt1 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt1 As Double = Nothing
    Public ItemAdd_Charge_Code9 As String = Nothing
    Public ItemAdd_Org_Charge_Amt9 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt9 As Double = Nothing
    Public ItemAdd_Charge_Code8 As String = Nothing
    Public ItemAdd_Org_Charge_Amt8 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt8 As Double = Nothing
    Public ItemAdd_Charge_Code7 As String = Nothing
    Public ItemAdd_Org_Charge_Amt7 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt7 As Double = Nothing
    Public ItemAdd_Charge_Code6 As String = Nothing
    Public ItemAdd_Org_Charge_Amt6 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt6 As Double = Nothing
    Public ItemAdd_Charge_Code5 As String = Nothing
    Public ItemAdd_Org_Charge_Amt5 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt5 As Double = Nothing
    Public ItemAdd_Charge_Code4 As String = Nothing
    Public ItemAdd_Org_Charge_Amt4 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt4 As Double = Nothing
    Public ItemAdd_Charge_Code3 As String = Nothing
    Public ItemAdd_Org_Charge_Amt3 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt3 As Double = Nothing
    Public ItemAdd_Charge_Code2 As String = Nothing
    Public ItemAdd_Org_Charge_Amt2 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt2 As Double = Nothing
    Public ItemAdd_Charge_Code10 As String = Nothing
    Public ItemAdd_Org_Charge_Amt10 As Double = Nothing
    Public ItemAdd_Calc_Charge_Amt10 As Double = Nothing
    Public Total_ItemAdd_Charge As Double = Nothing
    ''==================================================================

    Public Insurance_Base_Amt As Decimal
    Public Insurance_Per As Decimal

    Public Capex_SubCode As String = Nothing
    Public Capex_Code As String = Nothing
    Public Emergency As Boolean = Nothing
    Public Category As String = Nothing
    Public Against_Item_Wise_Tax_Rate As String = Nothing
    Public arrSrItem As List(Of clsSerializeInvenotry) = Nothing
    Public arrBatchItem As List(Of clsBatchInventory) = Nothing

    Public Item_Insurance_Base_Amt As Decimal = 0
    Public Item_Insurance_Apply_On As String = Nothing
    Public Item_Insurance_Rate As Decimal = 0
    Public Item_Insurance_Amt As Decimal = 0
    Public Item_Amt_After_Insurance As Decimal = 0
    Public QC_Status As String = Nothing
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal dtDocDate As DateTime, ByVal objHead As clsSRNHead, ByVal trans As SqlTransaction) As Boolean

        If (objHead.Arr IsNot Nothing AndAlso objHead.Arr.Count > 0) Then
            For Each obj As clsSRNDetail In objHead.Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "SRN_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                clsCommon.AddColumnsForChange(coll, "Row_Type", obj.Row_Type)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Bar_Code", obj.Bar_Code, True)
                clsCommon.AddColumnsForChange(coll, "Item_Desc", obj.Item_Desc)
                clsCommon.AddColumnsForChange(coll, "SRN_Qty", obj.SRN_Qty)
                clsCommon.AddColumnsForChange(coll, "Leak_Qty", obj.Leak_Qty)
                clsCommon.AddColumnsForChange(coll, "Burst_Qty", obj.Burst_Qty)
                clsCommon.AddColumnsForChange(coll, "Short_Qty", obj.Short_Qty)
                clsCommon.AddColumnsForChange(coll, "Rejected_Qty", obj.Rejected_Qty)
                clsCommon.AddColumnsForChange(coll, "Free_qty", obj.Freeqty)
                clsCommon.AddColumnsForChange(coll, "UOM_WEIGHT", obj.UOMWeight)
                clsCommon.AddColumnsForChange(coll, "UOM_WEIGHT_VALUE", obj.UOMWeightValue)

                clsCommon.AddColumnsForChange(coll, "Capex_Code", obj.Capex_Code, True)
                clsCommon.AddColumnsForChange(coll, "Capex_SubCode", obj.Capex_SubCode, True)
                clsCommon.AddColumnsForChange(coll, "Category", obj.Category, True)
                clsCommon.AddColumnsForChange(coll, "Emergency", IIf(obj.Emergency, 1, 0))

                clsCommon.AddColumnsForChange(coll, "MRN_Id", obj.MRN_Id, True)

                clsCommon.AddColumnsForChange(coll, "RGP_Id", obj.RGP_Id, True)
                clsCommon.AddColumnsForChange(coll, "PO_ID", obj.PO_ID, True)
                clsCommon.AddColumnsForChange(coll, "GRN_ID", obj.GRN_ID, True)
                clsCommon.AddColumnsForChange(coll, "Against_Schedule_Code", obj.Against_Schedule_Code, True)

                clsCommon.AddColumnsForChange(coll, "Balance_Qty", obj.Balance_Qty)
                clsCommon.AddColumnsForChange(coll, "Unit_code", obj.Unit_code)
                clsCommon.AddColumnsForChange(coll, "Location", obj.Location)
                clsCommon.AddColumnsForChange(coll, "Item_Cost", obj.Item_Cost)

                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Header_Discount_Per", obj.Header_Discount_Per)
                clsCommon.AddColumnsForChange(coll, "Header_Discount_Amount", obj.Header_Discount_Amount)
                clsCommon.AddColumnsForChange(coll, "Disc_Per", obj.Disc_Per)
                clsCommon.AddColumnsForChange(coll, "Detail_Discount_Amount", obj.Detail_Discount_Amount)
                clsCommon.AddColumnsForChange(coll, "Disc_Per_Unit", obj.Disc_Per_Unit)
                clsCommon.AddColumnsForChange(coll, "Disc_Amt_Per_Unit", obj.Disc_Amt_Per_Unit)
                clsCommon.AddColumnsForChange(coll, "Disc_Type", obj.Disc_Type)
                clsCommon.AddColumnsForChange(coll, "Disc_Amt", obj.Disc_Amt)
                clsCommon.AddColumnsForChange(coll, "Amt_Less_Discount", obj.Amt_Less_Discount)
                clsCommon.AddColumnsForChange(coll, "Taxable_Amount_Per", obj.Taxable_Amount_Per)
                clsCommon.AddColumnsForChange(coll, "Taxable_Amount", obj.Taxable_Amount)
                clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
                clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
                clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
                clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
                clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
                clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
                clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
                clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
                clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
                clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
                clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
                clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", obj.Total_Tax_Amt)
                clsCommon.AddColumnsForChange(coll, "Item_Net_Amt", obj.Item_Net_Amt)
                clsCommon.AddColumnsForChange(coll, "MRP", obj.MRP)
                clsCommon.AddColumnsForChange(coll, "Batch_No", obj.Batch_No)
                clsCommon.AddColumnsForChange(coll, "Bin_No", obj.Bin_No)
                clsCommon.AddColumnsForChange(coll, "Landed_Cost_Rate", obj.Landed_Cost_Rate)
                clsCommon.AddColumnsForChange(coll, "Landed_Cost_Amount", obj.Landed_Cost_Amount)
                clsCommon.AddColumnsForChange(coll, "Specification", obj.Specification)
                clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
                clsCommon.AddColumnsForChange(coll, "Is_Mannual_Amt", obj.Is_Mannual_Amt)
                ' ''clsCommon.AddColumnsForChange(coll, "Landed_Cost_Rate", obj.Landed_Cost_Rate)
                ' ''clsCommon.AddColumnsForChange(coll, "Landed_Cost_Amount", obj.Landed_Cost_Amount)
                If obj.MFG_Date.HasValue Then
                    clsCommon.AddColumnsForChange(coll, "MFG_Date", clsCommon.GetPrintDate(obj.MFG_Date, "dd-MMM-yyyy"))
                End If
                If obj.Expiry_Date.HasValue Then
                    clsCommon.AddColumnsForChange(coll, "Expiry_Date", clsCommon.GetPrintDate(obj.Expiry_Date, "dd-MMM-yyyy"))
                End If
                clsCommon.AddColumnsForChange(coll, "Assessable", obj.Assessable)
                clsCommon.AddColumnsForChange(coll, "AssessableAmt", obj.AssessableAmt)

                clsCommon.AddColumnsForChange(coll, "Fater_Code", obj.Fater_Code)
                clsCommon.AddColumnsForChange(coll, "Fater_Rate", obj.Fater_Rate)
                clsCommon.AddColumnsForChange(coll, "Fater_Amt", obj.Fater_Amt)

                clsCommon.AddColumnsForChange(coll, "PO_Qty", obj.PO_Qty)
                clsCommon.AddColumnsForChange(coll, "GRN_Qty", obj.GRN_Qty)
                clsCommon.AddColumnsForChange(coll, "MRN_Qty", obj.MRN_Qty)
                clsCommon.AddColumnsForChange(coll, "Schedule_Qty", obj.Schedule_Qty)
                clsCommon.AddColumnsForChange(coll, "RGP_Qty", obj.RGP_Qty)

                clsCommon.AddColumnsForChange(coll, "Total_RecTax_PerUnit", obj.Total_RecTax_PerUnit)
                clsCommon.AddColumnsForChange(coll, "Total_NonRecTax_PerUnit", obj.Total_NonRecTax_PerUnit)
                clsCommon.AddColumnsForChange(coll, "Total_AddtionalCost_PerUnit", obj.Total_AddtionalCost_PerUnit)
                clsCommon.AddColumnsForChange(coll, "Req_No", obj.Req_No, True)
                '' for abatement PO
                clsCommon.AddColumnsForChange(coll, "AbatementRate", obj.AbatementRate)
                clsCommon.AddColumnsForChange(coll, "AssessableMRP", obj.AssessableMRP)
                clsCommon.AddColumnsForChange(coll, "TotalAssessableMRP", obj.TotalAssessableMRP)

                clsCommon.AddColumnsForChange(coll, "Accepted_Amount", obj.Accepted_Amount)
                clsCommon.AddColumnsForChange(coll, "Rejected_Amount", obj.Rejected_Amount)
                clsCommon.AddColumnsForChange(coll, "Shortage_Amount", obj.Shortage_Amount)
                clsCommon.AddColumnsForChange(coll, "Leak_Amount", obj.Leak_Amount)
                clsCommon.AddColumnsForChange(coll, "Burst_Amount", obj.Leak_Amount)
                clsCommon.AddColumnsForChange(coll, "Amt_Less_Discount_Without_Shortage", obj.Amt_Less_Discount_Without_Shortage)

                ''==========================19/10/2016======================
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code1", obj.ItemAdd_Charge_Code1)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt1", obj.ItemAdd_Org_Charge_Amt1)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt1", obj.ItemAdd_Calc_Charge_Amt1)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code9", obj.ItemAdd_Charge_Code9)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt9", obj.ItemAdd_Org_Charge_Amt9)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt9", obj.ItemAdd_Calc_Charge_Amt9)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code8", obj.ItemAdd_Charge_Code8)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt8", obj.ItemAdd_Org_Charge_Amt8)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt8", obj.ItemAdd_Calc_Charge_Amt8)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code7", obj.ItemAdd_Charge_Code7)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt7", obj.ItemAdd_Org_Charge_Amt7)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt7", obj.ItemAdd_Calc_Charge_Amt7)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code6", obj.ItemAdd_Charge_Code6)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt6", obj.ItemAdd_Org_Charge_Amt6)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt6", obj.ItemAdd_Calc_Charge_Amt6)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code5", obj.ItemAdd_Charge_Code5)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt5", obj.ItemAdd_Org_Charge_Amt5)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt5", obj.ItemAdd_Calc_Charge_Amt5)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code4", obj.ItemAdd_Charge_Code4)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt4", obj.ItemAdd_Org_Charge_Amt4)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt4", obj.ItemAdd_Calc_Charge_Amt4)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code3", obj.ItemAdd_Charge_Code3)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt3", obj.ItemAdd_Org_Charge_Amt3)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt3", obj.ItemAdd_Calc_Charge_Amt3)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code2", obj.ItemAdd_Charge_Code2)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt2", obj.ItemAdd_Org_Charge_Amt2)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt2", obj.ItemAdd_Calc_Charge_Amt2)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Charge_Code10", obj.ItemAdd_Charge_Code10)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Org_Charge_Amt10", obj.ItemAdd_Org_Charge_Amt10)
                clsCommon.AddColumnsForChange(coll, "ItemAdd_Calc_Charge_Amt10", obj.ItemAdd_Calc_Charge_Amt10)
                clsCommon.AddColumnsForChange(coll, "Total_ItemAdd_Charge", obj.Total_ItemAdd_Charge)
                ''=================================================================
                clsCommon.AddColumnsForChange(coll, "Against_Item_Wise_Tax_Rate", obj.Against_Item_Wise_Tax_Rate, True)

                clsCommon.AddColumnsForChange(coll, "Insurance_Per", obj.Insurance_Per)
                clsCommon.AddColumnsForChange(coll, "Insurance_Base_Amt", obj.Insurance_Base_Amt)


                clsCommon.AddColumnsForChange(coll, "Item_Insurance_Base_Amt", obj.Item_Insurance_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "Item_Insurance_Apply_On", obj.Item_Insurance_Apply_On)
                clsCommon.AddColumnsForChange(coll, "Item_Insurance_Rate", obj.Item_Insurance_Rate)
                clsCommon.AddColumnsForChange(coll, "Item_Insurance_Amt", obj.Item_Insurance_Amt)
                clsCommon.AddColumnsForChange(coll, "Item_Amt_After_Insurance", obj.Item_Amt_After_Insurance)

                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Dim strLocationCode As String
                If clsCommon.myLen(objHead.Sublocation_Code) > 0 Then
                    strLocationCode = objHead.Sublocation_Code
                ElseIf clsCommon.myLen(objHead.Ship_To_Location) > 0 Then
                    strLocationCode = objHead.Ship_To_Location
                Else
                    strLocationCode = objHead.Bill_To_Location
                End If
                clsBatchInventory.SaveData("SRN", strDocNo, dtDocDate, "I", obj.Item_Code, strLocationCode, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                clsSerializeInvenotry.SaveData("SRN", strDocNo, dtDocDate, "I", obj.Item_Code, IIf(clsCommon.myLen(objHead.Ship_To_Location) > 0, objHead.Ship_To_Location, objHead.Bill_To_Location), obj.Line_No, obj.arrSrItem, trans)
            Next
        End If
        Return True
    End Function

    Public Shared Function GetBalanceSRNQty(ByVal strSRNCode As String, ByVal strICode As String, ByVal strCurrPINNo As String, ByVal strUOM As String, ByVal dblMRP As Double, ByVal dblAssessable As Double) As Double
        Dim qry As String = "select SUM(qty * RI) as Balance from(  " &
            " select TSPL_SRN_DETAIL.Item_Code as ICode,TSPL_SRN_DETAIL.SRN_Qty as Qty,1 as RI from TSPL_SRN_DETAIL left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No where TSPL_SRN_DETAIL.Status=0 and TSPL_SRN_HEAD.Status=1 and TSPL_SRN_DETAIL.SRN_No ='" + strSRNCode + "' and TSPL_SRN_DETAIL.Item_Code='" + strICode + "' and  TSPL_SRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_SRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_SRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' " &
            " union all " &
            " select TSPL_PI_DETAIL.Item_Code as ICode,TSPL_PI_DETAIL.PI_Qty as Qty,-1 as RI from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No where TSPL_PI_DETAIL.SRN_Id='" + strSRNCode + "'   and TSPL_PI_DETAIL.Item_Code='" + strICode + "'  and  TSPL_PI_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_PI_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_PI_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'  and TSPL_PI_DETAIL.PI_No not in ('" + strCurrPINNo + "')  " &
            " )Final "
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
    End Function

    Public Shared Function CompleteSRN(ByVal strDoccode As String, ByVal strICode As String, ByVal LineNo As Integer) As Boolean
        Dim qry As String = "update TSPL_SRN_DETAIL set Status=1 where SRN_No='" + strDoccode + "' and Line_No='" + clsCommon.myCstr(LineNo) + "' and Item_Code='" + strICode + "'"
        Return clsDBFuncationality.ExecuteNonQuery(qry)
    End Function
End Class

Public Class clsSRNRoadPermitDetail
#Region "Variables"
    Public roadvendor As String = Nothing
    Public roadpono As String = Nothing
    Public roadcode As String = Nothing
    Public roadissue_no As String = Nothing
    Public RoadpermitSRNNO As String = Nothing
#End Region

    Public Shared Function SaveData_RoadPermit(ByVal srnno As String, ByVal pono As String, ByVal arr As List(Of clsSRNRoadPermitDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = ""

            'If clsCommon.myLen(pono) <= 0 Then
            qry = "delete from TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL where srn_no='" + srnno + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            'End If

            Dim isSaved As Boolean = True
            Dim coll As New Hashtable()
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For Each objtr As clsSRNRoadPermitDetail In arr
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "PurchaseOrder_No", pono)
                    clsCommon.AddColumnsForChange(coll, "Vendor_Code", objtr.roadvendor)
                    clsCommon.AddColumnsForChange(coll, "Form_Code", objtr.roadcode)
                    clsCommon.AddColumnsForChange(coll, "Issue_No", objtr.roadissue_no)
                    clsCommon.AddColumnsForChange(coll, "srn_no", srnno)
                    clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                    clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))

                    'qry = "delete from TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL where TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.purchaseorder_no='" + objtr.roadpono + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.form_code='" + objtr.roadcode + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.issue_no='" + objtr.roadissue_no + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.vendor_code='" + objtr.roadvendor + "'"
                    'clsDBFuncationality.ExecuteNonQuery(qry, trans)

                    'If clsCommon.myLen(pono) > 0 Then
                    'isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL", OMInsertOrUpdate.Update, " TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.purchaseorder_no='" + objtr.roadpono + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.form_code='" + objtr.roadcode + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.issue_no='" + objtr.roadissue_no + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.vendor_code='" + objtr.roadvendor + "'", trans)
                    'Else
                    isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                    'End If
                Next
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class clsSRNCFORMDetail
#Region "Variables"
    Public cformvendor As String = Nothing
    Public cformpono As String = Nothing
    Public cformcode As String = Nothing
    Public cformissue_no As String = Nothing
    Public cformSRNNO As String = Nothing
#End Region

    Public Shared Function SaveData_CFORM(ByVal srnno As String, ByVal pono As String, ByVal arr As List(Of clsSRNCFORMDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = ""

            'If clsCommon.myLen(pono) <= 0 Then
            qry = "delete from TSPL_CFORM_ISSUE_RECEIVE_DETAIL where srn_no='" + srnno + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            'End If

            Dim isSaved As Boolean = True
            Dim coll As New Hashtable()
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For Each objtr As clsSRNCFORMDetail In arr
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "PurchaseOrder_No", pono)
                    clsCommon.AddColumnsForChange(coll, "Vendor_Code", objtr.cformvendor)
                    clsCommon.AddColumnsForChange(coll, "Form_Code", objtr.cformcode)
                    clsCommon.AddColumnsForChange(coll, "Issue_No", objtr.cformissue_no)
                    clsCommon.AddColumnsForChange(coll, "srn_no", srnno)
                    clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                    clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))

                    'qry = "delete from TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL where TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.purchaseorder_no='" + objtr.cformpono + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.form_code='" + objtr.cformcode + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.issue_no='" + objtr.cformissue_no + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.vendor_code='" + objtr.cformvendor + "'"
                    'clsDBFuncationality.ExecuteNonQuery(qry, trans)

                    'If clsCommon.myLen(pono) <= 0 Then
                    isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_CFORM_ISSUE_RECEIVE_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                    'Else
                    'isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_CFORM_ISSUE_RECEIVE_DETAIL", OMInsertOrUpdate.Update, " TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.purchaseorder_no='" + objtr.cformpono + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.form_code='" + objtr.cformcode + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.issue_no='" + objtr.cformissue_no + "' and TSPL_ROADPERMIT_ISSUE_RECEIVE_DETAIL.vendor_code='" + objtr.cformvendor + "'", trans)
                    'End If
                Next
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

<Serializable()>
Public Class clsSRNAdditionChargeInsurance
#Region "Variables"
    Public TR_Code As String = Nothing
    Public DocNo As String = Nothing
    Public AC_Code As String = Nothing
    Public AC_Name As String = Nothing ''Not a table Field
    Public Amount As Decimal
#End Region
    Public Shared Function SaveData(ByVal DocNo As String, ByVal DocDate As DateTime, ByVal arr As List(Of clsSRNAdditionChargeInsurance), ByVal trans As SqlTransaction) As Boolean
        Try
            Dim coll As New Hashtable()
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For Each objtr As clsSRNAdditionChargeInsurance In arr
                    coll = New Hashtable()
                    objtr.TR_Code = clsERPFuncationality.GetNextCode(trans, DocDate, clsDocType.Detail, clsDocTransactionType.Detail, "")
                    clsCommon.AddColumnsForChange(coll, "TR_Code", objtr.TR_Code)
                    clsCommon.AddColumnsForChange(coll, "SRN_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "AC_Code", objtr.AC_Code)
                    clsCommon.AddColumnsForChange(coll, "Amount", objtr.Amount)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SRN_ADDITION_CHARGE_INSURANCE", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function DeleteData(ByVal DocNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = "delete from TSPL_SRN_ADDITION_CHARGE_INSURANCE where SRN_No='" + DocNo + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Return True
    End Function
    Public Shared Function GetData(ByVal DocNo As String, ByVal trans As SqlTransaction) As List(Of clsSRNAdditionChargeInsurance)
        Dim qry As String = "select TSPL_SRN_ADDITION_CHARGE_INSURANCE.*,TSPL_Additional_Charges.Description as AC_Name  from TSPL_SRN_ADDITION_CHARGE_INSURANCE left outer join TSPL_Additional_Charges on TSPL_Additional_Charges.Code=TSPL_SRN_ADDITION_CHARGE_INSURANCE.AC_Code where TSPL_SRN_ADDITION_CHARGE_INSURANCE.SRN_No='" + DocNo + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Dim Arr_ACInsurance As List(Of clsSRNAdditionChargeInsurance) = Nothing
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Arr_ACInsurance = New List(Of clsSRNAdditionChargeInsurance)
            For Each dr As DataRow In dt.Rows
                Dim objtr As New clsSRNAdditionChargeInsurance()
                objtr.TR_Code = clsCommon.myCstr(dr("TR_Code"))
                objtr.DocNo = clsCommon.myCstr(dr("SRN_No"))
                objtr.AC_Code = clsCommon.myCstr(dr("AC_Code"))
                objtr.AC_Name = clsCommon.myCstr(dr("AC_Name"))
                objtr.Amount = clsCommon.myCstr(dr("Amount"))
                Arr_ACInsurance.Add(objtr)
            Next
        End If
        Return Arr_ACInsurance
    End Function
End Class
