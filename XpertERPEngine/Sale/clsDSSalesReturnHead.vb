Imports common
Imports System.Data.SqlClient
Public Class clsDSSalesReturnHead
#Region "Variables"
    Public TotalCAN As Double = 0
    Public ShippedCAN As Double = 0
    Public CrateQty As Integer = 0
    Public Is_Taxable As Integer = 0
    Public Gate_ReturnNo As String = Nothing
    Public Sub_Location_code As String = String.Empty
    Public CrateMan As Double = 0
    Public jaali As Double = 0
    Public Box As Double = 0
    Public Trans_type As String = String.Empty
    Public RoundOffAmount As Double = 0
    Public Return_Type As String = Nothing
    Public Shift_Type As String = Nothing
    Public Damage_Type As String = Nothing
    Public Cust_PO_No As String = Nothing
    Public Price_Group_Code As String = Nothing
    Public Invoice_Type As String = Nothing
    Public PROJECT_ID As String = Nothing
    Public Price_Code As String = Nothing
    Public Route_No As String = Nothing
    Public Route_Desc As String = Nothing
    Public HeadDisc_Per As Double = 0
    Public HeadDisc_Amt As Double = 0
    Public HeadDisc_PerAmt As Double = 0
    Public TotCashDiscAmt As Double = 0

    Public Document_Code As String = Nothing
    Public Document_Date As DateTime
    Public Customer_Code As String = Nothing
    Public Customer_Name As String = Nothing  'Not a table field
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public On_Hold As Boolean = Nothing
    Public Ref_No As String = Nothing
    Public Description As String = Nothing
    Public Remarks As String = Nothing
    Public Tax_Group As String = Nothing
    Public TaxGroupName As String = Nothing 'Not a table field
    Public Bill_To_Location As String = Nothing
    Public BillToLocationName As String = Nothing 'Not a table field
    Public Ship_To_Location As String = Nothing
    Public ShipToLocationName As String = Nothing 'Not a table field
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
    Public Discount_Base As Double = 0
    Public Discount_Amt As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public Total_Tax_Amt As Double = 0
    Public Total_Amt As Double = 0
    Public Comments As String = Nothing
    Public Comp_Code As String = Nothing
    Public Terms_Code As String = Nothing
    Public TermsName As String = Nothing
    Public Due_Date As DateTime? = Nothing
    Public Posting_Date As DateTime? = Nothing
    Public Carrier As String = Nothing
    Public Vehicle_Code As String = Nothing
    Public VehicleNo As String = Nothing
    Public GRNo As String = Nothing
    Public GENo As String = Nothing
    Public GEDate As Date? = Nothing
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
    Public Dept As String = Nothing
    Public Dept_Desc As String = Nothing
    Public Item_Type As String = Nothing
    Public Challan_No As String = Nothing
    Public Challan_Date As DateTime? = Nothing
    Public Inv_No As String = Nothing
    Public Inv_Date As DateTime? = Nothing
    Public Against_Invoice_No As String = Nothing
    Public Is_Internal As Boolean = False
    Public Tax_Calculation_Type As EnumTaxCalucationType
    Public Salesman_Code As String = Nothing
    Public Salesman_Name As String = Nothing
    Public Arr As List(Of clsDSSalesReturnDetail) = Nothing
    Public Booth_Arr As List(Of clsDSSalesReturnBookingDetail) = Nothing

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing
    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public Is_Cancelled As Integer = 0
    Public Screen_Type As String = ""
    Public IsSampling As Integer = 0
    Public Total_Comm_Amt As Double = 0
    Public Distributor_Commission_TotalAmt As Decimal = 0
    Public Transporter_Commission_TotalAmt As Decimal = 0
    Public Security_TotalAmt As Decimal = 0
    Public isMultipleReturn As Decimal = 0
    Public Is_Without_Invoice As Integer = 0

#End Region
    Public Shared Function funsaleReturnDairyPrint(ByVal Form_ID As String, ByVal strCancelDelete As String, ByVal txtDate As DateTime, ByVal StrCode As String, ByVal IsPDF As Boolean) As Boolean

        Dim TSPL_SD_SALE_RETURN_HEAD As String = Nothing
        Dim TSPL_SD_SALE_RETURN_detail As String = Nothing
        If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal Then
            TSPL_SD_SALE_RETURN_HEAD = "TSPL_SD_SALE_RETURN_HEAD_cancel_data"
            TSPL_SD_SALE_RETURN_detail = "TSPL_SD_SALE_RETURN_detail_cancel_data"
        ElseIf clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
            TSPL_SD_SALE_RETURN_HEAD = "TSPL_SD_SALE_RETURN_HEAD_Delete_data"
            TSPL_SD_SALE_RETURN_detail = "TSPL_SD_SALE_RETURN_detail_Delete_data"
        ElseIf clsCommon.CompairString(strCancelDelete, "Dispatch_Return") = CompairStringResult.Equal Then
            TSPL_SD_SALE_RETURN_HEAD = "TSPL_SD_SHIPMENT_RETURN_HEAD"
            TSPL_SD_SALE_RETURN_detail = "TSPL_SD_SHIPMENT_RETURN_detail"
        Else
            TSPL_SD_SALE_RETURN_HEAD = "TSPL_SD_SALE_RETURN_HEAD"
            TSPL_SD_SALE_RETURN_detail = "TSPL_SD_SALE_RETURN_detail"
        End If


        If clsCommon.myLen(StrCode) > 0 Then
            Dim Qry As String = "select "
            If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
                Qry += " 'Cancelled' As Report_Status, "
            Else
                Qry += " '' As Report_Status, "
            End If
            Qry += " 
  Final.* 
from 
  (
    select 
      Main_Final.*, 
      TSPL_COMPANY_MASTER.Access_Officer as FSSAI_NO, 
      TSPL_COMPANY_MASTER.Logo_Img, 
      1 As CopyType, 
      TSPL_COMPANY_MASTER.GSTReg_No As SellerGST, 
      TSPL_COMPANY_MASTER.Pan_No, 
      Rate_Default As RateLtr 
    from 
      (
        select 
          final.*, 
          tbl_Brand.Brand, 
          tbl_Brand.BRANDDESC, 
          Item_Desc + '   ' + isnull('HSN ' + HSN_Code, '') as Particulars 
        From 
          (
            select 
              TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate, 
              --case when TSPL_BOOKING_MATSER.Is_CashSale = 'Y' then TSPL_SD_SALE_INVOICE_HEAD.Payment_Terms else 'CREDIT' END AS PaymentTerms, 
              --TSPL_BOOKING_MATSER.Is_Distributor, 
              --TSPL_BOOKING_MATSER.Is_BPL, 
              --TSPL_BOOKING_MATSER.Is_CashSale, 
              --TSPL_BOOKING_MATSER.Is_DCS, 
              --TSPL_BOOKING_MATSER.Booking_Type, 
              TSPL_COMPANY_MASTER.CST_LST, 
              (
                Case When " + TSPL_SD_SALE_RETURN_HEAD + ".Invoice_Type = 'T' Then cast(
                  " + TSPL_SD_SALE_RETURN_HEAD + ".BarCode_Img as image
                ) End
              ) as BarCode_Img, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".VehicleNo as Manual_VehicleNo, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Payment_Terms, 
              --" + TSPL_SD_SALE_RETURN_HEAD + ".ReceiverName, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Security_TotalAmt, 
              --convert(
              --  varchar(12), 
              --  TSPL_SD_SALE_INVOICE_HEAD.Supply_Date, 
              --  103
              --) Supply_Date, 
              --case when TSPL_SD_SALE_INVOICE_HEAD.Shift_Type = 'AM' then 'Morning' else 'Evening' end as Shift_Type, 
              Case when tspl_item_master.Is_FreshItem = 1 then (
                (
                  " + TSPL_SD_SALE_RETURN_detail + ".ActualReturnQty * isnull(
                    TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 
                    1
                  )
                ) / coalesce(ITEMDETAIL3.LTR, 1)
              ) when tspl_item_master.Is_Ambient = 1 then (
                (
                   " + TSPL_SD_SALE_RETURN_detail + ".Qty * isnull(
                    TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 
                    1
                  )
                ) / coalesce(ITEMDETAIL3.kg, 1)
              ) end as QTY_LTRKG, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX1 as ITAX1, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX1_RATE AS ITAX1_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX1_Amt as ITAX1_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX2 as ITAX2, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX2_RATE AS ITAX2_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX2_Amt as ITAX2_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX3 AS ITAX3, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX3_Rate AS ITAX3_Rate, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX3_Amt as ITAX3_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX4 AS ITAX4, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX4_RATE AS ITAX4_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX4_Amt as ITAX4_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX5 as ITAX5, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX5_RATE AS ITAX5_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX5_Amt as ITAX5_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX6 as ITAX6, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX6_RATE AS ITAX6_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX6_Amt as ITAX6_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX7 AS ITAX7, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX7_Rate AS ITAX7_Rate, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX7_Amt as ITAX7_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX8 AS ITAX8, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX8_RATE AS ITAX8_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX8_Amt as ITAX8_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX9 AS ITAX9, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX9_Rate AS ITAX9_Rate, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX9_Amt as ITAX9_Amt, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX10 AS ITAX10, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX10_RATE AS ITAX10_RATE, 
               " + TSPL_SD_SALE_RETURN_detail + ".TAX10_Amt as ITAX10_Amt, 
              (
                Case When " + TSPL_SD_SALE_RETURN_HEAD + ".Invoice_Type = 'T' Then 'IRN : ' + " + TSPL_SD_SALE_RETURN_HEAD + ".IRN_No End
              ) As IRN_No,  
              Zone_Code, 
              ITEMDETAIL1.Conversion_Factor As CF, 
              TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".EInvoice_Type, 
              0 as LeakageDeduction_Freshsale, 
              0 as LeakageDeduction, 
              case when TSPL_ITEM_PRICE_MASTER.Price_Comp1 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount1, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp2 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount2, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp3 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount3, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp4 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount4, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp5 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount5, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp6 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount6, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp7 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount7, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp8 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount8, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp9 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount9, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp10 = 'SCM' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount10, 
                0
              ) end as SCM, 
              case when TSPL_ITEM_PRICE_MASTER.Price_Comp1 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount1, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp2 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount2, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp3 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount3, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp4 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount4, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp5 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount5, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp6 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount6, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp7 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount7, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp8 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount8, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp9 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount9, 
                0
              ) when TSPL_ITEM_PRICE_MASTER.Price_Comp10 = 'DIS-MARGIN' then isnull(
                 " + TSPL_SD_SALE_RETURN_detail + ".Price_Amount10, 
                0
              ) end as DIS_MARGIN, 
              TSPL_LOCATION_MASTER.Location_Desc, 
              TSPL_LOCATION_MASTER.Loc_Short_Name, 
              TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin, 
              (
                case when isnull(TSPL_LOCATION_MASTER.Phone1, '')<> '' then TSPL_LOCATION_MASTER.Phone1 when isnull(TSPL_LOCATION_MASTER.Phone2, '')<> '' then + ', ' + TSPL_LOCATION_MASTER.Phone2 end
              ) as Loc_Phone, 
              TSPL_LOCATION_MASTER.Email as Loc_Eamil, 
              '' as Loc_Website, 
              TSPL_COMPANY_MASTER.ISO_No, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Against_Invoice_No as Invoice_No, 
              convert(
                varchar(12), 
                TSPL_SD_SALE_INVOICE_HEAD.Document_date, 
                103
              ) as Invoice_Date, 
              customer_city_master.city_name as Cust_City, 
              '' as Against_Shipment_No, 
              CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, 
              '' as Electronic_Ref_No, 
              Tspl_customer_master.gstno as CustGSTNo, 
              TSPL_STATE_MASTER.gst_state_code, 
              tspl_location_master.gstno as LocGstNo, 
              '' as EWayBillNo, 
              '' as EWayBillDate, 
              TSPL_ITEM_MASTER.HSN_Code, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Remarks as InvRemarks, 
              '' as Delivery_Code, 
              ITEMDETAIL.Conversion_factor, 
              (
                CASE when  " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  case when  " + TSPL_SD_SALE_RETURN_detail + ".Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor, 0)= 0 then 0 else convert(
                    Decimal(18, 2), 
                     " + TSPL_SD_SALE_RETURN_detail + ".Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor / coalesce(ITEMDETAIL.Conversion_factor, 1)
                  ) end else 0 end
                ) end
              ) as QTY_Box, 
              TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, 
              Case When ISNULL(
                " + TSPL_SD_SALE_RETURN_HEAD + ".Vehicle_Code, 
                ''
              )<> '' Then " + TSPL_SD_SALE_RETURN_HEAD + ".Vehicle_Code WHEN ISNULL(
                " + TSPL_SD_SALE_RETURN_HEAD + ".VehicleNo, 
                ''
              )<> '' Then TSPL_VEHICLE_MASTER.Number Else " + TSPL_SD_SALE_RETURN_HEAD + ".vehicleNo End as vehicleNo, 
              Convert(
                varchar, TSPL_SD_SALE_INVOICE_HEAD.Document_Date, 
                103
              ) as Sale_Invoice_Date, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".RoundOffAmount, 
              TSPL_LOCATION_MASTER.Add1 as Loc_ADd1, 
              TSPL_LOCATION_MASTER.Add2 as LOC_ADD2, 
              TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, 
              TSPL_STATE_MASTER.State_Name as LocationState, 
              case when ISNULL(TSPL_LOCATION_MASTER.Phone1, '')= '(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end + Case When ISNULL(TSPL_LOCATION_MASTER.Phone2, '')<> '(+__)__________' Then ', ' + TSPL_LOCATION_MASTER.Phone2 Else '' End as LOCPhone, 
              TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code, 
              convert(
                varchar, " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Date, 
                103
              ) as Document_Date, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Description, 
              '' as Lorry_No, 
              TSPL_ITEM_MASTER.Sku_Seq, 
              CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then " + TSPL_SD_SALE_RETURN_detail + ".Item_Code + ' -Scheme' else " + TSPL_SD_SALE_RETURN_detail + ".Item_Code end as Item_Code, 
              (
                CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  " + TSPL_SD_SALE_RETURN_detail + ".Line_No
                ) end
              ) as Line_No, 
              CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then TSPL_ITEM_MASTER.Item_Desc + ' -Scheme' else TSPL_ITEM_MASTER.Item_Desc end as Item_Desc, 
              " + TSPL_SD_SALE_RETURN_detail + ".Crate as QtyCrates, 
              ITEMDETAIL2.Conversion_Factor As ConvFactInCrate, 
              CEILING(
                (
                  " + TSPL_SD_SALE_RETURN_detail + ".Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor
                )/ ITEMDETAIL2.Conversion_Factor
              ) As ConvQtyInCrate, 
              " + TSPL_SD_SALE_RETURN_detail + ".ActualUOM as Unit_code, 
              convert(
                Decimal(18, 3), 
                " + TSPL_SD_SALE_RETURN_detail + ".ActualQty +" + TSPL_SD_SALE_RETURN_detail + ".DamageQty
              ) as Qty_Default, 
              (
                CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  convert(
                    Decimal(18, 2), 
                    case when " + TSPL_SD_SALE_RETURN_detail + ".Qty > 0 then convert(
                      DECIMAL(18, 5), 
                      (
                        case when " + TSPL_SD_SALE_RETURN_detail + ".Sampling = 1 then 0 else " + TSPL_SD_SALE_RETURN_detail + ".Amount end
                      )/ (
                        " + TSPL_SD_SALE_RETURN_detail + ".ActualQty +" + TSPL_SD_SALE_RETURN_detail + ".DamageQty
                      )
                    ) else 0 end
                  )
                ) end
              ) as Rate_Default, 
              (
                CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  convert(
                    Decimal(18, 2), 
                    (" + TSPL_SD_SALE_RETURN_detail + ".ActualQty +" + TSPL_SD_SALE_RETURN_detail + ".DamageQty) * TSPL_ITEM_UOM_DETAIL.Conversion_Factor
                  )
                ) end
              ) as QtyPCS, 
              coalesce(
                case when " + TSPL_SD_SALE_RETURN_detail + ".Sampling = 1 then 0 else SUB_QTY end, 
                0
              ) as free_qty, 
              " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item as FreeSchemeInLitres, 
              (
                CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  case when " + TSPL_SD_SALE_RETURN_detail + ".Qty > 0 then convert(
                    DECIMAL(18, 5), 
                    (
                      case when " + TSPL_SD_SALE_RETURN_detail + ".Sampling = 1 then 0 else " + TSPL_SD_SALE_RETURN_detail + ".Amount end
                    )/ (
                      (" + TSPL_SD_SALE_RETURN_detail + ".ActualQty +" + TSPL_SD_SALE_RETURN_detail + ".DamageQty)* TSPL_ITEM_UOM_DETAIL.Conversion_Factor
                    )
                  ) else 0 end
                ) end
              ) as RatePerPcs, 
              (
                CASE when " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  (
                    case when " + TSPL_SD_SALE_RETURN_detail + ".Sampling = 1 then 0 else " + TSPL_SD_SALE_RETURN_detail + ".TAX1_Base_Amt end
                  )
                ) end
              ) as valueInRs, 
              '' as comp_add2, 
              '' as comp_add3, 
              TSPL_COMPANY_MASTER.Phone1 as CompPhone, 
              (
                CASE when  " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'Y' then 0 else (
                  coalesce(
                     " + TSPL_SD_SALE_RETURN_detail + ".Cash_Scheme_Amount, 
                    0
                  )
                ) end
              ) as Cash_Scheme_Amount, 
              isnull(schemeInCrates, 0) as schemeInCrates, 
              '' GrandTotalCrates, 
              TSPL_COMPANY_MASTER.Comp_Code, 
              TSPL_COMPANY_MASTER.Comp_Name, 
              TSPL_COMPANY_MASTER.Add1 as comp_add1, 
              TSPL_COMPANY_MASTER.Fax as comp_Fax, 
              TSPL_COMPANY_MASTER.Email as comp_Email, 
              TSPL_COMPANY_MASTER.Tin_No as comp_tinNo, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Customer_Code as cust_Code, 
              TSPL_CUSTOMER_MASTER.Customer_Name, 
              TSPL_CUSTOMER_MASTER.Add1 as cust_add1, 
              TSPL_CUSTOMER_MASTER.Add2 as cust_add2, 
              TSPL_CUSTOMER_MASTER.Add3 cust_add3, 
              case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1, '')= '(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end + Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2, '')<> '(+__)__________' Then ', ' + TSPL_CUSTOMER_MASTER.Phone2 Else '' End as CustPhone, 
              TSPL_CUSTOMER_MASTER.Fax as cust_fax, 
              TSPL_CUSTOMER_MASTER.State as Cust_state, 
              CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename, 
              TSPL_CUSTOMER_MASTER.Email as cust_Email, 
              TSPL_CUSTOMER_MASTER.WebSite as cust_website, 
              TSPL_CUSTOMER_MASTER.pan as Customer_Pan, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".Ack_No, 
                'NA'
              ) AS Ack_No, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Ack_Date, 
              TSPL_SD_SALE_INVOICE_HEAD.Invoice_Type As TaxableNonTaxable, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX1, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX1
              ) as TaxType1, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX1_Amt, 
                0.00
              ) As TAX1_Amt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX1_Rate, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX1_Amt as TAX1Amt, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX2
              ) as TaxType2, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX2, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX2_Amt, 
                0.00
              ) As TAX2_Amt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX2_Rate, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX2_Amt as TAX2Amt, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX3
              ) as TaxType3, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX3, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX3_Amt, 
                0.00
              ) As TAX3_Amt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX3_Rate, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX3_Amt as TAX3Amt, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX4
              ) as TaxType4, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX4, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX4_Amt, 
                0.00
              ) As TAX4_Amt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX4_Rate, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX4_Amt as TAX4Amt, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX5
              ) as TaxType5, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX5, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX5_Amt, 
                0.00
              ) As TAX5_Amt, 
              (
                select 
                  type 
                from 
                  TSPL_TAX_MASTER 
                where 
                  Tax_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".TAX6
              ) as TaxType6, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".TAX6, 
              IsNull(
                " + TSPL_SD_SALE_RETURN_HEAD + ".TAX6_Amt, 
                0.00
              ) As TAX6_Amt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Route_No, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Route_Desc, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Distributor_Commission_TotalAmt, 
              " + TSPL_SD_SALE_RETURN_HEAD + ".Transporter_Commission_TotalAmt, 
              --isnull(
              --  TSPL_SD_SALE_INVOICE_HEAD.Against_booking_no, 
              --  ''
              --)
			  '' as Against_Delivery_Code, 
              TabBatch.Batch_No as batchNO, 
              TabBatch.BATCH_QTY AS Batchqty, 
              Case when TSPL_CUSTOMER_MASTER.Credit_Customer = 'Y' THEN 'CREDIT' else '' end as Credit_Customer, 
              TSPL_SHIP_TO_LOCATION.Ship_To_Code, 
              TSPL_SHIP_TO_LOCATION.Ship_To_Desc, 
              TSPL_SHIP_TO_LOCATION.Ship_Address, 
              TSPL_SHIP_TO_LOCATION.Ship_City, 
              TSPL_SHIP_TO_LOCATION.Ship_State, 
              Convert(
                Varchar, TSPL_SHIP_TO_LOCATION.Ship_Pin_Code
              ) Ship_Pin_Code, 
              TSPL_SHIP_TO_LOCATION.Ship_PAN, 
              TSPL_SHIP_TO_LOCATION.Ship_GSTNO,
              --IsNull(
              --  TSPL_SD_SHIPMENT_DETAIL.Booth_Security_Amt, 
              --  0
              --)
			  0 as  Booth_Security_Amt 
            from 
              " + TSPL_SD_SALE_RETURN_detail + " 
              LEFT OUTER JOIN " + TSPL_SD_SALE_RETURN_HEAD + " ON " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code = " + TSPL_SD_SALE_RETURN_detail + ".DOCUMENT_CODE 
              left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Against_Invoice_No 
              --left outer join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE 
              --and TSPL_SD_SALE_INVOICE_DETAIL.Line_No = " + TSPL_SD_SALE_RETURN_detail + ".Line_No 
              --left outer join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SALE_INVOICE_HEAD.Against_Booking_No 
              left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              And TSPL_ITEM_UOM_DETAIL.UOM_Code = " + TSPL_SD_SALE_RETURN_detail + ".ActualUOM 
              LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              left join (
                select 
                  Conversion_factor, 
                  TSPL_ITEM_UOM_DETAIL.Item_code 
                from 
                  TSPL_ITEM_UOM_DETAIL 
                where 
                  UOM_code = 'BOX'
              ) as ITEMDETAIL on ITEMDETAIL.Item_code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              left join (
                select 
                  Conversion_factor, 
                  TSPL_ITEM_UOM_DETAIL.Item_code 
                from 
                  TSPL_ITEM_UOM_DETAIL 
                where 
                  UOM_code in (
                    select 
                      case when Is_FreshItem = 1 then 'LTR' else 'KG' end 
                    from 
                      TSPL_ITEM_MASTER 
                    where 
                      Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code
                  )
              ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              left join (
                SELECT 
                  * 
                FROM 
                  (
                    select 
                      item_code, 
                      uom_code, 
                      conversion_factor 
                    from 
                      TSPL_ITEM_UOM_DETAIL
                  ) I PIVOT (
                    Max(conversion_factor) FOR uom_code IN ([KG], [LTR])
                  ) P
              ) ITEMDETAIL3 ON " + TSPL_SD_SALE_RETURN_detail + ".Item_Code = ITEMDETAIL3.item_code 
              left join (
                select 
                  Conversion_factor, 
                  TSPL_ITEM_UOM_DETAIL.Item_code 
                from 
                  TSPL_ITEM_UOM_DETAIL 
                where 
                  ITEM_Code = TSPL_ITEM_UOM_DETAIL.Item_Code 
                  And UOM_code = 'Crate'
              ) as ITEMDETAIL2 on ITEMDETAIL2.Item_code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Comp_Code 
              left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Customer_Code 
              left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Bill_To_Location 
              LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code = TSPL_LOCATION_MASTER.State 
              left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = CUSTOMER_STATE_MASTER.STATE_CODE 
              left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code = customer_city_master.City_Code 
              Left Outer Join (
                select 
                  Ship_To_Code, 
                  Ship_To_Desc, 
                  Ship_To_Type_Desc, 
                  (
                    TSPL_SHIP_TO_LOCATION.Add1 + ' ' + TSPL_SHIP_TO_LOCATION.Add2 + ' ' + TSPL_SHIP_TO_LOCATION.Add3 + ' ' + TSPL_SHIP_TO_LOCATION.Add4
                  ) As Ship_Address, 
                  TSPL_CITY_MASTER.City_Name As Ship_City, 
                  TSPL_STATE_MASTER.STATE_NAME As Ship_State, 
                  Pin_Code As Ship_Pin_Code, 
                  PAN As Ship_PAN, 
                  GSTNO As Ship_GSTNO 
                from 
                  TSPL_SHIP_TO_LOCATION 
                  Left Outer Join TSPL_CITY_MASTER On TSPL_CITY_MASTER.City_Code = TSPL_SHIP_TO_LOCATION.City_Code 
                  Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE = TSPL_SHIP_TO_LOCATION.State
              ) As TSPL_SHIP_TO_LOCATION ON TSPL_SHIP_TO_LOCATION.Ship_To_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Ship_To_Location 
              LEFT OUTER JOIN TSPL_DISTRIBUTOR_COMMISSION_DETAIL ON TSPL_DISTRIBUTOR_COMMISSION_DETAIL.pk_id = " + TSPL_SD_SALE_RETURN_detail + ".Distributor_Commission_PKID Full 
              join (
                select 
                  DOCUMENT_CODE, 
                  Item_Code as Scheme_Item_Code, 
                  SUM(Qty) AS SUB_QTY, 
                  SUM(Crate) AS schemeInCrates 
                from 
                  " + TSPL_SD_SALE_RETURN_detail + " as inn 
                where 
                  DOCUMENT_CODE in ('" + StrCode + "') 
                  and inn.Scheme_Item = 'Y' 
                group by 
                  DOCUMENT_CODE, 
                  Item_Code
              ) TSPL_SD_SALE_RETURN_DETAIL_Sub on TSPL_SD_SALE_RETURN_DETAIL_sub.DOCUMENT_CODE = " + TSPL_SD_SALE_RETURN_detail + ".DOCUMENT_CODE 
              and TSPL_SD_SALE_RETURN_DETAIL_sub.Scheme_Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              LEFT OUTER JOIN TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id = " + TSPL_SD_SALE_RETURN_HEAD + ".Vehicle_Code 
              left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code = " + TSPL_SD_SALE_RETURN_detail + ".Price_code 
              and TSPL_ITEM_PRICE_MASTER.Location_Code = " + TSPL_SD_SALE_RETURN_detail + ".Location 
              and TSPL_ITEM_PRICE_MASTER.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
              and len(
                isnull(
                  " + TSPL_SD_SALE_RETURN_detail + ".Price_code, 
                  ''
                )
              )> 0 
              left outer join (
                select 
                  Document_Code, 
                  Parent_Line_No, 
                  STRING_AGG(
                    Batch_No, 
                    CHAR(10)
                  ) as Batch_No, 
                  STRING_AGG(
                    Qty, 
                    CHAR(10)
                  ) as Batch_Qty 
                from 
                  (
                    SELECT 
                      Document_Code, 
                      Batch_No, 
                      Qty, 
                      Parent_Line_No 
                    FROM 
                      TSPL_BATCH_ITEM 
                    WHERE 
                      TSPL_BATCH_ITEM.Document_Type = 'PS-SR'
                  ) x 
                group by 
                  Document_Code, 
                  Parent_Line_No
              ) TabBatch On TabBatch.Document_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code 
              And TabBatch.Parent_Line_No = " + TSPL_SD_SALE_RETURN_detail + ".Line_No 
            where 
              2 = 2 
              And " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code in ('" + StrCode + "') 
              And exists (
                select 
                  1 
                from 
                  (
                    select 
                      " + TSPL_SD_SALE_RETURN_detail + ".Item_Code, 
                      " + TSPL_SD_SALE_RETURN_detail + ".DOCUMENT_CODE, 
                      " + TSPL_SD_SALE_RETURN_detail + ".Line_No 
                    from 
                      " + TSPL_SD_SALE_RETURN_detail + " 
                      Left OUTER JOIN " + TSPL_SD_SALE_RETURN_HEAD + " ON " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code = " + TSPL_SD_SALE_RETURN_detail + ".DOCUMENT_CODE 
                      Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
                      And TSPL_ITEM_UOM_DETAIL.UOM_Code = " + TSPL_SD_SALE_RETURN_detail + ".Unit_code 
                      Left OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
                      Left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Comp_Code 
                      Left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Customer_Code 
                      Left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = " + TSPL_SD_SALE_RETURN_HEAD + ".Bill_To_Location 
                      Left OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code = TSPL_LOCATION_MASTER.State Full 
                      Join(
                        select 
                          DOCUMENT_CODE, 
                          Item_Code As Scheme_Item_Code, 
                          SUM(Qty) As SUB_QTY, 
                          SUM(Crate) AS schemeInCrates 
                        From 
                          " + TSPL_SD_SALE_RETURN_detail + " As inn 
                        Where 
                          DOCUMENT_CODE In ('" + StrCode + "') 
                          And inn.Scheme_Item = 'Y' 
                        group by 
                          DOCUMENT_CODE, 
                          Item_Code
                      ) TSPL_SD_SALE_RETURN_DETAIL_Sub on TSPL_SD_SALE_RETURN_DETAIL_sub.DOCUMENT_CODE = " + TSPL_SD_SALE_RETURN_HEAD + ".DOCUMENT_CODE 
                      and TSPL_SD_SALE_RETURN_DETAIL_sub.Scheme_Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".Item_Code 
                    where 
                      2 = 2 
                      And " + TSPL_SD_SALE_RETURN_HEAD + ".Document_Code In ('" + StrCode + "') 
                      And " + TSPL_SD_SALE_RETURN_detail + ".Scheme_Item = 'N'
                  ) xx 
                where 
                  xx.Item_Code = " + TSPL_SD_SALE_RETURN_detail + ".item_CODE 
                  And xx.DOCUMENT_CODE = " + TSPL_SD_SALE_RETURN_detail + ".DOCUMENT_CODE
              )
          ) as final 
          Left outer join (
            Select 
              Item_Code, 
              max([CATEGORY RM]) As [CATEGORY RM], 
              max([BRAND]) As [BRAND], 
              max([SUB BRAND]) As [SUB BRAND], 
              max([DESCRP]) As [DESCRP], 
              max([PACK]) As [PACK], 
              max([PACK SIZE]) As [PACK SIZE], 
              max([CATEGORY OT]) As [CATEGORY OT], 
              max([CATEGORY FA]) As [CATEGORY FA], 
              max([P TYPE]) As [P TYPE], 
              max([L TYPE]) As [L TYPE], 
              max([JW]) As [JW], 
              max([SCRAP]) As [SCRAP], 
              max([CATEGORY RMDESC]) As [CATEGORY RMDESC], 
              max([BRANDDESC]) As [BRANDDESC], 
              max([SUB BRANDDESC]) As [SUB BRANDDESC], 
              max([DESCRPDESC]) As [DESCRPDESC], 
              max([PACKDESC]) As [PACKDESC], 
              max([PACK SIZEDESC]) As [PACK SIZEDESC], 
              max([CATEGORY OTDESC]) As [CATEGORY OTDESC], 
              max([CATEGORY FADESC]) As [CATEGORY FADESC], 
              max([P TYPEDESC]) As [P TYPEDESC], 
              max([L TYPEDESC]) As [L TYPEDESC], 
              max([JWDESC]) As [JWDESC], 
              max([SCRAPDESC]) As [SCRAPDESC] 
            from 
              (
                Select 
                  * 
                from 
                  (
                    Select 
                      TSPL_ITEM_MASTER.Item_Code, 
                      TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code, 
                      TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code + 'DESC' as Item_Category_CodeDesc, 
                      TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values, 
                      TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc 
                    From 
                      TSPL_ITEM_MASTER 
                      Left outer join TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code 
                      Left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE = TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code 
                      And TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE = TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values 
                    where 
                      2 = 3
                  ) xx Pivot(
                    max(Item_Cagetory_Values) For Item_Category_Code In (
                      [CATEGORY RM], [BRAND], [SUB BRAND], 
                      [DESCRP], [PACK], [PACK SIZE], [CATEGORY OT], 
                      [CATEGORY FA], [P TYPE], [L TYPE], [JW], 
                      [SCRAP]
                    )
                  ) Pivt Pivot (
                    max(Category_Value_Desc) For Item_Category_CodeDesc In (
                      [CATEGORY RMDESC], [BRANDDESC], [SUB BRANDDESC], 
                      [DESCRPDESC], [PACKDESC], [PACK SIZEDESC], 
                      [CATEGORY OTDESC], [CATEGORY FADESC], 
                      [P TYPEDESC], [L TYPEDESC], [JWDESC], 
                      [SCRAPDESC]
                    )
                  ) Pivt1
              ) xxx 
            group by 
              Item_Code
          ) as tbl_Brand on tbl_Brand.Item_Code = final.item_Code
      ) AS Main_Final 
      left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code = Main_Final.comp_code
  ) Final"
            Dim frmCRV As New frmCrystalReportViewer()
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptSaleReturnInvoiceBKN", "Sale Return", clsCommon.GetPrintDate(txtDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                Else
                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptSaleReturnInvoice", "Sale Return", clsCommon.GetPrintDate(txtDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())

                End If

                'If dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
                '    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptSaleReturnInvoiceBKN", "Sale Return", clsCommon.GetPrintDate(txtDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'Else
                '    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptSaleReturnNonTaxableInvoiceBKN", "Sale Return", clsCommon.GetPrintDate(txtDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'End If
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceGNG", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceSKR", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJDH", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBHR", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal AndAlso dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceALW1", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceALW1", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceSWM1", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'Else
                '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                'End If
                frmCRV = Nothing
            Else
                Throw New Exception("Data Not Found!")
            End If

        End If
        Return True
    End Function
    Public Function SaveData(ByVal obj As clsDSSalesReturnHead, ByVal isNewEntry As Boolean) As Boolean
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
    Public Function SaveData(ByVal obj As clsDSSalesReturnHead, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim Trans_type_Str As String = ""
        Dim isSaved As Boolean = True
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Trans_type_Str = clsCommon.myCstr(obj.Trans_type) + "-SR"
        Try
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleSaleDairy, clsUserMgtCode.frmSaleReturndairy, obj.Bill_To_Location, obj.Document_Date, trans)
            '' Anubhooti 06-Sep-2014 BM00000003735
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Sales And Distribution", "Sale Return", obj.Bill_To_Location, obj.Document_Date, trans)
            ''
            clsSerializeInvenotry.DeleteData("Sale Return", obj.Document_Code, trans)
            clsBatchInventory.DeleteData(Trans_type_Str, obj.Document_Code, trans)
            'If Not isNewEntry Then
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)
            ' End If

            Dim qry As String = "delete from TSPL_SD_SALE_RETURN_DETAIL where Document_Code='" & obj.Document_Code & "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_SD_SALE_RETURN_BOOKING_DETAIL where Document_Code='" & obj.Document_Code & "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            'Dim strDocNo As String = ""
            If isNewEntry Then
                If obj.Is_Cancelled = 1 Then
                    'If obj.Trans_type = "PS" And obj.Screen_Type = "DS" Then
                    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmDairySaleReturn, clsDocTransactionType.SaleReturnCancel, obj.Bill_To_Location)
                    'Else
                    '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleReturnProductSale, clsDocTransactionType.SaleReturnCancel, obj.Bill_To_Location)
                    'End If
                Else
                    'If obj.Trans_type = "PS" And obj.Screen_Type = "DS" Then
                    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmDairySaleReturn, clsDocTransactionType.NA, obj.Bill_To_Location)
                    'Else
                    '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleReturnProductSale, clsDocTransactionType.NA, obj.Bill_To_Location)
                    'End If
                End If
                'If clsCommon.CompairString(obj.Item_Type, "F") = CompairStringResult.Equal Then
                '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleReturn, clsDocTransactionType.SNQuotationFinishedGoods, obj.Bill_To_Location)
                'Else
                '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleReturn, clsDocTransactionType.SNQuotationOther, obj.Bill_To_Location)
                'End If
            End If
            If (clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "TotalCAN", obj.TotalCAN)
            clsCommon.AddColumnsForChange(coll, "ShippedCAN", obj.ShippedCAN)
            clsCommon.AddColumnsForChange(coll, "CrateQty", obj.CrateQty)
            clsCommon.AddColumnsForChange(coll, "Is_Taxable", obj.Is_Taxable)
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Customer_Code", obj.Customer_Code)
            clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Internal", IIf(obj.Is_Internal, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Ref_No", obj.Ref_No)
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            clsCommon.AddColumnsForChange(coll, "Inv_No", obj.Inv_No)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Bill_To_Location", obj.Bill_To_Location)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Location", obj.Ship_To_Location)
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
            clsCommon.AddColumnsForChange(coll, "Total_Amt", obj.Total_Amt)
            clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
            clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)
            clsCommon.AddColumnsForChange(coll, "PROJECT_ID", obj.PROJECT_ID, True)
            clsCommon.AddColumnsForChange(coll, "Sub_Location_code", obj.Sub_Location_code, True)

            If clsCommon.myLen(obj.Due_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Due_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Carrier", obj.Carrier)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Code", obj.Vehicle_Code)
            clsCommon.AddColumnsForChange(coll, "VehicleNo", obj.VehicleNo)
            clsCommon.AddColumnsForChange(coll, "GRNo", obj.GRNo)
            clsCommon.AddColumnsForChange(coll, "GENo", obj.GENo)
            clsCommon.AddColumnsForChange(coll, "Dept", obj.Dept)
            clsCommon.AddColumnsForChange(coll, "Dept_Desc", obj.Dept_Desc)
            clsCommon.AddColumnsForChange(coll, "Item_Type", obj.Item_Type)
            clsCommon.AddColumnsForChange(coll, "Against_Invoice_No", obj.Against_Invoice_No, True)
            clsCommon.AddColumnsForChange(coll, "isMultipleReturn", obj.isMultipleReturn)
            If obj.GEDate.HasValue Then
                clsCommon.AddColumnsForChange(coll, "GEDate", clsCommon.GetPrintDate(obj.GEDate, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "GEDate", Nothing, True)
            End If
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
            clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
            clsCommon.AddColumnsForChange(coll, "Challan_Date", clsCommon.GetPrintDate(obj.Challan_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Inv_Date", clsCommon.GetPrintDate(obj.Inv_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Salesman_Code", obj.Salesman_Code, True)
            clsCommon.AddColumnsForChange(coll, "Salesman_Name", obj.Salesman_Name)

            '' currencyconversion
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", obj.ApplicableFrom, True)
            '' End currencyconversion

            clsCommon.AddColumnsForChange(coll, "Price_Code", obj.Price_Code)
            clsCommon.AddColumnsForChange(coll, "Route_No", obj.Route_No)
            clsCommon.AddColumnsForChange(coll, "Route_Desc", obj.Route_Desc)
            clsCommon.AddColumnsForChange(coll, "TotCashDiscAmt", obj.TotCashDiscAmt)
            clsCommon.AddColumnsForChange(coll, "Invoice_Type", obj.Invoice_Type)
            clsCommon.AddColumnsForChange(coll, "Price_Group_Code", obj.Price_Group_Code)
            clsCommon.AddColumnsForChange(coll, "Cust_PO_No", obj.Cust_PO_No)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Per", obj.HeadDisc_Per)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_PerAmt", obj.HeadDisc_PerAmt)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Amt", obj.HeadDisc_Amt)
            clsCommon.AddColumnsForChange(coll, "Return_Type", obj.Return_Type)
            clsCommon.AddColumnsForChange(coll, "Shift_Type", obj.Shift_Type, True)
            clsCommon.AddColumnsForChange(coll, "Damage_Type", obj.Damage_Type)
            clsCommon.AddColumnsForChange(coll, "CrateMan", obj.CrateMan)
            clsCommon.AddColumnsForChange(coll, "jaali", obj.jaali)
            clsCommon.AddColumnsForChange(coll, "Box", obj.Box)
            clsCommon.AddColumnsForChange(coll, "Gate_ReturnNo", obj.Gate_ReturnNo)
            clsCommon.AddColumnsForChange(coll, "Is_Cancelled", obj.Is_Cancelled)
            clsCommon.AddColumnsForChange(coll, "IsSampling", obj.IsSampling)
            'clsCommon.AddColumnsForChange(coll, "Total_Comm_Amt", obj.Total_Comm_Amt)
            clsCommon.AddColumnsForChange(coll, "Distributor_Commission_TotalAmt", obj.Distributor_Commission_TotalAmt)
            clsCommon.AddColumnsForChange(coll, "Transporter_Commission_TotalAmt", obj.Transporter_Commission_TotalAmt)
            clsCommon.AddColumnsForChange(coll, "Security_TotalAmt", obj.Security_TotalAmt)
            clsCommon.AddColumnsForChange(coll, "Is_Without_Invoice", obj.Is_Without_Invoice)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Trans_Type", obj.Trans_type)
                clsCommon.AddColumnsForChange(coll, "Screen_Type", obj.Screen_Type)
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_RETURN_HEAD", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_RETURN_HEAD", OMInsertOrUpdate.Update, "TSPL_SD_SALE_RETURN_HEAD.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            isSaved = isSaved AndAlso clsDSSalesReturnDetail.SaveData(obj.Document_Code, obj.Document_Date, Arr, obj.Trans_type, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)

            If obj.Booth_Arr IsNot Nothing AndAlso obj.Booth_Arr.Count > 0 Then
                isSaved = isSaved AndAlso clsDSSalesReturnBookingDetail.SaveData(obj.Document_Code, obj.Booth_Arr, trans)

            End If
            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_Code, obj.arrCustomFields, trans)
            '''' to save item weight unit
            qry = "update TSPL_SD_SALE_RETURN_DETAIL set Weight_UOM= (select Weight_UOM from TSPL_ITEM_MASTER where Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code)  where Document_Code='" + obj.Document_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            '''' 

            'Save Cancel Ticket No- ALF/16/05/18-000064
            If obj.Is_Cancelled = 1 Then
                qry = "update TSPL_SD_SALE_INVOICE_HEAD set CancelFlag='Y'  where Document_Code='" + obj.Against_Invoice_No + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            'Save Cancel Ticket No- ALF/16/05/18-000064

            isSaved = isSaved AndAlso clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "Document_Code", obj.Document_Code, "TSPL_SD_SALE_RETURN_HEAD", trans)
            'If isSaved Then
            '    trans.Commit()
            'End If
        Catch err As Exception
            'trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function
    Public Shared Function getGateReturnFinder(ByVal whrcls As String, ByVal curGE_CODE As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " select DOC_CODE as Code,Transfer_Date as Date,Cust_Code as Customer,From_Location_Code as Location from TSPL_GATE_ENTRY_RETURN_PS "
        str = clsCommon.ShowSelectForm("GateEntryRefRetFnd", qry, "Code", whrcls, curGE_CODE, "Code", isButtonClicked)
        Return str
    End Function
    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType) As clsDSSalesReturnHead
        Return GetData(strDocumentNo, NavType, Nothing)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsDSSalesReturnHead
        Dim obj As clsDSSalesReturnHead = Nothing
        Dim Trans_type_Str As String = ""
        Dim Trans_type As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Trans_Type from TSPL_SD_SALE_RETURN_HEAD where document_code='" + strPONo + "'", trans))
        Dim qry As String = "SELECT TSPL_SD_SALE_RETURN_HEAD.Sub_Location_code,TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount,TSPL_SD_SALE_RETURN_HEAD.TotalCAN,TSPL_SD_SALE_RETURN_HEAD.ShippedCAN, TSPL_SD_SALE_RETURN_HEAD.CrateQty,TSPL_SD_SALE_RETURN_HEAD.Is_Taxable,TSPL_SD_SALE_RETURN_HEAD.Gate_ReturnNo,TSPL_SD_SALE_RETURN_HEAD.CrateMan,TSPL_SD_SALE_RETURN_HEAD.Jaali,TSPL_SD_SALE_RETURN_HEAD.Box , TSPL_SD_SALE_RETURN_HEAD.Trans_Type,TSPL_SD_SALE_RETURN_HEAD.Return_Type,TSPL_SD_SALE_RETURN_HEAD.Shift_Type,TSPL_SD_SALE_RETURN_HEAD.Damage_Type,TSPL_SD_SALE_RETURN_HEAD.HeadDisc_PerAmt,TSPL_SD_SALE_RETURN_HEAD.Cust_PO_No,TSPL_SD_SALE_RETURN_HEAD.VehicleNo, TSPL_SD_SALE_RETURN_HEAD.Vehicle_Code,TSPL_SD_SALE_RETURN_HEAD.price_group_code , " &
        "TSPL_SD_SALE_RETURN_HEAD.Invoice_Type,TSPL_SD_SALE_RETURN_HEAD.HeadDisc_Per,TSPL_SD_SALE_RETURN_HEAD.HeadDisc_Amt,TSPL_SD_SALE_RETURN_HEAD.TotCashDiscAmt,TSPL_SD_SALE_RETURN_HEAD.Route_No,TSPL_SD_SALE_RETURN_HEAD.Route_Desc,TSPL_SD_SALE_RETURN_HEAD.Price_Code, TSPL_SD_SALE_RETURN_HEAD.Document_Code,TSPL_SD_SALE_RETURN_HEAD.Document_Date,TSPL_SD_SALE_RETURN_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_SD_SALE_RETURN_HEAD.Status,TSPL_SD_SALE_RETURN_HEAD.On_Hold,TSPL_SD_SALE_RETURN_HEAD.Ref_No,TSPL_SD_SALE_RETURN_HEAD.Description,TSPL_SD_SALE_RETURN_HEAD.Remarks,TSPL_SD_SALE_RETURN_HEAD.Tax_Group,TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location,TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location,TSPL_SD_SALE_RETURN_HEAD.TAX1,TSPL_SD_SALE_RETURN_HEAD.TAX1_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX1_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX2,TSPL_SD_SALE_RETURN_HEAD.TAX2_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX2_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX3,TSPL_SD_SALE_RETURN_HEAD.TAX3_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX3_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX3_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX4,TSPL_SD_SALE_RETURN_HEAD.TAX4_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX4_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX4_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX5,TSPL_SD_SALE_RETURN_HEAD.TAX5_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX5_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX5_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX6,TSPL_SD_SALE_RETURN_HEAD.TAX6_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX6_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX6_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX7,TSPL_SD_SALE_RETURN_HEAD.TAX7_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX7_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX7_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX8,TSPL_SD_SALE_RETURN_HEAD.TAX8_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX8_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX8_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX9,TSPL_SD_SALE_RETURN_HEAD.TAX9_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX9_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX9_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX10,TSPL_SD_SALE_RETURN_HEAD.TAX10_Rate,TSPL_SD_SALE_RETURN_HEAD.TAX10_Amt,TSPL_SD_SALE_RETURN_HEAD.TAX10_Base_Amt,TSPL_SD_SALE_RETURN_HEAD.Discount_Base,TSPL_SD_SALE_RETURN_HEAD.Discount_Amt,TSPL_SD_SALE_RETURN_HEAD.Amount_Less_Discount,TSPL_SD_SALE_RETURN_HEAD.Total_Tax_Amt,TSPL_SD_SALE_RETURN_HEAD.Comments,TSPL_SD_SALE_RETURN_HEAD.Comp_Code,TSPL_SD_SALE_RETURN_HEAD.Terms_Code,TSPL_SD_SALE_RETURN_HEAD.Due_Date ,TSPL_LOCATION_MASTER.Location_Desc as BillToLocationName,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as ShipToLocationName,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as TaxGroupName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_SD_SALE_RETURN_HEAD.Posting_Date,TSPL_SD_SALE_RETURN_HEAD.Total_Amt,TSPL_SD_SALE_RETURN_HEAD.Carrier,TSPL_SD_SALE_RETURN_HEAD.VehicleNo,TSPL_SD_SALE_RETURN_HEAD.GRNo,TSPL_SD_SALE_RETURN_HEAD.GENo,TSPL_SD_SALE_RETURN_HEAD.GEDate, TSPL_SD_SALE_RETURN_HEAD.Dept,TSPL_SD_SALE_RETURN_HEAD.Dept_Desc,TSPL_SD_SALE_RETURN_HEAD.Item_Type,TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No ,TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code1,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name1,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt1,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code2,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name2,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt2,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code3,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name3,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt3,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code4,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name4,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt4,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code5,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name5,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt5,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code6,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name6,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt6,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code7,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name7,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt7,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code8,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name8,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt8,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code9 ,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name9,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt9 ,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Code10 ,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name10,TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt10,TSPL_SD_SALE_RETURN_HEAD.Total_Add_Charge,TSPL_SD_SALE_RETURN_HEAD.Tax_Calculation_Type,TSPL_SD_SALE_RETURN_HEAD.Challan_No, TSPL_SD_SALE_RETURN_HEAD.Challan_Date, TSPL_SD_SALE_RETURN_HEAD.Inv_Date,TSPL_SD_SALE_RETURN_HEAD.Inv_No,TSPL_SD_SALE_RETURN_HEAD.Is_Internal ,TSPL_SD_SALE_RETURN_HEAD.Salesman_Code ,TSPL_SD_SALE_RETURN_HEAD.Salesman_Name,"
        qry += " TSPL_SD_SALE_RETURN_HEAD.CURRENCY_CODE,TSPL_SD_SALE_RETURN_HEAD.CONVRATE,TSPL_SD_SALE_RETURN_HEAD.APPLICABLEFROM,TSPL_SD_SALE_RETURN_HEAD.PROJECT_ID ,TSPL_SD_SALE_RETURN_HEAD.Is_Cancelled,TSPL_SD_SALE_RETURN_HEAD.Is_Without_Invoice,TSPL_SD_SALE_RETURN_HEAD.isMultipleReturn "
        qry += " FROM TSPL_SD_SALE_RETURN_HEAD"
        qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location "
        qry += " left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location "
        qry += " left outer join  TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code= TSPL_SD_SALE_RETURN_HEAD.Tax_Group "
        qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_SD_SALE_RETURN_HEAD.Terms_Code "
        qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_RETURN_HEAD.Customer_Code where TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + Trans_Type + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' "
        Dim whrCls As String = ""

        '-------richa 12/08/2014 Ticket No. BM00000003242---------
        Dim strwherecls As String = ""
        If clsCommon.CompairString(clsCommon.myCstr(NavType).ToUpper(), "CURRENT") <> CompairStringResult.Equal Then
            strwherecls = FrmMainTranScreen.CustomerPermission()
        End If


        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 And clsCommon.myLen(strwherecls) > 0 Then
            whrCls = " AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + strwherecls + ") "
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrCls = " AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        ElseIf clsCommon.myLen(strwherecls) > 0 Then
            whrCls = " AND TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + strwherecls + ")"
        End If
        '-----------------------------------------------------
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = (select MIN(Document_Code) from TSPL_SD_SALE_RETURN_HEAD WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Last
                qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = (select Max(Document_Code) from TSPL_SD_SALE_RETURN_HEAD WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Current
                qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = '" + strPONo + "'"
            Case NavigatorType.Next
                qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = (select Min(Document_Code) from TSPL_SD_SALE_RETURN_HEAD where Document_Code>'" + strPONo + "' " + whrCls + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = (select Max(Document_Code) from TSPL_SD_SALE_RETURN_HEAD where Document_Code<'" + strPONo + "' " + whrCls + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsDSSalesReturnHead()
            obj.TotalCAN = clsCommon.myCdbl(dt.Rows(0)("TotalCAN"))
            obj.ShippedCAN = clsCommon.myCdbl(dt.Rows(0)("ShippedCAN"))
            obj.CrateQty = clsCommon.myCdbl(dt.Rows(0)("CrateQty"))
            obj.Is_Cancelled = clsCommon.myCdbl(dt.Rows(0)("Is_Cancelled"))
            obj.Is_Without_Invoice = clsCommon.myCdbl(dt.Rows(0)("Is_Without_Invoice"))
            obj.Is_Taxable = clsCommon.myCdbl(dt.Rows(0)("Is_Taxable"))
            obj.Gate_ReturnNo = clsCommon.myCstr(dt.Rows(0)("Gate_ReturnNo"))
            obj.CrateMan = clsCommon.myCdbl(dt.Rows(0)("CrateMan"))
            obj.jaali = clsCommon.myCdbl(dt.Rows(0)("jaali"))
            obj.Box = clsCommon.myCdbl(dt.Rows(0)("Box"))

            obj.Trans_type = clsCommon.myCstr(dt.Rows(0)("Trans_type"))
            Trans_type_Str = clsCommon.myCstr(dt.Rows(0)("Trans_type")) + "-SR"
            obj.Invoice_Type = clsCommon.myCstr(dt.Rows(0)("Invoice_Type"))
            obj.Cust_PO_No = clsCommon.myCstr(dt.Rows(0)("Cust_PO_No"))
            obj.Price_Group_Code = clsCommon.myCstr(dt.Rows(0)("Price_Group_Code"))
            obj.Price_Code = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
            obj.Sub_Location_code = clsCommon.myCstr(dt.Rows(0)("Sub_Location_code"))
            obj.Route_No = clsCommon.myCstr(dt.Rows(0)("Route_No"))
            obj.Route_Desc = clsCommon.myCstr(dt.Rows(0)("Route_Desc"))
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Customer_Code = clsCommon.myCstr(dt.Rows(0)("Customer_Code"))
            obj.Customer_Name = clsCommon.myCstr(dt.Rows(0)("Customer_Name"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Is_Internal = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Internal")) = 1, True, False)
            obj.Ref_No = clsCommon.myCstr(dt.Rows(0)("Ref_No"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.isMultipleReturn = clsCommon.myCdbl(dt.Rows(0)("isMultipleReturn"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Bill_To_Location = clsCommon.myCstr(dt.Rows(0)("Bill_To_Location"))
            obj.Ship_To_Location = clsCommon.myCstr(dt.Rows(0)("Ship_To_Location"))
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
            obj.Total_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Amt"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            obj.Comments = clsCommon.myCstr(dt.Rows(0)("Comments"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            obj.BillToLocationName = clsCommon.myCstr(dt.Rows(0)("BillToLocationName"))
            obj.ShipToLocationName = clsCommon.myCstr(dt.Rows(0)("ShipToLocationName"))
            obj.TaxGroupName = clsCommon.myCstr(dt.Rows(0)("TaxGroupName"))
            obj.TermsName = clsCommon.myCstr(dt.Rows(0)("TermsName"))
            obj.PROJECT_ID = clsCommon.myCstr(dt.Rows(0)("PROJECT_ID"))
            obj.Vehicle_Code = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
            obj.VehicleNo = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
            If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
            End If


            obj.Challan_No = clsCommon.myCdbl(dt.Rows(0)("Challan_No"))
            obj.Carrier = clsCommon.myCstr(dt.Rows(0)("Carrier"))
            obj.Vehicle_Code = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
            obj.VehicleNo = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
            obj.GRNo = clsCommon.myCstr(dt.Rows(0)("GRNo"))
            obj.GENo = clsCommon.myCstr(dt.Rows(0)("GENo"))
            If dt.Rows(0)("GEDate") IsNot DBNull.Value Then
                obj.GEDate = clsCommon.myCDate(dt.Rows(0)("GEDate"))
            End If




            obj.Dept = clsCommon.myCstr(dt.Rows(0)("Dept"))
            obj.Dept_Desc = clsCommon.myCstr(dt.Rows(0)("Dept_Desc"))
            obj.Item_Type = clsCommon.myCstr(dt.Rows(0)("Item_Type"))

            obj.Against_Invoice_No = clsCommon.myCstr(dt.Rows(0)("Against_Invoice_No"))


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

            obj.Salesman_Code = clsCommon.myCstr(dt.Rows(0)("Salesman_Code"))
            obj.Salesman_Name = clsCommon.myCstr(dt.Rows(0)("Salesman_Name"))
            obj.HeadDisc_Per = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Per"))
            obj.HeadDisc_PerAmt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_PerAmt"))
            obj.HeadDisc_Amt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Amt"))
            obj.Return_Type = clsCommon.myCstr(dt.Rows(0)("Return_Type"))
            obj.Shift_Type = clsCommon.myCstr(dt.Rows(0)("Shift_Type"))
            obj.Damage_Type = clsCommon.myCstr(dt.Rows(0)("Damage_Type"))

            '' CURRENCYCONVERSION 
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            '' END CURRENCYCONVERSION 

            qry = "SELECT isnull(TSPL_SD_SALE_RETURN_DETAIL.Delivery_Code,'') as Delivery_Code,TSPL_SD_SALE_RETURN_DETAIL.CAN,  TSPL_SD_SALE_RETURN_DETAIL.Crate,TSPL_SD_SALE_RETURN_DETAIL.InvoiceCashScheme_Code,  TSPL_SD_SALE_RETURN_DETAIL.ItemwiseTaxCode,TSPL_SD_SALE_RETURN_DETAIL.Alter_UnitQty,TSPL_SD_SALE_RETURN_DETAIL.Rate_UnitQty,TSPL_SD_SALE_RETURN_DETAIL.InvoiceScheme_Code,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Amount,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Type,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Pers,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Code,TSPL_SD_SALE_RETURN_DETAIL.TAX_PAID ,TSPL_SD_SALE_RETURN_DETAIL.Item_Group, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Scheme_Item_UOM,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Qty,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Item_Code,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Type,TSPL_SD_SALE_RETURN_DETAIL.Alternate_UOM,TSPL_SD_SALE_RETURN_DETAIL.RATE_UOM , TSPL_SD_SALE_RETURN_DETAIL.OrgUnit_code,TSPL_SD_SALE_RETURN_DETAIL.Is_Mannual_Amt,TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Line_No, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Status,TSPL_SD_SALE_RETURN_DETAIL.Row_Type,TSPL_SD_SALE_RETURN_DETAIL.status,TSPL_SD_SALE_RETURN_DETAIL.Item_Code, " &
            "TSPL_ITEM_MASTER.Item_Desc,TSPL_SD_SALE_RETURN_DETAIL.Qty,TSPL_SD_SALE_RETURN_DETAIL.Free_Qty,TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code,TSPL_SD_SALE_RETURN_DETAIL.Balance_Qty,TSPL_SD_SALE_RETURN_DETAIL.Unit_code, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Location,TSPL_SD_SALE_RETURN_DETAIL.Item_Cost,TSPL_SD_SALE_RETURN_DETAIL.TAX1,TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX2,TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX3,TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX4, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX5,TSPL_SD_SALE_RETURN_DETAIL.TAX5_Rate, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX6,TSPL_SD_SALE_RETURN_DETAIL.TAX6_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX7,TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX8, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX8_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX9,TSPL_SD_SALE_RETURN_DETAIL.TAX9_Rate, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX10,TSPL_SD_SALE_RETURN_DETAIL.TAX10_Rate,TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Amount,TSPL_SD_SALE_RETURN_DETAIL.Disc_Per,TSPL_SD_SALE_RETURN_DETAIL.Disc_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt, " &
            "TSPL_LOCATION_MASTER.Location_Desc as LocationName,TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX2_Base_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX4_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX5_Base_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX6_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX7_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX8_Base_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.TAX9_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX10_Base_Amt,TSPL_SD_SALE_RETURN_DETAIL.MRP , " &
            "TSPL_SD_SALE_RETURN_DETAIL.Batch_No,TSPL_SD_SALE_RETURN_DETAIL.MFG_Date,TSPL_SD_SALE_RETURN_DETAIL.Expiry_Date, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Specification,TSPL_SD_SALE_RETURN_DETAIL.Remarks,TSPL_SD_SALE_RETURN_DETAIL.Assessable, " &
            "TSPL_SD_SALE_RETURN_DETAIL.AssessableAmt,TSPL_SD_SALE_RETURN_DETAIL.DamageQty,TSPL_SD_SALE_RETURN_DETAIL.Return_Amount,TSPL_SD_SALE_RETURN_DETAIL.Damage_Amount, " &
             "TSPL_SD_SALE_RETURN_DETAIL.Scheme_Applicable,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Code, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Scheme_Item,TSPL_SD_SALE_RETURN_DETAIL.Item_Tax,TSPL_SD_SALE_RETURN_DETAIL.Total_MRP_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Total_Basic_Amt,TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,TSPL_SD_SALE_RETURN_DETAIL.Cust_Discount, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Total_Cust_Discount,TSPL_SD_SALE_RETURN_DETAIL.ActualRate,TSPL_SD_SALE_RETURN_DETAIL.Cust_DiscountQty, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Price_code,TSPL_SD_SALE_RETURN_DETAIL.Abatement_Per,TSPL_SD_SALE_RETURN_DETAIL.Abatement_Amt, " &
            "TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,TSPL_SD_SALE_RETURN_DETAIL.Item_Weight,TSPL_SD_SALE_RETURN_DETAIL.Price_Date, " &
            "TSPL_SD_SALE_RETURN_DETAIL.Bin_No,TSPL_SD_SALE_RETURN_DETAIL.vendor_code,TSPL_SD_SALE_RETURN_DETAIL.vendor_desc,TSPL_SD_SALE_RETURN_DETAIL.PrincipleCode,TSPL_SD_SALE_RETURN_DETAIL.PrincipleDesc,TSPL_SD_SALE_RETURN_DETAIL.TotalItem_Weight,TSPL_SD_SALE_RETURN_DETAIL.Conv_Factor,TSPL_SD_SALE_RETURN_DETAIL.Purchase_Cost,TSPL_SD_SALE_RETURN_DETAIL.OrgRate,  " &
            "TSPL_SD_SALE_RETURN_DETAIL.HeadDiscPer,TSPL_SD_SALE_RETURN_DETAIL.HeadDiscPerAmt,TSPL_SD_SALE_RETURN_DETAIL.Markup_On,TSPL_SD_SALE_RETURN_DETAIL.Markup_Percent,TSPL_SD_SALE_RETURN_DETAIL.Landing_Cost,TSPL_SD_SALE_RETURN_DETAIL.HeadDiscAmt,TSPL_SD_SALE_RETURN_DETAIL.CustDiscPer,TSPL_SD_SALE_RETURN_DETAIL.CasdDiscScheme_Code,TSPL_SD_SALE_RETURN_DETAIL.ActualUOM,TSPL_SD_SALE_RETURN_DETAIL.ActualQty,TSPL_SD_SALE_RETURN_DETAIL.ActualConvAmt,TSPL_SD_SALE_RETURN_DETAIL.Distributor_Commission_PKID,TSPL_SD_SALE_RETURN_DETAIL.Distributor_Commission_Rate,TSPL_SD_SALE_RETURN_DETAIL.Distributor_Commission_RateWithTax,TSPL_SD_SALE_RETURN_DETAIL.Distributor_Commission_Amt,TSPL_SD_SALE_RETURN_DETAIL.Transporter_Commission_Rate,TSPL_SD_SALE_RETURN_DETAIL.Transporter_Commission_Amt,TSPL_SD_SALE_RETURN_DETAIL.Security_Rate,TSPL_SD_SALE_RETURN_DETAIL.Security_Amt,TSPL_SD_SALE_RETURN_DETAIL.Transporter,TSPL_SD_SALE_RETURN_DETAIL.Price_with_tax "
            qry += " FROM TSPL_SD_SALE_RETURN_DETAIL "
            qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SD_SALE_RETURN_DETAIL.Location "
            qry += " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code"
            qry += " where TSPL_SD_SALE_RETURN_DETAIL.Document_Code='" + obj.Document_Code + "' ORDER BY TSPL_SD_SALE_RETURN_DETAIL.Line_No"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsDSSalesReturnDetail)
                Dim objTr As clsDSSalesReturnDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsDSSalesReturnDetail
                    objTr.Delivery_Code = clsCommon.myCstr(dr("Delivery_Code"))
                    objTr.CAN = clsCommon.myCdbl(dr("CAN"))
                    objTr.Crate = clsCommon.myCdbl(dr("Crate"))
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(dr("ItemwiseTaxCode"))
                    objTr.Alter_UnitQty = clsCommon.myCdbl(dr("Alter_UnitQty"))
                    objTr.Rate_UnitQty = clsCommon.myCdbl(dr("Rate_UnitQty"))
                    objTr.InvoiceCashScheme_Code = clsCommon.myCstr(dr("InvoiceCashScheme_Code"))
                    objTr.InvoiceScheme_Code = clsCommon.myCstr(dr("InvoiceScheme_Code"))
                    objTr.Cash_Scheme_Code = clsCommon.myCstr(dr("Cash_Scheme_Code"))
                    objTr.Cash_Scheme_Type = clsCommon.myCstr(dr("Cash_Scheme_Type"))
                    objTr.Cash_Scheme_Pers = clsCommon.myCdbl(dr("Cash_Scheme_Pers"))
                    objTr.Cash_Scheme_Amount = clsCommon.myCdbl(dr("Cash_Scheme_Amount"))
                    objTr.Scheme_Type = clsCommon.myCstr(dr("Scheme_Type"))
                    objTr.Scheme_Qty = clsCommon.myCdbl(dr("Scheme_Qty"))
                    objTr.Scheme_Item_UOM = clsCommon.myCstr(dr("Scheme_Item_UOM"))
                    objTr.Scheme_Item_Code = clsCommon.myCstr(dr("Scheme_Item_Code"))
                    objTr.Alternate_UOM = clsCommon.myCstr(dr("Alternate_UOM"))
                    objTr.RATE_UOM = clsCommon.myCstr(dr("RATE_UOM"))
                    objTr.Item_Group = clsCommon.myCstr(dr("Item_Group"))
                    objTr.TAX_PAID = clsCommon.myCstr(dr("TAX_PAID"))
                    objTr.Price_with_tax = clsCommon.myCdbl(dr("Price_with_tax"))

                    objTr.OrgUnit_code = clsCommon.myCstr(dr("OrgUnit_code"))
                    objTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objTr.Row_Type = clsCommon.myCstr(dr("Row_Type"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Status = Convert.ToInt32(clsCommon.myCdbl(dr("Status")))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.Qty = clsCommon.myCdbl(dr("Qty"))
                    objTr.DamageQty = clsCommon.myCdbl(dr("DamageQty"))
                    objTr.Return_Amount = clsCommon.myCdbl(dr("Return_Amount"))
                    objTr.Damage_Amount = clsCommon.myCdbl(dr("Damage_Amount"))
                    objTr.Free_Qty = clsCommon.myCdbl(dr("Free_Qty"))
                    objTr.Invoice_Code = clsCommon.myCstr(dr("Invoice_Code"))

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
                    objTr.Disc_Per = clsCommon.myCdbl(dr("Disc_Per"))
                    objTr.Disc_Amt = clsCommon.myCdbl(dr("Disc_Amt"))
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(dr("Amt_Less_Discount"))
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(dr("Total_Tax_Amt"))
                    objTr.Item_Net_Amt = clsCommon.myCdbl(dr("Item_Net_Amt"))


                    objTr.Is_Mannual_Amt = clsCommon.myCdbl(dr("Is_Mannual_Amt"))

                    ' ''objTr.Landed_Cost_Rate = clsCommon.myCdbl(dr("Landed_Cost_Rate"))
                    ' ''objTr.Landed_Cost_Amount = clsCommon.myCdbl(dr("Landed_Cost_Amount"))

                    objTr.MRP = clsCommon.myCdbl(dr("MRP"))
                    objTr.Assessable = clsCommon.myCdbl(dr("Assessable"))
                    objTr.AssessableAmt = clsCommon.myCdbl(dr("AssessableAmt"))
                    objTr.Batch_No = clsCommon.myCstr(dr("Batch_No"))
                    If dr("MFG_Date") IsNot DBNull.Value Then
                        objTr.MFG_Date = clsCommon.myCDate(dr("MFG_Date"))
                    End If
                    If dr("Expiry_Date") IsNot DBNull.Value Then
                        objTr.Expiry_Date = clsCommon.myCDate(dr("Expiry_Date"))
                    End If
                    objTr.Specification = clsCommon.myCstr(dr("Specification"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))


                    objTr.Scheme_Applicable = clsCommon.myCstr(dr("Scheme_Applicable"))
                    objTr.Scheme_Code = clsCommon.myCstr(dr("Scheme_Code"))
                    objTr.Scheme_Item = clsCommon.myCstr(dr("Scheme_Item"))
                    objTr.Item_Tax = clsCommon.myCdbl(dr("Item_Tax"))
                    objTr.Total_MRP_Amt = clsCommon.myCdbl(dr("Total_MRP_Amt"))
                    objTr.Total_Basic_Amt = clsCommon.myCdbl(dr("Total_Basic_Amt"))
                    objTr.Total_Disc_Amt = clsCommon.myCdbl(dr("Total_Disc_Amt"))
                    objTr.Cust_Discount = clsCommon.myCdbl(dr("Cust_Discount"))
                    objTr.Total_Cust_Discount = clsCommon.myCdbl(dr("Total_Cust_Discount"))
                    objTr.ActualRate = clsCommon.myCdbl(dr("ActualRate"))
                    objTr.Cust_DiscountQty = clsCommon.myCdbl(dr("Cust_DiscountQty"))
                    objTr.Price_code = clsCommon.myCstr(dr("Price_code"))
                    If IsDBNull(dr("Price_Date")).ToString() = "" Then
                        objTr.Price_Date = clsCommon.myCDate(dr("Price_Date"))
                    End If


                    objTr.Abatement_Per = clsCommon.myCdbl(dr("Abatement_Per"))
                    objTr.Abatement_Amt = clsCommon.myCdbl(dr("Abatement_Amt"))
                    objTr.FOC_Item = clsCommon.myCdbl(dr("FOC_Item"))
                    objTr.Markup_On = clsCommon.myCstr(dr("Markup_On"))
                    objTr.Markup_Percent = clsCommon.myCdbl(dr("Markup_Percent"))
                    objTr.Landing_Cost = clsCommon.myCdbl(dr("Landing_Cost"))
                    objTr.HeadDiscAmt = clsCommon.myCdbl(dr("HeadDiscAmt"))
                    objTr.CustDiscPer = clsCommon.myCdbl(dr("CustDiscPer"))
                    objTr.CasdDiscScheme_Code = clsCommon.myCstr(dr("CasdDiscScheme_Code"))
                    objTr.Item_Weight = clsCommon.myCdbl(dr("Item_Weight"))
                    objTr.TotalItem_Weight = clsCommon.myCdbl(dr("TotalItem_Weight"))
                    objTr.Conv_Factor = clsCommon.myCdbl(dr("Conv_Factor"))
                    objTr.Purchase_Cost = clsCommon.myCdbl(dr("Purchase_Cost"))
                    objTr.OrgRate = clsCommon.myCdbl(dr("OrgRate"))
                    objTr.Bin_No = clsCommon.myCstr(dr("Bin_No"))
                    objTr.PrincipleCode = clsCommon.myCstr(dr("PrincipleCode"))
                    objTr.PrincipleDesc = clsCommon.myCstr(dr("PrincipleDesc"))
                    objTr.vendor_code = clsCommon.myCstr(dr("vendor_code"))
                    objTr.vendor_desc = clsCommon.myCstr(dr("vendor_desc"))
                    objTr.HeadDiscPer = clsCommon.myCdbl(dr("HeadDiscPer"))
                    objTr.HeadDiscPerAmt = clsCommon.myCdbl(dr("HeadDiscPerAmt"))
                    objTr.ActuaQty = clsCommon.myCdbl(dr("ActualQty"))
                    objTr.ActualUOM = clsCommon.myCstr(dr("ActualUOM"))
                    objTr.ActualReturnQty = clsCommon.myCdbl(dr("ActualConvAmt"))
                    objTr.Distributor_Commission_PKID = clsCommon.myCstr(dr("Distributor_Commission_PKID"))
                    objTr.Distributor_Commission_Rate = clsCommon.myCdbl(dr("Distributor_Commission_Rate"))
                    objTr.Distributor_Commission_RateWithTax = clsCommon.myCdbl(dr("Distributor_Commission_RateWithTax"))
                    objTr.Distributor_Commission_Amt = clsCommon.myCdbl(dr("Distributor_Commission_Amt"))
                    objTr.Transporter_Commission_Rate = clsCommon.myCdbl(dr("Transporter_Commission_Rate"))
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(dr("Transporter_Commission_Amt"))
                    objTr.Security_Rate = clsCommon.myCdbl(dr("Security_Rate"))
                    objTr.Security_Amt = clsCommon.myCdbl(dr("Security_Amt"))
                    objTr.Transporter = clsCommon.myCdbl(dr("Transporter"))
                    objTr.arrSrItem = clsSerializeInvenotry.GetData("Sale Return", objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    objTr.arrBatchItem = clsBatchInventory.GetData(Trans_type_Str, objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr.Add(objTr)
                Next
            End If

            qry = " select TSPL_SD_SALE_RETURN_Booking_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_Booking_DETAIL.Booth_Code,TSPL_Customer_Master.Customer_Name as Booth_Name,TSPL_SD_SALE_RETURN_Booking_DETAIL.item_Code,TSPL_ITEM_MASTER.Short_Description as Item_Name,TSPL_SD_SALE_RETURN_Booking_DETAIL.Unit_Code,TSPL_SD_SALE_RETURN_Booking_DETAIL.Return_Qty from TSPL_SD_SALE_RETURN_Booking_DETAIL
 left join TSPL_Customer_Master on TSPL_Customer_Master.Cust_Code=TSPL_SD_SALE_RETURN_Booking_DETAIL.Booth_Code
 left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_Booking_DETAIL.item_Code
 where TSPL_SD_SALE_RETURN_Booking_DETAIL.Document_Code='" & obj.Document_Code & "'"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Booth_Arr = New List(Of clsDSSalesReturnBookingDetail)
                Dim objBoothTr As clsDSSalesReturnBookingDetail
                For Each dr As DataRow In dt.Rows
                    objBoothTr = New clsDSSalesReturnBookingDetail
                    objBoothTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objBoothTr.Booth_Code = clsCommon.myCstr(dr("Booth_Code"))
                    objBoothTr.Booth_Name = clsCommon.myCstr(dr("Booth_Name"))
                    objBoothTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objBoothTr.Item_Name = clsCommon.myCstr(dr("Item_Name"))
                    objBoothTr.Unit_Code = clsCommon.myCstr(dr("Unit_Code"))
                    objBoothTr.Return_Qty = clsCommon.myCstr(dr("Return_Qty"))
                    obj.Booth_Arr.Add(objBoothTr)
                Next

            End If
        End If

        Return obj
    End Function

    Public Shared Function GetOriginalQty(ByVal strMrnNo As String, ByVal strICode As String, ByVal strUOM As String, ByVal dblAssessable As Double, ByVal dblMRP As Double, ByVal trans As SqlTransaction) As DataTable
        Dim qry As String = "Select TSPL_MRN_DETAIL.MRN_No,(TSPL_MRN_DETAIL.MRN_Qty+ISNULL(TSPL_MRN_DETAIL.Leak_Qty,0) +ISNULL(TSPL_MRN_DETAIL.Burst_Qty,0)+ISNULL(TSPL_MRN_DETAIL.Short_Qty,0)) as MRN_Qty,TSPL_GRN_DETAIL.GRN_No,(TSPL_GRN_DETAIL.GRN_Qty+ISNULL(TSPL_GRN_DETAIL.Leak_Qty,0) +ISNULL(TSPL_GRN_DETAIL.Burst_Qty,0)+ISNULL(TSPL_GRN_DETAIL.Short_Qty,0)) as GRN_Qty, TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty from TSPL_MRN_DETAIL left outer join TSPL_GRN_DETAIL on TSPL_GRN_DETAIL.GRN_No=TSPL_MRN_DETAIL.GRN_Id and TSPL_GRN_DETAIL.Item_Code=TSPL_MRN_DETAIL.Item_Code and TSPL_GRN_DETAIL.Unit_code=TSPL_MRN_DETAIL.Unit_code and isnull(TSPL_GRN_DETAIL.Assessable,0)=isnull(TSPL_MRN_DETAIL.Assessable,0) and isnull(TSPL_GRN_DETAIL.Item_Code,0)=isnull(TSPL_MRN_DETAIL.Item_Code ,0) left outer join TSPL_PURCHASE_ORDER_DETAIL on TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No= TSPL_GRN_DETAIL.PO_Id and TSPL_PURCHASE_ORDER_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code and TSPL_PURCHASE_ORDER_DETAIL.Unit_code=  TSPL_GRN_DETAIL.Unit_code and isnull(TSPL_PURCHASE_ORDER_DETAIL.Assessable,0)=  isnull(TSPL_GRN_DETAIL.Assessable,0) and isnull(TSPL_PURCHASE_ORDER_DETAIL.MRP,0)=  isnull(TSPL_GRN_DETAIL.MRP,0) where TSPL_MRN_DETAIL.MRN_No='" + strMrnNo + "' and TSPL_MRN_DETAIL.Item_Code='" + strICode + "' and TSPL_MRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_MRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_MRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'"
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function
    Public Shared Function UpdateInventoryMovement(ByVal obj As clsDSSalesReturnHead, ByVal trans As SqlTransaction, Optional ByVal UpdateInventory As Boolean = False) As Boolean

        Dim isSaved As Boolean = True
        Dim Trans_type_Str As String = ""
        Trans_type_Str = clsCommon.myCstr(obj.Trans_type) + "-SR"
        If UpdateInventory = True Then
            clsDBFuncationality.ExecuteNonQuery("update tspl_batch_item set Against_Inv_Movement_Trans_Id=null where Document_Code='" & obj.Document_Code & "'", trans)
            clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & obj.Document_Code & "'", trans)
        End If
        Dim qry As String = ""

        Dim ArrLocationDetails As List(Of clsItemLocationDetails) = New List(Of clsItemLocationDetails)()
        Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)

        Dim strFirstItemCodeNonItemRowType As String = GetFirstItemCode(obj.Arr)
        Dim strRgpNo As String = Nothing
        Dim intCounter As Integer = 0
        For Each objTr As clsDSSalesReturnDetail In obj.Arr
            intCounter = intCounter + 1
            If clsCommon.CompairString(objTr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                Dim strItemType As String = clsItemMaster.GetItemType(objTr.Item_Code, trans)
                Dim strItemTypeToSave As String = ""
                If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                    strItemTypeToSave = "RM"
                ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                    strItemTypeToSave = "OT"
                ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                    strItemTypeToSave = "FT"
                Else
                    strItemTypeToSave = strItemType
                    'Throw New Exception("Item Type not found: " + strItemType)
                End If
                Dim objLocationDetails As New clsItemLocationDetails()
                Dim ConvFac As Double = clsItemMaster.GetConvertionFactor(objTr.Item_Code, objTr.Unit_code, trans)
                If ConvFac = 0 Then
                    Throw New Exception("Conversion Factor found zero for item :" + objTr.Item_Code + " and Uom:'" + objTr.Unit_code)
                End If

                '''' FOR cogs START here

                Dim objInv As clsSNInvoiceHead
                Dim arr As New List(Of String)


                Dim dblCogsBasicCost As Double
                objInv = clsSNInvoiceHead.GetData(objTr.Invoice_Code, NavigatorType.Current, "", trans)
                If Not objInv Is Nothing Then
                    For Each objInvDetail As clsSNInvoiceDetail In objInv.Arr
                        If objInvDetail.Item_Code = objTr.Item_Code Then
                            dblCogsBasicCost = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost " &
                            "when Costing_Method=3 then LIFO_Cost end)/sum(Qty) as COst from TSPL_INVENTORY_MOVEMENT left outer join " &
                            "TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on " &
                            "TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where " &
                            "Source_Doc_No='" & objInvDetail.Shipment_Code & "' And TSPL_INVENTORY_MOVEMENT.Item_Code='" & objTr.Item_Code & "'  and  TSPL_INVENTORY_MOVEMENT.MRP=" & objTr.MRP & " ", trans))



                        End If
                    Next
                End If
                '''' COGS ENDS HERE


                objLocationDetails.Item_Code = objTr.Item_Code
                objLocationDetails.Item_Desc = objTr.Item_Desc
                objLocationDetails.Location_Code = objTr.Location
                objLocationDetails.Location_Desc = objTr.LocationName
                objLocationDetails.Item_Qty = objTr.Qty
                objLocationDetails.Amount = objTr.Amount
                objLocationDetails.MRP = objTr.MRP * ConvFac
                If objTr.MFG_Date.HasValue Then
                    objLocationDetails.MFG_Date = objTr.MFG_Date
                End If
                objLocationDetails.Batch_No = objTr.Batch_No
                If objTr.Expiry_Date.HasValue Then
                    objLocationDetails.Expiry_Date = objTr.Expiry_Date
                End If
                objLocationDetails.ItemType = strItemTypeToSave
                ArrLocationDetails.Add(objLocationDetails)

                Dim objInventoryMovemnt As New clsInventoryMovement()
                objInventoryMovemnt.InOut = "I"
                objInventoryMovemnt.Location_Code = IIf(clsCommon.myLen(clsCommon.myCstr(obj.Sub_Location_code)) > 0, obj.Sub_Location_code, objTr.Location)

                objInventoryMovemnt.Cust_Code = obj.Customer_Code
                objInventoryMovemnt.Cust_Name = obj.Customer_Name

                objInventoryMovemnt.Item_Code = objTr.Item_Code
                objInventoryMovemnt.Item_Desc = objTr.Item_Desc
                'objInventoryMovemnt.Qty = objTr.Qty
                objInventoryMovemnt.Qty = objTr.ActuaQty
                objInventoryMovemnt.UOM = IIf(clsCommon.myLen(clsCommon.myCstr(objTr.ActualUOM)) > 0, objTr.ActualUOM, objTr.Unit_code)
                objInventoryMovemnt.Is_Scheme_Item = objTr.Scheme_Item
                ''richa only for scheme item 
                If clsCommon.CompairString(objTr.Scheme_Item, "Y") = CompairStringResult.Equal Then
                    If Not objInv Is Nothing Then
                        For Each objInvDetail As clsSNInvoiceDetail In objInv.Arr
                            If objInvDetail.Item_Code = objTr.Item_Code Then
                                If clsCommon.CompairString(objInvDetail.Scheme_Item, "Y") = CompairStringResult.Equal Then
                                    dblCogsBasicCost = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost " &
                            "when Costing_Method=3 then LIFO_Cost end)/sum(Qty) as COst from TSPL_INVENTORY_MOVEMENT left outer join " &
                            "TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on " &
                            "TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where " &
                            "Source_Doc_No='" & objInvDetail.Shipment_Code & "' And TSPL_INVENTORY_MOVEMENT.Item_Code='" & objTr.Item_Code & "' and TSPL_INVENTORY_MOVEMENT.is_Scheme_item='Y' and  TSPL_INVENTORY_MOVEMENT.MRP='" & objInvDetail.MRP & "' ", trans))



                                End If

                            End If
                        Next
                    End If
                End If

                If clsCommon.CompairString(objTr.Scheme_Item, "Y") = CompairStringResult.Equal Then
                    objInventoryMovemnt.Basic_Cost = dblCogsBasicCost * objTr.ActuaQty
                Else
                    objInventoryMovemnt.Basic_Cost = dblCogsBasicCost * objTr.Qty
                End If

                objInventoryMovemnt.Avg_Cost = objInventoryMovemnt.Basic_Cost
                objInventoryMovemnt.LIFO_Cost = objInventoryMovemnt.Basic_Cost
                objInventoryMovemnt.FIFO_Cost = objInventoryMovemnt.Basic_Cost
                objInventoryMovemnt.CalculateAvgCost = False
                objInventoryMovemnt.MRP = objTr.MRP

                objInventoryMovemnt.Add_Cost = objTr.Total_Tax_Amt
                objInventoryMovemnt.Net_Cost = objTr.Total_Tax_Amt
                'If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "RM"
                'ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "OT"
                'ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "FT"
                'End If
                objInventoryMovemnt.ItemType = strItemTypeToSave
                ArrInventoryMovement.Add(objInventoryMovemnt)
            End If
        Next
        isSaved = isSaved AndAlso clsItemLocationDetails.SaveData(clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy"), ArrLocationDetails, trans)
        isSaved = isSaved AndAlso clsInventoryMovement.SaveData(Trans_type_Str, obj.Document_Code, obj.Document_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy"), ArrInventoryMovement, trans)

        ' done by priti on BHA/09/05/18-000022
        Dim AllowCrateCanPhysicalStock As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, trans))
        Dim dblCrateQty As Double = obj.CrateQty
        Dim dblCanQty As Double = obj.ShippedCAN
        If AllowCrateCanPhysicalStock = 1 Then
            If dblCrateQty > 0 Then
                Dim strCrateItem = clsDBFuncationality.getSingleValue("select top 1 Item_Code from TSPL_ITEM_MASTER where isnull(CRATE,0)=1", trans)
                If clsCommon.myLen(strCrateItem) > 0 Then
                    Dim dblRate As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ItemCrateRate, clsFixedParameterCode.ItemCrateRate, trans))
                    Dim strCrateUOM = clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Item_Code='" & strCrateItem & "' and Default_UOM=1", trans)
                    'Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strCrateItem, obj.Bill_To_Location, obj.Document_Code, obj.Document_Date, trans, strCrateUOM, 0)
                    Dim strItemType As String = clsItemMaster.GetItemType(strCrateItem, trans)
                    Dim strItemTypeToSave As String = ""
                    If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                        strItemTypeToSave = "RM"
                    ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                        strItemTypeToSave = "OT"
                    ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                        strItemTypeToSave = "FT"
                    Else
                        strItemTypeToSave = strItemType
                    End If
                    Dim objInventoryMovemnt As New clsInventoryMovement()
                    Dim ArrInventoryMovementCrate As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
                    objInventoryMovemnt.InOut = "I"
                    objInventoryMovemnt.Location_Code = IIf(clsCommon.myLen(clsCommon.myCstr(obj.Sub_Location_code)) > 0, obj.Sub_Location_code, obj.Bill_To_Location)

                    objInventoryMovemnt.Cust_Code = obj.Customer_Code
                    objInventoryMovemnt.Cust_Name = obj.Customer_Name

                    objInventoryMovemnt.Item_Code = strCrateItem
                    objInventoryMovemnt.Item_Desc = clsItemMaster.GetItemName(strCrateItem, trans)
                    objInventoryMovemnt.Qty = obj.CrateQty
                    objInventoryMovemnt.UOM = strCrateUOM
                    objInventoryMovemnt.Basic_Cost = dblRate
                    objInventoryMovemnt.MRP = 0
                    objInventoryMovemnt.Add_Cost = 0
                    objInventoryMovemnt.Net_Cost = 0
                    objInventoryMovemnt.ItemType = strItemTypeToSave
                    ArrInventoryMovementCrate.Add(objInventoryMovemnt)
                    clsInventoryMovement.SaveData(Trans_type_Str, obj.Document_Code, obj.Document_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy"), ArrInventoryMovementCrate, trans)

                    qry = "Update TSPL_SD_SALE_RETURN_HEAD set Crate_Item='" & strCrateItem & "',Crate_ItemUnit='" & strCrateUOM & "',Crate_ItemRate='" & dblRate & "' where Document_Code='" & obj.Document_Code & "' "
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                Else
                    Throw New Exception("Please Create item as Crate type Item.")
                    isSaved = False
                End If
            End If
            ' For Can

            If dblCanQty > 0 Then
                Dim strCanItem = clsDBFuncationality.getSingleValue("select top 1 Item_Code from TSPL_ITEM_MASTER where isnull(Can,0)=1", trans)
                If clsCommon.myLen(strCanItem) > 0 Then
                    Dim dblRate As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ItemCanRate, clsFixedParameterCode.ItemCanRate, trans))
                    Dim strCanUOM = clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Item_Code='" & strCanItem & "' and Default_UOM=1", trans)

                    Dim strItemType As String = clsItemMaster.GetItemType(strCanItem, trans)
                    Dim strItemTypeToSave As String = ""
                    If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                        strItemTypeToSave = "RM"
                    ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                        strItemTypeToSave = "OT"
                    ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                        strItemTypeToSave = "FT"
                    Else
                        strItemTypeToSave = strItemType
                    End If
                    Dim objInventoryMovemnt As New clsInventoryMovement()
                    Dim ArrInventoryMovementCan As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
                    objInventoryMovemnt.InOut = "I"
                    objInventoryMovemnt.Location_Code = IIf(clsCommon.myLen(clsCommon.myCstr(obj.Sub_Location_code)) > 0, obj.Sub_Location_code, obj.Bill_To_Location)

                    objInventoryMovemnt.Cust_Code = obj.Customer_Code
                    objInventoryMovemnt.Cust_Name = obj.Customer_Name

                    objInventoryMovemnt.Item_Code = strCanItem
                    objInventoryMovemnt.Item_Desc = clsItemMaster.GetItemName(strCanItem, trans)
                    objInventoryMovemnt.Qty = obj.ShippedCAN
                    objInventoryMovemnt.UOM = strCanUOM
                    objInventoryMovemnt.Basic_Cost = dblRate
                    objInventoryMovemnt.MRP = 0
                    objInventoryMovemnt.Add_Cost = 0
                    objInventoryMovemnt.Net_Cost = 0
                    objInventoryMovemnt.ItemType = strItemTypeToSave
                    ArrInventoryMovementCan.Add(objInventoryMovemnt)
                    clsInventoryMovement.SaveData(Trans_type_Str, obj.Document_Code, obj.Document_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy"), ArrInventoryMovementCan, trans)

                    qry = "Update TSPL_SD_SALE_RETURN_HEAD set Can_Item='" & strCanItem & "',Can_ItemUnit='" & strCanUOM & "',Can_ItemRate='" & dblRate & "' where Document_Code='" & obj.Document_Code & "' "
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                Else
                    Throw New Exception("Please Create item as Can type Item.")
                    isSaved = False
                End If

            End If
        End If
        Return True
    End Function
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            PostData(FormId, strDocNo, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal trans As SqlTransaction, Optional ByVal strARNoForRecreate As String = Nothing, Optional ByVal strVoucherForRecreate As String = Nothing, Optional ByVal IsSkipEInvoicePosting As Boolean = False) As Boolean
        Try
            Dim isSaved As Boolean = True
            Dim Trans_type_Str As String = ""
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If
            Dim obj As clsDSSalesReturnHead = clsDSSalesReturnHead.GetData(strDocNo, NavigatorType.Current, trans)
            Trans_type_Str = clsCommon.myCstr(obj.Trans_type) + "-SR"
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            '' Anubhooti 06-Sep-2014 BM00000003735 (Locked Transaction)
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleSaleDairy, clsUserMgtCode.frmSaleReturndairy, obj.Bill_To_Location, obj.Document_Date, trans)
            ''
            If (obj.Status = 1) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            If (obj.On_Hold) Then
                Throw New Exception("Transaction " + obj.Document_Code + " Is On Hold.Can't Post it")
            End If

            Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", obj.Document_Code, trans)
            If isResult = False Then
                trans.Commit()
                Return False
            End If
            Dim qry As String = ""
            If obj.isMultipleReturn = 0 Then
                UpdateInventoryMovement(obj, trans, False)
            End If

            Dim dblcount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count (*) as rowno from tspl_sd_sale_return_detail where document_Code ='" & obj.Document_Code & "' and scheme_item='N'", trans))
            If dblcount > 0 Then
                createARInvoice(obj, trans, strARNoForRecreate, strVoucherForRecreate)
            Else
                CreateJournalEntry(obj.Document_Code, trans, "")
            End If



            Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", obj.Against_Invoice_No, trans)

            ''richa agarwal 25 Dec,2020 check eInvoice Implementation
            If IsSkipEInvoicePosting = False Then
                If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Is_Taxable), "1") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True Then
                    If clsDSSalesReturnHead.EInvoice_Implementation(obj.Document_Code, obj.Bill_To_Location, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If

            qry = "Update TSPL_SD_SALE_RETURN_HEAD set Status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "',EInvoice_Type='" + ECustomerType + "',EInvoice_Posting_Date='" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") + "'"
            qry += " where Document_Code='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)

        Catch ex As Exception

            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Sub CreateJournalEntry(ByVal strCode As String, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing)
        Dim RecoControlACC As String = ""
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 0 Then
            RecoControlACC = "I"
        End If
        Dim isSkipCogsGL As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, trans)) = 0, False, True)
        If Not isSkipCogsGL Then
            Dim obj As New clsDSSalesReturnHead
            obj = clsDSSalesReturnHead.GetData(strCode, NavigatorType.Current, trans)
            Dim ArryLstGLAC As ArrayList = New ArrayList()
            Dim strInventoryControlAc As String = ""
            Dim strShipmentClearingAC As String = ""
            Dim dblTotalCost As Double = 0
            Dim strSql As String = String.Empty
            Dim strCogsAcct As String
            Dim ArryLst As ArrayList = New ArrayList()
            Dim SentschemecogsinAnotherAccount As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SentschemecogsinAnotherAccount, clsFixedParameterCode.SentschemecogsinAnotherAccount, trans)) = "1", True, False))

            strSql = "select inv.Is_Scheme_Item,inv.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from " &
                            " ( select isnull(TSPL_INVENTORY_MOVEMENT.Is_Scheme_Item,'N') as Is_Scheme_Item,Item_Code,Source_Doc_No,Avg_Cost,FIFO_Cost,LIFO_Cost from TSPL_INVENTORY_MOVEMENT where TSPL_INVENTORY_MOVEMENT.Source_Doc_No='" & obj.Document_Code & "' " &
                            " union all " &
                            " select isnull(TSPL_INVENTORY_MOVEMENT_NEW.Is_Scheme_Item,'N') as Is_Scheme_Item,Item_Code,Source_Doc_No,Avg_Cost,FIFO_Cost,LIFO_Cost from TSPL_INVENTORY_MOVEMENT_NEW where TSPL_INVENTORY_MOVEMENT_NEW.Source_Doc_No='" & obj.Document_Code & "' " &
                            " ) inv " &
                            " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=inv.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  where isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0 and Source_Doc_No='" & obj.Document_Code & "'"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                For Each dr As DataRow In dt.Rows
                    strInventoryControlAc = clsDBFuncationality.getSingleValue("SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                            " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                            " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                    strInventoryControlAc = clsERPFuncationality.ChangeGLAccountLocationSegment(strInventoryControlAc, obj.Bill_To_Location, trans)

                    If clsCommon.myLen(strInventoryControlAc) = 0 Then
                        Throw New Exception("Please set Inventory Control Account for first item")
                    End If
                    If SentschemecogsinAnotherAccount = True Then
                        If clsCommon.CompairString(dr("Is_Scheme_Item"), "Y") = CompairStringResult.Equal Then
                            strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Scheme from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                            If clsCommon.myLen(strCogsAcct) = 0 Then
                                Throw New Exception("Please set Cost Of Goods Scheme Account for first item")
                            End If
                        Else
                            strCogsAcct = clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                            If clsCommon.myLen(strCogsAcct) = 0 Then
                                Throw New Exception("Please set Cost of Goods Sold for first item")
                            End If
                        End If
                    Else
                        strCogsAcct = clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                        If clsCommon.myLen(strCogsAcct) = 0 Then
                            Throw New Exception("Please set Cost of Goods Sold for first item")
                        End If
                    End If

                    strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.Bill_To_Location, trans)

                    Dim Acc() As String = {strInventoryControlAc, clsCommon.myCdbl(dr("Cost")), "", "", "", "", "", "", RecoControlACC}
                    ArryLst.Add(Acc)

                    Dim Acc1() As String = {strCogsAcct, -1 * clsCommon.myCdbl(dr("Cost"))}
                    ArryLst.Add(Acc1)
                Next
            End If

            obj.Remarks = "Journal entry against Sale Return " & obj.Document_Code & " created only when scheme item return"
            If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then
                clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, obj.Remarks, "DS-SR", "Sale Return", obj.Document_Code, "", "C", obj.Customer_Code, obj.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, obj.Description, obj.Remarks)
            Else
                clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, trans, obj.Document_Date, obj.Remarks, "DS-SR", "Sale Return", obj.Document_Code, "", "C", obj.Customer_Code, obj.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, obj.Description, obj.Remarks)
            End If
        End If
    End Sub

    Public Shared Function CancelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal NavType As NavigatorType, ByVal isCancelByAdmin As Boolean) As Boolean
        '' created by Sanjay
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            Dim obj As clsDSSalesReturnHead = clsDSSalesReturnHead.GetData(Doc_No, NavType, trans)

            If obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If

            ''richa agarwal 06 Jan,2021
            Dim dtirn As DataTable = clsDBFuncationality.GetDataTable("select Einvoice_type,IRN_No from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & Doc_No & "'", trans)
            If dtirn IsNot Nothing AndAlso dtirn.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtirn.Rows(0)("Einvoice_type")), "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Is_Taxable), "1") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True AndAlso Not isCancelByAdmin Then
                    If ClsEInvoiceOFAPIs.EInvoice_Cancellation(Doc_No, clsCommon.myCstr(dtirn.Rows(0)("IRN_No")), obj.Bill_To_Location, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If


            clsItemLocationDetails.CheckCancelInventoryBalance(Form_Id, Doc_No, trans)

            '' transfer data into cancel table
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)

            '' cancel customer invoice data
            qry = "select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & Doc_No & "'"
            Dim Document_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Document_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Document_No, "TSPL_Customer_Invoice_Head", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
            End If

            '' cancel journal master data invoice
            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & Doc_No & "')"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            '' cancel custom field data
            clsCommonFunctionality.SaveCancelDataMultKey(objCommonVar.CurrentUserCode, Doc_No, "TSPL_CUSTOM_FIELD_VALUES", "Transaction_Code", "Program_Code", Form_Id, trans)

            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & Doc_No & "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Detail where Document_No in (Select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Head where Against_Sale_Return_No='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_BATCH_ITEM where  Document_Code='" & Doc_No & "' and Document_Type='FS-SR'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "' and Trans_Type='FS-SR'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_CUSTOM_FIELD_VALUES where Transaction_Code ='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_SD_SALE_RETURN_DETAIL where Document_Code='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

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

    Public Shared Function EInvoice_Implementation(ByVal strDocNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If

            Dim strtoken As String = ClsEInvoiceOFAPIs.IsGenerateAuthTokenNo_Required(objCommonVar.CurrentCompanyCode, strLocation, trans)
            If clsCommon.myLen(strtoken) > 0 Then
                '                Dim strQry As String = "select TSPL_Customer_master .Cust_Code ,TSPL_SD_SALE_RETURN_HEAD.Document_Code as DocNo,convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) as DocDate,case when TSPL_Customer_Invoice_Head.Document_Type='I' then 'INV' when TSPL_Customer_Invoice_Head.Document_Type ='D' then 'DBN' else 'CRN' end as DocType ,'B2B' as SupTyp, 'N'  as IgstOnIntra,Bill_To_Location.GSTNO as SellerGSTINNo ,Bill_To_Location.location_desc as SellerLglNm,TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm,Bill_To_Location.Add1 as SellerAdd1,Bill_To_Location.Add2 as SellerAdd2 ,Bill_To_Location.City_Code  as SellerLoc,Bill_To_Location.Pin_Code  as SellerPincode,BillToLocation_State_Master.GST_STATE_Code as SellerStcd,Bill_To_Location.Phone1 as SellerPhone,Bill_To_Location.Email as SellerEmail,TSPL_Customer_master.GSTNo as BuyerGSTINNo ,TSPL_Customer_master.Customer_Name as BuyerLglNm,TSPL_Customer_master.Alies_name as BuyerTrdNm,case when isnull(TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location,'')='' then Customer_State_Master.GST_STATE_Code else Ship_To_Location_State_Master.GST_STATE_Code end as BuyerPOS,TSPL_Customer_master.Add1 as BuyerAdd1,TSPL_Customer_master.Add2 as BuyerAdd2 ,tspl_city_master.City_Name as BuyerLoc,cast(TSPL_Customer_master.PIN_NO as int) as BuyerPincode,Customer_State_Master.GST_STATE_Code as BuyerStcd,TSPL_Customer_master.Phone1 as BuyerPhone,TSPL_Customer_master.Email as BuyerEmail,TSPL_SD_SALE_RETURN_DETAIL.Line_No as ItemSlNo,case when isnull(TSPL_SD_SALE_RETURN_DETAIL.Row_Type,'') ='Item' then 'N' else 'Y' end as ItemIsServc,COAlESCE(TSPL_ITEM_MASTER.Item_Desc,TSPL_Additional_Charges.DESCRIPTION ) AS ItemPrdDesc,COAlESCE(TSPL_ITEM_MASTER.HSN_Code,TSPL_Additional_Charges.SAC_CODE) AS ItemHsnCd,case when TSPL_SD_SALE_RETURN_DETAIL.Actualqty>0 then TSPL_SD_SALE_RETURN_DETAIL.Actualqty else TSPL_SD_SALE_RETURN_DETAIL.qty end  as ItemQty, case when isnull(TSPL_SD_SALE_RETURN_DETAIL.ActualUom,'')<>'' then TSPL_SD_SALE_RETURN_DETAIL.ActualUom else TSPL_SD_SALE_RETURN_DETAIL.Unit_code end as ItemUnit,TSPL_SD_SALE_RETURN_DETAIL.Item_Cost as ItemUnitPrice,TSPL_SD_SALE_RETURN_DETAIL.Amount as ItemTotAmt,TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt as ItemDiscount,TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount as ItemAssAmt,case when ISNULL(TSPL_SD_SALE_RETURN_DETAIL .tax1,'') ='IGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate when ISNULL(TSPL_SD_SALE_RETURN_DETAIL .tax1,'') ='CGST' AND   ISNULL(TSPL_SD_SALE_RETURN_DETAIL .tax2,'') ='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate+TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate  ELSE 0 end as ItemGstRt, case when TSPL_SD_SALE_RETURN_DETAIL .TAX1 ='SGST' AND TSPL_SD_SALE_RETURN_DETAIL .TAX2  ='CGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt when TSPL_SD_SALE_RETURN_DETAIL .TAX1 ='CGST' AND TSPL_SD_SALE_RETURN_DETAIL .TAX2  ='SGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt else 0 end ItemSgstAmt,case when TSPL_SD_SALE_RETURN_DETAIL .TAX1 ='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt else 0 end ItemIgstAmt,
                '        case when TSPL_SD_SALE_RETURN_DETAIL .TAX1 ='CGST' AND TSPL_SD_SALE_RETURN_DETAIL .TAX2  ='SGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt when TSPL_SD_SALE_RETURN_DETAIL .TAX1 ='SGST' AND TSPL_SD_SALE_RETURN_DETAIL .TAX2  ='CGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt else 0 end ItemCgstAmt,0 as ItemOthChrg,TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt-case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_SD_SALE_RETURN_DETAIL.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y' THEN  TSPL_SD_SALE_RETURN_DETAIL.TAX3_AMT ELSE 0 END as ItemTotItemVal,TSPL_SD_SALE_RETURN_HEAD .Amount_Less_Discount as ValDtlsAssVal,case when TSPL_SD_SALE_RETURN_HEAD .TAX1 ='CGST' AND TSPL_SD_SALE_RETURN_HEAD .TAX2  ='SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt when TSPL_SD_SALE_RETURN_HEAD .TAX1 ='SGST' AND TSPL_SD_SALE_RETURN_HEAD .TAX2  ='CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt else 0 end ValDtlsCgstVal, case when TSPL_SD_SALE_RETURN_HEAD .TAX1 ='SGST' AND TSPL_SD_SALE_RETURN_HEAD .TAX2  ='CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt when TSPL_SD_SALE_RETURN_HEAD .TAX1 ='CGST' AND TSPL_SD_SALE_RETURN_HEAD .TAX2  ='SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt else 0 end ValDtlsSgstVal,case when TSPL_SD_SALE_RETURN_HEAD .TAX1 ='IGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt else 0 end ValDtlsIgstVal,TSPL_SD_SALE_RETURN_HEAD.Discount_Amt as ValDtlsDiscount,case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_SD_SALE_RETURN_HEAD.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y' THEN  TSPL_SD_SALE_RETURN_HEAD.TAX3_AMT ELSE 0 END as ValDtlsOthChrg,TSPL_SD_SALE_RETURN_HEAD.Total_Amt as ValDtlsTotInvVal,TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount  as ValDtlsRndOffAmt
                '        from TSPL_SD_SALE_RETURN_HEAD
                '        Left Outer Join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_Return_No =TSPL_SD_SALE_RETURN_HEAD.Document_Code
                '        Left Outer Join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code  ='" & objCommonVar.CurrentCompanyCode & "'
                '        Left Outer Join TSPL_Customer_master on TSPL_Customer_master.Cust_Code  =TSPL_SD_SALE_RETURN_HEAD.Customer_Code
                '        left Outer Join TSPL_LOCATION_MASTER as Bill_To_Location on Bill_To_Location.Location_Code =TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location 
                '        left Outer Join TSPL_SHIP_TO_LOCATION as Ship_To_Location on Ship_To_Location.Ship_To_Code  =TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location  
                '        left outer join TSPL_SD_SALE_RETURN_DETAIL on TSPL_SD_SALE_RETURN_DETAIL.document_code=TSPL_SD_SALE_RETURN_HEAD.document_code
                '        left outer join tspl_item_master on tspl_item_master.Item_code=TSPL_SD_SALE_RETURN_DETAIL.Item_code
                '        left outer join TSPL_ADDITIONAL_CHARGES on TSPL_ADDITIONAL_CHARGES.CODE=TSPL_SD_SALE_RETURN_DETAIL.Item_code
                '        left outer join TSPL_STATE_MASTER as BillToLocation_State_Master on BillToLocation_State_Master.STATE_CODE  =Bill_To_Location.State
                '        left outer join TSPL_STATE_MASTER as Ship_To_Location_State_Master on Ship_To_Location_State_Master.STATE_CODE  =Ship_To_Location.State
                '        left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE  =TSPL_Customer_master.State 
                '        left outer join tspl_city_master on tspl_city_master.city_code=TSPL_Customer_master.City_Code
                '		left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code =TSPL_SD_SALE_RETURN_HEAD.Tax2
                'left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code =TSPL_SD_SALE_RETURN_HEAD.Tax3
                '        where TSPL_SD_SALE_RETURN_HEAD.Document_Code ='" & strDocNo & "'"
                Dim strQry As String = "select TSPL_Customer_master.Cust_Code, TSPL_SD_SALE_RETURN_HEAD.Document_Code as DocNo, convert( date, TSPL_SD_SALE_RETURN_HEAD.Document_Date, 103 ) as DocDate, case when TSPL_Customer_Invoice_Head.Document_Type = 'I' then 'INV' when TSPL_Customer_Invoice_Head.Document_Type = 'D' then 'DBN' else 'CRN' end as DocType, 'B2B' as SupTyp, 'N' as IgstOnIntra, Bill_To_Location.GSTNO as SellerGSTINNo, Bill_To_Location.location_desc as SellerLglNm, TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm, Bill_To_Location.Add1 as SellerAdd1, Bill_To_Location.Add2 as SellerAdd2, Bill_To_Location.City_Code as SellerLoc, Bill_To_Location.Pin_Code as SellerPincode, BillToLocation_State_Master.GST_STATE_Code as SellerStcd, Bill_To_Location.Phone1 as SellerPhone, Bill_To_Location.Email as SellerEmail, TSPL_Customer_master.GSTNo as BuyerGSTINNo, TSPL_Customer_master.Customer_Name as BuyerLglNm, TSPL_Customer_master.Alies_name as BuyerTrdNm, case when isnull( TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location, '' )= '' then Customer_State_Master.GST_STATE_Code else Ship_To_Location_State_Master.GST_STATE_Code end as BuyerPOS, TSPL_Customer_master.Add1 as BuyerAdd1, TSPL_Customer_master.Add2 as BuyerAdd2, tspl_city_master.City_Name as BuyerLoc, cast( TSPL_Customer_master.PIN_NO as int ) as BuyerPincode, Customer_State_Master.GST_STATE_Code as BuyerStcd, TSPL_Customer_master.Phone1 as BuyerPhone, TSPL_Customer_master.Email as BuyerEmail, TSPL_SD_SALE_RETURN_DETAIL.Line_No as ItemSlNo, case when isnull( TSPL_SD_SALE_RETURN_DETAIL.Row_Type, '' ) = 'Item' then 'N' else 'Y' end as ItemIsServc, COAlESCE( TSPL_ITEM_MASTER.Item_Desc, TSPL_Additional_Charges.DESCRIPTION ) AS ItemPrdDesc, COAlESCE( TSPL_ITEM_MASTER.HSN_Code, TSPL_Additional_Charges.SAC_CODE ) AS ItemHsnCd, case when TSPL_SD_SALE_RETURN_DETAIL.Actualqty > 0 or TSPL_SD_SALE_RETURN_DETAIL.DamageQty>0 then TSPL_SD_SALE_RETURN_DETAIL.Actualqty + TSPL_SD_SALE_RETURN_DETAIL.DamageQty else TSPL_SD_SALE_RETURN_DETAIL.qty end as ItemQty, case when isnull( TSPL_SD_SALE_RETURN_DETAIL.ActualUom, '' )<> '' then TSPL_SD_SALE_RETURN_DETAIL.ActualUom else TSPL_SD_SALE_RETURN_DETAIL.Unit_code end as ItemUnit,  Case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then( ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt + TSPL_SD_SALE_RETURN_DETAIL.Disc_Amt ) / TSPL_SD_SALE_RETURN_DETAIL.Qty ) else TSPL_SD_SALE_RETURN_DETAIL.Item_Cost end as ItemUnitPrice, case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt + TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt else TSPL_SD_SALE_RETURN_DETAIL.Amount end as ItemTotAmt, TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt as ItemDiscount, ( case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt else TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount end -( case when isnull( TSPL_SD_SALE_RETURN_DETAIL.FOC_Item, 0 )= 1 then TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt else 0 end ) ) as ItemAssAmt, Case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' ) = 'IGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate + TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' ) = 'CGST' AND ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' ) = 'SGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate + TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate ELSE 0 end as ItemGstRt, Case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_DETAIL.TAX3 = 'SGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX4 = 'CGST' then ( ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt )* ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate ) )/ 100 when TSPL_SD_SALE_RETURN_DETAIL.TAX3 = 'CGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX4 = 'SGST' then ( ( TSPL_SD_SALE_RETURN_DETAIL.TAX4_Base_Amt )* ( TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate ) )/ 100 else 0 end ) else ( case when TSPL_SD_SALE_RETURN_DETAIL.TAX1 = 'SGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX2 = 'CGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt when TSPL_SD_SALE_RETURN_DETAIL.TAX1 = 'CGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX2 = 'SGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt else 0 end ) end ItemSgstAmt, Case when TSPL_SD_SALE_RETURN_DETAIL.TAX1 = 'IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt else ( case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then( case when TSPL_SD_SALE_RETURN_DETAIL.TAX3 = 'IGST' then ( ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt )* ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate ) )/ 100 else 0 end ) else 0 end ) end ItemIgstAmt, case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_DETAIL.TAX3 = 'CGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX4 = 'SGST' then ( ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt )* ( TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate ) )/ 100 when TSPL_SD_SALE_RETURN_DETAIL.TAX3 = 'SGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX4 = 'CGST' then ( ( TSPL_SD_SALE_RETURN_DETAIL.TAX4_Base_Amt )* ( TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate ) )/ 100 else 0 end ) else ( case when TSPL_SD_SALE_RETURN_DETAIL.TAX1 = 'CGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX2 = 'SGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt when TSPL_SD_SALE_RETURN_DETAIL.TAX1 = 'SGST' AND TSPL_SD_SALE_RETURN_DETAIL.TAX2 = 'CGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt else 0 end ) end ItemCgstAmt, Case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_DETAIL.TAX5 = 'TCS' then 0 else 0 end ) else ( case when TSPL_SD_SALE_RETURN_DETAIL.tax3 = 'TCS' then 0 else 0 end ) end as ItemOthChrg, ( case when ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_DETAIL.tax2, '' )= 'KKF' then ( ( TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount -( case when isnull( TSPL_SD_SALE_RETURN_DETAIL.FOC_Item, 0 )= 1 then TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt else 0 end ) )+ TSPL_SD_SALE_RETURN_DETAIL.TAX1_AMT + TSPL_SD_SALE_RETURN_DETAIL.TAX2_AMT + TSPL_SD_SALE_RETURN_DETAIL.TAX3_AMT + TSPL_SD_SALE_RETURN_DETAIL.TAX4_AMT ) else ( ( TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount -( case when isnull( TSPL_SD_SALE_RETURN_DETAIL.FOC_Item, 0 )= 1 then TSPL_SD_SALE_RETURN_DETAIL.total_disc_amt else 0 end ) )+ TSPL_SD_SALE_RETURN_DETAIL.TAX1_AMT + ( case when isnull(TCS1.is_tcs, '')<> 'Y' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_AMT ELSE 0 END ) + ( case when isnull(TCS2.is_tcs, '')<> 'Y' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_AMT ELSE 0 END ) ) end ) as ItemTotItemVal,
Case when ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax2, '' )= 'KKF' then TSPL_SD_SALE_RETURN_HEAD.TAX3_Base_Amt else TSPL_SD_SALE_RETURN_HEAD.TAX1_Base_Amt end as ValDtlsAssVal, Case when ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_HEAD.TAX3 = 'CGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX4 = 'SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX3_Amt when TSPL_SD_SALE_RETURN_HEAD.TAX3 = 'SGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX4 = 'CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX4_Amt else 0 end ) else ( case when TSPL_SD_SALE_RETURN_HEAD.TAX1 = 'CGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX2 = 'SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt when TSPL_SD_SALE_RETURN_HEAD.TAX1 = 'SGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX2 = 'CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt else 0 end ) end ValDtlsCgstVal, Case when ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_HEAD.TAX3 = 'SGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX4 = 'CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX3_Amt when TSPL_SD_SALE_RETURN_HEAD.TAX3 = 'CGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX4 = 'SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX4_Amt else 0 end ) else ( case when TSPL_SD_SALE_RETURN_HEAD.TAX1 = 'SGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX2 = 'CGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt when TSPL_SD_SALE_RETURN_HEAD.TAX1 = 'CGST' AND TSPL_SD_SALE_RETURN_HEAD.TAX2 = 'SGST' then TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt else 0 end ) end ValDtlsSgstVal, Case when TSPL_SD_SALE_RETURN_HEAD.TAX1 = 'IGST' then TSPL_SD_SALE_RETURN_HEAD.TAX1_Amt else ( case when ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_HEAD.TAX3 = 'IGST' then TSPL_SD_SALE_RETURN_HEAD.TAX3_Amt else 0 end ) else 0 end ) end ValDtlsIgstVal, 0 as ValDtlsDiscount,
Case when ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax1, '' )= 'KKF' or ISNULL( TSPL_SD_SALE_RETURN_HEAD.tax2, '' )= 'KKF' then ( case when TSPL_SD_SALE_RETURN_HEAD.TAX5 = 'TCS' then TSPL_SD_SALE_RETURN_HEAD.TAX5_Amt else 0 end ) else ( case when TSPL_SD_SALE_RETURN_HEAD.tax3 = 'TCS' then TSPL_SD_SALE_RETURN_HEAD.TAX3_Amt else 0 end ) end as ValDtlsOthChrg, TSPL_SD_SALE_RETURN_HEAD.Total_Amt as ValDtlsTotInvVal, TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount As ValDtlsRndOffAmt 
from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_Return_No = TSPL_SD_SALE_RETURN_HEAD.Document_Code
Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrentCompanyCode & "' 
Left Outer Join TSPL_Customer_master on TSPL_Customer_master.Cust_Code = TSPL_SD_SALE_RETURN_HEAD.Customer_Code 
left Outer Join TSPL_LOCATION_MASTER as Bill_To_Location on Bill_To_Location.Location_Code = TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location
left Outer Join TSPL_SHIP_TO_LOCATION as Ship_To_Location on Ship_To_Location.Ship_To_Code = TSPL_SD_SALE_RETURN_HEAD.Ship_To_Location
left outer join TSPL_SD_SALE_RETURN_DETAIL on TSPL_SD_SALE_RETURN_DETAIL.document_code = TSPL_SD_SALE_RETURN_HEAD.document_code
left outer join tspl_item_master on tspl_item_master.Item_code = TSPL_SD_SALE_RETURN_DETAIL.Item_code
left outer join TSPL_ADDITIONAL_CHARGES on TSPL_ADDITIONAL_CHARGES.CODE = TSPL_SD_SALE_RETURN_DETAIL.Item_code 
left outer join TSPL_STATE_MASTER as BillToLocation_State_Master on BillToLocation_State_Master.STATE_CODE = Bill_To_Location.State
left outer join TSPL_STATE_MASTER as Ship_To_Location_State_Master on Ship_To_Location_State_Master.STATE_CODE = Ship_To_Location.State 
left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE = TSPL_Customer_master.State 
left outer join tspl_city_master on tspl_city_master.city_code = TSPL_Customer_master.City_Code 
left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code = TSPL_SD_SALE_RETURN_HEAD.Tax2 left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code = TSPL_SD_SALE_RETURN_HEAD.Tax3
where TSPL_SD_SALE_RETURN_HEAD.Document_Code = '" & strDocNo & "'"


                Dim objResult As Object = ClsEInvoiceOFAPIs.PostAuthTokenNo_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strQry, strLocation, trans, True)
                If objResult IsNot Nothing Then
                    'assign to variable
                    Dim AckNo As String = objResult.SelectToken("AckNo").ToString
                    Dim AckDt As String = objResult.SelectToken("AckDt").ToString
                    Dim Irn As String = objResult.SelectToken("Irn").ToString
                    Dim SignedQRCode As String = objResult.SelectToken("SignedQRCode").ToString
                    clsDBFuncationality.ExecuteNonQuery("update TSPL_SD_SALE_RETURN_HEAD set  IRN_No ='" & Irn & "',qr_code='" & SignedQRCode & "',ack_no='" & AckNo & "',ack_date='" & clsCommon.GetPrintDate(AckDt, "dd/MMM/yyyy hh:mm tt") & "' where Document_Code ='" & strDocNo & "'", trans)

                    Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(SignedQRCode)
                    clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_SD_SALE_RETURN_HEAD", "TSPL_SD_SALE_RETURN_HEAD.Document_Code='" & strDocNo & "'", trans)
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
    Public Shared Function ReverseAndUnpost(ByVal strCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ReverseAndUnpost(strCode, trans, False)
            trans.Commit()
        Catch ex As Exception
            Throw New Exception(ex.Message)
            trans.Rollback()
        End Try
        Return True
    End Function
    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal trans As SqlTransaction, ByVal isSkipBalanceCheck As Boolean) As Boolean

        Try
            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select Status from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" + strCode + "'"
            Dim Trans_type_Str As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Trans_Type from TSPL_SD_SALE_RETURN_HEAD where document_code='" + strCode + "'", trans))
            Trans_type_Str = Trans_type_Str + "-SR"
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If
            If isSkipBalanceCheck = False Then
                Dim dblBalance As Decimal = 0
                Dim obj As clsDSSalesReturnHead = clsDSSalesReturnHead.GetData(strCode, NavigatorType.Current, trans)
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleSaleDairy, clsUserMgtCode.frmSaleReturndairy, obj.Bill_To_Location, obj.Document_Date, trans)

                For Each objtr As clsDSSalesReturnDetail In obj.Arr
                    dblBalance = clsItemLocationDetails.getBalanceWithUnapproveForRMOther(objtr.Item_Code, IIf(clsCommon.myLen(clsCommon.myCstr(obj.Sub_Location_code)) > 0, obj.Sub_Location_code, objtr.Location), obj.Document_Code, obj.Document_Date, trans, IIf(clsCommon.myLen(clsCommon.myCstr(objtr.ActualUOM)) > 0, objtr.ActualUOM, objtr.Unit_code))
                    If dblBalance < objtr.ActuaQty Then
                        If Math.Abs(dblBalance - objtr.ActuaQty) > 0.01 Then
                            Throw New Exception("Balance will be going to -ve.Balance Qty : " + clsCommon.myCstr(dblBalance) + " and Entered Qty : " + clsCommon.myCstr(objtr.ActuaQty) + Environment.NewLine + "Item : " + objtr.Item_Code + Environment.NewLine + "Unit : " + IIf(clsCommon.myLen(clsCommon.myCstr(objtr.ActualUOM)) > 0, objtr.ActualUOM, objtr.Unit_code))
                        End If
                    End If

                    ''End of Main Item

                    ''Batch Item
                    If objtr.arrBatchItem IsNot Nothing AndAlso objtr.arrBatchItem.Count > 0 Then
                        For Each objBatch As clsBatchInventory In objtr.arrBatchItem
                            dblBalance = clsBatchInventory.GetBatchBalance(objBatch.Item_Code, objBatch.Location_Code, objBatch.Batch_No, objBatch.MRP, objBatch.UOM, objBatch.Document_Code, objBatch.Document_Type, trans)
                            If dblBalance < objBatch.Qty Then
                                Throw New Exception("Balance will be going to -ve.Balance Qty : " + clsCommon.myCstr(dblBalance) + " and Entered Qty : " + clsCommon.myCstr(objBatch.Qty) + Environment.NewLine + "Item : " + objBatch.Item_Code + Environment.NewLine + "Batch : " + objBatch.Batch_No + Environment.NewLine + "MRP : " + clsCommon.myCstr(objBatch.MRP) + Environment.NewLine + "Unit : " + objBatch.UOM)
                            End If
                        Next
                    End If

                    ''End of Batch Item
                Next
            End If




            Dim strARInvoiceNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No  from TSPL_Customer_Invoice_Head where Against_Sale_Return_No='" + strCode + "'", trans))

            If clsCommon.myLen(strARInvoiceNo) > 0 Then
                Qry = "select distinct Receipt_No from TSPL_RECEIPT_DETAIL where Document_No='" + strARInvoiceNo + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Qry = "Current Sale Return - " + strCode + ".AR No - " + strARInvoiceNo + " is used in following Receipt Entry -"
                    For Each dr As DataRow In dt.Rows
                        Qry += Environment.NewLine + clsCommon.myCstr(dr("Receipt_No"))
                    Next
                    Throw New Exception(Qry)
                End If


                Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='AR-CR' and Source_Doc_No='" + strARInvoiceNo + "'", trans)
                If clsCommon.myLen(VoucherNo) > 0 Then
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                    Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                    Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                End If
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strARInvoiceNo, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", "TSPL_REMITTANCE", "Document_No", trans)
                Qry = "delete from TSPL_Customer_Invoice_Detail where Document_No ='" + strARInvoiceNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_Customer_Invoice_Head where Document_No ='" + strARInvoiceNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT_New", "Source_Doc_No", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "tspl_batch_item", "document_code", trans)
            '' delete from inventory movement
            Qry = "update tspl_batch_item set against_inv_movement_trans_id=NULL where document_type='" + Trans_type_Str + "' and document_code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            Qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & strCode & "' and Trans_Type= '" + Trans_type_Str + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            Qry = "delete from TSPL_INVENTORY_MOVEMENT_New where Source_Doc_No='" & strCode & "' and Trans_Type= '" + Trans_type_Str + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            ''delete journal entry which has created against when only scheme item has returned
            Dim VoucherNo1 As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='DS-SR' and Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo1) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo1, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo1 + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo1 + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If


            Qry = "Update TSPL_SD_SALE_RETURN_HEAD set Status = 0 where Document_Code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            'Qry = "update TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL set is_reverse=1 where document_code='" + strCode + "' and trans_code='" + clsCommon.myCstr(clsUserMgtCode.frmSaleReturnProductSale) + "' and is_reverse=0"
            Qry = "update TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL set is_reverse=1 where document_code='" + strCode + "' and trans_code='" + clsCommon.myCstr(clsUserMgtCode.frmSaleReturndairy) + "' and is_reverse=0"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)
        Catch ex As Exception

            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function createARInvoice(ByVal obj As clsDSSalesReturnHead, ByVal trans As SqlTransaction, Optional ByVal strARNoForRecreate As String = Nothing, Optional ByVal strVoucherForRecreate As String = Nothing) As Boolean
        ''''''''''''''''''''''''''''''''''For Making AR Invoice
        Dim objCustInv As New clsCustomerInvoiceHead()
        ''objCustInv.Document_No ''Will be Generateed
        objCustInv.Document_Date = obj.Document_Date
        objCustInv.Document_Type = "C"
        objCustInv.Document_Total = obj.Total_Amt
        objCustInv.RoundOffAmount = obj.RoundOffAmount
        objCustInv.Customer_Code = obj.Customer_Code
        objCustInv.Customer_Name = obj.Customer_Name
        objCustInv.Posting_Date = obj.Document_Date
        objCustInv.Trans_Type = obj.Trans_type
        Dim qry As String = " select Cust_Account from TSPL_CUSTOMER_MASTER where Cust_Code='" + obj.Customer_Code + "'"
        objCustInv.Account_Set = clsDBFuncationality.getSingleValue(qry, trans)
        ''objCustInv.Order_No
        objCustInv.loc_code = clsLocation.GetSegmentCode(obj.Bill_To_Location, trans)
        objCustInv.Location_Code_Prefix = obj.Bill_To_Location
        objCustInv.On_Hold = 0
        objCustInv.Remarks = obj.Remarks
        objCustInv.Description = obj.Description
        objCustInv.Tax_Group = obj.Tax_Group
        objCustInv.TAX1 = obj.TAX1
        objCustInv.TAX1_Rate = obj.TAX1_Rate
        objCustInv.TAX1_Amt = obj.TAX1_Amt
        objCustInv.TAX2 = obj.TAX2
        objCustInv.TAX2_Rate = obj.TAX2_Rate
        objCustInv.TAX2_Amt = obj.TAX2_Amt
        objCustInv.TAX3 = obj.TAX3
        objCustInv.TAX3_Rate = obj.TAX3_Rate
        objCustInv.TAX3_Amt = obj.TAX3_Amt
        objCustInv.TAX4 = obj.TAX4
        objCustInv.TAX4_Rate = obj.TAX4_Rate
        objCustInv.TAX4_Amt = obj.TAX4_Amt
        objCustInv.TAX5 = obj.TAX5
        objCustInv.TAX5_Rate = obj.TAX5_Rate
        objCustInv.TAX5_Amt = obj.TAX5_Amt
        objCustInv.TAX6 = obj.TAX6
        objCustInv.TAX6_Rate = obj.TAX6_Rate
        objCustInv.TAX6_Amt = obj.TAX6_Amt
        objCustInv.TAX7 = obj.TAX7
        objCustInv.TAX7_Rate = obj.TAX7_Rate
        objCustInv.TAX7_Amt = obj.TAX7_Amt
        objCustInv.TAX8 = obj.TAX8
        objCustInv.TAX8_Rate = obj.TAX8_Rate
        objCustInv.TAX8_Amt = obj.TAX8_Amt
        objCustInv.TAX9 = obj.TAX9
        objCustInv.TAX9_Rate = obj.TAX9_Rate
        objCustInv.TAX9_Amt = obj.TAX9_Amt
        objCustInv.TAX10 = obj.TAX10
        objCustInv.TAX10_Rate = obj.TAX10_Rate
        objCustInv.TAX10_Amt = obj.TAX10_Amt
        objCustInv.Total_Tax = obj.Total_Tax_Amt
        objCustInv.Tax1_BAmount = obj.TAX1_Base_Amt
        objCustInv.Tax2_BAmount = obj.TAX2_Base_Amt
        objCustInv.Tax3_BAmount = obj.TAX3_Base_Amt
        objCustInv.Tax4_BAmount = obj.TAX4_Base_Amt
        objCustInv.Tax5_BAmount = obj.TAX5_Base_Amt
        objCustInv.Tax6_BAmount = obj.TAX6_Base_Amt
        objCustInv.Tax7_BAmount = obj.TAX7_Base_Amt
        objCustInv.Tax8_BAmount = obj.TAX8_Base_Amt
        objCustInv.Tax9_BAmount = obj.TAX9_Base_Amt
        objCustInv.Tax10_BAmount = obj.TAX10_Base_Amt
        objCustInv.Balance_Amt = obj.Total_Amt
        objCustInv.Terms_Code = obj.Terms_Code
        objCustInv.Return_Type = obj.Return_Type
        objCustInv.Route_No = obj.Route_No
        qry = "select Terms_Code,Terms_Desc,No_Days from TSPL_TERMS_MASTER where Terms_Code='" + obj.Terms_Code + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            objCustInv.Terms_Description = clsCommon.myCstr(dt.Rows(0)("Terms_Desc"))
            objCustInv.Due_Date = obj.Document_Date.AddDays(clsCommon.myCdbl(dt.Rows(0)("No_Days")))
        End If
        objCustInv.Discount_Percentage = IIf(obj.Discount_Base = 0, 0, obj.Discount_Amt * 100 / obj.Discount_Base)
        objCustInv.Discount_Base = obj.Discount_Base
        objCustInv.Discount_Amount = obj.Discount_Amt + obj.HeadDisc_Amt + obj.HeadDisc_PerAmt
        ''objCustInv.Amount_Less_Discount = 
        dt = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + objCustInv.Account_Set + "'", trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            objCustInv.Customer_Control_AC = clsCommon.myCstr(dt.Rows(0)("Receivable_Control_acct"))
            If clsCommon.myCdbl(obj.Discount_Amt) > 0 Then
                objCustInv.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Receipts_Discount_acct"))
            End If
        End If

        If obj.TAX1_Amt > 0 AndAlso clsCommon.myLen(obj.TAX1) > 0 Then
            objCustInv.TAX1_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
            objCustInv.TAX1_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX1_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX2_Amt > 0 AndAlso clsCommon.myLen(obj.TAX2) > 0 Then
            objCustInv.TAX2_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
            objCustInv.TAX2_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX2_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX3_Amt > 0 AndAlso clsCommon.myLen(obj.TAX3) > 0 Then
            objCustInv.TAX3_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
            objCustInv.TAX3_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX3_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX4_Amt > 0 AndAlso clsCommon.myLen(obj.TAX4) > 0 Then
            objCustInv.TAX4_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
            objCustInv.TAX4_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX4_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX5_Amt > 0 AndAlso clsCommon.myLen(obj.TAX5) > 0 Then
            objCustInv.TAX5_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
            objCustInv.TAX5_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX5_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX6_Amt > 0 AndAlso clsCommon.myLen(obj.TAX6) > 0 Then
            objCustInv.TAX6_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
            objCustInv.TAX6_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX6_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX7_Amt > 0 AndAlso clsCommon.myLen(obj.TAX7) > 0 Then
            objCustInv.TAX7_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
            objCustInv.TAX7_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX7_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX8_Amt > 0 AndAlso clsCommon.myLen(obj.TAX8) > 0 Then
            objCustInv.TAX8_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
            objCustInv.TAX8_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX8_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX9_Amt > 0 AndAlso clsCommon.myLen(obj.TAX9) > 0 Then
            objCustInv.TAX9_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
            objCustInv.TAX9_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX9_GLAC, obj.Bill_To_Location, trans)
        End If
        If obj.TAX10_Amt > 0 AndAlso clsCommon.myLen(obj.TAX10) > 0 Then
            objCustInv.TAX10_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
            objCustInv.TAX10_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX10_GLAC, obj.Bill_To_Location, trans)
        End If

        'objCustInv.RefDocType=
        objCustInv.RefDocNo = obj.Against_Invoice_No
        objCustInv.Add_Charge_Code1 = obj.Add_Charge_Code1
        objCustInv.Add_Charge_Name1 = obj.Add_Charge_Name1
        objCustInv.Add_Charge_Amt1 = obj.Add_Charge_Amt1
        objCustInv.Add_Charge_Code2 = obj.Add_Charge_Code2
        objCustInv.Add_Charge_Name2 = obj.Add_Charge_Name2
        objCustInv.Add_Charge_Amt2 = obj.Add_Charge_Amt2
        objCustInv.Add_Charge_Code3 = obj.Add_Charge_Code3
        objCustInv.Add_Charge_Name3 = obj.Add_Charge_Name3
        objCustInv.Add_Charge_Amt3 = obj.Add_Charge_Amt3
        objCustInv.Add_Charge_Code4 = obj.Add_Charge_Code4
        objCustInv.Add_Charge_Name4 = obj.Add_Charge_Name4
        objCustInv.Add_Charge_Amt4 = obj.Add_Charge_Amt4
        objCustInv.Add_Charge_Code5 = obj.Add_Charge_Code5
        objCustInv.Add_Charge_Name5 = obj.Add_Charge_Name5
        objCustInv.Add_Charge_Amt5 = obj.Add_Charge_Amt5
        objCustInv.Add_Charge_Code6 = obj.Add_Charge_Code6
        objCustInv.Add_Charge_Name6 = obj.Add_Charge_Name6
        objCustInv.Add_Charge_Amt6 = obj.Add_Charge_Amt6
        objCustInv.Add_Charge_Code7 = obj.Add_Charge_Code7
        objCustInv.Add_Charge_Name7 = obj.Add_Charge_Name7
        objCustInv.Add_Charge_Amt7 = obj.Add_Charge_Amt7
        objCustInv.Add_Charge_Code8 = obj.Add_Charge_Code8
        objCustInv.Add_Charge_Name8 = obj.Add_Charge_Name8
        objCustInv.Add_Charge_Amt8 = obj.Add_Charge_Amt8
        objCustInv.Add_Charge_Code9 = obj.Add_Charge_Code9
        objCustInv.Add_Charge_Name9 = obj.Add_Charge_Name9
        objCustInv.Add_Charge_Amt9 = obj.Add_Charge_Amt9
        objCustInv.Add_Charge_Code10 = obj.Add_Charge_Code10
        objCustInv.Add_Charge_Name10 = obj.Add_Charge_Name10
        objCustInv.Add_Charge_Amt10 = obj.Add_Charge_Amt10
        objCustInv.Total_Add_Charge = obj.Total_Add_Charge
        objCustInv.Tax_Calculation_Type = obj.Tax_Calculation_Type
        ''objCustInv.Status
        ''objCustInv.AgainstScrap
        objCustInv.Against_Sale_Return_No = obj.Document_Code
        Dim counter As Integer = 1
        objCustInv.Arr = New List(Of clsCustomerInvoiceDetail)

        '''' for return Qty
        'Ticket no-ERO/17/12/19-001145,Sample invoice return
        Dim IsSample As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select isnull(IsSampling,0) as IsSampling from TSPL_SD_SALE_INVOICE_HEAD where DOCUMENT_CODE='" & obj.Against_Invoice_No & "'", trans))
        For Each objTr As clsDSSalesReturnDetail In obj.Arr
            If clsCommon.CompairString(objTr.Scheme_Item, "N") = CompairStringResult.Equal Then
                Dim objCustInvTR As clsCustomerInvoiceDetail = New clsCustomerInvoiceDetail()
                If objTr.Return_Amount > 0 OrElse IsSample = True Then
                    objCustInvTR.SNo = counter
                    If clsCommon.CompairString(objTr.Row_Type, "Item") = CompairStringResult.Equal Then
                        dt = clsItemMaster.GetSaleReturnAccGLAC(objTr.Item_Code, trans)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Please set sale account for item" + objTr.Item_Code)
                        End If
                        objCustInvTR.GL_Account_Code = clsCommon.myCstr(dt.Rows(0)("Sales_Return_Account"))
                        objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                        objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                        ''richa agarwal 02 Jan,2019 BHA/03/12/18-000735 send S For Sales Account set Gl Account
                        objCustInvTR.Reco_Control_Account = "S"
                    Else ''for row type misl.
                        Dim objAC As clsAdditionalCharge = clsAdditionalCharge.GetData(objTr.Item_Code, NavigatorType.Current, trans)
                        If objAC Is Nothing Then
                            Throw New Exception("Please set GL Ac from addition charge" + objTr.Item_Code)
                        End If
                        objCustInvTR.GL_Account_Code = objAC.Account_Code
                        objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                        objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                    End If

                    objCustInvTR.Amount = objTr.Return_Amount
                    objCustInvTR.Discount_Per = objTr.Disc_Per
                    objCustInvTR.Discount = objTr.Disc_Amt
                    objCustInvTR.Amount_less_Discount = objTr.Return_Amount - objTr.Disc_Amt
                    objCustInvTR.TAX1 = objTr.TAX1
                    objCustInvTR.TAX1_Rate = objTr.TAX1_Rate
                    objCustInvTR.TAX1_Amt = objTr.TAX1_Amt
                    objCustInvTR.TAX1_Base_Amt = objTr.TAX1_Base_Amt
                    objCustInvTR.TAX2 = objTr.TAX2
                    objCustInvTR.TAX2_Rate = objTr.TAX2_Rate
                    objCustInvTR.TAX2_Amt = objTr.TAX2_Amt
                    objCustInvTR.TAX2_Base_Amt = objTr.TAX2_Base_Amt
                    objCustInvTR.TAX3 = objTr.TAX3
                    objCustInvTR.TAX3_Rate = objTr.TAX3_Rate
                    objCustInvTR.TAX3_Amt = objTr.TAX3_Amt
                    objCustInvTR.TAX3_Base_Amt = objTr.TAX3_Base_Amt
                    objCustInvTR.TAX4 = objTr.TAX4
                    objCustInvTR.TAX4_Rate = objTr.TAX4_Rate
                    objCustInvTR.TAX4_Amt = objTr.TAX4_Amt
                    objCustInvTR.TAX4_Base_Amt = objTr.TAX4_Base_Amt
                    objCustInvTR.TAX5 = objTr.TAX5
                    objCustInvTR.TAX5_Rate = objTr.TAX5_Rate
                    objCustInvTR.TAX5_Amt = objTr.TAX5_Amt
                    objCustInvTR.TAX5_Base_Amt = objTr.TAX5_Base_Amt
                    objCustInvTR.TAX6 = objTr.TAX6
                    objCustInvTR.TAX6_Rate = objTr.TAX6_Rate
                    objCustInvTR.TAX6_Amt = objTr.TAX6_Amt
                    objCustInvTR.TAX6_Base_Amt = objTr.TAX6_Base_Amt
                    objCustInvTR.TAX7 = objTr.TAX7
                    objCustInvTR.TAX7_Rate = objTr.TAX7_Rate
                    objCustInvTR.TAX7_Amt = objTr.TAX7_Amt
                    objCustInvTR.TAX7_Base_Amt = objTr.TAX7_Base_Amt
                    objCustInvTR.TAX8 = objTr.TAX8
                    objCustInvTR.TAX8_Rate = objTr.TAX8_Rate
                    objCustInvTR.TAX8_Amt = objTr.TAX8_Amt
                    objCustInvTR.TAX8_Base_Amt = objTr.TAX8_Base_Amt
                    objCustInvTR.TAX9 = objTr.TAX9
                    objCustInvTR.TAX9_Rate = objTr.TAX9_Rate
                    objCustInvTR.TAX9_Amt = objTr.TAX9_Amt
                    objCustInvTR.TAX9_Base_Amt = objTr.TAX9_Base_Amt
                    objCustInvTR.TAX10 = objTr.TAX10
                    objCustInvTR.TAX10_Rate = objTr.TAX10_Rate
                    objCustInvTR.TAX10_Amt = objTr.TAX10_Amt
                    objCustInvTR.TAX10_Base_Amt = objTr.TAX10_Base_Amt
                    objCustInvTR.Total_Tax = objTr.Total_Tax_Amt
                    objCustInvTR.Total_Amount = objTr.Item_Net_Amt
                    objCustInvTR.Remarks = objTr.Remarks
                    'objCustInvTR.Comments=objTr.Comments
                    objCustInv.Arr.Add(objCustInvTR)
                    counter += 1
                End If
            End If
        Next


        '''' for Damage Qty
        For Each objTr As clsDSSalesReturnDetail In obj.Arr
            If clsCommon.CompairString(objTr.Scheme_Item, "N") = CompairStringResult.Equal Then
                Dim objCustInvTR As clsCustomerInvoiceDetail = New clsCustomerInvoiceDetail()
                If objTr.Damage_Amount > 0 Then
                    objCustInvTR.SNo = counter
                    If clsCommon.CompairString(objTr.Row_Type, "Item") = CompairStringResult.Equal Then
                        dt = clsItemMaster.GetDamageAccGLAC(objTr.Item_Code, trans)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Please set Damage account for item" + objTr.Item_Code)
                        End If
                        objCustInvTR.GL_Account_Code = clsCommon.myCstr(dt.Rows(0)("Damaged_Goods"))
                        objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                        objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                    Else ''for row type misl.
                        Dim objAC As clsAdditionalCharge = clsAdditionalCharge.GetData(objTr.Item_Code, NavigatorType.Current, trans)
                        If objAC Is Nothing Then
                            Throw New Exception("Please set GL Ac from addition charge" + objTr.Item_Code)
                        End If
                        objCustInvTR.GL_Account_Code = objAC.Account_Code
                        objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                        objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                    End If

                    objCustInvTR.Amount = objTr.Damage_Amount
                    objCustInvTR.Discount_Per = objTr.Disc_Per
                    objCustInvTR.Discount = objTr.Disc_Amt
                    objCustInvTR.Amount_less_Discount = objTr.Damage_Amount
                    objCustInvTR.TAX1 = objTr.TAX1
                    objCustInvTR.TAX1_Rate = objTr.TAX1_Rate
                    objCustInvTR.TAX1_Amt = objTr.TAX1_Amt
                    objCustInvTR.TAX2 = objTr.TAX2
                    objCustInvTR.TAX2_Rate = objTr.TAX2_Rate
                    objCustInvTR.TAX2_Amt = objTr.TAX2_Amt
                    objCustInvTR.TAX3 = objTr.TAX3
                    objCustInvTR.TAX3_Rate = objTr.TAX3_Rate
                    objCustInvTR.TAX3_Amt = objTr.TAX3_Amt
                    objCustInvTR.TAX4 = objTr.TAX4
                    objCustInvTR.TAX4_Rate = objTr.TAX4_Rate
                    objCustInvTR.TAX4_Amt = objTr.TAX4_Amt
                    objCustInvTR.TAX5 = objTr.TAX5
                    objCustInvTR.TAX5_Rate = objTr.TAX5_Rate
                    objCustInvTR.TAX5_Amt = objTr.TAX5_Amt
                    objCustInvTR.TAX6 = objTr.TAX6
                    objCustInvTR.TAX6_Rate = objTr.TAX6_Rate
                    objCustInvTR.TAX6_Amt = objTr.TAX6_Amt
                    objCustInvTR.TAX7 = objTr.TAX7
                    objCustInvTR.TAX7_Rate = objTr.TAX7_Rate
                    objCustInvTR.TAX7_Amt = objTr.TAX7_Amt
                    objCustInvTR.TAX8 = objTr.TAX8
                    objCustInvTR.TAX8_Rate = objTr.TAX8_Rate
                    objCustInvTR.TAX8_Amt = objTr.TAX8_Amt
                    objCustInvTR.TAX9 = objTr.TAX9
                    objCustInvTR.TAX9_Rate = objTr.TAX9_Rate
                    objCustInvTR.TAX9_Amt = objTr.TAX9_Amt
                    objCustInvTR.TAX10 = objTr.TAX10
                    objCustInvTR.TAX10_Rate = objTr.TAX10_Rate
                    objCustInvTR.TAX10_Amt = objTr.TAX10_Amt
                    objCustInvTR.Total_Tax = objTr.Total_Tax_Amt
                    objCustInvTR.Total_Amount = objTr.Item_Net_Amt
                    objCustInvTR.Remarks = objTr.Remarks
                    'objCustInvTR.Comments=objTr.Comments
                    objCustInv.Arr.Add(objCustInvTR)
                    counter += 1
                End If
            End If
        Next


        '''''  For Price only return type
        If clsCommon.CompairString(obj.Return_Type, "P") = CompairStringResult.Equal Then
            For Each objTr As clsDSSalesReturnDetail In obj.Arr
                If clsCommon.CompairString(objTr.Scheme_Item, "N") = CompairStringResult.Equal Then
                    Dim objCustInvTR As clsCustomerInvoiceDetail = New clsCustomerInvoiceDetail()
                    If objTr.Return_Amount = 0 AndAlso objTr.Damage_Amount = 0 AndAlso objTr.Amt_Less_Discount > 0 Then
                        objCustInvTR.SNo = counter
                        If clsCommon.CompairString(objTr.Row_Type, "Item") = CompairStringResult.Equal Then
                            dt = clsItemMaster.GetSaleReturnAccGLAC(objTr.Item_Code, trans)
                            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                                Throw New Exception("Please set sale account for item" + objTr.Item_Code)
                            End If
                            objCustInvTR.GL_Account_Code = clsCommon.myCstr(dt.Rows(0)("Sales_Return_Account"))
                            objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                            objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                            ''richa agarwal 02 Jan,2019 BHA/03/12/18-000735 send S For Sales Account set Gl Account
                            objCustInvTR.Reco_Control_Account = "S"
                        Else ''for row type misl.
                            Dim objAC As clsAdditionalCharge = clsAdditionalCharge.GetData(objTr.Item_Code, NavigatorType.Current, trans)
                            If objAC Is Nothing Then
                                Throw New Exception("Please set GL Ac from addition charge" + objTr.Item_Code)
                            End If
                            objCustInvTR.GL_Account_Code = objAC.Account_Code
                            objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Bill_To_Location, trans)
                            objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                        End If

                        objCustInvTR.Amount = objTr.Amt_Less_Discount
                        objCustInvTR.Discount_Per = objTr.Disc_Per
                        objCustInvTR.Discount = objTr.Disc_Amt
                        objCustInvTR.Amount_less_Discount = objTr.Amt_Less_Discount
                        objCustInvTR.TAX1 = objTr.TAX1
                        objCustInvTR.TAX1_Rate = objTr.TAX1_Rate
                        objCustInvTR.TAX1_Amt = objTr.TAX1_Amt
                        objCustInvTR.TAX2 = objTr.TAX2
                        objCustInvTR.TAX2_Rate = objTr.TAX2_Rate
                        objCustInvTR.TAX2_Amt = objTr.TAX2_Amt
                        objCustInvTR.TAX3 = objTr.TAX3
                        objCustInvTR.TAX3_Rate = objTr.TAX3_Rate
                        objCustInvTR.TAX3_Amt = objTr.TAX3_Amt
                        objCustInvTR.TAX4 = objTr.TAX4
                        objCustInvTR.TAX4_Rate = objTr.TAX4_Rate
                        objCustInvTR.TAX4_Amt = objTr.TAX4_Amt
                        objCustInvTR.TAX5 = objTr.TAX5
                        objCustInvTR.TAX5_Rate = objTr.TAX5_Rate
                        objCustInvTR.TAX5_Amt = objTr.TAX5_Amt
                        objCustInvTR.TAX6 = objTr.TAX6
                        objCustInvTR.TAX6_Rate = objTr.TAX6_Rate
                        objCustInvTR.TAX6_Amt = objTr.TAX6_Amt
                        objCustInvTR.TAX7 = objTr.TAX7
                        objCustInvTR.TAX7_Rate = objTr.TAX7_Rate
                        objCustInvTR.TAX7_Amt = objTr.TAX7_Amt
                        objCustInvTR.TAX8 = objTr.TAX8
                        objCustInvTR.TAX8_Rate = objTr.TAX8_Rate
                        objCustInvTR.TAX8_Amt = objTr.TAX8_Amt
                        objCustInvTR.TAX9 = objTr.TAX9
                        objCustInvTR.TAX9_Rate = objTr.TAX9_Rate
                        objCustInvTR.TAX9_Amt = objTr.TAX9_Amt
                        objCustInvTR.TAX10 = objTr.TAX10
                        objCustInvTR.TAX10_Rate = objTr.TAX10_Rate
                        objCustInvTR.TAX10_Amt = objTr.TAX10_Amt
                        objCustInvTR.Total_Tax = objTr.Total_Tax_Amt
                        objCustInvTR.Total_Amount = objTr.Item_Net_Amt
                        objCustInvTR.Remarks = objTr.Remarks
                        'objCustInvTR.Comments=objTr.Comments
                        objCustInv.Arr.Add(objCustInvTR)
                        counter += 1
                    End If
                End If
            Next
        End If
        ''richa BHA/01/03/19-000830 send form id 06 March,2019
        objCustInv.SaveData(objCustInv, True, trans, clsUserMgtCode.frmSaleReturndairy, strVoucherForRecreate, strARNoForRecreate)
        clsCustomerInvoiceHead.PostData(clsUserMgtCode.frmSaleReturndairy, objCustInv.Document_No, "", trans)
        Return True
    End Function

    Private Shared Function GetFirstItemCode(ByVal Arr As List(Of clsDSSalesReturnDetail)) As String
        For Each objtr As clsDSSalesReturnDetail In Arr
            If clsCommon.CompairString(objtr.Row_Type, clsItemRowType.RowTypeItem) = CompairStringResult.Equal Then
                Return objtr.Item_Code
            End If
        Next
        Return ""
    End Function

    Public Shared Function DeleteData(ByVal strCode As String) As Boolean
        Dim Trans_type_Str As String = ""
        Dim isSaved As Boolean = False
        If (clsCommon.myLen(strCode) <= 0) Then
            Throw New Exception("Purchase Order No not found to Delete")
        End If
        Dim obj As clsDSSalesReturnHead = clsDSSalesReturnHead.GetData(strCode, NavigatorType.Current)
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Trans_type_Str = clsCommon.myCstr(obj.Trans_type) + "-SR"
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
            Try
                '' Anubhooti 06-Sep-2014 BM00000003735 (Remarks: Locked Transaction)
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleSaleDairy, clsUserMgtCode.frmSaleReturndairy, obj.Bill_To_Location, obj.Document_Date, trans)
                ''
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SALE_RETURN_HEAD", "Document_Code", "TSPL_SD_SALE_RETURN_DETAIL", "Document_Code", trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)

                If (obj.Status = 1) Then
                    Throw New Exception("Already Posted on :" + obj.Posting_Date)
                End If
                clsSerializeInvenotry.DeleteData("Sale Return", strCode, trans)
                clsBatchInventory.DeleteData(Trans_type_Str, strCode, trans)
                Dim qry As String = "delete from TSPL_SD_SALE_RETURN_DETAIL where Document_Code='" & strCode & "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SD_SALE_RETURN_BOOKING_DETAIL where Document_Code='" & strCode & "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & strCode & "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                isSaved = isSaved AndAlso clsCustomFieldValues.DeleteData(obj.Form_ID, strCode, trans)

                If (isSaved) Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            Catch ex As Exception
                trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function

    Public Shared Function IsValidCustomer(ByVal Arr As List(Of String), ByVal strVendorCode As String) As Boolean
        If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
            Dim qry As String = "select Document_Code,Customer_Code,Customer_Name from TSPL_SD_SALE_RETURN_HEAD where Document_Code in (" + clsCommon.GetMulcallString(Arr) + ") and Customer_Code not in ('" + strVendorCode + "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim msg As String = "Error. "
                For Each dr As DataRow In dt.Rows
                    msg += Environment.NewLine + "SRN No:" + clsCommon.myCstr(dr("Document_Code")) + " Is For Vendor Code: " + clsCommon.myCstr(dr("Customer_Code")) + " Vendor Name:" + clsCommon.myCstr(dr("Customer_Name"))
                Next
                Throw New Exception(msg)
            End If
        End If
        Return True
    End Function
    ''To be Uncomment
    '    Public Sub SRNPrintOut(ByVal FromDate As Date?, ByVal ToDate As Date?, ByVal IsDocTypeFinsihGoods As Boolean, ByVal ArrSrnNo As ArrayList, ByVal ArrVendor As ArrayList, ByVal ArrLocation As ArrayList)
    '        Dim qry As String

    '        Try
    '            If IsDocTypeFinsihGoods Then
    '                qry = "select Document_Code,MAX(ItemType )as ItemType,MAX(MRN_Date) as Document_Date,MAX(Customer_Name) as Customer_Name,MAX(GRNo) as GRNo,MAX(GENo) as GENo,MAX(GEDate) as GEDate,Item_Code,MAX(Item_Desc) as Item_Desc,MAX(VehicleNo) as VehicleNo, SUM(ISNULL( FCS,0)) as FCS, SUM(isnull(FBS,0))as FBS, SUM(ISNULL( FSH,0)) as FSH, SUM(ISNULL( ECS,0)) as ECS, SUM(ISNULL( EBS,0)) as EBS, SUM(Leak_Qty) as HF,SUM(Burst_Qty) as Burst,SUM(Short_Qty) as Short,MAX(Remarks) as Remarks,max(Ref_No)as Ref_No from( " & _
    '         "select TSPL_SD_SALE_RETURN_HEAD.Document_Code,TSPL_SD_SALE_RETURN_HEAD .Item_Type as ItemType," & _
    '         "(replace( CONVERT(varchar(11), TSPL_SD_SALE_RETURN_HEAD.Document_Date,104),'.','/')+' '+CONVERT(varchar(100),TSPL_SD_SALE_RETURN_HEAD.Document_Date,108) )as MRN_Date,TSPL_SD_SALE_RETURN_HEAD.Customer_Name,TSPL_SD_SALE_RETURN_HEAD.GRNo,TSPL_SD_SALE_RETURN_HEAD.GENo," & _
    '         "(case when LEN(TSPL_SD_SALE_RETURN_HEAD.GEDate)>0  then REPLACE( CONVERT(varchar(11), TSPL_SD_SALE_RETURN_HEAD.GEDate,104),'.','/') else '' end) as GEDate,TSPL_SD_SALE_RETURN_HEAD.VehicleNo,TSPL_SD_SALE_RETURN_HEAD.Remarks ,TSPL_SD_SALE_RETURN_HEAD.Ref_No,TSPL_SD_SALE_RETURN_DETAIL.Item_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Desc,TSPL_SD_SALE_RETURN_DETAIL.Unit_code," & _
    '         "case when Unit_code='FC' then Qty + ISNULL( Free_Qty,0) end as FCS, " & _
    '         "case when Unit_code='FB' then Qty + ISNULL( Free_Qty,0) end as FBS, " & _
    '         "case when Unit_code='SH' then Qty + ISNULL( Free_Qty,0) end as FSH, " & _
    '         "case when Unit_code='EC' then Qty + ISNULL( Free_Qty,0) end as ECS," & _
    '         "case when Unit_code='EB' then Qty + ISNULL( Free_Qty,0) end as EBS, " & _
    '         "TSPL_SD_SALE_RETURN_DETAIL.Leak_Qty,TSPL_SD_SALE_RETURN_DETAIL.Burst_Qty,TSPL_SD_SALE_RETURN_DETAIL.Short_Qty from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code= TSPL_SD_SALE_RETURN_DETAIL.Document_Code " & _
    '         " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location   where Item_Type ='F'"
    '                If FromDate.HasValue AndAlso ToDate.HasValue Then
    '                    qry += " and Convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)<=Convert(date,'" + ToDate + "',103) "
    '                End If

    '                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
    '                    qry += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
    '                End If
    '                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
    '                    qry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
    '                End If
    '                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
    '                    qry += " and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")" 'ADDED BY ABHISHEK AS ON 30 AUG 2012
    '                End If
    '                qry += " )xxx group by Document_Code,Item_Code order by Item_Desc"
    '                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    '                    common.clsCommon.MyMessageBoxShow("No Record Found")
    '                Else
    '                    PurchaseOrderViewer.funreport(dt, EnumTecxpertPaperSize.PaperSize10x6, "rptSRNCustomReport", "SRN Report")

    '                End If
    '            Else ''For RM Other Print out
    '                Dim strquery As String = "SELECT TSPL_SD_SALE_RETURN_HEAD.Document_Code, TSPL_SD_SALE_RETURN_HEAD.Document_Date,TSPL_SD_SALE_RETURN_HEAD.Customer_Name,(case when len(against_mrn)>0 then (select MRN_Date  from tspl_mrn_head where tspl_mrn_head.MRN_No =against_mrn) else Document_Date end ) as Challan_Date, TSPL_SD_SALE_RETURN_HEAD.Ref_No  " & _
    '                      "as Challan_No, TSPL_SD_SALE_RETURN_HEAD.Inv_No, TSPL_SD_SALE_RETURN_HEAD.Inv_Date, TSPL_SD_SALE_RETURN_HEAD.GRNo,TSPL_SD_SALE_RETURN_HEAD.Amount_Less_Discount ,TSPL_SD_SALE_RETURN_HEAD.GENo,TSPL_SD_SALE_RETURN_HEAD.Total_Amt, " & _
    '                      "TSPL_SD_SALE_RETURN_HEAD.GEDate, TSPL_SD_SALE_RETURN_HEAD.VehicleNo, TSPL_SD_SALE_RETURN_HEAD.Carrier,TSPL_SD_SALE_RETURN_HEAD.Remarks,TSPL_SD_SALE_RETURN_DETAIL.Landed_Cost_Rate,TSPL_SD_SALE_RETURN_DETAIL.Landed_Cost_Amount , TSPL_SD_SALE_RETURN_DETAIL.Item_Code,TSPL_SD_SALE_RETURN_DETAIL.Row_Type,TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount," & _
    '"TSPL_SD_SALE_RETURN_DETAIL.Item_Cost as basicRate,TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt as BasicTotal,TSPL_SD_SALE_RETURN_DETAIL.Unit_Cost_Tax_Rate as UCTR," & _
    '"TSPL_SD_SALE_RETURN_DETAIL.Unit_Cost_Tax as uctax,TSPL_SD_SALE_RETURN_DETAIL.Item_Desc,TSPL_SD_SALE_RETURN_DETAIL.Unit_code,TSPL_SD_SALE_RETURN_DETAIL.Qty,TSPL_SD_SALE_RETURN_DETAIL.Rejected_Qty,TSPL_SD_SALE_RETURN_HEAD.Customer_Code,TSPL_SD_SALE_RETURN_HEAD.Total_Amt,TSPL_SD_SALE_RETURN_DETAIL.ITEM_COST," & _
    ' "TSPL_VENDOR_MASTER.Add1 as venAdd1, TSPL_VENDOR_MASTER.Add2 as vanadd2, TSPL_VENDOR_MASTER.Add3 as venadd3, " & _
    '"tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax1_amt,0) as txt1amt,tax2.Tax_Code_Desc as tax2name," & _
    '"isnull (TSPL_SD_SALE_RETURN_HEAD.tax2_amt,0) as txt2amt,tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax3_amt,0) as txt3amt," & _
    '"tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax4_amt,0) as txt4amt,tax5.Tax_Code_Desc as tax5name," & _
    '"isnull (TSPL_SD_SALE_RETURN_HEAD.tax5_amt,0) as txt5amt,tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax6_amt,0) as txt6amt " & _
    '",tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax7_amt,0) as txt7amt,tax8.Tax_Code_Desc as tax8name," & _
    '"isnull (TSPL_SD_SALE_RETURN_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax9_amt,0) as txt9amt," & _
    '"tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax10_amt,0) as txt10amt, TSPL_COMPANY_MASTER.Comp_Name as compname, " & _
    '"TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_SD_SALE_RETURN_DETAIL.Qty," & _
    '"case when tax1.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax1_amt else null end as Tax1Recoverable," & _
    '"case when tax2.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.TAX2_Amt else null end as Tax2Recoverable, " & _
    '"case when tax3.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax3_amt else null end as Tax3Recoverable, " & _
    '"case when tax4.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax4_amt else null end as Tax4Recoverable, " & _
    '"case when tax5.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax5_amt else null end as Tax5Recoverable, " & _
    '"case when tax6.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax6_amt else null end as Tax6Recoverable," & _
    '"case when tax7.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax7_amt else null end as Tax7Recoverable, " & _
    '"case when tax8.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax8_amt else null end as Tax8Recoverable, " & _
    '"case when tax9.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax9_amt else null end as Tax9Recoverable," & _
    '"case when tax10.Tax_Recoverable='Y' then TSPL_SD_SALE_RETURN_HEAD.tax10_amt else null end as Tax10Recoverable, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX1_Rate ,0),103)+'%' as txt1Rate," & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX2_Rate   ,0),103)+'%' as txt2Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX3_Rate  ,0),103)+'%' as txt3Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX4_Rate  ,0),103)+'%' as txt4Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX5_Rate  ,0),103)+'%' as txt5Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX6_Rate  ,0),103)+'%' as txt6Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX7_Rate  ,0),103)+'%' as txt7Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX8_Rate  ,0),103)+'%' as txt8Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX9_Rate  ,0),103)+'%' as txt9Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SALE_RETURN_HEAD.TAX10_Rate  ,0),103)+'%' as txt10Rate," & _
    '"TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount as Value,(select SUM(rejected_qty) from TSPL_SD_SALE_RETURN_DETAIL where Document_Code=TSPL_SD_SALE_RETURN_HEAD.Document_Code) as Rej_qty, (select SUM(TSPL_MRN_DETAIL.MRN_Qty) from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_MRN_DETAIL on TSPL_MRN_DETAIL .MRN_No=TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code and TSPL_MRN_DETAIL.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code where Document_Code =TSPL_SD_SALE_RETURN_HEAD.Document_Code)as MrnTotQty, (select SUM(Qty) from TSPL_SD_SALE_RETURN_DETAIL where Document_Code=TSPL_SD_SALE_RETURN_HEAD.Document_Code) as SRNQtyTotal, (select case when COUNT(xxx.PI_No)>1 then Min(xxx.PI_No)+ ' *' else Min(xxx.PI_No)end as PINO from" & _
    '" ( select TSPL_PI_DETAIL.PI_No from TSPL_PI_DETAIL  where  TSPL_PI_DETAIL.SRN_Id= TSPL_SD_SALE_RETURN_HEAD.Document_Code " & _
    '" GROUP by TSPL_PI_DETAIL.PI_No)xxx) as PInvNo  ,    " & _
    '       " TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name1 as Add1Name, " & _
    '     " TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt1 as Add1 , " & _
    '     "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name2 as Add2Name, " & _
    '     "   TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt2 as Add2 , " & _
    '     "    TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name3 as Add3Name, " & _
    '     "   TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt3 as Add3 , " & _
    '     "    TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name4 as Add4Name, " & _
    '     "    TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt4 as Add4 , " & _
    '     "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name5 as Add5Name, " & _
    '      "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt5 as Add5 , " & _
    '      "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name6 as Add6Name, " & _
    '      "    TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt6 as Add6 , " & _
    '      "    TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name7 as Add7Name, " & _
    '      "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt7 as Add7 , " & _
    '      "       TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name8 as Add8Name, " & _
    '      "      TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt8 as Add8 , " & _
    '       "      TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name9 as Add9Name, " & _
    '       "      TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt9 as Add9 , " & _
    '       "      TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Name10 as Add10Name, " & _
    '       "     TSPL_SD_SALE_RETURN_HEAD.Add_Charge_Amt10 as Add10,TSPL_SD_SALE_RETURN_HEAD.Against_RGP,TSPL_SD_SALE_RETURN_DETAIL .Specification   " & _
    ' " FROM  TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.Document_Code = TSPL_SD_SALE_RETURN_HEAD.Document_Code " & _
    ' "INNER JOIN TSPL_COMPANY_MASTER ON TSPL_SD_SALE_RETURN_HEAD.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code  " & _
    ' "INNER JOIN TSPL_VENDOR_MASTER ON TSPL_SD_SALE_RETURN_HEAD.Customer_Code = TSPL_VENDOR_MASTER.Customer_Code " & _
    ' "left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_RETURN_HEAD.tax1  " & _
    ' "left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_RETURN_HEAD.tax2 " & _
    ' "left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .TAX3 " & _
    ' "left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_RETURN_HEAD .tax4 " & _
    ' "left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .tax5 " & _
    ' "left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX6  " & _
    ' "left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX7  " & _
    ' "left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX8 " & _
    ' "left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX9 " & _
    ' " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX10  " & _
    ' "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location  " & _
    ' " where TSPL_SD_SALE_RETURN_HEAD .Item_Type not in('F')"

    '                If FromDate.HasValue AndAlso ToDate.HasValue Then
    '                    strquery += " and Convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)<=Convert(date,'" + ToDate + "',103) "

    '                End If
    '                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
    '                    strquery += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
    '                End If
    '                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
    '                    strquery += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
    '                End If
    '                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
    '                    strquery += " and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")  "

    '                End If
    '                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strquery)
    '                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    '                    common.clsCommon.MyMessageBoxShow("No Record Found")
    '                Else
    '                    PurchaseOrderViewer.funreport(dt, "SRNReportThroughReport", "Store Receipt Report")
    '                End If
    '            End If

    '        Catch ex As Exception
    '            Throw New Exception(ex.Message)
    '        End Try
    '    End Sub

End Class

Public Class clsDSSalesReturnDetail
#Region "Variables"
    Public CAN As Double = 0
    Public Crate As Double = 0
    Public ItemwiseTaxCode As String = ""
    Public Alter_UnitQty As Double = 0
    Public Rate_UnitQty As Double = 0
    Public ActuaQty As Double = 0
    Public ActualReturnQty As Double = 0
    Public InvoiceScheme_Code As String = Nothing
    Public InvoiceCashScheme_Code As String = Nothing
    Public Alternate_UOM As String = Nothing
    Public RATE_UOM As String = Nothing
    Public ActualUOM As String = Nothing
    Public Item_Group As String = Nothing
    Public TAX_PAID As String = Nothing
    Public Scheme_Type As String = Nothing
    Public Scheme_Item_Code As String = Nothing
    Public Scheme_Qty As Decimal = Nothing
    Public Scheme_Item_UOM As String = Nothing
    Public Cash_Scheme_Code As String = Nothing
    Public Cash_Scheme_Type As String = Nothing
    Public Cash_Scheme_Pers As Decimal = Nothing
    Public Cash_Scheme_Amount As Decimal = Nothing

    Public OrgUnit_code As String = ""
    Public Document_Code As String = Nothing
    Public Line_No As Integer = 0
    Public Row_Type As String = Nothing
    Public Item_Code As String = Nothing
    Public Item_Desc As String = Nothing 'Not a Table Field
    Public Qty As Double = 0
    Public Balance_Qty As Double = 0
    Public Free_Qty As Double = 0
    Public Invoice_Code As String = Nothing
    Public Unit_code As String = Nothing '
    Public Location As String = Nothing '
    Public LocationName As String = Nothing 'Not a Table Field
    Public Item_Cost As Double = 0
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
    Public Disc_Per As Double = 0
    Public Disc_Amt As Double = 0
    Public Amt_Less_Discount As Double = 0
    Public Total_Tax_Amt As Double = 0
    Public Item_Net_Amt As Double = 0
    Public Status As Integer = 0
    Public MRP As Double = 0
    Public MFG_Date As Date? = Nothing
    Public Batch_No As String = Nothing
    Public Expiry_Date As Date? = Nothing
    Public Specification As String = Nothing
    Public Remarks As String = Nothing
    Public Assessable As Double = 0
    Public AssessableAmt As Double = 0
    Public Is_Mannual_Amt As Integer = Nothing
    Public SRNTax_Group As String = Nothing 'Not a Table Field

    Public Scheme_Applicable As String = Nothing
    Public Scheme_Code As String = Nothing
    Public Scheme_Item As String = Nothing
    Public Item_Tax As Double = 0
    Public Total_MRP_Amt As Double = 0
    Public Total_Basic_Amt As Double = 0
    Public Total_Disc_Amt As Double = 0
    Public Price_with_tax As Double = 0
    Public Cust_Discount As Double = 0
    Public Total_Cust_Discount As Double = 0
    Public Price_Amount1 As Double = 0
    Public Price_Amount2 As Double = 0
    Public Price_Amount3 As Double = 0
    Public Price_Amount4 As Double = 0
    Public Price_Amount5 As Double = 0
    Public Price_Amount6 As Double = 0
    Public Price_Amount7 As Double = 0
    Public Price_Amount8 As Double = 0
    Public Price_Amount9 As Double = 0
    Public Price_Amount10 As Double = 0
    Public ActualRate As Double = 0
    Public Cust_DiscountQty As Double = 0
    Public Price_code As String = Nothing
    Public Abatement_Per As Double = 0
    Public Abatement_Amt As Double = 0
    Public FOC_Item As Double = 0
    Public Price_Date As Date? = Nothing
    Public Item_Weight As Double = 0
    Public Conv_Factor As Double = 0
    Public TotalItem_Weight As Double = 0
    Public Markup_On As String = Nothing
    Public Markup_Percent As Double = 0
    Public Landing_Cost As Double = 0
    Public HeadDiscAmt As Double = 0
    Public CustDiscPer As Double = 0
    Public CasdDiscScheme_Code As String = Nothing
    Public Purchase_Cost As Double = 0
    Public OrgRate As Double = 0
    Public DamageQty As Double = 0
    Public Return_Amount As Double = 0
    Public Damage_Amount As Double = 0
    Public Bin_No As String = Nothing
    Public PrincipleCode As String = Nothing
    Public PrincipleDesc As String = Nothing
    Public vendor_code As String = Nothing
    Public vendor_desc As String = Nothing
    Public HeadDiscPer As Double = 0
    Public HeadDiscPerAmt As Double = 0
    Public Delivery_Code As String = Nothing
    Public Sampling As Integer = 0
    Public Distributor_Commission_PKID As String = ""
    Public Distributor_Commission_Rate As Decimal = 0
    Public Distributor_Commission_RateWithTax As Decimal = 0
    Public Distributor_Commission_Amt As Decimal = 0
    Public Transporter_Commission_Rate As Decimal = 0
    Public Transporter_Commission_Amt As Decimal = 0
    Public Security_Rate As Decimal = 0
    Public Security_Amt As Decimal = 0
    Public Transporter As String = ""
    Public arrSrItem As List(Of clsSerializeInvenotry) = Nothing
    Public arrBatchItem As List(Of clsBatchInventory) = Nothing
    Public Against_Dispatch As String = ""
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal dtDocDate As DateTime, ByVal Arr As List(Of clsDSSalesReturnDetail), ByVal Trans_Type As String, ByVal trans As SqlTransaction) As Boolean
        Dim Trans_type_Str As String = ""
        Trans_type_Str = clsCommon.myCstr(Trans_Type) + "-SR"
        Try
            If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
                For Each obj As clsDSSalesReturnDetail In Arr
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Sampling", obj.Sampling)
                    clsCommon.AddColumnsForChange(coll, "Delivery_Code", obj.Delivery_Code, True)
                    clsCommon.AddColumnsForChange(coll, "CAN", obj.CAN)
                    clsCommon.AddColumnsForChange(coll, "Crate", obj.Crate)
                    clsCommon.AddColumnsForChange(coll, "ItemwiseTaxCode", obj.ItemwiseTaxCode)
                    clsCommon.AddColumnsForChange(coll, "Alter_UnitQty", obj.Alter_UnitQty)
                    clsCommon.AddColumnsForChange(coll, "Rate_UnitQty", obj.Rate_UnitQty)
                    clsCommon.AddColumnsForChange(coll, "InvoiceScheme_Code", obj.InvoiceScheme_Code)
                    clsCommon.AddColumnsForChange(coll, "InvoiceCashScheme_Code", obj.InvoiceCashScheme_Code)
                    clsCommon.AddColumnsForChange(coll, "Cash_Scheme_Code", obj.Cash_Scheme_Code)
                    clsCommon.AddColumnsForChange(coll, "Cash_Scheme_Type", obj.Cash_Scheme_Type)
                    clsCommon.AddColumnsForChange(coll, "Cash_Scheme_Pers", obj.Cash_Scheme_Pers)
                    clsCommon.AddColumnsForChange(coll, "Cash_Scheme_Amount", obj.Cash_Scheme_Amount)

                    clsCommon.AddColumnsForChange(coll, "Scheme_Item_Code", obj.Scheme_Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Scheme_Item_UOM", obj.Scheme_Item_UOM)
                    clsCommon.AddColumnsForChange(coll, "Scheme_Qty", obj.Scheme_Qty)
                    clsCommon.AddColumnsForChange(coll, "Scheme_Type", obj.Scheme_Type)
                    clsCommon.AddColumnsForChange(coll, "Alternate_UOM", obj.Alternate_UOM, True)
                    clsCommon.AddColumnsForChange(coll, "RATE_UOM", obj.RATE_UOM, True)
                    clsCommon.AddColumnsForChange(coll, "Item_Group", obj.Item_Group)
                    clsCommon.AddColumnsForChange(coll, "TAX_PAID", obj.TAX_PAID)

                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                    clsCommon.AddColumnsForChange(coll, "Row_Type", obj.Row_Type)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                    clsCommon.AddColumnsForChange(coll, "OrgUnit_code", obj.OrgUnit_code)
                    clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)

                    clsCommon.AddColumnsForChange(coll, "Free_qty", obj.Free_Qty)

                    clsCommon.AddColumnsForChange(coll, "Invoice_Code", obj.Invoice_Code, True)

                    clsCommon.AddColumnsForChange(coll, "Balance_Qty", obj.Balance_Qty)
                    clsCommon.AddColumnsForChange(coll, "Unit_code", obj.Unit_code)
                    clsCommon.AddColumnsForChange(coll, "Location", obj.Location)
                    clsCommon.AddColumnsForChange(coll, "Item_Cost", obj.Item_Cost)

                    clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                    clsCommon.AddColumnsForChange(coll, "Disc_Per", obj.Disc_Per)
                    clsCommon.AddColumnsForChange(coll, "Disc_Amt", obj.Disc_Amt)
                    clsCommon.AddColumnsForChange(coll, "Amt_Less_Discount", obj.Amt_Less_Discount)
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
                    clsCommon.AddColumnsForChange(coll, "Batch_No", obj.Batch_No, True)

                    clsCommon.AddColumnsForChange(coll, "Specification", obj.Specification)
                    clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
                    clsCommon.AddColumnsForChange(coll, "Is_Mannual_Amt", obj.Is_Mannual_Amt)
                    If obj.MFG_Date.HasValue Then
                        clsCommon.AddColumnsForChange(coll, "MFG_Date", clsCommon.GetPrintDate(obj.MFG_Date, "dd-MMM-yyyy"))
                    End If
                    If obj.Expiry_Date.HasValue Then
                        clsCommon.AddColumnsForChange(coll, "Expiry_Date", clsCommon.GetPrintDate(obj.Expiry_Date, "dd-MMM-yyyy"))
                    End If
                    clsCommon.AddColumnsForChange(coll, "Assessable", obj.Assessable)
                    clsCommon.AddColumnsForChange(coll, "AssessableAmt", obj.AssessableAmt)

                    clsCommon.AddColumnsForChange(coll, "Scheme_Applicable", IIf(obj.Scheme_Applicable = "Yes", "Y", "N"))
                    clsCommon.AddColumnsForChange(coll, "Scheme_Code", obj.Scheme_Code)
                    clsCommon.AddColumnsForChange(coll, "Scheme_Item", IIf(obj.Scheme_Item = "Yes", "Y", "N"))
                    clsCommon.AddColumnsForChange(coll, "Item_Tax", obj.Item_Tax)
                    clsCommon.AddColumnsForChange(coll, "Total_MRP_Amt", obj.Total_MRP_Amt)
                    clsCommon.AddColumnsForChange(coll, "Total_Basic_Amt", obj.Total_Basic_Amt)
                    clsCommon.AddColumnsForChange(coll, "Total_Disc_Amt", obj.Total_Disc_Amt)
                    clsCommon.AddColumnsForChange(coll, "Cust_Discount", obj.Cust_Discount)
                    clsCommon.AddColumnsForChange(coll, "Total_Cust_Discount", obj.Total_Cust_Discount)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount1", obj.Price_Amount1)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount2", obj.Price_Amount2)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount3", obj.Price_Amount3)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount4", obj.Price_Amount4)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount5", obj.Price_Amount5)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount6", obj.Price_Amount6)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount7", obj.Price_Amount7)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount8", obj.Price_Amount8)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount9", obj.Price_Amount9)
                    clsCommon.AddColumnsForChange(coll, "Price_Amount10", obj.Price_Amount10)
                    clsCommon.AddColumnsForChange(coll, "ActualRate", obj.ActualRate)
                    clsCommon.AddColumnsForChange(coll, "Cust_DiscountQty", obj.Cust_DiscountQty)
                    clsCommon.AddColumnsForChange(coll, "Price_code", obj.Price_code)
                    clsCommon.AddColumnsForChange(coll, "Price_Date", clsCommon.GetPrintDate(obj.Price_Date, "dd/MMM/yyyy"))
                    clsCommon.AddColumnsForChange(coll, "Abatement_Per", obj.Abatement_Per)
                    clsCommon.AddColumnsForChange(coll, "Abatement_Amt", obj.Abatement_Amt)
                    clsCommon.AddColumnsForChange(coll, "FOC_Item", obj.FOC_Item)

                    clsCommon.AddColumnsForChange(coll, "Item_Weight", obj.Item_Weight)
                    clsCommon.AddColumnsForChange(coll, "Conv_Factor", obj.Conv_Factor)
                    clsCommon.AddColumnsForChange(coll, "TotalItem_Weight", obj.TotalItem_Weight)
                    clsCommon.AddColumnsForChange(coll, "Purchase_Cost", obj.Purchase_Cost)
                    clsCommon.AddColumnsForChange(coll, "OrgRate", obj.OrgRate)
                    clsCommon.AddColumnsForChange(coll, "DamageQty", obj.DamageQty)
                    clsCommon.AddColumnsForChange(coll, "Return_Amount", obj.Return_Amount)
                    clsCommon.AddColumnsForChange(coll, "Damage_Amount", obj.Damage_Amount)
                    clsCommon.AddColumnsForChange(coll, "Bin_No", obj.Bin_No)
                    clsCommon.AddColumnsForChange(coll, "PrincipleCode", obj.PrincipleCode)
                    clsCommon.AddColumnsForChange(coll, "PrincipleDesc", obj.PrincipleDesc)
                    clsCommon.AddColumnsForChange(coll, "vendor_code", obj.vendor_code)
                    clsCommon.AddColumnsForChange(coll, "vendor_desc", obj.vendor_desc)
                    clsCommon.AddColumnsForChange(coll, "HeadDiscAmt", obj.HeadDiscAmt)
                    clsCommon.AddColumnsForChange(coll, "HeadDiscPer", obj.HeadDiscPer)
                    clsCommon.AddColumnsForChange(coll, "HeadDiscPerAmt", obj.HeadDiscPerAmt)
                    clsCommon.AddColumnsForChange(coll, "ActualUOM", obj.ActualUOM)
                    clsCommon.AddColumnsForChange(coll, "ActualQty", obj.ActuaQty)
                    clsCommon.AddColumnsForChange(coll, "ActualConvAmt", obj.ActualReturnQty)
                    clsCommon.AddColumnsForChange(coll, "Distributor_Commission_PKID", obj.Distributor_Commission_PKID, True)
                    clsCommon.AddColumnsForChange(coll, "Distributor_Commission_Rate", obj.Distributor_Commission_Rate, True)
                    clsCommon.AddColumnsForChange(coll, "Distributor_Commission_RateWithTax", obj.Distributor_Commission_RateWithTax, True)
                    clsCommon.AddColumnsForChange(coll, "Distributor_Commission_Amt", obj.Distributor_Commission_Amt, True)
                    clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Rate", obj.Transporter_Commission_Rate, True)
                    clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Amt", obj.Transporter_Commission_Amt, True)
                    clsCommon.AddColumnsForChange(coll, "Security_Rate", obj.Security_Rate, True)
                    clsCommon.AddColumnsForChange(coll, "Security_Amt", obj.Security_Amt, True)
                    clsCommon.AddColumnsForChange(coll, "Transporter", obj.Transporter, True)
                    clsCommon.AddColumnsForChange(coll, "Price_with_tax", obj.Price_with_tax, True)
                    clsCommon.AddColumnsForChange(coll, "Against_Dispatch", obj.Against_Dispatch, True)

                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_RETURN_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                    If clsCommon.myLen(obj.Against_Dispatch) = 0 Then
                        clsSerializeInvenotry.SaveData("Sale Return", strDocNo, dtDocDate, "I", obj.Item_Code, obj.Location, obj.Line_No, obj.arrSrItem, trans)
                        clsBatchInventory.SaveData(Trans_type_Str, strDocNo, dtDocDate, "I", obj.Item_Code, obj.Location, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                    End If


                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return True
    End Function

    Public Shared Function GetBalanceSRNQty(ByVal strSRNCode As String, ByVal strICode As String, ByVal strCurrPINNo As String, ByVal strUOM As String, ByVal dblMRP As Double, ByVal dblAssessable As Double) As Double
        Dim qry As String = "select SUM(qty * RI) as Balance from(  " & _
            " select TSPL_SD_SALE_RETURN_DETAIL.Item_Code as ICode,TSPL_SD_SALE_RETURN_DETAIL.Qty as Qty,1 as RI from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code=TSPL_SD_SALE_RETURN_DETAIL.Document_Code where TSPL_SD_SALE_RETURN_DETAIL.Status=0 and TSPL_SD_SALE_RETURN_HEAD.Status=1 and TSPL_SD_SALE_RETURN_DETAIL.Document_Code ='" + strSRNCode + "' and TSPL_SD_SALE_RETURN_DETAIL.Item_Code='" + strICode + "' and  TSPL_SD_SALE_RETURN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_SD_SALE_RETURN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_SD_SALE_RETURN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' " & _
            " union all " & _
            " select TSPL_PI_DETAIL.Item_Code as ICode,TSPL_PI_DETAIL.PI_Qty as Qty,-1 as RI from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No where TSPL_PI_DETAIL.SRN_Id='" + strSRNCode + "'   and TSPL_PI_DETAIL.Item_Code='" + strICode + "'  and  TSPL_PI_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_PI_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_PI_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'  and TSPL_PI_DETAIL.PI_No not in ('" + strCurrPINNo + "')  " & _
            " )Final "
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
    End Function

    Public Shared Function CompleteSRN(ByVal strDoccode As String, ByVal strICode As String, ByVal LineNo As Integer) As Boolean
        Dim qry As String = "update TSPL_SD_SALE_RETURN_DETAIL set Status=1 where Document_Code='" + strDoccode + "' and Line_No='" + clsCommon.myCstr(LineNo) + "' and Item_Code='" + strICode + "'"
        Return clsDBFuncationality.ExecuteNonQuery(qry)
    End Function
End Class

Public Class clsDSSalesReturnBookingDetail
#Region "Variables"
    Public Document_Code As String = ""
    Public Booth_Code As String = ""
    Public Booth_Name As String = ""
    Public Item_Code As String = ""
    Public Item_Name As String = ""
    Public Unit_Code As String = ""
    Public Return_Qty As Double = 0

#End Region
    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsDSSalesReturnBookingDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
                For Each obj As clsDSSalesReturnBookingDetail In Arr
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Booth_Code", obj.Booth_Code)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Unit_Code", obj.Unit_Code)
                    clsCommon.AddColumnsForChange(coll, "Return_Qty", obj.Return_Qty)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_RETURN_BOOKING_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    'Public Shared Function GetData(ByVal strSRNCode As String, ByVal trans As SqlTransaction) As clsDSSalesReturnBookingDetail
    '    Dim obj As List(Of clsDSSalesReturnBookingDetail)
    '    Dim qry As String = "select SUM(qty * RI) as Balance from(  " &
    '        " select TSPL_SD_SALE_RETURN_DETAIL.Item_Code as ICode,TSPL_SD_SALE_RETURN_DETAIL.Qty as Qty,1 as RI from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code=TSPL_SD_SALE_RETURN_DETAIL.Document_Code where TSPL_SD_SALE_RETURN_DETAIL.Status=0 and TSPL_SD_SALE_RETURN_HEAD.Status=1 and TSPL_SD_SALE_RETURN_DETAIL.Document_Code ='" + strSRNCode + "' and TSPL_SD_SALE_RETURN_DETAIL.Item_Code='" + strICode + "' and  TSPL_SD_SALE_RETURN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_SD_SALE_RETURN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_SD_SALE_RETURN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' " &
    '        " union all " &
    '        " select TSPL_PI_DETAIL.Item_Code as ICode,TSPL_PI_DETAIL.PI_Qty as Qty,-1 as RI from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No where TSPL_PI_DETAIL.SRN_Id='" + strSRNCode + "'   and TSPL_PI_DETAIL.Item_Code='" + strICode + "'  and  TSPL_PI_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_PI_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_PI_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'  and TSPL_PI_DETAIL.PI_No not in ('" + strCurrPINNo + "')  " &
    '        " )Final "
    '    Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
    'End Function
End Class
