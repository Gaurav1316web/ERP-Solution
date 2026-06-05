Imports System.Data.SqlClient
Imports System.IO
'============================Update by preeti gupta====against ticket no[BM00000008159]
Public Class clsMCCMaterialSale
#Region "Variables"
    Public EWayBillDate As Date?
    Public EWayBillNo As String = Nothing
    Public Is_CashSale As String = "N"
    Public Total_Comm_Amt As Double = 0
    Public WayBillDate As Date?
    Public WayBillNo As String = Nothing
    Public SO_Validity As Integer = 0
    Public Commission_Apply As Integer
    Public Dispatch_date As Date?
    Public Dispatch_Terms As String = Nothing
    Public Payment_Terms As String = Nothing
    Public Dispatch_Period As Integer = 0
    Public Vehicle_Capacity As Integer = 0
    Public Road_Permit_No As String = Nothing
    Public Direct_Dispatch As Integer = 0
    Public Is_Delivered As Integer = 0
    Public Podate As DateTime? = Nothing
    Public Form_38_No As String = Nothing
    Public Cust_PO_No As String = Nothing
    Public Price_Group_Code As String = Nothing
    Public Invoice_No As String = Nothing
    Public Invoice_Type As String = Nothing
    Public Document_Code As String = Nothing
    Public isTaxExempted As Boolean = False '' Not a table field
    Public Document_Date As DateTime? = Nothing
    Public Customer_Code As String = Nothing
    Public Against_CF_Sale_Purchase_No As String = Nothing
    Public Customer_Name As String = Nothing  'Not a table field
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Rate_Status As Integer = 1
    Public On_Hold As Boolean = Nothing
    Public Ref_No As String = Nothing
    Public Description As String = Nothing
    Public Remarks As String = Nothing
    Public Tax_Group As String = Nothing
    Public TaxGroupName As String = Nothing 'Not a table field
    Public Bill_To_Location As String = Nothing
    Public Sub_Location_code As String = Nothing
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
    Public RoundOffAmount As Double = 0

    Public Comments As String = Nothing
    Public Comp_Code As String = Nothing
    Public Terms_Code As String = Nothing
    Public TermsName As String = Nothing
    Public Due_Date As DateTime? = Nothing
    Public Posting_Date As DateTime? = Nothing
    Public Carrier As String = Nothing
    Public VehicleNo As String = Nothing
    Public Vehicle_Code As String = Nothing
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
    Public Against_Sales_Order As String = Nothing
    Public Is_Internal As Boolean = False
    Public Tax_Calculation_Type As EnumTaxCalucationType

    Public Is_Create_Auto_Invoice As Boolean = False
    Public Sale_Invoice_No As String = Nothing
    Public Is_Create_Auto_Receipt As Boolean = False
    Public Is_Apply_TPT As Boolean = False
    Public Is_Add_TPT As Boolean = False
    Public TPT_Vendor As String = Nothing
    Public Recommended_By As String = Nothing
    Public Against_POS As String = Nothing
    Public DCS_Price_Code As String = Nothing
    Public Sale_Route_Code As String = Nothing
    Public Supply_Date As DateTime? = Nothing
    Public Shift_Type As String = Nothing
    Public Salesman_Code As String = Nothing
    Public Salesman_Name As String = Nothing
    Public Arr As List(Of clsMCCMaterialSaleDetail) = Nothing

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public PROJECT_ID As String = Nothing
    Public IS_TCS As Integer = 0
    Public Price_Code As String = Nothing
    Public Route_No As String = Nothing
    Public Route_Desc As String = Nothing
    Public HeadDisc_Per As Double = 0
    Public HeadDisc_Amt As Double = 0
    Public RateDiff_Per As Double = 0
    Public RateDiff_Amt As Double = 0
    Public Rate_Diff_Amount_Type As Integer
    Public Is_Rate_Diff_Per_Amt As Integer
    Public Gross_Amount As Double = 0
    Public TotCashDiscAmt As Double = 0
    Public HeadDisc_PerAmt As Double = 0
    Public Mannual_Invoice_No As Integer = 0
    Public InvoiceManualNowithPrefix As String = Nothing
    Public Is_Taxable As Boolean = False
    Public Electronic_Ref_No As String = Nothing
    Public No_Of_Instalment As Integer
    Public Exclude_KKF_And_Mandi As Boolean = False
    Public Receipt_No As String = ""
    Public ReceiptAmt As Double = 0
    Public ReceiverName As String = ""
    Public TotalSubsidyAmt As Double = 0
    Public TotalSubsidyDisAmt As Double = 0
    Public Transporter_Commission_TotalAmt As Double = 0
    Public Deduction_Type As String = Nothing
    Public Deduction As String = Nothing
    Public Bank_Code As String = Nothing
    Public IsEwayBill As Integer = 0
    Public No_Transporter As Integer = 0
    Public Transport_Id As String = ""
    Public Transporter_Name As String = ""

#End Region

    Public Function SaveData(ByVal obj As clsMCCMaterialSale, ByVal isNewEntry As Boolean) As Boolean
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
    Public Function GetInvoiceType(ByVal strCustCode As String, ByVal strLocation As String, ByVal strType As String, ByVal trans As SqlTransaction)
        Dim dt As DataTable


        Dim qry As String

        If clsCommon.CompairString(strType, "S") = CompairStringResult.Equal Then

            qry = "select State from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" & strLocation & "'"
        Else

            qry = "SELECT TSPL_LOCATION_MASTER.Excisable,TSPL_LOCATION_MASTER.State, " &
              "TSPL_LOCATION_MASTER.Sales_Tax_Group as LocalTaxGroup,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as Local_Tax_GroupName, " &
              "TSPL_LOCATION_MASTER.Sales_Tax_GroupIS as InterstateTaxGroup,TSPL_TAX_GROUP_MASTERIS.Tax_Group_Desc as Interstate_Tax_GroupName " &
              "FROM TSPL_LOCATION_MASTER left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' left outer join TSPL_TAX_GROUP_MASTER as TSPL_TAX_GROUP_MASTERIS on TSPL_TAX_GROUP_MASTERIS.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_GroupIS and TSPL_TAX_GROUP_MASTERIS.Tax_Group_Type='S' " &
              "WHERE TSPL_LOCATION_MASTER.Location_Code = '" + strLocation + "'"
        End If
        dt = clsDBFuncationality.GetDataTable(qry, trans)
        Dim strLocState As String = clsCommon.myCstr(dt.Rows(0)("State"))

        qry = "select Price_Code,price_CodeNon,State,Tin_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + strCustCode + "'"
        dt = clsDBFuncationality.GetDataTable(qry, trans)
        Dim strCustState As String = clsCommon.myCstr(dt.Rows(0)("State"))
        Dim strTinNo As String = clsCommon.myCstr(dt.Rows(0)("Tin_No"))
        If clsCommon.myLen(strTinNo) > 0 AndAlso clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
            Invoice_Type = "T"
        Else
            Invoice_Type = "R"
        End If
        Return Invoice_Type
    End Function

    Private Function AllowToSave(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = "select TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per, TSPL_VLC_MASTER_HEAD.MCC,TSPL_PAYMENT_CYCLE_MASTER.PC_CODE,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE" + Environment.NewLine +
                    "from TSPL_VENDOR_MASTER" + Environment.NewLine +
                    "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_VENDOR_MASTER.Vendor_Code" + Environment.NewLine +
                    "left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code" + Environment.NewLine +
                    "left outer join TSPL_MCC_MASTER on  TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC" + Environment.NewLine +
                    "left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle" + Environment.NewLine +
                    "where TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" + obj.Customer_Code + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.myCdbl(dt.Rows(0)("Credit_Limit_On_Milk_Receipt_Per") >= 0) Then
                    Dim dtDoc As Date = obj.Document_Date
                    Dim PaymentCycleType As String = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
                    Dim PaymentCycleValue As Integer = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
                    Dim dtFrom As Date = dtDoc
                    Dim dtTo As Date = dtDoc
                    If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
                        Dim ModResut As Integer = dtDoc.Day Mod PaymentCycleValue
                        If ModResut = 0 Then
                            ModResut = PaymentCycleValue
                        End If
                        dtFrom = New Date(dtDoc.Year, dtDoc.Month, (dtDoc.Day - (ModResut) + 1))
                        dtTo = dtFrom.AddDays(PaymentCycleValue - 1)
                        Dim dtNxtPay As DateTime = dtTo.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
                        If dtFrom.Month <> dtNxtPay.Month Then
                            dtTo = New Date(dtFrom.Year, dtFrom.Month, 1).AddMonths(1).AddDays(-1)
                        End If
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(clsCommon.GetPrintDate(dtFrom, "dd")) <> 1 Then
                            dtFrom = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                            dtTo = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                        End If
                        dtTo = DateAdd(DateInterval.Month, PaymentCycleValue, dtFrom)
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(clsCommon.GetPrintDate(dtFrom, "dd")) <> 1 Then
                            dtFrom = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                            dtTo = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                        End If
                        dtTo = DateAdd(DateInterval.Year, PaymentCycleValue, dtFrom)
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
                        Dim today As Date = dtFrom
                        Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
                        dtFrom = today.AddDays(-dayDiff)
                        dtTo = dtFrom.AddDays(6)
                    End If


                    qry = "select convert(decimal(18,2), isnull( sum(Amount*RI),0)) as Amount from (" + Environment.NewLine +
                        "select TSPL_MILK_SRN_HEAD.DOC_CODE,(AMOUNT*TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per/100) as Amount,1 as RI,1 as Chk from TSPL_MILK_SRN_DETAIL" + Environment.NewLine +
                        "left outer join TSPL_MILK_SRN_HEAD	on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE" + Environment.NewLine +
                        "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_MILK_SRN_HEAD.VSP_CODE" + Environment.NewLine +
                        "left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_MILK_SRN_HEAD.VSP_CODE" + Environment.NewLine +
                        "where TSPL_MILK_SRN_HEAD.Posted=1 and TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" + obj.Customer_Code + "' and TSPL_MILK_SRN_HEAD.DOC_DATE>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_MILK_SRN_HEAD.DOC_DATE<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per>=0" + Environment.NewLine +
                        "union all " + Environment.NewLine +
                        "select TSPL_SD_SHIPMENT_HEAD.Document_Code as DOC_CODE,TSPL_SD_SHIPMENT_HEAD.Total_Amt as Amount,-1 as RI,0 as chk  from TSPL_SD_SHIPMENT_HEAD where Trans_Type='MCC' and   TSPL_SD_SHIPMENT_HEAD.Customer_Code='" + obj.Customer_Code + "' " + Environment.NewLine +
                        "and TSPL_SD_SHIPMENT_HEAD.Document_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_SD_SHIPMENT_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_SD_SHIPMENT_HEAD.Document_Code not in ('" + obj.Document_Code + "')" + Environment.NewLine +
                        "union all" + Environment.NewLine +
                        "select TSPL_SD_SALE_RETURN_HEAD.Document_Code,TSPL_SD_SALE_RETURN_HEAD.Total_Amt,1 as RI,0 as chk from TSPL_SD_SALE_RETURN_HEAD where Trans_Type='MCC' and TSPL_SD_SALE_RETURN_HEAD.Customer_Code='" + obj.Customer_Code + "' and TSPL_SD_SALE_RETURN_HEAD.Status=1" + Environment.NewLine +
                        "and not exists(select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Return_Doc_No=TSPL_SD_SALE_RETURN_HEAD.Document_Code)" + Environment.NewLine +
                        "Union all" + Environment.NewLine +
                        "select TSPL_MILK_REJECT_HEAD.DOC_CODE,cast (((case when TSPL_MILK_REJECT_DETAIL.Is_Return = 0 and TSPL_MILK_REJECT_DETAIL.Defaulter='VSP' then TSPL_MILK_REJECT_DETAIL.SNF_Deduction_Amount+(case when TSPL_MILK_REJECT_DETAIL.Reject_Type='Curd' then TSPL_MILK_REJECT_DETAIL.FAT_Deduction_Amount else 0 end) else (case when TSPL_MILK_REJECT_DETAIL.Is_Return = 1 then TSPL_MILK_REJECT_DETAIL.Amount else (case when TSPL_MILK_REJECT_DETAIL.Is_Return = 3 then TSPL_MILK_REJECT_DETAIL.Amount else 0 end ) end ) end )*TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per/100) as decimal(18,2)) as Amount ,-1 as RI,0 as chk" + Environment.NewLine +
                        "from TSPL_MILK_REJECT_DETAIL" + Environment.NewLine +
                        "left outer join TSPL_MILK_REJECT_HEAD on TSPL_MILK_REJECT_HEAD.DOC_CODE=TSPL_MILK_REJECT_DETAIL.DOC_CODE" + Environment.NewLine +
                        "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_MILK_REJECT_DETAIL.VSP_CODE" + Environment.NewLine +
                        "left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_MILK_REJECT_DETAIL.VSP_CODE" + Environment.NewLine +
                        "where TSPL_MILK_REJECT_HEAD.Posted=1 and TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" + obj.Customer_Code + "' and TSPL_MILK_REJECT_HEAD.DOC_DATE>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_MILK_REJECT_HEAD.DOC_DATE<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per>=0" + Environment.NewLine +
                         ")xx  "
                    Dim dblBal As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                    If obj.Total_Amt > dblBal Then
                        Throw New Exception("Balance Amount of Milk Received : " + clsCommon.myCstr(dblBal) + "And Document Amount :" + clsCommon.myCstr(obj.Total_Amt))
                    End If
                End If
            End If

            qry = "select 1 from TSPL_ITEM_MASTER where TSPL_ITEM_MASTER.Deduction <>'" + obj.Deduction + "' and Item_Code in ("
            For Each objtr As clsMCCMaterialSaleDetail In obj.Arr
                If clsCommon.CompairString(objtr.Row_Type, "Item") = CompairStringResult.Equal Then
                    qry += "'" + objtr.Item_Code + "'"
                End If
            Next
            qry += ")"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("Deduction of All item should be [" + obj.Deduction + "]")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return True
    End Function

    Public Function SaveData(ByVal obj As clsMCCMaterialSale, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try
            '' Anubhooti 06-Sep-2014 BM00000003735 (Locked Transaction)
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Sales And Distribution", "Shipment/Sale Invoice", obj.Bill_To_Location, obj.Document_Date, trans)
            ''shivani
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmMCCMaterial, obj.Bill_To_Location, obj.Document_Date, trans)
            ''
            ''BHA/09/05/18-000015 By balwinder on 10/05/2018
            AllowToSave(obj, trans)
            clsSerializeInvenotry.DeleteData("SD-In", obj.Document_Code, trans)
            'If Not isNewEntry Then
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", "TSPL_SD_SHIPMENT_DETAIL", "Document_Code", trans)
            'End If
            Dim qry As String = "delete from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + obj.Document_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsBatchInventory.DeleteData("MCC-MSALE", obj.Document_Code, trans)
            Dim strDocNo As String = ""
            If isNewEntry Then
                If clsERPFuncationality.GetGSTStatus(obj.Document_Date) Then
                    If obj.isTaxExempted Then
                        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DCSDispatch, clsDocTransactionType.TaxExempted_ProductInvoice, obj.Bill_To_Location)
                    Else
                        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DCSDispatch, clsDocTransactionType.Other, obj.Bill_To_Location)
                    End If
                    'If obj.Is_Taxable Then
                    '    Dim strTaxType = clsLocationWiseTax.TaxType(obj.Bill_To_Location, obj.Customer_Code, "S", obj.Document_Date, trans)
                    '    If clsCommon.CompairString(strTaxType, "L") = CompairStringResult.Equal Then
                    '        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTLocal, obj.Bill_To_Location)
                    '    Else
                    '        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTInterstate, obj.Bill_To_Location)
                    '    End If
                    'Else
                    '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmShipmentProductSale, clsDocTransactionType.GSTNonTaxable, obj.Bill_To_Location)
                    'End If
                Else
                    If obj.isTaxExempted Then
                        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DCSDispatch, clsDocTransactionType.TaxExempted_ProductInvoice, obj.Bill_To_Location)
                    Else
                        obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DCSDispatch, clsDocTransactionType.Other, obj.Bill_To_Location)
                    End If
                End If
                'If clsCommon.CompairString(obj.Item_Type, "F") = CompairStringResult.Equal Then
                '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNShipment, clsDocTransactionType.SNQuotationFinishedGoods, obj.Bill_To_Location)
                'Else
                '    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNShipment, clsDocTransactionType.SNQuotationOther, obj.Bill_To_Location)
                'End If
            End If
            '''' for Invoice Type
            Dim AllowChangeInvoiceType As Boolean = False
            AllowChangeInvoiceType = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Allow_Change_InvoiceType from TSPL_inv_parameters", trans)) = 0, False, True)
            If obj.Is_Create_Auto_Invoice Then
                If AllowChangeInvoiceType Then
                    If clsCommon.myLen(obj.Invoice_Type) <= 0 Then
                        common.clsCommon.MyMessageBoxShow("Please select invoice  Type for creating auto invoice")
                        Return False
                    End If
                Else
                    If clsCommon.myLen(obj.Ship_To_Location) > 0 Then
                        obj.Invoice_Type = GetInvoiceType(obj.Customer_Code, obj.Bill_To_Location, "B", trans)
                    Else
                        obj.Invoice_Type = GetInvoiceType(obj.Customer_Code, obj.Bill_To_Location, "B", trans)
                    End If
                End If
            End If
            ''''' invoice type ends here
            If clsERPFuncationality.GetGSTStatus(obj.Document_Date) Then
                If obj.Is_Taxable Then
                    obj.Invoice_Type = "T"
                Else
                    obj.Invoice_Type = "N"
                End If
            End If



            If (clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            ''richa 13/08/2014 against Ticket No BM00000003502
            'If IsDBNull("cust_po_date") Then
            If clsCommon.myLen(obj.Podate) <= 0 Then
                clsCommon.AddColumnsForChange(coll, "cust_po_date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            Else
                clsCommon.AddColumnsForChange(coll, "cust_po_date", clsCommon.GetPrintDate(obj.Podate, "dd/MMM/yyyy hh:mm tt"))
            End If
            '--------------------------------------------------------
            '-----------------richa 27/06/2014 Ticket No .BM00000002982--------------------------------
            Dim isIncrementCounter As Boolean = True
            If obj.Mannual_Invoice_No > 0 OrElse clsCommon.myLen(obj.InvoiceManualNowithPrefix) > 0 Then
                isIncrementCounter = False
            End If
            Dim Doc_Code As String = Nothing

            ''richa agarwal 18/03/2015 sale invoice series generation setting based
            Dim Desc As String = String.Empty
            If obj.Mannual_Invoice_No > 0 Then
                If clsCommon.CompairString(obj.Invoice_Type, "T") = CompairStringResult.Equal Then
                    ''richa agarwal 18/03/2015
                    Desc = clsFixedParameter.GetData(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesTaxatMCCSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesTaxatMCCSale, trans)
                    If clsCommon.CompairString(Desc, "1") = CompairStringResult.Equal Then
                        Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceTax, obj.Bill_To_Location, False, isIncrementCounter)
                    Else
                        Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceTax, obj.Bill_To_Location, False, isIncrementCounter)
                    End If
                    ' Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceTax, obj.Bill_To_Location, False, isIncrementCounter)
                ElseIf clsCommon.CompairString(obj.Invoice_Type, "R") = CompairStringResult.Equal Then
                    ''richa agarwal 18/03/2015
                    Desc = clsFixedParameter.GetData(clsFixedParameterType.AllowToGenerateSaleInvoiceSeriesRetailatMCCSale, clsFixedParameterCode.AllowToGenerateSaleInvoiceSeriesRetailatMCCSale, trans)
                    If clsCommon.CompairString(Desc, "1") = CompairStringResult.Equal Then
                        Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceRetail, obj.Bill_To_Location, False, isIncrementCounter)
                    Else
                        Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceRetail, obj.Bill_To_Location, False, isIncrementCounter)
                    End If
                    '  Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceRetail, obj.Bill_To_Location, False, isIncrementCounter)
                ElseIf clsCommon.CompairString(obj.Invoice_Type, "S") = CompairStringResult.Equal Then
                    Doc_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceService, obj.Bill_To_Location, False, isIncrementCounter)
                End If

                Dim strDocLike As String = ""
                Dim arr As Array = Doc_Code.ToCharArray()
                For jj As Integer = 0 To arr.Length - clsCommon.myLen(obj.Mannual_Invoice_No) - 1
                    strDocLike += clsCommon.myCstr(arr(jj))
                Next
                strDocLike += clsCommon.myCstr(obj.Mannual_Invoice_No)
                Doc_Code = strDocLike
                obj.Mannual_Invoice_No = Doc_Code

            ElseIf clsCommon.myLen(obj.InvoiceManualNowithPrefix) > 0 Then
                Doc_Code = obj.InvoiceManualNowithPrefix
                obj.InvoiceManualNowithPrefix = Doc_Code
            End If

            clsCommon.AddColumnsForChange(coll, "Deduction_Type", obj.Deduction_Type, True)
            clsCommon.AddColumnsForChange(coll, "Deduction", obj.Deduction, True)
            clsCommon.AddColumnsForChange(coll, "Bank_Code", obj.Bank_Code, True)
            '---------------------------------------------------------------------------------------
            clsCommon.AddColumnsForChange(coll, "Customer_Code", obj.Customer_Code)
            clsCommon.AddColumnsForChange(coll, "Against_CF_Sale_Purchase_No", obj.Against_CF_Sale_Purchase_No, True)
            clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Internal", IIf(obj.Is_Internal, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Ref_No", obj.Ref_No)
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            clsCommon.AddColumnsForChange(coll, "Inv_No", obj.Inv_No)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Is_CashSale", obj.Is_CashSale)
            clsCommon.AddColumnsForChange(coll, "Bill_To_Location", obj.Bill_To_Location)
            clsCommon.AddColumnsForChange(coll, "Sub_Location_code", obj.Sub_Location_code)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Location", obj.Ship_To_Location)
            clsCommon.AddColumnsForChange(coll, "Transport_Id", obj.Transport_Id, True)
            clsCommon.AddColumnsForChange(coll, "Transporter_Name", obj.Transporter_Name, True)
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
            clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)
            clsCommon.AddColumnsForChange(coll, "PROJECT_ID", obj.PROJECT_ID, True)

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
            clsCommon.AddColumnsForChange(coll, "VehicleNo", obj.VehicleNo)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Code", obj.Vehicle_Code)
            clsCommon.AddColumnsForChange(coll, "GRNo", obj.GRNo)
            clsCommon.AddColumnsForChange(coll, "GENo", obj.GENo)
            clsCommon.AddColumnsForChange(coll, "Dept", obj.Dept)
            clsCommon.AddColumnsForChange(coll, "Dept_Desc", obj.Dept_Desc)
            clsCommon.AddColumnsForChange(coll, "Item_Type", obj.Item_Type)
            clsCommon.AddColumnsForChange(coll, "Against_Sales_Order", obj.Against_Sales_Order, True)
            clsCommon.AddColumnsForChange(coll, "Against_POS", obj.Against_POS, True)
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

            If obj.Challan_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Challan_Date", clsCommon.GetPrintDate(obj.Challan_Date, "dd/MMM/yyyy"))
            End If
            If obj.Inv_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Inv_Date", clsCommon.GetPrintDate(obj.Inv_Date, "dd/MMM/yyyy"))
            End If

            clsCommon.AddColumnsForChange(coll, "ParentDocNo", obj.Document_Code)
            clsCommon.AddColumnsForChange(coll, "Is_Create_Auto_Invoice", IIf(obj.Is_Create_Auto_Invoice, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Sale_Invoice_No", obj.Sale_Invoice_No)
            clsCommon.AddColumnsForChange(coll, "Is_Create_Auto_Receipt", IIf(obj.Is_Create_Auto_Receipt, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Apply_TPT", IIf(obj.Is_Apply_TPT, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Add_TPT", IIf(obj.Is_Add_TPT, 1, 0))
            clsCommon.AddColumnsForChange(coll, "TPT_Vendor", obj.TPT_Vendor, True)
            clsCommon.AddColumnsForChange(coll, "DCS_Price_Code", obj.DCS_Price_Code, True)
            ' clsCommon.AddColumnsForChange(coll, "Route_No", obj.Route_No, True)
            If obj.Supply_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Supply_Date", clsCommon.GetPrintDate(obj.Supply_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Supply_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
            End If
            clsCommon.AddColumnsForChange(coll, "Shift_type", obj.Shift_Type, True)
            clsCommon.AddColumnsForChange(coll, "Exclude_KKF_And_Mandi", IIf(obj.Exclude_KKF_And_Mandi, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Recommended_By", obj.Recommended_By, True)
            'clsCommon.AddColumnsForChange(coll, "Is_Create_Auto_Receipt", 1)
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
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Per", obj.HeadDisc_Per)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_PerAmt", obj.HeadDisc_PerAmt)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Amt", obj.HeadDisc_Amt)
            clsCommon.AddColumnsForChange(coll, "RateDiff_Per", obj.RateDiff_Per)
            clsCommon.AddColumnsForChange(coll, "Rate_Diff_Amount_Type", obj.Rate_Diff_Amount_Type, True)
            clsCommon.AddColumnsForChange(coll, "Is_Rate_Diff_Per_Amt", obj.Is_Rate_Diff_Per_Amt, True)
            clsCommon.AddColumnsForChange(coll, "RateDiff_Amt", obj.RateDiff_Amt)
            clsCommon.AddColumnsForChange(coll, "Gross_Amount", obj.Gross_Amount)
            clsCommon.AddColumnsForChange(coll, "TotCashDiscAmt", obj.TotCashDiscAmt)
            clsCommon.AddColumnsForChange(coll, "Invoice_Type", obj.Invoice_Type)
            clsCommon.AddColumnsForChange(coll, "Price_Group_Code", obj.Price_Group_Code)
            clsCommon.AddColumnsForChange(coll, "Cust_PO_No", obj.Cust_PO_No)
            clsCommon.AddColumnsForChange(coll, "Form_38_No", obj.Form_38_No)
            '-----------------richa 27/06/2014 Ticket No .BM00000002982--------------------------------
            clsCommon.AddColumnsForChange(coll, "Mannual_Invoice_No", obj.Mannual_Invoice_No)
            clsCommon.AddColumnsForChange(coll, "Mannual_Invoice_No_StringType", obj.InvoiceManualNowithPrefix)


            clsCommon.AddColumnsForChange(coll, "SO_Validity", obj.SO_Validity)
            clsCommon.AddColumnsForChange(coll, "Commission_Apply", obj.Commission_Apply)
            'clsCommon.AddColumnsForChange(coll, "Dispatch_date", clsCommon.GetPrintDate(obj.Dispatch_date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Dispatch_date", clsCommon.GetPrintDate(obj.Inv_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Vehicle_Capacity", obj.Vehicle_Capacity)
            clsCommon.AddColumnsForChange(coll, "Dispatch_Terms", obj.Dispatch_Terms)
            clsCommon.AddColumnsForChange(coll, "IS_TCS", obj.IS_TCS)
            clsCommon.AddColumnsForChange(coll, "Payment_Terms", obj.Payment_Terms)
            clsCommon.AddColumnsForChange(coll, "Dispatch_Period", obj.Dispatch_Period)
            clsCommon.AddColumnsForChange(coll, "Road_Permit_No", obj.Road_Permit_No)
            clsCommon.AddColumnsForChange(coll, "WayBillNo", obj.WayBillNo)
            clsCommon.AddColumnsForChange(coll, "Total_Comm_Amt", obj.Total_Comm_Amt)
            clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
            clsCommon.AddColumnsForChange(coll, "Receipt_No", obj.Receipt_No, True)
            clsCommon.AddColumnsForChange(coll, "ReceiptAmt", obj.ReceiptAmt, True)
            clsCommon.AddColumnsForChange(coll, "ReceiverName", obj.ReceiverName, True)
            clsCommon.AddColumnsForChange(coll, "TotalSubsidyAmt", obj.TotalSubsidyAmt, True)
            clsCommon.AddColumnsForChange(coll, "TotalSubsidyDisAmt", obj.TotalSubsidyDisAmt, True)
            clsCommon.AddColumnsForChange(coll, "Transporter_Commission_TotalAmt", obj.Transporter_Commission_TotalAmt)
            clsCommon.AddColumnsForChange(coll, "No_Transporter", obj.No_Transporter)
            If clsCommon.myLen(obj.Against_Sales_Order) = 0 Then
                obj.Direct_Dispatch = 1
            End If
            clsCommon.AddColumnsForChange(coll, "Direct_Dispatch", obj.Direct_Dispatch)

            clsCommon.AddColumnsForChange(coll, "No_Of_Instalment", obj.No_Of_Instalment)
            'clsCommon.AddColumnsForChange(coll, "WayBillDate", clsCommon.GetPrintDate(obj.WayBillDate, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Is_Taxable", IIf(obj.Is_Taxable, 1, 0))
            clsCommon.AddColumnsForChange(coll, "IsEwayBill", IIf(obj.IsEwayBill, 1, 0))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Trans_Type", "MCC")
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SHIPMENT_HEAD", OMInsertOrUpdate.Insert, "", trans)

                'If obj.Rate_Status = 2 Then
                '    qry = "insert into TSPL_TRANSACTION_APPROVAL(Screen_Name,Program_Code,Document_No,Doc_Date,approval_type,Approve,Created_By,Created_Date,Modified_By,Modified_Date,Comp_Code) " &
                '    "values ('MCC Material Sale','" & clsUserMgtCode.frmMCCMaterial & "','" & obj.Document_Code & "','" & clsCommon.GetPrintDate(obj.Document_Date, "dd-MMM-yyyy hh:mm:ss") & "','Rate',0,'" + objCommonVar.CurrentUserCode + "','" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") + "','" + objCommonVar.CurrentUserCode + "','" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") + "','" & objCommonVar.CurrentCompanyCode & "')"
                '    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'End If
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SHIPMENT_HEAD", OMInsertOrUpdate.Update, "TSPL_SD_SHIPMENT_HEAD.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            isSaved = isSaved AndAlso clsMCCMaterialSaleDetail.SaveData(obj.Document_Code, Arr, trans, obj.Document_Date)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", "TSPL_SD_SHIPMENT_DETAIL", "Document_Code", trans)

            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_Code, obj.arrCustomFields, trans)
            '''' to save item weight unit
            qry = "update TSPL_SD_shipment_DETAIL set Weight_UOM= (select Weight_UOM from TSPL_ITEM_MASTER where Item_Code=TSPL_SD_shipment_DETAIL.Item_Code)  where Document_Code='" + obj.Document_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            isSaved = isSaved AndAlso clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "Document_Code", obj.Document_Code, "TSPL_SD_SHIPMENT_HEAD", trans)

            ''''  for automatic invoice
            Dim objSI As clsPSInvoiceHead = ConvertShipmentToSaleInvoice(obj)
            If clsCommon.myLen(obj.Sale_Invoice_No) > 0 Then
                objSI.SaveData(objSI, False, False, clsDocType.DCSInvoice, trans, True)
            Else
                If clsCommon.CompairString(obj.Is_CashSale, "Y") = CompairStringResult.Equal Then
                    objSI.SaveData(objSI, True, True, clsDocType.DCSInvoice, trans, True)
                Else
                    objSI.SaveData(objSI, True, False, clsDocType.DCSInvoice, trans, True)
                End If

            End If
            Dim sQuery As String = "update TSPL_SD_SALE_INVOICE_HEAD set trans_type='MCC' where Against_Shipment_No='" & obj.Document_Code & "'"
            clsDBFuncationality.ExecuteNonQuery(sQuery, trans)
            'objSI.SaveData(objSI, True, trans)
            'obj.Sale_Invoice_No = objSI.Document_Code
            ''''  automatic invoice ends here
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function
    Public Shared Function UpdateAfterPosting(ByVal obj As clsMCCMaterialSale, ByVal DocumentCode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If obj IsNot Nothing And clsCommon.myLen(DocumentCode) > 0 Then
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "EWayBillNo", obj.EWayBillNo)
                clsCommon.AddColumnsForChange(coll, "Electronic_Ref_No", obj.Electronic_Ref_No)

                If clsCommon.myLen(obj.EWayBillDate) > 0 Then
                    clsCommon.AddColumnsForChange(coll, "EWayBillDate", clsCommon.GetPrintDate(obj.EWayBillDate, "dd/MMM/yyyy"))
                Else
                    clsCommon.AddColumnsForChange(coll, "EWayBillDate", Nothing, True)
                End If
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SHIPMENT_HEAD", OMInsertOrUpdate.Update, "TSPL_SD_SHIPMENT_HEAD.Document_Code='" + DocumentCode + "'", trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function checkSaveNotification(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim Count As Integer = 0
            Dim CreditLimit As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Credit_Limit from TSPL_CUSTOMER_MASTER WHERE Cust_Code='" + obj.Customer_Code + "'", trans))
            Dim qry As String
            Dim dt As DataTable = clsScreenNotificationSchedule.GetScreenNotificationInfo(clsUserMgtCode.frmSNShipment, trans)
            For Each dr As DataRow In dt.Rows
                'Criteria, Notification, Validation
                If clsCommon.CompairString(dr("Criteria"), "Credit days") = CompairStringResult.Equal Then
                    qry = "Select COUNT(*) from TSPL_SD_SHIPMENT_HEAD" &
        " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code" &
        " LEFT OUTER JOIN TSPL_Customer_Invoice_Head ON TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code" &
        " WHERE TSPL_SD_SHIPMENT_HEAD.Status = 1" &
        " AND TSPL_SD_SHIPMENT_HEAD.Customer_Code='" + obj.Customer_Code + "'" &
        " AND TSPL_SD_SHIPMENT_HEAD.Due_Date<'" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "'" &
        " AND ISNULL(TSPL_Customer_Invoice_Head.Balance_Amt,0)<>0"
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) > 0 Then
                        If clsCommon.CompairString(dr("Validation"), "Required Approval") = CompairStringResult.Equal Then
                            'clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Notification")))
                            If common.clsCommon.MyMessageBoxShow(clsCommon.myCstr(dr("Notification")) + Environment.NewLine + "Do you want to continue?.", "Load Out", MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
                                Dim frm As New FrmPWD(trans)
                                frm.strCode = clsFixedParameterCode.CreditLimitApproval
                                frm.strType = clsFixedParameterType.CreditLimitApproval
                                frm.ShowDialog()
                                If frm.isPasswordCorrect Then
                                    Count += 1
                                End If
                            Else
                                Return False
                            End If
                        Else
                            Throw New Exception(clsCommon.myCstr(dt.Rows(0)("Notification")))
                        End If
                    End If
                ElseIf clsCommon.CompairString(dr("Criteria"), "Credit Amount") = CompairStringResult.Equal Then
                    qry = "Select SUM(TSPL_Customer_Invoice_Head.Balance_Amt) from TSPL_SD_SHIPMENT_HEAD" &
        " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code" &
        " LEFT OUTER JOIN TSPL_Customer_Invoice_Head ON TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code" &
        " WHERE TSPL_SD_SHIPMENT_HEAD.Status = 1" &
        " AND TSPL_SD_SHIPMENT_HEAD.Customer_Code='" + obj.Customer_Code + "'" &
        " AND TSPL_SD_SHIPMENT_HEAD.Document_Date<'" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "'" &
        " AND ISNULL(TSPL_Customer_Invoice_Head.Balance_Amt,0)<>0"
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) > CreditLimit Then
                        If clsCommon.CompairString(dr("Validation"), "Required Approval") = CompairStringResult.Equal Then
                            'clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Notification")))
                            If common.clsCommon.MyMessageBoxShow(clsCommon.myCstr(dr("Notification")) + Environment.NewLine + "Do you want to continue?.", "Load Out", MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
                                Dim frm As New FrmPWD(trans)
                                frm.strCode = clsFixedParameterCode.CreditLimitApproval
                                frm.strType = clsFixedParameterType.CreditLimitApproval
                                frm.ShowDialog()
                                If frm.isPasswordCorrect Then
                                    Count += 1
                                End If
                            Else
                                Return False
                            End If
                        Else
                            Throw New Exception(clsCommon.myCstr(dt.Rows(0)("Notification")))
                        End If
                    End If
                End If
            Next
            If Count >= 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function checkApprovalDocument(ByVal Prog_code As String, ByVal document_no As String) As Boolean
        Dim sQuery As String
        sQuery = "select * from TSPL_TRANSACTION_APPROVAL where program_Code='" & Prog_code & "' and Document_No='" & document_no & "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(sQuery)
        If dt.Rows.Count <= 0 Then
            Return True
        Else
            If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Approve")), "0") = CompairStringResult.Equal Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType) As clsMCCMaterialSale
        Return GetData(strDocumentNo, NavType, Nothing)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsMCCMaterialSale
        Dim obj As clsMCCMaterialSale = Nothing
        Dim qry As String = "SELECT TSPL_SD_SHIPMENT_HEAD.No_Of_Instalment,TSPL_SD_SHIPMENT_HEAD.IS_TCS, TSPL_SD_SHIPMENT_HEAD.EWayBillNo,TSPL_SD_SHIPMENT_HEAD.EWayBillDate,TSPL_SD_SHIPMENT_HEAD.Road_Permit_No,TSPL_SD_SHIPMENT_HEAD.Is_Delivered,TSPL_SD_SHIPMENT_HEAD.HeadDisc_PerAmt,TSPL_SD_SHIPMENT_HEAD.RateDiff_Per,TSPL_SD_SHIPMENT_HEAD.Rate_Diff_Amount_Type,TSPL_SD_SHIPMENT_HEAD.Is_Rate_Diff_Per_Amt,TSPL_SD_SHIPMENT_HEAD.Gross_Amount,TSPL_SD_SHIPMENT_HEAD.RateDiff_Amt,TSPL_SD_SHIPMENT_HEAD.cust_po_date,TSPL_SD_SHIPMENT_HEAD.Cust_PO_No,TSPL_SD_SHIPMENT_HEAD.Vehicle_Code,TSPL_SD_SHIPMENT_HEAD.price_group_code,TSPL_SD_SHIPMENT_HEAD.Invoice_Type,TSPL_SD_SHIPMENT_HEAD.HeadDisc_Per,TSPL_SD_SHIPMENT_HEAD.HeadDisc_Amt,TSPL_SD_SHIPMENT_HEAD.TotCashDiscAmt,TSPL_SD_SHIPMENT_HEAD.Route_No,TSPL_SD_SHIPMENT_HEAD.Route_Desc,TSPL_SD_SHIPMENT_HEAD.Price_Code,TSPL_SD_SHIPMENT_HEAD.Document_Code,TSPL_SD_SHIPMENT_HEAD.Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_SD_SHIPMENT_HEAD.Status,TSPL_SD_SHIPMENT_HEAD.On_Hold,TSPL_SD_SHIPMENT_HEAD.Bank_Code,TSPL_SD_SHIPMENT_HEAD.Ref_No,TSPL_SD_SHIPMENT_HEAD.Description,TSPL_SD_SHIPMENT_HEAD.Is_CashSale,TSPL_SD_SHIPMENT_HEAD.Remarks,TSPL_SD_SHIPMENT_HEAD.Tax_Group,TSPL_SD_SHIPMENT_HEAD.Bill_To_Location,TSPL_SD_SHIPMENT_HEAD.Sub_Location_code,TSPL_SD_SHIPMENT_HEAD.Ship_To_Location,TSPL_SD_SHIPMENT_HEAD.TAX1,TSPL_SD_SHIPMENT_HEAD.TAX1_Rate,TSPL_SD_SHIPMENT_HEAD.TAX1_Amt,TSPL_SD_SHIPMENT_HEAD.TAX1_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX2,TSPL_SD_SHIPMENT_HEAD.TAX2_Rate,TSPL_SD_SHIPMENT_HEAD.TAX2_Amt,TSPL_SD_SHIPMENT_HEAD.TAX2_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX3,TSPL_SD_SHIPMENT_HEAD.TAX3_Rate,TSPL_SD_SHIPMENT_HEAD.TAX3_Amt,TSPL_SD_SHIPMENT_HEAD.TAX3_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX4,TSPL_SD_SHIPMENT_HEAD.TAX4_Rate,TSPL_SD_SHIPMENT_HEAD.TAX4_Amt,TSPL_SD_SHIPMENT_HEAD.TAX4_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX5,TSPL_SD_SHIPMENT_HEAD.TAX5_Rate,TSPL_SD_SHIPMENT_HEAD.TAX5_Amt,TSPL_SD_SHIPMENT_HEAD.TAX5_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX6,TSPL_SD_SHIPMENT_HEAD.TAX6_Rate,TSPL_SD_SHIPMENT_HEAD.TAX6_Amt,TSPL_SD_SHIPMENT_HEAD.TAX6_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX7,TSPL_SD_SHIPMENT_HEAD.TAX7_Rate,TSPL_SD_SHIPMENT_HEAD.TAX7_Amt,TSPL_SD_SHIPMENT_HEAD.TAX7_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX8,TSPL_SD_SHIPMENT_HEAD.TAX8_Rate,TSPL_SD_SHIPMENT_HEAD.TAX8_Amt,TSPL_SD_SHIPMENT_HEAD.TAX8_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX9,TSPL_SD_SHIPMENT_HEAD.TAX9_Rate,TSPL_SD_SHIPMENT_HEAD.TAX9_Amt,TSPL_SD_SHIPMENT_HEAD.TAX9_Base_Amt,TSPL_SD_SHIPMENT_HEAD.TAX10,TSPL_SD_SHIPMENT_HEAD.TAX10_Rate,TSPL_SD_SHIPMENT_HEAD.TAX10_Amt,TSPL_SD_SHIPMENT_HEAD.TAX10_Base_Amt,TSPL_SD_SHIPMENT_HEAD.Discount_Base,TSPL_SD_SHIPMENT_HEAD.Discount_Amt,TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount,TSPL_SD_SHIPMENT_HEAD.Total_Tax_Amt,TSPL_SD_SHIPMENT_HEAD.Comments,TSPL_SD_SHIPMENT_HEAD.Comp_Code,TSPL_SD_SHIPMENT_HEAD.Terms_Code,TSPL_SD_SHIPMENT_HEAD.Due_Date ,TSPL_LOCATION_MASTER.Location_Desc as BillToLocationName,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as ShipToLocationName,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as TaxGroupName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_SD_SHIPMENT_HEAD.Posting_Date,TSPL_SD_SHIPMENT_HEAD.Total_Amt,TSPL_SD_SHIPMENT_HEAD.Carrier,TSPL_SD_SHIPMENT_HEAD.VehicleNo,TSPL_SD_SHIPMENT_HEAD.GRNo,TSPL_SD_SHIPMENT_HEAD.GENo,TSPL_SD_SHIPMENT_HEAD.GEDate, TSPL_SD_SHIPMENT_HEAD.Dept,TSPL_SD_SHIPMENT_HEAD.Dept_Desc,TSPL_SD_SHIPMENT_HEAD.Item_Type,TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order ,TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code1,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name1,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt1,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code2,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name2,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt2,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code3,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name3,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt3,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code4,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name4,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt4,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code5,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name5,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt5,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code6,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name6,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt6,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code7,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name7,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt7,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code8,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name8,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt8,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code9 ,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name9,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt9 ,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Code10 ,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name10,TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt10,TSPL_SD_SHIPMENT_HEAD.Total_Add_Charge,TSPL_SD_SHIPMENT_HEAD.Tax_Calculation_Type,TSPL_SD_SHIPMENT_HEAD.Challan_No, TSPL_SD_SHIPMENT_HEAD.Challan_Date, TSPL_SD_SHIPMENT_HEAD.Inv_Date,TSPL_SD_SHIPMENT_HEAD.Inv_No,TSPL_SD_SHIPMENT_HEAD.Is_Internal,TSPL_SD_SHIPMENT_HEAD.Is_Create_Auto_Invoice,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No,TSPL_SD_SHIPMENT_HEAD.Is_Create_Auto_Receipt,TSPL_SD_SHIPMENT_HEAD.Is_Apply_TPT,TSPL_SD_SHIPMENT_HEAD.Is_Add_TPT,TSPL_SD_SHIPMENT_HEAD.Exclude_KKF_And_Mandi,TSPL_SD_SHIPMENT_HEAD.Recommended_By,TSPL_SD_SHIPMENT_HEAD.TPT_Vendor,TSPL_SD_SHIPMENT_HEAD.DCS_Price_Code,TSPL_SD_SHIPMENT_HEAD.Supply_Date,TSPL_SD_SHIPMENT_HEAD.Shift_Type,TSPL_SD_SHIPMENT_HEAD.Salesman_Code ,TSPL_SD_SHIPMENT_HEAD.Salesman_Name,  "
        qry += " TSPL_SD_SHIPMENT_HEAD.CURRENCY_CODE,TSPL_SD_SHIPMENT_HEAD.CONVRATE,TSPL_SD_SHIPMENT_HEAD.APPLICABLEFROM,TSPL_SD_SHIPMENT_HEAD.PRoject_ID ,TSPL_SD_SHIPMENT_HEAD.Mannual_Invoice_No,TSPL_SD_SHIPMENT_HEAD. Mannual_Invoice_No_StringType,TSPL_SD_SHIPMENT_HEAD.Form_38_No " &
        " ,TSPL_SD_SHIPMENT_HEAD.SO_Validity,TSPL_SD_SHIPMENT_HEAD.Commission_Apply,TSPL_SD_SHIPMENT_HEAD.Total_Comm_Amt,TSPL_SD_SHIPMENT_HEAD.Dispatch_date,TSPL_SD_SHIPMENT_HEAD.WayBillNo,TSPL_SD_SHIPMENT_HEAD.WayBillDate " &
        " ,TSPL_SD_SHIPMENT_HEAD.Dispatch_Terms,TSPL_SD_SHIPMENT_HEAD.Payment_Terms,TSPL_SD_SHIPMENT_HEAD.Dispatch_Period,TSPL_SD_SHIPMENT_HEAD.Vehicle_Capacity,TSPL_SD_SHIPMENT_HEAD.RoundOffAmount,TSPL_SD_SHIPMENT_HEAD.Is_Taxable, TSPL_SD_SHIPMENT_HEAD.Electronic_Ref_No " &
        ",TSPL_SD_SHIPMENT_HEAD.Receipt_No,TSPL_SD_SHIPMENT_HEAD.ReceiptAmt,TSPL_SD_SHIPMENT_HEAD.VehicleNo,TSPL_SD_SHIPMENT_HEAD.ReceiverName,TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt ,TSPL_SD_SHIPMENT_HEAD.TotalSubsidyDisAmt,TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt,TSPL_SD_SHIPMENT_HEAD.Deduction_Type ,TSPL_SD_SHIPMENT_HEAD.Deduction,TSPL_SD_SHIPMENT_HEAD.IsEwayBill,TSPL_SD_SHIPMENT_HEAD.No_Transporter,TSPL_SD_SHIPMENT_HEAD.Transport_Id,TSPL_SD_SHIPMENT_HEAD.Transporter_Name "
        qry += "  FROM TSPL_SD_SHIPMENT_HEAD "
        qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SD_SHIPMENT_HEAD.Bill_To_Location "
        qry += " left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SD_SHIPMENT_HEAD.Ship_To_Location "
        qry += " left outer join  TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code= TSPL_SD_SHIPMENT_HEAD.Tax_Group "
        qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_SD_SHIPMENT_HEAD.Terms_Code "
        qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code where 2=2"
        Dim whrCls As String = ""
        '-------richa 30/07/2014 Ticket No. BM00000003242---------
        Dim strwherecls As String = ""
        If clsCommon.CompairString(clsCommon.myCstr(NavType).ToUpper(), "CURRENT") <> CompairStringResult.Equal Then
            strwherecls = FrmMainTranScreen.CustomerPermission()
        End If
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrCls = " AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        'End If

        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 And clsCommon.myLen(strwherecls) > 0 Then
            whrCls = "  and TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + strwherecls + ") "
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrCls = "  and TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        ElseIf clsCommon.myLen(strwherecls) > 0 Then
            whrCls = "  and TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' AND TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + strwherecls + ")"
        Else
            whrCls = " and TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' "
        End If
        '-----------------------------------------------------

        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code = (select MIN(Document_Code) from TSPL_SD_SHIPMENT_HEAD WHERE 1=1 " + whrCls + ") "
            Case NavigatorType.Last
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code = (select Max(Document_Code) from TSPL_SD_SHIPMENT_HEAD WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Current
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code = '" + strPONo + "' "
            Case NavigatorType.Next
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code = (select Min(Document_Code) from TSPL_SD_SHIPMENT_HEAD where Document_Code>'" + strPONo + "' " + whrCls + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code = (select Max(Document_Code) from TSPL_SD_SHIPMENT_HEAD where Document_Code<'" + strPONo + "' " + whrCls + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsMCCMaterialSale()
            'obj.WayBillDate = clsCommon.myCDate(dt.Rows(0)("WayBillDate"))
            obj.WayBillNo = clsCommon.myCstr(dt.Rows(0)("WayBillNo"))
            If dt.Rows(0)("EWayBillDate") IsNot DBNull.Value Then
                obj.EWayBillDate = clsCommon.myCDate(dt.Rows(0)("EWayBillDate"))
            End If
            obj.Deduction_Type = clsCommon.myCstr(dt.Rows(0)("Deduction_Type"))
            obj.Deduction = clsCommon.myCstr(dt.Rows(0)("Deduction"))
            obj.Electronic_Ref_No = clsCommon.myCstr(dt.Rows(0)("Electronic_Ref_No"))
            obj.EWayBillNo = clsCommon.myCstr(dt.Rows(0)("EWayBillNo"))
            obj.SO_Validity = clsCommon.myCdbl(dt.Rows(0)("SO_Validity"))
            obj.Commission_Apply = clsCommon.myCdbl(dt.Rows(0)("Commission_Apply"))
            obj.Total_Comm_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Comm_Amt"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            obj.No_Of_Instalment = clsCommon.myCdbl(dt.Rows(0)("No_Of_Instalment"))
            obj.Dispatch_date = clsCommon.myCDate(dt.Rows(0)("Dispatch_date"))
            obj.Vehicle_Capacity = clsCommon.myCdbl(dt.Rows(0)("Vehicle_Capacity"))
            obj.Dispatch_Terms = clsCommon.myCstr(dt.Rows(0)("Dispatch_Terms"))
            obj.Payment_Terms = clsCommon.myCstr(dt.Rows(0)("Payment_Terms"))
            obj.Dispatch_Period = clsCommon.myCdbl(dt.Rows(0)("Dispatch_Period"))
            obj.Road_Permit_No = clsCommon.myCstr(dt.Rows(0)("Road_Permit_No"))
            obj.Transport_Id = clsCommon.myCstr(dt.Rows(0)("Transport_Id"))
            obj.Transporter_Name = clsCommon.myCstr(dt.Rows(0)("Transporter_Name"))
            obj.IS_TCS = clsCommon.myCstr(dt.Rows(0)("IS_TCS"))
            obj.No_Transporter = clsCommon.myCdbl(dt.Rows(0)("No_Transporter"))
            obj.IsEwayBill = clsCommon.myCdbl(dt.Rows(0)("IsEwayBill"))
            obj.Is_Delivered = clsCommon.myCdbl(dt.Rows(0)("Is_Delivered"))
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Customer_Code = clsCommon.myCstr(dt.Rows(0)("Customer_Code"))
            obj.Customer_Name = clsCommon.myCstr(dt.Rows(0)("Customer_Name"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Bank_Code = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
            obj.Is_Internal = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Internal")) = 1, True, False)
            obj.Ref_No = clsCommon.myCstr(dt.Rows(0)("Ref_No"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Is_CashSale = clsCommon.myCstr(dt.Rows(0)("Is_CashSale"))
            obj.Receipt_No = clsCommon.myCstr(dt.Rows(0)("Receipt_No"))
            obj.ReceiptAmt = clsCommon.myCdbl(dt.Rows(0)("ReceiptAmt"))
            obj.VehicleNo = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
            obj.ReceiverName = clsCommon.myCstr(dt.Rows(0)("ReceiverName"))
            obj.TotalSubsidyAmt = clsCommon.myCdbl(dt.Rows(0)("TotalSubsidyAmt"))
            obj.TotalSubsidyDisAmt = clsCommon.myCdbl(dt.Rows(0)("TotalSubsidyDisAmt"))
            obj.Transporter_Commission_TotalAmt = clsCommon.myCdbl(dt.Rows(0)("Transporter_Commission_TotalAmt"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Bill_To_Location = clsCommon.myCstr(dt.Rows(0)("Bill_To_Location"))
            obj.Sub_Location_code = clsCommon.myCstr(dt.Rows(0)("Sub_Location_code"))
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
            obj.Comments = clsCommon.myCstr(dt.Rows(0)("Comments"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.PROJECT_ID = clsCommon.myCstr(dt.Rows(0)("PROJECT_ID"))

            If dt.Rows(0)("Due_Date") IsNot DBNull.Value Then
                obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            End If

            '-----------------richa 26/06/2014 Ticket No .BM00000002982----------------------------
            obj.InvoiceManualNowithPrefix = clsCommon.myCstr(dt.Rows(0)("Mannual_Invoice_No_StringType"))
            obj.Mannual_Invoice_No = clsCommon.myCstr(dt.Rows(0)("Mannual_Invoice_No"))

            '--------------------------------------------------------------------------------------

            obj.BillToLocationName = clsCommon.myCstr(dt.Rows(0)("BillToLocationName"))
            obj.ShipToLocationName = clsCommon.myCstr(dt.Rows(0)("ShipToLocationName"))
            obj.TaxGroupName = clsCommon.myCstr(dt.Rows(0)("TaxGroupName"))
            obj.TermsName = clsCommon.myCstr(dt.Rows(0)("TermsName"))

            If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
            End If


            obj.Challan_No = clsCommon.myCdbl(dt.Rows(0)("Challan_No"))
            obj.Carrier = clsCommon.myCstr(dt.Rows(0)("Carrier"))
            obj.VehicleNo = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
            obj.Vehicle_Code = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
            obj.GRNo = clsCommon.myCstr(dt.Rows(0)("GRNo"))
            obj.GENo = clsCommon.myCstr(dt.Rows(0)("GENo"))
            If dt.Rows(0)("GEDate") IsNot DBNull.Value Then
                obj.GEDate = clsCommon.myCDate(dt.Rows(0)("GEDate"))
            End If




            obj.Dept = clsCommon.myCstr(dt.Rows(0)("Dept"))
            obj.Dept_Desc = clsCommon.myCstr(dt.Rows(0)("Dept_Desc"))
            obj.Item_Type = clsCommon.myCstr(dt.Rows(0)("Item_Type"))

            obj.Against_Sales_Order = clsCommon.myCstr(dt.Rows(0)("Against_Sales_Order"))


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
            If dt.Rows(0)("Challan_Date") IsNot DBNull.Value Then

                obj.Challan_Date = clsCommon.GetPrintDate((dt.Rows(0)("Challan_Date")), "dd/MMM/yyyy")
            End If

            If dt.Rows(0)("Inv_Date") IsNot DBNull.Value Then
                obj.Inv_Date = clsCommon.GetPrintDate((dt.Rows(0)("Inv_Date")), "dd/MMM/yyyy")
            End If

            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)
            obj.Is_Create_Auto_Invoice = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Create_Auto_Invoice")) = 1, True, False)
            obj.Sale_Invoice_No = clsCommon.myCstr(dt.Rows(0)("Sale_Invoice_No"))
            obj.Is_Create_Auto_Receipt = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Create_Auto_Receipt")) = 1, True, False)
            obj.Is_Apply_TPT = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Apply_TPT")) = 1, True, False)
            obj.Is_Add_TPT = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Add_TPT")) = 1, True, False)
            obj.Exclude_KKF_And_Mandi = IIf(clsCommon.myCdbl(dt.Rows(0)("Exclude_KKF_And_Mandi")) = 1, True, False)
            obj.TPT_Vendor = clsCommon.myCstr(dt.Rows(0)("TPT_Vendor"))
            obj.DCS_Price_Code = clsCommon.myCstr(dt.Rows(0)("DCS_Price_Code"))
            If dt.Rows(0)("Supply_Date") IsNot DBNull.Value Then
                obj.Supply_Date = clsCommon.myCstr(dt.Rows(0)("Supply_Date"))
            End If
            obj.Shift_Type = clsCommon.myCstr(dt.Rows(0)("Shift_Type"))
            obj.Recommended_By = clsCommon.myCstr(dt.Rows(0)("Recommended_By"))
            obj.Salesman_Code = clsCommon.myCstr(dt.Rows(0)("Salesman_Code"))
            obj.Salesman_Name = clsCommon.myCstr(dt.Rows(0)("Salesman_Name"))
            obj.Price_Code = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
            obj.Route_No = clsCommon.myCstr(dt.Rows(0)("Route_No"))
            obj.Route_Desc = clsCommon.myCstr(dt.Rows(0)("Route_Desc"))
            obj.HeadDisc_Per = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Per"))
            obj.HeadDisc_PerAmt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_PerAmt"))
            obj.HeadDisc_Amt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Amt"))
            obj.RateDiff_Per = clsCommon.myCdbl(dt.Rows(0)("RateDiff_Per"))
            obj.Rate_Diff_Amount_Type = clsCommon.myCdbl(dt.Rows(0)("Rate_Diff_Amount_Type"))
            obj.Is_Rate_Diff_Per_Amt = clsCommon.myCdbl(dt.Rows(0)("Is_Rate_Diff_Per_Amt"))
            obj.Gross_Amount = clsCommon.myCdbl(dt.Rows(0)("Gross_Amount"))
            obj.RateDiff_Amt = clsCommon.myCdbl(dt.Rows(0)("RateDiff_Amt"))
            obj.TotCashDiscAmt = clsCommon.myCdbl(dt.Rows(0)("TotCashDiscAmt"))
            obj.Invoice_Type = clsCommon.myCstr(dt.Rows(0)("Invoice_Type"))
            obj.Price_Group_Code = clsCommon.myCstr(dt.Rows(0)("Price_Group_Code"))
            obj.Cust_PO_No = clsCommon.myCstr(dt.Rows(0)("Cust_PO_No"))
            If dt.Rows(0)("cust_po_date") IsNot DBNull.Value Then
                obj.Podate = clsCommon.myCDate(dt.Rows(0)("cust_po_date"))
            End If
            obj.Form_38_No = clsCommon.myCstr(dt.Rows(0)("Form_38_No"))
            obj.Mannual_Invoice_No = clsCommon.myCdbl(dt.Rows(0)("Mannual_Invoice_No"))
            'richa Ticket No.BM00000002982-------------
            obj.InvoiceManualNowithPrefix = clsCommon.myCstr(dt.Rows(0)("Mannual_Invoice_No_StringType"))
            '  ---------------------------------------------
            '' CURRENCYCONVERSION 
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            '' END CURRENCYCONVERSION 
            obj.Invoice_No = clsDBFuncationality.getSingleValue("Select Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No='" & obj.Document_Code & "' ", trans)
            obj.Is_Taxable = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Taxable")) = 1, True, False)

            qry = "SELECT  TSPL_SD_SHIPMENT_DETAIL.ItemwiseTaxCode,TSPL_SD_SHIPMENT_DETAIL.OrgUnit_code,TSPL_SD_SHIPMENT_DETAIL.Is_Mannual_Amt,TSPL_SD_SHIPMENT_DETAIL.Document_Code,TSPL_SD_SHIPMENT_DETAIL.Line_No, " &
            "TSPL_SD_SHIPMENT_DETAIL.Status,TSPL_SD_SHIPMENT_DETAIL.Row_Type,TSPL_SD_SHIPMENT_DETAIL.status,TSPL_SD_SHIPMENT_DETAIL.Item_Code, " &
            "TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_MASTER.DCS_Sale_Zero_Cost,0)as DCS_Sale_Zero_Cost,TSPL_SD_SHIPMENT_DETAIL.Qty,TSPL_SD_SHIPMENT_DETAIL.Free_Qty,TSPL_SD_SHIPMENT_DETAIL.Order_Code, " &
            "TSPL_SD_SHIPMENT_DETAIL.Order_Code,TSPL_SD_SHIPMENT_DETAIL.Balance_Qty,TSPL_SD_SHIPMENT_DETAIL.Unit_code, " &
            "TSPL_SD_SHIPMENT_DETAIL.Location,TSPL_SD_SHIPMENT_DETAIL.Item_Cost,TSPL_SD_SHIPMENT_DETAIL.TAX1,TSPL_SD_SHIPMENT_DETAIL.TAX1_Rate, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX2,TSPL_SD_SHIPMENT_DETAIL.TAX2_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX3,TSPL_SD_SHIPMENT_DETAIL.TAX3_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX4 , " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX4_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX5,TSPL_SD_SHIPMENT_DETAIL.TAX5_Rate , " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX6,TSPL_SD_SHIPMENT_DETAIL.TAX6_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX7,TSPL_SD_SHIPMENT_DETAIL.TAX7_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX8, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX8_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX9,TSPL_SD_SHIPMENT_DETAIL.TAX9_Rate, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX10,TSPL_SD_SHIPMENT_DETAIL.TAX10_Rate,TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.Amount,TSPL_SD_SHIPMENT_DETAIL.Disc_Per,TSPL_SD_SHIPMENT_DETAIL.Disc_Amt,TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount, " &
            "TSPL_SD_SHIPMENT_DETAIL.Total_Tax_Amt,TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt,TSPL_LOCATION_MASTER.Location_Desc as LocationName, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX1_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX2_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX3_Base_Amt , " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX4_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX5_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX6_Base_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX7_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX8_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.TAX9_Base_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.TAX10_Base_Amt,TSPL_SD_SHIPMENT_DETAIL.MRP,TSPL_SD_SHIPMENT_DETAIL.Batch_No,TSPL_SD_SHIPMENT_DETAIL.MFG_Date, " &
            "TSPL_SD_SHIPMENT_DETAIL.Expiry_Date,TSPL_SD_SHIPMENT_DETAIL.Specification,TSPL_SD_SHIPMENT_DETAIL.Remarks,TSPL_SD_SHIPMENT_DETAIL.Assessable, " &
            "TSPL_SD_SHIPMENT_DETAIL.AssessableAmt,TSPL_SD_SHIPMENT_DETAIL.Bar_Code, " &
            "TSPL_SD_SHIPMENT_DETAIL.Scheme_Applicable,TSPL_SD_SHIPMENT_DETAIL.Scheme_Code, " &
            "TSPL_SD_SHIPMENT_DETAIL.Scheme_Item,TSPL_SD_SHIPMENT_DETAIL.Item_Tax,TSPL_SD_SHIPMENT_DETAIL.Total_MRP_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.Total_Basic_Amt,TSPL_SD_SHIPMENT_DETAIL.Total_Disc_Amt,TSPL_SD_SHIPMENT_DETAIL.Cust_Discount, " &
            "TSPL_SD_SHIPMENT_DETAIL.Total_Cust_Discount,TSPL_SD_SHIPMENT_DETAIL.ActualRate,TSPL_SD_SHIPMENT_DETAIL.Cust_DiscountQty, " &
            "TSPL_SD_SHIPMENT_DETAIL.Price_code,TSPL_SD_SHIPMENT_DETAIL.Abatement_Per,TSPL_SD_SHIPMENT_DETAIL.Abatement_Amt, " &
            "TSPL_SD_SHIPMENT_DETAIL.FOC_Item,TSPL_SD_SHIPMENT_DETAIL.Item_Weight,TSPL_SD_SHIPMENT_DETAIL.Price_Date, " &
            "TSPL_SD_SHIPMENT_DETAIL.HeadDiscPer,TSPL_SD_SHIPMENT_DETAIL.HeadDiscPerAmt,TSPL_SD_SHIPMENT_DETAIL.Bin_No,TSPL_SD_SHIPMENT_DETAIL.TotalItem_Weight,TSPL_SD_SHIPMENT_DETAIL.Conv_Factor,TSPL_SD_SHIPMENT_DETAIL.Purchase_Cost,TSPL_SD_SHIPMENT_DETAIL.OrgRate,  " &
            "TSPL_SD_SHIPMENT_DETAIL.vendor_code,TSPL_SD_SHIPMENT_DETAIL.vendor_desc,TSPL_SD_SHIPMENT_DETAIL.PrincipleCode,TSPL_SD_SHIPMENT_DETAIL.PrincipleDesc,TSPL_SD_SHIPMENT_DETAIL.Markup_On,TSPL_SD_SHIPMENT_DETAIL.Markup_Percent,TSPL_SD_SHIPMENT_DETAIL.Landing_Cost,TSPL_SD_SHIPMENT_DETAIL.HeadDiscAmt,TSPL_SD_SHIPMENT_DETAIL.CustDiscPer,TSPL_SD_SHIPMENT_DETAIL.CasdDiscScheme_Code " &
            ",TSPL_SD_SHIPMENT_DETAIL.Commission_Rate,TSPL_SD_SHIPMENT_DETAIL.Billing_Qty,TSPL_SD_SHIPMENT_DETAIL.Billing_Unit_Code,TSPL_SD_SHIPMENT_DETAIL.Commission_Party,TSPL_SD_SHIPMENT_DETAIL.Commission_Amt,TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Commission,TSPL_SD_SHIPMENT_DETAIL.Transporter_Commission_Rate,TSPL_SD_SHIPMENT_DETAIL.Transporter_Commission_Amt,TSPL_SD_SHIPMENT_DETAIL.REF_TPT_PK_ID,TSPL_SD_SHIPMENT_DETAIL.Disc_Per_Unit,TSPL_SD_SHIPMENT_DETAIL.Disc_Unit_Amt "
            qry += " FROM TSPL_SD_SHIPMENT_DETAIL "
            qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SD_SHIPMENT_DETAIL.Location "
            qry += " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code"
            qry += " where TSPL_SD_SHIPMENT_DETAIL.Document_Code='" + obj.Document_Code + "' ORDER BY TSPL_SD_SHIPMENT_DETAIL.Line_No  asc"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsMCCMaterialSaleDetail)
                Dim objTr As clsMCCMaterialSaleDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsMCCMaterialSaleDetail
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(dr("ItemwiseTaxCode"))
                    objTr.Commission_Rate = clsCommon.myCdbl(dr("Commission_Rate"))
                    objTr.Commission_Party = clsCommon.myCstr(dr("Commission_Party"))
                    objTr.Commission_Amt = clsCommon.myCdbl(dr("Commission_Amt"))
                    objTr.Amt_Less_Commission = clsCommon.myCdbl(dr("Amt_Less_Commission"))
                    objTr.Transporter_Commission_Rate = clsCommon.myCdbl(dr("Transporter_Commission_Rate"))
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(dr("Transporter_Commission_Amt"))
                    objTr.REF_TPT_PK_ID = clsCommon.myCdbl(dr("REF_TPT_PK_ID"))
                    objTr.OrgUnit_code = clsCommon.myCstr(dr("OrgUnit_code"))
                    objTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objTr.Row_Type = clsCommon.myCstr(dr("Row_Type"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Status = Convert.ToInt32(clsCommon.myCdbl(dr("Status")))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.Bar_Code = clsCommon.myCstr(dr("Bar_Code"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.Qty = clsCommon.myCdbl(dr("Qty"))
                    objTr.Billing_Qty = clsCommon.myCdbl(dr("Billing_Qty"))
                    objTr.Billing_Unit_code = clsCommon.myCstr(dr("Billing_Unit_code"))
                    objTr.DCS_Sale_Zero_Cost = clsCommon.myCdbl(dr("DCS_Sale_Zero_Cost"))
                    objTr.Free_Qty = clsCommon.myCdbl(dr("Free_Qty"))
                    objTr.Order_Code = clsCommon.myCstr(dr("Order_Code"))

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
                    objTr.Disc_Per_Unit = clsCommon.myCdbl(dr("Disc_Per_Unit"))
                    objTr.Disc_Unit_Amt = clsCommon.myCdbl(dr("Disc_Unit_Amt"))
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
                    'objTr.Unit_Cogs = clsCommon.myCdbl(dr("Unit_Cogs"))

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
                    If dr("Price_Date") IsNot DBNull.Value Then
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
                    objTr.Purchase_Cost = clsCommon.myCdbl(dr("Purchase_Cost"))
                    objTr.OrgRate = clsCommon.myCdbl(dr("OrgRate"))
                    objTr.Conv_Factor = clsCommon.myCdbl(dr("Conv_Factor"))
                    objTr.PrincipleCode = clsCommon.myCstr(dr("PrincipleCode"))
                    objTr.PrincipleDesc = clsCommon.myCstr(dr("PrincipleDesc"))
                    objTr.vendor_code = clsCommon.myCstr(dr("vendor_code"))
                    objTr.vendor_desc = clsCommon.myCstr(dr("vendor_desc"))
                    objTr.Bin_No = clsCommon.myCstr(dr("Bin_No"))
                    objTr.HeadDiscPer = clsCommon.myCdbl(dr("HeadDiscPer"))
                    objTr.HeadDiscPerAmt = clsCommon.myCdbl(dr("HeadDiscPerAmt"))
                    objTr.arrSrItem = clsSerializeInvenotry.GetData("SD-IN", objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    objTr.arrBatchItem = clsBatchInventory.GetData("MCC-MSALE", objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr.Add(objTr)
                Next
            End If
        End If

        Return obj
    End Function

    Public Shared Function GetOriginalQty(ByVal strMrnNo As String, ByVal strICode As String, ByVal strUOM As String, ByVal dblAssessable As Double, ByVal dblMRP As Double, ByVal trans As SqlTransaction) As DataTable
        Dim qry As String = "Select TSPL_MRN_DETAIL.MRN_No,(TSPL_MRN_DETAIL.MRN_Qty+ISNULL(TSPL_MRN_DETAIL.Leak_Qty,0) +ISNULL(TSPL_MRN_DETAIL.Burst_Qty,0)+ISNULL(TSPL_MRN_DETAIL.Short_Qty,0)) as MRN_Qty,TSPL_GRN_DETAIL.GRN_No,(TSPL_GRN_DETAIL.GRN_Qty+ISNULL(TSPL_GRN_DETAIL.Leak_Qty,0) +ISNULL(TSPL_GRN_DETAIL.Burst_Qty,0)+ISNULL(TSPL_GRN_DETAIL.Short_Qty,0)) as GRN_Qty, TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No,TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_Qty from TSPL_MRN_DETAIL left outer join TSPL_GRN_DETAIL on TSPL_GRN_DETAIL.GRN_No=TSPL_MRN_DETAIL.GRN_Id and TSPL_GRN_DETAIL.Item_Code=TSPL_MRN_DETAIL.Item_Code and TSPL_GRN_DETAIL.Unit_code=TSPL_MRN_DETAIL.Unit_code and isnull(TSPL_GRN_DETAIL.Assessable,0)=isnull(TSPL_MRN_DETAIL.Assessable,0) and isnull(TSPL_GRN_DETAIL.Item_Code,0)=isnull(TSPL_MRN_DETAIL.Item_Code ,0) left outer join TSPL_PURCHASE_ORDER_DETAIL on TSPL_PURCHASE_ORDER_DETAIL.PurchaseOrder_No= TSPL_GRN_DETAIL.PO_Id and TSPL_PURCHASE_ORDER_DETAIL.Item_Code=TSPL_GRN_DETAIL.Item_Code and TSPL_PURCHASE_ORDER_DETAIL.Unit_code=  TSPL_GRN_DETAIL.Unit_code and isnull(TSPL_PURCHASE_ORDER_DETAIL.Assessable,0)=  isnull(TSPL_GRN_DETAIL.Assessable,0) and isnull(TSPL_PURCHASE_ORDER_DETAIL.MRP,0)=  isnull(TSPL_GRN_DETAIL.MRP,0) where TSPL_MRN_DETAIL.MRN_No='" + strMrnNo + "' and TSPL_MRN_DETAIL.Item_Code='" + strICode + "' and TSPL_MRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_MRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_MRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'"
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Dim GeneratePDFForMobile As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GeneratePDFForMobile, clsFixedParameterCode.GeneratePDFForMobile, trans))
        Try
            PostData(FormId, strDocNo, trans)
            trans.Commit()
            If GeneratePDFForMobile = 1 Then
                UpdateFileInfo(FormId, strDocNo)
            End If
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean

        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Shipment No not found to Post")
            End If
            Dim obj As clsMCCMaterialSale = clsMCCMaterialSale.GetData(strDocNo, NavigatorType.Current, trans)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            '' Anubhooti 06-Sep-2014 BM00000003735 (Locked Transaction)
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Sales And Distribution", "Shipment/Sale Invoice", obj.Bill_To_Location, obj.Document_Date, trans)
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmMCCMaterial, obj.Bill_To_Location, obj.Document_Date, trans)

            ''
            If (obj.Status = 1) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            If (obj.On_Hold) Then
                Throw New Exception("Shipment No " + obj.Document_Code + " Is On Hold.Can't Post it")
            End If

            Dim qry As String = ""

            Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", obj.Document_Code, trans)
            If isResult = False Then
                trans.Commit()
                Return False
            End If

            ''to check Qty available or not ERO/16/08/19-000993
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then

                    For ii As Integer = 0 To obj.Arr.Count - 1
                        Dim strICode As String = clsCommon.myCstr(obj.Arr(ii).Item_Code)
                        Dim strIName As String = clsCommon.myCstr(obj.Arr(ii).Item_Desc)
                        Dim dblQty As Double = clsCommon.myCdbl(obj.Arr(ii).Qty)
                        Dim dblMRP As Double = clsCommon.myCdbl(obj.Arr(ii).MRP)
                        Dim strUOM As String = clsCommon.myCstr(obj.Arr(ii).Unit_code)
                        Dim strSchemeItem As String = clsCommon.myCstr(obj.Arr(ii).Scheme_Item)
                        Dim strProject As String = ""
                        If clsCommon.myLen(strICode) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Arr(ii).Row_Type), "Item") = CompairStringResult.Equal Then
                            If clsCommon.myLen(strICode) > 0 AndAlso clsCommon.CompairString(strSchemeItem, "N") = CompairStringResult.Equal Then
                                For jj As Integer = ii + 1 To obj.Arr.Count - 1
                                    Dim strInICode As String = clsCommon.myCstr(obj.Arr(jj).Item_Code)
                                    Dim strInUOM As String = clsCommon.myCstr(obj.Arr(jj).Unit_code)


                                    If clsCommon.CompairString(strICode, strInICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOM, strInUOM) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Arr(ii).Scheme_Item), "No") = CompairStringResult.Equal Then
                                        Throw New Exception("Item Code " + strICode + " and Unit " + strUOM + " is repeated at Row No" + clsCommon.myCstr(ii + 1) + " and " + clsCommon.myCstr(jj + 1))
                                    End If
                                Next
                            End If
                            Dim dblOuterConvFac As Double = clsItemMaster.GetConvertionFactor(strICode, strUOM, trans)
                            Dim Location As String = Nothing
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(obj.Bill_To_Location) & "'", trans)), "Y") = CompairStringResult.Equal Then
                                Location = obj.Sub_Location_code
                            Else
                                Location = obj.Bill_To_Location

                            End If
                            Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strICode, Location, obj.Document_Code, obj.Document_Date, trans, strUOM, dblMRP)
                            Dim dblEnteredQty As Double = dblQty
                            For jj As Integer = 0 To obj.Arr.Count - 1
                                If ii = jj Then
                                    Continue For
                                End If
                                Dim strICodeInner As String = clsCommon.myCstr(obj.Arr(jj).Item_Code)
                                Dim strUOMInner As String = clsCommon.myCstr(obj.Arr(jj).Item_Code)
                                Dim dblQtyInner As Double = clsCommon.myCdbl(obj.Arr(jj).Item_Code)
                                Dim dblInnerConvFac As Double = clsItemMaster.GetConvertionFactor(strICodeInner, strUOMInner, trans)
                                If dblQtyInner > 0 AndAlso clsCommon.CompairString(strICodeInner, strICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOMInner, strUOM) = CompairStringResult.Equal Then
                                    dblEnteredQty += dblQtyInner
                                End If
                            Next
                            dblEnteredQty = Math.Round(dblEnteredQty, 2, MidpointRounding.ToEven)
                            If dblEnteredQty > dblBalQty Then ' AndAlso clsCommon.CompairString(strSchemeItem, "No") = CompairStringResult.Equal Then
                                Throw New Exception("Item - " + strICode + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblEnteredQty) + " and Balance Quantity - " + clsCommon.myCstr(dblBalQty))
                            End If
                        End If
                    Next
                End If
            End If


            ''---------------


            UpdateInventoryMovement(obj, trans, False)

            CreateJournalEntry(obj.Document_Code, trans, strVoucherNoForRecreatedOnly)

            qry = "Update TSPL_SD_SHIPMENT_HEAD set "
            If clsCommon.myLen(obj.Against_Sales_Order) = 0 Then
                obj.Against_Sales_Order = "NULL"
            End If
            qry += "Against_Sales_Order=" & obj.Against_Sales_Order & ", Status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "',Sale_Invoice_No ='" + obj.Invoice_No + "' "
            qry += " where Document_Code='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            If objCommonVar.CreateAutoReceiptEntryDCSSale Then
                If clsCommon.CompairString(obj.Is_CashSale, "Y") = CompairStringResult.Equal Then
                    obj.RecieptEntryOfDCSSale(obj, trans)
                End If
            End If

            If obj.Is_Create_Auto_Invoice Then
                clsPSInvoiceHead.PostData("", obj.Invoice_No, trans)   ''obj.Sale_Invoice_No remove because it has not value.by bulk posting.
            End If

            'Comment by balwinder on 13/02/2026 as AR Invoice created after reduce the subsidy amount.
            'If (obj.TotalSubsidyAmt > 0) Then
            '    CreateARAdjustmentEntryAgainstSubsidy(obj, trans)
            'End If

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", trans)
            ''richa BHO/08/07/21-000019
            Dim CreateARAdjAPDebitnoteforEmployeesinMCCMS As Boolean = IIf(clsFixedParameter.GetData(clsFixedParameterType.CreateARAdjAPDebitnoteforEmployeesinMCCMS, clsFixedParameterCode.CreateARAdjAPDebitnoteforEmployeesinMCCMS, trans) = 1, True, False)
            If CreateARAdjAPDebitnoteforEmployeesinMCCMS = True Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Form_type,IsEmployee,Vendor_Type_CHA  from TSPL_VENDOR_MASTER where vendor_code='" & obj.Customer_Code & "'", trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Form_type")), "PTM") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("IsEmployee")), "1") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Vendor_Type_CHA")), "CV") = CompairStringResult.Equal Then
                        CreateARAdjustmentEntry(obj, trans)
                        APInvoice_DebitNote(obj, trans)
                    End If
                End If
            End If

            If clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.Sale_SMSATPOST, clsFixedParameterCode.Sale_SMSATPOST, trans) = "1", True, False)) Then
                SMSSENDONLY(obj, trans, True)
            End If
            Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceType(obj.Customer_Code, trans)
            If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso obj.Is_Taxable AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True AndAlso obj.Total_Amt > 0 Then
                If clsCommon.myLen(GetIRNNo(obj.Sale_Invoice_No, trans)) <= 0 Then
                    clsPSInvoiceHead.EInvoice_Implementation(obj.Sale_Invoice_No, obj.Bill_To_Location, trans, False)
                    If clsCommon.myLen(GetIRNNo(obj.Sale_Invoice_No, trans)) <= 0 Then
                        Throw New Exception("IRN No For Sales Invoice No [" + obj.Sale_Invoice_No + "] is not generated")
                    End If
                End If
                If obj.IsEwayBill = 1 Then
                    If objCommonVar.GenerateEWayBillWithEInvoice Then
                        If clsCommon.myLen(GetEWayBillNo(strDocNo, trans)) <= 0 Then
                            clsPSInvoiceHead.EInvoice_Implementation(obj.Sale_Invoice_No, obj.Bill_To_Location, trans, True)
                            If clsCommon.myLen(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans)) <= 0 Then
                                'Throw New Exception("E-Way Bill For Sales Invoice No [" + strDocNo + "] is not generated")
                            End If
                        End If
                    End If
                End If
            ElseIf clsCommon.CompairString(ECustomerType, "BC") = CompairStringResult.Equal Then
                Dim EnableDynamicQRCodeForB2CInvoice As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EnableDynamicQRCodeForB2CInvoice, clsFixedParameterCode.EnableDynamicQRCodeForB2CInvoice, trans)) = 1, True, False))
                If EnableDynamicQRCodeForB2CInvoice = True AndAlso clsERPFuncationality.GetQRCodeStatus(obj.Document_Date, trans) = True Then
                    clsPSInvoiceHead.EInvoice_ImplementationFor_CustomerType_BC(obj.Document_Code, obj.Bill_To_Location, trans)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GenereateDCSEInvoice(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim obj As clsMCCMaterialSale = clsMCCMaterialSale.GetData(strDocNo, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceType(obj.Customer_Code, trans)
            If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso obj.Is_Taxable AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True Then
                If clsCommon.myLen(GetIRNNo(obj.Sale_Invoice_No, trans)) <= 0 Then
                    clsPSInvoiceHead.EInvoice_Implementation(obj.Sale_Invoice_No, obj.Bill_To_Location, trans, False)
                    If clsCommon.myLen(GetIRNNo(obj.Sale_Invoice_No, trans)) <= 0 Then
                        Throw New Exception("IRN No For Sales Invoice No [" + obj.Sale_Invoice_No + "] is not generated")
                    End If
                End If
                'If obj.IsEwayBill = 1 Then
                '    If objCommonVar.GenerateEWayBillWithEInvoice Then
                '        If clsCommon.myLen(GetEWayBillNo(strDocNo, trans)) <= 0 Then
                '            clsPSInvoiceHead.EInvoice_Implementation(obj.Sale_Invoice_No, obj.Bill_To_Location, trans, True)
                '            If clsCommon.myLen(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans)) <= 0 Then
                '                'Throw New Exception("E-Way Bill For Sales Invoice No [" + strDocNo + "] is not generated")
                '            End If
                '        End If
                '    End If
                'End If
            ElseIf clsCommon.CompairString(ECustomerType, "BC") = CompairStringResult.Equal Then
                Dim EnableDynamicQRCodeForB2CInvoice As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EnableDynamicQRCodeForB2CInvoice, clsFixedParameterCode.EnableDynamicQRCodeForB2CInvoice, trans)) = 1, True, False))
                If EnableDynamicQRCodeForB2CInvoice = True AndAlso clsERPFuncationality.GetQRCodeStatus(obj.Document_Date, trans) = True Then
                    clsPSInvoiceHead.EInvoice_ImplementationFor_CustomerType_BC(obj.Document_Code, obj.Bill_To_Location, trans)
                End If
            End If
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetEWayBillNo(strDocNo As String, trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans))
    End Function
    Public Shared Function GetIRNNo(strDocNo As String, trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(IRN_No,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans))
    End Function

    Public Shared Sub UpdateFileInfo(ByVal FormId As String, ByVal strDocNo As String)
        Try
            Dim qry As String = "select Document_Code,Document_Date from TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No='" + strDocNo + "' and FILE_INFO is null"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim frmCRV As New frmCrystalReportViewer()
                    Dim PDFPath As String = funPrint(FormId, False, clsCommon.myCDate(dt.Rows(0)("Document_Date")), strDocNo, True)
                    Dim FileNo As Integer = clsAttachDocument.UploadWithHttpRequest(PDFPath, Path.GetFileName(PDFPath), FormId, strDocNo)
                    Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                    Try
                        If FileNo > 0 Then
                            qry = " UPDATE TSPL_SD_SALE_INVOICE_HEAD set FILE_INFO=" + clsCommon.myCstr(FileNo) + " where Against_Shipment_No='" + strDocNo + "'"
                            clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        End If
                        trans.Commit()
                    Catch ex As Exception
                        trans.Rollback()
                        Throw New Exception(ex.Message)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Public Shared Sub SMSSENDONLY(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction, ByVal isPost As Boolean)
        Try
            Dim strContactPerson As String = ""
            Dim strotherno As String = clsDBFuncationality.getSingleValue("select REPLACE( REPLACE( REPLACE(Phone1,'_',''),'(+)',''),'(+91)','') as Phone1 from TSPL_customer_MASTER where cust_code ='" & obj.Customer_Code & "' ", trans)
            If clsCommon.myLen(strotherno) > 0 Then
                If clsCommon.myLen(strotherno) = 10 Then
                    strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & obj.Customer_Code & "' ", trans)
                    Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmMCCMaterial + "'", trans)
                    Dim objSMSH As New clsSMSHead()
                    objSMSH.arrMobilNo = New List(Of String)()
                    objSMSH.arrMobilNo.Add(clsCommon.myCstr(strotherno))
                    If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
                        If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                            Dim strItem As String = ""
                            Dim strQty As String = ""
                            If (obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0) Then
                                For Each objdetail As clsMCCMaterialSaleDetail In obj.Arr
                                    If clsCommon.myLen(strItem) > 0 Then
                                        strItem += ", "
                                        strQty += ", "
                                    End If
                                    strItem += clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_ITEM_MASTER.item_desc from TSPL_ITEM_MASTER where item_code ='" & objdetail.Item_Code & "' ", trans))
                                    strQty += clsCommon.myCstr(objdetail.Qty)
                                Next
                            End If


                            objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.InvoiceNo, obj.Invoice_No)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SaleOrderNo, obj.Document_Code)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SaleOrderDate, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerNo, obj.Customer_Code)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerName, obj.Customer_Name)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(obj.Total_Amt))
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, clsUserMgtCode.frmMCCMaterial)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ContactPerson, strContactPerson)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ItemCode, strItem)
                            objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Qty, strQty)
                        End If

                        If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                            objSMSH.SaveData(clsUserMgtCode.frmMCCMaterial, objSMSH, trans)
                            objSMSH = Nothing
                            If Not isPost Then
                                clsCommon.MyMessageBoxShow("SMS Send Successfully")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Shared Function UpdateInventoryMovement(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction, Optional ByVal UpdateInventory As Boolean = False) As Boolean
        Try
            Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
            If UpdateInventory = True Then
                clsDBFuncationality.ExecuteNonQuery("update tspl_batch_item set Against_Inv_Movement_Trans_Id=null where Document_Code='" & obj.Document_Code & "'", trans)
                clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & obj.Document_Code & "'", trans)
            End If
            Dim strRgpNo As String = Nothing
            Dim intCounter As Integer = 0
            For Each objTr As clsMCCMaterialSaleDetail In obj.Arr
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
                    Dim objInventoryMovemnt As New clsInventoryMovement()
                    objInventoryMovemnt.InOut = "O"
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(obj.Bill_To_Location) & "'", trans)), "Y") = CompairStringResult.Equal Then
                        objInventoryMovemnt.Location_Code = obj.Sub_Location_code
                    Else
                        objInventoryMovemnt.Location_Code = objTr.Location

                    End If
                    objInventoryMovemnt.Cust_Code = obj.Customer_Code
                    objInventoryMovemnt.Cust_Name = obj.Customer_Name
                    objInventoryMovemnt.Item_Code = objTr.Item_Code
                    objInventoryMovemnt.Item_Desc = objTr.Item_Desc
                    objInventoryMovemnt.Qty = objTr.Qty + objTr.Free_Qty
                    objInventoryMovemnt.UOM = objTr.Unit_code
                    objInventoryMovemnt.Basic_Cost = objTr.Item_Cost
                    objInventoryMovemnt.MRP = objTr.MRP
                    objInventoryMovemnt.Add_Cost = objTr.Total_Tax_Amt
                    objInventoryMovemnt.Net_Cost = objTr.Total_Tax_Amt
                    objInventoryMovemnt.ItemType = strItemTypeToSave
                    ArrInventoryMovement.Add(objInventoryMovemnt)
                End If
            Next
            clsInventoryMovement.SaveData("MCC-MSALE", obj.Document_Code, obj.Document_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy"), ArrInventoryMovement, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Shared Function APInvoice_DebitNote(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try

            Dim DisCCodeForArAdj As String = ""
            Dim GLAcARAdj As String = ""
            Dim DiscDiscForArAdj As String = ""
            Dim GLAcDescARAdj As String = ""

            Dim objVendInv As New clsVedorInvoiceHead()
            ''objCustInv.Document_No ''Will be Generateed
            objVendInv.Invoice_Entry_Date = obj.Document_Date

            objVendInv.Document_Type = "D"
            '' ARNote = "Debit Note"

            objVendInv.Against_MCC_Material_Sale = obj.Document_Code
            objVendInv.loc_code = clsLocation.GetSegmentCode(obj.Bill_To_Location, trans)
            objVendInv.Document_Total = clsCommon.myCdbl(obj.Total_Amt)
            objVendInv.Vendor_Code = clsCommon.myCstr(obj.Customer_Code)

            objVendInv.Invoice_Type = "AP"
            objVendInv.Vendor_Name = clsCommon.myCstr(obj.Customer_Name)
            objVendInv.Posting_Date = objVendInv.Invoice_Entry_Date
            objVendInv.Vendor_Invoice_Date = objVendInv.Invoice_Entry_Date
            objVendInv.Account_Set = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT  ISNULL(Vendor_Account,'') AS [Vendor_Account] FROM TSPL_VENDOR_MASTER WHERE Vendor_Code ='" + objVendInv.Vendor_Code + "'", trans))
            If (clsCommon.myLen(objVendInv.Account_Set) < 0) Then
                Throw New Exception("Please set the vendor account set for vendor : " + objVendInv.Vendor_Code)
            End If

            objVendInv.Vendor_Control_AC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select payable_account from TSPL_VENDOR_ACCOUNT_SET  where Acct_Set_Code='" + objVendInv.Account_Set + "'", trans))
            objVendInv.Vendor_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendInv.Vendor_Control_AC, objVendInv.loc_code, True, trans)

            If clsCommon.myLen(objVendInv.Vendor_Control_AC) <= 0 Then
                Throw New Exception("Please set the vendor payable Account")
            End If

            objVendInv.On_Hold = 0
            objVendInv.Remarks = "AP Debit Note Against MCC Material Sale - " & obj.Document_Code & ""
            objVendInv.Description = "AP Debit Note Against MCC Material Sale - " & obj.Document_Code & ""
            objVendInv.Balance_Amt = clsCommon.myCdbl(obj.Total_Amt)
            objVendInv.Amount_Less_Discount = clsCommon.myCdbl(obj.Total_Amt)
            objVendInv.Discount_Base = clsCommon.myCdbl(obj.Total_Amt)

            objVendInv.Arr = New List(Of clsVedorInvoiceDetail)

            '' Detail Level Saving

            Dim VendAccSet As String = String.Empty


            Dim objVendInvTR As clsVedorInvoiceDetail = New clsVedorInvoiceDetail()

            objVendInvTR.Detail_Line_No = 1
            objVendInvTR.Amount = clsCommon.myCdbl(obj.Total_Amt)
            objVendInvTR.Amount_less_Discount = clsCommon.myCdbl(obj.Total_Amt)
            objVendInvTR.Total_Amount = clsCommon.myCdbl(obj.Total_Amt)


            DisCCodeForArAdj = clsFixedParameter.GetData(clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, trans)
            If clsCommon.myLen(DisCCodeForArAdj) <= 0 Then
                Throw New Exception("Please Map Discount code from Sale setting")
            End If
            DiscDiscForArAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
            GLAcARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Code from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
            GLAcDescARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))

            objVendInvTR.GL_Account_Code = GLAcARAdj
            objVendInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendInvTR.GL_Account_Code, objVendInv.loc_code, True, trans)
            objVendInvTR.GL_Account_Desc = GLAcDescARAdj
            objVendInv.Arr.Add(objVendInvTR)

            objVendInv.SaveData(objVendInv, True, trans)
            clsVedorInvoiceHead.PostData("", objVendInv.Document_No, "", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return isSaved
    End Function
    Public Shared Function CreateARAdjustmentEntry(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction)
        Try
            Dim objRcpt As clsAdjustmentEntryReceivables = Nothing
            Dim DisCCodeForArAdj As String = ""
            Dim GLAcARAdj As String = ""
            Dim DiscDiscForArAdj As String = ""
            Dim GLAcDescARAdj As String = ""
            objRcpt = New clsAdjustmentEntryReceivables
            objRcpt.Adjustment_No = ""
            objRcpt.Description = "AR Adjustment Against MCC Material Sale "
            objRcpt.Adjustment_Date = clsCommon.myCDate(obj.Document_Date)
            objRcpt.Customer_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Cust_Code  from TSPL_CUSTOMER_VENDOR_MAPPING where Vendor_Code ='" & clsCommon.myCstr(obj.Customer_Code) & "' ", trans))
            ''richa BHO/12/07/21-000026
            If clsCommon.myLen(objRcpt.Customer_No) <= 0 Then
                Throw New Exception("Please Map customer code with Vendor code " & obj.Customer_Code & " on Customer Vendor Mapping screen")
            End If

            objRcpt.Customer_Name = clsCommon.myCstr(clsCustomerMaster.GetName(objRcpt.Customer_No, trans))
            objRcpt.Doc_No = clsCommon.myCstr(obj.Sale_Invoice_No)
            Dim ReturnAmt As Decimal = 0

            objRcpt.ARInvoiceNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select document_no from TSPL_Customer_Invoice_Head where against_sale_no='" & obj.Sale_Invoice_No & "'", trans))
            objRcpt.Doc_Amount = clsCommon.myCdbl(obj.Total_Amt)
            objRcpt.Remarks = "AR Adjustment Against MCC Material Sale - " & obj.Sale_Invoice_No & ""
            objRcpt.Adjustment_Amount = clsCommon.myCdbl(obj.Total_Amt)
            objRcpt.Arr = New List(Of clsAdjustmentEntryReceivablesDetail)
            Dim objTrRcpt As New clsAdjustmentEntryReceivablesDetail()
            DisCCodeForArAdj = clsFixedParameter.GetData(clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, trans)
            If clsCommon.myLen(DisCCodeForArAdj) <= 0 Then
                Throw New Exception("Please Map Discount code from Sale setting")
            End If
            DiscDiscForArAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
            GLAcARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Code from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
            GLAcDescARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))

            objTrRcpt.Discount_Code = DisCCodeForArAdj
            objTrRcpt.Discount_Description = DiscDiscForArAdj
            If clsCommon.myLen(GLAcARAdj) <= 0 Then
                Throw New Exception("Please set account for Discount master:" + DisCCodeForArAdj)
            End If

            objTrRcpt.Account_No = GLAcARAdj
            objTrRcpt.Account_Description = GLAcDescARAdj
            objTrRcpt.Amount = clsCommon.myCdbl(obj.Total_Amt)
            objTrRcpt.Remarks = ""
            objRcpt.Arr.Add(objTrRcpt)
            If clsCommon.myCdbl(objTrRcpt.Amount) > 0 Then
                objRcpt.SaveData(objRcpt, True, trans)
                clsAdjustmentEntryReceivables.FunPost(objRcpt.Adjustment_No, trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function CreateARAdjustmentEntryAgainstSubsidy(ByVal obj As clsMCCMaterialSale, ByVal trans As SqlTransaction)
        Try
            Dim objRcpt As clsAdjustmentEntryReceivables = Nothing
            Dim GLAcARAdj As String = ""
            Dim GLAcDescARAdj As String = ""
            objRcpt = New clsAdjustmentEntryReceivables
            objRcpt.Adjustment_No = ""
            objRcpt.Description = "AR Adjustment Against Subsidy"
            objRcpt.Adjustment_Date = clsCommon.myCDate(obj.Document_Date)
            objRcpt.Customer_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Cust_Code  from TSPL_CUSTOMER_VENDOR_MAPPING where Vendor_Code ='" & clsCommon.myCstr(obj.Customer_Code) & "' ", trans))
            If clsCommon.myLen(objRcpt.Customer_No) <= 0 Then
                Throw New Exception("Please Map customer code with Vendor code " & obj.Customer_Code & " on Customer Vendor Mapping screen")
            End If

            objRcpt.Customer_Name = clsCommon.myCstr(clsCustomerMaster.GetName(objRcpt.Customer_No, trans))
            objRcpt.Doc_No = clsCommon.myCstr(obj.Sale_Invoice_No)
            Dim ReturnAmt As Decimal = 0

            objRcpt.ARInvoiceNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select document_no from TSPL_Customer_Invoice_Head where against_sale_no='" & obj.Sale_Invoice_No & "'", trans))
            objRcpt.Doc_Amount = clsCommon.myCdbl(obj.Total_Amt)
            objRcpt.Remarks = "AR Adjustment Against Subsidy"
            objRcpt.Adjustment_Amount = clsCommon.myCdbl(obj.TotalSubsidyAmt)
            objRcpt.Arr = New List(Of clsAdjustmentEntryReceivablesDetail)
            Dim objTrRcpt As New clsAdjustmentEntryReceivablesDetail()

            GLAcARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select isnull(Rate_Difference,'')Rate_Difference from TSPL_CUSTOMER_ACCOUNT_SET LEFT OUTER JOIN tspl_customer_master ON tspl_customer_master.Cust_Account = TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where tspl_customer_master.Cust_Code='" & objRcpt.Customer_No & "'", trans))
            GLAcDescARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select description from tspl_gl_accounts where account_code='" & GLAcARAdj & "'", trans))

            objTrRcpt.Discount_Code = ""
            objTrRcpt.Discount_Description = ""
            If clsCommon.myLen(GLAcARAdj) <= 0 Then
                Throw New Exception("Please set Rate Difference/Subsidy account for [" + objRcpt.Customer_No + "] customer in Customer Account Set :" + GLAcARAdj)
            End If

            objTrRcpt.Account_No = GLAcARAdj
            objTrRcpt.Account_Description = GLAcDescARAdj
            objTrRcpt.Amount = obj.TotalSubsidyAmt
            objTrRcpt.Remarks = ""
            objRcpt.Arr.Add(objTrRcpt)
            If clsCommon.myCdbl(objTrRcpt.Amount) > 0 Then
                objRcpt.SaveData(objRcpt, True, trans)
                clsAdjustmentEntryReceivables.FunPost(objRcpt.Adjustment_No, trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Function RecieptEntryOfDCSSale(ByVal objDCSSale As clsMCCMaterialSale, ByVal trans As SqlTransaction) As String
        Try
            Dim dblReceiptAmt As Double = (objDCSSale.Total_Amt - objDCSSale.TotalSubsidyAmt)
            If dblReceiptAmt > 0 Then
                Dim obj As New clsRcptEntryHeader()
                obj.Entry_Desc = "Shipment No - " + objDCSSale.Document_Code + " for Cash Memo No - " & objDCSSale.Document_Code & "  "
                obj.Receipt_Date = clsCommon.GetPrintDate(objDCSSale.Document_Date, "dd/MMM/yyyy")
                obj.Receipt_Post_Date = clsCommon.GetPrintDate(objDCSSale.Document_Date, "dd/MMM/yyyy")
                obj.Bank_Code = clsCommon.myCstr(objDCSSale.Bank_Code)

                obj.Receipt_Type = "R"
                obj.Payment_Code = clsCommon.myCstr(objDCSSale.Payment_Terms)
                obj.Cheque_No = ""
                obj.Cheque_Date = Nothing
                obj.Cust_Code = clsCommon.myCstr(objDCSSale.Customer_Code)
                obj.Receipt_Amount = dblReceiptAmt
                obj.Balance_Amt = dblReceiptAmt
                obj.UnApply_Amt = dblReceiptAmt
                obj.Apply_By = ""
                obj.Apply_To = ""
                obj.IsSalesmanType = "N"
                obj.Salesman_Code = ""
                obj.Salesman_Name = ""
                obj.SecurityDeposit = "N"

                obj.ArrTr = New List(Of clsReceiptDettail)

                Dim objTr As New clsReceiptDettail()
                objTr.Apply = "Y"
                objTr.Receipt_Type = "I"
                objTr.TagType = "C"
                objTr.Document_No = objDCSSale.Document_Code
                If clsCommon.myLen(objDCSSale.Document_Code) > 0 Then
                    objTr.Document_Date = objDCSSale.Document_Date
                End If
                objTr.Original_Amt = dblReceiptAmt
                objTr.Pending_Balance = dblReceiptAmt
                objTr.Applied_Amount = dblReceiptAmt
                objTr.Adjustment_No = ""
                objTr.Comment = objDCSSale.Remarks
                obj.ArrTr.Add(objTr)

                If obj.SaveData(obj, True, trans) Then
                    clsRcptEntryHeader.funRcptPost(obj.Receipt_No, trans, "MReceivable", objDCSSale.Bill_To_Location, True)
                    'clsCommon.MyMessageBoxShow("Receipt No: '" + obj.Receipt_No + "' create aganist DCS Sale no: '" + objDCSSale.Document_Code + "'")
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_SD_SHIPMENT_HEAD set Receipt_No ='" & obj.Receipt_No & "' where Document_Code='" + objDCSSale.Document_Code + "'", trans)
                End If
                '---------------------------------------------------------------------------
                Return obj.Receipt_No
            Else
                Throw New Exception("Receipt can't be created.Invoice amount is zero.")

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function

    Public Shared Function CreateJournalEntry(ByVal strCode As String, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing)
        Dim RecoControlACC As String = ""
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 0 Then
            RecoControlACC = "I"
        End If
        Dim isSkipCogsGL As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, trans)) = 0, False, True)

        Dim obj As New clsMCCMaterialSale
        obj = clsMCCMaterialSale.GetData(strCode, NavigatorType.Current, trans)
        Dim ArryLstGLAC As ArrayList = New ArrayList()
        Dim strInventoryControlAc As String = ""
        Dim strShipmentClearingAC As String = ""
        Dim dblTotalCost As Double = 0
        If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL
            strShipmentClearingAC = clsDBFuncationality.getSingleValue("SELECT PA.Shipment_Clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
              " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
               " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr.Item(0).Item_Code.ToString() + "'", trans)
            strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, obj.Bill_To_Location, trans)

            If clsCommon.myLen(strShipmentClearingAC) = 0 Then
                Throw New Exception("Please set Shipment clearing Account for first item")
            End If

            Dim dblCogsCost As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end) as COst from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & obj.Document_Code & "'", trans))
            '======update by preeti gupta Against ticket no[TEC/07/03/19-000435]
            Dim Acc() As String = {strShipmentClearingAC, dblCogsCost, "", "", "", "", "", "", "H"}
            ArryLstGLAC.Add(Acc)

            Dim strSql As String = "select TSPL_INVENTORY_MOVEMENT.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  where Source_Doc_No='" & obj.Document_Code & "'"
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
                    Dim Acc1() As String = {strInventoryControlAc, -1 * clsCommon.myCdbl(dr("Cost")), "", "", "", "", "", "", RecoControlACC}
                    ArryLstGLAC.Add(Acc1)
                    If clsCommon.CompairString(RecoControlACC, "I") = CompairStringResult.Equal Then
                        ''TEC/13/03/19-000447 by Richa on 14/03/2019 send Inventory control a/c into inventory movement table
                        clsInventoryMovement.UpdateInvControlAccount(clsCommon.myCstr(strCode), "MCC-MSALE", clsCommon.myCstr(dr("Item_Code")), "", strInventoryControlAc, "", trans)
                        ''------------------
                    End If
                Next
            End If
            '' BHA/30/10/18-000646 RICHA AGARWAL SEND CUSTOMER CODE AND CUSTOMER NAME INTO JOURNAL ENTRY AND TYPE C instead of O 30 Oct,2018
            If strVoucherNoForRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreatedOnly) > 0 Then
                clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, strVoucherNoForRecreatedOnly, trans, obj.Document_Date, obj.Remarks, "SD-SH", "Shipment", obj.Document_Code, "", "C", obj.Customer_Code, clsCustomerMaster.GetName(obj.Customer_Code, trans), objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstGLAC, , obj.Description, obj.Remarks)
            Else
                clsJournalMaster.FunGrnlEntryWithTrans(obj.Bill_To_Location, False, trans, obj.Document_Date, obj.Remarks, "SD-SH", "Shipment", obj.Document_Code, "", "C", obj.Customer_Code, clsCustomerMaster.GetName(obj.Customer_Code, trans), objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstGLAC, , obj.Description, obj.Remarks)
            End If

        End If  '' Done By Pankaj Jha For Skipping Cogs GL
        Return obj


    End Function

    Private Shared Function ConvertShipmentToSaleInvoice(ByVal objShipment As clsMCCMaterialSale) As clsPSInvoiceHead
        Dim obj As New clsPSInvoiceHead()
        obj = New clsPSInvoiceHead()
        obj.Deduction_Type = objShipment.Deduction_Type
        obj.Deduction = objShipment.Deduction
        obj.Item_Tax_Type = objShipment.isTaxExempted
        obj.podate = objShipment.Document_Date
        obj.Total_Comm_Amt = objShipment.Total_Comm_Amt
        obj.RoundOffAmount = objShipment.RoundOffAmount
        obj.Invoice_Type = objShipment.Invoice_Type
        obj.Document_Code = objShipment.Sale_Invoice_No
        obj.Document_Date = objShipment.Document_Date
        obj.Customer_Code = objShipment.Customer_Code
        obj.Customer_Name = objShipment.Customer_Name
        obj.Status = IIf(objShipment.Status = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
        obj.On_Hold = IIf(objShipment.On_Hold = 1, True, False)
        obj.Is_Internal = IIf(objShipment.Is_Internal = 1, True, False)
        obj.Ref_No = objShipment.Ref_No
        obj.No_Transporter = objShipment.No_Transporter
        obj.Description = objShipment.Description
        obj.Remarks = objShipment.Remarks
        obj.Bill_To_Location = objShipment.Bill_To_Location
        obj.Sub_Location_code = objShipment.Sub_Location_code
        obj.Ship_To_Location = objShipment.Ship_To_Location
        obj.Transport_Code = objShipment.Transport_Id
        obj.Transporter_Name = objShipment.Transporter_Name
        obj.Tax_Group = objShipment.Tax_Group
        obj.TAX1 = objShipment.TAX1
        obj.TAX1_Rate = objShipment.TAX1_Rate
        obj.TAX1_Base_Amt = objShipment.TAX1_Base_Amt
        obj.TAX1_Amt = objShipment.TAX1_Amt
        obj.TAX2 = objShipment.TAX2
        obj.TAX2_Rate = objShipment.TAX2_Rate
        obj.TAX2_Base_Amt = objShipment.TAX2_Base_Amt
        obj.TAX2_Amt = objShipment.TAX2_Amt
        obj.TAX3 = objShipment.TAX3
        obj.TAX3_Base_Amt = objShipment.TAX3_Base_Amt
        obj.TAX3_Rate = objShipment.TAX3_Rate
        obj.TAX3_Amt = objShipment.TAX3_Amt
        obj.TAX4 = objShipment.TAX4
        obj.TAX4_Rate = objShipment.TAX4_Rate
        obj.TAX4_Base_Amt = objShipment.TAX4_Base_Amt
        obj.TAX4_Amt = objShipment.TAX4_Amt
        obj.TAX5 = objShipment.TAX5
        obj.TAX5_Rate = objShipment.TAX5_Rate
        obj.TAX5_Base_Amt = objShipment.TAX5_Base_Amt
        obj.TAX5_Amt = objShipment.TAX5_Amt
        obj.TAX6 = objShipment.TAX6
        obj.TAX6_Rate = objShipment.TAX6_Rate
        obj.TAX6_Base_Amt = objShipment.TAX6_Base_Amt
        obj.TAX6_Amt = objShipment.TAX6_Amt
        obj.TAX7 = objShipment.TAX7
        obj.TAX7_Rate = objShipment.TAX7_Rate
        obj.TAX7_Base_Amt = objShipment.TAX7_Base_Amt
        obj.TAX7_Amt = objShipment.TAX7_Amt
        obj.TAX8 = objShipment.TAX8
        obj.TAX8_Rate = objShipment.TAX8_Rate
        obj.TAX8_Base_Amt = objShipment.TAX8_Base_Amt
        obj.TAX8_Amt = objShipment.TAX8_Amt
        obj.TAX9 = objShipment.TAX9
        obj.TAX9_Rate = objShipment.TAX9_Rate
        obj.TAX9_Base_Amt = objShipment.TAX9_Base_Amt
        obj.TAX9_Amt = objShipment.TAX9_Amt
        obj.TAX10 = objShipment.TAX10
        obj.TAX10_Rate = objShipment.TAX10_Rate
        obj.TAX10_Base_Amt = objShipment.TAX10_Base_Amt
        obj.TAX10_Amt = objShipment.TAX10_Amt
        obj.Total_Tax_Amt = objShipment.Total_Tax_Amt
        obj.Discount_Base = objShipment.Discount_Base
        obj.Discount_Amt = objShipment.Discount_Amt
        obj.Amount_Less_Discount = objShipment.Amount_Less_Discount
        obj.Total_Amt = objShipment.Total_Amt
        obj.Comments = objShipment.Comments
        obj.Comp_Code = objShipment.Comp_Code
        obj.Terms_Code = objShipment.Terms_Code
        obj.Due_Date = objShipment.Due_Date
        obj.BillToLocationName = objShipment.BillToLocationName
        obj.ShipToLocationName = objShipment.ShipToLocationName
        obj.TaxGroupName = objShipment.TaxGroupName
        obj.TermsName = objShipment.TermsName
        obj.PROJECT_ID = objShipment.PROJECT_ID
        obj.Route_No = objShipment.Route_No
        obj.Route_Desc = objShipment.Route_Desc
        obj.Price_Code = objShipment.Price_Code
        obj.HeadDisc_Per = objShipment.HeadDisc_Per
        obj.HeadDisc_Amt = objShipment.HeadDisc_Amt
        obj.HeadDisc_PerAmt = objShipment.HeadDisc_PerAmt
        obj.RateDiff_Per = objShipment.RateDiff_Per
        obj.RateDiff_Amt = objShipment.RateDiff_Amt
        obj.Gross_Amount = objShipment.Gross_Amount
        obj.TotalSubsidyAmt = objShipment.TotalSubsidyAmt
        obj.TotalSubsidyDisAmt = objShipment.TotalSubsidyDisAmt
        obj.TotCashDiscAmt = objShipment.TotCashDiscAmt
        obj.Cust_PO_No = objShipment.Cust_PO_No
        ' obj.podate = objShipment.Podate
        obj.No_Of_Instalment = objShipment.No_Of_Instalment


        If objShipment.Posting_Date IsNot Nothing Then
            obj.Posting_Date = objShipment.Posting_Date
        End If

        obj.Salesman_Code = objShipment.Salesman_Code
        obj.Salesman_Name = objShipment.Salesman_Name

        obj.Challan_No = objShipment.Challan_No
        obj.Carrier = objShipment.Carrier
        obj.Vehicle_Code = objShipment.Vehicle_Code
        obj.VehicleNo = objShipment.VehicleNo
        obj.GRNo = objShipment.GRNo
        obj.GENo = objShipment.GENo
        If objShipment.GEDate IsNot Nothing Then
            obj.GEDate = objShipment.GEDate
        End If




        obj.Dept = objShipment.Dept
        obj.Dept_Desc = objShipment.Dept_Desc
        obj.Item_Type = objShipment.Item_Type

        obj.Against_Shipment_No = objShipment.Document_Code


        obj.Add_Charge_Code1 = objShipment.Add_Charge_Code1
        obj.Add_Charge_Name1 = objShipment.Add_Charge_Name1
        obj.Add_Charge_Amt1 = objShipment.Add_Charge_Amt1

        obj.Add_Charge_Code2 = objShipment.Add_Charge_Code2
        obj.Add_Charge_Name2 = objShipment.Add_Charge_Name2
        obj.Add_Charge_Amt2 = objShipment.Add_Charge_Amt2

        obj.Add_Charge_Code3 = objShipment.Add_Charge_Code3
        obj.Add_Charge_Name3 = objShipment.Add_Charge_Name3
        obj.Add_Charge_Amt3 = objShipment.Add_Charge_Amt3

        obj.Add_Charge_Code4 = objShipment.Add_Charge_Code4
        obj.Add_Charge_Name4 = objShipment.Add_Charge_Name4
        obj.Add_Charge_Amt4 = objShipment.Add_Charge_Amt4

        obj.Add_Charge_Code5 = objShipment.Add_Charge_Code5
        obj.Add_Charge_Name5 = objShipment.Add_Charge_Name5
        obj.Add_Charge_Amt5 = objShipment.Add_Charge_Amt5

        obj.Add_Charge_Code6 = objShipment.Add_Charge_Code6
        obj.Add_Charge_Name6 = objShipment.Add_Charge_Name6
        obj.Add_Charge_Amt6 = objShipment.Add_Charge_Amt6

        obj.Add_Charge_Code7 = objShipment.Add_Charge_Code7
        obj.Add_Charge_Name7 = objShipment.Add_Charge_Name7
        obj.Add_Charge_Amt7 = objShipment.Add_Charge_Amt7

        obj.Add_Charge_Code8 = objShipment.Add_Charge_Code8
        obj.Add_Charge_Name8 = objShipment.Add_Charge_Name8
        obj.Add_Charge_Amt8 = objShipment.Add_Charge_Amt8

        obj.Add_Charge_Code9 = objShipment.Add_Charge_Code9
        obj.Add_Charge_Name9 = objShipment.Add_Charge_Name9
        obj.Add_Charge_Amt9 = objShipment.Add_Charge_Amt9

        obj.Add_Charge_Code10 = objShipment.Add_Charge_Code10
        obj.Add_Charge_Name10 = objShipment.Add_Charge_Name10
        obj.Add_Charge_Amt10 = objShipment.Add_Charge_Amt10

        obj.Total_Add_Charge = objShipment.Total_Add_Charge
        obj.Inv_No = objShipment.Inv_No
        If clsCommon.myLen(objShipment.Challan_Date) > 0 Then
            obj.Challan_Date = clsCommon.GetPrintDate(objShipment.Challan_Date, "dd/MMM/yyyy")
        End If

        If clsCommon.myLen(objShipment.Inv_Date) <= 0 Then
            obj.Inv_Date = ""
        Else
            obj.Inv_Date = clsCommon.GetPrintDate(objShipment.Inv_Date, "dd/MMM/yyyy")
        End If
        obj.SO_Validity = objShipment.SO_Validity
        obj.Commission_Apply = objShipment.Commission_Apply
        obj.Dispatch_date = objShipment.Inv_Date 'objShipment.Dispatch_date
        obj.Vehicle_Capacity = objShipment.Vehicle_Capacity
        obj.Dispatch_Terms = objShipment.Dispatch_Terms
        obj.Payment_Terms = objShipment.Payment_Terms
        obj.Dispatch_Period = objShipment.Dispatch_Period
        obj.WayBillNo = objShipment.WayBillNo
        obj.WayBillDate = objShipment.WayBillDate
        obj.Tax_Calculation_Type = IIf(objShipment.Tax_Calculation_Type = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)
        obj.Is_Create_Auto_Receipt = objShipment.Is_Create_Auto_Receipt
        '-----------------richa 27/06/2014 Ticket No .BM00000002982----------
        obj.Mannual_Document_Code = objShipment.Mannual_Invoice_No
        obj.InvoiceManualNowithPrefix = objShipment.InvoiceManualNowithPrefix
        obj.Is_Taxable = IIf(objShipment.Is_Taxable, 1, 0)
        obj.Transporter_Commission_TotalAmt = objShipment.Transporter_Commission_TotalAmt
        obj.Is_Apply_TPT = objShipment.Is_Apply_TPT
        obj.Is_Add_TPT = objShipment.Is_Add_TPT
        obj.Recommended_By = objShipment.Recommended_By
        obj.TPT_Vendor = objShipment.TPT_Vendor
        obj.Bank_Code = objShipment.Bank_Code
        obj.Exclude_KKF_And_Mandi = objShipment.Exclude_KKF_And_Mandi
        'obj.Invoice_Type
        '-------------------------------------------------------------------
        If (objShipment.Arr IsNot Nothing AndAlso objShipment.Arr.Count > 0) Then
            obj.Arr = New List(Of clsPSInvoiceHeadDetail)
            Dim objTr As clsPSInvoiceHeadDetail
            For Each objShipmentDetail As clsMCCMaterialSaleDetail In objShipment.Arr
                objTr = New clsPSInvoiceHeadDetail
                objTr.Transporter_Commission_Rate = objShipmentDetail.Transporter_Commission_Rate
                objTr.Transporter_Commission_Amt = objShipmentDetail.Transporter_Commission_Amt
                objTr.REF_TPT_PK_ID = objShipmentDetail.REF_TPT_PK_ID
                objTr.PrincipleCode = objShipmentDetail.PrincipleCode
                objTr.PrincipleDesc = objShipmentDetail.PrincipleDesc
                objTr.vendor_code = objShipmentDetail.vendor_code
                objTr.vendor_desc = objShipmentDetail.vendor_desc
                objTr.Document_Code = objShipmentDetail.Document_Code
                objTr.Row_Type = objShipmentDetail.Row_Type
                objTr.Line_No = objShipmentDetail.Line_No
                objTr.Status = Convert.ToInt32(objShipmentDetail.Status)
                objTr.Item_Code = objShipmentDetail.Item_Code
                objTr.Item_Desc = objShipmentDetail.Item_Desc
                objTr.Qty = objShipmentDetail.Qty
                objTr.Billing_Qty = objShipmentDetail.Billing_Qty
                objTr.Free_Qty = objShipmentDetail.Free_Qty
                objTr.Shipment_Code = objShipment.Document_Code
                objTr.Balance_Qty = objShipmentDetail.Balance_Qty
                objTr.Unit_code = objShipmentDetail.Unit_code
                objTr.Billing_Unit_code = objShipmentDetail.Billing_Unit_code
                objTr.Location = objShipmentDetail.Location
                objTr.LocationName = objShipmentDetail.LocationName
                objTr.Item_Cost = objShipmentDetail.Item_Cost
                objTr.TAX1 = objShipmentDetail.TAX1
                objTr.TAX1_Base_Amt = objShipmentDetail.TAX1_Base_Amt
                objTr.TAX1_Rate = objShipmentDetail.TAX1_Rate
                objTr.TAX1_Amt = objShipmentDetail.TAX1_Amt
                objTr.TAX2 = objShipmentDetail.TAX2
                objTr.TAX2_Base_Amt = objShipmentDetail.TAX2_Base_Amt
                objTr.TAX2_Rate = objShipmentDetail.TAX2_Rate
                objTr.TAX2_Amt = objShipmentDetail.TAX2_Amt
                objTr.TAX3 = objShipmentDetail.TAX3
                objTr.TAX3_Base_Amt = objShipmentDetail.TAX3_Base_Amt
                objTr.TAX3_Rate = objShipmentDetail.TAX3_Rate
                objTr.TAX3_Amt = objShipmentDetail.TAX3_Amt
                objTr.TAX4 = objShipmentDetail.TAX4
                objTr.TAX4_Base_Amt = objShipmentDetail.TAX4_Base_Amt
                objTr.TAX4_Rate = objShipmentDetail.TAX4_Rate
                objTr.TAX4_Amt = objShipmentDetail.TAX4_Amt
                objTr.TAX5 = objShipmentDetail.TAX5
                objTr.TAX5_Base_Amt = objShipmentDetail.TAX5_Base_Amt
                objTr.TAX5_Rate = objShipmentDetail.TAX5_Rate
                objTr.TAX5_Amt = objShipmentDetail.TAX5_Amt
                objTr.TAX6 = objShipmentDetail.TAX6
                objTr.TAX6_Base_Amt = objShipmentDetail.TAX6_Base_Amt
                objTr.TAX6_Rate = objShipmentDetail.TAX6_Rate
                objTr.TAX6_Amt = objShipmentDetail.TAX6_Amt
                objTr.TAX7 = objShipmentDetail.TAX7
                objTr.TAX7_Base_Amt = objShipmentDetail.TAX7_Base_Amt
                objTr.TAX7_Rate = objShipmentDetail.TAX7_Rate
                objTr.TAX7_Amt = objShipmentDetail.TAX7_Amt
                objTr.TAX8 = objShipmentDetail.TAX8
                objTr.TAX8_Base_Amt = objShipmentDetail.TAX8_Base_Amt
                objTr.TAX8_Rate = objShipmentDetail.TAX8_Rate
                objTr.TAX8_Amt = objShipmentDetail.TAX8_Amt
                objTr.TAX9 = objShipmentDetail.TAX9
                objTr.TAX9_Base_Amt = objShipmentDetail.TAX9_Base_Amt
                objTr.TAX9_Rate = objShipmentDetail.TAX9_Rate
                objTr.TAX9_Amt = objShipmentDetail.TAX9_Amt
                objTr.TAX10 = objShipmentDetail.TAX10
                objTr.TAX10_Base_Amt = objShipmentDetail.TAX10_Base_Amt
                objTr.TAX10_Rate = objShipmentDetail.TAX10_Rate
                objTr.TAX10_Amt = objShipmentDetail.TAX10_Amt
                objTr.Amount = objShipmentDetail.Amount
                objTr.Disc_Per = objShipmentDetail.Disc_Per
                objTr.Disc_Amt = objShipmentDetail.Disc_Amt
                objTr.Amt_Less_Discount = objShipmentDetail.Amt_Less_Discount
                objTr.Total_Tax_Amt = objShipmentDetail.Total_Tax_Amt
                objTr.Item_Net_Amt = objShipmentDetail.Item_Net_Amt


                objTr.Is_Mannual_Amt = objShipmentDetail.Is_Mannual_Amt

                objTr.MRP = objShipmentDetail.MRP
                objTr.Assessable = objShipmentDetail.Assessable
                objTr.AssessableAmt = objShipmentDetail.AssessableAmt
                objTr.Batch_No = objShipmentDetail.Batch_No
                If objShipmentDetail.MFG_Date IsNot Nothing Then
                    objTr.MFG_Date = objShipmentDetail.MFG_Date
                End If
                If objShipmentDetail.Expiry_Date IsNot Nothing Then
                    objTr.Expiry_Date = objShipmentDetail.Expiry_Date
                End If
                objTr.Specification = objShipmentDetail.Specification
                objTr.Remarks = objShipmentDetail.Remarks

                objTr.Scheme_Applicable = IIf(objShipmentDetail.Scheme_Applicable = "Y", "Yes", "No")
                objTr.Scheme_Code = objShipmentDetail.Scheme_Code
                objTr.Scheme_Item = IIf(objShipmentDetail.Scheme_Item = "Y", "Yes", "No")
                objTr.Item_Tax = objShipmentDetail.Item_Tax
                objTr.Total_MRP_Amt = objShipmentDetail.Total_MRP_Amt
                objTr.Total_Basic_Amt = objShipmentDetail.Total_Basic_Amt
                objTr.Total_Disc_Amt = objShipmentDetail.Total_Disc_Amt
                objTr.Cust_Discount = objShipmentDetail.Cust_Discount
                objTr.Total_Cust_Discount = objShipmentDetail.Total_Cust_Discount
                objTr.ActualRate = objShipmentDetail.ActualRate
                objTr.Cust_DiscountQty = objShipmentDetail.Cust_DiscountQty
                objTr.Price_code = objShipmentDetail.Price_code
                objTr.Price_Date = objShipmentDetail.Price_Date
                objTr.Abatement_Per = objShipmentDetail.Abatement_Per
                objTr.Abatement_Amt = objShipmentDetail.Abatement_Amt
                objTr.FOC_Item = objShipmentDetail.FOC_Item
                objTr.Markup_On = objShipmentDetail.Markup_On
                objTr.Markup_Percent = objShipmentDetail.Markup_Percent
                objTr.Landing_Cost = objShipmentDetail.Landing_Cost
                objTr.HeadDiscAmt = objShipmentDetail.HeadDiscAmt
                objTr.HeadDiscPer = objShipmentDetail.HeadDiscPer
                objTr.HeadDiscPerAmt = objShipmentDetail.HeadDiscPerAmt
                objTr.CustDiscPer = objShipmentDetail.CustDiscPer
                objTr.CasdDiscScheme_Code = objShipmentDetail.CasdDiscScheme_Code

                objTr.Item_Weight = objShipmentDetail.Item_Weight
                objTr.TotalItem_Weight = objShipmentDetail.TotalItem_Weight
                objTr.Conv_Factor = objShipmentDetail.Conv_Factor
                objTr.Purchase_Cost = objShipmentDetail.Purchase_Cost
                objTr.OrgRate = objShipmentDetail.OrgRate

                objTr.Price_Amount1 = objShipmentDetail.Price_Amount1
                objTr.Price_Amount2 = objShipmentDetail.Price_Amount2
                objTr.Price_Amount3 = objShipmentDetail.Price_Amount3
                objTr.Price_Amount4 = objShipmentDetail.Price_Amount4
                objTr.Price_Amount5 = objShipmentDetail.Price_Amount5
                objTr.Price_Amount6 = objShipmentDetail.Price_Amount6
                objTr.Price_Amount7 = objShipmentDetail.Price_Amount7
                objTr.Price_Amount8 = objShipmentDetail.Price_Amount8
                objTr.Price_Amount9 = objShipmentDetail.Price_Amount9
                objTr.Price_Amount10 = objShipmentDetail.Price_Amount10

                objTr.TAX1_Base_Amt = objShipmentDetail.TAX1_Base_Amt
                objTr.TAX2_Base_Amt = objShipmentDetail.TAX2_Base_Amt
                objTr.TAX3_Base_Amt = objShipmentDetail.TAX3_Base_Amt
                objTr.TAX4_Base_Amt = objShipmentDetail.TAX4_Base_Amt
                objTr.TAX5_Base_Amt = objShipmentDetail.TAX5_Base_Amt
                objTr.TAX6_Base_Amt = objShipmentDetail.TAX6_Base_Amt
                objTr.TAX7_Base_Amt = objShipmentDetail.TAX7_Base_Amt
                objTr.TAX8_Base_Amt = objShipmentDetail.TAX8_Base_Amt
                objTr.TAX9_Base_Amt = objShipmentDetail.TAX9_Base_Amt
                objTr.TAX10_Base_Amt = objShipmentDetail.TAX10_Base_Amt

                objTr.Commission_Rate = objShipmentDetail.Commission_Rate
                objTr.Commission_Party = objShipmentDetail.Commission_Party
                objTr.Commission_Amt = objShipmentDetail.Commission_Amt
                objTr.Amt_Less_Commission = objShipmentDetail.Amt_Less_Commission

                obj.Arr.Add(objTr)
            Next
        End If
        Return obj
    End Function
    Private Shared Function ConvertShipmentToSaleOrder(ByVal objShipment As clsMCCMaterialSale) As clsPSSalesOrder
        Dim obj As New clsPSSalesOrder()
        obj = New clsPSSalesOrder()
        obj.Auto_SaleOrder = 1
        obj.Total_Comm_Amt = objShipment.Total_Comm_Amt

        obj.Cust_PO_No = objShipment.Cust_PO_No
        obj.HeadDisc_Per = objShipment.HeadDisc_Per
        obj.HeadDisc_PerAmt = objShipment.HeadDisc_PerAmt
        obj.HeadDisc_Amt = objShipment.HeadDisc_Amt
        obj.Road_Permit_No = objShipment.Road_Permit_No
        obj.Price_Group_Code = objShipment.Price_Group_Code
        obj.Route_No = objShipment.Route_No
        obj.Route_Desc = objShipment.Route_Desc
        obj.Price_Code = objShipment.Price_Code
        obj.Document_Date = clsCommon.GetPrintDate(objShipment.Document_Date, "dd-MMM-yyyy hh:mm:ss")
        obj.CloseSO = "N"
        obj.Delivery_date = clsCommon.GetPrintDate(objShipment.Document_Date, "dd-MMM-yyyy hh:mm:ss")
        obj.Customer_Code = objShipment.Customer_Code
        obj.Customer_Name = objShipment.Customer_Name
        obj.Ref_No = objShipment.Ref_No
        obj.Total_Tax_Amt = objShipment.Total_Tax_Amt
        obj.Remarks = objShipment.Remarks
        obj.Bill_To_Location = objShipment.Bill_To_Location
        obj.Ship_To_Location = objShipment.Ship_To_Location
        obj.Comments = objShipment.Comments
        obj.On_Hold = objShipment.On_Hold
        obj.Mode_Of_Transport = "By Road"
        obj.Description = objShipment.Description
        obj.Tax_Group = objShipment.Tax_Group
        obj.SalesOrder_Type = ""
        obj.Item_Type = objShipment.Item_Type
        obj.Dept = objShipment.Dept
        obj.Dept_Desc = objShipment.Dept_Desc
        obj.PROJECT_ID = objShipment.PROJECT_ID
        obj.Approvel_Required = 0
        If clsCommon.myLen(objShipment.TAX1) > 0 Then
            obj.TAX1 = objShipment.TAX1
            obj.TAX1_Rate = objShipment.TAX1_Rate
            obj.TAX1_Base_Amt = objShipment.TAX1_Base_Amt
            obj.TAX1_Amt = objShipment.TAX1_Amt
        End If
        If clsCommon.myLen(objShipment.TAX2) > 0 Then
            obj.TAX2 = objShipment.TAX2
            obj.TAX2_Rate = objShipment.TAX2_Rate
            obj.TAX2_Base_Amt = objShipment.TAX2_Base_Amt
            obj.TAX2_Amt = objShipment.TAX2_Amt
        End If
        If clsCommon.myLen(objShipment.TAX3) > 0 Then
            obj.TAX3 = objShipment.TAX3
            obj.TAX3_Rate = objShipment.TAX3_Rate
            obj.TAX3_Base_Amt = objShipment.TAX3_Base_Amt
            obj.TAX3_Amt = objShipment.TAX3_Amt
        End If
        If clsCommon.myLen(objShipment.TAX4) > 0 Then
            obj.TAX4 = objShipment.TAX4
            obj.TAX4_Rate = objShipment.TAX4_Rate
            obj.TAX4_Base_Amt = objShipment.TAX4_Base_Amt
            obj.TAX4_Amt = objShipment.TAX4_Amt
        End If
        If clsCommon.myLen(objShipment.TAX5) > 0 Then
            obj.TAX5 = objShipment.TAX1
            obj.TAX5_Rate = objShipment.TAX5_Rate
            obj.TAX5_Base_Amt = objShipment.TAX5_Base_Amt
            obj.TAX5_Amt = objShipment.TAX5_Amt
        End If
        If clsCommon.myLen(objShipment.TAX6) > 0 Then
            obj.TAX6 = objShipment.TAX6
            obj.TAX6_Rate = objShipment.TAX6_Rate
            obj.TAX6_Base_Amt = objShipment.TAX6_Base_Amt
            obj.TAX6_Amt = objShipment.TAX6_Amt
        End If
        If clsCommon.myLen(objShipment.TAX7) > 0 Then
            obj.TAX7 = objShipment.TAX7
            obj.TAX7_Rate = objShipment.TAX7_Rate
            obj.TAX7_Base_Amt = objShipment.TAX7_Base_Amt
            obj.TAX7_Amt = objShipment.TAX7_Amt
        End If
        If clsCommon.myLen(objShipment.TAX8) > 0 Then
            obj.TAX8 = objShipment.TAX8
            obj.TAX8_Rate = objShipment.TAX8_Rate
            obj.TAX8_Base_Amt = objShipment.TAX8_Base_Amt
            obj.TAX8_Amt = objShipment.TAX8_Amt
        End If
        If clsCommon.myLen(objShipment.TAX9) > 0 Then
            obj.TAX9 = objShipment.TAX9
            obj.TAX9_Rate = objShipment.TAX9_Rate
            obj.TAX9_Base_Amt = objShipment.TAX9_Base_Amt
            obj.TAX9_Amt = objShipment.TAX9_Amt
        End If
        If clsCommon.myLen(objShipment.TAX1) > 0 Then
            obj.TAX10 = objShipment.TAX10
            obj.TAX10_Rate = objShipment.TAX10_Rate
            obj.TAX10_Base_Amt = objShipment.TAX10_Base_Amt
            obj.TAX10_Amt = objShipment.TAX10_Amt
        End If


        obj.Terms_Code = objShipment.Terms_Code
        obj.Due_Date = objShipment.Due_Date
        obj.Discount_Base = objShipment.Discount_Base
        obj.Discount_Amt = objShipment.Discount_Amt
        obj.Amount_Less_Discount = objShipment.Amount_Less_Discount
        obj.Total_Amt = objShipment.Total_Amt
        obj.Abandonment_No = 0
        'obj.Against_Quotation_No = txtReqNo.Value
        obj.Against_DeliveryNo = ""
        obj.SO_Validity = 0
        obj.Commission_Apply = objShipment.Commission_Apply
        obj.Dispatch_date = objShipment.Dispatch_date
        obj.Vehicle_Code = objShipment.Vehicle_Code
        obj.Vehicle_No = objShipment.VehicleNo
        obj.Vehicle_Capacity = objShipment.Vehicle_Capacity
        obj.Payment_Terms = objShipment.Payment_Terms
        obj.Dispatch_Terms = objShipment.Dispatch_Terms
        obj.Dispatch_Period = objShipment.Dispatch_Period
        'If clsCommon.myLen(txtReqNo.Value) = 0 Then
        '    txtDispatchDate.Value = txtDeliveryDate.Value.AddDays(txtDispatchPeriod.Value)
        'End If



        If clsCommon.myLen(obj.Add_Charge_Code1) > 0 Then
            obj.Add_Charge_Code1 = objShipment.Add_Charge_Code1
            obj.Add_Charge_Name1 = objShipment.Add_Charge_Name1
            obj.Add_Charge_Amt1 = objShipment.Add_Charge_Amt1
        End If
        If clsCommon.myLen(obj.Add_Charge_Code2) > 0 Then
            obj.Add_Charge_Code2 = objShipment.Add_Charge_Code2
            obj.Add_Charge_Name2 = objShipment.Add_Charge_Name2
            obj.Add_Charge_Amt2 = objShipment.Add_Charge_Amt2
        End If
        If clsCommon.myLen(obj.Add_Charge_Code3) > 0 Then
            obj.Add_Charge_Code3 = objShipment.Add_Charge_Code3
            obj.Add_Charge_Name3 = objShipment.Add_Charge_Name3
            obj.Add_Charge_Amt3 = objShipment.Add_Charge_Amt3
        End If
        If clsCommon.myLen(obj.Add_Charge_Code4) > 0 Then
            obj.Add_Charge_Code4 = objShipment.Add_Charge_Code4
            obj.Add_Charge_Name4 = objShipment.Add_Charge_Name4
            obj.Add_Charge_Amt4 = objShipment.Add_Charge_Amt4
        End If
        If clsCommon.myLen(obj.Add_Charge_Code5) > 0 Then
            obj.Add_Charge_Code5 = objShipment.Add_Charge_Code5
            obj.Add_Charge_Name1 = objShipment.Add_Charge_Name5
            obj.Add_Charge_Amt5 = objShipment.Add_Charge_Amt5
        End If
        If clsCommon.myLen(obj.Add_Charge_Code6) > 0 Then
            obj.Add_Charge_Code6 = objShipment.Add_Charge_Code6
            obj.Add_Charge_Name6 = objShipment.Add_Charge_Name6
            obj.Add_Charge_Amt6 = objShipment.Add_Charge_Amt6
        End If
        If clsCommon.myLen(obj.Add_Charge_Code7) > 0 Then
            obj.Add_Charge_Code7 = objShipment.Add_Charge_Code7
            obj.Add_Charge_Name7 = objShipment.Add_Charge_Name7
            obj.Add_Charge_Amt7 = objShipment.Add_Charge_Amt7
        End If
        If clsCommon.myLen(obj.Add_Charge_Code8) > 0 Then
            obj.Add_Charge_Code8 = objShipment.Add_Charge_Code8
            obj.Add_Charge_Name8 = objShipment.Add_Charge_Name8
            obj.Add_Charge_Amt8 = objShipment.Add_Charge_Amt8
        End If
        If clsCommon.myLen(obj.Add_Charge_Code9) > 0 Then
            obj.Add_Charge_Code9 = objShipment.Add_Charge_Code9
            obj.Add_Charge_Name9 = objShipment.Add_Charge_Name9
            obj.Add_Charge_Amt9 = objShipment.Add_Charge_Amt9
        End If
        If clsCommon.myLen(obj.Add_Charge_Code10) > 0 Then
            obj.Add_Charge_Code10 = objShipment.Add_Charge_Code10
            obj.Add_Charge_Name10 = objShipment.Add_Charge_Name10
            obj.Add_Charge_Amt10 = objShipment.Add_Charge_Amt10
        End If
        obj.Total_Add_Charge = objShipment.Total_Add_Charge

        obj.Salesman_Code = objShipment.Salesman_Code
        obj.Salesman_Name = objShipment.Salesman_Name

        obj.Arr = New List(Of clsPSSalesOrderDetail)
        For Each objShipmentDetail As clsMCCMaterialSaleDetail In objShipment.Arr

            Dim objTr As New clsPSSalesOrderDetail()
            objTr.Line_No = objShipmentDetail.Line_No
            objTr.Row_Type = objShipmentDetail.Row_Type
            objTr.Item_Code = objShipmentDetail.Item_Code
            objTr.Item_Desc = objShipmentDetail.Item_Desc
            objTr.Qty = objShipmentDetail.Qty
            objTr.Balance_Qty = 0
            objTr.Unit_code = objShipmentDetail.Unit_code
            objTr.OrgUnit_code = objShipmentDetail.OrgUnit_code
            ' objTr.BOOK_QTY_UOM = objShipmentDetail.Unit_code
            'objTr.BOOK_RATE_UOM = objShipmentDetail.Unit_code
            'objTr.Quotation_Code = clsCommon.myCstr(grow.Cells(colReqistionNo).Value)
            objTr.Location = objShipmentDetail.Location
            objTr.Ship_Party = obj.Customer_Code
            objTr.Location = objShipmentDetail.Location
            objTr.Item_Cost = objShipmentDetail.Item_Cost
            objTr.Amount = objShipmentDetail.Amount
            objTr.Disc_Per = objShipmentDetail.Disc_Per
            objTr.Disc_Amt = objShipmentDetail.Disc_Amt
            objTr.Amt_Less_Discount = objShipmentDetail.Amt_Less_Discount
            objTr.TAX1 = objShipmentDetail.TAX1
            objTr.TAX1_Base_Amt = objShipmentDetail.TAX1_Base_Amt
            objTr.TAX1_Rate = objShipmentDetail.TAX1_Rate
            objTr.TAX1_Amt = objShipmentDetail.TAX1_Amt
            objTr.TAX2 = objShipmentDetail.TAX2
            objTr.TAX2_Base_Amt = objShipmentDetail.TAX2_Base_Amt
            objTr.TAX2_Rate = objShipmentDetail.TAX2_Rate
            objTr.TAX2_Amt = objShipmentDetail.TAX2_Amt
            objTr.TAX3 = objShipmentDetail.TAX3
            objTr.TAX3_Base_Amt = objShipmentDetail.TAX3_Base_Amt
            objTr.TAX3_Rate = objShipmentDetail.TAX3_Rate
            objTr.TAX3_Amt = objShipmentDetail.TAX3_Amt
            objTr.TAX4 = objShipmentDetail.TAX4
            objTr.TAX4_Base_Amt = objShipmentDetail.TAX4_Base_Amt
            objTr.TAX4_Rate = objShipmentDetail.TAX4_Rate
            objTr.TAX4_Amt = objShipmentDetail.TAX4_Amt
            objTr.TAX5 = objShipmentDetail.TAX5
            objTr.TAX5_Base_Amt = objShipmentDetail.TAX5_Base_Amt
            objTr.TAX5_Rate = objShipmentDetail.TAX5_Rate
            objTr.TAX5_Amt = objShipmentDetail.TAX5_Amt
            objTr.TAX6 = objShipmentDetail.TAX6
            objTr.TAX6_Base_Amt = objShipmentDetail.TAX6_Base_Amt
            objTr.TAX6_Rate = objShipmentDetail.TAX6_Rate
            objTr.TAX6_Amt = objShipmentDetail.TAX6_Amt
            objTr.TAX7 = objShipmentDetail.TAX7
            objTr.TAX7_Base_Amt = objShipmentDetail.TAX7_Base_Amt
            objTr.TAX7_Rate = objShipmentDetail.TAX7_Rate
            objTr.TAX7_Amt = objShipmentDetail.TAX7_Amt
            objTr.TAX8 = objShipmentDetail.TAX8
            objTr.TAX8_Base_Amt = objShipmentDetail.TAX8_Base_Amt
            objTr.TAX8_Rate = objShipmentDetail.TAX8_Rate
            objTr.TAX8_Amt = objShipmentDetail.TAX8_Amt
            objTr.TAX9 = objShipmentDetail.TAX9
            objTr.TAX9_Base_Amt = objShipmentDetail.TAX9_Base_Amt
            objTr.TAX9_Rate = objShipmentDetail.TAX9_Rate
            objTr.TAX9_Amt = objShipmentDetail.TAX9_Amt
            objTr.TAX10 = objShipmentDetail.TAX10
            objTr.TAX10_Base_Amt = objShipmentDetail.TAX10_Base_Amt
            objTr.TAX10_Rate = objShipmentDetail.TAX10_Rate
            objTr.TAX10_Amt = objShipmentDetail.TAX10_Amt
            objTr.Total_Tax_Amt = objShipmentDetail.Total_Tax_Amt
            objTr.Item_Net_Amt = objShipmentDetail.Item_Net_Amt
            objTr.Specification = objShipmentDetail.Specification
            objTr.Remarks = objShipmentDetail.Remarks
            'objTr.Location = txtBillToLocation.Value 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)
            objTr.MRP = objShipmentDetail.MRP
            objTr.Scheme_Applicable = objShipmentDetail.Scheme_Applicable
            objTr.Scheme_Code = objShipmentDetail.Scheme_Code
            objTr.Scheme_Item = objShipmentDetail.Scheme_Item
            objTr.Item_Tax = objShipmentDetail.Item_Tax
            objTr.Total_MRP_Amt = objShipmentDetail.Total_MRP_Amt
            objTr.Total_Basic_Amt = objShipmentDetail.Total_Basic_Amt
            objTr.Total_Disc_Amt = objShipmentDetail.Total_Disc_Amt
            objTr.Cust_Discount = objShipmentDetail.Cust_Discount
            objTr.Total_Cust_Discount = objShipmentDetail.Total_Cust_Discount
            objTr.ActualRate = objShipmentDetail.ActualRate
            objTr.Cust_DiscountQty = objShipmentDetail.Cust_DiscountQty
            objTr.Price_Date = objShipmentDetail.Price_Date
            objTr.Price_code = objShipmentDetail.Price_code
            objTr.Abatement_Per = objShipmentDetail.Abatement_Per
            objTr.Abatement_Amt = objShipmentDetail.Abatement_Amt
            objTr.FOC_Item = objShipmentDetail.FOC_Item

            objTr.Item_Weight = objShipmentDetail.Item_Weight
            objTr.Conv_Factor = objShipmentDetail.Conv_Factor
            objTr.TotalItem_Weight = objShipmentDetail.TotalItem_Weight
            objTr.Batch_No = objShipmentDetail.Batch_No
            objTr.Bin_No = objShipmentDetail.Bin_No
            objTr.HeadDiscPer = objShipmentDetail.HeadDiscPer
            objTr.HeadDiscPerAmt = objShipmentDetail.HeadDiscPerAmt
            objTr.Expiry_Date = objShipmentDetail.Expiry_Date
            objTr.MFG_Date = objShipmentDetail.MFG_Date
            objTr.Markup_On = objShipmentDetail.Markup_On
            objTr.Markup_Percent = objShipmentDetail.Markup_Percent
            objTr.Landing_Cost = objShipmentDetail.Landing_Cost
            objTr.CustDiscPer = objShipmentDetail.CustDiscPer
            objTr.HeadDiscAmt = objShipmentDetail.HeadDiscAmt
            objTr.CasdDiscScheme_Code = objShipmentDetail.CasdDiscScheme_Code
            objTr.Purchase_Cost = objShipmentDetail.Purchase_Cost
            objTr.OrgRate = objShipmentDetail.OrgRate
            objTr.PrincipleCode = objShipmentDetail.PrincipleCode
            objTr.PrincipleDesc = objShipmentDetail.PrincipleDesc
            objTr.vendor_code = objShipmentDetail.vendor_code
            objTr.vendor_desc = objShipmentDetail.vendor_desc

            objTr.HeadDiscPer = objShipmentDetail.HeadDiscPer
            objTr.HeadDiscPerAmt = objShipmentDetail.HeadDiscPerAmt

            objTr.Commission_Rate = objShipmentDetail.Commission_Rate
            objTr.Commission_Party = objShipmentDetail.Commission_Party
            objTr.Commission_Amt = objShipmentDetail.Commission_Amt
            objTr.Amt_Less_Commission = objShipmentDetail.Amt_Less_Commission
            objTr.Ship_Party = obj.Customer_Code
            objTr.delivery_code = ""
            ''objTr.Assessable = clsCommon.myCdbl(grow.Cells(colAssessableRate).Value)
            ''objTr.AssessableAmt = clsCommon.myCdbl(grow.Cells(colAssessableAmount).Value)
            If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                obj.Arr.Add(objTr)
            End If
        Next

        Return obj
    End Function

    Private Shared Function GetFirstItemCode(ByVal Arr As List(Of clsMCCMaterialSaleDetail)) As String


        For Each objtr As clsMCCMaterialSaleDetail In Arr
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
        Dim obj As clsMCCMaterialSale = clsMCCMaterialSale.GetData(strCode, NavigatorType.Current, trans)

        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
            Try
                '' Anubhooti 06-Sep-2014 BM00000003735 (Locked Transaction)
                'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Sales And Distribution", "Shipment/Sale Invoice", obj.Bill_To_Location, obj.Document_Date, trans)
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmMCCMaterial, obj.Bill_To_Location, obj.Document_Date, trans)

                ''
                If (obj.Status = 1) Then
                    Throw New Exception("Already Posted on :" + obj.Posting_Date)
                End If

                clsSerializeInvenotry.DeleteData("SD-IN", strCode, trans)
                clsBatchInventory.DeleteData("MCC-MSALE", obj.Document_Code, trans)
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", "TSPL_SD_SHIPMENT_DETAIL", "Document_Code", trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", "TSPL_SD_SHIPMENT_DETAIL", "Document_Code", trans)
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", "TSPL_SD_SHIPMENT_DETAIL", "Document_Code", trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SALE_INVOICE_HEAD", "Against_Shipment_No", trans)
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SALE_INVOICE_HEAD", "Against_Shipment_No", trans)

                Dim qry As String = "delete from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SD_SALE_INVOICE_Detail where DOCUMENT_CODE in (select DOCUMENT_CODE from TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No='" + strCode + "')"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                clsCustomFieldValues.DeleteData(obj.Form_ID, strCode, trans)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
        Return True
    End Function

    Public Shared Function IsValidCustomer(ByVal Arr As List(Of String), ByVal strCustomerCode As String) As Boolean
        If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
            Dim qry As String = "select TSPL_SD_SHIPMENT_HEAD.Document_Code,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name from TSPL_SD_SHIPMENT_HEAD left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code where Document_Code in (" + clsCommon.GetMulcallString(Arr) + ") and Customer_Code not in ('" + strCustomerCode + "')"
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
    '         "select TSPL_SD_SHIPMENT_HEAD.Document_Code,TSPL_SD_SHIPMENT_HEAD .Item_Type as ItemType," & _
    '         "(replace( CONVERT(varchar(11), TSPL_SD_SHIPMENT_HEAD.Document_Date,104),'.','/')+' '+CONVERT(varchar(100),TSPL_SD_SHIPMENT_HEAD.Document_Date,108) )as MRN_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Name,TSPL_SD_SHIPMENT_HEAD.GRNo,TSPL_SD_SHIPMENT_HEAD.GENo," & _
    '         "(case when LEN(TSPL_SD_SHIPMENT_HEAD.GEDate)>0  then REPLACE( CONVERT(varchar(11), TSPL_SD_SHIPMENT_HEAD.GEDate,104),'.','/') else '' end) as GEDate,TSPL_SD_SHIPMENT_HEAD.VehicleNo,TSPL_SD_SHIPMENT_HEAD.Remarks ,TSPL_SD_SHIPMENT_HEAD.Ref_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Item_Desc,TSPL_SD_SHIPMENT_DETAIL.Unit_code," & _
    '         "case when Unit_code='FC' then Qty + ISNULL( Free_Qty,0) end as FCS, " & _
    '         "case when Unit_code='FB' then Qty + ISNULL( Free_Qty,0) end as FBS, " & _
    '         "case when Unit_code='SH' then Qty + ISNULL( Free_Qty,0) end as FSH, " & _
    '         "case when Unit_code='EC' then Qty + ISNULL( Free_Qty,0) end as ECS," & _
    '         "case when Unit_code='EB' then Qty + ISNULL( Free_Qty,0) end as EBS, " & _
    '         "TSPL_SD_SHIPMENT_DETAIL.Leak_Qty,TSPL_SD_SHIPMENT_DETAIL.Burst_Qty,TSPL_SD_SHIPMENT_DETAIL.Short_Qty from TSPL_SD_SHIPMENT_DETAIL left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code= TSPL_SD_SHIPMENT_DETAIL.Document_Code " & _
    '         " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_SD_SHIPMENT_HEAD.Bill_To_Location   where Item_Type ='F'"
    '                If FromDate.HasValue AndAlso ToDate.HasValue Then
    '                    qry += " and Convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103)<=Convert(date,'" + ToDate + "',103) "
    '                End If

    '                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
    '                    qry += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
    '                End If
    '                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
    '                    qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
    '                End If
    '                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
    '                    qry += " and TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")" 'ADDED BY ABHISHEK AS ON 30 AUG 2012
    '                End If
    '                qry += " )xxx group by Document_Code,Item_Code order by Item_Desc"
    '                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    '                    common.clsCommon.MyMessageBoxShow("No Record Found")
    '                Else
    '                    PurchaseOrderViewer.funreport(dt, EnumTecxpertPaperSize.PaperSize10x6, "rptSRNCustomReport", "SRN Report")

    '                End If
    '            Else ''For RM Other Print out
    '                Dim strquery As String = "SELECT TSPL_SD_SHIPMENT_HEAD.Document_Code, TSPL_SD_SHIPMENT_HEAD.Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Name,(case when len(against_mrn)>0 then (select MRN_Date  from tspl_mrn_head where tspl_mrn_head.MRN_No =against_mrn) else Document_Date end ) as Challan_Date, TSPL_SD_SHIPMENT_HEAD.Ref_No  " & _
    '                      "as Challan_No, TSPL_SD_SHIPMENT_HEAD.Inv_No, TSPL_SD_SHIPMENT_HEAD.Inv_Date, TSPL_SD_SHIPMENT_HEAD.GRNo,TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount ,TSPL_SD_SHIPMENT_HEAD.GENo,TSPL_SD_SHIPMENT_HEAD.Total_Amt, " & _
    '                      "TSPL_SD_SHIPMENT_HEAD.GEDate, TSPL_SD_SHIPMENT_HEAD.VehicleNo, TSPL_SD_SHIPMENT_HEAD.Carrier,TSPL_SD_SHIPMENT_HEAD.Remarks,TSPL_SD_SHIPMENT_DETAIL.Landed_Cost_Rate,TSPL_SD_SHIPMENT_DETAIL.Landed_Cost_Amount , TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Row_Type,TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount," & _
    '"TSPL_SD_SHIPMENT_DETAIL.Item_Cost as basicRate,TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt as BasicTotal,TSPL_SD_SHIPMENT_DETAIL.Unit_Cost_Tax_Rate as UCTR," & _
    '"TSPL_SD_SHIPMENT_DETAIL.Unit_Cost_Tax as uctax,TSPL_SD_SHIPMENT_DETAIL.Item_Desc,TSPL_SD_SHIPMENT_DETAIL.Unit_code,TSPL_SD_SHIPMENT_DETAIL.Qty,TSPL_SD_SHIPMENT_DETAIL.Rejected_Qty,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_SD_SHIPMENT_HEAD.Total_Amt,TSPL_SD_SHIPMENT_DETAIL.ITEM_COST," & _
    ' "TSPL_VENDOR_MASTER.Add1 as venAdd1, TSPL_VENDOR_MASTER.Add2 as vanadd2, TSPL_VENDOR_MASTER.Add3 as venadd3, " & _
    '"tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_HEAD.tax1_amt,0) as txt1amt,tax2.Tax_Code_Desc as tax2name," & _
    '"isnull (TSPL_SD_SHIPMENT_HEAD.tax2_amt,0) as txt2amt,tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_HEAD.tax3_amt,0) as txt3amt," & _
    '"tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_HEAD.tax4_amt,0) as txt4amt,tax5.Tax_Code_Desc as tax5name," & _
    '"isnull (TSPL_SD_SHIPMENT_HEAD.tax5_amt,0) as txt5amt,tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_HEAD.tax6_amt,0) as txt6amt " & _
    '",tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_HEAD.tax7_amt,0) as txt7amt,tax8.Tax_Code_Desc as tax8name," & _
    '"isnull (TSPL_SD_SHIPMENT_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_HEAD.tax9_amt,0) as txt9amt," & _
    '"tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_HEAD.tax10_amt,0) as txt10amt, TSPL_COMPANY_MASTER.Comp_Name as compname, " & _
    '"TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_SD_SHIPMENT_DETAIL.Qty," & _
    '"case when tax1.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax1_amt else null end as Tax1Recoverable," & _
    '"case when tax2.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt else null end as Tax2Recoverable, " & _
    '"case when tax3.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax3_amt else null end as Tax3Recoverable, " & _
    '"case when tax4.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax4_amt else null end as Tax4Recoverable, " & _
    '"case when tax5.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax5_amt else null end as Tax5Recoverable, " & _
    '"case when tax6.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax6_amt else null end as Tax6Recoverable," & _
    '"case when tax7.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax7_amt else null end as Tax7Recoverable, " & _
    '"case when tax8.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax8_amt else null end as Tax8Recoverable, " & _
    '"case when tax9.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax9_amt else null end as Tax9Recoverable," & _
    '"case when tax10.Tax_Recoverable='Y' then TSPL_SD_SHIPMENT_HEAD.tax10_amt else null end as Tax10Recoverable, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX1_Rate ,0),103)+'%' as txt1Rate," & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX2_Rate   ,0),103)+'%' as txt2Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX3_Rate  ,0),103)+'%' as txt3Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX4_Rate  ,0),103)+'%' as txt4Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX5_Rate  ,0),103)+'%' as txt5Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX6_Rate  ,0),103)+'%' as txt6Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX7_Rate  ,0),103)+'%' as txt7Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX8_Rate  ,0),103)+'%' as txt8Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX9_Rate  ,0),103)+'%' as txt9Rate, " & _
    '"convert(varchar,isnull (TSPL_SD_SHIPMENT_HEAD.TAX10_Rate  ,0),103)+'%' as txt10Rate," & _
    '"TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount as Value,(select SUM(rejected_qty) from TSPL_SD_SHIPMENT_DETAIL where Document_Code=TSPL_SD_SHIPMENT_HEAD.Document_Code) as Rej_qty, (select SUM(TSPL_MRN_DETAIL.MRN_Qty) from TSPL_SD_SHIPMENT_DETAIL left outer join TSPL_MRN_DETAIL on TSPL_MRN_DETAIL .MRN_No=TSPL_SD_SHIPMENT_DETAIL.Order_Code and TSPL_MRN_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code where Document_Code =TSPL_SD_SHIPMENT_HEAD.Document_Code)as MrnTotQty, (select SUM(Qty) from TSPL_SD_SHIPMENT_DETAIL where Document_Code=TSPL_SD_SHIPMENT_HEAD.Document_Code) as SRNQtyTotal, (select case when COUNT(xxx.PI_No)>1 then Min(xxx.PI_No)+ ' *' else Min(xxx.PI_No)end as PINO from" & _
    '" ( select TSPL_PI_DETAIL.PI_No from TSPL_PI_DETAIL  where  TSPL_PI_DETAIL.SRN_Id= TSPL_SD_SHIPMENT_HEAD.Document_Code " & _
    '" GROUP by TSPL_PI_DETAIL.PI_No)xxx) as PInvNo  ,    " & _
    '       " TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name1 as Add1Name, " & _
    '     " TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt1 as Add1 , " & _
    '     "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name2 as Add2Name, " & _
    '     "   TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt2 as Add2 , " & _
    '     "    TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name3 as Add3Name, " & _
    '     "   TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt3 as Add3 , " & _
    '     "    TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name4 as Add4Name, " & _
    '     "    TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt4 as Add4 , " & _
    '     "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name5 as Add5Name, " & _
    '      "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt5 as Add5 , " & _
    '      "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name6 as Add6Name, " & _
    '      "    TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt6 as Add6 , " & _
    '      "    TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name7 as Add7Name, " & _
    '      "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt7 as Add7 , " & _
    '      "       TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name8 as Add8Name, " & _
    '      "      TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt8 as Add8 , " & _
    '       "      TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name9 as Add9Name, " & _
    '       "      TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt9 as Add9 , " & _
    '       "      TSPL_SD_SHIPMENT_HEAD.Add_Charge_Name10 as Add10Name, " & _
    '       "     TSPL_SD_SHIPMENT_HEAD.Add_Charge_Amt10 as Add10,TSPL_SD_SHIPMENT_HEAD.Against_RGP,TSPL_SD_SHIPMENT_DETAIL .Specification   " & _
    ' " FROM  TSPL_SD_SHIPMENT_DETAIL INNER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_DETAIL.Document_Code = TSPL_SD_SHIPMENT_HEAD.Document_Code " & _
    ' "INNER JOIN TSPL_COMPANY_MASTER ON TSPL_SD_SHIPMENT_HEAD.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code  " & _
    ' "INNER JOIN TSPL_VENDOR_MASTER ON TSPL_SD_SHIPMENT_HEAD.Customer_Code = TSPL_VENDOR_MASTER.Customer_Code " & _
    ' "left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SHIPMENT_HEAD.tax1  " & _
    ' "left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SHIPMENT_HEAD.tax2 " & _
    ' "left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SHIPMENT_HEAD .TAX3 " & _
    ' "left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax4 " & _
    ' "left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SHIPMENT_HEAD .tax5 " & _
    ' "left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX6  " & _
    ' "left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX7  " & _
    ' "left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX8 " & _
    ' "left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX9 " & _
    ' " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX10  " & _
    ' "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_SD_SHIPMENT_HEAD.Bill_To_Location  " & _
    ' " where TSPL_SD_SHIPMENT_HEAD .Item_Type not in('F')"

    '                If FromDate.HasValue AndAlso ToDate.HasValue Then
    '                    strquery += " and Convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103)>=Convert(date,'" + FromDate + "',103)and Convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103)<=Convert(date,'" + ToDate + "',103) "

    '                End If
    '                If ArrLocation IsNot Nothing AndAlso ArrLocation.Count > 0 Then
    '                    strquery += "and TSPL_LOCATION_MASTER.Loc_Segment_Code  IN (" + clsCommon.GetMulcallString(ArrLocation) + ") "
    '                End If
    '                If ArrSrnNo IsNot Nothing AndAlso ArrSrnNo.Count > 0 Then
    '                    strquery += " and TSPL_SD_SHIPMENT_HEAD.Document_Code in (" + clsCommon.GetMulcallString(ArrSrnNo) + ")  "
    '                End If
    '                If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
    '                    strquery += " and TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")  "

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


    Public Shared Function ReverseAndUnpost(ByVal strCode As String, isDCSEntrySale As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ReverseAndUnpost(strCode, isDCSEntrySale, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function ReverseAndUnpost(ByVal strCode As String, isDCSEntrySale As Boolean, ByVal trans As SqlTransaction) As Boolean
        Try
            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select Status from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + strCode + "'"
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If
            Dim dt As DataTable
            If Not isDCSEntrySale Then
                Qry = " select distinct DOCUMENT_CODE from TSPL_SD_SHIPMENT_DETAIL where DOCUMENT_CODE='" + strCode + "' and REF_PK_ID is not null"
                dt = clsDBFuncationality.GetDataTable(Qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Throw New Exception("This Document No [" + strCode + "] is generated by DCS Entry sale. You can't reverse and unpost this.")
                End If
            End If
            Qry = " select isnull(GPCode,'') as GPCode from tspl_sd_shipment_head where Document_Code = '" & strCode & "' and isnull(GPCode,'')<>''  "
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("GatePass No [" + clscommon.mycstr(dt.Rows(0)("GPCode")) + "] is already generated with this Shipment No [" + strCode + "]. You can't reverse and unpost this.")
            End If
            Qry = "select distinct DOCUMENT_CODE from TSPL_SD_SALE_INVOICE_DETAIL where Shipment_Code='" + strCode + "'"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    clsPSInvoiceHead.ReverseAndUnpost(clsCommon.myCstr(dr("DOCUMENT_CODE")), trans)
                Next
            End If

            Qry = "select distinct adjustment_no from TSPL_Receipt_Adjustment_Header where  Doc_No= (select Sale_Invoice_no from TSPL_SD_SHIPMENT_head where document_code='" + strCode + "' )"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    clsAdjustmentEntryReceivables.ReverseAndUnpost(clsCommon.myCstr(dr("adjustment_no")), trans)
                    'connectSql.RunSp("sp_TSPL_Receipt_Adjustment_Detail_Delete", New SqlParameter("@Adjustment_No", clsCommon.myCstr(dr("adjustment_no"))))
                    'connectSql.RunSp("sp_TSPL_Receipt_Adjustment_Header_Delete", New SqlParameter("@Adjustment_No", clsCommon.myCstr(dr("adjustment_no"))))

                    Qry = "delete from TSPL_Receipt_Adjustment_Detail where Adjustment_No ='" + clsCommon.myCstr(dr("adjustment_no")) + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                    Qry = "delete from TSPL_Receipt_Adjustment_Header where Adjustment_No ='" + clsCommon.myCstr(dr("adjustment_no")) + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Next
            End If

            If objCommonVar.CreateAutoReceiptEntryDCSSale Then
                Qry = "select Receipt_No from TSPL_RECEIPT_DETAIL  where Document_No='" + strCode + "'"
                Dim dtReceipt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
                If dt IsNot Nothing AndAlso dtReceipt.Rows.Count > 0 Then
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_SD_SHIPMENT_HEAD set Receipt_No=null where Document_Code='" + strCode + "'", trans)
                    For Each dr As DataRow In dtReceipt.Rows
                        Dim strReceiptNo As String = clsCommon.myCstr(dr("Receipt_No"))
                        clsRcptEntryHeader.ReverseAndUnpost(strReceiptNo, trans)
                        Qry = "delete from TSPL_RECEIPT_DETAIL where Receipt_No='" + strReceiptNo + "'"
                        clsDBFuncationality.ExecuteNonQuery(Qry, trans)

                        Qry = "delete from TSPL_RECEIPT_HEADER where Receipt_No='" + strReceiptNo + "'"
                        clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                    Next
                End If
            End If

            Qry = "Select document_no from TSPL_VENDOR_INVOICE_HEAD where Against_MCC_Material_Sale='" + strCode + "'"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    clsVedorInvoiceHead.ReverseAndUnpost(clsCommon.myCstr(dr("document_no")), trans)
                    clsVedorInvoiceHead.DeleteData(clsCommon.myCstr(dr("document_no")), trans)
                Next
            End If


            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='SD-SH' and Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            clsBatchInventory.ReverseAndUnpost("MCC-MSALE", strCode, trans)

            ''richa ERO/18/06/19-000645 18 June ,2019
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT_NEW", "Source_Doc_No", trans)
            Qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" + strCode + "' and Trans_Type='MCC-MSALE' "
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            Qry = "delete from TSPL_INVENTORY_MOVEMENT_NEW where Source_Doc_No='" + strCode + "' and Trans_Type='MCC-MSALE' "
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            Qry = "Update TSPL_SD_SHIPMENT_HEAD set Status = 0 where Document_Code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SD_SHIPMENT_HEAD", "Document_Code", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


    Public Shared Function GetAttachQry(ByVal isCancel As Boolean, ByVal strDate As DateTime, ByVal strCode As String) As String
        Dim QryShowStatus As String = ""
        Dim EnableDynamicQRCodeForB2CInvoice As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EnableDynamicQRCodeForB2CInvoice, clsFixedParameterCode.EnableDynamicQRCodeForB2CInvoice, Nothing)) = "1", True, False))
        Dim MultiplySubsidyWithQuantity As Boolean = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MultiplySubsidyWithQuantity, clsFixedParameterCode.MultiplySubsidyWithQuantity, Nothing))

        Dim ShowStatusForSale As Double = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowStatusForSales' And Type ='ShowStatusForSales'")
        If clsCommon.CompairString(clsCommon.myCstr(ShowStatusForSale), "1") = CompairStringResult.Equal Then
            QryShowStatus = " ,(case when TSPL_SD_SHIPMENT_HEAD.status =1 then 'AUTHORIZED' else 'NOT AUTHORIZED' end) as Status "
        Else
            QryShowStatus = ""
        End If
        Dim SerialNo As String = ""
        Dim SerialNoColumn As String = ""
        Dim ShowSerialNoForSales As Double = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowSerialNoForSales' And Type ='ShowSerialNoForSales'")
        If clsCommon.CompairString(clsCommon.myCstr(ShowSerialNoForSales), "1") = CompairStringResult.Equal Then
            SerialNoColumn = " ,1 As SerialNoText , aa.Serial_No As [SerialNo]  "
            'SerialNo = "  left outer join TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL  on TSPL_SD_SHIPMENT_DETAIL.Item_Code  =TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Main_Item_Code And TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.IS_Principle=1 ANd TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Location_Code =TSPL_SD_SHIPMENT_DETAIL.Location "
            SerialNo = " left outer join (select distinct Doc_No,Serial_No,Main_Item_Code,Location_Code from TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL WHERE Is_principle='1' AND ISNULL(Serial_No,'')<>'' and Doc_No in (select Doc_No from TSPL_MF_PRINCIPLE_RECEIPT_HEAD where Status='1'))aa  on TSPL_SD_SHIPMENT_DETAIL.Item_Code  =AA.Main_Item_Code  ANd aa.Location_Code =TSPL_SD_SHIPMENT_DETAIL.Location  "
        Else
            SerialNoColumn = " ,0 As SerialNoText "
            SerialNo = ""
        End If
        Dim isShowQRcode As Integer = 0
        If clsERPFuncationality.GetQRCodeStatus(strDate) = True AndAlso EnableDynamicQRCodeForB2CInvoice = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EInvoice_Type from  TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No = '" + strCode + "'")), "BC") = CompairStringResult.Equal Then
            isShowQRcode = 1
        End If
        Dim Qry As String = "select *, itemcost/ConversionFactor As RateLtr from (  select "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
            Qry += "1 As CopyType,"
        End If

        Qry += " TSPL_SD_SHIPMENT_DETAIL.Batch_No as Manual_BatchNo,TabBatch.Batch_No, (CASE when TSPL_SD_SHIPMENT_DETAIL.Scheme_Item='Y' then 0 else ((case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1 then 0 else  TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount end)) end)  as valueInRs,
ITEMDETAIL1.Conversion_Factor As CF,TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor,Cast((TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ITEMDETAIL2.Conversion_Factor As Int) As ConvQtyInCrate,
convert(varchar(12),TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)Supply_Date,case when TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 'Morning' else 'Evening' end as Shift_Type,  ((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SHIPMENT_DETAIL.Qty)/ITEMDETAIL3.Conversion_Factor) as QTY_LTRKG, ITEMDETAIL2.UOM_Code As CCrateUOM,
Case when tspl_item_master.Is_FreshItem = 1 then 'LTR' When tspl_item_master.Is_Ambient = 1 then 'KG' End As Qty_LTRKG_UOM,((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SHIPMENT_DETAIL.Qty)/ItemConversionInPouch.Conversion_Factor) as QtyInPouch,
((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SHIPMENT_DETAIL.Qty)/ItemConversionInPACK.Conversion_Factor) as QtyInPack,
ISNull(TSPL_SD_SHIPMENT_DETAIL.Transporter_Commission_Rate,0.00) As TCRate,isnull(TSPL_SD_SALE_INVOICE_HEAD.Transporter_Commission_TotalAmt,0.00) as TCAmt,(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt/((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SHIPMENT_DETAIL.Qty)/ItemConversionInPouch.Conversion_Factor)) as Rate_Per_Pouch,
(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt/((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SHIPMENT_DETAIL.Qty)/ItemConversionInPACK.Conversion_Factor)) as Rate_Per_Pack,ItemConversionInPouch.UOM_Code As CPPouch_UOM_Code,ItemConversionInPACK.UOM_Code  As CPack_UOM_Code,"
        If isCancel Then
            Qry += " 'Cancelled' As Report_Status,"
        Else
            Qry += " '' As Report_Status,"
        End If
        Qry += " " + clsCommon.myCstr(isShowQRcode) + " as isShowQRcode, TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type ,  cast(TSPL_SD_SALE_INVOICE_HEAD.BarCode_Img as image) As BarCode_Img, TSPL_SD_SHIPMENT_HEAD.Bill_To_Location, TSPL_LOCATION_MASTER.Location_Desc, TSPL_SD_SHIPMENT_HEAD.Is_Taxable,TSPL_SD_SHIPMENT_HEAD.Electronic_Ref_No,TSPL_SD_SHIPMENT_HEAD.EWayBillNo, convert(varchar,TSPL_SD_SHIPMENT_HEAD.EWayBillDate,103) as EWayBillDate,  StateMasterForLocation.GST_STATE_CODE AS From_GstStateCode," &
          " TSPL_LOCATION_MASTER.GSTNO as From_Loc_GstinNo, TSPL_STATE_MASTER.GST_STATE_Code AS Cust_GstStateCode,  TSPL_CUSTOMER_MASTER.GSTNO as Cust_GstInNo,    case when coalesce(p_cust.GST_STATE_CODE,'')='' then TSPL_state_Master.GST_STATE_CODE       when coalesce(p_cust.GST_STATE_CODE,'')<>'' then p_cust .GST_STATE_CODE    end as P_GST_STATE_CODE," &
          " case when coalesce(p_cust.P_GSTIN_NO,'')='' then TSPL_CUSTOMER_MASTER .GSTNO  when coalesce(p_cust.P_GSTIN_NO,'')<>'' then p_cust .P_GSTIN_NO end as P_GSTIN_NO,    TSPL_ITEM_MASTER.HSN_Code,  " &
          " isnull(dtax1.Type,'') as tax1Type,isnull(dtax2.Type,'') as tax2Type,isnull(dtax3.Type,'') as tax3Type,isnull(dtax4.Type,'') as tax4Type,isnull(dtax5.Type,'') as tax5Type,isnull(dtax6.Type,'') as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type, " &
        " ISNULL(TSPL_SD_SHIPMENT_head.TAX1_Amt ,0)as DTax1_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX2_Amt ,0)as DTax2_Amt, ISNULL(TSPL_SD_SHIPMENT_head.TAX3_Amt ,0)as DTax3_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX4_Amt ,0)as DTax4_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX5_Amt ,0)as DTax5_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX6_Amt ,0)as DTax6_Amt, ISNULL(TSPL_SD_SHIPMENT_head.TAX7_Amt ,0)as DTax7_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX8_Amt ,0)as DTax8_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX9_Amt ,0)as DTax9_Amt,ISNULL(TSPL_SD_SHIPMENT_head.TAX10_Amt ,0)as DTax10_Amt," &
          "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX1_Rate,0)as DTax1_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX2_Rate,0)as DTax2_Rate," &
          "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX3_Rate,0)as DTax3_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX4_Rate,0)as DTax4_Rate," &
          "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX5_Rate,0)as DTax5_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX6_Rate,0)as DTax6_Rate," &
          "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX7_Rate,0)as DTax7_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX8_Rate,0)as DTax8_Rate," &
          "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX9_Rate,0)as DTax9_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX10_Rate,0)as DTax10_Rate," &
         " TSPL_LOCATION_MASTER.State as loc_state_code,tspl_location_master.HOAdd1 ,tspl_location_master.HOAdd2,TSPL_SD_SHIPMENT_HEAD.HeadDisc_PerAmt,'" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") + "' as RunDate,"
        Qry += " TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end"
        Qry += "+ Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ',Phone'+TSPL_COMPANY_MASTER.Phone1 end "
        Qry += "+  Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ',  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End "
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end "
        Qry += " as Comp_Address,TSPL_SD_SHIPMENT_HEAD.RoundOffAmount , "
        '' Anubhooti 20-Feb-2015 (Show VLC,VSP Name)
        Qry += " tspl_vlc_master_head.vlc_code_vlc_uploader As VLC_Code,tspl_vlc_master_head.vlc_name,tspl_vlc_master_head.VSP_Code, "
        ''
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            Qry += " TSPL_VENDOR_MASTER.Phone1 as VSP_PHONE, "
        End If
        Qry += "   case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add1  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add1 end as P_Add1, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add2  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add2 end as P_Add2, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add3  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add3 end as P_Add3, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PIN_Code   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Pin_No  end as P_PinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .CST    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_CST_No   end as P_CstNo,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Tin_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .p_Tin_No   end as P_TinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Email    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Email  end as P_Email,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Fax     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Fax   end as P_Fax, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PAN      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Cust_PAN    end as P_Cust_PAN,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Lst_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_LST_No    end as P_LstNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Cust_Code      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_code   end as P_CustCode, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Customer_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_name    end as P_Cust_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CITY_MASTER   .City_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_City_Name    end as P_City_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_state_Master.state_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_state_Name    end as P_State_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then     case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Phn    end as P_Cust_Phn,TSPL_CUSTOMER_MASTER.Cust_Code  ,TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1,TSPL_CUSTOMER_MASTER.Add2 as Cust_add2,TSPL_CUSTOMER_MASTER.Add3 as cust_add3,  case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End  as Cust_Phn,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_TinNo ,TSPL_CUSTOMER_MASTER.CST as Cust_CSTNo,TSPL_CUSTOMER_MASTER.Lst_No as Cust_LSTNo,TSPL_CUSTOMER_MASTER.Email as Cust_Email ,TSPL_CUSTOMER_MASTER.PAN as Customer_PAN,TSPL_CUSTOMER_MASTER.PIN_Code as Cust_PinCode,TSPL_CITY_MASTER.City_Name as Loctn_City_Name_Desc,TSPL_CUSTOMER_MASTER.Fax as Cust_Fax,TSPL_STATE_MASTER .STATE_NAME  as Cust_State_Name, "
        Qry += " TSPL_CITY_MASTER_ForCustomer.City_Name  as Cust_City_Name,case when (TSPL_SD_SHIPMENT_HEAD.Dispatch_Terms )='FE' then 'Freight Extra' else  TSPL_SD_SHIPMENT_HEAD.Dispatch_Terms  end  as Dispatch_Terms,TSPL_LOCATION_MASTER .Add1 as Loc_Add1,TSPL_LOCATION_MASTER.Add2 as Loc_ADd2,TSPL_LOCATION_MASTER.Add3  as Loc_Add3, StateMasterForLocation.State_Name as LocationState, TSPL_LOCATION_MASTER.Pin_Code as Loc_Pin_Code,TSPL_LOCATION_MASTER.TIN_No as Loc_TinNo,Case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as  Loc_Phn,TSPL_LOCATION_MASTER.Email as Loc_Email,TSPL_SD_SHIPMENT_HEAD.Total_Add_Charge,convert(varchar,TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Date ,103) as Delivery_Date, TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS,( case when TSPL_SD_SHIPMENT_HEAD.Invoice_Type='R' then 'Retail Invoice' else 'Tax Invoice' end) as Invoice_Type,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Cash_Sale_InvNo,convert(varchar(15),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Cash_Sale_InvDate,TSPL_SD_SHIPMENT_HEAD.Invoice_Type as Inv_Type,TSPL_SD_SHIPMENT_HEAD.Is_CashSale,TSPL_LOCATION_MASTER.Loc_Short_Name,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Inv_Date,103) as Inv_Date,case when (TSPL_SD_SHIPMENT_HEAD.Payment_Terms)='A' then 'Advance' else TSPL_SD_SHIPMENT_HEAD.Payment_Terms end  as Payment_Terms ,TSPL_SD_SHIPMENT_HEAD.Transporter_Name ,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No   ,Case When IsNull(TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_Date,'')<>'' Then convert(varchar,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_Date,103) Else convert(varchar,TSPL_SD_SHIPMENT_HEAD.Inv_Date,103) End as Sale_Invoice_Date,TSPL_COMPANY_MASTER.Tin_No as Comp_Tin_No,TSPL_COMPANY_MASTER.Pan_No as Comp_PANNO,TSPL_COMPANY_MASTER.CST_LST as Comp_CST_No,TSPL_COMPANY_MASTER.CINNo as Comp_CinNo,TSPL_COMPANY_MASTER.Pincode  as Comp_Pin_Code, TSPL_SHIP_TO_LOCATION.Ship_To_Desc as shipName, TSPL_SHIP_TO_LOCATION.add1 as ship_Add1, TSPL_SHIP_TO_LOCATION.Add2 as ship_add2 ,TSPL_SHIP_TO_LOCATION.Add3 as ship_add3  ,TSPL_SHIP_TO_LOCATION.Pin_Code,TSPL_CITY_MASTER.STATE_CODE  ,Tspl_City_master.City_Name,TSPL_EMPLOYEE_MASTER.Emp_Name as SalesManName,TSPL_SD_SHIPMENT_HEAD.Inv_No, TSPL_SD_SHIPMENT_HEAD.Dept_Desc , TSPL_SD_SHIPMENT_HEAD.Remarks ,  TSPL_SD_SHIPMENT_HEAD.Terms_Code,TSPL_SD_SHIPMENT_HEAD.VehicleNo ,TSPL_SD_SHIPMENT_HEAD.Challan_No,TSPL_SD_SHIPMENT_HEAD.RateDiff_Amt,TSPL_SD_SHIPMENT_HEAD.Gross_Amount,TSPL_VENDOR_MASTER.Zone_Code, "
        Qry += " TSPL_SD_SHIPMENT_DETAIL .Specification as  specification,   TSPL_SD_SHIPMENT_HEAD.sale_invoice_no as DocNo , TSPL_SD_SHIPMENT_HEAD.Description, "
        Qry += "  convert(varchar ,TSPL_SD_SHIPMENT_HEAD .Document_Date,103)as Document_Date ,FORMAT(TSPL_SD_SHIPMENT_HEAD.Document_Date, 'hh:mm:ss tt') AS DocumentTime, TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order, TSPL_SD_SHIPMENT_HEAD.Item_Type ,  TSPL_SD_SHIPMENT_HEAD.Customer_Code, "
        Qry += " TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1,TSPL_CUSTOMER_MASTER.add2 as customer_Add2,TSPL_CUSTOMER_MASTER.Add3 as customer_Add3 ,TSPL_CUSTOMER_MASTER.State as customer_city_State,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_Tin_No ,TSPL_CUSTOMER_MASTER.PIN_Code as Customer_Pin_Code , TSPL_SD_SHIPMENT_HEAD .Terms_Code as termscode ,TSPL_SD_SHIPMENT_HEAD .Ref_No as ref_no ,"
        Qry += " TSPL_SD_SHIPMENT_HEAD .Comments as comments ,  TSPL_SD_SHIPMENT_HEAD .Discount_Amt as dis_amt,"
        If MultiplySubsidyWithQuantity Then
            Qry += "  isnull(TSPL_SD_SHIPMENT_HEAD.TotalSubsidyDisAmt,0)  as dis_amt1,"
        Else
            Qry += "  TSPL_SD_SHIPMENT_DETAIL.Total_Disc_Amt  as dis_amt1,"
        End If
        Qry += " TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount  as aftrdiscount ,TSPL_SD_SHIPMENT_HEAD .Total_Amt as Total_amount,"
        Qry += " TSPL_SD_SHIPMENT_HEAD.Discount_Base as bfrdisc_amount, TSPL_COMPANY_MASTER.Access_Officer as FSSAI,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Tcan_No AS WebSite ,TSPL_COMPANY_MASTER.Phone1 AS COMP_PHONE,TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt ,TSPL_SD_SHIPMENT_HEAD.Recommended_By, "
        Qry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_detail.tax1_amt,0) as txt1amt, TSPL_SD_SHIPMENT_detail.TAX1_Rate,  tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SHIPMENT_detail.tax2_amt,0) as txt2amt,TSPL_SD_SHIPMENT_detail.TAX2_Rate, tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_detail.tax3_amt,0) as txt3amt,TSPL_SD_SHIPMENT_detail.TAX3_Rate,   tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_detail.tax4_amt,0) as txt4amt,TSPL_SD_SHIPMENT_detail.TAX4_Rate,   tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SHIPMENT_detail.tax5_amt,0) as txt5amt,TSPL_SD_SHIPMENT_detail.TAX5_Rate,   tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_detail.tax6_amt,0) as txt6amt,TSPL_SD_SHIPMENT_detail.TAX6_Rate,   tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_detail.tax7_amt,0) as txt7amt,TSPL_SD_SHIPMENT_detail.TAX7_Rate,   tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SHIPMENT_detail.tax8_amt,0) as txt8amt,TSPL_SD_SHIPMENT_detail.TAX8_Rate,    tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_detail.tax9_amt,0) as txt9amt,TSPL_SD_SHIPMENT_detail.TAX9_Rate,   tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_detail.tax10_amt,0) as txt10amt, TSPL_SD_SHIPMENT_detail.TAX10_Rate, "
        'Qry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_HEAD.tax1_amt,0) as txt1amt,  "
        'Qry += " tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SHIPMENT_HEAD.tax2_amt,0) as txt2amt,  "
        'Qry += " tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_HEAD.tax3_amt,0) as txt3amt,  "
        'Qry += " tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_HEAD.tax4_amt,0) as txt4amt,  "
        'Qry += " tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SHIPMENT_HEAD.tax5_amt,0) as txt5amt,  "
        'Qry += " tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_HEAD.tax6_amt,0) as txt6amt,  "
        'Qry += " tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_HEAD.tax7_amt,0) as txt7amt,  "
        'Qry += " tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SHIPMENT_HEAD.tax8_amt,0) as txt8amt,   "
        'Qry += " tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_HEAD.tax9_amt,0) as txt9amt,  "
        'Qry += " tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_HEAD.tax10_amt,0) as txt10amt,  "
        Qry += " isnull(TSPL_SD_SHIPMENT_HEAD .Total_Tax_Amt,0) as total_tax_amt, TSPL_SD_SHIPMENT_HEAD.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,TSPL_COMPANY_MASTER.ISO_No as ISO_No,TSPL_COMPANY_MASTER.ISO_Date,TSPL_COMPANY_MASTER.Access_Officer,ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+ Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as Phone,TSPL_COMPANY_MASTER.Fax as Comp_Fax,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.GSTReg_No As SellerGST,TSPL_COMPANY_MASTER.Pan_No,ISNULL(tspl_company_Master.ADD1,'') as Comp_add,"
        Qry += " TSPL_SD_SHIPMENT_DETAIL.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc + case when Scheme_Item ='Y' then ' (Free Scheme)' else '' end    as itemdesc, TSPL_SD_SHIPMENT_DETAIL.Row_Type,TSPL_SD_SHIPMENT_DETAIL.Qty as qty,TSPL_SD_SHIPMENT_DETAIL.unit_code as uom,TSPL_SD_SHIPMENT_DETAIL.item_cost as itemcost,TSPL_SD_SHIPMENT_DETAIL.amount as amount,TSPL_SD_SHIPMENT_HEAD.TAX1,TSPL_SD_SHIPMENT_HEAD.TAX2,TSPL_SD_SHIPMENT_HEAD.TAX3,TSPL_SD_SHIPMENT_HEAD.TAX4,TSPL_SD_SHIPMENT_HEAD.TAX5,TSPL_SD_SHIPMENT_HEAD.ReceiverName, isnull(TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt,0) as TotalSubsidyAmt "
        Qry += " " & QryShowStatus & " "
        Qry += " " & SerialNoColumn & "  "
        Qry += " ,TSPL_SD_SHIPMENT_HEAD.Vehicle_Code,Case When IsNull(TSPL_SD_SHIPMENT_HEAD.VehicleNo,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.VehicleNo Else TSPL_Vehicle_Master.Number End As Vehicle_Manual_No,Case when IsNull(TSPL_SD_SHIPMENT_HEAD.Route_No,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.Route_No End As Route_No,Case when IsNull(TSPL_SD_SHIPMENT_HEAD.Route_Desc,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.Route_Desc Else TSPL_Route_Master.Route_Desc End As Route_Desc,TSPL_COMPANY_MASTER.Bank_Name, TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.GSTINNo,TSPL_COMPANY_MASTER.GSTReg_No "
        If isCancel Then
            Qry += " from TSPL_SD_SHIPMENT_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_DETAIL  "
        Else
            Qry += " from TSPL_SD_SHIPMENT_DETAIL   "
        End If

        Qry += " " & SerialNo & " "
        If isCancel Then
            Qry += " left outer join TSPL_SD_SHIPMENT_HEAD_Cancel_Data As TSPL_SD_SHIPMENT_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code  =TSPL_SD_SHIPMENT_DETAIL.Document_Code   "
        Else
            Qry += " left outer join TSPL_SD_SHIPMENT_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code  =TSPL_SD_SHIPMENT_DETAIL.Document_Code   "
        End If
        Qry += "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code
       left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code in (select UOM_Code  from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Print_UOM=1)and TSPL_item_uom_detail.Print_UOM=1   ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code "
        Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And Print_UOM=1) as ITEMDETAIL3 on ITEMDETAIL3.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And UOM_code IN ('Crate','CFC','TIN','JAR','Carton','Box')) as ITEMDETAIL2 on ITEMDETAIL2.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_code IN ('CUP','Pouch')) as ItemConversionInPouch on ItemConversionInPouch.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_code In ('Matka','PACK')) as ItemConversionInPACK on ItemConversionInPACK.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code"
        Qry += " Left Outer Join TSPL_Route_Master On TSPL_Route_Master.Route_No=(Case when IsNull(TSPL_SD_SHIPMENT_HEAD.Route_No,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.Route_No Else TSPL_SD_SHIPMENT_HEAD.sale_Route_Code End)
Left Outer Join TSPL_Vehicle_Master On TSPL_Vehicle_Master.Vehicle_Id=TSPL_Route_Master.Vehicle_Code "
        Qry += " left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_SD_SHIPMENT_HEAD .Ship_To_Location "
        Qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_SHIP_TO_LOCATION.City_Code "
        Qry += " left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_SD_SHIPMENT_HEAD.Salesman_Code "
        Qry += " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SHIPMENT_HEAD.tax1  "
        Qry += " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SHIPMENT_HEAD.tax2  "
        Qry += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SHIPMENT_HEAD .TAX3  "
        Qry += " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax4  "
        Qry += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SHIPMENT_HEAD .tax5  "
        Qry += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX6  "
        Qry += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX7  "
        Qry += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX8  "
        Qry += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX9 "
        Qry += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX10     "
        Qry += "left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_SD_SHIPMENT_DETAIL.tax1  " &
            " left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_SD_SHIPMENT_DETAIL.tax2   " &
            " left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_SD_SHIPMENT_DETAIL.TAX3  " &
            " left outer join TSPL_TAX_MASTER as dtax4 on dtax4.Tax_Code= TSPL_SD_SHIPMENT_DETAIL.tax4  " &
             "  left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_SD_SHIPMENT_DETAIL.tax5  " &
               " left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX6  " &
                " left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX7  " &
                 "  left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX8   " &
                " left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX9 " &
            "left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX10  "
        Qry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_SD_SHIPMENT_HEAD.comp_code  "
        Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SHIPMENT_HEAD.Customer_Code  left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = TSPL_CUSTOMER_MASTER.Cust_Code   LEFT join (select TSPL_CUSTOMER_MASTER.GSTNO AS P_GSTIN_NO,GST_STATE_CODE, pan as P_Cust_PAN, Cust_Code as P_cust_code,Customer_Name as P_cust_name,Add1 as P_cust_add1 ,Add2  as P_cust_add2,add3  as P_cust_add3,PIN_Code as P_Pin_No,Tin_No as p_Tin_No ,State as P_state,Email as P_Email,fax as P_Fax,TSPL_CITY_MASTER.City_Name as  P_City_Name,TSPL_STATE_MASTER.STATE_NAME as P_State_Name  ,case when ISNULL(Phone1,'')='(+__)__________' then '' else Phone1 end +  Case When   ISNULL(Phone2,'')<>'(+__)__________' Then ', '+ Phone2 Else'' End as  P_Phn,CST as P_CST_No,Terms_Code as P_Terms,Lst_No as P_LST_No  from TSPL_CUSTOMER_MASTER   left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_CUSTOMER_MASTER.City_Code "
        Qry += " left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State =TSPL_STATE_MASTER.STATE_CODE  "
        Qry += " ) p_cust on p_cust.P_cust_code=TSPL_CUSTOMER_MASTER.Parent_Customer_No and TSPL_CUSTOMER_MASTER.Parent_Customer_YN='N'"
        '' Anubhooti 20-Feb-2015 (Show VLC,VSP Name)
        Qry += "  LEFT OUTER JOIN tspl_customer_vendor_mapping ON tspl_customer_vendor_mapping.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code   "
        Qry += "LEFT OUTER JOIN tspl_vlc_master_head ON tspl_customer_vendor_mapping.Vendor_Code =tspl_vlc_master_head.vsp_code"
        Qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_SD_SHIPMENT_HEAD.Bill_To_Location "
        Qry += " left outer join TSPL_CITY_MASTER  as TSPL_CITY_MASTER_ForCustomer on TSPL_CITY_MASTER_ForCustomer.City_Code =TSPL_CUSTOMER_MASTER.City_Code"
        Qry += " left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE = TSPL_CUSTOMER_MASTER.State"
        Qry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code "
        Qry += " left outer join TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE on TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code =TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS"
        Qry += " LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code "
        Qry += " LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State " &
        " LEFT OUTER JOIN TSPL_STATE_MASTER StateMasterForLocation ON StateMasterForLocation.State_Code=TSPL_LOCATION_MASTER.State " &
        " left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code 
left join ( select Document_Code,Parent_Line_No, STRING_AGG(Batch_No +' -'+convert(varchar(10),Qty)+'', ',') as Batch_No from(
SELECT Document_Code, Batch_No, Qty, Parent_Line_No FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='MCC-MSALE'
)x group by Document_Code,Parent_Line_No                
) TabBatch On TabBatch.Document_Code= TSPL_SD_SHIPMENT_head.Document_Code And TabBatch.Parent_Line_No = TSPL_SD_SHIPMENT_DETAIL.Line_No"
        Qry += "  where 2=2 "
        Qry += "  and  TSPL_SD_SHIPMENT_HEAD.Document_Code = '" + strCode + "' )xx "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
            Qry += " Left OUTER JOIN (Select 1 As COL1, 1 As COL2,  'ORIGINAL COPY' as CopyType1 UNION Select 1 as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XX.CopyType "
        End If
        Return Qry
    End Function
    Public Shared Function funPrint(ByVal Form_ID As String, ByVal isCancel As Boolean, ByVal strDate As DateTime, ByVal strCode As String, ByVal GetPDFPath As Boolean) As String
        Dim pdfPath As String = ""
        Try
            '====================================================
            'Dim strUrl As String = "upi://pay?pa=1234566@ICIC12345.ifsc.npci&pn=TEST&mc=0000&tr=API123456&am=10000&mam=10000&cu=INR&mode=01&b2cSellerGstin=29AABCT1332L000&b2cUPIID=null&b2cBankAcNo=1234566&b2cIFSCCode=ICIC12345&b2cInvNo=API123456&b2cInvDate=10-06-2021&b2cCGSTAmount=0&b2cSGSTAmount=0&b2cIGSTAmount=0&b2cCESS=0&size=150"
            'Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMy    QCCode(strUrl)
            'clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_SD_SALE_INVOICE_HEAD", "TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No='" & txtDocNo.Value & "'")
            '===================================================
            Dim TaxGroup As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Group from TSPL_SD_SHIPMENT_HEAD_Cancel_Data where Document_Code = '" + strCode + "'"))
            Dim IsMandiTax As Double = 0
            IsMandiTax = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code='" + TaxGroup + "' and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y')"))
            Dim IsExampted As Double = 0
            IsExampted = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(Is_Tax_Exempted)  from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" + TaxGroup + "' AND Tax_Group_Type='S' AND Is_Tax_Exempted=1"))

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetAttachQry(isCancel, strDate, strCode))
            If dt.Rows.Count > 0 Then
                ' ==============================Ticket No  ERO/27/11/18-000421 By Prabhakar for  Customer Dashboard on Print==============================================
                Dim dtCustomerOutstanding As DataTable = Nothing
                '=============================================================================
                Dim frmCRV As New frmCrystalReportViewer()
                If clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(dt.Rows(0)("Document_Date"))) Then
                    If clsCommon.myCdbl(dt.Rows(0)("Is_Taxable")) = 1 Then
                        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("LocationState")), clsCommon.myCstr(dt.Rows(0)("Cust_State_Name"))) = CompairStringResult.Equal Then
                            If IsExampted > 0 Then
                                If GetPDFPath Then
                                    pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                Else
                                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                End If
                            ElseIf IsMandiTax > 0 Then
                                If GetPDFPath Then
                                    pdfPath = frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                Else
                                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                End If
                                'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local_WithMandi", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            Else
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalCHU", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalCHU", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalJDH", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalJDH", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalRJS", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalRJS", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalSKR", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalSKR", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalTNK", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalTNK", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalAJM", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_LocalAJM", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                Else
                                    If GetPDFPath Then
                                        pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    Else
                                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                    End If
                                    'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                                End If
                            End If
                        Else
                            If IsExampted > 0 Then
                                If GetPDFPath Then
                                    pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                Else
                                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                End If
                            ElseIf IsMandiTax > 0 Then
                                If GetPDFPath Then
                                    pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Interstate With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                Else
                                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Interstate With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                                End If
                                'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState_WithMandi", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            Else
                                dtCustomerOutstanding = clsCustomerMaster.getCustomerOutstandingOfAmt_Can_Crate("'" & clsCommon.myCstr(dt.Rows(0)("Customer_Code")) & "'", clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")).AddDays(-1), "dd/MMM/yyyy"), clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")), "dd/MMM/yyyy"))
                                If GetPDFPath Then
                                    pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                Else
                                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                                End If
                            End If
                        End If
                    Else
                        If GetPDFPath Then
                            pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                        Else
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxableRJS", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxableCHU", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxableTNK", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            ElseIf clscommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSaleNonTaxableAJM", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            Else
                                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            End If
                        End If
                    End If
                Else
                    If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Is_CashSale")), "Y") = CompairStringResult.Equal Then
                        If GetPDFPath Then
                            pdfPath = frmCRV.funreport(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, "rptMccMaterialSaleCashMemo", "Cash Memo")
                        Else
                            frmCRV.funreport(Form_ID, CrystalReportFolder.MilkProcurement, dt, "rptMccMaterialSaleCashMemo", "Cash Memo")
                        End If
                    Else
                        If GetPDFPath Then
                            pdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
                        Else
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
                        End If
                    End If
                End If
                'frmCrystalReportViewer.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
                frmCRV = Nothing
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        If GetPDFPath Then
            Return pdfPath
        Else
            Return True
        End If
    End Function













End Class

Public Class clsMCCMaterialSaleDetail
#Region "Variables"
    Public ItemwiseTaxCode As String = Nothing
    Public OrgUnit_code As String = ""
    Public Commission_Party As String = Nothing
    Public Commission_Rate As Double = 0
    Public Commission_Amt As Double = 0
    Public Amt_Less_Commission As Double = 0

    Public Document_Code As String = Nothing
    Public Line_No As Integer = 0
    Public Row_Type As String = Nothing
    Public Item_Code As String = Nothing
    Public Item_Desc As String = Nothing 'Not a Table Field
    Public DCS_Sale_Zero_Cost As String = Nothing
    Public Bar_Code As String = Nothing
    Public Qty As Double = 0
    Public Billing_Qty As Double = 0
    Public Balance_Qty As Double = 0
    Public Free_Qty As Double = 0
    Public Order_Code As String = Nothing
    Public Unit_code As String = Nothing '
    Public Billing_Unit_code As String = Nothing '
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
    Public Disc_Per_Unit As Double = 0
    Public Disc_Unit_Amt As Double = 0
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
    'Public Unit_Cogs As Double = 0
    Public SRNTax_Group As String = Nothing 'Not a Table Field

    Public arrSrItem As List(Of clsSerializeInvenotry) = Nothing


    Public Scheme_Applicable As String = Nothing
    Public Scheme_Code As String = Nothing
    Public Scheme_Item As String = Nothing
    Public Item_Tax As Double = 0
    Public Total_MRP_Amt As Double = 0
    Public Total_Basic_Amt As Double = 0
    Public Total_Disc_Amt As Double = 0
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
    Public Price_Date As String = Nothing

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
    Public PrincipleCode As String = Nothing
    Public REF_PK_ID As Integer
    Public PrincipleDesc As String = Nothing
    Public vendor_code As String = Nothing
    Public vendor_desc As String = Nothing
    Public Bin_No As String = Nothing
    Public HeadDiscPer As Double = 0
    Public HeadDiscPerAmt As Double = 0
    Public Transporter_Commission_Rate As Double = 0
    Public Transporter_Commission_Amt As Double = 0
    Public REF_TPT_PK_ID As Integer = 0
    Public arrBatchItem As List(Of clsBatchInventory) = Nothing
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsMCCMaterialSaleDetail), ByVal trans As SqlTransaction, ByVal DocDate As DateTime) As Boolean

        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then

            For Each obj As clsMCCMaterialSaleDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "ItemwiseTaxCode", obj.ItemwiseTaxCode)
                clsCommon.AddColumnsForChange(coll, "Commission_Rate", obj.Commission_Rate)
                clsCommon.AddColumnsForChange(coll, "Commission_Party", obj.Commission_Party)
                clsCommon.AddColumnsForChange(coll, "Commission_Amt", obj.Commission_Amt)
                clsCommon.AddColumnsForChange(coll, "Amt_Less_Commission", obj.Amt_Less_Commission)
                clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Rate", obj.Transporter_Commission_Rate)
                clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Amt", obj.Transporter_Commission_Amt)
                clsCommon.AddColumnsForChange(coll, "OrgUnit_code", obj.OrgUnit_code)
                clsCommon.AddColumnsForChange(coll, "REF_TPT_PK_ID", obj.REF_TPT_PK_ID, True)
                clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                clsCommon.AddColumnsForChange(coll, "Row_Type", obj.Row_Type)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Bar_Code", obj.Bar_Code, True)
                clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)
                clsCommon.AddColumnsForChange(coll, "Billing_Qty", obj.Billing_Qty)
                clsCommon.AddColumnsForChange(coll, "Billing_Unit_code", obj.Billing_Unit_code)
                clsCommon.AddColumnsForChange(coll, "Free_qty", obj.Free_Qty)

                clsCommon.AddColumnsForChange(coll, "Order_Code", obj.Order_Code, True)

                clsCommon.AddColumnsForChange(coll, "Balance_Qty", obj.Balance_Qty)
                clsCommon.AddColumnsForChange(coll, "Unit_code", obj.Unit_code)
                clsCommon.AddColumnsForChange(coll, "Location", obj.Location)
                clsCommon.AddColumnsForChange(coll, "Item_Cost", obj.Item_Cost)

                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Disc_Per", obj.Disc_Per)
                clsCommon.AddColumnsForChange(coll, "Disc_Amt", obj.Disc_Amt)
                clsCommon.AddColumnsForChange(coll, "Disc_Per_Unit", obj.Disc_Per_Unit)
                clsCommon.AddColumnsForChange(coll, "Disc_Unit_Amt", obj.Disc_Unit_Amt)
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
                clsCommon.AddColumnsForChange(coll, "Batch_No", obj.Batch_No)

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
                Dim strSql As String = "select top 1 TSPL_ITEM_PRICE_MASTER.Price_Amount1 ,TSPL_ITEM_PRICE_MASTER.Price_Amount2 , " &
              "TSPL_ITEM_PRICE_MASTER.Price_Amount3 ,TSPL_ITEM_PRICE_MASTER.Price_Amount4 ,TSPL_ITEM_PRICE_MASTER.Price_Amount5 , " &
              "TSPL_ITEM_PRICE_MASTER.Price_Amount6 ,TSPL_ITEM_PRICE_MASTER.Price_Amount7 ,TSPL_ITEM_PRICE_MASTER.Price_Amount8 , " &
              "TSPL_ITEM_PRICE_MASTER.Price_Amount9 ,TSPL_ITEM_PRICE_MASTER.Price_Amount10 from TSPL_ITEM_PRICE_MASTER " &
             "where  TSPL_ITEM_PRICE_MASTER.Price_Code='" & obj.Price_code & "' and  TSPL_ITEM_PRICE_MASTER.Item_Code='" & obj.Item_Code & "' and  " &
             "TSPL_ITEM_PRICE_MASTER.Item_Basic_Net=" & obj.MRP & " and TSPL_ITEM_PRICE_MASTER.UOM='" & obj.Unit_code & "' "
                Dim dt = New DataTable()
                dt = clsDBFuncationality.GetDataTable(strSql, trans)
                If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                    obj.Price_Amount1 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount1"))
                    obj.Price_Amount2 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount2"))
                    obj.Price_Amount3 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount3"))
                    obj.Price_Amount4 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount4"))
                    obj.Price_Amount5 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount5"))
                    obj.Price_Amount6 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount6"))
                    obj.Price_Amount7 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount7"))
                    obj.Price_Amount8 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount8"))
                    obj.Price_Amount9 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount9"))
                    obj.Price_Amount10 = clsCommon.myCdbl(dt.Rows(0).Item("Price_Amount10"))
                End If

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
                If clsCommon.myLen(obj.Price_Date) > 0 Then
                    clsCommon.AddColumnsForChange(coll, "Price_Date", clsCommon.GetPrintDate(obj.Price_Date, "dd/MMM/yyyy"))
                End If
                clsCommon.AddColumnsForChange(coll, "Abatement_Per", obj.Abatement_Per)
                clsCommon.AddColumnsForChange(coll, "Abatement_Amt", obj.Abatement_Amt)
                clsCommon.AddColumnsForChange(coll, "FOC_Item", obj.FOC_Item)

                clsCommon.AddColumnsForChange(coll, "Item_Weight", obj.Item_Weight)
                clsCommon.AddColumnsForChange(coll, "Conv_Factor", obj.Conv_Factor)
                clsCommon.AddColumnsForChange(coll, "TotalItem_Weight", obj.TotalItem_Weight)
                clsCommon.AddColumnsForChange(coll, "Markup_On", obj.Markup_On)
                clsCommon.AddColumnsForChange(coll, "Markup_Percent", obj.Markup_Percent)
                clsCommon.AddColumnsForChange(coll, "Landing_Cost", obj.Landing_Cost)
                clsCommon.AddColumnsForChange(coll, "HeadDiscAmt", obj.HeadDiscAmt)
                clsCommon.AddColumnsForChange(coll, "CustDiscPer", obj.CustDiscPer)
                clsCommon.AddColumnsForChange(coll, "CasdDiscScheme_Code", obj.CasdDiscScheme_Code)
                clsCommon.AddColumnsForChange(coll, "Purchase_Cost", obj.Purchase_Cost)
                clsCommon.AddColumnsForChange(coll, "OrgRate", obj.OrgRate)
                clsCommon.AddColumnsForChange(coll, "PrincipleCode", obj.PrincipleCode)
                clsCommon.AddColumnsForChange(coll, "REF_PK_ID", obj.REF_PK_ID, True)
                clsCommon.AddColumnsForChange(coll, "PrincipleDesc", obj.PrincipleDesc)
                clsCommon.AddColumnsForChange(coll, "vendor_code", obj.vendor_code)
                clsCommon.AddColumnsForChange(coll, "vendor_desc", obj.vendor_desc)
                clsCommon.AddColumnsForChange(coll, "Bin_No", obj.Bin_No)
                clsCommon.AddColumnsForChange(coll, "HeadDiscPer", obj.HeadDiscPer)
                clsCommon.AddColumnsForChange(coll, "HeadDiscPerAmt", obj.HeadDiscPerAmt)
                'clsCommon.AddColumnsForChange(coll, "Unit_Cogs", clsItemLocationDetails.GetUnitCogs(obj.Item_Code, obj.Location, trans))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SHIPMENT_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                clsSerializeInvenotry.SaveData("SD-IN", strDocNo, DocDate, "O", obj.Item_Code, obj.Location, obj.Line_No, obj.arrSrItem, trans)
                clsBatchInventory.SaveData("MCC-MSALE", strDocNo, DocDate, "O", obj.Item_Code, obj.Location, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans) ' Change by Prabhakar

            Next
        End If
        Return True
    End Function

    Public Shared Function GetBalanceSRNQty(ByVal strSRNCode As String, ByVal strICode As String, ByVal strCurrPINNo As String, ByVal strUOM As String, ByVal dblMRP As Double, ByVal dblAssessable As Double) As Double
        Dim qry As String = "select SUM(qty * RI) as Balance from(  " &
            " select TSPL_SD_SHIPMENT_DETAIL.Item_Code as ICode,TSPL_SD_SHIPMENT_DETAIL.Qty as Qty,1 as RI from TSPL_SD_SHIPMENT_DETAIL left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code where TSPL_SD_SHIPMENT_DETAIL.Status=0 and TSPL_SD_SHIPMENT_HEAD.Status=1 and TSPL_SD_SHIPMENT_DETAIL.Document_Code ='" + strSRNCode + "' and TSPL_SD_SHIPMENT_DETAIL.Item_Code='" + strICode + "' and  TSPL_SD_SHIPMENT_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_SD_SHIPMENT_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_SD_SHIPMENT_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' " &
            " union all " &
            " select TSPL_PI_DETAIL.Item_Code as ICode,TSPL_PI_DETAIL.PI_Qty as Qty,-1 as RI from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No where TSPL_PI_DETAIL.SRN_Id='" + strSRNCode + "'   and TSPL_PI_DETAIL.Item_Code='" + strICode + "'  and  TSPL_PI_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_PI_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_PI_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'  and TSPL_PI_DETAIL.PI_No not in ('" + strCurrPINNo + "')  " &
            " )Final "
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
    End Function

    Public Shared Function CompleteSRN(ByVal strDoccode As String, ByVal strICode As String, ByVal LineNo As Integer) As Boolean
        Dim qry As String = "update TSPL_SD_SHIPMENT_DETAIL set Status=1 where Document_Code='" + strDoccode + "' and Line_No='" + clsCommon.myCstr(LineNo) + "' and Item_Code='" + strICode + "'"
        Return clsDBFuncationality.ExecuteNonQuery(qry)
    End Function
End Class
