Imports System.Data.SqlClient
Imports System.IO
Public Class clsDCSSaleEntry
#Region "Variables"
    Public Is_CashSale As String = "N"
    Public Total_Comm_Amt As Double = 0
    Public SO_Validity As Integer = 0
    Public Commission_Apply As Integer
    Public Dispatch_date As Date?
    Public Dispatch_Terms As String = Nothing
    Public Payment_Terms As String = Nothing
    Public Dispatch_Period As Integer = 0
    Public Vehicle_Capacity As Integer = 0
    Public Road_Permit_No As String = Nothing
    Public Bank_Code As String = Nothing
    Public Direct_Dispatch As Integer = 0
    Public Is_Delivered As Integer = 0
    Public Podate As DateTime? = Nothing
    Public Form_38_No As String = Nothing
    Public Cust_PO_No As String = Nothing
    Public Price_Group_Code As String = Nothing
    Public Document_Code As String = Nothing
    Public isTaxExempted As Boolean = False '' Not a table field
    Public Document_Date As DateTime? = Nothing
    Public Customer_Code As String = Nothing
    Public Customer_Name As String = Nothing  'Not a table field
    Public Is_Apply_TPT As Boolean = False
    Public TPT_Vendor As String = Nothing
    Public Exclude_KKF_And_Mandi As Boolean = False
    Public Recommended_By As String = Nothing
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Rate_Status As Integer = 1
    Public On_Hold As Boolean = Nothing
    Public Ref_No As String = Nothing
    Public Description As String = Nothing
    Public Remarks As String = Nothing
    Public Bill_To_Location As String = Nothing
    Public Sub_Location_code As String = Nothing
    Public BillToLocationName As String = Nothing 'Not a table field
    Public Ship_To_Location As String = Nothing
    Public ShipToLocationName As String = Nothing 'Not a table field
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
    'Public Inv_No As String = Nothing
    Public Against_Sales_Order As String = Nothing
    Public Is_Internal As Boolean = False
    Public Tax_Calculation_Type As EnumTaxCalucationType

    Public Is_Create_Auto_Receipt As Boolean = False
    Public Against_POS As String = Nothing

    Public Salesman_Code As String = Nothing
    Public Salesman_Name As String = Nothing
    Public Arr As List(Of clsDCSSaleEntryDetail) = Nothing

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public PROJECT_ID As String = Nothing
    Public No_Of_Instalment As Integer
    Public IS_TCS As Integer = 0
    Public Price_Code As String = Nothing
    Public DCS_Price_Code As String = Nothing
    Public Sale_Route_Code As String = Nothing
    Public Supply_Date As DateTime? = Nothing
    Public Shift_Type As String = Nothing
    Public Route_No As String = Nothing
    Public Route_Desc As String = Nothing
    Public HeadDisc_Per As Double = 0
    Public HeadDisc_Amt As Double = 0
    Public RateDiff_Per As Double = 0
    Public RateDiff_Amt As Double = 0
    Public Gross_Amount As Double = 0
    Public TotCashDiscAmt As Double = 0
    Public HeadDisc_PerAmt As Double = 0
    Public Receipt_No As String = ""
    Public ReceiptAmt As Double = 0
    Public ReceiverName As String = ""
    Public TotalSubsidyAmt As Double = 0
    Public TotalSubsidyDisAmt As Double = 0
    Public Transporter_Commission_TotalAmt As Double = 0


#End Region
    Public Function SaveData(ByVal obj As clsDCSSaleEntry, ByVal isNewEntry As Boolean) As Boolean
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
    Private Function AllowToSave(ByVal obj As clsDCSSaleEntry, ByVal trans As SqlTransaction) As Boolean
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
                    "select TSPL_DCS_SALE_ENTRY.Document_Code as DOC_CODE,TSPL_DCS_SALE_ENTRY.Total_Amt as Amount,-1 as RI,0 as chk  from TSPL_DCS_SALE_ENTRY where Trans_Type='MCC' and   TSPL_DCS_SALE_ENTRY.Customer_Code='" + obj.Customer_Code + "' " + Environment.NewLine +
                    "and TSPL_DCS_SALE_ENTRY.Document_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_DCS_SALE_ENTRY.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_DCS_SALE_ENTRY.Document_Code not in ('" + obj.Document_Code + "')" + Environment.NewLine +
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

        '   qry = "select 1 from TSPL_ITEM_MASTER where TSPL_ITEM_MASTER.Deduction <>'" + obj.Deduction + "' and Item_Code in ("
        For Each objtr As clsDCSSaleEntryDetail In obj.Arr
            If clsCommon.CompairString(objtr.Row_Type, "Item") = CompairStringResult.Equal Then
                qry += "'" + objtr.Item_Code + "'"
            End If
        Next
        ' qry += ")"
        'dt = clsDBFuncationality.GetDataTable(qry, trans)
        'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '    Throw New Exception("Deduction of All item should be [" + obj.Deduction + "]")
        'End If
        Return True
    End Function
    Public Function SaveData(ByVal obj As clsDCSSaleEntry, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmDCSSaleEntry, obj.Bill_To_Location, obj.Document_Date, trans)
            AllowToSave(obj, trans)
            clsSerializeInvenotry.DeleteData("DCS-SAL-ENT", obj.Document_Code, trans)
            'If Not isNewEntry Then
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DCS_SALE_ENTRY", "Document_Code", "TSPL_DCS_SALE_ENTRY_DETAIL", "Document_Code", trans)
            'End If

            Dim qry As String = "select distinct isnull(TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,'') DOCUMENT_CODE from TSPL_SD_SHIPMENT_DETAIL where REF_PK_ID in (select  PK_ID from TSPL_DCS_SALE_ENTRY_DETAIL where Document_Code='" + obj.Document_Code + "') "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    isSaved = isSaved AndAlso clsMCCMaterialSale.DeleteData(clsCommon.myCstr(dr("DOCUMENT_CODE")), trans)
                Next
            End If
            qry = "delete from TSPL_DCS_SALE_ENTRY_DETAIL where Document_Code='" + obj.Document_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsBatchInventory.DeleteData("DCS-SAL-ENT", obj.Document_Code, trans)
            Dim strDocNo As String = ""
            If isNewEntry Then
                obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DCSSaleEntry, "", obj.Bill_To_Location)
            End If
            If (clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            If clsCommon.myLen(obj.Podate) <= 0 Then
                clsCommon.AddColumnsForChange(coll, "cust_po_date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            Else
                clsCommon.AddColumnsForChange(coll, "cust_po_date", clsCommon.GetPrintDate(obj.Podate, "dd/MMM/yyyy hh:mm tt"))
            End If
            '---------------------------------------------------------------------------------------
            clsCommon.AddColumnsForChange(coll, "Customer_Code", obj.Customer_Code)
            clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Internal", IIf(obj.Is_Internal, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Ref_No", obj.Ref_No)
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            'clsCommon.AddColumnsForChange(coll, "Inv_No", obj.Inv_No)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Is_CashSale", obj.Is_CashSale)
            clsCommon.AddColumnsForChange(coll, "Bill_To_Location", obj.Bill_To_Location)
            clsCommon.AddColumnsForChange(coll, "Is_Apply_TPT", IIf(obj.Is_Apply_TPT, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Exclude_KKF_And_Mandi", IIf(obj.Exclude_KKF_And_Mandi, 1, 0))
            clsCommon.AddColumnsForChange(coll, "TPT_Vendor", obj.TPT_Vendor, True)
            clsCommon.AddColumnsForChange(coll, "Recommended_By", obj.Recommended_By, True)
            clsCommon.AddColumnsForChange(coll, "Sub_Location_code", obj.Sub_Location_code)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Location", obj.Ship_To_Location)
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

            clsCommon.AddColumnsForChange(coll, "ParentDocNo", obj.Document_Code)
            clsCommon.AddColumnsForChange(coll, "Is_Create_Auto_Receipt", IIf(obj.Is_Create_Auto_Receipt, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Salesman_Code", obj.Salesman_Code, True)
            clsCommon.AddColumnsForChange(coll, "Salesman_Name", obj.Salesman_Name)
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", obj.ApplicableFrom, True)
            clsCommon.AddColumnsForChange(coll, "Price_Code", obj.Price_Code)
            clsCommon.AddColumnsForChange(coll, "DCS_Price_Code", obj.DCS_Price_Code, True)
            '  clsCommon.AddColumnsForChange(coll, "Sale_Route_Code", obj.Sale_Route_Code, True)
            If obj.Supply_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Supply_Date", clsCommon.GetPrintDate(obj.Supply_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Supply_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
            End If
            clsCommon.AddColumnsForChange(coll, "Shift_type", obj.Shift_Type, True)
            clsCommon.AddColumnsForChange(coll, "Route_No", obj.Route_No)
            clsCommon.AddColumnsForChange(coll, "Route_Desc", obj.Route_Desc)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Per", obj.HeadDisc_Per)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_PerAmt", obj.HeadDisc_PerAmt)
            clsCommon.AddColumnsForChange(coll, "HeadDisc_Amt", obj.HeadDisc_Amt)
            clsCommon.AddColumnsForChange(coll, "RateDiff_Per", obj.RateDiff_Per)
            clsCommon.AddColumnsForChange(coll, "RateDiff_Amt", obj.RateDiff_Amt)
            clsCommon.AddColumnsForChange(coll, "Gross_Amount", obj.Gross_Amount)
            clsCommon.AddColumnsForChange(coll, "TotCashDiscAmt", obj.TotCashDiscAmt)
            clsCommon.AddColumnsForChange(coll, "Price_Group_Code", obj.Price_Group_Code)
            clsCommon.AddColumnsForChange(coll, "Cust_PO_No", obj.Cust_PO_No)
            clsCommon.AddColumnsForChange(coll, "Form_38_No", obj.Form_38_No)
            clsCommon.AddColumnsForChange(coll, "SO_Validity", obj.SO_Validity)
            clsCommon.AddColumnsForChange(coll, "Commission_Apply", obj.Commission_Apply)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Capacity", obj.Vehicle_Capacity)
            clsCommon.AddColumnsForChange(coll, "Dispatch_Terms", obj.Dispatch_Terms)
            clsCommon.AddColumnsForChange(coll, "IS_TCS", obj.IS_TCS)
            clsCommon.AddColumnsForChange(coll, "Payment_Terms", obj.Payment_Terms)
            clsCommon.AddColumnsForChange(coll, "Dispatch_Period", obj.Dispatch_Period)
            clsCommon.AddColumnsForChange(coll, "Road_Permit_No", obj.Road_Permit_No)
            clsCommon.AddColumnsForChange(coll, "Bank_Code", obj.Bank_Code, True)
            clsCommon.AddColumnsForChange(coll, "Total_Comm_Amt", obj.Total_Comm_Amt)
            clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
            clsCommon.AddColumnsForChange(coll, "Receipt_No", obj.Receipt_No, True)
            clsCommon.AddColumnsForChange(coll, "ReceiptAmt", obj.ReceiptAmt, True)
            clsCommon.AddColumnsForChange(coll, "ReceiverName", obj.ReceiverName, True)
            clsCommon.AddColumnsForChange(coll, "TotalSubsidyAmt", obj.TotalSubsidyAmt)
            clsCommon.AddColumnsForChange(coll, "TotalSubsidyDisAmt", obj.TotalSubsidyDisAmt)
            clsCommon.AddColumnsForChange(coll, "Transporter_Commission_TotalAmt", obj.Transporter_Commission_TotalAmt)

            If clsCommon.myLen(obj.Against_Sales_Order) = 0 Then
                obj.Direct_Dispatch = 1
            End If
            clsCommon.AddColumnsForChange(coll, "Direct_Dispatch", obj.Direct_Dispatch)
            clsCommon.AddColumnsForChange(coll, "No_Of_Instalment", obj.No_Of_Instalment)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Trans_Type", "MCC")
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_SALE_ENTRY", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_SALE_ENTRY", OMInsertOrUpdate.Update, "TSPL_DCS_SALE_ENTRY.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            isSaved = isSaved AndAlso clsDCSSaleEntryDetail.SaveData(obj.Document_Code, Arr, trans, obj.Document_Date)

            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_Code, obj.arrCustomFields, trans)
            qry = "update TSPL_DCS_SALE_ENTRY_DETAIL set Weight_UOM= (select Weight_UOM from TSPL_ITEM_MASTER where Item_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code)  where Document_Code='" + obj.Document_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            isSaved = isSaved AndAlso clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "Document_Code", obj.Document_Code, "TSPL_DCS_SALE_ENTRY", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DCS_SALE_ENTRY", "Document_Code", "TSPL_DCS_SALE_ENTRY_DETAIL", "Document_Code", trans)

        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function
    Public Shared Function checkSaveNotification(ByVal obj As clsDCSSaleEntry, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim Count As Integer = 0
            Dim CreditLimit As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Credit_Limit from TSPL_CUSTOMER_MASTER WHERE Cust_Code='" + obj.Customer_Code + "'", trans))
            Dim qry As String
            Dim dt As DataTable = clsScreenNotificationSchedule.GetScreenNotificationInfo(clsUserMgtCode.frmSNShipment, trans)
            For Each dr As DataRow In dt.Rows
                If clsCommon.CompairString(dr("Criteria"), "Credit days") = CompairStringResult.Equal Then
                    qry = "Select COUNT(*) from TSPL_DCS_SALE_ENTRY" &
        " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_DCS_SALE_ENTRY.Document_Code" &
        " LEFT OUTER JOIN TSPL_Customer_Invoice_Head ON TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code" &
        " WHERE TSPL_DCS_SALE_ENTRY.Status = 1" &
        " AND TSPL_DCS_SALE_ENTRY.Customer_Code='" + obj.Customer_Code + "'" &
        " AND TSPL_DCS_SALE_ENTRY.Due_Date<'" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "'" &
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
                    qry = "Select SUM(TSPL_Customer_Invoice_Head.Balance_Amt) from TSPL_DCS_SALE_ENTRY" &
        " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_DCS_SALE_ENTRY.Document_Code" &
        " LEFT OUTER JOIN TSPL_Customer_Invoice_Head ON TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code" &
        " WHERE TSPL_DCS_SALE_ENTRY.Status = 1" &
        " AND TSPL_DCS_SALE_ENTRY.Customer_Code='" + obj.Customer_Code + "'" &
        " AND TSPL_DCS_SALE_ENTRY.Document_Date<'" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "'" &
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
    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType) As clsDCSSaleEntry
        Return GetData(strDocumentNo, NavType, Nothing)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsDCSSaleEntry
        Dim obj As clsDCSSaleEntry = Nothing
        Dim qry As String = "SELECT TSPL_DCS_SALE_ENTRY.No_Of_Instalment,TSPL_DCS_SALE_ENTRY.IS_TCS,TSPL_DCS_SALE_ENTRY.Road_Permit_No,TSPL_DCS_SALE_ENTRY.Bank_Code,TSPL_DCS_SALE_ENTRY.Is_Delivered,TSPL_DCS_SALE_ENTRY.HeadDisc_PerAmt,TSPL_DCS_SALE_ENTRY.RateDiff_Per,TSPL_DCS_SALE_ENTRY.Gross_Amount,TSPL_DCS_SALE_ENTRY.RateDiff_Amt,TSPL_DCS_SALE_ENTRY.cust_po_date,TSPL_DCS_SALE_ENTRY.Cust_PO_No,TSPL_DCS_SALE_ENTRY.Vehicle_Code,TSPL_DCS_SALE_ENTRY.price_group_code,TSPL_DCS_SALE_ENTRY.HeadDisc_Per,TSPL_DCS_SALE_ENTRY.HeadDisc_Amt,TSPL_DCS_SALE_ENTRY.TotCashDiscAmt,TSPL_DCS_SALE_ENTRY.Route_No,TSPL_DCS_SALE_ENTRY.Route_Desc,TSPL_DCS_SALE_ENTRY.Price_Code,TSPL_DCS_SALE_ENTRY.DCS_Price_Code,TSPL_DCS_SALE_ENTRY.Sale_Route_Code,TSPL_DCS_SALE_ENTRY.Supply_Date,TSPL_DCS_SALE_ENTRY.Shift_Type,TSPL_DCS_SALE_ENTRY.Document_Code,TSPL_DCS_SALE_ENTRY.Document_Date,TSPL_DCS_SALE_ENTRY.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DCS_SALE_ENTRY.Status,TSPL_DCS_SALE_ENTRY.On_Hold,TSPL_DCS_SALE_ENTRY.Ref_No,TSPL_DCS_SALE_ENTRY.Description,TSPL_DCS_SALE_ENTRY.Is_Apply_TPT,TSPL_DCS_SALE_ENTRY.Exclude_KKF_And_Mandi,TSPL_DCS_SALE_ENTRY.TPT_Vendor,TSPL_DCS_SALE_ENTRY.Recommended_By,TSPL_DCS_SALE_ENTRY.Is_CashSale,TSPL_DCS_SALE_ENTRY.Remarks,TSPL_DCS_SALE_ENTRY.Bill_To_Location,TSPL_DCS_SALE_ENTRY.Sub_Location_code,TSPL_DCS_SALE_ENTRY.Ship_To_Location,TSPL_DCS_SALE_ENTRY.TAX1,TSPL_DCS_SALE_ENTRY.TAX1_Rate,TSPL_DCS_SALE_ENTRY.TAX1_Amt,TSPL_DCS_SALE_ENTRY.TAX1_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX2,TSPL_DCS_SALE_ENTRY.TAX2_Rate,TSPL_DCS_SALE_ENTRY.TAX2_Amt,TSPL_DCS_SALE_ENTRY.TAX2_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX3,TSPL_DCS_SALE_ENTRY.TAX3_Rate,TSPL_DCS_SALE_ENTRY.TAX3_Amt,TSPL_DCS_SALE_ENTRY.TAX3_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX4,TSPL_DCS_SALE_ENTRY.TAX4_Rate,TSPL_DCS_SALE_ENTRY.TAX4_Amt,TSPL_DCS_SALE_ENTRY.TAX4_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX5,TSPL_DCS_SALE_ENTRY.TAX5_Rate,TSPL_DCS_SALE_ENTRY.TAX5_Amt,TSPL_DCS_SALE_ENTRY.TAX5_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX6,TSPL_DCS_SALE_ENTRY.TAX6_Rate,TSPL_DCS_SALE_ENTRY.TAX6_Amt,TSPL_DCS_SALE_ENTRY.TAX6_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX7,TSPL_DCS_SALE_ENTRY.TAX7_Rate,TSPL_DCS_SALE_ENTRY.TAX7_Amt,TSPL_DCS_SALE_ENTRY.TAX7_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX8,TSPL_DCS_SALE_ENTRY.TAX8_Rate,TSPL_DCS_SALE_ENTRY.TAX8_Amt,TSPL_DCS_SALE_ENTRY.TAX8_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX9,TSPL_DCS_SALE_ENTRY.TAX9_Rate,TSPL_DCS_SALE_ENTRY.TAX9_Amt,TSPL_DCS_SALE_ENTRY.TAX9_Base_Amt,TSPL_DCS_SALE_ENTRY.TAX10,TSPL_DCS_SALE_ENTRY.TAX10_Rate,TSPL_DCS_SALE_ENTRY.TAX10_Amt,TSPL_DCS_SALE_ENTRY.TAX10_Base_Amt,TSPL_DCS_SALE_ENTRY.Discount_Base,TSPL_DCS_SALE_ENTRY.Discount_Amt,TSPL_DCS_SALE_ENTRY.Amount_Less_Discount,TSPL_DCS_SALE_ENTRY.Total_Tax_Amt,TSPL_DCS_SALE_ENTRY.Comments,TSPL_DCS_SALE_ENTRY.Comp_Code,TSPL_DCS_SALE_ENTRY.Terms_Code,TSPL_DCS_SALE_ENTRY.Due_Date ,TSPL_LOCATION_MASTER.Location_Desc as BillToLocationName,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as ShipToLocationName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_DCS_SALE_ENTRY.Posting_Date,TSPL_DCS_SALE_ENTRY.Total_Amt,TSPL_DCS_SALE_ENTRY.Carrier,TSPL_DCS_SALE_ENTRY.VehicleNo,TSPL_DCS_SALE_ENTRY.GRNo,TSPL_DCS_SALE_ENTRY.GENo,TSPL_DCS_SALE_ENTRY.GEDate, TSPL_DCS_SALE_ENTRY.Dept,TSPL_DCS_SALE_ENTRY.Dept_Desc,TSPL_DCS_SALE_ENTRY.Item_Type,TSPL_DCS_SALE_ENTRY.Against_Sales_Order ,TSPL_DCS_SALE_ENTRY.Against_Sales_Order,TSPL_DCS_SALE_ENTRY.Add_Charge_Code1,TSPL_DCS_SALE_ENTRY.Add_Charge_Name1,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt1,TSPL_DCS_SALE_ENTRY.Add_Charge_Code2,TSPL_DCS_SALE_ENTRY.Add_Charge_Name2,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt2,TSPL_DCS_SALE_ENTRY.Add_Charge_Code3,TSPL_DCS_SALE_ENTRY.Add_Charge_Name3,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt3,TSPL_DCS_SALE_ENTRY.Add_Charge_Code4,TSPL_DCS_SALE_ENTRY.Add_Charge_Name4,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt4,TSPL_DCS_SALE_ENTRY.Add_Charge_Code5,TSPL_DCS_SALE_ENTRY.Add_Charge_Name5,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt5,TSPL_DCS_SALE_ENTRY.Add_Charge_Code6,TSPL_DCS_SALE_ENTRY.Add_Charge_Name6,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt6,TSPL_DCS_SALE_ENTRY.Add_Charge_Code7,TSPL_DCS_SALE_ENTRY.Add_Charge_Name7,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt7,TSPL_DCS_SALE_ENTRY.Add_Charge_Code8,TSPL_DCS_SALE_ENTRY.Add_Charge_Name8,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt8,TSPL_DCS_SALE_ENTRY.Add_Charge_Code9 ,TSPL_DCS_SALE_ENTRY.Add_Charge_Name9,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt9 ,TSPL_DCS_SALE_ENTRY.Add_Charge_Code10 ,TSPL_DCS_SALE_ENTRY.Add_Charge_Name10,TSPL_DCS_SALE_ENTRY.Add_Charge_Amt10,TSPL_DCS_SALE_ENTRY.Total_Add_Charge,TSPL_DCS_SALE_ENTRY.Tax_Calculation_Type,TSPL_DCS_SALE_ENTRY.Challan_No, TSPL_DCS_SALE_ENTRY.Challan_Date,  TSPL_DCS_SALE_ENTRY.Is_Internal,TSPL_DCS_SALE_ENTRY.Is_Create_Auto_Receipt,TSPL_DCS_SALE_ENTRY.Salesman_Code ,TSPL_DCS_SALE_ENTRY.Salesman_Name,  "
        qry += " TSPL_DCS_SALE_ENTRY.CURRENCY_CODE,TSPL_DCS_SALE_ENTRY.CONVRATE,TSPL_DCS_SALE_ENTRY.APPLICABLEFROM,TSPL_DCS_SALE_ENTRY.PRoject_ID ,TSPL_DCS_SALE_ENTRY.Mannual_Invoice_No,TSPL_DCS_SALE_ENTRY. Mannual_Invoice_No_StringType,TSPL_DCS_SALE_ENTRY.Form_38_No " &
        " ,TSPL_DCS_SALE_ENTRY.SO_Validity,TSPL_DCS_SALE_ENTRY.Commission_Apply,TSPL_DCS_SALE_ENTRY.Total_Comm_Amt,TSPL_DCS_SALE_ENTRY.Dispatch_date" &
        " ,TSPL_DCS_SALE_ENTRY.Dispatch_Terms,TSPL_DCS_SALE_ENTRY.Payment_Terms,TSPL_DCS_SALE_ENTRY.Dispatch_Period,TSPL_DCS_SALE_ENTRY.Vehicle_Capacity,TSPL_DCS_SALE_ENTRY.RoundOffAmount,TSPL_DCS_SALE_ENTRY.Is_Taxable " &
        ",TSPL_DCS_SALE_ENTRY.Receipt_No,TSPL_DCS_SALE_ENTRY.ReceiptAmt,TSPL_DCS_SALE_ENTRY.VehicleNo,TSPL_DCS_SALE_ENTRY.ReceiverName,TSPL_DCS_SALE_ENTRY.TotalSubsidyAmt ,TSPL_DCS_SALE_ENTRY.TotalSubsidyDisAmt,TSPL_DCS_SALE_ENTRY.Transporter_Commission_TotalAmt,TSPL_DCS_SALE_ENTRY.No_Of_Instalment "
        qry += "  FROM TSPL_DCS_SALE_ENTRY "
        qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_DCS_SALE_ENTRY.Bill_To_Location "
        qry += " left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_DCS_SALE_ENTRY.Ship_To_Location "
        qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_DCS_SALE_ENTRY.Terms_Code "
        qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DCS_SALE_ENTRY.Customer_Code where 2=2"
        Dim whrCls As String = ""
        Dim strwherecls As String = ""
        If clsCommon.CompairString(clsCommon.myCstr(NavType).ToUpper(), "CURRENT") <> CompairStringResult.Equal Then
            strwherecls = FrmMainTranScreen.CustomerPermission()
        End If
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrCls = " AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        'End If

        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 And clsCommon.myLen(strwherecls) > 0 Then
            whrCls = "  and TSPL_DCS_SALE_ENTRY.Trans_Type='MCC' AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and TSPL_DCS_SALE_ENTRY.Customer_Code in (" + strwherecls + ") "
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrCls = "  and TSPL_DCS_SALE_ENTRY.Trans_Type='MCC' AND Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        ElseIf clsCommon.myLen(strwherecls) > 0 Then
            whrCls = "  and TSPL_DCS_SALE_ENTRY.Trans_Type='MCC' AND TSPL_DCS_SALE_ENTRY.Customer_Code in (" + strwherecls + ")"
        Else
            whrCls = " and TSPL_DCS_SALE_ENTRY.Trans_Type='MCC' "
        End If
        '-----------------------------------------------------

        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_DCS_SALE_ENTRY.Document_Code = (select MIN(Document_Code) from TSPL_DCS_SALE_ENTRY WHERE 1=1 " + whrCls + ") "
            Case NavigatorType.Last
                qry += " and TSPL_DCS_SALE_ENTRY.Document_Code = (select Max(Document_Code) from TSPL_DCS_SALE_ENTRY WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Current
                qry += " and TSPL_DCS_SALE_ENTRY.Document_Code = '" + strPONo + "' "
            Case NavigatorType.Next
                qry += " and TSPL_DCS_SALE_ENTRY.Document_Code = (select Min(Document_Code) from TSPL_DCS_SALE_ENTRY where Document_Code>'" + strPONo + "' " + whrCls + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_DCS_SALE_ENTRY.Document_Code = (select Max(Document_Code) from TSPL_DCS_SALE_ENTRY where Document_Code<'" + strPONo + "' " + whrCls + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsDCSSaleEntry()
            obj.SO_Validity = clsCommon.myCdbl(dt.Rows(0)("SO_Validity"))
            obj.Commission_Apply = clsCommon.myCdbl(dt.Rows(0)("Commission_Apply"))
            obj.Total_Comm_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Comm_Amt"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            If dt.Rows(0)("Dispatch_date") IsNot DBNull.Value Then
                obj.Dispatch_date = clsCommon.GetPrintDate((dt.Rows(0)("Dispatch_date")), "dd/MMM/yyyy")
            End If
            obj.No_Of_Instalment = clsCommon.myCdbl(dt.Rows(0)("No_Of_Instalment"))
            obj.Vehicle_Capacity = clsCommon.myCdbl(dt.Rows(0)("Vehicle_Capacity"))
            obj.Dispatch_Terms = clsCommon.myCstr(dt.Rows(0)("Dispatch_Terms"))
            obj.Payment_Terms = clsCommon.myCstr(dt.Rows(0)("Payment_Terms"))
            obj.Dispatch_Period = clsCommon.myCdbl(dt.Rows(0)("Dispatch_Period"))
            obj.Road_Permit_No = clsCommon.myCstr(dt.Rows(0)("Road_Permit_No"))
            obj.Bank_Code = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
            obj.IS_TCS = clsCommon.myCstr(dt.Rows(0)("IS_TCS"))
            obj.Is_Delivered = clsCommon.myCdbl(dt.Rows(0)("Is_Delivered"))
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Customer_Code = clsCommon.myCstr(dt.Rows(0)("Customer_Code"))
            obj.Customer_Name = clsCommon.myCstr(dt.Rows(0)("Customer_Name"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Is_Internal = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Internal")) = 1, True, False)
            obj.Ref_No = clsCommon.myCstr(dt.Rows(0)("Ref_No"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Is_CashSale = clsCommon.myCstr(dt.Rows(0)("Is_CashSale"))
            obj.Is_Apply_TPT = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Apply_TPT")) = 1, True, False)
            obj.Exclude_KKF_And_Mandi = IIf(clsCommon.myCdbl(dt.Rows(0)("Exclude_KKF_And_Mandi")) = 1, True, False)
            obj.TPT_Vendor = clsCommon.myCstr(dt.Rows(0)("TPT_Vendor"))
            obj.Recommended_By = clsCommon.myCstr(dt.Rows(0)("Recommended_By"))
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
            '--------------------------------------------------------------------------------------

            obj.BillToLocationName = clsCommon.myCstr(dt.Rows(0)("BillToLocationName"))
            obj.ShipToLocationName = clsCommon.myCstr(dt.Rows(0)("ShipToLocationName"))
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
            'obj.Inv_No = clsCommon.myCstr(dt.Rows(0)("Inv_No"))
            If dt.Rows(0)("Challan_Date") IsNot DBNull.Value Then

                obj.Challan_Date = clsCommon.GetPrintDate((dt.Rows(0)("Challan_Date")), "dd/MMM/yyyy")
            End If
            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)

            obj.Salesman_Code = clsCommon.myCstr(dt.Rows(0)("Salesman_Code"))
            obj.Salesman_Name = clsCommon.myCstr(dt.Rows(0)("Salesman_Name"))
            obj.Price_Code = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
            obj.DCS_Price_Code = clsCommon.myCstr(dt.Rows(0)("DCS_Price_Code"))
            obj.Sale_Route_Code = clsCommon.myCstr(dt.Rows(0)("Sale_Route_Code"))
            If dt.Rows(0)("Supply_Date") IsNot DBNull.Value Then
                obj.Supply_Date = clsCommon.myCstr(dt.Rows(0)("Supply_Date"))
            End If
            obj.Shift_Type = clsCommon.myCstr(dt.Rows(0)("Shift_Type"))
            obj.Route_No = clsCommon.myCstr(dt.Rows(0)("Route_No"))
            obj.Route_Desc = clsCommon.myCstr(dt.Rows(0)("Route_Desc"))
            obj.HeadDisc_Per = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Per"))
            obj.HeadDisc_PerAmt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_PerAmt"))
            obj.HeadDisc_Amt = clsCommon.myCdbl(dt.Rows(0)("HeadDisc_Amt"))
            obj.RateDiff_Per = clsCommon.myCdbl(dt.Rows(0)("RateDiff_Per"))
            obj.Gross_Amount = clsCommon.myCdbl(dt.Rows(0)("Gross_Amount"))
            obj.RateDiff_Amt = clsCommon.myCdbl(dt.Rows(0)("RateDiff_Amt"))
            obj.TotCashDiscAmt = clsCommon.myCdbl(dt.Rows(0)("TotCashDiscAmt"))
            obj.Price_Group_Code = clsCommon.myCstr(dt.Rows(0)("Price_Group_Code"))
            obj.Cust_PO_No = clsCommon.myCstr(dt.Rows(0)("Cust_PO_No"))
            If dt.Rows(0)("cust_po_date") IsNot DBNull.Value Then
                obj.Podate = clsCommon.myCDate(dt.Rows(0)("cust_po_date"))
            End If
            obj.Form_38_No = clsCommon.myCstr(dt.Rows(0)("Form_38_No"))
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            qry = "SELECT TSPL_DCS_SALE_ENTRY_DETAIL.PK_ID, TSPL_DCS_SALE_ENTRY_DETAIL.ItemwiseTaxCode,TSPL_DCS_SALE_ENTRY_DETAIL.OrgUnit_code,TSPL_DCS_SALE_ENTRY_DETAIL.Is_Mannual_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Document_Code,TSPL_DCS_SALE_ENTRY_DETAIL.Line_No, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Status,TSPL_DCS_SALE_ENTRY_DETAIL.Row_Type,TSPL_DCS_SALE_ENTRY_DETAIL.status,TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code, " &
            "TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_MASTER.DCS_Sale_Zero_Cost,0)as DCS_Sale_Zero_Cost,TSPL_DCS_SALE_ENTRY_DETAIL.Qty,TSPL_DCS_SALE_ENTRY_DETAIL.Free_Qty,TSPL_DCS_SALE_ENTRY_DETAIL.Order_Code, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Order_Code,TSPL_DCS_SALE_ENTRY_DETAIL.Balance_Qty,TSPL_DCS_SALE_ENTRY_DETAIL.Unit_code, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Location,TSPL_DCS_SALE_ENTRY_DETAIL.Item_Cost,TSPL_DCS_SALE_ENTRY_DETAIL.TAX1,TSPL_DCS_SALE_ENTRY_DETAIL.TAX1_Rate, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX1_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX2,TSPL_DCS_SALE_ENTRY_DETAIL.TAX2_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX2_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX3,TSPL_DCS_SALE_ENTRY_DETAIL.TAX3_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX3_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX4 , " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX4_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX4_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX5,TSPL_DCS_SALE_ENTRY_DETAIL.TAX5_Rate , " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX5_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX6,TSPL_DCS_SALE_ENTRY_DETAIL.TAX6_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX6_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX7,TSPL_DCS_SALE_ENTRY_DETAIL.TAX7_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX7_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX8, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX8_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX8_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX9,TSPL_DCS_SALE_ENTRY_DETAIL.TAX9_Rate, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX9_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX10,TSPL_DCS_SALE_ENTRY_DETAIL.TAX10_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.TAX10_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Amount,TSPL_DCS_SALE_ENTRY_DETAIL.Gross_Amount,TSPL_DCS_SALE_ENTRY_DETAIL.TotalSubsidyDisAmt,TSPL_DCS_SALE_ENTRY_DETAIL.Transporter_Commission_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.Transporter_Commission_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.RoundOffAmount,TSPL_DCS_SALE_ENTRY_DETAIL.REF_TPT_PK_ID,TSPL_DCS_SALE_ENTRY_DETAIL.RateDiff_Per,TSPL_DCS_SALE_ENTRY_DETAIL.RateDiff_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TotalSubsidyAmt,TSPL_DCS_SALE_ENTRY_DETAIL.Disc_Per,TSPL_DCS_SALE_ENTRY_DETAIL.Disc_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Amt_Less_Discount, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Total_Tax_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Item_Net_Amt,TSPL_LOCATION_MASTER.Location_Desc as LocationName, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX1_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX2_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX3_Base_Amt , " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX4_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX5_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX6_Base_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX7_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX8_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.TAX9_Base_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.TAX10_Base_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.MRP,TSPL_DCS_SALE_ENTRY_DETAIL.Batch_No,TSPL_DCS_SALE_ENTRY_DETAIL.MFG_Date, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Expiry_Date,TSPL_DCS_SALE_ENTRY_DETAIL.Specification,TSPL_DCS_SALE_ENTRY_DETAIL.Remarks,TSPL_DCS_SALE_ENTRY_DETAIL.Assessable, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.AssessableAmt,TSPL_DCS_SALE_ENTRY_DETAIL.Bar_Code, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Scheme_Applicable,TSPL_DCS_SALE_ENTRY_DETAIL.Scheme_Code, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Scheme_Item,TSPL_DCS_SALE_ENTRY_DETAIL.Item_Tax,TSPL_DCS_SALE_ENTRY_DETAIL.Total_MRP_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Total_Basic_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Total_Disc_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Cust_Discount, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Total_Cust_Discount,TSPL_DCS_SALE_ENTRY_DETAIL.ActualRate,TSPL_DCS_SALE_ENTRY_DETAIL.Cust_DiscountQty, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.Price_code,TSPL_DCS_SALE_ENTRY_DETAIL.Abatement_Per,TSPL_DCS_SALE_ENTRY_DETAIL.Abatement_Amt, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.FOC_Item,TSPL_DCS_SALE_ENTRY_DETAIL.Item_Weight,TSPL_DCS_SALE_ENTRY_DETAIL.Price_Date, " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.HeadDiscPer,TSPL_DCS_SALE_ENTRY_DETAIL.HeadDiscPerAmt,TSPL_DCS_SALE_ENTRY_DETAIL.Is_Add_TPT,TSPL_DCS_SALE_ENTRY_DETAIL.Bin_No,TSPL_DCS_SALE_ENTRY_DETAIL.TotalItem_Weight,TSPL_DCS_SALE_ENTRY_DETAIL.Conv_Factor,TSPL_DCS_SALE_ENTRY_DETAIL.Purchase_Cost,TSPL_DCS_SALE_ENTRY_DETAIL.OrgRate,  " &
            "TSPL_DCS_SALE_ENTRY_DETAIL.vendor_code,TSPL_DCS_SALE_ENTRY_DETAIL.vendor_desc,TSPL_DCS_SALE_ENTRY_DETAIL.PrincipleCode,TSPL_DCS_SALE_ENTRY_DETAIL.PrincipleDesc,TSPL_DCS_SALE_ENTRY_DETAIL.Markup_On,TSPL_DCS_SALE_ENTRY_DETAIL.Markup_Percent,TSPL_DCS_SALE_ENTRY_DETAIL.Landing_Cost,TSPL_DCS_SALE_ENTRY_DETAIL.HeadDiscAmt,TSPL_DCS_SALE_ENTRY_DETAIL.CustDiscPer,TSPL_DCS_SALE_ENTRY_DETAIL.CasdDiscScheme_Code " &
            ",TSPL_DCS_SALE_ENTRY_DETAIL.Commission_Rate,TSPL_DCS_SALE_ENTRY_DETAIL.Commission_Party,TSPL_DCS_SALE_ENTRY_DETAIL.Commission_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Amt_Less_Commission,TSPL_DCS_SALE_ENTRY_DETAIL.Disc_Per_Unit,TSPL_DCS_SALE_ENTRY_DETAIL.Disc_Unit_Amt,TSPL_DCS_SALE_ENTRY_DETAIL.Deduction_Type ,TSPL_DCS_SALE_ENTRY_DETAIL.Deduction,TSPL_DEDUCTION_MASTER.Description as Deduction_Name ,TSPL_DCS_SALE_ENTRY_DETAIL.Tax_Group,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as TaxGroupName "
            qry += " FROM TSPL_DCS_SALE_ENTRY_DETAIL "
            qry += "  Left outer join TSPL_DEDUCTION_MASTER On TSPL_DEDUCTION_MASTER.Code=TSPL_DCS_SALE_ENTRY_DETAIL.Deduction "
            qry += " left outer join  TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code= TSPL_DCS_SALE_ENTRY_DETAIL.Tax_Group "
            qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Location "
            qry += " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code"
            qry += " where TSPL_DCS_SALE_ENTRY_DETAIL.Document_Code='" + obj.Document_Code + "' ORDER BY TSPL_DCS_SALE_ENTRY_DETAIL.Line_No  asc"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsDCSSaleEntryDetail)
                Dim objTr As clsDCSSaleEntryDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsDCSSaleEntryDetail
                    objTr.PK_ID = clsCommon.myCdbl(dr("PK_ID"))
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(dr("ItemwiseTaxCode"))
                    objTr.Commission_Rate = clsCommon.myCdbl(dr("Commission_Rate"))
                    objTr.Commission_Party = clsCommon.myCstr(dr("Commission_Party"))
                    objTr.Commission_Amt = clsCommon.myCdbl(dr("Commission_Amt"))
                    objTr.Amt_Less_Commission = clsCommon.myCdbl(dr("Amt_Less_Commission"))
                    objTr.OrgUnit_code = clsCommon.myCstr(dr("OrgUnit_code"))
                    objTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objTr.Row_Type = clsCommon.myCstr(dr("Row_Type"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Status = Convert.ToInt32(clsCommon.myCdbl(dr("Status")))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.Bar_Code = clsCommon.myCstr(dr("Bar_Code"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.DCS_Sale_Zero_Cost = clsCommon.myCdbl(dr("DCS_Sale_Zero_Cost"))
                    objTr.Deduction = clsCommon.myCstr(dr("Deduction"))
                    objTr.Deduction_Name = clsCommon.myCstr(dr("Deduction_Name"))
                    objTr.Tax_Group = clsCommon.myCstr(dr("Tax_Group"))
                    objTr.TaxGroupName = clsCommon.myCstr(dr("TaxGroupName"))
                    objTr.Qty = clsCommon.myCdbl(dr("Qty"))


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
                    objTr.TotalSubsidyDisAmt = clsCommon.myCdbl(dr("TotalSubsidyDisAmt"))
                    objTr.Transporter_Commission_Rate = clsCommon.myCdbl(dr("Transporter_Commission_Rate"))
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(dr("Transporter_Commission_Amt"))
                    objTr.RoundOffAmount = clsCommon.myCdbl(dr("RoundOffAmount"))
                    objTr.REF_TPT_PK_ID = clsCommon.myCdbl(dr("REF_TPT_PK_ID"))
                    objTr.RateDiff_Per = clsCommon.myCdbl(dr("RateDiff_Per"))
                    objTr.RateDiff_Amt = clsCommon.myCdbl(dr("RateDiff_Amt"))
                    objTr.TotalSubsidyAmt = clsCommon.myCdbl(dr("TotalSubsidyAmt"))
                    objTr.Gross_Amount = clsCommon.myCdbl(dr("Gross_Amount"))
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
                    objTr.Is_Add_TPT = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Add_TPT")) = 1, True, False)
                    objTr.arrSrItem = clsSerializeInvenotry.GetData("DCS-SAL-ENT", objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    objTr.arrBatchItem = clsBatchInventory.GetData("DCS-SAL-ENT", objTr.Document_Code, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr.Add(objTr)
                Next
            End If
        End If

        Return obj
    End Function
    Public Shared Sub UpdateFileInfo(ByVal FormId As String, ByVal strDocNo As String)
        Try
            Dim qry As String = "select Document_Code,Document_Date from TSPL_DCS_SALE_ENTRY where Document_Code='" + strDocNo + "' and FILE_INFO is null"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim frmCRV As New frmCrystalReportViewer()
                    Dim PDFPath As String = funPrint(FormId, False, clsCommon.myCDate(dt.Rows(0)("Document_Date")), strDocNo, True)
                    Dim FileNo As Integer = clsAttachDocument.UploadWithHttpRequest(PDFPath, Path.GetFileName(PDFPath), FormId, strDocNo)
                    Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                    Try
                        If FileNo > 0 Then
                            qry = " UPDATE TSPL_DCS_SALE_ENTRY set FILE_INFO=" + clsCommon.myCstr(FileNo) + " where Document_Code='" + strDocNo + "'"
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
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Dim GeneratePDFForMobile As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GeneratePDFForMobile, clsFixedParameterCode.GeneratePDFForMobile, trans))
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If
            Dim obj As clsDCSSaleEntry = clsDCSSaleEntry.GetData(strDocNo, NavigatorType.Current, trans)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmDCSSaleEntry, obj.Bill_To_Location, obj.Document_Date, trans)
            If (obj.Status = 1) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            If (obj.On_Hold) Then
                Throw New Exception("Document No " + obj.Document_Code + " Is On Hold.Can't Post it")
            End If

            Dim qry As String = ""

            Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_DCS_SALE_ENTRY", "Document_Code", obj.Document_Code, trans)
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
            Dim AllowRoundOffAmountOnDCSSale As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowRoundOffAmountOnDCSSale, clsFixedParameterCode.AllowRoundOffAmountOnDCSSale, trans)) = "1", True, False))
            Dim Arr As New Dictionary(Of String, clsMCCMaterialSale)
            Dim strDedTaxGroup As String = ""
            Dim objDCSSale As New clsMCCMaterialSale()
            For Each objDCSEntrySale As clsDCSSaleEntryDetail In obj.Arr
                strDedTaxGroup = (objDCSEntrySale.Deduction + " " + objDCSEntrySale.Tax_Group).ToUpper()
                If Not Arr.ContainsKey(strDedTaxGroup) Then
                    objDCSSale = New clsMCCMaterialSale()
                    objDCSSale.Document_Date = obj.Podate
                    objDCSSale.Total_Comm_Amt = 0
                    objDCSSale.Is_Apply_TPT = obj.Is_Apply_TPT
                    objDCSSale.Is_Add_TPT = objDCSEntrySale.Is_Add_TPT
                    objDCSSale.Recommended_By = obj.Recommended_By
                    objDCSSale.TPT_Vendor = obj.TPT_Vendor
                    objDCSSale.No_Of_Instalment = obj.No_Of_Instalment
                    objDCSSale.Exclude_KKF_And_Mandi = obj.Exclude_KKF_And_Mandi
                    objDCSSale.RoundOffAmount = 0
                    Dim isTaxable As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select IsTaxable from tspl_item_master where item_code ='" + objDCSEntrySale.Item_Code + "'", trans) = 1)
                    objDCSSale.Invoice_Type = IIf(isTaxable, "T", "N")
                    objDCSSale.Document_Date = obj.Document_Date
                    objDCSSale.Customer_Code = obj.Customer_Code
                    objDCSSale.Customer_Name = obj.Customer_Name
                    objDCSSale.Status = IIf(obj.Status = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
                    objDCSSale.On_Hold = IIf(obj.On_Hold = 1, True, False)
                    objDCSSale.Is_Internal = IIf(obj.Is_Internal = 1, True, False)
                    objDCSSale.Ref_No = obj.Ref_No
                    objDCSSale.Description = obj.Description
                    objDCSSale.Remarks = obj.Remarks
                    objDCSSale.Bill_To_Location = obj.Bill_To_Location
                    objDCSSale.Sub_Location_code = obj.Sub_Location_code
                    objDCSSale.Ship_To_Location = obj.Ship_To_Location
                    objDCSSale.Bank_Code = obj.Bank_Code
                    objDCSSale.Payment_Terms = obj.Payment_Terms
                    objDCSSale.Is_CashSale = obj.Is_CashSale
                    objDCSSale.Tax_Group = objDCSEntrySale.Tax_Group
                    objDCSSale.TAX1 = objDCSEntrySale.TAX1
                    objDCSSale.TAX1_Rate = objDCSEntrySale.TAX1_Rate
                    objDCSSale.TAX1_Base_Amt = 0
                    objDCSSale.TAX1_Amt = 0
                    objDCSSale.TAX2 = objDCSEntrySale.TAX2
                    objDCSSale.TAX2_Rate = objDCSEntrySale.TAX2_Rate
                    objDCSSale.TAX2_Base_Amt = 0
                    objDCSSale.TAX2_Amt = 0
                    objDCSSale.TAX3 = objDCSEntrySale.TAX3
                    objDCSSale.TAX3_Base_Amt = 0
                    objDCSSale.TAX3_Rate = objDCSEntrySale.TAX3_Rate
                    objDCSSale.TAX3_Amt = 0
                    objDCSSale.TAX4 = objDCSEntrySale.TAX4
                    objDCSSale.TAX4_Rate = objDCSEntrySale.TAX4_Rate
                    objDCSSale.TAX4_Base_Amt = 0
                    objDCSSale.TAX4_Amt = 0
                    objDCSSale.TAX5 = objDCSEntrySale.TAX5
                    objDCSSale.TAX5_Rate = objDCSEntrySale.TAX5_Rate
                    objDCSSale.TAX5_Base_Amt = 0
                    objDCSSale.TAX5_Amt = 0
                    objDCSSale.TAX6 = objDCSEntrySale.TAX6
                    objDCSSale.TAX6_Rate = objDCSEntrySale.TAX6_Rate
                    objDCSSale.TAX6_Base_Amt = 0
                    objDCSSale.TAX6_Amt = 0
                    objDCSSale.TAX7 = objDCSEntrySale.TAX7
                    objDCSSale.TAX7_Rate = objDCSEntrySale.TAX7_Rate
                    objDCSSale.TAX7_Base_Amt = 0
                    objDCSSale.TAX7_Amt = 0
                    objDCSSale.TAX8 = objDCSEntrySale.TAX8
                    objDCSSale.TAX8_Rate = objDCSEntrySale.TAX8_Rate
                    objDCSSale.TAX8_Base_Amt = 0
                    objDCSSale.TAX8_Amt = 0
                    objDCSSale.TAX9 = objDCSEntrySale.TAX9
                    objDCSSale.TAX9_Rate = objDCSEntrySale.TAX9_Rate
                    objDCSSale.TAX9_Base_Amt = 0
                    objDCSSale.TAX9_Amt = 0
                    objDCSSale.TAX10 = objDCSEntrySale.TAX10
                    objDCSSale.TAX10_Rate = objDCSEntrySale.TAX10_Rate
                    objDCSSale.TAX10_Base_Amt = 0
                    objDCSSale.TAX10_Amt = 0
                    objDCSSale.Discount_Base = 0
                    objDCSSale.Discount_Amt = 0
                    objDCSSale.Amount_Less_Discount = 0
                    objDCSSale.Total_Amt = 0
                    objDCSSale.Total_Tax_Amt = 0
                    objDCSSale.Comments = obj.Comments
                    objDCSSale.Comp_Code = obj.Comp_Code
                    objDCSSale.Terms_Code = obj.Terms_Code
                    objDCSSale.Due_Date = obj.Due_Date
                    objDCSSale.BillToLocationName = obj.BillToLocationName
                    objDCSSale.ShipToLocationName = obj.ShipToLocationName
                    objDCSSale.TaxGroupName = objDCSEntrySale.TaxGroupName
                    objDCSSale.TermsName = obj.TermsName
                    objDCSSale.PROJECT_ID = obj.PROJECT_ID
                    objDCSSale.Route_No = obj.Route_No
                    objDCSSale.Route_Desc = obj.Route_Desc
                    objDCSSale.Price_Code = obj.Price_Code
                    objDCSSale.DCS_Price_Code = obj.DCS_Price_Code
                    objDCSSale.Supply_Date = obj.Supply_Date
                    objDCSSale.Shift_Type = obj.Shift_Type
                    objDCSSale.HeadDisc_Per = objDCSEntrySale.HeadDiscPer
                    objDCSSale.HeadDisc_Amt = 0
                    objDCSSale.HeadDisc_PerAmt = 0
                    objDCSSale.RateDiff_Per = 0
                    objDCSSale.RateDiff_Amt = 0
                    objDCSSale.Gross_Amount = 0
                    objDCSSale.TotalSubsidyAmt = 0
                    objDCSSale.TotalSubsidyDisAmt = 0
                    objDCSSale.Transporter_Commission_TotalAmt = 0
                    objDCSSale.TotCashDiscAmt = 0
                    objDCSSale.Cust_PO_No = obj.Cust_PO_No
                    If obj.Posting_Date IsNot Nothing Then
                        objDCSSale.Posting_Date = obj.Posting_Date
                    End If
                    objDCSSale.Inv_Date = obj.Document_Date
                    objDCSSale.Salesman_Code = obj.Salesman_Code
                    objDCSSale.Salesman_Name = obj.Salesman_Name
                    objDCSSale.Challan_No = obj.Challan_No
                    objDCSSale.Carrier = obj.Carrier
                    objDCSSale.Vehicle_Code = obj.Vehicle_Code
                    objDCSSale.VehicleNo = obj.VehicleNo
                    objDCSSale.GRNo = obj.GRNo
                    objDCSSale.GENo = obj.GENo
                    If obj.GEDate IsNot Nothing Then
                        objDCSSale.GEDate = obj.GEDate
                    End If
                    objDCSSale.Deduction_Type = objDCSEntrySale.Deduction_Type
                    objDCSSale.Deduction = objDCSEntrySale.Deduction
                    objDCSSale.Dept = obj.Dept
                    objDCSSale.Dept_Desc = obj.Dept_Desc
                    objDCSSale.Item_Type = obj.Item_Type
                    objDCSSale.Add_Charge_Code1 = obj.Add_Charge_Code1
                    objDCSSale.Add_Charge_Name1 = obj.Add_Charge_Name1
                    objDCSSale.Add_Charge_Amt1 = obj.Add_Charge_Amt1
                    objDCSSale.Add_Charge_Code2 = obj.Add_Charge_Code2
                    objDCSSale.Add_Charge_Name2 = obj.Add_Charge_Name2
                    objDCSSale.Add_Charge_Amt2 = obj.Add_Charge_Amt2
                    objDCSSale.Add_Charge_Code3 = obj.Add_Charge_Code3
                    objDCSSale.Add_Charge_Name3 = obj.Add_Charge_Name3
                    objDCSSale.Add_Charge_Amt3 = obj.Add_Charge_Amt3
                    objDCSSale.Add_Charge_Code4 = obj.Add_Charge_Code4
                    objDCSSale.Add_Charge_Name4 = obj.Add_Charge_Name4
                    objDCSSale.Add_Charge_Amt4 = obj.Add_Charge_Amt4
                    objDCSSale.Add_Charge_Code5 = obj.Add_Charge_Code5
                    objDCSSale.Add_Charge_Name5 = obj.Add_Charge_Name5
                    objDCSSale.Add_Charge_Amt5 = obj.Add_Charge_Amt5
                    objDCSSale.Add_Charge_Code6 = obj.Add_Charge_Code6
                    objDCSSale.Add_Charge_Name6 = obj.Add_Charge_Name6
                    objDCSSale.Add_Charge_Amt6 = obj.Add_Charge_Amt6
                    objDCSSale.Add_Charge_Code7 = obj.Add_Charge_Code7
                    objDCSSale.Add_Charge_Name7 = obj.Add_Charge_Name7
                    objDCSSale.Add_Charge_Amt7 = obj.Add_Charge_Amt7
                    objDCSSale.Add_Charge_Code8 = obj.Add_Charge_Code8
                    objDCSSale.Add_Charge_Name8 = obj.Add_Charge_Name8
                    objDCSSale.Add_Charge_Amt8 = obj.Add_Charge_Amt8
                    objDCSSale.Add_Charge_Code9 = obj.Add_Charge_Code9
                    objDCSSale.Add_Charge_Name9 = obj.Add_Charge_Name9
                    objDCSSale.Add_Charge_Amt9 = obj.Add_Charge_Amt9
                    objDCSSale.Add_Charge_Code10 = obj.Add_Charge_Code10
                    objDCSSale.Add_Charge_Name10 = obj.Add_Charge_Name10
                    objDCSSale.Add_Charge_Amt10 = obj.Add_Charge_Amt10
                    objDCSSale.Total_Add_Charge = obj.Total_Add_Charge
                    If clsCommon.myLen(obj.Challan_Date) <= 0 Then
                        objDCSSale.Challan_Date = ""
                    Else
                        objDCSSale.Challan_Date = clsCommon.GetPrintDate(obj.Challan_Date, "dd/MMM/yyyy")
                    End If
                    objDCSSale.SO_Validity = obj.SO_Validity
                    objDCSSale.Commission_Apply = obj.Commission_Apply
                    objDCSSale.Vehicle_Capacity = obj.Vehicle_Capacity
                    objDCSSale.Dispatch_Terms = obj.Dispatch_Terms
                    objDCSSale.Payment_Terms = obj.Payment_Terms
                    objDCSSale.Dispatch_Period = obj.Dispatch_Period
                    objDCSSale.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                    objDCSSale.Is_Create_Auto_Receipt = obj.Is_Create_Auto_Receipt
                    objDCSSale.Is_Create_Auto_Invoice = True
                    objDCSSale.Is_Taxable = isTaxable
                    objDCSSale.Arr = New List(Of clsMCCMaterialSaleDetail)
                    Arr.Add(strDedTaxGroup, objDCSSale)
                End If

                Dim objDCSSaleDetail As New clsMCCMaterialSaleDetail()
                objDCSSaleDetail.REF_PK_ID = objDCSEntrySale.PK_ID
                objDCSSaleDetail.REF_TPT_PK_ID = objDCSEntrySale.REF_TPT_PK_ID
                objDCSSaleDetail.PrincipleCode = objDCSEntrySale.PrincipleCode
                objDCSSaleDetail.PrincipleDesc = objDCSEntrySale.PrincipleDesc
                objDCSSaleDetail.vendor_code = objDCSEntrySale.vendor_code
                objDCSSaleDetail.vendor_desc = objDCSEntrySale.vendor_desc
                objDCSSaleDetail.Document_Code = objDCSEntrySale.Document_Code
                objDCSSaleDetail.Row_Type = objDCSEntrySale.Row_Type

                objDCSSaleDetail.Status = objDCSEntrySale.Status
                objDCSSaleDetail.Item_Code = objDCSEntrySale.Item_Code
                objDCSSaleDetail.Item_Desc = objDCSEntrySale.Item_Desc
                objDCSSaleDetail.Qty = objDCSEntrySale.Qty
                objDCSSaleDetail.Free_Qty = objDCSEntrySale.Free_Qty
                objDCSSaleDetail.Balance_Qty = objDCSEntrySale.Balance_Qty
                objDCSSaleDetail.Unit_code = objDCSEntrySale.Unit_code
                objDCSSaleDetail.Location = objDCSEntrySale.Location
                objDCSSaleDetail.LocationName = objDCSEntrySale.LocationName
                objDCSSaleDetail.Item_Cost = objDCSEntrySale.Item_Cost
                objDCSSaleDetail.TAX1 = objDCSEntrySale.TAX1
                objDCSSaleDetail.TAX1_Base_Amt = objDCSEntrySale.TAX1_Base_Amt
                objDCSSaleDetail.TAX1_Rate = objDCSEntrySale.TAX1_Rate
                objDCSSaleDetail.TAX1_Amt = objDCSEntrySale.TAX1_Amt
                objDCSSaleDetail.TAX2 = objDCSEntrySale.TAX2
                objDCSSaleDetail.TAX2_Base_Amt = objDCSEntrySale.TAX2_Base_Amt
                objDCSSaleDetail.TAX2_Rate = objDCSEntrySale.TAX2_Rate
                objDCSSaleDetail.TAX2_Amt = objDCSEntrySale.TAX2_Amt
                objDCSSaleDetail.TAX3 = objDCSEntrySale.TAX3
                objDCSSaleDetail.TAX3_Base_Amt = objDCSEntrySale.TAX3_Base_Amt
                objDCSSaleDetail.TAX3_Rate = objDCSEntrySale.TAX3_Rate
                objDCSSaleDetail.TAX3_Amt = objDCSEntrySale.TAX3_Amt
                objDCSSaleDetail.TAX4 = objDCSEntrySale.TAX4
                objDCSSaleDetail.TAX4_Base_Amt = objDCSEntrySale.TAX4_Base_Amt
                objDCSSaleDetail.TAX4_Rate = objDCSEntrySale.TAX4_Rate
                objDCSSaleDetail.TAX4_Amt = objDCSEntrySale.TAX4_Amt
                objDCSSaleDetail.TAX5 = objDCSEntrySale.TAX5
                objDCSSaleDetail.TAX5_Base_Amt = objDCSEntrySale.TAX5_Base_Amt
                objDCSSaleDetail.TAX5_Rate = objDCSEntrySale.TAX5_Rate
                objDCSSaleDetail.TAX5_Amt = objDCSEntrySale.TAX5_Amt
                objDCSSaleDetail.TAX6 = objDCSEntrySale.TAX6
                objDCSSaleDetail.TAX6_Base_Amt = objDCSEntrySale.TAX6_Base_Amt
                objDCSSaleDetail.TAX6_Rate = objDCSEntrySale.TAX6_Rate
                objDCSSaleDetail.TAX6_Amt = objDCSEntrySale.TAX6_Amt
                objDCSSaleDetail.TAX7 = objDCSEntrySale.TAX7
                objDCSSaleDetail.TAX7_Base_Amt = objDCSEntrySale.TAX7_Base_Amt
                objDCSSaleDetail.TAX7_Rate = objDCSEntrySale.TAX7_Rate
                objDCSSaleDetail.TAX7_Amt = objDCSEntrySale.TAX7_Amt
                objDCSSaleDetail.TAX8 = objDCSEntrySale.TAX8
                objDCSSaleDetail.TAX8_Base_Amt = objDCSEntrySale.TAX8_Base_Amt
                objDCSSaleDetail.TAX8_Rate = objDCSEntrySale.TAX8_Rate
                objDCSSaleDetail.TAX8_Amt = objDCSEntrySale.TAX8_Amt
                objDCSSaleDetail.TAX9 = objDCSEntrySale.TAX9
                objDCSSaleDetail.TAX9_Base_Amt = objDCSEntrySale.TAX9_Base_Amt
                objDCSSaleDetail.TAX9_Rate = objDCSEntrySale.TAX9_Rate
                objDCSSaleDetail.TAX9_Amt = objDCSEntrySale.TAX9_Amt
                objDCSSaleDetail.TAX10 = objDCSEntrySale.TAX10
                objDCSSaleDetail.TAX10_Base_Amt = objDCSEntrySale.TAX10_Base_Amt
                objDCSSaleDetail.TAX10_Rate = objDCSEntrySale.TAX10_Rate
                objDCSSaleDetail.TAX10_Amt = objDCSEntrySale.TAX10_Amt
                objDCSSaleDetail.Amount = objDCSEntrySale.Amount
                objDCSSaleDetail.Disc_Per = objDCSEntrySale.Disc_Per
                objDCSSaleDetail.Disc_Amt = objDCSEntrySale.Disc_Amt
                objDCSSaleDetail.Amt_Less_Discount = objDCSEntrySale.Amt_Less_Discount
                objDCSSaleDetail.Total_Tax_Amt = objDCSEntrySale.Total_Tax_Amt
                objDCSSaleDetail.Item_Net_Amt = objDCSEntrySale.Item_Net_Amt
                objDCSSaleDetail.Is_Mannual_Amt = objDCSEntrySale.Is_Mannual_Amt
                objDCSSaleDetail.MRP = objDCSEntrySale.MRP
                objDCSSaleDetail.Assessable = objDCSEntrySale.Assessable
                objDCSSaleDetail.AssessableAmt = objDCSEntrySale.AssessableAmt
                objDCSSaleDetail.Batch_No = objDCSEntrySale.Batch_No
                If objDCSEntrySale.MFG_Date IsNot Nothing Then
                    objDCSSaleDetail.MFG_Date = objDCSEntrySale.MFG_Date
                End If
                If objDCSEntrySale.Expiry_Date IsNot Nothing Then
                    objDCSSaleDetail.Expiry_Date = objDCSEntrySale.Expiry_Date
                End If
                objDCSSaleDetail.Specification = objDCSEntrySale.Specification
                objDCSSaleDetail.Remarks = objDCSEntrySale.Remarks
                objDCSSaleDetail.Scheme_Applicable = IIf(objDCSEntrySale.Scheme_Applicable = "Yes", "Y", "N")
                objDCSSaleDetail.Scheme_Code = objDCSEntrySale.Scheme_Code
                objDCSSaleDetail.Scheme_Item = IIf(objDCSEntrySale.Scheme_Item = "Yes", "Y", "N")
                objDCSSaleDetail.Item_Tax = objDCSEntrySale.Item_Tax
                objDCSSaleDetail.Total_MRP_Amt = objDCSEntrySale.Total_MRP_Amt
                objDCSSaleDetail.Total_Basic_Amt = objDCSEntrySale.Total_Basic_Amt
                objDCSSaleDetail.Total_Disc_Amt = objDCSEntrySale.Total_Disc_Amt
                objDCSSaleDetail.Cust_Discount = objDCSEntrySale.Cust_Discount
                objDCSSaleDetail.Total_Cust_Discount = objDCSEntrySale.Total_Cust_Discount
                objDCSSaleDetail.ActualRate = objDCSEntrySale.ActualRate
                objDCSSaleDetail.Cust_DiscountQty = objDCSEntrySale.Cust_DiscountQty
                objDCSSaleDetail.Price_code = objDCSEntrySale.Price_code
                objDCSSaleDetail.Price_Date = objDCSEntrySale.Price_Date
                objDCSSaleDetail.Abatement_Per = objDCSEntrySale.Abatement_Per
                objDCSSaleDetail.Abatement_Amt = objDCSEntrySale.Abatement_Amt
                objDCSSaleDetail.FOC_Item = objDCSEntrySale.FOC_Item
                objDCSSaleDetail.Markup_On = objDCSEntrySale.Markup_On
                objDCSSaleDetail.Markup_Percent = objDCSEntrySale.Markup_Percent
                objDCSSaleDetail.Landing_Cost = objDCSEntrySale.Landing_Cost
                objDCSSaleDetail.HeadDiscAmt = objDCSEntrySale.HeadDiscAmt
                objDCSSaleDetail.HeadDiscPer = objDCSEntrySale.HeadDiscPer
                objDCSSaleDetail.HeadDiscPerAmt = objDCSEntrySale.HeadDiscPerAmt
                objDCSSaleDetail.CustDiscPer = objDCSEntrySale.CustDiscPer
                objDCSSaleDetail.CasdDiscScheme_Code = objDCSEntrySale.CasdDiscScheme_Code
                objDCSSaleDetail.Item_Weight = objDCSEntrySale.Item_Weight
                objDCSSaleDetail.TotalItem_Weight = objDCSEntrySale.TotalItem_Weight
                objDCSSaleDetail.Conv_Factor = objDCSEntrySale.Conv_Factor
                objDCSSaleDetail.Purchase_Cost = objDCSEntrySale.Purchase_Cost
                objDCSSaleDetail.OrgRate = objDCSEntrySale.OrgRate
                objDCSSaleDetail.Price_Amount1 = objDCSEntrySale.Price_Amount1
                objDCSSaleDetail.Price_Amount2 = objDCSEntrySale.Price_Amount2
                objDCSSaleDetail.Price_Amount3 = objDCSEntrySale.Price_Amount3
                objDCSSaleDetail.Price_Amount4 = objDCSEntrySale.Price_Amount4
                objDCSSaleDetail.Price_Amount5 = objDCSEntrySale.Price_Amount5
                objDCSSaleDetail.Price_Amount6 = objDCSEntrySale.Price_Amount6
                objDCSSaleDetail.Price_Amount7 = objDCSEntrySale.Price_Amount7
                objDCSSaleDetail.Price_Amount8 = objDCSEntrySale.Price_Amount8
                objDCSSaleDetail.Price_Amount9 = objDCSEntrySale.Price_Amount9
                objDCSSaleDetail.Price_Amount10 = objDCSEntrySale.Price_Amount10
                objDCSSaleDetail.TAX1_Base_Amt = objDCSEntrySale.TAX1_Base_Amt
                objDCSSaleDetail.TAX2_Base_Amt = objDCSEntrySale.TAX2_Base_Amt
                objDCSSaleDetail.TAX3_Base_Amt = objDCSEntrySale.TAX3_Base_Amt
                objDCSSaleDetail.TAX4_Base_Amt = objDCSEntrySale.TAX4_Base_Amt
                objDCSSaleDetail.TAX5_Base_Amt = objDCSEntrySale.TAX5_Base_Amt
                objDCSSaleDetail.TAX6_Base_Amt = objDCSEntrySale.TAX6_Base_Amt
                objDCSSaleDetail.TAX7_Base_Amt = objDCSEntrySale.TAX7_Base_Amt
                objDCSSaleDetail.TAX8_Base_Amt = objDCSEntrySale.TAX8_Base_Amt
                objDCSSaleDetail.TAX9_Base_Amt = objDCSEntrySale.TAX9_Base_Amt
                objDCSSaleDetail.TAX10_Base_Amt = objDCSEntrySale.TAX10_Base_Amt
                objDCSSaleDetail.Commission_Rate = objDCSEntrySale.Commission_Rate
                objDCSSaleDetail.Commission_Party = objDCSEntrySale.Commission_Party
                objDCSSaleDetail.Commission_Amt = objDCSEntrySale.Commission_Amt
                objDCSSaleDetail.Amt_Less_Commission = objDCSEntrySale.Amt_Less_Commission
                objDCSSaleDetail.Transporter_Commission_Rate = objDCSEntrySale.Transporter_Commission_Rate
                objDCSSaleDetail.Transporter_Commission_Amt = objDCSEntrySale.Transporter_Commission_Amt
                objDCSSaleDetail.arrBatchItem = objDCSEntrySale.arrBatchItem
                objDCSSaleDetail.arrSrItem = objDCSEntrySale.arrSrItem

                Arr(strDedTaxGroup).TAX1_Amt += objDCSEntrySale.TAX1_Amt
                Arr(strDedTaxGroup).TAX1_Base_Amt += objDCSEntrySale.TAX1_Base_Amt
                Arr(strDedTaxGroup).TAX2_Amt += objDCSEntrySale.TAX2_Amt
                Arr(strDedTaxGroup).TAX2_Base_Amt += objDCSEntrySale.TAX2_Base_Amt
                Arr(strDedTaxGroup).TAX3_Amt += objDCSEntrySale.TAX3_Amt
                Arr(strDedTaxGroup).TAX3_Base_Amt += objDCSEntrySale.TAX3_Base_Amt
                Arr(strDedTaxGroup).TAX4_Amt += objDCSEntrySale.TAX4_Amt
                Arr(strDedTaxGroup).TAX4_Base_Amt += objDCSEntrySale.TAX4_Base_Amt
                Arr(strDedTaxGroup).TAX5_Amt += objDCSEntrySale.TAX5_Amt
                Arr(strDedTaxGroup).TAX5_Base_Amt += objDCSEntrySale.TAX5_Base_Amt
                Arr(strDedTaxGroup).TAX6_Amt += objDCSEntrySale.TAX6_Amt
                Arr(strDedTaxGroup).TAX6_Base_Amt += objDCSEntrySale.TAX6_Base_Amt
                Arr(strDedTaxGroup).TAX7_Amt += objDCSEntrySale.TAX7_Amt
                Arr(strDedTaxGroup).TAX7_Base_Amt += objDCSEntrySale.TAX7_Base_Amt
                Arr(strDedTaxGroup).TAX8_Amt += objDCSEntrySale.TAX8_Amt
                Arr(strDedTaxGroup).TAX8_Base_Amt += objDCSEntrySale.TAX8_Base_Amt
                Arr(strDedTaxGroup).TAX9_Amt += objDCSEntrySale.TAX9_Amt
                Arr(strDedTaxGroup).TAX9_Base_Amt += objDCSEntrySale.TAX9_Base_Amt
                Arr(strDedTaxGroup).TAX10_Amt += objDCSEntrySale.TAX10_Amt
                Arr(strDedTaxGroup).TAX10_Base_Amt += objDCSEntrySale.TAX10_Base_Amt
                Arr(strDedTaxGroup).Total_Tax_Amt += objDCSEntrySale.Total_Tax_Amt
                Arr(strDedTaxGroup).Amount_Less_Discount += objDCSEntrySale.Amt_Less_Discount
                Arr(strDedTaxGroup).Discount_Base += objDCSEntrySale.Total_Basic_Amt
                Arr(strDedTaxGroup).Discount_Amt += objDCSEntrySale.Disc_Amt
                Arr(strDedTaxGroup).HeadDisc_Amt += objDCSEntrySale.HeadDiscAmt
                Arr(strDedTaxGroup).HeadDisc_PerAmt += objDCSEntrySale.HeadDiscPerAmt
                Arr(strDedTaxGroup).TotCashDiscAmt += objDCSEntrySale.Total_Cust_Discount
                Arr(strDedTaxGroup).Total_Comm_Amt += objDCSEntrySale.Commission_Amt

                If AllowRoundOffAmountOnDCSSale Then
                    Dim lstDecml As New List(Of Decimal)
                    lstDecml = ClsScrapSaleHead.Calculate_RoundOffAmt(clsCommon.myCdbl(objDCSEntrySale.Item_Net_Amt), Nothing)
                    If lstDecml IsNot Nothing AndAlso lstDecml.Count > 0 Then
                        Arr(strDedTaxGroup).Total_Amt += clsCommon.myCdbl(lstDecml(0))
                        Arr(strDedTaxGroup).RoundOffAmount += clsCommon.myCdbl(lstDecml(1))
                    End If
                Else
                    Arr(strDedTaxGroup).Total_Amt += objDCSEntrySale.Item_Net_Amt
                    Arr(strDedTaxGroup).RoundOffAmount += 0
                End If


                Arr(strDedTaxGroup).RateDiff_Per += objDCSEntrySale.RateDiff_Per
                Arr(strDedTaxGroup).RateDiff_Amt += objDCSEntrySale.RateDiff_Amt
                Arr(strDedTaxGroup).TotalSubsidyAmt += objDCSEntrySale.TotalSubsidyAmt
                Arr(strDedTaxGroup).TotalSubsidyDisAmt += objDCSEntrySale.TotalSubsidyDisAmt
                Arr(strDedTaxGroup).Transporter_Commission_TotalAmt += objDCSEntrySale.Transporter_Commission_Amt
                Arr(strDedTaxGroup).Gross_Amount += objDCSEntrySale.Gross_Amount

                objDCSSaleDetail.Line_No = Arr(strDedTaxGroup).Arr.Count + 1
                Arr(strDedTaxGroup).Arr.Add(objDCSSaleDetail)

            Next

            qry = " update TSPL_BATCH_ITEM set In_Out_Type='X' where Document_Type='DCS-SAL-ENT' and Document_Code='" + strDocNo + "' "
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            For ii As Integer = 0 To Arr.Keys.Count - 1
                Dim objShipment As clsMCCMaterialSale = Arr.Item(Arr.Keys(ii))
                objShipment.SaveData(objShipment, True, trans)
                clsMCCMaterialSale.PostData(clsUserMgtCode.frmDCSSaleEntry, objShipment.Document_Code, trans)
            Next

            qry = "Update TSPL_DCS_SALE_ENTRY set "
            If clsCommon.myLen(obj.Against_Sales_Order) = 0 Then
                obj.Against_Sales_Order = "NULL"
            End If
            qry += "Against_Sales_Order=" & obj.Against_Sales_Order & ", Status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "' "
            qry += " where Document_Code='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            trans.Commit()

            If GeneratePDFForMobile = 1 Then
                For ii As Integer = 0 To Arr.Keys.Count - 1
                    Dim objShipment As clsMCCMaterialSale = Arr.Item(Arr.Keys(ii))
                    clsMCCMaterialSale.UpdateFileInfo(clsUserMgtCode.frmMCCMaterial, objShipment.Document_Code)
                Next
            ElseIf GeneratePDFForMobile = 2 Then
                UpdateFileInfo(FormId, strDocNo)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DCS_SALE_ENTRY", "Document_Code", trans)
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
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
            Throw New Exception("Document No not found to Delete")
        End If
        Dim obj As clsDCSSaleEntry = clsDCSSaleEntry.GetData(strCode, NavigatorType.Current, trans)

        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
            Try
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmDCSSaleEntry, obj.Bill_To_Location, obj.Document_Date, trans)

                ''
                If (obj.Status = 1) Then
                    Throw New Exception("Already Posted on :" + obj.Posting_Date)
                End If

                clsSerializeInvenotry.DeleteData("DCS-SAL-ENT", strCode, trans)
                clsBatchInventory.DeleteData("DCS-SAL-ENT", obj.Document_Code, trans)
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strCode, "TSPL_DCS_SALE_ENTRY", "Document_Code", "TSPL_DCS_SALE_ENTRY_DETAIL", "Document_Code", trans)
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, strCode, "TSPL_DCS_SALE_ENTRY", "Document_Code", "TSPL_DCS_SALE_ENTRY_DETAIL", "Document_Code", trans)

                Dim qry As String = "select distinct isnull(TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,'') DOCUMENT_CODE from TSPL_SD_SHIPMENT_DETAIL where REF_PK_ID in (select  PK_ID from TSPL_DCS_SALE_ENTRY_DETAIL where Document_Code='" + strCode + "') "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        clsMCCMaterialSale.DeleteData(clsCommon.myCstr(dr("DOCUMENT_CODE")), trans)
                    Next
                End If
                qry = "delete from TSPL_DCS_SALE_ENTRY_DETAIL where Document_Code='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_DCS_SALE_ENTRY where Document_Code='" + strCode + "'"
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
            Dim qry As String = "select TSPL_DCS_SALE_ENTRY.Document_Code,TSPL_DCS_SALE_ENTRY.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name from TSPL_DCS_SALE_ENTRY left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DCS_SALE_ENTRY.Customer_Code where Document_Code in (" + clsCommon.GetMulcallString(Arr) + ") and Customer_Code not in ('" + strCustomerCode + "')"
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
            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select Status from TSPL_DCS_SALE_ENTRY where Document_Code='" + strCode + "'"
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            Qry = "select distinct isnull(TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,'') DOCUMENT_CODE from TSPL_SD_SHIPMENT_DETAIL where REF_PK_ID in (select  PK_ID from TSPL_DCS_SALE_ENTRY_DETAIL where Document_Code='" + strCode + "') "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    clsMCCMaterialSale.ReverseAndUnpost(clsCommon.myCstr(dr("DOCUMENT_CODE")), True, trans)
                Next
            End If
            clsBatchInventory.ReverseAndUnpost("DCS-SAL-ENT", strCode, trans)

            Qry = "Update TSPL_DCS_SALE_ENTRY set Status = 0 where Document_Code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_DCS_SALE_ENTRY", "Document_Code", trans)
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
            QryShowStatus = " ,(case when TSPL_DCS_SALE_ENTRY.status =1 then 'AUTHORIZED' else 'NOT AUTHORIZED' end) as Status "
        Else
            QryShowStatus = ""
        End If
        Dim SerialNo As String = ""
        Dim SerialNoColumn As String = ""
        Dim ShowSerialNoForSales As Double = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowSerialNoForSales' And Type ='ShowSerialNoForSales'")
        If clsCommon.CompairString(clsCommon.myCstr(ShowSerialNoForSales), "1") = CompairStringResult.Equal Then
            SerialNoColumn = " ,1 As SerialNoText , aa.Serial_No As [SerialNo]  "
            'SerialNo = "  left outer join TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL  on TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code  =TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Main_Item_Code And TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.IS_Principle=1 ANd TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Location_Code =TSPL_DCS_SALE_ENTRY_DETAIL.Location "
            SerialNo = " left outer join (select distinct Doc_No,Serial_No,Main_Item_Code,Location_Code from TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL WHERE Is_principle='1' AND ISNULL(Serial_No,'')<>'' and Doc_No in (select Doc_No from TSPL_MF_PRINCIPLE_RECEIPT_HEAD where Status='1'))aa  on TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code  =AA.Main_Item_Code  ANd aa.Location_Code =TSPL_DCS_SALE_ENTRY_DETAIL.Location  "
        Else
            SerialNoColumn = " ,0 As SerialNoText "
            SerialNo = ""
        End If

        Dim TaxQry As String = "select DOCUMENT_CODE,Tax,max(Type)Type,max(Is_TCS)Is_TCS,TaxRate,SUM(TAX_Amt) as TAX_Amt from ( select * from ( "
        For ii As Integer = 1 To 10
            If ii > 1 Then
                TaxQry += " " + Environment.NewLine + " union all " + Environment.NewLine + ""
            End If
            TaxQry += " SELECT " + clsCommon.myCstr(ii) + " as Sno,TSPL_DCS_SALE_ENTRY_DETAIL.DOCUMENT_CODE, TAX" + clsCommon.myCstr(ii) + " as Tax,TAX" + clsCommon.myCstr(ii) + "_Rate as TaxRate, TAX" + clsCommon.myCstr(ii) + "_Amt as TAX_Amt,tspl_tax_master.Type,tspl_tax_master.Is_TCS,Tax_Code_Desc as taxname FROM TSPL_DCS_SALE_ENTRY_DETAIL  
            Left outer join tspl_tax_master on tspl_tax_master.Tax_Code=TSPL_DCS_SALE_ENTRY_DETAIL.tax" + clsCommon.myCstr(ii) + " where DOCUMENT_CODE='" + strCode + "' "
        Next
        TaxQry += " )xx where len(ISNULL(Tax,''))>0  ) xxx group by DOCUMENT_CODE,sno,Tax,TaxRate "
        Dim TotalTaxQry As String = ""
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(TaxQry)
        If dt.Rows.Count > 0 AndAlso dt IsNot Nothing Then
            TotalTaxQry = " select max(DOCUMENT_CODE)DOCUMENT_CODE"
            For ii As Integer = 0 To dt.Rows.Count - 1
                TotalTaxQry += "  ,sum(case when Type='" + dt.Rows(ii)("Type") + "' and Is_TCS='" + dt.Rows(ii)("Is_TCS") + "' then TAX_Amt else 0 end) as DTax" + clsCommon.myCstr(ii + 1) + "_Amt,max(case when Type='" + dt.Rows(ii)("Type") + "' and Is_TCS='" + dt.Rows(ii)("Is_TCS") + "' then TaxRate else 0 end) as DTax" + clsCommon.myCstr(ii + 1) + "_Rate, max(case when Type='" + dt.Rows(ii)("Type") + "' and Is_TCS='" + dt.Rows(ii)("Is_TCS") + "' then Type else '' end) as tax" + clsCommon.myCstr(ii + 1) + "Type "
            Next
            For ii As Integer = dt.Rows.Count + 1 To 10
                TotalTaxQry += ",'' as tax" + clsCommon.myCstr(ii) + "Type,0 as DTax" + clsCommon.myCstr(ii) + "_Rate,0 as DTax" + clsCommon.myCstr(ii) + "_Amt "
            Next

        End If

        Dim Qry As String = "select *, itemcost/ConversionFactor As RateLtr from (  select '" + objCommonVar.CurrentUserCode + "' as UserName, tabbatch.Batch_No, (CASE when TSPL_DCS_SALE_ENTRY_DETAIL.Scheme_Item='Y' then 0 else ((case when TSPL_DCS_SALE_ENTRY_DETAIL.Sampling=1 then 0 else  TSPL_DCS_SALE_ENTRY_DETAIL.Amt_Less_Discount end)) end)  as valueInRs,
ITEMDETAIL1.Conversion_Factor As CF,TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor,"
        If isCancel Then
            Qry += " 'Cancelled' As Report_Status,"
        Else
            Qry += " '' As Report_Status,"
        End If
        Qry += " TSPL_DCS_SALE_ENTRY.Bill_To_Location, TSPL_LOCATION_MASTER.Location_Desc, TSPL_DCS_SALE_ENTRY.Is_Taxable,  StateMasterForLocation.GST_STATE_CODE AS From_GstStateCode," &
          " TSPL_LOCATION_MASTER.GSTNO as From_Loc_GstinNo, TSPL_STATE_MASTER.GST_STATE_Code AS Cust_GstStateCode,  TSPL_CUSTOMER_MASTER.GSTNO as Cust_GstInNo,    case when coalesce(p_cust.GST_STATE_CODE,'')='' then TSPL_state_Master.GST_STATE_CODE       when coalesce(p_cust.GST_STATE_CODE,'')<>'' then p_cust .GST_STATE_CODE    end as P_GST_STATE_CODE," &
          " case when coalesce(p_cust.P_GSTIN_NO,'')='' then TSPL_CUSTOMER_MASTER .GSTNO  when coalesce(p_cust.P_GSTIN_NO,'')<>'' then p_cust .P_GSTIN_NO end as P_GSTIN_NO,    TSPL_ITEM_MASTER.HSN_Code,  " &
         " TSPL_LOCATION_MASTER.State as loc_state_code,tspl_location_master.HOAdd1 ,tspl_location_master.HOAdd2,TSPL_DCS_SALE_ENTRY.HeadDisc_PerAmt,'" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") + "' as RunDate,"
        Qry += " TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end"
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end"
        Qry += "+ Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ',Phone'+TSPL_COMPANY_MASTER.Phone1 end "
        Qry += "+  Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ',  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End "
        Qry += "  + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end "
        Qry += " as Comp_Address,TSPL_DCS_SALE_ENTRY.RoundOffAmount , "
        Qry += " tspl_vlc_master_head.vlc_code_vlc_uploader As VLC_Code,tspl_vlc_master_head.vlc_name,tspl_vlc_master_head.VSP_Code, "
        ''
        Qry += "   case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add1  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add1 end as P_Add1, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add2  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add2 end as P_Add2, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add3  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add3 end as P_Add3, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PIN_Code   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Pin_No  end as P_PinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .CST    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_CST_No   end as P_CstNo,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Tin_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .p_Tin_No   end as P_TinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Email    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Email  end as P_Email,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Fax     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Fax   end as P_Fax, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PAN      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Cust_PAN    end as P_Cust_PAN,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Lst_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_LST_No    end as P_LstNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Cust_Code      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_code   end as P_CustCode, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Customer_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_name    end as P_Cust_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CITY_MASTER   .City_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_City_Name    end as P_City_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_state_Master.state_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_state_Name    end as P_State_Name,"
        Qry += " case when coalesce(p_cust.P_cust_code,'')='' then     case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Phn    end as P_Cust_Phn,TSPL_CUSTOMER_MASTER.Cust_Code  ,TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1,TSPL_CUSTOMER_MASTER.Add2 as Cust_add2,TSPL_CUSTOMER_MASTER.Add3 as cust_add3,  case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End  as Cust_Phn,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_TinNo ,TSPL_CUSTOMER_MASTER.CST as Cust_CSTNo,TSPL_CUSTOMER_MASTER.Lst_No as Cust_LSTNo,TSPL_CUSTOMER_MASTER.Email as Cust_Email ,TSPL_CUSTOMER_MASTER.PAN as Customer_PAN,TSPL_CUSTOMER_MASTER.PIN_Code as Cust_PinCode,TSPL_CITY_MASTER.City_Name as Loctn_City_Name_Desc,TSPL_CUSTOMER_MASTER.Fax as Cust_Fax,TSPL_STATE_MASTER .STATE_NAME  as Cust_State_Name, "
        Qry += " TSPL_CITY_MASTER_ForCustomer.City_Name  as Cust_City_Name,case when (TSPL_DCS_SALE_ENTRY.Dispatch_Terms )='FE' then 'Freight Extra' else  TSPL_DCS_SALE_ENTRY.Dispatch_Terms  end  as Dispatch_Terms,TSPL_LOCATION_MASTER .Add1 as Loc_Add1,TSPL_LOCATION_MASTER.Add2 as Loc_ADd2,TSPL_LOCATION_MASTER.Add3  as Loc_Add3, StateMasterForLocation.State_Name as LocationState, TSPL_LOCATION_MASTER.Pin_Code as Loc_Pin_Code,TSPL_LOCATION_MASTER.TIN_No as Loc_TinNo,Case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as  Loc_Phn,TSPL_LOCATION_MASTER.Email as Loc_Email,TSPL_DCS_SALE_ENTRY.Total_Add_Charge,convert(varchar,TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Date ,103) as Delivery_Date, TSPL_DCS_SALE_ENTRY.Delivery_Code_PS,'NA' AS Ack_No,TSPL_DCS_SALE_ENTRY.Document_Date as Ack_Date,TSPL_DCS_SALE_ENTRY.Is_CashSale,TSPL_LOCATION_MASTER.Loc_Short_Name ,case when (TSPL_DCS_SALE_ENTRY.Payment_Terms)='A' then 'Advance' else TSPL_DCS_SALE_ENTRY.Payment_Terms end  as Payment_Terms ,TSPL_DCS_SALE_ENTRY.Transporter_Name    ,convert(varchar,TSPL_DCS_SALE_ENTRY.Sale_Invoice_Date,103) as Sale_Invoice_Date,TSPL_COMPANY_MASTER.Tin_No as Comp_Tin_No,TSPL_COMPANY_MASTER.Pan_No as Comp_PANNO,TSPL_COMPANY_MASTER.CST_LST as Comp_CST_No,TSPL_COMPANY_MASTER.CINNo as Comp_CinNo,TSPL_COMPANY_MASTER.Pincode  as Comp_Pin_Code, TSPL_SHIP_TO_LOCATION.Ship_To_Desc as shipName, TSPL_SHIP_TO_LOCATION.add1 as ship_Add1, TSPL_SHIP_TO_LOCATION.Add2 as ship_add2 ,TSPL_SHIP_TO_LOCATION.Add3 as ship_add3  ,TSPL_SHIP_TO_LOCATION.Pin_Code,TSPL_CITY_MASTER.STATE_CODE  ,Tspl_City_master.City_Name,TSPL_EMPLOYEE_MASTER.Emp_Name as SalesManName, TSPL_DCS_SALE_ENTRY.Dept_Desc , TSPL_DCS_SALE_ENTRY.Remarks ,  TSPL_DCS_SALE_ENTRY.Terms_Code,TSPL_DCS_SALE_ENTRY.VehicleNo ,TSPL_DCS_SALE_ENTRY.Challan_No,TSPL_DCS_SALE_ENTRY.RateDiff_Amt,TSPL_DCS_SALE_ENTRY.Gross_Amount,TSPL_VENDOR_MASTER.Zone_Code, "
        Qry += " TSPL_DCS_SALE_ENTRY_DETAIL .Specification as  specification,   TSPL_DCS_SALE_ENTRY.Document_Code as DocNo , TSPL_DCS_SALE_ENTRY.Description, "
        Qry += "  convert(varchar ,TSPL_DCS_SALE_ENTRY .Document_Date,103)as Document_Date , TSPL_DCS_SALE_ENTRY.Against_Sales_Order, TSPL_DCS_SALE_ENTRY.Item_Type ,  TSPL_DCS_SALE_ENTRY.Customer_Code, "
        Qry += " TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1,TSPL_CUSTOMER_MASTER.add2 as customer_Add2,TSPL_CUSTOMER_MASTER.Add3 as customer_Add3 ,TSPL_CUSTOMER_MASTER.State as customer_city_State,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_Tin_No ,TSPL_CUSTOMER_MASTER.PIN_Code as Customer_Pin_Code , TSPL_DCS_SALE_ENTRY .Terms_Code as termscode ,TSPL_DCS_SALE_ENTRY .Ref_No as ref_no ,"
        Qry += " TSPL_DCS_SALE_ENTRY .Comments as comments ,  TSPL_DCS_SALE_ENTRY .Discount_Amt as dis_amt,"
        If MultiplySubsidyWithQuantity Then
            Qry += "  isnull(TSPL_DCS_SALE_ENTRY.TotalSubsidyDisAmt,0)  as dis_amt1,"
        Else
            Qry += "  TSPL_DCS_SALE_ENTRY_DETAIL.Total_Disc_Amt  as dis_amt1,"
        End If
        Qry += " TSPL_DCS_SALE_ENTRY.Amount_Less_Discount  as aftrdiscount ,TSPL_DCS_SALE_ENTRY .Total_Amt as Total_amount,"
        Qry += " TSPL_DCS_SALE_ENTRY.Discount_Base as bfrdisc_amount, TSPL_COMPANY_MASTER.Access_Officer as FSSAI,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Tcan_No AS WebSite ,TSPL_COMPANY_MASTER.Phone1 AS COMP_PHONE,TSPL_DCS_SALE_ENTRY.Transporter_Commission_TotalAmt,TSPL_DCS_SALE_ENTRY.Recommended_By,  "
        Qry += " '' as tax1name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax1_amt,0) as txt1amt, TSPL_DCS_SALE_ENTRY_DETAIL.Tax1_Rate,  '' as tax2name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax2_amt,0) as txt2amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax2_Rate, '' as tax3name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax3_amt,0) as txt3amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax3_Rate, '' as tax4name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax4_amt,0) as txt4amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax4_Rate,'' as tax5name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax5_amt,0) as txt5amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax5_Rate,'' as tax6name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax6_amt,0) as txt6amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax6_Rate, '' as tax7name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax7_amt,0) as txt7amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax7_Rate, '' as tax8name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax8_amt,0) as txt8amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax8_Rate,'' as tax9name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax9_amt,0) as txt9amt,TSPL_DCS_SALE_ENTRY_DETAIL.Tax9_Rate,'' as tax10name,isnull (TSPL_DCS_SALE_ENTRY_DETAIL.Tax10_amt,0) as txt10amt, TSPL_DCS_SALE_ENTRY_DETAIL.Tax10_Rate, "
        Qry += " isnull(TSPL_DCS_SALE_ENTRY .Total_Tax_Amt,0) as total_tax_amt, TSPL_DCS_SALE_ENTRY.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+ Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as Phone,TSPL_COMPANY_MASTER.Fax as Comp_Fax,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.GSTReg_No As SellerGST,TSPL_COMPANY_MASTER.Pan_No,ISNULL(tspl_company_Master.ADD1,'') as Comp_add,"
        Qry += " TSPL_DCS_SALE_ENTRY_DETAIL.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc + case when TSPL_DCS_SALE_ENTRY_DETAIL.Scheme_Item ='Y' then ' (Free Scheme)' else '' end    as itemdesc, TSPL_DCS_SALE_ENTRY_DETAIL.Row_Type,TSPL_DCS_SALE_ENTRY_DETAIL.Qty as qty,TSPL_DCS_SALE_ENTRY_DETAIL.unit_code as uom,TSPL_DCS_SALE_ENTRY_DETAIL.item_cost as itemcost,TSPL_DCS_SALE_ENTRY_DETAIL.amount as amount,TSPL_DCS_SALE_ENTRY_DETAIL.Tax1,TSPL_DCS_SALE_ENTRY_DETAIL.Tax2,TSPL_DCS_SALE_ENTRY_DETAIL.Tax3,TSPL_DCS_SALE_ENTRY_DETAIL.Tax4,TSPL_DCS_SALE_ENTRY_DETAIL.Tax5,TSPL_DCS_SALE_ENTRY.ReceiverName, isnull(TSPL_DCS_SALE_ENTRY.TotalSubsidyAmt,0) as TotalSubsidyAmt "
        Qry += " " & QryShowStatus & " "
        Qry += " " & SerialNoColumn & "  "
        If isCancel Then
            Qry += " from TSPL_DCS_SALE_ENTRY_DETAIL_Cancel_Data As TSPL_DCS_SALE_ENTRY_DETAIL  "
        Else
            Qry += " from TSPL_DCS_SALE_ENTRY_DETAIL   "
        End If

        Qry += " " & SerialNo & " "
        If isCancel Then
            Qry += " left outer join TSPL_DCS_SALE_ENTRY_Cancel_Data As TSPL_DCS_SALE_ENTRY  on TSPL_DCS_SALE_ENTRY.Document_Code  =TSPL_DCS_SALE_ENTRY_DETAIL.Document_Code   "
        Else
            Qry += " left outer join TSPL_DCS_SALE_ENTRY  on TSPL_DCS_SALE_ENTRY.Document_Code  =TSPL_DCS_SALE_ENTRY_DETAIL.Document_Code   "
        End If
        Qry += "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Unit_code
       left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code in (select UOM_Code  from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Print_UOM=1)and TSPL_item_uom_detail.Print_UOM=1   ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code "
        Qry += " left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_DCS_SALE_ENTRY .Ship_To_Location "
        Qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_SHIP_TO_LOCATION.City_Code "
        Qry += " left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_DCS_SALE_ENTRY.Salesman_Code "
        Qry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_DCS_SALE_ENTRY.comp_code  "
        Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DCS_SALE_ENTRY.Customer_Code  left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = TSPL_CUSTOMER_MASTER.Cust_Code   LEFT join (select TSPL_CUSTOMER_MASTER.GSTNO AS P_GSTIN_NO,GST_STATE_CODE, pan as P_Cust_PAN, Cust_Code as P_cust_code,Customer_Name as P_cust_name,Add1 as P_cust_add1 ,Add2  as P_cust_add2,add3  as P_cust_add3,PIN_Code as P_Pin_No,Tin_No as p_Tin_No ,State as P_state,Email as P_Email,fax as P_Fax,TSPL_CITY_MASTER.City_Name as  P_City_Name,TSPL_STATE_MASTER.STATE_NAME as P_State_Name  ,case when ISNULL(Phone1,'')='(+__)__________' then '' else Phone1 end +  Case When   ISNULL(Phone2,'')<>'(+__)__________' Then ', '+ Phone2 Else'' End as  P_Phn,CST as P_CST_No,Terms_Code as P_Terms,Lst_No as P_LST_No  from TSPL_CUSTOMER_MASTER   left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_CUSTOMER_MASTER.City_Code "
        Qry += " left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State =TSPL_STATE_MASTER.STATE_CODE  "
        Qry += " ) p_cust on p_cust.P_cust_code=TSPL_CUSTOMER_MASTER.Parent_Customer_No and TSPL_CUSTOMER_MASTER.Parent_Customer_YN='N'"
        Qry += "  LEFT OUTER JOIN tspl_customer_vendor_mapping ON tspl_customer_vendor_mapping.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code   "
        Qry += "LEFT OUTER JOIN tspl_vlc_master_head ON tspl_customer_vendor_mapping.Vendor_Code =tspl_vlc_master_head.vsp_code"
        Qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_DCS_SALE_ENTRY.Bill_To_Location "
        Qry += " left outer join TSPL_CITY_MASTER  as TSPL_CITY_MASTER_ForCustomer on TSPL_CITY_MASTER_ForCustomer.City_Code =TSPL_CUSTOMER_MASTER.City_Code"
        Qry += " left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE = TSPL_CUSTOMER_MASTER.State"
        Qry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_DCS_SALE_ENTRY_DETAIL.Item_Code "
        Qry += " left outer join TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE on TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code =TSPL_DCS_SALE_ENTRY.Delivery_Code_PS"
        Qry += " LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code "
        Qry += " LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State "
        Qry += " LEFT OUTER JOIN TSPL_STATE_MASTER StateMasterForLocation On StateMasterForLocation.State_Code=TSPL_LOCATION_MASTER.State 
left join ( select Document_Code,Parent_Line_No, STRING_AGG(Batch_No +' -'+convert(varchar(10),Qty)+'', ',') as Batch_No from(
SELECT Document_Code, Batch_No, Qty, Parent_Line_No FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='DCS-SAL-ENT'
)x group by Document_Code,Parent_Line_No                
) TabBatch On TabBatch.Document_Code= TSPL_DCS_SALE_ENTRY.Document_Code And TabBatch.Parent_Line_No = TSPL_DCS_SALE_ENTRY_DETAIL.Line_No "
        Qry += "  where 2=2 "
        Qry += "  And  TSPL_DCS_SALE_ENTRY.Document_Code = '" + strCode + "' )xx "
        Qry += " left join ( " + IIf(clsCommon.myLen(TotalTaxQry) > 0, TotalTaxQry, " select '' as DOCUMENT_CODE,0 as DTax1_Amt,0 as DTax1_Rate,'' as tax1Type,0 as DTax2_Amt,0 as DTax2_Rate,'' as tax2Type,0 as DTax3_Amt,0 as DTax3_Rate,'' as tax3Type,0 as DTax4_Amt,0 as DTax4_Rate,'' as tax4Type,0 as DTax5_Amt,0 as DTax5_Rate,'' as tax5Type,0 as	DTax6_Amt,0 as DTax6_Rate,'' as	tax6Type,0 as	DTax7_Amt,0 as DTax7_Rate,	'' as tax7Type	,0 as	DTax8_Amt,0 as DTax8_Rate,'' as tax8Type,0 as	DTax9_Amt,0 as DTax9_Rate,'' as	tax9Type,0 as	DTax10_Amt,0 as	DTax10_Rate,'' as	tax10Type ") + " from ( " + TaxQry + " )xxxx )as tabTax on tabTax.DOCUMENT_CODE= xx.DocNo"
        Return Qry
    End Function
    Public Shared Function funPrint(ByVal Form_ID As String, ByVal isCancel As Boolean, ByVal strDate As DateTime, ByVal strCode As String, GetPDFPath As Boolean) As String
        Dim PdfPath As String = ""
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetAttachQry(isCancel, strDate, strCode))

            If dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If GetPDFPath Then
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                        PdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceDCSSaleEntryJDHNEW", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                    Else
                        PdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDCSSaleEntry", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                    End If
                    ' PdfPath = frmCRV.funsubreportWithdt(Form_ID, GetPDFPath, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDCSSaleEntry", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                Else
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceDCSSaleEntryJDHNEW", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDCSSaleEntryNAG", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")

                    Else
                            frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDCSSaleEntry", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                    End If

                    ' frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDCSSaleEntry", "DCS Sale Entry", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                End If
                    frmCRV = Nothing
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        If GetPDFPath Then
            Return PdfPath
        Else
            Return True
        End If
    End Function

End Class

Public Class clsDCSSaleEntryDetail
#Region "Variables"
    Public ItemwiseTaxCode As String = Nothing
    Public PK_ID As Integer
    Public OrgUnit_code As String = ""
    Public Commission_Party As String = Nothing
    Public Commission_Rate As Double = 0
    Public Commission_Amt As Double = 0
    Public Amt_Less_Commission As Double = 0
    Public Deduction As String = Nothing
    Public Deduction_Name As String = Nothing
    Public Deduction_Type As String = Nothing
    Public Tax_Group As String = Nothing
    Public TaxGroupName As String = Nothing
    Public Document_Code As String = Nothing
    Public Line_No As Integer = 0
    Public Row_Type As String = Nothing
    Public Item_Code As String = Nothing
    Public Item_Desc As String = Nothing
    Public DCS_Sale_Zero_Cost As String = Nothing
    Public Bar_Code As String = Nothing
    Public Qty As Double = 0
    Public Balance_Qty As Double = 0
    Public Free_Qty As Double = 0
    Public Order_Code As String = Nothing
    Public Unit_code As String = Nothing '
    Public Location As String = Nothing '
    Public LocationName As String = Nothing
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
    Public TotalSubsidyAmt As Double = 0
    Public TotalSubsidyDisAmt As Double = 0
    Public Transporter_Commission_Rate As Double = 0
    Public Transporter_Commission_Amt As Double = 0
    Public RoundOffAmount As Double = 0
    Public REF_TPT_PK_ID As Integer = 0
    Public RateDiff_Per As Double = 0
    Public RateDiff_Amt As Double = 0
    Public Gross_Amount As Double = 0
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
    Public PrincipleDesc As String = Nothing
    Public vendor_code As String = Nothing
    Public vendor_desc As String = Nothing
    Public Bin_No As String = Nothing
    Public HeadDiscPer As Double = 0
    Public HeadDiscPerAmt As Double = 0
    Public Is_Add_TPT As Boolean = False
    Public arrBatchItem As List(Of clsBatchInventory) = Nothing
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsDCSSaleEntryDetail), ByVal trans As SqlTransaction, ByVal DocDate As DateTime) As Boolean

        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then

            For Each obj As clsDCSSaleEntryDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "ItemwiseTaxCode", obj.ItemwiseTaxCode)
                clsCommon.AddColumnsForChange(coll, "Commission_Rate", obj.Commission_Rate)
                clsCommon.AddColumnsForChange(coll, "Commission_Party", obj.Commission_Party)
                clsCommon.AddColumnsForChange(coll, "Commission_Amt", obj.Commission_Amt)
                clsCommon.AddColumnsForChange(coll, "Amt_Less_Commission", obj.Amt_Less_Commission)
                clsCommon.AddColumnsForChange(coll, "OrgUnit_code", obj.OrgUnit_code)
                clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                clsCommon.AddColumnsForChange(coll, "Row_Type", obj.Row_Type)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
                clsCommon.AddColumnsForChange(coll, "Deduction", obj.Deduction)
                clsCommon.AddColumnsForChange(coll, "Deduction_Type", obj.Deduction_Type, True)
                clsCommon.AddColumnsForChange(coll, "Bar_Code", obj.Bar_Code, True)
                clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)

                clsCommon.AddColumnsForChange(coll, "Free_qty", obj.Free_Qty)

                clsCommon.AddColumnsForChange(coll, "Order_Code", obj.Order_Code, True)

                clsCommon.AddColumnsForChange(coll, "Balance_Qty", obj.Balance_Qty)
                clsCommon.AddColumnsForChange(coll, "Unit_code", obj.Unit_code)
                clsCommon.AddColumnsForChange(coll, "Location", obj.Location)
                clsCommon.AddColumnsForChange(coll, "Item_Cost", obj.Item_Cost)

                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "TotalSubsidyDisAmt", obj.TotalSubsidyDisAmt)
                clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Rate", obj.Transporter_Commission_Rate)
                clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Amt", obj.Transporter_Commission_Amt)
                clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
                clsCommon.AddColumnsForChange(coll, "REF_TPT_PK_ID", obj.REF_TPT_PK_ID, True)
                clsCommon.AddColumnsForChange(coll, "RateDiff_Per", obj.RateDiff_Per)
                clsCommon.AddColumnsForChange(coll, "RateDiff_Amt", obj.RateDiff_Amt)
                clsCommon.AddColumnsForChange(coll, "TotalSubsidyAmt", obj.TotalSubsidyAmt)
                clsCommon.AddColumnsForChange(coll, "Gross_Amount", obj.Gross_Amount)
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
                clsCommon.AddColumnsForChange(coll, "PrincipleDesc", obj.PrincipleDesc)
                clsCommon.AddColumnsForChange(coll, "vendor_code", obj.vendor_code)
                clsCommon.AddColumnsForChange(coll, "vendor_desc", obj.vendor_desc)
                clsCommon.AddColumnsForChange(coll, "Bin_No", obj.Bin_No)
                clsCommon.AddColumnsForChange(coll, "HeadDiscPer", obj.HeadDiscPer)
                clsCommon.AddColumnsForChange(coll, "HeadDiscPerAmt", obj.HeadDiscPerAmt)
                clsCommon.AddColumnsForChange(coll, "Is_Add_TPT", IIf(obj.Is_Add_TPT, 1, 0))
                'clsCommon.AddColumnsForChange(coll, "Unit_Cogs", clsItemLocationDetails.GetUnitCogs(obj.Item_Code, obj.Location, trans))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_SALE_ENTRY_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                clsSerializeInvenotry.SaveData("DCS-SAL-ENT", strDocNo, DocDate, "O", obj.Item_Code, obj.Location, obj.Line_No, obj.arrSrItem, trans)
                clsBatchInventory.SaveData("DCS-SAL-ENT", strDocNo, DocDate, "O", obj.Item_Code, obj.Location, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans) ' Change by Prabhakar
            Next
        End If
        Return True
    End Function

    Public Shared Function CompleteSRN(ByVal strDoccode As String, ByVal strICode As String, ByVal LineNo As Integer) As Boolean
        Dim qry As String = "update TSPL_DCS_SALE_ENTRY_DETAIL set Status=1 where Document_Code='" + strDoccode + "' and Line_No='" + clsCommon.myCstr(LineNo) + "' and Item_Code='" + strICode + "'"
        Return clsDBFuncationality.ExecuteNonQuery(qry)
    End Function
End Class
