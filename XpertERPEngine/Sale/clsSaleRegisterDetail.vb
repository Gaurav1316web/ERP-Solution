
'===========Rohit,BM00000007332,BM00000007715,BM00000007628==============On 26 Aug,2015=======
Imports common
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Telerik.WinControls

Public Class clsSaleRegisterDetail
#Region "Variables"
    Public IsSameBillShipParty As Integer = 0
    Public Ship_To_Party As String = Nothing
    Public Scheme_Tax_Group As String = Nothing
    Public Electronic_Ref_No As String = Nothing
    Public EWayBillDate As Date?
    Public EWayBillNo As String = Nothing
    Public GSTStatus As Boolean = False
    Public Is_Taxable As Integer = 0
    Public Nine_NR_No As String = Nothing
    Public Including_Insurance As Integer = 0
    Public VAT_InvoiceNo As String = Nothing
    Public VatInvoice_Type As String = Nothing
    Public Crate As Double = 0
    Public jaali As Double = 0
    Public Box As Double = 0
    Public Is_CustomerChanged As Integer = 0
    Public RoundOffAmount As Double = 0
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
    Public Is_Delivered As Integer = 0
    Public podate As DateTime? = Nothing
    Public Form_38_No As String = Nothing
    Public Cust_PO_No As String = Nothing
    Public Price_Group_Code As String = Nothing
    Public PROJECT_ID As String = Nothing
    Public Invoice_Type As String = Nothing
    Public Mannual_Document_Code As Integer
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
    Public Against_Shipment_No As String = Nothing
    Public Is_Internal As Boolean = False
    Public Tax_Calculation_Type As EnumTaxCalucationType
    Public Is_Create_Auto_Receipt As Boolean = False
    Public Salesman_Code As String = Nothing
    Public Salesman_Name As String = Nothing
    Public Arr As List(Of clsPSInvoiceHeadDetail) = Nothing

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing
    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public Against_C_Form As Boolean = False
    Public Price_Code As String = Nothing
    Public Route_No As String = Nothing
    Public Route_Desc As String = Nothing
    Public HeadDisc_Per As Double = 0
    Public HeadDisc_Amt As Double = 0
    Public HeadDisc_PerAmt As Double = 0
    Public TotCashDiscAmt As Double = 0
    Public Item_Tax_Type As Integer = 0
    Public InvoiceManualNowithPrefix As String = Nothing
    Public CancelFlag As String = Nothing
    '================Rohit ,add Trans_type to Get Type(PS,MCC) in Variable,Nov 10,2014===========
    Public Trans_type As String = String.Empty
    '======================================================
    Public Invoice_No_For_Supplementary As String = String.Empty
    Public Supplementary_Type As String = ""
    Public No_Of_Instalment As Integer
#End Region


    Public Shared Function GetMIS_ITem_GroupColumn() As String
        Dim MIS_Item_Group As String = ""
        Dim qry As String = ""
        qry = " select MAP.Custom_Field_Code from TSPL_CUSTOM_FIELD_MAPPING MAP " &
            " left join TSPL_CUSTOM_FIELD_HEAD CF on MAP.Custom_Field_Code=CF.Code " &
            " where CF.Name='MIS Item Group' and MAP.PROGRAM_CODE='" & clsUserMgtCode.itemStructure & "'"
        MIS_Item_Group = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Nothing))
        Return MIS_Item_Group
    End Function
    Public Shared Function GetPivotForFinalOuterQry(ByVal obj As clsSaleRegisterParameterType) As String
        Dim qryTaxQuery As String = ""
        Dim strPivotForFinalOuterQuery As String = ""
        Dim strPivotForFinalOuter As String = ""
        qryTaxQuery = GetTaxQuery(obj)
        strPivotForFinalOuter = " select distinct (select Distinct ',xx.['+tax1+']' from ( " & qryTaxQuery
        strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        strPivotForFinalOuterQuery = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))
        Return strPivotForFinalOuterQuery
    End Function
    Public Shared Function GetPivotForFinalOuterQryForDocumnetInfoLevel(ByVal obj As clsSaleRegisterParameterType) As String
        Dim qryTaxQuery As String = ""
        Dim strPivotForFinalOuterQuery As String = ""
        Dim strPivotForFinalOuter As String = ""
        qryTaxQuery = GetTaxQuery(obj)
        strPivotForFinalOuter = " select distinct (select Distinct ',xx.['+tax1+']' from ( " & qryTaxQuery
        strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        strPivotForFinalOuterQuery = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))
        Return strPivotForFinalOuterQuery
    End Function
    Public Shared Function GetTaxQuery(ByVal obj As clsSaleRegisterParameterType) As String
        Dim qryTaxQuery As String = ""
        Dim lstTables As New List(Of String)
        Dim lstTableDocDateCols As New List(Of String)
        If obj.Trans_Type_List.Contains("Dairy Sale") OrElse obj.Trans_Type_List.Contains("Fresh Sale") OrElse obj.Trans_Type_List.Contains("Product Sale") OrElse obj.Trans_Type_List.Contains("MCC Sale") OrElse obj.Trans_Type_List.Contains("Export Sale") OrElse obj.Trans_Type_List.Contains("CSA Sale") OrElse obj.Trans_Type_List.Contains("Merchant Trade") Then
            lstTables.Add("TSPL_SD_SALE_INVOICE_HEAD")
            lstTableDocDateCols.Add("Document_Date")
        End If
        If obj.Trans_Type_List.Contains("Misc Sale") Then
            lstTables.Add("TSPL_SCRAPINVOICE_HEAD")
            lstTableDocDateCols.Add("shipment_Date")
        End If
        If obj.Trans_Type_List.Contains("Fresh Sale Return") OrElse obj.Trans_Type_List.Contains("CSA Sale Return") OrElse obj.Trans_Type_List.Contains("Product Sale Return") OrElse obj.Trans_Type_List.Contains("MCC Sale Return") OrElse obj.Trans_Type_List.Contains("Export Sale Return") OrElse obj.Trans_Type_List.Contains("Sale Return") OrElse obj.Trans_Type_List.Contains("Bulk Sale Return") Then
            lstTables.Add("TSPL_SD_SALE_RETURN_HEAD")
            lstTableDocDateCols.Add("Document_Date")
        End If

        If obj.Trans_Type_List.Contains("Transfer") OrElse obj.Trans_Type_List.Contains("Transfer Return") Then
            lstTables.Add("TSPL_TRANSFER_ORDER_HEAD")
            lstTableDocDateCols.Add("Document_Date")
        End If

        If obj.Trans_Type_List.Contains("CSA Sale") Then
            lstTables.Add("TSPL_CSA_TRANSFER_HEAD")
            lstTableDocDateCols.Add("Transfer_Date")
        End If

        If obj.Trans_Type_List.Contains("Can Sale") Then
            lstTables.Add("TSPL_CANSALE_INVOICE_HEAD")
            lstTableDocDateCols.Add("Document_Date")
        End If
        If obj.Trans_Type_List.Contains("Bulk Sale") OrElse obj.Trans_Type_List.Contains("Bulk Sale Trade") Then
            lstTables.Add("TSPL_INVOICE_Master_BULKSALE")
            lstTableDocDateCols.Add("Document_Date")
        End If

        qryTaxQuery = GetTaxQuery(lstTables, lstTableDocDateCols, obj)
        Return qryTaxQuery
    End Function
    Public Shared Function GetTrasactionItemCodeQry(ByVal obj As clsSaleRegisterParameterType) As String
        Dim qry As String = ""
        Dim FirstQuery As Boolean = False
        'qry = " SELECT DISTINCT xxx.Code,  xxx.Name ,xxx.[Document Code],xxx.[Document Date],xxx.[Customer Code] FROM (" + Environment.NewLine
        If obj.Trans_Type_List.Contains("General Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='ALL'  and Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        If obj.Trans_Type_List.Contains("Fresh Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='FS' and Screen_Type='' and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Product Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='PS' and Screen_Type='' and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Dairy Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type IN ('FS','PS') and Screen_Type='DS' and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("MCC Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='MCC'  and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("CSA Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='CSA'  and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        If obj.Trans_Type_List.Contains("Export Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type ='EX'  and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Merchant Trade") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code From TSPL_SD_SALE_INVOICE_DETAIL INNER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE Where Trans_Type='EXP'  and TSPL_SD_SALE_INVOICE_HEAD.Document_Type ='MT' and cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Bulk Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT distinct TSPL_INVOICE_DETAIL_BULKSALE.Item_Code FROM TSPL_INVOICE_DETAIL_BULKSALE inner join TSPL_INVOICE_MASTER_BULKSALE on  TSPL_INVOICE_DETAIL_BULKSALE.Document_No=TSPL_INVOICE_MASTER_BULKSALE.Document_No where cast(TSPL_INVOICE_MASTER_BULKSALE.Document_Date as date)  between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Transfer") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " select DISTINCT Item_Code from TSPL_TRANSFER_ORDER_DETAIL INNER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No=TSPL_TRANSFER_ORDER_HEAD.Document_No WHERE CAST(TSPL_TRANSFER_ORDER_HEAD.Document_Date AS DATE) BETWEEN '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Misc Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " select DISTINCT Item_Code from TSPL_SCRAPINVOICE_DETAIL INNER JOIN TSPL_SCRAPINVOICE_HEAD ON TSPL_SCRAPINVOICE_DETAIL.invoice_No=TSPL_SCRAPINVOICE_HEAD.invoice_No" &
                  " WHERE CAST(TSPL_SCRAPINVOICE_HEAD.shipment_Date AS DATE) BETWEEN '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Misc Sale") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " select DISTINCT Item_Code from TSPL_SCRAPINVOICE_DETAIL INNER JOIN TSPL_SCRAPINVOICE_HEAD ON TSPL_SCRAPINVOICE_DETAIL.invoice_No=TSPL_SCRAPINVOICE_HEAD.invoice_No" &
                  " WHERE CAST(TSPL_SCRAPINVOICE_HEAD.shipment_Date AS DATE) BETWEEN '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='ALL' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        If obj.Trans_Type_List.Contains("Fresh Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Product Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='PS' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        If obj.Trans_Type_List.Contains("MCC Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        If obj.Trans_Type_List.Contains("CSA Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='CSA' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Export Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='EXP' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("CSA Sale Patti Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SD_SALE_RETURN_DETAIL.Item_Code FROM TSPL_SD_SALE_RETURN_DETAIL INNER JOIN TSPL_SD_SALE_RETURN_HEAD ON TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code WHERE TSPL_SD_SALE_RETURN_HEAD.Trans_Type='CPR' and  CAST(TSPL_SD_SALE_RETURN_HEAD.Document_Date AS DATE) BETWEEN  '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If

        If obj.Trans_Type_List.Contains("Bulk Sale Return") Then
            If FirstQuery Then
                qry = qry & " Union All "
            End If
            FirstQuery = True
            qry = qry & " SELECT DISTINCT TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code FROM TSPL_SALE_RETURN_DETAIL_BULKSALE INNER JOIN TSPL_SALE_RETURN_MASTER_BULKSALE ON TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No=TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No WHERE CAST(TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date AS DATE) BETWEEN '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "'"
        End If
        Return qry

    End Function


    Public Shared Function GetAllSaleTransactionTypeQuery(Optional ByVal DocFinder As Boolean = False) As String
        Dim qry As String
        Dim Donotshowtrasnfertransactionsbydefault As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.Donotshowtrasnfertransactionsbydefault, clsFixedParameterCode.Donotshowtrasnfertransactionsbydefault, Nothing)) = "1", True, False))
        If DocFinder = True Then
            qry = " SELECT DISTINCT xxx.Code,  xxx.Name " & If(DocFinder = False, "", ",xxx.[Document Code],xxx.[Document Date],xxx.[Customer Code]") & " FROM (" + Environment.NewLine &
                       " SELECT DISTINCT 'SD' As Code , 'General Sale' As Name " & If(DocFinder = False, "", ",Document_Code as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " From  TSPL_SD_SALE_INVOICE_HEAD Where Trans_Type='ALL' " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT Trans_Type+ISNULL(Document_Type,'')  As Code ,CASE " &
                       " WHEN Trans_Type ='CSA' THEN 'CSA Sale' " &
                       " WHEN Trans_Type ='EXP' AND Document_Type ='EX' THEN 'Export Sale' " &
                       " WHEN Trans_Type ='FS' THEN 'Fresh Sale'  " &
                       " WHEN Trans_Type ='MCC' THEN 'MCC Sale' " &
                       " WHEN Trans_Type ='PS' THEN 'Product Sale' " &
                       " WHEN Trans_Type ='EXP' AND Document_Type ='MT' THEN 'Merchant Trade' " &
                       " End " &
                       " As Name " & If(DocFinder = False, "", ",Document_Code as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " FROM TSPL_SD_SALE_INVOICE_HEAD Where Trans_Type <>'ALL'  union all select distinct 'CSA' as Code,'CSA Sale' as Name " & If(DocFinder = False, "", ",DOC_CODE as [Document Code],TRANSFER_DATE  as [Document Date],CUST_CODE as [Customer Code]") & " from tspl_csa_transfer_head " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'BS'  As Code ,'Bulk Sale' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " FROM TSPL_INVOICE_MASTER_BULKSALE  " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'BST'  As Code ,'Bulk Sale Trade' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " FROM TSPL_INVOICE_MASTER_BULKSALE " &
                       " UNION ALL " + Environment.NewLine &
                       " Select DISTINCT 'SS'  As Code ,'Misc Sale' As Name " & If(DocFinder = False, "", ",invoice_No as [Document Code],posting_Date  as [Document Date],cust_Code as [Customer Code]") & " FROM TSPL_SCRAPINVOICE_HEAD " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'BSR'  As Code ,'Bulk Sale Return' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " FROM TSPL_SALE_RETURN_MASTER_BULKSALE " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT  DISTINCT 'SDR'  As Code ,'General Sale Return'  As Name " & If(DocFinder = False, "", ",Document_Code as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " From TSPL_SD_SALE_RETURN_HEAD Where Trans_Type='ALL' " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT Trans_Type+'R' As Code ,CASE " &
                       " WHEN Trans_Type ='CSA' THEN 'CSA Sale Return' " &
                       " WHEN Trans_Type ='EXP' THEN 'Export Sale Return' " &
                       " WHEN Trans_Type ='FS'  THEN 'Fresh Sale Return'  " &
                       " WHEN Trans_Type ='MCC' THEN 'MCC Sale Return' " &
                       " WHEN Trans_Type ='PS'  THEN 'Product Sale Return' " &
                       " WHEN Trans_Type ='CPR'  THEN 'CSA Sale Patti Return' " &
                       " END " &
                       " As Name " & If(DocFinder = False, "", ",Document_Code as [Document Code],Document_Date  as [Document Date],Customer_Code as [Customer Code]") & " FROM TSPL_SD_SALE_RETURN_HEAD Where Trans_Type not in ('ALL','CPR') " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'Transfer'  As Code ,'Transfer' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],'' as [Customer Code]") & " FROM TSPL_TRANSFER_ORDER_HEAD " &
                        " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'Transfer Return'  As Code ,'Transfer Return' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],'' as [Customer Code]") & " FROM TSPL_TRANSFER_RETURN " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'MCC Transfer' As Code ,'MCC Transfer' As Name " & If(DocFinder = False, "", ",Receipt_Challan_No as [Document Code],Receipt_Challan_Date  as [Document Date],'' as [Customer Code]") & " FROM TSPL_MILK_TRANSFER_IN " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'Tanker Dispatch Return' As Code ,'Tanker Dispatch Return' As Name " & If(DocFinder = False, "", ",Document_No as [Document Code],Document_Date  as [Document Date],'' as [Customer Code]") & " FROM TSPL_MCC_Dispatch_Challan_Return " &
                       " UNION ALL " + Environment.NewLine &
                       " SELECT DISTINCT 'MCC Tanker Dispatch Return' As Code ,'MCC Tanker Dispatch Return' As Name " & If(DocFinder = False, "", ",Return_NO as [Document Code],Return_Date  as [Document Date],'' as [Customer Code]") & " FROM TSPL_MCC_Tanker_Dispatch_Return_head " &
                       " ) xxx "
        Else
            ' qry = "" & Environment.NewLine &
            '" select 'BS' as Code,'Bulk Sale' as Name " & Environment.NewLine &
            ' " union all" & Environment.NewLine & _
            '" select 'BSR','Bulk Sale Return' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            '" select 'BST' as Code,'Bulk Sale Trade' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            '" select 'CSA' as Code,'CSA Sale' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            '" select 'CSAR' as Code,'CSA Sale Return' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            '" select 'EXPEX','Export Sale' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            '" select 'EXPMT' as Code,'Merchant Trade' as Name " & Environment.NewLine & _
            '" Union All " & Environment.NewLine & _
            qry = " select 'FS' as Code,'Fresh Sale' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'FSR','Fresh Sale Return' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'MCC','MCC Sale' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                 " select 'MCCR','MCC Sale Return' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'PS' as Code,'Product Sale' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'PSR','Product Sale Return' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'SS' as Code,'Misc Sale' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'MCCFS' as Code,'MCC Sale Farmer' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'MCCFSR' as Code,'MCC Sale Return Farmer' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'DS' as Code,'Dairy Sale' as Name " &
                  " Union All " & Environment.NewLine &
                  " select 'DSR','Dairy Sale Return' as Name "
            '" Union All " & Environment.NewLine &
            '    " select 'CanSaleInvoice' as Code,'Can Sale' as Name "

            ''richa agarwal 12 Aug,2019 ERO/05/08/19-000985
            If Donotshowtrasnfertransactionsbydefault = True Then
                qry += " Union All " & Environment.NewLine &
                  " select 'MCC Transfer','MCC Transfer' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'Tanker Dispatch Return','Tanker Dispatch Return' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                   " select 'Transfer' as Code,'Transfer' as Name " & Environment.NewLine &
                  " Union All " & Environment.NewLine &
                  " select 'Transfer Return' as Code,'Transfer Return' as Name " & Environment.NewLine
            End If
        End If

        Return qry
    End Function

    Public Shared Function ReturnQuery(ByVal obj As clsSaleRegisterParameterType) As ArrayList
        Dim Stock_uom As Boolean = obj.stockinguom
        Dim From_Date As Date = obj.From_Date
        Dim To_Date As Date = obj.To_Date
        Dim Unit_Code As String = obj.Unit_Code
        Dim QryLst As New ArrayList

        Dim strCodeColumn As String = ""
        Dim strCodeColumnForVirtual As String = ""
        Dim strCodeColumnMax As String = ""
        Dim strCodeDescColumn As String = ""
        Dim strCodeDescColumnMax As String = ""
        Dim strPivotForFinalOuterQuery As String = ""
        Dim strCustomerAmountWiseTaxQuery As String = ""
        Dim strCategoryTable As String = ""
        Dim MIS_Item_Group As String = GetMIS_ITem_GroupColumn()
        Dim dtCategory As New DataTable
        Dim qryStarted As Boolean = False
        Dim ItemCategoryCondQry As String = GetTrasactionItemCodeQry(obj)
        Dim Batch_Wise As Boolean = obj.BatchWise
        ''====================Monika 11/04/2017
        If obj.rbtnCategorySected OrElse (clsCommon.CompairString(obj.ReportType, "Document Info Level") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.ReportType, "Customer Amount Wise") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.ReportType, "Document Detail") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.ReportType, "Net Sale") = CompairStringResult.Equal) Then ''if categories are checked from screen only then show as pivot.Or if document detail/info report opened then category pivot run
            dtCategory = clsDBFuncationality.GetDataTable("select ITEM_CATEGORY_CODE AS CodeColumn,ITEM_CATEGORY_CODE+' Desc' as CodeDescColumn,DESCRIPTION as DescColumn  from TSPL_ITEM_CATEGORY_LEVEL order by CATEGORY_LEVEL")
            If dtCategory IsNot Nothing AndAlso dtCategory.Rows.Count > 0 Then
                For ii As Integer = 0 To dtCategory.Rows.Count - 1
                    If ii <> 0 Then
                        strCodeColumn += ","
                        strCodeColumnForVirtual += ","
                        strCodeColumnMax += ","
                        strCodeDescColumn += ","
                        strCodeDescColumnMax += ","
                    End If
                    strCodeColumn += "[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeColumnForVirtual += "VirtualCategoryTabel.[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeColumnMax += "max([" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]) as [" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeDescColumn += "[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")) + "]"
                    strCodeDescColumnMax += "max([" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")).Trim() + "]) as [" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")).Trim() + "]"
                Next
                strCategoryTable = "select Item_Code," + strCodeColumnMax + "," + strCodeDescColumnMax + "  from (" + Environment.NewLine &
                " select * from ( " + Environment.NewLine &
                " select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code " + Environment.NewLine &
                " ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+' Desc' as Item_Category_CodeDesc " + Environment.NewLine &
                " ,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " + Environment.NewLine &
                " ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc " + Environment.NewLine &
                " from  TSPL_ITEM_MASTER  " + Environment.NewLine &
                " left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code " + Environment.NewLine &
                " left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values" + Environment.NewLine &
                " where 2=2 " + Environment.NewLine &
                " )xx" + Environment.NewLine &
                " Pivot " + Environment.NewLine &
                " ( max(Item_Cagetory_Values) for Item_Category_Code   in ( " + strCodeColumn + ")" + Environment.NewLine &
                " ) Pivt" + Environment.NewLine &
                " Pivot " + Environment.NewLine &
                " (" + Environment.NewLine &
                " max(Category_Value_Desc) for Item_Category_CodeDesc in (" + strCodeDescColumn + ")" + Environment.NewLine &
                " ) Pivt1 " + Environment.NewLine &
                " ) xxx group by Item_Code "
                ''End of Category Table start now.
            End If
        End If
        ''and TSPL_ITEM_MASTER.Item_Code in (" & ItemCategoryCondQry & " )
        ''Virtual Category Table start now.


        Dim strMCCMaterial As String = ""
        Dim qryTaxQuery As String = ""
        'Dim strPivotForOuter As String

        qryTaxQuery = GetTaxQuery(obj)

        'strPivotForOuter = "select distinct (select Distinct ',sum(isnull(final.['+tax1+'],0)) as ['+TAX1+']' from ( " & qryTaxQuery
        'strPivotForOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim dtTax As DataTable
        If clsCommon.myLen(qryTaxQuery) > 0 Then
            dtTax = clsDBFuncationality.GetDataTable(qryTaxQuery)
            If dtTax.Rows.Count <= 0 Then
                qryTaxQuery = "select 'EXEMPTED' as Tax1"
                dtTax = clsDBFuncationality.GetDataTable(qryTaxQuery)
            End If
        Else
            qryTaxQuery = "select 'EXEMPTED' as Tax1"
            dtTax = clsDBFuncationality.GetDataTable(qryTaxQuery)
            'dtTax = New DataTable
            'dtTax.Columns.Add("Tax1")
        End If

        '' create strPivotForOuterQuery
        Dim strPivotForOuterQuery As String = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotForOuterQuery = strPivotForOuterQuery & "sum(isnull(final.[" & clsCommon.myCstr(dr.Item("TAX1")) & "],0)) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
            Else
                strPivotForOuterQuery = strPivotForOuterQuery & "sum(isnull(final.[" & clsCommon.myCstr(dr.Item("TAX1")) & "],0)) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
            End If

        Next
        'Dim strPivotForOuterQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForOuter))
        '============================

        'Dim strPivotForFinalOuter As String
        'strPivotForFinalOuter = ""
        'strPivotForFinalOuter = " select distinct (select Distinct ',xx.['+tax1+']' from ( " & qryTaxQuery
        'strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        'strPivotForFinalOuterQuery = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))
        '' create strPivotForOuterQuery
        strPivotForFinalOuterQuery = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                'strPivotForFinalOuterQuery = strPivotForFinalOuterQuery & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
                strPivotForFinalOuterQuery = strPivotForFinalOuterQuery & " case when Scheme_Item='Y' then 0 else " & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "]" & " end " & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
                'strCustomerAmountWiseTaxQuery = " Sum ([" & clsCommon.myCstr(dr.Item("TAX1")) & "]) [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
                strCustomerAmountWiseTaxQuery = strCustomerAmountWiseTaxQuery & " Sum ([" & clsCommon.myCstr(dr.Item("TAX1")) & "]) [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"

            Else
                'strPivotForFinalOuterQuery = strPivotForFinalOuterQuery & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
                strPivotForFinalOuterQuery = strPivotForFinalOuterQuery & " case when Scheme_Item='Y' then 0 else " & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "]" & " end " & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
                'strCustomerAmountWiseTaxQuery = " Sum ([" & clsCommon.myCstr(dr.Item("TAX1")) & "]) [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
                strCustomerAmountWiseTaxQuery = strCustomerAmountWiseTaxQuery & " Sum ([" & clsCommon.myCstr(dr.Item("TAX1")) & "]) [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"

            End If
        Next
        obj.strPivotForFinalOuterQuery = strPivotForFinalOuterQuery
        obj.strCustomerAmountWiseTaxQuery = strCustomerAmountWiseTaxQuery
        'Dim strPivotForFinalOuterPercent As String
        'strPivotForFinalOuterPercent = " select distinct (select  Distinct ',xx.['+tax1+'%'+']' from ( " & qryTaxQuery
        'strPivotForFinalOuterPercent += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        Dim strPivotForFinalOuterPercentQuery As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuterPercent))
        '' update strPivotForFinalOuterPercentQuery
        strPivotForFinalOuterPercentQuery = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                'strPivotForFinalOuterPercentQuery = strPivotForFinalOuterPercentQuery & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
                strPivotForFinalOuterPercentQuery = strPivotForFinalOuterPercentQuery & " case when Scheme_Item='Y' then 0 else " & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%]" & " end " & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                'strPivotForFinalOuterPercentQuery = strPivotForFinalOuterPercentQuery & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
                strPivotForFinalOuterPercentQuery = strPivotForFinalOuterPercentQuery & " case when Scheme_Item='Y' then 0 else " & "xx.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%]" & " end " & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next

        Dim strPivotForTransfer_In As String
        'strPivotForFinalOuter = ""
        'strPivotForFinalOuter = " select distinct (select Distinct ',0 as ['+tax1+']' from ( " & qryTaxQuery
        'strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        'strPivotForTransfer_In = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))

        ''update strPivotForTransfer_In
        strPivotForTransfer_In = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotForTransfer_In = strPivotForTransfer_In & "0 as [" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
            Else
                strPivotForTransfer_In = strPivotForTransfer_In & "0 as [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
            End If
        Next

        'Dim strPivotFortRANSFER_INPercent As String
        'strPivotFortRANSFER_INPercent = " select distinct (select  Distinct ',0 as ['+tax1+'%'+']' from ( " & qryTaxQuery
        'strPivotFortRANSFER_INPercent += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        Dim strPivotFortRANSFER_INPercentQuery As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotFortRANSFER_INPercent))

        '' update strPivotFortRANSFER_INPercentQuery
        strPivotFortRANSFER_INPercentQuery = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotFortRANSFER_INPercentQuery = strPivotFortRANSFER_INPercentQuery & "0 as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                strPivotFortRANSFER_INPercentQuery = strPivotFortRANSFER_INPercentQuery & "0 as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next

        '===========

        'Dim strPivotForGroupOuter As String
        'strPivotForGroupOuter = "select SUBSTRING(ax,2,len(Ax)) from ("        
        'strPivotForGroupOuter += " select distinct (select Distinct ',max(isnull(final.['+tax1+'%'+'],0)) as ['+TAX1+'%'+']' from ( " & qryTaxQuery

        'strPivotForGroupOuter += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"
        Dim strPivotFoGrouprOuterQuery As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForGroupOuter))
        '' update strPivotFoGrouprOuterQuery
        strPivotFoGrouprOuterQuery = ""
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotFoGrouprOuterQuery = strPivotFoGrouprOuterQuery & "max(isnull(final.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%],0)) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                strPivotFoGrouprOuterQuery = strPivotFoGrouprOuterQuery & "max(isnull(final.[" & clsCommon.myCstr(dr.Item("TAX1")) & "%],0)) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next


        'Dim strPivotForOuterForBulk As String
        'strPivotForOuterForBulk = " select distinct (select Distinct ',0 as ['+TAX1+']' from ( " & qryTaxQuery

        'strPivotForOuterForBulk += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strPivotForOuterQueryforBulk As String '' = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForOuterForBulk))
        '' update strPivotForOuterQueryforBulk
        strPivotForOuterQueryforBulk = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotForOuterQueryforBulk = strPivotForOuterQueryforBulk & "(case when '" & clsCommon.myCstr(dr.Item("TAX1")) & "' = 'TCS' THEN -TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX1_AMT ELSE 0 END) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
            Else
                strPivotForOuterQueryforBulk = strPivotForOuterQueryforBulk & "(case when '" & clsCommon.myCstr(dr.Item("TAX1")) & "' = 'TCS' THEN -TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX1_AMT ELSE 0 END) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
            End If
        Next

        'Dim strDoublePivotForOuterForBulk As String

        'strDoublePivotForOuterForBulk = " select distinct (select Distinct ',0 as ['+tax1+'%'+']' from ( " & qryTaxQuery


        'strDoublePivotForOuterForBulk += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strDoublePivotForOuterQueryforBulk As String ' = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForOuterForBulk))
        '' update strDoublePivotForOuterQueryforBulk
        strDoublePivotForOuterQueryforBulk = ","
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strDoublePivotForOuterQueryforBulk = strDoublePivotForOuterQueryforBulk & "(case when '" & clsCommon.myCstr(dr.Item("TAX1")) & "' = 'TCS' THEN TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX1_RATE ELSE 0 END) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                strDoublePivotForOuterQueryforBulk = strDoublePivotForOuterQueryforBulk & "(case when '" & clsCommon.myCstr(dr.Item("TAX1")) & "' = 'TCS' THEN TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX1_RATE ELSE 0 END) as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next

        'Dim strPivotForInner As String
        'strPivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        'strPivotForInner += " select distinct (select Distinct ',['+tax1+']' from ( " & qryTaxQuery

        'strPivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strPivotForInnerQuery As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForInner))

        '' update strPivotForInnerQuery
        strPivotForInnerQuery = ""
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotForInnerQuery = strPivotForInnerQuery & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
            Else
                strPivotForInnerQuery = strPivotForInnerQuery & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
            End If
        Next

        '' taxcolumns for no tax 
        'strPivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        'strPivotForInner += " select distinct (select Distinct ',Null as ['+tax1+']' from ( " & qryTaxQuery

        'strPivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strPivotForInnerQueryNoTax As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForInner))

        '' update strPivotForInnerQueryNoTax
        strPivotForInnerQueryNoTax = ""
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strPivotForInnerQueryNoTax = strPivotForInnerQueryNoTax & "Null as [" & clsCommon.myCstr(dr.Item("TAX1")) & "]"
            Else
                strPivotForInnerQueryNoTax = strPivotForInnerQueryNoTax & "Null as [" & clsCommon.myCstr(dr.Item("TAX1")) & "],"
            End If
        Next

        'Dim strDoublePivotForInner As String
        'strDoublePivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        'strDoublePivotForInner += " select distinct (select Distinct ',['+tax1+'%'+']' from ( " & qryTaxQuery

        'strDoublePivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strDoublePivotForInnerQuery As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForInner))

        ' update strDoublePivotForInnerQuery
        strDoublePivotForInnerQuery = ""
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strDoublePivotForInnerQuery = strDoublePivotForInnerQuery & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                strDoublePivotForInnerQuery = strDoublePivotForInnerQuery & "[" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next

        '' tax rate columns for no tax 
        'strDoublePivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        'strDoublePivotForInner += " select distinct (select Distinct ',Null as ['+tax1+'%'+']' from ( " & qryTaxQuery

        'strDoublePivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strDoublePivotForInnerQueryNoTax As String '= clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForInner))
        ' update strDoublePivotForInnerQueryNoTax
        strDoublePivotForInnerQueryNoTax = ""
        For Each dr As DataRow In dtTax.Rows
            If dtTax.Rows.IndexOf(dr) = dtTax.Rows.Count - 1 Then
                strDoublePivotForInnerQueryNoTax = strDoublePivotForInnerQueryNoTax & "Null as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%]"
            Else
                strDoublePivotForInnerQueryNoTax = strDoublePivotForInnerQueryNoTax & "Null as [" & clsCommon.myCstr(dr.Item("TAX1")) & "%],"
            End If
        Next

        Dim qryQC As String = ""
        qryQC = " select Item_Code,MAX(Fat_Per) as Fat_Per,MAX(SNF_Per) as SNF_Per from (" &
                " select Item_QCP.Item_Code,Item_QCP.Code as Parameter_Code,(case when QCP.Type='FAT' then Item_QCP.Actual_Range end) as Fat_Per," &
                " (case when QCP.Type='SNF' then Item_QCP.Actual_Range  end) as SNF_Per from TSPL_ITEM_QC_PARAMETER_MASTER Item_QCP " &
                " left join TSPL_PARAMETER_MASTER QCP  on Item_QCP.Code=QCP.Code) as QC group by Item_Code"

        Dim qryKG As String = ""
        qryKG = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='KG'"
        'qryKG = " select distinct TSPL_ITEM_UOM_DETAIL.Item_Code,coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) as UOM_Code, " & _
        '        " (case when Weigt1.Contained_UOM='KG' then round(Weigt1.Contained_Qty/Weigt1.Container_Qty,4) " & _
        '        " when Weigt2.Container_UOM='KG' then round(Weigt2.Container_Qty/Weigt2.Contained_Qty,4) else 0 end)*TSPL_ITEM_MASTER.Weight_Value as Conversion_Factor " & _
        '        " from TSPL_ITEM_UOM_DETAIL " & _
        '        " left join TSPL_ITEM_MASTER on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION where Contained_UOM='KG') Weigt1 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt1.Container_UOM  and (TSPL_ITEM_MASTER.Product_Type=Weigt1.Product_Type or Weigt1.Product_Type='All') " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION where Container_UOM='KG') Weigt2 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt2.Contained_UOM and (TSPL_ITEM_MASTER.Product_Type=Weigt2.Product_Type or Weigt2.Product_Type='All')  " & _
        '        " where  2=2 and coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) is not null and coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) in ('KG','Ltr')"
        '' done by Panch Raj from Preeti User Id
        'qryKG = " select distinct TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Weight_UOM," & _
        '        " max(case when TSPL_ITEM_MASTER.Weight_UOM='KG' then 1 " & _
        '        " else  coalesce(Weigt1.Contained_Qty/Weigt1.Container_Qty,Weigt2.Container_Qty/Weigt2.Contained_Qty)end)*TSPL_ITEM_MASTER.Weight_Value " & _
        '        " as Conversion_Factor from TSPL_ITEM_MASTER " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION ) Weigt1 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt1.Container_UOM and Weigt1.Contained_UOM='KG'  " & _
        '        " and (TSPL_ITEM_MASTER.Product_Type=Weigt1.Product_Type or Weigt1.Product_Type='All') " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION ) Weigt2 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt2.Contained_UOM and  Weigt2.Container_UOM='KG' " & _
        '        " and (TSPL_ITEM_MASTER.Product_Type=Weigt2.Product_Type or Weigt2.Product_Type='All')" & _
        '        " where  2=2 and len(coalesce(TSPL_ITEM_MASTER.Weight_UOM,''))>0 group by TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Weight_UOM,TSPL_ITEM_MASTER.Weight_Value"
        Dim qryStock As String = ""
        qryStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL "

        '' query for transaction  UOM conversion
        Dim qryTransStock As String = ""
        If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not Stock_uom Then
            qryTransStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL "
        ElseIf clsCommon.myLen(Unit_Code) <= 0 AndAlso Stock_uom Then
            qryTransStock = "select Item_Code,max(UOM_Code) as UOM_Code,max(Conversion_Factor) as Conversion_Factor from TSPL_ITEM_UOM_DETAIL where  Stocking_Unit='Y'  group by Item_Code "
        Else
            qryTransStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='" & Unit_Code & "'"
        End If
        '===================Added By Preeti Gupta===================
        Dim qryRateStock As String = ""
        ' =============================================================
        Dim CompGstinNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GSTINNo from TSPL_COMPANY_MASTER where Comp_Code='" & objCommonVar.CurrentCompanyCode & "'"))

        '' end query for transaction  UOM conversion
        ' BM00000008464 BM00000008484
        '' query for structure and item group custom field========================BM00000008352
        '==============added by preeti Gupta against ticket No[BM00000009864]
        Dim strItemGroup As String = ""
        strItemGroup = " select Struct.Structure_Code,Structure_Descq,Struct_Val.Value as Item_Group,StructDtl.Description as Group_Description from TSPL_STRUCTURE_MASTER Struct left join (" &
                       " select Custom_field_Code,Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES where Program_Code='" & clsUserMgtCode.itemStructure & "'  " &
                       " and Custom_Field_Code='" & MIS_Item_Group & "') as Struct_Val  on Struct.Structure_Code=Struct_Val.Transaction_Code" &
                       " left join (select Custom_Field_Code,SNo,Value,Description from TSPL_CUSTOM_FIELD_DETAIL where Custom_Field_Code='" & MIS_Item_Group & "') as StructDtl on Struct_Val.Value=StructDtl.Value "

        Dim strSDCommonQuery As String = ""
        Dim strTaxColumns As String = ""
        Dim strSDJoinQry As String = ""
        Dim strSDTaxRate As String = ""
        Dim strSDTaxRateColumn As String = ""
        Dim strSDTaxRateBlankColumn As String = ""
        strSDTaxRateBlankColumn = " '' as _Type ,"
        strSDTaxRateColumn = "  ttr._Type ,"
        Dim strSDEndQry As String = ""
        If obj.Trans_Type_List.Contains("Can Sale") Then
            strSDTaxRateBlankColumn = " null as _Type ,"
            strSDEndQry = ",TSPL_CANSALE_INVOICE_DETAIL.TAX1+'%' as Tax1_Rate"
        ElseIf obj.Trans_Type_List.Contains("Bulk Sale") OrElse obj.Trans_Type_List.Contains("Bulk Sale Trade") Then
            strSDTaxRateBlankColumn = " null as _Type ,"
            strSDEndQry = ",TSPL_INVOICE_DETAIL_BULKSALE.TAX1+'%' as Tax1_Rate"
        Else
            strSDEndQry = ",TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1_Rate"
        End If
        ''strSDCommonQuery = " select TSPL_SD_SALE_INVOICE_HEAD.Description as Narration,case when TSPL_SD_SALE_INVOICE_HEAD.Against_C_Form = 1 then 'C' else '' End as Formtype,case when ISNUll(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,'')<>'' then TSPL_SD_SALE_INVOICE_HEAD.Document_Type else  CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='ALL' THEN 'SD' ELSE TSPL_SD_SALE_INVOICE_HEAD.Trans_Type END end as Trans_Type  ,TSPL_SD_SALE_INVOICE_HEAD.Status ,TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location, " & _
        ''                   " TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,TSPL_SD_SALE_INVOICE_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_INVOICE_HEAD.Document_Code , " & _
        ''                   " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Line_No , " & _
        ''                   " (case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 then 0 else TSPL_SD_SALE_INVOICE_DETAIL.Qty end) as Qty, " & _
        ''                   " TSPL_SD_SALE_INVOICE_DETAIL.Unit_code , TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as Item_Cost , " & _
        ''                   " (TSPL_SD_SALE_INVOICE_DETAIL.Amount *(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as AMount, " & _
        ''                   " TSPL_SD_SALE_INVOICE_DETAIL.Disc_Per ,(TSPL_SD_SALE_INVOICE_detail.Total_Disc_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) - case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Disc_Amt, " & _
        ''                   " (case when coalesce(TSPL_SD_SALE_INVOICE_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.sampling,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item,'')='Y' then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [Scheme Amount] , " & _
        ''                   " ((Amount-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='AS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Amt_Less_Discount , " & _
        ''                   " TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Total_Tax_AMt , " & _
        ''                   " (Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Total_Amt, " & _
        ''                   " case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then ''  when ManualVehicle <> '' then '' else TSPL_SD_SALE_INVOICE_HEAD.Vehicle_Code end as Vehicle_Code , case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then TSPL_SD_SALE_INVOICE_HEAD.VehicleNo  when ManualVehicle <> '' then ManualVehicle else COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_SD_SALE_INVOICE_HEAD.VEHICLENO) end as Vehicle_No, " & _
        ''                   " (case when TSPL_SD_SALE_INVOICE_DETAIL.Line_No=1 then (TSPL_SD_SALE_INVOICE_HEAD.Total_Add_Charge+coalesce(TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount,0))*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as  Additional_Charge, " & _
        ''                   " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Document Amt]," & _
        ''                   " TSPL_Customer_Invoice_Head.Discount_Amount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Document Discount Amt], " & _
        ''                   " TSPL_Customer_Invoice_Head.amount_less_Discount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Amount Before Tax]," & _
        ''                   " TSPL_Customer_Invoice_Head.total_tax*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Total Tax], " & _
        ''                   " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Total Add Charge], " & _
        ''                   " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " & _
        ''                   " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_INVOICE_HEAD.GRNo as [GR No],tspl_sd_shipment_head.gr_date as [GR Date],TSPL_SD_SALE_INVOICE_HEAD.WayBillNo as [WayBill No],TSPL_SD_SHIPMENT_HEAD.transport_id as [Transporter Code],case when len(TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual) > 0 then TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual else TSPL_SD_SHIPMENT_HEAD.Transporter_Name end as [Transporter Name],(case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS else  TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code end) as [Delivery No]  ,Shipment_Code as [Shipment No],TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.booking_no as [Booking No], TSPL_SD_SALE_INVOICE_DETAIL.MRP ,TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Code as [Scheme Code] ,TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Code as [Cash Scheme Code] ," & _
        ''                   " TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Amount*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [Cash Scheme Amount], TSPL_SD_SALE_INVOICE_DETAIL.Price_code as [Price Code],'' as Created_By,'' as Modify_By ,TSPL_SD_SALE_INVOICE_DETAIL.RATE_UOM,TSPL_SD_SALE_INVOICE_DETAIL.Conv_Factor,tspl_sd_sale_invoice_detail.Sampling,tspl_sd_sale_invoice_detail.Scheme_Item ," & _
        ''                   " (case when TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='S' then 'Supplementry Invoice' when TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then 'Credit Note' when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=1 then 'Tax Invoice' else 'Bill Of Supply' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN no Customer], " & _
        ''                   " (case when TSPL_SD_SALE_INVOICE_HEAD.Total_Tax_Amt<=0 and TSPL_SD_SALE_INVOICE_HEAD.Tax_Group<>'EXEMPTED' then ((Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) else null end) as [Nill Rate Amount]," & _
        ''                   " (case when TSPL_SD_SALE_INVOICE_HEAD.Tax_Group='EXEMPTED' then ((Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) else null end) as [Exempted Amount]," & _
        ''                   " 0 as [Non GST Supply],'N' as [Reverse Charge],coalesce(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,'') as [Export Type],TSPL_SD_SALE_INVOICE_HEAD.Loading_Port as Port," & _
        ''                   " TSPL_SD_SALE_INVOICE_HEAD.Against_Com_Inv_No as [Shipping Bill No],convert(varchar,TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Date,103) as [Shipping Bill Date]," & _
        ''                   " TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary as [Original Invoice No],convert(varchar,SupplInvoice.Document_Date,103) as [Original Invoice Date],SupplInvoice.Description as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,"

        ''strSDEndQry = " from TSPL_SD_SALE_INVOICE_DETAIL " & _
        ''                   " left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE " & _
        ''                   " left join TSPL_VEHICLE_MASTER on TSPL_SD_SALE_INVOICE_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " & _
        ''                   " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code " & _
        ''                   " left join TSPL_DELIVERY_NOTE_MASTER_FRESHSALE on TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No=TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code " & _
        ''                   " left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code " & _
        ''                   " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code " & _
        ''                   " LEFT JOIN TSPL_EX_COMMERCIAL_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Com_Inv_No=TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Code " & _
        ''                   " LEFT JOIN TSPL_SD_SALE_INVOICE_HEAD SupplInvoice ON TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary=SupplInvoice.Document_Code " & _
        ''                   " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX1=TAXM1.TAX_CODE " & _
        ''                   " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX2=TAXM2.TAX_CODE " & _
        ''                   " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX3=TAXM3.TAX_CODE "

        ''strSDJoinQry = "  where  isnull(TSPL_SD_SALE_INVOICE_HEAD.trans_type,'ALL') not in ('CSA')  " & _
        ''                   " and (case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='FS' then 'Fresh Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='PS' then 'Product Sale' " & _
        ''                   " when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='MCC' then 'MCC Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type <>'MT' then 'Export Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type ='MT' then 'Merchant Trade' WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='SD' then 'General Sale' " & _
        ''                   " else  TSPL_SD_SALE_INVOICE_HEAD.Trans_Type end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " & _
        ''                   " and  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '' '' filter conditions
        ''If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
        ''    strSDJoinQry += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        ''End If
        ''If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
        ''    strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        ''End If

        ''If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
        ''    strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        ''End If
        ''If clsCommon.myLen(obj.Document_Code) > 0 Then
        ''    strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Document_Code = '" & obj.Document_Code & "' "
        ''End If
        ' ''===================added by preeti gupta against ticket no[BM00000009858]
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowParticluarColumnInSalesRegisterForGopalJee, clsFixedParameterCode.ShowParticluarColumnInSalesRegisterForGopalJee, Nothing)) = 1 Then
            strMCCMaterial = " select  [Trans Type],[Document_date],[Document No],[Narration],[Location Code]as [Warehouse Code], [Location Name] as [Warehouse Name],Cust.Cust_Group_Code as [Customer Group Code],Cust_Group.Cust_Group_Desc as [Customer Group Description],[Customer Code],[Customer Name],Item_Group.Item_Group as [Product Group Code],Item_Group.Group_Description as [Product Group Description],Vehicle_Code as [Vehicle Code],Vehicle_No as [Vehicle No] "

            strMCCMaterial += " , [Item Code] as [Product Code],[Item Name] as [Product Name],cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3)) as [Pack Sold]," & IIf(clsCommon.myLen(Unit_Code) <= 0, IIf(Stock_uom = True, "TransStock.UOM_Code", "xx.[UOM]"), "'" & Unit_Code & "'") & " as [Unit of sale],"
            If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not Stock_uom Then
                strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  as [Rate] "
            Else
                strMCCMaterial += "  cast(( case when isnull(Rate_Stock_SU.Conversion_Factor,0)<=0 then ([Item Cost]) else ([Item Cost] * Rate_select_SU.Conversion_Factor)/ Rate_Stock_SU.Conversion_Factor end )"
                strMCCMaterial += " as Numeric(18,3)) as [Rate]"
            End If
            strMCCMaterial += " ,Amount " + strPivotForFinalOuterQuery + " ,[Total Tax Amount], [Total Amount] as [Total Amount],'' as Remarks,''as Relation,'' as CTYPE "

        Else
            strMCCMaterial = " select "
            If obj.ReportType = "Net Sale" OrElse obj.ReportType = "Customer Wise" Then
                strMCCMaterial += " IM.Avg_Cost as [COGS], "
            End If
            'added by richa  BHA/17/08/18-000441 route
            strMCCMaterial += "  case when len([Form Type])>0 then [Form Type] else   _Type end  as [Form Type],[Trans Type],[Location Code],[Location Name],Loc.State as [Location State],GSTState.GST_STATE_Code as [GST State Code],loc.GSTNO as [Dispatch Location GSTIN No], (CASE WHEN [Invoice Type]='T' THEN 'Tax' when [Invoice Type]='R' then 'Retail' when [Invoice Type]='N' then 'None' else [Invoice Type] end) as [Invoice Type],[Document No],[Document_date],CONVERT(varchar, [Supply_Date], 103) AS Supply_Date,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                strMCCMaterial += " dcscode as [DCS Uploader Code],"
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                strMCCMaterial += " TotalSubsidyAmt,Vehicle_Code As [Vehicle Code],Vehicle_No As [Vehicle No],cast(Additional_Charge As numeric(18,2)) As [Additional Amount],dcscode as [DCS Code],[Customer Code],[Customer Name],[Customer Address],Route_Table.Route_No As Route_no,Route_Table.Route_Desc,Cust.Struct_Code,xx.[Registered],[Composition],[City Code],[Place Of Supply],[Customer GST State Code],Cust.Cust_Group_Code As [Customer Group Code],Cust_Group.Cust_Group_Desc As [Customer Group Description],Cust.cust_category_code As [Customer Category],Cust.Zone_Code As [Customer Zone Code],Zone.Description As [Customer Zone Description], [Parent Customer No],[Parent Customer Code], [Parent Customer Name],coalesce(Cust.State_Code,Cust.State_Code) As [Customer State Code],coalesce(Cust.State_Name,Cust_Loc.State_Name) As [Customer State Desc],coalesce(Cust.Tin_No,cust_loc.Tin_No) As [Tin No],Item_Group.Item_Group As [Item Group Code],Item_Group.Group_Description As [Item Group Description] "
            Else
                strMCCMaterial += " TotalSubsidyAmt as [Rate Differnce],Vehicle_Code As [Vehicle Code],Vehicle_No As [Vehicle No],cast(Additional_Charge As numeric(18,2)) As [Additional Amount],dcscode as [DCS Code],[Customer Code],[Customer Name],[Customer Address],Route_Table.Route_No As Route_no,Route_Table.Route_Desc,Cust.Struct_Code,xx.[Registered],[Composition],[City Code],[Place Of Supply],[Customer GST State Code],Cust.Cust_Group_Code As [Customer Group Code],Cust_Group.Cust_Group_Desc As [Customer Group Description],Cust.cust_category_code As [Customer Category],Cust.Zone_Code As [Customer Zone Code],Zone.Description As [Customer Zone Description], [Parent Customer No],[Parent Customer Code], [Parent Customer Name],coalesce(Cust.State_Code,Cust.State_Code) As [Customer State Code],coalesce(Cust.State_Name,Cust_Loc.State_Name) As [Customer State Desc],coalesce(Cust.Tin_No,cust_loc.Tin_No) As [Tin No],Item_Group.Item_Group As [Item Group Code],Item_Group.Group_Description As [Item Group Description] "

            End If

            'strMCCMaterial += " TotalSubsidyAmt as [Rate Differnce],Vehicle_Code As [Vehicle Code],Vehicle_No As [Vehicle No],cast(Additional_Charge As numeric(18,2)) As [Additional Amount],dcscode as [DCS Code],[Customer Code],[Customer Name],[Customer Address],Route_Table.Route_No As Route_no,Route_Table.Route_Desc,Cust.Struct_Code,xx.[Registered],[Composition],[City Code],[Place Of Supply],[Customer GST State Code],Cust.Cust_Group_Code As [Customer Group Code],Cust_Group.Cust_Group_Desc As [Customer Group Description],Cust.cust_category_code As [Customer Category],Cust.Zone_Code As [Customer Zone Code],Zone.Description As [Customer Zone Description], [Parent Customer No],[Parent Customer Code], [Parent Customer Name],coalesce(Cust.State_Code,Cust.State_Code) As [Customer State Code],coalesce(Cust.State_Name,Cust_Loc.State_Name) As [Customer State Desc],coalesce(Cust.Tin_No,cust_loc.Tin_No) As [Tin No],Item_Group.Item_Group As [Item Group Code],Item_Group.Group_Description As [Item Group Description] "
            If clsCommon.myLen(strCategoryTable) > 0 Then
                ''richa agarwal to avoid ambiguous error
                '  strMCCMaterial += "," + strCodeColumn + "," + strCodeDescColumn
                strMCCMaterial += "," + strCodeColumnForVirtual + "," + strCodeDescColumn
            End If
            ' BM00000008438 BM00000008391           
            strMCCMaterial += " , [Item Code],[Item Name],[HSN Code],"
            If obj.UOMType = "Stock UOM" OrElse obj.ReportType = "Custom Type" Then
                strMCCMaterial += "cast(([Quantity]*Stock_SU.Conversion_Factor)/(Case When coalesce(TransStock.Conversion_Factor,1)=0 Then 1 Else coalesce(TransStock.Conversion_Factor,1) End) As Numeric(18,3)) As [Quantity],"
                strMCCMaterial += " " & IIf(clsCommon.myLen(Unit_Code) <= 0, IIf(Stock_uom = True, "TransStock.UOM_Code", "xx.[UOM]"), "'" & Unit_Code & "'") & " as [UOM],"
            Else
                strMCCMaterial += " Isnull(convert(decimal(18,2),( [Quantity] * isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1)) /I.Conversion_Factor),0) as Quantity,I.UOM_Code as UOM,"
            End If

            If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not Stock_uom Then
                'sanjay
                'strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  as [Item Rate] "
                If obj.UOMType = "Stock UOM" OrElse obj.ReportType = "Custom Type" Then
                    strMCCMaterial += " case when Scheme_Item='Y' then 0 else"
                    strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3)) "
                    strMCCMaterial += " end as [Item Rate]"
                Else
                    strMCCMaterial += " case when Scheme_Item='Y' then 0 else"
                    strMCCMaterial += " cast(([Item Cost]*I.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3)) "
                    strMCCMaterial += " end as [Item Rate]"

                End If
                'strMCCMaterial += " case when Scheme_Item='Y' then 0 else"
                'strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3)) "
                'strMCCMaterial += " end as [Item Rate]"
                'sanjay
            Else
                strMCCMaterial += " case when Scheme_Item='Y' then 0 else"
                strMCCMaterial += "  cast(( case when isnull(Rate_Stock_SU.Conversion_Factor,0)<=0 then ([Item Cost]) else ([Item Cost] * Rate_select_SU.Conversion_Factor)/ Rate_Stock_SU.Conversion_Factor end )"
                strMCCMaterial += " as Numeric(18,3)) end as [Item Rate]"
            End If


            strMCCMaterial += " ,Case When [trans type] In ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [Fat Per] else Item.STD_FATPER end as [FAT %],case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [SNF Per] else Item.STD_SNFPer end as [SNF %],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [Fat Per] else Item.STD_FATPER end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as numeric(18,3)) end) as [FAT KG],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [SNF Per] else Item.STD_SNFPer end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as Numeric(18,3)) end) as [SNF KG],(case when Scheme_Item='Y' then 0 else Amount+cast(Additional_Charge as numeric(18,2)) end) as Amount,[Discount Per] as [Discount %],  (coalesce( [Discount Amount],0)-coalesce([Scheme Amount],0))  as [Discount Amount],[Scheme Amount],case when Scheme_Item='Y' then 0 else [Amount Less Discount]+cast(Additional_Charge as numeric(18,2)) end as [Amount Less Discount]" + strPivotForFinalOuterQuery + " " + strPivotForFinalOuterPercentQuery + ", " &
                    " Convert(decimal(18,2),case when Scheme_Item='Y' then 0 else case when [trans type]='Fresh Sale Return' then [Amount Less Discount]+cast(Additional_Charge as numeric(18,2))  else ([Total Amount]-[Total Tax Amount]+cast(Additional_Charge as numeric(18,2))) end end) as [Sale Amount],Convert(decimal(18,2),case when Scheme_Item='Y' then 0 else ([Total Amount]-[Total Tax Amount]+MANDI_TAX_AMT+cast(Additional_Charge as numeric(18,2))) end) as [Sale Amount GST]," &
                    " case when Scheme_Item='Y' then 0 else [Total Tax Amount] end as [Total Tax Amount], " &
                    " Convert(decimal(18,2),case when Scheme_Item='Y' then 0 else (cast(Additional_Charge as numeric(18,2))+[Total Amount]) end) as [Total Amount],"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                ' strMCCMaterial += "([Total Amount] + isnull(Security_Amt,0)+ isnull(Transporter_Commission_Amt,0)) as [Gross Amount], "
                strMCCMaterial += "Gross_Amount as [Gross Amount], "
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                strMCCMaterial += "([Total Amount] - isnull(Transporter_Commission_Amt,0)) as [Gross Amount], "
            Else
                'strMCCMaterial += " ([Total Amount] + isnull(Security_Amt,0)) as [Gross Amount], "
                strMCCMaterial += " ([Total Amount]  - isnull(TotalSubsidyAmt,0)) as [Gross Amount], "
            End If
            strMCCMaterial += " cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.HeadDiscAmt,0) as numeric(18,2))[Discount Amt],[Ack No],CONVERT(varchar, ([ACK Date]), 103)[ACK Date],[IRN No],CustomerType,PaymentType,Payment_Terms as [Payment Terms],[Narration], " &
                    " [AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax]+ case when (coalesce([Total Tax Amount],0)<>0 or [Scheme Amount]<=0) and [Document No]<>'SRFS-003/15-16/000006' then 0 else coalesce([AR Document Discount Amt],0)  end as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge], " &
                     "  case when [trans type] in ('CSA Sale','CSA Sale Return') then (case when coalesce(item.GSOC_Acct,'')<>'' then  left( item.GSOC_Acct, Len( item.GSOC_Acct)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code else '' end) else case when [trans type] in ('Dairy Sale Return','Product Sale Return', 'Fresh Sale Return') then left(Item.Sales_Return_Account , Len(Item.Sales_Return_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code else left(Item.Sales_Account, Len(Item.Sales_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code  end  end as [Sales Account], " &
                    " [GR No],convert(varchar,[GR Date],103) as [GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code],[Scheme Type] as [Schemes Type] , [Cash Scheme Code] , [Cash Scheme Amount], [Price Code], case when Sampling=0 then  'N' else case when sampling=1 then'Y' end end as sampling,"

            If obj.ReportType = "Document Detail" OrElse obj.ReportType = "Customer Wise" Then
                strMCCMaterial += " case when Sampling=1 then cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3))*case when Scheme_Item='Y' then 0 else cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  end else 0 end  as [Sampling Amount], Cust.[RSM Code] as [RSM Code] ,Cust.[RSM Name] as [RSM Name] ,Cust.[ZSM Code] as [ZSM Code] ,cust.[ZSM Name] as [ZSM Name] ,Cust.[ASM Code] as [ASM Code] ,Cust.[ASM Name] as [ASM Name] ,Cust.[ASO Code]  as [ASO Code] ,Cust.[ASO Name] as [ASO Name] , "
            End If

            strMCCMaterial += " " & IIf(Batch_Wise = True, " Batch_No, ", " ") & " Scheme_Item as [Scheme Type],[Invoice Type GST],[GSTIN No Company],[GSTIN no Customer], case when Scheme_Item='Y' then 0 else [Nill Rate Amount] end [Nill Rate Amount],cast(Additional_Charge as numeric(18,2))+ [Exempted Amount] as [Exempted Amount],[Non GST Supply],[Reverse Charge],[Export Type],Port,[Shipping Bill No],[Shipping Bill Date],[Original Invoice No],[Original Invoice Date],[Reason for Revision],[Executive],cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0) as numeric(18,2))[Commission Amt] "

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                strMCCMaterial += ",cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0) as numeric(18,2))[Security Amt],cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.Transporter_Commission_Amt,0) as numeric(18,2)) as [Transporter Amt] "
                ' strMCCMaterial += ",cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0) as numeric(18,2))[Security Amt],dcscode as [DCS Uploader Code] "
            Else
                strMCCMaterial += ",cast(IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0) as numeric(18,2))[Security Amt] "
            End If
            If obj.ReportType = "Document Detail" Then
                strMCCMaterial += ",cast(IsNull(Transporter_Commission_Amt,0) as numeric(18,2))[TPT Amt] "
            End If
            If obj.ReportType = "Customer Amount Wise" Then
                strMCCMaterial += " ,Conv_Factor,ItemConversionInLTR.Conversion_factor,
                                    (Quantity*ItemConversionInLTR.Conversion_factor) as QTYLTR"
            End If
        End If


        '' ''richa agarwal add merchant trade trans_type in below qry BM00000008390 (Applied For DCC Also) 
        strMCCMaterial += " from ( "
        '' '' base union 1
        ''If obj.Trans_Type_List.Contains("Fresh Sale") OrElse obj.Trans_Type_List.Contains("Product Sale") OrElse obj.Trans_Type_List.Contains("MCC Sale") OrElse obj.Trans_Type_List.Contains("Export Sale") OrElse obj.Trans_Type_List.Contains("CSA Sale") OrElse obj.Trans_Type_List.Contains("Merchant Trade") Then
        ''    qryStarted = True
        ''    strMCCMaterial += " select max(final._Type) as _Type , max(final.FormType) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='EX' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' when Trans_Type='MT' then 'Merchant Trade' WHEN Trans_Type ='SD' then 'General Sale' else  Trans_Type end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date], max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) AS [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

        ''    ''Monika  QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
        ''    strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " & _
        ''        " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]- case when (Trans_Type ='FS') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No], [Booking No],MRP, [Scheme Code] ,[Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Created_By,final.Modify_By ,final. RATE_UOM,final. Conv_Factor,final.Sampling,final.Scheme_Item, " & _
        ''        " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT"
        ''    strMCCMaterial += " from ("
        ''    'strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"
        ''    strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        ''    '' query for no tax applied

        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & "  and (coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax2,'')='' " & _
        ''                      " and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax4,'')='' and " & _
        ''                      " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax5,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax6,'')='' and " & _
        ''                      " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax7,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax8,'')='' and " & _
        ''                      " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax9,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax10,'')='') )t "

        ''    '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + ") " & _

        ''    strMCCMaterial += "   union all"
        ''    '' query for tax1 applied
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"

        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax1 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and TSPL_SD_SALE_INVOICE_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "   union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX2 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax2_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX2+'%' as Tax2Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax2 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX3 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax3_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX3+'%' as Tax3Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax3 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "   union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX4 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax4_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX4+'%' as Tax4Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax4 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX5 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax5_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX5+'%' as Tax5Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax5 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX6 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax6_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX6+'%' as Tax6Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax6 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX7 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax7_AMt,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX7+'%' as Tax7Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax7 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX8 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax8_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX8+'%' as Tax8Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax8 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX9 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax9_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX9+'%' as Tax9Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax9 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += "  union all"
        ''    strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX10 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax10_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX10+'%' as Tax10Rate"
        ''    strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax10 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
        ''    strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
        ''    strMCCMaterial += " )final"
        ''    strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
        ''    strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
        ''    strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
        ''    strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        ''    strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " & _
        ''                      " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


        ''    ''Monika
        ''    ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "

        ''    'added by stuti on 01/05/2017
        ''    strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        ''    ''Monika
        ''    If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
        ''        strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        ''    End If
        ''    If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
        ''        strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        ''    End If
        ''    If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
        ''        strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        ''    End If
        ''    ''end here

        ''    strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No], [Booking No],MRP , [Scheme Code] , [Cash Scheme Code] , [Cash Scheme Amount] , [Price Code] ,final.Created_By,final.Modify_By,final.RATE_UOM,final.Conv_Factor,final.Sampling,final.Scheme_Item" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per

        ''End If
        '' add csa transfer 
        If obj.Trans_Type_List.Contains("CSA Sale") Then
            Dim stCSACommonQuery As String = " select '' as Narration,TSPL_CSA_TRANSFER_HEAD.Against_Form as [Formtype],'CSA Sale' as  Trans_Type,TSPL_CSA_TRANSFER_HEAD.Status ,TSPL_CSA_TRANSFER_HEAD.From_Location_Code as Bill_To_Location," &
                                              " TSPL_CSA_TRANSFER_HEAD.Cust_Code as Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,'CSA' as Invoice_Type,TSPL_CSA_TRANSFER_HEAD.Doc_Code as Document_Code ," &
                                              " convert(varchar,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103 ) as Document_Date , TSPL_CSA_TRANSFER_DETAIL.Item_Code , " &
                                              " TSPL_CSA_TRANSFER_DETAIL.Qty ,TSPL_CSA_TRANSFER_DETAIL.Unit_code ,TSPL_CSA_TRANSFER_DETAIL.Transfer_Rate*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Item_Cost , " &
                                              " coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Amount ,coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Per,0) as Disc_Per ,coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Amt,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Disc_Amt ,0 as [Scheme Amount] " &
                                              ", (coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)-coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Amt,0))*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Amt_Less_Discount ,0 as Total_Tax_Amt ,cast((TSPL_CSA_TRANSFER_DETAIL.Transfer_Rate * TSPL_CSA_TRANSFER_DETAIL.Qty) as Numeric(18,2))*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Total_Amt, " &
                                              " TSPL_CSA_TRANSFER_HEAD.Vehicle_Id as Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,0 as Additional_Charge, " &
                                              " '' as [AR Document No],0 as [AR Document Amt],0 as [AR Document Discount Amt], 0 as [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_CSA_TRANSFER_HEAD.GR_No as [GR No],TSPL_CSA_TRANSFER_HEAD.gr_date as [GR Date],TSPL_CSA_TRANSFER_HEAD.WayBill_No as [WayBill No],TSPL_CSA_TRANSFER_HEAD.Transport_Id as [Transporter Code] , case when len(TSPL_CSA_TRANSFER_HEAD.transporter_name_manual) > 0 then TSPL_CSA_TRANSFER_HEAD.transporter_name_manual else TSPL_TRANSPORT_MASTER.Transporter_Name end as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],TSPL_CSA_TRANSFER_DETAIL.MRP,'' as  [Scheme Code],'' as [Scheme Type] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_CSA_TRANSFER_HEAD.Created_By ,TSPL_CSA_TRANSFER_HEAD.Modify_By,TSPL_CSA_TRANSFER_DETAIL.Unit_code as RATE_UOM,0 as Conv_Factor, " &
                                              " 'Delivery Challan' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer],(case when TSPL_CSA_TRANSFER_HEAD.Total_Tax_Amt<=0 and TSPL_CSA_TRANSFER_HEAD.Tax_Group<>'EXEMPTED' then coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) else 0 end) as [Nill Rate Amount],(case when TSPL_CSA_TRANSFER_HEAD.Tax_Group='EXEMPTED' then coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) else 0 end) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_CSA_TRANSFER_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_CSA_TRANSFER_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_CSA_TRANSFER_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT ,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive],"

            strSDEndQry = " from TSPL_CSA_TRANSFER_DETAIL " &
                        " left outer join TSPL_CSA_TRANSFER_HEAD on TSPL_CSA_TRANSFER_HEAD.DOC_CODE =TSPL_CSA_TRANSFER_DETAIL.DOC_CODE " &
                        " left join TSPL_VEHICLE_MASTER on TSPL_CSA_TRANSFER_HEAD.Vehicle_Id=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                        " left join TSPL_TRANSPORT_MASTER on TSPL_CSA_TRANSFER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                        " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_CSA_TRANSFER_HEAD.Cust_Code " &
                        " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
                        " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_CSA_TRANSFER_DETAIL.TAX1=TAXM1.TAX_CODE " &
                        " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_CSA_TRANSFER_DETAIL.TAX2=TAXM2.TAX_CODE " &
                        " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_CSA_TRANSFER_DETAIL.TAX3=TAXM3.TAX_CODE "

            strSDJoinQry = " where 'CSA Sale' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                        " and convert(date,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,('" & To_Date & "'),103) "
            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CSA_TRANSFER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CSA_TRANSFER_HEAD.From_Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CSA_TRANSFER_HEAD.Cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_CSA_TRANSFER_HEAD.Doc_Code = '" & obj.Document_Code & "' "
            End If
            If qryStarted = True Then
                strMCCMaterial += " union all " '' base union 2
            End If
            qryStarted = True
            strMCCMaterial += " select max(final._Type) as _Type, max(final.FormType) as [Form Type],Trans_Type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " &
                              " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP, [Scheme Code],[Scheme Type] , [Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final. RATE_UOM,final. Conv_Factor,0 as Sampling,'N' as Scheme_Item, " &
                              " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive]"
            strMCCMaterial += " from ("
            'strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX1 ,0 as TAX1_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX1_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX1+'%' as Tax1Rate"
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            '' query for no tax applied& strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry &

            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & "  and (coalesce(TSPL_CSA_TRANSFER_DETAIL.tax1,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax3,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax5,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax7,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax9,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax10,'')='') )t  "
            '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            '' query for tax1 applied
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX1 ,0 as TAX1_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX1_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX1+'%' as Tax1Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax1 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T' "
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and TSPL_CSA_TRANSFER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX2 ,0 as TAX2_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX2_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX2+'%' as Tax2Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax2 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX3 ,0 as TAX3_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX3_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX3+'%' as Tax3Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax3 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX4 ,0 as TAX4_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX4_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX4+'%' as Tax4Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax4 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX5 ,0 as TAX5_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX5_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX5+'%' as Tax5Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax5 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX6 ,0 as TAX6_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX6_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX6+'%' as Tax6Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax6 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX7 ,0 as TAX7_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX7_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX7+'%' as Tax7Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax7 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX8 ,0 as TAX8_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX8_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX8+'%' as Tax8Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax8 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX9 ,0 as TAX9_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX9_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX9+'%' as Tax9Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax9 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX10 ,0 as TAX10_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX10_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX10+'%' as Tax10Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CSA_TRANSFER_DETAIL.tax10 and ttr.tax_Rate=TSPL_CSA_TRANSFER_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & stCSACommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_CSA_TRANSFER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " )final"
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
            strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
            strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
            strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
            strMCCMaterial += " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code"
            strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                              " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "



            ''Monika
            ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strMCCMaterial += " group by  final.Trans_Type,final.Status  ,final.Document_Code ,final.Item_Code ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Scheme Type], [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final. RATE_UOM,final.Conv_Factor,[Executive]" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per
        End If
        '' end of csa transfer

        '' base union 3
        'If obj.Trans_Type_List.Contains("Bulk Sale") OrElse obj.Trans_Type_List.Contains("Bulk Sale Trade") Then
        '    If qryStarted = True Then
        '        strMCCMaterial += " union all"
        '    End If
        '    qryStarted = True
        '    strMCCMaterial += " select '' as _Type,'' as [Form Type], case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end as [Trans Type] ,coalesce(TSPL_LOCATION_MASTER.Main_Location_Code,TSPL_LOCATION_MASTER.Location_Code) as [Location Code],TSPL_INVOICE_Master_BULKSALE.Posted as Status,coalesce(Main_Loc.Location_Desc,TSPL_LOCATION_MASTER.location_desc) as [Location Name] ,'Invoice' as [Invoice Type] ,TSPL_INVOICE_Master_BULKSALE.Document_No as [Document No] ,convert(varchar,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) as [Document_date],'' as [Narration],case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end as Vehicle_Code,case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end as Vehicle_No,(case when ROW_NUMBER() over (partition by TSPL_INVOICE_Master_BULKSALE.Document_No order by TSPL_INVOICE_DETAIL_BULKSALE.Item_Code )=1 then coalesce(TSPL_INVOICE_Master_BULKSALE.RoundOffAmount,0) else 0 end) as Additional_Charge, " & _
        '                      " TSPL_INVOICE_Master_BULKSALE.Customer_Code as [Customer Code],TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As [Customer Address] ,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER .GST_Registered as [Registered],TSPL_CUSTOMER_MASTER .GST_COMPOSITION as [Composition],TSPL_CUSTOMER_MASTER .City_Code as [City Code],TSPL_CITY_MASTER .City_Name as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,TSPL_CUSTOMER_MASTER.Parent_Customer_No as [Parent Customer No],"
        '    strMCCMaterial += " Parent_Master.Cust_Code as [Parent Customer Code],Parent_Master.Customer_Name as [Parent Customer Name], "
        '    strMCCMaterial += " TSPL_INVOICE_DETAIL_BULKSALE.Item_Code as [Item Code],tspl_item_master.Item_Desc as [Item Name],tspl_item_master.HSN_Code as [HSN Code] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceQty as [Quantity] ,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as [UOM],TSPL_INVOICE_DETAIL_BULKSALE.InvoiceRate as [Item Cost],TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatPer as [Fat Per] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFPer as [SNF Per] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatKG as [Fat Kg] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFKG as [SNF KG] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as Amount,0 as [Discount Per],0 as [Discount Amount],0 as [Scheme Amount],TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as [Amount Less Discount],0 as [Total Tax Amount] " + strPivotForOuterQueryforBulk + " " + strDoublePivotForOuterQueryforBulk + ",(TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount) as [Total Amount], " & _
        '                      " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," & _
        '                      " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " & _
        '                      " (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " & _
        '                      " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],'' as  [Delivery No]  ,'' as  [Shipment No],'' as [Booking No],0 AS MRP,'' as  [Scheme Code] ,'' as [Scheme Type],'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_INVOICE_Master_BULKSALE.Created_By as Created_By ,TSPL_INVOICE_Master_BULKSALE.Modified_By as Modify_By,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as RATE_UOM,0 as Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " & _
        '                      " 'Bill of Supply' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer],(TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT ,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive]" & _
        '                      " from TSPL_INVOICE_DETAIL_BULKSALE "

        '    strMCCMaterial += " left outer join TSPL_INVOICE_Master_BULKSALE on TSPL_INVOICE_Master_BULKSALE.Document_No =TSPL_INVOICE_DETAIL_BULKSALE.Document_No "
        '    strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_INVOICE_Master_BULKSALE.Location_Code"
        '    strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as Main_Loc on TSPL_LOCATION_MASTER.Main_Location_Code =Main_Loc.Location_Code"
        '    strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =TSPL_INVOICE_Master_BULKSALE.Customer_Code"
        '    strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No"
        '    strMCCMaterial += " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code"
        '    strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " & _
        '                      " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "
        '    strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_INVOICE_DETAIL_BULKSALE.Item_Code"
        '    strMCCMaterial += " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_INVOICE_Master_BULKSALE.Document_No  where 2=2 " & _
        '                        " and (case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " & _
        '                        " and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '    ''====Monika
        '    If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
        '        strMCCMaterial += " and TSPL_INVOICE_Master_BULKSALE.Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        '    End If
        '    If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
        '        strMCCMaterial += " and TSPL_INVOICE_DETAIL_BULKSALE.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        '    End If
        '    If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
        '        strMCCMaterial += " and TSPL_INVOICE_Master_BULKSALE.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        '    End If
        '    If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
        '        strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
        '    End If
        '    ''=======end here
        'End If

        If obj.Trans_Type_List.Contains("Bulk Sale") OrElse obj.Trans_Type_List.Contains("Bulk Sale Trade") Then
            strMCCMaterial += Environment.NewLine + "-- Bulk Sale/Trade Start" + Environment.NewLine
            If qryStarted = True Then
                strMCCMaterial += " union all"
            End If
            qryStarted = True
            strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode, max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No])[Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type,max(final.Formtype) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='Exp' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' WHEN (Trans_Type ='SD' or Trans_Type ='ALL') then 'General Sale' else Trans_Type  end  as [Trans Type],final.Loc_Code  as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.shipment_No  as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],'' as Vehicle_Code,'' as Vehicle_No,0 as Additional_Charge,final.cust_Code  as [Customer Code],MAX(final.CustAdd ) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.shipped_Qty  as [Quantity],final.Unit_code as [UOM],final.price  as [Item Cost], "

            strMCCMaterial += " max([Fat Per]) as [Fat Per],max([SNF Per]) as [SNF Per],max([Fat Kg]) as [Fat Kg],max([snf Kg]) as [SNF KG],final.ItemAmt as Amount ,final.DiscountPer  as [Discount Per],final.TotalDiscountAmt   as [Discount Amount],final.[Scheme Amount]   as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,isnull(final.TotalTaxAmt,0)  as [Total Tax Amount],final.Doc_Amt  as [Total Amount], " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Scheme Type] ,[Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final. RATE_UOM,final.Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive] " &
                " " & If(clsCommon.CompairString(obj.Program_Code, clsUserMgtCode.RptBulkSaleRegister) = CompairStringResult.Equal, ", sum(isnull(Fat_Amt,0)) as Fat_Amt,sum(isnull(SNF_Amt,0)) as SNF_Amt,sum(isnull(Standard_Rate ,0)) as Standard_Rate", "") & " " &
                " from ("

            Dim strScarpCommonQry As String = ""

            strScarpCommonQry += " select '' as Formtype,   case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end  as Trans_Type ,coalesce(TSPL_LOCATION_MASTER.Main_Location_Code,TSPL_LOCATION_MASTER.Location_Code) as Loc_Code,TSPL_INVOICE_Master_BULKSALE.Posted as Status,coalesce(Main_Loc.Location_Desc,TSPL_LOCATION_MASTER.location_desc) as [Location Name] ,'Invoice' as Invoice_Type ,TSPL_INVOICE_Master_BULKSALE.Document_No as shipment_No ,convert(varchar,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) as [Document_date],'' as [Narration], case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end as Vehicle_Code,case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end  as Vehicle_No,(case when ROW_NUMBER() over (partition by TSPL_INVOICE_Master_BULKSALE.Document_No order by TSPL_INVOICE_DETAIL_BULKSALE.Item_Code )=1 then coalesce(TSPL_INVOICE_Master_BULKSALE.RoundOffAmount,0) else 0 end) as Additional_Charge," &
    " TSPL_INVOICE_Master_BULKSALE.Customer_Code as cust_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd ,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER .GST_Registered as [Registered],TSPL_CUSTOMER_MASTER .GST_COMPOSITION as [Composition],TSPL_CUSTOMER_MASTER .City_Code as [City Code],TSPL_CITY_MASTER .City_Name as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,TSPL_CUSTOMER_MASTER.Parent_Customer_No as [Parent Customer No]," &
    " Parent_Master.Cust_Code as [Parent Customer Code],Parent_Master.Customer_Name as [Parent Customer Name], " &
    " TSPL_INVOICE_DETAIL_BULKSALE.Item_Code as Item_Code,tspl_item_master.Item_Desc as [Item Name],tspl_item_master.HSN_Code as [HSN Code] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceQty as shipped_Qty ,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as Unit_code,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceRate as Price,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatPer as [Fat Per] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFPer as [SNF Per] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatKG as [Fat Kg] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFKG as [SNF KG] ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as ItemAmt,0 as DiscountPer,0 as TotalDiscountAmt,0 as [Scheme Amount],
    TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as Amt_less_Discount,TSPL_INVOICE_DETAIL_BULKSALE.Total_Tax_Amt as TotalTaxAmt,(case when TSPL_INVOICE_Master_BULKSALE.InvoiceAgainst='Against Dispatch Trade' then TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount else TSPL_INVOICE_DETAIL_BULKSALE.Item_Net_Amt end) as Doc_Amt, " &
    " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
    " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
    " (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
    " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No],
    TSPL_INVOICE_Master_BULKSALE.EWayBillNo as [EWayBill No],convert(Varchar,TSPL_INVOICE_Master_BULKSALE.EWayBillDate,103) as [EWayBill Date],'' as [Transporter Code],'' as [Transporter Name],'' as  [Delivery No]  ,'' as  [Shipment No],0 AS MRP,'' as  [Scheme Code] ,'' as [Scheme Type],'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_INVOICE_Master_BULKSALE.Created_By as Created_By ,TSPL_INVOICE_Master_BULKSALE.Modified_By as Modify_By,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as RATE_UOM,0 as Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " &
    " 'Bill of Supply' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer],(TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],'' as [LUT No],0 AS MANDI_TAX_AMT,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive], " &
    " " & If(clsCommon.CompairString(obj.Program_Code, clsUserMgtCode.RptBulkSaleRegister) = CompairStringResult.Equal, " (case when TSPL_INVOICE_Master_BULKSALE.InvoiceAgainst='Against Dispatch Trade' then TSPL_Dispatch_Detail_BulkSale_Trade.FatAmount " &
    " else  TSPL_Dispatch_Detail_BulkSale.FatAmount end) as Fat_Amt, " &
    " (case when TSPL_INVOICE_Master_BULKSALE.InvoiceAgainst='Against Dispatch Trade' then TSPL_Dispatch_Detail_BulkSale_Trade.SNFAmount " &
    " else  TSPL_Dispatch_Detail_BulkSale.SNFAmount end) as SNF_Amt, " &
    " (case when TSPL_INVOICE_Master_BULKSALE.InvoiceAgainst='Against Dispatch Trade' then TSPL_Dispatch_Detail_BulkSale_Trade.StandardRate " &
    " else  TSPL_Dispatch_Detail_BulkSale.StandardRate end) as Standard_Rate,", "")

            strSDEndQry = " from TSPL_INVOICE_DETAIL_BULKSALE  " + Environment.NewLine +
" left outer join TSPL_INVOICE_Master_BULKSALE on TSPL_INVOICE_Master_BULKSALE.Document_No =TSPL_INVOICE_DETAIL_BULKSALE.Document_No "

            If clsCommon.CompairString(obj.Program_Code, clsUserMgtCode.RptBulkSaleRegister) = CompairStringResult.Equal Then
                strSDEndQry += " left join TSPL_Dispatch_Detail_BulkSale on TSPL_INVOICE_DETAIL_BULKSALE.Dispatch_Code=TSPL_Dispatch_Detail_BulkSale.Document_No and TSPL_INVOICE_DETAIL_BULKSALE.Item_Code=TSPL_Dispatch_Detail_BulkSale.Item_Code " + Environment.NewLine +
                 " left join TSPL_Dispatch_Detail_BulkSale_Trade on TSPL_INVOICE_DETAIL_BULKSALE.Dispatch_Code=TSPL_Dispatch_Detail_BulkSale_Trade.Document_No and TSPL_INVOICE_DETAIL_BULKSALE.Item_Code=TSPL_Dispatch_Detail_BulkSale_Trade.Item_Code "
            End If

            strSDEndQry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_INVOICE_Master_BULKSALE.Location_Code" + Environment.NewLine +
    " left outer join TSPL_LOCATION_MASTER as Main_Loc on TSPL_LOCATION_MASTER.Main_Location_Code =Main_Loc.Location_Code" + Environment.NewLine +
    " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =TSPL_INVOICE_Master_BULKSALE.Customer_Code" + Environment.NewLine +
    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No" + Environment.NewLine +
    " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code " + Environment.NewLine +
            " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " + Environment.NewLine +
    " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE " + Environment.NewLine +
    " left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_INVOICE_DETAIL_BULKSALE.Item_Code" + Environment.NewLine +
    " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_INVOICE_Master_BULKSALE.Document_No  "

            strSDJoinQry = " where 2=2 " + Environment.NewLine +
                                " and (case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                                " and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_INVOICE_Master_BULKSALE.Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_INVOICE_DETAIL_BULKSALE.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_INVOICE_Master_BULKSALE.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax

            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_INVOICE_DETAIL_BULKSALE.tax1,'')='' and coalesce(TSPL_INVOICE_DETAIL_BULKSALE.tax2,'')='' " &
                                  " and coalesce(TSPL_INVOICE_DETAIL_BULKSALE.tax3,'')='' and coalesce(TSPL_INVOICE_DETAIL_BULKSALE.tax4,'')='') )t "

            strMCCMaterial += " union all "

            '' query for tax1 applied
            strTaxColumns = " TSPL_INVOICE_DETAIL_BULKSALE.TAX1 ,TSPL_INVOICE_DETAIL_BULKSALE.TAX1_Amt ,TSPL_INVOICE_DETAIL_BULKSALE.TAX1_Rate, TSPL_INVOICE_DETAIL_BULKSALE.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_INVOICE_DETAIL_BULKSALE.tax1 and ttr.tax_Rate=TSPL_INVOICE_DETAIL_BULKSALE.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_INVOICE_DETAIL_BULKSALE.tax1<>'')"
            strMCCMaterial += " s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_INVOICE_DETAIL_BULKSALE.TAX2 ,TSPL_INVOICE_DETAIL_BULKSALE.TAX2_Amt ,TSPL_INVOICE_DETAIL_BULKSALE.TAX2_Rate, TSPL_INVOICE_DETAIL_BULKSALE.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_INVOICE_DETAIL_BULKSALE.tax2 and ttr.tax_Rate=TSPL_INVOICE_DETAIL_BULKSALE.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_INVOICE_DETAIL_BULKSALE.tax2<>'' )"
            strMCCMaterial += " s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = "TSPL_INVOICE_DETAIL_BULKSALE.TAX3 ,TSPL_INVOICE_DETAIL_BULKSALE.TAX3_Amt , TSPL_INVOICE_DETAIL_BULKSALE.TAX3_Rate, TSPL_INVOICE_DETAIL_BULKSALE.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_INVOICE_DETAIL_BULKSALE.tax3 and ttr.tax_Rate=TSPL_INVOICE_DETAIL_BULKSALE.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += "  select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_INVOICE_DETAIL_BULKSALE.tax3<>'' )"
            strMCCMaterial += " s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_INVOICE_DETAIL_BULKSALE.TAX4 ,TSPL_INVOICE_DETAIL_BULKSALE.TAX4_Amt ,TSPL_INVOICE_DETAIL_BULKSALE.TAX4_Rate, TSPL_INVOICE_DETAIL_BULKSALE.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_INVOICE_DETAIL_BULKSALE.tax4 and ttr.tax_Rate=TSPL_INVOICE_DETAIL_BULKSALE.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_INVOICE_DETAIL_BULKSALE.tax4<>'' )"
            strMCCMaterial += " s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " ) final left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Loc_Code  " &
                    " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code  " &
                    " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.cust_Code   " &
                    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No " &
                    " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                    " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.shipment_No  ,final.Item_Code ,final.Loc_Code  ,final.cust_Code  ,final.shipped_Qty  ,final.TotalTaxAmt  ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.price  ,final.ItemAmt  ,final.DiscountPer,final.TotalDiscountAmt  ,final.Amt_less_Discount   ,final.Amt_Less_Discount ,final.Doc_Amt,final.[Scheme Amount],Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP ,  [Scheme Code],[Scheme Type] , [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final.RATE_UOM,final. Conv_Factor "
            strMCCMaterial += Environment.NewLine + "-- End of Bulk Sale Start" + Environment.NewLine
        End If




        '--Can Sale By Balwinder on 12/06/2019 against ticket No GKD/12/06/19-000180 For Temp Report
        '' base union 3

        If obj.Trans_Type_List.Contains("Can Sale") Then
            strMCCMaterial += Environment.NewLine + "-- Can Sale Start" + Environment.NewLine
            If qryStarted = True Then
                strMCCMaterial += " union all"
            End If
            qryStarted = True
            strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt, '' as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No])[Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type,max(final.Formtype) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='Exp' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' WHEN (Trans_Type ='SD' or Trans_Type ='ALL') then 'General Sale' else Trans_Type  end  as [Trans Type],final.Loc_Code  as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.shipment_No  as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],'' as Vehicle_Code,'' as Vehicle_No,0 as Additional_Charge,final.cust_Code  as [Customer Code],MAX(final.CustAdd ) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.shipped_Qty  as [Quantity],final.Unit_code as [UOM],final.price  as [Item Cost], "

            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.ItemAmt as Amount ,final.DiscountPer  as [Discount Per],final.TotalDiscountAmt   as [Discount Amount],final.[Scheme Amount]   as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,isnull(final.TotalTaxAmt,0)  as [Total Tax Amount],final.Doc_Amt  as [Total Amount], " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Scheme Type] ,[Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final. RATE_UOM,final.Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive]  " &
                " from ("

            '' richa UDL/09/08/18-000213 do job work data separately
            Dim strScarpCommonQry As String = ""

            strScarpCommonQry += " select '' as Formtype,   'Can Sale'  as Trans_Type ,coalesce(TSPL_LOCATION_MASTER.Main_Location_Code,TSPL_LOCATION_MASTER.Location_Code) as Loc_Code,TSPL_CANSALE_INVOICE_HEAD.Posted as Status,coalesce(Main_Loc.Location_Desc,TSPL_LOCATION_MASTER.location_desc) as [Location Name] ,'Invoice' as Invoice_Type ,TSPL_CANSALE_INVOICE_HEAD.Document_No as shipment_No ,convert(varchar,TSPL_CANSALE_INVOICE_HEAD.Document_Date,103) as [Document_date],'' as [Narration], '' as Vehicle_Code,''  as Vehicle_No,(case when ROW_NUMBER() over (partition by TSPL_CANSALE_INVOICE_HEAD.Document_No order by TSPL_CANSALE_INVOICE_DETAIL.ItemCode )=1 then coalesce(TSPL_CANSALE_INVOICE_HEAD.RoundOffAmount,0) else 0 end) as Additional_Charge," &
                              " TSPL_CANSALE_INVOICE_HEAD.Customer_Code as cust_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd ,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER .GST_Registered as [Registered],TSPL_CUSTOMER_MASTER .GST_COMPOSITION as [Composition],TSPL_CUSTOMER_MASTER .City_Code as [City Code],TSPL_CITY_MASTER .City_Name as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,TSPL_CUSTOMER_MASTER.Parent_Customer_No as [Parent Customer No]," &
            " Parent_Master.Cust_Code as [Parent Customer Code],Parent_Master.Customer_Name as [Parent Customer Name], " &
            " TSPL_CANSALE_INVOICE_DETAIL.ItemCode as Item_Code,tspl_item_master.Item_Desc as [Item Name],tspl_item_master.HSN_Code as [HSN Code] ,TSPL_CANSALE_INVOICE_DETAIL.Qty as shipped_Qty ,TSPL_CANSALE_INVOICE_DETAIL.UOM as Unit_code,TSPL_CANSALE_INVOICE_DETAIL.MilkRate as Price,TSPL_CANSALE_INVOICE_DETAIL.FatPer as [Fat Per] ,TSPL_CANSALE_INVOICE_DETAIL.SNFPer as [SNF Per] ,TSPL_CANSALE_INVOICE_DETAIL.Fat_KG as [Fat Kg] ,TSPL_CANSALE_INVOICE_DETAIL.SNF_KG as [SNF KG] ,TSPL_CANSALE_INVOICE_DETAIL.MilkAmt as ItemAmt,0 as DiscountPer,0 as TotalDiscountAmt,0 as [Scheme Amount],TSPL_CANSALE_INVOICE_DETAIL.MilkAmt as Amt_less_Discount,TSPL_CANSALE_INVOICE_DETAIL.Total_Tax_Amt as TotalTaxAmt,TSPL_CANSALE_INVOICE_DETAIL.Item_Net_Amt as Doc_Amt, " &
                              " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                              " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                              " (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                              " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No],'' as [EWayBill No],'' as [EWayBill Date],'' as [Transporter Code],'' as [Transporter Name],'' as  [Delivery No]  ,'' as  [Shipment No],0 AS MRP,'' as  [Scheme Code] ,'' as [Scheme Type],'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_CANSALE_INVOICE_HEAD.Created_By as Created_By ,TSPL_CANSALE_INVOICE_HEAD.Modified_By as Modify_By,TSPL_CANSALE_INVOICE_DETAIL.UOM as RATE_UOM,0 as Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " &
                              " 'Bill of Supply' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer],(TSPL_CANSALE_INVOICE_DETAIL.TotalAmount) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],'' as [LUT No],0 AS MANDI_TAX_AMT,'' as [Executive], " &
                              " " & If(clsCommon.CompairString(obj.Program_Code, clsUserMgtCode.FrmCanSale) = CompairStringResult.Equal, " TSPL_CANSALE_INVOICE_DETAIL.FatAmount as Fat_Amt, " &
                              " TSPL_CANSALE_INVOICE_DETAIL.SNFAmount as SNF_Amt, TSPL_CANSALE_INVOICE_HEAD.Standard_Rate as Standard_Rate", "") + Environment.NewLine

            strSDEndQry = " from TSPL_CANSALE_INVOICE_DETAIL  " + Environment.NewLine +
                                "left outer join TSPL_CANSALE_INVOICE_HEAD on TSPL_CANSALE_INVOICE_HEAD.Document_No =TSPL_CANSALE_INVOICE_DETAIL.Document_No  " + Environment.NewLine +
                                "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_CANSALE_INVOICE_HEAD.Location_Code " + Environment.NewLine +
                                "left outer join TSPL_LOCATION_MASTER as Main_Loc on TSPL_LOCATION_MASTER.Main_Location_Code =Main_Loc.Location_Code " + Environment.NewLine +
                                "left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_CANSALE_INVOICE_HEAD.Customer_Code " + Environment.NewLine +
                                "LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No " + Environment.NewLine +
                                "left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code  " + Environment.NewLine +
                                "LEFT outer JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE  " + Environment.NewLine +
                                "left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_CANSALE_INVOICE_DETAIL.ItemCode " + Environment.NewLine +
                                "left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_CANSALE_INVOICE_HEAD.Document_No" + Environment.NewLine

            strSDJoinQry = " where 2=2 " + Environment.NewLine +
                                " and ('Can Sale') in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                                " and convert(date,TSPL_CANSALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_CANSALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CANSALE_INVOICE_HEAD.Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CANSALE_INVOICE_DETAIL.ItemCode in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CANSALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax

            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_CANSALE_INVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_CANSALE_INVOICE_DETAIL.tax2,'')='' " &
                                  " and coalesce(TSPL_CANSALE_INVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_CANSALE_INVOICE_DETAIL.tax4,'')='') )t "

            strMCCMaterial += " union all "
            '' query for tax1 applied
            strTaxColumns = " TSPL_CANSALE_INVOICE_DETAIL.TAX1 ,TSPL_CANSALE_INVOICE_DETAIL.TAX1_Amt ,TSPL_CANSALE_INVOICE_DETAIL.TAX1_Rate, TSPL_CANSALE_INVOICE_DETAIL.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CANSALE_INVOICE_DETAIL.tax1 and ttr.tax_Rate=TSPL_CANSALE_INVOICE_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_CANSALE_INVOICE_DETAIL.tax1<>'')"
            strMCCMaterial += " s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_CANSALE_INVOICE_DETAIL.TAX2 ,TSPL_CANSALE_INVOICE_DETAIL.TAX2_Amt ,TSPL_CANSALE_INVOICE_DETAIL.TAX2_Rate, TSPL_CANSALE_INVOICE_DETAIL.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CANSALE_INVOICE_DETAIL.tax2 and ttr.tax_Rate=TSPL_CANSALE_INVOICE_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_CANSALE_INVOICE_DETAIL.tax2<>'' )"
            strMCCMaterial += " s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = "TSPL_CANSALE_INVOICE_DETAIL.TAX3 ,TSPL_CANSALE_INVOICE_DETAIL.TAX3_Amt , TSPL_CANSALE_INVOICE_DETAIL.TAX3_Rate, TSPL_CANSALE_INVOICE_DETAIL.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CANSALE_INVOICE_DETAIL.tax3 and ttr.tax_Rate=TSPL_CANSALE_INVOICE_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += "  select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_CANSALE_INVOICE_DETAIL.tax3<>'' )"
            strMCCMaterial += " s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_CANSALE_INVOICE_DETAIL.TAX4 ,TSPL_CANSALE_INVOICE_DETAIL.TAX4_Amt ,TSPL_CANSALE_INVOICE_DETAIL.TAX4_Rate, TSPL_CANSALE_INVOICE_DETAIL.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_CANSALE_INVOICE_DETAIL.tax4 and ttr.tax_Rate=TSPL_CANSALE_INVOICE_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_CANSALE_INVOICE_DETAIL.tax4<>'' )"
            strMCCMaterial += " s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"


            strMCCMaterial += " ) final left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Loc_Code  " &
                    " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code  " &
                    " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.cust_Code   " &
                    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No " &
                    " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                    " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.shipment_No  ,final.Item_Code ,final.Loc_Code  ,final.cust_Code  ,final.shipped_Qty  ,final.TotalTaxAmt  ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.price  ,final.ItemAmt  ,final.DiscountPer,final.TotalDiscountAmt  ,final.Amt_less_Discount   ,final.Amt_Less_Discount ,final.Doc_Amt,final.[Scheme Amount],Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP ,  [Scheme Code],[Scheme Type] , [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final.RATE_UOM,final. Conv_Factor "
            strMCCMaterial += Environment.NewLine + "-- End of Can Sale Start" + Environment.NewLine
        End If

        '' base union 4
        If obj.Trans_Type_List.Contains("Misc Sale") Then
            If qryStarted = True Then
                strMCCMaterial += " union all"
            End If
            qryStarted = True
            strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt, '' as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No])[Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type,max(final.Formtype) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='Exp' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' WHEN Trans_Type ='SD' then 'General Sale' else Trans_Type  end  as [Trans Type],final.Loc_Code  as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.shipment_No  as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],'' as Vehicle_Code,'' as Vehicle_No,0 as Additional_Charge,final.cust_Code  as [Customer Code],MAX(final.CustAdd ) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.shipped_Qty  as [Quantity],final.Unit_code as [UOM],final.price  as [Item Cost], "

            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.ItemAmt as Amount ,final.DiscountPer  as [Discount Per],final.TotalDiscountAmt   as [Discount Amount],final.[Scheme Amount]   as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.TotalTaxAmt  as [Total Tax Amount],final.Doc_Amt  as [Total Amount], " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code],[Scheme Type] ,[Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final. RATE_UOM,final.Conv_Factor ,0 as Sampling,'N' as Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive] " &
                " from ("

            Dim strScarpCommonQry As String = ""
            strScarpCommonQry = " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,TSPL_SCRAPINVOICE_HEAD.Ack_No as [Ack No],TSPL_SCRAPINVOICE_HEAD.Ack_Date as [ACK Date],TSPL_SCRAPINVOICE_HEAD.IRN_No as [IRN No],TSPL_SCRAPINVOICE_HEAD.Description as Narration,'' as Formtype,'SS' as Trans_Type,TSPL_SCRAPINVOICE_HEAD.ispost as Status ,TSPL_SCRAPINVOICE_HEAD.Loc_Code,TSPL_SCRAPINVOICE_HEAD.cust_Code " &
                                " ,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd, " &
                                " case when isnull(Is_CashSale,'N')='Y' then 'Cash Sale Invoice' when isnull(Is_Scrap,'N')='Y' then 'Scrap Sale Invoice' else 'Misc Sale Invoice' end as Invoice_Type,TSPL_SCRAPINVOICE_HEAD.invoice_No as shipment_No ,convert(varchar,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103 ) as Document_Date , " &
                                " TSPL_SCRAPINVOICE_DETAIL.Item_Code ,TSPL_SCRAPINVOICE_DETAIL.shipped_Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code ," &
                                " TSPL_SCRAPINVOICE_DETAIL.price ,0 as InvoiceFatPer ,0 as InvoiceSNFPer ,0 as InvoiceFatKG ,0 as InvoiceSNFKG , " &
                                " TSPL_SCRAPINVOICE_DETAIL.ItemAmt ,TSPL_SCRAPINVOICE_DETAIL.DiscountPer,0 as [Scheme Amount] ,TSPL_SCRAPINVOICE_DETAIL.ItemNetAmt as Amt_less_Discount, " &
                                " TSPL_SCRAPINVOICE_DETAIL.TotalDiscountAmt , " &
                                " TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt ,TSPL_SCRAPINVOICE_DETAIL.TotalAmt+Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end as Doc_Amt,'' as Vehicle_Code,'' as Vehicle_No,Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end as Additional_Charge," &
                                " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                                " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                                " TSPL_Customer_Invoice_Head.amount_less_Discount as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                                " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge], " &
                                " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                                " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,'' as [GR No],NULL as [GR Date],'' as [WayBill No],'' as [Transporter Code] ,'' as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP,  '' as  [Scheme Code],'' as [Scheme Type] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code], '' as Created_By ,'' as Modify_By ,TSPL_SCRAPINVOICE_DETAIL.Unit_code as RATE_UOM,0 as Conv_Factor," &
                                " (case when TSPL_SCRAPINVOICE_HEAD.Is_Taxable=1 then 'Tax Invoice' else 'Bill Of Supply' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN no Customer],(case when TSPL_SCRAPINVOICE_HEAD.Total_Tax_Amt<=0 and TSPL_SCRAPINVOICE_HEAD.Tax_Group<>'EXEMPTED' then (TSPL_SCRAPINVOICE_DETAIL.TotalAmt+Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end) else null end) as [Nill Rate Amount],(case when TSPL_SCRAPINVOICE_HEAD.Tax_Group='EXEMPTED' then (TSPL_SCRAPINVOICE_DETAIL.TotalAmt+Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end) else null end) as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive],"

            strSDEndQry = " from TSPL_SCRAPINVOICE_DETAIL  " &
                          " left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No " &
                          " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.AgainstScrap=TSPL_SCRAPINVOICE_HEAD.invoice_No " &
                          " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SCRAPINVOICE_HEAD.Cust_Code " &
                          " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
                          " left outer join TSPL_SCRAPSALE_HEAD on TSPL_SCRAPINVOICE_HEAD.shipment_No=TSPL_SCRAPSALE_HEAD.shipment_No "

            strSDJoinQry = " where 2 = 2 AND 'Misc Sale' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                          "  and convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) <= convert(date,('" & To_Date & "'),103) "
            Dim AllowtoenterrateIntoJobWorkDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowtoenterrateIntoJobWorkDispatch, clsFixedParameterCode.AllowtoenterrateIntoJobWorkDispatch, Nothing)) = 1, True, False)
            If AllowtoenterrateIntoJobWorkDispatch = True Then
                strSDJoinQry += " and TSPL_SCRAPINVOICE_HEAD.Doc_Type <>'J'"
            End If

            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SCRAPINVOICE_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SCRAPINVOICE_HEAD.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SCRAPINVOICE_HEAD.cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_SCRAPINVOICE_HEAD.invoice_No = '" & obj.Document_Code & "' "
            End If

            'sanjay
            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.MiscSaleSubCategory) > 0 Then
                If obj.MiscSaleSubCategory = "S" Then
                    strSDJoinQry += " and ISNULL(IS_SCRAP,'N')='Y' "
                ElseIf obj.MiscSaleSubCategory = "C" Then
                    strSDJoinQry += " and ISNULL(Is_CashSale,'N')='Y' "
                End If
            End If
            'sanjay

            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            '' query for no tax applied



            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_SCRAPINVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax5,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax7,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax9,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax10,'')='') )t "
            '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            '' query for tax1 applied
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX1 ,TSPL_SCRAPINVOICE_DETAIL.TAX1_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX1_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax1 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax1<>'')"
            strMCCMaterial += " s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX2 ,TSPL_SCRAPINVOICE_DETAIL.TAX2_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX2_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax2 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax2<>'' )"
            strMCCMaterial += " s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = "TSPL_SCRAPINVOICE_DETAIL.TAX3 ,TSPL_SCRAPINVOICE_DETAIL.TAX3_Amt , TSPL_SCRAPINVOICE_DETAIL.TAX3_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax3 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += "  select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax3<>'' )"
            strMCCMaterial += " s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX4 ,TSPL_SCRAPINVOICE_DETAIL.TAX4_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX4_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax4 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax4<>'' )"
            strMCCMaterial += " s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX5 ,TSPL_SCRAPINVOICE_DETAIL.TAX5_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX5_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX5+'%' as tax5rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax5 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax5<>'' )"
            strMCCMaterial += " s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX6 ,TSPL_SCRAPINVOICE_DETAIL.TAX6_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX6_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX6+'%' as tax6rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax6 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax6<>'' )"
            strMCCMaterial += " s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX7 ,TSPL_SCRAPINVOICE_DETAIL.TAX7_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX7_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX7+'%' as tax7rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax7 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax7<>'' )"
            strMCCMaterial += " s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + ")) t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX8 ,TSPL_SCRAPINVOICE_DETAIL.TAX8_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX8_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX8+'%' as tax8rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax8 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax8<>'' )"
            strMCCMaterial += "  s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX9 ,TSPL_SCRAPINVOICE_DETAIL.TAX9_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX9_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX9+'%' as tax9rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax9 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += "  select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax9<>'')"
            strMCCMaterial += " s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX10 ,TSPL_SCRAPINVOICE_DETAIL.TAX10_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX10_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX10+'%' as tax10rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SCRAPINVOICE_DETAIL.tax10 and ttr.tax_Rate=TSPL_SCRAPINVOICE_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strScarpCommonQry & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax10<>'')"
            strMCCMaterial += " s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " ) final left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Loc_Code  " &
                " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code  " &
                " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.cust_Code   " &
                " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No " &
                " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


            ''Monika
            ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            ''=Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            ''==end here


            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.shipment_No  ,final.Item_Code ,final.Loc_Code  ,final.cust_Code  ,final.shipped_Qty  ,final.TotalTaxAmt  ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.price  ,final.ItemAmt  ,final.DiscountPer,final.TotalDiscountAmt  ,final.Amt_less_Discount   ,final.Amt_Less_Discount ,final.Doc_Amt,final.[Scheme Amount],Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP ,  [Scheme Code],[Scheme Type] , [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Created_By ,final.Modify_By,final.RATE_UOM,final. Conv_Factor " ' ," + strPivotFoGrouprOuterQuery + "  ,QC.FAT_Per,QC.SNF_Per
            '' query for return
        End If
        If obj.Trans_Type_List.Contains("Dairy Sale Return") OrElse obj.Trans_Type_List.Contains("Fresh Sale Return") OrElse obj.Trans_Type_List.Contains("CSA Sale Return") OrElse obj.Trans_Type_List.Contains("Product Sale Return") OrElse obj.Trans_Type_List.Contains("MCC Sale Return") OrElse obj.Trans_Type_List.Contains("Export Sale Return") OrElse obj.Trans_Type_List.Contains("Sale Return") Then

            Dim strSDRCommonQuery As String = ""
            strSDRCommonQuery = " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,TSPL_SD_SALE_RETURN_HEAD.Ack_No as [Ack No],TSPL_SD_SALE_RETURN_HEAD.Ack_Date as [ACK Date],TSPL_SD_SALE_RETURN_HEAD.IRN_No as [IRN No],(case when len(TSPL_SD_SALE_RETURN_HEAD.Description)<=0 then TSPL_SD_SALE_RETURN_HEAD.Comments else TSPL_SD_SALE_RETURN_HEAD.Description end) as Narration,'' as Formtype,(CASE WHEN TSPL_SD_SALE_RETURN_HEAD.Trans_Type='ALL' THEN 'SDR' ELSE TSPL_SD_SALE_RETURN_HEAD.Trans_Type+'R' END) as Trans_Type,TSPL_SD_SALE_RETURN_HEAD.Status ,TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location, " &
                              " TSPL_SD_SALE_RETURN_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_RETURN_HEAD.Document_Type,TSPL_SD_SALE_RETURN_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_RETURN_HEAD.Document_Code , " &
                              " convert(varchar,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_RETURN_DETAIL.Item_Code," & IIf(Batch_Wise = True, "TSPL_BATCH_ITEM.Batch_No ,", " ") & " TSPL_SD_SALE_RETURN_DETAIL.Line_No , " &
                              " " & IIf(Batch_Wise = True, " - TSPL_BATCH_ITEM.Qty ", " -TSPL_SD_SALE_RETURN_DETAIL.Qty  ") & " as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code ,TSPL_SD_SALE_RETURN_DETAIL.Item_Cost , " &
                              " -1* (coalesce(TSPL_SD_SALE_RETURN_DETAIL.Amount,0)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_RETURN_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as Amount ,TSPL_SD_SALE_RETURN_DETAIL.Disc_Per ,case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)=0 then -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)  + case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then 1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end else -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end as Disc_Amt,case when coalesce(FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then -1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) end  as [Scheme Amount] , " &
                              " -(Amount- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then Total_Disc_Amt else Total_Disc_Amt end  + case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type<>'FS' then case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end else 0 end)  as Amt_Less_Discount , " &
                              " -1* (coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_RETURN_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as Total_Tax_Amt ,-(Amount+coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt else coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end )  as Total_Amt,TSPL_SD_SALE_RETURN_HEAD.Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,-1* ((case when TSPL_SD_SALE_RETURN_DETAIL.Line_No=1 then (coalesce(TSPL_SD_SALE_RETURN_HEAD.Total_Add_Charge,0)+coalesce(TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount,0)) else 0 end)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_RETURN_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as Additional_Charge, " &
                              " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],-TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                              " -TSPL_Customer_Invoice_Head.Discount_Amount-coalesce(TSPL_SD_SALE_RETURN_HEAD.headDisc_AMt,0) as [AR Document Discount Amt], " &
                              " -TSPL_Customer_Invoice_Head.amount_less_Discount as [AR Amount Before Tax],-TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                              " -(TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge], " &
                              " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                              " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_RETURN_HEAD.GRNo as [GR No],TSPL_SD_SHIPMENT_HEAD.gr_date as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],TSPL_SD_SALE_RETURN_DETAIL.Delivery_Code as [Delivery No],Against_Shipment_No as [Shipment No],
                              
                              TSPL_SD_SALE_RETURN_DETAIL.MRP, TSPL_SD_SALE_RETURN_DETAIL.Scheme_Code ,case when TSPL_SD_SALE_RETURN_DETAIL.scheme_item='N' then '' else TSPL_SD_SALE_RETURN_DETAIL. scheme_type end as Scheme_Type,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Code ,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Amount*(-1) as Cash_Scheme_Amount ,TSPL_SD_SALE_RETURN_DETAIL.Price_code ,'' as Created_By ,'' as Modify_By,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as RATE_UOM,0 as Conv_Factor, TSPL_SD_SALE_RETURN_DETAIL.Sampling,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Item," &
                              " 'Credit Note' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN no Customer],(case when TSPL_SD_SALE_RETURN_HEAD.Total_Tax_Amt<=0 and TSPL_SD_SALE_RETURN_HEAD.Tax_Group<>'EXEMPTED' then -(Amount+coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then 0 else coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end ) else null end) as [Nill Rate Amount],(case when TSPL_SD_SALE_RETURN_HEAD.Tax_Group='EXEMPTED' then -(Amount+coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then 0 else coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end ) else null end) as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code as [Original Invoice No],convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as [Original Invoice Date],TSPL_SD_SALE_RETURN_HEAD.Description as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive],"
            strSDEndQry = " from TSPL_SD_SALE_RETURN_DETAIL " &
                                " left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code =TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE " &
                                " left join TSPL_VEHICLE_MASTER on TSPL_SD_SALE_RETURN_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                                " left join TSPL_Customer_Invoice_Head on TSPL_SD_SALE_RETURN_HEAD.Document_Code=  case when len(isnull(TSPL_Customer_Invoice_Head.Against_Sale_Return_No,''))>0  then TSPL_Customer_Invoice_Head.Against_Sale_Return_No else TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return end " &
                                " left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code " &
                                " left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                                "  " &
                                " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_RETURN_HEAD.Customer_Code " &
                                " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
                                " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_SD_SALE_RETURN_DETAIL.TAX1=TAXM1.TAX_CODE " &
                                " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_SD_SALE_RETURN_DETAIL.TAX2=TAXM2.TAX_CODE " &
                                " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_SD_SALE_RETURN_DETAIL.TAX3=TAXM3.TAX_CODE "

            If Batch_Wise = True Then
                strSDEndQry += "left outer join  TSPL_BATCH_ITEM on  TSPL_BATCH_ITEM.document_code=TSPL_SD_SALE_RETURN_HEAD.Document_Code and TSPL_BATCH_ITEM.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.item_code and TSPL_BATCH_ITEM.UOM=TSPL_SD_SALE_RETURN_DETAIL.Unit_code "
            End If
            strSDJoinQry = " WHERE 2=2 AND (case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type IN ('FS','PS') and TSPL_SD_SALE_RETURN_HEAD.Screen_Type ='DS' then 'Dairy Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='FS' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type ='' then 'Fresh Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='CSA' then 'CSA Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='PS' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type ='' then 'Product Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' then 'MCC Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='EXP' then 'Export Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='Bulk Sale' then 'Bulk Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='SS' then 'Misc Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type in ('SD','All') then 'General Sale Return' else TSPL_SD_SALE_RETURN_HEAD.trans_Type  end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                            " and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_RETURN_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If

            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = '" & obj.Document_Code & "' "
            End If
            If obj.Scheme_Type_List IsNot Nothing AndAlso obj.Scheme_Type_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_RETURN_DETAIL.Scheme_Type in (" + clsCommon.GetMulcallString(obj.Scheme_Type_List) + ") "
            End If

            '' base union 5
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No]) as [Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type, max(final.Formtype) as [Form Type],case when Trans_Type ='DSR' then 'Dairy Sale Return' when Trans_Type ='FSR' then 'Fresh Sale Return' when Trans_Type ='CSAR' then 'CSA Sale Return' when Trans_Type='PSR' then 'Product Sale Return' when Trans_Type='MCCR' then 'MCC Sale Return' when Trans_Type='EXPR' then 'Export Sale Return'when Trans_Type='Bulk Sale' then 'Bulk Sale Return' when Trans_Type ='SSR' then 'Misc Sale' when Trans_Type ='SDR' then 'General Sale Return' else trans_Type  end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost],"

            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += "  0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount]," & IIf(Batch_Wise = True, " final.Amount + final.Total_Tax_AMt ", " final.Total_Amt ") & "  as [Total Amount], " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]  - case when (Trans_Type ='DSR' or Trans_Type ='FSR' or Trans_Type ='PSR') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],final.[Transporter Name], [Delivery No]  , [Shipment No],MRP, Scheme_Code as [Scheme Code] , Scheme_Type as [Scheme Type],Cash_Scheme_Code as [Cash Scheme Code], Cash_Scheme_Amount as [Cash Scheme Amount] ,final.Price_code as [Price Code] ,final.Created_By ,final.Modify_By,final.RATE_UOM ,final.Conv_Factor ,final. Sampling,final.Scheme_Item," &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive] " &
                " " & IIf(Batch_Wise = True, ",final.Batch_No ", " ") & " from ( "
            'max(isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'')) as [Executive] "
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            '' query for no tax applied
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax1,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax3,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax5,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax7,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax9,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax10,'')='') )t "

            strMCCMaterial += "   union all"
            '' query for tax1 applied
            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX1 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt as TAX1_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax1 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX2 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt as TAX2_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax2 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX3 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt as TAX3_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax3 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax7 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX4 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt as TAX4_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax4 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX5 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt as TAX5_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX5_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX5+'%' as tax5rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax5 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"

            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX6 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt as TAX6_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX6_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX6+'%' as tax6rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax6 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX7 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt as TAX7_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX7+'%' as tax7rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax7 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX8 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt as TAX8_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX8_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX8+'%' as tax8rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax8 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX9 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt as TAX9_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX9_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX9+'%' as tax9rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax9 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX10 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt as TAX10_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX10_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX10+'%' as tax10rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_RETURN_DETAIL.tax10 and ttr.tax_Rate=TSPL_SD_SALE_RETURN_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " )final"
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
            strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
            strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
            strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
            strMCCMaterial += " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code "
            strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                              " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "

            ''Monika
            ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            ''=========Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If

            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If
            ''======end here

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , Scheme_Code ,Scheme_Type,Cash_Scheme_Code , Cash_Scheme_Amount ,final.Price_code ,final.Created_By ,final.Modify_By ,final.RATE_UOM,final.Conv_Factor ,final. Sampling,final.Scheme_Item,[Executive] " & IIf(Batch_Wise = True, ",final.Batch_No ", " ") & "" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per
        End If
        '' base union 6
        If obj.Trans_Type_List.Contains("Bulk Sale Return") Then

            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            '''' bulk sale return 
            strMCCMaterial += "  select '' as _Type,'' as [Form Type],'Bulk Sale Return' as [Trans Type] ,TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code as [Location Code],TSPL_SALE_RETURN_MASTER_BULKSALE.Posted as Status,TSPL_LOCATION_MASTER.Location_Desc as [Location Name] ,'Invoice' as [Invoice Type] ,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No as [Document No] ,convert(varchar,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) as [Document_date],'' as [Narration],'' as Vehicle_Code,'' as Vehicle_No,coalesce(-1 * TSPL_SALE_RETURN_MASTER_BULKSALE.roundoffamount,0) as Additional_Charge, " &
                               " TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code as [Customer Code],TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As [Customer Address] ,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER .GST_Registered as [Registered],TSPL_CUSTOMER_MASTER .GST_COMPOSITION as [Composition],TSPL_CUSTOMER_MASTER .City_Code as [City Code],TSPL_CITY_MASTER .City_Name as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,TSPL_CUSTOMER_MASTER.Parent_Customer_No as [Parent Customer No]," &
                               " Parent_Master.Cust_Code as [Parent Customer Code],Parent_Master.Customer_Name as [Parent Customer Name] ," &
                               " TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code as [Item Code],tspl_item_master.Item_Desc as [Item Name],tspl_item_master.HSN_Code as [HSN Code] ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceQty as [Quantity] ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as [UOM],TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceRate as [Item Cost],TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceFatPer as [Fat Per] ,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceSNFPer as [SNF Per] ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceFatKG as [Fat Kg] ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceSNFKG as [SNF KG]  ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as Amount,0 as [Discount Per],0 as [Discount Amount],0 as [Scheme Amount],-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as [Amount Less Discount] " + strPivotForOuterQueryforBulk + " " + strDoublePivotForOuterQueryforBulk + ",-TSPL_SALE_RETURN_DETAIL_BULKSALE.Total_Tax_Amt as [Total Tax Amount],-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as [Total Amount], " &
                               " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],-1 * TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                               " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                               " -1 * (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                               " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],NULL as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],'' as [Delivery No]  ,'' as  [Shipment No],0 AS MRP, '' as [Scheme Code],'' as [Scheme Type] ,'' as  [Cash Scheme Code] , 0 as [Cash Scheme Amount], '' as [Price Code] ,TSPL_SALE_RETURN_MASTER_BULKSALE.Created_By as Created_By,TSPL_SALE_RETURN_MASTER_BULKSALE.Modified_By as Modify_By,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as RATE_UOM,0 as Conv_Factor,0 as Sampling,'N' as Scheme_Item," &
                               " 'Credit Note' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer],-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as [Nill Rate Amount],0 as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],TSPL_SALE_RETURN_MASTER_BULKSALE.InvoiceNo as [Original Invoice No],convert(varchar,TSPL_INVOICE_MASTER_BULKSALE.Document_Date,103) as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT ,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive]" &
                               " from TSPL_SALE_RETURN_DETAIL_BULKSALE "

            strMCCMaterial += " left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No " &
                              " left join TSPL_INVOICE_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE.InvoiceNo= TSPL_INVOICE_MASTER_BULKSALE.Document_No "
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code"
            strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code"
            strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No"
            strMCCMaterial += " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code"
            strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                              " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "
            strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code"
            strMCCMaterial += " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_Return_No=TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No  where 2=2 " &
                                " and 'Bulk Sale Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                                " and convert(date,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

            ''=========Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If

            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            ''=end here
        End If
        If obj.Trans_Type_List.Contains("MCC Transfer") Then

            '' base union 7
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += "   select * " &
                " from (Select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,'' AS [Ack No],'' as [ACK Date],'' as [IRN No],'' as _Type, '' as [Form Type],'MCC Transfer' as [Trans Type] ,TSPL_MCC_Dispatch_Challan.MCC_Code as  [Location Code],TSPL_MCC_Dispatch_Challan.isPosted as " _
                & " Status, sendr.location_desc as  [Location Name],'MCC Transfer' as [Invoice Type],TSPL_MCC_Dispatch_Challan.Chalan_NO as [Document No] , " _
                & " convert(varchar,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103 ) as [Document_date],'' as [Narration], '' as Vehicle_Code,tm.tanker_NO as Vehicle_No,0  as Additional_Charge , " _
                & " tspl_milk_Transfer_In.location_Code as [Customer Code],'' AS [Customer Address],  recv.Location_Desc  as [Customer Name],recv.Registered as [Registered],'' as [Composition],'' as [City Code],recv.City_Code as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,'' as [Parent Customer No],'' as [Parent Customer Code]," _
                & " '' as [Parent Customer Name], TSPL_MCC_Dispatch_Challan.Item_Code as [Item Code], TSPL_MCC_Dispatch_Challan.Item_Desc as [Item Name],'' as [HSN Code] , TSPL_MCC_Dispatch_Challan.Net_Qty  as [Quantity] " _
                & " ,TSPL_MCC_Dispatch_Challan.UOM_Code as  [UOM] ,round(((TSPL_MCC_Dispatch_Challan.FAT_RATE *(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) " _
                & " * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)) +(TSPL_MCC_Dispatch_Challan.SNF_RATE  *(coalesce(cast(t_SNF_Recd.Param_Field_Value as float),0) * " _
                & " coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)))/coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,1) ,2) as  [Item Cost] ,  cast(t_FAT_Recd.Param_Field_Value as numeric(10,2)) as [FAT Per]" _
                & " ,cast(t_SNF_Recd.Param_Field_Value as numeric(10,2)) as [SNF Per],(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) " _
                & " as [FAT KG],(coalesce(cast(t_Snf_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) as [SNF KG],amount as  Amount ,0" _
                & " as [Discount Per]  ,0 as [Discount Amount],0 as [Scheme Amount] ,  amount as  [Amount Less Discount]   " & strPivotForTransfer_In & "" & strPivotFortRANSFER_INPercentQuery & ",   0 as [Total Tax Amount] " _
                & " ,amount as   [Total Amount],TSPL_vendor_Invoice_Head.Document_No as [AR Document No],TSPL_vendor_Invoice_Head.Document_Total [AR Document Amt]" _
                & " ,TSPL_vendor_Invoice_Head.Discount_Amount as [AR Document Discount Amt],TSPL_vendor_Invoice_Head.amount_less_Discount as [AR Amount Before Tax]," _
                & " TSPL_vendor_Invoice_Head.total_tax as [AR Total Tax],TSPL_vendor_Invoice_Head.total_Add_Charge as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No]," _
                & " TSPL_MCC_Dispatch_Challan.tanker_No as [Transporter Code],tm.description as [Transporter Name],'' as  [Delivery No],'' as [Shipment No],0 AS MRP,'' as  [Scheme Code],'' as [Scheme Type] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],'' as Created_By ,'' as Modify_By ,TSPL_MCC_Dispatch_Challan.UOM_Code as RATE_UOM,0 as Conv_Factor,0 as Sampling,'N' as Scheme_Item,'Delivery Challan' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],recv.GstNo as  [GSTIN No Customer],cast(null as numeric(18,2)) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT,'' as [Executive] from TSPL_MCC_Dispatch_Challan  left outer  join TSPL_MILK_TRANSFER_IN on TSPL_MILK_TRANSFER_IN.Dispatch_Challan_No " _
                & " =TSPL_MCC_Dispatch_Challan.Chalan_NO  LEFT JOIN tspl_location_Master  sendr ON sendr.Location_Code=TSPL_MCC_Dispatch_Challan.MCC_CODE " _
                & " left join tspl_Location_master on tspl_Location_master.location_code=TSPL_MCC_Dispatch_Challan.mcc_Code left join tspl_Location_master recv on " _
                & " recv.location_code=TSPL_MILK_TRANSFER_IN.Location_Code left join TSPL_STATE_MASTER on recv.State=TSPL_STATE_MASTER.STATE_CODE left join TSPL_vendor_Invoice_Head on  TSPL_vendor_Invoice_Head.vendor_Invoice_No" _
                & " =TSPL_MILK_TRANSFER_IN.Receipt_Challan_No left join tspl_tanker_Master tm on tm.tanker_no=TSPL_MCC_Dispatch_Challan.tanker_No Left Outer Join " _
                & " (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.* From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On " _
                & " TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'SNF') " _
                & " t_SNF_Recd On t_SNF_Recd.Chalan_NO   = TSPL_MCC_Dispatch_Challan.Chalan_NO   Left Outer Join (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.*" _
                & " From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = " _
                & " TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'FAT' ) t_FAT_Recd On t_FAT_Recd.Chalan_No " _
                & " = TSPL_MCC_Dispatch_Challan.Chalan_NO where 2=2 " &
                  " AND 'MCC Transfer' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                  " and convert(date,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103) <= convert(date,('" & To_Date & "'),103) )t "

        End If
        ''richa agarwal 25 May,2017
        If obj.Trans_Type_List.Contains("Tanker Dispatch Return") Then

            '' base union 8
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += "   select * from (Select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,'' AS [Ack No],'' as [ACK Date],'' as [IRN No],'' as _Type, '' as [Form Type],'Tanker Dispatch Return' as [Trans Type] ,TSPL_MCC_Dispatch_Challan.MCC_Code as  [Location Code],TSPL_MCC_Dispatch_Challan.isPosted as " _
                & " Status, sendr.location_desc as  [Location Name],'Tanker Dispatch Return' as [Invoice Type],TSPL_MCC_Dispatch_Challan_Return.Document_No as [Document No] , " _
                & " convert(varchar,TSPL_MCC_Dispatch_Challan_Return.Document_Date,103 ) as [Document_date],TSPL_MCC_Dispatch_Challan_Return.Remarks as [Narration], '' as Vehicle_Code,tm.tanker_NO as Vehicle_No,0  as Additional_Charge , " _
                & " tspl_milk_Transfer_In.location_Code as [Customer Code],'' AS [Customer Address],  recv.Location_Desc  as [Customer Name],sendr.Registered as [Registered],'' as [Composition],sendr.City_Code as [City Code],sendr.City_Code as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,'' as [Parent Customer No],'' as [Parent Customer Code]," _
                & " '' as [Parent Customer Name], TSPL_MCC_Dispatch_Challan.Item_Code as [Item Code], TSPL_MCC_Dispatch_Challan.Item_Desc as [Item Name],'' as [HSN Code] , TSPL_MCC_Dispatch_Challan.Net_Qty * -1  as [Quantity] " _
                & " ,TSPL_MCC_Dispatch_Challan.UOM_Code as  [UOM] ,round(((TSPL_MCC_Dispatch_Challan.FAT_RATE *(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) " _
                & " * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)) +(TSPL_MCC_Dispatch_Challan.SNF_RATE  *(coalesce(cast(t_SNF_Recd.Param_Field_Value as float),0) * " _
                & " coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)))/coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,1) ,2) as  [Item Cost] ,  t_FAT_Recd.Param_Field_Value as [FAT Per]" _
                & " ,t_SNF_Recd.Param_Field_Value as [SNF Per],(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) * -1 " _
                & " as [FAT KG],(coalesce(cast(t_Snf_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) * -1 as [SNF KG],amount * -1 as  Amount ,0" _
                & " as [Discount Per]  ,0 as [Discount Amount],0 as [Scheme Amount] ,  amount * -1 as  [Amount Less Discount]   " & strPivotForTransfer_In & "" & strPivotFortRANSFER_INPercentQuery & ",   0 as [Total Tax Amount] " _
                & " ,amount * -1 as   [Total Amount],TSPL_vendor_Invoice_Head.Document_No as [AR Document No],TSPL_vendor_Invoice_Head.Document_Total [AR Document Amt]" _
                & " ,TSPL_vendor_Invoice_Head.Discount_Amount as [AR Document Discount Amt],TSPL_vendor_Invoice_Head.amount_less_Discount as [AR Amount Before Tax]," _
                & " TSPL_vendor_Invoice_Head.total_tax as [AR Total Tax],TSPL_vendor_Invoice_Head.total_Add_Charge as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No]," _
                & " TSPL_MCC_Dispatch_Challan.tanker_No as [Transporter Code],tm.description as [Transporter Name],'' as  [Delivery No],'' as [Shipment No],0 AS MRP,'' as  [Scheme Code] ,'' as [Scheme Type],'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],'' as Created_By ,'' as Modify_By,TSPL_MCC_Dispatch_Challan.UOM_Code as RATE_UOM,0 as Conv_Factor,0 as Sampling,'N' as Scheme_Item,'Credit Note' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],sendr.GSTNO as [GSTIN No Customer],cast(null as numeric(18,2)) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT,'' as [Executive] from TSPL_MCC_Dispatch_Challan  left outer  join TSPL_MILK_TRANSFER_IN on TSPL_MILK_TRANSFER_IN.Dispatch_Challan_No " _
                & " =TSPL_MCC_Dispatch_Challan.Chalan_NO  LEFT JOIN tspl_location_Master  sendr ON sendr.Location_Code=TSPL_MCC_Dispatch_Challan.MCC_CODE " _
                & " left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.State_Code=sendr.state left join tspl_Location_master recv on " _
                & " recv.location_code=TSPL_MILK_TRANSFER_IN.Location_Code left join TSPL_vendor_Invoice_Head on  TSPL_vendor_Invoice_Head.vendor_Invoice_No" _
                & " =TSPL_MILK_TRANSFER_IN.Receipt_Challan_No left join tspl_tanker_Master tm on tm.tanker_no=TSPL_MCC_Dispatch_Challan.tanker_No Left Outer Join " _
                & " (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.* From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On " _
                & " TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'SNF') " _
                & " t_SNF_Recd On t_SNF_Recd.Chalan_NO   = TSPL_MCC_Dispatch_Challan.Chalan_NO   Left Outer Join (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.*" _
                & " From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = " _
                & " TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'FAT' ) t_FAT_Recd On t_FAT_Recd.Chalan_No " _
                & " = TSPL_MCC_Dispatch_Challan.Chalan_NO left outer join TSPL_MCC_Dispatch_Challan_Return on TSPL_MCC_Dispatch_Challan_Return.Challan_No =TSPL_MCC_Dispatch_Challan.Chalan_NO where 2=2 and TSPL_MCC_Dispatch_Challan_Return.Challan_No =TSPL_MCC_Dispatch_Challan.Chalan_NO " &
                  " AND 'Tanker Dispatch Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                  " and convert(date,TSPL_MCC_Dispatch_Challan_Return.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_MCC_Dispatch_Challan_Return.Document_Date,103) <= convert(date,('" & To_Date & "'),103) )t "

        End If
        If obj.Trans_Type_List.Contains("MCC Tanker Dispatch Return") OrElse obj.Trans_Type_List.Contains("Tanker Dispatch Return") Then

            '' base union 9
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += "   select * from (Select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,'' AS [Ack No],'' as [ACK Date],'' as [IRN No],'' as _Type, '' as [Form Type],'MCC Tanker Dispatch Return' as [Trans Type] ,TSPL_MCC_Tanker_Dispatch_Return_head.MCC_Code as  [Location Code],TSPL_MCC_Tanker_Dispatch_Return_head.isPosted as " _
                & " Status, sendr.location_desc as  [Location Name],'MCC Tanker Dispatch Return' as [Invoice Type],TSPL_MCC_Tanker_Dispatch_Return_head.Return_NO as [Document No] , " _
                & " convert(varchar,TSPL_MCC_Tanker_Dispatch_Return_head.Return_Date ,103 ) as [Document_date],'' as [Narration], '' as Vehicle_Code,tm.tanker_NO as Vehicle_No,0  as Additional_Charge , " _
                & " tspl_milk_Transfer_In.location_Code as [Customer Code],'' AS [Customer Address],  recv.Location_Desc  as [Customer Name],recv.Registered as [Registered],'' as [Composition],'' as [City Code],sendr.City_Code as [Place of Supply],TSPL_STATE_MASTER.GST_STATE_Code AS [Customer GST State Code] ,'' as [Parent Customer No],'' as [Parent Customer Code]," _
                & " '' as [Parent Customer Name], TSPL_MCC_Tanker_Dispatch_Return_head.Item_Code as [Item Code], TSPL_MCC_Tanker_Dispatch_Return_head.Item_Desc as [Item Name],'' as [HSN Code] , TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty * -1  as [Quantity] " _
                & " ,TSPL_MCC_Tanker_Dispatch_Return_head.UOM_Code as  [UOM] ,round(((TSPL_MCC_Tanker_Dispatch_Return_head.FAT_RATE *(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) " _
                & " * coalesce(TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty,0)/100)) +(TSPL_MCC_Tanker_Dispatch_Return_head.SNF_RATE  *(coalesce(cast(t_SNF_Recd.Param_Field_Value as float),0) * " _
                & " coalesce(TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty,0)/100)))/coalesce(TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty,1) ,2) as  [Item Cost] ,  t_FAT_Recd.Param_Field_Value as [FAT Per]" _
                & " ,t_SNF_Recd.Param_Field_Value as [SNF Per],(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty,0)/100) * -1 " _
                & " as [FAT KG],(coalesce(cast(t_Snf_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Tanker_Dispatch_Return_head.Net_Qty,0)/100) * -1 as [SNF KG],TSPL_MCC_Tanker_Dispatch_Return_head.Amount * -1 as  Amount ,0" _
                & " as [Discount Per]  ,0 as [Discount Amount],0 as [Scheme Amount] ,  TSPL_MCC_Tanker_Dispatch_Return_head.Amount * -1 as  [Amount Less Discount]   " & strPivotForTransfer_In & "" & strPivotFortRANSFER_INPercentQuery & ",   0 as [Total Tax Amount] " _
                & " ,TSPL_MCC_Tanker_Dispatch_Return_head.Amount * -1 as   [Total Amount],TSPL_vendor_Invoice_Head.Document_No as [AR Document No],TSPL_vendor_Invoice_Head.Document_Total [AR Document Amt]" _
                & " ,TSPL_vendor_Invoice_Head.Discount_Amount as [AR Document Discount Amt],TSPL_vendor_Invoice_Head.amount_less_Discount as [AR Amount Before Tax]," _
                & " TSPL_vendor_Invoice_Head.total_tax as [AR Total Tax],TSPL_vendor_Invoice_Head.total_Add_Charge as [AR Total Add Charge],'' as [GR No],null as [GR Date],'' as [WayBill No]," _
                & " TSPL_MCC_Tanker_Dispatch_Return_head.tanker_No as [Transporter Code],tm.description as [Transporter Name],'' as  [Delivery No],'' as [Shipment No],0 AS MRP,'' as  [Scheme Code],'' as [Scheme Type] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],'' as Created_By ,'' as Modify_By,TSPL_MCC_Tanker_Dispatch_Return_head.UOM_Code as RATE_UOM,0 as Conv_Factor,0 as Sampling,'N' as Scheme_Item,'Credit Note' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],sendr.GSTNO as [GSTIN no Customer],cast(null as numeric(18,2)) as [Nill Rate Amount],cast(null as numeric(18,2)) as [Exempted Amount],cast(null as numeric(18,2)) as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT,'' as [Executive] from TSPL_MCC_Dispatch_Challan  left outer  join TSPL_MILK_TRANSFER_IN on TSPL_MILK_TRANSFER_IN.Dispatch_Challan_No " _
                & " =TSPL_MCC_Dispatch_Challan.Chalan_NO  LEFT JOIN tspl_location_Master  sendr ON sendr.Location_Code=TSPL_MCC_Dispatch_Challan.MCC_CODE " _
                & " left join TSPL_STATE_MASTER on sendr.State=TSPL_STATE_MASTER.STATE_CODE left join tspl_Location_master recv on " _
                & " recv.location_code=TSPL_MILK_TRANSFER_IN.Location_Code left join TSPL_vendor_Invoice_Head on  TSPL_vendor_Invoice_Head.vendor_Invoice_No" _
                & " =TSPL_MILK_TRANSFER_IN.Receipt_Challan_No left join tspl_tanker_Master tm on tm.tanker_no=TSPL_MCC_Dispatch_Challan.tanker_No Left Outer Join " _
                & " (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.* From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On " _
                & " TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'SNF') " _
                & " t_SNF_Recd On t_SNF_Recd.Chalan_NO   = TSPL_MCC_Dispatch_Challan.Chalan_NO   Left Outer Join (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.*" _
                & " From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = " _
                & " TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'FAT' ) t_FAT_Recd On t_FAT_Recd.Chalan_No " _
                & " = TSPL_MCC_Dispatch_Challan.Chalan_NO left outer join TSPL_MCC_Tanker_Dispatch_Return_head on TSPL_MCC_Tanker_Dispatch_Return_head.Chalan_No =TSPL_MCC_Dispatch_Challan.Chalan_NO where 2=2 and TSPL_MCC_Tanker_Dispatch_Return_head.Chalan_No =TSPL_MCC_Dispatch_Challan.Chalan_NO " &
                  " AND 'MCC Tanker Dispatch Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                  " and convert(date,TSPL_MCC_Tanker_Dispatch_Return_head.Return_Date ,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_MCC_Tanker_Dispatch_Return_head.Return_Date ,103) <= convert(date,('" & To_Date & "'),103) )t "

        End If

        If obj.Trans_Type_List.Contains("Transfer") Then

            ''----------------------- end of richa agarwal
            Dim strTranferCommonQuery As String = ""
            strTranferCommonQuery = " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,TSPL_TRANSFER_ORDER_HEAD.Ack_No as [Ack No],TSPL_TRANSFER_ORDER_HEAD.Ack_Date as [ACK Date],TSPL_TRANSFER_ORDER_HEAD.IRN_No as [IRN No],TSPL_TRANSFER_ORDER_HEAD.Description as Narration,case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF = 1 then 'F' else '' End as Formtype,'Transfer' as Trans_Type,TSPL_TRANSFER_ORDER_HEAD.Status ,TSPL_TRANSFER_ORDER_HEAD.From_Location as Bill_To_Location, " &
                                    " GIT_Main_Loc.Location_Code as To_Location,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type as Invoice_Type,TSPL_TRANSFER_ORDER_HEAD.Document_No , " &
                                    " convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103 ) as Document_Date , TSPL_TRANSFER_ORDER_DETAIL.Item_Code , " &
                                    " (case when TSPL_TRANSFER_ORDER_DETAIL.In_Qty>0 then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_TRANSFER_ORDER_DETAIL.Out_Qty end) as Qty, " &
                                    " TSPL_TRANSFER_ORDER_DETAIL.Unit_code ,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost , " &
                                    " TSPL_TRANSFER_ORDER_DETAIL.Amount ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt ,0 as [Scheme Amount], " &
                                    " (TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Amt_Less_Discount ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt ,(TSPL_TRANSFER_ORDER_DETAIL.Amount+TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Total_Amt,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code,COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No) as Vehicle_No,0 as Additional_Charge, " &
                                    " '' as [AR Document No],0 as  [AR Document Amt],0 as [AR Document Discount Amt],0 as  [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_TRANSFER_ORDER_HEAD.GR_No as [GR No],TSPL_TRANSFER_ORDER_HEAD.gr_date as [GR Date],TSPL_TRANSFER_ORDER_HEAD.WayBill_No as [WayBill No],TSPL_TRANSFER_ORDER_HEAD.Transport_Id as [Transporter Code],case when len(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual) > 0 then TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual else TSPL_TRANSPORT_MASTER.Transporter_Name end  as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP, '' as [Scheme Code] ,'' as [Scheme Type],'' as  [Cash Scheme Code] , 0 as [Cash Scheme Amount], '' as [Price Code],'' as Created_By,'' as Modified_By,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as RATE_UOM,0 as Conv_Factor, " &
                                    " (CASE WHEN From_LOC.STATE=GIT_Main_Loc.TO_STATE THEN  'Delivery Challan' WHEN TSPL_TRANSFER_ORDER_HEAD.TOTAL_TAX_AMT>0 AND TSPL_TAX_MASTER.TYPE='IGST' then 'Tax Invoice-Stock Transfer' else 'Bill of Supply' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],'' as [GSTIN no Customer],(case when TSPL_TRANSFER_ORDER_Head.Total_Tax_Amt<=0 and TSPL_TRANSFER_ORDER_Head.Tax_Group<>'EXEMPTED' then (TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) else null end) as [Nill Rate Amount],(case when TSPL_TRANSFER_ORDER_HEAD.Tax_Group='EXEMPTED' then (TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) else null end) as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],'' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,'' as [Executive],"

            strSDEndQry = " from TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No " &
                          " left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_Head.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                          " left join TSPL_TRANSPORT_MASTER on TSPL_TRANSFER_ORDER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                          " left join ( select  max(Location_Code) as Location_Code,MAX(STATE) AS TO_STATE,GIT_Location from TSPL_LOCATION_MASTER where GIT_Location is not null " &
                          " group by GIT_Location ) GIT_Main_Loc on TSPL_TRANSFER_ORDER_HEAD.To_Location=GIT_Main_Loc.GIT_Location  " &
                          " LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE " &
                          " left join TSPL_TAX_MASTER on  TSPL_TRANSFER_ORDER_HEAD.TAX1=TSPL_TAX_MASTER.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_TRANSFER_ORDER_DETAIL.TAX1=TAXM1.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_TRANSFER_ORDER_DETAIL.TAX2=TAXM2.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_TRANSFER_ORDER_DETAIL.TAX3=TAXM3.TAX_CODE "

            strSDJoinQry = "  where Transfer_Type <>'I' AND 'Transfer' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                          " and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_ORDER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_ORDER_HEAD.From_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and GIT_Main_Loc.Location_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = '" & obj.Document_Code & "' "
            End If
            '' base union 10
            '' transaction unit conversion
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += " select  0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No])[Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type, max(Formtype) as [Form Type],Trans_type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_No as [Document No],final.Document_Date as [Document_date],max(Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.To_Location as [Customer Code],'' As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Location_Desc) as [Customer Name],max(TSPL_CUSTOMER_MASTER.Registered) as [Registered],'' as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CUSTOMER_MASTER .City_Code) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,'' as [Parent Customer No] ,'' as [Parent Customer Code],'' as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount]," &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,  [Scheme Code],[Scheme Type] , [Cash Scheme Code] ,  [Cash Scheme Amount],  [Price Code],final.Created_By as Created_By ,final.Modified_By as Modify_By,final.RATE_UOM,final.Conv_Factor,0 as Sampling,'N' as Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max(TSPL_CUSTOMER_MASTER.GSTNO) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive] " &
                " from ( "
            'strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,0 as TAX1_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "

            '' query for no tax applied
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax1,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax3,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax5,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax7,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax9,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax10,'')='') )t "
            '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " union all "
            '' quert for no tax applied
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax1 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX2 ,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax2 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX3 ,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax3 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX4 ,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax4 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX5 ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5+'%' as tax5rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax5 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"

            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX6 ,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX6+'%' as tax6rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax6 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX7 ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7+'%' as tax7rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax7 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX8 ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8+'%' as tax8rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax8 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX9 ,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9+'%' as tax9rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax9 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX10 ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10+'%' as tax10rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax10 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"

            strMCCMaterial += " )final"
            ''-------------------------
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
            strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location " &
                              " left join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE"



            ''Monika
            ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
            'strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            ''=Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            ''end here

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_No ,final.Item_Code ,final.Bill_To_Location ,final.To_Location ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,[Scheme Code] ,[Scheme Type],[Cash Scheme Code] ,[Cash Scheme Amount], [Price Code] ,final.Created_By ,final.Modified_By,final. RATE_UOM,final.Conv_Factor" '',QC.FAT_Per,QC.SNF_Per
            ''richa
        End If
        strMCCMaterial += Environment.NewLine + " --- QUERY FOR TRANSFER RETURN---------------------- ADDED BY RICHA AGARWAL ------------" + Environment.NewLine

        If obj.Trans_Type_List.Contains("Transfer Return") Then

            Dim strTranferReturnCommonQuery As String = ""
            strTranferReturnCommonQuery = " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' dcscode,'' as Supply_Date,'' AS Payment_Terms,'' as CustomerType,'' as PaymentType,TSPL_TRANSFER_ORDER_HEAD.Ack_No as [Ack No],TSPL_TRANSFER_ORDER_HEAD.Ack_Date as [ACK Date],TSPL_TRANSFER_ORDER_HEAD.IRN_No as [IRN No],(case when len(TSPL_TRANSFER_ORDER_HEAD.description)<=0 then TSPL_TRANSFER_ORDER_HEAD.Remarks else TSPL_TRANSFER_ORDER_HEAD.description end) as Narration,case	when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF = 1 then 'F' else '' End as Formtype,'Transfer Return' as Trans_Type,TSPL_TRANSFER_ORDER_HEAD.Status ,  TSPL_TRANSFER_ORDER_HEAD.From_Location as Bill_To_Location, " &
                                    " GIT_Main_Loc.Location_Code as To_Location,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type as Invoice_Type,TSPL_TRANSFER_RETURN.Document_No , " &
                                    " convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103 ) as Document_Date , TSPL_TRANSFER_ORDER_DETAIL.Item_Code , " &
                                    " -(case when TSPL_TRANSFER_ORDER_DETAIL.In_Qty>0 then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_TRANSFER_ORDER_DETAIL.Out_Qty end) as Qty, " &
                                    " TSPL_TRANSFER_ORDER_DETAIL.Unit_code ,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost , " &
                                    " -TSPL_TRANSFER_ORDER_DETAIL.Amount AS Amount ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per ,-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt as Disc_Amt ,0 as [Scheme Amount], " &
                                    " -(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Amt_Less_Discount ,-TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt as Total_Tax_Amt ,-(TSPL_TRANSFER_ORDER_DETAIL.Amount+TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Total_Amt,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code,COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No) as Vehicle_No,0 as Additional_Charge, " &
                                    " '' as [AR Document No],0 as  [AR Document Amt],0 as [AR Document Discount Amt],0 as  [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_TRANSFER_ORDER_HEAD.GR_No as [GR No],TSPL_TRANSFER_ORDER_HEAD.gr_date as [GR Date],TSPL_TRANSFER_ORDER_HEAD.WayBill_No as [WayBill No],TSPL_TRANSFER_ORDER_HEAD.Transport_Id as [Transporter Code],case when len(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual) > 0 then TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual else TSPL_TRANSPORT_MASTER.Transporter_Name end as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP, '' as [Scheme Code] ,'' as [Scheme Type], '' as [Cash Scheme Code] ,0 as   [Cash Scheme Amount],  '' as [Price Code],'' as Created_By ,'' as Modified_By,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as RATE_UOM,0 as Conv_Factor," &
                                    " (CASE WHEN From_LOC.STATE=GIT_Main_Loc.TO_STATE THEN  'Delivery Challan' WHEN TSPL_TRANSFER_ORDER_HEAD.TOTAL_TAX_AMT>0 AND TSPL_TAX_MASTER.TYPE='IGST' then 'Tax Invoice-Stock Transfer' else 'Bill of Supply' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],'' as [GSTIN no Customer],(case when TSPL_TRANSFER_ORDER_Head.Total_Tax_Amt<=0 and TSPL_TRANSFER_ORDER_Head.Tax_Group<>'EXEMPTED' then -(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) else null end) as [Nill Rate Amount],(case when TSPL_TRANSFER_ORDER_HEAD.Tax_Group='EXEMPTED' then -(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) else null end) as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],TSPL_TRANSFER_ORDER_HEAD.Document_No as [Original Invoice No],Convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as [Original Invoice Date],'' as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,'' as [Executive],"

            strSDEndQry = " from TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No " &
                " Left Outer Join TSPL_TRANSFER_RETURN on TSPL_TRANSFER_RETURN.Transfer_No=TSPL_TRANSFER_ORDER_HEAD.Document_No " &
                          " left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_Head.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                          " left join TSPL_TRANSPORT_MASTER on TSPL_TRANSFER_ORDER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                          " left join ( select  max(Location_Code) as Location_Code,max(state) as TO_STATE,GIT_Location from TSPL_LOCATION_MASTER where GIT_Location is not null " &
                          " group by GIT_Location ) GIT_Main_Loc on TSPL_TRANSFER_ORDER_HEAD.To_Location=GIT_Main_Loc.GIT_Location  " &
                          " LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE " &
                          " left join TSPL_TAX_MASTER on  TSPL_TRANSFER_ORDER_HEAD.TAX1=TSPL_TAX_MASTER.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_TRANSFER_ORDER_DETAIL.TAX1=TAXM1.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_TRANSFER_ORDER_DETAIL.TAX2=TAXM2.TAX_CODE " &
                          " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_TRANSFER_ORDER_DETAIL.TAX3=TAXM3.TAX_CODE "

            strSDJoinQry = " where TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ='O' AND 'Transfer Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                          " and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) AND ISNULL(TSPL_TRANSFER_RETURN.Document_No ,'')<>'' "

            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_ORDER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_ORDER_HEAD.From_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and  GIT_Main_Loc.Location_Code  in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_TRANSFER_RETURN.Document_No = '" & obj.Document_Code & "' "
            End If
            '' base union 11
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode, max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No])[Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type,max(Formtype) as [Form Type],Trans_type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER.Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_No as [Document No],final.Document_Date as [Document_date],max(Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge, final.To_Location as [Customer Code],'' As [Customer Address] ,max(TSPL_CUSTOMER_MASTER.Location_Desc) as [Customer Name],max(TSPL_LOCATION_MASTER.Registered) as [Registered],'' as [Composition],max(TSPL_LOCATION_MASTER.City_Code) as [City Code],max(TSPL_LOCATION_MASTER.City_Code) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,'' as [Parent Customer No] ,'' as [Parent Customer Code],'' as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "


            ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount]," &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No] ,MRP, [Scheme Code] ,[Scheme Type], [Cash Scheme Code] ,  [Cash Scheme Amount],  [Price Code] ,final.Created_By as Created_By ,final.Modified_By as Modify_By,final. RATE_UOM,final. Conv_Factor,0 as Sampling,'N' as Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max(TSPL_LOCATION_MASTER.GSTNO) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,'' as [Executive] " &
                " from ( "
            'strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,0 as TAX1_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "

            '' query for no tax applied
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax1,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax3,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax5,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax7,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax9,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax10,'')='') )t "

            strMCCMaterial += Environment.NewLine + " union all "
            '' quert for no tax applied
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT AS TAX1_AMT ,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax1 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX2 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt AS TAX2_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax2 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX3 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt AS TAX3_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax3 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX4 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt AS TAX4_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax4 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX5 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt AS TAX5_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5+'%' as tax5rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax5 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"

            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX6 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt AS TAX6_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX6+'%' as tax6rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax6 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX7 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt AS TAX7_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7+'%' as tax7rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax7 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX8 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt AS TAX8_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8+'%' as tax8rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax8 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX9 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt AS TAX9_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9+'%' as tax9rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax9 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX10 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt AS TAX10_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10+'%' as tax10rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_TRANSFER_ORDER_DETAIL.tax10 and ttr.tax_Rate=TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t" &
             " )final" &
             " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location " &
             " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code " &
             " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location " &
             " left join TSPL_STATE_MASTER on TSPL_LOCATION_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


            ''Monika
            ''" left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            ''Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            ''=End here

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_No ,final.Item_Code ,final.Bill_To_Location ,final.To_Location ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,  [Scheme Code] ,[Scheme Type],  [Cash Scheme Code] , [Cash Scheme Amount],  [Price Code] ,final.Created_By ,final.Modified_By,final. RATE_UOM,final.Conv_Factor" + Environment.NewLine '',QC.FAT_Per,QC.SNF_Per
        End If
        strMCCMaterial += Environment.NewLine + " --- QUERY FOR TRANSFER RETURN END---------------------- ADDED BY RICHA AGARWAL ------------" + Environment.NewLine
        If obj.Include_MCCFarmerSale = True Then
            If obj.Trans_Type_List.Contains("MCC Sale Farmer") Then
                strSDCommonQuery = " select TSPL_MCC_SALE_FARMER_HEAD.Remarks as Narration,'' as Formtype,'MCC Sale Farmer' as Trans_Type ," &
                                   " TSPL_MCC_SALE_FARMER_HEAD.Status ,TSPL_MCC_SALE_FARMER_HEAD.Bill_To_Location, " &
                                   " TSPL_MCC_SALE_FARMER_HEAD.Farmer_Code as Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2  As CustAdd, " &
                                   " TSPL_MCC_SALE_FARMER_HEAD.Invoice_Type AS Invoice_Type,TSPL_MCC_SALE_FARMER_HEAD.Document_Code , " &
                                   " convert(varchar,TSPL_MCC_SALE_FARMER_HEAD.Document_Date,103 ) as Document_Date , TSPL_MCC_Sale_Farmer_Detail.Item_Code,TSPL_MCC_Sale_Farmer_Detail.Line_No , " &
                                   " TSPL_MCC_Sale_Farmer_Detail.Qty as Qty,TSPL_MCC_Sale_Farmer_Detail.Unit_code , " &
                                   " TSPL_MCC_Sale_Farmer_Detail.Item_Cost*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0 " &
                                   " then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as Item_Cost , " &
                                   " (TSPL_MCC_Sale_Farmer_Detail.Amount *(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) - case when TSPL_MCC_SALE_FARMER_HEAD.trans_type='FS' then coalesce(TSPL_MCC_Sale_Farmer_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) else 0 end) as AMount, " &
                                   " TSPL_MCC_Sale_Farmer_Detail.Disc_Per ,(TSPL_MCC_Sale_Farmer_Detail.Total_Disc_Amt*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) - case when TSPL_MCC_SALE_FARMER_HEAD.trans_type='FS' then coalesce(TSPL_MCC_Sale_Farmer_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) else 0 end) as Disc_Amt, " &
                                   " (case when coalesce(TSPL_MCC_Sale_Farmer_Detail.FOC_Item,0)=1 or coalesce(TSPL_MCC_Sale_Farmer_Detail.sampling,0)=1 or coalesce(TSPL_MCC_Sale_Farmer_Detail.Scheme_Item,'')='Y' then Item_Net_Amt*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) end) as [Scheme Amount] , " &
                                   " ((Amount-Total_Disc_Amt)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end)- case when TSPL_MCC_SALE_FARMER_HEAD.trans_type='AS' then coalesce(TSPL_MCC_Sale_Farmer_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) else 0 end) as Amt_Less_Discount , " &
                                   " TSPL_MCC_Sale_Farmer_Detail.Total_Tax_Amt*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as Total_Tax_AMt , " &
                                   " (Amount+TSPL_MCC_Sale_Farmer_Detail.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0 then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as Total_Amt, " &
                                   " TSPL_MCC_SALE_FARMER_HEAD.Vehicle_Code as Vehicle_Code ,TSPL_MCC_SALE_FARMER_HEAD.VEHICLENO as Vehicle_No, " &
                                   " (case when TSPL_MCC_Sale_Farmer_Detail.Line_No=1 then (TSPL_MCC_SALE_FARMER_HEAD.Total_Add_Charge+coalesce(TSPL_MCC_SALE_FARMER_HEAD.RoundOffAmount,0))*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) else 0 end) as Additional_Charge, " &
                                   " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0 then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as [AR Document Amt]," &
                                   " TSPL_Customer_Invoice_Head.Discount_Amount*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as [AR Document Discount Amt], " &
                                   " TSPL_Customer_Invoice_Head.amount_less_Discount*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as [AR Amount Before Tax]," &
                                   " TSPL_Customer_Invoice_Head.total_tax*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as [AR Total Tax]," &
                                   " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) as [AR Total Add Charge], " &
                                   " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                                   " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_MCC_SALE_FARMER_HEAD.GRNo as [GR No]," &
                                   " cast(null as date) as [GR Date],TSPL_MCC_SALE_FARMER_HEAD.WayBillNo as [WayBill No],'' as [Transporter Code]," &
                                   " '' as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No], " &
                                   " TSPL_MCC_Sale_Farmer_Detail.MRP ,TSPL_MCC_Sale_Farmer_Detail.Scheme_Code as [Scheme Code] ,case when TSPL_MCC_Sale_Farmer_Detail.scheme_item='N' then '' else TSPL_MCC_Sale_Farmer_Detail. scheme_type end as [Scheme Type]" &
                                   " TSPL_MCC_Sale_Farmer_Detail.Cash_Scheme_Code as [Cash Scheme Code] ,TSPL_MCC_Sale_Farmer_Detail.Cash_Scheme_Amount as [Cash Scheme Amount],TSPL_MCC_Sale_Farmer_Detail.Price_code as [Price Code],'' as Created_By, " &
                                   " '' as  Modify_By ,TSPL_MCC_Sale_Farmer_Detail.RATE_UOM,TSPL_MCC_Sale_Farmer_Detail.Conv_Factor,TSPL_MCC_Sale_Farmer_Detail.Sampling," &
                                   " TSPL_MCC_Sale_Farmer_Detail.Scheme_Item ,(case when TSPL_MCC_SALE_FARMER_HEAD.Is_Taxable=1 then 'Tax Invoice' else 'Bill Of Supply' end) as [Invoice Type GST]," &
                                   " ' " & CompGstinNo & "' as [GSTIN No Company],'' as [GSTIN No Customer], (case when TSPL_MCC_SALE_FARMER_HEAD.Total_Tax_Amt<=0 and TSPL_MCC_SALE_FARMER_HEAD.Tax_Group<>'EXEMPTED'  then ((Amount+TSPL_MCC_Sale_Farmer_Detail.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end)) else null end) as [Nill Rate Amount]," &
                                   " (case when TSPL_MCC_SALE_FARMER_HEAD.Tax_Group='EXEMPTED' then ((Amount+TSPL_MCC_Sale_Farmer_Detail.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end)) else null end) as [Exempted Amount]," &
                                   " 0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,TSPL_MCC_SALE_FARMER_HEAD.Sale_Invoice_No as [Shipping Bill No],convert(varchar,TSPL_MCC_SALE_FARMER_HEAD.Inv_Date,103) as [Shipping Bill Date]," &
                                   " '' as [Original Invoice No],'' as [Original Invoice Date],'' as [Reason for Revision], " &
                                   " (CASE WHEN TAXM1.TYPE='M' THEN TSPL_MCC_Sale_Farmer_Detail.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_MCC_Sale_Farmer_Detail.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_MCC_Sale_Farmer_Detail.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,'' as [Executive], "

                strSDEndQry = " from TSPL_MCC_Sale_Farmer_Detail " &
                              " left outer join TSPL_MCC_SALE_FARMER_HEAD on TSPL_MCC_SALE_FARMER_HEAD.Document_Code =TSPL_MCC_Sale_Farmer_Detail.DOCUMENT_CODE " &
                              " left join TSPL_VEHICLE_MASTER on TSPL_MCC_SALE_FARMER_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                              " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_MCC_SALE_FARMER_HEAD.Document_Code " &
                              " /*left join TSPL_DELIVERY_NOTE_MASTER_FRESHSALE on TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No=TSPL_MCC_Sale_Farmer_Detail.Delivery_Code " &
                              " left join TSPL_SD_SHIPMENT_HEAD on TSPL_MCC_SALE_FARMER_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code */ " &
                              " LEFT OUTER JOIN TSPL_MP_MASTER as TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.MP_Code = TSPL_MCC_SALE_FARMER_HEAD.Farmer_Code " &
                              " /*LEFT JOIN TSPL_EX_COMMERCIAL_INVOICE_HEAD ON TSPL_MCC_SALE_FARMER_HEAD.Against_Com_Inv_No=TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Code " &
                              " LEFT JOIN TSPL_MCC_SALE_FARMER_HEAD SupplInvoice ON TSPL_MCC_SALE_FARMER_HEAD.Invoice_No_For_Supplementary=SupplInvoice.Document_Code */ " &
                              " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_MCC_Sale_Farmer_Detail.TAX1=TAXM1.TAX_CODE " &
                              " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_MCC_Sale_Farmer_Detail.TAX2=TAXM2.TAX_CODE " &
                              " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_MCC_Sale_Farmer_Detail.TAX3=TAXM3.TAX_CODE "
                strSDJoinQry = " where (case when TSPL_MCC_SALE_FARMER_HEAD.Trans_Type='MCC' then 'MCC Sale Farmer' else  TSPL_MCC_SALE_FARMER_HEAD.Trans_Type end) " &
                              " in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                              " and  convert(date,TSPL_MCC_SALE_FARMER_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) " &
                              " and convert(date,TSPL_MCC_SALE_FARMER_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

                '' filter conditions
                If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Farmer_Detail.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
                End If
                If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_SALE_FARMER_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
                End If

                If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_SALE_FARMER_HEAD.FARMER_CODE in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
                End If
                If clsCommon.myLen(obj.Document_Code) > 0 Then
                    strSDJoinQry += " and TSPL_MCC_SALE_FARMER_HEAD.Document_Code = '" & obj.Document_Code & "' "
                End If
                If obj.Scheme_Type_List IsNot Nothing AndAlso obj.Scheme_Type_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Farmer_Detail.Scheme_Type in (" + clsCommon.GetMulcallString(obj.Scheme_Type_List) + ") "
                End If

                'strMCCMaterial = " select case when len([Form Type])>0 then [Form Type] else   _Type end  as [Form Type],[Trans Type],[Location Code],[Location Name],Loc.State as [Location State],GSTState.GST_STATE_Code as [GST State Code],loc.GSTNO as [Dispatch Location GSTIN No], (CASE WHEN [Invoice Type]='T' THEN 'Tax' when [Invoice Type]='R' then 'Retail' when [Invoice Type]='N' then 'None' else [Invoice Type] end) as [Invoice Type],[Document No],[Document_date],[Narration],Vehicle_Code as [Vehicle Code],Vehicle_No as [Vehicle No],cast(Additional_Charge as numeric(18,2)) as [Additional Amount],[Customer Code],[Customer Name],[Customer Address],Cust.Struct_Code,xx.[Registered],[Composition],[City Code],[Place of Supply],[Customer GST State Code],Cust.Cust_Group_Code as [Customer Group Code],Cust_Group.Cust_Group_Desc as [Customer Group Description],Cust.Zone_Code as [Customer Zone Code],Zone.Description as [Customer Zone Description], [Parent Customer No],[Parent Customer Code], [Parent Customer Name],coalesce(Cust.State_Code,Cust.State_Code) as [Customer State Code],coalesce(Cust.State_Name,Cust_Loc.State_Name) as [Customer State Desc],coalesce(Cust.Tin_No,cust_loc.Tin_No) as [Tin No],Item_Group.Item_Group as [Item Group Code],Item_Group.Group_Description as [Item Group Description] "
                'If clsCommon.myLen(strCategoryTable) > 0 Then
                '    ''richa agarwal to avoid ambiguous error
                '    '  strMCCMaterial += "," + strCodeColumn + "," + strCodeDescColumn
                '    strMCCMaterial += "," + strCodeColumnForVirtual + "," + strCodeDescColumn
                'End If
                '' BM00000008438 BM00000008391           
                'strMCCMaterial += " , [Item Code],[Item Name],[HSN Code],cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3)) as [Quantity]," & IIf(clsCommon.myLen(Unit_Code) <= 0, IIf(Stock_uom = True, "TransStock.UOM_Code", "xx.[UOM]"), "'" & Unit_Code & "'") & " as [UOM],"
                'If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not Stock_uom Then
                '    strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  as [Item Rate] "
                'Else
                '    strMCCMaterial += "  cast(( case when isnull(Rate_Stock_SU.Conversion_Factor,0)<=0 then ([Item Cost]) else ([Item Cost] * Rate_select_SU.Conversion_Factor)/ Rate_Stock_SU.Conversion_Factor end )"
                '    strMCCMaterial += " as Numeric(18,3)) as [Item Rate]"
                'End If

                ' '' below code was removed by panch raj as per email dated 20-sep-2017:2.34PM (previously Sales amount=total amout-total tax amount+scheme amount. Now sales amount=Total amount-total tax amount)
                ' ''+ case when coalesce([Total Tax Amount],0)>0 then 0 else coalesce( [Scheme Amount],0) end
                'strMCCMaterial += " ,case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [Fat Per] else QC.fat_per end as [FAT %],case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [SNF Per] else QC.snf_per end as [SNF %],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [Fat Per] else QC.fat_per end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as numeric(18,3)) end) as [FAT KG],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [SNF Per] else snf_per end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as Numeric(18,3)) end) as [SNF KG],Amount,[Discount Per] as [Discount %],  (coalesce( [Discount Amount],0)-coalesce([Scheme Amount],0))  as [Discount Amount],[Scheme Amount],[Amount Less Discount]  as [Amount Less Discount]" + strPivotForFinalOuterQuery + " " + strPivotForFinalOuterPercentQuery + ",case when [trans type]='Fresh Sale Return' then [Amount Less Discount]  else ([Total Amount]-[Total Tax Amount]) end as [Sale Amount],([Total Amount]-[Total Tax Amount]+MANDI_TAX_AMT) as [Sale Amount GST],[Total Tax Amount], (cast(Additional_Charge as numeric(18,2))+[Total Amount]) as [Total Amount], " & _
                '" [AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax]+ case when (coalesce([Total Tax Amount],0)<>0 or [Scheme Amount]<=0) and [Document No]<>'SRFS-003/15-16/000006' then 0 else coalesce([AR Document Discount Amt],0)  end as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge], "
                ' ''richa agarwal change to show csa sales account for csa sale and csa sale return
                ''left(Item.Sales_Account, Len(Item.Sales_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code as [Sales Account],
                'strMCCMaterial += "  case when [trans type] in ('CSA Sale','CSA Sale Return') then (case when coalesce(item.GSOC_Acct,'')<>'' then  left( item.GSOC_Acct, Len( item.GSOC_Acct)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code else '' end) else left(Item.Sales_Account, Len(Item.Sales_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code  end as [Sales Account], " & _
                '" [GR No],convert(varchar,[GR Date],103) as [GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No], [Booking No],MRP, [Scheme Code] , [Cash Scheme Code] , [Cash Scheme Amount], [Price Code], case when Sampling=0 then  'N' else case when sampling=1 then'Y' end end as sampling, Scheme_Item as [Scheme Type],[Invoice Type GST],[GSTIN No Company],[GSTIN no Customer],[Nill Rate Amount],[Exempted Amount],[Non GST Supply],[Reverse Charge],[Export Type],Port,[Shipping Bill No],[Shipping Bill Date],[Original Invoice No],[Original Invoice Date],[Reason for Revision] "


                ' ''richa agarwal add merchant trade trans_type in below qry BM00000008390 (Applied For DCC Also) 
                'strMCCMaterial += " from ( "
                ' '' base union farmer sale
                'If obj.Trans_Type_List.Contains("Fresh Sale") OrElse obj.Trans_Type_List.Contains("Product Sale") OrElse obj.Trans_Type_List.Contains("MCC Sale") OrElse obj.Trans_Type_List.Contains("Export Sale") OrElse obj.Trans_Type_List.Contains("CSA Sale") OrElse obj.Trans_Type_List.Contains("Merchant Trade") Then
                '' base union 12
                If qryStarted = True Then
                    strMCCMaterial += " union all "
                End If
                qryStarted = True
                strMCCMaterial += " select max(Gross_Amount)Gross_Amount,max(TotalSubsidyAmt)TotalSubsidyAmt,max(DCSCODE) as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(Final.[Ack No]) as [Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type , max(final.FormType) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='EX' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' when Trans_Type='MT' then 'Merchant Trade' WHEN Trans_Type ='SD' then 'General Sale' else  Trans_Type end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date], max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) AS [Customer Address] ,max(TSPL_CUSTOMER_MASTER .MP_Name) as [Customer Name],'' as [Registered],'' as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,'' as [Parent Customer No] ,'' as [Parent Customer Code],'' as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

                ''Monika  QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
                strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " &
                    " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]- case when (Trans_Type ='FS') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code] ,[Scheme Type],[Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Created_By,final.Modify_By ,final. RATE_UOM,final. Conv_Factor,final.Sampling,final.Scheme_Item, " &
                    " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,'' as [Executive]"
                strMCCMaterial += " from ("
                'strTaxColumns = " TSPL_MCC_SALE_FARMER_HEAD.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_MCC_SALE_FARMER_HEAD.TAX1_Amt as Tax1_Amt, TSPL_MCC_SALE_FARMER_HEAD.TAX1_Rate,TSPL_MCC_SALE_FARMER_HEAD.TAX1+'%' as Tax1Rate"
                strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
                '' query for no tax applied

                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & "  and (coalesce(TSPL_MCC_Sale_Farmer_Detail.tax1,'')='' and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax2,'')='' " &
                                  " and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax3,'')='' and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax4,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Farmer_Detail.tax5,'')='' and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax6,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Farmer_Detail.tax7,'')='' and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax8,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Farmer_Detail.tax9,'')='' and coalesce(TSPL_MCC_Sale_Farmer_Detail.tax10,'')='') )t "

                '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + ") " & _

                strMCCMaterial += "   union all"
                '' query for tax1 applied
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX1 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX1_Amt as Tax1_Amt, TSPL_MCC_Sale_Farmer_Detail.TAX1_Rate,TSPL_MCC_Sale_Farmer_Detail.TAX1+'%' as Tax1Rate"

                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax1 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX1_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and TSPL_MCC_Sale_Farmer_Detail.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "   union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX2 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX2_Amt as Tax2_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX2_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX2+'%' as Tax2Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax2 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX3 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX3_Amt as Tax3_Amt, TSPL_MCC_Sale_Farmer_Detail.TAX3_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX3+'%' as Tax3Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax3 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "   union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX4 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX4_Amt as Tax4_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX4_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX4+'%' as Tax4Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax4 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX5 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX5_Amt as Tax5_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX5_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX5+'%' as Tax5Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax5 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX6 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX6_Amt as Tax6_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX6_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX6+'%' as Tax6Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax6 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX7 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX7_Amt as Tax7_AMt,TSPL_MCC_Sale_Farmer_Detail.TAX7_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX7+'%' as Tax7Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax7 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX8 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX8_Amt as Tax8_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX8_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX8+'%' as Tax8Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax8 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX9 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX9_Amt as Tax9_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX9_Rate, TSPL_MCC_Sale_Farmer_Detail.TAX9+'%' as Tax9Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax9 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Farmer_Detail.TAX10 ,(case when coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_SALE_FARMER_HEAD.convrate,0) end) * TSPL_MCC_Sale_Farmer_Detail.TAX10_Amt as Tax10_Amt,TSPL_MCC_Sale_Farmer_Detail.TAX10_Rate,TSPL_MCC_Sale_Farmer_Detail.TAX10+'%' as Tax10Rate"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Farmer_Detail.tax10 and ttr.tax_Rate=TSPL_MCC_Sale_Farmer_Detail.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Farmer_Detail.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += " )final"
                strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
                strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
                strMCCMaterial += " left outer join TSPL_MP_MASTER AS TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .MP_CODE =final.Customer_Code "
                'strMCCMaterial += " LEFT OUTER JOIN TSPL_MP_MASTER as Parent_Master ON Parent_Master.MP_CODE=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
                strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                                  " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State_Code=TSPL_STATE_MASTER.STATE_CODE "


                ''Monika
                ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "

                'added by stuti on 01/05/2017
                strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

                If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
                End If
                If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
                End If
                If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
                End If

                ''end here

                strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , [Scheme Code],[Scheme Type] , [Cash Scheme Code] , [Cash Scheme Amount] , [Price Code] ,final.Created_By,final.Modify_By,final.RATE_UOM,final.Conv_Factor,final.Sampling,final.Scheme_Item" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per

            End If
            If obj.Trans_Type_List.Contains("MCC Sale Return Farmer") Then

                Dim strSDRCommonQuery As String = ""
                strSDRCommonQuery = " select (case when len(TSPL_MCC_Sale_Return_Head_Farmer.Description)<=0 then TSPL_MCC_Sale_Return_Head_Farmer.Comments else TSPL_MCC_Sale_Return_Head_Farmer.Description end) as Narration,'' as Formtype,'MCC Sale Return Farmer' as Trans_Type,TSPL_MCC_Sale_Return_Head_Farmer.Status ,TSPL_MCC_Sale_Return_Head_Farmer.Bill_To_Location, " &
                                  " TSPL_MCC_Sale_Return_Head_Farmer.Farmer_Code as Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2  As CustAdd,COALESCE(TSPL_MCC_Sale_Return_Head_Farmer.Document_Type,TSPL_MCC_Sale_Return_Head_Farmer.Invoice_Type) AS Invoice_Type,TSPL_MCC_Sale_Return_Head_Farmer.Document_Code , " &
                                  " convert(varchar,TSPL_MCC_Sale_Return_Head_Farmer.Document_Date,103 ) as Document_Date , TSPL_MCC_Sale_Return_Detail_Farmer.Item_Code,TSPL_MCC_Sale_Return_Detail_Farmer.Line_No , " &
                                  " -TSPL_MCC_Sale_Return_Detail_Farmer.Qty as Qty ,TSPL_MCC_Sale_Return_Detail_Farmer.Unit_code ,TSPL_MCC_Sale_Return_Detail_Farmer.Item_Cost , " &
                                  " -coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Amount,0) as Amount ,TSPL_MCC_Sale_Return_Detail_Farmer.Disc_Per ,case when coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0)=0 then -coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0)  + case when coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.FOC_Item,0)=1 or coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.sampling,0)=1  then 1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0) end) else 0 end else -coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0) end as Disc_Amt,case when coalesce(FOC_Item,0)=1 or coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.sampling,0)=1  then -1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0) end) end  as [Scheme Amount] , " &
                                  " -(Amount- case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type='FS' then Total_Disc_Amt else Total_Disc_Amt end  + case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type<>'FS' then case when coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.FOC_Item,0)=1 or coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.sampling,0)=1  then Item_Net_Amt*(case when coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0)<=0  then 1 else coalesce(TSPL_MCC_Sale_Return_Head_Farmer.convrate,0) end) else 0 end else 0 end)  as Amt_Less_Discount , " &
                                  " -coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Tax_Amt,0) as Total_Tax_Amt ,-(Amount+coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Tax_Amt,0)- case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type='FS' then TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt else coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0) end )  as Total_Amt,TSPL_MCC_Sale_Return_Head_Farmer.Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,-(case when TSPL_MCC_Sale_Return_Detail_Farmer.Line_No=1 then (coalesce(TSPL_MCC_Sale_Return_Head_Farmer.Total_Add_Charge,0)+coalesce(TSPL_MCC_Sale_Return_Head_Farmer.RoundOffAmount,0)) else 0 end) as Additional_Charge, " &
                                  " '' as [AR Document No],0 as [AR Document Amt]," &
                                  " 0 as [AR Document Discount Amt], " &
                                  " 0 as [AR Amount Before Tax],0 as [AR Total Tax], " &
                                  " 0 as [AR Total Add Charge], " &
                                  " '' as Against_Sale_No,'' as Against_Sale_Return_No,'' as AgainstScrap, " &
                                  " '' as Against_VCGL,'' as Against_MCC_Material_Sale_Return,TSPL_MCC_Sale_Return_Head_Farmer.GRNo as [GR No],cast(null as date) as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],TSPL_MCC_Sale_Return_Detail_Farmer.Delivery_Code as [Delivery No],'' as [Shipment No],TSPL_MCC_Sale_Return_Detail_Farmer.MRP, TSPL_MCC_Sale_Return_Detail_Farmer.Scheme_Code ,TSPL_MCC_Sale_Return_Detail_Farmer.Cash_Scheme_Code ,TSPL_MCC_Sale_Return_Detail_Farmer.Cash_Scheme_Amount*(-1) as Cash_Scheme_Amount ,TSPL_MCC_Sale_Return_Detail_Farmer.Price_code ,'' as Created_By ,'' as Modify_By,TSPL_MCC_Sale_Return_Detail_Farmer.Unit_code as RATE_UOM,0 as Conv_Factor, TSPL_MCC_Sale_Return_Detail_Farmer.Sampling,TSPL_MCC_Sale_Return_Detail_Farmer.Scheme_Item," &
                                  " 'Credit Note' as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],'' as [GSTIN no Customer],(case when TSPL_MCC_Sale_Return_Head_Farmer.Total_Tax_Amt<=0 and TSPL_MCC_Sale_Return_Head_Farmer.Tax_Group<>'EXEMPTED' then -(Amount+coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Tax_Amt,0)- case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type='FS' then 0 else coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0) end ) else null end) as [Nill Rate Amount],(case when TSPL_MCC_Sale_Return_Head_Farmer.Tax_Group='EXEMPTED' then -(Amount+coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Tax_Amt,0)- case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type='FS' then 0 else coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.Total_Disc_Amt,0) end ) else null end) as [Exempted Amount],0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port,'' as [Shipping Bill No],'' as [Shipping Bill Date],TSPL_MCC_Sale_Return_Detail_Farmer.Invoice_Code as [Original Invoice No],convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as [Original Invoice Date],TSPL_MCC_Sale_Return_Head_Farmer.Description as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_MCC_Sale_Return_Detail_Farmer.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_MCC_Sale_Return_Detail_Farmer.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_MCC_Sale_Return_Detail_Farmer.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT,'' as [Executive],"
                strSDEndQry = " from TSPL_MCC_Sale_Return_Detail_Farmer " &
                                    " left outer join TSPL_MCC_Sale_Return_Head_Farmer on TSPL_MCC_Sale_Return_Head_Farmer.Document_Code =TSPL_MCC_Sale_Return_Detail_Farmer.DOCUMENT_CODE " &
                                    " left join TSPL_VEHICLE_MASTER on TSPL_MCC_Sale_Return_Head_Farmer.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                                    " left join TSPL_MCC_SALE_FARMER_HEAD as TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_MCC_Sale_Return_Detail_Farmer.Invoice_Code " &
                                    " LEFT OUTER JOIN TSPL_MP_MASTER as TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.MP_Code = TSPL_MCC_Sale_Return_Head_Farmer.Farmer_Code " &
                                    " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_MCC_Sale_Return_Detail_Farmer.TAX1=TAXM1.TAX_CODE " &
                                    " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_MCC_Sale_Return_Detail_Farmer.TAX2=TAXM2.TAX_CODE " &
                                    " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_MCC_Sale_Return_Detail_Farmer.TAX3=TAXM3.TAX_CODE "

                strSDJoinQry = " WHERE 2=2 AND (case when TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type='MCC' then 'MCC Sale Return Farmer' else  TSPL_MCC_Sale_Return_Head_Farmer.Trans_Type end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                                " and convert(date,TSPL_MCC_Sale_Return_Head_Farmer.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_MCC_Sale_Return_Head_Farmer.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

                '' filter conditions
                If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Return_Detail_Farmer.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
                End If
                If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Return_Head_Farmer.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
                End If

                If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Return_Head_Farmer.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
                End If
                If clsCommon.myLen(obj.Document_Code) > 0 Then
                    strSDJoinQry += " and TSPL_MCC_Sale_Return_Head_Farmer.Document_Code = '" & obj.Document_Code & "' "
                End If

                '' base union 5
                If qryStarted = True Then
                    strMCCMaterial += " union all "
                End If
                qryStarted = True
                strMCCMaterial += " select 0 as Gross_Amount,0 as TotalSubsidyAmt,'' as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(final.[Ack No]) as [Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type, max(final.Formtype) as [Form Type],case when Trans_Type ='DSR' then 'Dairy Sale Return' when Trans_Type ='FSR' then 'Fresh Sale Return' when Trans_Type ='CSAR' then 'CSA Sale Return' when Trans_Type='PSR' then 'Product Sale Return' when Trans_Type='MCCR' then 'MCC Sale Return' when Trans_Type='EXPR' then 'Export Sale Return'when Trans_Type='Bulk Sale' then 'Bulk Sale Return' when Trans_Type ='SSR' then 'Misc Sale' when Trans_Type ='SDR' then 'General Sale Return' else trans_Type  end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost],"

                ''Monika QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
                strMCCMaterial += "  0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount], " &
                    " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]  - case when (Trans_Type ='DSR' or Trans_Type ='FSR' or Trans_Type ='PSR') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],final.[Transporter Name], [Delivery No]  , [Shipment No],MRP, Scheme_Code as [Scheme Code] ,Scheme_Type as [Scheme Type],Cash_Scheme_Code as [Cash Scheme Code], Cash_Scheme_Amount as [Cash Scheme Amount] ,final.Price_code as [Price Code] ,final.Created_By ,final.Modify_By,final.RATE_UOM ,final.Conv_Factor ,final. Sampling,final.Scheme_Item," &
                    " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max(TSPL_EMPLOYEE_MASTER.Emp_Name) as [Executive] " &
                    " from ( "

                strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
                '' query for no tax applied
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax1,'')='' and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax2,'')='' " &
                                  " and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax3,'')='' and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax4,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax5,'')='' and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax6,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax7,'')='' and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax8,'')='' and " &
                                  " coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax9,'')='' and coalesce(TSPL_MCC_Sale_Return_Detail_Farmer.tax10,'')='') )t "

                strMCCMaterial += "   union all"
                '' query for tax1 applied
                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX1 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX1_Amt as TAX1_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX1_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX1+'%' as tax1rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax1 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX1_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

                strMCCMaterial += "   union all"
                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX2 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX2_Amt as TAX2_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX2_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX2+'%' as tax2rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax2 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX2_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX3 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX3_Amt as TAX3_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX3_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX3+'%' as tax3rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax3 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX3_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax7 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX7_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "   union all"
                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX4 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX4_Amt as TAX4_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX4_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX4+'%' as tax4rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax4 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX4_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"
                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX5 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX5_Amt as TAX5_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX5_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX5+'%' as tax5rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax5 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX5_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"

                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"

                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX6 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX6_Amt as TAX6_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX6_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX6+'%' as tax6rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax6 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX6_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"

                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX7 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX7_Amt as TAX7_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX7_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX7+'%' as tax7rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax7 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX7_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"

                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX8 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX8_Amt as TAX8_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX8_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX8+'%' as tax8rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax8 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX8_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"

                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX9 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX9_Amt as TAX9_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX9_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX9+'%' as tax9rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax9 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX9_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += "  union all"

                strTaxColumns = " TSPL_MCC_Sale_Return_Detail_Farmer.TAX10 ,-TSPL_MCC_Sale_Return_Detail_Farmer.TAX10_Amt as TAX10_Amt,TSPL_MCC_Sale_Return_Detail_Farmer.TAX10_Rate ,TSPL_MCC_Sale_Return_Detail_Farmer.TAX10+'%' as tax10rate  "
                strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_MCC_Sale_Return_Detail_Farmer.tax10 and ttr.tax_Rate=TSPL_MCC_Sale_Return_Detail_Farmer.TAX10_Rate and ttr._type<>'OH' and ttr._type<>'OH' and ttr.Tax_Type='S'"
                strMCCMaterial += " select * from (" & strSDRCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_MCC_Sale_Return_Detail_Farmer.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
                strMCCMaterial += " )final"
                strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
                strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
                strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
                strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
                strMCCMaterial += " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code"
                strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                                  " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "

                ''Monika
                ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
                'added by stuti on 01/05/2017
                strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

                ''=========Monika
                If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
                End If
                If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
                End If
                If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                    strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
                End If
                ''======end here

                If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                    strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
                End If
                If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                    strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
                End If

                strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , Scheme_Code ,Cash_Scheme_Code , Cash_Scheme_Amount ,final.Price_code ,final.Created_By ,final.Modify_By ,final.RATE_UOM,final.Conv_Factor ,final. Sampling,final.Scheme_Item " '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per

            End If
        End If
        '' Debit/Credit Note
        If obj.Include_Debit_Credit Then
            Dim strDebitCreditCommonQuery As String = ""
            strDebitCreditCommonQuery = " select TSPL_Customer_Invoice_Head.Description as Narration,'' as Formtype, 'Debit/Credit' as Trans_Type  ,TSPL_Customer_Invoice_Head.Status ,TSPL_Customer_Invoice_Head.Loc_Code as Bill_To_Location, " &
                                        " TSPL_Customer_Invoice_Head.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd," &
                                        " TSPL_Customer_Invoice_Head.Document_Type AS Invoice_Type,TSPL_Customer_Invoice_Head.Document_No as Document_Code , " &
                                        " convert(varchar,TSPL_Customer_Invoice_Head.Document_Date,103 ) as Document_Date , '' as Item_Code,TSPL_Customer_Invoice_detail.SNo as Line_No, " &
                                        " 0 as Qty,'' as Unit_code , 0 as Item_Cost ,TSPL_Customer_Invoice_detail.Amount* (case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end) *(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as AMount, " &
                                        " TSPL_Customer_Invoice_detail.Discount_Per as Disc_Per ,TSPL_Customer_Invoice_detail.Discount*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as Disc_Amt, 0 as [Scheme Amount] , " &
                                        " ((TSPL_Customer_Invoice_detail.Amount-TSPL_Customer_Invoice_detail.Discount)*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end)) as Amt_Less_Discount , " &
                                        " TSPL_Customer_Invoice_detail.Total_Tax*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as Total_Tax_AMt ," &
                                        " (TSPL_Customer_Invoice_detail.Amount+TSPL_Customer_Invoice_detail.Total_Tax-TSPL_Customer_Invoice_detail.Discount)*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as Total_Amt, '' as Vehicle_Code ,'' as Vehicle_No, " &
                                        " (case when TSPL_Customer_Invoice_detail.SNo=1 then (TSPL_Customer_Invoice_Head.Total_Add_Charge+coalesce(TSPL_Customer_Invoice_Head.RoundOffAmount,0))*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) else 0 end) as  Additional_Charge, " &
                                        " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as [AR Document Amt]," &
                                        " TSPL_Customer_Invoice_Head.Discount_Amount*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end)  as [AR Document Discount Amt], " &
                                        " TSPL_Customer_Invoice_Head.amount_less_Discount*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as [AR Amount Before Tax]," &
                                        " TSPL_Customer_Invoice_Head.total_tax*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as [AR Total Tax], " &
                                        " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount)*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end) as [AR Total Add Charge], " &
                                        " '' as Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,'' as AgainstScrap,'' as Against_VCGL,'' as Against_MCC_Material_Sale_Return,'' as [GR No],'' as [GR Date],'' as [WayBill No],'' as [Transporter Code]," &
                                        " '' as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No], 0 as MRP ,'' as [Scheme Code],'' as Scheme_Type ,0 as [Cash Scheme Code] ,0 as [Cash Scheme Amount],'' as [Price Code],TSPL_Customer_Invoice_Head.Created_By as Created_By,TSPL_Customer_Invoice_Head.Modify_By as Modify_By ," &
                                        " '' as RATE_UOM,0 as Conv_Factor,'' as Sampling,'' as Scheme_Item , (case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'Debit Note' when TSPL_Customer_Invoice_Head.Document_Type='C' then 'Credit Note' else 'Invoice' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN No Customer], " &
                                        " (case when TSPL_Customer_Invoice_Head.Total_Tax<=0 and TSPL_Customer_Invoice_Head.Tax_Group<>'EXEMPTED' then ((Amount+TSPL_Customer_Invoice_detail.Total_Tax-TSPL_Customer_Invoice_detail.Discount)*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end)) else null end) as [Nill Rate Amount]," &
                                        " (case when TSPL_Customer_Invoice_Head.Tax_Group='EXEMPTED' then ((Amount+TSPL_Customer_Invoice_detail.Total_Tax-TSPL_Customer_Invoice_detail.Discount)*(case when TSPL_Customer_Invoice_Head.Document_Type='D' then 1 else -1 end)*(case when coalesce(TSPL_Customer_Invoice_Head.convrate,0)<=0  then 1 else coalesce(TSPL_Customer_Invoice_Head.convrate,0) end)) else null end) as [Exempted Amount], " &
                                        " 0 as [Non GST Supply],'N' as [Reverse Charge],'' as [Export Type],'' as Port, " &
                                        " '' as [Shipping Bill No],'' as [Shipping Bill Date],TSPL_Customer_Invoice_Head.RefDocNo as [Original Invoice No],convert(varchar,RefInvoice.Document_Date,103) as [Original Invoice Date],'' as [Reason for Revision],0 AS MANDI_TAX_AMT,isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive], "

            strSDEndQry = " from TSPL_Customer_Invoice_detail " &
                          " left outer join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Document_No =TSPL_Customer_Invoice_detail.Document_No " &
                          " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_Customer_Invoice_Head.Customer_Code" &
                        " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
                         " left join TSPL_Customer_Invoice_Head RefInvoice on TSPL_Customer_Invoice_Head.RefDocNo=RefInvoice.Document_No "

            strSDJoinQry = " where TSPL_Customer_Invoice_Head.Document_Type in ('D','C')  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_Asset_Disposal,''))<=0 " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,''))<=0  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_Sale_No,''))<=0" &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_Sale_Return_No,''))<=0  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_Security_Receipt_No,''))<=0  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.Against_Service_Visit_Code,''))<=0  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.AgainstScrap,''))<=0  " &
                           " and len(coalesce(TSPL_Customer_Invoice_Head.AgainstScrapReturn,''))<=0 " &
                           " and convert(date,TSPL_Customer_Invoice_Head.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_Customer_Invoice_Head.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "
            '' filter conditions
            'If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            '    strSDJoinQry += " and TSPL_TRANSFER_ORDER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            'End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_Customer_Invoice_Head.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and  TSPL_Customer_Invoice_Head.Customer_Code  in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and  tspl_customer_master.cust_category_code  in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If

            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_Customer_Invoice_Head.Document_No = '" & obj.Document_Code & "' "
            End If
            '' base union 11
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True

            strMCCMaterial += " select  max(Gross_Amount)Gross_Amount,max(TotalSubsidyAmt)TotalSubsidyAmt,max(dcscode) as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(Final.[Ack No]) as [Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type , max(final.FormType) as [Form Type],case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='EX' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' when Trans_Type='MT' then 'Merchant Trade' WHEN Trans_Type ='SD' then 'General Sale' else  Trans_Type end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date], max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) AS [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

            ''Monika  QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]- case when (Trans_Type ='FS') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code],[Scheme Type] ,[Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Created_By,final.Modify_By ,final. RATE_UOM,final. Conv_Factor,final.Sampling,final.Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT,max([Executive]) as [Executive] " &
                " from ( "
            'strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,0 as TAX1_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "

            '' query for no tax applied
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & " and (coalesce(TSPL_Customer_Invoice_Detail.tax1,'')='' and coalesce(TSPL_Customer_Invoice_Detail.tax2,'')='' " &
                              " and coalesce(TSPL_Customer_Invoice_Detail.tax3,'')='' and coalesce(TSPL_Customer_Invoice_Detail.tax4,'')='' and " &
                              " coalesce(TSPL_Customer_Invoice_Detail.tax5,'')='' and coalesce(TSPL_Customer_Invoice_Detail.tax6,'')='' and " &
                              " coalesce(TSPL_Customer_Invoice_Detail.tax7,'')='' and coalesce(TSPL_Customer_Invoice_Detail.tax8,'')='' and " &
                              " coalesce(TSPL_Customer_Invoice_Detail.tax9,'')='' and coalesce(TSPL_Customer_Invoice_Detail.tax10,'')='') )t "

            strMCCMaterial += Environment.NewLine + " union all "
            '' quert for no tax applied
            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX1 ,-TSPL_Customer_Invoice_Detail.TAX1_AMT AS TAX1_AMT ,TSPL_Customer_Invoice_Detail.TAX1_Rate ,TSPL_Customer_Invoice_Detail.TAX1+'%' as tax1rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax1 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX1_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX2 ,-TSPL_Customer_Invoice_Detail.TAX2_Amt AS TAX2_Amt ,TSPL_Customer_Invoice_Detail.TAX2_Rate ,TSPL_Customer_Invoice_Detail.TAX2+'%' as tax2rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax2 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX3 ,-TSPL_Customer_Invoice_Detail.TAX3_Amt AS TAX3_Amt,TSPL_Customer_Invoice_Detail.TAX3_Rate ,TSPL_Customer_Invoice_Detail.TAX3+'%' as tax3rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax3 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX4 ,-TSPL_Customer_Invoice_Detail.TAX4_Amt AS TAX4_Amt,TSPL_Customer_Invoice_Detail.TAX4_Rate ,TSPL_Customer_Invoice_Detail.TAX4+'%' as tax4rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax4 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX5 ,-TSPL_Customer_Invoice_Detail.TAX5_Amt AS TAX5_Amt,TSPL_Customer_Invoice_Detail.TAX5_Rate ,TSPL_Customer_Invoice_Detail.TAX5+'%' as tax5rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax5 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"

            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX6 ,-TSPL_Customer_Invoice_Detail.TAX6_Amt AS TAX6_Amt,TSPL_Customer_Invoice_Detail.TAX6_Rate ,TSPL_Customer_Invoice_Detail.TAX6+'%' as tax6rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax6 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX7 ,-TSPL_Customer_Invoice_Detail.TAX7_Amt AS TAX7_Amt,TSPL_Customer_Invoice_Detail.TAX7_Rate ,TSPL_Customer_Invoice_Detail.TAX7+'%' as tax7rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax7 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX8 ,-TSPL_Customer_Invoice_Detail.TAX8_Amt AS TAX8_Amt,TSPL_Customer_Invoice_Detail.TAX8_Rate ,TSPL_Customer_Invoice_Detail.TAX8+'%' as tax8rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax8 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX9 ,-TSPL_Customer_Invoice_Detail.TAX9_Amt AS TAX9_Amt,TSPL_Customer_Invoice_Detail.TAX9_Rate ,TSPL_Customer_Invoice_Detail.TAX9+'%' as tax9rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax9 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"

            strTaxColumns = " TSPL_Customer_Invoice_Detail.TAX10 ,-TSPL_Customer_Invoice_Detail.TAX10_Amt AS TAX10_Amt,TSPL_Customer_Invoice_Detail.TAX10_Rate ,TSPL_Customer_Invoice_Detail.TAX10+'%' as tax10rate  "
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_Customer_Invoice_Detail.tax10 and ttr.tax_Rate=TSPL_Customer_Invoice_Detail.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='T'"
            strMCCMaterial += " select * from (" & strDebitCreditCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_Customer_Invoice_Detail.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t" &
             " )final" &
             " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location " &
             " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code " &
             " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =final.Customer_Code " &
             " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No " &
            " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
             " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
             " left join TSPL_STATE_MASTER on TSPL_LOCATION_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "

            strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Scheme Type], [Cash Scheme Code] , [Cash Scheme Amount] , [Price Code] ,final.Created_By,final.Modify_By,final.RATE_UOM,final.Conv_Factor,final.Sampling,final.Scheme_Item" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per

        End If



        ''ERO/18/05/18-000317 richa show delivery no and booking no in respective of dairy sale and dair sale return.
        If obj.Trans_Type_List.Contains("Dairy Sale") OrElse obj.Trans_Type_List.Contains("Fresh Sale") OrElse obj.Trans_Type_List.Contains("Product Sale") OrElse obj.Trans_Type_List.Contains("MCC Sale") OrElse obj.Trans_Type_List.Contains("Export Sale") OrElse obj.Trans_Type_List.Contains("CSA Sale") OrElse obj.Trans_Type_List.Contains("Merchant Trade") Then
            strSDCommonQuery = " select CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Document_Code = LEAD(TSPL_SD_SALE_INVOICE_HEAD.Document_Code) OVER (ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code) THEN 0 
                                    ELSE TSPL_SD_SALE_INVOICE_HEAD.Gross_Amount 
                                    END AS Gross_Amount,CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Document_Code = LEAD(TSPL_SD_SALE_INVOICE_HEAD.Document_Code) OVER (ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code) THEN 0 
                                    ELSE TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt
                                    END AS TotalSubsidyAmt ,tspl_vlc_master_Head.VLC_Code_VLC_Uploader AS DCSCODE,  (TSPL_SD_SHIPMENT_HEAD.Supply_Date)Supply_Date,
                                    CASE 
                                    WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type = 'FS' 
                                        THEN TSPL_BOOKING_MATSER.Payment_Terms
                                    WHEN TSPL_SD_SHIPMENT_HEAD.Is_CashSale = 'Y' 
                                        THEN 'CASH'
                                    ELSE 'CREDIT'
                                END AS Payment_Terms, Case when TSPL_BOOKING_MATSER.Is_BPL = 1 then 'BPL'
						  WHEN TSPL_BOOKING_MATSER.Is_Distributor = 1 THEN 'Distributor'
						  WHEN TSPL_BOOKING_MATSER.Is_DCS = 1 THEN 'DCS'
						  WHEN TSPL_BOOKING_MATSER.Is_CashSale= 'Y' THEN 'CASH SALE' ELSE 'OTHER' end as CustomerType,
                            Case when TSPL_BOOKING_MATSER.Is_CashSale= 'Y' THEN 'CASH' ELSE 'CREDIT' end as PaymentType,
						  --Case when TSPL_SD_SHIPMENT_HEAD.Is_CashSale = 'Y' then 'CASH' Else 'CREDIT' end as PaymentType, 
                          TSPL_SD_SALE_INVOICE_HEAD.Ack_No as [Ack No],TSPL_SD_SALE_INVOICE_HEAD.Ack_Date as [ACK Date],TSPL_SD_SALE_INVOICE_HEAD.IRN_No as [IRN No],TSPL_SD_SALE_INVOICE_HEAD.Screen_Type,TSPL_SD_SALE_INVOICE_HEAD.Description as Narration,case when TSPL_SD_SALE_INVOICE_HEAD.Against_C_Form = 1 then 'C' else '' End as Formtype,case when ISNUll(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,'')<>'' then TSPL_SD_SALE_INVOICE_HEAD.Document_Type else  CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='ALL' THEN 'SD' ELSE TSPL_SD_SALE_INVOICE_HEAD.Trans_Type END end as Trans_Type  ,TSPL_SD_SALE_INVOICE_HEAD.Status ,TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location, " &
                           " TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,TSPL_SD_SALE_INVOICE_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_INVOICE_HEAD.Document_Code , " &
                           " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_INVOICE_DETAIL.Item_Code," & IIf(Batch_Wise = True, "TSPL_BATCH_ITEM.Batch_No ,", " ") & " TSPL_SD_SALE_INVOICE_DETAIL.Line_No , " &
                           " (case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 then 0 else " & IIf(Batch_Wise = True, " TSPL_BATCH_ITEM.Qty ", " TSPL_SD_SALE_INVOICE_DETAIL.Qty  ") & " end) as Qty, " &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Unit_code ,
                           case when tspl_sd_sale_invoice_detail.Sampling=1 then 0  else TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) end as Item_Cost , " &
                           " ((TSPL_SD_SALE_INVOICE_DETAIL.Amount *(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_INVOICE_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as AMount, " &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Disc_Per ,(TSPL_SD_SALE_INVOICE_detail.Total_Disc_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) - case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Disc_Amt, " &
                           " (case when coalesce(TSPL_SD_SALE_INVOICE_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.sampling,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item,'')='Y' then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [Scheme Amount] , " &
                           " ((TSPL_SD_SALE_INVOICE_DETAIL.Amount-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='AS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Amt_Less_Discount , " &
                           " (TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_INVOICE_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as Total_Tax_AMt , " &
                           " (TSPL_SD_SALE_INVOICE_DETAIL.Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Total_Amt, " &
                           " case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then ''  when TSPL_SD_SHIPMENT_HEAD.ManualVehicle <> '' then '' else TSPL_SD_SALE_INVOICE_HEAD.Vehicle_Code end as Vehicle_Code " &
                           ", case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then TSPL_SD_SALE_INVOICE_HEAD.VehicleNo  when TSPL_SD_SHIPMENT_HEAD.ManualVehicle <> '' then TSPL_SD_SHIPMENT_HEAD.ManualVehicle else COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_SD_SALE_INVOICE_HEAD.VEHICLENO) end as Vehicle_No, " &
                           " ((case when TSPL_SD_SALE_INVOICE_DETAIL.Line_No=1 then (TSPL_SD_SALE_INVOICE_HEAD.Total_Add_Charge+coalesce(TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount,0))*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) " & IIf(Batch_Wise = True, " / TSPL_SD_SALE_INVOICE_DETAIL.Qty * TSPL_BATCH_ITEM.Qty ", "  ") & " as  Additional_Charge, " &
                           " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Document Amt]," &
                           " TSPL_Customer_Invoice_Head.Discount_Amount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Document Discount Amt], " &
                           " TSPL_Customer_Invoice_Head.amount_less_Discount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Amount Before Tax]," &
                           " TSPL_Customer_Invoice_Head.total_tax*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Total Tax], " &
                           " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [AR Total Add Charge], " &
                           " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                           " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_INVOICE_HEAD.GRNo as [GR No],tspl_sd_shipment_head.gr_date as [GR Date],TSPL_SD_SALE_INVOICE_HEAD.WayBillNo as [WayBill No],TSPL_SD_SHIPMENT_HEAD.transport_id as [Transporter Code],case when len(TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual) > 0 then TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual else TSPL_SD_SHIPMENT_HEAD.Transporter_Name end as [Transporter Name],(case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS when isnull(TSPL_SD_SALE_INVOICE_HEAD.Screen_type,'')='DS'  then TSPL_SD_SHIPMENT_HEAD.Against_Delivery_code else  TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code end) as [Delivery No]  ,Shipment_Code as [Shipment No],
                           
                           TSPL_SD_SALE_INVOICE_DETAIL.MRP ,TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Code as [Scheme Code] ,case when TSPL_SD_SALE_INVOICE_DETAIL.scheme_item='N' then '' else TSPL_SD_SALE_INVOICE_DETAIL. scheme_type end  as [Scheme Type],TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Code as [Cash Scheme Code] ," &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Amount*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as [Cash Scheme Amount], TSPL_SD_SALE_INVOICE_DETAIL.Price_code as [Price Code],'' as Created_By,'' as Modify_By ,TSPL_SD_SALE_INVOICE_DETAIL.RATE_UOM,TSPL_SD_SALE_INVOICE_DETAIL.Conv_Factor,tspl_sd_sale_invoice_detail.Sampling,tspl_sd_sale_invoice_detail.Scheme_Item ," &
                           " (case when TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='S' then 'Supplementry Invoice' when TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then 'Credit Note' when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=1 then 'Tax Invoice' else 'Bill Of Supply' end) as [Invoice Type GST],'" & CompGstinNo & "' as [GSTIN No Company],TSPL_CUSTOMER_MASTER.GSTNO as [GSTIN no Customer], " &
                           " (case when TSPL_SD_SALE_INVOICE_HEAD.Total_Tax_Amt<=0 and TSPL_SD_SALE_INVOICE_HEAD.Tax_Group<>'EXEMPTED' then ((TSPL_SD_SALE_INVOICE_DETAIL.Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) else null end) as [Nill Rate Amount]," &
                           " (case when TSPL_SD_SALE_INVOICE_HEAD.Tax_Group='EXEMPTED' then ((TSPL_SD_SALE_INVOICE_DETAIL.Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end)) else null end) as [Exempted Amount]," &
                           " 0 as [Non GST Supply],'N' as [Reverse Charge],coalesce(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,'') as [Export Type],TSPL_SD_SALE_INVOICE_HEAD.Loading_Port as Port," &
                           " TSPL_SD_SALE_INVOICE_HEAD.Against_Com_Inv_No as [Shipping Bill No],convert(varchar,TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Date,103) as [Shipping Bill Date]," &
                           " TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary as [Original Invoice No],convert(varchar,SupplInvoice.Document_Date,103) as [Original Invoice Date],SupplInvoice.Description as [Reason for Revision],(CASE WHEN TAXM1.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_AMT ELSE 0 END+CASE WHEN TAXM2.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_AMT ELSE 0 END+CASE WHEN TAXM3.TYPE='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_AMT ELSE 0 END) AS MANDI_TAX_AMT," &
                           " isnull(TSPL_EMPLOYEE_MASTER.Emp_Name,'') as [Executive],"
            strSDEndQry = " from TSPL_SD_SALE_INVOICE_DETAIL " &
                               " left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE " &
                               " left join TSPL_VEHICLE_MASTER on TSPL_SD_SALE_INVOICE_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                               " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code " &
                               "  " &
                               " left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code " &
                               " LEFT OUTER JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No  " &
                               " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code " &
                               " left outer join tspl_Customer_vendor_mapping on tspl_Customer_vendor_mapping.Cust_code=TSPL_CUSTOMER_MASTER.cust_code 
              left outer join tspl_vlc_master_Head on tspl_vlc_master_Head.VSP_Code=tspl_Customer_vendor_mapping.vendor_code " &
                                " left join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE= TSPL_CUSTOMER_MASTER.Service_Dealer_Code" &
                                " LEFT JOIN TSPL_EX_COMMERCIAL_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Against_Com_Inv_No=TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Code " &
                               " LEFT JOIN TSPL_SD_SALE_INVOICE_HEAD SupplInvoice ON TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary=SupplInvoice.Document_Code " &
                               " LEFT JOIN TSPL_TAX_MASTER TAXM1 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX1=TAXM1.TAX_CODE " &
                               " LEFT JOIN TSPL_TAX_MASTER TAXM2 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX2=TAXM2.TAX_CODE " &
                               " LEFT JOIN TSPL_TAX_MASTER TAXM3 ON TSPL_SD_SALE_INVOICE_DETAIL.TAX3=TAXM3.TAX_CODE "
            If Batch_Wise = True Then
                strSDEndQry += " left outer join  TSPL_BATCH_ITEM on  TSPL_BATCH_ITEM.document_code=TSPL_SD_SHIPMENT_HEAD.Document_Code and TSPL_BATCH_ITEM.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and TSPL_BATCH_ITEM.UOM=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code "
            End If

            strSDJoinQry = "  where  isnull(TSPL_SD_SALE_INVOICE_HEAD.trans_type,'ALL') not in ('CSA')  " &
                               " and (case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type  IN ('FS','PS') and TSPL_SD_SALE_INVOICE_HEAD.Screen_Type='DS' then 'Dairy Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='FS' and TSPL_SD_SALE_INVOICE_HEAD.Screen_Type='' then 'Fresh Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='PS' and TSPL_SD_SALE_INVOICE_HEAD.Screen_Type='' then 'Product Sale' " &
                               " when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='MCC' then 'MCC Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type <>'MT' then 'Export Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type ='MT' then 'Merchant Trade' WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='SD' then 'General Sale' " &
                               " else  TSPL_SD_SALE_INVOICE_HEAD.Trans_Type end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") "
            If obj.supply_Date Then
                strSDJoinQry += " and  convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= convert(date,('" & To_Date & "'),103) "
            Else
                strSDJoinQry += " and  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "
            End If

            '   " and  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

            '' filter conditions
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and tspl_customer_master.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strSDJoinQry += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_INVOICE_HEAD.Document_Code = '" & obj.Document_Code & "' "
            End If
            If obj.Scheme_Type_List IsNot Nothing AndAlso obj.Scheme_Type_List.Count > 0 Then
                strSDJoinQry += " and TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Type in (" + clsCommon.GetMulcallString(obj.Scheme_Type_List) + ") "
            End If

            '' base union 12
            If qryStarted = True Then
                strMCCMaterial += " union all "
            End If
            qryStarted = True
            strMCCMaterial += " select  max(Gross_Amount)Gross_Amount,max(TotalSubsidyAmt)TotalSubsidyAmt, max(dcscode) as dcscode,max(Supply_Date)Supply_Date,max(Payment_Terms)Payment_Terms,max(CustomerType)CustomerType,max(PaymentType)PaymentType,max(Final.[Ack No]) as [Ack No],max(final.[ACK Date])[ACK Date],max(final.[IRN No])[IRN No],max(final._Type) as _Type , max(final.FormType) as [Form Type],case when Trans_Type  IN ('FS','PS') and Screen_Type='DS' then 'Dairy Sale' when Trans_Type ='FS' and Screen_Type='' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' and Screen_Type='' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='EX' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' when Trans_Type='MT' then 'Merchant Trade' WHEN Trans_Type ='SD' then 'General Sale' else  Trans_Type end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date], max(final.Narration) as [Narration],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) AS [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name],max(TSPL_CUSTOMER_MASTER .GST_Registered) as [Registered],max(TSPL_CUSTOMER_MASTER .GST_COMPOSITION) as [Composition],max(TSPL_CUSTOMER_MASTER .City_Code) as [City Code],max(TSPL_CITY_MASTER .City_Name) as [Place of Supply],max(TSPL_STATE_MASTER.GST_STATE_Code) as [Customer GST State Code] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],max(tspl_item_master.HSN_Code) as [HSN Code],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], "

            ''Monika  QC.FAT_Per as [Fat Per],QC.SNF_Per as [SNF Per]
            strMCCMaterial += " 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount]," & IIf(Batch_Wise = True, " final.Amount + final.Total_Tax_AMt ", " final.Total_Amt ") & "  as [Total Amount],   " &
                " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]- case when (Trans_Type ='FS') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No] ,MRP, [Scheme Code] ,[Scheme Type],[Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Created_By,final.Modify_By ,final. RATE_UOM,final. Conv_Factor,final.Sampling,final.Scheme_Item, " &
                " max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],max([Nill Rate Amount]) as [Nill Rate Amount],max([Exempted Amount]) as [Exempted Amount],max([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max(MANDI_TAX_AMT) as MANDI_TAX_AMT" &
                " ,max([Executive]) as [Executive] " & IIf(Batch_Wise = True, ",final.Batch_No ", " ") & ""

            strMCCMaterial += " from ("
            'strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"
            strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
            '' query for no tax applied

            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateBlankColumn & strTaxColumns & strSDEndQry & strSDJoinQry & "  and (coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax2,'')='' " &
                              " and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax4,'')='' and " &
                              " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax5,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax6,'')='' and " &
                              " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax7,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax8,'')='' and " &
                              " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax9,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax10,'')='') )t "

            '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + ") " & _

            strMCCMaterial += "   union all"
            '' query for tax1 applied
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"

            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax1 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & "  and TSPL_SD_SALE_INVOICE_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX2 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax2_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX2+'%' as Tax2Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax2 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX3 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax3_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX3+'%' as Tax3Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax3 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "   union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX4 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax4_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX4+'%' as Tax4Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax4 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX5 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax5_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX5+'%' as Tax5Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax5 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX6 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax6_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX6+'%' as Tax6Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax6 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX7 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax7_AMt,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX7+'%' as Tax7Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax7 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX8 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax8_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX8+'%' as Tax8Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax8 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX9 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax9_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX9+'%' as Tax9Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax9 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += "  union all"
            strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX10 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Tax10_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX10+'%' as Tax10Rate"
            strSDTaxRate = " left join TSPL_TAX_RATES ttr  on ttr.tax_Code=TSPL_SD_SALE_INVOICE_DETAIL.tax10 and ttr.tax_Rate=TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate and ttr._type<>'OH' and ttr.Tax_Type='S'"
            strMCCMaterial += " select * from (" & strSDCommonQuery & strSDTaxRateColumn & strTaxColumns & strSDEndQry & strSDTaxRate & strSDJoinQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
            strMCCMaterial += " )final"
            strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
            strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
            strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
            strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
            strMCCMaterial += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER .City_Code =TSPL_CUSTOMER_MASTER.City_Code " &
                              " LEFT JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State=TSPL_STATE_MASTER.STATE_CODE "


            ''Monika
            ''strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "

            'added by stuti on 01/05/2017
            strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

            ''Monika
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If
            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMCCMaterial += " and final.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If obj.Login_User_Mapped_Customer_Category_List IsNot Nothing AndAlso obj.Login_User_Mapped_Customer_Category_List.Count > 0 Then
                strMCCMaterial += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(obj.Login_User_Mapped_Customer_Category_List) + ") "
            End If
            ''end here

            strMCCMaterial += " group by  final.Screen_Type,final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Scheme Type], [Cash Scheme Code] , [Cash Scheme Amount] , [Price Code] ,final.Created_By,final.Modify_By,final.RATE_UOM,final.Conv_Factor,final.Sampling,final.Scheme_Item " & IIf(Batch_Wise = True, ",final.Batch_No ", " ") & "" '', " + strPivotFoGrouprOuterQuery + " ,QC.FAT_Per,QC.SNF_Per

        End If
        strMCCMaterial += ") xx"

        strMCCMaterial += " Left Join (Select  DOCUMENT_CODE,Item_Code,unit_code,Distributor_Commission_Amt,Transporter_Commission_Amt,Security_Amt,HeadDiscAmt,HeadDiscPer from TSPL_SD_SHIPMENT_DETAIL Group By Item_Code,Unit_code,DOCUMENT_CODE,Distributor_Commission_Amt,Security_Amt,Transporter_Commission_Amt,HeadDiscAmt,HeadDiscPer)TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=xx.[Shipment No] And TSPL_SD_SHIPMENT_DETAIL.Item_Code=xx.[Item Code] and TSPL_SD_SHIPMENT_DETAIL.unit_code=xx.[UOM] "
        If obj.ReportType = ("Customer Amount Wise") Then
            strMCCMaterial += "  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='LTR') as ItemConversionInLTR on ItemConversionInLTR.Item_code=xx.[Item Code] "
        End If


        '===============Added by preeti Gupta ===
        strMCCMaterial += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =xx.[Location Code] "
        '=======================================
        ''richa agarwal change to show csa sales account for csa sale and csa sale return ''richa BHA/12/08/19-000922 add sale return account 
        '' strMCCMaterial += " left outer join (select ITEM.Item_Code,ITEM.Item_Desc,ITEM.Structure_Code,SA.Sales_Account from TSPL_ITEM_MASTER Item left join TSPL_SALES_ACCOUNTS SA on Item.Sale_Class_Code=SA.Sales_Class_Code) Item on Item.Item_Code =xx.[Item Code] "
        strMCCMaterial += " left outer join (select ITEM.Item_Code,ITEM.Item_Desc,ITEM.Structure_Code,SA.Sales_Account,ISNULL(SA.Sales_Return_Account,'') as Sales_Return_Account, isnull(CA.GSOC_Acct,'') as GSOC_Acct,Item.STD_FATPER,Item.STD_SNFPer from TSPL_ITEM_MASTER Item left join TSPL_SALES_ACCOUNTS SA on Item.Sale_Class_Code=SA.Sales_Class_Code left join TSPL_CUSTOMER_ACCOUNT_SET  CA on Item.Cust_Account =CA.Cust_Account ) Item on Item.Item_Code =xx.[Item Code] "
        ''--------------------------
        '' transaction unit conversion
        strMCCMaterial += " left join (" & qryTransStock & ") as  TransStock on xx.[Item Code]=TransStock.Item_Code "
        If Not Stock_uom Then
            strMCCMaterial += " and TransStock.UOM_Code=" & IIf(clsCommon.myLen(Unit_Code) <= 0, "xx.[UOM]", "'" & Unit_Code & "'") & " "
        End If
        ''end transaction unit conversion
        strMCCMaterial += " left join (" & strItemGroup & ") as Item_Group on Item.Structure_Code=Item_Group.Structure_Code "
        'strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location "
        strMCCMaterial += " left join (" & qryStock & ") as Stock_SU on xx.[Item Code]=Stock_SU.Item_Code and xx.[UOM]=Stock_SU.UOM_Code "
        '===============================Added By Preeti Gupta===================================
        strMCCMaterial += " left join (" & qryStock & ") as Rate_Stock_SU on xx.[Item Code]=Rate_Stock_SU.Item_Code and isnull(xx.[RATE_UOM],xx.UOM) =Rate_Stock_SU.UOM_Code  "
        strMCCMaterial += " left join (" & qryTransStock & ") as Rate_select_SU on xx.[Item Code]=Rate_select_SU.Item_Code  "
        If Not Stock_uom Then
            strMCCMaterial += " and Rate_select_SU.UOM_Code=" & IIf(clsCommon.myLen(Unit_Code) <= 0, "xx.[UOM]", "'" & Unit_Code & "'") & ""
        End If
        ' ============================================================================================ 'added by richa  BHA/17/08/18-000441 route
        strMCCMaterial += " left join (" & qryKG & ") as StockKG on xx.[Item Code]=StockKG.Item_Code  
         left join (select Cust_Code,Cust_Group_Code,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Struct_Code,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.state as State_Code,tspl_State_Master.State_Name,TSPL_CUSTOMER_MASTER.cust_category_code,RSM as [RSM Code],RSM_MASTER.Emp_Name AS [RSM Name],ZSM as [ZSM Code],ZSM_MASTER.Emp_Name AS [ZSM Name],ASM AS [ASM Code],ASM_MASTER.Emp_Name AS [ASM Name], ASO as [ASO Code],ASO_MASTER.Emp_Name AS [ASO Name] from TSPL_CUSTOMER_MASTER left join tspl_State_Master on tspl_State_Master.state_Code=TSPL_CUSTOMER_MASTER.state LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER AS RSM_MASTER ON RSM_MASTER.EMP_CODE = TSPL_CUSTOMER_MASTER.RSM  LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER AS ZSM_MASTER ON ZSM_MASTER.EMP_CODE = TSPL_CUSTOMER_MASTER.ZSM  LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER AS ASM_MASTER ON ASM_MASTER.EMP_CODE = TSPL_CUSTOMER_MASTER.ASM LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER AS ASO_MASTER ON ASO_MASTER.EMP_CODE = TSPL_CUSTOMER_MASTER.ASO) as Cust on xx.[Customer Code]=Cust.Cust_Code  left join (select location_Code as Cust_Code,Tin_No,State_Code,State_Name from TSPL_location_MASTER left join tspl_State_Master on tspl_State_Master.state_Code=TSPL_location_MASTER.state) as Cust_Loc on xx.[Customer Code]=Cust_Loc.Cust_Code  
         left join (select Cust_Group_Code,Cust_Group_Desc from TSPL_CUSTOMER_GROUP_MASTER) as Cust_Group on Cust.Cust_Group_Code=Cust_Group.Cust_Group_Code 
          left join (select Zone_Code,Description from TSPL_ZONE_MASTER) as Zone on Cust.Zone_Code=Zone.Zone_Code 
                           left join TSPL_LOCATION_MASTER as Loc on Loc.Location_Code=xx.[Location Code] 
                           left join TSPL_STATE_MASTER GSTState on loc.State=GSTState.STATE_CODE 
                            left outer join (select TSPL_SD_SALE_INVOICE_HEAD.Route_No,TSPL_SD_SALE_INVOICE_HEAD.Document_Code,TSPL_ROUTE_MASTER.Route_Desc  from  TSPL_SD_SALE_INVOICE_HEAD LEFT OUTER JOIN  TSPL_ROUTE_MASTER ON TSPL_ROUTE_MASTER.Route_No =TSPL_SD_SALE_INVOICE_HEAD.Route_No )  as Route_Table on Route_Table.Document_Code =[Document No] "
        If clsCommon.myLen(strCategoryTable) > 0 Then
            strMCCMaterial += " left outer join (" + strCategoryTable + ") as VirtualCategoryTabel on  VirtualCategoryTabel.Item_Code=xx.[Item Code]"
        End If

        ''==========Monika
        'strMCCMaterial += " left outer join (" + qryQC + ")QC on QC.item_code=xx.[Item Code] "
        ''================end here=====
        ' done by priti BHA/09/07/18-000139 for showing costing for sale return document ,BHA/03/12/18-000733 done by richa show cogs for product sale return,BHA/22/05/20-000950 cogs for misc sale 28 MAy,2020
        If obj.ReportType = "Net Sale" OrElse obj.ReportType = "Customer Wise" Then
            strMCCMaterial += " left join (select Source_Doc_No,Item_Code,UOM,Location_Code,Qty,max(Avg_Cost) as Avg_Cost from TSPL_INVENTORY_MOVEMENT group by Source_Doc_No,Item_Code,UOM,Location_Code,Qty) IM " &
                "on IM.Source_Doc_No=case when xx.[trans type] in ('Fresh Sale Return','Product Sale Return','Misc Sale') then xx.[Document No]else xx.[Shipment No] end  " &
                "and IM.Item_Code=xx.[Item Code] and IM.UOM=xx.UOM and IM.Location_Code=xx.[Location Code] and " &
                "IM.Qty=case when xx.[trans type] in ('Fresh Sale Return','Product Sale Return')  then -xx.Quantity else xx.Quantity end " 'and IM.Qty=xx.[Quantity]
        End If

        strMCCMaterial += " LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = xx.[Item Code] and TSPL_ITEM_UOM_DETAIL.UOM_Code = xx.uom	 
             Inner JOIN  ( select item_code,uom_code,conversion_factor from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON xx.[Item Code] = I.item_code  where 2 = 2  "
        ''convert(date,xx.Document_Date,103) >= convert(date,('" + From_Date + "'),103) and convert(date,xx.Document_Date,103) <= convert(date,('" + To_Date + "'),103) 
        ''===========Monika
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strMCCMaterial += " and xx.[Location Code] in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strMCCMaterial += " and xx.[Item Code] in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strMCCMaterial += " and xx.[Customer Code] in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If

        'added by richa  BHA/17/08/18-000441 
        If obj.Route_List IsNot Nothing AndAlso obj.Route_List.Count > 0 Then
            strMCCMaterial += " and isnull(Route_Table .Route_No,'') in (" + clsCommon.GetMulcallString(obj.Route_List) + ") "
        End If
        ''========end here



        QryLst.Add(strMCCMaterial)


        'Dim fullpath As String = "D:\RMWebBrowser\RMWebBrowser\bin\Debug" & "\Muhil.xml"
        'Dim fileinfo As System.IO.FileInfo = New System.IO.FileInfo(fullpath)
        'Dim XMLWriter As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(fullpath, System.Text.Encoding.UTF8)
        'XMLWriter.Formatting = Xml.Formatting.Indented
        'XMLWriter.WriteStartDocument()
        'XMLWriter.WriteElementString("Name", strMCCMaterial)

        QryLst.Add(strPivotForFinalOuterQuery)
        QryLst.Add(strCustomerAmountWiseTaxQuery)
        'clsCommon.MyMessageBoxShow("DOne")
        'Return Nothing
        Return QryLst
    End Function
    Public Shared Function ReturnQueryforDocumentInfoLevel(ByVal obj As clsSaleRegisterParameterType) As ArrayList
        Dim stock_uom As Boolean = obj.stockinguom
        Dim From_Date As Date = obj.From_Date
        Dim To_Date As Date = obj.To_Date
        Dim Unit_Code As String = obj.Unit_Code
        Dim QryLst As New ArrayList

        Dim strCodeColumn As String = ""
        Dim strCodeColumnForVirtual As String = ""
        Dim strCodeColumnMax As String = ""
        Dim strCodeDescColumn As String = ""
        Dim strCodeDescColumnMax As String = ""
        Dim strPivotForFinalOuterQuery As String = ""
        Dim strCategoryTable As String = ""
        Dim MIS_Item_Group As String = GetMIS_ITem_GroupColumn()
        Dim dtCategory As New DataTable
        ''===========below If cond.===Add by moinka 12/04/2017
        If obj.rbtnCategorySected OrElse (clsCommon.CompairString(obj.ReportType, "Document Info Level") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.ReportType, "Document Detail") = CompairStringResult.Equal) Then ''if categories are checked from screen only then show as pivot.Or if document detail/info report opened then category pivot run
            dtCategory = clsDBFuncationality.GetDataTable("select ITEM_CATEGORY_CODE AS CodeColumn,ITEM_CATEGORY_CODE+' Description' as CodeDescColumn,DESCRIPTION as DescColumn  from TSPL_ITEM_CATEGORY_LEVEL order by CATEGORY_LEVEL")
            If dtCategory IsNot Nothing AndAlso dtCategory.Rows.Count > 0 Then
                For ii As Integer = 0 To dtCategory.Rows.Count - 1
                    If ii <> 0 Then
                        strCodeColumn += ","
                        strCodeColumnForVirtual += ","
                        strCodeColumnMax += ","
                        strCodeDescColumn += ","
                        strCodeDescColumnMax += ","
                    End If
                    strCodeColumn += "[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeColumnForVirtual += "VirtualCategoryTabel.[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeColumnMax += "max([" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]) as [" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeColumn")).Trim() + "]"
                    strCodeDescColumn += "[" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")) + "]"
                    strCodeDescColumnMax += "max([" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")).Trim() + "]) as [" + clsCommon.myCstr(dtCategory.Rows(ii)("CodeDescColumn")).Trim() + "]"
                Next
                strCategoryTable = "select Item_Code," + strCodeColumnMax + "," + strCodeDescColumnMax + "  from (" + Environment.NewLine &
                " select * from ( " + Environment.NewLine &
                " select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code " + Environment.NewLine &
                " ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+' Description' as Item_Category_CodeDesc " + Environment.NewLine &
                " ,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " + Environment.NewLine &
                " ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc " + Environment.NewLine &
                " from  TSPL_ITEM_MASTER  " + Environment.NewLine &
                " left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code " + Environment.NewLine &
                " left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values" + Environment.NewLine &
                " where 2=2 " + Environment.NewLine &
                " )xx" + Environment.NewLine &
                " Pivot " + Environment.NewLine &
                " ( max(Item_Cagetory_Values) for Item_Category_Code   in ( " + strCodeColumn + ")" + Environment.NewLine &
                " ) Pivt" + Environment.NewLine &
                " Pivot " + Environment.NewLine &
                " (" + Environment.NewLine &
                " max(Category_Value_Desc) for Item_Category_CodeDesc in (" + strCodeDescColumn + ")" + Environment.NewLine &
                " ) Pivt1 " + Environment.NewLine &
                " ) xxx group by Item_Code "
                ''End of Category Table start now.
            End If
        End If

        ''Virtual Category Table start now.


        Dim strMCCMaterial As String = ""
        Dim qryTaxQuery As String = ""
        Dim strPivotForOuter As String
        Dim strPivotForOuterMax As String
        'Dim lstTables As New List(Of String)
        'lstTables.Add("TSPL_SD_SALE_INVOICE_DETAIL")
        'lstTables.Add("TSPL_SCRAPINVOICE_DETAIL")
        'lstTables.Add("TSPL_SD_SALE_RETURN_DETAIL")
        'lstTables.Add("TSPL_TRANSFER_ORDER_DETAIL")
        'lstTables.Add("TSPL_CSA_TRANSFER_DETAIL")
        qryTaxQuery = GetTaxQuery(obj)

        'strPivotForOuter = " select distinct (select Distinct ',sum(isnull(final.'+tax1+',0)) as '+TAX1 from ( " & qryTaxQuery
        strPivotForOuter = "select distinct (select Distinct ',sum(isnull(final.['+tax1+'],0)) as ['+TAX1+']' from ( " & qryTaxQuery
        strPivotForOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strPivotForOuterQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForOuter))
        '===============================================
        strPivotForOuterMax = "select distinct (select Distinct ',max(isnull(final.['+tax1+'],0)) as ['+TAX1+']' from ( " & qryTaxQuery
        strPivotForOuterMax += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strPivotForOuterQueryMAX As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForOuterMax))
        '============================

        Dim strPivotForFinalOuter As String
        strPivotForFinalOuter = ""
        strPivotForFinalOuter = " select distinct (select Distinct ',xx.['+tax1+']' from ( " & qryTaxQuery
        strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        strPivotForFinalOuterQuery = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))

        Dim strPivotForFinalOuterPercent As String
        strPivotForFinalOuterPercent = " select distinct (select  Distinct ',xx.['+tax1+'%'+']' from ( " & qryTaxQuery
        strPivotForFinalOuterPercent += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        Dim strPivotForFinalOuterPercentQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuterPercent))

        Dim strPivotForTransfer_In As String
        strPivotForFinalOuter = ""
        strPivotForFinalOuter = " select distinct (select Distinct ',0 as ['+tax1+']' from ( " & qryTaxQuery
        strPivotForFinalOuter += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        strPivotForTransfer_In = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForFinalOuter))

        Dim strPivotFortRANSFER_INPercent As String
        strPivotFortRANSFER_INPercent = " select distinct (select  Distinct ',0 as ['+tax1+'%'+']' from ( " & qryTaxQuery
        strPivotFortRANSFER_INPercent += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"
        Dim strPivotFortRANSFER_INPercentQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotFortRANSFER_INPercent))
        '===========

        Dim strPivotForGroupOuter As String
        strPivotForGroupOuter = "select SUBSTRING(ax,2,len(Ax)) from ("
        'strPivotForGroupOuter += " select distinct (select Distinct ',['+tax1+'%'+']' from ( " & qryTaxQuery
        strPivotForGroupOuter += " select distinct (select Distinct ',max(isnull(final.['+tax1+'%'+'],0)) as ['+TAX1+'%'+']' from ( " & qryTaxQuery

        strPivotForGroupOuter += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"
        Dim strPivotFoGrouprOuterQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForGroupOuter))



        Dim strPivotForOuterForBulk As String
        strPivotForOuterForBulk = " select distinct (select Distinct ',0 as ['+TAX1+']' from ( " & qryTaxQuery

        strPivotForOuterForBulk += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strPivotForOuterQueryforBulk As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForOuterForBulk))

        Dim strDoublePivotForOuterForBulk As String

        strDoublePivotForOuterForBulk = " select distinct (select Distinct ',0 as ['+tax1+'%'+']' from ( " & qryTaxQuery


        strDoublePivotForOuterForBulk += " )aa where len(isnull(TAX1,''))>0 for xml path('') )"

        Dim strDoublePivotForOuterQueryforBulk As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForOuterForBulk))


        Dim strPivotForInner As String
        strPivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        strPivotForInner += " select distinct (select Distinct ',['+tax1+']' from ( " & qryTaxQuery

        strPivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strPivotForInnerQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForInner))

        '' taxcolumns for no tax 
        strPivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        strPivotForInner += " select distinct (select Distinct ',Null as ['+tax1+']' from ( " & qryTaxQuery

        strPivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strPivotForInnerQueryNoTax As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strPivotForInner))


        Dim strDoublePivotForInner As String
        strDoublePivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        strDoublePivotForInner += " select distinct (select Distinct ',['+tax1+'%'+']' from ( " & qryTaxQuery

        strDoublePivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strDoublePivotForInnerQuery As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForInner))

        '' tax rate columns for no tax 
        strDoublePivotForInner = "select SUBSTRING(ax,2,len(Ax)) from ("
        strDoublePivotForInner += " select distinct (select Distinct ',Null as ['+tax1+'%'+']' from ( " & qryTaxQuery

        strDoublePivotForInner += " )a where len(isnull(TAX1,''))>0 for xml path('') )ax)Axx"

        Dim strDoublePivotForInnerQueryNoTax As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDoublePivotForInner))


        Dim qryQC As String = ""
        qryQC = " select Item_Code,MAX(Fat_Per) as Fat_Per,MAX(SNF_Per) as SNF_Per from (" &
                " select Item_QCP.Item_Code,Item_QCP.Code as Parameter_Code,(case when QCP.Type='FAT' then Item_QCP.Actual_Range end) as Fat_Per," &
                " (case when QCP.Type='SNF' then Item_QCP.Actual_Range  end) as SNF_Per from TSPL_ITEM_QC_PARAMETER_MASTER Item_QCP " &
                " left join TSPL_PARAMETER_MASTER QCP  on Item_QCP.Code=QCP.Code) as QC group by Item_Code"

        Dim qryKG As String = ""
        qryKG = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='KG'"
        'qryKG = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='KG'"
        'qryKG = " select distinct TSPL_ITEM_UOM_DETAIL.Item_Code,coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) as UOM_Code, " & _
        '        " (case when Weigt1.Contained_UOM='KG' then round(Weigt1.Contained_Qty/Weigt1.Container_Qty,4) " & _
        '        " when Weigt2.Container_UOM='KG' then round(Weigt2.Container_Qty/Weigt2.Contained_Qty,4) else 0 end)*TSPL_ITEM_MASTER.Weight_Value as Conversion_Factor " & _
        '        " from TSPL_ITEM_UOM_DETAIL " & _
        '        " left join TSPL_ITEM_MASTER on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION where Contained_UOM='KG') Weigt1 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt1.Container_UOM  and (TSPL_ITEM_MASTER.Product_Type=Weigt1.Product_Type or Weigt1.Product_Type='All') " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION where Container_UOM='KG') Weigt2 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt2.Contained_UOM and (TSPL_ITEM_MASTER.Product_Type=Weigt2.Product_Type or Weigt2.Product_Type='All')  " & _
        '        " where  2=2 and coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) is not null and coalesce(Weigt1.Container_UOM,Weigt2.Container_UOM) in ('KG','Ltr')"

        '' done by Panch Raj from Preeti User Id
        'qryKG = " select distinct TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Weight_UOM," & _
        '        " max(case when TSPL_ITEM_MASTER.Weight_UOM='KG' then 1 " & _
        '        " else  coalesce(Weigt1.Contained_Qty/Weigt1.Container_Qty,Weigt2.Container_Qty/Weigt2.Contained_Qty)end)*TSPL_ITEM_MASTER.Weight_Value " & _
        '        " as Conversion_Factor from TSPL_ITEM_MASTER " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION ) Weigt1 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt1.Container_UOM and Weigt1.Contained_UOM='KG'  " & _
        '        " and (TSPL_ITEM_MASTER.Product_Type=Weigt1.Product_Type or Weigt1.Product_Type='All') " & _
        '        " left join (select distinct Container_Qty,Container_UOM,Contained_Qty,Contained_UOM,Product_Type from TSPL_WEIGHT_CONVERSION ) Weigt2 " & _
        '        " on TSPL_ITEM_MASTER.Weight_UOM=Weigt2.Contained_UOM and  Weigt2.Container_UOM='KG' " & _
        '        " and (TSPL_ITEM_MASTER.Product_Type=Weigt2.Product_Type or Weigt2.Product_Type='All')" & _
        '        " where  2=2 and len(coalesce(TSPL_ITEM_MASTER.Weight_UOM,''))>0 group by TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Weight_UOM,TSPL_ITEM_MASTER.Weight_Value"
        Dim qryStock As String = ""
        qryStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL "

        '' query for transaction  UOM conversion
        Dim qryTransStock As String = ""
        If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not stock_uom Then
            qryTransStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL "
        ElseIf clsCommon.myLen(Unit_Code) <= 0 AndAlso stock_uom Then
            qryTransStock = "select Item_Code,max(UOM_Code) as UOM_Code,max(Conversion_Factor) as Conversion_Factor from TSPL_ITEM_UOM_DETAIL where  Stocking_Unit='Y'  group by Item_Code "
        Else
            qryTransStock = "select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='" & Unit_Code & "'"
        End If
        '===================Added By Preeti Gupta===================
        Dim qryRateStock As String = ""
        ' =============================================================

        '' end query for transaction  UOM conversion
        ' BM00000008464 BM00000008484
        '' query for structure and item group custom field========================BM00000008352
        '==============added by preeti Gupta against ticket No[BM00000009864,BM00000007115]
        Dim strItemGroup As String = ""
        strItemGroup = " select Struct.Structure_Code,Structure_Descq,Struct_Val.Value as Item_Group,StructDtl.Description as Group_Description from TSPL_STRUCTURE_MASTER Struct left join (" &
                       " select Custom_field_Code,Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES where Program_Code='" & clsUserMgtCode.itemStructure & "'  " &
                       " and Custom_Field_Code='" & MIS_Item_Group & "') as Struct_Val  on Struct.Structure_Code=Struct_Val.Transaction_Code" &
                       " left join (select Custom_Field_Code,SNo,Value,Description from TSPL_CUSTOM_FIELD_DETAIL where Custom_Field_Code='" & MIS_Item_Group & "') as StructDtl on Struct_Val.Value=StructDtl.Value "

        Dim strSDCommonQuery As String = ""
        Dim strTaxColumns As String = ""
        Dim strSDEndQry As String = ",TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1_Rate"
        strSDCommonQuery = " select case when ISNUll(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,'')<>'' then TSPL_SD_SALE_INVOICE_HEAD.Document_Type else  CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='ALL' THEN 'SD' ELSE TSPL_SD_SALE_INVOICE_HEAD.Trans_Type END end as Trans_Type  ,TSPL_SD_SALE_INVOICE_HEAD.Status ,TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location, " &
                           " TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_INVOICE_HEAD.Document_Type,TSPL_SD_SALE_INVOICE_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_INVOICE_HEAD.Document_Code , " &
                           " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Line_No , " &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Qty ,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code , TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as Item_Cost , " &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Amount *(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end as AMount,TSPL_SD_SALE_INVOICE_DETAIL.Disc_Per ,TSPL_SD_SALE_INVOICE_detail.Total_Disc_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) - case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end as Disc_Amt,case when coalesce(TSPL_SD_SALE_INVOICE_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.sampling,0)=1 or coalesce(TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item,'')='Y' then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) end as [Scheme Amount] , " &
                           " (TSPL_SD_SALE_INVOICE_detail.Amount-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)- case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='AS' then coalesce(TSPL_SD_SALE_INVOICE_Detail.Cash_Scheme_Amount,0)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end as Amt_Less_Discount , " &
                           " TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as Total_Tax_AMt ,(TSPL_SD_SALE_INVOICE_DETAIL.Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as Total_Amt,case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then '' else TSPL_SD_SALE_INVOICE_HEAD.Vehicle_Code end as Vehicle_Code , case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='PS' then TSPL_SD_SALE_INVOICE_HEAD.VehicleNo else COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_SD_SALE_INVOICE_HEAD.VEHICLENO) end as Vehicle_No,(case when TSPL_SD_SALE_INVOICE_DETAIL.Line_No=1 then (TSPL_SD_SALE_INVOICE_HEAD.Total_Add_Charge+coalesce(TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount,0))*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) else 0 end) as  Additional_Charge, " &
                           " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as [AR Document Amt]," &
                           " TSPL_Customer_Invoice_Head.Discount_Amount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as [AR Document Discount Amt], " &
                           " TSPL_Customer_Invoice_Head.amount_less_Discount*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as [AR Total Tax], " &
                           " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) as [AR Total Add Charge], " &
                           " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                           " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_INVOICE_HEAD.GRNo as [GR No],TSPL_SD_SHIPMENT_HEAD.gr_date as [GR Date],TSPL_SD_SALE_INVOICE_HEAD.WayBillNo as [WayBill No],TSPL_SD_SHIPMENT_HEAD.transport_id as [Transporter Code] ,case when len(TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual) > 0 then TSPL_SD_SHIPMENT_HEAD.Transporter_Name_Manual else TSPL_SD_SHIPMENT_HEAD.Transporter_Name end as [Transporter Name],(case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS else  TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code end) as [Delivery No]  ,Shipment_Code as [Shipment No],
                           
                           TSPL_SD_SALE_INVOICE_DETAIL.MRP ,TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Code as [Scheme Code] ,TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Code as [Cash Scheme Code] ,TSPL_SD_SALE_INVOICE_DETAIL.Cash_Scheme_Amount as [Cash Scheme Amount], TSPL_SD_SALE_INVOICE_DETAIL.Price_code as [Price Code],
                           case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code else case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then ''  else '' end end   as Ref_Doc_no,'' as Ref_doc_date," &
                            " " &
                            "  case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Created_By  when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then '' else '' end  Created_By , " &
                            "case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Modify_By  when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then '' else '' end  as Modify_By , " &
                            " case when TSPL_SD_SALE_INVOICE_HEAD.trans_type='PS' then TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Modify_date when TSPL_SD_SALE_INVOICE_HEAD.trans_type='FS' then '' else NULL end  as Modify_date ,TSPL_SD_SALE_INVOICE_DETAIL.RATE_UOM,TSPL_SD_SALE_INVOICE_DETAIL.Conv_Factor,tspl_sd_sale_invoice_detail.Sampling,tspl_sd_sale_invoice_detail.Scheme_Item ,"

        strSDEndQry = " from TSPL_SD_SALE_INVOICE_DETAIL " &
                           " left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE " &
                           " left join TSPL_VEHICLE_MASTER on TSPL_SD_SALE_INVOICE_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                           " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code " &
                           "  " &
                           "  " &
                           " left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code " &
                           " LEFT OUTER JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No " &
                           " left join TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE on TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code =TSPL_SD_SHIPMENT_head.Delivery_Code_ps " &
                           " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code " &
                           " where  isnull(TSPL_SD_SALE_INVOICE_HEAD.trans_type,'ALL') not in ('CSA')  " &
                           " and (case when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='FS' then 'Fresh Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='PS' then 'Product Sale' " &
                           " when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='MCC' then 'MCC Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type <>'MT' then 'Export Sale' when TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type ='MT' then 'Merchant Trade' WHEN TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='SD' then 'General Sale' " &
                           " else  TSPL_SD_SALE_INVOICE_HEAD.Trans_Type end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") "

        If obj.supply_Date Then
            strSDEndQry += " and  convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= convert(date,('" & To_Date & "'),103) "
        Else
            strSDEndQry += " and  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "
        End If
        ' and  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If

        If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
            strSDEndQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
        End If

        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_INVOICE_HEAD.Document_Code = '" & obj.Document_Code & "' "
        End If
        '===================added by preeti gupta against ticket no[BM00000009858]
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowParticluarColumnInSalesRegisterForGopalJee, clsFixedParameterCode.ShowParticluarColumnInSalesRegisterForGopalJee, Nothing)) = 1 Then
            strMCCMaterial = " select [Trans Type],[Document_date],[Document No],[Location Code]as [Warehouse Code], [Location Name] as [Warehouse Name],Cust.Cust_Group_Code as [Customer Group Code],Cust_Group.Cust_Group_Desc as [Customer Group Description],[Customer Code],[Customer Name],Item_Group.Item_Group as [Product Group Code],Item_Group.Group_Description as [Product Group Description],Vehicle_Code as [Vehicle Code],Vehicle_No as [Vehicle No] "

            strMCCMaterial += " , [Item Code] as [Product Code],[Item Name] as [Product Name],cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3)) as [Pack Sold]," & IIf(clsCommon.myLen(Unit_Code) <= 0, IIf(stock_uom = True, "TransStock.UOM_Code", "xx.[UOM]"), "'" & Unit_Code & "'") & " as [Unit of sale],"
            If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not stock_uom Then
                strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  as [Rate] "
            Else
                strMCCMaterial += "  cast(( case when isnull(Rate_Stock_SU.Conversion_Factor,0)<=0 then ([Item Cost]) else ([Item Cost] * Rate_select_SU.Conversion_Factor)/ Rate_Stock_SU.Conversion_Factor end )"
                strMCCMaterial += " as Numeric(18,3)) as [Rate]"
            End If
            strMCCMaterial += " ,Amount " + strPivotForFinalOuterQuery + " ,[Total Tax Amount], [Total Amount] as [Total Amount],'' as Remarks,''as Relation,'' as CTYPE "

        Else
            strMCCMaterial = " select [Trans Type],[Location Code],[Location Name],Loc.State as [Location State],loc.TIN_No as [Dispatch Location Tin No], (CASE WHEN [Invoice Type]='T' THEN 'Tax' when [Invoice Type]='R' then 'Retail' when [Invoice Type]='N' then 'None' else [Invoice Type] end) as [Invoice Type],[Document No],[Document_date],Vehicle_Code as [Vehicle Code],Vehicle_No as [Vehicle No],cast(Additional_Charge as numeric(18,2)) as [Additional Amount],[Customer Code],[Customer Address],Cust.Struct_Code,[Customer Name],Cust.Cust_Group_Code as [Customer Group Code],Cust_Group.Cust_Group_Desc as [Customer Group Description],Cust.cust_category_code as [Customer Category],Cust.Zone_Code as [Customer Zone Code],Zone.Description as [Customer Zone Description], [Parent Customer No],[Parent Customer Code], [Parent Customer Name],coalesce(Cust.State_Code,Cust.State_Code) as [Customer State Code],coalesce(Cust.State_Name,Cust_Loc.State_Name) as [Customer State Desc],coalesce(Cust.Tin_No,cust_loc.Tin_No) as [Tin No],Item_Group.Item_Group as [Item Group Code],Item_Group.Group_Description as [Item Group Description] "
            If clsCommon.myLen(strCategoryTable) > 0 Then
                ''richa agarwal to avoid ambiguous error
                '  strMCCMaterial += "," + strCodeColumn + "," + strCodeDescColumn
                strMCCMaterial += "," + strCodeColumnForVirtual + "," + strCodeDescColumn
            End If
            ' BM00000008438 BM00000008391
            'strMCCMaterial += " , [Item Code],[Item Name],cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3)) as [Quantity]," & IIf(clsCommon.myLen(Unit_Code) <= 0, "xx.[UOM]", "'" & Unit_Code & "'") & " as [UOM],[Item Cost] as [Item Rate],[Fat Per] as [FAT %],[SNF Per] as [SNF %],cast(([Quantity]*[Fat Per]*Stock_SU.Conversion_Factor)/(100*coalesce(StockKG.Conversion_Factor,1)) as numeric(18,3)) as [FAT KG],cast(([Quantity]*[SNF Per]*Stock_SU.Conversion_Factor)/(100*coalesce(StockKG.Conversion_Factor,1)) as Numeric(18,3)) as [SNF KG],Amount,[Discount Per] as [Discount %],  (coalesce( [Discount Amount],0)-coalesce([Scheme Amount],0))  as [Discount Amount],[Scheme Amount],[Amount Less Discount]  as [Amount Less Discount]" + strPivotForFinalOuterQuery + " " + strPivotForFinalOuterPercentQuery + ",case when [trans type]='Fresh Sale Retturn' then [Amount Less Discount] + case when coalesce([Total Tax Amount],0)<0 then 0 else coalesce( [Scheme Amount],0) end else ([Total Amount]-[Total Tax Amount] + case when coalesce([Total Tax Amount],0)>0 then 0 else coalesce( [Scheme Amount],0) end) end as [Sale Amount],[Total Tax Amount], (cast(Additional_Charge as numeric(18,2))+[Total Amount]) as [Total Amount], " & _
            strMCCMaterial += " , [Item Code],[Item Name],"

            If obj.UOMType = "Stock UOM" OrElse obj.ReportType = "Custom Type" Then
                strMCCMaterial += " cast(([Quantity]*Stock_SU.Conversion_Factor)/(case when coalesce(TransStock.Conversion_Factor,1)=0 then 1 else coalesce(TransStock.Conversion_Factor,1) end) as Numeric(18,3)) as [Quantity],"
                strMCCMaterial += " " & IIf(clsCommon.myLen(Unit_Code) <= 0, IIf(stock_uom = True, "TransStock.UOM_Code", "xx.[UOM]"), "'" & Unit_Code & "'") & " as [UOM],"
            Else
                strMCCMaterial += " Isnull(convert(decimal(18,2),( [Quantity] * isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1)) /I.Conversion_Factor),0) as Quantity,I.UOM_Code as UOM,"
            End If

            If clsCommon.myLen(Unit_Code) <= 0 AndAlso Not stock_uom Then
                strMCCMaterial += " cast(([Item Cost]*Stock_SU.Conversion_Factor)/(case when coalesce(rate_stock_su.Conversion_Factor,1)=0 then 1 else coalesce(rate_stock_su.Conversion_Factor,1) end) as Numeric(18,3))  as [Item Rate] "
            Else
                strMCCMaterial += "  cast(( case when isnull(Rate_Stock_SU.Conversion_Factor,0)<=0 then ([Item Cost]) else ([Item Cost] * Rate_select_SU.Conversion_Factor)/ Rate_Stock_SU.Conversion_Factor end )"
                strMCCMaterial += " as Numeric(18,3)) as [Item Rate]"
            End If
            strMCCMaterial += " , [Fat Per] as [FAT %],[SNF Per] as [SNF %],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [Fat Per] else QC.fat_per end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as numeric(18,3)) end) as [FAT KG],(case when coalesce(StockKG.Conversion_Factor,0)=0 then 0 else cast(([Quantity]* (case when [trans type] in ('Bulk Sale Trade','Bulk Sale','Bulk Sale Return','MCC Transfer','SS','Tanker Dispatch Return','MCC Tanker Dispatch Return') then [SNF Per] else snf_per end) *Stock_SU.Conversion_Factor)/(coalesce(StockKG.Conversion_Factor,1)*100) as Numeric(18,3)) end) as [SNF KG],Amount,[Discount Per] as [Discount %],  (coalesce( [Discount Amount],0)-coalesce([Scheme Amount],0))  as [Discount Amount],[Scheme Amount],[Amount Less Discount]  as [Amount Less Discount]" + strPivotForFinalOuterQuery + " " + strPivotForFinalOuterPercentQuery + ",case when [trans type]='Fresh Sale Return' then [Amount Less Discount]  else ([Total Amount]-[Total Tax Amount] + case when coalesce([Total Tax Amount],0)>0 then 0 else coalesce( [Scheme Amount],0) end) end as [Sale Amount],[Total Tax Amount], (cast(Additional_Charge as numeric(18,2))+[Total Amount]) as [Total Amount], " &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax]+ case when (coalesce([Total Tax Amount],0)<>0 or [Scheme Amount]<=0) and [Document No]<>'SRFS-003/15-16/000006' then 0 else coalesce([AR Document Discount Amt],0)  end as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge], "
            ''richa agarwal change to show csa sales account for csa sale and csa sale return
            'left(Item.Sales_Account, Len(Item.Sales_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code as [Sales Account],
            strMCCMaterial += "  case when [trans type] in ('CSA Sale','CSA Sale Return') then case when item.GSOC_Acct<>'' then  left( item.GSOC_Acct, Len( item.GSOC_Acct)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code else '' end  else  left(Item.Sales_Account, Len(Item.Sales_Account)-3)+  TSPL_LOCATION_MASTER.Loc_Segment_Code  end as [Sales Account], " &
            " xx.[GR No],case when convert(date,xx.[GR Date],103)='01/Jan/1990' or convert(date,xx.[GR Date],103)='01/Jan/0001' then NULL else  convert(varchar,xx.[GR Date],103) end as [GR Date]  ,[WayBill No],[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code] , [Cash Scheme Code] , [Cash Scheme Amount], [Price Code], case when Sampling=0 then  'N' else case when sampling=1 then'Y' end end as sampling, Scheme_Item as [Scheme Type],xx.Ref_Doc_no as [Ref Doc No],convert(varchar,xx.Ref_doc_date,103) as [Ref Doc Date],Created_By.User_Name as [Created By],Modify_By.User_Name as [Posted By],Convert(varchar,xx.Modify_date,103) as [Modify date] "
        End If


        ''richa agarwal add merchant trade trans_type in below qry BM00000008390 (Applied For DCC Also)
        strMCCMaterial += " from (select case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='EX' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' when Trans_Type='MT' then 'Merchant Trade' WHEN Trans_Type ='SD' then 'General Sale' else  Trans_Type end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) AS [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQueryMAX + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]- case when (Trans_Type ='FS') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code] ,[Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By,final.Modify_By,final.Modify_date ,final. RATE_UOM,final. Conv_Factor,final.Sampling,final.Scheme_Item "
        strMCCMaterial += " from ("
        'strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"
        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        '' query for no tax applied
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & "  and (coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax5,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax7,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax9,'')='' and coalesce(TSPL_SD_SALE_INVOICE_DETAIL.tax10,'')='') )t "

        '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + ") " & _

        strMCCMaterial += "   union all"
        '' query for tax1 applied
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX1 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt as Tax1_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1+'%' as Tax1Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & "  and TSPL_SD_SALE_INVOICE_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX2 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt as Tax2_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX2+'%' as Tax2Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX3 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt as Tax3_Amt, TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX3+'%' as Tax3Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX4 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt as Tax4_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX4+'%' as Tax4Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX5 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt as Tax5_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX5+'%' as Tax5Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX6 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt as Tax6_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX6+'%' as Tax6Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX7 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt as Tax7_AMt,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX7+'%' as Tax7Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX8 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt as Tax8_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX8+'%' as Tax8Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX9 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt as Tax9_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate, TSPL_SD_SALE_INVOICE_DETAIL.TAX9+'%' as Tax9Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_INVOICE_DETAIL.TAX10 ,(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end) * TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt as Tax10_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX10+'%' as Tax10Rate"
        strMCCMaterial += " select * from (" & strSDCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_INVOICE_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " )final"
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
        strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
        strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        'strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , [Scheme Code] , [Cash Scheme Code] , [Cash Scheme Amount] , [Price Code] ,final.Ref_Doc_no,final.Ref_doc_date,final.Created_By,final.Modify_By,final.Modify_date,final.Ref_Doc_no,final.Ref_doc_date,final.RATE_UOM,final.Conv_Factor,final.Sampling,final.Scheme_Item " '', " + strPivotFoGrouprOuterQuery + "

        '' add csa transfer 
        Dim stCSACommonQuery As String = " select  'CSA Sale' as  Trans_Type,TSPL_CSA_TRANSFER_HEAD.Status ,TSPL_CSA_TRANSFER_HEAD.From_Location_Code as Bill_To_Location," &
                                          " TSPL_CSA_TRANSFER_HEAD.Cust_Code as Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,'CSA' as Invoice_Type,TSPL_CSA_TRANSFER_HEAD.Doc_Code as Document_Code ," &
                                          " convert(varchar,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103 ) as Document_Date , TSPL_CSA_TRANSFER_DETAIL.Item_Code , " &
                                          " TSPL_CSA_TRANSFER_DETAIL.Qty ,TSPL_CSA_TRANSFER_DETAIL.Unit_code ,TSPL_CSA_TRANSFER_DETAIL.Transfer_Rate*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Item_Cost , " &
                                          " coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Amount ,coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Per,0) as Disc_Per ,coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Amt,0)*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Disc_Amt ,0 as [Scheme Amount] " &
                                          ", (coalesce(TSPL_CSA_TRANSFER_DETAIL.transfer_Rate*TSPL_CSA_TRANSFER_DETAIL.Qty,0)-coalesce(TSPL_CSA_TRANSFER_DETAIL.Disc_Amt,0))*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Amt_Less_Discount ,0 as Total_Tax_Amt ,cast((TSPL_CSA_TRANSFER_DETAIL.Transfer_Rate * TSPL_CSA_TRANSFER_DETAIL.Qty) as Numeric(18,2))*(case when coalesce(TSPL_CSA_TRANSFER_Head.convrate,0)<=0  then 1 else coalesce(TSPL_CSA_TRANSFER_Head.convrate,0) end) as Total_Amt, " &
                                          " TSPL_CSA_TRANSFER_HEAD.Vehicle_Id as Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,0 as Additional_Charge, " &
                                          " '' as [AR Document No],0 as [AR Document Amt],0 as [AR Document Discount Amt], 0 as [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_CSA_TRANSFER_HEAD.GR_No as [GR No],TSPL_CSA_TRANSFER_HEAD.gr_date as [GR Date],TSPL_CSA_TRANSFER_HEAD.WayBill_No as [WayBill No],TSPL_CSA_TRANSFER_HEAD.Transport_Id as [Transporter Code],case when len(TSPL_CSA_TRANSFER_HEAD.transporter_name_manual) > 0 then TSPL_CSA_TRANSFER_HEAD.transporter_name_manual else TSPL_TRANSPORT_MASTER.Transporter_Name end as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],TSPL_CSA_TRANSFER_DETAIL.MRP,'' as  [Scheme Code] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_CSA_TRANSFER_HEAD.DOC_CODE as Ref_Doc_no,TSPL_CSA_TRANSFER_HEAD.Transfer_Date as Ref_doc_date,TSPL_CSA_TRANSFER_HEAD.Created_By ,TSPL_CSA_TRANSFER_HEAD.Modify_By,TSPL_CSA_TRANSFER_HEAD.Modify_date,NULL as RATE_UOM,0 as Conv_Factor, "

        strSDEndQry = " from TSPL_CSA_TRANSFER_DETAIL " &
                    " left outer join TSPL_CSA_TRANSFER_HEAD on TSPL_CSA_TRANSFER_HEAD.DOC_CODE =TSPL_CSA_TRANSFER_DETAIL.DOC_CODE " &
                    " left join TSPL_VEHICLE_MASTER on TSPL_CSA_TRANSFER_HEAD.Vehicle_Id=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                    " left join TSPL_TRANSPORT_MASTER on TSPL_CSA_TRANSFER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_CSA_TRANSFER_HEAD.Cust_Code " &
                    " where 'CSA Sale' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                    " and convert(date,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_CSA_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,('" & To_Date & "'),103) "
        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_CSA_TRANSFER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_CSA_TRANSFER_HEAD.From_Location_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_CSA_TRANSFER_HEAD.Cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If
        If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
            strSDEndQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
        End If
        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_CSA_TRANSFER_HEAD.Doc_Code = '" & obj.Document_Code & "' "
        End If

        strMCCMaterial += " union all "
        strMCCMaterial += " select Trans_Type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount],   " &
                          " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No] ,MRP, [Scheme Code] , [Cash Scheme Code] , [Cash Scheme Amount], [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modify_By,final.Modify_date,final. RATE_UOM,final. Conv_Factor,0 as Sampling,'N' as Scheme_Item "
        strMCCMaterial += " from ("
        'strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX1 ,0 as TAX1_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX1_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX1+'%' as Tax1Rate"
        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        '' query for no tax applied
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & "  and (coalesce(TSPL_CSA_TRANSFER_DETAIL.tax1,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax3,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax5,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax7,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_CSA_TRANSFER_DETAIL.tax9,'')='' and coalesce(TSPL_CSA_TRANSFER_DETAIL.tax10,'')='') )t  "
        '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        '' query for tax1 applied
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX1 ,0 as TAX1_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX1_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX1+'%' as Tax1Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & "  and TSPL_CSA_TRANSFER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for Tax1Rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX2 ,0 as TAX2_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX2_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX2+'%' as Tax2Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX3 ,0 as TAX3_Amt, TSPL_CSA_TRANSFER_DETAIL.TAX3_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX3+'%' as Tax3Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX4 ,0 as TAX4_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX4_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX4+'%' as Tax4Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX5 ,0 as TAX5_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX5_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX5+'%' as Tax5Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX6 ,0 as TAX6_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX6_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX6+'%' as Tax6Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX7 ,0 as TAX7_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX7_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX7+'%' as Tax7Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX8 ,0 as TAX8_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX8_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX8+'%' as Tax8Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX9 ,0 as TAX9_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX9_Rate, TSPL_CSA_TRANSFER_DETAIL.TAX9+'%' as Tax9Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_CSA_TRANSFER_DETAIL.TAX10 ,0 as TAX10_Amt,TSPL_CSA_TRANSFER_DETAIL.TAX10_Rate,TSPL_CSA_TRANSFER_DETAIL.TAX10+'%' as Tax10Rate"
        strMCCMaterial += " select * from (" & stCSACommonQuery & strTaxColumns & strSDEndQry & " and TSPL_CSA_TRANSFER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " )final"
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
        strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
        strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        'strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code] , [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modify_By,final.Modify_date,final. RATE_UOM,final.Conv_Factor" '', " + strPivotFoGrouprOuterQuery + "

        '' end of csa transfer


        strMCCMaterial += " union all"
        strMCCMaterial += " select case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end as Trans_Type ,coalesce(TSPL_LOCATION_MASTER.Main_Location_Code,TSPL_LOCATION_MASTER.Location_Code) as Bill_To_Location,TSPL_INVOICE_Master_BULKSALE.Posted,coalesce(Main_Loc.Location_Desc,TSPL_LOCATION_MASTER.location_desc) as Location_Desc ,'Invoice' as Invoice_type ,TSPL_INVOICE_Master_BULKSALE.Document_No as Document_code ,convert(varchar,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) Document_Date,case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end as Vehicle_Code,case when isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')='' then TSPL_INVOICE_DETAIL_BULKSALE.TradeTanker_No else isnull(TSPL_INVOICE_DETAIL_BULKSALE.Tanker_Code ,'')  end as Vehicle_No,(case when ROW_NUMBER() over (partition by TSPL_INVOICE_Master_BULKSALE.Document_No order by TSPL_INVOICE_DETAIL_BULKSALE.Item_Code )=1 then coalesce(TSPL_INVOICE_Master_BULKSALE.RoundOffAmount,0) else 0 end) as Additional_Charge, " &
                          " TSPL_INVOICE_Master_BULKSALE.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Parent_Customer_No,"
        strMCCMaterial += " Parent_Master.Cust_Code as Parent_Customer_Code,Parent_Master.Customer_Name as Parent_Cust_Name, "
        strMCCMaterial += " TSPL_INVOICE_DETAIL_BULKSALE.Item_Code,tspl_item_master.Item_Desc ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceQty as Qty ,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceRate as Item_cost,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatPer ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFPer ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceFatKG ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceSNFKG  ,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as Amount,0 as Disc_per,0 as Disc_Amt,0 as [Scheme Amount],TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount as Amt_less_Discount,0 as Total_tax_amt " + strPivotForOuterQueryforBulk + " " + strDoublePivotForOuterQueryforBulk + ",(TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount) as Total_Amt, " &
                          " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                          " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                          " (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                          " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],NULL as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],'' as  [Delivery No]  ,'' as  [Shipment No],0 AS MRP,'' as  [Scheme Code] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],TSPL_Dispatch_BulkSale.Document_No as Ref_Doc_no ,TSPL_Dispatch_BulkSale.Document_Date as Ref_doc_date,TSPL_BulkSalePrice_MASTER.Created_By ,TSPL_BulkSalePrice_MASTER.Modified_By ,TSPL_BulkSalePrice_MASTER.Modified_Date ,NULL as RATE_UOM,0 as Conv_Factor ,0 as Sampling,'N' as Scheme_Item " &
                          " from TSPL_INVOICE_DETAIL_BULKSALE "

        strMCCMaterial += " left outer join TSPL_INVOICE_Master_BULKSALE on TSPL_INVOICE_Master_BULKSALE.Document_No =TSPL_INVOICE_DETAIL_BULKSALE.Document_No "
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_INVOICE_Master_BULKSALE.Location_Code"
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as Main_Loc on TSPL_LOCATION_MASTER.Main_Location_Code =Main_Loc.Location_Code"
        strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =TSPL_INVOICE_Master_BULKSALE.Customer_Code"
        strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No"
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_INVOICE_DETAIL_BULKSALE.Item_Code"
        strMCCMaterial += " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No=TSPL_INVOICE_Master_BULKSALE.Document_No "
        strMCCMaterial += " left join TSPL_Dispatch_BulkSale on TSPL_Dispatch_BulkSale.Document_No =TSPL_INVOICE_DETAIL_BULKSALE.Dispatch_Code "
        strMCCMaterial += " left join TSPL_BulkSalePrice_MASTER on TSPL_BulkSalePrice_MASTER.Price_Code =TSPL_Dispatch_BulkSale.Price_Code  where 2=2 " &
                            " and (case when InvoiceAgainst='Against Dispatch Trade' then 'Bulk Sale Trade'  else 'Bulk Sale' end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                            " and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_INVOICE_Master_BULKSALE.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "
        strMCCMaterial += " union all"
        strMCCMaterial += " select case when Trans_Type ='FS' then 'Fresh Sale' when Trans_Type ='CSA' then 'CSA Sale' when Trans_Type='PS' then 'Product Sale' when Trans_Type='MCC' then 'MCC Sale' when Trans_Type='Exp' then 'Export Sale'when Trans_Type='Bulk Sale' then 'Bulk Sale' when Trans_Type ='SS' then 'Misc Sale' WHEN Trans_Type ='SD' then 'General Sale' else Trans_Type  end  as [Trans Type],final.Loc_Code  as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.shipment_No  as [Document No],final.Document_Date as [Document_date],'' as Vehicle_Code,'' as Vehicle_No,0 as Additional_Charge,final.cust_Code  as [Customer Code],MAX(final.CustAdd ) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.shipped_Qty  as [Quantity],final.Unit_code as [UOM],final.price  as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.ItemAmt ,final.DiscountPer  as [Discount Per],final.TotalDiscountAmt   as [Discount Amount],final.[Scheme Amount]   as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.TotalTaxAmt  as [Total Tax Amount],final.Doc_Amt  as [Total Amount], " &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No],MRP , [Scheme Code] ,[Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modify_By,final.Modify_date,final. RATE_UOM,final.Conv_Factor ,0 as Sampling,'N' as Scheme_Item  from ("

        Dim strScarpCommonQry As String = ""
        strScarpCommonQry = " select 'SS' as Trans_Type,TSPL_SCRAPINVOICE_HEAD.ispost as Status ,TSPL_SCRAPINVOICE_HEAD.Loc_Code,TSPL_SCRAPINVOICE_HEAD.cust_Code " &
                            " ,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd, " &
                            " TSPL_SCRAPINVOICE_HEAD.Invoice_Type,TSPL_SCRAPINVOICE_HEAD.invoice_No as shipment_No ,convert(varchar,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103 ) as Document_Date , " &
                            " TSPL_SCRAPINVOICE_DETAIL.Item_Code ,TSPL_SCRAPINVOICE_DETAIL.shipped_Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code ," &
                            " TSPL_SCRAPINVOICE_DETAIL.price ,0 as InvoiceFatPer ,0 as InvoiceSNFPer ,0 as InvoiceFatKG ,0 as InvoiceSNFKG , " &
                            " TSPL_SCRAPINVOICE_DETAIL.ItemAmt ,TSPL_SCRAPINVOICE_DETAIL.DiscountPer,0 as [Scheme Amount] ,TSPL_SCRAPINVOICE_DETAIL.ItemNetAmt as Amt_less_Discount, " &
                            " TSPL_SCRAPINVOICE_DETAIL.TotalDiscountAmt , " &
                            " TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt ,TSPL_SCRAPINVOICE_DETAIL.TotalAmt+Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end as Doc_Amt,'' as Vehicle_Code,'' as Vehicle_No,Case when TSPL_SCRAPINVOICE_DETAIL.line_No=1 then coalesce(TSPL_SCRAPINVOICE_HEAD.add_Amt,0) else 0 end as Additional_Charge," &
                            " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                            " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                            " TSPL_Customer_Invoice_Head.amount_less_Discount as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                            " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge], " &
                            " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                            " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,'' as [GR No],NULL as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP,  '' as  [Scheme Code] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],'' as Ref_Doc_no,NULL as Ref_doc_date, '' as Created_By ,'' as Modify_By,NULL as Modify_date ,NULL as RATE_UOM,0 as Conv_Factor,"

        strSDEndQry = " from TSPL_SCRAPINVOICE_DETAIL  " &
                      " left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No " &
                      " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.AgainstScrap=TSPL_SCRAPINVOICE_HEAD.invoice_No " &
                      " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SCRAPINVOICE_HEAD.Cust_Code  where 2=2 " &
                      " AND 'Misc Sale' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                      " and convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SCRAPINVOICE_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SCRAPINVOICE_HEAD.Loc_Code in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SCRAPINVOICE_HEAD.cust_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If

        If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
            strSDEndQry += " and TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
        End If

        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_SCRAPINVOICE_HEAD.invoice_No = '" & obj.Document_Code & "' "
        End If
        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        '' query for no tax applied

        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and (coalesce(TSPL_SCRAPINVOICE_DETAIL.tax1,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax3,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax5,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax7,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_SCRAPINVOICE_DETAIL.tax9,'')='' and coalesce(TSPL_SCRAPINVOICE_DETAIL.tax10,'')='') )t "
        '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        '' query for tax1 applied
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX1 ,TSPL_SCRAPINVOICE_DETAIL.TAX1_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX1_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX1+'%' as tax1rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax1<>'')"
        strMCCMaterial += " s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX2 ,TSPL_SCRAPINVOICE_DETAIL.TAX2_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX2_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX2+'%' as tax2rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax2<>'' )"
        strMCCMaterial += " s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = "TSPL_SCRAPINVOICE_DETAIL.TAX3 ,TSPL_SCRAPINVOICE_DETAIL.TAX3_Amt , TSPL_SCRAPINVOICE_DETAIL.TAX3_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX3+'%' as tax3rate  "
        strMCCMaterial += "  select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax3<>'' )"
        strMCCMaterial += " s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX4 ,TSPL_SCRAPINVOICE_DETAIL.TAX4_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX4_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX4+'%' as tax4rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax4<>'' )"
        strMCCMaterial += " s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX5 ,TSPL_SCRAPINVOICE_DETAIL.TAX5_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX5_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX5+'%' as tax5rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax5<>'' )"
        strMCCMaterial += " s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX6 ,TSPL_SCRAPINVOICE_DETAIL.TAX6_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX6_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX6+'%' as tax6rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & "  and  TSPL_SCRAPINVOICE_DETAIL.tax6<>'' )"
        strMCCMaterial += " s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX7 ,TSPL_SCRAPINVOICE_DETAIL.TAX7_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX7_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX7+'%' as tax7rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax7<>'' )"
        strMCCMaterial += " s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + ")) t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX8 ,TSPL_SCRAPINVOICE_DETAIL.TAX8_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX8_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX8+'%' as tax8rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax8<>'' )"
        strMCCMaterial += "  s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"

        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX9 ,TSPL_SCRAPINVOICE_DETAIL.TAX9_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX9_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX9+'%' as tax9rate  "
        strMCCMaterial += "  select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax9<>'')"
        strMCCMaterial += " s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        strTaxColumns = " TSPL_SCRAPINVOICE_DETAIL.TAX10 ,TSPL_SCRAPINVOICE_DETAIL.TAX10_Amt ,TSPL_SCRAPINVOICE_DETAIL.TAX10_Rate, TSPL_SCRAPINVOICE_DETAIL.TAX10+'%' as tax10rate  "
        strMCCMaterial += " select * from (" & strScarpCommonQry & strTaxColumns & strSDEndQry & " and  TSPL_SCRAPINVOICE_DETAIL.tax10<>'')"
        strMCCMaterial += " s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"

        strMCCMaterial += " ) final left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Loc_Code  left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.cust_Code   LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        'strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.shipment_No  ,final.Item_Code ,final.Loc_Code  ,final.cust_Code  ,final.shipped_Qty  ,final.TotalTaxAmt  ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.price  ,final.ItemAmt  ,final.DiscountPer,final.TotalDiscountAmt  ,final.Amt_less_Discount   ,final.Amt_Less_Discount ,final.Doc_Amt,final.[Scheme Amount],Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name] , [Delivery No]  , [Shipment No] ,MRP ,  [Scheme Code] , [Cash Scheme Code] ,  [Cash Scheme Amount], [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modify_By,final.Modify_date,final.RATE_UOM,final. Conv_Factor " ' ," + strPivotFoGrouprOuterQuery + " 
        '' query for return

        Dim strSDRCommonQuery As String = ""
        'strSDRCommonQuery = " select (CASE WHEN TSPL_SD_SALE_RETURN_HEAD.Trans_Type='ALL' THEN 'SDR' ELSE TSPL_SD_SALE_RETURN_HEAD.Trans_Type+'R' END) as Trans_Type,TSPL_SD_SALE_RETURN_HEAD.Status ,TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location, " & _
        '                    " TSPL_SD_SALE_RETURN_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_RETURN_HEAD.Document_Type,TSPL_SD_SALE_RETURN_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_RETURN_HEAD.Document_Code , " & _
        '                    " convert(varchar,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_RETURN_DETAIL.Item_Code,TSPL_SD_SALE_RETURN_DETAIL.Line_No , " & _
        '                    " -TSPL_SD_SALE_RETURN_DETAIL.Qty as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code ,TSPL_SD_SALE_RETURN_DETAIL.Item_Cost , " & _
        '                    " -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Amount,0) as Amount ,TSPL_SD_SALE_RETURN_DETAIL.Disc_Per ,case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)=0 then -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)  + case when coalesce(FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then 1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end else -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end as Disc_Amt,case when coalesce(FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then -1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) end  as [Scheme Amount] , " & _
        '                    " -(Amount-Total_Disc_Amt+case when coalesce(FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end)  as Amt_Less_Discount , " & _
        '                    " -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0) as Total_Tax_Amt ,-(Amount+coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)-coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0))  as Total_Amt,TSPL_SD_SALE_RETURN_HEAD.Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,-(case when TSPL_SD_SALE_RETURN_DETAIL.Line_No=1 then (coalesce(TSPL_SD_SALE_RETURN_HEAD.Total_Add_Charge,0)+coalesce(TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount,0)) else 0 end) as Additional_Charge, " & _
        '                    " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],-TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," & _
        '                    " -TSPL_Customer_Invoice_Head.Discount_Amount-coalesce(TSPL_SD_SALE_RETURN_HEAD.headDisc_AMt,0) as [AR Document Discount Amt], " & _
        '                    " -TSPL_Customer_Invoice_Head.amount_less_Discount as [AR Amount Before Tax],-TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " & _
        '                    " -(TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge], " & _
        '                    " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " & _
        '                    " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_RETURN_HEAD.GRNo as [GR No],'' as [WayBill No],'' as [Transporter Name],TSPL_SD_SALE_RETURN_DETAIL.Delivery_Code as [Delivery No],Against_Shipment_No as [Shipment No],TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Booking_No as [Booking No],TSPL_SD_SALE_RETURN_DETAIL.MRP, "

        strSDRCommonQuery = " select (CASE WHEN TSPL_SD_SALE_RETURN_HEAD.Trans_Type='ALL' THEN 'SDR' ELSE TSPL_SD_SALE_RETURN_HEAD.Trans_Type+'R' END) as Trans_Type,TSPL_SD_SALE_RETURN_HEAD.Status ,TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location, " &
                          " TSPL_SD_SALE_RETURN_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd,COALESCE(TSPL_SD_SALE_RETURN_HEAD.Document_Type,TSPL_SD_SALE_RETURN_HEAD.Invoice_Type) AS Invoice_Type,TSPL_SD_SALE_RETURN_HEAD.Document_Code , " &
                          " convert(varchar,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103 ) as Document_Date , TSPL_SD_SALE_RETURN_DETAIL.Item_Code,TSPL_SD_SALE_RETURN_DETAIL.Line_No , " &
                          " -TSPL_SD_SALE_RETURN_DETAIL.Qty as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code ,TSPL_SD_SALE_RETURN_DETAIL.Item_Cost , " &
                          " -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Amount,0) as Amount ,TSPL_SD_SALE_RETURN_DETAIL.Disc_Per ,case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)=0 then -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0)  + case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then 1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end else -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end as Disc_Amt,case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then -1*coalesce(Item_Net_Amt,0)*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) end  as [Scheme Amount] , " &
                          " -(Amount- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then Total_Disc_Amt else Total_Disc_Amt end  + case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type<>'FS' then case when coalesce(TSPL_SD_SALE_RETURN_DETAIL.FOC_Item,0)=1 or coalesce(TSPL_SD_SALE_RETURN_DETAIL.sampling,0)=1  then Item_Net_Amt*(case when coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_RETURN_Head.convrate,0) end) else 0 end else 0 end)  as Amt_Less_Discount , " &
                          " -coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0) as Total_Tax_Amt ,-(Amount+coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,0)- case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS' then 0 else coalesce(TSPL_SD_SALE_RETURN_DETAIL.Total_Disc_Amt,0) end )  as Total_Amt,TSPL_SD_SALE_RETURN_HEAD.Vehicle_Code,TSPL_VEHICLE_MASTER.Number as Vehicle_No,-(case when TSPL_SD_SALE_RETURN_DETAIL.Line_No=1 then (coalesce(TSPL_SD_SALE_RETURN_HEAD.Total_Add_Charge,0)+coalesce(TSPL_SD_SALE_RETURN_HEAD.RoundOffAmount,0)) else 0 end) as Additional_Charge, " &
                          " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],-TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                          " -TSPL_Customer_Invoice_Head.Discount_Amount-coalesce(TSPL_SD_SALE_RETURN_HEAD.headDisc_AMt,0) as [AR Document Discount Amt], " &
                          " -TSPL_Customer_Invoice_Head.amount_less_Discount as [AR Amount Before Tax],-TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                          " -(TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge], " &
                          " TSPL_Customer_Invoice_Head.Against_Sale_No,TSPL_Customer_Invoice_Head.Against_Sale_Return_No,TSPL_Customer_Invoice_Head.AgainstScrap, " &
                          " TSPL_Customer_Invoice_Head.Against_VCGL,TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,TSPL_SD_SALE_RETURN_HEAD.GRNo as [GR No],TSPL_SD_SHIPMENT_HEAD.gr_date as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],TSPL_SD_SALE_RETURN_DETAIL.Delivery_Code as [Delivery No],Against_Shipment_No as [Shipment No],

TSPL_SD_SALE_RETURN_DETAIL.MRP, TSPL_SD_SALE_RETURN_DETAIL.Scheme_Code ,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Type,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Code ,TSPL_SD_SALE_RETURN_DETAIL.Cash_Scheme_Amount*(-1) as Cash_Scheme_Amount ,TSPL_SD_SALE_RETURN_DETAIL.Price_code ,'' as Ref_Doc_no,NULL as Ref_doc_date,'' as Created_By ,'' as Modify_By,NULL as Modify_date,NULL as RATE_UOM,0 as Conv_Factor, TSPL_SD_SALE_RETURN_DETAIL.Sampling,TSPL_SD_SALE_RETURN_DETAIL.Scheme_Item,"
        strSDEndQry = " from TSPL_SD_SALE_RETURN_DETAIL " &
                            " left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code =TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE " &
                            " left join TSPL_VEHICLE_MASTER on TSPL_SD_SALE_RETURN_HEAD.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                            " left join TSPL_Customer_Invoice_Head on TSPL_SD_SALE_RETURN_HEAD.Document_Code=  case when len(isnull(TSPL_Customer_Invoice_Head.Against_Sale_Return_No,''))>0  then TSPL_Customer_Invoice_Head.Against_Sale_Return_No else TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return end " &
                            " left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_RETURN_DETAIL.Invoice_Code " &
                            " left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                            "  " &
                            " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_RETURN_HEAD.Customer_Code WHERE 2=2 " &
                            " AND (case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='FS' then 'Fresh Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='CSA' then 'CSA Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='PS' then 'Product Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' then 'MCC Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='EXP' then 'Export Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='Bulk Sale' then 'Bulk Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type ='SS' then 'Misc Sale Return' when TSPL_SD_SALE_RETURN_HEAD.Trans_Type in ('SD','All') then 'General Sale Return' else TSPL_SD_SALE_RETURN_HEAD.trans_Type  end) in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                            " and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_RETURN_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If

        If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
            strSDEndQry += " and tspl_customer_master.cust_category_code in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
        End If

        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_SD_SALE_RETURN_HEAD.Document_Code = '" & obj.Document_Code & "' "
        End If

        strMCCMaterial += " union all "

        strMCCMaterial += " select case when Trans_Type ='FSR' then 'Fresh Sale Return' when Trans_Type ='CSAR' then 'CSA Sale Return' when Trans_Type='PSR' then 'Product Sale Return' when Trans_Type='MCCR' then 'MCC Sale Return' when Trans_Type='EXPR' then 'Export Sale Return'when Trans_Type='Bulk Sale' then 'Bulk Sale Return' when Trans_Type ='SSR' then 'Misc Sale' when Trans_Type ='SDR' then 'General Sale Return' else trans_Type  end  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_Code as [Document No],final.Document_Date as [Document_date],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.Customer_Code as [Customer Code],MAX(final.CustAdd) As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Customer_Name) as [Customer Name] ,max(TSPL_CUSTOMER_MASTER .Parent_Customer_No) as [Parent Customer No] ,max(Parent_Master.Cust_Code) as [Parent Customer Code],max(Parent_Master.Customer_Name) as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount], " &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]  - case when (Trans_Type ='FSR' or Trans_Type ='PSR') and [AR Document Amt]>0 then coalesce(final.[Scheme Amount],0) else 0 end ) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, Scheme_Code ,Cash_Scheme_Code , Cash_Scheme_Amount ,final.Price_code,final.Ref_Doc_no,final.Ref_doc_date ,final.Created_By ,final.Modify_By,final.Modify_date,final.RATE_UOM ,final.Conv_Factor ,final. Sampling,final.Scheme_Item from ( "

        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        '' query for no tax applied
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and (coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax1,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax3,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax5,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax7,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax9,'')='' and coalesce(TSPL_SD_SALE_RETURN_DETAIL.tax10,'')='') )t "

        strMCCMaterial += "   union all"
        '' query for tax1 applied
        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX1 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt as TAX1_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX1+'%' as tax1rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"

        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX2 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt as TAX2_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX2+'%' as tax2rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX3 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt as TAX3_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX3+'%' as tax3rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX4 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt as TAX4_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX4+'%' as tax4rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX5 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt as TAX5_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX5_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX5+'%' as tax5rate  "

        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX6 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt as TAX6_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX6_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX6+'%' as tax6rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX7 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt as TAX7_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX7+'%' as tax7rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX8 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt as TAX8_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX8_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX8+'%' as tax8rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX9 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt as TAX9_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX9_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX9+'%' as tax9rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_SD_SALE_RETURN_DETAIL.TAX10 ,-TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt as TAX10_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX10_Rate ,TSPL_SD_SALE_RETURN_DETAIL.TAX10+'%' as tax10rate  "
        strMCCMaterial += " select * from (" & strSDRCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_SD_SALE_RETURN_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " )final"
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
        strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =final.Customer_Code "
        strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        'strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_Code ,final.Item_Code,final.Line_No ,final.Bill_To_Location ,final.Customer_Code ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP , Scheme_Code ,Cash_Scheme_Code , Cash_Scheme_Amount ,final.Price_code ,final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modify_By,final.Modify_date ,final.RATE_UOM,final.Conv_Factor ,final. Sampling,final.Scheme_Item " '', " + strPivotFoGrouprOuterQuery + " 

        strMCCMaterial += " union all "

        '''' bulk sale return 
        strMCCMaterial += "  select 'Bulk Sale Return' as Trans_Type ,TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code as Bill_To_Location,TSPL_SALE_RETURN_MASTER_BULKSALE.Posted,TSPL_LOCATION_MASTER.Location_Desc ,'Invoice' as Invoice_type ,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No as Document_code ,convert(varchar,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) Document_Date,'' as Vehicle_Code,'' as Vehicle_No,coalesce(-1 * TSPL_SALE_RETURN_MASTER_BULKSALE.roundoffamount,0) as Additional_Charge, " &
                           " TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code,TSPL_CUSTOMER_MASTER.Add1 + ' ' + TSPL_CUSTOMER_MASTER.Add2 + ' ' + TSPL_CUSTOMER_MASTER.Add3 As CustAdd ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Parent_Customer_No," &
                           " Parent_Master.Cust_Code as Parent_Customer_Code,Parent_Master.Customer_Name as Parent_Cust_Name ," &
                           " TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code,tspl_item_master.Item_Desc ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceQty as Qty ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceRate as Item_cost,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceFatPer ,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceSNFPer ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceFatKG as InvoiceFatKG ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceSNFKG as InvoiceSNFKG  ,-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as Amount,0 as Disc_per,0 as Disc_Amt,0 as [Scheme Amount],-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as Amt_less_Discount,0 as Total_tax_amt " + strPivotForOuterQueryforBulk + " " + strDoublePivotForOuterQueryforBulk + ",-TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount as Total_Amt, " &
                           " TSPL_Customer_Invoice_Head.Document_No as [AR Document No],-1 * TSPL_Customer_Invoice_Head.Document_Total [AR Document Amt]," &
                           " TSPL_Customer_Invoice_Head.Discount_Amount as [AR Document Discount Amt], " &
                           " -1 * (TSPL_Customer_Invoice_Head.Document_Total-TSPL_Customer_Invoice_Head.total_tax-TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Amount Before Tax],TSPL_Customer_Invoice_Head.total_tax as [AR Total Tax], " &
                           " (TSPL_Customer_Invoice_Head.total_Add_Charge+TSPL_Customer_Invoice_Head.RoundOffAmount) as [AR Total Add Charge],'' as [GR No],NULL as [GR Date],'' as [WayBill No],'' as [Transporter Code],'' as [Transporter Name],'' as [Delivery No]  ,'' as  [Shipment No],0 AS MRP, '' as [Scheme Code] ,'' as  [Cash Scheme Code] , 0 as [Cash Scheme Amount], '' as [Price Code],'' as Ref_Doc_no,NULL as Ref_doc_date ,'' as Created_By ,'' as Modified_By,NULL as Modify_date,NULL as RATE_UOM,0 as Conv_Factor,0 as Sampling,'N' as Scheme_Item" &
                           " from TSPL_SALE_RETURN_DETAIL_BULKSALE "

        strMCCMaterial += " left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No "
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code"
        strMCCMaterial += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Cust_Code =TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code"
        strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No"
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_SALE_RETURN_DETAIL_BULKSALE.Item_Code"
        strMCCMaterial += " left join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_Return_No=TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No  where 2=2 " &
                            " and 'Bulk Sale Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                            " and convert(date,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        Dim strTranferCommonQuery As String = ""
        strTranferCommonQuery = " select 'Transfer' as Trans_Type,TSPL_TRANSFER_ORDER_HEAD.Status ,TSPL_TRANSFER_ORDER_HEAD.From_Location as Bill_To_Location, " &
                                " GIT_Main_Loc.Location_Code as To_Location,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type as Invoice_Type,TSPL_TRANSFER_ORDER_HEAD.Document_No , " &
                                " convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103 ) as Document_Date , TSPL_TRANSFER_ORDER_DETAIL.Item_Code , " &
                                " (case when TSPL_TRANSFER_ORDER_DETAIL.In_Qty>0 then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_TRANSFER_ORDER_DETAIL.Out_Qty end) as Qty, " &
                                " TSPL_TRANSFER_ORDER_DETAIL.Unit_code ,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost , " &
                                " TSPL_TRANSFER_ORDER_DETAIL.Amount ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt ,0 as [Scheme Amount], " &
                                " (TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Amt_Less_Discount ,0 as Total_Tax_Amt ,(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Total_Amt,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code,COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No) as Vehicle_No,0 as Additional_Charge, " &
                                " '' as [AR Document No],0 as  [AR Document Amt],0 as [AR Document Discount Amt],0 as  [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_TRANSFER_ORDER_HEAD.GR_No as [GR No],TSPL_TRANSFER_ORDER_HEAD.gr_date as [GR Date],TSPL_TRANSFER_ORDER_HEAD.WayBill_No as [WayBill No],TSPL_TRANSFER_ORDER_HEAD.Transport_Id as [Transporter Code],case when len(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual) > 0 then TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual  else TSPL_TRANSPORT_MASTER.Transporter_Name end as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP, '' as [Scheme Code] ,'' as  [Cash Scheme Code] , 0 as [Cash Scheme Amount], '' as [Price Code],'' as Ref_Doc_no,NULL as Ref_doc_date,'' as Created_By,'' as Modified_By,NULL as Modify_date,NULL as RATE_UOM,0 as Conv_Factor, "

        strSDEndQry = " from TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No " &
                      " left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_Head.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                      " left join TSPL_TRANSPORT_MASTER on TSPL_TRANSFER_ORDER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                      " left join ( select  max(Location_Code) as Location_Code,GIT_Location from TSPL_LOCATION_MASTER where GIT_Location is not null " &
                      " group by GIT_Location ) GIT_Main_Loc on TSPL_TRANSFER_ORDER_HEAD.To_Location=GIT_Main_Loc.GIT_Location  where Transfer_Type <>'I' " &
                      " AND 'Transfer' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                      " and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) "

        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_ORDER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_ORDER_HEAD.From_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and GIT_Main_Loc.Location_Code in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If


        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = '" & obj.Document_Code & "' "
        End If
        strMCCMaterial += " union all "
        strMCCMaterial += "   select *,0 as Sampling,'N' as Scheme_Item from (Select 'MCC Transfer' as Trans_Type ,TSPL_MCC_Dispatch_Challan.MCC_Code as  Bill_To_Location,TSPL_MCC_Dispatch_Challan.isPosted as " _
            & " Status, sendr.location_desc as  location_desc,'MCC Transfer' as Invoice_Type,TSPL_MCC_Dispatch_Challan.Chalan_NO as PI_NO , " _
            & " convert(varchar,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103 ) as PI_Date, '' as vehicledesc,tm.tanker_NO as Vehicle_No,0  as Additional_Charge , " _
            & " tspl_milk_Transfer_In.location_Code as Customer_Code,'' AS [CustAdd],  recv.Location_Desc  as Customer_Name ,'' as [Parent Vendor No],'' as [Parent Vendor Code]," _
            & " '' as [Parent Vendor Name], TSPL_MCC_Dispatch_Challan.Item_Code, TSPL_MCC_Dispatch_Challan.Item_Desc , TSPL_MCC_Dispatch_Challan.Net_Qty  as Qty " _
            & " ,TSPL_MCC_Dispatch_Challan.UOM_Code as  Unit_code ,round(((TSPL_MCC_Dispatch_Challan.FAT_RATE *(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) " _
            & " * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)) +(TSPL_MCC_Dispatch_Challan.SNF_RATE  *(coalesce(cast(t_SNF_Recd.Param_Field_Value as float),0) * " _
            & " coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100)))/coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,1) ,2) as  Item_Cost ,  t_FAT_Recd.Param_Field_Value as [FAT Per]" _
            & " ,t_SNF_Recd.Param_Field_Value as [SNF Per],(coalesce(cast(t_FAT_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) " _
            & " as [FAT KG],(coalesce(cast(t_Snf_Recd.Param_Field_Value as float),0) * coalesce(TSPL_MCC_Dispatch_Challan.Net_Qty,0)/100) as [SNF KG],amount as  Amount ,0" _
            & " as Disc_Per  ,0 as Disc_Amt,0 as [Scheme Amount] ,  amount as  Amt_Less_Discount   " & strPivotForTransfer_In & "" & strPivotFortRANSFER_INPercentQuery & ",   0 as Total_Tax_Amt " _
            & " ,amount as   Total_Amt,TSPL_vendor_Invoice_Head.Document_No as [AR Document No],TSPL_vendor_Invoice_Head.Document_Total [AR Document Amt]" _
            & " ,TSPL_vendor_Invoice_Head.Discount_Amount as [AR Document Discount Amt],TSPL_vendor_Invoice_Head.amount_less_Discount as [AR Amount Before Tax]," _
            & " TSPL_vendor_Invoice_Head.total_tax as [AR Total Tax],TSPL_vendor_Invoice_Head.total_Add_Charge as [AR Total Add Charge],'' as [GRNO],null as [GRN Date],'' as Way_BillNo," _
            & "  TSPL_MCC_Dispatch_Challan.tanker_No,tm.description,'' as  [delivery No],'' as [Shiping No],0 AS MRP,'' as  [Scheme Code] ,'' as [Cash Scheme Code] , 0 as [Cash Scheme Amount],'' as [Price Code],'' as Ref_Doc_no,NULL as Ref_doc_date,'' as Created_By ,'' as Modified_By,NULL as Modify_date,NULL as RATE_UOM,0 as Conv_Factor from TSPL_MCC_Dispatch_Challan  left outer  join TSPL_MILK_TRANSFER_IN on TSPL_MILK_TRANSFER_IN.Dispatch_Challan_No " _
            & " =TSPL_MCC_Dispatch_Challan.Chalan_NO  LEFT JOIN tspl_location_Master  sendr ON sendr.Location_Code=TSPL_MCC_Dispatch_Challan.MCC_CODE " _
            & " left join tspl_Location_master on tspl_Location_master.location_code=TSPL_MCC_Dispatch_Challan.mcc_Code left join tspl_Location_master recv on " _
            & " recv.location_code=TSPL_MILK_TRANSFER_IN.Location_Code left join TSPL_vendor_Invoice_Head on  TSPL_vendor_Invoice_Head.vendor_Invoice_No" _
            & " =TSPL_MILK_TRANSFER_IN.Receipt_Challan_No left join tspl_tanker_Master tm on tm.tanker_no=TSPL_MCC_Dispatch_Challan.tanker_No Left Outer Join " _
            & " (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.* From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On " _
            & " TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'SNF') " _
            & " t_SNF_Recd On t_SNF_Recd.Chalan_NO   = TSPL_MCC_Dispatch_Challan.Chalan_NO   Left Outer Join (Select TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.*" _
            & " From TSPL_MCC_Dispatch_Challan Left Outer Join TSPL_Mcc_Dispatch_Chalan_Parameter_Detail On TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Chalan_No  = " _
            & " TSPL_MCC_Dispatch_Challan.Chalan_NO  where TSPL_Mcc_Dispatch_Chalan_Parameter_Detail.Param_Type = 'FAT' ) t_FAT_Recd On t_FAT_Recd.Chalan_No " _
            & " = TSPL_MCC_Dispatch_Challan.Chalan_NO where 2=2 " &
              " AND 'MCC Transfer' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
              " and convert(date,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_MCC_Dispatch_Challan.Dispatch_Date,103) <= convert(date,('" & To_Date & "'),103) )t "
        '' transaction unit conversion
        strMCCMaterial += " union all "
        strMCCMaterial += " select Trans_type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER .Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_No as [Document No],final.Document_Date as [Document_date],Vehicle_Code,Vehicle_No,final.Additional_Charge,final.To_Location as [Customer Code],'' As [Customer Address] ,max(TSPL_CUSTOMER_MASTER .Location_Desc) as [Customer Name] ,'' as [Parent Customer No] ,'' as [Parent Customer Code],'' as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount]," &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],final.[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,  [Scheme Code] , [Cash Scheme Code] ,  [Cash Scheme Amount],  [Price Code],final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modified_By,final.Modify_date,final.RATE_UOM,final.Conv_Factor,0 as Sampling,'N' as Scheme_Item from ( "
        'strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,0 as TAX1_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "

        '' query for no tax applied
        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and (coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax1,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax3,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax5,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax7,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax9,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax10,'')='') )t "
        '" pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += " union all "
        '' quert for no tax applied
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX2 ,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX2+'%' as tax2rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX3 ,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX3+'%' as tax3rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX4 ,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4+'%' as tax4rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX5 ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5+'%' as tax5rate  "

        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX6 ,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX6+'%' as tax6rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX7 ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7+'%' as tax7rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX8 ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8+'%' as tax8rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX9 ,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9+'%' as tax9rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX10 ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10+'%' as tax10rate  "
        strMCCMaterial += " select * from (" & strTranferCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t"

        strMCCMaterial += " )final"
        ''-------------------------
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location "
        strMCCMaterial += " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code "
        strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location "
        'strMCCMaterial += " left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'strMCCMaterial += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER as Parent_Master ON Parent_Master.Cust_Code=TSPL_CUSTOMER_MASTER.Parent_Customer_No "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_No ,final.Item_Code ,final.Bill_To_Location ,final.To_Location ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,[Scheme Code] ,[Cash Scheme Code] ,[Cash Scheme Amount], [Price Code] ,final.Ref_Doc_no,final.Ref_doc_date,final.Created_By ,final.Modified_By,final.Modify_date,final. RATE_UOM,final.Conv_Factor"
        ''richa

        strMCCMaterial += Environment.NewLine + " --- QUERY FOR TRANSFER RETURN---------------------- ADDED BY RICHA AGARWAL ------------" + Environment.NewLine


        Dim strTranferReturnCommonQuery As String = ""
        strTranferReturnCommonQuery = " select 'Transfer Return' as Trans_Type,TSPL_TRANSFER_ORDER_HEAD.Status ,  TSPL_TRANSFER_ORDER_HEAD.From_Location as Bill_To_Location, " &
                                " GIT_Main_Loc.Location_Code as To_Location,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type as Invoice_Type,TSPL_TRANSFER_RETURN.Document_No , " &
                                " convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103 ) as Document_Date , TSPL_TRANSFER_ORDER_DETAIL.Item_Code , " &
                                " -(case when TSPL_TRANSFER_ORDER_DETAIL.In_Qty>0 then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_TRANSFER_ORDER_DETAIL.Out_Qty end) as Qty, " &
                                " TSPL_TRANSFER_ORDER_DETAIL.Unit_code ,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost , " &
                                " -TSPL_TRANSFER_ORDER_DETAIL.Amount AS Amount ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per ,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt ,0 as [Scheme Amount], " &
                                " -(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Amt_Less_Discount ,0 as Total_Tax_Amt ,-(TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt) as Total_Amt,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code,COALESCE(TSPL_VEHICLE_MASTER.Number,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No) as Vehicle_No,0 as Additional_Charge, " &
                                " '' as [AR Document No],0 as  [AR Document Amt],0 as [AR Document Discount Amt],0 as  [AR Amount Before Tax],0 as [AR Total Tax],0 as [AR Total Add Charge],TSPL_TRANSFER_ORDER_HEAD.GR_No as [GR No],TSPL_TRANSFER_ORDER_HEAD.gr_date as [GR Date],TSPL_TRANSFER_ORDER_HEAD.WayBill_No as [WayBill No],TSPL_TRANSFER_ORDER_HEAD.Transport_Id as [Transporter Code],case when len(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual) > 0 then TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual  else TSPL_TRANSPORT_MASTER.Transporter_Name end as [Transporter Name],'' as [Delivery No]  ,'' as [Shipment No],0 AS MRP, '' as [Scheme Code] , '' as [Cash Scheme Code] ,0 as   [Cash Scheme Amount],  '' as [Price Code],'' as Ref_Doc_no,NULL as Ref_doc_date,'' as Created_By ,'' as Modified_By,NULL as Modify_date,NULL as RATE_UOM,0 as Conv_Factor,"

        strSDEndQry = " from TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No " &
            " Left Outer Join TSPL_TRANSFER_RETURN on TSPL_TRANSFER_RETURN.Transfer_No=TSPL_TRANSFER_ORDER_HEAD.Document_No " &
                      " left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_Head.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                      " left join TSPL_TRANSPORT_MASTER on TSPL_TRANSFER_ORDER_HEAD.Transport_Id=TSPL_TRANSPORT_MASTER.Transport_Id " &
                      " left join ( select  max(Location_Code) as Location_Code,GIT_Location from TSPL_LOCATION_MASTER where GIT_Location is not null " &
                      " group by GIT_Location ) GIT_Main_Loc on TSPL_TRANSFER_ORDER_HEAD.To_Location=GIT_Main_Loc.GIT_Location  where TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ='O'  " &
                      " AND 'Transfer Return' in (" & clsCommon.GetMulcallString(obj.Trans_Type_List) & ") " &
                      " and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= convert(date,('" & To_Date & "'),103) AND ISNULL(TSPL_TRANSFER_RETURN.Document_No ,'')<>'' "

        '' filter conditions
        If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_ORDER_DETAIL.Item_Code in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        End If
        If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_ORDER_HEAD.From_Location in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        End If

        If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
            strSDEndQry += " and  GIT_Main_Loc.Location_Code  in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        End If
        If clsCommon.myLen(obj.Document_Code) > 0 Then
            strSDEndQry += " and TSPL_TRANSFER_RETURN.Document_No = '" & obj.Document_Code & "' "
        End If

        strMCCMaterial += " union all "
        strMCCMaterial += " select Trans_type  as [Trans Type],final.Bill_To_Location as [Location Code],final.Status  ,max(TSPL_LOCATION_MASTER.Location_Desc) as [Location Name] ,(final.Invoice_Type) as [Invoice Type],final.Document_No as [Document No],final.Document_Date as [Document_date],Vehicle_Code,Vehicle_No,final.Additional_Charge, final.To_Location as [Customer Code],'' As [Customer Address] ,max(TSPL_CUSTOMER_MASTER.Location_Desc) as [Customer Name] ,'' as [Parent Customer No] ,'' as [Parent Customer Code],'' as [Parent Customer Name], final.Item_Code as [Item Code],max(tspl_item_master.Item_Desc) as [Item Name],final.Qty as [Quantity],final.Unit_code as [UOM],final.Item_Cost as [Item Cost], 0 as [Fat Per],0 as [SNF Per],0 as [Fat Kg],0 as [SNF KG],final.Amount,final.Disc_Per as [Discount Per],final.Disc_Amt as [Discount Amount],final.[Scheme Amount] as [Scheme Amount],final.Amt_Less_Discount  as [Amount Less Discount] " + strPivotForOuterQuery + ", " + strPivotFoGrouprOuterQuery + " ,final.Total_Tax_Amt as [Total Tax Amount],final.Total_Amt as [Total Amount]," &
            " [AR Document No], [AR Document Amt],[AR Document Discount Amt],([AR Document Amt]-[AR Total Tax]-[AR Total Add Charge]) as  [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],[GR No],[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP, [Scheme Code] , [Cash Scheme Code] ,  [Cash Scheme Amount],  [Price Code],final.Ref_Doc_no,final. Ref_doc_date ,final.Created_By ,final.Modified_By,final.Modify_date,final. RATE_UOM,final. Conv_Factor,0 as Sampling,'N' as Scheme_Item from ( "
        'strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,0 as TAX1_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "

        '' query for no tax applied
        strTaxColumns = strPivotForInnerQueryNoTax & "," & strDoublePivotForInnerQueryNoTax
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and (coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax1,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax2,'')='' " &
                          " and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax3,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax4,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax5,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax6,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax7,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax8,'')='' and " &
                          " coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax9,'')='' and coalesce(TSPL_TRANSFER_ORDER_DETAIL.tax10,'')='') )t "

        strMCCMaterial += Environment.NewLine + " union all "
        '' quert for no tax applied
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX1 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX1_AMT AS TAX1_AMT ,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX1+'%' as tax1rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax1<>'' )s pivot (sum(tax1_amt) for tax1 in (" + strPivotForInnerQuery + "))t pivot (min(tax1_rate) for tax1rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX2 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt AS TAX2_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX2+'%' as tax2rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax2<>'' )s pivot (sum(tax2_amt) for tax2 in (" + strPivotForInnerQuery + "))t pivot (min(tax2_rate) for tax2rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX3 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt AS TAX3_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX3+'%' as tax3rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax3<>'' )s pivot (sum(tax3_amt) for tax3 in (" + strPivotForInnerQuery + "))t pivot (min(tax3_rate) for tax3rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "   union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX4 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt AS TAX4_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4+'%' as tax4rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax4<>'' )s pivot (sum(tax4_amt) for tax4 in (" + strPivotForInnerQuery + "))t pivot (min(tax4_rate) for tax4rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"
        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX5 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt AS TAX5_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5+'%' as tax5rate  "

        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax5<>'' )s pivot (sum(tax5_amt) for tax5 in (" + strPivotForInnerQuery + "))t pivot (min(tax5_rate) for tax5rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX6 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt AS TAX6_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX6+'%' as tax6rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax6<>'')s pivot (sum(tax6_amt) for tax6 in (" + strPivotForInnerQuery + "))t pivot (min(tax6_rate) for tax6rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX7 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt AS TAX7_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7+'%' as tax7rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax7<>'' )s pivot (sum(tax7_amt) for tax7 in (" + strPivotForInnerQuery + "))t pivot (min(tax7_rate) for tax7rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX8 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt AS TAX8_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8+'%' as tax8rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax8<>'' )s pivot (sum(tax8_amt) for tax8 in (" + strPivotForInnerQuery + "))t pivot (min(tax8_rate) for tax8rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX9 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt AS TAX9_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9+'%' as tax9rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax9<>'')s pivot (sum(tax9_amt) for tax9 in (" + strPivotForInnerQuery + "))t pivot (min(tax9_rate) for tax9rate in (" + strDoublePivotForInnerQuery + "))t"
        strMCCMaterial += "  union all"

        strTaxColumns = " TSPL_TRANSFER_ORDER_DETAIL.TAX10 ,-TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt AS TAX10_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10+'%' as tax10rate  "
        strMCCMaterial += " select * from (" & strTranferReturnCommonQuery & strTaxColumns & strSDEndQry & " and TSPL_TRANSFER_ORDER_DETAIL.tax10<>'' )s pivot (sum(tax10_amt) for tax10 in (" + strPivotForInnerQuery + "))t pivot (min(tax10_rate) for tax10rate in (" + strDoublePivotForInnerQuery + "))t" &
         " )final" &
        " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =final.Bill_To_Location " &
         " left outer join tspl_item_master on tspl_item_master.Item_Code =final.Item_Code " &
         " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location "
        '" left outer join " & "(" & qryQC & ") as QC" & " on QC.Item_Code =final.Item_Code "
        'added by stuti on 01/05/2017
        strMCCMaterial += " where convert(date,final.Document_Date,103) >= convert(date,('" & From_Date & "'),103) and convert(date,final.Document_Date,103) <= convert(date,('" & To_Date & "'),103)"

        strMCCMaterial += " group by  final.Trans_Type,final .Status  ,final.Document_No ,final.Item_Code ,final.Bill_To_Location ,final.To_Location ,final.Qty ,final.Total_Tax_Amt ,final.Invoice_Type ,final.Document_Date ,final.Unit_code ,final.Item_Cost ,final.Amount ,final.Disc_Per ,final.Disc_Amt,final.[Scheme Amount] ,final.Amt_Less_Discount ,final.Total_Amt,Vehicle_Code,Vehicle_No,final.Additional_Charge,[AR Document No], [AR Document Amt],[AR Document Discount Amt], [AR Amount Before Tax],[AR Total Tax],[AR Total Add Charge],[GR No],final.[GR Date],[WayBill No],final.[Transporter Code],[Transporter Name], [Delivery No]  , [Shipment No],MRP,  [Scheme Code] ,  [Cash Scheme Code] , [Cash Scheme Amount],  [Price Code] ,final.Ref_Doc_no,final. Ref_doc_date,final.Created_By ,final.Modified_By,final.Modify_date,final. RATE_UOM,final.Conv_Factor" + Environment.NewLine

        strMCCMaterial += Environment.NewLine + " --- QUERY FOR TRANSFER RETURN END---------------------- ADDED BY RICHA AGARWAL ------------" + Environment.NewLine


        strMCCMaterial += ") xx"
        '===============Added by preeti Gupta ===
        strMCCMaterial += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =xx.[Location Code] "
        strMCCMaterial += " left join TSPL_USER_MASTER as Created_By on Created_By.user_Code =xx.Created_By "
        strMCCMaterial += " left join TSPL_USER_MASTER as Modify_By on Modify_By.user_Code =xx.Modify_By "
        '=======================================
        ''richa agarwal change to show csa sales account for csa sale and csa sale return
        '' strMCCMaterial += " left outer join (select ITEM.Item_Code,ITEM.Item_Desc,ITEM.Structure_Code,SA.Sales_Account from TSPL_ITEM_MASTER Item left join TSPL_SALES_ACCOUNTS SA on Item.Sale_Class_Code=SA.Sales_Class_Code) Item on Item.Item_Code =xx.[Item Code] "
        strMCCMaterial += " left outer join (select ITEM.Item_Code,ITEM.Item_Desc,ITEM.Structure_Code,SA.Sales_Account, isnull(CA.GSOC_Acct,'') as GSOC_Acct from TSPL_ITEM_MASTER Item left join TSPL_SALES_ACCOUNTS SA on Item.Sale_Class_Code=SA.Sales_Class_Code left join TSPL_CUSTOMER_ACCOUNT_SET  CA on Item.Cust_Account =CA.Cust_Account ) Item on Item.Item_Code =xx.[Item Code] "
        ''--------------------------
        '' transaction unit conversion
        strMCCMaterial += " inner join (" & qryTransStock & ") as  TransStock on xx.[Item Code]=TransStock.Item_Code  "
        If Not stock_uom Then
            strMCCMaterial += " and TransStock.UOM_Code=" & IIf(clsCommon.myLen(Unit_Code) <= 0, "xx.[UOM]", "'" & Unit_Code & "'") & ""
        End If
        ''end transaction unit conversion
        strMCCMaterial += " left join (" & strItemGroup & ") as Item_Group on Item.Structure_Code=Item_Group.Structure_Code "
        'strMCCMaterial += " left outer join TSPL_LOCATION_MASTER as TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .Location_Code =final.To_location "
        strMCCMaterial += " left join (" & qryStock & ") as Stock_SU on xx.[Item Code]=Stock_SU.Item_Code and xx.[UOM]=Stock_SU.UOM_Code "
        '===============================Added By Preeti Gupta===================================
        strMCCMaterial += " left join (" & qryStock & ") as Rate_Stock_SU on xx.[Item Code]=Rate_Stock_SU.Item_Code and isnull(xx.[RATE_UOM],xx.UOM) =Rate_Stock_SU.UOM_Code  "
        strMCCMaterial += " inner join (" & qryTransStock & ") as Rate_select_SU on xx.[Item Code]=Rate_select_SU.Item_Code  "
        If Not stock_uom Then
            strMCCMaterial += " and Rate_select_SU.UOM_Code=" & IIf(clsCommon.myLen(Unit_Code) <= 0, "xx.[UOM]", "'" & Unit_Code & "'") & ""
        End If
        ' ============================================================================================
        strMCCMaterial += " left join (" & qryKG & ") as StockKG on xx.[Item Code]=StockKG.Item_Code  "
        strMCCMaterial += " left join (select Cust_Code,Cust_Group_Code,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Struct_Code,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.state as State_Code,tspl_State_Master.State_Name,TSPL_CUSTOMER_MASTER.cust_category_code from TSPL_CUSTOMER_MASTER left join tspl_State_Master on tspl_State_Master.state_Code=TSPL_CUSTOMER_MASTER.state) as Cust on xx.[Customer Code]=Cust.Cust_Code  left join (select location_Code as Cust_Code,Tin_No,State_Code,State_Name from TSPL_location_MASTER left join tspl_State_Master on tspl_State_Master.state_Code=TSPL_location_MASTER.state) as Cust_Loc on xx.[Customer Code]=Cust_Loc.Cust_Code  "
        strMCCMaterial += " left join (select Cust_Group_Code,Cust_Group_Desc from TSPL_CUSTOMER_GROUP_MASTER) as Cust_Group on Cust.Cust_Group_Code=Cust_Group.Cust_Group_Code "
        strMCCMaterial += " left join (select Zone_Code,Description from TSPL_ZONE_MASTER) as Zone on Cust.Zone_Code=Zone.Zone_Code " &
                          " left join TSPL_LOCATION_MASTER as Loc on Loc.Location_Code=xx.[Location Code] "
        If clsCommon.myLen(strCategoryTable) > 0 Then
            strMCCMaterial += " left outer join (" + strCategoryTable + ") as VirtualCategoryTabel on  VirtualCategoryTabel.Item_Code=xx.[Item Code] 
 LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = xx.[Item Code] and TSPL_ITEM_UOM_DETAIL.UOM_Code = xx.uom	 
 Inner JOIN  ( select item_code,uom_code,conversion_factor from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON xx.[Item Code] = I.item_code "
        End If
        ''==========Monika
        strMCCMaterial += " left outer join (" + qryQC + ")QC on QC.item_code=xx.[Item Code] "
        ''================end here=====
        strMCCMaterial += " where 2 = 2  and  convert(date,xx.Document_Date,103) >= convert(date,('" + From_Date + "'),103) and convert(date,xx.Document_Date,103) <= convert(date,('" + To_Date + "'),103) " ' + clsCommon.myCstr(IIf(clsCommon.myLen(txtUOM.Value) > 0, "and xx.[UOM]='" + txtUOM.Value + "' ", ""))
        QryLst.Add(strMCCMaterial)


        'Dim fullpath As String = "D:\RMWebBrowser\RMWebBrowser\bin\Debug" & "\Muhil.xml"
        'Dim fileinfo As System.IO.FileInfo = New System.IO.FileInfo(fullpath)
        'Dim XMLWriter As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(fullpath, System.Text.Encoding.UTF8)
        'XMLWriter.Formatting = Xml.Formatting.Indented
        'XMLWriter.WriteStartDocument()
        'XMLWriter.WriteElementString("Name", strMCCMaterial)
        QryLst.Add(strPivotForFinalOuterQuery)
        'clsCommon.MyMessageBoxShow("DOne")
        'Return Nothing
        Return QryLst
    End Function
    Public Shared Function GetTaxQuery(ByVal lstTables As List(Of String), ByVal lstTableDocDateCols As List(Of String), ByVal obj As clsSaleRegisterParameterType) As String
        Dim qry As String = String.Empty
        If Not lstTables Is Nothing AndAlso lstTables.Count > 0 Then
            For Each TableName As String In lstTables
                If clsCommon.CompairString(TableName, "TSPL_CANSALE_INVOICE_HEAD") = CompairStringResult.Equal OrElse clsCommon.CompairString(TableName, "TSPL_INVOICE_Master_BULKSALE") = CompairStringResult.Equal Then
                    For intloop As Integer = 1 To 4
                        If clsCommon.myLen(qry) <= 0 Then
                            qry = "select distinct TAX" & intloop & " from " & TableName & " where cast(" & lstTableDocDateCols.Item(lstTables.IndexOf(TableName)) & " as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "' and coalesce(Tax" & intloop & ",'')<>''"
                        Else
                            qry = qry & " Union  " & "select distinct TAX" & intloop & " from " & TableName & " where cast(" & lstTableDocDateCols.Item(lstTables.IndexOf(TableName)) & " as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "' and coalesce(Tax" & intloop & ",'')<>''"
                        End If
                    Next
                Else
                    For intloop As Integer = 1 To 10
                        If clsCommon.myLen(qry) <= 0 Then
                            qry = "select distinct TAX" & intloop & " from " & TableName & " where cast(" & lstTableDocDateCols.Item(lstTables.IndexOf(TableName)) & " as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "' and coalesce(Tax" & intloop & ",'')<>''"
                        Else
                            qry = qry & " Union  " & "select distinct TAX" & intloop & " from " & TableName & " where cast(" & lstTableDocDateCols.Item(lstTables.IndexOf(TableName)) & " as Date) between '" & clsCommon.GetPrintDate(obj.From_Date, "dd-MMM-yyyy") & "' and '" & clsCommon.GetPrintDate(obj.To_Date, "dd-MMM-yyyy") & "' and coalesce(Tax" & intloop & ",'')<>''"
                        End If
                    Next
                End If
            Next
        Else
            Return qry
        End If
        Return qry
    End Function
    Public Shared Function GetReportData(ByVal obj As clsSaleRegisterParameterType) As DataTable
        Dim strRunQuery As String = ""
        strRunQuery = GetReportDataQuery(obj)
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(strRunQuery)
        Return dt
        'Dim strPivotForFinalOuterQuery As String = ""
        'Dim qryList As ArrayList
        'qryList = ReturnQuery(obj)
        'Dim strMain As String = qryList(0)
        'strPivotForFinalOuterQuery = qryList(1)

        'If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
        '    strMain += " and xx.[Item Code] in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
        'End If
        'If obj.Trans_Type_List IsNot Nothing AndAlso obj.Trans_Type_List.Count > 0 Then
        '    strMain += " and xx.[Trans Type] in (" + clsCommon.GetMulcallString(obj.Trans_Type_List) + ") "
        'End If
        'If obj.State_List IsNot Nothing AndAlso obj.State_List.Count > 0 Then
        '    strMain += " and Loc.State in (" + clsCommon.GetMulcallString(obj.State_List) + ") "
        'End If
        'If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
        '    strMain += " and xx.[Location Code] in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
        'End If

        'If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
        '    strMain += " and xx.[Customer Code] in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
        'End If
        'If obj.Item_Group_List IsNot Nothing AndAlso obj.Item_Group_List.Count > 0 Then
        '    strMain += " and coalesce(Item_Group.Item_Group,'') in (" + clsCommon.GetMulcallString(obj.Item_Group_List) + ") "
        'End If
        ' '' Done by Panch raj against Ticket No:BM00000007277
        'If obj.Cust_Group_Code_List IsNot Nothing AndAlso obj.Cust_Group_Code_List.Count > 0 Then
        '    strMain += " and coalesce(Cust.Cust_Group_Code,'') in (" + clsCommon.GetMulcallString(obj.Cust_Group_Code_List) + ") "
        'End If
        'If clsCommon.myLen(obj.Document_Code) > 0 Then
        '    strMain += " and xx.[Document No] = '" & obj.Document_Code & "' "
        'End If
        'If clsCommon.myLen(obj.Other_Cond) > 0 Then
        '    strMain += obj.Other_Cond
        'End If
        'If obj.ReportType = "Total Sale" Then
        '    strRunQuery = "select sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final"
        'ElseIf obj.ReportType = "Location Wise" Then
        '    strRunQuery = "select [Location Code],[Location Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name]"
        'ElseIf obj.ReportType = "Item Group Wise" Then
        '    strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description]"
        'ElseIf obj.ReportType = "Customer Group Wise" Then
        '    strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description]"
        'ElseIf obj.ReportType = "Item Wise" Then
        '    strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name]"
        'ElseIf obj.ReportType = "Customer Wise" Then
        '    strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Item Code],[Item Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name],[Customer Code],[Customer Name]"
        'ElseIf obj.ReportType = "Document Wise" Then
        '    strRunQuery = "select [Document No],[Document_date],[Trans Type],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],max([TIN No]) as [TIN No],max([GR No]) as [GR No],max([WayBill No]) as [WayBill No],max([Transporter Name]) as [Transporter Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Discount Amount]) as [Discount Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],max([AR Document No]) as [AR Document No], max([AR Document Amt]) as [AR Document Amt],max([AR Document Discount Amt]) as [AR Document Discount Amt] , case when max(coalesce([AR Amount Before Tax],0))>0 then  max([AR Amount Before Tax]) else  min([AR Amount Before Tax]) end as [AR Amount Before Tax],max([AR Total Tax]) as [AR Total Tax],max([AR Total Add Charge]) as [AR Total Add Charge] from (" & strMain & ") as Final group by [Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Trans Type] order by convert(date,[Document_Date],103),[Document No]"
        'ElseIf obj.ReportType = "Document Detail" Then
        '    strRunQuery = strMain & "order by convert(date,[Document_Date],103),[Document No]"
        'End If

    End Function
    Public Shared Function GetReportDataReaderForPrint(ByVal obj As clsSaleRegisterParameterType) As String
        Dim strRunQuery As String = ""
        strRunQuery = GetReportDataQuery(obj)
        'Dim rd As SqlDataReader = clsDBFuncationality.GetDataReader(strRunQuery)
        Dim PrintQuery As String = strRunQuery
        Return PrintQuery
    End Function
    Public Shared Function GetReportDataReader(ByVal obj As clsSaleRegisterParameterType) As SqlDataReader
        Dim strRunQuery As String = ""
        strRunQuery = GetReportDataQuery(obj)
        Dim rd As SqlDataReader = clsDBFuncationality.GetDataReader(strRunQuery)
        Return rd
    End Function
    Public Shared Function GetReportDataQuery(ByVal obj As clsSaleRegisterParameterType) As String
        Dim strRunQuery As String = ""
        Dim strPivotForFinalOuterQuery As String = ""
        Dim strCustomerAmountWiseTaxQuery As String = ""
        Dim qryList As ArrayList
        Dim qryListDocInfoLevel As ArrayList



        Dim strMainDocInfoLevel As String = Nothing

        Dim strMain As String = Nothing

        If obj.ReportType = "Document Info Level" Then
            qryListDocInfoLevel = ReturnQueryforDocumentInfoLevel(obj)
            strMainDocInfoLevel = qryListDocInfoLevel(0)
            strPivotForFinalOuterQuery = qryListDocInfoLevel(1)
        Else
            qryList = ReturnQuery(obj)
            strMain = qryList(0)
            strPivotForFinalOuterQuery = qryList(1)
            strCustomerAmountWiseTaxQuery = qryList(2)

        End If

        If obj.ReportType = "Document Info Level" Then

            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMainDocInfoLevel += " and xx.[Item Code] in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            If obj.Trans_Type_List IsNot Nothing AndAlso obj.Trans_Type_List.Count > 0 Then
                strMainDocInfoLevel += " and xx.[Trans Type] in (" + clsCommon.GetMulcallString(obj.Trans_Type_List) + ") "
            End If
            If obj.State_List IsNot Nothing AndAlso obj.State_List.Count > 0 Then
                strMainDocInfoLevel += " and Loc.State in (" + clsCommon.GetMulcallString(obj.State_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMainDocInfoLevel += " and xx.[Location Code] in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMainDocInfoLevel += " and xx.[Customer Code] in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Item_Group_List IsNot Nothing AndAlso obj.Item_Group_List.Count > 0 Then
                strMainDocInfoLevel += " and coalesce(Item_Group.Item_Group,'') in (" + clsCommon.GetMulcallString(obj.Item_Group_List) + ") "
            End If
            '' Done by Panch raj against Ticket No:BM00000007277
            If obj.Cust_Group_Code_List IsNot Nothing AndAlso obj.Cust_Group_Code_List.Count > 0 Then
                strMainDocInfoLevel += " and coalesce(Cust.Cust_Group_Code,'') in (" + clsCommon.GetMulcallString(obj.Cust_Group_Code_List) + ") "
            End If
            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMainDocInfoLevel += " and coalesce(Cust.cust_category_code,'') in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If
            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strMainDocInfoLevel += " and xx.[Document No] = '" & obj.Document_Code & "' "
            End If
            If clsCommon.myLen(obj.Other_Cond) > 0 Then
                strMainDocInfoLevel += obj.Other_Cond
            End If
        Else

            If obj.Item_Code_List IsNot Nothing AndAlso obj.Item_Code_List.Count > 0 Then
                strMain += " and xx.[Item Code] in (" + clsCommon.GetMulcallString(obj.Item_Code_List) + ") "
            End If
            'If obj.Trans_Type_List IsNot Nothing AndAlso obj.Trans_Type_List.Count > 0 Then
            '    strMain += " and xx.[Trans Type] in (" + clsCommon.GetMulcallString(obj.Trans_Type_List) + ") "
            'End If
            If obj.State_List IsNot Nothing AndAlso obj.State_List.Count > 0 Then
                strMain += " and Loc.State in (" + clsCommon.GetMulcallString(obj.State_List) + ") "
            End If
            If obj.Location_Code_List IsNot Nothing AndAlso obj.Location_Code_List.Count > 0 Then
                strMain += " and xx.[Location Code] in (" + clsCommon.GetMulcallString(obj.Location_Code_List) + ") "
            End If

            If obj.Customer_Code_List IsNot Nothing AndAlso obj.Customer_Code_List.Count > 0 Then
                strMain += " and xx.[Customer Code] in (" + clsCommon.GetMulcallString(obj.Customer_Code_List) + ") "
            End If
            If obj.Item_Group_List IsNot Nothing AndAlso obj.Item_Group_List.Count > 0 Then
                strMain += " and coalesce(Item_Group.Item_Group,'') in (" + clsCommon.GetMulcallString(obj.Item_Group_List) + ") "
            End If
            '' Done by Panch raj against Ticket No:BM00000007277
            If obj.Cust_Group_Code_List IsNot Nothing AndAlso obj.Cust_Group_Code_List.Count > 0 Then
                strMain += " and coalesce(Cust.Cust_Group_Code,'') in (" + clsCommon.GetMulcallString(obj.Cust_Group_Code_List) + ") "
            End If

            If obj.Customer_Category_List IsNot Nothing AndAlso obj.Customer_Category_List.Count > 0 Then
                strMain += " and coalesce(Cust.cust_category_code,'') in (" + clsCommon.GetMulcallString(obj.Customer_Category_List) + ") "
            End If

            If clsCommon.myLen(obj.Document_Code) > 0 Then
                strMain += " and xx.[Document No] = '" & obj.Document_Code & "' "
            End If
            If clsCommon.myLen(obj.Other_Cond) > 0 Then
                strMain += obj.Other_Cond
            End If

            If obj.Zone_List IsNot Nothing AndAlso obj.Zone_List.Count > 0 Then
                strMain += " and coalesce(Cust.Zone_Code,'') in (" + clsCommon.GetMulcallString(obj.Zone_List) + ") "
            End If

            If obj.RSM_List IsNot Nothing AndAlso obj.RSM_List.Count > 0 Then
                strMain += " and coalesce(Cust.[RSM Code],'') in (" + clsCommon.GetMulcallString(obj.RSM_List) + ") "
            End If

            If obj.ZSM_List IsNot Nothing AndAlso obj.ZSM_List.Count > 0 Then
                strMain += " and coalesce(Cust.[ZSM Code],'') in (" + clsCommon.GetMulcallString(obj.ZSM_List) + ") "
            End If

            If obj.ASM_List IsNot Nothing AndAlso obj.ASM_List.Count > 0 Then
                strMain += " and coalesce(Cust.[ASM Code],'') in (" + clsCommon.GetMulcallString(obj.ASM_List) + ") "
            End If

            If obj.ASO_List IsNot Nothing AndAlso obj.ASO_List.Count > 0 Then
                strMain += " and coalesce(Cust.[ASO Code],'') in (" + clsCommon.GetMulcallString(obj.ASO_List) + ") "
            End If


        End If

        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowParticluarColumnInSalesRegisterForGopalJee, clsFixedParameterCode.ShowParticluarColumnInSalesRegisterForGopalJee, Nothing)) = 1 Then
            If obj.ReportType = "Total Sale" Then
                strRunQuery = " select sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final"
            ElseIf obj.ReportType = "Location Wise" Then
                strRunQuery = "select [Warehouse Code] ,[Warehouse Name] ,sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Warehouse Code] ,[Warehouse Name] "
            ElseIf obj.ReportType = "Item Group Wise" Then
                strRunQuery = "select [Warehouse Code] ,[Warehouse Name] ,[Product Group Code],[Product Group Description],sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ")  as Final group by [Warehouse Code] ,[Warehouse Name] ,[Product Group Code] ,[Product Group Description] "
            ElseIf obj.ReportType = "Customer Group Wise" Then
                strRunQuery = "select [Warehouse Code] ,[Warehouse Name] ,[Product Group Code]  ,[Product Group Description] ,[Customer Group Code],[Customer Group Description],sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ")  as Final group by [Warehouse Code] ,[Warehouse Name] ,[Product Group Code] ,[Product Group Description] ,[Customer Group Code],[Customer Group Description]"
            ElseIf obj.ReportType = "Item Wise" Then
                strRunQuery = "select [Warehouse Code] ,[Warehouse Name],[Product Group Code] ,[Product Group Description],[Product Code] ,[Product Name],sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Warehouse Code] ,[Warehouse Name],[Product Group Code]  ,[Product Group Description] ,[Product Code] ,[Product Name] "
            ElseIf obj.ReportType = "Customer Wise" Then
                strRunQuery = "select [Warehouse Code],[Warehouse Name],[product Group Code],[Product Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Product Code],[Product Name],sum(Amount) as Amount,sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ")  as final group by [Warehouse Code],[warehouse Name],[Product Group Code] ,[Product Group Description] ,[Customer Group Code],[Customer Group Description],[Product Code] ,[Product Name],[Customer Code],[Customer Name]"
            ElseIf obj.ReportType = "Document Wise" Then
                strRunQuery = "select [Trans Type],[Document_date],[Document No],max([Narration]) as [Narration], [Warehouse Code],  [Warehouse Name], [Customer Group Code], [Customer Group Description],[Customer Code],[Customer Name],max([Customer GST State Code]) as [Customer GST State Code] ,sum(Amount) as Amount ,sum([Total Tax Amount]) as [Total Tax Amount], sum([Total Amount]) as [Total Amount] from (" & strMain & ") as final group by [Document No],[Warehouse Code] ,[Warehouse Name] ,[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Trans Type] order by convert(date,[Document_Date],103),[Document No]"
            ElseIf obj.ReportType = "Document Detail" Then
                strRunQuery = strMain & "order by convert(date,[Document_Date],103),[Document No]"
            End If
        Else
            If obj.ReportType = "Total Sale" Then
                strRunQuery = "select sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final"
            ElseIf obj.ReportType = "Location Wise" Then
                strRunQuery = "select [Location Code],[Location Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name]"
            ElseIf obj.ReportType = "Item Group Wise" Then
                strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description]"
            ElseIf obj.ReportType = "Customer Group Wise" Then
                strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description]"
            ElseIf obj.ReportType = "Item Wise" Then
                'KUNAL > WHOLLY JOY > TICKET : VERBAL DISCUSSION > DATE : 3 -JAN -2016 
                strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Item Code],[Item Name],Isnull(Sum(Quantity),0)Quantity,isnull(max(UOM),'') [UOM] ,sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Item Code],[Item Name]  "
            ElseIf obj.ReportType = "Customer Wise" Then
                ''richa BHA/23/08/19-000924
                'strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],max( [Customer Zone Code]) as  [Customer Zone Code],Route_no,Route_desc,[Item Code],[Item Name],[Scheme Type], Sampling,sum([Quantity]) as [Total Quantity],max(UOM) as UOM,sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],sum(COGS) as COGS from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name],[Customer Code],[Customer Name],Route_No,Route_desc, [Scheme Type], Sampling"
                strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],max([Customer Category]) as  [Customer Category],max( [Customer Zone Code]) as  [Customer Zone Code],Route_no,Route_desc,max( [RSM Code]) as [RSM Code] ,max( [RSM Name]) as [RSM Name] ,max( [ZSM Code]) as [ZSM Code] ,max( [ZSM Name]) as [ZSM Name] ,max( [ASM Code]) as [ASM Code] ,max( [ASM Name]) as [ASM Name] ,max( [ASO Code])  as [ASO Code] ,max( [ASO Name]) as [ASO Name] , [Item Code],[Item Name],[Scheme Type], Sampling,sum([Sampling Amount] ) as [Sampling Amount],sum([Quantity]) as [Total Quantity],max(UOM) as UOM,sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],sum(COGS) as COGS,case when isnull(max(stockLtr.UOM_Code ),'') ='Ltr' then  (case when coalesce(max(stockLtr.Conversion_Factor),0)=0 then 0 else cast((sum([Quantity])* max(Stock_SU.Conversion_Factor))/(coalesce(max(stockLtr.Conversion_Factor),1)) as numeric(18,3)) end) else 0 end  as [Ltr Qty] from (" & strMain & ") as Final " & Environment.NewLine &
                    " left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on final.[Item Code]=Stock_SU.Item_Code and final.[UOM]=Stock_SU.UOM_Code" & Environment.NewLine &
                    " left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='Ltr') as StockLtr on final.[Item Code]=StockLtr.Item_Code " & Environment.NewLine &
                    " group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name],[Customer Code],[Customer Name],Route_No,Route_desc, [Scheme Type], Sampling"

            ElseIf obj.ReportType = "Customer Amount Wise" Then
                ''richa BHA/23/08/19-000924
                'strRunQuery = "select [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],max( [Customer Zone Code]) as  [Customer Zone Code],Route_no,Route_desc,[Item Code],[Item Name],[Scheme Type], Sampling,sum([Quantity]) as [Total Quantity],max(UOM) as UOM,sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],sum(COGS) as COGS from (" & strMain & ") as Final group by [Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Item Code],[Item Name],[Customer Code],[Customer Name],Route_No,Route_desc, [Scheme Type], Sampling"
                strRunQuery = "select  max([Location Code]) as [Location Code],max([Location Name]) as [Location Name],max([Item Group Code]) as [Item Group Code], max([Item Group Description])[Item Group Description],max([Customer Group Code])[Customer Group Code],max([Customer Group Description]) as [Customer Group Description],[Customer Code],max([Customer Name]) as [Customer Name],
                               max([Customer Category]) as  [Customer Category],max( [Customer Zone Code]) as  [Customer Zone Code],max(Route_no) as Route_no ,max(Route_desc) as Route_desc, isnull(SUM(QTYltr),0) as [Total Quantity],max(UOM) as UOM,sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],
                               sum([Additional Amount]) as [Total Additional Amount], " + strCustomerAmountWiseTaxQuery + "
                               sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount]
                             from (" & strMain & ") as Final " & Environment.NewLine &
                 " left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on final.[Item Code]=Stock_SU.Item_Code and final.[UOM]=Stock_SU.UOM_Code" & Environment.NewLine &
                 " left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='Ltr') as StockLtr on final.[Item Code]=StockLtr.Item_Code " & Environment.NewLine &
                 " group by [Customer Code]"




            ElseIf obj.ReportType = "Document Wise" Then
                If obj.QuickLoad Then
                    strRunQuery = "select [Document No],[Document_date],CONVERT(varchar, [Supply_Date], 103) AS Supply_Date,[Trans Type],[Location Code],[Location Name],max([Location State]) as [Location State],max([GST State Code]) as [GST State Code],max([Dispatch Location GSTIN No]) as [Dispatch Location GSTIN No],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Customer Category],max(Route_no) as Route_no,max(Route_Desc) as Route_Desc,max([City Code]) as [City Code],max([Place of Supply]) as [Place of Supply],max([Customer GST State Code]) as [Customer GST State Code],max([Invoice Type]) as [Invoice Type],max([TIN No]) as [TIN No],max([GR No]) as [GR No],max([GR Date]) as [GR Date],max([WayBill No]) as [WayBill No],max([Transporter Code]) as [Transporter Code],max([Transporter Name]) as [Transporter Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Sale Amount GST]) as [Total Sale Amount GST],sum([Discount Amount]) as [Discount Amount],sum([Rate Differnce]) as [Rate Differnce],sum([Gross Amount])as [Gross Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],max([AR Document No]) as [AR Document No], max([AR Document Amt]) as [AR Document Amt],max([AR Document Discount Amt]) as [AR Document Discount Amt] , case when max(coalesce([AR Amount Before Tax],0))>0 then  max([AR Amount Before Tax]) else  min([AR Amount Before Tax]) end as [AR Amount Before Tax],max([AR Total Tax]) as [AR Total Tax],max([AR Total Add Charge]) as [AR Total Add Charge],max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],sum([Nill Rate Amount]) as [Nill Rate Amount],sum([Exempted Amount]) as [Exempted Amount],sum([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max([Ack No])[Ack No],CONVERT(varchar, max([ACK Date]), 103)[ACK Date],max([IRN No])[IRN No],max(CustomerType)CustomerType,max(PaymentType)PaymentType,max([Payment Terms])[Payment Terms] ,max([Narration]) as [Narration]  from (" & strMain & ") as Final group by [Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Supply_Date],[Trans Type],[Customer Category] order by convert(date,[Document_Date],103),[Document No]"
                Else
                    strRunQuery = "select [Document No],[Document_date],CONVERT(varchar, [Supply_Date], 103) AS Supply_Date,[Trans Type],[Location Code],[Location Name],max([Location State]) as [Location State],max([GST State Code]) as [GST State Code],max([Dispatch Location GSTIN No]) as [Dispatch Location GSTIN No],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Customer Category],max(Route_no) as Route_no,max(Route_Desc) as Route_Desc,max([City Code]) as [City Code],max([Place of Supply]) as [Place of Supply],max([Customer GST State Code]) as [Customer GST State Code],max([Invoice Type]) as [Invoice Type],max([TIN No]) as [TIN No],max([GR No]) as [GR No],max([GR Date]) as [GR Date],max([WayBill No]) as [WayBill No],max([Transporter Code]) as [Transporter Code],max([Transporter Name]) as [Transporter Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Sale Amount GST]) as [Total Sale Amount GST],sum([Discount Amount]) as [Discount Amount],sum([Rate Differnce]) as [Rate Differnce],sum([Gross Amount])as [Gross Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],max([AR Document No]) as [AR Document No], max([AR Document Amt]) as [AR Document Amt],max([AR Document Discount Amt]) as [AR Document Discount Amt] , case when max(coalesce([AR Amount Before Tax],0))>0 then  max([AR Amount Before Tax]) else  min([AR Amount Before Tax]) end as [AR Amount Before Tax],max([AR Total Tax]) as [AR Total Tax],max([AR Total Add Charge]) as [AR Total Add Charge],max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],sum([Nill Rate Amount]) as [Nill Rate Amount],sum([Exempted Amount]) as [Exempted Amount],sum([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision],max([Ack No])[Ack No],CONVERT(varchar, max([ACK Date]), 103)[ACK Date],max([IRN No])[IRN No],max(CustomerType)CustomerType,max(PaymentType)PaymentType,max([Payment Terms])[Payment Terms] ,max([Narration]) as [Narration] from (" & strMain & ") as Final group by [Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Supply_Date],[Trans Type],[Customer Category] order by convert(date,[Document_Date],103),[Document No]"
                End If

            ElseIf obj.ReportType = "Document Detail" Then
                If obj.QuickLoad Then
                    strRunQuery = strMain & " order by convert(date,[Document_Date],103),[Document No]"
                    'strRunQuery = "Select * from ( " & strMain & "  )xy where  xy.Report_UOM_Qty>0 order by convert(date,xy.[Document_Date],103),xy.[Document No]"
                Else
                    strRunQuery = strMain & " order by convert(date,[Document_Date],103),[Document No]"
                    'strRunQuery = "Select * from ( " & strMain & "  )xy where  xy.Report_UOM_Qty>0 order by convert(date,xy.[Document_Date],103),xy.[Document No]"
                End If

            ElseIf obj.ReportType = "Net Sale" Then
                'strRunQuery = "select [Document No],[Document_date],[Trans Type],max([Narration]) as [Narration],[Location Code],[Location Name],max([Location State]) as [Location State],max([GST State Code]) as [GST State Code],max([Dispatch Location GSTIN No]) as [Dispatch Location GSTIN No],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],max([City Code]) as [City Code],max([Place of Supply]) as [Place of Supply],max([Customer GST State Code]) as [Customer GST State Code],max([Invoice Type]) as [Invoice Type],max([TIN No]) as [TIN No],max([GR No]) as [GR No],max([GR Date]) as [GR Date],max([WayBill No]) as [WayBill No],max([Transporter Code]) as [Transporter Code],max([Transporter Name]) as [Transporter Name],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Sale Amount GST]) as [Total Sale Amount GST],sum([Discount Amount]) as [Discount Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount],max([AR Document No]) as [AR Document No], max([AR Document Amt]) as [AR Document Amt],max([AR Document Discount Amt]) as [AR Document Discount Amt] , case when max(coalesce([AR Amount Before Tax],0))>0 then  max([AR Amount Before Tax]) else  min([AR Amount Before Tax]) end as [AR Amount Before Tax],max([AR Total Tax]) as [AR Total Tax],max([AR Total Add Charge]) as [AR Total Add Charge],max([Invoice Type GST]) as [Invoice Type GST],max([GSTIN No Company]) as [GSTIN No Company],max([GSTIN No Customer]) as [GSTIN No Customer],sum([Nill Rate Amount]) as [Nill Rate Amount],sum([Exempted Amount]) as [Exempted Amount],sum([Non GST Supply]) as [Non GST Supply],max([Reverse Charge]) as [Reverse Charge],max([Export Type]) as [Export Type],max(Port) as Port,max([Shipping Bill No]) as [Shipping Bill No],max([Shipping Bill Date]) as [Shipping Bill Date],max([Original Invoice No]) as [Original Invoice No],max([Original Invoice Date]) as [Original Invoice Date],max([Reason for Revision]) as [Reason for Revision]  from (" & strMain & ") as Final group by [Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Trans Type] order by convert(date,[Document_Date],103),[Document No]"
                'strRunQuery = "select [Document_date],[Document No],[Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Item Code],[Item Name],[Trans Type],max([Narration]) as [Narration],max([Location State]) as [Location State],max([Invoice Type]) as [Invoice Type],sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG],sum([Sale Amount]) as [Total Sale Amount],sum([Sale Amount GST]) as [Total Sale Amount GST],sum([Discount Amount]) as [Discount Amount],sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount],sum([Total Amount] ) as [Total Amount] from (" & strMain & ") as Final group by [Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Document_date],[Trans Type] order by convert(date,[Document_Date],103),[Document No]"
                'If obj.QuickLoad Then
                '    strRunQuery = strMain & " order by convert(date,[Document_Date],103),[Document No]"
                'Else
                '    strRunQuery = strMain & " order by convert(date,[Document_Date],103),[Document No]"
                'End If
                'strRunQuery = "select [Trans Type],Document_date as [Document_date],[Document No],[Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name],[Item Code],[Item Name] " & _
                '",sum(COALESCE([FAT KG],0)) as [Total FAT KG],sum(COALESCE([SNF KG],0)) as [Total SNF KG] " & _
                '",sum([Quantity]) as [Quantity],[UOM],max([Item Rate]) as [Item Rate],sum([Sale Amount]) as [Total Sale Amount] " & _
                '",sum([Additional Amount]) as [Total Additional Amount],sum([Total Tax Amount]) as [Total Tax Amount] " & _
                '",sum([Total Amount] ) as [Total Amount] " & _
                '",max([Vehicle No]) as [Vehicle No] " & _
                '",max([Narration]) as [Narration] " & _
                '",max([Customer Zone Description]) as [Customer Zone Description] " & _
                '",max([Executive]) as [Sales Person Name],max([Scheme Type]) as [Scheme Type],max([COGS]) as [COGS] " & " from (" & strMain & " " & _
                '") as Final group by [Document_date],[Document No],[Location Code],[Location Name],[Customer Group Code],[Customer Group Description],[Customer Code],[Customer Name]" & _
                '",[Item Code],[Item Name],[UOM],[Item Group Code],[Item Group Description],[Trans Type],[Scheme Type] order by convert(date,[Document_Date],103),[Document No]"


                strRunQuery = "select [Trans Type],Document_date as [Document_date],[Document No],[Location Code],[Location Name],[Item Group Code],[Item Group Description],[Customer Group Code],[Customer Group Description], " &
                "[Customer Code],[Customer Name],Route_no,Route_Desc,[Item Code],[Item Name] " &
                ",COALESCE([FAT KG],0) as [Total FAT KG],COALESCE([SNF KG],0) as [Total SNF KG] " &
                ",[Quantity],[UOM],[Item Rate],[Sale Amount] as [Total Sale Amount] " &
                ",[Additional Amount] as [Total Additional Amount],[Total Tax Amount] " &
                ",[Total Amount] " &
                ",[Vehicle No] " &
                ",[Narration] " &
                ", [Customer Category] " &
                ", [Customer Zone Description] " &
                ",[Executive] as [Sales Person Name], [Scheme Type],[COGS] " & " from (" & strMain & " " &
                ") as Final order by convert(date,[Document_Date],103),[Document No]"

            ElseIf obj.ReportType = "Document Info Level" Then
                strRunQuery = strMainDocInfoLevel & "order by convert(date,[Document_Date],103),[Document No]"
            End If
        End If


        Return strRunQuery
    End Function

End Class

