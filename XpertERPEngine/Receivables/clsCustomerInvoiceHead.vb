Imports common
Imports System.Data.SqlClient
Public Class clsCustomerInvoiceHead

#Region "Variables"
    Public LeakageAmount As Decimal = 0
    Public TAX1_ExciseFOCAmt As Double = 0
    Public TAX2_ExciseFOCAmt As Double = 0
    Public TAX3_ExciseFOCAmt As Double = 0
    Public TAX4_ExciseFOCAmt As Double = 0
    Public Trans_Type As String = Nothing
    Public Return_Type As String = Nothing
    Public PROJECT_ID As String = Nothing
    Public Document_No As String = Nothing
    Public Document_Date As DateTime
    Public Customer_Code As String = Nothing
    Public Customer_Name As String = Nothing
    '-------- added by usha--
    Public loc_code As String = Nothing
    Public Location_Code_Prefix As String = Nothing
    '-------end---
    Public Posting_Date As DateTime? = Nothing
    Public Account_Set As String = Nothing
    Public Document_Type As String = Nothing
    Public Order_No As String = Nothing
    Public Document_Total As Double = 0
    Public isCardSale As Integer = 0
    Public On_Hold As Boolean = Nothing
    Public Remarks As String = Nothing
    Public Description As String = Nothing
    Public Tax_Group As String = Nothing
    Public RefDocType As String = Nothing
    Public RefDocNo As String = Nothing
    Public TAX1 As String = Nothing
    Public TAX1_Rate As Double = 0
    Public Tax1_BAmount As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Rate As Double = 0
    Public Tax2_BAmount As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Rate As Double = 0
    Public Tax3_BAmount As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Rate As Double = 0
    Public Tax4_BAmount As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Rate As Double = 0
    Public Tax5_BAmount As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Rate As Double = 0
    Public Tax6_BAmount As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Rate As Double = 0
    Public Tax7_BAmount As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Rate As Double = 0
    Public Tax8_BAmount As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Rate As Double = 0
    Public Tax9_BAmount As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Rate As Double = 0
    Public Tax10_BAmount As Double = 0
    Public TAX10_Amt As Double = 0
    Public Total_Tax As Double = 0
    Public Terms_Code As String = Nothing
    Public Terms_Description As String = Nothing
    Public Due_Date As DateTime
    Public Discount_Percentage As Double = 0
    Public Discount_Base As Double = 0
    Public Discount_Amount As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public Comp_Code As String = Nothing
    Public Balance_Amt As Double = 0
    Public Customer_Control_AC As String = Nothing
    Public Discount_GL_AC As String = Nothing
    Public TAX1_GLAC As String = Nothing
    Public TAX2_GLAC As String = Nothing
    Public TAX3_GLAC As String = Nothing
    Public TAX4_GLAC As String = Nothing
    Public TAX5_GLAC As String = Nothing
    Public TAX6_GLAC As String = Nothing
    Public TAX7_GLAC As String = Nothing
    Public TAX8_GLAC As String = Nothing
    Public TAX9_GLAC As String = Nothing
    Public TAX10_GLAC As String = Nothing
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
    Public RoundOffAmount As Double = 0
    Public TotalSubsidyAmt As Double = 0
    Public Is_Add_TPT As Boolean = False
    Public AgainstScrap As String = Nothing
    Public AgainstScrapReturn As String = Nothing
    Public Against_Sale_No As String = Nothing
    Public Against_Sale_Return_No As String = Nothing
    Public Against_MCC_Material_Sale_Return As String = Nothing
    Public Tax_Calculation_Type As EnumTaxCalucationType
    Public SecurityDeposit As Boolean = False
    Public Is_Einvoice As Integer = 0
    Public SecurityDepositType As Char = Nothing
    Public DateAndTime As DateTime?
    Public TapalNo As String = String.Empty
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Arr As List(Of clsCustomerInvoiceDetail) = Nothing

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public Against_VCGL As String = Nothing
    Public Against_Service_Visit_Code As String = Nothing
    Public Against_Asset_Disposal As String = Nothing
    Public Form_ID As String = ""
    Public Route_No As String = ""

    Public irn_no As String = Nothing
    Public Ack_No As String = Nothing
    Public QR_Code As String = Nothing
    Public Ack_Date As DateTime?

    Public Is_Against_Security_Receipt As Boolean
    Public Against_Security_Receipt_No As String
    Public Against_Subsidy_No As String
    Public AgainstServiceInvoice As String = Nothing
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing

#End Region

    Public Shared Function UpdateEInvoiceAfterPosting(ByVal obj As clsCustomerInvoiceHead, ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If obj IsNot Nothing And clsCommon.myLen(strDocumentNo) > 0 Then
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "IRN_No", obj.irn_no)
                clsCommon.AddColumnsForChange(coll, "Ack_No", obj.Ack_No)
                clsCommon.AddColumnsForChange(coll, "Ack_Date", clsCommon.GetPrintDate(obj.Ack_Date, "dd/MMM/yyyy hh:mm tt"))
                clsCommon.AddColumnsForChange(coll, "QR_Code", obj.QR_Code)


                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Customer_Invoice_HEAD", OMInsertOrUpdate.Update, "TSPL_Customer_Invoice_HEAD.Document_No='" + strDocumentNo + "'", trans)
                '    clsCommon.AddColumnsForChange(coll, "Electronic_Ref_No", obj.Electronic_Ref_No)
                '    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SCRAPSALE_HEAD", OMInsertOrUpdate.Update, "TSPL_SCRAPSALE_HEAD.shipment_No='" + ShipmentNo + "'", trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Function SaveData(ByVal obj As clsCustomerInvoiceHead, ByVal isNewEntry As Boolean, ByVal FormId As String, Optional ByVal IsDirectEntry As Boolean = False) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = obj.SaveData(obj, isNewEntry, trans, FormId, "", "", IsDirectEntry)
            If (isSaved) Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function

    Public Function SaveData(ByVal obj As clsCustomerInvoiceHead, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction, ByVal FormId As String, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing, Optional ByVal strARNoForRecreateOnly As String = Nothing, Optional ByVal IsDirectEntry As Boolean = False) As Boolean
        Dim isSaved As Boolean = True
        Try

            If clsCommon.myLen(obj.loc_code) <= 0 Then
            Throw New Exception("Please first select Location")
        End If
        Dim qry As String
        If obj.Is_Against_Security_Receipt Then
            qry = clsRcptEntryHeader.GetAgainstSercurityQry(obj.Customer_Code, obj.Document_No, obj.Against_Security_Receipt_No)
            Dim dtBal As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If obj.Discount_Base > clsCommon.myCdbl(dtBal.Rows(0)("Amount")) Then
                Throw New Exception("Against receipt No " + obj.Against_Security_Receipt_No + " Balance Amount " + clsCommon.myCstr(dtBal.Rows(0)("Amount")) + " And Document Base amount " + clsCommon.myCstr(obj.Discount_Base))
            End If
            If Not clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal Then
                Throw New Exception("Invoice type should be Credit note")
            End If

        End If

            clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleReceivable, clsUserMgtCode.mbtnARInvoiceEntry, obj.loc_code, obj.Document_Date, trans)

            'If Not isNewEntry Then
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", "TSPL_REMITTANCE", "Document_No", trans)
            'End If

            qry = "delete from TSPL_Customer_Invoice_Detail where Document_No='" + obj.Document_No + "'"
        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_REMITTANCE where Document_No='" + obj.Document_No + "'"
        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim strDocNo As String = ""
        If (isNewEntry) Then
            Dim CreateSeperateSeriesforRefDocARinvforCreditdebit As Integer = 0
            CreateSeperateSeriesforRefDocARinvforCreditdebit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateSeperateSeriesforRefDocARinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocARinvforCreditdebit, trans))
                If obj.Arr.Count <= 0 Then
                    Throw New Exception("Please fill at list one Account")
                End If
                Dim isLocationCodeSegment As Boolean
                Dim strLocation As String
                If objCommonVar.ApplyLocationWisePrefix Then
                    isLocationCodeSegment = False
                    strLocation = obj.Location_Code_Prefix
                Else
                    isLocationCodeSegment = True
                    strLocation = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                    If clsCommon.myLen(strLocation) <= 0 Then
                        Throw New Exception("Please enter account wiht location segment")
                    End If
                End If

                If clsCommon.myLen(strARNoForRecreateOnly) > 0 Then
                obj.Document_No = strARNoForRecreateOnly
            Else
                ''richa 4 Aug,2017 -- separate series will be generated in case of Supplementary invoice and credit note
                Dim str_Invoice_No_For_Supplementary As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select ISNULL(Invoice_No_For_Supplementary,'') from TSPL_SD_SALE_INVOICE_HEAD WHERE Trans_Type ='PS' AND Document_Code ='" & clsCommon.myCstr(obj.Against_Sale_No) & "'", trans))
                If clsCommon.myLen(str_Invoice_No_For_Supplementary) > 0 Then
                    If (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SupplementaryARInvoice, "", strLocation, isLocationCodeSegment)
                        ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SupplementaryARCreditNote, "", strLocation, isLocationCodeSegment)
                        End If
                ElseIf clsCommon.CompairString(clsCommon.myCstr(obj.AgainstServiceInvoice), "Y") = CompairStringResult.Equal Then
                    If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARServiceInvoice, "", strLocation, isLocationCodeSegment)
                        ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARServiceCreditNote, "", strLocation, isLocationCodeSegment)
                        ElseIf (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARServiceDebitNote, "", strLocation, isLocationCodeSegment)
                        End If
                Else
                    If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) Then
                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARInvoice, "", strLocation, isLocationCodeSegment)
                        ElseIf (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                        If clsCommon.myLen(RefDocNo) > 0 And IsDirectEntry = True And CreateSeperateSeriesforRefDocARinvforCreditdebit = 1 Then
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARDebitNote, clsDocTransactionType.DebitRefDoc, strLocation, isLocationCodeSegment)
                            Else
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARDebitNote, clsDocTransactionType.NA, strLocation, isLocationCodeSegment)
                            End If
                      
                    ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                        If clsCommon.myLen(RefDocNo) > 0 And IsDirectEntry = True And CreateSeperateSeriesforRefDocARinvforCreditdebit = 1 Then
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARCreditNote, clsDocTransactionType.CreditRefDoc, strLocation, isLocationCodeSegment)
                            Else
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.ARCreditNote, clsDocTransactionType.NA, strLocation, isLocationCodeSegment)
                            End If
                    End If
                End If
                ''-------
            End If


        End If
        If (clsCommon.myLen(obj.Document_No) <= 0) Then
            Throw New Exception("Error in Document Code Generation")
        End If

        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Trans_Type", obj.Trans_Type)
        clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "Customer_Code", obj.Customer_Code)
        clsCommon.AddColumnsForChange(coll, "Customer_Name", obj.Customer_Name)
            '--------added by usha----
            clsCommon.AddColumnsForChange(coll, "Loc_code", obj.loc_code)
            clsCommon.AddColumnsForChange(coll, "Location_Code_Prefix", obj.Location_Code_Prefix, True)
            '--------------
            clsCommon.AddColumnsForChange(coll, "Account_Set", obj.Account_Set)
        clsCommon.AddColumnsForChange(coll, "Document_Type", obj.Document_Type)
        clsCommon.AddColumnsForChange(coll, "RefDocType", obj.RefDocType)
        clsCommon.AddColumnsForChange(coll, "RefDocNo", obj.RefDocNo)
        clsCommon.AddColumnsForChange(coll, "Order_No", obj.Order_No)
        clsCommon.AddColumnsForChange(coll, "Document_Total", obj.Document_Total)
        clsCommon.AddColumnsForChange(coll, "isCardSale", obj.isCardSale)

        clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
        clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
        clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
        clsCommon.AddColumnsForChange(coll, "PROJECT_ID", obj.PROJECT_ID, True)
        clsCommon.AddColumnsForChange(coll, "Return_Type", obj.Return_Type)
        clsCommon.AddColumnsForChange(coll, "SecurityDeposit", IIf(obj.SecurityDeposit = True, "Y", "N"))
        clsCommon.AddColumnsForChange(coll, "SecurityDepositType", obj.SecurityDepositType)
        clsCommon.AddColumnsForChange(coll, "AgainstServiceInvoice", obj.AgainstServiceInvoice)
        clsCommon.AddColumnsForChange(coll, "TAX1_ExciseFOCAmt", obj.TAX1_ExciseFOCAmt)
        clsCommon.AddColumnsForChange(coll, "TAX2_ExciseFOCAmt", obj.TAX2_ExciseFOCAmt)
        clsCommon.AddColumnsForChange(coll, "TAX3_ExciseFOCAmt", obj.TAX3_ExciseFOCAmt)
        clsCommon.AddColumnsForChange(coll, "TAX4_ExciseFOCAmt", obj.TAX4_ExciseFOCAmt)
        clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
        clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax1_BAmount", obj.Tax1_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
        clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax2_BAmount", obj.Tax2_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
        clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax3_BAmount", obj.Tax3_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
        clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax4_BAmount", obj.Tax4_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
        clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax5_BAmount", obj.Tax5_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
        clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax6_BAmount", obj.Tax6_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
        clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax7_BAmount", obj.Tax7_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
        clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax8_BAmount", obj.Tax8_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
        clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax9_BAmount", obj.Tax9_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
        clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax10_BAmount", obj.Tax10_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
        clsCommon.AddColumnsForChange(coll, "Total_Tax", obj.Total_Tax)

        clsCommon.AddColumnsForChange(coll, "Customer_Control_AC", obj.Customer_Control_AC, True)
        clsCommon.AddColumnsForChange(coll, "Discount_GL_AC", obj.Discount_GL_AC, True)
        clsCommon.AddColumnsForChange(coll, "TAX1_GLAC", obj.TAX1_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX2_GLAC", obj.TAX2_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX3_GLAC", obj.TAX3_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX4_GLAC", obj.TAX4_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX5_GLAC", obj.TAX5_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX6_GLAC", obj.TAX6_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX7_GLAC", obj.TAX7_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX8_GLAC", obj.TAX8_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX9_GLAC", obj.TAX9_GLAC, True)
        clsCommon.AddColumnsForChange(coll, "TAX10_GLAC", obj.TAX10_GLAC, True)
            clsCommon.AddColumnsForChange(coll, "LeakageAmount", obj.LeakageAmount)
            clsCommon.AddColumnsForChange(coll, "Is_Einvoice", obj.Is_Einvoice)

            ''richa agarwal changes on 12 Aug,2016 against AR Aging report for invoice type
            If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) AndAlso clsCommon.myLen(obj.Terms_Code) <= 0 Then
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select TSPL_CUSTOMER_MASTER.Terms_Code ,TSPL_TERMS_MASTER.Terms_Desc,TSPL_TERMS_MASTER.No_Days   from TSPL_CUSTOMER_MASTER left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code =TSPL_CUSTOMER_MASTER.Terms_Code where TSPL_CUSTOMER_MASTER.Cust_Code ='" & clsCommon.myCstr(obj.Customer_Code) & "'", trans)
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                obj.Terms_Code = clsCommon.myCstr(dt1.Rows(0)("Terms_Code"))
                obj.Terms_Description = clsCommon.myCstr(dt1.Rows(0)("Terms_Desc"))
                obj.Due_Date = clsCommon.myCDate(obj.Document_Date).AddDays(clsCommon.myCdbl(dt1.Rows(0)("No_Days")))
            Else
                Throw New Exception("Please enter Terms Code for Customer " & obj.Customer_Code & " in Customer Master")
            End If
        End If

        ''richa agarwal changes on 06 JUNE,2019 against TEC/04/06/19-000512 DEBIT NOTE AND CREDIT NOTE IF TERMS CODE IS NULL
        If (clsCommon.CompairString(obj.Document_Type, "I") <> CompairStringResult.Equal) AndAlso clsCommon.myLen(obj.Terms_Code) <= 0 Then
            obj.Due_Date = obj.Document_Date
        End If

        clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
        clsCommon.AddColumnsForChange(coll, "Terms_Description", obj.Terms_Description)
        clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MMM/yyyy"))
        clsCommon.AddColumnsForChange(coll, "Discount_Percentage", obj.Discount_Percentage)
        clsCommon.AddColumnsForChange(coll, "Discount_Base", clsCommon.myCdbl(obj.Discount_Base))
        clsCommon.AddColumnsForChange(coll, "Discount_Amount", obj.Discount_Amount)
        clsCommon.AddColumnsForChange(coll, "Amount_Less_Discount", obj.Amount_Less_Discount)
        clsCommon.AddColumnsForChange(coll, "Balance_Amt", obj.Document_Total)
        clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
        clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
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
            clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
            clsCommon.AddColumnsForChange(coll, "TotalSubsidyAmt", obj.TotalSubsidyAmt)
            clsCommon.AddColumnsForChange(coll, "Is_Add_TPT", IIf(obj.Is_Add_TPT, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
        clsCommon.AddColumnsForChange(coll, "AgainstScrap", obj.AgainstScrap)
            clsCommon.AddColumnsForChange(coll, "Against_Sale_No", obj.Against_Sale_No, True)
            clsCommon.AddColumnsForChange(coll, "Against_Sale_Return_No", obj.Against_Sale_Return_No)
        clsCommon.AddColumnsForChange(coll, "AgainstScrapReturn", obj.AgainstScrapReturn)
        '' Anubhooti 18-Mar-2015 (Save Against_VCGL)
        clsCommon.AddColumnsForChange(coll, "Against_VCGL", obj.Against_VCGL, True)
        ''
        '' Anubhooti 30-Oct-2015 (Save Against_Service_Visit_Code)
        clsCommon.AddColumnsForChange(coll, "Against_Service_Visit_Code", obj.Against_Service_Visit_Code, True)
        ''
        clsCommon.AddColumnsForChange(coll, "Against_MCC_Material_Sale_Return", obj.Against_MCC_Material_Sale_Return, True)
        clsCommon.AddColumnsForChange(coll, "Against_Asset_Disposal", obj.Against_Asset_Disposal, True)

        clsCommon.AddColumnsForChange(coll, "Is_Against_Security_Receipt", IIf(obj.Is_Against_Security_Receipt, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Against_Security_Receipt_No", obj.Against_Security_Receipt_No, True)
            clsCommon.AddColumnsForChange(coll, "Against_Subsidy_No", obj.Against_Subsidy_No, True)
            clsCommon.AddColumnsForChange(coll, "Route_No", obj.Route_No, True)

            '' currencyconversion BM00000007917
            If clsCommon.myLen(obj.CURRENCY_CODE) = 0 Then
            'obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select BaseCurrencyCode from TSPL_COMPANY_MASTER where Comp_Code='" & objCommonVar.CurrentCompanyCode & "'", trans))
            'clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            'clsCommon.AddColumnsForChange(coll, "ConvRate", 1)

            obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(CURRENCY_CODE,'') as CURRENCY_CODE from TSPL_CUSTOMER_MASTER where Cust_Code ='" & obj.Customer_Code & "'", trans))
            If clsCommon.myLen(obj.CURRENCY_CODE) > 0 Then
                Dim dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(obj.Document_Date, obj.CURRENCY_CODE, trans)
                If dt.Rows.Count = 0 Then
                    If obj.CURRENCY_CODE = objCommonVar.BaseCurrencyCode Then
                        obj.ConvRate = 1
                    Else
                        Throw New Exception("Conversion rate not entered for currency '" & obj.CURRENCY_CODE & "'")
                    End If
                Else
                    obj.ConvRate = dt.Rows(0).Item("Rate")
                    obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd-MMM-yyyy")
                End If
            Else
                Throw New Exception("Please enter currency for Customer'" & obj.Customer_Code & "' in Customer master.")
            End If

            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
        Else
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
        End If

        If clsCommon.myLen(obj.ApplicableFrom) > 0 Then
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", clsCommon.GetPrintDate(obj.ApplicableFrom, "dd/MMM/yyyy"), True)
        Else
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", Nothing, True)
        End If

        clsCommon.AddColumnsForChange(coll, "TapalNo", obj.TapalNo, True)
        If clsCommon.myLen(obj.DateAndTime) > 0 Then
            clsCommon.AddColumnsForChange(coll, "DateAndTime", clsCommon.GetPrintDate(obj.DateAndTime, "dd/MMM/yyyy hh:mm tt"))
        Else
            clsCommon.AddColumnsForChange(coll, "DateAndTime", Nothing, True)
        End If

        '' End currencyconversion

        If isNewEntry Then
            clsCommon.AddColumnsForChange(coll, "Document_No", obj.Document_No)
            clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt"))
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Customer_Invoice_Head", OMInsertOrUpdate.Insert, "", trans)
        Else
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Customer_Invoice_Head", OMInsertOrUpdate.Update, "Document_No='" + obj.Document_No + "'", trans)
        End If
            isSaved = isSaved AndAlso clsCustomerInvoiceDetail.SaveData(obj.Document_No, Arr, trans)
            isSaved = isSaved AndAlso clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", "TSPL_REMITTANCE", "Document_No", trans)

            '' Anubhooti 05-Dec-2014 (Commented CreateGLEntty and check new GLEntryFunction)
            ' isSaved = isSaved AndAlso CreateGLEntty(obj, trans, True, FormId)

            ''richa 21 Dec,2018  TEC/02/11/18-000360 create journal entry for opening in case of Credit or debit note as Journal Master table instead of journal master op table
            'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, trans)
            Dim JEWithOPening As Boolean = False
            If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                If clsCommon.myCDate(clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                    JEWithOPening = True
                End If
            End If

            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
            isSaved = isSaved AndAlso CreateJournalEntryForOpening(obj, trans, True, FormId, strVoucherNoForRecreateOnly)
        Else
            isSaved = isSaved AndAlso CreateGLEntryForAllCases(obj, trans, True, FormId, strVoucherNoForRecreateOnly)
        End If
        ''-----------------



        isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_No, obj.arrCustomFields, trans)

        isSaved = isSaved AndAlso clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "Document_No", obj.Document_No, "TSPL_Customer_Invoice_Head", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String) As clsCustomerInvoiceHead
        Return GetData(strDocumentNo, Nothing)
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As clsCustomerInvoiceHead
        Dim obj As clsCustomerInvoiceHead = Nothing
        Dim qry As String = "Select * from TSPL_Customer_Invoice_Head where Document_No='" + strDocumentNo + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsCustomerInvoiceHead()
            obj.LeakageAmount = clsCommon.myCdbl(dt.Rows(0)("LeakageAmount"))
            obj.Document_No = clsCommon.myCstr(dt.Rows(0)("Document_No"))
            obj.Document_Date = clsCommon.myCstr(dt.Rows(0)("Document_Date"))
            obj.Customer_Code = clsCommon.myCstr(dt.Rows(0)("Customer_Code"))
            obj.Customer_Name = clsCommon.myCstr(dt.Rows(0)("Customer_Name"))
            '----------added by usha
            obj.loc_code = clsCommon.myCstr(dt.Rows(0)("Loc_code"))
            obj.Location_Code_Prefix = clsCommon.myCstr(dt.Rows(0)("Location_Code_Prefix"))
            '-----------
            If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
            Else
                obj.Posting_Date = Nothing
            End If
            obj.AgainstServiceInvoice = clsCommon.myCstr(dt.Rows(0)("AgainstServiceInvoice"))
            obj.Account_Set = clsCommon.myCstr(dt.Rows(0)("Account_Set"))
            obj.Document_Type = clsCommon.myCstr(dt.Rows(0)("Document_Type"))
            obj.RefDocType = clsCommon.myCstr(dt.Rows(0)("RefDocType"))
            obj.RefDocNo = clsCommon.myCstr(dt.Rows(0)("RefDocNo"))
            obj.Order_No = clsCommon.myCstr(dt.Rows(0)("Order_No"))
            obj.Route_No = clsCommon.myCstr(dt.Rows(0)("Route_No"))
            obj.Document_Total = clsCommon.myCdbl(dt.Rows(0)("Document_Total"))
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.SecurityDeposit = IIf(clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("SecurityDeposit")), "Y") = CompairStringResult.Equal, True, False)
            obj.SecurityDepositType = clsCommon.myCstr(dt.Rows(0)("SecurityDepositType"))
            obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
            obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
            obj.TAX1_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
            obj.Tax1_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax1_BAmount"))
            obj.TAX1_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            obj.TAX2 = clsCommon.myCstr(dt.Rows(0)("TAX2"))
            obj.TAX2_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
            obj.Tax2_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax2_BAmount"))
            obj.TAX2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            obj.TAX3 = clsCommon.myCstr(dt.Rows(0)("TAX3"))
            obj.Tax3_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax3_BAmount"))
            obj.TAX3_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
            obj.TAX3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            obj.TAX4 = clsCommon.myCstr(dt.Rows(0)("TAX4"))
            obj.TAX4_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
            obj.Tax4_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax4_BAmount"))
            obj.TAX4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            obj.TAX5 = clsCommon.myCstr(dt.Rows(0)("TAX5"))
            obj.TAX5_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
            obj.Tax5_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax5_BAmount"))
            obj.TAX5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
            obj.TAX6 = clsCommon.myCstr(dt.Rows(0)("TAX6"))
            obj.TAX6_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
            obj.Tax6_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax6_BAmount"))
            obj.TAX6_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
            obj.TAX7 = clsCommon.myCstr(dt.Rows(0)("TAX7"))
            obj.TAX7_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX7_Rate"))
            obj.Tax7_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax7_BAmount"))
            obj.TAX7_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
            obj.TAX8 = clsCommon.myCstr(dt.Rows(0)("TAX8"))
            obj.TAX8_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX8_Rate"))
            obj.Tax8_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax8_BAmount"))
            obj.TAX8_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
            obj.TAX9 = clsCommon.myCstr(dt.Rows(0)("TAX9"))
            obj.TAX9_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX9_Rate"))
            obj.Tax9_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax9_BAmount"))
            obj.TAX9_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))
            obj.TAX10 = clsCommon.myCstr(dt.Rows(0)("TAX10"))
            obj.TAX10_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX10_Rate"))
            obj.Tax10_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax10_BAmount"))
            obj.TAX10_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
            obj.Total_Tax = clsCommon.myCdbl(dt.Rows(0)("Total_Tax"))
            obj.TAX1_ExciseFOCAmt = clsCommon.myCdbl(dt.Rows(0)("TAX1_ExciseFOCAmt"))
            obj.TAX2_ExciseFOCAmt = clsCommon.myCdbl(dt.Rows(0)("TAX2_ExciseFOCAmt"))
            obj.TAX3_ExciseFOCAmt = clsCommon.myCdbl(dt.Rows(0)("TAX3_ExciseFOCAmt"))
            obj.TAX4_ExciseFOCAmt = clsCommon.myCdbl(dt.Rows(0)("TAX4_ExciseFOCAmt"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Terms_Description = clsCommon.myCstr(dt.Rows(0)("Terms_Description"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            obj.Discount_Percentage = clsCommon.myCdbl(dt.Rows(0)("Discount_Percentage"))
            obj.Discount_Base = clsCommon.myCdbl(dt.Rows(0)("Discount_Base"))
            obj.Discount_Amount = clsCommon.myCdbl(dt.Rows(0)("Discount_Amount"))
            obj.Amount_Less_Discount = clsCommon.myCdbl(dt.Rows(0)("Amount_Less_Discount"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Balance_Amt = clsCommon.myCdbl(dt.Rows(0)("Balance_Amt"))
            obj.Customer_Control_AC = clsCommon.myCstr(dt.Rows(0)("Customer_Control_AC"))
            obj.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Discount_GL_AC"))
            obj.TAX1_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC"))
            obj.TAX2_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC"))
            obj.TAX3_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC"))
            obj.TAX4_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC"))
            obj.TAX5_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC"))
            obj.TAX6_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC"))
            obj.TAX7_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC"))
            obj.TAX8_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC"))
            obj.TAX9_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC"))
            obj.TAX10_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC"))
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
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.Total_Add_Charge = clsCommon.myCdbl(dt.Rows(0)("Total_Add_Charge"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            obj.TotalSubsidyAmt = clsCommon.myCdbl(dt.Rows(0)("TotalSubsidyAmt"))
            obj.Is_Add_TPT = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Add_TPT")) = 1, True, False)
            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)
            obj.AgainstScrap = clsCommon.myCstr(dt.Rows(0)("AgainstScrap"))
            obj.Against_Sale_No = clsCommon.myCstr(dt.Rows(0)("Against_Sale_No"))
            obj.Against_Sale_Return_No = clsCommon.myCstr(dt.Rows(0)("Against_Sale_Return_No"))
            '' Richa 28-Mar-2019 (AgainstScrapReturn) ERO/19/03/19-000515
            obj.AgainstScrapReturn = clsCommon.myCstr(dt.Rows(0)("AgainstScrapReturn"))
            '' Anubhooti 18-Mar-2015 (Fetch Against_VCGL)
            obj.Against_VCGL = clsCommon.myCstr(dt.Rows(0)("Against_VCGL"))
            obj.Against_Asset_Disposal = clsCommon.myCstr(dt.Rows(0)("Against_Asset_Disposal"))
            ''
            '' Anubhooti 30-Oct-2015 (Fetch Against_Service_Visit_Code)
            obj.Against_Service_Visit_Code = clsCommon.myCstr(dt.Rows(0)("Against_Service_Visit_Code"))
            obj.Trans_Type = clsCommon.myCstr(dt.Rows(0)("Trans_Type"))
            ''
            obj.Against_MCC_Material_Sale_Return = clsCommon.myCstr(dt.Rows(0)("Against_MCC_Material_Sale_Return"))
            obj.PROJECT_ID = clsCommon.myCstr(dt.Rows(0)("PROJECT_ID"))
            obj.Return_Type = clsCommon.myCstr(dt.Rows(0)("Return_Type"))

            '' CURRENCYCONVERSION 
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            '' END CURRENCYCONVERSION 
            If IsDBNull(dt.Rows(0)("DateAndTime")) = True Then
                obj.DateAndTime = Nothing
            Else
                obj.DateAndTime = clsCommon.myCstr(dt.Rows(0)("DateAndTime"))
            End If
            obj.TapalNo = clsCommon.myCstr(dt.Rows(0)("TapalNo"))
            obj.Is_Einvoice = clsCommon.myCdbl(dt.Rows(0)("Is_Einvoice"))
            obj.Is_Against_Security_Receipt = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_Against_Security_Receipt")) = 1, True, False)
            obj.Against_Security_Receipt_No = clsCommon.myCstr(dt.Rows(0)("Against_Security_Receipt_No"))

            qry = "Select * from TSPL_Customer_Invoice_Detail where Document_No='" + strDocumentNo + "' ORDER BY SNo"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsCustomerInvoiceDetail)
                Dim objTr As clsCustomerInvoiceDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsCustomerInvoiceDetail
                    objTr.Document_No = clsCommon.myCstr(dr("Document_No"))
                    objTr.SNo = clsCommon.myCstr(dr("SNo"))
                    objTr.AddChargeCode = clsCommon.myCstr(dr("AddChargeCode"))
                    objTr.AddChargeDesc = clsCommon.myCstr(dr("AddChargeDesc"))
                    objTr.GL_Account_Code = clsCommon.myCstr(dr("GL_Account_Code"))
                    objTr.GL_Account_Desc = clsCommon.myCstr(dr("GL_Account_Desc"))
                    objTr.Reco_Control_Account = clsCommon.myCstr(dr("Reco_Control_Account"))
                    objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                    objTr.Discount_Per = clsCommon.myCdbl(dr("Discount_Per"))
                    objTr.Discount = clsCommon.myCdbl(dr("Discount"))
                    objTr.Amount_less_Discount = clsCommon.myCdbl(dr("Amount_less_Discount"))
                    objTr.TAX1 = clsCommon.myCstr(dr("TAX1"))
                    objTr.TAX1_Rate = clsCommon.myCdbl(dr("TAX1_Rate"))
                    objTr.TAX1_Amt = clsCommon.myCdbl(dr("TAX1_Amt"))
                    objTr.TAX2 = clsCommon.myCstr(dr("TAX2"))
                    objTr.TAX2_Rate = clsCommon.myCdbl(dr("TAX2_Rate"))
                    objTr.TAX2_Amt = clsCommon.myCdbl(dr("TAX2_Amt"))
                    objTr.TAX3 = clsCommon.myCstr(dr("TAX3"))
                    objTr.TAX3_Rate = clsCommon.myCdbl(dr("TAX3_Rate"))
                    objTr.TAX3_Amt = clsCommon.myCdbl(dr("TAX3_Amt"))
                    objTr.TAX4 = clsCommon.myCstr(dr("TAX4"))
                    objTr.TAX4_Rate = clsCommon.myCdbl(dr("TAX4_Rate"))
                    objTr.TAX4_Amt = clsCommon.myCdbl(dr("TAX4_Amt"))
                    objTr.TAX5 = clsCommon.myCstr(dr("TAX5"))
                    objTr.TAX5_Rate = clsCommon.myCdbl(dr("TAX5_Rate"))
                    objTr.TAX5_Amt = clsCommon.myCdbl(dr("TAX5_Amt"))
                    objTr.TAX6 = clsCommon.myCstr(dr("TAX6"))
                    objTr.TAX6_Rate = clsCommon.myCdbl(dr("TAX6_Rate"))
                    objTr.TAX6_Amt = clsCommon.myCdbl(dr("TAX6_Amt"))
                    objTr.TAX7 = clsCommon.myCstr(dr("TAX7"))
                    objTr.TAX7_Rate = clsCommon.myCdbl(dr("TAX7_Rate"))
                    objTr.TAX7_Amt = clsCommon.myCdbl(dr("TAX7_Amt"))
                    objTr.TAX8 = clsCommon.myCstr(dr("TAX8"))
                    objTr.TAX8_Rate = clsCommon.myCdbl(dr("TAX8_Rate"))
                    objTr.TAX8_Amt = clsCommon.myCdbl(dr("TAX8_Amt"))
                    objTr.TAX9 = clsCommon.myCstr(dr("TAX9"))
                    objTr.TAX9_Rate = clsCommon.myCdbl(dr("TAX9_Rate"))
                    objTr.TAX9_Amt = clsCommon.myCdbl(dr("TAX9_Amt"))
                    objTr.TAX10 = clsCommon.myCstr(dr("TAX10"))
                    objTr.TAX10_Rate = clsCommon.myCdbl(dr("TAX10_Rate"))
                    objTr.TAX10_Amt = clsCommon.myCdbl(dr("TAX10_Amt"))
                    objTr.Total_Tax = clsCommon.myCdbl(dr("Total_Tax"))
                    objTr.Total_Amount = clsCommon.myCdbl(dr("Total_Amount"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))
                    objTr.Comments = clsCommon.myCstr(dr("Comments"))

                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(dr("TAX1_Base_Amt"))
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(dr("TAX2_Base_Amt"))
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(dr("TAX3_Base_Amt"))
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(dr("TAX4_Base_Amt"))
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(dr("TAX5_Base_Amt"))
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(dr("TAX6_Base_Amt"))
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(dr("TAX7_Base_Amt"))
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(dr("TAX8_Base_Amt"))
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(dr("TAX9_Base_Amt"))
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(dr("TAX10_Base_Amt"))

                    ''Richa
                    objTr.Hirerachy_Code = clsCommon.myCstr(dr("Hirerachy_Code"))
                    objTr.Cost_Centre_Code = clsCommon.myCstr(dr("Cost_Centre_Code"))
                    objTr.Hirerachy_Code1 = clsCommon.myCstr(dr("Hirerachy_Code1"))
                    objTr.Hirerachy_Code2 = clsCommon.myCstr(dr("Hirerachy_Code2"))
                    objTr.Hirerachy_Code3 = clsCommon.myCstr(dr("Hirerachy_Code3"))
                    objTr.Hirerachy_Code4 = clsCommon.myCstr(dr("Hirerachy_Code4"))
                    objTr.Distributor_Commission_Amt = clsCommon.myCdbl(dr("Distributor_Commission_Amt"))
                    objTr.Promotional_GL_Account_Code = clsCommon.myCstr(dr("Promotional_GL_Account_Code"))
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(dr("Transporter_Commission_Amt"))
                    objTr.Transporter_GL_Account_Code = clsCommon.myCstr(dr("Transporter_GL_Account_Code"))
                    objTr.Security_Amt = clsCommon.myCdbl(dr("Security_Amt"))
                    objTr.SD_GL_Account_Code = clsCommon.myCstr(dr("SD_GL_Account_Code"))
                    obj.Arr.Add(objTr)
                Next
            End If
        End If
        Return obj
    End Function
    Public Shared Function EInvoice_Implementation(ByVal strDocNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If

            Dim strtoken As String = ClsEInvoiceOFAPIs.IsGenerateAuthTokenNo_Required(objCommonVar.CurrentCompanyCode, strLocation, trans)
            If clsCommon.myLen(strtoken) > 0 Then
                '                Dim strQry As String = "select TSPL_Customer_master .Cust_Code ,TSPL_Customer_Invoice_Head.document_No as DocNo,convert(date,TSPL_Customer_Invoice_Head.Document_Date,103) as DocDate,case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'DBN' when TSPL_Customer_Invoice_Head.Document_Type ='C' then 'CRN' else 'INV' end as DocType ,'B2B' as SupTyp, 'N'  as IgstOnIntra,TSPL_LOCATION_MASTER.GSTNO as SellerGSTINNo ,TSPL_LOCATION_MASTER.location_desc as SellerLglNm,TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm,TSPL_LOCATION_MASTER.Add1 as SellerAdd1,TSPL_LOCATION_MASTER.Add2 as SellerAdd2 ,TSPL_LOCATION_MASTER.City_Code  as SellerLoc,TSPL_LOCATION_MASTER.Pin_Code  as SellerPincode,Location_State_Master.GST_STATE_Code as SellerStcd,TSPL_LOCATION_MASTER.Phone1 as SellerPhone,TSPL_LOCATION_MASTER.Email as SellerEmail,TSPL_Customer_master.GSTNo as BuyerGSTINNo ,TSPL_Customer_master.Customer_Name as BuyerLglNm,TSPL_Customer_master.Alies_name as BuyerTrdNm,Customer_State_Master.GST_STATE_Code  as BuyerPOS,TSPL_Customer_master.Add1 as BuyerAdd1,TSPL_Customer_master.Add2 as BuyerAdd2 ,tspl_city_master.City_Name as BuyerLoc,cast(TSPL_Customer_master.PIN_NO as int) as BuyerPincode,Customer_State_Master.GST_STATE_Code as BuyerStcd,TSPL_Customer_master.Phone1 as BuyerPhone,TSPL_Customer_master.Email as BuyerEmail,TSPL_Customer_Invoice_Detail.SNo as ItemSlNo, 'Y' as ItemIsServc,TSPL_ADDITIONAL_charges.Description AS ItemPrdDesc,TSPL_ADDITIONAL_charges.Sac_code AS ItemHsnCd,
                '0 as ItemQty,'OTH' as ItemUnit,0 as ItemUnitPrice,TSPL_Customer_Invoice_Detail.Amount as ItemTotAmt,TSPL_Customer_Invoice_Detail.Discount as ItemDiscount,TSPL_Customer_Invoice_Detail.Amount_Less_Discount as ItemAssAmt,case when ISNULL(TSPL_Customer_Invoice_Detail .tax1,'') ='IGST' THEN TSPL_Customer_Invoice_Detail.TAX1_Rate when ISNULL(TSPL_Customer_Invoice_Detail .tax1,'') ='CGST' AND   ISNULL(TSPL_Customer_Invoice_Detail .tax2,'') ='SGST'  THEN TSPL_Customer_Invoice_Detail.TAX1_Rate+TSPL_Customer_Invoice_Detail.TAX2_Rate  ELSE 0 end as ItemGstRt, case when TSPL_Customer_Invoice_Detail .TAX1 ='SGST' AND TSPL_Customer_Invoice_Detail .TAX2  ='CGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt when TSPL_Customer_Invoice_Detail .TAX1 ='CGST' AND TSPL_Customer_Invoice_Detail .TAX2  ='SGST' then TSPL_Customer_Invoice_Detail.TAX2_Amt else 0 end ItemSgstAmt,case when TSPL_Customer_Invoice_Detail .TAX1 ='IGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt else 0 end ItemIgstAmt,case when TSPL_Customer_Invoice_Detail .TAX1 ='CGST' AND TSPL_Customer_Invoice_Detail .TAX2  ='SGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt when TSPL_Customer_Invoice_Detail .TAX1 ='SGST' AND TSPL_Customer_Invoice_Detail .TAX2  ='CGST' then TSPL_Customer_Invoice_Detail.TAX2_Amt else 0 end ItemCgstAmt,0 as ItemOthChrg,TSPL_Customer_Invoice_Detail.total_amount-case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_Customer_Invoice_Detail.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y' THEN  TSPL_Customer_Invoice_Detail.TAX3_AMT ELSE 0 END as ItemTotItemVal,TSPL_Customer_Invoice_Head .Discount_Base  as ValDtlsAssVal,case when TSPL_Customer_Invoice_Head .TAX1 ='CGST' AND TSPL_Customer_Invoice_Head .TAX2  ='SGST' then TSPL_Customer_Invoice_Head.TAX1_Amt when TSPL_Customer_Invoice_Head .TAX1 ='SGST' AND TSPL_Customer_Invoice_Head .TAX2  ='CGST' then TSPL_Customer_Invoice_Head.TAX2_Amt else 0 end ValDtlsCgstVal, case when TSPL_Customer_Invoice_Head .TAX1 ='SGST' AND TSPL_Customer_Invoice_Head .TAX2  ='CGST' then TSPL_Customer_Invoice_Head.TAX1_Amt when TSPL_Customer_Invoice_Head .TAX1 ='CGST' AND TSPL_Customer_Invoice_Head .TAX2  ='SGST' then TSPL_Customer_Invoice_Head.TAX2_Amt else 0 end ValDtlsSgstVal,case when TSPL_Customer_Invoice_Head .TAX1 ='IGST' then TSPL_Customer_Invoice_Head.TAX1_Amt else 0 end ValDtlsIgstVal,TSPL_Customer_Invoice_Head.Discount_Amount as ValDtlsDiscount,case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_Customer_Invoice_Head.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y' THEN  TSPL_Customer_Invoice_Head.TAX3_AMT ELSE 0 END as ValDtlsOthChrg,TSPL_Customer_Invoice_Head.document_total as ValDtlsTotInvVal,TSPL_Customer_Invoice_Head.RoundOffAmount  as ValDtlsRndOffAmt
                'from TSPL_Customer_Invoice_Head
                'Left Outer Join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code  ='" & objCommonVar.CurrentCompanyCode & "'
                'Left Outer Join TSPL_Customer_master on TSPL_Customer_master.Cust_Code  =TSPL_Customer_Invoice_Head.Customer_Code
                'left Outer Join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code ='" + strLocation + "'
                'left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.document_No=TSPL_Customer_Invoice_Head.document_No
                'left outer join TSPL_ADDITIONAL_charges on TSPL_ADDITIONAL_charges.Code=TSPL_Customer_Invoice_Detail.AddChargeCode
                'left outer join TSPL_STATE_MASTER as LOCATION_State_Master on LOCATION_State_Master.STATE_CODE  =TSPL_LOCATION_MASTER.State
                'left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE  =TSPL_Customer_master.State 
                'left outer join tspl_city_master on tspl_city_master.city_code=TSPL_Customer_master.City_Code
                'left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code =TSPL_Customer_Invoice_Head.Tax2
                'left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code =TSPL_Customer_Invoice_Head.Tax3
                'where TSPL_Customer_Invoice_Head.document_No ='" & strDocNo & "'"
                Dim strQry As String = "select TSPL_Customer_master .Cust_Code ,TSPL_Customer_Invoice_Head.document_No as DocNo,convert(date,TSPL_Customer_Invoice_Head.Document_Date,103) as DocDate,case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'DBN' when TSPL_Customer_Invoice_Head.Document_Type ='C' then 'CRN' else 'INV' end as DocType ,'B2B' as SupTyp, 'N'  as IgstOnIntra,TSPL_LOCATION_MASTER.GSTNO as SellerGSTINNo ,TSPL_LOCATION_MASTER.location_desc as SellerLglNm,TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm,TSPL_LOCATION_MASTER.Add1 as SellerAdd1,TSPL_LOCATION_MASTER.Add2 as SellerAdd2 ,TSPL_LOCATION_MASTER.City_Code  as SellerLoc,TSPL_LOCATION_MASTER.Pin_Code  as SellerPincode,Location_State_Master.GST_STATE_Code as SellerStcd,TSPL_LOCATION_MASTER.Phone1 as SellerPhone,TSPL_LOCATION_MASTER.Email as SellerEmail,TSPL_Customer_master.GSTNo as BuyerGSTINNo ,TSPL_Customer_master.Customer_Name as BuyerLglNm,TSPL_Customer_master.Alies_name as BuyerTrdNm,Customer_State_Master.GST_STATE_Code  as BuyerPOS,TSPL_Customer_master.Add1 as BuyerAdd1,TSPL_Customer_master.Add2 as BuyerAdd2 ,tspl_city_master.City_Name as BuyerLoc,cast(TSPL_Customer_master.PIN_NO as int) as BuyerPincode,Customer_State_Master.GST_STATE_Code as BuyerStcd,TSPL_Customer_master.Phone1 as BuyerPhone,TSPL_Customer_master.Email as BuyerEmail,TSPL_Customer_Invoice_Detail.SNo as ItemSlNo, 'Y' as ItemIsServc,TSPL_ADDITIONAL_charges.Description AS ItemPrdDesc,TSPL_ADDITIONAL_charges.Sac_code AS ItemHsnCd,
0 as ItemQty,'OTH' as ItemUnit,0 as ItemUnitPrice,TSPL_Customer_Invoice_Detail.Amount as ItemTotAmt,TSPL_Customer_Invoice_Detail.Discount as ItemDiscount,TSPL_Customer_Invoice_Detail.Amount_Less_Discount as ItemAssAmt,
Case when ISNULL( TSPL_Customer_Invoice_Detail.tax1, '' ) = 'IGST' THEN TSPL_Customer_Invoice_Detail.TAX1_Rate 
 when ISNULL(TSPL_Customer_Invoice_Detail.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Detail.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Detail.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX4 = 'SGST' then TSPL_Customer_Invoice_Detail.TAX3_Rate + TSPL_Customer_Invoice_Detail.TAX4_Rate else TSPL_Customer_Invoice_Detail.TAX3_Rate  end)
when ISNULL( TSPL_Customer_Invoice_Detail.tax1, '' ) = 'CGST' AND ISNULL( TSPL_Customer_Invoice_Detail.tax2, '' ) = 'SGST' THEN TSPL_Customer_Invoice_Detail.TAX1_Rate + TSPL_Customer_Invoice_Detail.TAX2_Rate ELSE 0 end  as ItemGstRt,
Case when ISNULL(TSPL_Customer_Invoice_Detail.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Detail.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Detail.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX4 = 'SGST' then ((TSPL_Customer_Invoice_Detail.TAX3_Base_Amt )* (TSPL_Customer_Invoice_Detail.TAX3_Rate ) )/100 when TSPL_Customer_Invoice_Detail.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX4 = 'SGST' then ((TSPL_Customer_Invoice_Detail.TAX4_Base_Amt)* (TSPL_Customer_Invoice_Detail.TAX4_Rate ) )/100 else 0 end) else (case when TSPL_Customer_Invoice_Detail.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX2 = 'SGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt when TSPL_Customer_Invoice_Detail.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX2 = 'SGST' then TSPL_Customer_Invoice_Detail.TAX2_Amt else 0 end) end ItemSgstAmt,
Case when TSPL_Customer_Invoice_Detail.TAX1 = 'IGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt else (case when ISNULL(TSPL_Customer_Invoice_Detail.tax1,'')= 'KKF' or ISNULL(TSPL_Customer_Invoice_Detail.tax2,'')= 'KKF' then(case when TSPL_Customer_Invoice_Detail.TAX3 = 'IGST' then ( (TSPL_Customer_Invoice_Detail.TAX3_Base_Amt)* (TSPL_Customer_Invoice_Detail.TAX3_Rate))/ 100 else 0 end) else 0 end)  end ItemIgstAmt, 
case when ISNULL(TSPL_Customer_Invoice_Detail.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Detail.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Detail.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX4 = 'SGST' then ((TSPL_Customer_Invoice_Detail.TAX3_Base_Amt )* (TSPL_Customer_Invoice_Detail.TAX3_Rate ) )/100 when TSPL_Customer_Invoice_Detail.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX4 = 'SGST' then ((TSPL_Customer_Invoice_Detail.TAX4_Base_Amt )* (TSPL_Customer_Invoice_Detail.TAX4_Rate) )/100 else 0 end) else (case when TSPL_Customer_Invoice_Detail.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX2 = 'SGST' then TSPL_Customer_Invoice_Detail.TAX1_Amt when TSPL_Customer_Invoice_Detail.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Detail.TAX2 = 'SGST' then TSPL_Customer_Invoice_Detail.TAX2_Amt else 0 end) end ItemCgstAmt,
0 as ItemOthChrg,
( case when ISNULL(TSPL_Customer_Invoice_Detail.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Detail.tax2,'')='KKF' then ((TSPL_Customer_Invoice_Detail.Amount_less_Discount)+ ((case when TSPL_Customer_Invoice_Head.TAX3='IGST' then TSPL_Customer_Invoice_Detail.TAX1_AMT + TSPL_Customer_Invoice_Detail.TAX2_AMT +TSPL_Customer_Invoice_Head.TAX3_Amt else TSPL_Customer_Invoice_Detail.TAX1_AMT + TSPL_Customer_Invoice_Detail.TAX2_AMT +TSPL_Customer_Invoice_Detail.TAX3_AMT + TSPL_Customer_Invoice_Detail.TAX4_AMT end))) else ((TSPL_Customer_Invoice_Detail.Amount_less_Discount )+ TSPL_Customer_Invoice_Detail.TAX1_AMT + (case when isnull(TCS1.is_tcs, '')<> 'Y' THEN TSPL_Customer_Invoice_Detail.TAX2_AMT ELSE 0 END) + (case when isnull(TCS2.is_tcs, '')<> 'Y' THEN TSPL_Customer_Invoice_Detail.TAX3_AMT ELSE 0 END)) end)  as ItemTotItemVal,
Case when ISNULL(TSPL_Customer_Invoice_Head.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Head.tax2,'')='KKF' then TSPL_Customer_Invoice_Head.Tax3_BAmount else TSPL_Customer_Invoice_Head.Tax1_BAmount end as ValDtlsAssVal,
Case when ISNULL(TSPL_Customer_Invoice_Head.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Head.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Head.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX4 = 'SGST' then TSPL_Customer_Invoice_Head.TAX3_Amt  when TSPL_Customer_Invoice_Head.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX4 = 'SGST' then TSPL_Customer_Invoice_Head.TAX4_Amt else 0 end) else (case when TSPL_Customer_Invoice_Head.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX2 = 'SGST' then TSPL_Customer_Invoice_Head.TAX1_Amt when TSPL_Customer_Invoice_Head.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX2 = 'SGST' then TSPL_Customer_Invoice_Head.TAX2_Amt else 0 end) end ValDtlsCgstVal,
Case when ISNULL(TSPL_Customer_Invoice_Head.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Head.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Head.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX4 = 'SGST' then TSPL_Customer_Invoice_Head.TAX3_Amt when TSPL_Customer_Invoice_Head.TAX3 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX4 = 'SGST' then TSPL_Customer_Invoice_Head.TAX4_Amt else 0 end) else (case when TSPL_Customer_Invoice_Head.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX2 = 'SGST' then TSPL_Customer_Invoice_Head.TAX1_Amt when TSPL_Customer_Invoice_Head.TAX1 = 'CGST' AND TSPL_Customer_Invoice_Head.TAX2 = 'SGST' then TSPL_Customer_Invoice_Head.TAX2_Amt else 0 end) end ValDtlsSgstVal,
Case when TSPL_Customer_Invoice_Head.TAX1 = 'IGST' then TSPL_Customer_Invoice_Head.TAX1_Amt else (case when ISNULL(TSPL_Customer_Invoice_Head.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Head.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Head.TAX3 = 'IGST' then TSPL_Customer_Invoice_Head.TAX3_Amt else 0 end) else 0 end )end ValDtlsIgstVal, 
TSPL_Customer_Invoice_Head.Discount_Amount as ValDtlsDiscount,
Case when TSPL_Customer_Invoice_Head.TAX1 = 'IGST' then isnull(TSPL_Customer_Invoice_Head.TAX2_Amt,0) else(
Case when ISNULL(TSPL_Customer_Invoice_Head.tax1,'')='KKF' or ISNULL(TSPL_Customer_Invoice_Head.tax2,'')='KKF' then (case when TSPL_Customer_Invoice_Head.TAX3='IGST' then TSPL_Customer_Invoice_Head.TAX4_Amt else (case when    
TSPL_Customer_Invoice_Head.TAX5='TCS' then TSPL_Customer_Invoice_Head.TAX5_Amt else 0 end) end) else (case when TSPL_Customer_Invoice_Head.tax3='TCS' then TSPL_Customer_Invoice_Head.TAX3_Amt else 0 end) end) end as ValDtlsOthChrg,
TSPL_Customer_Invoice_Head.document_total as ValDtlsTotInvVal,
TSPL_Customer_Invoice_Head.RoundOffAmount  as ValDtlsRndOffAmt
from TSPL_Customer_Invoice_Head
Left Outer Join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code  ='" & objCommonVar.CurrentCompanyCode & "'
Left Outer Join TSPL_Customer_master on TSPL_Customer_master.Cust_Code  =TSPL_Customer_Invoice_Head.Customer_Code
left Outer Join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code ='" + strLocation + "'
left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.document_No=TSPL_Customer_Invoice_Head.document_No
left outer join TSPL_ADDITIONAL_charges on TSPL_ADDITIONAL_charges.Code=TSPL_Customer_Invoice_Detail.AddChargeCode
left outer join TSPL_STATE_MASTER as LOCATION_State_Master on LOCATION_State_Master.STATE_CODE  =TSPL_LOCATION_MASTER.State
left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE  =TSPL_Customer_master.State 
left outer join tspl_city_master on tspl_city_master.city_code=TSPL_Customer_master.City_Code
left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code =TSPL_Customer_Invoice_Head.Tax2
left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code =TSPL_Customer_Invoice_Head.Tax3
where TSPL_Customer_Invoice_Head.document_No ='" & strDocNo & "'"


                Dim objResult As Object = ClsEInvoiceOFAPIs.PostAuthTokenNo_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strQry, strLocation, trans, True)
                If objResult IsNot Nothing Then
                    'assign to variable
                    Dim AckNo As String = objResult.SelectToken("AckNo").ToString
                    Dim AckDt As String = objResult.SelectToken("AckDt").ToString
                    Dim Irn As String = objResult.SelectToken("Irn").ToString
                    Dim SignedQRCode As String = objResult.SelectToken("SignedQRCode").ToString
                    clsDBFuncationality.ExecuteNonQuery("update TSPL_Customer_Invoice_Head set  IRN_No ='" & Irn & "',qr_code='" & SignedQRCode & "',ack_no='" & AckNo & "',ack_date='" & clsCommon.GetPrintDate(AckDt, "dd/MMM/yyyy hh:mm tt") & "' where document_No ='" & strDocNo & "'", trans)

                    Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(SignedQRCode)
                    clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_Customer_Invoice_Head", "TSPL_Customer_Invoice_Head.document_No='" & strDocNo & "'", trans)
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
    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal strRefDocNo As String) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = clsCustomerInvoiceHead.PostData(FormId, strDocNo, strRefDocNo, trans)
            If isSaved Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal strRefDocNo As String, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean
        Try


            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If

            Dim obj As clsCustomerInvoiceHead = clsCustomerInvoiceHead.GetData(strDocNo, trans)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If

            clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleReceivable, clsUserMgtCode.mbtnARInvoiceEntry, obj.loc_code, obj.Document_Date, trans)

            If (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            If (obj.On_Hold) Then
                Throw New Exception("Document No " + obj.Document_No + " Is currently On Hold.Can't Post it")
            End If
            If clsCommon.myLen(obj.Customer_Control_AC) <= 0 Then
                Throw New Exception("Customer's Control A/C Not found")
            End If

            Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_Customer_Invoice_Head", "Document_No", strDocNo, trans)
            If isResult = False Then
                trans.Commit()
                Return False
            End If

            ''richa 21 Dec,2018  TEC/02/11/18-000360 create journal entry for opening in case of Credit or debit note as Journal Master table instead of journal master op table
            'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, trans)
            Dim JEWithOPening As Boolean = False
            If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                If clsCommon.myCDate(clsCommon.GetPrintDate(obj.Document_Date, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                    JEWithOPening = True
                End If
            End If

            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
                CreateJournalEntryForOpening(obj, trans, False, FormId, strVoucherNoForRecreatedOnly)
            Else
                CreateGLEntryForAllCases(obj, trans, False, FormId, strVoucherNoForRecreatedOnly)
            End If
            ''-----------------
            '' SMS RElated Function
            ''  CreateSMSContent(obj, trans)
            '' End
            Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceType(obj.Customer_Code, trans)
            ''richa agarwal 31 Dec,2020 check eInvoice Implementation
            If clsCommon.myLen(clsCommon.myCstr(obj.Tax_Group)) > 0 Then
                Dim isTaxTaxable As String = "N"
                Dim LocationCode As String = ""
                '' IsMainPlant removed for RCDFCF 26-DEc-2025 (Ashok Sir)
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RCDFCF") = CompairStringResult.Equal Then
                    LocationCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Location_Master where Loc_Segment_Code='" + obj.loc_code + "'", trans))

                Else
                    LocationCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Location_Master where Loc_Segment_Code='" + obj.loc_code + "' and IsMainPlant=1 ", trans))

                End If
                isTaxTaxable = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'Y' from TSPL_TAX_GROUP_MASTER where Tax_Group_Code ='" & obj.Tax_Group & "' and Is_Tax_Exempted =0 and Tax_Group_Type ='S'", trans))
                ''If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(isTaxTaxable), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.AgainstServiceInvoice), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "I") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True Then
                If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(isTaxTaxable), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.AgainstServiceInvoice), "Y") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True AndAlso obj.Is_Einvoice = 1 Then
                    If clsCustomerInvoiceHead.EInvoice_Implementation(obj.Document_No, LocationCode, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If
            ''------------------------------

            Dim qry As String = "update TSPL_REMITTANCE set Remit_TDS='N' where Document_No='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Update TSPL_Customer_Invoice_Head set Posting_Date='" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") + "',Modify_By='" + objCommonVar.CurrentUserCode + "' , Status=1,EInvoice_Type='" + ECustomerType + "' where Document_No='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_Customer_Invoice_HEAD", "Document_No", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function CancelData(ByVal Doc_No As String, ByVal isCancelByAdmin As Boolean) As Boolean
        '' created by Sanjay date 31-12-2020
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            Dim obj As clsCustomerInvoiceHead = clsCustomerInvoiceHead.GetData(Doc_No, trans)

            If obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If

            ''richa agarwal 01 Jan,2021 check eInvoice Cancel Implementation
            Dim dtirn As DataTable = clsDBFuncationality.GetDataTable("select Einvoice_type,IRN_No,Loc_Code from TSPL_Customer_Invoice_Head where Document_No='" & obj.Document_No & "'", trans)
            If dtirn IsNot Nothing AndAlso dtirn.Rows.Count > 0 Then
                Dim isTaxTaxable As String = "N"
                isTaxTaxable = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'Y' from TSPL_TAX_GROUP_MASTER where Tax_Group_Code ='" & obj.Tax_Group & "' and Is_Tax_Exempted =0 and Tax_Group_Type ='S'", trans))
                If clsCommon.CompairString(clsCommon.myCstr(dtirn.Rows(0)("Einvoice_type")), "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(isTaxTaxable), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.AgainstServiceInvoice), "Y") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True AndAlso Not isCancelByAdmin Then
                    If ClsEInvoiceOFAPIs.EInvoice_Cancellation(obj.Document_No, clsCommon.myCstr(dtirn.Rows(0)("IRN_No")), obj.loc_code, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If
            ''------------------------------

            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_Customer_Invoice_Head", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)

            '' cancel journal master data invoice
            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No ='" & Doc_No & "'"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            ''delete data from multiple tables

            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Detail where Document_No ='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Head where Document_No='" & Doc_No & "'"
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

    Shared Sub CreateSMSContent(ByVal obj As clsCustomerInvoiceHead, ByVal trans As SqlTransaction)
        Dim Form_ID As String = clsUserMgtCode.mbtnARInvoiceEntry
        Dim itemName As String = ""
        Dim DO_Date As Date

        Dim dtCustomerOutstanding As DataTable = Nothing
        Dim dtCrateCan As DataTable = Nothing
        Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + Form_ID + "'", trans)
        If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
            Dim qry As String = "select distinct tspl_customer_master.Customer_Name,substring (tspl_customer_master.Phone1 ,6,10) as MobileNo,tspl_customer_master.Email from tspl_customer_master"
            qry += " left outer join tspl_customer_invoice_head on tspl_customer_invoice_head.Customer_Code=tspl_customer_master.Cust_Code"
            qry += " where 2=2 and len(replace( isnull(substring(tspl_customer_master.Phone1,6,10),''),'_',''))>0 and Cust_Code='" + obj.Customer_Code + "' "
            Dim dtParty As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

            Dim qry1 As String = "select convert(varchar(12),tspl_customer_invoice_head.Document_Date,103) as Do_Date ,tspl_customer_invoice_head.Document_Total            from tspl_customer_invoice_head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.DOCUMENT_No=tspl_customer_invoice_head.Document_No "
            qry1 += "  where tspl_customer_invoice_head.Document_No='" & obj.Document_No & "' "
            Dim dtDocWise As DataTable = clsDBFuncationality.GetDataTable(qry1, trans)


            If dtParty IsNot Nothing AndAlso dtParty.Rows.Count > 0 Then
                If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                    Dim objSMSH As New clsSMSHead()
                    objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Code, obj.Customer_Code)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Name, obj.Customer_Name)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_No, obj.Document_No)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.DocAmount, clsCommon.myFormat(obj.Document_Total))


                    DO_Date = clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy")

                    '' Outstanding Amount
                    dtCustomerOutstanding = getCustomerOutstandingOfAmt_Can_Crate("'" & obj.Customer_Code & "'", clsCommon.GetPrintDate(clsCommon.myCDate(DO_Date).AddDays(-1), "dd/MMM/yyyy"), clsCommon.GetPrintDate(clsCommon.myCDate(DO_Date), "dd/MMM/yyyy"), trans)
                    ''end

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.OutStandingAmt, clsCommon.myFormat(dtCustomerOutstanding.Rows(0)("BalAmt")))

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CRT, clsCommon.myFormat(dtCustomerOutstanding.Rows(0)("CrateClosing")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CAN, clsCommon.myFormat(dtCustomerOutstanding.Rows(0)("CanQtyClosing")))

                    objSMSH.arrMobilNo = New List(Of String)()
                    objSMSH.arrMobilNo.Add(clsCommon.myCstr(dtParty.Rows(0)("MobileNo")))
                    objSMSH.SaveData(Form_ID, objSMSH, trans)
                    objSMSH = Nothing
                End If

            End If

        End If
    End Sub

    'Public Shared Function getCustomerOutstandingOfAmt_Can_Crate(ByVal strCustomer As String, ByVal strfromdate As String, ByVal strtodate As String, ByVal Trans As SqlTransaction) As DataTable
    'Dim dt As DataTable = Nothing
    '    Try
    'Dim ConvRate As String = "ConvRate"
    'Dim BaseQryForCustomer As String = clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, ConvRate, strCustomer, False, strfromdate, strtodate, False, False, True, Trans)

    'Dim BaseQry As String = " Select convert(varchar,FinalGroup.DocDate,103) as Date ,FinalGroup.acode ,max(FinalGroup.AName) as Aname,sum(FinalGroup.DrAmt) as DrAmt,sum(FinalGroup.CrAmt) as Cr_Amt,sum(FinalGroup.CanDebit) as CanDebit,sum(FinalGroup.CanCredit) as CanCredit,sum(FinalGroup.CrateDebit) as CrateDebit,sum(FinalGroup.CrateCredit) as CrateCredit from ( Select final.DocDate , max(ACode) as acode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName, SUM(convert(decimal(18,2),DrAmt*  Final.ConvRate)) as DrAmt, " & _
    '" SUM(convert(decimal(18,2),CrAmt)) as CrAmt,0 as CanDebit,0 as CanCredit,0 as CrateDebit,0 as CrateCredit " & _
    '" FROM ( " & BaseQryForCustomer & " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code " + Environment.NewLine & _
    '" Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " + Environment.NewLine & _
    '"where  CONVERT(DATE,final.DocDate,103) >= '" + strfromdate + "' AND CONVERT(DATE,final.DocDate,103) <= '" + strtodate + "' AND LEN(ACode)>0  GROUP BY CONVERT(DATE,final.DocDate,103)" + Environment.NewLine & _
    '" Union All " & Environment.NewLine & _
    '" select '" + strfromdate + "' as Date,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as DrAmt,0 as Cr_Amt,0 as CanDebit,0 as CanCredit,0 as CrateDebit,0 as CrateCredit  from TSPL_CUSTOMER_MASTER where  TSPL_CUSTOMER_MASTER.Cust_Code in (" & strCustomer & ") " & Environment.NewLine & _
    '" Union All " & Environment.NewLine & _
    '" select '" + strtodate + "' as Date,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as DrAmt,0 as Cr_Amt,0 as CanDebit,0 as CanCredit,0 as CrateDebit,0 as CrateCredit  from TSPL_CUSTOMER_MASTER where  TSPL_CUSTOMER_MASTER.Cust_Code in (" & strCustomer & ")" & Environment.NewLine & _
    '" Union All " & Environment.NewLine & _
    '" select convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code ,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as DrAmt,0 as Cr_Amt,isnull(TSPL_SD_SALE_INVOICE_HEAD.ShippedCan,0) as CanDebit,0 as CanCredit,isnull(TSPL_SD_SALE_INVOICE_HEAD.crate,0) as CrateDebit,0 as CrateCredit from TSPL_SD_SALE_INVOICE_HEAD  " & Environment.NewLine & _
    '" left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " & Environment.NewLine & _
    '" where (isnull(TSPL_SD_SALE_INVOICE_HEAD.ShippedCan,0)>0 or isnull(TSPL_SD_SALE_INVOICE_HEAD.crate,0)>0) and TSPL_SD_SALE_INVOICE_HEAD.Status =1 and CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date ,103) >= '" + strfromdate + "' AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <='" + strtodate + "'  " & Environment.NewLine & _
    '" and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in  (" & strCustomer & ") " & Environment.NewLine & _
    '" Union All " & Environment.NewLine & _
    '" select convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_RETURN_HEAD.Customer_Code ,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as DrAmt,0 as Cr_Amt,0 as CanDebit,isnull(TSPL_SD_SALE_RETURN_HEAD.ShippedCan,0) as CanCredit,0 as CrateDebit,isnull(TSPL_SD_SALE_RETURN_HEAD.CrateQty,0) as CrateCredit from TSPL_SD_SALE_RETURN_HEAD " & Environment.NewLine & _
    '" left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SALE_RETURN_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code " & Environment.NewLine & _
    '"  where (isnull(TSPL_SD_SALE_RETURN_HEAD.ShippedCan,0)>0 or isnull(TSPL_SD_SALE_RETURN_HEAD.CrateQty,0)>0) and TSPL_SD_SALE_RETURN_HEAD.Status =1 and CONVERT(DATE,Document_Date ,103) >= '" + strfromdate + "' AND CONVERT(DATE,Document_Date,103) <= '" + strtodate + "' " & Environment.NewLine & _
    '" and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" & strCustomer & ") " & Environment.NewLine & _
    '" Union All " & Environment.NewLine & _
    '" select convert(date,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date,103) as Document_Date,TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code ,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as DrAmt,0 as Cr_Amt,0 as CanDebit,isnull(TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.CanQtyRec,0)+isnull(TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.CanAdjustment,0) as CanCredit,0 as CrateDebit,isnull(TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.CrateQtyRecd,0)+isnull(TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Adjustment,0) as CrateCredit from TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE " & Environment.NewLine & _
    '" left outer join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE  on TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No " & Environment.NewLine & _
    '" left outer join TSPL_CUSTOMER_MASTER on TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " & Environment.NewLine & _
    '" where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted =1 and CONVERT(DATE,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date ,103) >= '" + strfromdate + "' AND CONVERT(DATE,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date,103) <='" + strtodate + "'  " & Environment.NewLine & _
    '" and TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code in (" & strCustomer & ") " & Environment.NewLine & _
    '    " ) FinalGroup group by FinalGroup.DocDate ,FinalGroup.acode order by FinalGroup.DocDate desc"
    '        Return clsDBFuncationality.GetDataTable(BaseQry, Trans)

    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    '    Return dt
    'End Function
    ''richa agarwal 23 Nov,2018
    Public Shared Function getCustomerOutstandingAmtWithOPeningAndClosing(ByVal strCustomer As String, ByVal strfromdate As String, ByVal strtodate As String, ByVal ConvRate As String, ByVal Trans As SqlTransaction) As String
        Try
            Dim BaseQryForCustomer As String = Nothing
            Dim BaseQryForCustomerforOpening As String = Nothing
            Dim BaseQry As String = Nothing

            BaseQryForCustomer = clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, ConvRate, strCustomer, False, strfromdate, strtodate, False, False, False, Trans)
            BaseQryForCustomerforOpening = clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, ConvRate, strCustomer, True, strfromdate, strtodate, False, False, False, Trans)

            BaseQry = " Select '" + clsCommon.GetPrintDate(strtodate, "dd/MM/yyyy") + "' as DocDate, ACode,  MAX(AName) as AName, SUM(convert(decimal(18,2),OpngBal)) as OpngBal, SUM(convert(decimal(18,2),DrAmt)) as DrAmt, SUM(convert(decimal(18,2),CrAmt)) as CrAmt,  ( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))  as BalAmt  From (" + Environment.NewLine & _
            "  Select max(DocDate) as DocDate, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName,  '' as CurrencyCode, null as ConvRate, SUM(DrAmt*ConvRate)-SUM(CrAmt) as OpngBal, 0 as DrAmt, 0 as CrAmt from  ( " + BaseQryForCustomerforOpening + " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code where  CONVERT(DATE,final.DocDate,103) < '" + strfromdate + "' AND LEN(ACode)>0 AND TSPL_CUSTOMER_MASTER.Status='N'  GROUP BY ACode" + Environment.NewLine & _
            Environment.NewLine + " UNION ALL" + Environment.NewLine & _
            " Select  max(DocDate) as DocDate, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName,  MAX(Final.Currency_Code) as Currency_Code, MAX(Final.ConvRate) as ConvRate, 0 as OpngBal, SUM(convert(decimal(18,2),DrAmt*  Final.ConvRate)) as DrAmt, " & Environment.NewLine & _
            " SUM(convert(decimal(18,2),CrAmt)) as CrAmt FROM ( " + BaseQryForCustomer + " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code " + Environment.NewLine & _
            " Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " + Environment.NewLine & _
            "where  CONVERT(DATE,final.DocDate,103) >= '" + strfromdate + "' AND CONVERT(DATE,final.DocDate,103) <= '" + strtodate + "' AND LEN(ACode)>0 AND TSPL_CUSTOMER_MASTER.Status='N' GROUP BY ACode ) XXX GROUP BY ACode  " + Environment.NewLine

            Return BaseQry
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getCustomerOutstandingCrateWithOPeningAndClosing(ByVal strCustomer As String, ByVal strfromdate As String, ByVal strtodate As String, ByVal ConvRate As String) As String
        Try
            Dim BaseQry As String = Nothing

            BaseQry = " ( Select convert(varchar,Doc_Date,103) as Doc_Date,Customer_Code,Customer_Name, OpencrateQty, CrateQtyRecd ,CrateOutQty, CrateAdjQty ,SUM(CrateQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code) as CrateQtyClosing,OpenCanQty , CanQtyRecd  ,CanOutQty , CanAdjQty ,SUM(CanQtyClosing ) OVER (Partition BY Customer_Code ORDER BY Customer_Code) as CanQtyClosing from ( " & Environment.NewLine & _
            " select  pp.Doc_Date  as Doc_Date,TSPL_CUSTOMER_MASTER.Route_No,TSPL_CUSTOMER_MASTER.Route_Desc,pp.Vehicle_Code,tspl_vehicle_master.Number as Vehicle_Name ,pp.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,pp.OpencrateQty as OpencrateQty,pp.OpenJaaliQty  as OpenJaaliQty,pp.OpenBoxQty  as OpenBoxQty,pp.OpenCanQty  as OpenCanQty,pp.CrateQtyRecd  as CrateQtyRecd,pp.JaaliQtyRecd  as JaaliQtyRecd,pp.BoxQtyRecd  as BoxQtyRecd ,pp.CanQtyRecd  as CanQtyRecd,pp.CrateOutQty  as CrateOutQty,pp.jaaliOutQty  as jaaliOutQty ,pp.boxOutQty  as boxOutQty,pp.CanOutQty  as CanOutQty ,pp.CrateQtyClosing as CrateQtyClosing, pp.JaaliQtyClosing as  JaaliQtyClosing, pp.BoxQtyClosing as BoxQtyClosing, pp.CanQtyClosing as CanQtyClosing,pp.CrateAdjQty , pp.JaaliAdjQty  , pp.BoxAdjQty , pp.CanAdjQty from ( " & Environment.NewLine & _
            " select  max(convert(date,Doc_Date,103))  as Doc_Date,max(xx.Vehicle_Code) as Vehicle_Code,xx.Customer_Code,sum(xx.OpencrateQty) as OpencrateQty,sum(xx.OpenJaaliQty ) as OpenJaaliQty,sum(xx.OpenBoxQty )  as OpenBoxQty,sum(xx.OpenCanQty )  as OpenCanQty,sum(xx.CrateQtyRecd) as CrateQtyRecd,sum(xx.JaaliQtyRecd) as JaaliQtyRecd,sum(xx.BoxQtyRecd) as BoxQtyRecd,sum(xx.CanQtyRecd) as CanQtyRecd,sum(xx.CrateOutQty ) as CrateOutQty,sum(xx.jaaliOutQty ) as jaaliOutQty ,sum(xx.boxOutQty ) as boxOutQty,sum(xx.CanOutQty ) as CanOutQty, sum(xx.CrateAdjQty ) as CrateAdjQty ,sum(xx.JaaliAdjQty )as JaaliAdjQty  ,sum(xx.BoxAdjQty )  as BoxAdjQty ,sum(xx.CanAdjQty )  as CanAdjQty,(sum(xx.OpencrateQty)+sum(xx.CrateOutQty )-sum(xx.CrateQtyRecd)-sum(xx.CrateAdjQty )) as CrateQtyClosing, (sum(xx.OpenJaaliQty)+sum(xx.jaaliOutQty)-sum(xx.JaaliQtyRecd)-sum(xx.JaaliAdjQty )) as JaaliQtyClosing, (sum(xx.OpenBoxQty)+sum(xx.boxOutQty )-sum(xx.BoxQtyRecd)-sum(xx.BoxAdjQty )) as BoxQtyClosing  , (sum(xx.OpenCanQty)+sum(xx.CanOutQty )-sum(xx.CanQtyRecd)-sum(xx.CanAdjQty )) as CanQtyClosing  from (select  max(convert(date,'" + strfromdate + "' ,103)) as Doc_Date,max(Opening.Vehicle_Code) as Vehicle_Code , Opening.Customer_Code ,sum(Opening.OpencrateQty*Type) as OpencrateQty,sum(Opening.OpenJaaliQty*Type ) as OpenJaaliQty,sum(Opening.OpenBoxQty *Type)  as OpenBoxQty,sum(Opening.OpenCanQty *Type)  as OpenCanQty,sum(Opening.CrateQtyRecd *Type) as CrateQtyRecd,sum(Opening.JaaliQtyRecd*Type ) as JaaliQtyRecd,sum(Opening.BoxQtyRecd*Type ) as BoxQtyRecd,sum(Opening.CanQtyRecd*Type ) as CanQtyRecd,sum(Opening.CrateOutQty*Type ) as CrateOutQty,sum(Opening.jaaliOutQty*Type ) as jaaliOutQty ,sum(Opening.boxOutQty*Type ) as boxOutQty,sum(Opening.CanOutQty*Type ) as CanOutQty,sum(Opening.CrateAdjQty*Type ) as CrateAdjQty,sum(Opening.JaaliAdjQty*Type ) as JaaliAdjQty,sum(Opening.BoxAdjQty *Type) as  BoxAdjQty ,sum(Opening.CanAdjQty *Type) as  canAdjQty from  ( select TSPL_SD_SHIPMENT_HEAD.Document_Date    as Document_Date,TSPL_SD_SHIPMENT_HEAD.Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,1 as Type  ,'O' as Type1,TSPL_SD_SHIPMENT_HEAD.Crate as OpencrateQty,TSPL_SD_SHIPMENT_HEAD.jaali as OpenJaaliQty, TSPL_SD_SHIPMENT_HEAD.box  as OpenBoxQty , TSPL_SD_SHIPMENT_HEAD.ShippedCAN  as OpenCanQty  ,0 as CrateQtyRecd,0  as JaaliQtyRecd ,0 as BoxQtyRecd ,0 as CanQtyRecd ,0 as CrateOutQty,0 as jaaliOutQty,0 as boxOutQty ,0 as CanOutQty,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty  ,0 as  CanAdjQty  from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS'  union all  select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as OpencrateQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty   as OpenjaaliQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty  as OpenboxQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty as OpenCanQty , 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O'  union all select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment,0)  as OpencrateQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment,0)    as OpenjaaliQty , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment,0)  as OpenboxQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment,0)  as OpenCanQty, 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I'  )as Opening WHERE convert(date,Document_Date,103)<(convert(date,'" + strfromdate + "' ,103)) group by Customer_Code " & Environment.NewLine & _
            " UNION All " & Environment.NewLine & _
            " select Document_Date,Vehicle_Code,Customer_Code,0 as OpencrateQty,0 as OpenjaaliQty ,0 as OpenboxQty,0 as OpenCanQty,Case When [Type]=1 Then CrateQtyRecd Else 0 End as CrateQtyRecd,Case When [Type]=1 Then JaaliQtyRecd Else 0 End as JaaliQtyRecd,Case When [Type]=1 Then BoxQtyRecd Else 0 End as BoxQtyRecd,Case When [Type]=1 Then CANQtyRec Else 0 End as CANQtyRec, " & Environment.NewLine & _
            " Case When [Type]=-1 Then CrateOutQty Else 0 End as CrateOutQty,Case When [Type]=-1 Then jaaliOutQty Else 0 End as jaaliOutQty,Case When [Type]=-1 Then boxOutQty Else 0 End as boxOutQty,Case When [Type]=-1 Then CanOutQty Else 0 End as CanOutQty,Case When [Type]=1 Then CrateAdjQty Else 0 End as CrateAdjQty,Case When [Type]=1 Then JaaliAdjQty Else 0 End as JaaliAdjQty,Case When [Type]=1 Then BoxAdjQty Else 0 End as BoxAdjQty,Case When [Type]=1 Then CANAdjustment Else 0 End as CANAdjustment " & Environment.NewLine & _
            " from ((select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCanQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE" & Environment.NewLine & _
            " left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  " & Environment.NewLine & _
            "where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I') " & Environment.NewLine & _
            " union all" & Environment.NewLine & _
            " (select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCANQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE" & Environment.NewLine & _
            " left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No " & Environment.NewLine & _
            " where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O' " & Environment.NewLine & _
            " union all " & Environment.NewLine & _
            " select TSPL_SD_SHIPMENT_HEAD.Document_Date   as Document_Date,TSPL_SD_SHIPMENT_HEAD.Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,-1 as Type  ,'O' as Type1,0 as OpencrateQty,0  as OpenBoxQty ,0 as OpenJaaliQty,0 as OpenCANQty, 0 as CrateQtyRecd, 0 JaaliQtyRecd ,  0  as BoxQtyRecd,0 CanQtyRecd  ,TSPL_SD_SHIPMENT_HEAD.Crate as CrateOutQty,jaali  as jaaliOutQty,TSPL_SD_SHIPMENT_HEAD.Box as boxOutQty,TSPL_SD_SHIPMENT_HEAD.ShippedCAN as CanOutQty ,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty,0 as  CanAdjQty    from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS')  ) as Closing " & Environment.NewLine & _
            " WHERE convert(date,Document_Date ,103)>= convert(date,'" + strfromdate + "' ,103) AND convert(date,Document_Date,103)<=convert(date,'" + strtodate + "' ,103) " & Environment.NewLine & _
            " ) as xx where 2=2   and xx.Customer_Code  in (" & strCustomer & ") " & Environment.NewLine & _
            " GROUP BY Customer_Code " & Environment.NewLine & _
            " ) as pp   left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =pp.Customer_Code " & Environment.NewLine & _
            " left join tspl_vehicle_master on tspl_vehicle_master.vehicle_id=pp.vehicle_code where 2=2  ) YYY )"


            Return BaseQry
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    ''richa BHA/19/09/18-000561
    Public Shared Function getCustomerOutstandingOfAmt_Can_Crate(ByVal strCustomer As String, ByVal strfromdate As String, ByVal strtodate As String, ByVal Trans As SqlTransaction) As DataTable
        Dim dt As DataTable = Nothing
        Try
            Dim BaseQryForCustomer As String = Nothing
            Dim BaseQryForCustomerforOpening As String = Nothing
            Dim BaseQry As String = Nothing
            Dim ConvRate As String = "ConvRate"

            Dim BaseqryfromDate_Amount As String = Nothing
            Dim BaseqryToDate_Amount As String = Nothing

            Dim BaseqryfromDate_Crate As String = Nothing
            Dim BaseqryToDate_Crate As String = Nothing
            '' for Customer outstanding
            BaseqryfromDate_Amount = getCustomerOutstandingAmtWithOPeningAndClosing(strCustomer, strfromdate, strfromdate, ConvRate, Trans)
            BaseqryToDate_Amount = getCustomerOutstandingAmtWithOPeningAndClosing(strCustomer, strtodate, strtodate, ConvRate, Trans)
            ''for crate outstanding
            BaseqryfromDate_Crate = getCustomerOutstandingCrateWithOPeningAndClosing(strCustomer, strfromdate, strfromdate, ConvRate)
            BaseqryToDate_Crate = getCustomerOutstandingCrateWithOPeningAndClosing(strCustomer, strtodate, strtodate, ConvRate)

            BaseQry = "  Select convert(varchar,FinalGroup.DocDate,103) as Date ,FinalGroup.acode as Cust_code ,max(FinalGroup.AName) as Name,sum(FinalGroup.OpngBal) as OpngBal,sum(FinalGroup.DrAmt) as DrAmt,sum(FinalGroup.CrAmt) as CrAmt,sum(FinalGroup.BalAmt) as BalAmt,sum(FinalGroup.CrateOpng) as CrateOpng,sum(FinalGroup.CrateReceived) as CrateReceived,sum(FinalGroup.CrateIssue) as CrateIssue,sum(FinalGroup.CrateClosing) as CrateClosing,sum(FinalGroup.OpenCanQty) as OpenCanQty ,sum(FinalGroup.CanQtyRecd) as CanQtyRecd  ,sum(FinalGroup.CanOutQty) as CanOutQty ,sum(FinalGroup.CanAdjQty) as CanAdjQty ,sum(FinalGroup.CanQtyClosing) as CanQtyClosing  from  ( " & Environment.NewLine & _
            " Select  DocDate, ACode, AName, OpngBal, DrAmt, CrAmt, BalAmt,0 as CrateOpng,0 as CrateReceived,0 as CrateIssue,0 as CrateClosing,0 as OpenCanQty ,0 as CanQtyRecd  ,0 as CanOutQty ,0 as CanAdjQty ,0 as CanQtyClosing  from (" & BaseqryfromDate_Amount & ") Z " & Environment.NewLine & _
            " Union All " & Environment.NewLine & _
            " Select  DocDate, ACode, AName, OpngBal, DrAmt, CrAmt, BalAmt,0 as CrateOpng,0 as CrateReceived,0 as CrateIssue,0 as CrateClosing,0 as OpenCanQty ,0 as CanQtyRecd  ,0 as CanOutQty ,0 as CanAdjQty ,0 as CanQtyClosing  from (" & BaseqryToDate_Amount & ") X " & Environment.NewLine & _
            " Union All " & Environment.NewLine & _
            " Select Doc_Date,Customer_Code,Customer_Name,0 as OpngBal, 0 as DrAmt, 0 as  CrAmt, 0 as  BalAmt,  OpencrateQty, CrateQtyRecd ,CrateOutQty, CrateQtyClosing ,OpenCanQty , CanQtyRecd  ,CanOutQty , CanAdjQty ,CanQtyClosing from (" & BaseqryfromDate_Crate & ") Y " & Environment.NewLine & _
            " Union All " & Environment.NewLine & _
            " Select Doc_Date,Customer_Code,Customer_Name,0 as OpngBal, 0 as DrAmt, 0 as  CrAmt, 0 as  BalAmt,  OpencrateQty, CrateQtyRecd ,CrateOutQty, CrateQtyClosing ,OpenCanQty , CanQtyRecd  ,CanOutQty , CanAdjQty ,CanQtyClosing from (" & BaseqryToDate_Crate & ") S " & Environment.NewLine & _
            " Union All " & Environment.NewLine & _
            " select '" + clsCommon.GetPrintDate(strfromdate, "dd/MM/yyyy") + "' as Date,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as OpngBal, 0 as DrAmt, 0 as  CrAmt, 0 as  BalAmt,0 as CrateOpng,0 as CrateReceived,0 as CrateIssue,0 as CrateClosing,0 as OpenCanQty ,0 as CanQtyRecd  ,0 as CanOutQty ,0 as CanAdjQty ,0 as CanQtyClosing  from TSPL_CUSTOMER_MASTER where  TSPL_CUSTOMER_MASTER.Cust_Code in (" & strCustomer & ") " & Environment.NewLine & _
            " Union All " & Environment.NewLine & _
            " select '" + clsCommon.GetPrintDate(strtodate, "dd/MM/yyyy") + "'   as Date ,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name  as Name ,0 as OpngBal, 0 as DrAmt, 0 as  CrAmt, 0 as  BalAmt,0 as CrateOpng,0 as CrateReceived,0 as CrateIssue,0 as CrateClosing,0 as OpenCanQty ,0 as CanQtyRecd  ,0 as CanOutQty ,0 as CanAdjQty ,0 as CanQtyClosing   from TSPL_CUSTOMER_MASTER where  TSPL_CUSTOMER_MASTER.Cust_Code in (" & strCustomer & ")" & Environment.NewLine & _
                " ) FinalGroup group by FinalGroup.DocDate ,FinalGroup.acode order by FinalGroup.DocDate desc"

            Return clsDBFuncationality.GetDataTable(BaseQry, Trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return dt
    End Function
    ''--------

    Private Shared Function GetTaxAmt(ByVal objPIDetail As clsCustomerInvoiceDetail, ByVal tans As SqlTransaction) As Double
        Dim dblTotalTax As Double = 0
        Dim isTaxRecoverable As Boolean = False
        If objPIDetail.TAX1_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX1, tans) Then
            dblTotalTax += objPIDetail.TAX1_Amt
        End If
        If objPIDetail.TAX2_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX2, tans) Then
            dblTotalTax += objPIDetail.TAX2_Amt
        End If
        If objPIDetail.TAX3_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX3, tans) Then
            dblTotalTax += objPIDetail.TAX3_Amt
        End If
        If objPIDetail.TAX4_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX4, tans) Then
            dblTotalTax += objPIDetail.TAX4_Amt
        End If
        If objPIDetail.TAX5_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX5, tans) Then
            dblTotalTax += objPIDetail.TAX5_Amt
        End If
        If objPIDetail.TAX6_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX6, tans) Then
            dblTotalTax += objPIDetail.TAX6_Amt
        End If
        If objPIDetail.TAX7_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX7, tans) Then
            dblTotalTax += objPIDetail.TAX7_Amt
        End If
        If objPIDetail.TAX8_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX8, tans) Then
            dblTotalTax += objPIDetail.TAX8_Amt
        End If
        If objPIDetail.TAX9_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX9, tans) Then
            dblTotalTax += objPIDetail.TAX9_Amt
        End If
        If objPIDetail.TAX10_Amt > 0 AndAlso Not clsTaxMaster.IsTaxRecoverableAC(objPIDetail.TAX10, tans) Then
            dblTotalTax += objPIDetail.TAX10_Amt
        End If
        Return dblTotalTax
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
    Public Shared Function DeleteData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = False
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Document No not found to Delete")
        End If
        Dim obj As clsCustomerInvoiceHead = clsCustomerInvoiceHead.GetData(strDocNo, trans)
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
            Try
                clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleReceivable, clsUserMgtCode.mbtnARInvoiceEntry, obj.loc_code, obj.Document_Date, trans)

                If obj.Status = ERPTransactionStatus.Approved Then
                    Throw New Exception("Already Post on :" + obj.Posting_Date)
                End If
                'If (clsCommon.myLen(obj.Posting_Date) > 0) Then
                '    Throw New Exception("Already Post on :" + obj.Posting_Date)
                'End If
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
                'clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_Customer_Invoice_Detail", "Document_No", trans)

                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
                Dim qry As String = "delete from TSPL_Customer_Invoice_Detail where Document_No='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)

                ''changes done by richa agarwal in journal master query from 'AR-CN' to 'AR-CR'
                '' changed by Panch Raj against Ticket No: BM00000008161
                '' delete journal details
                'qry = "delete from TSPL_JOURNAL_DETAILS where Journal_No in (select Journal_No from TSPL_JOURNAL_MASTER where Source_Code in ('AR-IN','AR-CR','AR-DN') and Source_Doc_No='" + strDocNo + "')"
                'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'qry = "delete from TSPL_JOURNAL_MASTER   where Source_Code in ('AR-IN','AR-CR','AR-DN') and Source_Doc_No='" + strDocNo + "'"
                'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)


                '' TO DELETE DATA FROM JOURNAL MASTER 
                Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where TSPL_JOURNAL_MASTER.Source_Code in ('AR-IN','AR-DN','AR-CR','GL-JE') AND  TSPL_JOURNAL_MASTER.Source_Doc_No='" + strDocNo + "'", trans)
                If clsCommon.myLen(VoucherNo) > 0 Then
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                    qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If

                '' TO DELETE DATA FROM JOURNAL MASTER OPENING TABLE ''richa TEC/02/11/18-000360 24 Dec,2018
                Dim VoucherNoOP As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER_OP where TSPL_JOURNAL_MASTER_OP.Source_Code in ('AR-IN','AR-DN','AR-CR','GL-JE') AND  TSPL_JOURNAL_MASTER_OP.Source_Doc_No='" + strDocNo + "'", trans)
                If clsCommon.myLen(VoucherNoOP) > 0 Then
                    qry = "delete from TSPL_JOURNAL_DETAILS_OP where Voucher_No ='" + VoucherNoOP + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    qry = "delete from TSPL_JOURNAL_MASTER_OP where Voucher_No ='" + VoucherNoOP + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If

                'Ticket No-TEC/06/09/19-001003,Save Deleted data ,sanjay
                'clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_Customer_Invoice_Head", "Document_No", trans)

                qry = "delete from TSPL_REMITTANCE where Document_No='" + strDocNo + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_Customer_Invoice_Head where Document_No='" + strDocNo + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                isSaved = isSaved AndAlso clsCustomFieldValues.DeleteData(obj.Form_ID, strDocNo, trans)
                'If (isSaved) Then
                '    trans.Commit()
                'Else
                '    trans.Rollback()
                'End If
            Catch ex As Exception
                'trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function
    ''richa agarwal 22 Jan,2019 TEC/22/01/19-000406
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
        Return ReverseAndUnpost(strCode, trans, True)
    End Function
    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal trans As SqlTransaction, ByVal isCheckRefenceDoc As Boolean) As Boolean
        ' Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim Qry As String = "select TSPL_Customer_Invoice_Head .Status from TSPL_Customer_Invoice_Head where 2=2 and Document_No='" + strCode + "'"
            If isCheckRefenceDoc Then
                Qry += " and (isnull(TSPL_Customer_Invoice_Head.AgainstScrapReturn,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Sale_Return_No,'')='' and  ISNULL (TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,'')='' and isnull(TSPL_Customer_Invoice_Head.Against_Asset_Disposal,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Sale_No,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Security_Receipt_No,'')='' and isnull(TSPL_Customer_Invoice_Head.Against_Service_Visit_Code,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Subsidy_No,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_VCGL,'')='' and ISNULL (TSPL_Customer_Invoice_Head.AgainstScrap,'')='')  "
            End If
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            Qry = "select TSPL_Customer_Invoice_Head .Document_No from TSPL_Customer_Invoice_Head where 2=2 and Document_No='" + strCode + "' and Status='1' "
            If isCheckRefenceDoc Then
                Qry += "and (isnull(TSPL_Customer_Invoice_Head.AgainstScrapReturn,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Sale_Return_No,'')='' and  ISNULL (TSPL_Customer_Invoice_Head.Against_MCC_Material_Sale_Return,'')='' and isnull(TSPL_Customer_Invoice_Head.Against_Asset_Disposal,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Sale_No,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Security_Receipt_No,'')='' and isnull(TSPL_Customer_Invoice_Head.Against_Service_Visit_Code,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_Subsidy_No,'')='' and ISNULL (TSPL_Customer_Invoice_Head.Against_VCGL,'')='' and ISNULL (TSPL_Customer_Invoice_Head.AgainstScrap,'')='')  "
            End If

            Dim dt7 As DataTable = clsDBFuncationality.GetDataTable("select Document_Date,Loc_Code from TSPL_Customer_Invoice_Head  where  Document_No ='" + strCode + "'", trans)
            If dt7 IsNot Nothing AndAlso dt7.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleReceivable, clsUserMgtCode.mbtnARInvoiceEntry, clsCommon.myCstr(dt7.Rows(0)("Loc_Code")), clsCommon.myCDate(dt7.Rows(0)("Document_Date")), trans)
            End If

            Dim strcustomerInvocieNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(Qry, trans))
            If clsCommon.myLen(strcustomerInvocieNo) <= 0 Then
                Throw New Exception("Transaction cannot be Reverse becuase it has created Against Invoice No.")
            End If

            Dim dt As DataTable = Nothing

            '' to check Customer Incentive entry
            Qry = "select Doc_Code  from TSPL_CUSTOMER_INCENTIVE_DETAIL where TR_Code in (select RefDocNo from TSPL_Customer_Invoice_Head  where RefDocType = 'CUS_INC_ENT' and Document_No ='" & strCode & "')"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Qry = "Current document is used in following Customer Incentive Entry -"
                For Each dr As DataRow In dt.Rows
                    Qry += Environment.NewLine + clsCommon.myCstr(dr("Doc_Code"))
                Next
                Throw New Exception(Qry)
            End If


            '' to check receipt header ERO/25/09/19-001039
            Qry = "select Receipt_No from TSPL_RECEIPT_HEADER where isnull(IsChkReverse,'') ='N' and Applied_Receipt ='" + strCode + "' "
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Qry = "Current document is used in following Receipt -"
                For Each dr As DataRow In dt.Rows
                    Qry += Environment.NewLine + clsCommon.myCstr(dr("Receipt_No"))
                Next
                Throw New Exception(Qry)
            End If


            '' to check receipt detail 
            Qry = "select distinct tspl_receipt_detail.Receipt_No from TSPL_RECEIPT_HEADER left outer join tspl_receipt_detail on tspl_receipt_header.Receipt_no=tspl_receipt_detail.Receipt_no where isnull(tspl_receipt_header.IsChkReverse,'') ='N' and tspl_receipt_detail.Document_No ='" + strCode + "' "
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Qry = "Current document is used in following Receipt -"
                For Each dr As DataRow In dt.Rows
                    Qry += Environment.NewLine + clsCommon.myCstr(dr("Receipt_No"))
                Next
                Throw New Exception(Qry)
            End If

            '' TO DELETE DATA FROM JOURNAL MASTER 
            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where TSPL_JOURNAL_MASTER.Source_Code in ('AR-IN','AR-DN','AR-CR','GL-JE') AND  TSPL_JOURNAL_MASTER.Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            '' TO DELETE DATA FROM JOURNAL MASTER OPENING TABLE ''richa TEC/02/11/18-000360 24 Dec,2018
            Dim VoucherNoOP As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER_OP where TSPL_JOURNAL_MASTER_OP.Source_Code in ('AR-IN','AR-DN','AR-CR','GL-JE') AND  TSPL_JOURNAL_MASTER_OP.Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNoOP) > 0 Then
                Qry = "delete from TSPL_JOURNAL_DETAILS_OP where Voucher_No ='" + VoucherNoOP + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER_OP where Voucher_No ='" + VoucherNoOP + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            ''Change status to unpost
            Qry = "update TSPL_Customer_Invoice_Head set Status=0 where Document_No='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_Customer_Invoice_HEAD", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    '---------------------------------------

    Public Shared Function GetSecurityDepositType() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Value", GetType(String))

        dt.Rows.Add("Select", "")
        dt.Rows.Add("Security", "S")
        dt.Rows.Add("Crate Security", "C")
        dt.Rows.Add("Refrigerator Security", "R")
        dt.Rows.Add("Others", "O")
        Return dt
    End Function
    ''richa TEC/02/11/18-000360 on 21 Dec,2018
    Public Shared Function CreateJournalEntryForOpening(ByVal obj As clsCustomerInvoiceHead, ByVal trans As SqlTransaction, ByVal isForUnpostedTransaction As Boolean, ByVal FormId As String, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean
        Try
            Dim qry As String = "select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No='" + obj.Document_No + "' "
            Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))

            If strVoucherNoForRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreatedOnly) > 0 Then
                strVoucherNo = strVoucherNoForRecreatedOnly
            End If

            ''===========these column is for GL entry=========do for multicurrency==if currency is differ form basecurrency=======15/04/2015==(Monika)================================
            Dim coll As New Hashtable()
            If clsCommon.myLen(obj.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(obj.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
                coll = New Hashtable()
                clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE)
                clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
                clsCommon.AddColumnsForChange(coll, "ConvRateOld", obj.ConvRate)
            End If
            ''===============================================================================================


            Dim isTaxRecoverable As Boolean = False
            Dim NonRecovTaxAmount As Decimal = 0
            Dim Tax_Liability_Account As String = ""
            Dim strEntryDesc As String = ""
            Dim strSrcType As String = ""
            Dim strSrcDesc As String = ""
            Dim strRemarks As String = ""
            If (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                strEntryDesc = "Customer Opening"
                strSrcType = "GL-JE"
                strSrcDesc = "AR Debit"
                strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  "
            ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                strEntryDesc = "Customer Opening"
                strSrcType = "GL-JE"
                strSrcDesc = "AR Credit Note"
                strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  "

            End If
            Dim strQ As String = Nothing
            Dim ArryLst As ArrayList = New ArrayList()
            If (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                Dim strReceivable_Control_acct As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Receivable_Control_acct ,'') as Receivable_Control_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                If clsCommon.myLen(strReceivable_Control_acct) = 0 Then
                    Throw New Exception("Please set Receivable Control Account for customer - " + obj.Customer_Code)
                End If
                strReceivable_Control_acct = clsERPFuncationality.ChangeGLAccountLocationSegment(strReceivable_Control_acct, obj.loc_code, True, trans)
                Dim Acc1() As String = {strReceivable_Control_acct, (obj.Document_Total)}
                ArryLst.Add(Acc1)


                Dim strCustomer_Opening_Clearing_AC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Customer_Opening_Clearing_AC ,'') as Customer_Opening_Clearing_AC from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                If clsCommon.myLen(strCustomer_Opening_Clearing_AC) = 0 Then
                    Throw New Exception("Please set Customer Opening Clearing Account for customer - " + obj.Customer_Code)
                End If
                strCustomer_Opening_Clearing_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strCustomer_Opening_Clearing_AC, obj.loc_code, True, trans)
                Dim Acc2() As String = {strCustomer_Opening_Clearing_AC, (obj.Document_Total) * -1}
                ArryLst.Add(Acc2)

            ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                Dim strReceivable_Control_acct As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Receivable_Control_acct ,'') as Receivable_Control_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                If clsCommon.myLen(strReceivable_Control_acct) = 0 Then
                    Throw New Exception("Please set Receivable Control Account for customer - " + obj.Customer_Code)
                End If
                strReceivable_Control_acct = clsERPFuncationality.ChangeGLAccountLocationSegment(strReceivable_Control_acct, obj.loc_code, True, trans)
                Dim Acc1() As String = {strReceivable_Control_acct, (obj.Document_Total) * -1}
                ArryLst.Add(Acc1)


                Dim strCustomer_Opening_Clearing_AC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Customer_Opening_Clearing_AC ,'') as Customer_Opening_Clearing_AC from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                If clsCommon.myLen(strCustomer_Opening_Clearing_AC) = 0 Then
                    Throw New Exception("Please set Customer Opening Clearing Account for customer - " + obj.Customer_Code)
                End If
                strCustomer_Opening_Clearing_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strCustomer_Opening_Clearing_AC, obj.loc_code, True, trans)
                Dim Acc2() As String = {strCustomer_Opening_Clearing_AC, (obj.Document_Total)}
                ArryLst.Add(Acc2)
            End If


            '' clsJournalMaster.FunGrnlEntryWithTrans(True, 0, "", "N", LocSegmentCode, True, isForUnpostedTransaction, strVoucherNo, trans, clsCommon.myCDate(dtReceipt.Rows(0)("Receipt_Post_Date")), strRemarks, "GL-JE", strSourceTypeDesc, strDocNo, "", IIf(clsCommon.CompairString(clsCommon.myCstr(dtReceipt.Rows(0)("Receipt_Type")), "P") = CompairStringResult.Equal, "C", "O"), dtReceipt.Rows(0)("Cust_Code").ToString(), dtReceipt.Rows(0)("Customer_Name").ToString(), objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , dtReceipt.Rows(0)("Reference").ToString(), dtReceipt.Rows(0)("Narration").ToString(), coll)
            clsJournalMaster.FunGrnlEntryWithTrans(True, 0, "", "N", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, obj.Document_Date, strEntryDesc, strSrcType, strSrcDesc, obj.Document_No, obj.Description, "C", obj.Customer_Code, obj.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, "", strRemarks, Nothing, coll)

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
        Return True
    End Function
    ''-------------------------------------------------
    Public Shared Function CreateGLEntryForAllCases(ByVal obj As clsCustomerInvoiceHead, ByVal trans As SqlTransaction, ByVal isForUnpostedTransaction As Boolean, ByVal FormId As String, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean
        Try

            'Ticket no- TEC/04/04/19-000465 Sanjay Setting-AllowPurchaseAccounting
            Dim RecoControlACC As String = ""
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 0 Then
            RecoControlACC = "I"
        End If
        Dim dblDiscountTaxamt As Double = 0
        Dim isSkipCogsGL As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, trans)) = 0, False, True)
        'Dim IsAllowPurchaseAccounting As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 0, False, True)
        '' work done by parteek on 11/07/2018
        Dim Against_Subsidy_No As String = ""
        Dim EnableSudsidyCreditNote As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableDistributorSubsidy, clsFixedParameterCode.EnableDistributorSubsidy, trans)) = 0, False, True)
        If EnableSudsidyCreditNote = True Then
            Dim qry1 As String = "select Against_Subsidy_No from TSPL_CUSTOMER_INVOICE_HEAD where document_no='" + obj.Document_No + "'"
            Against_Subsidy_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry1, trans))
        End If
        '' end
        Dim qry As String = "select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No='" + obj.Document_No + "' and Source_Code in ('AR-IN','AR-DN','AR-CR')"
        Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))

        If strVoucherNoForRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreatedOnly) > 0 Then
            strVoucherNo = strVoucherNoForRecreatedOnly
        End If

        ''===========these column is for GL entry=========do for multicurrency==if currency is differ form basecurrency=======15/04/2015==(Monika)================================
        Dim coll As New Hashtable()
        If clsCommon.myLen(obj.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(obj.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ConvRateOld", obj.ConvRate)
        End If
        ''===============================================================================================


        Dim isTaxRecoverable As Boolean = False
        Dim NonRecovTaxAmount As Decimal = 0
        Dim Tax_Liability_Account As String = ""
        Dim strEntryDesc As String = ""
        Dim strSrcType As String = ""
        Dim strSrcDesc As String = ""
        Dim strRemarks As String = ""
        If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) AndAlso (clsCommon.CompairString(FormId, "CSA-SALE") <> CompairStringResult.Equal) Then
            strEntryDesc = "AR Invoice Entry Against-"
            strSrcType = "AR-IN"
            strSrcDesc = "AR Invoice"
        ElseIf (clsCommon.CompairString(FormId, "CSA-SALE") = CompairStringResult.Equal) Then
            strEntryDesc = "AR Invoice Against CSA Sale Patti No." + obj.Against_Sale_No
            strSrcType = "AR-IN"
            strSrcDesc = "AR Invoice"
        ElseIf (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
            strEntryDesc = "AR Debit Note Entry Against-"
            strSrcType = "AR-DN"
            strSrcDesc = "AR Debit"
            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  "
        ElseIf clsCommon.myLen(clsCommon.myCstr(obj.Against_MCC_Material_Sale_Return)) > 0 Then
            strEntryDesc = "AR Credit Note Entry Against-"
            strSrcType = "AR-CR"
            strSrcDesc = "AR Credit Note"
            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Mcc Material Sale Return No -: " & obj.Against_MCC_Material_Sale_Return & " "
        ElseIf clsCommon.myLen(clsCommon.myCstr(Against_Subsidy_No)) > 0 Then
            strEntryDesc = "AR Credit Note Entry Against-"
            strSrcType = "AR-CR"
            strSrcDesc = "AR Credit Note"
            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  Against Transport Subsidy ref doc No " & obj.Document_No & " "
        ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
            strEntryDesc = "AR Credit Note Entry Against-"
            strSrcType = "AR-CR"
            strSrcDesc = "AR Credit Note"
            ''richa ERO/19/03/19-000515 
            If clsCommon.myLen(clsCommon.myCstr(obj.AgainstScrapReturn)) > 0 Then
                strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Sale Return No " & obj.AgainstScrapReturn & " "
            Else
                strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Sale Return No " & obj.Against_Sale_Return_No & " "
            End If


        End If
        Dim ArryLst As ArrayList = New ArrayList()
        '' *********************************************** Conditionally GL Entry(Fresh,Bulk,CSA) ******************************************* 


        '''''''' For Invoice GL entry 
        If clsCommon.myLen(obj.Against_Sale_No) > 0 Then
                If ((clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal)) AndAlso (FormId = "" OrElse FormId = clsUserMgtCode.frmSaleDispatchDairy OrElse FormId = clsUserMgtCode.FrmSalesOrderDispatch OrElse FormId = "FreshSaleInvoice" OrElse FormId = "CSA-SALE" OrElse clsCommon.CompairString(FormId, "INVOICE-PS") = CompairStringResult.Equal) _
                OrElse (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal AndAlso clsCommon.CompairString(FormId, "INVOICE-PS") = CompairStringResult.Equal) Then

                    Dim objInv As clsSNInvoiceHead
                    Dim isTaxExcisable As Boolean = False
                    Dim arr As New List(Of String)
                    Dim dblCogsCost As Double
                    Dim strCogsAcct As String
                    Dim strQuery As String = ""
                    Dim dtAllData As DataTable = Nothing
                    objInv = clsSNInvoiceHead.GetData(obj.Against_Sale_No, NavigatorType.Current, "", trans)
                    Dim strReceivedAgaingstPSDO As String = Nothing
                    Dim objBalAdvTaxAmt As clsSOAdvanceAdjustmentKnockOff = Nothing
                    If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True Then
                        'strReceivedAgaingstPSDO = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Delivery_Code_PS  from TSPL_SD_SALE_INVOICE_HEAD left outer join " & _
                        '"TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code  " & _
                        '"where TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE='" & obj.Against_Sale_No & "' and Delivery_Code_PS in (select Delivery_Code_PS from TSPL_RECEIPT_HEADER)", trans))
                        'strQuery = GetAdvanceTaxAmtForPS(strReceivedAgaingstPSDO, obj.Against_Sale_No)
                        'dtAllData = clsDBFuncationality.GetDataTable(strQuery, trans)



                        objBalAdvTaxAmt = clsSOAdvanceAdjustmentKnockOff.GetBalanceAdvanceAmt(obj.Against_Sale_No, trans)

                    End If

                    ''''' GL entry for Tax and retail Invoicefrm
                    '' Updated by richa agarwal add condition in below line invoice type=S----------------- added A for TAx Exempted type,I= inter state
                    If clsCommon.CompairString(objInv.Invoice_Type, "T") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "E") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "R") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "S") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "N") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "A") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "I") = CompairStringResult.Equal Then

                        ''  for tax gl entry start here
                        Dim objTM As clsTaxMaster
                        Dim dblExcise As Double = 0
                        If obj.TAX1_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX1, trans)
                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX1, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX1, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set Tax Net Payable Account of Tax Authority " + obj.TAX1)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX1_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX1_Amt

                                    '' For excisable FOC entry start here on 26/10/2016 for product sale
                                    If obj.TAX1_ExciseFOCAmt > 0 Then
                                        Acc1 = {objTM.Tax_Net_Payable, obj.TAX1_ExciseFOCAmt}
                                        ArryLst.Add(Acc1)
                                    End If
                                    '' For excisable FOC entry ends here



                                End If
                            End If

                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX1_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX1)
                            End If

                            Dim AccInvDR() As String = {obj.TAX1_GLAC, -1 * obj.TAX1_Amt}
                            ArryLst.Add(AccInvDR)
                            If obj.TAX1_ExciseFOCAmt > 0 Then
                                AccInvDR = {obj.TAX1_GLAC, -1 * obj.TAX1_ExciseFOCAmt}
                                ArryLst.Add(AccInvDR)
                            End If
                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX1_Amt
                            End If
                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX1_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX1, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX1)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX1)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX1_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX1_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX2_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX2, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX2, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX2, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX2)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX2_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX2_Amt

                                    '' For excisable FOC entry start here on 26/10/2016 for product sale
                                    If obj.TAX2_ExciseFOCAmt > 0 Then
                                        Acc1 = {objTM.Tax_Net_Payable, obj.TAX2_ExciseFOCAmt}
                                        ArryLst.Add(Acc1)
                                        dblExcise += obj.TAX2_ExciseFOCAmt
                                    End If
                                    '' For excisable FOC entry ends here

                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX2_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX2)
                            End If

                            Dim AccInvDR() As String = {obj.TAX2_GLAC, -1 * obj.TAX2_Amt}
                            ArryLst.Add(AccInvDR)
                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX2_Amt
                            End If
                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX2_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX2, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX2)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX2)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX2_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX2_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX3_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX3, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX3, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX3, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX3)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX3_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX3_Amt
                                    '' For excisable FOC entry start here on 26/10/2016 for product sale
                                    If obj.TAX3_ExciseFOCAmt > 0 Then
                                        Acc1 = {objTM.Tax_Net_Payable, obj.TAX3_ExciseFOCAmt}
                                        ArryLst.Add(Acc1)
                                        dblExcise += obj.TAX3_ExciseFOCAmt
                                    End If
                                    '' For excisable FOC entry ends here

                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX3_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX3)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX3_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX3_GLAC, -1 * obj.TAX3_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX3_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX3, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX3)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX3)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX3_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX3_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX4_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX4, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX4, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX4, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX4)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX4_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX4_Amt
                                    '' For excisable FOC entry start here on 26/10/2016 for product sale
                                    If obj.TAX4_ExciseFOCAmt > 0 Then
                                        Acc1 = {objTM.Tax_Net_Payable, obj.TAX4_ExciseFOCAmt}
                                        ArryLst.Add(Acc1)
                                        dblExcise += obj.TAX4_ExciseFOCAmt
                                    End If
                                    '' For excisable FOC entry ends here

                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX4_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX4)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX4_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX4_GLAC, -1 * obj.TAX4_Amt}
                            ArryLst.Add(AccInvDR)
                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX4_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX4, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX4)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX4)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX4_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX4_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If
                        If obj.TAX5_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX5, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX5, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX5, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX5)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX5_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX5_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX5_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX5)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX5_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX5_GLAC, -1 * obj.TAX5_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX5_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX5, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX5)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX5)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX5_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX5_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX6_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX6, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX6, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX6, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX6)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX6_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX6_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX6_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX6)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX6_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX6_GLAC, -1 * obj.TAX6_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX6_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX6, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX6)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX6)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX6_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX6_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX7_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX7, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX7, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX7, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX7)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX7_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX7_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX7_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX7)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX7_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX7_GLAC, -1 * obj.TAX7_Amt}
                            ArryLst.Add(AccInvDR)
                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX7_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX7, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX7)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX7)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX7_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX7_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX8_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX8, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX8, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX8, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX8)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX8_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX8_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX8_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX8)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX8_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX8_GLAC, -1 * obj.TAX8_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX8_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX8, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX8)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX8)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX8_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX8_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If


                        If obj.TAX9_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX9, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX9, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX9, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX9)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX9_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX9_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX9_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX9)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX9_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX9_GLAC, -1 * obj.TAX9_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX9_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX9, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX9)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX9)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX9_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX9_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If

                        If obj.TAX10_Amt <> 0 Then
                            isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX10, trans)

                            ' for excisable tax start here
                            isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX10, trans)
                            If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX10, trans)
                                If objTM IsNot Nothing Then
                                    If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                        Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX10)
                                    End If
                                    objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX10_Amt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX10_Amt
                                End If
                            End If
                            'Excisable tax ends here

                            If clsCommon.myLen(obj.TAX10_GLAC) <= 0 Then
                                Throw New Exception("GL Acount not found for" + obj.TAX10)
                            End If

                            If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                                dblExcise += obj.TAX10_Amt
                            End If
                            Dim AccInvDR() As String = {obj.TAX10_GLAC, -1 * obj.TAX10_Amt}
                            ArryLst.Add(AccInvDR)

                            If clsCommon.CompairString(obj.Trans_Type, "PS") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True AndAlso objBalAdvTaxAmt IsNot Nothing Then
                                If objBalAdvTaxAmt.TAX10_Amt <> 0 Then
                                    objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX10, trans)
                                    If objTM IsNot Nothing Then
                                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                                            Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + obj.TAX10)
                                        End If
                                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                                        If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                                            Throw New Exception("Please set Tax Liablity Account of Tax Authority " + obj.TAX10)
                                        End If
                                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, obj.loc_code, True, trans)
                                        Dim Acc2() As String = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX10_Amt}
                                        ArryLst.Add(Acc2)
                                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objBalAdvTaxAmt.TAX10_Amt}
                                        ArryLst.Add(Acc3)
                                    End If
                                End If

                            End If
                        End If





                        ''  tax gl entry ends here
                        '' FOR Additional Cost START here
                        If obj.Add_Charge_Amt1 <> 0 Then
                            Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                            AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                            End If

                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt2 <> 0 Then
                            Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                            AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                            End If

                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt3 <> 0 Then
                            Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                            AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                            End If

                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt4 <> 0 Then
                            Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                            AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                            End If

                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt5 <> 0 Then
                            Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                            If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                            End If
                            AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)

                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt6 <> 0 Then
                            Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                            AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                            End If
                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt7 <> 0 Then
                            Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                            AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                            End If
                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt8 <> 0 Then
                            Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                            AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                            End If
                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt9 <> 0 Then
                            Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                            AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                            End If
                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                            ArryLst.Add(AccAddCostCR)
                        End If
                        If obj.Add_Charge_Amt10 <> 0 Then
                            Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                            AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)

                            If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                                Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                            End If
                            Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                            ArryLst.Add(AccAddCostCR)
                        End If

                        '' Additional cost ends here

                        ''richa agarwal added on 02-jan-2015
                        If clsCommon.CompairString(objInv.Invoice_Type, "S") <> CompairStringResult.Equal Then
                            Dim isFirstTime As Boolean = True
                            For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                                'Dim dblLedgeerNonRecoverableAmt As Double = clsCustomerInvoiceHead.GetTaxAmt(objTR, trans)
                                'Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount_less_Discount)}
                                Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount_less_Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                ArryLst.Add(AccInvDR)


                                If FormId = clsUserMgtCode.frmSaleDispatchDairy Then
                                    'Dim AccInvTaxDR() As String = {objTR.GL_Account_Code, -1 * (objTR.TAX1_Amt + objTR.TAX2_Amt + objTR.TAX3_Amt + objTR.TAX4_Amt + objTR.TAX5_Amt + objTR.TAX6_Amt + objTR.TAX7_Amt)}
                                    'ArryLst.Add(AccInvTaxDR)
                                End If

                                If isFirstTime AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                                    Dim AccExciseDR() As String = {objTR.GL_Account_Code, -1 * dblExcise, "", "", "", "", "", "", objTR.Reco_Control_Account}
                                    ArryLst.Add(AccExciseDR)

                                End If
                                isFirstTime = False
                                ''''''added by priti for discount entry of invoice

                                If objTR.Amount_less_Discount = 0 AndAlso FormId = "FreshSaleInvoice" Then
                                    Dim AccDiscDR() As String = {objTR.GL_Account_Code, 1 * (objTR.Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                    ArryLst.Add(AccDiscDR)
                                End If
                                If FormId = clsUserMgtCode.frmSaleDispatchDairy AndAlso objTR.Discount > 0 AndAlso objTR.Amount_less_Discount = 0 Then
                                    Dim AccDDiscDR() As String = {objTR.GL_Account_Code, 1 * (objTR.Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                    ArryLst.Add(AccDDiscDR)

                                    Dim AccDiscTaxDR() As String = {objTR.GL_Account_Code, 1 * (objTR.TAX1_Amt + objTR.TAX2_Amt + objTR.TAX3_Amt + objTR.TAX4_Amt + objTR.TAX5_Amt + objTR.TAX6_Amt + objTR.TAX7_Amt), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                    ArryLst.Add(AccDiscTaxDR)
                                    dblDiscountTaxamt += objTR.TAX1_Amt + objTR.TAX2_Amt + objTR.TAX3_Amt + objTR.TAX4_Amt + objTR.TAX5_Amt + objTR.TAX6_Amt + objTR.TAX7_Amt
                                End If
                                If objTR.Distributor_Commission_Amt > 0 Then
                                    Dim AccDDiscDR() As String = {objTR.Promotional_GL_Account_Code, 1 * (objTR.Distributor_Commission_Amt)}
                                    ArryLst.Add(AccDDiscDR)

                                    Dim AccDiscTaxDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Distributor_Commission_Amt)}
                                    ArryLst.Add(AccDiscTaxDR)
                                End If
                                If objTR.Transporter_Commission_Amt > 0 Then
                                    Dim strLocationt As String = ""
                                    Dim strACWithLocationt As String = ""
                                    Dim AccDiscDR() As String = Nothing
                                    Dim AccDiscTaxDR() As String = Nothing
                                    If clsCommon.CompairString(obj.Trans_Type, "MCC") = CompairStringResult.Equal Then

                                        If obj.Is_Add_TPT Then
                                            AccDiscTaxDR = {objTR.Transporter_GL_Account_Code, -1 * (objTR.Transporter_Commission_Amt)}
                                        Else
                                            AccDiscTaxDR = {objTR.Transporter_GL_Account_Code, (objTR.Transporter_Commission_Amt)}
                                        End If
                                        ArryLst.Add(AccDiscTaxDR)
                                    Else
                                        strLocationt = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                                        strACWithLocationt = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocationt, True, trans)
                                        AccDiscDR = {strACWithLocationt, 1 * (objTR.Transporter_Commission_Amt)}
                                        ArryLst.Add(AccDiscDR)
                                        AccDiscTaxDR = {objTR.Transporter_GL_Account_Code, -1 * (objTR.Transporter_Commission_Amt)}
                                        ArryLst.Add(AccDiscTaxDR)
                                    End If

                                End If

                                If objTR.Security_Amt > 0 Then
                                    Dim strLocations As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                                    Dim strACWithLocations As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocations, True, trans)

                                    Dim AccDiscDR() As String = {strACWithLocations, 1 * (objTR.Security_Amt)}
                                    ArryLst.Add(AccDiscDR)

                                    Dim AccDiscTaxDR() As String = {objTR.SD_GL_Account_Code, -1 * (objTR.Security_Amt)}
                                    ArryLst.Add(AccDiscTaxDR)
                                End If


                                ''''''code ends here    
                            Next
                            Dim AllowCrateCanPhysicalStock As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, trans))
                            If AllowCrateCanPhysicalStock = 1 Then
                                ' DOne by priti BHA/15/06/18-000055
                                If FormId = clsUserMgtCode.frmSaleDispatchDairy Then

                                    ' FOr Crate
                                    Dim strCrateItem = ""
                                    Dim strCrateUOM = ""
                                    Dim dblCrateRate As Integer = 0
                                    Dim dblCrateQty As Integer = 0
                                    Dim strReturnable_ContainerAC As String = ""
                                    Dim strContainerDepositAC As String = ""
                                    Dim Acc() As String = Nothing
                                    Dim Acc1() As String = Nothing
                                    qry = "select Crate_Item,Crate_ItemUnit,Crate_ItemRate,Crate from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & objInv.Document_Code & "'"
                                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                                    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                        strCrateItem = clsCommon.myCstr(dt.Rows(0)("Crate_Item"))
                                        strCrateUOM = clsCommon.myCstr(dt.Rows(0)("Crate_ItemUnit"))
                                        dblCrateRate = clsCommon.myCdbl(dt.Rows(0)("Crate_ItemRate"))
                                        dblCrateQty = clsCommon.myCdbl(dt.Rows(0)("Crate"))
                                    End If
                                    If dblCrateQty > 0 Then
                                        strReturnable_ContainerAC = clsCommon.myCstr(clsItemMaster.GetReturnableConGLAC(strCrateItem, trans))
                                        If clsCommon.myLen(strReturnable_ContainerAC) = 0 Then
                                            Throw New Exception("Please set Returnable Container Account for item - " + strCrateItem)
                                        End If
                                        strReturnable_ContainerAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strReturnable_ContainerAC, obj.loc_code, True, trans)
                                        Acc = {strReturnable_ContainerAC, -1 * (dblCrateQty * dblCrateRate)}
                                        ArryLst.Add(Acc)
                                        strContainerDepositAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Container_Deposit ,'') as Container_Deposit from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & objInv.Customer_Code & "'", trans))
                                        If clsCommon.myLen(strContainerDepositAC) = 0 Then
                                            Throw New Exception("Please set Container Deposit Account for customer - " + objInv.Customer_Code)
                                        End If
                                        strContainerDepositAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strContainerDepositAC, obj.loc_code, True, trans)
                                        Acc1 = {strContainerDepositAC, (dblCrateQty * dblCrateRate)}
                                        ArryLst.Add(Acc1)
                                    End If
                                    ' FOr Can
                                    Dim strCanItem = ""
                                    Dim strCanUOM = ""
                                    Dim dblCanRate As Integer = 0
                                    Dim dblCanQty As Integer = 0
                                    qry = "select Can_Item,Can_ItemUnit,Can_ItemRate,ShippedCAN from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & objInv.Document_Code & "'"
                                    dt = clsDBFuncationality.GetDataTable(qry, trans)
                                    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                        strCanItem = clsCommon.myCstr(dt.Rows(0)("Can_Item"))
                                        strCanUOM = clsCommon.myCstr(dt.Rows(0)("Can_ItemUnit"))
                                        dblCanRate = clsCommon.myCdbl(dt.Rows(0)("Can_ItemRate"))
                                        dblCanQty = clsCommon.myCdbl(dt.Rows(0)("ShippedCAN"))
                                    End If
                                    If dblCanQty > 0 Then
                                        strReturnable_ContainerAC = clsCommon.myCstr(clsItemMaster.GetReturnableConGLAC(strCanItem, trans))
                                        If clsCommon.myLen(strReturnable_ContainerAC) = 0 Then
                                            Throw New Exception("Please set Returnable Container Account for item " + strCanItem)
                                        End If
                                        strReturnable_ContainerAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strReturnable_ContainerAC, obj.loc_code, True, trans)
                                        Acc = Nothing
                                        Acc = {strReturnable_ContainerAC, -1 * (dblCanQty * dblCanRate)}
                                        ArryLst.Add(Acc)
                                        '====preeti gupta=============
                                        strContainerDepositAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Container_Deposit ,'') as Container_Deposit from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & objInv.Customer_Code & "'", trans))
                                        If clsCommon.myLen(strContainerDepositAC) = 0 Then
                                            Throw New Exception("Please set Container Deposit Account for customer - " + objInv.Customer_Code)
                                        End If
                                        strContainerDepositAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strContainerDepositAC, obj.loc_code, True, trans)
                                        '-----------------------------
                                        Acc1 = Nothing
                                        Acc1 = {strContainerDepositAC, (dblCanQty * dblCanRate)}
                                        ArryLst.Add(Acc1)
                                    End If
                                End If
                            End If
                            Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                            Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)

                            ' done by priti SWA/20/08/18-000045 for swadesh to add leakage account in fresh journal entry
                            Dim dblLeakAmount As Double = 0
                            Dim LeakageAcct As String = ""
                            If FormId = "FreshSaleInvoice" Then
                                If obj.LeakageAmount > 0 Then
                                    dblLeakAmount = obj.LeakageAmount
                                    LeakageAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Leakage_Deduction from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + obj.Account_Set + "'", trans))
                                    LeakageAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(LeakageAcct, strLocation, True, trans)
                                    If clsCommon.myLen(LeakageAcct) <= 0 Then
                                        Throw New Exception("Please set Leakage account set of customer account set :" + obj.Account_Set)
                                    End If
                                End If

                                Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                                ArryLst.Add(AccInvCR)

                                If dblLeakAmount > 0 Then
                                    AccInvCR = Nothing
                                    AccInvCR = {LeakageAcct, dblLeakAmount}
                                    ArryLst.Add(AccInvCR)
                                End If
                            Else
                                'If clsCommon.CompairString(obj.Trans_Type, "MCC") = CompairStringResult.Equal AndAlso obj.TotalSubsidyAmt > 0 Then
                                If obj.TotalSubsidyAmt > 0 Then
                                    Dim strLocationt As String = ""
                                    Dim strRate_Difference_act As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Rate_Difference ,'') as Rate_Difference from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                                    If clsCommon.myLen(strRate_Difference_act) = 0 Then
                                        Throw New Exception("Please set Rate Difference Account for customer - " + obj.Customer_Code)
                                    End If
                                    strRate_Difference_act = clsERPFuncationality.ChangeGLAccountLocationSegment(strRate_Difference_act, obj.loc_code, True, trans)
                                    Dim Acc1() As String = {strRate_Difference_act, (obj.TotalSubsidyAmt)}
                                    ArryLst.Add(Acc1)
                                End If

                                Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                                ArryLst.Add(AccInvCR)

                                If dblDiscountTaxamt > 0 Then
                                    AccInvCR = {strACWithLocation, -dblDiscountTaxamt}
                                    ArryLst.Add(AccInvCR)
                                End If
                            End If







                            ' for  rounding off account
                            If obj.RoundOffAmount <> 0 Then
                                Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))

                                If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                                    Throw New Exception("Please set round off account in Sales Setting")
                                End If

                                '================Changed By Rohit on Apr 3,2015 .showing Error on Post Dispatch .Because it was searching AccountSegmentof Location not Segment.==============
                                ' strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, objInv.Bill_To_Location, True, trans)

                                strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, objInv.Bill_To_Location, False, trans)
                                '==========================================================================================================================================================
                                Dim AccRoundInvCR() As String = {strACRoundInvCr, -1 * obj.RoundOffAmount}
                                ArryLst.Add(AccRoundInvCR)

                            End If

                            If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL'And Not IsAllowPurchaseAccounting
                                Dim SentschemecogsinAnotherAccount As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SentschemecogsinAnotherAccount, clsFixedParameterCode.SentschemecogsinAnotherAccount, trans)) = "1", True, False))
                                For Each objInvDetail As clsSNInvoiceDetail In objInv.Arr
                                    Dim strCode As String = objInvDetail.Shipment_Code
                                    ''richa agarwal need to change control account for scheme item as well as for non scheme item  20 Dec,2018 add ((FormId = clsUserMgtCode.frmSaleDispatchDairy) AndAlso SentschemecogsinAnotherAccount = True) in below condition ERO/20/12/18-000448
                                    If (FormId <> clsUserMgtCode.frmSaleDispatchDairy) OrElse FormId = clsUserMgtCode.FrmSalesOrderDispatch OrElse ((FormId = clsUserMgtCode.frmSaleDispatchDairy) AndAlso SentschemecogsinAnotherAccount = True) Then
                                        If Not arr.Contains(strCode) Then

                                            arr.Add(strCode)
                                            dblCogsCost += clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end) as COst from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & objInvDetail.Shipment_Code & "'", trans))

                                            ''richa agarwal need to change control account for scheme item as well as for non scheme item  20 Dec,2018 ERO/20/12/18-000448
                                            ''''' for cogs entry item wise
                                            Dim strSql As String = "select isnull(TSPL_INVENTORY_MOVEMENT.Is_Scheme_Item,'N') as Is_Scheme_Item,TSPL_INVENTORY_MOVEMENT.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code " &
                                        "where Source_Doc_No='" & objInvDetail.Shipment_Code & "'"
                                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                                            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                                For Each dr As DataRow In dt.Rows
                                                    ''richa agarwal need to change control account for scheme item as well as for non scheme item 20 Dec,2018 ERO/20/12/18-000448
                                                    If SentschemecogsinAnotherAccount = True Then
                                                        ''richa agarwal 13 Sep,2019 pick cost of goods scheme account when is sampling check box is on too  ERO/11/09/19-001027
                                                        If clsCommon.CompairString(dr("Is_Scheme_Item"), "Y") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(issampling,0) from TSPL_SD_SALE_INVOICE_HEAD where Document_Code ='" & objInv.Document_Code & "'", trans)), "1") = CompairStringResult.Equal Then
                                                            strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Scheme from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                                                            If clsCommon.myLen(strCogsAcct) = 0 Then
                                                                Throw New Exception("Please set Cost Of Goods Scheme Account for first item")
                                                            End If
                                                        Else
                                                            strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                                                            If clsCommon.myLen(strCogsAcct) = 0 Then
                                                                Throw New Exception("Please set Cost Of Goods Account for first item")
                                                            End If
                                                        End If
                                                    Else
                                                        strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                                                        If clsCommon.myLen(strCogsAcct) = 0 Then
                                                            Throw New Exception("Please set Cost Of Goods Account for first item")
                                                        End If
                                                    End If
                                                    ''-----------------------
                                                    '=================rohit Done on Nov 11,2014 =====Discussed with Priti Mam and Balwinder Sir==============

                                                    strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                                                    Dim Acc1() As String = {strCogsAcct, clsCommon.myCdbl(dr("Cost"))}
                                                    ArryLst.Add(Acc1)
                                                Next
                                            End If
                                            ''''' cogs entry item wise ends here
                                        End If
                                    Else


                                        ''''' for cogs entry item wise for Dairy Sale
                                        qry = "select case when sum(QTy) > 0 then  sum (case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end)/sum(Qty) else 0 end as Cost from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code " &
                                    "where Source_Doc_No='" & objInvDetail.Shipment_Code & "'  and " &
                                    "TSPL_INVENTORY_MOVEMENT.Item_Code='" & objInvDetail.Item_Code & "' and TSPL_INVENTORY_MOVEMENT.UOM='" & objInvDetail.Unit_code & "'"
                                        Dim dblUnitCost As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                                        If dblUnitCost > 0 Then
                                            strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + objInvDetail.Item_Code + "'", trans))
                                            If clsCommon.myLen(strCogsAcct) = 0 Then
                                                Throw New Exception("Please set Cost Of Goods Account for first item")
                                            End If
                                            '=================rohit Done on Nov 11,2014 =====Discussed with Priti Mam and Balwinder Sir==============
                                            strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                                            Dim dblCost As Double = Math.Round(dblUnitCost * objInvDetail.Qty, 2)
                                            Dim Acc1() As String = {strCogsAcct, dblCost}
                                            ArryLst.Add(Acc1)
                                            dblCogsCost += dblCost
                                        End If
                                        ''''' cogs entry item wise ends here
                                    End If



                                Next


                                Dim strShipmentClearingAC = clsDBFuncationality.getSingleValue("SELECT PA.Shipment_Clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                      " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                       " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.Arr.Item(0).Item_Code.ToString() + "'", trans)
                                If clsCommon.myLen(strShipmentClearingAC) = 0 Then
                                    Throw New Exception("Please set Shipment clearing Account for first item")
                                End If
                                strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, objInv.Bill_To_Location, trans)

                                'Dim strCogsAcct = clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + objInv.Arr.Item(0).Item_Code.ToString() + "'", trans)
                                'If clsCommon.myLen(strCogsAcct) = 0 Then
                                '    Throw New Exception("Please set Cost of Goods Sold Account for first item")
                                'End If
                                'strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, objInv.Bill_To_Location, trans)

                                Dim Acc() As String = {strShipmentClearingAC, -1 * dblCogsCost, "", "", "", "", "", "", "H"}


                                ArryLst.Add(Acc)

                            End If  '' Done By Pankaj Jha For Skipping Cogs GL


                            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Sale Invoice No " & objInv.Document_Code & " "
                        Else
                            ''richa 
                            ''''' GL entry for Service Invoice start here
                            For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                                Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount_less_Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                ArryLst.Add(AccInvDR)
                            Next
                            Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                            Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)

                            Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                            ArryLst.Add(AccInvCR)

                            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Service Invoice No " & objInv.Document_Code & " "

                        End If

                    Else
                        ''''' GL entry for Service Invoice start here
                        For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                            Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount_less_Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                            ArryLst.Add(AccInvDR)
                        Next
                        Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                        Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)

                        Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                        ArryLst.Add(AccInvCR)

                        strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Service Invoice No " & objInv.Document_Code & " "
                        If clsCommon.CompairString(FormId, "CSA-SALE") = CompairStringResult.Equal Then
                            strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For CSA Sale Patti No " & objInv.Document_Code & " "
                        End If
                    End If
                    ''If Credit Note then reverse amount 
                    If clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal Then
                        Dim ArryLstNew As ArrayList = New ArrayList()
                        For Each Str() As String In ArryLst
                            Dim strNew() As String = {Str(0), -1 * Str(1), If(Str.Length >= 3, Str(2), ""), If(Str.Length >= 4, Str(3), ""), If(Str.Length >= 5, Str(4), ""), If(Str.Length >= 6, Str(5), "")}
                            ArryLstNew.Add(strNew)
                        Next

                        ArryLst = New ArrayList
                        For Each Str() As String In ArryLstNew
                            Dim strNew() As String = {Str(0), Str(1), If(Str.Length >= 3, Str(2), ""), If(Str.Length >= 4, Str(3), ""), If(Str.Length >= 5, Str(4), ""), If(Str.Length >= 6, Str(5), "")}
                            ArryLst.Add(strNew)
                        Next

                    End If


                    '' cOGS ENDS HERE


                ElseIf ((clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal)) AndAlso FormId = "BulkSaleInvoice" Then
                    Dim objInv As ClsInvoiceBulkSale
                    Dim arr As New List(Of String)
                    Dim dblCogsCost As Decimal
                    Dim strCogsAcct As String
                    objInv = ClsInvoiceBulkSale.GetData(obj.Against_Sale_No, "", NavigatorType.Current, trans)
                    ''''' GL entry for Tax and retail Invoice
                    'If clsCommon.CompairString(objInv.Invoice_Type, "T") = CompairStringResult.Equal OrElse clsCommon.CompairString(objInv.Invoice_Type, "R") = CompairStringResult.Equal Then


                    ''  for tax gl entry start here
                    Dim objTM As clsTaxMaster
                    Dim dblExcise As Double = 0
                    Dim isTaxExcisable As Boolean = False

                    If obj.TAX1_Amt <> 0 Then
                        isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX1, trans)
                        ' for excisable tax start here
                        isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX1, trans)
                        If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                            objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX1, trans)
                            If objTM IsNot Nothing Then
                                If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                    Throw New Exception("Please set Tax Net Payable Account of Tax Authority " + obj.TAX1)
                                End If
                                objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX1_Amt}
                                ArryLst.Add(Acc1)
                                dblExcise += obj.TAX1_Amt

                                '' For excisable FOC entry start here on 26/10/2016 for product sale
                                If obj.TAX1_ExciseFOCAmt > 0 Then
                                    Acc1 = {objTM.Tax_Net_Payable, obj.TAX1_ExciseFOCAmt}
                                    ArryLst.Add(Acc1)
                                End If
                                '' For excisable FOC entry ends here
                            End If
                        End If

                        'Excisable tax ends here

                        If clsCommon.myLen(obj.TAX1_GLAC) <= 0 Then
                            Throw New Exception("GL Acount not found for" + obj.TAX1)
                        End If

                        Dim AccInvDR() As String = {obj.TAX1_GLAC, -1 * obj.TAX1_Amt}
                        ArryLst.Add(AccInvDR)
                        If obj.TAX1_ExciseFOCAmt > 0 Then
                            AccInvDR = {obj.TAX1_GLAC, -1 * obj.TAX1_ExciseFOCAmt}
                            ArryLst.Add(AccInvDR)
                        End If
                        If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                            dblExcise += obj.TAX1_Amt
                        End If

                    End If

                    If obj.TAX2_Amt <> 0 Then
                        isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX2, trans)

                        ' for excisable tax start here
                        isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX2, trans)
                        If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                            objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX2, trans)
                            If objTM IsNot Nothing Then
                                If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                    Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX2)
                                End If
                                objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX2_Amt}
                                ArryLst.Add(Acc1)
                                dblExcise += obj.TAX2_Amt

                                '' For excisable FOC entry start here on 26/10/2016 for product sale
                                If obj.TAX2_ExciseFOCAmt > 0 Then
                                    Acc1 = {objTM.Tax_Net_Payable, obj.TAX2_ExciseFOCAmt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX2_ExciseFOCAmt
                                End If
                                '' For excisable FOC entry ends here

                            End If
                        End If
                        'Excisable tax ends here

                        If clsCommon.myLen(obj.TAX2_GLAC) <= 0 Then
                            Throw New Exception("GL Acount not found for" + obj.TAX2)
                        End If

                        Dim AccInvDR() As String = {obj.TAX2_GLAC, -1 * obj.TAX2_Amt}
                        ArryLst.Add(AccInvDR)
                        If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                            dblExcise += obj.TAX2_Amt
                        End If

                    End If

                    If obj.TAX3_Amt <> 0 Then
                        isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX3, trans)

                        ' for excisable tax start here
                        isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX3, trans)
                        If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                            objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX3, trans)
                            If objTM IsNot Nothing Then
                                If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                    Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX3)
                                End If
                                objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                                Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX3_Amt}
                                ArryLst.Add(Acc1)
                                dblExcise += obj.TAX3_Amt
                                '' For excisable FOC entry start here on 26/10/2016 for product sale
                                If obj.TAX3_ExciseFOCAmt > 0 Then
                                    Acc1 = {objTM.Tax_Net_Payable, obj.TAX3_ExciseFOCAmt}
                                    ArryLst.Add(Acc1)
                                    dblExcise += obj.TAX3_ExciseFOCAmt
                                End If
                                '' For excisable FOC entry ends here

                            End If
                        End If
                        'Excisable tax ends here

                        If clsCommon.myLen(obj.TAX3_GLAC) <= 0 Then
                            Throw New Exception("GL Acount not found for" + obj.TAX3)
                        End If

                        If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                            dblExcise += obj.TAX3_Amt
                        End If
                        Dim AccInvDR() As String = {obj.TAX3_GLAC, -1 * obj.TAX3_Amt}
                        ArryLst.Add(AccInvDR)

                    End If
                    ''tax ends here

                    '''' FOR Additional Cost START here
                    If obj.Add_Charge_Amt1 <> 0 Then
                        Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                        AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                        End If

                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt2 <> 0 Then
                        Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                        AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                        End If

                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt3 <> 0 Then
                        Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                        AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                        End If

                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt4 <> 0 Then
                        Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                        AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                        End If

                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt5 <> 0 Then
                        Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                        If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                        End If
                        AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)

                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt6 <> 0 Then
                        Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                        AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                        End If
                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt7 <> 0 Then
                        Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                        AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                        End If
                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt8 <> 0 Then
                        Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                        AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                        End If
                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt9 <> 0 Then
                        Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                        AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                        End If
                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    If obj.Add_Charge_Amt10 <> 0 Then
                        Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                        AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)

                        If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                            Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                        End If
                        Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                        ArryLst.Add(AccAddCostCR)
                    End If
                    '' Additional cost ends here

                    Dim isFirstTime As Boolean = True
                    For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                        'Dim dblLedgeerNonRecoverableAmt As Double = clsCustomerInvoiceHead.GetTaxAmt(objTR, trans)
                        Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount), "", "", "", "", "", "", objTR.Reco_Control_Account}

                        ArryLst.Add(AccInvDR)
                        isFirstTime = False
                    Next
                    Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                    Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)

                    ''richa agarwal 14/10/2014
                    'Dim creditamount As Double = 0
                    'creditamount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select SUM(InvoiceAmount) from TSPL_INVOICE_DETAIL_BULKSALE  where Document_No='" & objInv.Document_No & "'", trans))
                    ' Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                    'Dim AccInvCR() As String = {strACWithLocation, obj.Discount_Base}
                    'Dim AccInvCR() As String = {strACWithLocation, obj.Discount_Base + obj.RoundOffAmount}
                    Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                    ArryLst.Add(AccInvCR)


                    If obj.RoundOffAmount <> 0 Then
                        Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))

                        If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                            Throw New Exception("Please set round off account in Sales Setting")
                        End If

                        strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, strLocation, True, trans)
                        Dim AccRoundInvCR() As String = {strACRoundInvCr, -1 * obj.RoundOffAmount}
                        ArryLst.Add(AccRoundInvCR)
                    End If

                    ''============

                    If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL'And Not IsAllowPurchaseAccounting
                        Dim Costincaseoflossandgain As Decimal = 0
                        For Each objInvDetail As ClsInvoiceDetailBulkSale In objInv.arrInvoiceDetailBulkSale
                            Dim strCode As String = objInvDetail.Dispatch_Code
                            If Not arr.Contains(strCode) Then
                                arr.Add(strCode)
                                '' changes by richa agarwal against ticket BM00000006070
                                ''updation by richa agarwal according to gain or loss amount
                                Dim dblCogsCosttemp As Decimal = 0
                                dblCogsCosttemp = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end) as COst from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & objInvDetail.Dispatch_Code & "'", trans))
                                ' dblCogsCost += clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end)/Qty) * " & objInvDetail.InvoiceQty & " as COst from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & objInvDetail.Dispatch_Code & "'", trans))
                                dblCogsCost += dblCogsCosttemp
                                ''''' for cogs entry item wise
                                Dim strSql As String = String.Empty
                                If clsCommon.CompairString(objInv.InvoiceAgainst, "Against Dispatch") = CompairStringResult.Equal Then
                                    ''richa agarwal 10 Jan,2019
                                    Dim dblInvoiceQty As Double = 0
                                    Dim UseKGLitreConversionInBulkSaleAsperCLRCalculation As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.UseKGLitreConversionInBulkSaleAsperCLRCalculation, clsFixedParameterCode.UseKGLitreConversionInBulkSaleAsperCLRCalculation, trans)) = 1, True, False))
                                    'If UseKGLitreConversionInBulkSaleAsperCLRCalculation = True Then
                                    '    dblInvoiceQty = Math.Round(clsCommon.myCdbl(clsItemMaster.GetQtyInLtrFromKgByCLR(clsCommon.myCdbl(objInvDetail.InvoiceQty), clsCommon.myCdbl(objInvDetail.CLR))), 2)

                                    'Else
                                    '    dblInvoiceQty = objInvDetail.InvoiceQty
                                    'End If
                                    ''richa ERO/25/02/19-000499
                                    dblInvoiceQty = IIf(UseKGLitreConversionInBulkSaleAsperCLRCalculation = True, objInvDetail.InvoiceQty_in_Ltr, objInvDetail.InvoiceQty)
                                    ''-----------------------------------end
                                    strSql = "select TSPL_INVENTORY_MOVEMENT_NEW.Item_Code,((case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end)/TSPL_INVENTORY_MOVEMENT_NEW.Qty)* " & dblInvoiceQty & "  as Cost from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code" &
                               " where Source_Doc_No='" & objInvDetail.Dispatch_Code & "'"
                                Else
                                    strSql = "select TSPL_INVENTORY_MOVEMENT_NEW.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code " &
                                "where Source_Doc_No='" & objInvDetail.Dispatch_Code & "'"
                                End If
                                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                                If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                    For Each dr As DataRow In dt.Rows
                                        strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                                        If clsCommon.myLen(strCogsAcct) = 0 Then
                                            Throw New Exception("Please set Cost of Goods Sold for first item")
                                        End If
                                        ''richa agarwal discussed with Balwinder sir
                                        'strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, trans)
                                        strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                                        ''----------------------------------
                                        Dim Acc1() As String = {strCogsAcct, clsCommon.myCdbl(dr("Cost"))}

                                        ArryLst.Add(Acc1)

                                        Costincaseoflossandgain = clsCommon.myCdbl(dr("Cost"))
                                    Next
                                End If
                                ''richa agarwal 06/04/2015
                                If clsCommon.CompairString(objInv.InvoiceAgainst, "Against Dispatch") = CompairStringResult.Equal Then
                                    ' If dblCogsCost <> Costincaseoflossandgain Then ''richa dblCogsCosttemp <> Costincaseoflossandgain  as per ranjana mam and balwinder sir ERO/10/01/19-000462
                                    If Math.Abs(dblCogsCosttemp - Costincaseoflossandgain) > 0.001 Then
                                        Dim strGainorLossAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT SA.Gain_Loss_Account  FROM TSPL_ITEM_MASTER AS IM INNER JOIN  TSPL_SALES_ACCOUNTS  AS SA ON IM.Sale_Class_Code  = SA.Sales_Class_Code  INNER JOIN  TSPL_GL_ACCOUNTS AS GLA ON SA.Gain_Loss_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.arrInvoiceDetailBulkSale.Item(0).Item_Code.ToString() + "'", trans))
                                        If clsCommon.myLen(strGainorLossAC) = 0 Then
                                            Throw New Exception("Please set Gain/Loss Account for first item")
                                        End If
                                        strGainorLossAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strGainorLossAC, obj.loc_code, True, trans)

                                        'Dim Acc2() As String = {strGainorLossAC, 1 * (dblCogsCost - Costincaseoflossandgain)}
                                        Dim Acc2() As String = {strGainorLossAC, 1 * (dblCogsCosttemp - Costincaseoflossandgain)}
                                        ArryLst.Add(Acc2)

                                    End If
                                End If
                                ''------------------
                                ''''' cogs entry item wise ends here
                            End If
                        Next

                        ' ''richa agarwal 23/02/2015
                        'If clsCommon.CompairString(objInv.InvoiceAgainst, "Against Dispatch") = CompairStringResult.Equal Then
                        '    If dblCogsCost <> Costincaseoflossandgain Then
                        '        Dim strGainorLossAC = clsDBFuncationality.getSingleValue("SELECT SA.Gain_Loss_Account  FROM TSPL_ITEM_MASTER AS IM INNER JOIN  TSPL_SALES_ACCOUNTS  AS SA ON IM.Sale_Class_Code  = SA.Sales_Class_Code  INNER JOIN  TSPL_GL_ACCOUNTS AS GLA ON SA.Gain_Loss_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.arrInvoiceDetailBulkSale.Item(0).Item_Code.ToString() + "'", trans)
                        '        If clsCommon.myLen(strGainorLossAC) = 0 Then
                        '            Throw New Exception("Please set Gain/Loss Account for first item")
                        '        End If
                        '        strGainorLossAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strGainorLossAC, obj.loc_code, True, trans)

                        '        Dim Acc2() As String = {strGainorLossAC, 1 * (dblCogsCost - Costincaseoflossandgain)}
                        '        ArryLst.Add(Acc2)

                        '    End If
                        'End If
                        ' ''------------------


                        Dim strShipmentClearingAC = clsDBFuncationality.getSingleValue("SELECT PA.Shipment_Clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                     " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.arrInvoiceDetailBulkSale.Item(0).Item_Code.ToString() + "'", trans)
                        If clsCommon.myLen(strShipmentClearingAC) = 0 Then
                            Throw New Exception("Please set Shipment clearing Account for first item")
                        End If
                        ''richa 13/09/2014 change 
                        '  strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, objInv.Location_Code, trans)
                        strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, obj.loc_code, True, trans)

                        Dim Acc() As String = {strShipmentClearingAC, -1 * dblCogsCost, "", "", "", "", "", "", "H"}
                        ArryLst.Add(Acc)



                    End If  '' Done By Pankaj Jha For Skipping Cogs GL

                    ''richa agarwal
                    If clsCommon.CompairString(objInv.InvoiceAgainst, "Against Dispatch") = CompairStringResult.Equal Then
                        strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Invoice Bulk Sale No " & objInv.Document_No & " "
                    Else
                        strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Invoice Bulk Sale Trade No " & objInv.Document_No & " "
                    End If

                    ''=====================
                    'strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Sale Invoice No " & objInv.Document_No & " "
                ElseIf ((clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal)) AndAlso FormId = "CanSaleInvoice" Then
                    Dim objInv As ClsCanSaleInvoice
                Dim arr As New List(Of String)
                Dim dblCogsCost As Double
                Dim strCogsAcct As String
                objInv = ClsCanSaleInvoice.GetData(obj.Against_Sale_No, "", NavigatorType.Current, trans)

                ''  for tax gl entry start here
                Dim objTM As clsTaxMaster
                Dim dblExcise As Double = 0
                Dim isTaxExcisable As Boolean = False

                If obj.TAX1_Amt <> 0 Then
                    isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX1, trans)
                    ' for excisable tax start here
                    isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX1, trans)
                    If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                        objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX1, trans)
                        If objTM IsNot Nothing Then
                            If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                Throw New Exception("Please set Tax Net Payable Account of Tax Authority " + obj.TAX1)
                            End If
                            objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                            Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX1_Amt}
                            ArryLst.Add(Acc1)
                            dblExcise += obj.TAX1_Amt

                            '' For excisable FOC entry start here on 26/10/2016 for product sale
                            If obj.TAX1_ExciseFOCAmt > 0 Then
                                Acc1 = {objTM.Tax_Net_Payable, obj.TAX1_ExciseFOCAmt}
                                ArryLst.Add(Acc1)
                            End If
                            '' For excisable FOC entry ends here
                        End If
                    End If

                    'Excisable tax ends here

                    If clsCommon.myLen(obj.TAX1_GLAC) <= 0 Then
                        Throw New Exception("GL Acount not found for" + obj.TAX1)
                    End If

                    Dim AccInvDR() As String = {obj.TAX1_GLAC, -1 * obj.TAX1_Amt}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_ExciseFOCAmt > 0 Then
                        AccInvDR = {obj.TAX1_GLAC, -1 * obj.TAX1_ExciseFOCAmt}
                        ArryLst.Add(AccInvDR)
                    End If
                    If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                        dblExcise += obj.TAX1_Amt
                    End If

                End If

                If obj.TAX2_Amt <> 0 Then
                    isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX2, trans)

                    ' for excisable tax start here
                    isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX2, trans)
                    If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                        objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX2, trans)
                        If objTM IsNot Nothing Then
                            If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX2)
                            End If
                            objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                            Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX2_Amt}
                            ArryLst.Add(Acc1)
                            dblExcise += obj.TAX2_Amt

                            '' For excisable FOC entry start here on 26/10/2016 for product sale
                            If obj.TAX2_ExciseFOCAmt > 0 Then
                                Acc1 = {objTM.Tax_Net_Payable, obj.TAX2_ExciseFOCAmt}
                                ArryLst.Add(Acc1)
                                dblExcise += obj.TAX2_ExciseFOCAmt
                            End If
                            '' For excisable FOC entry ends here

                        End If
                    End If
                    'Excisable tax ends here

                    If clsCommon.myLen(obj.TAX2_GLAC) <= 0 Then
                        Throw New Exception("GL Acount not found for" + obj.TAX2)
                    End If

                    Dim AccInvDR() As String = {obj.TAX2_GLAC, -1 * obj.TAX2_Amt}
                    ArryLst.Add(AccInvDR)
                    If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                        dblExcise += obj.TAX2_Amt
                    End If

                End If

                If obj.TAX3_Amt <> 0 Then
                    isTaxRecoverable = clsTaxMaster.ISTaxRecoverableAC(obj.TAX3, trans)

                    ' for excisable tax start here
                    isTaxExcisable = clsTaxMaster.IsTaxExcisable(obj.TAX3, trans)
                    If isTaxExcisable AndAlso clsCommon.CompairString(obj.Trans_Type, "CSA") <> CompairStringResult.Equal Then
                        objTM = clsTaxMaster.GetTaxDetailsForSale(obj.TAX3, trans)
                        If objTM IsNot Nothing Then
                            If clsCommon.myLen(objTM.Tax_Net_Payable) <= 0 Then
                                Throw New Exception("Please set  Tax Net Payable Account of Tax Authority " + obj.TAX3)
                            End If
                            objTM.Tax_Net_Payable = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Net_Payable, obj.loc_code, True, trans)
                            Dim Acc1() As String = {objTM.Tax_Net_Payable, obj.TAX3_Amt}
                            ArryLst.Add(Acc1)
                            dblExcise += obj.TAX3_Amt
                            '' For excisable FOC entry start here on 26/10/2016 for product sale
                            If obj.TAX3_ExciseFOCAmt > 0 Then
                                Acc1 = {objTM.Tax_Net_Payable, obj.TAX3_ExciseFOCAmt}
                                ArryLst.Add(Acc1)
                                dblExcise += obj.TAX3_ExciseFOCAmt
                            End If
                            '' For excisable FOC entry ends here

                        End If
                    End If
                    'Excisable tax ends here

                    If clsCommon.myLen(obj.TAX3_GLAC) <= 0 Then
                        Throw New Exception("GL Acount not found for" + obj.TAX3)
                    End If

                    If clsCommon.CompairString(obj.Trans_Type, "CSA") = CompairStringResult.Equal Then
                        dblExcise += obj.TAX3_Amt
                    End If
                    Dim AccInvDR() As String = {obj.TAX3_GLAC, -1 * obj.TAX3_Amt}
                    ArryLst.Add(AccInvDR)

                End If
                ''tax ends here


                '''' FOR Additional Cost START here
                If obj.Add_Charge_Amt1 <> 0 Then
                    Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                    AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                    End If

                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt2 <> 0 Then
                    Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                    AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                    End If

                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt3 <> 0 Then
                    Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                    AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                    End If

                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt4 <> 0 Then
                    Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                    AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                    End If

                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt5 <> 0 Then
                    Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                    If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                    End If
                    AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)

                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt6 <> 0 Then
                    Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                    AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                    End If
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt7 <> 0 Then
                    Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                    AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                    End If
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt8 <> 0 Then
                    Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                    AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                    End If
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt9 <> 0 Then
                    Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                    AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                    End If
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                    ArryLst.Add(AccAddCostCR)
                End If
                If obj.Add_Charge_Amt10 <> 0 Then
                    Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                    AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)

                    If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                        Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                    End If
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                    ArryLst.Add(AccAddCostCR)
                End If
                '' Additional cost ends here

                Dim isFirstTime As Boolean = True
                For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                    ''Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Total_Amount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                    Dim AccInvDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Amount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                    ArryLst.Add(AccInvDR)
                    isFirstTime = False
                Next
                Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)

                ''Dim AccInvCR() As String = {strACWithLocation, obj.Discount_Base + obj.RoundOffAmount}
                Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                ArryLst.Add(AccInvCR)


                If obj.RoundOffAmount <> 0 Then
                    Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))

                    If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                        Throw New Exception("Please set round off account in Sales Setting")
                    End If

                    strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, strLocation, True, trans)
                    Dim AccRoundInvCR() As String = {strACRoundInvCr, -1 * obj.RoundOffAmount}
                    ArryLst.Add(AccRoundInvCR)
                End If


                If Not isSkipCogsGL Then
                    Dim Costincaseoflossandgain As Double = 0
                    For Each objInvDetail As ClsCanSaleInvoiceDetail In objInv.arrCanSaleInvoiceDetail
                        Dim strCode As String = objInv.CanSale_Dispatch_No
                        If Not arr.Contains(strCode) Then
                            arr.Add(strCode)
                            ''updation by richa agarwal according to gain or loss amount
                            Dim dblCogsCosttemp As Double = 0
                            dblCogsCosttemp = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end) as COst from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & objInv.CanSale_Dispatch_No & "'", trans))
                            ' dblCogsCost += clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (sum(case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end)/Qty) * " & objInvDetail.InvoiceQty & " as COst from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Source_Doc_No='" & objInvDetail.Dispatch_Code & "'", trans))
                            dblCogsCost += dblCogsCosttemp
                            ''''' for cogs entry item wise
                            Dim strSql As String = String.Empty

                            strSql = "select TSPL_INVENTORY_MOVEMENT_NEW.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT_NEW left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT_NEW.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code " & _
                            "where Source_Doc_No='" & objInv.CanSale_Dispatch_No & "'"

                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                For Each dr As DataRow In dt.Rows
                                    strCogsAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans))
                                    If clsCommon.myLen(strCogsAcct) = 0 Then
                                        Throw New Exception("Please set Cost of Goods Sold for first item")
                                    End If
                                    strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)

                                    Dim Acc1() As String = {strCogsAcct, clsCommon.myCdbl(dr("Cost"))}

                                    ArryLst.Add(Acc1)

                                    'Costincaseoflossandgain = clsCommon.myCdbl(dr("Cost"))
                                Next
                            End If

                            'If dblCogsCosttemp <> Costincaseoflossandgain Then
                            '    Dim strGainorLossAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT SA.Gain_Loss_Account  FROM TSPL_ITEM_MASTER AS IM INNER JOIN  TSPL_SALES_ACCOUNTS  AS SA ON IM.Sale_Class_Code  = SA.Sales_Class_Code  INNER JOIN  TSPL_GL_ACCOUNTS AS GLA ON SA.Gain_Loss_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.arrCanSaleInvoiceDetail.Item(0).ItemCode.ToString() + "'", trans))
                            '    If clsCommon.myLen(strGainorLossAC) = 0 Then
                            '        Throw New Exception("Please set Gain/Loss Account for first item")
                            '    End If
                            '    strGainorLossAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strGainorLossAC, obj.loc_code, True, trans)

                            '    Dim Acc2() As String = {strGainorLossAC, 1 * (dblCogsCosttemp - Costincaseoflossandgain)}
                            '    ArryLst.Add(Acc2)

                            'End If

                            ''''' cogs entry item wise ends here
                        End If
                    Next



                    Dim strShipmentClearingAC = clsDBFuncationality.getSingleValue("SELECT PA.Shipment_Clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " & _
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " & _
                     " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + objInv.arrCanSaleInvoiceDetail.Item(0).ItemCode.ToString() + "'", trans)
                    If clsCommon.myLen(strShipmentClearingAC) = 0 Then
                        Throw New Exception("Please set Shipment clearing Account for first item")
                    End If
                    ''richa 13/09/2014 change 
                    '  strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, objInv.Location_Code, trans)
                    strShipmentClearingAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strShipmentClearingAC, obj.loc_code, True, trans)

                    Dim Acc() As String = {strShipmentClearingAC, -1 * dblCogsCost, "", "", "", "", "", "", "H"}
                    ArryLst.Add(Acc)

                End If

                ''richa 25 May,2015 BHA/09/05/18-000021
                If objInv.CanInventoryType = 1 Then
                    Dim strReturnable_ContainerAC As String = clsCommon.myCstr(clsItemMaster.GetReturnableConGLAC(objInv.CanItemCode, trans))
                    If clsCommon.myLen(strReturnable_ContainerAC) = 0 Then
                        Throw New Exception("Please set Returnable Container Account for item")
                    End If
                    strReturnable_ContainerAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strReturnable_ContainerAC, obj.loc_code, True, trans)
                    Dim Acc() As String = {strReturnable_ContainerAC, -1 * (objInv.TotalNoofCans * objInv.CanItemRate)}
                    ArryLst.Add(Acc)

                    Dim strContainerDepositAC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Container_Deposit ,'') as Container_Deposit from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & objInv.Customer_Code & "'", trans))
                    If clsCommon.myLen(strContainerDepositAC) = 0 Then
                        Throw New Exception("Please set Container Deposit Account for customer")
                    End If
                    strContainerDepositAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strContainerDepositAC, obj.loc_code, True, trans)
                    Dim Acc1() As String = {strContainerDepositAC, (objInv.TotalNoofCans * objInv.CanItemRate)}
                    ArryLst.Add(Acc1)
                End If
                ''------------------
                strRemarks = " AR invoice for customer: " + obj.Customer_Code + " - " + obj.Customer_Name + "  For Invoice Can Sale No " & objInv.Document_No & " "

            End If
        ElseIf clsCommon.myLen(obj.AgainstScrap) > 0 Then
            If obj.TAX1_Amt <> 0 Then
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX1)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX1_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX2_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX2)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX2_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX3_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX3)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX3_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX4_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX4)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX4_Amt}
                ArryLst.Add(AccInvDR)
            End If
            If obj.TAX5_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX5)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX5_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX6_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX6)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX6_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX7_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX7)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX7_Amt}
                ArryLst.Add(AccInvDR)

            End If

            If obj.TAX8_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX8)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX8_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX9_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX9)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX9_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX10_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX10)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX10_Amt}
                ArryLst.Add(AccInvDR)
            End If

            '' FOR Additional Cost START here
            If obj.Add_Charge_Amt1 <> 0 Then
                Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt2 <> 0 Then
                Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt3 <> 0 Then
                Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt4 <> 0 Then
                Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt5 <> 0 Then
                Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt6 <> 0 Then
                Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt7 <> 0 Then
                Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt8 <> 0 Then
                Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt9 <> 0 Then
                Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt10 <> 0 Then
                Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                ArryLst.Add(AccAddCostCR)
            End If

            For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                Dim dblLedgeerNonRecoverableAmt As Double = 0
                Dim AccInvDR1() As String = {objTR.GL_Account_Code, -1 * objTR.Amount_less_Discount, "", "", "", "", "", "", objTR.Reco_Control_Account}
                ArryLst.Add(AccInvDR1)
            Next

            If obj.RoundOffAmount <> 0 Then
                Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))
                If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                    Throw New Exception("Please set round off account in Sales Setting")
                End If
                strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, obj.loc_code, True, trans)
                Dim AccRoundInvCR() As String = {strACRoundInvCr, -1 * obj.RoundOffAmount}
                ArryLst.Add(AccRoundInvCR)
            End If

            Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
            Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)
            Dim AccInvCR1() As String = {strACWithLocation, obj.Document_Total}
            ArryLst.Add(AccInvCR1)

            If Not isSkipCogsGL Then 'And Not IsAllowPurchaseAccounting
                Dim strInventoryControlAc As String
                Dim strCogsAcct As String
                Dim strSql As String = "select TSPL_INVENTORY_MOVEMENT.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  where Source_Doc_No='" & obj.AgainstScrap & "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                    For Each dr As DataRow In dt.Rows
                        strInventoryControlAc = clsDBFuncationality.getSingleValue("SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " & _
                        " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " & _
                        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                        strInventoryControlAc = clsERPFuncationality.ChangeGLAccountLocationSegment(strInventoryControlAc, obj.loc_code, True, trans)

                        If clsCommon.myLen(strInventoryControlAc) = 0 Then
                            Throw New Exception("Please set Inventory Control Account for first item")
                        End If

                        strCogsAcct = clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                        If clsCommon.myLen(strCogsAcct) = 0 Then
                            Throw New Exception("Please set Cost of Goods Sold for first item")
                        End If
                        strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                        'Sanjay Journal entry RecoControlACC
                        Dim Acc() As String = {strInventoryControlAc, -1 * clsCommon.myCdbl(dr("Cost")), "", "", "", "", "", "", RecoControlACC}
                        ArryLst.Add(Acc)
                        Dim Acc1() As String = {strCogsAcct, clsCommon.myCdbl(dr("Cost"))}
                        ArryLst.Add(Acc1)
                    Next
                End If
            End If
        ElseIf clsCommon.myLen(obj.AgainstScrapReturn) > 0 Then
            If obj.TAX1_Amt <> 0 Then
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX1)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX1_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX2_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX2)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX2_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX3_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX3)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX3_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX4_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX4)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX4_Amt}
                ArryLst.Add(AccInvDR)
            End If
            If obj.TAX5_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX5)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX5_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX6_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX6)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX6_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX7_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX7)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX7_Amt}
                ArryLst.Add(AccInvDR)

            End If

            If obj.TAX8_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX8)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX8_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX9_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX9)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX9_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX10_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX10)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, 1 * obj.TAX10_Amt}
                ArryLst.Add(AccInvDR)
            End If

            '' FOR Additional Cost START here
            If obj.Add_Charge_Amt1 <> 0 Then
                Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, 1 * obj.Add_Charge_Amt1}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt2 <> 0 Then
                Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, 1 * obj.Add_Charge_Amt2}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt3 <> 0 Then
                Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, 1 * obj.Add_Charge_Amt3}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt4 <> 0 Then
                Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt5 <> 0 Then
                Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, 1 * obj.Add_Charge_Amt5}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt6 <> 0 Then
                Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, 1 * obj.Add_Charge_Amt6}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt7 <> 0 Then
                Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, 1 * obj.Add_Charge_Amt7}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt8 <> 0 Then
                Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt9 <> 0 Then
                Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, 1 * obj.Add_Charge_Amt9}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt10 <> 0 Then
                Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, 1 * obj.Add_Charge_Amt10}
                ArryLst.Add(AccAddCostCR)
            End If

            For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                Dim dblLedgeerNonRecoverableAmt As Double = 0
                Dim AccInvDR1() As String = {objTR.GL_Account_Code, 1 * objTR.Amount_less_Discount, "", "", "", "", "", "", objTR.Reco_Control_Account}
                ArryLst.Add(AccInvDR1)
            Next

            If obj.RoundOffAmount <> 0 Then
                Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))
                If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                    Throw New Exception("Please set round off account in Sales Setting")
                End If
                strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, obj.loc_code, True, trans)
                Dim AccRoundInvCR() As String = {strACRoundInvCr, obj.RoundOffAmount}
                ArryLst.Add(AccRoundInvCR)
            End If

            Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
            Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, strLocation, True, trans)
            Dim AccInvCR1() As String = {strACWithLocation, -1 * obj.Document_Total}
            ArryLst.Add(AccInvCR1)

            If Not isSkipCogsGL Then 'And Not IsAllowPurchaseAccounting
                Dim strInventoryControlAc As String
                Dim strCogsAcct As String
                ''richa agarwal use scarp sale invoice no instead of against scrap sale return no to fetch cogs amaount ERO/19/03/19-000515 28 March,2019
                Dim strSql As String = "select TSPL_INVENTORY_MOVEMENT.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from TSPL_INVENTORY_MOVEMENT left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  where Source_Doc_No='" & obj.RefDocNo & "' "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                    For Each dr As DataRow In dt.Rows
                        strInventoryControlAc = clsDBFuncationality.getSingleValue("SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " & _
                        " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " & _
                        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                        strInventoryControlAc = clsERPFuncationality.ChangeGLAccountLocationSegment(strInventoryControlAc, obj.loc_code, True, trans)

                        If clsCommon.myLen(strInventoryControlAc) = 0 Then
                            Throw New Exception("Please set Inventory Control Account for first item")
                        End If

                        strCogsAcct = clsDBFuncationality.getSingleValue("select Cost_Of_Goods_Sold from TSPL_ITEM_MASTER left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                        If clsCommon.myLen(strCogsAcct) = 0 Then
                            Throw New Exception("Please set Cost of Goods Sold for first item")
                        End If
                        strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                        '' richa agarwal 28 Mar,2019 add account into inventory movement table,ERO/19/03/19-000515
                        Dim Acc() As String = {strInventoryControlAc, 1 * clsCommon.myCdbl(dr("Cost")), "", "", "", "", "", "", RecoControlACC}
                        ArryLst.Add(Acc)
                        If clsCommon.CompairString(RecoControlACC, "I") = CompairStringResult.Equal Then
                            clsInventoryMovement.UpdateInvControlAccount(obj.AgainstScrapReturn, "MS-SR", clsCommon.myCstr(dr("Item_Code")), strInventoryControlAc, "", "", trans)
                        End If
                        Dim Acc1() As String = {strCogsAcct, -1 * clsCommon.myCdbl(dr("Cost"))}
                        ArryLst.Add(Acc1)
                    Next
                End If
            End If
        ElseIf clsCommon.myLen(obj.Against_Asset_Disposal) > 0 Then
            'Dim qryAsset As String = clsAssetScrapSaleHead.GetAssetDisposalJEQuery(obj.Against_Asset_Disposal)
            Dim qryAsset As String = clsAssetScrapSaleHead.GetAssetDisposalJEQuery_ARInvoice(obj.Against_Asset_Disposal)
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qryAsset, trans)

            '' customer control acc debit with sale amt
            Dim Acc() As String = {clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, obj.loc_code, True, trans), obj.Document_Total}
            ArryLst.Add(Acc)
            For Each dr As DataRow In dt.Rows
                '' get accounts
                'Dim Receivable_Control_acct As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("Receivable_Control_acct")), obj.loc_code, trans)
                Dim Disposal_Account As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("Disposal_Account")), obj.loc_code, True, trans)
                Dim Disposal_Cost_Account As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("Disposal_Cost_Account")), obj.loc_code, True, trans)
                Dim Ac_Accum_Dep As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("Ac_Accum_Dep")), obj.loc_code, True, trans)
                Dim Ac_Control As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("Ac_Control")), obj.loc_code, True, trans)
                Dim PROFIT_AC As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("PROFIT_AC")), obj.loc_code, True, trans)
                Dim LOSS_AC As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dr.Item("LOSS_AC")), obj.loc_code, True, trans)
                Dim strcostCentre As String = clsCommon.myCstr(dr.Item("CostCenter_Code"))
                Dim strHirerachy_Code As String = clsCommon.myCstr(dr.Item("Hirerachy_Code"))

                '' get amounts
                Dim Book_Source_value As Decimal = clsCommon.myCdbl(dr.Item("Book_Source_value"))
                Dim Sale_Amount As Decimal = clsCommon.myCdbl(dr.Item("Sale_Amount"))
                Dim Perm_Dep_Amount As Decimal = clsCommon.myCdbl(dr.Item("Final_Dep_Amount"))

                '' disposal acc credit with sale amt
                Dim Acc1() As String = {Disposal_Account, -1 * Sale_Amount, "", "", strHirerachy_Code, strcostCentre}
                ArryLst.Add(Acc1)

                '' Accumulated Depreciation a/c debit with depreciation amt
                Dim Acc2() As String = {Ac_Accum_Dep, Perm_Dep_Amount, "", "", strHirerachy_Code, strcostCentre}
                ArryLst.Add(Acc2)

                '' Assets Control a/c credit with book source value 
                Dim Acc3() As String = {Ac_Control, -1 * Book_Source_value, "", "", strHirerachy_Code, strcostCentre}
                ArryLst.Add(Acc3)

                If (Sale_Amount - Book_Source_value + Perm_Dep_Amount) > 0 Then
                    '' profit a/c credit with depreciation amt
                    Dim Acc4() As String = {PROFIT_AC, -1 * (Sale_Amount - Book_Source_value + Perm_Dep_Amount), "", "", strHirerachy_Code, strcostCentre}
                    ArryLst.Add(Acc4)
                Else
                    '' loss a/c debit with depreciation amt
                    Dim Acc5() As String = {LOSS_AC, -1 * (Sale_Amount - Book_Source_value + Perm_Dep_Amount), "", "", strHirerachy_Code, strcostCentre}
                    ArryLst.Add(Acc5)
                End If

                '' customer control acc debit with sale amt
                Dim Acc6() As String = {Disposal_Cost_Account, Sale_Amount, "", "", strHirerachy_Code, strcostCentre}
                ArryLst.Add(Acc6)
            Next
            If obj.TAX1_Amt <> 0 Then
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX1)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX1_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX2_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX2)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX2_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX3_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX3)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX3_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX4_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX4)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX4_Amt}
                ArryLst.Add(AccInvDR)
            End If
            If obj.TAX5_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX5)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX5_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX6_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX6)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX6_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX7_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX7)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX7_Amt}
                ArryLst.Add(AccInvDR)

            End If

            If obj.TAX8_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX8)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX8_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX9_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX9)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX9_Amt}
                ArryLst.Add(AccInvDR)
            End If

            If obj.TAX10_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX10)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX10_Amt}
                ArryLst.Add(AccInvDR)
            End If

            '' FOR Additional Cost START here
            If obj.Add_Charge_Amt1 <> 0 Then
                Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt2 <> 0 Then
                Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt3 <> 0 Then
                Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt4 <> 0 Then
                Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt5 <> 0 Then
                Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt6 <> 0 Then
                Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt7 <> 0 Then
                Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt8 <> 0 Then
                Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt9 <> 0 Then
                Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                ArryLst.Add(AccAddCostCR)
            End If
            If obj.Add_Charge_Amt10 <> 0 Then
                Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                End If
                Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                ArryLst.Add(AccAddCostCR)
            End If

        ElseIf clsCommon.myLen(obj.Against_Sale_Return_No) > 0 OrElse _
               clsCommon.myLen(obj.Against_MCC_Material_Sale_Return) > 0 OrElse _
               clsCommon.myLen(obj.Against_VCGL) > 0 OrElse _
           (clsCommon.myLen(obj.Against_Sale_No) <= 0 AndAlso clsCommon.myLen(obj.AgainstScrap) <= 0 AndAlso clsCommon.myLen(obj.Against_Sale_Return_No) <= 0 AndAlso clsCommon.myLen(obj.Against_MCC_Material_Sale_Return) <= 0 AndAlso clsCommon.myLen(obj.Against_VCGL) <= 0) Then ''BM00000007723 For direct AR Invoice
            If obj.TAX1_Amt <> 0 Then
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX1, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX1)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX1_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX1_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX2_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX2)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX2_Amt}
                    ArryLst.Add(AccInvDR)

                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX2_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX3_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX3)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX3_Amt}
                    ArryLst.Add(AccInvDR)


                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX3_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX4_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX4)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX4_Amt}
                    ArryLst.Add(AccInvDR)


                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX4_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If
            If obj.TAX5_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX5)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX5_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX5_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX6_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX6)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX6_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX6_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX7_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX7)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX7_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX7_Amt}
                    ArryLst.Add(AccInvDR)
                End If
            End If

            If obj.TAX8_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX8)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX8_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX8_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX9_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX9)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX9_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX9_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            If obj.TAX10_Amt <> 0 Then
                'isTaxRecoverable = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans)
                Tax_Liability_Account = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
                If clsCommon.myLen(Tax_Liability_Account) <= 0 Then
                    Throw New Exception("Liability Acount not found for Tax " + obj.TAX10)
                End If
                Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax_Liability_Account, obj.loc_code, True, trans)
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccInvDR() As String = {Tax_Liability_Account, -1 * obj.TAX10_Amt}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {Tax_Liability_Account, obj.TAX10_Amt}
                    ArryLst.Add(AccInvDR)
                End If

            End If

            '' FOR Additional Cost START here
            If obj.Add_Charge_Amt1 <> 0 Then
                Dim AddCharge_GL_Acc1 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code1, trans)
                AddCharge_GL_Acc1 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc1, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc1) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code1 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, -1 * obj.Add_Charge_Amt1}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc1, obj.Add_Charge_Amt1}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt2 <> 0 Then
                Dim AddCharge_GL_Acc2 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code2, trans)
                AddCharge_GL_Acc2 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc2, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc2) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code2 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, -1 * obj.Add_Charge_Amt2}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc2, obj.Add_Charge_Amt2}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt3 <> 0 Then
                Dim AddCharge_GL_Acc3 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code3, trans)
                AddCharge_GL_Acc3 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc3, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc3) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code3 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, -1 * obj.Add_Charge_Amt3}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc3, obj.Add_Charge_Amt3}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt4 <> 0 Then
                Dim AddCharge_GL_Acc4 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code4, trans)
                AddCharge_GL_Acc4 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc4, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc4) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code4 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, -1 * obj.Add_Charge_Amt4}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc4, obj.Add_Charge_Amt4}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt5 <> 0 Then
                Dim AddCharge_GL_Acc5 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code5, trans)
                AddCharge_GL_Acc5 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc5, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc5) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code5 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, -1 * obj.Add_Charge_Amt5}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc5, obj.Add_Charge_Amt5}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt6 <> 0 Then
                Dim AddCharge_GL_Acc6 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code6, trans)
                AddCharge_GL_Acc6 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc6, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc6) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code6 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, -1 * obj.Add_Charge_Amt6}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc6, obj.Add_Charge_Amt6}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt7 <> 0 Then
                Dim AddCharge_GL_Acc7 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code7, trans)
                AddCharge_GL_Acc7 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc7, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc7) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code7 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, -1 * obj.Add_Charge_Amt7}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc7, obj.Add_Charge_Amt7}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt8 <> 0 Then
                Dim AddCharge_GL_Acc8 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code8, trans)
                AddCharge_GL_Acc8 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc8, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc8) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code8 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, -1 * obj.Add_Charge_Amt8}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc8, obj.Add_Charge_Amt8}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            If obj.Add_Charge_Amt9 <> 0 Then
                Dim AddCharge_GL_Acc9 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code9, trans)
                AddCharge_GL_Acc9 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc9, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc9) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code9 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, -1 * obj.Add_Charge_Amt9}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc9, obj.Add_Charge_Amt9}
                    ArryLst.Add(AccAddCostCR)
                End If
            End If
            If obj.Add_Charge_Amt10 <> 0 Then
                Dim AddCharge_GL_Acc10 = clsAdditionalCharge.GetAdditonalChACC(obj.Add_Charge_Code10, trans)
                AddCharge_GL_Acc10 = clsERPFuncationality.ChangeGLAccountLocationSegment(AddCharge_GL_Acc10, obj.loc_code, True, trans)
                If clsCommon.myLen(AddCharge_GL_Acc10) <= 0 Then
                    Throw New Exception("Additional GL Account for " & obj.Add_Charge_Code10 & " Not found")
                End If
                If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, -1 * obj.Add_Charge_Amt10}
                    ArryLst.Add(AccAddCostCR)
                Else
                    Dim AccAddCostCR() As String = {AddCharge_GL_Acc10, obj.Add_Charge_Amt10}
                    ArryLst.Add(AccAddCostCR)
                End If

            End If
            '' Additional cost ends here

            Dim isFirstTime As Boolean = True
            '' Anubhooti 19-Mar-2015 (IF Entry is against VCGL then GL will get opposite(DR/CR))
            If clsCommon.myLen(obj.Against_VCGL) <= 0 Then 'AndAlso clsCommon.CompairString(obj.Trans_Type, "VC") <> CompairStringResult.Equal
                    For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                        Dim dblLedgeerNonRecoverableAmt As Double = 0
                        ''richa agarwal 14/05/2015 BM00000006615 credit gl account in case of direct ar invoice which type of Invoice
                        If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                            ' Dim AccInvDR1() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount * -1}
                            ''richa 15/03/2017
                            'Dim AccInvDR1() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount * -1}
                            ' Dim AccInvDR1() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount * -1, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code}
                            Dim AccInvDR1() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount * -1, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4, objTR.Reco_Control_Account}
                            ArryLst.Add(AccInvDR1)
                        Else
                            ''richa 15/03/2017
                            '  Dim AccInvDR() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount}
                            ' Dim AccInvDR() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code}
                            Dim AccInvDR() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4, objTR.Reco_Control_Account}
                            ArryLst.Add(AccInvDR)
                        End If
                        ''--------------------------------------
                        isFirstTime = False

                        ''''''added by priti for discount entry of Return
                        If FormId = "FreshSaleReturn" Then
                            If objTR.Amount_less_Discount = 0 AndAlso objTR.Discount > 0 Then
                                Dim AccDiscDR() As String = {objTR.GL_Account_Code, -1 * (objTR.Discount), "", "", "", "", "", "", objTR.Reco_Control_Account}
                                ArryLst.Add(AccDiscDR)
                            End If
                        End If
                        ''''''code ends here

                        If clsCommon.myLen(obj.Against_MCC_Material_Sale_Return) > 0 AndAlso objTR.Transporter_Commission_Amt > 0 Then
                            Dim strLocationt As String = ""
                            Dim strACWithLocationt As String = ""
                            Dim AccDiscDR() As String = Nothing
                            Dim AccDiscTaxDR() As String = Nothing

                            If obj.Is_Add_TPT Then
                                AccDiscTaxDR = {objTR.Transporter_GL_Account_Code, (objTR.Transporter_Commission_Amt)}
                            Else
                                AccDiscTaxDR = {objTR.Transporter_GL_Account_Code, -1 * (objTR.Transporter_Commission_Amt)}
                            End If
                            ArryLst.Add(AccDiscTaxDR)
                        End If
                    Next

                    ''--------done by richa SWA/12/09/18-000052 for swadesh to add leakage account in fresh journal entry
                    Dim dblLeakAmount As Double = 0
                    Dim LeakageAcct As String = ""
                    If clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal AndAlso FormId = "FreshSaleReturn" Then
                        If obj.LeakageAmount > 0 Then
                            dblLeakAmount = obj.LeakageAmount
                            LeakageAcct = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Leakage_Deduction from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + obj.Account_Set + "'", trans))
                            LeakageAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(LeakageAcct, obj.loc_code, True, trans)
                            If clsCommon.myLen(LeakageAcct) <= 0 Then
                                Throw New Exception("Please set Leakage account set of customer account set :" + obj.Account_Set)
                            End If
                            If dblLeakAmount > 0 Then
                                Dim AccLeakageAcc() As String = {LeakageAcct, dblLeakAmount * -1}
                                ArryLst.Add(AccLeakageAcc)
                            End If
                        End If
                    End If

                    ''------------------------------------

                    Dim AllowCrateCanPhysicalStock As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, trans))
                    If AllowCrateCanPhysicalStock = 1 Then
                        ' DOne by priti BHA/15/06/18-000055
                        If FormId = clsUserMgtCode.frmSaleReturndairy Then

                            ' FOr Crate
                            Dim strCrateItem = ""
                            Dim strCrateUOM = ""
                            Dim dblCrateRate As Integer = 0
                            Dim dblCrateQty As Integer = 0
                            Dim strReturnable_ContainerAC As String = ""
                            Dim strContainerDepositAC As String = ""
                            Dim Acc() As String = Nothing
                            Dim Acc1() As String = Nothing
                            qry = "select Crate_Item,Crate_ItemUnit,Crate_ItemRate,CrateQty from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & obj.Against_Sale_Return_No & "'"
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                strCrateItem = clsCommon.myCstr(dt.Rows(0)("Crate_Item"))
                                strCrateUOM = clsCommon.myCstr(dt.Rows(0)("Crate_ItemUnit"))
                                dblCrateRate = clsCommon.myCdbl(dt.Rows(0)("Crate_ItemRate"))
                                dblCrateQty = clsCommon.myCdbl(dt.Rows(0)("CrateQty"))
                            End If
                            If dblCrateQty > 0 Then
                                strReturnable_ContainerAC = clsCommon.myCstr(clsItemMaster.GetReturnableConGLAC(strCrateItem, trans))
                                If clsCommon.myLen(strReturnable_ContainerAC) = 0 Then
                                    Throw New Exception("Please set Returnable Container Account for item - " + strCrateItem)
                                End If
                                strReturnable_ContainerAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strReturnable_ContainerAC, obj.loc_code, True, trans)
                                Acc = {strReturnable_ContainerAC, 1 * (dblCrateQty * dblCrateRate)}
                                ArryLst.Add(Acc)
                                strContainerDepositAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Container_Deposit ,'') as Container_Deposit from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                                If clsCommon.myLen(strContainerDepositAC) = 0 Then
                                    Throw New Exception("Please set Container Deposit Account for customer - " + obj.Customer_Code)
                                End If
                                strContainerDepositAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strContainerDepositAC, obj.loc_code, True, trans)
                                Acc1 = {strContainerDepositAC, -1 * (dblCrateQty * dblCrateRate)}
                                ArryLst.Add(Acc1)
                            End If
                            ' FOr Can
                            Dim strCanItem = ""
                            Dim strCanUOM = ""
                            Dim dblCanRate As Integer = 0
                            Dim dblCanQty As Integer = 0
                            qry = "select Can_Item,Can_ItemUnit,Can_ItemRate,ShippedCAN from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & obj.Against_Sale_Return_No & "'"
                            dt = clsDBFuncationality.GetDataTable(qry, trans)
                            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                                strCanItem = clsCommon.myCstr(dt.Rows(0)("Can_Item"))
                                strCanUOM = clsCommon.myCstr(dt.Rows(0)("Can_ItemUnit"))
                                dblCanRate = clsCommon.myCdbl(dt.Rows(0)("Can_ItemRate"))
                                dblCanQty = clsCommon.myCdbl(dt.Rows(0)("ShippedCAN"))
                            End If
                            If dblCanQty > 0 Then
                                strReturnable_ContainerAC = clsCommon.myCstr(clsItemMaster.GetReturnableConGLAC(strCanItem, trans))
                                If clsCommon.myLen(strReturnable_ContainerAC) = 0 Then
                                    Throw New Exception("Please set Returnable Container Account for item " + strCanItem)
                                End If
                                strReturnable_ContainerAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strReturnable_ContainerAC, obj.loc_code, True, trans)
                                Acc = Nothing
                                Acc = {strReturnable_ContainerAC, 1 * (dblCanQty * dblCanRate)}
                                ArryLst.Add(Acc)

                                ''richa agarwal 6 Sep,2018
                                strContainerDepositAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Container_Deposit ,'') as Container_Deposit from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                                If clsCommon.myLen(strContainerDepositAC) = 0 Then
                                    Throw New Exception("Please set Container Deposit Account for customer - " + obj.Customer_Code)
                                End If
                                strContainerDepositAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strContainerDepositAC, obj.loc_code, True, trans)
                                ''------------------
                                Acc1 = Nothing
                                Acc1 = {strContainerDepositAC, -1 * (dblCanQty * dblCanRate)}
                                ArryLst.Add(Acc1)
                            End If
                        End If
                    End If

                Else '' New Part 
                    For Each objTR As clsCustomerInvoiceDetail In obj.Arr
                        Dim dblLedgeerNonRecoverableAmt As Double = 0
                        ''richa agarwal 21/07/2015  debit/credit customer account in case of vcgl
                        If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                            Dim AccInvDR1() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount * -1, "", "", "", "", "", "", objTR.Reco_Control_Account}
                            ArryLst.Add(AccInvDR1)
                        Else
                            Dim AccInvDR() As String = {objTR.GL_Account_Code, objTR.Amount_less_Discount, "", "", "", "", "", "", objTR.Reco_Control_Account}
                            ArryLst.Add(AccInvDR)
                        End If
                        isFirstTime = False
                    Next
                End If
                ''richa agarwal 24/11/2014

                If obj.RoundOffAmount <> 0 Then
                    Dim strACRoundInvCr As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DefaultRoundOffGLAccount, clsFixedParameterCode.DefaultRoundOffGLAccount, trans))
                    If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                        Throw New Exception("Please set round off account in Sales Setting")
                    End If
                    strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, obj.loc_code, True, trans)
                    Dim AccRoundInvCR() As String = {strACRoundInvCr, obj.RoundOffAmount}
                    ArryLst.Add(AccRoundInvCR)
                End If
                If clsCommon.myLen(obj.Against_MCC_Material_Sale_Return) > 0 AndAlso obj.TotalSubsidyAmt > 0 Then
                    Dim strRate_Difference_act As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_CUSTOMER_ACCOUNT_SET.Rate_Difference ,'') as Rate_Difference from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Account  =TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code ='" & obj.Customer_Code & "'", trans))
                    If clsCommon.myLen(strRate_Difference_act) = 0 Then
                        Throw New Exception("Please set Rate Difference Account for customer - " + obj.Customer_Code)
                    End If
                    strRate_Difference_act = clsERPFuncationality.ChangeGLAccountLocationSegment(strRate_Difference_act, obj.loc_code, True, trans)
                    Dim AccSub() As String = {strRate_Difference_act, (-obj.TotalSubsidyAmt)}
                    ArryLst.Add(AccSub)
                End If
                ''------------------------

                Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
                Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Customer_Control_AC, obj.loc_code, True, trans)

                '' Anubhooti 19-Mar-2015 (IF Entry is against VCGL then GL will get opposite(DR/CR))
                If clsCommon.myLen(obj.Against_VCGL) <= 0 Then 'AndAlso clsCommon.CompairString(obj.Trans_Type, "VC") = CompairStringResult.Equal
                    ''richa agarwal 14/05/2015 BM00000006615 debit customer account in case of direct ar invoice which type of Invoice
                    If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                        Dim AccInvCR1() As String = {strACWithLocation, obj.Document_Total}
                        ArryLst.Add(AccInvCR1)
                    Else
                        ''richa agarwal add/subtract round off amount from customer amount
                        Dim AccInvCR() As String = {strACWithLocation, -1 * obj.Document_Total}
                        ' Dim AccInvCR() As String = {strACWithLocation, -1 * (obj.Document_Total - obj.RoundOffAmount)}
                        ArryLst.Add(AccInvCR)
                    End If

                Else '' New Part
                    ''richa agarwal 21/07/2015  debit/credit customer account in case of vcgl
                    If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                        Dim AccInvCR1() As String = {strACWithLocation, obj.Document_Total}
                        ArryLst.Add(AccInvCR1)
                    Else
                        Dim AccInvCR() As String = {strACWithLocation, -1 * obj.Document_Total}
                        ArryLst.Add(AccInvCR)
                    End If
                    'Dim AccInvCR() As String = {strACWithLocation, obj.Document_Total}
                    'ArryLst.Add(AccInvCR)
                    ''-----------------------------
                End If

                If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL'And Not IsAllowPurchaseAccounting
                    '' FOR cogs START here
                    If ((clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal)) Then
                        ' Dim objInv As clsSNInvoiceHead
                        Dim strFirstItem As String
                        Dim arr As New List(Of String)
                        ' Dim dblCogsCost As Double
                        Dim strInventoryControlAc As String
                        Dim strCogsAcct As String
                        Dim strInvNo As String = String.Empty
                        Dim strSql As String = String.Empty
                        Dim SentschemecogsinAnotherAccount As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SentschemecogsinAnotherAccount, clsFixedParameterCode.SentschemecogsinAnotherAccount, trans)) = "1", True, False))
                        ''richa agarwal on 27 March,2019 against ticket no ERO/19/03/19-000517
                        strInvNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Invoice_Code from TSPL_sd_SALE_RETURN_DETAIL where DOCUMENT_CODE='" & IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return) & "'", trans))
                        strFirstItem = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Item_Code from TSPL_sd_SALE_RETURN_DETAIL where DOCUMENT_CODE='" & IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return) & "'", trans))
                        ' removed crate and can item for cogs return for BHA/27/07/18-000197 bharat
                        strSql = "select inv.Is_Scheme_Item,inv.Item_Code,case when Costing_Method=0 then Avg_Cost when Costing_Method=1 then Avg_Cost when Costing_Method=2 then FIFO_Cost when Costing_Method=3 then LIFO_Cost end as Cost from " &
                            " ( select isnull(TSPL_INVENTORY_MOVEMENT.Is_Scheme_Item,'N') as Is_Scheme_Item,Item_Code,Source_Doc_No,Avg_Cost,FIFO_Cost,LIFO_Cost from TSPL_INVENTORY_MOVEMENT where TSPL_INVENTORY_MOVEMENT.Source_Doc_No='" & IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return) & "' " &
                            " union all " &
                            " select isnull(TSPL_INVENTORY_MOVEMENT_NEW.Is_Scheme_Item,'N') as Is_Scheme_Item,Item_Code,Source_Doc_No,Avg_Cost,FIFO_Cost,LIFO_Cost from TSPL_INVENTORY_MOVEMENT_NEW where TSPL_INVENTORY_MOVEMENT_NEW.Source_Doc_No='" & IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return) & "' " &
                            " ) inv " &
                            " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=inv.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  where isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0 and Source_Doc_No='" & IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return) & "'"

                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(strSql, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            For Each dr As DataRow In dt.Rows
                                strInventoryControlAc = clsDBFuncationality.getSingleValue("SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                            " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                            " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'", trans)
                                strInventoryControlAc = clsERPFuncationality.ChangeGLAccountLocationSegment(strInventoryControlAc, obj.loc_code, True, trans)

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

                                strCogsAcct = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAcct, obj.loc_code, True, trans)
                                '' change ends here ''richa agarwal BHA/27/11/18-000719 send I for inventory account into journal detail table
                                Dim Acc() As String = {strInventoryControlAc, clsCommon.myCdbl(dr("Cost")), "", "", "", "", "", "", RecoControlACC}
                                ArryLst.Add(Acc)
                                ''---------------''richa agarwal BHA/27/11/18-000719 27 Dec,2018 add account into inventory movement table
                                Dim trans_type As String = String.Empty
                                If FormId = clsUserMgtCode.frmSaleReturndairy OrElse clsCommon.CompairString(FormId, "FreshSaleReturn") = CompairStringResult.Equal Then
                                    trans_type = "FS-SR"
                                ElseIf FormId = "" Then
                                    If clsCommon.myLen(obj.Against_MCC_Material_Sale_Return) > 0 Then
                                        trans_type = "MCC-MSR"
                                    Else
                                        trans_type = "PS-SR"
                                    End If
                                    ''richa TEC/21/02/19-000428 21 Feb,2019
                                ElseIf clsCommon.CompairString(FormId, "SaleReturnBS") = CompairStringResult.Equal Then
                                    trans_type = "SaleReturnBS"
                                End If
                                If clsCommon.CompairString(RecoControlACC, "I") = CompairStringResult.Equal Then
                                    clsInventoryMovement.UpdateInvControlAccount(IIf(clsCommon.myLen(obj.Against_Sale_Return_No) > 0, obj.Against_Sale_Return_No, obj.Against_MCC_Material_Sale_Return), trans_type, clsCommon.myCstr(dr("Item_Code")), strInventoryControlAc, "", "", trans)
                                End If
                                '------------------
                                Dim Acc1() As String = {strCogsAcct, -1 * clsCommon.myCdbl(dr("Cost"))}
                                ArryLst.Add(Acc1)
                            Next
                        End If
                    End If
                End If
            Else
                Throw New Exception("Document is not implemented")
            End If
            If clsCommon.CompairString(FormId, "CSA-SALE") <> CompairStringResult.Equal Then
                strEntryDesc = strEntryDesc + obj.Document_No
            End If
            ''richa 6 Feb,2020
            Dim objJE As New clsJEExtraColumns
        If clsCommon.myLen(obj.TapalNo) > 0 Or clsCommon.myLen(obj.DateAndTime) > 0 Then
            objJE.TapalNo = clsCommon.myCstr(obj.TapalNo)
            If clsCommon.myLen(obj.DateAndTime) > 0 Then
                objJE.DateAndTime = obj.DateAndTime
            End If
        End If

        'If IsAllowPurchaseAccounting Then
        clsJournalMaster.FunGrnlEntryWithTrans(obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, obj.Document_Date, strEntryDesc, strSrcType, strSrcDesc, obj.Document_No, obj.Description, "C", obj.Customer_Code, obj.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, "", strRemarks, Nothing, coll, objJE)
            'End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsCustomerInvoiceDetail
#Region "Variables"
    Public Reco_Control_Account As String = Nothing
    Public SNo As Integer = 0
    Public Document_No As String = Nothing
    Public GL_Account_Code As String = Nothing
    Public GL_Account_Desc As String = Nothing
    Public Amount As Double = 0
    Public Discount_Per As Double = 0
    Public Discount As Double = 0
    Public Amount_less_Discount As String = Nothing
    Public TAX1 As String = Nothing
    Public TAX1_Rate As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Rate As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Rate As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Rate As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Rate As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Rate As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Rate As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Rate As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Rate As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Rate As Double = 0
    Public TAX10_Amt As Double = 0
    Public Total_Tax As Double = 0
    Public Total_Amount As Double = 0
    Public Remarks As String = Nothing
    Public Comments As String = Nothing
    Public TAX1_Base_Amt As Double = 0
    Public TAX2_Base_Amt As Double = 0
    Public TAX3_Base_Amt As Double = 0
    Public TAX4_Base_Amt As Double = 0
    Public TAX5_Base_Amt As Double = 0
    Public TAX6_Base_Amt As Double = 0
    Public TAX7_Base_Amt As Double = 0
    Public TAX8_Base_Amt As Double = 0
    Public TAX9_Base_Amt As Double = 0
    Public TAX10_Base_Amt As Double = 0
    Public Distributor_Commission_Amt As Double = 0
    Public Transporter_Commission_Amt As Double = 0
    Public Security_Amt As Double = 0
    '' richa
    Public Hirerachy_Code As String = String.Empty
    Public Cost_Centre_Code As String = String.Empty
    Public Hirerachy_Code1 As String = String.Empty
    Public Hirerachy_Code2 As String = String.Empty
    Public Hirerachy_Code3 As String = String.Empty
    Public Hirerachy_Code4 As String = String.Empty
    Public AddChargeCode As String = String.Empty
    Public AddChargeDesc As String = String.Empty
    Public Promotional_GL_Account_Code As String = String.Empty
    Public Transporter_GL_Account_Code As String = String.Empty
    Public SD_GL_Account_Code As String = String.Empty

#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsCustomerInvoiceDetail), ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each obj As clsCustomerInvoiceDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "SNo", obj.SNo)
                clsCommon.AddColumnsForChange(coll, "GL_Account_Code", obj.GL_Account_Code)
                clsCommon.AddColumnsForChange(coll, "GL_Account_Desc", obj.GL_Account_Desc)
                clsCommon.AddColumnsForChange(coll, "Reco_Control_Account", obj.Reco_Control_Account)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Discount_Per", obj.Discount_Per)
                clsCommon.AddColumnsForChange(coll, "Discount", obj.Discount)
                clsCommon.AddColumnsForChange(coll, "Amount_less_Discount", obj.Amount_less_Discount)
                clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
                clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
                clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
                clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
                clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
                clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
                clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
                clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
                clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
                clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
                clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
                clsCommon.AddColumnsForChange(coll, "Total_Tax", obj.Total_Tax)
                clsCommon.AddColumnsForChange(coll, "Total_Amount", obj.Total_Amount)
                clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
                clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)
                clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code", obj.Hirerachy_Code, True)
                clsCommon.AddColumnsForChange(coll, "Cost_Centre_Code", obj.Cost_Centre_Code, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code1", obj.Hirerachy_Code1, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code2", obj.Hirerachy_Code2, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code3", obj.Hirerachy_Code3, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code4", obj.Hirerachy_Code4, True)
                clsCommon.AddColumnsForChange(coll, "AddChargeCode", obj.AddChargeCode, True)
                clsCommon.AddColumnsForChange(coll, "AddChargeDesc", obj.AddChargeDesc, True)
                clsCommon.AddColumnsForChange(coll, "Distributor_Commission_Amt", obj.Distributor_Commission_Amt, True)
                clsCommon.AddColumnsForChange(coll, "Promotional_GL_Account_Code", obj.Promotional_GL_Account_Code, True)
                clsCommon.AddColumnsForChange(coll, "Transporter_Commission_Amt", obj.Transporter_Commission_Amt, True)
                clsCommon.AddColumnsForChange(coll, "Transporter_GL_Account_Code", obj.Transporter_GL_Account_Code, True)
                clsCommon.AddColumnsForChange(coll, "Security_Amt", obj.Security_Amt, True)
                clsCommon.AddColumnsForChange(coll, "SD_GL_Account_Code", obj.SD_GL_Account_Code, True)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Customer_Invoice_Detail", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If
        Return True
    End Function



End Class

Public Class clsOPInvoiceForTCS

#Region "Variables"
    Public FINANCIAL_YEAR_CODE As String = Nothing
    Public CUSTOMER_CODE As String = Nothing
    Public SALE_AMT As Decimal = 0

#End Region
    Public Function SaveData(ByVal obj As clsOPInvoiceForTCS, ByVal isNewEntry As Boolean) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = obj.SaveData(obj, isNewEntry, trans)
            If (isSaved) Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function
    Public Function SaveData(ByVal obj As clsOPInvoiceForTCS, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "FINANCIAL_YEAR_CODE", obj.FINANCIAL_YEAR_CODE)
            clsCommon.AddColumnsForChange(coll, "CUSTOMER_CODE", obj.CUSTOMER_CODE)
            clsCommon.AddColumnsForChange(coll, "SALE_AMT", obj.SALE_AMT)

            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt"))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_OP_invoice_for_TCS", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_OP_invoice_for_TCS", OMInsertOrUpdate.Update, "CUSTOMER_CODE='" + obj.CUSTOMER_CODE + "'", trans)

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

End Class