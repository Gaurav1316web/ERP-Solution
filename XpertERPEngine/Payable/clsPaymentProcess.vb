Imports common
Imports System.Data.SqlClient
'Check in prabhakar 22/06/2020
Public Class clsPaymentProcessHead
#Region "Variables"
    Public Doc_No As String = ""
    Public Doc_Date As String = ""
    Public From_Date As String = ""
    Public To_Date As String = ""
    Public Loc_Seg_Code As String = ""
    Public MCC_Code_Selected As String = ""
    Public Created_By As String = ""
    Public Created_Date As String = ""
    Public Modified_By As String = ""
    Public Modified_Date As String = ""
    Public Comp_Code As String = ""
    Public DocRefNoForUploader As String = ""
    Public Area_Location_Code As String = ""
    Public Location_Code_Prefix As String = Nothing
    Public ArrPPDetail As List(Of clsPaymentProcessDetail) = Nothing
    Public arrClsPaymentProcessInvoices As List(Of clsPaymentProcessInvoices) = Nothing
    Public arrClsPaymentProcessMccSale As List(Of clsPaymentProcessMCCSale) = Nothing
    Public arrClsPaymentProcessMccSaleReturn As List(Of clsPaymentProcessMCCSaleReturn) = Nothing
    Public arrClsPaymentProcessItemIssue As List(Of clsPaymentProcessItemIssue) = Nothing
    Public arrClsPaymentProcessItemIssueReturn As List(Of clsPaymentProcessItemIssueReturn) = Nothing
    Public arrClsPaymentProcessDeductions As List(Of clsPaymentProcessDeduction) = Nothing
    Public arrClsPaymentProcessCreditNote As List(Of clsPaymentProcessCreditNote) = Nothing
    Public arrclsPaymentProcessSaving As List(Of clsPaymentProcessSaving) = Nothing
    Public arrclsPaymentProcessCompulsory As List(Of clsPaymentProcessCompulsory) = Nothing
    Public ArrPPAdvancePayment As List(Of clsPaymentProcessAdvancePayment) = Nothing
    Public ArrPPAssetLost As List(Of clsPaymentProcessAssetLost) = Nothing

    Public dtPPDetail As DataTable = Nothing
    Public dtClsPaymentProcessInvoices As DataTable = Nothing
    Public dtClsPaymentProcessMccSale As DataTable = Nothing
    Public dtClsPaymentProcessMccSaleReturn As DataTable = Nothing
    Public dtClsPaymentProcessItemIssue As DataTable = Nothing
    Public dtClsPaymentProcessItemIssueReturn As DataTable = Nothing
    Public dtClsPaymentProcessDeductions As DataTable = Nothing
    Public dtClsPaymentProcessCreditNote As DataTable = Nothing
    Public dtclsPaymentProcessSaving As DataTable = Nothing
    Public dtclsPaymentProcessCompulsory As DataTable = Nothing
    Public dtPPAdvancePayment As DataTable = Nothing
    Public dtPPAssetLost As DataTable = Nothing

    Public isPosted As Integer = 0
    Public isPrePosted As Integer = 0
    Dim Posting_Date As String = ""
    Public PaymentDesc As String = ""

    Public Is_Skip_Previous_Item_Issue As Boolean
    Public Is_Skip_Previous_Item_Issue_Return As Boolean
    Public Is_Skip_Previous_MCC_Sale As Boolean
    Public Is_Skip_Previous_MCC_Sale_Return As Boolean
    Public Is_Skip_Previous_Credit_Note As Boolean
    Public Is_Skip_Previous_Debit_Note As Boolean
    Public Is_Skip_Previous_Advacee_Payment As Boolean
    Public ArrPPSkipDoc As List(Of clsPaymentProcessSkipDoc) = Nothing
    Public FarmType As String = "PP"
    Dim SetCowFatPer As Decimal = 0

    Public ArrIndex As Dictionary(Of String, clsPaymentProcessIndex)
#End Region

    Public Shared Function SaveData(ByVal obj As clsPaymentProcessHead, ByVal isNewEntry As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(obj, trans, isNewEntry)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function SaveData(ByVal obj As clsPaymentProcessHead, ByVal trans As SqlTransaction, ByVal isNewEntry As Boolean) As Boolean
        Dim issaved As Boolean = True
        Try


            If isNewEntry Then
                If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.CheckUnpostedPaymentProcess, clsFixedParameterCode.CheckUnpostedPaymentProcess, trans)) > 0 Then
                    Dim qry As String = "select Doc_No from TSPL_PAYMENT_PROCESS_HEAD where Loc_Seg_Code='" + obj.Loc_Seg_Code + "' and isPosted='0' and Doc_No not in ('" + obj.Doc_No + "')"
                    If clsCommon.myLen(obj.MCC_Code_Selected) > 0 Then
                        qry += " and MCC_Code_Selected='" + obj.MCC_Code_Selected + "'"
                    End If

                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                    If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                        qry = "Unposted Payment process found " + Environment.NewLine
                        For Each dr As DataRow In dt1.Rows
                            qry += clsCommon.myCstr(dr("Doc_No")) + ","
                        Next
                        qry += Environment.NewLine + "First delete or post these documents"
                        Throw New Exception(qry)
                    End If
                End If
            End If

            'If Not isNewEntry Then
            '    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Doc_No, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)
            'End If

            clsPaymentProcessInvoices.deleteData(obj.Doc_No, trans)
            clsPaymentProcessMCCSale.deleteData(obj.Doc_No, trans)
            clsPaymentProcessMCCSaleReturn.deleteData(obj.Doc_No, trans)
            clsPaymentProcessItemIssue.deleteData(obj.Doc_No, trans)
            clsPaymentProcessItemIssueReturn.deleteData(obj.Doc_No, trans)
            clsPaymentProcessDeduction.deleteData(obj.Doc_No, trans)
            clsPaymentProcessCreditNote.deleteData(obj.Doc_No, trans)
            clsPaymentProcessSaving.deleteData(obj.Doc_No, trans)
            clsPaymentProcessCompulsory.deleteData(obj.Doc_No, trans)
            clsPaymentProcessDetail.deleteData(obj.Doc_No, trans)
            clsPaymentProcessAdvancePayment.deleteData(obj.Doc_No, trans)
            clsPaymentProcessAssetLost.deleteData(obj.Doc_No, trans)

            Dim dt As Date = clsCommon.myCDate(clsCommon.GetPrintDate(obj.Doc_Date, "dd/MMM/yyyy"))

            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, obj.Area_Location_Code, clsCommon.myCDate(obj.Doc_Date), trans)

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Area_Location_Code", obj.Area_Location_Code, True)
            clsCommon.AddColumnsForChange(coll, "FarmType", "PP")
            clsCommon.AddColumnsForChange(coll, "Doc_Date", clsCommon.GetPrintDate(obj.Doc_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "From_Date", clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "To_Date", clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Loc_Seg_Code", clsCommon.myCstr(obj.Loc_Seg_Code))
            clsCommon.AddColumnsForChange(coll, "Location_Code_Prefix", clsCommon.myCstr(obj.Location_Code_Prefix), True)
            clsCommon.AddColumnsForChange(coll, "PaymentDesc", clsCommon.myCstr(obj.PaymentDesc))
            clsCommon.AddColumnsForChange(coll, "MCC_Code_Selected", obj.MCC_Code_Selected, True)
            clsCommon.AddColumnsForChange(coll, "isPosted", 0)
            clsCommon.AddColumnsForChange(coll, "isPrePosted", 0)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "DocRefNoForUploader", clsCommon.myCstr(obj.DocRefNoForUploader))
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            'clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            'clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))

            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_Item_Issue", IIf(obj.Is_Skip_Previous_Item_Issue, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_Item_Issue_Return", IIf(obj.Is_Skip_Previous_Item_Issue_Return, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_MCC_Sale", IIf(obj.Is_Skip_Previous_MCC_Sale, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_MCC_Sale_Return", IIf(obj.Is_Skip_Previous_MCC_Sale_Return, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_Credit_Note", IIf(obj.Is_Skip_Previous_Credit_Note, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_Debit_Note", IIf(obj.Is_Skip_Previous_Debit_Note, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Skip_Previous_Advacee_Payment", IIf(obj.Is_Skip_Previous_Advacee_Payment, 1, 0))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))

                Dim MultipleFinderFillAuto As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MultipleFinderFillAuto, clsFixedParameterCode.MultipleFinderFillAuto, trans)) = 1)
                If objCommonVar.ApplyLocationWisePrefix Then
                    obj.Doc_No = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.PaymentProcess, "", obj.Location_Code_Prefix, False)
                Else
                    If MultipleFinderFillAuto Then
                        obj.Doc_No = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.PaymentProcess, "", "", False)
                    Else
                        obj.Doc_No = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.PaymentProcess, "", obj.Loc_Seg_Code, True)
                    End If
                End If


                'PaymentProcessWithoutLoc
                clsCommon.AddColumnsForChange(coll, "Doc_No", obj.Doc_No)
                If clsCommon.myLen(obj.Doc_No) <= 0 Then
                    Throw New Exception("Error In Doc No Generation")
                End If
                issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_HEAD", OMInsertOrUpdate.Insert, "", trans)
            Else
                issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_HEAD", OMInsertOrUpdate.Update, "Doc_No= '" + obj.Doc_No + "'", trans)
            End If


            issaved = issaved AndAlso clsPaymentProcessInvoices.SaveData(obj.Doc_No, obj.arrClsPaymentProcessInvoices, trans)
            issaved = issaved AndAlso clsPaymentProcessMCCSale.SaveData(obj.Doc_No, obj.arrClsPaymentProcessMccSale, trans)
            issaved = issaved AndAlso clsPaymentProcessMCCSaleReturn.SaveData(obj.Doc_No, obj.arrClsPaymentProcessMccSaleReturn, trans)
            issaved = issaved AndAlso clsPaymentProcessItemIssue.SaveData(obj.Doc_No, obj.arrClsPaymentProcessItemIssue, trans)
            issaved = issaved AndAlso clsPaymentProcessItemIssueReturn.SaveData(obj.Doc_No, obj.arrClsPaymentProcessItemIssueReturn, trans)
            issaved = issaved AndAlso clsPaymentProcessDeduction.SaveData(obj.Doc_No, obj.arrClsPaymentProcessDeductions, trans)
            issaved = issaved AndAlso clsPaymentProcessCreditNote.SaveData(obj.Doc_No, obj.arrClsPaymentProcessCreditNote, trans)
            issaved = issaved AndAlso clsPaymentProcessSaving.SaveData(obj.Doc_No, obj.arrclsPaymentProcessSaving, trans)
            issaved = issaved AndAlso clsPaymentProcessCompulsory.SaveData(obj.Doc_No, obj.arrclsPaymentProcessCompulsory, trans)
            issaved = issaved AndAlso clsPaymentProcessDetail.SaveData(obj.Doc_No, obj.ArrPPDetail, trans)
            issaved = issaved AndAlso clsPaymentProcessAssetLost.SaveData(obj.Doc_No, obj.ArrPPAssetLost, trans)
            issaved = issaved AndAlso clsPaymentProcessAdvancePayment.SaveData(obj.Doc_No, obj.ArrPPAdvancePayment, obj.ArrPPDetail, trans) ''It Should be at last
            issaved = issaved AndAlso clsPaymentProcessSkipDoc.SaveData(obj.Doc_No, obj.ArrPPSkipDoc, trans) 'It Should be at last
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Doc_No, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return issaved
    End Function

    Public Shared Function PrePostData(ByVal DocNo As String, ByVal arrSavingDocs As ArrayList) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim obj As clsPaymentProcessHead = clsPaymentProcessHead.getData(DocNo, NavigatorType.Current, trans, "", True, False, False)
            If obj.isPrePosted = 1 Then
                Throw New Exception("Transaction should be pending")
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + DocNo + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, clsCommon.myCstr(dt.Rows(0)("Area_Location_Code")), clsCommon.myCDate(dt.Rows(0)("Doc_Date")), trans)
            End If

            Dim qry As String = ""
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                Dim baseQry As String = "select TSPL_VENDOR_INVOICE_HEAD.Document_No,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,TSPL_VENDOR_INVOICE_HEAD.Posting_Date, TSPL_VENDOR_INVOICE_HEAD.Document_Total,Against_TransferToSavingPKID 
            from TSPL_PAYMENT_PROCESS_SAVING 
            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
            where 2=2 and Doc_No='" + obj.Doc_No + "' 
            and TSPL_VENDOR_INVOICE_HEAD.Against_TransferToSavingPKID is not null"

                qry = "select xxx.*,TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,TSPL_PAYMENT_PROCESS_DETAIL.PP_Detail_No from (
            select Vendor_Code,min(Against_TransferToSavingPKID) as Min_Against_TransferToSavingPKID,max(Against_TransferToSavingPKID) as Max_Against_TransferToSavingPKID,sum(Document_Total) as Document_Total from (" + baseQry + ")xx group by Vendor_Code 
            )xxx
            left outer join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" + obj.Doc_No + "'  and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE=xxx.Vendor_Code 
            where xxx.Document_Total>TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount "
                dt = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        If clsCommon.myCDecimal(dr("Min_Against_TransferToSavingPKID")) = clsCommon.myCDecimal(dr("Max_Against_TransferToSavingPKID")) Then
                            UpdateTransferToSaving(clsCommon.myCDecimal(dr("Min_Against_TransferToSavingPKID")), clsCommon.myCDecimal(dr("Payable_Amount")), trans, clsCommon.myCstr(dr("PP_Detail_No")))
                        Else
                            qry = baseQry + " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code='" + clsCommon.myCstr(dr("Vendor_Code")) + "' order by TSPL_VENDOR_INVOICE_HEAD.Document_Total desc"
                            Dim dtInner As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                            If dtInner IsNot Nothing AndAlso dtInner.Rows.Count > 0 Then
                                Dim BalAmt As Decimal = clsCommon.myCDecimal(dr("Document_Total")) - clsCommon.myCDecimal(dr("Payable_Amount"))
                                For Each drInner As DataRow In dtInner.Rows
                                    If BalAmt > clsCommon.myCDecimal(drInner("Document_Total")) Then
                                        UpdateTransferToSaving(clsCommon.myCDecimal(drInner("Against_TransferToSavingPKID")), 0, trans, clsCommon.myCstr(dr("PP_Detail_No")))
                                        BalAmt -= clsCommon.myCDecimal(drInner("Document_Total"))
                                    Else
                                        BalAmt -= clsCommon.myCDecimal(drInner("Document_Total"))
                                        UpdateTransferToSaving(clsCommon.myCDecimal(drInner("Against_TransferToSavingPKID")), Math.Abs(BalAmt), trans, clsCommon.myCstr(dr("PP_Detail_No")))
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If

            If arrSavingDocs IsNot Nothing AndAlso arrSavingDocs.Count > 0 Then
                For Each str As String In arrSavingDocs
                    clsVedorInvoiceHead.ReverseAndUnpost(str, trans)
                    clsVedorInvoiceHead.DeleteData(str, trans)
                Next
            End If

            qry = " update TSPL_PAYMENT_PROCESS_HEAD set isPrePosted=1, Posting_Date='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy") & "' where doc_no='" & obj.Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, DocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Shared Sub UpdateTransferToSaving(intPKID As Integer, NewAmount As Decimal, trans As SqlTransaction, strPP_Detail_No As String)
        Dim qry As String = "update TSPL_TRANSFER_TO_SAVING_DETAIL set Org_Amount=Amount where PK_ID='" + clsCommon.myCstr(intPKID) + "' and Org_Amount is null"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "update TSPL_TRANSFER_TO_SAVING_DETAIL set Amount=" + clsCommon.myCstr(NewAmount) + " where PK_ID='" + clsCommon.myCstr(intPKID) + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "select TSPL_VENDOR_INVOICE_HEAD.Document_No,TSPL_JOURNAL_MASTER.Voucher_No 
from TSPL_VENDOR_INVOICE_HEAD 
left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No=TSPL_VENDOR_INVOICE_HEAD.Document_No 
where Against_TransferToSavingPKID=" + clsCommon.myCstr(intPKID) + ""
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        For Each dr As DataRow In dt.Rows
            qry = "update TSPL_VENDOR_INVOICE_HEAD set Document_Total=" + clsCommon.myCstr(NewAmount) + "
,Balance_Amt=" + clsCommon.myCstr(NewAmount) + "
,Amount_Less_Discount=" + clsCommon.myCstr(NewAmount) + "
,Discount_Base=" + clsCommon.myCstr(NewAmount) + "
where Document_No='" + clsCommon.myCstr(dr("Document_No")) + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "update TSPL_VENDOR_INVOICE_DETAIL set Amount=" + clsCommon.myCstr(NewAmount) + "
,Amount_less_Discount=" + clsCommon.myCstr(NewAmount) + "
,Total_Amount=" + clsCommon.myCstr(NewAmount) + "
where Document_No='" + clsCommon.myCstr(dr("Document_No")) + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim objVI As clsVedorInvoiceHead = clsVedorInvoiceHead.GetData(clsCommon.myCstr(dr("Document_No")), "", trans)
            clsVedorInvoiceHead.CreateJournalEntry(objVI, "", objVI.Posting_Date, trans)
        Next


        qry = "update TSPL_PAYMENT_PROCESS_DETAIL set Saving_Amount=NewSaving_Amount from (
select TSPL_PAYMENT_PROCESS_DETAIL.PP_Detail_No, TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount, TSPL_VENDOR_INVOICE_HEAD.Document_Total as NewSaving_Amount from TSPL_PAYMENT_PROCESS_DETAIL
inner join TSPL_PAYMENT_PROCESS_SAVING on TSPL_PAYMENT_PROCESS_SAVING.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No  
inner join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
where TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code and TSPL_PAYMENT_PROCESS_DETAIL.PP_Detail_No='" + strPP_Detail_No + "'
 ) xx inner join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.PP_Detail_No=xx.PP_Detail_No"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

    End Sub

    Public Shared Function ProcessData(ByVal DocNo As String, ByVal Desc As String) As Boolean
        Return ProcessData(DocNo, Desc, True)
    End Function

    Public Shared Function ProcessData(ByVal DocNo As String, ByVal Desc As String, ByVal ShowProgressBAR As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ProcessData(DocNo, Desc, ShowProgressBAR, trans)
            'Throw New Exception("Balwinder singh premi")
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function ProcessData(ByVal DocNo As String, ByVal Desc As String, ByVal ShowProgressBAR As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim obj As clsPaymentProcessHead = clsPaymentProcessHead.getData(DocNo, NavigatorType.Current, trans, "", False, True)
        Dim i As Integer = 0
        Dim Counter As Integer = 0
        Dim AdjAmt As Decimal = 0
        Dim ReturnAdjAmt As Decimal = 0

        Dim todaydt As Date = clsCommon.GETSERVERDATE(trans)
        Dim objRcpt As clsAdjustmentEntryReceivables = Nothing
        Dim objPayAdj As clsPaymentAdjustmentEntry = Nothing
        Dim DisCCodeForArAdj As String = ""
        Dim GLAcARAdj As String = ""
        Dim DiscDiscForArAdj As String = ""
        Dim GLAcDescARAdj As String = ""
        Dim objPay As clsPaymentHeader = Nothing
        Dim objTr As New clsPaymentDetail()
        Dim isProgressBarShownLocal As Boolean = False
        Dim arrARForSameMCCSaleAndReturn As New Dictionary(Of String, Decimal)
        Try
            If obj.ArrPPDetail IsNot Nothing And obj.ArrPPDetail.Count > 0 Then
                If obj.isPrePosted = 0 Then
                    Throw New Exception("Transaction should be Pre posted")
                End If
                DisCCodeForArAdj = clsFixedParameter.GetData(clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, trans)
                If clsCommon.myLen(DisCCodeForArAdj) <= 0 Then
                    Throw New Exception("Please Map Discount code from Sale setting")
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select  Code,Description,Account_Code,Account_Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    Throw New Exception("Discount Code not Exists [" + DisCCodeForArAdj + "] as Defined is Fixed Parameter [" + clsFixedParameterType.DiscountCodeForArAdj + "]")
                End If
                DiscDiscForArAdj = clsCommon.myCstr(dt.Rows(0)("Description"))
                GLAcARAdj = clsCommon.myCstr(dt.Rows(0)("Account_Code"))
                If clsCommon.myLen(GLAcARAdj) <= 0 Then
                    Throw New Exception("Please set account for Discount master:" + DisCCodeForArAdj)
                End If
                GLAcDescARAdj = clsCommon.myCstr(dt.Rows(0)("Account_Description"))
                If obj.arrClsPaymentProcessMccSale IsNot Nothing AndAlso obj.arrClsPaymentProcessMccSale.Count > 0 Then
                    If ShowProgressBAR Then
                        clsCommon.ProgressBarPercentShow()
                        isProgressBarShownLocal = True
                    End If
                    For i = 0 To obj.arrClsPaymentProcessMccSale.Count - 1
                        If ShowProgressBAR Then
                            clsCommon.ProgressBarPercentUpdate(i * 100 / obj.arrClsPaymentProcessMccSale.Count, "Updating AR Adjustment Record " & (i + 1) & " Of " & obj.arrClsPaymentProcessMccSale.Count)
                        End If
                        objRcpt = New clsAdjustmentEntryReceivables
                        objRcpt.Adjustment_No = ""
                        objRcpt.Description = " Adjustment Against Bulk Payment Process "
                        objRcpt.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                        objRcpt.Customer_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Cust_Code  from TSPL_CUSTOMER_VENDOR_MAPPING where Vendor_Code ='" & obj.arrClsPaymentProcessMccSale(i).Customer_CODE & "' ", trans))
                        objRcpt.Customer_Name = clsCommon.myCstr(clsCustomerMaster.GetName(objRcpt.Customer_No, trans))
                        objRcpt.Doc_No = clsCommon.myCstr(obj.arrClsPaymentProcessMccSale(i).Sale_Doc_No)
                        Dim ReturnAmt As Decimal = 0
                        For ireturn As Integer = 0 To obj.arrClsPaymentProcessMccSaleReturn.Count - 1
                            If clsCommon.myCstr(obj.arrClsPaymentProcessMccSale(i).Sale_Doc_No) = clsCommon.myCstr(obj.arrClsPaymentProcessMccSaleReturn(ireturn).Sale_Doc_No) Then
                                ReturnAmt += obj.arrClsPaymentProcessMccSaleReturn(ireturn).Amount
                                If arrARForSameMCCSaleAndReturn.ContainsKey(obj.arrClsPaymentProcessMccSale(i).AR_Invoice_No) Then
                                    arrARForSameMCCSaleAndReturn.Remove(obj.arrClsPaymentProcessMccSale(i).AR_Invoice_No)
                                End If
                                If arrARForSameMCCSaleAndReturn.ContainsKey(obj.arrClsPaymentProcessMccSaleReturn(ireturn).AR_Invoice_No) Then
                                    arrARForSameMCCSaleAndReturn.Remove(obj.arrClsPaymentProcessMccSaleReturn(ireturn).AR_Invoice_No)
                                End If
                                arrARForSameMCCSaleAndReturn.Add(obj.arrClsPaymentProcessMccSale(i).AR_Invoice_No, obj.arrClsPaymentProcessMccSaleReturn(ireturn).Amount)
                                arrARForSameMCCSaleAndReturn.Add(obj.arrClsPaymentProcessMccSaleReturn(ireturn).AR_Invoice_No, obj.arrClsPaymentProcessMccSaleReturn(ireturn).Amount)
                            End If
                        Next
                        objRcpt.ARInvoiceNo = clsCommon.myCstr(obj.arrClsPaymentProcessMccSale(i).AR_Invoice_No)
                        objRcpt.Doc_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSale(i).Amount) - ReturnAmt
                        objRcpt.Remarks = ""
                        objRcpt.Adjustment_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSale(i).Amount - obj.arrClsPaymentProcessMccSale(i).Reduce_Deduc_Amt) - ReturnAmt
                        objRcpt.Arr = New List(Of clsAdjustmentEntryReceivablesDetail)
                        Dim objTrRcpt As New clsAdjustmentEntryReceivablesDetail()
                        objTrRcpt.Discount_Code = DisCCodeForArAdj
                        objTrRcpt.Discount_Description = DiscDiscForArAdj


                        objTrRcpt.Account_No = GLAcARAdj
                        objTrRcpt.Account_Description = GLAcDescARAdj
                        objTrRcpt.Amount = clsCommon.myCDecimal(clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSale(i).Amount - obj.arrClsPaymentProcessMccSale(i).Reduce_Deduc_Amt - ReturnAmt))
                        objTrRcpt.Remarks = ""
                        objRcpt.Arr.Add(objTrRcpt)
                        If clsCommon.myCDecimal(objTrRcpt.Amount) > 0 Then
                            objRcpt.SaveData(objRcpt, True, trans)
                            clsAdjustmentEntryReceivables.FunPost(objRcpt.Adjustment_No, trans)
                        End If
                    Next

                    If ShowProgressBAR Then
                        isProgressBarShownLocal = False
                        clsCommon.ProgressBarPercentHide()
                    End If
                End If
                If ShowProgressBAR Then
                    clsCommon.ProgressBarPercentShow()
                    isProgressBarShownLocal = True
                End If


                For i = 0 To obj.ArrPPDetail.Count - 1
                    If clsCommon.CompairString(obj.ArrPPDetail(i).VSP_CODE, "DCS00687") = CompairStringResult.Equal Then
                        Dim x As Integer = 0
                    End If
                    If ShowProgressBAR Then
                        clsCommon.ProgressBarPercentUpdate(i * 100 / obj.ArrPPDetail.Count, " Creating Payment Entry  Record " & (i + 1) & " Of " & obj.ArrPPDetail.Count)
                    End If
                    If Not obj.ArrPPDetail(i).Is_select Then
                        Continue For
                    End If
                    ''Becuause of Knock off amount will always be create.
                    CreateApplyDocumentForAdavancePayment(clsCommon.myCDate(obj.Doc_Date), obj.ArrPPAdvancePayment, obj.ArrPPDetail(i), obj.Loc_Seg_Code, trans)
                    ReturnAdjAmt = 0
                    AdjAmt = 0
                    For Counter = 0 To obj.arrClsPaymentProcessMccSale.Count - 1
                        If clsCommon.CompairString(obj.arrClsPaymentProcessMccSale(Counter).Customer_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                            AdjAmt = AdjAmt + (obj.arrClsPaymentProcessMccSale(Counter).Amount - obj.arrClsPaymentProcessMccSale(Counter).Reduce_Deduc_Amt)
                            Dim ReturnAmt As Decimal = 0
                            For ireturn As Integer = 0 To obj.arrClsPaymentProcessMccSaleReturn.Count - 1
                                If clsCommon.myCstr(obj.arrClsPaymentProcessMccSale(Counter).Sale_Doc_No) = clsCommon.myCstr(obj.arrClsPaymentProcessMccSaleReturn(ireturn).Sale_Doc_No) Then
                                    ReturnAmt += clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSaleReturn(ireturn).Amount)
                                    obj.arrClsPaymentProcessMccSaleReturn(ireturn).Amount = 0
                                End If
                            Next
                            AdjAmt -= ReturnAmt
                        End If
                    Next
                    Dim arrCreditNoteAdjustAmt As New Dictionary(Of String, Decimal)
                    If AdjAmt > 0 Then
                        Dim objTrPay As clsPaymentAdjustmentEntryDetail = Nothing
                        If AdjAmt > obj.ArrPPDetail(i).Total_Invoice_Amount Then ''ERO/03/09/19-001016 by blaiwnder on 05/09/2019
                            Dim AmtToAdjustInCreditNote As Decimal = AdjAmt - obj.ArrPPDetail(i).Total_Invoice_Amount
                            Dim dclRedDeduction As Decimal = 0
                            For Counter = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRTo
                                If clsCommon.CompairString(obj.arrClsPaymentProcessDeductions(Counter).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                    dclRedDeduction += obj.arrClsPaymentProcessDeductions(Counter).Reduce_Deduc_Amt
                                End If
                            Next

                            If ((obj.ArrPPDetail(i).Credit_Note_Amount + obj.ArrPPDetail(i).Head_Load_Amount + obj.ArrPPDetail(i).Compulsory_Amount + dclRedDeduction) - obj.ArrPPDetail(i).Deduction_Amount - obj.ArrPPDetail(i).Advance_Payment_Amount) >= AmtToAdjustInCreditNote Then

                                If obj.arrClsPaymentProcessCreditNote IsNot Nothing And obj.arrClsPaymentProcessCreditNote.Count > 0 Then
                                    For k As Integer = 0 To obj.arrClsPaymentProcessCreditNote.Count - 1
                                        If clsCommon.CompairString(obj.arrClsPaymentProcessCreditNote(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                            Dim tAmt As Decimal = 0
                                            If AmtToAdjustInCreditNote >= obj.arrClsPaymentProcessCreditNote.Item(k).Amount Then
                                                tAmt = obj.arrClsPaymentProcessCreditNote.Item(k).Amount
                                                AmtToAdjustInCreditNote -= tAmt
                                            Else
                                                tAmt = AmtToAdjustInCreditNote
                                                AmtToAdjustInCreditNote = 0
                                            End If
                                            AdjAmt -= tAmt

                                            objPayAdj = New clsPaymentAdjustmentEntry
                                            objPayAdj.Adjustment_No = ""
                                            objPayAdj.Description = " AP Adjustment Against Bulk Payment Process"
                                            objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                                            objPayAdj.Vendor_No = obj.ArrPPDetail(i).VSP_CODE
                                            objPayAdj.Vendor_Name = obj.ArrPPDetail(i).VSP_NAME
                                            objPayAdj.Doc_No = obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No
                                            objPayAdj.Doc_Amount = obj.arrClsPaymentProcessCreditNote.Item(k).Amount
                                            objPayAdj.Remarks = "Credit note Adjusment"
                                            objPayAdj.Adjustment_Amount = clsCommon.myCDecimal(tAmt)
                                            objPayAdj.Against_Payment_Process = DocNo
                                            objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                                            objTrPay = New clsPaymentAdjustmentEntryDetail()
                                            objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                                            objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                                            objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                                            objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                                            objTrPay.Amount = clsCommon.myCDecimal(tAmt)
                                            objTrPay.Remarks = "Credit note Adjusment"
                                            objPayAdj.Arr.Add(objTrPay)
                                            objPayAdj.SaveData(objPayAdj, True, trans)
                                            clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                                            arrCreditNoteAdjustAmt.Add(obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No, tAmt)
                                            If AmtToAdjustInCreditNote = 0 Then
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If

                                If AmtToAdjustInCreditNote <> 0 Then
                                    If obj.arrclsPaymentProcessCompulsory IsNot Nothing And obj.arrclsPaymentProcessCompulsory.Count > 0 Then
                                        For k As Integer = 0 To obj.arrclsPaymentProcessCompulsory.Count - 1
                                            If clsCommon.CompairString(obj.arrclsPaymentProcessCompulsory(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                                Dim tAmt As Decimal = 0
                                                If AmtToAdjustInCreditNote >= obj.arrclsPaymentProcessCompulsory.Item(k).Amount Then
                                                    tAmt = obj.arrclsPaymentProcessCompulsory.Item(k).Amount
                                                    AmtToAdjustInCreditNote -= tAmt
                                                Else
                                                    tAmt = AmtToAdjustInCreditNote
                                                    AmtToAdjustInCreditNote = 0
                                                End If
                                                AdjAmt -= tAmt

                                                objPayAdj = New clsPaymentAdjustmentEntry
                                                objPayAdj.Adjustment_No = ""
                                                objPayAdj.Description = " AP Adjustment Against Bulk Payment Process"
                                                objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                                                objPayAdj.Vendor_No = obj.ArrPPDetail(i).VSP_CODE
                                                objPayAdj.Vendor_Name = obj.ArrPPDetail(i).VSP_NAME
                                                objPayAdj.Doc_No = obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No
                                                objPayAdj.Doc_Amount = obj.arrclsPaymentProcessCompulsory.Item(k).Amount
                                                objPayAdj.Remarks = "Credit note Adjusment"
                                                objPayAdj.Adjustment_Amount = clsCommon.myCDecimal(tAmt)
                                                objPayAdj.Against_Payment_Process = DocNo
                                                objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                                                objTrPay = New clsPaymentAdjustmentEntryDetail()
                                                objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                                                objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                                                objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                                                objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                                                objTrPay.Amount = clsCommon.myCDecimal(tAmt)
                                                objTrPay.Remarks = "Credit note Adjusment"
                                                objPayAdj.Arr.Add(objTrPay)
                                                objPayAdj.SaveData(objPayAdj, True, trans)
                                                clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                                                arrCreditNoteAdjustAmt.Add(obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No, tAmt)
                                                If AmtToAdjustInCreditNote = 0 Then
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                End If

                                If AmtToAdjustInCreditNote <> 0 Then
                                    If obj.ArrPPDetail(i).Head_Load_Amount > 0 Then
                                        Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from TSPL_VENDOR_INVOICE_HEAD where Document_Type='C' and RefDocType='Milk_HE' and Against_MillkPurchaseInvoice_No='" + obj.ArrPPDetail(i).Milk_Purchase_Invoice_No + "'", trans))
                                        If clsCommon.myLen(strDocNo) > 0 Then
                                            Dim tAmt As Decimal = 0
                                            If AmtToAdjustInCreditNote >= obj.ArrPPDetail(i).Head_Load_Amount Then
                                                tAmt = obj.ArrPPDetail(i).Head_Load_Amount
                                                AmtToAdjustInCreditNote -= tAmt
                                            Else
                                                tAmt = AmtToAdjustInCreditNote
                                                AmtToAdjustInCreditNote = 0
                                            End If
                                            AdjAmt -= tAmt

                                            objPayAdj = New clsPaymentAdjustmentEntry
                                            objPayAdj.Adjustment_No = ""
                                            objPayAdj.Description = " AP Adjustment Against Bulk Payment Process Head Load"
                                            objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                                            objPayAdj.Vendor_No = obj.ArrPPDetail(i).VSP_CODE
                                            objPayAdj.Vendor_Name = obj.ArrPPDetail(i).VSP_NAME
                                            objPayAdj.Doc_No = strDocNo
                                            objPayAdj.Doc_Amount = obj.ArrPPDetail(i).Head_Load_Amount
                                            objPayAdj.Remarks = "Credit note Adjusment Head Load"
                                            objPayAdj.Adjustment_Amount = clsCommon.myCDecimal(tAmt)
                                            objPayAdj.Against_Payment_Process = DocNo
                                            objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                                            objTrPay = New clsPaymentAdjustmentEntryDetail()
                                            objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                                            objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                                            objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                                            objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                                            objTrPay.Amount = clsCommon.myCDecimal(tAmt)
                                            objTrPay.Remarks = "Credit note Adjusment Head Load"
                                            objPayAdj.Arr.Add(objTrPay)
                                            objPayAdj.SaveData(objPayAdj, True, trans)
                                            clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                                            arrCreditNoteAdjustAmt.Add(DocNo, tAmt)
                                            'If AmtToAdjustInCreditNote = 0 Then
                                            '    Exit For
                                            'End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ''AdjAmt and balance amt is having gap of TDS Amount.
                        objPayAdj = New clsPaymentAdjustmentEntry
                        objPayAdj.Adjustment_No = ""
                        objPayAdj.Description = " AP Adjustment Against Bulk Payment Process "
                        objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                        objPayAdj.Vendor_No = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_CODE)
                        objPayAdj.Vendor_Name = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_NAME)
                        objPayAdj.Doc_No = clsCommon.myCstr(obj.ArrPPDetail(i).AP_Invoice_No)
                        objPayAdj.Doc_Amount = clsCommon.myCDecimal(obj.ArrPPDetail(i).Total_Invoice_Amount)
                        objPayAdj.Remarks = clsCommon.myCstr("")
                        objPayAdj.Adjustment_Amount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select balance_amt from TSPL_VENDOR_INVOICE_HEAD where Document_No='" + obj.ArrPPDetail(i).AP_Invoice_No + "'", trans))
                        objPayAdj.Against_Payment_Process = DocNo
                        objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                        objTrPay = New clsPaymentAdjustmentEntryDetail()
                        objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                        objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                        objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                        objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                        objTrPay.Amount = objPayAdj.Adjustment_Amount

                        objTrPay.Remarks = clsCommon.myCstr("")
                        objPayAdj.Arr.Add(objTrPay)
                        objPayAdj.SaveData(objPayAdj, True, trans)
                        clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                        If ShowProgressBAR Then
                            isProgressBarShownLocal = False
                            clsCommon.ProgressBarPercentHide()
                        End If
                    End If

                    AdjAmt -= (obj.ArrPPDetail(i).Credit_Note_Amount + obj.ArrPPDetail(i).Compulsory_Amount + obj.ArrPPDetail(i).Head_Load_Amount) ''Becuase not effet the ap adjustment.
                    '========BM00000007337===================
                    For Counter = 0 To obj.arrClsPaymentProcessMccSaleReturn.Count - 1
                        If clsCommon.CompairString(obj.arrClsPaymentProcessMccSaleReturn(Counter).Customer_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                            ReturnAdjAmt = ReturnAdjAmt + (obj.arrClsPaymentProcessMccSaleReturn(Counter).Amount)
                        End If
                    Next
                    If ReturnAdjAmt > 0 Then
                        objPayAdj = New clsPaymentAdjustmentEntry
                        objPayAdj.Adjustment_No = ""
                        objPayAdj.Description = " AP Return Adjustment Against Bulk Payment Process "
                        objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                        objPayAdj.Vendor_No = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_CODE)
                        objPayAdj.Vendor_Name = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_NAME)
                        objPayAdj.Doc_No = clsCommon.myCstr(obj.ArrPPDetail(i).AP_Invoice_No)
                        objPayAdj.Doc_Amount = clsCommon.myCDecimal(obj.ArrPPDetail(i).Total_Invoice_Amount)
                        objPayAdj.Remarks = clsCommon.myCstr("")
                        objPayAdj.Adjustment_Amount = clsCommon.myCDecimal(ReturnAdjAmt)
                        objPayAdj.Adjust_Type = "R"
                        objPayAdj.Against_Payment_Process = DocNo
                        objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                        Dim objTrPay As New clsPaymentAdjustmentEntryDetail()
                        objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                        objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                        objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                        objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                        objTrPay.Amount = clsCommon.myCDecimal(ReturnAdjAmt)
                        objTrPay.Remarks = clsCommon.myCstr("")
                        objPayAdj.Arr.Add(objTrPay)
                        objPayAdj.SaveData(objPayAdj, True, trans)
                        clsPaymentAdjustmentEntry.FunPostReverseEntry(objPayAdj.Adjustment_No, trans)
                    End If

                    If Not obj.ArrPPDetail(i).is_Hold_Payment_Process_Saving AndAlso obj.ArrPPDetail(i).Saving_Amount > 0 Then
                        If clsCommon.myLen(obj.ArrPPDetail.Item(i).Bank_Code_Saving) <= 0 Then
                            Throw New Exception("Saving account of DCS[" + obj.ArrPPDetail.Item(i).VSP_CODE + "] not found")
                        End If
                        objPay = New clsPaymentHeader()
                        objPay.Against_PP_Detail_No = obj.ArrPPDetail(i).PP_Detail_No
                        objPay.Payment_No = ""
                        objPay.Entry_Desc = "Saving " + Desc + " " + DocNo
                        objPay.Payment_Date = clsCommon.myCDate(obj.Doc_Date)
                        objPay.Payment_Post_Date = clsCommon.myCDate(obj.Doc_Date)
                        objPay.Bank_Code = obj.ArrPPDetail.Item(i).Bank_Code_Saving
                        objPay.Payment_Type = "PY"
                        objPay.Vendor_Code = obj.ArrPPDetail.Item(i).VSP_CODE
                        objPay.Vendor_Name = obj.ArrPPDetail.Item(i).VSP_NAME
                        objPay.Payment_Code = obj.ArrPPDetail.Item(i).Payment_Mode_Saving
                        'objPay.Cheque_No = obj.ArrPPDetail.Item(i).Cheque_No
                        'If Not obj.ArrPPDetail.Item(i).Cheque_Dated Is Nothing Then
                        '    objPay.Cheque_Date = obj.ArrPPDetail.Item(i).Cheque_Dated
                        'End If

                        objPay.Account_Payee = 0
                        objPay.memorndmamt = "0"
                        objPay.Applied_Payment = clsCommon.myCstr(obj.ArrPPDetail.Item(i).AP_Invoice_No)
                        objPay.Is_Security = 0
                        objPay.IsChkReverse = "N"
                        objPay.Bank_Charges = 0
                        objPay.Saving = True
                        objPay.ArrTr = New List(Of clsPaymentDetail)
                        If obj.arrclsPaymentProcessSaving IsNot Nothing And obj.arrclsPaymentProcessSaving.Count > 0 Then
                            For k As Integer = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).SavingFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).SavingTo
                                If clsCommon.CompairString(obj.arrclsPaymentProcessSaving(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                    Dim tAmt As Decimal = clsCommon.myCDecimal(obj.arrclsPaymentProcessSaving.Item(k).Amount)
                                    If tAmt > 0 Then
                                        objTr = New clsPaymentDetail()
                                        objTr.Apply = "1"
                                        objTr.Payment_Type = "PY"
                                        objTr.Document_No = clsCommon.myCstr(obj.arrclsPaymentProcessSaving.Item(k).AP_Invoice_No)
                                        objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrclsPaymentProcessSaving.Item(k).Amount)
                                        objTr.Applied_Amount = tAmt
                                        objTr.Pending_Balance = 0
                                        Dim vendorInvNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & obj.arrclsPaymentProcessSaving.Item(k).AP_Invoice_No & "'", trans))
                                        objTr.Net_Balance = 0
                                        objTr.Vendor_Invoice_No = vendorInvNo
                                        objTr.Security_Amount = 0
                                        objPay.ArrTr.Add(objTr)
                                    End If
                                End If
                            Next
                        End If
                        objPay.Payment_Amount = obj.ArrPPDetail(i).Saving_Amount
                        objPay.Balance_Amt = obj.ArrPPDetail(i).Saving_Amount
                        objPay.Location_Code = clsCommon.myCstr(obj.Loc_Seg_Code)
                        objPay.Entry_Desc = obj.PaymentDesc + " " + DocNo

                        objPay.SaveData(objPay, True, trans, True)
                        clsPaymentHeader.PostData(objPay.Payment_No, "Payable", trans)
                    End If

                    If obj.ArrPPDetail(i).is_Hold_Payment_Process Then
                        Dim XTotalAmount As Decimal = 0
                        For Counter = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRTo
                            If clsCommon.CompairString(obj.arrClsPaymentProcessDeductions(Counter).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                Dim XAmount As Decimal = obj.arrClsPaymentProcessDeductions(Counter).Amount - obj.arrClsPaymentProcessDeductions(Counter).Reduce_Deduc_Amt
                                If XAmount > 0 Then
                                    XTotalAmount += XAmount
                                    objPayAdj = New clsPaymentAdjustmentEntry
                                    objPayAdj.Adjustment_No = "" ''To Be Generated
                                    objPayAdj.Description = "AP Debit Note Adjustment Against Hold Payment Process"
                                    objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                                    objPayAdj.Vendor_No = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_CODE)
                                    objPayAdj.Vendor_Name = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_NAME)
                                    objPayAdj.Doc_No = clsCommon.myCstr(obj.arrClsPaymentProcessDeductions(Counter).AP_Invoice_No)
                                    objPayAdj.Doc_Amount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Document_Total from TSPL_Vendor_Invoice_Head where Document_No='" + obj.arrClsPaymentProcessDeductions(Counter).AP_Invoice_No + "'", trans))
                                    objPayAdj.Remarks = clsCommon.myCstr("")
                                    objPayAdj.Adjustment_Amount = XAmount
                                    objPayAdj.Against_Payment_Process = DocNo
                                    objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                                    Dim objTrPay As New clsPaymentAdjustmentEntryDetail()
                                    objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                                    objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                                    objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                                    objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                                    objTrPay.Amount = XAmount
                                    objTrPay.Remarks = clsCommon.myCstr("")
                                    objPayAdj.Arr.Add(objTrPay)
                                    objPayAdj.SaveData(objPayAdj, True, trans)
                                    clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                                End If
                            End If
                        Next
                        If XTotalAmount > 0 Then
                            For Counter = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).CRFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).CRTo
                                If clsCommon.CompairString(obj.arrClsPaymentProcessCreditNote(Counter).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                    'Dim XAmount As Decimal = obj.arrClsPaymentProcessCreditNote(Counter).Amount
                                    Dim XAmount As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Balance_Amt from TSPL_Vendor_Invoice_Head where Document_No='" + obj.arrClsPaymentProcessCreditNote(Counter).AP_Invoice_No + "'", trans))
                                    If XAmount > 0 Then
                                        XTotalAmount -= XAmount
                                        objPayAdj = New clsPaymentAdjustmentEntry
                                        objPayAdj.Adjustment_No = "" ''To Be Generated
                                        objPayAdj.Description = "AP Credit Note Adjustment Against Hold Payment Process"
                                        objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                                        objPayAdj.Vendor_No = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_CODE)
                                        objPayAdj.Vendor_Name = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_NAME)
                                        objPayAdj.Doc_No = clsCommon.myCstr(obj.arrClsPaymentProcessCreditNote(Counter).AP_Invoice_No)
                                        objPayAdj.Doc_Amount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Document_Total from TSPL_Vendor_Invoice_Head where Document_No='" + obj.arrClsPaymentProcessCreditNote(Counter).AP_Invoice_No + "'", trans))
                                        objPayAdj.Remarks = clsCommon.myCstr("")
                                        objPayAdj.Adjustment_Amount = XAmount
                                        objPayAdj.Against_Payment_Process = DocNo
                                        objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                                        Dim objTrPay As New clsPaymentAdjustmentEntryDetail()
                                        objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                                        objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                                        objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                                        objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                                        objTrPay.Amount = XAmount
                                        objTrPay.Remarks = clsCommon.myCstr("")
                                        objPayAdj.Arr.Add(objTrPay)
                                        objPayAdj.SaveData(objPayAdj, True, trans)
                                        clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                                    End If
                                End If
                            Next
                        End If
                        XTotalAmount = XTotalAmount - obj.ArrPPDetail(i).Head_Load_Amount
                        If XTotalAmount > 0 Then
                            objPayAdj = New clsPaymentAdjustmentEntry
                            objPayAdj.Adjustment_No = "" ''To Be Generated
                            objPayAdj.Description = "AP Invoice Adjustment Against Hold Payment Process"
                            objPayAdj.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                            objPayAdj.Vendor_No = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_CODE)
                            objPayAdj.Vendor_Name = clsCommon.myCstr(obj.ArrPPDetail(i).VSP_NAME)
                            objPayAdj.Doc_No = obj.ArrPPDetail(i).AP_Invoice_No
                            objPayAdj.Doc_Amount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Document_Total from TSPL_Vendor_Invoice_Head where Document_No='" + obj.ArrPPDetail(i).AP_Invoice_No + "'", trans))
                            objPayAdj.Remarks = clsCommon.myCstr("")
                            objPayAdj.Adjustment_Amount = XTotalAmount
                            objPayAdj.Against_Payment_Process = DocNo
                            objPayAdj.Arr = New List(Of clsPaymentAdjustmentEntryDetail)
                            Dim objTrPay As New clsPaymentAdjustmentEntryDetail()
                            objTrPay.Discount_Code = clsCommon.myCstr(DisCCodeForArAdj)
                            objTrPay.Discount_Description = clsCommon.myCstr(DiscDiscForArAdj)
                            objTrPay.Account_No = clsCommon.myCstr(GLAcARAdj)
                            objTrPay.Account_Description = clsCommon.myCstr(GLAcDescARAdj)
                            objTrPay.Amount = XTotalAmount
                            objTrPay.Remarks = clsCommon.myCstr("")
                            objPayAdj.Arr.Add(objTrPay)
                            objPayAdj.SaveData(objPayAdj, True, trans)
                            clsPaymentAdjustmentEntry.FunPost(objPayAdj.Adjustment_No, trans)
                        End If
                        Continue For
                    End If


                    ''------------------Payment Entry Start Here
                    objPay = New clsPaymentHeader()
                    objPay.Against_PP_Detail_No = obj.ArrPPDetail(i).PP_Detail_No
                    objPay.Payment_No = ""
                    objPay.Entry_Desc = Desc + " " + DocNo
                    objPay.Payment_Date = clsCommon.myCDate(obj.Doc_Date)
                    objPay.Payment_Post_Date = clsCommon.myCDate(obj.Doc_Date)
                    objPay.Bank_Code = obj.ArrPPDetail.Item(i).Bank_Code
                    objPay.Payment_Type = "PY"
                    objPay.Vendor_Code = obj.ArrPPDetail.Item(i).VSP_CODE
                    objPay.Vendor_Name = obj.ArrPPDetail.Item(i).VSP_NAME
                    objPay.Payment_Code = obj.ArrPPDetail.Item(i).Payment_Mode
                    objPay.Cheque_No = obj.ArrPPDetail.Item(i).Cheque_No
                    If Not obj.ArrPPDetail.Item(i).Cheque_Dated Is Nothing Then
                        objPay.Cheque_Date = obj.ArrPPDetail.Item(i).Cheque_Dated
                    End If

                    objPay.Account_Payee = 0
                    objPay.memorndmamt = "0"
                    objPay.Applied_Payment = clsCommon.myCstr(obj.ArrPPDetail.Item(i).AP_Invoice_No)
                    objPay.Is_Security = 0
                    objPay.IsChkReverse = "N"
                    objPay.Bank_Charges = 0

                    objPay.ArrTr = New List(Of clsPaymentDetail)
                    objTr = New clsPaymentDetail()
                    objTr.Apply = "1"
                    objTr.Payment_Type = "PY"
                    objTr.Document_No = clsCommon.myCstr(obj.ArrPPDetail.Item(i).AP_Invoice_No)
                    objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Total_Invoice_Amount)
                    objTr.Applied_Amount = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Total_Invoice_Amount + ReturnAdjAmt - AdjAmt - obj.ArrPPDetail.Item(i).Advance_Payment_Amount_Knock_Off - (obj.ArrPPDetail(i).Credit_Note_Amount + obj.ArrPPDetail(i).Compulsory_Amount))
                    objTr.Pending_Balance = 0
                    objTr.Vendor_Invoice_No = obj.ArrPPDetail.Item(i).Milk_Purchase_Invoice_No
                    objTr.Net_Balance = 0
                    objTr.Security_Amount = 0
                    objPay.ArrTr.Add(objTr)

                    Dim dtExtra As DataTable = clsDBFuncationality.GetDataTable("select Head_Load_Doc_No,Vsp_Own_System_Doc_No,Deduction_Doc_No from TSPL_PAYMENT_PROCESS_INVOICE where Doc_No='" + obj.ArrPPDetail.Item(i).Doc_No + "' and AP_Invoice_No='" + obj.ArrPPDetail.Item(i).AP_Invoice_No + "'", trans)
                    If dtExtra IsNot Nothing AndAlso dtExtra.Rows.Count > 0 Then
                        If clsCommon.myLen(dtExtra.Rows(0)("Vsp_Own_System_Doc_No")) > 0 Then
                            objTr = New clsPaymentDetail()
                            objTr.Apply = "1"
                            objTr.Payment_Type = "PY"
                            objTr.Document_No = clsCommon.myCstr(dtExtra.Rows(0)("Vsp_Own_System_Doc_No"))
                            objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Vsp_Own_System_Amount)
                            objTr.Applied_Amount = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Vsp_Own_System_Amount)
                            objTr.Pending_Balance = 0
                            objTr.Vendor_Invoice_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & clsCommon.myCstr(objTr.Document_No) & "'", trans))
                            objTr.Net_Balance = 0
                            objTr.Security_Amount = 0
                            objPay.ArrTr.Add(objTr)
                        End If
                        If clsCommon.myLen(dtExtra.Rows(0)("Head_Load_Doc_No")) > 0 Then
                            objTr = New clsPaymentDetail()
                            objTr.Apply = "1"
                            objTr.Payment_Type = "PY"
                            objTr.Document_No = clsCommon.myCstr(dtExtra.Rows(0)("Head_Load_Doc_No"))

                            objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Head_Load_Amount)
                            objTr.Applied_Amount = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Head_Load_Amount)
                            objTr.Pending_Balance = 0
                            objTr.Vendor_Invoice_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & clsCommon.myCstr(objTr.Document_No) & "'", trans))
                            objTr.Net_Balance = 0
                            objTr.Security_Amount = 0
                            objPay.ArrTr.Add(objTr)
                        End If
                        If clsCommon.myLen(dtExtra.Rows(0)("Deduction_Doc_No")) > 0 Then
                            objTr = New clsPaymentDetail()
                            objTr.Apply = "1"
                            objTr.Payment_Type = "PY"
                            objTr.Document_No = clsCommon.myCstr(dtExtra.Rows(0)("Deduction_Doc_No"))
                            objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Deduction_Amount)
                            objTr.Applied_Amount = clsCommon.myCDecimal(obj.ArrPPDetail.Item(i).Deduction_Amount)
                            objTr.Pending_Balance = 0
                            objTr.Vendor_Invoice_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & clsCommon.myCstr(objTr.Document_No) & "'", trans))
                            objTr.Net_Balance = 0
                            objTr.Security_Amount = 0
                            objPay.ArrTr.Add(objTr)
                        End If
                    End If

                    If obj.arrClsPaymentProcessItemIssue IsNot Nothing And obj.arrClsPaymentProcessItemIssue.Count > 0 Then
                        For j As Integer = 0 To obj.arrClsPaymentProcessItemIssue.Count - 1
                            If clsCommon.CompairString(obj.arrClsPaymentProcessItemIssue(j).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                objTr = New clsPaymentDetail()
                                objTr.Apply = "1"
                                objTr.Payment_Type = "PY"
                                objTr.Document_No = clsCommon.myCstr(obj.arrClsPaymentProcessItemIssue.Item(j).AP_Invoice_No)
                                objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrClsPaymentProcessItemIssue.Item(j).Amount)
                                objTr.Applied_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessItemIssue.Item(j).Amount - obj.arrClsPaymentProcessItemIssue.Item(j).Reduce_Deduc_Amt)
                                objTr.Pending_Balance = obj.arrClsPaymentProcessItemIssue.Item(j).Reduce_Deduc_Amt
                                objTr.Vendor_Invoice_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & clsCommon.myCstr(obj.arrClsPaymentProcessItemIssue.Item(j).AP_Invoice_No) & "'", trans))
                                objTr.Net_Balance = obj.arrClsPaymentProcessItemIssue.Item(j).Reduce_Deduc_Amt
                                objTr.Security_Amount = 0
                                objPay.ArrTr.Add(objTr)
                            End If
                        Next
                    End If
                    If obj.arrClsPaymentProcessItemIssueReturn IsNot Nothing And obj.arrClsPaymentProcessItemIssueReturn.Count > 0 Then
                        For j As Integer = 0 To obj.arrClsPaymentProcessItemIssueReturn.Count - 1
                            If clsCommon.CompairString(obj.arrClsPaymentProcessItemIssueReturn(j).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                objTr = New clsPaymentDetail()
                                objTr.Apply = "1"
                                objTr.Payment_Type = "PY"
                                objTr.Document_No = clsCommon.myCstr(obj.arrClsPaymentProcessItemIssueReturn.Item(j).AP_Invoice_No)
                                objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrClsPaymentProcessItemIssueReturn.Item(j).Amount)
                                objTr.Applied_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessItemIssueReturn.Item(j).Amount)
                                objTr.Pending_Balance = 0
                                objTr.Vendor_Invoice_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & clsCommon.myCstr(obj.arrClsPaymentProcessItemIssueReturn.Item(j).AP_Invoice_No) & "'", trans))
                                objTr.Net_Balance = 0
                                objTr.Security_Amount = 0
                                objPay.ArrTr.Add(objTr)
                            End If
                        Next
                    End If
                    If obj.arrClsPaymentProcessDeductions IsNot Nothing And obj.arrClsPaymentProcessDeductions.Count > 0 Then
                        For k As Integer = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).DRTo
                            If clsCommon.CompairString(obj.arrClsPaymentProcessDeductions(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                objTr = New clsPaymentDetail()
                                objTr.Apply = "1"
                                objTr.Payment_Type = "PY"
                                objTr.Document_No = clsCommon.myCstr(obj.arrClsPaymentProcessDeductions.Item(k).AP_Invoice_No)
                                objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrClsPaymentProcessDeductions.Item(k).Amount)
                                objTr.Applied_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessDeductions.Item(k).Amount - obj.arrClsPaymentProcessDeductions.Item(k).Reduce_Deduc_Amt)
                                objTr.Pending_Balance = obj.arrClsPaymentProcessDeductions.Item(k).Reduce_Deduc_Amt
                                Dim vendorInvNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & obj.arrClsPaymentProcessDeductions.Item(k).AP_Invoice_No & "'", trans))
                                objTr.Net_Balance = obj.arrClsPaymentProcessDeductions.Item(k).Reduce_Deduc_Amt
                                objTr.Vendor_Invoice_No = vendorInvNo
                                objTr.Security_Amount = 0
                                objPay.ArrTr.Add(objTr)
                            End If
                        Next
                    End If

                    If obj.arrClsPaymentProcessCreditNote IsNot Nothing And obj.arrClsPaymentProcessCreditNote.Count > 0 Then
                        For k As Integer = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).CRFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).CRTo
                            If clsCommon.CompairString(obj.arrClsPaymentProcessCreditNote(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                Dim tAmt As Decimal = clsCommon.myCDecimal(obj.arrClsPaymentProcessCreditNote.Item(k).Amount)
                                If arrCreditNoteAdjustAmt.ContainsKey(obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No) Then
                                    tAmt -= arrCreditNoteAdjustAmt(obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No)
                                End If
                                If tAmt > 0 Then
                                    objTr = New clsPaymentDetail()
                                    objTr.Apply = "1"
                                    objTr.Payment_Type = "PY"
                                    objTr.Document_No = clsCommon.myCstr(obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No)
                                    objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrClsPaymentProcessCreditNote.Item(k).Amount)
                                    objTr.Applied_Amount = tAmt
                                    objTr.Pending_Balance = 0
                                    Dim vendorInvNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & obj.arrClsPaymentProcessCreditNote.Item(k).AP_Invoice_No & "'", trans))
                                    objTr.Net_Balance = 0
                                    objTr.Vendor_Invoice_No = vendorInvNo
                                    objTr.Security_Amount = 0
                                    objPay.ArrTr.Add(objTr)
                                End If
                            End If
                        Next
                    End If

                    If obj.arrclsPaymentProcessCompulsory IsNot Nothing And obj.arrclsPaymentProcessCompulsory.Count > 0 Then
                        For k As Integer = obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).COMPULSORYFrom To obj.ArrIndex(obj.ArrPPDetail(i).VSP_CODE.ToUpper()).COMPULSORYTo
                            If clsCommon.CompairString(obj.arrclsPaymentProcessCompulsory(k).Vendor_CODE, obj.ArrPPDetail(i).VSP_CODE) = CompairStringResult.Equal Then
                                Dim tAmt As Decimal = clsCommon.myCDecimal(obj.arrclsPaymentProcessCompulsory.Item(k).Amount)
                                If arrCreditNoteAdjustAmt.ContainsKey(obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No) Then
                                    tAmt -= arrCreditNoteAdjustAmt(obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No)
                                End If
                                If tAmt > 0 Then
                                    objTr = New clsPaymentDetail()
                                    objTr.Apply = "1"
                                    objTr.Payment_Type = "PY"
                                    objTr.Document_No = clsCommon.myCstr(obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No)
                                    objTr.Original_Invoice_Amt = clsCommon.myCDecimal(obj.arrclsPaymentProcessCompulsory.Item(k).Amount)
                                    objTr.Applied_Amount = tAmt
                                    objTr.Pending_Balance = 0
                                    Dim vendorInvNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Invoice_No  from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & obj.arrclsPaymentProcessCompulsory.Item(k).AP_Invoice_No & "'", trans))
                                    objTr.Net_Balance = 0
                                    objTr.Vendor_Invoice_No = vendorInvNo
                                    objTr.Security_Amount = 0
                                    objPay.ArrTr.Add(objTr)
                                End If
                            End If
                        Next
                    End If

                    objPay.Payment_Amount = obj.ArrPPDetail(i).Payable_Amount
                    objPay.Balance_Amt = obj.ArrPPDetail(i).Payable_Amount
                    objPay.Location_Code = clsCommon.myCstr(obj.Loc_Seg_Code)
                    objPay.Entry_Desc = obj.PaymentDesc + " " + DocNo

                    objPay.SaveData(objPay, True, trans, True)
                    clsPaymentHeader.PostData(objPay.Payment_No, "Payable", trans)
                    ''------------------Payment Entry End Here
                    'Throw New Exception("Balwinder singh premi")
                Next
                '==============Add Code for save Mcc Sale Return Document Payment Details==================
                If obj.arrClsPaymentProcessMccSaleReturn IsNot Nothing AndAlso obj.arrClsPaymentProcessMccSaleReturn.Count > 0 Then
                    DisCCodeForArAdj = clsFixedParameter.GetData(clsFixedParameterType.DiscountCodeForArAdj, clsFixedParameterCode.DiscountCodeForArAdj, trans)
                    If clsCommon.myLen(DisCCodeForArAdj) <= 0 Then
                        Throw New Exception("Please Map Discount code from Sale setting")
                    End If
                    DiscDiscForArAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
                    GLAcARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Code from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
                    GLAcDescARAdj = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select  Account_Description from TSPL_Discount_Master WHERE Code='" & DisCCodeForArAdj & "'", trans))
                    If ShowProgressBAR Then
                        clsCommon.ProgressBarPercentShow()
                        isProgressBarShownLocal = True
                    End If
                    For i = 0 To obj.arrClsPaymentProcessMccSaleReturn.Count - 1
                        If ShowProgressBAR Then
                            clsCommon.ProgressBarPercentUpdate(i * 100 / obj.arrClsPaymentProcessMccSaleReturn.Count, "Updating AR Adjustment Record " & (i + 1) & " Of " & obj.arrClsPaymentProcessMccSaleReturn.Count)
                        End If
                        objRcpt = New clsAdjustmentEntryReceivables
                        objRcpt.Adjustment_No = ""
                        objRcpt.Description = " Return Adjustment Against Bulk Payment Process "
                        objRcpt.Adjustment_Date = clsCommon.myCDate(obj.Doc_Date)
                        objRcpt.Customer_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Cust_Code  from TSPL_CUSTOMER_VENDOR_MAPPING where Vendor_Code ='" & obj.arrClsPaymentProcessMccSaleReturn(i).Customer_CODE & "' ", trans))
                        objRcpt.Customer_Name = clsCommon.myCstr(clsCustomerMaster.GetName(objRcpt.Customer_No, trans))
                        objRcpt.Doc_No = clsCommon.myCstr(obj.arrClsPaymentProcessMccSaleReturn(i).Return_Doc_No)
                        objRcpt.ARInvoiceNo = clsCommon.myCstr(obj.arrClsPaymentProcessMccSaleReturn(i).AR_Invoice_No)
                        objRcpt.Doc_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSaleReturn(i).Amount)
                        objRcpt.Remarks = ""
                        objRcpt.Adjustment_Amount = clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSaleReturn(i).Amount)
                        objRcpt.Arr = New List(Of clsAdjustmentEntryReceivablesDetail)
                        Dim objTrRcpt As New clsAdjustmentEntryReceivablesDetail()
                        objTrRcpt.Discount_Code = DisCCodeForArAdj
                        objTrRcpt.Discount_Description = DiscDiscForArAdj
                        objTrRcpt.Account_No = GLAcARAdj
                        objTrRcpt.Account_Description = GLAcDescARAdj
                        Dim ReturnAmt As Decimal = 0
                        objTrRcpt.Amount = clsCommon.myCDecimal(clsCommon.myCDecimal(obj.arrClsPaymentProcessMccSaleReturn(i).Amount))
                        objTrRcpt.Remarks = ""
                        objRcpt.Arr.Add(objTrRcpt)
                        If clsCommon.myCDecimal(objTrRcpt.Amount) > 0 Then
                            objRcpt.SaveData(objRcpt, True, trans)
                            clsAdjustmentEntryReceivables.FunPostReverseEntry(objRcpt.Adjustment_No, trans)
                        End If
                    Next

                    If ShowProgressBAR Then
                        isProgressBarShownLocal = False
                        clsCommon.ProgressBarPercentHide()
                    End If

                End If
                '===========================================================================================================
                If ShowProgressBAR Then
                    isProgressBarShownLocal = False
                    clsCommon.ProgressBarPercentHide()
                End If

                Dim qry As String = " update TSPL_PAYMENT_PROCESS_HEAD set isPosted=1, Posting_Date='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy") & "' where doc_no='" & obj.Doc_No & "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Doc_No, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

                If arrARForSameMCCSaleAndReturn IsNot Nothing AndAlso arrARForSameMCCSaleAndReturn.Count > 0 Then
                    For Each kvp As KeyValuePair(Of String, Decimal) In arrARForSameMCCSaleAndReturn
                        clsReceiptDettail.funBalanceAmtSave(kvp.Key, kvp.Value, trans, "", "C")
                    Next
                End If
            End If
        Catch ex As Exception
            If ShowProgressBAR Then
                If isProgressBarShownLocal Then
                    clsCommon.ProgressBarPercentHide()
                End If
            End If
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Shared Function CreateApplyDocumentForAdavancePayment(ByVal dtTo As DateTime, ByVal ArrPPAdvancePayment As List(Of clsPaymentProcessAdvancePayment), ByVal objtr As clsPaymentProcessDetail, ByVal strLocSeg As String, ByVal trans As SqlTransaction)
        Dim KFAmt As Decimal = objtr.Advance_Payment_Amount_Knock_Off
        If objtr.Advance_Payment_Amount_Knock_Off > 0 Then
            If ArrPPAdvancePayment IsNot Nothing AndAlso ArrPPAdvancePayment.Count > 0 Then
                For Each objAdvancePayment As clsPaymentProcessAdvancePayment In ArrPPAdvancePayment
                    If clsCommon.CompairString(objAdvancePayment.Vendor_Code, objtr.VSP_CODE) = CompairStringResult.Equal Then
                        Dim objPay As clsPaymentHeader = New clsPaymentHeader()
                        objPay.Payment_No = ""
                        objPay.Payment_Date = dtTo
                        objPay.Payment_Post_Date = dtTo
                        objPay.Bank_Code = objtr.Bank_Code
                        objPay.Payment_Type = "AD"
                        objPay.Vendor_Code = objtr.VSP_CODE
                        objPay.Vendor_Name = objtr.VSP_NAME
                        objPay.Payment_Amount = objAdvancePayment.Amount_Knock_Off
                        objPay.Balance_Amt = objAdvancePayment.Amount_Knock_Off
                        objPay.Location_GL_Code = strLocSeg
                        objPay.Entry_Desc = "Apply document of Advance payment"
                        objPay.Applied_Payment = objAdvancePayment.Payment_No
                        objPay.Against_PP_Detail_No_Advance_Payment = objtr.PP_Detail_No

                        objPay.ArrTr = New List(Of clsPaymentDetail)

                        Dim objPayTr As clsPaymentDetail = New clsPaymentDetail()
                        objPayTr.Apply = "1"
                        objPayTr.Payment_Type = "AD"
                        objPayTr.Document_No = objtr.AP_Invoice_No
                        objPayTr.Net_Balance = objtr.Total_Invoice_Amount
                        objPayTr.Original_Invoice_Amt = objtr.Total_Invoice_Amount   '' objtr.Advance_Payment_Amount_Knock_Off + objtr.Payable_Amount
                        objPayTr.Applied_Amount = objAdvancePayment.Amount_Knock_Off
                        KFAmt -= objAdvancePayment.Amount_Knock_Off
                        objPayTr.Pending_Balance = KFAmt


                        objPayTr.Vendor_Invoice_No = objtr.Milk_Purchase_Invoice_No

                        objPayTr.Security_Amount = 0
                        objPayTr.ConvRateOld = 1
                        objPay.ArrTr.Add(objPayTr)

                        objPay.SaveData(objPay, True, trans, True)
                        clsPaymentHeader.PostData(objPay.Payment_No, "Payable", trans)
                    End If
                Next
            End If
        End If

        Return True
    End Function

    Shared Function createDebitNote(ByVal objtr As clsPaymentProcessDetail, ByVal strLocSeg As String, ByVal trans As SqlTransaction)
        If objtr.MP_Net_Amount > 0 AndAlso objtr.VSP_Amount > objtr.MP_Net_Amount Then
            Dim dblAmt As Decimal = objtr.VSP_Amount - objtr.MP_Net_Amount
            Dim objVendorInvHead As New clsVedorInvoiceHead()
            Dim objVendorInvDetail As New clsVedorInvoiceDetail()
            objVendorInvHead.isDeduction = 1
            Dim dtDed As DataTable = clsDBFuncationality.GetDataTable("select code from TSPL_DEDUCTION_MASTER  where Is_Default_VSP_Deduction=1", trans)
            If dtDed Is Nothing OrElse dtDed.Rows.Count <= 0 Then
                Throw New Exception("Please set default VSP deduction code")
            End If
            objVendorInvDetail.DeductionCode = clsCommon.myCstr(dtDed.Rows(0)("code"))

            'objVendorInvHead.Document_No = txtDocNo.Value'ToBeGenerated
            objVendorInvHead.Invoice_Entry_Date = clsCommon.GetPrintDate(objtr.AP_Invoice_Date.AddDays(1), "dd/MMM/yyyy")
            objVendorInvHead.Vendor_Code = objtr.VSP_CODE
            objVendorInvHead.Vendor_Name = objtr.VSP_NAME
            objVendorInvHead.Vendor_Invoice_No = "" ''No Need to send vendor invoice no because it is of debit note type
            objVendorInvHead.Invoice_Type = "AP"
            objVendorInvHead.Vendor_Invoice_Date = objVendorInvHead.Invoice_Entry_Date
            objVendorInvHead.loc_code = strLocSeg
            'objVendorInvHead.Irregular_loc_code = obj.Irregular_MCC_CODE
            objVendorInvHead.Description = "AP Debit Note Against VSP Vs MP Collection : " & objtr.PP_Detail_No & " .VSP : " & objVendorInvHead.Vendor_Name & "(" + objVendorInvHead.Vendor_Code + ")"
            'objVendorInvHead.PROJECT_ID = 1 'obj.PROJECT_ID
            objVendorInvHead.Account_Set = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Vendor_Account from TSPL_VENDOR_MASTER where Vendor_Code ='" + objVendorInvHead.Vendor_Code + "'", trans))
            If (clsCommon.myLen(objVendorInvHead.Account_Set) < 0) Then
                Throw New Exception("Please set the vendor Account Set For Vendor : " + objVendorInvHead.Vendor_Name)
            End If

            objVendorInvHead.Document_Type = "D" ''For Purchase Invoice Type
            ''objVendorInvHead.PO_Number = obj.p

            '' ''added by priti
            objVendorInvHead.RefDocType = "VSP-MP"
            objVendorInvHead.RefDocNo = objtr.PP_Detail_No
            'objVendorInvHead.Ref_SNo = objtr.SAMPLE_NO
            '' '' priti ends here
            'objVendorInvHead.Order_No = txtOrderNo.Text
            ' objVendorInvHead.Total_Tax = 0

            objVendorInvHead.On_Hold = False
            'Dim srndate As String = ""
            'Dim srncode As String = ""
            'Dim Vlc_Code As String = ""
            'Dim Vlc_Name As String = ""
            'For Each objTr As clsMilkPurchaseInvoiceMCCDetail In obj.ObjList
            '    If clsCommon.myLen(objTr.SRN_CODE) > 0 Then
            '        Dim query As String = "select doc_date,vd.VLC_Code,VLC_Name from TSPL_Milk_SRN_HEAD sh inner join TSPL_VLC_MASTER_HEAD vd on sh.VLC_CODE=vd.VLC_Code where DOc_Code ='" + objTr.SRN_CODE + "' "
            '        Dim Dt_SRN As DataTable = clsDBFuncationality.GetDataTable(query, trans)
            '        srndate = IIf(srndate = "", clsCommon.myCDate(CStr(Dt_SRN.Rows(0).Item("Doc_Date")), "dd/MMM/yyyy"), srndate & "," & clsCommon.myCDate(CStr(Dt_SRN.Rows(0).Item("Doc_Date")), "dd/MMM/yyyy"))
            '        srncode = IIf(srncode = "", objTr.SRN_CODE, srncode & "," & objTr.SRN_CODE)
            '        Vlc_Code = IIf(Vlc_Code = "", Dt_SRN.Rows(0).Item("VLC_Code").ToString, Vlc_Code & "," & Dt_SRN.Rows(0).Item("VLC_Code").ToString)
            '        Vlc_Name = IIf(Vlc_Name = "", Dt_SRN.Rows(0).Item("VLC_Name").ToString, Vlc_Name & "," & Dt_SRN.Rows(0).Item("VLC_name").ToString)
            '    End If
            'Next



            'objVendorInvHead.Description = "VSP : " + obj.VSP_CODE + "/" + vendor_name + "VLC : " + Vlc_Code + "/" + Vlc_Name + " .Against PI Invoice No " + obj.DOC_CODE + "-" + srncode + "-" + srndate
            'objVendorInvHead.Tax_Calculation_Type = Nothing
            'objVendorInvHead.Tax_Group = Nothing
            'If (clsCommon.myLen(obj.TAX1) > 0) Then
            '    objVendorInvHead.TAX1 = obj.TAX1
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX1, trans) Then
            '        objVendorInvHead.TAX1_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX1, trans)
            '        objVendorInvHead.TAX1_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX1_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX1_Rate = obj.TAX1_Rate
            '    objVendorInvHead.Tax1_BAmount = obj.TAX1_Base_Amt
            '    objVendorInvHead.TAX1_Amt = obj.TAX1_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX2) > 0) Then
            '    objVendorInvHead.TAX2 = obj.TAX2
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans) Then
            '        objVendorInvHead.TAX2_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX2, trans)
            '        objVendorInvHead.TAX2_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX2_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX2_Rate = obj.TAX2_Rate
            '    objVendorInvHead.Tax2_BAmount = obj.TAX2_Base_Amt
            '    objVendorInvHead.TAX2_Amt = obj.TAX2_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX3) > 0) Then
            '    objVendorInvHead.TAX3 = obj.TAX3
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans) Then
            '        objVendorInvHead.TAX3_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX3, trans)
            '        objVendorInvHead.TAX3_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX3_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX3_Rate = obj.TAX3_Rate
            '    objVendorInvHead.Tax3_BAmount = obj.TAX3_Base_Amt
            '    objVendorInvHead.TAX3_Amt = obj.TAX3_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX4) > 0) Then
            '    objVendorInvHead.TAX4 = obj.TAX4
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans) Then
            '        objVendorInvHead.TAX4_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX4, trans)
            '        objVendorInvHead.TAX4_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX4_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX4_Rate = obj.TAX4_Rate
            '    objVendorInvHead.Tax4_BAmount = obj.TAX4_Base_Amt
            '    objVendorInvHead.TAX4_Amt = obj.TAX4_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX5) > 0) Then
            '    objVendorInvHead.TAX5 = obj.TAX5
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans) Then
            '        objVendorInvHead.TAX5_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX5, trans)
            '        objVendorInvHead.TAX5_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX5_GLAC, obj.MCC_CODE, trans)

            '    End If
            '    objVendorInvHead.TAX5_Rate = obj.TAX5_Rate
            '    objVendorInvHead.Tax5_BAmount = obj.TAX5_Base_Amt
            '    objVendorInvHead.TAX5_Amt = obj.TAX5_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX6) > 0) Then
            '    objVendorInvHead.TAX6 = obj.TAX6
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans) Then
            '        objVendorInvHead.TAX6_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX6, trans)
            '        objVendorInvHead.TAX6_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX6_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX6_Rate = obj.TAX6_Rate
            '    objVendorInvHead.Tax6_BAmount = obj.TAX6_Base_Amt
            '    objVendorInvHead.TAX6_Amt = obj.TAX6_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX7) > 0) Then
            '    objVendorInvHead.TAX7 = obj.TAX7
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans) Then
            '        objVendorInvHead.TAX7_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX7, trans)
            '        objVendorInvHead.TAX7_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX7_GLAC, obj.MCC_CODE, trans)

            '    End If
            '    objVendorInvHead.TAX7_Rate = obj.TAX7_Rate
            '    objVendorInvHead.Tax7_BAmount = obj.TAX7_Base_Amt
            '    objVendorInvHead.TAX7_Amt = obj.TAX7_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX8) > 0) Then
            '    objVendorInvHead.TAX8 = obj.TAX8
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans) Then
            '        objVendorInvHead.TAX8_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX8, trans)
            '        objVendorInvHead.TAX8_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX8_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX8_Rate = obj.TAX8_Rate
            '    objVendorInvHead.Tax8_BAmount = obj.TAX8_Base_Amt
            '    objVendorInvHead.TAX8_Amt = obj.TAX8_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX9) > 0) Then
            '    objVendorInvHead.TAX9 = obj.TAX9
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans) Then
            '        objVendorInvHead.TAX9_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX9, trans)
            '        objVendorInvHead.TAX9_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX9_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX9_Rate = obj.TAX9_Rate
            '    objVendorInvHead.Tax9_BAmount = obj.TAX9_Base_Amt
            '    objVendorInvHead.TAX9_Amt = obj.TAX9_Amt
            'End If
            'If (clsCommon.myLen(obj.TAX10) > 0) Then
            '    objVendorInvHead.TAX10 = obj.TAX10
            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans) Then
            '        objVendorInvHead.TAX10_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX10, trans)
            '        objVendorInvHead.TAX10_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX10_GLAC, obj.MCC_CODE, trans)
            '    End If
            '    objVendorInvHead.TAX10_Rate = obj.TAX10_Rate
            '    objVendorInvHead.Tax10_BAmount = obj.TAX10_Base_Amt
            '    objVendorInvHead.TAX10_Amt = obj.TAX10_Amt
            'End If

            'objVendorInvHead.Terms_Code = obj.Terms_Code
            'objVendorInvHead.Terms_Description = obj.TermsName
            objVendorInvHead.Due_Date = objVendorInvHead.Invoice_Entry_Date

            'objVendorInvHead.Against_POInvoice_No = obj.DOC_CODE
            'objVendorInvHead.Against_MillkPurchaseInvoice_No = obj.DOC_CODE

            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Acct_Set_Code,Payable_Account,Discount_Account,Deduction_ACCOUNT from TSPL_VENDOR_ACCOUNT_SET  where Acct_Set_Code='" + objVendorInvHead.Account_Set + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                objVendorInvHead.Vendor_Control_AC = clsCommon.myCstr(dt.Rows(0)("Payable_Account"))
                objVendorInvHead.Vendor_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.Vendor_Control_AC, strLocSeg, True, trans)
                If clsCommon.myCDecimal(objVendorInvHead.Discount_Amount) > 0 Then
                    objVendorInvHead.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Discount_Account"))
                    objVendorInvHead.Discount_GL_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.Discount_GL_AC, strLocSeg, True, trans)
                End If
            End If
            If clsCommon.myLen(objVendorInvHead.Vendor_Control_AC) <= 0 Then
                Throw New Exception("Please set the vendor payable Account")
            End If

            'objVendorInvHead.Total_Add_Charge = obj.Total_Add_Charge

            'objVendorInvHead.Add_Charge_Code1 = obj.Add_Charge_Code1
            'objVendorInvHead.Add_Charge_Name1 = obj.Add_Charge_Name1
            'objVendorInvHead.Add_Charge_Amt1 = obj.Add_Charge_Amt1

            'objVendorInvHead.Add_Charge_Code2 = obj.Add_Charge_Code2
            'objVendorInvHead.Add_Charge_Name2 = obj.Add_Charge_Name2
            'objVendorInvHead.Add_Charge_Amt2 = obj.Add_Charge_Amt2

            'objVendorInvHead.Add_Charge_Code3 = obj.Add_Charge_Code3
            'objVendorInvHead.Add_Charge_Name3 = obj.Add_Charge_Name3
            'objVendorInvHead.Add_Charge_Amt3 = obj.Add_Charge_Amt3

            'objVendorInvHead.Add_Charge_Code4 = obj.Add_Charge_Code4
            'objVendorInvHead.Add_Charge_Name4 = obj.Add_Charge_Name4
            'objVendorInvHead.Add_Charge_Amt4 = obj.Add_Charge_Amt4

            'objVendorInvHead.Add_Charge_Code5 = obj.Add_Charge_Code5
            'objVendorInvHead.Add_Charge_Name5 = obj.Add_Charge_Name5
            'objVendorInvHead.Add_Charge_Amt5 = obj.Add_Charge_Amt5

            'objVendorInvHead.Add_Charge_Code6 = obj.Add_Charge_Code6
            'objVendorInvHead.Add_Charge_Name6 = obj.Add_Charge_Name6
            'objVendorInvHead.Add_Charge_Amt6 = obj.Add_Charge_Amt6

            'objVendorInvHead.Add_Charge_Code7 = obj.Add_Charge_Code7
            'objVendorInvHead.Add_Charge_Name7 = obj.Add_Charge_Name7
            'objVendorInvHead.Add_Charge_Amt7 = obj.Add_Charge_Amt7

            'objVendorInvHead.Add_Charge_Code8 = obj.Add_Charge_Code8
            'objVendorInvHead.Add_Charge_Name8 = obj.Add_Charge_Name8
            'objVendorInvHead.Add_Charge_Amt8 = obj.Add_Charge_Amt8

            'objVendorInvHead.Add_Charge_Code9 = obj.Add_Charge_Code9
            'objVendorInvHead.Add_Charge_Name9 = obj.Add_Charge_Name9
            'objVendorInvHead.Add_Charge_Amt9 = obj.Add_Charge_Amt9

            'objVendorInvHead.Add_Charge_Code10 = obj.Add_Charge_Code10
            'objVendorInvHead.Add_Charge_Name10 = obj.Add_Charge_Name10
            'objVendorInvHead.Add_Charge_Amt10 = obj.Add_Charge_Amt10


            objVendorInvHead.Arr = New List(Of clsVedorInvoiceDetail)
            Dim ii As Integer = 0
            Dim isFirstTime As Boolean = True
            ' Dim strFirstItemCode As String = GetFirstItemCode(obj.ObjList)
            'objVendorInvHead.Empty_Amount = obj.Tot_Empty_Amount
            objVendorInvHead.Total_Landed_Amt = 0

            objVendorInvHead.ArrAssetEMI = New List(Of clsAPInvoiceAssetEMIDetails)()



            ''Set AP Invvoice Detail Table

            Dim strInvCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("Deduction_ACCOUNT"))
            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                Throw New Exception("Please set Deduction Account for Vendor Account set Code :" + clsCommon.myCstr(dt.Rows(0)("Acct_Set_Code")))
            End If
            strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, strLocSeg, True, trans)




            ii = ii + 1
            objVendorInvDetail.Detail_Line_No = ii
            objVendorInvDetail.GL_Account_Code = strInvCtrlAC
            objVendorInvDetail.GL_Account_Desc = clsGLAccount.GetName(strInvCtrlAC, trans)
            objVendorInvDetail.Amount = dblAmt

            objVendorInvDetail.Discount_Per = 0
            objVendorInvDetail.Discount = 0
            objVendorInvDetail.Amount_less_Discount = dblAmt
            objVendorInvDetail.Total_Tax = 0
            objVendorInvDetail.Total_Amount = dblAmt
            objVendorInvDetail.Landed_Amount = dblAmt
            ''End of Set AP Invvoice Detail Table

            If (clsCommon.myLen(objVendorInvDetail.GL_Account_Code) > 0) Then
                objVendorInvHead.Arr.Add(objVendorInvDetail)
            End If

            ''Set AP Invvoice Header Table
            objVendorInvHead.Total_Landed_Amt += dblAmt
            objVendorInvHead.Discount_Base += dblAmt
            objVendorInvHead.Discount_Amount += 0
            objVendorInvHead.Amount_Less_Discount += dblAmt
            objVendorInvHead.Document_Total += dblAmt
            objVendorInvHead.Balance_Amt += dblAmt
            ''End of Set AP Invvoice Header Table

            objVendorInvHead.Empty_Amount = 0 'obj.Tot_Empty_Amount
            If objVendorInvHead.Empty_Amount > 0 Then
                If clsCommon.myLen(objVendorInvHead.Empty_Account) <= 0 Then
                    Throw New Exception("Please set Inventory Control Empties")
                End If
                objVendorInvHead.Document_Total += objVendorInvHead.Empty_Amount
            End If
            If (objVendorInvHead.Arr Is Nothing OrElse objVendorInvHead.Arr.Count <= 0) Then
                Throw New Exception("No GL Account Found For AP Invoice")
            End If
            ''multicurrency
            'objVendorInvHead.CURRENCY_CODE = obj.CURRENCY_CODE
            'objVendorInvHead.ConvRate = 1
            objVendorInvHead.ApplicableFrom = clsCommon.GetPrintDate(objVendorInvHead.Invoice_Entry_Date, "dd/MMM/yyyy")
            ''end multicurrency

            objVendorInvHead.SaveData(objVendorInvHead, True, trans)
            clsVedorInvoiceHead.PostData("", objVendorInvHead.Document_No, "", trans, "", False)
        End If
        Return True
    End Function

    Public Shared Function getData(ByVal strCode As String, ByVal navtype As NavigatorType, Optional ByVal trans As SqlTransaction = Nothing, Optional Vendorcode As String = "", Optional isGetDT As Boolean = False, Optional ApplyIndex As Boolean = False, Optional GetDetailsData As Boolean = True) As clsPaymentProcessHead
        Dim obj As clsPaymentProcessHead = Nothing
        Try
            Dim whrCls As String = String.Empty
            Dim qst As String = " select *   From TSPL_PAYMENT_PROCESS_HEAD   where 1=1 " & whrCls
            'Dim qst As String = "select Area_Location_Code,*   From TSPL_PAYMENT_PROCESS_HEAD  left outer join tspl_mcc_master on tspl_mcc_master.Mcc_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
            'where 1=1 " & whrCls
            Select Case navtype
                Case NavigatorType.Current
                    qst += " and TSPL_PAYMENT_PROCESS_HEAD.Doc_No in ('" + strCode + "')"
                Case NavigatorType.Next
                    qst += " and TSPL_PAYMENT_PROCESS_HEAD.Doc_No in (select min(Doc_No ) from TSPL_PAYMENT_PROCESS_HEAD where Doc_No  >'" + strCode + "' " & whrCls & ") "
                Case NavigatorType.First
                    qst += " and TSPL_PAYMENT_PROCESS_HEAD.Doc_No in (select MIN(Doc_No ) from TSPL_PAYMENT_PROCESS_HEAD where 1=1 " & whrCls & ") "
                Case NavigatorType.Last
                    qst += " and TSPL_PAYMENT_PROCESS_HEAD.Doc_No in (select Max(Doc_No ) from TSPL_PAYMENT_PROCESS_HEAD where 1=1 " & whrCls & ") "
                Case NavigatorType.Previous
                    qst += " and TSPL_PAYMENT_PROCESS_HEAD.Doc_No in (select Max(Doc_No ) from TSPL_PAYMENT_PROCESS_HEAD where Doc_No  <'" + strCode + "' " & whrCls & ") "
            End Select
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qst, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj = New clsPaymentProcessHead
                obj.FarmType = "PP"
                obj.Area_Location_Code = clsCommon.myCstr(dt.Rows(0)("Area_Location_Code"))
                obj.Doc_No = clsCommon.myCstr(dt.Rows(0)("Doc_No"))
                obj.DocRefNoForUploader = clsCommon.myCstr(dt.Rows(0)("DocRefNoForUploader"))
                obj.Doc_Date = clsCommon.myCDate(dt.Rows(0)("Doc_Date"))
                obj.From_Date = clsCommon.myCDate(dt.Rows(0)("From_Date"))
                obj.To_Date = clsCommon.myCDate(dt.Rows(0)("To_Date"))
                obj.Loc_Seg_Code = clsCommon.myCstr(dt.Rows(0)("Loc_Seg_Code"))
                obj.Location_Code_Prefix = clsCommon.myCstr(dt.Rows(0)("Location_Code_Prefix"))
                obj.isPrePosted = clsCommon.myCDecimal(dt.Rows(0)("isPrePosted"))
                obj.isPosted = clsCommon.myCDecimal(dt.Rows(0)("isPosted"))
                obj.PaymentDesc = clsCommon.myCstr(dt.Rows(0)("PaymentDesc"))
                If obj.isPosted = 1 Then
                    obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
                End If
                obj.MCC_Code_Selected = clsCommon.myCstr(dt.Rows(0)("MCC_Code_Selected"))
                obj.Is_Skip_Previous_Item_Issue = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_Item_Issue")) = 1, True, False)
                obj.Is_Skip_Previous_Item_Issue_Return = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_Item_Issue_Return")) = 1, True, False)
                obj.Is_Skip_Previous_MCC_Sale = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_MCC_Sale")) = 1, True, False)
                obj.Is_Skip_Previous_MCC_Sale_Return = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_MCC_Sale_Return")) = 1, True, False)
                obj.Is_Skip_Previous_Credit_Note = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_Credit_Note")) = 1, True, False)
                obj.Is_Skip_Previous_Debit_Note = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_Debit_Note")) = 1, True, False)
                obj.Is_Skip_Previous_Advacee_Payment = IIf(clsCommon.myCDecimal(dt.Rows(0)("Is_Skip_Previous_Advacee_Payment")) = 1, True, False)
                If GetDetailsData Then
                    If isGetDT Then
                        obj.dtClsPaymentProcessInvoices = clsPaymentProcessInvoices.getDataDT(obj.Doc_No, trans)

                        obj.dtPPDetail = clsPaymentProcessDetail.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessMccSale = clsPaymentProcessMCCSale.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessMccSaleReturn = clsPaymentProcessMCCSaleReturn.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessItemIssue = clsPaymentProcessItemIssue.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessItemIssueReturn = clsPaymentProcessItemIssueReturn.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessDeductions = clsPaymentProcessDeduction.getDataDT(obj.Doc_No, trans)
                        obj.dtClsPaymentProcessCreditNote = clsPaymentProcessCreditNote.getDataDT(obj.Doc_No, trans)
                        obj.dtclsPaymentProcessSaving = clsPaymentProcessSaving.getDataDT(obj.Doc_No, trans)
                        obj.dtclsPaymentProcessCompulsory = clsPaymentProcessCompulsory.getDataDT(obj.Doc_No, trans)
                        obj.dtPPAdvancePayment = clsPaymentProcessAdvancePayment.getDataDT(obj.Doc_No, trans)
                        obj.dtPPAssetLost = clsPaymentProcessAssetLost.getDataDT(obj.Doc_No, trans)
                    Else
                        obj.ArrIndex = New Dictionary(Of String, clsPaymentProcessIndex)

                        obj.ArrPPDetail = clsPaymentProcessDetail.getData(ApplyIndex, obj.ArrIndex, obj.Doc_No, trans)
                        obj.arrClsPaymentProcessInvoices = clsPaymentProcessInvoices.getData(ApplyIndex, obj.Doc_No, trans)
                        obj.arrClsPaymentProcessMccSale = clsPaymentProcessMCCSale.getData(obj.Doc_No, trans)
                        obj.arrClsPaymentProcessMccSaleReturn = clsPaymentProcessMCCSaleReturn.getData(obj.Doc_No, trans)
                        obj.arrClsPaymentProcessItemIssue = clsPaymentProcessItemIssue.getData(obj.Doc_No, trans)
                        obj.arrClsPaymentProcessItemIssueReturn = clsPaymentProcessItemIssueReturn.getData(obj.Doc_No, trans)
                        obj.arrClsPaymentProcessDeductions = clsPaymentProcessDeduction.getData(ApplyIndex, obj.ArrIndex, obj.Doc_No, trans)
                        obj.arrClsPaymentProcessCreditNote = clsPaymentProcessCreditNote.getData(ApplyIndex, obj.ArrIndex, obj.Doc_No, trans)
                        obj.arrclsPaymentProcessSaving = clsPaymentProcessSaving.getData(ApplyIndex, obj.ArrIndex, obj.Doc_No, trans)
                        obj.arrclsPaymentProcessCompulsory = clsPaymentProcessCompulsory.getData(ApplyIndex, obj.ArrIndex, obj.Doc_No, trans)
                        obj.ArrPPAdvancePayment = clsPaymentProcessAdvancePayment.getData(obj.Doc_No, trans)
                        obj.ArrPPAssetLost = clsPaymentProcessAssetLost.getData(obj.Doc_No, trans)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return obj
    End Function

    Public Shared Function deleteData(ByVal DocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try


            deleteData(DocNo, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim isDeleted As Boolean = True
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + DocNo + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, clsCommon.myCstr(dt.Rows(0)("Loc_Seg_Code")), clsCommon.myCDate(dt.Rows(0)("Doc_Date")), trans)
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, clsCommon.myCstr(dt.Rows(0)("Area_Location_Code")), clsCommon.myCDate(dt.Rows(0)("Doc_Date")), trans)

            End If
            clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, DocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, DocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

            isDeleted = isDeleted AndAlso clsPaymentProcessInvoices.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessMCCSale.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessItemIssue.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessItemIssueReturn.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessDeduction.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessCreditNote.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessSaving.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessCompulsory.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessDetail.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessAdvancePayment.deleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessSkipDoc.DeleteData(DocNo, trans)
            isDeleted = isDeleted AndAlso clsPaymentProcessAssetLost.deleteData(DocNo, trans)
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_HEAD where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Try
            Dim str As String = ""
            Dim qry As String = ""
            If (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowMCCFinderInPaymentProcess, clsFixedParameterCode.ShowMCCFinderInPaymentProcess, Nothing)) = 1) = True Then
                qry = " select TSPL_PAYMENT_PROCESS_HEAD.Doc_No as [DocumentNo] ,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date as [Doc Date] ,TSPL_PAYMENT_PROCESS_HEAD.From_Date as [From Date] ,TSPL_PAYMENT_PROCESS_HEAD.To_Date as [To Date] ,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code as [Plant Code],TSPL_GL_SEGMENT_CODE.description as [Plant Name], isnull (TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected,'') as [MCC Code]  , isnull (TSPL_MCC_MASTER.MCC_NAME,'') as [MCC Name] ,TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code as AreaCode,AreaMaster.Location_Desc as AreaName,TSPL_PAYMENT_PROCESS_HEAD.Location_Code_Prefix as [Location Code Prefix],TSPL_PAYMENT_PROCESS_HEAD.Created_By as [Created By] ,TSPL_PAYMENT_PROCESS_HEAD.Created_Date as [Created Date] ,TSPL_PAYMENT_PROCESS_HEAD.Modified_By as [Modified By] ,TSPL_PAYMENT_PROCESS_HEAD.Modified_Date as [Modified Date] ,TSPL_PAYMENT_PROCESS_HEAD.Comp_Code as [Comp Code] ,case when isnull(TSPL_PAYMENT_PROCESS_HEAD.isPrePosted,0)=0 then 'NO' else 'YES' end as [Posting Status] ,TSPL_PAYMENT_PROCESS_HEAD.Posting_Date as [Posting Date],TSPL_PAYMENT_PROCESS_HEAD.DocRefNoForUploader as [NEFT Uploader Ref No],PMode.Payment_Mode as [Payment Mode],PMode.Payable_Amount as [Payable Amount] " &
                " From TSPL_PAYMENT_PROCESS_HEAD 
left outer join TSPL_LOCATION_MASTER as AreaMaster on AreaMaster.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
left outer join TSPL_GL_SEGMENT_CODE  on TSPL_GL_SEGMENT_CODE.segment_code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code 
left join (select Doc_No as PP_Code,Max(Payment_Mode) as Payment_Mode,sum(Payable_Amount) as Payable_Amount  from TSPL_PAYMENT_PROCESS_DETAIL group by Doc_No) PMode on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=PMode.PP_Code " &
                " left Outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   "
            Else
                qry = " select TSPL_PAYMENT_PROCESS_HEAD.Doc_No as [DocumentNo] ,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date as [Doc Date] ,TSPL_PAYMENT_PROCESS_HEAD.From_Date as [From Date] ,TSPL_PAYMENT_PROCESS_HEAD.To_Date as [To Date] ,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code as [MCC Code],TSPL_GL_SEGMENT_CODE.description as [MCC Name],TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code as AreaCode,AreaMaster.Location_Desc as AreaName ,TSPL_PAYMENT_PROCESS_HEAD.Created_By as [Created By] ,TSPL_PAYMENT_PROCESS_HEAD.Created_Date as [Created Date] ,TSPL_PAYMENT_PROCESS_HEAD.Modified_By as [Modified By] ,TSPL_PAYMENT_PROCESS_HEAD.Modified_Date as [Modified Date] ,TSPL_PAYMENT_PROCESS_HEAD.Comp_Code as [Comp Code] ,case when isnull(TSPL_PAYMENT_PROCESS_HEAD.isPrePosted,0)=0 then 'NO' else 'YES' end as [Posting Status] ,TSPL_PAYMENT_PROCESS_HEAD.Posting_Date as [Posting Date],TSPL_PAYMENT_PROCESS_HEAD.DocRefNoForUploader as [NEFT Uploader Ref No],PMode.Payment_Mode as [Payment Mode],PMode.Payable_Amount as [Payable Amount] 
From TSPL_PAYMENT_PROCESS_HEAD 
left outer join TSPL_LOCATION_MASTER as AreaMaster on AreaMaster.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
left outer join TSPL_GL_SEGMENT_CODE  on TSPL_GL_SEGMENT_CODE.segment_code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code 
left join (select Doc_No as PP_Code,Max(Payment_Mode) as Payment_Mode,sum(Payable_Amount) as Payable_Amount  from TSPL_PAYMENT_PROCESS_DETAIL group by Doc_No) PMode on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=PMode.PP_Code  "
            End If
            str = clsCommon.ShowSelectForm("fndPayProcess", qry, "DocumentNo", whrcls, curcode, "DocumentNo", isButtonClicked, "Doc_Date")
            Return str
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function Unpost(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(strDocNo) <= 0 Then
                Throw New Exception("Payment process Document no Not found to  unpost")
            End If

            Dim qry As String = "select TSPL_PAYMENT_PROCESS_HEAD.isPosted,TSPL_PAYMENT_PROCESS_HEAD.isPrePosted,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code,TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + strDocNo + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("Payment process Document no Not found to  unpost")
            Else
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, clsCommon.myCstr(dt.Rows(0)("Area_Location_Code")), clsCommon.myCDate(dt.Rows(0)("Doc_Date")), trans)
            End If
            If clsCommon.myCDecimal(dt.Rows(0)("isPosted")) = 1 Then
                Throw New Exception("Payment process processed can't unpost it")
            End If
            If clsCommon.myCDecimal(dt.Rows(0)("isPrePosted")) = 0 Then
                Throw New Exception("Payment process should be posted to  unpost")
            End If

            Dim DocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_No from TSPL_BANK_ADVISE Where Payment_Process_Document_No='" + strDocNo + "'", trans))
            If clsCommon.myLen(DocNo) > 0 Then
                clsBankAdvise.ReverseAndUnpost(DocNo, trans)
                clsBankAdvise.deleteData(DocNo, trans)
            End If
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

            qry = "update TSPL_PAYMENT_PROCESS_HEAD set isPrePosted=0, Posting_Date=null where doc_no='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function CreateBankAdvise(ByVal strCode As String, ByVal strDate As DateTime) As Boolean
        Dim isNewEntry As Boolean = False
        Dim IsBankAdviseStartDate As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.BankAdviseRequired, clsFixedParameterCode.BankAdviseRequired, Nothing))
        If clsCommon.myLen(IsBankAdviseStartDate) <= 0 Then
            Dim obj As clsBankAdvise = New clsBankAdvise()
            obj.Document_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_No from TSPL_BANK_ADVISE Where Payment_Process_Document_No='" + strCode + "'"))
            obj.Document_Date = strDate
            obj.Payment_Process_Document_No = strCode
            obj.Remarks = "Auto"
            If obj.Document_No IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0 Then
                isNewEntry = False
            Else
                isNewEntry = True
            End If
            If clsBankAdvise.SaveData(obj, isNewEntry) Then
                If clsCommon.myLen(obj.Document_No) > 0 Then
                    clsBankAdvise.postData(obj.Document_No)
                End If
            End If
        End If
        Return True
    End Function

    Public Shared Function ReverseAndUnpost(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code,TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code,TSPL_PAYMENT_PROCESS_HEAD.Doc_Date from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + strDocNo + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.frmPaymentProcess, clsCommon.myCstr(dt.Rows(0)("Area_Location_Code")), clsCommon.myCDate(dt.Rows(0)("Doc_Date")), trans)

            End If

            If clsCommon.myLen(strDocNo) <= 0 Then
                Throw New Exception("Payment process Document no Not found to reverse and unpost")
            End If
            Dim qry As String = "select isPosted from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + strDocNo + "'"
            If clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry, trans)) <> 1 Then
                Throw New Exception("Payment process Document no should be posted to reverse and unpost")
            End If

            '-------MCC Sale Return adjustment
            qry = "Delete from TSPL_JOURNAL_DETAILS where Voucher_No in ( select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Sale_Doc_No=TSPL_Receipt_Adjustment_Header.Doc_No and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.AR_Invoice_No=TSPL_Receipt_Adjustment_Header.ARInvoiceNo and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No='" + strDocNo + "')))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Sale_Doc_No=TSPL_Receipt_Adjustment_Header.Doc_No and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.AR_Invoice_No=TSPL_Receipt_Adjustment_Header.ARInvoiceNo and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No='" + strDocNo + "')) "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_Receipt_Adjustment_Detail where Adjustment_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Sale_Doc_No=TSPL_Receipt_Adjustment_Header.Doc_No and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.AR_Invoice_No=TSPL_Receipt_Adjustment_Header.ARInvoiceNo and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No='" + strDocNo + "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Sale_Doc_No=TSPL_Receipt_Adjustment_Header.Doc_No and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.AR_Invoice_No=TSPL_Receipt_Adjustment_Header.ARInvoiceNo and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No='" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            '-------End of MCC Sale Return adjustment

            '-------Payment Entry
            qry = "Delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "')))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_PAYMENT_BANK_CHARGES_TAX where  Payment_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_PAYMENT_DETAIL_GST where  Payment_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_PAYMENT_DETAIL where Payment_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Delete from TSPL_PAYMENT_HEADER where Against_PP_Detail_No in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            '-------End of Payment Entry


            '---------Payable adjustment
            qry = "Delete from TSPL_JOURNAL_DETAILS where Voucher_No in ( select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Payment_Adjustment_Header  where Against_Payment_Process in ('" + strDocNo + "')))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Payment_Adjustment_Header where Against_Payment_Process in ('" + strDocNo + "') )"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from  TSPL_Payment_Adjustment_Detail where Adjustment_No in ( select Adjustment_No from TSPL_Payment_Adjustment_Header where Against_Payment_Process in ('" + strDocNo + "') )"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_Payment_Adjustment_Header where Against_Payment_Process in ('" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            '---------End of Payable adjustment

            '---------------Advance Payment
            qry = "Delete from TSPL_JOURNAL_DETAILS where Voucher_No in ( select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No_Advance_Payment in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" + strDocNo + "' )))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No_Advance_Payment in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" + strDocNo + "' ))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_PAYMENT_DETAIL where Payment_No in (select Payment_No from TSPL_PAYMENT_HEADER where Against_PP_Detail_No_Advance_Payment in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" + strDocNo + "' ))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_PAYMENT_HEADER where Against_PP_Detail_No_Advance_Payment in (select PP_Detail_No from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" + strDocNo + "' )"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            '---------------End of Advance Payment

            '----------- Receipt Adjustment 
            qry = "Delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE where TSPL_PAYMENT_PROCESS_MCC_SALE.Doc_No='" + strDocNo + "' and TSPL_Receipt_Adjustment_Header.Doc_No=TSPL_PAYMENT_PROCESS_MCC_SALE.Sale_Doc_No and TSPL_Receipt_Adjustment_Header.ARInvoiceNo=TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_No)))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_JOURNAL_MASTER where Source_Doc_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE where TSPL_PAYMENT_PROCESS_MCC_SALE.Doc_No='" + strDocNo + "' and TSPL_Receipt_Adjustment_Header.Doc_No=TSPL_PAYMENT_PROCESS_MCC_SALE.Sale_Doc_No and TSPL_Receipt_Adjustment_Header.ARInvoiceNo=TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_No))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_Receipt_Adjustment_Detail where Adjustment_No in ( select Adjustment_No from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE where TSPL_PAYMENT_PROCESS_MCC_SALE.Doc_No='" + strDocNo + "' and TSPL_Receipt_Adjustment_Header.Doc_No=TSPL_PAYMENT_PROCESS_MCC_SALE.Sale_Doc_No and TSPL_Receipt_Adjustment_Header.ARInvoiceNo=TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_No))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "Delete from TSPL_Receipt_Adjustment_Header where exists (select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE where TSPL_PAYMENT_PROCESS_MCC_SALE.Doc_No='" + strDocNo + "' and TSPL_Receipt_Adjustment_Header.Doc_No=TSPL_PAYMENT_PROCESS_MCC_SALE.Sale_Doc_No and TSPL_Receipt_Adjustment_Header.ARInvoiceNo=TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_No)"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            '----------- End of Receipt Adjustment 

            ''Update AP Invoice Balance Amont Type Invoice
            qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_Amt=Document_Total where Document_No in (select AP_Invoice_No from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            ''Update AP Invoice Balance Amont Type Debit/Credit Note/ItemIssue/ItemIssueRetrun
            qry = "update TSPL_VENDOR_INVOICE_HEAD  set Balance_Amt=xx.BalanceAmt from (" + Environment.NewLine +
            " select AP_Invoice_No,( Amount) as BalanceAmt from TSPL_PAYMENT_PROCESS_DEDUCTION where Doc_No='" + strDocNo + "'  " + Environment.NewLine +
            " union all " + Environment.NewLine +
            " select AP_Invoice_No,(Amount) as BalanceAmt from TSPL_PAYMENT_PROCESS_CREDIT_NOTE where Doc_No='" + strDocNo + "'  " + Environment.NewLine +
            "union all" + Environment.NewLine +
            "select AP_Invoice_No,(Amount) as BalanceAmt from TSPL_PAYMENT_PROCESS_ITEM_ISSUE where Doc_No='" + strDocNo + "'" + Environment.NewLine +
            "union all" + Environment.NewLine +
            "select AP_Invoice_No,(Amount) as BalanceAmt from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN  where Doc_No='" + strDocNo + "'" + Environment.NewLine +
            " )xx " + Environment.NewLine +
            "inner join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=xx.AP_Invoice_No"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            ''Update Head Load Credit note AP Invoice Balance Amont Type Invoice
            qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_amt=Document_Total where RefDocType = 'Milk_HE' and RefDocNo in (select Milk_Purchase_Invoice_No from TSPL_PAYMENT_PROCESS_INVOICE where doc_no='" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_amt=Document_Total where Document_No in (
select AP_Invoice_No from TSPL_PAYMENT_PROCESS_COMPULSORY where Doc_No='" + strDocNo + "'
union all
select AP_Invoice_No from TSPL_PAYMENT_PROCESS_SAVING where Doc_No='" + strDocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'Update Advance Payment entry Balance Amount
            qry = "update TSPL_PAYMENT_HEADER set Balance_Amt=Balance_Amt+xx.Amount_Knock_Off from (" + Environment.NewLine +
            "select  Payment_No,Amount_Knock_Off from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT where Doc_No='" + strDocNo + "'" + Environment.NewLine +
            ")xx inner join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=xx.Payment_No"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            ''UpdateAR Invoice balance amount for MCC Sale/MCC Sale Return
            qry = "update TSPL_Customer_Invoice_Head set Balance_Amt=Balance_Amt+xx.BalanceAmt from (" + Environment.NewLine +
            " select AR_Invoice_No,(Amount-Reduce_Deduc_Amt) as BalanceAmt  from TSPL_PAYMENT_PROCESS_MCC_SALE where Doc_No='" + strDocNo + "'" + Environment.NewLine +
            " union all " + Environment.NewLine +
            " select AR_Invoice_No,(Amount-Reduce_Deduc_Amt) as BalanceAmt  from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where Doc_No='" + strDocNo + "' " + Environment.NewLine +
            " )xx inner join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Document_No=xx.AR_Invoice_No"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "update TSPL_PAYMENT_PROCESS_HEAD set isPosted=0, Posting_Date=null where doc_no='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_PAYMENT_PROCESS_HEAD", "Doc_No", "TSPL_PAYMENT_PROCESS_DETAIL", "Doc_No", trans)

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function DeleteWithVSPBill(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(strDocNo) <= 0 Then
                Throw New Exception("Payment process Document no Not found to Delete with VSP Bill")
            End If
            Dim qry As String = "select isPosted,From_Date,To_Date from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + strDocNo + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.myCDecimal(dt.Rows(0)("isPosted")) <> 0 Then
                    Throw New Exception("Payment process Document no should be Unposted to Delete with VSP Bill")
                End If
            Else
                Throw New Exception("Document No [" + strDocNo + "] not found")
            End If


            qry = "select Location_Code from TSPL_LOCATION_MASTER where Loc_Segment_Code=(select Loc_Seg_Code from TSPL_PAYMENT_PROCESS_HEAD where Doc_No='" + strDocNo + "') and  Rejected_Type='N' and Location_Category='MCC' "
            Dim strMCCCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))

            If clsCommon.myLen(strMCCCode) <= 0 Then
                Throw New Exception("MCC Code not found")
            End If
            deleteData(strDocNo, trans)
            DeleteOnlyBill(strMCCCode, Nothing, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function DeleteOnlyBill(ByVal strMCCCode As String, ByVal DocDate As DateTime?) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            DeleteOnlyBill(strMCCCode, DocDate, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetNoOfDoc(ByVal strMCCCode As String, ByVal DocDate As DateTime?) As String
        Dim strWhr As String = "(select DOC_CODE from TSPL_MILK_PURCHASE_INVOICE_HEAD where TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE='" + strMCCCode + "' "
        If DocDate IsNot Nothing Then
            strWhr += " and TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(DocDate), "dd/MMM/yyyy hh:mm:ss tt") + "' and  TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(DocDate), "dd/MMM/yyyy hh:mm:ss tt") + "'"
        End If
        strWhr += " and DOC_CODE not in (select Milk_Purchase_Invoice_No from TSPL_PAYMENT_PROCESS_DETAIL))"
        Return strWhr
    End Function
    Public Shared Function DeleteOnlyBill(ByVal strMCCCode As String, ByVal DocDate As DateTime?, ByVal trans As SqlTransaction) As Boolean
        Dim strWhr As String = GetNoOfDoc(strMCCCode, DocDate)

        Dim qry As String = " select min(TSPL_MILK_SRN_HEAD.DOC_DATE) as MinDOC_DATE,max(TSPL_MILK_SRN_HEAD.DOC_DATE) as MaxDOC_DATE from " &
        "TSPL_MILK_SRN_HEAD " &
        " left outer join TSPL_MILK_PURCHASE_INVOICE_DETAIL on TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE=TSPL_MILK_SRN_HEAD.DOC_CODE " &
        " where TSPL_MILK_PURCHASE_INVOICE_DETAIL.DOC_CODE in  " + strWhr + "  "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            qry = "select Doc_No from tspl_provision_entry where Prov_type = 'Chilling Charge' and Prog_Code = '" + clsUserMgtCode.MilkVSPPayment + "' and Loc_Code = '" + strMCCCode + "' and convert(date, Doc_Date,103)>='" + clsCommon.GetPrintDate(dt.Rows(0)("MinDOC_DATE"), "dd/MMM/yyyy") + "' and convert(date, Doc_Date,103)<='" + clsCommon.GetPrintDate(dt.Rows(0)("MaxDOC_DATE"), "dd/MMM/yyyy") + "'"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    clsProvisionEntry.ReverseAndUnpost(clsCommon.myCstr(dr("Doc_No")), trans)
                    clsProvisionEntry.deleteData(clsCommon.myCstr(dr("Doc_No")), trans)
                Next
            End If
        End If

        qry = "select Payment_No from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Entry_Desc = 'Apply document for Asset Lost'  and len(isnull(TSPL_PAYMENT_HEADER.Applied_Payment,''))>0 and Reference in " + strWhr + " "
        dt = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                clsPaymentHeader.ReverseAndUnpost(clsCommon.myCstr(dr("Payment_No")), trans)
                clsPaymentHeader.fundelete("", dr("Payment_No"), "", trans)
            Next
        End If


        Dim strRefDocType As String = "('DED-MAP','TIP-DED','VSP-COM','VSP-CMP','VSP-QLT','VSP-PVK','VSP-DIT','PRO-VFC','PRO-VFD','NCM-DED','CM-DED','ASL-DED','PRO-LCS','PRO-STD','OWD-CRE','CHC-CRE','CRE-DHL','OWD-CRD','OWD-DBT','DCS-ADD','DCS-DED','DCS-QAT','DCS-LYT','VSP-NGT','OWD-RJM')"
        'Delete deduction Entry

        qry = "delete from TSPL_INVENTORY_MOVEMENT_NEW where Trans_Type='IC-AD' and source_doc_no in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + "))"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_ADJUSTMENT_DETAIL where Adjustment_No in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + "))"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_ADJUSTMENT_HEADER where Adjustment_No in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + "))"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_JOURNAL_DETAILS  where Voucher_No in (select voucher_no from TSPL_JOURNAL_MASTER where Source_Doc_No in   (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + "  and RefDocType in " + strRefDocType + "))"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No in   (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_REMITTANCE where Document_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhr + " and RefDocType in " + strRefDocType + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'End of Delete deduction Entry

        'Delete MilkReject Entry
        'Dim strWhrReject As String = "( select DOC_CODE from TSPL_MILK_REJECT_HEAD  " + Environment.NewLine +
        '"inner Join (  " + Environment.NewLine +
        '"select MCC_Code,max(FromDate) as FromDate,max(ToDate) as ToDate from (" + Environment.NewLine +
        '"select TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code,DATEADD(day,((-1*TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE)+1), TSPL_MILK_PURCHASE_INVOICE_HEAD.Doc_Date) as FromDate ,TSPL_MILK_PURCHASE_INVOICE_HEAD.Doc_Date as ToDate" + Environment.NewLine +
        '"from TSPL_MILK_PURCHASE_INVOICE_HEAD " + Environment.NewLine +
        '"left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_code=TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_code" + Environment.NewLine +
        '"left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle" + Environment.NewLine +
        '"where TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE in " + strWhr + "" + Environment.NewLine +
        '")x Group by MCC_Code" + Environment.NewLine +
        '")Inv on Inv.MCC_Code=TSPL_MILK_REJECT_HEAD.MCC_CODE and convert(date, TSPL_MILK_REJECT_HEAD.DOC_DATE,103)>=inv.FromDate and convert(date, TSPL_MILK_REJECT_HEAD.DOC_DATE,103)<=inv.ToDate)"

        'qry = "delete from TSPL_INVENTORY_MOVEMENT_NEW where Trans_Type='IC-AD' and source_doc_no in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ')))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'qry = "delete from TSPL_ADJUSTMENT_DETAIL where Adjustment_No in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ')))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'qry = "delete from TSPL_ADJUSTMENT_HEADER where Adjustment_No in ( select Adjustment_No from TSPL_ADJUSTMENT_HEADER where Against_AP_Invoice_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ')))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)


        'qry = "delete from TSPL_JOURNAL_DETAILS  where Voucher_No in (select voucher_no from TSPL_JOURNAL_MASTER where Source_Doc_No in   (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + "  and RefDocType in ('MILK-REJ')))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No in   (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ'))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'qry = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ'))"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'qry = "delete from TSPL_VENDOR_INVOICE_HEAD where RefDocNo in " + strWhrReject + " and RefDocType in ('MILK-REJ')"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        'End of Delete deduction Entry



        Dim settVSPDayWiseIncentiveAtSRN As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.VSPDayWiseIncentiveAtSRN, clsFixedParameterCode.VSPDayWiseIncentiveAtSRN, trans)) > 0)
        qry = "update TSPL_MILK_SRN_DETAIL set VSP_Commission_Apply=0, VSP_Deduction_Apply=0"
        If Not settVSPDayWiseIncentiveAtSRN Then
            qry += " ,VSP_Day_Wise_Incentive=0,VSP_Day_Wise_Incentive_Rate=0 "
        End If
        qry += ",Farmer_Pro_Code=null,VSP_Mapping_Code_Day_Wise_Incentive=null,VSP_Mapping_Code=null  where DOC_CODE in (select TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE from TSPL_MILK_PURCHASE_INVOICE_DETAIL where DOC_CODE in " + strWhr + ")  "
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "update TSPL_MILK_SRN_Head set Is_Incentive_Created='N' where DOC_CODE in (select TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE from TSPL_MILK_PURCHASE_INVOICE_DETAIL where DOC_CODE in " + strWhr + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)


        qry = "delete from TSPL_REMITTANCE where Document_No in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where Against_MillkPurchaseInvoice_No in " + strWhr + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No  in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where Against_MillkPurchaseInvoice_No in " + strWhr + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No  in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No  in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where Against_MillkPurchaseInvoice_No in " + strWhr + "))"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = " delete from TSPL_JOURNAL_MASTER where Source_Doc_No  in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where Against_MillkPurchaseInvoice_No in " + strWhr + ")"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_VENDOR_INVOICE_HEAD where Against_MillkPurchaseInvoice_No in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_DETAIL where DOC_CODE in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_PI_REMITTANCE where  Document_No in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_MP_COLLEC_SUMMARY where  DOC_CODE in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)


        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_INCENTIVEDETAIL where MILK_DOC_Code in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_PRO_LOSS where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_OWN_BMC where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_OWN_BMC_REJECT where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_OWN_BMC_EXPANSE where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_CHILLING_CHARGES where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_DAY_WISE_HEAD_LOAD where InvoiceNo in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_MILK_PURCHASE_INVOICE_HEAD  where DOC_CODE in " + strWhr + ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Return True
    End Function

    '========================================
    'Public Shared Function PaymentProcessDrCrPrint(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal strMCC_Code As String) As Boolean
    'Public Shared Function PaymentProcessDrCrPrint(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal strMCC_Code As String) As Boolean
    Public Shared Function PaymentProcessDrCrPrint(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal strMCC_Code As String, ByVal strVSPCode As String, ByVal CycleWiseData As Boolean) As Boolean
        Try
            Dim sQuery As String = ""
            Dim settMaxFATPerLimit As Decimal = 0
            Dim settMaxSNFPerLimit As Decimal = 0
            Dim dtDebit As New DataTable
            Dim dtCredit As New DataTable
            Dim dtMilkType As New DataTable
            Dim AreaWiseBilling As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = 1)
            settMaxFATPerLimit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxFATPerLimitforReport, clsFixedParameterCode.MaxFATPerLimitforReport, Nothing))
            settMaxSNFPerLimit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxSNFPerLimitforReport, clsFixedParameterCode.MaxSNFPerLimitforReport, Nothing))

            Dim dt As New DataTable
            Dim Flag As Boolean = False
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse
                clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Flag = True
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                sQuery = " select Description,(Amount)Amount from ( select TSPL_DEDUCTION_MASTER.Description as Description,"
            Else
                sQuery = " select Description,Sum(Amount)Amount from ( select TSPL_DEDUCTION_MASTER.Description as Description,"
            End If

            If Flag Then
                sQuery += " sum(TSPL_PAYMENT_PROCESS_DEDUCTION.Amount-TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt) as Amount "
            Else
                sQuery += " sum(TSPL_VENDOR_INVOICE_DETAIL.Amount) as Amount"
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                sQuery += " ,max(TSPL_DEDUCTION_MASTER.SNo)SNo "
            End If
            sQuery += " from TSPL_PAYMENT_PROCESS_DEDUCTION
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No "

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                sQuery += " left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_VENDOR_INVOICE_DETAIL.Document_No "
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                sQuery += " left outer join ( select code, Description,SNo  from TSPL_DCS_ADDITION_DEDUCTION
union 
select  Code , Description,Sequence_No AS SNo from TSPL_DEDUCTION_MASTER)  as TSPL_DEDUCTION_MASTER on ( TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
where TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in (" + strDocNo + ") "

                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE In (" + strVSPCode + ")"
                End If


                sQuery += " and Len(TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE ) > 0 "
            Else
                sQuery += " left outer join ( select code, Description  from TSPL_DCS_ADDITION_DEDUCTION
union 
select  Code , Description from TSPL_DEDUCTION_MASTER) as TSPL_DEDUCTION_MASTER on ( TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
where TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in (" + strDocNo + ")  "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE In (" + strVSPCode + ")"
                End If

                sQuery += " and Len(TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE ) > 0 "
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                sQuery += " and TSPL_VENDOR_INVOICE_HEAD.Saving=0 "
            End If

            sQuery += " group by TSPL_DEDUCTION_MASTER.Description  
union all"

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                sQuery += " Select * from(select 'TDS' as Description,isnull(sum(isnull(TSPL_PAYMENT_PROCESS_DETAIL.TDS_Amount,0)),0) as Amount,99 as SNo from TSPL_PAYMENT_PROCESS_DETAIL 
                            where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ")"
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                End If

                sQuery += " )x where Amount>0 
                             union all
                            select TSPL_DEDUCTION_MASTER.Description as Description,Sum(TSPL_PAYMENT_PROCESS_MCC_SALE.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE.Reduce_Deduc_Amt) as Amount,max(Sequence_No) AS SNo
                            from TSPL_PAYMENT_PROCESS_MCC_SALE 
                            left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE
                            left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no
                            left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
                            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
                            left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
                            where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in (" + strDocNo + ") "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE In (" + strVSPCode + ")"
                End If
                sQuery += " group by TSPL_DEDUCTION_MASTER.Description
                            
)xx order by  xx.SNo "
                '')xx group by xx.Description order by  xx.SNo "
            Else
                sQuery += " Select * from(select 'TDS' as Description,isnull(sum(isnull(TSPL_PAYMENT_PROCESS_DETAIL.TDS_Amount,0)),0) as Amount from TSPL_PAYMENT_PROCESS_DETAIL 
                            where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                End If

                sQuery += " )x where Amount>0 
                            union all
                            select TSPL_DEDUCTION_MASTER.Description as Description,Sum(TSPL_PAYMENT_PROCESS_MCC_SALE.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE.Reduce_Deduc_Amt) as Amount
                            from TSPL_PAYMENT_PROCESS_MCC_SALE 
                            left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE
                            left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no
                            left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
                            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
                            left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
                            where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in (" + strDocNo + ") "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE In (" + strVSPCode + ")"
                End If
                sQuery += " group by TSPL_DEDUCTION_MASTER.Description
                            )xx 
                            group by xx.Description order by  xx.Description  "
            End If

            dtDebit = clsDBFuncationality.GetDataTable(sQuery)


            Dim strHeadLoadColumnName As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.HeadLoadDescriptionInPaymentProcessPrint, clsFixedParameterCode.HeadLoadDescriptionInPaymentProcessPrint, Nothing))
            sQuery = "  Select 
 Description,
 sum(Amount) as Amount from (
Select  Doc_No, Description,Sequence_No,Code, sum (Amount) As Amount from ( 
Select  TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No,TSPL_DEDUCTION_MASTER.Description,TSPL_DEDUCTION_MASTER.Sequence_No,TSPL_DEDUCTION_MASTER.Code,sum(TSPL_VENDOR_INVOICE_DETAIL.Amount) As Amount 
From TSPL_PAYMENT_PROCESS_CREDIT_NOTE
Left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
Left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
Left outer join (  Select code, 0 As Sequence_No, Description  from TSPL_DCS_ADDITION_DEDUCTION --where len(isnull(MappingCode,''))<=0
                            union
                            Select  Code , Sequence_No,Description from TSPL_DEDUCTION_MASTER
union
                                Select 'DCS-LYT'  Code ,0 as Sequence_No,'Loyalty' as  Description 
union
                                    Select 'DCS-QAT'  Code ,0 as Sequence_No,'QAP' as  Description 
)  TSPL_DEDUCTION_MASTER on (TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode Or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction Or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_HEAD.RefDocType)
where TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE In (" + strVSPCode + ")"
            End If
            sQuery += " And len(TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE) > 0
Group by TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No, TSPL_DEDUCTION_MASTER.Description, TSPL_DEDUCTION_MASTER.Code, Sequence_No

union all 
                            select TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No,TSPL_DEDUCTION_MASTER.Description as Description,88 as Sequence_No,'' AS Code ,Sum(TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Reduce_Deduc_Amt) as Amount
                            from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN 
                            left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Customer_CODE
                            left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no
                            left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
                            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
                            left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
                            where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.doc_no  in (" + strDocNo + ") 
							Group by TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No, TSPL_DEDUCTION_MASTER.Description, TSPL_DEDUCTION_MASTER.Code, Sequence_No "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") <> CompairStringResult.Equal And clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") <> CompairStringResult.Equal Then
                sQuery += " Union
                            Select  TT.Doc_No,coalesce (mapping.mmDescription, TT.Description) ,0 As sequence_no, coalesce (mapping.mmCode, TT.Code) As Code,TT.Amount
                            FROM(
                            select TSPL_PAYMENT_PROCESS_SAVING.Doc_No, TSPL_DEDUCTION_MASTER.Description, TSPL_DEDUCTION_MASTER.Code, sum(TSPL_VENDOR_INVOICE_DETAIL.Amount) as Amount from TSPL_PAYMENT_PROCESS_SAVING
                            Left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
                            Left outer join ( Select Code, 0 As Sequence_No, Description  from TSPL_DCS_ADDITION_DEDUCTION 
                            union
                                            Select  Code ,Sequence_No ,Description  from TSPL_DEDUCTION_MASTER
                            )  TSPL_DEDUCTION_MASTER on (TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode Or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
                            where TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ") "
                'If CycleWiseData = True Then
                '    sQuery += " and TSPL_VENDOR_INVOICE_DETAIL.Customer_CODE In (" + strVSPCode + ")"
                'End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    sQuery += " And TSPL_DEDUCTION_MASTER.Code !='PDP' "
                End If
                sQuery += " group by TSPL_PAYMENT_PROCESS_SAVING.Doc_No,TSPL_DEDUCTION_MASTER.Description ,TSPL_DEDUCTION_MASTER.Code,TSPL_DEDUCTION_MASTER.Sequence_No 
                            )TT
                            left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
                            left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION on  DEDUCTION.Code=MAPPING.MappingCode
                            WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code "
            End If

            '            sQuery += "  Select Case TT.Doc_No,coalesce (mapping.mmDescription, TT.Description) ,0 As sequence_no, coalesce (mapping.mmCode, TT.Code) As Code,TT.Amount
            'FROM(
            'select TSPL_PAYMENT_PROCESS_SAVING.Doc_No, TSPL_DEDUCTION_MASTER.Description, TSPL_DEDUCTION_MASTER.Code, sum(TSPL_VENDOR_INVOICE_DETAIL.Amount) as Amount from TSPL_PAYMENT_PROCESS_SAVING
            'Left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
            'Left outer join ( Select Code, 0 As Sequence_No, Description  from TSPL_DCS_ADDITION_DEDUCTION 
            'union
            '                    Select Case Code ,Sequence_No ,Description  from TSPL_DEDUCTION_MASTER
            ')  TSPL_DEDUCTION_MASTER on (TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode Or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
            'where TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ")"

            '                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            '                sQuery += " And TSPL_DEDUCTION_MASTER.Code !='PDP' "
            '            End If

            '            sQuery += " group by TSPL_PAYMENT_PROCESS_SAVING.Doc_No,TSPL_DEDUCTION_MASTER.Description ,TSPL_DEDUCTION_MASTER.Code,TSPL_DEDUCTION_MASTER.Sequence_No 
            ')TT
            'left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
            'left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION on  DEDUCTION.Code=MAPPING.MappingCode
            'WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code"

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                sQuery += " ) Final group by  Doc_No, Description ,Code,Sequence_No
UNION ALL
 
select 
TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,'" + strHeadLoadColumnName + "' as Description,0 as Sequence_No,'' AS Code,TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount as Amount from 
TSPL_PAYMENT_PROCESS_DETAIL INNER JOIN
TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
where Document_Type='C' and RefDocType='Milk_HE'  and TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                End If
                sQuery += " )DD
group by code,Sequence_No,Description order by  Sequence_No,code asc "
            Else

                sQuery += " Union
SELECT TT.Doc_No,coalesce (mapping.mmDescription, TT.Description) ,0 as Sequence_No, coalesce (mapping.mmCode, TT.Code) AS Code,TT.Amount
FROM (
select TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No,TSPL_DEDUCTION_MASTER.Description ,TSPL_DEDUCTION_MASTER.Code
,sum(TSPL_VENDOR_INVOICE_DETAIL.Amount) as Amount from TSPL_PAYMENT_PROCESS_COMPULSORY
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No"

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    sQuery += " left outer join ( select Code, 0 as Sequence_No,Description,0 as Is_Transfer_To_Saving    from TSPL_DCS_ADDITION_DEDUCTION 
                    union 
                    select  Code ,Sequence_No, Description,Is_Transfer_To_Saving    from TSPL_DEDUCTION_MASTER
                    )  TSPL_DEDUCTION_MASTER on (TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
                    where TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No in (" + strDocNo + ") and isnull(TSPL_DEDUCTION_MASTER.Is_Transfer_To_Saving,0)=0 "
                Else
                    sQuery += " left outer join ( select Code, 0 as Sequence_No,Description  from TSPL_DCS_ADDITION_DEDUCTION 
                    union 
                    select  Code ,Sequence_No, Description  from TSPL_DEDUCTION_MASTER
                    )  TSPL_DEDUCTION_MASTER on (TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
                    where TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No in (" + strDocNo + ") "
                End If


                sQuery += "    group by TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No,TSPL_DEDUCTION_MASTER.Description ,TSPL_DEDUCTION_MASTER.Code ,TSPL_DEDUCTION_MASTER.Sequence_No
)TT left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
                     left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION
                     on  DEDUCTION.Code=MAPPING.MappingCode
                     WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code
                     ) Final group by  Doc_No, Description ,Code,Sequence_No
UNION ALL

 
select 
TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,'" + strHeadLoadColumnName + "' as Description,0 as Sequence_No,'' AS Code,TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount as Amount from 
TSPL_PAYMENT_PROCESS_DETAIL INNER JOIN
TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
where Document_Type='C' and RefDocType='Milk_HE'  and TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ")"
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                End If

                'sQuery += " union all 
                '            select TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No,TSPL_DEDUCTION_MASTER.Description as Description,88 as Sequence_No,'' AS Code ,(TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Reduce_Deduc_Amt) as Amount
                '            from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN 
                '            left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Customer_CODE
                '            left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no
                '            left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
                '            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
                '            left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
                '            where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.doc_no  in (" + strDocNo + ")  "
                'If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                '    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                'End If

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    sQuery += " UNION ALL 
                                select  TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,'Milk Payment' as Description,0 as Sequence_No,'' AS Code,"
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                        sQuery += " (TSPL_PAYMENT_PROCESS_DETAIL.SRN_Net_Amount) as Amount  "
                    Else
                        sQuery += " (TSPL_PAYMENT_PROCESS_DETAIL.Milk_Amount) as Amount  "
                    End If


                    sQuery += "   from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") "
                    If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                        sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code In (" + strVSPCode + ")"
                    End If
                End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
                    sQuery += " UNION ALL 
                                 select TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo as Doc_No,TSPL_DCS_ADDITION_DEDUCTION.Description,0 as Sequence_No,'' AS Code ,TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE
			                    left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.DCS_Addition_Deduction
								left outer join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo 
								Where convert(date, TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_Date,103)>= '" + clsCommon.GetPrintDate(CycleFromDate) + "'   
                                    and convert(date,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_Date,103)<= '" + clsCommon.GetPrintDate(CycleToDate) + "' 
									and TSPL_DCS_ADDITION_DEDUCTION.Description='PURCHASE EXP.' )DD
                                    group by Description order by  Description asc "
                Else
                    sQuery += " )DD
group by code,Sequence_No,Description order by  Description desc"
                End If
            End If

            dtCredit = clsDBFuncationality.GetDataTable(sQuery)

            sQuery = "select 
Comp_Code,Comp_Name,max(MCC_NAME) as MCC_NAME,Regn_No
,max(Milk_Qty) as Milk_Qty,max(Milk_Amount) as Milk_Amount,max(Date1) as Date1,max(Date2) as Date2,max(comp_add1) as comp_add1,max(comp_add2) as comp_add2,max(comp_add3) as comp_add3,max(comp_Fax) as comp_Fax,max(comp_Email) as comp_Email,max(CompPhone) as CompPhone,max(Pincode) as Pincode,max(Tcan_No) as Tcan_No,max(Doc_Date) as Doc_Date,min(from_date) as from_date, max(to_date) as to_date,max(Milk_Qty) as Milk_Qty,max(Milk_Amount) as Milk_Amount,max(DebitAmount) as DebitAmount,max(CreditAmount) as CreditAmount,max(SaleReturnAmount)SaleReturnAmount,max(Hold_Payable_Amount) as Hold_Payable_Amount,sum(HeadLoadAmount) as HeadLoadAmount "

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                sQuery += " , sum(TotalOutStandingAmt) as TotalOutStandingAmt "
            End If

            sQuery += "  from (select TSPL_PAYMENT_PROCESS_HEAD.doc_no
                        ,TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name,"
            If AreaWiseBilling = True Then
                sQuery += " xxxSetLocation.Location_Desc as MCC_NAME, "
            Else
                sQuery += " TSPL_MCC_MASTER.MCC_NAME,"
            End If
            sQuery += "TSPL_COMPANY_MASTER.Regn_No,'" + CycleFromDate + " ' as Date1,'" + CycleToDate + "' as Date2
                    ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , TSPL_COMPANY_MASTER.Add2 as  comp_add2 
                    ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , cast(TSPL_COMPANY_MASTER.logo_img as image) as logo_img,tspl_company_master.Pincode,tspl_company_master.Tcan_No
                                        ,convert(varchar(12),TSPL_PAYMENT_PROCESS_HEAD.Doc_Date,103) as Doc_Date
                    ,convert(varchar(12),TSPL_PAYMENT_PROCESS_HEAD.from_date,103) as from_date
                    ,convert(varchar(12),TSPL_PAYMENT_PROCESS_HEAD.to_date,103) as to_date"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                sQuery += ",(Select SUM(case when TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code='99' then (ISNULL(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount, 0))end)
					From TSPL_PAYMENT_PROCESS_DETAIL Where Doc_No In (" + strDocNo + ")"
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
                End If
                sQuery += " )AS TotalOutStandingAmt "
            End If

            sQuery += "  ,(SELECT sum(TSPL_PAYMENT_PROCESS_DETAIL.Milk_Qty) FROM TSPL_PAYMENT_PROCESS_DETAIL WHERE Doc_no in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS Milk_Qty
                    ,(SELECT sum(TSPL_PAYMENT_PROCESS_DETAIL.Milk_Amount) FROM TSPL_PAYMENT_PROCESS_DETAIL WHERE Doc_no in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS Milk_Amount                 
                    ,(SELECT sum(TSPL_PAYMENT_PROCESS_DEDUCTION.Amount) FROM TSPL_PAYMENT_PROCESS_DEDUCTION WHERE TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_no in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS DebitAmount
                    ,(SELECT sum(TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount) FROM TSPL_PAYMENT_PROCESS_CREDIT_NOTE WHERE TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS CreditAmount
            ,(SELECT sum(TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Amount) FROM TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN WHERE TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Customer_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS SaleReturnAmount
                    ,(SELECT sum(Payable_Amount) FROM TSPL_PAYMENT_PROCESS_DETAIL WHERE TSPL_PAYMENT_PROCESS_DETAIL.is_Hold_Payment_Process=1 and TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ")"
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
            End If
            sQuery += ") AS Hold_Payable_Amount
                    from TSPL_PAYMENT_PROCESS_HEAD                    
                    left outer join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code =TSPL_PAYMENT_PROCESS_HEAD.Comp_Code"
            If AreaWiseBilling = True Then
                sQuery += "	  Left Outer Join( 
					  SELECT TSPL_LOCATION_MASTER.Location_Desc, TSPL_LOCATION_MASTER.Location_Code,TSPL_MCC_MASTER.MCC_Code FROM TSPL_PAYMENT_PROCESS_HEAD 
										  LEFT OUTER JOIN TSPL_MCC_MASTER ON TSPL_MCC_MASTER.Area_Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
										  LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_MCC_MASTER.Area_Location_Code
					   ) xxxSetLocation On xxxSetLocation.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code"
            Else
                sQuery += "  Left outer join TSPL_MCC_MASTER On TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected"
            End If

            sQuery += " where TSPL_PAYMENT_PROCESS_HEAD.doc_no in (" + strDocNo + "))XXY
					 left outer join 
					 (select 
                        TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,'" + strHeadLoadColumnName + "' as Description,'' AS Code,sum(TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount) as HeadLoadAmount from 
                        TSPL_PAYMENT_PROCESS_DETAIL INNER JOIN
                        TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
                        where Document_Type='C' and RefDocType='Milk_HE'  and TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") "
            If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                sQuery += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
            End If
            sQuery += " group by 
                        TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,Description)final1  on final1.Doc_No=xxy.doc_no "

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                sQuery += " left outer join (
                         select TSPL_PAYMENT_PROCESS_DETAIL.Doc_No,Sum(TSPL_PAYMENT_PROCESS_DETAIL.Reduce_Deduc_Amt) As 'TotalOutStandingAmt' from TSPL_PAYMENT_PROCESS_DETAIL  
                         left outer join (select DOC_CODE,cast( sum(FATKg) as decimal(18,3)) as FATKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(FATKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as FATPer ,
                         cast( sum(SNFKg) as decimal(18,3)) as SNFKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(SNFKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as SNFPer  
                         from (select DOC_CODE, ACC_Qty,FAT_PER,SNF_PER, cast(ACC_Qty*FAT_PER/100 as decimal(18,2)) as FATKg, cast( ACC_Qty*SNF_PER/100 as decimal(18,2)) as SNFKg 
                         from TSPL_MILK_PURCHASE_INVOICE_DETAIL )xx group by DOC_CODE ) as TabFATSNFDetail on TabFATSNFDetail.DOC_CODE=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
                         where TSPL_PAYMENT_PROCESS_DETAIL.doc_no in (" + strDocNo + ") Group By TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                         ) As TotalOutStanding On TotalOutStanding.Doc_No=final1.Doc_No"
            End If
            sQuery += " group by Comp_Code,Comp_Name,Regn_No "
            dt = clsDBFuncationality.GetDataTable(sQuery)


            'Dim BaseQty As String = clsPaymentProcessHead.Load_Report_Paymnet_RCDF_BaseQuery1(strDocNo, CycleFromDate, CycleToDate, strLoc, strVSPCode, "", "", "")
            Dim BaseQty As String = ""
            If CycleWiseData = True Then
                BaseQty = clsPaymentProcessHead.Load_Report_Paymnet_RCDF_BaseQuery1(strDocNo, CycleFromDate, CycleToDate, strLoc, "", "", "", strVSPCode, True)
            Else
                BaseQty = clsPaymentProcessHead.Load_Report_Paymnet_RCDF_BaseQuery1(strDocNo, CycleFromDate, CycleToDate, strLoc, "", "", "", "", False)
            End If
            'Dim BaseQty As String = clsPaymentProcessHead.Load_Report_Paymnet_RCDF_BaseQuery1(strDocNo, CycleFromDate, CycleToDate, strLoc, "", "", "", "", False)
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then cast( sum (FATQTY) * 100 /sum( Qty) as decimal(18,2))  as FATPer
            If True Then
                Dim qry As String = "select 'SWEET' as Code,'SWEET' as Name union all select 'GOODFAIL' as Code,'GOODFAIL' as Name union all select Code,Description from TSPL_MILK_REJECT_TYPE"
                Dim dtRejType As DataTable = clsDBFuncationality.GetDataTable(qry)

                sQuery = " select QBD,sum( Qty) as Qty, CASE WHEN SUM(Qty) = 0 THEN 0 ELSE CAST(SUM(FATQTY) * 100 / SUM(Qty) AS DECIMAL(18,2)) END AS FATPer,
                           CASE WHEN SUM(Qty) = 0 THEN 0 ELSE CAST(SUM(SNFQTY) * 100 / SUM(Qty) AS DECIMAL(18,2)) End as SNFPer, "

                'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                '    sQuery += " sum(isnull(SRN_NET_AMOUNT,0)) AS Amt,round(sum (FATQTY),2,1)  as FATKG,round(sum (SNFQTY),2,1) as SNFKG "
                'Else
                sQuery += " sum(isnull(SRN_NET_AMOUNT,0))+(sum(isnull(PPSRN_RO_Amount,0))*-1) AS Amt,round(sum (FATQTY),2,1)  as FATKG,round(sum (SNFQTY),2,1) as SNFKG "
                'End If
                If AreaWiseBilling = True Then
                    sQuery += " ,MAX(Location_Desc)Area "
                End If
                sQuery += " from ( " + BaseQty + " ) XXXX group by QBD "
                dtMilkType = clsDBFuncationality.GetDataTable(sQuery)
                For Each drRejType As DataRow In dtRejType.Rows
                    dt.Columns.Add("" + clsCommon.myCstr(drRejType("Code")) + "Qty" + "", GetType(Decimal))
                    dt.Columns.Add("" + clsCommon.myCstr(drRejType("Code")) + "Amt" + "", GetType(Decimal))
                    For Each drMilkType As DataRow In dtMilkType.Rows
                        If clsCommon.CompairString(clsCommon.myCstr(drRejType("Code")), clsCommon.myCstr(drMilkType("QBD"))) = CompairStringResult.Equal Then
                            dt.Rows(0)(clsCommon.myCstr("" + drRejType("Code")) + "Qty" + "") = clsCommon.myCDecimal(drMilkType("Qty"))
                            dt.Rows(0)(clsCommon.myCstr("" + drRejType("Code")) + "Amt" + "") = clsCommon.myCDecimal(drMilkType("Amt"))
                            Exit For
                        End If
                    Next
                Next
                dt.AcceptChanges()
            End If


            sQuery = " select CowBuffalo_Type,sum( Qty) as Qty,  CASE WHEN SUM(Qty) = 0 THEN 0 ELSE CAST(SUM(FATQTY) * 100 / SUM(Qty) AS DECIMAL(18,2)) END AS FATPer,
                           CASE WHEN SUM(Qty) = 0 THEN 0 ELSE CAST(SUM(SNFQTY) * 100 / SUM(Qty) AS DECIMAL(18,2)) End as SNFPer, "

            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            '    sQuery += " sum(SRN_NET_AMOUNT) As Amt, "
            'Else
            sQuery += " sum(SRN_NET_AMOUNT)+(sum(PPSRN_RO_Amount)*-1) As Amt, "
            'End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                sQuery += " (sum (FATQTY))  As FATKG,(sum (SNFQTY)) As SNFKG "
            Else
                sQuery += " round(sum (FATQTY),2,1)  As FATKG,round(sum (SNFQTY),2,1) As SNFKG  "
            End If
            'If AreaWiseBilling = True Then
            '    sQuery += "  ,MAX(Location_Desc) As MCC_NAME"
            'End If
            sQuery += " from ( " + BaseQty + " ) XXXX group by CowBuffalo_Type"
            dtMilkType = clsDBFuncationality.GetDataTable(sQuery)

            Dim dtAdvice As DataTable = Nothing
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                Dim ssql = "Select ROW_NUMBER() over ( order by GRPColumn) As SNO , * from ( Select max(CycleRange) As CycleRange, max(GRPColumn) As GRPColumn,max(Comp_Name) As Comp_Name,max(Comp_address) As Comp_address, max(From_Date) As From_Date,max(GSTReg_No) As GSTReg_No,max(Fiscal_Name) As Fiscal_Name,max(CycleNo) As CycleNo,max(Date_Range) As Date_Range,Bank_Code,Branch_Name,max(Bank_Code_Desc) As Bank_Code_Desc, max (Payee_Joint_IFSC_Code) As Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) As Payee_Joint_Account_No ,sum(Payable_Amount) As Payable_Amount,sum(Payable_Amount1)Amt
,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
from(
select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code as GRPColumn,TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
TSPL_COMPANY_MASTER.Comp_Name
,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc,  TSPL_Fiscal_Year_Master.Fiscal_Name
,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Code as Branch_Name,TSPL_Vendor_MASTER.Bank_Code as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No, (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))  as Payable_Amount ,
Round(isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount, 0), 0)-Round (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0)  As Payable_Amount1   from TSPL_PAYMENT_PROCESS_DETAIL 
Left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
Left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                Left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
Left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date And TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
Left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
Left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected

 Left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code And  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' 
                Left outer join TSPL_PAYMENT_CYCLE_GENERATED On convert(Date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date, 103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) And convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)  And TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code  
where  TSPL_PAYMENT_PROCESS_HEAD.From_Date >= Convert(Date, ('" + CycleFromDate + "'),103) and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<=convert(date,('" + CycleToDate + "'),103)    and  TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + strMCC_Code + "' And TSPL_MCC_MASTER.MCC_Code='" + strMCC_Code + "'
 )xxx group by Bank_Code, Branch_Name)xxxx order by GRPColumn "
                '                Dim ssql = "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No,max(payamt) as Amount ,max(Payable_Amount) as Payable_Amount
                ',max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
                'from (
                'select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code as GRPColumn,TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
                'TSPL_COMPANY_MASTER.Comp_Name
                ',TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
                ',case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
                ',TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc,  TSPL_Fiscal_Year_Master.Fiscal_Name
                ',TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,TSPL_Vendor_MASTER.Bank_Code as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No, Cast((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)) as decimal(18)) as Payable_Amount,isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0) as payamt from TSPL_PAYMENT_PROCESS_DETAIL 
                'left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                'left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                'left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
                'left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                'left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                'left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                ' left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' 
                'left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code  
                'where  TSPL_PAYMENT_PROCESS_HEAD.From_Date>=convert(date,('" + CycleFromDate + "'),103) and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<=convert(date,('" + CycleToDate + "'),103)    and TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount>0 
                ' )xxx group by Bank_Code,Branch_Name 

                ' )xxxx order by GRPColumn "
                dtAdvice = clsDBFuncationality.GetDataTable(ssql)
            End If
            Dim dtSavingat2 As DataTable = Nothing
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                Dim sqlqry = " select TSPL_DEDUCTION_MASTER.Description as Description, 
                                sum(TSPL_PAYMENT_PROCESS_DEDUCTION.Amount-TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt) as Amount  from TSPL_PAYMENT_PROCESS_DEDUCTION
                                left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No 
                                left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_VENDOR_INVOICE_DETAIL.Document_No 
                                left outer join ( select code, Description  from TSPL_DCS_ADDITION_DEDUCTION
                                union 
                                select  Code , Description from TSPL_DEDUCTION_MASTER) as TSPL_DEDUCTION_MASTER on ( TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode or TSPL_DEDUCTION_MASTER.code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction)
                                where TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in ( " + strDocNo + " )  "
                If CycleWiseData = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
                    sqlqry += " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code In (" + strVSPCode + ")"
                End If

                sqlqry += " and Len(TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE ) > 0 
                                and TSPL_VENDOR_INVOICE_HEAD.Saving In ('1','2') group by TSPL_DEDUCTION_MASTER.Description   "
                dtSavingat2 = clsDBFuncationality.GetDataTable(sqlqry)
            End If


            If dt IsNot Nothing And dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                ' If Flag Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreCHU", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType, "SubPPBA.rpt", dtAdvice)
                    'frmCRV.funsubreportWithdt(Nothing, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreCHU", "", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType, "SubPaymentProcessBankAdvice.rpt", dtAdvice)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreGNG", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreJPR", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreJDH", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreTNK", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreRJS", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreUDP", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreBKN", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreAreaWise", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreJAL", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreNEW", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCreALW", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType, "rptSavingat2.rpt", dtSavingat2)
                Else
                    frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, False, CrystalReportFolder.MilkProcurement, dt, dtDebit, "rptPaymentProcessDebCre", "rptPaymentProcessDebCre.rpt", clsCommon.myCDate(CycleFromDate), "SubPaymentProcessDebit.rpt", "SubPaymentProcessCredit.rpt", dtCredit, "rptPRMilkType.rpt", dtMilkType)
                End If

                frmCRV = Nothing
            Else
                clsCommon.MyMessageBoxShow("No Data Found")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


    'Public Shared Function Load_Report_Paymnet_RCDF_BaseQuery1(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String) As String
    ' Public Shared Function Load_Report_Paymnet_RCDF_BaseQuery1(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String) As String
    Public Shared Function Load_Report_Paymnet_RCDF_BaseQuery1(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal strVSPCode As String, ByVal cyclewisedata As Boolean) As String
        Dim companyADD, CompName, CompCode As String
        Dim settMaxFATPerLimit As Decimal = 0
        Dim settMaxSNFPerLimit As Decimal = 0
        Dim AreaWiseBilling As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = 1)
        settMaxFATPerLimit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxFATPerLimitforReport, clsFixedParameterCode.MaxFATPerLimitforReport, Nothing))
        settMaxSNFPerLimit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxSNFPerLimitforReport, clsFixedParameterCode.MaxSNFPerLimitforReport, Nothing))

        Dim ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster As Boolean = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, clsFixedParameterCode.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, Nothing)) > 0, True, False)
        Dim ApplyMilkTypeBuffaloCowOnPrint As Boolean = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyMilkTypeBuffaloCowOnPrint, clsFixedParameterCode.ApplyMilkTypeBuffaloCowOnPrint, Nothing)) > 0, True, False)

        Dim sQuery As String = ""
        sQuery += " select   TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_COMPANY_MASTER.City_Code)>0 then ', '+TSPL_COMPANY_MASTER.City_Code else ' ' end + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end  as comp_address from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(sQuery)
        companyADD = dt1.Rows(0).Item("comp_address")

        sQuery = ""
        sQuery += " select   TSPL_COMPANY_MASTER.Comp_Name  from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(sQuery)
        CompName = dt2.Rows(0).Item("Comp_Name")


        sQuery = ""
        sQuery += " select   TSPL_COMPANY_MASTER.comp_code  from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
        Dim dt5 As DataTable = clsDBFuncationality.GetDataTable(sQuery)
        CompCode = dt5.Rows(0).Item("Comp_Code")

        Dim fromDate As String = CycleFromDate ' dtpFromDate.Value
        Dim Todate As String = CycleToDate  ' dtpToDate.Value


        Dim whrcls As String = " where 2=2 "
        Dim whrcls1 As String = " where 2=2 "
        Dim whrclsItemWise As String = " where 2=2 "

        whrcls += "  and convert(date,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) <=convert(date,('" + Todate + "'),103) and TBL_BILL_DETAILS.Doc_No IN (" + strDocNo + ") "
        'If clsCommon.myLen(strVSPCode) > 0 Then
        '    whrcls += "  and TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE  in ( " + strVSPCode + ")"
        'End If
        If clsCommon.myLen(strRoutecode) > 0 Then
            whrcls += "  and TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE  in ( " + strRoutecode + ")"
        End If

        'Comment by balwinder on 29/03/2023 as strLoc is segment not mcc code
        'If clsCommon.myLen(strLoc) > 0 Then ' TSPL_LOCATION_MASTER.Loc_Segment_Code
        '    whrcls += " and TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE IN ('" + strLoc + "') "
        'End If
        If clsCommon.myLen(strBank) > 0 Then
            whrcls += " and TBL_BILL_DETAILS.Bank_Code in (" + strBank + ")  "
        End If
        If clsCommon.myLen(strHoldUnhold) > 0 Then
            whrcls += " and TBL_BILL_DETAILS.is_Hold_Payment_Process  = " + strHoldUnhold + "  "
        End If
        If cyclewisedata = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
            whrcls += " and TBL_BILL_DETAILS.VSP_CODE In (" + strVSPCode + ") "
        End If
        whrcls += " and not exists(select 1 from TSPL_MILK_PURCHASE_INVOICE_PRO_LOSS where TSPL_MILK_PURCHASE_INVOICE_PRO_LOSS.InvoiceNo=TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE) "

        ''Comment BSP on 23/07/2024
        'whrcls1 += "  and convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) <=convert(date,('" + Todate + "'),103) "
        'If clsCommon.myLen(strVSPCode) > 0 Then
        '    whrcls1 += "  and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE  in ( " + strVSPCode + ")" '  & clsCommon.GetMulcallString(txtVSP.arrValueMember) &
        'End If

        whrcls1 += "  and TSPL_PAYMENT_PROCESS_Head.doc_no in ( " + strDocNo + " ) "


        If clsCommon.myLen(strLoc) > 0 Then
            whrcls1 += " and TSPL_PAYMENT_PROCESS_Head.loc_seg_code    IN (" + strLoc + ") " 'fndLoc.Value
        End If
        whrclsItemWise += " and final.doc_no in ( " + strDocNo + " )"
        whrclsItemWise += "  and convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) <=convert(date,('" + Todate + "'),103) "
        Dim strPC_FATValue As String = 0
        Dim strPC_SNFValue As String = 0
        Dim dtPC_FAT_SNF As DataTable = clsDBFuncationality.GetDataTable("select  top 1 round ( Price_Chart_FAT_Ratio * Price_Chart_Rate / nullif (Price_Chart_FAT_Per,0) , 2,1 ) as FAT_PCValue , round (  Price_Chart_SNF_Ratio * Price_Chart_Rate / nullif (Price_Chart_SNF_Per,0),2,1)  as SNF_PCValue  from TSPL_PRICE_CHART_PLANNING order by convert (date, Planning_Date,103)")
        If dtPC_FAT_SNF IsNot Nothing AndAlso dtPC_FAT_SNF.Rows.Count > 0 Then
            strPC_FATValue = clsCommon.myCDecimal(dtPC_FAT_SNF.Rows(0)("FAT_PCValue"))
            strPC_SNFValue = clsCommon.myCDecimal(dtPC_FAT_SNF.Rows(0)("SNF_PCValue"))
        End If
        Dim CycleNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select max(Name) as CycleNo from TSPL_PAYMENT_CYCLE_GENERATED where convert(varchar, From_Date,103) = convert(varchar, '" + fromDate + "',103) "))
        Dim BaseQry As String = ""
        BaseQry = "select  TBL_BILL_DETAILS.Doc_No as PPDoc_No,'" + clsCommon.myCstr(clsCommon.GetPrintDate(fromDate, "dd/MM/yyyy")) + "' as PPDoc_Date,'" + CycleNo + "' as CycleNo, "
        If AreaWiseBilling = True Then
            BaseQry += " xxxSetLocation.Location_Desc, "

        End If
        BaseQry += " TBL_BILL_DETAILS.BillNo, TBL_BILL_DETAILS.BillDate, '" + strPC_FATValue + "' as PC_FATValue, '" + strPC_SNFValue + "' as PC_SNFValue, "



        BaseQry += " isnull(TSPL_VENDOR_MASTER.Actual_charges,0) as Actual_charges,isnull (TSPL_VENDOR_MASTER.Rate_Head_Load,0) as Rate_Head_Load ,isnull(TSPL_MILK_PURCHASE_INVOICE_HEAD.Handling_Charges_Amount,0) as Handling_Charges_Amount , TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE as MPD ,convert(varchar,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) as MPI_Date,   TSPL_MCC_MASTER.add1 +case when len(TSPL_MCC_MASTER.add2)>0 then ', '+TSPL_MCC_MASTER.add2 else '' end + case when LEN(TSPL_COMPANY_MASTER.City_Code)>0 then ', '+MCC_City.City_Name  else ' ' end + case when len(TSPL_MCC_MASTER.State_Code )>0 then MCC_State.STATE_NAME else '' end  as MCC_address, "
        BaseQry += "    '" & fromDate & "'  as fromDate ,'" & Todate & "'  as Todate"
        BaseQry += " ,'" & companyADD & "'  as companyADD, '" & CompName & "'  as CompName,'" & CompCode & "'  as CompCode,TSPL_COMPANY_MASTER .Logo_Img   as compLogo1 ,TSPL_COMPANY_MASTER .Logo_Img2 as compLogo2,TSPL_COMPANY_MASTER.Vat_Reg_No,TSPL_COMPANY_MASTER.Phone1 as Comp_Phone1,TSPL_COMPANY_MASTER.Phone2 as Comp_Phone2,TSPL_COMPANY_MASTER.Regn_No,coalesce(PaymentProcess.Total_EMP_Amount,0) as Total_EMP_Amount,coalesce(PaymentProcess.Incentive_Amount,0) as Incentive_Amount ,coalesce(PaymentProcess.Incentive_EMP_Amount,0) as Incentive_EMP_Amount ,coalesce(PaymentProcess.EMP_Amount,0) as EMP_Amount ,coalesce(PaymentProcess.Vsp_Own_System_Amount,0) as Vsp_Own_System_Amount ,coalesce(PaymentProcess.Head_Load_Amount,0) as Head_Load_Amount ,coalesce(PaymentProcess.Payable_Amount,0) as Payable_Amount,coalesce(PaymentProcess.Credit_Note_Amount,0)as Credit_Note_Amount,coalesce(PaymentProcess.Deduction_Amount,0)*(-1) as Deduction_Amount,coalesce(PaymentProcess.Item_Issue_Amount,0)*(-1) as Item_Issue_Amount,coalesce(PaymentProcess.Item_Issue_Return_Amount,0) as Item_Issue_Return_Amount,coalesce(PaymentProcess.MCC_Sale_Amount,0)*(-1) as MCC_Sale_Amount ,coalesce(PaymentProcess.MCC_Sale_Return_Amount,0) as MCC_Sale_Return_Amount, TSPL_MCC_MASTER.add1 + TSPL_MCC_MASTER.add2 as addd,TSPL_MILK_SRN_DETAIL.UOM_Code,TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty "
        BaseQry += " ,case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else  (Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0)) end as Standard_Rate" + Environment.NewLine +
        ",case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else Cast( (((Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0))*Price_Chart.Fat_ratio)/Price_Chart.FAT_Pers) as decimal(18,2)) end as Standard_FAT_Rate" + Environment.NewLine +
        ",case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else  Cast( (((Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0))*Price_Chart.SNF_Ratio)/Price_Chart.SNF_Pers) as decimal(18,2)) end as Standard_SNF_Rate" + Environment.NewLine +
        ",TSPL_MILK_PURCHASE_INVOICE_DETAIL.AMOUNT as Net_AMOUNT,TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_RO_Amount , TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE , convert(varchar,TSPL_MILK_SRN_head.DOC_DATE,103) as DOC_DATE,TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE ,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,'')='' then TSPL_MILK_SRN_HEAD.shift else TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.shift end as SHIFT,"
        BaseQry += " TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE ,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.Bank_Code as Vendor_Bank_Code, TSPL_VENDOR_MASTER.Bank_Name as Vendor_Bank_Name, TSPL_VENDOR_MASTER.Account_Type as Vendor_Bank_Account_Type1 , TSPL_VENDOR_MASTER.AccountType2 as Vendor_Bank_Account_Type2 , TSPL_VENDOR_MASTER.Account_No as Vendor_Account_No1, TSPL_VENDOR_MASTER.AccNo2 as Vendor_Account_No2,TSPL_MCC_ROUTE_MASTER .Route_Name  ,TSPL_MCC_MASTER .MCC_NAME ,case when isnull(TSPL_MILK_SRN_HEAD.Dock_Collection_Milk_Type,'')='' then 'Mix' else TSPL_MILK_SRN_HEAD.Dock_Collection_Milk_Type end as Type ,TSPL_MILK_SRN_DETAIL.CLR,TSPL_MILK_SRN_HEAD.SAMPLE_NO ,TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,"
        BaseQry += " TSPL_VLC_MASTER_HEAD.VLC_Name ,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.TOTAL_PaymentCOMMISSION,0) as [EMP],coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.incentive_head,0) as Incentive,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.total_head_load_amount,0) as HEDAmt,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.total_Own_Asset_Amount,0) as AstAMT,coalesce(Total_dEDUCTION_AMOUNT,0) as DedAmt"
        BaseQry += " ,TSPL_VLC_MASTER_HEAD.Village_Code, TSPL_VILLAGE_MASTER.Village_Name,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
            BaseQry += " 'Mix' as CowBuffalo_Type "
        Else
            BaseQry += "" + IIf(ApplyMilkTypeBuffaloCowOnPrint = True, " case when TSPL_PRICE_CHART_PLANNING.Dock_Collection_Milk_Type = 'M' then 'Buffalo' else 'Cow' end  ", "'Mixed'") + "  as CowBuffalo_Type "
        End If
        BaseQry += ",(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)) as SRN_Net_Amount
,TSPL_MILK_PURCHASE_INVOICE_HEAD.Total_Basic_AMOUNT "
        If ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster = True Then
            BaseQry += " ,coalesce(TSPL_Primary_Vehicle_Master.Vehicle,TSPL_MILK_SRN_HEAD.VEHICLE_CODE) as VEHICLE_CODE "
        Else
            BaseQry += " ,TSPL_MILK_SRN_HEAD.VEHICLE_CODE "
        End If
        BaseQry += ",cast( case when  TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty=0 then 0 else (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))/TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty end as decimal(18,2)) as RATE
,TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER 
,TSPL_MILK_SRN_DETAIL.FAT_KG as FATQTY
,cast(case when round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER*TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 )=0 then 0 else ( cast( round((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer)/round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER*TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 ) ) end as decimal(18,2)) as FAT_Rate
,cast( round((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer) as FAT_Amount
,TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER
,TSPL_MILK_SRN_DETAIL.SNF_KG as SNFQTY
,cast(case when round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER *TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 )=0 then 0 else (cast(((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)))-round( (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer)/round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER *TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1)) end as decimal(18,2)) as SNF_Rate
,cast(((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)))-round( (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer) as SNF_Amount "
        ', " + IIf(objCommonVar.CurrentCompanyCode = "RCDF", "TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type",
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            'BaseQry += "  ,CASE WHEN (CASE WHEN TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No IS NOT NULL THEN ISNULL(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type, 'SWEET') 
            '                ELSE (CASE WHEN TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No IS NOT NULL THEN ISNULL(TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type, 'SWEET') END) END) = 'forfeited' THEN 'SWEET' 
            '                    ELSE (CASE WHEN TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No IS NOT NULL THEN ISNULL(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type, 'SWEET') 
            '                 ELSE (CASE WHEN TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No IS NOT NULL THEN ISNULL(TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type, 'SWEET') END) END) 
            '                END AS QBD"
            BaseQry += " ,CASE WHEN TSPL_MILK_SRN_DETAIL.FAT_PER < " & settMaxFATPerLimit & " OR TSPL_MILK_SRN_DETAIL.SNF_PER < " & settMaxSNFPerLimit & " THEN 'GOODFAIL'
                            WHEN (CASE WHEN TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No IS NOT NULL 
                           THEN ISNULL(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type, 'SWEET') WHEN TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No IS NOT NULL 
                           THEN ISNULL(TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type, 'SWEET')END) = 'forfeited'
                           THEN 'SWEET'
                            ELSE 
                           (CASE WHEN TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No IS NOT NULL 
                           THEN ISNULL(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type, 'SWEET') WHEN TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No IS NOT NULL 
                            THEN ISNULL(TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type, 'SWEET')END)END AS QBD "
        Else
            BaseQry += "  ,(case when TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No is not null then isnull (TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,'SWEET') else (case when TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No is not null then isnull (TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type,'SWEET') end) end) as QBD "
        End If



        BaseQry += " , '1' as QAP,TBL_BILL_DETAILS.BILLSRL  
,TabSaving.Item_Desc as SavingDesc, TabSaving.[Amount] as [SavingAmount]
,(case when row_number() over(partition by TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE order by TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE)=1
then PaymentProcess.PPSRN_RO_Amount else 0 end) as PPSRN_RO_Amount
         from TSPL_MILK_PURCHASE_INVOICE_DETAIL  " + Environment.NewLine + " Inner Join TSPL_MILK_PURCHASE_INVOICE_HEAD On TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE =TSPL_MILK_PURCHASE_INVOICE_DETAIL.DOC_CODE  " + Environment.NewLine + " left outer join TSPL_MILK_SRN_HEAD  on TSPL_MILK_SRN_HEAD .DOC_CODE  =TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE " + Environment.NewLine
        BaseQry += " left outer join TSPL_MILK_SRN_DETAIL   on TSPL_MILK_SRN_DETAIL .DOC_CODE  =TSPL_MILK_SRN_HEAD.DOC_CODE " + Environment.NewLine
        BaseQry += " " + Environment.NewLine + " Left Outer Join TSPL_VENDOR_MASTER On"
        BaseQry += " TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE =TSPL_VENDOR_MASTER.Vendor_Code And TSPL_VENDOR_MASTER.Form_Type = 'VSP'  " + Environment.NewLine + " Left Outer Join TSPL_MCC_MASTER On TSPL_MILK_PURCHASE_INVOICE_HEAD .MCC_CODE = TSPL_MCC_MASTER.MCC_Code  " + Environment.NewLine
        '       If AreaWiseBilling = True Then
        '           BaseQry += " Left Outer Join( select TSPL_PAYMENT_PROCESS_HEAD.Doc_No,tspl_location_master.Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD  
        'left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
        ') xxxSetLocation On xxxSetLocation.Location_Code=TSPL_MCC_MASTER.area_Location_code"
        '       Else
        '           BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code"
        '       End If

        ' " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code 

        BaseQry += " Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE =TSPL_MCC_ROUTE_MASTER.Route_Code " + Environment.NewLine + " left outer join TSPL_VLC_MASTER_HEAD on"
        BaseQry += " TSPL_VLC_MASTER_HEAD.VLC_Code =TSPL_MILK_PURCHASE_INVOICE_DETAIL.VLC_NO  " + Environment.NewLine
        BaseQry += " left join TSPL_CITY_MASTER  as MCC_City on MCC_City.city_code=TSPL_MCC_MASTER.City_code " + Environment.NewLine
        BaseQry += " left join TSPL_STATE_MASTER as MCC_State on MCC_State.STATE_CODE =TSPL_MCC_MASTER.State_Code " + Environment.NewLine
        BaseQry += " left join TSPL_Primary_Vehicle_Master on TSPL_Primary_Vehicle_Master.VEHICLE_CODE = TSPL_MILK_SRN_HEAD.VEHICLE_CODE " + Environment.NewLine
        BaseQry += " left join  (select VLC_Code, VSP_CODE,sum(Total_EMP_Amount) as Total_EMP_Amount,sum(Incentive_Amount) as Incentive_Amount,sum(Incentive_EMP_Amount) as Incentive_EMP_Amount,sum(EMP_Amount) as EMP_Amount,sum(Vsp_Own_System_Amount) as Vsp_Own_System_Amount,sum(Head_Load_Amount) as Head_Load_Amount,sum(Payable_Amount) as Payable_Amount,sum(Credit_Note_Amount)as Credit_Note_Amount,sum(Deduction_Amount) as Deduction_Amount,sum(Item_Issue_Amount) as Item_Issue_Amount,sum(Item_Issue_Return_Amount) as Item_Issue_Return_Amount,sum(MCC_Sale_Amount) as MCC_Sale_Amount ,sum(MCC_Sale_Return_Amount) as MCC_Sale_Return_Amount,SUM(SRN_RO_Amount) as PPSRN_RO_Amount,MAX(Area_Location_Code)Area_Location_Code from (select TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code, TSPL_PAYMENT_PROCESS_DETAIL.SRN_RO_Amount,TSPL_PAYMENT_PROCESS_DETAIL.Incentive_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.Incentive_EMP_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.EMP_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.Vsp_Own_System_Amount,TSPL_PAYMENT_PROCESS_DETAIL.Total_EMP_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount , TSPL_VLC_MASTER_HEAD.VLC_Code  ,TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE ,TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,TSPL_PAYMENT_PROCESS_DETAIL.Credit_Note_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.Deduction_Amount  ,TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Return_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Amount,TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Amount ,TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Return_Amount  from TSPL_PAYMENT_PROCESS_DETAIL"
        BaseQry += " left join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No =TSPL_PAYMENT_PROCESS_DETAIL.Doc_No"
        BaseQry += " left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader =TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader " & whrcls1 & ""
        If cyclewisedata = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
        End If

        BaseQry += " ) as pp group by VSP_CODE,VLC_Code"
        BaseQry += " ) as PaymentProcess on "
        BaseQry += "  PaymentProcess.vsp_code = TSPL_MILK_PURCHASE_INVOICE_Head.vsp_code And PaymentProcess.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_Code " + Environment.NewLine
        BaseQry += " left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_MILK_PURCHASE_INVOICE_Head.Comp_Code left join TSPL_PRICE_CHART_PLANNING on TSPL_PRICE_CHART_PLANNING.Planning_Code=TSPL_MILK_SRN_DETAIL.Price_Code " + Environment.NewLine
        BaseQry += " left join (select distinct FAT_Pers,SNF_Pers,Ratio as Fat_ratio,SNF_Ratio, Milk_Rate,TSPL_MILK_PRICE_MASTER.Price_Code,TSPL_FAT_SNF_UPLOADER_MASTER.code    from TSPL_FAT_SNF_UPLOADER_MASTER inner join  TSPL_MILK_PRICE_MASTER  on TSPL_MILK_PRICE_MASTER.Price_Code=TSPL_FAT_SNF_UPLOADER_MASTER.Price_Code) as  Price_Chart    on TSPL_MILK_SRN_DETAIL.Price_Code=Price_Chart.Code "
        BaseQry += " left outer join TSPL_VILLAGE_MASTER on TSPL_VILLAGE_MASTER.Village_Code = TSPL_VLC_MASTER_HEAD.Village_Code " + Environment.NewLine +
        " left join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Uploader_TR_No
          LEFT OUTER JOIN TSPL_MILK_SHIFT_UPLOADER_DETAIL ON TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No "

        If AreaWiseBilling = True Then
            BaseQry += " Left Outer Join(select MAX(TSPL_PAYMENT_PROCESS_HEAD.Doc_No)Doc_No,MAX(tspl_location_master.Location_Desc)Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD  
                         left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
						 GROUP BY tspl_location_master.Location_Code
                         ) xxxSetLocation On xxxSetLocation.Location_Code=PaymentProcess.area_Location_code "
        Else
            BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code "
        End If
        BaseQry += " left outer join (select TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE , TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No as BillNo , convert(varchar,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_Date,103) as BillDate, TSPL_PAYMENT_PROCESS_DETAIL.SNo as BILLSRL, TSPL_PAYMENT_PROCESS_DETAIL.Doc_No, TSPL_PAYMENT_PROCESS_DETAIL. is_Hold_Payment_Process, TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in ( " + strDocNo + " ) "
        If cyclewisedata = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += " and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
        End If

        BaseQry += ") as TBL_BILL_DETAILS on TBL_BILL_DETAILS.VSP_CODE =
          TSPL_MILK_PURCHASE_INVOICE_Head.vsp_code and TBL_BILL_DETAILS.BillNo=TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE
          left outer join (select VSP_Code,max(Item_Desc) as Item_Desc, sum([Amount]) as [Amount] from (
			 select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Item_Desc,(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ") "
        If cyclewisedata = True AndAlso clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_CODE In (" + strVSPCode + ")"
        End If
        BaseQry += ")x group by VSP_Code)TabSaving on TabSaving.VSP_Code=TBL_BILL_DETAILS.VSP_CODE"
        BaseQry += "  " & whrcls & " "
        'sQuery = " select CowBuffalo_Type,sum( Qty) as Qty, round( sum (FATQTY) * 100 /sum( Qty),2,1)  as FATPer, round( sum (SNFQTY) * 100 / sum(Qty),2,1) as SNFPer,sum(SRN_NET_AMOUNT)+(sum(PPSRN_RO_Amount)*-1) AS Amt,round(sum (FATQTY),2,1)  as FATKG,round(sum (SNFQTY),2,1) as SNFKG from ( " + BaseQry + " ) XXXX group by CowBuffalo_Type "

        Return BaseQry
    End Function
    Public Shared Function Load_Report_Paymnet_RCDF(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal QryForShowData As Boolean, ByVal strZone As String) As String
        Return Load_Report_Paymnet_RCDF(strDocNo, CycleFromDate, CycleToDate, strLoc, strVSPCode, strRoutecode, strBank, strHoldUnhold, QryForShowData, False, strZone, False)
    End Function

    Public Shared Function Load_Report_Paymnet_RCDF(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal QryForShowData As Boolean, ByVal strZone As String, ByVal isprintHindi As Boolean) As String
        Return Load_Report_Paymnet_RCDF(strDocNo, CycleFromDate, CycleToDate, strLoc, strVSPCode, strRoutecode, strBank, strHoldUnhold, QryForShowData, False, "", isprintHindi)
    End Function
    Public Shared Function Load_Report_Paymnet_RCDF(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal QryForShowData As Boolean) As String
        Return Load_Report_Paymnet_RCDF(strDocNo, CycleFromDate, CycleToDate, strLoc, strVSPCode, strRoutecode, strBank, strHoldUnhold, QryForShowData, False, Nothing)
    End Function
    Public Shared Function Load_Report_Paymnet_RCDF(ByVal strDocNo As String, ByVal CycleFromDate As String, ByVal CycleToDate As String, ByVal strLoc As String, ByVal strVSPCode As String, ByVal strRoutecode As String, ByVal strBank As String, ByVal strHoldUnhold As String, ByVal QryForShowData As Boolean, ByVal isPDFPath As Boolean, ByVal strZone As String, ByVal isprintHindi As Boolean) As String
        Dim PDFPath As String = ""
        Dim companyADD As String
        Dim User_Name As String = objCommonVar.CurrentUser
        Dim PickHeadLoadRateFromSecretaryMaster As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.PickHeadLoadRateFromSecretaryMaster, clsFixedParameterCode.PickHeadLoadRateFromSecretaryMaster, Nothing)) = 1)
        Dim SetCowFatPer As Decimal = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.CowFATPer, clsFixedParameterCode.CowFATPer, Nothing))
        Dim IncentiveRate As Decimal = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryIncentiveRate, clsFixedParameterCode.MPIncentiveEntryIncentiveRate, Nothing))
        Dim AreaWiseBilling As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = 1)
        Dim PaymentProcessInHindi As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.PaymentProcessPrintInHindi, clsFixedParameterCode.PaymentProcessPrintInHindi, Nothing)) = 1)
        Dim PickDataFromMasterOrTransaction As Boolean = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.PickDataFromMasterOrTransaction, clsFixedParameterType.PickDataFromMasterOrTransaction, Nothing)) = 0, True, False)
        Dim ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster As Boolean = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, clsFixedParameterCode.ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster, Nothing)) > 0, True, False)
        Dim sQuery As String = ""
        sQuery += " select   TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_COMPANY_MASTER.City_Code)>0 then ', '+TSPL_COMPANY_MASTER.City_Code else ' ' end + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end  as comp_address from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(sQuery)
        companyADD = dt1.Rows(0).Item("comp_address")
        Dim fromDate As String = CycleFromDate
        Dim Todate As String = CycleToDate
        Dim whrcls As String = " where 2=2 "
        Dim whrcls1 As String = " where 2=2 "
        Dim whrclsItemWise As String = " where 2=2 "
        whrcls += "  and convert(date,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) <=convert(date,('" + Todate + "'),103) and TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty > 0 "
        If clsCommon.myLen(strVSPCode) > 0 Then
            whrcls += "  and TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE  in ( " + strVSPCode + ")"
        End If
        If clsCommon.myLen(strRoutecode) > 0 Then
            whrcls += "  and TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE  in ( " + strRoutecode + ")"
        End If
        If clsCommon.myLen(strLoc) > 0 Then
            whrcls += " and TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE IN (" + strLoc + ") "
        End If
        If clsCommon.myLen(strBank) > 0 Then
            whrcls += " and TBL_BILL_DETAILS.Bank_Code in (" + strBank + ")  "
        End If
        If clsCommon.myLen(strHoldUnhold) > 0 Then
            whrcls += " and TBL_BILL_DETAILS.is_Hold_Payment_Process  = " + strHoldUnhold + "  "
        End If
        If clsCommon.myLen(strZone) > 0 Then
            whrcls += "   and TSPL_VENDOR_MASTER.Zone_Code  in ( " + strZone + ")"
        End If
        ''Comment By Balwinder on 23/07/2024
        'whrcls += " and not exists(select 1 from TSPL_MILK_PURCHASE_INVOICE_PRO_LOSS where TSPL_MILK_PURCHASE_INVOICE_PRO_LOSS.InvoiceNo=TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE) "

        whrcls1 += "  and convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) <=convert(date,('" + Todate + "'),103) "
        If clsCommon.myLen(strVSPCode) > 0 Then
            whrcls1 += "  and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE  in ( " + strVSPCode + ")" '  & clsCommon.GetMulcallString(txtVSP.arrValueMember) &
        End If
        whrcls1 += "  and TSPL_PAYMENT_PROCESS_Head.doc_no in ( " + strDocNo + " ) "
        'If clsCommon.myLen(strLoc) > 0 Then
        '    whrcls1 += " and TSPL_PAYMENT_PROCESS_Head.loc_seg_code    IN (" + strLoc + ") " 'fndLoc.Value
        'End If
        whrclsItemWise += " and final.doc_no in ( " + strDocNo + " )"
        whrclsItemWise += "  and convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103)>=convert(date,('" + fromDate + "'),103) and convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) <=convert(date,('" + Todate + "'),103) "
        Dim CycleNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select max(Name) as CycleNo from TSPL_PAYMENT_CYCLE_GENERATED where convert(varchar, From_Date,103) = convert(varchar, '" + fromDate + "',103) "))
        Dim BaseQry As String = ""
        BaseQry = "select '" & User_Name & "' as User_Name, "
        If PickDataFromMasterOrTransaction Then
            BaseQry += " TSPL_MILK_SRN_DETAIL.Head_Load_Rate as Rate_Head_Load,PaymentProcess.Payee_Joint_Account_No as Vendor_Account_No1,PaymentProcess.Payee_Joint_Branch_Name as Vendor_Bank_Code,PaymentProcess.VSP_NAME AS Vendor_Name, "
        Else
            BaseQry += " TSPL_MILK_SRN_DETAIL.Head_Load_Rate as Rate_Head_Load,TSPL_VENDOR_MASTER.Account_No as Vendor_Account_No1, TSPL_VENDOR_MASTER.Bank_Code as Vendor_Bank_Code,TSPL_VENDOR_MASTER.Vendor_Name, "
        End If

        BaseQry += " CycleWise.Head_Load_Cycle,CycleWise.Rate as CycleWiseRate,TSPL_COMPANY_MASTER.GSTReg_No, TBL_BILL_DETAILS.Doc_No as PPDoc_No,'" + clsCommon.myCstr(clsCommon.GetPrintDate(fromDate, "dd/MM/yyyy")) + "' as PPDoc_Date,'" + CycleNo + "' as CycleNo, TBL_BILL_DETAILS.BillNo, TBL_BILL_DETAILS.BillDate, round ( TSPL_PRICE_CHART_PLANNING.Price_Chart_FAT_Ratio * TSPL_PRICE_CHART_PLANNING.Price_Chart_Rate / nullif (TSPL_PRICE_CHART_PLANNING.Price_Chart_FAT_Per,0) , 2,1 ) as PC_FATValue , round (  TSPL_PRICE_CHART_PLANNING.Price_Chart_SNF_Ratio * TSPL_PRICE_CHART_PLANNING.Price_Chart_Rate / nullif (TSPL_PRICE_CHART_PLANNING.Price_Chart_SNF_Per,0),2,1)  as PC_SNFValue,  isnull(TSPL_VENDOR_MASTER.Actual_charges,0) as Actual_charges, "

        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
        'If PickHeadLoadRateFromSecretaryMaster Then
        '    'BaseQry += " TSPL_VENDOR_MASTER.Rate_Head_Load as Rate_Head_Load, 
        '    BaseQry += " TSPL_MILK_SRN_DETAIL.Head_Load_Rate as Rate_Head_Load,"
        'Else
        '    'BaseQry += " ISnull(Headload.Head_Load_Rate, 0)Rate_Head_Load, "
        '    BaseQry += " TSPL_MILK_SRN_DETAIL.Head_Load_Rate as Rate_Head_Load, "
        'End If
        'BaseQry += " ISnull(Headload.Head_Load_Rate, 0)Rate_Head_Load, "
        'End If

        BaseQry += " isnull(TSPL_MILK_PURCHASE_INVOICE_HEAD.Handling_Charges_Amount,0) as Handling_Charges_Amount , TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE as MPD ,convert(varchar,TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) as MPI_Date,   TSPL_MCC_MASTER.add1 +case when len(TSPL_MCC_MASTER.add2)>0 then ', '+TSPL_MCC_MASTER.add2 else '' end + case when LEN(TSPL_COMPANY_MASTER.City_Code)>0 then ', '+MCC_City.City_Name  else ' ' end + case when len(TSPL_MCC_MASTER.State_Code )>0 then MCC_State.STATE_NAME else '' end  as MCC_address, 
'" & fromDate & "'  as fromDate ,'" & Todate & "'  as Todate
,'" & companyADD & "'  as companyADD, '" & objCommonVar.CurrentCompanyName & "'  as CompName,TSPL_COMPANY_MASTER.Comp_Code1,TSPL_COMPANY_MASTER.Comp_Name  as CompName1,'" & objCommonVar.CurrentCompanyCode & "'  as CompCode,TSPL_COMPANY_MASTER .Logo_Img   as compLogo1 ,TSPL_COMPANY_MASTER .Logo_Img2 as compLogo2,TSPL_COMPANY_MASTER.Pan_No As Comp_PanNo,TSPL_COMPANY_MASTER.Vat_Reg_No,replace(replace(TSPL_COMPANY_MASTER.Phone1 ,'(+__)__________',''),'(+91)','') as Comp_Phone1,replace(replace(TSPL_COMPANY_MASTER.Phone2 ,'(+__)__________',''),'(+91)','') as Comp_Phone2,TSPL_COMPANY_MASTER.Regn_No,coalesce(PaymentProcess.Total_EMP_Amount,0) as Total_EMP_Amount,coalesce(PaymentProcess.Incentive_Amount,0) as Incentive_Amount ,coalesce(PaymentProcess.Incentive_EMP_Amount,0) as Incentive_EMP_Amount ,coalesce(PaymentProcess.EMP_Amount,0) as EMP_Amount ,coalesce(PaymentProcess.Vsp_Own_System_Amount,0) as Vsp_Own_System_Amount ,coalesce(PaymentProcess.Head_Load_Amount,0) as Head_Load_Amount ,coalesce(PaymentProcess.Payable_Amount,0) as Payable_Amount,coalesce(PaymentProcess.Credit_Note_Amount,0)as Credit_Note_Amount,coalesce(PaymentProcess.Deduction_Amount,0)*(-1) as Deduction_Amount,coalesce(PaymentProcess.Item_Issue_Amount,0)*(-1) as Item_Issue_Amount,coalesce(PaymentProcess.Item_Issue_Return_Amount,0) as Item_Issue_Return_Amount,coalesce(PaymentProcess.MCC_Sale_Amount,0)*(-1) as MCC_Sale_Amount ,coalesce(PaymentProcess.MCC_Sale_Return_Amount,0) as MCC_Sale_Return_Amount, TSPL_MCC_MASTER.add1 + TSPL_MCC_MASTER.add2 as addd,TSPL_MILK_SRN_DETAIL.UOM_Code,TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty, ROUND((TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/TSPL_ITEM_UOM_DETAIL.Conversion_Factor),3) as 'Milk Weight LTR1'
,case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else  (Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0)) end as Standard_Rate
,case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else Cast( (((Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0))*Price_Chart.Fat_ratio)/Price_Chart.FAT_Pers) as decimal(18,2)) end as Standard_FAT_Rate
,case when TSPL_MILK_SRN_DETAIL.AMOUNT=0 then 0 else  Cast( (((Price_Chart.milk_rate+isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive_Rate,0))*Price_Chart.SNF_Ratio)/Price_Chart.SNF_Pers) as decimal(18,2)) end as Standard_SNF_Rate 
,TSPL_MILK_PURCHASE_INVOICE_DETAIL.AMOUNT as Net_AMOUNT,TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_RO_Amount , TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE ,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader, convert(varchar,TSPL_MILK_SRN_head.DOC_DATE,103) as DOC_DATE,TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE ,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,'')='' then TSPL_MILK_SRN_HEAD.shift else TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.shift end as SHIFT,"
        If AreaWiseBilling Then
            BaseQry += " xxxRoutename.ROUTE_NAME as  ROUTE_NAME,xxxRoutename.ROUTE_NO as route_code "
        Else
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                BaseQry += " isnull(TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO,'') AS ROUTE_CODE,isnull(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,'') as Route_Name "
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                BaseQry += " isnull(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,'') as Route_Name ,isnull(TSPL_BULK_ROUTE_MASTER.ROUTE_NO,'') AS ROUTE_CODE,isnull(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME_HINDI,'') as ROUTE_NAME_HINDI "
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                BaseQry += " TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE ,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME as Route_Name ,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME_HINDI as ROUTE_NAME_HINDI "
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                BaseQry += " TSPL_VLC_MASTER_HEAD.Route_Code as ROUTE_CODE,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME as Route_Name,isnull(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME_HINDI,'') as ROUTE_NAME_HINDI "
            Else
                BaseQry += " TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE ,TSPL_MCC_ROUTE_MASTER.Route_Name  "
            End If
        End If


        'If AreaWiseBilling = True Then
        '    BaseQry += " ,tspl_mcc_master.MCC_NAME where MCC_Code='" & clsCommon.myCstr(strLoc) & "'"
        'Else
        '    BaseQry += " "
        'End If
        BaseQry += ",TSPL_MILK_SRN_HEAD.ROUTE_CODE As SRN_ROUTE,TSPL_VENDOR_MASTER.Vendor_Name_Hindi, TSPL_VENDOR_MASTER.Bank_Name as Vendor_Bank_Name, TSPL_VENDOR_MASTER.Account_Type as Vendor_Bank_Account_Type1 , TSPL_VENDOR_MASTER.AccountType2 as Vendor_Bank_Account_Type2 , "

        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
        'BaseQry += " PaymentProcess.Payee_Joint_Account_No as Vendor_Account_No1,PaymentProcess.Payee_Joint_Branch_Name as Vendor_Bank_Code, "
        ' Else
        'BaseQry += " TSPL_VENDOR_MASTER.Account_No as Vendor_Account_No1, TSPL_VENDOR_MASTER.Bank_Code as Vendor_Bank_Code, "
        ' End If
        BaseQry += " PaymentProcess.Bank_Account_No_Saving as Vendor_Account_No2,TSPL_VENDOR_MASTER.PAN As DCS_PAN, "
        If AreaWiseBilling = True Then
            BaseQry += "  xxxSetLocation.Location_Desc as MCC_NAME "
        Else
            BaseQry += " TSPL_MCC_MASTER .MCC_NAME ,TSPL_MCC_MASTER.MCC_Name_Hindi "
        End If
        BaseQry += " ,case when isnull(TSPL_MILK_SRN_HEAD.Dock_Collection_Milk_Type,'')='' then 'Mix' else TSPL_MILK_SRN_HEAD.Dock_Collection_Milk_Type end as Type ,TSPL_MILK_SRN_DETAIL.CLR,TSPL_MILK_SRN_HEAD.SAMPLE_NO ,TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VENDOR_MASTER.Registered_PDCS_CLUSTER as DCS_Type,
TSPL_VLC_MASTER_HEAD.VLC_Name ,TSPL_VLC_MASTER_HEAD.VLC_Name_Hindi,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.TOTAL_PaymentCOMMISSION,0) as [EMP],coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.incentive_head,0) as Incentive,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.total_head_load_amount,0) as HEDAmt,coalesce(TSPL_MILK_PURCHASE_INVOICE_HEAD.total_Own_Asset_Amount,0) as AstAMT,coalesce(Total_dEDUCTION_AMOUNT,0) as DedAmt
,TSPL_VLC_MASTER_HEAD.Village_Code, TSPL_VILLAGE_MASTER.Village_Name,TSPL_VILLAGE_MASTER.Village_Name_Hindi,(case when TSPL_PRICE_CHART_PLANNING.Dock_Collection_Milk_Type='C' then 'Cow' else 'Buffalo' end) as CowBuffalo_Type "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            BaseQry += " ,case when TSPL_VENDOR_MASTER.Registered_PDCS_CLUSTER = 'PDCS' then TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader + '       CC       ' + TSPL_VENDOR_MASTER.Vendor_Name else TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader +'        '+ TSPL_VENDOR_MASTER.Vendor_Name end as Society,
                         '1'  as QAP"
        Else
            BaseQry += " ,'1' AS QAP "
        End If
        BaseQry += "     ,(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)) as SRN_Net_Amount
,TSPL_MILK_PURCHASE_INVOICE_HEAD.Total_Basic_AMOUNT "
        If ShowVehicleNoSeparatelyInPrimaryTransVehicleMaster = True Then
            BaseQry += " ,coalesce(TSPL_Primary_Vehicle_Master.Vehicle,TSPL_MILK_SRN_HEAD.VEHICLE_CODE) as VEHICLE_CODE "
        Else
            BaseQry += " ,TSPL_MILK_SRN_HEAD.VEHICLE_CODE "
        End If
        BaseQry += ",cast( case when  TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty=0 then 0 else (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))/TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty end as decimal(18,2)) as RATE
,TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER 
,round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER*TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,3,1 ) as FATQTY
,cast(case when round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER*TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 )=0 then 0 else ( cast( round((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer)/round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.FAT_PER*TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 ) ) end as decimal(18,2)) as FAT_Rate
,cast( round((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer) as FAT_Amount
,TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER
,round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER *TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,3,1 ) as SNFQTY 
,cast(case when round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER *TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1 )=0 then 0 else (cast(((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)))-round( (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer)/round(TSPL_MILK_PURCHASE_INVOICE_DETAIL.SNF_PER *TSPL_MILK_PURCHASE_INVOICE_DETAIL.Qty/100,2,1)) end as decimal(18,2)) as SNF_Rate
,cast(((TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0)))-round( (TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_Net_Amount+ isnull(TSPL_MILK_SRN_DETAIL.VSP_Day_Wise_Incentive,0))*isnull(TSPL_MILK_SRN_DETAIL.FAT_Ratio,0),0) as integer) as SNF_Amount
,(case when  TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No is not null then isnull (TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,'SWEET') else (case when  TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No is not null then isnull (TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type,'SWEET') end) end) as QBD
,TBL_BILL_DETAILS.BILLSRL,TabSaving.Item_Desc as SavingDesc,TabCompulsory.Item_Desc as CompulsoryDesc,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            BaseQry += " TabOverheadSaving.Amount as OverheadSaving,TabOverheadCompulsory.Amount as OverheadCompulsory,  "
        End If

        '    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
        '    BaseQry += " Round(ISNULL(TabSaving.[Amount],0), 0) + ISNULL(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0) as [SavingAmount],TSPL_TRANSFER_TO_SAVING_DETAIL.Amount "
        'Else
        BaseQry += " Round(ISNULL(TabSaving.[Amount],0), 0) as [SavingAmount],Round(ISNULL(TabCompulsory.Amount,0),0) as [CompulsoryAmount]"
        'End If
        BaseQry += ",convert(varchar,PaymentProcess.Created_Date,103) as Created_Date"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            BaseQry += ",Tab_TSPL_PRICE_CHART_PLANNING_TSDDCF.Rate_Per as Planing_Rate_Per,Tab_TSPL_PRICE_CHART_PLANNING_TSDDCF.Fixed_Rate as Planing_Fixed_Rate,tspl_vendor_master.Email"
        End If
        BaseQry += ",convert(varchar,TSPL_MILK_PURCHASE_INVOICE_HEAD.Created_Date,103) as MPI_CREATED_Date,TSPL_MILK_SRN_DETAIL.ACC_Qty_LTR as 'Milk Weight LTR'," + clsCommon.myCstr(IncentiveRate) + " as IncentiveRate,TSPL_MILK_PURCHASE_INVOICE_DETAIL.Cans,FORMAT(CONVERT(date,'" + clsCommon.myCstr(clsCommon.GetPrintDate(Todate, "dd/MM/yyyy")) + "' ,103),'ddd') as DaysName,PaymentProcess.Compulsory_Amount,TSPL_VENDOR_MASTER.PAN
                    ,TSPL_VENDOR_MASTER.RegistrationNo,Price_Chart.CircularNo,PaymentProcess.zone_code "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            BaseQry += " ,Round(ISNULL(TrnferAmt.Trsfer_Amt,0), 0) as Trsfer_Amt "
        End If

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal AndAlso PaymentProcessInHindi = True AndAlso isprintHindi Then
            BaseQry += " ,TabMilkPurchase.Amount as DCSEXPAmt "
        End If


        BaseQry += "  from TSPL_MILK_PURCHASE_INVOICE_DETAIL  " + Environment.NewLine + " Inner Join TSPL_MILK_PURCHASE_INVOICE_HEAD On TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE =TSPL_MILK_PURCHASE_INVOICE_DETAIL.DOC_CODE     left outer join (Select InvoiceNo,Sum(Head_Load_Cycle)Head_Load_Cycle,max(Rate)Rate from TSPL_MILK_PURCHASE_INVOICE_DAY_WISE_HEAD_LOAD group by InvoiceNo ) as CycleWise on CycleWise.InvoiceNo=TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE  " + Environment.NewLine + " left outer join TSPL_MILK_SRN_HEAD  on TSPL_MILK_SRN_HEAD .DOC_CODE  =TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE " + Environment.NewLine
        BaseQry += " left outer join TSPL_MILK_SRN_DETAIL   on TSPL_MILK_SRN_DETAIL .DOC_CODE  =TSPL_MILK_SRN_HEAD.DOC_CODE " + Environment.NewLine
        BaseQry += "  " + Environment.NewLine + " Left Outer Join TSPL_VENDOR_MASTER On"
        BaseQry += " TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE =TSPL_VENDOR_MASTER.Vendor_Code And TSPL_VENDOR_MASTER.Form_Type = 'VSP'  " + Environment.NewLine
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") <> CompairStringResult.Equal Then
            BaseQry += " Left Outer Join TSPL_MCC_MASTER On TSPL_MILK_PURCHASE_INVOICE_HEAD .MCC_CODE = TSPL_MCC_MASTER.MCC_Code  " + Environment.NewLine
        End If

        '       If AreaWiseBilling = True Then
        '           BaseQry += " Left Outer Join( select TSPL_PAYMENT_PROCESS_HEAD.Doc_No,tspl_location_master.Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD  
        'left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
        ') xxxSetLocation On xxxSetLocation.Location_Code=TSPL_MCC_MASTER.area_Location_code"
        '       Else
        '           BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code " + Environment.NewLine
        '       End If
        'BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code " + Environment.NewLine
        BaseQry += " left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code =TSPL_MILK_PURCHASE_INVOICE_DETAIL.VLC_NO  " + Environment.NewLine
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
            BaseQry += "  Left Outer Join TSPL_BULK_ROUTE_MASTER On TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE =TSPL_BULK_ROUTE_MASTER.ROUTE_NO" + Environment.NewLine
        ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            BaseQry += "  Left Outer Join TSPL_BULK_ROUTE_MASTER On TSPL_VLC_MASTER_HEAD.Route_Code =TSPL_BULK_ROUTE_MASTER.ROUTE_NO" + Environment.NewLine
        Else
            BaseQry += "  Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE =TSPL_MCC_ROUTE_MASTER.Route_Code " + Environment.NewLine
        End If


        'BaseQry += " left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code =TSPL_MILK_PURCHASE_INVOICE_DETAIL.VLC_NO  " + Environment.NewLine
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
            BaseQry += " Left Outer Join TSPL_MCC_MASTER On TSPL_VLC_MASTER_HEAD.MCC = TSPL_MCC_MASTER.MCC_Code " + Environment.NewLine
        End If

        'If AreaWiseBilling = True Then
        '    BaseQry += " Left Outer Join( select TSPL_PAYMENT_PROCESS_HEAD.Doc_No,tspl_location_master.Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD  
        '                 left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
        '                 ) xxxSetLocation On xxxSetLocation.Location_Code=TSPL_MCC_MASTER.area_Location_code"
        'Else
        '    BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code " + Environment.NewLine
        'End If
        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
        'BaseQry += " left join(SELECT TSPL_HEAD_LOAD.Document_No,TSPL_HEAD_LOAD_DCS.VLC_CODE,MAX(TSPL_HEAD_LOAD_DCS.Head_Load_Rate) Head_Load_Rate FROM (
        'SELECT DISTINCT MAX(TSPL_HEAD_LOAD.Document_No) Document_No,MAX( TSPL_HEAD_LOAD_DCS.VLC_CODE) VLC_CODE FROM TSPL_HEAD_LOAD
        '                       LEFT OUTER JOIN TSPL_HEAD_LOAD_DCS ON TSPL_HEAD_LOAD_DCS.Document_No=TSPL_HEAD_LOAD.Document_No
        '                       GROUP BY TSPL_HEAD_LOAD_DCS.VLC_CODE) TSPL_HEAD_LOAD
        '                       LEFT OUTER JOIN TSPL_HEAD_LOAD_DCS ON TSPL_HEAD_LOAD_DCS.Document_No=TSPL_HEAD_LOAD.Document_No AND
        '                       TSPL_HEAD_LOAD_DCS.VLC_CODE=TSPL_HEAD_LOAD.VLC_CODE
        '                       GROUP BY TSPL_HEAD_LOAD.Document_No,TSPL_HEAD_LOAD_DCS.VLC_CODE)Headload on TSPL_VLC_MASTER_HEAD.VLC_Code = Headload.VLC_CODE "
        'BaseQry += " Left Outer Join TSPL_HEAD_LOAD_DCS On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_HEAD_LOAD_DCS.VLC_CODE " + Environment.NewLine
        'End If
        'Comment by balwinder on 26/03/2024 as TSPL_TRANSFER_TO_SAVING_DETAIL is not used in this query
        'BaseQry += " left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_VLC_MASTER_HEAD.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code " + Environment.NewLine
        BaseQry += " left outer join TSPL_PRICE_CHART_PLANNING on TSPL_PRICE_CHART_PLANNING.Planning_Code=TSPL_MILK_SRN_DETAIL.Price_Code " + Environment.NewLine
        BaseQry += " left join TSPL_CITY_MASTER  as MCC_City on MCC_City.city_code=TSPL_MCC_MASTER.City_code " + Environment.NewLine
        BaseQry += " left join TSPL_STATE_MASTER as MCC_State on MCC_State.STATE_CODE =TSPL_MCC_MASTER.State_Code " + Environment.NewLine
        BaseQry += " left join TSPL_Primary_Vehicle_Master on TSPL_Primary_Vehicle_Master.VEHICLE_CODE = TSPL_MILK_SRN_HEAD.VEHICLE_CODE 
left outer join  (select TSPL_PRICE_CHART_PLANNING_TSDDCF.Planning_Code, TSPL_PRICE_CHART_PLANNING_TSDDCF.Rate_Per,TSPL_PRICE_CHART_PLANNING_TSDDCF.Fixed_Rate from (select Planning_Code,max(SNo) as SNo from TSPL_PRICE_CHART_PLANNING_TSDDCF group by Planning_Code) TabMaxPlanning inner join TSPL_PRICE_CHART_PLANNING_TSDDCF on TSPL_PRICE_CHART_PLANNING_TSDDCF.Planning_Code=TabMaxPlanning.Planning_Code and TSPL_PRICE_CHART_PLANNING_TSDDCF.SNo=TabMaxPlanning.SNo) as Tab_TSPL_PRICE_CHART_PLANNING_TSDDCF on  Tab_TSPL_PRICE_CHART_PLANNING_TSDDCF.Planning_Code=TSPL_MILK_SRN_DETAIL.Price_Code
left join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_MILK_SRN_DETAIL.Item_Code and  TSPL_ITEM_UOM_DETAIL.UOM_Code='LTR'"


        BaseQry += " left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_MILK_PURCHASE_INVOICE_Head.Comp_Code " + Environment.NewLine
        BaseQry += " left join (select distinct FAT_Pers,SNF_Pers,Ratio as Fat_ratio,SNF_Ratio, Milk_Rate,TSPL_MILK_PRICE_MASTER.Price_Code,TSPL_FAT_SNF_UPLOADER_MASTER.code,TSPL_FAT_SNF_UPLOADER_MASTER.CircularNo    from TSPL_FAT_SNF_UPLOADER_MASTER inner join  TSPL_MILK_PRICE_MASTER  on TSPL_MILK_PRICE_MASTER.Price_Code=TSPL_FAT_SNF_UPLOADER_MASTER.Price_Code) as  Price_Chart    on TSPL_MILK_SRN_DETAIL.Price_Code=Price_Chart.Code "
        BaseQry += " left outer join TSPL_VILLAGE_MASTER on TSPL_VILLAGE_MASTER.Village_Code = TSPL_VLC_MASTER_HEAD.Village_Code " + Environment.NewLine +
        " left join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Uploader_TR_No  
        Left Join(select max(created_date) As created_date, VLC_Code, VSP_CODE,max(VSP_NAME)VSP_NAME, sum(Total_EMP_Amount) As Total_EMP_Amount, sum(Incentive_Amount) As Incentive_Amount, sum(Incentive_EMP_Amount) As Incentive_EMP_Amount, sum(EMP_Amount) As EMP_Amount, sum(Vsp_Own_System_Amount) As Vsp_Own_System_Amount, sum(Head_Load_Amount) As Head_Load_Amount, sum(Payable_Amount) As Payable_Amount, sum(Credit_Note_Amount)As Credit_Note_Amount, sum(Deduction_Amount) As Deduction_Amount, sum(Item_Issue_Amount) As Item_Issue_Amount, sum(Item_Issue_Return_Amount) As Item_Issue_Return_Amount, sum(MCC_Sale_Amount) As MCC_Sale_Amount , sum(MCC_Sale_Return_Amount) As MCC_Sale_Return_Amount,sum(Compulsory_Amount) as Compulsory_Amount,MAX(Area_Location_Code)Area_Location_Code,max(Payee_Joint_Account_No)Payee_Joint_Account_No,max(Payee_Joint_Branch_Name)Payee_Joint_Branch_Name,max(Zone_Code)Zone_Code,max(Bank_Account_No_Saving)Bank_Account_No_Saving from (Select TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code,TSPL_PAYMENT_PROCESS_HEAD.Created_Date  , TSPL_PAYMENT_PROCESS_DETAIL.Incentive_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Incentive_EMP_Amount, TSPL_PAYMENT_PROCESS_DETAIL.EMP_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Vsp_Own_System_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Total_EMP_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount, TSPL_VLC_MASTER_HEAD.VLC_Code, TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE,TSPL_PAYMENT_PROCESS_DETAIL.VSP_NAME, TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Credit_Note_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Deduction_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Return_Amount, TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Amount, TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Amount, TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Return_Amount,TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Name,TSPL_PAYMENT_PROCESS_DETAIL.Zone_Code,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving  from TSPL_PAYMENT_PROCESS_DETAIL"
        BaseQry += " left join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No =TSPL_PAYMENT_PROCESS_DETAIL.Doc_No"
        BaseQry += " left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader =TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader " & whrcls1 & ""
        BaseQry += " ) as pp group by VSP_CODE,VLC_Code"
        BaseQry += " ) as PaymentProcess on "
        BaseQry += "  PaymentProcess.vsp_code = TSPL_MILK_PURCHASE_INVOICE_Head.vsp_code And PaymentProcess.VLC_Code = (case when TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No Is Not null then TSPL_MILK_SRN_HEAD.VLC_Code else TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code end) " + Environment.NewLine

        BaseQry += "    left outer join (select TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE , TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No as BillNo , convert(varchar,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_Date,103) as BillDate, TSPL_PAYMENT_PROCESS_DETAIL.SNo as BILLSRL, TSPL_PAYMENT_PROCESS_DETAIL.Doc_No, TSPL_PAYMENT_PROCESS_DETAIL. is_Hold_Payment_Process, TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code from TSPL_PAYMENT_PROCESS_DETAIL where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in ( " + strDocNo + " )"
        If clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += "and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE In (" + strVSPCode + ")"
        End If
        BaseQry += " ) as TBL_BILL_DETAILS on TBL_BILL_DETAILS.VSP_CODE =TSPL_MILK_PURCHASE_INVOICE_Head.vsp_code"

        If AreaWiseBilling = True Then
            BaseQry += "  Left Outer Join(select MAX(TSPL_PAYMENT_PROCESS_HEAD.Doc_No)Doc_No,MAX(tspl_location_master.Location_Desc)Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD  
                         left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code
						 GROUP BY tspl_location_master.Location_Code
                         ) xxxSetLocation On xxxSetLocation.Location_Code=PaymentProcess.area_Location_code"
        Else
            BaseQry += " left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_Code " + Environment.NewLine
        End If


        BaseQry += "  LEFT OUTER JOIN TSPL_MILK_SHIFT_UPLOADER_DETAIL ON TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No   "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") <> CompairStringResult.Equal Then
            BaseQry += " left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO= TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO " + Environment.NewLine
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
            BaseQry += " left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO= TSPL_MILK_SRN_HEAD.ROUTE_CODE  " + Environment.NewLine
        End If
        If AreaWiseBilling = True Then
            BaseQry += "  left outer join (select TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,TSPL_BULK_ROUTE_MASTER.ROUTE_NO from TSPL_BULK_ROUTE_MASTER) as xxxRoutename on xxxRoutename.ROUTE_NO= TSPL_VLC_MASTER_HEAD.route_code " + Environment.NewLine
        End If
        BaseQry += "left outer join (select VSP_Code,max(Item_Desc) as Item_Desc, sum([Amount]) as [Amount] from (
			 select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Item_Desc,(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_COMPULSORY 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No in (" + strDocNo + ")"
        If clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_Code In (" + strVSPCode + ") "
        End If
        BaseQry += " )x group by VSP_Code)TabCompulsory on TabCompulsory.VSP_Code=TBL_BILL_DETAILS.VSP_CODE"

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            BaseQry += " 	left outer join (select VSP_Code,max(Item_Desc) as Item_Desc, sum([Amount]) as [Amount] from (
			 select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Item_Desc,(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ") 
   "
            If clsCommon.myLen(strVSPCode) > 0 Then
                BaseQry += " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code in (" + strVSPCode + ") "
            End If
            BaseQry += "  and TSPL_DCS_ADDITION_DEDUCTION.Description='DCS OVERHEAD' ) x group by VSP_Code)TabOverheadSaving on TabOverheadSaving.VSP_Code=TBL_BILL_DETAILS.VSP_CODE "

            BaseQry += " left outer join (select VSP_Code,max(Item_Desc) as Item_Desc, sum([Amount]) as [Amount] from (
			 select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Item_Desc,(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_COMPULSORY 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No in (" + strDocNo + ") "
             If clsCommon.myLen(strVSPCode) > 0 Then
                BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_Code In  (" + strVSPCode + ") "
            End If


            BaseQry += " and TSPL_DCS_ADDITION_DEDUCTION.Description='DCS OVERHEAD' )x group by VSP_Code)TabOverheadCompulsory on TabOverheadCompulsory.VSP_Code=TBL_BILL_DETAILS.VSP_CODE  "
        End If


        BaseQry += " left outer join (select VSP_Code,max(Item_Desc) as Item_Desc, sum([Amount]) as [Amount] from (
			 select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Item_Desc,(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ") "
        If clsCommon.myLen(strVSPCode) > 0 Then
            BaseQry += " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code in (" + strVSPCode + ") "
        End If
        BaseQry += " )x group by VSP_Code)TabSaving on TabSaving.VSP_Code=TBL_BILL_DETAILS.VSP_CODE"

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            BaseQry += " Left outer join (Select VSP_Code, sum(Trsfer_Amt) as Trsfer_Amt from (
                        Select  Vendor_Code As VSP_Code,Document_No,Document_Total As Trsfer_Amt  from TSPL_PAYMENT_PROCESS_SAVING
                        Left OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
                        Left OUTER JOIN TSPL_PAYMENT_PROCESS_HEAD ON TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                        where  2 = 2 And Against_TransferToSavingPKID Is Not NULL
                        And TSPL_PAYMENT_PROCESS_SAVING.Doc_No in (" + strDocNo + ") "
            If clsCommon.myLen(strVSPCode) > 0 Then
                BaseQry += " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code in (" + strVSPCode + ") "
            End If
            BaseQry += " )x group by VSP_Code) TrnferAmt on TrnferAmt.VSP_Code=TBL_BILL_DETAILS.VSP_CODE "
        End If


        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal AndAlso PaymentProcessInHindi = True AndAlso isprintHindi Then
            BaseQry += " left outer join (SELECT ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition,zz.Addition_Hindi,SUM(ZZ.Amount) AS Amount,max(ManAddDed) as ManAddDed FROM(
select  Final.VSP_Uploader_Code, Final.VSP_Code ,'' as Vendor_NAME,Final.Item_Desc as Addition,final.Addition_Hindi, sum(Amount) as [Amount],max(ManAddDed) as ManAddDed  from (
select TSPL_VENDOR_INVOICE_HEAD.Document_No, TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VSP_Code ,TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VLC_Code_VLC_Uploader, 
case when isnull(TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type,'')='Addition' then TSPL_MULTIPLE_DEDUCTION_DETAIL.Deduction_Desc 
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='DCS-QAT' THEN 'QAP'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='DCS-LYT' THEN 'Loyalty'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='OWD-CRE' THEN 'Own BMC Expanse'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='OWD-CRD' THEN 'FAT SNF SHORTAGE'
when TSPL_DCS_ADDITION_DEDUCTION.Description is null then TSPL_VENDOR_INVOICE_HEAD.RefDocType
else TSPL_DCS_ADDITION_DEDUCTION.Description  end as Item_Desc ,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindi, 0 as FAT_Amount,0 as SNF_Amount , TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount as Amount ,Convert (varchar,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) as  AP_Invoice_Date,  0 as Is_Default_Pashu_Vikas_Kos, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code
,TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction,case when TSPL_MULTIPLE_DEDUCTION_head.Document_No is not null then 1 else 0 end as ManAddDed
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE   
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
left outer join TSPL_MULTIPLE_DEDUCTION_DETAIL on TSPL_MULTIPLE_DEDUCTION_DETAIL.Against_Deduction_DocNo = TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_MULTIPLE_DEDUCTION_head on TSPL_MULTIPLE_DEDUCTION_head.Document_No = TSPL_MULTIPLE_DEDUCTION_DETAIL.Document_No 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE
where  TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No in (" + strDocNo + ") "
            If clsCommon.myLen(strVSPCode) > 0 Then
                BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_Code in (" + strVSPCode + ") "
            End If

            BaseQry += " ) Final group by Final.VSP_Uploader_Code, Final.VSP_Code , Final.Item_Desc,final.Addition_Hindi  

union all

SELECT TT.VSP_Uploader_Code,TT.VSP_Code,TT.Vendor_NAME,coalesce (mapping.mmDescription, TT.Addition) AS Addition,coalesce (mapping.HindiDescription, TT.Addition_Hindi) AS Addition_Hindi
,TT.Amount,0 as ManAddDed  FROM (
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Addition,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindi,
TSPL_DCS_ADDITION_DEDUCTION.Code,
(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_SAVING.Doc_No in  (" + strDocNo + ")"
            If clsCommon.myLen(strVSPCode) > 0 Then
                BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_Code in (" + strVSPCode + ") "
            End If
            BaseQry += " )TT
left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,MAPPING.Description_Hindi HindiDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION on  DEDUCTION.Code=MAPPING.MappingCode WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code)ZZ  where Addition='P_DCS ADM. EXP' GROUP BY  ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition,zz.Addition_Hindi )TabMilkPurchase on TabMilkPurchase.VSP_Code=TBL_BILL_DETAILS.VSP_CODE "
        End If


        BaseQry += "  " & whrcls & " "
        Dim dt As New DataTable
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            sQuery = BaseQry + " order by  cast(VLC_Code_VLC_Uploader as int),BillNo,convert(datetime,coalesce(TSPL_MILK_SRN_HEAD.DOC_DATE,TSPL_MILK_SRN_head.DOC_DATE),103),shift desc"
        ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
            sQuery = BaseQry + " order by  cast(Mcc_Code_VLC_Uploader as int) asc, convert(datetime,coalesce(TSPL_MILK_SRN_HEAD.DOC_DATE,TSPL_MILK_SRN_head.DOC_DATE),103)"
        Else
            sQuery = BaseQry + " order by " + IIf(objCommonVar.CurrentCompanyCode = "RCDF", " ", "BILLSRL asc,") + " vsp_code,convert(datetime,coalesce(TSPL_MILK_SRN_HEAD.DOC_DATE,TSPL_MILK_SRN_head.DOC_DATE),103),shift desc"
        End If

        If QryForShowData = True Then
            Return BaseQry
        End If
        dt = clsDBFuncationality.GetDataTable(sQuery)

        Dim dtFATNSFDCNote As New DataTable
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDF") = CompairStringResult.Equal Then

            sQuery = "select Vendor_CODE as VSP_Code,trans_type,round( (coalesce(Amount,0)*tab.FAT_Amount)/(tab.FAT_Amount+tab.SNF_Amount),0) as  FAT_Amount,Amount - round( (coalesce(Amount,0)*tab.FAT_Amount)/(tab.FAT_Amount+tab.SNF_Amount),0) as SNF_Amount,coalesce(Amount,0) as Amount from (" + Environment.NewLine +
"select doc_no,Vendor_CODE,SNo,trans_type,sum(Amount) as Amount from( " + Environment.NewLine +
"select case when RefDocType ='VSP-COM' then 'EMP' else (case when RefDocType ='PRO-VFC' then 'PRO ADD.' else 'OTHER ADD.' end ) end as trans_type,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.doc_no,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount,1 as   Show_FAT_SNF , case when RefDocType ='VSP-COM' then 1 else (case when RefDocType ='PRO-VFC' then 3 else 2 end ) end as SNo" + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_CREDIT_NOTE   " + Environment.NewLine +
"left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No" + Environment.NewLine +
"where TSPL_VENDOR_INVOICE_HEAD.RefDocType not in ('VSP-DIT') and TSPL_VENDOR_INVOICE_HEAD.Description = 'AP Credit Note For VSP Commission' and TSPL_PAYMENT_PROCESS_CREDIT_NOTE.doc_no  in (" + strDocNo + ")" + Environment.NewLine +
"union all " + Environment.NewLine +
"select case when RefDocType='VSP-QLT' then 'CBD' else case when RefDocType='MILK-REJ' then 'Rejection Ch.' else (case when RefDocType ='PRO-VFD' then 'PRO DED.' else 'OTHER DED.' end ) end end as trans_type,TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no,TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No ,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE, " + Environment.NewLine +
" TSPL_PAYMENT_PROCESS_DEDUCTION.Amount*(-1) as Amount ,case when RefDocType='MILK-REJ' then 1 else TSPL_DEDUCTION_MASTER.Show_FAT_SNF  end Show_FAT_SNF, case when RefDocType='VSP-QLT' then 4 else case when RefDocType='MILK-REJ' then 5 else (case when RefDocType ='PRO-VFD' then 7 else 6 end ) end end as SNo" + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_DEDUCTION " + Environment.NewLine +
"left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No" + Environment.NewLine +
"left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no and TSPL_VENDOR_INVOICE_DETAIL.Detail_Line_No=1" + Environment.NewLine +
"left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode" + Environment.NewLine +
" where TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no in ( " + strDocNo + " ) and (TSPL_DEDUCTION_MASTER.Show_FAT_SNF=1 or RefDocType='MILK-REJ')" + Environment.NewLine +
")xxx group by doc_no, Vendor_CODE,SNo,trans_type" + Environment.NewLine +
        " ) as final " + Environment.NewLine +
        " left join TSPL_PAYMENT_PROCESS_head on TSPL_PAYMENT_PROCESS_head.doc_no=final.doc_no" + Environment.NewLine +
        "left outer join (select vsp_code,sum(FAT_Amount) as FAT_Amount,sum(SNF_Amount) as SNF_Amount from (" + BaseQry + ")x group by vsp_code)Tab on tab.VSP_CODE=final.Vendor_CODE"
            sQuery += " " & whrclsItemWise & " "
            dtFATNSFDCNote = clsDBFuncationality.GetDataTable(sQuery)
        End If


        sQuery = " select * from ( select Customer_CODE as VSP_Code,trans_type,Item_Code as VLC_Code_VLC_Uploader ,Item_Desc,0 as FAT_Amount,0 as SNF_Amount,coalesce(Amount,0) as Amount , AP_Invoice_Date,  isnull(TSPL_DEDUCTION_MASTER.Is_Default_Pashu_Vikash_Kos,0) as Is_Default_Pashu_Vikash_Kos, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code from (" + Environment.NewLine +
        "select max(trans_type) as trans_type,max(doc_no) as doc_no,max(shipment_doc_no) as AP_Invoice_No,(item_code) as	item_code	, max(Item_Desc) as	Item_Desc,customer_code	,sum(Paymnet_Amount) as	Paymnet_Amount	,sum(Amount) as	Amount	,max(Show_FAT_SNF) as	Show_FAT_SNF , max( AP_Invoice_Date) as AP_Invoice_Date from(" + Environment.NewLine +
        "select 'MCC Sale' as trans_type,TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no,TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no,TSPL_SD_SHIPMENT_detail.item_code,tspl_item_master.item_desc,TSPL_PAYMENT_PROCESS_MCC_SALE.customer_code,TSPL_PAYMENT_PROCESS_MCC_SALE.amount as Paymnet_Amount,TSPL_SD_SHIPMENT_detail.Item_net_amt*(-1) as Amount ,0 as Show_FAT_SNF  , case when len (TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_Date) > 0 then convert (varchar,TSPL_PAYMENT_PROCESS_MCC_SALE.AR_Invoice_Date,103) else '' end as AP_Invoice_Date  " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_MCC_SALE " + Environment.NewLine +
"left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no " + Environment.NewLine +
"left join tspl_item_master on tspl_item_master.item_code=TSPL_SD_SHIPMENT_detail.item_code    where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in ( " + strDocNo + ")" + Environment.NewLine +
")xxx group by customer_code,item_code" + Environment.NewLine +
"union all " + Environment.NewLine +
"select 'MCC Sale Return' as trans_type,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.Doc_No ,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.return_doc_no,TSPL_SD_SALE_RETURN_DETAIL.item_code,tspl_item_master.item_desc,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.customer_code,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.amount as Paymnet_Amount ,TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt,0 as Show_FAT_SNF , case when len (TSPL_PAYMENT_PROCESS_MCC_SALE_Return.AR_Invoice_Date) > 0 then convert (varchar,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.AR_Invoice_Date,103) else '' end as AP_Invoice_Date " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_MCC_SALE_Return  " + Environment.NewLine +
"left join TSPL_SD_SALE_RETURN_DETAIL on TSPL_SD_SALE_RETURN_DETAIL.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_Return.return_doc_no " + Environment.NewLine +
"left join tspl_item_master on tspl_item_master.item_code=TSPL_SD_SALE_RETURN_DETAIL.item_code   " + Environment.NewLine +
"union all
select max(trans_type) as trans_type,max(doc_no) as doc_no,max(Payment_No) as AP_Invoice_No,'' as	item_code	,'Asset Lost' as Item_Desc,Vendor_Code as customer_code,0 as	Paymnet_Amount	,sum(Amount) as	Amount	,max(Show_FAT_SNF) as	Show_FAT_SNF,  '' as AP_Invoice_Date from(
select 'Asset Lost' as trans_type,TSPL_PAYMENT_PROCESS_ASSET_LOST.Doc_No,TSPL_PAYMENT_PROCESS_ASSET_LOST.Payment_No,TSPL_PAYMENT_PROCESS_ASSET_LOST.Vendor_Code,TSPL_PAYMENT_PROCESS_ASSET_LOST.Payment_Amount*(-1) as Amount ,0 as Show_FAT_SNF 
from TSPL_PAYMENT_PROCESS_ASSET_LOST 
 where TSPL_PAYMENT_PROCESS_ASSET_LOST.Doc_No  in ( " + strDocNo + " )
)xxx group by Vendor_Code" + Environment.NewLine +
"union all " + Environment.NewLine +
"select 'Item Issue' as trans_type,TSPL_PAYMENT_PROCESS_ITEM_ISSUE.Doc_No ,TSPL_PAYMENT_PROCESS_ITEM_ISSUE.item_issue_doc_no,TSPL_VSPItem_DETAIL.item_code,tspl_item_master.item_desc,TSPL_PAYMENT_PROCESS_ITEM_ISSUE.Vendor_code,TSPL_PAYMENT_PROCESS_ITEM_ISSUE.amount as Paymnet_Amount ,TSPL_VSPItem_DETAIL.Item_Net_Amt*(-1) as Item_Net_Amt ,0 as Show_FAT_SNF , '' as AP_Invoice_Date " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_ITEM_ISSUE  " + Environment.NewLine +
"left join TSPL_VSPItem_DETAIL on TSPL_VSPItem_DETAIL.doc_no=TSPL_PAYMENT_PROCESS_ITEM_ISSUE.item_issue_doc_no " + Environment.NewLine +
"left join tspl_item_master on tspl_item_master.item_code=TSPL_VSPItem_DETAIL.item_code    " + Environment.NewLine +
"union all " + Environment.NewLine +
"select 'Item Issue Return' as trans_type,TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.Doc_No ,TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.item_issue_return_no,TSPL_VSPItem_DETAIL.item_code,tspl_item_master.item_desc,TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.Vendor_code,TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.amount as Paymnet_Amount ,TSPL_VSPItem_DETAIL.Item_Net_Amt,0 as Show_FAT_SNF  , '' as AP_Invoice_Date " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN  " + Environment.NewLine +
"left join TSPL_VSPItem_DETAIL on TSPL_VSPItem_DETAIL.doc_no=TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.item_issue_return_no " + Environment.NewLine +
"left join tspl_item_master on tspl_item_master.item_code=TSPL_VSPItem_DETAIL.item_code    " + Environment.NewLine +
"union all " + Environment.NewLine +
"select max(trans_type) as trans_type,max(doc_no) as doc_no,max(AP_Invoice_No) as AP_Invoice_No,max(Vendor_CODE) as	Vendor_CODE	,Item_Desc,Vendor_CODE	,sum(Paymnet_Amount) as	Paymnet_Amount	,sum(Item_Net_AMount) as	Item_Net_AMount	,max(Show_FAT_SNF) as	Show_FAT_SNF, AP_Invoice_Date  from( " + Environment.NewLine +
"select RefDocType,'Deduction' as trans_type,TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no,TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No ,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE, " + Environment.NewLine +
"case when TSPL_DEDUCTION_MASTER.Code is not null then TSPL_DEDUCTION_MASTER.Description else TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Desc end as Item_Desc ,0 as Paymnet_Amount ,TSPL_PAYMENT_PROCESS_DEDUCTION.Amount*(-1) as Item_Net_AMount ,TSPL_DEDUCTION_MASTER.Show_FAT_SNF, convert (varchar,TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_Date,103) as AP_Invoice_Date   " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_DEDUCTION " + Environment.NewLine +
"left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No" + Environment.NewLine +
"left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no and TSPL_VENDOR_INVOICE_DETAIL.Detail_Line_No=1" + Environment.NewLine +
"left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode" + Environment.NewLine + " where TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no in ( " + strDocNo + " ) and not ( RefDocType='MILK-REJ')" + Environment.NewLine +     ' TSPL_DEDUCTION_MASTER.Show_FAT_SNF=1 or
")xxx group by RefDocType,Vendor_CODE,Item_Desc,AP_Invoice_Date " + Environment.NewLine
        sQuery += " ) as final "
        sQuery += " left join TSPL_PAYMENT_PROCESS_head on TSPL_PAYMENT_PROCESS_head.doc_no=final.doc_no 
left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Description=final.Item_Desc "
        '"left outer join (select distinct VSP_Code ,VLC_Code_VLC_Uploader from TSPL_VLC_MASTER_HEAD where TSPL_VLC_MASTER_HEAD.VSP_Code IN (" + strVSPCode + ") ) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code = Item_Code"
        sQuery += "left outer join TSPL_VLC_MASTER_HEAD  on TSPL_VLC_MASTER_HEAD.VSP_Code = Item_Code "
        sQuery += " " & whrclsItemWise & "  Union All  "
        sQuery += " select TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VSP_Code ,TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type   , TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VLC_Code_VLC_Uploader, TSPL_MULTIPLE_DEDUCTION_DETAIL.Deduction_Desc as Item_Desc , 0 as FAT_Amount,0 as SNF_Amount , TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount as Amount ,Convert (varchar,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) as  AP_Invoice_Date,  0 as Is_Default_Pashu_Vikas_Kos, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE   
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_MULTIPLE_DEDUCTION_DETAIL on TSPL_MULTIPLE_DEDUCTION_DETAIL.Against_Deduction_DocNo = TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_MULTIPLE_DEDUCTION_head on TSPL_MULTIPLE_DEDUCTION_head.Document_No = TSPL_MULTIPLE_DEDUCTION_DETAIL.Document_No 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE
where  TSPL_PAYMENT_PROCESS_CREDIT_NOTE.doc_no in ( " + strDocNo + ")
and TSPL_VENDOR_INVOICE_HEAD.Description <> 'AP Credit Note For VSP Commission'
and TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type = 'Addition' 
) Final order by Is_Default_Pashu_Vikash_Kos desc , trans_type desc "
        Dim dtgv As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        'sQuery = "select  TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE,convert(varchar, TSPL_MILK_REJECT_HEAD.DOC_DATE,103) as DOC_DATE,TSPL_MILK_REJECT_HEAD.SHIFT,TSPL_MILK_REJECT_DETAIL.Reject_Type,TSPL_MILK_REJECT_DETAIL.MILK_WEIGHT,UOM_Code,TSPL_MILK_REJECT_DETAIL.RATE, TSPL_PAYMENT_PROCESS_DEDUCTION.Amount   " + Environment.NewLine +
        '"from TSPL_PAYMENT_PROCESS_DEDUCTION " + Environment.NewLine +
        '"left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No" + Environment.NewLine +
        '"left outer join TSPL_MILK_REJECT_DETAIL on TSPL_MILK_REJECT_DETAIL.DOC_CODE=TSPL_VENDOR_INVOICE_HEAD.RefDocNo and TSPL_MILK_REJECT_DETAIL.SAMPLE_NO=TSPL_VENDOR_INVOICE_HEAD.Ref_SNo " + Environment.NewLine +
        '"left outer join TSPL_MILK_REJECT_HEAD on TSPL_MILK_REJECT_HEAD.DOC_CODE=TSPL_MILK_REJECT_DETAIL.DOC_CODE" + Environment.NewLine +
        '"where TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no in ( " + strDocNo + ") and RefDocType='MILK-REJ'"
        'Dim dtRej As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        sQuery = " select  TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight,TSPL_MILK_SRN_DETAIL.UOM_Code,TSPL_MILK_SRN_DETAIL.RATE, TSPL_PAYMENT_PROCESS_DEDUCTION.Amount   
                   from TSPL_PAYMENT_PROCESS_DEDUCTION 
                   left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No
                   left outer join TSPL_MILK_SRN_HEAD on TSPL_MILK_SRN_HEAD.SAMPLE_NO = TSPL_VENDOR_INVOICE_HEAD.Ref_SNo
                   left outer join TSPL_MILK_SRN_DETAIL on TSPL_MILK_SRN_DETAIL.DOC_CODE=TSPL_MILK_SRN_HEAD.DOC_CODE
                   left outer join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No = TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No
                   where TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no in (" + strDocNo + ") and RefDocType='MILK-REJ'"
        Dim dtRej As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        sQuery = "select TSPL_PAYMENT_PROCESS_MCC_SALE.customer_code,  TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no,TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no,TSPL_SD_SHIPMENT_detail.item_code,tspl_item_master.item_desc" + Environment.NewLine +
",TSPL_SD_SHIPMENT_detail.Qty as MILK_WEIGHT,case when isnull(TSPL_SD_SHIPMENT_detail.Qty,0)<=0 then 0 else cast(TSPL_PAYMENT_PROCESS_MCC_SALE.Amount/TSPL_SD_SHIPMENT_detail.Qty as decimal(18,2)) end as RATE,TSPL_PAYMENT_PROCESS_MCC_SALE.Amount  " + Environment.NewLine +
"from TSPL_PAYMENT_PROCESS_MCC_SALE " + Environment.NewLine +
"left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no " + Environment.NewLine +
"left join tspl_item_master on tspl_item_master.item_code=TSPL_SD_SHIPMENT_detail.item_code    " + Environment.NewLine +
"where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in (" + strDocNo + ")"
        Dim dtSale As DataTable = clsDBFuncationality.GetDataTable(sQuery)



        Dim dtSaving As DataTable = Nothing
        Dim sQuerySaving As String = "SELECT TT.VSP_Uploader_Code,TT.VSP_Code,TT.Vendor_NAME,coalesce (mapping.mmDescription, TT.Addition) AS Addition,coalesce (mapping.HindiDescription, TT.Addition_Hindi) AS Addition_Hindi
,TT.Amount,0 as ManAddDed  FROM (
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Addition,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindi,
TSPL_DCS_ADDITION_DEDUCTION.Code,
(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount]  from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where  TSPL_PAYMENT_PROCESS_SAVING.Doc_No in  (" + strDocNo + ") "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            sQuerySaving += " and TSPL_DCS_ADDITION_DEDUCTION.Description is not null "
        End If

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
            sQuerySaving += " and TSPL_DCS_ADDITION_DEDUCTION.Saving = 2 "
        End If

        sQuerySaving += " )TT
left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,MAPPING.Description_Hindi HindiDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION on  DEDUCTION.Code=MAPPING.MappingCode WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            Dim sQuerySavingData As String = "SELECT ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition,SUM(ZZ.Amount) AS Amount
                 FROM (" + sQuerySaving + " )ZZ  GROUP BY ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition "
            dtSaving = clsDBFuncationality.GetDataTable(sQuerySavingData)
        End If


        sQuery = "SELECT ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition,zz.Addition_Hindi,SUM(ZZ.Amount) AS Amount,max(ManAddDed) as ManAddDed FROM(
select  Final.VSP_Uploader_Code, Final.VSP_Code ,'' as Vendor_NAME,Final.Item_Desc as Addition,final.Addition_Hindi, sum(Amount) as [Amount],max(ManAddDed) as ManAddDed  from (
select TSPL_VENDOR_INVOICE_HEAD.Document_No, TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VSP_Code ,TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as VLC_Code_VLC_Uploader, 
case when isnull(TSPL_MULTIPLE_DEDUCTION_DETAIL.trans_type,'')='Addition' then TSPL_MULTIPLE_DEDUCTION_DETAIL.Deduction_Desc 
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='DCS-QAT' THEN 'QAP'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='DCS-LYT' THEN 'Loyalty'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='OWD-CRE' THEN 'Own BMC Expanse'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='OWD-CRD' THEN 'FAT SNF SHORTAGE'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='CHC-CRE' THEN 'BMC CHARGES'
WHEN TSPL_VENDOR_INVOICE_HEAD.RefDocType='CRE-DHL' THEN 'Head Load'
when TSPL_DCS_ADDITION_DEDUCTION.Description is null then TSPL_VENDOR_INVOICE_HEAD.RefDocType
else TSPL_DCS_ADDITION_DEDUCTION.Description  end as Item_Desc ,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindis,
COALESCE(TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi, TSPL_DEDUCTION_MASTER.Description_Hindi) AS Addition_Hindi,
0 as FAT_Amount,0 as SNF_Amount , TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount as Amount ,Convert (varchar,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) as  AP_Invoice_Date,  0 as Is_Default_Pashu_Vikas_Kos, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code
,TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction,case when TSPL_MULTIPLE_DEDUCTION_head.Document_No is not null then 1 else 0 end as ManAddDed
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE   
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.document_no=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
left outer join TSPL_MULTIPLE_DEDUCTION_DETAIL on TSPL_MULTIPLE_DEDUCTION_DETAIL.Against_Deduction_DocNo = TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_MULTIPLE_DEDUCTION_head on TSPL_MULTIPLE_DEDUCTION_head.Document_No = TSPL_MULTIPLE_DEDUCTION_DETAIL.Document_No 
left outer join TSPL_DEDUCTION_MASTER ON TSPL_DEDUCTION_MASTER.Code = TSPL_MULTIPLE_DEDUCTION_DETAIL.DeductionCode
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE
where  TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No in ( " + strDocNo + " ) and TSPL_VENDOR_INVOICE_HEAD.RefDocType not in ('CAP-MSN-CDCS','CAP-MSN','CAP-OMSN') and (Hide_In_Milk_Bill_Print = '0' OR Hide_In_Milk_Bill_Print IS NULL)
) Final group by Final.VSP_Uploader_Code, Final.VSP_Code , Final.Item_Desc,final.Addition_Hindi "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal And clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") <> CompairStringResult.Equal And clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") <> CompairStringResult.Equal And clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") <> CompairStringResult.Equal And clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") <> CompairStringResult.Equal Then
            sQuery = sQuery + " union all " + sQuerySaving
        End If

        sQuery = sQuery + " union all
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME,'"
        sQuery = sQuery + clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.HeadLoadDescriptionInPaymentProcessPrint, clsFixedParameterCode.HeadLoadDescriptionInPaymentProcessPrint, Nothing))
        sQuery = sQuery + "' as Addition,'' as Addition_Hindi,TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount as Amount,0 as ManAddDed 
from TSPL_PAYMENT_PROCESS_DETAIL 
INNER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
where Document_Type='C' and RefDocType='Milk_HE' and TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ")"

        sQuery += " Union all
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE,'' as Vendor_NAME,
TSPL_DCS_ADDITION_DEDUCTION.Description as Addition,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindi,
TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount as Amount,0 as ManAddDed  
from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE
left outer join TSPL_MILK_PURCHASE_INVOICE_HEAD on TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.DCS_Addition_Deduction
--Left outer JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No=TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo
--left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_Code
left outer join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No =TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo
where TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Nature_Type='C' and TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") "

        sQuery += " union all
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME
,TSPL_DEDUCTION_MASTER.Description as Addition,'' as Addition_Hindi ,(TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Amount) as [Amount] ,0 as ManAddDed
from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Customer_CODE
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no
left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Doc_No in  (" + strDocNo + ") "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") <> CompairStringResult.Equal Then
            sQuery += "union all
SELECT TT.VSP_Uploader_Code,TT.VSP_Code,TT.Vendor_NAME,coalesce (mapping.mmDescription, TT.Addition) AS Addition ,TT.Addition_Hindi ,TT.Amount,0 as ManAddDed  FROM (
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_VLC_MASTER_HEAD.VSP_Code,'' as Vendor_NAME
,TSPL_DCS_ADDITION_DEDUCTION.Description as Addition,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Addition_Hindi ,TSPL_DCS_ADDITION_DEDUCTION.Code,
(TSPL_VENDOR_INVOICE_HEAD.Document_Total) as [Amount] 
from TSPL_PAYMENT_PROCESS_COMPULSORY left join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No 
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.document_no=TSPL_VENDOR_INVOICE_HEAD.document_no
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction
where TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No in (" + strDocNo + ") and TSPL_VENDOR_INVOICE_HEAD.RefDocType not in ('CAP-MSN-CDCS','CAP-MSN','CAP-OMSN') "

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                sQuery += " and TSPL_DCS_ADDITION_DEDUCTION.Saving <> 2 "
            End If

            sQuery += ")TT
left join (select MAPPING.Code mmCode,MAPPING.Description mmDescription,DEDUCTION.CODE AS ddCode from TSPL_DCS_ADDITION_DEDUCTION as MAPPING
left join TSPL_DCS_ADDITION_DEDUCTION as DEDUCTION on  DEDUCTION.Code=MAPPING.MappingCode WHERE  len(isnull(MAPPING.MappingCode,''))>0)mapping on mapping.ddCode=TT.Code "
        End If

        sQuery += ")ZZ  where Addition!='Notview' GROUP BY  ZZ.VSP_Uploader_Code,ZZ.VSP_Code,ZZ.Vendor_NAME,ZZ.Addition,zz.Addition_Hindi"
        Dim dtAddition As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        sQuery = "select Final.VSP_Uploader_Code, Final.Vendor_CODE, max(Vendor_NAME)Vendor_NAME,
Max(case when Ded_Desc IS NULL OR len(Ded_Desc)<=0 then Ded_Code else Ded_Desc end ) as Ded_Code,
COALESCE(max(Ded_Code_Hindi), max(Hindi))  AS Ded_Code_Hindi, 
sum(Amount) as [Amount],max(ManAddDed) as ManAddDed,sum(Reduce_Deduc_Amt)Reduce_Deduc_Amt from (
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE, TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_NAME,
COALESCE(TSPL_DEDUCTION_MASTER.Code, TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code) AS Ded_Code,
COALESCE(TSPL_DEDUCTION_MASTER.Description, TSPL_PAYMENT_PROCESS_DEDUCTION.ded_desc) as Ded_Desc,
COALESCE(TSPL_DEDUCTION_MASTER.Description_Hindi, DedMaster.Description_Hindi)  as Ded_Code_Hindi,"

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
            sQuery += " TSPL_PAYMENT_PROCESS_DEDUCTION.Amount as Amount "
        Else
            sQuery += " (TSPL_PAYMENT_PROCESS_DEDUCTION.Amount-TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt) as Amount "
        End If

        sQuery += ",case when TSPL_MULTIPLE_DEDUCTION_head.Document_No is not null then 1 else 0 end as ManAddDed,TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi AS Hindi 
from TSPL_PAYMENT_PROCESS_DEDUCTION 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE
left outer join TSPL_MULTIPLE_DEDUCTION_DETAIL on TSPL_MULTIPLE_DEDUCTION_DETAIL.Against_Deduction_DocNo = TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No
left outer join TSPL_MULTIPLE_DEDUCTION_head on TSPL_MULTIPLE_DEDUCTION_head.Document_No = TSPL_MULTIPLE_DEDUCTION_DETAIL.Document_No 
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code
left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_DCS_ADDITION_DEDUCTION.Deduction
left outer join TSPL_DEDUCTION_MASTER DedMaster  on DedMaster.Code=TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code
left outer join TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No
where  "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            sQuery += " TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in (" + strDocNo + ") and isnull(TSPL_DEDUCTION_MASTER.Is_Transfer_To_Saving,0)=0 and TSPL_VENDOR_INVOICE_HEAD.Saving=0 and (Hide_In_Milk_Bill_Print = '0' OR Hide_In_Milk_Bill_Print IS NULL) "
        Else
            sQuery += " TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in (" + strDocNo + ") and TSPL_VENDOR_INVOICE_HEAD.RefDocType not in ('CAP-MSN-CDCS','CAP-MSN','CAP-OMSN') and (Hide_In_Milk_Bill_Print = '0' OR Hide_In_Milk_Bill_Print IS NULL) "
        End If

        sQuery += "   union all
select TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE as Vendor_CODE,VSP_NAME as Vendor_NAME,'TDS' as Ded_Code,'TDS' as Ded_Desc,'' as Ded_Code_Hindi,TSPL_PAYMENT_PROCESS_DETAIL.TDS_Amount as Amount,0 as ManAddDed,0 as Reduce_Deduc_Amt,'' as Hindi 
from TSPL_PAYMENT_PROCESS_DETAIL 
where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ") and TSPL_PAYMENT_PROCESS_DETAIL.TDS_Amount>0 
union all
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader,TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE,'' as Vendor_NAME,TSPL_DCS_ADDITION_DEDUCTION.Description as Ded_Code,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Ded_Desc,'' as Ded_Code_Hindi,TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount as Amount,0 as ManAddDed,0 as Reduce_Deduc_Amt,'' as Hindi  from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE
left outer join TSPL_MILK_PURCHASE_INVOICE_HEAD on TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.DCS_Addition_Deduction
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_Code 
left outer join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No =TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo
where TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Nature_Type='D' and TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount<>0 AND TSPL_PAYMENT_PROCESS_DETAIL.Doc_No in (" + strDocNo + ")
union ALL
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader AS VSP_Uploader_Code, TSPL_PAYMENT_PROCESS_MCC_SALE.customer_code as Vendor_CODE, 
TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_NAME  as Vendor_NAME,
TSPL_DEDUCTION_MASTER.Code as Ded_Code,TSPL_DEDUCTION_MASTER.Description as Ded_Desc,'' as Ded_Code_Hindi,(TSPL_PAYMENT_PROCESS_MCC_SALE.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE.Reduce_Deduc_Amt) as Amount,1 as ManAddDed,0 as Reduce_Deduc_Amt,'' as Hindi 
from TSPL_PAYMENT_PROCESS_MCC_SALE 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no
left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction
where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in (" + strDocNo + ") and (TSPL_PAYMENT_PROCESS_MCC_SALE.Amount-TSPL_PAYMENT_PROCESS_MCC_SALE.Reduce_Deduc_Amt)>0 
) Final where Amount>0 group by  final.VSP_Uploader_Code, Final.Vendor_CODE, Final.Ded_Code
--group by  final.VSP_Uploader_Code, Final.Vendor_CODE, Vendor_NAME,Case when isnull(Final.Deduction_Type,'')='' then Final.Ded_Code Else Final.Deduction_Type End,final.Hindi 

"
        Dim dtDeduction As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        sQuery = "select Final.VSP_Uploader_Code, Final.Vendor_CODE, Vendor_NAME, Max(Ded_Desc) as Ded_Code,Case When max(Hindi) is null then max(Description_Hindi) else max(Hindi) end as Hindi
, sum(Amount) as [Amount] from (
select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VSP_Uploader_Code,TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE,
TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_NAME, TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code,TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Desc,
TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt as Amount ,TSPL_DCS_ADDITION_DEDUCTION.Description_Hindi as Hindi ,TSPL_DEDUCTION_MASTER.Description_Hindi
from TSPL_PAYMENT_PROCESS_DEDUCTION 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code
left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_PAYMENT_PROCESS_DEDUCTION.Ded_Code
where  TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No in (" + strDocNo + ") 
union all

select 
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader AS VSP_Uploader_Code, TSPL_PAYMENT_PROCESS_MCC_SALE.customer_code as Vendor_CODE, 
TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_NAME  as Vendor_NAME,TSPL_DEDUCTION_MASTER.Code as Ded_Code,
TSPL_DEDUCTION_MASTER.Description as Ded_Desc,TSPL_PAYMENT_PROCESS_MCC_SALE.Reduce_Deduc_Amt as Amount,'' as Hindi,'' as Description_Hindi
from TSPL_PAYMENT_PROCESS_MCC_SALE 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no
left join TSPL_SD_SHIPMENT_detail on TSPL_SD_SHIPMENT_detail.document_code=TSPL_PAYMENT_PROCESS_MCC_SALE.shipment_doc_no and TSPL_SD_SHIPMENT_detail.Line_No=1
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SHIPMENT_detail.Item_Code
left outer join TSPL_DEDUCTION_MASTER  on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SHIPMENT_HEAD.Deduction

where TSPL_PAYMENT_PROCESS_MCC_SALE.doc_no  in (" + strDocNo + ") 
) Final group by  final.VSP_Uploader_Code, Final.Vendor_CODE, Vendor_NAME, Final.Ded_Code having sum(Amount)>0"
        Dim dtReduceDeduction As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        'CorrectionQuery

        sQuery = " Select * from ( Select TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No as Document_No, TSPL_PAYMENT_PROCESS_HEAD.To_Date as Doc_Date,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No,
TSPL_VENDOR_INVOICE_HEAD.Posting_Date as AP_Invoice_Date,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 'Debit' else 'Credit' end as Document_Type,TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,
TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount,0 as Reduce_Deduc_Amt, 
CAST(TSPL_MILK_SRN_HEAD.DOC_DATE AS DATE) AS SRNDate
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE
left join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
left  join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No
left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left join TSPL_MILK_SRN_HEAD on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_VENDOR_INVOICE_HEAD.RefDocNo
where 2=2  and RefDocType In('CAP-MSN-CDCS','CAP-MSN','CAP-OMSN') and TSPL_PAYMENT_PROCESS_CREDIT_NOTE.doc_no  in (" + strDocNo + ") 

union all
Select TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No as Document_No, TSPL_PAYMENT_PROCESS_HEAD.To_Date as Doc_Date,TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No ,
TSPL_VENDOR_INVOICE_HEAD.Posting_Date as AP_Invoice_Date,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 'Debit' else 'Credit' end as Document_Type,TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,
TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_DEDUCTION.Amount,
TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt,CAST(TSPL_MILK_SRN_HEAD.DOC_DATE AS DATE) AS SRNDate
from TSPL_PAYMENT_PROCESS_DEDUCTION
left join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No
left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
left  join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No
left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left join TSPL_MILK_SRN_HEAD on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_VENDOR_INVOICE_HEAD.RefDocNo
where 2=2 and RefDocType In('CAP-MSN-CDCS','CAP-MSN','CAP-OMSN') and TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no  in (" + strDocNo + ") )XX order by SRNDate  "
        Dim dtCorrection As DataTable = clsDBFuncationality.GetDataTable(sQuery)

        ''SUSPENSE QUERY

        sQuery = " Select *,case when Document_Type='C' Then 'Addition' else 'Deduction' end as TypeDoc from ( Select TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No as Document_No, TSPL_PAYMENT_PROCESS_HEAD.To_Date as Doc_Date,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No,
TSPL_VENDOR_INVOICE_HEAD.Posting_Date as AP_Invoice_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type,TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,
TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Amount,0 as Reduce_Deduc_Amt, 
CAST(TSPL_MILK_COLLECTION_DCS.Document_Date AS DATE) AS SRNDate
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE
left join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
left  join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No
left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left join TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT on TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT.PK_ID=TSPL_VENDOR_INVOICE_HEAD.RefDocNo
LEFT JOIN TSPL_MILK_COLLECTION_DCS ON TSPL_MILK_COLLECTION_DCS.Document_No=TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT.Document_No
where 2=2  and RefDocType In('SUS-ADJC') and TSPL_PAYMENT_PROCESS_CREDIT_NOTE.doc_no  in (" + strDocNo + ") 

union all
Select TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No as Document_No, TSPL_PAYMENT_PROCESS_HEAD.To_Date as Doc_Date,TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No ,
TSPL_VENDOR_INVOICE_HEAD.Posting_Date as AP_Invoice_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type,TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,
TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_DEDUCTION.Amount,
TSPL_PAYMENT_PROCESS_DEDUCTION.Reduce_Deduc_Amt,CAST(TSPL_MILK_COLLECTION_DCS.Document_Date AS DATE) AS SRNDate
from TSPL_PAYMENT_PROCESS_DEDUCTION
left join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_DEDUCTION.AP_Invoice_No
left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
left  join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No
left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left join TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT on TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT.PK_ID=TSPL_VENDOR_INVOICE_HEAD.RefDocNo
LEFT JOIN TSPL_MILK_COLLECTION_DCS ON TSPL_MILK_COLLECTION_DCS.Document_No=TSPL_MILK_COLLECTION_DCS_SUSPENSE_ADJUSTMENT.Document_No
where 2=2 and RefDocType In('SUS-ADJD') and TSPL_PAYMENT_PROCESS_DEDUCTION.doc_no  in (" + strDocNo + ") )XX order by SRNDate  "
        Dim dtsuspense As DataTable = clsDBFuncationality.GetDataTable(sQuery)


        If dt IsNot Nothing And dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            'If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDF") = CompairStringResult.Equal Then
            '    PDFPath = frmCRV.funsubreportWithdt(isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtgv, "crptMilkPurchaseBillPaymentProcess", "", Nothing, "SubMilkPurchaseBill.rpt", "Address.rpt", Nothing, "SubMilkPurchaseBillRejection.rpt", dtRej, "SubMilkPurchaseBillMCCSale.rpt", dtSale, "SubMilkPurchaseBillFATSNFDebitCreditNote.rpt", dtFATNSFDCNote)
            'Else
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                Dim dtAdditionFinance As DataTable = Nothing
                Dim rows As DataRow() = dtAddition.Select("ManAddDed=0")
                If rows Is Nothing OrElse rows.Length > 0 Then
                    dtAdditionFinance = rows.CopyToDataTable()
                Else
                    dtAdditionFinance = New DataTable()
                    dtAdditionFinance.Columns.Add("VSP_Uploader_Code", GetType(String))
                    dtAdditionFinance.Columns.Add("VSP_Code", GetType(String))
                    dtAdditionFinance.Columns.Add("Vendor_NAME", GetType(String))
                    dtAdditionFinance.Columns.Add("Addition", GetType(String))
                    dtAdditionFinance.Columns.Add("Amount", GetType(Decimal))
                    dtAdditionFinance.Columns.Add("ManAddDed", GetType(Integer))

                    Dim dr As DataRow = dtAdditionFinance.NewRow()
                    dr("VSP_Uploader_Code") = "XXXYYYZZZ"
                    dr("VSP_Code") = "XXXYYYZZZ"
                    dr("Vendor_NAME") = "XXXYYYZZZ"
                    dr("Addition") = "XXXYYYZZZ"
                    dr("Amount") = 0
                    dr("ManAddDed") = 0
                    dtAdditionFinance.Rows.Add(dr)
                End If

                Dim dtAdditionOther As DataTable = Nothing
                rows = dtAddition.Select("ManAddDed=1")
                If rows Is Nothing OrElse rows.Length > 0 Then
                    dtAdditionOther = rows.CopyToDataTable()
                Else
                    dtAdditionOther = New DataTable()
                    dtAdditionOther.Columns.Add("VSP_Uploader_Code", GetType(String))
                    dtAdditionOther.Columns.Add("VSP_Code", GetType(String))
                    dtAdditionOther.Columns.Add("Vendor_NAME", GetType(String))
                    dtAdditionOther.Columns.Add("Addition", GetType(String))
                    dtAdditionOther.Columns.Add("Amount", GetType(Decimal))
                    dtAdditionOther.Columns.Add("ManAddDed", GetType(Integer))

                    Dim dr As DataRow = dtAdditionOther.NewRow()
                    dr("VSP_Uploader_Code") = "XXXYYYZZZ"
                    dr("VSP_Code") = "XXXYYYZZZ"
                    dr("Vendor_NAME") = "XXXYYYZZZ"
                    dr("Addition") = "XXXYYYZZZ"
                    dr("Amount") = 0
                    dr("ManAddDed") = 1
                    dtAdditionOther.Rows.Add(dr)
                End If

                Dim dtDeductionFinance As DataTable = Nothing
                rows = dtDeduction.Select("ManAddDed=0")
                If rows Is Nothing OrElse rows.Length > 0 Then
                    dtDeductionFinance = rows.CopyToDataTable()
                Else
                    dtDeductionFinance = New DataTable()
                    dtDeductionFinance.Columns.Add("VSP_Uploader_Code", GetType(String))
                    dtDeductionFinance.Columns.Add("Vendor_CODE", GetType(String))
                    dtDeductionFinance.Columns.Add("Vendor_NAME", GetType(String))
                    dtDeductionFinance.Columns.Add("Ded_Code", GetType(String))
                    dtDeductionFinance.Columns.Add("Amount", GetType(Decimal))
                    dtDeductionFinance.Columns.Add("ManAddDed", GetType(Integer))

                    Dim dr As DataRow = dtDeductionFinance.NewRow()
                    dr("VSP_Uploader_Code") = "XXXYYYZZZ"
                    dr("Vendor_CODE") = "XXXYYYZZZ"
                    dr("Vendor_NAME") = "XXXYYYZZZ"
                    dr("Ded_Code") = "XXXYYYZZZ"
                    dr("Amount") = 0
                    dr("ManAddDed") = 0
                    dtDeductionFinance.Rows.Add(dr)
                End If

                Dim dtDeductionOther As DataTable = Nothing
                rows = dtDeduction.Select("ManAddDed=1")
                If rows Is Nothing OrElse rows.Length > 0 Then
                    dtDeductionOther = rows.CopyToDataTable()
                Else
                    dtDeductionOther = New DataTable()
                    dtDeductionOther.Columns.Add("VSP_Uploader_Code", GetType(String))
                    dtDeductionOther.Columns.Add("Vendor_CODE", GetType(String))
                    dtDeductionOther.Columns.Add("Vendor_NAME", GetType(String))
                    dtDeductionOther.Columns.Add("Ded_Code", GetType(String))
                    dtDeductionOther.Columns.Add("Amount", GetType(Decimal))
                    dtDeductionOther.Columns.Add("ManAddDed", GetType(Integer))

                    Dim dr As DataRow = dtDeductionOther.NewRow()
                    dr("VSP_Uploader_Code") = "XXXYYYZZZ"
                    dr("Vendor_CODE") = "XXXYYYZZZ"
                    dr("Vendor_NAME") = "XXXYYYZZZ"
                    dr("Ded_Code") = "XXXYYYZZZ"
                    dr("Amount") = 0
                    dr("ManAddDed") = 1
                    dtDeductionOther.Rows.Add(dr)
                End If

                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAdditionFinance, "crptMilkPurchaseBillPaymentProcessNewJPR", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeductionFinance, "subReduceDeduction.rpt", dtReduceDeduction, "subSaving.rpt", dtSaving, "SubAdditionOther.rpt", dtAdditionOther, "SubDeductionOther.rpt", dtDeductionOther, "SubCorrectionWise.rpt", dtCorrection, "SubMilkCorrectionDetail.rpt", dtsuspense)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewGNG", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal AndAlso PaymentProcessInHindi = True AndAlso isprintHindi Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewJDHInHindi", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewJDH", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewALW", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewKTA", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewTNKBharatpur", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewJAL", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
                '
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewBKN", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
                '
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewTNK", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)

                '
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewJHL", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewSKR", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDF") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewRJS", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BNS") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNewBNS", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal AndAlso PaymentProcessInHindi = True AndAlso isprintHindi Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessUDPInHindi", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessUDP", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal AndAlso PaymentProcessInHindi = True AndAlso isprintHindi Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessAJMInHindi", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessAJM", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessCHT", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            Else
                PDFPath = frmCRV.funsubreportWithdt(clsUserMgtCode.frmPaymentProcess, isPDFPath, CrystalReportFolder.MilkProcurement, dt, dtAddition, "crptMilkPurchaseBillPaymentProcessNew", "", Nothing, "subAddition.rpt", "subDeduction.rpt", dtDeduction, "subReduceDeduction.rpt", dtReduceDeduction)
            End If
            frmCRV = Nothing
        Else
            clsCommon.MyMessageBoxShow("No Data Found")
        End If


        '    frmCRV = Nothing
        'Else
        '    clsCommon.MyMessageBoxShow("No Data Found")
        'End If
        Return PDFPath
    End Function


    '========================================
End Class

Public Class clsPaymentProcessDetail
#Region "Variables"
    Public Area_Location_Code As String = ""
    Public PP_Detail_No As String = ""
    Public Doc_No As String = ""
    Public Is_select As Boolean = False
    Public SNo As String = ""
    Public Milk_Purchase_Invoice_No As String = ""
    Public Milk_Purchase_Invoice_Date As DateTime
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As DateTime
    Public VLC_CODE_Uploader As String = ""
    Public VLC_Name As String = ""
    Public MCC_Code As String = ""
    Public VSP_CODE As String = ""
    Public VSP_NAME As String = ""
    Public Main_VSP_CODE As String = ""
    Public Main_VSP_NAME As String = ""
    Public Payee_Joint_Name As String = ""
    Public Payee_Joint_Bank_Code As String = ""
    Public Payee_Joint_Bank_Name As String = ""
    Public Payee_Joint_Branch_Code As String = ""
    Public Payee_Joint_Branch_Name As String = ""
    Public Payee_Joint_Account_No As String = ""
    Public Payee_Joint_IFSC_Code As String = ""
    Public Bank_Code As String = ""
    Public Bank_Desc As String = ""
    Public Payment_Mode As String = ""
    Public Cheque_No As String = ""
    Public Payee_Joint_Name_Saving As String = ""
    Public Bank_Code_Saving As String = ""
    Public Bank_Desc_Saving As String = ""
    Public Bank_Account_No_Saving As String = ""
    Public Payment_Mode_Saving As String = ""
    Public Zone_Code As String = ""

    Public Milk_Qty As Decimal = 0
    Public Milk_Amount As Decimal = 0
    Public Incentive_Amount As Decimal = 0
    Public EMP_Amount As Decimal = 0
    Public Incentive_EMP_Amount As Decimal = 0
    Public Total_EMP_Amount As Decimal = 0
    Public Total As Decimal = 0
    Public Total_Invoice_Amount As Decimal = 0
    Public TDS_Amount As Decimal = 0
    Public Vsp_Own_System_Amount As Decimal = 0
    Public Head_Load_Amount As Decimal = 0
    Public Invoice_Deduction_Amount As Decimal = 0
    Public Reduce_Deduc_Amt As Decimal = 0
    Public MCC_Sale_Amount As Decimal = 0
    Public MCC_Sale_Return_Amount As Decimal = 0
    Public Item_Issue_Amount As Decimal = 0
    Public Item_Issue_Return_Amount As Decimal = 0
    Public Deduction_Amount As Decimal = 0
    Public Asset_Lost_Amount As Decimal = 0
    Public Handling_Charges_Amount As Decimal = 0
    Public SRN_Net_Amount As Decimal = 0
    Public SRN_RO_Amount As Decimal = 0
    Public Advance_Payment_Amount As Decimal = 0
    Public Advance_Payment_Amount_Knock_Off As Decimal = 0

    Public Credit_Note_Amount As Decimal = 0
    Public Compulsory_Amount As Decimal = 0
    Public Saving_Amount As Decimal = 0
    Public Payable_Amount As Decimal = 0
    Public VSP_Amount As Decimal = 0
    Public Cheque_Dated As Date? = Nothing
    Public Service_Charge_Amt As Decimal = 0
    Public MP_VSP_Diff_Amount As Decimal = 0
    Public is_Hold_Payment_Process As Boolean = False

    Public is_Hold_Payment_Process_Saving As Boolean = False
    Public is_Hold_Payment_Process_Saving_Auto As Boolean = False
    Public is_Hold_Payment_Process_Saving_Manual As Boolean = False

    Public MP_Amount As Decimal
    Public MP_EMP As Decimal
    Public MP_Incentive As Decimal
    Public MP_IncentiveEMP As Decimal
    Public MP_Net_Amount As Decimal
    Public FarmerPayment As Decimal = 0
    Public VSP_Excess_Amount As Decimal = 0
    Public MP_Total_Deduction As Decimal = 0
    Public NextCycleDebitNote As Decimal = 0
    Public FarmerMilkQty As Decimal = 0
    Public FarmerSaleAmount As Decimal = 0
    Public FarmerSaleReturnAmount As Decimal = 0
    Public FarmerAdjustmentAmount As Decimal = 0
    Public FarmerPayableAmount As Decimal = 0
    Public PrevCycleDebitNote As Decimal = 0
    Public PrevCycleDebitNoteMP As Decimal = 0
    Public NextCycleDebitNoteMP As Decimal = 0

    Public CalFATPer As Decimal
    Public CalFATKG As Decimal
    Public CalSNFPer As Decimal
    Public CalSNFKg As Decimal




#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessDetail), ByVal tran As SqlTransaction) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    arr.Item(i).PP_Detail_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select max(PP_Detail_No) from TSPL_PAYMENT_PROCESS_DETAIL", tran))
                    If clsCommon.myLen(arr.Item(i).PP_Detail_No) <= 0 Then
                        arr.Item(i).PP_Detail_No = "TR0000000000000001"
                    Else
                        arr.Item(i).PP_Detail_No = clsCommon.incval(arr.Item(i).PP_Detail_No)
                    End If
                    clsCommon.AddColumnsForChange(coll, "PP_Detail_No", arr.Item(i).PP_Detail_No)
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "Is_select", IIf(arr.Item(i).Is_select, 1, 0))
                    clsCommon.AddColumnsForChange(coll, "SNo", i + 1)
                    clsCommon.AddColumnsForChange(coll, "Milk_Purchase_Invoice_No", arr.Item(i).Milk_Purchase_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "Milk_Purchase_Invoice_Date", clsCommon.GetPrintDate(arr.Item(i).Milk_Purchase_Invoice_Date, "dd/MMM/yyyy hh:mm tt"))
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", clsCommon.GetPrintDate(arr.Item(i).AP_Invoice_Date, "dd/MMM/yyyy hh:mm tt"))
                    clsCommon.AddColumnsForChange(coll, "VLC_CODE_Uploader", arr.Item(i).VLC_CODE_Uploader)
                    clsCommon.AddColumnsForChange(coll, "VLC_Name", arr.Item(i).VLC_Name)
                    clsCommon.AddColumnsForChange(coll, "MCC_Code", arr.Item(i).MCC_Code, True)
                    clsCommon.AddColumnsForChange(coll, "VSP_CODE", arr.Item(i).VSP_CODE)
                    clsCommon.AddColumnsForChange(coll, "VSP_NAME", arr.Item(i).VSP_NAME)
                    clsCommon.AddColumnsForChange(coll, "Main_VSP_CODE", arr.Item(i).Main_VSP_CODE)
                    clsCommon.AddColumnsForChange(coll, "Main_VSP_NAME", arr.Item(i).Main_VSP_NAME)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Name", arr.Item(i).Payee_Joint_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Bank_Code", arr.Item(i).Payee_Joint_Bank_Code)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Bank_Name", arr.Item(i).Payee_Joint_Bank_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Branch_Code", arr.Item(i).Payee_Joint_Branch_Code)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Branch_Name", arr.Item(i).Payee_Joint_Branch_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Account_No", arr.Item(i).Payee_Joint_Account_No)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_IFSC_Code", arr.Item(i).Payee_Joint_IFSC_Code)
                    clsCommon.AddColumnsForChange(coll, "Bank_Code", clsCommon.myCstr(arr.Item(i).Bank_Code))
                    clsCommon.AddColumnsForChange(coll, "Bank_Desc", clsCommon.myCstr(arr.Item(i).Bank_Desc))
                    clsCommon.AddColumnsForChange(coll, "Payment_Mode", clsCommon.myCstr(arr.Item(i).Payment_Mode))
                    clsCommon.AddColumnsForChange(coll, "Cheque_No", clsCommon.myCstr(arr.Item(i).Cheque_No))
                    If Not arr.Item(i).Cheque_Dated Is Nothing Then
                        clsCommon.AddColumnsForChange(coll, "Cheque_Dated", clsCommon.GetPrintDate(clsCommon.myCstr(arr.Item(i).Cheque_Dated), "dd/MMM/yyyy"), True)
                    End If
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Name_Saving", clsCommon.myCstr(arr.Item(i).Payee_Joint_Name_Saving), True)
                    clsCommon.AddColumnsForChange(coll, "Bank_Code_Saving", clsCommon.myCstr(arr.Item(i).Bank_Code_Saving))
                    clsCommon.AddColumnsForChange(coll, "Bank_Desc_Saving", clsCommon.myCstr(arr.Item(i).Bank_Desc_Saving))
                    clsCommon.AddColumnsForChange(coll, "Bank_Account_No_Saving", clsCommon.myCstr(arr.Item(i).Bank_Account_No_Saving))
                    clsCommon.AddColumnsForChange(coll, "Payment_Mode_Saving", clsCommon.myCstr(arr.Item(i).Payment_Mode_Saving))
                    clsCommon.AddColumnsForChange(coll, "Zone_Code", clsCommon.myCstr(arr.Item(i).Zone_Code))

                    clsCommon.AddColumnsForChange(coll, "VSP_Amount", clsCommon.myCDecimal(arr.Item(i).VSP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Handling_Charges_Amount", clsCommon.myCDecimal(arr.Item(i).Handling_Charges_Amount))
                    clsCommon.AddColumnsForChange(coll, "SRN_Net_Amount", clsCommon.myCDecimal(arr.Item(i).SRN_Net_Amount))
                    clsCommon.AddColumnsForChange(coll, "SRN_RO_Amount", clsCommon.myCDecimal(arr.Item(i).SRN_RO_Amount))

                    clsCommon.AddColumnsForChange(coll, "MP_Amount", clsCommon.myCDecimal(arr.Item(i).MP_Amount))
                    clsCommon.AddColumnsForChange(coll, "MP_EMP", clsCommon.myCDecimal(arr.Item(i).MP_EMP))
                    clsCommon.AddColumnsForChange(coll, "MP_Incentive", clsCommon.myCDecimal(arr.Item(i).MP_Incentive))
                    clsCommon.AddColumnsForChange(coll, "MP_IncentiveEMP", clsCommon.myCDecimal(arr.Item(i).MP_IncentiveEMP))
                    clsCommon.AddColumnsForChange(coll, "MP_Net_Amount", clsCommon.myCDecimal(arr.Item(i).MP_Net_Amount))

                    clsCommon.AddColumnsForChange(coll, "MP_VSP_Diff_Amount", clsCommon.myCDecimal(arr.Item(i).MP_VSP_Diff_Amount))
                    clsCommon.AddColumnsForChange(coll, "Milk_Qty", clsCommon.myCDecimal(arr.Item(i).Milk_Qty))
                    clsCommon.AddColumnsForChange(coll, "Milk_Amount", clsCommon.myCDecimal(arr.Item(i).Milk_Amount))
                    clsCommon.AddColumnsForChange(coll, "Incentive_Amount", clsCommon.myCDecimal(arr.Item(i).Incentive_Amount))
                    clsCommon.AddColumnsForChange(coll, "EMP_Amount", clsCommon.myCDecimal(arr.Item(i).EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Incentive_EMP_Amount", clsCommon.myCDecimal(arr.Item(i).Incentive_EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Total_EMP_Amount", clsCommon.myCDecimal(arr.Item(i).Total_EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Total", clsCommon.myCDecimal(arr.Item(i).Total))
                    clsCommon.AddColumnsForChange(coll, "TDS_Amount", arr.Item(i).TDS_Amount, True)
                    clsCommon.AddColumnsForChange(coll, "Total_Invoice_Amount", clsCommon.myCDecimal(arr.Item(i).Total_Invoice_Amount))
                    clsCommon.AddColumnsForChange(coll, "Vsp_Own_System_Amount", clsCommon.myCDecimal(arr.Item(i).Vsp_Own_System_Amount))
                    clsCommon.AddColumnsForChange(coll, "Head_Load_Amount", clsCommon.myCDecimal(arr.Item(i).Head_Load_Amount))
                    clsCommon.AddColumnsForChange(coll, "Invoice_Deduction_Amount", clsCommon.myCstr(arr.Item(i).Invoice_Deduction_Amount))
                    clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCstr(arr.Item(i).Reduce_Deduc_Amt))
                    clsCommon.AddColumnsForChange(coll, "MCC_Sale_Amount", clsCommon.myCstr(arr.Item(i).MCC_Sale_Amount))
                    clsCommon.AddColumnsForChange(coll, "MCC_Sale_Return_Amount", clsCommon.myCDecimal(arr.Item(i).MCC_Sale_Return_Amount))
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Amount", clsCommon.myCDecimal(arr.Item(i).Item_Issue_Amount))
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Return_Amount", clsCommon.myCDecimal(arr.Item(i).Item_Issue_Return_Amount))
                    clsCommon.AddColumnsForChange(coll, "Deduction_Amount", clsCommon.myCDecimal(arr.Item(i).Deduction_Amount))
                    clsCommon.AddColumnsForChange(coll, "Asset_Lost_Amount", clsCommon.myCDecimal(arr.Item(i).Asset_Lost_Amount))
                    clsCommon.AddColumnsForChange(coll, "Credit_Note_Amount", clsCommon.myCDecimal(arr.Item(i).Credit_Note_Amount))
                    clsCommon.AddColumnsForChange(coll, "Compulsory_Amount", clsCommon.myCDecimal(arr.Item(i).Compulsory_Amount))
                    clsCommon.AddColumnsForChange(coll, "Saving_Amount", clsCommon.myCDecimal(arr.Item(i).Saving_Amount))
                    clsCommon.AddColumnsForChange(coll, "Payable_Amount", clsCommon.myCDecimal(arr.Item(i).Payable_Amount))
                    clsCommon.AddColumnsForChange(coll, "Service_Charge_Amt", clsCommon.myCDecimal(arr.Item(i).Service_Charge_Amt))
                    clsCommon.AddColumnsForChange(coll, "Advance_Payment_Amount", clsCommon.myCDecimal(arr.Item(i).Advance_Payment_Amount))
                    clsCommon.AddColumnsForChange(coll, "Advance_Payment_Amount_Knock_Off", clsCommon.myCDecimal(arr.Item(i).Advance_Payment_Amount_Knock_Off))
                    clsCommon.AddColumnsForChange(coll, "is_Hold_Payment_Process", IIf(arr.Item(i).is_Hold_Payment_Process, 1, 0))

                    clsCommon.AddColumnsForChange(coll, "is_Hold_Payment_Process_Saving", IIf(arr.Item(i).is_Hold_Payment_Process_Saving, 1, 0))
                    clsCommon.AddColumnsForChange(coll, "is_Hold_Payment_Process_Saving_Auto", IIf(arr.Item(i).is_Hold_Payment_Process_Saving_Auto, 1, 0))
                    clsCommon.AddColumnsForChange(coll, "is_Hold_Payment_Process_Saving_Manual", IIf(arr.Item(i).is_Hold_Payment_Process_Saving_Manual, 1, 0))

                    clsCommon.AddColumnsForChange(coll, "VSP_Excess_Amount", clsCommon.myCDecimal(arr.Item(i).VSP_Excess_Amount))
                    clsCommon.AddColumnsForChange(coll, "MP_Total_Deduction", clsCommon.myCDecimal(arr.Item(i).MP_Total_Deduction))
                    clsCommon.AddColumnsForChange(coll, "NextCycleDebitNote", clsCommon.myCDecimal(arr.Item(i).NextCycleDebitNote))
                    clsCommon.AddColumnsForChange(coll, "FarmerPayment", clsCommon.myCDecimal(arr.Item(i).FarmerPayment))

                    clsCommon.AddColumnsForChange(coll, "FarmerMilkQty", clsCommon.myCDecimal(arr.Item(i).FarmerMilkQty))
                    clsCommon.AddColumnsForChange(coll, "FarmerSaleAmount", clsCommon.myCDecimal(arr.Item(i).FarmerSaleAmount))
                    clsCommon.AddColumnsForChange(coll, "FarmerSaleReturnAmount", clsCommon.myCDecimal(arr.Item(i).FarmerSaleReturnAmount))
                    clsCommon.AddColumnsForChange(coll, "FarmerAdjustmentAmount", clsCommon.myCDecimal(arr.Item(i).FarmerAdjustmentAmount))
                    clsCommon.AddColumnsForChange(coll, "FarmerPayableAmount", clsCommon.myCDecimal(arr.Item(i).FarmerPayableAmount))
                    clsCommon.AddColumnsForChange(coll, "PrevCycleDebitNote", clsCommon.myCDecimal(arr.Item(i).PrevCycleDebitNote))

                    clsCommon.AddColumnsForChange(coll, "PrevCycleDebitNoteMP", clsCommon.myCDecimal(arr.Item(i).PrevCycleDebitNoteMP))
                    clsCommon.AddColumnsForChange(coll, "NextCycleDebitNoteMP", clsCommon.myCDecimal(arr.Item(i).NextCycleDebitNoteMP))

                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_DETAIL", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select cast(TSPL_PAYMENT_PROCESS_DETAIL.Is_select as bit) as Is_select 
,TSPL_PAYMENT_PROCESS_DETAIL.Doc_No  
,TSPL_PAYMENT_PROCESS_DETAIL.SNo  
,cast(TSPL_PAYMENT_PROCESS_DETAIL.is_Hold_Payment_Process as bit) as is_Hold_Payment_Process
,cast(TSPL_PAYMENT_PROCESS_DETAIL.is_Hold_Payment_Process_Saving as bit) as is_Hold_Payment_Process_Saving
,cast(TSPL_PAYMENT_PROCESS_DETAIL.is_Hold_Payment_Process_Saving_Auto as bit) as is_Hold_Payment_Process_Saving_Auto
,cast(TSPL_PAYMENT_PROCESS_DETAIL.is_Hold_Payment_Process_Saving_Manual as bit) as is_Hold_Payment_Process_Saving_Manual
,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No  
,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_Date  
,TSPL_PAYMENT_PROCESS_DETAIL.AP_Invoice_No 
,TSPL_PAYMENT_PROCESS_DETAIL.AP_Invoice_Date 
,TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader 
,TSPL_PAYMENT_PROCESS_DETAIL.VLC_Name 
,TSPL_PAYMENT_PROCESS_DETAIL.MCC_Code 
,TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE 
,TSPL_PAYMENT_PROCESS_DETAIL.VSP_NAME 
,TSPL_PAYMENT_PROCESS_DETAIL.Main_VSP_CODE 
,TSPL_PAYMENT_PROCESS_DETAIL.Main_VSP_NAME 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Bank_Code 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Bank_Name 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Code 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Name 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code 
,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code 
,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc
,TSPL_PAYMENT_PROCESS_DETAIL.Zone_Code
,TSPL_PAYMENT_PROCESS_DETAIL.Payment_Mode 
,TSPL_PAYMENT_PROCESS_DETAIL.Cheque_No 
,TSPL_PAYMENT_PROCESS_DETAIL.Cheque_Dated 
,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name_Saving 
,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code_Saving 
,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc_Saving 
,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving
,TSPL_PAYMENT_PROCESS_DETAIL.Payment_Mode_Saving 
,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Qty 
 ,TabFATSNFDetail.FATPer 
,TabFATSNFDetail.FATKg 
,TabFATSNFDetail.SNFPer 
,TabFATSNFDetail.SNFKg 
,TSPL_PAYMENT_PROCESS_DETAIL.VSP_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_EMP 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_Incentive 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_IncentiveEMP 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_Net_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.MP_VSP_Diff_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Milk_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Incentive_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.EMP_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Incentive_EMP_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Total_EMP_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Total 
,TSPL_PAYMENT_PROCESS_DETAIL.SRN_RO_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.SRN_Net_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Handling_Charges_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.TDS_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Total_Invoice_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Vsp_Own_System_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Head_Load_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Deduction_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Reduce_Deduc_Amt 
,TSPL_PAYMENT_PROCESS_DETAIL.Service_Charge_Amt 
,TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.MCC_Sale_Return_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Item_Issue_Return_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Deduction_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Asset_Lost_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Advance_Payment_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Advance_Payment_Amount_Knock_Off 
,TSPL_PAYMENT_PROCESS_DETAIL.Credit_Note_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount
,TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount 
,TSPL_PAYMENT_PROCESS_DETAIL.Invoice_Deduction_Amount 
,TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME
from   TSPL_PAYMENT_PROCESS_DETAIL 
left outer join (select DOC_CODE,cast( sum(FATKg) as decimal(18,3)) as FATKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(FATKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as FATPer ,cast( sum(SNFKg) as decimal(18,3)) as SNFKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(SNFKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as SNFPer  from (select DOC_CODE, ACC_Qty,FAT_PER,SNF_PER,cast(ACC_Qty*FAT_PER/100 as decimal(18,2)) as FATKg, cast(ACC_Qty * SNF_PER / 100 As Decimal(18,2)) As SNFKg from TSPL_MILK_PURCHASE_INVOICE_DETAIL )xx group by DOC_CODE  ) As TabFATSNFDetail On TabFATSNFDetail.DOC_CODE= TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No 
left outer join TSPL_MILK_PURCHASE_INVOICE_HEAD on TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE
where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No ='" & doc_No & "' order by TSPL_PAYMENT_PROCESS_DETAIL.SNo"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try


    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByRef ArrIndex As Dictionary(Of String, clsPaymentProcessIndex), ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessDetail)
        Try
            Dim arr As New List(Of clsPaymentProcessDetail)
            Dim obj As New clsPaymentProcessDetail
            Dim qry As String = "select TSPL_PAYMENT_PROCESS_DETAIL.*,TabFATSNFDetail.FATKg,TabFATSNFDetail.FATPer,TabFATSNFDetail.SNFKg,TabFATSNFDetail.SNFPer from TSPL_PAYMENT_PROCESS_DETAIL " +
                " left outer join (select DOC_CODE,cast( sum(FATKg) as decimal(18,3)) as FATKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(FATKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as FATPer ,cast( sum(SNFKg) as decimal(18,3)) as SNFKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(SNFKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as SNFPer " +
                " from (select DOC_CODE, ACC_Qty,FAT_PER,SNF_PER,cast(ACC_Qty*FAT_PER/100 as decimal(18,2)) as FATKg,cast( ACC_Qty*SNF_PER/100 as decimal(18,2)) as SNFKg from TSPL_MILK_PURCHASE_INVOICE_DETAIL )xx group by DOC_CODE " +
                ") as TabFATSNFDetail on TabFATSNFDetail.DOC_CODE=TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No" + Environment.NewLine +
                " where TSPL_PAYMENT_PROCESS_DETAIL.Doc_No='" & doc_No & "' "
            If ApplyIndex Then
                qry += "  order by TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE "
            Else
                qry += "  order by TSPL_PAYMENT_PROCESS_DETAIL.SNo "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessDetail
                    obj.PP_Detail_No = clsCommon.myCstr(dtbl.Rows(i)("PP_Detail_No"))
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.Is_select = IIf(clsCommon.myCDecimal(dtbl.Rows(i)("Is_select")) > 0, True, False)
                    obj.SNo = clsCommon.myCstr(dtbl.Rows(i)("SNo"))
                    obj.Milk_Purchase_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("Milk_Purchase_Invoice_No"))
                    obj.Milk_Purchase_Invoice_Date = clsCommon.myCDate(dtbl.Rows(i)("Milk_Purchase_Invoice_Date"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCDate(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.VLC_CODE_Uploader = clsCommon.myCstr(dtbl.Rows(i)("VLC_CODE_Uploader"))
                    obj.VLC_Name = clsCommon.myCstr(dtbl.Rows(i)("VLC_Name"))
                    obj.MCC_Code = clsCommon.myCstr(dtbl.Rows(i)("MCC_Code"))
                    obj.VSP_CODE = clsCommon.myCstr(dtbl.Rows(i)("VSP_CODE"))
                    obj.VSP_NAME = clsCommon.myCstr(dtbl.Rows(i)("VSP_NAME"))
                    obj.Main_VSP_CODE = clsCommon.myCstr(dtbl.Rows(i)("Main_VSP_CODE"))
                    obj.Main_VSP_NAME = clsCommon.myCstr(dtbl.Rows(i)("Main_VSP_NAME"))
                    obj.Payee_Joint_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Name"))
                    obj.Payee_Joint_Bank_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Bank_Code"))
                    obj.Payee_Joint_Bank_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Bank_Name"))
                    obj.Payee_Joint_Branch_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Branch_Code"))
                    obj.Payee_Joint_Branch_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Branch_Name"))
                    obj.Payee_Joint_Account_No = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Account_No"))
                    obj.Payee_Joint_IFSC_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_IFSC_Code"))
                    obj.Bank_Code = clsCommon.myCstr(dtbl.Rows(i)("Bank_Code"))
                    obj.Bank_Desc = clsCommon.myCstr(dtbl.Rows(i)("Bank_Desc"))
                    obj.Payment_Mode = clsCommon.myCstr(dtbl.Rows(i)("Payment_Mode"))
                    obj.Cheque_No = clsCommon.myCstr(dtbl.Rows(i)("Cheque_No"))
                    If clsCommon.myLen(dtbl.Rows(i)("Cheque_Dated")) > 0 Then
                        obj.Cheque_Dated = clsCommon.myCstr(dtbl.Rows(i)("Cheque_Dated"))
                    End If
                    obj.Payee_Joint_Name_Saving = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Name_Saving"))
                    obj.Bank_Code_Saving = clsCommon.myCstr(dtbl.Rows(i)("Bank_Code_Saving"))
                    obj.Bank_Desc_Saving = clsCommon.myCstr(dtbl.Rows(i)("Bank_Desc_Saving"))
                    obj.Payment_Mode_Saving = clsCommon.myCstr(dtbl.Rows(i)("Payment_Mode_Saving"))
                    obj.Bank_Account_No_Saving = clsCommon.myCstr(dtbl.Rows(i)("Bank_Account_No_Saving"))

                    obj.VSP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("VSP_Amount"))
                    obj.Handling_Charges_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Handling_Charges_Amount"))

                    obj.SRN_Net_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("SRN_Net_Amount"))
                    obj.SRN_RO_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("SRN_RO_Amount"))

                    obj.MP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Amount"))
                    obj.MP_EMP = clsCommon.myCDecimal(dtbl.Rows(i)("MP_EMP"))
                    obj.MP_Incentive = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Incentive"))
                    obj.MP_IncentiveEMP = clsCommon.myCDecimal(dtbl.Rows(i)("MP_IncentiveEMP"))
                    obj.MP_Net_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Net_Amount"))

                    obj.MP_VSP_Diff_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MP_VSP_Diff_Amount"))
                    obj.Milk_Qty = clsCommon.myCDecimal(dtbl.Rows(i)("Milk_Qty"))
                    obj.Milk_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Milk_Amount"))
                    obj.Incentive_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Incentive_Amount"))
                    obj.EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("EMP_Amount"))
                    obj.Incentive_EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Incentive_EMP_Amount"))
                    obj.Total_EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Total_EMP_Amount"))
                    obj.Total = clsCommon.myCDecimal(dtbl.Rows(i)("Total"))
                    obj.TDS_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("TDS_Amount"))
                    obj.Total_Invoice_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Total_Invoice_Amount"))
                    obj.Vsp_Own_System_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Vsp_Own_System_Amount"))
                    obj.Head_Load_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Head_Load_Amount"))
                    obj.Invoice_Deduction_Amount = clsCommon.myCstr(dtbl.Rows(i)("Invoice_Deduction_Amount"))
                    obj.Reduce_Deduc_Amt = clsCommon.myCstr(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    obj.MCC_Sale_Amount = clsCommon.myCstr(dtbl.Rows(i)("MCC_Sale_Amount"))
                    obj.MCC_Sale_Return_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MCC_Sale_Return_Amount"))
                    obj.Item_Issue_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Item_Issue_Amount"))
                    obj.Item_Issue_Return_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Item_Issue_Return_Amount"))
                    obj.Deduction_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Deduction_Amount"))
                    obj.Asset_Lost_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Asset_Lost_Amount"))
                    obj.Credit_Note_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Credit_Note_Amount"))
                    obj.Saving_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Saving_Amount"))
                    obj.Compulsory_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Compulsory_Amount"))
                    obj.Payable_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Payable_Amount"))
                    obj.Service_Charge_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Service_Charge_Amt"))
                    obj.Advance_Payment_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Advance_Payment_Amount"))
                    obj.Advance_Payment_Amount_Knock_Off = clsCommon.myCDecimal(dtbl.Rows(i)("Advance_Payment_Amount_Knock_Off"))
                    obj.is_Hold_Payment_Process = IIf(clsCommon.myCDecimal(dtbl.Rows(i)("is_Hold_Payment_Process")) = 1, True, False)

                    obj.is_Hold_Payment_Process_Saving = IIf(clsCommon.myCDecimal(dtbl.Rows(i)("is_Hold_Payment_Process_Saving")) = 1, True, False)
                    obj.is_Hold_Payment_Process_Saving_Auto = IIf(clsCommon.myCDecimal(dtbl.Rows(i)("is_Hold_Payment_Process_Saving_Auto")) = 1, True, False)
                    obj.is_Hold_Payment_Process_Saving_Manual = IIf(clsCommon.myCDecimal(dtbl.Rows(i)("is_Hold_Payment_Process_Saving_Manual")) = 1, True, False)
                    obj.VSP_Excess_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("VSP_Excess_Amount"))
                    obj.MP_Total_Deduction = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Total_Deduction"))
                    obj.NextCycleDebitNote = clsCommon.myCDecimal(dtbl.Rows(i)("NextCycleDebitNote"))
                    obj.FarmerPayment = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerPayment"))

                    obj.FarmerMilkQty = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerMilkQty"))
                    obj.FarmerSaleAmount = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerSaleAmount"))
                    obj.FarmerSaleReturnAmount = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerSaleReturnAmount"))
                    obj.FarmerAdjustmentAmount = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerAdjustmentAmount"))
                    obj.FarmerPayableAmount = clsCommon.myCDecimal(dtbl.Rows(i)("FarmerPayableAmount"))
                    obj.PrevCycleDebitNote = clsCommon.myCDecimal(dtbl.Rows(i)("PrevCycleDebitNote"))

                    obj.PrevCycleDebitNoteMP = clsCommon.myCDecimal(dtbl.Rows(i)("PrevCycleDebitNoteMP"))
                    obj.NextCycleDebitNoteMP = clsCommon.myCDecimal(dtbl.Rows(i)("NextCycleDebitNoteMP"))

                    obj.CalFATKG = clsCommon.myCDecimal(dtbl.Rows(i)("FATKg"))
                    obj.CalFATPer = clsCommon.myCDecimal(dtbl.Rows(i)("FATPer"))
                    obj.CalSNFKg = clsCommon.myCDecimal(dtbl.Rows(i)("SNFKg"))
                    obj.CalSNFPer = clsCommon.myCDecimal(dtbl.Rows(i)("SNFPer"))
                    arr.Add(obj)
                    If ApplyIndex Then
                        Dim objidx As New clsPaymentProcessIndex
                        objidx.COMPULSORYTo = -1
                        objidx.CRTo = -1
                        objidx.DRTo = -1
                        objidx.SavingTo = -1
                        ArrIndex.Add(clsCommon.myCstr(dtbl.Rows(i)("VSP_CODE")).ToUpper(), objidx)
                    End If
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_DETAIL where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessInvoices
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public Milk_Purchase_Invoice_No As String = ""
    Public Milk_Purchase_Invoice_Date As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As String = ""
    Public VLC_CODE As String = ""
    Public VSP_CODE As String = ""
    Public VSP_NAME As String = ""
    Public Payee_Joint_Name As String = ""
    Public Payee_Joint_Ac_No As String = ""
    Public Payee_Joint_Bank_Code As String = ""
    Public Payee_Joint_Bank_Name As String = ""
    Public Payee_Joint_Branch_Code As String = ""
    Public Payee_Joint_Branch_Name As String = ""
    Public Payee_Joint_IFSC_Code As String = ""
    Public Milk_Qty As Decimal = 0
    Public Inv_Amount As Decimal = 0
    Public Inv_EMP_Amount As Decimal = 0
    Public Handling_Charges_Amount As Decimal = 0
    Public SRN_Net_Amount As Decimal = 0
    Public SRN_RO_Amount As Decimal = 0
    Public Inv_Amt_EMP_Amount As Decimal = 0
    Public Inv_Incentive_Amount As Decimal = 0
    Public Inv_Incentive_EMP_Amount As Decimal = 0
    Public Gross_Amount As Decimal = 0
    Public TDS_Amount As Decimal = 0
    Public Vsp_Own_System_Amount As Decimal = 0
    Public Head_Load_Amount As Decimal = 0
    Public Deduction_Amount As Decimal = 0
    Public Vsp_Own_System_Doc_No As String = ""
    Public Head_Load_Doc_No As String = ""
    Public Deduction_Doc_No As String = ""
    Public Reduce_Deduc_Amt As Decimal = 0
    Public Bank_Code As String = ""
    Public Bank_Desc As String = ""
    Public Payment_Mode As String = ""
    Public Cheque_No As String = ""
    Public Service_Charge_Amt As Decimal = 0
    Public ActualVSPCode As String = ""
    Public ActualVSPName As String = ""

    Public MP_Amount As Decimal
    Public MP_EMP As Decimal
    Public MP_Incentive As Decimal
    Public MP_IncentiveEMP As Decimal
    Public MP_Net_Amount As Decimal

    Public CalFATPer As Decimal
    Public CalFATKG As Decimal
    Public CalSNFPer As Decimal
    Public CalSNFKg As Decimal
    Public MCC_Code As String = ""
    Public Area_Location_Code As String = ""

#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessInvoices), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then

                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    'clsCommon.AddColumnsForChange(coll, "Area_Location_Code", arr.Item(i).Area_Location_Code)
                    clsCommon.AddColumnsForChange(coll, "Milk_Purchase_Invoice_No", arr.Item(i).Milk_Purchase_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "Milk_Purchase_Invoice_Date", arr.Item(i).Milk_Purchase_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", arr.Item(i).AP_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "VLC_CODE", arr.Item(i).VLC_CODE)
                    clsCommon.AddColumnsForChange(coll, "MCC_Code", arr.Item(i).MCC_Code, True)
                    clsCommon.AddColumnsForChange(coll, "VSP_CODE", arr.Item(i).VSP_CODE)
                    clsCommon.AddColumnsForChange(coll, "VSP_NAME", arr.Item(i).VSP_NAME)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Name", arr.Item(i).Payee_Joint_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Bank_Code", arr.Item(i).Payee_Joint_Bank_Code)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Bank_Name", arr.Item(i).Payee_Joint_Bank_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Branch_Code", arr.Item(i).Payee_Joint_Branch_Code)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Branch_Name", arr.Item(i).Payee_Joint_Branch_Name)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_Ac_No", arr.Item(i).Payee_Joint_Ac_No)
                    clsCommon.AddColumnsForChange(coll, "Payee_Joint_IFSC_Code", arr.Item(i).Payee_Joint_IFSC_Code)
                    clsCommon.AddColumnsForChange(coll, "Milk_Qty", clsCommon.myCDecimal(arr.Item(i).Milk_Qty))
                    clsCommon.AddColumnsForChange(coll, "Inv_Amount", clsCommon.myCDecimal(arr.Item(i).Inv_Amount))
                    clsCommon.AddColumnsForChange(coll, "Inv_EMP_Amount", clsCommon.myCDecimal(arr.Item(i).Inv_EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Handling_Charges_Amount", clsCommon.myCDecimal(arr.Item(i).Handling_Charges_Amount))
                    clsCommon.AddColumnsForChange(coll, "SRN_Net_Amount", clsCommon.myCDecimal(arr.Item(i).SRN_Net_Amount))
                    clsCommon.AddColumnsForChange(coll, "SRN_RO_Amount", clsCommon.myCDecimal(arr.Item(i).SRN_RO_Amount))
                    clsCommon.AddColumnsForChange(coll, "Inv_Amt_EMP_Amount", clsCommon.myCDecimal(arr.Item(i).Inv_Amt_EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Inv_Incentive_Amount", clsCommon.myCDecimal(arr.Item(i).Inv_Incentive_Amount))
                    clsCommon.AddColumnsForChange(coll, "Inv_Incentive_EMP_Amount", clsCommon.myCDecimal(arr.Item(i).Inv_Incentive_EMP_Amount))
                    clsCommon.AddColumnsForChange(coll, "Gross_Amount", clsCommon.myCDecimal(arr.Item(i).Gross_Amount))
                    clsCommon.AddColumnsForChange(coll, "TDS_Amount", arr.Item(i).TDS_Amount, True)
                    clsCommon.AddColumnsForChange(coll, "Vsp_Own_System_Amount", clsCommon.myCDecimal(arr.Item(i).Vsp_Own_System_Amount))
                    clsCommon.AddColumnsForChange(coll, "Head_Load_Amount", clsCommon.myCDecimal(arr.Item(i).Head_Load_Amount))
                    clsCommon.AddColumnsForChange(coll, "Deduction_Amount", clsCommon.myCDecimal(arr.Item(i).Deduction_Amount))
                    clsCommon.AddColumnsForChange(coll, "Vsp_Own_System_Doc_No", clsCommon.myCstr(arr.Item(i).Vsp_Own_System_Doc_No))
                    clsCommon.AddColumnsForChange(coll, "Head_Load_Doc_No", clsCommon.myCstr(arr.Item(i).Head_Load_Doc_No))
                    clsCommon.AddColumnsForChange(coll, "Deduction_Doc_No", clsCommon.myCstr(arr.Item(i).Deduction_Doc_No))
                    clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCDecimal(arr.Item(i).Reduce_Deduc_Amt))
                    clsCommon.AddColumnsForChange(coll, "Bank_Code", clsCommon.myCstr(arr.Item(i).Bank_Code))
                    clsCommon.AddColumnsForChange(coll, "Bank_Desc", clsCommon.myCstr(arr.Item(i).Bank_Desc))
                    clsCommon.AddColumnsForChange(coll, "Payment_Mode", clsCommon.myCstr(arr.Item(i).Payment_Mode))
                    clsCommon.AddColumnsForChange(coll, "Cheque_No", clsCommon.myCstr(arr.Item(i).Cheque_No))
                    clsCommon.AddColumnsForChange(coll, "Service_Charge_Amt", clsCommon.myCstr(arr.Item(i).Service_Charge_Amt))

                    clsCommon.AddColumnsForChange(coll, "MP_Amount", clsCommon.myCstr(arr.Item(i).MP_Amount))
                    clsCommon.AddColumnsForChange(coll, "MP_EMP", clsCommon.myCstr(arr.Item(i).MP_EMP))
                    clsCommon.AddColumnsForChange(coll, "MP_Incentive", clsCommon.myCstr(arr.Item(i).MP_Incentive))
                    clsCommon.AddColumnsForChange(coll, "MP_IncentiveEMP", clsCommon.myCstr(arr.Item(i).MP_IncentiveEMP))
                    clsCommon.AddColumnsForChange(coll, "MP_Net_Amount", clsCommon.myCstr(arr.Item(i).MP_Net_Amount))

                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_INVOICE", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select cast(1 as bit) as Sel 
,TSPL_PAYMENT_PROCESS_INVOICE.*  
,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader 
,TSPL_VLC_MASTER_HEAD.VLC_Name 
 ,TabFATSNFDetail.FATPer  
,TabFATSNFDetail.FATKg  
,TabFATSNFDetail.SNFPer  
,TabFATSNFDetail.SNFKg 
,TSPL_VENDOR_MASTER.Parent_Vendor_Code as ActualVSPCode 
,ParVen.Vendor_Name as ActualVSPName,TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,TSPL_MILK_PURCHASE_INVOICE_HEAD.FILE_INFO
from TSPL_PAYMENT_PROCESS_INVOICE 
left outer join TSPL_MILK_PURCHASE_INVOICE_HEAD on TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE=TSPL_PAYMENT_PROCESS_INVOICE.Milk_Purchase_Invoice_No
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_PURCHASE_INVOICE_HEAD.ROUTE_CODE
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_INVOICE.VSP_CODE 
left outer join TSPL_VENDOR_MASTER as ParVen on ParVen.Vendor_Code=TSPL_VENDOR_MASTER.Parent_Vendor_Code  
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_INVOICE.VSP_CODE
left outer join (select DOC_CODE,cast(sum(FATKg) as decimal(18,3)) as FATKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(FATKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as FATPer ,cast( sum(SNFKg) as decimal(18,3)) as SNFKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(SNFKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as SNFPer from(select DOC_CODE, ACC_Qty, FAT_PER, SNF_PER, ACC_Qty * FAT_PER / 100 As FATKg, ACC_Qty * SNF_PER / 100 As SNFKg from TSPL_MILK_PURCHASE_INVOICE_DETAIL )xx group by DOC_CODE  ) as TabFATSNFDetail on TabFATSNFDetail.DOC_CODE=TSPL_PAYMENT_PROCESS_INVOICE.Milk_Purchase_Invoice_No 
where TSPL_PAYMENT_PROCESS_INVOICE.Doc_No ='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_INVOICE.slno as integer)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessInvoices)
        Try
            Dim arr As New List(Of clsPaymentProcessInvoices)
            Dim obj As New clsPaymentProcessInvoices
            Dim q As String = "select TSPL_PAYMENT_PROCESS_INVOICE.*,TSPL_VENDOR_MASTER.Parent_Vendor_Code as ActualVSPCode,ParVen.Vendor_Name as ActualVSPName,TabFATSNFDetail.FATKg,TabFATSNFDetail.FATPer,TabFATSNFDetail.SNFKg,TabFATSNFDetail.SNFPer from TSPL_PAYMENT_PROCESS_INVOICE left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_INVOICE.VSP_CODE left outer join TSPL_VENDOR_MASTER as ParVen on ParVen.Vendor_Code=TSPL_VENDOR_MASTER.Parent_Vendor_Code" +
            " left outer join (select DOC_CODE,cast( sum(FATKg) as decimal(18,3)) as FATKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(FATKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as FATPer ,cast( sum(SNFKg) as decimal(18,3)) as SNFKg,cast(case when sum(ACC_Qty)=0 then 0 else sum(SNFKg)*100/sum(ACC_Qty) end as decimal(18,2) ) as SNFPer " +
            " from (select DOC_CODE, ACC_Qty,FAT_PER,SNF_PER,ACC_Qty*FAT_PER/100 as FATKg,ACC_Qty*SNF_PER/100 as SNFKg from TSPL_MILK_PURCHASE_INVOICE_DETAIL )xx group by DOC_CODE" +
            " ) as TabFATSNFDetail on TabFATSNFDetail.DOC_CODE=TSPL_PAYMENT_PROCESS_INVOICE.Milk_Purchase_Invoice_No" +
            " where TSPL_PAYMENT_PROCESS_INVOICE.Doc_No='" & doc_No & "'"
            If ApplyIndex Then
                q += " order by TSPL_PAYMENT_PROCESS_INVOICE.VSP_CODE "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessInvoices
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.Milk_Purchase_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("Milk_Purchase_Invoice_No"))
                    obj.Milk_Purchase_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("Milk_Purchase_Invoice_Date"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.VLC_CODE = clsCommon.myCstr(dtbl.Rows(i)("VLC_CODE"))
                    obj.MCC_Code = clsCommon.myCstr(dtbl.Rows(i)("MCC_Code"))
                    obj.VSP_CODE = clsCommon.myCstr(dtbl.Rows(i)("VSP_CODE"))
                    obj.VSP_NAME = clsCommon.myCstr(dtbl.Rows(i)("VSP_NAME"))
                    obj.Payee_Joint_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Name"))
                    obj.Payee_Joint_Bank_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Bank_Code"))
                    obj.Payee_Joint_Bank_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Bank_Name"))
                    obj.Payee_Joint_Branch_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Branch_Code"))
                    obj.Payee_Joint_Branch_Name = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Branch_Name"))
                    obj.Payee_Joint_IFSC_Code = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_IFSC_Code"))
                    obj.Payee_Joint_Ac_No = clsCommon.myCstr(dtbl.Rows(i)("Payee_Joint_Ac_No"))
                    obj.Milk_Qty = clsCommon.myCDecimal(dtbl.Rows(i)("Milk_Qty"))
                    obj.Inv_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Inv_Amount"))
                    obj.Inv_EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Inv_EMP_Amount"))
                    obj.Handling_Charges_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Handling_Charges_Amount"))
                    obj.SRN_Net_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("SRN_Net_Amount"))
                    obj.SRN_RO_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("SRN_RO_Amount"))
                    obj.Inv_Amt_EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Inv_Amt_EMP_Amount"))
                    obj.Inv_Incentive_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Inv_Incentive_Amount"))
                    obj.Inv_Incentive_EMP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Inv_Incentive_EMP_Amount"))
                    obj.Gross_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Gross_Amount"))
                    obj.TDS_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("TDS_Amount"))
                    obj.Vsp_Own_System_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Vsp_Own_System_Amount"))
                    obj.Head_Load_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Head_Load_Amount"))
                    obj.Deduction_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Deduction_Amount"))
                    obj.Vsp_Own_System_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Vsp_Own_System_Doc_No"))
                    obj.Head_Load_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Head_Load_Doc_No"))
                    obj.Deduction_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Deduction_Doc_No"))
                    obj.Reduce_Deduc_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    obj.Bank_Code = clsCommon.myCstr(dtbl.Rows(i)("Bank_Code"))
                    obj.Bank_Desc = clsCommon.myCstr(dtbl.Rows(i)("Bank_Desc"))
                    obj.Payment_Mode = clsCommon.myCstr(dtbl.Rows(i)("Payment_Mode"))
                    obj.Cheque_No = clsCommon.myCstr(dtbl.Rows(i)("Cheque_No"))
                    obj.Service_Charge_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Service_Charge_Amt"))
                    obj.ActualVSPCode = clsCommon.myCstr(dtbl.Rows(i)("ActualVSPCode"))
                    obj.ActualVSPName = clsCommon.myCstr(dtbl.Rows(i)("ActualVSPName"))

                    obj.MP_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Amount"))
                    obj.MP_EMP = clsCommon.myCDecimal(dtbl.Rows(i)("MP_EMP"))
                    obj.MP_Incentive = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Incentive"))
                    obj.MP_IncentiveEMP = clsCommon.myCDecimal(dtbl.Rows(i)("MP_IncentiveEMP"))
                    obj.MP_Net_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("MP_Net_Amount"))

                    obj.CalFATKG = clsCommon.myCDecimal(dtbl.Rows(i)("FATKg"))
                    obj.CalFATPer = clsCommon.myCDecimal(dtbl.Rows(i)("FATPer"))
                    obj.CalSNFKg = clsCommon.myCDecimal(dtbl.Rows(i)("SNFKg"))
                    obj.CalSNFPer = clsCommon.myCDecimal(dtbl.Rows(i)("SNFPer"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_INVOICE where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessMCCSale
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public Shipment_Doc_No As String = ""
    Public Shipment_Doc_Date As String = ""
    Public Sale_Doc_No As String = ""
    Public Sale_Doc_Date As String = ""
    Public AR_Invoice_No As String = ""
    Public AR_Invoice_Date As String = ""
    Public Customer_CODE As String = ""
    Public Customer_NAME As String = ""
    Public Item_Code As String = ""
    Public Item_Desc As String = ""
    Public Amount As Decimal = 0
    Public Reduce_Deduc_Amt As Decimal = 0
    Public Instalment_Amt As Decimal = 0
    Public Original_Balance_Amount As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessMCCSale), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                'Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_MCC_SALE where  Doc_No='" & arr.Item(0).Doc_No & "'"
                'clsDBFuncationality.ExecuteNonQuery(qry, tran)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "Shipment_Doc_No", arr.Item(i).Shipment_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Shipment_Doc_Date", arr.Item(i).Shipment_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "Sale_Doc_No", arr.Item(i).Sale_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Sale_Doc_Date", arr.Item(i).Sale_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "AR_Invoice_No", arr.Item(i).AR_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AR_Invoice_Date", arr.Item(i).AR_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Customer_CODE", arr.Item(i).Customer_CODE)
                    clsCommon.AddColumnsForChange(coll, "Customer_NAME", arr.Item(i).Customer_NAME)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", arr.Item(i).Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Item_Desc", arr.Item(i).Item_Desc)
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(arr.Item(i).Amount))
                    clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCDecimal(arr.Item(i).Reduce_Deduc_Amt))
                    clsCommon.AddColumnsForChange(coll, "Instalment_Amount", clsCommon.myCDecimal(arr.Item(i).Instalment_Amt))
                    clsCommon.AddColumnsForChange(coll, "Original_Balance_Amount", clsCommon.myCDecimal(arr.Item(i).Original_Balance_Amount))
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_MCC_SALE", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select  cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_MCC_SALE.*,TSPL_DEDUCTION_MASTER.Code as DeductionType,TSPL_DEDUCTION_MASTER.Description as DeductionDesc 
from TSPL_PAYMENT_PROCESS_MCC_SALE 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_MCC_SALE.Customer_CODE
left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_PAYMENT_PROCESS_MCC_SALE.Sale_Doc_No    
left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Code=TSPL_SD_SALE_INVOICE_HEAD.Deduction 
where TSPL_PAYMENT_PROCESS_MCC_SALE.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_MCC_SALE.SLNO as int)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessMCCSale)
        Try
            Dim arr As New List(Of clsPaymentProcessMCCSale)
            Dim obj As New clsPaymentProcessMCCSale
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_MCC_SALE where Doc_No='" & doc_No & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessMCCSale
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.Shipment_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Shipment_Doc_No"))
                    obj.Shipment_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Shipment_Doc_Date"))
                    obj.Sale_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Sale_Doc_No"))
                    obj.Sale_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Sale_Doc_Date"))
                    obj.AR_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AR_Invoice_No"))
                    obj.AR_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AR_Invoice_Date"))
                    obj.Customer_CODE = clsCommon.myCstr(dtbl.Rows(i)("Customer_CODE"))
                    obj.Customer_NAME = clsCommon.myCstr(dtbl.Rows(i)("Customer_NAME"))
                    obj.Item_Code = clsCommon.myCstr(dtbl.Rows(i)("Item_Code"))
                    obj.Item_Desc = clsCommon.myCstr(dtbl.Rows(i)("Item_Desc"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    obj.Reduce_Deduc_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    obj.Instalment_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Instalment_Amount"))
                    obj.Original_Balance_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Original_Balance_Amount"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_MCC_SALE where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessItemIssue
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public Item_Issue_Doc_No As String = ""
    Public Item_Issue_Doc_Date As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As String = ""
    Public Vendor_CODE As String = ""
    Public Vendor_NAME As String = ""
    Public Item_Code As String = ""
    Public Item_Desc As String = ""
    Public Amount As Decimal = 0
    Public Reduce_Deduc_Amt As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessItemIssue), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                'Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_ITEM_ISSUE where  Doc_No='" & arr.Item(0).Doc_No & "'"
                'clsDBFuncationality.ExecuteNonQuery(qry, tran)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Doc_No", arr.Item(i).Item_Issue_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Doc_Date", arr.Item(i).Item_Issue_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", arr.Item(i).AP_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Vendor_CODE", arr.Item(i).Vendor_CODE)
                    clsCommon.AddColumnsForChange(coll, "Vendor_NAME", arr.Item(i).Vendor_NAME)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", arr.Item(i).Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Item_Desc", arr.Item(i).Item_Desc)
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(arr.Item(i).Amount))
                    clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCDecimal(arr.Item(i).Reduce_Deduc_Amt))
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_ITEM_ISSUE", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_ITEM_ISSUE.* 
from TSPL_PAYMENT_PROCESS_ITEM_ISSUE 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_ITEM_ISSUE.Vendor_CODE
where Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_ITEM_ISSUE.SLNO as int) "
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessItemIssue)
        Try
            Dim arr As New List(Of clsPaymentProcessItemIssue)
            Dim obj As New clsPaymentProcessItemIssue
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_ITEM_ISSUE where Doc_No='" & doc_No & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessItemIssue
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.Item_Issue_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Item_Issue_Doc_No"))
                    obj.Item_Issue_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Item_Issue_Doc_Date"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_NAME"))
                    obj.Item_Code = clsCommon.myCstr(dtbl.Rows(i)("Item_Code"))
                    obj.Item_Desc = clsCommon.myCstr(dtbl.Rows(i)("Item_Desc"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    obj.Reduce_Deduc_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_ITEM_ISSUE where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class clsPaymentProcessItemIssueReturn
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public Item_Issue_Return_No As String = ""
    Public Item_Issue_Doc_No As String = ""
    Public Item_Issue_Return_Date As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As String = ""
    Public Vendor_CODE As String = ""
    Public Vendor_NAME As String = ""
    Public Item_Code As String = ""
    Public Item_Desc As String = ""
    Public Amount As Decimal = 0
    'Public Reduce_Deduc_Amt As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessItemIssueReturn), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Return_No", arr.Item(i).Item_Issue_Return_No)
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Doc_No", arr.Item(i).Item_Issue_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Item_Issue_Return_Date", arr.Item(i).Item_Issue_Return_Date)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", arr.Item(i).AP_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Vendor_CODE", arr.Item(i).Vendor_CODE)
                    clsCommon.AddColumnsForChange(coll, "Vendor_NAME", arr.Item(i).Vendor_NAME)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", arr.Item(i).Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Item_Desc", arr.Item(i).Item_Desc)
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(arr.Item(i).Amount))
                    'clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCDecimal(arr.Item(i).Reduce_Deduc_Amt))
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select  cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.* 
from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.Vendor_CODE
where TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN.SLNO as int)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessItemIssueReturn)
        Try
            Dim arr As New List(Of clsPaymentProcessItemIssueReturn)
            Dim obj As New clsPaymentProcessItemIssueReturn
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN where Doc_No='" & doc_No & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessItemIssueReturn
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.Item_Issue_Return_No = clsCommon.myCstr(dtbl.Rows(i)("Item_Issue_Return_No"))
                    obj.Item_Issue_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Item_Issue_Doc_No"))
                    obj.Item_Issue_Return_Date = clsCommon.myCstr(dtbl.Rows(i)("Item_Issue_Return_Date"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_NAME"))
                    obj.Item_Code = clsCommon.myCstr(dtbl.Rows(i)("Item_Code"))
                    obj.Item_Desc = clsCommon.myCstr(dtbl.Rows(i)("Item_Desc"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    'obj.Reduce_Deduc_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class clsPaymentProcessDeduction
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As String = ""
    Public Vendor_CODE As String = ""
    Public Vendor_NAME As String = ""
    Public Ded_Code As String = ""
    Public Ded_Desc As String = ""
    Public Amount As Decimal = 0
    Public Reduce_Deduc_Amt As Decimal = 0
    Public Area_Location_Code As String = ""
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessDeduction), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                'Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_DEDUCTION where  Doc_No='" & arr.Item(0).Doc_No & "'"
                'clsDBFuncationality.ExecuteNonQuery(qry, tran)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", i + 1)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", arr.Item(i).AP_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Vendor_CODE", arr.Item(i).Vendor_CODE)
                    clsCommon.AddColumnsForChange(coll, "Vendor_NAME", arr.Item(i).Vendor_NAME)
                    clsCommon.AddColumnsForChange(coll, "Ded_Code", arr.Item(i).Ded_Code)
                    clsCommon.AddColumnsForChange(coll, "Ded_Desc", arr.Item(i).Ded_Desc)
                    'clsCommon.AddColumnsForChange(coll, "Area_Location_Code", clsCommon.myCstr(arr.Item(i).Area_Location_Code))
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(arr.Item(i).Amount))
                    clsCommon.AddColumnsForChange(coll, "Reduce_Deduc_Amt", clsCommon.myCDecimal(arr.Item(i).Reduce_Deduc_Amt))
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_DEDUCTION", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            'Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("SELECT Area_Location_Code FROM TSPL_PAYMENT_PROCESS_HEAD")
            'Dim AreaLocationCode As String = ""

            'If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
            '    ' Assuming you want the first value in the DataTable
            '    AreaLocationCode = clsCommon.myCstr(dt1.Rows(0)("Area_Location_Code"))
            'End If
            'Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select  Area_Location_Code from tspl_mcc_master")
            'Dim lst As New List(Of String)
            'If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then

            '    For Each dr As DataRow In dt1.Rows
            '        lst.Add(clsCommon.myCstr(dr("Area_Location_Code")))
            '    Next
            'End If
            'Dim AreaLocationCode As String = clsCommon.GetMulcallString(lst)

            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_DEDUCTION.* 
from TSPL_PAYMENT_PROCESS_DEDUCTION 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE
left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_no=TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_no

where TSPL_PAYMENT_PROCESS_DEDUCTION.Doc_No='" & doc_No & "'order by cast(TSPL_PAYMENT_PROCESS_DEDUCTION.SLNO as int)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByRef ArrIndex As Dictionary(Of String, clsPaymentProcessIndex), ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessDeduction)
        Try
            Dim arr As New List(Of clsPaymentProcessDeduction)
            Dim obj As New clsPaymentProcessDeduction
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_DEDUCTION where Doc_No='" & doc_No & "'"
            If ApplyIndex Then
                q += " order by TSPL_PAYMENT_PROCESS_DEDUCTION.Vendor_CODE "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                Dim IdxVendor As String = clsCommon.myCstr(dtbl.Rows(0)("Vendor_CODE")).ToUpper()
                Dim IdxFrom As Integer = 0
                Dim IdxTo As Integer = -1
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    Try
                        If ApplyIndex Then
                            If (Not (clsCommon.CompairString(IdxVendor, clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))) = CompairStringResult.Equal)) Then
                                IdxTo = i - 1
                                ArrIndex(IdxVendor).DRFrom = IdxFrom
                                ArrIndex(IdxVendor).DRTo = IdxTo
                                IdxVendor = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE")).ToUpper()
                                IdxFrom = i
                            End If
                            If i = dtbl.Rows.Count - 1 Then
                                IdxTo = i
                                ArrIndex(IdxVendor).DRFrom = IdxFrom
                                ArrIndex(IdxVendor).DRTo = IdxTo
                            End If
                        End If
                    Catch ex As Exception
                        Dim x As Integer = i
                        Dim y As Integer = i
                    End Try


                    obj = New clsPaymentProcessDeduction
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_NAME"))
                    obj.Ded_Code = clsCommon.myCstr(dtbl.Rows(i)("Ded_Code"))
                    obj.Ded_Desc = clsCommon.myCstr(dtbl.Rows(i)("Ded_Desc"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    'obj.Area_Location_Code = clsCommon.myCstr(dtbl.Rows(i)("Area_Location_Code"))
                    obj.Reduce_Deduc_Amt = clsCommon.myCDecimal(dtbl.Rows(i)("Reduce_Deduc_Amt"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_DEDUCTION where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessCreditNote
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Date As String = ""
    Public Vendor_CODE As String = ""
    Public Vendor_NAME As String = ""
    Public TDS_Amount As Decimal = 0
    Public Amount As Decimal = 0
    Public Area_Location_Code As String = ""
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessCreditNote), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_Date", arr.Item(i).AP_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Vendor_CODE", arr.Item(i).Vendor_CODE)
                    clsCommon.AddColumnsForChange(coll, "Vendor_NAME", arr.Item(i).Vendor_NAME)
                    clsCommon.AddColumnsForChange(coll, "TDS_Amount", arr.Item(i).TDS_Amount)
                    clsCommon.AddColumnsForChange(coll, "Amount", arr.Item(i).Amount)
                    'clsCommon.AddColumnsForChange(coll, "Area_Location_Code", clsCommon.myCstr(arr.Item(i).Area_Location_Code))

                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_CREDIT_NOTE", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As DataTable
        Try
            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.*,case when len(isnull(TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,''))>0 then TSPL_VENDOR_INVOICE_DETAIL.DeductionCode else TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction end as DeductionCode , case when len(isnull(TSPL_VENDOR_INVOICE_DETAIL.DeductionCode,''))>0 then TSPL_VENDOR_INVOICE_DETAIL.Deduction_Desc else TSPL_DCS_ADDITION_DEDUCTION.Description end as Deduction_Desc 
from TSPL_PAYMENT_PROCESS_CREDIT_NOTE 
left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No =TSPL_PAYMENT_PROCESS_CREDIT_NOTE.AP_Invoice_No
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_VENDOR_INVOICE_DETAIL.Document_No 
left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code=TSPL_VENDOR_INVOICE_DETAIL.DCS_Addition_Deduction  
left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
where TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_CREDIT_NOTE.SLNO as int)"
            '            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_CREDIT_NOTE.* 
            'from TSPL_PAYMENT_PROCESS_CREDIT_NOTE 
            'left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE
            'where TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_CREDIT_NOTE.SLNO as int)"

            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByRef ArrIndex As Dictionary(Of String, clsPaymentProcessIndex), ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessCreditNote)
        Try
            Dim arr As New List(Of clsPaymentProcessCreditNote)
            Dim obj As New clsPaymentProcessCreditNote
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_CREDIT_NOTE where Doc_No='" & doc_No & "'"
            If ApplyIndex Then
                q += " order by TSPL_PAYMENT_PROCESS_CREDIT_NOTE.Vendor_CODE "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                Dim IdxVendor As String = clsCommon.myCstr(dtbl.Rows(0)("Vendor_CODE")).ToUpper()
                Dim IdxFrom As Integer = 0
                Dim IdxTo As Integer = -1
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    If ApplyIndex Then
                        If (Not (clsCommon.CompairString(IdxVendor, clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))) = CompairStringResult.Equal)) Then
                            IdxTo = i - 1
                            ArrIndex(IdxVendor).CRFrom = IdxFrom
                            ArrIndex(IdxVendor).CRTo = IdxTo

                            IdxVendor = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE")).ToUpper()
                            IdxFrom = i
                        End If
                        If i = dtbl.Rows.Count - 1 Then
                            IdxTo = i
                            ArrIndex(IdxVendor).CRFrom = IdxFrom
                            ArrIndex(IdxVendor).CRTo = IdxTo
                        End If
                    End If

                    obj = New clsPaymentProcessCreditNote
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_CODE"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_NAME"))
                    obj.TDS_Amount = clsCommon.myCDecimal(dtbl.Rows(i)("TDS_Amount"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    ' obj.Area_Location_Code = clsCommon.myCstr(dtbl.Rows(i)("Area_Location_Code"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_CREDIT_NOTE where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessMCCSaleReturn
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public Return_Doc_No As String = ""
    Public Return_Doc_Type As String = ""
    Public Return_Doc_Date As String = ""
    Public Shipment_Doc_No As String = ""
    Public Shipment_Doc_Date As String = ""
    Public Sale_Doc_No As String = ""
    Public Sale_Doc_Date As String = ""
    Public AR_Invoice_No As String = ""
    Public AR_Invoice_Date As String = ""
    Public Customer_CODE As String = ""
    Public Customer_NAME As String = ""
    Public Item_Code As String = ""
    Public Item_Desc As String = ""
    Public Amount As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessMCCSaleReturn), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_MCC_SALE_Return where  Doc_No='" & DocNo & "'"
                clsDBFuncationality.ExecuteNonQuery(qry, tran)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "Return_Doc_No", arr.Item(i).Return_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Return_Doc_TYpe", arr.Item(i).Return_Doc_Type)
                    clsCommon.AddColumnsForChange(coll, "Return_Doc_Date", arr.Item(i).Return_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "Shipment_Doc_No", arr.Item(i).Shipment_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Shipment_Doc_Date", arr.Item(i).Shipment_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "Sale_Doc_No", arr.Item(i).Sale_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Sale_Doc_Date", arr.Item(i).Sale_Doc_Date)
                    clsCommon.AddColumnsForChange(coll, "AR_Invoice_No", arr.Item(i).AR_Invoice_No)
                    clsCommon.AddColumnsForChange(coll, "AR_Invoice_Date", arr.Item(i).AR_Invoice_Date)
                    clsCommon.AddColumnsForChange(coll, "Customer_CODE", arr.Item(i).Customer_CODE)
                    clsCommon.AddColumnsForChange(coll, "Customer_NAME", arr.Item(i).Customer_NAME)
                    clsCommon.AddColumnsForChange(coll, "Item_Code", arr.Item(i).Item_Code)
                    clsCommon.AddColumnsForChange(coll, "Item_Desc", arr.Item(i).Item_Desc)
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(arr.Item(i).Amount))
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_MCC_SALE_Return", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As DataTable
        Try
            Dim q As String = "select  cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_MCC_SALE_Return.* 
from TSPL_PAYMENT_PROCESS_MCC_SALE_Return 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_MCC_SALE_Return.Customer_CODE
where TSPL_PAYMENT_PROCESS_MCC_SALE_Return.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_MCC_SALE_Return.SLNO as int) "
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessMCCSaleReturn)
        Try
            Dim arr As New List(Of clsPaymentProcessMCCSaleReturn)
            Dim obj As New clsPaymentProcessMCCSaleReturn
            Dim q As String = "select * from TSPL_PAYMENT_PROCESS_MCC_SALE_Return where Doc_No='" & doc_No & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsPaymentProcessMCCSaleReturn
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.Shipment_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Shipment_Doc_No"))
                    obj.Shipment_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Shipment_Doc_Date"))
                    obj.Return_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Return_Doc_No"))
                    obj.Return_Doc_Type = clsCommon.myCstr(dtbl.Rows(i)("Return_Doc_Type"))
                    obj.Return_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Return_Doc_Date"))
                    obj.Sale_Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Sale_Doc_No"))
                    obj.Sale_Doc_Date = clsCommon.myCstr(dtbl.Rows(i)("Sale_Doc_Date"))
                    obj.AR_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AR_Invoice_No"))
                    obj.AR_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("AR_Invoice_Date"))
                    obj.Customer_CODE = clsCommon.myCstr(dtbl.Rows(i)("Customer_CODE"))
                    obj.Customer_NAME = clsCommon.myCstr(dtbl.Rows(i)("Customer_NAME"))
                    obj.Item_Code = clsCommon.myCstr(dtbl.Rows(i)("Item_Code"))
                    obj.Item_Desc = clsCommon.myCstr(dtbl.Rows(i)("Item_Desc"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Amount"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_MCC_SALE_Return where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessAdvancePayment
#Region "Variables"
    Public Doc_No As String = ""
    Public SNo As Integer = 0
    Public Payment_No As String = ""
    Public Vendor_Code As String = ""
    Public Payment_Amount As Decimal = 0
    Public Installment_Amount As Decimal = 0
    Public Payment_Balance As Decimal = 0
    Public Amount_Knock_Off As Decimal = 0
    ''No a table column
    Public Payment_Date As String = ""
    Public Vendor_Name As String = ""
    Public No_Of_EMI As String = ""
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessAdvancePayment), ByVal arrPPD As List(Of clsPaymentProcessDetail), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SNo", arr.Item(i).SNo)
                    clsCommon.AddColumnsForChange(coll, "Payment_No", arr.Item(i).Payment_No)
                    clsCommon.AddColumnsForChange(coll, "Vendor_Code", arr.Item(i).Vendor_Code)
                    clsCommon.AddColumnsForChange(coll, "Payment_Amount", arr.Item(i).Payment_Amount)
                    clsCommon.AddColumnsForChange(coll, "Installment_Amount", arr.Item(i).Installment_Amount)
                    clsCommon.AddColumnsForChange(coll, "Payment_Balance", arr.Item(i).Payment_Balance)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If

            If arrPPD IsNot Nothing AndAlso arrPPD.Count > 0 Then
                For i = 0 To arrPPD.Count - 1
                    'BHA/23/05/18-000034 by balwinder on 25/05/2018
                    'Dim qry As String = "select TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No,TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_Balance from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No where TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Doc_No='" + DocNo + "' and TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Vendor_Code='" + arrPPD(i).VSP_CODE + "' order by TSPL_PAYMENT_HEADER.Payment_Date "
                    Dim qry As String = "select Payment_No,case when No_Of_EMI=0 then Payment_Balance else case when InsatllmentAmt<Payment_Balance then InsatllmentAmt else Payment_Balance end end as Payment_Balance from (" + Environment.NewLine +
                    "select TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No,TSPL_PAYMENT_HEADER.Payment_Date,TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_Balance" + Environment.NewLine +
                    ",isnull(TSPL_PAYMENT_HEADER.No_Of_EMI,0) as No_Of_EMI" + Environment.NewLine
                    ''Comment by balwinder on 11/12/2018 due to give option to Change installment amount.
                    '",cast( case when TSPL_PAYMENT_HEADER.Payment_Amount<>0 and isnull(TSPL_PAYMENT_HEADER.No_Of_EMI,0)<>0 then  TSPL_PAYMENT_HEADER.Payment_Amount/isnull(TSPL_PAYMENT_HEADER.No_Of_EMI,0) else 0 end as decimal(18,2)) as InsatllmentAmt" + Environment.NewLine + _
                    qry += ",cast( case when TSPL_PAYMENT_HEADER.Payment_Amount<>0 and isnull(TSPL_PAYMENT_HEADER.No_Of_EMI,0)<>0 then  TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Installment_Amount else 0 end as decimal(18,2)) as InsatllmentAmt " + Environment.NewLine +
                    "from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT " + Environment.NewLine +
                    "left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No " + Environment.NewLine +
                    "where TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Doc_No='" + DocNo + "' and TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Vendor_Code='" + arrPPD(i).VSP_CODE + "'" + Environment.NewLine +
                    ")x" + Environment.NewLine +
                    "order by Payment_Date "
                    Dim KnockOffAmt As Decimal = arrPPD(i).Advance_Payment_Amount_Knock_Off
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, tran)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        For Each dr As DataRow In dt.Rows
                            If KnockOffAmt <= 0 Then
                                Exit For
                            End If
                            Dim ApplicableAmt As Decimal = 0
                            If KnockOffAmt >= clsCommon.myCDecimal(dr("Payment_Balance")) Then
                                ApplicableAmt = clsCommon.myCDecimal(dr("Payment_Balance"))
                            Else
                                ApplicableAmt = KnockOffAmt
                            End If
                            KnockOffAmt -= ApplicableAmt

                            Dim coll As New Hashtable()
                            clsCommon.AddColumnsForChange(coll, "Amount_Knock_Off", ApplicableAmt)
                            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT", OMInsertOrUpdate.Update, "Doc_No='" + DocNo + "'  and Payment_No='" + clsCommon.myCstr(dr("Payment_No")) + "'", tran)
                        Next
                    End If
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function AutoKnockOff(ByVal DocNo As String, ByVal trans As SqlTransaction) As Boolean

        Return True
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.*, TSPL_PAYMENT_HEADER.Payment_Date, TSPL_VENDOR_MASTER.Vendor_Name,TSPL_PAYMENT_HEADER.No_Of_EMI 
from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT  
left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No 
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Vendor_Code 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Vendor_Code
where TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Doc_No='" & doc_No & "' order by  TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.SNo  "
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessAdvancePayment)
        Try
            Dim arr As New List(Of clsPaymentProcessAdvancePayment)
            Dim obj As New clsPaymentProcessAdvancePayment
            Dim q As String = "select TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.*, TSPL_PAYMENT_HEADER.Payment_Date, TSPL_VENDOR_MASTER.Vendor_Name,TSPL_PAYMENT_HEADER.No_Of_EMI from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Payment_No left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT.Vendor_Code where Doc_No='" & doc_No & "' order by SNo"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsPaymentProcessAdvancePayment
                    obj.Doc_No = clsCommon.myCstr(dt.Rows(i)("Doc_No"))
                    obj.SNo = clsCommon.myCstr(dt.Rows(i)("SNo"))
                    obj.Payment_No = clsCommon.myCstr(dt.Rows(i)("Payment_No"))
                    obj.Payment_Date = clsCommon.GetPrintDate(dt.Rows(i)("Payment_Date"), "dd/MM/yyyy")
                    obj.Vendor_Code = clsCommon.myCstr(dt.Rows(i)("Vendor_Code"))
                    obj.Vendor_Name = clsCommon.myCstr(dt.Rows(i)("Vendor_Name"))
                    obj.Payment_Amount = clsCommon.myCDecimal(dt.Rows(i)("Payment_Amount"))
                    obj.Installment_Amount = clsCommon.myCDecimal(dt.Rows(i)("Installment_Amount"))
                    obj.Payment_Balance = clsCommon.myCDecimal(dt.Rows(i)("Payment_Balance"))
                    obj.Amount_Knock_Off = clsCommon.myCDecimal(dt.Rows(i)("Amount_Knock_Off"))
                    obj.No_Of_EMI = clsCommon.myCDecimal(dt.Rows(i)("No_Of_EMI"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class clsPaymentProcessSkipDoc
#Region "Variables"
    Public Doc_No As String = ""
    Public Source_Doc_No As String = ""
    Public Source_Doc_Type As String = ""
    Public Vendor_Code As String = ""
    Public Balance_Amount As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessSkipDoc), ByVal tran As SqlTransaction) As Boolean
        Try
            Dim qry As String = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where  Source_Doc_Type='MCC-SALE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Sale_Doc_No from TSPL_PAYMENT_PROCESS_MCC_SALE where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='MCC-SALE-RET' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Return_Doc_No from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='VSP-ITEM-ISSUE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Item_Issue_Doc_No from TSPL_PAYMENT_PROCESS_ITEM_ISSUE where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='VSP-ITEM-ISSUE-RETURN' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Item_Issue_Return_No  from TSPL_PAYMENT_PROCESS_ITEM_ISSUE_RETURN where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='DEBIT-NOTE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select AP_Invoice_No from TSPL_PAYMENT_PROCESS_DEDUCTION where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='CREDIT-NOTE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select AP_Invoice_No from TSPL_PAYMENT_PROCESS_CREDIT_NOTE where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='MILK-PUR-INVOICE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Milk_Purchase_Invoice_No  from TSPL_PAYMENT_PROCESS_DETAIL where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Source_Doc_Type='ADVANCE' and Doc_No='" + DocNo + "' and Source_Doc_No in (select Payment_No  from TSPL_PAYMENT_PROCESS_ADVANCE_PAYMENT where Doc_No='" + DocNo + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)

            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i As Integer = 0 To arr.Count - 1
                    qry = "Delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where Doc_No='" + DocNo + "' and Source_Doc_No='" + arr.Item(i).Source_Doc_No + "' and Source_Doc_Type='" + arr.Item(i).Source_Doc_Type + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, tran)

                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "Source_Doc_No", arr.Item(i).Source_Doc_No)
                    clsCommon.AddColumnsForChange(coll, "Source_Doc_Type", arr.Item(i).Source_Doc_Type)
                    clsCommon.AddColumnsForChange(coll, "Vendor_Code", arr.Item(i).Vendor_Code)
                    clsCommon.AddColumnsForChange(coll, "Balance_Amount", arr.Item(i).Balance_Amount)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function DeleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_SKIP_DOCUMENT where  Doc_No='" & DocNo & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


End Class



Public Class clsPaymentProcessAssetLost
#Region "Variables"
    Public Doc_No As String = ""
    Public SNo As Integer = 0
    Public Payment_No As String = ""
    Public Vendor_Code As String = ""
    Public Payment_Amount As Decimal = 0
    ''No a table column
    Public Payment_Date As String = ""
    Public Vendor_Name As String = ""
    Public No_Of_EMI As String = ""
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessAssetLost), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SNo", arr.Item(i).SNo)
                    clsCommon.AddColumnsForChange(coll, "Payment_No", arr.Item(i).Payment_No)
                    clsCommon.AddColumnsForChange(coll, "Vendor_Code", arr.Item(i).Vendor_Code)
                    clsCommon.AddColumnsForChange(coll, "Payment_Amount", arr.Item(i).Payment_Amount)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_ASSET_LOST", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_ASSET_LOST.*, TSPL_PAYMENT_HEADER.Payment_Date, TSPL_VENDOR_MASTER.Vendor_Name,TSPL_PAYMENT_HEADER.No_Of_EMI 
from TSPL_PAYMENT_PROCESS_ASSET_LOST  
left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ASSET_LOST.Payment_No 
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_ASSET_LOST.Vendor_Code 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code
where TSPL_PAYMENT_PROCESS_ASSET_LOST.Doc_No='" & doc_No & "' order by TSPL_PAYMENT_PROCESS_ASSET_LOST.SNo"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessAssetLost)
        Try
            Dim arr As New List(Of clsPaymentProcessAssetLost)
            Dim obj As New clsPaymentProcessAssetLost
            Dim q As String = "select TSPL_PAYMENT_PROCESS_ASSET_LOST.*, TSPL_PAYMENT_HEADER.Payment_Date, TSPL_VENDOR_MASTER.Vendor_Name,TSPL_PAYMENT_HEADER.No_Of_EMI from TSPL_PAYMENT_PROCESS_ASSET_LOST  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_PROCESS_ASSET_LOST.Payment_No left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_ASSET_LOST.Vendor_Code where Doc_No='" & doc_No & "' order by SNo"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsPaymentProcessAssetLost
                    obj.Doc_No = clsCommon.myCstr(dt.Rows(i)("Doc_No"))
                    obj.SNo = clsCommon.myCstr(dt.Rows(i)("SNo"))
                    obj.Payment_No = clsCommon.myCstr(dt.Rows(i)("Payment_No"))
                    obj.Payment_Date = clsCommon.GetPrintDate(dt.Rows(i)("Payment_Date"), "dd/MM/yyyy")
                    obj.Vendor_Code = clsCommon.myCstr(dt.Rows(i)("Vendor_Code"))
                    obj.Vendor_Name = clsCommon.myCstr(dt.Rows(i)("Vendor_Name"))
                    obj.Payment_Amount = clsCommon.myCDecimal(dt.Rows(i)("Payment_Amount"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_ASSET_LOST where  Doc_No='" & DocNo & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, tran)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

End Class


Public Class clsPaymentProcessSaving
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Type As String = "" ''Not a Table Column
    Public AP_Invoice_Date As String = "" ''Not a Table Column 
    Public Vendor_CODE As String = "" ''Not a Table Column
    Public Vendor_NAME As String = "" ''Not a Table Column
    Public Amount As Decimal = 0 ''Not a Table Column
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessSaving), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_SAVING", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, ByVal trans As SqlTransaction) As DataTable
        Try
            Dim q As String = "select  cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_SAVING.*,TSPL_VENDOR_INVOICE_HEAD.Posting_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type
,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,TSPL_VENDOR_INVOICE_HEAD.Vendor_Name,TSPL_VENDOR_INVOICE_HEAD.Document_Total
from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
where TSPL_PAYMENT_PROCESS_SAVING.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_SAVING.SLNO as int)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByRef ArrIndex As Dictionary(Of String, clsPaymentProcessIndex), ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessSaving)
        Try
            Dim arr As New List(Of clsPaymentProcessSaving)
            Dim obj As New clsPaymentProcessSaving
            Dim q As String = "select TSPL_PAYMENT_PROCESS_SAVING.*,TSPL_VENDOR_INVOICE_HEAD.Posting_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type
,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,TSPL_VENDOR_INVOICE_HEAD.Vendor_Name,TSPL_VENDOR_INVOICE_HEAD.Document_Total
from TSPL_PAYMENT_PROCESS_SAVING left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No where TSPL_PAYMENT_PROCESS_SAVING.Doc_No='" & doc_No & "'"
            If ApplyIndex Then
                q += " order by TSPL_VENDOR_INVOICE_HEAD.Vendor_Code "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                Dim IdxVendor As String = clsCommon.myCstr(dtbl.Rows(0)("Vendor_Code")).ToUpper()
                Dim IdxFrom As Integer = 0
                Dim IdxTo As Integer = -1
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    If ApplyIndex Then
                        If (Not (clsCommon.CompairString(IdxVendor, clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code"))) = CompairStringResult.Equal)) Then
                            IdxTo = i - 1
                            ArrIndex(IdxVendor).SavingFrom = IdxFrom
                            ArrIndex(IdxVendor).SavingTo = IdxTo

                            IdxVendor = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code")).ToUpper()
                            IdxFrom = i
                        End If
                        If i = dtbl.Rows.Count - 1 Then
                            IdxTo = i
                            ArrIndex(IdxVendor).SavingFrom = IdxFrom
                            ArrIndex(IdxVendor).SavingTo = IdxTo
                        End If
                    End If

                    obj = New clsPaymentProcessSaving
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Type = clsCommon.myCstr(dtbl.Rows(i)("Document_Type"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("Posting_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Name"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Document_Total"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_SAVING where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessCompulsory
#Region "Variables"
    Public Doc_No As String = ""
    Public SLNO As String = ""
    Public AP_Invoice_No As String = ""
    Public AP_Invoice_Type As String = "" ''Not a Table Column
    Public AP_Invoice_Date As String = "" ''Not a Table Column 
    Public Vendor_CODE As String = "" ''Not a Table Column
    Public Vendor_NAME As String = "" ''Not a Table Column
    Public Amount As Decimal = 0 ''Not a Table Column
#End Region

    Public Shared Function SaveData(ByVal DocNo As String, ByVal arr As List(Of clsPaymentProcessCompulsory), Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then

                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Doc_No", DocNo)
                    clsCommon.AddColumnsForChange(coll, "SLNO", arr.Item(i).SLNO)
                    clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", arr.Item(i).AP_Invoice_No)
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_PROCESS_COMPULSORY", OMInsertOrUpdate.Insert, "", tran)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getDataDT(ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As DataTable
        Try
            Dim arr As New List(Of clsPaymentProcessCompulsory)
            Dim obj As New clsPaymentProcessCompulsory
            Dim q As String = "select  cast(1 as bit) as Sel,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_PAYMENT_PROCESS_COMPULSORY.*,TSPL_VENDOR_INVOICE_HEAD.Posting_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type
,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,TSPL_VENDOR_INVOICE_HEAD.Vendor_Name,TSPL_VENDOR_INVOICE_HEAD.Document_Total
from TSPL_PAYMENT_PROCESS_COMPULSORY 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
where TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No='" & doc_No & "' order by cast(TSPL_PAYMENT_PROCESS_COMPULSORY.SLNO as int)"
            Return clsDBFuncationality.GetDataTable(q, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal ApplyIndex As Boolean, ByRef ArrIndex As Dictionary(Of String, clsPaymentProcessIndex), ByVal doc_No As String, Optional ByVal trans As SqlTransaction = Nothing) As List(Of clsPaymentProcessCompulsory)
        Try
            Dim arr As New List(Of clsPaymentProcessCompulsory)
            Dim obj As New clsPaymentProcessCompulsory
            Dim q As String = "select TSPL_PAYMENT_PROCESS_COMPULSORY.*,TSPL_VENDOR_INVOICE_HEAD.Posting_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Type
,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,TSPL_VENDOR_INVOICE_HEAD.Vendor_Name,TSPL_VENDOR_INVOICE_HEAD.Document_Total
from TSPL_PAYMENT_PROCESS_COMPULSORY 
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No 
where TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No='" & doc_No & "'"
            If ApplyIndex Then
                q += " order by TSPL_VENDOR_INVOICE_HEAD.Vendor_Code "
            End If
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q, trans)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                Dim IdxVendor As String = clsCommon.myCstr(dtbl.Rows(0)("Vendor_Code")).ToUpper()
                Dim IdxFrom As Integer = 0
                Dim IdxTo As Integer = -1

                For i As Integer = 0 To dtbl.Rows.Count - 1
                    If ApplyIndex Then
                        If (Not (clsCommon.CompairString(IdxVendor, clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code"))) = CompairStringResult.Equal)) Then
                            IdxTo = i - 1
                            ArrIndex(IdxVendor).COMPULSORYFrom = IdxFrom
                            ArrIndex(IdxVendor).COMPULSORYTo = IdxTo

                            IdxVendor = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code")).ToUpper()
                            IdxFrom = i
                        End If
                        If i = dtbl.Rows.Count - 1 Then
                            IdxTo = i
                            ArrIndex(IdxVendor).COMPULSORYFrom = IdxFrom
                            ArrIndex(IdxVendor).COMPULSORYTo = IdxTo
                        End If
                    End If

                    obj = New clsPaymentProcessCompulsory
                    obj.Doc_No = clsCommon.myCstr(dtbl.Rows(i)("Doc_No"))
                    obj.SLNO = clsCommon.myCstr(dtbl.Rows(i)("SLNO"))
                    obj.AP_Invoice_No = clsCommon.myCstr(dtbl.Rows(i)("AP_Invoice_No"))
                    obj.AP_Invoice_Type = clsCommon.myCstr(dtbl.Rows(i)("Document_Type"))
                    obj.AP_Invoice_Date = clsCommon.myCstr(dtbl.Rows(i)("Posting_Date"))
                    obj.Vendor_CODE = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Code"))
                    obj.Vendor_NAME = clsCommon.myCstr(dtbl.Rows(i)("Vendor_Name"))
                    obj.Amount = clsCommon.myCDecimal(dtbl.Rows(i)("Document_Total"))
                    arr.Add(obj)
                Next
            End If
            Return arr
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function deleteData(ByVal DocNo As String, ByVal tran As SqlTransaction) As Boolean
        Try
            Dim isDeleted As Boolean = True
            Dim qry As String = "delete from TSPL_PAYMENT_PROCESS_COMPULSORY where  Doc_No='" & DocNo & "'"
            isDeleted = isDeleted AndAlso clsDBFuncationality.ExecuteNonQuery(qry, tran)
            Return isDeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsPaymentProcessIndex
    Public SavingFrom As Integer
    Public SavingTo As Integer
    Public DRFrom As Integer
    Public DRTo As Integer
    Public CRFrom As Integer
    Public CRTo As Integer
    Public COMPULSORYFrom As Integer
    Public COMPULSORYTo As Integer
End Class

