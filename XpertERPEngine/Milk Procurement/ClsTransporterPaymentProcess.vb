Imports common
Imports System.Data.SqlClient

Public Class ClsTransporterPaymentProcess

#Region "Variables"
    Public Document_Code As String = Nothing
    Public Document_Date As DateTime
    Public From_Date As DateTime
    Public To_Date As DateTime
    Public Tanker_No As String = Nothing
    Public Toll_Tax As Double = 0
    Public Ice_Charge As Double = 0
    'Public Fat_Shortage As Double = 0
    'Public Snf_Shortage As Double = 0
    'Public Fat_Snf_Shortage As Double = 0
    'Public Fat_Rate As Double = 0
    Public Snf_Rate As Double = 0
    Public Fat_Shortage As Double = 0
    Public Fat_Shortage_NMG As Double = 0
    Public Snf_Shortage As Double = 0
    Public Snf_Shortage_NMG As Double = 0
    Public Fat_Snf_Shortage As Double = 0
    Public Fat_Rate As Double = 0
    Public Fat_Rate_NMG As Double = 0
    Public Snf_Rate_NMG As Double = 0

    Public Tanker_Capacity As Double = 0
    Public KM_Rate As Double = 0
    Public Total_Amount As Double = 0
    Public Gross_Amount As Double = 0
    Public Total_Addition As Double = 0
    Public Total_Deduction As Double = 0
    Public Diesel_Rate_Plus As Double = 0
    Public Diesel_Rate_Minus As Double = 0
    Public Total_Diesel As Double = 0
    Public Prorata_Amt As Double = 0
    Public Total_Before_Calc As Double = 0
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Posted_Date As DateTime?
    Public Arr As List(Of ClsTransporterPaymentProcessDetail) = Nothing
    Public ArrDT As DataTable = Nothing
    Public Comment As String = Nothing
    Public Remarks As String = Nothing
    Public Type As String = Nothing
    Public Is_Private As Boolean = False
    Public Total_PF_Amount As Double = 0
    Public Total_ESI_Amount As Double = 0
    Public Labour_Charge As Double = 0
    Public Labour_Amt As Double = 0
    Public Employee_PF As Double = 0
    Public Employer_PF As Double = 0
    Public Admin_Charge As Double = 0
    Public EDLI_Charge As Double = 0
    Public ESI_Employee As Double = 0
    Public ESI_Employer As Double = 0
    Public COESI_PER As Double = 0
    Public EMPESI_PER As Double = 0
    Public EMPEPF_PER As Double = 0
    Public ACCOEPF_PER As Double = 0



#End Region

    Public Function SaveData(ByVal obj As ClsTransporterPaymentProcess, ByVal isNewEntry As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            obj.SaveData(obj, isNewEntry, trans)
            trans.Commit()
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function

    Public Function SaveData(ByVal obj As ClsTransporterPaymentProcess, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean

        Dim isSaved As Boolean = True

        Dim qry As String = "delete from TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL where Document_Code='" + obj.Document_Code + "'"
        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

        If (isNewEntry) Then
            obj.Document_Code = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Document_Date), clsDocType.TransporterPaymentProcess, "", Nothing)
        End If

        If (clsCommon.myLen(obj.Document_Code) <= 0) Then
            Throw New Exception("Error in Document Code Generation")
        End If

        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "From_Date", clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "To_Date", clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "Comment", obj.Comment)
        clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
        clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
        clsCommon.AddColumnsForChange(coll, "Posted_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Posted_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
        If isNewEntry Then
            clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
            clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", OMInsertOrUpdate.Insert, "", trans)
        Else
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", OMInsertOrUpdate.Update, "Document_Code='" + obj.Document_Code + "'", trans)

        End If
        isSaved = isSaved AndAlso ClsTransporterPaymentProcessDetail.SaveData(obj.Document_Code, Arr, trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", "Document_Code", "TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL", "Document_Code", trans)
        Return isSaved
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

        Dim obj As ClsTransporterPaymentProcess = ClsTransporterPaymentProcess.GetData(strDocNo, NavigatorType.Current, False, trans)
        clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", "Document_Code", "TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL", "Document_Code", trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", "Document_Code", "TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL", "Document_Code", trans)

        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
            Try
                If obj.Status = ERPTransactionStatus.Approved Then
                    Throw New Exception("Already Post on :" + obj.Posted_Date)
                End If
                Dim qry As String = "delete from TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL where Document_Code='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where Document_Code='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType, ByVal isGetDT As Boolean, ByVal trans As SqlTransaction) As ClsTransporterPaymentProcess
        Dim obj As ClsTransporterPaymentProcess = Nothing

        Dim qry As String = "Select TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.* from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where 2=2 "
        Dim whrClas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code = (select MIN(Document_Code) from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Last
                qry += " and TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code = (select Max(Document_Code) from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Next
                qry += " and TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code = (select Min(Document_Code) from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where Document_Code>'" + strDocumentNo + "' " + whrClas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code = (select Max(Document_Code) from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD where Document_Code<'" + strDocumentNo + "' " + whrClas + ")"
            Case NavigatorType.Current
                qry += " and TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code = '" + strDocumentNo + "'"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New ClsTransporterPaymentProcess()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCstr(dt.Rows(0)("Document_Date"))
            obj.From_Date = clsCommon.myCstr(dt.Rows(0)("From_Date"))
            obj.To_Date = clsCommon.myCstr(dt.Rows(0)("To_Date"))
            obj.Comment = clsCommon.myCstr(dt.Rows(0)("Comment"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Type = clsCommon.myCstr(dt.Rows(0)("Type"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 0, ERPTransactionStatus.Pending, ERPTransactionStatus.Approved)

            qry = "Select TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL.* from TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL 
                   where TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL.Document_Code='" + obj.Document_Code + "' ORDER BY TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL.PK_ID"
            obj.ArrDT = clsDBFuncationality.GetDataTable(qry, trans)

            If (obj.ArrDT IsNot Nothing AndAlso obj.ArrDT.Rows.Count > 0) Then
                obj.Arr = New List(Of ClsTransporterPaymentProcessDetail)
                For Each dr As DataRow In obj.ArrDT.Rows
                    Dim objTr As New ClsTransporterPaymentProcessDetail
                    objTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objTr.PK_ID = clsCommon.myCstr(dr("PK_ID"))
                    objTr.Transporter_Bill_No = clsCommon.myCstr(dr("Transporter_Bill_No"))
                    objTr.Transporter_Bill_Date = clsCommon.myCDate(dr("Transporter_Bill_Date"))
                    objTr.Tanker_No = clsCommon.myCstr(dr("Tanker_No"))
                    objTr.KM = clsCommon.myCdbl(dr("KM"))
                    objTr.Transporter_Code = clsCommon.myCstr(dr("Transporter_Code"))
                    objTr.Type = clsCommon.myCstr(dr("Type"))
                    objTr.Bank_Code = clsCommon.myCstr(dr("Bank_Code"))
                    objTr.Bank_Name = clsCommon.myCstr(dr("Bank_Name"))
                    objTr.IFSC_Code = clsCommon.myCstr(dr("IFSC_Code"))
                    objTr.Amount = clsCommon.myCstr(dr("Amount"))
                    obj.Arr.Add(objTr)
                Next
            End If
        End If
        Return obj
    End Function

    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = ClsTransporterPaymentProcess.PostData(strDocNo, trans)
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

    Public Shared Function PostData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Document No not found to Post")
        End If

        Dim obj As ClsTransporterPaymentProcess = ClsTransporterPaymentProcess.GetData(strDocNo, NavigatorType.Current, False, trans)

        If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
            Throw New Exception("No Data found to Post")
        End If

        If (clsCommon.myLen(obj.Posted_Date) > 0) Then
            Throw New Exception("Already Post on :" + obj.Posted_Date)
        End If

        'CreateAPInvoiceHeader(obj, trans)
        qry = "Update TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD set Posted_Date='" + clsCommon.GetPrintDate(clsCommon.myCDate(clsCommon.GETSERVERDATE(trans)), "dd/MMM/yyyy hh:mm:ss tt") + "',Status=1 ,Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", "Document_Code", trans)

        Return True
    End Function

End Class


Public Class ClsTransporterPaymentProcessDetail

#Region "Variables"
    Public PK_ID As Integer = 0
    Public Document_Code As String = Nothing
    Public Amount As Decimal = 0
    Public Transporter_Bill_No As String = Nothing
    Public Type As String = Nothing
    Public Bank_Code As String = Nothing
    Public Bank_Name As String = Nothing
    Public IFSC_Code As String = Nothing
    Public Transporter_Code As String = Nothing
    Public GPS_KM As Decimal = 0
    Public KM As Decimal = 0
    Public Quantity_KG As Decimal = 0
    Public Diesel_RD As Decimal = 0
    Public BMC_Date As String = Nothing
    Public Ice_Box As String = Nothing
    Public arr As List(Of clsfrmVLCMaster) = Nothing
    Public Transporter_Bill_Date As DateTime
    Public Tanker_No As String = Nothing

    'Public BalanceAmount As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of ClsTransporterPaymentProcessDetail), ByVal trans As SqlTransaction) As Boolean

        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each obj As ClsTransporterPaymentProcessDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Transporter_Bill_No", obj.Transporter_Bill_No)
                clsCommon.AddColumnsForChange(coll, "Transporter_Bill_Date", clsCommon.GetPrintDate(obj.Transporter_Bill_Date, "dd/MMM/yyyy"))
                clsCommon.AddColumnsForChange(coll, "Tanker_No", obj.Tanker_No)
                clsCommon.AddColumnsForChange(coll, "KM", obj.KM)
                clsCommon.AddColumnsForChange(coll, "Type", obj.Type)
                clsCommon.AddColumnsForChange(coll, "Transporter_Code", obj.Transporter_Code)
                clsCommon.AddColumnsForChange(coll, "Bank_Code", obj.Bank_Code, True)
                clsCommon.AddColumnsForChange(coll, "Bank_Name", obj.Bank_Name, True)
                clsCommon.AddColumnsForChange(coll, "IFSC_Code", obj.IFSC_Code, True)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If

        Return True
    End Function



End Class

