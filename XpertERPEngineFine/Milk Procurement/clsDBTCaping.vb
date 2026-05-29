Imports System.Data.SqlClient

Public Class clsDBTCaping
#Region "variables"
    Public Document_Code As String = Nothing
    Public Document_Date As Date
    Public Reco_Code As String = Nothing
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Posting_Date As Date? = Nothing
    Public arr As List(Of clsDBTCapingDetail) = Nothing

#End Region
    Public Shared Function SaveData(ByVal obj As clsDBTCaping, ByVal isNewEntry As Boolean) As Boolean
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
    Public Shared Function SaveData(ByVal obj As clsDBTCaping, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        Try
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DBTCappingCheck, "", obj.Document_Date, trans)
            qry = "delete from TSPL_DBT_CAPING_DETAIL where Document_Code='" & obj.Document_Code & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            clsCommon.AddColumnsForChange(coll, "Reco_Code", obj.Reco_Code)
            If isNewEntry Then
                obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DBTCaping, "", "")
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_CAPING", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_CAPING", OMInsertOrUpdate.Update, "TSPL_DBT_CAPING.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            clsDBTCapingDetail.saveData(obj.Document_Code, obj.arr, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DBT_CAPING", "Document_Code", "TSPL_DBT_CAPING_DETAIL", "Document_Code", "", "", trans)
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction, ByVal whrclas As String) As clsDBTCaping
        Dim obj As clsDBTCaping = Nothing
        Dim Arr As List(Of clsDBTCaping) = Nothing
        Dim qry As String = "Select TSPL_DBT_CAPING.* from TSPL_DBT_CAPING  where 2=2 " + whrclas
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_DBT_CAPING.Document_Code = (select MIN(Document_Code) from TSPL_DBT_CAPING WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Last
                qry += " and TSPL_DBT_CAPING.Document_Code = (select Max(Document_Code) from TSPL_DBT_CAPING WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Current
                qry += " and TSPL_DBT_CAPING.Document_Code = '" + strCode + "' "
            Case NavigatorType.Next
                qry += " and TSPL_DBT_CAPING.Document_Code = (select Min(Document_Code) from TSPL_DBT_CAPING where Document_Code>'" + strCode + "' " + whrclas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_DBT_CAPING.Document_Code = (select Max(Document_Code) from TSPL_DBT_CAPING where Document_Code<'" + strCode + "' " + whrclas + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsDBTCaping()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Reco_Code = clsCommon.myCstr(dt.Rows(0)("Reco_Code"))
            obj.Status = IIf(clsCommon.myCDecimal(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
            End If
            'obj.dt = clsDBTCapingDetail.getData(obj.Document_Code, SelectedZone, trans)
        End If
        Return obj
    End Function
    Public Shared Function DeleteData(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            DeleteData(strDocNo, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function DeleteData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Document No not found to Delete")
        End If
        Try
            clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_CAPING", "Document_Code", "TSPL_DBT_CAPING_DETAIL", "Document_Code", "", "", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_CAPING", "Document_Code", "TSPL_DBT_CAPING_DETAIL", "Document_Code", "", "", trans)
            Dim qry As String = ""
            qry = "delete from TSPL_DBT_CAPING_DETAIL where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_CAPING where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = "  Select TSPL_DBT_CAPING.Document_Code as Code,Convert(varchar,TSPL_DBT_CAPING.Document_Date,103) as Date,Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) as [From Date],Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To,103)as [To Date],Reco_Code as Zone,case when isnull(TSPL_DBT_CAPING.Status,0)=0 then 'Pending' else 'Approved' end as Status 
from TSPL_DBT_CAPING 
left outer join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=TSPL_DBT_CAPING.Reco_Code "
        str = clsCommon.ShowSelectForm("DCMPInc#F", qry, "Code", "2=2 " + whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If
            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")

            Dim obj As clsDBTCaping = clsDBTCaping.GetData(strDocNo, NavigatorType.Current, trans, "")
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            Dim qry As String = "Update TSPL_DBT_CAPING set Status=1 , Posting_Date='" + strPostDate + "',Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DBT_CAPING", "Document_Code", "TSPL_DBT_CAPING_DETAIL", "Document_Code", "", "", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
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
    Public Shared Function ReverseAndUnpost(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            'Throw New Exception("Not implemented")
            Dim obj As clsDBTCaping = clsDBTCaping.GetData(strDocNo, NavigatorType.Current, trans, "")
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Reverse And UnPost")
            End If
            If Not obj.Status = ERPTransactionStatus.Approved Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If
            Dim qry As String = "select TSPL_DBT_NEFT.Document_Code 
from  TSPL_DCS_MP_INCENTIVE_RECO_HEAD 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.From_Date=TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date and TSPL_DBT_NEFT.To_Date=TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To
where TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code='" + obj.Reco_Code + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.myLen(dt.Rows(0)("Document_Code")) > 0 Then
                    Throw New Exception("DBT NEFT No [" + dt.Rows(0)("Document_Code") + "] created you cant unpost current document")
                End If
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Status", 0)
            clsCommon.AddColumnsForChange(coll, "Posted_By", Nothing, True)
            clsCommon.AddColumnsForChange(coll, "Posting_Date", Nothing, True)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_CAPING", OMInsertOrUpdate.Update, "Document_Code='" + obj.Document_Code + "'", trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DBT_CAPING", "Document_Code", "TSPL_DBT_CAPING_DETAIL", "Document_Code", "", "", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsDBTCapingDetail
#Region "Variable"
    Public PK_Id As Integer = Nothing
    Public Document_Code As String = Nothing
    Public BMC_Code As String = Nothing
    Public BMC_Uploader_Code As String = Nothing ''Not a Table Column
    Public BMC_Name As String = Nothing ''Not a Table Column
    Public DCS_Code As String = Nothing
    Public DCS_Uploader_Code As String = Nothing ''Not a Table Column
    Public DCS_Name As String = Nothing ''Not a Table Column
    Public MP_Code As String = Nothing
    Public MP_Uploader_Code As String = Nothing ''Not a Table Column
    Public MP_Name As String = Nothing ''Not a Table Column
    Public Qty As Decimal = 0
    Public Capping_Qty As Decimal = 0
    Public Capping_Status As Boolean = False
    Public Capping_Increase_By As String = Nothing
    Public Capping_Increase_Date As DateTime? = Nothing
    Public Capping_Increase_Remarks As String = Nothing
#End Region

    Public Shared Function saveData(ByVal strDocNo As String, ByVal arrObj As List(Of clsDBTCapingDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            If arrObj IsNot Nothing Then
                For Each obj As clsDBTCapingDetail In arrObj
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "BMC_Code", obj.BMC_Code)
                    clsCommon.AddColumnsForChange(coll, "DCS_Code", obj.DCS_Code)
                    clsCommon.AddColumnsForChange(coll, "MP_Code", obj.MP_Code)
                    clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)
                    clsCommon.AddColumnsForChange(coll, "Capping_Qty", obj.Capping_Qty)
                    clsCommon.AddColumnsForChange(coll, "Capping_Status", IIf(obj.Capping_Status, 1, 0))
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_CAPING_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getData(ByVal document_Code As String, ByVal whrcls As String) As DataTable
        Dim qry As String = "select TSPL_DBT_CAPING_DETAIL.PK_Id,ROW_NUMBER() over (order by TSPL_DBT_CAPING_DETAIL.PK_Id) as SNo
,TSPL_DBT_CAPING_DETAIL.BMC_Code,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as BMC_Uploader_Code,TSPL_MCC_MASTER.MCC_NAME as BMC_Name
,TSPL_DBT_CAPING_DETAIL.DCS_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as DCS_Uploader_Code,TSPL_VLC_MASTER_HEAD.VLC_Name as DCS_Name,
TSPL_DBT_CAPING_DETAIL.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as MP_Uploader_Code,TSPL_MP_MASTER.MP_Name,
TSPL_DBT_CAPING_DETAIL.Qty,TSPL_DBT_CAPING_DETAIL.Capping_Qty,TSPL_DBT_CAPING_DETAIL.Capping_Status,TSPL_DBT_CAPING_DETAIL.Capping_Increase_By,TSPL_DBT_CAPING_DETAIL.Capping_Increase_Date,TSPL_DBT_CAPING_DETAIL.Capping_Increase_Remarks ,TSPL_MP_MASTER.THIRD_PARTY_CODE
from TSPL_DBT_CAPING_DETAIL
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_DBT_CAPING_DETAIL.MP_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DBT_CAPING_DETAIL.DCS_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_DBT_CAPING_DETAIL.BMC_Code
where TSPL_DBT_CAPING_DETAIL.Document_Code='" + document_Code + "' " + whrcls
        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function CappingIncrease(ByVal strPKId As String, ByVal Remarks As String, ByVal CappingQtyPerDay As Integer, ByVal DaysInMonth As Integer) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If (clsCommon.myLen(strPKId) <= 0) Then
                Throw New Exception("Invalid Farmer Detail")
            End If
            Dim qry As String = "select pk_id,Qty,MP_Code,Capping_Qty,Capping_Status from TSPL_DBT_CAPING_DETAIL where pk_id=" + strPKId + ""
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("Invalid Farmer Details")
            End If
            If clsCommon.myCDecimal(dt.Rows(0)("Capping_Status")) = 1 Then
                Throw New Exception("Farmer capping Details should be invalid type")
            End If
            If clsCommon.myCDecimal(dt.Rows(0)("Qty")) > (CappingQtyPerDay * DaysInMonth) Then
                Throw New Exception("DBT Qty [" + clsCommon.myCstr(dt.Rows(0)("Qty")) + "] and Capping Qty [" + clsCommon.myCstr((CappingQtyPerDay * DaysInMonth)) + "]")
            End If

            qry = "Update TSPL_MP_MASTER set  DBT_Capping_Qty='" + clsCommon.myCstr(CappingQtyPerDay) + "' where MP_Code='" + clsCommon.myCstr(dt.Rows(0)("MP_Code")) + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(dt.Rows(0)("MP_Code")), "TSPL_MP_MASTER", "MP_Code", trans)


            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")
            qry = "Update TSPL_DBT_CAPING_DETAIL set  Capping_Status=1,Capping_Qty=" + clsCommon.myCstr((CappingQtyPerDay * DaysInMonth)) + " , Capping_Increase_Date='" + strPostDate + "',Capping_Increase_By='" + objCommonVar.CurrentUserCode + "',Capping_Increase_Remarks='" + Remarks + "' where PK_Id='" + strPKId + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class
