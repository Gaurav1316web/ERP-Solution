Imports System.Data.SqlClient
Imports common
Public Class clsDBTNEFT
#Region "variables"
    Public Document_Code As String = Nothing
    Dim CycleWiseDBT As Boolean = False
    Public Document_Date As Date
    Public From_Date As Date
    Public DBT_Revise_Payment As Integer
    Public To_Date As Date
    Public Bank_Letter_Date As Date
    Public Remarks As String = ""
    Public Zone_Code As String = ""
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public RCDF_Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Posted_Date As Date? = Nothing
    Public arr As List(Of clsDBTNEFTDetail) = Nothing
    Public arrInvalid As List(Of clsDBTNEFTDetailInvalid) = Nothing
    Public arrHold As List(Of clsDBTNEFTDetailHold) = Nothing
    Public arrVLC As ArrayList = Nothing
    Public arrMCC As ArrayList = Nothing
#End Region

    Public Shared Function SaveData(ByVal obj As clsDBTNEFT, ByVal isNewEntry As Boolean) As Boolean
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_DBT_NEFT.Document_Date,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code= TSPL_DBT_NEFT_DETAIL.Document_Code
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.pk_id= TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code= TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
 where TSPL_DBT_NEFT_DETAIL.Document_Code ='" + obj.Document_Code + "'", trans)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DBTNEFTUploader, clsCommon.myCstr(dt.Rows(0)("MCC_Code")), clsCommon.myCDate(dt.Rows(0)("Document_Date")), trans)
            End If

            qry = "delete from TSPL_DBT_NEFT_DETAIL where Document_Code='" & obj.Document_Code & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_NEFT_DETAIL_INVALID where Document_Code='" & obj.Document_Code & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_NEFT_DETAIL_HOLD where Document_Code='" & obj.Document_Code & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Bank_Letter_Date", clsCommon.GetPrintDate(obj.Bank_Letter_Date, "dd/MMM/yyyy hh:mm tt"))

            clsCommon.AddColumnsForChange(coll, "From_Date", clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "To_Date", clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            clsCommon.AddColumnsForChange(coll, "DBT_Revise_Payment", obj.DBT_Revise_Payment)
            clsCommon.AddColumnsForChange(coll, "Zone_Code", obj.Zone_Code, True)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            If isNewEntry Then
                obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DBTNEFT, "", "")
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT", OMInsertOrUpdate.Insert, "", trans)
            Else

                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT", OMInsertOrUpdate.Update, "TSPL_DBT_NEFT.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            clsDBTNEFTDetail.saveData(obj.arr, obj.Document_Code, trans)
            clsDBTNEFTDetailInvalid.saveData(obj.arrInvalid, obj.Document_Code, trans)
            clsDBTNEFTDetailHold.saveData(obj.arrHold, obj.Document_Code, trans)
            SaveHistoryData(obj.Document_Code, trans)
            trans.Commit()
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function

    Private Shared Sub SaveHistoryData(document_Code As String, trans As SqlTransaction)
        clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, document_Code, "TSPL_DBT_NEFT", "Document_Code", "TSPL_DBT_NEFT_DETAIL", "Document_Code", "TSPL_DBT_NEFT_DETAIL_INVALID", "Document_Code", "TSPL_DBT_NEFT_DETAIL_HOLD", "Document_Code", "", "", "", "", "", "", trans)
    End Sub

    Public Shared Function SaveBankLetter(ByVal obj As clsDBTNEFT) As Boolean
        Dim trans As SqlTransaction = Nothing
        If SaveBankLetter(obj, trans) Then
            Return True
        End If
        Return True
    End Function

    Public Shared Function SaveBankLetter(ByVal obj As clsDBTNEFT, ByRef trans As SqlTransaction) As Boolean
        trans = clsDBFuncationality.GetTransactin()
        Try
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Bank_Letter_Date", clsCommon.GetPrintDate(obj.Bank_Letter_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT", OMInsertOrUpdate.Update, "TSPL_DBT_NEFT.Document_Code='" + obj.Document_Code + "'", trans)
            SaveHistoryData(obj.Document_Code, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsDBTNEFT
        Return GetData(strCode, NavType, Nothing)
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsDBTNEFT
        Dim obj As clsDBTNEFT = Nothing
        Dim Arr As List(Of clsDBTNEFT) = Nothing
        Dim qry As String = "Select TSPL_DBT_NEFT.* from TSPL_DBT_NEFT where 2=2 "
        Dim whrclas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_DBT_NEFT.Document_Code = (select MIN(Document_Code) from TSPL_DBT_NEFT WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Last
                qry += " and TSPL_DBT_NEFT.Document_Code = (select Max(Document_Code) from TSPL_DBT_NEFT WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Current
                qry += " and TSPL_DBT_NEFT.Document_Code = '" + strCode + "' "
            Case NavigatorType.Next
                qry += " and TSPL_DBT_NEFT.Document_Code = (select Min(Document_Code) from TSPL_DBT_NEFT where Document_Code>'" + strCode + "' " + whrclas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_DBT_NEFT.Document_Code = (select Max(Document_Code) from TSPL_DBT_NEFT where Document_Code<'" + strCode + "' " + whrclas + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsDBTNEFT()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Bank_Letter_Date = clsCommon.myCDate(dt.Rows(0)("Bank_Letter_Date"))
            obj.From_Date = clsCommon.myCDate(dt.Rows(0)("From_Date"))
            obj.To_Date = clsCommon.myCDate(dt.Rows(0)("To_Date"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.DBT_Revise_Payment = clsCommon.myCdbl(dt.Rows(0)("DBT_Revise_Payment"))
            obj.Zone_Code = clsCommon.myCstr(dt.Rows(0)("Zone_Code"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.RCDF_Status = IIf(clsCommon.myCdbl(dt.Rows(0)("RCDF_Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            If dt.Rows(0)("Posted_Date") IsNot DBNull.Value Then
                obj.Posted_Date = clsCommon.myCDate(dt.Rows(0)("Posted_Date"))
            End If
            obj.arrMCC = New ArrayList
            obj.arrVLC = New ArrayList
            obj.arr = clsDBTNEFTDetail.getData(obj.Document_Code, trans, obj.arrMCC, obj.arrVLC)
            obj.arrInvalid = clsDBTNEFTDetailInvalid.getData(obj.Document_Code, trans, obj.arrMCC, obj.arrVLC)
            obj.arrHold = clsDBTNEFTDetailHold.getData(obj.Document_Code, trans, obj.arrMCC, obj.arrVLC)
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

            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_DBT_NEFT.Document_Date,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code= TSPL_DBT_NEFT_DETAIL.Document_Code
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.pk_id= TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code= TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code

 where TSPL_DBT_NEFT_DETAIL.Document_Code ='" + strDocNo + "'", trans)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DBTNEFTUploader, clsCommon.myCstr(dt.Rows(0)("MCC_Code")), clsCommon.myCDate(dt.Rows(0)("Document_Date")), trans)

            End If
            clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_NEFT", "Document_Code", "TSPL_DBT_NEFT_DETAIL", "Document_Code", "TSPL_DBT_NEFT_DETAIL_INVALID", "Document_Code", trans)
            SaveHistoryData(strDocNo, trans)

            Dim qry As String = ""
            qry = "delete from TSPL_DBT_NEFT_DETAIL where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_NEFT_DETAIL_INVALID where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_NEFT_DETAIL_HOLD where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_NEFT where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = "Select TSPL_DBT_NEFT.Document_Code as Code,Convert(varchar,TSPL_DBT_NEFT.Document_Date,103) as Date
          ,TSPL_DBT_NEFT.Remarks as [Remarks],Convert(varchar,TSPL_DBT_NEFT.From_Date,103) as [From Date],Convert(varchar,TSPL_DBT_NEFT.To_Date,103) as [To Date],Zone_code as Zone
          ,case when isnull(Status,0)=0 then 'Pending' else 'Approved' end as Status,case when isnull(RCDF_Status,0)=0 then 'Pending' else 'Approved' end as [RCDF Status]
          from TSPL_DBT_NEFT "
        str = clsCommon.ShowSelectForm("DPTNeft#F", qry, "Code", whrcls, curcode, "Code", isButtonClicked, "Document_Date")
        Return str
    End Function
    Public Shared Function PostData(ByVal strDocNo As String, ByVal excelPath As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            PostData(strDocNo, excelPath, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
        Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostData(ByVal strDocNo As String, ByVal excelPath As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_DBT_NEFT.Document_Date,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code= TSPL_DBT_NEFT_DETAIL.Document_Code
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.pk_id= TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code= TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
 where TSPL_DBT_NEFT_DETAIL.Document_Code ='" + strDocNo + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DBTNEFTUploader, clsCommon.myCstr(dt.Rows(0)("MCC_Code")), clsCommon.myCDate(dt.Rows(0)("Document_Date")), trans)
            End If
            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")
            Dim obj As clsDBTNEFT = clsDBTNEFT.GetData(strDocNo, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            'SaveHistoryData(obj.Document_Code, trans)

            If (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Already Post on :" + obj.Posted_Date)
            End If
            Dim qry As String = "Update TSPL_DBT_NEFT set Status=1,RCDF_Status=0, Posted_Date='" + strPostDate + "',Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_NEFT", "Document_Code", trans)

            SetLotNo(obj.Document_Code, trans)



            '================================
            CreateEmailContent(obj, strDocNo, excelPath, trans)
            '================================
            Dim flag As Boolean = False
            Try
                qry = "select 1 from TSPL_MASTER.dbo.TSPL_APP_LOCATION where  code not in ('7544','7545') and DataBase_Name in ('" + objCommonVar.CurrDatabase + "') "
                dt = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    flag = True
                End If
            Catch ex As Exception
            End Try


            If flag Then
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "DB_Name", objCommonVar.CurrDatabase)
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
                clsCommon.AddColumnsForChange(coll, "From_Date", clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy"))
                clsCommon.AddColumnsForChange(coll, "To_Date", clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy"))
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommonFunctionality.UpdateDataTable(coll, objCommonVar.RCDFDB + "TSPL_DBT_NEFT_RCDF", OMInsertOrUpdate.Insert, "", trans)

                Dim strPKID As String = clsDBFuncationality.getSingleValue("select max(PK_ID) as PK_ID from " + objCommonVar.RCDFDB + "TSPL_DBT_NEFT_RCDF", trans)

                qry = "insert into " + objCommonVar.RCDFDB + "TSPL_ATTACHMENTS (Code,FormId,TransactionId,SNo,FileName,FileData,COMMENTS,Created_By,Created_Date,Modified_By,Modified_Date)
select '" + strPKID + "'+CODE as Code,'" + clsUserMgtCode.DBTPayment + "' as FormId,'" + strPKID + "' as TransactionId,SNo,FileName,FileData,COMMENTS,'" + objCommonVar.CurrentUserCode + "' as Created_By,GETDATE() as Created_Date,'" + objCommonVar.CurrentUserCode + "' as Modified_By,GETDATE() as  Modified_Date from TSPL_ATTACHMENTS where TransactionId='" + obj.Document_Code + "' and 2=(case when FormId='DBT-NEFT-UPL' then (case when Created_By='" + objCommonVar.CurrentUserCode + "' then 2 else 3 end ) else 2 end )"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Shared Function SetLotNo(document_Code As String, trans As SqlTransaction) As Boolean
        Dim intLotNo As Integer = 9
        Dim intNoOfRecordForLotNo As Integer = clsFixedParameter.GetData(clsFixedParameterType.PDAccountPaymanager, clsFixedParameterCode.NoOfRecordForLotNo, trans)
        Dim MaxPKID As String = ""
        Dim dt As DataTable = GetTRData(document_Code, MaxPKID, intNoOfRecordForLotNo, trans)
        While (dt IsNot Nothing AndAlso dt.Rows.Count > 0)
            intLotNo += 1
            If intLotNo > 99 Then
                Throw New Exception("Invalid Lot No [" + clsCommon.myCstr(intLotNo) + "]")
            End If
            Dim qry As String = "Update TSPL_DBT_NEFT_DETAIL set Lot_No ='" + clsCommon.myCstr(intLotNo) + "' where document_Code='" + document_Code + "' 
and TSPL_DBT_NEFT_DETAIL.PK_Id>='" + clsCommon.myCstr(dt.Rows(0)("PK_Id")) + "' and TSPL_DBT_NEFT_DETAIL.PK_Id<='" + clsCommon.myCstr(dt.Rows(dt.Rows.Count - 1)("PK_Id")) + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            MaxPKID = clsCommon.myCstr(dt.Rows(dt.Rows.Count - 1)("PK_Id"))
            dt = GetTRData(document_Code, MaxPKID, intNoOfRecordForLotNo, trans)
        End While
        Return True
    End Function

    Private Shared Function GetTRData(document_Code As String, ByVal maxPKID As String, ByVal NoOFRecords As String, trans As SqlTransaction) As DataTable
        Dim qry As String = "select top " & NoOFRecords & "  TSPL_DBT_NEFT_DETAIL.PK_Id from TSPL_DBT_NEFT_DETAIL 
where  TSPL_DBT_NEFT_DETAIL.Document_Code = '" + document_Code + "'  "
        If clsCommon.myLen(maxPKID) > 0 Then
            qry += " and  TSPL_DBT_NEFT_DETAIL.PK_Id > '" & maxPKID & "'"
        End If
        qry += " order by TSPL_DBT_NEFT_DETAIL.PK_Id"
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function

    Shared Sub CreateEmailContent(ByVal obj As clsDBTNEFT, ByVal strDocNo As String, ByVal strExcelPath As String, ByVal trans As SqlTransaction)

        Dim Form_ID As String = clsUserMgtCode.DBTNEFTUploader
        If clsCommon.myLen(strDocNo) > 0 Then
            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + Form_ID + "'", trans)
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
                Dim qry As String = "select Email from TSPL_BANK_MASTER where NEFT_DBT_Default = 1 "
                Dim bankEmail As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                If clsCommon.myLen(bankEmail) > 0 Then
                    Dim objSMSH As New clsEMailHead()
                    objSMSH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                    objSMSH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                    objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_No, obj.Document_Code)
                    objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                    objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.PaymentCycleFromDate, clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy"))
                    objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.PaymentCycleToDate, clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy"))
                    'objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Code, obj.Customer_Code)
                    'objSMSH.Email_Text = objSMSH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Name, obj.Customer_Name)
                    objSMSH.Attachment_1_Path = strExcelPath
                    objSMSH.arrEMail = New List(Of String)()
                    objSMSH.arrEMail.Add(clsCommon.myCstr(bankEmail))
                    objSMSH.SaveData(Form_ID, objSMSH, trans)
                End If
            End If
        End If

    End Sub



    Public Shared Function ReverseDataRCDF(ByVal strPKID As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim qry As String = "select DB_Name,Document_Code,ISNULL(status,0) as status,DB_Name from TSPL_DBT_NEFT_RCDF where PK_Id = " + strPKID + ""
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                qry = "select top 1  1 from " & clsCommon.myCstr(dt.Rows(0)("DB_Name")) & ".dbo.TSPL_DBT_NEFT_DETAIL 
inner join " & clsCommon.myCstr(dt.Rows(0)("DB_Name")) & ".dbo.TSPL_DBT_NEFT_BANK_RESPONSE on TSPL_DBT_NEFT_BANK_RESPONSE.Ref_PK_Id= TSPL_DBT_NEFT_DETAIL.PK_Id
where TSPL_DBT_NEFT_DETAIL.Document_Code='" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "'"
                Dim dtBR As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dtBR IsNot Nothing AndAlso dtBR.Rows.Count > 0 Then
                    Throw New Exception("Bank response received against the document [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "]")
                End If

                If clsCommon.myCDecimal(dt.Rows(0)("status")) = 0 Then
                    Throw New Exception("Already reverse Doument [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "]")
                End If

                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strPKID, "TSPL_DBT_NEFT_RCDF", "PK_Id", "", "", trans)

                qry = "delete from TSPL_DBT_NEFT_RCDF where  PK_Id=" + strPKID + ""
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from " + clsCommon.myCstr(dt.Rows(0)("DB_Name")) + ".dbo.TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL where  Document_Code in ('" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "')"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "RCDF_Status", 0)
                clsCommonFunctionality.UpdateDataTable(coll, clsCommon.myCstr(dt.Rows(0)("DB_Name")) + ".dbo." + "TSPL_DBT_NEFT", OMInsertOrUpdate.Update, "Document_Code='" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "'", trans)
            End If

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Function funPrintBankLetter(ByVal strDocNo As String, ByVal isPDFPath As Boolean) As String
        Dim strPAth As String = ""
        Try
            CycleWiseDBT = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CycleWiseDBT, clsFixedParameterCode.CycleWiseDBT, Nothing)) = 1, True, False)
            If clsCommon.myLen(strDocNo) <= 0 Then
                Throw New Exception("Doument not found to Print")
            End If
            Dim reportDateTime As String = (clsDBFuncationality.getSingleValue("select convert(varchar , getdate(),113) as Date_Time"))
            Dim isPending As String = ERPTransactionStatus.Pending
            Dim status As String = ""
            Dim qry As String = "select  Document_Date,Bank_Letter_Date,To_Date,From_Date,TSPL_DBT_NEFT_DETAIL.Amount as Total_Amt,TSPL_DBT_NEFT_DETAIL.DCS,TSPL_DBT_NEFT_DETAIL.[Farmer Code],TSPL_DBT_NEFT_DETAIL.Total_Milk,TSPL_DBT_NEFT_DETAIL.Rate 
from TSPL_DBT_NEFT 
left outer join (select TSPL_DBT_NEFT_DETAIL.document_code,max(TSPL_DBT_NEFT_DETAIL.Rem_Name) as Rem_Name,max(TSPL_DBT_NEFT_DETAIL.Rem_Account_No) as Rem_Account_No,
sum(TSPL_DBT_NEFT_DETAIL.Amount) as Amount , COUNT( DISTINCT TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code) AS DCS, Max(TSPL_MP_INCENTIVE_ENTRY_HEAD.Incetive_Rate)as Rate ,
COUNT( DISTINCT TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code)  as [Farmer Code],sum(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty) as Total_Milk from TSPL_DBT_NEFT_DETAIL   
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
where   TSPL_DBT_NEFT_DETAIL.document_code ='" + strDocNo + "' group by TSPL_DBT_NEFT_DETAIL.document_code 
) TSPL_DBT_NEFT_DETAIL   on TSPL_DBT_NEFT.Document_Code = TSPL_DBT_NEFT_DETAIL.Document_Code  
where TSPL_DBT_NEFT.Document_Code = '" + strDocNo + "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("Invalid Document No [" + strDocNo + "]")
            End If
            Dim Doc_Date As String = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Bank_Letter_Date")), "dd/MM/yyyy")
            Dim ToDate As String = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("To_Date")), "dd/MM/yyyy")
            Dim FromDate As String = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("From_Date")), "dd/MM/yyyy")
            Dim strTotal_Amt_Format As String = clsCommon.myFormat(clsCommon.myCDecimal(dt.Rows(0)("Total_Amt")))
            Dim strTotal_Amt As String = clsCommon.myCstr(clsCommon.myCDecimal(dt.Rows(0)("Total_Amt")))
            Dim strDCS As String = clsCommon.myFormat(clsCommon.myCDecimal(dt.Rows(0)("DCS")), False, False, True, 0, False)
            Dim strFarmerCode As String = clsCommon.myFormat(clsCommon.myCDecimal(dt.Rows(0)("Farmer Code")), False, False, True, 0, False)
            Dim strTotal_Milk As String = clsCommon.myFormat(clsCommon.myCDecimal(dt.Rows(0)("Total_Milk")))
            Dim strRate As String = clsCommon.myFormat(clsCommon.myCDecimal(dt.Rows(0)("Rate")))

            If isPending = "1" Then
                status = "Pending"
            End If
            Dim User_Name As String = objCommonVar.CurrentUser
            Try
                If clsCommon.myLen(strDocNo) <= 0 Then
                    Throw New Exception("Please select Document No")
                End If

                qry = "select case when DBT_Revise_Payment = 1 then 'DCS filling data in saras pro app but payment not paid' else '' end as DBTRevisePayment, ( select TSPL_USER_MASTER.User_Name from TSPL_APPROVAL_LEVEL_SCREEN LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_USER_MASTER.User_Code = TSPL_APPROVAL_LEVEL_SCREEN.User_Code where Module_Code = 'MMMProc' and TRANS_Code = 'DBT-NEFT-UPL' and No_Of_Level = 1) as 'P&IIncharge'
                ,(select TSPL_USER_MASTER.User_Name from TSPL_APPROVAL_LEVEL_SCREEN LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_USER_MASTER.User_Code = TSPL_APPROVAL_LEVEL_SCREEN.User_Code where Module_Code = 'MMMProc' and TRANS_Code = 'DBT-NEFT-UPL' and No_Of_Level = 2) as 'AccountHead',
              (select TSPL_USER_MASTER.User_Name from TSPL_APPROVAL_LEVEL_SCREEN  LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_USER_MASTER.User_Code = TSPL_APPROVAL_LEVEL_SCREEN.User_Code where Module_Code = 'MMMProc' and TRANS_Code = 'DBT-NEFT-UPL' and No_Of_Level = 3) as 'ManagingDirector', tspl_company_master.Logo_Img , tspl_company_master.Logo_Img2 , tspl_company_master.Comp_Name , tspl_company_master.Add1 , tspl_company_master.Add2, tspl_company_master.Add3,
tspl_company_master.City_Code, tspl_company_master.Pincode, tspl_company_master.Email,REPLACE( RIGHT( Phone1,11),'_','') as Phone1 ,TSPL_DBT_NEFT.Document_Code as Doc_No ,"

                If CycleWiseDBT Then
                    qry += " '" + clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") + "'  + ' To ' + '" + clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") + "'  as PayCycle,"
                Else
                    qry += "DATENAME(MONTH, CONVERT(date, '" & ToDate & "', 103)) + '-' + DATENAME(YEAR, CONVERT(date, '" & ToDate & "', 103)) as PayCycle,"
                End If
                qry += " '" & Doc_Date & "' as Date, '" & reportDateTime & "' as Date_Time , '" & status & "' as Pending , '" & User_Name & "' as User_Name
," + strRate + " AS Rate,'" + strFarmerCode + "' as  [Farmer Code],'" + strDCS + "' as DCS,'" + strTotal_Milk + "' as Total_Milk," + strTotal_Amt + " as Total_Amt, '" + strTotal_Amt_Format + "' as Total_Amt_Format ,'" & FromDate & "' as From_Date, '" & ToDate & "' as To_Date, TSPL_DBT_NEFT.To_Date
,TSPL_BANK_BRANCH_MASTER.BRANCH_NAME,TSPL_BANK_MASTER.ADD1 as BankADD1,TSPL_BANK_MASTER.ADD2 as BankADD2,TSPL_BANK_MASTER.ADD3 as BankADD3,TSPL_BANK_MASTER.ADD4 as BankADD4,TSPL_BANK_MASTER.BANKACCNUMBER,TSPL_BANK_MASTER.DESCRIPTION as BankName
from TSPL_DBT_NEFT  
LEFT OUTER JOIN tspl_company_master ON tspl_company_master.comp_code = '" + objCommonVar.CurrentCompanyCode + "'
left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.NEFT_DBT_Default=1
left outer join TSPL_BANK_BRANCH_MASTER on TSPL_BANK_BRANCH_MASTER.Bank_CODE=TSPL_BANK_MASTER.BANK_CODE
WHERE  TSPL_DBT_NEFT.document_code ='" + strDocNo + "'"
                dt = clsDBFuncationality.GetDataTable(qry)
                Dim Apply_PD_Account_Date As Date = clsDBFuncationality.getSingleValue("SELECT [TSPL_APP_LOCATION].Apply_PD_Account_Date FROM [TSPL_MASTER].[dbo].[TSPL_APP_LOCATION] WHERE Union_Report=1 and DataBase_Name IN ('" + objCommonVar.CurrDatabase + "') ORDER BY [TSPL_APP_LOCATION].Location_Name")

                Dim frmCRV As New frmCrystalReportViewer()
                If clsCommon.myCDate(ToDate) >= clsCommon.myCDate(Apply_PD_Account_Date) Then
                    strPAth = frmCRV.funreport("ABC", isPDFPath, CrystalReportFolder.MilkProcurement, dt, "crptDBTNEFTUploaderBankLetter", "Bank Letter NEFT Uploader")
                Else
                    strPAth = frmCRV.funreport("ABC", isPDFPath, CrystalReportFolder.MilkProcurement, dt, "crptDBTNEFTUploaderBankLetterYesBank", "Bank Letter NEFT Uploader")
                End If
                frmCRV = Nothing
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        Catch ex As Exception
            If isPDFPath Then
                Throw New Exception("Errow While making attachment" + ex.Message)
            Else
                clsCommon.MyMessageBoxShow(ex.Message)
            End If
        End Try
        Return strPAth
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
            Dim obj As clsDBTNEFT = clsDBTNEFT.GetData(strCode, NavigatorType.Current, trans)
            If obj.RCDF_Status = ERPTransactionStatus.Approved Then
                Throw New Exception("Transaction Approved by RCDF can't reverse it.")
            End If
            If Not obj.Status = ERPTransactionStatus.Approved Then
                Throw New Exception("Transaction should be Approved to reverse.")
            End If
            Dim flag As Boolean = False
            Dim qry As String
            Try
                qry = "select 1 from TSPL_MASTER.dbo.TSPL_APP_LOCATION where  code not in ('5888','6888') and DataBase_Name in ('" + objCommonVar.CurrDatabase + "') "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    flag = True
                End If
            Catch ex As Exception
            End Try

            If flag Then
                qry = "delete from " + objCommonVar.RCDFDB + "TSPL_DBT_NEFT_RCDF where DB_Name='" + objCommonVar.CurrDatabase + "' and Document_Code='" + obj.Document_Code + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

            qry = "delete from TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL where Document_Code in ('" + obj.Document_Code + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_ATTACHMENTS where TransactionId in ('" + obj.Document_Code + "') and FormId='" + clsUserMgtCode.DBTNEFTUploader + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "Update TSPL_DBT_NEFT set Status=0, Posted_Date=null,Posted_By=null where Document_Code='" + obj.Document_Code + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function ClearApprovalLevel(ByVal strCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ClearApprovalLevel(strCode, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function ClearApprovalLevel(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String

            Dim obj As clsDBTNEFT = clsDBTNEFT.GetData(strCode, NavigatorType.Current, trans)
            If obj.RCDF_Status = ERPTransactionStatus.Approved Then
                Throw New Exception("Transaction Approved by RCDF can't reverse it.")
            End If
            If Not obj.Status = ERPTransactionStatus.Approved Then
                qry = "delete from TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL where Document_Code in ('" + obj.Document_Code + "')"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_ATTACHMENTS where TransactionId in ('" + obj.Document_Code + "') and FormId='" + clsUserMgtCode.DBTNEFTUploader + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsDBTNEFTDetail
#Region "Variable"
    Public PK_Id As Integer
    Public Document_Code As String = Nothing
    Public Against_MP_Incentive_TR As Integer
    Public SNo As Integer
    Public Rem_Account_No As String = Nothing
    Public Rem_Name As String = Nothing
    Public VLC_Uploader_Code As String = Nothing
    Public MP_Uploader_Code As String = Nothing
    Public Amount As Decimal = 0
    Public MP_IFSC_No As String = Nothing
    Public MP_Account_No As String = Nothing
    Public MP_Bank As String = Nothing
    Public MP_Mobile_No As String = Nothing
    Public MP_Name As String = Nothing
    Public Transaction As String = Nothing
#End Region
    Public Shared Function saveData(ByVal arrObj As List(Of clsDBTNEFTDetail), ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim coll As Hashtable
            If arrObj IsNot Nothing Then
                For Each obj As clsDBTNEFTDetail In arrObj
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Against_MP_Incentive_TR", obj.Against_MP_Incentive_TR)
                    clsCommon.AddColumnsForChange(coll, "SNo", obj.SNo)
                    clsCommon.AddColumnsForChange(coll, "Rem_Account_No", obj.Rem_Account_No)
                    clsCommon.AddColumnsForChange(coll, "Rem_Name", obj.Rem_Name)
                    clsCommon.AddColumnsForChange(coll, "VLC_Uploader_Code", obj.VLC_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "MP_Uploader_Code", obj.MP_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                    clsCommon.AddColumnsForChange(coll, "MP_Bank", obj.MP_Bank)
                    clsCommon.AddColumnsForChange(coll, "MP_Mobile_No", obj.MP_Mobile_No)
                    clsCommon.AddColumnsForChange(coll, "MP_IFSC_No", obj.MP_IFSC_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Account_No", obj.MP_Account_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Name", obj.MP_Name)
                    clsCommon.AddColumnsForChange(coll, "[Transaction]", obj.Transaction)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getData(ByVal strDocNo As String, ByVal trans As SqlTransaction, ByRef ArrMCC As ArrayList, ByRef ArrVLC As ArrayList) As List(Of clsDBTNEFTDetail)
        Try

            Dim arrObj As List(Of clsDBTNEFTDetail) = Nothing
            Dim obj As clsDBTNEFTDetail = Nothing
            Dim qry As String = "Select TSPL_DBT_NEFT_DETAIL.*,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
from TSPL_DBT_NEFT_DETAIL 
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
where TSPL_DBT_NEFT_DETAIL.Document_Code='" & strDocNo & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of clsDBTNEFTDetail)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsDBTNEFTDetail()
                    obj.PK_Id = clsCommon.myCdbl(dt.Rows(i)("PK_Id"))
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.Against_MP_Incentive_TR = clsCommon.myCdbl(dt.Rows(i)("Against_MP_Incentive_TR"))
                    obj.SNo = i + 1
                    obj.Rem_Account_No = clsCommon.myCstr(dt.Rows(i)("Rem_Account_No"))
                    obj.Rem_Name = clsCommon.myCstr(dt.Rows(i)("Rem_Name"))
                    obj.VLC_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Uploader_Code"))
                    obj.MP_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("MP_Uploader_Code"))
                    obj.Amount = clsCommon.myCdbl(dt.Rows(i)("Amount"))
                    obj.MP_IFSC_No = clsCommon.myCstr(dt.Rows(i)("MP_IFSC_No"))
                    obj.MP_Account_No = clsCommon.myCstr(dt.Rows(i)("MP_Account_No"))
                    obj.MP_Bank = clsCommon.myCstr(dt.Rows(i)("MP_Bank"))
                    obj.MP_Mobile_No = clsCommon.myCstr(dt.Rows(i)("MP_Mobile_No"))
                    obj.MP_Name = clsCommon.myCstr(dt.Rows(i)("MP_Name"))
                    obj.Transaction = clsCommon.myCstr(dt.Rows(i)("Transaction"))
                    arrObj.Add(obj)
                    If Not ArrMCC.Contains(clsCommon.myCstr(dt.Rows(i)("MCC_Code"))) Then
                        ArrMCC.Add(clsCommon.myCstr(dt.Rows(i)("MCC_Code")))
                    End If
                    If Not ArrVLC.Contains(clsCommon.myCstr(dt.Rows(i)("VLC_Code"))) Then
                        ArrVLC.Add(clsCommon.myCstr(dt.Rows(i)("VLC_Code")))
                    End If
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsDBTNEFTDetailInvalid
#Region "Variable"
    Public PK_Id As Integer
    Public Document_Code As String = Nothing
    Public Against_MP_Incentive_TR As Integer
    Public SNo As Integer
    Public Rem_Account_No As String = Nothing
    Public Rem_Name As String = Nothing
    Public VLC_Uploader_Code As String = Nothing
    Public MP_Uploader_Code As String = Nothing
    Public Amount As Decimal = 0
    Public MP_IFSC_No As String = Nothing
    Public MP_Account_No As String = Nothing
    Public MP_Bank As String = Nothing
    Public MP_Mobile_No As String = Nothing
    Public MP_Name As String = Nothing
    Public Transaction As String = Nothing
#End Region
    Public Shared Function saveData(ByVal arrObj As List(Of clsDBTNEFTDetailInvalid), ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim coll As Hashtable
            If arrObj IsNot Nothing Then
                For Each obj As clsDBTNEFTDetailInvalid In arrObj
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Against_MP_Incentive_TR", obj.Against_MP_Incentive_TR)
                    clsCommon.AddColumnsForChange(coll, "SNo", obj.SNo)
                    clsCommon.AddColumnsForChange(coll, "Rem_Account_No", obj.Rem_Account_No)
                    clsCommon.AddColumnsForChange(coll, "Rem_Name", obj.Rem_Name)
                    clsCommon.AddColumnsForChange(coll, "VLC_Uploader_Code", obj.VLC_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "MP_Uploader_Code", obj.MP_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                    clsCommon.AddColumnsForChange(coll, "MP_Bank", obj.MP_Bank)
                    clsCommon.AddColumnsForChange(coll, "MP_Mobile_No", obj.MP_Mobile_No)
                    clsCommon.AddColumnsForChange(coll, "MP_IFSC_No", obj.MP_IFSC_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Account_No", obj.MP_Account_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Name", obj.MP_Name)
                    clsCommon.AddColumnsForChange(coll, "[Transaction]", obj.Transaction)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT_DETAIL_INVALID", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getData(ByVal strDocNo As String, ByVal trans As SqlTransaction, ByRef ArrMCC As ArrayList, ByRef ArrVLC As ArrayList) As List(Of clsDBTNEFTDetailInvalid)
        Try

            Dim arrObj As List(Of clsDBTNEFTDetailInvalid) = Nothing
            Dim obj As clsDBTNEFTDetailInvalid = Nothing
            Dim qry As String = "Select TSPL_DBT_NEFT_DETAIL_INVALID.*,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
from TSPL_DBT_NEFT_DETAIL_INVALID 
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL_INVALID.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
where TSPL_DBT_NEFT_DETAIL_INVALID.Document_Code='" & strDocNo & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of clsDBTNEFTDetailInvalid)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsDBTNEFTDetailInvalid()
                    obj.PK_Id = clsCommon.myCdbl(dt.Rows(i)("PK_Id"))
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.Against_MP_Incentive_TR = clsCommon.myCdbl(dt.Rows(i)("Against_MP_Incentive_TR"))
                    obj.SNo = i + 1
                    obj.Rem_Account_No = clsCommon.myCstr(dt.Rows(i)("Rem_Account_No"))
                    obj.Rem_Name = clsCommon.myCstr(dt.Rows(i)("Rem_Name"))
                    obj.VLC_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Uploader_Code"))
                    obj.MP_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("MP_Uploader_Code"))
                    obj.Amount = clsCommon.myCdbl(dt.Rows(i)("Amount"))
                    obj.MP_IFSC_No = clsCommon.myCstr(dt.Rows(i)("MP_IFSC_No"))
                    obj.MP_Account_No = clsCommon.myCstr(dt.Rows(i)("MP_Account_No"))
                    obj.MP_Bank = clsCommon.myCstr(dt.Rows(i)("MP_Bank"))
                    obj.MP_Mobile_No = clsCommon.myCstr(dt.Rows(i)("MP_Mobile_No"))
                    obj.MP_Name = clsCommon.myCstr(dt.Rows(i)("MP_Name"))
                    obj.Transaction = clsCommon.myCstr(dt.Rows(i)("Transaction"))
                    arrObj.Add(obj)
                    If Not ArrMCC.Contains(clsCommon.myCstr(dt.Rows(i)("MCC_Code"))) Then
                        ArrMCC.Add(clsCommon.myCstr(dt.Rows(i)("MCC_Code")))
                    End If
                    If Not ArrVLC.Contains(clsCommon.myCstr(dt.Rows(i)("VLC_Code"))) Then
                        ArrVLC.Add(clsCommon.myCstr(dt.Rows(i)("VLC_Code")))
                    End If
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsDBTNEFTPerforma
    Public Const colSlNo As String = "SNo"
    'Public Const colUniqueNo As String = "Unique No"
    Public Const colAgainstMPIncetive As String = "DBT TR No"
    Public Const colFarmerCode As String = "Farmer Code"
    Public Const colSociety As String = "Society"
    Public Const colSocietyName As String = "Society Name"
    Public Const colZoneName As String = "Zone Name"
    Public Const colMPUploaderCode As String = "MP Uploader Code"
    Public Const colAmount As String = "AMOUNT"
    Public Const colMPIFSCCode As String = "IFSC CODE"
    Public Const colMPAccountNo As String = "BENEFICERY ACCOUNT  NO."
    Public Const colMPName As String = "BENEFICERY NAME"
    Public Const colMPBank As String = "Bank"
    Public Const colMPMobileNo As String = "Mobile No"
    Public Const colTHIRDPARTYCODE As String = "Third Party Code"



    Public Shared Function GetDefault() As DataTable
        Dim dt As New DataTable
        Try
            dt.Columns.Add("Code", GetType(String))

            Dim dr As DataRow = dt.NewRow()
            dr("Code") = colSlNo
            dt.Rows.Add(dr)


            dr = dt.NewRow()
            dr("Code") = colAgainstMPIncetive
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colFarmerCode
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colZoneName
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colSociety
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colSocietyName
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPUploaderCode
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colAmount
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPIFSCCode
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPAccountNo
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPName
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPBank
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Code") = colMPMobileNo
            dt.Rows.Add(dr)


            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return dt
    End Function
End Class

Public Class clsDBTNEFTRCDF
    Public PK_Id As Integer
    Public Sanction_Number As String = Nothing
    Public Sanction_Date As Date? = Nothing
    Public Sanction_Amount As Decimal

    Public Shared Function PostDataRCDF(ByVal Arr As List(Of clsDBTNEFTRCDF)) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            For Each obj As clsDBTNEFTRCDF In Arr
                Dim qry As String = "select DB_Name,Document_Code,ISNULL(status,0) as status,DB_Name from TSPL_DBT_NEFT_RCDF where PK_Id = " + clsCommon.myCstr(obj.PK_Id) + ""
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If clsCommon.myCDecimal(dt.Rows(0)("status")) = 1 Then
                        Throw New Exception("Already posted Doument [" + clsCommon.myCDecimal(dt.Rows(0)("Document_Code")) + "]")
                    End If

                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Status", 1)
                    clsCommon.AddColumnsForChange(coll, "Post_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "Post_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                    clsCommon.AddColumnsForChange(coll, "Sanction_Number", obj.Sanction_Number)
                    If obj.Sanction_Date IsNot Nothing Then
                        clsCommon.AddColumnsForChange(coll, "Sanction_Date", clsCommon.GetPrintDate(obj.Sanction_Date, "dd/MMM/yyyy"))
                    End If
                    clsCommon.AddColumnsForChange(coll, "Sanction_Amount", obj.Sanction_Amount)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT_RCDF", OMInsertOrUpdate.Update, "PK_Id=" + clsCommon.myCstr(obj.PK_Id) + "", trans)

                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "RCDF_Status", 1)
                    clsCommon.AddColumnsForChange(coll, "Sanction_Number", obj.Sanction_Number)
                    If obj.Sanction_Date IsNot Nothing Then
                        clsCommon.AddColumnsForChange(coll, "Sanction_Date", clsCommon.GetPrintDate(obj.Sanction_Date, "dd/MMM/yyyy"))
                    End If
                    clsCommon.AddColumnsForChange(coll, "Sanction_Amount", obj.Sanction_Amount)
                    clsCommon.AddColumnsForChange(coll, "RCDF_Post_By", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "RCDF_Post_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))

                    clsCommonFunctionality.UpdateDataTable(coll, clsCommon.myCstr(dt.Rows(0)("DB_Name")) + ".dbo." + "TSPL_DBT_NEFT", OMInsertOrUpdate.Update, "Document_Code='" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "'", trans)
                End If
            Next
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

End Class

Public Class clsDBTNEFTDetailHold
#Region "Variable"
    Public PK_Id As Integer
    Public Document_Code As String = Nothing
    Public Against_MP_Incentive_TR As Integer
    Public SNo As Integer
    Public Rem_Account_No As String = Nothing
    Public Rem_Name As String = Nothing
    Public VLC_Uploader_Code As String = Nothing
    Public MP_Uploader_Code As String = Nothing
    Public Amount As Decimal = 0
    Public MP_IFSC_No As String = Nothing
    Public MP_Account_No As String = Nothing
    Public MP_Bank As String = Nothing
    Public MP_Mobile_No As String = Nothing
    Public MP_Name As String = Nothing
    Public Transaction As String = Nothing
#End Region
    Public Shared Function saveData(ByVal arrObj As List(Of clsDBTNEFTDetailHold), ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim coll As Hashtable
            If arrObj IsNot Nothing Then
                For Each obj As clsDBTNEFTDetailHold In arrObj
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Against_MP_Incentive_TR", obj.Against_MP_Incentive_TR)
                    clsCommon.AddColumnsForChange(coll, "SNo", obj.SNo)
                    clsCommon.AddColumnsForChange(coll, "Rem_Account_No", obj.Rem_Account_No)
                    clsCommon.AddColumnsForChange(coll, "Rem_Name", obj.Rem_Name)
                    clsCommon.AddColumnsForChange(coll, "VLC_Uploader_Code", obj.VLC_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "MP_Uploader_Code", obj.MP_Uploader_Code)
                    clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                    clsCommon.AddColumnsForChange(coll, "MP_Bank", obj.MP_Bank)
                    clsCommon.AddColumnsForChange(coll, "MP_Mobile_No", obj.MP_Mobile_No)
                    clsCommon.AddColumnsForChange(coll, "MP_IFSC_No", obj.MP_IFSC_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Account_No", obj.MP_Account_No)
                    clsCommon.AddColumnsForChange(coll, "MP_Name", obj.MP_Name)
                    clsCommon.AddColumnsForChange(coll, "[Transaction]", obj.Transaction)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_NEFT_DETAIL_HOLD", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getData(ByVal strDocNo As String, ByVal trans As SqlTransaction, ByRef ArrMCC As ArrayList, ByRef ArrVLC As ArrayList) As List(Of clsDBTNEFTDetailHold)
        Try

            Dim arrObj As List(Of clsDBTNEFTDetailHold) = Nothing
            Dim obj As clsDBTNEFTDetailHold = Nothing
            Dim qry As String = "Select TSPL_DBT_NEFT_DETAIL_HOLD.*,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
from TSPL_DBT_NEFT_DETAIL_HOLD 
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL_HOLD.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
where TSPL_DBT_NEFT_DETAIL_HOLD.Document_Code='" & strDocNo & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of clsDBTNEFTDetailHold)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsDBTNEFTDetailHold()
                    obj.PK_Id = clsCommon.myCdbl(dt.Rows(i)("PK_Id"))
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.Against_MP_Incentive_TR = clsCommon.myCdbl(dt.Rows(i)("Against_MP_Incentive_TR"))
                    obj.SNo = i + 1
                    obj.Rem_Account_No = clsCommon.myCstr(dt.Rows(i)("Rem_Account_No"))
                    obj.Rem_Name = clsCommon.myCstr(dt.Rows(i)("Rem_Name"))
                    obj.VLC_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Uploader_Code"))
                    obj.MP_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("MP_Uploader_Code"))
                    obj.Amount = clsCommon.myCdbl(dt.Rows(i)("Amount"))
                    obj.MP_IFSC_No = clsCommon.myCstr(dt.Rows(i)("MP_IFSC_No"))
                    obj.MP_Account_No = clsCommon.myCstr(dt.Rows(i)("MP_Account_No"))
                    obj.MP_Bank = clsCommon.myCstr(dt.Rows(i)("MP_Bank"))
                    obj.MP_Mobile_No = clsCommon.myCstr(dt.Rows(i)("MP_Mobile_No"))
                    obj.MP_Name = clsCommon.myCstr(dt.Rows(i)("MP_Name"))
                    obj.Transaction = clsCommon.myCstr(dt.Rows(i)("Transaction"))
                    arrObj.Add(obj)
                    If Not ArrMCC.Contains(clsCommon.myCstr(dt.Rows(i)("MCC_Code"))) Then
                        ArrMCC.Add(clsCommon.myCstr(dt.Rows(i)("MCC_Code")))
                    End If
                    If Not ArrVLC.Contains(clsCommon.myCstr(dt.Rows(i)("VLC_Code"))) Then
                        ArrVLC.Add(clsCommon.myCstr(dt.Rows(i)("VLC_Code")))
                    End If
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class