Imports System.Data.SqlClient

Public Class clsDBTMonthlyFarmerMilk
#Region "variables"
    Public Document_Code As String = Nothing
    Public Document_Date As Date
    Public DBT_Reco_Code As String = ""
    Public From_Date As Date ''Not a Table Column
    Public To_Date As Date ''Not a Table Column
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Posted_Date As Date? = Nothing
    Public arr As List(Of clsDBTMonthlyFarmerMilkDetail) = Nothing

#End Region
    Public Shared Function SaveData(ByVal obj As clsDBTMonthlyFarmerMilk, ByVal isNewEntry As Boolean) As Boolean
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

    Public Shared Function SaveData(ByVal obj As clsDBTMonthlyFarmerMilk, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        Try
            qry = "delete from TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL where Document_Code='" & obj.Document_Code & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "DBT_Reco_Code", obj.DBT_Reco_Code)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            If isNewEntry Then
                obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.DBTMonthlyFarmerMilk, "", "")
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_MONTHLY_FARMER_MILK", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_MONTHLY_FARMER_MILK", OMInsertOrUpdate.Update, "TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            clsDBTMonthlyFarmerMilkDetail.SaveData(obj.From_Date, obj.arr, obj.Document_Code, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DBT_MONTHLY_FARMER_MILK", "Document_Code", "TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL", "Document_Code", trans)
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsDBTMonthlyFarmerMilk
        Return GetData(strCode, NavType, Nothing)
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsDBTMonthlyFarmerMilk
        Return GetData(strCode, NavType, trans, True)
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction, ByVal PickDetailsData As Boolean) As clsDBTMonthlyFarmerMilk
        Dim obj As clsDBTMonthlyFarmerMilk = Nothing
        Dim Arr As List(Of clsDBTMonthlyFarmerMilk) = Nothing
        Dim qry As String = "Select TSPL_DBT_MONTHLY_FARMER_MILK.*,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To from TSPL_DBT_MONTHLY_FARMER_MILK " &
            " Left Outer Join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=TSPL_DBT_MONTHLY_FARMER_MILK.DBT_Reco_Code  where 2=2 "
        Dim whrclas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code = (select MIN(Document_Code) from TSPL_DBT_MONTHLY_FARMER_MILK WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Last
                qry += " and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code = (select Max(Document_Code) from TSPL_DBT_MONTHLY_FARMER_MILK WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Current
                qry += " and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code = '" + strCode + "' "
            Case NavigatorType.Next
                qry += " and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code = (select Min(Document_Code) from TSPL_DBT_MONTHLY_FARMER_MILK where Document_Code>'" + strCode + "' " + whrclas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code = (select Max(Document_Code) from TSPL_DBT_MONTHLY_FARMER_MILK where Document_Code<'" + strCode + "' " + whrclas + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsDBTMonthlyFarmerMilk()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.DBT_Reco_Code = clsCommon.myCstr(dt.Rows(0)("DBT_Reco_Code"))
            obj.From_Date = clsCommon.myCDate(dt.Rows(0)("Reco_Date"))
            obj.To_Date = clsCommon.myCDate(dt.Rows(0)("Reco_Date_To"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            If dt.Rows(0)("Posted_Date") IsNot DBNull.Value Then
                obj.Posted_Date = clsCommon.myCDate(dt.Rows(0)("Posted_Date"))
            End If
            If PickDetailsData Then
                obj.arr = clsDBTMonthlyFarmerMilkDetail.getData(obj.Document_Code, trans)
            End If
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
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_MONTHLY_FARMER_MILK", "Document_Code", "TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL", "Document_Code", trans)
            Dim qry As String = "delete from TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DBT_MONTHLY_FARMER_MILK where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = "Select TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code as Code,Convert(varchar,TSPL_DBT_MONTHLY_FARMER_MILK.Document_Date,103) as Date
,TSPL_DBT_MONTHLY_FARMER_MILK.DBT_Reco_Code as [DBT Reco Code] ,Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) as [From Date],Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To,103) as [To Date]
,case when isnull(TSPL_DBT_MONTHLY_FARMER_MILK.Status,0)=0 then 'Pending' else 'Approved' end as Status 
from TSPL_DBT_MONTHLY_FARMER_MILK 
Left Outer Join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=TSPL_DBT_MONTHLY_FARMER_MILK.DBT_Reco_Code "
        str = clsCommon.ShowSelectForm("DBTMFM#F", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If
            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")
            Dim obj As clsDBTMonthlyFarmerMilk = clsDBTMonthlyFarmerMilk.GetData(strDocNo, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Already Post on :" + obj.Posted_Date)
            End If

            'Dim Arr As New List(Of String, clsMPIncentiveEntry)
            Dim SettMCCOneDBTOneDoc As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.OneDBTOneDoc, trans))
            Dim SettMPIncentiveEntryIncentiveRate As Decimal = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryIncentiveRate, clsFixedParameterCode.MPIncentiveEntryIncentiveRate, trans))
            If clsCommon.myLen(SettMCCOneDBTOneDoc) <= 0 Then
                SettMCCOneDBTOneDoc = obj.arr(0).MCC_Code
            End If
            Dim qry As String = "select Uom_Code from TSPL_Mcc_UOM_DETAIL where Stocking_Unit='Y' and  MCC_CODE='" + SettMCCOneDBTOneDoc + "'"
            Dim UOM As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))

            Dim objMPE As New clsMPIncentiveEntry()
            objMPE.Document_Date = obj.Document_Date
            objMPE.MCC_Code = SettMCCOneDBTOneDoc
            objMPE.From_Date = obj.From_Date
            objMPE.To_Date = obj.To_Date
            objMPE.Incetive_Rate = SettMPIncentiveEntryIncentiveRate
            objMPE.FATSNFPer = 2
            Dim objMPETr As New clsMPIncentiveEntryDetail
            objMPE.arr = New List(Of clsMPIncentiveEntryDetail)
            For Each objMFMD In obj.arr
                objMPETr = New clsMPIncentiveEntryDetail()
                objMPETr.SNo = objMPE.arr.Count + 1
                objMPETr.VLC_Code = objMFMD.VLC_Code
                objMPETr.MP_Code = objMFMD.MP_Code
                'objMPETr.MP_Account_No = clsCommon.myCstr(grow.Cells(clsMPIncetiveEntryColumns.colMPAccountNo).Value)
                'objMPETr.MP_Bank = clsCommon.myCstr(grow.Cells(clsMPIncetiveEntryColumns.colMPBank).Value)
                objMPETr.Qty = objMFMD.Qty
                objMPETr.UOM = UOM
                objMPETr.Amount = objMFMD.Qty * SettMPIncentiveEntryIncentiveRate
                objMPETr.Amount_Actual = objMPETr.Amount

                'objMPETr.MP_Phone_No = clsCommon.myCstr(grow.Cells(clsMPIncetiveEntryColumns.colMPPhoneNo).Value)
                'objMPETr.MP_Aadhar_No = clsCommon.myCstr(grow.Cells(clsMPIncetiveEntryColumns.colMPAadharNo).Value)
                'objMPETr.MP_IFSC_No = clsCommon.myCstr(grow.Cells(clsMPIncetiveEntryColumns.colMPIFSCCode).Value)

                'objMPETr.FAT = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colFAT).Value)
                'objMPETr.SNF = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colSNF).Value)

                'objMPETr.Pashu_Aahar_Qty = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colPashuAaharQty).Value)
                'objMPETr.Pashu_Aahar_Amount = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colPashuAaharAmt).Value)
                'objMPETr.Mineral_Mixture_Qty = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colMineralMixtureQty).Value)
                'objMPETr.Mineral_Mixture_Amount = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colMineralMixtureAmt).Value)
                'objMPETr.Sailej_Qty = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colSailejQty).Value)
                'objMPETr.Sailej_Amount = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colSailejAmt).Value)
                'objMPETr.Rahat_Kampekat_Feed_Qty = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colRahatKampekatFeedQty).Value)
                'objMPETr.Rahat_Kampekat_Feed_Amount = clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colRahatKampekatFeedAmt).Value)
                objMPETr.Total_Amount = objMPETr.Amount
                objMPE.arr.Add(objMPETr)
            Next

            clsMPIncentiveEntry.SaveData(objMPE, True, trans)

            qry = "Update TSPL_DBT_MONTHLY_FARMER_MILK set Status=1, Posted_Date='" + strPostDate + "',Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DBT_MONTHLY_FARMER_MILK", "Document_Code", "TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL", "Document_Code", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsDBTMonthlyFarmerMilkDetail
#Region "Variable"
    Public PK_Id As Integer = Nothing
    Public Document_Code As String = Nothing
    Public MCC_Code As String = Nothing ''Not a Table Column
    Public VLC_Code As String = Nothing
    Public VLC_Uploader_Code As String = Nothing ''Not a Table Column
    Public VLC_Name As String = Nothing ''Not a Table Column

    Public MP_Code As String = Nothing
    Public MP_Uploader_Code As String = Nothing ''Not a Table Column
    Public MP_Name As String = Nothing ''Not a Table Column
    Public Collection_Qty As Decimal
    Public Qty As Decimal
#End Region
    Public Shared Function SaveData(ByVal FromDate As Date, ByVal arrObj As List(Of clsDBTMonthlyFarmerMilkDetail), ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try

            Dim qry As String = ""
            Dim coll As Hashtable
            If arrObj IsNot Nothing Then
                For Each obj As clsDBTMonthlyFarmerMilkDetail In arrObj
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "VLC_Code", obj.VLC_Code)
                    clsCommon.AddColumnsForChange(coll, "MP_Code", obj.MP_Code)
                    clsCommon.AddColumnsForChange(coll, "Cycle_No", 1)
                    clsCommon.AddColumnsForChange(coll, "Cycle_Month", FromDate.Month)
                    clsCommon.AddColumnsForChange(coll, "Cycle_Year", FromDate.Year)
                    clsCommon.AddColumnsForChange(coll, "Collection_Qty", obj.Collection_Qty)
                    clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As List(Of clsDBTMonthlyFarmerMilkDetail)
        Try
            Dim arrObj As List(Of clsDBTMonthlyFarmerMilkDetail) = Nothing
            Dim obj As clsDBTMonthlyFarmerMilkDetail = Nothing
            Dim qry As String = "Select TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.*,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MP_MASTER.MP_Code_VLC_Uploader,TSPL_MP_MASTER.MP_Name,TSPL_VLC_MASTER_HEAD.MCC 
from TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL 
Left Outer Join TSPL_MP_MASTER On TSPL_MP_MASTER.MP_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.MP_Code   
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.VLC_Code
where TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Document_Code='" & strDocNo & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of clsDBTMonthlyFarmerMilkDetail)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsDBTMonthlyFarmerMilkDetail()
                    obj.PK_Id = clsCommon.myCdbl(dt.Rows(i)("PK_Id"))
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.MCC_Code = clsCommon.myCstr(dt.Rows(i)("MCC"))
                    obj.VLC_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Code"))
                    obj.VLC_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Code_VLC_Uploader"))
                    obj.VLC_Name = clsCommon.myCstr(dt.Rows(i)("VLC_Name"))
                    obj.MP_Code = clsCommon.myCstr(dt.Rows(i)("MP_Code"))
                    obj.MP_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("MP_Code_VLC_Uploader"))
                    obj.MP_Name = clsCommon.myCstr(dt.Rows(i)("MP_Name"))
                    obj.Collection_Qty = clsCommon.myCdbl(dt.Rows(i)("Collection_Qty"))
                    obj.Qty = clsCommon.myCdbl(dt.Rows(i)("Qty"))
                    arrObj.Add(obj)
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function GetQry(ByVal strDocNo As String) As String
        Dim qry As String = ""
        qry = "Select  TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as MP_Uploader_Code,TSPL_MP_MASTER.MP_Name,TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Collection_Qty,TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Qty,TSPL_MP_MASTER.THIRD_PARTY_CODE
from TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL 
left outer join TSPL_DBT_MONTHLY_FARMER_MILK on TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Document_Code
Left Outer Join TSPL_MP_MASTER On TSPL_MP_MASTER.MP_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.MP_Code   
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.VLC_Code
where TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Document_Code='" + strDocNo + "' order by TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.PK_Id"
        Return qry
    End Function
End Class