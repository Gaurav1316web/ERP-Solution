Imports System.Data.SqlClient
Imports common
Public Class clsMPDCSInsentiveReco
#Region "variables"
    Public Document_Code As String = Nothing
    Public Document_Date As Date
    Public Reco_Date As Date
    Public Reco_Date_To As Date
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public Posting_Date As Date? = Nothing
    Public Zone_Code As String = Nothing
    Public Apply_FAT_Above As Decimal = 0
    Public Apply_SNF_Above As Decimal = 0
    Public Third_Party_Integrated As Integer = Nothing
    Public arr As List(Of clsMPDCSInsentiveRecoDetail) = Nothing

#End Region

    Public Shared Function SaveData(ByVal obj As clsMPDCSInsentiveReco, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        Try

            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DCSMPIncentiveReco, obj.arr(0).MCC_Code, obj.Document_Date, trans)
            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where Document_Code='" & obj.Document_Code & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where Document_Code='" & obj.Document_Code & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'obj.Reco_Date = obj.Reco_Date ''New Date(obj.Reco_Date.Year, obj.Reco_Date.Month, 1)
            'obj.Reco_Date_To = obj.Reco_Date_To
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Reco_Date", clsCommon.GetPrintDate(obj.Reco_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Reco_Date_To", clsCommon.GetPrintDate(obj.Reco_Date_To, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            clsCommon.AddColumnsForChange(coll, "Zone_Code", obj.Zone_Code, True)
            clsCommon.AddColumnsForChange(coll, "Third_Party_Integrated", obj.Third_Party_Integrated)
            clsCommon.AddColumnsForChange(coll, "Apply_FAT_Above", obj.Apply_FAT_Above, True)
            clsCommon.AddColumnsForChange(coll, "Apply_SNF_Above", obj.Apply_SNF_Above, True)

            If isNewEntry Then
                qry = "select max(Document_Code)as Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Reco_Date='" + clsCommon.GetPrintDate(obj.Reco_Date) + "'"
                obj.Document_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                If clsCommon.myLen(obj.Document_Code) <= 0 Then
                    obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.MPDCSIncentiveReco, "", "")
                ElseIf obj.Document_Code.Contains(".") Then
                    obj.Document_Code = clsCommon.incval(obj.Document_Code)
                Else
                    obj.Document_Code = obj.Document_Code + ".01"
                End If
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", OMInsertOrUpdate.Update, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            clsMPDCSInsentiveRecoDetail.saveData(obj.Document_Code, obj.arr, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", "Document_Code", trans)
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsMPDCSInsentiveReco
        Return GetData(strCode, NavType, "", Nothing)
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal SelectedZone As String, ByVal trans As SqlTransaction) As clsMPDCSInsentiveReco
        Dim obj As clsMPDCSInsentiveReco = Nothing
        Dim Arr As List(Of clsMPDCSInsentiveReco) = Nothing
        Dim qry As String = "Select TSPL_DCS_MP_INCENTIVE_RECO_HEAD.* from TSPL_DCS_MP_INCENTIVE_RECO_HEAD  where 2=2 "
        Dim whrclas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = (select MIN(Document_Code) from TSPL_DCS_MP_INCENTIVE_RECO_HEAD WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Last
                qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = (select Max(Document_Code) from TSPL_DCS_MP_INCENTIVE_RECO_HEAD WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Current
                qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = '" + strCode + "' " + whrclas + " "
            Case NavigatorType.Next
                qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = (select Min(Document_Code) from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code>'" + strCode + "' " + whrclas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = (select Max(Document_Code) from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code<'" + strCode + "' " + whrclas + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsMPDCSInsentiveReco()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Zone_Code = clsCommon.myCstr(dt.Rows(0)("Zone_Code"))
            obj.Reco_Date = clsCommon.myCDate(dt.Rows(0)("Reco_Date"))
            obj.Reco_Date_To = clsCommon.myCDate(dt.Rows(0)("Reco_Date_To"))
            obj.Status = IIf(clsCommon.myCDecimal(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)
            obj.Third_Party_Integrated = clsCommon.myCDecimal(dt.Rows(0)("Third_Party_Integrated"))
            obj.Apply_FAT_Above = clsCommon.myCDecimal(dt.Rows(0)("Apply_SNF_Above"))
            obj.Apply_SNF_Above = clsCommon.myCDecimal(dt.Rows(0)("Apply_SNF_Above"))
            If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                obj.Posting_Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
            End If
            obj.arr = clsMPDCSInsentiveRecoDetail.getData(obj.Document_Code, SelectedZone, trans)
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

            Dim Location_code As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MCC_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where Document_Code='" + strDocNo + "'", trans))

            If (clsCommon.myLen(Location_code) <= 0) Then
                Location_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MCC_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where Document_Code='" + strDocNo + "'", trans))
            End If
            Dim Document_Date As DateTime = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select top 1 Document_Date from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code='" + strDocNo + "'", trans))
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DCSMPIncentiveReco, Location_code, Document_Date, trans)
            clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", "Document_Code", trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", "Document_Code", trans)

            Dim qry As String = ""
            qry = "delete from TSPL_ATTACHMENTS where FormId='" + clsUserMgtCode.DCSMPIncentiveReco + "' and TransactionId='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = "Select TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code as Code,Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Date,103) as Date
           , SUBSTRING( REPLACE(convert(varchar, TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,106),' ',' /'),5,10) as [Reco Date],Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) as [From Date],Convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To,103) as [To Date],zone_code as Zone
          ,case when isnull(Status,0)=0 then 'Pending' else 'Approved' end as Status 
          from TSPL_DCS_MP_INCENTIVE_RECO_HEAD  "
        str = clsCommon.ShowSelectForm("DCMPInc#F", qry, "Code", whrcls, curcode, "Code", isButtonClicked, "Document_Date")
        Return str
    End Function
    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Document No not found to Post")
            End If

            Dim Location_code As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MCC_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where Document_Code='" + strDocNo + "'", trans))

            If (clsCommon.myLen(Location_code) <= 0) Then
                Location_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MCC_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where Document_Code='" + strDocNo + "'", trans))
            End If

            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")

            Dim obj As clsMPDCSInsentiveReco = clsMPDCSInsentiveReco.GetData(strDocNo, NavigatorType.Current, "", trans)
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.DCSMPIncentiveReco, Location_code, obj.Document_Date, trans)


            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If (obj.Status = ERPTransactionStatus.Approved) Then
                Throw New Exception("Already Post on :" + obj.Posting_Date)
            End If
            Dim qry As String = ""
            Dim ExtrColumn As String = " DBT_Capping_Apply=0"
            Dim SettDBTMilkQtyCapping As Integer = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DBTMilkQtyCapping, clsFixedParameterCode.DBTMilkQtyCapping, trans))
            If SettDBTMilkQtyCapping > 0 Then
                qry = "Update TSPL_MP_MASTER set DBT_Capping_Qty=" + clsCommon.myCstr(SettDBTMilkQtyCapping) + " where DBT_Capping_Qty is null"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                ExtrColumn = " DBT_Capping_Apply=1"
            End If
            qry = "Update TSPL_DCS_MP_INCENTIVE_RECO_HEAD set Status=1 ," + ExtrColumn + ", Posting_Date='" + strPostDate + "',Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", trans)

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
            Dim obj As clsMPDCSInsentiveReco = clsMPDCSInsentiveReco.GetData(strDocNo, NavigatorType.Current, "", trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                clsCommon.MyMessageBoxShow("No Data found to Reverse And UnPost")
            End If

            If Not obj.Status = ERPTransactionStatus.Approved Then
                clsCommon.MyMessageBoxShow("Transaction status should be posted for reverse and unpost")
            End If

            Dim qry As String = "select Document_Code from TSPL_DBT_CAPING where Reco_Code='" + strDocNo + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("DBT Capping Document No [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "] is generated.you cannot unpost it")
            End If

            qry = "select Document_Code from TSPL_DBT_NEFT where From_Date='" + clsCommon.GetPrintDate(obj.Reco_Date, "dd/MMM/yyyy") + "' and To_Date='" + clsCommon.GetPrintDate(obj.Reco_Date_To, "dd/MMM/yyyy") + "'"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("DBT NEFT Document No [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "] is generated.you cannot unpost it")
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Status", 0)
            clsCommon.AddColumnsForChange(coll, "Posted_By", Nothing, True)
            clsCommon.AddColumnsForChange(coll, "Posting_Date", Nothing, True)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", OMInsertOrUpdate.Update, "Document_Code='" + strDocNo + "'", trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", "Document_Code", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


End Class

Public Class clsMPDCSInsentiveRecoDetail
#Region "Variable"
    Public PK_Id As Integer = Nothing

    Public Document_Code As String = Nothing
    Public SNo As Integer

    Public Route_Code As String = Nothing
    Public Route_Name As String = Nothing

    Public MCC_Code As String = Nothing
    Public MCC_Name As String = Nothing ''Not a Table Column

    Public Zone_Code As String = Nothing ''Not a Table Column
    Public Zone_Name As String = Nothing ''Not a Table Column

    Public VLC_Code As String = Nothing
    Public VLC_Uploader_Code As String = Nothing ''Not a Table Column
    Public VLC_Name As String = Nothing ''Not a Table Column

    Public Cycle_Year As Integer
    Public Cycle_Month As Integer
    Public Cycle_No As Integer
    Public Qty As Decimal = 0
    Public UOM As String = Nothing
    Public FAT As Decimal = 0
    Public SNF As Decimal = 0
    Public Amount As Decimal = 0
    Public MP_Count As Decimal = 0
    Public MP_Qty As Decimal = 0
    Public MP_FAT As Decimal = 0
    Public MP_SNF As Decimal = 0
    Public MP_Amount As Decimal = 0

    Public Diff_Qty As Decimal = 0
    Public Diff_FAT As Decimal = 0
    Public Diff_SNF As Decimal = 0
    Public Diff_Amount As Decimal = 0

    Public Reco_Staus As Boolean = False ''Not a Table Column
    Public Reco_Staus_OLD As Boolean = False ''Not a Table Column

    Public MPCollectionQty As Decimal = 0 ''Not a Table Column


#End Region
    Public Shared Function saveDataZone(ByVal strDocNo As String, ByVal strZone As String, ByVal arrObj As List(Of clsMPDCSInsentiveRecoDetail)) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim qry As String = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where PK_Id in (
select PK_Id from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL 
inner join tspl_vlc_master_head on tspl_vlc_master_head.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code
inner join tspl_vendor_master on tspl_vendor_master.Vendor_Code=tspl_vlc_master_head.vsp_code
where Document_Code='" + strDocNo + "' and tspl_vendor_master.zone_Code in (" + strZone + ") ) "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where PK_Id in (
select PK_Id from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID 
inner join tspl_vlc_master_head on tspl_vlc_master_head.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.VLC_Code
inner join tspl_vendor_master on tspl_vendor_master.Vendor_Code=tspl_vlc_master_head.vsp_code
where Document_Code='" + strDocNo + "' and tspl_vendor_master.zone_Code in (" + strZone + ") ) "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)


            saveData(strDocNo, arrObj, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
        End Try
        Return True
    End Function

    Public Shared Function saveData(ByVal strDocNo As String, ByVal arrObj As List(Of clsMPDCSInsentiveRecoDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            If arrObj IsNot Nothing AndAlso arrObj.Count > 0 Then
                For Each obj As clsMPDCSInsentiveRecoDetail In arrObj
                    saveDataOne(strDocNo, obj, trans)

                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Shared Sub saveDataOne(strDocNo As String, obj As clsMPDCSInsentiveRecoDetail, trans As SqlTransaction)
        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
        clsCommon.AddColumnsForChange(coll, "SNo", obj.SNo)
        clsCommon.AddColumnsForChange(coll, "MCC_Code", obj.MCC_Code)
        clsCommon.AddColumnsForChange(coll, "VLC_Code", obj.VLC_Code)

        clsCommon.AddColumnsForChange(coll, "Cycle_Year", obj.Cycle_Year)
        clsCommon.AddColumnsForChange(coll, "Cycle_Month", obj.Cycle_Month)
        clsCommon.AddColumnsForChange(coll, "Cycle_No", obj.Cycle_No)

        clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)
        clsCommon.AddColumnsForChange(coll, "UOM", obj.UOM)
        clsCommon.AddColumnsForChange(coll, "FAT", obj.FAT)
        clsCommon.AddColumnsForChange(coll, "SNF", obj.SNF)
        clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
        clsCommon.AddColumnsForChange(coll, "MP_Count", obj.MP_Count)
        clsCommon.AddColumnsForChange(coll, "MP_Qty", obj.MP_Qty)
        clsCommon.AddColumnsForChange(coll, "MP_FAT", obj.MP_FAT)
        clsCommon.AddColumnsForChange(coll, "MP_SNF", obj.MP_SNF)
        clsCommon.AddColumnsForChange(coll, "MP_Amount", obj.MP_Amount)

        clsCommon.AddColumnsForChange(coll, "Diff_Qty", obj.Diff_Qty)
        clsCommon.AddColumnsForChange(coll, "Diff_FAT", obj.Diff_FAT)
        clsCommon.AddColumnsForChange(coll, "Diff_SNF", obj.Diff_SNF)
        clsCommon.AddColumnsForChange(coll, "Diff_Amount", obj.Diff_Amount)
        clsCommonFunctionality.UpdateDataTable(coll, IIf(obj.Reco_Staus, "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID"), OMInsertOrUpdate.Insert, "", trans)

    End Sub

    Public Shared Function getData(ByVal strDocNo As String, ByVal SelectedZone As String, ByVal trans As SqlTransaction) As List(Of clsMPDCSInsentiveRecoDetail)
        Try
            Dim arrObj As List(Of clsMPDCSInsentiveRecoDetail) = Nothing
            Dim obj As clsMPDCSInsentiveRecoDetail = Nothing
            Dim qry As String = "Select TSPL_VLC_MASTER_HEAD.Route_Code,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.*,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader
,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MCC_MASTER.MCC_NAME,TSPL_VENDOR_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description as Zone_Name,TabCollectionQty.Qty as MPCollectionQty
from ( select PK_Id,Document_Code,SNo,Cycle_Year,Cycle_Month,Cycle_No,MCC_Code,VLC_Code,Qty,UOM,FAT,SNF,Amount,MP_Count,MP_Qty,MP_FAT,MP_SNF,MP_Amount,Diff_Qty,Diff_FAT,Diff_SNF,Diff_Amount,1 as Reco_Staus from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code='" & strDocNo & "'
union all
select PK_Id,Document_Code,SNo,Cycle_Year,Cycle_Month,Cycle_No,MCC_Code,VLC_Code,Qty,UOM,FAT,SNF,Amount,MP_Count,MP_Qty,MP_FAT,MP_SNF,MP_Amount,Diff_Qty,Diff_FAT,Diff_SNF,Diff_Amount,0 as Reco_Staus from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.Document_Code='" & strDocNo & "' ) as TSPL_DCS_MP_INCENTIVE_RECO_DETAIL 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code 
left outer join ( select VLC_Code,sum(Qty) as Qty from TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL 
left outer join TSPL_DBT_MONTHLY_FARMER_MILK on TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code=TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL.Document_Code
where TSPL_DBT_MONTHLY_FARMER_MILK.DBT_Reco_Code='" + strDocNo + "' group by VLC_Code ) TabCollectionQty on TabCollectionQty.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code
where 2=2 "

            If clsCommon.myLen(SelectedZone) > 0 Then
                qry += " and TSPL_VENDOR_MASTER.Zone_Code in (" + SelectedZone + ")"
            End If
            qry += "Order by TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.SNo"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of clsMPDCSInsentiveRecoDetail)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New clsMPDCSInsentiveRecoDetail()
                    obj.PK_Id = clsCommon.myCDecimal(dt.Rows(i)("PK_Id"))
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.SNo = i + 1
                    obj.Route_Code = clsCommon.myCstr(dt.Rows(i)("Route_Code"))
                    obj.Route_Name = clsCommon.myCstr(dt.Rows(i)("Route_Name"))

                    obj.MCC_Code = clsCommon.myCstr(dt.Rows(i)("MCC_Code"))
                    obj.MCC_Name = clsCommon.myCstr(dt.Rows(i)("MCC_Name"))

                    obj.VLC_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Code"))
                    obj.VLC_Uploader_Code = clsCommon.myCstr(dt.Rows(i)("VLC_Code_VLC_Uploader"))
                    obj.VLC_Name = clsCommon.myCstr(dt.Rows(i)("VLC_Name"))

                    obj.Zone_Code = clsCommon.myCstr(dt.Rows(i)("Zone_Code"))
                    obj.Zone_Name = clsCommon.myCstr(dt.Rows(i)("Zone_Name"))

                    obj.Cycle_Year = clsCommon.myCDecimal(dt.Rows(i)("Cycle_Year"))
                    obj.Cycle_Month = clsCommon.myCDecimal(dt.Rows(i)("Cycle_Month"))
                    obj.Cycle_No = clsCommon.myCDecimal(dt.Rows(i)("Cycle_No"))

                    obj.Qty = clsCommon.myCDecimal(dt.Rows(i)("Qty"))
                    obj.UOM = clsCommon.myCstr(dt.Rows(i)("UOM"))
                    obj.FAT = clsCommon.myCDecimal(dt.Rows(i)("FAT"))
                    obj.SNF = clsCommon.myCDecimal(dt.Rows(i)("SNF"))
                    obj.Amount = clsCommon.myCDecimal(dt.Rows(i)("Amount"))

                    obj.MP_Count = clsCommon.myCDecimal(dt.Rows(i)("MP_Count"))
                    obj.MP_Qty = clsCommon.myCDecimal(dt.Rows(i)("MP_Qty"))
                    obj.MP_FAT = clsCommon.myCDecimal(dt.Rows(i)("MP_FAT"))
                    obj.MP_SNF = clsCommon.myCDecimal(dt.Rows(i)("MP_SNF"))
                    obj.MP_Amount = clsCommon.myCDecimal(dt.Rows(i)("MP_Amount"))

                    obj.Diff_Qty = clsCommon.myCDecimal(dt.Rows(i)("Diff_Qty"))
                    obj.Diff_FAT = clsCommon.myCDecimal(dt.Rows(i)("Diff_FAT"))
                    obj.Diff_SNF = clsCommon.myCDecimal(dt.Rows(i)("Diff_SNF"))
                    obj.Diff_Amount = clsCommon.myCDecimal(dt.Rows(i)("Diff_Amount"))
                    obj.MPCollectionQty = clsCommon.myCDecimal(dt.Rows(i)("MPCollectionQty"))

                    obj.Reco_Staus = (clsCommon.myCDecimal(dt.Rows(i)("Reco_Staus")) = 1)
                    arrObj.Add(obj)
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function EditDCSQty(ByVal strDocNo As String, ByVal strRemarks As String, ByVal Arr As List(Of clsMPDCSInsentiveRecoDetail)) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim obj As clsMPDCSInsentiveReco = clsMPDCSInsentiveReco.GetData(strDocNo, NavigatorType.Current, "", trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                clsCommon.MyMessageBoxShow("No Data found to Reverse And UnPost")
            End If

            If obj.Status = ERPTransactionStatus.Approved Then
                clsCommon.MyMessageBoxShow("Transaction status should be pending")
            End If

            Dim qry As String = "select Document_Code from TSPL_DBT_CAPING where Reco_Code='" + strDocNo + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("DBT Capping Document No [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "] is generated.you cannot unpost it")
            End If

            qry = "select Document_Code from TSPL_DBT_NEFT where From_Date='" + clsCommon.GetPrintDate(obj.Reco_Date, "dd/MMM/yyyy") + "' and To_Date='" + clsCommon.GetPrintDate(obj.Reco_Date_To, "dd/MMM/yyyy") + "'"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("DBT NEFT Document No [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "] is generated.you cannot unpost it")
            End If
            If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                clsCommon.AddColumnsForChange(coll, "Edit_Remarks", strRemarks)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", OMInsertOrUpdate.Update, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code='" + obj.Document_Code + "'", trans)

                For Each objTR As clsMPDCSInsentiveRecoDetail In Arr
                    If objTR.Reco_Staus_OLD Then
                        qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where Document_Code='" & obj.Document_Code & "' and PK_Id='" + clsCommon.myCstr(objTR.PK_Id) + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    Else
                        qry = "delete from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID where Document_Code='" & obj.Document_Code & "'  and PK_Id='" + clsCommon.myCstr(objTR.PK_Id) + "' "
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    End If
                    saveDataOne(obj.Document_Code, objTR, trans)
                Next
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", "Document_Code", "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", "Document_Code", trans)
                trans.Commit()
            Else
                Throw New Exception("No Data found to edit")
            End If
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class
