Imports common
Imports System.Data.SqlClient


Public Class clsBMCTransporterBill

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
    Public Arr As List(Of clsBMCTransporterBillDetail) = Nothing
    Public ArrDT As DataTable = Nothing
    Public Comment As String = Nothing
    Public Remarks As String = Nothing
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


    'Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
    '    Try
    '        Dim str As String = ""
    '        Dim qry As String = ""

    '        qry = "select TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code ,convert(varchar,TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_date,103) as Document_date,TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No,case when TSPL_BMC_TRANSPORTER_BILL_HEAD.status =1  then 'Approved' else 'Pending' end as Status   
    '                         from TSPL_BMC_TRANSPORTER_BILL_HEAD "
    '        str = clsCommon.ShowSelectForm("fndPayProcess", qry, "Document_Code", "", curcode, "Document_Code", isButtonClicked, "Document_date")

    '        Return str
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function
    Public Function SaveData(ByVal obj As clsBMCTransporterBill, ByVal isNewEntry As Boolean) As Boolean
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

    Public Function SaveData(ByVal obj As clsBMCTransporterBill, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True

        Dim qry As String = "delete from TSPL_BMC_TRANSPORTER_BILL_DETAIL where Document_Code='" + obj.Document_Code + "'"
        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)


        'If obj.Arr.Count <= 0 Then
        '    Throw New Exception("No detail found to save")
        'End If

        If (isNewEntry) Then
            obj.Document_Code = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Document_Date), clsDocType.BMCTransporterBill, "", Nothing)
        End If
        If (clsCommon.myLen(obj.Document_Code) <= 0) Then
            Throw New Exception("Error in Document Code Generation")
        End If

        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "From_Date", clsCommon.GetPrintDate(obj.From_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "To_Date", clsCommon.GetPrintDate(obj.To_Date, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "Tanker_No", obj.Tanker_No)
        clsCommon.AddColumnsForChange(coll, "Toll_Tax", obj.Toll_Tax)
        clsCommon.AddColumnsForChange(coll, "Ice_Charge", obj.Ice_Charge)


        'clsCommon.AddColumnsForChange(coll, "Fat_Shortage", obj.Fat_Shortage)
        'clsCommon.AddColumnsForChange(coll, "Snf_Shortage", obj.Snf_Shortage)
        'clsCommon.AddColumnsForChange(coll, "Fat_Snf_Shortage", obj.Fat_Snf_Shortage)
        'clsCommon.AddColumnsForChange(coll, "Fat_Rate", obj.Fat_Rate)
        'clsCommon.AddColumnsForChange(coll, "Snf_Rate", obj.Snf_Rate)
        clsCommon.AddColumnsForChange(coll, "Fat_Shortage", obj.Fat_Shortage)
        clsCommon.AddColumnsForChange(coll, "Fat_Shortage_NMG", obj.Fat_Shortage_NMG)
        clsCommon.AddColumnsForChange(coll, "Snf_Shortage", obj.Snf_Shortage)
        clsCommon.AddColumnsForChange(coll, "Snf_Shortage_NMG", obj.Snf_Shortage_NMG)
        clsCommon.AddColumnsForChange(coll, "Fat_Snf_Shortage", obj.Fat_Snf_Shortage)
        clsCommon.AddColumnsForChange(coll, "Fat_Rate", obj.Fat_Rate)
        clsCommon.AddColumnsForChange(coll, "Fat_Rate_NMG", obj.Fat_Rate_NMG)
        clsCommon.AddColumnsForChange(coll, "Snf_Rate", obj.Snf_Rate)
        clsCommon.AddColumnsForChange(coll, "Snf_Rate_NMG", obj.Snf_Rate_NMG)

        clsCommon.AddColumnsForChange(coll, "Tanker_Capacity", obj.Tanker_Capacity)
        clsCommon.AddColumnsForChange(coll, "KM_Rate", obj.KM_Rate)
        clsCommon.AddColumnsForChange(coll, "Total_Amount", obj.Total_Amount)
        clsCommon.AddColumnsForChange(coll, "Gross_Amount", obj.Gross_Amount)
        clsCommon.AddColumnsForChange(coll, "Total_Addition", obj.Total_Addition)
        clsCommon.AddColumnsForChange(coll, "Total_Deduction", obj.Total_Deduction)
        clsCommon.AddColumnsForChange(coll, "Diesel_Rate_Plus", obj.Diesel_Rate_Plus)
        clsCommon.AddColumnsForChange(coll, "Diesel_Rate_Minus", obj.Diesel_Rate_Minus)
        clsCommon.AddColumnsForChange(coll, "Total_Diesel", obj.Total_Diesel)
        clsCommon.AddColumnsForChange(coll, "Prorata_Amt", obj.Prorata_Amt)
        clsCommon.AddColumnsForChange(coll, "Total_Before_Calc", obj.Total_Before_Calc)

        clsCommon.AddColumnsForChange(coll, "Comment", obj.Comment)
        clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
        clsCommon.AddColumnsForChange(coll, "Is_Private", IIf(obj.Is_Private, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Labour_Charge", obj.Labour_Charge)
        clsCommon.AddColumnsForChange(coll, "Labour_Amt", obj.Labour_Amt)
        clsCommon.AddColumnsForChange(coll, "Employee_PF", obj.Employee_PF)
        clsCommon.AddColumnsForChange(coll, "Employer_PF", obj.Employer_PF)
        clsCommon.AddColumnsForChange(coll, "Admin_Charge", obj.Admin_Charge)
        clsCommon.AddColumnsForChange(coll, "EDLI_Charge", obj.EDLI_Charge)
        clsCommon.AddColumnsForChange(coll, "Total_PF_Amount", obj.Total_PF_Amount)
        clsCommon.AddColumnsForChange(coll, "ESI_Employee", obj.ESI_Employee)
        clsCommon.AddColumnsForChange(coll, "ESI_Employer", obj.ESI_Employer)
        clsCommon.AddColumnsForChange(coll, "Total_ESI_Amount", obj.Total_ESI_Amount)
        clsCommon.AddColumnsForChange(coll, "COESI_PER", obj.COESI_PER)
        clsCommon.AddColumnsForChange(coll, "EMPESI_PER", obj.EMPESI_PER)
        clsCommon.AddColumnsForChange(coll, "EMPEPF_PER", obj.EMPEPF_PER)
        clsCommon.AddColumnsForChange(coll, "ACCOEPF_PER", obj.ACCOEPF_PER)
        clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
        clsCommon.AddColumnsForChange(coll, "Posted_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Posted_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
        If isNewEntry Then
            'obj.Document_Code = clsERPFuncationality.GetNextCode(trans, DateTime.Now, clsDocType.DistributorRouteTagging, "", "")
            'If clsCommon.myLen(obj.Document_Code) <= 0 Then
            '    Throw New Exception("Error in Code Generation")
            'End If
            clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
            clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_BMC_TRANSPORTER_BILL_HEAD", OMInsertOrUpdate.Insert, "", trans)
        Else
            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_BMC_TRANSPORTER_BILL_HEAD", OMInsertOrUpdate.Update, "Document_Code='" + obj.Document_Code + "'", trans)
        End If
        isSaved = isSaved AndAlso clsBMCTransporterBillDetail.SaveData(obj.Document_Code, Arr, trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_Code, "TSPL_BMC_TRANSPORTER_BILL_HEAD", "Document_Code", "TSPL_BMC_TRANSPORTER_BILL_DETAIL", "Document_Code", trans)
        Return isSaved

    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType, ByVal isGetDT As Boolean, ByVal trans As SqlTransaction) As clsBMCTransporterBill
        Dim obj As clsBMCTransporterBill = Nothing

        Dim qry As String = "Select TSPL_BMC_TRANSPORTER_BILL_HEAD.* from TSPL_BMC_TRANSPORTER_BILL_HEAD where 2=2 "
        Dim whrClas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = (select MIN(Document_Code) from TSPL_BMC_TRANSPORTER_BILL_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Last
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = (select Max(Document_Code) from TSPL_BMC_TRANSPORTER_BILL_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Next
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = (select Min(Document_Code) from TSPL_BMC_TRANSPORTER_BILL_HEAD where Document_Code>'" + strDocumentNo + "' " + whrClas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = (select Max(Document_Code) from TSPL_BMC_TRANSPORTER_BILL_HEAD where Document_Code<'" + strDocumentNo + "' " + whrClas + ")"
            Case NavigatorType.Current
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = '" + strDocumentNo + "'"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsBMCTransporterBill()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCstr(dt.Rows(0)("Document_Date"))
            obj.From_Date = clsCommon.myCstr(dt.Rows(0)("From_Date"))
            obj.To_Date = clsCommon.myCstr(dt.Rows(0)("To_Date"))
            obj.Tanker_No = clsCommon.myCstr(dt.Rows(0)("Tanker_No"))
            obj.Toll_Tax = clsCommon.myCdbl(dt.Rows(0)("Toll_Tax"))
            obj.Ice_Charge = clsCommon.myCdbl(dt.Rows(0)("Ice_Charge"))
            obj.Fat_Shortage = clsCommon.myCdbl(dt.Rows(0)("Fat_Shortage"))

            obj.Fat_Rate_NMG = clsCommon.myCdbl(dt.Rows(0)("Fat_Rate_NMG"))
            obj.Snf_Rate_NMG = clsCommon.myCdbl(dt.Rows(0)("Snf_Rate_NMG"))
            obj.Fat_Shortage_NMG = clsCommon.myCdbl(dt.Rows(0)("Fat_Shortage_NMG"))
            obj.Snf_Shortage_NMG = clsCommon.myCdbl(dt.Rows(0)("Snf_Shortage_NMG"))


            obj.Snf_Shortage = clsCommon.myCdbl(dt.Rows(0)("Snf_Shortage"))
            obj.Fat_Snf_Shortage = clsCommon.myCdbl(dt.Rows(0)("Fat_Snf_Shortage"))
            obj.Fat_Rate = clsCommon.myCdbl(dt.Rows(0)("Fat_Rate"))
            obj.Snf_Rate = clsCommon.myCdbl(dt.Rows(0)("Snf_Rate"))
            obj.Tanker_Capacity = clsCommon.myCdbl(dt.Rows(0)("Tanker_Capacity"))
            obj.KM_Rate = clsCommon.myCdbl(dt.Rows(0)("KM_Rate"))
            obj.Total_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Amount"))
            obj.Gross_Amount = clsCommon.myCdbl(dt.Rows(0)("Gross_Amount"))
            obj.Total_Addition = clsCommon.myCdbl(dt.Rows(0)("Total_Addition"))
            obj.Total_Deduction = clsCommon.myCdbl(dt.Rows(0)("Total_Deduction"))
            obj.Diesel_Rate_Plus = clsCommon.myCdbl(dt.Rows(0)("Diesel_Rate_Plus"))
            obj.Diesel_Rate_Minus = clsCommon.myCdbl(dt.Rows(0)("Diesel_Rate_Minus"))
            obj.Total_Diesel = clsCommon.myCdbl(dt.Rows(0)("Total_Diesel"))
            obj.Total_Before_Calc = clsCommon.myCdbl(dt.Rows(0)("Total_Before_Calc"))
            obj.Prorata_Amt = clsCommon.myCdbl(dt.Rows(0)("Prorata_Amt"))
            obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 0, ERPTransactionStatus.Pending, ERPTransactionStatus.Approved)

            obj.Comment = clsCommon.myCstr(dt.Rows(0)("Comment"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Is_Private = (clsCommon.myCDecimal(dt.Rows(0)("Is_Private")) = 1)
            obj.Labour_Charge = clsCommon.myCdbl(dt.Rows(0)("Labour_Charge"))
            obj.Labour_Amt = clsCommon.myCdbl(dt.Rows(0)("Labour_Amt"))
            obj.Employee_PF = clsCommon.myCdbl(dt.Rows(0)("Employee_PF"))
            obj.Employer_PF = clsCommon.myCdbl(dt.Rows(0)("Employer_PF"))
            obj.Admin_Charge = clsCommon.myCdbl(dt.Rows(0)("Admin_Charge"))
            obj.EDLI_Charge = clsCommon.myCdbl(dt.Rows(0)("EDLI_Charge"))
            obj.Total_PF_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_PF_Amount"))
            obj.ESI_Employee = clsCommon.myCdbl(dt.Rows(0)("ESI_Employee"))
            obj.ESI_Employer = clsCommon.myCdbl(dt.Rows(0)("ESI_Employer"))
            obj.Total_ESI_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_ESI_Amount"))
            obj.COESI_PER = clsCommon.myCdbl(dt.Rows(0)("COESI_PER"))
            obj.EMPESI_PER = clsCommon.myCdbl(dt.Rows(0)("EMPESI_PER"))
            obj.EMPEPF_PER = clsCommon.myCdbl(dt.Rows(0)("EMPEPF_PER"))
            obj.ACCOEPF_PER = clsCommon.myCdbl(dt.Rows(0)("ACCOEPF_PER"))

            qry = "Select TSPL_BMC_TRANSPORTER_BILL_DETAIL.* from TSPL_BMC_TRANSPORTER_BILL_DETAIL 
                   where TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code='" + obj.Document_Code + "' ORDER BY TSPL_BMC_TRANSPORTER_BILL_DETAIL.PK_ID"
            obj.ArrDT = clsDBFuncationality.GetDataTable(qry, trans)

            If (obj.ArrDT IsNot Nothing AndAlso obj.ArrDT.Rows.Count > 0) Then
                obj.Arr = New List(Of clsBMCTransporterBillDetail)
                For Each dr As DataRow In obj.ArrDT.Rows
                    Dim objTr As New clsBMCTransporterBillDetail
                    objTr.Document_Code = clsCommon.myCstr(dr("Document_Code"))
                    objTr.PK_ID = clsCommon.myCstr(dr("PK_ID"))
                    objTr.MCC_Document_Code = clsCommon.myCstr(dr("MCC_Document_Code"))
                    objTr.Station_1 = clsCommon.myCstr(dr("Station_1"))
                    objTr.Station_2 = clsCommon.myCstr(dr("Station_2"))
                    objTr.Station_3 = clsCommon.myCstr(dr("Station_3"))
                    objTr.Station_4 = clsCommon.myCstr(dr("Station_4"))
                    objTr.Category = clsCommon.myCstr(dr("Category"))
                    objTr.Trip = clsCommon.myCdbl(dr("trip"))
                    objTr.GPS_KM = clsCommon.myCdbl(dr("GPS_KM"))
                    objTr.KM = clsCommon.myCdbl(dr("KM"))
                    objTr.Quantity_KG = clsCommon.myCdbl(dr("Quantity_KG"))
                    objTr.Diesel_RD = clsCommon.myCdbl(dr("Diesel_RD"))
                    objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                    'objTr.BMC_Date = clsCommon.myCdbl(dr("Document_Date"))
                    objTr.BMC_Date = clsCommon.myCstr(dr("Document_Date"))
                    objTr.Ice_Box = clsCommon.myCstr(dr("Ice_Box"))
                    'clsCommon.AddColumnsForChange(coll, "Ice_Box", obj.Ice_Box)
                    'objTr.BalanceAmount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Balance_Amt from TSPL_VENDOR_INVOICE_HEAD left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code where TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader= '" + clsCommon.myCstr(dr("VLC_Code_VLC_Uploader")) + "' and Transfer_To_Saving=1"))
                    obj.Arr.Add(objTr)
                Next
            End If

        End If
        Return obj
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

        Dim obj As clsBMCTransporterBill = clsBMCTransporterBill.GetData(strDocNo, NavigatorType.Current, False, trans)
        clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_BMC_TRANSPORTER_BILL_HEAD", "Document_Code", "TSPL_BMC_TRANSPORTER_BILL_DETAIL", "Document_Code", trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_BMC_TRANSPORTER_BILL_HEAD", "Document_Code", "TSPL_BMC_TRANSPORTER_BILL_DETAIL", "Document_Code", trans)

        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
            Try
                If obj.Status = ERPTransactionStatus.Approved Then
                    Throw New Exception("Already Post on :" + obj.Posted_Date)
                End If
                Dim qry As String = "delete from TSPL_BMC_TRANSPORTER_BILL_DETAIL where Document_Code='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_BMC_TRANSPORTER_BILL_HEAD where Document_Code='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function

    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = clsBMCTransporterBill.PostData(strDocNo, trans)
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

        Dim obj As clsBMCTransporterBill = clsBMCTransporterBill.GetData(strDocNo, NavigatorType.Current, False, trans)

        If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
            Throw New Exception("No Data found to Post")
        End If

        If (clsCommon.myLen(obj.Posted_Date) > 0) Then
            Throw New Exception("Already Post on :" + obj.Posted_Date)
        End If

        CreateAPInvoiceHeader(obj, trans)
        qry = "Update TSPL_BMC_TRANSPORTER_BILL_HEAD set Posted_Date='" + clsCommon.GetPrintDate(clsCommon.myCDate(clsCommon.GETSERVERDATE(trans)), "dd/MMM/yyyy hh:mm:ss tt") + "',Status=1 ,Posted_By='" + objCommonVar.CurrentUserCode + "' where Document_Code='" + strDocNo + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_BMC_TRANSPORTER_BILL_HEAD", "Document_Code", trans)

        Return True

    End Function

    Private Shared Function CreateAPInvoiceHeader(ByVal obj As clsBMCTransporterBill, ByVal trans As SqlTransaction) As Boolean
        Dim VendAccSet As String = String.Empty
        Dim dtDed As DataTable = clsDBFuncationality.GetDataTable("select Freight_Provision from TSPL_VENDOR_ACCOUNT_SET  where Acct_Set_Code='TPT'", trans)
        If dtDed Is Nothing OrElse dtDed.Rows.Count <= 0 Then
            Throw New Exception("Please set Freight Provision Account")
        End If

        Dim fullStation As String = obj.ArrDT.Rows(1)("Station_1").ToString()
        Dim station1 As String = ""

        If fullStation.Contains(">>") Then
            station1 = fullStation.Substring(fullStation.IndexOf(">>") + 2)
        Else
            station1 = fullStation ' Or set it to "" or some default value if >> is not present
        End If
        'Dim station1 As String = obj.ArrDT.Rows(1)("Station_1").ToString()
        'Dim qry1 As String = clsDBFuncationality.getSingleValue(" select TSPL_BULK_route_master_location.Location_Code from TSPL_BULK_route_master_location
        '                                left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_BULK_route_master_location.Location_Code
        '                                    where TSPL_LOCATION_MASTER.Location_Desc='" + station1 + "'", trans)
        Dim Loc_Seg As String = clsDBFuncationality.getSingleValue(" select top 1 Loc_Segment_Code from TSPL_LOCATION_MASTER", trans)
        For Each objTr As clsBMCTransporterBillDetail In obj.Arr
            Dim objVendInv As New clsVedorInvoiceHead()
            objVendInv.Invoice_Entry_Date = obj.Document_Date
            objVendInv.Document_Type = "I"
            objVendInv.Invoice_Type = "AP"
            objVendInv.loc_code = clsCommon.myCstr(Loc_Seg)
            'objVendInv.loc_code = "JPR"
            objVendInv.Document_Total = objTr.Amount
            objVendInv.Posting_Date = obj.Document_Date
            objVendInv.Vendor_Invoice_Date = obj.Document_Date
            objVendInv.Description = "BMC Transporter Bill Debit Note"
            objVendInv.RefDocNo = objTr.Document_Code
            objVendInv.RefDocType = "TP_BL"
            objVendInv.Balance_Amt = objTr.Amount
            objVendInv.Arr = New List(Of clsVedorInvoiceDetail)

            '' Detail Level Saving
            Dim objVendInvTR As clsVedorInvoiceDetail = New clsVedorInvoiceDetail()
            objVendInvTR.Amount = objTr.Amount
            objVendInvTR.Amount_less_Discount = objTr.Amount
            objVendInvTR.Total_Amount = objTr.Amount


            objVendInvTR.GL_Account_Code = clsCommon.myCstr(dtDed.Rows(0)("Freight_Provision"))
            If clsCommon.myLen(objVendInvTR.GL_Account_Code) <= 0 Then
                Throw New Exception("Please set GL Account Code for Freight_Provision")
            End If
            objVendInvTR.GL_Account_Desc = clsGLAccount.GetName(objVendInvTR.GL_Account_Code, trans)

            Dim qry2 As String = clsDBFuncationality.getSingleValue(" select Tanker_Transporter_Code from TSPL_TANKER_MASTER where Tanker_No='" + obj.Tanker_No + "'", trans)
            Dim qry3 As String = clsDBFuncationality.getSingleValue(" select Vendor_Name from TSPL_VENDOR_MASTER where Vendor_Code='" + qry2 + "'", trans)

            Dim qry4 As String = "Select TSPL_VENDOR_ACCOUNT_SET.Payable_Account,TSPL_VENDOR_MASTER.GSTRegistered,TSPL_VENDOR_MASTER.Vendor_Account,TSPL_VENDOR_MASTER.Terms_Code ,TSPL_TERMS_MASTER.Terms_Desc,TSPL_TERMS_MASTER.No_Days,TSPL_VENDOR_MASTER.Vendor_Name   
from TSPL_VENDOR_MASTER 
left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code =TSPL_VENDOR_MASTER.Terms_Code 
left outer join TSPL_VENDOR_ACCOUNT_SET on TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code = TSPL_VENDOR_MASTER.Vendor_Account
where TSPL_VENDOR_MASTER.Vendor_Code ='" + qry2 + "'"
            Dim dtVendor As DataTable = clsDBFuncationality.GetDataTable(qry4, trans)
            If dtVendor IsNot Nothing AndAlso dtVendor.Rows.Count > 0 Then
                objVendInv.Terms_Code = clsCommon.myCstr(dtVendor.Rows(0)("Terms_Code"))
                objVendInv.Terms_Description = clsCommon.myCstr(dtVendor.Rows(0)("Terms_Desc"))
                objVendInv.Due_Date = obj.Document_Date.AddDays(clsCommon.myCdbl(dtVendor.Rows(0)("No_Days")))
                objVendInv.Account_Set = clsCommon.myCstr(dtVendor.Rows(0)("Vendor_Account"))
                objVendInv.GSTRegistered = clsCommon.myCDecimal(dtVendor.Rows(0)("GSTRegistered"))
                objVendInv.Vendor_Control_AC = clsCommon.myCstr(dtVendor.Rows(0)("Payable_Account"))
            Else
                Throw New Exception("Please define vendor account set for vendor [" + qry2 + "] ")
            End If
            objVendInv.Vendor_Code = qry2
            objVendInv.Vendor_Name = qry3
            ' objVendInv.MCC_Code = obj.MCC_Code
            'objVendInv.MCC_Name = obj.MCC_Name
            objVendInv.Arr = New List(Of clsVedorInvoiceDetail)

            objVendInv.Arr.Add(objVendInvTR)
            objVendInv.SaveData(objVendInv, True, trans)
            clsVedorInvoiceHead.PostData("", objVendInv.Document_No, "", trans)

            'objVendInv.Document_No = ""
            'objVendInv.SaveData(objVendInv, True, trans)
            'clsVedorInvoiceHead.PostData("", objVendInv.Document_No, "", trans)
        Next
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
            Dim Qry As String = "Select Posted_Date from TSPL_BMC_TRANSPORTER_BILL_HEAD WHERE Document_Code='" + strDocNo + "'"
            If clsCommon.myLen(clsDBFuncationality.getSingleValue(Qry, trans)) <= 0 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            Dim obj As clsBMCTransporterBill = clsBMCTransporterBill.GetData(strDocNo, NavigatorType.Current, False, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_Code) <= 0) Then
                Throw New Exception("No Data found to Reverse And UnPost")
            End If

            For Each objtr As clsBMCTransporterBillDetail In obj.Arr
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Document_No from tspl_vendor_invoice_head where RefDocNo= '" + clsCommon.myCstr(objtr.Document_Code) + "'", trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim strAPDocCode As String = clsCommon.myCstr(dr("Document_No"))
                        If clsCommon.myLen(strAPDocCode) > 0 Then
                            '            Dim dtCheck As DataTable = clsDBFuncationality.GetDataTable("select Doc_No from TSPL_PAYMENT_PROCESS_DEDUCTION where AP_Invoice_No='" + strAPDocCode + "'", trans)
                            '            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                            '                Throw New Exception("Used In Payment Process No [" + clsCommon.myCstr(dtCheck.Rows(0)("Doc_No")) + "] in Deduction ")
                            '            End If
                            '            dtCheck = clsDBFuncationality.GetDataTable("select Doc_No from TSPL_PAYMENT_PROCESS_CREDIT_NOTE where AP_Invoice_No='" + strAPDocCode + "'", trans)
                            '            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                            '                Throw New Exception("Used In Payment Process No [" + clsCommon.myCstr(dtCheck.Rows(0)("Doc_No")) + "] in Addition")
                            '            End If
                            clsVedorInvoiceHead.ReverseAndUnpost(strAPDocCode, trans)
                            clsVedorInvoiceHead.DeleteData(strAPDocCode, trans)
                        End If
                    Next
                End If
            Next

                Qry = "Update TSPL_BMC_TRANSPORTER_BILL_HEAD set Posted_By=null,Posted_Date=NULL, Modify_By='" + objCommonVar.CurrentUserCode + "',Status=0 where Document_Code='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsBMCTransporterBillDetail

#Region "Variables"
    Public PK_ID As Integer = 0
    Public Document_Code As String = Nothing
    Public Amount As Decimal = 0
    Public MCC_Document_Code As String = Nothing
    Public Station_1 As String = Nothing
    Public Station_2 As String = Nothing
    Public Station_3 As String = Nothing
    Public Station_4 As String = Nothing
    Public Category As String = Nothing
    Public Trip As Decimal = 0
    Public GPS_KM As Decimal = 0
    Public KM As Decimal = 0
    Public Quantity_KG As Decimal = 0
    Public Diesel_RD As Decimal = 0
    Public BMC_Date As String = Nothing
    Public Ice_Box As String = Nothing
    Public arr As List(Of clsfrmVLCMaster) = Nothing

    'Public BalanceAmount As Decimal = 0
#End Region
    Public Shared Function SaveData1(ByVal strDocNo As String, ByVal isNewEntry As Boolean, ByVal obj As clsBMCTransporterBillDetail, ByVal arr As List(Of clsBMCTransporterBillDetail), ByVal trans As SqlTransaction) As Boolean
        If (arr IsNot Nothing AndAlso arr.Count > 0) Then
            For Each obj1 As clsBMCTransporterBillDetail In arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "MCC_Document_Code", obj.MCC_Document_Code)
                clsCommon.AddColumnsForChange(coll, "Station_1", obj.Station_1)
                clsCommon.AddColumnsForChange(coll, "Station_2", obj.Station_2)
                clsCommon.AddColumnsForChange(coll, "Station_3", obj.Station_3)
                clsCommon.AddColumnsForChange(coll, "Station_4", obj.Station_4)
                clsCommon.AddColumnsForChange(coll, "Trip", obj.Trip)
                clsCommon.AddColumnsForChange(coll, "GPS_KM", obj.GPS_KM)
                clsCommon.AddColumnsForChange(coll, "KM", obj.KM)
                clsCommon.AddColumnsForChange(coll, "Quantity_KG", obj.Quantity_KG)
                clsCommon.AddColumnsForChange(coll, "Diesel_RD", obj.Diesel_RD)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Document_Date", obj.BMC_Date)
                clsCommon.AddColumnsForChange(coll, "Ice_Box", obj.Ice_Box)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_BMC_TRANSPORTER_BILL_DETAIL", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If
        Return True

    End Function



    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsBMCTransporterBillDetail), ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each obj As clsBMCTransporterBillDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "MCC_Document_Code", obj.MCC_Document_Code, True)
                clsCommon.AddColumnsForChange(coll, "Station_1", obj.Station_1)
                clsCommon.AddColumnsForChange(coll, "Station_2", obj.Station_2)
                clsCommon.AddColumnsForChange(coll, "Station_3", obj.Station_3)
                clsCommon.AddColumnsForChange(coll, "Station_4", obj.Station_4)
                clsCommon.AddColumnsForChange(coll, "Category", obj.Category)
                clsCommon.AddColumnsForChange(coll, "Trip", obj.Trip)
                clsCommon.AddColumnsForChange(coll, "GPS_KM", obj.GPS_KM)
                clsCommon.AddColumnsForChange(coll, "KM", obj.KM)
                clsCommon.AddColumnsForChange(coll, "Quantity_KG", obj.Quantity_KG)
                clsCommon.AddColumnsForChange(coll, "Diesel_RD", obj.Diesel_RD)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.BMC_Date, "dd/MMM/yyyy"))
                clsCommon.AddColumnsForChange(coll, "Ice_Box", obj.Ice_Box)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_BMC_TRANSPORTER_BILL_DETAIL", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If
        Dim qry As String = "select Tanker_No,Document_Date,Trip,sum(1) as Rep from (
select TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No,CONVERT(Date, TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Date,103) as Document_Date,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Trip,  case when TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code='" & strDocNo & "' then 1 else 0 end as chk
from TSPL_BMC_TRANSPORTER_BILL_DETAIL 
left outer join TSPL_BMC_TRANSPORTER_BILL_HEAD on TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code=TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code
) xx group by Tanker_No,Document_Date,Trip having sum(chk)>0 and sum(1)>1"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Throw New Exception("Repeated Entry found on [" & clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")), "dd/MM/yyyy") & "] for Trip [" + clsCommon.myCstr(dt.Rows(0)("Trip")) + "]")
        End If
        Return True
    End Function

End Class
