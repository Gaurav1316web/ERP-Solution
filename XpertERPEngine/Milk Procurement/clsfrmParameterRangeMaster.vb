Imports common
Imports System.Data.SqlClient
Public Class clsfrmParameterRangeMaster
#Region "Variables"
    Public Procurement_Type As String = Nothing
    Public MIKL_TYPE_CODE As String = Nothing
    Public code As String = Nothing
    Public desc As String = Nothing
    Public Lrange As Decimal = Nothing
    Public Urange As Decimal = Nothing
    Public Value As String = Nothing
    Public Eff_date As Date = Nothing
    Public End_date As Date? = Nothing
    Public Condition As String = Nothing
    Public Condition_Value As String = Nothing
    Public Status As String = Nothing
    Public Loc_Code As String = Nothing
    Public PK_Id As String = Nothing
    Public Vendor_Class As String = Nothing
    Public IsReject As Integer = 0
#End Region
    Public Shared Function SaveData(ByVal arr As List(Of clsfrmParameterRangeMaster), ByVal trans As SqlTransaction, ByVal isImport As Boolean) As Boolean
        Try
            'Dim whrCls As String = ""
            'If Not clsMccMaster.isCurrentUserHO(trans) Then
            '    If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            '        whrCls = " and loc_code in (" & objCommonVar.strCurrUserLocations & "  )"
            '    End If
            'End If
            Dim isSaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                'If Not isImport Then
                '    Dim qry As String = "delete from tspl_parameter_range_master where 1=1 " & whrCls
                '    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'End If
                For Each obj As clsfrmParameterRangeMaster In arr
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "comp_code", objCommonVar.CurrentCompanyCode)
                    'clsCommon.AddColumnsForChange(coll, "PK_Id", obj.PK_Id)
                    clsCommon.AddColumnsForChange(coll, "code", obj.code)
                    clsCommon.AddColumnsForChange(coll, "lower_range", obj.Lrange)
                    clsCommon.AddColumnsForChange(coll, "upper_range", obj.Urange)
                    clsCommon.AddColumnsForChange(coll, "value", obj.Value)
                    clsCommon.AddColumnsForChange(coll, "Loc_Code", obj.Loc_Code, True)
                    'clsCommon.AddColumnsForChange(coll, "Condition", obj.Condition)
                    clsCommon.AddColumnsForChange(coll, "Condition_Value", obj.Condition_Value)
                    clsCommon.AddColumnsForChange(coll, "Status", obj.Status)
                    clsCommon.AddColumnsForChange(coll, "IsReject", clsCommon.myCdbl(obj.IsReject))
                    clsCommon.AddColumnsForChange(coll, "effective_date", clsCommon.GetPrintDate(obj.Eff_date, "dd/MMM/yyyy"))
                    If obj.End_date Is Nothing Then
                        clsCommon.AddColumnsForChange(coll, "End_date", Nothing, True)
                    Else
                        clsCommon.AddColumnsForChange(coll, "End_date", clsCommon.GetPrintDate(obj.End_date, "dd/MMM/yyyy"))
                    End If

                    clsCommon.AddColumnsForChange(coll, "created_by", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "created_date", clsCommon.myCstr(clsCommon.GETSERVERDATE(trans).ToString("dd/MM/yyyy")))
                    clsCommon.AddColumnsForChange(coll, "modified_by", objCommonVar.CurrentUserCode)
                    clsCommon.AddColumnsForChange(coll, "modified_date", clsCommon.myCstr(clsCommon.GETSERVERDATE(trans).ToString("dd/MM/yyyy")))
                    clsCommon.AddColumnsForChange(coll, "Vendor_Class", obj.Vendor_Class)
                    clsCommon.AddColumnsForChange(coll, "MIKL_TYPE_CODE", obj.MIKL_TYPE_CODE, True)
                    clsCommon.AddColumnsForChange(coll, "Procurement_Type", obj.Procurement_Type)
                    'Dim qry As String = "delete from tspl_parameter_range_master where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code='" + clsCommon.myCstr(obj.code) + "' and lower_range='" + clsCommon.myCstr(obj.Lrange) + "' and upper_range='" + clsCommon.myCstr(obj.Urange) + "' and value='" + clsCommon.myCstr(obj.Value) + "'"
                    '' and effective_date='" + clsCommon.myCstr(Convert.ToDateTime(obj.Eff_date).ToString("dd/MMM/yyyy")) + "'
                    'Dim qry1 As String = " select count(*) from tspl_parameter_range_master where code='" & obj.code & "' and Lower_Range='" & obj.Lrange & "' and upper_Range='" & obj.Lrange & "' and Value='" & obj.Value & "' and Condition_value='" & obj.Condition_Value & "' and effective_date='" & obj.Eff_date & "' and status='" & obj.Status & "'"
                    'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry1, trans) >= 0) Then

                    'Else

                    'End If
                    If clsCommon.myLen(obj.PK_Id) > 0 Then
                        isSaved = SaveDataForHistory(obj.PK_Id, trans)
                        isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PARAMETER_RANGE_MASTER", OMInsertOrUpdate.Update, "TSPL_PARAMETER_RANGE_MASTER.pk_Id='" & obj.PK_Id & "'", trans)
                    Else
                        isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PARAMETER_RANGE_MASTER", OMInsertOrUpdate.Insert, "", trans)
                    End If
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.PK_Id, "TSPL_PARAMETER_RANGE_MASTER", "pk_Id", trans)

                Next
                'trans.Commit()
            End If

            Return True
        Catch ex As Exception
            Try
                trans.Rollback()
            Catch ex1 As Exception
            End Try
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function SaveDataForHistory(ByVal strCode As String, Optional ByVal trans As SqlTransaction = Nothing) As Boolean
        Dim isSaved As Boolean = True
        Try
            Dim strHistoryDate As DateTime = clsCommon.GETSERVERDATE(trans)
            '' TSPL_INVOICE_MASTER_BULKSALE_History
            Dim qry As String = String.Empty
            Dim strInvColumns As String = clsERPFuncationality.GetTableColumnNameForQry("tspl_parameter_range_master", trans)
            qry = "INSERT INTO tspl_parameter_range_master_history (" + strInvColumns + ") SELECT  " + strInvColumns + " FROM tspl_parameter_range_master WHERE PK_Id='" + clsCommon.myCstr(strCode) + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            ''
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return isSaved
    End Function
End Class

Public Class clsfrmParameterMaster
#Region "Variables"
    Public nature As String = Nothing
    Public code As String = Nothing
    Public desc As String = Nothing
    Public type As String = Nothing
    Public Param_For As String = Nothing
    Public IsMandatory As Integer = 0
    Public IsCanType As Integer = 0
    Public IsBulkSale As Integer = 0
    Public IsMCC_Qc As Integer = 0
    Public IsForPrintOnQC As Integer = 0
    Public IsForPrintOnDispatch As Integer = 0
    Public IsForMilkGrade As Integer = 0
    Public Parametertype As String = Nothing
    Public IsProduction As Integer = 0
    Public IsProductionMandatory As Integer = 0
    Public Is_Milk_Sample As Boolean = False
#End Region

    Public Shared Function isCLRParmExist() As Boolean
        Dim rValue As Boolean = False
        Dim whrCls As String = String.Empty
        If clsERPFuncationality.isCurrentUserMCC Then
            whrCls = " where Param_for='MCC' or Param_for='BOTH'"
        Else
            whrCls = " where Param_for='PLANT' or Param_for='BOTH'"
        End If
        Dim qry As String = " select count(*) from tspl_parameter_Master  " & whrCls & " and type='CLR'"

        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) > 0 Then
            rValue = True
        Else
            rValue = False
        End If
        Return rValue
    End Function
    Public Shared Function isFATParmExist() As Boolean
        Return isFATParmExist(False)
    End Function

    Public Shared Function isFATParmExist(ByVal isForProduction As Boolean) As Boolean
        Dim rValue As Boolean = False
        Dim whrCls As String = String.Empty
        If isForProduction Then
            whrCls = " where IsProduction=1"
        Else
            If clsERPFuncationality.isCurrentUserMCC Then
                whrCls = " where Param_for='MCC' or Param_for='BOTH'"
            Else
                whrCls = " where Param_for='PLANT' or Param_for='BOTH'"
            End If
        End If
        
        Dim qry As String = " select count(*) from tspl_parameter_Master  " & whrCls & " and type='FAT'"

        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) > 0 Then
            rValue = True
        Else
            rValue = False
        End If
        Return rValue
    End Function

    Public Shared Function isSNFParmExist() As Boolean
        Return isSNFParmExist(False)
    End Function
    Public Shared Function isSNFParmExist(ByVal isForProduction As Boolean) As Boolean
        Dim rValue As Boolean = False
        Dim whrCls As String = String.Empty
        If isForProduction Then
            whrCls = " where IsProduction=1"
        Else
            If clsERPFuncationality.isCurrentUserMCC Then
                whrCls = " where Param_for='MCC' or Param_for='BOTH'"
            Else
                whrCls = " where Param_for='PLANT' or Param_for='BOTH'"
            End If
        End If
        Dim qry As String = " select count(*) from tspl_parameter_Master  " & whrCls & " and type='SNF'"

        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) > 0 Then
            rValue = True
        Else
            rValue = False
        End If
        Return rValue
    End Function
    Public Shared Function isCFParmExist() As Boolean
        Dim rValue As Boolean = False
        Dim whrCls As String = String.Empty
        If clsERPFuncationality.isCurrentUserMCC Then
            whrCls = " where Param_for='MCC' or Param_for='BOTH'"
        Else
            whrCls = " where Param_for='PLANT' or Param_for='BOTH'"
        End If
        Dim qry As String = " select count(*) from tspl_parameter_Master  " & whrCls & " and type='CF'"

        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) > 0 Then
            rValue = True
        Else
            rValue = False
        End If
        Return rValue
    End Function
    Public Shared Function SaveData(ByVal strCode As String, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction, ByVal obj As clsfrmParameterMaster) As Boolean
        Try
            Dim isSaved As Boolean = True

            If isNewEntry Then
                If clsCommon.myLen(obj.code) <= 0 Then
                    obj.code = clsERPFuncationality.GetNextCode(trans, clsCommon.GETSERVERDATE(trans), clsDocType.ParamMaster, "", "")
                End If
            End If

            strCode = obj.code

            Dim coll As New Hashtable()

            clsCommon.AddColumnsForChange(coll, "comp_code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Code", strCode)
            clsCommon.AddColumnsForChange(coll, "description", obj.desc)
            clsCommon.AddColumnsForChange(coll, "type", obj.type)
            clsCommon.AddColumnsForChange(coll, "Nature", obj.nature)
            clsCommon.AddColumnsForChange(coll, "Param_For", obj.Param_For)
            ''richa Against Ticket No. BM00000003713 on 03/09/2014
            clsCommon.AddColumnsForChange(coll, "IsMandatory", obj.IsMandatory)
            clsCommon.AddColumnsForChange(coll, "IsCanType", obj.IsCanType)
            clsCommon.AddColumnsForChange(coll, "IsBulkSale", obj.IsBulkSale)
            clsCommon.AddColumnsForChange(coll, "IsMCC_Qc", obj.IsMCC_Qc)
            clsCommon.AddColumnsForChange(coll, "IsForMilkGrade", obj.IsForMilkGrade)
            clsCommon.AddColumnsForChange(coll, "Parametertype", obj.Parametertype)
            clsCommon.AddColumnsForChange(coll, "IsProduction", obj.IsProduction)
            clsCommon.AddColumnsForChange(coll, "IsProductionMandatory", obj.IsProductionMandatory)
            clsCommon.AddColumnsForChange(coll, "Is_Milk_Sample", IIf(obj.Is_Milk_Sample, 1, 0))
            'If obj.IsForPrintOnDispatch = 1 Then
            '    clsCommon.AddColumnsForChange(coll, "IsForPrintOnDispatch", obj.IsForPrintOnDispatch)
            'End If
            'If obj.IsForPrintOnQC = 1 Then
            '    clsCommon.AddColumnsForChange(coll, "IsForPrintOnQC", obj.IsForPrintOnQC)
            'End If

            If clsCommon.CompairString(obj.Param_For, "Plant") = CompairStringResult.Equal Then
                clsCommon.AddColumnsForChange(coll, "IsForPrintOnDispatch", 0)
            End If
            ''----------------------------------------------------
            clsCommon.AddColumnsForChange(coll, "modified_by", clsCommon.myCstr(objCommonVar.CurrentUserCode))
            clsCommon.AddColumnsForChange(coll, "modified_date", clsCommon.myCstr(clsCommon.GETSERVERDATE(trans).ToString("dd/MM/yyyy")))

            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "created_by", clsCommon.myCstr(objCommonVar.CurrentUserCode))
                clsCommon.AddColumnsForChange(coll, "created_date", clsCommon.myCstr(clsCommon.GETSERVERDATE(trans).ToString("dd/MM/yyyy")))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PARAMETER_MASTER", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PARAMETER_MASTER", OMInsertOrUpdate.Update, " TSPL_PARAMETER_MASTER.Code='" + obj.code + "' and tspl_parameter_master.comp_code='" + objCommonVar.CurrentCompanyCode + "'", trans)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.code, "TSPL_PARAMETER_MASTER", "Code", trans)

            trans.Commit()
            Return True
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsfrmParameterMaster
        Dim obj As New clsfrmParameterMaster()
        Try
            Dim qry As String = "select * from tspl_parameter_master"

            Select Case NavType
                Case NavigatorType.Current
                    qry += " where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code='" + strCode + "'"
                Case NavigatorType.First
                    qry += " where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code in (select min(code) from tspl_parameter_master where comp_code='" + objCommonVar.CurrentCompanyCode + "')"
                Case NavigatorType.Last
                    qry += " where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code in (select max(code) from tspl_parameter_master where comp_code='" + objCommonVar.CurrentCompanyCode + "')"
                Case NavigatorType.Next
                    qry += " where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code in (select min(code) from tspl_parameter_master where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code>'" + strCode + "')"
                Case NavigatorType.Previous
                    qry += " where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code in (select max(code) from tspl_parameter_master where comp_code='" + objCommonVar.CurrentCompanyCode + "' and code<'" + strCode + "')"
            End Select

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            'Dim obj As New clsfrmParameterMaster()
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj.code = clsCommon.myCstr(dt.Rows(0)("code"))
                obj.desc = clsCommon.myCstr(dt.Rows(0)("description"))
                obj.type = clsCommon.myCstr(dt.Rows(0)("type"))
                obj.nature = clsCommon.myCstr(dt.Rows(0)("nature"))
                obj.Param_For = clsCommon.myCstr(dt.Rows(0)("Param_For"))
                obj.IsMandatory = clsCommon.myCdbl(dt.Rows(0)("IsMandatory"))
                obj.IsCanType = clsCommon.myCdbl(dt.Rows(0)("IsCanType"))
                obj.IsBulkSale = clsCommon.myCdbl(dt.Rows(0)("IsBulkSale"))
                obj.IsMCC_Qc = clsCommon.myCdbl(dt.Rows(0)("IsMCC_Qc"))
                obj.IsForPrintOnDispatch = clsCommon.myCdbl(dt.Rows(0)("IsForPrintOnDispatch"))
                obj.IsForPrintOnQC = clsCommon.myCdbl(dt.Rows(0)("IsForPrintOnQC"))
                obj.IsForMilkGrade = clsCommon.myCdbl(dt.Rows(0)("IsForMilkGrade"))
                obj.Parametertype = clsCommon.myCstr(dt.Rows(0)("Parametertype"))
                obj.IsProduction = clsCommon.myCdbl(dt.Rows(0)("IsProduction"))
                obj.IsProductionMandatory = clsCommon.myCdbl(dt.Rows(0)("IsProductionMandatory"))
                obj.Is_Milk_Sample = (clsCommon.myCdbl(dt.Rows(0)("Is_Milk_Sample")) = 1)
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return obj
    End Function

End Class

Public Class clsfrmTankerMaster
#Region "Variable"
    Public code As String = Nothing
    Public ESIRule As String = Nothing
    Public EPFRule As String = Nothing
    Public desc As String = Nothing
    Public tankerno As String = Nothing
    Public storagecap As Decimal = Nothing
    Public IceCharge As Decimal = Nothing
    Public year As String = Nothing
    Public inner As String = Nothing
    Public outer As String = Nothing
    Public tanker_name As String = Nothing
    Public shift_chrg As Decimal = Nothing
    Public avg_km_rate As Decimal = Nothing
    Public diesel_rate As Decimal = Nothing
    Public day_rental As Decimal = Nothing
    Public week_rental As Decimal = Nothing
    Public month_rental As Decimal = Nothing
    Public rate_km As Decimal = Nothing
    Public ltr_kg As Decimal = Nothing
    Public rate_ltr As Decimal = Nothing
    Public StorageCapacityDesc As String = Nothing
    Public RentalType As String = Nothing
    Public RentalAmount As Double = 0
    Public Rate_Type As String = String.Empty
    Public Price_Ltr_KG As Double = 0
    Public Labour_Charge As Double = 0
    Public Status As String = String.Empty
    Public Total_Chamber As Integer
    Public arrChamber As List(Of clsTankerChamberDetail) = Nothing
    Public Provision_Min_Qty As Integer
    Public Inactive As Boolean = False
    Public PrivateTanker As Boolean = False
#End Region

    Public Shared Function GetFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Try
            Dim qry As String = " select TSPL_TANKER_MASTER.Tanker_No as [TankerNo],TSPL_TANKER_MASTER.tanker_name as [Tanker Description],TSPL_TANKER_MASTER.tanker_transporter_code as [Transporter No],TSPL_VENDOR_MASTER.Vendor_Name as [Tanker Transporter Name],(TSPL_VENDOR_MASTER.Add1+' '+TSPL_VENDOR_MASTER.Add2+' '+TSPL_VENDOR_MASTER.Add3) as Address,TSPL_VENDOR_MASTER.City_Code_Desc as City,TSPL_VENDOR_MASTER.State,TSPL_VENDOR_MASTER.Country,(TSPL_VENDOR_MASTER.Phone1+' '+TSPL_VENDOR_MASTER.Phone2) as [Phone No],TSPL_VENDOR_MASTER.Email,TSPL_TANKER_MASTER.Storage_Capacity as [Storage Cap.(mt)],TSPL_TANKER_MASTER.StorageCapacityDesc as [Storage Capacity Description],TSPL_TANKER_MASTER.Year as [Year of Manufacturing],TSPL_TANKER_MASTER.Inner_SS as [Inner SS],TSPL_TANKER_MASTER.Outer_SS as [Outer SS],TSPL_TANKER_MASTER.Shift_Charges as [Charges per Day],TSPL_TANKER_MASTER.Avg_KM_Ltr as [Avg. KM per Ltr.],TSPL_TANKER_MASTER.Diesel_Rate as [Rate of Diesel],TSPL_TANKER_MASTER.Price_KM as [Rate per KM],TSPL_TANKER_MASTER.Rate_type as [Rate Type],TSPL_TANKER_MASTER.Price_Ltr_Kg as [Price Ltr/KG],TSPL_TANKER_MASTER.Rental_type as [Rental Type],TSPL_TANKER_MASTER.Rental_Amount as [Rental Amount],TSPL_TANKER_MASTER.Status from TSPL_VENDOR_MASTER right outer join TSPL_TANKER_MASTER on TSPL_TANKER_MASTER.Tanker_Transporter_Code=TSPL_VENDOR_MASTER.Vendor_Code"
            str = clsCommon.ShowSelectForm("TNDFND", qry, "TankerNo", whrcls, curcode, "TankerNo", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return str
    End Function

    Public Shared Function SaveData(ByVal tankerno As String, ByVal isNewEntry As Boolean, ByVal obj As clsfrmTankerMaster) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(tankerno, isNewEntry, obj, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function SaveData(ByVal tankerno As String, ByVal isNewEntry As Boolean, ByVal obj As clsfrmTankerMaster, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            clsDBFuncationality.ExecuteNonQuery("delete from TSPL_TANKER_CHAMBER_DETAIL where Tanker_No='" + obj.tankerno + "'", trans)
            If obj.Inactive Then
                Dim qry As String = "select ROUTE_NO from TSPL_bulk_route_master where Tanker_No='" + tankerno + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Throw New Exception("You can't Inactive Tanker No [" + tankerno + "] it is tagged with route no [" + clsCommon.myCstr(dt.Rows(0)("ROUTE_NO")) + "]")
                End If
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "comp_code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Tanker_Transporter_Code", obj.code)
            clsCommon.AddColumnsForChange(coll, "PFRULE_CODE", obj.EPFRule, True)
            clsCommon.AddColumnsForChange(coll, "ESIRULE_CODE", obj.ESIRule, True)
            clsCommon.AddColumnsForChange(coll, "Description", obj.desc)
            clsCommon.AddColumnsForChange(coll, "Tanker_No", obj.tankerno)
            clsCommon.AddColumnsForChange(coll, "Tanker_Name", obj.tanker_name)
            clsCommon.AddColumnsForChange(coll, "Storage_Capacity", obj.storagecap)
            clsCommon.AddColumnsForChange(coll, "Ice_Charge", obj.IceCharge)
            ''richa Against Ticket No. BM00000003713 on 03/09/2014
            clsCommon.AddColumnsForChange(coll, "StorageCapacityDesc", obj.StorageCapacityDesc)
            ''----------------------------------------------------
            clsCommon.AddColumnsForChange(coll, "Year", obj.year)
            clsCommon.AddColumnsForChange(coll, "Inner_SS", obj.inner)
            clsCommon.AddColumnsForChange(coll, "Outer_SS", obj.outer)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Shift_Charges", obj.shift_chrg)
            clsCommon.AddColumnsForChange(coll, "Avg_KM_Ltr", obj.avg_km_rate)
            clsCommon.AddColumnsForChange(coll, "Diesel_Rate", obj.diesel_rate)
            'clsCommon.AddColumnsForChange(coll, "Rental_Day", obj.day_rental)
            'clsCommon.AddColumnsForChange(coll, "Rental_Week", obj.week_rental)
            'clsCommon.AddColumnsForChange(coll, "Rental_Month", obj.month_rental)
            clsCommon.AddColumnsForChange(coll, "Rental_type", clsCommon.myCstr(obj.RentalType))
            clsCommon.AddColumnsForChange(coll, "Rental_Amount", clsCommon.myCdbl(obj.RentalAmount))
            clsCommon.AddColumnsForChange(coll, "Price_KM", obj.rate_km)
            'clsCommon.AddColumnsForChange(coll, "Price_Ltr", obj.rate_ltr)
            'clsCommon.AddColumnsForChange(coll, "Ltr_KG", obj.ltr_kg)
            clsCommon.AddColumnsForChange(coll, "Rate_Type", clsCommon.myCstr(obj.Rate_Type))
            clsCommon.AddColumnsForChange(coll, "Price_Ltr_KG", clsCommon.myCdbl(obj.Price_Ltr_KG))
            clsCommon.AddColumnsForChange(coll, "Labour_Charge", clsCommon.myCdbl(obj.Labour_Charge))
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.myCstr(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy")))
            clsCommon.AddColumnsForChange(coll, "Status", clsCommon.myCstr(obj.Status))

            clsCommon.AddColumnsForChange(coll, "Total_Chamber", obj.Total_Chamber)
            clsCommon.AddColumnsForChange(coll, "Provision_Min_Qty", obj.Provision_Min_Qty)
            clsCommon.AddColumnsForChange(coll, "Inactive", IIf(obj.Inactive, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Private_Tanker", IIf(obj.PrivateTanker, 1, 0))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.myCstr(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy")))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TANKER_MASTER", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TANKER_MASTER", OMInsertOrUpdate.Update, " TSPL_TANKER_MASTER.Tanker_No='" + obj.tankerno + "'", trans)
            End If
            clsTankerChamberDetail.SaveData(obj.tankerno, obj.arrChamber, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(obj.tankerno), "TSPL_TANKER_MASTER", "Tanker_No", "TSPL_TANKER_CHAMBER_DETAIL", "Tanker_No", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetData(ByVal tankerno As String, ByVal NavType As NavigatorType) As clsfrmTankerMaster
        Try
            Dim obj As New clsfrmTankerMaster()
            Dim qry As String = "select TSPL_TANKER_MASTER.tanker_transporter_code as Vendor_Code,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_TANKER_MASTER.Tanker_No,TSPL_TANKER_MASTER.tanker_name,TSPL_TANKER_MASTER.Storage_Capacity,TSPL_TANKER_MASTER.Ice_Charge,TSPL_TANKER_MASTER.StorageCapacityDesc,TSPL_TANKER_MASTER.Year,TSPL_TANKER_MASTER.Inner_SS,TSPL_TANKER_MASTER.Outer_SS,TSPL_TANKER_MASTER.Shift_Charges,TSPL_TANKER_MASTER.Avg_KM_Ltr,TSPL_TANKER_MASTER.Diesel_Rate,TSPL_TANKER_MASTER.Price_KM,TSPL_TANKER_MASTER.Rate_Type,TSPL_TANKER_MASTER.Price_Ltr_Kg,TSPL_TANKER_MASTER.Rental_Type,TSPL_TANKER_MASTER.Rental_Amount,TSPL_TANKER_MASTER.Status,TSPL_TANKER_MASTER.Total_Chamber ,TSPL_TANKER_MASTER.Provision_Min_Qty,TSPL_TANKER_MASTER.Inactive ,TSPL_TANKER_MASTER.Private_Tanker,TSPL_TANKER_MASTER.PFRULE_CODE,TSPL_TANKER_MASTER.ESIRULE_CODE,TSPL_TANKER_MASTER.Labour_Charge
from TSPL_VENDOR_MASTER 
right outer join TSPL_TANKER_MASTER on TSPL_TANKER_MASTER.Tanker_Transporter_Code=TSPL_VENDOR_MASTER.Vendor_Code "

            Select Case (NavType)
                Case NavigatorType.Current
                    qry += " where TSPL_TANKER_MASTER.tanker_no='" + tankerno + "' and TSPL_VENDOR_MASTER.Form_Type='TTM'"
                Case NavigatorType.First
                    qry += " where TSPL_TANKER_MASTER.tanker_no in (select min(TSPL_TANKER_MASTER.tanker_no) from TSPL_TANKER_MASTER) and TSPL_VENDOR_MASTER.Form_Type='TTM'"
                Case NavigatorType.Last
                    qry += " where TSPL_TANKER_MASTER.tanker_no in (select max(TSPL_TANKER_MASTER.tanker_no) from TSPL_TANKER_MASTER) and TSPL_VENDOR_MASTER.Form_Type='TTM'"
                Case NavigatorType.Next
                    qry += " where TSPL_TANKER_MASTER.tanker_no in (select min(TSPL_TANKER_MASTER.tanker_no) from TSPL_TANKER_MASTER where TSPL_TANKER_MASTER.tanker_no>'" + tankerno + "') and TSPL_VENDOR_MASTER.Form_Type='TTM'"
                Case NavigatorType.Previous
                    qry += " where TSPL_TANKER_MASTER.tanker_no in (select max(TSPL_TANKER_MASTER.tanker_no) from TSPL_TANKER_MASTER where TSPL_TANKER_MASTER.tanker_no<'" + tankerno + "') and TSPL_VENDOR_MASTER.Form_Type='TTM'"
            End Select

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj.code = clsCommon.myCstr(dt.Rows(0)("vendor_code"))
                obj.desc = clsCommon.myCstr(dt.Rows(0)("vendor_name"))
                obj.tankerno = clsCommon.myCstr(dt.Rows(0)("tanker_no"))
                obj.tanker_name = clsCommon.myCstr(dt.Rows(0)("tanker_name"))
                obj.storagecap = clsCommon.myCdbl(dt.Rows(0)("storage_capacity"))
                obj.IceCharge = clsCommon.myCdbl(dt.Rows(0)("Ice_Charge"))
                obj.shift_chrg = clsCommon.myCdbl(dt.Rows(0)("Shift_Charges"))
                obj.avg_km_rate = clsCommon.myCdbl(dt.Rows(0)("Avg_KM_Ltr"))
                obj.diesel_rate = clsCommon.myCdbl(dt.Rows(0)("Diesel_Rate"))
                'obj.day_rental = clsCommon.myCdbl(dt.Rows(0)("Rental_Day"))
                'obj.week_rental = clsCommon.myCdbl(dt.Rows(0)("Rental_Week"))
                'obj.month_rental = clsCommon.myCdbl(dt.Rows(0)("Rental_Month"))
                obj.RentalType = clsCommon.myCstr(dt.Rows(0)("Rental_Type"))
                obj.RentalAmount = clsCommon.myCdbl(dt.Rows(0)("Rental_Amount"))
                obj.Rate_Type = clsCommon.myCstr(dt.Rows(0)("Rate_type"))
                obj.Price_Ltr_KG = clsCommon.myCdbl(dt.Rows(0)("Price_Ltr_Kg"))
                obj.Labour_Charge = clsCommon.myCdbl(dt.Rows(0)("Labour_Charge"))
                obj.Status = clsCommon.myCstr(dt.Rows(0)("Status"))
                'obj.ltr_kg = clsCommon.myCdbl(dt.Rows(0)("Ltr_KG"))
                'obj.rate_ltr = clsCommon.myCdbl(dt.Rows(0)("Price_Ltr"))
                obj.rate_km = clsCommon.myCdbl(dt.Rows(0)("Price_KM"))
                ''richa Against Ticket No. BM00000003713 on 03/09/2014
                obj.StorageCapacityDesc = clsCommon.myCstr(dt.Rows(0)("StorageCapacityDesc"))
                obj.Provision_Min_Qty = clsCommon.myCdbl(dt.Rows(0)("Provision_Min_Qty"))
                ''----------------------------------------------------
                If clsCommon.myCstr(dt.Rows(0)("year")) Is DBNull.Value Or clsCommon.myCstr(dt.Rows(0)("year")) IsNot Nothing Then
                    obj.year = clsCommon.GETSERVERDATE().ToString("yyyy")
                Else
                    obj.year = clsCommon.myCstr(clsCommon.GetPrintDate(dt.Rows(0)("year"), "yyyy"))
                End If
                obj.inner = clsCommon.myCstr(dt.Rows(0)("inner_ss"))
                obj.outer = clsCommon.myCstr(dt.Rows(0)("outer_ss"))
                obj.Total_Chamber = clsCommon.myCdbl(dt.Rows(0)("Total_Chamber"))
                obj.Inactive = (clsCommon.myCDecimal(dt.Rows(0)("Inactive")) = 1)
                obj.PrivateTanker = (clsCommon.myCDecimal(dt.Rows(0)("Private_Tanker")) = 1)
                obj.arrChamber = clsTankerChamberDetail.GetData(obj.tankerno)
                obj.ESIRule = clsCommon.myCstr(dt.Rows(0)("ESIRULE_CODE"))
                obj.EPFRule = clsCommon.myCstr(dt.Rows(0)("PFRULE_CODE"))
            End If
            Return obj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class clsTankerChamberDetail
#Region "Variable"
    Public Tanker_No As String = Nothing
    Public Chamber_No As Integer = 0
    Public Chamber_Description As String = Nothing
#End Region

    Public Shared Function SaveData(ByVal strTankerNo As String, ByVal arr As List(Of clsTankerChamberDetail), ByVal trans As SqlTransaction) As Boolean
        Try
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                For Each obj As clsTankerChamberDetail In arr
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Tanker_No", strTankerNo)
                    clsCommon.AddColumnsForChange(coll, "Chamber_No", obj.Chamber_No)
                    clsCommon.AddColumnsForChange(coll, "Chamber_Description", obj.Chamber_Description)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TANKER_CHAMBER_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strTankerNo), "TSPL_TANKER_CHAMBER_DETAIL", "Tanker_No", trans)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetData(ByVal strTankerNo As String) As List(Of clsTankerChamberDetail)
        Dim arr As List(Of clsTankerChamberDetail) = Nothing
        Try
            Dim qry As String = "select * FROM  TSPL_TANKER_CHAMBER_DETAIL where Tanker_No='" + strTankerNo + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arr = New List(Of clsTankerChamberDetail)
                For Each dr As DataRow In dt.Rows
                    Dim obj As New clsTankerChamberDetail
                    obj.Tanker_No = clsCommon.myCstr(dr("Tanker_No"))
                    obj.Chamber_No = clsCommon.myCstr(dr("Chamber_No"))
                    obj.Chamber_Description = clsCommon.myCstr(dr("Chamber_Description"))
                    arr.Add(obj)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return arr
    End Function
End Class

Public Class clsfrmVillageMaster
#Region "Variables"
    Public villcode As String = Nothing
    Public villname As String = Nothing
    Public citycode As String = Nothing
    Public cityname As String = Nothing
    Public statecode As String = Nothing
    Public statename As String = Nothing
    Public countrycode As String = Nothing
    Public countryname As String = Nothing
    Public add1 As String = Nothing
    Public add2 As String = Nothing
    Public pinno As String = Nothing

    Public Surveyor_Name As String = Nothing
    Public Survey_Date As DateTime? = Nothing
    Public Total_Population As Double = 0
    Public Tehsil As String = Nothing
    Public Total_Voting As Double = 0
    Public Pradhan_Name As String = Nothing
    Public Pradhan_Contact_No As String = Nothing
    Public Distance_From_Center As Double = 0
    Public Irrigation_Source As String = Nothing
    Public Village_Area As String = Nothing
    Public Distance_From_MCC As Double = 0
    Public Cow_In_Milk As Double = 0
    Public Buffalo_In_Milk As Double = 0
    Public Cow_Dry As Double = 0
    Public Buffalo_Dry As Double = 0
    Public Total_Animals As Double = 0
    Public Milk_Production_Per_Day_Cow As Double = 0
    Public Milk_Production_Per_Day_Buffalo As Double = 0
    Public Marketable_Surplus_Per_Day_Cow As Double = 0
    Public Marketable_Surplus_Per_Day_Buffalo As Double = 0
    Public Expected_Milk_Per_Day_Cow As Double = 0
    Public Expected_Milk_Per_Day_Buffalo As Double = 0
#End Region

    Public Shared Function SaveData(ByVal obj As clsfrmVillageMaster, ByVal isNewEntry As Boolean) As Boolean
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

    Public Shared Function SaveData(ByVal obj As clsfrmVillageMaster, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim coll As New Hashtable()
            If isNewEntry AndAlso clsCommon.myLen(obj.villcode) <= 0 Then
                obj.villcode = clsERPFuncationality.GetNextCode(trans, clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"), clsDocType.VILLAGEMASTER, "", "")
            End If
            clsCommon.AddColumnsForChange(coll, "comp_code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "village_code", obj.villcode)
            clsCommon.AddColumnsForChange(coll, "village_name", obj.villname)
            clsCommon.AddColumnsForChange(coll, "add1", obj.add1)
            clsCommon.AddColumnsForChange(coll, "city_code", obj.citycode)
            clsCommon.AddColumnsForChange(coll, "state_code", obj.statecode)
            clsCommon.AddColumnsForChange(coll, "country_code", obj.countrycode)
            clsCommon.AddColumnsForChange(coll, "pin_no", obj.pinno)
            clsCommon.AddColumnsForChange(coll, "modified_by", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "modified_date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Surveyor_Name", obj.Surveyor_Name)
            If obj.Survey_Date Is Nothing Then
                clsCommon.AddColumnsForChange(coll, "Survey_Date", Nothing, True)
            Else
                clsCommon.AddColumnsForChange(coll, "Survey_Date", clsCommon.GetPrintDate(obj.Survey_Date, "dd/MMM/yyyy"))
            End If


            clsCommon.AddColumnsForChange(coll, "Total_Population", obj.Total_Population)
            clsCommon.AddColumnsForChange(coll, "Tehsil", obj.Tehsil)
            clsCommon.AddColumnsForChange(coll, "Total_Voting", obj.Total_Voting)
            clsCommon.AddColumnsForChange(coll, "Pradhan_Name", obj.Pradhan_Name)
            clsCommon.AddColumnsForChange(coll, "Pradhan_Contact_No", obj.Pradhan_Contact_No)
            clsCommon.AddColumnsForChange(coll, "Distance_From_Center", obj.Distance_From_Center)
            clsCommon.AddColumnsForChange(coll, "Irrigation_Source", obj.Irrigation_Source)
            clsCommon.AddColumnsForChange(coll, "Village_Area", obj.Village_Area)
            clsCommon.AddColumnsForChange(coll, "Distance_From_MCC", obj.Distance_From_MCC)
            clsCommon.AddColumnsForChange(coll, "Cow_In_Milk", obj.Cow_In_Milk)
            clsCommon.AddColumnsForChange(coll, "Buffalo_In_Milk", obj.Buffalo_In_Milk)
            clsCommon.AddColumnsForChange(coll, "Cow_Dry", obj.Cow_Dry)
            clsCommon.AddColumnsForChange(coll, "Buffalo_Dry", obj.Buffalo_Dry)
            clsCommon.AddColumnsForChange(coll, "Total_Animals", obj.Cow_In_Milk + obj.Buffalo_In_Milk + obj.Cow_Dry + obj.Buffalo_Dry)
            clsCommon.AddColumnsForChange(coll, "Milk_Production_Per_Day_Cow", obj.Milk_Production_Per_Day_Cow)
            clsCommon.AddColumnsForChange(coll, "Milk_Production_Per_Day_Buffalo", obj.Milk_Production_Per_Day_Buffalo)
            clsCommon.AddColumnsForChange(coll, "Marketable_Surplus_Per_Day_Cow", obj.Marketable_Surplus_Per_Day_Cow)
            clsCommon.AddColumnsForChange(coll, "Marketable_Surplus_Per_Day_Buffalo", obj.Marketable_Surplus_Per_Day_Buffalo)
            clsCommon.AddColumnsForChange(coll, "Expected_Milk_Per_Day_Cow", obj.Expected_Milk_Per_Day_Cow)
            clsCommon.AddColumnsForChange(coll, "Expected_Milk_Per_Day_Buffalo", obj.Expected_Milk_Per_Day_Buffalo)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "created_by", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "created_date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VILLAGE_MASTER", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VILLAGE_MASTER", OMInsertOrUpdate.Update, " TSPL_VILLAGE_MASTER.Village_code='" + obj.villcode + "'", trans)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.villcode, "TSPL_VILLAGE_MASTER", "Village_code", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsfrmVillageMaster
        Try
            Dim qry As String = "select TSPL_VILLAGE_MASTER.*,tspl_city_master.city_name,tspl_state_master.state_name,tspl_country_master.country_name from TSPL_VILLAGE_MASTER left outer join tspl_city_master on tspl_city_master.city_code=TSPL_VILLAGE_MASTER.city_code left outer join tspl_state_master on tspl_state_master.state_code=TSPL_VILLAGE_MASTER.state_code left outer join tspl_country_master on tspl_country_master.country_code=TSPL_VILLAGE_MASTER.country_code"
            Select Case NavType
                Case NavigatorType.Current
                    qry += " where TSPL_VILLAGE_MASTER.village_code='" + strCode + "'"
                Case NavigatorType.First
                    qry += " where TSPL_VILLAGE_MASTER.village_code  in (select min(t.village_code) from tspl_village_master t)"
                Case NavigatorType.Last
                    qry += " where TSPL_VILLAGE_MASTER.village_code  in (select max(t.village_code) from tspl_village_master t)"
                Case NavigatorType.Previous
                    qry += " where TSPL_VILLAGE_MASTER.village_code  in (select max(t.village_code) from tspl_village_master t where t.village_code<'" + strCode + "')"
                Case NavigatorType.Next
                    qry += " where TSPL_VILLAGE_MASTER.village_code  in (select min(t.village_code) from tspl_village_master t where t.village_code>'" + strCode + "')"
            End Select
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim obj As New clsfrmVillageMaster()

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj.villcode = clsCommon.myCstr(dt.Rows(0)("village_code"))
                obj.villname = clsCommon.myCstr(dt.Rows(0)("village_name"))
                obj.add1 = clsCommon.myCstr(dt.Rows(0)("add1"))
                'obj.add2 = clsCommon.myCstr(dt.Rows(0)("add2"))
                obj.citycode = clsCommon.myCstr(dt.Rows(0)("city_code"))
                obj.cityname = clsCommon.myCstr(dt.Rows(0)("city_name"))
                obj.statecode = clsCommon.myCstr(dt.Rows(0)("state_code"))
                obj.statename = clsCommon.myCstr(dt.Rows(0)("state_name"))
                obj.countrycode = clsCommon.myCstr(dt.Rows(0)("country_code"))
                obj.countryname = clsCommon.myCstr(dt.Rows(0)("country_name"))
                obj.pinno = clsCommon.myCstr(dt.Rows(0)("pin_no"))


                obj.Surveyor_Name = clsCommon.myCstr(dt.Rows(0)("Surveyor_Name"))
                If dt.Rows(0)("Survey_Date") IsNot DBNull.Value Then
                    obj.Survey_Date = clsCommon.myCstr(dt.Rows(0)("Survey_Date"))
                End If
                obj.Total_Population = clsCommon.myCdbl(dt.Rows(0)("Total_Population"))
                obj.Tehsil = clsCommon.myCstr(dt.Rows(0)("Tehsil"))
                obj.Total_Voting = clsCommon.myCdbl(dt.Rows(0)("Total_Voting"))
                obj.Pradhan_Name = clsCommon.myCstr(dt.Rows(0)("Pradhan_Name"))
                obj.Pradhan_Contact_No = clsCommon.myCstr(dt.Rows(0)("Pradhan_Contact_No"))
                obj.Distance_From_Center = clsCommon.myCdbl(dt.Rows(0)("Distance_From_Center"))
                obj.Irrigation_Source = clsCommon.myCstr(dt.Rows(0)("Irrigation_Source"))
                obj.Village_Area = clsCommon.myCstr(dt.Rows(0)("Village_Area"))
                obj.Distance_From_MCC = clsCommon.myCdbl(dt.Rows(0)("Distance_From_MCC"))
                obj.Cow_In_Milk = clsCommon.myCdbl(dt.Rows(0)("Cow_In_Milk"))
                obj.Buffalo_In_Milk = clsCommon.myCdbl(dt.Rows(0)("Buffalo_In_Milk"))
                obj.Cow_Dry = clsCommon.myCdbl(dt.Rows(0)("Cow_Dry"))
                obj.Buffalo_Dry = clsCommon.myCdbl(dt.Rows(0)("Buffalo_Dry"))
                obj.Total_Animals = clsCommon.myCdbl(dt.Rows(0)("Total_Animals"))
                obj.Milk_Production_Per_Day_Cow = clsCommon.myCdbl(dt.Rows(0)("Milk_Production_Per_Day_Cow"))
                obj.Milk_Production_Per_Day_Buffalo = clsCommon.myCdbl(dt.Rows(0)("Milk_Production_Per_Day_Buffalo"))
                obj.Marketable_Surplus_Per_Day_Cow = clsCommon.myCdbl(dt.Rows(0)("Marketable_Surplus_Per_Day_Cow"))
                obj.Marketable_Surplus_Per_Day_Buffalo = clsCommon.myCdbl(dt.Rows(0)("Marketable_Surplus_Per_Day_Buffalo"))
                obj.Expected_Milk_Per_Day_Cow = clsCommon.myCdbl(dt.Rows(0)("Expected_Milk_Per_Day_Cow"))
                obj.Expected_Milk_Per_Day_Buffalo = clsCommon.myCdbl(dt.Rows(0)("Expected_Milk_Per_Day_Buffalo"))
            End If
            Return obj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class