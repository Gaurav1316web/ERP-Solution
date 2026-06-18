Imports common
Imports System.Data.SqlClient
Public Class clsTransferDCC
#Region "Variables"
    Public Requisition_Id As String = Nothing
    Public GLVoucher_No As String = Nothing
    Public Is_MandiTax As Integer = 0
    Public For_Repair As Integer = 0
    Public TransferIndent_No As String = String.Empty
    Public InternalTransfer As Integer = 0
    Public JobWorkTransfer As Integer = 0
    Public ProdRequestTransfer As Integer = 0
    Public IsJobWorkType As Integer = 0
    Public Electronic_Ref_No As String = Nothing
    Public EWayBillDate As Date?
    Public EWayBillNo As String = Nothing
    Public GSTStatus As Boolean = False
    Public Is_Taxable As Integer = 0
    Public Freight_Type As String = Nothing
    Public EmptyCharge As Decimal = Nothing
    Public FixedCharge As Decimal = Nothing
    Public Loading_Advice_No As String = String.Empty
    Public Vehicle_Charge As Decimal = Nothing
    Public Vehicle_Capacity As Decimal = Nothing
    Public Total_Item_Wt As Decimal = Nothing
    Public Gross_Item_Wt As Decimal = Nothing
    Public Vehicle_Mannual_No As String = Nothing
    Public IsEwayBill As Integer = 0

    Public Form38 As Boolean = False
    Public Document_No As String = Nothing
    Public Document_Date As DateTime = Nothing
    Public Delivery_date As String = Nothing
    Public Delivery_Duration As String = ""
    Public Transfer_Type As String = Nothing
    Public GP_Item_Type As String = Nothing
    Public Price_Code As String = Nothing
    Public Type As String = Nothing
    Public Item_Tax_Type As Integer = 0
    Public Status As ERPTransactionStatus = ERPTransactionStatus.Pending

    Public On_Hold As Boolean = Nothing
    Public Ref_No As String = Nothing
    Public Description As String = Nothing
    Public Remarks As String = Nothing
    Public From_Location As String = Nothing
    Public To_Location As String = Nothing
    Public To_Location_Main As String = Nothing
    Public From_LocationName As String = Nothing
    Public To_LocationName As String = Nothing
    Public Vehicle_Code As String = Nothing
    Public Vehicle_No As String = Nothing
    Public Km_Reading As String = Nothing
    Public Document_Type As String = Nothing
    Public GR_No As String = Nothing
    Public GR_Date As Date = Nothing
    Public Waybill_No As String = Nothing
    Public Waybill_Date As Date? = Nothing
    Public Transport_Id As String = Nothing
    Public Transporter_Name_Manual As String = Nothing

    Public Tax_Group As String = Nothing
    Public TaxGroupName As String = Nothing 'Not a table field
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
    Public Total_Tax_Amt As Double = 0
    Public Discount_Base As Double = 0
    Public Discount_Amt As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public Total_Amt_Less_Tax As Double = 0
    Public DOC_Total_Amt As Double = 0
    Public Mode_Of_Transport As String = Nothing
    Public Comments As String = Nothing
    Public Comp_Code As String = Nothing
    Public Terms_Code As String = Nothing
    Public TermsName As String = Nothing
    Public Terms_Remark As String = Nothing
    Public Due_Date As String = Nothing
    Public Posting_Date As DateTime?

    Public Modify_By As String = Nothing
    Public Modify_Date As String = Nothing
    Public Created_By As String = Nothing
    Public Created_Date As String = Nothing
    Public Is_AgainstFormF As Double = 0
    Public Tax_Calculation_Type As EnumTaxCalucationType
    Public TransferOutNo As String = Nothing
    Public Is_Status_IN As String = String.Empty
    Public Arr As List(Of clsTransferDCCDetail) = Nothing

    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public AutoTransfer As Double = 0
    Public RMDA_Code As String = ""

    'Private Shared strExcisable As String
    'Private Shared strFromLType As String
    'Private Shared strToLType As String
    '======================Added by Preeti Gupta================
    Public Crate_IN As Double = 0
    Public Box_IN As Double = 0
    Public jaali_IN As Double = 0
    Public Crate_Out As Double = 0
    Public jaali_Out As Double = 0
    Public Box_Out As Double = 0
    Public Secondary_Code As String = Nothing
    Public Freight_Distance As Integer = 0
#End Region


    ''Note Very Important If any change mad in PO Head or PO Detail table allso update it's History table.

    Public Shared Function CreateMultiplePOs(ByVal ArrPO As List(Of clsTransferDCC)) As List(Of String)
        Dim isSaved As Boolean = True
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Dim ArrCreatedPO As New List(Of String)()
        Try
            For Each obj As clsTransferDCC In ArrPO
                obj.SaveData(obj, True, False, trans)
                ArrCreatedPO.Add(obj.Document_No)
            Next
            '' Throw New Exception("Exception")
            trans.Commit()
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return ArrCreatedPO
    End Function


    Public Shared Function GetProvisionCharge(ByVal Frm_Loc_Code As String, ByVal to_loc_code As String, ByVal gross_wt As Decimal, ByVal Capacity As Decimal, ByVal Transport_Id As String) As DataTable
        Dim value As Decimal = 0
        Dim Ret_Dt As New DataTable()
        Ret_Dt.Columns.Add("FixedCharge", GetType(Decimal))
        Ret_Dt.Columns.Add("EmptyCharge", GetType(Decimal))
        Ret_Dt.Columns.Add("FreightCharge", GetType(Decimal))
        Ret_Dt.Columns.Add("FreightType", GetType(String))

        Dim dr As DataRow = Nothing
        Dim city As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select location_code from tspl_location_master where location_code='" + to_loc_code + "'"))

        Dim qry As String = "select top 1 capacitymt,freight,Fixed,Type from TSPL_ROUTE_FREIGHT_DETAILS where location_code='" + Frm_Loc_Code + "' and Tolocation_code='" + city + "' and capacitymt='" + clsCommon.myCstr(Capacity) + "' and transport_id='" + Transport_Id + "' and TransType='T' order by effective_date desc"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim capacitymt As Decimal = clsCommon.myCdbl(dt.Rows(0)("capacitymt"))
            Dim charge As Decimal = clsCommon.myCdbl(dt.Rows(0)("freight"))
            Dim Fixed As Decimal = clsCommon.myCdbl(dt.Rows(0)("Fixed"))
            Dim Freight_Type As String = clsCommon.myCstr(dt.Rows(0)("Type"))
            Dim FixedCharge As Double = Fixed
            Dim EmptyCharge As Double = charge

            'DONE FOR mt TO kg CONVERSION
            Dim Weight_MT_Unit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select description from TSPL_FIXED_PARAMETER where code='" + clsFixedParameterCode.VehicleCapacityUnit + "' and type='" + clsFixedParameterType.VehicleCapacityUnit + "'"))
            Dim Gross_Weight_Unit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select description from TSPL_FIXED_PARAMETER where code='" + clsFixedParameterCode.GrossWeightUnit + "' and type='" + clsFixedParameterType.GrossWeightUnit + "'"))
            qry = "select top 1 CF from (select (case when (Container_UOM='" & Gross_Weight_Unit & "' and Contained_UOM='" & Weight_MT_Unit & "') then round(Contained_Qty/Container_Qty,4) else case when (Container_UOM='" & Weight_MT_Unit & "' and Contained_UOM='" & Gross_Weight_Unit & "') then round(Container_Qty/Contained_Qty,4) end end) as CF,product_type from TSPL_WEIGHT_CONVERSION where product_type in ('ALL'))aa where isnull(cast(CF as varchar),'')<>'' order by Product_Type desc"
            Dim gross_uom_cnvrsn As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
            If clsCommon.CompairString(Weight_MT_Unit, Gross_Weight_Unit) = CompairStringResult.Equal Then
                gross_uom_cnvrsn = 1
            End If
            If gross_uom_cnvrsn <= 0 Then
                Throw New Exception("Provide weight conversion of unit [" + Gross_Weight_Unit + "] to [" + Weight_MT_Unit + "] at Weight conversion screen.")
            End If
            gross_wt = gross_wt * gross_uom_cnvrsn
            ''====================================================

            If gross_wt > capacitymt Then
                If charge > 0 Then
                    value = System.Math.Round((charge / capacitymt) * gross_wt, 2)
                Else
                    value = System.Math.Round(Fixed, 2)
                End If
            ElseIf gross_wt <= capacitymt Then
                value = charge + Fixed
            End If

            dr = Ret_Dt.NewRow()
            dr("FreightType") = Freight_Type
            dr("FixedCharge") = FixedCharge
            dr("FreightCharge") = value
            dr("EmptyCharge") = EmptyCharge
            Ret_Dt.Rows.Add(dr)
        End If

        Return Ret_Dt
    End Function
    '' TEPORARY WORK DUE TO UMVAALABOLITY OF CLSTAGGROUPMASTER CLASS
    Public Shared Function getFinder_ExcisableTaxGroup(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " SELECT distinct TSPL_TAX_GROUP_MASTER.TAX_Group_Code as 'Code'," &
                            " (CASE WHEN TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' THEN 'Sales' " &
                            " when TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' then  'Purchase' else 'Transfer' END) as 'Transaction Type', " &
                            " Tax_Group_Desc as Description,TSPL_TAX_GROUP_MASTER.Currency_Code From TSPL_TAX_GROUP_MASTER " &
                            " inner join TSPL_TAX_GROUP_DETAILS on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code " &
                            " inner join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=    TSPL_TAX_GROUP_DETAILS.Tax_Code "
        str = clsCommon.ShowSelectForm("RPTTXGRPFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function

    Public Shared Function GetTaxDetailsTransfer(ByVal GrpCode As String, Optional ByVal trans As SqlTransaction = Nothing) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='T') AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.Tax_Recoverable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='T' and TSPL_TAX_MASTER.type='E' order by Trans_Code"
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function

    Public Shared Function GetTaxGroupData(ByVal strCode As String, ByVal strType As String) As clsTaxGroupMaster
        Dim obj As clsTaxGroupMaster = Nothing
        Dim qry As String = "select Tax_Group_Code,Tax_Group_Desc,Tax_Group_Type,CURRENCY_CODE from TSPL_TAX_GROUP_MASTER Where Tax_Group_Code='" + strCode + "' and Tax_Group_Type='" + strType + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsTaxGroupMaster()
            obj.Tax_Group_Code = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Code"))
            obj.Tax_Group_Desc = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            obj.Tax_Group_Type = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Type"))
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))

            qry = "select  Trans_Code,Tax_Group_Code,Tax_Group_Type,Tax_Code,Tax_Code_Desc,Taxable,Surtax,Surtax_Tax_Code,Surtax_Tax_Code_Desc from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code='" + strCode + "'  and Tax_Group_Type='" + strType + "' order by Trans_Code"
            dt = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsTaxGroupDetail)
                Dim objTR As clsTaxGroupDetail = Nothing
                For Each dr As DataRow In dt.Rows
                    objTR = New clsTaxGroupDetail()
                    objTR.Trans_Code = clsCommon.myCdbl(dr("Trans_Code"))
                    objTR.Tax_Group_Code = clsCommon.myCstr(dr("Tax_Group_Code"))
                    objTR.Tax_Group_Type = clsCommon.myCstr(dr("Tax_Group_Type"))
                    objTR.Tax_Code = clsCommon.myCstr(dr("Tax_Code"))
                    objTR.Tax_Code_Desc = clsCommon.myCstr(dr("Tax_Code_Desc"))
                    objTR.Taxable = clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal
                    objTR.Surtax = clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal
                    objTR.Surtax_Tax_Code = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                    objTR.Surtax_Tax_Code_Desc = clsCommon.myCstr(dr("Surtax_Tax_Code_Desc"))

                    obj.Arr.Add(objTR)
                Next
            End If
        End If
        Return obj
    End Function
    ''RICHA UDL/14/05/18-000161
    Public Shared Function CancelData(ByVal FormId As String, ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            CancelData(FormId, strDocNo, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function CancelData(ByVal FormId As String, ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean

        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Transafer No not found to Cancel")
            End If
            Dim obj As clsTransferDCC = clsTransferDCC.GetData(strDocNo, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0) Then
                Throw New Exception("No Data found to Cancel")
            End If

            If obj.Status = ERPTransactionStatus.Approved Then
                Throw New Exception("Transaction is already Posted, so it should not be Cancelled.")
            End If
            If obj.Status = ERPTransactionStatus.Cancel Then
                Throw New Exception("Transaction is already Cancelled.")
            End If
            Dim qry As String = "Update TSPL_TRANSFER_ORDER_HEAD set Status=2, Posting_Date='" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") + "',Modify_By='" + objCommonVar.CurrentUserCode + "' where Document_No='" + obj.Document_No + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function CancelDelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal NavType As NavigatorType) As Boolean
        '' created by Sanjay
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            Dim obj As clsTransferDCC = clsTransferDCC.GetData(Doc_No, NavType, trans)

            If obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If

            ''richa agarwal 24 Dec,2020
            Dim dtirn As DataTable = clsDBFuncationality.GetDataTable("select Einvoice_type,IRN_No,Is_Taxable,From_Location from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & Doc_No & "'", trans)
            If dtirn IsNot Nothing AndAlso dtirn.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(obj.Transfer_Type), "O") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Is_Taxable), "1") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.IsJobWorkType), "0") = CompairStringResult.Equal Then
                    If ClsEInvoiceOFAPIs.EInvoice_Cancellation(Doc_No, clsCommon.myCstr(dtirn.Rows(0)("IRN_No")), clsCommon.myCstr(dtirn.Rows(0)("From_Location")), trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If
            ''----------

            clsItemLocationDetails.CheckCancelInventoryBalance(Form_Id, Doc_No, trans)
            '' transfer data into cancel table
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_TRANSFER_ORDER_HEAD", "Document_No", "TSPL_TRANSFER_ORDER_DETAIL", "Document_No", trans)


            '' cancel journal master data shipment
            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No='" & Doc_No & "'"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            '' cancel custom field data
            clsCommonFunctionality.SaveCancelDataMultKey(objCommonVar.CurrentUserCode, Doc_No, "TSPL_CUSTOM_FIELD_VALUES", "Transaction_Code", "Program_Code", Form_Id, trans)


            ''delete data from multiple tables


            qry = "delete from TSPL_CUSTOM_FIELD_VALUES where Transaction_Code in ('" & Doc_No & "') "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)


            qry = "delete from TSPL_BATCH_ITEM where  Document_Code='" & Doc_No & "' and Document_Type IN ('Transfer','ITransfer')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "' and Trans_Type IN ('Transfer','ITransfer')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)


            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No  ='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & Doc_No & "' "
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

    ''END  TEPORARY WORK DUE TO UMVAALABOLITY OF CLSTAGGROUPMASTER CLASS
    Public Function SaveData(ByVal obj As clsTransferDCC, ByVal isNewEntry As Boolean, ByVal isMakeAbandomentNo As Boolean) As Boolean
        Dim isSaved As Boolean = True
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(obj, isNewEntry, isMakeAbandomentNo, trans)
            trans.Commit()
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function

    Public Function SaveData(ByVal obj As clsTransferDCC, ByVal isNewEntry As Boolean, ByVal isMakeAbandomentNo As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Transfer Order", "Transfer Order", obj.From_Location, obj.TransferOrder_Date, trans)

            If Not isNewEntry Then
                Dim Status As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Status from TSPL_Transfer_ORDER_HEAD Where Document_No='" + obj.Document_No + "'", trans))
                If Status = 1 Then
                    Throw New Exception("This document is already posted.")
                End If
            End If
            Dim trans_Type As String = ""
            If obj.InternalTransfer = 1 Then
                trans_Type = "ITransfer"
            ElseIf obj.JobWorkTransfer = 1 Then
                trans_Type = "JTransfer"
            Else
                trans_Type = "Transfer"
            End If
            'clsSerializeInvenotry.DeleteData(trans_Type, obj.Document_No, trans)
            'clsBatchInventory.DeleteData(trans_Type, obj.Document_No, trans)
            'Dim qry As String = "delete from TSPL_Transfer_ORDER_DETAIL where Document_No='" + obj.Document_No + "'"
            'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Dim strDocNo As String = ""
            GSTStatus = clsERPFuncationality.GetGSTStatus(obj.Document_Date)
            If isNewEntry Then
                If GSTStatus Then
                    If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateCommonSeriesLocationwiseForAllSale, clsFixedParameterCode.CreateCommonSeriesLocationwiseForAllSale, trans)) = 0 Then
                        ' if common series setting is OFF
                        If obj.Is_Taxable = 1 Then
                            If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                    If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransfer, obj.From_Location)
                                    End If
                                ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)
                                End If
                                '=============Added by preeti Gupta=================
                            ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, obj.From_Location)
                                '============================================================

                            Else
                                Dim intExemptedType As Integer = clsDBFuncationality.getSingleValue("select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & obj.Tax_Group & "'", trans)
                                If clsCommon.CompairString(clsCommon.myCstr(intExemptedType), "1") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTBillofSupply, obj.From_Location)
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTTaxable, obj.From_Location)
                                End If
                            End If
                            obj.Document_Type = "Stock Transfer (Tax)"
                            'in case of transfer-in ,prefix generated on the basis of GIT ,main location.
                        Else
                            '======Sanjeet(31/01/2018)Dispatch Challan Code Generation For Repair Check=============
                            If obj.For_Repair = 1 Then
                                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)

                                    '===============Added by preeti gupta======================
                                ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, obj.From_Location)
                                    ''===========================================================================
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTNonTaxable, obj.From_Location)
                                End If
                                '===========================
                            Else
                                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                    If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                        Else
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransfer, obj.From_Location)
                                        End If
                                    ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)
                                    End If
                                    '===============Added by preeti gupta======================
                                ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.From_Location)
                                    ''===========================================================================

                                Else
                                    If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                                            If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)
                                            Else
                                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransferOut, obj.From_Location)
                                            End If
                                        End If
                                    ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)
                                        End If
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTNonTaxable, obj.From_Location)
                                    End If

                                End If

                                obj.Document_Type = "Stock Transfer (Local)"
                            End If
                        End If
                    Else
                        'if common series setting is ON
                        'Done by priti UDL/26/04/18-000131
                        If obj.Is_Taxable = 1 Then
                            If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                    If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransfer, obj.From_Location)
                                    End If
                                ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)

                                End If
                                '=============Added by preeti Gupta=================
                            ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, obj.From_Location)
                                '============================================================
                            Else
                                Dim intExemptedType As Integer = clsDBFuncationality.getSingleValue("select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & obj.Tax_Group & "'", trans)
                                If clsCommon.CompairString(clsCommon.myCstr(intExemptedType), "1") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.CommonSaleSeries, clsDocTransactionType.GSTBillofSupply, obj.From_Location)
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.CommonSaleSeries, clsDocTransactionType.GSTTaxable, obj.From_Location)
                                End If
                            End If
                            obj.Document_Type = "Stock Transfer (Tax)"
                            'in case of transfer-in ,prefix generated on the basis of GIT ,main location.
                        Else
                            '======Sanjeet(31/01/2018)Dispatch Challan Code Generation For Repair Check=============
                            If obj.For_Repair = 1 Then
                                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)
                                    '=============Added by preeti Gupta=================
                                ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, obj.From_Location)
                                    '============================================================
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.CommonSaleSeries, clsDocTransactionType.GSTDeliveryChallan, obj.From_Location)
                                End If
                                '===========================
                            Else
                                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                                    If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)
                                        Else
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransfer, obj.From_Location)
                                        End If
                                    ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransfer, obj.From_Location)

                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTIn, obj.To_Location)
                                    End If
                                    '=============Added by preeti Gupta=================
                                ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmTransferGST, clsDocTransactionType.GSTReturn, obj.From_Location)
                                    '============================================================

                                Else
                                    If clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                                            If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)
                                            Else
                                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransferOut, obj.From_Location)
                                            End If

                                        End If
                                    ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                        If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                                            obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)
                                        End If
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.CommonSaleSeries, clsDocTransactionType.GSTDeliveryChallan, obj.From_Location)
                                    End If

                                End If

                                obj.Document_Type = "Stock Transfer (Local)"
                            End If
                        End If
                    End If


                Else
                    Dim strExcise As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Excisable from TSPL_LOCATION_MASTER where Location_Code='" + obj.From_Location + "'", trans)) = "T", True, False)
                    If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then

                        If obj.Item_Tax_Type = 2 AndAlso strExcise = True Then   'If AllowedDockPosition Items Have Same Item_Tax_Type i.e. Excisable
                            'Dim strExcise As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Excisable from TSPL_LOCATION_MASTER where Location_Code='" + obj.From_Location + "'", trans)) = "T", True, False)
                            'If strExcise = False Then
                            '    Throw New Exception("Both Location and Item should be excisable.")
                            'End If
                            If clsCommon.myCBool(obj.InternalTransfer) = True Then
                                If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal Then
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.InternalTransferOut, obj.From_Location)
                                End If
                            ElseIf clsCommon.CompairString(obj.JobWorkTransfer, "1") = CompairStringResult.Equal Then
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.JWTransferOut, obj.From_Location)

                            Else
                                obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceExcise, obj.From_Location)
                            End If

                        Else
                            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TransferWithProductionSale_Retail_Series, obj.From_Location, trans)) = 10 And clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal Then
                                If clsCommon.myCBool(clsDBFuncationality.getSingleValue("Select Case When Excisable='T' Then 'True' Else 'False' End from TSPL_LOCATION_MASTER WHere Location_Code='" + obj.From_Location + "'", trans)) = True Then
                                    ''changes by richa agarwal Sale Invoice Series against ticket no BM00000005919 on 18/03/2015
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceTax, obj.From_Location)
                                    'obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceExcise, obj.From_Location)
                                    obj.Document_Type = "Tax Invoice"
                                Else
                                    ''changes by richa agarwal Sale Invoice Series against ticket no BM00000005919 on 18/03/2015
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceRetail, obj.From_Location)
                                    ' obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.frmSaleInvoiceProductSale, clsDocTransactionType.SaleInvoiceRetail, obj.From_Location)
                                    obj.Document_Type = "Retail Invoice"
                                End If
                            Else
                                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TransferLocalInterState, clsFixedParameterType.TransferLocalInterState, trans)) = 1 Then
                                    If clsDBFuncationality.getSingleValue("Select Case When (select State from TSPL_LOCATION_MASTER WHERE Location_Code='" & obj.From_Location & "')=(select State from TSPL_LOCATION_MASTER WHERE Location_Code='" & obj.To_Location & "') Then 'True' Else 'False' End", trans) = True Then
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferLocalOut, obj.From_Location)
                                        obj.Document_Type = "Stock Transfer (Local)"
                                    Else
                                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferInterStateOut, obj.From_Location)
                                        obj.Document_Type = "Stock Transfer (InterState)"
                                    End If
                                Else
                                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferOut, obj.From_Location)
                                    obj.Document_Type = "Stock Transfer"
                                End If
                            End If
                        End If
                    ElseIf clsCommon.CompairString(obj.Transfer_Type, "R") = CompairStringResult.Equal Then
                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferReject, obj.From_Location)
                        obj.Document_Type = "Stock Transfer"
                    ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then 'in case of transfer-in ,prefix generated on the basis of GIT ,main location.
                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferIn, obj.To_Location)
                        obj.Document_Type = "Stock Transfer"
                    Else
                        obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferIn, obj.From_Location)
                        obj.Document_Type = "Stock Transfer"
                    End If


                    ''=======================for secondary code=================
                    If obj.Item_Tax_Type = 2 AndAlso strExcise = True AndAlso clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.GenerateSecondryCode, clsFixedParameterCode.GenerateSecondryCode, trans)) = 1, True, False)) = True Then
                        If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TransferWithProductionSale_Retail_Series, obj.From_Location, trans)) = 10 And clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal Then
                                If clsCommon.myCBool(clsDBFuncationality.getSingleValue("Select Case When Excisable='T' Then 'True' Else 'False' End from TSPL_LOCATION_MASTER WHere Location_Code='" + obj.From_Location + "'", trans)) = True Then
                                    obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceTax, obj.From_Location)
                                Else
                                    obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.SNSaleInvoice, clsDocTransactionType.SaleInvoiceRetail, obj.From_Location)
                                End If
                            Else
                                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TransferLocalInterState, clsFixedParameterType.TransferLocalInterState, trans)) = 1 Then
                                    If clsDBFuncationality.getSingleValue("Select Case When (select State from TSPL_LOCATION_MASTER WHERE Location_Code='" & obj.From_Location & "')=(select State from TSPL_LOCATION_MASTER WHERE Location_Code='" & obj.To_Location & "') Then 'True' Else 'False' End", trans) = True Then
                                        obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferLocalOut, obj.From_Location)
                                    Else
                                        obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferInterStateOut, obj.From_Location)
                                    End If
                                Else
                                    obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferOut, obj.From_Location)
                                End If
                            End If
                        ElseIf clsCommon.CompairString(obj.Transfer_Type, "R") = CompairStringResult.Equal Then
                            obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferReject, obj.From_Location)
                        ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then 'in case of transfer-in ,prefix generated on the basis of GIT ,main location.
                            obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferIn, obj.To_Location)
                        Else
                            obj.Secondary_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.TransferDCC, clsDocTransactionType.TransferIn, obj.From_Location)
                        End If
                    End If
                    ''========================================================
                End If
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Item_Tax_Type", obj.Item_Tax_Type)
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Delivery_date", clsCommon.GetPrintDate(obj.Delivery_date, "dd/MM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Delivery_Duration", obj.Delivery_Duration)
            clsCommon.AddColumnsForChange(coll, "Transfer_Type", obj.Transfer_Type)
            clsCommon.AddColumnsForChange(coll, "IsEwayBill", obj.IsEwayBill)

            clsCommon.AddColumnsForChange(coll, "GP_Item_Type", obj.GP_Item_Type)
            clsCommon.AddColumnsForChange(coll, "Price_Code", obj.Price_Code)

            clsCommon.AddColumnsForChange(coll, "TransferIndent_No", obj.TransferIndent_No, True)
            clsCommon.AddColumnsForChange(coll, "On_Hold", IIf(obj.On_Hold, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Ref_No", obj.Ref_No)
            clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "From_Location", obj.From_Location)
            clsCommon.AddColumnsForChange(coll, "To_Location", obj.To_Location)
            clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
            clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
            clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
            clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
            clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
            clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
            clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
            clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
            clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
            clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
            clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
            clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
            clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
            clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", obj.Total_Tax_Amt)
            clsCommon.AddColumnsForChange(coll, "Discount_Base", obj.Discount_Base)
            clsCommon.AddColumnsForChange(coll, "Discount_Amt", obj.Discount_Amt)
            clsCommon.AddColumnsForChange(coll, "Amount_Less_Discount", obj.Amount_Less_Discount)
            clsCommon.AddColumnsForChange(coll, "Total_Amt_Less_Tax", obj.Total_Amt_Less_Tax)
            clsCommon.AddColumnsForChange(coll, "DOC_Total_Amt", obj.DOC_Total_Amt)
            clsCommon.AddColumnsForChange(coll, "Mode_Of_Transport", obj.Mode_Of_Transport)
            clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)

            clsCommon.AddColumnsForChange(coll, "Vehicle_Code", obj.Vehicle_Code)
            clsCommon.AddColumnsForChange(coll, "Vehicle_No", obj.Vehicle_No)
            clsCommon.AddColumnsForChange(coll, "Km_Reading", obj.Km_Reading)
            clsCommon.AddColumnsForChange(coll, "Is_AgainstFormF", obj.Is_AgainstFormF)
            clsCommon.AddColumnsForChange(coll, "TransferOutNo", obj.TransferOutNo)
            clsCommon.AddColumnsForChange(coll, "Type", obj.Type)
            clsCommon.AddColumnsForChange(coll, "AutoTransfer", obj.AutoTransfer)
            clsCommon.AddColumnsForChange(coll, "Form38", IIf(obj.Form38, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Secondary_Code", obj.Secondary_Code)

            If clsCommon.myLen(obj.Due_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Due_Date", Nothing, True)
            End If

            clsCommon.AddColumnsForChange(coll, "For_Repair", IIf(obj.For_Repair, 1, 0))
            clsCommon.AddColumnsForChange(coll, "InternalTRansfer", IIf(obj.InternalTransfer, 1, 0))
            clsCommon.AddColumnsForChange(coll, "JobWorkTransfer", IIf(obj.JobWorkTransfer, 1, 0))
            clsCommon.AddColumnsForChange(coll, "ProdRequestTransfer", IIf(obj.ProdRequestTransfer, 1, 0))
            clsCommon.AddColumnsForChange(coll, "IsJobWorkType", IIf(obj.IsJobWorkType, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Requisition_Id", clsCommon.myCstr(obj.Requisition_Id), True)
            clsCommon.AddColumnsForChange(coll, "WayBill_No", obj.Waybill_No)
            ''richa 03/03/2015
            If clsCommon.myLen(obj.Waybill_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "WayBill_Date", clsCommon.GetPrintDate(obj.Waybill_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "WayBill_Date", Nothing, True)
            End If
            ''----------------
            clsCommon.AddColumnsForChange(coll, "GR_No", obj.GR_No)
            clsCommon.AddColumnsForChange(coll, "GR_Date", clsCommon.GetPrintDate(obj.GR_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Transport_Id", obj.Transport_Id)
            clsCommon.AddColumnsForChange(coll, "Transporter_Name_Manual", obj.Transporter_Name_Manual)
            clsCommon.AddColumnsForChange(coll, "Freight_Distance", obj.Freight_Distance)
            clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
            clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
            clsCommon.AddColumnsForChange(coll, "Terms_Remark", obj.Terms_Remark)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))

            '' currencyconversion
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", obj.ApplicableFrom, True)
            '' End currencyconversion
            clsCommon.AddColumnsForChange(coll, "RMDA_Code", obj.RMDA_Code)

            clsCommon.AddColumnsForChange(coll, "Vehicle_Mannual_No", obj.Vehicle_Mannual_No)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Capacity", obj.Vehicle_Capacity)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Charge", obj.Vehicle_Charge)
            clsCommon.AddColumnsForChange(coll, "Total_Item_Wt", obj.Total_Item_Wt)
            clsCommon.AddColumnsForChange(coll, "Gross_Item_Wt", obj.Gross_Item_Wt)

            clsCommon.AddColumnsForChange(coll, "Freight_Type", obj.Freight_Type)
            clsCommon.AddColumnsForChange(coll, "FixedCharge", obj.FixedCharge)
            clsCommon.AddColumnsForChange(coll, "EmptyCharge", obj.EmptyCharge)

            ''richa agarwal 
            If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I' ", trans)), obj.TransferOutNo) = CompairStringResult.Equal) Or (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I' ", trans)), "") = CompairStringResult.Equal) Then
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='Y' where Document_No='" & obj.TransferOutNo & "' and Transfer_Type='O'", trans)
                Else
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='N' where Document_No=(Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I') and Transfer_Type='O'", trans)
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='Y' where Document_No='" & obj.TransferOutNo & "' and Transfer_Type='O'", trans)
                End If
            End If
            ''-----------------

            ''preeti gupta
            clsCommon.AddColumnsForChange(coll, "Crate_IN", obj.Crate_IN)
            clsCommon.AddColumnsForChange(coll, "jaali_IN", obj.jaali_IN)
            clsCommon.AddColumnsForChange(coll, "Box_IN", obj.Box_IN)
            clsCommon.AddColumnsForChange(coll, "Crate_Out", obj.Crate_Out)
            clsCommon.AddColumnsForChange(coll, "jaali_Out", obj.jaali_Out)
            clsCommon.AddColumnsForChange(coll, "Box_Out", obj.Box_Out)
            ''=================
            clsCommon.AddColumnsForChange(coll, "Is_Taxable", obj.Is_Taxable)
            clsCommon.AddColumnsForChange(coll, "Is_MandiTax", obj.Is_MandiTax)
            clsCommon.AddColumnsForChange(coll, "GLVoucher_No", obj.GLVoucher_No)
            clsCommon.AddColumnsForChange(coll, "Loading_Advice_No", obj.Loading_Advice_No)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Document_No", obj.Document_No)
                clsCommon.AddColumnsForChange(coll, "Document_Type", obj.Document_Type)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSFER_ORDER_HEAD", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSFER_ORDER_HEAD", OMInsertOrUpdate.Update, "TSPL_TRANSFER_ORDER_HEAD.Document_No='" + obj.Document_No + "'", trans)
            End If

            isSaved = isSaved AndAlso clsTransferDCCDetail.SaveData(obj.Document_No, Arr, trans, obj.Document_Date, obj.From_Location, obj.To_Location, trans_Type, obj.Transfer_Type, obj.IsJobWorkType, obj.InternalTransfer, obj.JobWorkTransfer)
            isSaved = isSaved AndAlso clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(obj.Document_No), "TSPL_TRANSFER_ORDER_HEAD", "DOCUMENT_NO", "TSPL_TRANSFER_ORDER_DETAIL", "DOCUMENT_NO", trans)

            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_No, obj.arrCustomFields, trans)
            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal And clsCommon.myLen(obj.GLVoucher_No) > 0 Then
                Dim strLoadInNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_TRANSFER_ORDER_HEAD where TransferOutNo='" & obj.Document_No & "'", trans)
                isSaved = isSaved AndAlso UpdateLoadinwithLoadout(obj, strLoadInNo, trans)
            End If
            If isNewEntry AndAlso clsCommon.myLen(obj.Requisition_Id) > 0 Then
                Dim isProductionStoreReqDoc As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue(" select count (*)  from TSPL_PP_REQUISITION_HEAD where Requisition_Id = '" + obj.Requisition_Id + "' ", trans))
                If isProductionStoreReqDoc = True Then
                    isSaved = isSaved AndAlso clsProductionRequistionHead.CloseProdStoreRequest(obj.Requisition_Id, trans)
                Else
                    isSaved = isSaved AndAlso clsRequistionHead.CloseprData(obj.Requisition_Id, "Y", trans)
                End If
            End If

        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function
    Public Shared Function eWayBill_Implementation(ByVal strDocNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If
            Dim strQry As String = "select 
TSPL_TRANSFER_ORDER_HEAD.Transfer_Type as supplyType,'1' as subSupplyType,TSPL_TRANSFER_ORDER_HEAD.Remarks as subSupplyDesc,'INV' as docType,  TSPL_TRANSFER_ORDER_HEAD.Document_No as docNo,  convert(date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as docDate,  FromLocation.GSTNO as fromGstin,TSPL_COMPANY_MASTER.Comp_Name as fromTrdName,FromLocation.Add1 as fromAddr1,FromLocation.Add2 as fromAddr2,FromLocation.city_code as fromPlace,Seller_State_Master.GST_STATE_Code as actFromStateCode,FromLocation.Pin_Code as fromPincode,Seller_State_Master.GST_STATE_Code as fromStateCode,ToLocation.GSTNo as toGstin,ToLocation.location_desc as toTrdName,  Tolocation.Add1 as toAddr1,Tolocation.Add2 as toAddr2,Tolocation.City_Code as toPlace,Tolocation.Pin_Code as toPincode,  Buyer_State_Master.GST_STATE_Code as actToStateCode,Buyer_State_Master.GST_STATE_Code as toStateCode,'1' AS transactionType,  FromLocation.GSTNO as dispatchFromGSTIN,TSPL_COMPANY_MASTER.Comp_Name as dispatchFromTradeName,ToLocation.GSTNo AS shipToGSTIN,  ToLocation.location_desc AS shipToTradeName,TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt AS totalValue,
'0' as cgstValue,'0' as sgstValue,'0' as igstValue,'0' as cessValue,'0' as cessNonAdvolValue,TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt as totInvValue,'1' as transMode,
TSPL_TRANSFER_ORDER_HEAD.Km_Reading as transDistance,
'' AS transporterName,tspl_vendor_master.GSTFinalNo AS transporterId,
  '' as transDocNo,'' as transDocDate,
  TSPL_VEHICLE_MASTER.Number as vehicleNo,
  'R' as vehicleType,
  
  TSPL_ITEM_MASTER.Item_Desc AS productName, 
  TSPL_ITEM_MASTER.Item_Desc AS productDesc,TSPL_ITEM_MASTER.HSN_Code AS hsnCode,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as quantity, TSPL_TRANSFER_ORDER_DETAIL.Unit_code as qtyUnit,TSPL_TRANSFER_ORDER_DETAIL.Item_cost as ItemUnitPrice,  TSPL_TRANSFER_ORDER_DETAIL.Amount as ItemTotAmt,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt as ItemDiscount,TSPL_TRANSFER_ORDER_DETAIL.Amount as taxableAmount,'0' as sgstRate,'0' as cgstRate,'0' as igstRate,'0' as cessRate,0 as NO_Transporter
from 
  TSPL_TRANSFER_ORDER_HEAD 
  Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' 
  left Outer Join TSPL_LOCATION_MASTER as ToLocation on ToLocation.Location_Code = TSPL_TRANSFER_ORDER_HEAD.To_Location 
  left Outer Join TSPL_LOCATION_MASTER as FromLocation on FromLocation.Location_Code = TSPL_TRANSFER_ORDER_HEAD.From_Location 
  left outer join TSPL_TRANSFER_ORDER_DETAIL on TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No 
  left outer join tspl_item_master on tspl_item_master.Item_code = TSPL_TRANSFER_ORDER_DETAIL.Item_code 
  left outer join TSPL_STATE_MASTER as Seller_State_Master on Seller_State_Master.STATE_CODE = FromLocation.State 
  left outer join TSPL_STATE_MASTER as Buyer_State_Master on Buyer_State_Master.STATE_CODE = ToLocation.State 
  left outer join tspl_city_master on tspl_city_master.city_code = ToLocation.City_Code 
  left outer join tspl_city_master as BuyerCity on BuyerCity.city_code = FromLocation.City_Code 
  left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code = TSPL_TRANSFER_ORDER_HEAD.Tax2 
  left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code = TSPL_TRANSFER_ORDER_HEAD.Tax3 
  Left Outer Join tspl_vendor_master on tspl_vendor_master.vendor_code = TSPL_TRANSFER_ORDER_HEAD.Transport_Id 
  Left Outer Join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id = TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code 
where 
  TSPL_TRANSFER_ORDER_HEAD.Document_No = '" & strDocNo & "' AND TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O' AND TSPL_TRANSFER_ORDER_HEAD.IsJobWorkType = 0 "
            Dim objResult As Object = ClsEInvoiceOFAPIs.PostEWayBill(objCommonVar.CurrentCompanyCode, strQry, strLocation, trans)
            If objResult IsNot Nothing Then
                Dim EWayBillNo As String = objResult.SelectToken("data.ewayBillNo").ToString
                Dim EWayBillDate As String = objResult.SelectToken("data.ewayBillDate").ToString
                Dim EWayBillValidDate As String = objResult.SelectToken("data.validUpto").ToString
                Dim EWayBillRemarks As String = objResult.SelectToken("data.alert").ToString
                Dim CompGSTNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GSTReg_No from TSPL_COMPANY_MASTER ", trans))
                Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(EWayBillNo + "/" + CompGSTNo + "/" + clsCommon.GetPrintDate(EWayBillValidDate, "dd/MMM/yyyy hh:mm tt"))
                clsDBFuncationality.ExecuteNonQuery("update TSPL_TRANSFER_ORDER_HEAD set  EWayBillNo ='" & EWayBillNo & "',EWayBillDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(EWayBillDate), "dd/MMM/yyyy hh:mm tt") & "',EWayBillValidDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(EWayBillValidDate), "dd/MMM/yyyy hh:mm tt") & "',EWayBillRemarks='" & EWayBillRemarks & "' where TSPL_TRANSFER_ORDER_HEAD.Document_No ='" & strDocNo & "'", trans)
                clsDBFuncationality.UpdateImage("EWayBill_QR_Code", TempByte, "TSPL_TRANSFER_ORDER_HEAD", "TSPL_TRANSFER_ORDER_HEAD.Document_No='" & strDocNo & "'", trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try
        Return True
    End Function


    Public Shared Function EInvoice_Implementation(ByVal strDocNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If

            Dim strtoken As String = ClsEInvoiceOFAPIs.IsGenerateAuthTokenNo_Required(objCommonVar.CurrentCompanyCode, strLocation, trans)
            If clsCommon.myLen(strtoken) > 0 Then
                Dim strQry As String = "select TSPL_TRANSFER_ORDER_HEAD.Document_No as DocNo,convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as DocDate,'INV' as DocType ,'B2B' as SupTyp, 'N'  as IgstOnIntra,FromLocation.GSTNO as SellerGSTINNo ,FromLocation.location_desc as SellerLglNm,TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm,FromLocation.Add1 as SellerAdd1,FromLocation.Add2 as SellerAdd2 ,FromLocation.city_code  as SellerLoc,FromLocation.Pin_Code  as SellerPincode,Seller_State_Master .GST_STATE_Code as SellerStcd,FromLocation.Phone1 as SellerPhone,FromLocation.Email as SellerEmail,ToLocation.GSTNo as BuyerGSTINNo ,ToLocation.location_desc as BuyerLglNm,ToLocation.location_desc as BuyerTrdNm, Buyer_State_Master.GST_STATE_Code as BuyerPOS,Tolocation.Add1 as BuyerAdd1,Tolocation.Add2 as BuyerAdd2 ,Tolocation.City_Code as BuyerLoc,Tolocation.Pin_Code as BuyerPincode,Buyer_State_Master.GST_STATE_Code as BuyerStcd,Tolocation.Phone1 as BuyerPhone,Tolocation.Email as BuyerEmail,TSPL_TRANSFER_ORDER_DETAIL.Line_No as ItemSlNo, 'N' as ItemIsServc,TSPL_ITEM_MASTER.Item_Desc AS ItemPrdDesc,TSPL_ITEM_MASTER.HSN_Code AS ItemHsnCd,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as ItemQty, TSPL_TRANSFER_ORDER_DETAIL.Unit_code as ItemUnit,TSPL_TRANSFER_ORDER_DETAIL.Item_cost as ItemUnitPrice,TSPL_TRANSFER_ORDER_DETAIL.Amount as ItemTotAmt,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt as ItemDiscount,TSPL_TRANSFER_ORDER_DETAIL.Amount-TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt as ItemAssAmt,case when ISNULL(TSPL_TRANSFER_ORDER_DETAIL .tax1,'') ='IGST' THEN TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate when ISNULL(TSPL_TRANSFER_ORDER_DETAIL .tax1,'') ='CGST' AND   ISNULL(TSPL_TRANSFER_ORDER_DETAIL .tax2,'') ='SGST'  THEN TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate+TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate  ELSE 0 end as ItemGstRt, case when TSPL_TRANSFER_ORDER_DETAIL .TAX1 ='SGST' AND TSPL_TRANSFER_ORDER_DETAIL .TAX2  ='CGST' then TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt when TSPL_TRANSFER_ORDER_DETAIL .TAX1 ='CGST' AND TSPL_TRANSFER_ORDER_DETAIL .TAX2  ='SGST' then TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt else 0 end ItemSgstAmt,case when TSPL_TRANSFER_ORDER_DETAIL .TAX1 ='IGST' then TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt else 0 end ItemIgstAmt,
case when TSPL_TRANSFER_ORDER_DETAIL .TAX1 ='CGST' AND TSPL_TRANSFER_ORDER_DETAIL .TAX2  ='SGST' then TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt
when TSPL_TRANSFER_ORDER_DETAIL .TAX1 ='SGST' AND TSPL_TRANSFER_ORDER_DETAIL .TAX2  ='CGST' then TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt 
else 0 end ItemCgstAmt,0 as ItemOthChrg,TSPL_TRANSFER_ORDER_DETAIL.item_net_amt-case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_TRANSFER_ORDER_DETAIL.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y'  THEN  TSPL_TRANSFER_ORDER_DETAIL.TAX3_AMT ELSE 0 END as ItemTotItemVal,TSPL_TRANSFER_ORDER_HEAD .discount_base as ValDtlsAssVal,
case when TSPL_TRANSFER_ORDER_HEAD .TAX1 ='CGST' AND TSPL_TRANSFER_ORDER_HEAD .TAX2  ='SGST' then TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt
when TSPL_TRANSFER_ORDER_HEAD .TAX1 ='SGST' AND TSPL_TRANSFER_ORDER_HEAD .TAX2  ='CGST' then TSPL_TRANSFER_ORDER_HEAD.TAX2_Amt 
else 0 end ValDtlsCgstVal, case when TSPL_TRANSFER_ORDER_HEAD .TAX1 ='SGST' AND TSPL_TRANSFER_ORDER_HEAD .TAX2  ='CGST' then TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt when TSPL_TRANSFER_ORDER_HEAD .TAX1 ='CGST' AND TSPL_TRANSFER_ORDER_HEAD .TAX2  ='SGST' then TSPL_TRANSFER_ORDER_HEAD.TAX2_Amt else 0 end ValDtlsSgstVal,case when TSPL_TRANSFER_ORDER_HEAD .TAX1 ='IGST' then TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt else 0 end ValDtlsIgstVal,TSPL_TRANSFER_ORDER_HEAD.Discount_Amt as ValDtlsDiscount,case when isnull(TCS1.is_tcs,'')='Y' THEN  TSPL_TRANSFER_ORDER_HEAD.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y'  THEN  TSPL_TRANSFER_ORDER_HEAD.TAX3_AMT ELSE 0 END as ValDtlsOthChrg,TSPL_TRANSFER_ORDER_HEAD.Doc_Total_amt  as ValDtlsTotInvVal,TSPL_TRANSFER_ORDER_HEAD.RoundOffAmount  as ValDtlsRndOffAmt
,ISNULL(tspl_vendor_master.GSTFinalNo,'') AS EwbTransId,ISNULL(tspl_vendor_master.Vendor_Name,'') AS EwbTransName,TSPL_TRANSFER_ORDER_HEAD.Freight_Distance as EwbDistance,isnull(TSPL_VEHICLE_MASTER.Number,'') as EwbVehNo
from TSPL_TRANSFER_ORDER_HEAD
Left Outer Join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code  ='" & objCommonVar.CurrentCompanyCode & "'
left Outer Join TSPL_LOCATION_MASTER as ToLocation on ToLocation.GIT_Location  =TSPL_TRANSFER_ORDER_HEAD.To_Location  
left Outer Join TSPL_LOCATION_MASTER as FromLocation on FromLocation.Location_Code =TSPL_TRANSFER_ORDER_HEAD.From_Location    
left outer join TSPL_TRANSFER_ORDER_DETAIL on TSPL_TRANSFER_ORDER_DETAIL.Document_No =TSPL_TRANSFER_ORDER_HEAD.Document_No
left outer join tspl_item_master on tspl_item_master.Item_code=TSPL_TRANSFER_ORDER_DETAIL.Item_code
left outer join TSPL_STATE_MASTER as Seller_State_Master on Seller_State_Master.STATE_CODE  =FromLocation.State
left outer join TSPL_STATE_MASTER as Buyer_State_Master on Buyer_State_Master.STATE_CODE  =ToLocation.State
left outer join tspl_city_master on tspl_city_master.city_code=ToLocation.City_Code
left outer join tspl_city_master as BuyerCity on BuyerCity.city_code=FromLocation.City_Code
left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code =TSPL_TRANSFER_ORDER_HEAD.Tax2
left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code =TSPL_TRANSFER_ORDER_HEAD.Tax3
Left Outer Join tspl_vendor_master on tspl_vendor_master.vendor_code  =TSPL_TRANSFER_ORDER_HEAD.Transport_Id
Left Outer Join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code
where TSPL_TRANSFER_ORDER_HEAD.Document_No  ='" & strDocNo & "' AND TSPL_TRANSFER_ORDER_HEAD.transfer_type  ='O' AND TSPL_TRANSFER_ORDER_HEAD.IsJobWorkType =0"

                Dim objResult As Object = ClsEInvoiceOFAPIs.PostAuthTokenNo_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strQry, strLocation, trans)
                If objResult IsNot Nothing Then
                    'assign to variable
                    Dim AckNo As String = objResult.SelectToken("AckNo").ToString
                    Dim AckDt As String = objResult.SelectToken("AckDt").ToString
                    Dim Irn As String = objResult.SelectToken("Irn").ToString
                    Dim SignedQRCode As String = objResult.SelectToken("SignedQRCode").ToString
                    clsDBFuncationality.ExecuteNonQuery("update TSPL_TRANSFER_ORDER_HEAD set  IRN_No ='" & Irn & "',qr_code='" & SignedQRCode & "',ack_no='" & AckNo & "',ack_date='" & clsCommon.GetPrintDate(AckDt, "dd/MMM/yyyy hh:mm tt") & "' where TSPL_TRANSFER_ORDER_HEAD.Document_No ='" & strDocNo & "'", trans)

                    Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(SignedQRCode)
                    clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_TRANSFER_ORDER_HEAD", "TSPL_TRANSFER_ORDER_HEAD.Document_No='" & strDocNo & "'", trans)

                    If objCommonVar.GenerateEWayBillWithEInvoice = True Then
                        Dim EwbNo As String = objResult.SelectToken("EwbNo").ToString
                        Dim EwbDt As String = objResult.SelectToken("EwbDt").ToString
                        Dim EwbValidTill As String = objResult.SelectToken("EwbValidTill").ToString
                        Dim Remarks As String = objResult.SelectToken("Remarks").ToString
                        If clsCommon.myLen(EwbDt) > 0 Then
                            EwbDt = clsCommon.GetPrintDate(EwbDt, "dd/MMM/yyyy hh:mm tt")
                        End If
                        If clsCommon.myLen(EwbValidTill) > 0 Then
                            EwbValidTill = clsCommon.GetPrintDate(EwbValidTill, "dd/MMM/yyyy hh:mm tt")
                        End If
                        clsDBFuncationality.ExecuteNonQuery("update TSPL_TRANSFER_ORDER_HEAD set  EWayBillNo ='" & EwbNo & "',EwayBillDate=(CASE WHEN LEN('" & EwbDt & "')>0   THEN '" & EwbDt & "' ELSE NULL END) ,EwayBillValidDate=(CASE WHEN LEN('" & EwbValidTill & "')>0   THEN '" & EwbValidTill & "' ELSE NULL END)  , EWayBillRemarks = '" & Remarks & "'  where Document_No ='" & strDocNo & "' ", trans)
                    End If

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

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType) As clsTransferDCC
        Return GetData(strDocumentNo, NavType, Nothing)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As clsTransferDCC
        Dim obj As clsTransferDCC = Nothing
        Dim qry As String = "SELECT TSPL_Transfer_ORDER_Head.TransferIndent_No,TSPL_Transfer_ORDER_Head.Loading_Advice_No,TSPL_Transfer_ORDER_Head.GLVoucher_No, TSPL_Transfer_ORDER_Head.Electronic_Ref_No, TSPL_Transfer_ORDER_Head.Is_MandiTax,TSPL_Transfer_ORDER_Head.EWayBillNo,TSPL_Transfer_ORDER_Head.EWayBillDate,TSPL_Transfer_ORDER_Head.Is_Taxable, TSPL_Transfer_ORDER_Head.Freight_Type,TSPL_Transfer_ORDER_Head.FixedCharge,TSPL_Transfer_ORDER_Head.EmptyCharge,TSPL_Transfer_ORDER_Head.Crate_IN ,TSPL_Transfer_ORDER_Head.Box_IN ,TSPL_Transfer_ORDER_Head.jaali_IN ,TSPL_Transfer_ORDER_Head.Crate_Out ,TSPL_Transfer_ORDER_Head.jaali_Out ,TSPL_Transfer_ORDER_Head.Box_Out, TSPL_TRANSFER_ORDER_HEAD.GP_Item_Type,TSPL_TRANSFER_ORDER_HEAD.Price_Code,TSPL_TRANSFER_ORDER_HEAD.Is_Status_IN,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,TSPL_TRANSFER_ORDER_HEAD.Gross_Item_Wt,TSPL_TRANSFER_ORDER_HEAD.Total_Item_Wt,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Capacity,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Charge,TSPL_TRANSFER_ORDER_HEAD.Form38,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code,TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_HEAD.Km_Reading, " &
        " TSPL_TRANSFER_ORDER_HEAD.TransferOutNo,TSPL_TRANSFER_ORDER_HEAD.Type,TSPL_TRANSFER_ORDER_HEAD.Document_No,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ," &
        " TSPL_TRANSFER_ORDER_HEAD.Document_Date,TSPL_TRANSFER_ORDER_HEAD.Secondary_Code, " &
        " TSPL_TRANSFER_ORDER_HEAD.Status,TSPL_TRANSFER_ORDER_HEAD.On_Hold,TSPL_TRANSFER_ORDER_HEAD.Ref_No,TSPL_TRANSFER_ORDER_HEAD.Description, " &
        " TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_TRANSFER_ORDER_HEAD.Tax_Group,TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_TRANSFER_ORDER_HEAD.GR_No, TSPL_TRANSFER_ORDER_HEAD.GR_Date, TSPL_TRANSFER_ORDER_HEAD.WayBill_No, TSPL_TRANSFER_ORDER_HEAD.WayBill_Date, ISNULL(TSPL_TRANSFER_ORDER_HEAD.Transport_Id,'') AS Transport_Id," &
        " TSPL_TRANSFER_ORDER_HEAD.To_Location, TSPL_LOCATION_MASTER_2.Location_Code as To_Location_MAIN, TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX1_Base_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate," &
        " TSPL_TRANSFER_ORDER_HEAD.TAX2_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX2_Base_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX3," &
        " TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX3_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX3_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX4_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX4_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX5_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX5_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX6,TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX6_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX6_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX7,TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX7_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX7_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX8,TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX8_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX8_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX9,TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX9_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX9_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.TAX10,TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX10_Amt,TSPL_TRANSFER_ORDER_HEAD.TAX10_Base_Amt, " &
        " TSPL_TRANSFER_ORDER_HEAD.Discount_Base,TSPL_TRANSFER_ORDER_HEAD.Discount_Amt,TSPL_TRANSFER_ORDER_HEAD.Amount_Less_Discount,TSPL_TRANSFER_ORDER_HEAD.Total_Amt_Less_Tax, " &
        " TSPL_TRANSFER_ORDER_HEAD.Total_Tax_Amt,TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_HEAD.Mode_Of_Transport, " &
        " TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Comments,TSPL_TRANSFER_ORDER_HEAD.Comp_Code, " &
        " TSPL_TRANSFER_ORDER_HEAD.Terms_Code,TSPL_TRANSFER_ORDER_HEAD.Terms_Remark,TSPL_TRANSFER_ORDER_HEAD.Due_Date , " &
        " TSPL_LOCATION_MASTER.Location_Desc as From_LocationName,TSPL_LOCATION_MASTER_1.Location_Desc as To_LocationName, " &
        " TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as TaxGroupName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_TRANSFER_ORDER_HEAD.Posting_Date, " &
        " TSPL_TRANSFER_ORDER_HEAD.Delivery_date, TSPL_TRANSFER_ORDER_HEAD.Delivery_Duration,TSPL_TRANSFER_ORDER_HEAD.Item_Type, " &
        " TSPL_TRANSFER_ORDER_HEAD.Modify_By,TSPL_TRANSFER_ORDER_HEAD.Modify_Date,TSPL_TRANSFER_ORDER_HEAD.Created_By, " &
        " TSPL_TRANSFER_ORDER_HEAD.Created_Date,TSPL_TRANSFER_ORDER_HEAD.Tax_Calculation_Type,TSPL_TRANSFER_ORDER_HEAD.IsEwayBill, " &
        " TSPL_TRANSFER_ORDER_HEAD.CURRENCY_CODE,TSPL_TRANSFER_ORDER_HEAD.CONVRATE,TSPL_TRANSFER_ORDER_HEAD.ApplicableFrom,TSPL_TRANSFER_ORDER_HEAD.RMDA_Code,ISNULL(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual,'') AS Transporter_Name_Manual,TSPL_Transfer_ORDER_Head.For_Repair,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer,TSPL_TRANSFER_ORDER_HEAD.JobWorkTransfer,TSPL_TRANSFER_ORDER_HEAD.ProdRequestTransfer,TSPL_TRANSFER_ORDER_HEAD.IsJobWorkType,Requisition_Id,TSPL_TRANSFER_ORDER_HEAD.Freight_Distance FROM TSPL_TRANSFER_ORDER_HEAD left " &
        " outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_HEAD.From_Location " &
        " left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location " &
        " left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.GIT_Location=TSPL_TRANSFER_ORDER_HEAD.To_Location" &
        " left outer join  TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code= TSPL_TRANSFER_ORDER_HEAD.Tax_Group " &
        " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_TRANSFER_ORDER_HEAD.Terms_Code where 2=2"

        Dim whrClas As String = ""

        '=get locations======Done By Monika for location rights 19/01/2015===========
        Dim arrLoc As String = ""
        Dim LOCobj As New clsMCCCodes()
        LOCobj = clsMCCCodes.GetData(trans)
        If LOCobj.arrLocCodes IsNot Nothing AndAlso clsCommon.myLen(LOCobj.arrLocCodes) > 0 Then
            arrLoc = LOCobj.arrLocCodes
        End If
        '==================================================
        If clsCommon.myLen(arrLoc) > 0 Then
            whrClas = " AND (case when transfer_type='O' then from_location else to_location end) in (" + arrLoc + ")"
        End If

        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = (select MIN(Document_No) from TSPL_TRANSFER_ORDER_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Last
                qry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = (select Max(Document_No) from TSPL_TRANSFER_ORDER_HEAD where 1=1 " + whrClas + ")"
            Case NavigatorType.Next
                qry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = (select Min(Document_No) from TSPL_TRANSFER_ORDER_HEAD where Document_No>'" + strPONo + "' " + whrClas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = (select Max(Document_No) from TSPL_TRANSFER_ORDER_HEAD where Document_No<'" + strPONo + "' " + whrClas + ")"
            Case NavigatorType.Current
                qry += " and TSPL_TRANSFER_ORDER_HEAD.Document_No = '" + strPONo + "'"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsTransferDCC()
            obj.Form38 = IIf(clsCommon.myCdbl(dt.Rows(0)("Form38")) = 1, True, False)
            obj.Document_No = clsCommon.myCstr(dt.Rows(0)("Document_No"))
            obj.Document_Date = clsCommon.myCstr(dt.Rows(0)("Document_Date"))
            obj.Delivery_date = clsCommon.myCstr(dt.Rows(0)("Delivery_date"))
            obj.Delivery_Duration = clsCommon.myCstr(dt.Rows(0)("Delivery_Duration"))
            obj.Transfer_Type = clsCommon.myCstr(dt.Rows(0)("Transfer_Type"))

            obj.Secondary_Code = clsCommon.myCstr(dt.Rows(0)("Secondary_Code"))

            obj.Freight_Type = clsCommon.myCstr(dt.Rows(0)("Freight_Type"))
            obj.FixedCharge = clsCommon.myCdbl(dt.Rows(0)("FixedCharge"))
            obj.EmptyCharge = clsCommon.myCdbl(dt.Rows(0)("EmptyCharge"))

            obj.Price_Code = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
            obj.GP_Item_Type = clsCommon.myCstr(dt.Rows(0)("GP_Item_Type"))
            obj.TransferIndent_No = clsCommon.myCstr(dt.Rows(0)("TransferIndent_No"))

            'obj.Status = IIf(clsCommon.myCdbl(dt.Rows(0)("Status")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)

            If clsCommon.myCdbl(dt.Rows(0)("Status")) = 1 Then
                obj.Status = ERPTransactionStatus.Approved
            ElseIf clsCommon.myCdbl(dt.Rows(0)("Status")) = 2 Then
                obj.Status = ERPTransactionStatus.Cancel
            Else
                obj.Status = ERPTransactionStatus.Pending
            End If
            obj.On_Hold = IIf(clsCommon.myCdbl(dt.Rows(0)("On_Hold")) = 1, True, False)
            obj.Ref_No = clsCommon.myCstr(dt.Rows(0)("Ref_No"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Is_Status_IN = clsCommon.myCstr(dt.Rows(0)("Is_Status_IN"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.IsEwayBill = clsCommon.myCdbl(dt.Rows(0)("IsEwayBill"))
            obj.From_Location = clsCommon.myCstr(dt.Rows(0)("From_Location"))
            obj.To_Location = clsCommon.myCstr(dt.Rows(0)("To_Location"))
            obj.To_Location_Main = clsCommon.myCstr(dt.Rows(0)("To_Location_Main"))
            obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
            obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
            obj.TAX1_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
            obj.TAX1_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Base_Amt"))
            obj.TAX1_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            obj.TAX2 = clsCommon.myCstr(dt.Rows(0)("TAX2"))
            obj.TAX2_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
            obj.TAX2_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Base_Amt"))
            obj.TAX2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            obj.TAX3 = clsCommon.myCstr(dt.Rows(0)("TAX3"))
            obj.TAX3_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Base_Amt"))
            obj.TAX3_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
            obj.TAX3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            obj.TAX4 = clsCommon.myCstr(dt.Rows(0)("TAX4"))
            obj.TAX4_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
            obj.TAX4_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Base_Amt"))
            obj.TAX4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            obj.TAX5 = clsCommon.myCstr(dt.Rows(0)("TAX5"))
            obj.TAX5_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
            obj.TAX5_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Base_Amt"))
            obj.TAX5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
            obj.TAX6 = clsCommon.myCstr(dt.Rows(0)("TAX6"))
            obj.TAX6_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
            obj.TAX6_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Base_Amt"))
            obj.TAX6_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
            obj.TAX7 = clsCommon.myCstr(dt.Rows(0)("TAX7"))
            obj.TAX7_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX7_Rate"))
            obj.TAX7_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Base_Amt"))
            obj.TAX7_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
            obj.TAX8 = clsCommon.myCstr(dt.Rows(0)("TAX8"))
            obj.TAX8_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX8_Rate"))
            obj.TAX8_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Base_Amt"))
            obj.TAX8_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
            obj.TAX9 = clsCommon.myCstr(dt.Rows(0)("TAX9"))
            obj.TAX9_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX9_Rate"))
            obj.TAX9_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Base_Amt"))
            obj.TAX9_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))
            obj.TAX10 = clsCommon.myCstr(dt.Rows(0)("TAX10"))
            obj.TAX10_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX10_Rate"))
            obj.TAX10_Base_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Base_Amt"))
            obj.TAX10_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
            obj.Total_Tax_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
            obj.Discount_Base = clsCommon.myCdbl(dt.Rows(0)("Discount_Base"))
            obj.Discount_Amt = clsCommon.myCdbl(dt.Rows(0)("Discount_Amt"))
            obj.Amount_Less_Discount = clsCommon.myCdbl(dt.Rows(0)("Amount_Less_Discount"))
            obj.Total_Amt_Less_Tax = clsCommon.myCdbl(dt.Rows(0)("Total_Amt_Less_Tax"))
            obj.DOC_Total_Amt = clsCommon.myCdbl(dt.Rows(0)("DOC_Total_Amt"))
            obj.Mode_Of_Transport = clsCommon.myCstr(dt.Rows(0)("Mode_Of_Transport"))
            obj.Comments = clsCommon.myCstr(dt.Rows(0)("Comments"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Terms_Remark = clsCommon.myCstr(dt.Rows(0)("Terms_Remark"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            obj.From_LocationName = clsCommon.myCstr(dt.Rows(0)("From_LocationName"))
            obj.To_LocationName = clsCommon.myCstr(dt.Rows(0)("To_LocationName"))
            obj.TaxGroupName = clsCommon.myCstr(dt.Rows(0)("TaxGroupName"))
            obj.TermsName = clsCommon.myCstr(dt.Rows(0)("TermsName"))
            If clsCommon.myLen(dt.Rows(0)("Posting_Date")) > 0 Then
                obj.Posting_Date = clsCommon.myCstr(dt.Rows(0)("Posting_Date"))
            End If
            obj.Is_AgainstFormF = IIf(clsCommon.myCdbl(dt.Rows(0)("Is_AgainstFormF")) = 1, True, False)
            obj.Vehicle_Code = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
            obj.Vehicle_No = clsCommon.myCstr(dt.Rows(0)("Vehicle_No"))
            obj.Km_Reading = clsCommon.myCstr(dt.Rows(0)("Km_Reading"))
            obj.TransferOutNo = clsCommon.myCstr(dt.Rows(0)("TransferOutNo"))
            obj.Type = clsCommon.myCstr(dt.Rows(0)("Type"))

            obj.Waybill_No = clsCommon.myCstr(dt.Rows(0)("WayBill_No"))
            If clsCommon.myLen(dt.Rows(0)("WayBill_No")) <= 0 Then
                obj.Waybill_Date = clsCommon.GETSERVERDATE(trans)
            Else
                obj.Waybill_Date = clsCommon.myCDate(dt.Rows(0)("waybill_date"))
            End If
            obj.GR_No = clsCommon.myCstr(dt.Rows(0)("GR_No"))
            If clsCommon.myLen(dt.Rows(0)("GR_No")) <= 0 Then
                obj.GR_Date = clsCommon.GETSERVERDATE(trans)
            Else
                obj.GR_Date = clsCommon.myCDate(dt.Rows(0)("GR_Date"))
            End If
            obj.Transport_Id = clsCommon.myCstr(dt.Rows(0)("Transport_Id"))
            obj.Transporter_Name_Manual = clsCommon.myCstr(dt.Rows(0)("Transporter_Name_Manual"))
            obj.Freight_Distance = clsCommon.myCdbl(dt.Rows(0)("Freight_Distance"))
            obj.Modify_By = clsCommon.myCstr(dt.Rows(0)("Modify_By"))
            obj.Modify_Date = clsCommon.myCstr(dt.Rows(0)("Modify_Date"))
            obj.Created_By = clsCommon.myCstr(dt.Rows(0)("Created_By"))
            obj.Created_Date = clsCommon.myCstr(dt.Rows(0)("Created_Date"))
            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)

            obj.RMDA_Code = clsCommon.myCstr(dt.Rows(0)("RMDA_Code"))

            '' CURRENCYCONVERSION 
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If
            '' END CURRENCYCONVERSION 

            obj.Vehicle_Capacity = clsCommon.myCdbl(dt.Rows(0)("Vehicle_Capacity"))
            obj.Gross_Item_Wt = clsCommon.myCdbl(dt.Rows(0)("Gross_Item_Wt"))
            obj.Vehicle_Charge = clsCommon.myCdbl(dt.Rows(0)("Vehicle_Charge"))
            obj.Vehicle_Mannual_No = clsCommon.myCstr(dt.Rows(0)("Vehicle_Mannual_No"))
            obj.Total_Item_Wt = clsCommon.myCdbl(dt.Rows(0)("Total_Item_Wt"))
            '=========added By preeti Gupta===============
            obj.Crate_IN = clsCommon.myCdbl(dt.Rows(0)("Crate_IN"))
            obj.jaali_IN = clsCommon.myCdbl(dt.Rows(0)("jaali_IN"))
            obj.Box_IN = clsCommon.myCdbl(dt.Rows(0)("Box_IN"))
            obj.Crate_Out = clsCommon.myCdbl(dt.Rows(0)("Crate_Out"))
            obj.jaali_Out = clsCommon.myCdbl(dt.Rows(0)("jaali_Out"))
            obj.Box_Out = clsCommon.myCdbl(dt.Rows(0)("Box_Out"))
            '=============================================
            If dt.Rows(0)("EWayBillDate") IsNot DBNull.Value Then
                obj.EWayBillDate = clsCommon.myCDate(dt.Rows(0)("EWayBillDate"))
            End If
            obj.EWayBillNo = clsCommon.myCstr(dt.Rows(0)("EWayBillNo"))
            obj.Electronic_Ref_No = clsCommon.myCstr(dt.Rows(0)("Electronic_Ref_No"))
            obj.Is_MandiTax = clsCommon.myCdbl(dt.Rows(0)("Is_MandiTax"))
            obj.Is_Taxable = clsCommon.myCdbl(dt.Rows(0)("Is_Taxable"))
            obj.For_Repair = clsCommon.myCdbl(dt.Rows(0)("For_Repair"))
            obj.GLVoucher_No = clsCommon.myCstr(dt.Rows(0)("GLVoucher_No"))
            obj.Loading_Advice_No = clsCommon.myCstr(dt.Rows(0)("Loading_Advice_No"))
            obj.InternalTransfer = clsCommon.myCdbl(dt.Rows(0)("InternalTransfer"))
            obj.JobWorkTransfer = clsCommon.myCdbl(dt.Rows(0)("JobWorkTransfer"))
            obj.ProdRequestTransfer = clsCommon.myCdbl(dt.Rows(0)("ProdRequestTransfer"))
            obj.IsJobWorkType = clsCommon.myCdbl(dt.Rows(0)("IsJobWorkType"))
            obj.Requisition_Id = clsCommon.myCstr(dt.Rows(0)("Requisition_Id"))
            Dim StrCrateTransferFromBooking As String = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, trans)
            If StrCrateTransferFromBooking = "1" Then
                Dim TransactionTypeQry As String = "select Transfer_Type from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + obj.Document_No + "'"
                Dim TransactionTypeValue = clsDBFuncationality.getSingleValue(TransactionTypeQry, trans)
                If TransactionTypeValue <> "I" Then
                    qry = "SELECT TSPL_TRANSFER_ORDER_DETAIL.ItemwiseTaxCode,TSPL_TRANSFER_ORDER_DETAIL.GatePassNo, TSPL_TRANSFER_ORDER_DETAIL.Bin_No, TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Unit_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_MT_Wt,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Line_No, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Status,TSPL_TRANSFER_ORDER_DETAIL.Row_Type,TSPL_TRANSFER_ORDER_DETAIL.Item_Code, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Item_Desc,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo,TSPL_TRANSFER_ORDER_DETAIL.Unit_code, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Location,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Per, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX1,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX2,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX3,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX4,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX5,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX6,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX7,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX8,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX9,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX10,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Amount,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Amt_Less_Discount,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt, " &
                "TSPL_LOCATION_MASTER.Location_Desc as LocationName,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Base_Amt , " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX3_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Base_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX6_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Base_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.TAX9_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Base_Amt, " &
                "TSPL_TRANSFER_ORDER_DETAIL.Specification,TSPL_TRANSFER_ORDER_DETAIL.Remarks, " &
                "TSPL_TRANSFER_ORDER_DETAIL.In_Qty,TSPL_TRANSFER_ORDER_DETAIL.Leak,TSPL_TRANSFER_ORDER_DETAIL.Breakage,TSPL_TRANSFER_ORDER_DETAIL.Shortage, " &
                "TSPL_TRANSFER_ORDER_DETAIL.MRP FROM TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_DETAIL.Location  left Outer Join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.GatePassNo  where TSPL_TRANSFER_ORDER_DETAIL.Document_No='" + obj.Document_No + "' ORDER BY TSPL_TRANSFER_ORDER_DETAIL.Line_No"
                Else
                    qry = "SELECT TSPL_TRANSFER_ORDER_DETAIL.ItemwiseTaxCode,TSPL_TRANSFER_ORDER_DETAIL.GatePassNo, TSPL_TRANSFER_ORDER_DETAIL.Bin_No, TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Unit_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_MT_Wt,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Line_No, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Status,TSPL_TRANSFER_ORDER_DETAIL.Row_Type,TSPL_TRANSFER_ORDER_DETAIL.Item_Code, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Item_Desc,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo,TSPL_TRANSFER_ORDER_DETAIL.Unit_code, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Location,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Per, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX1,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX2,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX3,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX4,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX5,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX6,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX7,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX8,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX9,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX10,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Amount,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Amt_Less_Discount,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt, " &
                            "TSPL_LOCATION_MASTER.Location_Desc as LocationName,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Base_Amt , " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX3_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX6_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX9_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Specification,TSPL_TRANSFER_ORDER_DETAIL.Remarks, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.In_Qty,TSPL_TRANSFER_ORDER_DETAIL.Leak,TSPL_TRANSFER_ORDER_DETAIL.Breakage,TSPL_TRANSFER_ORDER_DETAIL.Shortage, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.MRP FROM TSPL_TRANSFER_ORDER_Head left Outer join  TSPL_TRANSFER_ORDER_DETAIL on TSPL_TRANSFER_ORDER_Head.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_DETAIL.Location where TSPL_TRANSFER_ORDER_DETAIL.Document_No='" + obj.Document_No + "' ORDER BY TSPL_TRANSFER_ORDER_DETAIL.Line_No"
                End If
            Else
                qry = "SELECT TSPL_TRANSFER_ORDER_DETAIL.ItemwiseTaxCode,TSPL_TRANSFER_ORDER_DETAIL.GatePassNo, TSPL_TRANSFER_ORDER_DETAIL.Bin_No, TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Unit_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Wt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_MT_Wt,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Line_No, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Status,TSPL_TRANSFER_ORDER_DETAIL.Row_Type,TSPL_TRANSFER_ORDER_DETAIL.Item_Code, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Item_Desc,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo,TSPL_TRANSFER_ORDER_DETAIL.Unit_code, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Location,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Per, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX1,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX2,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX3,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX4,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX5,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX6,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX7,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX8,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX9,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX10,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Amount,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per,TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Amt_Less_Discount,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt, " &
                            "TSPL_LOCATION_MASTER.Location_Desc as LocationName,TSPL_TRANSFER_ORDER_DETAIL.TAX1_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX2_Base_Amt , " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX3_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX6_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.TAX9_Base_Amt,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Base_Amt, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.Specification,TSPL_TRANSFER_ORDER_DETAIL.Remarks, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.In_Qty,TSPL_TRANSFER_ORDER_DETAIL.Leak,TSPL_TRANSFER_ORDER_DETAIL.Breakage,TSPL_TRANSFER_ORDER_DETAIL.Shortage, " &
                            "TSPL_TRANSFER_ORDER_DETAIL.MRP FROM TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_DETAIL.Location where TSPL_TRANSFER_ORDER_DETAIL.Document_No='" + obj.Document_No + "' ORDER BY TSPL_TRANSFER_ORDER_DETAIL.Line_No"
            End If

            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsTransferDCCDetail)
                Dim objTr As clsTransferDCCDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsTransferDCCDetail
                    objTr.Document_No = clsCommon.myCstr(dr("Document_No"))
                    objTr.Row_Type = clsCommon.myCstr(dr("Row_Type"))
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(dr("ItemwiseTaxCode"))
                    objTr.Line_No = Convert.ToInt32(clsCommon.myCdbl(dr("Line_No")))
                    objTr.Status = Convert.ToInt32(clsCommon.myCdbl(dr("Status")))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.GatePassNo = clsCommon.myCstr(dr("GatePassNo"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.Out_Qty = clsCommon.myCdbl(dr("Out_Qty"))
                    objTr.TransferOutNo = clsCommon.myCstr(dr("TransferOutNo"))
                    objTr.In_Qty = clsCommon.myCdbl(dr("In_Qty"))
                    objTr.Alt_Unit_Code = clsCommon.myCstr(dr("Alt_Unit_Code"))
                    objTr.Leak = clsCommon.myCdbl(dr("Leak"))
                    objTr.Breakage = clsCommon.myCdbl(dr("Breakage"))
                    objTr.Shortage = clsCommon.myCdbl(dr("Shortage"))
                    objTr.Unit_code = clsCommon.myCstr(dr("Unit_code"))
                    objTr.Location = clsCommon.myCstr(dr("Location"))
                    objTr.LocationName = clsCommon.myCstr(dr("LocationName"))
                    objTr.Item_Cost = clsCommon.myCdbl(dr("Item_Cost"))
                    objTr.TAX1 = clsCommon.myCstr(dr("TAX1"))
                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(dr("TAX1_Base_Amt"))
                    objTr.TAX1_Rate = clsCommon.myCdbl(dr("TAX1_Rate"))
                    objTr.TAX1_Amt = clsCommon.myCdbl(dr("TAX1_Amt"))
                    objTr.TAX2 = clsCommon.myCstr(dr("TAX2"))
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(dr("TAX2_Base_Amt"))
                    objTr.TAX2_Rate = clsCommon.myCdbl(dr("TAX2_Rate"))
                    objTr.TAX2_Amt = clsCommon.myCdbl(dr("TAX2_Amt"))
                    objTr.TAX3 = clsCommon.myCstr(dr("TAX3"))
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(dr("TAX3_Base_Amt"))
                    objTr.TAX3_Rate = clsCommon.myCdbl(dr("TAX3_Rate"))
                    objTr.TAX3_Amt = clsCommon.myCdbl(dr("TAX3_Amt"))
                    objTr.TAX4 = clsCommon.myCstr(dr("TAX4"))
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(dr("TAX4_Base_Amt"))
                    objTr.TAX4_Rate = clsCommon.myCdbl(dr("TAX4_Rate"))
                    objTr.TAX4_Amt = clsCommon.myCdbl(dr("TAX4_Amt"))
                    objTr.TAX5 = clsCommon.myCstr(dr("TAX5"))
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(dr("TAX5_Base_Amt"))
                    objTr.TAX5_Rate = clsCommon.myCdbl(dr("TAX5_Rate"))
                    objTr.TAX5_Amt = clsCommon.myCdbl(dr("TAX5_Amt"))
                    objTr.TAX6 = clsCommon.myCstr(dr("TAX6"))
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(dr("TAX6_Base_Amt"))
                    objTr.TAX6_Rate = clsCommon.myCdbl(dr("TAX6_Rate"))
                    objTr.TAX6_Amt = clsCommon.myCdbl(dr("TAX6_Amt"))
                    objTr.TAX7 = clsCommon.myCstr(dr("TAX7"))
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(dr("TAX7_Base_Amt"))
                    objTr.TAX7_Rate = clsCommon.myCdbl(dr("TAX7_Rate"))
                    objTr.TAX7_Amt = clsCommon.myCdbl(dr("TAX7_Amt"))
                    objTr.TAX8 = clsCommon.myCstr(dr("TAX8"))
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(dr("TAX8_Base_Amt"))
                    objTr.TAX8_Rate = clsCommon.myCdbl(dr("TAX8_Rate"))
                    objTr.TAX8_Amt = clsCommon.myCdbl(dr("TAX8_Amt"))
                    objTr.TAX9 = clsCommon.myCstr(dr("TAX9"))
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(dr("TAX9_Base_Amt"))
                    objTr.TAX9_Rate = clsCommon.myCdbl(dr("TAX9_Rate"))
                    objTr.TAX9_Amt = clsCommon.myCdbl(dr("TAX9_Amt"))
                    objTr.TAX10 = clsCommon.myCstr(dr("TAX10"))
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(dr("TAX10_Base_Amt"))
                    objTr.TAX10_Rate = clsCommon.myCdbl(dr("TAX10_Rate"))
                    objTr.TAX10_Amt = clsCommon.myCdbl(dr("TAX10_Amt"))
                    objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                    objTr.Disc_Per = clsCommon.myCdbl(dr("Disc_Per"))
                    objTr.Disc_Amt = clsCommon.myCdbl(dr("Disc_Amt"))
                    objTr.Abatement_Per = clsCommon.myCdbl(dr("Abatement_Per"))
                    objTr.Abatement_Amt = clsCommon.myCdbl(dr("Abatement_Amt"))
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(dr("Amt_Less_Discount"))
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(dr("Total_Tax_Amt"))
                    objTr.Item_Net_Amt = clsCommon.myCdbl(dr("Item_Net_Amt"))

                    objTr.Specification = clsCommon.myCstr(dr("Specification"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))
                    objTr.MRP = clsCommon.myCdbl(dr("MRP"))
                    'If StrCrateTransferFromBooking = "1" Then
                    '    objTr.FOCItem = clsCommon.myCdbl(dr("FOC_Item"))
                    'End If


                    objTr.Item_Unit_Wt = clsCommon.myCdbl(dr("Item_Unit_Wt"))
                    objTr.Item_Net_Wt = clsCommon.myCdbl(dr("Item_Net_Wt"))
                    objTr.Item_Net_MT_Wt = clsCommon.myCdbl(dr("Item_Net_MT_Wt"))
                    objTr.Bin_No = clsCommon.myCstr(dr("Bin_No"))
                    Dim DocType As String = ""
                    If obj.JobWorkTransfer = 1 Then
                        DocType = "JTransfer"
                    ElseIf obj.InternalTransfer Then
                        DocType = "ITransfer"
                    Else
                        DocType = "Transfer"
                    End If
                    objTr.arrSrItem = clsSerializeInvenotry.GetData(DocType, objTr.Document_No, objTr.Item_Code, objTr.Line_No, trans)
                    objTr.arrBatchItem = clsBatchInventory.GetData(DocType, objTr.Document_No, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr.Add(objTr)
                Next
            End If
        End If

        Return obj
    End Function
    Public Shared Function UpdateLoadinwithLoadout(ByVal obj As clsTransferDCC, ByVal LoadInNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        If (obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0) Then
            qry = clsDBFuncationality.ExecuteNonQuery("  update TSPL_TRANSFER_ORDER_HEAD set " &
            "TAX1_Base_Amt='" & clsCommon.myCdbl(obj.TAX1_Base_Amt) & "', TAX1_Amt='" & clsCommon.myCdbl(obj.TAX1_Amt) & "'," &
            "TAX2_Base_Amt='" & clsCommon.myCdbl(obj.TAX2_Base_Amt) & "', TAX2_Amt='" & clsCommon.myCdbl(obj.TAX2_Amt) & "', " &
            "TAX3_Base_Amt='" & clsCommon.myCdbl(obj.TAX3_Base_Amt) & "', TAX3_Amt='" & clsCommon.myCdbl(obj.TAX3_Base_Amt) & "', " &
            "TAX4_Base_Amt='" & clsCommon.myCdbl(obj.TAX4_Base_Amt) & "', TAX4_Amt='" & clsCommon.myCdbl(obj.TAX4_Base_Amt) & "', " &
            "TAX5_Base_Amt='" & clsCommon.myCdbl(obj.TAX5_Base_Amt) & "', TAX5_Amt='" & clsCommon.myCdbl(obj.TAX5_Base_Amt) & "', " &
            "Discount_Base='" & clsCommon.myCdbl(obj.Discount_Base) & "',Discount_Amt='" & clsCommon.myCdbl(obj.Discount_Amt) & "',Amount_Less_Discount='" & clsCommon.myCdbl(obj.Amount_Less_Discount) & "', " &
            "Total_Tax_Amt='" & clsCommon.myCdbl(obj.Total_Tax_Amt) & "',DOC_Total_Amt='" & clsCommon.myCdbl(obj.DOC_Total_Amt) & "',Total_Amt_Less_Tax='" & clsCommon.myCdbl(obj.Total_Amt_Less_Tax) & "', " &
            "Total_Item_Wt='" & clsCommon.myCdbl(obj.Total_Item_Wt) & "' where Document_No='" & LoadInNo & "'", trans)
            For Each objTr As clsTransferDCCDetail In obj.Arr
                qry = clsDBFuncationality.ExecuteNonQuery(" update TSPL_TRANSFER_ORDER_DETAIL set " &
            "Out_Qty='" & clsCommon.myCdbl(objTr.Out_Qty) & "', In_Qty='" & clsCommon.myCdbl(objTr.Out_Qty) & "', " &
            "TAX1_Base_Amt='" & clsCommon.myCdbl(objTr.TAX1_Base_Amt) & "', TAX1_Amt='" & clsCommon.myCdbl(objTr.TAX1_Amt) & "'," &
            "TAX2_Base_Amt='" & clsCommon.myCdbl(objTr.TAX2_Base_Amt) & "', TAX2_Amt='" & clsCommon.myCdbl(objTr.TAX2_Amt) & "', " &
            "TAX3_Base_Amt='" & clsCommon.myCdbl(objTr.TAX3_Base_Amt) & "', TAX3_Amt='" & clsCommon.myCdbl(objTr.TAX3_Base_Amt) & "', " &
            "TAX4_Base_Amt='" & clsCommon.myCdbl(objTr.TAX4_Base_Amt) & "', TAX4_Amt='" & clsCommon.myCdbl(objTr.TAX4_Base_Amt) & "', " &
            "TAX5_Base_Amt='" & clsCommon.myCdbl(objTr.TAX5_Base_Amt) & "', TAX5_Amt='" & clsCommon.myCdbl(objTr.TAX5_Base_Amt) & "', " &
            "Amount='" & clsCommon.myCdbl(objTr.Amount) & "',Disc_Per='" & clsCommon.myCdbl(objTr.Disc_Per) & "',Disc_Amt='" & clsCommon.myCdbl(objTr.Disc_Amt) & "', " &
            "Amt_Less_Discount='" & clsCommon.myCdbl(objTr.Amt_Less_Discount) & "',Total_Tax_Amt='" & clsCommon.myCdbl(objTr.Total_Tax_Amt) & "', " &
            "Item_Net_Amt='" & clsCommon.myCdbl(objTr.Item_Net_Amt) & "',Item_Net_Wt='" & clsCommon.myCdbl(objTr.Item_Net_Wt) & "', " &
            "Item_Net_MT_Wt='" & clsCommon.myCdbl(objTr.Item_Net_MT_Wt) & "',Item_Cost='" + clsCommon.myCstr(objTr.Item_Cost) + "' " &
            "where Document_No='" & LoadInNo & "' and Item_Code='" & objTr.Item_Code & "' and Unit_code='" & objTr.Unit_code & "' ", trans)
            Next

            Dim objLoadiN As clsTransferDCC = clsTransferDCC.GetData(LoadInNo, NavigatorType.Current, trans)
            clsBatchInventory.ReverseAndUnpost("Transfer", LoadInNo, trans)

            qry = clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_INVENTORY_MOVEMENT where source_doc_no='" & LoadInNo & "'", trans)
            HitInvenotryMovement(trans, objLoadiN, "")
            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='MM-TF' and Source_Doc_No='" + LoadInNo + "'", trans)
            PostTransferJournalEntryOnly(objLoadiN, VoucherNo, trans)
        End If
        Return True
    End Function



    Public Shared Function UpdateAfterPosting(ByVal obj As clsTransferDCC, ByVal TransferNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If obj IsNot Nothing And clsCommon.myLen(obj.Document_No) > 0 Then
                Dim coll As New Hashtable()

                clsCommon.AddColumnsForChange(coll, "EWayBillNo", obj.EWayBillNo)
                clsCommon.AddColumnsForChange(coll, "Electronic_Ref_No", obj.Electronic_Ref_No)
                If clsCommon.myLen(obj.EWayBillDate) > 0 Then
                    clsCommon.AddColumnsForChange(coll, "EWayBillDate", clsCommon.GetPrintDate(obj.EWayBillDate, "dd/MMM/yyyy"))
                Else
                    clsCommon.AddColumnsForChange(coll, "EWayBillDate", Nothing, True)
                End If
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSFER_ORDER_HEAD", OMInsertOrUpdate.Update, "TSPL_TRANSFER_ORDER_HEAD.document_no='" + obj.Document_No + "'", trans)


            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function postTransfer(ByVal strCode As String, Optional ByVal ProvisionAllow As Boolean = False, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            postTransfer(trans, strCode, ProvisionAllow, strVoucherNoForRecreateOnly)
            If objCommonVar.InternalSMSEmailinPurchaseModule = True Then
                Dim obj As clsTransferDCC = clsTransferDCC.GetData(strCode, NavigatorType.Current, trans)
                CreateInternalEmailSMS(obj, trans)
            End If
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function postTransfer(ByVal trans As SqlTransaction, ByVal strCode As String, Optional ByVal ProvisionAllow As Boolean = False, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Return postTransfer(False, trans, strCode, ProvisionAllow, strVoucherNoForRecreateOnly)
    End Function
    Public Shared Function postTransfer(ByVal isUpdateDocAccordingToInvtory As Boolean, ByVal trans As SqlTransaction, ByVal strCode As String, Optional ByVal ProvisionAllow As Boolean = False, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Try
            Dim obj As clsTransferDCC = clsTransferDCC.GetData(strCode, NavigatorType.Current, trans)
            For ii As Integer = 0 To obj.Arr.Count - 1
                If obj.Arr(ii).Item_Cost <= 0 Then
                    Throw New Exception("Unable to Post.Invalid cost of Item [" + obj.Arr(ii).Item_Code + "]")
                End If
            Next

            ''richa agarwal 23 Dec,2020 check eInvoice Implementation
            If clsCommon.CompairString(clsCommon.myCstr(obj.Transfer_Type), "O") = CompairStringResult.Equal AndAlso obj.IsEwayBill = 1 Then

                If clsTransferDCC.eWayBill_Implementation(obj.Document_No, obj.From_Location, trans) = True Then
                Else
                    Throw New Exception("Invalid JSON Value")
                End If
            End If

            If (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TransferJEForLocationMapping, clsFixedParameterCode.TransferJEForLocationMapping, trans)) > 0) Then
                postTransfer(isUpdateDocAccordingToInvtory, strCode, trans, ProvisionAllow, strVoucherNoForRecreateOnly)
            Else
                postTransferNew(isUpdateDocAccordingToInvtory, trans, obj, False, ProvisionAllow, strVoucherNoForRecreateOnly)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_TRANSFER_ORDER_HEAD", "Document_No", trans)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Shared Sub CreateInternalEmailSMS(ByVal obj As clsTransferDCC, ByVal trans As SqlTransaction)

        Dim itemName As String = ""
        Dim UOM As String = ""
        Dim qty As String = ""
        Dim ItemDetail As String = ""
        Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.Transfer + "2" + "'", trans)

        Dim qry As String = "select TSPL_USER_MASTER.User_Code from TSPL_USER_MASTER " &
              " left join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Created_By=TSPL_USER_MASTER.User_Code " &
              " where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + obj.Document_No + "'"
        Dim StrUserCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        Dim arrMobileNo As New List(Of String)
        Dim arrMailID As List(Of String) = clsERPFuncationality.ReportingMailIdandPhone(StrUserCode, arrMobileNo, trans)

        If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 AndAlso ((arrMailID IsNot Nothing AndAlso arrMailID.Count > 0) Or (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0)) Then

            'Dim qry1 As String = "select TSPL_TRANSFER_ORDER_DETAIL.Unit_code,CAST((case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty else TSPL_TRANSFER_ORDER_DETAIL.In_Qty end) AS decimal(18,2)) as Qty,isnull(TSPL_TRANSFER_ORDER_DETAIL.Item_Desc,'') as item_desc "
            'qry1 += "  from TSPL_TRANSFER_ORDER_DETAIL left join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No "
            'qry1 += "  where TSPL_TRANSFER_ORDER_DETAIL.Document_No='" & obj.Document_No & "' ORDER BY TSPL_TRANSFER_ORDER_DETAIL.Line_No"
            'Dim dtDocWise As DataTable = clsDBFuncationality.GetDataTable(qry1, trans)

            'For ii As Integer = 0 To dtDocWise.Rows.Count - 1
            '    itemName = clsCommon.myCstr(dtDocWise.Rows(ii)("item_desc"))
            '    UOM = clsCommon.myCstr(dtDocWise.Rows(ii)("Unit_Code"))
            '    qty = clsCommon.myCstr(dtDocWise.Rows(ii)("Qty"))

            '    ItemDetail += itemName + " " + UOM + "-" + qty + ","
            'Next

            If (obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0) Then
                For Each objdetail As clsTransferDCCDetail In obj.Arr
                    itemName = clsCommon.myCstr(objdetail.Item_Desc)
                    UOM = clsCommon.myCstr(objdetail.Unit_code)
                    qty = clsCommon.myCstr(IIf(obj.Transfer_Type = "O", objdetail.Out_Qty, objdetail.In_Qty))
                    ItemDetail += itemName + " " + UOM + "-" + qty + Environment.NewLine
                Next
            End If

            If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 AndAlso (arrMailID IsNot Nothing AndAlso arrMailID.Count > 0) Then
                Dim objEmailH As New clsEMailHead()
                objEmailH.arrEMail = New List(Of String)()
                objEmailH.arrEMail = arrMailID
                objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.DOC_NO, obj.Document_No)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myCstr(obj.DOC_Total_Amt))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)
                objEmailH.SaveData(clsUserMgtCode.Transfer, objEmailH, trans)
                objEmailH = Nothing

            End If

            If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 AndAlso (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0) Then
                Dim objSMSH As New clsSMSHead()
                objSMSH.arrMobilNo = New List(Of String)()
                objSMSH.arrMobilNo = arrMobileNo
                objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_NO, obj.Document_No)
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myCstr(obj.DOC_Total_Amt))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)
                objSMSH.SaveData(clsUserMgtCode.Transfer, objSMSH, trans)
                objSMSH = Nothing
            End If
        End If
    End Sub

    Private Shared Function CreateProvision(ByVal objTrans As clsTransferDCC, ByVal trans As SqlTransaction) As Boolean
        Dim obj As New clsProvisionEntry()
        Try
            obj = New clsProvisionEntry()
            obj.isNewEntry = True
            obj.Doc_Date = objTrans.Document_Date
            obj.Vendor_Code = objTrans.Transport_Id
            obj.Vendor_Desc = clsTransferDCC.GetTransporterName(objTrans.Transport_Id, trans)
            obj.Vendor_Type = "Transporter For Transfer"
            obj.Status = "No"
            obj.Ref_Doc_No = objTrans.TransferOutNo
            obj.Prov_type = "Freight"
            obj.Amount = objTrans.Vehicle_Charge
            obj.Prog_Code = clsUserMgtCode.Transfer
            obj.Prov_Month = Month(objTrans.Document_Date)
            obj.Prov_Year = Year(objTrans.Document_Date)
            obj.Loc_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + objTrans.TransferOutNo + "'", trans))
            obj.Loc_Desc = clsLocation.GetName(obj.Loc_Code, trans)
            obj.Freight_Type = objTrans.Freight_Type
            obj.FixedCharge = objTrans.FixedCharge
            obj.EmptyCharge = objTrans.EmptyCharge
            If clsProvisionEntry.SaveData(obj, trans) Then
                If clsProvisionEntry.PostData(obj.Doc_No, trans, False) Then

                End If
            End If

            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            obj = Nothing
        End Try

    End Function

    Public Shared Function postTransfer(ByVal strCode As String, ByVal trans As SqlTransaction, Optional ByVal ProvisionAllow As Boolean = False, Optional strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Return postTransfer(False, strCode, trans, ProvisionAllow, strVoucherNoForRecreateOnly)
    End Function

    Public Shared Function postTransfer(ByVal isUpdateDocAccordingToInvtory As Boolean, ByVal strCode As String, ByVal trans As SqlTransaction, Optional ByVal ProvisionAllow As Boolean = False, Optional strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Try
            Dim isSaved As Boolean = True
            Dim sql As String = Nothing
            If (clsCommon.myLen(strCode) <= 0) Then
                Throw New Exception("Transfer No. not found to Post")
            End If
            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")
            Dim obj As clsTransferDCC = clsTransferDCC.GetData(strCode, NavigatorType.Current, trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If obj.Status = ERPTransactionStatus.Cancel Then
                Throw New Exception("Transaction is already Cancelled, so it should not be Posted.")
            End If

            If clsCommon.myCBool(obj.InternalTransfer) = True Then
            Else
                If ProvisionAllow AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Transport_Id) > 0 Then
                    CreateProvision(obj, trans)
                End If
            End If


            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Material Management", "Transfer(Load-Out)", obj.From_Location, obj.Document_Date, trans)
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Material Management", "Transfer(Load-In)", obj.To_Location, obj.Document_Date, trans)
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "R") = CompairStringResult.Equal Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Material Management", "Transfer(Load-Out)", obj.From_Location, obj.Document_Date, trans)
            End If
            If clsCommon.CompairString(obj.Status, "Y") = CompairStringResult.Equal Then
                Throw New Exception("Already Post on :" + clsCommon.GetPrintDate(obj.Modify_Date, "dd/MM/yyyy"))
            End If
            '' code commented by panch raj on 17/12/2014 because inventory movement does not hit on excisable case
            'If clsCommon.CompairString(strExcisable, "F") = CompairStringResult.Equal Then
            HitInvenotryMovement(trans, obj, strPostDate)
            If isUpdateDocAccordingToInvtory Then
                UpdateDocAccordingToInventoryMovement(obj.Document_No, trans)
                obj = clsTransferDCC.GetData(obj.Document_No, NavigatorType.Current, trans)
            End If
            'comment by Balwinder cogs is handled in JE funcion on 02/02/2016
            If Not clsCommon.myCBool(obj.InternalTransfer) Then
                PostTransferJournalEntryOnly(obj, strVoucherNoForRecreateOnly, trans)
            End If
            'Dim ELocationType = clsDBFuncationality.getSingleValue("select case when Registered=1 then 'BB' else 'BC' end as Type from TSPL_LOCATION_MASTER where location_code='" + obj.To_Location + "'", trans)
            Dim ELocationType As String = ""
            If obj.IsJobWorkType = 0 AndAlso (clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal) Then
                ELocationType = clsDBFuncationality.getSingleValue("select case when (CASE WHEN ToLocation.Registered=0 AND  ToLocationGIT.Registered=0 THEN 0 ELSE 1 END)=1 then 'BB' else 'BC' end AS Type from TSPL_TRANSFER_ORDER_HEAD left Outer Join TSPL_LOCATION_MASTER as ToLocation on ToLocation.Location_Code  =TSPL_TRANSFER_ORDER_HEAD.To_Location left Outer Join TSPL_LOCATION_MASTER as ToLocationGIT on ToLocationGIT.GIT_Location  =TSPL_TRANSFER_ORDER_HEAD.To_Location where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + obj.Document_No + "'", trans)
            End If
            Dim qry As String = "Update TSPL_TRANSFER_ORDER_HEAD set Status=1, Posting_Date='" + strPostDate + "',Modify_By='" + objCommonVar.CurrentUserCode + "',EInvoice_Type='" + ELocationType + "' where Document_No='" + obj.Document_No + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function PostTransferJournalEntryOnly(ByVal obj As clsTransferDCC, ByVal strVoucherNoForRecreateOnly As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        If (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateJEForTransfer, clsFixedParameterCode.CreateJEForTransfer, trans)) > 0) Then
            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                CreateTransferOutJE(obj, trans, strVoucherNoForRecreateOnly)
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                CreateTransferInJE(obj, trans, strVoucherNoForRecreateOnly)
                '==============Added by preeti Gupta Against Ticket no[]
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                CreateTransferReturn(obj, trans, strVoucherNoForRecreateOnly)
                '=========================================================================
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "R") = CompairStringResult.Equal Then
                Dim strGITLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER   where Location_Code in (SELECT GIT_Location from TSPL_LOCATION_MASTER WHERE Location_Code='" + obj.From_Location + "')", trans))
                If clsCommon.myLen(strGITLocationSegment) <= 0 Then
                    Throw New Exception("GIT Location not found for location" + obj.From_Location)
                End If
                Dim strLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER   where Location_Code = '" + obj.From_Location + "' ", trans))
                Dim strToLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER   where Location_Code = '" + obj.To_Location + "' ", trans))
                If clsCommon.myLen(strToLocationSegment) <= 0 Then
                    Throw New Exception("Location segment not found for location" + obj.To_Location)
                End If
                Dim strToLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Location_Code  from TSPL_LOCATION_MASTER where GIT_Location='" + obj.To_Location + "'", trans))
                Dim arrlist As New ArrayList()
                For Each objtr As clsTransferDCCDetail In obj.Arr
                    qry = "SELECT TSPL_PURCHASE_ACCOUNTS.Other_1 AS Rejected_Acc FROM TSPL_ITEM_MASTER  INNER JOIN   TSPL_PURCHASE_ACCOUNTS  ON TSPL_ITEM_MASTER.Purchase_Class_Code = TSPL_PURCHASE_ACCOUNTS .Purchase_Class_Code INNER JOIN  TSPL_GL_ACCOUNTS ON TSPL_PURCHASE_ACCOUNTS .Inv_Control_Account = TSPL_GL_ACCOUNTS.Account_Code WHERE TSPL_ITEM_MASTER.Item_Code='" + objtr.Item_Code + "'"
                    Dim Rejected_Acc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                    If clsCommon.myLen(Rejected_Acc) <= 0 Then
                        Throw New Exception("Please set Rejected account for purchase account set")
                    End If

                    Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(Rejected_Acc, strGITLocationSegment, True, trans)
                    Dim fromGITLocation() As String = {strTemp, objtr.Amount}
                    arrlist.Add(fromGITLocation)

                    '' strInvCtrlAcc to location
                    'qry = "SELECT TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account AS strInvCtrlAc FROM TSPL_ITEM_MASTER  INNER JOIN   TSPL_PURCHASE_ACCOUNTS  ON TSPL_ITEM_MASTER.Purchase_Class_Code = TSPL_PURCHASE_ACCOUNTS .Purchase_Class_Code INNER JOIN  TSPL_GL_ACCOUNTS ON TSPL_PURCHASE_ACCOUNTS .Inv_Control_Account = TSPL_GL_ACCOUNTS.Account_Code WHERE TSPL_ITEM_MASTER.Item_Code='" + objtr.Item_Code + "'"
                    qry = "SELECT TSPL_PURCHASE_ACCOUNTS.Stock_Transfer_Acc AS strInvCtrlAc FROM TSPL_ITEM_MASTER  INNER JOIN   TSPL_PURCHASE_ACCOUNTS  ON TSPL_ITEM_MASTER.Purchase_Class_Code = TSPL_PURCHASE_ACCOUNTS .Purchase_Class_Code INNER JOIN  TSPL_GL_ACCOUNTS ON TSPL_PURCHASE_ACCOUNTS .Stock_Transfer_Acc = TSPL_GL_ACCOUNTS.Account_Code WHERE TSPL_ITEM_MASTER.Item_Code='" + objtr.Item_Code + "'"
                    Dim strInvCtrlAc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                    If clsCommon.myLen(Rejected_Acc) <= 0 Then
                        Throw New Exception("Please set Stock Transfer account for purchase account set")
                    End If

                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAc, strToLocationSegment, True, trans)
                    Dim ToLocation() As String = {strTemp, -1 * objtr.Amount}
                    arrlist.Add(ToLocation)
                Next
                If clsCommon.CompairString(obj.Type, "Excise") = CompairStringResult.Equal Then
                    JE_Excisable_Common(obj, arrlist, strLocationSegment, trans)
                End If

                Dim GLDesc As String = "Journal Entry Against Stock Transfer- Doc No." & obj.Document_No & " "
                Dim Remarks As String = "Journal Entry against Stock Transfer from location -" & obj.To_Location & " For Doc No. " & obj.Document_No & ""

                If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , Remarks, obj.Remarks)
                End If
            End If
        End If
        Return True
    End Function

    Public Shared Function postTransferNew(ByVal trans As SqlTransaction, ByVal obj As clsTransferDCC, ByVal isForJournalEntryOnly As Boolean, Optional ByVal ProvisionAllow As Boolean = False, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Return postTransferNew(False, trans, obj, isForJournalEntryOnly, ProvisionAllow, strVoucherNoForRecreateOnly)
    End Function
    Public Shared Function postTransferNew(ByVal isUpdateDocAccordingToInvtory As Boolean, ByVal trans As SqlTransaction, ByVal obj As clsTransferDCC, ByVal isForJournalEntryOnly As Boolean, Optional ByVal ProvisionAllow As Boolean = False, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing) As Boolean
        Try
            If obj.Status = ERPTransactionStatus.Cancel Then
                Throw New Exception("Transaction is already Cancelled, so it should not be Posted.")
            End If

            Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt")
            Dim PostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy")

            Dim LoadInAmt As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" select DOC_Total_Amt from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + obj.Document_No + "'", trans))
            Dim strTranferDate As String = clsCommon.GetPrintDate(clsCommon.myCDate(clsDBFuncationality.getSingleValue("Select Document_Date from TSPL_TRANSFER_ORDER_HEAD WHERE Document_No='" + obj.Document_No + "'", trans)), "dd/MMM/yyyy")

            'strExcisable = clsDBFuncationality.getSingleValue("select excisable from TSPL_LOCATION_MASTER where Location_Code='" + obj.From_Location + "'", trans)
            'strToLType = clsDBFuncationality.getSingleValue("select Location_Type from TSPL_LOCATION_MASTER where Location_Code='" + obj.To_Location + "'", trans)
            'strFromLType = clsDBFuncationality.getSingleValue("select Location_Type from TSPL_LOCATION_MASTER where Location_Code='" + obj.From_Location + "'", trans)

            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Material Management", "Transfer(Load-Out)", obj.From_Location, obj.Document_Date, trans)
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Material Management", "Transfer(Load-In)", obj.To_Location, obj.Document_Date, trans)
            End If

            If ProvisionAllow AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Transport_Id) > 0 Then
                CreateProvision(obj, trans)
            End If

            HitInvenotryMovement(trans, obj, strTranferDate)

            If isUpdateDocAccordingToInvtory Then
                UpdateDocAccordingToInventoryMovement(obj.Document_No, trans)
                obj = clsTransferDCC.GetData(obj.Document_No, NavigatorType.Current, trans)
            End If
            If clsCommon.myCBool(obj.InternalTransfer) = True Then
            Else
                PostTransferNewJournalEntryOnly(obj, strVoucherNoForRecreateOnly, trans)
            End If

            'End If
            Dim ELocationType As String = ""
            If obj.IsJobWorkType = 0 AndAlso (clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal) Then
                ELocationType = clsDBFuncationality.getSingleValue("select case when (CASE WHEN ToLocation.Registered=0 AND  ToLocationGIT.Registered=0 THEN 0 ELSE 1 END)=1 then 'BB' else 'BC' end AS Type from TSPL_TRANSFER_ORDER_HEAD left Outer Join TSPL_LOCATION_MASTER as ToLocation on ToLocation.Location_Code  =TSPL_TRANSFER_ORDER_HEAD.To_Location left Outer Join TSPL_LOCATION_MASTER as ToLocationGIT on ToLocationGIT.GIT_Location  =TSPL_TRANSFER_ORDER_HEAD.To_Location where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + obj.Document_No + "'", trans)
            End If
            Dim qry As String = "Update TSPL_TRANSFER_ORDER_HEAD set Status=1, Posting_Date='" + strPostDate + "',Modify_By='" + objCommonVar.CurrentUserCode + "',EInvoice_Type='" + ELocationType + "' where Document_No='" + obj.Document_No + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Return True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function UpdateDocAccordingToInventoryMovement(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = "update TSPL_TRANSFER_ORDER_detail set Item_Cost=case when x.qty=0 then 0 else  x.Avg_Cost/qty end,Amount=x.Avg_Cost,Amt_Less_Discount=x.Avg_Cost,Item_Net_Amt=x.Avg_Cost" + Environment.NewLine +
        "from (select Source_Doc_No,Item_Code,Qty,UOM,Avg_Cost from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" + strDocNo + "' and InOut='O' and Trans_Type='Transfer')x" + Environment.NewLine +
        "inner join TSPL_TRANSFER_ORDER_detail on TSPL_TRANSFER_ORDER_detail.document_no =x.Source_Doc_No and TSPL_TRANSFER_ORDER_detail.Item_Code =x.Item_Code  and TSPL_TRANSFER_ORDER_detail.Unit_code =x.UOM"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "update TSPL_TRANSFER_ORDER_HEAD set Discount_Base=x.Amount,Amount_Less_Discount=x.Amount,DOC_Total_Amt=x.Amount,Total_Amt_Less_Tax=x.Amount 
        from (select max(TSPL_TRANSFER_ORDER_detail.document_no) as document_no, sum(Amount) as Amount from TSPL_TRANSFER_ORDER_detail where TSPL_TRANSFER_ORDER_detail.document_no='" + strDocNo + "')x
        inner join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=x.document_no"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Return True
    End Function

    Public Shared Function PostTransferNewJournalEntryOnly(ByVal obj As clsTransferDCC, ByVal strVoucherNoForRecreateOnly As String, ByVal trans As SqlTransaction) As Boolean
        Dim settStockTranferFromTransferPriceAndInvJVWithAvgCost As Boolean = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.StockTranferFromTransferPriceAndInvJVWithAvgCost, clsFixedParameterCode.StockTranferFromTransferPriceAndInvJVWithAvgCost, trans)) = 1)
        Dim strExcisable As String = clsDBFuncationality.getSingleValue("select excisable from TSPL_LOCATION_MASTER where Location_Code='" + obj.From_Location + "'", trans)
        If strExcisable = "F" Then
            Dim Sql As String = ""
            Dim strFromInvAcc As String = ""
            Dim strFromInvAccDesc As String = ""
            Dim strToInvAcc As String = ""
            Dim strToInvAccDesc As String = ""
            Dim strShpClrAcc As String = ""
            Dim strShpClrAccDesc As String = ""
            Sql = "SELECT Loc_Segment_Code from TSPL_LOCATION_MASTER WHERE Location_Code='" + Convert.ToString(obj.From_Location) + "'"
            Dim fromLocSegCode As String = connectSql.RunScalar(trans, Sql)
            Sql = "SELECT Loc_Segment_Code from TSPL_LOCATION_MASTER WHERE Location_Code='" + Convert.ToString(obj.To_Location) + "'"
            Dim toLocSegCode As String = connectSql.RunScalar(trans, Sql)
            ''richa agarwal 24 Jan,2020
            Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)
            Dim arrlist As New ArrayList()
            For Each objtr As clsTransferDCCDetail In obj.Arr
                Sql = "SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + objtr.Item_Code + "'"
                strFromInvAcc = connectSql.RunScalar(trans, Sql)
                strFromInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strFromInvAcc, obj.From_Location, trans)
                strToInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strFromInvAcc, obj.To_Location, trans)
                Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strFromInvAcc + "'"
                strFromInvAccDesc = connectSql.RunScalar(trans, Sql)
                If strFromInvAccDesc Is Nothing Then
                    Throw New Exception("Inventory Control Account " + strFromInvAcc + " not found.")
                End If
                Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strToInvAcc + "'"
                strToInvAccDesc = connectSql.RunScalar(trans, Sql)
                If strToInvAccDesc Is Nothing Then
                    Throw New Exception("Inventory Control Account " + strToInvAcc + " not found.")
                End If
                Dim objSeg As Accountsegment = Accountsegment.Getaccountcodedesc(strFromInvAcc, trans)

                If settStockTranferFromTransferPriceAndInvJVWithAvgCost AndAlso clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                    Sql = "select avg_Cost  from tspl_inventory_movement where source_doc_no='" + obj.Document_No + "' and Trans_Type='Transfer' and inOut='O' and item_code='" + objtr.Item_Code + "'"
                    Dim dblAvgCost As Decimal = clsDBFuncationality.getSingleValue(Sql, trans)

                    'Dim BankAcc() As String = {strFromInvAcc, dblAvgCost * (-1)}
                    Dim BankAcc() As String = {strFromInvAcc, dblAvgCost * (-1), "", "", "", "", "", "", "I"}
                    arrlist.Add(BankAcc)

                    Sql = "select TSPL_ITEM_MASTER.Sale_Class_Code, TSPL_SALES_ACCOUNTS.Cost_Of_Goods_Sold from TSPL_ITEM_MASTER" + Environment.NewLine +
                    "left outer join TSPL_SALES_ACCOUNTS on TSPL_SALES_ACCOUNTS.Sales_Class_Code=TSPL_ITEM_MASTER.Sale_Class_Code " + Environment.NewLine +
                    "where TSPL_ITEM_MASTER.Item_Code='" + objtr.Item_Code + "'"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Sql, trans)
                    If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                        Throw New Exception("Not a Valid Sales Account set for Item [" + objtr.Item_Code + "]")
                    End If
                    Dim strCogsAc As String = clsCommon.myCstr(dt.Rows(0)("Cost_Of_Goods_Sold"))
                    If clsCommon.myLen(strCogsAc) <= 0 Then
                        Throw New Exception("Please set Cost of Goods Sold Account Of Sales Account Set [" + clsCommon.myCstr(dt.Rows(0)("Sale_Class_Code")) + "] and Item [" + objtr.Item_Code + "].")
                    End If
                    strCogsAc = clsERPFuncationality.ChangeGLAccountLocationSegment(strCogsAc, obj.From_Location, trans)
                    Dim BankAcc1() As String = {strCogsAc, (dblAvgCost - objtr.Item_Net_Amt)}
                    arrlist.Add(BankAcc1)
                Else
                    If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                        Dim BankAcc() As String = {strFromInvAcc, objtr.Item_Net_Amt * (-1), "", "", "", "", "", "", "I"}
                        arrlist.Add(BankAcc)
                    Else
                        Dim BankAcc() As String = {strFromInvAcc, objtr.Item_Net_Amt * (-1)}
                        arrlist.Add(BankAcc)
                    End If
                End If




                strToInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strToInvAcc, obj.To_Location, trans)
                objSeg = Accountsegment.Getaccountcodedesc(strToInvAcc, trans)

                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    Dim DocAmtWithTax() As String = {strToInvAcc, objtr.Item_Net_Amt, "", "", "", "", "", "", "I"}
                    arrlist.Add(DocAmtWithTax)
                Else
                    Dim DocAmtWithTax() As String = {strToInvAcc, objtr.Item_Net_Amt}
                    arrlist.Add(DocAmtWithTax)
                End If


                If PickTCAForStockTransferAndTankerDispatch = True Then
                    Dim strTempAC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                    " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + objtr.Item_Code + "'", trans))
                    If clsCommon.myLen(strTempAC) <= 0 Then
                        Throw New Exception("Please Map Transfer Clearing Account For  for item " + objtr.Item_Code.ToString() + "  (" & objtr.Item_Desc.ToString() & ")")
                    End If
                    If clsCommon.myLen(strTempAC) > 0 Then
                        strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, fromLocSegCode, True, trans)
                        Dim acc3() As String = {strTempAC, objtr.Item_Net_Amt}
                        arrlist.Add(acc3)

                        strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, toLocSegCode, True, trans)
                        Dim acc4() As String = {strTempAC, -1 * objtr.Item_Net_Amt}
                        arrlist.Add(acc4)
                    End If
                Else
                    Dim strTempAC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Transfer_Clearing FROM TSPL_ITEM_MASTER INNER JOIN TSPL_PURCHASE_ACCOUNTS on TSPL_ITEM_MASTER.Purchase_Class_Code = TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objtr.Item_Code + "'", trans))
                    If clsCommon.myLen(strTempAC) > 0 Then
                        strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, fromLocSegCode, True, trans)
                        Dim acc3() As String = {strTempAC, objtr.Item_Net_Amt}
                        arrlist.Add(acc3)

                        strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, toLocSegCode, True, trans)
                        Dim acc4() As String = {strTempAC, -1 * objtr.Item_Net_Amt}
                        arrlist.Add(acc4)
                    End If
                End If
            Next

#Region "Taxes"
            Dim taxAmt As Decimal
            Dim ttlTotalTaxAmt As Decimal = 0
            Dim strTaxCode As String
            Dim strNetPayAcc As String = ""
            Dim strNetPayAccDesc As String = ""

            '''' Tax1 ''''''***********************
            strTaxCode = obj.TAX1
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX1_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX1_Amt), 2)
                Else
                    If obj.TAX1_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX1_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    '' Anubhooti 09-Sep-2014 (Not Passed trans Replace(connectSql.RunScalar(Sql))
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax1

            '''' Tax2 ''''''***********************
            strTaxCode = obj.TAX2
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX2_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX2_Amt), 2)
                Else
                    If obj.TAX2_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX2_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax2

            '''' Tax3 ''''''***********************
            strTaxCode = obj.TAX3
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX3_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX3_Amt), 2)
                Else
                    If obj.TAX3_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX3_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax3

            '''' Tax4 ''''''***********************
            strTaxCode = obj.TAX4
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX4_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX4_Amt), 2)
                Else
                    If obj.TAX4_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX4_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    ' connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax4

            '''' Tax5 ''''''***********************
            strTaxCode = obj.TAX5
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX5_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX5_Amt), 2)
                Else
                    If obj.TAX5_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX5_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax5

            '''' Tax6 ''''''***********************
            strTaxCode = obj.TAX6
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX6_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX6_Amt), 2)
                Else
                    If obj.TAX5_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX6_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax6

            '''' Tax7 ''''''***********************
            strTaxCode = obj.TAX7
            If clsCommon.myLen(strTaxCode) > 0 Then
                If obj.TAX7_Amt.ToString().Substring(0, 1) = "-" Then
                    taxAmt = Math.Round(CDec(obj.TAX7_Amt), 2)
                Else
                    If obj.TAX5_Amt = 0 Then
                        taxAmt = 0
                    Else
                        taxAmt = Math.Round(CDec(obj.TAX7_Amt), 2)
                    End If
                End If
                Sql = "SELECT Tax_Recoverable_Account FROM TSPL_TAX_MASTER WHERE Tax_Code = '" + strTaxCode + "'"
                If Not connectSql.RunScalar(trans, Sql).ToString() = "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(connectSql.RunScalar(trans, Sql), obj.From_Location, trans)
                    'strNetPayAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
                End If
                If Not strNetPayAcc = "" Then
                    Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strNetPayAcc + "'"
                    strNetPayAccDesc = connectSql.RunScalar(trans, Sql)
                End If
                If Not taxAmt = 0 AndAlso strNetPayAcc <> "" Then
                    strNetPayAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strNetPayAcc, obj.From_Location, trans)
                    Dim obj1 As Accountsegment = Accountsegment.Getaccountcodedesc(strNetPayAcc, trans)
                    'connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strNetPayAcc), New SqlParameter("@Account_Desc", strNetPayAccDesc), New SqlParameter("@Amount", taxAmt * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", obj1.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", obj1.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", obj1.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", obj1.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", obj1.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", obj1.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", obj1.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", obj1.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", obj1.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", obj1.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", obj1.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", obj1.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", obj1.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", obj1.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", obj1.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", obj1.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", obj1.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", obj1.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", obj1.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", obj1.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", obj1.Account_Seg_Desc10))
                    'lineNo = lineNo + 1
                    Dim Tax1() As String = {strNetPayAcc, taxAmt * (-1)}
                    arrlist.Add(Tax1)
                End If
                ttlTotalTaxAmt = ttlTotalTaxAmt + taxAmt
            End If
            '*********** End Tax7
#End Region

            If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then
                If ((clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal)) Then
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)
                End If
            Else
                If ((clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal)) Then
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)
                End If
            End If
        End If
        Return True
    End Function


    Public Shared Function CreateTransferInJE(obj As clsTransferDCC, trans As SqlTransaction, strVoucherNoForRecreateOnly As String) As Boolean
        Dim isSkipCogsGL As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, trans)) = 0, False, True)
        Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)

        Dim rValue As Boolean = False
        Try
            Dim COGT_AC As String = String.Empty
            Dim Stock_Transfer_Ac As String = String.Empty
            Dim Inventory_Control_Ac As String = String.Empty
            Dim Branch_Ac As String = String.Empty
            Dim Inventory_Control_Ac_FromLoc As String = String.Empty
            Dim Inventory_Control_Ac_ToLoc As String = String.Empty
            Dim Inventory_Control_Ac_GITLoc As String = String.Empty
            Dim Stock_Transfer_Ac_FromLoc As String = String.Empty
            Dim Stock_Transfer_Ac_GITLoc As String = String.Empty
            Dim GIT_LOC As String = String.Empty
            Dim CostingMethod As Integer = 0
            Dim CostOfItem As Double = 0
            Dim dt As Date = clsCommon.GETSERVERDATE(trans)
            Dim FromLocSeg As String = String.Empty
            Dim ToLocSeg As String = String.Empty
            If obj.Status = 1 Then
                dt = obj.Posting_Date
            End If
            Dim FromLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location  from TSPL_TRANSFER_ORDER_HEAD where Document_No  ='" & obj.TransferOutNo & "'", trans))
            FromLocSeg = FromLocation   'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & obj.From_Location & "'", trans))
            ToLocSeg = obj.To_Location    'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & obj.To_Location & "'", trans))

            Dim FromLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans))
            Dim ToLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans))

            Dim TransitLossAc As String = ""
            Dim Branch_AcFromLoc As String = ""
            'GIT_LOC = obj.From_Location  'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GIT_Location  from tspl_location_master where location_code='" & FromLocation & "'", trans))
            'If clsCommon.myLen(GIT_LOC) <= 0 Then
            'Throw New Exception("Please Map GIT Location For Location : " & FromLocation)
            'End If
            Dim GIT_LOC_SEG As String = String.Empty
            ''richa agarwal 15 nov ,2016
            Dim IGnoreGITAccount As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, trans))
            If IGnoreGITAccount = "0" Then
                GIT_LOC = obj.From_Location
                GIT_LOC_SEG = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & GIT_LOC & "'", trans))
                If clsCommon.myLen(GIT_LOC_SEG) <= 0 Then
                    Throw New Exception(" Location Segment Not Found in Location Master, For Location : " & GIT_LOC)
                End If
                GIT_LOC = GIT_LOC_SEG
            End If
            ''--------------------------

            Dim qry As String = ""

            If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                Dim ArryLst As ArrayList = New ArrayList()
                For i As Integer = 0 To obj.Arr.Count - 1
                    If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL
                        CostingMethod = clsInventoryMovementNew.getCostingMethod(obj.Arr(i).Item_Code, trans)
                        'qry = " select " & IIf(CostingMethod = 1, "avg_cost", IIf(CostingMethod = 2, "FIFO_Cost", IIf(CostingMethod = 3, "LIFO_Cost", " 0 "))) & " from tspl_inventory_movement where source_doc_no='" & obj.TransferOutNo & "' and Item_Code='" & obj.Arr(i).Item_Code & "' and InOut='O' and Trans_Type='Transfer' "
                        ' CostOfItem = clsInventoryMovement.GetCost(CostingMethod, obj.Arr(i).Item_Code, obj.From_Location, obj.Arr(i).Out_Qty, obj.Document_Date, clsCommon.GetPrintDate(dt, "dd/MMM/yyyy"), True, trans, "tspl_inventory_movement")
                        CostOfItem = Math.Round((obj.Arr(i).Out_Qty * obj.Arr(i).Item_Cost), 2)
                    Else
                        CostOfItem = 0
                    End If  '' Done By Pankaj Jha For Skipping Cogs GL
                    Inventory_Control_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_Transfer_In from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Inventory_Control_Ac) <= 0 Then
                        Throw New Exception("Please Map Stock Transfer In A/C From Purchase Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If
                    'If CostOfItem > 0 Then
                    '    COGT_AC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select COGT_AC from  tspl_sales_accounts where Sales_class_code=(select Sale_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    '    If clsCommon.myLen(COGT_AC) <= 0 Then
                    '        Throw New Exception("Please Map Cost Of Goods Transfer A/C From Sales Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    '    End If
                    '    COGT_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(COGT_AC, FromLocSeg, True, trans)
                    '    Inventory_Control_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, FromLocSeg, True, trans)
                    'End If
                    Inventory_Control_Ac_ToLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, ToLocationSegment, True, trans)
                    Inventory_Control_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, FromLocationSegment, True, False, trans)


                    ''richa agarwal 16 Jan,2020
                    If PickTCAForStockTransferAndTankerDispatch = True Then
                        Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                        " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(i).Item_Code + "'", trans))
                        If clsCommon.myLen(Branch_Ac) <= 0 Then
                            Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(i).Item_Code + "  (" & obj.Arr(i).Item_Desc & ")")
                        End If
                        Branch_Ac = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac, ToLocationSegment, True, trans)
                    Else
                        Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "'", trans))
                        If clsCommon.myLen(Branch_Ac) <= 0 Then
                            Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)))
                        End If
                    End If

                    Dim TransferGainLossAc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Transfer_Gain_Loss_Ac from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(TransferGainLossAc) <= 0 Then
                        Throw New Exception("Please Map Transfer Profit/Loss A/c For Item : " & obj.Arr(i).Item_Code)
                    End If

                    'TransitLossAc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Transfer_Gain_Loss_Ac from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    'If clsCommon.myLen(TransitLossAc) <= 0 Then
                    '    Throw New Exception("Please Map Transit Loss A/c For Item : " & obj.Arr(i).Item_Code)
                    'End If

                    Stock_Transfer_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_transfer_Ac from  tspl_sales_accounts where Sales_class_code=(select Sale_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Stock_Transfer_Ac) <= 0 Then
                        Throw New Exception("Please Map Stock Transfer A/C From Sales Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If

                    ''richa agarwal 15 nov ,2016
                    If IGnoreGITAccount = "0" Then
                        Inventory_Control_Ac_GITLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, GIT_LOC, True, False, trans)
                        Stock_Transfer_Ac_GITLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_Ac, GIT_LOC, True, False, trans)
                    End If

                    'Stock_Transfer_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_Ac, FromLocSeg, True, trans)
                    'Branch_Ac = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac, FromLocSeg, True, trans)

                    Dim Amt As Double = clsCommon.myCdbl(obj.Arr(i).Item_Cost * obj.Arr(i).In_Qty)
                    Dim OutAmt As Double = clsCommon.myCdbl(obj.Arr(i).Item_Cost * obj.Arr(i).Out_Qty)

                    ArryLst.Add(New String() {Branch_Ac, Amt * -1})
                    ''richa BHA/27/11/18-000714 10 Dec,2018
                    Dim strItemProductType As String = clsItemMaster.GetItemProductType(obj.Arr(i).Item_Code, trans)
                    If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = "1", True, False)) = True Then
                        ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, Amt, "", "", "", "", "", "", "P"})
                    Else
                        ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, Amt, "", "", "", "", "", "", "I"})
                        clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, Inventory_Control_Ac_ToLoc, "", "I", trans)
                    End If
                    ''------------------

                    ''richa agarwal 15 nov ,2016
                    If IGnoreGITAccount = "0" Then
                        ArryLst.Add(New String() {Inventory_Control_Ac_GITLoc, Amt * -1})

                        ArryLst.Add(New String() {Stock_Transfer_Ac_GITLoc, Amt})
                    End If
                    TransferGainLossAc = clsERPFuncationality.ChangeGLAccountLocationSegment(TransferGainLossAc, ToLocationSegment, True, trans)
                    Dim DiffAmt As Double = 0

                    If CostOfItem > OutAmt Then
                        DiffAmt = Math.Abs(CostOfItem - OutAmt)
                        ''richa BHA/27/11/18-000714 10 Dec,2018
                        If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = "1", True, False)) = True Then
                            ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, DiffAmt, "", "", "", "", "", "", "P"})
                        Else
                            ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, DiffAmt, "", "", "", "", "", "", "I"})
                            clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, Inventory_Control_Ac_ToLoc, "", "I", trans)
                        End If
                        ''------------------
                        ArryLst.Add(New String() {TransferGainLossAc, DiffAmt * -1})
                    ElseIf CostOfItem < OutAmt AndAlso CostOfItem <> 0 Then
                        DiffAmt = Math.Abs(OutAmt - CostOfItem)
                        ''richa BHA/27/11/18-000714 10 Dec,2018
                        If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = "1", True, False)) = True Then
                            ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, DiffAmt * -1, "", "", "", "", "", "", "P"})
                        Else
                            ArryLst.Add(New String() {Inventory_Control_Ac_ToLoc, DiffAmt * -1, "", "", "", "", "", "", "I"})
                            clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, "", Inventory_Control_Ac_ToLoc, "I", trans)
                        End If
                        ''------------------

                        ArryLst.Add(New String() {TransferGainLossAc, DiffAmt})
                    End If
                    'ArryLst.Add(New String() {COGT_AC, CostOfItem})
                    'ArryLst.Add(New String() {Inventory_Control_Ac_FromLoc, CostOfItem * -1})
                    If obj.Arr(i).Out_Qty <> obj.Arr(i).In_Qty Then

                        If obj.Arr(i).Out_Qty < obj.Arr(i).In_Qty Then
                            TransitLossAc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Transfer_Gain_Loss_Ac from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                            If clsCommon.myLen(TransitLossAc) <= 0 Then
                                Throw New Exception("Please Map Transit Loss A/c For Item : " & obj.Arr(i).Item_Code)
                            End If
                            TransitLossAc = clsERPFuncationality.ChangeGLAccountLocationSegment(TransitLossAc, FromLocationSegment, True, False, trans)
                            Branch_AcFromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "'", trans))
                            If clsCommon.myLen(Branch_AcFromLoc) <= 0 Then
                                Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & FromLocationSegment & " To " & ToLocSeg)
                            End If
                            DiffAmt = (obj.Arr(i).In_Qty * obj.Arr(i).Item_Cost) - (obj.Arr(i).Out_Qty * obj.Arr(i).Item_Cost)
                            ArryLst.Add(New String() {TransitLossAc, DiffAmt * -1})
                            ArryLst.Add(New String() {Branch_AcFromLoc, DiffAmt})
                        ElseIf obj.Arr(i).Out_Qty > obj.Arr(i).In_Qty Then
                            TransitLossAc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Transfer_Gain_Loss_Ac from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                            If clsCommon.myLen(TransitLossAc) <= 0 Then
                                Throw New Exception("Please Map Transit Loss A/c For Item : " & obj.Arr(i).Item_Code)
                            End If
                            TransitLossAc = clsERPFuncationality.ChangeGLAccountLocationSegment(TransitLossAc, FromLocationSegment, True, False, trans)
                            Branch_AcFromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "'", trans))
                            If clsCommon.myLen(Branch_AcFromLoc) <= 0 Then
                                Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & FromLocationSegment & " To " & ToLocSeg)
                            End If
                            DiffAmt = (obj.Arr(i).Out_Qty * obj.Arr(i).Item_Cost) - (obj.Arr(i).In_Qty * obj.Arr(i).Item_Cost)
                            ArryLst.Add(New String() {TransitLossAc, DiffAmt})
                            ArryLst.Add(New String() {Branch_AcFromLoc, DiffAmt * -1})
                        End If
                    End If

                    'rValue = clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, True, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , obj.Remarks, obj.Remarks)
                Next
                If (clsCommon.CompairString(obj.Type, "Excise") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Type, "Mandi") = CompairStringResult.Equal) OrElse (clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True) Then
                    JE_Excisable_Common(obj, ArryLst, ToLocationSegment, trans)
                End If

                Dim GLDesc As String = "Journal Entry Against Stock Transfer- Doc No." & obj.Document_No & " "
                Dim Remarks As String = "Journal Entry against Stock Transfer from location -" & FromLocation & " For Doc No. " & obj.Document_No & ", Transfer Out Doc No: " & obj.TransferOutNo

                If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.To_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.To_Location, False, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return rValue
    End Function

    Public Shared Function CreateTransferOutJE(obj As clsTransferDCC, trans As SqlTransaction, strVoucherNoForRecreateOnly As String) As Boolean
        Dim rValue As Boolean = False
        Dim isSkipCogsGL As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SkipCogsEntry, clsFixedParameterCode.SkipCogsEntry, trans)) = 0, False, True)
        Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)
        ''TCA means Transfer Clearing Account
        Try
            Dim COGT_AC As String = String.Empty
            Dim Stock_Transfer_Ac As String = String.Empty
            Dim Inventory_Control_Ac As String = String.Empty
            Dim Branch_Ac As String = String.Empty
            Dim Inventory_Control_Ac_FromLoc As String = String.Empty
            Dim Inventory_Control_Ac_GITLoc As String = String.Empty
            Dim Tax_Liability_Account As String = String.Empty
            Dim Recoverabe_Account As String = String.Empty
            Dim Stock_Transfer_Ac_FromLoc As String = String.Empty
            Dim Stock_Transfer_Ac_GITLoc As String = String.Empty
            Dim GIT_LOC As String = String.Empty
            Dim CostingMethod As Integer = 0
            Dim CostOfItem As Double = 0
            Dim dt As Date = clsCommon.GETSERVERDATE(trans)
            Dim FromLocSeg As String = String.Empty
            Dim ToLocSeg As String = String.Empty
            If obj.Status = 1 Then
                dt = obj.Posting_Date
            End If

            FromLocSeg = obj.From_Location
            ToLocSeg = obj.To_Location_Main

            Dim FromLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans))
            Dim ToLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans))

            Dim GIT_LOC_SEG As String = String.Empty
            ''richa agarwal 15 nov ,2016
            Dim IGnoreGITAccount As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.IGnoreGITAccount, clsFixedParameterCode.IGnoreGITAccount, trans))
            If IGnoreGITAccount = "0" Then
                GIT_LOC = obj.To_Location
                GIT_LOC_SEG = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & GIT_LOC & "'", trans))
                If clsCommon.myLen(GIT_LOC_SEG) <= 0 Then
                    Throw New Exception(" Location Segment Not Found in Location Master, For Location : " & GIT_LOC)
                End If
                GIT_LOC = GIT_LOC_SEG
            End If
            If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                Dim ArryLst As ArrayList = New ArrayList()
                For i As Integer = 0 To obj.Arr.Count - 1
                    If Not isSkipCogsGL Then    '' Done By Pankaj Jha For Skipping Cogs GL
                        CostingMethod = clsInventoryMovementNew.getCostingMethod(obj.Arr(i).Item_Code, trans)
                        CostOfItem = (obj.Arr(i).Out_Qty * obj.Arr(i).Item_Cost)
                    Else
                        CostOfItem = 0
                    End If  '' Done By Pankaj Jha For Skipping Cogs GL
                    Inventory_Control_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_transfer_Acc from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Inventory_Control_Ac) <= 0 Then
                        Throw New Exception("Please Map  Stock Transfer A/C From Purchase Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If
                    If CostOfItem > 0 Then
                        COGT_AC = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select COGT_AC from  tspl_sales_accounts where Sales_class_code=(select Sale_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                        If clsCommon.myLen(COGT_AC) <= 0 Then
                            Throw New Exception("Please Map Cost Of Goods Transfer A/C From Sales Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                        End If
                        COGT_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(COGT_AC, FromLocationSegment, True, trans)
                        Inventory_Control_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, FromLocationSegment, True, trans)
                    End If
                    ''richa agarwal 16 Jan,2020
                    If PickTCAForStockTransferAndTankerDispatch = True Then
                        Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                        " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(i).Item_Code + "'", trans))
                        If clsCommon.myLen(Branch_Ac) <= 0 Then
                            Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(i).Item_Code + "  (" & obj.Arr(i).Item_Desc & ")")
                        End If
                        Branch_Ac = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac, FromLocationSegment, True, trans)
                    Else
                        Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "'", trans))
                        If clsCommon.myLen(Branch_Ac) <= 0 Then
                            Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)))
                        End If
                    End If
                    Stock_Transfer_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_transfer_Ac from  tspl_sales_accounts where Sales_class_code=(select Sale_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Stock_Transfer_Ac) <= 0 Then
                        Throw New Exception("Please Map Stock Transfer A/C From Sales Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If
                    ''richa 15 Nov,2016
                    If IGnoreGITAccount = "0" Then
                        Inventory_Control_Ac_GITLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Inventory_Control_Ac, GIT_LOC, True, False, trans)
                        Stock_Transfer_Ac_GITLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_Ac, GIT_LOC, True, False, trans)
                    End If
                    Stock_Transfer_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_Ac, FromLocationSegment, True, trans)
                    Dim Amt As Double = clsCommon.myCdbl(obj.Arr(i).Out_Qty * obj.Arr(i).Item_Cost)
                    ArryLst.Add(New String() {Branch_Ac, Amt})
                    ''richa BHA/27/11/18-000714 10 Dec,2018
                    Dim strItemProductType As String = clsItemMaster.GetItemProductType(obj.Arr(i).Item_Code, trans)
                    If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = "1", True, False)) = True Then
                        ArryLst.Add(New String() {Stock_Transfer_Ac_FromLoc, Amt * -1, "", "", "", "", "", "", "S"})
                    Else
                        ArryLst.Add(New String() {Stock_Transfer_Ac_FromLoc, Amt * -1, "", "", "", "", "", "", "I"})
                        clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, "", Stock_Transfer_Ac_FromLoc, "O", trans)
                    End If
                    ''------------------
                    ''richa 15 Nov,2016
                    If IGnoreGITAccount = "0" Then
                        ArryLst.Add(New String() {Inventory_Control_Ac_GITLoc, Amt})
                        ArryLst.Add(New String() {Stock_Transfer_Ac_GITLoc, Amt * -1})
                    End If
                Next
                If (clsCommon.CompairString(obj.Type, "Excise") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Type, "Mandi") = CompairStringResult.Equal) OrElse clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True Then
                    JE_Excisable_Common(obj, ArryLst, FromLocationSegment, trans)
                End If
                Dim GLDesc As String = "Journal Entry Against Stock Transfer- Doc No." & obj.Document_No & " "
                Dim Remarks As String = "Journal Entry against Stock Transfer from location -" & obj.To_Location & " For Doc No. " & obj.Document_No & ""
                If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then ''because if voucher no known then recreate it with same no.
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return rValue
    End Function
    Public Shared Function CreateTransferReturn(obj As clsTransferDCC, trans As SqlTransaction, strVoucherNoForRecreateOnly As String) As Boolean
        Dim rValue As Boolean = False
        Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)
        Try

            Dim Stock_Transfer_AcOut As String = String.Empty
            Dim Stock_Transfer_AcIn As String = String.Empty
            Dim Stock_Transfer_Ac_FromLoc As String = String.Empty
            Dim Stock_Transfer_Ac_ToLoc As String = String.Empty
            Dim Branch_Ac_FromLoc As String = String.Empty
            Dim Branch_Ac_ToLoc As String = String.Empty
            Dim dt As Date = clsCommon.GETSERVERDATE(trans)
            Dim FromLocSeg As String = String.Empty
            Dim ToLocSeg As String = String.Empty
            If obj.Status = 1 Then
                dt = obj.Posting_Date
            End If
            Dim FromLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location  from TSPL_TRANSFER_ORDER_HEAD where Document_No  ='" & obj.Document_No & "'", trans))
            FromLocSeg = FromLocation
            ToLocSeg = obj.To_Location
            'Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "'", trans))
            'If clsCommon.myLen(Branch_Ac) <= 0 Then
            '    Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)))
            'End If
            Dim FromLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans))
            Dim ToLocationSegment As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans))
            Dim qry As String = ""
            If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                Dim ArryLst As ArrayList = New ArrayList()
                For i As Integer = 0 To obj.Arr.Count - 1
                    Stock_Transfer_AcIn = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_Transfer_In from  tspl_purchase_accounts where purchase_class_code=(select purchase_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Stock_Transfer_AcIn) <= 0 Then
                        Throw New Exception("Please Map Stock Transfer In A/C From Purchase Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If
                    Stock_Transfer_AcOut = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stock_transfer_Ac from  tspl_sales_accounts where Sales_class_code=(select Sale_class_code  from tspl_item_master where Item_Code='" & obj.Arr(i).Item_Code & "') ", trans))
                    If clsCommon.myLen(Stock_Transfer_AcOut) <= 0 Then
                        Throw New Exception("Please Map Stock Transfer A/C From Sales Account Set For Item : " & obj.Arr(i).Item_Code & " (" & obj.Arr(i).Item_Desc & ")")
                    End If

                    ''richa agarwal 16 Jan,2020
                    If PickTCAForStockTransferAndTankerDispatch = True Then
                        Branch_Ac_FromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                        " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(i).Item_Code + "'", trans))
                        If clsCommon.myLen(Branch_Ac_FromLoc) <= 0 Then
                            Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(i).Item_Code + "  (" & obj.Arr(i).Item_Desc & ")")
                        End If
                        Branch_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac_FromLoc, FromLocationSegment, True, trans)

                        Branch_Ac_ToLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                    " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(i).Item_Code + "'", trans))
                        If clsCommon.myLen(Branch_Ac_ToLoc) <= 0 Then
                            Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(i).Item_Code + "  (" & obj.Arr(i).Item_Desc & ")")
                        End If
                        Branch_Ac_ToLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac_ToLoc, ToLocationSegment, True, trans)
                    Else
                        'debit
                        Branch_Ac_FromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "'", trans))
                        If clsCommon.myLen(Branch_Ac_FromLoc) <= 0 Then
                            Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)))
                        End If

                        'credit
                        Branch_Ac_ToLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "'", trans))
                        If clsCommon.myLen(Branch_Ac_ToLoc) <= 0 Then
                            Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)))
                        End If
                    End If


                    Stock_Transfer_Ac_ToLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_AcIn, ToLocationSegment, True, False, trans) ''UDL/24/07/19-000308 by balwinder on 31/07/2019
                    Stock_Transfer_Ac_FromLoc = clsERPFuncationality.ChangeGLAccountLocationSegment(Stock_Transfer_AcOut, FromLocationSegment, True, trans)

                    Dim Amt As Double = clsCommon.myCdbl(obj.Arr(i).Item_Cost * obj.Arr(i).In_Qty)
                    Dim OutAmt As Double = clsCommon.myCdbl(obj.Arr(i).Item_Cost * obj.Arr(i).Out_Qty)
                    Dim BranchAmt As Double = clsCommon.myCdbl(obj.Arr(i).Item_Net_Amt)

                    ''richa BHA/27/11/18-000714 10 Dec,2018
                    Dim strItemProductType As String = clsItemMaster.GetItemProductType(obj.Arr(i).Item_Code, trans)
                    If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = "1", True, False)) = True Then
                        ArryLst.Add(New String() {Stock_Transfer_Ac_ToLoc, Amt, "", "", "", "", "", "", "P"})
                        ArryLst.Add(New String() {Stock_Transfer_Ac_FromLoc, Amt * -1, "", "", "", "", "", "", "S"})
                    Else
                        ArryLst.Add(New String() {Stock_Transfer_Ac_ToLoc, Amt, "", "", "", "", "", "", "I"})
                        ArryLst.Add(New String() {Stock_Transfer_Ac_FromLoc, Amt * -1, "", "", "", "", "", "", "I"})

                        clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, "", Stock_Transfer_Ac_FromLoc, "O", trans)
                        clsInventoryMovement.UpdateInvControlAccount(obj.Document_No, "Transfer", obj.Arr(i).Item_Code, Stock_Transfer_Ac_ToLoc, "", "I", trans)
                    End If
                    ''------------------

                    ArryLst.Add(New String() {Branch_Ac_FromLoc, BranchAmt})
                    ArryLst.Add(New String() {Branch_Ac_ToLoc, BranchAmt * -1})
                Next
                If (clsCommon.CompairString(obj.Type, "Excise") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Type, "Mandi") = CompairStringResult.Equal) OrElse (clsERPFuncationality.GetGSTStatus(obj.Document_Date) = True) Then
                    JE_Excisable_CommonReturnTransfer(obj, ArryLst, ToLocationSegment, FromLocationSegment, trans)

                End If

                Dim GLDesc As String = "Journal Entry Against Stock Transfer- Doc No." & obj.Document_No & " "
                Dim Remarks As String = "Journal Entry against Stock Transfer from location -" & FromLocation & " For Doc No. " & obj.Document_No & ", Transfer Out Doc No: " & obj.TransferOutNo

                If strVoucherNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreateOnly) > 0 Then
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.To_Location, False, strVoucherNoForRecreateOnly, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                Else
                    clsJournalMaster.FunGrnlEntryWithTrans(obj.To_Location, False, trans, obj.Document_Date, GLDesc, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, , Remarks, obj.Remarks)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return rValue
    End Function


    Public Shared Function JE_Excisable_Mapping_Off(ByVal obj As clsTransferDCC, ByVal strLocationSegment As String, ByVal trans As SqlTransaction) As Boolean

        Dim Sql As String = ""
        Dim strFromInvAcc As String = ""
        Dim strFromInvAccDesc As String = ""
        Dim strToInvAcc As String = ""
        Dim strToInvAccDesc As String = ""
        Dim strShpClrAcc As String = ""
        Dim strShpClrAccDesc As String = ""
        Sql = "SELECT Loc_Segment_Code from TSPL_LOCATION_MASTER WHERE Location_Code='" + Convert.ToString(obj.From_Location) + "'"
        Dim fromLocSegCode As String = connectSql.RunScalar(trans, Sql)
        Sql = "SELECT Loc_Segment_Code from TSPL_LOCATION_MASTER WHERE Location_Code='" + Convert.ToString(obj.To_Location) + "'"
        Dim toLocSegCode As String = connectSql.RunScalar(trans, Sql)
        Sql = "SELECT PA.Inv_Control_Account FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
       " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
        " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr.Item(0).Item_Code.ToString() + "'"
        strFromInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strFromInvAcc, obj.From_Location, trans)
        strToInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strFromInvAcc, obj.To_Location, trans)
        'strFromInvAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), fromLocSegCode)
        'strToInvAcc = connectSql.RunScalar(trans, Sql).Replace(connectSql.RunScalar(trans, Sql).ToString().Substring(connectSql.RunScalar(trans, Sql).ToString().Length - 3, 3), toLocSegCode)
        Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strFromInvAcc + "'"
        strFromInvAccDesc = connectSql.RunScalar(trans, Sql)
        If strFromInvAccDesc Is Nothing Then
            Throw New Exception("Inventory Control Account " + strFromInvAcc + " not found.")
        End If
        Sql = "SELECT  Description FROM TSPL_GL_ACCOUNTS WHERE  Account_Code = '" + strToInvAcc + "'"
        strToInvAccDesc = connectSql.RunScalar(trans, Sql)
        If strToInvAccDesc Is Nothing Then
            Throw New Exception("Inventory Control Account " + strToInvAcc + " not found.")
        End If

        Dim qry As String = ""
        'qry = " select Doc.Document_No,TaxM1.Tax_Liability_Account as Tax1_GLAC,TaxM2.Tax_Liability_Account as Tax2_GLAC," & _
        '         " TaxM3.Tax_Liability_Account as Tax3_GLAC,TaxM4.Tax_Liability_Account as Tax4_GLAC," & _
        '         " TaxM5.Tax_Liability_Account as Tax5_GLAC,TaxM6.Tax_Liability_Account as Tax6_GLAC, " & _
        '         " TaxM7.Tax_Liability_Account as Tax7_GLAC,TaxM8.Tax_Liability_Account as Tax8_GLAC, " & _
        '         " TaxM9.Tax_Liability_Account as Tax9_GLAC,TaxM10.Tax_Liability_Account as Tax10_GLAC, " & _
        '         " TaxM1.Tax_Net_Payable as Tax1_GLAC_Payable,TaxM2.Tax_Net_Payable as Tax2_GLAC_Payable, " & _
        '         " TaxM3.Tax_Net_Payable as Tax3_GLAC_Payable,TaxM4.Tax_Net_Payable as Tax4_GLAC_Payable, " & _
        '         " TaxM5.Tax_Net_Payable as Tax5_GLAC_Payable,TaxM6.Tax_Net_Payable as Tax6_GLAC_Payable, " & _
        '         " TaxM7.Tax_Net_Payable as Tax7_GLAC_Payable,TaxM8.Tax_Net_Payable as Tax8_GLAC_Payable, " & _
        '         " TaxM9.Tax_Net_Payable as Tax9_GLAC_Payable,TaxM10.Tax_Net_Payable as Tax10_GLAC_Payable from TSPL_TRANSFER_ORDER_HEAD doc " & _
        '         " left join TSPL_TAX_MASTER TaxM1 on Doc.TAX1=TaxM1.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM2 on Doc.TAX2=TaxM2.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM3 on Doc.TAX3=TaxM3.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM4 on Doc.TAX4=TaxM4.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM5 on Doc.TAX5=TaxM5.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM6 on Doc.TAX6=TaxM6.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM7 on Doc.TAX7=TaxM7.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM8 on Doc.TAX8=TaxM8.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM9 on Doc.TAX9=TaxM9.Tax_Code " & _
        '         " left join TSPL_TAX_MASTER TaxM10 on Doc.TAX10=TaxM10.Tax_Code where doc.Document_No='" & obj.Document_No & "'"
        'Dim dtTax As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        'If dtTax.Rows.Count = 0 Then
        '    Throw New Exception("Tax details of transfer document not found.")
        'End If
        'Dim TAX1_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax1_GLAC"))
        'Dim TAX2_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax2_GLAC"))
        'Dim TAX3_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax3_GLAC"))
        'Dim TAX4_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax4_GLAC"))
        'Dim TAX5_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax5_GLAC"))
        'Dim TAX6_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax6_GLAC"))
        'Dim TAX7_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax7_GLAC"))
        'Dim TAX8_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax8_GLAC"))
        'Dim TAX9_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax9_GLAC"))
        'Dim TAX10_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax10_GLAC"))

        'Dim TAX1_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX1_GLAC_Payable"))
        'Dim TAX2_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX2_GLAC_Payable"))
        'Dim TAX3_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX3_GLAC_Payable"))
        'Dim TAX4_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX4_GLAC_Payable"))
        'Dim TAX5_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX5_GLAC_Payable"))
        'Dim TAX6_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX6_GLAC_Payable"))
        'Dim TAX7_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX7_GLAC_Payable"))
        'Dim TAX8_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX8_GLAC_Payable"))
        'Dim TAX9_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX9_GLAC_Payable"))
        'Dim TAX10_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX10_GLAC_Payable"))

        ' '' validation for gl
        'If obj.TAX1_Amt > 0 Then
        '    If clsCommon.myLen(TAX1_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX1)
        '    End If
        '    If clsCommon.myLen(TAX1_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX1)
        '    End If
        'End If
        'If obj.TAX2_Amt > 0 Then
        '    If clsCommon.myLen(TAX2_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX2)
        '    End If
        '    If clsCommon.myLen(TAX2_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX2)
        '    End If
        'End If
        'If obj.TAX3_Amt > 0 Then
        '    If clsCommon.myLen(TAX3_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX3)
        '    End If
        '    If clsCommon.myLen(TAX3_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX3)
        '    End If
        'End If
        'If obj.TAX4_Amt > 0 Then
        '    If clsCommon.myLen(TAX4_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX4)
        '    End If
        '    If clsCommon.myLen(TAX4_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX4)
        '    End If
        'End If
        'If obj.TAX5_Amt > 0 Then
        '    If clsCommon.myLen(TAX5_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX5)
        '    End If
        '    If clsCommon.myLen(TAX5_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX5)
        '    End If
        'End If
        'If obj.TAX6_Amt > 0 Then
        '    If clsCommon.myLen(TAX6_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX6)
        '    End If
        '    If clsCommon.myLen(TAX6_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX6)
        '    End If
        'End If

        'If obj.TAX7_Amt > 0 Then
        '    If clsCommon.myLen(TAX7_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX7)
        '    End If
        '    If clsCommon.myLen(TAX7_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX7)
        '    End If
        'End If

        'If obj.TAX8_Amt > 0 Then
        '    If clsCommon.myLen(TAX8_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX8)
        '    End If
        '    If clsCommon.myLen(TAX8_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX8)
        '    End If
        'End If

        'If obj.TAX9_Amt > 0 Then
        '    If clsCommon.myLen(TAX9_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX9)
        '    End If
        '    If clsCommon.myLen(TAX9_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX9)
        '    End If
        'End If


        'If obj.TAX10_Amt > 0 Then
        '    If clsCommon.myLen(TAX10_GLAC) <= 0 Then
        '        Throw New Exception("Tax Liability Acount not found for" + obj.TAX10)
        '    End If
        '    If clsCommon.myLen(TAX10_GLAC_Payable) <= 0 Then
        '        Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX10)
        '    End If
        'End If
        'strFromInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strFromInvAcc, obj.From_Location, trans)
        Dim objSeg As Accountsegment = Accountsegment.Getaccountcodedesc(strFromInvAcc, trans)
        ' connectSql.RunSpTransaction(trans, "sp_TSPL_JOURNAL_DETAILS_INSERT", New SqlParameter("@Journal_No", Jrnl), New SqlParameter("@Voucher_No", StrVoucher), New SqlParameter("@Voucher_Date", strTranferDate), New SqlParameter("@Detail_Line_No", lineNo), New SqlParameter("@Account_code", strFromInvAcc), New SqlParameter("@Account_Desc", strFromInvAccDesc), New SqlParameter("@Amount", fromShipmentCogs * (-1)), New SqlParameter("@Description", clsCommon.myCstr(obj.description)), New SqlParameter("@Reference", ""), New SqlParameter("@Posting_Date", strTranferDate), New SqlParameter("@Account_Type", "I"), New SqlParameter("@Account_Group_Code", objSeg.Account_Group_Code), New SqlParameter("@Account_Seg_Code1", objSeg.Account_Seg_Code1), New SqlParameter("@Account_Seg_Desc1", objSeg.Account_Seg_Desc1), New SqlParameter("@Account_Seg_Code2", objSeg.Account_Seg_Code2), New SqlParameter("@Account_Seg_Desc2", objSeg.Account_Seg_Desc2), New SqlParameter("@Account_Seg_Code3", objSeg.Account_Seg_Code3), New SqlParameter("@Account_Seg_Desc3", objSeg.Account_Seg_Desc3), New SqlParameter("@Account_Seg_Code4", objSeg.Account_Seg_Code4), New SqlParameter("@Account_Seg_Desc4", objSeg.Account_Seg_Desc4), New SqlParameter("@Account_Seg_Code5", objSeg.Account_Seg_Code5), New SqlParameter("@Account_Seg_Desc5", objSeg.Account_Seg_Desc5), New SqlParameter("@Account_Seg_Code6", objSeg.Account_Seg_Code6), New SqlParameter("@Account_Seg_Desc6", objSeg.Account_Seg_Desc6), New SqlParameter("@Account_Seg_Code7", objSeg.Account_Seg_Code7), New SqlParameter("@Account_Seg_Desc7", objSeg.Account_Seg_Desc7), New SqlParameter("@Account_Seg_Code8", objSeg.Account_Seg_Code8), New SqlParameter("@Account_Seg_Desc8", objSeg.Account_Seg_Desc8), New SqlParameter("@Account_Seg_Code9", objSeg.Account_Seg_Code9), New SqlParameter("@Account_Seg_Desc9", objSeg.Account_Seg_Desc9), New SqlParameter("@Account_Seg_Code10", objSeg.Account_Seg_Code10), New SqlParameter("@Account_Seg_Desc10", objSeg.Account_Seg_Desc10))
        Dim arrlist As New ArrayList()
        Dim BankAcc() As String = {strFromInvAcc, obj.DOC_Total_Amt * (-1)}
        arrlist.Add(BankAcc)


        '  Dim taxAmt As Decimal
        Dim ttlTotalTaxAmt As Decimal = 0
        ' Dim strTaxCode As String
        Dim strNetPayAcc As String = ""
        Dim strNetPayAccDesc As String = ""




        strToInvAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strToInvAcc, obj.To_Location, trans)
        objSeg = Accountsegment.Getaccountcodedesc(strToInvAcc, trans)


        Dim DocAmtWithTax() As String = {strToInvAcc, obj.DOC_Total_Amt}
        arrlist.Add(DocAmtWithTax)

        Dim strTempAC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT TOP(1) Transfer_Clearing FROM TSPL_PURCHASE_ACCOUNTS", trans))
        If clsCommon.myLen(strTempAC) > 0 Then
            strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, fromLocSegCode, True, trans)
            Dim acc3() As String = {strTempAC, obj.DOC_Total_Amt}
            arrlist.Add(acc3)

            strTempAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strTempAC, toLocSegCode, True, trans)
            Dim acc4() As String = {strTempAC, -1 * obj.DOC_Total_Amt}
            arrlist.Add(acc4)
        End If
        JE_Excisable_Common(obj, arrlist, strLocationSegment, trans)
        If ((clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal)) Then
            clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)
        Else
            clsJournalMaster.FunGrnlEntryWithTrans(obj.From_Location, False, trans, obj.Document_Date, obj.Remarks, "MM-TF", "Stock Transfer", obj.Document_No, "", "I", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, arrlist, , obj.Remarks, obj.Remarks)

        End If
        Return True
    End Function

    Public Shared Function JE_Excisable_Common(ByVal obj As clsTransferDCC, ByVal arrlist As ArrayList, ByVal strLocationSegment As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        qry = " select Doc.Document_No,TaxM1.Tax_Liability_Account as Tax1_GLAC,TaxM2.Tax_Liability_Account as Tax2_GLAC," &
                 " TaxM3.Tax_Liability_Account as Tax3_GLAC,TaxM4.Tax_Liability_Account as Tax4_GLAC," &
                 " TaxM5.Tax_Liability_Account as Tax5_GLAC,TaxM6.Tax_Liability_Account as Tax6_GLAC, " &
                 " TaxM7.Tax_Liability_Account as Tax7_GLAC,TaxM8.Tax_Liability_Account as Tax8_GLAC, " &
                 " TaxM9.Tax_Liability_Account as Tax9_GLAC,TaxM10.Tax_Liability_Account as Tax10_GLAC, " &
                 " TaxM1.Tax_Net_Payable as Tax1_GLAC_Payable,TaxM2.Tax_Net_Payable as Tax2_GLAC_Payable, " &
                 " TaxM3.Tax_Net_Payable as Tax3_GLAC_Payable,TaxM4.Tax_Net_Payable as Tax4_GLAC_Payable, " &
                 " TaxM5.Tax_Net_Payable as Tax5_GLAC_Payable,TaxM6.Tax_Net_Payable as Tax6_GLAC_Payable, " &
                 " TaxM7.Tax_Net_Payable as Tax7_GLAC_Payable,TaxM8.Tax_Net_Payable as Tax8_GLAC_Payable, " &
                 " TaxM9.Tax_Net_Payable as Tax9_GLAC_Payable,TaxM10.Tax_Net_Payable as Tax10_GLAC_Payable, " &
                 " TaxM1.Tax_Recoverable_Account as Tax1_RecovAct,TaxM2.Tax_Recoverable_Account as Tax2_RecovAct, " &
                 " TaxM3.Tax_Recoverable_Account as Tax3_RecovAct,TaxM4.Tax_Recoverable_Account as Tax4_RecovAct, " &
                 " TaxM5.Tax_Recoverable_Account as Tax5_RecovAct,TaxM6.Tax_Recoverable_Account as Tax6_RecovAct, " &
                 " TaxM7.Tax_Recoverable_Account as Tax7_RecovAct,TaxM8.Tax_Recoverable_Account as Tax8_RecovAct, " &
                 " TaxM9.Tax_Recoverable_Account as Tax9_RecovAct,TaxM10.Tax_Recoverable_Account as Tax10_RecovAct from TSPL_TRANSFER_ORDER_HEAD doc " &
                 " left join TSPL_TAX_MASTER TaxM1 on Doc.TAX1=TaxM1.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM2 on Doc.TAX2=TaxM2.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM3 on Doc.TAX3=TaxM3.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM4 on Doc.TAX4=TaxM4.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM5 on Doc.TAX5=TaxM5.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM6 on Doc.TAX6=TaxM6.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM7 on Doc.TAX7=TaxM7.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM8 on Doc.TAX8=TaxM8.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM9 on Doc.TAX9=TaxM9.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM10 on Doc.TAX10=TaxM10.Tax_Code where doc.Document_No='" & obj.Document_No & "'"
        Dim dtTax As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dtTax.Rows.Count = 0 Then
            Throw New Exception("Tax details of transfer document not found.")
        End If
        Dim Tax1_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax1_RecovAct"))
        Dim Tax2_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax2_RecovAct"))
        Dim Tax3_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax3_RecovAct"))
        Dim Tax4_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax4_RecovAct"))
        Dim Tax5_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax5_RecovAct"))
        Dim Tax6_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax6_RecovAct"))
        Dim Tax7_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax7_RecovAct"))
        Dim Tax8_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax8_RecovAct"))
        Dim Tax9_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax9_RecovAct"))
        Dim Tax10_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax10_RecovAct"))

        Dim TAX1_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax1_GLAC"))
        Dim TAX2_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax2_GLAC"))
        Dim TAX3_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax3_GLAC"))
        Dim TAX4_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax4_GLAC"))
        Dim TAX5_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax5_GLAC"))
        Dim TAX6_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax6_GLAC"))
        Dim TAX7_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax7_GLAC"))
        Dim TAX8_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax8_GLAC"))
        Dim TAX9_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax9_GLAC"))
        Dim TAX10_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax10_GLAC"))

        Dim TAX1_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX1_GLAC_Payable"))
        Dim TAX2_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX2_GLAC_Payable"))
        Dim TAX3_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX3_GLAC_Payable"))
        Dim TAX4_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX4_GLAC_Payable"))
        Dim TAX5_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX5_GLAC_Payable"))
        Dim TAX6_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX6_GLAC_Payable"))
        Dim TAX7_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX7_GLAC_Payable"))
        Dim TAX8_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX8_GLAC_Payable"))
        Dim TAX9_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX9_GLAC_Payable"))
        Dim TAX10_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX10_GLAC_Payable"))

        '' validation for gl
        If clsERPFuncationality.GetGSTStatus(obj.Document_Date) = False Then
            If obj.TAX1_Amt > 0 Then
                If clsCommon.myLen(TAX1_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX1)
                End If
                If clsCommon.myLen(TAX1_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX1)
                End If
            End If
            If obj.TAX2_Amt > 0 Then
                If clsCommon.myLen(TAX2_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX2)
                End If
                If clsCommon.myLen(TAX2_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX2)
                End If
            End If
            If obj.TAX3_Amt > 0 Then
                If clsCommon.myLen(TAX3_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX3)
                End If
                If clsCommon.myLen(TAX3_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX3)
                End If
            End If
            If obj.TAX4_Amt > 0 Then
                If clsCommon.myLen(TAX4_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX4)
                End If
                If clsCommon.myLen(TAX4_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX4)
                End If
            End If
            If obj.TAX5_Amt > 0 Then
                If clsCommon.myLen(TAX5_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX5)
                End If
                If clsCommon.myLen(TAX5_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX5)
                End If
            End If
            If obj.TAX6_Amt > 0 Then
                If clsCommon.myLen(TAX6_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX6)
                End If
                If clsCommon.myLen(TAX6_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX6)
                End If
            End If

            If obj.TAX7_Amt > 0 Then
                If clsCommon.myLen(TAX7_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX7)
                End If
                If clsCommon.myLen(TAX7_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX7)
                End If
            End If

            If obj.TAX8_Amt > 0 Then
                If clsCommon.myLen(TAX8_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX8)
                End If
                If clsCommon.myLen(TAX8_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX8)
                End If
            End If

            If obj.TAX9_Amt > 0 Then
                If clsCommon.myLen(TAX9_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX9)
                End If
                If clsCommon.myLen(TAX9_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX9)
                End If
            End If


            If obj.TAX10_Amt > 0 Then
                If clsCommon.myLen(TAX10_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX10)
                End If
                If clsCommon.myLen(TAX10_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX10)
                End If
            End If
        Else
            If obj.TAX1_Amt > 0 Then
                If clsCommon.myLen(TAX1_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX1)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX1, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX1_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX1)
                        End If
                    Else
                        If clsCommon.myLen(Tax1_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX1)
                        End If
                    End If
                End If
            End If
            If obj.TAX2_Amt > 0 Then
                If clsCommon.myLen(TAX2_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX2)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX2, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX2_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX2)
                        End If
                    Else
                        If clsCommon.myLen(Tax2_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX2)
                        End If
                    End If
                End If
            End If
            If obj.TAX3_Amt > 0 Then
                If clsCommon.myLen(TAX3_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX3)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX3, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX3_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX3)
                        End If
                    Else
                        If clsCommon.myLen(Tax3_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX3)
                        End If
                    End If
                End If
            End If
            If obj.TAX4_Amt > 0 Then
                If clsCommon.myLen(TAX4_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX4)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX4, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX4_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX4)
                        End If
                    Else
                        If clsCommon.myLen(Tax4_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX4)
                        End If
                    End If
                End If
            End If
            If obj.TAX5_Amt > 0 Then
                If clsCommon.myLen(TAX5_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX5)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX5, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX5_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX5)
                        End If
                    Else
                        If clsCommon.myLen(Tax5_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX5)
                        End If
                    End If
                End If
            End If
            If obj.TAX6_Amt > 0 Then
                If clsCommon.myLen(TAX6_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX6)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX6, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX6_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX6)
                        End If
                    Else
                        If clsCommon.myLen(Tax6_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX6)
                        End If
                    End If
                End If
            End If

            If obj.TAX7_Amt > 0 Then
                If clsCommon.myLen(TAX7_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX7)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX7, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX7_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX7)
                        End If
                    Else
                        If clsCommon.myLen(Tax7_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX7)
                        End If
                    End If
                End If
            End If

            If obj.TAX8_Amt > 0 Then
                If clsCommon.myLen(TAX8_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX8)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX8, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX8_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX8)
                        End If
                    Else
                        If clsCommon.myLen(Tax8_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX8)
                        End If
                    End If
                End If
            End If

            If obj.TAX9_Amt > 0 Then
                If clsCommon.myLen(TAX9_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX9)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX9, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX9_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX9)
                        End If
                    Else
                        If clsCommon.myLen(Tax9_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX9)
                        End If
                    End If
                End If
            End If

            If obj.TAX10_Amt > 0 Then
                If clsCommon.myLen(TAX10_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX10)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX10, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX10_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX10)
                        End If
                    Else
                        If clsCommon.myLen(Tax10_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX10)
                        End If
                    End If
                End If
            End If
        End If
        '' taxes - from locaton
        '' richa agarwal net payable a/c should be debit and tax liablity a/c should be credit BM00000009115.
        If clsERPFuncationality.GetGSTStatus(obj.Document_Date) = False Then
            If obj.TAX1_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX1_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX1_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX2_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX2_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX2_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX3_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX3_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX3_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX4_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX4_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX4_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX5_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX5_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX5_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX6_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX6_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX6_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX6_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX6_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX7_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX7_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX7_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX7_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX7_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX8_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX8_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX8_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX8_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX8_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX9_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX9_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX9_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX9_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX9_Amt}
                arrlist.Add(accCR)
            End If


            If obj.TAX10_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX10_GLAC_Payable, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX10_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX10_GLAC, strLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX10_Amt}
                arrlist.Add(accCR)
            End If
        Else
            Dim Branch_Ac As String = ""
            If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                Dim FromLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location  from TSPL_TRANSFER_ORDER_HEAD where Document_No  ='" & obj.TransferOutNo & "'", trans))
                Dim FromLocSeg = FromLocation   'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & obj.From_Location & "'", trans))
                Dim ToLocSeg = obj.To_Location    'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from tspl_location_master where location_code='" & obj.To_Location & "'", trans))

                ''richa agarwal 24 Jan,2020
                Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)
                If PickTCAForStockTransferAndTankerDispatch = True Then
                    Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                    " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(0).Item_Code + "'", trans))
                    If clsCommon.myLen(Branch_Ac) <= 0 Then
                        Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(0).Item_Code + "  (" & obj.Arr(0).Item_Desc & ")")
                    End If
                    Branch_Ac = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac, strLocationSegment, True, trans)
                Else
                    Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "'", trans))
                    If clsCommon.myLen(Branch_Ac) <= 0 Then
                        Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)))
                    End If
                End If


            Else
                ''richa agarwal 24 Jan,2020
                Dim PickTCAForStockTransferAndTankerDispatch As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickTCAForStockTransferAndTankerDispatch, clsFixedParameterCode.PickTCAForStockTransferAndTankerDispatch, trans)) = 0, False, True)
                If PickTCAForStockTransferAndTankerDispatch = True Then
                    Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PA.transfer_clearing FROM TSPL_ITEM_MASTER AS IM INNER JOIN " &
                    " TSPL_PURCHASE_ACCOUNTS AS PA ON IM.Purchase_Class_Code = PA.Purchase_Class_Code INNER JOIN " &
                    " TSPL_GL_ACCOUNTS AS GLA ON PA.Inv_Control_Account = GLA.Account_Code WHERE IM.Item_Code='" + obj.Arr(0).Item_Code + "'", trans))
                    If clsCommon.myLen(Branch_Ac) <= 0 Then
                        Throw New Exception("Please Map Transfer Clearing Account For  for item " + obj.Arr(0).Item_Code + "  (" & obj.Arr(0).Item_Desc & ")")
                    End If
                    Branch_Ac = clsERPFuncationality.ChangeGLAccountLocationSegment(Branch_Ac, strLocationSegment, True, trans)
                Else
                    Branch_Ac = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.From_Location & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.To_Location_Main & "'", trans)) & "'", trans))
                    If clsCommon.myLen(Branch_Ac) <= 0 Then
                        Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.From_Location & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.To_Location_Main & "'", trans)))
                    End If
                End If


            End If

            If obj.TAX1_Amt > 0 Then
                '' debit
                Dim strTemp As String = ""
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX1, "MANDI") = CompairStringResult.Equal Then
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC_Payable, strLocationSegment, True, trans)
                    Else
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strLocationSegment, True, trans)
                    End If

                    Dim accDR() As String = {strTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    strTemp = Branch_Ac 'update by preeti gupta [UDL/10/09/18-000217]
                    Dim accCR() As String = {strTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)


                Else
                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strLocationSegment, True, trans)
                    Dim accDR() As String = {strTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)


                End If


                ''credit


            End If
            If obj.TAX2_Amt > 0 Then
                '' debit
                Dim strTemp As String = ""
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX2, "MANDI") = CompairStringResult.Equal Then
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC_Payable, strLocationSegment, True, trans)
                    Else
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax2_RecovAct, strLocationSegment, True, trans)
                    End If
                    Dim accDR() As String = {strTemp, 1 * obj.TAX2_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, -1 * obj.TAX2_Amt}
                    arrlist.Add(accCR)
                Else
                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC, strLocationSegment, True, trans)
                    Dim accDR() As String = {strTemp, -1 * obj.TAX2_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, 1 * obj.TAX2_Amt}
                    arrlist.Add(accCR)
                End If


            End If
            If obj.TAX3_Amt > 0 Then
                '' debit
                Dim strTemp As String = ""
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX3, "MANDI") = CompairStringResult.Equal Then
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC_Payable, strLocationSegment, True, trans)
                    Else
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax3_RecovAct, strLocationSegment, True, trans)
                    End If
                    Dim accDR() As String = {strTemp, 1 * obj.TAX3_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, -1 * obj.TAX3_Amt}
                    arrlist.Add(accCR)
                Else
                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC, strLocationSegment, True, trans)
                    Dim accDR() As String = {strTemp, -1 * obj.TAX3_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, obj.TAX3_Amt}
                    arrlist.Add(accCR)
                End If


            End If
            If obj.TAX4_Amt > 0 Then
                '' debit
                Dim strTemp As String = ""
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX4, "MANDI") = CompairStringResult.Equal Then
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC_Payable, strLocationSegment, True, trans)
                    Else
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax4_RecovAct, strLocationSegment, True, trans)
                    End If
                    Dim accDR() As String = {strTemp, 1 * obj.TAX4_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, -1 * obj.TAX4_Amt}
                    arrlist.Add(accCR)
                Else
                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC, strLocationSegment, True, trans)
                    Dim accDR() As String = {strTemp, -1 * obj.TAX4_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, obj.TAX4_Amt}
                    arrlist.Add(accCR)
                End If


            End If
            If obj.TAX5_Amt > 0 Then
                '' debit
                Dim strTemp As String = ""
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX5, "MANDI") = CompairStringResult.Equal Then
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC_Payable, strLocationSegment, True, trans)
                    Else
                        strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax5_RecovAct, strLocationSegment, True, trans)
                    End If
                    Dim accDR() As String = {strTemp, 1 * obj.TAX5_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, -1 * obj.TAX5_Amt}
                    arrlist.Add(accCR)
                Else
                    strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC, strLocationSegment, True, trans)
                    Dim accDR() As String = {strTemp, -1 * obj.TAX5_Amt}
                    arrlist.Add(accDR)

                    ''credit
                    strTemp = Branch_Ac
                    Dim accCR() As String = {strTemp, obj.TAX5_Amt}
                    arrlist.Add(accCR)
                End If

            End If
            If obj.TAX6_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX6_GLAC, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, -1 * obj.TAX6_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = Branch_Ac
                Dim accCR() As String = {strTemp, -1 * obj.TAX6_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX7_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX7_GLAC, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, -1 * obj.TAX7_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = Branch_Ac
                Dim accCR() As String = {strTemp, obj.TAX7_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX8_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX8_GLAC, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, -1 * obj.TAX8_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = Branch_Ac
                Dim accCR() As String = {strTemp, obj.TAX8_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX9_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX9_GLAC, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, -1 * obj.TAX9_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = Branch_Ac
                Dim accCR() As String = {strTemp, obj.TAX9_Amt}
                arrlist.Add(accCR)
            End If


            If obj.TAX10_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX10_GLAC, strLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, -1 * obj.TAX10_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = Branch_Ac
                Dim accCR() As String = {strTemp, obj.TAX10_Amt}
                arrlist.Add(accCR)
            End If
        End If
        Return True
    End Function
    Public Shared Function JE_Excisable_CommonReturnTransfer(ByVal obj As clsTransferDCC, ByVal arrlist As ArrayList, ByVal strFromLocationSegment As String, ByVal strToLocationSegment As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        qry = " select Doc.Document_No,TaxM1.Tax_Liability_Account as Tax1_GLAC,TaxM2.Tax_Liability_Account as Tax2_GLAC," &
                 " TaxM3.Tax_Liability_Account as Tax3_GLAC,TaxM4.Tax_Liability_Account as Tax4_GLAC," &
                 " TaxM5.Tax_Liability_Account as Tax5_GLAC,TaxM6.Tax_Liability_Account as Tax6_GLAC, " &
                 " TaxM7.Tax_Liability_Account as Tax7_GLAC,TaxM8.Tax_Liability_Account as Tax8_GLAC, " &
                 " TaxM9.Tax_Liability_Account as Tax9_GLAC,TaxM10.Tax_Liability_Account as Tax10_GLAC, " &
                 " TaxM1.Tax_Net_Payable as Tax1_GLAC_Payable,TaxM2.Tax_Net_Payable as Tax2_GLAC_Payable, " &
                 " TaxM3.Tax_Net_Payable as Tax3_GLAC_Payable,TaxM4.Tax_Net_Payable as Tax4_GLAC_Payable, " &
                 " TaxM5.Tax_Net_Payable as Tax5_GLAC_Payable,TaxM6.Tax_Net_Payable as Tax6_GLAC_Payable, " &
                 " TaxM7.Tax_Net_Payable as Tax7_GLAC_Payable,TaxM8.Tax_Net_Payable as Tax8_GLAC_Payable, " &
                 " TaxM9.Tax_Net_Payable as Tax9_GLAC_Payable,TaxM10.Tax_Net_Payable as Tax10_GLAC_Payable, " &
                 " TaxM1.Tax_Recoverable_Account as Tax1_RecovAct,TaxM2.Tax_Recoverable_Account as Tax2_RecovAct, " &
                 " TaxM3.Tax_Recoverable_Account as Tax3_RecovAct,TaxM4.Tax_Recoverable_Account as Tax4_RecovAct, " &
                 " TaxM5.Tax_Recoverable_Account as Tax5_RecovAct,TaxM6.Tax_Recoverable_Account as Tax6_RecovAct, " &
                 " TaxM7.Tax_Recoverable_Account as Tax7_RecovAct,TaxM8.Tax_Recoverable_Account as Tax8_RecovAct, " &
                 " TaxM9.Tax_Recoverable_Account as Tax9_RecovAct,TaxM10.Tax_Recoverable_Account as Tax10_RecovAct from TSPL_TRANSFER_ORDER_HEAD doc " &
                 " left join TSPL_TAX_MASTER TaxM1 on Doc.TAX1=TaxM1.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM2 on Doc.TAX2=TaxM2.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM3 on Doc.TAX3=TaxM3.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM4 on Doc.TAX4=TaxM4.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM5 on Doc.TAX5=TaxM5.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM6 on Doc.TAX6=TaxM6.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM7 on Doc.TAX7=TaxM7.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM8 on Doc.TAX8=TaxM8.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM9 on Doc.TAX9=TaxM9.Tax_Code " &
                 " left join TSPL_TAX_MASTER TaxM10 on Doc.TAX10=TaxM10.Tax_Code where doc.Document_No='" & obj.Document_No & "'"
        Dim dtTax As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dtTax.Rows.Count = 0 Then
            Throw New Exception("Tax details of transfer document not found.")
        End If
        Dim Tax1_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax1_RecovAct"))
        Dim Tax2_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax2_RecovAct"))
        Dim Tax3_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax3_RecovAct"))
        Dim Tax4_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax4_RecovAct"))
        Dim Tax5_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax5_RecovAct"))
        Dim Tax6_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax6_RecovAct"))
        Dim Tax7_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax7_RecovAct"))
        Dim Tax8_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax8_RecovAct"))
        Dim Tax9_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax9_RecovAct"))
        Dim Tax10_RecovAct As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax10_RecovAct"))

        Dim TAX1_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax1_GLAC"))
        Dim TAX2_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax2_GLAC"))
        Dim TAX3_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax3_GLAC"))
        Dim TAX4_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax4_GLAC"))
        Dim TAX5_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax5_GLAC"))
        Dim TAX6_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax6_GLAC"))
        Dim TAX7_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax7_GLAC"))
        Dim TAX8_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax8_GLAC"))
        Dim TAX9_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax9_GLAC"))
        Dim TAX10_GLAC As String = clsCommon.myCstr(dtTax.Rows(0).Item("Tax10_GLAC"))

        Dim TAX1_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX1_GLAC_Payable"))
        Dim TAX2_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX2_GLAC_Payable"))
        Dim TAX3_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX3_GLAC_Payable"))
        Dim TAX4_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX4_GLAC_Payable"))
        Dim TAX5_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX5_GLAC_Payable"))
        Dim TAX6_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX6_GLAC_Payable"))
        Dim TAX7_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX7_GLAC_Payable"))
        Dim TAX8_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX8_GLAC_Payable"))
        Dim TAX9_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX9_GLAC_Payable"))
        Dim TAX10_GLAC_Payable As String = clsCommon.myCstr(dtTax.Rows(0).Item("TAX10_GLAC_Payable"))

        '' validation for gl
        If clsERPFuncationality.GetGSTStatus(obj.Document_Date) = False Then
            If obj.TAX1_Amt > 0 Then
                If clsCommon.myLen(TAX1_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX1)
                End If
                If clsCommon.myLen(TAX1_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX1)
                End If
            End If
            If obj.TAX2_Amt > 0 Then
                If clsCommon.myLen(TAX2_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX2)
                End If
                If clsCommon.myLen(TAX2_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX2)
                End If
            End If
            If obj.TAX3_Amt > 0 Then
                If clsCommon.myLen(TAX3_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX3)
                End If
                If clsCommon.myLen(TAX3_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX3)
                End If
            End If
            If obj.TAX4_Amt > 0 Then
                If clsCommon.myLen(TAX4_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX4)
                End If
                If clsCommon.myLen(TAX4_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX4)
                End If
            End If
            If obj.TAX5_Amt > 0 Then
                If clsCommon.myLen(TAX5_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX5)
                End If
                If clsCommon.myLen(TAX5_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX5)
                End If
            End If
            If obj.TAX6_Amt > 0 Then
                If clsCommon.myLen(TAX6_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX6)
                End If
                If clsCommon.myLen(TAX6_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX6)
                End If
            End If

            If obj.TAX7_Amt > 0 Then
                If clsCommon.myLen(TAX7_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX7)
                End If
                If clsCommon.myLen(TAX7_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX7)
                End If
            End If

            If obj.TAX8_Amt > 0 Then
                If clsCommon.myLen(TAX8_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX8)
                End If
                If clsCommon.myLen(TAX8_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX8)
                End If
            End If

            If obj.TAX9_Amt > 0 Then
                If clsCommon.myLen(TAX9_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX9)
                End If
                If clsCommon.myLen(TAX9_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX9)
                End If
            End If


            If obj.TAX10_Amt > 0 Then
                If clsCommon.myLen(TAX10_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX10)
                End If
                If clsCommon.myLen(TAX10_GLAC_Payable) <= 0 Then
                    Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX10)
                End If
            End If
        Else
            If obj.TAX1_Amt > 0 Then
                If clsCommon.myLen(TAX1_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX1)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX1, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX1_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX1)
                        End If
                    Else
                        If clsCommon.myLen(Tax1_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX1)
                        End If
                    End If
                End If
            End If
            If obj.TAX2_Amt > 0 Then
                If clsCommon.myLen(TAX2_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX2)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX2, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX2_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX2)
                        End If
                    Else
                        If clsCommon.myLen(Tax2_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX2)
                        End If
                    End If
                End If
            End If
            If obj.TAX3_Amt > 0 Then
                If clsCommon.myLen(TAX3_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX3)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX3, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX3_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX3)
                        End If
                    Else
                        If clsCommon.myLen(Tax3_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX3)
                        End If
                    End If
                End If
            End If
            If obj.TAX4_Amt > 0 Then
                If clsCommon.myLen(TAX4_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX4)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX4, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX4_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX4)
                        End If
                    Else
                        If clsCommon.myLen(Tax4_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX4)
                        End If
                    End If
                End If
            End If
            If obj.TAX5_Amt > 0 Then
                If clsCommon.myLen(TAX5_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX5)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX5, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX5_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX5)
                        End If
                    Else
                        If clsCommon.myLen(Tax5_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX5)
                        End If
                    End If
                End If
            End If
            If obj.TAX6_Amt > 0 Then
                If clsCommon.myLen(TAX6_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX6)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX6, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX6_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX6)
                        End If
                    Else
                        If clsCommon.myLen(Tax6_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX6)
                        End If
                    End If
                End If
            End If

            If obj.TAX7_Amt > 0 Then
                If clsCommon.myLen(TAX7_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX7)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX7, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX7_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX7)
                        End If
                    Else
                        If clsCommon.myLen(Tax7_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX7)
                        End If
                    End If
                End If
            End If

            If obj.TAX8_Amt > 0 Then
                If clsCommon.myLen(TAX8_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX8)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX8, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX8_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX8)
                        End If
                    Else
                        If clsCommon.myLen(Tax8_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX8)
                        End If
                    End If
                End If
            End If

            If obj.TAX9_Amt > 0 Then
                If clsCommon.myLen(TAX9_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX9)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX9, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX9_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX9)
                        End If
                    Else
                        If clsCommon.myLen(Tax9_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX9)
                        End If
                    End If
                End If
            End If

            If obj.TAX10_Amt > 0 Then
                If clsCommon.myLen(TAX10_GLAC) <= 0 Then
                    Throw New Exception("Tax Liability Acount not found for" + obj.TAX10)
                End If
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(obj.TAX10, "MANDI") = CompairStringResult.Equal Then
                        If clsCommon.myLen(TAX10_GLAC_Payable) <= 0 Then
                            Throw New Exception("Tax Net Payable Acount not found for" + obj.TAX10)
                        End If
                    Else
                        If clsCommon.myLen(Tax10_RecovAct) <= 0 Then
                            Throw New Exception("Tax Recoverable Acount not found for" + obj.TAX10)
                        End If
                    End If
                End If
            End If
        End If
        '' taxes - from locaton
        '' richa agarwal net payable a/c should be debit and tax liablity a/c should be credit BM00000009115.
        If clsERPFuncationality.GetGSTStatus(obj.Document_Date) = False Then
            If obj.TAX1_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX1_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX1_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX2_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX2_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX2_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX2_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX3_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX3_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX3_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX3_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX4_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX4_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX4_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX4_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX5_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX5_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX5_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX5_Amt}
                arrlist.Add(accCR)
            End If
            If obj.TAX6_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX6_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX6_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX6_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX6_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX7_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX7_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX7_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX7_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX7_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX8_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX8_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX8_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX8_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX8_Amt}
                arrlist.Add(accCR)
            End If

            If obj.TAX9_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX9_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX9_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX9_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX9_Amt}
                arrlist.Add(accCR)
            End If


            If obj.TAX10_Amt > 0 Then
                '' debit
                Dim strTemp As String = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX10_GLAC_Payable, strToLocationSegment, True, trans)
                Dim accDR() As String = {strTemp, obj.TAX10_Amt}
                arrlist.Add(accDR)

                ''credit
                strTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX10_GLAC, strFromLocationSegment, True, trans)
                Dim accCR() As String = {strTemp, -1 * obj.TAX10_Amt}
                arrlist.Add(accCR)
            End If
        Else


            Dim strToTemp As String = ""
            Dim strFromTemp As String = ""
            'Dim Branch_Ac_FromLoc As String = ""
            'Dim Branch_Ac_To_Loc As String = ""

            'Branch_Ac_FromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)) & "'", trans))
            'If clsCommon.myLen(Branch_Ac_FromLoc) <= 0 Then
            '    Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & ToLocSeg & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & FromLocSeg & "'", trans)))
            'End If

            'Branch_Ac_To_Loc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Branch_account from tspl_branch_account_mapping where From_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.From_Location & "'", trans)) & "' and to_location='" & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.To_Location_Main & "'", trans)) & "'", trans))
            'If clsCommon.myLen(Branch_Ac_To_Loc) <= 0 Then
            '    Throw New Exception("Please Map Account For Branch Account Mapping For Location From  " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.From_Location & "'", trans)) & " To " & clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_segment_code from tspl_location_master where Location_Code='" & obj.To_Location_Main & "'", trans)))
            'End If

            If obj.TAX1_Amt > 0 Then

                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)
                    Dim accDR() As String = {strFromTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strToTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
            If obj.TAX2_Amt > 0 Then

                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
            If obj.TAX3_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
            If obj.TAX4_Amt > 0 Then

                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
            If obj.TAX5_Amt > 0 Then

                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
            If obj.TAX6_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If

            If obj.TAX7_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If

            If obj.TAX8_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If

            End If

            If obj.TAX9_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If


            If obj.TAX10_Amt > 0 Then
                If clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                    strToTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(Tax1_RecovAct, strToLocationSegment, True, trans)
                    strFromTemp = clsERPFuncationality.ChangeGLAccountLocationSegment(TAX1_GLAC, strFromLocationSegment, True, trans)

                    Dim accDR() As String = {strToTemp, 1 * obj.TAX1_Amt}
                    arrlist.Add(accDR)
                    Dim accCR() As String = {strFromTemp, -1 * obj.TAX1_Amt}
                    arrlist.Add(accCR)
                End If
            End If
        End If
        Return True
    End Function

    Private Shared Function FunReturnLocation(ByVal loc As String, Optional ByVal trans As SqlTransaction = Nothing) As String
        If trans Is Nothing Then
            Return connectSql.RunScalar("select Location_Type  from TSPL_LOCATION_MASTER where Location_Code = '" + loc + "'")
        Else
            Return connectSql.RunScalar(trans, "select Location_Type  from TSPL_LOCATION_MASTER where Location_Code = '" + loc + "'")
        End If
        Return ""
    End Function

    Private Shared Function HitInvenotryMovement(ByVal trans As SqlTransaction, ByVal obj As clsTransferDCC, ByVal strTranferDate As String) As Boolean
        'Dim SettPickAvgCost As Boolean = True ''clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.InDocMandatoryOnInternalTransfer, clsFixedParameterCode.InDocMandatoryOnInternalTransfer, trans)) = 1, True, False))
        Dim InDocMandatoryOnInternalTransfer As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.InDocMandatoryOnInternalTransfer, clsFixedParameterCode.InDocMandatoryOnInternalTransfer, trans)) = 1, True, False))
        Dim ArrInventoryMovementOut As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
        Dim ArrInventoryMovementIn As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
        Dim addcost As Decimal = 0.0
        Dim trnasfer As Decimal = 0
        Dim PunchingTime As DateTime = New DateTime(obj.Document_Date.Year, obj.Document_Date.Month, obj.Document_Date.Day, obj.Document_Date.Hour, obj.Document_Date.Minute, obj.Document_Date.Second)
        For Each objTr As clsTransferDCCDetail In obj.Arr
            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Transfer_Type, "R") = CompairStringResult.Equal Then
                trnasfer = objTr.Out_Qty
            ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Transfer_Type, "T") = CompairStringResult.Equal Then
                trnasfer = objTr.In_Qty
            End If
            Dim basicost As Double = objTr.Item_Cost
            Dim itemcost As Double
            'If SettPickAvgCost Then
            '    If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
            '        Dim objCost As MIlkComponentType = clsInventoryMovementNew.GetAvgCost("", objTr.Item_Code, obj.From_Location, objTr.Out_Qty, objTr.Unit_code, 0, 0, obj.Document_Date, obj.Document_Date, False, trans)
            '        itemcost = Math.Round(objCost.FAT_Cost, 2) + Math.Round(objCost.SNF_Cost, 2)
            '    ElseIf clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then

            '    End If
            'Else
            itemcost = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select DISTINCT Amount from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD  on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No" &
                         " where TSPL_TRANSFER_ORDER_HEAD.Document_No  ='" + obj.Document_No + "'  and TSPL_TRANSFER_ORDER_DETAIL.Item_Code ='" + objTr.Item_Code + "' ", trans))
            If itemcost = 0 Then
                itemcost = clsCommon.myCdbl(objTr.Amount)
            End If
            'End If

            Dim reccost As Double = 0.0
            Dim netcost As Double = clsCommon.myCdbl(itemcost)
            Dim strItemType As String = clsItemMaster.GetItemType(objTr.Item_Code, trans)
            Dim strItemTypeToSave As String = ""
            If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                strItemTypeToSave = "RM"
            ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                strItemTypeToSave = "OT"
            ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                strItemTypeToSave = "FT"
            Else
                strItemTypeToSave = strItemType
            End If

            Dim objInventoryMovemnt As New clsInventoryMovement()
            objInventoryMovemnt.InOut = "O"
            objInventoryMovemnt.Location_Code = obj.From_Location

            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                If clsCommon.CompairString(obj.InternalTransfer, 1) = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.JobWorkTransfer, 1) = CompairStringResult.Equal Then
                    objInventoryMovemnt.Other_Location_Code = obj.To_Location
                    objInventoryMovemnt.Other_Location_Desc = obj.To_LocationName
                Else
                    objInventoryMovemnt.Other_Location_Code = clsLocation.GetGITMainLocation(obj.To_Location, trans)
                    objInventoryMovemnt.Other_Location_Desc = clsLocation.GetName(objInventoryMovemnt.Other_Location_Code, trans)
                End If
            Else
                objInventoryMovemnt.Other_Location_Code = obj.To_Location
                objInventoryMovemnt.Other_Location_Desc = obj.To_LocationName
            End If



            objInventoryMovemnt.Item_Code = objTr.Item_Code
            objInventoryMovemnt.Item_Desc = objTr.Item_Desc
            objInventoryMovemnt.Qty = trnasfer
            objInventoryMovemnt.UOM = objTr.Unit_code
            objInventoryMovemnt.MRP = objTr.MRP
            objInventoryMovemnt.Add_Cost = addcost
            objInventoryMovemnt.Net_Cost = netcost
            objInventoryMovemnt.ItemType = strItemTypeToSave
            objInventoryMovemnt.Basic_Cost = basicost
            objInventoryMovemnt.Rec_Cost = reccost
            objInventoryMovemnt.Punching_Date = PunchingTime
            If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
            ElseIf InDocMandatoryOnInternalTransfer = True AndAlso clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
            Else
                ArrInventoryMovementOut.Add(objInventoryMovemnt)
            End If




            Dim objInventoryMovemnt1 As New clsInventoryMovement()
            objInventoryMovemnt1.InOut = "I"
            If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                objInventoryMovemnt1.Location_Code = obj.To_Location
            ElseIf InDocMandatoryOnInternalTransfer = True AndAlso clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                objInventoryMovemnt1.Location_Code = obj.To_Location
            Else
                objInventoryMovemnt1.Location_Code = obj.To_Location
            End If


            If clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
                objInventoryMovemnt1.Other_Location_Code = obj.From_Location
                objInventoryMovemnt1.Other_Location_Desc = obj.From_LocationName
            Else
                objInventoryMovemnt1.Other_Location_Code = GetTransferOutFromLocation(obj.TransferOutNo, trans)
                objInventoryMovemnt1.Other_Location_Desc = clsLocation.GetName(objInventoryMovemnt1.Other_Location_Code, trans)
            End If


            objInventoryMovemnt1.Item_Code = objTr.Item_Code
            objInventoryMovemnt1.Item_Desc = objTr.Item_Desc
            objInventoryMovemnt1.Qty = trnasfer
            objInventoryMovemnt1.UOM = objTr.Unit_code
            objInventoryMovemnt1.MRP = objTr.MRP
            objInventoryMovemnt1.Add_Cost = addcost
            objInventoryMovemnt1.Net_Cost = netcost
            objInventoryMovemnt1.ItemType = strItemTypeToSave
            objInventoryMovemnt1.Basic_Cost = basicost
            objInventoryMovemnt1.Rec_Cost = 0
            objInventoryMovemnt1.Punching_Date = PunchingTime

            If clsCommon.CompairString(obj.IsJobWorkType, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
            ElseIf InDocMandatoryOnInternalTransfer = True AndAlso clsCommon.CompairString(obj.InternalTransfer, "1") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Transfer_Type, "O") = CompairStringResult.Equal Then
            Else
                ArrInventoryMovementIn.Add(objInventoryMovemnt1)
            End If
        Next
        Dim TransType As String = ""
        If obj.JobWorkTransfer = 1 Then
            TransType = "JTransfer"
        ElseIf obj.InternalTransfer Then
            TransType = "ITransfer"
        Else
            TransType = "Transfer"
        End If
        clsInventoryMovement.SaveData(TransType, obj.Document_No, PunchingTime, clsCommon.GetPrintDate(PunchingTime, "dd/MM/yyyy"), ArrInventoryMovementOut, trans)
        clsInventoryMovement.SaveData(TransType, obj.Document_No, PunchingTime, clsCommon.GetPrintDate(PunchingTime, "dd/MM/yyyy"), ArrInventoryMovementIn, trans)

        Return True
    End Function

    Public Shared Function GetTransferOutFromLocation(ByVal strTransferOutNo As String, ByVal trans As SqlTransaction) As String
        Try
            Dim qry As String = "select inn.From_Location from TSPL_TRANSFER_ORDER_HEAD as inn where inn.Document_No ='" + strTransferOutNo + "'"
            Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
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
    Public Shared Function DeleteData(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        If (clsCommon.myLen(strCode) <= 0) Then
            Throw New Exception("Transfer Order No not found to Delete")
        End If
        Dim obj As clsTransferDCC = clsTransferDCC.GetData(strCode, NavigatorType.Current, trans)
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
            Try
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_TRANSFER_ORDER_HEAD", "DOCUMENT_NO", "TSPL_TRANSFER_ORDER_DETAIL", "DOCUMENT_NO", trans)
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strCode, "TSPL_TRANSFER_ORDER_HEAD", "DOCUMENT_NO", "TSPL_TRANSFER_ORDER_DETAIL", "DOCUMENT_NO", trans)

                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Transfer Order", "Transfer Order", obj.From_Location, obj.Document_Date, trans)
                If (obj.Status = 1) Then
                    Throw New Exception("Already Posted on :" + obj.Posting_Date)
                End If
                Dim DocType As String = ""
                If obj.JobWorkTransfer = 1 Then
                    DocType = "JTransfer"
                ElseIf obj.InternalTransfer Then
                    DocType = "ITransfer"
                Else
                    DocType = "Transfer"
                End If
                clsSerializeInvenotry.DeleteData(DocType, obj.Document_No, trans)
                clsBatchInventory.DeleteData(DocType, obj.Document_No, trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_TRANSFER_ORDER_HEAD", "Document_No", "TSPL_TRANSFER_ORDER_DETAIL", "Document_No", trans)
                Dim qry As String = "delete from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                clsCustomFieldValues.DeleteData(obj.Form_ID, strCode, trans)

                ''richa agarwal 
                If clsCommon.CompairString(obj.Transfer_Type, "I") = CompairStringResult.Equal Then
                    If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I' ", trans)), obj.TransferOutNo) = CompairStringResult.Equal) Or (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I' ", trans)), "") = CompairStringResult.Equal) Then
                        clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='N' where Document_No='" & obj.TransferOutNo & "' and Transfer_Type='O'", trans)
                    Else
                        clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='Y' where Document_No=(Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & obj.Document_No & "' and Transfer_Type='I') and Transfer_Type='O'", trans)
                        clsDBFuncationality.ExecuteNonQuery("Update TSPL_TRANSFER_ORDER_HEAD set Is_Status_IN='N' where Document_No='" & obj.TransferOutNo & "' and Transfer_Type='O'", trans)
                    End If
                End If
                If clsCommon.myLen(clsCommon.myCstr(obj.Requisition_Id)) > 0 Then
                    'Dim isProductionStoreReqDoc As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue(" select count (*)  from TSPL_PP_REQUISITION_HEAD where Requisition_Id = '" + obj.Requisition_Id + "' "))
                    Dim isProductionStoreReqDoc As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue(" select count (*)  from TSPL_PP_REQUISITION_HEAD where Requisition_Id = '" + obj.Requisition_Id + "' ", trans))
                    If isProductionStoreReqDoc = True Then
                        clsDBFuncationality.ExecuteNonQuery(" update TSPL_PP_REQUISITION_HEAD set close_yn = 'N' where Requisition_Id = '" + obj.Requisition_Id + "' ", trans)
                    End If
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
        Return True
    End Function

    Public Shared Function IsValidVendorForPO(ByVal ArrPONo As List(Of String), ByVal strVendorCode As String) As Boolean
        If ArrPONo IsNot Nothing AndAlso ArrPONo.Count > 0 Then
            Dim qry As String = "select Document_No,Vendor_Code,Vendor_Name from TSPL_TRANSFER_ORDER_HEAD where Document_No  in (" + clsCommon.GetMulcallString(ArrPONo) + ") and Vendor_Code not in ('" + strVendorCode + "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim msg As String = "Error. "
                For Each dr As DataRow In dt.Rows
                    msg += Environment.NewLine + "PO No:" + clsCommon.myCstr(dr("Document_No")) + " Is For Vendor Code: " + clsCommon.myCstr(dr("Vendor_Code")) + " Vendor Name:" + clsCommon.myCstr(dr("Vendor_Name"))
                Next
                Throw New Exception(msg)
            End If
        End If
        Return True
    End Function

    Public Shared Function IsValidTaxGroupForPO(ByVal ArrPONo As List(Of String), ByVal strTaxGroupCode As String) As Boolean
        If ArrPONo IsNot Nothing AndAlso ArrPONo.Count > 0 Then
            Dim qry As String = "select Document_No,Tax_Group from TSPL_TRANSFER_ORDER_HEAD where Document_No  in (" + clsCommon.GetMulcallString(ArrPONo) + ") and Tax_Group not in ('" + strTaxGroupCode + "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim msg As String = "Error. "
                For Each dr As DataRow In dt.Rows
                    msg += Environment.NewLine + "PO No:" + clsCommon.myCstr(dr("Document_No")) + " .Tax Group is: " + clsCommon.myCstr(dr("Tax_Group"))
                Next
                Throw New Exception(msg)
            End If
        End If
        Return True
    End Function

    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal blnUpdateLoadInwithLoadOut As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ReverseAndUnpost(strCode, blnUpdateLoadInwithLoadOut, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal blnUpdateLoadInwithLoadOut As Boolean, ByVal trans As SqlTransaction) As Boolean
        Return ReverseAndUnpost(strCode, blnUpdateLoadInwithLoadOut, trans, True)
    End Function

    Public Shared Function ReverseAndUnpost(ByVal strCode As String, ByVal blnUpdateLoadInwithLoadOut As Boolean, ByVal trans As SqlTransaction, ByVal CheckGatePass As Boolean) As Boolean
        Try
            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select Status,Transfer_Type,(case when Transfer_Type='O' then (select  inn.Document_No from TSPL_TRANSFER_ORDER_HEAD as inn where inn.TransferOutNo=TSPL_TRANSFER_ORDER_HEAD.Document_No ) else null end) as LoadOutNo,TransferOutNo,InternalTransfer,ProdRequestTransfer, JobWorkTransfer from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + strCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("Document no not found")
            End If

            If Not clsCommon.myCdbl(dt.Rows(0)("Status")) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If



            If blnUpdateLoadInwithLoadOut = False Then
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Transfer_Type")), "O") = CompairStringResult.Equal Then
                    If clsCommon.myLen(dt.Rows(0)("LoadOutNo")) > 0 Then
                        Throw New Exception("Loadin no -" + clsCommon.myCstr(dt.Rows(0)("Transfer_Type")) + " found")
                    End If
                Else
                    Qry = "delete from TSPL_PROVISION_ENTRY where Ref_Doc_No='" & clsCommon.myCstr(dt.Rows(0)("TransferOutNo")) & "' and Prog_Code='" + clsUserMgtCode.Transfer + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                End If
            Else
                If Not clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Transfer_Type")), "O") = CompairStringResult.Equal Then
                    Qry = "delete from TSPL_PROVISION_ENTRY where Ref_Doc_No='" & clsCommon.myCstr(dt.Rows(0)("TransferOutNo")) & "' and Prog_Code='" + clsUserMgtCode.Transfer + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                End If
            End If

            Dim InternalTransfer As Integer = clsCommon.myCdbl(dt.Rows(0)("InternalTransfer"))
            Dim JobWorkTransfer As Integer = clsCommon.myCdbl(dt.Rows(0)("JobWorkTransfer"))
            Dim transType As String = ""
            If JobWorkTransfer = 1 Then
                transType = "JTransfer"
            ElseIf InternalTransfer Then
                transType = "ITransfer"
            Else
                transType = "Transfer"
            End If
            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='MM-TF' and Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            ''richa agarwal 12 July,2020 if transfer Out no used into Dairy Gatepass then it will not get reversed.
            If CheckGatePass Then
                Qry = "select GPCode from TSPL_DAIRYSALE_GATEPASS_MASTER where isnull(AgainstTransferNo,'')='" + strCode + "' and isnull(AgainstTransferNo,'')<>''"
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
                If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                    Throw New Exception("Document no should not reversed because it has used in Gate Pass No " & clsCommon.myCstr(dt1.Rows(0)("GPCode")) & "")
                End If
            End If



            ''-----------------

            'Dim obj As clsTransferDCC = clsTransferDCC.GetData(strCode, NavigatorType.Current, trans)
            'PostTransferJournalEntryOnly(obj, VoucherNo, trans)
            clsBatchInventory.ReverseAndUnpost(transType, strCode, trans)

            ''richa agarwal 10 Feb,2017 delete items from serial item table against transfer no
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_SERIAL_ITEM", "Document_Code", trans)
            Qry = " delete from TSPL_SERIAL_ITEM where Document_Code ='" + strCode + "' and Document_Type ='" & transType & "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            ''------------------
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)
            Qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" + strCode + "' and Trans_Type='" & transType & "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_INVENTORY_MOVEMENT_New", "Source_Doc_No", trans)
            Qry = "delete from TSPL_INVENTORY_MOVEMENT_New where Source_Doc_No='" + strCode + "' and Trans_Type='" & transType & "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)



            If blnUpdateLoadInwithLoadOut = False Then
                Qry = "Update TSPL_TRANSFER_ORDER_HEAD set Status = 0 where Document_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            Else
                Qry = "Update TSPL_TRANSFER_ORDER_HEAD set Status = 0,GLVoucher_No='" & VoucherNo & "' where Document_No='" + strCode + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strCode, "TSPL_TRANSFER_ORDER_HEAD", "Document_No", trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetTransporterName(ByVal strTransporterCode As String, Optional ByVal trans As SqlTransaction = Nothing) As String
        Try
            Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" + strTransporterCode + "'", trans))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function GetAttachQry(ByVal strCode As String) As String
        Return GetAttachQry(strCode, Nothing)
    End Function

    Public Shared Function GetAttachQry(ByVal strCode As String, ByVal isCancelOrDelete As String) As String
        '' Add Delivery Challan query by Vinod
        Dim strCancelOrDelete As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
        Else
            strCancelOrDelete = Nothing
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
        End If
        Dim strQuery As String
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            strQuery = "(select TSPL_TRANSFER_ORDER_DETAIL.line_no,TSPL_TRANSFER_ORDER_HEAD.Price_Code, CASE WHEN TSPL_ITEM_MASTER.IsTaxable = 1 THEN 'T' ELSE 'NT' end as Taxable, tspl_item_master.hsn_code,tspl_location_master.GSTNo,TSPL_LOCATION_MASTER_1.GSTNo as To_Location_GSTNo,TSPL_TRANSFER_ORDER_HEAD.Remarks ,tspl_location_master.HOAdd1 ,TSPL_LOCATION_MASTER .HOAdd2,TSPL_TRANSFER_ORDER_HEAD.GR_No,TSPL_TRANSFER_ORDER_HEAD.document_type ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) as GR_Date,TSPL_TRANSFER_ORDER_HEAD.Transport_Id ,ISNULL(TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual,'') AS Transporter_Name_Manual,TSPL_TRANSFER_ORDER_HEAD.WayBill_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.WayBill_Date,103) as WayBill_Date,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,tspl_location_master_For_Location.City_Code as Location_City_Name, coalesce(cast(convert(decimal(18,0),(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * tspl_item_uom_detail.conversion_factor)/alt_convrsn.conversion_factor) as varchar)+' '+TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,'') as Alt_Unit_Code,TSPL_TRANSPORT_MASTER.Transporter_Name,(case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF=1 then 'Against F-Form Due' else '' end) as Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Document_No  as[STN_NO] , tspl_transfer_order_head.Document_Date as [Date_N_Time_issue],"
            strQuery += "  TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as Discount , TSPL_TRANSFER_ORDER_DETAIL .Document_No as ref_doc_no ,"
            strQuery += " TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_TRANSFER_ORDER_HEAD.To_Location ,TSPL_LOCATION_MASTER_1.Location_Desc as To_LocationName ,TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc ,"
            strQuery += " TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1,"
            strQuery += " TSPL_LOCATION_MASTER.Location_Code as From_Location_Code,  TSPL_LOCATION_MASTER.Location_Desc as From_Location_Dec,TSPL_LOCATION_MASTER.Add1 as From_Location_Add1 ,"
            strQuery += " TSPL_LOCATION_MASTER.Add2 as From_Location_Add2 , TSPL_LOCATION_MASTER.Add3 as From_Location_Add3,TSPL_LOCATION_MASTER.Add4  as From_Location_Add4,TSPL_LOCATION_MASTER.City_Code as From_Location_City_Code , "
            strQuery += " TSPL_STATE_MASTER.STATE_NAME  as From_Location_State,TSPL_LOCATION_MASTER.Pin_Code as From_Location_Pin_Code ,TSPL_LOCATION_MASTER.Country as From_Location_Country,  TSPL_LOCATION_MASTER.Telphone "
            strQuery += " as From_Location_Telphone,TSPL_LOCATION_MASTER.Email as From_Location_Email,TSPL_LOCATION_MASTER.TIN_No AS From_Location_tin_no, TSPL_LOCATION_MASTER.CST_No AS From_Location_cstNo,TSPL_ITEM_MASTER.Weight_UOM as UOM2,"
            strQuery += " TSPL_ITEM_MASTER.Weight_Value as Weight, TSPL_LOCATION_MASTER_1.Location_Code AS To_Location_Code,TSPL_LOCATION_MASTER.add1 as transpotor,TSPL_LOCATION_MASTER_1.Add1 AS To_Location_Add1, "
            strQuery += " TSPL_LOCATION_MASTER_1.Add2 as To_Location_Add2 ,TSPL_LOCATION_MASTER_1.Add3 as To_Location_Add3,TSPL_LOCATION_MASTER_1.Add4 as To_Location_Add4, TSPL_LOCATION_MASTER_1.Location_Desc as To_Location_Desc  ,"
            strQuery += " TSPL_LOCATION_MASTER_1.City_Code as To_Location_City_Code, StateMaster_ToLocation.State_Name as To_Location_State, TSPL_LOCATION_MASTER_1.Pin_Code as To_Location_Pin_Code,  TSPL_LOCATION_MASTER_1.Country as To_Location_Country,"
            strQuery += " TSPL_LOCATION_MASTER_1.Telphone as To_Location_Telphone, TSPL_LOCATION_MASTER_1.Email as To_Location_Email ,  TSPL_LOCATION_MASTER_1 .TIN_No as to_location_tin_no, TSPL_LOCATION_MASTER_1 .CST_No as to_location_cstno,TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX3,TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX6 "
            strQuery += ",TSPL_COMPANY_MASTER.Comp_Name AS CompName "
            strQuery += ",TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end +"
            strQuery += " case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end +"
            strQuery += " case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end +"
            strQuery += " case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end +"
            strQuery += " Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ' ,Phone'+TSPL_COMPANY_MASTER.Phone1 end + Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then '  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End +"
            strQuery += " case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end "
            strQuery += " as Company_Address, TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount,TSPL_COMPANY_MASTER .Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,tspl_company_master.GSTReg_No ,  tspl_company_master.Pan_No,tspl_company_master.Access_Officer as FSSAI_No,TSPL_COMPANY_MASTER.add1,TSPL_COMPANY_MASTER.add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Phone2,tspl_company_master.State,TSPL_TRANSFER_ORDER_HEAD.Requisition_Id, "
            strQuery += " InCrate.Conversion_Factor As [ConversionInCrate],case when coalesce(InCrate.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InCrate.Conversion_factor,1)) end as QtyInCrate,
                              case when coalesce(InPouch.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InPouch.Conversion_factor,1)) end as QtyInPouch,tabbatch.Batch_No,tabbatch.Batch_Qty "
            strQuery += " from " & tbl_TSPL_TRANSFER_ORDER_Detail & ""
            strQuery += " left outer join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No"
            strQuery += "  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location  "
            strQuery += " left outer join TSPL_LOCATION_MASTER AS TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.Location_Code =  TSPL_TRANSFER_ORDER_HEAD.To_Location "
            strQuery += " INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE "
            strQuery += " Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code "
            strQuery += " LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code "
            strQuery += " LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State "
            strQuery += " left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code "
            strQuery += " left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code "
            strQuery += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL 	  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Crate_type='Y') as InCrate on InCrate.Item_code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code "
            strQuery += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Packet_Type='Y') as InPouch on InPouch.Item_code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code "
            strQuery += " left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id "
            strQuery += " LEFT OUTER JOIN tspl_location_master  AS tspl_location_master_For_Location ON tspl_location_master_For_Location.GIT_Location =TSPL_TRANSFER_ORDER_HEAD.To_Location "
            strQuery += " left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State " &
            " LEFT OUTER JOIN TSPL_STATE_MASTER StateMaster_ToLocation ON StateMaster_ToLocation.State_Code=TSPL_LOCATION_MASTER_1.State
 left outer join ( select Document_Code,Parent_Line_No, STRING_AGG(Batch_No,CHAR(10)) as Batch_No , STRING_AGG(Qty,CHAR(10)) as Batch_Qty from(
SELECT Document_Code, Batch_No, Qty, Parent_Line_No FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='ITransfer' and In_Out_Type = 'O'
)x group by Document_Code,Parent_Line_No        
) TabBatch On TabBatch.Document_Code= TSPL_TRANSFER_ORDER_Detail.Document_No And TabBatch.Parent_Line_No = TSPL_TRANSFER_ORDER_Detail.Line_No"
            strQuery += " where 2=2   "
            strQuery += "  and  TSPL_TRANSFER_ORDER_HEAD. Document_No = '" + strCode + "' )"


            strQuery = "select '" & strCancelOrDelete & "' As IsCancelled,* from " & strQuery & " xx cross apply ( select top 1 TSPL_ITEM_PRICE_PLAN_DETAIL.Item_MRP,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price from TSPL_ITEM_PRICE_PLAN_DETAIL
                  left outer join TSPL_ITEM_PRICE_PLAN_HEADER on TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code 
                  where xx.Price_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code and xx.Item_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code
                  order by TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Date desc)
                  as price order by xx .line_no  "

        Else
            strQuery = "select '" & strCancelOrDelete & "' As IsCancelled,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1+TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3 as Comp_Address,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.GSTReg_No,TSPL_COMPANY_MASTER.Logo_Img,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual,frmLocation.Location_Desc as From_Location,ToLocation.Location_Desc as To_Location,TSPL_TRANSFER_ORDER_HEAD.Document_No,TSPL_TRANSFER_ORDER_HEAD.Document_Date,TSPL_TRANSFER_ORDER_HEAD.Requisition_Id,TSPL_REQUISITION_HEAD.Requisition_Date,TSPL_TRANSFER_ORDER_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,TSPL_TRANSFER_ORDER_DETAIL.Unit_code,case when isnull(TSPL_BATCH_ITEM.Batch_No,'')='' then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_BATCH_ITEM.Qty end as Qty,TSPL_BATCH_ITEM.Batch_No,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type
from " & tbl_TSPL_TRANSFER_ORDER_HEAD & "
left join " & tbl_TSPL_TRANSFER_ORDER_Detail & " on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
left join TSPL_ITEM_MASTER on TSPL_TRANSFER_ORDER_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code=TSPL_VEHICLE_MASTER.Vehicle_Id
left join TSPL_LOCATION_MASTER as frmLocation on TSPL_TRANSFER_ORDER_HEAD.From_Location=frmLocation.Location_Code
left join TSPL_LOCATION_MASTER as ToLocation on TSPL_TRANSFER_ORDER_HEAD.To_Location=ToLocation.Location_Code
left join TSPL_BATCH_ITEM on TSPL_TRANSFER_ORDER_DETAIL.Item_Code=TSPL_BATCH_ITEM.Item_Code and TSPL_TRANSFER_ORDER_DETAIL.Document_No=TSPL_BATCH_ITEM.Document_Code and TSPL_BATCH_ITEM.In_Out_Type='I'
left join TSPL_VENDOR_MASTER on TSPL_VEHICLE_MASTER.Transport_Id=TSPL_VENDOR_MASTER.Vendor_Code
left join TSPL_COMPANY_MASTER on TSPL_TRANSFER_ORDER_HEAD.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code

left join TSPL_REQUISITION_HEAD on TSPL_TRANSFER_ORDER_HEAD.Requisition_Id=TSPL_REQUISITION_HEAD.Requisition_Id

where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + strCode + "'"
        End If


        Return strQuery
    End Function

    Public Shared Function GetSTAMlkPrint(ByVal StrCode As String) As String
        Return GetSTAMlkPrint(StrCode, Nothing)
    End Function
    Public Shared Function GetSTAMlkPrint(ByVal StrCode As String, ByVal isCancelOrDelete As String) As String
        Dim strCancelOrDelete As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
        Else
            strCancelOrDelete = Nothing
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
        End If
        Dim Qry As String = "Select '" & strCancelOrDelete & "' As IsCancelled,CONCAT(Frm_StateGST,ToLoc_PANNO,GSTEntity,GSTBlank,GSTDegit) as LocGST,'" + objCommonVar.CurrentUserCode + "' as printedBy,* from (select TSPL_TRANSFER_ORDER_HEAD.Price_Code,TSPL_TRANSFER_ORDER_HEAD.Tax_Group,ISNULL(TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted,0) AS Is_Tax_Exempted ,TSPL_TRANSFER_ORDER_HEAD.Is_MandiTax, TSPL_TRANSFER_ORDER_HEAD.Electronic_Ref_No,
                       TSPL_LOCATION_MASTER.GSTNO as GSTIN_No ,TSPL_STATE_MASTER.GST_STATE_Code as From_Gst_StateCode,
                       TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_ITEM_MASTER.HSN_Code ,TSPL_LOCATION_MASTER.GSTNO as frm_GSTINNo,TSPL_LOCATION_MASTER_2.GSTEntity,case when len(TSPL_LOCATION_MASTER_2.GSTBlank   )>0 then TSPL_LOCATION_MASTER_2.GSTBlank  else 'Z' end    as GSTBlank,TSPL_LOCATION_MASTER_2.GSTDegit ,TSPL_STATE_MASTER.GST_STATE_Code as Frm_StateGST,
                       TSPL_TRANSFER_ORDER_HEAD.EWayBillNo ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as EWayBillDate ,convert(varchar(10),TSPL_COMPANY_MASTER.insurance_valid_date,103) as insurance_valid_date,
                       TSPL_COMPANY_MASTER.Pan_No as ToLoc_PANNO,  TSPL_TRANSFER_ORDER_HEAD.Is_taxable,TSPL_TRANSFER_ORDER_HEAD.For_Repair,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer, TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1 as dTAX1, TSPL_TRANSFER_ORDER_DETAIL.TAX2 as dTAX2, TSPL_TRANSFER_ORDER_DETAIL.TAX3 as  dTAX3, TSPL_TRANSFER_ORDER_DETAIL.TAX4 as  dTAX4, TSPL_TRANSFER_ORDER_DETAIL.TAX5 as  dTAX5, TSPL_TRANSFER_ORDER_DETAIL.TAX6 as  dTAX6, TSPL_TRANSFER_ORDER_DETAIL.TAX7 as  dTAX7, TSPL_TRANSFER_ORDER_DETAIL.TAX8 as dTAX8, TSPL_TRANSFER_ORDER_DETAIL.TAX9 as dTAX9, TSPL_TRANSFER_ORDER_DETAIL.TAX10 as  dTAX10, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate as dTAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate as dTAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate as dTAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate as dTAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate as dTAX5_Rate  ,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate as dTAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate as dTAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate as dTAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate as dTAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate as dTAX10_Rate,  dtax1.Type as tax1Type,dtax2.Type as tax2Type,dtax3.Type as tax3Type,dtax4.Type as tax4Type,dtax5.Type as tax5Type,dtax6.Type as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type, case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.crate_In else TSPL_TRANSFER_ORDER_HEAD.crate_out end as crate,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.jaali_in else TSPL_TRANSFER_ORDER_HEAD.jaali_Out end as jaali,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.box_in else TSPL_TRANSFER_ORDER_HEAD.box_Out end as box,TSPL_TRANSFER_ORDER_DETAIL.Line_No, '1' as CopyType,0 as Alter_UnitQty ,tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt,TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate,					 
					   TSPL_TRANSFER_ORDER_HEAD.Transfer_Type, TSPL_LOCATION_MASTER.TIN_No as Loc_Tin_No,TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end  + case when len(TSPL_STATE_MASTER.STATE_NAME   )>0 then ', '+ TSPL_STATE_MASTER.STATE_NAME  else ' ' end    as Location_Address_GP,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_STATE_MASTER.state_code as frm_State_code,tspl_location_master.HOAdd1 ,TSPL_LOCATION_MASTER .HOAdd2,TSPL_TRANSFER_ORDER_HEAD.GR_No,TSPL_TRANSFER_ORDER_HEAD.document_type ,case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date ,TSPL_TRANSFER_ORDER_HEAD.WayBill_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as WayBill_Date,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,
					 TSPL_TRANSFER_ORDER_HEAD.transport_id,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual ,(case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF=1 then 'Against F-Form Due' else '' end) as Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Document_No  as[STN_NO] ,
                       tspl_transfer_order_head.Document_Date as [Date_N_Time_issue],  TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as Discount , TSPL_TRANSFER_ORDER_DETAIL .Document_No as ref_doc_no , TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_LOCATION_MASTER.Location_Desc as From_LocationName,TSPL_LOCATION_MASTER_2.Location_Desc as To_LocationName,TSPL_TRANSFER_ORDER_HEAD.To_Location ,
                       TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1, TSPL_LOCATION_MASTER.Location_Code as From_Location_Code,  TSPL_LOCATION_MASTER.Location_Desc as From_Location_Dec,TSPL_LOCATION_MASTER.Add1 as From_Location_Add1 , TSPL_LOCATION_MASTER.Add2 as From_Location_Add2 , TSPL_LOCATION_MASTER.Add3 as From_Location_Add3,TSPL_LOCATION_MASTER.Add4  as From_Location_Add4,TSPL_LOCATION_MASTER.City_Code as From_Location_City_Code ,  TSPL_STATE_MASTER.STATE_NAME  as From_Location_State,TSPL_LOCATION_MASTER.Pin_Code as From_Location_Pin_Code ,TSPL_LOCATION_MASTER.Country as From_Location_Country,  TSPL_LOCATION_MASTER.Telphone  as From_Location_Telphone,TSPL_LOCATION_MASTER.Email as From_Location_Email,TSPL_LOCATION_MASTER.TIN_No AS From_Location_tin_no, TSPL_LOCATION_MASTER.CST_No AS From_Location_cstNo,TSPL_ITEM_MASTER.Weight_UOM as UOM2, TSPL_ITEM_MASTER.Weight_Value as Weight,((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinLTR.Conversion_Factor ) As TotalQty,
isnull(((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinCrate.Conversion_Factor),0 ) As CrateQty,
isnull(((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinPouch.Conversion_Factor ),0) As PouchQty,
CFinLTR.UOM_Code As TotalQtyUOM,TSPL_LOCATION_MASTER.add1 as transpotor,TSPL_LOCATION_MASTER_2.Add1 AS To_Location_Add1,TSPL_LOCATION_MASTER_2.Pin_Code As To_Location_PinCode, TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX3,TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX6 ,TSPL_COMPANY_MASTER.Comp_Name AS CompName ,
                        'PAN No - '+TSPL_COMPANY_MASTER.Pan_No +',GSTIN - '+TSPL_COMPANY_MASTER.GSTReg_No As CompPanGST,
                        TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount ,  tspl_company_master.Pan_No,tspl_company_master.Access_Officer as FSSAI_No,
                        tspl_company_master.State,TSPL_TRANSFER_ORDER_HEAD.Requisition_Id,TSPL_COMPANY_MASTER.GSTReg_No ,TSPL_COMPANY_MASTER.add1,TSPL_COMPANY_MASTER.add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Phone2
                        from " & tbl_TSPL_TRANSFER_ORDER_Detail & "   
                      inner join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "    on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No  
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
                        left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                        INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='LTR') As CFinLTR On CFinLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='Crate') As CFinCrate On CFinCrate.Item_Code=TSPL_ITEM_MASTER.Item_Code
						Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='Pouch') As CFinPouch On CFinPouch.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        
                        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code 
                       left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
                       left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
                        left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 
                        left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    
                        left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   
                        left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   
                        left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   
                        left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    
                        left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    
                        left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  
                        left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    
                        left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10    
                        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL .tax1 
                        left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2    
                        left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3   
                        left outer join TSPL_TAX_MASTER as dtax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4   
                        left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5   
                        left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6    
                        left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7    
                        left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8  
                        left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9    
                        left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10  
                        LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' 
                        where 2=2 and TSPL_TRANSFER_ORDER_HEAD. Document_No = '" + StrCode + "' and TSPL_ITEM_MASTER.IsTaxable =0)xxx
						cross apply ( select top 1 ((xxx.Quantity)*Item_Basic_Price) as TotalAmount,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_MRP,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price from TSPL_ITEM_PRICE_PLAN_DETAIL
                        left outer join TSPL_ITEM_PRICE_PLAN_HEADER on TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code 
                        where xxx.Price_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code and xxx.Item_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code and TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date <=   xxx.Date_N_Time_issue  
                        order by TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Date desc)  as price 
                        LEFT OUTER JOIN (Select 1 as COL1, 1 as COL2,  'ORIGINAL' as CopyType1 ) YYY ON YYY.COL1=XXX.CopyType ORDER BY Line_No,YYY.CopyType1  "
        Return Qry
    End Function

    Public Shared Function GetSTAProductQry(ByVal StrCode As String) As String
        Return GetSTAProductQry(StrCode, Nothing)
    End Function
    Public Shared Function GetSTAProductQry(ByVal StrCode As String, ByVal isCancelOrDelete As String) As String
        Dim strCancelOrDelete As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
        Else
            strCancelOrDelete = Nothing
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
        End If
        Dim Qry As String = "Select '" & strCancelOrDelete & "' As IsCancelled,CONCAT(Frm_StateGST,ToLoc_PANNO,GSTEntity,GSTBlank,GSTDegit) as LocGST,Batch_No,isnull(Batch_Qty,0)Batch_Qty,* from (select TSPL_TRANSFER_ORDER_HEAD.Price_Code,TSPL_TRANSFER_ORDER_HEAD.Tax_Group,ISNULL(TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted,0) AS Is_Tax_Exempted ,TSPL_TRANSFER_ORDER_HEAD.Is_MandiTax, TSPL_TRANSFER_ORDER_HEAD.Electronic_Ref_No,
                       TSPL_LOCATION_MASTER.GSTNO as GSTIN_No ,TSPL_STATE_MASTER.GST_STATE_Code as From_Gst_StateCode,
                       TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_ITEM_MASTER.HSN_Code ,TSPL_LOCATION_MASTER.GSTNO as frm_GSTINNo ,TSPL_LOCATION_MASTER_2.GSTEntity,case when len(TSPL_LOCATION_MASTER_2.GSTBlank   )>0 then TSPL_LOCATION_MASTER_2.GSTBlank  else 'Z' end    as GSTBlank,TSPL_LOCATION_MASTER_2.GSTDegit,TSPL_STATE_MASTER.GST_STATE_Code as Frm_StateGST,TSPL_TRANSFER_ORDER_HEAD.EWayBillNo ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as EWayBillDate ,convert(varchar(10),TSPL_COMPANY_MASTER.insurance_valid_date,103) as insurance_valid_date,TSPL_COMPANY_MASTER.Pan_No as ToLoc_PANNO, TSPL_TRANSFER_ORDER_HEAD.Is_taxable,TSPL_TRANSFER_ORDER_HEAD.For_Repair,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer, TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1 as dTAX1, TSPL_TRANSFER_ORDER_DETAIL.TAX2 as dTAX2, TSPL_TRANSFER_ORDER_DETAIL.TAX3 as  dTAX3, TSPL_TRANSFER_ORDER_DETAIL.TAX4 as  dTAX4, TSPL_TRANSFER_ORDER_DETAIL.TAX5 as  dTAX5, TSPL_TRANSFER_ORDER_DETAIL.TAX6 as  dTAX6, TSPL_TRANSFER_ORDER_DETAIL.TAX7 as  dTAX7, TSPL_TRANSFER_ORDER_DETAIL.TAX8 as dTAX8, TSPL_TRANSFER_ORDER_DETAIL.TAX9 as dTAX9, TSPL_TRANSFER_ORDER_DETAIL.TAX10 as  dTAX10, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate as dTAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate as dTAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate as dTAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate as dTAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate as dTAX5_Rate  ,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate as dTAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate as dTAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate as dTAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate as dTAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate as dTAX10_Rate,  dtax1.Type as tax1Type,dtax2.Type as tax2Type,dtax3.Type as tax3Type,dtax4.Type as tax4Type,dtax5.Type as tax5Type,dtax6.Type as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type, case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.crate_In else TSPL_TRANSFER_ORDER_HEAD.crate_out end as crate,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.jaali_in else TSPL_TRANSFER_ORDER_HEAD.jaali_Out end as jaali,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.box_in else TSPL_TRANSFER_ORDER_HEAD.box_Out end as box,TSPL_TRANSFER_ORDER_DETAIL.Line_No, '1' as CopyType,0 as Alter_UnitQty ,tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt,TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate, TSPL_TRANSFER_ORDER_HEAD.Transfer_Type, TSPL_LOCATION_MASTER.TIN_No as Loc_Tin_No,TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end  + case when len(TSPL_STATE_MASTER.STATE_NAME   )>0 then ', '+ TSPL_STATE_MASTER.STATE_NAME  else ' ' end    as Location_Address_GP,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_STATE_MASTER.state_code as frm_State_code,tspl_location_master.HOAdd1 ,TSPL_LOCATION_MASTER .HOAdd2,TSPL_TRANSFER_ORDER_HEAD.GR_No,TSPL_TRANSFER_ORDER_HEAD.document_type ,case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date ,TSPL_TRANSFER_ORDER_HEAD.WayBill_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as WayBill_Date,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,TSPL_TRANSFER_ORDER_HEAD.transport_id,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual ,(case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF=1 then 'Against F-Form Due' else '' end) as Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Document_No  as[STN_NO] ,
                       tspl_transfer_order_head.Document_Date as [Date_N_Time_issue],  TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as Discount , TSPL_TRANSFER_ORDER_DETAIL .Document_No as ref_doc_no , TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_LOCATION_MASTER.Location_Desc as From_LocationName,TSPL_LOCATION_MASTER_2.Location_Desc as To_LocationName,TSPL_TRANSFER_ORDER_HEAD.To_Location ,
                       TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1, TSPL_LOCATION_MASTER.Location_Code as From_Location_Code,  TSPL_LOCATION_MASTER.Location_Desc as From_Location_Dec,TSPL_LOCATION_MASTER.Add1 as From_Location_Add1 , TSPL_LOCATION_MASTER.Add2 as From_Location_Add2 , TSPL_LOCATION_MASTER.Add3 as From_Location_Add3,TSPL_LOCATION_MASTER.Add4  as From_Location_Add4,TSPL_LOCATION_MASTER.City_Code as From_Location_City_Code ,  TSPL_STATE_MASTER.STATE_NAME  as From_Location_State,TSPL_LOCATION_MASTER.Pin_Code as From_Location_Pin_Code ,TSPL_LOCATION_MASTER.Country as From_Location_Country,  TSPL_LOCATION_MASTER.Telphone  as From_Location_Telphone,TSPL_LOCATION_MASTER.Email as From_Location_Email,TSPL_LOCATION_MASTER.TIN_No AS From_Location_tin_no, TSPL_LOCATION_MASTER.CST_No AS From_Location_cstNo,TSPL_ITEM_MASTER.Weight_UOM as UOM2, TSPL_ITEM_MASTER.Weight_Value as Weight,((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinLTR.Conversion_Factor ) As TotalQty, CFinLTR.UOM_Code As TotalQtyUOM,TSPL_LOCATION_MASTER_2.Add1 AS To_Location_Add1,TSPL_LOCATION_MASTER_2.Pin_Code As To_Location_PinCode, TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX3,TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX6 ,TSPL_COMPANY_MASTER.Comp_Name AS CompName ,
                        'PAN No - '+TSPL_COMPANY_MASTER.Pan_No +',GSTIN - '+TSPL_COMPANY_MASTER.GSTReg_No As CompPanGST,
                        TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount ,  tspl_company_master.Pan_No,tspl_company_master.Access_Officer as FSSAI_No,tspl_company_master.State,TSPL_TRANSFER_ORDER_HEAD.Requisition_Id,TSPL_COMPANY_MASTER.GSTReg_No ,TSPL_COMPANY_MASTER.add1,TSPL_COMPANY_MASTER.add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Phone2,tabbatch.Batch_No,tabbatch.Batch_Qty
                        from " & tbl_TSPL_TRANSFER_ORDER_Detail & " 
                         inner join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No  
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
                        left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                        INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='LTR') As CFinLTR On CFinLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code 
                       left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
                       left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
                        left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 
                        left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    
                        left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   
                        left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   
                        left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   
                        left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    
                        left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    
                        left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  
                        left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    
                        left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10    
                        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL .tax1 
                        left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2    
                        left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3   
                        left outer join TSPL_TAX_MASTER as dtax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4   
                        left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5   
                        left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6    
                        left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7    
                        left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8  
                        left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9    
                        left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10  
                        LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T'
                        left outer join ( select Document_Code,Parent_Line_No, STRING_AGG(Batch_No,CHAR(10)) as Batch_No , STRING_AGG(Qty,CHAR(10)) as Batch_Qty from(
                        SELECT Document_Code, Batch_No, Qty, Parent_Line_No FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='ITransfer' and In_Out_Type = 'O'
                        )x group by Document_Code,Parent_Line_No ) TabBatch On TabBatch.Document_Code= TSPL_TRANSFER_ORDER_Detail.Document_No And TabBatch.Parent_Line_No = TSPL_TRANSFER_ORDER_Detail.Line_No
                        where 2=2 and TSPL_TRANSFER_ORDER_HEAD. Document_No = '" + StrCode + "' and TSPL_ITEM_MASTER.IsTaxable =1)xxx
                        cross apply ( select top 1 ((xxx.Quantity)*Item_Basic_Price) as TotalAmount,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1 as PTax1,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Rate as PTax1_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Amt*xxx.Quantity) as PTax1_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2 as PTax2,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Rate as PTax2_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Amt*xxx.Quantity) as PTax2_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3 as PTax3,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Rate as PTax3_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Amt * xxx.Quantity) as PTax3_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4 as PTax4,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Rate as PTax4_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Amt * xxx.Quantity) as PTax4_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5 as PTax5,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Rate as PTax5_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Amt * xxx.Quantity) as PTax5_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6 as PTax6,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Rate as PTax6_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Amt * xxx.Quantity) as PTax6_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7 as PTax7,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Rate as PTax7_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Amt * xxx.Quantity) as PTax7_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8 as PTax8,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Rate as PTax8_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Amt * xxx.Quantity) as PTax8_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9 as PTax9,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Rate as PTax9_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Amt * xxx.Quantity) as PTax9_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10 as PTax10,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Rate as PTax10_Rate,Convert(decimal(18,2),TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Amt * xxx.Quantity) as PTax10_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_MRP,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Selling_Price,((xxx.Quantity)*Item_Selling_Price) as valueInRs from TSPL_ITEM_PRICE_PLAN_DETAIL
                  left outer join TSPL_ITEM_PRICE_PLAN_HEADER on TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code 
                  where xxx.Price_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code and xxx.Item_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code and TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date <=   xxx.Date_N_Time_issue  
                  order by TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Date desc) as price 
          LEFT OUTER JOIN (
Select 1 as COL1, 1 as COL2,  'ORIGINAL' as CopyType1 ) YYY ON YYY.COL1=XXX.CopyType ORDER BY Line_No,YYY.COL2 "
        Return Qry
    End Function
End Class

Public Class clsTransferDCCDetail
#Region "Variables"
    Public Document_No As String = Nothing
    Public Line_No As Integer = 0
    Public Row_Type As String = Nothing
    Public Status As Integer = 0
    Public Item_Code As String = Nothing
    Public GatePassNo As String = Nothing
    Public Item_Desc As String = Nothing
    Public Out_Qty As Double = 0 '
    Public TransferOutNo As String = Nothing
    Public In_Qty As Double = 0 ''Not a tale field
    Public Shortage As Double = 0 '
    Public Breakage As Double = 0
    Public Leak As Double = 0
    Public Alt_Unit_Code As String = Nothing
    Public Unit_code As String = Nothing
    Public Price As Double = 0
    Public Item_Cost As Double = 0
    Public Location As String = Nothing
    Public LocationName As String = Nothing
    Public TAX1 As String = Nothing
    Public TAX1_Base_Amt As Double = 0
    Public TAX1_Rate As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Base_Amt As Double = 0
    Public TAX2_Rate As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Base_Amt As Double = 0
    Public TAX3_Rate As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Base_Amt As Double = 0
    Public TAX4_Rate As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Base_Amt As Double = 0
    Public TAX5_Rate As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Base_Amt As Double = 0
    Public TAX6_Rate As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Base_Amt As Double = 0
    Public TAX7_Rate As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Base_Amt As Double = 0
    Public TAX8_Rate As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Base_Amt As Double = 0
    Public TAX9_Rate As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Base_Amt As Double = 0
    Public TAX10_Rate As Double = 0
    Public TAX10_Amt As Double = 0
    Public Amount As Double = 0
    Public Disc_Per As Double = 0
    Public Disc_Amt As Double = 0
    Public Abatement_Per As Double = 0
    Public Abatement_Amt As Double = 0
    Public Amt_Less_Discount As Double = 0
    Public Total_Tax_Amt As Double = 0
    Public Item_Net_Amt As Double = 0
    Public Specification As String = Nothing
    Public Remarks As String = Nothing
    Public MRP As Double = 0
    Public FOCItem As Double = 0

    Public Item_Unit_Wt As Decimal = Nothing
    Public Item_Net_Wt As Decimal = Nothing
    Public Item_Net_MT_Wt As Decimal = Nothing
    Public arrSrItem As List(Of clsSerializeInvenotry) = Nothing
    Public arrBatchItem As List(Of clsBatchInventory) = Nothing
    Public Bin_No As String = 0
    Public ItemwiseTaxCode As String = ""
#End Region

    ''Note Very Important If any change mad in PO Head or PO Detail table allso update it's History table.
    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsTransferDCCDetail), ByVal trans As SqlTransaction, ByVal DocDate As DateTime, ByVal strFromLoc As String, ByVal strToLoc As String, ByVal Trans_Type As String, ByVal Transfer_Type As String, ByVal IsJobWorkType As Int16, ByVal IsInternalTransfer As Int16, ByVal isJobWorkTransfer As Int16) As Boolean
        clsSerializeInvenotry.DeleteData(Trans_Type, strDocNo, trans)
        clsBatchInventory.DeleteData(Trans_Type, strDocNo, trans)
        Dim qry As String = "delete from TSPL_Transfer_ORDER_DETAIL where Document_No='" + strDocNo + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each obj As clsTransferDCCDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                clsCommon.AddColumnsForChange(coll, "Row_Type", obj.Row_Type)
                clsCommon.AddColumnsForChange(coll, "ItemwiseTaxCode", obj.ItemwiseTaxCode)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Item_Desc", obj.Item_Desc)
                clsCommon.AddColumnsForChange(coll, "Out_Qty", obj.Out_Qty)
                clsCommon.AddColumnsForChange(coll, "In_Qty", obj.In_Qty)
                clsCommon.AddColumnsForChange(coll, "Item_Unit_Wt", obj.Item_Unit_Wt)
                clsCommon.AddColumnsForChange(coll, "Item_Net_Wt", obj.Item_Net_Wt)
                clsCommon.AddColumnsForChange(coll, "Item_Net_MT_Wt", obj.Item_Net_MT_Wt)
                clsCommon.AddColumnsForChange(coll, "GatePassNo", obj.GatePassNo, True)
                clsCommon.AddColumnsForChange(coll, "Breakage", obj.Breakage)
                clsCommon.AddColumnsForChange(coll, "Shortage", obj.Shortage)
                clsCommon.AddColumnsForChange(coll, "Leak", obj.Leak)
                clsCommon.AddColumnsForChange(coll, "TransferOutNo", obj.TransferOutNo, True)
                clsCommon.AddColumnsForChange(coll, "Alt_Unit_Code", obj.Alt_Unit_Code)
                clsCommon.AddColumnsForChange(coll, "Unit_code", obj.Unit_code)
                clsCommon.AddColumnsForChange(coll, "Location", obj.Location)
                clsCommon.AddColumnsForChange(coll, "Item_Cost", obj.Item_Cost)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Disc_Per", obj.Disc_Per)
                clsCommon.AddColumnsForChange(coll, "Disc_Amt", obj.Disc_Amt)
                clsCommon.AddColumnsForChange(coll, "Abatement_Per", obj.Abatement_Per)
                clsCommon.AddColumnsForChange(coll, "Abatement_Amt", obj.Abatement_Amt)
                clsCommon.AddColumnsForChange(coll, "Amt_Less_Discount", obj.Amt_Less_Discount)
                clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
                clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
                clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
                clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
                clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
                clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
                clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
                clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
                clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
                clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
                clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
                clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", obj.Total_Tax_Amt)
                clsCommon.AddColumnsForChange(coll, "Item_Net_Amt", obj.Item_Net_Amt)
                clsCommon.AddColumnsForChange(coll, "Specification", obj.Specification)
                clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
                clsCommon.AddColumnsForChange(coll, "MRP", obj.MRP)
                clsCommon.AddColumnsForChange(coll, "Bin_No", obj.Bin_No)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_TRANSFER_ORDER_DETAIL", OMInsertOrUpdate.Insert, "", trans)

                If (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyFEFO, clsFixedParameterCode.ApplyFEFO, trans)) = 1 And (IsInternalTransfer = 1 OrElse isJobWorkTransfer = 1)) Then
                    If clsCommon.CompairString(Transfer_Type, "O") = CompairStringResult.Equal Then
                        clsSerializeInvenotry.SaveData(Trans_Type, strDocNo, DocDate, "O", obj.Item_Code, strFromLoc, obj.Line_No, obj.arrSrItem, trans)
                        clsBatchInventory.SaveData(Trans_Type, strDocNo, DocDate, "O", obj.Item_Code, strFromLoc, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                    ElseIf clsCommon.CompairString(Transfer_Type, "I") = CompairStringResult.Equal Then
                        clsSerializeInvenotry.SaveData(Trans_Type, strDocNo, DocDate, "I", obj.Item_Code, strFromLoc, obj.Line_No, obj.arrSrItem, trans)
                        clsBatchInventory.SaveData(Trans_Type, strDocNo, DocDate, "I", obj.Item_Code, strFromLoc, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                    End If
                Else
                    clsSerializeInvenotry.SaveData(Trans_Type, strDocNo, DocDate, "O", obj.Item_Code, strFromLoc, obj.Line_No, obj.arrSrItem, trans)
                    clsSerializeInvenotry.SaveData(Trans_Type, strDocNo, DocDate, "I", obj.Item_Code, strToLoc, obj.Line_No, obj.arrSrItem, trans)

                    clsBatchInventory.SaveData(Trans_Type, strDocNo, DocDate, "O", obj.Item_Code, strFromLoc, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                    clsBatchInventory.SaveData(Trans_Type, strDocNo, DocDate, "I", obj.Item_Code, strToLoc, obj.Line_No, obj.MRP, obj.Unit_code, obj.arrBatchItem, trans)
                End If

            Next
        End If
        Return True
    End Function
    Public Shared Function GetMainLocationOfGIT(ByVal GIT_Loc As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select top 1 Location_Code from TSPL_LOCATION_MASTER where GIT_Location='" & GIT_Loc & "'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
    End Function

    Public Shared Function GetBalancePOQty(ByVal strPOCode As String, ByVal strICode As String, ByVal strCurrGRNNo As String, ByVal strUOM As String, ByVal dblMRP As Double, ByVal dblAssessable As Double) As Double
        Dim qry As String = "select SUM(qty * RI) as Balance from(  " &
            " select TSPL_TRANSFER_ORDER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_ORDER_DETAIL.TransferOrder_Qty as Qty,1 as RI from TSPL_TRANSFER_ORDER_DETAIL left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No where TSPL_TRANSFER_ORDER_DETAIL.Status=0 and TSPL_TRANSFER_ORDER_HEAD.Status=1 and TSPL_TRANSFER_ORDER_DETAIL.Document_No ='" + strPOCode + "' and TSPL_TRANSFER_ORDER_DETAIL.Item_Code='" + strICode + "' and  TSPL_TRANSFER_ORDER_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_TRANSFER_ORDER_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_TRANSFER_ORDER_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "'" &
            " union all " &
            " select  TSPL_SRN_DETAIL.Item_Code as ICode,((TSPL_SRN_DETAIL.SRN_Qty)+(TSPL_SRN_DETAIL.Leak_Qty)+(TSPL_SRN_DETAIL.Burst_Qty)+(TSPL_SRN_DETAIL.Short_Qty)) as Qty,-1 as RI from TSPL_SRN_DETAIL left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No where TSPL_SRN_DETAIL.PO_Id='" + strPOCode + "'   and TSPL_SRN_DETAIL.Item_Code='" + strICode + "' and  TSPL_SRN_DETAIL.Unit_code='" + strUOM + "' and isnull(TSPL_SRN_DETAIL.MRP,0)='" + clsCommon.myCstr(dblMRP) + "' and isnull(TSPL_SRN_DETAIL.Assessable,0)='" + clsCommon.myCstr(dblAssessable) + "' and TSPL_SRN_DETAIL.SRN_No not in ('" + strCurrGRNNo + "')  " &
            " )Final "
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
    End Function

    Public Shared Function CompletePO(ByVal strReqCode As String, ByVal strICode As String, ByVal LineNo As Integer) As Boolean
        Dim qry As String = "update TSPL_TRANSFER_ORDER_DETAIL set Status ='1' where Document_No='" + strReqCode + "' and Line_No='" + clsCommon.myCstr(LineNo) + "' and Item_Code='" + strICode + "'"
        Return clsDBFuncationality.ExecuteNonQuery(qry)
    End Function
End Class




