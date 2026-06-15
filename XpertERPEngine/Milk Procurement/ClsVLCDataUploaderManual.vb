'--------Created By Richa 05/09/2014 Against Ticket No BM00000003791
Imports System.Data.SqlClient
Imports common
Public Class ClsVLCDataUploaderManual
#Region "variables"
    Public Document_Code As String = Nothing
    Public Document_Date As Date
    Public VLC_Code As String = Nothing
    Public VLC_NAME As String = Nothing
    Public Route_Code As String = Nothing
    Public Route_Name As String = Nothing
    Public Shift As String = Nothing
    Public MCC_Code As String = ""
    Public MCC_Name As String = ""
    Public arrVLCDetail As List(Of ClsVLCDataUploaderManualdetail) = Nothing
    Public Dock_Collection_Milk_Type As String = ""
#End Region
    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = "Select TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code as Code,Convert(varchar,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) as Date, " &
            " TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code as [VLC Code],TSPL_VLC_MASTER_HEAD.VLC_Name as [VLC NAME] ,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [VLC Uploader Code],TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code as [Route Code], " &
            " TSPL_MCC_ROUTE_MASTER.Route_Name as [Route Name],TSPL_VLC_DATA_UPLOADER_MASTER.Shift,TSPL_VLC_MASTER_HEAD.MCC AS [MCC CODE],TSPL_MCC_MASTER.MCC_NAME AS [MCC NAME] from TSPL_VLC_DATA_UPLOADER_MASTER " &
            " Left Outer Join TSPL_MCC_ROUTE_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code= TSPL_MCC_ROUTE_MASTER.Route_Code " &
            " Left Outer Join TSPL_VLC_MASTER_HEAD on TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code " &
            " Left Outer Join TSPL_MCC_MASTER on TSPL_VLC_MASTER_HEAD.MCC=TSPL_MCC_MASTER.MCC_CODE "
        str = clsCommon.ShowSelectForm("VLCData", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function

    Public Shared Function GetQuery(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String, ByVal Mcc_Code_Uploder As String, ByVal tran As SqlTransaction) As String
        Return GetQuery(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code, Mcc_Code_Uploder, tran, 0, 0, 0)
    End Function
    Public Shared Function GetQuery(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String, ByVal Mcc_Code_Uploder As String, ByVal tran As SqlTransaction, ByVal Qty As Decimal, ByVal FAT As Decimal, ByVal SNF As Decimal) As String
        Dim qry As String
        Dim qryPDU As String = ""
        Dim qryManual As String = ""
        qryPDU = " select 'PDU' as Type,VLCData.Doc_No,convert(varchar,VLCData.Doc_Date,103) as Doc_Date,convert(varchar,VLCData.File_Date,103) as File_Date ,VLCData.MCC_Code,VLCData.VLC_CODE,VLC.VLC_Code_VLC_Uploader,VLCData.Route_No,VLCData.MP_CODE,MP.MP_Name ,VLCData.shift as [Shift]," &
              " VLCData.Milk_Type,VLCData.qty as VLC_Qty,VLCData.Uom_Code,VLCData.fat as VLC_Fat,VLCData.snf as VLC_SNF,VLCData.fat_KG,VLCData.snf_KG,VLCData.water as VLC_Water,VLCData.Rate,VLCData.Amount,mcc.Mcc_Code_VLC_Uploader as Mcc_Code_VLC_Uploader from TSPL_VLC_DATA_UPLOADER VLCData " &
              " left join TSPL_VLC_MASTER_HEAD VLC on VLCData.VLC_CODE=VLC.VLC_Code_VLC_Uploader and VLC.MCC=VLCData.MCC_Code left join TSPL_MP_MASTER MP ON VLCData.MP_CODE =MP.MP_Code left join TSPL_MCC_MASTER MCC on VLCData.MCC_Code=mcc.MCC_Code where 2=2  "
        If clsCommon.myLen(Doc_Date) > 0 Then
            qryPDU = qryPDU & " and VLCData.File_Date=convert(date,'" & clsCommon.GetPrintDate(Doc_Date, "dd/MMM/yyyy") & "',103)"
        End If
        If clsCommon.myLen(Shift) > 0 Then
            qryPDU = qryPDU & " and VLCData.shift='" & IIf(clsCommon.CompairString(Shift, "Morning") = CompairStringResult.Equal Or clsCommon.CompairString(Shift, "M") = CompairStringResult.Equal, "M", "E") & "'"
        End If
        If clsCommon.myLen(MCC_Code) > 0 Then
            qryPDU = qryPDU & " and VLCData.MCC_Code='" & MCC_Code & "'"
        End If
        If clsCommon.myLen(Mcc_Code_Uploder) > 0 Then
            qryPDU = qryPDU & " and mcc.Mcc_Code_VLC_Uploader='" & Mcc_Code_Uploder & "'"
        End If
        If clsCommon.myLen(VLC_Code) > 0 Then
            qryPDU = qryPDU & " and VLCData.VLC_Code='" & VLC_Code & "'"
        End If
        If clsCommon.myLen(VLC_Code_Uploader) > 0 Then
            qryPDU = qryPDU & " and VLC.VLC_Code_VLC_Uploader='" & VLC_Code_Uploader & "'"
        End If
        If clsCommon.myLen(Route_Code) > 0 Then
            qryPDU = qryPDU & " and VLCData.Route_No='" & Route_Code & "'"
        End If
        If clsCommon.myLen(MP_Code) > 0 Then
            qryPDU = qryPDU & " and VLCData.MP_CODE='" & MP_Code & "'"
        End If
        If Qty > 0 Then
            qryPDU = qryPDU & " and VLCData.qty='" & clsCommon.myCstr(Qty) & "'"
        End If
        If FAT > 0 Then
            qryPDU = qryPDU & " and VLCData.fat='" & clsCommon.myCstr(FAT) & "'"
        End If
        If SNF > 0 Then
            qryPDU = qryPDU & " and VLCData.snf='" & clsCommon.myCstr(SNF) & "'"
        End If
        qryManual = " select MFinal.Type,MFinal.Document_Code,MFinal.Doc_Date as Document_Date,MFinal.Document_Date,MFinal.MCC,MFinal.VLC_Code,MFinal.VLC_Code_VLC_Uploader,MFinal.Route_Code,MFinal.Farmer_Code," &
                    " MFinal.MP_Name,MFinal.Shift,MFinal.Milk_Type,MFinal.VLC_Qty,MFinal.Unit_Code,MFinal.VLC_Fat,MFinal.VLC_SNF,MFinal.VLC_FAT_KG*Conv.CF,MFinal.VLC_SNF_KG*Conv.CF,MFinal.VLC_Water,MFinal.Rate,MFinal.Amount,MFinal.Mcc_Code_VLC_Uploader from " &
                    " ( select 'Manual' as Type,VLCM.Document_Code,convert(varchar,VLCM.Document_Date,103) as Doc_Date,convert(varchar,VLCM.Document_Date,103) as Document_Date,VLC.MCC,VLCM.VLC_Code,VLC.VLC_Code_VLC_Uploader,VLCM.Route_Code,VLCD.Farmer_Code,MP.MP_Name ,VLCM.Shift, " &
                    " '' as Milk_Type,VLCD.Qty as VLC_Qty,VLCD.Unit_Code,VLCD.FatPer as VLC_Fat,VLCD.SNFPer as VLC_SNF,(VLCD.Qty*VLCD.FatPer/100) as VLC_FAT_KG, " &
                    " (VLCD.Qty*VLCD.SNFPer/100) as VLC_SNF_KG,null as VLC_Water,VLCD.Rate,VLCD.Amount,'' as Mcc_Code_VLC_Uploader from TSPL_VLC_DATA_UPLOADER_MASTER VLCM " &
                    " inner join TSPL_VLC_DATA_UPLOADER_DETAIL VLCD on VLCM.Document_Code=VLCD.Document_Code " &
                    " left join TSPL_VLC_MASTER_HEAD VLC on VLCM.VLC_CODE=VLC.VLC_Code " &
                    " left join TSPL_MP_MASTER MP ON VLCD.Farmer_Code=MP.MP_Code where 2=2 "

        If clsCommon.myLen(Doc_Date) > 0 Then
            qryManual = qryManual & " and cast(VLCM.Document_Date as date)='" & clsCommon.GetPrintDate(Doc_Date, "dd/MMM/yyyy") & "'"
        End If
        If clsCommon.myLen(Shift) > 0 Then
            qryManual = qryManual & " and VLCM.shift='" & IIf(clsCommon.CompairString(Shift, "Morning") = CompairStringResult.Equal Or clsCommon.CompairString(Shift, "M") = CompairStringResult.Equal, "MORNING", "EVENING") & "'"
        End If
        If clsCommon.myLen(MCC_Code) > 0 Then
            qryManual = qryManual & " and VLC.MCC='" & MCC_Code & "'"
        End If
        If clsCommon.myLen(VLC_Code) > 0 Then
            qryManual = qryManual & " and VLCM.VLC_Code='" & VLC_Code & "'"
        End If
        If clsCommon.myLen(VLC_Code_Uploader) > 0 Then
            qryManual = qryManual & " and VLC.VLC_Code_VLC_Uploader='" & VLC_Code_Uploader & "'"
        End If
        If clsCommon.myLen(Route_Code) > 0 Then
            qryManual = qryManual & " and VLCM.Route_Code='" & Route_Code & "'"
        End If
        If clsCommon.myLen(MP_Code) > 0 Then
            qryManual = qryManual & " and MP.MP_Code_VLC_Uploader='" & MP_Code & "'"
        End If

        If Qty > 0 Then
            qryManual = qryManual & " and VLCD.Qty='" & clsCommon.myCstr(Qty) & "'"
        End If
        If FAT > 0 Then
            qryManual = qryManual & " and VLCD.FatPer='" & clsCommon.myCstr(FAT) & "'"
        End If
        If SNF > 0 Then
            qryManual = qryManual & " and VLCD.SNFPer='" & clsCommon.myCstr(SNF) & "'"
        End If

        qryManual = qryManual & ") as MFinal "
        ''richa agarwal TEC/28/03/19-000462 add item structure on setting based
        Dim ItemStructureMandatoryOnWeightConversion As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ItemStructureMandatoryOnWeightConversion, clsFixedParameterCode.ItemStructureMandatoryOnWeightConversion, tran)) = 1, True, False))
        If ItemStructureMandatoryOnWeightConversion = True Then
            qryManual = qryManual & " left join (Select yyy.FromUOM,yyy.TOUOM,max(CF) as CF,Structure_code From (  Select Container_UOM as FromUOM, Contained_UOM as TOUOM, Container_Qty*Contained_Qty as CF,Structure_code " &
                     " from TSPL_WEIGHT_CONVERSION UNION All Select Contained_UOM as FromUOM, Container_UOM as TOUOM, Container_Qty/Contained_Qty as CF,Structure_code from TSPL_WEIGHT_CONVERSION " &
                     " UNION All   Select Contained_UOM as FromUOM, Contained_UOM as TOUOM, 1 as CF,Structure_code from TSPL_WEIGHT_CONVERSION " &
                     " UNION All Select Container_UOM as FromUOM, Container_UOM as TOUOM, 1 as CF,Structure_code from TSPL_WEIGHT_CONVERSION  )as yyy group by yyy.FromUOM,yyy.TOUOM,yyy.Structure_code ) as Conv on MFinal.Unit_Code=Conv.FromUOM and Conv.TOUOM='KG' where Conv.Structure_code =(select Structure_Code  from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, Nothing)) & "')"
        Else
            qryManual = qryManual & " left join (Select yyy.FromUOM,yyy.TOUOM,max(CF) as CF From (  Select Container_UOM as FromUOM, Contained_UOM as TOUOM, Container_Qty*Contained_Qty as CF " &
                     " from TSPL_WEIGHT_CONVERSION UNION All Select Contained_UOM as FromUOM, Container_UOM as TOUOM, Container_Qty/Contained_Qty as CF from TSPL_WEIGHT_CONVERSION " &
                     " UNION All   Select Contained_UOM as FromUOM, Contained_UOM as TOUOM, 1 as CF from TSPL_WEIGHT_CONVERSION " &
                     " UNION All Select Container_UOM as FromUOM, Container_UOM as TOUOM, 1 as CF from TSPL_WEIGHT_CONVERSION  )as yyy group by yyy.FromUOM,yyy.TOUOM ) as Conv on MFinal.Unit_Code=Conv.FromUOM and Conv.TOUOM='KG'"
        End If

        qry = "select * from (" & qryPDU & "Union All " & qryManual & ") as Final where 2=2 "
        If clsCommon.myLen(Type) > 0 Then
            qry = qry & " and Type='" & Type & "'"
        End If
        Return qry
    End Function
    Public Shared Function GetQuery1(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String) As String
        Dim qry As String
        qry = GetQuery(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code, "", Nothing)
        qry = "select sum(VLC_Qty) as VLC_Qty,sum(Fat_KG) as Fat_KG,sum(SNF_KG) as SNF_KG,(100*sum(Fat_KG)/sum(VLC_Qty)) as Fat_Per,(100*sum(SNF_KG)/sum(VLC_Qty)) as SNF_Per from (" & qry & ") as Avg"
        Return qry
    End Function

    Public Shared Function GetQueryDT(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String, ByVal Mcc_Code_uploader As String) As DataTable
        Return GetQueryDT(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code, Mcc_Code_uploader, Nothing)
    End Function

    Public Shared Function GetQueryDT(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String, ByVal Mcc_Code_uploader As String, ByVal tran As SqlTransaction) As DataTable
        Return GetQueryDT(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code, Mcc_Code_uploader, tran, 0, 0, 0)
    End Function

    Public Shared Function GetQueryDT(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String, ByVal Mcc_Code_uploader As String, ByVal tran As SqlTransaction, ByVal Qty As Decimal, ByVal FAT As Decimal, ByVal SNF As Decimal) As DataTable
        Dim dt As DataTable
        Dim qry As String = GetQuery(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code, Mcc_Code_uploader, tran, Qty, FAT, SNF)
        dt = clsDBFuncationality.GetDataTable(qry, tran)
        Return dt
    End Function
    Public Shared Function GetQueryDT1(ByVal Type As String, ByVal Doc_Date As String, ByVal Shift As String, ByVal MCC_Code As String, ByVal VLC_Code As String, ByVal VLC_Code_Uploader As String, ByVal Route_Code As String, ByVal MP_Code As String) As DataTable
        Dim dt As DataTable
        Dim qry As String = GetQuery1(Type, Doc_Date, Shift, MCC_Code, VLC_Code, VLC_Code_Uploader, Route_Code, MP_Code)
        dt = clsDBFuncationality.GetDataTable(qry)
        Return dt
    End Function

    Public Shared Function SaveData(ByVal obj As ClsVLCDataUploaderManual, ByVal isNewEntry As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(obj, isNewEntry, trans)
            trans.Commit()
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function
    Public Shared Function SaveData(ByVal obj As ClsVLCDataUploaderManual, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""


        Try
            '' apply validaton for Farmer payment: If Farmer Payment is On and Bill is generated for the VSP for which milk data is to be uploaded
            '' get MPPayment setting
            Dim mcc_Code As String = clsDBFuncationality.getSingleValue("select location_Code from tspl_mcc_route_master inner join tspl_location_master on location_Code=mcc_Code where route_Code='" & obj.Route_Code & "'", trans)

            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.FrmVLCDataUploaderManual, mcc_Code, obj.Document_Date, trans)

            Dim IsMP As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.isFarmerPaymentCycle, clsFixedParameterCode.isFarmerPaymentCycle, trans))
            If IsMP = 1 Then
                Dim arrVLC As New ArrayList
                arrVLC.Add(obj.VLC_Code)
                Dim arr As ArrayList = clsVlcDataUploader.CheckBillStatusByVLC(arrVLC, obj.Document_Date, trans)
                If Not arr Is Nothing AndAlso arr.Count > 0 Then
                    Dim VLC As String = clsCommon.GetMulcallStringWithComma(arr)
                    Throw New Exception("Bill is generated for Some VLCs:" & Environment.NewLine & VLC)
                End If
            End If

            'Dim mcc_Code As String = clsDBFuncationality.getSingleValue("select location_Code from tspl_mcc_route_master inner join tspl_location_master on location_Code=mcc_Code where route_Code='" & obj.Route_Code & "'", trans)
            qry = "delete from TSPL_VLC_DATA_UPLOADER_DETAIL where Document_Code='" & obj.Document_Code & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            If isNewEntry Then
                qry = "select Document_Code from TSPL_VLC_DATA_UPLOADER_MASTER where CONVERT(date, TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)='" + clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.[Shift] = '" + obj.Shift + "'  and VLC_code='" + obj.VLC_Code + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Throw New Exception("Document No [" + clsCommon.myCstr(dt.Rows(0)("Document_Code")) + "] is already generated for VLC [" + obj.VLC_Code + "].Please Add in this document")
                End If
                obj.Document_Code = clsERPFuncationality.GetNextCode(trans, obj.Document_Date, clsDocType.VLCDataUploaderManual, "", mcc_Code)
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "VLC_Code", obj.VLC_Code)
            clsCommon.AddColumnsForChange(coll, "Route_Code", obj.Route_Code)
            clsCommon.AddColumnsForChange(coll, "Shift", obj.Shift)
            clsCommon.AddColumnsForChange(coll, "Dock_Collection_Milk_Type", obj.Dock_Collection_Milk_Type)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
            '' update Sync Satatus
            clsCommon.AddColumnsForChange(coll, "SYNC_STATUS", 0)


            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
                clsCommon.AddColumnsForChange(coll, "Document_Code", obj.Document_Code)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VLC_DATA_UPLOADER_MASTER", OMInsertOrUpdate.Insert, "", trans)
                CreateSMSContentMP(obj, trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VLC_DATA_UPLOADER_MASTER", OMInsertOrUpdate.Update, "TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code='" + obj.Document_Code + "'", trans)
            End If
            ClsVLCDataUploaderManualdetail.saveData(obj.Document_Date, mcc_Code, obj.VLC_Code, obj.arrVLCDetail, obj.Document_Code, trans)

        Catch err As Exception

            Throw New Exception(err.Message)
        End Try
        Return True
    End Function

    Shared Sub CreateSMSContentMP(ByVal obj As ClsVLCDataUploaderManual, ByVal trans As SqlTransaction)
        'SMSCode Start
        Dim strSMSContent As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT SMS_Text from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmVlcdataUploadar + "'", trans))
        If clsCommon.myLen(strSMSContent) > 0 Then
            For Each objtr As ClsVLCDataUploaderManualdetail In obj.arrVLCDetail
                Dim qry As String = "select TSPL_MCC_MASTER.Plant_Code,TSPL_LOCATION_MASTER.Location_Desc as PlantName,TSPL_MCC_MASTER.MCC_Code,TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader, 
                    TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Name, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,
                    TSPL_MP_MASTER.MP_Code,TSPL_MP_MASTER.MP_Name,TSPL_MP_MASTER.MP_Code_VLC_Uploader,TSPL_MP_MASTER.Telphone from TSPL_MP_MASTER 
                    left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code='" + obj.VLC_Code + "'
                    left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
                    left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_MCC_MASTER.Plant_Code
                    where TSPL_MP_MASTER.MP_Code='" + objtr.Farmer_Code + "' and len(isnull(TSPL_MP_MASTER.Telphone,''))>0"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim objSMSH As New clsSMSHead()
                    objSMSH.SMS_Text = strSMSContent

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.PlantCode, clsCommon.myCstr(dt.Rows(0)("Plant_Code")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.PlantName, clsCommon.myCstr(dt.Rows(0)("PlantName")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.MCCCode, clsCommon.myCstr(dt.Rows(0)("MCC_Code")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.MCCName, clsCommon.myCstr(dt.Rows(0)("MCC_NAME")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.MCCUploaderCode, clsCommon.myCstr(dt.Rows(0)("Mcc_Code_VLC_Uploader")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCCode, clsCommon.myCstr(dt.Rows(0)("VLC_Code")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCName, clsCommon.myCstr(dt.Rows(0)("VLC_Name")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCUploaderCode, clsCommon.myCstr(dt.Rows(0)("VLC_Code_VLC_Uploader")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.MPCode, clsCommon.myCstr(dt.Rows(0)("MP_Code")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderMP, clsCommon.myCstr(dt.Rows(0)("MP_Name")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.MPUploaderCode, clsCommon.myCstr(dt.Rows(0)("MP_Code_VLC_Uploader")))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderDate, clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy"))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderShift, obj.Shift)

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderMilkType, obj.Dock_Collection_Milk_Type)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderQty, clsCommon.myFormat(objtr.Qty))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.UOM, objtr.Unit_Code)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderFat, clsCommon.myFormat(objtr.FatPer))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderSNF, clsCommon.myFormat(objtr.SNFPer))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderRate, clsCommon.myFormat(objtr.Rate))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.VLCDataUploaderAmt, clsCommon.myFormat(objtr.Amount))

                    objSMSH.arrMobilNo = New List(Of String)()
                    objSMSH.arrMobilNo.Add(clsCommon.myCstr(dt.Rows(0)("Telphone")))
                    objSMSH.SaveData(clsUserMgtCode.frmVlcdataUploadar, objSMSH, trans)
                    objSMSH = Nothing
                End If
            Next
        End If
        'SMSCode End Start
    End Sub
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As ClsVLCDataUploaderManual
        Return GetData(strCode, NavType, Nothing)
    End Function
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction) As ClsVLCDataUploaderManual
        Dim obj As ClsVLCDataUploaderManual = Nothing
        Dim Arr As List(Of ClsVLCDataUploaderManual) = Nothing
        Dim qry As String = "Select TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date, " &
            " TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code,TSPL_MCC_ROUTE_MASTER.Route_Name, " &
            " TSPL_VLC_DATA_UPLOADER_MASTER.Shift,TSPL_VLC_MASTER_HEAD.MCC AS MCC_CODE,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_DATA_UPLOADER_MASTER.Dock_Collection_Milk_Type from TSPL_VLC_DATA_UPLOADER_MASTER " &
            " Left Outer Join TSPL_MCC_ROUTE_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code= TSPL_MCC_ROUTE_MASTER.Route_Code " &
            " Left Outer Join TSPL_VLC_MASTER_HEAD on TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code " &
            " Left Outer Join TSPL_MCC_MASTER on TSPL_VLC_MASTER_HEAD.MCC=TSPL_MCC_MASTER.MCC_CODE where 2=2 "
        Dim whrclas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code = (select MIN(Document_Code) from TSPL_VLC_DATA_UPLOADER_MASTER WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Last
                qry += " and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code = (select Max(Document_Code) from TSPL_VLC_DATA_UPLOADER_MASTER WHERE 1=1 " + whrclas + " )"
            Case NavigatorType.Current
                qry += " and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code = '" + strCode + "' "
            Case NavigatorType.Next
                qry += " and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code = (select Min(Document_Code) from TSPL_VLC_DATA_UPLOADER_MASTER where Document_Code>'" + strCode + "' " + whrclas + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code = (select Max(Document_Code) from TSPL_VLC_DATA_UPLOADER_MASTER where Document_Code<'" + strCode + "' " + whrclas + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New ClsVLCDataUploaderManual()
            obj.Document_Code = clsCommon.myCstr(dt.Rows(0)("Document_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.VLC_Code = clsCommon.myCstr(dt.Rows(0)("VLC_Code"))
            obj.VLC_NAME = clsCommon.myCstr(dt.Rows(0)("VLC_Name"))
            obj.Route_Code = clsCommon.myCstr(dt.Rows(0)("Route_Code"))
            obj.Route_Name = clsCommon.myCstr(dt.Rows(0)("Route_Name"))
            obj.Shift = clsCommon.myCstr(dt.Rows(0)("Shift"))
            obj.MCC_Code = clsCommon.myCstr(dt.Rows(0)("MCC_Code"))
            obj.MCC_Name = clsCommon.myCstr(dt.Rows(0)("MCC_Name"))
            obj.Dock_Collection_Milk_Type = clsCommon.myCstr(dt.Rows(0)("Dock_Collection_Milk_Type"))
            obj.arrVLCDetail = ClsVLCDataUploaderManualdetail.getData(obj.Document_Code, trans)
        End If
        Return obj
    End Function

    Public Shared Function DeleteData(ByVal arr As ArrayList) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            For Each strDocNo As String In arr
                DeleteData(strDocNo, trans)
            Next
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
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
            Dim obj As ClsVLCDataUploaderManual = ClsVLCDataUploaderManual.GetData(strDocNo, NavigatorType.Current, trans)
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0 Then
                Dim mcc_Code As String = clsDBFuncationality.getSingleValue("select location_Code from tspl_mcc_route_master inner join tspl_location_master on location_Code=mcc_Code where route_Code='" & obj.Route_Code & "'", trans)
                clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleMCCMilkProcurement, clsUserMgtCode.FrmVLCDataUploaderManual, mcc_Code, obj.Document_Date, trans)
                Dim qry As String = ""
                qry = "delete from TSPL_VLC_DATA_UPLOADER_DETAIL where Document_Code='" + strDocNo + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_VLC_DATA_UPLOADER_MASTER where Document_Code='" + strDocNo + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                '' LOG FOR SYNC DATA
                clsSyncHeadTables.SaveSyncDelete("TSPL_VLC_DATA_UPLOADER_MASTER", strDocNo, trans)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class
Public Class ClsVLCDataUploaderManualdetail
#Region "Variable"
    Public Document_Code As String = Nothing
    Public Farmer_Code As String = Nothing
    Public Farmer_Code_vlc_Uploader As String = Nothing
    Public Farmer_Name As String = Nothing
    Public Unit_Code As String = Nothing
    Public Qty As Double = 0
    Public FatPer As Double = 0
    Public SNFPer As Double = 0
    Public Rate As Double = 0
    Public Amount As Double = 0
    Public Reject_Type As String
    Public Milk_Type As String
#End Region
    Public Shared Function saveData(ByVal TransDate As Date, ByVal MCC As String, ByVal VLC As String, ByVal arrObj As List(Of ClsVLCDataUploaderManualdetail), ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim issaved As Boolean = True
            Dim coll As Hashtable

            If arrObj IsNot Nothing Then

                For Each obj As ClsVLCDataUploaderManualdetail In arrObj
                    coll = New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Document_Code", strDocNo)
                    clsCommon.AddColumnsForChange(coll, "Farmer_Code", obj.Farmer_Code)
                    clsCommon.AddColumnsForChange(coll, "Unit_Code", obj.Unit_Code)
                    clsCommon.AddColumnsForChange(coll, "Qty", obj.Qty)
                    clsCommon.AddColumnsForChange(coll, "FatPer", obj.FatPer)
                    clsCommon.AddColumnsForChange(coll, "SNFPer", obj.SNFPer)
                    clsCommon.AddColumnsForChange(coll, "Rate", obj.Rate)
                    clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                    clsCommon.AddColumnsForChange(coll, "Reject_Type", obj.Reject_Type, True)
                    clsCommon.AddColumnsForChange(coll, "Milk_Type", obj.Milk_Type, True)
                    If clsCommon.myLen(obj.Reject_Type) <= 0 Then
                        Dim objIncetive As clsMPIncetiveDetail = clsMPIncetive.GetIncentive(obj.SNFPer, MCC, VLC, TransDate, trans)
                        If objIncetive IsNot Nothing AndAlso clsCommon.myLen(objIncetive.TRCode) > 0 Then
                            clsCommon.AddColumnsForChange(coll, "Incentive_TRCode", objIncetive.TRCode)
                            clsCommon.AddColumnsForChange(coll, "Incentive_Amt", Math.Round(obj.Qty * objIncetive.Slab_Value, 2))
                        End If
                    End If
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VLC_DATA_UPLOADER_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function getData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As List(Of ClsVLCDataUploaderManualdetail)
        Try
            Dim arrObj As List(Of ClsVLCDataUploaderManualdetail) = Nothing
            Dim obj As ClsVLCDataUploaderManualdetail = Nothing
            Dim qry As String = "Select TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code,Mp_Code_vlc_Uploader,TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code,TSPL_VLC_DATA_UPLOADER_DETAIL.Unit_Code,TSPL_VLC_DATA_UPLOADER_DETAIL.Qty,TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer,TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer,TSPL_MP_MASTER.MP_Name as [Farmer_Name],TSPL_VLC_DATA_UPLOADER_DETAIL.Rate,TSPL_VLC_DATA_UPLOADER_DETAIL.Amount,TSPL_VLC_DATA_UPLOADER_DETAIL.Reject_Type,TSPL_VLC_DATA_UPLOADER_DETAIL.Milk_Type 
from TSPL_VLC_DATA_UPLOADER_DETAIL 
Left Outer Join TSPL_MP_MASTER On TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code=TSPL_MP_MASTER .MP_Code  
where Document_Code='" & strDocNo & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                arrObj = New List(Of ClsVLCDataUploaderManualdetail)
                For i As Integer = 0 To dt.Rows.Count - 1
                    obj = New ClsVLCDataUploaderManualdetail()
                    obj.Document_Code = clsCommon.myCstr(dt.Rows(i)("Document_Code"))
                    obj.Farmer_Code = clsCommon.myCstr(dt.Rows(i)("Farmer_Code"))
                    obj.Farmer_Code_vlc_Uploader = clsCommon.myCstr(dt.Rows(i)("Mp_Code_vlc_Uploader"))
                    obj.Farmer_Name = clsCommon.myCstr(dt.Rows(i)("Farmer_Name"))
                    obj.Unit_Code = clsCommon.myCstr(dt.Rows(i)("Unit_Code"))
                    obj.FatPer = clsCommon.myCdbl(dt.Rows(i)("FatPer"))
                    obj.SNFPer = clsCommon.myCdbl(dt.Rows(i)("SNFPer"))
                    obj.Rate = clsCommon.myCdbl(dt.Rows(i)("Rate"))
                    obj.Amount = clsCommon.myCdbl(dt.Rows(i)("Amount"))
                    obj.Qty = clsCommon.myCdbl(dt.Rows(i)("Qty"))
                    obj.Reject_Type = clsCommon.myCstr(dt.Rows(i)("Reject_Type"))
                    obj.Milk_Type = clsCommon.myCstr(dt.Rows(i)("Milk_Type"))
                    arrObj.Add(obj)
                Next
            End If
            Return arrObj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

