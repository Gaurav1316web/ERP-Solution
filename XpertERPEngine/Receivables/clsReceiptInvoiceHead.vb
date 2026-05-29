''Changes done by priti mam replace by balwinder

Imports common
Imports System.Data
Imports System.Data.SqlClient
Imports Telerik.WinControls.UI
Public Class clsReceiptInvoiceHead
#Region "Variables"
    Public Receipt_Location As String = Nothing
    Public Receipt_Tax_Amt As Double = 0
    Public Invoice_Tax_Amt As Double = 0
    Public Document_Date As DateTime
    Public Doc_Code As String = Nothing
    Public Receipt_No As String = Nothing
    Public Description As String = Nothing
    Public POSTED As ERPTransactionStatus = ERPTransactionStatus.Pending
    Dim Created_By As String = Nothing
    Dim Created_Date As DateTime = Nothing
    Dim Modify_By As String = Nothing
    Dim Modify_Date As DateTime = Nothing
    Dim Comp_Code As String = Nothing
    Public UOMMO As String = Nothing
    Public Arr As List(Of clsReceiptInvoiceDetails) = Nothing

#End Region

    Public Shared Function funPrintServiceInvoicePrint(ByVal Form_ID As String, ByVal strCanceldelete As String, ByVal txtDate As DateTime, ByVal StrCode As String, ByVal IsEInvoiceApply As Integer) As Boolean
        Dim TSPL_Customer_Invoice_Head As String = Nothing
        Dim TSPL_Customer_Invoice_Detail As String = Nothing

        If clsCommon.CompairString(strCanceldelete, "Cancel") = CompairStringResult.Equal Then

            TSPL_Customer_Invoice_Head = "TSPL_Customer_Invoice_Head_Cancel_Data"
            TSPL_Customer_Invoice_Detail = "TSPL_Customer_Invoice_Detail_Cancel_Data"
        ElseIf clsCommon.CompairString(strCanceldelete, "Delete") = CompairStringResult.Equal Then
            TSPL_Customer_Invoice_Head = "TSPL_Customer_Invoice_Head_Delete_data"
            TSPL_Customer_Invoice_Detail = "TSPL_Customer_Invoice_Detail_Delete_data"
        Else
            TSPL_Customer_Invoice_Head = "TSPL_Customer_Invoice_Head"
            TSPL_Customer_Invoice_Detail = "TSPL_Customer_Invoice_Detail"
        End If

        Dim qry As String = "   Select"

        If clsCommon.CompairString(strCanceldelete, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCanceldelete, "Delete") = CompairStringResult.Equal Then
            qry += " 'Cancelled' As Report_Status, "
        Else
            qry += " '' As Report_Status, "
        End If
        qry += " tspl_company_master.Access_Officer as FSSAI, cast(" + TSPL_Customer_Invoice_Head + ".BarCode_Img As image) As BarCode_Img, isnull(" + TSPL_Customer_Invoice_Head + ".IRN_No,'') as IRN_No,isnull (" + TSPL_Customer_Invoice_Head + ".Ack_No,'') as Ack_No,case when len(isnull (" + TSPL_Customer_Invoice_Head + ".Ack_No,'')) > 0 then convert (varchar,  " + TSPL_Customer_Invoice_Head + ".Ack_Date,103) else ''  end as Ack_Date, " + clsCommon.myCstr(IsEInvoiceApply) + " as  IsEInvoiceApply,'1' as CopyType, TSPL_COMPANY_MASTER.TIN_NO , TSPL_COMPANY_MASTER.CST_LST ,TSPL_COMPANY_MASTER.Pan_No,tspl_location_master.PAN_NO as LocationPAN," &
                                    " TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd ," &
                                    " " &
                                    " tspl_state_master_For_Comp.GST_STATE_Code as Comp_GST_STATE_CODE , tspl_state_master_For_Comp.State_Name as Comp_State_Name,tspl_state_master_For_Comp.STATE_CODE as Comp_State_Code,TSPL_COMPANY_MASTER.Email as Comp_Email, TSPL_COMPANY_MASTER.CINNO as Comp_CINNO, TSPL_COMPANY_MASTER.GSTReg_No as Comp_GSTReg_No," &
                                    " TSPL_COMPANY_MASTER.Add1 as Comp_Add1, TSPL_COMPANY_MASTER.Add2 as Comp_Add2, TSPL_COMPANY_MASTER.add3 as Comp_Add3," &
                                    " " &
                                    " XXX.Customer_Code ,XXX.Customer_Name ," &
                                    " TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1, TSPL_CUSTOMER_MASTER.Add2 as Cust_Add2, TSPL_CUSTOMER_MASTER .Add3 as Cust_Add3,TSPL_CUSTOMER_MASTER.GSTNO AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,TSPL_STATE_MASTER.state_name AS Vendor_State_Name,xxx.Tin_No as Customer_Tin_No,XXX.PAN as Customer_PAN_NO," &
                                    " " &
                                    " XXX.Loc_Code,locAdd,XXX.Loc_Add1,XXX.Loc_Add2,XXX.Loc_Add3,XXX.Loc_City_Code,XXX.Loc_City_Name,XXX.Loc_State_Name,XXX.Loc_Pin_code,XXX.Location_Desc," &
                                    " TSPL_LOCATION_MASTER.Email as Loc_Email,tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo," &
                                    " " &
                                    "  XXX.RefDocNo    , XXX.Description  ,XXX.Document_No ,Convert (varchar,XXX.Document_Date,103) as Document_Date ,XXX.Status , XXX.AddChargeCode ,XXX.AddChargeDesc,XXX.SAC_Code,XXX.SNo,XXX.Amount_less_Discount,XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Created_By ,XXX.Modify_By , XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,XXX.Taxdesc1,XXX.Taxdesc2,XXX.Taxdesc3,XXX.Taxdesc4,XXX.Taxdesc5,XXX.Taxdesc6,XXX.Taxdesc7,XXX.Taxdesc8,XXX.Taxdesc9,XXX.Taxdesc10,XXX.TAX1_Amt,XXX.TAX2_Amt,XXX.TAX3_Amt,XXX.TAX4_Amt,XXX.TAX5_Amt,XXX.TAX6_Amt,XXX.TAX7_Amt,XXX.TAX8_Amt,XXX.TAX9_Amt,XXX.TAX10_Amt,XXX.Tax1_Rate,XXX.Tax2_Rate,XXX.Tax3_Rate,XXX.Tax4_Rate,XXX.Tax5_Rate,XXX.Tax6_Rate,XXX.Tax7_Rate,XXX.Tax8_Rate,XXX.Tax9_Rate,XXX.Tax10_Rate,TSPL_COMPANY_MASTER.Bank_Name,TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.BankBranchAddress,TSPL_LOCATION_MASTER.ACType,TSPL_LOCATION_MASTER.accountholdername,TSPL_CUSTOMER_MASTER.PIN_NO  ,TSPL_CUSTOMER_MASTER.PIN_Code," + TSPL_Customer_Invoice_Head + ".Balance_Amt, " + TSPL_Customer_Invoice_Head + ".Discount_Amount, " + TSPL_Customer_Invoice_Head + ".Total_Tax  from (" &
                                    "  " &
                                    "  select distinct Loc_Code,RefDocNo, locAdd,Final.Loc_Add1,final.Loc_Add2,Final.Loc_Add3,final.Loc_City_Code,Final.Loc_City_Name,Final.Loc_State_Name,Final.Loc_Pin_code,Location_Desc,Tin_No,PAN ,final.Document_No ,final.Document_Date, final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , final.Modify_By  ,final .Comp_Code,Description ,final.AddChargeCode ,final.AddChargeDesc,final.SAC_Code,final.SNo,final.Amount_less_Discount," &
                                    "  final.Taxdesc1,final.Taxdesc2,final.Taxdesc3,final.Taxdesc4,final.Taxdesc5,final.Taxdesc6,final.Taxdesc7,final.Taxdesc8,final.Taxdesc9,final.Taxdesc10,final.TAX1_Amt,final.TAX2_Amt,final.TAX3_Amt,final.TAX4_Amt,final.TAX5_Amt,final.TAX6_Amt,final.TAX7_Amt,final.TAX8_Amt,final.TAX9_Amt,final.TAX10_Amt,final.Tax1_Rate,final.Tax2_Rate,final.Tax3_Rate,final.Tax4_Rate,final.Tax5_Rate,final.Tax6_Rate,final.Tax7_Rate,final.Tax8_Rate,final.Tax9_Rate,final.Tax10_Rate" &
                                    "   from (" &
                                    "  " &
                                    "  select isnull(" + TSPL_Customer_Invoice_Head + ".RefDocNo,'') as RefDocNo, " + TSPL_Customer_Invoice_Head + ".Loc_Code, TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end + case When TSPL_LOCATION_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.City_Code, 103) end+ Case When TSPL_LOCATION_MASTER.State='' Then '' else ', '+Convert(Varchar, loc_state.STATE_NAME ) end +  Case When TSPL_LOCATION_MASTER.Pin_code='' Then '' Else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Pin_code, 103)  end  as locAdd, " &
                                    "  TSPL_LOCATION_MASTER.Add1 as Loc_Add1, TSPL_LOCATION_MASTER.Add2 as Loc_Add2, TSPL_LOCATION_MASTER.Add3 as Loc_Add3, TSPL_LOCATION_MASTER.City_Code as Loc_City_Code,TSPL_CITY_MASTER.City_Name as Loc_City_Name , loc_state.STATE_NAME as Loc_State_Name,TSPL_LOCATION_MASTER.Pin_code as Loc_Pin_code ," &
                                    "  Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN," + TSPL_Customer_Invoice_Head + ".Description," + TSPL_Customer_Invoice_Detail + ".AddChargeCode," + TSPL_Customer_Invoice_Detail + ".AddChargeDesc,TSPL_Additional_Charges.SAC_Code," + TSPL_Customer_Invoice_Detail + ".SNo," + TSPL_Customer_Invoice_Detail + ".Amount_less_Discount," &
                                    "  " &
                                    "   " + TSPL_Customer_Invoice_Head + ".Document_No, Document_Date , case when " + TSPL_Customer_Invoice_Head + ".Status=1 then 'Authorized' else 'UnAuthorized' end as Status ,  case when " + TSPL_Customer_Invoice_Head + ".Document_Type='I' then 'Tax Invoice' else case when " + TSPL_Customer_Invoice_Head + ".Document_Type='D' then 'Debit Note' else case when " + TSPL_Customer_Invoice_Head + ".Document_Type='C' then 'Credit Note' else '' end end end as Document_Type, Account_Set,Document_Total as DocAmt,  Customer_Code ," + TSPL_Customer_Invoice_Head + ".Customer_Name,tspl_user_master.User_Name as Created_By,user_master_modify.User_Name as Modify_By  ," &
                                    " " &
                                    "   " + TSPL_Customer_Invoice_Head + ".Comp_Code ," &
                                    " " &
                                    "    isnull(Tspl_Tax1.Tax_Code,'') as Taxdesc1,isnull(Tspl_Tax2.Tax_Code,'') as Taxdesc2,isnull(Tspl_Tax3.Tax_Code,'') as Taxdesc3,isnull(Tspl_Tax4.Tax_Code,'') as Taxdesc4,isnull(Tspl_Tax5.Tax_Code,'') as Taxdesc5,isnull(Tspl_Tax6.Tax_Code,'') as Taxdesc6,isnull(Tspl_Tax7.Tax_Code,'') as Taxdesc7,isnull(Tspl_Tax8.Tax_Code,'') as Taxdesc8,isnull(Tspl_Tax8.Tax_Code,'') as Taxdesc9,isnull(Tspl_Tax10.Tax_Code,'') as Taxdesc10, " + TSPL_Customer_Invoice_Detail + ".Tax1_Rate," + TSPL_Customer_Invoice_Detail + ".Tax2_Rate," + TSPL_Customer_Invoice_Detail + ".Tax3_Rate," + TSPL_Customer_Invoice_Detail + ".Tax4_Rate," + TSPL_Customer_Invoice_Detail + ".Tax5_Rate," + TSPL_Customer_Invoice_Detail + ".Tax6_Rate," + TSPL_Customer_Invoice_Detail + ".Tax7_Rate," + TSPL_Customer_Invoice_Detail + ".Tax8_Rate," + TSPL_Customer_Invoice_Detail + ".Tax9_Rate," + TSPL_Customer_Invoice_Detail + ".Tax10_Rate," &
                                    "    " + TSPL_Customer_Invoice_Detail + ".TAX1_Amt," + TSPL_Customer_Invoice_Detail + ".TAX2_Amt," + TSPL_Customer_Invoice_Detail + ".TAX3_Amt," + TSPL_Customer_Invoice_Detail + ".TAX4_Amt," + TSPL_Customer_Invoice_Detail + ".TAX5_Amt," + TSPL_Customer_Invoice_Detail + ".TAX6_Amt," + TSPL_Customer_Invoice_Detail + ".TAX7_Amt," + TSPL_Customer_Invoice_Detail + ".TAX8_Amt," + TSPL_Customer_Invoice_Detail + ".TAX9_Amt," + TSPL_Customer_Invoice_Detail + ".TAX10_Amt," + TSPL_Customer_Invoice_Detail + ".Total_Tax" &
                                    "  " &
                                    "  from " + TSPL_Customer_Invoice_Head + "  " &
                                    " left outer join " + TSPL_Customer_Invoice_Detail + " on  " + TSPL_Customer_Invoice_Detail + ".Document_No = " + TSPL_Customer_Invoice_Head + ".Document_No" &
                                    "  " &
                                    "     left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=" + TSPL_Customer_Invoice_Head + ".Customer_Code "

        If objCommonVar.RCDFCFP Then
            qry += "  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Loc_Segment_Code=" + TSPL_Customer_Invoice_Head + ".Loc_Code   "
        Else
            qry += "  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=" + TSPL_Customer_Invoice_Head + ".Loc_Code   "
        End If

        qry += " 	left outer join TSPL_STATE_MASTER as loc_state on loc_state.STATE_CODE =TSPL_LOCATION_MASTER.State   left outer join tspl_user_master on tspl_user_master.User_Code=" + TSPL_Customer_Invoice_Head + ".Created_By  left outer join tspl_user_master as user_master_modify on user_master_modify.User_Code=" + TSPL_Customer_Invoice_Head + ".Modify_By " &
                                    " left outer join TSPL_Additional_Charges on 	TSPL_Additional_Charges.Code = " + TSPL_Customer_Invoice_Detail + ".AddChargeCode" &
                                    " " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax1 on Tspl_Tax1.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX1 " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax2 on Tspl_Tax2.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX2  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax3 on Tspl_Tax3.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX3  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax4 on Tspl_Tax4.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX4  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax5 on Tspl_Tax5.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX5  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax6 on Tspl_Tax6.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX6  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax7 on Tspl_Tax7.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX7  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax8 on Tspl_Tax8.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX8  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax9 on Tspl_Tax9.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX9  " &
                                    "  left outer join TSPL_TAX_MASTER Tspl_Tax10 on Tspl_Tax10.Tax_Code =" + TSPL_Customer_Invoice_Head + ".TAX10 " &
                                    "  left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code = TSPL_LOCATION_MASTER.City_Code " &
                                    " 	where " + TSPL_Customer_Invoice_Head + ".Document_No ='" & StrCode & "'" &
                                    " " &
                                    " 	 )final   )XXX left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = XXX.Customer_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE "
        If objCommonVar.RCDFCFP Then
            qry += " left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Loc_Segment_Code = XXX.Loc_Code "
        Else
            qry += " left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code = XXX.Loc_Code "
        End If


        qry += "  left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state    " &
                                    " " &
                                    "  left outer join tspl_state_master as tspl_state_master_For_Comp on tspl_state_master_For_Comp.state_code = TSPL_COMPANY_MASTER.State" &
                                    "  left outer join " + TSPL_Customer_Invoice_Head + " on " + TSPL_Customer_Invoice_Head + ".Document_No=XXX.DOCUMENT_NO "

        '&
        '                        "  ) XXXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE' as CopyType1) YYY ON YYY.COL1=XXXX.CopyType ORDER BY YYY.COL2 ,XXXX.SNO  " &
        '                        " " &
        '                        " "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        qry = "  select TSPL_Additional_Charges.SAC_Code, sum (" + TSPL_Customer_Invoice_Detail + ".Amount_Less_Discount) as Amount_Less_Discount  from " + TSPL_Customer_Invoice_Detail + " " &
                  "  left outer join TSPL_Additional_Charges on 	TSPL_Additional_Charges.Code = " + TSPL_Customer_Invoice_Detail + ".AddChargeCode " &
                  " where  Document_No ='" & StrCode & "' group by TSPL_Additional_Charges.SAC_Code "
        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            'frmCRV.funsubreportWithdt(CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvc", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            If objCommonVar.RCDFCFP Then
                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvcnew", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            Else
                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvcnewForAll", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            End If
            frmCRV = Nothing
        Else
            clsCommon.MyMessageBoxShow("No Data Found")
        End If

        Return True
    End Function




    Public Shared Function funARInvoicePrint(ByVal Form_ID As String, ByVal isCancel As Boolean, ByVal txtDate As DateTime, ByVal StrCode As String, ByVal IsEInvoiceApply As Integer, ByVal SettingCostCenterlevel As Boolean) As Boolean
        Dim TSPL_Customer_Invoice_Head As String = Nothing
        Dim TSPL_JOURNAL_DETAILS As String = Nothing
        Dim TSPL_JOURNAL_MASTER As String = Nothing
        If isCancel Then
            TSPL_Customer_Invoice_Head = "TSPL_Customer_Invoice_Head_cancel_data"
            TSPL_JOURNAL_DETAILS = "TSPL_JOURNAL_DETAILS_cancel_data"
            TSPL_JOURNAL_MASTER = "TSPL_JOURNAL_MASTER_cancel_data"
        Else
            TSPL_Customer_Invoice_Head = "TSPL_Customer_Invoice_Head"
            TSPL_JOURNAL_DETAILS = "TSPL_JOURNAL_DETAILS"
            TSPL_JOURNAL_MASTER = "TSPL_JOURNAL_MASTER"

        End If

        Dim qry As String = "SELECT"
        If isCancel Then
            qry += " 'Cancelled' As Report_Status, "
        Else
            qry += " '' As Report_Status, "
        End If
        qry += "cast(BarCode_Img as image) As BarCode_Img,isnull (IRN_No,'') as IRN_No,isnull (Ack_No,'') as Ack_No,case when len(isnull (Ack_No,'')) > 0 then convert (varchar, Ack_Date,103) else ''  end as Ack_Date, " + clsCommon.myCstr(IsEInvoiceApply) + " as  IsEInvoiceApply," &
              " XXX.RefDocNo, tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo ,TSPL_CUSTOMER_MASTER.GSTNO AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,TSPL_STATE_MASTER.state_name AS Vendor_State_Name, XXX.Loc_Code, locAdd, XXX.Location_Desc,xxx.Tin_No,XXX.PAN,TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd , " &
            "XXX.Description,XXX.Account_Code,XXX.Account_Desc ,XXX.DrAmt ,XXX.CrAmt ,XXX.Document_No ,XXX.Document_Date,XXX.Status , " &
            "XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Customer_Code ,XXX.Customer_Name ,XXX.Created_By ,XXX.Modify_By ,XXX.Detail_Line_No , " &
            "XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,XXX.Cost_Centre_Code,XXX.Cost_Center_Fin_Name,XXX.Hirerachy_Code,XXX.HIRERACHY_Name,XXX.Hirerachy_Code3,XXX.Hirerachy_Code4,XXX.TapalNo,XXX.DateAndTime   from " &
            "(select distinct Loc_Code,RefDocNo, locAdd,Location_Desc,Tin_No,PAN,(final.Account_Code),final.Account_Desc ,final.DrAmt ,final.CrAmt ,final.Document_No ,final.Document_Date, " &
            "final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , " &
            "final.Modify_By ,final.Detail_Line_No ,final .Comp_Code,Description ,final.Cost_Centre_Code,final.Cost_Center_Fin_Name,final.Hirerachy_Code,final.HIRERACHY_Name,final.Hirerachy_Code3 ,final.Hirerachy_Code4,final.TapalNo,final.DateAndTime  from " &
            "(select isnull(" + TSPL_Customer_Invoice_Head + ".RefDocNo,'') as RefDocNo, " + TSPL_Customer_Invoice_Head + ".Loc_Code," &
            " TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end + case When TSPL_LOCATION_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.City_Code, 103) end+ Case When TSPL_LOCATION_MASTER.State='' Then '' else ', '+Convert(Varchar, loc_state.STATE_NAME ) end +  Case When TSPL_LOCATION_MASTER.Pin_code='' Then '' Else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Pin_code, 103)  end  as locAdd," &
            " Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN," + TSPL_Customer_Invoice_Head + ".Description,case  when " + TSPL_JOURNAL_DETAILS + ".Amount >=0 then " + TSPL_JOURNAL_DETAILS + ".Amount else 0 end as DrAmt ," &
            " case  when " + TSPL_JOURNAL_DETAILS + ".Amount <0 then " + TSPL_JOURNAL_DETAILS + ".Amount*-1 else 0 end as CrAmt, " + TSPL_Customer_Invoice_Head + ".Document_No, Document_Date , case when " + TSPL_Customer_Invoice_Head + ".Status=1 then 'Authorized' else 'UnAuthorized' end as Status , " &
            " case WHEN isnull(AgainstServiceInvoice,'')='Y' then 'Tax Invoice' when " + TSPL_Customer_Invoice_Head + ".Document_Type='I' then 'Invoice' else case when " + TSPL_Customer_Invoice_Head + ".Document_Type='D' then 'Debit Note' else case when " + TSPL_Customer_Invoice_Head + ".Document_Type='C' then 'Credit Note' else '' end end end as Document_Type," &
            " Account_Set,Document_Total as DocAmt,  Customer_Code ," + TSPL_Customer_Invoice_Head + ".Customer_Name,tspl_user_master.User_Name as Created_By,user_master_modify.User_Name as Modify_By  , " + TSPL_JOURNAL_DETAILS + ".Detail_Line_No as Detail_Line_No ," &
            " " + TSPL_JOURNAL_DETAILS + ".Account_code as Account_Code , " + TSPL_JOURNAL_DETAILS + ".Account_Desc as Account_Desc ," + TSPL_JOURNAL_DETAILS + ".Amount , 0 as Discount ," + TSPL_JOURNAL_DETAILS + ".Amount as Amount_less_Discount , 0 Total_Tax ," + TSPL_JOURNAL_DETAILS + ".Amount as Total_Amount  ," &
            " " + TSPL_Customer_Invoice_Head + ".Comp_Code ," + TSPL_JOURNAL_DETAILS + ".Cost_Centre_Code,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name," + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code,TSPL_HIRERACHY_LEVEL_MASTER.Description as HIRERACHY_Name," + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code3 ," + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code4," + TSPL_Customer_Invoice_Head + ".TapalNo," + TSPL_Customer_Invoice_Head + ".DateAndTime  from " + TSPL_Customer_Invoice_Head + " " &
            " left outer join " + TSPL_JOURNAL_MASTER + " on " + TSPL_JOURNAL_MASTER + ".Source_Doc_No = " + TSPL_Customer_Invoice_Head + ".Document_No left outer join " + TSPL_JOURNAL_DETAILS + " on " + TSPL_JOURNAL_DETAILS + ".Voucher_No = " + TSPL_JOURNAL_MASTER + ".Voucher_No   " &
            " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code= " + TSPL_Customer_Invoice_Head + ".Customer_Code  left outer join TSPL_LOCATION_MASTER on left(TSPL_LOCATION_MASTER.Location_Code,3)=" + TSPL_Customer_Invoice_Head + ".Loc_Code  " &
            " left outer join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code = " + TSPL_JOURNAL_DETAILS + ".Cost_Centre_Code  left outer join TSPL_HIRERACHY_LEVEL_MASTER on TSPL_HIRERACHY_LEVEL_MASTER.Hirerachy_Code = " + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code  left outer join TSPL_STATE_MASTER as loc_state on loc_state.STATE_CODE =TSPL_LOCATION_MASTER.State  " &
            " left outer join tspl_user_master on tspl_user_master.User_Code=" + TSPL_Customer_Invoice_Head + ".Created_By " &
            " left outer join tspl_user_master as user_master_modify on user_master_modify.User_Code=" + TSPL_Customer_Invoice_Head + ".Modify_By " &
            "where " + TSPL_Customer_Invoice_Head + ".Document_No ='" + StrCode + "' and Rejected_Type='N' )final   )XXX left outer join " &
            "TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = XXX.Customer_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code = XXX.Loc_Code left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state  left outer join " + TSPL_Customer_Invoice_Head + " on " + TSPL_Customer_Invoice_Head + ".Document_No=XXX.Document_No  order by XXX.Detail_Line_No  "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            If SettingCostCenterlevel Then
                frmCRV.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptAPInvc_Hierarchy", "AR INVOICE", clsCommon.myCDate(dt.Rows(0)("Document_Date")))
            Else
                frmCRV.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptAPInvc", "AR INVOICE", clsCommon.myCDate(dt.Rows(0)("Document_Date")))
            End If
            frmCRV = Nothing
        Else
            'clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            clsCommon.MyMessageBoxShow("No Data Found")

        End If

        Return True
    End Function


    Public Function SaveData(ByVal obj As clsReceiptInvoiceHead, ByVal isNewEntry As Boolean) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            SaveData(obj, isNewEntry, "", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Function SaveData(ByVal obj As clsReceiptInvoiceHead, ByVal isNewEntry As Boolean, ByVal strAdjustmentNoTemp As String, ByVal trans As SqlTransaction) As Boolean
        Dim cntr As Integer = 0
        Dim isSaved As Boolean = True
        Try
            Dim qry As String = "delete from TSPL_Receipt_InvoiceMapping_Detail where Doc_Code='" + obj.Doc_Code + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Dim strDocNo As String = ""
            Dim dtCurrent As DateTime = clsCommon.GETSERVERDATE(trans)
            If clsCommon.myLen(strAdjustmentNoTemp) > 0 Then
                obj.Doc_Code = strAdjustmentNoTemp
                'isNewEntry = True
            Else
                isNewEntry = True
                If isNewEntry Then
                    obj.Doc_Code = clsERPFuncationality.GetNextCode(trans, clsCommon.GETSERVERDATE(trans), clsDocType.ReceiptInvoiceMapping, "", "")
                End If
            End If
            If (clsCommon.myLen(obj.Doc_Code) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If
            Dim dblchkBalnce As Double = GetReceiptBalanceTaxAmount(obj.Receipt_No, obj.Doc_Code, trans)
            Dim dblcalAmt As Double = 0
            For Each objtr As clsReceiptInvoiceDetails In obj.Arr
                dblcalAmt += objtr.Total_Tax_Amt
            Next
            If dblcalAmt > dblchkBalnce Then
                Throw New Exception("Receipt balance tax amount" + clsCommon.myCstr(dblchkBalnce) + " and used amount " + clsCommon.myCstr(dblcalAmt))
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Document_Date", clsCommon.GetPrintDate(obj.Document_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            clsCommon.AddColumnsForChange(coll, "Receipt_No", obj.Receipt_No)
            clsCommon.AddColumnsForChange(coll, "Receipt_Location", obj.Receipt_Location)
            clsCommon.AddColumnsForChange(coll, "Receipt_Tax_Amt", obj.Receipt_Tax_Amt)
            clsCommon.AddColumnsForChange(coll, "Invoice_Tax_Amt", obj.Invoice_Tax_Amt)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modified_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", clsCommon.GetPrintDate(dtCurrent, "dd/MMM/yyyy hh:mm tt"))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(dtCurrent, "dd/MMM/yyyy hh:mm tt"))
                clsCommon.AddColumnsForChange(coll, "Doc_Code", obj.Doc_Code)
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Receipt_InvoiceMapping_Head", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Receipt_InvoiceMapping_Head", OMInsertOrUpdate.Update, "Doc_Code='" + obj.Doc_Code + "'", trans)
            End If
            isSaved = isSaved AndAlso clsReceiptInvoiceDetails.SaveData(obj.Doc_Code, Arr, isNewEntry, trans)

        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function GetData(ByVal strDocNo As String, ByVal NavType As NavigatorType) As clsReceiptInvoiceHead

        Dim obj As clsReceiptInvoiceHead = Nothing
        Dim qry As String = "SELECT * from TSPL_Receipt_InvoiceMapping_Head  where 2=2"
        Dim whrClas As String = ""
        Select Case NavType
            Case NavigatorType.First
                qry += " and Doc_Code = (select MIN(Doc_Code) from TSPL_Receipt_InvoiceMapping_Head where 1=1 " + whrClas + ")"
            Case NavigatorType.Last
                qry += " and Doc_Code = (select Max(Doc_Code) from TSPL_Receipt_InvoiceMapping_Head where 1=1 " + whrClas + ")"
            Case NavigatorType.Next
                qry += " and Doc_Code = (select Min(Doc_Code) from TSPL_Receipt_InvoiceMapping_Head where Doc_Code>'" + strDocNo + "' " + whrClas + ")"
            Case NavigatorType.Previous
                qry += " and Doc_Code = (select Max(Doc_Code) from TSPL_Receipt_InvoiceMapping_Head where Doc_Code<'" + strDocNo + "' " + whrClas + ")"
            Case NavigatorType.Current
                qry += " and Doc_Code = '" + strDocNo + "'"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsReceiptInvoiceHead()
            obj.Doc_Code = clsCommon.myCstr(dt.Rows(0)("Doc_Code"))
            obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Receipt_No = clsCommon.myCstr(dt.Rows(0)("Receipt_No"))
            obj.Receipt_Location = clsCommon.myCstr(dt.Rows(0)("Receipt_Location"))
            obj.Invoice_Tax_Amt = clsCommon.myCdbl(dt.Rows(0)("Invoice_Tax_Amt"))
            obj.Receipt_Tax_Amt = clsCommon.myCdbl(dt.Rows(0)("Receipt_Tax_Amt"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.POSTED = IIf(clsCommon.myCdbl(dt.Rows(0)("isPosted")) = 1, ERPTransactionStatus.Approved, ERPTransactionStatus.Pending)

            qry = "SELECT* from TSPL_Receipt_InvoiceMapping_Detail  where  Doc_Code='" + obj.Doc_Code + "'"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsReceiptInvoiceDetails)
                Dim objTr As clsReceiptInvoiceDetails
                For Each dr As DataRow In dt.Rows
                    objTr = New clsReceiptInvoiceDetails()
                    objTr.Doc_Code = clsCommon.myCstr(dr("Doc_Code"))
                    objTr.InvoiceNo = clsCommon.myCstr(dr("InvoiceNo"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))
                    objTr.InvoiceLocation = clsCommon.myCstr(dr("InvoiceLocation"))
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(dr("Total_Tax_Amt"))
                    objTr.Tax1_Amt = clsCommon.myCdbl(dr("Tax1_Amt"))
                    objTr.Tax2_Amt = clsCommon.myCdbl(dr("Tax2_Amt"))
                    objTr.Tax3_Amt = clsCommon.myCdbl(dr("Tax3_Amt"))
                    objTr.Tax4_Amt = clsCommon.myCdbl(dr("Tax4_Amt"))
                    objTr.Tax5_Amt = clsCommon.myCdbl(dr("Tax5_Amt"))
                    objTr.Tax6_Amt = clsCommon.myCdbl(dr("Tax6_Amt"))
                    objTr.Tax7_Amt = clsCommon.myCdbl(dr("Tax7_Amt"))
                    objTr.Tax8_Amt = clsCommon.myCdbl(dr("Tax8_Amt"))
                    objTr.Tax9_Amt = clsCommon.myCdbl(dr("Tax9_Amt"))
                    objTr.Tax10_Amt = clsCommon.myCdbl(dr("Tax10_Amt"))

                    obj.Arr.Add(objTr)
                Next
            End If
        End If
        Return obj
    End Function


    Public Shared Function DeleteData(ByVal strCode As String) As Boolean
        Dim isSaved As Boolean = False
        If (clsCommon.myLen(strCode) <= 0) Then
            Throw New Exception("Document No not found to Delete")
        End If
        Dim obj As New clsReceiptInvoiceHead()
        obj = clsReceiptInvoiceHead.GetData(strCode, NavigatorType.Current)
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Doc_Code) > 0) Then
            Try
                If (clsCommon.CompairString(obj.Posted, "Y") = CompairStringResult.Equal) Then
                    Throw New Exception("Already Posted .")
                End If
                Dim qry As String = "delete from TSPL_Receipt_InvoiceMapping_Detail where Doc_Code='" + strCode + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from TSPL_Receipt_InvoiceMapping_Head where Doc_Code='" + strCode + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsReceiptInvoiceDetails In obj.Arr
                        qry = "Update tspl_sd_sale_invoice_head set IsAdvanceTaxGlEntry=0  where Document_Code ='" + objTr.InvoiceNo + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    Next
                End If

                If (isSaved) Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            Catch ex As Exception
                trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function
    Public Shared Function CreateGLEntryForAllCases(ByVal obj As clsReceiptInvoiceHead, ByVal trans As SqlTransaction) As Boolean
        Dim ArryLst As ArrayList = New ArrayList()
        Dim intInvoiceLocCount As Integer = 0
        Dim intInvoiceCount As Integer = 0
        Dim objRec As clsRcptEntryHeader
        objRec = clsRcptEntryHeader.GetData(obj.Receipt_No, NavigatorType.Current, trans)
        Dim StrInvoice = clsDBFuncationality.getSingleValue("select isnull((Select distinct ' '+TSPL_Receipt_InvoiceMapping_Detail.InvoiceNo+' ,  ' from TSPL_Receipt_InvoiceMapping_Detail  where  TSPL_Receipt_InvoiceMapping_Detail.Doc_CODE='" & obj.Doc_Code & "' for xml path('')),'')  as DocNo ", trans)
        Dim strRemarks = "Advance Tax Entry for Customer : " + objRec.Cust_Code + " Receipt No : " + obj.Receipt_No & " Incoice No : " + StrInvoice + " "
        intInvoiceCount = obj.Arr.Count
        For Each objTr As clsReceiptInvoiceDetails In obj.Arr
            If Not clsCommon.CompairString(obj.Receipt_Location, objTr.InvoiceLocation) = CompairStringResult.Equal Then
                intInvoiceLocCount += 1
            End If
        Next
 
        Dim isApplyBrachAccounting As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, trans)) = 1, True, False)
        If objRec.Tax_Amount_Advance > 0 Then
            Dim objTM As clsTaxMaster
            If objRec.TAX1_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX1, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX1)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX1)
                    End If

                    'Cases start here
                    If ((obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0) OrElse (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX1_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then

                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax1_Amt}
                            ArryLst.Add(Acc3)

                        Next
                    ElseIf ((obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax1_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax1_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc5)

                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax1_Amt}
                                ArryLst.Add(Acc3)
                            End If


                        Next
                    ElseIf (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX1_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX1_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc5)

                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX1_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next

                    ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX2_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax2, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax2)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax2)
                    End If

                    'Cases start here
                    If ((obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0) OrElse (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX2_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then

                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax2_Amt}
                            ArryLst.Add(Acc3)

                        Next
                    ElseIf ((obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax2_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax2_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax2_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX2_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX2_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX2_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX3_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX3, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX3)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX3)
                    End If

                    'Cases start here
                    If ((obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0) OrElse (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX3_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then

                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax3_Amt}
                            ArryLst.Add(Acc3)

                        Next
                    ElseIf ((obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax3_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax3_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax3_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX3_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX3_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX3_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX4_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX4, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX4)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX4)
                    End If

                    'Cases start here
                    If ((obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0) OrElse (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX4_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then

                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax4_Amt}
                            ArryLst.Add(Acc3)

                        Next
                    ElseIf ((obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax4_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax4_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax4_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX4_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX4_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX4_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX5_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax5, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax5)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax5)
                    End If

                    'Cases start here
                    If ((obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0) OrElse (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX5_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax5_Amt}
                            ArryLst.Add(Acc3)
                        Next
                    ElseIf ((obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0) OrElse (obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax5_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax5_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax5_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf (obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)
                            If isApplyBrachAccounting AndAlso Not clsCommon.CompairString(strInvoiceLoc, objRec.Location_GL_Code) = CompairStringResult.Equal Then


                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX5_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX5_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX5_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX6_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax6, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax6)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax5)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX6_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX6_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX7_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax7, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax7)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax7)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX7_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX7_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX8_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax8, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax8)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax8)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX8_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX8_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX9_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax9, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax9)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax9)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX9_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX9_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX10_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax10, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax10)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax10)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX10_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX10_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If
        End If
        Dim coll As New Hashtable()
        If clsCommon.myLen(objRec.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(objRec.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", objRec.CURRENCY_CODE)
            clsCommon.AddColumnsForChange(coll, "ConvRate", objRec.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ConvRateOld", objRec.ConvRate)
        End If
        clsJournalMaster.FunGrnlEntryWithTrans(objRec.Location_GL_Code, True, trans, clsCommon.myCDate(obj.Document_Date), obj.Description, "RC-AD", "Receipt Advance Tax Knock off", obj.Doc_Code, obj.Description, "C", objRec.Cust_Code, objRec.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, objRec.Reference, strRemarks, Nothing, coll)

        Return True

    End Function

    Public Shared Function CreateGLEntryForAllCasesOLD(ByVal obj As clsReceiptInvoiceHead, ByVal trans As SqlTransaction) As Boolean
        Dim TAX1_Amt As Double = 0
        Dim TAX2_Amt As Double = 0
        Dim TAX3_Amt As Double = 0
        Dim TAX4_Amt As Double = 0
        Dim TAX5_Amt As Double = 0
        Dim TAX6_Amt As Double = 0
        Dim TAX7_Amt As Double = 0
        Dim TAX8_Amt As Double = 0
        Dim TAX9_Amt As Double = 0
        Dim TAX10_Amt As Double = 0

        Dim TAX1Invoice_Amt As Double = 0
        Dim TAX2Invoice_Amt As Double = 0
        Dim TAX3Invoice_Amt As Double = 0
        Dim TAX4Invoice_Amt As Double = 0
        Dim TAX5Invoice_Amt As Double = 0
        Dim ArryLst As ArrayList = New ArrayList()
        Dim intInvoiceLocCount As Integer = 0
        Dim intInvoiceCount As Integer = 0
        Dim objRec As clsRcptEntryHeader
        objRec = clsRcptEntryHeader.GetData(obj.Receipt_No, NavigatorType.Current, trans)
        Dim StrInvoice = clsDBFuncationality.getSingleValue("select isnull((Select distinct ' '+TSPL_Receipt_InvoiceMapping_Detail.InvoiceNo+' ,  ' from TSPL_Receipt_InvoiceMapping_Detail  where  TSPL_Receipt_InvoiceMapping_Detail.Doc_CODE='" & obj.Doc_Code & "' for xml path('')),'')  as DocNo ", trans)
        Dim strRemarks = "Advance Tax Entry for Customer : " + objRec.Cust_Code + " Receipt No : " + obj.Receipt_No & " Incoice No : " + StrInvoice + " "
        intInvoiceCount = obj.Arr.Count
        For Each objTr As clsReceiptInvoiceDetails In obj.Arr
            If Not clsCommon.CompairString(obj.Receipt_Location, objTr.InvoiceLocation) = CompairStringResult.Equal Then
                intInvoiceLocCount += 1
            End If
        Next
        Dim qry As String = "select sum(isnull(TAX1_Amt,0) * RI) as TAX1_Amt,sum(isnull(TAX2_Amt,0) * RI) as TAX2_Amt,sum(isnull(TAX3_Amt,0) * RI) as TAX3_Amt, " & _
            "sum(isnull(TAX4_Amt,0) * RI) as TAX4_Amt,sum(isnull(TAX5_Amt,0) * RI) as TAX5_Amt ,sum(isnull(TAX6_Amt,0) * RI) as TAX6_Amt, " & _
            "sum(isnull(TAX7_Amt,0) * RI) as TAX7_Amt,sum(isnull(TAX8_Amt,0) * RI) as TAX8_Amt,sum(isnull(TAX9_Amt,0) * RI) as TAX9_Amt, " & _
            "sum(isnull(TAX10_Amt,0) * RI) as TAX10_Amt from (  " & _
            "select tax1_amt,tax2_amt,tax3_amt,tax4_amt,tax5_amt,tax6_amt,tax7_amt,tax8_amt,tax9_amt,tax10_amt,1 as RI from TSPL_RECEIPT_HEADER where Receipt_No='" & obj.Receipt_No & "'  " & _
            "union all  " & _
            "select sum(tax1_amt) as tax1_amt,sum(tax2_amt) as tax2_amt,sum(tax3_amt) as tax3_amt,sum(tax4_amt) as tax4_amt,sum(tax5_amt) as tax5_amt, " & _
            "sum(tax6_amt) as tax6_amt,sum(tax7_amt) as tax7_amt,sum(tax8_amt) as tax8_amt,sum(tax9_amt) as tax9_amt,sum(tax10_amt) as tax10_amt, " & _
            "-1 as RI from TSPL_RECEIPT_ADVANCE_ADJUSTMENT_KNOCKOFF  where Receipt_no='" & obj.Receipt_No & "' ) a "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            objRec.TAX1_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            objRec.TAX2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            objRec.TAX3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            objRec.TAX4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            objRec.TAX5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
            objRec.TAX6_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
            objRec.TAX7_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
            objRec.TAX8_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
            objRec.TAX9_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))
            objRec.TAX10_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
        End If
        qry = "select sum(Tax1_Amt) as Tax1_Amt,sum(Tax2_Amt) as Tax2_Amt,sum(Tax3_Amt) as Tax3_Amt,sum(Tax4_Amt) as Tax4_Amt,sum(Tax5_Amt) as Tax5_Amt from TSPL_RECEIPT_INVOICEMAPPING_DETAIL where Doc_CODE='" & obj.Doc_Code & "'"
        dt = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            TAX1Invoice_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            TAX2Invoice_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            TAX3Invoice_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            TAX4Invoice_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            TAX5Invoice_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
        End If

        'If obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then ' CASE IF Amount and location same

        'ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then ' CASE Advance Amount > Invoice and location same

        'ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then ' CASE Advance Amount > Invoice and location Different

        'ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount = 0 Then ' CASE Advance Amount < Invoice and location same

        'ElseIf obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1 Then ' CASE Advance Amount = Invoice and location Different and one invoice

        'ElseIf obj.Receipt_Tax_Amt = obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then ' CASE Advance Amount = Invoice and location Different and multiple invoices

        'ElseIf obj.Receipt_Tax_Amt < obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then ' CASE Advance Amount = Invoice and location Different and multiple invoices

        'ElseIf obj.Receipt_Tax_Amt > obj.Invoice_Tax_Amt AndAlso intInvoiceLocCount > 0 Then ' CASE Advance Amount = Invoice and location Different and multiple invoices

        'End If
        Dim isApplyBrachAccounting As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, trans)) = 1, True, False)
        If objRec.Tax_Amount_Advance > 0 Then
            Dim objTM As clsTaxMaster
            If objRec.TAX1_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX1, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX1)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX1)
                    End If

                    'Cases start here
                    If ((objRec.TAX1_Amt = TAX1Invoice_Amt AndAlso intInvoiceLocCount = 0) OrElse (objRec.TAX1_Amt < TAX1Invoice_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX1_Amt}
                        ArryLst.Add(Acc3)
                        TAX1_Amt = objRec.TAX1_Amt
                    ElseIf objRec.TAX1_Amt > TAX1Invoice_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax1_Amt}
                            ArryLst.Add(Acc3)
                            TAX1_Amt += objTR.Tax1_Amt
                        Next
                    ElseIf ((objRec.TAX1_Amt > TAX1Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX1_Amt = TAX1Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX1_Amt = TAX1Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax1_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax1_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc5)

                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objTR.Tax1_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX1_Amt += objTR.Tax1_Amt

                        Next
                    ElseIf (objRec.TAX1_Amt < TAX1Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX1_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX1_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc5)

                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objRec.TAX1_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX1_Amt += objRec.TAX1_Amt
                        Next

                    ElseIf objRec.TAX1_Amt < TAX1Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX2_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax2, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax2)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax2)
                    End If

                    'Cases start here
                    If ((objRec.TAX2_Amt = TAX2Invoice_Amt AndAlso intInvoiceLocCount = 0) OrElse (objRec.TAX2_Amt < TAX2Invoice_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX2_Amt}
                        ArryLst.Add(Acc3)
                        TAX2_Amt = objRec.TAX2_Amt
                    ElseIf objRec.TAX2_Amt > TAX2Invoice_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax2_Amt}
                            ArryLst.Add(Acc3)
                            TAX2_Amt += objTR.Tax2_Amt
                        Next
                    ElseIf ((objRec.TAX2_Amt > TAX2Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX2_Amt = TAX2Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX2_Amt = TAX2Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax2_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax2_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objTR.Tax2_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX2_Amt += objTR.Tax2_Amt
                        Next
                    ElseIf (objRec.TAX2_Amt < TAX2Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX2_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX2_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objRec.TAX2_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX2_Amt += objRec.TAX2_Amt
                        Next
                    ElseIf objRec.TAX2_Amt < TAX2Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX3_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX3, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX3)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX3)
                    End If

                    'Cases start here
                    If ((objRec.TAX3_Amt = TAX3Invoice_Amt AndAlso intInvoiceLocCount = 0) OrElse (objRec.TAX3_Amt < TAX3Invoice_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX3_Amt}
                        ArryLst.Add(Acc3)
                        TAX3_Amt = objRec.TAX3_Amt
                    ElseIf objRec.TAX3_Amt > TAX3Invoice_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax3_Amt}
                            ArryLst.Add(Acc3)
                            TAX3_Amt += objTR.Tax3_Amt
                        Next
                    ElseIf ((objRec.TAX3_Amt > TAX3Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX3_Amt = TAX3Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX3_Amt = TAX3Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax3_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax3_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objTR.Tax3_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX3_Amt += objTR.Tax3_Amt
                        Next
                    ElseIf (objRec.TAX3_Amt < TAX3Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX3_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX3_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objRec.TAX3_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX3_Amt += objRec.TAX3_Amt
                        Next
                    ElseIf objRec.TAX3_Amt < TAX3Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX4_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.TAX4, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.TAX4)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.TAX4)
                    End If

                    'Cases start here
                    If ((objRec.TAX4_Amt = TAX4Invoice_Amt AndAlso intInvoiceLocCount = 0) OrElse (objRec.TAX4_Amt < TAX4Invoice_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX4_Amt}
                        ArryLst.Add(Acc3)
                        TAX4_Amt = objRec.TAX4
                    ElseIf objRec.TAX4_Amt > TAX4Invoice_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax4_Amt}
                            ArryLst.Add(Acc3)
                            TAX4_Amt += objTR.Tax4_Amt
                        Next
                    ElseIf ((objRec.TAX4_Amt > TAX4Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX4_Amt = TAX4Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX4_Amt = TAX4Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax4_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax4_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objTR.Tax4_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX4_Amt += objTR.Tax4_Amt
                        Next
                    ElseIf (objRec.TAX4_Amt < TAX4Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX4_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX4_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objRec.TAX4_Amt}
                                ArryLst.Add(Acc3)
                            End If
                            TAX4_Amt += objRec.TAX4_Amt
                        Next
                    ElseIf objRec.TAX4_Amt < TAX4Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX5_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax5, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax5)
                    End If
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax5)
                    End If

                    'Cases start here
                    If ((objRec.TAX5_Amt = TAX5Invoice_Amt AndAlso intInvoiceLocCount = 0) OrElse (objRec.TAX5_Amt < TAX5Invoice_Amt AndAlso intInvoiceLocCount = 0)) Then
                        '1 CASE  ( IF Amount and location same )  - 
                        '3 CASE (3 CASE Advance Amount < Invoice and location same)
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                        ArryLst.Add(Acc2)
                        Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX5_Amt}
                        ArryLst.Add(Acc3)
                    ElseIf objRec.TAX5_Amt > TAX5Invoice_Amt AndAlso intInvoiceLocCount = 0 Then
                        '2 CASE Advance Amount > Invoice and location same
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                            ArryLst.Add(Acc2)
                            Dim Acc3() As String = {objTM.Tax_Liability_Account, objTR.Tax5_Amt}
                            ArryLst.Add(Acc3)
                        Next
                    ElseIf ((objRec.TAX5_Amt > TAX5Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX5_Amt = TAX5Invoice_Amt AndAlso intInvoiceLocCount > 0) OrElse (objRec.TAX5_Amt = TAX5Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1)) Then
                        '4 CASE( Advance Amount = Invoice Amount and location Different ) 
                        '5 CASE Advance Amount = Invoice and location Different and multiple invoices
                        '7(CASE Advance Amount > Invoice Amount and location Different)


                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objTR.Tax5_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objTR.Tax5_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objTR.Tax5_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf (objRec.TAX5_Amt < TAX5Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount = 1) Then
                        ' CASE( Advance Amount < Invoice Amount and location Different but only one invoice ) 

                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                        For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                            If isApplyBrachAccounting Then
                                Dim strInvoiceLoc = clsERPFuncationality.GetLocationSegment(objTR.InvoiceLocation, trans)

                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, strInvoiceLoc, True, trans)
                                Dim strInvoiceBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(strInvoiceLoc, objRec.Location_GL_Code, trans)
                                If clsCommon.myLen(strInvoiceBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + strInvoiceLoc + " and to location " + objRec.Location_GL_Code)
                                End If
                                Dim strReceiptBranchAcct As String = ClsBranchAccountMapping.GetBranchAccount(objRec.Location_GL_Code, strInvoiceLoc, trans)
                                If clsCommon.myLen(strReceiptBranchAcct) <= 0 Then
                                    Throw New Exception("Please set Branch account mapping with from location " + objRec.Location_GL_Code + " and to location " + strInvoiceLoc)
                                End If
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {strReceiptBranchAcct, objRec.TAX5_Amt}
                                ArryLst.Add(Acc3)
                                Dim Acc4() As String = {objTM.Tax_Liability_Account, objRec.TAX5_Amt}
                                ArryLst.Add(Acc4)
                                Dim Acc5() As String = {strInvoiceBranchAcct, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc5)
                            Else
                                objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                                Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc2)
                                Dim Acc3() As String = {objTM.Tax_Liability_Account, -1 * objRec.TAX5_Amt}
                                ArryLst.Add(Acc3)
                            End If

                        Next
                    ElseIf objRec.TAX5_Amt < TAX5Invoice_Amt AndAlso intInvoiceLocCount > 0 AndAlso intInvoiceCount > 1 Then '6 CASE Advance Amount = Invoice and location Different and multiple invoices
                        Throw New Exception("Multiple Invoice is applied and Total Invoice Tax Amount is greater than Receipt Tax Amount So amount cannot be knock off")

                    End If

                End If
            End If

            If objRec.TAX6_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax6, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax6)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax5)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX6_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX6_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX7_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax7, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax7)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax7)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX7_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX7_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX8_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax8, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax8)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax8)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX8_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX8_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX9_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax9, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax9)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax9)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX9_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX9_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If

            If objRec.TAX10_Amt > 0 Then
                objTM = clsTaxMaster.GetTaxDetailsForSale(objRec.tax10, trans)
                If objTM IsNot Nothing Then
                    If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                        Throw New Exception("Please set Tax Deposit Control Account of Tax Authority " + objRec.tax10)
                    End If
                    objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, objRec.Location_GL_Code, True, trans)
                    If clsCommon.myLen(objTM.Tax_Liability_Account) <= 0 Then
                        Throw New Exception("Please set Tax Liablity Account of Tax Authority " + objRec.tax10)
                    End If
                    objTM.Tax_Liability_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Liability_Account, objRec.Location_GL_Code, True, trans)
                    Dim Acc2() As String = {objTM.DepositControl, -1 * objRec.TAX10_Amt}
                    ArryLst.Add(Acc2)
                    Dim Acc3() As String = {objTM.Tax_Liability_Account, objRec.TAX10_Amt}
                    ArryLst.Add(Acc3)
                End If
            End If
        End If
        If (TAX1_Amt > 0 OrElse TAX2_Amt > 0 OrElse TAX3_Amt > 0 OrElse TAX4_Amt > 0 OrElse TAX5_Amt > 0 OrElse TAX6_Amt > 0 OrElse TAX7_Amt > 0 OrElse TAX8_Amt > 0 OrElse TAX9_Amt > 0 OrElse TAX10_Amt > 0) Then
            qry = "insert  into TSPL_RECEIPT_ADVANCE_ADJUSTMENT_KNOCKOFF(Receipt_No,TAX1_Amt,TAX2_Amt,TAX3_Amt,TAX4_Amt,TAX5_Amt,TAX6_Amt,TAX7_Amt,TAX8_Amt,TAX9_Amt,TAX10_Amt ) " + Environment.NewLine + _
                " values ('" & obj.Receipt_No & "','" & TAX1_Amt & "','" & TAX2_Amt & "','" & TAX3_Amt & "','" & TAX4_Amt & "','" & TAX5_Amt & "','" & TAX6_Amt & "','" & TAX7_Amt & "','" & TAX8_Amt & "','" & TAX9_Amt & "','" & TAX10_Amt & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        End If

        Dim coll As New Hashtable()
        If clsCommon.myLen(objRec.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(objRec.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", objRec.CURRENCY_CODE)
            clsCommon.AddColumnsForChange(coll, "ConvRate", objRec.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ConvRateOld", objRec.ConvRate)
        End If
        clsJournalMaster.FunGrnlEntryWithTrans(objRec.Location_GL_Code, False, trans, clsCommon.myCDate(obj.Document_Date), obj.Description, "RC-AD", "Receipt Advance Tax Knock off", obj.Doc_Code, obj.Description, "C", objRec.Cust_Code, objRec.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, objRec.Reference, strRemarks, Nothing, coll)

        Return True

    End Function
    Public Shared Function PostData(ByVal strDocNo As String, ByVal isCheckForPosted As Boolean) As Boolean
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Code not found to Post")
        End If
        Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt")
        Dim obj As clsReceiptInvoiceHead = clsReceiptInvoiceHead.GetData(strDocNo, NavigatorType.Current)

        If (obj Is Nothing OrElse clsCommon.myLen(obj.Doc_Code) <= 0) Then
            Throw New Exception("No Data found to Post")
        End If
        If (isCheckForPosted AndAlso obj.POSTED = 1) Then
            Throw New Exception("Already Post on :")
        End If
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim qry As String = "Update TSPL_Receipt_InvoiceMapping_Head set isPOSTED=1, Modified_Date='" + strPostDate + "',Modified_By='" + objCommonVar.CurrentUserCode + "' where DOC_CODE ='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                For Each objTr As clsReceiptInvoiceDetails In obj.Arr
                    qry = "Update tspl_sd_sale_invoice_head set IsAdvanceTaxGlEntry=1  where Document_Code ='" + objTr.InvoiceNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                Next
            End If
            CreateGLEntryForAllCases(obj, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function ReverseAndUnpost(ByVal strCode As String) As Boolean
        Dim obj As clsReceiptInvoiceHead = clsReceiptInvoiceHead.GetData(strCode, NavigatorType.Current)
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select isPosted from TSPL_RECEIPT_INVOICEMAPPING_HEAD where Doc_Code='" + strCode + "'"
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='RC-AD' and Source_Doc_No='" + strCode + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            For Each objTR As clsReceiptInvoiceDetails In obj.Arr
                Qry = "update TSPL_SD_SALE_INVOICE_HEAD set IsAdvanceTaxGlEntry=0  where Document_Code='" & objTR.InvoiceNo & "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            Next

            Qry = "update TSPL_RECEIPT_INVOICEMAPPING_HEAD set isPosted=0 where Doc_Code='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetReceiptBalanceTaxAmount(ByVal strReceiptNo As String, ByVal strCurrDocNo As String, ByVal trans As SqlTransaction) As Decimal
        Dim qry As String = "select sum(Tax_Amount_Advance*RI) as BalanceAmt from (" + Environment.NewLine + _
        "Select TSPL_RECEIPT_HEADER.Receipt_No,TSPL_RECEIPT_HEADER.Tax_Amount_Advance,1 as RI from  TSPL_RECEIPT_HEADER WHERE Receipt_No='" + strReceiptNo + "' " + Environment.NewLine + _
        "union all" + Environment.NewLine + _
        "select TSPL_Receipt_InvoiceMapping_Head.Receipt_No,TSPL_Receipt_InvoiceMapping_detail.Total_Tax_Amt,-1 as RI from TSPL_Receipt_InvoiceMapping_detail" + Environment.NewLine + _
        "left outer join TSPL_Receipt_InvoiceMapping_Head on TSPL_Receipt_InvoiceMapping_Head.Doc_Code=TSPL_Receipt_InvoiceMapping_detail.Doc_CODE" + Environment.NewLine + _
         "where TSPL_Receipt_InvoiceMapping_Head.Receipt_No='" + strReceiptNo + "'  and TSPL_Receipt_InvoiceMapping_Head.Doc_CODE not in ('" + strCurrDocNo + "' )" + Environment.NewLine + _
         ")x group by Receipt_No"
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
    End Function
End Class
Public Class clsReceiptInvoiceDetails
#Region "Variables"
    Public Line_No As String = Nothing
    Public Doc_Code As String = Nothing
    Public InvoiceNo As String = Nothing
    Public Remarks As String = Nothing
    Public InvoiceLocation As String = Nothing
    Public Total_Tax_Amt As Double = 0
    Public Tax1_Amt As Double = 0
    Public Tax2_Amt As Double = 0
    Public Tax3_Amt As Double = 0
    Public Tax4_Amt As Double = 0
    Public Tax5_Amt As Double = 0
    Public Tax6_Amt As Double = 0
    Public Tax7_Amt As Double = 0
    Public Tax8_Amt As Double = 0
    Public Tax9_Amt As Double = 0
    Public Tax10_Amt As Double = 0
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsReceiptInvoiceDetails), ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean

        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each objtr As clsReceiptInvoiceDetails In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Doc_Code", strDocNo)
                clsCommon.AddColumnsForChange(coll, "InvoiceNo", objtr.InvoiceNo)
                clsCommon.AddColumnsForChange(coll, "Line_No", objtr.Line_No)
                clsCommon.AddColumnsForChange(coll, "Remarks", objtr.Remarks)
                clsCommon.AddColumnsForChange(coll, "InvoiceLocation", objtr.InvoiceLocation)
                clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", objtr.Total_Tax_Amt)

                Dim qry = "select Total_Tax_Amt, TAX1_Amt,TAX2_Amt,TAX3_Amt,TAX4_Amt,TAX5_Amt,TAX6_Amt,TAX7_Amt,TAX8_Amt,TAX9_Amt,TAX10_Amt from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & objtr.InvoiceNo & "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt")) > 0 Then
                        objtr.Tax1_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax2_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax3_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax4_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax5_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax6_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax7_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax8_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax9_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                        objtr.Tax10_Amt = objtr.Total_Tax_Amt * clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt")) / clsCommon.myCdbl(dt.Rows(0)("Total_Tax_Amt"))
                    End If
                End If

                clsCommon.AddColumnsForChange(coll, "Tax1_Amt", objtr.Tax1_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax2_Amt", objtr.Tax2_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax3_Amt", objtr.Tax3_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax4_Amt", objtr.Tax4_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax5_Amt", objtr.Tax5_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax6_Amt", objtr.Tax6_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax7_Amt", objtr.Tax7_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax8_Amt", objtr.Tax8_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax9_Amt", objtr.Tax9_Amt)
                clsCommon.AddColumnsForChange(coll, "Tax10_Amt", objtr.Tax10_Amt)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Receipt_InvoiceMapping_Detail", OMInsertOrUpdate.Insert, "", trans)

                'qry = "Update TSPL_Receipt_InvoiceMapping_Detail set Tax1_Amt='" & objtr.Tax1_Amt & "', " & _
                '    "Tax2_Amt='" & objtr.Tax2_Amt & "', Tax3_Amt='" & objtr.Tax3_Amt & "', Tax4_Amt='" & objtr.Tax4_Amt & "',  " & _
                '    "Tax5_Amt='" & objtr.Tax5_Amt & "', Tax6_Amt='" & objtr.Tax6_Amt & "', Tax7_Amt='" & objtr.Tax7_Amt & "',  " & _
                '    "Tax8_Amt='" & objtr.Tax8_Amt & "', Tax9_Amt='" & objtr.Tax9_Amt & "', Tax10_Amt='" & objtr.Tax10_Amt & "' where InvoiceNo='" & objtr.InvoiceNo & "'"
                'clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Next
        End If
        Return True
    End Function
End Class
