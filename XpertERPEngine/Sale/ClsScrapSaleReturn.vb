Imports System.Data.SqlClient

Public Class ClsScrapSaleHeadReturn
#Region "Variables"
    Public Document_No As String = Nothing
    Public shipment_No As String = Nothing
    Public Specification As String = Nothing
    Public invoice_No As String = Nothing
    Public Status As String = Nothing
    Public Po_No As String = Nothing
    Public NRG_No As String = Nothing
    Public cust_Code As String = Nothing
    Public cust_Name As String = Nothing
    Public shipment_Date As Date
    Public posting_Date As String = Nothing
    Public expship_Date As String = Nothing
    Public Loc_Code As String = Nothing
    Public Loc_Name As String = Nothing
    Public Vehicle_Id As String = Nothing
    Public ToLoc_Code As String = Nothing
    Public CreateInvoice As String = Nothing
    Public Excisable As String = Nothing
    Public Description As String = Nothing
    Public reff As String = Nothing
    Public Tax_Group As String = Nothing
    Public Tax_Desc As String = Nothing
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
    Public Addcost As String = Nothing
    Public AddcostDesc As String = Nothing
    Public Add_Amt As Double = 0
    Public Before_Add_Amt As Double = 0
    Public After_Add_Amt As Double = 0
    Public Discount_Base As Double = 0
    Public Discount_Amt As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public Total_Tax_Amt As Double = 0
    Public ship_Total_Amt As Double = 0
    Public doc_Amt As Double = 0
    Public Comp_Code As String = Nothing
    Public ispost As Integer = 0

    Public AddCode1 As String = Nothing
    Public AddDesc1 As String = Nothing
    Public AddAmt1 As Double = 0
    Public AddCode2 As String = Nothing
    Public AddDesc2 As String = Nothing
    Public AddAmt2 As Double = 0
    Public AddCode3 As String = Nothing
    Public AddDesc3 As String = Nothing
    Public AddAmt3 As Double = 0
    Public AddCode4 As String = Nothing
    Public AddDesc4 As String = Nothing
    Public AddAmt4 As Double = 0
    Public AddCode5 As String = Nothing
    Public AddDesc5 As String = Nothing
    Public AddAmt5 As Double = 0
    Public AddCode6 As String = Nothing
    Public AddDesc6 As String = Nothing
    Public AddAmt6 As Double = 0
    Public AddCode7 As String = Nothing
    Public AddDesc7 As String = Nothing
    Public AddAmt7 As Double = 0
    Public AddCode8 As String = Nothing
    Public AddDesc8 As String = Nothing
    Public AddAmt8 As Double = 0
    Public AddCode9 As String = Nothing
    Public AddDesc9 As String = Nothing
    Public AddAmt9 As Double = 0
    Public AddCode10 As String = Nothing
    Public AddDesc10 As String = Nothing
    Public AddAmt10 As Double = 0
    Public Terms_Code As String = Nothing
    Public Due_Date As String = Nothing
    Public Unit_Code As String = Nothing

    Public Inter_Branch As Boolean = False
    Public Tax_Calculation_Type As EnumTaxCalucationType
    Public strInvoiceNo As String = Nothing
    Public Invoice_Type As String = Nothing
    'Public ispost As ERPTransactionStatus = ERPTransactionStatus.Pending
    Public is_Asset_Type As Boolean = False
    Public Arr1 As List(Of ClsScrapSaleDetailReturn) = Nothing
    Public Transporter_code As String = Nothing
    Public Transporter_Name As String = Nothing
    Public Vehicle_code As String = Nothing
    Public VAT_InvoiceNo As String = Nothing
    Public VatInvoice_Type As String = Nothing
    Public Is_Scrap As String = "N"
    Public Is_CashSale As String = "N"
    Public InvoiceManualNowithPrefix As String = Nothing
    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing
    Public RoundOffAmount As Double = 0

    Public Total_Gross_Weight As Double = 0
    Public Total_Net_Weight As Double = 0
    Public Doc_Type As String
    Public Is_Taxable As Boolean = False
    Public Is_ManualTCS As Integer = 0
    Public EWayBillNo As String
    Public EWayBillDate As Date? = Nothing
    Public Electronic_Ref_No As String
    Public Is_Cancelled As Integer = 0
    Public Inter_Unit_Sale As Integer = 0
    Public IsEinvoice As Integer = 0
    Public IsEwayBill As Integer = 0
    Public Gate_Entry_No As String = Nothing
#End Region
    Public Shared Function frmScrapSaleReturnPrint(ByVal Form_ID As String, ByVal strCancelDelete As String, ByVal strDate As DateTime, ByVal StrDocNo As String, ByVal isPrint As Boolean, ByVal location As String) As Boolean
        Dim qry As String = Nothing
        Dim TSPL_SCRAPSALE_HEAD_RETURN As String = Nothing
        Dim TSPL_SCRAPSALE_DETAIL_RETURN As String = Nothing
        If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal Then
            TSPL_SCRAPSALE_HEAD_RETURN = "TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data"
            TSPL_SCRAPSALE_DETAIL_RETURN = "TSPL_SCRAPSALE_DETAIL_RETURN_cancel_data"
        ElseIf clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
            TSPL_SCRAPSALE_HEAD_RETURN = "TSPL_SCRAPSALE_HEAD_RETURN_Delete_data"
            TSPL_SCRAPSALE_DETAIL_RETURN = "TSPL_SCRAPSALE_DETAIL_RETURN_Delete_data"
        Else
            TSPL_SCRAPSALE_HEAD_RETURN = "TSPL_SCRAPSALE_HEAD_RETURN"
            TSPL_SCRAPSALE_DETAIL_RETURN = "TSPL_SCRAPSALE_DETAIL_RETURN"
        End If
        Dim InvoiceNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select invoice_No from TSPL_SCRAPINVOICE_HEAD where shipment_No in (select shipment_No from TSPL_SCRAPSALE_HEAD_RETURN where Document_No='" + StrDocNo + "')"))
        Dim Address As String

        If clsCommon.myLen(location) > 0 Then
            Address = "(Select  MAX(TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end  + Case When TSPL_LOCATION_MASTER.Add4 ='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add4 ,103) end) from TSPL_LOCATION_MASTER   where Location_Code = ('" + location + "'))  "
        Else
            Address = "(TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end) "

        End If


        Dim dt As DataTable
        qry = "select h.Description,h.Reff," + TSPL_SCRAPSALE_HEAD_RETURN + ".Vehicle_Id,TSPL_GL_SEGMENT_CODE.Description as Vdescription, h.Created_By ,h.Modify_By , tspl_company_master.logo_img,tspl_company_master.logo_img2,tspl_company_master.Comp_Name  as CompanyName,tspl_company_master.Tin_No as CompanyTin," + Address + " as  CompanyAddress,h.invoice_No as ScrapInvoice,h.shipment_Date  as ScrapInvoiceDate," &
                                   "h.cust_Name as CustomerName,cm.PAN as Cust_PAN,h.cust_Code as CustomerCode,case when len(cm.Add1)>0 then cm.Add1 else '' end +" &
                                   "case when len(cm.Add2)>0 then ',' else '' end + " &
                                   "case when len(cm.Add2)>0 then cm.Add2 else ''end +" &
                                   "case when len(cm.Add3)>0 then ',' else '' end +" &
                                   "case when len(cm.Add3)>0 then cm.Add3 else '' end   as CustomerAddress," &
                                   "" + TSPL_SCRAPSALE_HEAD_RETURN + ".NRG_No  as NrgNo,cm.CST as CstNo,TSPL_LOCATION_MASTER.Tin_No as TinNo,d.Item_Code as ItemCode," &
                                   " (Select ISNULL(Cheapter_Heads,'') from TSPL_ITEM_MASTER WHere TSPL_ITEM_MASTER.Item_Code=d.Item_Code)as [ChapterHead], " &
                                   "d.Item_Desc as Desciption,d.invoice_Qty as Quantity ,d.unit_code as Uom,d.price as Rate," &
                                   "d.ItemAmt as Amount, d.TAX1_Amt as Dtax1_Amt, d.DiscountAmt ,h.TAX1 as TaxRateDesc1, ROUND(h.TAX1_Amt, 0) as TaxRate1,h.TAX2 as TaxRateDesc2,ROUND(h.TAX2_Amt,0) as TaxRate2,h.TAX3 as TaxRateDesc3 ,ROUND(h.TAX3_Amt,0) as TaxRate3,h.TAX4 as TaxRateDesc4,ROUND(h.TAX4_Amt,0) as TaxRate4,h.TAX5 as TaxRateDesc5,ROUND(h.TAX5_Amt,0) as TaxRate5,h.TAX6 as TaxRateDesc6,ROUND(h.TAX6_Amt,0) as TaxRate6,h.TAX7 as TaxRateDesc7,ROUND(h.TAX7_Amt,0) as TaxRate7,h.TAX8 as  TaxRateDesc8,ROUND(h.TAX8_Amt,0) as TaxRate8,h.TAX9 as TaxRateDesc9,ROUND(h.TAX9_Amt,0) as TaxRate9,h.TAX10 as TaxRateDesc10,ROUND(h.TAX10_Amt,0) as  TaxRate10, " &
                                   " h.TAX1_Rate , h.TAX2_Rate, h.TAX3_Rate, h.TAX4_Rate, h.TAX5_Rate, h.TAX6_Rate, h.TAX7_Rate, h.TAX8_Rate, h.TAX9_Rate, h.TAX10_Rate, " &
                                   " " + TSPL_SCRAPSALE_HEAD_RETURN + ".Excisable " &
                                   "from TSPL_SCRAPINVOICE_HEAD h left outer join TSPL_SCRAPINVOICE_DETAIL d on h.invoice_No =d.invoice_No " &
                                   "left outer join TSPL_CUSTOMER_MASTER   cm on h.cust_Code =cm.Cust_Code  left outer join tspl_company_master   on tspl_company_master.Comp_Code ='" + objCommonVar.CurrentCompanyCode + "'left outer join " + TSPL_SCRAPSALE_HEAD_RETURN + "  on h.shipment_No =" + TSPL_SCRAPSALE_HEAD_RETURN + ".shipment_No left outer join TSPL_LOCATION_MASTER on  h.Loc_Code=TSPL_LOCATION_MASTER.Location_Code left outer join TSPL_GL_SEGMENT_CODE on TSPL_GL_SEGMENT_CODE.Segment_code  =" + TSPL_SCRAPSALE_HEAD_RETURN + ".Vehicle_Id where h.invoice_No='" + clsCommon.myCstr(InvoiceNo) + "'"
        dt = clsDBFuncationality.GetDataTable(qry)
        '====for KDIL(by shivani)==============='
        Dim Qry1 As String = "select "
        If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
            Qry1 += " ' Cancelled' As Report_Status, "
        Else
            Qry1 += " '' As Report_Status, "
        End If

        Qry1 += "  " + TSPL_SCRAPSALE_HEAD_RETURN + ".Is_Taxable,tspl_company_master.comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1, TSPL_COMPANY_MASTER.Add2 as Comp_Add2, TSPL_COMPANY_MASTER.Add3 as Comp_Add3, TSPL_COMPANY_MASTER.Email as Comp_Email, TSPL_COMPANY_MASTER.Phone1 as Comp_Phone1, TSPL_COMPANY_MASTER.Phone2 as Comp_Phone2, TSPL_COMPANY_MASTER.Pan_No as Comp_Pan_No, cast(TSPL_COMPANY_MASTER.Logo_Img as image) as Logo_Img, cast(TSPL_COMPANY_MASTER.Logo_Img2 as image) as Logo_Img2 , TSPL_COMPANY_MASTER.GSTREg_No as Comp_GSTREg_No , TSPL_COMPANY_MASTER.CINNO as Comp_CINNO, TSPL_COMPANY_MASTER.Access_Officer as Comp_Access_Officer," + TSPL_SCRAPSALE_DETAIL_RETURN + ".TotalAmt,isnull(" + TSPL_SCRAPSALE_HEAD_RETURN + ".RoundOffAmount,0) as RoundOffAmount,cast(" + TSPL_SCRAPSALE_HEAD_RETURN + ".BarCode_Img as image) As BarCode_Img,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".IRN_No,'') as IRN_No,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".Ack_No,'') as Ack_No,case when len(isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".Ack_No,'')) > 0 then convert (varchar, " + TSPL_SCRAPSALE_HEAD_RETURN + ".Ack_Date,103) else ''  end as Ack_Date,TSPL_SCRAPINVOICE_HEAD.EWayBillNo,convert(varchar,TSPL_SCRAPINVOICE_HEAD.EWayBillDate,103) as EWayBillDate,'' as Electronic_Ref_No, TSPL_ITEM_MASTER.HSN_Code, " &
                                            "FromState.GST_STATE_CODE as From_GstStateCode,FromLocation.GSTNO as From_Loc_GstinNo,Customer_State.GST_STATE_CODE as Cust_GstStateCode,Customer_State.STATE_NAME AS Cust_StateName,TSPL_CUSTOMER_MASTER.GSTNO as Cust_GstInNo,ISNULL(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".DiscountAmt,0) AS DiscountAmt," &
                                            " convert(varchar,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) as shipment_Date, tspl_customer_master.PAN as Cust_Pan, TSPL_SCRAPINVOICE_HEAD.Description,FromLocation.HOAdd1 as frmHO1,FromLocation.HOAdd2 as frmHO2,ToLocation.HOAdd1 as to_HO1,ToLocation.HOAdd2 as to_HO2," &
                                        "TSPL_SCRAPINVOICE_HEAD.invoice_No," + TSPL_SCRAPSALE_HEAD_RETURN + ".Document_No,convert(varchar," + TSPL_SCRAPSALE_HEAD_RETURN + ".Return_ship_Date,103) as Return_Date,convert(varchar,TSPL_SCRAPINVOICE_HEAD.posting_Date,103) as Invoice_Date," + TSPL_SCRAPSALE_HEAD_RETURN + ".Loc_Code as From_Location,FromLocation.Location_Desc as [From Location Desc],(FromLocation.Add1+FromLocation.Add2+FromLocation.Add3+FromLocation.Add4)as [From Address] ," &
                                        "FromLocation.Pin_Code,FromLocation.TIN_No,FromLocation.CST_No,FromLocation.State as From_State,FromState.State_Name as frm_State_name  ," + TSPL_SCRAPSALE_HEAD_RETURN + ".ToLoc_Code as To_Location,ToLocation.Location_Desc as To_Location_Desc,(ToLocation.Add1+ToLocation.Add2+ToLocation.Add3+ToLocation.Add4)as [To Address]," &
                                        "tspl_customer_master.Pin_Code as [To Pin Code],tspl_customer_master.TIN_No as [To TIN No],tspl_customer_master.CST as [To CST No],tspl_customer_master.Phone1 as [To phone],ToLocation.State as To_State,ToState.State_Name as To_state_name," + TSPL_SCRAPSALE_DETAIL_RETURN + ".Item_Code,TSPL_ITEM_MASTER.Item_Desc,Shipped_Qty," &
                                        "" + TSPL_SCRAPSALE_DETAIL_RETURN + ".Unit_Code,Price,ItemAmt," + TSPL_SCRAPSALE_HEAD_RETURN + ".Invoice_Type," + TSPL_SCRAPSALE_HEAD_RETURN + ".Total_Tax_Amt," + TSPL_SCRAPSALE_HEAD_RETURN + ".Doc_Amt,	TAX1 .Tax_Code as tax1name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax1_rate,0) as txt1Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax1_amt,0) as txt1amt, "
        Qry1 += " tax2.Tax_Code as tax2name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax2_rate,0) as txt2Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax2_amt,0) as txt2amt,"
        Qry1 += " tax3.Tax_Code as tax3name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax3_rate,0) as txt3Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax3_amt,0) as txt3amt, "
        Qry1 += " tax4.Tax_Code as tax4name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax4_rate,0) as txt4Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax4_amt,0) as txt4amt, "
        Qry1 += " tax5.Tax_Code as tax5name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax5_rate,0) as txt5Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax5_amt,0) as txt5amt, "
        Qry1 += "  tax6.Tax_Code as tax6name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax6_rate,0) as txt6Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax6_amt,0) as txt6amt, "
        Qry1 += "  tax7.Tax_Code as tax7name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax7_rate,0) as txt7Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax7_amt,0) as txt7amt, "
        Qry1 += " tax8.Tax_Code as tax8name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax8_rate,0) as txt8Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax8_amt,0) as txt8amt,  "
        Qry1 += "  tax9.Tax_Code as tax9name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax9_rate,0) as txt9Rate,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax9_amt,0) as txt9amt,'' as cin,'' as pan, '' as companyaddress,"
        Qry1 += " tax10.Tax_Code as tax10name,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax10_amt,0) as txt10amt ," + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc1,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt1,0) as AddAmt1,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc2,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt2,0) as AddAmt2,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc3,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt3,0) as AddAmt3,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc4,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt4,0) as AddAmt4,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc5,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt5,0) as AddAmt5,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc6,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt6,0) as AddAmt6,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc7,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt7,0) as AddAmt7,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc8,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt8,0) as AddAmt8,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc9,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt9,0) as AddAmt9,"
        Qry1 += " " + TSPL_SCRAPSALE_HEAD_RETURN + ".AddDesc10,isnull (" + TSPL_SCRAPSALE_HEAD_RETURN + ".AddAmt10,0) as AddAmt10,"

        Qry1 += "dtax1.Type as tax1Type,dtax2.Type as tax2Type,dtax3.Type as tax3Type,dtax4.Type as tax4Type,dtax5.Type as tax5Type,dtax6.Type as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX1_Amt ,0) as DTax1_Amt, isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX2_Amt ,0) as DTax2_Amt," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX3_Amt ,0) as DTax3_Amt, isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX4_Amt ,0) as DTax4_Amt," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX5_Amt ,0) as DTax5_Amt, isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX6_Amt ,0) as DTax6_Amt," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX7_Amt ,0) as DTax7_Amt, isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX8_Amt ,0) as DTax8_Amt," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX9_Amt ,0) as DTax9_Amt, isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX10_Amt ,0) as DTax10_Amt," &
                         " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX1_Rate,0) as DTax1_Rate,  isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX2_Rate,0) as DTax2_Rate, " &
                          " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX3_Rate,0) as DTax3_Rate,  isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX4_Rate,0) as DTax4_Rate," &
                           " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX5_Rate,0) as DTax5_Rate,  isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX6_Rate,0) as DTax6_Rate," &
                            " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX7_Rate,0) as DTax7_Rate,  isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX8_Rate,0) as DTax8_Rate," &
                             " isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX9_Rate,0) as DTax9_Rate,  isnull(" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX10_Rate,0) as DTax10_Rate,"

        Qry1 += "" + TSPL_SCRAPSALE_HEAD_RETURN + ".vehicle_code," + TSPL_SCRAPSALE_HEAD_RETURN + ".Transport_code,Transporter_name," + TSPL_SCRAPSALE_HEAD_RETURN + ".Cust_code," + TSPL_SCRAPSALE_HEAD_RETURN + ".cust_name,tspl_customer_master.add1,tspl_customer_master.add2,tspl_customer_master.add3,tspl_city_master.city_name," + TSPL_SCRAPSALE_HEAD_RETURN + ".Electronic_ref_No from " + TSPL_SCRAPSALE_HEAD_RETURN + " left join " + TSPL_SCRAPSALE_DETAIL_RETURN + " on " + TSPL_SCRAPSALE_DETAIL_RETURN + ".shipment_No=" + TSPL_SCRAPSALE_HEAD_RETURN + ".shipment_No left join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD. shipment_No=" + TSPL_SCRAPSALE_HEAD_RETURN + ".shipment_No left join TSPL_LOCATION_MASTER as FromLocation on FromLocation.Location_Code=" + TSPL_SCRAPSALE_HEAD_RETURN + ".Loc_Code left join TSPL_LOCATION_MASTER as ToLocation on ToLocation.Location_Code=" + TSPL_SCRAPSALE_HEAD_RETURN + ".ToLoc_Code left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=" + TSPL_SCRAPSALE_DETAIL_RETURN + ".Item_Code left join TSPL_STATE_MASTER as  FromState on FromState.State_Code=FromLocation.State"
        Qry1 += " left join TSPL_STATE_MASTER as  ToState on ToState.State_Code=ToLocation.State"
        Qry1 += " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =" + TSPL_SCRAPSALE_HEAD_RETURN + ".tax1 "
        Qry1 += "  left outer join tspl_tax_master as tax2 on tax2.tax_code = " + TSPL_SCRAPSALE_HEAD_RETURN + ".tax2 "
        Qry1 += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX3 "
        Qry1 += "  left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= " + TSPL_SCRAPSALE_HEAD_RETURN + " .tax4 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=" + TSPL_SCRAPSALE_HEAD_RETURN + " .tax5 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX6 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX7 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX8 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX9 "
        Qry1 += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =" + TSPL_SCRAPSALE_HEAD_RETURN + " .TAX10 "

        Qry1 += "left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =" + TSPL_SCRAPSALE_DETAIL_RETURN + ".tax1  " &
                     " left outer join tspl_tax_master as dtax2 on dtax2.tax_code = " + TSPL_SCRAPSALE_DETAIL_RETURN + ".tax2 " &
                      "left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX3  " &
                       " left outer join TSPL_TAX_MASTER as dtax4 on dtax4.Tax_Code= " + TSPL_SCRAPSALE_DETAIL_RETURN + ".tax4 " &
                        "left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=" + TSPL_SCRAPSALE_DETAIL_RETURN + ".tax5 " &
                         " left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX6  " &
                         "left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX7  " &
                         " left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code = " + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX8  " &
                         " left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX9 " &
                          " left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =" + TSPL_SCRAPSALE_DETAIL_RETURN + ".TAX10 " &
                    " left join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.Transport_Id=" + TSPL_SCRAPSALE_HEAD_RETURN + ".Transport_code  left join tspl_customer_master on tspl_customer_master.cust_code=" + TSPL_SCRAPSALE_HEAD_RETURN + ".cust_code "
        Qry1 += " left join TSPL_STATE_MASTER as Customer_State on tspl_customer_master .State=Customer_State.STATE_CODE "
        Qry1 += " left join tspl_city_master on tspl_city_master.city_code=tspl_customer_master.city_code "
        Qry1 += "left outer join tspl_company_master on tspl_company_master.comp_code = " + TSPL_SCRAPSALE_HEAD_RETURN + ".Comp_Code where " + TSPL_SCRAPSALE_HEAD_RETURN + ".Document_No='" + StrDocNo + "'"

        Dim dt1 As DataTable
        dt1 = clsDBFuncationality.GetDataTable(Qry1)

        If isPrint Then
            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDFCF") = CompairStringResult.Equal Then
                If clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(dt1.Rows(0)("Invoice_Date"))) Then
                    Dim frmCRV As New frmCrystalReportViewer()

                    If clsCommon.myCdbl(dt1.Rows(0)("Is_Taxable")) = 1 Then
                        frmCRV.funreport(Form_ID, CrystalReportFolder.KwalitySalesReport, dt1, "rptMaterialSalesReturn_RCDF", "")
                    Else
                        frmCRV.funreport(Form_ID, CrystalReportFolder.KwalitySalesReport, dt1, "rptMaterialSaleaReturn_RCDFCF_NT", "")

                    End If
                    frmCRV = Nothing

                End If
            End If

        Else
            'frmCRV.funreport(CrystalReportFolder.PurchaseOrder, dt, EnumTecxpertPaperSize.PaperSize10x12, "ScrapSaleInvoice4Excise", "ScrapSaleInvoiceRpt")
        End If
        Return True
    End Function
    Public Shared Function HistoryUpdate(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_SCRAPSALE_HEAD_Return", "Document_No", "TSPL_SCRAPSALE_DETAIL_RETURN", "Document_No", trans)
        Return True
    End Function

    Public Function SaveData(ByVal obj As ClsScrapSaleHeadReturn, ByVal strScrapSaleInvoiceNo As String, ByVal isNewEntry As Boolean) As Boolean
        Dim isSaved As Boolean = True
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Dim Desc As String = String.Empty
        Dim VatInvoiceType As String = Nothing
        Try
            'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase Order", "Scrap Invoice", obj.Loc_Code, obj.shipment_Date, trans)
            'If Not isNewEntry Then
            '    HistoryUpdate(obj.Document_No, trans)
            'End If
            Dim qry As String = "delete from TSPL_SCRAPSALE_DETAIL_RETURN where Document_No='" + obj.Document_No + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_SCRAPINVOICE_DETAIL where invoice_No='" + obj.strInvoiceNo + "'"
            'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_SCRAPINVOICE_HEAD where invoice_No='" + obj.strInvoiceNo + "'"
            'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

            clsBatchInventory.DeleteData("MS-SR", obj.Document_No, trans)

            Dim strDocNo As String = ""
            If isNewEntry Then
                If obj.Is_Cancelled = 1 Then
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.shipment_Date, clsDocType.ScrapReturn, clsDocTransactionType.SaleReturnCancel, obj.Loc_Code)
                Else
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, obj.shipment_Date, clsDocType.ScrapReturn, clsDocTransactionType.NA, obj.Loc_Code)
                End If
            End If

            If (clsCommon.myLen(obj.Document_No) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If
            '-----------------------------
            Dim coll As New Hashtable()

            clsCommon.AddColumnsForChange(coll, "shipment_No", obj.shipment_No)
            obj.invoice_No = clsDBFuncationality.getSingleValue("select invoice_No from TSPL_SCRAPINVOICE_HEAD where TSPL_SCRAPINVOICE_HEAD.shipment_No='" & obj.shipment_No & "'", trans)
            clsCommon.AddColumnsForChange(coll, "invoice_No", obj.invoice_No)
            clsCommon.AddColumnsForChange(coll, "Doc_Type", obj.Doc_Type)
            clsCommon.AddColumnsForChange(coll, "Status", obj.Status)
            clsCommon.AddColumnsForChange(coll, "Po_No", obj.Po_No)
            clsCommon.AddColumnsForChange(coll, "NRG_No", obj.NRG_No)
            clsCommon.AddColumnsForChange(coll, "cust_Code", obj.cust_Code)
            clsCommon.AddColumnsForChange(coll, "cust_Name", obj.cust_Name)
            clsCommon.AddColumnsForChange(coll, "Return_ship_Date", clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt "))
            clsCommon.AddColumnsForChange(coll, "posting_Date", clsCommon.GetPrintDate(obj.posting_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "expship_Date", clsCommon.GetPrintDate(obj.expship_Date, "dd/MMM/yyyy hh:mm tt"))
            clsCommon.AddColumnsForChange(coll, "Loc_Code", obj.Loc_Code)
            clsCommon.AddColumnsForChange(coll, "Loc_Name", obj.Loc_Name)
            clsCommon.AddColumnsForChange(coll, "Vehicle_Id", obj.Vehicle_Id)
            clsCommon.AddColumnsForChange(coll, "ToLoc_Code", obj.ToLoc_Code)
            clsCommon.AddColumnsForChange(coll, "CreateInvoice", obj.CreateInvoice)
            clsCommon.AddColumnsForChange(coll, "Excisable", obj.Excisable)
            clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
            '============Add By Rohit on Apr,14-2015==============
            clsCommon.AddColumnsForChange(coll, "Transport_code", obj.Transporter_code)
            clsCommon.AddColumnsForChange(coll, "Vehicle_code", obj.Vehicle_code)
            '===========================================================
            clsCommon.AddColumnsForChange(coll, "reff", obj.reff)
            clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
            clsCommon.AddColumnsForChange(coll, "Tax_Desc", obj.Tax_Desc)
            clsCommon.AddColumnsForChange(coll, "Add_Amt", obj.Add_Amt)
            clsCommon.AddColumnsForChange(coll, "Before_Add_Amt", obj.Before_Add_Amt)
            clsCommon.AddColumnsForChange(coll, "Discount_Base", obj.Discount_Base)
            clsCommon.AddColumnsForChange(coll, "Discount_Amt", obj.Discount_Amt)
            clsCommon.AddColumnsForChange(coll, "Amount_Less_Discount", obj.Amount_Less_Discount)
            clsCommon.AddColumnsForChange(coll, "Total_Tax_Amt", obj.Total_Tax_Amt)
            clsCommon.AddColumnsForChange(coll, "ship_Total_Amt", obj.ship_Total_Amt)
            clsCommon.AddColumnsForChange(coll, "doc_Amt", obj.doc_Amt)
            clsCommon.AddColumnsForChange(coll, "RoundOffAmount", obj.RoundOffAmount)
            clsCommon.AddColumnsForChange(coll, "ispost", 0)
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
            clsCommon.AddColumnsForChange(coll, "AddCode1", obj.AddCode1)
            clsCommon.AddColumnsForChange(coll, "AddDesc1", obj.AddDesc1)
            clsCommon.AddColumnsForChange(coll, "AddAmt1", obj.AddAmt1)
            clsCommon.AddColumnsForChange(coll, "AddCode2", obj.AddCode2)
            clsCommon.AddColumnsForChange(coll, "AddDesc2", obj.AddDesc2)
            clsCommon.AddColumnsForChange(coll, "AddAmt2", obj.AddAmt2)
            clsCommon.AddColumnsForChange(coll, "AddCode3", obj.AddCode3)
            clsCommon.AddColumnsForChange(coll, "AddDesc3", obj.AddDesc3)
            clsCommon.AddColumnsForChange(coll, "AddAmt3", obj.AddAmt3)
            clsCommon.AddColumnsForChange(coll, "AddCode4", obj.AddCode4)
            clsCommon.AddColumnsForChange(coll, "AddDesc4", obj.AddDesc4)
            clsCommon.AddColumnsForChange(coll, "AddAmt4", obj.AddAmt4)
            clsCommon.AddColumnsForChange(coll, "AddCode5", obj.AddCode5)
            clsCommon.AddColumnsForChange(coll, "AddDesc5", obj.AddDesc5)
            clsCommon.AddColumnsForChange(coll, "AddAmt5", obj.AddAmt5)
            clsCommon.AddColumnsForChange(coll, "AddCode6", obj.AddCode6)
            clsCommon.AddColumnsForChange(coll, "AddDesc6", obj.AddDesc6)
            clsCommon.AddColumnsForChange(coll, "AddAmt6", obj.AddAmt6)
            clsCommon.AddColumnsForChange(coll, "AddCode7", obj.AddCode7)
            clsCommon.AddColumnsForChange(coll, "AddDesc7", obj.AddDesc7)
            clsCommon.AddColumnsForChange(coll, "AddAmt7", obj.AddAmt7)
            clsCommon.AddColumnsForChange(coll, "AddCode8", obj.AddCode8)
            clsCommon.AddColumnsForChange(coll, "AddDesc8", obj.AddDesc8)
            clsCommon.AddColumnsForChange(coll, "AddAmt8", obj.AddAmt8)
            clsCommon.AddColumnsForChange(coll, "AddCode9", obj.AddCode9)
            clsCommon.AddColumnsForChange(coll, "AddDesc9", obj.AddDesc9)
            clsCommon.AddColumnsForChange(coll, "AddAmt9", obj.AddAmt9)
            clsCommon.AddColumnsForChange(coll, "AddCode10", obj.AddCode10)
            clsCommon.AddColumnsForChange(coll, "AddDesc10", obj.AddDesc10)
            clsCommon.AddColumnsForChange(coll, "AddAmt10", obj.AddAmt10)
            clsCommon.AddColumnsForChange(coll, "Inter_Branch", IIf(obj.Inter_Branch = True, 1, 0))
            clsCommon.AddColumnsForChange(coll, "is_Asset_Type", IIf(obj.is_Asset_Type, 1, 0))
            ''richa agarwal 19/03/2015
            clsCommon.AddColumnsForChange(coll, "Invoice_Type", obj.Invoice_Type)
            clsCommon.AddColumnsForChange(coll, "Is_ManualTCS", IIf(obj.Is_ManualTCS, 1, 0))
            ''-------------------

            '-------Ravi--------
            clsCommon.AddColumnsForChange(coll, "VAT_InvoiceNo", obj.VAT_InvoiceNo)
            clsCommon.AddColumnsForChange(coll, "VatInvoice_Type", obj.VatInvoice_Type)
            clsCommon.AddColumnsForChange(coll, "Is_Scrap", obj.Is_Scrap)
            'done by stuti
            clsCommon.AddColumnsForChange(coll, "Is_CashSale", obj.Is_CashSale)
            ''-------------------

            clsCommon.AddColumnsForChange(coll, "Total_Gross_Weight", obj.Total_Gross_Weight)
            clsCommon.AddColumnsForChange(coll, "Total_Net_Weight", obj.Total_Net_Weight)
            clsCommon.AddColumnsForChange(coll, "Inter_Unit_Sale", obj.Inter_Unit_Sale)
            clsCommon.AddColumnsForChange(coll, "IsEinvoice", obj.IsEinvoice)
            clsCommon.AddColumnsForChange(coll, "IsEwayBill", obj.IsEwayBill)

            If clsCommon.myLen(obj.Due_Date) > 0 Then
                clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Due_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
            clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Is_Taxable", IIf(obj.Is_Taxable, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Is_Cancelled", obj.Is_Cancelled)
            clsCommon.AddColumnsForChange(coll, "Gate_Entry_No", obj.Gate_Entry_No)
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Document_No", obj.Document_No)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SCRAPSALE_HEAD_Return", OMInsertOrUpdate.Insert, "", trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SCRAPSALE_HEAD_Return", OMInsertOrUpdate.Update, "TSPL_SCRAPSALE_HEAD_Return.Document_No='" + obj.Document_No + "'", trans)
            End If

            isSaved = isSaved AndAlso ClsScrapSaleDetailReturn.SaveData(obj.Document_No, Arr1, trans, obj.shipment_Date, obj.Loc_Code, obj.strInvoiceNo)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_SCRAPSALE_HEAD_Return", "Document_No", "TSPL_SCRAPSALE_DETAIL_Return", "Document_No", trans)

            'If (obj.CreateInvoice = 1) Then
            '    isSaved = isSaved AndAlso scrapinvoicehead.SaveDatainvoiceReturn(obj.Document_No, strScrapSaleInvoiceNo, trans, obj.Invoice_Type, Arr1)
            'End If

            isSaved = isSaved AndAlso clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_No, obj.arrCustomFields, trans)


            If isSaved Then
                trans.Commit()
            End If
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function CancelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal NavType As NavigatorType) As Boolean
        '' created by Sanjay
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            Dim obj As ClsScrapSaleHeadReturn = ClsScrapSaleHeadReturn.GetData(Doc_No, NavType, trans, False)

            If obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If

            ''richa agarwal 06 Jan,2021 check eInvoice Cancellation
            Dim dtirn As DataTable = clsDBFuncationality.GetDataTable("select Einvoice_type,IRN_No from TSPL_SCRAPSALE_HEAD_RETURN where Document_No='" & Doc_No & "'", trans)
            If dtirn IsNot Nothing AndAlso dtirn.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtirn.Rows(0)("Einvoice_type")), "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Is_Taxable), "1") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.shipment_Date, trans) = True Then
                    If ClsEInvoiceOFAPIs.EInvoice_Cancellation(Doc_No, clsCommon.myCstr(dtirn.Rows(0)("IRN_No")), obj.Loc_Code, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If

            clsItemLocationDetails.CheckCancelInventoryBalance(Form_Id, Doc_No, trans)

            '' transfer data into cancel table
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "TSPL_SCRAPSALE_HEAD_RETURN", "Document_No", "TSPL_SCRAPSALE_DETAIL_RETURN", "Document_No", trans)

            '' cancel customer invoice data
            qry = "select Document_No from TSPL_Customer_Invoice_Head  where AgainstScrapReturn='" & Doc_No & "'"
            Dim Document_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Document_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Document_No, "TSPL_Customer_Invoice_Head", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
            End If

            '' cancel journal master data invoice
            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where AgainstScrapReturn='" & Doc_No & "')"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            '' cancel custom field data
            clsCommonFunctionality.SaveCancelDataMultKey(objCommonVar.CurrentUserCode, Doc_No, "TSPL_CUSTOM_FIELD_VALUES", "Transaction_Code", "Program_Code", Form_Id, trans)

            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where AgainstScrapReturn='" & Doc_No & "'))"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No in (select Document_No from TSPL_Customer_Invoice_Head  where AgainstScrapReturn='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Detail where Document_No in (Select Document_No from TSPL_Customer_Invoice_Head  where AgainstScrapReturn='" & Doc_No & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_Customer_Invoice_Head where AgainstScrapReturn='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_BATCH_ITEM where  Document_Code='" & Doc_No & "' and Document_Type='MS-SR'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "' and Trans_Type='MS-SR'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_CUSTOM_FIELD_VALUES where Transaction_Code ='" & Doc_No & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_SCRAPSALE_DETAIL_RETURN where Document_No='" & Doc_No & "' "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_SCRAPSALE_HEAD_RETURN where Document_No='" & Doc_No & "' "
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



    Public Function InvoiceType(ByVal Location As String, ByVal Customer As String, ByVal trans As SqlTransaction) As String
        Dim dt As DataTable
        Dim strloc As String = Nothing
        Dim qry As String = Nothing
        Dim strInvoiceType As String = Nothing
        strloc = Location
        qry = "SELECT TSPL_LOCATION_MASTER.Excisable,TSPL_LOCATION_MASTER.State, " &
          "TSPL_LOCATION_MASTER.Sales_Tax_Group as LocalTaxGroup,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as Local_Tax_GroupName, " &
          "TSPL_LOCATION_MASTER.Sales_Tax_GroupIS as InterstateTaxGroup,TSPL_TAX_GROUP_MASTERIS.Tax_Group_Desc as Interstate_Tax_GroupName " &
          "FROM TSPL_LOCATION_MASTER left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' left outer join TSPL_TAX_GROUP_MASTER as TSPL_TAX_GROUP_MASTERIS on TSPL_TAX_GROUP_MASTERIS.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_GroupIS and TSPL_TAX_GROUP_MASTERIS.Tax_Group_Type='S' " &
          "WHERE TSPL_LOCATION_MASTER.Location_Code = '" + strloc + "'"


        dt = clsDBFuncationality.GetDataTable(qry, trans)
        Dim strLocState As String = clsCommon.myCstr(dt.Rows(0)("State"))

        qry = "select Price_Code,price_CodeNon,State,Tin_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + Customer + "'"
        dt = clsDBFuncationality.GetDataTable(qry, trans)
        Dim strCustState As String = clsCommon.myCstr(dt.Rows(0)("State"))
        Dim strTinNo As String = clsCommon.myCstr(dt.Rows(0)("Tin_No"))

        If clsCommon.myLen(strTinNo) > 0 AndAlso clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
            strInvoiceType = "T"
        Else
            If clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
                strInvoiceType = "R"
            Else
                strInvoiceType = "I" 'Interstate series
            End If

        End If

        Return strInvoiceType
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal NavType As NavigatorType) As ClsScrapSaleHeadReturn
        Return GetData(strDocumentNo, NavType, Nothing, False)
    End Function

    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction, ByVal isAssetType As Boolean) As ClsScrapSaleHeadReturn
        Return GetData(strPONo, NavType, trans, isAssetType, "")
    End Function
    Public Shared Function GetData(ByVal strPONo As String, ByVal NavType As NavigatorType, ByVal trans As SqlTransaction, ByVal isAssetType As Boolean, ByVal strDocType As String) As ClsScrapSaleHeadReturn
        Dim obj As ClsScrapSaleHeadReturn = Nothing
        Dim qry As String = "SELECT  TSPL_SCRAPSALE_HEAD_Return.Gate_Entry_No,TSPL_SCRAPSALE_HEAD_Return.Document_No,TSPL_SCRAPSALE_HEAD_Return.Doc_Type, TSPL_SCRAPSALE_HEAD_Return.Inter_Branch, TSPL_SCRAPSALE_HEAD_Return.Invoice_Type,TSPL_SCRAPSALE_HEAD_Return.shipment_No,TSPL_SCRAPSALE_HEAD_Return.Return_ship_Date,TSPL_SCRAPSALE_HEAD_Return.cust_Code,TSPL_SCRAPSALE_HEAD_Return.Po_No,TSPL_SCRAPSALE_HEAD_Return.NRG_no,TSPL_SCRAPSALE_HEAD_Return.invoice_No,TSPL_SCRAPSALE_HEAD_Return.cust_Name,TSPL_SCRAPSALE_HEAD_Return.Return_ship_Date,TSPL_SCRAPSALE_HEAD_Return.posting_date,TSPL_SCRAPSALE_HEAD_Return.expship_date,TSPL_SCRAPSALE_HEAD_Return.Status,TSPL_SCRAPSALE_HEAD_Return.CreateInvoice,TSPL_SCRAPSALE_HEAD_Return.ispost,TSPL_SCRAPSALE_HEAD_Return.Is_ManualTCS,TSPL_SCRAPSALE_HEAD_Return.Description,TSPL_SCRAPSALE_HEAD_Return.reff,TSPL_SCRAPSALE_HEAD_Return.Tax_Group, TSPL_SCRAPSALE_HEAD_Return.tax_desc, TSPL_SCRAPSALE_HEAD_Return.loc_code, TSPL_SCRAPSALE_HEAD_Return.Vehicle_Id,TSPL_SCRAPSALE_HEAD_Return.loc_Name,TSPL_SCRAPSALE_HEAD_Return.ToLoc_code,TSPL_SCRAPSALE_HEAD_Return.TAX1,TSPL_SCRAPSALE_HEAD_Return.TAX1_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX1_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX1_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX2,TSPL_SCRAPSALE_HEAD_Return.TAX2_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX2_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX2_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX3,TSPL_SCRAPSALE_HEAD_Return.TAX3_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX3_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX3_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX4,TSPL_SCRAPSALE_HEAD_Return.TAX4_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX4_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX4_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX5,TSPL_SCRAPSALE_HEAD_Return.TAX5_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX5_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX5_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX6,TSPL_SCRAPSALE_HEAD_Return.TAX6_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX6_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX6_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX7,TSPL_SCRAPSALE_HEAD_Return.TAX7_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX7_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX7_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX8,TSPL_SCRAPSALE_HEAD_Return.TAX8_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX8_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX8_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX9,TSPL_SCRAPSALE_HEAD_Return.TAX9_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX9_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX9_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX10,TSPL_SCRAPSALE_HEAD_Return.TAX10_Rate,TSPL_SCRAPSALE_HEAD_Return.TAX10_Amt,TSPL_SCRAPSALE_HEAD_Return.TAX10_Base_Amt,TSPL_SCRAPSALE_HEAD_Return.Addcost,TSPL_SCRAPSALE_HEAD_Return.AddCostDesc,TSPL_SCRAPSALE_HEAD_Return.Add_Amt,TSPL_SCRAPSALE_HEAD_Return.Before_add_Amt,TSPL_SCRAPSALE_HEAD_Return.After_Add_Amt,TSPL_SCRAPSALE_HEAD_Return.Discount_Base,TSPL_SCRAPSALE_HEAD_Return.Discount_Amt,TSPL_SCRAPSALE_HEAD_Return.Amount_Less_Discount,TSPL_SCRAPSALE_HEAD_Return.Total_Tax_Amt,TSPL_SCRAPSALE_HEAD_Return.ship_total_Amt,TSPL_SCRAPSALE_HEAD_Return.Doc_Amt,TSPL_SCRAPSALE_HEAD_Return.Comp_Code,TSPL_SCRAPSALE_HEAD_Return.Terms_Code,TSPL_SCRAPSALE_HEAD_Return.Due_Date ,TSPL_SCRAPSALE_HEAD_Return.AddCode1,TSPL_SCRAPSALE_HEAD_Return.AddDesc1,TSPL_SCRAPSALE_HEAD_Return.AddAmt1,TSPL_SCRAPSALE_HEAD_Return.AddCode2,TSPL_SCRAPSALE_HEAD_Return.AddDesc2,TSPL_SCRAPSALE_HEAD_Return.AddAmt2,TSPL_SCRAPSALE_HEAD_Return.AddCode3,TSPL_SCRAPSALE_HEAD_Return.AddDesc3,TSPL_SCRAPSALE_HEAD_Return.AddAmt3,TSPL_SCRAPSALE_HEAD_Return.AddCode4,TSPL_SCRAPSALE_HEAD_Return.AddDesc4,TSPL_SCRAPSALE_HEAD_Return.AddAmt4,TSPL_SCRAPSALE_HEAD_Return.AddCode5,TSPL_SCRAPSALE_HEAD_Return.AddDesc5,TSPL_SCRAPSALE_HEAD_Return.AddAmt5,TSPL_SCRAPSALE_HEAD_Return.AddCode6,TSPL_SCRAPSALE_HEAD_Return.AddDesc6,TSPL_SCRAPSALE_HEAD_Return.AddAmt6,TSPL_SCRAPSALE_HEAD_Return.AddCode7,TSPL_SCRAPSALE_HEAD_Return.AddDesc7,TSPL_SCRAPSALE_HEAD_Return.AddAmt7,TSPL_SCRAPSALE_HEAD_Return.AddCode8,TSPL_SCRAPSALE_HEAD_Return.AddDesc8,TSPL_SCRAPSALE_HEAD_Return.AddAmt8,TSPL_SCRAPSALE_HEAD_Return.AddCode9,TSPL_SCRAPSALE_HEAD_Return.AddDesc9,TSPL_SCRAPSALE_HEAD_Return.AddAmt9,TSPL_SCRAPSALE_HEAD_Return.AddCode10,TSPL_SCRAPSALE_HEAD_Return.AddDesc10,TSPL_SCRAPSALE_HEAD_Return.AddAmt10,TSPL_SCRAPSALE_HEAD_Return.Tax_Calculation_Type,TSPL_LOCATION_MASTER.Location_Desc as ToLocationName,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as ShipToLocationName,TSPL_TERMS_MASTER.Terms_Desc as TermsName,TSPL_SCRAPSALE_HEAD_Return.EXCISABLE, " &
        "(select invoice_No from TSPL_SCRAPINVOICE_HEAD where TSPL_SCRAPINVOICE_HEAD.shipment_No=TSPL_SCRAPSALE_HEAD_Return.shipment_No) as ScrapInvoiceNo,TSPL_SCRAPSALE_HEAD_Return.is_Asset_Type,Transport_code,Transporter_name,Vehicle_code,TSPL_SCRAPSALE_HEAD_Return.VAT_InvoiceNo,TSPL_SCRAPSALE_HEAD_Return.VatInvoice_Type,TSPL_SCRAPSALE_HEAD_Return.Is_Scrap,TSPL_SCRAPSALE_HEAD_Return.Is_CashSale,TSPL_SCRAPSALE_HEAD_Return.Total_Gross_Weight,TSPL_SCRAPSALE_HEAD_Return.Total_Net_Weight,TSPL_SCRAPSALE_HEAD_Return.RoundOffAmount,TSPL_SCRAPSALE_HEAD_Return.Is_Taxable,TSPL_SCRAPSALE_HEAD_Return.EWayBillNo,TSPL_SCRAPSALE_HEAD_Return.Electronic_Ref_No,TSPL_SCRAPSALE_HEAD_Return.EWayBillDate, TSPL_SCRAPSALE_HEAD_Return.Is_Cancelled,TSPL_SCRAPSALE_HEAD_Return.Inter_Unit_Sale,TSPL_SCRAPSALE_HEAD_Return.IsEinvoice,TSPL_SCRAPSALE_HEAD_Return.IsEwayBill FROM TSPL_SCRAPSALE_HEAD_Return left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SCRAPSALE_HEAD_Return.loc_code  left join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.Transport_Id=Transport_code  left outer join TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SCRAPSALE_HEAD_Return.loc_code  left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_SCRAPSALE_HEAD_Return.Terms_Code where 2=2"
        Dim whrCls As String = ""

        If isAssetType Then
            whrCls += " and is_Asset_Type=1 "
        End If
        If clsCommon.myLen(strDocType) > 0 Then
            whrCls += " and Doc_Type ='" + strDocType + "'"
        End If

        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrCls = " AND loc_code in (" + objCommonVar.strCurrUserLocations + ")"
        End If
        Select Case NavType
            Case NavigatorType.First
                qry += " and TSPL_SCRAPSALE_HEAD_Return.document_No = (select MIN(document_No) from TSPL_SCRAPSALE_HEAD_Return WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Last
                qry += " and TSPL_SCRAPSALE_HEAD_Return.document_No = (select Max(document_No) from TSPL_SCRAPSALE_HEAD_Return WHERE 1=1 " + whrCls + ")"
            Case NavigatorType.Current
                qry += " and TSPL_SCRAPSALE_HEAD_Return.document_No = '" + strPONo + "'"
            Case NavigatorType.Next
                qry += " and TSPL_SCRAPSALE_HEAD_Return.document_No = (select Min(document_No) from TSPL_SCRAPSALE_HEAD_Return where document_No>'" + strPONo + "' " + whrCls + ")"
            Case NavigatorType.Previous
                qry += " and TSPL_SCRAPSALE_HEAD_Return.document_No = (select Max(document_No) from TSPL_SCRAPSALE_HEAD_Return where document_No<'" + strPONo + "' " + whrCls + ")"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New ClsScrapSaleHeadReturn()
            obj.Gate_Entry_No = clsCommon.myCstr(dt.Rows(0)("Gate_Entry_No"))
            obj.Is_Cancelled = clsCommon.myCdbl(dt.Rows(0)("Is_Cancelled"))
            obj.Document_No = clsCommon.myCstr(dt.Rows(0)("Document_No"))
            obj.shipment_No = clsCommon.myCstr(dt.Rows(0)("shipment_No"))
            obj.invoice_No = clsCommon.myCstr(dt.Rows(0)("invoice_No"))
            obj.Doc_Type = clsCommon.myCstr(dt.Rows(0)("Doc_Type"))
            obj.Po_No = clsCommon.myCstr(dt.Rows(0)("Po_No"))
            obj.NRG_No = clsCommon.myCstr(dt.Rows(0)("NRG_No"))
            obj.Status = clsCommon.myCstr(dt.Rows(0)("Status"))
            obj.ispost = clsCommon.myCstr(dt.Rows(0)("ispost"))
            obj.cust_Code = clsCommon.myCstr(dt.Rows(0)("cust_Code"))
            obj.cust_Name = clsCommon.myCstr(dt.Rows(0)("cust_Name"))
            obj.shipment_Date = clsCommon.myCDate(dt.Rows(0)("Return_ship_Date"))
            obj.posting_Date = clsCommon.myCDate(dt.Rows(0)("posting_Date"))
            obj.expship_Date = clsCommon.myCDate(dt.Rows(0)("expship_Date"))
            obj.Loc_Code = clsCommon.myCstr(dt.Rows(0)("Loc_Code"))
            obj.Loc_Name = clsCommon.myCstr(dt.Rows(0)("Loc_Name"))
            obj.Vehicle_Id = clsCommon.myCstr(dt.Rows(0)("Vehicle_Id"))
            obj.ToLoc_Code = clsCommon.myCstr(dt.Rows(0)("ToLoc_Code"))
            obj.CreateInvoice = clsCommon.myCstr(dt.Rows(0)("CreateInvoice"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.reff = clsCommon.myCstr(dt.Rows(0)("reff"))
            obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
            obj.Transporter_code = clsCommon.myCstr(dt.Rows(0)("Transport_code"))
            obj.Transporter_Name = clsCommon.myCstr(dt.Rows(0)("Transporter_name"))
            obj.Vehicle_code = clsCommon.myCstr(dt.Rows(0)("Vehicle_code"))
            obj.Tax_Desc = clsCommon.myCstr(dt.Rows(0)("Tax_Desc"))
            obj.Add_Amt = clsCommon.myCstr(dt.Rows(0)("Add_Amt"))
            obj.Before_Add_Amt = clsCommon.myCstr(dt.Rows(0)("Before_Add_Amt"))
            obj.Discount_Base = clsCommon.myCstr(dt.Rows(0)("Discount_Base"))
            obj.Discount_Amt = clsCommon.myCstr(dt.Rows(0)("Discount_Amt"))
            obj.Amount_Less_Discount = clsCommon.myCstr(dt.Rows(0)("Amount_Less_Discount"))
            obj.Total_Tax_Amt = clsCommon.myCstr(dt.Rows(0)("Total_tax_amt"))
            obj.ship_Total_Amt = clsCommon.myCstr(dt.Rows(0)("ship_total_amt"))
            obj.doc_Amt = clsCommon.myCstr(dt.Rows(0)("Doc_Amt"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            obj.Excisable = clsCommon.myCstr(dt.Rows(0)("EXCISABLE"))
            obj.Inter_Unit_Sale = clsCommon.myCdbl(dt.Rows(0)("Inter_Unit_Sale"))
            obj.IsEinvoice = clsCommon.myCdbl(dt.Rows(0)("IsEinvoice"))
            obj.IsEwayBill = clsCommon.myCdbl(dt.Rows(0)("IsEwayBill"))
            obj.Inter_Branch = IIf(clsCommon.myCdbl(dt.Rows(0)("Inter_Branch")) = 1, True, False)
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
            obj.Is_ManualTCS = clsCommon.myCdbl(dt.Rows(0)("Is_ManualTCS"))
            obj.AddCode1 = clsCommon.myCstr(dt.Rows(0)("AddCode1"))
            obj.AddDesc1 = clsCommon.myCstr(dt.Rows(0)("AddDesc1"))
            obj.AddAmt1 = clsCommon.myCdbl(dt.Rows(0)("AddAmt1"))
            obj.AddCode2 = clsCommon.myCstr(dt.Rows(0)("AddCode2"))
            obj.AddDesc2 = clsCommon.myCstr(dt.Rows(0)("AddDesc2"))
            obj.AddAmt2 = clsCommon.myCdbl(dt.Rows(0)("AddAmt2"))
            obj.AddCode3 = clsCommon.myCstr(dt.Rows(0)("AddCode3"))
            obj.AddDesc3 = clsCommon.myCstr(dt.Rows(0)("AddDesc3"))
            obj.AddAmt3 = clsCommon.myCdbl(dt.Rows(0)("AddAmt3"))
            obj.AddCode4 = clsCommon.myCstr(dt.Rows(0)("AddCode4"))
            obj.AddDesc4 = clsCommon.myCstr(dt.Rows(0)("AddDesc4"))
            obj.AddAmt4 = clsCommon.myCdbl(dt.Rows(0)("AddAmt4"))
            obj.AddCode5 = clsCommon.myCstr(dt.Rows(0)("AddCode5"))
            obj.AddDesc5 = clsCommon.myCstr(dt.Rows(0)("AddDesc5"))
            obj.AddAmt5 = clsCommon.myCdbl(dt.Rows(0)("AddAmt5"))
            obj.AddCode6 = clsCommon.myCstr(dt.Rows(0)("AddCode6"))
            obj.AddDesc6 = clsCommon.myCstr(dt.Rows(0)("AddDesc6"))
            obj.AddAmt6 = clsCommon.myCdbl(dt.Rows(0)("AddAmt6"))
            obj.AddCode7 = clsCommon.myCstr(dt.Rows(0)("AddCode7"))
            obj.AddDesc7 = clsCommon.myCstr(dt.Rows(0)("AddDesc7"))
            obj.AddAmt7 = clsCommon.myCdbl(dt.Rows(0)("AddAmt7"))
            obj.AddCode8 = clsCommon.myCstr(dt.Rows(0)("AddCode8"))
            obj.AddDesc8 = clsCommon.myCstr(dt.Rows(0)("AddDesc8"))
            obj.AddAmt8 = clsCommon.myCdbl(dt.Rows(0)("AddAmt8"))
            obj.AddCode9 = clsCommon.myCstr(dt.Rows(0)("AddCode9"))
            obj.AddDesc9 = clsCommon.myCstr(dt.Rows(0)("AddDesc9"))
            obj.AddAmt9 = clsCommon.myCdbl(dt.Rows(0)("AddAmt9"))
            obj.AddCode10 = clsCommon.myCstr(dt.Rows(0)("AddCode10"))
            obj.AddDesc10 = clsCommon.myCstr(dt.Rows(0)("AddDesc10"))
            obj.AddAmt10 = clsCommon.myCdbl(dt.Rows(0)("AddAmt10"))

            obj.is_Asset_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("is_Asset_Type")) = 1, True, False)
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            ''richa agarwal 19/03/2015
            obj.Invoice_Type = clsCommon.myCstr(dt.Rows(0)("Invoice_Type"))
            ''-----------
            '------Ravi----------
            obj.VAT_InvoiceNo = clsCommon.myCstr(dt.Rows(0)("VAT_InvoiceNo"))
            obj.VatInvoice_Type = clsCommon.myCstr(dt.Rows(0)("VatInvoice_Type"))
            obj.Is_Scrap = clsCommon.myCstr(dt.Rows(0)("Is_Scrap"))
            ''-----------
            'done by stuti
            obj.Is_CashSale = clsCommon.myCstr(dt.Rows(0)("Is_CashSale"))

            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)

            obj.strInvoiceNo = clsCommon.myCstr(dt.Rows(0)("ScrapInvoiceNo"))


            obj.Total_Gross_Weight = clsCommon.myCdbl(dt.Rows(0)("Total_Gross_Weight"))
            obj.Total_Net_Weight = clsCommon.myCdbl(dt.Rows(0)("Total_Net_Weight"))

            obj.Is_Taxable = If(clsCommon.myCdbl(dt.Rows(0)("Is_Taxable")) > 0, True, False)
            obj.EWayBillNo = clsCommon.myCstr(dt.Rows(0)("EWayBillNo"))
            If dt.Rows(0)("EWayBillDate") IsNot DBNull.Value Then
                obj.EWayBillDate = clsCommon.myCDate(dt.Rows(0)("EWayBillDate"))
            End If
            obj.Electronic_Ref_No = clsCommon.myCstr(dt.Rows(0)("Electronic_Ref_No"))
            qry = "SELECT TSPL_SCRAPSALE_DETAIL_Return.ItemwiseTaxCode, TSPL_SCRAPSALE_DETAIL_Return.shipment_No,TSPL_SCRAPSALE_DETAIL_Return.Document_No,TSPL_SCRAPSALE_DETAIL_Return.Specification,TSPL_SCRAPSALE_DETAIL_Return.Line_No,TSPL_SCRAPSALE_DETAIL_Return.Item_Code,TSPL_SCRAPSALE_DETAIL_Return.Item_Desc,TSPL_SCRAPSALE_DETAIL_Return.unit_code,TSPL_SCRAPSALE_DETAIL_Return.Shipped_Qty,TSPL_SCRAPSALE_DETAIL_Return.pending_qty,TSPL_SCRAPSALE_DETAIL_Return.price,TSPL_SCRAPSALE_DETAIL_Return.DiscountPer,TSPL_SCRAPSALE_DETAIL_Return.DiscountAmt,TSPL_SCRAPSALE_DETAIL_Return.TotalTaxAmt,TSPL_SCRAPSALE_DETAIL_Return.NetPriceAmt,TSPL_SCRAPSALE_DETAIL_Return.ItemAmt,TSPL_SCRAPSALE_DETAIL_Return.TotalDiscountAmt,TSPL_SCRAPSALE_DETAIL_Return.ItemNetAmt,TSPL_SCRAPSALE_DETAIL_Return.TotalAmt,TSPL_SCRAPSALE_DETAIL_Return.Specification,TSPL_SCRAPSALE_DETAIL_Return.TAX1,TSPL_SCRAPSALE_DETAIL_Return.TAX1_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX1_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX2,TSPL_SCRAPSALE_DETAIL_Return.TAX2_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX2_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX3,TSPL_SCRAPSALE_DETAIL_Return.TAX3_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX3_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX4,TSPL_SCRAPSALE_DETAIL_Return.TAX4_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX4_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX5,TSPL_SCRAPSALE_DETAIL_Return.TAX5_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX5_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX6,TSPL_SCRAPSALE_DETAIL_Return.TAX6_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX6_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX7,TSPL_SCRAPSALE_DETAIL_Return.TAX7_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX7_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX8,TSPL_SCRAPSALE_DETAIL_Return.TAX8_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX8_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX9,TSPL_SCRAPSALE_DETAIL_Return.TAX9_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX9_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX10,TSPL_SCRAPSALE_DETAIL_Return.TAX10_Rate,TSPL_SCRAPSALE_DETAIL_Return.TAX10_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX1_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX2_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX3_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX4_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX5_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX6_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX7_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX8_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX9_Base_Amt,TSPL_SCRAPSALE_DETAIL_Return.TAX10_Base_Amt,Asset_Code,TSPL_SCRAPSALE_DETAIL_Return.FAT,TSPL_SCRAPSALE_DETAIL_Return.SNF  FROM TSPL_SCRAPSALE_DETAIL_Return  where TSPL_SCRAPSALE_DETAIL_Return.Document_No='" + obj.Document_No + "' ORDER BY TSPL_SCRAPSALE_DETAIL_Return.Line_No"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr1 = New List(Of ClsScrapSaleDetailReturn)
                Dim objTr As ClsScrapSaleDetailReturn
                For Each dr As DataRow In dt.Rows
                    objTr = New ClsScrapSaleDetailReturn
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(dr("ItemwiseTaxCode"))
                    objTr.document_No = clsCommon.myCstr(dr("document_No"))
                    objTr.shipment_No = clsCommon.myCstr(dr("shipment_No"))
                    objTr.Line_No = clsCommon.myCstr(dr("Line_No"))
                    objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                    objTr.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                    objTr.Unit_Code = clsCommon.myCstr(dr("Unit_Code"))
                    objTr.shipped_Qty = clsCommon.myCdbl(dr("shipped_Qty"))

                    'sanjay 
                    'objTr.shQty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("", trans))
                    objTr.pending_Qty = clsCommon.myCdbl(dr("pending_qty"))
                    'sanjay

                    objTr.FAT = clsCommon.myCdbl(dr("FAT"))
                    objTr.SNF = clsCommon.myCdbl(dr("SNF"))

                    objTr.Asset_Code = clsCommon.myCstr(dr("Asset_Code"))
                    objTr.price = clsCommon.myCdbl(dr("price"))
                    objTr.DiscountPer = clsCommon.myCdbl(dr("DiscountPer"))
                    objTr.DiscountAmt = clsCommon.myCdbl(dr("DiscountAmt"))
                    objTr.TotalTaxAmt = clsCommon.myCdbl(dr("TotalTaxAmt"))
                    objTr.NetPriceAmt = clsCommon.myCdbl(dr("NetPriceAmt"))
                    objTr.ItemAmt = clsCommon.myCdbl(dr("ItemAmt"))
                    objTr.TotalDiscountAmt = clsCommon.myCdbl(dr("TotalDiscountAmt"))
                    objTr.ItemNetAmt = clsCommon.myCdbl(dr("ItemNetAmt"))
                    objTr.TotalAmt = clsCommon.myCdbl(dr("TotalAmt"))

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
                    objTr.Specification = clsCommon.myCstr(dr("Specification"))
                    objTr.arrBatchItem = clsBatchInventory.GetData("MS-SR", objTr.document_No, objTr.Item_Code, objTr.Line_No, trans)
                    obj.Arr1.Add(objTr)
                Next
            End If
        End If

        Return obj
    End Function
    Private Shared Function GetFirstItemCode(ByVal Arr As List(Of ClsScrapSaleDetailReturn)) As String
        For Each objtr As ClsScrapSaleDetailReturn In Arr
            Return objtr.Item_Code
        Next
        Return ""
    End Function
    Public Shared Function PostData(ByVal strDocNo As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim isSaved As Boolean = True
            Dim isscrap As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Shipment No not found to Post")
            End If
            ''Dim strPostDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy hh:mm tt")

            Dim obj As ClsScrapSaleHeadReturn = ClsScrapSaleHeadReturn.GetData(strDocNo, NavigatorType.Current, trans, False)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.shipment_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If

            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase Order", "Scrap Invoice", obj.Loc_Code, obj.shipment_Date, trans)

            If (obj.ispost = 1) Then
                Throw New Exception("Already Post on :" + obj.posting_Date)
            End If
            If (obj.Status = 1) Then
                Throw New Exception("Shipment No " + obj.shipment_No + " Is On Hold.Can't Post it")
            End If


            Dim qry As String = ""

            If obj.is_Asset_Type Then

                qry = "Update TSPL_SCRAPSALE_HEAD_RETURN set ispost=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
                qry += " where Shipment_No='" + strDocNo + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "Update TSPL_SCRAPSALE_HEAD_RETURN set ispost=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt ") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
                qry += " where invoice_No='" + obj.strInvoiceNo + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                trans.Commit()
                Return True
            End If



            Dim ArrLocationDetails As List(Of clsItemLocationDetails) = New List(Of clsItemLocationDetails)()
            Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)

            Dim strFirstItemCodeNonItemRowType As String = GetFirstItemCode(obj.Arr1)
            Dim strRgpNo As String = Nothing
            Dim intCounter As Integer = 0
            For Each objTr As ClsScrapSaleDetailReturn In obj.Arr1
                intCounter = intCounter + 1
                qry = "select TSPL_ITEM_MASTER.Purchase_Class_Code,TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_PURCHASE_ACCOUNTS.Inv_Payable_Clearing,TSPL_PURCHASE_ACCOUNTS.Assembly_Cost_Credit,TSPL_PURCHASE_ACCOUNTS.Breakage_Gl_Account  from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objTr.Item_Code + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    Throw New Exception("Please set Purchase Account set for item " + objTr.Item_Code + "(" + objTr.Item_Desc + ")")
                End If


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
                    'Throw New Exception("Item Type not found: " + strItemType)
                End If
                Dim objLocationDetails As New clsItemLocationDetails()
                Dim ConvFac As Double = clsItemMaster.GetConvertionFactor(objTr.Item_Code, objTr.Unit_Code, trans)
                If ConvFac = 0 Then
                    Throw New Exception("Conversion Factor found zero for item :" + objTr.Item_Code + " and Uom:'" + objTr.Unit_Code)
                End If


                Dim arr As New List(Of String)

                Dim objInventoryMovemnt As New clsInventoryMovement()
                objInventoryMovemnt.InOut = "I"
                objInventoryMovemnt.Location_Code = obj.Loc_Code

                objInventoryMovemnt.Cust_Code = obj.cust_Code
                objInventoryMovemnt.Cust_Name = obj.cust_Name

                objInventoryMovemnt.Item_Code = objTr.Item_Code
                objInventoryMovemnt.Item_Desc = objTr.Item_Desc
                objInventoryMovemnt.Qty = objTr.shipped_Qty
                objInventoryMovemnt.UOM = objTr.Unit_Code
                objInventoryMovemnt.Basic_Cost = objTr.price
                objInventoryMovemnt.MRP = 0

                objInventoryMovemnt.Add_Cost = objTr.TotalTaxAmt
                objInventoryMovemnt.Net_Cost = objTr.TotalTaxAmt
                'If clsCommon.CompairString(strItemType, "R") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "RM"
                'ElseIf clsCommon.CompairString(strItemType, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(strItemType, "O") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "OT"
                'ElseIf clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal Then
                '    objInventoryMovemnt.ItemType = "FT"
                'End If
                objInventoryMovemnt.ItemType = strItemTypeToSave
                ArrInventoryMovement.Add(objInventoryMovemnt)
                ' End If
            Next
            isSaved = isSaved AndAlso clsItemLocationDetails.SaveData(clsCommon.GetPrintDate(obj.shipment_Date, "dd/MM/yyyy"), ArrLocationDetails, trans)
            isSaved = isSaved AndAlso clsInventoryMovement.SaveData("MS-SR", obj.Document_No, obj.shipment_Date, clsCommon.GetPrintDate(obj.shipment_Date, "dd/MM/yyyy"), ArrInventoryMovement, trans)


            createARInvoice(obj, trans)

            ''richa DO COMMENT ON 28 mARCH,2019 BECUASE THIS TABLE IS not used in this transaction
            'qry = "Update TSPL_SD_SALE_RETURN_HEAD set Status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
            'qry += " where Document_Code='" + strDocNo + "'"
            'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_SCRAPINVOICE_HEAD", "invoice_No", obj.invoice_No, trans)
            ''richa agarwal 28 Dec,2020 check eInvoice Implementation
            If obj.IsEinvoice = 1 Then
                If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso obj.Is_Taxable AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.shipment_Date, trans) = True Then
                    If ClsScrapSaleHeadReturn.EInvoice_Implementation(obj.Document_No, obj.Loc_Code, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If



            qry = "Update TSPL_SCRAPSALE_HEAD_RETURN set ispost=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt") + "',Modify_By='" + objCommonVar.CurrentUserCode + "',EInvoice_Type='" + ECustomerType + "'"
            qry += " where Document_No='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strDocNo), "TSPL_SCRAPSALE_HEAD_Return", "Document_No", trans)
            If objCommonVar.InternalSMSEmailinPurchaseModule = True Then
                CreateInternalEmailSMS_ScrapReturn(obj, trans)
            End If

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
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
                Dim strQry As String = "select TSPL_Customer_master .Cust_Code ,tspl_scrapsale_head_return.invoice_No as DocNo,convert(date,tspl_scrapsale_head_return.return_ship_date,103) as DocDate,case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'DBN' when TSPL_Customer_Invoice_Head.Document_Type ='I' then 'INV' else 'CRN' end as DocType ,'B2B' as SupTyp, 'N'  as IgstOnIntra,Bill_To_Location.GSTNO as SellerGSTINNo ,Bill_To_Location.location_desc as SellerLglNm,TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm,Bill_To_Location.Add1 as SellerAdd1,Bill_To_Location.Add2 as SellerAdd2 ,Bill_To_Location.city_code as SellerLoc,Bill_To_Location.Pin_Code  as SellerPincode,BillToLocation_State_Master.GST_STATE_Code as SellerStcd,Bill_To_Location.Phone1 as SellerPhone,Bill_To_Location.Email as SellerEmail,TSPL_Customer_master.GSTNo as BuyerGSTINNo ,TSPL_Customer_master.Customer_Name as BuyerLglNm,TSPL_Customer_master.alies_name as BuyerTrdNm,case when isnull(tspl_scrapsale_head_return.ToLoc_Code,'')='' then Customer_State_Master.GST_STATE_Code else Ship_To_Location_State_Master.GST_STATE_Code end as BuyerPOS,TSPL_Customer_master.Add1 as BuyerAdd1,TSPL_Customer_master.Add2 as BuyerAdd2 ,tspl_city_master.City_Name as BuyerLoc,cast(TSPL_Customer_master.PIN_NO as int) as BuyerPincode,Customer_State_Master.GST_STATE_Code as BuyerStcd,TSPL_Customer_master.Phone1 as BuyerPhone,TSPL_Customer_master.Email as BuyerEmail,tspl_scrapsale_detail_return.Line_No as ItemSlNo, 'N' as ItemIsServc,TSPL_ITEM_MASTER.Item_Desc AS ItemPrdDesc,TSPL_ITEM_MASTER.HSN_Code AS ItemHsnCd,tspl_scrapsale_detail_return.shipped_qty as ItemQty, tspl_scrapsale_detail_return.Unit_code
as ItemUnit,tspl_scrapsale_detail_return.price as ItemUnitPrice,tspl_scrapsale_detail_return.ItemAmt as ItemTotAmt,tspl_scrapsale_detail_return.DiscountAmt as ItemDiscount,tspl_scrapsale_detail_return.ItemNetAmt as ItemAssAmt,case when ISNULL(tspl_scrapsale_detail_return .tax1,'') ='IGST' THEN tspl_scrapsale_detail_return.TAX1_Rate when ISNULL(tspl_scrapsale_detail_return .tax1,'') ='CGST' AND   ISNULL(tspl_scrapsale_detail_return .tax2,'') ='SGST'  THEN tspl_scrapsale_detail_return.TAX1_Rate+tspl_scrapsale_detail_return.TAX2_Rate  ELSE 0 end as ItemGstRt, case when tspl_scrapsale_detail_return .TAX1 ='SGST' AND tspl_scrapsale_detail_return .TAX2  ='CGST' then tspl_scrapsale_detail_return.TAX1_Amt when tspl_scrapsale_detail_return .TAX1 ='CGST' AND tspl_scrapsale_detail_return .TAX2  ='SGST' then tspl_scrapsale_detail_return.TAX2_Amt else 0 end ItemSgstAmt,case when tspl_scrapsale_detail_return .TAX1 ='IGST' then tspl_scrapsale_detail_return.TAX1_Amt else 0 end ItemIgstAmt,case when tspl_scrapsale_detail_return .TAX1 ='CGST' AND tspl_scrapsale_detail_return .TAX2  ='SGST' then tspl_scrapsale_detail_return.TAX1_Amt when tspl_scrapsale_detail_return .TAX1 ='SGST' AND tspl_scrapsale_detail_return .TAX2  ='CGST' then tspl_scrapsale_detail_return.TAX2_Amt else 0 end ItemCgstAmt,0 as ItemOthChrg,tspl_scrapsale_detail_return.TotalAmt-case when isnull(TCS1.is_tcs,'')='Y' THEN  tspl_scrapsale_detail_return.TAX2_AMT when  isnull(TCS2.is_tcs,'')='Y' THEN  tspl_scrapsale_detail_return.TAX3_AMT ELSE 0 END as ItemTotItemVal,tspl_scrapsale_head_return .Amount_Less_Discount as ValDtlsAssVal,case when tspl_scrapsale_head_return .TAX1 ='CGST' AND tspl_scrapsale_head_return .TAX2  ='SGST' then tspl_scrapsale_head_return.TAX1_Amt when tspl_scrapsale_head_return .TAX1 ='SGST' AND tspl_scrapsale_head_return .TAX2  ='CGST' then tspl_scrapsale_head_return.TAX2_Amt else 0 end ValDtlsCgstVal, case when tspl_scrapsale_head_return .TAX1 ='SGST' AND tspl_scrapsale_head_return .TAX2  ='CGST' then tspl_scrapsale_head_return.TAX1_Amt when tspl_scrapsale_head_return .TAX1 ='CGST' AND tspl_scrapsale_head_return .TAX2  ='SGST' then tspl_scrapsale_head_return.TAX2_Amt else 0 end ValDtlsSgstVal,case when tspl_scrapsale_head_return .TAX1 ='IGST' then tspl_scrapsale_head_return.TAX1_Amt else 0 end ValDtlsIgstVal,tspl_scrapsale_head_return.Discount_Amt as ValDtlsDiscount,case when isnull(TCS1.is_tcs,'')='Y' THEN  tspl_scrapsale_head_return.TAX2_AMT when isnull(TCS2.is_tcs,'')='Y' THEN  tspl_scrapsale_head_return.TAX3_AMT ELSE 0 END as ValDtlsOthChrg,tspl_scrapsale_head_return.Doc_Amt  as ValDtlsTotInvVal,tspl_scrapsale_head_return.RoundOffAmount  as ValDtlsRndOffAmt
from tspl_scrapsale_head_return
Left Outer Join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Against_Sale_No =tspl_scrapsale_head_return.invoice_No
Left Outer Join TSPL_COMPANY_MASTER  on TSPL_COMPANY_MASTER.Comp_Code  ='" & objCommonVar.CurrentCompanyCode & "'
Left Outer Join TSPL_Customer_master on TSPL_Customer_master.Cust_Code  =tspl_scrapsale_head_return.Cust_Code
left Outer Join TSPL_LOCATION_MASTER as Bill_To_Location on Bill_To_Location.Location_Code =tspl_scrapsale_head_return.Loc_Code 
left Outer Join TSPL_SHIP_TO_LOCATION as Ship_To_Location on Ship_To_Location.Ship_To_Code =tspl_scrapsale_head_return.ToLoc_Code   
left outer join tspl_scrapsale_detail_return on tspl_scrapsale_detail_return.document_no=tspl_scrapsale_head_return.document_no
left outer join tspl_item_master on tspl_item_master.Item_code=tspl_scrapsale_detail_return.Item_code
left outer join TSPL_STATE_MASTER as BillToLocation_State_Master on BillToLocation_State_Master.STATE_CODE  =Bill_To_Location.State
left outer join TSPL_STATE_MASTER as Ship_To_Location_State_Master on Ship_To_Location_State_Master.STATE_CODE  =Ship_To_Location.State
left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE  =TSPL_Customer_master.State 
left outer join tspl_city_master on tspl_city_master.city_code=TSPL_Customer_master.City_Code
left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code =tspl_scrapsale_head_return.Tax2
left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code =tspl_scrapsale_head_return.Tax3
  where tspl_scrapsale_head_return.document_no ='" & strDocNo & "'"

                Dim objResult As Object = ClsEInvoiceOFAPIs.PostAuthTokenNo_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strQry, strLocation, trans, True)
                If objResult IsNot Nothing Then
                    'assign to variable
                    Dim AckNo As String = objResult.SelectToken("AckNo").ToString
                    Dim AckDt As String = objResult.SelectToken("AckDt").ToString
                    Dim Irn As String = objResult.SelectToken("Irn").ToString
                    Dim SignedQRCode As String = objResult.SelectToken("SignedQRCode").ToString
                    clsDBFuncationality.ExecuteNonQuery("update tspl_scrapsale_head_return set  IRN_No ='" & Irn & "',qr_code='" & SignedQRCode & "',ack_no='" & AckNo & "',ack_date='" & clsCommon.GetPrintDate(AckDt, "dd/MMM/yyyy hh:mm tt") & "' where document_no ='" & strDocNo & "'", trans)

                    Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(SignedQRCode)
                    clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "tspl_scrapsale_head_return", "tspl_scrapsale_head_return.document_no='" & strDocNo & "'", trans)
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
    Private Shared Sub CreateInternalEmailSMS_ScrapReturn(ByVal obj As ClsScrapSaleHeadReturn, ByVal trans As SqlTransaction)
        Dim itemName As String = ""
        Dim UOM As String = ""
        Dim qty As String = ""
        Dim ItemDetail As String = ""
        Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.ScrapSaleRetrun + "2" + "'", trans)

        Dim qry As String = "select TSPL_USER_MASTER.User_Code from TSPL_USER_MASTER "
        qry += " left join TSPL_SCRAPSALE_HEAD_RETURN on TSPL_SCRAPSALE_HEAD_RETURN.Created_By=TSPL_USER_MASTER.User_Code  "
        qry += " where TSPL_SCRAPSALE_HEAD_RETURN.document_no='" + obj.Document_No + "'"
        Dim StrUserCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        Dim arrMobileNo As New List(Of String)
        Dim arrMailID As List(Of String) = clsERPFuncationality.ReportingMailIdandPhone(StrUserCode, arrMobileNo, trans)

        If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 AndAlso ((arrMailID IsNot Nothing AndAlso arrMailID.Count > 0) Or (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0)) Then

            If (obj.Arr1 IsNot Nothing AndAlso obj.Arr1.Count > 0) Then
                For Each objdetail As ClsScrapSaleDetailReturn In obj.Arr1

                    itemName = clsCommon.myCstr(objdetail.Item_Desc)
                    UOM = clsCommon.myCstr(objdetail.Unit_Code)
                    qty = clsCommon.myCstr(objdetail.shipped_Qty)

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
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myFormat(obj.doc_Amt))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.CustomerNo, obj.cust_Code)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(clsEmailSMSConstants.CustomerName, obj.cust_Name)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)

                objEmailH.SaveData(clsUserMgtCode.ScrapSaleRetrun, objEmailH, trans)
                objEmailH = Nothing

            End If

            If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 AndAlso (arrMobileNo IsNot Nothing AndAlso arrMobileNo.Count > 0) Then
                Dim objSMSH As New clsSMSHead()
                objSMSH.arrMobilNo = New List(Of String)()
                objSMSH.arrMobilNo = arrMobileNo

                objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))

                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_NO, obj.Document_No)
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.DOC_Date, clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy"))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.Doc_Amount, clsCommon.myFormat(obj.doc_Amt))
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.CustomerNo, obj.cust_Code)
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(clsEmailSMSConstants.CustomerName, obj.cust_Name)
                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ItemDetail, ItemDetail)

                objSMSH.SaveData(clsUserMgtCode.ScrapSaleRetrun, objSMSH, trans)
                objSMSH = Nothing
            End If
        End If


    End Sub


    ''richa ERO/19/03/19-000515 added on 28 March,2019
    Public Shared Function ReverseAndUnpost(ByVal strCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(strCode) <= 0 Then
                Throw New Exception("Transaction No not found for reverse and unpost")
            End If

            Dim Qry As String = "select ispost from TSPL_SCRAPSALE_HEAD_RETURN where Document_No='" + strCode + "'"
            If Not clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) = 1 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            '' For Journal Entry of ar Invoice credit note
            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where source_code='AR-CR' and source_doc_no=(Select Document_No from TSPL_Customer_Invoice_Head where AgainstScrapReturn ='" & clsCommon.myCstr(strCode) & "')", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_Customer_Invoice_Head", "AgainstScrapReturn", trans)
            Qry = "Delete from TSPL_Customer_Invoice_Detail where Document_No =(Select Document_No from TSPL_Customer_Invoice_Head where AgainstScrapReturn ='" & clsCommon.myCstr(strCode) & "')"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            Qry = "Delete from TSPL_Customer_Invoice_Head where AgainstScrapReturn ='" & clsCommon.myCstr(strCode) & "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_INVENTORY_MOVEMENT", "Source_Doc_No", trans)
            Qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" + strCode + "' and Trans_Type='MS-SR'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            clsBatchInventory.ReverseAndUnpost("MS-SR", strCode, trans)

            Qry = "Update TSPL_SCRAPSALE_HEAD_RETURN set ispost = 0 where Document_No='" + strCode + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_SCRAPSALE_HEAD_Return", "Document_No", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function createARInvoice(ByVal obj As ClsScrapSaleHeadReturn, ByVal trans As SqlTransaction, Optional ByVal strARNoForRecreate As String = Nothing, Optional ByVal strVoucherForRecreate As String = Nothing) As Boolean
        ''''''''''''''''''''''''''''''''''For Making AR Invoice
        Dim objCustInv As New clsCustomerInvoiceHead()
        ''objCustInv.Document_No ''Will be Generateed
        objCustInv.Document_Date = obj.shipment_Date
        objCustInv.Document_Type = "C"
        objCustInv.Document_Total = obj.doc_Amt
        objCustInv.Customer_Code = obj.cust_Code
        objCustInv.Customer_Name = obj.cust_Name
        objCustInv.Posting_Date = obj.shipment_Date
        objCustInv.Trans_Type = "SC"
        Dim qry As String = " select Cust_Account from TSPL_CUSTOMER_MASTER where Cust_Code='" + obj.cust_Code + "'"
        objCustInv.Account_Set = clsDBFuncationality.getSingleValue(qry, trans)
        ''objCustInv.Order_No
        objCustInv.loc_code = clsLocation.GetSegmentCode(obj.Loc_Code, trans)
        objCustInv.Location_Code_Prefix = obj.Loc_Code
        objCustInv.On_Hold = 0
        objCustInv.Remarks = obj.Description
        objCustInv.Description = obj.Description
        objCustInv.Tax_Group = obj.Tax_Group
        objCustInv.TAX1 = obj.TAX1
        objCustInv.TAX1_Rate = obj.TAX1_Rate
        objCustInv.TAX1_Amt = obj.TAX1_Amt
        objCustInv.TAX2 = obj.TAX2
        objCustInv.TAX2_Rate = obj.TAX2_Rate
        objCustInv.TAX2_Amt = obj.TAX2_Amt
        objCustInv.TAX3 = obj.TAX3
        objCustInv.TAX3_Rate = obj.TAX3_Rate
        objCustInv.TAX3_Amt = obj.TAX3_Amt
        objCustInv.TAX4 = obj.TAX4
        objCustInv.TAX4_Rate = obj.TAX4_Rate
        objCustInv.TAX4_Amt = obj.TAX4_Amt
        objCustInv.TAX5 = obj.TAX5
        objCustInv.TAX5_Rate = obj.TAX5_Rate
        objCustInv.TAX5_Amt = obj.TAX5_Amt
        objCustInv.TAX6 = obj.TAX6
        objCustInv.TAX6_Rate = obj.TAX6_Rate
        objCustInv.TAX6_Amt = obj.TAX6_Amt
        objCustInv.TAX7 = obj.TAX7
        objCustInv.TAX7_Rate = obj.TAX7_Rate
        objCustInv.TAX7_Amt = obj.TAX7_Amt
        objCustInv.TAX8 = obj.TAX8
        objCustInv.TAX8_Rate = obj.TAX8_Rate
        objCustInv.TAX8_Amt = obj.TAX8_Amt
        objCustInv.TAX9 = obj.TAX9
        objCustInv.TAX9_Rate = obj.TAX9_Rate
        objCustInv.TAX9_Amt = obj.TAX9_Amt
        objCustInv.TAX10 = obj.TAX10
        objCustInv.TAX10_Rate = obj.TAX10_Rate
        objCustInv.TAX10_Amt = obj.TAX10_Amt
        objCustInv.Total_Tax = obj.Total_Tax_Amt
        objCustInv.Tax1_BAmount = obj.TAX1_Base_Amt
        objCustInv.Tax2_BAmount = obj.TAX2_Base_Amt
        objCustInv.Tax3_BAmount = obj.TAX3_Base_Amt
        objCustInv.Tax4_BAmount = obj.TAX4_Base_Amt
        objCustInv.Tax5_BAmount = obj.TAX5_Base_Amt
        objCustInv.Tax6_BAmount = obj.TAX6_Base_Amt
        objCustInv.Tax7_BAmount = obj.TAX7_Base_Amt
        objCustInv.Tax8_BAmount = obj.TAX8_Base_Amt
        objCustInv.Tax9_BAmount = obj.TAX9_Base_Amt
        objCustInv.Tax10_BAmount = obj.TAX10_Base_Amt
        objCustInv.Balance_Amt = obj.doc_Amt
        objCustInv.Terms_Code = obj.Terms_Code
        objCustInv.Return_Type = "I"
        qry = "select Terms_Code,Terms_Desc,No_Days from TSPL_TERMS_MASTER where Terms_Code='" + obj.Terms_Code + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            objCustInv.Terms_Description = clsCommon.myCstr(dt.Rows(0)("Terms_Desc"))
            objCustInv.Due_Date = obj.shipment_Date.AddDays(clsCommon.myCdbl(dt.Rows(0)("No_Days")))
        End If
        objCustInv.Discount_Percentage = IIf(obj.Discount_Base = 0, 0, obj.Discount_Amt * 100 / obj.Discount_Base)
        objCustInv.Discount_Base = obj.Discount_Base
        objCustInv.Discount_Amount = obj.Discount_Amt
        ''objCustInv.Amount_Less_Discount = 
        dt = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + objCustInv.Account_Set + "'", trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            objCustInv.Customer_Control_AC = clsCommon.myCstr(dt.Rows(0)("Receivable_Control_acct"))
            If clsCommon.myCdbl(obj.Discount_Amt) > 0 Then
                objCustInv.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Receipts_Discount_acct"))
            End If
        End If

        If obj.TAX1_Amt > 0 AndAlso clsCommon.myLen(obj.TAX1) > 0 Then
            objCustInv.TAX1_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
            objCustInv.TAX1_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX1_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX2_Amt > 0 AndAlso clsCommon.myLen(obj.TAX2) > 0 Then
            objCustInv.TAX2_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
            objCustInv.TAX2_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX2_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX3_Amt > 0 AndAlso clsCommon.myLen(obj.TAX3) > 0 Then
            objCustInv.TAX3_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
            objCustInv.TAX3_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX3_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX4_Amt > 0 AndAlso clsCommon.myLen(obj.TAX4) > 0 Then
            objCustInv.TAX4_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
            objCustInv.TAX4_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX4_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX5_Amt > 0 AndAlso clsCommon.myLen(obj.TAX5) > 0 Then
            objCustInv.TAX5_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
            objCustInv.TAX5_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX5_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX6_Amt > 0 AndAlso clsCommon.myLen(obj.TAX6) > 0 Then
            objCustInv.TAX6_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
            objCustInv.TAX6_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX6_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX7_Amt > 0 AndAlso clsCommon.myLen(obj.TAX7) > 0 Then
            objCustInv.TAX7_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
            objCustInv.TAX7_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX7_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX8_Amt > 0 AndAlso clsCommon.myLen(obj.TAX8) > 0 Then
            objCustInv.TAX8_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
            objCustInv.TAX8_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX8_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX9_Amt > 0 AndAlso clsCommon.myLen(obj.TAX9) > 0 Then
            objCustInv.TAX9_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
            objCustInv.TAX9_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX9_GLAC, obj.Loc_Code, trans)
        End If
        If obj.TAX10_Amt > 0 AndAlso clsCommon.myLen(obj.TAX10) > 0 Then
            objCustInv.TAX10_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
            objCustInv.TAX10_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInv.TAX10_GLAC, obj.Loc_Code, trans)
        End If

        'objCustInv.RefDocType=
        objCustInv.RefDocNo = obj.invoice_No
        objCustInv.Add_Charge_Code1 = obj.AddCode1
        objCustInv.Add_Charge_Name1 = obj.AddDesc1
        objCustInv.Add_Charge_Amt1 = obj.AddAmt1
        objCustInv.Add_Charge_Code2 = obj.AddCode2
        objCustInv.Add_Charge_Name2 = obj.AddDesc2
        objCustInv.Add_Charge_Amt2 = obj.AddAmt2
        objCustInv.Add_Charge_Code3 = obj.AddCode3
        objCustInv.Add_Charge_Name3 = obj.AddDesc3
        objCustInv.Add_Charge_Amt3 = obj.AddAmt3
        objCustInv.Add_Charge_Code4 = obj.AddCode4
        objCustInv.Add_Charge_Name4 = obj.AddDesc4
        objCustInv.Add_Charge_Amt4 = obj.AddAmt4
        objCustInv.Add_Charge_Code5 = obj.AddCode5
        objCustInv.Add_Charge_Name5 = obj.AddDesc5
        objCustInv.Add_Charge_Amt5 = obj.AddAmt5
        objCustInv.Add_Charge_Code6 = obj.AddCode6
        objCustInv.Add_Charge_Name6 = obj.AddDesc6
        objCustInv.Add_Charge_Amt6 = obj.AddAmt6
        objCustInv.Add_Charge_Code7 = obj.AddCode7
        objCustInv.Add_Charge_Name7 = obj.AddDesc7
        objCustInv.Add_Charge_Amt7 = obj.AddAmt7
        objCustInv.Add_Charge_Code8 = obj.AddCode8
        objCustInv.Add_Charge_Name8 = obj.AddDesc8
        objCustInv.Add_Charge_Amt8 = obj.AddAmt8
        objCustInv.Add_Charge_Code9 = obj.AddCode9
        objCustInv.Add_Charge_Name9 = obj.AddDesc9
        objCustInv.Add_Charge_Amt9 = obj.AddAmt9
        objCustInv.Add_Charge_Code10 = obj.AddCode10
        objCustInv.Add_Charge_Name10 = obj.AddDesc10
        objCustInv.Add_Charge_Amt10 = obj.AddAmt10
        objCustInv.Total_Add_Charge = obj.AddAmt1 + obj.AddAmt2 + obj.AddAmt3 + obj.AddAmt5 + obj.AddAmt5 + obj.AddAmt6 + obj.AddAmt7 + obj.AddAmt8 + obj.AddAmt9 + obj.AddAmt10
        objCustInv.Tax_Calculation_Type = obj.Tax_Calculation_Type
        ''objCustInv.Status
        ''objCustInv.AgainstScrap
        objCustInv.AgainstScrapReturn = obj.Document_No
        Dim counter As Integer = 1
        objCustInv.Arr = New List(Of clsCustomerInvoiceDetail)

        '''' for return Qty
        For Each objTr As ClsScrapSaleDetailReturn In obj.Arr1

            Dim objCustInvTR As clsCustomerInvoiceDetail = New clsCustomerInvoiceDetail()
            If objTr.ItemAmt > 0 Then
                objCustInvTR.SNo = counter

                dt = clsItemMaster.GetSaleReturnAccGLAC(objTr.Item_Code, trans)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    Throw New Exception("Please set sale account for item" + objTr.Item_Code)
                End If
                objCustInvTR.GL_Account_Code = clsCommon.myCstr(dt.Rows(0)("Sales_Return_Account"))
                objCustInvTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objCustInvTR.GL_Account_Code, obj.Loc_Code, trans)
                objCustInvTR.GL_Account_Desc = clsGLAccount.GetName(objCustInvTR.GL_Account_Code, trans)
                objCustInvTR.Reco_Control_Account = "S"

                objCustInvTR.Amount = objTr.ItemAmt
                objCustInvTR.Discount_Per = objTr.DiscountPer
                objCustInvTR.Discount = objTr.DiscountAmt
                objCustInvTR.Amount_less_Discount = (objTr.ItemAmt - objTr.DiscountAmt)
                objCustInvTR.TAX1 = objTr.TAX1
                objCustInvTR.TAX1_Rate = objTr.TAX1_Rate
                objCustInvTR.TAX1_Amt = objTr.TAX1_Amt
                objCustInvTR.TAX1_Base_Amt = objTr.TAX1_Base_Amt
                objCustInvTR.TAX2 = objTr.TAX2
                objCustInvTR.TAX2_Rate = objTr.TAX2_Rate
                objCustInvTR.TAX2_Amt = objTr.TAX2_Amt
                objCustInvTR.TAX2_Base_Amt = objTr.TAX2_Base_Amt
                objCustInvTR.TAX3 = objTr.TAX3
                objCustInvTR.TAX3_Rate = objTr.TAX3_Rate
                objCustInvTR.TAX3_Amt = objTr.TAX3_Amt
                objCustInvTR.TAX3_Base_Amt = objTr.TAX3_Base_Amt
                objCustInvTR.TAX4 = objTr.TAX4
                objCustInvTR.TAX4_Rate = objTr.TAX4_Rate
                objCustInvTR.TAX4_Amt = objTr.TAX4_Amt
                objCustInvTR.TAX4_Base_Amt = objTr.TAX4_Base_Amt
                objCustInvTR.TAX5 = objTr.TAX5
                objCustInvTR.TAX5_Rate = objTr.TAX5_Rate
                objCustInvTR.TAX5_Amt = objTr.TAX5_Amt
                objCustInvTR.TAX5_Base_Amt = objTr.TAX5_Base_Amt
                objCustInvTR.TAX6 = objTr.TAX6
                objCustInvTR.TAX6_Rate = objTr.TAX6_Rate
                objCustInvTR.TAX6_Amt = objTr.TAX6_Amt
                objCustInvTR.TAX6_Base_Amt = objTr.TAX6_Base_Amt
                objCustInvTR.TAX7 = objTr.TAX7
                objCustInvTR.TAX7_Rate = objTr.TAX7_Rate
                objCustInvTR.TAX7_Amt = objTr.TAX7_Amt
                objCustInvTR.TAX7_Base_Amt = objTr.TAX7_Base_Amt
                objCustInvTR.TAX8 = objTr.TAX8
                objCustInvTR.TAX8_Rate = objTr.TAX8_Rate
                objCustInvTR.TAX8_Amt = objTr.TAX8_Amt
                objCustInvTR.TAX8_Base_Amt = objTr.TAX8_Base_Amt
                objCustInvTR.TAX9 = objTr.TAX9
                objCustInvTR.TAX9_Rate = objTr.TAX9_Rate
                objCustInvTR.TAX9_Amt = objTr.TAX9_Amt
                objCustInvTR.TAX9_Base_Amt = objTr.TAX9_Base_Amt
                objCustInvTR.TAX10 = objTr.TAX10
                objCustInvTR.TAX10_Rate = objTr.TAX10_Rate
                objCustInvTR.TAX10_Amt = objTr.TAX10_Amt
                objCustInvTR.TAX10_Base_Amt = objTr.TAX10_Base_Amt
                objCustInvTR.Total_Tax = objTr.TotalTaxAmt
                objCustInvTR.Total_Amount = objTr.TotalAmt
                objCustInvTR.Remarks = objTr.Description
                'objCustInvTR.Comments=objTr.Comments
                objCustInv.Arr.Add(objCustInvTR)
                counter += 1
            End If

        Next

        ''richa BHA/01/03/19-000830 send form id 06 March,2019
        objCustInv.SaveData(objCustInv, True, trans, "", strVoucherForRecreate, strARNoForRecreate)
        clsCustomerInvoiceHead.PostData("", objCustInv.Document_No, "", trans)
        Return True
    End Function

    Public Shared Function DeleteData(ByVal strCode As String) As Boolean
        Dim isSaved As Boolean = False
        If (clsCommon.myLen(strCode) <= 0) Then
            Throw New Exception("ShipmentNo not found to Delete")
        End If
        Dim obj As ClsScrapSaleHeadReturn = ClsScrapSaleHeadReturn.GetData(strCode, NavigatorType.Current)
        Dim frm As New FrmFreeTxtBox1
        frm.Text = "Remarks for Delete"
        frm.ShowDialog()
        If clsCommon.myLen(frm.strRmks) <= 0 Then
            Return False
        End If


        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.shipment_No) > 0) Then
            Try
                '  clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Purchase Order", "Scrap Invoice", obj.Loc_Code, obj.shipment_Date, trans)

                If (obj.ispost = 1) Then
                    Throw New Exception("Already Posted on :" + obj.posting_Date)
                End If
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strCode), "TSPL_SCRAPSALE_HEAD_Return", "Document_No", "TSPL_SCRAPSALE_DETAIL_RETURN", "Document_No", trans)

                HistoryUpdate(strCode, trans)
                Dim qry As String = "delete from TSPL_SCRAPSALE_DETAIL_RETURN where Document_No='" + strCode + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)

                qry = "delete from tspl_scrapsale_head_RETURN where Document_No='" + strCode + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)


                isSaved = isSaved AndAlso clsCustomFieldValues.DeleteData(obj.Form_ID, strCode, trans)

                clsBatchInventory.DeleteData("MS-SR", strCode, trans)

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
End Class
Public Class ClsScrapSaleDetailReturn
#Region "Variables"
    Public ItemwiseTaxCode As String = Nothing
    Public document_No As String = Nothing
    Public shipment_No As String = Nothing
    Public Specification As String = Nothing
    Public Line_No As Integer = 0
    Public Item_Code As String = Nothing
    Public Item_Desc As String = Nothing
    Public shipped_Qty As Double = 0
    'sanjay
    Public shQty As Double = 0
    'Public PendingQty As Double = 0
    'sanjay
    Public price As Double = 0
    Public balance_Qty As Double = 0
    Public DiscountPer As Double = 0
    Public DiscountAmt As Double = 0
    Public Tax As Double = 0
    Public NetPriceAmt As String = Nothing
    Public ItemAmt As String = Nothing
    Public TotalDiscountAmt As Double = 0
    Public TotalTaxAmt As Double = 0
    Public ItemNetAmt As Double = 0
    Public TotalAmt As Double = 0 '
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
    Public Code As String = Nothing   'Additiona charges code
    Public Description As String = Nothing   'Additiona charges description
    Public invoice_Qty As Double = 0
    Public pending_Qty As Double = 0
    Public Unit_Code As String = ""
    Public Asset_Code As String = ""

    Public FAT As Double = 0
    Public SNF As Double = 0

    Public arrBatchItem As List(Of clsBatchInventory) = Nothing

#End Region



    Shared Function SaveData(Document_No As String, Arr1 As List(Of ClsScrapSaleDetailReturn), trans As SqlTransaction, shipment_Date As Date, Loc_Code As String, strInvoiceNo As String) As Boolean
        If (Arr1 IsNot Nothing AndAlso Arr1.Count > 0) Then
            For Each obj As ClsScrapSaleDetailReturn In Arr1
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Document_No", Document_No)
                clsCommon.AddColumnsForChange(coll, "shipment_No", obj.shipment_No)
                clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Item_Desc", obj.Item_Desc)
                clsCommon.AddColumnsForChange(coll, "Unit_Code", obj.Unit_Code)
                clsCommon.AddColumnsForChange(coll, "shipped_Qty", obj.shipped_Qty)
                clsCommon.AddColumnsForChange(coll, "price", obj.price)
                clsCommon.AddColumnsForChange(coll, "DiscountPer", obj.DiscountPer)
                clsCommon.AddColumnsForChange(coll, "DiscountAmt", obj.DiscountAmt)
                clsCommon.AddColumnsForChange(coll, "TotalTaxAmt", obj.TotalTaxAmt)
                clsCommon.AddColumnsForChange(coll, "ItemAmt", obj.ItemAmt)
                clsCommon.AddColumnsForChange(coll, "NetPriceAmt", obj.NetPriceAmt)
                clsCommon.AddColumnsForChange(coll, "TotalDiscountAmt", obj.TotalDiscountAmt)
                clsCommon.AddColumnsForChange(coll, "ItemNetAmt", obj.ItemNetAmt)
                clsCommon.AddColumnsForChange(coll, "TotalAmt", obj.TotalAmt)
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
                clsCommon.AddColumnsForChange(coll, "pending_Qty", obj.pending_Qty)
                clsCommon.AddColumnsForChange(coll, "Specification", obj.Specification)
                clsCommon.AddColumnsForChange(coll, "Asset_Code", obj.Asset_Code, True)

                clsCommon.AddColumnsForChange(coll, "FAT", obj.FAT)
                clsCommon.AddColumnsForChange(coll, "SNF", obj.SNF)
                clsCommon.AddColumnsForChange(coll, "ItemwiseTaxCode", obj.ItemwiseTaxCode)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SCRAPSALE_DETAIL_Return", OMInsertOrUpdate.Insert, "", trans)
                clsBatchInventory.SaveData("MS-SR", Document_No, shipment_Date, "I", obj.Item_Code, Loc_Code, obj.Line_No, 0, obj.Unit_Code, obj.arrBatchItem, trans)


            Next
        End If
        Return True
    End Function
    '' ScrapInvoice Post Entry
    Public Shared Function PostData(ByVal strDocNo As String, ByVal IsCreateInventory As Boolean, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreateOnly As String = Nothing, Optional ByVal strArInvNoForRecreateOnly As String = Nothing, Optional strARVoucherNoForRecreatedOnly As String = Nothing, Optional strReturnInvoice As String = Nothing) As Boolean
        Try
            Dim isSaved As Boolean = True
            Dim istrue As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Invoice No not found to Post")
            End If
            Dim obj As ClsScrapSaleHeadReturn = ClsScrapSaleHeadReturn.GetData(strDocNo, NavigatorType.Current, trans, False)

            If (obj Is Nothing OrElse clsCommon.myLen(obj.invoice_No) <= 0) Then
                Throw New Exception("No Data found to Post")
            End If
            If (obj.ispost = 1) Then
                Throw New Exception("Already Post on :" + obj.posting_Date)
            End If
            If (obj.Status = 1) Then
                Throw New Exception("Invoice No " + obj.invoice_No + " Is On Hold.Can't Post it")
            End If

            Dim qry As String = ""
            Dim ArrInventoryMovement As List(Of clsInventoryMovement) = New List(Of clsInventoryMovement)
            Dim IsRejectedItemFound As Boolean = False



            For Each objTr As ClsScrapSaleDetailReturn In obj.Arr1
                '--------------------------------------------------------------\
                Dim objInventoryMovemnt As New clsInventoryMovement()
                objInventoryMovemnt.InOut = "I"
                objInventoryMovemnt.Location_Code = obj.Loc_Code
                objInventoryMovemnt.Cust_Code = obj.cust_Code
                objInventoryMovemnt.Cust_Name = obj.cust_Name
                objInventoryMovemnt.Item_Code = objTr.Item_Code
                objInventoryMovemnt.Item_Desc = objTr.Item_Desc
                objInventoryMovemnt.Qty = objTr.shipped_Qty
                objInventoryMovemnt.UOM = objTr.Unit_Code
                objInventoryMovemnt.Basic_Cost = objTr.price
                objInventoryMovemnt.Add_Cost = objTr.TotalTaxAmt
                objInventoryMovemnt.Net_Cost = objTr.TotalAmt
                objInventoryMovemnt.ItemType = "RM"
                ArrInventoryMovement.Add(objInventoryMovemnt)
                '------------------------------------------------------------=-=

                'sanjay
                'Dim qryupdatepen As String = "update TSPL_SCRAPSALE_DETAIL_Return set pending_qty='" + clsCommon.myCstr(objTr.pending_Qty) + "' where Document_no='" + obj.shipment_No + "'  and item_code='" + objTr.Item_Code + "' "
                'clsDBFuncationality.ExecuteNonQuery(qryupdatepen, trans)
                'sanjay
            Next


            If IsCreateInventory Then
                isSaved = isSaved AndAlso clsInventoryMovement.SaveData("ScrapReturn", obj.invoice_No, clsCommon.GetPrintDate(obj.shipment_Date, "dd/MM/yyyy"), clsCommon.GetPrintDate(obj.shipment_Date, "dd/MM/yyyy"), ArrInventoryMovement, trans)
            End If

            If istrue = False Then
                Return False
            End If

            Dim strRMDANo As String = ""
            qry = "Update TSPL_SCRAPINVOICE_HEAD set ispost=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt ") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
            qry += " where invoice_No='" + strDocNo + "'"
            isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)



            '---------------------For AR-Invoice Entry-----------------------------

            Dim Auto_Gen_No As Boolean = True

            Dim objCust As New clsCustomerInvoiceHead()
            ' objCust.Document_No = "'"
            If strArInvNoForRecreateOnly IsNot Nothing AndAlso clsCommon.myLen(strArInvNoForRecreateOnly) > 0 Then
                Auto_Gen_No = False
                objCust.Document_No = strArInvNoForRecreateOnly
            End If
            objCust.Document_Date = obj.shipment_Date
            objCust.Customer_Code = obj.cust_Code
            objCust.Customer_Name = obj.cust_Name
            objCust.loc_code = obj.Loc_Code
            objCust.Posting_Date = obj.posting_Date
            objCust.Account_Set = clsDBFuncationality.getSingleValue("select Cust_Account from TSPL_CUSTOMER_MASTER where Cust_Code='" + obj.cust_Code + "'", trans)
            If clsCommon.CompairString(obj.Doc_Type, "J") = CompairStringResult.Equal Then
                objCust.Document_Type = "D"
            Else
                objCust.Document_Type = "C"
            End If



            objCust.Order_No = ""
            objCust.Document_Total = obj.doc_Amt
            objCust.RoundOffAmount = obj.RoundOffAmount
            objCust.On_Hold = "0"
            objCust.Remarks = ""
            objCust.Description = ""
            objCust.Tax_Group = obj.Tax_Group
            objCust.RefDocType = ""
            objCust.RefDocNo = ""

            objCust.TAX1 = obj.TAX1
            objCust.TAX1_Rate = obj.TAX1_Rate
            objCust.Tax1_BAmount = obj.TAX1_Base_Amt
            objCust.TAX1_Amt = obj.TAX1_Amt

            objCust.TAX2 = obj.TAX2
            objCust.TAX2_Rate = obj.TAX2_Rate
            objCust.Tax2_BAmount = obj.TAX2_Base_Amt
            objCust.TAX2_Amt = obj.TAX2_Amt

            objCust.TAX3 = obj.TAX3
            objCust.TAX3_Rate = obj.TAX3_Rate
            objCust.Tax3_BAmount = obj.TAX3_Base_Amt
            objCust.TAX3_Amt = obj.TAX3_Amt

            objCust.TAX4 = obj.TAX4
            objCust.TAX4_Rate = obj.TAX4_Rate
            objCust.Tax4_BAmount = obj.TAX4_Base_Amt
            objCust.TAX4_Amt = obj.TAX4_Amt

            objCust.TAX5 = obj.TAX5
            objCust.TAX5_Rate = obj.TAX5_Rate
            objCust.Tax5_BAmount = obj.TAX5_Base_Amt
            objCust.TAX5_Amt = obj.TAX5_Amt

            objCust.TAX6 = obj.TAX6
            objCust.TAX6_Rate = obj.TAX6_Rate
            objCust.Tax6_BAmount = obj.TAX6_Base_Amt
            objCust.TAX6_Amt = obj.TAX6_Amt

            objCust.TAX7 = obj.TAX7
            objCust.TAX7_Rate = obj.TAX7_Rate
            objCust.Tax7_BAmount = obj.TAX7_Base_Amt
            objCust.TAX7_Amt = obj.TAX7_Amt

            objCust.TAX8 = obj.TAX8
            objCust.TAX8_Rate = obj.TAX8_Rate
            objCust.Tax8_BAmount = obj.TAX8_Base_Amt
            objCust.TAX8_Amt = obj.TAX8_Amt

            objCust.TAX9 = obj.TAX9
            objCust.TAX9_Rate = obj.TAX9_Rate
            objCust.Tax9_BAmount = obj.TAX9_Base_Amt
            objCust.TAX9_Amt = obj.TAX9_Amt

            objCust.TAX10 = obj.TAX10
            objCust.TAX10_Rate = obj.TAX10_Rate
            objCust.Tax10_BAmount = obj.TAX10_Base_Amt
            objCust.TAX10_Amt = obj.TAX10_Amt

            objCust.Total_Tax = obj.Total_Tax_Amt
            objCust.Terms_Code = obj.Terms_Code
            objCust.Terms_Description = clsDBFuncationality.getSingleValue("select Terms_Desc from TSPL_TERMS_MASTER where Terms_Code='" + obj.Terms_Code + "'", trans)
            objCust.Due_Date = obj.Due_Date
            objCust.Discount_Percentage = 0
            objCust.Discount_Base = obj.Discount_Base
            objCust.Discount_Amount = obj.Discount_Amt
            objCust.Amount_Less_Discount = obj.Amount_Less_Discount
            objCust.Comp_Code = obj.Comp_Code
            objCust.Balance_Amt = obj.doc_Amt

            ''richa agarwal 25/06/2015 change location of account against BM00000007177
            '  objCust.Customer_Control_AC = clsDBFuncationality.getSingleValue("select Receivable_Control_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.cust_Code + "'", trans)
            Dim strCustomerCntrlAcc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Receivable_Control_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.cust_Code + "'", trans))
            objCust.Customer_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strCustomerCntrlAcc, obj.Loc_Code, trans)
            '--------------------------------
            If obj.Discount_Amt <> 0 Then
                ''richa agarwal 25/06/2015 change location of account against BM00000007177
                'objCust.Discount_GL_AC = clsDBFuncationality.getSingleValue("select Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.cust_Code + "'", trans)
                Dim strDiscountGLAcc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.cust_Code + "'", trans))
                objCust.Customer_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strDiscountGLAcc, obj.Loc_Code, trans)
                ''-------------------------------------
            End If
            objCust.TAX1_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX1, trans)
            objCust.TAX2_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX2, trans)
            objCust.TAX3_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX3, trans)
            objCust.TAX4_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX4, trans)
            objCust.TAX5_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX5, trans)
            objCust.TAX6_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX6, trans)
            objCust.TAX7_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX7, trans)
            objCust.TAX8_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX8, trans)
            objCust.TAX9_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX9, trans)
            objCust.TAX10_GLAC = clsTaxMaster.GetTaxPayAC(obj.TAX10, trans)
            objCust.Add_Charge_Code1 = obj.AddCode1
            objCust.Add_Charge_Name1 = obj.AddDesc1
            objCust.Add_Charge_Amt1 = obj.AddAmt1

            objCust.Add_Charge_Code2 = obj.AddCode2
            objCust.Add_Charge_Name2 = obj.AddDesc2
            objCust.Add_Charge_Amt2 = obj.AddAmt2

            objCust.Add_Charge_Code3 = obj.AddCode3
            objCust.Add_Charge_Name3 = obj.AddDesc3
            objCust.Add_Charge_Amt3 = obj.AddAmt3

            objCust.Add_Charge_Code4 = obj.AddCode4
            objCust.Add_Charge_Name4 = obj.AddDesc4
            objCust.Add_Charge_Amt4 = obj.AddAmt4

            objCust.Add_Charge_Code5 = obj.AddCode5
            objCust.Add_Charge_Name5 = obj.AddDesc5
            objCust.Add_Charge_Amt5 = obj.AddAmt5

            objCust.Add_Charge_Code6 = obj.AddCode6
            objCust.Add_Charge_Name6 = obj.AddDesc6
            objCust.Add_Charge_Amt6 = obj.AddAmt6

            objCust.Add_Charge_Code7 = obj.AddCode7
            objCust.Add_Charge_Name7 = obj.AddDesc7
            objCust.Add_Charge_Amt7 = obj.AddAmt7

            objCust.Add_Charge_Code8 = obj.AddCode8
            objCust.Add_Charge_Name8 = obj.AddDesc8
            objCust.Add_Charge_Amt8 = obj.AddAmt8

            objCust.Add_Charge_Code9 = obj.AddCode9
            objCust.Add_Charge_Name9 = obj.AddDesc9
            objCust.Add_Charge_Amt9 = obj.AddAmt9

            objCust.Add_Charge_Code10 = obj.AddCode10
            objCust.Add_Charge_Name10 = obj.AddDesc10
            objCust.Add_Charge_Amt10 = obj.AddAmt10

            objCust.Total_Add_Charge = obj.Add_Amt
            objCust.Balance_Amt = obj.doc_Amt
            objCust.Tax_Calculation_Type = "0"
            objCust.AgainstScrapReturn = strReturnInvoice

            objCust.RefDocNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select document_No from tspl_customer_invoice_head where AgainstScrapReturn='" & strReturnInvoice & "'", trans))


            objCust.Arr = New List(Of clsCustomerInvoiceDetail)
            For Each objout As ClsScrapSaleDetailReturn In obj.Arr1
                Dim objtr As New clsCustomerInvoiceDetail()
                'objtr.invoice_No = invoice
                objtr.SNo = objout.Line_No
                ''richa agarwal 25/06/2015 change location of account against BM00000007177
                'objtr.GL_Account_Code = clsDBFuncationality.getSingleValue("select TSPL_SALES_ACCOUNTS.Sales_Account  from TSPL_SALES_ACCOUNTS left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Sale_Class_Code=TSPL_SALES_ACCOUNTS.Sales_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objout.Item_Code + "'", trans)
                'objtr.GL_Account_Desc = clsDBFuncationality.getSingleValue("select TSPL_SALES_ACCOUNTS.Sales_Class_Desc  from TSPL_SALES_ACCOUNTS left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Sale_Class_Code=TSPL_SALES_ACCOUNTS.Sales_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objout.Item_Code + "'", trans)
                Dim strGLAcc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_SALES_ACCOUNTS.Sales_Account  from TSPL_SALES_ACCOUNTS left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Sale_Class_Code=TSPL_SALES_ACCOUNTS.Sales_Class_Code where TSPL_ITEM_MASTER.Item_Code='" + objout.Item_Code + "'", trans))
                objtr.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(strGLAcc, obj.Loc_Code, trans)
                objtr.GL_Account_Desc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_GL_ACCOUNTS where Account_Code='" + clsCommon.myCstr(objtr.GL_Account_Code) + "'", trans))
                ''-------------------------------------
                objtr.Amount = objout.ItemAmt
                objtr.Discount_Per = objout.DiscountPer
                objtr.Discount = objout.DiscountAmt
                objtr.Amount_less_Discount = objout.ItemNetAmt

                objtr.TAX1 = objout.TAX1
                objtr.TAX1_Rate = objout.TAX1_Rate
                objtr.TAX1_Amt = objout.TAX1_Amt
                objtr.TAX1_Base_Amt = objout.TAX1_Base_Amt


                objtr.TAX2 = objout.TAX2
                objtr.TAX2_Rate = objout.TAX2_Rate
                objtr.TAX2_Amt = objout.TAX2_Amt
                objtr.TAX2_Base_Amt = objout.TAX2_Base_Amt

                objtr.TAX3 = objout.TAX3
                objtr.TAX3_Rate = objout.TAX3_Rate
                objtr.TAX3_Amt = objout.TAX3_Amt
                objtr.TAX3_Base_Amt = objout.TAX3_Base_Amt

                objtr.TAX4 = objout.TAX4
                objtr.TAX4_Rate = objout.TAX4_Rate
                objtr.TAX4_Amt = objout.TAX4_Amt
                objtr.TAX4_Base_Amt = objout.TAX4_Base_Amt

                objtr.TAX5 = objout.TAX5
                objtr.TAX5_Rate = objout.TAX5_Rate
                objtr.TAX5_Amt = objout.TAX5_Amt
                objtr.TAX5_Base_Amt = objout.TAX4_Base_Amt

                objtr.TAX6 = objout.TAX6
                objtr.TAX6_Rate = objout.TAX6_Rate
                objtr.TAX6_Amt = objout.TAX6_Amt
                objtr.TAX6_Base_Amt = objout.TAX6_Base_Amt

                objtr.TAX7 = objout.TAX7
                objtr.TAX7_Rate = objout.TAX7_Rate
                objtr.TAX7_Amt = objout.TAX7_Amt
                objtr.TAX7_Base_Amt = objout.TAX7_Base_Amt

                objtr.TAX8 = objout.TAX8
                objtr.TAX8_Rate = objout.TAX8_Rate
                objtr.TAX8_Amt = objout.TAX8_Amt
                objtr.TAX8_Base_Amt = objout.TAX8_Base_Amt

                objtr.TAX9 = objout.TAX9
                objtr.TAX9_Rate = objout.TAX9_Rate
                objtr.TAX9_Amt = objout.TAX9_Amt
                objtr.TAX9_Base_Amt = objout.TAX9_Base_Amt

                objtr.TAX10 = objout.TAX10
                objtr.TAX10_Rate = objout.TAX10_Rate
                objtr.TAX10_Amt = objout.TAX10_Amt
                objtr.TAX10_Base_Amt = objout.TAX10_Base_Amt

                objtr.Total_Tax = objout.TotalTaxAmt
                objtr.Total_Amount = objout.TotalAmt
                objtr.Remarks = ""
                objtr.Comments = ""
                objCust.Arr.Add(objtr)
            Next

            ''richa agarwal 08/07/2015 BM00000007330
            'If objCust.SaveData(objCust, Auto_Gen_No, trans, "", strARVoucherNoForRecreatedOnly) Then
            '    qry = "Update TSPL_Customer_Invoice_Head set status=1, Posting_Date='" + clsCommon.GetPrintDate(obj.shipment_Date, "dd/MMM/yyyy hh:mm tt ") + "',Modify_By='" + objCommonVar.CurrentUserCode + "'"
            '    qry += " where AgainstScrap='" + obj.invoice_No + "'"

            'End If
            objCust.SaveData(objCust, Auto_Gen_No, trans, "", strARVoucherNoForRecreatedOnly)
            clsCustomerInvoiceHead.PostData("", objCust.Document_No, "", trans)
            ''------------------------------
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class
