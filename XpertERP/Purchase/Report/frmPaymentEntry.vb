'Developed By -BibhuPrasad Parida
'Database - TSPLERP
'Table - tspl_customer_category_master
'Start Date -
'End Date -

''''13/10/2011---Updation by --[Pankaj Kumar Chaudhary]-- in Report Query (Added New Fields [Add1, Add2, Add3, Add4, City_Code, State From Location_Master])
''''07/06/2012---Updation By --[Pankaj Kumar]--In Report Printing the Vendor name should be printed with Vendor Code----Chenge In Print Query
Imports XpertERPEngine
Imports System.Data
Imports System.Data.SqlClient
Imports common
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine


Public Class FrmPaymentEntry
    Inherits FrmMainTranScreen
#Region "variables"
    Dim dr As SqlDataReader
    Dim ds As DataSet
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Public Shared strCode As String = ""
#End Region
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            Dim StartDate As String = dtpFromDate.Value
            Dim EndDate As String = dtpToDate.Value

            Dim ArrDoc As New ArrayList
            If ChkDocumentSelect.IsChecked And cbgDocument.CheckedValue.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select atleast single Document or select all.", Me.Text)
                Exit Sub
            Else
                ArrDoc = cbgDocument.CheckedValue
            End If

            Dim ArrVendor As New ArrayList
            If chkVendorSelect.IsChecked And cbgVendor.CheckedValue.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select atleast single Vendor or select all.", Me.Text)
                Exit Sub
            Else
                ArrVendor = cbgVendor.CheckedValue
            End If

            Dim Arrlocation As New ArrayList
            If chkLocSelect.IsChecked And cbgLocation.CheckedValue.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select atleast single Location or select all.", Me.Text)
                Exit Sub
            Else
                Arrlocation = cbgLocation.CheckedValue
            End If

            funReport(StartDate, EndDate, ArrDoc, ArrVendor, Arrlocation)
        Catch ex As Exception

        End Try
        
    End Sub
    Private Sub funPrint()
        Try
            Dim LocCode As String = String.Empty
            Dim Vendor As String = String.Empty
            Dim DocNo As String = String.Empty
            Dim dttemp As New DataTable()
            Dim fromdate As String = clsCommon.GetPrintDate(dtpFromDate.Value, "dd/MM/yyyy")
            Dim todate As String = clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy")
            If cbgDocument.CheckedValue.Count > 0 OrElse cbgVendor.CheckedValue.Count > 0 Then
                fromdate = ""
                todate = ""
            End If
            Dim arrDocument As ArrayList = Nothing
            Dim arrVendor As ArrayList = Nothing
            Dim arrlocation As ArrayList = Nothing
            If ChkDocumentSelect.IsChecked AndAlso cbgDocument.CheckedValue.Count <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please Select at least one Document.", Me.Text)
                Return
            ElseIf (ChkDocumentSelect.IsChecked) Then
                arrDocument = cbgDocument.CheckedValue
                DocNo = clsCommon.GetMulcallString(cbgDocument.CheckedValue)
                DocNo = DocNo.Replace("'", "")
            ElseIf Not (ChkDocumentSelect.IsChecked) Then
                arrDocument = cbgDocument.AllValue
            End If
            If chkVendorSelect.IsChecked AndAlso cbgVendor.CheckedValue.Count <= 0 Then
                common.clsCommon.MyMessageBoxShow("Please Select at least one Vendor.", Me.Text)
                Return
            ElseIf (chkVendorSelect.IsChecked) Then
                arrVendor = cbgVendor.CheckedValue
                Vendor = clsCommon.GetMulcallString(cbgVendor.CheckedValue)
                Vendor = Vendor.Replace("'", "")
            End If
            If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count <= 0 Then
                common.clsCommon.MyMessageBoxShow("Please Select at least one Location.", Me.Text)
                Return
            ElseIf (chkLocSelect.IsChecked) Then
                arrlocation = cbgLocation.CheckedValue
                LocCode = clsCommon.GetMulcallString(cbgLocation.CheckedValue)
                LocCode = LocCode.Replace("'", "")
            End If
            Dim qry As String = "select Payment_No   from TSPL_PAYMENT_HEADER  where 2=2"
            If arrDocument IsNot Nothing AndAlso arrDocument.Count > 0 Then
                qry += " and Payment_No in (" + clsCommon.GetMulcallString(arrDocument) + ")"
            End If
            If arrVendor IsNot Nothing AndAlso arrVendor.Count > 0 Then
                qry += "and Vendor_Code in (" + clsCommon.GetMulcallString(arrVendor) + ")"
            End If

            Dim dtdocument As DataTable = clsDBFuncationality.GetDataTable(qry)
            dttemp.Columns.Add("FromDate")
            dttemp.Columns.Add("ToDate")
            dttemp.Columns.Add("Location", GetType(System.String))
            dttemp.Columns.Add("Vendor", GetType(System.String))
            dttemp.Columns.Add("Document", GetType(System.String))
            dttemp.Columns.Add("PaymentNo", GetType(System.String))
            dttemp.Columns.Add("PaymentDate")
            dttemp.Columns.Add("VendorNameRemitTO", GetType(System.String))
            dttemp.Columns.Add("Cheque_No", GetType(System.String))
            dttemp.Columns.Add("VoucherType", GetType(System.String))
            dttemp.Columns.Add("Payment_Code", GetType(System.String))
            dttemp.Columns.Add("BANK_CODE", GetType(System.String))
            dttemp.Columns.Add("BankName", GetType(System.String))
            dttemp.Columns.Add("Bank_acct")
            dttemp.Columns.Add("Compaddress1", GetType(System.String))
            dttemp.Columns.Add("CreatedBy", GetType(System.String))
            dttemp.Columns.Add("AuthorisedBy", GetType(System.String))
            dttemp.Columns.Add("CompanyCode", GetType(System.String))
            dttemp.Columns.Add("Logo_Img")
            dttemp.Columns.Add("Logo_Img2")
            dttemp.Columns.Add("Payment_Type", GetType(System.String))
            dttemp.Columns.Add("AccountCode", GetType(System.String))
            dttemp.Columns.Add("AccountName", GetType(System.String))
            dttemp.Columns.Add("Entry_Desc", GetType(System.String))
            dttemp.Columns.Add("Remarks", GetType(System.String))
            dttemp.Columns.Add("Comment", GetType(System.String))
            dttemp.Columns.Add("Debit", GetType(System.Decimal))
            dttemp.Columns.Add("Credit", GetType(System.Decimal))
            dttemp.Columns.Add("SegName", GetType(System.String))
            dttemp.Columns.Add("saleinvoiceno", GetType(System.String))
            dttemp.Columns.Add("salesmancode", GetType(System.String))
            dttemp.Columns.Add("salesmanname", GetType(System.String))
            dttemp.Columns.Add("txtno", GetType(System.String))
            dttemp.Columns.Add("EntryDesc", GetType(System.String))

            Dim dt As DataTable
            Dim recvalue As String
            For Each drreceipt As DataRow In dtdocument.Rows
                recvalue = drreceipt("Payment_No").ToString()
                'Dim qryParticual = "select max(EntryDesc) from("
                '' qryParticual += "select TSPL_PAYMENT_DETAIL.Vendor_Invoice_No+ case when len(RTRIM(TSPL_PAYMENT_DETAIL.Vendor_Invoice_No))>0 then ', ' else ' ' end +Vendor_Invoice_Date  as EntryDesc from tspl_payment_header join TSPL_PAYMENT_DETAIL on tspl_payment_header.Payment_No =TSPL_PAYMENT_DETAIL.Payment_No  join TSPL_VENDOR_INVOICE_HEAD  on TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_PAYMENT_DETAIL.Document_No where tspl_payment_header.Payment_No = Final.Payment_No"
                'qryParticual += "select TSPL_PAYMENT_DETAIL.Vendor_Invoice_No+ case when len(RTRIM(TSPL_PAYMENT_DETAIL.Vendor_Invoice_No))>0 then ', ' else ' ' end +Vendor_Invoice_Date  as EntryDesc from tspl_payment_header join TSPL_PAYMENT_DETAIL on tspl_payment_header.Payment_No =TSPL_PAYMENT_DETAIL.Payment_No  join TSPL_VENDOR_INVOICE_HEAD  on TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_PAYMENT_DETAIL.Document_No where tspl_payment_header.Payment_No in (" + clsCommon.GetMulcallString(ArrDocument) + ")"
                'qryParticual += ") xxx"
                qry = "select  '" + dtpFromDate.Value + "' as FromDate,'" + dtpToDate.Value + "' as ToDate, '" + LocCode + "' as Location,'" + Vendor + "' as Vendor,'" + DocNo + "'as Document,Final.Payment_No as PaymentNo,Payment_Date as PaymentDate,case when Vendor_Name =Remit_To then Vendor_Name else case when len(isnull(Vendor_Name,''))>0 and len(isnull(Remit_To,''))>0 then Vendor_Name+' / '+Remit_To else Vendor_Name+''+ Remit_To   end end  as VendorNameRemitTO,Cheque_No ,case when  Payment_Code ='SETTLEMENT' then 'Settlement Payment Voucher' else CASE WHEN Payment_Code ='Cheque' then 'Bank Payment Voucher' else CASE WHEN Payment_Code ='Cash' then 'Cash Payment Voucher'  ELSE CASE WHEN tspl_payment_header.Payment_Code=(select Bank_Code  from TSPL_PAYMENT_HEADER WHERE tspl_payment_header.Payment_No ='" + recvalue + "' )THEN  (select Bank_Code  from TSPL_PAYMENT_HEADER WHERE Payment_No ='" + recvalue + "')+' Voucher'" & _
                      " END END end end  as [VoucherType], tspl_payment_header.Payment_Code,tspl_payment_header.BANK_CODE,TSPL_BANK_MASTER.DESCRIPTION as BankName, (select BANKACC from TSPL_BANK_MASTER where Bank_Code =tspl_payment_header.BANK_CODE) as Bank_acct,(select max(ADD1 + case when len(add2)> 0 then ',' else '' end + ADD2 +case when len(add3)> 0 then ','else '' end +ADD3+case when len(add4)> 0 then ',' else '' end +ADD4+case when len(City_Code)> 0 then ',' else '' end +City_Code +case when len(STATE)> 0 then ',' else '' end  +STATE) from tspl_location_master where loc_segment_code =(substring (TSPL_BANK_MASTER .BANKACC ,LEN(TSPL_BANK_MASTER .BANKACC)-2,3))and TSPL_LOCATION_MASTER .Location_Type ='Physical')as Compaddress1, TSPL_USER_MASTER.User_Name as CreatedBy,Case when (TSPL_PAYMENT_HEADER.Posted='1' or TSPL_PAYMENT_HEADER.Posted='P') then tspl_payment_header.Modify_By when (TSPL_PAYMENT_HEADER.Posted='0' or TSPL_PAYMENT_HEADER.Posted='N') then null end as AuthorisedBy,TSPL_COMPANY_MASTER.Comp_Name as CompanyCode,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2, TSPL_PAYMENT_HEADER.Payment_Type , final.ACCode as AccountCode,final.Description as AccountName,Final .Entry_Desc,final.Remarks,final.Comment ,Final.DrAmt as Debit,final.CrAmt as Credit ,(select Description from TSPL_GL_SEGMENT_CODE where  Seg_No='7' and Segment_code=substring (final.ACCode,LEN(final.ACCode)-2,3)) as SegName,Final .saleinvoiceno ,Final .salesmancode ,Final .salesmanname,final.txtno,final.EntryDesc from ( "
                qry += " select tspl_payment_header.Payment_No, TSPL_BANK_MASTER.BANKACC as ACCode,TSPL_GL_ACCOUNTS.Description,tspl_payment_header.Entry_Desc,'' as Remarks,'' as Comment,0 as DrAmt,Payment_Amount as CrAmt,1 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc  from tspl_payment_header  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=TSPL_PAYMENT_HEADER.Location_Code  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BANK_MASTER.BANKACC      "
                qry += "where  exists (select 1 from TSPL_LOCATION_MASTER where TSPL_LOCATION_MASTER .Loc_Segment_Code=substring(TSPL_BANK_MASTER .BANKACC ,len(TSPL_BANK_MASTER .BANKACC)-2,3) "

                If arrlocation IsNot Nothing AndAlso arrlocation.Count > 0 Then
                    qry += " and  TSPL_LOCATION_MASTER.Loc_Segment_Code  in (" + clsCommon.GetMulcallString(arrlocation) + ")"

                End If
                qry += " )"
                qry += " union all"
                qry += " select tspl_payment_header.Payment_No,Debit_Account as ACCode, TSPL_GL_ACCOUNTS.Description,'' as Entry_Desc,'' as Remarks,'' as Comment,TSPL_PAYMENT_HEADER.Payment_Amount as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account "
                qry += "  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code "
                qry += " where  exists (select 1 from TSPL_LOCATION_MASTER where TSPL_LOCATION_MASTER .Loc_Segment_Code=substring(TSPL_BANK_MASTER .BANKACC ,len(TSPL_BANK_MASTER .BANKACC)-2,3) "


                If arrlocation IsNot Nothing AndAlso arrlocation.Count > 0 Then
                    qry += " and TSPL_LOCATION_MASTER.Loc_Segment_Code  in (" + clsCommon.GetMulcallString(arrlocation) + ")"

                End If
                qry += " )"
                qry += "  and  TSPL_PAYMENT_HEADER.Payment_Type='AV'    "
                qry += " union all"
                qry += " select distinct tspl_payment_header.Payment_No, TSPL_PAYMENT_HEADER.Debit_Account as Account_Code,(select Description from TSPL_GL_ACCOUNTS where Account_Code=TSPL_PAYMENT_HEADER.Debit_Account) as Description,tspl_payment_header.Entry_Desc,'' as Remarks,TSPL_PAYMENT_DETAIL.Comment,TSPL_PAYMENT_DETAIL.Applied_Amount as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno ,TSPL_PAYMENT_DETAIL.Vendor_Invoice_No+ case when len(RTRIM(TSPL_PAYMENT_DETAIL.Vendor_Invoice_No))>0 then ', ' else ' ' end +Vendor_Invoice_Date  as EntryDesc  from tspl_payment_header  "
                qry += " left outer join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_HEADER .Payment_No =TSPL_PAYMENT_DETAIL .Payment_No left outer join  TSPL_VENDOR_INVOICE_HEAD  on TSPL_PAYMENT_DETAIL.Document_No =TSPL_VENDOR_INVOICE_HEAD.Document_No "
                qry += "  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code "
                qry += "where  exists (select 1 from TSPL_LOCATION_MASTER where TSPL_LOCATION_MASTER .Loc_Segment_Code=substring(TSPL_BANK_MASTER .BANKACC ,len(TSPL_BANK_MASTER .BANKACC)-2,3) "
                If arrlocation IsNot Nothing AndAlso arrlocation.Count > 0 Then
                    qry += " and TSPL_LOCATION_MASTER.Loc_Segment_Code  in (" + clsCommon.GetMulcallString(arrlocation) + ")"

                End If
                qry += " ) "
                qry += " and  TSPL_PAYMENT_HEADER.Payment_Type ='PY'"

                qry += " UNION all "
                qry += " select tspl_payment_header.Payment_No, Account_Code,TSPL_PAYMENT_DETAIL.Description ,tspl_payment_header.Entry_Desc,TSPL_PAYMENT_DETAIL .Remarks ,'' as Comment,Net_Balance as DrAmt,0 as CrAmt,0 as OrderDRCR ," & _
                " (case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and ( TSPL_PAYMENT_HEADER.apply_by='sale invoice' or TSPL_PAYMENT_HEADER.apply_by='Load Out' )     then apply_to else '' end ) as saleinvoiceno," & _
                " (case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice'   then TSPL_SALE_INVOICE_HEAD .Salesman_Code  else " & _
                " case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out'   then  TSPL_SHIPMENT_MASTER.Salesman_Code else  '' end end  ) as salesmancode," & _
                "(case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice'     then a.emp_name else " & _
                " case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out' then b.emp_name  else '' end end ) as emp_name  ,(case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice' then   'Sale Invoice No' else case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out' then 'Load Out No' else ''end end )  as txtno,'' as EntryDesc " & _
                " from TSPL_PAYMENT_DETAIL " & _
                " left outer join tspl_payment_header on tspl_payment_header.Payment_No=TSPL_PAYMENT_DETAIL.Payment_No    " & _
               " left outer join TSPL_SALE_INVOICE_HEAD on TSPL_PAYMENT_HEADER.Apply_To =TSPL_SALE_INVOICE_HEAD.Sale_Invoice_No " & _
                " left outer join TSPL_SHIPMENT_MASTER on TSPL_SHIPMENT_MASTER.Shipment_No =TSPL_PAYMENT_HEADER.Apply_To      " & _
                " left outer join TSPL_EMPLOYEE_MASTER as a on TSPL_SALE_INVOICE_HEAD.Salesman_Code = a.EMP_CODE   " & _
                " left outer join TSPL_EMPLOYEE_MASTER as b   on   TSPL_SHIPMENT_MASTER.Salesman_Code =b.EMP_CODE " & _
                "  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code " & _
                " where  exists (select 1 from TSPL_LOCATION_MASTER where TSPL_LOCATION_MASTER .Loc_Segment_Code=substring(TSPL_BANK_MASTER .BANKACC ,len(TSPL_BANK_MASTER .BANKACC)-2,3) "
                If arrlocation IsNot Nothing AndAlso arrlocation.Count > 0 Then
                    qry += " and TSPL_LOCATION_MASTER.Loc_Segment_Code  in (" + clsCommon.GetMulcallString(arrlocation) + ")"
                End If
                qry += " )"
                qry += " and Account_Code not in ('NULL')"
                qry += " ) Final left outer join tspl_payment_header on TSPL_PAYMENT_HEADER.Payment_No=final.Payment_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_USER_MASTER on TSPL_USER_MASTER.User_Code=tspl_payment_header.Created_By  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code  where 2=2 "
                If arrDocument IsNot Nothing AndAlso arrDocument.Count > 0 Then
                    qry += "and  tspl_payment_header.Payment_No ='" + recvalue + "'"
                End If

                If arrVendor IsNot Nothing AndAlso arrVendor.Count > 0 Then
                    qry += "and tspl_payment_header.Vendor_Code in (" + clsCommon.GetMulcallString(arrVendor) + ")"
                End If
                If clsCommon.myLen(fromdate) > 0 Then
                    qry += " and convert(date,tspl_payment_header.payment_date,103) >= convert(date,'" + fromdate + "',103) "
                End If
                If clsCommon.myLen(todate) > 0 Then
                    qry += " and convert(date,tspl_payment_header.payment_date,103) <= convert(date,'" + todate + "',103) "
                End If
                qry += "  and tspl_payment_header.payment_type not in ('AV')  order by final.Payment_No,OrderDRCR  "
                dt = clsDBFuncationality.GetDataTable(qry)
                For count As Integer = 0 To dt.Rows.Count - 1
                    Dim newrow As DataRow = dttemp.NewRow()
                    newrow("FromDate") = dt.Rows(count)("FromDate")
                    newrow("ToDate") = dt.Rows(count)("ToDate")
                    newrow("Location") = Convert.ToString(dt.Rows(count)("Location"))
                    newrow("Vendor") = Convert.ToString(dt.Rows(count)("Vendor"))
                    newrow("Document") = Convert.ToString(dt.Rows(count)("Document"))
                    newrow("PaymentNo") = Convert.ToString(dt.Rows(count)("PaymentNo"))
                    newrow("PaymentDate") = dt.Rows(count)("PaymentDate")
                    newrow("VendorNameRemitTO") = Convert.ToString(dt.Rows(count)("VendorNameRemitTO"))
                    newrow("Cheque_No") = dt.Rows(count)("Cheque_No")
                    newrow("VoucherType") = dt.Rows(count)("VoucherType")
                    newrow("Payment_Code") = Convert.ToString(dt.Rows(count)("Payment_Code"))
                    newrow("BANK_CODE") = Convert.ToString(dt.Rows(count)("BANK_CODE"))
                    newrow("BankName") = Convert.ToString(dt.Rows(count)("BankName"))
                    newrow("Bank_acct") = Convert.ToString(dt.Rows(count)("Bank_acct"))
                    newrow("Compaddress1") = Convert.ToString(dt.Rows(count)("Compaddress1"))
                    newrow("CreatedBy") = Convert.ToString(dt.Rows(count)("CreatedBy"))
                    newrow("AuthorisedBy") = Convert.ToString(dt.Rows(count)("AuthorisedBy"))
                    newrow("CompanyCode") = Convert.ToString(dt.Rows(count)("CompanyCode"))
                    newrow("Logo_Img") = dt.Rows(count)("Logo_Img")
                    newrow("Logo_Img2") = dt.Rows(count)("Logo_Img2")
                    newrow("Payment_Type") = Convert.ToString(dt.Rows(count)("Payment_Type"))
                    newrow("AccountCode") = Convert.ToString(dt.Rows(count)("AccountCode"))
                    newrow("AccountName") = Convert.ToString(dt.Rows(count)("AccountName"))
                    newrow("Entry_Desc") = Convert.ToString(dt.Rows(count)("Entry_Desc"))
                    newrow("Remarks") = dt.Rows(count)("Remarks")
                    newrow("Comment") = dt.Rows(count)("Comment")
                    newrow("Debit") = dt.Rows(count)("Debit")
                    newrow("Credit") = dt.Rows(count)("Credit")
                    newrow("SegName") = Convert.ToString(dt.Rows(count)("SegName"))
                    newrow("saleinvoiceno") = dt.Rows(count)("saleinvoiceno")
                    newrow("salesmancode") = dt.Rows(count)("salesmancode")
                    newrow("salesmanname") = dt.Rows(count)("salesmanname")
                    newrow("txtno") = dt.Rows(count)("txtno")
                    newrow("EntryDesc") = dt.Rows(count)("EntryDesc")
                    dttemp.Rows.Add(newrow)
                Next
            Next
            Dim frmCRV As New frmCrystalReportViewer()
            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.Purchase, dttemp, "PaymentEntryReport", "Payment Details")
            frmCRV = Nothing
            ' funReport(fromdate, todate, arrDocument, arrVendor)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(me,ex.Message,Me.text)
        End Try
    End Sub

    Public Sub funReport(ByVal fromDate As String, ByVal toDate As String, ByVal ArrDocument As ArrayList, ByVal ArrVendor As ArrayList, ByVal ArrLocation As ArrayList)
        Try
            Dim frmCRV As New frmCrystalReportViewer()
            Dim dt1 As DataTable
            ''richa KDI/11/07/18-000399, KDI/11/07/18-000398 pick vendor name from vendor master
            Dim qry As String = "select TSPL_BANK_MASTER.Bank_Type, case when  TSPL_BANK_MASTER.Bank_Type = 'C' and TSPL_PAYMENT_HEADER.Payment_Type not in ( 'AD','RC') then 'Cash Payment Voucher'  when TSPL_BANK_MASTER.Bank_Type = 'B' and TSPL_PAYMENT_HEADER.Payment_Type not in ( 'AD','RC') then 'Bank Payment Voucher'  when  TSPL_BANK_MASTER.Bank_Type = 'C' and TSPL_PAYMENT_HEADER.Payment_Type not in ( 'AD') and TSPL_PAYMENT_HEADER.Payment_Type = 'RC' then 'Cash Receipt Voucher'  when TSPL_BANK_MASTER.Bank_Type = 'B' and TSPL_PAYMENT_HEADER.Payment_Type not in ( 'AD') and TSPL_PAYMENT_HEADER.Payment_Type = 'RC' then 'Bank Receipt Voucher' else '' end as HeadingVoucher, TSPL_PAYMENT_HEADER.TapalNo,TSPL_PAYMENT_HEADER.DateAndTime,Final.Payment_No as PaymentNo,Payment_Date as PaymentDate,TSPL_PAYMENT_HEADER.Vendor_Code,case when TSPL_VENDOR_MASTER.Vendor_Name = Remit_To then TSPL_PAYMENT_HEADER.Vendor_Code+' - '+TSPL_VENDOR_MASTER.Vendor_Name else case when len(isnull(TSPL_VENDOR_MASTER.Vendor_Name,''))>0 and len(isnull(Remit_To,''))>0 then TSPL_VENDOR_MASTER.Vendor_Name+' / '+Remit_To Else case When len(isnull(TSPL_PAYMENT_HEADER.Vendor_Code,''))>0 then TSPL_PAYMENT_HEADER.Vendor_Code+' - '+TSPL_VENDOR_MASTER.Vendor_Name  else TSPL_PAYMENT_HEADER.Vendor_Code+''+TSPL_VENDOR_MASTER.Vendor_Name+''+ Remit_To   end end end as VendorNameRemitTO,Cheque_No,convert(varchar,Cheque_Date,103) as Cheque_Date ,case when  TSPL_PAYMENT_HEADER.Payment_Code ='SETTLEMENT' then 'Settlement Payment Voucher' else CASE WHEN TSPL_PAYMENT_HEADER.Payment_Code ='Cheque' then 'Bank Payment Voucher' else CASE WHEN TSPL_PAYMENT_HEADER.Payment_Code ='Cash' then 'Cash Payment Voucher' else CASE WHEN TSPL_PAYMENT_HEADER.Payment_Code='PETTYCASH' THEN 'Petty Cash Payment Voucher'  " &
        " END END end end as [VoucherType], tspl_payment_header.Payment_Code,tspl_payment_header.BANK_CODE,TSPL_BANK_MASTER.DESCRIPTION as BankName, (select BANKACC from TSPL_BANK_MASTER where Bank_Code =tspl_payment_header.BANK_CODE) as Bank_acct,(select max(ADD1 + case when len(add2)> 0 then ',' else '' end + ADD2 +case when len(add3)> 0 then ','else '' end +ADD3+case when len(add4)> 0 then ',' else '' end +ADD4+case when len(City_Code)> 0 then ',' else '' end +City_Code +case when len(STATE)> 0 then ',' else '' end  +STATE) from tspl_location_master  where 1=1 "
            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AddressOnPaymentVoucher, clsFixedParameterCode.AddressOnPaymentVoucherOnBankBasis, Nothing)) = 1 Then
                qry = qry & " and Location_Code =(substring (TSPL_BANK_MASTER .BANKACC ,LEN(TSPL_BANK_MASTER .BANKACC)-2,3))and TSPL_LOCATION_MASTER .Location_Type ='Physical')as Compaddress1"
            Else
                qry = qry & " and Location_Code =tspl_payment_header.Location_GL_Code)as Compaddress1"
            End If

            qry = qry & ", TSPL_USER_MASTER.User_Name as CreatedBy,Case when (TSPL_PAYMENT_HEADER.Posted='1' or TSPL_PAYMENT_HEADER.Posted='P') then tspl_payment_header.Modify_By when (TSPL_PAYMENT_HEADER.Posted='0' or TSPL_PAYMENT_HEADER.Posted='N') then null end as AuthorisedBy,Case when (TSPL_PAYMENT_HEADER.Posted='1' or TSPL_PAYMENT_HEADER.Posted='P') then 'Y' when (TSPL_PAYMENT_HEADER.Posted='0' or TSPL_PAYMENT_HEADER.Posted='N') then 'N' end as Posted,TSPL_COMPANY_MASTER.Comp_Name as CompanyCode,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2, TSPL_PAYMENT_HEADER.Payment_Type , final.ACCode as AccountCode,final.Description as AccountName,Final .Entry_Desc,final.Remarks,final.Comment ,case when TSPL_PAYMENT_HEADER.Payment_Type='MI' and TSPL_PAYMENT_HEADER.isReceipt = 1 then  final.CrAmt else Final.DrAmt end  as Debit,case when TSPL_PAYMENT_HEADER.Payment_Type='MI' and TSPL_PAYMENT_HEADER.isReceipt = 1 then  Final.DrAmt else final.CrAmt end as Credit ,(select Top 1 Description from TSPL_GL_SEGMENT_CODE where  Seg_No='7' and Segment_code=Right(TSPL_BANK_MASTER .BANKACC,3)) as SegName,Final .saleinvoiceno ,Final .salesmancode ,Final .salesmanname,final.txtno,final.EntryDesc,final.Hirerachy_Level_Code,final.Cost_Center_Fin_Code,  Case When case when TSPL_PAYMENT_HEADER.Payment_Type='MI' and TSPL_PAYMENT_HEADER.isReceipt = 1 then  final.CrAmt else Final.DrAmt end>0 then 0 Else 1 END as OrderDRCR  from ( " ' Case When DrAmt>0 then 0 Else 1 END as OrderDRCR / Final.DrAmt as Debit,final.CrAmt as Credit
            ''richa agarwal 30/06/2015
            ' qry += " select tspl_payment_header.Payment_No, TSPL_BANK_MASTER.BANKACC as ACCode,TSPL_GL_ACCOUNTS.Description,tspl_payment_header.Entry_Desc,'' as Remarks,'' as Comment,0 as DrAmt, Payment_Amount+Bank_Charges as CrAmt,1 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc  from tspl_payment_header  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BANK_MASTER.BANKACC   Where Payment_Type not in ('RC','SR')  "
            qry += " select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No, TSPL_BANK_MASTER.BANKACC as ACCode,TSPL_GL_ACCOUNTS.Description,tspl_payment_header.Entry_Desc,'' as Remarks,'' as Comment,0 as DrAmt, case when isnull(tspl_payment_header.BASE_CURRENCY_CODE,'')=isnull(tspl_payment_header.CURRENCY_CODE,'') then Payment_Amount+Bank_Charges else (Payment_Amount* tspl_payment_header.ConvRate )+Bank_Charges end  as CrAmt,1 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code  from tspl_payment_header  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BANK_MASTER.BANKACC   Where Payment_Type not in ('RC','SR')  "
            ''------------------------
            qry += "union all"
            ''richa agarwal 30/06/2015
            'qry += "  select TSPL_PAYMENT_HEADER .Payment_No ,Debit_Account as ACCode, Description,TSPL_PAYMENT_HEADER.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment,TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account where TSPL_PAYMENT_HEADER.Payment_Type='OA'"
            qry += "  select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, TSPL_PAYMENT_HEADER .Payment_No ,Debit_Account as ACCode, Description,TSPL_PAYMENT_HEADER.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment, case when isnull(tspl_payment_header.BASE_CURRENCY_CODE,'')=isnull(tspl_payment_header.CURRENCY_CODE,'') then TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) else (TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0))* tspl_payment_header.ConvRate end  as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc ,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account where TSPL_PAYMENT_HEADER.Payment_Type='OA'"
            ''----------------------------
            qry += " union all"
            ''richa agarwal 01/07/2015
            'qry += " select tspl_payment_header.Payment_No,Debit_Account as ACCode, Description,tspl_payment_header.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment,TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account where TSPL_PAYMENT_HEADER.Payment_Type='AV'    "
            qry += " select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No,Debit_Account as ACCode, Description,tspl_payment_header.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment,case when isnull(tspl_payment_header.BASE_CURRENCY_CODE,'')=isnull(tspl_payment_header.CURRENCY_CODE,'') then TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) else (TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0))* tspl_payment_header.ConvRate end  as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc ,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account where TSPL_PAYMENT_HEADER.Payment_Type='AV'    "
            ''----------------------------------
            '-----Added By--Pankaj kUmar on---23/11/2012-----------While Adding new PaymentType[Receipt]---
            qry += " Union All"
            ''richa agarwal 01/07/2015
            'qry += " select tspl_payment_header.Payment_No,Credit_Account  as ACCode, Description,tspl_payment_header.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment,0 as DrAmt,TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) as CrAmt,1 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Credit_Account  where TSPL_PAYMENT_HEADER.Payment_Type in ('RC','SR')"
            qry += " select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No,Credit_Account  as ACCode, Description,tspl_payment_header.Entry_Desc as Entry_Desc,'' as Remarks,'' as Comment,0 as DrAmt,case when isnull(tspl_payment_header.BASE_CURRENCY_CODE,'')=isnull(tspl_payment_header.CURRENCY_CODE,'') then TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0) else (TSPL_PAYMENT_HEADER.Payment_Amount+ISNULL(TSPL_PAYMENT_HEADER.TDS_Amount,0))* tspl_payment_header.ConvRate end  as CrAmt,1 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc ,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Credit_Account  where TSPL_PAYMENT_HEADER.Payment_Type in ('RC','SR')"
            ''------------------
            qry += " Union All"
            ''richa agarwal 01/07/2015
            'qry += " select tspl_payment_header.Payment_No, TSPL_BANK_MASTER.BANKACC as ACCode,TSPL_GL_ACCOUNTS.Description,tspl_payment_header.Entry_Desc,'' as Remarks,'' as Comment,Payment_Amount as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc  from tspl_payment_header  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BANK_MASTER.BANKACC Where Payment_Type in ('RC','SR')"
            qry += " select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No, TSPL_BANK_MASTER.BANKACC as ACCode,TSPL_GL_ACCOUNTS.Description,tspl_payment_header.Entry_Desc,'' as Remarks,'' as Comment,case when isnull(tspl_payment_header.BASE_CURRENCY_CODE,'')=isnull(tspl_payment_header.CURRENCY_CODE,'') then Payment_Amount else Payment_Amount * tspl_payment_header.ConvRate end  as DrAmt,0 as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc ,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code  from tspl_payment_header  left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BANK_MASTER.BANKACC Where Payment_Type in ('RC','SR')"
            ''-------------------------
            '----------------------------------------------------------------------------------------------
            '' ------------------- Security Refund Union Starts------------------------------------------------------
            '' ------------------- Security Refund Ends--------------------------------------------------------------
            qry += " union all"
            '' credit amount not shown in case when receipt type of payment is used as a invoice into grid level
            'qry += " Select Payment_No, MAX(Account_Code) as Account_Code, MAX(Description) as Description, MAX(Entry_Desc) as Entry_Desc, '' as Remarks, MAX(Comment) as Comment, SUM(DrAmt)-SUM(CrAmt) as DrAmt, 0 as CrAmt, 0 as OrderdrCr, '' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno, MAX(EntryDesc1) as Entry_Desc,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from  (select  tspl_payment_header.Payment_No, TSPL_PAYMENT_HEADER.Debit_Account as Account_Code,(select Description from TSPL_GL_ACCOUNTS where Account_Code=TSPL_PAYMENT_HEADER.Debit_Account) as Description,tspl_payment_header.Entry_Desc,'' as Remarks,TSPL_PAYMENT_DETAIL.Comment, Case When TSPL_VENDOR_INVOICE_HEAD .Document_Type<>'D' Then TSPL_PAYMENT_DETAIL.Applied_Amount * tspl_payment_header.ConvRate Else 0 END as DrAmt, case WHEN TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then TSPL_PAYMENT_DETAIL.Applied_Amount * tspl_payment_header.ConvRate Else 0 END as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno ,TSPL_PAYMENT_DETAIL.Vendor_Invoice_No+ case when len(RTRIM(TSPL_PAYMENT_DETAIL.Vendor_Invoice_No))>0 then ', ' else ' ' end +Vendor_Invoice_Date  as EntryDesc1, Bank_Charges from tspl_payment_header "
            qry += " Select  max(Payment_Type) as Payment_Type,max(isReceipt) as isReceipt, Payment_No, MAX(Account_Code) as Account_Code, MAX(Description) as Description, MAX(Entry_Desc) as Entry_Desc, '' as Remarks, MAX(Comment) as Comment, SUM(DrAmt)-SUM(CrAmt) as DrAmt, 0 as CrAmt, 0 as OrderdrCr, '' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno, MAX(EntryDesc1) as Entry_Desc,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from  (select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No, TSPL_PAYMENT_HEADER.Debit_Account as Account_Code,(select Description from TSPL_GL_ACCOUNTS where Account_Code=TSPL_PAYMENT_HEADER.Debit_Account) as Description,tspl_payment_header.Entry_Desc,'' as Remarks,TSPL_PAYMENT_DETAIL.Comment, Case When (ISNULL(TSPL_VENDOR_INVOICE_HEAD .Document_Type,'')<>'D' or (select TPH.Payment_Type  from TSPL_PAYMENT_HEADER TPH where TPH.payment_no=TSPL_PAYMENT_DETAIL.Document_No )='RC') Then TSPL_PAYMENT_DETAIL.Applied_Amount * tspl_payment_header.ConvRate Else 0 END as DrAmt, case WHEN TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then TSPL_PAYMENT_DETAIL.Applied_Amount * tspl_payment_header.ConvRate Else 0 END as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno ,TSPL_PAYMENT_DETAIL.Vendor_Invoice_No+ case when len(RTRIM(TSPL_PAYMENT_DETAIL.Vendor_Invoice_No))>0 then ', ' else ' ' end +Vendor_Invoice_Date  as EntryDesc1, Bank_Charges from tspl_payment_header "
            qry += " left outer join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_HEADER .Payment_No =TSPL_PAYMENT_DETAIL .Payment_No left outer join  TSPL_VENDOR_INVOICE_HEAD  on TSPL_PAYMENT_DETAIL.Document_No =TSPL_VENDOR_INVOICE_HEAD.Document_No where TSPL_PAYMENT_HEADER.Payment_Type ='PY' ) aaa Group  by Payment_No "
            qry += "  union all select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, tspl_payment_header.Payment_No,(select Branch_GL_AC  from TSPL_REMITTANCE Where TSPL_REMITTANCE.Document_No = TSPL_PAYMENT_HEADER.Payment_No ) as ACCode,( select Description  from TSPL_GL_ACCOUNTS where Account_Code =(select Branch_GL_AC  from TSPL_REMITTANCE Where TSPL_REMITTANCE.Document_No = TSPL_PAYMENT_HEADER.Payment_No )) Description,'' as Entry_Desc,'' as Remarks,'' as Comment,0  as DrAmt,TSPL_PAYMENT_HEADER.TDS_Amount as CrAmt,0 as OrderDRCR,'' as saleinvoiceno,'' as salesmancode,'' as salesmanname,'' as txtno,'' as EntryDesc,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from TSPL_PAYMENT_HEADER  left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=TSPL_PAYMENT_HEADER.Debit_Account  where TSPL_PAYMENT_HEADER.Payment_Type='OA'  or TSPL_PAYMENT_HEADER.Payment_Type='AV'  "
            qry += " UNION all "
            qry += " select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt,  tspl_payment_header.Payment_No, Account_Code,TSPL_PAYMENT_DETAIL.Description ,tspl_payment_header.Entry_Desc,TSPL_PAYMENT_DETAIL .Remarks ,'' as Comment,(Case when Net_Balance >0 then Net_Balance else 0.0 end) as DrAmt,(Case when Net_Balance < 0 then Net_Balance *-1  else 0 end) As CrAmt,0 as OrderDRCR ," &
            " (case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and ( TSPL_PAYMENT_HEADER.apply_by='sale invoice' or TSPL_PAYMENT_HEADER.apply_by='Load Out' )     then apply_to else '' end ) as saleinvoiceno," &
            " (case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice'   then TSPL_SALE_INVOICE_HEAD .Salesman_Code  else " &
            " case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out'   then  TSPL_SHIPMENT_MASTER.Salesman_Code else  '' end end  ) as salesmancode," &
            "(case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice'     then a.emp_name else " &
            " case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out' then b.emp_name  else '' end end ) as emp_name  ,(case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='sale invoice' then   'Sale Invoice No' else case when TSPL_PAYMENT_HEADER.Payment_Type='MI'  and TSPL_PAYMENT_HEADER.apply_by='Load Out' then 'Load Out No' else ''end end )  as txtno,'' as EntryDesc,TSPL_HIRERACHY_LEVEL_MASTER.Description as  Hirerachy_Level_Code,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name  as Cost_Center_Fin_Code  " &
            " from TSPL_PAYMENT_DETAIL " &
            " left outer join tspl_payment_header on tspl_payment_header.Payment_No=TSPL_PAYMENT_DETAIL.Payment_No " &
            " left outer join TSPL_SALE_INVOICE_HEAD on TSPL_PAYMENT_HEADER.Apply_To =TSPL_SALE_INVOICE_HEAD.Sale_Invoice_No "
            qry += " left outer join TSPL_SHIPMENT_MASTER on TSPL_SHIPMENT_MASTER.Shipment_No =TSPL_PAYMENT_HEADER.Apply_To      " &
            " left outer join TSPL_EMPLOYEE_MASTER as a on TSPL_SALE_INVOICE_HEAD.Salesman_Code = a.EMP_CODE   " &
            " left outer join TSPL_EMPLOYEE_MASTER as b   on   TSPL_SHIPMENT_MASTER.Salesman_Code =b.EMP_CODE " &
            " left outer join TSPL_COST_CENTRE_FINANCIAL on TSPL_PAYMENT_DETAIL.Cost_Center_Fin_Code=TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code " &
            "left outer join TSPL_HIRERACHY_LEVEL_MASTER on TSPL_PAYMENT_DETAIL.Hirerachy_Level_Code=TSPL_HIRERACHY_LEVEL_MASTER.Hirerachy_Code " &
                  "where ISNULL(Account_Code,'') <>'' "
            qry += " UNION ALL"
            qry += " Select tspl_payment_header.Payment_Type,tspl_payment_header.isReceipt, TSPL_PAYMENT_HEADER.Payment_No, TSPL_PAYMENT_HEADER.Bank_Charges_Ac as Account_Code, TSPL_GL_ACCOUNTS.Description, '' as Entry_Desc, '' as Remarks, '' as Comments, case when ISNULL(Bank_Charges,0)>0 then Bank_Charges end as DrAmt,case when ISNULL(Bank_Charges,0)<0 then Bank_Charges*-1 end as CrAmt, 0 as OrderDrCr, '' as saleInvoiceNo, '' as SalesmanCode, '' as SalesmanName, '' as txtNo, '' as EntryDesc ,'' AS Hirerachy_Level_Code,'' AS Cost_Center_Fin_Code from TSPL_PAYMENT_HEADER "
            qry += " Left Outer Join TSPL_GL_ACCOUNTS ON TSPL_PAYMENT_HEADER.Bank_Charges_Ac=TSPL_GL_ACCOUNTS.Account_Code Where (ISNULL(Bank_Charges_Ac,'')<>'' And ISNULL(Bank_Charges,0)<>0) And Payment_Type not in ('RC','SR')"
            qry += " ) Final left outer join tspl_payment_header on TSPL_PAYMENT_HEADER.Payment_No=final.Payment_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=tspl_payment_header.Bank_Code  left outer join TSPL_USER_MASTER on TSPL_USER_MASTER.User_Code=tspl_payment_header.Created_By  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = tspl_payment_header.Vendor_Code  where 2=2 "

            If ArrDocument IsNot Nothing AndAlso ArrDocument.Count > 0 Then
                qry += "and  tspl_payment_header.Payment_No in (" + clsCommon.GetMulcallString(ArrDocument) + ")"
            End If

            If ArrVendor IsNot Nothing AndAlso ArrVendor.Count > 0 Then
                qry += "and tspl_payment_header.Vendor_Code in (" + clsCommon.GetMulcallString(ArrVendor) + ")"
            End If

            If clsCommon.myLen(fromDate) > 0 Then
                qry += " and convert(date,tspl_payment_header.payment_date,103) >= convert(date,'" + fromDate + "',103) "
            End If
            If clsCommon.myLen(toDate) > 0 Then
                qry += " and convert(date,tspl_payment_header.payment_date,103) <= convert(date,'" + toDate + "',103) "
            End If
            qry += "  and tspl_payment_header.payment_type not in ('AD')   "  'order by OrderDRCR, final.Payment_No 
            qry = " select TSPL_VENDOR_MASTER.GSTFinalNo,TSPL_STATE_MASTER_vendor.GST_STATE_Code,xxxx.*,TSPL_VENDOR_MASTER.*,TSPL_PAYMENT_HEADER.Vendor_Bank_Code ,TSPL_PAYMENT_HEADER.Vendor_Bank_Name ,TSPL_PAYMENT_HEADER.Vendor_IFSC_Code ,TSPL_PAYMENT_HEADER.Vendor_Branch_Name ,TSPL_PAYMENT_HEADER.Vendor_Bank_ACNo from (  " + qry + " ) xxxx left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = xxxx.Vendor_Code left outer join TSPL_STATE_MASTER as  TSPL_STATE_MASTER_vendor on TSPL_STATE_MASTER_vendor.State_Code=TSPL_VENDOR_MASTER.State_Code left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=xxxx.PaymentNo order by xxxx.OrderDRCR,xxxx.PaymentNo "

            dt1 = clsDBFuncationality.GetDataTable(qry)


            '' added By Richa Agarwal 29 aug,2018 BHA/28/08/18-000487   For SubReport
            Dim InvoiceSubReport As String = "Select TSPL_PAYMENT_DETAIL.Payment_No  ,TSPL_PAYMENT_DETAIL.Document_No ,CONVERT(VARCHAR,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) AS Document_Date ,TSPL_PAYMENT_DETAIL.Original_Invoice_Amt AS Original_Amt,TSPL_PAYMENT_DETAIL.Applied_Amount From TSPL_PAYMENT_DETAIL LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No =TSPL_PAYMENT_DETAIL.Payment_No LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_PAYMENT_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No where 1=1 AND TSPL_PAYMENT_HEADER.Payment_Type <>'MI' AND TSPL_PAYMENT_DETAIL.Payment_No in (" & clsCommon.GetMulcallString(ArrDocument) & ") "
            Dim frmcrystal As New frmCrystalReportViewer()
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "VIZAG") = CompairStringResult.Equal Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.Purchase, dt1, EnumTecxpertPaperSize.HalfLegal85x7, "PaymentEntryReport4VIZAG", "Payment Details")
                Else
                    frmcrystal.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.Purchase, dt1, clsDBFuncationality.GetDataTable(InvoiceSubReport), "PaymentEntryReport4GUNTUR", "Payment Details", clsCommon.myCDate(dt1.Rows(0)("PaymentDate")), "rptPaymentDetailWithInvoice.rpt")
                    'frmCRV.funreport(CrystalReportFolder.Purchase, dt1, EnumTecxpertPaperSize.NA, "PaymentEntryReport4GUNTUR", "Payment Details", clsCommon.myCDate(dt1.Rows(0)("PaymentDate")),)
                End If
            Else
                common.clsCommon.MyMessageBoxShow("Data not found.")

            End If
            frmCRV = Nothing


        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        funreset()
    End Sub
    Sub funreset()
        dtpFromDate.Value = Date.Today
        dtpToDate.Value = Date.Today
        chkDocumentAll.IsChecked = True
        chkVendorAll.IsChecked = True
        cbgLocation.DataSource = clsLocation.GetLocationSegments()

        chkDocumentAll.IsChecked = True
        chkVendorAll.IsChecked = True
        chkLocAll.IsChecked = True
        LoadVendor()
        LoadDocument()
    End Sub
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.frmPaymentEntry)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")

        End If
        'btnSave.Visible = MyBase.isModifyFlag
        'btnPost.Visible = MyBase.isPostFlag
        'btnDelete.Visible = MyBase.isDeleteFlag

    End Sub
    Sub loadlocation()
        'Dim qry As String = " select Segment_code as Code, Description  from TSPL_GL_SEGMENT_CODE where Seg_No='7' "
        cbgLocation.DataSource = clsLocation.GetLocationSegments()
        cbgLocation.ValueMember = "Code"
        cbgLocation.DisplayMember = "Name"
    End Sub
    Private Sub FrmPaymentEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetUserMgmtNew()
        loadlocation()
        'cbgLocation.DataSource = clsLocation.GetLocationSegments()
        dtpFromDate.Value = Date.Today
        dtpToDate.Value = Date.Today
        'If objCommonVar.CurrentUserCode <> "ADMIN" Then
        '    If funSetUserAccess() = False Then Exit Sub
        'End If
        chkDocumentAll.IsChecked = True
        chkVendorAll.IsChecked = True

        LoadVendor()
        LoadDocument()
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnReset, "Press Alt+N Adding New Trasnaction")
        ButtonToolTip.SetToolTip(btnPrint, "Press Ctrl+P for Print ")

    End Sub

    Sub LoadVendor()

        Dim strqry As String = "select Vendor_Code as [Vendor Number],Vendor_Name as [Vendor Name] from TSPL_VENDOR_MASTER  WHERE Status='N' "
        cbgVendor.DataSource = clsDBFuncationality.GetDataTable(strqry)
        cbgVendor.ValueMember = "Vendor Number"
        cbgVendor.DisplayMember = "Vendor Name"
    End Sub

    Sub LoadDocument()
        Dim strqry As String = "select  Payment_No as [Document Number]," & _
"Payment_Date ,Bank_Code ,Vendor_Code ,Vendor_Name  from tspl_payment_header where CONVERT(Date,Payment_Date ,103)>=Convert(Date,'" + clsCommon.GetPrintDate(dtpFromDate.Value, "dd/MM/yyyy") + "',103) and CONVERT(Date,payment_date,103)<=Convert(Date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") + "',103) "
        cbgDocument.DataSource = clsDBFuncationality.GetDataTable(strqry)
        cbgDocument.ValueMember = "Document Number"
        cbgDocument.DisplayMember = "Payment_Date"
    End Sub




    'This will check the authorization of user to access the screen.If authorize then it will allow user to access the screen.
    'Private Function funSetUserAccess() As Boolean
    '    Try

    '        Dim strRights As String
    '        Dim strTemp() As String
    '        Dim strProgCode = "PAY-ENT-RPT"
    '        strRights = enuUserRights.enuRead & "," & enuUserRights.enuModify & "," & enuUserRights.enuDelete
    '        strRights = modUserMgt.funGetPermissions(strRights, strProgCode)
    '        strTemp = Split(strRights, ",")
    '        If strTemp(0) = "0" Then
    '            MsgBox("Permission Denied", MsgBoxStyle.Critical, Me.Text)
    '            funSetUserAccess = False
    '            blnRead = False
    '            Me.Close()
    '            Exit Function
    '        Else
    '            blnRead = True
    '        End If
    '        If strTemp(1) = "0" Then 'Grant modify access

    '        End If
    '        If strTemp(2) = "0" Then 'Grant modify access

    '        End If

    '        funSetUserAccess = True
    '    Catch er As Exception
    '        myMessages.myExceptions(er)
    '    End Try
    'End Function

    Private Sub chkVendorAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkVendorAll.ToggleStateChanged, chkVendorSelect.ToggleStateChanged
        cbgVendor.Enabled = chkVendorSelect.IsChecked
    End Sub


    Private Sub chkDocumentAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkDocumentAll.ToggleStateChanged, ChkDocumentSelect.ToggleStateChanged
        cbgDocument.Enabled = ChkDocumentSelect.IsChecked
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dtpFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged
        Dim strqry As String = "select  Payment_No as [Document Number]," & _
            "Payment_Date ,Bank_Code ,Vendor_Code ,Vendor_Name  from tspl_payment_header where CONVERT(Date,Payment_Date ,103)>=Convert(Date,'" + clsCommon.GetPrintDate(dtpFromDate.Value, "dd/MM/yyyy") + "',103) and CONVERT(Date,payment_date,103)<=Convert(Date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") + "',103) "
        cbgDocument.DataSource = clsDBFuncationality.GetDataTable(strqry)
        cbgDocument.ValueMember = "Document Number"
        cbgDocument.DisplayMember = "Payment_Date"
    End Sub

    Private Sub dtpToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged
        Dim strqry As String = "select  Payment_No as [Document Number]," & _
                    "Payment_Date ,Bank_Code ,Vendor_Code ,Vendor_Name  from tspl_payment_header where CONVERT(Date,Payment_Date ,103)>=Convert(Date,'" + clsCommon.GetPrintDate(dtpFromDate.Value, "dd/MM/yyyy") + "',103) and CONVERT(Date,payment_date,103)<=Convert(Date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") + "',103) "
        cbgDocument.DataSource = clsDBFuncationality.GetDataTable(strqry)
        cbgDocument.ValueMember = "Document Number"
        cbgDocument.DisplayMember = "Payment_Date"
    End Sub

    Private Sub FrmPaymentEntry_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag Then
            'SaveData()
        ElseIf e.Control AndAlso e.KeyCode = Keys.P Then
            funPrint()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag Then
            'DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C Then
            Close()
        ElseIf e.Alt And e.KeyCode = Keys.N Then
            funreset()
        End If
    End Sub

    Private Sub chkLocAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkLocAll.ToggleStateChanged, chkLocSelect.ToggleStateChanged
        cbgLocation.Enabled = Not chkLocAll.IsChecked
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class
