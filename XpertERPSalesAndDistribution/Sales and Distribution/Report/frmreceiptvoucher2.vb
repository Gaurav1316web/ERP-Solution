'' richa agarwal against ticket no. BM00000006889
Imports System.Data
Imports XpertERPEngine
Imports System.Data.SqlClient
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports common
Imports System.IO
Public Class Frmreceiptvoucher2
    Inherits FrmMainTranScreen
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim userCode, companyCode As String
    Public Sub New(ByVal user As String, ByVal company As String)
        InitializeComponent()
        userCode = user
        companyCode = company
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub fndreceiptfrom_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'fndreceiptfrom.ConnectionString = connectSql.SqlCon()
        'fndreceiptfrom.Query = "select Receipt_No as [Receipt No], Entry_Desc as [Description] from TSPL_RECEIPT_HEADER  "
        'fndreceiptfrom.ValueToSelect = "Receipt No"
        'fndreceiptfrom.Caption = "Receipt Detail"
        'fndreceiptfrom.ValueToSelect1 = "Description"
    End Sub
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.receiptreport)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
            
        End If
        'btnSave.Visible = MyBase.isModifyFlag
        'btnPost.Visible = MyBase.isPostFlag
        'btnDelete.Visible = MyBase.isDeleteFlag
    End Sub

    Private Sub fndreceiptto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'fndreceiptto.ConnectionString = connectSql.SqlCon()
        'fndreceiptto.Query = "select Receipt_No as [Receipt No], Entry_Desc as [Description] from TSPL_RECEIPT_HEADER  "
        'fndreceiptto.ValueToSelect = "Receipt No"
        'fndreceiptto.Caption = "Receipt Detail"
        'fndreceiptto.ValueToSelect1 = "Description"
    End Sub

    Private Sub fndcustomerfrom_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'fndcustomerfrom.ConnectionString = connectSql.SqlCon()
        'fndcustomerfrom.Query = "select cust_code as [Customer Code], Customer_Name as [Customer Name] from tspl_customer_master"
        'fndcustomerfrom.ValueToSelect = "Customer Code"
        'fndcustomerfrom.Caption = "Customer Details"
        'fndcustomerfrom.ValueToSelect1 = "Customer Name"
    End Sub

    Private Sub fndcustomerto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'fndcustomerto.ConnectionString = connectSql.SqlCon()
        'fndcustomerto.Query = "select cust_code as [Customer Code], Customer_Name as [Customer Name] from tspl_customer_master"
        'fndcustomerto.ValueToSelect = "Customer Code"
        'fndcustomerto.Caption = "Customer Details"
        'fndcustomerto.ValueToSelect1 = "Customer Name"

    End Sub

    Private Sub btnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprint.Click
        Print()
    End Sub

    Sub Print()
        If chkLocSelect.IsChecked = True AndAlso cbgLocation.CheckedValue.Count <= 0 Then
            common.clsCommon.MyMessageBoxShow("Please select atleast one Location")
            Return
        End If
        Dim fromdate As String = clsCommon.GetPrintDate(dtpfrom.Value, "dd/MM/yyyy")
        Dim todate As String = clsCommon.GetPrintDate(dtpto.Value, "dd/MM/yyyy")
        Dim arrreceipt As ArrayList = Nothing
        If chkReceiptSelect.IsChecked = True Then
            arrreceipt = cbgReceipt.CheckedValue
        ElseIf chkReceiptSelect.IsChecked = False Then
            arrreceipt = cbgReceipt.AllValue
        End If
        Dim arrlocation As ArrayList = Nothing
        If chkLocSelect.IsChecked = True Then
            arrlocation = cbgLocation.CheckedValue
        ElseIf chkLocSelect.IsChecked = False Then
            arrlocation = cbgLocation.AllValue
        End If
        Dim arrcustomer As ArrayList = Nothing
        If chkCustomerSelect.IsChecked = True Then
            arrcustomer = cbgCustomer.CheckedValue
        End If


        PrintData(fromdate, todate, chkReceiptSelect.IsChecked, arrreceipt, chkCustomerSelect.IsChecked, arrcustomer, chkLocSelect.IsChecked, arrlocation)
    End Sub

    Sub PrintData(ByVal fromdate As String, ByVal todate As String, ByVal chkReceiptSelect As Boolean, ByVal arrreceipt As ArrayList, ByVal chkCustomerSelect As Boolean, ByVal arrcustomer As ArrayList, ByVal chklocationselect As Boolean, ByVal arrlocation As ArrayList)
        Dim dttemp As New DataTable()
        Dim strTDocNo As String = String.Empty
        Dim recvalue1 As String = String.Empty
        Dim arrDocNo1 As New List(Of String)() ''---- '' added By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
        Try
            If chkReceiptSelect AndAlso arrreceipt.Count <= 0 Then
                common.clsCommon.MyMessageBoxShow("Please Select at least one Receipt.")
                Return
            ElseIf (chkReceiptSelect) = False Then
                arrreceipt = arrreceipt
            End If
            If chklocationselect AndAlso arrlocation.Count <= 0 Then
                common.clsCommon.MyMessageBoxShow("Please Select at least one Location.")
                Return
            ElseIf (chklocationselect) = False Then
                arrlocation = arrlocation
            End If
            If chkCustomerSelect And (arrcustomer Is Nothing OrElse arrcustomer.Count <= 0) Then
                common.clsCommon.MyMessageBoxShow("Please Select at least one Customer.")
                Return
            ElseIf (chkCustomerSelect) Then
                arrcustomer = arrcustomer
            End If

            Dim qry As String = "select Receipt_No  from TSPL_RECEIPT_HEADER where 2=2"
            If arrreceipt IsNot Nothing AndAlso arrreceipt.Count > 0 Then
                qry += " and Receipt_No in (" + clsCommon.GetMulcallString(arrreceipt) + ")"
            End If
            If arrcustomer IsNot Nothing AndAlso arrcustomer.Count > 0 Then
                qry += "and Cust_Code in (" + clsCommon.GetMulcallString(arrcustomer) + ")"
            End If
            Dim dtdocument As DataTable = clsDBFuncationality.GetDataTable(qry)
            dttemp.Columns.Add("Cheque_No", GetType(System.String))
            dttemp.Columns.Add("Cheque_Date", GetType(System.String))
            dttemp.Columns.Add("Cheque_From", GetType(System.String))
            dttemp.Columns.Add("Entry_Desc", GetType(System.String))
            dttemp.Columns.Add("Receipt_No", GetType(System.String))
            dttemp.Columns.Add("Receipt_Post_Date")
            dttemp.Columns.Add("Created_By")
            dttemp.Columns.Add("Modify_By")
            dttemp.Columns.Add("Detail_Line_No")
            dttemp.Columns.Add("Account_code", GetType(System.String))
            dttemp.Columns.Add("Account_Desc", GetType(System.String))
            dttemp.Columns.Add("Document_No", GetType(System.String))
            dttemp.Columns.Add("Remarks", GetType(System.String))
            dttemp.Columns.Add("Comment", GetType(System.String))
            dttemp.Columns.Add("Amount", GetType(System.Decimal))
            dttemp.Columns.Add("Description", GetType(System.String))
            dttemp.Columns.Add("Cust_Code", GetType(System.String))
            dttemp.Columns.Add("GSTNO", GetType(System.String))
            dttemp.Columns.Add("GST_STATE_Code", GetType(System.String))
            dttemp.Columns.Add("Customer_Name", GetType(System.String))
            dttemp.Columns.Add("BankName", GetType(System.String))
            dttemp.Columns.Add("CompanyName", GetType(System.String))
            Dim logo1 As DataColumn = New DataColumn()
            logo1.DataType = System.Type.GetType("System.Byte[]")

            Dim logo2 As DataColumn = New DataColumn()
            logo2.DataType = System.Type.GetType("System.Byte[]")
            logo1.Caption = "logo1"
            logo2.Caption = "logo2"
            dttemp.Columns.Add(logo1)
            dttemp.Columns.Add(logo2)
            'dttemp.Columns.Add("logo1", GetType(Image))
            'dttemp.Columns.Add("logo2", GetType(Image))
            dttemp.Columns.Add("Posted", GetType(System.String))
            dttemp.Columns.Add("SegName", GetType(System.String))
            dttemp.Columns.Add("vouchertype", GetType(System.String))
            dttemp.Columns.Add("SaleInvoice", GetType(System.String))
            dttemp.Columns.Add("TapalNo", GetType(System.String))
            dttemp.Columns.Add("DateAndTime")
            dttemp.Columns.Add("DocNo", GetType(System.String)) '' --- '' added By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
            Dim dt As DataTable
            Dim recvalue As String
            For Each drreceipt As DataRow In dtdocument.Rows
                recvalue = drreceipt("Receipt_No").ToString()
                recvalue1 = drreceipt("Receipt_No").ToString()
                Dim strquery As String
                Dim strQuerydocno As String = String.Empty
                Dim strqueryremarks As String = String.Empty
                Dim strquerycomment As String = String.Empty
                ' Dim strdoccomment As String
                strquery = "select  TSPL_RECEIPT_DETAIL.Document_No as DocNo,TSPL_RECEIPT_DETAIL.Document_No + CASE WHEN len(TSPL_SALE_INVOICE_HEAD.Sale_Invoice_Date) < 0 THEN '   ' ELSE ',' + CONVERT(varchar(10), TSPL_SALE_INVOICE_HEAD.Sale_Invoice_Date, 103) END AS Document_No,   TSPL_RECEIPT_DETAIL.Remarks, TSPL_RECEIPT_DETAIL.Comment from TSPL_RECEIPT_DETAIL   LEFT OUTER JOIN TSPL_SALE_INVOICE_HEAD ON TSPL_SALE_INVOICE_HEAD.Sale_Invoice_No= TSPL_RECEIPT_DETAIL.Document_No  where TSPL_RECEIPT_DETAIL.Receipt_No ='" + recvalue + "' "
                dt = clsDBFuncationality.GetDataTable(strquery)
                Dim strdocno As String = String.Empty
                Dim strremark As String = String.Empty   '''''''''''''''' Added by Abhishek kumar as on 26/09/2012 due to remarks repeating
                Dim strcomment As String = ""
                Dim arrDocNo As New List(Of String)()
                Dim arrRmks As New List(Of String)()
                Dim arrComment As New List(Of String)()
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows

                        Dim strTRmks As String = String.Empty
                        Dim strTComment As String = String.Empty
                        Dim DocNo As String = clsCommon.myCstr(dr("DocNo")) '-- '' added By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
                        strTDocNo = clsCommon.myCstr(dr("Document_No"))
                        strTRmks = clsCommon.myCstr(dr("Remarks"))
                        strTComment = clsCommon.myCstr(dr("Comment"))

                        arrDocNo1.Add(DocNo) '' added By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
                        If clsCommon.myLen(strTDocNo) > 0 AndAlso Not arrDocNo.Contains(strTDocNo) Then
                            If clsCommon.myLen(strdocno) > 0 Then
                                strdocno += ", "
                            End If
                            strdocno += strTDocNo
                        Else
                            arrDocNo.Add(strTDocNo)
                        End If
                        If clsCommon.myLen(strTComment) > 0 AndAlso Not arrComment.Contains(strTComment) Then
                            If clsCommon.myLen(strcomment) > 0 Then
                                strcomment += ", "
                            End If
                            strcomment += strTComment
                        Else
                            arrComment.Add(strTComment)
                        End If
                        If clsCommon.myLen(strTRmks) > 0 AndAlso Not arrDocNo.Contains(strTRmks) Then
                            If clsCommon.myLen(strremark) > 0 Then

                                strremark += ","
                            End If
                            strremark += strTRmks
                        Else
                            arrRmks.Add(strremark)
                        End If

                    Next

                End If


                strremark = strremark.Replace("'", "''")
                '' ---- '' added One Column DocNo in this Query By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
                strquery = "SELECT  GSTNO, GST_STATE_Code,   Cheque_No,  Cheque_Date, Entry_Desc, Receipt_No,convert(varchar, Receipt_Post_Date,103)as Receipt_Post_Date, Created_By, Modify_By, Detail_Line_No, Account_code, Account_Desc, " &
                                       " '" + strdocno + "' as Document_No,(Substring('" & strdocno & "',0,charindex(',','" & strdocno & "')))as DocNo,''as Remarks, rmk as Comment, Amount, Description, Cust_Code, Customer_Name, BankName, " &
                                       " (SELECT     Comp_Name FROM TSPL_COMPANY_MASTER WHERE      (Comp_Code = '" + objCommonVar.CurrentCompanyCode + "')) AS CompanyName, " &
                                       " (SELECT     Logo_Img FROM          TSPL_COMPANY_MASTER WHERE      (Comp_Code = '" + objCommonVar.CurrentCompanyCode + "')) AS logo1, " &
                                       " (SELECT     Logo_Img2 FROM          TSPL_COMPANY_MASTER WHERE      (Comp_Code = '" + objCommonVar.CurrentCompanyCode + "')) AS logo2, " &
                                       " (SELECT     Description FROM TSPL_GL_SEGMENT_CODE WHERE      (Seg_No = '7') AND (Segment_code = SUBSTRING(xxx.Account_code, LEN(xxx.Account_code) - 2, 3))) AS SegName," &
                "case "
                '' Anubhooti 11-Feb-2015 (If Receipt Type =Refund Then VoucherType Heading should be 'Payment Voucher')
                strquery += " WHEN (Select Receipt_Type From TSPL_RECEIPT_HEADER Where Receipt_No ='" + recvalue + "') ='F' THEN 'Payment Voucher'" &
                " when  xxx.Payment_Code ='cheque' then 'Bank Receipt Voucher' when xxx.Payment_Code ='check' then 'Bank Receipt Voucher'when xxx.Payment_Code ='SETTLEMENT' then'Cash Receipt Voucher'when xxx.Payment_Code ='SETTLEB'then 'Bank Receipt Voucher' when xxx.Payment_Code ='CASH' then 'Cash Receipt Voucher'  " &
                "  when xxx.Payment_Code ='NEFT' THEN 'NEFT Receipt Voucher'when xxx.Payment_Code ='RTGS' THEN 'RTGS Receipt Voucher' when xxx.Payment_Code ='DD' THEN 'DD Receipt Voucher' when xxx.Payment_Code='OTHER' then '" + recvalue + "'+ ' Voucher'    else (select distinct TSPL_PAYMENT_CODE.Payment_Code  from TSPL_PAYMENT_CODE left outer join TSPL_RECEIPT_HEADER on TSPL_PAYMENT_CODE.Payment_Code =TSPL_RECEIPT_HEADER.Payment_Code  where Receipt_No='" + recvalue + "') +'  Voucher' " &
                " end   as vouchertype ,xxx.IsShowDocumentNo, Cheque_From,case when Posted ='N' Then 'UN-APPROVED' else 'P' end as Posted,SaleInvoice,BRNCH_AC ,GL_ACC,xxx.TapalNo,xxx.DateAndTime  FROM  (" &
                " select TSPL_RECEIPT_HEADER.Cheque_No,TSPL_RECEIPT_HEADER.Cheque_Date,TSPL_RECEIPT_HEADER.Entry_Desc, xx.Receipt_No,TSPL_RECEIPT_HEADER.Receipt_Post_Date,TSPL_RECEIPT_HEADER.Created_By,TSPL_RECEIPT_HEADER.Modify_By,0 as Detail_Line_No,xx.ACCode as Account_code,xx.Account_Desc,xx.Amount,'' as Description,TSPL_RECEIPT_HEADER.Cust_Code, TSPL_RECEIPT_HEADER.Customer_Name,   TSPL_BANK_MASTER.DESCRIPTION AS BankName ,TSPL_RECEIPT_HEADER.Payment_Code,TSPL_CUSTOMER_MASTER.GSTNO, TSPL_STATE_MASTER.GST_STATE_Code,xx.IsShowDocumentNo,tspl_gl_accounts.Account_Seg_Code7, TSPL_RECEIPT_HEADER.Cheque_From, xx.Entry_Desc as rmk,TSPL_RECEIPT_HEADER.Posted , SaleInvoice " + Environment.NewLine &
                " ,ISNULL(SUBSTRING(TSPL_BRANCH_ACCOUNT_MAPPING.BRANCH_ACCOUNT,0,LEN(TSPL_BRANCH_ACCOUNT_MAPPING.BRANCH_ACCOUNT)-3),'') AS BRNCH_AC,SUBSTRING(TSPL_GL_ACCOUNTS.Account_Code,0,LEN(TSPL_GL_ACCOUNTS.Account_Code)-3) AS GL_ACC,TSPL_RECEIPT_HEADER.TapalNo,TSPL_RECEIPT_HEADER.DateAndTime " &
                " from(" + Environment.NewLine

                strquery += " SELECT TSPL_RECEIPT_HEADER.Receipt_No,  TSPL_JOURNAL_DETAILS.Account_code  as ACCode,TSPL_JOURNAL_DETAILS.Account_Desc,TSPL_JOURNAL_DETAILS.Amount as Amount," + Environment.NewLine &
                " CASE WHEN LEN(ISNULL(TSPL_RECEIPT_HEADER.UnApplied_No,''))>0 THEN 0 ELSE 1 END as IsShowDocumentNo, '' as Entry_Desc ,'' as SaleInvoice" + Environment.NewLine &
                " FROM TSPL_RECEIPT_HEADER" + Environment.NewLine &
                " LEFT OUTER JOIN TSPL_JOURNAL_MASTER ON TSPL_JOURNAL_MASTER .Source_Doc_No =TSPL_RECEIPT_HEADER .Receipt_No " + Environment.NewLine &
                " LEFT OUTER JOIN TSPL_JOURNAL_DETAILS  ON TSPL_JOURNAL_MASTER .Voucher_No  =TSPL_JOURNAL_DETAILS .Voucher_No " + Environment.NewLine &
                " )xx " + Environment.NewLine &
                " left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No=xx.Receipt_No" + Environment.NewLine &
                " left  outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code=xx.ACCode" + Environment.NewLine &
                 " left  outer join TSPL_BRANCH_ACCOUNT_MAPPING on TSPL_GL_ACCOUNTS.Account_Code=TSPL_BRANCH_ACCOUNT_MAPPING.Branch_Account " & Environment.NewLine &
                " LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE =TSPL_RECEIPT_HEADER.Bank_Code left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_RECEIPT_HEADER.Cust_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= TSPL_STATE_MASTER.State_Code " + Environment.NewLine &
                " ) AS xxx " + Environment.NewLine &
                " WHERE   2=2 "

                Dim StrAllowBranchAcconReceiptPrint = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowBranchAcconReceiptPrint, clsFixedParameterCode.AllowBranchAcconReceiptPrint, Nothing)) = 1, True, False))

                If StrAllowBranchAcconReceiptPrint = False Then
                    strquery += " AND BRNCH_AC =''"
                End If
                If clsCommon.myLen(fromdate) > 0 Then
                    strquery += " and convert(date,Receipt_Post_Date,103) >= convert(date,'" + fromdate + "',103)"
                End If

                If clsCommon.myLen(todate) > 0 Then
                    strquery += " and convert(date,Receipt_Post_Date,103) <= convert(date,'" + todate + "',103)"
                End If

                If arrreceipt IsNot Nothing AndAlso arrreceipt.Count > 0 Then
                    strquery += "and Receipt_No='" + recvalue + "'"
                End If
                If arrcustomer IsNot Nothing AndAlso arrcustomer.Count > 0 Then
                    strquery += "and Cust_Code in (" + clsCommon.GetMulcallString(arrcustomer) + ")"
                End If
                If arrlocation IsNot Nothing AndAlso arrlocation.Count > 0 Then
                    strquery += "and xxx. Account_Seg_Code7  in (" + clsCommon.GetMulcallString(arrlocation) + ")"
                End If
                dt = clsDBFuncationality.GetDataTable(strquery)
                For count As Integer = 0 To dt.Rows.Count - 1
                    Dim newrow As DataRow = dttemp.NewRow()
                    newrow("Cheque_No") = Convert.ToString(dt.Rows(count)("Cheque_No"))
                    If clsCommon.myLen(dt.Rows(count)("Cheque_No")) <= 0 Then
                        newrow("Cheque_Date") = Nothing
                    Else
                        newrow("Cheque_Date") = dt.Rows(count)("Cheque_Date")
                    End If
                    newrow("Entry_Desc") = Convert.ToString(dt.Rows(count)("Entry_Desc"))
                    newrow("Receipt_No") = Convert.ToString(dt.Rows(count)("Receipt_No"))
                    newrow("Receipt_Post_Date") = dt.Rows(count)("Receipt_Post_Date")
                    newrow("Created_By") = Convert.ToString(dt.Rows(count)("Created_By"))
                    newrow("Modify_By") = Convert.ToString(dt.Rows(count)("Modify_By"))
                    newrow("Detail_Line_No") = Convert.ToString(dt.Rows(count)("Detail_Line_No"))

                    newrow("Account_code") = Convert.ToString(dt.Rows(count)("Account_code"))
                    newrow("Account_Desc") = Convert.ToString(dt.Rows(count)("Account_Desc"))

                    If clsCommon.myCdbl(dt.Rows(count)("IsShowDocumentNo")) = 1 Then
                        newrow("Document_No") = Convert.ToString(dt.Rows(count)("Document_No"))
                    End If
                    newrow("DocNo") = Convert.ToString(dt.Rows(count)("DocNo")) '' added By Abhishek kumar as on 11 july 2012  get DocNo For SubReport
                    newrow("Remarks") = Convert.ToString(dt.Rows(count)("Remarks"))
                    newrow("Comment") = Convert.ToString(dt.Rows(count)("Comment"))
                    'newrow("Amount") = Convert.ToString(dt.Rows(count)("Amount"))
                    newrow("Amount") = clsCommon.myCdbl(dt.Rows(count)("Amount"))
                    newrow("Description") = Convert.ToString(dt.Rows(count)("Description"))
                    newrow("Cust_Code") = Convert.ToString(dt.Rows(count)("Cust_Code"))
                    newrow("GSTNO") = Convert.ToString(dt.Rows(count)("GSTNO"))
                    newrow("GST_STATE_Code") = Convert.ToString(dt.Rows(count)("GST_STATE_Code"))
                    newrow("Customer_Name") = Convert.ToString(dt.Rows(count)("Customer_Name"))
                    newrow("BankName") = Convert.ToString(dt.Rows(count)("BankName"))
                    newrow("CompanyName") = Convert.ToString(dt.Rows(count)("CompanyName"))
                    If clsCommon.myLen(dt.Rows(0)("logo1")) > 0 Then
                        newrow("Column1") = DirectCast(dt.Rows(0)("logo1"), Byte())
                    End If
                    If clsCommon.myLen(dt.Rows(0)("logo2")) > 0 Then
                        newrow("Column2") = DirectCast(dt.Rows(0)("logo2"), Byte())
                    End If

                    newrow("Posted") = Convert.ToString(dt.Rows(count)("Posted"))
                    newrow("SegName") = Convert.ToString(dt.Rows(count)("SegName"))
                    newrow("vouchertype") = Convert.ToString(dt.Rows(count)("vouchertype"))
                    newrow("Cheque_From") = Convert.ToString(dt.Rows(count)("Cheque_From"))
                    newrow("SaleInvoice") = Convert.ToString(dt.Rows(count)("SaleInvoice"))
                    newrow("TapalNo") = Convert.ToString(dt.Rows(count)("TapalNo"))
                    newrow("DateAndTime") = dt.Rows(count)("DateAndTime")
                    dttemp.Rows.Add(newrow)
                Next
            Next
            '' added By Abhishek kumar as on 11 july 2012   For SubReport
            Dim adjustSubReport As String = "select finalQry .*,TSPL_COMPANY_MASTER.Comp_Name as compname, TSPL_COMPANY_MASTER.Logo_Img as Image1, TSPL_COMPANY_MASTER.Logo_Img2 as Image2,(select max(ADD1 + case when len(add2)> 0 then ',' else '' end + ADD2 +case when len(add3)> 0 then ','else '' end +ADD3+case when len(add4)> 0 then ',' else '' end +ADD4+case when len(City_Code)> 0 then ',' else '' end +City_Code +case when len(STATE)> 0 then ',' else '' end  +STATE) from tspl_location_master where Location_Code in(select Location_Code from TSPL_LOCATION_MASTER where Loc_Segment_Code =(substring (finalqry.AcctNo ,LEN(finalqry.AcctNo)-2,5)))   )as address from (select xx.Adjustment_No,Convert(varchar,xx.Adjustment_Date,103) as Adjustment_Date ,xx.Customer_No ,xx.Customer_Name ,xx.AcctNo ,xx.AcctDesc ,xx.DbtAmt ,xx.CrAmt ,xx.Comp_Code,xx.Doc_No  from " &
                  " (SELECT  TSPL_Receipt_Adjustment_Detail.Adjustment_No  ,Adjustment_Date,TSPL_Receipt_Adjustment_Header.Customer_No ,  (select Customer_Name from TSPL_CUSTOMER_MASTER where cust_Code =TSPL_Receipt_Adjustment_Header .Customer_No )as Customer_Name,TSPL_Receipt_Adjustment_Detail.Account_No as AcctNo,Account_Description as AcctDesc,Amount as DbtAmt,0 as CrAmt,TSPL_Receipt_Adjustment_Header .Comp_Code,TSPL_Receipt_Adjustment_Header .Doc_No  FROM TSPL_Receipt_Adjustment_Detail left outer join TSPL_Receipt_Adjustment_Header on  TSPL_Receipt_Adjustment_Detail.Adjustment_No = TSPL_Receipt_Adjustment_Header .Adjustment_No   where TSPL_Receipt_Adjustment_Header .Doc_No  in (" & clsCommon.GetMulcallString(arrDocNo1) & ")" &
                   " union all " &
                   " select TSPL_Receipt_Adjustment_Header.Adjustment_No ,TSPL_Receipt_Adjustment_Header .Adjustment_Date  ,TSPL_Receipt_Adjustment_Header .Customer_No,(select Customer_Name from TSPL_CUSTOMER_MASTER where cust_Code =TSPL_Receipt_Adjustment_Header .Customer_No )as Customer_Name,(select TSPL_CUSTOMER_ACCOUNT_SET .Receivable_Control_acct from TSPL_CUSTOMER_ACCOUNT_SET where TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account =TSPL_CUSTOMER_master.Cust_Account) as Acct,(select Description from TSPL_GL_ACCOUNTS where Account_Code =(select TSPL_CUSTOMER_ACCOUNT_SET .Receivable_Control_acct  from TSPL_CUSTOMER_ACCOUNT_SET where TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account =TSPL_CUSTOMER_master.Cust_Account))as AcctDesc,0 as DbtAmt,TSPL_Receipt_Adjustment_Header.Adjustment_Amount as CrAmt ,TSPL_Receipt_Adjustment_Header .Comp_Code,TSPL_Receipt_Adjustment_Header .Doc_No from TSPL_Receipt_Adjustment_Header left outer join TSPL_CUSTOMER_MASTER on TSPL_Receipt_Adjustment_Header .Customer_No =tspl_customer_master.Cust_Code  where TSPL_Receipt_Adjustment_Header .Doc_No  in (" & clsCommon.GetMulcallString(arrDocNo1) & "))as xx )as finalQry left outer join TSPL_COMPANY_MASTER on finalQry .Comp_Code =TSPL_COMPANY_MASTER .Comp_Code "
            '------ Code Ends Here --------------
            '' added By Richa Agarwal 29 aug,2018 BHA/28/08/18-000487   For SubReport
            Dim InvoiceSubReport As String = "Select TSPL_RECEIPT_DETAIL.Receipt_No ,TSPL_RECEIPT_DETAIL.Document_No ,convert(varchar,TSPL_RECEIPT_DETAIL.Document_Date,103) as Document_Date,TSPL_RECEIPT_DETAIL.Original_Amt ,TSPL_RECEIPT_DETAIL.Applied_Amount From TSPL_RECEIPT_DETAIL left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_DETAIL.Receipt_No =TSPL_RECEIPT_HEADER.Receipt_No where TSPL_RECEIPT_HEADER.Receipt_Type  <>'M' AND TSPL_RECEIPT_DETAIL.Receipt_No in (" & clsCommon.GetMulcallString(arrreceipt) & ") "

            Dim frmcrystal As New frmCrystalReportViewer()
            Dim strReceiptDate As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Receipt_Date from TSPL_RECEIPT_HEADER where Receipt_No='" + recvalue1 + "'"))

            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDFCF") = CompairStringResult.Equal Then
                frmcrystal.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.SalesReport, dttemp, clsDBFuncationality.GetDataTable(adjustSubReport), "receipt1", "Receipt Report", clsCommon.myCDate(strReceiptDate), "AdjustmentSubReport.rpt", "rptReceiptDetailWithInvoice.rpt", clsDBFuncationality.GetDataTable(InvoiceSubReport))

            Else
                frmcrystal.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.SalesReport, dttemp, clsDBFuncationality.GetDataTable(adjustSubReport), "receipt", "Receipt Report", clsCommon.myCDate(strReceiptDate), "AdjustmentSubReport.rpt", "rptReceiptDetailWithInvoice.rpt", clsDBFuncationality.GetDataTable(InvoiceSubReport))

            End If

            ' frmcrystal.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.SalesReport, dttemp, clsDBFuncationality.GetDataTable(adjustSubReport), "receipt", "Receipt Report", clsCommon.myCDate(strReceiptDate), "AdjustmentSubReport.rpt", "rptReceiptDetailWithInvoice.rpt", clsDBFuncationality.GetDataTable(InvoiceSubReport))
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message.ToString())
        End Try
    End Sub


    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub
    Private Sub Frmreceiptvoucher2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtpfrom.Value = clsCommon.GETSERVERDATE().Date
        dtpto.Value = clsCommon.GETSERVERDATE().Date
        chkCustomerAll.IsChecked = True
        chkReceiptAll.IsChecked = True
        LoadReceipt()
        LoadCustomer()
        loadlocation()
        ButtonToolTip.SetToolTip(btnprint, "Press Ctrl+P for Print")
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        cbgLocation.DataSource = clsLocation.GetLocationSegments()
        SetUserMgmtNew()
    End Sub
    Sub loadlocation()
        ' Dim qry As String = " select Segment_code as Code, Description  from TSPL_GL_SEGMENT_CODE where Seg_No='7' "
        cbgLocation.DataSource = clsLocation.GetLocationSegments()
        cbgLocation.ValueMember = "Code"
        cbgLocation.DisplayMember = "Name"
    End Sub
    Sub LoadReceipt()
        Dim strquery As String = "select Receipt_No as [Receipt No], Receipt_Date as [Receipt Date] from TSPL_RECEIPT_HEADER where (Receipt_Post_Date >= Convert(DATE,'" + dtpfrom.Value + "',103) AND Receipt_Post_Date <= CONVERT(DATE,'" + dtpto.Value + "',103))"
        cbgReceipt.DataSource = clsDBFuncationality.GetDataTable(strquery)
        cbgReceipt.ValueMember = "Receipt No"
        cbgReceipt.DisplayMember = "Description"
    End Sub
    Sub LoadCustomer()
        Dim strquery As String = "select cust_code as [Customer Code], Customer_Name as [Customer Name] from tspl_customer_master"
        cbgCustomer.DataSource = clsDBFuncationality.GetDataTable(strquery)
        cbgCustomer.ValueMember = "Customer Code"
        cbgCustomer.DisplayMember = "Customer Name"
    End Sub
    Private Sub chkReceiptAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkReceiptAll.ToggleStateChanged, chkReceiptSelect.ToggleStateChanged
        cbgReceipt.Enabled = Not chkReceiptAll.IsChecked
        If chkReceiptSelect.IsChecked Then
            chkCustomerAll.IsChecked = chkReceiptSelect.IsChecked
        End If
    End Sub
    Private Sub chkCustomerAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkCustomerAll.ToggleStateChanged, chkCustomerSelect.ToggleStateChanged
        cbgCustomer.Enabled = Not chkCustomerAll.IsChecked
        If chkCustomerSelect.IsChecked Then
            chkReceiptAll.IsChecked = chkCustomerSelect.IsChecked
        End If
    End Sub
    Private Sub dtpfrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpfrom.ValueChanged
        LoadReceipt()
    End Sub
    Private Sub dtpto_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpto.ValueChanged
        LoadReceipt()
    End Sub

    Private Sub Frmreceiptvoucher2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.P Then
            Print()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag Then
            'SaveData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isPostFlag Then
            'DeleteData()
        ElseIf e.Alt And e.KeyCode = Keys.C Then
            Close()
       
        End If
    End Sub


    Private Sub cbgLocation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbgLocation.Load

    End Sub

    Private Sub chkLocAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkLocAll.ToggleStateChanged
        cbgLocation.Enabled = Not chkLocAll.IsChecked
    End Sub
End Class
