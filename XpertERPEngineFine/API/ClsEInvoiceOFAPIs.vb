Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Public Class ClsEInvoiceOFAPIs
#Region "Variables"
    '' Public HttpWebRequest()
#End Region
    Public Shared Function IsGenerateAuthTokenNo_Required(ByVal strCompCode As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As String
        Dim strToken As String = String.Empty
        Try
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            Dim TokenTimeReGenForEinvoice As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.TokenTimeReGenForEinvoice, clsFixedParameterCode.TokenTimeReGenForEinvoice, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_EInvoiceToken where Location_Code='" & strLocation & "'", trans)
                    If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                        'Dim isGenerateTokenRequired As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where  DATEDIFF (minute,getdate(),dateadd(MINUTE ,360,ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,360,ResponseTime ))<=58", trans))
                        'Dim isGenerateTokenRequired As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where Location_Code='" & strLocation & "' and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,60,ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,60,ResponseTime ))<=58", trans))
                        Dim isGenerateTokenRequired As String = String.Empty
                        If clsCommon.myCdbl(TokenTimeReGenForEinvoice) > 0 Then
                            isGenerateTokenRequired = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where Location_Code='" & strLocation & "' and  DATEDIFF (minute,getdate(),dateadd(MINUTE ," & TokenTimeReGenForEinvoice & ",ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ," & TokenTimeReGenForEinvoice & ",ResponseTime ))<=" & clsCommon.myCdbl(TokenTimeReGenForEinvoice) - 2 & "", trans))
                        Else
                            isGenerateTokenRequired = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where Location_Code='" & strLocation & "' and convert(smalldatetime,expirytime,103)>convert(smalldatetime,getdate(),103)", trans))
                        End If
                        If clsCommon.myLen(isGenerateTokenRequired) > 0 AndAlso clsCommon.CompairString(isGenerateTokenRequired, "No") = CompairStringResult.Equal Then
                            strToken = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select AuthToken from TSPL_EInvoiceToken where Location_Code='" & strLocation & "'", trans))
                        Else
                            GenerateAuthTokenNo(strCompCode, strLocation, trans)
                            strToken = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select AuthToken from TSPL_EInvoiceToken where Location_Code='" & strLocation & "'", trans))
                        End If
                    Else
                        GenerateAuthTokenNo(strCompCode, strLocation, trans)
                        strToken = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select AuthToken from TSPL_EInvoiceToken where Location_Code='" & strLocation & "'", trans))
                    End If
                Else
                    strToken = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 client_secret As TokenId from tspl_einvoiceHeader_info where VendorName='CLEARTAX' and Location_Code='" & strLocation & "'", trans))
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return strToken
    End Function
    Public Shared Function GenerateAuthTokenNo(ByVal strCompCode As String, ByVal strLocation As String, ByVal trans As SqlTransaction)
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateAuthToken_Get' and Location_Code='" & strLocation & "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Const Tls12Value As Integer = &HC00
                Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                ' Set the security protocol globally for all HTTPS requests
                System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                httpRequest.ContentType = "application/json"
                httpRequest.Headers.Add("username", clsCommon.myCstr(dt.Rows(0)("username")))
                httpRequest.Headers.Add("password", clsCommon.myCstr(dt.Rows(0)("password")))
                httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                httpRequest.Method = "GET"
                httpRequest.KeepAlive = True
                httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                '' required validation httresponse
                Dim strResponseDate As DateTime = clsCommon.GETSERVERDATE(trans)
                Dim strResult As String = String.Empty
                Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                    strResult = streamReader.ReadToEnd()
                End Using
                Dim jObj As JObject = JObject.Parse(strResult)
                'Dim strToken As String = (jObj.SelectToken("data")).SelectToken("AuthToken").ToString()
                Dim objResult As Object = Nothing
                Dim strExpiryDate As DateTime?
                Dim strToken As String = String.Empty
                If TypeOf jObj.SelectToken("data") Is JObject Then
                    objResult = jObj.SelectToken("data")
                    strToken = objResult.SelectToken("AuthToken").ToString
                    strExpiryDate = objResult.SelectToken("TokenExpiry").ToString
                Else
                    Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                    Throw New Exception(strStatusDesc)
                End If
                ''insert token and datetime value into tspl_einnvoice_token table
                If clsCommon.myLen(strToken) > 0 Then
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_EInvoiceToken where Location_Code='" & strLocation & "' ", trans)
                    If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                        clsDBFuncationality.ExecuteNonQuery("update TSPL_EInvoiceToken set AuthToken='" & strToken & "',ResponseTime='" & clsCommon.GetPrintDate(strResponseDate, "dd/MMM/yyyy hh:mm tt") & "',ExpiryTime='" & clsCommon.GetPrintDate(strExpiryDate, "dd/MMM/yyyy hh:mm tt") & "' where Location_Code='" & strLocation & "'", trans)
                    Else
                        clsDBFuncationality.ExecuteNonQuery("Insert Into TSPL_EInvoiceToken values ('" & strToken & "','" & clsCommon.GetPrintDate(strResponseDate, "dd/MMM/yyyy hh:mm tt") & "','" & strLocation & "','" & clsCommon.GetPrintDate(strExpiryDate, "dd/MMM/yyyy hh:mm tt") & "')", trans)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostAuthTokenNo_withInvoiceData_ForCustomerType_BC(ByVal strCompCode As String, ByVal strTokenNo As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateDynamicQRCode' and VendorName='MASTERGST' and Location_Code='" & strLocation & "'"
                Else
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateDynamicQRCode' and VendorName='CLEARTAX' and Location_Code='" & strLocation & "'"
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("username", clsCommon.myCstr(dt.Rows(0)("username")))
                        httpRequest.Headers.Add("sgstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        Dim dtDetails As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
                        If dtDetails IsNot Nothing AndAlso dtDetails.Rows.Count > 0 Then
                            httpRequest.Headers.Add("docno", clsCommon.myCstr(dtDetails.Rows(0)("docno")))
                            httpRequest.Headers.Add("docdate", clsCommon.myCstr(dtDetails.Rows(0)("docdate")))
                            httpRequest.Headers.Add("totinvval", clsCommon.myCstr(dtDetails.Rows(0)("totinvval")))
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("upiid"))) > 0 Then
                                httpRequest.Headers.Add("upiid", clsCommon.myCstr(dtDetails.Rows(0)("upiid")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("bankaccno"))) > 0 Then
                                httpRequest.Headers.Add("bankaccno", clsCommon.myCstr(dtDetails.Rows(0)("bankaccno")))
                            Else
                                Throw New Exception("Please fill Bank Acc No into Location Master for Location " & clsCommon.myCstr(dtDetails.Rows(0)("Bill_To_Location")) & "")
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("bankifsccode"))) > 0 Then
                                httpRequest.Headers.Add("bankifsccode", clsCommon.myCstr(dtDetails.Rows(0)("bankifsccode")))
                            Else
                                Throw New Exception("Please fill Bank IFSC Code into Location Master for Location " & clsCommon.myCstr(dtDetails.Rows(0)("Bill_To_Location")) & "")
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("accountholdername"))) > 0 Then
                                httpRequest.Headers.Add("accountholdername", clsCommon.myCstr(dtDetails.Rows(0)("accountholdername")))
                            Else
                                Throw New Exception("Please fill A/c holder Name into Location Master for Location " & clsCommon.myCstr(dtDetails.Rows(0)("Bill_To_Location")) & "")
                            End If
                            httpRequest.Headers.Add("igstamount", clsCommon.myCstr(dtDetails.Rows(0)("igstamount")))
                            httpRequest.Headers.Add("cgstamount", clsCommon.myCstr(dtDetails.Rows(0)("cgstamount")))
                            httpRequest.Headers.Add("sgstamount", clsCommon.myCstr(dtDetails.Rows(0)("sgstamount")))
                            httpRequest.Headers.Add("cessamount", clsCommon.myCstr(dtDetails.Rows(0)("cessamount")))
                        End If
                        httpRequest.Method = "GET"
                    Else
                        httpRequest.Headers.Add("x-cleartax-auth-token", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("owner_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "PUT"
                    End If
                    httpRequest.KeepAlive = True
                    ''Dim StreamWriter As StreamWriter = New StreamWriter(httpRequest.GetRequestStream())
                    ''------
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult)
                        If TypeOf jObj.SelectToken("data") Is JObject Then
                            objResult = jObj.SelectToken("data")
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function IsEWayBillAPIAuthenticate(ByVal strCompCode As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Dim IsAuthenticate As Boolean = False
        Try
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            Dim TokenTimeReGenForEWayBill As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.TokenTimeReGenForEWayBill, clsFixedParameterCode.TokenTimeReGenForEWayBill, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_EWayBillToken where Location_Code='" & strLocation & "'", trans)
                    If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                        'Dim isGenerateTokenRequired As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where  DATEDIFF (minute,getdate(),dateadd(MINUTE ,360,ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,360,ResponseTime ))<=58", trans))
                        'Dim isGenerateTokenRequired As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EINVOICETOKEN where Location_Code='" & strLocation & "' and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,60,ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ,60,ResponseTime ))<=58", trans))
                        Dim isGenerateTokenRequired As String = String.Empty
                        If clsCommon.myCdbl(TokenTimeReGenForEWayBill) > 0 Then
                            isGenerateTokenRequired = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EWayBillToken where Location_Code='" & strLocation & "' and  DATEDIFF (minute,getdate(),dateadd(MINUTE ," & TokenTimeReGenForEWayBill & ",ResponseTime ))>0 and  DATEDIFF (minute,getdate(),dateadd(MINUTE ," & TokenTimeReGenForEWayBill & ",ResponseTime ))<=" & clsCommon.myCdbl(TokenTimeReGenForEWayBill) - 2 & "", trans))
                        Else
                            isGenerateTokenRequired = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'No' from TSPL_EWayBillToken where Location_Code='" & strLocation & "' and convert(smalldatetime,expirytime,103)>convert(smalldatetime,getdate(),103)", trans))
                        End If
                        If clsCommon.myLen(isGenerateTokenRequired) > 0 AndAlso clsCommon.CompairString(isGenerateTokenRequired, "No") = CompairStringResult.Equal Then
                            IsAuthenticate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsAuthenticate from TSPL_EWayBillToken where Location_Code='" & strLocation & "'", trans))
                        Else
                            IsAuthenticate = GetEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                            'IsAuthenticate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsAuthenticate from TSPL_EWayBillToken where Location_Code='" & strLocation & "'", trans))
                        End If
                    Else
                        IsAuthenticate = GetEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                        'IsAuthenticate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsAuthenticate from TSPL_EWayBillToken where Location_Code='" & strLocation & "'", trans))
                    End If
                Else
                    'strToken = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 client_secret As TokenId from tspl_einvoiceHeader_info where VendorName='CLEARTAX' and Location_Code='" & strLocation & "'", trans))
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return IsAuthenticate
    End Function
    Public Shared Function GetEWayBillAPIAuthenticate(ByVal strCompCode As String, ByVal strLocation As String, ByVal trans As SqlTransaction)
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='authenticateEwayBill_Get' and Location_Code='" & strLocation & "' ", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                ' Set the TLS protocol version
                Const Tls12Value As Integer = &HC00
                Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                ' Set the security protocol globally for all HTTPS requests
                System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
                'ServicePointManager.SecurityProtocol = SecurityProtocolType.Or SecurityProtocolType.Tls11
                httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                httpRequest.ContentType = "application/json"
                httpRequest.Accept = "application/json"
                httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                httpRequest.Method = "GET"
                httpRequest.KeepAlive = True
                httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                '' required validation httresponse
                Dim strResponseDate As DateTime = clsCommon.GETSERVERDATE(trans)
                Dim strResult As String = String.Empty
                Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                    strResult = streamReader.ReadToEnd()
                End Using
                Dim jObj As JObject = JObject.Parse(strResult)
                Dim objResult As Object = Nothing
                Dim Status_cd As String = String.Empty
                Dim strStatusDesc As String = String.Empty
                If TypeOf jObj.SelectToken("") Is JObject Then
                    objResult = jObj.SelectToken("")
                    Status_cd = objResult.SelectToken("status_cd").ToString
                    If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                        strStatusDesc = (jObj.SelectToken("status_desc")).ToString()
                        Throw New Exception(strStatusDesc)
                    Else
                        Dim strUpdate As String = "update TSPL_EWayBillToken set IsAuthenticate=1,ResponseTime='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") & "',ExpiryTime='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans).AddHours(6), "dd/MMM/yyyy hh:mm tt") & "'  where Location_Code='" & strLocation & "'"
                        clsDBFuncationality.ExecuteNonQuery(strUpdate, trans)
                    End If
                Else
                    strStatusDesc = (jObj.SelectToken("status_desc")).ToString()
                    Throw New Exception(strStatusDesc)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function PostEWayBill(ByVal strCompCode As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Return PostEWayBill(strCompCode, strQry, strLocation, trans, False)
    End Function
    Public Shared Function PostEWayBill(ByVal strCompCode As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction, ByVal stopEWayBill As Boolean) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='genewaybill' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    'Dim jObj As JObject
                    Dim strEwayBillDetails As String = String.Empty
                    Dim dtDetails As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        If dtDetails IsNot Nothing AndAlso dtDetails.Rows.Count > 0 Then
                            If clsCommon.myCdbl(dtDetails(0)("NO_Transporter")) = 0 Then
                                Dim objEWayBill As Object = New clsEwayBillDetail()
                                objEWayBill.supplyType = clsCommon.myCstr(dtDetails(0)("supplyType"))
                                objEWayBill.subSupplyType = clsCommon.myCstr(dtDetails(0)("subSupplyType"))
                                objEWayBill.subSupplyDesc = clsCommon.myCstr(dtDetails(0)("subSupplyDesc"))
                                objEWayBill.docType = clsCommon.myCstr(dtDetails(0)("docType"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("docNo"))) <= 15 Then
                                    objEWayBill.docNo = clsCommon.myCstr(dtDetails(0)("docNo"))
                                Else
                                    Throw New Exception("Document No Should <=15")
                                End If
                                objEWayBill.docDate = clsCommon.myCstr(clsCommon.GetPrintDate(dtDetails(0)("docDate"), "dd/MM/yyyy"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("fromGstin"))) = 15 Then
                                    objEWayBill.fromGstin = clsCommon.myCstr(dtDetails(0)("fromGstin"))
                                Else
                                    Throw New Exception("From GSTIN No. not Found/Invalid!")
                                End If
                                objEWayBill.fromTrdName = clsCommon.myCstr(dtDetails(0)("fromTrdName"))
                                objEWayBill.fromAddr1 = clsCommon.myCstr(dtDetails(0)("fromAddr1"))
                                objEWayBill.fromAddr2 = clsCommon.myCstr(dtDetails(0)("fromAddr2"))
                                objEWayBill.fromPlace = clsCommon.myCstr(dtDetails(0)("fromPlace"))
                                objEWayBill.actFromStateCode = clsCommon.myCdbl(dtDetails(0)("actFromStateCode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("fromPincode"))) = 6 Then
                                    objEWayBill.fromPincode = clsCommon.myCdbl(dtDetails(0)("fromPincode"))
                                Else
                                    Throw New Exception("From Pincode not Found/Invalid!")
                                End If
                                objEWayBill.fromStateCode = clsCommon.myCdbl(dtDetails(0)("fromStateCode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("toGstin"))) = 15 OrElse clsCommon.CompairString(clsCommon.myCstr(dtDetails.Rows(0)("toGstin")), "URP") = CompairStringResult.Equal Then
                                    objEWayBill.toGstin = clsCommon.myCstr(dtDetails(0)("toGstin"))
                                Else
                                    Throw New Exception("To GSTIN No. not Found/Invalid!")
                                End If
                                objEWayBill.toTrdName = clsCommon.myCstr(dtDetails(0)("toTrdName"))
                                objEWayBill.toAddr1 = clsCommon.myCstr(dtDetails(0)("toAddr1"))
                                objEWayBill.toAddr2 = clsCommon.myCstr(dtDetails(0)("toAddr2"))
                                objEWayBill.toPlace = clsCommon.myCstr(dtDetails(0)("toPlace"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("toPincode"))) = 6 Then
                                    objEWayBill.toPincode = clsCommon.myCdbl(dtDetails(0)("toPincode"))
                                Else
                                    Throw New Exception("To Pincode not Found/Invalid!")
                                End If
                                objEWayBill.actToStateCode = clsCommon.myCdbl(dtDetails(0)("actToStateCode"))
                                objEWayBill.toStateCode = clsCommon.myCdbl(dtDetails(0)("toStateCode"))
                                objEWayBill.transactionType = clsCommon.myCdbl(dtDetails(0)("transactionType"))
                                'objEWayBill.dispatchFromGSTIN = clsCommon.myCstr(dtDetails(0)("dispatchFromGSTIN"))
                                objEWayBill.dispatchFromTradeName = clsCommon.myCstr(dtDetails(0)("dispatchFromTradeName"))
                                'objEWayBill.shipToGSTIN = clsCommon.myCstr(dtDetails(0)("shipToGSTIN"))
                                objEWayBill.shipToTradeName = clsCommon.myCstr(dtDetails(0)("shipToTradeName"))
                                objEWayBill.totalValue = clsCommon.myCdbl(dtDetails(0)("totalValue"))
                                objEWayBill.cgstValue = clsCommon.myCdbl(dtDetails(0)("cgstValue"))
                                objEWayBill.sgstValue = clsCommon.myCdbl(dtDetails(0)("sgstValue"))
                                objEWayBill.igstValue = clsCommon.myCdbl(dtDetails(0)("igstValue"))
                                objEWayBill.cessValue = clsCommon.myCdbl(dtDetails(0)("cessValue"))
                                objEWayBill.cessNonAdvolValue = clsCommon.myCdbl(dtDetails(0)("cessNonAdvolValue"))
                                objEWayBill.totInvValue = clsCommon.myCdbl(dtDetails(0)("totInvValue"))
                                objEWayBill.transMode = clsCommon.myCstr(dtDetails(0)("transMode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("transDistance"))) > 0 Then
                                    objEWayBill.transDistance = clsCommon.myCstr(dtDetails(0)("transDistance"))
                                End If
                                objEWayBill.transporterName = clsCommon.myCstr(dtDetails(0)("transporterName"))
                                objEWayBill.transporterId = clsCommon.myCstr(dtDetails(0)("transporterId"))
                                objEWayBill.transDocNo = clsCommon.myCstr(dtDetails(0)("transDocNo"))
                                objEWayBill.transDocDate = clsCommon.myCstr(dtDetails(0)("transDocDate"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("vehicleNo"))) > 0 Then
                                    objEWayBill.vehicleNo = clsCommon.myCstr(dtDetails(0)("vehicleNo"))
                                Else
                                    Throw New Exception("Vehicle No Not Found!")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("vehicleType"))) > 0 Then
                                    objEWayBill.vehicleType = clsCommon.myCstr(dtDetails(0)("vehicleType"))
                                Else
                                    Throw New Exception("Vehicle type Not Found!")
                                End If
                                objEWayBill.ItemList = New List(Of clsItem)
                                For Each dr As DataRow In dtDetails.Rows
                                    Dim item As clsItem = New clsItem()
                                    item.productName = clsCommon.myCstr(dr("productName"))
                                    item.productDesc = clsCommon.myCstr(dr("productDesc"))
                                    item.hsnCode = clsCommon.myCdbl(dr("hsnCode"))
                                    item.quantity = clsCommon.myCdbl(dr("quantity"))
                                    item.qtyUnit = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(GST_UNIT_CODE ,'') FROM TSPL_UNIT_MASTER WHERE UNIT_CODE='" & clsCommon.myCstr(dr("qtyUnit")).ToUpper & "'", trans))
                                    item.taxableAmount = clsCommon.myCdbl(dr("taxableAmount"))
                                    item.sgstRate = clsCommon.myCdbl(dr("sgstRate"))
                                    item.cgstRate = clsCommon.myCdbl(dr("cgstRate"))
                                    item.igstRate = clsCommon.myCdbl(dr("igstRate"))
                                    item.cessRate = clsCommon.myCdbl(dr("cessRate"))
                                    ' Add item to the list
                                    objEWayBill.itemList.Add(item)
                                Next
                                strEwayBillDetails = JsonConvert.SerializeObject(objEWayBill)
                            Else
                                Dim objEWayBill As Object = New clsEwayBillDetailWithoutTrans()
                                objEWayBill.supplyType = clsCommon.myCstr(dtDetails(0)("supplyType"))
                                objEWayBill.subSupplyType = clsCommon.myCstr(dtDetails(0)("subSupplyType"))
                                objEWayBill.subSupplyDesc = clsCommon.myCstr(dtDetails(0)("subSupplyDesc"))
                                objEWayBill.docType = clsCommon.myCstr(dtDetails(0)("docType"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("docNo"))) <= 15 Then
                                    objEWayBill.docNo = clsCommon.myCstr(dtDetails(0)("docNo"))
                                Else
                                    Throw New Exception("Document No Should <=15")
                                End If
                                objEWayBill.docDate = clsCommon.myCstr(clsCommon.GetPrintDate(dtDetails(0)("docDate"), "dd/MM/yyyy"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("fromGstin"))) = 15 Then
                                    objEWayBill.fromGstin = clsCommon.myCstr(dtDetails(0)("fromGstin"))
                                Else
                                    Throw New Exception("From GSTIN No. not Found/Invalid!")
                                End If
                                objEWayBill.fromTrdName = clsCommon.myCstr(dtDetails(0)("fromTrdName"))
                                objEWayBill.fromAddr1 = clsCommon.myCstr(dtDetails(0)("fromAddr1"))
                                objEWayBill.fromAddr2 = clsCommon.myCstr(dtDetails(0)("fromAddr2"))
                                objEWayBill.fromPlace = clsCommon.myCstr(dtDetails(0)("fromPlace"))
                                objEWayBill.actFromStateCode = clsCommon.myCdbl(dtDetails(0)("actFromStateCode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("fromPincode"))) = 6 Then
                                    objEWayBill.fromPincode = clsCommon.myCdbl(dtDetails(0)("fromPincode"))
                                Else
                                    Throw New Exception("From Pincode not Found/Invalid!")
                                End If
                                objEWayBill.fromStateCode = clsCommon.myCdbl(dtDetails(0)("fromStateCode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("toGstin"))) = 15 OrElse clsCommon.CompairString(clsCommon.myCstr(dtDetails.Rows(0)("toGstin")), "URP") = CompairStringResult.Equal Then
                                    objEWayBill.toGstin = clsCommon.myCstr(dtDetails(0)("toGstin"))
                                Else
                                    Throw New Exception("To GSTIN No. not Found/Invalid!")
                                End If
                                objEWayBill.toTrdName = clsCommon.myCstr(dtDetails(0)("toTrdName"))
                                objEWayBill.toAddr1 = clsCommon.myCstr(dtDetails(0)("toAddr1"))
                                objEWayBill.toAddr2 = clsCommon.myCstr(dtDetails(0)("toAddr2"))
                                objEWayBill.toPlace = clsCommon.myCstr(dtDetails(0)("toPlace"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("toPincode"))) = 6 Then
                                    objEWayBill.toPincode = clsCommon.myCdbl(dtDetails(0)("toPincode"))
                                Else
                                    Throw New Exception("To Pincode not Found/Invalid!")
                                End If
                                objEWayBill.actToStateCode = clsCommon.myCdbl(dtDetails(0)("actToStateCode"))
                                objEWayBill.toStateCode = clsCommon.myCdbl(dtDetails(0)("toStateCode"))
                                objEWayBill.transactionType = clsCommon.myCdbl(dtDetails(0)("transactionType"))
                                'objEWayBill.dispatchFromGSTIN = clsCommon.myCstr(dtDetails(0)("dispatchFromGSTIN"))
                                objEWayBill.dispatchFromTradeName = clsCommon.myCstr(dtDetails(0)("dispatchFromTradeName"))
                                'objEWayBill.shipToGSTIN = clsCommon.myCstr(dtDetails(0)("shipToGSTIN"))
                                objEWayBill.shipToTradeName = clsCommon.myCstr(dtDetails(0)("shipToTradeName"))
                                objEWayBill.totalValue = clsCommon.myCdbl(dtDetails(0)("totalValue"))
                                objEWayBill.cgstValue = clsCommon.myCdbl(dtDetails(0)("cgstValue"))
                                objEWayBill.sgstValue = clsCommon.myCdbl(dtDetails(0)("sgstValue"))
                                objEWayBill.igstValue = clsCommon.myCdbl(dtDetails(0)("igstValue"))
                                objEWayBill.cessValue = clsCommon.myCdbl(dtDetails(0)("cessValue"))
                                objEWayBill.cessNonAdvolValue = clsCommon.myCdbl(dtDetails(0)("cessNonAdvolValue"))
                                objEWayBill.totInvValue = clsCommon.myCdbl(dtDetails(0)("totInvValue"))
                                objEWayBill.transMode = clsCommon.myCstr(dtDetails(0)("transMode"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("transDistance"))) > 0 Then
                                    objEWayBill.transDistance = clsCommon.myCstr(dtDetails(0)("transDistance"))
                                End If
                                'objEWayBill.transporterName = clsCommon.myCstr(dtDetails(0)("transporterName"))
                                'objEWayBill.transporterId = clsCommon.myCstr(dtDetails(0)("transporterId"))
                                'objEWayBill.transDocNo = clsCommon.myCstr(dtDetails(0)("transDocNo"))
                                objEWayBill.transDocDate = clsCommon.myCstr(dtDetails(0)("transDocDate"))
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("vehicleNo"))) > 0 Then
                                    objEWayBill.vehicleNo = clsCommon.myCstr(dtDetails(0)("vehicleNo"))
                                Else
                                    Throw New Exception("Vehicle No Not Found!")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(dtDetails(0)("vehicleType"))) > 0 Then
                                    objEWayBill.vehicleType = clsCommon.myCstr(dtDetails(0)("vehicleType"))
                                Else
                                    Throw New Exception("Vehicle type Not Found!")
                                End If
                                objEWayBill.ItemList = New List(Of clsItem)
                                For Each dr As DataRow In dtDetails.Rows
                                    Dim item As clsItem = New clsItem()
                                    item.productName = clsCommon.myCstr(dr("productName"))
                                    item.productDesc = clsCommon.myCstr(dr("productDesc"))
                                    item.hsnCode = clsCommon.myCdbl(dr("hsnCode"))
                                    item.quantity = clsCommon.myCdbl(dr("quantity"))
                                    item.qtyUnit = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(GST_UNIT_CODE ,'') FROM TSPL_UNIT_MASTER WHERE UNIT_CODE='" & clsCommon.myCstr(dr("qtyUnit")).ToUpper & "'", trans))
                                    item.taxableAmount = clsCommon.myCdbl(dr("taxableAmount"))
                                    item.sgstRate = clsCommon.myCdbl(dr("sgstRate"))
                                    item.cgstRate = clsCommon.myCdbl(dr("cgstRate"))
                                    item.igstRate = clsCommon.myCdbl(dr("igstRate"))
                                    item.cessRate = clsCommon.myCdbl(dr("cessRate"))
                                    ' Add item to the list
                                    objEWayBill.itemList.Add(item)
                                Next
                                strEwayBillDetails = JsonConvert.SerializeObject(objEWayBill)
                            End If

                            streamWriter.Write(strEwayBillDetails)
                            End If
                    End Using
                    Try
                        Dim coll As New Hashtable()
                        clsCommon.AddColumnsForChange(coll, "Document_Code", clsCommon.myCstr(dtDetails.Rows(0)("DocNo")))
                        clsCommon.AddColumnsForChange(coll, "JSON", strEwayBillDetails)
                        clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                        clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_INVOICE_JSON", OMInsertOrUpdate.Insert, "", trans)
                    Catch ex As Exception
                    End Try
                    ''------
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult1 As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult1 = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult1)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult1)
                        If TypeOf jObj.SelectToken("") Is JObject Then
                            objResult = jObj.SelectToken("")
                            Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                            If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                                Dim statusDec As String = objResult.SelectToken("error").ToString
                                objResult = Nothing
                                Throw New Exception(statusDec)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function

    Public Shared Function PostAuthTokenNo_withInvoiceData(ByVal strCompCode As String, ByVal strTokenNo As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Return PostAuthTokenNo_withInvoiceData(strCompCode, strTokenNo, strQry, strLocation, trans, False)
    End Function
    Public Shared Function PostAuthTokenNo_withInvoiceData(ByVal strCompCode As String, ByVal strTokenNo As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction, ByVal stopEWayBill As Boolean) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateAuthToken_IRN' and VendorName='MASTERGST' and Location_Code='" & strLocation & "'"
                Else
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateAuthToken_IRN' and VendorName='CLEARTAX' and Location_Code='" & strLocation & "'"
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("username", clsCommon.myCstr(dt.Rows(0)("username")))
                        httpRequest.Headers.Add("auth-token", strTokenNo)
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    Else
                        httpRequest.Headers.Add("x-cleartax-auth-token", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("owner_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "PUT"
                    End If
                    httpRequest.KeepAlive = True
                    'Dim jObj As JObject
                    Dim strInvoiceDetails As String = String.Empty
                    Dim dtDetails As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        If dtDetails IsNot Nothing AndAlso dtDetails.Rows.Count > 0 Then
                            Dim objInvDetails As Object
                            If objCommonVar.GenerateEWayBillWithEInvoice = True AndAlso Not stopEWayBill Then
                                objInvDetails = New clsEInvoiceDetailsWithEWayBill()
                            Else
                                objInvDetails = New clsEInvoiceDetails()
                            End If
                            objInvDetails.Version = "1.1"
                            objInvDetails.ItemList = New List(Of ClsItemList)
                            ''Transaction Info
                            objInvDetails.TranDtls.Add("TaxSch", "GST")
                            objInvDetails.TranDtls.Add("SupTyp", clsCommon.myCstr(dtDetails.Rows(0)("SupTyp")))
                            objInvDetails.TranDtls.Add("RegRev", "N")
                            objInvDetails.TranDtls.Add("IgstOnIntra", clsCommon.myCstr(dtDetails.Rows(0)("IgstOnIntra")))
                            ''Document info
                            objInvDetails.DocDtls.Add("Typ", clsCommon.myCstr(dtDetails.Rows(0)("DocType")))
                            objInvDetails.DocDtls.Add("No", clsCommon.myCstr(dtDetails.Rows(0)("DocNo")))
                            objInvDetails.DocDtls.Add("Dt", clsCommon.myCDate(dtDetails.Rows(0)("DocDate")))
                            ''Seller Info 
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerGSTINNo"))) = 15 Then
                                objInvDetails.SellerDtls.Add("Gstin", clsCommon.myCstr(dtDetails.Rows(0)("SellerGSTINNo")))
                            Else
                                Throw New Exception("Seller GSTIN No. not Found/Invalid!")
                            End If
                            objInvDetails.SellerDtls.Add("LglNm", clsCommon.myCstr(dtDetails.Rows(0)("SellerLglNm")))
                            objInvDetails.SellerDtls.Add("TrdNm", clsCommon.myCstr(dtDetails.Rows(0)("SellerTrdNm")))
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerAdd1"))) > 0 Then
                                objInvDetails.SellerDtls.Add("Addr1", clsCommon.myCstr(dtDetails.Rows(0)("SellerAdd1")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerAdd2"))) > 0 Then
                                objInvDetails.SellerDtls.Add("Addr2", clsCommon.myCstr(dtDetails.Rows(0)("SellerAdd2")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerLoc"))) > 0 Then
                                objInvDetails.SellerDtls.Add("Loc", clsCommon.myCstr(dtDetails.Rows(0)("SellerLoc")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerPincode"))) = 6 Then
                                objInvDetails.SellerDtls.Add("Pin", dtDetails.Rows(0)("SellerPincode"))
                            Else
                                Throw New Exception("Seller Pincode not Found/Invalid!")
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerStcd"))) > 0 Then
                                objInvDetails.SellerDtls.Add("Stcd", clsCommon.myCstr(dtDetails.Rows(0)("SellerStcd")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerPhone"))) > 0 Then
                                Dim strPhone As String = clsCommon.myCstr(dtDetails.Rows(0)("SellerPhone"))
                                If strPhone.Contains(")") Then
                                    Dim strindex As Integer = strPhone.IndexOf(")") + 1
                                    strPhone = strPhone.Substring(strindex, strPhone.Length - strindex)
                                    strPhone = strPhone.Replace("_", "")
                                Else
                                    strPhone = strPhone.Replace("_", "")
                                End If
                                If clsCommon.myLen(strPhone) > 0 Then
                                    objInvDetails.SellerDtls.Add("Ph", strPhone)
                                End If
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("SellerEmail"))) > 0 Then
                                objInvDetails.SellerDtls.Add("Em", clsCommon.myCstr(dtDetails.Rows(0)("SellerEmail")))
                            End If
                            ''Buyer Info 
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerGSTINNo"))) = 15 Then
                                objInvDetails.BuyerDtls.Add("Gstin", clsCommon.myCstr(dtDetails.Rows(0)("BuyerGSTINNo")))
                            Else
                                Throw New Exception("Buyer GSTIN No. not Found/Invalid!")
                            End If
                            objInvDetails.BuyerDtls.Add("LglNm", clsCommon.myCstr(dtDetails.Rows(0)("BuyerLglNm")))
                            objInvDetails.BuyerDtls.Add("TrdNm", clsCommon.myCstr(dtDetails.Rows(0)("BuyerTrdNm")))
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerPOS"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Pos", clsCommon.myCstr(dtDetails.Rows(0)("BuyerPOS")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerAdd1"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Addr1", clsCommon.myCstr(dtDetails.Rows(0)("BuyerAdd1")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerAdd2"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Addr2", clsCommon.myCstr(dtDetails.Rows(0)("BuyerAdd2")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerLoc"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Loc", clsCommon.myCstr(dtDetails.Rows(0)("BuyerLoc")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerPincode"))) = 6 Then
                                objInvDetails.BuyerDtls.Add("Pin", dtDetails.Rows(0)("BuyerPincode"))
                            Else
                                Throw New Exception("Buyer Pincode not Found/Invalid!")
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerStcd"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Stcd", clsCommon.myCstr(dtDetails.Rows(0)("BuyerStcd")))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerPhone"))) > 0 Then
                                Dim strPhone As String = clsCommon.myCstr(dtDetails.Rows(0)("BuyerPhone"))
                                If strPhone.Contains(")") Then
                                    Dim strindex As Integer = strPhone.IndexOf(")") + 1
                                    strPhone = strPhone.Substring(strindex, strPhone.Length - strindex)
                                    strPhone = strPhone.Replace("_", "")
                                Else
                                    strPhone = strPhone.Replace("_", "")
                                End If
                                If clsCommon.myLen(strPhone) > 0 Then
                                    objInvDetails.BuyerDtls.Add("Ph", strPhone)
                                End If
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("BuyerEmail"))) > 0 Then
                                objInvDetails.BuyerDtls.Add("Em", clsCommon.myCstr(dtDetails.Rows(0)("BuyerEmail")))
                            End If
                            For Each dr As DataRow In dtDetails.Rows
                                Dim objItemList As New ClsItemList()
                                objItemList.SlNo = clsCommon.myCstr(dr("ItemSlNo"))
                                objItemList.IsServc = clsCommon.myCstr(dr("ItemIsServc"))
                                objItemList.PrdDesc = clsCommon.myCstr(dr("ItemPrdDesc"))
                                objItemList.HsnCd = clsCommon.myCstr(dr("ItemHsnCd"))
                                objItemList.Qty = clsCommon.myCdbl(dr("ItemQty"))
                                If clsCommon.CompairString(clsCommon.myCstr(dr("ItemIsServc")).ToUpper, "Y") = CompairStringResult.Equal Then
                                    objItemList.Unit = "OTH"
                                Else
                                    'If clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "CASE") = CompairStringResult.Equal Then
                                    '    objItemList.Unit = "CTN"
                                    'ElseIf clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "KG") = CompairStringResult.Equal Then
                                    '    objItemList.Unit = "KGS"
                                    'ElseIf clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "JAR") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "POUCH") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "TIN") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "PKT") = CompairStringResult.Equal Then
                                    '    objItemList.Unit = "PCS"
                                    'ElseIf clsCommon.CompairString(clsCommon.myCstr(dr("ItemUnit")).ToUpper, "BOTTLE") = CompairStringResult.Equal Then
                                    '    objItemList.Unit = "BTL"
                                    'Else
                                    '    objItemList.Unit = clsCommon.myCstr(dr("ItemUnit")).ToUpper
                                    'End If
                                    objItemList.Unit = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(GST_UNIT_CODE ,'') FROM TSPL_UNIT_MASTER WHERE UNIT_CODE='" & clsCommon.myCstr(dr("ItemUnit")).ToUpper & "'", trans))
                                End If
                                objItemList.UnitPrice = Math.Round(clsCommon.myCdbl(dr("ItemUnitPrice")), 3)
                                objItemList.TotAmt = Math.Round(clsCommon.myCdbl(dr("ItemTotAmt")), 2)
                                objItemList.Discount = Math.Round(clsCommon.myCdbl(dr("ItemDiscount")), 2)
                                objItemList.AssAmt = Math.Round(clsCommon.myCdbl(dr("ItemAssAmt")), 2)
                                objItemList.GstRt = Math.Round(clsCommon.myCdbl(dr("ItemGstRt")), 2)
                                objItemList.SgstAmt = Math.Round(clsCommon.myCdbl(dr("ItemSgstAmt")), 2)
                                objItemList.IgstAmt = Math.Round(clsCommon.myCdbl(dr("ItemIgstAmt")), 2)
                                objItemList.CgstAmt = Math.Round(clsCommon.myCdbl(dr("ItemCgstAmt")), 2)
                                objItemList.OthChrg = Math.Round(clsCommon.myCdbl(dr("ItemOthChrg")), 2)
                                objItemList.TotItemVal = Math.Round(clsCommon.myCdbl(dr("ItemTotItemVal")), 2)
                                objInvDetails.ItemList.Add(objItemList)
                            Next
                            ''ValDtls info
                            Dim ValDtlsDiscount As Double = 0
                            objInvDetails.ValDtls.Add("AssVal", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsAssVal")), 2))
                            objInvDetails.ValDtls.Add("CgstVal", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsCgstVal")), 2))
                            objInvDetails.ValDtls.Add("SgstVal", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsSgstVal")), 2))
                            objInvDetails.ValDtls.Add("IgstVal", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsIgstVal")), 2))
                            objInvDetails.ValDtls.Add("Discount", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsDiscount")), 2))
                            'objInvDetails.ValDtls.Add("Discount", Math.Round(clsCommon.myCdbl(ValDtlsDiscount), 2))
                            objInvDetails.ValDtls.Add("OthChrg", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsOthChrg")), 2))
                            objInvDetails.ValDtls.Add("TotInvVal", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsTotInvVal")), 2))
                            objInvDetails.ValDtls.Add("RndOffAmt", Math.Round(clsCommon.myCdbl(dtDetails.Rows(0)("ValDtlsRndOffAmt")), 2))
                            'EwbDtls info

                            If objCommonVar.GenerateEWayBillWithEInvoice = True Then
                                If Not stopEWayBill Then
                                    If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("EwbTransId"))) > 0 Then
                                        objInvDetails.EwbDtls.Add("TransId", clsCommon.myCstr(dtDetails.Rows(0)("EwbTransId")))
                                        objInvDetails.EwbDtls.Add("TransName", clsCommon.myCstr(dtDetails.Rows(0)("EwbTransName")))
                                    End If
                                    'objInvDetails.EwbDtls.Add("Distance", Convert.ToInt32(dtDetails.Rows(0)("EwbDistance")))
                                    objInvDetails.EwbDtls.Add("Distance", 0)
                                    If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("EwbVehNo"))) > 0 Then
                                        objInvDetails.EwbDtls.Add("VehNo", clsCommon.myCstr(dtDetails.Rows(0)("EwbVehNo")))
                                    End If
                                    objInvDetails.EwbDtls.Add("VehType", "R")
                                    objInvDetails.EwbDtls.Add("TransMode", "1")
                                End If
                            End If


                            strInvoiceDetails = JsonConvert.SerializeObject(objInvDetails)
                            If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                                strInvoiceDetails = "[{" + """transaction"":" + strInvoiceDetails + "}]"
                            End If
                            streamWriter.Write(strInvoiceDetails)
                        End If
                    End Using
                    Try
                        Dim coll As New Hashtable()
                        clsCommon.AddColumnsForChange(coll, "Document_Code", clsCommon.myCstr(dtDetails.Rows(0)("DocNo")))
                        clsCommon.AddColumnsForChange(coll, "JSON", strInvoiceDetails)
                        clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                        clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_INVOICE_JSON", OMInsertOrUpdate.Insert, "", trans)
                    Catch ex As Exception
                    End Try
                    ''------
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult)
                        If TypeOf jObj.SelectToken("data") Is JObject Then
                            objResult = jObj.SelectToken("data")
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function

    Public Shared Function PostAuthTokenNo_EWAYBillOnly(ByVal strCompCode As String, ByVal strTokenNo As String, ByVal strQry As String, ByVal strLocation As String, ByVal trans As SqlTransaction, ByVal strIRNNo As String) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateAuthToken_IRN' and VendorName='MASTERGST' and Location_Code='" & strLocation & "'"
                Else
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='GenerateAuthToken_IRN' and VendorName='CLEARTAX' and Location_Code='" & strLocation & "'"
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    qry = clsCommon.myCstr(dt.Rows(0)("Url"))
                    qry = qry.Replace("GENERATE", "GENERATE_EWAYBILL")
                    httpRequest = CType(WebRequest.Create(qry), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("username", clsCommon.myCstr(dt.Rows(0)("username")))
                        httpRequest.Headers.Add("auth-token", strTokenNo)
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    Else
                        httpRequest.Headers.Add("x-cleartax-auth-token", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("owner_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "PUT"
                    End If
                    httpRequest.KeepAlive = True
                    'Dim jObj As JObject
                    Dim strInvoiceDetails As String = String.Empty
                    Dim dtDetails As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        If dtDetails IsNot Nothing AndAlso dtDetails.Rows.Count > 0 Then
                            Dim objInvDetails As New clsWayBillOnly
                            objInvDetails.Irn = strIRNNo
                            objInvDetails.Distance = 0 ''Convert.ToInt32(dtDetails.Rows(0)("EwbDistance"))
                            objInvDetails.TransMode = 1
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("EwbTransId"))) > 0 Then
                                objInvDetails.TransId = clsCommon.myCstr(dtDetails.Rows(0)("EwbTransId"))
                                objInvDetails.TransName = clsCommon.myCstr(dtDetails.Rows(0)("EwbTransName"))
                            End If
                            If clsCommon.myLen(clsCommon.myCstr(dtDetails.Rows(0)("EwbVehNo"))) > 0 Then
                                objInvDetails.VehNo = clsCommon.myCstr(dtDetails.Rows(0)("EwbVehNo"))
                            End If
                            objInvDetails.VehType = "R"
                            strInvoiceDetails = JsonConvert.SerializeObject(objInvDetails)
                            If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                                strInvoiceDetails = "[{" + """transaction"":" + strInvoiceDetails + "}]"
                            End If
                            streamWriter.Write(strInvoiceDetails)
                        End If
                    End Using
                    Try
                        Dim coll As New Hashtable()
                        clsCommon.AddColumnsForChange(coll, "Document_Code", clsCommon.myCstr(dtDetails.Rows(0)("DocNo")))
                        clsCommon.AddColumnsForChange(coll, "JSON", strInvoiceDetails)
                        clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                        clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm:ss tt"))
                        clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_INVOICE_JSON", OMInsertOrUpdate.Insert, "", trans)
                    Catch ex As Exception
                    End Try
                    ''------
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        Dim jSONObj As JArray = JArray.Parse(strResult)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult)
                        If TypeOf jObj.SelectToken("data") Is JObject Then
                            objResult = jObj.SelectToken("data")
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            'Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function CancelIRN_withInvoiceData(ByVal strCompCode As String, ByVal strTokenNo As String, ByVal strIRN As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='CancelIRN' and VendorName='MASTERGST' and Location_Code='" & strLocation & "' "
                Else
                    qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='CancelIRN' and VendorName='CLEARTAX' and Location_Code='" & strLocation & "' "
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("username", clsCommon.myCstr(dt.Rows(0)("username")))
                        httpRequest.Headers.Add("auth-token", strTokenNo)
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    Else
                        httpRequest.Headers.Add("x-cleartax-auth-token", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("owner_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "PUT"
                    End If
                    httpRequest.KeepAlive = True
                    'Dim jObj As JObject
                    Dim strInvoiceDetails As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        If clsCommon.myLen(clsCommon.myCstr(strIRN)) > 0 Then
                            If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                                Dim objInvDetails As New clsEInvoiceCancelDetails_ClearTax()
                                objInvDetails.irn = strIRN
                                objInvDetails.CnlRsn = "1"
                                objInvDetails.CnlRem = "Wrong Entry"
                                strInvoiceDetails = JsonConvert.SerializeObject(objInvDetails)
                                strInvoiceDetails = "[" + strInvoiceDetails + "]"
                            Else
                                Dim objInvDetails As New clsEInvoiceCancelDetails()
                                objInvDetails.Irn = strIRN
                                objInvDetails.CnlRsn = "1"
                                objInvDetails.CnlRem = "Wrong Entry"
                                strInvoiceDetails = JsonConvert.SerializeObject(objInvDetails)
                            End If
                            streamWriter.Write(strInvoiceDetails)
                        Else
                            Throw New Exception("Please send IRN number")
                        End If
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        Dim jSONObj As JArray = JArray.Parse(strResult)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult)
                        If TypeOf jObj.SelectToken("data") Is JObject Then
                            objResult = jObj.SelectToken("data")
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function EInvoice_Cancellation(ByVal strDocNo As String, ByVal strIRNNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If
            Dim strtoken As String = ClsEInvoiceOFAPIs.IsGenerateAuthTokenNo_Required(objCommonVar.CurrentCompanyCode, strLocation, trans)
            If clsCommon.myLen(strtoken) > 0 Then
                Dim objResult As Object = ClsEInvoiceOFAPIs.CancelIRN_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strIRNNo, strLocation, trans)
                If objResult IsNot Nothing Then
                    Return True
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

    Public Shared Function GetGSTINDetails(ByVal strCompCode As String, ByVal strGSTNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='getGSTINDetail' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    Dim Url As String = clsCommon.myCstr(dt.Rows(0)("Url")) + "&GSTIN=" + strGSTNo
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(Url)), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "GET"
                    End If
                    httpRequest.KeepAlive = True
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult1 As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult1 = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult1)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult1)
                        If TypeOf jObj.SelectToken("") Is JObject Then
                            objResult = jObj.SelectToken("")
                            Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                            If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                                Dim statusDec As String = objResult.SelectToken("error").ToString
                                objResult = Nothing
                                Throw New Exception(statusDec)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function GetEwayBillDetail(ByVal strCompCode As String, ByVal strewbno As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='getewaybill' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    Dim Url As String = clsCommon.myCstr(dt.Rows(0)("Url")) + "&ewbNo=" + strewbno
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(Url)), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "GET"
                    End If
                    httpRequest.KeepAlive = True
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult1 As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult1 = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult1)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult1)
                        If TypeOf jObj.SelectToken("") Is JObject Then
                            objResult = jObj.SelectToken("")
                            Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                            If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                                Dim statusDec As String = objResult.SelectToken("error").ToString
                                objResult = Nothing
                                Throw New Exception(statusDec)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function GetEwayBillByDate(ByVal strCompCode As String, ByVal strewbDate As DateTime, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim originalDate As DateTime = clsCommon.GetPrintDate(strewbDate, "dd/MM/yyyy")
            Dim formattedDate As String = originalDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
            Dim urlEncodedDate As String = HttpUtility.UrlEncode(formattedDate)
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='getewaybillsbydate' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    Dim Url As String = clsCommon.myCstr(dt.Rows(0)("Url")) + "&date=" + urlEncodedDate
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(Url)), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "GET"
                    End If
                    httpRequest.KeepAlive = True
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult1 As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult1 = streamReader.ReadToEnd()
                    End Using
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "CLEARTAX") = CompairStringResult.Equal Then
                        ''
                        Dim jSONObj As JArray = JArray.Parse(strResult1)
                        Dim jObj As JObject = jSONObj(0)
                        If TypeOf jObj.SelectToken("govt_response") Is JObject Then
                            objResult = jObj.SelectToken("govt_response")
                            Dim status As String = objResult.SelectToken("Success").ToString
                            If clsCommon.CompairString(status, "N") = CompairStringResult.Equal Then
                                Dim strStatusDesc As String = (objResult.SelectToken("ErrorDetails")).ToString()
                                objResult = Nothing
                                Throw New Exception(strStatusDesc)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    Else
                        Dim jObj As JObject = JObject.Parse(strResult1)
                        If TypeOf jObj.SelectToken("") Is JObject Then
                            objResult = jObj.SelectToken("")
                            Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                            If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                                Dim statusDec As String = objResult.SelectToken("error").ToString
                                objResult = Nothing
                                Throw New Exception(statusDec)
                            End If
                            Return objResult
                        Else
                            Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                            objResult = Nothing
                            Throw New Exception(strStatusDesc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function CancelEWayBill(ByVal strCompCode As String, ByVal strEwayBillNo As String, ByVal strRemark As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='Cancelewaybill' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    Dim strEwbCancelDetails As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        Dim objEWayBill As Object = New clsEwayBillCancel()
                        objEWayBill.ewbno = clsCommon.myCdbl(strEwayBillNo)
                        objEWayBill.cancelRsnCode = clsCommon.myCdbl(2)
                        objEWayBill.cancelRmrk = clsCommon.myCstr(strRemark)
                        strEwbCancelDetails = JsonConvert.SerializeObject(objEWayBill)
                        streamWriter.Write(strEwbCancelDetails)
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    Dim jObj As JObject = JObject.Parse(strResult)
                    If TypeOf jObj.SelectToken("") Is JObject Then
                        objResult = jObj.SelectToken("")
                        Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                        If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                            Dim statusDec As String = objResult.SelectToken("error").ToString
                            objResult = Nothing
                            Throw New Exception(statusDec)
                        End If
                        Return objResult
                    Else
                        Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                        objResult = Nothing
                        Throw New Exception(strStatusDesc)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function EWayBill_Update_Transporter(ByVal strCompCode As String, ByVal strEwayBillNo As String, ByVal strTranspoterID As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='updatetransporter' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    Dim strEwbUpdateTranspoter As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        Dim objUpdateTranspoter As Object = New clsEwayBillUpdateTranspoter()
                        objUpdateTranspoter.ewbno = clsCommon.myCdbl(strEwayBillNo)
                        objUpdateTranspoter.transporterId = clsCommon.myCstr(strTranspoterID)
                        strEwbUpdateTranspoter = JsonConvert.SerializeObject(objUpdateTranspoter)
                        streamWriter.Write(strEwbUpdateTranspoter)
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    Dim jObj As JObject = JObject.Parse(strResult)
                    If TypeOf jObj.SelectToken("") Is JObject Then
                        objResult = jObj.SelectToken("")
                        Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                        If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                            Dim statusDec As String = objResult.SelectToken("error").ToString
                            objResult = Nothing
                            Throw New Exception(statusDec)
                        End If
                        Return objResult
                    Else
                        Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                        objResult = Nothing
                        Throw New Exception(strStatusDesc)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function EWayBill_Update_PartB(ByVal strCompCode As String, ByVal obj As clsEwayBillUpdatePartB_Vehicle, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='vehewb' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    Dim strUpdateVehicle As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        Dim objUpdateVehicle As Object = New clsEwayBillUpdatePartB_Vehicle()
                        objUpdateVehicle.ewbno = clsCommon.myCdbl(obj.ewbNo)
                        objUpdateVehicle.vehicleNo = clsCommon.myCstr(obj.vehicleNo)
                        objUpdateVehicle.fromPlace = clsCommon.myCstr(obj.fromPlace)
                        objUpdateVehicle.fromState = clsCommon.myCdbl(obj.fromState)
                        objUpdateVehicle.reasonCode = clsCommon.myCstr(obj.reasonCode)
                        objUpdateVehicle.reasonRem = clsCommon.myCstr(obj.reasonRem)
                        objUpdateVehicle.transDocNo = clsCommon.myCstr(obj.transDocNo)
                        objUpdateVehicle.transDocDate = clsCommon.myCstr(obj.transDocDate)
                        objUpdateVehicle.transMode = clsCommon.myCstr(obj.transMode)
                        strUpdateVehicle = JsonConvert.SerializeObject(objUpdateVehicle)
                        streamWriter.Write(strUpdateVehicle)
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    Dim jObj As JObject = JObject.Parse(strResult)
                    If TypeOf jObj.SelectToken("") Is JObject Then
                        objResult = jObj.SelectToken("")
                        Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                        If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                            Dim statusDec As String = objResult.SelectToken("error").ToString
                            objResult = Nothing
                            Throw New Exception(statusDec)
                        End If
                        Return objResult
                    Else
                        Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                        objResult = Nothing
                        Throw New Exception(strStatusDesc)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function

    Public Shared Function EWayBill_ExtendValidity(ByVal strCompCode As String, ByVal obj As clsextendvalidity, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='extendvalidity' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    Dim strextendvalidity As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        Dim objextendvalidity As Object = New clsextendvalidity()
                        objextendvalidity.ewbno = clsCommon.myCdbl(obj.ewbNo)
                        objextendvalidity.vehicleNo = clsCommon.myCstr(obj.vehicleNo)
                        objextendvalidity.fromPlace = clsCommon.myCstr(obj.fromPlace)
                        objextendvalidity.fromState = clsCommon.myCdbl(obj.fromState)
                        objextendvalidity.remainingDistance = clsCommon.myCdbl(obj.remainingDistance)
                        objextendvalidity.transDocNo = clsCommon.myCstr(obj.transDocNo)
                        objextendvalidity.transDocDate = clsCommon.myCstr(obj.transDocDate)
                        objextendvalidity.transMode = clsCommon.myCstr(obj.transMode)
                        objextendvalidity.extnRsnCode = clsCommon.myCstr(obj.extnRsnCode)
                        objextendvalidity.extnRemarks = clsCommon.myCstr(obj.extnRemarks)
                        objextendvalidity.fromPincode = clsCommon.myCstr(obj.fromPincode)
                        objextendvalidity.consignmentStatus = clsCommon.myCstr(obj.consignmentStatus)
                        objextendvalidity.transitType = clsCommon.myCstr(obj.transitType)
                        objextendvalidity.addressLine1 = clsCommon.myCstr(obj.addressLine1)
                        objextendvalidity.addressLine2 = clsCommon.myCstr(obj.addressLine2)
                        objextendvalidity.addressLine3 = clsCommon.myCstr(obj.addressLine3)
                        strextendvalidity = JsonConvert.SerializeObject(objextendvalidity)
                        streamWriter.Write(strextendvalidity)
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    Dim jObj As JObject = JObject.Parse(strResult)
                    If TypeOf jObj.SelectToken("") Is JObject Then
                        objResult = jObj.SelectToken("")
                        Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                        If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                            Dim statusDec As String = objResult.SelectToken("error").ToString
                            objResult = Nothing
                            Throw New Exception(statusDec)
                        End If
                        Return objResult
                    Else
                        Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                        objResult = Nothing
                        Throw New Exception(strStatusDesc)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
    Public Shared Function EWayBill_Close(ByVal strCompCode As String, ByVal obj As clsEwayBillClose, ByVal strLocation As String, ByVal trans As SqlTransaction) As Object
        Dim httpRequest As HttpWebRequest = Nothing
        Dim httpResponse As HttpWebResponse = Nothing
        Dim objResult As Object = Nothing
        Try
            Dim qry As String = String.Empty
            Dim strEInvoiceVendor As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EInvoiceVendor, clsFixedParameterCode.EInvoiceVendor, trans))
            If clsCommon.myLen(clsCommon.myCstr(strEInvoiceVendor)) > 0 Then
                If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                    Dim IsAuthenticate As Boolean = IsEWayBillAPIAuthenticate(strCompCode, strLocation, trans)
                    If IsAuthenticate Then
                        qry = "Select * from TSPL_EInvoiceHeader_Info where Comp_Code='" & strCompCode & "' and RequiredFor ='clsewb' and Location_Code='" & strLocation & "'"
                    End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Const Tls12Value As Integer = &HC00
                    Dim Tls12Protocol As System.Net.SecurityProtocolType = CType(Tls12Value, System.Net.SecurityProtocolType)
                    ' Set the security protocol globally for all HTTPS requests
                    System.Net.ServicePointManager.SecurityProtocol = Tls12Protocol
                    httpRequest = CType(WebRequest.Create(clsCommon.myCstr(dt.Rows(0)("Url"))), HttpWebRequest)
                    httpRequest.ContentType = "application/json"
                    If clsCommon.CompairString(strEInvoiceVendor.ToUpper(), "MASTERGST") = CompairStringResult.Equal Then
                        httpRequest.Headers.Add("ip_address", clsCommon.myCstr(dt.Rows(0)("ip_address")))
                        httpRequest.Headers.Add("client_id", clsCommon.myCstr(dt.Rows(0)("client_id")))
                        httpRequest.Headers.Add("client_secret", clsCommon.myCstr(dt.Rows(0)("client_secret")))
                        httpRequest.Headers.Add("gstin", clsCommon.myCstr(dt.Rows(0)("gstin")))
                        httpRequest.Method = "POST"
                    End If
                    httpRequest.KeepAlive = True
                    Dim strCloseEWB As String = String.Empty
                    Using streamWriter = New StreamWriter(httpRequest.GetRequestStream())
                        Dim objCloseEWB As Object = New clsEwayBillClose()
                        objCloseEWB.ewbNo = clsCommon.myCdbl(obj.ewbNo)
                        objCloseEWB.closureDate = clsCommon.myCstr(obj.closureDate)
                        objCloseEWB.remarks = clsCommon.myCstr(obj.remarks)

                        strCloseEWB = JsonConvert.SerializeObject(objCloseEWB)
                        streamWriter.Write(strCloseEWB)
                    End Using
                    httpResponse = CType(httpRequest.GetResponse(), HttpWebResponse)
                    '' required validation httresponse
                    Dim strResult As String = String.Empty
                    Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                        strResult = streamReader.ReadToEnd()
                    End Using
                    Dim jObj As JObject = JObject.Parse(strResult)
                    If TypeOf jObj.SelectToken("") Is JObject Then
                        objResult = jObj.SelectToken("")
                        Dim Status_cd As String = objResult.SelectToken("status_cd").ToString
                        If clsCommon.CompairString(Status_cd, "0") = CompairStringResult.Equal Then
                            Dim statusDec As String = objResult.SelectToken("error").ToString
                            objResult = Nothing
                            Throw New Exception(statusDec)
                        End If
                        Return objResult
                    Else
                        Dim strStatusDesc As String = (jObj.SelectToken("status_desc")).ToString()
                        objResult = Nothing
                        Throw New Exception(strStatusDesc)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return objResult
    End Function
End Class
Public Class clsEInvoiceCancelDetails
    Public Property Irn As String = String.Empty
    Public Property CnlRsn As String = String.Empty
    Public Property CnlRem As String = String.Empty
End Class
Public Class clsEInvoiceCancelDetails_ClearTax
    Public Property irn As String = String.Empty
    Public Property CnlRsn As String = String.Empty
    Public Property CnlRem As String = String.Empty
End Class
Public Class clsEInvoiceDetails
    Public Property Version As String = String.Empty
    Public Property TranDtls As New Dictionary(Of String, String)
    Public Property DocDtls As New Dictionary(Of String, String)
    Public Property SellerDtls As New Dictionary(Of String, Object)
    Public Property BuyerDtls As New Dictionary(Of String, Object)
    Public Property ItemList As List(Of ClsItemList) = Nothing
    Public Property ValDtls As New Dictionary(Of String, Double)
End Class
Public Class clsEInvoiceDetailsWithEWayBill
    Public Property Version As String = String.Empty
    Public Property TranDtls As New Dictionary(Of String, String)
    Public Property DocDtls As New Dictionary(Of String, String)
    Public Property SellerDtls As New Dictionary(Of String, Object)
    Public Property BuyerDtls As New Dictionary(Of String, Object)
    Public Property ItemList As List(Of ClsItemList) = Nothing
    Public Property ValDtls As New Dictionary(Of String, Double)
    Public Property EwbDtls As New Dictionary(Of String, Object)
End Class
Public Class ClsItemList
    Public Property SlNo As String = String.Empty
    Public Property IsServc As String = String.Empty
    Public Property PrdDesc As String = String.Empty
    Public Property HsnCd As String = String.Empty
    Public Property Qty As Double = 0
    Public Property Unit As String = String.Empty
    Public Property UnitPrice As Double = 0
    Public Property TotAmt As Double = 0
    Public Property Discount As Double = 0
    Public Property AssAmt As Double = 0
    Public Property GstRt As Double = 0
    Public Property SgstAmt As Double = 0
    Public Property IgstAmt As Double = 0
    Public Property CgstAmt As Double = 0
    Public Property OthChrg As Double = 0
    Public Property TotItemVal As Double = 0
End Class
Public Class clsWayBillOnly
    Public Property Irn As String
    Public Property Distance As Integer
    Public Property TransMode As String
    Public Property TransId As String
    Public Property TransName As String
    Public Property TransDocDt As String
    Public Property TransDocNo As String
    Public Property VehNo As String
    Public Property VehType As String
End Class
Public Class clsEwayBillDetail
    Public Property supplyType As String
    Public Property subSupplyType As String
    Public Property subSupplyDesc As String
    Public Property docType As String
    Public Property docNo As String
    Public Property docDate As String
    Public Property fromGstin As String
    Public Property fromTrdName As String
    Public Property fromAddr1 As String
    Public Property fromAddr2 As String
    Public Property fromPlace As String
    Public Property actFromStateCode As Integer
    Public Property fromPincode As Integer
    Public Property fromStateCode As Integer
    Public Property toGstin As String
    Public Property toTrdName As String
    Public Property toAddr1 As String
    Public Property toAddr2 As String
    Public Property toPlace As String
    Public Property toPincode As Integer
    Public Property actToStateCode As Integer
    Public Property toStateCode As Integer
    Public Property transactionType As Integer
    'Public Property dispatchFromGSTIN As String
    Public Property dispatchFromTradeName As String
    'Public Property shipToGSTIN As String
    Public Property shipToTradeName As String
    Public Property totalValue As Double
    Public Property cgstValue As Double
    Public Property sgstValue As Double
    Public Property igstValue As Double
    Public Property cessValue As Double
    Public Property cessNonAdvolValue As Double
    Public Property totInvValue As Double
    Public Property transMode As String
    Public Property transDistance As String
    Public Property transporterName As String
    Public Property transporterId As String
    Public Property transDocNo As String
    Public Property transDocDate As String
    Public Property vehicleNo As String
    Public Property vehicleType As String
    Public Property itemList As List(Of clsItem)
End Class
Public Class clsEwayBillDetailWithoutTrans
    Public Property supplyType As String
    Public Property subSupplyType As String
    Public Property subSupplyDesc As String
    Public Property docType As String
    Public Property docNo As String
    Public Property docDate As String
    Public Property fromGstin As String
    Public Property fromTrdName As String
    Public Property fromAddr1 As String
    Public Property fromAddr2 As String
    Public Property fromPlace As String
    Public Property actFromStateCode As Integer
    Public Property fromPincode As Integer
    Public Property fromStateCode As Integer
    Public Property toGstin As String
    Public Property toTrdName As String
    Public Property toAddr1 As String
    Public Property toAddr2 As String
    Public Property toPlace As String
    Public Property toPincode As Integer
    Public Property actToStateCode As Integer
    Public Property toStateCode As Integer
    Public Property transactionType As Integer
    'Public Property dispatchFromGSTIN As String
    Public Property dispatchFromTradeName As String
    'Public Property shipToGSTIN As String
    Public Property shipToTradeName As String
    Public Property totalValue As Double
    Public Property cgstValue As Double
    Public Property sgstValue As Double
    Public Property igstValue As Double
    Public Property cessValue As Double
    Public Property cessNonAdvolValue As Double
    Public Property totInvValue As Double
    Public Property transMode As String
    Public Property transDistance As String
    'Public Property transporterName As String
    'Public Property transporterId As String
    'Public Property transDocNo As String
    Public Property transDocDate As String
    Public Property vehicleNo As String
    Public Property vehicleType As String
    Public Property itemList As List(Of clsItem)
End Class
Public Class clsEwayBillCancel
    Public Property ewbNo As Long
    Public Property cancelRsnCode As Integer
    Public Property cancelRmrk As String

End Class
Public Class clsEwayBillClose
    Public Property ewbNo As Long
    Public Property closureDate As DateTime
    Public Property remarks As String

End Class
Public Class clsEwayBillUpdateTranspoter
    Public Property ewbNo As Long
    Public Property transporterId As String

End Class
Public Class clsEwayBillUpdatePartB_Vehicle
    Public Property ewbNo As Long
    Public Property vehicleNo As String
    Public Property fromPlace As String
    Public Property fromState As Integer
    Public Property reasonCode As String
    Public Property reasonRem As String
    Public Property transDocNo As String
    Public Property transDocDate As String
    Public Property transMode As String

End Class
Public Class clsItem
    Public Property productName As String
    Public Property productDesc As String
    Public Property hsnCode As Integer
    Public Property quantity As Integer
    Public Property qtyUnit As String
    Public Property taxableAmount As Double
    Public Property sgstRate As Double
    Public Property cgstRate As Double
    Public Property igstRate As Double
    Public Property cessRate As Double
End Class
Public Class clsextendvalidity
    Public Property ewbNo As Long
    Public Property vehicleNo As String
    Public Property fromPlace As String
    Public Property fromState As Integer
    Public Property remainingDistance As Integer
    Public Property transDocNo As String
    Public Property transDocDate As String
    Public Property transMode As String
    Public Property extnRsnCode As Integer
    Public Property extnRemarks As String
    Public Property fromPincode As Integer
    Public Property consignmentStatus As String
    Public Property transitType As String
    Public Property addressLine1 As String
    Public Property addressLine2 As String
    Public Property addressLine3 As String
End Class
Public Class clsUpdateTranspoter
    Public Property ewbNo As Long
    Public Property transporterId As String
End Class

