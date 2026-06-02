Imports common
Imports System.Data.SqlClient
Imports Telerik.WinControls.UI
Imports Telerik.WinControls
Imports System.Data
Imports System.Windows.Forms
Imports System.Configuration
Imports Telerik.Collections.Generic
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization
Public Class clsTaxGroupMaster
    Public Tax_Group_Code As String = Nothing
    Public Tax_Group_Desc As String = Nothing
    Public Tax_Group_Type As String = Nothing
    '' multicurrency columns
    Public CURRENCY_CODE As String

    Public Arr As List(Of clsTaxGroupDetail) = Nothing
    Public Function GetDataForPurchase(ByVal strCode As String) As clsTaxGroupMaster
        Return GetData(strCode, "P")
    End Function
    Public Function GetDataForSale(ByVal strCode As String) As clsTaxGroupMaster
        Return GetData(strCode, "S")
    End Function

    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " select Tax_Group_Code as [Code],Tax_Group_Desc as [Description],Tax_Group_Type as [Tax Group Type],CURRENCY_CODE as [Currency Code],Created_By as [Created By],Created_Date as [Created Date],Modify_By as [Modify By],Modify_Date as [Modify Date],Comp_Code as [Company Code],Excisable as [Excisable],VAT,STax ,Tax_Group_Code_InterState as [Tax Group Code Inter State],Tax_Group_Description_InterState as [Tax Group Description Inter State],Is_Transfer as [Is Transfer] from TSPL_TAX_GROUP_MASTER "
        str = clsCommon.ShowSelectForm("RPTTXGRPFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    '----------------End of Code For Get Finder--------------------------------------------------------------'


    Public Function GetData(ByVal strCode As String, ByVal strType As String) As clsTaxGroupMaster
        Dim obj As clsTaxGroupMaster = Nothing
        Dim qry As String = "select Tax_Group_Code,Tax_Group_Desc,Tax_Group_Type,CURRENCY_CODE from TSPL_TAX_GROUP_MASTER Where Tax_Group_Code='" + strCode + "' and Tax_Group_Type='" + strType + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsTaxGroupMaster()
            obj.Tax_Group_Code = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Code"))
            obj.Tax_Group_Desc = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            obj.Tax_Group_Type = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Type"))
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))

            qry = "select  TSPL_TAX_GROUP_DETAILS.Trans_Code,TSPL_TAX_GROUP_DETAILS.Tax_Group_Code,TSPL_TAX_GROUP_DETAILS.Tax_Group_Type,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,TSPL_TAX_GROUP_DETAILS.Taxable,TSPL_TAX_GROUP_DETAILS.Surtax,TSPL_TAX_GROUP_DETAILS.Surtax_Tax_Code,TSPL_TAX_GROUP_DETAILS.Surtax_Tax_Code_Desc,TSPL_TAX_MASTER.Excisable from TSPL_TAX_GROUP_DETAILS   left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where Tax_Group_Code='" + strCode + "'  and Tax_Group_Type='" + strType + "' order by Trans_Code"
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
                    objTR.Excisable = clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal
                    objTR.Surtax_Tax_Code = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                    objTR.Surtax_Tax_Code_Desc = clsCommon.myCstr(dr("Surtax_Tax_Code_Desc"))
                    obj.Arr.Add(objTR)
                Next
            End If
        End If
        Return obj
    End Function
    Public Shared Function GetTaxGroupFromTaxGroupMaster(ByVal strTransLocation As String, ByVal strVendorCustomerCode As String, ByVal strTaxType As String, ByVal isTaxable As Boolean, ByVal tranDate As Date?) As String
        Return GetTaxGroupFromTaxGroupMaster(strTransLocation, strVendorCustomerCode, strTaxType, isTaxable, tranDate, Nothing)
    End Function
    Public Shared Function GetTaxGroupFromTaxGroupMaster(ByVal strTransLocation As String, ByVal strVendorCustomerCode As String, ByVal strTaxType As String, ByVal isTaxable As Boolean, ByVal tranDate As Date?, ByVal trans As SqlTransaction) As String
        Dim qry As String = "  "
        Dim TaxQry = " SELECT Tax_Group_Code From TSPL_TAX_GROUP_MASTER where 2=2 "
        Dim GSTStatus As Boolean = clsERPFuncationality.GetGSTStatus(tranDate)
        If GSTStatus = False OrElse (isTaxable AndAlso GSTStatus) Then
            qry += " select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strTransLocation + "' union all   "
            If clsCommon.CompairString("S", strTaxType) = CompairStringResult.Equal Then
                qry += "  select   State from TSPL_CUSTOMER_MASTER where Cust_Code='" + strVendorCustomerCode + "' "
            ElseIf clsCommon.CompairString("P", strTaxType) = CompairStringResult.Equal Then
                qry += "   select  State_Code as State from TSPL_VENDOR_MASTER where Vendor_Code='" + strVendorCustomerCode + "' "
            Else
                Throw New Exception("Please enter valid Tax Type it should be 'P' or 'S'")
            End If
            qry += " )x "
            Dim State As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.CompairString(State, "L") = CompairStringResult.Equal Then
                TaxQry += " and Default_Type=3 "
            ElseIf clsCommon.CompairString(State, "I") = CompairStringResult.Equal Then
                TaxQry += " and Default_Type=2 "
            End If
        Else
            TaxQry += " and Default_Type=1 "
        End If

        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(TaxQry, trans))

    End Function

    Public Shared Function GetTaxDetailsByLocation(ByVal GrpCode As String, ByVal strTaxType As String, ByVal strVendorCustomerCode As String, ByVal strLocation As String) As DataTable
        Dim openTaxcond As String = ""
        Dim whrCls As String = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type='" + strTaxType + "' "
        Dim whrCls_taxgrp As String = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" + strTaxType + "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" + strTaxType + "' "
        'If Without_State_Condition Then
        '    openTaxcond = " top 1 "
        '    ''as discussed with ranjana mam only exact type used 
        '    If clsCommon.CompairString(strTaxType, "S") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxType, "T") = CompairStringResult.Equal Then
        '        whrCls = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type in ('" + strTaxType + "') "
        '        whrCls_taxgrp = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type in ('" + strTaxType + "') and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type in ('" + strTaxType + "') "
        '    End If
        'End If

        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,isnull(("
        qry += "select " + openTaxcond + " Tax_Rate from TSPL_LOCATION_WISE_TAX_MASTER WHERE TSPL_LOCATION_WISE_TAX_MASTER.Is_Default_Tax=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code " + whrCls + " and TSPL_LOCATION_WISE_TAX_MASTER.Location_Code='" + strLocation + "' "

        ' If Not Without_State_Condition Then
        qry += " and Tax_Category in (select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocation + "' union all   "

        If clsCommon.CompairString("S", strTaxType) = CompairStringResult.Equal Then
            qry += "  select   State from TSPL_CUSTOMER_MASTER where Cust_Code='" + strVendorCustomerCode + "' "
        ElseIf clsCommon.CompairString("P", strTaxType) = CompairStringResult.Equal Then
            qry += "   select  State_Code as State from TSPL_VENDOR_MASTER where Vendor_Code='" + strVendorCustomerCode + "' "
        Else
            Throw New Exception("Please enter valid Tax Type it should be 'P' or 'S'")
        End If
        qry += ")x)"
        ' End If


        qry += "),0) AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.IS_TCS from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code  where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' " + whrCls_taxgrp + " order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetailsByLocation(ByVal GrpCode As String, ByVal strTaxType As String, ByVal strVendorCustomerCode As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As DataTable
        Dim openTaxcond As String = ""
        Dim whrCls As String = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type='" + strTaxType + "' "
        Dim whrCls_taxgrp As String = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" + strTaxType + "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" + strTaxType + "' "
        'If Without_State_Condition Then
        '    openTaxcond = " top 1 "
        '    ''as discussed with ranjana mam only exact type used 
        '    If clsCommon.CompairString(strTaxType, "S") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxType, "T") = CompairStringResult.Equal Then
        '        whrCls = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type in ('" + strTaxType + "') "
        '        whrCls_taxgrp = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type in ('" + strTaxType + "') and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type in ('" + strTaxType + "') "
        '    End If
        'End If

        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,isnull(("
        qry += "select " + openTaxcond + " Tax_Rate from TSPL_LOCATION_WISE_TAX_MASTER WHERE TSPL_LOCATION_WISE_TAX_MASTER.Is_Default_Tax=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code " + whrCls + " and TSPL_LOCATION_WISE_TAX_MASTER.Location_Code='" + strLocation + "' "

        ' If Not Without_State_Condition Then
        qry += " and Tax_Category in (select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocation + "' union all   "

        If clsCommon.CompairString("S", strTaxType) = CompairStringResult.Equal Then
            qry += "  select   State from TSPL_CUSTOMER_MASTER where Cust_Code='" + strVendorCustomerCode + "' "
        ElseIf clsCommon.CompairString("P", strTaxType) = CompairStringResult.Equal Then
            qry += "   select  State_Code as State from TSPL_VENDOR_MASTER where Vendor_Code='" + strVendorCustomerCode + "' "
        Else
            Throw New Exception("Please enter valid Tax Type it should be 'P' or 'S'")
        End If
        qry += ")x)"
        ' End If


        qry += "),0) AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.IS_TCS from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code  where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' " + whrCls_taxgrp + " order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function
    Public Shared Function GetTaxDetailsByLocation(ByVal GrpCode As String, ByVal strTaxType As String, ByVal strVendorCustomerCode As String, ByVal strLocation As String, ByVal Without_State_Condition As Boolean) As DataTable
        Dim openTaxcond As String = ""
        Dim whrCls As String = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type='" + strTaxType + "' "
        Dim whrCls_taxgrp As String = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" + strTaxType + "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" + strTaxType + "' "
        If Without_State_Condition Then
            openTaxcond = " top 1 "
            ''as discussed with ranjana mam only exact type used 
            If clsCommon.CompairString(strTaxType, "S") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxType, "T") = CompairStringResult.Equal Then
                whrCls = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type in ('" + strTaxType + "') "
                whrCls_taxgrp = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type in ('" + strTaxType + "') and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type in ('" + strTaxType + "') "
            End If
        End If

        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,isnull(("
        qry += "select " + openTaxcond + " Tax_Rate from TSPL_LOCATION_WISE_TAX_MASTER WHERE TSPL_LOCATION_WISE_TAX_MASTER.Is_Default_Tax=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code " + whrCls + " and TSPL_LOCATION_WISE_TAX_MASTER.Location_Code='" + strLocation + "' "

        If Not Without_State_Condition Then
            qry += " and Tax_Category in (select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocation + "' union all   "

            If clsCommon.CompairString("S", strTaxType) = CompairStringResult.Equal Then
                qry += "  select   State from TSPL_CUSTOMER_MASTER where Cust_Code='" + strVendorCustomerCode + "' "
            ElseIf clsCommon.CompairString("P", strTaxType) = CompairStringResult.Equal Then
                qry += "   select  State_Code as State from TSPL_VENDOR_MASTER where Vendor_Code='" + strVendorCustomerCode + "' "
            Else
                Throw New Exception("Please enter valid Tax Type it should be 'P' or 'S'")
            End If
            qry += ")x)"
        End If


        qry += "),0) AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.IS_TCS from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code  where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' " + whrCls_taxgrp + " order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetailsByLocation(ByVal GrpCode As String, ByVal strTaxType As String, ByVal strVendorCustomerCode As String, ByVal strLocation As String, ByVal ItemCode As String, ByVal DocDate As Date) As DataTable
        Dim openTaxcond As String = ""
        Dim strjoin As String = String.Empty
        Dim whrCls As String = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type='" + strTaxType + "' "
        Dim whrCls_taxgrp As String = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" + strTaxType + "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" + strTaxType + "' "
        'If Without_State_Condition Then
        '    openTaxcond = " top 1 "
        '    ''as discussed with ranjana mam only exact type used 
        '    If clsCommon.CompairString(strTaxType, "S") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxType, "T") = CompairStringResult.Equal Then
        '        whrCls = " and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type in ('" + strTaxType + "') "
        '        whrCls_taxgrp = " and TSPL_TAX_GROUP_MASTER.Tax_Group_Type in ('" + strTaxType + "') and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type in ('" + strTaxType + "') "
        '    End If
        'End If

        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount," 'isnull(("
        'qry += "select " + openTaxcond + " Tax_Rate from TSPL_LOCATION_WISE_TAX_MASTER WHERE TSPL_LOCATION_WISE_TAX_MASTER.Is_Default_Tax=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code " + whrCls + " and TSPL_LOCATION_WISE_TAX_MASTER.Location_Code='" + strLocation + "' "

        'If Not Without_State_Condition Then
        '    qry += " and Tax_Category in (select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocation + "' union all   "

        '    If clsCommon.CompairString("S", strTaxType) = CompairStringResult.Equal Then
        '        qry += "  select   State from TSPL_CUSTOMER_MASTER where Cust_Code='" + strVendorCustomerCode + "' "
        '    ElseIf clsCommon.CompairString("P", strTaxType) = CompairStringResult.Equal Then
        '        qry += "   select  State_Code as State from TSPL_VENDOR_MASTER where Vendor_Code='" + strVendorCustomerCode + "' "
        '    Else
        '        Throw New Exception("Please enter valid Tax Type it should be 'P' or 'S'")
        '    End If
        '    qry += ")x)"
        'End If
        If clsCommon.myLen(ItemCode) > 0 Then
            qry += " X.TAX_Rate as TaxRate, "
            'whrCls_taxgrp += "  and TSPL_ITEM_WISE_TAX_GROUP.Item_Code IN(select top 1 TSPL_ITEM_WISE_TAX_GROUP.Item_Code from TSPL_ITEM_WISE_TAX left join TSPL_ITEM_WISE_TAX_GROUP on TSPL_ITEM_WISE_TAX.HCODE=TSPL_ITEM_WISE_TAX_GROUP.HCODE where TSPL_ITEM_WISE_TAX_GROUP.Item_Code='" + ItemCode + "' order by TSPL_ITEM_WISE_TAX.DOC_DATE desc) "
            strjoin = "  left join ( select top 1 TSPL_ITEM_WISE_TAX_GROUP.Item_Code, TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code, TSPL_ITEM_WISE_TAX_AUTHORITY.TAX_Rate, TSPL_ITEM_WISE_TAX.DOC_DATE from TSPL_ITEM_WISE_TAX left join TSPL_ITEM_WISE_TAX_GROUP on TSPL_ITEM_WISE_TAX.HCODE = TSPL_ITEM_WISE_TAX_GROUP.HCODE left join TSPL_ITEM_WISE_TAX_AUTHORITY on TSPL_ITEM_WISE_TAX_GROUP.DCODE=TSPL_ITEM_WISE_TAX_AUTHORITY.DCODE where TSPL_ITEM_WISE_TAX_GROUP.Item_Code = '" + ItemCode + "' and TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code='" + GrpCode + "' and TSPL_ITEM_WISE_TAX.DOC_DATE<='" + clsCommon.GetPrintDate(DocDate) + "' order by TSPL_ITEM_WISE_TAX.DOC_DATE desc) X on TSPL_TAX_GROUP_DETAILS.Tax_Group_Code=X.Tax_Group_Code "
            '          strjoin = " LEFT JOIN (SELECT TSPL_ITEM_WISE_TAX_GROUP.Item_Code,TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code, TSPL_ITEM_WISE_TAX_AUTHORITY.TAX_Rate, TSPL_ITEM_WISE_TAX_AUTHORITY.Tax_Authority,TSPL_ITEM_WISE_TAX.DOC_DATE 
            '  FROM  TSPL_ITEM_WISE_TAX  LEFT JOIN TSPL_ITEM_WISE_TAX_GROUP ON TSPL_ITEM_WISE_TAX.HCODE = TSPL_ITEM_WISE_TAX_GROUP.HCODE LEFT JOIN TSPL_ITEM_WISE_TAX_AUTHORITY ON TSPL_ITEM_WISE_TAX_GROUP.DCODE = TSPL_ITEM_WISE_TAX_AUTHORITY.DCODE 
            '  WHERE  TSPL_ITEM_WISE_TAX_GROUP.Item_Code = '" + ItemCode + "' AND TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code = '" + GrpCode + "'  AND TSPL_ITEM_WISE_TAX.DOC_DATE = ( SELECT MAX(DOC_DATE) 
            '        FROM TSPL_ITEM_WISE_TAX LEFT JOIN TSPL_ITEM_WISE_TAX_GROUP ON TSPL_ITEM_WISE_TAX.HCODE = TSPL_ITEM_WISE_TAX_GROUP.HCODE  LEFT JOIN TSPL_ITEM_WISE_TAX_AUTHORITY ON TSPL_ITEM_WISE_TAX_GROUP.DCODE = TSPL_ITEM_WISE_TAX_AUTHORITY.DCODE 
            '  WHERE TSPL_ITEM_WISE_TAX_GROUP.Item_Code = '" + ItemCode + "' AND TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code = '" + GrpCode + "'  and Status=1  )
            ') X ON TSPL_TAX_GROUP_DETAILS.Tax_Group_Code = X.Tax_Group_Code  AND TSPL_TAX_GROUP_DETAILS.Tax_Code = X.Tax_Authority "
        Else
            qry += " 0 as TaxRate, "
        End If

        qry += " TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.IS_TCS from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code " + strjoin + "  where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' " + whrCls_taxgrp + " order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetailsByLocationForTransfer(ByVal GrpCode As String, ByVal strTaxType As String, ByVal strLocationToCode As String, ByVal strLocation As String) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,isnull(("
        qry += "select  Tax_Rate from TSPL_LOCATION_WISE_TAX_MASTER WHERE TSPL_LOCATION_WISE_TAX_MASTER.Is_Default_Tax=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code and TSPL_LOCATION_WISE_TAX_MASTER.Tax_Type='" + strTaxType + "' and TSPL_LOCATION_WISE_TAX_MASTER.Location_Code='" + strLocation + "' and Tax_Category in (select case when MIN(x.State)=MAX(x.State) then 'L' else 'I' end  from  (select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocation + "' union all select State   from TSPL_LOCATION_MASTER where Location_Code='" + strLocationToCode + "'  "
        qry += ")x)),0) AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable,TSPL_TAX_MASTER.Excisable,TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.IS_TCS from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" + strTaxType + "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" + strTaxType + "' order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetails(ByVal GrpCode As String) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.Is_TCS  from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetailsOfSale(ByVal GrpCode As String, ByVal strTaxType As String) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,Surtax,Surtax_Tax_Code,isnull((select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P'),0) AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.Is_TCS  from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='" & strTaxType & "' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='" & strTaxType & "' order by Trans_Code"
        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetails_Union(ByVal GrpCode As String, ByVal HCode As String) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate=(select TAX_Rate from TSPL_SAC_WISE_TAX_AUTHORITY where HCODE='" & HCode & "' and TSPL_SAC_WISE_TAX_AUTHORITY.Tax_Authority=TSPL_TAX_GROUP_DETAILS.Tax_Code ) AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable , TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.Is_TCS  from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"

        Return clsDBFuncationality.GetDataTable(qry)
    End Function
    Public Shared Function GetTaxDetailsscrap(ByVal GrpCode As String) As DataTable
        Return (GetTaxDetailsscrap(GrpCode, Nothing))
    End Function

    Public Shared Function GetTaxDetailsscrap(ByVal GrpCode As String, ByVal trans As SqlTransaction) As DataTable
        ''UDL/11/01/19-000253 by balwinder on 11/01/2019
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,TSPL_TAX_GROUP_DETAILS.Tax_Code,TSPL_TAX_GROUP_DETAILS.Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='s') AS TaxRate,TSPL_TAX_GROUP_DETAILS.Taxable, TSPL_TAX_MASTER.Excisable,TSPL_TAX_MASTER.Type,TSPL_TAX_MASTER.Tax_Recoverable,TSPL_TAX_GROUP_DETAILS.Tax_On_Base_Amount from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + GrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='s' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='s' order by Trans_Code"
        Return clsDBFuncationality.GetDataTable(qry, trans)
    End Function

    Public Shared Function GetNameOfSaleType(ByVal GrpCode As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Tax_Group_Desc from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" + GrpCode + "' and Tax_Group_Type='S'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
    End Function

    Public Shared Function GetNameOfPurchaseType(ByVal GrpCode As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Tax_Group_Desc from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" + GrpCode + "' and Tax_Group_Type='P'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
    End Function

    Public Shared Function GetName(ByVal GrpCode As String, ByVal Type As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Tax_Group_Desc from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" + GrpCode + "' and Tax_Group_Type='" + Type + "'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
    End Function

    Public Shared Function IsHavingRecoverableTaxAuthority(ByVal strTaxGroup As String, ByVal strType As String, ByVal trans As SqlTransaction)
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Code from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code  where Tax_Group_Code='" + strTaxGroup + "' and Tax_Group_Type='" + strType + "' and Tax_Recoverable='Y'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function
End Class

Public Class clsTaxGroupDetail
    Public Trans_Code As Integer = 0
    Public Tax_Group_Code As String = Nothing
    Public Tax_Group_Type As String = Nothing
    Public Tax_Code As String = Nothing
    Public Tax_Code_Desc As String = Nothing
    Public Taxable As Boolean = False
    Public Surtax As Boolean = False
    Public Excisable As Boolean = False
    Public Surtax_Tax_Code As String = Nothing
    Public Surtax_Tax_Code_Desc As String = Nothing
End Class

Public Class clsTaxMaster

    Public Tax_Code As String = Nothing
    Public Tax_Code_Desc As String = Nothing
    Public Tax_Liability_Account As String = Nothing
    Public Tax_Recoverable As Boolean = False
    Public Tax_Recoverable_Account As String = Nothing
    Public Tax_Recover_Rate As Double
    Public Tax_Net_Payable As String = Nothing
    Public Excisable As Boolean = False
    Public Type As String = Nothing
    Public Tax_Recoverable_Account2 As String = Nothing
    Public Tax_Recover_Rate2 As Double
    Public Tax_Recoverable_Account3 As String = Nothing
    Public Tax_Recover_Rate3 As Double
    Public Tax_Recoverable_Account4 As String = Nothing
    Public Tax_Recover_Rate4 As Double
    Public Tax_Recoverable_Account5 As String = Nothing
    Public Tax_Recover_Rate5 As Double
    Public PayableControl As String = Nothing
    Public DepositControl As String = Nothing


    ''Public Tax_Liability_Account As String = Nothing
    ''Public Tax_Recoverable_Account As String = Nothing
    ''Public Tax_Net_Payable As String = Nothing
    ''Public Type As String = Nothing

    Public Shared Function GetTaxDetailsForSale(ByVal strCode As String, ByVal trans As SqlTransaction) As clsTaxMaster
        Dim obj As clsTaxMaster = Nothing
        Dim qry As String = "select Tax_Liability_Account,Tax_Recoverable_Account,Tax_Net_Payable,Type,PayableControl,DepositControl from TSPL_TAX_MASTER where Tax_Code='" + strCode + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            obj = New clsTaxMaster()
            obj.Tax_Liability_Account = clsCommon.myCstr(dt.Rows(0)("Tax_Liability_Account"))
            obj.Tax_Recoverable_Account = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account"))
            obj.Tax_Net_Payable = clsCommon.myCstr(dt.Rows(0)("Tax_Net_Payable"))
            obj.DepositControl = clsCommon.myCstr(dt.Rows(0)("DepositControl"))
            obj.PayableControl = clsCommon.myCstr(dt.Rows(0)("PayableControl"))
            obj.Type = clsCommon.myCstr(dt.Rows(0)("Type"))
        End If
        Return obj
    End Function
    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = ""
        Dim whrqry As String = ""
        Dim GstApplicable As Boolean = False
        Dim GSTActiveTaxRateGroup As Boolean = False

        qry = " select Tax_Code as [Code],Tax_Code_Desc as [Description],Tax_Liability_Account as [Tax Liability Account],Tax_Recoverable as [Tax Recoverable],Tax_Recoverable_Account as [Tax Recoverable Account],Tax_Net_Payable as [Tax Net Payable],Created_By as [Created By],Created_Date as [Created Date],Modify_By as [Modify By],Modify_Date as [Modify Date],Comp_Code as [Company Code],Excisable as [Excisable],Type,Tax_Recoverable_Account2 as [Tax Recoverable Account2],Tax_Recoverable_Account3 as [Tax Recoverable Account3],Tax_Recoverable_Account4 as [Tax Recoverable Account4],Tax_Recoverable_Account5 as [Tax Recoverable Account5],Tax_Recover_Rate as [Tax Recover Rate],Tax_Recover_Rate2 as [Tax Recover Rate2],Tax_Recover_Rate3 as [Tax Recover Rate3],Tax_Recover_Rate4 as [Tax Recover Rate4],Tax_Recover_Rate5 as [Tax Recover Rate5],CURRENCY_CODE as [Currency Code],ConvRate as [Coversion Rate],ApplicableFrom  as [Applicable From],PayableControl as [Payable Control],DepositControl as [Deposit Control] from tspl_Tax_master "

        str = clsCommon.ShowSelectForm("RPTTXMSTFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    '----------------End of Code For Get Finder--------------------------------------------------------------'
    Public Shared Function GetTaxRecoverableAC(ByVal Tax_Code As String) As String
        Return GetTaxRecoverableAC(Tax_Code, Nothing)
    End Function

    Public Shared Function GetData(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As clsTaxMaster
        Dim obj As clsTaxMaster = Nothing
        If clsCommon.myLen(Tax_Code) > 0 Then
            Dim qry As String = "select * from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj = New clsTaxMaster
                obj.Tax_Code = clsCommon.myCstr(dt.Rows(0)("Tax_Code"))
                obj.Tax_Code_Desc = clsCommon.myCstr(dt.Rows(0)("Tax_Code_Desc"))
                obj.Tax_Liability_Account = clsCommon.myCstr(dt.Rows(0)("Tax_Liability_Account"))
                obj.Tax_Recoverable = IIf(clsCommon.CompairString("Y", clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))) = CompairStringResult.Equal, True, False)
                obj.Tax_Recoverable_Account = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account"))
                obj.Tax_Recover_Rate = clsCommon.myCdbl(dt.Rows(0)("Tax_Recover_Rate"))
                obj.Tax_Net_Payable = clsCommon.myCstr(dt.Rows(0)("Tax_Net_Payable"))
                obj.Excisable = IIf(clsCommon.CompairString("Y", clsCommon.myCstr(dt.Rows(0)("Excisable"))) = CompairStringResult.Equal, True, False)
                obj.Type = clsCommon.myCstr(dt.Rows(0)("Type"))
                obj.Tax_Recoverable_Account2 = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account2"))
                obj.Tax_Recover_Rate2 = clsCommon.myCdbl(dt.Rows(0)("Tax_Recover_Rate2"))
                obj.Tax_Recoverable_Account3 = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account3"))
                obj.Tax_Recover_Rate3 = clsCommon.myCdbl(dt.Rows(0)("Tax_Recover_Rate3"))
                obj.Tax_Recoverable_Account4 = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account4"))
                obj.Tax_Recover_Rate4 = clsCommon.myCdbl(dt.Rows(0)("Tax_Recover_Rate4"))
                obj.Tax_Recoverable_Account5 = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable_Account5"))
                obj.Tax_Recover_Rate5 = clsCommon.myCdbl(dt.Rows(0)("Tax_Recover_Rate5"))
                obj.DepositControl = clsCommon.myCstr(dt.Rows(0)("DepositControl"))
                obj.PayableControl = clsCommon.myCstr(dt.Rows(0)("PayableControl"))
            End If
        End If
        Return obj
    End Function

    Public Shared Function GetTaxRecoverableAC(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Tax_Recoverable_Account from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "'"
        Return clsDBFuncationality.getSingleValue(qry, trans)
    End Function

    Public Shared Function GetTaxRateChangable(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = "select Is_Change_Rate from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "'"
        Return (clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) = 1)
    End Function

    Public Shared Function GetTaxPayAC(ByVal Tax_Code As String) As String
        Return GetTaxPayAC(Tax_Code, Nothing)
    End Function

    Public Shared Function GetTaxPayAC(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Tax_Liability_Account from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "'"
        Return clsDBFuncationality.getSingleValue(qry, trans)
    End Function



    Public Shared Function IsTaxRecoverableAC(ByVal Tax_Code As String) As Boolean
        Return IsTaxRecoverableAC(Tax_Code, Nothing)
    End Function

    Public Shared Function ISTaxRecoverableAC(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = ""
        If clsCommon.myLen(Tax_Code) > 0 Then
            qry = "select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "'"
            qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        End If
        Dim bool As Boolean = IIf(clsCommon.CompairString(qry, "N") = CompairStringResult.Equal, False, True)
        Return bool
    End Function

    Public Shared Function IsTaxExcisable(ByVal Tax_Code As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = "select 1 from TSPL_TAX_MASTER where Tax_Code='" + Tax_Code + "' and Type='E'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    Public Shared Function GetTaxDetailsForPurchase(ByVal strGrpCode As String) As DataTable
        Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,Tax_Code,Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,Taxable,TSPL_TAX_GROUP_MASTER.Excisable,TSPL_TAX_GROUP_MASTER.Excisable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + strGrpCode + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"
        Return clsDBFuncationality.GetDataTable(qry)
    End Function

    Public Shared Function GetName(ByVal strCode As String) As String
        Dim qry As String = "select Tax_Code_Desc from TSPL_TAX_MASTER where  tax_code ='" + strCode + "'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
    End Function
    Private Shared Function CreateTableForTax() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("Tax1Rate", GetType(Double))
        dt.Columns.Add("Tax2Rate", GetType(Double))
        dt.Columns.Add("Tax3Rate", GetType(Double))
        dt.Columns.Add("Tax4Rate", GetType(Double))
        dt.Columns.Add("Tax5Rate", GetType(Double))
        dt.Columns.Add("Tax6Rate", GetType(Double))

        dt.Columns.Add("Tax1BaseAmt", GetType(Double))
        dt.Columns.Add("Tax2BaseAmt", GetType(Double))
        dt.Columns.Add("Tax3BaseAmt", GetType(Double))
        dt.Columns.Add("Tax4BaseAmt", GetType(Double))
        dt.Columns.Add("Tax5BaseAmt", GetType(Double))
        dt.Columns.Add("Tax6BaseAmt", GetType(Double))

        dt.Columns.Add("Tax1Amt", GetType(Double))
        dt.Columns.Add("Tax2Amt", GetType(Double))
        dt.Columns.Add("Tax3Amt", GetType(Double))
        dt.Columns.Add("Tax4Amt", GetType(Double))
        dt.Columns.Add("Tax5Amt", GetType(Double))
        dt.Columns.Add("Tax6Amt", GetType(Double))
        Return dt
    End Function
    Public Shared Function GetExcisableTaxRates(ByVal strICode As String, ByVal dblMRP As Double, ByVal strStartDate As String, ByVal dblItemBasicPrice As Double, ByVal strUOM As String, ByVal strTaxGroup As String, ByVal strPriceCode As String) As DataRow
        Dim drReturn As DataRow = CreateTableForTax().NewRow()
        Try
            Dim qry As String = "select "
            qry += " isnull(TAX1_Rate,0)*(case when TSPL_TAX_MASTER1.Excisable='Y' then 1 else  0 end) as TAX1_Rate,"
            qry += " isnull(TAX2_Rate,0)*(case when TSPL_TAX_MASTER2.Excisable='Y' then 1 else  0 end) as TAX2_Rate,"
            qry += " isnull(TAX3_Rate,0)*(case when TSPL_TAX_MASTER3.Excisable='Y' then 1 else  0 end) as TAX3_Rate,"
            qry += " isnull(TAX4_Rate,0)*(case when TSPL_TAX_MASTER4.Excisable='Y' then 1 else  0 end) as TAX4_Rate,"
            qry += " isnull(TAX5_Rate,0)*(case when TSPL_TAX_MASTER5.Excisable='Y' then 1 else  0 end) as TAX5_Rate,"
            qry += " isnull(TAX6_Rate,0)*(case when TSPL_TAX_MASTER6.Excisable='Y' then 1 else  0 end) as TAX6_Rate,(Item_Basic_Net * Abatement_Rate)/100 as BaseAssessableAmt,"
            qry += " Item_Code, Item_Basic_Net, Start_Date, Item_Basic_Price, UOM, Tax_group, Price_Code, TAX1, TAX2, TAX3, TAX4, TAX5, TAX6"
            qry += " from TSPL_ITEM_PRICE_MASTER"
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER1 on TSPL_TAX_MASTER1.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX1 "
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER2 on TSPL_TAX_MASTER2.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX2 "
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER3 on TSPL_TAX_MASTER3.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX3 "
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER4 on TSPL_TAX_MASTER4.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX4 "
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER5 on TSPL_TAX_MASTER5.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX5 "
            qry += " left outer join TSPL_TAX_MASTER as TSPL_TAX_MASTER6 on TSPL_TAX_MASTER6.Tax_Code=TSPL_ITEM_PRICE_MASTER.TAX6 "
            qry += " where Item_Code='" + strICode + "' and Item_Basic_Net='" + clsCommon.myCstr(dblMRP) + "' and Start_Date='" + strStartDate + "'  and Item_Basic_Price='" + clsCommon.myCstr(dblItemBasicPrice) + "' and UOM='" + strUOM + "' and Tax_group='" + strTaxGroup + "'  and Price_Code='" + strPriceCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)


            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim dblBaseAmt As Double = clsCommon.myCdbl(dt.Rows(0)("BaseAssessableAmt"))
                Dim dtTaxDetails As DataTable = clsTaxGroupMaster.GetTaxDetailsscrap(strTaxGroup)
                dtTaxDetails.Columns.Add("BaseAmt", GetType(Double))
                dtTaxDetails.Columns.Add("TaxAmt", GetType(Double))
                If dtTaxDetails IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If dtTaxDetails.Rows.Count >= 1 Then
                        dtTaxDetails.Rows(0)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
                    End If
                    If dtTaxDetails.Rows.Count >= 2 Then
                        dtTaxDetails.Rows(1)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
                    End If
                    If dtTaxDetails.Rows.Count >= 3 Then
                        dtTaxDetails.Rows(2)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
                    End If
                    If dtTaxDetails.Rows.Count >= 4 Then
                        dtTaxDetails.Rows(3)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
                    End If
                    If dtTaxDetails.Rows.Count >= 5 Then
                        dtTaxDetails.Rows(4)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
                    End If
                    If dtTaxDetails.Rows.Count >= 6 Then
                        dtTaxDetails.Rows(5)("TaxRate") = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
                    End If
                    For ii As Integer = 1 To dtTaxDetails.Rows.Count
                        Dim dblTBaseAmt As Double = 0
                        If clsCommon.CompairString(clsCommon.myCstr(dtTaxDetails.Rows(ii - 1)("Surtax")), "Y") = CompairStringResult.Equal Then
                            Dim strTSurTaxCode As String = clsCommon.myCstr(dtTaxDetails.Rows(ii - 1)("Surtax_Tax_Code"))
                            For jj As Integer = 1 To ii - 1
                                If clsCommon.CompairString(strTSurTaxCode, clsCommon.myCstr(dtTaxDetails.Rows(jj - 1)("Tax_Code"))) = CompairStringResult.Equal Then
                                    dblTBaseAmt += clsCommon.myCdbl(dtTaxDetails.Rows(jj - 1)("TaxAmt"))
                                End If
                            Next
                            dtTaxDetails.Rows(ii - 1)("BaseAmt") = dblTBaseAmt
                            dtTaxDetails.Rows(ii - 1)("TaxAmt") = dblTBaseAmt * clsCommon.myCdbl(dtTaxDetails.Rows(ii - 1)("TaxRate")) / 100
                        Else
                            dblTBaseAmt = dblBaseAmt
                            If clsCommon.CompairString(clsCommon.myCstr(dtTaxDetails.Rows(ii - 1)("Taxable")), "Y") = CompairStringResult.Equal Then
                                For jj As Integer = 1 To ii - 1
                                    If clsCommon.CompairString(clsCommon.myCstr(dtTaxDetails.Rows(jj - 1)("Taxable")), "Y") = CompairStringResult.Equal Then
                                        dblTBaseAmt += clsCommon.myCdbl(dtTaxDetails.Rows(jj - 1)("TaxAmt"))
                                    End If
                                Next
                            End If
                            dtTaxDetails.Rows(ii - 1)("BaseAmt") = dblTBaseAmt
                            dtTaxDetails.Rows(ii - 1)("TaxAmt") = dblTBaseAmt * clsCommon.myCdbl(dtTaxDetails.Rows(ii - 1)("TaxRate")) / 100
                        End If
                    Next
                End If

                For index As Integer = 1 To 6
                    If dtTaxDetails.Rows.Count >= index Then
                        drReturn("Tax" + clsCommon.myCstr(index) + "Rate") = clsCommon.myCdbl(dtTaxDetails.Rows(index - 1)("TaxRate"))
                        drReturn("Tax" + clsCommon.myCstr(index) + "BaseAmt") = clsCommon.myCdbl(dtTaxDetails.Rows(index - 1)("BaseAmt"))
                        drReturn("Tax" + clsCommon.myCstr(index) + "Amt") = clsCommon.myCdbl(dtTaxDetails.Rows(index - 1)("TaxAmt"))
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return drReturn
    End Function

End Class


Public Class clsTempItemTaxDetails
    Public AuthorityCode As String = Nothing
    Public AuthorityName As String = Nothing
    Public Rate As Double = 0
    Public BaseAmt As Double = 0
    Public TaxAmt As Double = 0
    Public isSurTax As Boolean = False
    Public SurTax As String = Nothing
    Public IsTaxable As Boolean = False
    Public TaxOnBaseAmount As Boolean = False
End Class

Public Class clsTempDrCrAmt
    Public DrAmt As Double = 0
    Public CrAmt As Double = 0
End Class

Public Class clsTempUOMForConversion
    Public FB As Double = 0
    Public Raw As Double = 0
    Public Converted As Double = 0
    Public OZ As Double = 0
    Public Shared Function GetConvertsionFactors(ByVal strICode As String, ByVal trans As SqlTransaction) As clsTempUOMForConversion
        Dim obj As clsTempUOMForConversion = Nothing
        Dim qry As String = "select UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where Item_Code='" + strICode + "'"
        Dim dtUOM As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dtUOM IsNot Nothing AndAlso dtUOM.Rows.Count > 0 Then
            obj = New clsTempUOMForConversion()
            For Each dr As DataRow In dtUOM.Rows
                If clsCommon.CompairString(clsCommon.myCstr(dr("UOM_Code")), "FB") = CompairStringResult.Equal Then
                    obj.Raw = clsCommon.myCdbl(dr("Conversion_Factor"))
                ElseIf clsCommon.CompairString(clsCommon.myCstr(dr("UOM_Code")), "8oz") = CompairStringResult.Equal Then
                    obj.OZ = clsCommon.myCdbl(dr("Conversion_Factor"))
                ElseIf clsCommon.CompairString(clsCommon.myCstr(dr("UOM_Code")), "Con") = CompairStringResult.Equal Then
                    obj.Converted = clsCommon.myCdbl(dr("Conversion_Factor"))
                End If
            Next
        End If
        Return obj
    End Function

End Class


Public Class clsDepartment
    Public Code As String = ""
    Public Name As String = ""

    Public Shared Function Finder(ByVal strCode As String, ByVal isButtonClicked As Boolean) As clsDepartment
        Dim obj As clsDepartment = Nothing
        Dim qry As String = "select Segment_code as Code,Description as Name from TSPL_GL_SEGMENT_CODE"
        Dim WhrCls As String = "Seg_No='3'"
        strCode = clsCommon.ShowSelectForm("DepartmentFilter", qry, "Code", WhrCls, strCode, "Code", isButtonClicked)
        If clsCommon.myLen(strCode) > 0 Then
            qry = "select Description as Name from TSPL_GL_SEGMENT_CODE where Seg_No='3' and Segment_code ='" + strCode + "' "
            obj = New clsDepartment()
            obj.Code = strCode
            obj.Name = clsDBFuncationality.getSingleValue(qry)
        End If
        Return obj
    End Function
End Class



Public Class clsItemPriceMaster
    Public item_Description As String

    Public Item_Code As String = ""
    Public Item_MRP As Double
    Public Item_Baisc_Price As Double
    Public Abatement As Double
    Public Item_Basic_Net As Double
    Public Start_Date As String
    Public End_Date As Date
    Public Empty_Value_Shell As Double
    Public Can_Edit As Char
    Public NetLTPT As Double
    Public Price_Category As String
    Public Item_Basic_Price As Double
    Public Tax_group As String
    Public TAX1 As String
    Public TAX1_Rate As Double
    Public TAX1_Amt As Double
    Public TAX2 As String
    Public TAX2_Rate As Double
    Public TAX2_Amt As Double
    Public TAX3 As String
    Public TAX3_Rate As Double
    Public TAX3_Amt As Double
    Public TAX4 As String
    Public TAX4_Rate As Double
    Public TAX4_Amt As Double
    Public TAX5 As String
    Public TAX5_Rate As Double
    Public TAX5_Amt As Double
    Public TAX6 As String
    Public TAX6_Rate As Double
    Public TAX6_Amt As Double
    Public TAX7 As String
    Public TAX7_Rate As Double
    Public TAX7_Amt As Double
    Public TAX8 As String
    Public TAX8_Rate As Double
    Public TAX8_Amt As Double
    Public TAX9 As String
    Public TAX9_Rate As Double
    Public TAX9_Amt As Double
    Public TAX10 As String
    Public TAX10_Rate As Double
    Public TAX10_Amt As Double
    Public Price_Code As String
    Public Price_Comp1 As String
    Public Price_Comp_Desc1 As String
    Public Price_Rate1 As Double
    Public Price_Amount1 As Double
    Public Price_Comp2 As String
    Public Price_Comp_Desc2 As String
    Public Price_Rate2 As Double
    Public Price_Amount2 As Double
    Public Price_Comp3 As String
    Public Price_Comp_Desc3 As String
    Public Price_Rate3 As Double
    Public Price_Amount3 As Double
    Public Price_Comp4 As String
    Public Price_Comp_Desc4 As String
    Public Price_Rate4 As Double
    Public Price_Amount4 As Double
    Public Price_Comp5 As String
    Public Price_Comp_Desc5 As String
    Public Price_Rate5 As Double
    Public Price_Amount5 As Double
    Public Price_Comp6 As String
    Public Price_Comp_Desc6 As String
    Public Price_Rate6 As Double
    Public Price_Amount6 As Double
    Public Price_Comp7 As String
    Public Price_Comp_Desc7 As String
    Public Price_Rate7 As Double
    Public Price_Amount7 As Double
    Public Price_Comp8 As String
    Public Price_Comp_Desc8 As String
    Public Price_Rate8 As Double
    Public Price_Amount8 As Double
    Public Price_Comp9 As String
    Public Price_Comp_Desc9 As String
    Public Price_Rate9 As Double
    Public Price_Amount9 As Double
    Public Price_Comp10 As String
    Public Price_Comp_Desc10 As String
    Public Price_Rate10 As Double
    Public Price_Amount10 As Double
    Public Item_Rate As Double
    Public Liquid_Rate As Double
    Public Stock_Rate As Double
    Public Abatement_Rate As Double
    Public UOM As String
    Public Price_Code_Desc As String
    Public Empty_Value_Bottle As Double
    Public Item_Price_Id_No As Integer

    Public Shared Function SaveDataFromExport(ByVal arr As List(Of clsItemPriceMaster)) As Boolean
        If arr Is Nothing OrElse arr.Count <= 0 Then
            Throw New Exception("No Data Found to Save")
        End If
        Dim isSaved As Boolean = True
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try

            Dim intPrevPriceID As Integer = -1
            Dim intItem_Price_Id_No As Integer
            Dim strCurrDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd-MMM-yyyy")
            For Each obj As clsItemPriceMaster In arr
                Dim qry As String = "delete from tspl_item_price_master where item_code='" + obj.Item_Code + "' and Uom='" + obj.UOM + "' and  start_date=Convert(date,'" + obj.Start_Date + "',103) and price_code='" + obj.Price_Code + "' and item_basic_net='" + clsCommon.myCstr(obj.Item_Basic_Net) + "' and Item_Basic_Price='" + clsCommon.myCstr(obj.Item_Basic_Price) + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                If Not intPrevPriceID = obj.Item_Price_Id_No Then
                    qry = "select max(Item_Price_Id_No) from TSPL_ITEM_PRICE_MASTER"
                    intItem_Price_Id_No = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) + 1
                    intPrevPriceID = obj.Item_Price_Id_No
                End If
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Item_Price_Id_No", intItem_Price_Id_No)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Item_MRP", obj.Item_MRP)
                clsCommon.AddColumnsForChange(coll, "Abatement", obj.Abatement)
                clsCommon.AddColumnsForChange(coll, "Item_Basic_Net", obj.Item_Basic_Net)
                clsCommon.AddColumnsForChange(coll, "Start_Date", obj.Start_Date)
                clsCommon.AddColumnsForChange(coll, "Empty_Value_Shell", obj.Empty_Value_Shell)
                clsCommon.AddColumnsForChange(coll, "NetLTPT", obj.NetLTPT)
                clsCommon.AddColumnsForChange(coll, "Item_Basic_Price", obj.Item_Basic_Price)
                clsCommon.AddColumnsForChange(coll, "Tax_group", obj.Tax_group)

                clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
                clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)

                clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
                clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)

                clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
                clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)


                clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
                clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)

                clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
                clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)

                clsCommon.AddColumnsForChange(coll, "Price_Code", obj.Price_Code)

                clsCommon.AddColumnsForChange(coll, "Price_Comp1", obj.Price_Comp1)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc1", obj.Price_Comp_Desc1)
                clsCommon.AddColumnsForChange(coll, "Price_Rate1", obj.Price_Rate1)
                clsCommon.AddColumnsForChange(coll, "Price_Amount1", obj.Price_Amount1)

                clsCommon.AddColumnsForChange(coll, "Price_Comp2", obj.Price_Comp2)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc2", obj.Price_Comp_Desc2)
                clsCommon.AddColumnsForChange(coll, "Price_Rate2", obj.Price_Rate2)
                clsCommon.AddColumnsForChange(coll, "Price_Amount2", obj.Price_Amount2)

                clsCommon.AddColumnsForChange(coll, "Price_Comp3", obj.Price_Comp3)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc3", obj.Price_Comp_Desc3)
                clsCommon.AddColumnsForChange(coll, "Price_Rate3", obj.Price_Rate3)
                clsCommon.AddColumnsForChange(coll, "Price_Amount3", obj.Price_Amount3)

                clsCommon.AddColumnsForChange(coll, "Price_Comp4", obj.Price_Comp4)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc4", obj.Price_Comp_Desc4)
                clsCommon.AddColumnsForChange(coll, "Price_Rate4", obj.Price_Rate4)
                clsCommon.AddColumnsForChange(coll, "Price_Amount4", obj.Price_Amount4)

                clsCommon.AddColumnsForChange(coll, "Price_Comp5", obj.Price_Comp5)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc5", obj.Price_Comp_Desc5)
                clsCommon.AddColumnsForChange(coll, "Price_Rate5", obj.Price_Rate5)
                clsCommon.AddColumnsForChange(coll, "Price_Amount5", obj.Price_Amount5)

                clsCommon.AddColumnsForChange(coll, "Price_Comp6", obj.Price_Comp6)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc6", obj.Price_Comp_Desc6)
                clsCommon.AddColumnsForChange(coll, "Price_Rate6", obj.Price_Rate6)
                clsCommon.AddColumnsForChange(coll, "Price_Amount6", obj.Price_Amount6)

                clsCommon.AddColumnsForChange(coll, "Price_Comp7", obj.Price_Comp7)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc7", obj.Price_Comp_Desc7)
                clsCommon.AddColumnsForChange(coll, "Price_Rate7", obj.Price_Rate7)
                clsCommon.AddColumnsForChange(coll, "Price_Amount7", obj.Price_Amount7)

                clsCommon.AddColumnsForChange(coll, "Price_Comp8", obj.Price_Comp8)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc8", obj.Price_Comp_Desc8)
                clsCommon.AddColumnsForChange(coll, "Price_Rate8", obj.Price_Rate8)
                clsCommon.AddColumnsForChange(coll, "Price_Amount8", obj.Price_Amount8)

                clsCommon.AddColumnsForChange(coll, "Price_Comp9", obj.Price_Comp9)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc9", obj.Price_Comp_Desc9)
                clsCommon.AddColumnsForChange(coll, "Price_Rate9", obj.Price_Rate9)
                clsCommon.AddColumnsForChange(coll, "Price_Amount9", obj.Price_Amount9)

                clsCommon.AddColumnsForChange(coll, "Price_Comp10", obj.Price_Comp10)
                clsCommon.AddColumnsForChange(coll, "Price_Comp_Desc10", obj.Price_Comp_Desc10)
                clsCommon.AddColumnsForChange(coll, "Price_Rate10", obj.Price_Rate10)
                clsCommon.AddColumnsForChange(coll, "Price_Amount10", obj.Price_Amount10)

                clsCommon.AddColumnsForChange(coll, "Item_Rate", obj.Item_Rate)
                clsCommon.AddColumnsForChange(coll, "Liquid_Rate", obj.Liquid_Rate)
                clsCommon.AddColumnsForChange(coll, "Stock_Rate", obj.Stock_Rate)
                clsCommon.AddColumnsForChange(coll, "Abatement_Rate", obj.Abatement_Rate)
                clsCommon.AddColumnsForChange(coll, "UOM", obj.UOM)
                clsCommon.AddColumnsForChange(coll, "Price_Code_Desc", obj.Price_Code_Desc)
                clsCommon.AddColumnsForChange(coll, "Empty_Value_Bottle", obj.Empty_Value_Bottle)

                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", strCurrDate)
                clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Modify_Date", strCurrDate)
                clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)

                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_ITEM_PRICE_MASTER", OMInsertOrUpdate.Insert, "", trans)
            Next

            If isSaved Then
                trans.Commit()
            End If

        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


    Public Shared Function FinderForItemPrices() As clsItemPriceMaster
        Dim obj As clsItemPriceMaster = Nothing
        Dim qry As String = "select Final.Item_Code as Code,TSPL_ITEM_MASTER.Item_Desc as [Description],Final.UOM,Final.Item_Basic_Net as MRP,Final.Abatement as [Assessble Amount] from( " & _
        " select Item_Code,UOM,Item_Basic_Net,Abatement from TSPL_ITEM_PRICE_MASTER group by Item_Code,UOM,Item_Basic_Net,Abatement " & _
        " )Final " & _
        " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER .Item_Code=Final.Item_Code " & _
        " order by Final.Item_Code,Final.UOM,Final.Item_Basic_Net,Final.Abatement"
        Dim dr As DataRow = clsCommon.ShowSelectFormForRow("POFGItem", qry)
        If dr IsNot Nothing Then
            obj = New clsItemPriceMaster
            obj.Item_Code = clsCommon.myCstr(dr("Code"))
            obj.item_Description = clsCommon.myCstr(dr("Description"))
            obj.UOM = clsCommon.myCstr(dr("UOM"))
            obj.Item_MRP = clsCommon.myCdbl(dr("MRP"))
            obj.Abatement = clsCommon.myCdbl(dr("Assessble Amount"))
        End If
        Return obj
    End Function
    Public Shared Function GetMRPOfFinishItem(ByVal strFaterCode As String, ByVal strItemUOM As String) As Double
        Return GetMRPOfFinishItem(strFaterCode, strItemUOM, Nothing)
    End Function
    Public Shared Function GetMRPOfFinishItem(ByVal strFaterCode As String, ByVal strItemUOM As String, ByVal trans As SqlTransaction) As Double
        Dim strNewUOM As String = ""
        If clsCommon.CompairString(strItemUOM, "FC") = CompairStringResult.Equal Then
            strNewUOM = "EC"
        ElseIf clsCommon.CompairString(strItemUOM, "FB") = CompairStringResult.Equal Then
            strNewUOM = "EB"
        End If
        Dim qry As String = "select Item_Basic_Net from TSPL_ITEM_PRICE_MASTER where Item_Code='" + strFaterCode + "' and UOM='" + strNewUOM + "'"
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
    End Function
End Class


'==========BM00000007746,Rohit(On 27-Aug-2015)===================
Public Class clsVendorMaster

    Public Vendor_Code As String = ""
    Public Vendor_Name As String = ""
    Public Add1 As String
    Public Add2 As String
    Public Add3 As String
    Public Closing_Date As String
    Public Vendor_Group_Code As String = ""
    Public Vendor_Group_Code_Desc As String = Nothing
    Public City_Code As String = Nothing
    Public City_Code_Desc As String = Nothing
    Public State As String = Nothing
    Public Country As String
    Public Phone1 As String
    Public Phone2 As String
    Public Fax As String
    Public Email As String
    Public WebSite As String
    Public Contact_Person_Name As String
    Public Contact_Person_Phone As String
    Public Contact_Person_Fax As String
    Public Contact_Person_Website As String
    Public Contact_Person_Email As String
    Public Terms_Code As String = ""
    Public Terms_Code_Desc As String = ""
    Public Vendor_Account As String
    Public Vendor_Account_Desc As String
    Public Payment_Code As String
    Public Payment_Code_Desc As String
    Public Bank_Code As String
    Public Bank_Code_Desc As String
    Public Tax_Group As String = ""
    Public Tax_Group_Desc As String = ""
    Public Ven_Type_Code As String
    Public Ven_Type_Desc As String
    Public TAX1 As String
    Public TAX1_Rate As Double
    Public TAX2 As String
    Public TAX2_Rate As Double
    Public TAX3 As String
    Public TAX3_Rate As Double
    Public TAX4 As String
    Public TAX4_Rate As Double
    Public TAX5 As String
    Public TAX5_Rate As Double
    Public TAX6 As String
    Public TAX6_Rate As Double
    Public TAX7 As String
    Public TAX7_Rate As Double
    Public TAX8 As String
    Public TAX8_Rate As Double
    Public TAX9 As String
    Public TAX9_Rate As Double
    Public TAX10 As String
    Public TAX10_Rate As Double
    Public Service_Tax_No As String
    Public Tin_No As String
    Public Lst_No As String
    Public Status As Char
    Public OnHold As Char
    Public Transporter As Char
    Public Remarks1 As String
    Public Remarks2 As String
    Public Additional1 As String
    Public Additional2 As String
    Public Additional3 As String
    Public Credit_Limit As Double
    Public Created_By As String
    Public Created_Date As String
    Public Modify_By As String
    Public Modify_Date As String
    Public Comp_Code As String
    Public CST As String
    Public ECC As String
    Public Range As String
    Public Collectorate As String
    Public PAN As String
    Public Is_Gross_Receipt As Integer = 0
    Public Inter_Branch As Char
    Public CURRENCY_CODE As String = ""
    Public franchise_yn As Char
    Public Form_Type As String
    Public State_Code As String
    Public Country_Code As String
    Public Service_charges As Double
    Public commision_pers As Double
    Public incentive As String
    Public incentive_days As Double
    Public vsp_payment As String
    Public VSP_Payee_Name As String
    Public Zila As String
    Public Tehsil As String
    Public Branch_Name As String
    Public IFCI_Code As String
    Public Account_No As String
    Public Industry_Type As String
    Public Industry_Person As String
    Public Agreement As String
    Public Security_Cheque As String
    Public No_of_Installment As Double
    Public Amount_of_Installment As Double
    Public IsPermanent As Char
    Public IsTemporary As Char
    Public Joint_Name As String
    Public Service_Charge_Type As String
    Public Is_Parent_Vendor As Char
    Public Parent_Vendor_Code As String
    Public branch_code As String
    Public Category_Struct_Code As String
    Public Bank_Name As String
    Public IFSC_Code As String
    Public Account_Type As String
    Public Vendor_Type As String
    Public payment_commision_pers As Double
    Public Pin_Code As String
    Public Security_Amount As Double
    Public AMC_Charge As Double
    Public AMCU As String
    Public Billing_Date As Date? = Nothing
    Public Is_Chilling_vendor As String
    Public TDS_Branch_Code As String
    Public Deduction_Code As String
    Public TDS_State_Code As String
    Public TDS_Vendor_Type As String
    Public TDS_Status As String
    Public Is_TDS_Applicable As Integer = 0
    Public Nature As String
    Public Actual_charges As Double
    Public Joint_Bank_Code As String
    Public Joint_Account_No As String
    Public Start_Date As Date? = Nothing
    Public End_Date As Date? = Nothing
    Public CSA_Type As Char
    Public Start_Period As Date? = Nothing
    Public Expired_Period As Date? = Nothing
    Public PC_CODE As String
    Public IsBlankCheque As Integer = 0
    Public Alies_Name As String
    Public Vendor_Type_CHA As String
    Public Is_Head_Load As String
    Public Rate_Head_Load As Double
    Public Service_Basis_Head_Load As String
    Public Is_Own_Asset As String
    Public Rate_Own_Asset As Double
    Public Service_Basis_Own_Asset As String
    Public IsVendorInvoiceNo As Integer = 0
    Public CHA_DOC_NO As String
    Public Standard_Security_Amount As Double
    Public Is_TC_Certified As Char
    Public TC_Certified As String
    Public MP_Code As String
    Public MP_Name As String
    Public Cheque_In_Favour_Of As String
    Public is_Drip_Saver As String
    Public Other_For_PAN As Integer = 0
    Public Joint_Branch_Name As String
    Public Joint_IFSC_Code As String
    Public VSP_Farmer_Billing As Integer = 0
    Public EMP_Type As String
    Public EMP_Fixed_Amount As Double
    Public Actual_charges_Slab As Integer = 0
    Public Actual_charges_Slab2 As Integer = 0
    Public Actual_charges2 As Double
    Public Actual_charges_Slab3 As Integer = 0
    Public Actual_charges3 As Double
    Public Actual_charges_Slab4 As Integer = 0
    Public Actual_charges4 As Double
    Public Actual_charges_Slab5 As Integer = 0
    Public Actual_charges5 As Double
    Public Vendor_Distance As String
    Public Apply_Mult_Incentive As Integer = 0
    Public MinimumQtyRequired_pertrip As Double
    Public OldName As String
    Public Security_Deduction_Amount As Double
    Public Interest_Per As Double
    Public Minimum_Interest As Double
    Public SSI_No As String
    Public Is_Blacklist As Integer = 0
    Public Service_Charge_Per_Unit As Double
    Public is_Hold_Payment_Process As Integer = 0
    Public Is_Inactive_In_Milk_Procurement As Integer = 0
    Public DFOption As String
    Public BusinessCondition As String
    Public GSTRegistered As Integer = 0
    Public GST_Composition_scheme As Integer = 0
    Public GSTEntity As String
    Public GSTLastEntity As String
    Public GSTFinalNo As String
    Public GSTMiddle As String
    Public Weight As Double
    Public JWPriceCode As String
    Public CorrectionFat As Decimal
    Public CorrectionSNF As Decimal
    Public IsEmployee As Integer = 0
    Public Handling_Charges_Per As Double
    Public Credit_Limit_On_Milk_Receipt_Per As Double
    Public isBulkProcurement As Integer = 0
    Public isHighClass As Integer = 0
    Public Bulk_ROUTE_NO As String
    Public Monthly_Rent As Double
    Public EMP_CODE As String
    Public TIP_Buffalo As Double
    Public TIP_Cow As Double
    Public TIP_Mix As Double
    Public Registration_No As String
    Public Aadhar_No As String
    Public Care_Of As String
    Public IsTCSnotApplicable As Integer = 0
    Public Isbuyerfilereturninlasttwoyears As Integer = 0
    Public IsTCS_TDSamountgreaterthan50KpreviousYear As Integer = 0
    Public SecChequeNoLac1 As String
    Public SecChequeNoRs100 As String
    Public TDS_Branch_Code_Service As String
    Public Deduction_Code_Service As String
    Public TDS_Status_Service As String
    Public TDS_Vendor_Type_Service As String
    Public TDS_State_Code_Service As String
    Public DistanceKM_Head_Load As Double
    Public Vendor_Name_Hindi As String
    Public Registered_PDCS_CLUSTER As String
    Public StartDate As Date? = Nothing
    Public BankCode2 As String
    Public BankName2 As String
    Public Credit2 As Double
    Public IFSCCode2 As String
    Public AccNo2 As String
    Public AccountType2 As String
    Public BankBranch2 As String
    Public SecurityCharges2 As String
    Public SupervisorOrRP As String
    Public Active_Date As Date? = Nothing
    Public Gender As String
    Public RegistrationNo As String
    Public RegistrationDate As Date? = Nothing
    Public DISTRICT_Code As String
    Public Zone_Code As String
    Public CAST_CATEGORY_CODE As String
    Public BLOCK_CODE As String
    Public Company_Bank As String
    Public REVENUE_VILLAGE_CODE As String
    Public GRAMPANCHAYAT_CODE As String
    Public PANCHAYAT_SAMITI_CODE As String
    Public VIDHAN_SABHA_CODE As String
    Public IsAllowSkip_PurchaseQC As Integer = 0
    Public OEM As Integer = 0
    Public Is_Provisional As Integer = 0
    Public Company_Bank_Current As String
    Public In_Active_CF As Char
    Public Is_Default_Grower As Integer = 0



    Public VSP_CODE As String
    Public Charge_CODE As String
    Public GL_CODE As String
    Public Updated_date As DateTime
    Public Charge_Desc As String
    Public GL_DESC As String
    Public Rate As String


    Public CUSTOMER_FORM_TYPE As String = String.Empty

    Public Function SaveData(ByVal obj As clsVendorMaster, ByVal isNewEntry As Boolean) As Boolean
        Dim ISSaved As Boolean = True
        Dim trans As SqlTransaction = Nothing
        Try
            trans = clsDBFuncationality.GetTransactin()
            ISSaved = SaveData(obj, isNewEntry, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            clsCommon.MyMessageBoxShow(ex.Message)
            Return False
            'Throw New Exception(ex.Message)
        End Try
        Return ISSaved
    End Function

    Public Function SaveData(ByVal obj As clsVendorMaster, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try
            If (obj.Is_Default_Grower = 1) Then
                Dim qry As String = "update TSPL_VENDOR_MASTER set Is_Default_Grower=0 where Is_Default_Grower=1"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Form_Type", obj.Form_Type)
            clsCommon.AddColumnsForChange(coll, "Vendor_Name", obj.Vendor_Name)
            clsCommon.AddColumnsForChange(coll, "Vendor_Name_Hindi", obj.Vendor_Name_Hindi, True, True)
            clsCommon.AddColumnsForChange(coll, "Pin_Code", obj.Pin_Code)
            clsCommon.AddColumnsForChange(coll, "Add1", obj.Add1)
            clsCommon.AddColumnsForChange(coll, "Add2", obj.Add2)
            clsCommon.AddColumnsForChange(coll, "Add3", obj.Add3)
            If clsCommon.CompairString(obj.Status, "Y") = CompairStringResult.Equal Then
                clsCommon.AddColumnsForChange(coll, "Closing_Date", obj.Closing_Date)
            Else
                clsCommon.AddColumnsForChange(coll, "Closing_Date", Nothing)
            End If
            clsCommon.AddColumnsForChange(coll, "Vendor_Group_Code", obj.Vendor_Group_Code)
            clsCommon.AddColumnsForChange(coll, "Vendor_Group_Code_Desc", obj.Vendor_Group_Code_Desc)
            clsCommon.AddColumnsForChange(coll, "City_Code", obj.City_Code)
            clsCommon.AddColumnsForChange(coll, "City_Code_Desc", obj.City_Code_Desc)
            clsCommon.AddColumnsForChange(coll, "State", obj.State)
            clsCommon.AddColumnsForChange(coll, "Country", obj.Country)
            clsCommon.AddColumnsForChange(coll, "Phone1", obj.Phone1)
            clsCommon.AddColumnsForChange(coll, "Phone2", obj.Phone2)
            clsCommon.AddColumnsForChange(coll, "Fax", obj.Fax)
            clsCommon.AddColumnsForChange(coll, "Email", obj.Email)
            clsCommon.AddColumnsForChange(coll, "WebSite", obj.WebSite)
            clsCommon.AddColumnsForChange(coll, "Contact_Person_Name", obj.Contact_Person_Name)
            clsCommon.AddColumnsForChange(coll, "Contact_Person_Phone", obj.Contact_Person_Phone)
            clsCommon.AddColumnsForChange(coll, "Contact_Person_Fax", obj.Contact_Person_Fax)
            clsCommon.AddColumnsForChange(coll, "Contact_Person_Website", obj.Contact_Person_Website)
            clsCommon.AddColumnsForChange(coll, "Contact_Person_Email", obj.Contact_Person_Email)
            clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
            clsCommon.AddColumnsForChange(coll, "Terms_Code_Desc", obj.Terms_Code_Desc)
            clsCommon.AddColumnsForChange(coll, "Vendor_Account", obj.Vendor_Account)
            clsCommon.AddColumnsForChange(coll, "Vendor_Account_Desc", obj.Vendor_Account_Desc)
            clsCommon.AddColumnsForChange(coll, "Payment_Code", obj.Payment_Code)
            clsCommon.AddColumnsForChange(coll, "Payment_Code_Desc", obj.Payment_Code_Desc)
            clsCommon.AddColumnsForChange(coll, "Bank_Code", obj.Bank_Code)
            clsCommon.AddColumnsForChange(coll, "Bank_Code_Desc", obj.Bank_Code_Desc)
            clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
            clsCommon.AddColumnsForChange(coll, "Tax_Group_Desc", obj.Tax_Group_Desc)
            clsCommon.AddColumnsForChange(coll, "Ven_Type_Code", obj.Ven_Type_Code)
            clsCommon.AddColumnsForChange(coll, "Ven_Type_Desc", obj.Ven_Type_Desc)
            clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
            clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
            clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
            clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
            clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
            clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
            clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
            clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
            clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
            clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
            clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
            clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
            clsCommon.AddColumnsForChange(coll, "Service_Tax_No", obj.Service_Tax_No)
            clsCommon.AddColumnsForChange(coll, "Tin_No", obj.Tin_No)
            clsCommon.AddColumnsForChange(coll, "Lst_No", obj.Lst_No)
            clsCommon.AddColumnsForChange(coll, "Status", obj.Status)
            clsCommon.AddColumnsForChange(coll, "OnHold", obj.OnHold)
            clsCommon.AddColumnsForChange(coll, "Transporter", obj.Transporter)
            clsCommon.AddColumnsForChange(coll, "Remarks1", obj.Remarks1)
            clsCommon.AddColumnsForChange(coll, "Remarks2", obj.Remarks2)
            clsCommon.AddColumnsForChange(coll, "Additional1", obj.Additional1)
            clsCommon.AddColumnsForChange(coll, "Additional2", obj.Additional2)
            clsCommon.AddColumnsForChange(coll, "Additional3", obj.Additional3)
            clsCommon.AddColumnsForChange(coll, "Credit_Limit", obj.Credit_Limit)
            clsCommon.AddColumnsForChange(coll, "Is_Gross_Receipt", obj.Is_Gross_Receipt)
            clsCommon.AddColumnsForChange(coll, "Inter_Branch", obj.Inter_Branch)
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "franchise_yn", obj.franchise_yn, True)
            clsCommon.AddColumnsForChange(coll, "State_Code", obj.State_Code)
            clsCommon.AddColumnsForChange(coll, "Country_Code", obj.Country_Code)
            clsCommon.AddColumnsForChange(coll, "Service_charges", obj.Service_charges)
            clsCommon.AddColumnsForChange(coll, "commision_pers", obj.commision_pers)
            clsCommon.AddColumnsForChange(coll, "incentive", obj.incentive)
            clsCommon.AddColumnsForChange(coll, "incentive_days", obj.incentive_days)
            clsCommon.AddColumnsForChange(coll, "vsp_payment", obj.vsp_payment)
            clsCommon.AddColumnsForChange(coll, "VSP_Payee_Name", obj.VSP_Payee_Name)
            clsCommon.AddColumnsForChange(coll, "Zila", obj.Zila)
            clsCommon.AddColumnsForChange(coll, "Tehsil", obj.Tehsil)
            clsCommon.AddColumnsForChange(coll, "Branch_Name", obj.Branch_Name, True)
            clsCommon.AddColumnsForChange(coll, "IFCI_Code", obj.IFCI_Code, True)
            clsCommon.AddColumnsForChange(coll, "Account_No", obj.Account_No, True)
            clsCommon.AddColumnsForChange(coll, "Industry_Type", obj.Industry_Type)
            clsCommon.AddColumnsForChange(coll, "Industry_Person", obj.Industry_Person)
            clsCommon.AddColumnsForChange(coll, "Agreement", obj.Agreement)
            clsCommon.AddColumnsForChange(coll, "Security_Cheque", obj.Security_Cheque)
            clsCommon.AddColumnsForChange(coll, "No_of_Installment", obj.No_of_Installment)
            clsCommon.AddColumnsForChange(coll, "Amount_of_Installment", obj.Amount_of_Installment)
            clsCommon.AddColumnsForChange(coll, "IsPermanent", obj.IsPermanent)
            clsCommon.AddColumnsForChange(coll, "IsTemporary", obj.IsTemporary)
            clsCommon.AddColumnsForChange(coll, "Joint_Name", obj.Joint_Name)
            clsCommon.AddColumnsForChange(coll, "Service_Charge_Type", obj.Service_Charge_Type)
            clsCommon.AddColumnsForChange(coll, "Is_Parent_Vendor", obj.Is_Parent_Vendor)
            clsCommon.AddColumnsForChange(coll, "Parent_Vendor_Code", obj.Parent_Vendor_Code)
            clsCommon.AddColumnsForChange(coll, "branch_code", obj.branch_code)
            clsCommon.AddColumnsForChange(coll, "Category_Struct_Code", obj.Category_Struct_Code)
            clsCommon.AddColumnsForChange(coll, "Bank_Name", obj.Bank_Name, True)
            clsCommon.AddColumnsForChange(coll, "IFSC_Code", obj.IFSC_Code, True)
            clsCommon.AddColumnsForChange(coll, "Account_Type", obj.Account_Type, True)
            clsCommon.AddColumnsForChange(coll, "Vendor_Type", obj.Vendor_Type)
            clsCommon.AddColumnsForChange(coll, "payment_commision_pers", obj.payment_commision_pers)
            clsCommon.AddColumnsForChange(coll, "Security_Amount", obj.Security_Amount)
            clsCommon.AddColumnsForChange(coll, "AMC_Charge", obj.AMC_Charge)
            clsCommon.AddColumnsForChange(coll, "AMCU", obj.AMCU)
            If obj.Billing_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Billing_Date", clsCommon.GetPrintDate(obj.Billing_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Billing_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "Is_Chilling_vendor", obj.Is_Chilling_vendor)
            clsCommon.AddColumnsForChange(coll, "TDS_Branch_Code", obj.TDS_Branch_Code)
            clsCommon.AddColumnsForChange(coll, "Deduction_Code", obj.Deduction_Code)
            clsCommon.AddColumnsForChange(coll, "TDS_State_Code", obj.TDS_State_Code)
            clsCommon.AddColumnsForChange(coll, "TDS_Vendor_Type", obj.TDS_Vendor_Type)
            clsCommon.AddColumnsForChange(coll, "TDS_Status", obj.TDS_Status)
            clsCommon.AddColumnsForChange(coll, "Is_TDS_Applicable", obj.Is_TDS_Applicable)
            clsCommon.AddColumnsForChange(coll, "Nature", obj.Nature)
            clsCommon.AddColumnsForChange(coll, "Actual_charges", obj.Actual_charges)
            clsCommon.AddColumnsForChange(coll, "Joint_Bank_Code", obj.Joint_Bank_Code)
            If obj.Start_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Start_Date", clsCommon.GetPrintDate(obj.Start_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Start_Date", Nothing, True)
            End If
            If obj.End_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "End_Date", clsCommon.GetPrintDate(obj.End_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "End_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "CSA_Type", obj.CSA_Type)
            If obj.Start_Period IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Start_Period", clsCommon.GetPrintDate(obj.Start_Period, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Start_Period", Nothing, True)
            End If
            If obj.Expired_Period IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Expired_Period", clsCommon.GetPrintDate(obj.Expired_Period, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Expired_Period", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "PC_CODE", obj.PC_CODE)
            clsCommon.AddColumnsForChange(coll, "IsBlankCheque", obj.IsBlankCheque)
            clsCommon.AddColumnsForChange(coll, "Alies_Name", obj.Alies_Name)
            clsCommon.AddColumnsForChange(coll, "Vendor_Type_CHA", obj.Vendor_Type_CHA)
            clsCommon.AddColumnsForChange(coll, "Is_Head_Load", obj.Is_Head_Load)
            clsCommon.AddColumnsForChange(coll, "Rate_Head_Load", obj.Rate_Head_Load)
            clsCommon.AddColumnsForChange(coll, "Service_Basis_Head_Load", obj.Service_Basis_Head_Load)
            clsCommon.AddColumnsForChange(coll, "Is_Own_Asset", obj.Is_Own_Asset)
            clsCommon.AddColumnsForChange(coll, "Rate_Own_Asset", obj.Rate_Own_Asset)
            clsCommon.AddColumnsForChange(coll, "Service_Basis_Own_Asset", obj.Service_Basis_Own_Asset)
            clsCommon.AddColumnsForChange(coll, "IsVendorInvoiceNo", obj.IsVendorInvoiceNo)
            clsCommon.AddColumnsForChange(coll, "CHA_DOC_NO", obj.CHA_DOC_NO)
            clsCommon.AddColumnsForChange(coll, "Standard_Security_Amount", obj.Standard_Security_Amount)
            clsCommon.AddColumnsForChange(coll, "Is_TC_Certified", obj.Is_TC_Certified)
            clsCommon.AddColumnsForChange(coll, "TC_Certified", obj.TC_Certified)
            clsCommon.AddColumnsForChange(coll, "MP_Code", obj.MP_Code)
            clsCommon.AddColumnsForChange(coll, "MP_Name", obj.MP_Name)
            clsCommon.AddColumnsForChange(coll, "Cheque_In_Favour_Of", obj.Cheque_In_Favour_Of)
            clsCommon.AddColumnsForChange(coll, "is_Drip_Saver", obj.is_Drip_Saver)
            clsCommon.AddColumnsForChange(coll, "Other_For_PAN", obj.Other_For_PAN)
            clsCommon.AddColumnsForChange(coll, "Joint_Branch_Name", obj.Joint_Branch_Name)
            clsCommon.AddColumnsForChange(coll, "Joint_IFSC_Code", obj.Joint_IFSC_Code)
            clsCommon.AddColumnsForChange(coll, "VSP_Farmer_Billing", obj.VSP_Farmer_Billing)
            clsCommon.AddColumnsForChange(coll, "EMP_Type", obj.EMP_Type)
            clsCommon.AddColumnsForChange(coll, "EMP_Fixed_Amount", obj.EMP_Fixed_Amount)
            clsCommon.AddColumnsForChange(coll, "Actual_charges_Slab", obj.Actual_charges_Slab)
            clsCommon.AddColumnsForChange(coll, "Actual_charges_Slab2", obj.Actual_charges_Slab2)
            clsCommon.AddColumnsForChange(coll, "Actual_charges2", obj.Actual_charges2)
            clsCommon.AddColumnsForChange(coll, "Actual_charges_Slab3", obj.Actual_charges_Slab3)
            clsCommon.AddColumnsForChange(coll, "Actual_charges3", obj.Actual_charges3)
            clsCommon.AddColumnsForChange(coll, "Actual_charges_Slab4", obj.Actual_charges_Slab4)
            clsCommon.AddColumnsForChange(coll, "Actual_charges4", obj.Actual_charges4)
            clsCommon.AddColumnsForChange(coll, "Actual_charges_Slab5", obj.Actual_charges_Slab5)
            clsCommon.AddColumnsForChange(coll, "Actual_charges5", obj.Actual_charges5)
            clsCommon.AddColumnsForChange(coll, "Vendor_Distance", obj.Vendor_Distance)
            clsCommon.AddColumnsForChange(coll, "Apply_Mult_Incentive", obj.Apply_Mult_Incentive)
            clsCommon.AddColumnsForChange(coll, "MinimumQtyRequired_pertrip", obj.MinimumQtyRequired_pertrip)
            clsCommon.AddColumnsForChange(coll, "OldName", obj.OldName)
            clsCommon.AddColumnsForChange(coll, "Security_Deduction_Amount", obj.Security_Deduction_Amount)
            clsCommon.AddColumnsForChange(coll, "Interest_Per", obj.Interest_Per)
            clsCommon.AddColumnsForChange(coll, "Minimum_Interest", obj.Minimum_Interest)
            clsCommon.AddColumnsForChange(coll, "SSI_No", obj.SSI_No)
            clsCommon.AddColumnsForChange(coll, "Is_Blacklist", obj.Is_Blacklist)
            clsCommon.AddColumnsForChange(coll, "Service_Charge_Per_Unit", obj.Service_Charge_Per_Unit)
            clsCommon.AddColumnsForChange(coll, "is_Hold_Payment_Process", obj.is_Hold_Payment_Process)
            clsCommon.AddColumnsForChange(coll, "Is_Inactive_In_Milk_Procurement", obj.Is_Inactive_In_Milk_Procurement)
            clsCommon.AddColumnsForChange(coll, "DFOption", obj.DFOption)
            clsCommon.AddColumnsForChange(coll, "BusinessCondition", obj.BusinessCondition)
            clsCommon.AddColumnsForChange(coll, "GSTRegistered", obj.GSTRegistered)
            clsCommon.AddColumnsForChange(coll, "GST_Composition_scheme", obj.GST_Composition_scheme)
            clsCommon.AddColumnsForChange(coll, "GSTEntity", obj.GSTEntity)
            clsCommon.AddColumnsForChange(coll, "GSTLastEntity", obj.GSTLastEntity)
            clsCommon.AddColumnsForChange(coll, "GSTFinalNo", obj.GSTFinalNo)
            clsCommon.AddColumnsForChange(coll, "GSTMiddle", obj.GSTMiddle)
            clsCommon.AddColumnsForChange(coll, "Weight", obj.Weight)
            clsCommon.AddColumnsForChange(coll, "JWPriceCode", obj.JWPriceCode)
            clsCommon.AddColumnsForChange(coll, "CorrectionFat", obj.CorrectionFat)
            clsCommon.AddColumnsForChange(coll, "CorrectionSNF", obj.CorrectionSNF)
            clsCommon.AddColumnsForChange(coll, "IsEmployee", obj.IsEmployee)
            clsCommon.AddColumnsForChange(coll, "Handling_Charges_Per", obj.Handling_Charges_Per)
            clsCommon.AddColumnsForChange(coll, "Credit_Limit_On_Milk_Receipt_Per", obj.Credit_Limit_On_Milk_Receipt_Per)
            clsCommon.AddColumnsForChange(coll, "isBulkProcurement", obj.isBulkProcurement)
            clsCommon.AddColumnsForChange(coll, "isHighClass", obj.isHighClass)
            clsCommon.AddColumnsForChange(coll, "Bulk_ROUTE_NO", obj.Bulk_ROUTE_NO, True)
            clsCommon.AddColumnsForChange(coll, "Monthly_Rent", obj.Monthly_Rent)
            clsCommon.AddColumnsForChange(coll, "EMP_CODE", obj.EMP_CODE)
            clsCommon.AddColumnsForChange(coll, "TIP_Buffalo", obj.TIP_Buffalo)
            clsCommon.AddColumnsForChange(coll, "TIP_Cow", obj.TIP_Cow)
            clsCommon.AddColumnsForChange(coll, "TIP_Mix", obj.TIP_Mix)
            clsCommon.AddColumnsForChange(coll, "Registration_No", obj.Registration_No)
            clsCommon.AddColumnsForChange(coll, "Aadhar_No", obj.Aadhar_No)
            clsCommon.AddColumnsForChange(coll, "Care_Of", obj.Care_Of)
            clsCommon.AddColumnsForChange(coll, "IsTCSnotApplicable", obj.IsTCSnotApplicable, True)
            clsCommon.AddColumnsForChange(coll, "Isbuyerfilereturninlasttwoyears", obj.Isbuyerfilereturninlasttwoyears)
            clsCommon.AddColumnsForChange(coll, "IsTCS_TDSamountgreaterthan50KpreviousYear", obj.IsTCS_TDSamountgreaterthan50KpreviousYear)
            clsCommon.AddColumnsForChange(coll, "SecChequeNoLac1", obj.SecChequeNoLac1)
            clsCommon.AddColumnsForChange(coll, "SecChequeNoRs100", obj.SecChequeNoRs100)
            clsCommon.AddColumnsForChange(coll, "TDS_Branch_Code_Service", obj.TDS_Branch_Code_Service)
            clsCommon.AddColumnsForChange(coll, "Deduction_Code_Service", obj.Deduction_Code_Service)
            clsCommon.AddColumnsForChange(coll, "TDS_Status_Service", obj.TDS_Status_Service)
            clsCommon.AddColumnsForChange(coll, "TDS_Vendor_Type_Service", obj.TDS_Vendor_Type_Service)
            clsCommon.AddColumnsForChange(coll, "TDS_State_Code_Service", obj.TDS_State_Code_Service)
            clsCommon.AddColumnsForChange(coll, "DistanceKM_Head_Load", obj.DistanceKM_Head_Load)
            clsCommon.AddColumnsForChange(coll, "Registered_PDCS_CLUSTER", obj.Registered_PDCS_CLUSTER, True)
            If obj.StartDate IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "StartDate", clsCommon.GetPrintDate(obj.StartDate, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "StartDate", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "BankCode2", obj.BankCode2, True)
            clsCommon.AddColumnsForChange(coll, "BankName2", obj.BankName2, True)
            clsCommon.AddColumnsForChange(coll, "Credit2", obj.Credit2)
            clsCommon.AddColumnsForChange(coll, "IFSCCode2", obj.IFSCCode2, True)
            clsCommon.AddColumnsForChange(coll, "AccNo2", obj.AccNo2)
            clsCommon.AddColumnsForChange(coll, "AccountType2", obj.AccountType2)
            clsCommon.AddColumnsForChange(coll, "BankBranch2", obj.BankBranch2)
            clsCommon.AddColumnsForChange(coll, "SecurityCharges2", obj.SecurityCharges2)
            clsCommon.AddColumnsForChange(coll, "SupervisorOrRP", obj.SupervisorOrRP)
            If obj.Active_Date IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "Active_Date", clsCommon.GetPrintDate(obj.Active_Date, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "Active_Date", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "Gender", obj.Gender, True)
            clsCommon.AddColumnsForChange(coll, "RegistrationNo", obj.RegistrationNo, True)
            If obj.RegistrationDate IsNot Nothing Then
                clsCommon.AddColumnsForChange(coll, "RegistrationDate", clsCommon.GetPrintDate(obj.RegistrationDate, "dd/MMM/yyyy"))
            Else
                clsCommon.AddColumnsForChange(coll, "RegistrationDate", Nothing, True)
            End If
            clsCommon.AddColumnsForChange(coll, "DISTRICT_Code", obj.DISTRICT_Code, True)
            clsCommon.AddColumnsForChange(coll, "Zone_Code", obj.Zone_Code, True)
            clsCommon.AddColumnsForChange(coll, "Cast_Category_Code", obj.CAST_CATEGORY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Block_Code", obj.BLOCK_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Company_Bank", obj.Company_Bank)
            clsCommon.AddColumnsForChange(coll, "Revenue_Village_Code", obj.REVENUE_VILLAGE_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Grampanchayat_Code", obj.GRAMPANCHAYAT_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Panchayat_Samiti_Code", obj.PANCHAYAT_SAMITI_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Vidhan_Sabha_Code", obj.VIDHAN_SABHA_CODE, True)
            clsCommon.AddColumnsForChange(coll, "IsAllowSkipPurchaseQC", obj.IsAllowSkip_PurchaseQC)
            clsCommon.AddColumnsForChange(coll, "OEM", obj.OEM)
            clsCommon.AddColumnsForChange(coll, "Is_Provisional", obj.Is_Provisional)
            clsCommon.AddColumnsForChange(coll, "Company_Bank_Current", obj.Company_Bank_Current)
            clsCommon.AddColumnsForChange(coll, "In_Active_CF", obj.In_Active_CF)
            clsCommon.AddColumnsForChange(coll, "Is_Default_Grower", obj.Is_Default_Grower)
            clsCommon.AddColumnsForChange(coll, "Modify_By", obj.Modify_By)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Comp_Code", obj.Comp_Code)
            clsCommon.AddColumnsForChange(coll, "CST", obj.CST)
            clsCommon.AddColumnsForChange(coll, "ECC", obj.ECC)
            clsCommon.AddColumnsForChange(coll, "Range", obj.Range)
            clsCommon.AddColumnsForChange(coll, "Collectorate", obj.Collectorate)
            clsCommon.AddColumnsForChange(coll, "PAN", obj.PAN)

            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Vendor_Code", obj.Vendor_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", obj.Created_By)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_MASTER", OMInsertOrUpdate.Insert, "", trans)
                '' isSaved = isSaved AndAlso clsLocationCustomerMappings.SaveData(obj2, obj.Arr, trans)
            Else
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_MASTER", OMInsertOrUpdate.Update, "TSPL_VENDOR_MASTER.Vendor_Code='" + obj.Vendor_Code + "'", trans)
                'isSaved = isSaved AndAlso clsLocationCustomerMapping.SaveData(obj2, obj.Arr, trans)
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
            Return False
        End Try
        Return isSaved
    End Function

    Public Shared Function getQueryForVSPMaster(ByVal strVspCode As String) As String
        Dim qry As String = " Select Vendor_Name, Vendor_Group_Code,Vendor_Group_Code_Desc,Status,OnHold,Convert(Date,Closing_Date,103),Add1,Add2,Add3,City_Code,City_Code_Desc,State,Country,Phone1,Phone2,Fax,Email,WebSite,Contact_Person_Name,Contact_Person_Phone,Contact_Person_Fax,Contact_Person_Website,Contact_Person_Email,Terms_Code,Terms_Code_Desc,Vendor_Account,Vendor_Account_Desc,Payment_Code,Payment_Code_Desc,Ven_Type_Code,Ven_Type_Desc,Bank_Code,Bank_Code_Desc,Service_Tax_No,Lst_No,Tin_No,Credit_Limit,Tax_Group,Tax_Group_Desc,TAX1,TAX1_Rate,TAX2,TAX2_Rate,TAX3,TAX3_Rate,TAX4,TAX4_Rate,TAX5,TAX5_Rate,TAX6 ,TAX6_Rate ,TAX7 ,TAX7_Rate ,TAX8 ,TAX8_Rate ,TAX9 ,TAX9_Rate ,TAX10 ,TAX10_Rate ,Remarks1 ,Remarks2 ,Additional1 ,Additional2 ,Additional3,transporter,CST,ECC,Range,Collectorate,PAN,is_Gross_Receipt,Inter_branch,currency_code,franchise_yn,state_code,country_code,vsp_payment,incentive_days,incentive,commision_pers,payment_commision_pers,Service_charges,VSP_Payee_Name,Service_Charge_Type,Joint_Name,Branch_Name,Account_No,Bank_Name,IFSC_Code,Account_Type,Security_Amount,AMCU,Amc_Charge,Billing_date,Nature,Actual_charges,joint_bank_Code,Joint_Account_No,Agreement,Start_Date,End_Date,PC_Code,Is_Head_Load,Rate_Head_Load,Service_Basis_Head_Load,Is_Own_Asset,Rate_Own_Asset,Service_Basis_Own_Asset,joint_bank_code,Standard_security_Amount,MP_code,MP_Name,Cheque_In_Favour_Of,Pin_code,is_drip_saver,isnull(Joint_Branch_Name,'') as Joint_Branch_Name,isnull(Joint_IFSC_Code,'') as Joint_IFSC_Code,EMP_Type,EMP_Fixed_Amount,Actual_charges_Slab,Actual_charges_Slab2,Actual_charges2,Actual_charges_Slab3,Actual_charges3,Actual_charges_Slab4,Actual_charges4,Actual_charges_Slab5,Actual_charges5,Apply_Mult_Incentive,Security_Deduction_Amount,Interest_Per,Minimum_Interest,Is_Blacklist,Service_Charge_Per_Unit,is_Hold_Payment_Process,Is_Inactive_In_Milk_Procurement,GSTRegistered,GSTEntity,GSTLastEntity,GSTFinalNo,CorrectionFat,CorrectionSNF,Handling_Charges_Per,Credit_Limit_On_Milk_Receipt_Per,Monthly_Rent,TIP_Buffalo,TIP_Cow,TIP_Mix,Aadhar_No,Care_Of,Isbuyerfilereturninlasttwoyears,IsTCS_TDSamountgreaterthan50KpreviousYear,Is_TDS_Applicable,TDS_Branch_Code,Deduction_Code,TDS_Vendor_Type,TDS_Status,TDS_State_Code,SecChequeNoLac1,SecChequeNoRs100,DistanceKM_Head_Load,  BankCode2,BankName2,Credit2, IFSCCode2 ,AccNo2,AccountType2,BankBranch2,SecurityCharges2,Registered_PDCS_CLUSTER,StartDate,SupervisorOrRP,Vendor_name_Hindi,tspl_vendor_master.Company_Bank from tspl_vendor_master where vendor_code='" + strVspCode + "' and form_type='VSP' "
        Return qry
    End Function
    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " select Vendor_Code as [Code],Vendor_Name as [Vendor Name],ISNULL(Alies_Name,'') As [Alies Name],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [VLC Uploader Code],Add1,Add2,Add3,Closing_Date as [Closing Date],Vendor_Group_Code as [Vendor Group Code],Vendor_Group_Code_Desc as [Vendor Group Description],City_Code as [City Code],City_Code_Desc as [City Description],State,Country,Phone1,Phone2,Fax,Email,WebSite,Contact_Person_Name as [Contact Person Name],Contact_Person_Phone as [Contact Person Phone],Contact_Person_Fax as [Contact Person FAX],Contact_Person_Website as [Contact Person Website],Contact_Person_Email as [Contact Person Email],Terms_Code as [Terms Code],Terms_Code_Desc as [Terms Code Description],Vendor_Account as [Vendor Account],Vendor_Account_Desc as [Vendor Account Description],Payment_Code as [Payment Code],Payment_Code_Desc as [Payment Code Description],Bank_Code as [Bank Code],Bank_Code_Desc as [Bank Description],Tax_Group as [Tax Group],Tax_Group_Desc as [Tax Group Description],Ven_Type_Code as [Vendor Type Code],Ven_Type_Desc as [Vendor Type Description],TAX1,TAX1_Rate as [TAX1 Rate],TAX2,TAX2_Rate as [Tax2 Rate],TAX3,TAX3_Rate as [Tax3 Rate],TAX4,TAX4_Rate as [Tax4 Rate],TAX5,TAX5_Rate as [Tax5 Rate],TAX6,TAX6_Rate as [Tax6 Rate],TAX7,TAX7_Rate as [Tax7 Rate],TAX8,TAX8_Rate as [Tax8 Rate],TAX9,TAX9_Rate as [Tax9 Rate],TAX10,TAX10_Rate as [Tax10 Rate],Service_Tax_No as [Service Tax No],Tin_No as [TIN No],Lst_No as [LST No],(select case when Status='N' then 'Active' else 'In Active' end ) as Status,OnHold as [On Hold],Transporter,Remarks1,Remarks2,Additional1,Additional2,Additional3,Credit_Limit as [Credit Limit],tspl_vendor_master.Created_By as [Created By],tspl_vendor_master.Created_Date as [Created Date],Modify_By as [Modify By],Modify_Date as [Modify Date],tspl_vendor_master.Comp_Code as [Company Code],CST,ECC,Range,Collectorate,PAN,Is_Gross_Receipt as [Is Gross Receipt],Inter_Branch as [Inter Branch],CURRENCY_CODE as [Currency Code],franchise_yn as [Is Franchise] 
from tspl_vendor_master 
left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VSP_Code=tspl_vendor_master.Vendor_Code
"
        str = clsCommon.ShowSelectForm("RPTVENDFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    '----------------End of Code For Get Finder--------------------------------------------------------------'
    Public Shared Function isGrossReceipt(ByVal strVCode As String) As Boolean
        Dim qry As String = "select is_Gross_Receipt from TSPL_VENDOR_MASTER where Vendor_Code='" + strVCode + "'"
        Return IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1, True, False)
    End Function

    Public Shared Function IsAllowSkipPurchaseQC(ByVal strVCode As String) As Boolean
        Dim qry As String = "select IsAllowSkipPurchaseQC from TSPL_VENDOR_MASTER where Vendor_Code='" + strVCode + "'"
        Return IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1, True, False)
    End Function

    Public Shared Function GetName(ByVal strCode As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select Vendor_Name from TSPL_VENDOR_MASTER where Vendor_Code='" + strCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            Return ""
        End If
        Return clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
    End Function

    Public Shared Function IsGSTRegisteredVendor(ByVal strCode As String, ByVal trans As SqlTransaction) As Boolean
        If objCommonVar.TreatUnregisteredVendorAsRegisteredVendor Then
            Return True
        End If

        Dim qry As String = "select 1 from TSPL_VENDOR_MASTER where Vendor_Code='" + strCode + "' and GSTRegistered=1 "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    'done by stuti against purchase points
    Public Shared Function checkisIDSapplicable(ByVal vendorgroupcode As String, ByVal trans As SqlTransaction) As Boolean
        Dim qry As String = "select TSPL_VENDOR_group.is_TdsApplicable from TSPL_VENDOR_group where TSPL_VENDOR_group.Ven_Group_Code='" + clsCommon.myCstr(vendorgroupcode) + "'"
        Return IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) = 1, True, False)
    End Function

    Public Shared Function GetData(ByVal strCode As String, ByVal trans As SqlTransaction) As clsVendorMaster
        Dim obj As clsVendorMaster = Nothing
        Dim qry As String = "select Vendor_Name,Vendor_Account, CURRENCY_CODE from TSPL_VENDOR_MASTER where Vendor_Code='" + strCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            Throw New Exception("Vendor Code -" + strCode + ". Not Exist")
        End If
        obj = New clsVendorMaster()
        obj.Vendor_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
        obj.Vendor_Account = clsCommon.myCstr(dt.Rows(0)("Vendor_Account"))
        Return obj

    End Function

    Public Shared Function Save_VSP_Data(ByVal strDocNo As String, ByVal Arr As List(Of clsVendorMaster), ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            'Dim sQuery As String = "delete from TSPL_MCC_VSP_ChargeCategory_MAPPING where VSP_Code='" & strDocNo & "'"
            'clsDBFuncationality.ExecuteNonQuery(sQuery, trans)
            For Each obj As clsVendorMaster In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "VSP_CODE", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Charge_CODE", obj.Charge_CODE)
                clsCommon.AddColumnsForChange(coll, "Updated_date", clsCommon.GetPrintDate(obj.Updated_date, "dd-MMM-yyyy hh:mm:ss tt"))
                clsCommon.AddColumnsForChange(coll, "Rate", obj.Rate)
                Dim check As Integer = clsDBFuncationality.getSingleValue("select count(*) from TSPL_MCC_VSP_ChargeCategory_MAPPING where vsp_code='" & strDocNo & "' and Charge_CODE='" & obj.Charge_CODE & "'")
                If check <= 0 Then
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_MCC_VSP_ChargeCategory_MAPPING", OMInsertOrUpdate.Insert, "TSPL_MCC_VSP_ChargeCategory_MAPPING.VSP_CODE='" + strDocNo + "'", trans)
                Else
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_MCC_VSP_ChargeCategory_MAPPING", OMInsertOrUpdate.Update, "TSPL_MCC_VSP_ChargeCategory_MAPPING.VSP_CODE='" + strDocNo + "'  and Charge_CODE='" & obj.Charge_CODE & "'", trans)
                End If
            Next
        End If
        Return True
    End Function

    Public Shared Function GetChargesData(ByVal strCode As String, ByVal NavType As NavigatorType) As List(Of clsVendorMaster)
        Try
            Dim whrcls As String = ""
            Dim objtr As New clsVendorMaster

            whrcls = " and cm.VSP_code ='" & strCode & "'"
            Dim qry As String = "select Charge_CODE,Description,GL_CODE,GL_DESC,RaTE from TSPL_MCC_VSP_ChargeCategory_MAPPING cm inner join " _
            & " TSPL_Charge_Category cc on cc.Charge_Cat_Code=Charge_Code " & whrcls & ""

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            Dim obj As New List(Of clsVendorMaster)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                For Each dr As DataRow In dt.Rows
                    objtr = New clsVendorMaster

                    objtr.VSP_CODE = strCode
                    objtr.Charge_CODE = clsCommon.myCstr(dr("Charge_CODE"))
                    objtr.Charge_Desc = clsCommon.myCstr(dr("Description"))
                    objtr.GL_CODE = clsCommon.myCstr(dr("GL_CODE"))
                    objtr.GL_DESC = clsCommon.myCstr(dr("GL_DESC"))
                    objtr.Rate = clsCommon.myCstr(dr("Rate"))
                    obj.Add(objtr)
                Next
            End If
            Return obj
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function GetVendorLedgerBaseQry(ByVal strPortrait As Boolean, ByVal strLandscape As Boolean, ByVal IsOnlyForAgainstSalary As Boolean, ByVal strfromdate As String, ByVal strtodate As String, ByVal strvendor As String, ByVal isOpening As Boolean, ByVal strCurrencyType As String, ByVal FormType As String, ByVal isVendorWise As Boolean, ByVal isVendorGroupWise As Boolean, ByVal isDocWise As Boolean, ByVal isIncludeApplyDocument As Boolean) As String
        Dim strtempBaseQryforopening As String = String.Empty
        Dim strtempBaseQryforopeningForMIS As String = String.Empty
        Dim strtempBaseQry As String = String.Empty

        strtempBaseQryforopening = GetVendorLedgerBaseQryForOpening(strPortrait, strLandscape, IsOnlyForAgainstSalary, strfromdate, strtodate, strvendor, isOpening, strCurrencyType, FormType, isVendorWise, isVendorGroupWise, isDocWise, isIncludeApplyDocument)

        If isDocWise = True Then

            If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine &
                    " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
                    "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
                    "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
                    "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine &
                    " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) else  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
                    " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
                    "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end else " + Environment.NewLine &
                    "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
                    " case when InnQuery.DocType='Pur.Invoice' then InnQuery.CrAmt-InnQuery.DrAmt else 0 end as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine &
                    "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_M1ASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code  "
            Else


                strtempBaseQry = "  Select isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) as EXCHANGE_GAIN_AMT,isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) as EXCHANGE_LOSS_AMT, InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) - case when (DocType)='Reverse Payment' then  (Select isnull((PH.EXCHANGE_LOSS_AMT  ),0) from TSPL_PAYMENT_HEADER PH where PH.Payment_No =(  Select Document_No   from TSPL_BANK_REVERSE where Reverse_Code =InnQuery.DocNo )) else 0 end as  CrAmt ," + Environment.NewLine &
                   " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
                   "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
                   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
                   "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine &
                   " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) ELSE  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
                   " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
                   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then CASE WHEN ((DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")-isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0))<0 THEN 0 ELSE (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")-isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) END  else CASE WHEN (DocType)<>'IM' THEN (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) ELSE (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") END end else " + Environment.NewLine &
                   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
                   " InnQuery.CrAmt-InnQuery.DrAmt  as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine &
                   "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code  "


            End If

        Else
            ''richa to exclude exchange gain loss documnets for only apply documnets when include apply doc.checkbox is off  KDI/14/06/2018-000364
            Dim strExcludeEXcforApplyDocumnets As String = " where InnQuery.DocNo not in ( Select Payment_No  from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Payment_Type  ='AD' and   TSPL_PAYMENT_HEADER.Posted='1' and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT >0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT >0)  "
            If isOpening = True Then
                strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <'" + strfromdate + "'  " + Environment.NewLine
            Else
                strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' " + Environment.NewLine
            End If
            If clsCommon.myLen(strvendor) > 0 Then
                strExcludeEXcforApplyDocumnets += " and TSPL_PAYMENT_HEADER.Vendor_Code in (" & strvendor & ")"
            End If
            strExcludeEXcforApplyDocumnets += Environment.NewLine & " Union All" & Environment.NewLine &
                " Select Reverse_Code  from TSPL_BANK_REVERSE where Document_No in ( Select Payment_No  from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Payment_Type  ='AD' and   TSPL_PAYMENT_HEADER.Posted='1' and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT >0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT >0)  "
            If isOpening = True Then
                strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <'" + strfromdate + "'  " + Environment.NewLine
            Else
                strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' " + Environment.NewLine
            End If
            If clsCommon.myLen(strvendor) > 0 Then
                strExcludeEXcforApplyDocumnets += " and TSPL_PAYMENT_HEADER.Vendor_Code in (" & strvendor & ")"
            End If

            strExcludeEXcforApplyDocumnets += " ) ) "

            '----------------------


            strtempBaseQry = "  Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  "
            If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                strtempBaseQry += " CONVERT(DECIMAL(18,2), CASE WHEN InnQuery.DocType NOT IN ('Pur.Invoice','Receipt')  THEN case when InnQuery.DocType not in ('TDS','AP Invoice') THEN InnQuery.CrAmt  *(case when (DocType) NOT IN ('EXC','Credit Note','IM','Reverse Payment') then   InnQuery.ConvRate else 1 end) when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo)='C' then InnQuery.CrAmt ELSE CASE WHEN InnQuery.DocType in ('TDS') and (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  )='I' or (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo   )='C' then (Select ISNULL (Actual_Total_TDS ,0)  from TSPL_REMITTANCE where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  ) * -1 else 0 end END ELSE 0 END) as CrAmt ," &
                      " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Adjustment') then case when InnQuery.DocType not in ('AP Invoice','TDS') then InnQuery.DrAmt WHEN InnQuery.DocType in ('TDS') and   (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  )='D' then (Select ISNULL (Actual_Total_TDS ,0)  from TSPL_REMITTANCE where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  ) * -1 else case when  InnQuery.DocType in ('AP Invoice') AND  (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')='')='D' then InnQuery.DrAmt else 0 end end else 0 end) DrAmt, " &
                      " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Debit Note','Reverse Payment','Credit Note','EXC','IM') then case when InnQuery.DocType not in ('AP Invoice','TDS') then CONVERT(DECIMAL(18,2),InnQuery.Purchase) else case when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo)='I'   then CONVERT(DECIMAL(18,2),InnQuery.Purchase) when  InnQuery.DocType in ('AP Invoice')  and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')<>'')='D' then CONVERT(DECIMAL(18,2),InnQuery.Purchase) *-1  else 0 end end else 0 end) Purchase, " &
                      " CONVERT(DECIMAL(18,2), case when InnQuery.DocType  in ('Payment','On Account','Advance','Receipt','TDS','EXC','IM') then InnQuery.Payments * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) else 0 end) as Payments , "

            Else
                strtempBaseQry += " CONVERT(DECIMAL(18,2),InnQuery.CrAmt) AS CrAmt,CONVERT(DECIMAL(18,2),InnQuery.DrAmt) AS DrAmt,CONVERT(DECIMAL(18,2),InnQuery.Purchase) AS Purchase,CONVERT(DECIMAL(18,2),InnQuery.Payments) AS Payments , "
            End If

            strtempBaseQry += " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code from " &
               "(Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end)  as  CrAmt ," + Environment.NewLine &
               " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
               "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
               "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
               "case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine &
               "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine &
               "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
               "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
               "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
               " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine &
               "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code " & IIf(isIncludeApplyDocument = False, " " & strExcludeEXcforApplyDocumnets & " ", "") & "  ) InnQuery left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =InnQuery.VCode "

            ''richa KDI/15/05/19-000453,KDI/15/05/19-000452
            If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then

                strtempBaseQry = "  Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode,TSPL_VENDOR_MASTER.Vendor_Name as VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  " &
                  "  CONVERT(DECIMAL(18,2),InnQuery.CrAmt) AS CrAmt,CONVERT(DECIMAL(18,2),InnQuery.DrAmt) AS DrAmt,0 as Purchase,0 as Payments, " &
                  " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code from " &
                  " (Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine &
                  " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
                  " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
                  " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
                  " case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine &
                  " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine &
                  " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
                  "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
                  "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
                  " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine &
                  "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code " & IIf(isIncludeApplyDocument = False, " " & strExcludeEXcforApplyDocumnets & " ", "") & " ) InnQuery left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =InnQuery.VCode "


            End If

        End If
        Return strtempBaseQry
    End Function
    Public Shared Function GetVendorLedgerBaseQryForOpening(ByVal strPortrait As Boolean, ByVal strLandscape As Boolean, ByVal IsOnlyForAgainstSalary As Boolean, ByVal strfromdate As String, ByVal strtodate As String, ByVal strvendor As String, ByVal isOpening As Boolean, ByVal strCurrencyType As String, ByVal FormType As String, ByVal isVendorWise As Boolean, ByVal isVendorGroupWise As Boolean, ByVal isDocWise As Boolean, ByVal isIncludeApplyDocument As Boolean) As String
        Dim strtempBaseQry As String = String.Empty


        Try

            Dim strShowReferenceDocNoofAPInvoice As String = "   (case when isnull(TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ,'')<>'' then TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No when isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No  ,'')<>'' then TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No when isnull(TSPL_VENDOR_INVOICE_HEAD.Against_BulkMillkPurchaseInvoice_No  ,'')<>'' then TSPL_VENDOR_INVOICE_HEAD.Against_BulkMillkPurchaseInvoice_No when isnull(TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No  ,'')<>'' then TSPL_VENDOR_INVOICE_HEAD.Against_MillkPurchaseInvoice_No when isnull(TSPL_VENDOR_INVOICE_HEAD.Against_VCGL   ,'')<>'' then TSPL_VENDOR_INVOICE_HEAD.Against_VCGL  else '' end )as Reference_Doc_No,"

            Dim strTaxRecovarableQuery As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then " &
       " ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'  " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end + " &
       " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y' " &
       " then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end) "

            Dim strQryForRMDA As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and LEN(ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,''))>0 then (Select sum(TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt) from TSPL_TRANSFER_ORDER_HEAD where Status=1 and TSPL_TRANSFER_ORDER_HEAD.RMDA_Code in (select TSPL_SRN_HEAD.RMDA_No from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_PI_HEAD.Against_SRN where TSPL_PI_HEAD.PI_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No  )) else 0 end )"
            strtempBaseQry = " select " & strShowReferenceDocNoofAPInvoice & " ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,TSPL_VENDOR_INVOICE_HEAD.Due_Date, '' AS PO_SRN, TSPL_VENDOR_INVOICE_HEAD.vendor_code as VCode,TSPL_VENDOR_INVOICE_HEAD.vendor_name as VName,TSPL_VENDOR_INVOICE_HEAD .Document_No as DocNo ,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'Debit Note' else case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'Credit Note' else 'AP Invoice' end  end as DocType,convert(date,Invoice_Entry_Date,103) as DocDate,((Description) + (case when Description='' then case when TSPL_VENDOR_INVOICE_HEAD.Remarks <>'' then TSPL_VENDOR_INVOICE_HEAD.Remarks +' - ' else ''  end else ' - 'end) +(RefDocNo)+'Vendor Invoice No-' +Vendor_Invoice_No+' Date-' +Vendor_Invoice_Date  ) as DocNarr,'' as ChequeDetails, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE, TSPL_VENDOR_INVOICE_HEAD.ConvRate,"

            strtempBaseQry += " case when TSPL_VENDOR_INVOICE_HEAD.Document_Type IN('I','C') Then document_total " + strTaxRecovarableQuery + " else 0 end + ( case when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX1 ,'')='WCT' THEN TAX1_Amt * -1  when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2 ,'')='WCT' THEN TAX2_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3 ,'')='WCT' THEN TAX3_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4 ,'')='WCT' THEN TAX4_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5 ,'')='WCT' THEN TAX5_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6 ,'')='WCT' THEN TAX6_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7 ,'')='WCT' THEN TAX7_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,'')='WCT' THEN TAX8_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9 ,'')='WCT' THEN TAX9_Amt * -1 when isnull(TSPL_VENDOR_INVOICE_HEAD.TAX10 ,'')='WCT' THEN TAX10_Amt * -1 ELSE 0 END) as CrAmt, " &
              "  case when TSPL_VENDOR_INVOICE_HEAD.Document_Type IN('D') then Document_Total " + strTaxRecovarableQuery + " else 0 end as DrAmt, Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' Then document_total " + strTaxRecovarableQuery + " Else 0 End as Purchase, 0 as Payments, Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then document_total " + strTaxRecovarableQuery + " Else 0 End as DrNote, Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then document_total " + strTaxRecovarableQuery + " Else 0 End as CrNote, TSPL_VENDOR_INVOICE_HEAD.Document_Type,balance_Amt as Balance_Amount, ( select   GL_Account_Code from TSPL_VENDOR_INVOICE_DETAIL where Detail_Line_No='1' and Document_No=tspl_vendor_invoice_head.Document_No  ) as account, convert(date,tspl_vendor_invoice_head.Posting_Date,103) as Posting_Date ,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'AP-DN' when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, tspl_vendor_invoice_head.Comp_Code from tspl_vendor_invoice_head where ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<>''  and LEN(ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,''))<=0   " + Environment.NewLine



            If isOpening = True Then
                strtempBaseQry += "  and  convert(date,Invoice_Entry_Date,103)  <'" + strfromdate + "'  " + Environment.NewLine
            Else
                strtempBaseQry += "  and  convert(date,Invoice_Entry_Date,103)  >='" + strfromdate + "' and  convert(date,Invoice_Entry_Date,103)  <='" + strtodate + "' " + Environment.NewLine
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_VENDOR_INVOICE_HEAD.vendor_code in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine
            ''------------- code inserted for WCT type of documents
            strtempBaseQry += " Select * from ( select " & strShowReferenceDocNoofAPInvoice & "  ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,TSPL_VENDOR_INVOICE_HEAD.Due_Date, '' AS PO_SRN, vendor_code as VCode,vendor_name as VName,TSPL_VENDOR_INVOICE_HEAD .Document_No as DocNo ,'WCT' as DocType,convert(date,Invoice_Entry_Date,103) as DocDate,((Description) + (case when Description='' then case when TSPL_VENDOR_INVOICE_HEAD.Remarks <>'' then TSPL_VENDOR_INVOICE_HEAD.Remarks +' - ' else ''  end else ' - 'end) +(RefDocNo)+'Vendor Invoice No-' +Vendor_Invoice_No+' Date-' +Vendor_Invoice_Date  ) as DocNarr,'' as ChequeDetails, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE, TSPL_VENDOR_INVOICE_HEAD.ConvRate," &
            " 0 as CrAmt,case when isnull(TAX1 ,'')='WCT' THEN TAX1_Amt * -1  when isnull(TAX2 ,'')='WCT' THEN TAX2_Amt * -1 when isnull(TAX3 ,'')='WCT' THEN TAX3_Amt * -1 when isnull(TAX4 ,'')='WCT' THEN TAX4_Amt * -1 when isnull(TAX5 ,'')='WCT' THEN TAX5_Amt * -1 when isnull(TAX6 ,'')='WCT' THEN TAX6_Amt * -1 when isnull(TAX7 ,'')='WCT' THEN TAX7_Amt * -1 when isnull(TAX8 ,'')='WCT' THEN TAX8_Amt * -1 when isnull(TAX9 ,'')='WCT' THEN TAX9_Amt * -1 when isnull(TAX10 ,'')='WCT' THEN TAX10_Amt * -1 ELSE 0 END  DrAmt, " &
            " Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' Then document_total Else 0 End as Purchase, 0 as Payments, Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then document_total Else 0 End as DrNote, Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then document_total Else 0 End as CrNote, TSPL_VENDOR_INVOICE_HEAD.Document_Type,balance_Amt as Balance_Amount, ( select   GL_Account_Code from TSPL_VENDOR_INVOICE_DETAIL where Detail_Line_No='1' and Document_No=tspl_vendor_invoice_head.Document_No  ) as account, convert(date,tspl_vendor_invoice_head.Posting_Date,103) as Posting_Date ,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'AP-DN' when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, tspl_vendor_invoice_head.Comp_Code from tspl_vendor_invoice_head where ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<>''  and LEN(ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,''))<=0  ) FinalWCt where (FinalWCt.CrAmt <>0 or FinalWCt.DrAmt <>0) " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,FinalWCt.DocDate ,103) <'" + strfromdate + "'  "
            Else
                strtempBaseQry += " and  convert(date,FinalWCt.DocDate ,103)  >='" + strfromdate + "' and  convert(date,FinalWCt.DocDate ,103)  <='" + strtodate + "' "
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and FinalWCt.VCode in (" & strvendor & ")"
            End If
            strtempBaseQry += "  UNION ALL" + Environment.NewLine
            ''---------------------------

            strtempBaseQry += " Select  Reference_Doc_No,Emp_type,Emp_Adv_Type,Due_Date, PO_SRN, VCode, VName, DocNo, DocType, DocDate, DocNarr, ChequeDetails, Currency_Code, ConvRate, CrAmt, DrAmt, "
            If isVendorWise Or isVendorGroupWise Then
                strtempBaseQry += "case when DocType='Pur.Invoice' then  CrAmt-DrAmt else 0 end as Purchase,"
            Else
                strtempBaseQry += " CrAmt-DrAmt as Purchase,"
            End If
            strtempBaseQry += " 0 as Payments, 0 as drNote, 0 as CrNote, Document_Type, Balance_Amount, account, Posting_Date, GLDocType, Comp_Code from (" + Environment.NewLine &
            " select " & strShowReferenceDocNoofAPInvoice & "  ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,TSPL_VENDOR_INVOICE_HEAD.Due_Date ,CASE WHEN isnull(TSPL_PI_HEAD.Against_PO,'') <> '' THEN isnull(TSPL_PI_HEAD.Against_PO,'')  ELSE ''  + isnull(TSPL_PI_HEAD.Against_SRN,'') END  AS PO_SRN ,"
            If strPortrait = True Then

                strtempBaseQry += " TSPL_PI_HEAD.Vendor_Code as VCode,TSPL_PI_HEAD.Vendor_Name as VName, TSPL_VENDOR_INVOICE_HEAD.Document_No as DocNo,'Pur.Invoice' as DocType,convert(date,TSPL_PI_HEAD.PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' - '+'SRN No-'+TSPL_PI_HEAD.Against_SRN+ ' ,Vendor Invoice No-'+ TSPL_PI_HEAD.Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date)+ (SELECT ' & Items Are- '+STUFF((SELECT ', ' + t2.Item_Desc+'('+Convert(Varchar,COnvert(Decimal(18,2),t2.Item_Cost))+' Rs)' FROM TSPL_PI_DETAIL t2 WHERE t2.PI_No = TSPL_PI_HEAD.PI_No FOR XML PATH ('')),1,2,'')) as DocNarr,'' as ChequeDetails, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE, TSPL_VENDOR_INVOICE_HEAD.ConvRate, (case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " else 0 end) as CrAmt, (case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " else 0 end ) as DrAmt, TSPL_VENDOR_INVOICE_HEAD.Document_Type, TSPL_VENDOR_INVOICE_HEAD.Balance_Amt" + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=TSPL_PI_HEAD.Bill_To_Location) as account, convert(date,TSPL_PI_HEAD.Posting_Date,103) as Posting_Date ,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'AP-DN' when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, TSPL_PI_HEAD.Comp_Code from TSPL_PI_HEAD  left outer join TSPL_PJV_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PJV_HEAD.Invoice_No left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No=TSPL_PI_HEAD.Vendor_Invoice_No and TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No=TSPL_PI_HEAD.PI_No where TSPL_PI_HEAD.Status=1" + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += " and  convert(date,TSPL_PI_HEAD.PI_Date ,103)  <'" + strfromdate + "'"
                Else
                    strtempBaseQry += " and  convert(date,TSPL_PI_HEAD.PI_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PI_HEAD.PI_Date ,103)  <='" + strtodate + "'"
                End If


                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_PI_HEAD.Vendor_Code in (" & strvendor & ")"
                End If
            ElseIf strLandscape = True Then
                strtempBaseQry += " TSPL_PI_HEAD.Vendor_Code as VCode,TSPL_PI_HEAD.Vendor_Name as VName, TSPL_VENDOR_INVOICE_HEAD.Document_No as DocNo,'Pur.Invoice' as DocType,convert(date,TSPL_PI_HEAD.PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' ,Vendor Invoice No-'+ TSPL_PI_HEAD.Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date) as DocNarr,'' as ChequeDetails, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE, TSPL_VENDOR_INVOICE_HEAD.ConvRate, (case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " else 0 end) as CrAmt,  (case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " else 0 end ) as DrAmt, TSPL_VENDOR_INVOICE_HEAD.Document_Type, TSPL_VENDOR_INVOICE_HEAD.Balance_Amt " + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=TSPL_PI_HEAD.Bill_To_Location) as account, convert(date,TSPL_PI_HEAD.Posting_Date,103) as Posting_Date ,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'AP-DN' when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, TSPL_PI_HEAD.Comp_Code from TSPL_PI_HEAD  left outer join TSPL_PJV_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PJV_HEAD.Invoice_No left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No=TSPL_PI_HEAD.Vendor_Invoice_No and TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No=TSPL_PI_HEAD.PI_No where TSPL_PI_HEAD.Status=1" + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += " and  convert(date,TSPL_PI_HEAD.PI_Date ,103) <'" + strfromdate + "' "
                Else
                    strtempBaseQry += " and  convert(date,TSPL_PI_HEAD.PI_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PI_HEAD.PI_Date ,103)  <='" + strtodate + "'"
                End If

                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_PI_HEAD.Vendor_Code in (" & strvendor & ")"
                End If
            End If
            strtempBaseQry += ") XXX"
            strtempBaseQry += " UNION ALL" + Environment.NewLine

            strtempBaseQry += " select " & strShowReferenceDocNoofAPInvoice & " ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,TSPL_VENDOR_INVOICE_HEAD.Due_Date, '' AS PO_SRN, TSPL_REMITTANCE.Vendor_Code, TSPL_REMITTANCE.Vendor_Name, TSPL_REMITTANCE.Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, Case When TSPL_REMITTANCE.Document_Type IN ('I','C','D') Then TSPL_VENDOR_INVOICE_HEAD.Description Else TSPL_PAYMENT_HEADER.Entry_Desc End as DocNarr, '', Case When TSPL_REMITTANCE.Document_Type in ('I','D','C') Then TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE Else TSPL_PAYMENT_HEADER.CURRENCY_CODE End as Currency_Code, Case When TSPL_REMITTANCE.Document_Type in ('I','D','C') Then TSPL_VENDOR_INVOICE_HEAD.ConvRate Else TSPL_PAYMENT_HEADER.ConvRate End as ConvRate, case when TSPL_REMITTANCE.Document_Type IN('I','C','OA','AV') AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' then 0 else Actual_Total_TDS END as CrAmt, case when TSPL_REMITTANCE.Document_Type IN('I','C','OA','AV') AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' then "

            strtempBaseQry += " Actual_Total_TDS "

            strtempBaseQry += " else 0 END, case when TSPL_REMITTANCE.Document_Type ='I' AND TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as Purchase, case when TSPL_REMITTANCE.Document_Type IN "
            If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                strtempBaseQry += "('I','D','C')"
            Else
                strtempBaseQry += "('I','D')"
            End If
            strtempBaseQry += " AND TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount >0 then 0 else Actual_Total_TDS END as Payments, case when TSPL_REMITTANCE.Document_Type ='D' AND TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as DrNotes, 0 as CrNotes, 'TDS',0, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right( Branch_GL_AC,3) End as account, Case When ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date,'')<>'' Then TSPL_VENDOR_INVOICE_HEAD.Posting_Date Else TSPL_PAYMENT_HEADER.Payment_Post_Date End as Posting_Date,case when TSPL_REMITTANCE.Document_Type in ('C') then 'AP-CN' when TSPL_REMITTANCE.Document_Type in ('I') then 'AP-IN' when TSPL_REMITTANCE.Document_Type in ('D') then 'AP-DN' else case when TSPL_REMITTANCE.Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end end as GLDocType, TSPL_REMITTANCE.Comp_Code from TSPL_REMITTANCE Left Outer Join TSPL_VENDOR_INVOICE_HEAD ON TSPL_REMITTANCE.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No=TSPL_REMITTANCE.Document_No where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.is_For_TDS,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1   " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,TSPL_REMITTANCE.Document_Date  ,103)  <'" + strfromdate + "' "
            Else
                strtempBaseQry += " and  convert(date,TSPL_REMITTANCE.Document_Date  ,103)  >='" + strfromdate + "' and  convert(date,TSPL_REMITTANCE.Document_Date ,103)  <='" + strtodate + "'"
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_REMITTANCE.Vendor_Code in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine &
            " select '' as Reference_Doc_No, ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_REMITTANCE.Vendor_Code, TSPL_REMITTANCE.Vendor_Name ,TSPL_REMITTANCE.Document_No ,'TDS REVERSE' as [DocType],convert(date,TSPL_BANK_REVERSE.Reversal_Date,103)as Document_Date,'', '', TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when (TSPL_PAYMENT_HEADER.Payment_No=TSPL_BANK_REVERSE.Document_No) then Actual_Total_TDS else null end AS CrAmt, 0 as DrAmt, 0 as Purchase, case when (TSPL_PAYMENT_HEADER.Payment_No=TSPL_BANK_REVERSE.Document_No) then Actual_Total_TDS*-1 else 0 end as Payments, 0 as DrNote, 0 as CrNote, 'TDS',0, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right( Branch_GL_AC,3) End as account, TSPL_PAYMENT_HEADER.Payment_Date as Posting_Date,'RV-TA' as GLDocType, TSPL_REMITTANCE.Comp_Code from TSPL_REMITTANCE left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No=TSPL_REMITTANCE.Document_No inner join TSPL_BANK_REVERSE on TSPL_PAYMENT_HEADER.Payment_No=TSPL_BANK_REVERSE.Document_No where Remit_TDS is not null and TSPL_BANK_REVERSE.Post ='P' and Branch_GL_AC  is not null and Actual_Total_TDS<>0 AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1   AND 1=2" + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += "  and  convert(date,TSPL_REMITTANCE.Document_Date  ,103)  <'" + strfromdate + "' "
            Else
                strtempBaseQry += "  and  convert(date,TSPL_REMITTANCE.Document_Date  ,103)  >='" + strfromdate + "' and  convert(date,TSPL_REMITTANCE.Document_Date ,103)  <='" + strtodate + "' "
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_REMITTANCE.Vendor_Code in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine
            If objCommonVar.IsMultiCurrencyCompany = True Then
                strtempBaseQry += "select '' as Reference_Doc_No, ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_BANK_REVERSE.vendor_code as VCode, TSPL_BANK_REVERSE.vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, TSPL_BANK_REVERSE.Document_No+case when TSPL_PAYMENT_HEADER.Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(TSPL_BANK_REVERSE.cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate,"

                strtempBaseQry += " case when TSPL_PAYMENT_HEADER.Payment_Type IN ('RC','AD') then 0 when TSPL_PAYMENT_HEADER.Payment_Type='PY' then case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then  TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else  TSPL_PAYMENT_DETAIL.Applied_Amount  end ELSE TSPL_BANK_REVERSE.Amount end as CrAmt, " &
           "case when TSPL_PAYMENT_HEADER.Payment_Type IN ('AD') then case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then  TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else  TSPL_PAYMENT_DETAIL.Applied_Amount  end WHEN TSPL_PAYMENT_HEADER.Payment_Type IN ('RC') then TSPL_BANK_REVERSE.Amount else 0 end  as DrAmt, 0 as Purchase, case when TSPL_PAYMENT_HEADER.Payment_Type IN('OA','AV') then -1*TSPL_BANK_REVERSE.Amount when TSPL_PAYMENT_HEADER.Payment_Type='PY' then -1*TSPL_PAYMENT_DETAIL.Applied_Amount ELSE TSPL_BANK_REVERSE.Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code When TSPL_PAYMENT_HEADER.Payment_Type in ('PY') Then case when isnull(TSPL_VENDOR_INVOICE_HEAD.Loc_Code,'')='' then TSPL_PAYMENT_HEADER.Location_GL_Code else TSPL_VENDOR_INVOICE_HEAD.Loc_Code end Else CASE WHEN TSPL_PAYMENT_HEADER.Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM TSPL_PAYMENT_HEADER PR WHERE PR.Payment_No =TSPL_PAYMENT_HEADER.Applied_Payment )  ELSE substring(TSPL_PAYMENT_HEADER.Debit_Account , len(TSPL_PAYMENT_HEADER.Debit_Account )-2,3) END End as account, TSPL_BANK_REVERSE.Reversal_Date as Posting_Date,'RV-TA' as GLDocType, TSPL_BANK_REVERSE.Comp_Code from TSPL_BANK_REVERSE" &
           " LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No = TSPL_BANK_REVERSE.Document_No" &
           " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No" &
           " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No" &
           " left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_REVERSE.Bank_Code" &
           " where Source_Type ='AP' And TSPL_BANK_REVERSE.Post='P' AND TSPL_BANK_REVERSE.Vendor_Code<> '' AND TSPL_PAYMENT_HEADER.IsChkReverse='Y' AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1  " & IIf(isIncludeApplyDocument = False, " and TSPL_PAYMENT_HEADER.Payment_Type<>'AD' ", "") & " " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += " and  convert(date,TSPL_BANK_REVERSE.Reversal_Date ,103) <'" + strfromdate + "'  "
                Else
                    strtempBaseQry += " and  convert(date,TSPL_BANK_REVERSE.Reversal_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_BANK_REVERSE.Reversal_Date  ,103)  <='" + strtodate + "' "
                End If

                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_BANK_REVERSE.vendor_code in (" & strvendor & ")"
                End If
                strtempBaseQry += " ----------- For bank reverse entry--------------- " + Environment.NewLine
            Else
                strtempBaseQry += " select '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_BANK_REVERSE.vendor_code as VCode, TSPL_BANK_REVERSE.vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, TSPL_BANK_REVERSE.Document_No as DocNarr,(TSPL_BANK_REVERSE.cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when   TSPL_PAYMENT_HEADER.Payment_Type='RC' then 0 ELSE TSPL_BANK_REVERSE.Amount end as CrAmt, case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then TSPL_BANK_REVERSE.Amount else 0 end  as DrAmt, 0 as Purchase, case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then TSPL_BANK_REVERSE.Amount ELSE -1*TSPL_BANK_REVERSE.Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right(BANKACC,3) End as account, TSPL_BANK_REVERSE.Reversal_Date as Posting_Date,'RV-TA' as GLDocType, TSPL_BANK_REVERSE.Comp_Code from TSPL_BANK_REVERSE  LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No = TSPL_BANK_REVERSE.Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_REVERSE.Bank_Code where Source_Type ='AP' And Post='P' AND TSPL_BANK_REVERSE.Vendor_Code<> '' AND TSPL_PAYMENT_HEADER.IsChkReverse='Y' AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1   " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += " and  convert(date,TSPL_BANK_REVERSE.Reversal_Date ,103)  <'" + strfromdate + "'  "
                Else
                    strtempBaseQry += " and  convert(date,TSPL_BANK_REVERSE.Reversal_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_BANK_REVERSE.Reversal_Date  ,103)  <='" + strtodate + "' "
                End If

                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_BANK_REVERSE.vendor_code in (" & strvendor & ")"
                End If
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine &
            " select   " & strShowReferenceDocNoofAPInvoice & " ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, VC_Code as VCode, VC_Name as VName, TSPL_VCGL_Head.Document_No as DocNo,case when TSPL_VCGL_Head.Document_Type='V' then 'Vendor' else '' end as DocType,CONVERT(Date,TSPL_VCGL_Head.Document_Date,103) as DocDate, TSPL_VCGL_Head.Remarks as DocNarr, '' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, (case when Amount_Type='Dr' then Amount else 0 end) as CrAmt,(case when Amount_Type='Cr' then Amount else 0 end) as DrAmt, 0 as Purchase, 0 as Payments, case when Amount_Type='Cr' then Amount else 0 end as DrNote, case when Amount_Type='Dr' then Amount else 0 end as CrNote, (case when  TSPL_VCGL_Head.Document_Type='V' then 'Vendor' else '' end ) as Document_Type ,0 as Balance_Amount, (TSPL_VCGL_Head.Location_Segment) as account, TSPL_VCGL_Head.Posting_Date as Posting_Date,'VC-GL' as GLDocType,TSPL_VCGL_Head.Comp_Code from TSPL_VCGL_Head LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_VCGL=TSPL_VCGL_Head.Document_No where TSPL_VCGL_Head.Document_Type='v' and TSPL_VCGL_Head.Status='1' AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_VCGL,'')=''  " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,TSPL_VCGL_Head.Document_Date ,103)  <'" + strfromdate + "'"
            Else
                strtempBaseQry += " and  convert(date,TSPL_VCGL_Head.Document_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_VCGL_Head.Document_Date,103)  <='" + strtodate + "'"
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and VC_Code in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine &
            " select    " & strShowReferenceDocNoofAPInvoice & " ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,NULL AS Due_Date, '' AS PO_SRN, TSPL_VCGL_Detail.VCGL_Code as VCode, TSPL_VCGL_Detail.VCGL_Name as VName, TSPL_VCGL_Detail.Document_No as DocNo, TSPL_VCGL_Detail.Row_Type as DocType,CONVERT(date,TSPL_VCGL_Head.Document_Date,103) as DocDate,TSPL_VCGL_Head.Remarks as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, TSPL_VCGL_Detail.Cr_Amount as CrAmt,TSPL_VCGL_Detail.Dr_Amount as DrAmt, 0 as Purchase, 0 as Payments, TSPL_VCGL_Detail.Dr_Amount as DrNote, TSPL_VCGL_Detail.Cr_Amount as CrNote, TSPL_VCGL_Detail.Row_Type as Document_Type,0 as Balance_Amount, (TSPL_VCGL_Head.Location_Segment) as account, TSPL_VCGL_Head.Posting_Date as Posting_Date ,'VC-GL' as GLDocType, TSPL_VCGL_Head.Comp_Code from  TSPL_VCGL_Detail left outer join  TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No=TSPL_VCGL_Detail.Document_No LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_VCGL=TSPL_VCGL_Head.Document_No where Row_Type='Vendor'  and TSPL_VCGL_Head.Status='1' AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_VCGL,'')=''" + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,TSPL_VCGL_Head.Document_Date ,103)  <'" + strfromdate + "'"
            Else
                strtempBaseQry += " and  convert(date,TSPL_VCGL_Head.Document_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_VCGL_Head.Document_Date,103)  <='" + strtodate + "'"
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and VC_Code in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine
            strtempBaseQry += "  select    " & strShowReferenceDocNoofAPInvoice & "  ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,NULL AS Due_Date, '' AS PO_SRN, TSPL_Payment_Adjustment_Header.Vendor_No as VCode, TSPL_Payment_Adjustment_Header.Vendor_Name as VName, TSPL_Payment_Adjustment_Header.Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date,TSPL_Payment_Adjustment_Header.Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +TSPL_Payment_Adjustment_Header.Doc_No as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, CASE WHEN ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Type,'')='D' THEN  TSPL_Payment_Adjustment_Header.Adjustment_Amount ELSE 0 END as CrAmt, CASE WHEN ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Type,'')<>'D' THEN  TSPL_Payment_Adjustment_Header.Adjustment_Amount ELSE 0 END  as DrAmt, 0 as Purchase, 0 as Payments, TSPL_Payment_Adjustment_Header.Adjustment_Amount as DrNote, 0 as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, TSPL_VENDOR_INVOICE_HEAD.Loc_Code as account, TSPL_Payment_Adjustment_Header.Post_Date as Posting_Date ,'' as GLDocType, TSPL_Payment_Adjustment_Header.Comp_Code from  TSPL_Payment_Adjustment_Header  LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No  =TSPL_Payment_Adjustment_Header.Doc_No  where isnull(TSPL_Payment_Adjustment_Header.Is_Post,'') ='Y' and TSPL_Payment_Adjustment_Header.Adjust_Type='P'  " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103) <'" + strfromdate + "'  "
            Else
                strtempBaseQry += " and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date,103)  <='" + strtodate + "' "
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_Payment_Adjustment_Header.Vendor_No in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine
            strtempBaseQry += "  select    " & strShowReferenceDocNoofAPInvoice & " ISNULL(TSPL_VENDOR_INVOICE_HEAD.Employee_Type ,'') AS Emp_type,'' as Emp_Adv_Type,NULL AS Due_Date, '' AS PO_SRN, TSPL_Payment_Adjustment_Header.Vendor_No as VCode, TSPL_Payment_Adjustment_Header.Vendor_Name as VName, TSPL_Payment_Adjustment_Header.Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date,TSPL_Payment_Adjustment_Header.Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +TSPL_Payment_Adjustment_Header.Doc_No as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, CASE WHEN ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Type,'')<>'D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END as CrAmt,  CASE WHEN ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Type,'')='D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END as DrAmt, 0 as Purchase, 0 as Payments, 0 as DrNote, TSPL_Payment_Adjustment_Detail.Amount as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, TSPL_VENDOR_INVOICE_HEAD.Loc_Code as account, TSPL_Payment_Adjustment_Header.Post_Date as Posting_Date ,'' as GLDocType, TSPL_Payment_Adjustment_Header.Comp_Code from  TSPL_Payment_Adjustment_Detail  left outer join  TSPL_Payment_Adjustment_Header on TSPL_Payment_Adjustment_Header.Adjustment_No =TSPL_Payment_Adjustment_Detail .Adjustment_No  LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No  =TSPL_Payment_Adjustment_Header.Doc_No  where isnull(TSPL_Payment_Adjustment_Header.Is_Post,'') ='Y' and TSPL_Payment_Adjustment_Header.Adjust_Type='R'  " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += " and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103)  <'" + strfromdate + "' "
            Else
                strtempBaseQry += " and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_Payment_Adjustment_Header.Adjustment_Date,103)  <='" + strtodate + "'"
            End If

            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_Payment_Adjustment_Header.Vendor_No in (" & strvendor & ")"
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine &
            " Select max(Reference_Doc_No) as Reference_Doc_No,max(Emp_type ) as Emp_type,max(Emp_Adv_Type) as Emp_Adv_Type ,NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code from (" + Environment.NewLine &
            " select  '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,NULL AS Due_Date,'' AS PO_SRN, tspl_payment_header.vendor_code as VCode, tspl_payment_header.vendor_name as VName, TSPL_PAYMENT_HEADER.payment_no as DocNo, case when TSPL_PAYMENT_HEADER.Payment_Type='AV' then 'Advance' when TSPL_PAYMENT_HEADER.Payment_Type='OA' then 'On Account' when TSPL_PAYMENT_HEADER.Payment_Type='PY' then 'Payment' when TSPL_PAYMENT_HEADER.Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when TSPL_PAYMENT_HEADER.payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when TSPL_PAYMENT_HEADER.payment_type IN('OA','AV') then Payment_Amount When TSPL_PAYMENT_HEADER.payment_type IN('PY') Then (Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, TSPL_PAYMENT_HEADER.Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('PY') Then case when isnull(TSPL_VENDOR_INVOICE_HEAD.Loc_Code,'')='' then TSPL_PAYMENT_HEADER.Location_GL_Code else TSPL_VENDOR_INVOICE_HEAD.Loc_Code end When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_PAYMENT_HEADER.Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when TSPL_PAYMENT_HEADER.Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, tspl_payment_header.Comp_Code from tspl_payment_header LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_MASTER on tspl_payment_header.Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No Where tspl_payment_header.Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1   " + Environment.NewLine

            If isOpening = True Then
                strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.payment_date ,103)  <'" + strfromdate + "'  "
            Else
                strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.payment_date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.payment_date,103)  <='" + strtodate + "' "
            End If


            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and tspl_payment_header.vendor_code in (" & strvendor & ")"
            End If
            strtempBaseQry += " ) XX Group By XX.account, XX.DocNo"
            If isDocWise = False Then

                strtempBaseQry += " UNION ALL" + Environment.NewLine &
    " select '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_VENDOR_MASTER.Vendor_Code as VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName,TSPL_PAYMENT_HEADER.Payment_No as DocNo,'EXC' as [DocType],convert(date,Payment_Date, 103) as DocDate, case when ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)=1 Then 'Security' Else '' + 'Exchange Gain/Loss against Journal Entry No  '+isnull(TSPL_JOURNAL_MASTER.Voucher_No,'')    end as DocNarr,(TSPL_PAYMENT_HEADER.cheque_No+'-'+ CONVERT(VARCHAR,Cheque_Date , 103))as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT AS CrAmt,  TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT as DrAmt, 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, 0 as Balance_Amount,isnull(TSPL_VENDOR_INVOICE_HEAD.Loc_Code ,'') as account, TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date,'RV-TA' as GLDocType, TSPL_PAYMENT_HEADER.Comp_Code from TSPL_PAYMENT_HEADER  LEFT OUTER JOIN TSPL_Vendor_Master ON TSPL_Vendor_Master.Vendor_Code =TSPL_PAYMENT_HEADER.Vendor_Code " + Environment.NewLine &
    " left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code" + Environment.NewLine &
    " left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =TSPL_PAYMENT_HEADER.Payment_No " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No and TSPL_PAYMENT_DETAIL.Payment_Line_No =1 " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No  " + Environment.NewLine &
    " where TSPL_PAYMENT_HEADER.Posted=1 and  TSPL_PAYMENT_HEADER.Payment_Type  in ('PY','AD') " + Environment.NewLine &
    " AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1 and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT<>0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  <>0)  " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  <'" + strfromdate + "' "
                Else
                    strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "'"
                End If

                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_VENDOR_MASTER.Vendor_Code in (" & strvendor & ")"
                End If
                strtempBaseQry += " ---------------------to find gain or loss amount for payment---------------" + Environment.NewLine &
    " UNION ALL" + Environment.NewLine &
    " Select '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN,TSPL_VENDOR_MASTER.Vendor_Code as VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName, TSPL_BANK_REVERSE.Reverse_Code as DocNo," + Environment.NewLine &
    " 'EXC'  as DocType  ,convert(date,TSPL_BANK_REVERSE.Reversal_Date,103) as DocDate,case when ISNULL(TSPL_PAYMENT_HEADER.Is_Security ,0)=1 Then 'Security' Else '' + 'Exchange Gain/Loss against Journal Entry No  '+isnull(TSPL_JOURNAL_MASTER.Voucher_No,'')    end as DocNarr,(rtrim(Entry_Desc) + (case when len(RTRIM(Entry_Desc))>0 and len(RTRIM(TSPL_PAYMENT_HEADER.Cheque_No))>0  then ' / ' else '' end)  + (case when LEN(ISNULL(TSPL_BANK_REVERSE.Cheque_No,''))>0 then 'Cheque No. - ' + TSPL_BANK_REVERSE.Cheque_No +  ' - '+ convert(varchar ,TSPL_PAYMENT_HEADER . Cheque_Date ,103)else '' end)) as ChequeDetails, TSPL_PAYMENT_HEADER.Currency_Code, 1 as ConvRate,  TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT AS CrAmt, TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT as DrAmt,  0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," + Environment.NewLine &
    " 'RV',0, isnull(TSPL_VENDOR_INVOICE_HEAD.Loc_Code ,'') as account, TSPL_PAYMENT_HEADER.Payment_Date as Posting_Date,'RV-TA' as GLDocType,TSPL_PAYMENT_HEADER.Comp_Code " + Environment.NewLine &
    " from TSPL_BANK_REVERSE  LEFT OUTER JOIN TSPL_VENDOR_MASTER  ON TSPL_VENDOR_MASTER.Vendor_Code =TSPL_BANK_REVERSE.Vendor_Code   " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_Vendor_TYPE_MASTER  ON TSPL_Vendor_TYPE_MASTER.Ven_Type_Code  = TSPL_VENDOR_MASTER.Ven_Type_Code   " + Environment.NewLine &
    " left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =TSPL_BANK_REVERSE.Reverse_Code  " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.PAYMENT_NO =TSPL_BANK_REVERSE.Document_No " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No and TSPL_PAYMENT_DETAIL.Payment_Line_No =1 " + Environment.NewLine &
    " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No " + Environment.NewLine &
    " where  TSPL_BANK_REVERSE.Reverse_Document='Payments' AND TSPL_BANK_REVERSE.Post ='P'" + Environment.NewLine &
    " and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT<>0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  <>0) " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += "and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  <'" + strfromdate + "' "
                Else
                    strtempBaseQry += "and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "'"
                End If

                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and TSPL_VENDOR_MASTER.Vendor_Code in (" & strvendor & ")"
                End If
                strtempBaseQry += " ----------------------to find gain or loss amount FOR BANK REVERSE ---------------" + Environment.NewLine

            End If

            If isIncludeApplyDocument = True Then
                strtempBaseQry += " UNION ALL" + Environment.NewLine &
                " select '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type, NULL AS Due_Date,'' AS PO_SRN, tspl_payment_header.vendor_code as VCode, tspl_payment_header.vendor_name as VName, TSPL_PAYMENT_HEADER.payment_no as DocNo, CASE WHEN TSPL_PAYMENT_HEADER.Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS CrAmt , 0  AS  DrAmt ,0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," + Environment.NewLine &
                " TSPL_PAYMENT_HEADER.Payment_Type as Document_Type,Payment_Amount as Balance_Amount,  CASE WHEN TSPL_PAYMENT_HEADER.Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM TSPL_PAYMENT_HEADER PR WHERE PR.Payment_No =TSPL_PAYMENT_HEADER.Applied_Payment )  ELSE substring(TSPL_PAYMENT_HEADER.Debit_Account , len(TSPL_PAYMENT_HEADER.Debit_Account )-2,3) END as account,TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_PAYMENT_HEADER.Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when TSPL_PAYMENT_HEADER.Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, tspl_payment_header.Comp_Code " + Environment.NewLine &
                " from tspl_payment_header LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_MASTER on tspl_payment_header.Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No Where tspl_payment_header.Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1  " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += "   and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)<'" + strfromdate + "'  "
                Else
                    strtempBaseQry += "   and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' "
                End If
                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and tspl_payment_header.vendor_code in (" & strvendor & ")"
                End If
                strtempBaseQry += " --------------- INVOICE AGAINST APPLY DOCUMENT" + Environment.NewLine &
                " UNION ALL" + Environment.NewLine &
                " SELECT max(Reference_Doc_No) as Reference_Doc_No, max(Emp_type ) as Emp_type,max(Emp_Adv_Type) as Emp_Adv_Type ,NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code FROM (select  '' as Reference_Doc_No,ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,NULL AS Due_Date,'' AS PO_SRN, tspl_payment_header.vendor_code as VCode, tspl_payment_header.vendor_name as VName, TSPL_PAYMENT_HEADER.payment_no as DocNo,CASE WHEN TSPL_PAYMENT_HEADER.Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, 0  AS CrAmt,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS  DrAmt," + Environment.NewLine &
                " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote,TSPL_PAYMENT_HEADER.Payment_Type as Document_Type,Payment_Amount as Balance_Amount, " + Environment.NewLine &
                " TSPL_VENDOR_INVOICE_HEAD.Loc_Code as account,TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_PAYMENT_HEADER.Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when TSPL_PAYMENT_HEADER.Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, tspl_payment_header.Comp_Code  " + Environment.NewLine &
                " from tspl_payment_header LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_MASTER on tspl_payment_header.Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No Where tspl_payment_header.Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1  " + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103) <'" + strfromdate + "' "
                Else
                    strtempBaseQry += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' "
                End If
                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and tspl_payment_header.vendor_code in (" & strvendor & ")"
                End If
                strtempBaseQry += " )INV  GROUP BY  DocNo,account  " + Environment.NewLine &
                " ------- APPLY DOCUMENT ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION " + Environment.NewLine &
                " UNION ALL " + Environment.NewLine &
                " SELECT  max(Reference_Doc_No) as Reference_Doc_No,max(Emp_type ) as Emp_type,max(Emp_Adv_Type) as Emp_Adv_Type ,NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code FROM ( " + Environment.NewLine &
                " select '' as Reference_Doc_No, ISNULL(TSPL_PAYMENT_HEADER.Employee_Type ,'') AS Emp_type,ISNULL(TSPL_PAYMENT_HEADER.Employee_Advance_Type ,'') as Emp_Adv_Type,NULL AS Due_Date,'' AS PO_SRN, tspl_payment_header.vendor_code as VCode, tspl_payment_header.vendor_name as VName, TSPL_BANK_REVERSE.Reverse_Code as DocNo,'Reverse Payment' as DocType , Convert(date,TSPL_BANK_REVERSE.Reversal_Date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((TSPL_PAYMENT_HEADER.Cheque_No) + (case when TSPL_PAYMENT_HEADER.Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when TSPL_VENDOR_INVOICE_HEAD.Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end  AS CrAmt,0 AS  DrAmt, " + Environment.NewLine &
                " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote,TSPL_PAYMENT_HEADER.Payment_Type as Document_Type,Payment_Amount as Balance_Amount,  " + Environment.NewLine &
                " TSPL_VENDOR_INVOICE_HEAD.Loc_Code as account,TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_PAYMENT_HEADER.Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when TSPL_PAYMENT_HEADER.Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, tspl_payment_header.Comp_Code  " + Environment.NewLine &
                " from TSPL_BANK_REVERSE LEFT OUTER JOIN tspl_payment_header ON tspl_payment_header.Payment_No =TSPL_BANK_REVERSE.Document_No  LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_MASTER on tspl_payment_header.Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No Where" + Environment.NewLine &
                " TSPL_BANK_REVERSE.Reverse_Document='Payments' AND TSPL_BANK_REVERSE.Post ='P' AND tspl_payment_header.Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)<>1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1" + Environment.NewLine

                If isOpening = True Then
                    strtempBaseQry += "  and  convert(date,TSPL_BANK_REVERSE.Reversal_Date,103)  <'" + strfromdate + "'  "
                Else
                    strtempBaseQry += "  and  convert(date,TSPL_BANK_REVERSE.Reversal_Date,103)  >='" + strfromdate + "' and  convert(date,TSPL_BANK_REVERSE.Reversal_Date,103)  <='" + strtodate + "' "
                End If
                If clsCommon.myLen(strvendor) > 0 Then
                    strtempBaseQry += " and tspl_payment_header.vendor_code in (" & strvendor & ")"
                End If
                strtempBaseQry += " )INV  GROUP BY  DocNo,account " + Environment.NewLine &
                " -------------- FOR BANK REVERSE ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION" + Environment.NewLine
            End If
            strtempBaseQry += " UNION ALL" + Environment.NewLine &
                      " Select  max(FinalQryForAdvance.Reference_Doc_No) as Reference_Doc_No,max(FinalQryForAdvance.Emp_type) as Emp_type,max(FinalQryForAdvance.Emp_Adv_Type) as Emp_Adv_Type,null as Due_Date ,max(FinalQryForAdvance.PO_SRN) as PO_SRN, max(FinalQryForAdvance.VCode) as VCode,max(FinalQryForAdvance.VName) as VName,FinalQryForAdvance.DocNo,max(FinalQryForAdvance.DocType) as DocType,max(FinalQryForAdvance.DocDate) as DocDate, max(FinalQryForAdvance.DocNarr) as DocNarr, max(FinalQryForAdvance.ChequeDetails) as ChequeDetails, max(FinalQryForAdvance.CURRENCY_CODE) as CURRENCY_CODE, max(FinalQryForAdvance.ConvRate) as ConvRate,sum(CrAmt) as CrAmt,sum(DrAmt) as DrAmt, sum(Purchase) as Purchase, sum(Payments) as Payments, sum(DrNote) as DrNote, sum(CrNote) as CrNote, max(FinalQryForAdvance.Document_Type) as Document_Type ,0 as Balance_Amount, max(FinalQryForAdvance.account) as account " + Environment.NewLine &
                      " , max(FinalQryForAdvance.Posting_Date) as Posting_Date,max(FinalQryForAdvance.GLDocType) as GLDocType,max(FinalQryForAdvance.Comp_Code) as Comp_Code " + Environment.NewLine &
                      " from (select " & strShowReferenceDocNoofAPInvoice & " TSPL_VENDOR_INVOICE_DETAIL.GL_Account_Code, '' AS Emp_type, 'AAS' as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  as VCode, TSPL_VENDOR_INVOICE_HEAD.Vendor_Name as VName, TSPL_VENDOR_INVOICE_HEAD.Document_No as DocNo,'AAS' as DocType,CONVERT(Date,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) as DocDate, TSPL_VENDOR_INVOICE_HEAD.Remarks as DocNarr, '' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate,case when TSPL_VENDOR_INVOICE_DETAIL.Amount <0 then TSPL_VENDOR_INVOICE_DETAIL.Amount*-1  else 0 end as CrAmt,case when TSPL_VENDOR_INVOICE_DETAIL.Amount >0 then TSPL_VENDOR_INVOICE_DETAIL.Amount else 0 end as DrAmt, 0 as Purchase, 0 as Payments, case when TSPL_VENDOR_INVOICE_DETAIL.Amount >0 then TSPL_VENDOR_INVOICE_DETAIL.Amount else 0 end  as DrNote, case when TSPL_VENDOR_INVOICE_DETAIL.Amount <0 then TSPL_VENDOR_INVOICE_DETAIL.Amount*-1  else 0 end as CrNote, TSPL_VENDOR_INVOICE_HEAD.Document_Type as Document_Type ,0 as Balance_Amount, (TSPL_VENDOR_INVOICE_HEAD.Loc_Code ) as account, TSPL_VENDOR_INVOICE_HEAD.Posting_Date as Posting_Date,case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then 'AP-DN' when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType,TSPL_VENDOR_INVOICE_HEAD.Comp_Code from TSPL_VENDOR_INVOICE_DETAIL  LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_VENDOR_INVOICE_DETAIL.Document_No " + Environment.NewLine &
                      " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_INVOICE_HEAD.Vendor_Code =TSPL_VENDOR_MASTER.Vendor_Code " + Environment.NewLine &
                      " left outer join TSPL_VENDOR_ACCOUNT_SET on TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code =TSPL_VENDOR_MASTER.Vendor_Group_Code" + Environment.NewLine &
                      " left outer join  TSPL_GL_ACCOUNTS AS GL1 ON GL1.account_code=TSPL_VENDOR_INVOICE_DETAIL.GL_Account_Code" + Environment.NewLine &
                      " left outer join  TSPL_GL_ACCOUNTS AS GL2 ON GL2.account_code=TSPL_VENDOR_ACCOUNT_SET.Advance_Against_Salary " + Environment.NewLine &
                      " where  ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<>''   " + Environment.NewLine
            If isOpening = True Then
                strtempBaseQry += "  and  convert(date,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date  ,103) <'" + strfromdate + "' "
            Else
                strtempBaseQry += "  and  convert(date,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date  ,103)  >='" + strfromdate + "' and  convert(date,TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date  ,103)  <='" + strtodate + "' "
            End If
            If clsCommon.myLen(strvendor) > 0 Then
                strtempBaseQry += " and TSPL_VENDOR_INVOICE_HEAD.Vendor_Code in (" & strvendor & ")"
            End If
            strtempBaseQry += " and TSPL_VENDOR_MASTER.Vendor_Group_Code  ='EMPLOYEES' and GL1.Account_seg_code1=GL2.Account_seg_code1) FinalQryForAdvance" + Environment.NewLine &
            " group by FinalQryForAdvance.DocNo,FinalQryForAdvance.GL_Account_Code"

            strtempBaseQry += " UNION ALL" + Environment.NewLine &
             "Select max(Reference_Doc_No) as Reference_Doc_No, MAX(Emp_type) as Emp_type, 'S' as Emp_Adv_Type,NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code from (" + Environment.NewLine &
            " select '' as Reference_Doc_No, ISNULL(tspl_payment_header.Employee_Type ,'') AS Emp_type,ISNULL(tspl_payment_header.Employee_Advance_Type ,'') as Emp_Adv_Type,NULL AS Due_Date,'' AS PO_SRN, tspl_payment_header.vendor_code as VCode, tspl_payment_header.vendor_name as VName, TSPL_PAYMENT_HEADER.payment_no as DocNo, case when TSPL_PAYMENT_HEADER.Payment_Type='AV' then 'Advance' when TSPL_PAYMENT_HEADER.Payment_Type='OA' then 'On Account' when TSPL_PAYMENT_HEADER.Payment_Type='PY' then 'Payment' when TSPL_PAYMENT_HEADER.Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when TSPL_PAYMENT_HEADER.payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when TSPL_PAYMENT_HEADER.payment_type IN('OA','AV') then Payment_Amount When TSPL_PAYMENT_HEADER.payment_type IN('PY') Then (Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, TSPL_PAYMENT_HEADER.Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('PY') Then TSPL_VENDOR_INVOICE_HEAD.Loc_Code When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_PAYMENT_HEADER.Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when TSPL_PAYMENT_HEADER.Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, tspl_payment_header.Comp_Code from tspl_payment_header LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_MASTER on tspl_payment_header.Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No Where tspl_payment_header.Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)=1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1   " + Environment.NewLine &
            " ) XX Group By XX.account, XX.DocNo" + Environment.NewLine &
            " UNION ALL" + Environment.NewLine &
            " select '' as Reference_Doc_No, ISNULL(tspl_payment_header.Employee_Type ,'') AS Emp_type,ISNULL(tspl_payment_header.Employee_Advance_Type  ,'S') as Emp_Adv_Type,TSPL_PAYMENT_HEADER.Payment_Date As Due_Date, '' AS PO_SRN, TSPL_REMITTANCE.Vendor_Code, TSPL_REMITTANCE.Vendor_Name, TSPL_REMITTANCE.Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, TSPL_PAYMENT_HEADER.Entry_Desc as DocNarr, '', TSPL_PAYMENT_HEADER.CURRENCY_CODE as Currency_Code, TSPL_PAYMENT_HEADER.ConvRate as ConvRate, case when TSPL_REMITTANCE.Document_Type IN('OA','AV') then 0 else Actual_Total_TDS END as CrAmt, case when TSPL_REMITTANCE.Document_Type IN('OA','AV') then Actual_Total_TDS else 0 END as DrAmt, 0 as Purchase, case when TSPL_REMITTANCE.Document_Type IN('OA','AV') then Actual_Total_TDS else Actual_Total_TDS END as Payments, 0 as DrNotes, 0 as CrNotes, 'TDS',0, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code Else right( Branch_GL_AC,3) End as account, TSPL_PAYMENT_HEADER.Payment_Post_Date as Posting_Date, case when TSPL_REMITTANCE.Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end as GLDocType, TSPL_REMITTANCE.Comp_Code from TSPL_REMITTANCE LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No=TSPL_REMITTANCE.Document_No where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)=1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1" + Environment.NewLine &
            " UNION ALL" + Environment.NewLine &
            " select '' as Reference_Doc_No, ISNULL(tspl_payment_header.Employee_Type ,'') AS Emp_type,ISNULL(tspl_payment_header.Employee_Advance_Type ,'S') as Emp_Adv_Type,null AS Due_Date, '' AS PO_SRN, TSPL_BANK_REVERSE.vendor_code as VCode, TSPL_BANK_REVERSE.vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, TSPL_BANK_REVERSE.Document_No+case when TSPL_PAYMENT_HEADER.Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(TSPL_BANK_REVERSE.cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, TSPL_PAYMENT_HEADER.CURRENCY_CODE, TSPL_PAYMENT_HEADER.ConvRate, case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then 0 when TSPL_PAYMENT_HEADER.Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount ELSE TSPL_BANK_REVERSE.Amount end as CrAmt, case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then TSPL_BANK_REVERSE.Amount else 0 end  as DrAmt, 0 as Purchase, case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then TSPL_BANK_REVERSE.Amount when TSPL_PAYMENT_HEADER.Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount*-1 ELSE TSPL_BANK_REVERSE.Amount*-1 end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When TSPL_PAYMENT_HEADER.Payment_Type in ('AV','OA','RC') Then TSPL_PAYMENT_HEADER.Location_GL_Code When TSPL_PAYMENT_HEADER.Payment_Type in ('PY') Then TSPL_VENDOR_INVOICE_HEAD.Loc_Code Else right(BANKACC,3) End as account, TSPL_BANK_REVERSE.Reversal_Date as Posting_Date,'RV-TA' as GLDocType, TSPL_BANK_REVERSE.Comp_Code from TSPL_BANK_REVERSE LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No = TSPL_BANK_REVERSE.Document_No LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_REVERSE.Bank_Code where Source_Type ='AP' And TSPL_BANK_REVERSE.Post='P' AND TSPL_BANK_REVERSE.Vendor_Code<> '' AND TSPL_PAYMENT_HEADER.IsChkReverse='Y' AND ISNULL(TSPL_PAYMENT_HEADER.Advance_Against_Salary,0)=1 AND ISNULL(TSPL_PAYMENT_HEADER.Is_Security,0)<>1  "

            '   strtempBaseQryforopening = strtempBaseQry
            '  Dim strBaseQryforVendor As String = String.Empty

            'If isDocWise = True Then

            '    If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
            '        strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine &
            '        " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
            '        "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
            '        "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
            '        "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine &
            '        " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) else  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
            '        " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
            '        "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end else " + Environment.NewLine &
            '        "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
            '        " case when InnQuery.DocType='Pur.Invoice' then InnQuery.CrAmt-InnQuery.DrAmt else 0 end as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine &
            '        "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_M1ASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code  "
            '    Else


            '        strtempBaseQry = "  Select isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) as EXCHANGE_GAIN_AMT,isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) as EXCHANGE_LOSS_AMT, InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) - case when (DocType)='Reverse Payment' then  (Select isnull((PH.EXCHANGE_LOSS_AMT  ),0) from TSPL_PAYMENT_HEADER PH where PH.Payment_No =(  Select Document_No   from TSPL_BANK_REVERSE where Reverse_Code =InnQuery.DocNo )) else 0 end as  CrAmt ," + Environment.NewLine &
            '       " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
            '       "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
            '       "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
            '       "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine &
            '       " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT ),0) ELSE  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
            '       " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
            '       "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then CASE WHEN ((DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")-isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0))<0 THEN 0 ELSE (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")-isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) END  else CASE WHEN (DocType)<>'IM' THEN (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) ELSE (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") END end else " + Environment.NewLine &
            '       "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
            '       " InnQuery.CrAmt-InnQuery.DrAmt  as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine &
            '       "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code  "


            '    End If

            'Else
            '    ''richa to exclude exchange gain loss documnets for only apply documnets when include apply doc.checkbox is off  KDI/14/06/2018-000364
            '    Dim strExcludeEXcforApplyDocumnets As String = " where InnQuery.DocNo not in ( Select Payment_No  from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Payment_Type  ='AD' and   TSPL_PAYMENT_HEADER.Posted='1' and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT >0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT >0)  "
            '    If isOpening = True Then
            '        strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <'" + strfromdate + "'  " + Environment.NewLine
            '    Else
            '        strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' " + Environment.NewLine
            '    End If
            '    If clsCommon.myLen(strvendor) > 0 Then
            '        strExcludeEXcforApplyDocumnets += " and TSPL_PAYMENT_HEADER.Vendor_Code in (" & strvendor & ")"
            '    End If
            '    strExcludeEXcforApplyDocumnets += Environment.NewLine & " Union All" & Environment.NewLine &
            '    " Select Reverse_Code  from TSPL_BANK_REVERSE where Document_No in ( Select Payment_No  from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Payment_Type  ='AD' and   TSPL_PAYMENT_HEADER.Posted='1' and (TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT >0 or TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT >0)  "
            '    If isOpening = True Then
            '        strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <'" + strfromdate + "'  " + Environment.NewLine
            '    Else
            '        strExcludeEXcforApplyDocumnets += "  and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  >='" + strfromdate + "' and  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103)  <='" + strtodate + "' " + Environment.NewLine
            '    End If
            '    If clsCommon.myLen(strvendor) > 0 Then
            '        strExcludeEXcforApplyDocumnets += " and TSPL_PAYMENT_HEADER.Vendor_Code in (" & strvendor & ")"
            '    End If

            '    strExcludeEXcforApplyDocumnets += " ) ) "

            '    '----------------------


            '    strtempBaseQry = "  Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  "
            '    If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
            '        strtempBaseQry += " CONVERT(DECIMAL(18,2), CASE WHEN InnQuery.DocType NOT IN ('Pur.Invoice','Receipt')  THEN case when InnQuery.DocType not in ('TDS','AP Invoice') THEN InnQuery.CrAmt  *(case when (DocType) NOT IN ('EXC','Credit Note','IM','Reverse Payment') then   InnQuery.ConvRate else 1 end) when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo)='C' then InnQuery.CrAmt ELSE CASE WHEN InnQuery.DocType in ('TDS') and (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  )='I' or (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo   )='C' then (Select ISNULL (Actual_Total_TDS ,0)  from TSPL_REMITTANCE where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  ) * -1 else 0 end END ELSE 0 END) as CrAmt ," &
            '          " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Adjustment') then case when InnQuery.DocType not in ('AP Invoice','TDS') then InnQuery.DrAmt WHEN InnQuery.DocType in ('TDS') and   (Select ISNULL(TSPL_REMITTANCE.Document_Type,'')  from TSPL_REMITTANCE LEFT OUTER join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No = TSPL_REMITTANCE.Document_No  and  isnull(TSPL_VENDOR_INVOICE_HEAD.Against_PurchaseReturn_No,'')='' where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  )='D' then (Select ISNULL (Actual_Total_TDS ,0)  from TSPL_REMITTANCE where TSPL_REMITTANCE.Document_No=InnQuery.DocNo  ) * -1 else case when  InnQuery.DocType in ('AP Invoice') AND  (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')='')='D' then InnQuery.DrAmt else 0 end end else 0 end) DrAmt, " &
            '          " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Debit Note','Reverse Payment','Credit Note','EXC','IM') then case when InnQuery.DocType not in ('AP Invoice','TDS') then CONVERT(DECIMAL(18,2),InnQuery.Purchase) else case when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo)='I'   then CONVERT(DECIMAL(18,2),InnQuery.Purchase) when  InnQuery.DocType in ('AP Invoice')  and (Select Document_Type  from tspl_vendor_invoice_head where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')<>'')='D' then CONVERT(DECIMAL(18,2),InnQuery.Purchase) *-1  else 0 end end else 0 end) Purchase, " &
            '          " CONVERT(DECIMAL(18,2), case when InnQuery.DocType  in ('Payment','On Account','Advance','Receipt','TDS','EXC','IM') then InnQuery.Payments * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) else 0 end) as Payments , "

            '    Else
            '        strtempBaseQry += " CONVERT(DECIMAL(18,2),InnQuery.CrAmt) AS CrAmt,CONVERT(DECIMAL(18,2),InnQuery.DrAmt) AS DrAmt,CONVERT(DECIMAL(18,2),InnQuery.Purchase) AS Purchase,CONVERT(DECIMAL(18,2),InnQuery.Payments) AS Payments , "
            '    End If

            '    strtempBaseQry += " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code from " &
            '   "(Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end)  as  CrAmt ," + Environment.NewLine &
            '   " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
            '   "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
            '   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
            '   "case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine &
            '   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine &
            '   "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
            '   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
            '   "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
            '   " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine &
            '   "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code " & IIf(isIncludeApplyDocument = False, " " & strExcludeEXcforApplyDocumnets & " ", "") & "  ) InnQuery left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =InnQuery.VCode "

            '    ''richa KDI/15/05/19-000453,KDI/15/05/19-000452
            '    If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then

            '        strtempBaseQryforopeningForMIS = "  Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode,TSPL_VENDOR_MASTER.Vendor_Name as VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  " &
            '      "  CONVERT(DECIMAL(18,2),InnQuery.CrAmt) AS CrAmt,CONVERT(DECIMAL(18,2),InnQuery.DrAmt) AS DrAmt,0 as Purchase,0 as Payments, " &
            '      " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code from " &
            '      " (Select InnQuery.Reference_Doc_No,InnQuery.Emp_type,InnQuery.Emp_Adv_Type,InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine &
            '      " case when len( (isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'')))<=0 then " + Environment.NewLine &
            '      " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine &
            '      " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") +isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine &
            '      " case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") -isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine &
            '      " case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine &
            '      " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(TSPL_PAYMENT_HEADER.Location_GL_Code,'') then " + Environment.NewLine &
            '      "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else isnull((TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine &
            '      "case when isnull((TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(strCurrencyType) = "1", 1, " InnQuery." & strCurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine &
            '      " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code    from  (" & strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine &
            '      "  left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_PAYMENT_HEADER.Bank_Code " & IIf(isIncludeApplyDocument = False, " " & strExcludeEXcforApplyDocumnets & " ", "") & " ) InnQuery left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =InnQuery.VCode "


            '    End If

            'End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return strtempBaseQry

    End Function

    Public Shared Function GetOutStandingQry(ByVal AgeOfDate As Date, ByVal CutOfDate As Date, ByVal DocTypeList As ArrayList, ByVal strCurrency As String, ByVal IsOnDueDate As String, Optional ByVal VendorList As ArrayList = Nothing, Optional ByVal LocationList As ArrayList = Nothing, Optional ByVal VendorGroupList As ArrayList = Nothing) As String
        Try
            Dim StrAgeOfDate As String = ""
            Dim StrCutOfDate As String = ""

            StrAgeOfDate = clsCommon.myCstr(clsCommon.GetPrintDate(AgeOfDate, "dd/MMM/yyyy"))
            StrCutOfDate = clsCommon.GetPrintDate(AgeOfDate, "dd/MMM/yyyy")
            Dim Qry As String = "Select max(Vendor_Invoice_Date) as Vendor_Invoice_Date, max(VI_Due_Date) as VI_Due_Date, XXX.Vendor_Code, MAX(TSPL_VENDOR_MASTER.Vendor_Name) as Vendor_Name,max(tspl_vendor_master.Vendor_Group_Code ) as Vendor_Group_Code, MAX(TSPL_VENDOR_MASTER.Vendor_Group_Code_Desc) as Vendor_Group_Code_Desc, DocNo, MAX(Document_Type) as Document_Type, MAX(DocDate) as DocDate, MAX(Posting_Date) as Posting_Date, MAX(Due_Date) as Due_Date, " & IIf(strCurrency = "1", "MAX(Currency)", "'INR'") & "  as Currency, MAX(ConvRate) as ConvRate, SUM(convert(decimal(18,2),Document_Total*" & strCurrency & ")) as Document_Total, MAX(datedifference) as datedifference, SUM(convert(decimal(18,2),Payment_Amount*" & strCurrency & ")) as Payment_Amount, MAX(invno) as invno, MAX(XXX.Comp_Code) as Comp_Code, MAX(Location) as Location,MAX(XXX.Terms_Code)  AS Terms_Code,MAX(REFDOCNO) AS REFDOCNO,Max(RefDocType) as RefDocType,max(docPosted ) as docPosted  from (" + Environment.NewLine & _
            " select case when ISNULL(tspl_vendor_invoice_head.Posting_Date,'')='' then 0 else 1 end as docPosted,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocNo ,'') REFDOCNO , ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocType  ,'') RefDocType , COnvert(Date,TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date , 103) as Vendor_Invoice_Date,COnvert(Date,TSPL_VENDOR_INVOICE_HEAD.VI_Due_Date , 103) as VI_Due_Date,  Vendor_code,Vendor_Name ,Document_No as DocNo , Document_Type , "

            If clsCommon.CompairString(IsOnDueDate, "DueDate") = CompairStringResult.Equal OrElse clsCommon.CompairString(IsOnDueDate, "DocumentDate") = CompairStringResult.Equal Then
                Qry += "COnvert(Date,Invoice_Entry_Date, 103) as DocDate ,Convert(Date,Posting_Date, 103) as Posting_Date, Convert(Date,Due_Date, 103) as Due_Date , "
            ElseIf clsCommon.CompairString(IsOnDueDate, "VIDate") = CompairStringResult.Equal OrElse clsCommon.CompairString(IsOnDueDate, "VIDueDate") = CompairStringResult.Equal Then
                Qry += "COnvert(Date,Vendor_Invoice_Date, 103) as DocDate ,Convert(Date,Posting_Date, 103) as Posting_Date, Convert(Date,VI_Due_Date, 103) as Due_Date ,"
            End If

            Qry += "(Case When TSPL_VENDOR_INVOICE_HEAD.is_For_TDS=1 Then TSPL_VENDOR_INVOICE_HEAD.Document_Total Else TSPL_VENDOR_INVOICE_HEAD.Document_Total-TDS_Actual_Amount-ISNULL((Select case when vh.Balance_Amt =TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount else 0 end then 0 else vh.Balance_Amt end from TSPL_VENDOR_INVOICE_HEAD as VH Where VH.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code AND VH.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No AND VH.Document_Type IN ('D','C') and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1 AND ISNULL(VH.Against_PurchaseReturn_No,'')<>''),0) " & _
            "  - ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I'  and ISNULL( TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount ,0)=0  then ( select  sum(isnull(Document_Total,0)) - sum(isnull(inn.TDS_Actual_Amount ,0)) from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No AND CONVERT(Date,TSPL_PR_HEAD.PR_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) )  and inn.Document_Type='D'   and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ) else case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then ( select  sum(isnull(Document_Total,0))-sum(isnull(inn.TDS_Actual_Amount ,0)) from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No AND CONVERT(Date,TSPL_PR_HEAD.PR_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103))  and inn.Document_Type='D'   and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ) else 0 end end,0)  " & Environment.NewLine & _
            " - ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I'  and ISNULL( TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount ,0)=0  then ( select  sum(isnull(Document_Total,0)) - sum(isnull(inn.TDS_Actual_Amount ,0)) from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D'   and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ) else case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then ( select  sum(isnull(Document_Total,0))" & Environment.NewLine & _
            " -sum(isnull(inn.TDS_Actual_Amount ,0)) from TSPL_VENDOR_INVOICE_HEAD as inn where  inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ) else 0 end end,0) "


            Qry += "  End) " &
            " -ISNULL((Select SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_DETAIL.Payment_No Where TSPL_PAYMENT_HEADER.Posted=1 AND TSPL_PAYMENT_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND  CONVERT(DATE,TSPL_PAYMENT_HEADER.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103)),0) " &
            " +ISNULL((Select sum(Applied_Amount) from TSPL_PAYMENT_DETAIL LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_HEADER.Payment_No=TSPL_PAYMENT_DETAIL.Payment_No left outer join TSPL_BANK_REVERSE on TSPL_BANK_REVERSE.Document_No = TSPL_PAYMENT_HEADER.Payment_No Where TSPL_BANK_REVERSE.Reverse_Document ='Payments' and TSPL_BANK_REVERSE.Source_Type ='AP' and TSPL_PAYMENT_HEADER.Posted=1 AND TSPL_BANK_REVERSE.Post ='P' and TSPL_PAYMENT_HEADER.IsChkReverse='Y' AND TSPL_PAYMENT_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND CONVERT(DATE,TSPL_BANK_REVERSE.Reversal_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103)),0) " &
" - ISNULL((Select SUM(TSPL_PAYMENT_HEADER.Payment_Amount ) from TSPL_PAYMENT_HEADER  LEFT OUTER JOIN TSPL_BANK_REVERSE ON TSPL_BANK_REVERSE.Document_No = TSPL_PAYMENT_HEADER.Payment_No  AND ISNULL(TSPL_BANK_REVERSE.Source_Type,'') ='AP' AND TSPL_BANK_REVERSE.Post ='P' WHERE ISNULL(TSPL_BANK_REVERSE.REVERSE_CODE,'' )='' and  ISNULL(TSPL_PAYMENT_HEADER.Posted,0)=1 AND TSPL_PAYMENT_HEADER.Payment_Type  in ('AD') AND TSPL_PAYMENT_HEADER.Applied_Payment =TSPL_VENDOR_INVOICE_HEAD.Document_No AND  CONVERT(DATE,TSPL_PAYMENT_HEADER.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) ),0) " &
 " - isnull((case when TSPL_VENDOR_INVOICE_HEAD.Document_Type IN ('I') then (Select SUM(Document_Total) - SUM(TDS_Actual_Amount) from TSPL_VENDOR_INVOICE_HEAD VH WHERE isnull(VH.Posting_Date,'')<>'' AND vh.Document_Type='D' AND VH.Document_No  =TSPL_VENDOR_INVOICE_HEAD.Document_No and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1  AND VH.Document_No  <>'' ) else 0 end ),0) " &
            " + isnull((case when TSPL_VENDOR_INVOICE_HEAD.Document_Type IN ('I') then (Select SUM(Document_Total) from TSPL_VENDOR_INVOICE_HEAD VH WHERE isnull(VH.Posting_Date,'')<>'' AND VH.Document_Type='C' AND VH.Document_No  =TSPL_VENDOR_INVOICE_HEAD.Document_No  and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1 AND VH.Document_No  <>'' ) else 0 end ),0) " &
            " - isnull((select SUM(inn.Document_Total)  from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Document_Type='D' and len(isnull(inn.Against_POInvoice_No,''))>0 and inn.Against_POInvoice_No= TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1  and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_For_TDS ,0)=1 ),0)" &
            " + isnull((case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and LEN(ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,''))>0 then 0 else (Select sum(TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt) from TSPL_TRANSFER_ORDER_HEAD where Status=1 and TSPL_TRANSFER_ORDER_HEAD.RMDA_Code in (select TSPL_SRN_HEAD.RMDA_No from TSPL_PI_DETAIL left outer join TSPL_PI_HEAD on TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_PI_HEAD.Against_SRN where TSPL_PI_HEAD.PI_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and isnull(TSPL_SRN_HEAD.RMDA_No,'')<>'' and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1))   end ),0)" &
" -(SELECT SUM(TAXAMUNT) FROM (Select  ISNULL((case when VIH.GSTRegistered =0 or VIH.RCM=1 then  ( case when len(isnull(VIH.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.Tax1)='Y'  then VIH.TAX1_Amt else 0 end +  case when len(isnull(VIH.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX2)='Y'  then VIH.TAX2_Amt else 0 end +  case when len(isnull(VIH.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX3)='Y'  then VIH.TAX3_Amt else 0 end +  case when len(isnull(VIH.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX4)='Y'  then VIH.TAX4_Amt else 0 end +  case when len(isnull(VIH.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX5)='Y'  then VIH.TAX5_Amt else 0 end +  case when len(isnull(VIH.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX6)='Y'   then VIH.TAX6_Amt else 0 end +  case when len(isnull(VIH.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX7)='Y'  then VIH.TAX7_Amt else 0 end +  case when len(isnull(VIH.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX8)='Y'  then VIH.TAX8_Amt else 0 end +  case when len(isnull(VIH.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.TAX9)='Y'  then VIH.TAX9_Amt else 0 end +  case when len(isnull(VIH.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =VIH.Tax10)='Y'  then VIH.TAX10_Amt else 0 end ) else 0 end),0) TAXAMUNT from TSPL_VENDOR_INVOICE_HEAD VIH where VIH.Document_No =TSPL_VENDOR_INVOICE_HEAD.Document_No) Z)" &
 " as [Document_Total],"

            If clsCommon.CompairString(IsOnDueDate, "DueDate") = CompairStringResult.Equal Then
                ' Qry += " DATEDIFF(dd,convert(date," + IIf(IsOnDueDate = True, "Due_Date", "Invoice_Entry_Date") + ",103),'" & StrAgeOfDate & "') as datedifference, "
                Qry += " DATEDIFF(dd,convert(date,Due_Date,103),'" & StrAgeOfDate & "') as datedifference, "
            ElseIf clsCommon.CompairString(IsOnDueDate, "DocumentDate") = CompairStringResult.Equal Then
                Qry += " DATEDIFF(dd,convert(date,Invoice_Entry_Date,103),'" & StrAgeOfDate & "') as datedifference, "
            ElseIf clsCommon.CompairString(IsOnDueDate, "VIDate") = CompairStringResult.Equal Then
                Qry += " DATEDIFF(dd,convert(date,Vendor_Invoice_Date,103),'" & StrAgeOfDate & "') as datedifference, "
            ElseIf clsCommon.CompairString(IsOnDueDate, "VIDueDate") = CompairStringResult.Equal Then
                Qry += " DATEDIFF(dd,convert(date,VI_Due_Date,103),'" & StrAgeOfDate & "') as datedifference, "
            End If

            'Qry += " 0 as Payment_Amount , Vendor_Invoice_No  as invno, TSPL_VENDOR_INVOICE_HEAD.Comp_Code, TSPL_VENDOR_INVOICE_HEAD.Loc_Code AS Location,TSPL_VENDOR_INVOICE_HEAD.Terms_Code, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE as Currency, TSPL_VENDOR_INVOICE_HEAD.ConvRate from TSPL_VENDOR_INVOICE_HEAD" & _
            '" where ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<> '' And (Len(ISNULL(Against_POInvoice_No, ''))<=0 AND Len(ISNULL(Against_PurchaseReturn_No, ''))<=0) AND RefDocType <> 'AP' or Document_Type IN ('I','D','C')   " + Environment.NewLine & _
            '" UNION ALL" + Environment.NewLine & _
            '" select TSPL_PAYMENT_HEADER.Posted as docPosted,'' AS  REFDOCNO ,'' as RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date,Vendor_code,Vendor_Name ,TSPL_PAYMENT_HEADER.Payment_No as DocNo ,TSPL_PAYMENT_HEADER.Payment_Type as Document_Type , Convert(Date,Payment_Date, 103) as DocDate , Convert(Date,Payment_Post_Date, 103) as Posting_Date, Convert(Date,Payment_Date, 103) as Due_Date , " & _
            '" Payment_Amount+TDS_Amount - isnull((case when TSPL_PAYMENT_HEADER.Payment_Type in('AV','OA','RC') then (select Amount from TSPL_BANK_REVERSE where Source_Type='AP' and Document_No=TSPL_PAYMENT_HEADER.Payment_No AND TSPL_BANK_REVERSE.Post<>'N' AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND TSPL_BANK_REVERSE.Reversal_Date<=Convert(Date,'" & StrCutOfDate & "', 103)) else 0 end ),0) " & _
            '"  -isnull((case when TSPL_PAYMENT_HEADER.Payment_Type in('AV','OA') then (Select SUM(PH.Payment_Amount) from TSPL_PAYMENT_HEADER PH WHERE PH.Posted='1' AND PH.Payment_Type='AD' AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'' and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No in (select PPH.Payment_No from TSPL_PAYMENT_HEADER PPH where PPH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) and isnull(TSPL_BANK_REVERSE.POST,'')='P' )) AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND  CONVERT(DATE,PH.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) )else 0 end ),0) " & _
            '" -isnull((Select sum(isnull(Applied_amount,0)) from TSPL_PAYMENT_DETAIL left outer join TSPL_PAYMENT_HEADER TPH on  TSPL_PAYMENT_DETAIL.Payment_No =TPH.Payment_No where TSPL_PAYMENT_DETAIL .Document_No =TSPL_PAYMENT_HEADER.Payment_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )AND  CONVERT(DATE,TPH.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) and TPH.Posted='1'),0) " & _
            '" as [Document_Total]" & _
            '" ,DATEDIFF(dd,convert(date,Payment_Date,103), '" & StrAgeOfDate & "') as datedifference, Case When TSPL_PAYMENT_HEADER.Payment_Type='RC' Then Payment_Amount*-1 Else Payment_Amount End as Payment_Amount,''  as invno, TSPL_PAYMENT_HEADER.Comp_Code, RIGHT(TSPL_BANK_MASTER.BANKACC , 3) as Location,'' AS Terms_Code, TSPL_PAYMENT_HEADER.CURRENCY_CODE as Currency, TSPL_PAYMENT_HEADER.ConvRate from TSPL_PAYMENT_HEADER" & _
            '" LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_PAYMENT_HEADER.Bank_Code= TSPL_BANK_MASTER.BANK_CODE where TSPL_PAYMENT_HEADER.Payment_Type NOT IN ('PY','MI') And "

            'UDL/03/10/18-000226 richa UDL/10/10/18-000229
            Qry += " 0 as Payment_Amount , Vendor_Invoice_No  as invno, TSPL_VENDOR_INVOICE_HEAD.Comp_Code, TSPL_VENDOR_INVOICE_HEAD.Loc_Code AS Location,TSPL_VENDOR_INVOICE_HEAD.Terms_Code, TSPL_VENDOR_INVOICE_HEAD.CURRENCY_CODE as Currency, TSPL_VENDOR_INVOICE_HEAD.ConvRate from TSPL_VENDOR_INVOICE_HEAD" & _
           " where ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<> '' And (Len(ISNULL(Against_POInvoice_No, ''))<=0 AND Len(ISNULL(Against_PurchaseReturn_No, ''))<=0) AND RefDocType <> 'AP' or Document_Type IN ('I','D','C')   " + Environment.NewLine & _
           " UNION ALL" + Environment.NewLine & _
           " select TSPL_PAYMENT_HEADER.Posted as docPosted,'' AS  REFDOCNO ,'' as RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date,Vendor_code,Vendor_Name ,TSPL_PAYMENT_HEADER.Payment_No as DocNo ,TSPL_PAYMENT_HEADER.Payment_Type as Document_Type , Convert(Date,Payment_Date, 103) as DocDate , Convert(Date,Payment_Post_Date, 103) as Posting_Date, Convert(Date,Payment_Date, 103) as Due_Date , " & _
           " Payment_Amount+TDS_Amount - isnull((case when TSPL_PAYMENT_HEADER.Payment_Type in('AV','OA','RC') then (select Amount from TSPL_BANK_REVERSE where Source_Type='AP' and Document_No=TSPL_PAYMENT_HEADER.Payment_No AND TSPL_BANK_REVERSE.Post<>'N' AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND TSPL_BANK_REVERSE.Reversal_Date<=Convert(Date,'" & StrCutOfDate & "', 103)) else 0 end ),0) " & _
           "  -isnull((case when TSPL_PAYMENT_HEADER.Payment_Type in('AV','OA') then (Select SUM(PH.Payment_Amount) from TSPL_PAYMENT_HEADER PH WHERE PH.Posted='1' AND PH.Payment_Type='AD' AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'' and PH.Payment_No not in (select TSPL_BANK_REVERSE.Document_No from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No in (select PPH.Payment_No from TSPL_PAYMENT_HEADER PPH where PPH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) and isnull(TSPL_BANK_REVERSE.POST,'')='P' )) AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1 AND  CONVERT(DATE,PH.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) )else 0 end ),0) " & _
           " -isnull((Select sum(isnull(Applied_amount,0)) from TSPL_PAYMENT_DETAIL left outer join TSPL_PAYMENT_HEADER TPH on  TSPL_PAYMENT_DETAIL.Payment_No =TPH.Payment_No where TSPL_PAYMENT_DETAIL .Document_No =TSPL_PAYMENT_HEADER.Payment_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )AND  CONVERT(DATE,TPH.Payment_Date,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) and TPH.Posted='1'),0) " & _
           " as [Document_Total]" & _
           " ,DATEDIFF(dd,convert(date,Payment_Date,103), '" & StrAgeOfDate & "') as datedifference, Case When TSPL_PAYMENT_HEADER.Payment_Type='RC' Then Payment_Amount*-1 Else Payment_Amount End as Payment_Amount,''  as invno, TSPL_PAYMENT_HEADER.Comp_Code, RIGHT(TSPL_BANK_MASTER.BANKACC , 3) as Location,'' AS Terms_Code, TSPL_PAYMENT_HEADER.CURRENCY_CODE as Currency, TSPL_PAYMENT_HEADER.ConvRate from TSPL_PAYMENT_HEADER" & _
           " LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_PAYMENT_HEADER.Bank_Code= TSPL_BANK_MASTER.BANK_CODE where TSPL_PAYMENT_HEADER.Payment_Type NOT IN ('PY','MI') And "


            Qry += " (Posted='P' or Posted='1')  AND ISNULL(TSPL_PAYMENT_HEADER.is_security,0)<>1" + Environment.NewLine &
            " UNION ALL" + Environment.NewLine &
            " select TSPL_VCGL_Head.Status as docPosted,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocNo ,'') REFDOCNO ,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocType  ,'') RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date, VC_Code as Vendor_code,VC_Name as Vendor_Name, TSPL_VCGL_Head.Document_No as DocNo, case when Amount_Type='Cr' then 'D' else  'C' end as Document_Type,CONVERT(Date,TSPL_VCGL_Head.Document_Date,103) as DocDate," &
            " CONVERT(Date,TSPL_VCGL_Head.Posting_Date ,103) as Posting_Date, CONVERT(Date,TSPL_VCGL_Head.Document_Date,103) as Due_Date, Amount as [Document_Total], DATEDIFF(dd,convert(date,TSPL_VCGL_Head.Posting_Date,103) , '" & StrAgeOfDate & "') as datedifference, amount as Payment_Amount,'' as invno, TSPL_VCGL_Head.Comp_Code ,  TSPL_VCGL_Head.Location_Segment as Location,'' AS Terms_Code, CM.BaseCurrencyCode as Currency, 1 as ConvRate from TSPL_VCGL_Head" &
            " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_VCGL=TSPL_VCGL_Head.Document_No" &
            " LEFT OUTER JOIN TSPL_COMPANY_MASTER CM ON CM.Comp_Code=TSPL_VCGL_Head.Comp_Code" &
            " where TSPL_VCGL_Head.Document_Type='v' and TSPL_VCGL_Head.Status='1' AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_VCGL,'')='' and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1" + Environment.NewLine &
             " UNION ALL" + Environment.NewLine
            ''richa agarwal against ticket number BM00000007350 to show payment adjustment header impact on vendor ap invoice aging
            Qry += " select case when isnull(TSPL_Payment_Adjustment_Header.Is_Post ,'') ='Y' then 1 else 0 end  as docPosted,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocNo ,'') REFDOCNO ,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocType  ,'') RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date,TSPL_Payment_Adjustment_Header.Vendor_No   as Vendor_code,TSPL_Payment_Adjustment_Header.Vendor_Name  as Vendor_Name, TSPL_Payment_Adjustment_Header.Doc_No  as DocNo," & _
             " 'Adjustment' as Document_Type,CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103) as DocDate, CONVERT(Date,TSPL_Payment_Adjustment_Header.Post_Date ,103) as Posting_Date," & _
             " CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103) as Due_Date, TSPL_Payment_Adjustment_Header.Adjustment_Amount  * -1  as [Document_Total], " & _
             " DATEDIFF(dd,convert(date,TSPL_Payment_Adjustment_Header.Post_Date ,103) ,'" & StrAgeOfDate & "') as datedifference, TSPL_Payment_Adjustment_Header.Adjustment_Amount  as Payment_Amount," & _
             " '' as invno, TSPL_Payment_Adjustment_Header.Comp_Code ,  TSPL_VENDOR_INVOICE_HEAD.Loc_Code  as Location,'' AS Terms_Code, CM.BaseCurrencyCode as Currency," & _
             " 1 as ConvRate from TSPL_Payment_Adjustment_Header  LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_Payment_Adjustment_Header.Doc_No" & _
             " LEFT OUTER JOIN TSPL_COMPANY_MASTER CM ON CM.Comp_Code=TSPL_Payment_Adjustment_Header.Comp_Code " & _
             " where  isnull(TSPL_Payment_Adjustment_Header.Is_Post ,'') ='Y' AND ISNULL(TSPL_Payment_Adjustment_Header.Doc_No ,'')<>'' and TSPL_Payment_Adjustment_Header.Adjust_Type='P' AND  CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) " & _
             " UNION ALL" + Environment.NewLine & _
             " select case when isnull(TSPL_Payment_Adjustment_Header.Is_Post ,'') ='Y' then 1 else 0 end  as docPosted,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocNo ,'') REFDOCNO ,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocType  ,'') RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date,TSPL_Payment_Adjustment_Header.Vendor_No   as Vendor_code,TSPL_Payment_Adjustment_Header.Vendor_Name  as Vendor_Name, TSPL_Payment_Adjustment_Header.Doc_No  as DocNo," & _
             " 'Adjustment' as Document_Type,CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103) as DocDate, CONVERT(Date,TSPL_Payment_Adjustment_Header.Post_Date ,103) as Posting_Date," & _
             " CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103) as Due_Date, TSPL_Payment_Adjustment_Header.Adjustment_Amount   as [Document_Total], " & _
             " DATEDIFF(dd,convert(date,TSPL_Payment_Adjustment_Header.Post_Date ,103) ,'" & StrAgeOfDate & "') as datedifference, TSPL_Payment_Adjustment_Header.Adjustment_Amount * -1  as Payment_Amount," & _
             " '' as invno, TSPL_Payment_Adjustment_Header.Comp_Code ,  TSPL_VENDOR_INVOICE_HEAD.Loc_Code  as Location,'' AS Terms_Code, CM.BaseCurrencyCode as Currency," & _
             " 1 as ConvRate from TSPL_Payment_Adjustment_Header  LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Document_No =TSPL_Payment_Adjustment_Header.Doc_No" & _
             " LEFT OUTER JOIN TSPL_COMPANY_MASTER CM ON CM.Comp_Code=TSPL_Payment_Adjustment_Header.Comp_Code " & _
             " where  isnull(TSPL_Payment_Adjustment_Header.Is_Post ,'') ='Y' AND ISNULL(TSPL_Payment_Adjustment_Header.Doc_No ,'')<>'' and TSPL_Payment_Adjustment_Header.Adjust_Type='R' AND  CONVERT(Date,TSPL_Payment_Adjustment_Header.Adjustment_Date ,103)<=CONVERT(DATE,'" & StrCutOfDate & "',103) " & _
             " UNION ALL" + Environment.NewLine & _
            " select TSPL_VCGL_Head.Status as docPosted,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocNo ,'') REFDOCNO ,ISNULL(TSPL_VENDOR_INVOICE_HEAD.RefDocType  ,'') RefDocType , null as Vendor_Invoice_Date,null as VI_Due_Date,TSPL_VCGL_Detail.VCGL_Code  as Vendor_code, TSPL_VCGL_Detail.VCGL_Name  as Vendor_Name, TSPL_VCGL_Detail.Document_No as DocNo, case when TSPL_VCGL_Detail.Dr_Amount=0.0 then 'C' when TSPL_VCGL_Detail.Cr_Amount=0.0 then 'D' end as Document_Type, CONVERT(Date,TSPL_VCGL_Head.Document_Date,103) as DocDate, CONVERT(Date,TSPL_VCGL_Head.Posting_Date ,103) as Posting_Date, CONVERT(Date,TSPL_VCGL_Head.Document_Date,103) as Due_Date, case when TSPL_VCGL_Detail.Dr_Amount=0.0 then TSPL_VCGL_Detail.Cr_Amount when TSPL_VCGL_Detail.Cr_Amount=0.0 then TSPL_VCGL_Detail.Dr_Amount end as [Document_Total] ,DATEDIFF(dd,convert(date,TSPL_VCGL_Head.Posting_Date,103), '" & StrAgeOfDate & "') as datedifference, amount as Payment_Amount,'' as invno , TSPL_VCGL_Head.Comp_Code ,  TSPL_VCGL_Head.Location_Segment as Location,'' AS Terms_Code, CM.BaseCurrencyCode as Currency, 1 as ConvRate from  TSPL_VCGL_Detail" & _
            " left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No=TSPL_VCGL_Detail.Document_No" & _
            " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD ON TSPL_VENDOR_INVOICE_HEAD.Against_VCGL=TSPL_VCGL_Head.Document_No" & _
            " LEFT OUTER JOIN TSPL_COMPANY_MASTER CM ON CM.Comp_Code=TSPL_VCGL_Head.Comp_Code" & _
            " where Row_Type='Vendor'  and TSPL_VCGL_Head.Status='1' AND ISNULL(TSPL_VENDOR_INVOICE_HEAD.Against_VCGL,'')='' and ISNULL( TSPL_VENDOR_INVOICE_HEAD.is_Security,0)<>1" + Environment.NewLine & _
            " ) XXX LEFT OUTER JOIN TSPL_VENDOR_MASTER ON XXX.Vendor_Code=TSPL_VENDOR_MASTER.Vendor_Code" & _
             " where XXX.Document_Type in (" & clsCommon.GetMulcallString(DocTypeList) & ")"
            '" where Document_Total<>0 AND XXX.Document_Type in (" & clsCommon.GetMulcallString(DocTypeList) & ")"
            If VendorList IsNot Nothing AndAlso VendorList.Count > 0 Then
                Qry += " and XXX.Vendor_Code in (" & clsCommon.GetMulcallString(VendorList) & ") "
            End If
            If LocationList IsNot Nothing AndAlso LocationList.Count > 0 Then
                Qry += " and XXX.Location in (" & clsCommon.GetMulcallString(LocationList) & ") "
            End If
            If VendorGroupList IsNot Nothing AndAlso VendorGroupList.Count > 0 Then
                Qry += " AND TSPL_VENDOR_MASTER.Vendor_Group_Code IN (" & clsCommon.GetMulcallString(VendorGroupList) & ") "
            End If

            If clsCommon.CompairString(strCurrency, "1") = CompairStringResult.Equal Then
                Qry += " AND TSPL_VENDOR_MASTER .CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  "
            End If

            'Qry += " AND XXX.DocNo NOT IN (sELECT TSPL_PAYMENT_DETAIL.Document_No FROM TSPL_PAYMENT_DETAIL WHERE TSPL_PAYMENT_DETAIL.Document_No=XXX.DocNo  AND TSPL_PAYMENT_DETAIL.Applied_Amount =XXX.Document_Total  AND XXX.Document_Type IN ('C','D') AND ISNULL(TSPL_PAYMENT_DETAIL.Document_No,'')<>'') Group By XXX.Vendor_Code, XXX.DocNo"
            Qry += " AND XXX.DocNo NOT IN (sELECT TSPL_PAYMENT_DETAIL.Document_No FROM TSPL_PAYMENT_DETAIL LEFT OUTER JOIN TSPL_PAYMENT_HEADER ON TSPL_PAYMENT_DETAIL.Payment_No =TSPL_PAYMENT_HEADER.Payment_No left outer join TSPL_BANK_REVERSE on TSPL_BANK_REVERSE.Document_No = TSPL_PAYMENT_HEADER.Payment_No WHERE TSPL_PAYMENT_DETAIL.Document_No=XXX.DocNo  AND TSPL_PAYMENT_DETAIL.Applied_Amount =XXX.Document_Total  AND XXX.Document_Type IN ('C','D') AND TSPL_PAYMENT_HEADER.Posted=1 AND ISNULL(TSPL_PAYMENT_DETAIL.Document_No,'')<>'' AND Convert(Date, TSPL_PAYMENT_HEADER.Payment_Date , 103) <=Convert(Date,'" + StrCutOfDate + "', 103) and (TSPL_BANK_REVERSE.Document_No<>TSPL_PAYMENT_HEADER.Payment_No AND TSPL_BANK_REVERSE.Post ='P' AND CONVERT(DATE,TSPL_BANK_REVERSE.Reversal_Date ,103)<=Convert(Date,'" + StrCutOfDate + "', 103)) )" & _
            " Group By XXX.Vendor_Code, XXX.DocNo"
            ''" and XXX.DocNo not in ( Select TSPL_VENDOR_INVOICE_HEAD.RefDocNo  from TSPL_VENDOR_INVOICE_HEAD where TSPL_VENDOR_INVOICE_HEAD .RefDocNo =XXX.DocNo and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' AND Convert(Date, TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date  , 103) <=Convert(Date,'" + StrCutOfDate + "', 103)AND  ISNULL(tspl_vendor_invoice_head.Posting_Date,'')<> '' and TSPL_VENDOR_INVOICE_HEAD .Document_No <>'' and TSPL_VENDOR_INVOICE_HEAD.is_For_TDS =1) " & _


            Return Qry
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    
    Public Shared Function GetVendorBaseQry(ByRef objFilter As structVendorFilter) As String
        Dim strtempBaseQry As String = String.Empty
        Dim TSPL_VENDOR_INVOICE_HEAD As String = "TSPL_VENDOR_INVOICE_HEAD"
        Dim TSPL_PAYMENT_HEADER = "TSPL_PAYMENT_HEADER"
        Dim TSPL_Payment_Adjustment_Header = "TSPL_Payment_Adjustment_Header"
        Dim TSPL_BANK_REVERSE As String = "TSPL_BANK_REVERSE"
        Dim TSPL_REMITTANCE As String = "TSPL_REMITTANCE"
        Dim TSPL_VCGL_Head As String = "TSPL_VCGL_Head"
        Dim TSPL_PI_HEAD As String = "TSPL_PI_HEAD"
        Dim OP_TypeCol As String = "'' as OP_Type "
        If objFilter.IS_GIT Then
            TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD_WIN"
            TSPL_PAYMENT_HEADER = "TSPL_PAYMENT_HEADER_WIN"
            TSPL_Payment_Adjustment_Header = "TSPL_Payment_Adjustment_Header_WIN"
            TSPL_BANK_REVERSE = "TSPL_BANK_REVERSE_WIN"
            TSPL_REMITTANCE = "TSPL_REMITTANCE_WIN"
            TSPL_VCGL_Head = "TSPL_VCGL_Head_WIN"
            TSPL_PI_HEAD = "TSPL_PI_HEAD_WIN"
            OP_TypeCol = " OP_Type"
        End If
        Try
            If Not objFilter.IsOnlyForAgainstSalary Then
                Dim strQryForRMDA As String = " - (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' and LEN(ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No,''))>0 then (Select sum(TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt) from TSPL_TRANSFER_ORDER_HEAD where Status=1 and TSPL_TRANSFER_ORDER_HEAD.RMDA_Code in (select TSPL_SRN_HEAD.RMDA_No from TSPL_PI_DETAIL left outer join " & TSPL_PI_HEAD & " on " & TSPL_PI_HEAD & ".PI_No=TSPL_PI_DETAIL.PI_No left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=" & TSPL_PI_HEAD & ".Against_SRN where " & TSPL_PI_HEAD & ".PI_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No  )) else 0 end )"
                strtempBaseQry = " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date, '' AS PO_SRN, vendor_code as VCode,vendor_name as VName," & TSPL_VENDOR_INVOICE_HEAD & " .Document_No as DocNo ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'Debit Note' else case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'Credit Note' else 'AP Invoice' end  end as DocType,convert(date,Invoice_Entry_Date,103) as DocDate,((Description) + (case when Description='' then case when " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks <>'' then " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks +' - ' else ''  end else ' - 'end) +(RefDocNo)+'Vendor Invoice No-' +Vendor_Invoice_No+' Date-' +Vendor_Invoice_Date  ) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate,"
                ', case when "& TSPL_VENDOR_INVOICE_HEAD &".Document_Type IN('I','C') Then document_total else 0 end as CrAmt
                strtempBaseQry += " case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type IN('I','C') Then document_total else 0 end + ( case when isnull(TAX1 ,'')='WCT' THEN TAX1_Amt * -1  when isnull(TAX2 ,'')='WCT' THEN TAX2_Amt * -1 when isnull(TAX3 ,'')='WCT' THEN TAX3_Amt * -1 when isnull(TAX4 ,'')='WCT' THEN TAX4_Amt * -1 when isnull(TAX5 ,'')='WCT' THEN TAX5_Amt * -1 when isnull(TAX6 ,'')='WCT' THEN TAX6_Amt * -1 when isnull(TAX7 ,'')='WCT' THEN TAX7_Amt * -1 when isnull(TAX8 ,'')='WCT' THEN TAX8_Amt * -1 when isnull(TAX9 ,'')='WCT' THEN TAX9_Amt * -1 when isnull(TAX10 ,'')='WCT' THEN TAX10_Amt * -1 ELSE 0 END) as CrAmt, " & _
                    "  case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type IN('D') then Document_Total else 0 end as DrAmt, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' Then document_total Else 0 End as Purchase, 0 as Payments, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then document_total Else 0 End as DrNote, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' Then document_total Else 0 End as CrNote, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,balance_Amt as Balance_Amount, ( select   GL_Account_Code from TSPL_VENDOR_INVOICE_DETAIL where Detail_Line_No='1' and Document_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  ) as account, convert(date," & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_VENDOR_INVOICE_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VENDOR_INVOICE_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_VENDOR_INVOICE_HEAD & " where ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,'')<>''  and LEN(ISNULL( " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No,''))<=0   " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine
                ''------------- code inserted for WCT type of documents
                strtempBaseQry += " Select * from ( select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date, '' AS PO_SRN, vendor_code as VCode,vendor_name as VName," & TSPL_VENDOR_INVOICE_HEAD & " .Document_No as DocNo ,'WCT' as DocType,convert(date,Invoice_Entry_Date,103) as DocDate,((Description) + (case when Description='' then case when " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks <>'' then " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks +' - ' else ''  end else ' - 'end) +(RefDocNo)+'Vendor Invoice No-' +Vendor_Invoice_No+' Date-' +Vendor_Invoice_Date  ) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate," & _
                " 0 as CrAmt,case when isnull(TAX1 ,'')='WCT' THEN TAX1_Amt * -1  when isnull(TAX2 ,'')='WCT' THEN TAX2_Amt * -1 when isnull(TAX3 ,'')='WCT' THEN TAX3_Amt * -1 when isnull(TAX4 ,'')='WCT' THEN TAX4_Amt * -1 when isnull(TAX5 ,'')='WCT' THEN TAX5_Amt * -1 when isnull(TAX6 ,'')='WCT' THEN TAX6_Amt * -1 when isnull(TAX7 ,'')='WCT' THEN TAX7_Amt * -1 when isnull(TAX8 ,'')='WCT' THEN TAX8_Amt * -1 when isnull(TAX9 ,'')='WCT' THEN TAX9_Amt * -1 when isnull(TAX10 ,'')='WCT' THEN TAX10_Amt * -1 ELSE 0 END  DrAmt, " & _
                " Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' Then document_total Else 0 End as Purchase, 0 as Payments, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then document_total Else 0 End as DrNote, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' Then document_total Else 0 End as CrNote, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,balance_Amt as Balance_Amount, ( select   GL_Account_Code from TSPL_VENDOR_INVOICE_DETAIL where Detail_Line_No='1' and Document_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  ) as account, convert(date," & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_VENDOR_INVOICE_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VENDOR_INVOICE_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_VENDOR_INVOICE_HEAD & " where ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,'')<>''  and LEN(ISNULL( " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No,''))<=0  ) FinalWCt where (FinalWCt.CrAmt <>0 or FinalWCt.DrAmt <>0) " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine
                ''---------------------------

                strtempBaseQry += " Select Due_Date, PO_SRN, VCode, VName, DocNo, DocType, DocDate, DocNarr, ChequeDetails, Currency_Code, ConvRate, CrAmt, DrAmt, "
                If objFilter.VendorWise Or objFilter.VendorGroupWise Then
                    strtempBaseQry += "case when DocType='Pur.Invoice' then  CrAmt-DrAmt else 0 end as Purchase,"
                Else
                    strtempBaseQry += " CrAmt-DrAmt as Purchase,"
                End If
                strtempBaseQry += " 0 as Payments, 0 as drNote, 0 as CrNote, Document_Type, Balance_Amount, account, Posting_Date, GLDocType, Comp_Code,OP_Type from (" + Environment.NewLine & _
                " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date ,CASE WHEN isnull(" & TSPL_PI_HEAD & ".Against_PO,'') <> '' THEN isnull(" & TSPL_PI_HEAD & ".Against_PO,'')  ELSE ''  + isnull(" & TSPL_PI_HEAD & ".Against_SRN,'') END  AS PO_SRN ,"
                ' If rbPortrait.IsChecked = True Then
                If objFilter.strPortrait = True Then
                    strtempBaseQry += " " & TSPL_PI_HEAD & ".Vendor_Code as VCode," & TSPL_PI_HEAD & ".Vendor_Name as VName, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No as DocNo,'Pur.Invoice' as DocType,convert(date," & TSPL_PI_HEAD & ".PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' - '+'SRN No-'+" & TSPL_PI_HEAD & ".Against_SRN+ ' ,Vendor Invoice No-'+ " & TSPL_PI_HEAD & ".Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date)+ (SELECT ' & Items Are- '+STUFF((SELECT ', ' + t2.Item_Desc+'('+Convert(Varchar,COnvert(Decimal(18,2),t2.Item_Cost))+' Rs)' FROM TSPL_PI_DETAIL t2 WHERE t2.PI_No = " & TSPL_PI_HEAD & ".PI_No FOR XML PATH ('')),1,2,'')) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end) as CrAmt, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end ) as DrAmt, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type, " & TSPL_VENDOR_INVOICE_HEAD & ".Balance_Amt" + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=" & TSPL_PI_HEAD & ".Bill_To_Location) as account, convert(date," & TSPL_PI_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_PI_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PI_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PI_HEAD & "  left outer join TSPL_PJV_HEAD on " & TSPL_PI_HEAD & ".PI_No=TSPL_PJV_HEAD.Invoice_No left outer join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Invoice_No=" & TSPL_PI_HEAD & ".Vendor_Invoice_No and " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No=" & TSPL_PI_HEAD & ".PI_No where " & TSPL_PI_HEAD & ".Status=1" + Environment.NewLine
                    ' ElseIf rbLandScape.IsChecked = True Then
                ElseIf objFilter.strLandscape = True Then
                    strtempBaseQry += " " & TSPL_PI_HEAD & ".Vendor_Code as VCode," & TSPL_PI_HEAD & ".Vendor_Name as VName, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No as DocNo,'Pur.Invoice' as DocType,convert(date," & TSPL_PI_HEAD & ".PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' ,Vendor Invoice No-'+ " & TSPL_PI_HEAD & ".Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end) as CrAmt,  (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end ) as DrAmt, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type, " & TSPL_VENDOR_INVOICE_HEAD & ".Balance_Amt " + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=" & TSPL_PI_HEAD & ".Bill_To_Location) as account, convert(date," & TSPL_PI_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_PI_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PI_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PI_HEAD & "  left outer join TSPL_PJV_HEAD on " & TSPL_PI_HEAD & ".PI_No=TSPL_PJV_HEAD.Invoice_No left outer join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Invoice_No=" & TSPL_PI_HEAD & ".Vendor_Invoice_No and " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No=" & TSPL_PI_HEAD & ".PI_No where " & TSPL_PI_HEAD & ".Status=1" + Environment.NewLine
                End If
                strtempBaseQry += ") XXX"
                strtempBaseQry += " UNION ALL" + Environment.NewLine
                '  BM00000008275 BM00000008234 BM00000008238 case when "& TSPL_REMITTANCE &".Document_Type IN('I','C','OA','AV') AND ISNULL("& TSPL_VENDOR_INVOICE_HEAD &".Against_PurchaseReturn_No,'')='' then 0 else Actual_Total_TDS END, case when "& TSPL_REMITTANCE &".Document_Type IN('I','C','OA','AV') AND ISNULL("& TSPL_VENDOR_INVOICE_HEAD &".Against_PurchaseReturn_No,'')='' then Actual_Total_TDS else 0 END, 0 as Purchase, case when "& TSPL_REMITTANCE &".Document_Type IN('I','C','OA','AV') AND ISNULL("& TSPL_VENDOR_INVOICE_HEAD &".Against_PurchaseReturn_No,'')='' then -1*Actual_Total_TDS else Actual_Total_TDS END as Payments CHANGED TO " case when "& TSPL_REMITTANCE &".Document_Type ='D' AND "& TSPL_VENDOR_INVOICE_HEAD &".TDS_Actual_Amount >0 then -1* Actual_Total_TDS else Actual_Total_TDS END As Payments " as per ashok/amit sir. 26-Oct-2015
                strtempBaseQry += " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name, " & TSPL_REMITTANCE & ".Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, Case When " & TSPL_REMITTANCE & ".Document_Type IN ('I','C','D') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Description Else " & TSPL_PAYMENT_HEADER & ".Entry_Desc End as DocNarr, '', Case When " & TSPL_REMITTANCE & ".Document_Type in ('I','D','C') Then " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE Else " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE End as Currency_Code, Case When " & TSPL_REMITTANCE & ".Document_Type in ('I','D','C') Then " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate Else " & TSPL_PAYMENT_HEADER & ".ConvRate End as ConvRate, case when " & TSPL_REMITTANCE & ".Document_Type IN('I','C','OA','AV') AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' then 0 else Actual_Total_TDS END as CrAmt, case when " & TSPL_REMITTANCE & ".Document_Type IN('I','C','OA','AV') AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' then "

                strtempBaseQry += " Actual_Total_TDS "

                strtempBaseQry += " else 0 END, case when " & TSPL_REMITTANCE & ".Document_Type ='I' AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as Purchase, case when " & TSPL_REMITTANCE & ".Document_Type IN "
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry += "('I','D','C')"
                Else
                    strtempBaseQry += "('I','D')"
                End If
                strtempBaseQry += " AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then 0 else Actual_Total_TDS END as Payments, case when " & TSPL_REMITTANCE & ".Document_Type ='D' AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as DrNotes, 0 as CrNotes, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, Case When ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,'')<>'' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date Else " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date End as Posting_Date,case when " & TSPL_REMITTANCE & ".Document_Type in ('C') then 'AP-CN' when " & TSPL_REMITTANCE & ".Document_Type in ('I') then 'AP-IN' when " & TSPL_REMITTANCE & ".Document_Type in ('D') then 'AP-DN' else case when " & TSPL_REMITTANCE & ".Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end end as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " Left Outer Join " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_REMITTANCE & ".Document_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Document_No LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".is_For_TDS,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name ," & TSPL_REMITTANCE & ".Document_No ,'TDS REVERSE' as [DocType],convert(date," & TSPL_BANK_REVERSE & ".Reversal_Date,103)as Document_Date,'', '', " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when (" & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No) then Actual_Total_TDS else null end AS CrAmt, 0 as DrAmt, 0 as Purchase, case when (" & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No) then Actual_Total_TDS*-1 else 0 end as Payments, 0 as DrNote, 0 as CrNote, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No inner join " & TSPL_BANK_REVERSE & " on " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No where Remit_TDS is not null and " & TSPL_BANK_REVERSE & ".Post ='P' and Branch_GL_AC  is not null and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   AND 1=2" + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine
                If objCommonVar.IsMultiCurrencyCompany = True Then
                    strtempBaseQry += "select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No+case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate,"
                    '    " ----------- For bank reverse entry--------------- " + Environment.NewLine
                    strtempBaseQry += " case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN ('RC','AD') then 0 when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type ='D' then  TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else  TSPL_PAYMENT_DETAIL.Applied_Amount  end ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, " & _
               "case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN ('AD') then TSPL_PAYMENT_DETAIL.Applied_Amount WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type IN ('RC') then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN('OA','AV') then -1*" & TSPL_BANK_REVERSE & ".Amount when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then -1*TSPL_PAYMENT_DETAIL.Applied_Amount ELSE " & TSPL_BANK_REVERSE & ".Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code Else CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM " & TSPL_PAYMENT_HEADER & " PR WHERE PR.Payment_No =" & TSPL_PAYMENT_HEADER & ".Applied_Payment )  ELSE substring(" & TSPL_PAYMENT_HEADER & ".Debit_Account , len(" & TSPL_PAYMENT_HEADER & ".Debit_Account )-2,3) END End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & "" & _
               " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No" & _
               " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No" & _
               " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No" & _
               " left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code" & _
               " where Source_Type ='AP' And " & TSPL_BANK_REVERSE & ".Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  " & IIf(objFilter.IncludeApplyDoc = False, " and " & TSPL_PAYMENT_HEADER & ".Payment_Type<>'AD' ", "") & " " + Environment.NewLine & _
                   " ----------- For bank reverse entry--------------- " + Environment.NewLine
                Else
                    strtempBaseQry += " select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when   " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 0 ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount ELSE -1*" & TSPL_BANK_REVERSE & ".Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(BANKACC,3) End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & "  LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code where Source_Type ='AP' And Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   " + Environment.NewLine
                End If
                strtempBaseQry += " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, VC_Code as VCode, VC_Name as VName, " & TSPL_VCGL_Head & ".Document_No as DocNo,case when " & TSPL_VCGL_Head & ".Document_Type='V' then 'Vendor' else '' end as DocType,CONVERT(Date," & TSPL_VCGL_Head & ".Document_Date,103) as DocDate, " & TSPL_VCGL_Head & ".Remarks as DocNarr, '' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, (case when Amount_Type='Dr' then Amount else 0 end) as CrAmt,(case when Amount_Type='Cr' then Amount else 0 end) as DrAmt, 0 as Purchase, 0 as Payments, case when Amount_Type='Cr' then Amount else 0 end as DrNote, case when Amount_Type='Dr' then Amount else 0 end as CrNote, (case when  " & TSPL_VCGL_Head & ".Document_Type='V' then 'Vendor' else '' end ) as Document_Type ,0 as Balance_Amount, (" & TSPL_VCGL_Head & ".Location_Segment) as account, " & TSPL_VCGL_Head & ".Posting_Date as Posting_Date,'VC-GL' as GLDocType," & TSPL_VCGL_Head & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VCGL_Head & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_VCGL_Head & " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL=" & TSPL_VCGL_Head & ".Document_No where " & TSPL_VCGL_Head & ".Document_Type='v' and " & TSPL_VCGL_Head & ".Status='1' AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL,'')=''  " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " select NULL AS Due_Date, '' AS PO_SRN, TSPL_VCGL_Detail.VCGL_Code as VCode, TSPL_VCGL_Detail.VCGL_Name as VName, TSPL_VCGL_Detail.Document_No as DocNo, TSPL_VCGL_Detail.Row_Type as DocType,CONVERT(date," & TSPL_VCGL_Head & ".Document_Date,103) as DocDate," & TSPL_VCGL_Head & ".Remarks as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, TSPL_VCGL_Detail.Cr_Amount as CrAmt,TSPL_VCGL_Detail.Dr_Amount as DrAmt, 0 as Purchase, 0 as Payments, TSPL_VCGL_Detail.Dr_Amount as DrNote, TSPL_VCGL_Detail.Cr_Amount as CrNote, TSPL_VCGL_Detail.Row_Type as Document_Type,0 as Balance_Amount, (" & TSPL_VCGL_Head & ".Location_Segment) as account, " & TSPL_VCGL_Head & ".Posting_Date as Posting_Date ,'VC-GL' as GLDocType, " & TSPL_VCGL_Head & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VCGL_Head & "." & OP_TypeCol, OP_TypeCol) & " from  TSPL_VCGL_Detail left outer join  " & TSPL_VCGL_Head & " on " & TSPL_VCGL_Head & ".Document_No=TSPL_VCGL_Detail.Document_No LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL=" & TSPL_VCGL_Head & ".Document_No where Row_Type='Vendor'  and " & TSPL_VCGL_Head & ".Status='1' AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL,'')=''" + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine

                strtempBaseQry += "  select NULL AS Due_Date, '' AS PO_SRN, " & TSPL_Payment_Adjustment_Header & ".Vendor_No as VCode, " & TSPL_Payment_Adjustment_Header & ".Vendor_Name as VName, " & TSPL_Payment_Adjustment_Header & ".Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date," & TSPL_Payment_Adjustment_Header & ".Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +" & TSPL_Payment_Adjustment_Header & ".Doc_No as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, CASE WHEN ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,'')='D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END as CrAmt, CASE WHEN ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,'')<>'D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END  as DrAmt, 0 as Purchase, 0 as Payments, TSPL_Payment_Adjustment_Detail.Amount as DrNote, 0 as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account, " & TSPL_Payment_Adjustment_Header & ".Post_Date as Posting_Date ,'' as GLDocType, " & TSPL_Payment_Adjustment_Header & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_Payment_Adjustment_Header & "." & OP_TypeCol, OP_TypeCol) & " from  TSPL_Payment_Adjustment_Detail  left outer join  " & TSPL_Payment_Adjustment_Header & " on " & TSPL_Payment_Adjustment_Header & ".Adjustment_No =TSPL_Payment_Adjustment_Detail .Adjustment_No  LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  =" & TSPL_Payment_Adjustment_Header & ".Doc_No  where isnull(" & TSPL_Payment_Adjustment_Header & ".Is_Post,'') ='Y' and " & TSPL_Payment_Adjustment_Header & ".Adjust_Type='P'  " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine
                strtempBaseQry += "  select NULL AS Due_Date, '' AS PO_SRN, " & TSPL_Payment_Adjustment_Header & ".Vendor_No as VCode, " & TSPL_Payment_Adjustment_Header & ".Vendor_Name as VName, " & TSPL_Payment_Adjustment_Header & ".Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date," & TSPL_Payment_Adjustment_Header & ".Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +" & TSPL_Payment_Adjustment_Header & ".Doc_No as DocNarr,'' as ChequeDetails, 'INR' as CURRENCY_CODE, 1 as ConvRate, CASE WHEN ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,'')<>'D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END as CrAmt,  CASE WHEN ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,'')='D' THEN  TSPL_Payment_Adjustment_Detail.Amount ELSE 0 END as DrAmt, 0 as Purchase, 0 as Payments, 0 as DrNote, TSPL_Payment_Adjustment_Detail.Amount as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account, " & TSPL_Payment_Adjustment_Header & ".Post_Date as Posting_Date ,'' as GLDocType, " & TSPL_Payment_Adjustment_Header & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_Payment_Adjustment_Header & "." & OP_TypeCol, OP_TypeCol) & " from  TSPL_Payment_Adjustment_Detail  left outer join  " & TSPL_Payment_Adjustment_Header & " on " & TSPL_Payment_Adjustment_Header & ".Adjustment_No =TSPL_Payment_Adjustment_Detail .Adjustment_No  LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  =" & TSPL_Payment_Adjustment_Header & ".Doc_No  where isnull(" & TSPL_Payment_Adjustment_Header & ".Is_Post,'') ='Y' and " & TSPL_Payment_Adjustment_Header & ".Adjust_Type='R'  " + Environment.NewLine & _
 " UNION ALL" + Environment.NewLine & _
                " Select NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code,Max(OP_Type) as OP_Type from (" + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='AV' then 'Advance' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='OA' then 'On Account' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then 'Payment' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".payment_type IN('OA','AV') then Payment_Amount When " & TSPL_PAYMENT_HEADER & ".payment_type IN('PY') Then (Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No Where " & TSPL_PAYMENT_HEADER & ".Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   " + Environment.NewLine & _
                " ) XX Group By XX.account, XX.DocNo"
                If objFilter.DocumentWise = False Then
                    strtempBaseQry += " UNION ALL" + Environment.NewLine & _
        " select null AS Due_Date, '' AS PO_SRN, TSPL_VENDOR_MASTER.Vendor_Code as VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName," & TSPL_PAYMENT_HEADER & ".Payment_No as DocNo,'EXC' as [DocType],convert(date,Payment_Date, 103) as DocDate, case when ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)=1 Then 'Security' Else '' + 'Exchange Gain/Loss against Journal Entry No  '+isnull(TSPL_JOURNAL_MASTER.Voucher_No,'')    end as DocNarr,(" & TSPL_PAYMENT_HEADER & ".cheque_No+'-'+ CONVERT(VARCHAR,Cheque_Date , 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, " & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT AS CrAmt,  " & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT as DrAmt, 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, 0 as Balance_Amount,isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code ,'') as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PAYMENT_HEADER & "  LEFT OUTER JOIN TSPL_Vendor_Master ON TSPL_Vendor_Master.Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code " + Environment.NewLine & _
        " left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code" + Environment.NewLine & _
        " left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =" & TSPL_PAYMENT_HEADER & ".Payment_No " + Environment.NewLine & _
        " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No and TSPL_PAYMENT_DETAIL.Payment_Line_No =1 " + Environment.NewLine & _
        " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No  " + Environment.NewLine & _
        " where " & TSPL_PAYMENT_HEADER & ".Posted=1 and  " & TSPL_PAYMENT_HEADER & ".Payment_Type  in ('PY','AD') " + Environment.NewLine & _
        " AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1 and (" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT<>0 or " & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT  <>0)  " + Environment.NewLine & _
        " ---------------------to find gain or loss amount for payment---------------" + Environment.NewLine & _
        " UNION ALL" + Environment.NewLine & _
        " Select null AS Due_Date, '' AS PO_SRN,TSPL_VENDOR_MASTER.Vendor_Code as VCode, TSPL_VENDOR_MASTER.Vendor_Name as VName, " & TSPL_BANK_REVERSE & ".Reverse_Code as DocNo," + Environment.NewLine & _
        " 'EXC'  as DocType  ,convert(date," & TSPL_BANK_REVERSE & ".Reversal_Date,103) as DocDate,case when ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security ,0)=1 Then 'Security' Else '' + 'Exchange Gain/Loss against Journal Entry No  '+isnull(TSPL_JOURNAL_MASTER.Voucher_No,'')    end as DocNarr,(rtrim(Entry_Desc) + (case when len(RTRIM(Entry_Desc))>0 and len(RTRIM(" & TSPL_PAYMENT_HEADER & ".Cheque_No))>0  then ' / ' else '' end)  + (case when LEN(ISNULL(" & TSPL_BANK_REVERSE & ".Cheque_No,''))>0 then 'Cheque No. - ' + " & TSPL_BANK_REVERSE & ".Cheque_No +  ' - '+ convert(varchar ," & TSPL_PAYMENT_HEADER & " . Cheque_Date ,103)else '' end)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".Currency_Code, 1 as ConvRate,  " & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT AS CrAmt, " & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT as DrAmt,  0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," + Environment.NewLine & _
        " 'RV',0, isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code ,'') as account, " & TSPL_PAYMENT_HEADER & ".Payment_Date as Posting_Date,'RV-TA' as GLDocType," & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " " + Environment.NewLine & _
        " from " & TSPL_BANK_REVERSE & "  LEFT OUTER JOIN TSPL_VENDOR_MASTER  ON TSPL_VENDOR_MASTER.Vendor_Code =" & TSPL_BANK_REVERSE & ".Vendor_Code   " + Environment.NewLine & _
        " LEFT OUTER JOIN TSPL_Vendor_TYPE_MASTER  ON TSPL_Vendor_TYPE_MASTER.Ven_Type_Code  = TSPL_VENDOR_MASTER.Ven_Type_Code   " + Environment.NewLine & _
        " left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =" & TSPL_BANK_REVERSE & ".Reverse_Code  " + Environment.NewLine & _
        " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".PAYMENT_NO =" & TSPL_BANK_REVERSE & ".Document_No " + Environment.NewLine & _
        " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No and TSPL_PAYMENT_DETAIL.Payment_Line_No =1 " + Environment.NewLine & _
        " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No " + Environment.NewLine & _
        " where  " & TSPL_BANK_REVERSE & ".Reverse_Document='Payments' AND " & TSPL_BANK_REVERSE & ".Post ='P'" + Environment.NewLine & _
        " and (" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT<>0 or " & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT  <>0) " + Environment.NewLine & _
        " ----------------------to find gain or loss amount FOR BANK REVERSE ---------------" + Environment.NewLine

                End If

                If objFilter.IncludeApplyDoc Then
                    strtempBaseQry += " UNION ALL" + Environment.NewLine & _
                    " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS CrAmt , 0  AS  DrAmt ,0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," + Environment.NewLine & _
                    " " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount,  CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM " & TSPL_PAYMENT_HEADER & " PR WHERE PR.Payment_No =" & TSPL_PAYMENT_HEADER & ".Applied_Payment )  ELSE substring(" & TSPL_PAYMENT_HEADER & ".Debit_Account , len(" & TSPL_PAYMENT_HEADER & ".Debit_Account )-2,3) END as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " " + Environment.NewLine & _
                    " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No Where " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  " + Environment.NewLine & _
                    " --------------- INVOICE AGAINST APPLY DOCUMENT" + Environment.NewLine & _
                    " UNION ALL" + Environment.NewLine & _
                    " SELECT NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code,Max(OP_Type) as OP_Type FROM (select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo,CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, 0  AS CrAmt,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS  DrAmt," + Environment.NewLine & _
                    " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, " + Environment.NewLine & _
                    " " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & "  " + Environment.NewLine & _
                    " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No Where " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  " + Environment.NewLine & _
                    " )INV  GROUP BY  DocNo,account  " + Environment.NewLine & _
                    " ------- APPLY DOCUMENT ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION " + Environment.NewLine & _
                    " UNION ALL " + Environment.NewLine & _
                    " SELECT NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code,Max(OP_Type) as OP_Type FROM ( " + Environment.NewLine & _
                    " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_BANK_REVERSE & ".Reverse_Code as DocNo,'Reverse Payment' as DocType , Convert(date," & TSPL_BANK_REVERSE & ".Reversal_Date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((" & TSPL_PAYMENT_HEADER & ".Cheque_No) + (case when " & TSPL_PAYMENT_HEADER & ".Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type  in ('D') then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end  AS CrAmt,0 AS  DrAmt, " + Environment.NewLine & _
                    " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount,  " + Environment.NewLine & _
                    " " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & "  " + Environment.NewLine & _
                    " from " & TSPL_BANK_REVERSE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No =" & TSPL_BANK_REVERSE & ".Document_No  LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No Where" + Environment.NewLine & _
                    " " & TSPL_BANK_REVERSE & ".Reverse_Document='Payments' AND " & TSPL_BANK_REVERSE & ".Post ='P' AND " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1" + Environment.NewLine & _
                    " )INV  GROUP BY  DocNo,account " + Environment.NewLine & _
                    " -------------- FOR BANK REVERSE ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION" + Environment.NewLine
                End If


            Else
                strtempBaseQry = "Select NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code,Max(OP_Type) as OP_Type from (" + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='AV' then 'Advance' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='OA' then 'On Account' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then 'Payment' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".payment_type IN('OA','AV') then Payment_Amount When " & TSPL_PAYMENT_HEADER & ".payment_type IN('PY') Then (Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No Where " & TSPL_PAYMENT_HEADER & ".Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   " + Environment.NewLine & _
                " ) XX Group By XX.account, XX.DocNo" + Environment.NewLine & _
            " UNION ALL" + Environment.NewLine & _
                " select " & TSPL_PAYMENT_HEADER & ".Payment_Date As Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name, " & TSPL_REMITTANCE & ".Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, " & TSPL_PAYMENT_HEADER & ".Entry_Desc as DocNarr, '', " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE as Currency_Code, " & TSPL_PAYMENT_HEADER & ".ConvRate as ConvRate, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then 0 else Actual_Total_TDS END as CrAmt, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then Actual_Total_TDS else 0 END as DrAmt, 0 as Purchase, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then Actual_Total_TDS else Actual_Total_TDS END as Payments, 0 as DrNotes, 0 as CrNotes, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_REMITTANCE & ".Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1" + Environment.NewLine & _
            " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No+case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 0 when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount*-1 ELSE " & TSPL_BANK_REVERSE & ".Amount*-1 end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code Else right(BANKACC,3) End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code where Source_Type ='AP' And " & TSPL_BANK_REVERSE & ".Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  "

            End If
            objFilter.strtempBaseQryforopening = strtempBaseQry
            '  Dim strBaseQryforVendor As String = String.Empty
            ''BM00000008527
            If clsCommon.CompairString(clsCommon.myCstr(objFilter.CurrencyType), "") = CompairStringResult.Equal Then
                'LoadCurrencyType()
            End If
            If objFilter.DocumentWise = True Then
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine & _
                    " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
                    "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
                    "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine & _
                    " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) else  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
                    " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then   (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end else " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
                    " case when InnQuery.DocType='Pur.Invoice' then InnQuery.CrAmt-InnQuery.DrAmt else 0 end as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine & _
                    "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_M1ASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  "
                Else
                    strtempBaseQry = "  Select isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) as EXCHANGE_GAIN_AMT,isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) as EXCHANGE_LOSS_AMT, InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) - case when (DocType)='Reverse Payment' then  (Select isnull((PH.EXCHANGE_LOSS_AMT  ),0) from " & TSPL_PAYMENT_HEADER & " PH where PH.Payment_No =(  Select Document_No   from " & TSPL_BANK_REVERSE & " where Reverse_Code =InnQuery.DocNo )) else 0 end as  CrAmt ," + Environment.NewLine & _
                   " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
                   "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
                   "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine & _
                   " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) ELSE  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
                   " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then CASE WHEN ((DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")-isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0))<0 THEN 0 ELSE (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")-isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) END  else CASE WHEN (DocType)<>'IM' THEN (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) ELSE (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") END end else " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
                   " InnQuery.CrAmt-InnQuery.DrAmt  as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine & _
                   "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  "
                End If
            Else
                strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  "
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry += " CONVERT(DECIMAL(18,2), CASE WHEN InnQuery.DocType NOT IN ('Pur.Invoice','Receipt')  THEN case when InnQuery.DocType not in ('TDS','AP Invoice') THEN InnQuery.CrAmt  *(case when (DocType) NOT IN ('EXC','Credit Note','IM','Reverse Payment') then   InnQuery.ConvRate else 1 end) when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo)='C' then InnQuery.CrAmt ELSE CASE WHEN InnQuery.DocType in ('TDS') and (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  )='I' or (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo   )='C' then (Select ISNULL (Actual_Total_TDS ,0)  from " & TSPL_REMITTANCE & " where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  ) * -1 else 0 end END ELSE 0 END) as CrAmt ," & _
                      " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Adjustment') then case when InnQuery.DocType not in ('AP Invoice','TDS') then InnQuery.DrAmt WHEN InnQuery.DocType in ('TDS') and   (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  )='D' then (Select ISNULL (Actual_Total_TDS ,0)  from " & TSPL_REMITTANCE & " where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  ) * -1 else case when  InnQuery.DocType in ('AP Invoice') AND  (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')='')='D' then InnQuery.DrAmt else 0 end end else 0 end) DrAmt, " & _
                      " CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Debit Note','Reverse Payment','Credit Note','EXC','IM') then case when InnQuery.DocType not in ('AP Invoice','TDS') then CONVERT(DECIMAL(18,2),InnQuery.Purchase) else case when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo)='I'	 then CONVERT(DECIMAL(18,2),InnQuery.Purchase) when  InnQuery.DocType in ('AP Invoice')  and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')<>'')='D' then CONVERT(DECIMAL(18,2),InnQuery.Purchase) *-1  else 0 end end else 0 end) Purchase, " & _
                      " CONVERT(DECIMAL(18,2), case when InnQuery.DocType  in ('Payment','On Account','Advance','Receipt','TDS','EXC','IM') then InnQuery.Payments * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) else 0 end) as Payments , "
                Else
                    strtempBaseQry += " CONVERT(DECIMAL(18,2),InnQuery.CrAmt) AS CrAmt,CONVERT(DECIMAL(18,2),InnQuery.DrAmt) AS DrAmt,CONVERT(DECIMAL(18,2),InnQuery.Purchase) AS Purchase,CONVERT(DECIMAL(18,2),InnQuery.Payments) AS Payments , "
                End If

                strtempBaseQry += " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE from " & _
               "(Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end)  as  CrAmt ," + Environment.NewLine & _
               " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
               "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
               "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
               "case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine & _
               "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine & _
               "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
               "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
               "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
               " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & objFilter.strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine & _
               "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code) InnQuery  "


                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    objFilter.strtempBaseQryforopeningForMIS = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  " & _
                  " InnQuery.CrAmt,InnQuery.DrAmt,0 as Purchase,0 as Payments, " & _
                  " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE from " & _
                  " (Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  *(case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine & _
                  " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
                  " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
                  " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
                  " case when isnull((DrAmt),0)>0 AND (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else isnull((DrAmt),0) end end else " + Environment.NewLine & _
                  " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else case when (DocType)<>'EXC' then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") else CASE WHEN  isnull((DrAmt),0)>0 THEN isnull(DrAmt,0) ELSE isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT  ),0) END end  end end else " + Environment.NewLine & _
                  " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
                  "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND  (DocType)<>'EXC' then   (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else case when (DocType)<>'EXC' then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
                  "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
                  " CONVERT(DECIMAL(18,2),InnQuery.CrAmt-InnQuery.DrAmt) as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & objFilter.strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine & _
                  "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code) InnQuery  "
                End If

            End If
            objFilter.strtempBaseQryforopening = GetVendorOpening_GIT(Nothing, Today) ' strtempBaseQry
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return strtempBaseQry

    End Function
    Public Function GetVendorBaseQryForVendorCurrency(ByRef objFilter As structVendorFilter) As String
        Dim strtempBaseQry As String = String.Empty
        Dim TSPL_VENDOR_INVOICE_HEAD As String = "TSPL_VENDOR_INVOICE_HEAD"
        Dim TSPL_PAYMENT_HEADER = "TSPL_PAYMENT_HEADER"
        Dim TSPL_Payment_Adjustment_Header = "TSPL_Payment_Adjustment_Header"
        Dim TSPL_BANK_REVERSE As String = "TSPL_BANK_REVERSE"
        Dim TSPL_REMITTANCE As String = "TSPL_REMITTANCE"
        Dim TSPL_VCGL_Head As String = "TSPL_VCGL_Head"
        Dim TSPL_PI_HEAD As String = "TSPL_PI_HEAD"
        Dim OP_TypeCol As String = "'' as OP_Type "
        If objFilter.IS_GIT Then
            TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD_WIN"
            TSPL_PAYMENT_HEADER = "TSPL_PAYMENT_HEADER_WIN"
            TSPL_Payment_Adjustment_Header = "TSPL_Payment_Adjustment_Header_WIN"
            TSPL_BANK_REVERSE = "TSPL_BANK_REVERSE_WIN"
            TSPL_REMITTANCE = "TSPL_REMITTANCE_WIN"
            TSPL_VCGL_Head = "TSPL_VCGL_Head_WIN"
            TSPL_PI_HEAD = "TSPL_PI_HEAD_WIN"
            OP_TypeCol = " OP_Type"
        End If
        Try
            If Not objFilter.IsOnlyForAgainstSalary Then
                Dim strQryForRMDA As String = " - (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' and LEN(ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No,''))>0 then (Select sum(TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt) from TSPL_TRANSFER_ORDER_HEAD where Status=1 and TSPL_TRANSFER_ORDER_HEAD.RMDA_Code in (select TSPL_SRN_HEAD.RMDA_No from TSPL_PI_DETAIL left outer join " & TSPL_PI_HEAD & " on " & TSPL_PI_HEAD & ".PI_No=TSPL_PI_DETAIL.PI_No left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=" & TSPL_PI_HEAD & ".Against_SRN where " & TSPL_PI_HEAD & ".PI_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No  )) else 0 end )"
                strtempBaseQry = " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date, '' AS PO_SRN, " & TSPL_VENDOR_INVOICE_HEAD & ".vendor_code as VCode," & TSPL_VENDOR_INVOICE_HEAD & ".vendor_name as VName," & TSPL_VENDOR_INVOICE_HEAD & " .Document_No as DocNo ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'Debit Note' else case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'Credit Note' else 'AP Invoice' end  end as DocType,convert(date,Invoice_Entry_Date,103) as DocDate,((Description) + (case when Description='' then case when " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks <>'' then " & TSPL_VENDOR_INVOICE_HEAD & ".Remarks +' - ' else ''  end else ' - 'end) +(RefDocNo)+'Vendor Invoice No-' +Vendor_Invoice_No+' Date-' +Vendor_Invoice_Date  ) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate, case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type IN('I','C') Then document_total else 0 end as CrAmt, case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type IN('D') then Document_Total else 0 end as DrAmt, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' Then document_total Else 0 End as Purchase, 0 as Payments, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then document_total Else 0 End as DrNote, Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' Then document_total Else 0 End as CrNote, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type,balance_Amt as Balance_Amount, ( select   GL_Account_Code from TSPL_VENDOR_INVOICE_DETAIL where Detail_Line_No='1' and Document_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  ) as account, convert(date," & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_VENDOR_INVOICE_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VENDOR_INVOICE_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_VENDOR_INVOICE_HEAD & " left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Code where ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,'')<>''  and LEN(ISNULL( " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No,''))<=0 AND   " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  " + Environment.NewLine & _
                    " UNION ALL" + Environment.NewLine & _
                 " Select Due_Date, PO_SRN, VCode, VName, DocNo, DocType, DocDate, DocNarr, ChequeDetails, Currency_Code, ConvRate, CrAmt, DrAmt, "
                If objFilter.VendorWise Or objFilter.VendorGroupWise Then
                    strtempBaseQry += "case when DocType='Pur.Invoice' then  CrAmt-DrAmt else 0 end as Purchase,"
                Else
                    strtempBaseQry += " CrAmt-DrAmt as Purchase,"
                End If
                strtempBaseQry += " 0 as Payments, 0 as drNote, 0 as CrNote, Document_Type, Balance_Amount, account, Posting_Date, GLDocType, Comp_Code,OP_Type from (" + Environment.NewLine & _
                " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date ,CASE WHEN isnull(" & TSPL_PI_HEAD & ".Against_PO,'') <> '' THEN isnull(" & TSPL_PI_HEAD & ".Against_PO,'')  ELSE ''  + isnull(" & TSPL_PI_HEAD & ".Against_SRN,'') END  AS PO_SRN ,"
                If objFilter.strPortrait = True Then
                    strtempBaseQry += " " & TSPL_PI_HEAD & ".Vendor_Code as VCode," & TSPL_PI_HEAD & ".Vendor_Name as VName, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No as DocNo,'Pur.Invoice' as DocType,convert(date," & TSPL_PI_HEAD & ".PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' - '+'SRN No-'+" & TSPL_PI_HEAD & ".Against_SRN+ ' ,Vendor Invoice No-'+ " & TSPL_PI_HEAD & ".Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date)+ (SELECT ' & Items Are- '+STUFF((SELECT ', ' + t2.Item_Desc+'('+Convert(Varchar,COnvert(Decimal(18,2),t2.Item_Cost))+' Rs)' FROM TSPL_PI_DETAIL t2 WHERE t2.PI_No = " & TSPL_PI_HEAD & ".PI_No FOR XML PATH ('')),1,2,'')) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end) as CrAmt, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end ) as DrAmt, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type, " & TSPL_VENDOR_INVOICE_HEAD & ".Balance_Amt" + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=" & TSPL_PI_HEAD & ".Bill_To_Location) as account, convert(date," & TSPL_PI_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_PI_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PI_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PI_HEAD & "  left outer join TSPL_PJV_HEAD on " & TSPL_PI_HEAD & ".PI_No=TSPL_PJV_HEAD.Invoice_No left outer join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Invoice_No=" & TSPL_PI_HEAD & ".Vendor_Invoice_No and " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No=" & TSPL_PI_HEAD & ".PI_No where " & TSPL_PI_HEAD & ".Status=1" + Environment.NewLine
                ElseIf objFilter.strLandscape = True Then
                    strtempBaseQry += " " & TSPL_PI_HEAD & ".Vendor_Code as VCode," & TSPL_PI_HEAD & ".Vendor_Name as VName, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No as DocNo,'Pur.Invoice' as DocType,convert(date," & TSPL_PI_HEAD & ".PI_Date,103) as DocDate,(TSPL_PJV_HEAD.PJV_No+' ,Vendor Invoice No-'+ " & TSPL_PI_HEAD & ".Vendor_Invoice_No +' Date-' +Vendor_Invoice_Date) as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE, " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate, (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='I' then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end) as CrAmt,  (case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Total else 0 end ) as DrAmt, " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type, " & TSPL_VENDOR_INVOICE_HEAD & ".Balance_Amt " + strQryForRMDA + " as Balance_Amount, ( select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code=" & TSPL_PI_HEAD & ".Bill_To_Location) as account, convert(date," & TSPL_PI_HEAD & ".Posting_Date,103) as Posting_Date ,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' then 'AP-DN' when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='C' then 'AP-CN' else 'AP-IN' end as GLDocType, " & TSPL_PI_HEAD & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PI_HEAD & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PI_HEAD & "  left outer join TSPL_PJV_HEAD on " & TSPL_PI_HEAD & ".PI_No=TSPL_PJV_HEAD.Invoice_No left outer join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Invoice_No=" & TSPL_PI_HEAD & ".Vendor_Invoice_No and " & TSPL_VENDOR_INVOICE_HEAD & ".Against_POInvoice_No=" & TSPL_PI_HEAD & ".PI_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Code where " & TSPL_PI_HEAD & ".Status=1 AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'" + Environment.NewLine
                End If
                strtempBaseQry += ") XXX" + Environment.NewLine & _
                 " UNION ALL" + Environment.NewLine & _
                 " select " & TSPL_VENDOR_INVOICE_HEAD & ".Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name, " & TSPL_REMITTANCE & ".Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, Case When " & TSPL_REMITTANCE & ".Document_Type IN ('I','C','D') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Description Else " & TSPL_PAYMENT_HEADER & ".Entry_Desc End as DocNarr, '', Case When " & TSPL_REMITTANCE & ".Document_Type in ('I','D','C') Then " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE Else " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE End as Currency_Code, Case When " & TSPL_REMITTANCE & ".Document_Type in ('I','D','C') Then " & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate Else " & TSPL_PAYMENT_HEADER & ".ConvRate End as ConvRate, case when " & TSPL_REMITTANCE & ".Document_Type IN('I','C','OA','AV') AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' then 0 else Actual_Total_TDS END as CrAmt, case when " & TSPL_REMITTANCE & ".Document_Type IN('I','C','OA','AV') AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' then "
                'If clsCommon.CompairString(FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                '    strtempBaseQry += " 0 "
                'Else
                '    strtempBaseQry += " Actual_Total_TDS "
                'End If
                strtempBaseQry += " Actual_Total_TDS "
                strtempBaseQry += " else 0 END, case when " & TSPL_REMITTANCE & ".Document_Type ='I' AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as Purchase, case when " & TSPL_REMITTANCE & ".Document_Type IN "
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry += "('I','D','C')"
                Else
                    strtempBaseQry += "('I','D')"
                End If
                strtempBaseQry += " AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then 0 else Actual_Total_TDS END as Payments, case when " & TSPL_REMITTANCE & ".Document_Type ='D' AND " & TSPL_VENDOR_INVOICE_HEAD & ".TDS_Actual_Amount >0 then -1*Actual_Total_TDS else 0 END as DrNotes, 0 as CrNotes, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, Case When ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date,'')<>'' Then " & TSPL_VENDOR_INVOICE_HEAD & ".Posting_Date Else " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date End as Posting_Date,case when " & TSPL_REMITTANCE & ".Document_Type in ('C') then 'AP-CN' when " & TSPL_REMITTANCE & ".Document_Type in ('I') then 'AP-IN' when " & TSPL_REMITTANCE & ".Document_Type in ('D') then 'AP-DN' else case when " & TSPL_REMITTANCE & ".Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end end as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " Left Outer Join " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_REMITTANCE & ".Document_No=" & TSPL_VENDOR_INVOICE_HEAD & ".Document_No LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_VENDOR_INVOICE_HEAD & ".Vendor_Code where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".is_For_TDS,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1 AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'   " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name ," & TSPL_REMITTANCE & ".Document_No ,'TDS REVERSE' as [DocType],convert(date," & TSPL_BANK_REVERSE & ".Reversal_Date,103)as Document_Date,'', '', " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when (" & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No) then Actual_Total_TDS else null end AS CrAmt, 0 as DrAmt, 0 as Purchase, case when (" & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No) then Actual_Total_TDS*-1 else 0 end as Payments, 0 as DrNote, 0 as CrNote, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No inner join " & TSPL_BANK_REVERSE & " on " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_BANK_REVERSE & ".Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_REMITTANCE & ".Vendor_Code where Remit_TDS is not null and " & TSPL_BANK_REVERSE & ".Post ='P' and Branch_GL_AC  is not null and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine
                If objCommonVar.IsMultiCurrencyCompany = True Then
                    strtempBaseQry += "select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No+case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate," & _
                    " case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN ('RC','AD') then 0 when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type  in ('D') then  TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else  TSPL_PAYMENT_DETAIL.Applied_Amount  end ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, " & _
                "case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN ('RC','AD') then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type IN('OA','AV') then -1*" & TSPL_BANK_REVERSE & ".Amount when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then -1*TSPL_PAYMENT_DETAIL.Applied_Amount ELSE " & TSPL_BANK_REVERSE & ".Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code Else CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM " & TSPL_PAYMENT_HEADER & " PR WHERE PR.Payment_No =" & TSPL_PAYMENT_HEADER & ".Applied_Payment )  ELSE substring(" & TSPL_PAYMENT_HEADER & ".Debit_Account , len(" & TSPL_PAYMENT_HEADER & ".Debit_Account )-2,3) END End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & "" & _
                " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No" & _
                " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No" & _
                " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No" & _
                " left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code " & _
                " left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_BANK_REVERSE & ".Vendor_Code" & _
                " where Source_Type ='AP' And " & TSPL_BANK_REVERSE & ".Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' " & IIf(objFilter.IncludeApplyDoc = False, " and " & TSPL_PAYMENT_HEADER & ".Payment_Type<>'AD' ", "") & " " + Environment.NewLine & _
                " --------- for bank reverse entry transactions --------------- " + Environment.NewLine
                Else
                    strtempBaseQry += " select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when   " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 0 ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount ELSE -1*" & TSPL_BANK_REVERSE & ".Amount end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(BANKACC,3) End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & "  LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_BANK_REVERSE & ".Vendor_Code where Source_Type ='AP' And Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'   " + Environment.NewLine
                End If
                strtempBaseQry += " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, VC_Code as VCode, VC_Name as VName, " & TSPL_VCGL_Head & ".Document_No as DocNo,case when " & TSPL_VCGL_Head & ".Document_Type='V' then 'Vendor' else '' end as DocType,CONVERT(Date," & TSPL_VCGL_Head & ".Document_Date,103) as DocDate, " & TSPL_VCGL_Head & ".Remarks as DocNarr, '' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE as CURRENCY_CODE," & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate as ConvRate, (case when Amount_Type='Dr' then Amount else 0 end) as CrAmt,(case when Amount_Type='Cr' then Amount else 0 end) as DrAmt, 0 as Purchase, 0 as Payments, case when Amount_Type='Cr' then Amount else 0 end as DrNote, case when Amount_Type='Dr' then Amount else 0 end as CrNote, (case when  " & TSPL_VCGL_Head & ".Document_Type='V' then 'Vendor' else '' end ) as Document_Type ,0 as Balance_Amount, (" & TSPL_VCGL_Head & ".Location_Segment) as account, " & TSPL_VCGL_Head & ".Posting_Date as Posting_Date,'VC-GL' as GLDocType," & TSPL_VCGL_Head & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_VCGL_Head & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_VCGL_Head & " LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL=" & TSPL_VCGL_Head & ".Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_VCGL_Head & ".VC_Code where " & TSPL_VCGL_Head & ".Document_Type='v' and " & TSPL_VCGL_Head & ".Status='1' AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL,'')=''  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " select NULL AS Due_Date, '' AS PO_SRN, TSPL_VCGL_Detail.VCGL_Code as VCode, TSPL_VCGL_Detail.VCGL_Name as VName, TSPL_VCGL_Detail.Document_No as DocNo, TSPL_VCGL_Detail.Row_Type as DocType,CONVERT(date," & TSPL_VCGL_Head & ".Document_Date,103) as DocDate," & TSPL_VCGL_Head & ".Remarks as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE as CURRENCY_CODE," & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate as ConvRate, TSPL_VCGL_Detail.Cr_Amount as CrAmt,TSPL_VCGL_Detail.Dr_Amount as DrAmt, 0 as Purchase, 0 as Payments, TSPL_VCGL_Detail.Dr_Amount as DrNote, TSPL_VCGL_Detail.Cr_Amount as CrNote, TSPL_VCGL_Detail.Row_Type as Document_Type,0 as Balance_Amount, (" & TSPL_VCGL_Head & ".Location_Segment) as account, " & TSPL_VCGL_Head & ".Posting_Date as Posting_Date ,'VC-GL' as GLDocType, " & TSPL_VCGL_Head & ".Comp_Code from  TSPL_VCGL_Detail left outer join  " & TSPL_VCGL_Head & " on " & TSPL_VCGL_Head & ".Document_No=TSPL_VCGL_Detail.Document_No LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL=" & TSPL_VCGL_Head & ".Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_VCGL_Head & ".VC_Code where Row_Type='Vendor'  and " & TSPL_VCGL_Head & ".Status='1' AND ISNULL(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_VCGL,'')=''  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                "  select NULL AS Due_Date, '' AS PO_SRN, " & TSPL_Payment_Adjustment_Header & ".Vendor_No as VCode, " & TSPL_Payment_Adjustment_Header & ".Vendor_Name as VName, " & TSPL_Payment_Adjustment_Header & ".Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date," & TSPL_Payment_Adjustment_Header & ".Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +" & TSPL_Payment_Adjustment_Header & ".Doc_No as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE as CURRENCY_CODE," & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate as ConvRate, 0 as CrAmt,TSPL_Payment_Adjustment_Detail.Amount  as DrAmt, 0 as Purchase, 0 as Payments, TSPL_Payment_Adjustment_Detail.Amount as DrNote, 0 as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account, " & TSPL_Payment_Adjustment_Header & ".Post_Date as Posting_Date ,'' as GLDocType, " & TSPL_Payment_Adjustment_Header & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_Payment_Adjustment_Header & "." & OP_TypeCol, OP_TypeCol) & " from  TSPL_Payment_Adjustment_Detail  left outer join  " & TSPL_Payment_Adjustment_Header & " on " & TSPL_Payment_Adjustment_Header & ".Adjustment_No =TSPL_Payment_Adjustment_Detail .Adjustment_No  LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  =" & TSPL_Payment_Adjustment_Header & ".Doc_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_Payment_Adjustment_Header & ".Vendor_No where isnull(" & TSPL_Payment_Adjustment_Header & ".Is_Post,'') ='Y'  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' and " & TSPL_Payment_Adjustment_Header & ".Adjust_Type='P' " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                "  select NULL AS Due_Date, '' AS PO_SRN, " & TSPL_Payment_Adjustment_Header & ".Vendor_No as VCode, " & TSPL_Payment_Adjustment_Header & ".Vendor_Name as VName, " & TSPL_Payment_Adjustment_Header & ".Adjustment_No as DocNo, 'Adjustment' as DocType,CONVERT(date," & TSPL_Payment_Adjustment_Header & ".Adjustment_Date,103) as DocDate,'Against Vendor Invoice No. ' +" & TSPL_Payment_Adjustment_Header & ".Doc_No as DocNarr,'' as ChequeDetails, " & TSPL_VENDOR_INVOICE_HEAD & ".CURRENCY_CODE as CURRENCY_CODE," & TSPL_VENDOR_INVOICE_HEAD & ".ConvRate as ConvRate, TSPL_Payment_Adjustment_Detail.Amount as CrAmt,0  as DrAmt, 0 as Purchase, 0 as Payments, 0 as DrNote, TSPL_Payment_Adjustment_Detail.Amount as CrNote, 'PAE' as Document_Type,0 as Balance_Amount, " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account, " & TSPL_Payment_Adjustment_Header & ".Post_Date as Posting_Date ,'' as GLDocType, " & TSPL_Payment_Adjustment_Header & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_Payment_Adjustment_Header & "." & OP_TypeCol, OP_TypeCol) & " from  TSPL_Payment_Adjustment_Detail  left outer join  " & TSPL_Payment_Adjustment_Header & " on " & TSPL_Payment_Adjustment_Header & ".Adjustment_No =TSPL_Payment_Adjustment_Detail .Adjustment_No  LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No  =" & TSPL_Payment_Adjustment_Header & ".Doc_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_Payment_Adjustment_Header & ".Vendor_No where isnull(" & TSPL_Payment_Adjustment_Header & ".Is_Post,'') ='Y'  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' and " & TSPL_Payment_Adjustment_Header & ".Adjust_Type='R' " + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " Select NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code,Max(OP_Type) as OP_Type from (" + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='AV' then 'Advance' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='OA' then 'On Account' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then 'Payment' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".payment_type IN('OA','AV') then Payment_Amount When " & TSPL_PAYMENT_HEADER & ".payment_type IN('PY') Then (Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code Where " & TSPL_PAYMENT_HEADER & ".Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  " + Environment.NewLine & _
                " ) XX Group By XX.account, XX.DocNo" + Environment.NewLine

                If objFilter.IncludeApplyDoc Then
                    strtempBaseQry += " UNION ALL" + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type ='D' then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS CrAmt , 0  AS  DrAmt ,0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," + Environment.NewLine & _
                " " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' THEN (sELECT Location_GL_Code  FROM " & TSPL_PAYMENT_HEADER & " PR WHERE PR.Payment_No =" & TSPL_PAYMENT_HEADER & ".Applied_Payment )  ELSE substring(" & TSPL_PAYMENT_HEADER & ".Debit_Account , len(" & TSPL_PAYMENT_HEADER & ".Debit_Account )-2,3) END as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " " + Environment.NewLine & _
                " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code Where " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  " + Environment.NewLine & _
                " --------------- INVOICE AGAINST APPLY DOCUMENT" + Environment.NewLine & _
                " UNION ALL" + Environment.NewLine & _
                " SELECT NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code,Max(OP_Type) as OP_Type FROM (select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo,CASE WHEN " & TSPL_PAYMENT_HEADER & ".Payment_Type ='AD' THEN 'IM' else null end as DocType , COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, 0  AS CrAmt,case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type ='D' then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end AS  DrAmt," + Environment.NewLine & _
                " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, " + Environment.NewLine & _
                " " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & "  " + Environment.NewLine & _
                " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code Where " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' " + Environment.NewLine & _
                " )INV  GROUP BY  DocNo,account  " + Environment.NewLine & _
                " ------- APPLY DOCUMENT ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION " + Environment.NewLine & _
                " UNION ALL " + Environment.NewLine & _
                " SELECT NULL AS Due_Date,'' AS PO_SRN,  MAX(VCode) AS VCode,MAX(VName) as VName,DocNo AS DocNo,MAX(DocType) AS DocType,MAX(DocDate) AS DocDate,MAX(DocNarr) AS DocNarr,MAX(ChequeDetails) AS ChequeDetails  ,MAX(Currency_Code) AS Currency_Code,MAX(ConvRate) AS ConvRate,SUM(CrAmt) AS CrAmt,SUM(DrAmt) AS DrAmt,SUM(Purchase) AS Purchase,SUM(Payments) AS Payments  ,MAX(DrNote) AS DrNote,MAX(CrNote) AS CrNote, MAX(Document_Type) AS Document_Type,SUM(Balance_Amount) AS Balance_Amount,account , MAX(Posting_Date) AS Posting_Date,MAX(GLDocType) AS GLDocType,MAX(Comp_Code) AS Comp_Code,Max(OP_Type) as OP_Type FROM ( " + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_BANK_REVERSE & ".Reverse_Code as DocNo,'Reverse Payment' as DocType , Convert(date," & TSPL_BANK_REVERSE & ".Reversal_Date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr,((" & TSPL_PAYMENT_HEADER & ".Cheque_No) + (case when " & TSPL_PAYMENT_HEADER & ".Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type ='D' then TSPL_PAYMENT_DETAIL.Applied_Amount * -1 else TSPL_PAYMENT_DETAIL.Applied_Amount  end  AS CrAmt,0 AS  DrAmt, " + Environment.NewLine & _
                " 0 as Purchase, 0 as Payments, 0 as DrNote, 0 as CrNote," & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount,  " + Environment.NewLine & _
                " " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code as account," & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & "  " + Environment.NewLine & _
                " from " & TSPL_BANK_REVERSE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No =" & TSPL_BANK_REVERSE & ".Document_No  LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code Where" + Environment.NewLine & _
                " " & TSPL_BANK_REVERSE & ".Reverse_Document='Payments' AND " & TSPL_BANK_REVERSE & ".Post ='P' AND " & TSPL_PAYMENT_HEADER & ".Payment_Type='AD' AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)<>1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'" + Environment.NewLine & _
                " )INV  GROUP BY  DocNo,account " + Environment.NewLine & _
                " -------------- FOR BANK REVERSE ENTRY WHEN LOCATION IS CHANGED OF APPLIED DOCUMENT WITH DOCUMENT LOCATION" + Environment.NewLine

                End If

            Else
                strtempBaseQry = " Select NULL as Due_Date, MAX(PO_SRN) as PO_SRN, MAX(VCode) as VCode, MAX(VName) as VName, DocNo, MAX(DocType) as DocType, MAX(DocDate) as DocDate, MAX(DocNarr) as DocNarr, MAX(ChequeDetails) as ChequeDetails, MAX(Currency_Code) as Currency_Code, MAX(ConvRate) as ConvRate, SUM(CrAmt) as CrAmt, SUM(DrAmt) DrAmt, 0 as Purchase, SUM(DrAmt)-SUM(CrAmt) as Payments, 0 as DrNote, 0 as CrNote, MAX(Document_Type) as Document_Type, SUM(Balance_Amount) as Balance_Amount, account, MAX(Posting_Date) as Posting_Date, MAX(GLDocType) as GLDocType, MAX(Comp_Code) as Comp_Code,Max(OP_Type) as OP_Type from (" + Environment.NewLine & _
                " select  NULL AS Due_Date,'' AS PO_SRN, " & TSPL_PAYMENT_HEADER & ".vendor_code as VCode, " & TSPL_PAYMENT_HEADER & ".vendor_name as VName, " & TSPL_PAYMENT_HEADER & ".payment_no as DocNo, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='AV' then 'Advance' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='OA' then 'On Account' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then 'Payment' when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 'Receipt' else 'Mislleneous' end as DocType, COnvert(date,payment_date, 103) as DocDate, ((entry_desc)+(case when entry_desc='' then '' else ' - 'end)+(reference)+(case when reference='' then '' else ' - 'end)+  ( narration)) as DocNarr, ((Cheque_No) + (case when Cheque_No='' then '' else ' - ' end) +CONVERT(Varchar, Cheque_Date, 103)) as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".payment_type in ('RC') then Payment_Amount else 0 end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".payment_type IN('OA','AV') then Payment_Amount When " & TSPL_PAYMENT_HEADER & ".payment_type IN('PY') Then (Case When " & TSPL_VENDOR_INVOICE_HEAD & ".Document_Type='D' Then -TSPL_PAYMENT_DETAIL.Applied_Amount Else TSPL_PAYMENT_DETAIL.Applied_Amount End) else 0 end as DrAmt, " & TSPL_PAYMENT_HEADER & ".Payment_Type as Document_Type,Payment_Amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right(TSPL_BANK_MASTER.BANKACC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY','AV','OA','RC') then 'AP-PY' else case when " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('MI') then 'AP-MI' else '' end  end as GLDocType, " & TSPL_PAYMENT_HEADER & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_PAYMENT_HEADER & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_PAYMENT_HEADER & " LEFT OUTER JOIN TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No left outer join TSPL_BANK_MASTER on " & TSPL_PAYMENT_HEADER & ".Bank_Code=TSPL_BANK_MASTER.BANK_CODE LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_PAYMENT_HEADER & ".Vendor_Code Where " & TSPL_PAYMENT_HEADER & ".Payment_Type NOT IN ('MI','AD') AND (Posted='P' or Posted='1') AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  " + Environment.NewLine & _
                " ) XX Group By XX.account, XX.DocNo" + Environment.NewLine & _
            " UNION ALL" + Environment.NewLine & _
                " select " & TSPL_PAYMENT_HEADER & ".Payment_Date As Due_Date, '' AS PO_SRN, " & TSPL_REMITTANCE & ".Vendor_Code, " & TSPL_REMITTANCE & ".Vendor_Name, " & TSPL_REMITTANCE & ".Document_No, 'TDS' as [DocType], convert(date,Document_Date,103)as Document_Date, " & TSPL_PAYMENT_HEADER & ".Entry_Desc as DocNarr, '', " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE as Currency_Code, " & TSPL_PAYMENT_HEADER & ".ConvRate as ConvRate, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then 0 else Actual_Total_TDS END as CrAmt, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then Actual_Total_TDS else 0 END as DrAmt, 0 as Purchase, case when " & TSPL_REMITTANCE & ".Document_Type IN('OA','AV') then Actual_Total_TDS else Actual_Total_TDS END as Payments, 0 as DrNotes, 0 as CrNotes, 'TDS',0, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code Else right( Branch_GL_AC,3) End as account, " & TSPL_PAYMENT_HEADER & ".Payment_Post_Date as Posting_Date, case when " & TSPL_REMITTANCE & ".Document_Type in ('AV','OA','RC') then 'AP-PY' else '' end as GLDocType, " & TSPL_REMITTANCE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_REMITTANCE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_REMITTANCE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No=" & TSPL_REMITTANCE & ".Document_No left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_REMITTANCE & ".Vendor_Code where Remit_TDS is not null   and   Branch_GL_AC  is not null  and Actual_Total_TDS<>0 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1  AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' " + Environment.NewLine & _
            " UNION ALL" + Environment.NewLine & _
                " select null AS Due_Date, '' AS PO_SRN, " & TSPL_BANK_REVERSE & ".vendor_code as VCode, " & TSPL_BANK_REVERSE & ".vendor_name as VName,Reverse_Code as DocNo, 'Reverse Payment' as [DocType],convert(date,Reversal_Date, 103) as DocDate, " & TSPL_BANK_REVERSE & ".Document_No+case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then ', '+TSPL_PAYMENT_DETAIL.Document_No Else '' End as DocNarr,(" & TSPL_BANK_REVERSE & ".cheque_No+'-'+ CONVERT(VARCHAR,Pay_Rec_Date, 103))as ChequeDetails, " & TSPL_PAYMENT_HEADER & ".CURRENCY_CODE, " & TSPL_PAYMENT_HEADER & ".ConvRate, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then 0 when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount ELSE " & TSPL_BANK_REVERSE & ".Amount end as CrAmt, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount else 0 end  as DrAmt, 0 as Purchase, case when " & TSPL_PAYMENT_HEADER & ".Payment_Type='RC' then " & TSPL_BANK_REVERSE & ".Amount when " & TSPL_PAYMENT_HEADER & ".Payment_Type='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount*-1 ELSE " & TSPL_BANK_REVERSE & ".Amount*-1 end as Payments, 0 as DrNote, 0 as CrNote, 'RV'as Document_Type, amount as Balance_Amount, Case When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('AV','OA','RC') Then " & TSPL_PAYMENT_HEADER & ".Location_GL_Code When " & TSPL_PAYMENT_HEADER & ".Payment_Type in ('PY') Then " & TSPL_VENDOR_INVOICE_HEAD & ".Loc_Code Else right(BANKACC,3) End as account, " & TSPL_BANK_REVERSE & ".Reversal_Date as Posting_Date,'RV-TA' as GLDocType, " & TSPL_BANK_REVERSE & ".Comp_Code," & IIf(objFilter.IS_GIT = True, TSPL_BANK_REVERSE & "." & OP_TypeCol, OP_TypeCol) & " from " & TSPL_BANK_REVERSE & " LEFT OUTER JOIN " & TSPL_PAYMENT_HEADER & " ON " & TSPL_PAYMENT_HEADER & ".Payment_No = " & TSPL_BANK_REVERSE & ".Document_No LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=" & TSPL_PAYMENT_HEADER & ".Payment_No LEFT OUTER JOIN " & TSPL_VENDOR_INVOICE_HEAD & " ON " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No=TSPL_PAYMENT_DETAIL.Document_No left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=" & TSPL_BANK_REVERSE & ".Bank_Code left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =" & TSPL_BANK_REVERSE & ".Vendor_Code where Source_Type ='AP' And " & TSPL_BANK_REVERSE & ".Post='P' AND " & TSPL_BANK_REVERSE & ".Vendor_Code<> '' AND " & TSPL_PAYMENT_HEADER & ".IsChkReverse='Y' AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Advance_Against_Salary,0)=1 AND ISNULL(" & TSPL_PAYMENT_HEADER & ".Is_Security,0)<>1   AND TSPL_VENDOR_MASTER.CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "' "

            End If
            objFilter.strtempBaseQryforopening = strtempBaseQry
            If clsCommon.CompairString(clsCommon.myCstr(objFilter.CurrencyType), "") = CompairStringResult.Equal Then
                'LoadCurrencyType()
            End If
            If objFilter.DocumentWise = True Then
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) as  CrAmt ," + Environment.NewLine & _
                    " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
                    "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
                    "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine & _
                    " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) else  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
                    " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then   (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end else " + Environment.NewLine & _
                    "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
                    " case when InnQuery.DocType='Pur.Invoice' then InnQuery.CrAmt-InnQuery.DrAmt else 0 end as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine & _
                    "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_M1ASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  "
                Else
                    strtempBaseQry = "  Select isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) as EXCHANGE_GAIN_AMT,isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) as EXCHANGE_LOSS_AMT, InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  * (case when (DocType)<>'EXC' then   InnQuery.ConvRate else 1 end) - case when (DocType)='Reverse Payment' then  (Select isnull((PH.EXCHANGE_LOSS_AMT  ),0) from " & TSPL_PAYMENT_HEADER & " PH where PH.Payment_No =(  Select Document_No   from " & TSPL_BANK_REVERSE & " where Reverse_Code =InnQuery.DocNo )) else 0 end as  CrAmt ," + Environment.NewLine & _
                   " case when len( (isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'')))<=0 then " + Environment.NewLine & _
                   "case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=(InnQuery.account) then " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*   " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) else " + Environment.NewLine & _
                   "case when isnull((DrAmt),0)>0 then (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) else 0 end end else " + Environment.NewLine & _
                   " case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 AND (DocType)<>'EXC' then  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") -isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT ),0) ELSE  (DrAmt*  " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0) end end else " + Environment.NewLine & _
                   " case when (RIGHT(TSPL_BANK_MASTER.BANKACC,3))=isnull(" & TSPL_PAYMENT_HEADER & ".Location_GL_Code,'') then " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then CASE WHEN ((DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")-isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0))<0 THEN 0 ELSE (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")-isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0) END  else (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_GAIN_AMT ),0)  end else " + Environment.NewLine & _
                   "case when isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)>0 then  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ") +isnull((" & TSPL_PAYMENT_HEADER & ".EXCHANGE_LOSS_AMT),0)  else  (DrAmt* " & IIf(clsCommon.myCstr(objFilter.CurrencyType) = "1", 1, " InnQuery." & objFilter.CurrencyType & "") & ")  end end end as DrAmt, " + Environment.NewLine & _
                   " InnQuery.CrAmt-InnQuery.DrAmt  as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & strtempBaseQry & " ) InnQuery " + Environment.NewLine & _
                   "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  "

                End If
            Else
                strtempBaseQry = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  "
                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    strtempBaseQry += "  CONVERT(DECIMAL(18,2),CASE WHEN InnQuery.DocType NOT IN ('Pur.Invoice','Receipt')  THEN case when InnQuery.DocType not in ('TDS','AP Invoice') THEN InnQuery.CrAmt  when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo)='C' then InnQuery.CrAmt ELSE CASE WHEN InnQuery.DocType in ('TDS') and (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  )='I' or (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo   )='C' then (Select ISNULL (Actual_Total_TDS ,0)  from " & TSPL_REMITTANCE & " where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  ) * -1 else 0 end END ELSE 0 END) as CrAmt ," & _
                      "  CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Adjustment') then case when InnQuery.DocType not in ('AP Invoice','TDS') then InnQuery.DrAmt WHEN InnQuery.DocType in ('TDS') and   (Select ISNULL(" & TSPL_REMITTANCE & ".Document_Type,'')  from " & TSPL_REMITTANCE & " LEFT OUTER join " & TSPL_VENDOR_INVOICE_HEAD & " on " & TSPL_VENDOR_INVOICE_HEAD & ".Document_No = " & TSPL_REMITTANCE & ".Document_No  and  isnull(" & TSPL_VENDOR_INVOICE_HEAD & ".Against_PurchaseReturn_No,'')='' where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  )='D' then (Select ISNULL (Actual_Total_TDS ,0)  from " & TSPL_REMITTANCE & " where " & TSPL_REMITTANCE & ".Document_No=InnQuery.DocNo  ) * -1 else case when  InnQuery.DocType in ('AP Invoice') AND  (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')='')='D' then InnQuery.DrAmt else 0 end end else 0 end) DrAmt, " & _
                      "  CONVERT(DECIMAL(18,2),case when InnQuery.DocType not in ('Payment','On Account','Advance','Receipt','Debit Note','Reverse Payment','Credit Note','EXC','IM') then case when InnQuery.DocType not in ('AP Invoice','TDS') then InnQuery.Purchase else case when  InnQuery.DocType in ('AP Invoice') and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo)='I'	 then InnQuery.Purchase when  InnQuery.DocType in ('AP Invoice')  and (Select Document_Type  from " & TSPL_VENDOR_INVOICE_HEAD & " where document_no=InnQuery.DocNo AND ISNULL(Against_PurchaseReturn_No,'')<>'')='D' then InnQuery.Purchase *-1  else 0 end end else 0 end) Purchase, " & _
                      "  CONVERT(DECIMAL(18,2),case when InnQuery.DocType  in ('Payment','On Account','Advance','Receipt','TDS','EXC','IM') then InnQuery.Payments  else 0 end) as Payments , "
                Else
                    strtempBaseQry += "  CONVERT(DECIMAL(18,2),InnQuery.CrAmt) as CrAmt, CONVERT(DECIMAL(18,2),InnQuery.DrAmt) as DrAmt, CONVERT(DECIMAL(18,2),InnQuery.Purchase) as Purchase, CONVERT(DECIMAL(18,2),InnQuery.Payments) as Payments , "
                End If
                strtempBaseQry += " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE from " & _
                "(Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt  as  CrAmt ," + Environment.NewLine & _
                " InnQuery.DrAmt as DrAmt,InnQuery.CrAmt-InnQuery.DrAmt as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & objFilter.strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine & _
                "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  " & _
                  " left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =InnQuery.VCode where TSPL_VENDOR_MASTER .CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  ) InnQuery"


                If clsCommon.CompairString(objFilter.FormType, clsUserMgtCode.MISCreditorReport) = CompairStringResult.Equal Then
                    objFilter.strtempBaseQryforopeningForMIS = "  Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate,  " & _
                    " InnQuery.CrAmt,InnQuery.DrAmt,0 as Purchase,0 as Payments, " & _
                    " InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE from " & _
                    " (Select InnQuery.Due_Date, InnQuery.PO_SRN, InnQuery.VCode, InnQuery.VName, InnQuery.DocNo, InnQuery.DocType, InnQuery.DocDate, InnQuery.DocNarr, InnQuery.ChequeDetails, InnQuery.Currency_Code, InnQuery.ConvRate, InnQuery.CrAmt as  CrAmt ," + Environment.NewLine & _
                    " InnQuery.DrAmt AS DrAmt ,InnQuery.CrAmt-InnQuery.DrAmt as Purchase, InnQuery.Payments, InnQuery.drNote, InnQuery.CrNote, InnQuery.Document_Type, InnQuery.Balance_Amount, InnQuery.account,InnQuery.Posting_Date, InnQuery.GLDocType, InnQuery.Comp_Code,InnQuery.OP_TYPE    from  (" & objFilter.strtempBaseQryforopening & " ) InnQuery " + Environment.NewLine & _
                    "  left outer join " & TSPL_PAYMENT_HEADER & " on " & TSPL_PAYMENT_HEADER & ".Payment_No  =InnQuery.DocNo LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=" & TSPL_PAYMENT_HEADER & ".Bank_Code  " & _
                    " left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER .Vendor_Code =InnQuery.VCode where TSPL_VENDOR_MASTER .CURRENCY_CODE <>'" & objCommonVar.BaseCurrencyCode & "'  ) InnQuery"
                End If

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return strtempBaseQry

    End Function
    Public Shared Function UpdateVendorSummaryData(ByVal QryCond As String, ByVal trans As SqlTransaction) As Boolean
        If clsCommon.myLen(QryCond) > 0 Then
            QryCond = " and " & QryCond
        End If
        '' update summary transaction data
        Dim qry As String = GetBaseQueryWIN(QryCond)
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        '' UPDATE CLOSING DATA
        qry = ReturnClosingUpdateQry(QryCond)
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        qry = "delete from TSPL_VENDOR_INVOICE_HEAD_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_PAYMENT_HEADER_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_Payment_Adjustment_Header_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_VCGL_Head_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_BANK_REVERSE_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_REMITTANCE_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_PI_HEAD_WIN where 2=2  " & QryCond & ""
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Return True
    End Function
    Public Shared Function GetBaseQueryWIN(ByVal QryCond As String) As String
        'Dim qry As String = GetBaseQuery(QryCond, True)
        Dim qry As String = "select * from (select * from View_VENDOR_DATA_GIT union all select * from View_VENDOR_DATA_Currency_GIT ) View_VENDOR_DATA_GIT where 2=2 "
        If clsCommon.myLen(QryCond) > 0 Then
            qry = qry & " And " & QryCond
        End If
        Dim qryWin As String = " select VCode,CURRENCY_CODE,cast(DocDate as date) as Punching_Date, Coalesce(sum(DrAmt*(case when OP_Type='I' then 1 else -1 end)),0) as DrAmt," & _
                               " Coalesce(sum(CrAmt*(case when OP_Type='I' then -1 else 1 end)),0) as CrAmt, " & _
                               " (Coalesce(sum(DrAmt*(case when OP_Type='I' then 1 else -1 end)),0)-Coalesce(sum(CrAmt*(case when OP_Type='I' then -1 else 1 end)),0)) as TRANS_AMOUNT " & _
                               " from (" & qry & ") as BaseTable group by VCode,CURRENCY_CODE,cast(DocDate as date)"
        qryWin = " MERGE INTO TSPL_VENDOR_SUMMARY_DL  A" & _
                 " USING(" & qryWin & ") TA " & _
                 " ON (A.VENDOR_CODE=TA.VCode AND A.CURRENCY_CODE=TA.CURRENCY_CODE AND A.TRANS_DATE=TA.PUNCHING_DATE) " & _
                 " WHEN MATCHED THEN " & _
                 " update  SET A.TRANS_AMOUNT=A.TRANS_AMOUNT+TA.TRANS_AMOUNT,A.TOTAL_DR=A.TOTAL_DR+TA.DrAmt,A.TOTAL_CR=A.TOTAL_CR+TA.CrAmt,  A.CL_AMOUNT=A.CL_AMOUNT+TA.TRANS_AMOUNT " & _
                 " WHEN NOT MATCHED THEN  " & _
                 " insert  (VENDOR_CODE,TRANS_DATE,CURRENCY_CODE,TOTAL_DR,TOTAL_CR,TRANS_AMOUNT,CL_AMOUNT)  " & _
                 " VALUES(TA.VCode,TA.Punching_Date,TA.CURRENCY_CODE,TA.DrAmt,TA.CrAmt,TA.TRANS_AMOUNT,TA.TRANS_AMOUNT);"
        Return qryWin
    End Function
    Public Shared Function ReturnClosingUpdateQry(ByVal QryCond As String) As String
        ''''AND CL.TRANS_DATE=TEMP_DATA.TRANS_DATE
        Dim qry As String = " MERGE INTO TSPL_VENDOR_SUMMARY_DL  A" & _
                            " USING(SELECT CL.VENDOR_CODE,CL.CURRENCY_CODE,CL.TRANS_DATE,CL.CL_TRANS_AMOUNT" & _
                            " FROM ( " & _
                            " select TSPL_VENDOR_SUMMARY_DL.VENDOR_CODE,CURRENCY_CODE,TRANS_DATE, " & _
                            " sum(TRANS_AMOUNT) over (partition by VENDOR_CODE,CURRENCY_CODE order by VENDOR_CODE,CURRENCY_CODE,TRANS_DATE) as CL_TRANS_AMOUNT " & _
                            " from TSPL_VENDOR_SUMMARY_DL) CL " & _
                            " INNER JOIN (SELECT DISTINCT VENDOR_CODE,CURRENCY_CODE FROM " & _
                            " (SELECT VENDOR_CODE,CURRENCY_CODE,cast(Posting_Date as date) AS TRANS_DATE FROM TSPL_VENDOR_INVOICE_HEAD_WIN " & _
                            " UNION ALL " & _
                            " SELECT Vendor_Code,CURRENCY_CODE,cast(Payment_Date as date) AS TRANS_DATE FROM TSPL_PAYMENT_HEADER_WIN " & _
                            " UNION ALL " & _
                            " SELECT Vendor_No,'" & objCommonVar.BaseCurrencyCode & "' AS CURRENCY_CODE,cast(Adjustment_Date as date) AS TRANS_DATE FROM TSPL_Payment_Adjustment_Header_WIN" & _
                            " UNION ALL " & _
                            " SELECT VC_Code,'" & objCommonVar.BaseCurrencyCode & "' AS CURRENCY_CODE,cast(Document_Date as date) AS TRANS_DATE FROM TSPL_VCGL_Head_WIN WHERE Document_Type='V' " & _
                            " UNION ALL " & _
                            " select Vendor_Code,'" & objCommonVar.BaseCurrencyCode & "' AS CURRENCY_CODE,cast(Reversal_Date as date) AS TRANS_DATE " & _
                            " from TSPL_BANK_REVERSE_WIN where len(coalesce(Vendor_Code,''))>0 " & _
                            " union all " & _
                            " select VENDOR_CODE,'" & objCommonVar.BaseCurrencyCode & "' AS CURRENCY_CODE,cast(PI_Date as date) AS TRANS_DATE " & _
                            " from TSPL_PI_HEAD_WIN where len(coalesce(VENDOR_CODE,''))>0" & _
                            " ) TEMP  WHERE 2=2 " & QryCond & " ) TEMP_DATA ON CL.VENDOR_CODE=TEMP_DATA.VENDOR_CODE " & _
                            " AND CL.CURRENCY_CODE=TEMP_DATA.CURRENCY_CODE  " & _
                            " ) TA  " & _
                            " ON (A.VENDOR_CODE=TA.VENDOR_CODE AND A.CURRENCY_CODE=TA.CURRENCY_CODE AND A.TRANS_DATE=TA.TRANS_DATE) " & _
                            " WHEN MATCHED THEN  " & _
                            " update " & _
                            " SET A.CL_AMOUNT=TA.CL_TRANS_AMOUNT;"

        Return qry
    End Function
    Public Shared Function GetVendorOpening_GIT(ByVal VendorList As ArrayList, ByVal Trans_Date As Date) As String
        Dim cond As String = ""
        Dim qry As String = " select max(Due_Date) as Due_Date,max(PO_SRN) as PO_SRN,VCode,VName,max(DocNo) as DocNo,max(DocType) as DocType,max(DocDate) as DocDate,max(DocNarr) as DocNarr," & _
                            " max(ChequeDetails) as ChequeDetails,max(CURRENCY_CODE) as CURRENCY_CODE,max(ConvRate) as ConvRate,sum(DrAmt) as DrAmt,sum(CrAmt) as CrAmt," & _
                            " sum(Purchase) as Purchase,sum(Payments) as Payments,sum(drNote) as drNote,sum(crNote) as crNote,max(Document_Type) as Document_Type,sum(Balance_Amount) as Balance_Amount " & _
                            " ,max(account) as account,max(Posting_Date) as Posting_Date,max(GLDocType) as GLDocType,max(Comp_Code) as Comp_Code from ( " & _
                            " select '' as Due_Date,'' as PO_SRN,VENDOR_Code as VCode,'' as VName,'' as DocNo,'' as DocType,Trans_Date as DocDate,'' as DocNarr,'' as ChequeDetails,CURRENCY_CODE, " & _
                            " 1 as ConvRate,TOTAL_DR as DrAmt,TOTAL_CR as CrAmt,0 as Purchase,0 as Payments,0 as drNote,0 as crNote,'' as Document_Type,0 as Balance_Amount,'' as account, " & _
                            " null as Posting_Date,'' as GLDocType,'' as Comp_Code from TSPL_VENDOR_SUMMARY_DL where Trans_Date<'" & clsCommon.GetPrintDate(Trans_Date, "dd-MMM-yyyy") & "'"
        If Not VendorList Is Nothing AndAlso VendorList.Count > 0 Then
            qry = qry & " and " & " VENDOR_Code in " & "(" & clsCommon.GetMulcallStringWithComma(VendorList) & ")"
        End If
        qry = qry & " union all " & _
              " select Due_Date,PO_SRN,VCode,VName,DocNo,DocType,DocDate,DocNarr,ChequeDetails,CURRENCY_CODE,ConvRate,DrAmt,CrAmt,Purchase,Payments,drNote,crNote, " & _
              " Document_Type,Balance_Amount,account,Posting_Date,GLDocType,Comp_Code from View_VENDOR_DATA_GIT where DocDate<'" & clsCommon.GetPrintDate(Trans_Date, "dd-MMM-yyyy") & "' "

        If Not VendorList Is Nothing AndAlso VendorList.Count > 0 Then
            qry = qry & " and " & " VCode in " & "(" & clsCommon.GetMulcallStringWithComma(VendorList) & ")"
        End If

        qry = qry & " ) as OP_GIT group by VCode,VName"

        Return qry
    End Function

    Public Shared Function GetCorrectionVSP(ByVal strVSPCode As String, ByVal trans As SqlTransaction)
        Dim obj As New clsVendorMaster
        Dim qry As String = "select CorrectionFat,CorrectionSNF from TSPL_VENDOR_MASTER where Vendor_Code='" + strVSPCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            obj.CorrectionFat = clsCommon.myCdbl(dt.Rows(0)("CorrectionFat"))
            obj.CorrectionSNF = clsCommon.myCdbl(dt.Rows(0)("CorrectionSNF"))
        End If
        dt = Nothing
        Return obj
    End Function

    Public Shared Function GetVendorRegisterORNonRegister(ByVal strCode As String, ByVal Trans As SqlTransaction) As String
        Dim qry As String = " select case when  GSTRegistered = 1 then 'Registered'  when  GSTRegistered = 0 then 'Non Registered'  end  from tspl_vendor_master where Vendor_Code = '" + strCode + "' "
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Trans))
    End Function
    Public Shared Function GetVendorGSTINNo(ByVal strCode As String, ByVal Trans As SqlTransaction) As String
        Dim qry As String = " select GSTFinalNo  from tspl_vendor_master where Vendor_Code =  '" + strCode + "' "
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Trans))
    End Function
    Public Shared Function GetVendorEmailID(ByVal strCode As String, ByVal Trans As SqlTransaction) As String
        Dim qry As String = " select Email  from tspl_vendor_master where Vendor_Code =  '" + strCode + "' "
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Trans))
    End Function

    Public Shared Function IsVLCDripSaver(ByVal strVLCCode As String, ByVal tran As SqlTransaction) As Boolean ''ERO/06/03/19-000508 by balwinder on 08/04/2019

        Dim qry As String = "select TSPL_VENDOR_MASTER.is_Drip_Saver from TSPL_VLC_MASTER_HEAD" + Environment.NewLine +
        "left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code" + Environment.NewLine +
        "where TSPL_VLC_MASTER_HEAD.VLC_Code='" + strVLCCode + "' and TSPL_VENDOR_MASTER.is_Drip_Saver='Y'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, tran)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        End If
        Return False

    End Function

    Public Shared Function GetVendorOutstandingForTCSTaxApplicableOnFY(ByVal strVendor As String, ByVal strDocumentDate As Date, Optional ByVal trans As SqlTransaction = Nothing) As String
        Dim balanceAmt As Double = 0
        Dim FinancialYear As String = String.Empty
        Dim strStartDate As Date
        Dim ConsiderPreviousandCurrentFYForTCSTaxCustOutstanding As Boolean = IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ConsiderPreviousCurrentFYForTCSTaxVendOutstanding, clsFixedParameterCode.ConsiderPreviousCurrentFYForTCSTaxVendOutstanding, trans)) = "1", True, False)
        Try
            If ConsiderPreviousandCurrentFYForTCSTaxCustOutstanding = True Then
                FinancialYear = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT CASE WHEN DatePart(Month, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') >= 4 THEN DatePart(Year, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') - 1 ELSE DatePart(Year, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') -1 END AS Fiscal_Year", trans))
                strStartDate = "01/Apr/" + clsCommon.myCstr(clsCommon.myCdbl(FinancialYear))
                Dim strEndDate As Date = "31/Mar/" + clsCommon.myCstr(clsCommon.myCdbl(FinancialYear) + 1)
                Dim strwhrcls As String = " and (len(isnull(Against_PurchaseReturn_No,''))>0 or len(isnull(Against_POInvoice_No,''))>0 or len(isnull(Against_BulkMillkPurchaseInvoice_No,''))>0 or len(isnull(Against_MillkPurchaseInvoice_No,''))>0)"
                Dim PANNumber As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PAN FROM TSPL_VENDOR_MASTER WHERE Vendor_Code = '" & strVendor & "'", trans))

                Dim strqry As String = " select isnull(sum(final.totalAmount),0) from ( select case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then TSPL_VENDOR_INVOICE_HEAD.Document_Total *-1 else TSPL_VENDOR_INVOICE_HEAD.Document_Total end as totalAmount,TSPL_VENDOR_INVOICE_HEAD.* " + Environment.NewLine +
                " from TSPL_VENDOR_INVOICE_HEAD " + Environment.NewLine +
                " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  " + Environment.NewLine +
                " where 1=1 And (TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" & strVendor & "'  "
                If clsCommon.myLen(PANNumber) > 0 Then
                    strqry += "  Or isnull(TSPL_VENDOR_MASTER.PAN,'') ='" & PANNumber & "' "
                End If
                strqry += " ) and TSPL_VENDOR_INVOICE_HEAD.Posting_Date is not null " + Environment.NewLine +
                " And convert(date,TSPL_VENDOR_INVOICE_HEAD.Posting_Date ,103)  >= '" + clsCommon.GetPrintDate(strStartDate, "dd/MMM/yyyy") + "'  and  convert(date,TSPL_VENDOR_INVOICE_HEAD.Posting_Date ,103)  <='" + clsCommon.GetPrintDate(strEndDate, "dd/MMM/yyyy") + "' " & strwhrcls & " )final "

                balanceAmt = clsDBFuncationality.getSingleValue(strqry, trans)
                If ConsiderPreviousandCurrentFYForTCSTaxCustOutstanding = True Then
                    Dim AmountToCheckCustomerOutstandingForTCSTax As Double = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AmountToCheckVendorOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckVendorOutstandingForTCSTax, Nothing))
                    If balanceAmt < AmountToCheckCustomerOutstandingForTCSTax Then
                        balanceAmt = GetVendorOutstandingForTCSTaxApplicableOnPreviousFY(strVendor, strDocumentDate, trans)
                    End If
                End If
            Else
                balanceAmt = GetVendorOutstandingForTCSTaxApplicableOnPreviousFY(strVendor, strDocumentDate, trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return balanceAmt
    End Function
    Public Shared Function GetVendorOutstandingForTCSTaxApplicableOnPreviousFY(ByVal strVendor As String, ByVal strDocumentDate As Date, Optional ByVal trans As SqlTransaction = Nothing) As String
        Dim balanceAmt As Double = 0
        Try
            Dim FinancialYear As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT CASE WHEN DatePart(Month, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') >= 4 THEN DatePart(Year, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') + 1 ELSE DatePart(Year, '" + clsCommon.GetPrintDate(strDocumentDate, "dd/MMM/yyyy") + "') END AS Fiscal_Year", trans))
            Dim strStartDate As Date = "01/Apr/" + clsCommon.myCstr(clsCommon.myCdbl(FinancialYear - 1))
            Dim strEndDate As Date = "31/Mar/" + FinancialYear
            Dim strwhrcls As String = " and (len(isnull(Against_PurchaseReturn_No,''))>0 or len(isnull(Against_POInvoice_No,''))>0 or len(isnull(Against_BulkMillkPurchaseInvoice_No,''))>0)"


            Dim PANNumber As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT PAN FROM TSPL_VENDOR_MASTER WHERE Vendor_Code = '" & strVendor & "'", trans))

            Dim strqry As String = " select isnull(sum(final.totalAmount),0) from ( select case when TSPL_VENDOR_INVOICE_HEAD.Document_Type ='C' then TSPL_VENDOR_INVOICE_HEAD.Document_Total *-1 else TSPL_VENDOR_INVOICE_HEAD.Document_Total end as totalAmount  " + Environment.NewLine +
                " from TSPL_VENDOR_INVOICE_HEAD " + Environment.NewLine +
                " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  " + Environment.NewLine +
 " where 1=1 And (TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" & clsCommon.myCstr(strVendor) & "' "
            If clsCommon.myLen(PANNumber) > 0 Then
                strqry += "  Or isnull(TSPL_VENDOR_MASTER.PAN,'') ='" & PANNumber & "' "
            End If
            strqry += " ) and TSPL_VENDOR_INVOICE_HEAD.Posting_Date is not null " + Environment.NewLine +
 " And convert(date,TSPL_VENDOR_INVOICE_HEAD.Posting_Date ,103)  >= '" + clsCommon.GetPrintDate(strStartDate, "dd/MMM/yyyy") + "'  and  convert(date,TSPL_VENDOR_INVOICE_HEAD.Posting_Date ,103)  <='" + clsCommon.GetPrintDate(strEndDate, "dd/MMM/yyyy") + "' " & strwhrcls & " )final "

            balanceAmt = clsDBFuncationality.getSingleValue(strqry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return balanceAmt
    End Function

    Public Shared Function IsTCSApplied(ByVal strVendorCode As String, ByVal DocDate As Date, ByVal Trans As SqlTransaction) As Boolean
        Dim flag As Boolean = False

        'Dim qry As String = "Select top 1 1 from TSPL_PI_HEAD where vendor_Code =  '" + strVendorCode + "' and isnull(ActualTCSBaseAmount,0)>1 and PI_Date<'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(DocDate), "dd/MMM/yyyy hh:mm:ss tt") + "' ""
        Dim qry As String = "select top 1 1 from TSPL_PI_HEAD where vendor_Code='" + strVendorCode + "' and isnull(ActualTCSBaseAmount,0)>1 and PI_Date<'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(DocDate), "dd/MMM/yyyy hh:mm:ss tt") + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            flag = True
        End If

        Return flag
    End Function
    Public Shared Function GetVendorData(ByVal strCode As String, ByVal trans As SqlTransaction) As clsVendorMaster
        Dim obj As clsVendorMaster = Nothing
        Dim qry As String = "select * from TSPL_VENDOR_MASTER where Vendor_Code='" + strCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            Throw New Exception("Vendor Code -" + strCode + ". Not Exist")
        End If
        obj = New clsVendorMaster()
        obj.Vendor_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Code"))
        obj.Vendor_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
        obj.Add1 = clsCommon.myCstr(dt.Rows(0)("Add1"))
        obj.Add2 = clsCommon.myCstr(dt.Rows(0)("Add2"))
        obj.Add3 = clsCommon.myCstr(dt.Rows(0)("Add3"))
        If dt.Rows(0)("Closing_Date") Is DBNull.Value Then
            obj.Closing_Date = Nothing
        Else
            obj.Closing_Date = clsCommon.myCstr(dt.Rows(0)("Closing_Date"))
        End If
        obj.Vendor_Group_Code = clsCommon.myCstr(dt.Rows(0)("Vendor_Group_Code"))
        obj.Vendor_Group_Code_Desc = clsCommon.myCstr(dt.Rows(0)("Vendor_Group_Code_Desc"))
        obj.City_Code = clsCommon.myCstr(dt.Rows(0)("City_Code"))
        obj.City_Code_Desc = clsCommon.myCstr(dt.Rows(0)("City_Code_Desc"))
        obj.State = clsCommon.myCstr(dt.Rows(0)("State"))
        obj.Country = clsCommon.myCstr(dt.Rows(0)("Country"))
        obj.Phone1 = clsCommon.myCstr(dt.Rows(0)("Phone1"))
        obj.Phone2 = clsCommon.myCstr(dt.Rows(0)("Phone2"))
        obj.Fax = clsCommon.myCstr(dt.Rows(0)("Fax"))
        obj.Email = clsCommon.myCstr(dt.Rows(0)("Email"))
        obj.WebSite = clsCommon.myCstr(dt.Rows(0)("WebSite"))
        obj.Contact_Person_Name = clsCommon.myCstr(dt.Rows(0)("Contact_Person_Name"))
        obj.Contact_Person_Phone = clsCommon.myCstr(dt.Rows(0)("Contact_Person_Phone"))
        obj.Contact_Person_Fax = clsCommon.myCstr(dt.Rows(0)("Contact_Person_Fax"))
        obj.Contact_Person_Website = clsCommon.myCstr(dt.Rows(0)("Contact_Person_Website"))
        obj.Contact_Person_Email = clsCommon.myCstr(dt.Rows(0)("Contact_Person_Email"))
        obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
        obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
        obj.Vendor_Account = clsCommon.myCstr(dt.Rows(0)("Vendor_Account"))
        obj.Vendor_Account_Desc = clsCommon.myCstr(dt.Rows(0)("Vendor_Account_Desc"))
        obj.Payment_Code = clsCommon.myCstr(dt.Rows(0)("Payment_Code"))
        obj.Payment_Code_Desc = clsCommon.myCstr(dt.Rows(0)("Payment_Code_Desc"))
        obj.Bank_Code = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
        obj.Bank_Code_Desc = clsCommon.myCstr(dt.Rows(0)("Bank_Code_Desc"))
        obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
        obj.Tax_Group_Desc = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
        obj.Ven_Type_Code = clsCommon.myCstr(dt.Rows(0)("Ven_Type_Code"))
        obj.Ven_Type_Desc = clsCommon.myCstr(dt.Rows(0)("Ven_Type_Desc"))
        obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
        obj.TAX1_Rate = clsCommon.myCstr(dt.Rows(0)("TAX1_Rate"))
        obj.TAX2 = clsCommon.myCstr(dt.Rows(0)("TAX2"))
        obj.TAX2_Rate = clsCommon.myCstr(dt.Rows(0)("TAX2_Rate"))
        obj.TAX3 = clsCommon.myCstr(dt.Rows(0)("TAX3"))
        obj.TAX3_Rate = clsCommon.myCstr(dt.Rows(0)("TAX3_Rate"))
        obj.TAX4 = clsCommon.myCstr(dt.Rows(0)("TAX4"))
        obj.TAX4_Rate = clsCommon.myCstr(dt.Rows(0)("TAX4_Rate"))
        obj.TAX5 = clsCommon.myCstr(dt.Rows(0)("TAX5"))
        obj.TAX5_Rate = clsCommon.myCstr(dt.Rows(0)("TAX5_Rate"))
        obj.TAX6 = clsCommon.myCstr(dt.Rows(0)("TAX6"))
        obj.TAX6_Rate = clsCommon.myCstr(dt.Rows(0)("TAX6_Rate"))
        obj.TAX7 = clsCommon.myCstr(dt.Rows(0)("TAX7"))
        obj.TAX7_Rate = clsCommon.myCstr(dt.Rows(0)("TAX7_Rate"))
        obj.TAX8 = clsCommon.myCstr(dt.Rows(0)("TAX8"))
        obj.TAX8_Rate = clsCommon.myCstr(dt.Rows(0)("TAX8_Rate"))
        obj.TAX9 = clsCommon.myCstr(dt.Rows(0)("TAX9"))
        obj.TAX9_Rate = clsCommon.myCstr(dt.Rows(0)("TAX9_Rate"))
        obj.TAX10 = clsCommon.myCstr(dt.Rows(0)("TAX10"))
        obj.TAX10_Rate = clsCommon.myCstr(dt.Rows(0)("TAX10_Rate"))
        obj.Service_Tax_No = clsCommon.myCstr(dt.Rows(0)("Service_Tax_No"))
        obj.Tin_No = clsCommon.myCstr(dt.Rows(0)("Tin_No"))
        obj.Lst_No = clsCommon.myCstr(dt.Rows(0)("Lst_No"))
        obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
        obj.State = clsCommon.myCstr(dt.Rows(0)("Status"))
        obj.OnHold = clsCommon.myCstr(dt.Rows(0)("OnHold"))
        obj.Transporter = clsCommon.myCstr(dt.Rows(0)("Transporter"))
        obj.Remarks1 = clsCommon.myCstr(dt.Rows(0)("Remarks1"))
        obj.Remarks2 = clsCommon.myCstr(dt.Rows(0)("Remarks2"))
        obj.Additional1 = clsCommon.myCstr(dt.Rows(0)("Additional1"))
        obj.Additional2 = clsCommon.myCstr(dt.Rows(0)("Additional2"))
        obj.Additional3 = clsCommon.myCstr(dt.Rows(0)("Additional3"))
        obj.Credit_Limit = clsCommon.myCstr(dt.Rows(0)("Credit_Limit"))
        obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
        obj.CST = clsCommon.myCstr(dt.Rows(0)("CST"))
        obj.ECC = clsCommon.myCstr(dt.Rows(0)("ECC"))
        obj.Range = clsCommon.myCstr(dt.Rows(0)("Range"))
        obj.Collectorate = clsCommon.myCstr(dt.Rows(0)("Collectorate"))
        obj.PAN = clsCommon.myCstr(dt.Rows(0)("PAN"))
        obj.Is_Gross_Receipt = clsCommon.myCstr(dt.Rows(0)("Is_Gross_Receipt"))
        obj.Inter_Branch = clsCommon.myCstr(dt.Rows(0)("Inter_Branch"))
        obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
        obj.franchise_yn = clsCommon.myCstr(dt.Rows(0)("franchise_yn"))
        obj.Form_Type = clsCommon.myCstr(dt.Rows(0)("Form_Type"))
        obj.State_Code = clsCommon.myCstr(dt.Rows(0)("State_Code"))
        obj.Country_Code = clsCommon.myCstr(dt.Rows(0)("Country_Code"))
        obj.Service_charges = clsCommon.myCdbl(dt.Rows(0)("Service_charges"))
        obj.commision_pers = clsCommon.myCdbl(dt.Rows(0)("commision_pers"))
        obj.incentive = clsCommon.myCstr(dt.Rows(0)("incentive"))
        obj.incentive_days = clsCommon.myCdbl(dt.Rows(0)("incentive_days"))
        obj.vsp_payment = clsCommon.myCstr(dt.Rows(0)("vsp_payment"))
        obj.VSP_Payee_Name = clsCommon.myCstr(dt.Rows(0)("VSP_Payee_Name"))
        obj.Zila = clsCommon.myCstr(dt.Rows(0)("Zila"))
        obj.Tehsil = clsCommon.myCstr(dt.Rows(0)("Tehsil"))
        obj.Branch_Name = clsCommon.myCstr(dt.Rows(0)("Branch_Name"))
        obj.IFCI_Code = clsCommon.myCstr(dt.Rows(0)("IFCI_Code"))
        obj.Account_No = clsCommon.myCstr(dt.Rows(0)("Account_No"))
        obj.Industry_Type = clsCommon.myCstr(dt.Rows(0)("Industry_Type"))
        obj.Industry_Person = clsCommon.myCstr(dt.Rows(0)("Industry_Person"))
        obj.Agreement = clsCommon.myCstr(dt.Rows(0)("Agreement"))
        obj.Security_Cheque = clsCommon.myCstr(dt.Rows(0)("Security_Cheque"))
        obj.No_of_Installment = clsCommon.myCdbl(dt.Rows(0)("No_of_Installment"))
        obj.Amount_of_Installment = clsCommon.myCdbl(dt.Rows(0)("Amount_of_Installment"))
        obj.IsPermanent = clsCommon.myCstr(dt.Rows(0)("IsPermanent"))
        obj.IsTemporary = clsCommon.myCstr(dt.Rows(0)("IsTemporary"))
        obj.Joint_Name = clsCommon.myCstr(dt.Rows(0)("Joint_Name"))
        obj.Service_Charge_Type = clsCommon.myCstr(dt.Rows(0)("Service_Charge_Type"))
        obj.Is_Parent_Vendor = clsCommon.myCstr(dt.Rows(0)("Is_Parent_Vendor"))
        obj.Parent_Vendor_Code = clsCommon.myCstr(dt.Rows(0)("Parent_Vendor_Code"))
        obj.branch_code = clsCommon.myCstr(dt.Rows(0)("branch_code"))
        obj.Category_Struct_Code = clsCommon.myCstr(dt.Rows(0)("Category_Struct_Code"))
        obj.Bank_Name = clsCommon.myCstr(dt.Rows(0)("Bank_Name"))
        obj.IFSC_Code = clsCommon.myCstr(dt.Rows(0)("IFSC_Code"))
        obj.Account_Type = clsCommon.myCstr(dt.Rows(0)("Account_Type"))
        obj.Vendor_Type = clsCommon.myCstr(dt.Rows(0)("Vendor_Type"))
        obj.payment_commision_pers = clsCommon.myCdbl(dt.Rows(0)("payment_commision_pers"))
        obj.Pin_Code = clsCommon.myCstr(dt.Rows(0)("Pin_Code"))
        obj.Security_Amount = clsCommon.myCdbl(dt.Rows(0)("Security_Amount"))
        obj.AMC_Charge = clsCommon.myCdbl(dt.Rows(0)("AMC_Charge"))
        obj.AMCU = clsCommon.myCstr(dt.Rows(0)("AMCU"))
        If dt.Rows(0)("Billing_Date") Is DBNull.Value Then
            obj.Billing_Date = Nothing
        Else
            obj.Billing_Date = clsCommon.myCstr(dt.Rows(0)("Billing_Date"))
        End If
        obj.Is_Chilling_vendor = clsCommon.myCstr(dt.Rows(0)("Is_Chilling_vendor"))
        obj.TDS_Branch_Code = clsCommon.myCstr(dt.Rows(0)("TDS_Branch_Code"))
        obj.Deduction_Code = clsCommon.myCstr(dt.Rows(0)("Deduction_Code"))
        obj.TDS_State_Code = clsCommon.myCstr(dt.Rows(0)("TDS_State_Code"))
        obj.TDS_Vendor_Type = clsCommon.myCstr(dt.Rows(0)("TDS_Vendor_Type"))
        obj.TDS_Status = clsCommon.myCstr(dt.Rows(0)("TDS_Status"))
        obj.Is_TDS_Applicable = clsCommon.myCstr(dt.Rows(0)("Is_TDS_Applicable"))
        obj.Nature = clsCommon.myCstr(dt.Rows(0)("Nature"))
        obj.Actual_charges = clsCommon.myCdbl(dt.Rows(0)("Actual_charges"))
        obj.Joint_Bank_Code = clsCommon.myCstr(dt.Rows(0)("Joint_Bank_Code"))
        obj.Joint_Account_No = clsCommon.myCstr(dt.Rows(0)("Joint_Account_No"))
        If dt.Rows(0)("Start_Date") Is DBNull.Value Then
            obj.Start_Date = Nothing
        Else
            obj.Start_Date = clsCommon.myCstr(dt.Rows(0)("Start_Date"))
        End If
        If dt.Rows(0)("End_Date") Is DBNull.Value Then
            obj.End_Date = Nothing
        Else
            obj.End_Date = clsCommon.myCstr(dt.Rows(0)("End_Date"))
        End If
        obj.CSA_Type = clsCommon.myCstr(dt.Rows(0)("CSA_Type"))
        If dt.Rows(0)("Start_Period") Is DBNull.Value Then
            obj.Start_Period = Nothing
        Else
            obj.Start_Period = clsCommon.myCstr(dt.Rows(0)("Start_Period"))
        End If
        If dt.Rows(0)("Expired_Period") Is DBNull.Value Then
            obj.Expired_Period = Nothing
        Else
            obj.Expired_Period = clsCommon.myCstr(dt.Rows(0)("Expired_Period"))
        End If
        obj.PC_CODE = clsCommon.myCstr(dt.Rows(0)("PC_CODE"))
        obj.IsBlankCheque = clsCommon.myCstr(dt.Rows(0)("IsBlankCheque"))
        obj.Alies_Name = clsCommon.myCstr(dt.Rows(0)("Alies_Name"))
        obj.Vendor_Type_CHA = clsCommon.myCstr(dt.Rows(0)("Vendor_Type_CHA"))
        obj.Is_Head_Load = clsCommon.myCstr(dt.Rows(0)("Is_Head_Load"))
        obj.Rate_Head_Load = clsCommon.myCdbl(dt.Rows(0)("Rate_Head_Load"))
        obj.Service_Basis_Head_Load = clsCommon.myCstr(dt.Rows(0)("Service_Basis_Head_Load"))
        obj.Is_Own_Asset = clsCommon.myCstr(dt.Rows(0)("Is_Own_Asset"))
        obj.Rate_Own_Asset = clsCommon.myCdbl(dt.Rows(0)("Rate_Own_Asset"))
        obj.Service_Basis_Own_Asset = clsCommon.myCstr(dt.Rows(0)("Service_Basis_Own_Asset"))
        obj.IsVendorInvoiceNo = clsCommon.myCstr(dt.Rows(0)("IsVendorInvoiceNo"))
        obj.CHA_DOC_NO = clsCommon.myCstr(dt.Rows(0)("CHA_DOC_NO"))
        obj.Standard_Security_Amount = clsCommon.myCdbl(dt.Rows(0)("Standard_Security_Amount"))
        obj.Is_TC_Certified = clsCommon.myCstr(dt.Rows(0)("Is_TC_Certified"))
        obj.TC_Certified = clsCommon.myCstr(dt.Rows(0)("TC_Certified"))
        obj.MP_Code = clsCommon.myCstr(dt.Rows(0)("MP_Code"))
        obj.MP_Name = clsCommon.myCstr(dt.Rows(0)("MP_Name"))
        obj.Cheque_In_Favour_Of = clsCommon.myCstr(dt.Rows(0)("Cheque_In_Favour_Of"))
        obj.is_Drip_Saver = clsCommon.myCstr(dt.Rows(0)("is_Drip_Saver"))
        obj.Other_For_PAN = clsCommon.myCstr(dt.Rows(0)("Other_For_PAN"))
        obj.Joint_Branch_Name = clsCommon.myCstr(dt.Rows(0)("Joint_Branch_Name"))
        obj.Joint_IFSC_Code = clsCommon.myCstr(dt.Rows(0)("Joint_IFSC_Code"))
        obj.VSP_Farmer_Billing = clsCommon.myCstr(dt.Rows(0)("VSP_Farmer_Billing"))
        obj.EMP_Type = clsCommon.myCstr(dt.Rows(0)("EMP_Type"))
        obj.EMP_Fixed_Amount = clsCommon.myCdbl(dt.Rows(0)("EMP_Fixed_Amount"))
        obj.Actual_charges_Slab = clsCommon.myCdbl(dt.Rows(0)("Actual_charges_Slab"))
        obj.Actual_charges_Slab2 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges_Slab2"))
        obj.Actual_charges2 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges2"))
        obj.Actual_charges_Slab3 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges_Slab3"))
        obj.Actual_charges3 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges3"))
        obj.Actual_charges_Slab4 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges_Slab4"))
        obj.Actual_charges4 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges4"))
        obj.Actual_charges_Slab5 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges_Slab5"))
        obj.Actual_charges5 = clsCommon.myCdbl(dt.Rows(0)("Actual_charges5"))
        obj.Vendor_Distance = clsCommon.myCstr(dt.Rows(0)("Vendor_Distance"))
        obj.Apply_Mult_Incentive = clsCommon.myCdbl(dt.Rows(0)("Apply_Mult_Incentive"))
        obj.MinimumQtyRequired_pertrip = clsCommon.myCdbl(dt.Rows(0)("MinimumQtyRequired_pertrip"))
        obj.OldName = clsCommon.myCstr(dt.Rows(0)("OldName"))
        obj.Security_Deduction_Amount = clsCommon.myCdbl(dt.Rows(0)("Security_Deduction_Amount"))
        obj.Interest_Per = clsCommon.myCdbl(dt.Rows(0)("Interest_Per"))
        obj.Minimum_Interest = clsCommon.myCdbl(dt.Rows(0)("Minimum_Interest"))
        obj.SSI_No = clsCommon.myCstr(dt.Rows(0)("SSI_No"))
        obj.Is_Blacklist = clsCommon.myCdbl(dt.Rows(0)("Is_Blacklist"))
        obj.Service_Charge_Per_Unit = clsCommon.myCdbl(dt.Rows(0)("Service_Charge_Per_Unit"))
        obj.is_Hold_Payment_Process = clsCommon.myCdbl(dt.Rows(0)("is_Hold_Payment_Process"))
        obj.Is_Inactive_In_Milk_Procurement = clsCommon.myCdbl(dt.Rows(0)("Is_Inactive_In_Milk_Procurement"))
        obj.DFOption = clsCommon.myCstr(dt.Rows(0)("DFOption"))
        obj.BusinessCondition = clsCommon.myCstr(dt.Rows(0)("BusinessCondition"))
        obj.GSTRegistered = clsCommon.myCdbl(dt.Rows(0)("GSTRegistered"))
        obj.GST_Composition_scheme = clsCommon.myCdbl(dt.Rows(0)("GST_Composition_scheme"))
        obj.GSTEntity = clsCommon.myCstr(dt.Rows(0)("GSTEntity"))
        obj.GSTLastEntity = clsCommon.myCstr(dt.Rows(0)("GSTLastEntity"))
        obj.GSTFinalNo = clsCommon.myCstr(dt.Rows(0)("GSTFinalNo"))
        obj.GSTMiddle = clsCommon.myCstr(dt.Rows(0)("GSTMiddle"))
        obj.Weight = clsCommon.myCdbl(dt.Rows(0)("Weight"))
        obj.JWPriceCode = clsCommon.myCstr(dt.Rows(0)("JWPriceCode"))
        obj.CorrectionFat = clsCommon.myCdbl(dt.Rows(0)("CorrectionFat"))
        obj.CorrectionSNF = clsCommon.myCdbl(dt.Rows(0)("CorrectionSNF"))
        obj.IsEmployee = clsCommon.myCdbl(dt.Rows(0)("IsEmployee"))
        obj.Handling_Charges_Per = clsCommon.myCdbl(dt.Rows(0)("Handling_Charges_Per"))
        obj.Credit_Limit_On_Milk_Receipt_Per = clsCommon.myCdbl(dt.Rows(0)("Credit_Limit_On_Milk_Receipt_Per"))
        obj.isBulkProcurement = clsCommon.myCdbl(dt.Rows(0)("isBulkProcurement"))
        obj.isHighClass = clsCommon.myCdbl(dt.Rows(0)("isHighClass"))
        obj.Bulk_ROUTE_NO = clsCommon.myCstr(dt.Rows(0)("Bulk_ROUTE_NO"))
        obj.Monthly_Rent = clsCommon.myCdbl(dt.Rows(0)("Monthly_Rent"))
        obj.EMP_CODE = clsCommon.myCstr(dt.Rows(0)("EMP_CODE"))
        obj.TIP_Buffalo = clsCommon.myCdbl(dt.Rows(0)("TIP_Buffalo"))
        obj.TIP_Cow = clsCommon.myCdbl(dt.Rows(0)("TIP_Cow"))
        obj.TIP_Mix = clsCommon.myCdbl(dt.Rows(0)("TIP_Mix"))
        obj.Registration_No = clsCommon.myCstr(dt.Rows(0)("Registration_No"))
        obj.Aadhar_No = clsCommon.myCstr(dt.Rows(0)("Aadhar_No"))
        obj.Care_Of = clsCommon.myCstr(dt.Rows(0)("Care_Of"))
        obj.IsTCSnotApplicable = clsCommon.myCdbl(dt.Rows(0)("IsTCSnotApplicable"))
        obj.Isbuyerfilereturninlasttwoyears = clsCommon.myCdbl(dt.Rows(0)("Isbuyerfilereturninlasttwoyears"))
        obj.IsTCS_TDSamountgreaterthan50KpreviousYear = clsCommon.myCdbl(dt.Rows(0)("IsTCS_TDSamountgreaterthan50KpreviousYear"))
        obj.SecChequeNoLac1 = clsCommon.myCstr(dt.Rows(0)("SecChequeNoLac1"))
        obj.SecChequeNoRs100 = clsCommon.myCstr(dt.Rows(0)("SecChequeNoRs100"))
        obj.TDS_Branch_Code_Service = clsCommon.myCstr(dt.Rows(0)("TDS_Branch_Code_Service"))
        obj.Deduction_Code_Service = clsCommon.myCstr(dt.Rows(0)("Deduction_Code_Service"))
        obj.TDS_Status_Service = clsCommon.myCstr(dt.Rows(0)("TDS_Status_Service"))
        obj.TDS_Vendor_Type_Service = clsCommon.myCstr(dt.Rows(0)("TDS_Vendor_Type_Service"))
        obj.TDS_State_Code_Service = clsCommon.myCstr(dt.Rows(0)("TDS_State_Code_Service"))
        obj.DistanceKM_Head_Load = clsCommon.myCdbl(dt.Rows(0)("DistanceKM_Head_Load"))
        obj.Vendor_Name_Hindi = clsCommon.myCstr(dt.Rows(0)("Vendor_Name_Hindi"))
        obj.Registered_PDCS_CLUSTER = clsCommon.myCstr(dt.Rows(0)("Registered_PDCS_CLUSTER"))
        If dt.Rows(0)("StartDate") Is DBNull.Value Then
            obj.StartDate = Nothing
        Else
            obj.StartDate = clsCommon.myCstr(dt.Rows(0)("StartDate"))
        End If
        obj.BankCode2 = clsCommon.myCstr(dt.Rows(0)("BankCode2"))
        obj.BankName2 = clsCommon.myCstr(dt.Rows(0)("BankName2"))
        obj.Credit2 = clsCommon.myCdbl(dt.Rows(0)("Credit2"))
        obj.IFSCCode2 = clsCommon.myCstr(dt.Rows(0)("IFSCCode2"))
        obj.AccNo2 = clsCommon.myCstr(dt.Rows(0)("AccNo2"))
        obj.AccountType2 = clsCommon.myCstr(dt.Rows(0)("AccountType2"))
        obj.BankBranch2 = clsCommon.myCstr(dt.Rows(0)("BankBranch2"))
        obj.SecurityCharges2 = clsCommon.myCstr(dt.Rows(0)("SecurityCharges2"))
        obj.SupervisorOrRP = clsCommon.myCstr(dt.Rows(0)("SupervisorOrRP"))
        If dt.Rows(0)("Active_Date") Is DBNull.Value Then
            obj.Active_Date = Nothing
        Else
            obj.Active_Date = clsCommon.myCDate(dt.Rows(0)("Active_Date"))
        End If
        obj.Gender = clsCommon.myCstr(dt.Rows(0)("Gender"))
        obj.RegistrationNo = clsCommon.myCstr(dt.Rows(0)("RegistrationNo"))
        If dt.Rows(0)("RegistrationDate") Is DBNull.Value Then
            obj.RegistrationDate = Nothing
        Else
            obj.RegistrationDate = clsCommon.myCstr(dt.Rows(0)("RegistrationDate"))
        End If
        obj.DISTRICT_Code = clsCommon.myCstr(dt.Rows(0)("DISTRICT_Code"))
        obj.Zone_Code = clsCommon.myCstr(dt.Rows(0)("Zone_Code"))
        obj.CAST_CATEGORY_CODE = clsCommon.myCstr(dt.Rows(0)("CAST_CATEGORY_CODE"))
        obj.BLOCK_CODE = clsCommon.myCstr(dt.Rows(0)("BLOCK_CODE"))
        obj.Company_Bank = clsCommon.myCstr(dt.Rows(0)("Company_Bank"))
        obj.REVENUE_VILLAGE_CODE = clsCommon.myCstr(dt.Rows(0)("REVENUE_VILLAGE_CODE"))
        obj.GRAMPANCHAYAT_CODE = clsCommon.myCstr(dt.Rows(0)("GRAMPANCHAYAT_CODE"))
        obj.PANCHAYAT_SAMITI_CODE = clsCommon.myCstr(dt.Rows(0)("PANCHAYAT_SAMITI_CODE"))
        obj.VIDHAN_SABHA_CODE = clsCommon.myCstr(dt.Rows(0)("VIDHAN_SABHA_CODE"))
        obj.IsAllowSkip_PurchaseQC = clsCommon.myCdbl(dt.Rows(0)("IsAllowSkipPurchaseQC"))
        obj.OEM = clsCommon.myCdbl(dt.Rows(0)("OEM"))
        obj.Is_Provisional = clsCommon.myCdbl(dt.Rows(0)("Is_Provisional"))
        obj.Company_Bank_Current = clsCommon.myCstr(dt.Rows(0)("Company_Bank_Current"))
        obj.In_Active_CF = clsCommon.myCstr(dt.Rows(0)("In_Active_CF"))
        obj.Is_Default_Grower = clsCommon.myCdbl(dt.Rows(0)("Is_Default_Grower"))

        Return obj

    End Function

End Class
Public Structure structVendorFilter
    Dim VendorWise As Boolean
    Dim VendorGroupWise As Boolean
    Dim FormType As String
    Dim DocumentWise As Boolean
    Dim IncludeApplyDoc As Boolean
    Dim CurrencyType As String
    Dim strPortrait As Boolean
    Dim strLandscape As Boolean
    Dim IsOnlyForAgainstSalary As Boolean
    Dim strtempBaseQryforopening As String
    Dim strtempBaseQryforopeningForMIS As String
    Dim IS_GIT As Boolean
End Structure


Public Class clsShipToLocation
    Public Ship_To_Code As String = ""
    Public Ship_To_Desc As String = ""
    Public Ship_To_Type As String = ""
    Public Ship_To_Type_Code As String = ""
    Public Ship_To_Type_Desc As String = ""
    Public Add1 As String = ""
    Public Add2 As String = ""
    Public Add3 As String = ""
    Public Add4 As String = ""
    Public City_Code As String = ""
    Public State As String = ""
    Public Pin_Code As String = ""
    Public Country As String = ""
    Public Telphone As String = ""
    Public Email As String = ""

    Public Shared Function GetName(ByVal strCode As String, ByVal Trans As SqlTransaction) As String
        Dim qry As String = "select Ship_To_Desc from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" + strCode + "'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Trans))
    End Function


    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " select TSPL_SHIP_TO_LOCATION.Ship_To_Code as [Code] ,TSPL_SHIP_TO_LOCATION.Ship_To_Desc as [Ship To Description] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type as [Ship To Type] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code as [Ship To Type Code] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Desc as [Ship To Type Description] ,TSPL_SHIP_TO_LOCATION.Add1 as [Address1] ,TSPL_SHIP_TO_LOCATION.Add2 as [Address2] ,TSPL_SHIP_TO_LOCATION.Add3 as [Address3] ,TSPL_SHIP_TO_LOCATION.Add4 as [Address4] ,TSPL_SHIP_TO_LOCATION.City_Code as [City Code] ,TSPL_SHIP_TO_LOCATION.State as [State] ,TSPL_SHIP_TO_LOCATION.Pin_Code as [Pin Code] ,TSPL_SHIP_TO_LOCATION.Country as [Country] ,TSPL_SHIP_TO_LOCATION.Telphone as [Telphone] ,TSPL_SHIP_TO_LOCATION.Email as [Email] ,TSPL_SHIP_TO_LOCATION.Created_By as [Created By] ,TSPL_SHIP_TO_LOCATION.Created_Date as [Created Date] ,TSPL_SHIP_TO_LOCATION.Modify_By as [Modify By] ,TSPL_SHIP_TO_LOCATION.Modify_Date as [Modify Date] ,TSPL_SHIP_TO_LOCATION.Comp_Code as [Company Code] ,TSPL_SHIP_TO_LOCATION.Tin_No as [Tin No] ,TSPL_SHIP_TO_LOCATION.CST_No as [Cst No] ,TSPL_SHIP_TO_LOCATION.Loc_Code as [Location Code],TSPL_LOCATION_MASTER.location_desc as [Location Description]  From TSPL_SHIP_TO_LOCATION  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SHIP_TO_LOCATION.loc_code "
        str = clsCommon.ShowSelectForm("SHPTOLOCFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    '----------------End of Code For Get Finder--------------------------------------------------------------'
    Public Shared Function GetData(ByVal strCode As String, ByVal NavType As NavigatorType) As clsShipToLocation
        Dim obj As clsShipToLocation = Nothing
        Dim qry As String = "select TSPL_SHIP_TO_LOCATION.* from TSPL_SHIP_TO_LOCATION where 2=2 "
        Select Case NavType
            Case NavigatorType.First
                qry += " and Ship_To_Code = (select MIN(Ship_To_Code) from TSPL_SHIP_TO_LOCATION )"
            Case NavigatorType.Last
                qry += " and Ship_To_Code = (select Max(Ship_To_Code) from TSPL_SHIP_TO_LOCATION )"
            Case NavigatorType.Next
                qry += " and Ship_To_Code = (select Min(Ship_To_Code) from TSPL_SHIP_TO_LOCATION where Ship_To_Code > '" + strCode + "')"
            Case NavigatorType.Previous
                qry += " and Ship_To_Code = (select Max(Ship_To_Code) from TSPL_SHIP_TO_LOCATION where Ship_To_Code < '" + strCode + "')"
            Case NavigatorType.Current
                qry += " and Ship_To_Code = '" + strCode + "'"
        End Select
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            obj = New clsShipToLocation()
            obj.Ship_To_Code = clsCommon.myCstr(dt.Rows(0)("Ship_To_Code"))
            obj.Ship_To_Desc = clsCommon.myCstr(dt.Rows(0)("Ship_To_Desc"))
            obj.Ship_To_Type = clsCommon.myCstr(dt.Rows(0)("Ship_To_Type"))
            obj.Ship_To_Type_Code = clsCommon.myCstr(dt.Rows(0)("Ship_To_Type_Code"))
            obj.Ship_To_Type_Desc = clsCommon.myCstr(dt.Rows(0)("Ship_To_Type_Desc"))
            obj.Add1 = clsCommon.myCstr(dt.Rows(0)("Add1"))
            obj.Add2 = clsCommon.myCstr(dt.Rows(0)("Add2"))
            obj.Add3 = clsCommon.myCstr(dt.Rows(0)("Add3"))
            obj.Add4 = clsCommon.myCstr(dt.Rows(0)("Add4"))
            obj.City_Code = clsCommon.myCstr(dt.Rows(0)("City_Code"))
            obj.State = clsCommon.myCstr(dt.Rows(0)("State"))
            obj.Pin_Code = clsCommon.myCstr(dt.Rows(0)("Pin_Code"))
            obj.Country = clsCommon.myCstr(dt.Rows(0)("Country"))
            obj.Telphone = clsCommon.myCstr(dt.Rows(0)("Telphone"))
            obj.Email = clsCommon.myCstr(dt.Rows(0)("Email"))
        End If
        Return obj
    End Function

    Public Shared Function SaveData(ByVal obj As clsShipToLocation, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isNewEntry As Boolean = False
            Dim strDocNo As String = ""
            If clsCommon.myLen(obj.Ship_To_Code) <= 0 Then
                obj.Ship_To_Code = clsERPFuncationality.GetNextCode(trans, clsCommon.GETSERVERDATE(trans), clsDocType.ShipToMaster, "", "")
                isNewEntry = True
            End If

            If (clsCommon.myLen(obj.Ship_To_Code) <= 0) Then
                Throw New Exception("Error in Document Code Generation")
            End If

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Ship_To_Desc", obj.Ship_To_Desc)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Type", obj.Ship_To_Type)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Type_Code", obj.Ship_To_Type_Code)
            clsCommon.AddColumnsForChange(coll, "Ship_To_Type_Desc", obj.Ship_To_Type_Desc)
            clsCommon.AddColumnsForChange(coll, "Add1", obj.Add1)
            clsCommon.AddColumnsForChange(coll, "Add2", obj.Add2)
            clsCommon.AddColumnsForChange(coll, "Add3", obj.Add3)
            clsCommon.AddColumnsForChange(coll, "Add4", obj.Add4)
            clsCommon.AddColumnsForChange(coll, "City_Code", obj.City_Code)
            clsCommon.AddColumnsForChange(coll, "State", obj.State)
            clsCommon.AddColumnsForChange(coll, "Pin_Code", obj.Pin_Code)
            clsCommon.AddColumnsForChange(coll, "Country", obj.Country)
            clsCommon.AddColumnsForChange(coll, "Telphone", obj.Telphone)
            clsCommon.AddColumnsForChange(coll, "Email", obj.Email)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
            If isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Ship_To_Code", obj.Ship_To_Code)
                clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
                clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SHIP_TO_LOCATION", OMInsertOrUpdate.Insert, "", trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SHIP_TO_LOCATION", OMInsertOrUpdate.Update, "TSPL_SHIP_TO_LOCATION.Ship_To_Code='" + obj.Ship_To_Code + "'", trans)
            End If
        Catch err As Exception
            Throw New Exception(err.Message)
        End Try
        Return True
    End Function
End Class
Public Class clsTaxRateMaster
    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Dim qry As String = " select Tax_Code as [Code],Tax_Type as [Tax Type],Tax_Rate_Code as [Tax Rate Code],Tax_Rate_Desc as [Tax Rate Description],Tax_Rate as [Tax Rate],Created_By as [Created By],Created_Date as [Created By],Modify_By as [Modify By],Modify_Date as [Modify Date],Comp_Code as [Company Code] from tspl_tax_rates "
        str = clsCommon.ShowSelectForm("RPTXRTFND", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Return str
    End Function
    '----------------End of Code For Get Finder--------------------------------------------------------------'

    Public Shared Function GetGLAccountForPurchase(ByVal strTaxCode As String, ByVal dblTaxRate As Double, ByVal Trans As SqlTransaction, ByVal isThrowExceptionIfNotExist As Boolean) As String
        Dim qry As String = " select GL_Account_Code from TSPL_TAX_RATES where Tax_Code='" + strTaxCode + "' and Tax_Type='P' and Tax_Rate='" + clsCommon.myCdbl(dblTaxRate) + "'"
        Dim strGLCode As String = clsDBFuncationality.getSingleValue(qry, Trans)
        If isThrowExceptionIfNotExist AndAlso clsCommon.myLen(strGLCode) <= 0 Then
            Throw New Exception("Please set GL Account for Purchase Tax Authority" + strTaxCode + " With Tax Rate" + clsCommon.myCstr(dblTaxRate))
        End If
        Return strGLCode
    End Function

    Public Shared Function GetGLAccountForSale(ByVal strTaxCode As String, ByVal dblTaxRate As Double, ByVal Trans As SqlTransaction, ByVal isThrowExceptionIfNotExist As Boolean) As String
        Dim qry As String = " select GL_Account_Code from TSPL_TAX_RATES where Tax_Code='" + strTaxCode + "' and Tax_Type='S' and Tax_Rate='" + clsCommon.myCdbl(dblTaxRate) + "'"
        Dim strGLCode As String = clsDBFuncationality.getSingleValue(qry, Trans)
        If isThrowExceptionIfNotExist AndAlso clsCommon.myLen(strGLCode) <= 0 Then
            Throw New Exception("Please set GL Account for Sale Tax Authority" + strTaxCode + " With Tax Rate" + clsCommon.myCstr(dblTaxRate))
        End If
        Return strGLCode
    End Function
End Class
Public Class clsItemWithBalanceFINDER
    Public Shared Function GetBalance(ByVal LocCode As String) As String
        Dim qry As String = "select ICode,SUM(Qty * case when TransType='' then 1 else 0 end)as [Qty On Hand],SUM(Qty * case when TransType='' then 0 else 1 end)as CommitQty,SUM(Qty *RI )as [Available Balance] from ( "
        qry += "select distinct xxx.TransType,xxx.TransCode,xxx.DocNo, xxx.ICode,xxx.Location,xxx.MRP,xxx.RI,xxx.Qty from( "
        qry += "select xx.TransType,xx.TransCode,xx.DocNo,xx.ICode,xx.Location,xx.MRP as MRPOld,xx.Qty as OldQty,xx.RI,'' as UOM, "
        qry += "(case when ''='FB' then xx.Qty * (select Conversion_Factor from TSPL_ITEM_UOM_DETAIL where TSPL_ITEM_UOM_DETAIL.Item_Code=xx.ICode and TSPL_ITEM_UOM_DETAIL.UOM_Code= (case when xx.UOM_Code='FB' then 'FC' else case when xx.UOM_Code='FC' then 'FB' end end) ) else case when ''='FC' then xx.Qty /(select Conversion_Factor from TSPL_ITEM_UOM_DETAIL where TSPL_ITEM_UOM_DETAIL.Item_Code=xx.ICode and TSPL_ITEM_UOM_DETAIL.UOM_Code=xx.UOM_Code) else 0 end end) as Qty, "
        qry += "(case when 'FC'='FB' then xx.MRP / (select Conversion_Factor from TSPL_ITEM_UOM_DETAIL where TSPL_ITEM_UOM_DETAIL.Item_Code=xx.ICode and TSPL_ITEM_UOM_DETAIL.UOM_Code= (case when xx.UOM_Code='FB' then 'FC' else case when xx.UOM_Code='FC' then 'FB' end end) ) else case when 'FC'='FC' then xx.MRP *(select Conversion_Factor from TSPL_ITEM_UOM_DETAIL where TSPL_ITEM_UOM_DETAIL.Item_Code=xx.ICode and TSPL_ITEM_UOM_DETAIL.UOM_Code=xx.UOM_Code) else 0 end end) as MRP from ( "
        qry += "select '' as TransType,'' as TransCode,'' as DocNo, Item_Code as ICode,Location_Code as Location,MRPNew as MRP  ,SUM(QtyNew*RI) as Qty,1 as RI,max(UOMNew) as UOM_Code from( "
        qry += "select Trans_Id,Item_Code ,Location_Code,case when InOut='I' then 1 else -1 end as RI,Qty/Conversion_Factor as QtyNew,MRP*Conversion_Factor as MRPNew,UOMNew from( "
        qry += "select TSPL_INVENTORY_MOVEMENT.Trans_Id, TSPL_INVENTORY_MOVEMENT.Item_Code ,TSPL_INVENTORY_MOVEMENT.Location_Code , TSPL_INVENTORY_MOVEMENT.InOut,TSPL_INVENTORY_MOVEMENT.Qty ,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,TSPL_INVENTORY_MOVEMENT.MRP,TSPL_INVENTORY_MOVEMENT.UOM ,case when Conversion_Factor=1 then TSPL_ITEM_UOM_DETAIL.UOM_Code else '' end as UOMNew from "
        qry += "TSPL_INVENTORY_MOVEMENT  left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_INVENTORY_MOVEMENT.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_INVENTORY_MOVEMENT.UOM where TSPL_INVENTORY_MOVEMENT.Qty<>0  "
        qry += ")xxx   )xxxx group by Item_Code,Location_Code,MRPNew "
        qry += " union all "
        qry += "select 'Shipment' as TransType,'LoadOut' as TransCode,TSPL_SHIPMENT_MASTER.Shipment_No as DocNo, TSPL_SHIPMENT_DETAILS.Item_Code as ICode,TSPL_SHIPMENT_MASTER.Location as Location,TSPL_SHIPMENT_DETAILS.MRP_Amt as MRP,TSPL_SHIPMENT_DETAILS.Shipped_Qty as Qty,-1 as RI,TSPL_SHIPMENT_DETAILS.Unit_code from TSPL_SHIPMENT_DETAILS left outer join TSPL_SHIPMENT_MASTER on TSPL_SHIPMENT_MASTER.Shipment_No= TSPL_SHIPMENT_DETAILS.Shipment_No where TSPL_SHIPMENT_MASTER.Is_Post='N' and TSPL_SHIPMENT_MASTER.Shipment_Type='Sale' and TSPL_SHIPMENT_DETAILS.Shipped_Qty <>0   and TSPL_SHIPMENT_MASTER.Shipment_No not in ('') "
        qry += " union all "
        qry += "select  'Transfer' as TransType,'MMTrans' as TransCode,TSPL_TRANSFER_HEAD.Transfer_No as DocNo, TSPL_TRANSFER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_HEAD.From_Location as Locaion,TSPL_TRANSFER_DETAIL.MRP,TSPL_TRANSFER_DETAIL.Item_Qty as Qty,-1 as RI ,TSPL_TRANSFER_DETAIL.Uom from TSPL_TRANSFER_DETAIL left outer join TSPL_TRANSFER_HEAD on TSPL_TRANSFER_HEAD.Transfer_No= TSPL_TRANSFER_DETAIL.Transfer_No where TSPL_TRANSFER_HEAD.Post='N' and TSPL_TRANSFER_HEAD.Transfer_Type='LO' and TSPL_TRANSFER_DETAIL.Item_Qty<>0 and TSPL_TRANSFER_DETAIL.Transfer_No not in ('') "
        qry += " union all "
        qry += "select  'Purchase Return' as TransType,'PurchaseReturn' as TransCode,TSPL_PR_HEAD.PR_No as DocNo, TSPL_PR_DETAIL.Item_Code as ICode,TSPL_PR_DETAIL.Location as Locaion,TSPL_PR_DETAIL.MRP,TSPL_PR_DETAIL.PR_Qty as Qty,-1 as RI,TSPL_PR_DETAIL.Unit_code AS Uom from TSPL_PR_DETAIL left outer join TSPL_PR_HEAD on TSPL_PR_HEAD.PR_No=TSPL_PR_DETAIL.PR_No where TSPL_PR_HEAD.Status=0 and TSPL_PR_DETAIL.PR_Qty<>0  and TSPL_PR_DETAIL.PR_No not in ('') "
        qry += " union all "
        qry += "select 'IC-AD' as TransType,'ICAdj' as TransCode,TSPL_ADJUSTMENT_HEADER.Adjustment_No as DocNo, TSPL_ADJUSTMENT_DETAIL.Item_Code as ICode,TSPL_ADJUSTMENT_HEADER.Loc_Code as Locaion,TSPL_ADJUSTMENT_DETAIL.mrp,TSPL_ADJUSTMENT_DETAIL.Item_Quantity as Qty,-1 as RI,TSPL_ADJUSTMENT_DETAIL.Unit_Code AS Uom from TSPL_ADJUSTMENT_DETAIL left outer join TSPL_ADJUSTMENT_HEADER on TSPL_ADJUSTMENT_HEADER.Adjustment_No=TSPL_ADJUSTMENT_DETAIL.Adjustment_No where TSPL_ADJUSTMENT_HEADER.Posted='N' and TSPL_ADJUSTMENT_DETAIL.Item_Quantity<>0  and TSPL_ADJUSTMENT_DETAIL.Adjustment_Type in ('BD','QD') and TSPL_ADJUSTMENT_HEADER.Adjustment_No not in ('') "
        qry += " union all "
        qry += "select 'Warehouse Breakage' as TransType,'WareHouseBreakage' as TransCode,TSPL_WH_BREAKAGE_HEAD.Document_No as DocNo,  TSPL_WH_BREAKAGE_DETAIL.Item_Code as ICode,TSPL_WH_BREAKAGE_HEAD.Loc_Code as Locaion,TSPL_WH_BREAKAGE_DETAIL.mrp, Breakage_Qty+Leakage_Qty+Shortage_Qty as Qty, -1 as RI, TSPL_WH_BREAKAGE_DETAIL.Unit_Code AS Uom from TSPL_WH_BREAKAGE_DETAIL left outer join TSPL_WH_BREAKAGE_HEAD  on TSPL_WH_BREAKAGE_HEAD.Document_No =TSPL_WH_BREAKAGE_DETAIL.Document_No where TSPL_WH_BREAKAGE_HEAD.Is_Post<>1 and TSPL_WH_BREAKAGE_DETAIL.Item_Quantity<>0 AND TSPL_WH_BREAKAGE_HEAD.Document_No  not in ('') "
        qry += " )xx)xxx where 2=2 "

        qry += " union all "

        qry += "select  xx.TransType,xx.TransCode,xx.DocNo, xx.ICode,xx.Location,xx.RI ,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,(FinalUOM.Conversion_Factor*(xx.Qty*1/TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) as Qty from ( "
        qry += "select '' as TransType,'' as TransCode,'' as DocNo, Item_Code as ICode,Location_Code as Location ,SUM(QtyNew*RI) as Qty,1 as RI,UOMNew as UOM  from( "
        qry += "select Trans_Id,Item_Code ,Location_Code,case when InOut='I' then 1 else -1 end as RI,Qty as QtyNew,UOMNew from( select TSPL_INVENTORY_MOVEMENT.Trans_Id, TSPL_INVENTORY_MOVEMENT.Item_Code ,TSPL_INVENTORY_MOVEMENT.Location_Code , TSPL_INVENTORY_MOVEMENT.InOut,TSPL_INVENTORY_MOVEMENT.Qty   ,TSPL_INVENTORY_MOVEMENT.UOM as UOMNew  from TSPL_INVENTORY_MOVEMENT  where TSPL_INVENTORY_MOVEMENT.Qty<>0 "
        qry += ")xxx   )xxxx group by Item_Code,Location_Code,UOMNew "
        qry += " union all "
        qry += "select 'Purchase Return' as TransType,'PurchaseReturn' as TransCode,TSPL_PR_HEAD.PR_No as DocNo, TSPL_PR_DETAIL.Item_Code as ICode,TSPL_PR_DETAIL.Location as Locaion,TSPL_PR_DETAIL.PR_Qty as Qty,-1 as RI,TSPL_PR_DETAIL.Unit_code AS Uom  from TSPL_PR_DETAIL  left outer join TSPL_PR_HEAD on TSPL_PR_HEAD.PR_No=TSPL_PR_DETAIL.PR_No where TSPL_PR_HEAD.Status=0 and TSPL_PR_DETAIL.PR_Qty<>0   and TSPL_PR_DETAIL.PR_No not in ('') "
        qry += " union all "
        qry += "select 'IC-AD' as TransType,'ICAdj' as TransCode,TSPL_ADJUSTMENT_HEADER.Adjustment_No as DocNo, TSPL_ADJUSTMENT_DETAIL.Item_Code as ICode,TSPL_ADJUSTMENT_HEADER.Loc_Code as Locaion,TSPL_ADJUSTMENT_DETAIL.Item_Quantity as Qty,-1 as RI,TSPL_ADJUSTMENT_DETAIL.Unit_Code AS Uom  from TSPL_ADJUSTMENT_DETAIL  left outer join TSPL_ADJUSTMENT_HEADER on TSPL_ADJUSTMENT_HEADER.Adjustment_No=TSPL_ADJUSTMENT_DETAIL.Adjustment_No where TSPL_ADJUSTMENT_HEADER.Posted='N' and TSPL_ADJUSTMENT_DETAIL.Item_Quantity<>0  and TSPL_ADJUSTMENT_DETAIL.Adjustment_Type in ('BD','QD') and TSPL_ADJUSTMENT_HEADER.Adjustment_No not in ('') "
        qry += " union all "
        qry += "select 'RGP' as TransType,'RGP' as TransCode,TSPL_RGP_HEAD.RGP_No as DocNo, TSPL_RGP_DETAIL.Item_Code as ICode,TSPL_RGP_HEAD.Location as Locaion,TSPL_RGP_DETAIL.RGP_Qty as Qty,-1 as RI,TSPL_RGP_DETAIL.Unit_code AS Uom  from TSPL_RGP_DETAIL  left outer join TSPL_RGP_HEAD on TSPL_RGP_HEAD.RGP_No=TSPL_RGP_DETAIL.RGP_No where TSPL_RGP_HEAD.Status=0 and TSPL_RGP_DETAIL.RGP_Qty<>0   and TSPL_RGP_DETAIL.RGP_No not in ('') "
        qry += " union all "
        qry += "select 'Scrap' as TransType,'ScrapInvoice' as TransCode,TSPL_SCRAPSALE_HEAD.shipment_No as DocNo, TSPL_SCRAPSALE_DETAIL.Item_Code as ICode,TSPL_SCRAPSALE_HEAD.Loc_Code as Locaion,TSPL_SCRAPSALE_DETAIL.shipped_Qty as Qty,-1 as RI,TSPL_SCRAPSALE_DETAIL.Unit_code AS Uom  from TSPL_SCRAPSALE_DETAIL  left outer join TSPL_SCRAPSALE_HEAD on TSPL_SCRAPSALE_HEAD.shipment_No=TSPL_SCRAPSALE_DETAIL.shipment_No where TSPL_SCRAPSALE_HEAD.IsPost=0 and TSPL_SCRAPSALE_DETAIL.shipped_Qty<>0 and TSPL_SCRAPSALE_DETAIL.shipment_No not in ('') "
        qry += " union all "
        qry += "select 'Issue/Return/Transfer' as TransType,'IssueReturnTransfer' as TransCode,TSPL_IssueReturn_HEAD.Doc_No as DocNo, TSPL_IssueReturn_DETAIL.Item_Code as ICode,TSPL_IssueReturn_HEAD.From_Location as Locaion,TSPL_IssueReturn_DETAIL.Issued_Qty as Qty,-1 as RI,TSPL_IssueReturn_DETAIL.Unit_code AS Uom  from TSPL_IssueReturn_DETAIL  left outer join TSPL_IssueReturn_HEAD on TSPL_IssueReturn_HEAD.Doc_No=TSPL_IssueReturn_DETAIL.Doc_No where TSPL_IssueReturn_HEAD.Status=0 and TSPL_IssueReturn_DETAIL.Issued_Qty<>0 and TSPL_IssueReturn_DETAIL.Doc_No not in ('') "
        qry += " union all "
        qry += "select 'Assemblies' as TransType,'Assemblies' as TransCode,TSPL_PJC_ASSEMBLIES.CODE as DocNo, Main_Item_Code as ICode,LOCATION_CODE as Location,QUANTITY,(case when TRANSACTION_TYPE='Assembly' then 1  else -1 end) as RI, BUILD_ITEM_UNIT_CODE as UnitCode from TSPL_PJC_ASSEMBLIES where TSPL_PJC_ASSEMBLIES.POSTED=0 and TSPL_PJC_ASSEMBLIES.CODE  not in ('') "
        qry += " union all "
        qry += "select 'MF Issue' as TransType,'MFIssue' as TransCode,TSPL_MF_ISSUE_DETAIL.ISSUE_CODE as DocNo ,TSPL_MF_ISSUE_DETAIL.ITEM_CODE as ICode, TSPL_MF_ISSUE.LOCATION_CODE as Location,TSPL_MF_ISSUE_DETAIL.ISSUE_QTY as Qty,1 AS RI,TSPL_MF_ISSUE_DETAIL.UNIT_CODE as UOM  from TSPL_MF_ISSUE_DETAIL INNER JOIN TSPL_MF_ISSUE ON TSPL_MF_ISSUE_DETAIL.ISSUE_CODE=TSPL_MF_ISSUE.ISSUE_CODE  WHERE TSPL_MF_ISSUE.POSTED=0 AND TSPL_MF_ISSUE.ISSUE_CODE NOT IN ('') "
        qry += " union all "
        qry += "select 'MF Issue' as TransType,'MFIssue' as TransCode,TSPL_MF_ISSUE_DETAIL.ISSUE_CODE as DocNo,TSPL_MF_ISSUE_DETAIL.ITEM_CODE as ICode, TSPL_MF_ISSUE.LOCATION_CODE_FROM as Location,TSPL_MF_ISSUE_DETAIL.ISSUE_QTY as Qty,-1 AS RI,TSPL_MF_ISSUE_DETAIL.UNIT_CODE as UOM from TSPL_MF_ISSUE_DETAIL INNER JOIN TSPL_MF_ISSUE ON TSPL_MF_ISSUE_DETAIL.ISSUE_CODE=TSPL_MF_ISSUE.ISSUE_CODE WHERE TSPL_MF_ISSUE.POSTED=0 AND TSPL_MF_ISSUE.ISSUE_CODE NOT IN ('') "
        qry += " union all "
        qry += "select 'MF Return' as TransType,'MFReturn' as TransCode,TSPL_MF_RETURN.RETURN_CODE  as DocNo, TSPL_MF_RETURN_DETAIL.ITEM_CODE as ICode,TSPL_MF_RETURN.LOCATION_CODE as Location,TSPL_MF_RETURN_DETAIL.RETURN_QTY as Qty,1 AS RI,TSPL_MF_RETURN_DETAIL.UNIT_CODE as UOM from TSPL_MF_RETURN_DETAIL INNER JOIN TSPL_MF_RETURN ON TSPL_MF_RETURN_DETAIL.RETURN_CODE=TSPL_MF_RETURN.RETURN_CODE  WHERE TSPL_MF_RETURN.POSTED=0 AND TSPL_MF_RETURN.RETURN_CODE NOT IN ('') "
        qry += " union all "
        qry += "select 'MF Return' as TransType,'MFReturn' as TransCode,TSPL_MF_RETURN.RETURN_CODE as DocNo,TSPL_MF_RETURN_DETAIL.ITEM_CODE as ICode,TSPL_MF_RETURN.LOCATION_CODE_FROM as Location,TSPL_MF_RETURN_DETAIL.RETURN_QTY as Qty,-1 AS RI,TSPL_MF_RETURN_DETAIL.UNIT_CODE as UOM from TSPL_MF_RETURN_DETAIL INNER JOIN TSPL_MF_RETURN ON TSPL_MF_RETURN_DETAIL.RETURN_CODE=TSPL_MF_RETURN.RETURN_CODE  WHERE TSPL_MF_RETURN.POSTED=0 AND TSPL_MF_RETURN.RETURN_CODE NOT IN ('') "
        qry += " union all "
        qry += "select  'Assemblies' as TransType,'Assemblies' as TransCode,TSPL_PJC_ASSEMBLIES.CODE as DocNo, TSPL_MF_BOM_DETAIL.CONSM_ITEM_CODE AS ICode,TSPL_PJC_ASSEMBLIES.LOCATION_CODE as Location, TSPL_MF_BOM_DETAIL.CONSM_QUANTITY*(TSPL_PJC_ASSEMBLIES.QUANTITY/TSPL_PJC_ASSEMBLIES.BUILD_QUANTITY) as Qty, (case when TSPL_PJC_ASSEMBLIES.TRANSACTION_TYPE='Assembly' then  -1 else  1 end) AS RI, TSPL_MF_BOM_DETAIL.CONSM_ITEM_UNIT_CODE as UnitCode from TSPL_PJC_ASSEMBLIES  inner join TSPL_MF_BOM_HEAD on TSPL_MF_BOM_HEAD.BOM_CODE=TSPL_PJC_ASSEMBLIES.BOM_CODE inner JOIN TSPL_MF_BOM_DETAIL ON TSPL_MF_BOM_HEAD.BOM_CODE=TSPL_MF_BOM_DETAIL.BOM_CODE  where TSPL_PJC_ASSEMBLIES.POSTED=0  and TSPL_PJC_ASSEMBLIES.CODE  not in ('') "
        qry += ")xx left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=xx.ICode and TSPL_ITEM_UOM_DETAIL.UOM_Code=xx.UOM  left outer join TSPL_ITEM_UOM_DETAIL as FinalUOM on FinalUOM.Item_Code=xx.ICode "
        qry += ")FinalQry "
        If LocCode IsNot Nothing AndAlso clsCommon.myLen(LocCode) > 0 Then
            qry += "where FinalQry.Location in ('" + LocCode + "') "
        End If
        qry += "group by ICode"

        Return qry
    End Function

    Public Shared Function GetFinder(ByVal whrCls As String, ByVal strCurrCode As String, ByVal isButtonClicked As Boolean, ByVal LocationCode As String) As String
        Dim Icode As String = ""
        Dim BalQry As String = GetBalance(LocationCode)

        Dim qry As String = " select Item_Code as [Code] ,Item_Desc as [Item Description],BalQry.[Qty On Hand],BalQry.CommitQty as [Committed Qty],BalQry.[Available Balance] ,Structure_Code as [Structure Code] ,Structure_Desc as [Structure Description] ,Purchase_Class_Code as [Purchase Class Code] ,Sale_Class_Code as [Sale Class Code] ,Unit_Code as [Unit Code] ,Deafult_Price as [Deafult Price] ,Father_Code as [Father Code] ,Father_QTy as [Father Quantity] ,Cheapter_Heads as [Cheapter Heads] ,Mother_Code as [Mother Code] ,Mother_Qty as [Mother Quantity] ,Opening_Balance as [Opening Balance] ,Two_Count_Status as [Two Count Status] ,Three_Count_Status as [Three Count Status] ,Server_Type as [Server Type] ,Mfg_Date as [Manufacturing Date] ,Batch_No as [Batch No] ,Best_Befor_UseDate as [Best Befor Use Date] ,Item_Type as [Item Type] ,Created_By as [Created By] ,Created_Date as [Created Date] ,Modify_By as [Modify By] ,Modify_Date as [Modify Date] ,Comp_Code as [Company Code] ,Flavour_Seq as [Flavour Sequence] ,Pack_Seq as [Pack Sequence] ,Sku_Seq as [SKU Sequence] ,show as [Show] ,item_category as [Item Category],tech_shelf_life as [Tech Shelf Life] ,Cost as [Cost] ,Sub_item_category as [Item Sub Category] ,TypeOfItm as [Type Of Item] ,NoMRP as [No MRP] ,Morning as [Morning] ,PROD_ITEM_CATEGORY_CODE as [Prod Item Category Code] ,Rate as [Rate] ,Active as [Active] ,Is_Slice as SliceType,AlternativeItem as [Alternative Item] ,ItemSpecification as [Item Specification] ,Item_Category_Struct_Code as [Item Category Structure Code] ,SubItemType as [item Sub Type] ,Is_Serial_Item as [Is Serial Item] ,Serial_Counter as [Serial Counter] ,WARRANTY_CODE as [Warranty Code] ,Is_Pick_Auto_SrNo as [Is Pick Auto Srno] ,Asset_Life as [Asset Life] ,Warranty_Period as [Warranty Period]  From tspl_item_master "
        qry += " left outer join (" + BalQry + ")BalQry on BalQry.icode=tspl_item_master.item_code"
        Icode = clsCommon.ShowSelectForm("ITMFND", qry, "Code", whrCls, strCurrCode, "Code", isButtonClicked)

        Return Icode
    End Function


End Class