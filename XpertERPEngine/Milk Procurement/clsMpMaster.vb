Imports common
Imports System.Data.SqlClient
Public Class clsMpMaster
    Public MP_CODE_VLC_UPLOADER As String = String.Empty
    Public MP_Code As String = String.Empty
    Public MP_Name As String = String.Empty
    Public MCC_Code As String = String.Empty
    Public Villege_Code As String = String.Empty
    Public Father_Name As String = String.Empty
    Public Add1 As String = String.Empty
    Public Add2 As String = String.Empty
    Public Zila As String = String.Empty
    Public Tehsil As String = String.Empty
    Public City_code As String = String.Empty
    Public State_Code As String = String.Empty
    Public Country_code As String = String.Empty
    Public Pin_code As String = String.Empty
    Public Telphone As String = String.Empty
    Public Email As String = String.Empty
    Public Fax As String = String.Empty
    Public Jan_Aadhar_No As String = String.Empty
    Public THIRD_PARTY_CODE As String = String.Empty
    Public Third_Party_Source As String = String.Empty
    Public DOB As Date = Nothing
    Public Education As String = String.Empty
    Public Land_Holding As Double = 0
    Public No_Of_Buffaloes As Integer = 0
    Public No_Of_Cows As Integer = 0
    Public No_Of_breedable_milk_animal As Integer = 0
    Public Milk_production As Double = 0
    Public Milk_Home_consumption As Double = 0
    Public Milk_For_sale As Double = 0
    Public PayeeName As String = Nothing
    Public BankName As String = Nothing
    Public BankBranch As String = Nothing
    Public BankCityCode As String = Nothing
    Public BankStateCode As String = Nothing
    Public IFCICode As String = Nothing
    Public AccountNO As String = Nothing
    Public Created_By As String = Nothing
    Public Created_Date As String = Nothing
    Public Modified_By As String = Nothing
    Public Modified_Date As String = Nothing
    Public Comp_Code As String = Nothing
    Public isNewEntry As Boolean = False
    Public arrBuffaloesDetail As List(Of clsBuffaloesDetails) = Nothing
    Public arrCowDetail As List(Of clsCowDetails) = Nothing
    Public arrAnimalDetail As List(Of clsAnimalDetails) = Nothing
    Public No_Of_Animal As Double = 0
    Public No_Of_Children_member As Double = 0
    Public No_Of_Adult_member As Double = 0
    Public No_Of_Total_Dependent_member As Double = 0
    Public Account_Type As String = Nothing
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing
    Public Form_Id As String = String.Empty
    Public Gender As String = Nothing
    Public MaritalStatus As String = Nothing
    Public IsVSP As Integer = 0
    Public InActive As Integer = 0
    Public TypeOfFormer As String = Nothing
    Public Milk_Availability_In_Lean_Season As Double = 0
    Public Milk_Availability_In_Flush_Season As Double = 0
    Public MpPicture As Byte() = Nothing
    Public Cust_Account As String = ""
    Public Acct_Set_Code As String = ""
    Public Cust_Account_Desc As String = ""
    Public Vendor_Acc_Code_Desc As String = ""
    Public TOLERANCE As Decimal = 0
    Public Incentive_Account As String = Nothing
    Public MP_Name_Hindi As String = Nothing
    Public ArrMPIncentiveMapping As ArrayList
    Public DISTRICT_Code As String = Nothing
    Public BLOCK_CODE As String = Nothing
    Public Zone_Code As String = Nothing
    Public CAST_CATEGORY_CODE As String = Nothing
    Public REVENUE_VILLAGE_CODE As String = Nothing
    Public GRAMPANCHAYAT_CODE As String = Nothing
    Public PANCHAYAT_SAMITI_CODE As String = Nothing
    Public VIDHAN_SABHA_CODE As String = Nothing
    Public Shared Function getMPNameOnMPCodeForVLCUplader(ByVal mPCode As String, ByVal Vlc_Code As String, ByVal trans As SqlTransaction) As String
        Dim mPName As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select mp_name from TSPL_MP_MASTER where MP_Code_vlc_uploader='" & mPCode & "' and vlc_Code='" & Vlc_Code & "' ", trans))
        Return mPName
    End Function
    Public Shared Function getMP_Code(ByVal MPCode As String, ByVal Vlc_Code As String, ByVal trans As SqlTransaction) As String
        Dim strCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MP_Code from TSPL_MP_MASTER where MP_Code='" & MPCode & "'  and vlc_Code='" & Vlc_Code & "'  ", trans))
        Return strCode
    End Function
    '----------------Code For Get Finder--------------------------------------------------------------------'
    Public Shared Function getFinder(ByVal whrcls As String, ByVal curcode As String, ByVal isButtonClicked As Boolean) As String
        Dim str As String = ""
        Try
            Dim qry As String = " select tspl_mp_master.MP_Code as [Code] ,tspl_mp_master.MP_Name as [MP Name] ,tspl_mp_master.VLC_Code as [DCS Code],VLCH.VLC_Name as [DCS Name],VLCH.VLC_Code_VLC_Uploader as [DCS Uploader Code],tspl_mp_master.Village_Code as [Village Code] ,tspl_mp_master.Father_Name as [Father Name] ,tspl_mp_master.Add1 as [Address1] ,tspl_mp_master.Add2 as [Address2] ,tspl_mp_master.Zila as [Zila] ,tspl_mp_master.Tehsil as [Tehsil] ,tspl_mp_master.City_code as [City Code] ,tspl_mp_master.State_Code as [State Code] ,tspl_mp_master.Country_code as [Country Code] ,tspl_mp_master.Pin_code as [Pin Code] ,tspl_mp_master.Telphone as [Telphone] ,tspl_mp_master.Email as [Email] ,tspl_mp_master.Fax as [Fax],tspl_mp_master.Jan_Aadhar_No as [Jan Aadhar No] ,tspl_mp_master.THIRD_PARTY_CODE as [THIRD PARTY CODE] ,tspl_mp_master.Third_Party_Source as [THIRD PARTY SOURCE],tspl_mp_master.DOB as [Date Of Birth] ,tspl_mp_master.Education as [Education] ,tspl_mp_master.Land_Holding as [Land Holding] ,tspl_mp_master.No_Of_Buffaloes as [No Of Buffaloes] ,tspl_mp_master.No_Of_Cows as [No Of Cows] ,tspl_mp_master.No_Of_breedable_milk_animal as [No Of Breedable Milk Animal] ,tspl_mp_master.Milk_production as [Total Milk Production] ,tspl_mp_master.Milk_Home_consumption as [Total Milk Home Consumption] ,tspl_mp_master.Milk_For_sale as [Remaining Milk For Sale] ,tspl_mp_master.PayeeName as [Payee Name] ,tspl_mp_master.BankName as [Bank Name] ,tspl_mp_master.BankBranch as [Bank Branch] ,tspl_mp_master.BankCityCode as [Bank City Code] ,tspl_mp_master.BankStateCode as [Bank State Code] ,tspl_mp_master.IFCICode as [IFCI Code] ,convert(varchar,tspl_mp_master.AccountNO) as [Account No] ,tspl_mp_master.Created_By as [Created By] ,tspl_mp_master.Created_Date as [Created Date] ,tspl_mp_master.Modified_By as [Modified By] ,tspl_mp_master.Modified_Date as [Modified Date] ,tspl_mp_master.Comp_Code as [Company Code],tspl_mp_master.Mp_code_Vlc_uploader as [MP Code VLC Uploder],VLCH.MCC as [MCC Code]  From tspl_mp_master left join TSPL_VLC_MASTER_HEAD VLCH on tspl_mp_master.VLC_Code=VLCH.VLC_Code"
            str = clsCommon.ShowSelectForm("FNDMPMST", qry, "Code", whrcls, curcode, "Code", isButtonClicked)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return str
    End Function
    Public Shared Function ExportDataTable(ByVal strMPCode As ArrayList, ByVal frmMe As RadForm, ByVal exportSheet As String) As Boolean
        Try
            Dim whrQry As String = Nothing
            If clsCommon.myLen(strMPCode) > 0 AndAlso strMPCode IsNot Nothing Then
                whrQry = "where tspl_mp_Master.Mp_code In (" + clsCommon.GetMulcallString(strMPCode) + ")"
            End If
            Dim strQry As String = Nothing
            If exportSheet.Contains("BlankSheet") Then
                strQry = "select  ''  As 'MP Code','' As 'MP Name' ,'' as 'Payee Name' "
            Else
                strQry = "select  tspl_mp_Master.mp_code  As 'MP Code',tspl_mp_Master.mp_Name As 'MP Name',tspl_mp_Master.PayeeName as 'Payee Name' from tspl_mp_Master
                        " + whrQry + ""
            End If

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt.Rows.Count > 0 Then
                transportSql.ExporttoExcel(dt, frmMe)
                dt = Nothing
            Else
                Throw New Exception("No data found")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    '----------------End of Code For Get Finder--------------------------------------------------------------'
    Public Shared Function deleteData(ByVal strcode As String, ByVal progcode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isdeleted As Boolean = True
            'isdeleted = isdeleted And clsCowDetails.deleteData(progcode, strcode, trans)
            'isdeleted = isdeleted And clsBuffaloesDetails.deleteData(progcode, strcode, trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strcode, "tspl_mp_master", "MP_Code", trans)

            isdeleted = isdeleted And clsAnimalDetails.deleteData(progcode, strcode, trans)
            Dim qry As String = "delete from TSPL_MP_INCENTIVE where  mp_code='" & strcode & "'"
            isdeleted = isdeleted And clsDBFuncationality.ExecuteNonQuery(qry, trans)
            qry = "delete from tspl_mp_master where  mp_code='" & strcode & "'"
            isdeleted = isdeleted And clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Return isdeleted
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try

    End Function

    
    Public Shared Function loadData(ByVal strcode As String, ByVal navtype As NavigatorType, ByVal progcode As String) As clsMpMaster
        Dim obj As New clsMpMaster
        Try
            'obj.arrCowDetail = New List(Of clsCowDetails)
            'obj.arrBuffaloesDetail = New List(Of clsBuffaloesDetails)
            obj.arrAnimalDetail = New List(Of clsAnimalDetails)
            Dim qst As String = " select *   From tspl_mp_master   where 1=1"
            Select Case navtype
                Case NavigatorType.Current
                    qst += " and tspl_mp_master.Mp_Code in ('" + strcode + "')"
                Case NavigatorType.Next
                    qst += " and tspl_mp_master.Mp_Code in (select min(Mp_Code ) from tspl_mp_master where Mp_Code  >'" + strcode + "')"
                Case NavigatorType.First
                    qst += " and tspl_mp_master.Mp_Code in (select MIN(Mp_Code ) from tspl_mp_master)"
                Case NavigatorType.Last
                    qst += " and tspl_mp_master.Mp_Code in (select Max(Mp_Code ) from tspl_mp_master)"
                Case NavigatorType.Previous
                    qst += " and tspl_mp_master.Mp_Code in (select Max(Mp_Code ) from tspl_mp_master where Mp_Code  <'" + strcode + "')"
            End Select
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qst)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                obj.MCC_Code = clsCommon.myCstr(dt.Rows(0)("vlc_Code"))
                obj.MP_Code = clsCommon.myCstr(dt.Rows(0)("MP_Code"))
                obj.MP_Name = clsCommon.myCstr(dt.Rows(0)("MP_NAME"))
                obj.Father_Name = clsCommon.myCstr(dt.Rows(0)("Father_NAME"))
                obj.Villege_Code = clsCommon.myCstr(dt.Rows(0)("Village_Code"))
                obj.Add1 = clsCommon.myCstr(dt.Rows(0)("Add1"))
                obj.Add2 = clsCommon.myCstr(dt.Rows(0)("Add2"))
                obj.Zila = clsCommon.myCstr(dt.Rows(0)("Zila"))
                obj.Tehsil = clsCommon.myCstr(dt.Rows(0)("Tehsil"))
                obj.City_code = clsCommon.myCstr(dt.Rows(0)("City_code"))
                obj.State_Code = clsCommon.myCstr(dt.Rows(0)("State_Code"))
                obj.Country_code = clsCommon.myCstr(dt.Rows(0)("Country_code"))
                obj.Pin_code = clsCommon.myCstr(dt.Rows(0)("Pin_code"))
                obj.Telphone = clsCommon.myCstr(dt.Rows(0)("Telphone"))
                obj.Email = clsCommon.myCstr(dt.Rows(0)("Email"))
                obj.Fax = clsCommon.myCstr(dt.Rows(0)("Fax"))
                obj.Jan_Aadhar_No = clsCommon.myCstr(dt.Rows(0)("Jan_Aadhar_No"))
                obj.THIRD_PARTY_CODE = clsCommon.myCstr(dt.Rows(0)("THIRD_PARTY_CODE"))
                obj.Third_Party_Source = clsCommon.myCstr(dt.Rows(0)("Third_Party_Source"))
                If dt.Rows(0)("DOB") IsNot DBNull.Value Then
                    obj.DOB = clsCommon.myCDate(dt.Rows(0)("DOB"))
                End If

                obj.MP_CODE_VLC_UPLOADER = clsCommon.myCstr(dt.Rows(0)("MP_CODE_VLC_UPLOADER"))
                obj.Education = clsCommon.myCstr(dt.Rows(0)("Education"))
                obj.Land_Holding = clsCommon.myCdbl(dt.Rows(0)("Land_Holding"))
                'obj.No_Of_Buffaloes = clsCommon.myCdbl(dt.Rows(0)("No_Of_Buffaloes"))
                'obj.No_Of_Cows = clsCommon.myCdbl(dt.Rows(0)("No_Of_Cows"))
                'obj.No_Of_Animal = clsCommon.myCdbl(dt.Rows(0)("No_Of_Animal"))
                obj.No_Of_breedable_milk_animal = clsCommon.myCdbl(dt.Rows(0)("No_Of_breedable_milk_animal"))
                obj.Milk_production = clsCommon.myCdbl(dt.Rows(0)("Milk_production"))
                obj.Milk_Home_consumption = clsCommon.myCdbl(dt.Rows(0)("Milk_Home_consumption"))
                obj.Milk_For_sale = clsCommon.myCdbl(dt.Rows(0)("Milk_For_sale"))
                obj.PayeeName = clsCommon.myCstr(dt.Rows(0)("PayeeName"))
                obj.BankName = clsCommon.myCstr(dt.Rows(0)("BankName"))
                obj.BankBranch = clsCommon.myCstr(dt.Rows(0)("BankBranch"))
                obj.BankCityCode = clsCommon.myCstr(dt.Rows(0)("BankCityCode"))
                obj.BankStateCode = clsCommon.myCstr(dt.Rows(0)("BankStateCode"))
                obj.IFCICode = clsCommon.myCstr(dt.Rows(0)("IFCICode"))
                obj.No_Of_Children_member = clsCommon.myCdbl(dt.Rows(0)("No_Of_Children_member"))
                obj.No_Of_Adult_member = clsCommon.myCdbl(dt.Rows(0)("No_Of_Adult_member"))
                obj.No_Of_Total_Dependent_member = clsCommon.myCdbl(dt.Rows(0)("No_Of_Total_Dependent_member"))
                obj.Account_Type = clsCommon.myCstr(dt.Rows(0)("Account_Type"))
                obj.AccountNO = clsCommon.myCstr(dt.Rows(0)("AccountNO"))
                '===============================================
                obj.IsVSP = clsCommon.myCstr(dt.Rows(0)("Is_VSP"))
                obj.InActive = clsCommon.myCstr(dt.Rows(0)("Active"))
                obj.TypeOfFormer = clsCommon.myCstr(dt.Rows(0)("Type_Of_Former"))
                obj.Gender = clsCommon.myCstr(dt.Rows(0)("Gender"))
                obj.MaritalStatus = clsCommon.myCstr(dt.Rows(0)("Marital_Status"))
                obj.Milk_Availability_In_Lean_Season = clsCommon.myCdbl(dt.Rows(0)("Milk_Avalb_Lean_Season"))
                obj.Milk_Availability_In_Flush_Season = clsCommon.myCdbl(dt.Rows(0)("Milk_Avalb_flush_Season"))

                obj.Cust_Account = clsCommon.myCstr(dt.Rows(0)("Cust_Account"))
                obj.Acct_Set_Code = clsCommon.myCstr(dt.Rows(0)("Acct_Set_Code"))
                obj.Incentive_Account = clsCommon.myCstr(dt.Rows(0)("Incentive_Account"))
                If String.IsNullOrEmpty(clsCommon.myCstr(dt.Rows(0)("TOLERANCE"))) = False Then
                    obj.TOLERANCE = clsCommon.myCdbl(dt.Rows(0)("TOLERANCE"))
                Else
                    obj.TOLERANCE = Nothing
                End If

                obj.Cust_Account_Desc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Acct_Desc from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" & obj.Cust_Account & "'"))
                obj.Vendor_Acc_Code_Desc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Acct_Set_Desc from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code='" & obj.Acct_Set_Code & "'"))
                obj.MP_Name_Hindi = clsCommon.myCstr(dt.Rows(0)("MP_Name_Hindi"))
                obj.CAST_CATEGORY_CODE = clsCommon.myCstr(dt.Rows(0)("CAST_CATEGORY_CODE"))
                obj.DISTRICT_Code = clsCommon.myCstr(dt.Rows(0)("DISTRICT_Code"))
                obj.BLOCK_CODE = clsCommon.myCstr(dt.Rows(0)("BLOCK_CODE"))
                obj.Zone_Code = clsCommon.myCstr(dt.Rows(0)("Zone_Code"))
                obj.REVENUE_VILLAGE_CODE = clsCommon.myCstr(dt.Rows(0)("REVENUE_VILLAGE_CODE"))
                obj.GRAMPANCHAYAT_CODE = clsCommon.myCstr(dt.Rows(0)("GRAMPANCHAYAT_CODE"))
                obj.PANCHAYAT_SAMITI_CODE = clsCommon.myCstr(dt.Rows(0)("PANCHAYAT_SAMITI_CODE"))
                obj.VIDHAN_SABHA_CODE = clsCommon.myCstr(dt.Rows(0)("VIDHAN_SABHA_CODE"))
                '===================================================
                'obj.arrBuffaloesDetail = clsBuffaloesDetails.LoadData(progcode, obj.MP_Code)
                'obj.arrCowDetail = clsCowDetails.LoadData(progcode, obj.MP_Code)
                obj.arrAnimalDetail = clsAnimalDetails.LoadData(progcode, obj.MP_Code)
                obj.ArrMPIncentiveMapping = clsMPIncentiveMapping.GetData(obj.MP_Code)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return obj
    End Function
    Public Shared Function SaveData(ByVal obj As clsMpMaster, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = ""
            If obj.InActive = 0 Then
                qry = "Delete from TSPL_MP_INCENTIVE where MP_CODE ='" + obj.MP_Code + "' "
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            If obj.InActive = 0 Then
                qry = "select AccountNO from TSPL_MP_MASTER where MP_Code not in ('" + obj.MP_Code + "') and Active=0 and AccountNO='" + obj.AccountNO + "' "
                Dim dtACNo As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dtACNo IsNot Nothing AndAlso dtACNo.Rows.Count > 0 Then
                    Throw New Exception("Account number [" + obj.AccountNO + "] is already in use.")
                End If
            End If

            Dim issaved As Boolean = True
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "MP_Code", obj.MP_Code)
            clsCommon.AddColumnsForChange(coll, "VLC_Code", obj.MCC_Code, True)
            clsCommon.AddColumnsForChange(coll, "MP_NAME", obj.MP_Name)
            clsCommon.AddColumnsForChange(coll, "Village_Code", obj.Villege_Code)
            clsCommon.AddColumnsForChange(coll, "Father_Name", obj.Father_Name)
            clsCommon.AddColumnsForChange(coll, "Add1", obj.Add1)
            clsCommon.AddColumnsForChange(coll, "Add2", obj.Add2)
            clsCommon.AddColumnsForChange(coll, "Zila", obj.Zila)
            clsCommon.AddColumnsForChange(coll, "Tehsil", obj.Tehsil)
            clsCommon.AddColumnsForChange(coll, "City_code", obj.City_code, True)
            clsCommon.AddColumnsForChange(coll, "State_Code", obj.State_Code, True)
            clsCommon.AddColumnsForChange(coll, "Pin_Code", obj.Pin_code)
            clsCommon.AddColumnsForChange(coll, "Country_code", obj.Country_code, True)
            clsCommon.AddColumnsForChange(coll, "Telphone", obj.Telphone)
            clsCommon.AddColumnsForChange(coll, "Email", obj.Email)
            clsCommon.AddColumnsForChange(coll, "Fax", obj.Fax)
            clsCommon.AddColumnsForChange(coll, "Jan_Aadhar_No", obj.Jan_Aadhar_No)
            clsCommon.AddColumnsForChange(coll, "THIRD_PARTY_CODE", obj.THIRD_PARTY_CODE)
            clsCommon.AddColumnsForChange(coll, "Third_Party_Source", obj.Third_Party_Source)
            clsCommon.AddColumnsForChange(coll, "DOB", clsCommon.GetPrintDate(obj.DOB, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "MP_CODE_VLC_UPLOADER", clsCommon.myCstr(obj.MP_CODE_VLC_UPLOADER))
            clsCommon.AddColumnsForChange(coll, "Education", obj.Education)
            clsCommon.AddColumnsForChange(coll, "Land_Holding", obj.Land_Holding)
            'clsCommon.AddColumnsForChange(coll, "No_Of_Buffaloes", obj.No_Of_Buffaloes)
            'clsCommon.AddColumnsForChange(coll, "No_Of_Cows", obj.No_Of_Cows)
            'clsCommon.AddColumnsForChange(coll, "No_Of_Animal", obj.No_Of_Animal)
            clsCommon.AddColumnsForChange(coll, "No_Of_Children_member", obj.No_Of_Children_member)
            clsCommon.AddColumnsForChange(coll, "No_Of_Adult_member", obj.No_Of_Adult_member)
            clsCommon.AddColumnsForChange(coll, "No_Of_Total_Dependent_member", obj.No_Of_Total_Dependent_member)
            clsCommon.AddColumnsForChange(coll, "No_Of_breedable_milk_animal", obj.No_Of_breedable_milk_animal)
            clsCommon.AddColumnsForChange(coll, "Milk_production", obj.Milk_production)
            clsCommon.AddColumnsForChange(coll, "Milk_Home_consumption", obj.Milk_Home_consumption)
            clsCommon.AddColumnsForChange(coll, "Milk_For_sale", obj.Milk_For_sale)
            clsCommon.AddColumnsForChange(coll, "PayeeName", clsCommon.myCleanAlphaOnly(obj.PayeeName))
            clsCommon.AddColumnsForChange(coll, "BankName", obj.BankName)
            clsCommon.AddColumnsForChange(coll, "BankBranch", obj.BankBranch)
            clsCommon.AddColumnsForChange(coll, "BankCityCode", obj.BankCityCode)
            clsCommon.AddColumnsForChange(coll, "BankStateCode", obj.BankStateCode)
            clsCommon.AddColumnsForChange(coll, "IFCICode", obj.IFCICode)
            clsCommon.AddColumnsForChange(coll, "AccountNO", obj.AccountNO)
            clsCommon.AddColumnsForChange(coll, "Account_Type", obj.Account_Type)
            clsCommon.AddColumnsForChange(coll, "Modified_By", obj.Modified_By)
            clsCommon.AddColumnsForChange(coll, "Modified_Date", obj.Modified_Date)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", obj.Comp_Code)
            clsCommon.AddColumnsForChange(coll, "Is_VSP", obj.IsVSP)
            clsCommon.AddColumnsForChange(coll, "Active", obj.InActive)
            clsCommon.AddColumnsForChange(coll, "Type_Of_Former", obj.TypeOfFormer)
            clsCommon.AddColumnsForChange(coll, "Gender", obj.Gender)
            clsCommon.AddColumnsForChange(coll, "Marital_Status", obj.MaritalStatus)
            clsCommon.AddColumnsForChange(coll, "MP_Picture", obj.MpPicture)
            clsCommon.AddColumnsForChange(coll, "Milk_Avalb_Lean_Season", obj.Milk_Availability_In_Lean_Season)
            clsCommon.AddColumnsForChange(coll, "Milk_Avalb_flush_Season", obj.Milk_Availability_In_Flush_Season)

            clsCommon.AddColumnsForChange(coll, "Cust_Account", obj.Cust_Account, True)
            clsCommon.AddColumnsForChange(coll, "Acct_Set_Code", obj.Acct_Set_Code, True)
            clsCommon.AddColumnsForChange(coll, "TOLERANCE", obj.TOLERANCE, True)
            clsCommon.AddColumnsForChange(coll, "Incentive_Account", obj.Incentive_Account, True)
            clsCommon.AddColumnsForChange(coll, "MP_Name_Hindi", obj.MP_Name_Hindi, True, True)
            clsCommon.AddColumnsForChange(coll, "CAST_CATEGORY_CODE", obj.CAST_CATEGORY_CODE, True) 'IIf(clsCommon.myLen(obj.CAST_CATEGORY_CODE) > 0, obj.CAST_CATEGORY_CODE, Nothing)
            clsCommon.AddColumnsForChange(coll, "DISTRICT_Code", obj.DISTRICT_Code, True)
            clsCommon.AddColumnsForChange(coll, "BLOCK_CODE", obj.BLOCK_CODE, True)
            clsCommon.AddColumnsForChange(coll, "Zone_Code", obj.Zone_Code, True)
            clsCommon.AddColumnsForChange(coll, "REVENUE_VILLAGE_CODE", obj.REVENUE_VILLAGE_CODE, True)
            clsCommon.AddColumnsForChange(coll, "GRAMPANCHAYAT_CODE", obj.GRAMPANCHAYAT_CODE, True)
            clsCommon.AddColumnsForChange(coll, "PANCHAYAT_SAMITI_CODE", obj.PANCHAYAT_SAMITI_CODE, True)
            clsCommon.AddColumnsForChange(coll, "VIDHAN_SABHA_CODE", obj.VIDHAN_SABHA_CODE, True)

            If obj.isNewEntry Then
                clsCommon.AddColumnsForChange(coll, "Created_By", obj.Created_By)
                clsCommon.AddColumnsForChange(coll, "Created_Date", obj.Created_Date)
                issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "tspl_mp_master", OMInsertOrUpdate.Insert, "", trans)
            Else
                'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.MP_Code, "tspl_mp_master", "MP_Code", trans)

                issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "tspl_mp_master", OMInsertOrUpdate.Update, "tspl_mp_master.mp_code='" + obj.MP_Code + "'", trans)
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.MP_Code, "tspl_mp_master", "MP_Code", trans)

            'qry = "select 1 from TSPL_MP_MASTER where MP_Code not in ('" & obj.MP_Code & " ') and AccountNO='" & obj.AccountNO & "'"
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '    Throw New Exception("Duplicate account no [" + obj.AccountNO + "]")
            'End If
            qry = ""

            'issaved = issaved And clsBuffaloesDetails.SaveData(obj.arrBuffaloesDetail, trans)
            'issaved = issaved And clsCowDetails.SaveData(obj.arrCowDetail, trans)
            If obj.arrAnimalDetail Is Nothing OrElse obj.arrAnimalDetail.Count = 0 Then
                qry = "delete from tspl_Animal_Details where prog_code ='" + obj.Form_Id + "' and trans_code='" + obj.MP_Code + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
            End If
            issaved = issaved And clsAnimalDetails.SaveData(obj.arrAnimalDetail, trans)
            issaved = issaved AndAlso clsCustomFieldValues.SaveData(obj.Form_Id, obj.MP_Code, obj.arrCustomFields, trans)
            issaved = issaved AndAlso clsMPIncentiveMapping.SaveData(obj.MP_Code, obj.Modified_Date, obj.ArrMPIncentiveMapping, trans)


            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function



    Public Shared Function GetLedgerSummaryDt(ByVal From_Date As Date, ByVal To_Date As Date, ByVal lstCNF As ArrayList, ByVal lstDistr As ArrayList, ByVal trans As SqlTransaction, ByVal lstMCC As ArrayList) As DataTable
        Dim qry As String = GetLedgerSummaryQry(From_Date, To_Date, lstCNF, lstDistr, lstMCC)
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Return dt
    End Function
    Public Shared Function GetLedgerDetailDt(ByVal From_Date As Date, ByVal To_Date As Date, ByVal lstCNF As ArrayList, ByVal lstDistr As ArrayList, ByVal trans As SqlTransaction, ByVal lstMCC As ArrayList) As DataTable
        Dim qry As String = GetLedgerDetailQry(From_Date, To_Date, lstCNF, lstDistr, lstMCC)
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Return dt
    End Function
    Public Shared Function GetLedgerSummaryQry(ByVal From_Date As Date, ByVal To_Date As Date, ByVal lstVSP As ArrayList, ByVal lstMP As ArrayList, ByVal lstMCC As ArrayList) As String
        'Dim BaseQry As String = ""
        'Dim Qry As String = ""
        'BaseQry = GetLedgerBaseQry(From_Date, To_Date, lstCNF, lstDistr, False)
        Dim BaseQry As String = ""
        Dim BaseQryOP As String = ""
        Dim Qry As String = ""
        BaseQryOP = GetLedgerBaseQry(From_Date, To_Date, lstVSP, lstMP, True)
        BaseQry = GetLedgerBaseQry(From_Date, To_Date, lstVSP, lstMP, False)

        Qry = " select FARMER_CODE,Farmer_Name,VSP_CODE,VSP_Name,sum(Opening) as Opening,sum(Debit) as Debit,sum(Credit) as Credit,cast( sum(Opening+Debit-Credit) as decimal(18,2)) as Closing from (" &
              " select Opening.FARMER_CODE,Opening.FARMER_NAME,Opening.VSP_CODE,Opening.VSP_NAME,Null as Trans_Type,null as Doc_No," &
              " Null as Doc_Date,'Opening' as Remarks,Cast ( sum(Opening.Debit-Opening.Credit) as decimal(18,2)) as Opening,0 as Debit,0 as Credit,Cast ( sum(Opening.Debit-Opening.Credit) as decimal(18,2)) as Balance " &
              " from (" & BaseQryOP & ") as Opening left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=Opening.VSP_CODE " &
              " where cast(Opening.Farmer_Invoice_Date as date) <'" & clsCommon.GetPrintDate(From_Date, "dd/MMM/yyyy") & "' "
        If lstMCC IsNot Nothing AndAlso lstMCC.Count > 0 Then
            Qry += "  and TSPL_VLC_MASTER_HEAD.MCC in (" + clsCommon.GetMulcallString(lstMCC) + ") "
        End If
        Qry += " group by Opening.FARMER_CODE,Opening.FARMER_NAME,Opening.VSP_CODE,Opening.VSP_NAME " &
              " union all " &
              " select Trans.FARMER_CODE,Trans.Farmer_Name,Trans.VSP_CODE,Trans.VSP_Name,Trans.Trans_Type,Trans.Farmer_Invoice_No as Doc_No, " &
              " Trans.Farmer_Invoice_Date as Doc_Date,Trans.Remarks,0 as Opening,Trans.Debit,Trans.Credit,(Trans.Debit-Trans.Credit) as Balance " &
              " from (" & BaseQry & ") as Trans left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=Trans.VSP_CODE " &
              " where cast(Trans.Farmer_Invoice_Date as date) >= '" & clsCommon.GetPrintDate(From_Date, "dd/MMM/yyyy") & "' and cast(Trans.Farmer_Invoice_Date as date)<= '" & clsCommon.GetPrintDate(To_Date, "dd/MMM/yyyy") & "' "
        If lstMCC IsNot Nothing AndAlso lstMCC.Count > 0 Then
            Qry += "  and TSPL_VLC_MASTER_HEAD.MCC in (" + clsCommon.GetMulcallString(lstMCC) + ") "
        End If
        Qry += " ) as Final where 2=2 group by FARMER_CODE,FARMER_NAME,VSP_CODE,VSP_NAME"

        Return Qry
    End Function
    Public Shared Function GetLedgerDetailQry(ByVal From_Date As Date, ByVal To_Date As Date, ByVal lstVSP As ArrayList, ByVal lstMP As ArrayList, ByVal lstMCC As ArrayList) As String
        Dim BaseQry As String = ""
        Dim BaseQryOP As String = ""
        Dim Qry As String = ""
        BaseQryOP = GetLedgerBaseQry(From_Date, To_Date, lstVSP, lstMP, True)
        BaseQry = GetLedgerBaseQry(From_Date, To_Date, lstVSP, lstMP, False)
        ' Ticket No : ERO/07/05/19-000587 By Prabhakar
        Qry = " select *,row_number() over (partition by FARMER_CODE order by FARMER_CODE,Doc_Date) as Tr_id from (" &
              " select Opening.FARMER_CODE,FARMER_NAME,Opening.VSP_CODE,VSP_Name,Null as Trans_Type,null as Doc_No," &
              " Null as Doc_Date,'Opening' as Remarks,'' as PaymentProcessNo,'' as APAdjustmentNo ,cast(sum(Opening.Debit-Opening.Credit) as decimal(18,2)) as Debit,0 as Credit,cast( sum(Opening.Debit-Opening.Credit)as decimal(18,2)) as Balance " &
              " from (" & BaseQryOP & ") as Opening left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=Opening.VSP_CODE  " &
              " where cast(Opening.Farmer_Invoice_Date as date) <'" & clsCommon.GetPrintDate(From_Date, "dd/MMM/yyyy") & "' "
        If lstMCC IsNot Nothing AndAlso lstMCC.Count > 0 Then
            Qry += "  and TSPL_VLC_MASTER_HEAD.MCC in (" + clsCommon.GetMulcallString(lstMCC) + ") "
        End If
        Qry += " group by Opening.FARMER_CODE,Opening.FARMER_NAME,Opening.VSP_CODE,Opening.VSP_NAME " &
              " union all " &
              " select Trans.FARMER_CODE,FARMER_NAME,Trans.VSP_CODE,VSP_Name,Trans.Trans_Type,Trans.Farmer_Invoice_No as Doc_No, " &
              " Trans.Farmer_Invoice_Date as Doc_Date,Trans.Remarks,Trans.PaymentProcessNo,Trans.APAdjustmentNo,Trans.Debit,Trans.Credit,(Trans.Debit-Trans.Credit) as Balance " &
              " from (" & BaseQry & ") as Trans left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=Trans.VSP_CODE  " &
              " where cast(Trans.Farmer_Invoice_Date as date)>= '" & clsCommon.GetPrintDate(From_Date, "dd/MMM/yyyy") & "' and cast(Trans.Farmer_Invoice_Date as date)<= '" & clsCommon.GetPrintDate(To_Date, "dd/MMM/yyyy") & "' "
        If lstMCC IsNot Nothing AndAlso lstMCC.Count > 0 Then
            Qry += "  and TSPL_VLC_MASTER_HEAD.MCC in (" + clsCommon.GetMulcallString(lstMCC) + ") "
        End If
        Qry += " ) as Final where 2=2 "
        Qry = "select *,round(sum(outermost.Balance) over (partition by outermost.FARMER_CODE order by outermost.FARMER_CODE,outermost.Tr_Id),2) as RunningBalance from (" & Qry & ") as Outermost"
        Return Qry
    End Function
    Public Shared Function GetName(ByVal MP_Code As String, ByVal trans As SqlTransaction) As String
        Dim qry As String = "select MP_Name from tspl_mp_master where MP_Code='" & MP_Code & "'"
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
    End Function
    Public Shared Function GetLedgerBaseQry(ByVal From_Date As Date, ByVal To_Date As Date, ByVal lstVSP As ArrayList, ByVal lstMP As ArrayList, ByVal Is_Opening As Boolean) As String
        Dim Qry As String = ""
        Dim CondDisp As String = ""
        Dim CondRec As String = ""
        Dim condVCGL As String = ""
        If Not lstVSP Is Nothing AndAlso lstVSP.Count > 0 Then
            CondDisp = CondDisp & " and MPD.VSP_CODE in (" & clsCommon.GetMulcallString(lstVSP) & ")"
            CondRec = CondRec & " and TSPL_VENDOR_MASTER.VENDOR_CODE in (" & clsCommon.GetMulcallString(lstVSP) & ")"
            condVCGL = condVCGL & " and TSPL_VENDOR_MASTER.VENDOR_CODE in (" & clsCommon.GetMulcallString(lstVSP) & ")"
        End If
        If Not lstMP Is Nothing AndAlso lstMP.Count > 0 Then
            CondDisp = CondDisp & " and MPD.FARMER_CODE in (" & clsCommon.GetMulcallString(lstMP) & ")"
            CondRec = CondRec & "and PAY.FARMER_CODE in (" & clsCommon.GetMulcallString(lstMP) & ")"
            condVCGL = condVCGL & " and TSPL_MP_MASTER.MP_Code in (" & clsCommon.GetMulcallString(lstMP) & ")"
        End If
        Dim qryMP As String = "" 'clsMilkPurchaseInvoiceMCC.GetBaseQueryWithMP(From_Date,To_Date,
        If Is_Opening Then
            qryMP = clsMilkPurchaseInvoiceMCC.GetBaseQueryWithMP(Nothing, From_Date, "", lstVSP)
        Else
            qryMP = clsMilkPurchaseInvoiceMCC.GetBaseQueryWithMP(From_Date, To_Date, "", lstVSP)
        End If
        qryMP = " select Coll.Type as Trans_Type,(Billed.DOC_CODE + '/' + Coll.MP_CODE) AS Farmer_Invoice_No,cast(Billed.DOC_DATE as date) AS Farmer_Invoice_Date," &
                " Billed.VSP_CODE,VSP.VENDOR_NAME AS VSP_NAME,Coll.MP_CODE,Coll.MP_NAME AS Farmer_Name,'' as Remarks,0 as Debit, Coll.Amount as Credit " &
                " from (" & qryMP & ") Coll inner join (select MPH.DOC_CODE,coalesce(MPS.MANUAL_DOC_NO,MPS.PD_DOC_NO) AS Coll_Doc_Code,MPH.DOC_DATE,MPS.SHIFT,MPS.VSP_CODE,MPS.MCC_CODE,MPS.VLC_Code " &
                " from TSPL_MILK_PURCHASE_INVOICE_MP_COLLEC_SUMMARY MPS inner join TSPL_MILK_PURCHASE_INVOICE_HEAD MPH ON MPS.DOC_CODE=MPH.DOC_CODE) as Billed " &
                " on Coll.Doc_No=Billed.Coll_Doc_Code and Coll.VLC_CODE=Billed.VLC_CODE " &
                " left join TSPL_MP_PAY_PROCESS_DETAIL on TSPL_MP_PAY_PROCESS_DETAIL.Farmer_Invoice_No=Billed.DOC_CODE+'/'+TSPL_MP_PAY_PROCESS_DETAIL.Farmer_Code and TSPL_MP_PAY_PROCESS_DETAIL.Farmer_Code=Coll.MP_CODE and TSPL_MP_PAY_PROCESS_DETAIL.VSP_CODE=Billed.VSP_CODE " &
                " LEFT JOIN TSPL_MP_MASTER MP ON Coll.MP_CODE=MP.MP_CODE " &
                " LEFT JOIN TSPL_VENDOR_MASTER VSP ON VSP.VENDOR_CODE=Billed.VSP_CODE where TSPL_MP_PAY_PROCESS_DETAIL.Farmer_Invoice_No IS NULL and MP.MP_Farmer_Billing=1"
        If Not lstVSP Is Nothing AndAlso lstVSP.Count > 0 Then
            qryMP = qryMP & " and Billed.VSP_CODE in (" & clsCommon.GetMulcallString(lstVSP) & ")"
        End If
        If Not lstMP Is Nothing AndAlso lstMP.Count > 0 Then
            qryMP = qryMP & " and Coll.MP_CODE in (" & clsCommon.GetMulcallString(lstMP) & ")"
        End If
        qryMP = "SELECT Trans_Type,Farmer_Invoice_No,Farmer_Invoice_Date,VSP_CODE,VSP_NAME,MP_CODE,Farmer_Name,max(Remarks) as Remarks,sum(Debit) as Debit,sum(Credit) as Credit,'' as PaymentProcessNo  FROM (" & qryMP & ") Final group by Trans_Type,Farmer_Invoice_No,Farmer_Invoice_Date,VSP_CODE,VSP_NAME,MP_CODE,Farmer_Name"
        ''query change by Panch Raj on 01-may-2018 against ticket : KDI/30/04/18-000281
        Qry = " select 'Invoice' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE,TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_Name as Farmer_Name,MPH.PaymentDesc as Remarks,0 as Debit,round(MPD.Milk_Amount,2) as Credit,MPD.Doc_No as PaymentProcessNo,MPD.AP_Adjustment_No as APAdjustmentNo" &
              " from TSPL_MP_PAY_PROCESS_DETAIL MPD " &
              " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE left outer join TSPL_Payment_Adjustment_Header on TSPL_Payment_Adjustment_Header.Adjustment_No=MPD.AP_Adjustment_No  WHERE 2=2 and   TSPL_Payment_Adjustment_Header.Is_Post='Y'     " & CondDisp & "" &
              " UNION ALL " &
              " select 'Incentive' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE,TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,0 as Debit,round(MPD.Incentive_Amount,2) as Credit,MPD.Doc_No as PaymentProcessNo ,tspl_journal_Master.Voucher_No as APAdjustmentNo " &
              " from TSPL_MP_PAY_PROCESS_DETAIL MPD " &
              " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code    left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE left outer join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No and tspl_journal_Master.Source_Code='MP-IV' WHERE 2=2 and isnull(MPD.Incentive_Amount,0)>0 " & CondDisp & "" &
              " UNION ALL " &
              " select 'Dedution' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE, TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,round(MPD.Deduction_Amount,2) as Debit,0 as Credit,MPD.Doc_No as PaymentProcessNo,(case when TSPL_Payment_Adjustment_Header.Adjustment_No is not null then TSPL_Payment_Adjustment_Header.Adjustment_No else tspl_journal_Master.Voucher_No end) as  APAdjustmentNo " &
              " from TSPL_MP_PAY_PROCESS_DETAIL MPD " &
              " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code    left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE left outer join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No and tspl_journal_Master.Source_Code='MP-DE' left outer join TSPL_Payment_Adjustment_Header on TSPL_Payment_Adjustment_Header.Doc_No=MPD.AP_Invoice_No and TSPL_Payment_Adjustment_Header.Description = 'AP Return Adjustment Against Bulk Payment Process for extra amount to be paid by VSP' and TSPL_Payment_Adjustment_Header.Adjust_Type = 'R' WHERE 2=2 and isnull(MPD.Deduction_Amount,0)>0 " & CondDisp & "" &
            " UNION ALL " &
            " select 'Advance Recovery' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE, TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,round(MPD.Total_Advance_Amount_Recovery,2) as Debit,0 as Credit,MPD.Doc_No as PaymentProcessNo, tspl_journal_Master.Voucher_No as APAdjustmentNo   " &
            " from TSPL_MP_PAY_PROCESS_DETAIL MPD   " &
            " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No  " &
            " left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code     " &
            " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE  " &
            " inner join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No And tspl_journal_Master.Source_Code='MP-ST'  " &
            " WHERE 2=2 and isnull(MPD.Total_Advance_Amount_Recovery,0)>0 " & CondDisp & "" &
            " union all" &
             " select 'Advance Recovery' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE, TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,0 as Debit,round(MPD.Total_Advance_Amount_Recovery,2) as Credit,MPD.Doc_No as PaymentProcessNo, tspl_journal_Master.Voucher_No as APAdjustmentNo   " &
            " from TSPL_MP_PAY_PROCESS_DETAIL MPD   " &
            " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No  " &
            " left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code     " &
            " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE  " &
            " inner join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No And tspl_journal_Master.Source_Code='MP-ST'  " &
            " WHERE 2=2 and isnull(MPD.Total_Advance_Amount_Recovery,0)>0 " & CondDisp & "" &
            " union all" &
            " select 'Loan Recovery' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE, TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,round(MPD.Total_Loan_Payment_Recovery,2) as Debit,0 as Credit,MPD.Doc_No as PaymentProcessNo, tspl_journal_Master.Voucher_No as APAdjustmentNo  " &
            " from TSPL_MP_PAY_PROCESS_DETAIL MPD  " &
            " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No " &
            " left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code    " &
            " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE " &
            " inner join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No and tspl_journal_Master.Source_Code='MP-LT' " &
            " WHERE 2=2 And isnull(MPD.Total_Loan_Payment_Recovery,0)>0  " & CondDisp & "" &
            " union all" &
            " select 'Loan Recovery' AS Trans_Type,MPD.Farmer_Invoice_No,MPD.Farmer_Invoice_Date,MPD.VSP_CODE, TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,MPD.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,MPH.PaymentDesc as Remarks,0 as Debit,round(MPD.Total_Loan_Payment_Recovery,2) as Credit,MPD.Doc_No as PaymentProcessNo, tspl_journal_Master.Voucher_No as APAdjustmentNo  " &
            " from TSPL_MP_PAY_PROCESS_DETAIL MPD  " &
            " LEFT JOIN TSPL_PAYMENT_PROCESS_HEAD MPH ON MPD.Doc_No=MPH.Doc_No " &
            " left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = MPD.Farmer_Code    " &
            " left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = MPD.VSP_CODE " &
            " inner join tspl_journal_Master on tspl_journal_Master.source_doc_no=MPH.Doc_No and tspl_journal_Master.Source_Code='MP-LT' " &
            " WHERE 2=2 And isnull(MPD.Total_Loan_Payment_Recovery,0)>0  " & CondDisp & ""
        Qry += " union all " & 
              " select 'Payment' as Trans_Type,PAY.Payment_No,PAY.Payment_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.Farmer_Code,TSPL_MP_MASTER.MP_Name as Farmer_Name,'Misc Payment No-' + coalesce(Pay.Misc_Payment_No,'NA') + ', ' + PAY.Entry_Desc,PAY.Payment_Amount as Debit,0 as Credit,PAY.Payment_Process_Code,PAY.Payment_No as APAdjustmentNo from TSPL_MP_PAY_HEAD PAY " &
              " left join TSPL_MP_MASTER on PAY.Farmer_Code=TSPL_MP_MASTER.MP_Code " &
              " left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code " &
              " left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code WHERE 2=2 and isnull(PAY.Posted,0)=1 " & CondRec & "" &
              " union all "
        '" select 'Payment Adjustment(+)' as Trans_Type,PAY.Adjustment_No,PAY.Adjustment_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.Farmer_Code,TSPL_MP_MASTER.MP_Name as Farmer_Name,PAY.Description,round(ADJD.Amount,2) as Debit,0 as Credit,LEFT(replace(Description,'Adjstment amount for next payment cycle->Payment Process doc No: ','') , CHARINDEX(',',replace(Description,'Adjstment amount for next payment cycle->Payment Process doc No: ','') )-1) as PaymentProcessNo,'' as APAdjustmentNo  " &
        '" from TSPL_MP_PAY_ADJ_HEAD PAY  " &
        '" left join (select Adjustment_No,sum(Amount) as Amount from TSPL_MP_PAY_ADJ_DETAIL group by Adjustment_No) ADJD on PAY.Adjustment_No=ADJD.Adjustment_No " &
        '" left join TSPL_MP_MASTER on PAY.Farmer_Code=TSPL_MP_MASTER.MP_Code " &
        '" left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code " &
        '" left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code " &
        '" where PAY.Adjustment_Type='Payment' and isnull(PAY.Is_Post,'Y')='Y' " & CondRec & "" &
        '" union all " &
        '" select 'Payment Adjustment(-)' as Trans_Type,PAY.Adjustment_No,PAY.Adjustment_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.Farmer_Code,TSPL_MP_MASTER.MP_NAME as Farmer_Name,PAY.Description,0 as Debit,round(ADJD.Amount,2) as Credit ,LEFT(replace(Description,'Invoice Setoff ->Payment Process doc No: ','') , CHARINDEX(',',replace(Description,'Invoice Setoff ->Payment Process doc No: ','') )-1) as PaymentProcessNo ,'' as APAdjustmentNo " &
        '" from TSPL_MP_PAY_ADJ_HEAD PAY " &
        ' " left join (select Adjustment_No,sum(Amount) as Amount from TSPL_MP_PAY_ADJ_DETAIL group by Adjustment_No) ADJD on PAY.Adjustment_No=ADJD.Adjustment_No " &
        '" left join TSPL_MP_MASTER on PAY.Farmer_Code=TSPL_MP_MASTER.MP_Code " &
        '" left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code " &
        '" left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code " &
        '" where PAY.Adjustment_Type='Invoice' and isnull(PAY.Is_Post,'Y')='Y' " & CondRec & " " &
        '" Union All " &
        Qry += " select 'Farmer Sale(+)' as Trans_Type,PAY.Document_Code,PAY.Document_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.Farmer_Code,TSPL_MP_MASTER.MP_Name,PAY.Description," &
              " round(PAY.Total_Amt,2) as Debit,0 as Credit ,TSPL_MP_PAY_PROCESS_MCC_SALE.Doc_No as PaymentProcessNo,PAY.Document_Code as APAdjustmentNo from TSPL_MCC_Sale_Farmer_Head PAY   left join TSPL_MP_MASTER on PAY.Farmer_Code=TSPL_MP_MASTER.MP_Code  left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code  left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code left join TSPL_MP_PAY_PROCESS_MCC_SALE on TSPL_MP_PAY_PROCESS_MCC_SALE.Shipment_Doc_No=PAY.Document_Code  " &
              " where  2=2 and ISNULL(PAY.Status,0)=1 " & CondRec & " " &
              " union all  " &
              " select 'Farmer Sale Return(-)' as Trans_Type,PAY.Document_Code,PAY.Document_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.Farmer_Code,TSPL_MP_MASTER.MP_Name,PAY.Description, " &
              " 0 as Debit,round(PAY.Total_Amt,2) as Credit,TSPL_MP_PAY_PROCESS_MCC_SALE_RETURN.Doc_No as PaymentProcessNo,PAY.Document_Code as APAdjustmentNo  from TSPL_MCC_SALE_RETURN_HEAD_FARMER PAY   left join TSPL_MP_MASTER on PAY.Farmer_Code=TSPL_MP_MASTER.MP_Code  left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code  left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code left outer join TSPL_MP_PAY_PROCESS_MCC_SALE_RETURN on TSPL_MP_PAY_PROCESS_MCC_SALE_RETURN.Return_Doc_No=PAY.Document_Code " &
              " where  2=2 and isnull(PAY.Status,0)=1 " & CondRec & "" + Environment.NewLine +
              " union all  " &
              " select  case when PAY.isReceipt = 1 then 'Receipt' else 'Farmer Advance' end as Trans_Type,PAY.Payment_No Document_Code,PAY.Payment_Date as Document_Date,TSPL_VLC_MASTER_HEAD.VSP_Code,TSPL_VENDOR_MASTER.Vendor_Name,PAY.MP_Code_For_Advance as Farmer_Code,TSPL_MP_MASTER.MP_Name,PAY.Entry_Desc as Description,case when PAY.isReceipt = 0 then	round(PAY.Payment_Amount,2) else 0 end as Debit,case when PAY.isReceipt = 1 then	round(PAY.Payment_Amount,2) else 0 end as Credit ,TSPL_PAYMENT_PROCESS_FARMER_MP_ADVANCE.Doc_No as PaymentProcessNo,PAY.Payment_No as APAdjustmentNo  " + Environment.NewLine +
              "from TSPL_PAYMENT_HEADER as PAY" + Environment.NewLine +
              "inner join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=PAY.MP_Code_For_Advance" + Environment.NewLine +
              "left join TSPL_VLC_MASTER_HEAD on TSPL_MP_MASTER.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code  " + Environment.NewLine +
              "left join TSPL_VENDOR_MASTER on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code " &
              "left outer join TSPL_PAYMENT_PROCESS_FARMER_MP_ADVANCE on TSPL_PAYMENT_PROCESS_FARMER_MP_ADVANCE.Payment_No=PAY.Payment_No " &
              " where  2=2 and isnull(PAY.Posted,0)=1  " & CondRec.Replace("PAY.FARMER_CODE", "PAY.MP_Code_For_Advance") & " " + Environment.NewLine +
        " union all " + Environment.NewLine +
        "select 'VCGL' AS Trans_Type,tspl_vcgl_head.document_no,tspl_vcgl_head.document_Date,TSPL_VENDOR_MASTER.vendor_Code as VSP_CODE,TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,TSPL_VCGL_Detail.vcgl_code as Farmer_Code,TSPL_MP_MASTER.MP_Name as Farmer_Name,tspl_vcgl_head.remarks as Remarks,TSPL_VCGL_Detail.Dr_Amount as Debit,TSPL_VCGL_Detail.Cr_Amount as Credit,'' as PaymentProcessNo,tspl_vcgl_head.document_no as APAdjustmentNo 
from TSPL_VCGL_Detail 
 Left outer join tspl_vcgl_head on tspl_vcgl_head.document_no = TSPL_VCGL_Detail.document_no 
 left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = TSPL_VCGL_Detail.vcgl_code 
 left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_MASTER.VLC_Code
 Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = TSPL_VLC_MASTER_HEAD.VSP_Code
 where TSPL_VCGL_Detail.Row_type='Farmer' and tspl_vcgl_head.status=1 " & condVCGL & "" +
        Environment.NewLine + " union all " + Environment.NewLine +
 "select 'VCGL' AS Trans_Type,tspl_vcgl_head.document_no,tspl_vcgl_head.document_Date,TSPL_VENDOR_MASTER.vendor_Code as VSP_CODE,TSPL_VENDOR_MASTER.Vendor_Name as VSP_Name,TSPL_VCGL_HEAD.VC_Code as Farmer_Code,TSPL_MP_MASTER.MP_Name as Farmer_Name,tspl_vcgl_head.remarks as Remarks,TSPL_VCGL_Detail.Cr_Amount as Debit,TSPL_VCGL_Detail.Dr_Amount as Credit,'' as PaymentProcessNo,tspl_vcgl_head.document_no as APAdjustmentNo 
from TSPL_VCGL_Detail 
 Left outer join tspl_vcgl_head on tspl_vcgl_head.document_no = TSPL_VCGL_Detail.document_no 
 left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code = TSPL_VCGL_HEAD.VC_Code 
 left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_MASTER.VLC_Code
 Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.VENDOR_CODE = TSPL_VLC_MASTER_HEAD.VSP_Code
 where TSPL_VCGL_HEAD.Document_Type='F' and TSPL_VCGL_HEAD.status=1" & condVCGL & ""

        Return Qry
    End Function

    Public Shared Function IsDBTDone(ByVal MPCode As String, ByVal trans As SqlTransaction) As Boolean
        Dim retval As Boolean = False
        Dim qry As String = "select top 1  TSPL_DBT_NEFT_DETAIL.PK_Id from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code=TSPL_DBT_NEFT_DETAIL.Document_Code
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
where TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code='" + MPCode + "' order by TSPL_DBT_NEFT.Document_Date desc"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            retval = True
        End If
        Return retval
    End Function

    Friend Shared Sub SaveDefaultMP(ByVal strMPUploaderNo As String, ByVal strMPName As String, ByVal strMPIFSC As String, ByVal strMPAccountNo As String, ByVal strVLCCode As String)
        If True Then
            Dim trans = clsDBFuncationality.GetTransactin()
            Try
                Dim UseVLCUploaderCodeMPUploaderCodeInMCCProcurement As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, clsFixedParameterCode.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, trans)) = 1, True, False)
                Dim SettBankIFSCCodeValidateByService As Boolean = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.BankIFSCCodeValidateByService, clsFixedParameterCode.BankIFSCCodeValidateByService, trans)) > 1) ''Means 2 ERP or 3 Service And ERP
                Dim EnableBankFromMaster As Boolean = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.EnableBankFromMaster & "'", trans)) = 0, False, True)
                Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")

                Dim arrExistCols As New List(Of String)
                Dim dtDefault = New DataTable()
                Dim newBlankRow1 As DataRow = dtDefault.NewRow
                dtDefault.Rows.Add(newBlankRow1)
                Dim objDefaultTemplate As clsExportTemplate = clsExportTemplate.GetDefaultData("MP-IMP-TMP", trans)
                If (objDefaultTemplate IsNot Nothing AndAlso clsCommon.myLen(objDefaultTemplate.Export_Code) > 0) Then
                    If objDefaultTemplate.Arr IsNot Nothing AndAlso objDefaultTemplate.Arr.Count > 0 Then
                        For Each objTr As clsExportTemplateDetail In objDefaultTemplate.Arr
                            If clsCommon.myLen(objTr.Column_Name) > 0 Then
                                arrExistCols.Add(objTr.Column_Name)
                                Dim newColumn As New DataColumn(clsCommon.myCstr(objTr.Column_Name), GetType(System.String))
                                dtDefault.Columns.Add(newColumn)
                                dtDefault.Rows(0).Item(clsCommon.myCstr(objTr.Column_Name)) = clsCommon.myCstr(objTr.Column_Header)
                            End If
                        Next
                    End If
                Else
                    Throw New Exception("Please set default Templeate in Farmer Master")
                End If

                Dim obj As New clsMpMaster()
                Dim strData As String = ""
                If clsCommon.myLen(strVLCCode) > 0 Then
                    strData = strVLCCode
                ElseIf UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = True Then
                    If clsCommon.myLen(clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPUploaderCode))) <= 0 Then
                        Throw New Exception("MP Uploader Code Can Not Be Left Blank")
                    End If
                    strData = clsDBFuncationality.getSingleValue("select VLC_Code from TSPL_VLC_MASTER_HEAD where VLC_Code_VLC_Uploader='" & clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colDCSUploaderCode)) & "'", trans)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("VLC Uploader Code Not Found")
                    End If
                Else
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colDCSVLCCode))
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("VLC Code Can Not Be Left Blank")
                    End If
                    If clsCommon.myLen(strData) > 30 Then
                        Throw New Exception("VLC Code Can Not Be Larger Then 30 Charachter")
                    End If
                End If

                obj.MCC_Code = strData
                strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPCode))

                obj.MP_Code = strData

                If clsCommon.myLen(strMPName) > 0 Then
                    strData = strMPName
                Else
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPName))
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("MP Name Can Not Be Left Blank")
                    End If
                    If clsCommon.myLen(strData) > 50 Then
                        Throw New Exception("MP Name Can Not Be Larger Then 50 Charachter")
                    End If
                End If
                obj.MP_Name = strData


                If arrExistCols.Contains(clsMasterDefault.colMPFatherName) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPFatherName)) > 0 Then
                        obj.Father_Name = dtDefault.Rows(0).Item(clsMasterDefault.colMPFatherName)
                    End If
                End If



                strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAddress1))
                If clsCommon.myLen(strData) <= 0 Then
                    Throw New Exception("Address1 Can Not Be Left Blank")
                End If
                If clsCommon.myLen(strData) > 50 Then
                    Throw New Exception("Address1 Can Not Be Larger Then 50 Charachter")
                End If
                obj.Add1 = strData

                If arrExistCols.Contains(clsMasterDefault.colMPAddress2) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPAddress2)) > 0 Then
                        obj.Add2 = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAddress2))
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPZila) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPZila)) > 0 Then
                        obj.Zila = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPZila))
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPTehsil) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPTehsil)) > 0 Then
                        obj.Tehsil = dtDefault.Rows(0).Item(clsMasterDefault.colMPTehsil)
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPPinCode) Then
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPPinCode))
                    obj.Pin_code = strData
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPCityCode) Then
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPCityCode))
                    obj.City_code = strData
                End If

                If arrExistCols.Contains(clsMasterDefault.colMPStateCode) Then
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPStateCode))
                    obj.State_Code = strData
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPCountryCode) Then
                    strData = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPCountryCode))
                    obj.Country_code = strData
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPTelphone) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPTelphone)) > 0 Then
                        obj.Telphone = dtDefault.Rows(0).Item(clsMasterDefault.colMPTelphone)
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPEmail) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPEmail)) > 0 Then
                        Dim check As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(dtDefault.Rows(0).Item(clsMasterDefault.colMPEmail), "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
                        If check.Success Then
                            obj.Email = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPEmail))
                        Else
                            Throw New Exception("Email Is In Invalid Format. It Shoud Be as UserName@Domain Format ")
                        End If
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPAadharNo) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPAadharNo)) > 0 Then
                        'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*) from tspl_mp_master where fax='" & clsCommon.myCstr(dtDefault.Rows(0).Item(colMPAadharNo)) & "' and MP_Code<>'" & clsCommon.myCstr(obj.MP_Code) & "'", trans)) > 0 Then
                        '    Throw New Exception("Same Aadhar No is exist with another MP so please change Aadhar No because Aadhar No is unique.")
                        'End If
                        obj.Fax = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAadharNo))
                    End If
                End If

                If arrExistCols.Contains(clsMasterDefault.colMPJanAadharNo) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPJanAadharNo)) > 0 Then
                        'If clsCommon.myLen(dtDefault.Rows(0).Item(colMPJanAadharNo)) <> 10 Then
                        '    Throw New Exception("Invalid Jan Aadhar No.Please Enter 10 Digit Jan Aadhar No")
                        'End If
                        obj.Jan_Aadhar_No = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPJanAadharNo))
                        'Else
                        '    If SettJanAadharNoMandatory Then
                        '        Throw New Exception("Please Fill Jan Aadhar No")
                        '    End If
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPDateofBirth) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPDateofBirth)) > 0 Then
                        If IsDate(dtDefault.Rows(0).Item(clsMasterDefault.colMPDateofBirth)) Then
                            obj.DOB = clsCommon.myCDate(dtDefault.Rows(0).Item(clsMasterDefault.colMPDateofBirth), "dd/MMM/yyyy")
                        Else
                            Throw New Exception("Invalid Date Found For Date Of Birth")
                        End If
                    End If
                End If


                'MP Uploader Code
                If clsCommon.myLen(strMPUploaderNo) > 0 Then
                    obj.MP_CODE_VLC_UPLOADER = strMPUploaderNo
                ElseIf clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPUploaderCode)) > 0 Then
                    obj.MP_CODE_VLC_UPLOADER = dtDefault.Rows(0).Item(clsMasterDefault.colMPUploaderCode)
                Else
                    Throw New Exception("Please Fill MP Uploader Code")
                End If

                Dim qqq As String = "select COUNT(*) from TSPL_MP_MASTER where MP_Code_vlc_uploader='" & obj.MP_CODE_VLC_UPLOADER & "' and mp_code<>'" & obj.MP_Code & "' and vlc_Code='" & obj.MCC_Code & "'"
                If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qqq, trans)) >= 1 Then
                    Throw New Exception("Duplicate MP uploader Code")
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPEducation) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPEducation)) > 0 Then
                        obj.Education = dtDefault.Rows(0).Item(clsMasterDefault.colMPEducation)
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPTOLERANCE) Then
                    If clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPTOLERANCE)) > 0 Then
                        If IsNumeric(dtDefault.Rows(0).Item(clsMasterDefault.colMPTOLERANCE)) = False Then
                            Throw New Exception("TOLERANCE value should be Numeric")
                        End If
                        If clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPTOLERANCE)) > 100 Then
                            Throw New Exception("TOLERANCE value should be less then 100.")
                        End If
                        obj.TOLERANCE = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPTOLERANCE))
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPTOLERANCE) Then
                    obj.Land_Holding = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPTOLERANCE))
                End If
                obj.No_Of_Animal = 0
                If arrExistCols.Contains(clsMasterDefault.colMPNoOfMilchAnimal) Then
                    obj.No_Of_breedable_milk_animal = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPNoOfMilchAnimal))
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPTotalMilkProduction) Then
                    obj.Milk_production = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPTotalMilkProduction))
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPMilkForSelfConsumption) Then
                    obj.Milk_Home_consumption = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPMilkForSelfConsumption))
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPMilkForSale) Then
                    obj.Milk_For_sale = clsCommon.myCdbl(dtDefault.Rows(0).Item(clsMasterDefault.colMPMilkForSale))
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPPayeeName) Then
                    obj.PayeeName = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPPayeeName))
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPCustomerAccSet) Then
                    obj.Cust_Account = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPCustomerAccSet))
                    If clsCommon.myLen(obj.Cust_Account) > 0 Then
                        '' check valid acc set
                        qqq = "select Cust_Account from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" & obj.Cust_Account & "' "
                        Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                        If clsCommon.myLen(checkCode) <= 0 Then
                            Throw New Exception("Invalid Customer Account Set- " & obj.Cust_Account & "")
                        End If
                    End If
                End If
                If arrExistCols.Contains(clsMasterDefault.colMPVendorAccSet) Then
                    obj.Acct_Set_Code = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPVendorAccSet))
                    If clsCommon.myLen(obj.Acct_Set_Code) > 0 Then
                        '' check valid acc set
                        qqq = "select Acct_Set_Code from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code='" & obj.Acct_Set_Code & "' "
                        Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                        If clsCommon.myLen(checkCode) <= 0 Then
                            Throw New Exception("Invalid Vendor Account Set- " & obj.Acct_Set_Code & " ")
                        End If
                    End If
                End If
                If SettBankIFSCCodeValidateByService Then
                    If clsCommon.myLen(strMPIFSC) > 0 Then
                        Dim arrFilter As New Dictionary(Of String, String)
                        arrFilter.Add("IFSC", strMPIFSC)
                        arrFilter.Add("OutPutType", "1")
                        Dim dt As DataTable = Xtra.GetDataFromAPI(EnumAPI.BankIFSC, "GetIFSC", arrFilter)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Invalid IFSC Code")
                        End If
                        If clsCommon.CompairString(dt.Rows(0)("Result"), "true") = CompairStringResult.Equal Then
                            obj.BankName = clsCommon.myCstr(dt.Rows(0)("BANK"))
                            obj.BankBranch = clsCommon.myCstr(dt.Rows(0)("BRANCH"))
                            obj.BankStateCode = clsCommon.myCstr(dt.Rows(0)("STATE"))
                            obj.BankCityCode = clsCommon.myCstr(dt.Rows(0)("CITY"))
                            obj.IFCICode = clsCommon.myCstr(dt.Rows(0)("IFSC"))
                            If clsCommon.myLen(strMPAccountNo) > 0 Then
                                obj.AccountNO = strMPAccountNo
                            Else
                                If arrExistCols.Contains(clsMasterDefault.colMPAccountNo) Then
                                    If clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E+") OrElse clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E-") Then
                                        obj.AccountNO = Decimal.Parse(clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)), System.Globalization.NumberStyles.Float)
                                    Else
                                        obj.AccountNO = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo))
                                    End If
                                End If
                            End If
                        Else
                            Throw New Exception(dt.Rows(0)("Response"))
                        End If
                    End If
                ElseIf EnableBankFromMaster AndAlso arrExistCols.Contains(clsMasterDefault.colMPBankCode) AndAlso clsCommon.myLen(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode)) > 0 Then
                    If clsDBFuncationality.getSingleValue("select count(*) from TSPL_Vendor_Bank_Master where bank_code='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "'", trans) > 0 Then
                        Dim bnkDt As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_Vendor_Bank_Master where bank_code='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "'", trans)
                        obj.BankName = clsCommon.myCstr(bnkDt.Rows(0)("BANK_CODE"))
                        obj.BankCityCode = clsCommon.myCstr(bnkDt.Rows(0)("City_Code"))
                        obj.BankStateCode = clsCommon.myCstr(bnkDt.Rows(0)("State_Code"))
                        If clsCommon.myLen(strMPAccountNo) > 0 Then
                            obj.AccountNO = strMPAccountNo
                        Else
                            If arrExistCols.Contains(clsMasterDefault.colMPAccountNo) Then
                                If clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E+") OrElse clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E-") Then
                                    obj.AccountNO = Decimal.Parse(clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)), System.Globalization.NumberStyles.Float)
                                Else
                                    obj.AccountNO = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo))
                                End If
                            End If
                        End If
                        obj.BankBranch = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch))
                        obj.IFCICode = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode))
                        If clsCommon.myLen(obj.IFCICode) > 0 Then
                            qqq = "select Bank_IFSC_Code from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "' and Bank_IFSC_Code = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode) & "' "
                            Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                            If clsCommon.myLen(checkCode) <= 0 Then
                                Throw New Exception("Invalid IFSC Code for Bank Code - " & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & " ")
                            End If
                            If clsCommon.myLen(obj.BankBranch) <= 0 Then
                                qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "' and Bank_IFSC_Code = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode) & "' "
                                Dim BranchCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                If clsCommon.myLen(BranchCode) > 0 Then
                                    obj.BankBranch = BranchCode
                                    dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch) = BranchCode
                                End If
                            End If

                        End If
                        If clsCommon.myLen(obj.BankBranch) > 0 Then
                            qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "' and Branch_Name = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch) & "' "
                            Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                            If clsCommon.myLen(checkCode) <= 0 Then
                                Throw New Exception("Invalid Bank Branch for Bank Code - " & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & " ")
                            End If
                            If clsCommon.myLen(obj.IFCICode) <= 0 Then
                                qqq = "select Bank_IFSC_Code from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "' and Branch_Name = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch) & "' "
                                Dim IFSCCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                If clsCommon.myLen(IFSCCode) > 0 Then
                                    obj.IFCICode = IFSCCode
                                    dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode) = IFSCCode
                                End If
                            End If

                        End If

                        If clsCommon.myLen(obj.BankBranch) > 0 AndAlso clsCommon.myLen(obj.IFCICode) > 0 Then
                            qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "' and Branch_Name = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch) & "' and Bank_IFSC_Code = '" & dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode) & "' "
                            Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                            If clsCommon.myLen(checkCode) <= 0 Then
                                Throw New Exception("Invalid IFSC Code for Branch - [" & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankBranch) & "] ,Bank Code - " & dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode) & "")
                            End If
                        End If
                    End If
                ElseIf EnableBankFromMaster = False Then
                    If arrExistCols.Contains(clsMasterDefault.colMPBankCode) Then
                        obj.BankName = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode))
                        obj.BankCityCode = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode))
                    End If

                    If arrExistCols.Contains(clsMasterDefault.colMPBankStateCode) Then
                        obj.BankStateCode = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankStateCode))
                    End If

                    If clsCommon.myLen(strMPAccountNo) > 0 Then
                        obj.AccountNO = strMPAccountNo
                    Else
                        If arrExistCols.Contains(clsMasterDefault.colMPAccountNo) Then
                            If clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E+") OrElse clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)).Contains("E-") Then
                                obj.AccountNO = Decimal.Parse(clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo)), System.Globalization.NumberStyles.Float)
                            Else
                                obj.AccountNO = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPAccountNo))
                            End If
                        End If
                    End If
                    If arrExistCols.Contains(clsMasterDefault.colMPBankCode) Then
                        obj.BankBranch = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPBankCode))
                    End If
                    If clsCommon.myLen(strMPIFSC) > 0 Then
                        obj.IFCICode = strMPIFSC
                    Else
                        If arrExistCols.Contains(clsMasterDefault.colMPIFSCCode) Then
                            obj.IFCICode = clsCommon.myCstr(dtDefault.Rows(0).Item(clsMasterDefault.colMPIFSCCode))
                        End If
                    End If
                End If
                If clsMpMaster.IsDBTDone(obj.MP_Code, trans) Then
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable("select AccountNO,IFCICode From TSPL_MP_MASTER where MP_Code='" + obj.MP_Code + "'", trans)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        obj.AccountNO = clsCommon.myCstr(dt.Rows(0)("AccountNO"))
                        obj.IFCICode = clsCommon.myCstr(dt.Rows(0)("IFCICode"))
                    End If
                End If

                'If clsCommon.myLen(obj.AccountNO) > 0 Then
                '    If clsCommon.MySpecialChars(obj.AccountNO, EnumSpecialChactersType.Digit) Then
                '        Throw New Exception("Invalid AccountNO [" + clsCommon.myCstr(dtDefault.Rows(0).Item(colMPAccountNo)) + "]")
                '    End If
                'End If

                obj.ArrMPIncentiveMapping = New ArrayList()



                '=====================================================================================

                If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & obj.MP_Code & "'", trans) = 0 Then
                    Dim dt As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
                    If clsCommon.myLen(obj.MP_Code) <= 0 Then
                        obj.MP_Code = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.MPMaster, "", "")
                    End If

                    If obj.No_Of_breedable_milk_animal > 0 Then
                        Dim objAnimal As clsAnimalDetails
                        obj.arrAnimalDetail = New List(Of clsAnimalDetails)
                        For j As Integer = 0 To obj.No_Of_breedable_milk_animal
                            objAnimal = New clsAnimalDetails
                            objAnimal.Prog_Code = clsUserMgtCode.frmMPMaster
                            objAnimal.Trans_Code = obj.MP_Code
                            objAnimal.Line_No = (j + 1)
                            objAnimal.Bread_of_Animal = "N/A"
                            objAnimal.Type_Of_Animal = "N/A"
                            obj.arrAnimalDetail.Add(objAnimal)
                        Next
                    End If

                    obj.isNewEntry = True
                    obj.Modified_By = objCommonVar.CurrentUserCode
                    obj.Modified_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                    obj.Comp_Code = objCommonVar.CurrentCompanyCode
                    obj.Created_By = objCommonVar.CurrentUserCode
                    obj.Created_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                    clsMpMaster.SaveData(obj, trans)
                Else
                    obj.Modified_By = objCommonVar.CurrentUserCode
                    obj.Modified_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                    obj.isNewEntry = False
                    clsMpMaster.SaveData(obj, trans)
                End If
                trans.Commit()
            Catch ex As Exception
                trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub
End Class

Public Class clsBuffaloesDetails
    Public Prog_Code As String = Nothing
    Public Trans_Code As String = Nothing
    Public Line_No As Integer = Nothing
    Public Bread_of_Buffalo As String = Nothing
    Public Shared Function SaveData(ByVal arr As List(Of clsBuffaloesDetails), ByVal Trans As SqlTransaction) As Boolean
        Try
            Dim i As Integer = 0

            Dim issaved As Boolean = True
            If arr.Count > 0 Then
                deleteData(arr.Item(0).Prog_Code, arr.Item(0).Trans_Code, Trans)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Prog_Code", arr.Item(i).Prog_Code)
                    clsCommon.AddColumnsForChange(coll, "Trans_Code", arr.Item(i).Trans_Code)
                    clsCommon.AddColumnsForChange(coll, "Line_No", arr.Item(i).Line_No)
                    clsCommon.AddColumnsForChange(coll, "Bread_of_Buffalo", arr.Item(i).Bread_of_Buffalo)
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "tspl_Buffaloes_Details", OMInsertOrUpdate.Insert, "", Trans)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function SaveData(ByVal obj As clsBuffaloesDetails, ByVal Trans As SqlTransaction) As Boolean
        Try
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Prog_Code", obj.Prog_Code)
            clsCommon.AddColumnsForChange(coll, "Trans_Code", obj.Trans_Code)
            clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
            clsCommon.AddColumnsForChange(coll, "Bread_of_Buffalo", obj.Bread_of_Buffalo)
            If clsDBFuncationality.getSingleValue("select count(*) from TSPL_Gen_Set_Detail where prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "'", Trans) = 0 Then
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_Buffaloes_Details", OMInsertOrUpdate.Insert, "", Trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_Buffaloes_Details", OMInsertOrUpdate.Update, "  prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "'", Trans)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function LoadData(ByVal pcode As String, ByVal tcode As String) As List(Of clsBuffaloesDetails)
        Dim arr As New List(Of clsBuffaloesDetails)
        Try
            Dim obj As New clsBuffaloesDetails
            Dim q As String = "select * from tspl_Buffaloes_Details where Prog_Code='" & pcode & "' and Trans_Code='" & tcode & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsBuffaloesDetails
                    obj.Prog_Code = clsCommon.myCstr(dtbl.Rows(i)("Prog_Code"))
                    obj.Trans_Code = clsCommon.myCstr(dtbl.Rows(i)("Trans_Code"))
                    obj.Line_No = clsCommon.myCdbl(dtbl.Rows(i)("Line_No"))
                    obj.Bread_of_Buffalo = clsCommon.myCstr(dtbl.Rows(i)("Bread_of_Buffalo"))
                    arr.Add(obj)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return arr
    End Function
    Public Shared Function deleteData(ByVal pcode As String, ByVal tcode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = "delete from tspl_Buffaloes_Details where prog_code='" & pcode & "' and trans_code='" & tcode & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
    End Function
End Class

Public Class clsCowDetails
    Public Prog_Code As String = Nothing
    Public Trans_Code As String = Nothing
    Public Line_No As Integer = Nothing
    Public Bread_of_cow As String = Nothing
    Public Shared Function LoadData(ByVal pcode As String, ByVal tcode As String) As List(Of clsCowDetails)
        Dim arr As New List(Of clsCowDetails)
        Try
            Dim obj As New clsCowDetails
            Dim q As String = "select * from tspl_cow_Details where Prog_Code='" & pcode & "' and Trans_Code='" & tcode & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsCowDetails
                    obj.Prog_Code = clsCommon.myCstr(dtbl.Rows(i)("Prog_Code"))
                    obj.Trans_Code = clsCommon.myCstr(dtbl.Rows(i)("Trans_Code"))
                    obj.Line_No = clsCommon.myCdbl(dtbl.Rows(i)("Line_No"))
                    obj.Bread_of_cow = clsCommon.myCstr(dtbl.Rows(i)("bread_of_cow"))
                    arr.Add(obj)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return arr
    End Function
    Public Shared Function SaveData(ByVal arr As List(Of clsCowDetails), ByVal Trans As SqlTransaction) As Boolean
        Try
            Dim i As Integer = 0
            Dim issaved As Boolean = True
            If arr.Count > 0 Then
                deleteData(arr.Item(0).Prog_Code, arr.Item(0).Trans_Code, Trans)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Prog_Code", arr.Item(i).Prog_Code)
                    clsCommon.AddColumnsForChange(coll, "Trans_Code", arr.Item(i).Trans_Code)
                    clsCommon.AddColumnsForChange(coll, "Line_No", arr.Item(i).Line_No)
                    clsCommon.AddColumnsForChange(coll, "bread_of_cow", arr.Item(i).Bread_of_cow)
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "tspl_Cow_Details", OMInsertOrUpdate.Insert, "", Trans)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function deleteData(ByVal pcode As String, ByVal tcode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = "delete from tspl_cow_Details where prog_code='" & pcode & "' and trans_code='" & tcode & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
    End Function
    Public Shared Function SaveData(ByVal obj As clsCowDetails, ByVal Trans As SqlTransaction) As Boolean
        Try
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Prog_Code", obj.Prog_Code)
            clsCommon.AddColumnsForChange(coll, "Trans_Code", obj.Trans_Code)
            clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
            clsCommon.AddColumnsForChange(coll, "bread_of_cow", obj.Bread_of_cow)
            If clsDBFuncationality.getSingleValue("select count(*) from tspl_cow_Details where prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "'", Trans) = 0 Then
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_cow_Details", OMInsertOrUpdate.Insert, "", Trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_cow_Details", OMInsertOrUpdate.Update, "  prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "'", Trans)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
End Class

Public Class clsAnimalDetails
    Public Prog_Code As String = Nothing
    Public Trans_Code As String = Nothing
    Public Line_No As Integer = Nothing
    Public Bread_of_Animal As String = Nothing
    Public Type_Of_Animal As String = String.Empty
    Public Count_Of_animal As Integer = 0
    Public Animal_Staus As String = Nothing
    Public MP_Code As String = String.Empty
    Public Shared Function SaveData(ByVal arr As List(Of clsAnimalDetails), ByVal Trans As SqlTransaction) As Boolean
        Try
            Dim i As Integer = 0

            Dim issaved As Boolean = True
            If arr IsNot Nothing AndAlso arr.Count > 0 Then
                deleteData(arr.Item(0).Prog_Code, arr.Item(0).Trans_Code, Trans)
                For i = 0 To arr.Count - 1
                    Dim coll As New Hashtable()
                    clsCommon.AddColumnsForChange(coll, "Prog_Code", arr.Item(i).Prog_Code)
                    clsCommon.AddColumnsForChange(coll, "Trans_Code", arr.Item(i).Trans_Code)
                    clsCommon.AddColumnsForChange(coll, "Line_No", arr.Item(i).Line_No)
                    clsCommon.AddColumnsForChange(coll, "Type_Of_Animal", arr.Item(i).Type_Of_Animal)
                    clsCommon.AddColumnsForChange(coll, "Bread_of_Animal", arr.Item(i).Bread_of_Animal)
                    clsCommon.AddColumnsForChange(coll, "Count_of_Animal", arr.Item(i).Count_Of_animal)
                    clsCommon.AddColumnsForChange(coll, "Animal_staus", arr.Item(i).Animal_Staus)
                    issaved = issaved And clsCommonFunctionality.UpdateDataTable(coll, "tspl_Animal_Details", OMInsertOrUpdate.Insert, "", Trans)
                Next
            End If
            Return issaved
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function SaveData(ByVal obj As clsAnimalDetails, ByVal Trans As SqlTransaction) As Boolean
        Return SaveData(True, obj, Trans)
    End Function

    Public Shared Function SaveData(ByVal isRunDeleteQuery As Boolean, ByVal obj As clsAnimalDetails, ByVal Trans As SqlTransaction) As Boolean
        Try
            If isRunDeleteQuery Then
                deleteData(obj.Prog_Code, obj.Trans_Code, Trans)
            End If
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Prog_Code", obj.Prog_Code)
            clsCommon.AddColumnsForChange(coll, "Trans_Code", obj.Trans_Code)
            clsCommon.AddColumnsForChange(coll, "Line_No", obj.Line_No)
            clsCommon.AddColumnsForChange(coll, "Type_Of_Animal", obj.Type_Of_Animal)
            clsCommon.AddColumnsForChange(coll, "Bread_of_Animal", obj.Bread_of_Animal)
            clsCommon.AddColumnsForChange(coll, "Count_of_Animal", obj.Count_Of_animal)
            clsCommon.AddColumnsForChange(coll, "Animal_Staus", obj.Animal_Staus)
            If clsDBFuncationality.getSingleValue("select count(*) from tspl_Animal_Details where prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "' and Type_Of_Animal='" & obj.Type_Of_Animal & "' ", Trans) = 0 Then
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_Animal_Details", OMInsertOrUpdate.Insert, "", Trans)
            Else
                clsCommonFunctionality.UpdateDataTable(coll, "tspl_Animal_Details", OMInsertOrUpdate.Update, "  prog_code='" & obj.Prog_Code & "' and trans_code='" & obj.Trans_Code & "' and Line_No='" & obj.Line_No & "' and Type_Of_Animal='" & obj.Type_Of_Animal & "' ", Trans)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function LoadData(ByVal pcode As String, ByVal tcode As String) As List(Of clsAnimalDetails)
        Dim arr As New List(Of clsAnimalDetails)
        Try
            Dim obj As New clsAnimalDetails
            Dim q As String = "select * from tspl_Animal_Details where Prog_Code='" & pcode & "' and Trans_Code='" & tcode & "'"
            Dim dtbl As DataTable = clsDBFuncationality.GetDataTable(q)
            If dtbl IsNot Nothing AndAlso dtbl.Rows.Count > 0 Then
                For i As Integer = 0 To dtbl.Rows.Count - 1
                    obj = New clsAnimalDetails
                    obj.Prog_Code = clsCommon.myCstr(dtbl.Rows(i)("Prog_Code"))
                    obj.Trans_Code = clsCommon.myCstr(dtbl.Rows(i)("Trans_Code"))
                    obj.Line_No = clsCommon.myCdbl(dtbl.Rows(i)("Line_No"))
                    obj.Bread_of_Animal = clsCommon.myCstr(dtbl.Rows(i)("Bread_of_Animal"))
                    obj.Type_Of_Animal = clsCommon.myCstr(dtbl.Rows(i)("Type_Of_Animal"))
                    obj.Count_Of_animal = clsCommon.myCstr(dtbl.Rows(i)("Count_Of_Animal"))
                    obj.Animal_Staus = clsCommon.myCstr(dtbl.Rows(i)("Animal_Staus"))
                    arr.Add(obj)
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return arr
    End Function
    Public Shared Function deleteData(ByVal pcode As String, ByVal tcode As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim qry As String = "delete from tspl_Animal_Details where prog_code='" & pcode & "' and trans_code='" & tcode & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
    End Function
End Class

Public Class clsMPIncentiveMapping
#Region "Variables"
    Public TR_CODE As String = Nothing
    Public MP_CODE As String = Nothing
    Public INCENTIVE_CODE As String = Nothing

#End Region

    Public Shared Function SaveData(ByVal strMPCode As String, ByVal dtDocDate As DateTime, ByVal Arr As ArrayList, ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each strIncentiveCode As String In Arr
                Dim coll As New Hashtable()
                Dim strTRCode As String = clsERPFuncationality.GetNextCode(trans, dtDocDate, clsDocType.Detail, clsDocTransactionType.Detail, "")
                clsCommon.AddColumnsForChange(coll, "TR_CODE", strTRCode)
                clsCommon.AddColumnsForChange(coll, "MP_CODE", strMPCode)
                clsCommon.AddColumnsForChange(coll, "INCENTIVE_CODE", strIncentiveCode) ' strIncentiveCode.INCENTIVE_CODE
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_MP_INCENTIVE", OMInsertOrUpdate.Insert, "", trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, clsCommon.myCstr(strTRCode), "TSPL_MP_INCENTIVE", "TR_CODE", trans)

            Next
        End If
        Return True
    End Function
    Public Shared Function GetData(ByVal strCode As String) As ArrayList ' , ByVal trans As SqlTransaction
        Dim arr As ArrayList = Nothing
        Dim qry As String = "Select TSPL_MP_INCENTIVE.* from TSPL_MP_INCENTIVE Where MP_CODE='" + strCode + "'  "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            arr = New ArrayList()
            For Each dr As DataRow In dt.Rows
                arr.Add(clsCommon.myCstr(dr("INCENTIVE_CODE")))
            Next
        End If
        Return arr
    End Function

End Class
