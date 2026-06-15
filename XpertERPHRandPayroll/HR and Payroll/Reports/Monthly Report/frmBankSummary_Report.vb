'--01/08/2013--form Add By- Pradeep Sharma ---------
Imports common
Imports System.Data
Imports System.Data.SqlClient
Imports XpertERPEngine
Imports Telerik.WinControls.UI
Imports System.IO

Public Class frmBankSummary_Report
    Inherits FrmMainTranScreen
    Dim ButtonToolTip As ToolTip = New ToolTip()
#Region "Variable"
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Dim Qry As String

#End Region

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If rdbEmployeeBankAdvice.IsChecked = True Then
            PrintAdvice()
        Else
            PrintData(True)
        End If

    End Sub

    Private Sub frmBankSummary_Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetUserMgmtNew()
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+P for Save/Update ")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        RadPageView1.SelectedPage = RadPageViewPage1
    End Sub
    'Sub LoadData()

    '    qry = ""
    '    qry += " SELECT DISTINCT TSPL_EMPLOYEE_MASTER .BANK_CODE ,TSPL_BANK_MASTER.DESCRIPTION AS 'Bank Name' "
    '    qry += " FROM TSPL_GENERATE_SALARY "
    '    qry += " LEFT OUTER JOIN TSPL_GENERATE_SALARY_ATTENDANCE  ON TSPL_GENERATE_SALARY_ATTENDANCE  .SALARY_GENERATION_CODE = TSPL_GENERATE_SALARY .SALARY_GENERATION_CODE  left outer join TSPL_EMPLOYEE_MASTER on TSPL_GENERATE_SALARY_ATTENDANCE .EMP_CODE= TSPL_EMPLOYEE_MASTER .EMP_CODE  "
    '    Qry += " left outer join TSPL_BANK_MASTER  on TSPL_BANK_MASTER .BANK_CODE = TSPL_EMPLOYEE_MASTER .BANK_CODE "
    '    Qry += " Left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code= TSPL_GENERATE_SALARY.Location_Code"
    '    Qry += " where TSPL_GENERATE_SALARY.PAY_PERIOD_CODE ='" + txtFromPP.Value + "' "

    '    If txtLocationMult.arrValueMember IsNot Nothing AndAlso txtLocationMult.arrValueMember.Count > 0 Then
    '        Qry += " and TSPL_LOCATION_MASTER.Location_Code  in (" + clsCommon.GetMulcallString(txtLocationMult.arrValueMember) + ") "
    '    End If


    '    qry += " ORDER BY TSPL_BANK_MASTER.DESCRIPTION "
    '    cbgLocation.DataSource = clsDBFuncationality.GetDataTable(qry)
    '    cbgLocation.ValueMember = "BANK_CODE"
    '    cbgLocation.DisplayMember = "Bank Name"
    'End Sub
    Private Sub txtCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtFromPP._MYValidating
        Dim qry As String = "SELECT PAY_PERIOD_CODE AS 'Code',(DATEDIFF(DAY,date_from,date_to)+1) as 'Total days', " _
            & " PAY_PERIOD_NAME as 'Pay Period Name' FROM TSPL_PAYPERIOD_MASTER  "
        txtFromPP.Value = clsCommon.ShowSelectForm("TSPL_PAYPERIOD_MASTER", qry, "Code", "POSTED=1 AND FREEZED=0", txtFromPP.Value, "", isButtonClicked)
        lblFrompp.Text = clsPayPeriodMaster.GetName(txtFromPP.Value, Nothing)

    End Sub
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.frmBankSummary_Report)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
            'Me.Close()
            'Exit Function
        End If
        '' Anubhooti 23-July-2014 (BM00000003141)
        'btnSave.Visible = MyBase.isModifyFlag
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Sub PrintAdvice()
        Try
            If clsCommon.myLen(txtFromPP.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please Select Pay Period.", Me.Text)
                Return
            End If

            Dim Comp_Logo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Logo_Img from TSPL_COMPANY_MASTER where Comp_code='" + objCommonVar.CurrentCompanyCode + "'"))
            Dim Qry As String = ""
            Qry = " select  '" + Comp_Logo + "'  as Logo_Img,max(TSPL_LOCATION_MASTER.Add1+Case When ISNULL(TSPL_LOCATION_MASTER.Add2,'')='' Then '' else ', '+TSPL_LOCATION_MASTER.Add2+ Case When ISNULL(TSPL_LOCATION_MASTER.Add3,'')='' Then '' Else ', '+TSPL_LOCATION_MASTER.Add3+ Case When ISNULL(TSPL_STATE_MASTER.State_Name ,'')='' Then '' else '-'+CONVERT(varchar, TSPL_STATE_MASTER.State_Name) End End End ) as Loc_Add,
                    sum( case when TSPL_PAYHEAD_MASTER.isearning=1 then  TSPL_GENERATE_SALARY_PAYHEADS.ACTUAL_AMOUNT else -TSPL_GENERATE_SALARY_PAYHEADS.ACTUAL_AMOUNT end  )as 'sum amt', 
                    max(TSPL_GENERATE_SALARY.PAY_PERIOD_CODE) as pay_period, '" + objCommonVar.CurrentCompanyName + "'  as Company_Name,'" + objCommonVar.CurrentCompanyCode + "' as Comp_Code ,MAX(TSPL_LOCATION_MASTER.Location_Code) as Location_Code, (TSPL_EMPLOYEE_MASTER.BANK_CODE) as Bank_Name,max(TSPL_BANK_MASTER.DESCRIPTION) as Bank_Desc ,max(TSPL_EMPLOYEE_MASTER.Emp_Name)Emp_Name,TSPL_EMPLOYEE_MASTER.EMP_CODE,max(TSPL_EMPLOYEE_MASTER.BANK_ACC_NO)BANK_ACC_NO

from TSPL_GENERATE_SALARY_PAYHEADS 
left outer join TSPL_GENERATE_SALARY on TSPL_GENERATE_SALARY_PAYHEADS.SALARY_GENERATION_CODE =TSPL_GENERATE_SALARY.SALARY_GENERATION_CODE  
left join TSPL_PAYHEAD_MASTER on TSPL_PAYHEAD_MASTER.Pay_HEAD_Code=TSPL_GENERATE_SALARY_PAYHEADS.Pay_HEAD_Code  
left outer join TSPL_PAYPERIOD_MASTER on TSPL_PAYPERIOD_MASTER.PAY_PERIOD_CODE =TSPL_GENERATE_SALARY.PAY_PERIOD_CODE   
left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_GENERATE_SALARY_PAYHEADS.EMP_CODE 
left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=TSPL_EMPLOYEE_MASTER.BANK_CODE
left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_GENERATE_SALARY.LOCATION_CODE   
left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State  
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER  .Comp_Code =TSPL_LOCATION_MASTER.Comp_code   
left join TSPL_DESIGNATION_MASTER on TSPL_DESIGNATION_MASTER .Designation_id =TSPL_EMPLOYEE_MASTER.Designation  
left outer join TSPL_DEVISION_MASTER   on TSPL_DEVISION_MASTER .DEVISION_CODE  =TSPL_GENERATE_SALARY.DEVISION_CODE  
left outer join TSPL_PAYMENT_MODE PM on TSPL_EMPLOYEE_MASTER.PAYMENT_MODE_New  = PM.Code 
WHERE  TSPL_GENERATE_SALARY.PAY_PERIOD_CODE = '" + txtFromPP.Value + "' "
            If txtLocationMult.arrValueMember IsNot Nothing AndAlso txtLocationMult.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_LOCATION_MASTER.LOCATION_CODE  in (" + clsCommon.GetMulcallString(txtLocationMult.arrValueMember) + ") "
            End If

            Qry += "Group by TSPL_EMPLOYEE_MASTER.BANK_CODE,TSPL_EMPLOYEE_MASTER.EMP_CODE "

            Dim DT As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If DT IsNot Nothing AndAlso DT.Rows.Count > 0 Then
                Dim frmcrystal As New frmCrystalReportViewer()
                frmcrystal.funreport(MyBase.Form_ID, CrystalReportFolder.HRPayroll, DT, "crptBankSummaryAdvice", "Bank Summary Report")
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found", Me.Text)
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub PrintData(ByVal is_Print As Boolean)
        '' changed by Panch raj against Ticket No:BM00000007951
        Try
            If clsCommon.myLen(txtFromPP.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please Select Pay Period.", Me.Text)
                Return
            End If


            Dim Comp_Logo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Logo_Img from TSPL_COMPANY_MASTER where Comp_code='" + objCommonVar.CurrentCompanyCode + "'"))
            Dim Qry As String = ""
            Qry = ""
            Dim PM As Integer
            If rbtnPM.CheckState = CheckState.Checked Then
                PM = 1
            Else
                PM = 0
            End If
            Dim InnerQry As String = ""
            'InnerQry = " select '" + Comp_Logo + "' as Logo_Img, (case when " & PM & "=1 then PM.Name else T3.Bank_Name end) as Bank_Name,max(TSPL_LOCATION_MASTER.Add1+Case When ISNULL(TSPL_LOCATION_MASTER.Add2,'')='' Then '' else ', '+TSPL_LOCATION_MASTER.Add2+ Case When ISNULL(TSPL_LOCATION_MASTER.Add3,'')='' Then '' Else ', '+TSPL_LOCATION_MASTER.Add3+ Case When ISNULL(TSPL_STATE_MASTER.State_Name ,'')='' Then '' else '-'+CONVERT(varchar, TSPL_STATE_MASTER.State_Name) End End End ) as Loc_Add" & _
            '           " ,COUNT (T3.EMP_CODE) AS 'No of Emp',  SUM(T1.NET_SALARY) as 'sum amt' , max(T2.PAY_PERIOD_CODE) as PAY_PERIOD_CODE,  '" + objCommonVar.CurrentCompanyName + "' as Company_Name,max(TSPL_DEVISION_MASTER.DEVISION_CODE) as DEVISION_CODE,max(TSPL_DEVISION_MASTER.DEVISION_NAME)  as DEVISION_NAME ,max(Location_Desc)as Location_Desc ,max(T2.Location_Code) as Location_Code,'" + objCommonVar.CurrentCompanyCode + "' as Comp_Code " & _
            InnerQry = "select  '" + Comp_Logo + "'  as Logo_Img,max(TSPL_LOCATION_MASTER.Add1+Case When ISNULL(TSPL_LOCATION_MASTER.Add2,'')='' Then '' else ', '+TSPL_LOCATION_MASTER.Add2+ Case When ISNULL(TSPL_LOCATION_MASTER.Add3,'')='' Then '' Else ', '+TSPL_LOCATION_MASTER.Add3+ Case When ISNULL(TSPL_STATE_MASTER.State_Name ,'')='' Then '' else '-'+CONVERT(varchar, TSPL_STATE_MASTER.State_Name) End End End ) as Loc_Add,sum( case when TSPL_PAYHEAD_MASTER.isearning=1 then  TSPL_GENERATE_SALARY_PAYHEADS.ACTUAL_AMOUNT else -TSPL_GENERATE_SALARY_PAYHEADS.ACTUAL_AMOUNT end  )as 'sum amt', max(TSPL_GENERATE_SALARY.PAY_PERIOD_CODE) as pay_period, '" + objCommonVar.CurrentCompanyName + "'  as Company_Name,'" + objCommonVar.CurrentCompanyCode + "' as Comp_Code ,(TSPL_LOCATION_MASTER.Location_Code) as Location_Code, (case when " & PM & "=1 then isnull(PM.Name,'') else TSPL_EMPLOYEE_MASTER.Bank_Name end) as Bank_Name " &
             " from TSPL_GENERATE_SALARY_PAYHEADS left outer join TSPL_GENERATE_SALARY on TSPL_GENERATE_SALARY_PAYHEADS.SALARY_GENERATION_CODE =TSPL_GENERATE_SALARY.SALARY_GENERATION_CODE " &
             " left join TSPL_PAYHEAD_MASTER on TSPL_PAYHEAD_MASTER.Pay_HEAD_Code=TSPL_GENERATE_SALARY_PAYHEADS.Pay_HEAD_Code " &
             " left outer join TSPL_PAYPERIOD_MASTER on TSPL_PAYPERIOD_MASTER.PAY_PERIOD_CODE =TSPL_GENERATE_SALARY.PAY_PERIOD_CODE  " &
             " left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_GENERATE_SALARY_PAYHEADS.EMP_CODE " &
             " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code =TSPL_GENERATE_SALARY.LOCATION_CODE  " &
             " left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
             " left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER  .Comp_Code =TSPL_LOCATION_MASTER.Comp_code  " &
             " left join TSPL_DESIGNATION_MASTER on TSPL_DESIGNATION_MASTER .Designation_id =TSPL_EMPLOYEE_MASTER.Designation " &
             " left outer join TSPL_DEVISION_MASTER   on TSPL_DEVISION_MASTER .DEVISION_CODE  =TSPL_GENERATE_SALARY.DEVISION_CODE " &
             " left outer join TSPL_PAYMENT_MODE PM on TSPL_EMPLOYEE_MASTER.PAYMENT_MODE_New  = PM.Code WHERE " &
             " TSPL_GENERATE_SALARY.PAY_PERIOD_CODE = '" + txtFromPP.Value + "' "
            If txtLocationMult.arrValueMember IsNot Nothing AndAlso txtLocationMult.arrValueMember.Count > 0 Then
                InnerQry += " AND TSPL_LOCATION_MASTER.LOCATION_CODE  in (" + clsCommon.GetMulcallString(txtLocationMult.arrValueMember) + ") "
            End If
            If txtDivisionMult.arrValueMember IsNot Nothing AndAlso txtDivisionMult.arrValueMember.Count > 0 Then
                InnerQry += " AND TSPL_DEVISION_MASTER.Devision_Code  in (" + clsCommon.GetMulcallString(txtDivisionMult.arrValueMember) + ") "
            End If
            InnerQry += " GROUP BY (case when " & PM & "=1 then PM.Name else TSPL_EMPLOYEE_MASTER.Bank_Name end), TSPL_LOCATION_MASTER.Location_Code"


            If is_Print = True Then
                'Qry = "select * from (" & InnerQry & ") XXX WHERE 2=2"
                Qry = "(select ROW_NUMBER() Over(Order By (Select 1)) As 'SNo',* from ((" & InnerQry & ")as ll left join " &
                " (select count (TSPL_EMPLOYEE_MASTER.EMP_CODE) as No_of_employee ,Bank_Name  as Bank  from TSPL_EMPLOYEE_MASTER left join TSPL_GENERATE_SALARY_ATTENDANCE  on TSPL_GENERATE_SALARY_ATTENDANCE.EMP_CODE =TSPL_EMPLOYEE_MASTER.EMP_CODE left join TSPL_GENERATE_SALARY on TSPL_GENERATE_SALARY.SALARY_GENERATION_CODE = TSPL_GENERATE_SALARY_ATTENDANCE.SALARY_GENERATION_CODE  where  TSPL_GENERATE_SALARY.PAY_PERIOD_CODE = '" + txtFromPP.Value + "' "
                If txtLocationMult.arrValueMember IsNot Nothing AndAlso txtLocationMult.arrValueMember.Count > 0 Then
                    InnerQry += " and  TSPL_GENERATE_SALARY.LOCATION_CODE  in (" + clsCommon.GetMulcallString(txtLocationMult.arrValueMember) + ") "
                End If
                If txtDivisionMult.arrValueMember IsNot Nothing AndAlso txtDivisionMult.arrValueMember.Count > 0 Then
                    InnerQry += " AND TSPL_GENERATE_SALARY.Devision_Code  in (" + clsCommon.GetMulcallString(txtDivisionMult.arrValueMember) + ") "
                End If
                Qry += " group by Bank_Name)as mm on mm.Bank = ll.Bank_Name ))"
            Else
                Qry = "select Bank_Name as [Bank Name],No_of_employee as [No of Employees],[sum amt] as [Amount Transferred] from (select * from ((" & InnerQry & ")as ll left join " &
                " (select count (TSPL_EMPLOYEE_MASTER.EMP_CODE) as No_of_employee ,Bank_Name  as Bank  from TSPL_EMPLOYEE_MASTER left join TSPL_GENERATE_SALARY_ATTENDANCE  on TSPL_GENERATE_SALARY_ATTENDANCE.EMP_CODE =TSPL_EMPLOYEE_MASTER.EMP_CODE left join TSPL_GENERATE_SALARY on TSPL_GENERATE_SALARY.SALARY_GENERATION_CODE = TSPL_GENERATE_SALARY_ATTENDANCE.SALARY_GENERATION_CODE  where  TSPL_GENERATE_SALARY.PAY_PERIOD_CODE = '" + txtFromPP.Value + "' "
                If txtLocationMult.arrValueMember IsNot Nothing AndAlso txtLocationMult.arrValueMember.Count > 0 Then
                    InnerQry += " and  TSPL_GENERATE_SALARY.LOCATION_CODE  in (" + clsCommon.GetMulcallString(txtLocationMult.arrValueMember) + ") "
                End If
                If txtDivisionMult.arrValueMember IsNot Nothing AndAlso txtDivisionMult.arrValueMember.Count > 0 Then
                    InnerQry += " AND TSPL_GENERATE_SALARY.Devision_Code  in (" + clsCommon.GetMulcallString(txtDivisionMult.arrValueMember) + ") "
                End If
                Qry += " group by Bank_Name)as mm on mm.Bank = ll.Bank_Name )) as XXX WHERE 2=2"
            End If

            If txtmultBank.arrValueMember IsNot Nothing AndAlso txtmultBank.arrValueMember.Count > 0 Then
                Qry += " AND XXX.Bank_Name  in (" + clsCommon.GetMulcallString(txtmultBank.arrValueMember) + ")  "

            End If


            'Qry += "and Bank_Name <> '' "
            Dim DT As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If DT IsNot Nothing AndAlso DT.Rows.Count > 0 Then
                If Not is_Print Then
                    Gv1.DataSource = Nothing
                    Gv1.DataSource = DT
                    SetGridFormationOFGV1()
                    ReStoreGridLayout()
                    RadPageView1.SelectedPage = RadPageViewPage2
                Else
                    'formatgrid()
                    Dim frmcrystal As New frmCrystalReportViewer()
                    frmcrystal.funreport(MyBase.Form_ID, CrystalReportFolder.HRPayroll, DT, "crptBankSummary", "Bank Summary Report")
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub formatgrid()

        Gv1.TableElement.TableHeaderHeight = 25
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = False

        Next
        Gv1.Columns("Logo_Img").IsVisible = False
        Gv1.Columns("Logo_Img").Width = 100
        Gv1.Columns("Logo_Img").HeaderText = "Image"

        Gv1.Columns("Loc_Add").IsVisible = True
        Gv1.Columns("Loc_Add").Width = 100
        Gv1.Columns("Loc_Add").HeaderText = "Location Address"

        Gv1.Columns("pay_period").IsVisible = True
        Gv1.Columns("pay_period").Width = 100
        Gv1.Columns("pay_period").HeaderText = "Pay Period"

        Gv1.Columns("Company_Name").IsVisible = False
        Gv1.Columns("Company_Name").Width = 100
        Gv1.Columns("Company_Name").HeaderText = "Company Name"

        Gv1.Columns("Comp_Code").IsVisible = False
        Gv1.Columns("Comp_Code").Width = 100
        Gv1.Columns("Comp_Code").HeaderText = "Comp_Code"

        Gv1.Columns("Location_Code").IsVisible = True
        Gv1.Columns("Location_Code").Width = 100
        Gv1.Columns("Location_Code").HeaderText = "Location Code"

        Gv1.Columns("Bank_Name").IsVisible = True
        Gv1.Columns("Bank_Name").Width = 100
        Gv1.Columns("Bank_Name").HeaderText = "Bank Name"

        Gv1.Columns("sum amt").IsVisible = True
        Gv1.Columns("sum amt").Width = 100
        Gv1.Columns("sum amt").HeaderText = "Net Amt"

        'Gv1.Columns("No of Emp").IsVisible = True
        'Gv1.Columns("No of Emp").Width = 100
        'Gv1.Columns("No of Emp").HeaderText = "No Of Employee"


        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim item1 As New GridViewSummaryItem("Amount Transferred", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item1)

        Dim item2 As New GridViewSummaryItem("No of Employees", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item2)
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        RadPageView1.SelectedPage = RadPageViewPage2
        Gv1.AllowAddNewRow = False
        Gv1.ShowGroupPanel = True
        Gv1.ReadOnly = True
        Gv1.EnableFiltering = True
        Gv1.EnableGrouping = True
        Gv1.BestFitColumns()
    End Sub

    Sub SetGridFormationOFGV1()




        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim item1 As New GridViewSummaryItem("Amount Transferred", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item1)

        Dim item2 As New GridViewSummaryItem("No of Employees", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item2)
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        RadPageView1.SelectedPage = RadPageViewPage2
        Gv1.AllowAddNewRow = False
        Gv1.ShowGroupPanel = True
        Gv1.ReadOnly = True
        Gv1.EnableFiltering = True
        Gv1.EnableGrouping = True
        Gv1.BestFitColumns()
    End Sub

    Private Sub txtLocationMult__My_Click(sender As Object, e As EventArgs) Handles txtLocationMult._My_Click
        Dim qry As String = " select Location_Code as Code,Location_Desc as [Name] from TSPL_LOCATION_MASTER where LOCATION_CODE IN (select DISTINCT LOCATION_CODE from TSPL_GENERATE_SALARY where PAY_PERIOD_CODE='" & txtFromPP.Value & "') "
        txtLocationMult.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "Code", "Name", txtLocationMult.arrValueMember, txtLocationMult.arrDispalyMember)
    End Sub

    Private Sub txtDivisionMult__My_Click(sender As Object, e As EventArgs) Handles txtDivisionMult._My_Click
        Dim qry As String = " select DEVISION_CODE as Code,DEVISION_NAME as Name from TSPL_DEVISION_MASTER"
        txtDivisionMult.arrValueMember = clsCommon.ShowMultipleSelectForm("DivMulSel", qry, "Code", "Name", txtDivisionMult.arrValueMember, txtDivisionMult.arrDispalyMember)
    End Sub

    Private Sub txtmultBank__My_Click(sender As Object, e As EventArgs) Handles txtmultBank._My_Click
        Qry = ""
        'Qry += " SELECT DISTINCT TSPL_EMPLOYEE_MASTER .BANK_CODE as Code ,TSPL_BANK_MASTER.DESCRIPTION AS Name "
        'Qry += " FROM TSPL_GENERATE_SALARY "
        'Qry += " LEFT OUTER JOIN TSPL_GENERATE_SALARY_ATTENDANCE  ON TSPL_GENERATE_SALARY_ATTENDANCE  .SALARY_GENERATION_CODE = TSPL_GENERATE_SALARY .SALARY_GENERATION_CODE  left outer join TSPL_EMPLOYEE_MASTER on TSPL_GENERATE_SALARY_ATTENDANCE .EMP_CODE= TSPL_EMPLOYEE_MASTER .EMP_CODE  "
        'Qry += " left outer join TSPL_BANK_MASTER  on TSPL_BANK_MASTER .BANK_CODE = TSPL_EMPLOYEE_MASTER .BANK_CODE "
        'Qry += " Left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code= TSPL_GENERATE_SALARY.Location_Code"
        'Qry += " where TSPL_GENERATE_SALARY.PAY_PERIOD_CODE ='" + txtFromPP.Value + "' "
        If rbtnPM.CheckState = CheckState.Checked Then
            Qry = "select Code , Name  from TSPL_PAYMENT_MODE "
        Else
            Qry = " SELECT DISTINCT TSPL_EMPLOYEE_MASTER .Bank_Name  as Code  FROM TSPL_GENERATE_SALARY  " &
             " LEFT OUTER JOIN TSPL_GENERATE_SALARY_ATTENDANCE  ON TSPL_GENERATE_SALARY_ATTENDANCE  .SALARY_GENERATION_CODE = TSPL_GENERATE_SALARY .SALARY_GENERATION_CODE  " &
             " left outer join TSPL_EMPLOYEE_MASTER on TSPL_GENERATE_SALARY_ATTENDANCE .EMP_CODE= TSPL_EMPLOYEE_MASTER .EMP_CODE   " &
             " Left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code= TSPL_GENERATE_SALARY.Location_Code " &
             " where TSPL_GENERATE_SALARY.PAY_PERIOD_CODE ='" + txtFromPP.Value + "' and Bank_Name <> '' "
        End If

        txtmultBank.arrValueMember = clsCommon.ShowMultipleSelectForm("BankMulSel", Qry, "Code", "Code", txtmultBank.arrValueMember, txtmultBank.arrDispalyMember)
    End Sub


    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        PageSetupReport_ID = MyBase.Form_ID
        TemplateGridview = Gv1
        GetReportID()
        PrintData(False)
    End Sub

    Sub GetReportID()
        Dim VarID As String = ""
        If rbtnPM.IsChecked Then
            VarID += "_P"
        ElseIf rbtnBankwise.isChecked Then
            VarID += "_B"
        End If
        Gv1.VarID = VarID
    End Sub

    Private Sub ExportGrid(ByVal exporter As EnumExportTo)
        Try
            If Gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                Dim strTemp As String = ""
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.frmBankSummary_Report & "'"))
                arrHeader.Add("Pay Period : " & txtFromPP.Value)
                If Not txtLocationMult.arrValueMember Is Nothing Then
                    arrHeader.Add("Location : " & clsCommon.GetMulcallStringWithComma(txtLocationMult.arrValueMember))
                Else
                    arrHeader.Add("Location : All")
                End If
                If Not txtDivisionMult.arrValueMember Is Nothing Then
                    arrHeader.Add("Division : " & clsCommon.GetMulcallStringWithComma(txtDivisionMult.arrValueMember))
                Else
                    arrHeader.Add("Division : All")
                End If
                If exporter = EnumExportTo.Excel Then
                    'Dim sfd As SaveFileDialog = New SaveFileDialog()
                    'Dim filePath As String
                    'sfd.FileName = Me.Text
                    'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
                    'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    '    filePath = sfd.FileName
                    'Else
                    '    Exit Sub
                    'End If
                    'transportSql.exportdata(Gv1, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1)) 'frm.Text)
                    transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
                    transportSql.QuickExportToExcel(Gv1, "", Me.Text, , arrHeader)
                    'transportSql.exportdataChilRows(Gv1, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), , arrHeader)
                    'common.clsCommon.MyMessageBoxShow("Exported Successfully.")
                    'Process.Start(filePath)
                Else
                    transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
                    clsCommon.MyExportToPDF(Me.Text, Gv1, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
                End If
            Else
                common.clsCommon.MyMessageBoxShow(Me, "No Data Found to Export.", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        ExportGrid(EnumExportTo.Excel)
    End Sub

    Private Sub rmiPDF_Click(sender As Object, e As EventArgs) Handles rmiPDF.Click
        ExportGrid(EnumExportTo.PDF)
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(PageSetupReport_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= Gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To Gv1.Columns.Count - 1 Step ii + 1
                        Gv1.Columns(ii).IsVisible = False
                        Gv1.Columns(ii).VisibleInColumnChooser = True
                    Next

                    Gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub rmSaveLayout_Click(sender As Object, e As EventArgs) Handles rmSaveLayout.Click
        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = PageSetupReport_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            Gv1.SaveLayout(obj.GridLayout)
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            obj.GridColumns = Gv1.ColumnCount
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub
End Class
