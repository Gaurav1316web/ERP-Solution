
Imports common
Imports System.IO
Public Class rptFarmerMilkCollectionReport
    Inherits FrmMainTranScreen

    Dim Slot1 As DateTime = Nothing
    Dim Slot2 As DateTime = Nothing
    Dim Month1 As String = Nothing
    Dim Month2 As String = Nothing
    Dim Month3 As String = Nothing
    Dim Month4 As String = Nothing
    Dim Month5 As String = Nothing
    Dim Month6 As String = Nothing
    Dim Month7 As String = Nothing
    Dim Month8 As String = Nothing
    Dim Month9 As String = Nothing
    Dim Month10 As String = Nothing
    Dim Month11 As String = Nothing
    Dim Month12 As String = Nothing
    Private Sub rptFarmerMilkCollectionReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        funreset()
        txtFromMonth.Value = clsCommon.GETSERVERDATE()
        TxtToMonth.Value = clsCommon.GETSERVERDATE().AddMonths(-1)

        If clsCommon.myLen(objCommonVar.CurrentUnionDataBase) > 0 Then
            Dim Union As ArrayList = Nothing
            Dim qry As String = " Select DataBase_Name from TSPL_USER_MASTER where User_Code = '" + objCommonVar.CurrentUserCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Union = New ArrayList()
                For Each drZone As DataRow In dt.Rows
                    Union.Add(clsCommon.myCstr(drZone("DataBase_Name")))
                Next
            End If
            txtUnion.arrValueMember = Union
        End If
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        funreset()
    End Sub

    Sub funreset()
        gv1.DataSource = Nothing
        txtUnion.arrValueMember = Nothing
        txtFromMonth.Value = clsCommon.GETSERVERDATE()
        TxtToMonth.Value = clsCommon.GETSERVERDATE().AddMonths(-1)
        RadPageView1.SelectedPage = RadPageViewPage1
        EnableDisableCtrl(True)
    End Sub

    Private Sub txtUnion__My_Click(sender As Object, e As EventArgs) Handles txtUnion._My_Click
        Try
            Dim dt As DataTable
            Dim qry As String = ""

            If clsCommon.myLen(objCommonVar.CurrentUnionDataBase) > 0 Then
                qry = " Select DataBase_Name as [DataBase Name] from TSPL_USER_MASTER where User_Code = '" + objCommonVar.CurrentUserCode + "' "
                txtUnion.arrValueMember = clsCommon.ShowMultipleSelectForm("SaleUnionDs", qry, "DataBase Name", "", txtUnion.arrValueMember, Nothing)

            Else
                dt = clsDBFuncationality.GetDataTable("SELECT name FROM master.dbo.sysdatabases  WHERE name = 'TSPL_MASTER'")
                If (dt Is Nothing OrElse dt.Rows.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Database[TSPL_MASTER] not found")
                    Exit Sub
                End If

                qry = "SELECT [TSPL_APP_LOCATION].Location_Name as Location,[TSPL_APP_LOCATION].DataBase_Name as [DataBase Name] FROM [TSPL_MASTER].[dbo].[TSPL_APP_LOCATION] WHERE Union_Report=1 ORDER BY [TSPL_APP_LOCATION].Location_Name"

                txtUnion.arrValueMember = clsCommon.ShowMultipleSelectForm("DBTUnionPay", qry, "DataBase Name", "", txtUnion.arrValueMember, Nothing)

            End If
            'dt = clsDBFuncationality.GetDataTable("SELECT name FROM master.dbo.sysdatabases  WHERE name = 'TSPL_MASTER'")
            'If (dt Is Nothing OrElse dt.Rows.Count <= 0) Then
            '    common.clsCommon.MyMessageBoxShow(Me, "Database[TSPL_MASTER] not found")
            '    Exit Sub
            'End If

            'qry = "SELECT [TSPL_APP_LOCATION].Location_Name as Location,[TSPL_APP_LOCATION].DataBase_Name as [DataBase Name] FROM [TSPL_MASTER].[dbo].[TSPL_APP_LOCATION] WHERE Union_Report=1 ORDER BY [TSPL_APP_LOCATION].Location_Name"

            'txtUnion.arrValueMember = clsCommon.ShowMultipleSelectForm("DBTUnionPay", qry, "DataBase Name", "", txtUnion.arrValueMember, Nothing)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim SM As Integer = txtFromMonth.Value.Month
            Dim SY As Integer = txtFromMonth.Value.Year

            Dim CD As New DateTime(SY, SM, 1)

            Dim SM1 As Integer = TxtToMonth.Value.Month
            Dim SY1 As Integer = TxtToMonth.Value.Year
            'Last day of To Month
            Dim LastDay As Integer = DateTime.DaysInMonth(SY1, SM1)

            Dim CD1 As New DateTime(SY1, SM1, LastDay)
            Slot1 = clsCommon.GetPrintDate(CD, "dd/MMM/yyyy")
            Slot2 = clsCommon.GetPrintDate(CD1, "dd/MMM/yyyy")

            Dim qry As String = ""
            Dim qryies As String = ""
            Dim Baseqry As String = ""
            Dim Baseqry1 As String = ""
            Dim Baseqry2 As String = ""
            Dim dbNames As String = ""
            Dim portDt As New DataTable
            Dim dtGrandTotal As DataTable
            Dim FromDate As Date = txtFromMonth.Value
            Dim ToDate As Date = TxtToMonth.Value

            Dim MonthCount As Integer = ((ToDate.Year - FromDate.Year) * 12) + (ToDate.Month - FromDate.Month)

            Dim CurrentDate As Date = New Date(FromDate.Year, FromDate.Month, 1)
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("SELECT name FROM master.dbo.sysdatabases  WHERE name = 'TSPL_MASTER'")
            If (dt Is Nothing OrElse dt.Rows.Count <= 0) Then
                common.clsCommon.MyMessageBoxShow(Me, "Database[TSPL_MASTER] not found", Me.Text)
                gv1.DataSource = Nothing
                Exit Sub
            End If
            Dim ss As String = clsCommon.GetMulcallString(txtUnion.arrValueMember)

            Dim arrUnion As New ArrayList()
            arrUnion.Add(objCommonVar.CurrComp_Code1)
            If clsCommon.myLen(objCommonVar.CurrentUnionDataBase) > 0 Then
                qry = " Select TSPL_USER_MASTER.DataBase_Name,[TSPL_APP_LOCATION].Location_Name from TSPL_USER_MASTER 
                    left outer join TSPL_MASTER.dbo.[TSPL_APP_LOCATION] on [TSPL_APP_LOCATION].DataBase_Name=TSPL_USER_MASTER.DataBase_Name where User_Code = '" + objCommonVar.CurrentUserCode + "' "
                dt = clsDBFuncationality.GetDataTable(qry)
                'txtUnion.arrValueMember = clsCommon.ShowMultipleSelectForm("SaleUnionDs", Qry, "DataBase Name", "", txtUnion.arrValueMember, Nothing)
            Else
                If objCommonVar.RCDFCFP Then
                    dt = clsMilkUnion.UnionDBName()
                Else
                    dt = clsMilkUnion.UnionDBName1(arrUnion)
                End If
            End If

            'If txtUnion.arrValueMember Is Nothing Then
            '    qry = " select  [TSPL_APP_LOCATION].PD_Account_Prefix as PortNo,[TSPL_APP_LOCATION].Location_Name,[TSPL_APP_LOCATION].DataBase_Name
            '                from TSPL_MASTER.dbo.TSPL_APP_LOCATION WHERE 2=2 order by [TSPL_APP_LOCATION].Location_Name "
            'Else
            '    qry = " select  [TSPL_APP_LOCATION].PD_Account_Prefix as PortNo,[TSPL_APP_LOCATION].Location_Name,[TSPL_APP_LOCATION].DataBase_Name
            '            from TSPL_MASTER.dbo.TSPL_APP_LOCATION WHERE 2=2 and [TSPL_APP_LOCATION].DataBase_Name  in (" + ss + ") 
            '            order by [TSPL_APP_LOCATION].Location_Name "
            'End If
            ''dt = clsMilkUnion.UnionDBName()
            'dt = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                For ii As Integer = 0 To dt.Rows.Count - 1
                    If ii > 0 Then
                        qryies += " UNION ALL "
                    Else
                        qryies += " WITH Data AS ( "
                    End If

                    qryies += " 
    SELECT  '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,
        V.VLC_Code,M.MP_CODE,Doc_Date,Qty,0 as RecoQty
    FROM [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_DATA_UPLOADER U
    LEFT JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_MASTER_HEAD V
        ON V.Vlc_Code_VLC_Uploader = U.VLC_CODE
    LEFT JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPl_MP_MAster M
        ON M.MP_Code_VLC_Uploader = U.MP_CODE
       AND M.VLC_Code = V.VLC_Code
    WHERE CONVERT(DATE,Doc_Date)
          BETWEEN '" + clsCommon.GetPrintDate(Slot1) + "' AND '" + clsCommon.GetPrintDate(Slot2) + "'

    UNION ALL

    SELECT  '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,
        UM.VLC_CODE,M.MP_CODE,UM.Document_Date,D.Qty,0 as RecoQty
    FROM [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_DATA_UPLOADER_DETAIL D
    INNER JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_DATA_UPLOADER_MASTER UM
        ON UM.Document_Code = D.Document_Code
    LEFT JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPl_MP_MAster M
        ON M.MP_Code = D.Farmer_Code
    WHERE CONVERT(DATE,UM.Document_Date)
          BETWEEN '" + clsCommon.GetPrintDate(Slot1) + "' AND '" + clsCommon.GetPrintDate(Slot2) + "'

union all

select  '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code, NULL AS MP_CODE, TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date as Document_Date,
0 AS qty  ,TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Qty as RecoQty
from [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL
Left Outer Join [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD On [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code
Left Join [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_MASTER_HEAD On [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_MASTER_HEAD.VLC_Code=[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code
Where 1=1 and convert(date,[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) >='" + clsCommon.GetPrintDate(Slot1) + "' And convert(date,[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103)<= '" + clsCommon.GetPrintDate(Slot2) + "'

union all

select  '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.VLC_Code,NULL AS MP_CODE,  TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date as Document_Date,
0 AS qty  ,TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.Qty as RecoQty
from [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID
Left Outer Join [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD On [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.Document_Code
Left Join [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_MASTER_HEAD On [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_VLC_MASTER_HEAD.VLC_Code=[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.VLC_Code
Where 1=1 and convert(date,[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) >='" + clsCommon.GetPrintDate(Slot1) + "' And convert(date,[" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103)<= '" + clsCommon.GetPrintDate(Slot2) + "'

"
                Next

            End If

            qryies += " ), "

            '           For ii As Integer = 0 To dt.Rows.Count - 1
            '               If ii > 0 Then
            '                   qryies += " UNION ALL "
            '               Else
            '                   qryies += " RecoData AS ( "
            '               End If

            '               qryies += "  SELECT UnionName,
            '       FORMAT(Reco_Date,'MMM-yy') AS DateMonth,
            '       SUM(Qty) AS RecoQty
            '   FROM
            '   (
            '       SELECT '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,
            '           H.Reco_Date,
            '           D.Qty
            '       FROM [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL D
            '       INNER JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD H
            '           ON H.Document_Code = D.Document_Code
            '       WHERE H.Reco_Date >= '" + clsCommon.GetPrintDate(Slot1) + "'
            '         AND H.Reco_Date <= '" + clsCommon.GetPrintDate(Slot2) + "'

            '       UNION ALL

            '       SELECT '" + clsCommon.myCstr(dt.Rows(ii).Item("Location_Name")) + "' AS UnionName,
            '           H.Reco_Date,
            '           D.Qty
            '       FROM [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID D
            '       INNER JOIN [" + clsCommon.myCstr(dt.Rows(ii).Item("DataBase_Name")) + "].[dbo].TSPL_DCS_MP_INCENTIVE_RECO_HEAD H
            '           ON H.Document_Code = D.Document_Code
            '       WHERE H.Reco_Date >= '" + clsCommon.GetPrintDate(Slot1) + "'
            '         AND H.Reco_Date <= '" + clsCommon.GetPrintDate(Slot2) + "'
            '   ) X
            '   GROUP BY FORMAT(Reco_Date,'MMM-yy'),UnionName
            '"
            '           Next
            '           qryies += " ), "
            qryies += " MonthlySummary AS  
(
    Select  Data.UnionName,FORMAT(Doc_Date,'MMM-yy') AS DateMonth,COUNT(DISTINCT VLC_Code) AS DCS,
        count(DISTINCT MP_CODE) As Farmer,SUM(Qty) As Qty,SUM(Qty)*5 As Amount,SUM(RecoQty) AS RecoQty,SUM(RecoQty)*5 AS RecoAmt FROM Data
    Group BY   Data.UnionName, Year(Doc_Date), Month(Doc_Date), Format(Doc_Date,'MMM-yy')
)

Select 
     ROW_NUMBER() OVER(ORDER BY UnionName) As SNo,UnionName "

                For i As Integer = 0 To MonthCount
                Dim MonthCode As String = CurrentDate.ToString("MMM-yy")

                qryies += "  ,MAX(Case When DateMonth='" + MonthCode + "' THEN DCS END)     AS [" + MonthCode + "_DCS],
    MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN Farmer END)  AS [" + MonthCode + "_Farmer],
    MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + MonthCode + "_Qty],
    MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (CAST(RecoQty AS DECIMAL(18,2))) END)     AS [" + MonthCode + "_RecoQty],
    MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + MonthCode + "_Amount],
    MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (Cast (RecoAmt AS DECIMAL(18,2))) END)  AS [" + MonthCode + "_RecoAmount]"

                CurrentDate = CurrentDate.AddMonths(1)
            Next
            qryies += " ,SUM(CAST(Qty AS DECIMAL(18,2))) AS Total_Qty,sum(cast(RecoQty as decimal(18,2))) as Total_RecoQty
                        ,sum(cast(RecoAmt as decimal(18,2))) as Total_RecoAmount,SUM(CAST(Amount AS DECIMAL(18,2))) AS Total_Amount
                        ,Sum(Cast(RecoAmt-Amount as decimal(18,2))) as Difference,Sum(Cast((Qty *100)/NULLIF(RecoQty, 0) as decimal(18,2)) ) as Percentage"
            '        For i As Integer = 0 To MonthCount
            '            Dim MonthCode As String = CurrentDate.ToString("MMM-yy")

            '            qryies += "  ,ISNULL(MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN CAST(Qty AS DECIMAL(18,2)) END),0)     AS [" + MonthCode + "_DCS],
            'MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN Farmer END)  AS [" + MonthCode + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + MonthCode + "_Qty],
            'MAX(CASE WHEN DateMonth='" + MonthCode + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + MonthCode + "_Amount]"

            '            CurrentDate = CurrentDate.AddMonths(1)
            '        Next

            '        qryies += " MAX(CASE WHEN DateMonth='" + Month1 + "' THEN DCS END)     AS [" + Month1 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month1 + "' THEN Farmer END)  AS [" + Month1 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month1 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month1 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month1 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month1 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month2 + "' THEN DCS END)     AS [" + Month2 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month2 + "' THEN Farmer END)  AS [" + Month2 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month2 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month2 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month2 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month2 + "_Amount],   

            'MAX(CASE WHEN DateMonth='" + Month3 + "' THEN DCS END)     AS [" + Month3 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month3 + "' THEN Farmer END)  AS [" + Month3 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month3 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month3 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month3 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month3 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month4 + "' THEN DCS END)     AS [" + Month4 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month4 + "' THEN Farmer END)  AS [" + Month4 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month4 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month4 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month4 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month4 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month5 + "' THEN DCS END)     AS [" + Month5 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month5 + "' THEN Farmer END)  AS [" + Month5 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month5 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month5 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month5 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month5 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month6 + "' THEN DCS END)     AS [" + Month6 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month6 + "' THEN Farmer END)  AS [" + Month6 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month6 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month6 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month6 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month6 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month7 + "' THEN DCS END)     AS [" + Month7 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month7 + "' THEN Farmer END)  AS [" + Month7 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month7 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month7 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month7 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END) AS [" + Month7 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month8 + "' THEN DCS END)     AS [" + Month8 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month8 + "' THEN Farmer END)  AS [" + Month8 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month8 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month8 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month8 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month8 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month9 + "' THEN DCS END)     AS [" + Month9 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month9 + "' THEN Farmer END)  AS [" + Month9 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month9 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month9 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month9 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month9 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month10 + "' THEN DCS END)     AS [" + Month10 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month10 + "' THEN Farmer END)  AS [" + Month10 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month10 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)     AS [" + Month10 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month10 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month10 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month11 + "' THEN DCS END)     AS [" + Month11 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month11 + "' THEN Farmer END)  AS [" + Month11 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month11 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)    AS [" + Month11 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month11 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month11 + "_Amount],

            'MAX(CASE WHEN DateMonth='" + Month12 + "' THEN DCS END)     AS [" + Month12 + "_DCS],
            'MAX(CASE WHEN DateMonth='" + Month12 + "' THEN Farmer END)  AS [" + Month12 + "_Farmer],
            'MAX(CASE WHEN DateMonth='" + Month12 + "' THEN (CAST(Qty AS DECIMAL(18,2))) END)      AS [" + Month12 + "_Qty],
            'MAX(CASE WHEN DateMonth='" + Month12 + "' THEN (Cast (Amount AS DECIMAL(18,2))) END)  AS [" + Month12 + "_Amount] "

            qryies += "  FROM MonthlySummary
GROUP BY UnionName
ORDER BY UnionName;
"
            portDt = clsDBFuncationality.GetDataTable(qryies)
            gv1.DataSource = Nothing
            gv1.Rows.Clear()
            gv1.Columns.Clear()
            gv1.GroupDescriptors.Clear()
            gv1.MasterView.Refresh()
            gv1.GroupDescriptors.Clear()
            gv1.EnableFiltering = True
            gv1.EnableFiltering = False
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            If portDt.Rows.Count > 0 Then
                gv1.DataSource = portDt
                gv1.BestFitColumns()
                'SetGridFormation()
                'ReStoreGridLayout()
                gv1.MasterTemplate.AutoExpandGroups = True
                RadPageView1.SelectedPage = RadPageViewPage2
                gv1.BestFitColumns()
                EnableDisableCtrl(False)
                View()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub View()
        If gv1.Rows.Count > 0 Then

            Dim FromDate As Date = txtFromMonth.Value
            Dim ToDate As Date = TxtToMonth.Value

            Dim MonthCount As Integer = ((ToDate.Year - FromDate.Year) * 12) + (ToDate.Month - FromDate.Month)

            Dim CurrentDate As Date = New Date(FromDate.Year, FromDate.Month, 1)
            Dim view As New ColumnGroupsViewDefinition()

            view.ColumnGroups.Add(New GridViewColumnGroup("Union"))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("SNo").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("UnionName").Name)

            For i As Integer = 0 To MonthCount

                Dim MonthCode As String = CurrentDate.ToString("MMM-yy")

                'view.ColumnGroups.Add(New GridViewColumnGroup(MonthCode))
                Dim groupIndex As Integer = view.ColumnGroups.Count
                Dim ColDCS As String = MonthCode & "_DCS"
                Dim ColFarmer As String = MonthCode & "_Farmer"
                Dim ColQty As String = MonthCode & "_Qty"
                Dim ColRecoQty As String = MonthCode & "_RecoQty"
                Dim ColDBTAmt As String = MonthCode & "_Amount"
                Dim ColRecoAmt As String = MonthCode & "_RecoAmount"

                gv1.Columns(ColDCS).HeaderText = "DCS"
                gv1.Columns(ColFarmer).HeaderText = "Farmer Count"
                gv1.Columns(ColQty).HeaderText = "Farmer Qty"
                gv1.Columns(ColQty).HeaderText = "Reco Qty"
                gv1.Columns(ColDBTAmt).HeaderText = "DBT Amount"
                gv1.Columns(ColRecoAmt).HeaderText = "Reco Amount"

                'view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.GetPrintDate(MonthCode & " ", "MMM-yy")))
                view.ColumnGroups.Add(New GridViewColumnGroup("[" & clsCommon.GetPrintDate(MonthCode & " ", "MMM-yy") & "]"))
                view.ColumnGroups(groupIndex).Rows.Add(New GridViewColumnGroupRow())

                'view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns(MonthCode & "_DCS").Name)
                'view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns(MonthCode & "_Farmer").Name)
                'view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns(MonthCode & "_Qty").Name)
                'view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns(MonthCode & "_Amount").Name)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColDCS)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColFarmer)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColQty)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColRecoQty)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColDBTAmt)
                view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(ColRecoAmt)

                CurrentDate = CurrentDate.AddMonths(1)

            Next
            Dim groupIndex1 As Integer = view.ColumnGroups.Count
            Dim ColTotalQty As String = "Total_Qty"
            Dim ColTotalRecoQty As String = "Total_RecoQty"
            Dim ColTotalDBTAmt As String = "Total_Amount"
            Dim ColTotalRecoAmt As String = "Total_RecoAmount"
            Dim ColDifference As String = "Difference"
            Dim ColPercentage As String = "Percentage"

            gv1.Columns(ColTotalQty).HeaderText = "Total Qty"
            gv1.Columns(ColTotalRecoQty).HeaderText = "Total RecoQty"
            gv1.Columns(ColTotalDBTAmt).HeaderText = "Total Amount"
            gv1.Columns(ColTotalRecoAmt).HeaderText = "Total RecoAmount"
            gv1.Columns(ColDifference).HeaderText = "Difference"
            gv1.Columns(ColPercentage).HeaderText = "Percentage"

            view.ColumnGroups.Add(New GridViewColumnGroup("Total"))
            view.ColumnGroups(groupIndex1).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColTotalQty)
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColTotalRecoQty)
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColTotalDBTAmt)
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColTotalRecoAmt)
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColDifference)
            view.ColumnGroups(groupIndex1).Rows(0).ColumnNames.Add(ColPercentage)

            '           Dim view As New ColumnGroupsViewDefinition()
            '           view.ColumnGroups.Add(New GridViewColumnGroup(" "))
            '           view.ColumnGroups.Add(New GridViewColumnGroup(" "))
            '           view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
            '           view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Union Name").Name)
            '           'view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Item_Code").Name)
            '           'View.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Item_Desc").Name)
            '           'view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("UOM").Name)

            '           Dim qry As String = "  select DISTINCT DATEMONTH FROM (SELECT FORMAT(max(DOC_DATE), 'MMM-yy') as DateMonth
            '   FROM TSPL_VLC_DATA_UPLOADER     Where 1=1 
            'and convert(date,TSPL_VLC_DATA_UPLOADER.doc_date,103) >=convert(date,'" + clsCommon.GetPrintDate(Slot1) + "',103) 
            'And convert(date,TSPL_VLC_DATA_UPLOADER.doc_date,103)<= convert(date,'" + clsCommon.GetPrintDate(Slot2) + "',103)
            'group by doc_date)XX ORDER BY DATEMONTH DESC"
            '           Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            '           dt.Rows.Add("Jan-26")
            '           dt.Rows.Add("Feb-26")

            '           If dt.Rows.Count > 0 Then
            '               For Each row As DataRow In dt.Rows

            '                   Dim MonthCode As String = row("DATEMONTH").ToString().Trim().ToUpper()

            '                   ' Add group
            '                   view.ColumnGroups.Add(New GridViewColumnGroup(MonthCode))
            '                   Dim groupIndex As Integer = view.ColumnGroups.Count - 1

            '                   ' Add row
            '                   view.ColumnGroups(groupIndex).Rows.Add(New GridViewColumnGroupRow())
            '                   view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns("DCS").Name)
            '                   view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns("CountFarmer").Name)
            '                   view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns("Qty").Name)
            '                   view.ColumnGroups(groupIndex).Rows(0).ColumnNames.Add(gv1.Columns("Amount").Name)
            '               Next
            '           End If

            gv1.ViewDefinition = view
        End If
    End Sub

    Sub EnableDisableCtrl(ByVal val As Boolean)
        RadGroupBox1.Enabled = val
    End Sub
    Sub SetGridFormat()
        gv1.AutoExpandGroups = True
        'gv1.ShowGroupPanel = True
        gv1.ShowGroupPanel = False
        gv1.ShowRowHeaderColumn = False
        gv1.AllowAddNewRow = False
        gv1.AllowDeleteRow = False
        gv1.EnableFiltering = True
        gv1.ShowFilteringRow = True
        'gv1.ShowGroupPanel = False

        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).BestFit()
        Next
    End Sub
    Private Sub txtFromMonth_ValueChanged(sender As Object, e As EventArgs) Handles txtFromMonth.ValueChanged
        Try
            Dim SM As Integer = txtFromMonth.Value.Month
            Dim SY As Integer = txtFromMonth.Value.Year

            Dim CD As New DateTime(SY, SM, 1)
            Slot1 = clsCommon.GetPrintDate(CD, "dd/MMM/yyyy")
            Month()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Month()
        If clsCommon.myLen(txtFromMonth.Value) > 0 Then
            Dim SM As Integer = txtFromMonth.Value.Month
            Dim SY As Integer = txtFromMonth.Value.Year

            Dim CD As New DateTime(SY, SM, 1)
            'If rbtnYearly.IsChecked Then
            Slot2 = clsCommon.GetPrintDate(CD.AddMonths(12).AddDays(-1), "dd/MMM/yyyy")
            TxtToMonth.Value = txtFromMonth.Value.AddMonths(11)
            Month4 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(3), "MMM-yy")
            Month5 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(4), "MMM-yy")
            Month6 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(5), "MMM-yy")
            Month7 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(6), "MMM-yy")
            Month8 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(7), "MMM-yy")
            Month9 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(8), "MMM-yy")
            Month10 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(9), "MMM-yy")
            Month11 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(10), "MMM-yy")
            Month12 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(11), "MMM-yy")
            'ElseIf rbtnQuarterly.IsChecked Then
            '    Slot2 = clsCommon.GetPrintDate(CD.AddMonths(3).AddDays(-1), "dd/MMM/yyyy")
            '    TxtToMonth.Value = txtFromMonth.Value.AddMonths(2)
            'End If
            Month1 = clsCommon.GetPrintDate(txtFromMonth.Value, "MMM-yy")
            Month2 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(1), "MMM-yy")
            Month3 = clsCommon.GetPrintDate(txtFromMonth.Value.AddMonths(2), "MMM-yy")


        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ExportGrid(ByVal exporter As EnumExportTo)
        Try
            If gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            Dim strHeading As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptFarmerMilkCollectionReport & "'"))

            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            arrHeader.Add("Report Name : " + strHeading)
            If exporter = EnumExportTo.Excel Then
                transportSql.exportdata(gv1, "", "", False, arrHeader, False, True, True)
            Else
                clsCommon.MyExportToPDF(strHeading, gv1, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        ExportGrid(EnumExportTo.Excel)
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        ExportGrid(EnumExportTo.PDF)
    End Sub
End Class