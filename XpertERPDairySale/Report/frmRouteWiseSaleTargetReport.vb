Imports System.Text
Imports common
Public Class frmRouteWiseSaleTargetReport

#Region "Variables"
    Dim dtItem As DataTable = Nothing
#End Region
    Private Sub frmRouteWiseSaleTargetReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            txtToDate.Value = clsCommon.GETSERVERDATE()
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            Reset()
            EnableDisableFields(True)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        Gv1.GroupDescriptors.Clear()
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()
        Gv1.MasterView.Refresh()
    End Sub

    Sub EnableDisableFields(ByVal isEnable As Boolean)
        txtToDate.Enabled = isEnable
        If isEnable Then
            RadPageView1.SelectedPage = RadPageViewPage1
        Else
            RadPageView1.SelectedPage = RadPageViewPage2
        End If
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            Dim strQry As String = ReturnQry(False)
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Gv1.DataSource = Nothing
                Gv1.Rows.Clear()
                Gv1.Columns.Clear()
                Gv1.GroupDescriptors.Clear()
                Gv1.MasterTemplate.SummaryRowsBottom.Clear()
                Gv1.MasterView.Refresh()
                Gv1.DataSource = dt
                Gv1.EnableFiltering = True
                Gv1.AllowAddNewRow = False
                Gv1.AllowDragToGroup = False
                Gv1.ReadOnly = True
                EnableDisableFields(False)

                For rowI As Integer = 0 To Gv1.Rows.Count - 1
                    For colI As Integer = 2 To Gv1.Columns.Count - 1
                        Dim cellValue As Object = Gv1.Rows(rowI).Cells(Gv1.Columns(colI).Name).Value
                        If cellValue IsNot Nothing AndAlso IsNumeric(cellValue) Then
                            Dim col As GridViewDecimalColumn = TryCast(Gv1.Columns(Gv1.Columns(colI).Name), GridViewDecimalColumn)
                            If col IsNot Nothing Then
                                col.FormatString = "{0:N0}"
                            End If
                            col = Nothing
                        End If
                    Next
                Next
                View()
                Gv1.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub View()
        If Gv1.Rows.Count > 0 Then
            Dim strQry As String = ReturnItemSubQry()
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                Dim view As New ColumnGroupsViewDefinition()
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns("Code").Name)
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns("Description").Name)

                For Each dtRow As DataRow In dt.Rows
                    i += 1
                    view.ColumnGroups.Add(New GridViewColumnGroup(Gv1.Columns(i + 1).Name))
                    view.ColumnGroups(view.ColumnGroups.Count - 1).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns(i + 1).Name)
                    Gv1.Columns(Gv1.Columns(i + 1).Name).HeaderText = clsCommon.myCstr(dtRow("UOM_Code"))
                Next

                view.ColumnGroups.Add(New GridViewColumnGroup("TGT MILK"))
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns("TGT MILK").Name)
                Gv1.Columns("TGT MILK").HeaderText = clsCommon.GetPrintDate(txtToDate.Value, "MMM-yyyy").ToUpper()

                view.ColumnGroups.Add(New GridViewColumnGroup("ACHV MILK"))
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns("ACHV MILK").Name)
                Gv1.Columns("ACHV MILK").HeaderText = clsCommon.GetPrintDate(txtToDate.Value, "MMM-yyyy").ToUpper()

                view.ColumnGroups.Add(New GridViewColumnGroup("PROGRESS"))
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(view.ColumnGroups.Count - 1).Rows(0).ColumnNames.Add(Gv1.Columns("PROGRESS").Name)
                Gv1.Columns("PROGRESS").HeaderText = "%"

                Gv1.ViewDefinition = view
            End If

            'Dim view As New ColumnGroupsViewDefinition()
            'view.ColumnGroups.Add(New GridViewColumnGroup(""))
            'view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            'view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("SNo").Name)
            'view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Item").Name)
            'view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("OPBal").Name)

            'view.ColumnGroups.Add(New GridViewColumnGroup("SALE"))
            'view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())

            'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Credit_Amt").Name)
            'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Cash_Amt").Name)

            'view.ColumnGroups.Add(New GridViewColumnGroup(""))
            'view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())

            'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Total").Name)
            'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Amt_Ded").Name)
            'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Depo_Amt").Name)
            'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Closing_Bal").Name)
            'Gv1.ViewDefinition = view
        End If
    End Sub

    Public Function ReturnItemSubQry() As String
        Dim strQry As String = "Select tspl_chapter_head.Chapter_Head_Code,MAX(ItemUOMDetails.Description)Description,ItemUOMDetails.UOM_Code from(
select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Item_Desc,tspl_chapter_head.Chapter_Head_Code,tspl_chapter_head.Description,
TSPL_ITEM_UOM_DETAIL.Default_UOM, TSPL_ROUTE_WISE_SALE_TARGET.UOM As UOM_Code from tspl_chapter_head
Left Outer Join TSPL_ROUTE_WISE_SALE_TARGET On TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category=tspl_chapter_head.Chapter_Head_Code
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Sub_Group_Type=TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category
Left Outer Join TSPL_ITEM_UOM_DETAIL On TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
where IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Inactive,0)=0 And TSPL_ROUTE_WISE_SALE_TARGET.Status=1 And Months='" & clsCommon.GetPrintDate(txtToDate.Value, "MMM/yyyy") & "'
Union All
select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Item_Desc,tspl_chapter_head.Chapter_Head_Code,tspl_chapter_head.Description,
TSPL_ITEM_UOM_DETAIL.Default_UOM,TSPL_ITEM_UOM_DETAIL.UOM_Code from tspl_chapter_head
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Sub_Group_Type=tspl_chapter_head.Chapter_Head_Code
Left Outer Join TSPL_ITEM_UOM_DETAIL On TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code And TSPL_ITEM_UOM_DETAIL.Default_UOM=1
Where tspl_chapter_head.Chapter_Head_Code Not In (Select Item_Sub_Category from TSPL_ROUTE_WISE_SALE_TARGET Where IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Inactive,0)=0 And TSPL_ROUTE_WISE_SALE_TARGET.Status=1 And Months='" & clsCommon.GetPrintDate(txtToDate.Value, "MMM/yyyy") & "' Group By Item_Sub_Category)
)ItemUOMDetails Left Outer join tspl_chapter_head On tspl_chapter_head.chapter_head_code=ItemUOMDetails.Chapter_Head_Code
Where ISNULL(UOM_Code,'')<>'' 
Group By tspl_chapter_head.Chapter_Head_Code,UOM_Code,tspl_chapter_head.Seq_No
Order By tspl_chapter_head.Seq_No"
        Return strQry
    End Function

    Public Function ReturnDatailsDataQry(ByVal isPrint As Boolean, ByVal FromDate As DateTime, ByVal ToDate As DateTime) As String
        Dim strQry As String = Nothing
        'ReturnItemSubQry()
        'Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
        If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
            Dim sbLTR As New StringBuilder()
            Dim sbKG As New StringBuilder()
            Dim strItemType As New StringBuilder()
            Dim strItemTypeIN As New StringBuilder()
            Dim i As Integer = 0
            For Each UOM In dtItem.Rows
                If clsCommon.myLen(UOM("Chapter_Head_Code")) <> 0 AndAlso clsCommon.myLen(UOM("UOM_Code")) > 0 AndAlso clsCommon.CompairString(UOM("UOM_Code"), "LTR") = CompairStringResult.Equal Then
                    If clsCommon.myLen(sbLTR) > 0 Then
                        sbLTR.Append(",")
                    End If
                    sbLTR.Append("'" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "'")
                End If
                If clsCommon.myLen(UOM("Chapter_Head_Code")) <> 0 AndAlso clsCommon.myLen(UOM("UOM_Code")) > 0 AndAlso clsCommon.CompairString(UOM("UOM_Code"), "KG") = CompairStringResult.Equal Then
                    If clsCommon.myLen(sbKG) > 0 Then
                        sbKG.Append(",")
                    End If
                    sbKG.Append("'" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "'")
                End If

                If isPrint Then
                    strItemType.Append(",Max([DescGroup" & clsCommon.myCstr(i + 1) & "]) As [DescGroup" & clsCommon.myCstr(i + 1) & "]")
                    strItemType.Append(",Max([DescGroupUOM" & clsCommon.myCstr(i + 1) & "]) As [DescGroupUOM" & clsCommon.myCstr(i + 1) & "]")
                    strItemType.Append(",SUM([" & clsCommon.myCstr(UOM("Description")) & "]) As [Group" & clsCommon.myCstr(i + 1) & "]")
                Else
                    strItemType.Append(",SUM([" & clsCommon.myCstr(UOM("Description")) & "]) As [" & clsCommon.myCstr(UOM("Description")) & "]")
                End If

                strItemTypeIN.Append(" '" & clsCommon.myCstr(UOM("Description")) & "' As [DescGroup" & clsCommon.myCstr(i + 1) & "],")
                strItemTypeIN.Append(" '" & clsCommon.myCstr(UOM("UOM_Code")) & "' As [DescGroupUOM" & clsCommon.myCstr(i + 1) & "],")
                If clsCommon.myLen(UOM("UOM_Code")) > 0 Then
                    strItemTypeIN.Append(" Sum(QtyIn" & clsCommon.myCstr(UOM("UOM_Code")) & ")*Case When Max(DetailData.Document_Date)='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' And Max(Item_Sub_Group_Type)='" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "' Then 1 Else 0 End As '" & clsCommon.myCstr(UOM("Description")) & "',")
                Else
                    strItemTypeIN.Append(" 0 As '" & clsCommon.myCstr(UOM("Description")) & "',")
                End If

                i += 1
            Next


            strQry = Nothing
            strQry = ";WITH DetailData AS
(SELECT TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Document_Date,TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,
TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code,TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Desc,
IsNull(TSPL_CUSTOMER_GROUP_MASTER.IsGoverment,0)IsGoverment,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,
TSPL_DEMAND_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.Unit_code,TSPL_ITEM_UOM_DETAIL.Default_UOM,
TSPL_ITEM_UOM_DETAIL.UOM_Code As [Default_UOM_Code],TSPL_ITEM_MASTER.Item_Sub_Group_Type,tspl_chapter_head.Description,
TSPL_ITEM_UOM_DETAIL.Conversion_Factor,
TSPL_ITEM_UOM_DETAIL_IN_LTR.Conversion_Factor As [CFinLTR],TSPL_ITEM_UOM_DETAIL_IN_KG.Conversion_Factor As [CFinKG],"

            If clsCommon.myLen(sbLTR) > 0 Then
                strQry += " Case When TSPL_ITEM_MASTER.Item_Sub_Group_Type IN (" & clsCommon.myCstr(sbLTR) & ") Then (TSPL_DEMAND_BOOKING_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/TSPL_ITEM_UOM_DETAIL_IN_LTR.Conversion_Factor) Else 0 End As QtyInLTR ,"
            Else
                strQry += " 0 As QtyInLTR, "
            End If
            If clsCommon.myLen(sbKG) > 0 Then
                strQry += " Case When TSPL_ITEM_MASTER.Item_Sub_Group_Type IN (" & clsCommon.myCstr(sbKG) & ") Then (TSPL_DEMAND_BOOKING_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/TSPL_ITEM_UOM_DETAIL_IN_KG.Conversion_Factor) Else 0 End As QtyInKG ,"
            Else
                strQry += " 0 As QtyInKG,"
            End If

            strQry += " TSPL_ROUTE_WISE_SALE_TARGET.Months,IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Target_Qty,0)Target_Qty,TSPL_ROUTE_WISE_SALE_TARGET.Route_And_Group,tspl_chapter_head.Seq_No FROM TSPL_DEMAND_BOOKING_DETAIL
Left Outer Join TSPL_DEMAND_BOOKING_MASTER On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
Left Outer Join TSPL_CUSTOMER_GROUP_MASTER On TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code

LEFT Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code

Left Outer Join TSPL_ITEM_UOM_DETAIL On TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code 
				And TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
				--And TSPL_ITEM_UOM_DETAIL.Default_UOM=1

Left Outer Join TSPL_ITEM_UOM_DETAIL As TSPL_ITEM_UOM_DETAIL_IN_LTR On TSPL_ITEM_UOM_DETAIL_IN_LTR.Item_Code=TSPL_ITEM_MASTER.Item_Code 				
				And TSPL_ITEM_UOM_DETAIL_IN_LTR.UOM_Code='LTR'
Left Outer Join TSPL_ITEM_UOM_DETAIL As TSPL_ITEM_UOM_DETAIL_IN_KG On TSPL_ITEM_UOM_DETAIL_IN_KG.Item_Code=TSPL_ITEM_MASTER.Item_Code 
				And TSPL_ITEM_UOM_DETAIL_IN_KG.UOM_Code='KG'

Left Outer Join tspl_chapter_head On tspl_chapter_head.Chapter_Head_Code=TSPL_ITEM_MASTER.Item_Sub_Group_Type
Left Outer Join TSPL_ROUTE_MASTER On TSPL_ROUTE_MASTER.Route_No=TSPL_DEMAND_BOOKING_MASTER.Route_No

Left Outer Join(Select Months,Item_Sub_Category,Target_On,Case When	IsNull(Route_Code,'')<>'' Then Route_Code Else Cust_Group_Code End As Route_And_Group,	Target_Qty 
from (Select TSPL_ROUTE_WISE_SALE_TARGET.Months,TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category,TSPL_ROUTE_WISE_SALE_TARGET.Target_On,TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code,
TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Cust_Group_Code,IsNull(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Target_Qty,0) AS Target_Qty 
from TSPL_ROUTE_WISE_SALE_TARGET_DETAIL
Left Outer Join TSPL_ROUTE_WISE_SALE_TARGET  On TSPL_ROUTE_WISE_SALE_TARGET.Document_Code=TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Document_Code 
Where ISNULL(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code, '') <> '' And TSPL_ROUTE_WISE_SALE_TARGET.Status=1 "
            If Not isPrint Then
                strQry += " And TSPL_ROUTE_WISE_SALE_TARGET.Months='" & clsCommon.GetPrintDate(txtToDate.Value, "MMM/yyyy") & "'"

            End If
            strQry += " Union All
Select TSPL_ROUTE_WISE_SALE_TARGET.Months,TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category,TSPL_ROUTE_WISE_SALE_TARGET.Target_On,TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code,
TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Cust_Group_Code,IsNull(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Target_Qty,0) AS Target_Qty 
from TSPL_ROUTE_WISE_SALE_TARGET_DETAIL
Left Outer Join TSPL_ROUTE_WISE_SALE_TARGET  On TSPL_ROUTE_WISE_SALE_TARGET.Document_Code=TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Document_Code 
Where ISNULL(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Cust_Group_Code, '') <> '' And TSPL_ROUTE_WISE_SALE_TARGET.Status=1 "
            If Not isPrint Then
                strQry += " And TSPL_ROUTE_WISE_SALE_TARGET.Months='" & clsCommon.GetPrintDate(txtToDate.Value, "MMM/yyyy") & "'"
            End If
            strQry += ")SALE_TARGET)TSPL_ROUTE_WISE_SALE_TARGET On TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category=TSPL_ITEM_MASTER.Item_Sub_Group_Type 
And
TSPL_ROUTE_WISE_SALE_TARGET.Route_And_Group=Case When TSPL_ROUTE_WISE_SALE_TARGET.Target_On=1 Then TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code Else TSPL_Route_Master.Route_No End

where 2=2 And IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Route_And_Group,'')<>'' and TSPL_DEMAND_BOOKING_MASTER.Document_Date >= '01/" & clsCommon.GetPrintDate(FromDate, "MMM/yyyy") & "' And TSPL_DEMAND_BOOKING_MASTER.Document_Date <= '" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "'
--and TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code='ARMY U'
--And IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Target_Qty,0)<>0
)"
            Else
                Throw New Exception("Item sub group type not found !")
        End If
        Return strQry
    End Function

    Public Function ReturnQry(ByVal isPrint As Boolean) As String
        Dim strQry As String
        Try
            strQry = ReturnItemSubQry()
            dtItem = clsDBFuncationality.GetDataTable(strQry)
            If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
                Dim sbLTR As New StringBuilder()
                Dim sbKG As New StringBuilder()
                Dim strItemType As New StringBuilder()
                Dim strItemTypeGroup As New StringBuilder()
                Dim strItemTypeIN As New StringBuilder()
                Dim i As Integer = 0
                For Each UOM In dtItem.Rows
                    If clsCommon.myLen(UOM("Chapter_Head_Code")) <> 0 AndAlso clsCommon.myLen(UOM("UOM_Code")) > 0 AndAlso clsCommon.CompairString(UOM("UOM_Code"), "LTR") = CompairStringResult.Equal Then
                        If clsCommon.myLen(sbLTR) > 0 Then
                            sbLTR.Append(",")
                        End If
                        sbLTR.Append("'" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "'")
                    End If
                    If clsCommon.myLen(UOM("Chapter_Head_Code")) <> 0 AndAlso clsCommon.myLen(UOM("UOM_Code")) > 0 AndAlso clsCommon.CompairString(UOM("UOM_Code"), "KG") = CompairStringResult.Equal Then
                        If clsCommon.myLen(sbKG) > 0 Then
                            sbKG.Append(",")
                        End If
                        sbKG.Append("'" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "'")
                    End If

                    If isPrint Then
                        strItemType.Append(",Max([DescGroup" & clsCommon.myCstr(i + 1) & "]) As [DescGroup" & clsCommon.myCstr(i + 1) & "]")
                        strItemType.Append(",Max([DescGroupUOM" & clsCommon.myCstr(i + 1) & "]) As [DescGroupUOM" & clsCommon.myCstr(i + 1) & "]")
                        strItemType.Append(",SUM([" & clsCommon.myCstr(UOM("Description")) & "]) As [Group" & clsCommon.myCstr(i + 1) & "]")

                        strItemTypeGroup.Append(",Max([DescGroup" & clsCommon.myCstr(i + 1) & "]) As [DescGroup" & clsCommon.myCstr(i + 1) & "]")
                        strItemTypeGroup.Append(",Max([DescGroupUOM" & clsCommon.myCstr(i + 1) & "]) As [DescGroupUOM" & clsCommon.myCstr(i + 1) & "]")
                        strItemTypeGroup.Append(",SUM([Group" & clsCommon.myCstr(i + 1) & "]) As [Group" & clsCommon.myCstr(i + 1) & "]")
                    Else
                        strItemType.Append(",SUM([" & clsCommon.myCstr(UOM("Description")) & "]) As [" & clsCommon.myCstr(UOM("Description")) & "]")
                    End If

                    strItemTypeIN.Append(" '" & clsCommon.myCstr(UOM("Description")) & "' As [DescGroup" & clsCommon.myCstr(i + 1) & "],")
                    strItemTypeIN.Append(" '" & clsCommon.myCstr(UOM("UOM_Code")) & "' As [DescGroupUOM" & clsCommon.myCstr(i + 1) & "],")
                    If clsCommon.myLen(UOM("UOM_Code")) > 0 Then
                        strItemTypeIN.Append(" (QtyIn" & clsCommon.myCstr(UOM("UOM_Code")) & ")*Case When (DetailData.Document_Date)='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' And (Item_Sub_Group_Type)='" & clsCommon.myCstr(UOM("Chapter_Head_Code")) & "' Then 1 Else 0 End As '" & clsCommon.myCstr(UOM("Description")) & "',")
                    Else
                        strItemTypeIN.Append(" 0 As '" & clsCommon.myCstr(UOM("Description")) & "',")
                    End If
                    i += 1
                Next


                strQry = ReturnDatailsDataQry(isPrint, txtToDate.Value, txtToDate.Value)
                strQry += " Select "
                If isPrint Then
                    strQry += " Day(Convert(Date,'" & txtToDate.Value & "',103)) As FilterDays,'" & clsCommon.GetPrintDate(txtToDate.Value, "MMM-yyyy").ToUpper() & "' As FilterMonthYear,"
                End If
                strQry += "final.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_STATE_MASTER.STATE_NAME,
TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from(
Select Code,MAX(Description)Description " & clsCommon.myCstr(strItemType) & ",Max(Target_Qty)[TGT MILK],SUM([ACHV MILK])[ACHV MILK],Case When Max(Target_Qty)<>0 Then Round(((SUM([ACHV MILK])/Max(Target_Qty))*100),0) Else 0 End As 'PROGRESS' from (
Select 
DetailData.Route_And_Group As Code,(Route_Desc) As Description, "
                strQry += clsCommon.myCstr(strItemTypeIN)
                strQry += " (DetailData.Months)Months,IsNull((DetailData.Target_Qty),0)Target_Qty,
((QtyInLTR)*Case When Item_Sub_Group_Type='MILK' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "') As 'ACHV MILK'
from DetailData Where DetailData.IsGoverment=0
)xyz Group By Code

Union All

Select Code,MAX(Description)Description " & clsCommon.myCstr(strItemType) & ",Max(Target_Qty)[TGT MILK],SUM([ACHV MILK])[ACHV MILK],Case When Max(Target_Qty)<>0 Then Round(((SUM([ACHV MILK])/Max(Target_Qty))*100),0) Else 0 End As 'PROGRESS' from (
Select 
DetailData.Route_And_Group As Code,(TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Desc) As Description, "
                strQry += clsCommon.myCstr(strItemTypeIN)
                strQry += " (DetailData.Months)Months,IsNull((DetailData.Target_Qty),0)Target_Qty,
((QtyInLTR)*Case When Item_Sub_Group_Type='MILK' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "') As 'ACHV MILK'
from DetailData
Inner Join TSPL_CUSTOMER_GROUP_MASTER On TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=DetailData.Route_And_Group
Where TSPL_CUSTOMER_GROUP_MASTER.IsGoverment=1
)xyz Group By Code

Union All

Select Null As Code,'CITY(LTR)' As Description " & IIf(isPrint, clsCommon.myCstr(strItemTypeGroup), clsCommon.myCstr(strItemType)) & ",Sum([TGT MILK])[TGT MILK],SUM([ACHV MILK])[ACHV MILK],Case When Sum([TGT MILK])<>0 Then Round(((SUM([ACHV MILK])/Sum([TGT MILK]))*100),0) Else 0 End As 'PROGRESS' from (
Select Code,MAX(Description)Description " & clsCommon.myCstr(strItemType) & ",Max(Target_Qty)[TGT MILK],SUM([ACHV MILK])[ACHV MILK] from (
Select 
DetailData.Route_And_Group As Code,(Route_Desc) As Description, "
                strQry += clsCommon.myCstr(strItemTypeIN)
                strQry += " (DetailData.Months)Months,IsNull((DetailData.Target_Qty),0)Target_Qty,
((QtyInLTR)*Case When  Item_Sub_Group_Type='MILK' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "') As 'ACHV MILK'
from DetailData Where DetailData.IsGoverment<>1
)xyz Group By Code)CityTotal

Union All

Select Null As Code,'TOTAL(LTR)' As Description " & IIf(isPrint, clsCommon.myCstr(strItemTypeGroup), clsCommon.myCstr(strItemType)) & ",Sum([TGT MILK])[TGT MILK],SUM([ACHV MILK])[ACHV MILK],Case When Sum([TGT MILK])<>0 Then Round(((SUM([ACHV MILK])/Sum([TGT MILK]))*100),0) Else 0 End As 'PROGRESS' from (
Select Code,MAX(Description)Description " & clsCommon.myCstr(strItemType) & ",Max(Target_Qty)[TGT MILK],SUM([ACHV MILK])[ACHV MILK] from (
Select DetailData.Route_And_Group As Code,(Route_Desc) As Description, "
                strQry += clsCommon.myCstr(strItemTypeIN)
                strQry += " (DetailData.Months)Months,IsNull((DetailData.Target_Qty),0)Target_Qty,
((QtyInLTR)*Case When Item_Sub_Group_Type='MILK' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "') As 'ACHV MILK'
from DetailData Where DetailData.IsGoverment=0
Union All
Select DetailData.Route_And_Group As Code,(TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Desc) As Description, "
                strQry += clsCommon.myCstr(strItemTypeIN)
                strQry += " (DetailData.Months)Months,IsNull((DetailData.Target_Qty),0)Target_Qty,
((QtyInLTR)*Case When Item_Sub_Group_Type='MILK' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "') As 'ACHV MILK'
from DetailData 
Inner Join TSPL_CUSTOMER_GROUP_MASTER On TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=DetailData.Route_And_Group
Where TSPL_CUSTOMER_GROUP_MASTER.IsGoverment=1
)xyz Group By Code)AllTotal )final
Left Outer Join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State ORDER BY CASE WHEN Code IS NULL THEN 1 ELSE 0 END, Code"
            Else
                Throw New Exception("Item sub group type not found !")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return strQry
    End Function

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim strQry As New StringBuilder()
            strQry.Append(ReturnQry(True))
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(clsCommon.myCstr(strQry))
            strQry = Nothing
            strQry = New StringBuilder()
            Dim ToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            Dim fromDate As String = "01/" & clsCommon.GetPrintDate(txtToDate.Value.AddYears(-1), "MMM/yyyy")
            strQry.Append(ReturnDatailsDataQry(True, fromDate, ToDate))
            strQry.Append(" Select [S.No.],MAX(Details)Details,MAX(Item_Sub_Group_Type)Item_Sub_Group_Type,Sum([Details Qty])[Details Qty],Seq_No from (")
            If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each rows In dtItem.Rows
                    If i <> 0 Then
                        strQry.Append(" Union All ")
                    End If

                    strQry.Append(" Select 1 As [S.No.],'" & clsCommon.GetPrintDate(txtToDate.Value.AddYears(-1), "MMM-yy").ToUpper() & " ACHV AVG' As Details,Item_Sub_Group_Type,
(Sum(QtyInLTR)*Case When Max(DetailData.Document_Date)<='" & clsCommon.GetPrintDate(txtToDate.Value.AddYears(-1), "dd/MMM/yyyy") & "' And Item_Sub_Group_Type='" & clsCommon.myCstr(rows("Chapter_Head_Code")) & "' Then 1 Else 0 End)/DAY('" & clsCommon.GetPrintDate(txtToDate.Value.AddYears(-1), "dd/MMM/yyyy") & "') As [Details Qty],DetailData.Seq_No
from DetailData 
Where Document_Date >= '" & fromDate & "' And Document_Date <= '" & ToDate & "'
Group BY Item_Sub_Group_Type,Seq_No

Union All

Select 2 As [S.No.],'" & clsCommon.GetPrintDate(txtToDate.Value, "MMM-yy").ToUpper() & " ACHV AVG' As Details,Item_Sub_Group_Type,
(Sum(QtyInLTR)*Case When Max(DetailData.Document_Date)<='" & ToDate & "' And Item_Sub_Group_Type='" & clsCommon.myCstr(rows("Chapter_Head_Code")) & "' Then 1 Else 0 End)/DAY('" & ToDate & "') As [Details Qty],DetailData.Seq_No
from DetailData 
Where Document_Date >= '01/" & clsCommon.GetPrintDate(ToDate, "MMM/yyyy") & "' And Document_Date <= '" & ToDate & "'
Group BY Item_Sub_Group_Type,Seq_No


Union All

Select 3 As [S.No.],'SALE LTR/KG TILL DATE' As Details,Item_Sub_Group_Type,
Case When Sum(QtyInLTR)>0 Then Sum(QtyInLTR) When Sum(QtyInKG)>0  Then Sum(QtyInKG)	Else 0 End As [Details Qty],DetailData.Seq_No
from DetailData Where Document_Date >= '01/" & clsCommon.GetPrintDate(ToDate, "MMM/yyyy") & "' And Document_Date <= '" & ToDate & "'
Group BY Item_Sub_Group_Type,Seq_No

Union All

Select 4 As [S.No.],'AVG LTR/KG TILL DATE' As Details,Item_Sub_Group_Type,
Case When Sum(QtyInLTR)>0 Then
    (Sum(QtyInLTR)*Case When Max(DetailData.Document_Date)<='" & ToDate & "' And Item_Sub_Group_Type='" & clsCommon.myCstr(rows("Chapter_Head_Code")) & "' Then 1 Else 0 End)/DAY('" & ToDate & "')
	When Sum(QtyInKG)>0 Then
	(Sum(QtyInKG)*Case When Max(DetailData.Document_Date)<='" & ToDate & "' And Item_Sub_Group_Type='" & clsCommon.myCstr(rows("Chapter_Head_Code")) & "' Then 1 Else 0 End)/DAY('" & ToDate & "')
	Else 0 End
As [Details Qty],DetailData.Seq_No
from DetailData 
Where Document_Date >= '01/" & clsCommon.GetPrintDate(ToDate, "MMM/yyyy") & "' And Document_Date <= '" & ToDate & "'
Group BY Item_Sub_Group_Type,Seq_No

Union All

Select 5 As [S.No.],'PROGRESS %' As Details,Item_Sub_Group_Type,
Case When Sum(Target_Qty)<>0 Then Round((((Sum(QtyInLTR)*Case When Max(DetailData.Document_Date)<='" & ToDate & "' And Max(Item_Sub_Group_Type)='" & clsCommon.myCstr(rows("Chapter_Head_Code")) & "' Then 1 Else 0 End)/DAY('" & ToDate & "')/SUM(Target_Qty))*100),0) Else 0 End As  [Details Qty],DetailData.Seq_No
from DetailData 
Where Document_Date >= '01/" & clsCommon.GetPrintDate(ToDate, "MMM/yyyy") & "' And Document_Date <= '" & ToDate & "'
Group BY Item_Sub_Group_Type,Seq_No ")
                    i += 1
                Next
            End If
            strQry.Append(" )final Group By [S.No.],Seq_No Order By [S.No.],Seq_No")
            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(clsCommon.myCstr(strQry))
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frm As New frmCrystalReportViewer()
                frm.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dt, dt2, "crptRouteWiseSaleTargetReport", "Route Wise Sale Target Report", "crptSubRouteWiseSaleTargetReportSummary.rpt")
                'frm.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dt2, Nothing, "crptSubRouteWiseSaleTargetReportSummary", "Route Wise Sale Target Report")
                frm = Nothing
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class