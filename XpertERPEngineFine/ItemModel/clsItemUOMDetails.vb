Imports System.Data.SqlClient
Public Class clsItemUOMDetails
    Public Item_Code As String = ""
    Public UOM_Code As String = ""
    Public UOM_Description As String = ""
    Public InterUnionItem As String = ""
    Public Conversion_Factor As Double = 0
    Public Stocking_Unit As String = ""
    Public Default_UOM As Integer = 0
    Public Pieces As Integer = 0
    Public Gross_Weight As Double = 0
    Public Net_Weight As Double = 0
    Public Job_Work_Rate As Double = 0
    Public Item_Cost As Decimal
    Public Custom_Conversion As Boolean
    Public Print_UOM As Integer = 0
    Public Report_UOM As Integer = 0
    Public Decimal_UOM As Integer = 0
    Public Bulk_UOM As Integer = 0
    Public Billing_UOM As Integer = 0
    Public Loose_UOM As Integer = 0
    Public RMProcessLoss_UOM As Integer = 0


    Public Shared Function GetEntryUOM() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "0"
        dr("Name") = "Crate And Pouch"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "3"
        dr("Name") = "Crate and LTR"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "4"
        dr("Name") = "Pouch and LTR"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "1"
        dr("Name") = "Crate"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "5"
        dr("Name") = "Pouch"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "2"
        dr("Name") = "LTR"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "6"
        dr("Name") = "ALL"
        dt.Rows.Add(dr)

        Return dt
    End Function

    Public Shared Function SaveData(ByVal strICode As String, ByVal ArrUomDetails As List(Of clsItemUOMDetails), ByVal ArrDatabase As List(Of String), ByVal trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Dim StockUnitItemCost As Decimal = 0
        For Each obj As clsItemUOMDetails In ArrUomDetails
            If clsCommon.CompairString(obj.Stocking_Unit, "Y") = CompairStringResult.Equal Then
                StockUnitItemCost = obj.Item_Cost
                Exit For
            End If
        Next

        For Each obj As clsItemUOMDetails In ArrUomDetails
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Item_Code", strICode)
            clsCommon.AddColumnsForChange(coll, "UOM_Code", obj.UOM_Code)
            clsCommon.AddColumnsForChange(coll, "UOM_Description", obj.UOM_Description)
            clsCommon.AddColumnsForChange(coll, "Conversion_Factor", obj.Conversion_Factor)
            clsCommon.AddColumnsForChange(coll, "Stocking_Unit", obj.Stocking_Unit)
            ''added by richa agarwal against ticket no BM00000004327
            clsCommon.AddColumnsForChange(coll, "Default_UOM", obj.Default_UOM)
            clsCommon.AddColumnsForChange(coll, "InterUnionItem", obj.InterUnionItem, True)

            clsCommon.AddColumnsForChange(coll, "Pieces", obj.Pieces)
            '======================================

            clsCommon.AddColumnsForChange(coll, "Gross_Weight", obj.Gross_Weight)
            clsCommon.AddColumnsForChange(coll, "Net_Weight", obj.Net_Weight)
            clsCommon.AddColumnsForChange(coll, "Job_Work_Rate", obj.Job_Work_Rate)
            clsCommon.AddColumnsForChange(coll, "Custom_Conversion", IIf(obj.Custom_Conversion, 1, 0))
            clsCommon.AddColumnsForChange(coll, "Print_UOM", obj.Print_UOM)
            clsCommon.AddColumnsForChange(coll, "Report_UOM", obj.Report_UOM)
            clsCommon.AddColumnsForChange(coll, "Bulk_UOM", obj.Bulk_UOM)
            clsCommon.AddColumnsForChange(coll, "Billing_UOM", obj.Billing_UOM)
            clsCommon.AddColumnsForChange(coll, "Loose_UOM", obj.Loose_UOM)
            clsCommon.AddColumnsForChange(coll, "Decimal_UOM", obj.Decimal_UOM)
            clsCommon.AddColumnsForChange(coll, "ProcessLoss_UOM", obj.RMProcessLoss_UOM)
            clsCommon.AddColumnsForChange(coll, "Item_Cost", Math.Round(StockUnitItemCost * obj.Conversion_Factor, 2, MidpointRounding.AwayFromZero))

            isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTableInSelectedDatabase(coll, ArrDatabase, "TSPL_ITEM_UOM_DETAIL", OMInsertOrUpdate.Insert, "", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strICode, "TSPL_ITEM_UOM_DETAIL", "Item_Code", trans)

        Next
        Return isSaved
    End Function

    Public Shared Function GetName(ByVal strcode As String)
        Dim qry As String = "select Unit_Desc from TSPL_UNIT_MASTER where Unit_Code='" + strcode + "'"
        Return clsDBFuncationality.getSingleValue(qry)
    End Function

    Public Shared Function GetItemUOMCost(ByVal tranDate As DateTime, ByVal strcode As String, ByVal strUOM As String, ByVal tran As SqlTransaction) As Decimal
        Dim retval As Decimal = 0
        Dim qry As String = "select Item_Cost from TSPL_ITEM_UOM_DETAIL_Hist_Data where item_code='" + strcode + "' and  UOM_Code='" + strUOM + "' and Hist_On>='" + clsCommon.GetPrintDate(tranDate, "dd/MMM/yyyy hh:mm:ss tt") + "' order by Hist_On  "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, tran)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                retval = clsCommon.myCdbl(dr("Item_Cost"))
                If retval > 0 Then
                    Exit For
                End If
            Next
        End If
        If retval <= 0 Then
            qry = "select Item_Cost from tspl_item_uom_detail where item_code='" + strcode + "' and  UOM_Code='" + strUOM + "'"
            dt = clsDBFuncationality.GetDataTable(qry, tran)
            retval = clsCommon.myCdbl(dt.Rows(0)("Item_Cost"))
        End If
        Return retval
    End Function
End Class