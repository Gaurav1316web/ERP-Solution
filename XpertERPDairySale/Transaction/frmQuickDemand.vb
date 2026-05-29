Imports System.Data.SqlClient
Imports System.IO
Imports common
Public Class frmQuickDemand
    Inherits FrmMainTranScreen
#Region "Variables"
    Private isNewEntry As Boolean = False
    Dim UOMCrate As String = ""
    Dim isLastcellvaluechanged As Boolean = False
    Dim AllowRouteWiseDemandEntryInDecimal As Boolean = False
    Dim UOMPouch As String = ""
    Dim UOMLtr As String = ""
    Dim SetShiftTimeOut As String = ""
    Dim isInsideLoadData As Boolean = False
    Dim allowtoselectShift As Boolean = False
    Dim ConvertPouchtoCrate As Boolean = False
    Dim DontCreateForPouch As Boolean = False
    Dim isLoadData As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim QuickDemandUOMCrate As Boolean = False
    Dim QuickDemandUOMPouch As Boolean = False
    Dim QuickDemandUOMLtr As Boolean = False
    Const colLineNo As String = "colLineNo"
    Const colCustCode As String = "colCustCode"
    Const colCustPhone As String = "colCustPhone"
    Const colRouteNo As String = "colRouteNo"
    Const colSetZero As String = "colSetZero"
    Const colItemCode As String = "colItemCode"
    Const colCrate As String = "colCrate"
#End Region
    Public Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.frmbookingdairy)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
            Me.Close()
            Exit Sub
        End If
        btnSave.Visible = MyBase.isModifyFlag
        'btnExport.Visible = MyBase.isExport
        'btnPost.Visible = MyBase.isPostFlag
        If MyBase.isExport = True Then
            rmiImport.Enabled = True
            rmiExcel.Enabled = True
        Else
            rmiImport.Enabled = False
            rmiExcel.Enabled = False
        End If
    End Sub
    Private Sub frmQuickDemand_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetShiftTimeOut = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SetShiftTimeOut, clsFixedParameterCode.SetShiftTimeOut, Nothing))
        allowtoselectShift = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowToSelectShift, clsFixedParameterCode.AllowToSelectShift, Nothing)) = 1, True, False)
        QuickDemandUOMCrate = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickDemandUOMCrate, clsFixedParameterCode.QuickDemandUOMCrate, Nothing)) = 1, True, False)
        QuickDemandUOMPouch = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickDemandUOMPouch, clsFixedParameterCode.QuickDemandUOMPouch, Nothing)) = 1, True, False)
        QuickDemandUOMLtr = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickDemandUOMLtr, clsFixedParameterCode.QuickDemandUOMLtr, Nothing)) = 1, True, False)
        ConvertPouchtoCrate = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ConvertPouchtoCrate, clsFixedParameterCode.ConvertPouchtoCrate, Nothing)) = 1, True, False)
        DontCreateForPouch = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DontCreateForPouch, clsFixedParameterCode.DontCreateForPouch, Nothing)) = 1, True, False)
        AllowRouteWiseDemandEntryInDecimal = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowRouteWiseDemandEntryInDecimal, clsFixedParameterCode.AllowRouteWiseDemandEntryInDecimal, Nothing)) = 1, True, False)

        If allowtoselectShift = True Then
            txtDate.Enabled = True
            cmbShift.Enabled = True
        Else
            txtDate.Enabled = False
            cmbShift.Enabled = False
        End If
        LoadShiftType()
        AddNew()
        SetUserMgmtNew()
        'btnSave.Visible = False
        'txtDate.Value = clsCommon.GETSERVERDATE()
        'LoadData(txtDate.Value, cmbShift.SelectedValue, objCommonVar.CurrentUserCode, True)
        isInsideLoadData = False
    End Sub
    Sub LoadBlankGrid()
        Dim qry As String = String.Empty
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        gv1.DataSource = Nothing
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.IsPinned = True
        'repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gv1.MasterTemplate.Columns.Add(repoLineNo)
        Dim repoCustCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCustCode.FormatString = ""
        repoCustCode.HeaderText = "Booth Code"
        repoCustCode.Name = colCustCode
        repoCustCode.HeaderImage = My.Resources.search4
        repoCustCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoCustCode.Width = 100
        repoCustCode.IsVisible = True
        repoCustCode.IsPinned = True
        repoCustCode.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoCustCode)
        Dim repoCustPhone As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCustPhone.FormatString = ""
        repoCustPhone.HeaderText = "Phone No"
        repoCustPhone.Name = colCustPhone
        repoCustPhone.Width = 150
        repoCustPhone.IsVisible = True
        repoCustPhone.IsPinned = True
        repoCustPhone.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCustPhone)
        Dim repoRouteNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRouteNo.FormatString = ""
        repoRouteNo.HeaderText = "Route No"
        repoRouteNo.Name = colRouteNo
        repoRouteNo.Width = 150
        repoRouteNo.IsVisible = True
        repoRouteNo.IsPinned = True
        repoRouteNo.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoRouteNo)
        Dim repoSetZero As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoSetZero = New GridViewDecimalColumn()
        repoSetZero.FormatString = ""
        repoSetZero.HeaderText = "Set Zero"
        repoSetZero.Name = colSetZero
        repoSetZero.Width = 50
        repoSetZero.ReadOnly = False
        repoSetZero.IsPinned = True
        'repoSetZero.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gv1.MasterTemplate.Columns.Add(repoSetZero)
        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        qry = "select * from (select 'Fresh' as FreshAmbient,tspl_item_master.Item_Code,tspl_item_master.Short_Description ,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description as ItemDescNew,1 as RowNo,tspl_item_master.Sku_Seq   from tspl_item_master 
    left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_FreshItem =1 AND TSPL_ITEM_MASTER.Is_Milk_Pouch =1 and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0  and Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
    and TSPL_ITEM_UOM_DETAIL.Uom_code in ('" + UOMCrate + "','" + UOMPouch + "','" + UOMLtr + "')
    union
    select 'Fresh' as FreshAmbient,tspl_item_master.Item_Code ,tspl_item_master.Short_Description,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description as ItemDescNew,1 as RowNo,tspl_item_master.Sku_Seq  from tspl_item_master 
    left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_FreshItem =1 AND TSPL_ITEM_MASTER.Is_Milk_Pouch =1 and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0  and Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
    and TSPL_ITEM_UOM_DETAIL.Uom_code in ('" + UOMCrate + "','" + UOMPouch + "','" + UOMLtr + "')"
        If Not clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            qry += " union all
    Select 'Ambient' as FreshAmbient,tspl_item_master.Item_Code ,tspl_item_master.Short_Description,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description  as ItemDescNew,2 as RowNo,tspl_item_master.Sku_Seq   from tspl_item_master 
    left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_Ambient=1   and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0  and Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
    and TSPL_ITEM_UOM_DETAIL.Default_UOM=1"
        End If
        qry += " )z order by RowNo,Sku_Seq,Item_Code"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim i As Integer = 1
            Dim obj As ItemValueClass = New ItemValueClass()
            For Each dr As DataRow In dt.Rows
                repoIName = New GridViewTextBoxColumn()
                repoIName.FormatString = ""
                repoIName.HeaderText = clsCommon.myCstr(dr("UOM_Code"))
                obj = New ItemValueClass()
                obj.itemCode = clsCommon.myCstr(dr("Item_Code"))
                obj.itemDesc = clsCommon.myCstr(dr("Item_Desc"))
                obj.Unit_code = clsCommon.myCstr(dr("UOM_Code"))
                obj.IsFreshAmbient = clsCommon.myCstr(dr("FreshAmbient"))
                obj.ShortDesc = clsCommon.myCstr(dr("Short_Description"))
                repoIName.Tag = obj
                repoIName.Name = colItemCode + clsCommon.myCstr(i)
                repoIName.Width = 150
                repoIName.IsVisible = True
                i = i + 1
                gv1.MasterTemplate.Columns.Add(repoIName)
            Next
        End If
        Dim repoCrate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCrate = New GridViewDecimalColumn()
        repoCrate.FormatString = ""
        repoCrate.HeaderText = "Crate"
        repoCrate.Name = colCrate
        repoCrate.Width = 80
        repoCrate.Minimum = 0
        repoCrate.ReadOnly = True
        repoCrate.IsVisible = True
        repoCrate.IsPinned = True
        repoCrate.PinPosition = PinnedColumnPosition.Right
        gv1.MasterTemplate.Columns.Add(repoCrate)
        gv1.BestFitColumns()
        gv1.Rows.AddNew()
        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.EnableFiltering = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
        View()
    End Sub
    Sub View()
        Try
            If gv1.Rows.Count > 0 Then
                Dim view As New ColumnGroupsViewDefinition()
                view.ColumnGroups.Add(New GridViewColumnGroup("Booth"))
                view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colLineNo).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colCustCode).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colCustPhone).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colRouteNo).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colSetZero).Name)
                view.ColumnGroups(0).IsPinned = True
                Dim TempColGroupCount As Integer = 1
                Dim obj As ItemValueClass = New ItemValueClass()
                Dim i As Integer = 1
                For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
                    If obj1 IsNot Nothing Then
                        If clsCommon.CompairString(obj1.IsFreshAmbient, "Fresh") = CompairStringResult.Equal Then
                            view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(obj1.ShortDesc)))
                            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                            If QuickDemandUOMCrate AndAlso QuickDemandUOMPouch AndAlso QuickDemandUOMLtr Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                                dblcolumns = dblcolumns + 1
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                                dblcolumns = dblcolumns + 1
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMCrate AndAlso QuickDemandUOMPouch Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                                dblcolumns = dblcolumns + 1
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMPouch AndAlso QuickDemandUOMLtr Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                                dblcolumns = dblcolumns + 1
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMCrate AndAlso QuickDemandUOMLtr Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                                dblcolumns = dblcolumns + 1
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMCrate Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMPouch Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            ElseIf QuickDemandUOMLtr Then
                                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            End If
                            TempColGroupCount += 1
                        Else
                            view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(obj1.ShortDesc)))
                            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            TempColGroupCount += 1
                        End If
                    End If
                Next
                view.ColumnGroups.Add(New GridViewColumnGroup("Total"))
                view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colCrate).Name)
                view.ColumnGroups(TempColGroupCount).IsPinned = True
                view.ColumnGroups(TempColGroupCount).PinPosition = PinnedColumnPosition.Right
                'MergeHorizontally(gv1, 0, gv1.Rows.Count - 1)
                gv1.ViewDefinition = view
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    Sub AddNew()
        If QuickDemandUOMCrate Then
            UOMCrate = "Crate"
        End If
        If QuickDemandUOMPouch Then
            UOMPouch = "Pouch"
        End If
        If QuickDemandUOMLtr Then
            UOMLtr = "LTR"
        End If
        LoadBlankGrid()
        btnExport.Visible = False
        Dim CurrDateTime As DateTime = clsCommon.GETSERVERDATE
        Dim EndTime As DateTime = clsCommon.GetPrintDate(SetShiftTimeOut, "dd/MMM/yyyy hh:mm tt")
        If CurrDateTime.TimeOfDay < EndTime.TimeOfDay Then
            txtDate.Value = clsCommon.GetPrintDate(CurrDateTime)
            cmbShift.SelectedValue = "Evening"
        Else
            txtDate.Value = clsCommon.GetPrintDate(CurrDateTime.AddDays(1))
            cmbShift.SelectedValue = "Morning"
        End If
        txtBoothName.Text = ""
        txtDistributor.Text = ""
        txtAddress.Text = ""
        btnSave.Enabled = True
    End Sub
    Sub LoadShiftType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "Morning"
        dr("Name") = "Morning"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Evening"
        dr("Name") = "Evening"
        dt.Rows.Add(dr)
        cmbShift.DataSource = dt
        cmbShift.ValueMember = "Code"
        cmbShift.DisplayMember = "Name"
    End Sub
    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub
    Private Sub gv1_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If Not isInsideLoadData Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.Name = colCustCode Then
                        GenerateLineNo()
                        Dim qry As String = "select Cust_Code,Status from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "'"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            If Not clsCommon.CompairString(dt.Rows(0)("Status"), "N") = CompairStringResult.Equal Then
                                qry = clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value)
                                Dim rt As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + qry + "'"))
                                'gv1.CurrentRow.Cells(colCustCode).Value = ""
                                LoadBlankGrid()
                                Throw New Exception("Inactive Customer [ " + qry + " ], Route No [" + rt + " ]")
                            End If
                        End If
                        gv1.CurrentRow.Cells(colCustCode).Value = clsDistributorRouteTagging.getFinder(" Status='N' and IsDistributor='N' and form_type not in('TPT','VSP') ", clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value), False)
                        Dim isExistingCust As Boolean = FindCustInGrid(gv1.CurrentRow.Cells(colCustCode).Value)
                        gv1.CurrentRow.Cells(colCustPhone).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Phone1 from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "'"))
                        gv1.CurrentRow.Cells(colRouteNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "'"))
                        gv1.CurrentRow.Cells(colSetZero).Value = 1
                        GetBoothDetail()
                        qry = "select top 1 TSPL_DEMAND_BOOKING_MASTER.Document_No from TSPL_DEMAND_BOOKING_MASTER 
where convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") + "' and TSPL_DEMAND_BOOKING_MASTER.ShiftType='" + cmbShift.Text + "' and TSPL_DEMAND_BOOKING_MASTER.Route_No='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colRouteNo).Value) + "' and tspl_demand_booking_master.posted=1 and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=0"
                        Dim isDemandPosted As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                        Dim cust As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRouteNo).Value)
                        If clsCommon.myLen(isDemandPosted) > 0 Then
                            LoadBlankGrid()
                            Throw New Exception("Demand Already Posted for Route_No Code [" + clsCommon.myCstr(cust) + "]")
                        End If
                        FindDemand(gv1.CurrentRow.Cells(colCustCode).Value, gv1.CurrentRow.Cells(colRouteNo).Value)
                        gv1.CurrentRow.Cells(colCustCode).ReadOnly = True
                    End If
                    If e.Column.Name = colSetZero Then
                        UpdateCurrentRow(gv1.CurrentRow.Index)
                    End If
                    If e.Column.Index >= 5 Then
                        UpdateCurrentRow(gv1.CurrentRow.Index)
                        SetGridFocus(True)
                    End If

                    isCellValueChangedOpen = False
                End If
                isInsideLoadData = False
            End If
            'SetGridFocus()
        Catch ex As Exception
            isInsideLoadData = False
            isCellValueChangedOpen = False
            'LoadBlankGrid()
            SetGridFocus(False)
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        Try
            Dim obj As New clsDemandSheet()
            obj.DEMAND_Date = clsCommon.GetPrintDate(txtDate.Value)
            obj.Cust_Code = gv1.Rows(IntRowNo).Cells(colCustCode).Value
            obj.Route_No = gv1.Rows(IntRowNo).Cells(colRouteNo).Value
            obj.Set_Zero = gv1.Rows(IntRowNo).Cells(colSetZero).Value
            obj.ShiftType = cmbShift.SelectedValue
            If gv1.Rows(IntRowNo).Cells(colSetZero).Value = 0 Then
                For dbColumn As Integer = 4 To gv1.Columns.Count - 2
                    gv1.Rows(IntRowNo).Cells(dbColumn).Value = "0"
                    Dim k As Integer = 1
                    For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                        Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                        k = k + 1
                        If obj1 IsNot Nothing Then
                            If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 Then
                                obj.Item_Code = clsCommon.myCstr(obj1.itemCode)
                                obj.Unit_Code = clsCommon.myCstr(obj1.Unit_code)
                                obj.Qty = 0
                                Try
                                    If clsCommon.myLen(obj.Cust_Code) > 0 Then
                                        Dim status As Boolean = obj.SaveData(obj)
                                    End If
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            End If
                        End If
                    Next
                Next
            Else
                If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colCustCode).Value)) > 0 Then
                    Dim dbltotCrate As Decimal = 0
                    Dim dblTotalCrateRowWise As Decimal = 0
                    Dim k As Integer = 1
                    Dim ccount As Integer = 0
                    For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                        Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                        k = k + 1
                        If obj1 IsNot Nothing Then
                            If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 Then
                                obj.Item_Code = clsCommon.myCstr(obj1.itemCode)
                                Dim cellValue As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(dblcolumns).Value)
                                If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                                    If ItemCrateType = 1 Then
                                        Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                        Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                        Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                        Dim ItempouchCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                                        If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                                            Dim DispatchQty As Decimal = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(dblcolumns).Value) * ItemConvFactor
                                            If AllowRouteWiseDemandEntryInDecimal Then
                                                If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & obj.Route_No & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                                    If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                        gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                        Throw New Exception("Decimal values are not allowed.")
                                                    End If
                                                Else
                                                    Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))
                                                    If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                        gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                        Throw New Exception("Please Enter Valid Qty for " + obj1.ShortDesc)
                                                    End If
                                                End If
                                            Else
                                                If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" + obj1.itemCode + "' and  AllowEntryInDecimal=1")) = 0 Then
                                                    If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                        gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                        Throw New Exception("Decimal values are not allowed.")
                                                    End If
                                                Else
                                                    Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))
                                                    If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                        gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                        Throw New Exception("Please Enter Valid Qty for " + obj1.ShortDesc)
                                                    End If
                                                End If
                                            End If

                                            If ConvertPouchtoCrate Then
                                                If DispatchQty > (CrateConvFactor / 2) Then
                                                    dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                                Else
                                                    dblTotalCrateRowWise = 1
                                                End If
                                            ElseIf DontCreateForPouch Then
                                                dblTotalCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                            Else
                                                If DispatchQty > (CrateConvFactor / 2) Then
                                                    dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                                Else
                                                    dblTotalCrateRowWise = 0
                                                End If
                                            End If
                                        End If
                                    End If
                                    dbltotCrate = dbltotCrate + dblTotalCrateRowWise
                                    gv1.Rows(IntRowNo).Cells(colCrate).Value = dbltotCrate
                                Else
                                    Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                    Dim CrateConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' "))
                                    Dim ItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                    If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                                        Dim DispatchQty As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(dblcolumns).Value) * ItemConvFactor
                                        If AllowRouteWiseDemandEntryInDecimal Then
                                            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & obj.Route_No & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                                If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                    gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                    Throw New Exception("Decimal values are not allowed.")
                                                End If
                                            Else
                                                Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))
                                                If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                    gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                    Throw New Exception("Please Enter Valid Qty for " + obj1.ShortDesc)
                                                End If
                                            End If
                                        Else
                                            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" + obj1.itemCode + "' and  AllowEntryInDecimal=1")) = 0 Then
                                                If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                    gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                    Throw New Exception("Decimal values are not allowed.")
                                                End If
                                            Else
                                                Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))
                                                If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                    gv1.Rows(IntRowNo).Cells(dblcolumns).Value = ""
                                                    Throw New Exception("Please Enter Valid Qty for " + obj1.ShortDesc)
                                                End If
                                            End If
                                        End If

                                    End If
                                End If
                                obj.Qty = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(dblcolumns).Value)
                                obj.Unit_Code = clsCommon.myCstr(obj1.Unit_code)
                                Try
                                    Dim status As Boolean = obj.SaveData(obj)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                                ccount += obj.Qty
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    'Public Sub LoadData(ByVal CurrDate As Date, ByVal Shift As String, ByVal CurrUser As String, ByVal isSummary As Boolean)
    '    Try
    '        LoadBlankGrid()
    '        isInsideLoadData = True
    '        Dim IntRowNo As Decimal = 0
    '        Dim ccount As Integer = 0
    '        Dim lstobj As List(Of clsDemandSheet)
    '        lstobj = clsDemandSheet.GetData(CurrDate, Shift, objCommonVar.CurrentUserCode, Nothing)
    '        If (lstobj IsNot Nothing AndAlso lstobj.Count > 0) Then
    '            For Each obj As clsDemandSheet In lstobj
    '                gv1.Rows(IntRowNo).Cells(colLineNo).Value = IntRowNo + 1
    '                gv1.Rows(IntRowNo).Cells(colCustCode).Value = obj.Cust_Code
    '                gv1.Rows(IntRowNo).Cells(colCustPhone).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Phone1 from TSPL_CUSTOMER_MASTER where Cust_Code='" + obj.Cust_Code + "'"))
    '                'gv1.Rows(IntRowNo).Cells(colSetZero).Value = obj.Set_Zero
    '                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
    '                    For Each objTr As clsDemandSheetDetails In obj.Arr
    '                        gv1.Rows(IntRowNo).Cells(colRouteNo).Value = objTr.Route_No
    '                        gv1.Rows(IntRowNo).Cells(colSetZero).Value = objTr.Set_Zero
    '                        'For dblrows As Integer = 0 To gv1.Rows.Count - 1
    '                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colCustCode).Value), objTr.Cust_Code) = CompairStringResult.Equal Then
    '                            Dim k As Integer = 1
    '                            For columns = 5 To gv1.Columns.Count - 2
    '                                Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
    '                                k = k + 1
    '                                If clsCommon.CompairString(objTr.Item_Code, clsCommon.myCstr(obj1.itemCode)) = CompairStringResult.Equal Then
    '                                    gv1.Rows(IntRowNo).Cells(columns).Value = objTr.Qty
    '                                    'gv1.Rows(IntRowNo).Cells(colCrate).Value = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colCrate).Value) + objTr.Qty
    '                                End If
    '                            Next
    '                        End If
    '                        ' Next
    '                        'ccount += objTr.Qty
    '                        CountTotal()
    '                    Next
    '                End If
    '                IntRowNo += 1
    '                'gv1.Rows(IntRowNo).Cells(colCrate).Value = ccount
    '                gv1.Rows.AddNew()
    '            Next
    '            '
    '        End If
    '        'GvRowFridge()
    '        If isSummary Then
    '            Dim summaryRowItem As New GridViewSummaryRowItem()
    '            For colcount As Integer = 5 To gv1.Columns.Count - 2
    '                gv1.Columns(colItemCode + clsCommon.myCstr(colcount - 4)).FormatString = "{0:n2}"
    '            Next
    '            For colcount As Integer = 5 To gv1.Columns.Count - 2
    '                Dim TotalCount As New GridViewSummaryItem(colItemCode + clsCommon.myCstr(colcount - 4), "{0:n2}", GridAggregateFunction.Sum)
    '                summaryRowItem.Add(TotalCount)
    '            Next
    '            gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
    '            gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    '            gv1.AutoSizeRows = False
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            If gv1.Rows.Count > 0 Then
                If clsCommon.myLen(gv1.Rows(0).Cells(colCustCode).Value) > 0 Then
                    SetDemandBooking(clsCommon.myCstr(gv1.Rows(0).Cells(colCustCode).Value), txtDate.Value, cmbShift.Text, 0)
                    LoadBlankGrid()
                Else
                    Throw New Exception("Please fill data at line no 1.")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub SaveData()
        Try
            For Each grow As GridViewRowInfo In gv1.Rows
                If grow.Cells(colCustCode).Value <> "" Then

                    Dim document As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_NO from TSPL_DEMAND_BOOKING_MASTER where convert(date,Document_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and ShiftType='" + clsCommon.myCstr(cmbShift.SelectedValue) + "' and Route_No='" + clsCommon.myCstr(grow.Cells(colRouteNo).Value) + "' and IsIndividualCustomer=0"))
                    If clsCommon.myLen(document) > 0 Then
                        Dim k As Integer = 1
                        For columns = 5 To gv1.Columns.Count - 2
                            Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                            k = k + 1
                            Dim ExistsItem As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + document + "' and Cust_Code='" + clsCommon.myCstr(grow.Cells(colCustCode).Value) + "' and Item_Code='" + clsCommon.myCstr(obj1.itemCode) + "'"))
                            Dim UpdateQry As String = ""
                            If ExistsItem > 0 Then
                                If clsCommon.myCdbl(grow.Cells(columns).Value) > 0 Then
                                    UpdateQry = " update TSPL_DEMAND_BOOKING_DETAIL set Qty=" + clsCommon.myCstr(grow.Cells(columns).Value) + " where Document_No='" + document + "' and Cust_Code='" + clsCommon.myCstr(grow.Cells(colCustCode).Value) + "' and Item_Code='" + clsCommon.myCstr(obj1.itemCode) + "'"
                                    clsDBFuncationality.ExecuteNonQuery(UpdateQry)
                                Else
                                    Dim TRCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TR_Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + document + "' and Cust_Code='" + clsCommon.myCstr(grow.Cells(colCustCode).Value) + "' and Item_Code='" + clsCommon.myCstr(obj1.itemCode) + "'"))
                                    UpdateQry = "delete TSPL_BOOKING_DETAIL  where Against_DemandBooking_TR_Code='" + TRCode + "' and Cust_Code='" + clsCommon.myCstr(grow.Cells(colCustCode).Value) + "' and Item_Code='" + clsCommon.myCstr(obj1.itemCode) + "'"
                                    clsDBFuncationality.ExecuteNonQuery(UpdateQry)
                                    UpdateQry = " Delete TSPL_DEMAND_BOOKING_DETAIL  where Document_No='" + document + "' and Cust_Code='" + clsCommon.myCstr(grow.Cells(colCustCode).Value) + "' and Item_Code='" + clsCommon.myCstr(obj1.itemCode) + "'"
                                    clsDBFuncationality.ExecuteNonQuery(UpdateQry)
                                End If
                            Else
                                If clsCommon.myCdbl(grow.Cells(columns).Value) > 0 Then
                                    Dim objUDD As New clsUpdateDemandDetails()
                                    objUDD.Document_No = document
                                    objUDD.Cust_Code = grow.Cells(colCustCode).Value
                                    objUDD.ShiftType = cmbShift.SelectedValue
                                    objUDD.Item_Code = obj1.itemCode
                                    objUDD.Unit_code = obj1.Unit_code
                                    objUDD.Qty = grow.Cells(columns).Value
                                    objUDD.Price_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER where Cust_Code='" & clsCommon.myCstr(objUDD.Cust_Code) & "'"))
                                    objUDD.Line_No = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_DEMAND_BOOKING_DETAIL  where Document_No='" + objUDD.Document_No + "'")) + 1
                                    objUDD.Rate = obj1.ItemRate
                                    objUDD.Vehicle_Code = clsDBFuncationality.getSingleValue("select vehicle_code from TSPL_ROUTE_MASTER where Route_No='" + grow.Cells(colRouteNo).Value + "'")
                                    Dim status As Boolean = clsUpdateDemandDetails.InsertNewItem(objUDD, clsCommon.GetPrintDate(txtDate.Value))
                                End If
                                'clsCommon.MyMessageBoxShow(Me, UpdateQry, Me.Text)
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub SetGridFocus(ByVal isSave As Boolean)
        If gv1.CurrentCell IsNot Nothing Then
            Dim setNxtRow As Boolean = False
            Dim k As Integer = gv1.CurrentColumn.Index
            Dim r As Integer = 0
            If gv1.CurrentCell.ColumnInfo.Name = colCustCode Then
                gv1.CurrentColumn = gv1.Columns(colItemCode + clsCommon.myCstr(1))
            ElseIf gv1.CurrentCell.ColumnInfo.Name = colCustPhone Then
                gv1.CurrentColumn = gv1.Columns(colItemCode + clsCommon.myCstr(1))
            ElseIf gv1.CurrentCell.ColumnInfo.Name = colSetZero Then
                gv1.CurrentColumn = gv1.Columns(colItemCode + clsCommon.myCstr(1))
            ElseIf k >= 5 AndAlso k < gv1.Columns.Count - 2 Then
                If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(gv1.CurrentRow.Index).Cells(colCustCode).Value)) > 0 Then
                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k - 4)).Tag, ItemValueClass)
                    k = k - 3
                    r = k
                    If obj1 IsNot Nothing Then
                        If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso r <= gv1.Columns.Count - 5 Then  'AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(dblcolumns).Value) > 0
                            If gv1.CurrentCell.ColumnInfo.Name = colItemCode + clsCommon.myCstr(k - 1) Then
                                Dim obj2 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(r)).Tag, ItemValueClass)
                                If obj2 IsNot Nothing Then
                                    gv1.CurrentColumn = gv1.Columns(colItemCode + clsCommon.myCstr(r))
                                    Exit Sub
                                Else
                                    setNxtRow = True
                                End If
                            End If
                        Else
                            setNxtRow = True
                        End If
                    End If
                    If setNxtRow Then
                        'SetDemandBooking(clsCommon.myCstr(gv1.Rows(gv1.CurrentRow.Index).Cells(colCustCode).Value), txtDate.Value, cmbShift.Text, gv1.CurrentRow.Index)
                        If gv1.Rows.Count > gv1.CurrentRow.Index + 1 Then
                            gv1.CurrentRow = gv1.Rows(gv1.CurrentRow.Index + 1)
                        End If
                        gv1.CurrentColumn = gv1.Columns(colCustCode)
                        'GvRowFridge()
                    End If
                Else
                    clsCommon.MyMessageBoxShow(Me, "Please select Booth")
                    LoadBlankGrid()
                End If
            Else
                'FridgeCurrentRow(gv1.CurrentRow.Index)
                'If gv1.Rows.Count > gv1.CurrentRow.Index + 1 Then
                '    gv1.CurrentRow = gv1.Rows(gv1.CurrentRow.Index + 1)
                'End If
                If isSave OrElse Not isLastcellvaluechanged Then
                    'UpdateCurrentRow(0)
                    If clsCommon.myLen(gv1.Rows(0).Cells(colCustCode).Value) > 0 Then
                        SetDemandBooking(clsCommon.myCstr(gv1.Rows(0).Cells(colCustCode).Value), txtDate.Value, cmbShift.Text, 0)
                    End If
                    gv1.CurrentColumn = gv1.Columns(colCustCode)

                End If

            End If
        End If
    End Sub
    Public Sub SetDemandBooking(ByVal strCustCode As String, ByVal DemandData As DateTime, ByVal strShiftType As String, ByVal intRow As Integer)
        Try
            Dim qry As String = String.Empty
            Dim RouteNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + strCustCode + "'"))
            If clsCommon.myLen(RouteNo) > 0 Then
                qry = "Select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No = '" + RouteNo + "' and  CONVERT(varchar, CAST(Document_Date AS datetime), 103) ='" & clsCommon.GetPrintDate(DemandData, "dd/MM/yyyy") & "' and ShiftType = '" & strShiftType & "' and IsIndividualCustomer=0"
            End If
            Dim DocumentNO As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            Dim lineNo As Integer = 1
            If clsCommon.myLen(DocumentNO) > 0 Then
                Dim location_Code As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + DocumentNO + "'"))
                qry = "select * from TSPL_DEMAND_SHEET where Cust_Code='" + strCustCode + "' and convert(date,demand_date,103)='" + clsCommon.GetPrintDate(DemandData, "dd/MMM/yyyy") + "' and ShiftType='" + strShiftType + "' and Created_By='" + objCommonVar.CurrentUserCode + "'"
                Dim dtDS As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dtDS IsNot Nothing AndAlso dtDS.Rows.Count > 0 Then
                    Dim DObj As New List(Of clsDemandBookingSaleDetail)
                    For Each dr As DataRow In dtDS.Rows
                        Dim objDBD As New clsDemandBookingSaleDetail
                        objDBD.Line_No = lineNo
                        objDBD.Document_No = DocumentNO
                        objDBD.Cust_Code = strCustCode
                        objDBD.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                        objDBD.Item_Desc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Short_Description from TSPL_ITEM_MASTER where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'"))
                        objDBD.Unit_code = clsCommon.myCstr(dr("Unit_Code"))
                        objDBD.Qty = clsCommon.myCdbl(dr("Qty"))
                        objDBD.REF_PK_ID = clsCommon.myCstr(clsCommon.myCdbl(dr("PK_ID")))
                        objDBD.Price_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER where Cust_Code='" + strCustCode + "'"))
                        objDBD.ShiftType = strShiftType
                        objDBD.Vehicle_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vehicle_No from TSPL_VEHICLE_MASTER left join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where TSPL_ROUTE_MASTER.Route_No='" + RouteNo + "'"))
                        objDBD.Trip_No = 1
                        objDBD.Source_By = "ERP"
                        objDBD.TotalCrates_ItemWise = clsCommon.myCdbl(dr("Qty"))
                        Dim CrateConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(objDBD.Item_Code) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='Ltr' "))
                        Dim ItemConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(objDBD.Item_Code) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(objDBD.Unit_code) & "' "))
                        If CrateConvFactor_Ltr > 0 And ItemConvFactor_Ltr > 0 Then
                            Dim DispatchQty As Double = clsCommon.myCdbl(objDBD.Qty) * ItemConvFactor_Ltr
                            objDBD.TotalLtr_ItemWise = (DispatchQty / CrateConvFactor_Ltr)
                        End If
                        qry = " Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, XXXE.Tax_group,XXXE.TAX1_Rate, " &
                        " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
                        "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
                        " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
                        " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
                        " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
                        " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt, XXXE.TAX5_Amt,XXXE.TAX6_Amt,XXXE.TAX7_Amt,XXXE.TAX8_Amt,XXXE.TAX9_Amt,XXXE.TAX10_Amt,XXXE.Against_Plan_TR_Code  from ( " &
                        "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                        "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
                        "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, TSPL_ITEM_PRICE_MASTER.Tax_group,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
                        "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
                        " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.TAX5_Amt,TSPL_ITEM_PRICE_MASTER.TAX6_Amt,TSPL_ITEM_PRICE_MASTER.TAX7_Amt,   TSPL_ITEM_PRICE_MASTER.TAX8_Amt,TSPL_ITEM_PRICE_MASTER.TAX9_Amt,TSPL_ITEM_PRICE_MASTER.TAX10_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                        "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                        "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(clsCommon.CompairString(clsCommon.myCstr(cmbShift.SelectedValue), "Morning") = CompairStringResult.Equal, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + IIf(clsCommon.CompairString(clsCommon.myCstr(cmbShift.SelectedValue), "Morning") = CompairStringResult.Equal, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end)  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
                        "TSPL_ITEM_PRICE_MASTER.Price_Code='" & objDBD.Price_Code & "' and UOM='" & objDBD.Unit_code & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & objDBD.Item_Code & "' AND Location_Code='" & clsCommon.myCstr(location_Code) & "'" &
                        ") XXXE WHERE RowNo=1  "
                        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt2.Rows.Count > 0 Then
                            Dim dblRate As Decimal = clsCommon.myCDecimal(dt2.Rows(0).Item("Item_Basic_Price"))
                            If dblRate = 0 Then
                                Throw New Exception("Please Fill Selling Price for Location " & location_Code & "  for item " & clsCommon.myCstr(objDBD.Item_Desc) & Environment.NewLine)
                            End If
                            objDBD.Rate = dblRate
                            objDBD.ItemNetAmount = Math.Round(clsCommon.myCdbl(objDBD.Qty) * clsCommon.myCdbl(dblRate), 2)
                            objDBD.TAX_Group = clsCommon.myCstr(dt2.Rows(0).Item("TAX_Group"))
                            objDBD.TAX1 = clsCommon.myCstr(dt2.Rows(0).Item("TAX1"))
                            objDBD.TAX1_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX1_Rate"))
                            objDBD.TAX1_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX1_Rate / 100), 2)
                            objDBD.TAX1_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX2 = clsCommon.myCstr(dt2.Rows(0).Item("TAX2"))
                            objDBD.TAX2_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX2_Rate"))
                            objDBD.TAX2_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX2_Rate / 100), 2)
                            objDBD.TAX2_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX3 = clsCommon.myCstr(dt2.Rows(0).Item("TAX3"))
                            objDBD.TAX3_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX3_Rate"))
                            objDBD.TAX3_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX3_Rate / 100), 2)
                            objDBD.TAX3_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX4 = clsCommon.myCstr(dt2.Rows(0).Item("TAX4"))
                            objDBD.TAX4_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX4_Rate"))
                            objDBD.TAX4_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX4_Rate / 100), 2)
                            objDBD.TAX4_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX5 = clsCommon.myCstr(dt2.Rows(0).Item("TAX5"))
                            objDBD.TAX5_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX5_Rate"))
                            objDBD.TAX5_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX5_Rate / 100), 2)
                            objDBD.TAX5_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX6 = clsCommon.myCstr(dt2.Rows(0).Item("TAX6"))
                            objDBD.TAX6_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX6_Rate"))
                            objDBD.TAX6_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX6_Rate / 100), 2)
                            objDBD.TAX6_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX7 = clsCommon.myCstr(dt2.Rows(0).Item("TAX7"))
                            objDBD.TAX7_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX7_Rate"))
                            objDBD.TAX7_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX7_Rate / 100), 2)
                            objDBD.TAX7_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX8 = clsCommon.myCstr(dt2.Rows(0).Item("TAX8"))
                            objDBD.TAX8_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX8_Rate"))
                            objDBD.TAX8_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX8_Rate / 100), 2)
                            objDBD.TAX8_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX9 = clsCommon.myCstr(dt2.Rows(0).Item("TAX9"))
                            objDBD.TAX9_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX9_Rate"))
                            objDBD.TAX9_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX9_Rate / 100), 2)
                            objDBD.TAX9_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX10 = clsCommon.myCstr(dt2.Rows(0).Item("TAX10"))
                            objDBD.TAX10_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX10_Rate"))
                            objDBD.TAX10_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX10_Rate / 100), 2)
                            objDBD.TAX10_Base_Amt = objDBD.ItemNetAmount
                            DObj.Add(objDBD)
                        End If
                        lineNo += 1
                    Next
                    Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                    Try
                        Dim obj As clsDemandBookingSale = clsDemandBookingSale.GetData(DocumentNO, NavigatorType.Current, trans)
                        qry = "delete from TSPL_DEMAND_BOOKING_DETAIL where tr_code in (select tr_code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + obj.Document_No + "' and Cust_Code='" + strCustCode + "') "
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        clsDemandBookingSaleDetail.SaveData(DocumentNO, DemandData, DObj, trans, location_Code, strShiftType, True, False, RouteNo)
                        clsDemandBookingSale.SaveDemandHistoryData(obj, DObj, "Save Quick Demand", "ERP", objCommonVar.CurrentUserCode, trans)
                        trans.Commit()
                        clsCommon.MyMessageBoxShow(Me, "Saved Successfully", Me.Text)
                        LoadBlankGrid()
                    Catch ex As Exception
                        trans.Rollback()
                        Throw New Exception(ex.Message)
                    End Try
                End If
            Else
                Dim dblTotalAmount As Double = 0
                Dim dblTotalCrate As Double = 0
                Dim dblTotalltr As Double = 0
                Dim DBObj As New clsDemandBookingSale
                DBObj.Route_No = RouteNo
                DBObj.Document_Date = DemandData
                DBObj.Location_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Route_Master where Route_No='" + RouteNo + "' "))
                DBObj.IsIndividualCustomer = 0
                DBObj.City_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select City_Code from TSPL_ROUTE_MASTER where Route_No='" + RouteNo + "'"))
                DBObj.ItemType = "Fresh"
                DBObj.ShiftType = strShiftType
                DBObj.Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER where Cust_Code='" + strCustCode + "'"))
                DBObj.TripNo = 1
                DBObj.Arr = New List(Of clsDemandBookingSaleDetail)
                qry = "select * from TSPL_DEMAND_SHEET where Cust_Code='" + strCustCode + "' and convert(date,demand_date,103)='" + clsCommon.GetPrintDate(DemandData, "dd/MMM/yyyy") + "' and ShiftType='" + strShiftType + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim objDBD As New clsDemandBookingSaleDetail
                        objDBD.Cust_Code = strCustCode
                        objDBD.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                        objDBD.Item_Desc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Short_Description from TSPL_ITEM_MASTER where Item_Code='" + clsCommon.myCstr(dr("Item_Code")) + "'"))
                        objDBD.Unit_code = clsCommon.myCstr(dr("Unit_Code"))
                        objDBD.Qty = clsCommon.myCdbl(dr("Qty"))
                        objDBD.REF_PK_ID = clsCommon.myCstr(clsCommon.myCdbl(dr("PK_ID")))
                        objDBD.Price_Code = DBObj.Price_code
                        objDBD.ShiftType = strShiftType
                        objDBD.Vehicle_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vehicle_No from TSPL_VEHICLE_MASTER left join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where TSPL_ROUTE_MASTER.Route_No='" + RouteNo + "'"))
                        objDBD.Line_No = lineNo
                        objDBD.Trip_No = 1
                        objDBD.TotalCrates_ItemWise = clsCommon.myCdbl(dr("Qty"))
                        Dim CrateConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(objDBD.Item_Code) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='Ltr' "))
                        Dim ItemConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(objDBD.Item_Code) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(objDBD.Unit_code) & "' "))
                        If CrateConvFactor_Ltr > 0 And ItemConvFactor_Ltr > 0 Then
                            Dim DispatchQty As Double = clsCommon.myCdbl(objDBD.Qty) * ItemConvFactor_Ltr
                            objDBD.TotalLtr_ItemWise = (DispatchQty / CrateConvFactor_Ltr)
                        End If
                        qry = " Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, XXXE.Tax_group,XXXE.TAX1_Rate, " &
                        " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
                        "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
                        " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
                        " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
                        " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
                        " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt, XXXE.TAX5_Amt,XXXE.TAX6_Amt,XXXE.TAX7_Amt,XXXE.TAX8_Amt,XXXE.TAX9_Amt,XXXE.TAX10_Amt,XXXE.Against_Plan_TR_Code  from ( " &
                        "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                        "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
                        "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, TSPL_ITEM_PRICE_MASTER.Tax_group,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
                        "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
                        " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
                        " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.TAX5_Amt,TSPL_ITEM_PRICE_MASTER.TAX6_Amt,TSPL_ITEM_PRICE_MASTER.TAX7_Amt,   TSPL_ITEM_PRICE_MASTER.TAX8_Amt,TSPL_ITEM_PRICE_MASTER.TAX9_Amt,TSPL_ITEM_PRICE_MASTER.TAX10_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                        "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                        "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
                        "TSPL_ITEM_PRICE_MASTER.Price_Code='" & objDBD.Price_Code & "' and UOM='" & objDBD.Unit_code & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & objDBD.Item_Code & "' AND Location_Code='" & clsCommon.myCstr(DBObj.Location_Code) & "'  " &
                        ") XXXE WHERE RowNo=1  "
                        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt2.Rows.Count > 0 Then
                            Dim dblRate As Decimal = clsCommon.myCDecimal(dt2.Rows(0).Item("Item_Basic_Price"))
                            If dblRate = 0 Then
                                Throw New Exception("Please Fill Selling Price for Location " & DBObj.Location_Code & "  for item " & clsCommon.myCstr(objDBD.Item_Desc) & Environment.NewLine)
                            End If
                            objDBD.Rate = dblRate
                            objDBD.ItemNetAmount = Math.Round(clsCommon.myCdbl(objDBD.Qty) * clsCommon.myCdbl(dblRate), 2)
                            objDBD.TAX_Group = clsCommon.myCstr(dt2.Rows(0).Item("TAX_Group"))
                            objDBD.TAX1 = clsCommon.myCstr(dt2.Rows(0).Item("TAX1"))
                            objDBD.TAX1_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX1_Rate"))
                            objDBD.TAX1_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX1_Rate / 100), 2)
                            objDBD.TAX1_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX2 = clsCommon.myCstr(dt2.Rows(0).Item("TAX2"))
                            objDBD.TAX2_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX2_Rate"))
                            objDBD.TAX2_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX2_Rate / 100), 2)
                            objDBD.TAX2_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX3 = clsCommon.myCstr(dt2.Rows(0).Item("TAX3"))
                            objDBD.TAX3_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX3_Rate"))
                            objDBD.TAX3_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX3_Rate / 100), 2)
                            objDBD.TAX3_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX4 = clsCommon.myCstr(dt2.Rows(0).Item("TAX4"))
                            objDBD.TAX4_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX4_Rate"))
                            objDBD.TAX4_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX4_Rate / 100), 2)
                            objDBD.TAX4_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX5 = clsCommon.myCstr(dt2.Rows(0).Item("TAX5"))
                            objDBD.TAX5_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX5_Rate"))
                            objDBD.TAX5_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX5_Rate / 100), 2)
                            objDBD.TAX5_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX6 = clsCommon.myCstr(dt2.Rows(0).Item("TAX6"))
                            objDBD.TAX6_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX6_Rate"))
                            objDBD.TAX6_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX6_Rate / 100), 2)
                            objDBD.TAX6_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX7 = clsCommon.myCstr(dt2.Rows(0).Item("TAX7"))
                            objDBD.TAX7_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX7_Rate"))
                            objDBD.TAX7_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX7_Rate / 100), 2)
                            objDBD.TAX7_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX8 = clsCommon.myCstr(dt2.Rows(0).Item("TAX8"))
                            objDBD.TAX8_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX8_Rate"))
                            objDBD.TAX8_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX8_Rate / 100), 2)
                            objDBD.TAX8_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX9 = clsCommon.myCstr(dt2.Rows(0).Item("TAX9"))
                            objDBD.TAX9_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX9_Rate"))
                            objDBD.TAX9_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX9_Rate / 100), 2)
                            objDBD.TAX9_Base_Amt = objDBD.ItemNetAmount
                            objDBD.TAX10 = clsCommon.myCstr(dt2.Rows(0).Item("TAX10"))
                            objDBD.TAX10_Rate = clsCommon.myCdbl(dt2.Rows(0).Item("TAX10_Rate"))
                            objDBD.TAX10_Amt = Math.Round(objDBD.ItemNetAmount * (objDBD.TAX10_Rate / 100), 2)
                            objDBD.TAX10_Base_Amt = objDBD.ItemNetAmount
                            DBObj.Arr.Add(objDBD)
                            dblTotalAmount += objDBD.ItemNetAmount
                            dblTotalCrate += objDBD.Qty
                            dblTotalltr += objDBD.TotalLtr_ItemWise
                        End If
                        lineNo += 1
                    Next
                    DBObj.TotalQtyInCrates = dblTotalCrate
                    DBObj.DocumentAmount = dblTotalAmount
                    DBObj.TotalQtyInLtr = dblTotalltr
                    If DBObj.SaveData(DBObj, True) Then
                        clsCommon.MyMessageBoxShow(Me, "Saved Successfully", Me.Text)
                        LoadBlankGrid()
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub CountTotal()
        Dim Ccount As Integer = 0
        If gv1.Rows.Count > 0 Then
            For ii As Integer = 5 To gv1.Columns.Count - 2
                Ccount += gv1.CurrentRow.Cells(ii).Value
            Next
            gv1.CurrentRow.Cells(colCrate).Value = Ccount
        End If
    End Sub
    Private Sub frmDemand_Sheet_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                SetGridFocus(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub FridgeCurrentRow(ByVal intRow As Integer)
        Try
            For colcount As Integer = 0 To gv1.Columns.Count - 2
                gv1.Rows(intRow).Cells(colcount).ReadOnly = True
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub GvRowFridge()
        Try
            If gv1.Rows.Count > 2 Then
                For rowcount As Integer = 0 To gv1.Rows.Count - 3
                    For colcount As Integer = 0 To gv1.Columns.Count - 2
                        gv1.Rows(rowcount).Cells(colcount).ReadOnly = True
                    Next
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        Export(EnumExportTo.Excel)
    End Sub
    Private Sub Export(ByVal exporter As EnumExportTo)
        Try
            If gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            If exporter = EnumExportTo.Excel Then
                transportSql.applyExportTemplate(gv1, PageSetupReport_ID)
                transportSql.exportdata(gv1, "", Me.Text, , Nothing, False, False, False)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmiImport_Click(sender As Object, e As EventArgs) Handles rmiImport.Click
        Dim gv As New UserControls.MyRadGridView
        Dim arr As New List(Of Integer)
        Try
            Dim obj As clsDemandSheet = New clsDemandSheet()
            Dim DemandDate As DateTime = Nothing
            Dim Shift As String = Nothing
            Dim CurrDateTime As DateTime = clsCommon.GETSERVERDATE
            Dim EndTime As DateTime = clsCommon.GetPrintDate(SetShiftTimeOut, "dd/MMM/yyyy hh:mm tt")
            If CurrDateTime.TimeOfDay < EndTime.TimeOfDay Then
                DemandDate = clsCommon.GetPrintDate(CurrDateTime)
                Shift = "Evening"
            Else
                DemandDate = clsCommon.GetPrintDate(CurrDateTime.AddDays(1))
                Shift = "Morning"
            End If
            Dim strItems As New List(Of String)
            strItems.Add("Line No")
            strItems.Add("Booth Code")
            strItems.Add("Phone No")
            strItems.Add("Route No")
            strItems.Add("Set Zero")
            Me.Controls.Add(gv)
            Dim strQry As String = " select Short_Description from TSPL_ITEM_MASTER where Is_DisplayDemand=1 order by Sku_Seq "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing And dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    strItems.Add(clsCommon.myCstr(dr("Short_Description")))
                Next
            Else
            End If
            If transportSql.importExcel(gv, strItems.ToArray()) Then
                For Each grow As GridViewRowInfo In gv.Rows
                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells(1).Value)) > 0 Then
                        obj.DEMAND_Date = clsCommon.GetPrintDate(DemandDate)
                        obj.ShiftType = Shift
                        obj.Cust_Code = clsCommon.myCstr(grow.Cells(1).Value)
                        obj.Route_No = clsCommon.myCstr(grow.Cells(3).Value)
                        obj.Set_Zero = clsCommon.myCdbl(grow.Cells(4).Value)
                        Dim NoofItems As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(Short_Description) from TSPL_ITEM_MASTER where Is_DisplayDemand=1"))
                        For ii As Integer = 1 To NoofItems
                            Dim ItemCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Item_Code from tspl_item_master where Short_Description='" + clsCommon.myCstr(strItems(ii + 4)) + "'and Is_DisplayDemand=1  order by Sku_Seq"))
                            obj.Item_Code = ItemCode
                            obj.Qty = clsCommon.myCdbl(grow.Cells(ii + 4).Value)
                            If clsCommon.myCdbl(grow.Cells(4).Value) = 0 Then
                                obj.Qty = 0
                                Try
                                    Dim status As Boolean = obj.SaveData(obj)
                                Catch ex As Exception
                                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                                End Try
                            ElseIf clsCommon.myCdbl(grow.Cells(ii + 4).Value) > 0 Then
                                obj.Qty = clsCommon.myCdbl(grow.Cells(ii + 4).Value)
                                Try
                                    Dim status As Boolean = obj.SaveData(obj)
                                Catch ex As Exception
                                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                                End Try
                            End If
                        Next
                    Else
                        Throw New Exception("Booth Code Not Found at Line no" + clsCommon.myCstr(grow.Cells(0).Value))
                    End If
                Next
                common.clsCommon.MyMessageBoxShow(Me, "Data Uploaded Successfully", Me.Text)
                'LoadData(clsCommon.GetPrintDate(DemandDate), Shift, objCommonVar.CurrentUserCode, False)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub FindDemand(ByVal CustCode As String, ByVal routeNo As String)
        Try
            Dim strQry As String = String.Empty
            If clsCommon.myLen(CustCode) > 0 Then
                Dim ExistsCust As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(PK_ID) from TSPL_DEMAND_SHEET where convert(date,DEMAND_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and ShiftType='" + clsCommon.myCstr(cmbShift.SelectedValue) + "' and Cust_Code='" + CustCode + "'and Created_By='" + objCommonVar.CurrentUserCode + "'"))
                If ExistsCust = 0 Then
                    Dim document As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_NO from TSPL_DEMAND_BOOKING_MASTER where convert(date,Document_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and ShiftType='" + clsCommon.myCstr(cmbShift.SelectedValue) + "' and Route_No='" + clsCommon.myCstr(routeNo) + "' and IsIndividualCustomer=0"))
                    If clsCommon.myLen(document) > 0 Then
                        strQry = " select Item_Code,Unit_Code,Qty from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + document + "' and Cust_Code='" + CustCode + "' "
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
                        If dt IsNot Nothing And dt.Rows.Count > 0 Then
                            Dim obj As New clsDemandSheet()
                            For Each dr As DataRow In dt.Rows
                                obj.DEMAND_Date = clsCommon.GetPrintDate(txtDate.Value)
                                obj.Cust_Code = clsCommon.myCstr(CustCode)
                                obj.Route_No = clsCommon.myCstr(routeNo)
                                obj.Set_Zero = clsCommon.myCdbl(1)
                                obj.ShiftType = clsCommon.myCstr(cmbShift.SelectedValue)
                                obj.Item_Code = clsCommon.myCstr(dr.Item("Item_Code"))
                                obj.Unit_Code = clsCommon.myCstr(dr.Item("Unit_Code"))
                                obj.Qty = clsCommon.myCdbl(dr.Item("qty"))
                                Dim k As Integer = 1
                                For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                                    k = k + 1
                                    If obj1 IsNot Nothing Then
                                        If clsCommon.CompairString(obj.Item_Code, obj1.itemCode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Unit_Code, obj1.Unit_code) = CompairStringResult.Equal Then
                                            If obj.Qty > 0 Then
                                                gv1.CurrentRow.Cells(dblcolumns).Value = obj.Qty
                                            End If
                                        End If
                                    End If
                                Next
                                Try
                                    Dim status As Boolean = obj.SaveData(obj)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            Next
                        Else
                            Dim k As Integer = 1
                            For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                                Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                                k = k + 1
                                If obj1 IsNot Nothing Then
                                    gv1.CurrentRow.Cells(dblcolumns).Value = 0
                                End If
                            Next

                        End If
                    Else
                        Dim k As Integer = 1
                        For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                            Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                            k = k + 1
                            If obj1 IsNot Nothing Then
                                gv1.CurrentRow.Cells(dblcolumns).Value = 0
                            End If
                        Next
                    End If
                Else

                    strQry = "select Item_Code,Unit_Code,Qty from TSPL_DEMAND_SHEET where convert(date,DEMAND_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and ShiftType='" + clsCommon.myCstr(cmbShift.SelectedValue) + "' and Cust_Code='" + CustCode + "'and Created_By='" + objCommonVar.CurrentUserCode + "'"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
                    If dt IsNot Nothing And dt.Rows.Count > 0 Then
                        For Each dr As DataRow In dt.Rows
                            Dim k As Integer = 1
                            For dblcolumns As Integer = 5 To gv1.Columns.Count - 2
                                Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                                k = k + 1
                                If obj1 IsNot Nothing Then
                                    If clsCommon.CompairString(clsCommon.myCstr(dr.Item("Item_Code")), obj1.itemCode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(dr.Item("Unit_Code")), obj1.Unit_code) = CompairStringResult.Equal Then
                                        gv1.CurrentRow.Cells(dblcolumns).Value = clsCommon.myCdbl(dr.Item("qty"))
                                        'gv1.CurrentRow.Cells(colCrate).Value = clsCommon.myCdbl(gv1.CurrentRow.Cells(colCrate).Value) + clsCommon.myCdbl(dr.Item("qty"))
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub gv1_FilterChanging(sender As Object, e As GridViewCollectionChangingEventArgs) Handles gv1.FilterChanging
        isInsideLoadData = True
    End Sub
    Private Sub gv1_CellBeginEdit(sender As Object, e As GridViewCellCancelEventArgs) Handles gv1.CellBeginEdit
        isInsideLoadData = False
        If e.Column.Index = gv1.Columns.Count - 2 Then
            isLastcellvaluechanged = True
        End If
    End Sub
    Public Function FindCustInGrid(ByVal CustCode As String) As Boolean
        If gv1.Rows.Count > 2 Then
            For dblrows As Integer = 0 To gv1.Rows.Count - 3
                If clsCommon.myLen(CustCode) > 0 Then
                    If clsCommon.CompairString(gv1.Rows(dblrows).Cells(colCustCode).Value, CustCode) = CompairStringResult.Equal Then
                        gv1.Rows.Remove(gv1.Rows(dblrows))
                        GenerateLineNo()
                        Return True
                    End If
                End If
            Next
        End If
        Return False
    End Function
    Private Sub gv1_UserDeletingRow(sender As Object, e As GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        Try
            If (myMessages.deleteConfirm()) Then
                Dim strQry As String = " delete TSPL_DEMAND_SHEET where DEMAND_Date='" + clsCommon.GetPrintDate(txtDate.Value) + "' and ShiftType='" + clsCommon.myCstr(cmbShift.SelectedValue) + "' and Cust_Code='" + gv1.CurrentRow.Cells(colCustCode).Value + "' and Created_By='" + objCommonVar.CurrentUserCode + "' "
                clsDBFuncationality.ExecuteNonQuery(strQry)
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub GenerateLineNo()
        Dim count As Integer = 1
        For Each grow As GridViewRowInfo In gv1.Rows
            grow.Cells(colLineNo).Value = count
            count += 1
        Next
    End Sub
    'Private Sub gv1_CurrentRowChanged(sender As Object, e As CurrentRowChangedEventArgs) Handles gv1.CurrentRowChanged
    '    If gv1.CurrentRow IsNot Nothing AndAlso Not e.CurrentRow.Index < 0 Then
    '        GetBoothDetail()
    '    End If
    'End Sub
    Public Sub GetBoothDetail()
        Try
            If clsCommon.myLen(gv1.CurrentRow.Cells(colCustCode)) > 0 Then
                txtBoothName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "'"))
                txtDistributor.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code =(select Distributor_Code from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "')"))
                txtAddress.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Add1+Add2+Add3 from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) + "'"))
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        AddNew()
    End Sub
End Class