Imports common
Imports System
Imports System.Text
Public Class frmRouteWiseSaleTarget

#Region "Variables"
    Private isNewEntry As Boolean = False
    Dim isInsideLoadData As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Const colSNo As String = "COLSNO"
    Const colRouteCode As String = "COLROUTECODE"
    Const colRouteName As String = "COLROUTENAME"
    Const colGroupCode As String = "COLGROUPCODE"
    Const colGroupName As String = "COLGROUPNAME"
    Const colTargetQty As String = "COLTARGETQTY"
#End Region

    Private Sub frmRouteWiseSaleTarget_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Reset()
            'createTable()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub createTable()
        Dim coll As Dictionary(Of String, String)
        coll = New Dictionary(Of String, String)()
        coll.Add("Document_Code", "varchar(30) NOT NULL Primary Key")
        coll.Add("Document_Date", "DateTime not NULL")
        coll.Add("Months", "Varchar(10) not NULL")
        coll.Add("UOM", "Varchar(10) NULL")
        coll.Add("Target_On", "integer NULL")
        coll.Add("Item_Sub_Category", "varchar(20) Not NULL references tspl_chapter_head(chapter_head_Code)")
        coll.Add("Remarks", "varchar(200) NULL")
        coll.Add("Status", "integer null")
        coll.Add("Created_By", "varchar(12)  NULL")
        coll.Add("Created_Date", "Datetime NOT NULL")
        coll.Add("Modified_By", "varchar(12) NOT NULL")
        coll.Add("Modified_Date", "Datetime NOT NULL")
        coll.Add("Posted_By", "varchar(12) NULL")
        coll.Add("Posted_Date", "Datetime NULL")
        coll.Add("Inactive", "integer null")
        coll.Add("Inactive_By", "varchar(12) NULL")
        coll.Add("Inactive_Date", "Datetime NULL")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_ROUTE_WISE_SALE_TARGET", coll, Nothing, True, True, "", "Document_Code", "Document_Date", True)

        coll = New Dictionary(Of String, String)()
        coll.Add("PK_ID", "integer NOT NULL identity NOT FOR REPLICATION primary key")
        coll.Add("Document_Code", "varchar(30) NOT NULL references TSPL_ROUTE_WISE_SALE_TARGET(Document_Code)")
        coll.Add("Route_Code", "varchar(12) NULL references TSPL_ROUTE_MASTER(Route_No)")
        coll.Add("Cust_Group_Code", "varchar(12) NULL references TSPL_CUSTOMER_GROUP_MASTER(Cust_Group_Code)")
        coll.Add("Target_Qty", "Decimal(18,2) NULL")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_ROUTE_WISE_SALE_TARGET_DETAIL", coll, Nothing, True, True, "TSPL_ROUTE_WISE_SALE_TARGET", "DOCUMENT_CODE", "", True)
    End Sub

    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        Try
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        isInsideLoadData = False
        txtDocumentNo.Value = Nothing
        txtDocumentDate.Value = clsCommon.GETSERVERDATE()
        txtMonth.Value = txtDocumentDate.Value
        cboUOM.SelectedItem = Nothing
        txtItemSubCateg.Value = Nothing
        lblItemSubCategName.Text = Nothing
        chkInactive.Checked = False
        lblStatus.Status = ERPTransactionStatus.Pending
        txtRemarks.Text = Nothing
        btnSave.Text = "Save"
        btnReverseUnpost.Visible = False
        rbtnRoute.Checked = True
        rbtnGroup.Checked = False
        LoadBlankGrid(False)
        EnableDisableFields(True)
    End Sub

    Sub EnableDisableFields(ByVal isEnable As Boolean)
        btnSave.Enabled = isEnable
        btnPost.Enabled = isEnable
        btnDelete.Enabled = isEnable
        chkInactive.Visible = Not isEnable
        chkInactive.Enabled = Not isEnable
    End Sub

    Private Sub LoadBlankGrid(ByVal isReadOnly As Boolean)
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()

        Dim repoSNO As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSNO.FormatString = ""
        repoSNO.HeaderText = "S.No"
        repoSNO.Name = colSNo
        repoSNO.Width = 50
        repoSNO.ReadOnly = True
        repoSNO.IsVisible = True
        Gv1.MasterTemplate.Columns.Add(repoSNO)

        If rbtnRoute.Checked Then
            Dim repoRouteCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            repoRouteCode.FormatString = ""
            repoRouteCode.HeaderText = "Route Code"
            repoRouteCode.HeaderImage = My.Resources.search4
            repoRouteCode.TextImageRelation = TextImageRelation.TextBeforeImage
            repoRouteCode.Name = colRouteCode
            repoRouteCode.Width = 150
            repoRouteCode.ReadOnly = isReadOnly
            repoRouteCode.IsVisible = True
            Gv1.MasterTemplate.Columns.Add(repoRouteCode)

            Dim repoRouteName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            repoRouteName.FormatString = ""
            repoRouteName.HeaderText = "Route Name"
            repoRouteName.Name = colRouteName
            repoRouteName.Width = 300
            repoRouteName.ReadOnly = True
            repoRouteName.IsVisible = True
            Gv1.MasterTemplate.Columns.Add(repoRouteName)
        Else
            Dim repoGroupCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            repoGroupCode.FormatString = ""
            repoGroupCode.HeaderText = "Group Code"
            repoGroupCode.HeaderImage = My.Resources.search4
            repoGroupCode.TextImageRelation = TextImageRelation.TextBeforeImage
            repoGroupCode.Name = colGroupCode
            repoGroupCode.Width = 150
            repoGroupCode.ReadOnly = isReadOnly
            repoGroupCode.IsVisible = True
            Gv1.MasterTemplate.Columns.Add(repoGroupCode)

            Dim repoGroupName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            repoGroupName.FormatString = ""
            repoGroupName.HeaderText = "Group Name"
            repoGroupName.Name = colGroupName
            repoGroupName.Width = 300
            repoGroupName.ReadOnly = True
            repoGroupName.IsVisible = True
            Gv1.MasterTemplate.Columns.Add(repoGroupName)
        End If

        Dim repoTargetQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTargetQty.FormatString = "{0:n2}"
        repoTargetQty.HeaderText = "Target Qty"
        repoTargetQty.Name = colTargetQty
        repoTargetQty.Width = 200
        repoTargetQty.ReadOnly = isReadOnly
        repoTargetQty.ShowUpDownButtons = True
        Gv1.MasterTemplate.Columns.Add(repoTargetQty)

        Gv1.Rows.AddNew()
        Gv1.AllowDeleteRow = True
        Gv1.AllowAddNewRow = False
        Gv1.ShowGroupPanel = False
        Gv1.AllowColumnReorder = False
        Gv1.AllowRowReorder = False
        Gv1.EnableSorting = False
        Gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.AutoSizeRows = False
    End Sub

    Private Sub txtItemSubCateg__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtItemSubCateg._MYValidating
        Try
            Dim Qry As String = "select chapter_head_Code as Code,Description as [Description] from tspl_chapter_head"
            txtItemSubCateg.Value = clsCommon.ShowSelectForm("@ItemSubCateg", Qry, "Code", Nothing, txtItemSubCateg.Value, "Code", isButtonClicked)
            ItemSubCategory()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub ItemSubCategory()
        Try
            lblItemSubCategName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from tspl_chapter_head Where chapter_head_Code='" & txtItemSubCateg.Value & "'"))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub Gv1_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles Gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column Is Gv1.Columns(colRouteCode) Then
                        OpenRouteList(False)
                    ElseIf e.Column Is Gv1.Columns(colGroupCode) Then
                        OpenGroupList(False)
                    ElseIf e.Column Is Gv1.Columns(colTargetQty) Then
                        Gv1.Rows.AddNew()
                    End If
                End If
                isCellValueChangedOpen = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub OpenRouteList(ByVal isButtonClick As Boolean)
        Try
            Dim Qry As String = "Select Route_No As [Code],Route_Desc As [Description] from TSPL_ROUTE_MASTER "
            '            Dim whr As String = "  Route_No Not In (Select TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code from TSPL_ROUTE_WISE_SALE_TARGET_DETAIL
            'Inner Join TSPL_ROUTE_WISE_SALE_TARGET On TSPL_ROUTE_WISE_SALE_TARGET.Document_Code=TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Document_Code Where TRY_CONVERT(DATE, '01/' + TSPL_ROUTE_WISE_SALE_TARGET.Months, 103) <= TRY_CONVERT(DATE, '" & txtMonth.Value & "', 103) And IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Inactive,0)=0 And TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category='" & txtItemSubCateg.Value & "' ) "
            Gv1.CurrentRow.Cells(colRouteCode).Value = clsCommon.ShowSelectForm("@RouteCode", Qry, "Code", Nothing, Nothing, "Code", True)
            Gv1.CurrentRow.Cells(colRouteName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Route_Desc from TSPL_ROUTE_MASTER Where Route_No='" & clsCommon.myCstr(Gv1.CurrentRow.Cells(colRouteCode).Value) & "'"))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub OpenGroupList(ByVal isButtonClick As Boolean)
        Try
            Dim Qry As String = " SELECT Cust_Group_Code as Code,Cust_Group_Desc As [Description] FROM TSPL_CUSTOMER_GROUP_MASTER "
            '            Dim whr As String = "  Route_No Not In (Select TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code from TSPL_ROUTE_WISE_SALE_TARGET_DETAIL
            'Inner Join TSPL_ROUTE_WISE_SALE_TARGET On TSPL_ROUTE_WISE_SALE_TARGET.Document_Code=TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Document_Code Where TRY_CONVERT(DATE, '01/' + TSPL_ROUTE_WISE_SALE_TARGET.Months, 103) <= TRY_CONVERT(DATE, '" & txtMonth.Value & "', 103) And IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Inactive,0)=0 And TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category='" & txtItemSubCateg.Value & "' ) "
            Gv1.CurrentRow.Cells(colGroupCode).Value = clsCommon.ShowSelectForm("@GroupCode", Qry, "Code", Nothing, Nothing, "Code", True)
            Gv1.CurrentRow.Cells(colGroupName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Cust_Group_Desc from TSPL_CUSTOMER_GROUP_MASTER Where Cust_Group_Code='" & clsCommon.myCstr(Gv1.CurrentRow.Cells(colGroupCode).Value) & "'"))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub RefreshSNO()
        If Gv1 IsNot Nothing AndAlso Gv1.Rows.Count > 0 Then
            For i As Integer = 1 To Gv1.Rows.Count
                Gv1.Rows(i - 1).Cells(colSNo).Value = i
            Next
        End If
    End Sub

    Function AllowToSave() As Boolean
        Try
            Xtra.TransactionValidity(txtDocumentDate.Value)
            If clsCommon.myLen(cboUOM.SelectedItem) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select UOM !", Me.Text)
                Return False
            End If
            If clsCommon.myLen(txtItemSubCateg.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Item Sub Category can't ne blank !", Me.Text)
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveData(False)
    End Sub

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        SaveData(True)
    End Sub

    Sub SaveData(ByVal isPost As Boolean)
        Try
            If isPost AndAlso clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                Throw New Exception("Document Code not found to post !")
            End If
            If AllowToSave() Then
                Dim obj As New clsRouteWiseSaleTarget()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtDocumentDate.Value
                obj.Month = txtMonth.Value
                obj.UOM = clsCommon.myCstr(cboUOM.SelectedItem)
                If rbtnRoute.Checked Then
                    obj.Target_On = 0
                Else
                    obj.Target_On = 1
                End If
                obj.Item_Sub_Category = txtItemSubCateg.Value
                obj.Remarks = txtRemarks.Text
                Dim sbRoute As New StringBuilder()
                If Gv1 IsNot Nothing AndAlso Gv1.Rows.Count > 0 Then
                    obj.Arr = New List(Of clsRouteWiseSaleTargetDetail)
                    For Each row In Gv1.Rows
                        Dim chkQry As String = "Select "
                        If rbtnRoute.Checked Then
                            chkQry += " Count(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code) "
                        Else
                            chkQry += " Count(TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Cust_Group_Code) "
                        End If
                        chkQry += " From TSPL_ROUTE_WISE_SALE_TARGET_DETAIL
Inner Join TSPL_ROUTE_WISE_SALE_TARGET On TSPL_ROUTE_WISE_SALE_TARGET.Document_Code=TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Document_Code Where TRY_CONVERT(DATE, '01/' + TSPL_ROUTE_WISE_SALE_TARGET.Months, 103) <= TRY_CONVERT(DATE, '" & txtMonth.Value & "', 103) And IsNull(TSPL_ROUTE_WISE_SALE_TARGET.Inactive,0)=0 And TSPL_ROUTE_WISE_SALE_TARGET.Item_Sub_Category='" & txtItemSubCateg.Value & "' "
                        If rbtnRoute.Checked Then
                            chkQry += " And TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Route_Code ='" & clsCommon.myCstr(row.Cells(colRouteCode).Value) & "' "
                        Else
                            chkQry += " And TSPL_ROUTE_WISE_SALE_TARGET_DETAIL.Cust_Group_Code='" & clsCommon.myCstr(row.Cells(colGroupCode).Value) & "' "
                        End If
                        If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                            chkQry += " And TSPL_ROUTE_WISE_SALE_TARGET.Document_Code <>'" & txtDocumentNo.Value & "'"
                        End If
                        If clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(chkQry)) > 0 Then
                            If clsCommon.myLen(sbRoute) <> 0 Then
                                sbRoute.Append(",")
                            End If
                            sbRoute.Append(clsCommon.myCstr(row.Cells(colRouteCode).Value))
                        Else
                            Dim objTr As New clsRouteWiseSaleTargetDetail()
                            If rbtnRoute.Checked Then
                                objTr.Route_Code = clsCommon.myCstr(row.Cells(colRouteCode).Value)
                            Else
                                objTr.Group_Code = clsCommon.myCstr(row.Cells(colGroupCode).Value)
                            End If
                            objTr.Target_Qty = clsCommon.myCDecimal(row.Cells(colTargetQty).Value)
                            If clsCommon.myLen(objTr.Route_Code) > 0 OrElse clsCommon.myLen(objTr.Group_Code) > 0 Then
                                obj.Arr.Add(objTr)
                            End If
                            objTr = Nothing
                        End If
                    Next
                    If clsCommon.myLen(sbRoute) <> 0 Then
                        clsCommon.MyMessageBoxShow(Me, clsCommon.myCstr(IIf(rbtnRoute.Checked, "Route ", "Group ")) & clsCommon.myCstr(sbRoute) & " are already exist.", Me.Text)
                        sbRoute = Nothing
#Disable Warning
                        Exit Sub
#Enable Warning
                    End If

                End If
                sbRoute = Nothing
                If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                    isNewEntry = True
                End If
                If obj.SaveData(obj, isNewEntry) Then
                    If isPost Then
                        If clsCommon.MyMessageBoxShow(Me, "Are you sure to post data ?", Me.Text, MessageBoxButtons.YesNo) = DialogResult.Yes AndAlso obj.PostData(obj.Document_Code) Then
                            clsCommon.MyMessageBoxShow(Me, "Data Posted Successfully.", Me.Text)
                        End If
                    Else
                        clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully.", Me.Text)
                    End If
                    LoadData(obj.Document_Code, NavigatorType.Current, False)
                End If
                obj = Nothing
            End If
            isNewEntry = False
        Catch ex As Exception
            isNewEntry = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Public Sub LoadData(ByVal strDoc As String, ByVal NavTyep As NavigatorType, ByVal isCC As Boolean)
        Dim obj As clsRouteWiseSaleTarget = Nothing
        Try
            Reset()
            obj = New clsRouteWiseSaleTarget()
            obj = obj.GetData(strDoc, NavTyep)
            If obj IsNot Nothing Then
                isInsideLoadData = True
                If Not isCC Then
                    txtDocumentNo.Value = obj.Document_Code
                    txtDocumentDate.Value = obj.Document_Date
                    txtMonth.Value = obj.Month
                    btnSave.Text = "Update"
                Else
                    txtDocumentNo.Value = Nothing
                    txtDocumentDate.Value = clsCommon.GETSERVERDATE()
                    txtMonth.Value = txtDocumentDate.Value
                    btnSave.Text = "Save"
                End If
                cboUOM.SelectedValue = obj.UOM
                If obj.Target_On = 1 Then
                    rbtnGroup.Checked = True
                    rbtnRoute.Checked = False
                Else
                    rbtnRoute.Checked = True
                    rbtnGroup.Checked = False
                End If
                txtItemSubCateg.Value = obj.Item_Sub_Category
                ItemSubCategory()
                txtRemarks.Text = obj.Remarks
                If Not isCC AndAlso obj.Status = 1 Then
                    lblStatus.Status = ERPTransactionStatus.Approved
                    EnableDisableFields(False)
                Else
                    lblStatus.Status = ERPTransactionStatus.Pending
                    EnableDisableFields(True)
                End If
                If obj.Inactive = 1 Then
                    chkInactive.Checked = True
                    chkInactive.Enabled = False
                Else
                    chkInactive.Checked = False
                    chkInactive.Enabled = True
                End If
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    If obj.Status = 1 Then
                        LoadBlankGrid(True)
                    Else
                        LoadBlankGrid(False)
                    End If
                    For Each row In obj.Arr
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colSNo).Value = row.LineNo
                        If rbtnRoute.Checked Then
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colRouteCode).Value = row.Route_Code
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colRouteName).Value = row.Route_Name
                        Else
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colGroupCode).Value = row.Group_Code
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colGroupName).Value = row.Group_Name
                        End If
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colTargetQty).Value = row.Target_Qty
                        Gv1.Rows.AddNew()
                    Next
                End If
                isInsideLoadData = False
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            obj = Nothing
            Throw New Exception(ex.Message)
        End Try
        obj = Nothing
    End Sub


    Private Sub Gv1_CellValidated(sender As Object, e As CellValidatedEventArgs) Handles Gv1.CellValidated
        Try
            setGridFocus()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub setGridFocus()
        Try
            If Gv1.CurrentColumn IsNot Nothing AndAlso clsCommon.CompairString(Gv1.CurrentColumn.Name, colTargetQty) = CompairStringResult.Equal Then
                If rbtnRoute.Checked Then
                    Gv1.CurrentColumn = Gv1.Columns(colRouteCode)
                Else
                    Gv1.CurrentColumn = Gv1.Columns(colGroupCode)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub Gv1_CurrentColumnChanged(sender As Object, e As CurrentColumnChangedEventArgs) Handles Gv1.CurrentColumnChanged
        Try
            RefreshSNO()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub Gv1_UserDeletedRow(sender As Object, e As GridViewRowEventArgs) Handles Gv1.UserDeletedRow
        Try
            RefreshSNO()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocumentNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDocumentNo._MYValidating
        Try
            Dim Qry As String = "Select Document_Code As DocumentCode,Document_Date As [Document Date],Months,Item_Sub_Category As [Item Sub Category],Case When IsNull(Status,0)=0 Then 'Pending' Else 'Approved' End As Status from TSPL_ROUTE_WISE_SALE_TARGET"
            txtDocumentNo.Value = clsCommon.ShowSelectForm("@Doc", Qry, "DocumentCode", "", txtDocumentNo.Value, "TSPL_ROUTE_WISE_SALE_TARGET.Document_Date desc", isButtonClicked)
            If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                LoadData(txtDocumentNo.Value, NavigatorType.Current, False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                Throw New Exception("Document Code not found to delete !")
            End If
            If clsCommon.MyMessageBoxShow(Me, "Are you sure to delete ?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                Dim obj As New clsRouteWiseSaleTarget()
                If obj.DeleteData(txtDocumentNo.Value) Then
                    clsCommon.MyMessageBoxShow(Me, "Data deleted successfully.")
                    Reset()
                End If
                obj = Nothing
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocumentNo__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles txtDocumentNo._MYNavigator
        Try
            LoadData(txtDocumentNo.Value, NavType, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnCC_Click(sender As Object, e As EventArgs) Handles btnCC.Click
        Try
            Dim Qry As String = "Select Document_Code As DocumentCode,Document_Date As [Document Date],Months,Item_Sub_Category As [Item Sub Category],Case When IsNull(Status,0)=0 Then 'Pending' Else 'Approved' End As Status from TSPL_ROUTE_WISE_SALE_TARGET"
            Dim strCode As String = clsCommon.ShowSelectForm("@Doc", Qry, "DocumentCode", "", txtDocumentNo.Value, "TSPL_ROUTE_WISE_SALE_TARGET.Document_Date desc", True)
            LoadData(strCode, NavigatorType.Current, True)
            txtDocumentNo.Value = Nothing
            EnableDisableFields(True)
            isInsideLoadData = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub chkInactive_CheckedChanged(sender As Object, e As EventArgs) Handles chkInactive.CheckedChanged
        Try
            If chkInactive.Checked AndAlso Not isInsideLoadData Then
                If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                    Throw New Exception("Document Code not found !")
                End If
                If clsCommon.MyMessageBoxShow(Me, "Are you sure to inactive ?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                    Dim obj As New clsRouteWiseSaleTarget()
                    If obj.DataInactive(txtDocumentNo.Value) Then
                        clsCommon.MyMessageBoxShow(Me, "Data inactive successfully.", Me.Text)
                        LoadData(txtDocumentNo.Value, Nothing, False)
                    End If
                    obj = Nothing
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub frmRouteWiseSaleTarget_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Try
            If e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
                btnAddNew.PerformClick()
            ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso btnSave.Enabled AndAlso MyBase.isModifyFlag Then
                SaveData(False)
            ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso btnSave.Enabled AndAlso MyBase.isDeleteFlag Then
                btnDelete.PerformClick()
            ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnPost.Enabled AndAlso MyBase.isPostFlag Then
                btnPost.PerformClick()
            ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
                Me.Close()
            ElseIf e.Alt AndAlso e.Shift AndAlso e.Control AndAlso e.KeyCode = Keys.F12 Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverseUnpost.Visible = True
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnRoute_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnRoute.CheckedChanged
        CheckedRadioButton()
    End Sub

    Private Sub rbtnGroup_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnGroup.CheckedChanged
        CheckedRadioButton()
    End Sub

    Sub CheckedRadioButton()
        Try
            LoadBlankGrid(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnReverseUnpost_Click(sender As Object, e As EventArgs) Handles btnReverseUnpost.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                If clsCommon.MyMessageBoxShow(Me, "Are you sure to Reverse ?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                    Dim obj As New clsRouteWiseSaleTarget()
                    If obj.ReverseAndUnpost(txtDocumentNo.Value) Then
                        clsCommon.MyMessageBoxShow(Me, "Data Reverse Successfully.")
                    End If
                    obj = Nothing
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Document Code not found to reverse !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class