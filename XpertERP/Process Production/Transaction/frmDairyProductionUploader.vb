Imports common
Imports System.IO

Public Class frmDairyProductionUploader
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim isNewEntry As Boolean = False
    Const ColPKID As String = "ColPKID"
    Const ColSNo As String = "ColSNo"
    Const colBatchDate As String = "colShiftDate"
    Const colShift As String = "colShift"
    Const colItemCode As String = "colItemCode"
    Const colItemName As String = "colItemName"
    Const colQty As String = "colQty"
    Const colUOM As String = "colUOM"
    Const colBatchNo As String = "colBatchNo"
    Const ColQCStatus As String = "ColQCStatus"


    Const ReportID As String = "Produploader"
    Dim isInsideLoadData As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim arrLoc As String = Nothing
    Dim atchqry As String = ""
#End Region

    Private Sub frmDairyProductionUploader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LOCATIONRIGTHS()
        AddNew()
        btnReverse.Visible = False
        'Dim coll As Dictionary(Of String, String)

        'coll = New Dictionary(Of String, String)
        'coll.Add("Document_No", "Varchar(30) not null Primary key")
        'coll.Add("Document_Date", "datetime NOT NULL")
        'coll.Add("Location_FG", "Varchar(12) not null references TSPL_LOCATION_MASTER(Location_Code)")
        'coll.Add("Location_RM", "Varchar(12) not null references TSPL_LOCATION_MASTER(Location_Code)")
        'coll.Add("Location_PK", "Varchar(12) not null references TSPL_LOCATION_MASTER(Location_Code)")
        'coll.Add("Batch_No", "Varchar(200) not null")
        'coll.Add("Batch_Date", "Date not null")
        'coll.Add("Description", "Varchar(200) null")
        'coll.Add("Status", "Integer NOT NULL DEFAULT 0")
        'coll.Add("Created_By", "varchar(12) NOT NULL")
        'coll.Add("Created_Date", "Datetime NOT NULL")
        'coll.Add("Modified_By", "varchar(12) NOT NULL")
        'coll.Add("Modified_Date", "Datetime NOT NULL")
        'coll.Add("Posted_Date", "datetime null")
        'coll.Add("Posted_By", "varchar(12)  NULL")
        'clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_PRODUCTION_UPLOADER_HEAD", coll, Nothing, True, True, "", "Document_No", "Document_Date", True)

        'coll = New Dictionary(Of String, String)
        'coll.Add("PK_ID", "integer NOT NULL identity NOT FOR REPLICATION primary key")
        'coll.Add("Document_No", "Varchar(30) not null references TSPL_PRODUCTION_UPLOADER_HEAD(Document_No)")
        'coll.Add("Batch_No", "Varchar(200) not null")
        'coll.Add("Batch_Date", "Date NOT NULL")
        'coll.Add("Item_Code", "Varchar(50) not null references TSPL_ITEM_MASTER(Item_Code)")
        'coll.Add("Qty", "Decimal(18,2) null")
        'coll.Add("UOM", "Varchar(20) null")
        'coll.Add("Shift_Code", "Varchar(30) not null references tspl_shift_master(SHIFT_CODE)")
        'coll.Add("QC_Status", "Integer NOT NULL DEFAULT 0")
        'coll.Add("BOM_Code", "Varchar(30) null references TSPL_PP_BOM_HEAD(BOM_CODE)")
        'clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_PRODUCTION_UPLOADER_DETAIL", coll, Nothing, True, True, "TSPL_PRODUCTION_UPLOADER_HEAD", "Document_No", "")

    End Sub

    Private Sub LOCATIONRIGTHS()
        Dim obj As New clsMCCCodes()
        Try
            obj = clsMCCCodes.GetData()
            If obj.arrLocCodes IsNot Nothing AndAlso clsCommon.myLen(obj.arrLocCodes) > 0 Then
                arrLoc = obj.arrLocCodes
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub

    Sub AddNew()
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        txtDocNo.Value = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtDesc.Text = ""
        txtLocationFG.Value = ""
        lblLocationFG.Text = ""
        txtLocationRM.Value = ""
        lblLocationRM.Text = ""
        txtLocationPK.Value = ""
        lblLocationPK.Text = ""
        txtBatchNo.Text = ""
        txtBatchDate.Value = txtDate.Value
        isNewEntry = True
        UsLock1.Status = ERPTransactionStatus.Pending
        LoadBlankGrid()
        gv1.Rows.AddNew()
    End Sub

    Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()

        Dim repoTextBox As New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "PK ID"
        repoTextBox.Name = ColPKID
        repoTextBox.TextImageRelation = TextImageRelation.TextBeforeImage
        repoTextBox.Width = 200
        repoTextBox.ReadOnly = True
        repoTextBox.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        Dim repoNumBox As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoNumBox = New GridViewDecimalColumn()
        repoNumBox.FormatString = ""
        repoNumBox.HeaderText = "SNo"
        repoNumBox.Name = ColSNo
        repoNumBox.Width = 40
        repoNumBox.ReadOnly = True
        repoNumBox.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoNumBox)

        repoTextBox = New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "Item Code"
        repoTextBox.Name = colItemCode
        repoTextBox.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoTextBox.TextImageRelation = TextImageRelation.TextBeforeImage
        repoTextBox.Width = 100
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        repoTextBox = New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "Item Name"
        repoTextBox.Name = colItemName
        repoTextBox.Width = 150
        repoTextBox.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        repoNumBox = New GridViewDecimalColumn()
        repoNumBox.FormatString = ""
        repoNumBox.HeaderText = "Qty"
        repoNumBox.Name = colQty
        repoNumBox.Width = 80
        repoNumBox.Minimum = 0
        repoNumBox.ShowUpDownButtons = False
        repoNumBox.Step = 0
        repoNumBox.DecimalPlaces = 0
        repoNumBox.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoNumBox)

        repoTextBox = New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "UOM"
        repoTextBox.Name = colUOM
        repoTextBox.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoTextBox.TextImageRelation = TextImageRelation.TextBeforeImage
        repoTextBox.Width = 100
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        repoTextBox = New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "Batch no"
        repoTextBox.Name = colBatchNo
        repoTextBox.TextImageRelation = TextImageRelation.TextBeforeImage
        repoTextBox.Width = 100
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        Dim repoDateBox As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoDateBox.Format = DateTimePickerFormat.Custom
        repoDateBox.CustomFormat = "dd/MM/yyyy"
        repoDateBox.FormatString = "{0:dd/MM/yyyy}"
        repoDateBox.HeaderText = "Batch Date"
        repoDateBox.Name = colBatchDate
        repoDateBox.ReadOnly = False
        repoDateBox.IsVisible = True
        repoDateBox.Width = 100
        gv1.MasterTemplate.Columns.Add(repoDateBox)

        repoTextBox = New GridViewTextBoxColumn()
        repoTextBox.FormatString = ""
        repoTextBox.HeaderText = "Shift"
        repoTextBox.Name = colShift
        repoTextBox.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoTextBox.TextImageRelation = TextImageRelation.TextBeforeImage
        repoTextBox.Width = 100
        gv1.MasterTemplate.Columns.Add(repoTextBox)

        Dim repoCheck As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoCheck.FormatString = ""
        repoCheck.HeaderText = "QC Status"
        repoCheck.Name = ColQCStatus
        repoCheck.Width = 60
        repoCheck.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCheck)

        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = True
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40

        gv1.AllowDeleteRow = True

        ReStoreGridLayout()
    End Sub

    Private Sub gv1_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column Is gv1.Columns(colItemCode) Then
                        OpenItem(False)
                    ElseIf e.Column Is gv1.Columns(colUOM) Then
                        OpenUOM(False)
                    ElseIf e.Column Is gv1.Columns(colShift) Then
                        OpenShiftCode(False)
                        'ElseIf e.Column Is gv1.Columns(colBatchNo) Then
                        '  OpenBatchDate()
                    End If
                    isCellValueChangedOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub OpenBatchDate()

        Dim itemCode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value)
        Dim BACTH As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colBatchNo).Value)
        Dim qry As String = "SELECT Expiry_Date " &
                        "FROM TSPL_BATCH_ITEM " &
                        "WHERE Item_Code = '" & itemCode & "' AND Batch_No= '" + BACTH + "'"

        Dim expiryDate As Object = clsDBFuncationality.getSingleValue(qry)

        If expiryDate IsNot Nothing AndAlso Not IsDBNull(expiryDate) Then
            gv1.CurrentRow.Cells(colBatchDate).Value =
            Convert.ToDateTime(expiryDate).ToString("dd/MM/yyyy")
        Else
            gv1.CurrentRow.Cells(colBatchDate).Value = ""
        End If

    End Sub
    Sub OpenItem(ByVal isButtonClicked As Boolean)
        gv1.CurrentRow.Cells(colItemCode).Value = clsItemMaster.getFinder(" tspl_item_master.item_type in ('F') and tspl_item_master.Active='1' ", clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value), isButtonClicked)
        gv1.CurrentRow.Cells(colItemName).Value = clsItemMaster.GetItemName(clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value), Nothing)
    End Sub

    Sub OpenUOM(ByVal isButtonClick As Boolean)
        If clsCommon.myLen(gv1.CurrentRow.Cells(colItemCode).Value) <= 0 Then
            clsCommon.MyMessageBoxShow("Please first select item.")
            Exit Sub
        End If
        Dim qry As String = "select UOM_Code as Code,UOM_Description as Description,Conversion_Factor as [Conversion Factor],Stocking_Unit as [Stocking Unit],Weight from TSPL_ITEM_UOM_DETAIL "
        gv1.CurrentRow.Cells(colUOM).Value = clsCommon.myCstr(clsCommon.ShowSelectForm("PPUOMFND", qry, "Code", " item_code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value) + "'", clsCommon.myCstr(gv1.CurrentRow.Cells(colUOM).Value), "Code", isButtonClick))
    End Sub


    Sub OpenShiftCode(ByVal isButtonClicked As Boolean)
        Dim qry As String = "select shift_code as Code,shift_name as Description,from_time as [From Time],to_time as [To Time],interval_time as [Interval Time],fsthalf_adjust_min as [First Half Adjustment],sechalf_adjust_min as [Second Half Adjustment] from tspl_shift_master"
        gv1.CurrentRow.Cells(colShift).Value = clsCommon.ShowSelectForm("PUSFTFND", qry, "Code", "", clsCommon.myCstr(gv1.CurrentRow.Cells(colShift).Value), "Code", isButtonClicked)
        'gv1.CurrentRow.Cells(colShiftName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select shift_name from tspl_shift_master where shift_code='" + shiftcode + "'"))
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(ReportID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(ReportID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                        gv1.Columns(ii).IsVisible = False
                        gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveData()
    End Sub

    Private Function AllowToSave() As Boolean
        Xtra.TransactionValidity(txtDate.Value)
        'For ii As Integer = 0 To gv1.RowCount - 1
        '    If clsCommon.myLen(gv1.Rows(ii).Cells(colLocCode).Value) > 0 Then

        '        If objCommonVar.AddValidationofMilkTypeinsample Then
        '            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value), "M") = CompairStringResult.Equal Then
        '                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) < objCommonVar.FatMinMix OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) > objCommonVar.FatMaxMix Then
        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] FAT Should be in Range [" + clsCommon.myCstr(objCommonVar.FatMinMix) + " - " + clsCommon.myCstr(objCommonVar.FatMaxMix) + "]")
        '                ElseIf clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) < objCommonVar.SNFMinMix OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) > objCommonVar.SNFMaxMix Then

        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] SNF Should be in Range [" + clsCommon.myCstr(objCommonVar.SNFMinMix) + " - " + clsCommon.myCstr(objCommonVar.SNFMaxMix) + "]")
        '                End If
        '            ElseIf clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value), "C") = CompairStringResult.Equal Then
        '                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) < objCommonVar.FatMinCow OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) > objCommonVar.FatMaxCow Then
        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] FAT Should be in Range [" + clsCommon.myCstr(objCommonVar.FatMinCow) + " - " + clsCommon.myCstr(objCommonVar.FatMaxCow) + "]")
        '                ElseIf clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) < objCommonVar.SNFMinCow OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) > objCommonVar.SNFMaxCow Then
        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] SNF Should be in Range [" + clsCommon.myCstr(objCommonVar.SNFMinCow) + " - " + clsCommon.myCstr(objCommonVar.SNFMaxCow) + "]")
        '                End If
        '            ElseIf clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value), "B") = CompairStringResult.Equal Then
        '                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) < objCommonVar.FatMinBuff OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colFATPer).Value) > objCommonVar.FatMaxBuff Then
        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] FAT Should be in Range [" + clsCommon.myCstr(objCommonVar.FatMinBuff) + " - " + clsCommon.myCstr(objCommonVar.FatMaxBuff) + "]")
        '                ElseIf clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) < objCommonVar.SNFMinBuff OrElse clsCommon.myCdbl(gv1.Rows(ii).Cells(colSNFPer).Value) > objCommonVar.SNFMaxBuff Then
        '                    Throw New Exception("Milk Type [" + clsCommon.myCstr(gv1.Rows(ii).Cells(colDockCollectionMilkType).Value) + "] SNF Should be in Range [" + clsCommon.myCstr(objCommonVar.SNFMinBuff) + " - " + clsCommon.myCstr(objCommonVar.SNFMaxBuff) + "]")
        '                End If
        '            Else
        '                Throw New Exception("Milk Type should be M/B/C")
        '            End If
        '        End If
        '        If settLastMilkReceiptQtyTollerance > 0 Then
        '            Dim qty As Decimal = clsCommon.myCdbl(gv1.Rows(ii).Cells(colMilkWeight).Value)
        '            If qty > 0 Then
        '                Dim dtLastQty As DataTable = clsDBFuncationality.GetDataTable("select QTY,cast( case when (QTY-(QTY*" + clsCommon.myCstr(settLastMilkReceiptQtyTollerance) + "/100)) < 0 then 0 else (QTY-(QTY*" + clsCommon.myCstr(settLastMilkReceiptQtyTollerance) + "/100)) end as decimal(18,2)) as MinQty,cast( (QTY+(QTY*" + clsCommon.myCstr(settLastMilkReceiptQtyTollerance) + "/100)) as decimal(18,2)) as MaxQty from TSPL_MILK_SRN_DETAIL where DOC_CODE in (select top 1  DOC_CODE from tspl_milk_srn_head where DOC_DATE<'" + clsCommon.GetPrintDate(clsCommon.myCDate(gv1.Rows(ii).Cells(colShiftDate).Value), "dd/MMM/yyyy") + "' and VLC_CODE='" + clsCommon.myCstr(gv1.Rows(ii).Cells(colVLCCode).Value) + "' order by doc_date desc,SHIFT)")
        '                If dtLastQty IsNot Nothing AndAlso dtLastQty.Rows.Count > 0 Then
        '                    If qty < clsCommon.myCdbl(dtLastQty.Rows(0)("MinQty")) OrElse qty > clsCommon.myCdbl(dtLastQty.Rows(0)("MaxQty")) Then
        '                        If clsCommon.MyMessageBoxShow("Row No [" + clsCommon.myCstr(ii + 1) + "] Qty [" + clsCommon.myCstr(qty) + "] Tollerance [" + clsCommon.myCstr(settLastMilkReceiptQtyTollerance) + "] and Valid Qty Range [" + clsCommon.myCstr(dtLastQty.Rows(0)("MinQty")) + "-" + clsCommon.myCstr(dtLastQty.Rows(0)("MaxQty")) + "]" + Environment.NewLine + "Do you want to continue...", Me.Text, MessageBoxButtons.YesNo, WinControls.RadMessageIcon.Question) = System.Windows.Forms.DialogResult.No Then
        '                            Return False
        '                        End If
        '                    End If
        '                End If
        '            End If
        '        End If

        '    End If
        'Next
        Return True
    End Function

    Private Sub RadMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem1.Click
        If clsCommon.myLen(ReportID) > 0 Then
            gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = ReportID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv1.SaveLayout(obj.GridLayout)
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            obj.GridColumns = gv1.ColumnCount
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", "Information")
            End If
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub RadMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem2.Click
        clsGridLayout.DeleteData(ReportID, objCommonVar.CurrentUserCode)
    End Sub

    Private Sub FrmSerializeItemIn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.Escape Then
            CancelPressed()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIRC
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverse.Visible = True
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
            End If
        End If
    End Sub

    Sub CancelPressed()
        Me.Close()
    End Sub

    Private Sub gv1_UserDeletedRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        RefeshSNO()
    End Sub

    Private Sub gv1_UserDeletingRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        'If gv1.RowCount <= clsCommon.myCdbl(lblQty.Text) Then
        '    e.Cancel = True
        '    Exit Sub
        'End If

        'If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
        '    e.Cancel = True
        'End If
    End Sub

    Sub RefeshSNO()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(ColSNo).Value = ii
        Next
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CancelPressed()
    End Sub

    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If gv1.RowCount > 0 Then
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(ColSNo).Value = clsCommon.myCdbl(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                gv1.Rows.AddNew()
                gv1.CurrentRow = gv1.Rows(intCurrRow)
            End If
        End If
    End Sub

    Private Sub txtDocNo__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim qst As String = "select count(*) from TSPL_PRODUCTION_UPLOADER_HEAD where Document_No='" + txtDocNo.Value + "'"
            Dim count As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qst))
            If count = 0 Then
                txtDocNo.MyReadOnly = False
            Else
                txtDocNo.MyReadOnly = True
            End If
            LoadData(txtDocNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDocNo._MYValidating
        Dim qry As String = "select TSPL_PRODUCTION_UPLOADER_HEAD.Document_No,convert (varchar,TSPL_PRODUCTION_UPLOADER_HEAD.Document_Date,103) as Document_Date,TSPL_PRODUCTION_UPLOADER_HEAD.Description,case when TSPL_PRODUCTION_UPLOADER_HEAD.Status=1 then 'Posted' else 'Pending' end as Status
,TSPL_PRODUCTION_UPLOADER_HEAD.Location_FG as [FG Location Code]  ,TSPL_LOCATION_MASTER_FG.Location_Desc as [FG Location Name]  
,TSPL_PRODUCTION_UPLOADER_HEAD.Location_RM as [RM Location Code]  ,TSPL_LOCATION_MASTER_RM.Location_Desc as [RM Location Name]  
,TSPL_PRODUCTION_UPLOADER_HEAD.Location_PK as [Packing Location Code]  ,TSPL_LOCATION_MASTER_PK.Location_Desc as [Packing Location Name]  
from TSPL_PRODUCTION_UPLOADER_HEAD 
left outer join TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_FG on TSPL_LOCATION_MASTER_FG.Location_Code=TSPL_PRODUCTION_UPLOADER_HEAD.Location_FG 
left outer join TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_RM on TSPL_LOCATION_MASTER_RM.Location_Code=TSPL_PRODUCTION_UPLOADER_HEAD.Location_RM 
left outer join TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_PK on TSPL_LOCATION_MASTER_PK.Location_Code=TSPL_PRODUCTION_UPLOADER_HEAD.Location_PK "
        LoadData(clsCommon.ShowSelectForm("PUFINDOC", qry, "Document_No", "", txtDocNo.Value, "Document_No", isButtonClicked, "Document_Date"), NavigatorType.Current)
    End Sub

    Public Function SaveData() As Boolean
        Try
            If (AllowToSave()) Then
                Dim obj As New clsDairyProductionUploader()
                obj.Document_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Description = txtDesc.Text
                obj.Location_FG = txtLocationFG.Value
                obj.Location_RM = txtLocationRM.Value
                obj.Location_PK = txtLocationPK.Value
                obj.Batch_No = txtBatchNo.Text
                obj.Batch_Date = txtBatchDate.Value
                obj.Arr = New List(Of clsDairyProductionUploaderDetail)
                For ii As Integer = 0 To gv1.RowCount - 1
                    If clsCommon.myLen(gv1.Rows(ii).Cells(colItemCode).Value) > 0 Then
                        Dim objTr As New clsDairyProductionUploaderDetail()
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colBatchNo).Value) <= 0 Then
                            objTr.Batch_No = txtBatchNo.Text
                        Else
                            objTr.Batch_No = clsCommon.myCstr(gv1.Rows(ii).Cells(colBatchNo).Value)
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colBatchDate).Value) <= 0 Then
                            objTr.Batch_Date = txtBatchDate.Value
                        Else
                            objTr.Batch_Date = clsCommon.myCDate(gv1.Rows(ii).Cells(colBatchDate).Value)
                        End If
                        'objTr.Batch_Date = clsCommon.myCDate(gv1.Rows(ii).Cells(colBatchDate).Value)
                        objTr.Item_Code = clsCommon.myCstr(gv1.Rows(ii).Cells(colItemCode).Value)
                        objTr.Qty = clsCommon.myCdbl(gv1.Rows(ii).Cells(colQty).Value)
                        objTr.UOM = clsCommon.myCstr(gv1.Rows(ii).Cells(colUOM).Value)
                        objTr.Shift_Code = clsCommon.myCstr(gv1.Rows(ii).Cells(colShift).Value)
                        If clsCommon.myLen(objTr.UOM) <= 0 Then
                            Throw New Exception("Please select UOM at line no [" + clsCommon.myCstr(ii + 1) + "]")
                        End If
                        If clsCommon.myLen(objTr.Shift_Code) <= 0 Then
                            Throw New Exception("Please select Shift at line no [" + clsCommon.myCstr(ii + 1) + "]")
                        End If
                        Dim MinShelfLife As Integer = Val(clsDBFuncationality.getSingleValue(
    "SELECT ISNULL(Min_shelf_life,0) " &
    "FROM tspl_item_master " &
    "WHERE Item_Code='" & objTr.Item_Code & "'"))

                        ' Expiry Date = Batch Date + Min Shelf Life Days
                        Dim ExpiryDate As Date = objTr.Batch_Date.AddDays(MinShelfLife)

                        Dim strQry As String = "UPDATE TSPL_BATCH_ITEM " &
                       "SET Expiry_Date = '" & Format(ExpiryDate, "dd-MMM-yyyy") & "' " &
                       "WHERE Item_Code = '" & objTr.Item_Code & "' " &
                       "AND Batch_No = '" & objTr.Batch_No & "'"

                        clsDBFuncationality.ExecuteNonQuery(strQry)

                        objTr.QC_Status = clsCommon.myCBool(gv1.Rows(ii).Cells(ColQCStatus).Value)
                        objTr.ArrQC = TryCast(gv1.Rows(ii).Cells(ColQCStatus).Tag, List(Of clsDairyProductionUploaderQC))
                        obj.Arr.Add(objTr)
                    End If
                Next
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    Throw New Exception("Please Fill at list one Item")
                End If
                obj.SaveData(obj, isNewEntry)
                clsCommon.MyMessageBoxShow(Me, "Data saved successfully", Me.Text)
                LoadData(obj.Document_No, NavigatorType.Current)
            End If
        Catch ex As Exception
            frmSRN.IsPoSavedAuto = False
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Function

    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            isInsideLoadData = True
            LoadBlankGrid()
            Dim obj As New clsDairyProductionUploader()
            obj = clsDairyProductionUploader.GetData(strCode, NavTyep, Nothing, False)

            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                isNewEntry = False
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                Else
                    btnSave.Enabled = True
                    btnPost.Enabled = True
                    btnDelete.Enabled = True
                End If
                txtDocNo.Value = obj.Document_No
                txtDate.Value = obj.Document_Date
                txtDesc.Text = obj.Description
                UsLock1.Status = obj.Status
                txtLocationFG.Value = obj.Location_FG
                lblLocationFG.Text = obj.Location_FG_Name

                txtLocationRM.Value = obj.Location_RM
                lblLocationRM.Text = obj.Location_RM_Name

                txtLocationPK.Value = obj.Location_PK
                lblLocationPK.Text = obj.Location_PK_Name
                txtBatchNo.Text = obj.Batch_No
                txtBatchDate.Value = obj.Batch_Date
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsDairyProductionUploaderDetail In obj.Arr
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColPKID).Value = objTr.PK_ID
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColSNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = objTr.Batch_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchDate).Value = objTr.Batch_Date
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemCode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemName).Value = objTr.Item_Name
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUOM).Value = objTr.UOM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colShift).Value = objTr.Shift_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColQCStatus).Value = objTr.QC_Status
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColQCStatus).Tag = objTr.ArrQC
                    Next
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Please select document no to delete")
            End If
            If clsCommon.MyMessageBoxShow("Delete the current document" + Environment.NewLine + "Are you sure ? ", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                clsDairyProductionUploader.DeleteData(txtDocNo.Value)
                clsCommon.MyMessageBoxShow(Me, "Data delete successfully", Me.Text)
                AddNew()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("No document found to post")
            End If
            If clsCommon.MyMessageBoxShow("Post the Current Document [" + txtDocNo.Value + "]" + Environment.NewLine + "Are You Sure.", Me.Text, MessageBoxButtons.YesNo, WinControls.RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

                clsDairyProductionUploader.PostData(txtDocNo.Value)

                clsCommon.MyMessageBoxShow("Data posted successfully", Me.Text)
                LoadData(txtDocNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadMenuItem4_Click(sender As Object, e As EventArgs) Handles RadMenuItem4.Click
        Dim qry As String = "select '01/JAN/2020' as Date, null as 'Item Code',null as 'Qty',null as 'UOM',null as 'Shift'"
        transportSql.ExporttoExcelWithoutFilter(qry, "", "", Me)
    End Sub

    Private Sub RadMenuItem5_Click(sender As Object, e As EventArgs) Handles RadMenuItem5.Click
        Try
            Dim gv1 As New UserControls.MyRadGridView
            Me.Controls.Add(gv1)
            If transportSql.importExcel(gv1, "Date", "Item Code", "Qty", "UOM", "Shift") Then
                Dim Arr As New List(Of clsDairyProductionUploaderDetail)
                Dim ii As Integer = 0
                Dim dtt As DataTable = TryCast(gv1.DataSource, DataTable)
                dtt.Columns.Add("ErrorDesc", "".GetType())
                Try
                    Dim qry As String = ""
                    Dim ErrCount As Integer = 0
                    clsCommon.ProgressBarShow()
                    For ii = 0 To gv1.RowCount - 1
                        If clsCommon.myLen(gv1.Rows(ii).Cells("Item Code").Value) > 0 Then
                            Dim objTr As New clsDairyProductionUploaderDetail()
                            objTr.Batch_Date = clsCommon.myCDate(gv1.Rows(ii).Cells("Date").Value)
                            objTr.Item_Code = clsCommon.myCstr(gv1.Rows(ii).Cells("Item Code").Value)
                            If clsCommon.myLen(objTr.Item_Code) > 0 Then
                                qry = "select item_code from TSPL_ITEM_MASTER where item_code='" + objTr.Item_Code + "'"
                                objTr.Item_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                                If clsCommon.myLen(objTr.Item_Code) <= 0 Then
                                    dtt.Rows(ii)("ErrorDesc") = "Invalid Item code " + clsCommon.myCstr(gv1.Rows(ii).Cells("Item Code").Value) & " at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                    ErrCount = ErrCount + 1 : GoTo ExitLOOP
                                End If
                            Else
                                dtt.Rows(ii)("ErrorDesc") = "Pleae enter Item code at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                ErrCount = ErrCount + 1 : GoTo ExitLOOP
                            End If
                            objTr.Qty = clsCommon.myCdbl(gv1.Rows(ii).Cells("Qty").Value)

                            objTr.UOM = clsCommon.myCstr(gv1.Rows(ii).Cells("UOM").Value)
                            If clsCommon.myLen(objTr.UOM) > 0 Then
                                qry = "select UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_Code='" + objTr.UOM + "' and Item_Code='" + objTr.Item_Code + "'"
                                objTr.UOM = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                                If clsCommon.myLen(objTr.UOM) <= 0 Then
                                    dtt.Rows(ii)("ErrorDesc") = "Invalid UOM " + clsCommon.myCstr(gv1.Rows(ii).Cells("UOM").Value) & " at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                    ErrCount = ErrCount + 1 : GoTo ExitLOOP
                                End If
                            Else
                                dtt.Rows(ii)("ErrorDesc") = "Pleae enter UOM at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                ErrCount = ErrCount + 1 : GoTo ExitLOOP
                            End If

                            objTr.Shift_Code = clsCommon.myCstr(gv1.Rows(ii).Cells("Shift").Value)
                            If clsCommon.myLen(objTr.Shift_Code) > 0 Then
                                qry = "select SHIFT_CODE from tspl_shift_master where SHIFT_CODE='" + objTr.Shift_Code + "'"
                                objTr.Shift_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                                If clsCommon.myLen(objTr.Shift_Code) <= 0 Then
                                    dtt.Rows(ii)("ErrorDesc") = "Invalid Shift_Code " + clsCommon.myCstr(gv1.Rows(ii).Cells("Shift").Value) & " at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                    ErrCount = ErrCount + 1 : GoTo ExitLOOP
                                End If
                            Else
                                dtt.Rows(ii)("ErrorDesc") = "Pleae enter Shift at Line No :" & clsCommon.myCstr(ii + 2) & " "
                                ErrCount = ErrCount + 1 : GoTo ExitLOOP
                            End If
                            Arr.Add(objTr)
                        End If
ExitLOOP:
                    Next

                    dtt.DefaultView.RowFilter = "ErrorDesc<>''"
                    dtt = dtt.DefaultView.ToTable

                    If dtt.Rows.Count > 0 Then
                        clsCommon.ProgressBarHide()
                        common.clsCommon.MyMessageBoxShow("Error in " & dtt.Rows.Count & " Records.", Me.Text, MessageBoxButtons.OK)
                        Dim ff As New FrmFreeGrid
                        ff.ReportID = "UnImportedList"
                        ff.Text = "Record Could not Loaded"
                        ff.dt = dtt
                        ff.ShowDialog()
                        Exit Sub
                    End If
                    clsCommon.ProgressBarUpdate("Loading data in Grid.Please wait...")
                    AddRowsByImportAfterValidation(Arr, False)
                Catch ex As Exception
                    clsCommon.ProgressBarHide()
                    Throw New Exception("Error at Row No" + clsCommon.myCstr(ii + 1) + Environment.NewLine + ex.Message)
                Finally
                    clsCommon.ProgressBarHide()
                End Try
            End If
            Me.Controls.Remove(gv1)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadMenuItem6_Click(sender As Object, e As EventArgs)
        Try
            'If Not System.IO.Directory.Exists("C:\\ERPTempFolder") Then
            '    System.IO.Directory.CreateDirectory("C:\\ERPTempFolder")
            'End If

            'If Not File.Exists(Application.StartupPath + "\XpertBennyDecrptor.exe") Then
            '    LocalException("Please add File - XpertBennyDecrptor.exe.")
            'End If
            'Dim qry As String
            'Dim strOPFile As String = "C:\\ERPTempFolder\BSP.CSV"
            'Dim ofd As FolderBrowserDialog = New FolderBrowserDialog()
            'If clsCommon.myLen(strFolderPath) > 0 Then
            '    ofd.SelectedPath = strFolderPath
            'End If
            'If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '    strFolderPath = ofd.SelectedPath
            '    Dim totalFiles As Decimal = Directory.GetFiles(ofd.SelectedPath, "*.BDF").Count
            '    Dim FilesCounter As Integer = 1
            '    Dim Arr As New List(Of clsDairyProductionUploaderDetail)
            '    clsCommon.ProgressBarPercentShow()
            '    Try
            '        For Each FileName As String In Directory.GetFiles(ofd.SelectedPath, "*.BDF")
            '            Try
            '                If System.IO.File.Exists(strOPFile) Then
            '                    File.Delete(strOPFile)
            '                End If
            '                Process.Start(Application.StartupPath + "\XpertBennyDecrptor.exe", "-i " + FileName + " -o " + strOPFile + " -s , -f")
            '                Dim dt As DataTable = transportSql.GetExcelData(strOPFile, Path.GetFileName(FileName))

            '                For ii As Integer = 0 To dt.Rows.Count - 1
            '                    If clsCommon.myLen(dt.Rows(ii)("Extended_Code")) > 0 Then
            '                        If clsCommon.myLen(fndMCCCode.Value) <= 0 Then
            '                            Dim dtMCC As DataTable = clsDBFuncationality.GetDataTable("select MCC_Code,Mcc_Name from TSPL_Mcc_MASTER where MCC_Code='" + clsCommon.myCstr(dt.Rows(ii)("CP_CODE")) + "' ")
            '                            If dtMCC Is Nothing OrElse dtMCC.Rows.Count <= 0 Then
            '                                LocalException("Not a valid MCC " + clsCommon.myCstr(dt.Rows(ii)("CP_CODE")))
            '                            Else
            '                                fndMCCCode.Value = clsCommon.myCstr(dtMCC.Rows(0)("MCC_Code"))
            '                                LblMccName.Text = clsCommon.myCstr(dtMCC.Rows(0)("Mcc_Name"))
            '                            End If
            '                        End If
            '                        If Not clsCommon.CompairString(clsCommon.myCstr(dt.Rows(ii)("CP_CODE")), fndMCCCode.Value) = CompairStringResult.Equal Then
            '                            LocalException("MCC Should be same for all collection.Selected MCC " + fndMCCCode.Value + "And reading MCC-" + clsCommon.myCstr(dt.Rows(ii)("CP_CODE")))
            '                        End If

            '                        Dim objTr As New clsDairyProductionUploaderDetail()
            '                        objTr.SNo = ii + 1
            '                        objTr.Shift_Date = clsCommon.myCDate(dt.Rows(ii)("Date"))
            '                        objTr.Shift = clsCommon.myCstr(dt.Rows(ii)("SHIFT"))
            '                        qry = "select VLC_Code from TSPL_VLC_MASTER_HEAD where MCC='" + fndMCCCode.Value + "' and VLC_Code_VLC_Uploader='" + clsCommon.myCstr(dt.Rows(ii)("Extended_Code")) + "'"
            '                        objTr.VLC_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            '                        If clsCommon.myLen(objTr.VLC_Code) <= 0 Then
            '                            LocalException("Invalid VSP Uploader code " + clsCommon.myCstr(dt.Rows(ii)("Extended_Code")))
            '                        End If
            '                        objTr.Milk_Weight = clsCommon.myCdbl(dt.Rows(ii)("Quantity"))
            '                        objTr.No_Of_Cans = Math.Ceiling(objTr.Milk_Weight / MilkWeight_Setting)
            '                        objTr.FAT = clsCommon.myCdbl(dt.Rows(ii)("FAT"))
            '                        objTr.SNF = clsCommon.myCdbl(dt.Rows(ii)("SNF"))
            '                        objTr.Dock_Collection_Milk_Type = clsCommon.myCstr(dt.Rows(ii)("MILK_Type"))
            '                        If clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "MIXED") = CompairStringResult.Equal Then
            '                            objTr.Dock_Collection_Milk_Type = "M"
            '                        ElseIf clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "COW") = CompairStringResult.Equal Then
            '                            objTr.Dock_Collection_Milk_Type = "C"
            '                        ElseIf clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "Buffalo") = CompairStringResult.Equal Then
            '                            objTr.Dock_Collection_Milk_Type = "B"
            '                        End If
            '                        If Not (clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "M") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "C") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.Dock_Collection_Milk_Type, "B") = CompairStringResult.Equal) Then
            '                            LocalException("Milk Type can be M,B or C")
            '                        End If
            '                        Arr.Add(objTr)
            '                    End If
            '                Next
            '                dt = Nothing
            '            Catch ex As Exception
            '                LocalException("Error in file " + FileName + Environment.NewLine + ex.Message)
            '            End Try
            '            clsCommon.ProgressBarPercentUpdate(((FilesCounter) * 100 / totalFiles), "Reading File " + clsCommon.myCstr(FilesCounter) & "/" & clsCommon.myCstr(totalFiles) & "")
            '            FilesCounter += 1
            '        Next
            '        AddRowsByImportAfterValidation(Arr, True)
            '        clsCommon.ProgressBarPercentHide()
            'Catch ex As Exception
            '    clsCommon.ProgressBarPercentHide()
            '    LocalException(ex.Message)
            'End Try
            '    End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LocalException(ByVal str As String)
        Throw New Exception(str)
    End Sub

    Sub AddRowsByImportAfterValidation(ByVal Arr As List(Of clsDairyProductionUploaderDetail), ByVal isShowPercent As Boolean)
        Try
            isInsideLoadData = True
            If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                LoadBlankGrid()
                Dim ii As Decimal = 1
                For Each objTr As clsDairyProductionUploaderDetail In Arr
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColSNo).Value = gv1.Rows.Count
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchDate).Value = objTr.Batch_Date
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colItemCode).Value = objTr.Item_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colItemName).Value = objTr.Item_Name
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUOM).Value = objTr.UOM
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colShift).Value = objTr.Shift_Code
                    If isShowPercent Then
                        clsCommon.ProgressBarPercentUpdate(((ii) * 100 / Arr.Count), "Loading data in grid" + clsCommon.myCstr(ii) + "/" + clsCommon.myCstr(Arr.Count) + "")
                    Else
                        clsCommon.ProgressBarUpdate("Loading data in grid " + clsCommon.myCstr(ii) + "/" + clsCommon.myCstr(Arr.Count) + "")
                    End If
                    ii += 1
                    gv1.Refresh()
                Next
            End If
        Catch ex As Exception
        Finally
            isInsideLoadData = False
        End Try
    End Sub

    Private Sub txtLocationFG__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocationFG._MYValidating
        If clsCommon.myLen(arrLoc) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No location rights.", Me.Text)
            Exit Sub
        End If
        'txtLocationFG.Value = clsLocation.getFinder(" tspl_location_master.location_code in (" + arrLoc + ") and isnull(csa_type,'N')<>'Y' and isnull(Is_Section,'N')<>'Y' and isnull(Is_Sub_Location,'N')<>'Y' and Location_Category<>'MCC'", txtLocationFG.Value, isButtonClicked)
        txtLocationFG.Value = clsLocation.getFinder(" tspl_location_master.location_code in (" + arrLoc + ") and isnull(csa_type,'N')<>'Y' and isnull(Is_Section,'N')<>'Y'  and Location_Category<>'MCC'", txtLocationFG.Value, isButtonClicked)
        lblLocationFG.Text = clsLocation.GetName(txtLocationFG.Value, Nothing)
    End Sub

    Private Sub txtLocationRM__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocationRM._MYValidating
        If clsCommon.myLen(arrLoc) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No location rights.", Me.Text)
            Exit Sub
        End If
        txtLocationRM.Value = clsLocation.getFinder(" tspl_location_master.location_code in (" + arrLoc + ") and isnull(csa_type,'N')<>'Y' and isnull(Is_Section,'N')<>'Y' and isnull(Is_Sub_Location,'N')<>'N' and Location_Category<>'MCC'", txtLocationRM.Value, isButtonClicked)
        lblLocationRM.Text = clsLocation.GetName(txtLocationRM.Value, Nothing)
    End Sub

    Private Sub txtLocationPK__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocationPK._MYValidating
        If clsCommon.myLen(arrLoc) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No location rights.", Me.Text)
            Exit Sub
        End If
        'txtLocationPK.Value = clsLocation.getFinder(" tspl_location_master.location_code in (" + arrLoc + ") and isnull(csa_type,'N')<>'Y' and isnull(Is_Section,'N')<>'Y' and isnull(Is_Sub_Location,'N')<>'Y' and Location_Category<>'MCC'", txtLocationPK.Value, isButtonClicked)
        txtLocationPK.Value = clsLocation.getFinder(" tspl_location_master.location_code in (" + arrLoc + ") and isnull(csa_type,'N')<>'Y' and isnull(Is_Section,'N')<>'Y'  and Location_Category<>'MCC'", txtLocationPK.Value, isButtonClicked)
        lblLocationPK.Text = clsLocation.GetName(txtLocationPK.Value, Nothing)
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityold.ShowTransHistoryData(txtDocNo.Value, "Document_No", "TSPL_PRODUCTION_UPLOADER_HEAD", "TSPL_PRODUCTION_UPLOADER_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If common.clsCommon.MyMessageBoxShow("Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                If clsDairyProductionUploader.ReverseAndUnpost(txtDocNo.Value) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnShowInventory_Click(sender As Object, e As EventArgs) Handles btnShowInventory.Click
        clsOpenInventory.ShowInventoryDatails(txtDocNo.Value)
    End Sub

    Private Sub gv1_CellDoubleClick(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellDoubleClick
        Try
            Dim qry As String
            If e.Column Is gv1.Columns(ColQCStatus) Then
                If clsCommon.myLen(gv1.CurrentRow.Cells(colItemCode).Value) > 0 Then
                    Dim Arr As List(Of clsDairyProductionUploaderQC) = TryCast(gv1.CurrentRow.Cells(ColQCStatus).Tag, List(Of clsDairyProductionUploaderQC))
                    qry = "select Code as [QC Code],cast(Actual_Range as varchar) as [Standard Range],'' as [Value] From TSPL_ITEM_QC_PARAMETER_MASTER where Item_Code like '" + clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value) + "'
union all
select 'QC Status' as  [QC Code],'Y/N' as [Standard Range],'" + IIf(clsCommon.myCBool(gv1.CurrentRow.Cells(ColQCStatus).Value), "Y", "N") + "' as [Value] "
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                        Throw New Exception("QC Parameter not defined for item [" + clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value) + "]")
                    End If
                    If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                        For ii As Integer = 0 To dt.Rows.Count - 1
                            For Each objtr As clsDairyProductionUploaderQC In Arr
                                If clsCommon.CompairString(objtr.QC_Code, clsCommon.myCstr(dt.Rows(ii)("QC Code"))) = CompairStringResult.Equal Then
                                    dt.Rows(ii)("Value") = objtr.Value
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                    Dim frm As New FrmFreeGrid
                    frm.dt = dt
                    frm.arrEditableColumn = New List(Of String)
                    frm.arrEditableColumn.Add("Value")
                    frm.strFormName = "Fill QC Details of item [" + clsCommon.myCstr(gv1.CurrentRow.Cells(colItemCode).Value) + "]"
                    frm.ReportID = Me.Form_ID
                    frm.WindowState = FormWindowState.Normal
                    frm.ShowDialog()
                    If frm.dt IsNot Nothing AndAlso frm.dt.Rows.Count > 0 Then
                        Dim ArrTemp As New List(Of clsDairyProductionUploaderQC)
                        Dim obj As clsDairyProductionUploaderQC = Nothing
                        For ii As Integer = 0 To frm.dt.Rows.Count - 1
                            If ii = frm.dt.Rows.Count - 1 Then
                                gv1.CurrentRow.Cells(ColQCStatus).Value = (clsCommon.CompairString(clsCommon.myCstr(frm.dt.Rows(ii)("Value")), "Y") = CompairStringResult.Equal)
                                Exit For
                            End If
                            obj = New clsDairyProductionUploaderQC()
                            obj.QC_Code = clsCommon.myCstr(frm.dt.Rows(ii)("QC Code"))
                            obj.Value = clsCommon.myCstr(frm.dt.Rows(ii)("Value"))
                            ArrTemp.Add(obj)
                        Next
                        gv1.CurrentRow.Cells(ColQCStatus).Tag = ArrTemp
                    Else
                        gv1.CurrentRow.Cells(ColQCStatus).Tag = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrintNew_Click(sender As Object, e As EventArgs) Handles btnPrintNew.Click
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No data found to Print", Me.Text)
        Else
            funPrint(False, txtDate.Value, txtDocNo.Value)
        End If
    End Sub
    Public Sub funPrint(ByVal isCancel As Boolean, ByVal strDate As DateTime, ByVal strCode As String)
        Try
            Dim Qry = clsDairyProductionUploader.GetAttachQry(Form_ID, False, strDate, strCode)
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'If dt.Rows.Count > 0 Then
            '    Dim frmCRV As New frmCrystalReportViewer()
            '    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PRODUCTION, dt, "crptProductionUploaderPrint", "")
            'Else
            '    clsCommon.MyMessageBoxShow(Me, "No data found to Print", Me.Text)
            'End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

End Class
