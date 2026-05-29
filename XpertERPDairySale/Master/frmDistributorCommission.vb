Imports common
Public Class frmDistributorCommission
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim isNewEntry As Boolean = False
    Dim SeparateDemandMilkandProduct As Boolean = False
    Dim EnableProductSaleForJPR As Boolean = False
    Dim EnableVehicleType As Boolean = False
    Const ColSNo As String = "ColSNo"
    Const ColRouteCode As String = "ColRouteCode"
    Const ColRouteName As String = "ColRouteName"
    Const colCustCode As String = "colCustCode"
    Const colCustName As String = "colCustName"
    Const colCRate As String = "colCRate"
    Const colTRate As String = "colTRate"
    Const colSecRate As String = "colSecRate"
    Const colBoothSecRate As String = "colBoothSecRate"
    Private isCellValueChangedOpen As Boolean = False
    Private isInsideLoadData As Boolean = False
#End Region
    Public Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.frmBookingProductSale)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPrint.Visible = MyBase.isPrintFlag
        ' btnImport.Visible = MyBase.isExport
        btnReverse.Visible = False
        If MyBase.isExport = True Then
            rmiImport.Enabled = True
            rmiExport.Enabled = True
        Else
            rmiImport.Enabled = False
            rmiExport.Enabled = False
        End If
        'If MyBase.isReverse Then
        '    btnReverse.Enabled = True
        'Else
        '    btnReverse.Enabled = False
        'End If
        'If btnSave.Visible = True Then
        '    btnImport.Enabled = True
        'Else
        '    btnImport.Enabled = False
        'End If
    End Sub
    Private Sub frmDistributorCommission_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SeparateDemandMilkandProduct = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SeparateDemandMilkandProduct, clsFixedParameterCode.SeparateDemandMilkandProduct, Nothing)) = 1, True, False)
        EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)
        EnableVehicleType = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableVehicleType, clsFixedParameterCode.EnableVehicleType, Nothing)) = 1, True, False)
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtApplicableDate.Value = clsCommon.GETSERVERDATE()
        txtInActiveDate.Value = clsCommon.GETSERVERDATE()
        SetUserMgmtNew()
        AddNew()

    End Sub
    Sub LoadType()
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = Nothing

        dr = dt.NewRow()
        dr("Code") = "M"
        dr("Name") = "Milk"
        dt.Rows.Add(dr)

        If EnableProductSaleForJPR Then

            dr = dt.NewRow()
            dr("Code") = "P"
            dr("Name") = "Product"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "I"
            dr("Name") = "Ice Cream"
            dt.Rows.Add(dr)

        End If
        dr = dt.NewRow()
        dr("Code") = "O"
        dr("Name") = "Other"
        dt.Rows.Add(dr)

        cmbItemType.DataSource = dt
        cmbItemType.ValueMember = "Code"
        cmbItemType.DisplayMember = "Name"
    End Sub
    Sub LoadVehicleType()
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = Nothing

        dr = dt.NewRow()
        dr("Code") = "AC"
        dr("Name") = "A/C"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Non-AC"
        dr("Name") = "Non A/C"
        dt.Rows.Add(dr)
            dr = dt.NewRow()
        dr("Code") = "Open"
        dr("Name") = "Open"
        dt.Rows.Add(dr)


        cmbVehicleType.DataSource = dt
        cmbVehicleType.ValueMember = "Code"
        cmbVehicleType.DisplayMember = "Name"
    End Sub
    Private Sub frmDistributorCommission_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
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
    Sub LoadBlankGrid()
        GV1.Rows.Clear()
        GV1.Columns.Clear()
        Dim repoNumBox As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoNumBox = New GridViewDecimalColumn()
        repoNumBox.FormatString = ""
        repoNumBox.HeaderText = "SNo"
        repoNumBox.Name = ColSNo
        repoNumBox.Width = 40
        repoNumBox.ReadOnly = True
        repoNumBox.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        GV1.MasterTemplate.Columns.Add(repoNumBox)

        Dim repoRouteCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRouteCode.FormatString = ""
        repoRouteCode.HeaderText = "Route Code"
        repoRouteCode.Name = ColRouteCode
        repoRouteCode.IsVisible = True
        repoRouteCode.ReadOnly = True
        repoRouteCode.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        GV1.MasterTemplate.Columns.Add(repoRouteCode)

        Dim repoRouteName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRouteName.FormatString = ""
        repoRouteName.HeaderText = "Route Name"
        repoRouteName.Name = ColRouteName
        repoRouteName.IsVisible = True
        repoRouteName.ReadOnly = True
        repoRouteName.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        GV1.MasterTemplate.Columns.Add(repoRouteName)

        Dim repoCustCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCustCode.FormatString = ""
        If rbtnCommission.IsChecked Then
            repoCustCode.HeaderText = "Distributor/Transpoter Code"
        Else
            repoCustCode.HeaderText = "Transpoter Code"
        End If
        repoCustCode.Name = colCustCode
        repoCustCode.Width = 150
        repoCustCode.IsVisible = True
        repoCustCode.ReadOnly = True
        repoCustCode.TextImageRelation = TextImageRelation.TextBeforeImage
        'repoCustCode.ReadOnly = True
        GV1.MasterTemplate.Columns.Add(repoCustCode)

        Dim repoCustName = New GridViewTextBoxColumn()
        repoCustName.FormatString = ""
        If rbtnCommission.IsChecked Then
            repoCustName.HeaderText = "Distributor/Transpoter Name"
        Else
            repoCustName.HeaderText = "Transpoter Name"
        End If
        repoCustName.Name = colCustName
        repoCustName.Width = 200
        repoCustName.IsVisible = True
        repoCustName.ReadOnly = True
        GV1.MasterTemplate.Columns.Add(repoCustName)

        Dim repoTextBox = New GridViewDecimalColumn()
        repoTextBox.FormatString = "{0:n4}"
        repoTextBox.HeaderText = "Distributor Rate Per UOM"
        repoTextBox.Name = colCRate
        repoTextBox.Width = 120
        repoTextBox.Minimum = 0
        repoNumBox.ShowUpDownButtons = False
        repoNumBox.Step = 0
        repoTextBox.DecimalPlaces = 4
        repoTextBox.IsVisible = True
        'repoTextBox.ReadOnly = True
        GV1.MasterTemplate.Columns.Add(repoTextBox)

        Dim repoTRate = New GridViewDecimalColumn()
        repoTRate.FormatString = "{0:n4}"
        repoTRate.HeaderText = "Transporter Rate Per UOM"
        repoTRate.Name = colTRate
        repoTRate.Width = 120
        repoTRate.Minimum = 0
        repoTRate.ShowUpDownButtons = False
        repoTRate.Step = 0
        repoTRate.DecimalPlaces = 4
        repoTRate.IsVisible = True
        'repoTextBox.ReadOnly = True
        GV1.MasterTemplate.Columns.Add(repoTRate)

        If chkSecurity.Checked Then
            Dim SecRateTextBox = New GridViewDecimalColumn()
            SecRateTextBox.FormatString = "{0:n2}"
            SecRateTextBox.HeaderText = "Security Rate %"
            SecRateTextBox.Name = colSecRate
            SecRateTextBox.Width = 80
            SecRateTextBox.Minimum = 0
            SecRateTextBox.ShowUpDownButtons = False
            SecRateTextBox.Step = 0
            SecRateTextBox.DecimalPlaces = 2
            SecRateTextBox.IsVisible = True
            GV1.MasterTemplate.Columns.Add(SecRateTextBox)


            Dim SecBoothRateTextBox = New GridViewDecimalColumn()
            SecBoothRateTextBox.FormatString = "{0:n2}"
            SecBoothRateTextBox.HeaderText = "Booth Security Rate %"
            SecBoothRateTextBox.Name = colBoothSecRate
            SecBoothRateTextBox.Width = 80
            SecBoothRateTextBox.Minimum = 0
            SecBoothRateTextBox.ShowUpDownButtons = False
            SecBoothRateTextBox.Step = 0
            SecBoothRateTextBox.DecimalPlaces = 2
            SecBoothRateTextBox.IsVisible = True
            GV1.MasterTemplate.Columns.Add(SecBoothRateTextBox)
        End If
        GV1.Enabled = True
        GV1.AllowAddNewRow = False
        GV1.ShowGroupPanel = False
        GV1.AllowColumnReorder = False
        GV1.AllowRowReorder = False
        GV1.EnableSorting = False
        GV1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        GV1.MasterTemplate.ShowRowHeaderColumn = False
        GV1.TableElement.TableHeaderHeight = 40
        GV1.AllowDeleteRow = True
        GV1.Rows.AddNew()
    End Sub
    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub
    Public Sub AddNew()
        isNewEntry = True
        lblStatus.Status = ERPTransactionStatus.Pending
        txtDocNo.Value = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtInActiveDate.Value = clsCommon.GETSERVERDATE()
        txtApplicableDate.Value = txtDate.Value
        rbtnCommission.IsChecked = True
        rbtnTranspotation.IsChecked = False
        rbtnTranspotation.Visible = False
        chkSecurity.Checked = False
        txtDistributorTagging.Value = ""
        txtItems.arrValueMember = Nothing
        txtItems.arrDispalyMember = Nothing
        txtUOM.Value = ""
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnGo.Enabled = True
        btnDelete.Enabled = False
        chkInActive.Checked = False
        'cmbItemType.Enabled = False
        LoadType()
        LoadVehicleType()
        LoadBlankGrid()
        If EnableVehicleType Then
            lblVehicleType.Visible = True
            cmbVehicleType.Visible = True
            cmbVehicleType.SelectedValue = "AC"
        Else
            lblVehicleType.Visible = False
            cmbVehicleType.Visible = False
        End If
    End Sub
    Private Sub txtUOM__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtUOM._MYValidating
        Try
            Dim qry As String = "select Unit_Code as Code,Unit_Desc as Description from TSPL_UNIT_MASTER"
            txtUOM.Value = clsCommon.ShowSelectForm("fndUOMMaster@CustomerWiseSalesReport", qry, "Code", "", txtUOM.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtItems__My_Click(sender As Object, e As EventArgs) Handles txtItems._My_Click
        Try
            Dim qry As String = ""
            qry = " select Item_code as Code,Item_Desc,Short_Description  from TSPL_ITEM_MASTER where  "
            If clsCommon.CompairString(cmbItemType.SelectedValue, "M") = CompairStringResult.Equal Then
                qry += " TypeOfItm='M' "
            ElseIf clsCommon.CompairString(cmbItemType.SelectedValue, "P") = CompairStringResult.Equal Then
                qry += " TypeOfItm='P' "
            ElseIf clsCommon.CompairString(cmbItemType.SelectedValue, "I") = CompairStringResult.Equal Then
                qry += " TypeOfItm='I' "
            Else
                qry += " TypeOfItm='O' "
            End If
            txtItems.arrValueMember = clsCommon.ShowMultipleSelectForm("CustomerCode@CustWiseSaleRpt", qry, "Code", "Code", txtItems.arrValueMember, txtItems.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            LoadBlankGrid()
            If clsCommon.myLen(txtUOM.Value) > 0 Then
                If clsCommon.myLen(txtDistributorTagging.Value) > 0 Then
                    Dim obj As New clsDistributorCommission
                    If txtItems.arrValueMember Is Nothing Then
                        Dim strItems As String = "select Item_code  from TSPL_ITEM_MASTER where Item_Type='F'"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(strItems)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            obj.Items = New ArrayList()
                            For Each dr As DataRow In dt.Rows
                                obj.Items.Add(clsCommon.myCstr(dr("Item_Code")))
                            Next
                        End If
                    Else
                        obj.Items = txtItems.arrValueMember
                    End If
                    For Each items In obj.Items
                        Dim qry As String = "select * from TSPL_Distributor_commission_Head 
left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_Distributor_commission_Head.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No 
where TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code='" & items & "' and TSPL_Distributor_commission_Head.Distributor_Tagging_Code='" & txtDistributorTagging.Value & "'"
                        If EnableVehicleType Then
                            qry += " and TSPL_Distributor_commission_Head.Vehicle_Type='" & cmbVehicleType.SelectedValue & "'"
                        End If
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            Throw New Exception("Item : " & items & " is already exists with " & txtDistributorTagging.Value & " ! Please select different item/items")
                        End If
                    Next
                    Dim StrQry As String = "select TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No,TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DISTRIBUTOR_ROUTE.ItemType 
from TSPL_DISTRIBUTOR_ROUTE
left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
left join TSPL_CUSTOMER_MASTER on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
where TSPL_DISTRIBUTOR_ROUTE.Code='" + txtDistributorTagging.Value + "' "
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrQry)
                    If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                        Dim i As Integer = 1
                        For Each dr As DataRow In dt1.Rows
                            GV1.Rows(GV1.Rows.Count - 1).Cells(ColSNo).Value = i
                            GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteCode).Value = clsCommon.myCstr(dr("Route_No"))
                            GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteName).Value = clsDBFuncationality.getSingleValue("select Route_Desc from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(dr("Route_No")) + "' ")
                            GV1.Rows(GV1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                            GV1.Rows(GV1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                            cmbItemType.Text = clsCommon.myCstr(dr("ItemType"))
                            i = i + 1
                            GV1.Rows.AddNew()
                        Next
                    Else
                        clsCommon.MyMessageBoxShow(Me, "Not Found!", Me.Text)
                    End If
                Else
                    clsCommon.MyMessageBoxShow(Me, "Please Select Distributor Tagging", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Please Select Commission UOM", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmiImport_Click(sender As Object, e As EventArgs) Handles rmiImport.Click
        Import()
    End Sub
    Private Sub rmiExport_Click(sender As Object, e As EventArgs) Handles rmiExport.Click
        Export()
    End Sub
    '   Public Sub ImportTPT()
    '       Try
    '           Dim gv As New UserControls.MyRadGridView
    '           Me.Controls.Add(gv)
    '           Dim obj As New List(Of clsDistributorCommissionDetails)
    '           Dim currentdate As Date = Date.Today
    '           If clsCommon.myLen(txtUOM.Value) > 0 Then
    '               If transportSql.importExcel(gv, "Docuemnt Code", "Invoice No", "Customer Code") Then
    '                   'Dim trans As SqlTransaction = Nothing
    '                   Dim linno As Integer = 0
    '                   Dim TempNewRecord As Boolean = False
    '                   Try
    '                       'trans = clsDBFuncationality.GetTransactin()
    '                       clsCommon.ProgressBarPercentShow()
    '                       For Each grow As GridViewRowInfo In gv.Rows
    '                           linno += 1
    '                           ''Updating Shipment Detail
    '                           ''update Invoice Head
    '                           ''-------------------
    '                           Dim strQry As String = " update TSPL_SD_SHIPMENT_DETAIL set Distributor_Commission_PKID=xxx.DisPKID,Transporter_Commission_Rate=xxx.Rate,Distributor_Commission_RateWithTax=xxx.Rate,
    '   Transporter_Commission_Amt=xxx.Dis_Commision_Amt
    'from (
    '   select xx.Customer_Code,xx.Item_Code,xx.PK_ID,xx.QtyInCommission_Uom,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate as Rate,(xx.QtyInCommission_Uom * TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate) as Dis_Commision_Amt,xx.DisPKID
    '   from
    '   (
    '   select CurrentUOMCF.Conversion_Factor as CurrUintCF,LTRUOMMCF.Conversion_Factor as LtrUintCF,TSPL_SD_Shipment_Head.Customer_Code, TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Qty,(TSPL_SD_SHIPMENT_DETAIL.Qty*CurrentUOMCF.Conversion_Factor /LTRUOMMCF.Conversion_Factor) as QtyInCommission_Uom,TSPL_SD_SHIPMENT_DETAIL.PK_ID,
    '   (select top 1 TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID from TSPL_DISTRIBUTOR_COMMISSION_HEAD
    '   left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    '   left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    '   where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<=TSPL_SD_SHIPMENT_HEAD.Document_Date and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'  and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code=TSPL_SD_SHIPMENT_HEAD.Route_No
    '   order by TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date desc,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No desc) as DisPKID
    '   from TSPL_SD_SHIPMENT_HEAD
    '   left join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE
    '   left join TSPL_ITEM_UOM_DETAIL CurrentUOMCF on CurrentUOMCF.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and CurrentUOMCF.UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code
    '   left join TSPL_ITEM_UOM_DETAIL LTRUOMMCF on LTRUOMMCF.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and LTRUOMMCF.UOM_Code='LTR'
    '   where TSPL_SD_SHIPMENT_HEAD.Document_Code='" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value) + "'
    '   )xx 
    '   left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID=XX.DisPKID
    '   )xxx inner join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.pk_id=xxx.PK_ID  "
    '                           clsDBFuncationality.ExecuteNonQuery(strQry)
    '                           clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Shipment Detail  Doc No:" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value))
    '                           ''update Shipment Head
    '                           ''-------------------
    '                           strQry = "update TSPL_SD_SHIPMENT_HEAD set Transporter_Commission_TotalAmt=xx.Distributor_Commission_Amt,	
    '   Gross_Amount=(Total_Amt-xx.Distributor_Commission_Amt),Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'
    '   from 
    '   (
    '   select TSPL_SD_SHIPMENT_DETAIL.Document_Code, sum(TSPL_SD_SHIPMENT_DETAIL.Transporter_Commission_Amt) as Distributor_Commission_Amt from TSPL_SD_SHIPMENT_HEAD
    '   left join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.Document_Code where TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE='" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value) + "'
    '   group by TSPL_SD_SHIPMENT_DETAIL.Document_Code
    '   )XX
    '   where TSPL_SD_SHIPMENT_HEAD.Document_Code=xx.DOCUMENT_CODE"
    '                           clsDBFuncationality.ExecuteNonQuery(strQry)
    '                           clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Shipment Head Doc No:" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value))
    '                           ''update Invoice Detail
    '                           ''-------------------
    '                           strQry = " update TSPL_SD_SALE_INVOICE_DETAIL set Distributor_Commission_PKID=xxx.DisPKID,Transporter_Commission_Rate=xxx.Rate,
    '   Transporter_Commission_Amt=xxx.Dis_Commision_Amt  from (
    '   select xx.Customer_Code,xx.Item_Code,xx.QtyInCommission_Uom,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate as Rate,(xx.QtyInCommission_Uom * TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate) as Dis_Commision_Amt,xx.DisPKID,xx.Unit_code,xx.Line_No,xx.DOCUMENT_CODE
    '   from
    '   (
    '   select CurrentUOMCF.Conversion_Factor as CurrUintCF,LTRUOMMCF.Conversion_Factor as LtrUintCF,TSPL_SD_SALE_INVOICE_Head.Customer_Code, TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Qty,(TSPL_SD_SALE_INVOICE_DETAIL.Qty*CurrentUOMCF.Conversion_Factor /LTRUOMMCF.Conversion_Factor) as QtyInCommission_Uom ,--TSPL_SD_SALE_INVOICE_DETAIL.PK_ID,
    '   TSPL_SD_SALE_INVOICE_DETAIL.Unit_code,TSPL_SD_SALE_INVOICE_DETAIL.Line_No,TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE,
    '   (select top 1 TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID from TSPL_DISTRIBUTOR_COMMISSION_HEAD
    '   left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    '   left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    '   where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<=TSPL_SD_SALE_INVOICE_Head.Document_Date and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'  and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code=TSPL_SD_SALE_INVOICE_Head.Route_No
    '   order by TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date desc,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No desc) as DisPKID
    '   from TSPL_SD_SALE_INVOICE_Head
    '   left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Head.DOCUMENT_CODE
    '   left join TSPL_ITEM_UOM_DETAIL CurrentUOMCF on CurrentUOMCF.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and CurrentUOMCF.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
    '   left join TSPL_ITEM_UOM_DETAIL LTRUOMMCF on LTRUOMMCF.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and LTRUOMMCF.UOM_Code='LTR'
    '   where TSPL_SD_SALE_INVOICE_Head.Document_Code ='" + clsCommon.myCstr(grow.Cells("Invoice No").Value) + "'
    '   )xx 
    '   left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID=XX.DisPKID
    '   )xxx inner join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.Item_Code=xxx.Item_Code and TSPL_SD_SALE_INVOICE_DETAIL.Line_No=xxx.Line_No and TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=xxx.DOCUMENT_CODE "
    '                           clsDBFuncationality.ExecuteNonQuery(strQry)
    '                           clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Invoice Detail Doc No:" + clsCommon.myCstr(grow.Cells("Invoice No").Value))
    '                           ''update Invoice Head
    '                           ''-------------------
    '                           strQry = " update TSPL_SD_SALE_INVOICE_Head set Transporter_Commission_TotalAmt=xx.Distributor_Commission_Amt,
    '  Gross_Amount=(Total_Amt-xx.Distributor_Commission_Amt),Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'
    '   from 
    '   (
    '   select TSPL_SD_SALE_INVOICE_DETAIL.Document_Code, sum(TSPL_SD_SALE_INVOICE_DETAIL.Transporter_Commission_Amt) as Distributor_Commission_Amt from TSPL_SD_SALE_INVOICE_Head
    '   left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Head.Document_Code where TSPL_SD_SALE_INVOICE_Head.DOCUMENT_CODE ='" + clsCommon.myCstr(grow.Cells("Invoice No").Value) + "'
    '   group by TSPL_SD_SALE_INVOICE_DETAIL.Document_Code
    '   )XX
    '   where TSPL_SD_SALE_INVOICE_Head.Document_Code=xx.DOCUMENT_CODE "
    '                           clsDBFuncationality.ExecuteNonQuery(strQry)
    '                           clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Invoice head Doc No:" + clsCommon.myCstr(grow.Cells("Invoice No").Value))
    '                       Next
    '                       clsCommon.ProgressBarPercentHide()
    '                   Catch ex As Exception
    '                       clsCommon.ProgressBarPercentHide()
    '                       clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '                   End Try
    '               Else
    '                   clsCommon.MyMessageBoxShow(Me, "Excel Sheet is not in expected format", Me.Text)
    '               End If
    '           Else
    '               clsCommon.MyMessageBoxShow(Me, "Please Select Commission UOM", Me.Text)
    '           End If
    '           'clsCommon.ProgressBarHide()
    '           Me.Controls.Remove(gv)
    '       Catch ex As Exception
    '           'clsCommon.ProgressBarHide()
    '           clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '       End Try
    '   End Sub
    'Public Sub Import()
    '    Try
    '        Dim gv As New UserControls.MyRadGridView
    '        Me.Controls.Add(gv)
    '        Dim obj As New List(Of clsDistributorCommissionDetails)
    '        Dim currentdate As Date = Date.Today
    '        If clsCommon.myLen(txtUOM.Value) > 0 Then
    '            If transportSql.importExcel(gv, "Docuemnt Code", "Invoice No", "Customer Code") Then
    '                'Dim trans As SqlTransaction = Nothing
    '                Dim linno As Integer = 0
    '                Dim TempNewRecord As Boolean = False
    '                Try
    '                    'trans = clsDBFuncationality.GetTransactin()
    '                    clsCommon.ProgressBarPercentShow()
    '                    For Each grow As GridViewRowInfo In gv.Rows
    '                        linno += 1
    '                        ''Updating Shipment Detail
    '                        ''update Invoice Head
    '                        ''-------------------
    '                        Dim strQry As String = " update TSPL_SD_SHIPMENT_DETAIL set Distributor_Commission_PKID=xxx.DisPKID,Distributor_Commission_Rate=xxx.Rate,Distributor_Commission_RateWithTax=xxx.Rate,
    'Distributor_Commission_Amt=xxx.Dis_Commision_Amt,Disc_Amt=xxx.Dis_Commision_Amt,Total_Disc_Amt=xxx.Dis_Commision_Amt,Amt_Less_Discount=( Total_Basic_Amt-xxx.Dis_Commision_Amt),Item_Net_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt),TAX1_Base_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt),TAX2_Base_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt)   from (
    'select xx.Customer_Code,xx.Item_Code,xx.PK_ID,xx.QtyInCommission_Uom,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate,(xx.QtyInCommission_Uom * TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate) as Dis_Commision_Amt,xx.DisPKID
    'from
    '(
    'select CurrentUOMCF.Conversion_Factor as CurrUintCF,LTRUOMMCF.Conversion_Factor as LtrUintCF,TSPL_SD_Shipment_Head.Customer_Code, TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Qty,(TSPL_SD_SHIPMENT_DETAIL.Qty*CurrentUOMCF.Conversion_Factor /LTRUOMMCF.Conversion_Factor) as QtyInCommission_Uom,TSPL_SD_SHIPMENT_DETAIL.PK_ID,
    '(select top 1 TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID from TSPL_DISTRIBUTOR_COMMISSION_HEAD
    'left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    'left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    'where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<=TSPL_SD_SHIPMENT_HEAD.Document_Date and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'  and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code=TSPL_SD_SHIPMENT_HEAD.Route_No
    'order by TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date desc,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No desc) as DisPKID
    'from TSPL_SD_SHIPMENT_HEAD
    'left join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE
    'left join TSPL_ITEM_UOM_DETAIL CurrentUOMCF on CurrentUOMCF.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and CurrentUOMCF.UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code
    'left join TSPL_ITEM_UOM_DETAIL LTRUOMMCF on LTRUOMMCF.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and LTRUOMMCF.UOM_Code='LTR'
    'where TSPL_SD_SHIPMENT_HEAD.Document_Code='" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value) + "'
    ')xx 
    'left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID=XX.DisPKID
    ')xxx inner join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.pk_id=xxx.PK_ID "
    '                        clsDBFuncationality.ExecuteNonQuery(strQry)
    '                        clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Shipment Detail  Doc No:" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value))
    '                        ''update Shipment Head
    '                        ''-------------------
    '                        strQry = "update TSPL_SD_SHIPMENT_HEAD set Distributor_Commission_TotalAmt=xx.Distributor_Commission_Amt,Discount_Amt=xx.Distributor_Commission_Amt,
    'Amount_Less_Discount=(Discount_Base-xx.Distributor_Commission_Amt),Total_Amt=(Discount_Base-xx.Distributor_Commission_Amt),Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "',TAX1_Base_Amt=(Discount_Base-xx.Distributor_Commission_Amt),TAX2_Base_Amt=(Discount_Base-xx.Distributor_Commission_Amt)
    'from 
    '(
    'select TSPL_SD_SHIPMENT_DETAIL.Document_Code, sum(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt) as Distributor_Commission_Amt from TSPL_SD_SHIPMENT_HEAD
    'left join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.Document_Code where TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE='" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value) + "'
    'group by TSPL_SD_SHIPMENT_DETAIL.Document_Code
    ')XX
    'where TSPL_SD_SHIPMENT_HEAD.Document_Code=xx.DOCUMENT_CODE"
    '                        clsDBFuncationality.ExecuteNonQuery(strQry)
    '                        clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Shipment Head Doc No:" + clsCommon.myCstr(grow.Cells("Docuemnt Code").Value))
    '                        ''update Invoice Detail
    '                        ''-------------------
    '                        strQry = " update TSPL_SD_SALE_INVOICE_DETAIL set Distributor_Commission_PKID=xxx.DisPKID,Distributor_Commission_Rate=xxx.Rate,Distributor_Commission_RateWithTax=xxx.Rate,
    'Distributor_Commission_Amt=xxx.Dis_Commision_Amt,Disc_Amt=xxx.Dis_Commision_Amt,Total_Disc_Amt=xxx.Dis_Commision_Amt,Amt_Less_Discount=( Total_Basic_Amt-xxx.Dis_Commision_Amt),Item_Net_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt),TAX1_Base_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt),TAX2_Base_Amt=( Total_Basic_Amt-xxx.Dis_Commision_Amt)  from (
    'select xx.Customer_Code,xx.Item_Code,xx.QtyInCommission_Uom,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate,(xx.QtyInCommission_Uom * TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate) as Dis_Commision_Amt,xx.DisPKID,xx.Unit_code,xx.Line_No,xx.DOCUMENT_CODE
    'from
    '(
    'select CurrentUOMCF.Conversion_Factor as CurrUintCF,LTRUOMMCF.Conversion_Factor as LtrUintCF,TSPL_SD_SALE_INVOICE_Head.Customer_Code, TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Qty,(TSPL_SD_SALE_INVOICE_DETAIL.Qty*CurrentUOMCF.Conversion_Factor /LTRUOMMCF.Conversion_Factor) as QtyInCommission_Uom ,--TSPL_SD_SALE_INVOICE_DETAIL.PK_ID,
    'TSPL_SD_SALE_INVOICE_DETAIL.Unit_code,TSPL_SD_SALE_INVOICE_DETAIL.Line_No,TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE,
    '(select top 1 TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID from TSPL_DISTRIBUTOR_COMMISSION_HEAD
    'left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    'left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
    'where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<=TSPL_SD_SALE_INVOICE_Head.Document_Date and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "'  and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code=TSPL_SD_SALE_INVOICE_Head.Route_No
    'order by TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date desc,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No desc) as DisPKID
    'from TSPL_SD_SALE_INVOICE_Head
    'left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Head.DOCUMENT_CODE
    'left join TSPL_ITEM_UOM_DETAIL CurrentUOMCF on CurrentUOMCF.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and CurrentUOMCF.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
    'left join TSPL_ITEM_UOM_DETAIL LTRUOMMCF on LTRUOMMCF.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and LTRUOMMCF.UOM_Code='LTR'
    'where TSPL_SD_SALE_INVOICE_Head.Document_Code ='" + clsCommon.myCstr(grow.Cells("Invoice No").Value) + "'
    ')xx 
    'left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID=XX.DisPKID
    ')xxx inner join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.Item_Code=xxx.Item_Code and TSPL_SD_SALE_INVOICE_DETAIL.Line_No=xxx.Line_No and TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=xxx.DOCUMENT_CODE "
    '                        clsDBFuncationality.ExecuteNonQuery(strQry)
    '                        clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Invoice Detail Doc No:" + clsCommon.myCstr(grow.Cells("Invoice No").Value))
    '                        ''update Invoice Head
    '                        ''-------------------
    '                        strQry = "update TSPL_SD_SALE_INVOICE_Head set Distributor_Commission_TotalAmt=xx.Distributor_Commission_Amt,Discount_Amt=xx.Distributor_Commission_Amt,
    'Amount_Less_Discount=(Discount_Base-xx.Distributor_Commission_Amt),Total_Amt=(Discount_Base-xx.Distributor_Commission_Amt),Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer Code").Value) + "',TAX1_Base_Amt=(Discount_Base-xx.Distributor_Commission_Amt),TAX2_Base_Amt=(Discount_Base-xx.Distributor_Commission_Amt)
    'from 
    '(
    'select TSPL_SD_SALE_INVOICE_DETAIL.Document_Code, sum(TSPL_SD_SALE_INVOICE_DETAIL.Distributor_Commission_Amt) as Distributor_Commission_Amt from TSPL_SD_SALE_INVOICE_Head
    'left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Head.Document_Code where TSPL_SD_SALE_INVOICE_Head.DOCUMENT_CODE ='" + clsCommon.myCstr(grow.Cells("Invoice No").Value) + "'
    'group by TSPL_SD_SALE_INVOICE_DETAIL.Document_Code
    ')XX
    'where TSPL_SD_SALE_INVOICE_Head.Document_Code=xx.DOCUMENT_CODE"
    '                        clsDBFuncationality.ExecuteNonQuery(strQry)
    '                        clsCommon.ProgressBarPercentUpdate(((linno) * 100 / (gv.Rows.Count)), "Updating Invoice head Doc No:" + clsCommon.myCstr(grow.Cells("Invoice No").Value))
    '                    Next
    '                    clsCommon.ProgressBarPercentHide()
    '                Catch ex As Exception
    '                    clsCommon.ProgressBarPercentHide()
    '                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '                End Try
    '            Else
    '                clsCommon.MyMessageBoxShow(Me, "Excel Sheet is not in expected format", Me.Text)
    '            End If
    '        Else
    '            clsCommon.MyMessageBoxShow(Me, "Please Select Commission UOM", Me.Text)
    '        End If
    '        'clsCommon.ProgressBarHide()
    '        Me.Controls.Remove(gv)
    '    Catch ex As Exception
    '        'clsCommon.ProgressBarHide()
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Public Sub Import()
        Try
            Dim gv As New UserControls.MyRadGridView
            Me.Controls.Add(gv)
            Dim obj As New List(Of clsDistributorCommissionDetails)
            Dim currentdate As Date = Date.Today
            If clsCommon.myLen(txtUOM.Value) > 0 Then
                If transportSql.importExcel(gv, "Route Code", "Distributor Code", "Is Transpotation", "Rate", "Transporter Rate", "Security Rate", "Booth Security Rate") Then
                    'Dim trans As SqlTransaction = Nothing
                    Dim linno As Integer = 0
                    Dim TempNewRecord As Boolean = False
                    Try
                        'trans = clsDBFuncationality.GetTransactin()
                        clsCommon.ProgressBarShow()
                        For Each grow As GridViewRowInfo In gv.Rows
                            Dim Arr As New clsDistributorCommissionDetails()
                            linno += 1
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Route Code").Value))) Then
                                Continue For
                            Else
                                Dim str As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Route_No from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(grow.Cells("Route Code").Value) + "'"))
                                If clsCommon.CompairString(str, clsCommon.myCstr(grow.Cells("Route Code").Value)) = CompairStringResult.Equal Then
                                    Arr.Route_Code = clsCommon.myCstr(grow.Cells("Route Code").Value)
                                Else
                                    Continue For
                                End If
                            End If
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Distributor Code").Value))) Or (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Is Transpotation").Value))) Then
                                Continue For
                            Else
                                Dim str As String = String.Empty
                                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Is Transpotation").Value), "Y") = CompairStringResult.Equal Then
                                    str = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select cust_Code from TSPL_Customer_Master where cust_Code='" + clsCommon.myCstr(grow.Cells("Distributor Code").Value) + "' and Form_Type='TPT'"))
                                Else
                                    str = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select cust_Code from TSPL_Customer_Master where cust_Code='" + clsCommon.myCstr(grow.Cells("Distributor Code").Value) + "' and IsDistributor='Y'"))
                                End If
                                If clsCommon.CompairString(str, clsCommon.myCstr(grow.Cells("Distributor Code").Value)) = CompairStringResult.Equal Then
                                    Arr.Distributor_Code = clsCommon.myCstr(grow.Cells("Distributor Code").Value)
                                Else
                                    Continue For
                                End If
                            End If
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Rate").Value))) Then
                                Continue For
                            Else
                                Arr.Rate = clsCommon.myCDecimal(grow.Cells("Rate").Value)
                            End If
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Transporter Rate").Value))) Then
                                Continue For
                            Else
                                Arr.Transporter_Rate = clsCommon.myCDecimal(grow.Cells("Transporter Rate").Value)
                            End If
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Security Rate").Value))) Then
                                Continue For
                            Else
                                Arr.Security_Rate = clsCommon.myCDecimal(grow.Cells("Security Rate").Value)
                            End If
                            If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("Booth Security Rate").Value))) Then
                                Continue For
                            Else
                                Arr.Booth_Security_Rate = clsCommon.myCDecimal(grow.Cells("Booth Security Rate").Value)
                            End If
                            obj.Add(Arr)
                        Next
                        clsCommon.ProgressBarHide()
                        If clsCommon.MyMessageBoxShow(Me, "Total Correct Document [" + clsCommon.myCstr(obj.Count) + "] out of [" + clsCommon.myCstr(linno) + "] Are You Sure.", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                            Dim sl As Integer = 1
                            'AddNew()
                            LoadBlankGrid()
                            If obj IsNot Nothing AndAlso obj.Count > 0 Then
                                isInsideLoadData = True
                                For Each objTr As clsDistributorCommissionDetails In obj
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(ColSNo).Value = sl
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteCode).Value = objTr.Route_Code
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteName).Value = clsDBFuncationality.getSingleValue("select Route_Desc from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(objTr.Route_Code) + "' ")
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(colCustCode).Value = objTr.Distributor_Code
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(colCustName).Value = clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_Customer_Master where cust_code='" + objTr.Distributor_Code + "'")
                                    GV1.Rows(GV1.Rows.Count - 1).Cells(colCRate).Value = objTr.Rate
                                    If chkSecurity.Checked Then
                                        GV1.Rows(GV1.Rows.Count - 1).Cells(colSecRate).Value = objTr.Security_Rate
                                        GV1.Rows(GV1.Rows.Count - 1).Cells(colBoothSecRate).Value = objTr.Booth_Security_Rate
                                    End If
                                    sl += 1
                                    GV1.Rows.AddNew()
                                Next
                                isInsideLoadData = False
                            End If
                            common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                        Else
                            common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Failed", Me.Text, MessageBoxButtons.OK)
                        End If
                        clsCommon.ProgressBarHide()
                    Catch ex As Exception
                        clsCommon.ProgressBarHide()
                        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                    End Try
                Else
                    clsCommon.MyMessageBoxShow(Me, "Excel Sheet is not in expected format", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Please Select Commission UOM", Me.Text)
            End If
            'clsCommon.ProgressBarHide()
            Me.Controls.Remove(gv)
        Catch ex As Exception
            'clsCommon.ProgressBarHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub Export()
        Try
            Dim str As String = "select Route_Code as [Route Code],Distributor_Code as [Distributor Code],'' as [Is Transpotation],Rate as [Rate],Transporter_Rate as [Transporter Rate],Security_Rate as [Security Rate],Booth_Security_Rate As [Booth Security Rate] from TSPL_Distributor_Commission_Detail"
            Dim whrCls As String = ""
            ListImpExpColumnsMandatory = New List(Of String)({"Route Code", "Distributor Code", "Is Transpotation", "Rate", "Transporter Rate", "Security Rate", "Booth Security Rate"})
            transportSql.ExporttoExcel(str, whrCls, Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub RefeshSNO()
        For ii As Integer = 1 To GV1.Rows.Count
            GV1.Rows(ii - 1).Cells(ColSNo).Value = ii
        Next
    End Sub
    Private Sub gv1_UserDeletedRow(sender As Object, e As GridViewRowEventArgs) Handles GV1.UserDeletedRow
        Try
            RefeshSNO()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_CurrentColumnChanged(sender As Object, e As CurrentColumnChangedEventArgs) Handles GV1.CurrentColumnChanged
        If GV1.RowCount > 0 Then
            Dim intCurrRow As Integer = GV1.CurrentRow.Index
            GV1.CurrentRow.Cells(ColSNo).Value = clsCommon.myCDecimal(intCurrRow + 1)
            If intCurrRow = GV1.Rows.Count - 1 Then
                GV1.Rows.AddNew()
                GV1.CurrentRow = GV1.Rows(intCurrRow)
            End If
        End If
    End Sub
    Private Sub gv1_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles GV1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                End If
                isCellValueChangedOpen = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub LoadData(ByVal strDocNo As String, ByVal NavTyep As NavigatorType)
        Try
            isInsideLoadData = True
            Dim obj As New clsDistributorCommission()
            obj = clsDistributorCommission.GetData(strDocNo, NavTyep, Nothing)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Doc_No) > 0) Then
                isNewEntry = False
                LoadBlankGrid()
                AddNew()
                If obj.IsPosted = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    lblStatus.Status = ERPTransactionStatus.Approved
                    btnDelete.Enabled = False
                Else
                    btnSave.Enabled = True
                    btnPost.Enabled = True
                    btnSave.Text = "Update"
                    btnDelete.Enabled = True
                    lblStatus.Status = ERPTransactionStatus.Pending
                End If
                btnGo.Enabled = False
                txtDocNo.Value = obj.Doc_No
                txtDate.Value = obj.Document_Date
                txtApplicableDate.Value = obj.Applicable_Date
                txtInActiveDate.Value = obj.InActive_date
                txtUOM.Value = obj.Commision_UOM
                txtDistributorTagging.Value = obj.Distributor_Tagging_Code
                cmbItemType.SelectedValue = obj.Item_Type
                cmbVehicleType.SelectedValue = obj.Vehicle_Type
                If obj.IS_Transpotation Then
                    rbtnTranspotation.IsChecked = True
                Else
                    rbtnCommission.IsChecked = True
                End If
                If obj.IS_Security Then
                    chkSecurity.Checked = True
                End If
                If obj.IN_Active Then
                    chkInActive.Checked = True
                End If
                txtItems.arrValueMember = obj.Items
                Dim sl As Integer = 1
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsDistributorCommissionDetails In obj.Arr
                        GV1.Rows(GV1.Rows.Count - 1).Cells(ColSNo).Value = sl
                        GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteCode).Value = objTr.Route_Code
                        GV1.Rows(GV1.Rows.Count - 1).Cells(ColRouteName).Value = clsDBFuncationality.getSingleValue("select Route_Desc from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(objTr.Route_Code) + "' ")
                        GV1.Rows(GV1.Rows.Count - 1).Cells(colCustCode).Value = objTr.Distributor_Code
                        GV1.Rows(GV1.Rows.Count - 1).Cells(colCustName).Value = clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_Customer_Master where cust_code='" + objTr.Distributor_Code + "'")
                        GV1.Rows(GV1.Rows.Count - 1).Cells(colCRate).Value = objTr.Rate
                        GV1.Rows(GV1.Rows.Count - 1).Cells(colTRate).Value = objTr.Transporter_Rate
                        If chkSecurity.Checked Then
                            GV1.Rows(GV1.Rows.Count - 1).Cells(colSecRate).Value = objTr.Security_Rate
                            GV1.Rows(GV1.Rows.Count - 1).Cells(colBoothSecRate).Value = objTr.Booth_Security_Rate
                        End If
                        sl += 1
                        GV1.Rows.AddNew()
                    Next
                    'GV1.Rows.AddNew()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub
    Private Function AllowToSave() As Boolean
        If clsCommon.myLen(txtUOM.Value) <= 0 Then
            txtUOM.Focus()
            clsCommon.MyMessageBoxShow(Me, "Please select UOM", Me.Text)
            Return False
        End If
        If clsCommon.myLen(txtDistributorTagging.Value) <= 0 Then
            txtUOM.Focus()
            clsCommon.MyMessageBoxShow(Me, "Please select Distributor Tagging", Me.Text)
            Return False
        End If
        Return True
    End Function
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveData()
    End Sub
    Public Sub SaveData()
        Try
            If (AllowToSave()) Then
                Dim obj As New clsDistributorCommission()
                obj.Doc_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                If txtItems.arrValueMember Is Nothing Then
                    Dim strItems As String = "select Item_code  from TSPL_ITEM_MASTER where  "
                    If clsCommon.CompairString(cmbItemType.SelectedValue, "M") = CompairStringResult.Equal Then
                        strItems += " TypeOfItm='M' "
                    ElseIf clsCommon.CompairString(cmbItemType.SelectedValue, "P") = CompairStringResult.Equal Then
                        strItems += " TypeOfItm='P' "
                    ElseIf clsCommon.CompairString(cmbItemType.SelectedValue, "I") = CompairStringResult.Equal Then
                        strItems += " TypeOfItm='I' "
                    Else
                        strItems += " TypeOfItm='O' "
                    End If
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(strItems)
                    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                        obj.Items = New ArrayList()
                        For Each dr As DataRow In dt.Rows
                            obj.Items.Add(clsCommon.myCstr(dr("Item_Code")))
                        Next
                    End If
                Else
                    obj.Items = txtItems.arrValueMember
                End If
                If chkInActive.Checked Then
                    obj.IN_Active = True
                End If
                If rbtnTranspotation.IsChecked Then
                    obj.IS_Transpotation = True
                ElseIf rbtnCommission.IsChecked Then
                    obj.IS_Transpotation = False
                End If
                If chkSecurity.Checked Then
                    obj.IS_Security = True
                End If
                obj.Applicable_Date = txtApplicableDate.Value
                obj.InActive_date = txtInActiveDate.Value
                obj.Commision_UOM = txtUOM.Value
                obj.Distributor_Tagging_Code = txtDistributorTagging.Value
                obj.Item_Type = cmbItemType.SelectedValue
                If EnableVehicleType Then
                    obj.Vehicle_Type = cmbVehicleType.SelectedValue

                End If
                obj.Arr = GetTRData()
                obj.SaveData(obj, isNewEntry)
                clsCommon.MyMessageBoxShow(Me, "Data saved successfully", Me.Text)
                LoadData(obj.Doc_No, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Function GetTRData() As List(Of clsDistributorCommissionDetails)
        Dim Arr As New List(Of clsDistributorCommissionDetails)
        For ii As Integer = 0 To GV1.RowCount - 1
            If clsCommon.myLen(GV1.Rows(ii).Cells(colCustCode).Value) > 0 Then
                'If clsCommon.myCDecimal(GV1.Rows(ii).Cells(colCRate).Value) > 0 Then
                Dim objTr As New clsDistributorCommissionDetails()
                'objTr.SNo = ii + 1
                objTr.Distributor_Code = clsCommon.myCstr(GV1.Rows(ii).Cells(colCustCode).Value)
                objTr.Route_Code = clsCommon.myCstr(GV1.Rows(ii).Cells(ColRouteCode).Value)
                objTr.Rate = clsCommon.myCDecimal(GV1.Rows(ii).Cells(colCRate).Value)
                objTr.Transporter_Rate = clsCommon.myCDecimal(GV1.Rows(ii).Cells(colTRate).Value)
                If chkSecurity.Checked Then
                    objTr.Security_Rate = clsCommon.myCDecimal(GV1.Rows(ii).Cells(colSecRate).Value)
                    objTr.Booth_Security_Rate = clsCommon.myCDecimal(GV1.Rows(ii).Cells(colBoothSecRate).Value)
                Else
                    objTr.Security_Rate = 0
                    objTr.Booth_Security_Rate = 0
                End If

                Arr.Add(objTr)
                'End If
            End If
        Next
        Return Arr
    End Function
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub
    Public Sub PostData()
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("No document found to post")
            End If
            If clsCommon.MyMessageBoxShow(Me, "Post the Current Document [" + txtDocNo.Value + "]" + Environment.NewLine + "Are You Sure.", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                clsDistributorCommission.PostData(clsCommon.myCstr(txtDocNo.Value))
                clsCommon.MyMessageBoxShow(Me, "Data posted successfully", Me.Text)
                LoadData(clsCommon.myCstr(txtDocNo.Value), NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDocNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDocNo._MYValidating
        Try
            If txtDocNo.MyReadOnly OrElse isButtonClicked Then
                Dim whrClas As String = "2=2"
                Dim qry As String = "select Doc_No as Code,Document_Date,Applicable_Date,Commision_UOM,Created_By,Created_Date,Modified_By,Modified_Date,IsPosted,Posted_Date from TSPL_Distributor_Commission_Head"
                LoadData(clsCommon.myCstr(clsCommon.ShowSelectForm("RPTBDMST", qry, "Code", whrClas, txtDocNo.Value, "Code", isButtonClicked, "Document_Date")), NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDocNo__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim qry As String = "select count(*) from TSPL_Distributor_Commission_Head where Doc_No='" + txtDocNo.Value + "'"
            Dim count As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry))
            If count = 0 Then
                txtDocNo.MyReadOnly = False
            Else
                txtDocNo.MyReadOnly = True
            End If
            LoadData(clsCommon.myCstr(txtDocNo.Value), NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub txtDistributorTagging__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDistributorTagging._MYValidating
        Try
            Dim qry As String = "select  x.Code as Code,x.Start_Date,x.Remarks,x.NoOfRoute as [No of Route],x.NoOfDistributor as [No. of Distributor] 
from(
select TSPL_DISTRIBUTOR_ROUTE.Code as Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,count(distinct TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No) as NoOfRoute,count(distinct TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code) as NoOfDistributor
from TSPL_DISTRIBUTOR_ROUTE
left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
where 
-- not exists (select 1 from TSPL_DISTRIBUTOR_COMMISSION_HEAD where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Distributor_Tagging_Code =TSPL_DISTRIBUTOR_ROUTE.Code) and 
TSPL_DISTRIBUTOR_ROUTE.Status=1"
            If rbtnCommission.IsChecked Then
                qry += " and IS_Transpoter=0 "
            ElseIf rbtnTranspotation.IsChecked Then
                qry += " and IS_Transpoter=1 "
            End If
            qry += " Group by TSPL_DISTRIBUTOR_ROUTE.Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks
) X"
            txtDistributorTagging.Value = clsCommon.ShowSelectForm("DistributorTaggingFinder", qry, "Code", "", txtDistributorTagging.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                Dim Qry As String = Nothing
                Qry = "select ROW_NUMBER() OVER(ORDER BY (Select 1) ASC) AS [S.No.],
                        TSPL_Distributor_Commission_Head.Doc_No,TSPL_Distributor_Commission_Detail.Route_Code,TSPL_Distributor_Commission_Head.Document_Date,
                        TSPL_Distributor_Commission_Head.Applicable_Date,TSPL_Distributor_Commission_Head.Commision_UOM,TSPL_Distributor_Commission_Detail.Rate,(TSPL_Customer_Master.Cust_Code+' - '+ TSPL_Customer_Master.Customer_Name) As Distributor,
                        TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2
                        from  TSPL_Distributor_Commission_Detail 
                        Left outer Join TSPL_Distributor_Commission_Head On TSPL_Distributor_Commission_Head.Doc_No=TSPL_Distributor_Commission_Detail.Doc_No
                        Left Outer Join TSPL_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_Distributor_Commission_Detail.Distributor_Code
                        Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code=TSPL_Customer_Master.Comp_Code
                        where 2=2 and TSPL_Distributor_Commission_Head.Doc_No = '" + txtDocNo.Value + "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptDistributorCommission", "Distributor Commission Rate", Nothing)
                    frmCRV = Nothing
                End If
            Else
                clsCommon.MyMessageBoxShow("No data found", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnCommission_CheckStateChanged(sender As Object, e As EventArgs) Handles rbtnCommission.CheckStateChanged
        If rbtnCommission.IsChecked Then
            rbtnTranspotation.IsChecked = False
        End If
    End Sub
    Private Sub rbtnTranspotation_CheckStateChanged(sender As Object, e As EventArgs) Handles rbtnTranspotation.CheckStateChanged
        If rbtnTranspotation.IsChecked Then
            rbtnCommission.IsChecked = False
        End If
    End Sub
    Private Sub chkSecurity_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkSecurity.ToggleStateChanged
        If chkSecurity.Checked Then
            rbtnCommission.Enabled = False
            rbtnTranspotation.Enabled = False
            LoadBlankGrid()
        Else
            rbtnCommission.Enabled = True
            rbtnTranspotation.Enabled = True
            LoadBlankGrid()
        End If
    End Sub
    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                If clsCommon.MyMessageBoxShow("Unpost the current transaction" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    clsDistributorCommission.ReverseAndUnpost(txtDocNo.Value)
                    clsCommon.MyMessageBoxShow(Me, "Tansaction unposted succesffuly", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub DeleteData()
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "You Cannot Delete Record", Me.Text)
            Exit Sub
        End If
        funDelete()
    End Sub
    Sub funDelete()
        Try
            Dim Reason As String = ""
            If (myMessages.deleteConfirm()) Then
                If (clsDistributorCommission.DeleteData(txtDocNo.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        DeleteData()
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocNo.Value, "doc_no", "TSPL_Distributor_Commission_Head", "TSPL_Distributor_Commission_Detail")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
End Class