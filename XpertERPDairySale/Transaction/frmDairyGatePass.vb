' '' '' '' ''Created by Sanjeet 31/01/2018 ========
Imports common
Imports System.Data.SqlClient
Imports System
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports System.IO
Public Class frmDairyGatePass
    Inherits FrmMainTranScreen
    'Sanjay Ticket No-MIL/08/07/19-000104 Add Salesman
#Region "Variables"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim strQuery As String
    Dim SetDefaultShiftTime As String = ""
    Dim VehicleNofromDispatch As Boolean = False
    Dim EnableProductSaleForJPR As Boolean = False
    Dim AllowManualCrateForDispatch As Boolean = False
    Dim ApplyDepartmentRoute As Boolean = False
    Dim isDepartmentRoute As Boolean = False
    Dim ServerDateTimeForTaxableInvoice As Boolean = False
    Dim strQueryCANCRate As String
    Dim dt As DataTable
    Private isNewEntry As Boolean = False
    Dim strVehicleNo As List(Of String)
    Private AllowGatePassDemandTripWise As Boolean = False
    Private GatepassForTaxableandNonTaxableItems As Boolean = False
    Const ColApply As String = "ColApply"
    Const ColDocNo As String = "ColDocNo"
    Const ColDocDate As String = "ColDocDatet"
    Const ColToSalesmanCode As String = "ColToSalesmanCode"
    Const ColToSalesmanname As String = "ColToSalesmanname"
    Const ColRoute_No As String = "ColRoute_No"
    Const ColRoute_Desc As String = "ColRoute_Desc"
    Const ColType As String = "ColType"
    Const ColPriceCode As String = "ColPriceCode"
    Const ColPriceDesc As String = "ColPriceDesc"
    Const ColCustCode As String = "ColCustCode"
    Const ColCustName As String = "ColCustName"
    Dim isInsideLoadData As Boolean = False
    Dim blnLoad As Boolean = False
    Const colItemCode As String = "colItemCode"
    Const colItemDesc As String = "colItemDesc"
    Const colUnit As String = "colUnit"
    Const colQty As String = "colQty"
    Const colCrateIssue As String = "colCrateIssue"
    Const colLineNo As String = "colLineNo"
    Const colPKID As String = "colPKID"
    Const colHSNCode As String = "colHSNCode"
    Const colSchemeItem As String = "colSchemeItem"

    Const colCTDocCode As String = "colCTDocCode"
    Const colCTICode As String = "colCTICode"
    Const colCTCode As String = "colCTCode"
    Const ColCTName As String = "ColCTName"
    Const ColCTQty As String = "ColCTQty"

    Dim atchqry As String = ""
    Dim AlternateVechileforGatePass As Double
    Dim isCreateProvisionOfTransporterInDairyDispatch As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim isCTQtyUpdate As Boolean = False
    Dim SettCreateProvisionOnOpeningAndClosingKM As Boolean = False
    Dim IsLoadingSlipMandatory As Boolean = False
    Dim CreateAutoGatePass As Boolean = False
    Dim CreateGatePassFromDemand As Boolean = False
    Public arrShipmentFromMultiple As ArrayList
    Public Property routeno As String

    Public Property CreditCustomer As String
    Public Property txtlocation As String
    Public Property vehicleno As String
    Public Property vehicle_desc As String
    Public Property DriverName As String
    Public Property docdate As Date?
    Public Property Supplydate As Date?
    Public Property Shifttype As String = Nothing
    Public Property GenerateCustomerWiseGatePass As Boolean = False
    Public Property ShipmentDocNo As String
    ''ERO/03/05/19-000584 by balwindr on 06/05/2019
    ''ERO/03/05/19-000584 by balwindr on 06/05/2019
    Dim VehicleDesc As String = Nothing
    Dim OneTimeCheck As Boolean = False
    Dim EnableDispatch As Boolean = False
    Dim EnableLocation As Boolean = False
    Dim settFileUpload As Boolean = False
    Dim DifferentCrateTypeForFGItem As Boolean = False
    Dim ApplySaleRouteInDCSSale As Boolean = False
#End Region
    Private Sub SetUserMgmtNew()
        Me.Form_ID = clsUserMgtCode.frmDairyGatePass
        MyBase.SetUserMgmt(clsUserMgtCode.frmDairyGatePass)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPrint.Visible = MyBase.isPrintFlag
        RadSplitButton1.Visible = MyBase.isPrintFlag
        btnClKM.Visible = MyBase.isModifyFlag
        btnPrint2.Visible = MyBase.isPrintFlag
        btnGPCancel.Visible = MyBase.isCancel_Flag
        btnReverse.Visible = False

        'If MyBase.isReverse Then
        '    btnReverse.Enabled = True
        'Else
        '    btnReverse.Enabled = False
        'End If
    End Sub
    Private Sub FrmGatePassENtry1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ApplySaleRouteInDCSSale = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ApplySaleRouteInDCSSale, clsFixedParameterCode.ApplySaleRouteInDCSSale, Nothing)) = "1", True, False))
            settFileUpload = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.FileUpload, clsUserMgtCode.frmDairyGatePass, Nothing)) = 1)
            'CreateTable()
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                chkGhee.Visible = True
            Else
                chkGhee.Visible = False
            End If
            'Dim coll As Dictionary(Of String, String)
            'coll = New Dictionary(Of String, String)()
            'coll.Add("Closing_Date", "Datetime null")
            'coll.Add("ShiftType", "varchar(20) NULL")
            'clsCommonFunctionality.CreateOrAlterTable("TSPL_DAIRYSALE_GATEPASS_MASTER", coll)
            'coll = New Dictionary(Of String, String)()
            'coll.Add("Production_Remarks", "varchar(200) NULL")
            'coll.Add("GPCode", "varchar(30) NULL")
            'clsCommonFunctionality.CreateOrAlterTable("TSPL_DEMAND_BOOKING_DETAIL", coll)
            SetUserMgmtNew()
            isNewEntry = True

            txtDate.Value = clsCommon.GETSERVERDATE()
            txtSupplyDate.Value = txtDate.Value
            txtGatepassDate.Value = clsCommon.GETSERVERDATE()
            'CheckCreateCapacity = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CrateCapacityCheck, clsFixedParameterCode.CrateCapacityCheck, Nothing)))
            isCreateProvisionOfTransporterInDairyDispatch = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateProvisionOfTransporterInDairyDispatch, clsFixedParameterCode.CreateProvisionOfTransporterInDairyDispatch, Nothing)))
            IsLoadingSlipMandatory = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.IsLoadingSlipMandatory, clsFixedParameterCode.IsLoadingSlipMandatory, Nothing)) = 1, True, False)
            AllowGatePassDemandTripWise = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowGatePassDemandTripWise, clsFixedParameterCode.AllowGatePassDemandTripWise, Nothing)) = 1, True, False)
            GatepassForTaxableandNonTaxableItems = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GatepassForTaxableandNonTaxableItems, clsFixedParameterCode.GatepassForTaxableandNonTaxableItems, Nothing)) = 1, True, False)
            EnableLocation = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableLocation, clsFixedParameterCode.EnableLocation, Nothing)) = 1, True, False)
            SettCreateProvisionOnOpeningAndClosingKM = (clsCommon.myCdbl(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateProvisionOnOpeningAndClosingKM, clsFixedParameterCode.CreateProvisionOnOpeningAndClosingKM, Nothing))) = 1)
            CreateGatePassFromDemand = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateGatePassFromDemand, clsFixedParameterCode.CreateGatePassFromDemand, Nothing)))
            VehicleNofromDispatch = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.VehicleNofromDispatch, clsFixedParameterCode.VehicleNofromDispatch, Nothing)))
            EnableProductSaleForJPR = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)))
            AllowManualCrateForDispatch = clsCommon.myCBool(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowManualCrateForDispatch, clsFixedParameterCode.AllowManualCrateForDispatch, Nothing)))
            SetDefaultShiftTime = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SetDefaultShiftTime, clsFixedParameterCode.SetDefaultShiftTime, Nothing))
            ApplyDepartmentRoute = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyDepartmentRoute, clsFixedParameterCode.ApplyDepartmentRoute, Nothing)) = 1, True, False)
            DifferentCrateTypeForFGItem = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DifferentCrateTypeForFGItem, clsFixedParameterCode.DifferentCrateTypeForFGItem, Nothing)) = 1, True, False)
            CreateAutoGatePass = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateAutoGatePass, clsFixedParameterCode.CreateAutoGatePass, Nothing)) = 1, True, False)
            ServerDateTimeForTaxableInvoice = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ServerDateTimeForTaxableInvoice, clsFixedParameterCode.ServerDateTimeForTaxableInvoice, Nothing)) = 1, True, False)

            Addnew()
            LoadBlankGrid()
            Panel2.Visible = SettCreateProvisionOnOpeningAndClosingKM
            cmbitemtype.Text = "Select"
            txtTransporter.MaxLength = 100
            txtSalesman.MaxLength = 100
            txtRemarks.MaxLength = 200
            txtComments.MaxLength = 200
            If isCreateProvisionOfTransporterInDairyDispatch = True Then
                btnPost.Visible = True
                btnPost.Enabled = False
            Else
                ' btnPost.Visible = False
            End If
            If Not DifferentCrateTypeForFGItem Then
                RadPageView1.Pages("rpvpCrateType").Item.Visibility = ElementVisibility.Collapsed
            End If
            AlternateVechileforGatePass = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowAlternateVechileforFreshSale, clsFixedParameterCode.ShowAlternateVechileforFreshSale, Nothing))
            'funFillGrid()
            txtTransporter.Enabled = False
            If isCreateProvisionOfTransporterInDairyDispatch = True Then
                txtTransporter.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'"))
            End If
            If clsCommon.myLen(Me.Tag) > 0 Then
                LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
            End If
            txtmultiBooking.Enabled = False
            If IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TollTaxMaster, clsFixedParameterCode.TollTaxMaster, Nothing)) = 0, False, True) = False Then
                lblTollAmount.Visible = True
                txtTollAmount.Visible = True
            Else
                lblTollAmount.Visible = False
                txtTollAmount.Visible = False
            End If
            multipleDelteVisible(False)
            If CreateGatePassFromDemand = True Then
                RadGroupBox3.Visible = True
                txtLocCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
                If clsCommon.myLen(txtLocCode.Value) > 0 Then
                    txtLocDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocCode.Value + "'"))
                End If
                txtLocCode.Enabled = False
                txtmultiBooking.Visible = False
            Else
                RadGroupBox3.Visible = True
            End If
            fndRouteNo.Value = routeno
            txtRouteName.Text = clsDBFuncationality.getSingleValue("select Route_Desc from TSPL_ROUTE_MASTER where Route_No='" & fndRouteNo.Value & "'")
            txtLocCode.Value = txtlocation
            txtLocDesc.Text = clsDBFuncationality.getSingleValue("select  Location_Desc  from TSPL_LOCATION_MASTER where Location_Code='" & txtLocCode.Value & "'")
            txtVehicle.Value = vehicleno
            lblVehicleDesc.Text = vehicle_desc 'clsDBFuncationality.getSingleValue("select Description from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'")
            txtDriverName.Text = DriverName
            If docdate IsNot Nothing AndAlso clsCommon.myLen(docdate) > 0 Then
                txtDate.Value = docdate
            End If
            If Supplydate IsNot Nothing AndAlso clsCommon.myLen(Supplydate) > 0 Then
                txtSupplyDate.Value = Supplydate
            End If

            If clsCommon.CompairString(Shifttype, "AM") = CompairStringResult.Equal Then
                rbtnMorning.IsChecked = True
                txtTripNo.Text = "1"
                funFillGrid()
                If CreateAutoGatePass Then
                    txtLoadingSlip.Text = "1"
                    btnSave_Click(btnSave, New EventArgs())
                    btnPrint2_Click(btnPrint2, New EventArgs())
                    Me.Close()
                End If
            ElseIf clsCommon.CompairString(Shifttype, "PM") = CompairStringResult.Equal Then
                rbtnEvening.IsChecked = True
                txtTripNo.Text = "1"
                funFillGrid()
                If CreateAutoGatePass Then
                    txtLoadingSlip.Text = "1"
                    btnSave_Click(btnSave, New EventArgs())
                    btnPrint2_Click(btnPrint2, New EventArgs())
                    Me.Close()
                End If
            End If
            If GatepassForTaxableandNonTaxableItems Then
                RadGroupBox1.Visible = True
            End If
            '  LoadData(txtCode.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub CheckTimeShift()
        Try
            Dim inputDate As DateTime
            If DateTime.TryParse(txtDate.Value.ToString(), inputDate) Then
                If inputDate.TimeOfDay < New TimeSpan(12, 0, 0) Then
                    rbtnMorning.IsChecked = True
                    rbtnEvening.IsChecked = False
                Else
                    rbtnEvening.IsChecked = True
                    rbtnMorning.IsChecked = False
                End If
            Else
                Throw New Exception("Invalid date format")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub CreateTable()
        '    Dim coll As Dictionary(Of String, String)
        '    coll = New Dictionary(Of String, String)()
        '    coll.Add("GPCode", "Varchar(30) not null  PRIMARY KEY")
        '    coll.Add("GPDate", "datetime not NULL")
        '    coll.Add("Vehicle_Id", "varchar(12) NOT NULL")
        '    coll.Add("Vehicle_Number", "varchar(50) NOT NULL")
        '    coll.Add("Item_Type", "varchar(10) NULL")
        '    coll.Add("Comp_Code", "VARCHAR(8)")
        '    coll.Add("Created_By", "varchar(12) NOT NULL")
        '    coll.Add("Created_Date", "Datetime NOT NULL")
        '    coll.Add("Modified_By", "varchar(12) NOT NULL")
        '    coll.Add("Modified_Date", "Datetime NOT NULL")
        '    coll.Add("Transporter", "varchar(100) NULL")
        '    coll.Add("Remarks", "varchar(200) NULL")
        '    coll.Add("Comments", "varchar(200) NULL")
        '    coll.Add("Post", "char(1)  Not NUll Default 'N'")
        '    coll.Add("Location_Code", "varchar(12)  NULL")
        '    coll.Add("Location_Desc", "varchar(50)  NULL")
        '    coll.Add("Route_No", "varchar(12)  NULL")
        '    coll.Add("TotalCAN", "float  null")
        '    coll.Add("TotalCrate", "float  null")
        '    coll.Add("Opening_Km", "Decimal(18,2) null")
        '    coll.Add("Closing_Km", "Decimal(18,2) null")
        '    coll.Add("Distance_In_Route", "Decimal(18,2) null")
        '    coll.Add("Price_KM_In_Vehicle", "Decimal(18,2) null")
        '    coll.Add("Toll_Amount", "Decimal(18,2) null")
        '    coll.Add("Salesman", "varchar(100) NULL")
        '    coll.Add("Closing_Date", "Datetime null")
        '    coll.Add("IsTransfer", "int not null default 0")
        '    coll.Add("AgainstTransferNo", "Varchar(30) null References TSPL_TRANSFER_ORDER_HEAD(Document_No)")
        '    coll.Add("ShiftType", "varchar(20) NULL")
        '    coll.Add("Loading_Slip", "varchar(20) NULL")
        '    coll.Add("GatePass_Date", "datetime NULL")
        '    coll.Add("Status", "char(1)  NUll")
        '    coll.Add("Driver_Name", "varchar(100) NULL")
        '    coll.Add("Driver_ContactNo", "varchar(15) NULL")
        '    clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_DAIRYSALE_GATEPASS_MASTER", coll, Nothing, True, False, "", "GPCode", "GPDate")
        AlterShipmentTable()
        CreateBalancingTable()
    End Sub
    Private Sub AlterShipmentTable()
        Dim coll As Dictionary(Of String, String)
        coll = New Dictionary(Of String, String)()
        coll.Add("PK_ID", "integer NOT NULL identity NOT FOR REPLICATION primary key")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_SD_SHIPMENT_DETAIL", coll, Nothing, True, True, "TSPL_SD_SHIPMENT_HEAD", "DOCUMENT_CODE", "")
    End Sub
    Private Sub CreateBalancingTable()
        Dim coll As Dictionary(Of String, String)
        coll = New Dictionary(Of String, String)()
        coll.Add("PK_ID", "integer NOT NULL REFERENCES TSPL_SD_SHIPMENT_DETAIL(PK_ID)")
        coll.Add("GPCode", "Varchar(30) NOT NULL REFERENCES TSPL_DAIRYSALE_GATEPASS_MASTER(GPCode)")
        coll.Add("Item_Code", "Varchar(50) NULL")
        coll.Add("Unit_Code", "Varchar(12) NULL")
        coll.Add("GP_Qty", "Decimal(18,2) NULL")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL", coll, Nothing, True, False, "TSPL_DAIRYSALE_GATEPASS_MASTER", "GPCode", "")
        '=====================
    End Sub
    Sub LoadgvCrateType()
        gvCrateType.Rows.Clear()
        gvCrateType.Columns.Clear()
        Dim CrateTypeDocCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        CrateTypeDocCode.FormatString = ""
        CrateTypeDocCode.HeaderText = "Document Code"
        CrateTypeDocCode.Name = colCTDocCode
        CrateTypeDocCode.HeaderImage = My.Resources.search4
        CrateTypeDocCode.TextImageRelation = TextImageRelation.TextBeforeImage
        CrateTypeDocCode.Width = 100
        CrateTypeDocCode.ReadOnly = True
        CrateTypeDocCode.IsVisible = False
        gvCrateType.MasterTemplate.Columns.Add(CrateTypeDocCode)
        Dim resCrateICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        resCrateICode.FormatString = ""
        resCrateICode.HeaderText = "Item Code"
        resCrateICode.Name = colCTICode
        resCrateICode.Width = 80
        resCrateICode.ReadOnly = True
        resCrateICode.IsVisible = False
        gvCrateType.MasterTemplate.Columns.Add(resCrateICode)
        Dim resCratetype As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        resCratetype.FormatString = ""
        resCratetype.HeaderText = "Crate Type Code"
        resCratetype.Name = colCTCode
        resCratetype.Width = 80
        resCratetype.ReadOnly = True
        resCratetype.IsVisible = False
        gvCrateType.MasterTemplate.Columns.Add(resCratetype)
        Dim resCTname As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        resCTname.FormatString = ""
        resCTname.HeaderText = "Crate Type Name"
        resCTname.Name = ColCTName
        resCTname.Width = 200
        resCTname.ReadOnly = True
        resCTname.IsVisible = True
        gvCrateType.MasterTemplate.Columns.Add(resCTname)

        Dim repoCrateQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCrateQty.FormatString = ""
        repoCrateQty.HeaderText = "Crate Qty"
        repoCrateQty.Name = ColCTQty
        repoCrateQty.Width = 120
        repoCrateQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCrateQty.ShowUpDownButtons = False
        repoCrateQty.ReadOnly = False
        repoCrateQty.IsVisible = True
        gvCrateType.MasterTemplate.Columns.Add(repoCrateQty)

        gvCrateType.AllowAddNewRow = False
        gvCrateType.ShowGroupPanel = False
        gvCrateType.AllowColumnReorder = True
        gvCrateType.AllowRowReorder = False
        gvCrateType.EnableSorting = False
        gvCrateType.Rows.AddNew()
        gvCrateType.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvCrateType.MasterTemplate.ShowRowHeaderColumn = False
        gvCrateType.TableElement.TableHeaderHeight = 40
    End Sub
    Private Sub LoadBlankGrid()
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        Gv1.AllowDeleteRow = True
        Gv1.AllowAddNewRow = False
        Dim repoPK_ID As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPK_ID = New GridViewDecimalColumn()
        repoPK_ID.FormatString = ""
        repoPK_ID.HeaderText = "ID"
        repoPK_ID.Name = colPKID
        repoPK_ID.Width = 50
        repoPK_ID.ReadOnly = True
        repoPK_ID.IsVisible = False
        repoPK_ID.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Gv1.MasterTemplate.Columns.Add(repoPK_ID)
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Gv1.MasterTemplate.Columns.Add(repoLineNo)
        Dim ItemCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ItemCode.FormatString = ""
        ItemCode.HeaderText = "Item Code"
        ItemCode.Name = colItemCode
        ItemCode.Width = 100
        ItemCode.ReadOnly = True
        Gv1.MasterTemplate.Columns.Add(ItemCode)
        Dim ItemDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ItemDesc.FormatString = ""
        ItemDesc.HeaderText = "Item Desc"
        ItemDesc.Name = colItemDesc
        ItemDesc.Width = 100
        ItemDesc.ReadOnly = True
        Gv1.MasterTemplate.Columns.Add(ItemDesc)
        Dim CustCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        CustCode.FormatString = ""
        CustCode.HeaderText = "Cust Code"
        CustCode.Name = ColCustCode
        CustCode.Width = 100
        CustCode.ReadOnly = True
        CustCode.IsVisible = False
        Gv1.MasterTemplate.Columns.Add(CustCode)
        Dim CustName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        CustName.FormatString = ""
        CustName.HeaderText = "Cust Name"
        CustName.Name = ColCustName
        CustName.Width = 100
        CustName.ReadOnly = True
        CustName.IsVisible = False
        Gv1.MasterTemplate.Columns.Add(CustName)
        Dim Unit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        Unit.FormatString = ""
        Unit.HeaderText = "Unit"
        Unit.Name = colUnit
        Unit.Width = 100
        Unit.ReadOnly = True
        Gv1.MasterTemplate.Columns.Add(Unit)
        Dim HSN As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        HSN.FormatString = ""
        HSN.HeaderText = "HSN Code"
        HSN.Name = colHSNCode
        HSN.Width = 100
        HSN.ReadOnly = True
        Gv1.MasterTemplate.Columns.Add(HSN)

        Dim SchemeItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        SchemeItem.FormatString = ""
        SchemeItem.HeaderText = "Qty Scheme Item"
        SchemeItem.Name = colSchemeItem
        SchemeItem.Width = 100
        SchemeItem.ReadOnly = True
        Gv1.MasterTemplate.Columns.Add(SchemeItem)

        Dim repoQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQty = New GridViewDecimalColumn()
        repoQty.FormatString = ""
        repoQty.HeaderText = "Quantity"
        repoQty.Name = colQty
        repoQty.Width = 80
        repoQty.Minimum = 0
        If clsCommon.myCstr(txtCode.Value) IsNot Nothing AndAlso clsCommon.myLen(txtCode.Value) > 0 Then
            repoQty.ReadOnly = True
        Else
            repoQty.ReadOnly = False
        End If
        repoQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Gv1.MasterTemplate.Columns.Add(repoQty)
        Dim repoisCrateIssue As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoisCrateIssue = New GridViewDecimalColumn()
        repoisCrateIssue.FormatString = ""
        repoisCrateIssue.HeaderText = "Crate Qty"
        repoisCrateIssue.Name = colCrateIssue
        repoisCrateIssue.Width = 80
        repoisCrateIssue.Minimum = 0
        If clsCommon.myCstr(txtCode.Value) IsNot Nothing AndAlso clsCommon.myLen(txtCode.Value) > 0 Then
            repoisCrateIssue.ReadOnly = True
        Else
            repoisCrateIssue.ReadOnly = False
        End If
        repoisCrateIssue.IsVisible = True
        repoisCrateIssue.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Gv1.MasterTemplate.Columns.Add(repoisCrateIssue)
        Gv1.ShowGroupPanel = False
        Gv1.AllowColumnReorder = False
        Gv1.AllowRowReorder = False
        Gv1.EnableSorting = False
        Gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Gv1.AllowAddNewRow = False
    End Sub
    Private Function LoadQuery(ByVal strItemCode As String, ByVal EnableDispatch As Boolean) As String
        Try


            Dim strItem As String = String.Empty
            If CreateGatePassFromDemand Then
                If clsCommon.myLen(strItemCode) > 0 Then
                    strItem = " and TSPL_DEMAND_BOOKING_DETAIL.Item_Code='" & strItemCode & "'"
                End If
                strQuery = "select TSPL_DEMAND_BOOKING_MASTER.Document_No as [Document No],Document_Date as [Document Date],TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,Customer_Name, " &
                       "TSPL_DEMAND_BOOKING_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_DEMAND_BOOKING_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,'' as Scheme_Item   " &
                       "from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_no=TSPL_DEMAND_BOOKING_DETAIL.DOCUMENT_no " &
                       "left outer join TSPL_ITEM_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                       "left outer join TSPL_CUSTOMER_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                       "where convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='" & clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MMM/yyyy") & " ' And isnull(GPCode,'') = ''"
                If rbtn_Milk.IsChecked Then
                    strQuery += "and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'"

                End If
                strQuery += " and TSPL_DEMAND_BOOKING_MASTER.Location_Code='" & txtLocCode.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code='" + txtVehicle.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.Item_Code <> '' " & strItem & " "
                If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                    strQuery += "  and TSPL_DEMAND_BOOKING_MASTER.route_no='" & fndRouteNo.Value & "'"
                End If
            Else
                If clsCommon.myLen(strItemCode) > 0 Then
                    strItem = " and TSPL_SD_SHIPMENT_DETAIL.Item_Code='" & strItemCode & "'"
                End If
                If AlternateVechileforGatePass.Equals(1) Then
                    Dim StrChkAvQuery As String = "select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                    "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                    "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        StrChkAvQuery += "Left outer join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                    End If

                    StrChkAvQuery += " left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                    "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                    "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                    "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_HEAD.AlternateVehicle='" + txtVehicle.Value + "'   and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                    If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                        StrChkAvQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                    End If
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        If chkGhee.Checked Then
                            StrChkAvQuery += " and TSPL_BOOKING_MATSER.Is_GHEE = 1 "
                        Else
                            StrChkAvQuery += " and isnull(TSPL_BOOKING_MATSER.Is_GHEE,0) = 0 "
                        End If
                    End If

                    If ApplySaleRouteInDCSSale Then
                        StrChkAvQuery += " " & Environment.NewLine & " union all 
                   select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                   "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                   "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "

                        StrChkAvQuery += " left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                        "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                        "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                        "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                            StrChkAvQuery += "  and TSPL_SD_SHIPMENT_HEAD.Sale_Route_Code='" + fndRouteNo.Value + "'"
                        End If
                    End If

                    dt = New DataTable()
                    dt = clsDBFuncationality.GetDataTable(StrChkAvQuery)
                    If dt.Rows.Count > 0 Then
                        strQuery = "select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                        "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                        "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "

                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                            strQuery += "Left outer join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                        End If
                        strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code Left outer join TSPL_CUSTOMER_MASTER On TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                        "where convert(Date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                        "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and TSPL_SD_SHIPMENT_HEAD.AlternateVehicle='" + txtVehicle.Value + "' and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                        End If
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                            If chkGhee.Checked Then
                                strQuery += " and TSPL_BOOKING_MATSER.Is_GHEE = 1 "
                            Else
                                strQuery += " and isnull(TSPL_BOOKING_MATSER.Is_GHEE,0) = 0 "
                            End If
                        End If
                        If ApplySaleRouteInDCSSale Then
                            strQuery += " " & Environment.NewLine & " union all 
                        select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                        "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                        "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "
                            strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code Left outer join TSPL_CUSTOMER_MASTER On TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                            "where convert(Date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                            "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Sale_Route_Code='" + fndRouteNo.Value + "'"
                            End If
                        End If

                    Else
                        strQuery = "select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                       "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                       "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                            strQuery += "Left outer join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                        End If
                        strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                       "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                       "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                       "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and TSPL_SD_SHIPMENT_HEAD.Vehicle_Code='" + txtVehicle.Value + "' and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & " "
                        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                        End If
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                            If chkGhee.Checked Then
                                strQuery += " and TSPL_BOOKING_MATSER.Is_GHEE = 1 "
                            Else
                                strQuery += " and isnull(TSPL_BOOKING_MATSER.Is_GHEE,0) = 0 "
                            End If
                        End If


                        If ApplySaleRouteInDCSSale Then
                            strQuery += " " & Environment.NewLine & " union all 
                        select TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                       "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code,Scheme_Item  " &
                       "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "

                            strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                           "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                           "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                           "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & " "
                            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Sale_Route_Code='" + fndRouteNo.Value + "'"
                            End If
                        End If

                    End If
                Else
                    strQuery = "select TSPL_SD_SHIPMENT_DETAIL.PK_ID,TSPL_SD_SHIPMENT_HEAD.VehicleNo, TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                      "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit"
                    If AllowGatePassDemandTripWise Then
                        strQuery += " ,TSPL_SD_SHIPMENT_DETAIL.Qty as Qty, TSPL_SD_SHIPMENT_DETAIL.trip_no "
                    Else
                        strQuery += " ,Qty"
                    End If
                    strQuery += ",TSPL_ITEM_MASTER.HSN_Code,Scheme_Item,TSPL_SD_SHIPMENT_DETAIL.Crate as [Crate_Issue]  " &
                      "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        strQuery += "Left outer join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                    End If
                    strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                      "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  "
                    'If AllowGatePassDemandTripWise Then
                    '    strQuery += "   left outer  join  TSPL_SD_SHIPMENT_BOOKING_DETAIL on TSPL_SD_SHIPMENT_DETAIL.Document_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE And TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.item_code And TSPL_SD_SHIPMENT_DETAIL.Unit_code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Unit_code and TSPL_SD_SHIPMENT_DETAIL.Trip_No='" + txtTripNo.Text + "'"
                    'End If
                    strQuery += "where convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)='" & clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                      "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and (case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end)='" + txtVehicle.Value + "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                    If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                        strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" & fndRouteNo.Value & "'"
                    End If
                    If rbtnMorning.IsChecked Then
                        strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM'"
                    Else
                        strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM'"
                    End If
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                        If chkIndividualCustomer.Checked Then
                            strQuery += " and TSPL_SD_SHIPMENT_HEAD.IsIndividualCustomer=1 and TSPL_SD_SHIPMENT_HEAD.Demand_UniqueID in('" & txtDemandNo.Value & "-T','" & txtDemandNo.Value & "-NT' )"
                            txtTripNo.Text = "1"
                        Else
                            strQuery += " and TSPL_SD_SHIPMENT_HEAD.IsIndividualCustomer=0 "
                        End If
                    End If
                    If EnableDispatch Then
                        'strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Status=0"
                    Else
                        strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Status=1"
                        If GatepassForTaxableandNonTaxableItems Then
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                                If rdbTaxable.IsChecked Then
                                    strQuery &= " And TSPL_ITEM_MASTER.Is_CrateType=0  "
                                ElseIf rdbNonTaxable.IsChecked Then
                                    strQuery &= " And TSPL_ITEM_MASTER.Is_CrateType=1  "
                                End If
                            Else
                                If rdbTaxable.IsChecked Then
                                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Is_Taxable=1 "
                                ElseIf rdbNonTaxable.IsChecked Then
                                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Is_Taxable=0 "
                                End If
                            End If
                        End If
                    End If
                    If GenerateCustomerWiseGatePass Then
                        strQuery += " and TSPL_SD_SHIPMENT_HEAD.Document_Code='" & ShipmentDocNo & "' "
                    End If
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        If chkGhee.Checked Then
                            strQuery += " and TSPL_BOOKING_MATSER.Is_GHEE = 1 "
                        Else
                            strQuery += " and isnull(TSPL_BOOKING_MATSER.Is_GHEE,0) = 0 "
                        End If
                    End If
                    If EnableProductSaleForJPR Then
                        If rbtn_Milk.IsChecked Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='S'"
                        ElseIf rbtn_product.IsChecked Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='P'"
                        ElseIf rbtn_IceCream.IsChecked Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='I'"
                        End If
                    End If

                    If AllowGatePassDemandTripWise Then
                        If clsCommon.myLen(txtTripNo.Text) > 0 Then
                            strQuery += "  and TSPL_SD_SHIPMENT_DETAIL.Trip_No='" + txtTripNo.Text + "'"
                        Else
                            Throw New Exception("Please Enter Trip No. ")
                        End If

                    End If
                    strQueryCANCRate = "select sum(crate) as crate ,sum(ShippedCAN) as ShippedCAN from TSPL_SD_SHIPMENT_HEAD " &
                                "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                       "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                       "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and (case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end)='" + txtVehicle.Value + "' "
                    If txtmultiBooking.arrValueMember IsNot Nothing AndAlso txtmultiBooking.arrValueMember.Count > 0 Then
                        strQueryCANCRate += " and TSPL_SD_SHIPMENT_HEAD.document_code in (" + clsCommon.GetMulcallString(txtmultiBooking.arrValueMember) + ")"
                    End If
                    If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                        strQueryCANCRate += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                    End If
                    'dt = New DataTable()
                    'dt = clsDBFuncationality.GetDataTable(strQueryCANCRate)
                    'If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                    '    txtCanQty.Text = clsCommon.myCdbl(dt.Rows(0)("ShippedCAN"))
                    '    txtCrateQty.Text = clsCommon.myCdbl(dt.Rows(0)("crate"))
                    'End If
                    If ApplySaleRouteInDCSSale Then
                        strQuery += " " & Environment.NewLine & " union all  " & Environment.NewLine & ""
                        strQuery += "select TSPL_SD_SHIPMENT_DETAIL.PK_ID,TSPL_SD_SHIPMENT_HEAD.VehicleNo, TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document No],TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                          "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit"
                        If AllowGatePassDemandTripWise Then
                            strQuery += " ,TSPL_SD_SHIPMENT_DETAIL.Qty as Qty, TSPL_SD_SHIPMENT_DETAIL.trip_no "
                        Else
                            strQuery += " ,Qty"
                        End If
                        strQuery += ",TSPL_ITEM_MASTER.HSN_Code,Scheme_Item,TSPL_SD_SHIPMENT_DETAIL.Crate as [Crate_Issue]  " &
                          "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "

                        strQuery += "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                          "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  "

                        strQuery += "where convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103)='" & clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                          "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Sale_Route_Code='" & fndRouteNo.Value & "'"
                        End If

                        If EnableDispatch Then
                        Else
                            strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Status=1"
                            If GatepassForTaxableandNonTaxableItems Then
                                If rdbTaxable.IsChecked Then
                                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Is_Taxable=1 "
                                ElseIf rdbNonTaxable.IsChecked Then
                                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Is_Taxable=0 "
                                End If
                            End If
                        End If
                        If GenerateCustomerWiseGatePass Then
                            strQuery += " and TSPL_SD_SHIPMENT_HEAD.Document_Code='" & ShipmentDocNo & "' "
                        End If
                        If EnableProductSaleForJPR Then
                            If rbtn_Milk.IsChecked Then
                                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='S'"
                            ElseIf rbtn_product.IsChecked Then
                                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='P'"
                            ElseIf rbtn_IceCream.IsChecked Then
                                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.Item_Type='I'"
                            End If
                        End If

                        If AllowGatePassDemandTripWise Then
                            If clsCommon.myLen(txtTripNo.Text) > 0 Then
                                strQuery += "  and TSPL_SD_SHIPMENT_DETAIL.Trip_No='" + txtTripNo.Text + "'"
                            Else
                                Throw New Exception("Please Enter Trip No. ")
                            End If

                        End If
                        strQueryCANCRate = "select sum(crate) as crate ,sum(ShippedCAN) as ShippedCAN from TSPL_SD_SHIPMENT_HEAD " &
                                    "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                           "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   isnull(GPCode,'') = '' and " &
                           "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and (case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end)='" + txtVehicle.Value + "' "
                        If txtmultiBooking.arrValueMember IsNot Nothing AndAlso txtmultiBooking.arrValueMember.Count > 0 Then
                            strQueryCANCRate += " and TSPL_SD_SHIPMENT_HEAD.document_code in (" + clsCommon.GetMulcallString(txtmultiBooking.arrValueMember) + ")"
                        End If
                        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                            strQueryCANCRate += "  and TSPL_SD_SHIPMENT_HEAD.Sale_Route_Code='" + fndRouteNo.Value + "'"
                        End If
                    End If

                End If
            End If
            Return strQuery
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    '===================Added by preeti gupta against ticket no[BHA/17/08/18-000444] ========================
    Private Function LoadDoubleClickQuery(ByVal strItemCode As String) As String
        Dim strItem As String = String.Empty
        If clsCommon.myLen(strItemCode) > 0 Then
            strItem = " and TSPL_SD_SHIPMENT_DETAIL.Item_Code='" & strItemCode & "'"
        End If
        If AlternateVechileforGatePass.Equals(1) Then
            Dim StrChkAvQuery As String = "select '" & clsUserMgtCode.frmSaleDispatchDairy & "' as [Trans Type],TSPL_SD_SHIPMENT_HEAD.Document_Code as DocNo,Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code  " &
                "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   GPCode ='" + txtCode.Value + "' and " &
                "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_HEAD.AlternateVehicle='" + txtVehicle.Value + "'   and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                StrChkAvQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
            End If
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(StrChkAvQuery)
            If dt.Rows.Count > 0 Then
                strQuery = "select '" & clsUserMgtCode.frmSaleDispatchDairy & "' as [Trans Type], TSPL_SD_SHIPMENT_HEAD.Document_Code as [DocNo],Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                    "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code  " &
                    "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                    "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                    "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And   GPCode ='" + txtCode.Value + "' and " &
                    "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and TSPL_SD_SHIPMENT_HEAD.AlternateVehicle='" + txtVehicle.Value + "' and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & "  "
                If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                End If
            Else
                strQuery = "select '" & clsUserMgtCode.frmSaleDispatchDairy & "' as [Trans Type], TSPL_SD_SHIPMENT_HEAD.Document_Code as [DocNo],Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                   "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,TSPL_ITEM_MASTER.HSN_Code  " &
                   "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE " &
                   "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                   "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                   "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And    GPCode ='" + txtCode.Value + "' and " &
                   "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "' and TSPL_SD_SHIPMENT_HEAD.Vehicle_Code='" + txtVehicle.Value + "' and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " & strItem & " "
                If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                    strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
                End If
            End If
        Else
            strQuery = "select '" & clsUserMgtCode.frmSaleDispatchDairy & "' as [Trans Type],TSPL_SD_SHIPMENT_HEAD.Document_Code as [DocNo],Document_Date as [Document Date],Customer_Code,Customer_Name, " &
                    "TSPL_SD_SHIPMENT_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_SD_SHIPMENT_DETAIL.Unit_code as Unit,Qty,(Case When UsedGPDetail.GPUsedQty>0 Then (Qty-UsedGPDetail.GPUsedQty) Else Qty End) As [Balance Qty],TSPL_ITEM_MASTER.HSN_Code  " &
                    "from tspl_sd_shipment_head left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_ITEM_MASTER on TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code " &
                    "left outer join TSPL_CUSTOMER_MASTER on TSPL_SD_SHIPMENT_HEAD.Customer_Code=TSPL_CUSTOMER_MASTER.Cust_Code  " &
                    " left outer Join (Select PK_ID,	Max(GPCode)GPCode,	Item_Code,	Unit_Code,	Sum(IsNull(GP_Qty,0))GPUsedQty from TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL Group By PK_ID,Item_Code,Unit_Code)UsedGPDetail On UsedGPDetail.PK_ID=TSPL_SD_SHIPMENT_DETAIL.PK_ID " &
                    "where convert(date,Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' " & 'and TSPL_SD_SHIPMENT_DETAIL.PK_ID='" + clsCommon.myCstr(Gv1.CurrentRow.Cells(colPKID).Value) + "'" &'GPCode ='" + txtCode.Value + "' and " &
                    " and TSPL_SD_SHIPMENT_HEAD.Bill_To_Location='" & txtLocCode.Value & "'  and TSPL_SD_SHIPMENT_HEAD.Vehicle_Code='" + txtVehicle.Value + "'  and TSPL_SD_SHIPMENT_DETAIL.Item_Code <> '' " '" & strItem & "  "
            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                strQuery += "  and TSPL_SD_SHIPMENT_HEAD.route_no='" + fndRouteNo.Value + "'"
            End If
        End If
        Return strQuery
    End Function
    '=================================================================
    Private Function GetQuery(ByVal strItemCode As String, ByVal strGPCode As String) As String
        Dim strItem As String = String.Empty
        If clsCommon.myLen(strItemCode) > 0 Then
            strItem = " and TSPL_SD_SHIPMENT_DETAIL.Item_Code='" & strItemCode & "'"
        End If
        strQuery = "select TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode as [Document No],GPDate as [Document Date],'' as Customer_Code,'' as Customer_Name, " &
            "TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code as [Item Code],Item_Desc as [Item Desc],TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_code as Unit,Qty,TSPL_DAIRYSALE_GATEPASS_DETAIL.HSN_Code  " &
            "from TSPL_DAIRYSALE_GATEPASS_MASTER left outer join TSPL_DAIRYSALE_GATEPASS_DETAIL on TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode  " &
            "left outer join TSPL_ITEM_MASTER on TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code   " &
            "where TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode='" & strGPCode & "' "
        Return strQuery
    End Function
    Private Sub funFillGrid()
        Try
            LoadBlankGrid()
            Dim itemdispatchQty As New List(Of String)
            Dim totalCrate As Integer = 0
            If ApplyDepartmentRoute Then
                isDepartmentRoute = clsCommon.myCBool(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Department_Route from TSPL_ROUTE_MASTER where Route_No='" + fndRouteNo.Value + "'")))
            End If
            Dim totalCan As Integer = 0
            Dim qry As String = LoadQuery("", EnableDispatch)
            If arrShipmentFromMultiple IsNot Nothing AndAlso arrShipmentFromMultiple.Count > 0 Then
                qry += " and TSPL_SD_SHIPMENT_HEAD.Document_Code in (" + clsCommon.GetMulcallString(arrShipmentFromMultiple) + ") "
            End If
            strQuery = "Select xxfinal.PK_ID,xxfinal.[Item Code]" + IIf(VehicleNofromDispatch, ",max(VehicleNo) as VehicleNo", "") + ",max(xxfinal.[Item Desc]) as [Item Desc],xxfinal.Unit,xxfinal.Quantity,Case When Sum(TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.GP_Qty)>0 Then IsNull(((xxfinal.Quantity)-Sum(TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.GP_Qty)),0) Else xxfinal.Quantity End As BalanceQty,max(xxfinal.HSN_Code) as HSN_Code,max(xxfinal.Customer_Name) as Customer_Name,max(xxfinal.Scheme_Item)Scheme_Item,sum([Crate_Issue]) as [Crate_Issue]  from (select PK_ID,[Item Code],max([Item Desc]) as [Item Desc],Unit,sum(qty) as Quantity,max(HSN_Code) as HSN_Code, max(Customer_Name) as Customer_Name ,max(Scheme_Item)Scheme_Item,sum([Crate_Issue]) as [Crate_Issue] "
            If VehicleNofromDispatch Then
                strQuery += " ,max(VehicleNo) as VehicleNo "
            End If
            If AllowGatePassDemandTripWise Then
                strQuery += " ,max(trip_no) as Trip_No "
            End If
            strQuery += " from(" & qry & ") final  group by [Item Code],Unit,PK_ID )xxfinal
                        left Outer Join TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL ON TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.PK_ID=xxfinal.PK_ID and TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.Item_Code=xxfinal.[Item Code]"
            If AllowGatePassDemandTripWise Then
                If clsCommon.myLen(txtTripNo.Text) > 0 Then
                    strQuery += " and TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.Unit_Code=xxfinal.Unit and TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.TRip_No=xxfinal.Trip_No "
                Else
                    Throw New Exception("Please Enter Trip No.")
                End If
            End If
            strQuery += " Group By xxfinal.PK_ID, xxfinal.[Item Code], xxfinal.Unit, xxfinal.Quantity Having (Case When Sum(TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.GP_Qty)>0 Then IsNull(((xxfinal.Quantity)-Sum(TSPL_DAIRYSALE_GATEPASS_Shipment_DETAIL.GP_Qty)),0) Else xxfinal.Quantity End )>0 Order By PK_ID"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(strQuery)
            Dim intLineNo As Integer = 0
            Dim strV As String = ""
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                Dim lstDRobj As New List(Of clsDRDetail)
                ControlFields(False)
                For Each dr As DataRow In dt.Rows
                    If clsCommon.myCDecimal(dr("BalanceQty")) > 0 Then
                        Dim DRobj As New clsDRDetail
                        Gv1.Rows.AddNew()
                        intLineNo += 1
                        If VehicleNofromDispatch Then

                            strV = clsCommon.myCstr(dr("VehicleNo"))
                            If Not strVehicleNo.Contains(strV) Then
                                strVehicleNo.Add(strV)
                            End If
                        End If
                        If intLineNo = 1 Then
                            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                            '                                txtDistributorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from tspl_customer_master where cust_code in(select  top 1 x.Cust_Code 
                            'from(
                            'select TSPL_DISTRIBUTOR_ROUTE.Code as Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,max(TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code) as cust_code
                            'from TSPL_DISTRIBUTOR_ROUTE
                            'left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
                            'where  TSPL_DISTRIBUTOR_ROUTE.Status=1 and IS_Transpoter=0 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "'
                            ' Group by TSPL_DISTRIBUTOR_ROUTE.Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks
                            ') X)"))
                            Dim strQry As String
                            If clsCommon.myLen(CreditCustomer) > 0 Then
                                txtDistributorName.Text = CreditCustomer
                            Else
                                strQry = "select Customer_Name from tspl_customer_master where cust_code in(select TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code  from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER
left join TSPL_DISTRIBUTOR_ROUTE on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code=TSPL_DISTRIBUTOR_ROUTE.Code
                where TSPL_DISTRIBUTOR_ROUTE.Status=1 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code in(
select top 1 TSPL_DISTRIBUTOR_ROUTE.Code from TSPL_DISTRIBUTOR_ROUTE
left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
where TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end) order by Start_Date desc)
 and TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end))"
                                txtDistributorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strQry))
                            End If
                            '                            Dim strQry As String = "select Customer_Name from tspl_customer_master where cust_code in(select TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code  from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER
                            'left join TSPL_DISTRIBUTOR_ROUTE on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code=TSPL_DISTRIBUTOR_ROUTE.Code
                            '                where TSPL_DISTRIBUTOR_ROUTE.Status=1 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code in(
                            'select top 1 TSPL_DISTRIBUTOR_ROUTE.Code from TSPL_DISTRIBUTOR_ROUTE
                            'left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
                            'where TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end) order by Start_Date desc)
                            ' and TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end))"
                            '                                txtDistributorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strQry))
                            strQry = "select TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code  from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER
left join TSPL_DISTRIBUTOR_ROUTE on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code=TSPL_DISTRIBUTOR_ROUTE.Code
                where TSPL_DISTRIBUTOR_ROUTE.Status=1 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code in(
select top 1 TSPL_DISTRIBUTOR_ROUTE.Code from TSPL_DISTRIBUTOR_ROUTE
left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
where TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end) order by Start_Date desc)
 and TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end)"
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustCode).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strQry))
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustName).Value = clsCommon.myCstr(txtDistributorName.Text)
                            'Else
                            '    txtDistributorName.Text = clsCommon.myCstr(dr("Customer_Name"))
                            '    Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustCode).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Code from TSPL_CUSTOMER_MASTER where Customer_Name='" + clsCommon.myCstr(txtDistributorName.Text) + "'"))
                            '    Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustName).Value = clsCommon.myCstr(txtDistributorName.Text)
                            'End If
                        End If
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colPKID).Value = clsCommon.myCDecimal(dr("PK_ID"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colPKID).Tag = clsCommon.myCDecimal(dr("PK_ID"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colLineNo).Value = intLineNo
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemCode).Value = clsCommon.myCstr(dr("Item Code"))
                        DRobj.Item_Code = clsCommon.myCstr(dr("Item Code"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemDesc).Value = clsCommon.myCstr(dr("Item Desc"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colCrateIssue).Value = clsCommon.myCdbl(dr("Crate_Issue"))
                        DRobj.Unit_Code = clsCommon.myCstr(dr("Unit"))
                        DRobj.Crate = clsCommon.myCdbl(dr("Crate_Issue"))
                        If clsCommon.myCDecimal(dr("BalanceQty")) > 0 Then
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCDecimal(dr("BalanceQty"))
                            DRobj.Qty = clsCommon.myCDecimal(dr("Quantity"))
                        Else
                            Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCDecimal(dr("Quantity"))
                            DRobj.Qty = clsCommon.myCDecimal(dr("Quantity"))
                        End If


                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colHSNCode).Value = clsCommon.myCstr(dr("HSN_Code"))
                        Gv1.Rows(Gv1.Rows.Count - 1).Cells(colSchemeItem).Value = clsCommon.myCstr(dr("Scheme_Item"))
                        If clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Crate") = CompairStringResult.Equal Then
                            totalCrate += clsCommon.myRoundOFF(clsCommon.myCdbl(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value), 0, 6)
                        ElseIf clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Ltr") = CompairStringResult.Equal OrElse clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Pouch") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
                                Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dr("Item Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(dr("Unit")) & "'"))
                                Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dr("Item Code")) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dr("Item Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(dr("Unit")) & "' "))
                                Dim ItempouchCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dr("Item Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                                Dim DispatchQty As Decimal = clsCommon.myCdbl(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value) * ItemConvFactor
                                If DispatchQty > (CrateConvFactor / 2) Then
                                    totalCrate += Math.Ceiling(DispatchQty / CrateConvFactor)
                                End If
                            End If

                        End If
                        lstDRobj.Add(DRobj)
                        If clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Can") = CompairStringResult.Equal Then
                            totalCan += clsCommon.myCdbl(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value)
                        End If
                        txtDate.Enabled = False
                        'txtSupplyDate.Enabled = False
                    End If
                Next
                ' **************************************************************************************************
                If isDepartmentRoute Then
                    totalCrate = 0
                    Dim groupbyItem = From i In lstDRobj
                                      Group By i.Item_Code, i.Unit_Code Into Group
                                      Select New With {
                        Key .Item = Item_Code,
                        Key .Unit = Unit_Code,
                        Key .TotalQty = Group.Sum(Function(x) x.Qty)
                    }
                    For Each result In groupbyItem
                        Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(result.Item) & "' and tspl_unit_master.Crate_Type ='Y' "))
                        Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(result.Item) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(result.Unit) & "' "))
                        Dim DispatchQty As Decimal = clsCommon.myCdbl(result.TotalQty) * ItemConvFactor

                        totalCrate += clsCommon.myRoundOFF(DispatchQty / CrateConvFactor, 0, 9)
                        'Console.WriteLine($"Item: {result.Item}, Unit: {result.Unit}, Total Quantity: {result.TotalQty}")
                    Next
                End If
                If VehicleNofromDispatch Then
                    lblVehicleDesc.Text = clsCommon.GetMulcallStringWithComma(strVehicleNo)
                End If
                txtmultiBooking.Enabled = True
                'Dim strAllDoc As String = " select STUFF((SELECT ',' + Document_Code from (select distinct PPPP.Document_Code from  ( " & qry & "    ) As PPPP  ) Final FOR XML PATH('')), 1, 1, '') "
                Dim strAllDoc As String = " select distinct PPPP.[Document No] from  ( " & qry & "    ) As PPPP   "
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(strAllDoc)
                Dim list As New ArrayList
                If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                    For Each dr As DataRow In dt1.Rows
                        list.Add(dr("Document No"))
                    Next
                End If
                txtmultiBooking.arrValueMember = list
                txtCrateQty.Text = totalCrate
                txtCanQty.Text = totalCan
                If AllowManualCrateForDispatch Then
                    totalCrate = 0
                    Dim groupbyItem = From i In lstDRobj
                                      Group By i.Item_Code, i.Unit_Code Into Group
                                      Select New With {
                        Key .Item = Item_Code,
                        Key .Unit = Unit_Code,
                        Key .TotalQty = Group.Sum(Function(x) x.Crate)
                    }
                    For Each result In groupbyItem
                        Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(result.Item) & "' and tspl_unit_master.Crate_Type ='Y' "))
                        'Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(result.Item) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(result.Unit) & "' "))
                        'Dim DispatchQty As Decimal = clsCommon.myCdbl(result.TotalQty) * ItemConvFactor
                        'If DispatchQty > (CrateConvFactor / 2) AndAlso CrateConvFactor > 0 Then
                        '    totalCrate += clsCommon.myRoundOFF((DispatchQty / CrateConvFactor), 0, 4)
                        'End If
                        If CrateConvFactor > 0 Then
                            totalCrate += result.TotalQty
                        End If
                    Next
                    txtCrateQty.Text = totalCrate

                End If
                If DifferentCrateTypeForFGItem AndAlso Not isCTQtyUpdate Then
                    LoadgvCrateType()
                    Dim lstCTstr As List(Of String) = New List(Of String)
                    Dim ctintRow As Integer = 0
                    Dim groupbyItem = From i In lstDRobj
                                      Group By i.Item_Code, i.Unit_Code Into Group
                                      Select New With {
                        Key .Item = Item_Code,
                        Key .Unit = Unit_Code,
                        Key .TotalQty = Group.Sum(Function(x) x.Crate)
                    }
                    For Each result In groupbyItem
                        Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(result.Item) & "'"))
                        If ItemCrateType = 1 Then
                            Dim strCtCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select CrateType_Item  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(result.Item) & "'"))
                            If Not lstCTstr.Contains(strCtCode) Then
                                ctintRow += 1
                                gvCrateType.Rows.AddNew()
                                gvCrateType.Rows(ctintRow - 1).Cells(colCTICode).Value = clsCommon.myCstr(result.Item)
                                gvCrateType.Rows(ctintRow - 1).Cells(colCTCode).Value = strCtCode
                                gvCrateType.Rows(ctintRow - 1).Cells(ColCTName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Item_Desc from TSPL_ITEM_MASTER where Item_Code='" & strCtCode & "'"))
                                gvCrateType.Rows(ctintRow - 1).Cells(ColCTQty).Value = clsCommon.myCdbl(result.TotalQty)
                                lstCTstr.Add(strCtCode)
                            Else
                                For intinnerRow As Integer = 0 To gvCrateType.Rows.Count - 1
                                    If clsCommon.CompairString(strCtCode, clsCommon.myCstr(gvCrateType.Rows(intinnerRow).Cells(colCTCode).Value)) = CompairStringResult.Equal Then
                                        gvCrateType.Rows(intinnerRow).Cells(ColCTQty).Value += clsCommon.myCdbl(result.TotalQty)
                                    End If
                                Next
                            End If
                        End If
                    Next
                    UpdateCTQty()
                End If

            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found.", Me.Text)
                ' **************************************************************************************************
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gvCrateType_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gvCrateType.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column Is gvCrateType.Columns(ColCTQty) Then
                        UpdateCTQty()
                        isCTQtyUpdate = True
                    End If
                    isCellValueChangedOpen = False

                End If
            End If
        Catch ex As Exception
            isCellValueChangedOpen = False
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub UpdateCTQty()
        Try
            Dim dbltotCrate As Integer = 0
            If gvCrateType IsNot Nothing AndAlso gvCrateType.Rows.Count > 0 Then
                For intRow As Integer = 0 To gvCrateType.Rows.Count - 1
                    dbltotCrate += clsCommon.myCdbl(gvCrateType.Rows(intRow).Cells(ColCTQty).Value)
                Next
            End If
            txtCrateQty.Text = dbltotCrate
        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try

    End Sub

    Private Sub funLoadGrid(ByVal strGPCOde As String)
        Try
            LoadBlankGrid()
            Dim qry As String = GetQuery("", strGPCOde)
            Dim totalCrate As Integer = 0
            Dim totalCan As Integer = 0
            'strQuery = "select [Item Code],max([Item Desc]) as [Item Desc],Unit,sum(qty) as Quantity,max(HSN_Code) as HSN_Code from ( " & qry & " ) final group by [Item Code],Unit "
            'strQuery = "select TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID,[Item Code],max([Item Desc]) as [Item Desc],Unit,sum(qty) as Quantity,max(HSN_Code) as HSN_Code from ( " & qry & " ) final
            '            Left Outer Join TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL On TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode=final.[Document No] and TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Item_Code=final.[Item Code]
            '            group by [Item Code],Unit,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID "
            strQuery = " select TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Item_Code,max([Item_Desc]) as [Item Desc],TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Unit_Code as Unit,CONVERT(decimal(18,3), MAX(GP_Qty)) Quantity ,max(TSPL_DAIRYSALE_GATEPASS_DETAIL.HSN_Code)HSN_Code,max(TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Scheme_Item)Scheme_Item,max(TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Crate_Qty) as Crate_Qty    from TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL
                        left outer join TSPL_DAIRYSALE_GATEPASS_DETAIL on TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode  
						left outer join TSPL_ITEM_MASTER on TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code   
						where TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode= '" & strGPCOde & "'  
						group by TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Item_Code,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.Unit_Code,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID "
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(strQuery)
            Dim intLineNo As Integer = 0
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                For Each dr As DataRow In dt.Rows
                    Gv1.Rows.AddNew()
                    intLineNo += 1
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colPKID).Value = clsCommon.myCDecimal(dr("PK_ID"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colLineNo).Value = intLineNo
                    'Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemCode).Value = clsCommon.myCstr(dr("Item Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemCode).Value = clsCommon.myCstr(dr("Item_Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemDesc).Value = clsCommon.myCstr(dr("Item Desc"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCstr(dr("Quantity"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colHSNCode).Value = clsCommon.myCstr(dr("HSN_Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustCode).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Code from TSPL_CUSTOMER_MASTER where Customer_Name='" + clsCommon.myCstr(txtDistributorName.Text) + "'"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(ColCustName).Value = clsCommon.myCstr(txtDistributorName.Text)
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colSchemeItem).Value = clsCommon.myCstr(dr("Scheme_Item"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colCrateIssue).Value = clsCommon.myCdbl(dr("Crate_Qty"))
                    If clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Crate") = CompairStringResult.Equal Then
                        totalCrate += clsCommon.myRoundOFF(clsCommon.myCdbl(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value), 0, 6)
                    End If
                    If clsCommon.CompairString(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value, "Can") = CompairStringResult.Equal Then
                        totalCan += clsCommon.myCdbl(Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value)
                    End If
                    txtDate.Enabled = False
                    'txtSupplyDate.Enabled = False
                Next
                If Not isDepartmentRoute Then
                    If Not AllowManualCrateForDispatch Then
                        txtCrateQty.Text = totalCrate
                        txtCanQty.Text = totalCan
                    End If
                End If


            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "GatePass Entry", MessageBoxButtons.OK, Me.Text)
        End Try
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            btnPost.Enabled = True
            btnSave.Enabled = True
            btnPost.Enabled = True
            isInsideLoadData = True
            Addnew()
            LoadBlankGrid()
            Dim obj As New clsDairyGatePassEntry()
            obj = clsDairyGatePassEntry.GetData(strCode, NavTyep, "FS")
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.GPCode) > 0) Then
                isNewEntry = False
                If clsCommon.myLen(obj.GPCode) > 0 Then
                    btnGo.Enabled = False
                    RadGroupBox3.Enabled = False
                Else
                    RadGroupBox3.Enabled = True
                    btnGo.Enabled = True
                End If
                btnSave.Text = "Update"
                chkGhee.Checked = obj.Is_GHEE
                ControlFields(False)
                If isCreateProvisionOfTransporterInDairyDispatch = True Then
                    If obj.Post = "Y" Then
                        btnSave.Enabled = False
                        btnPost.Enabled = False
                    Else
                        btnSave.Enabled = True
                        btnPost.Enabled = True
                    End If
                Else
                    btnSave.Enabled = True
                    btnPost.Enabled = True
                    'btnPost.Visible = False
                End If
                If obj.Post = "Y" Then
                    UsLock1.Status = ERPTransactionStatus.Approved
                    btnDelete.Enabled = False
                    btnPost.Enabled = False
                    btnSave.Enabled = False
                Else
                    If obj.Status = "Y" Then
                        UsLock1.Status = ERPTransactionStatus.Cancel
                        btnSave.Enabled = False
                        btnDelete.Enabled = False
                        'btnPrint.Enabled = False
                        btnPost.Enabled = False
                    Else
                        UsLock1.Status = ERPTransactionStatus.Pending
                        btnDelete.Enabled = True
                        btnPrint.Enabled = True
                        btnPost.Enabled = True
                        btnSave.Enabled = True
                    End If
                End If
                txtCode.Value = obj.GPCode
                txtVehicle.Value = obj.Vehicle_Id
                lblVehicleDesc.Text = obj.Vehicle_Number
                If EnableProductSaleForJPR Then
                    If clsCommon.CompairString(obj.Item_Type, "M") = CompairStringResult.Equal Then
                        rbtn_Milk.IsChecked = True
                    ElseIf clsCommon.CompairString(obj.Item_Type, "P") = CompairStringResult.Equal Then
                        rbtn_product.IsChecked = True
                    ElseIf clsCommon.CompairString(obj.Item_Type, "I") = CompairStringResult.Equal Then
                        rbtn_IceCream.IsChecked = True

                    End If
                End If

                'cmbitemtype.Text = obj.Item_Type
                txtTransporter.Text = obj.Transporter
                txtSalesman.Text = obj.Salesman
                txtComments.Text = obj.Comments
                txtRemarks.Text = obj.Remarks
                isInsideLoadData = True
                'txtmultiBooking.arrValueMember = obj.AgainstDocumentCode
                txtDate.Value = obj.GPDate
                txtGatepassDate.Value = obj.GatePassDate
                txtLocCode.Value = obj.Location_Code
                txtLocDesc.Text = obj.Location_Desc
                txtTripNo.Text = obj.Trip_No
                If obj.Supply_Date IsNot Nothing Then
                    txtSupplyDate.Value = obj.Supply_Date
                End If
                '===============Added by preeti Gupta Against ticket no[BHA/17/08/18-000444]
                fndRouteNo.Value = obj.Route_No
                txtRouteName.Text = obj.Route_Desc
                If ApplyDepartmentRoute Then
                    isDepartmentRoute = clsCommon.myCBool(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Department_Route from TSPL_ROUTE_MASTER where Route_No='" + fndRouteNo.Value + "'")))
                End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    chkIndividualCustomer.Checked = IIf(obj.IsIndividualCustomer = 1, True, False)
                    If clsCommon.myLen(obj.Demand_UniqueID) Then
                        txtDemandNo.Value = obj.Demand_UniqueID
                    End If
                End If
                txtCanQty.Text = obj.TotalCAN
                txtCrateQty.Text = obj.TotalCrate
                'If clsCommon.myLen(obj.AgainstDocumentCode) > 0 Then
                '    Dim stringArray() As String = obj.AgainstDocumentCode.Split(","c)
                '    Dim myList As New ArrayList(stringArray)
                '    txtmultiBooking.arrValueMember = myList
                'End If

                txtmultiBooking.Enabled = False
                'btnSave.Enabled = False
                '=========================================================
                txtOpKM.Value = obj.Opening_Km
                txtClKM.Value = obj.Closing_Km
                txtTollAmount.Text = clsCommon.myCdbl(obj.Toll_Amount)
                lblClosingDate.Text = clsCommon.myCstr(obj.Closing_Date)
                If clsCommon.CompairString(obj.IsTransfer, "1") = CompairStringResult.Equal Then
                    chkAgainstTransfer.Checked = True
                    FndTransferNo.Value = obj.AgainstTransferNo
                End If
                If clsCommon.myCdbl(obj.Closing_Km) > 0 Then
                    btnClKM.Enabled = False
                Else
                    btnClKM.Enabled = True
                End If
                If clsCommon.CompairString(obj.ShiftType, "Morning") = CompairStringResult.Equal Then
                    rbtnMorning.IsChecked = True
                ElseIf clsCommon.CompairString(obj.ShiftType, "Evening") = CompairStringResult.Equal Then
                    rbtnEvening.IsChecked = True
                End If
                txtLoadingSlip.Text = obj.Loading_Slip
                txtDriverName.Text = obj.Driver_Name
                txtShipToLocation.Text = obj.Ship_To_Location

                txtDriverMobNo.Text = obj.Driver_ContactNo
                txtDistributorName.Text = obj.DistributorName
                funLoadGrid(txtCode.Value)
                If DifferentCrateTypeForFGItem Then
                    If obj.ArrCrateType IsNot Nothing AndAlso obj.ArrCrateType.Count > 0 Then
                        LoadgvCrateType()
                        Dim dbtotcrate As Integer = 0
                        For Each items As clsDairyGPCrateDetail In obj.ArrCrateType
                            gvCrateType.Rows(gvCrateType.Rows.Count - 1).Cells(colCTDocCode).Value = items.GPCode
                            gvCrateType.Rows(gvCrateType.Rows.Count - 1).Cells(colCTICode).Value = items.Item_Code
                            gvCrateType.Rows(gvCrateType.Rows.Count - 1).Cells(colCTCode).Value = items.CRATE_TYPE_CODE
                            gvCrateType.Rows(gvCrateType.Rows.Count - 1).Cells(ColCTName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Item_Desc from TSPL_ITEM_MASTER where Item_Code='" & items.CRATE_TYPE_CODE & "' and CRATE=1"))
                            gvCrateType.Rows(gvCrateType.Rows.Count - 1).Cells(ColCTQty).Value = items.CRATE_QTY
                            dbtotcrate += items.CRATE_QTY
                            gvCrateType.Rows.AddNew()
                        Next
                        txtCrateQty.Text = dbtotcrate
                    End If
                End If
            End If
            isInsideLoadData = False
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
        End Try
    End Sub

    Function AllowToSave() As Boolean
        Xtra.TransactionValidity(txtDate.Value)

        If AllowFutureDateTransaction(txtDate.Value, Nothing) = False Then
            txtDate.Select()
            Return False
        End If
        If ServerDateTimeForTaxableInvoice Then
            Dim serverDateTime As DateTime = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MMM/yyyy")
            If clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") <> serverDateTime Then
                Throw New Exception("Document Date and Supply Date Should be Server Date Time")
            End If
        End If
        If chkAgainstTransfer.Checked = True Then
            If clsCommon.myLen(FndTransferNo.Value) <= 0 Then
                FndTransferNo.Focus()
                Throw New Exception("Please select Transfer No")
            End If
            If clsCommon.myLen(txtVehicle.Value) <= 0 Then
                txtVehicle.Focus()
                Throw New Exception("Please select Vehicle")
            End If
        End If
        If clsCommon.myLen(fndRouteNo.Value) <= 0 Then
            fndRouteNo.Focus()
            Throw New Exception("Please select Route")
        End If
        If SettCreateProvisionOnOpeningAndClosingKM Then
            If txtOpKM.Value <= 0 Then
                txtOpKM.Focus()
                Throw New Exception("Please enter opening KM")
            End If
        End If
        If IsLoadingSlipMandatory Then
            If clsCommon.myLen(txtLoadingSlip.Text) <= 0 Then
                txtLoadingSlip.Focus()
                Throw New Exception("Please Enter Loading Slip")
            End If
        End If
        If clsCommon.myCdbl(txtCrateQty.Text) > 0 Then
            Dim CheckQty As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select IsNull(CrateCapacity,0)CrateCapacity from TSPL_VEHICLE_MASTER where Vehicle_Id='" + clsCommon.myCstr(txtVehicle.Value) + "'"))
            If clsCommon.myCdbl(txtCrateQty.Text) > CheckQty AndAlso CheckQty > 0 Then
                Throw New Exception("Vehicle Capacity:" + clsCommon.myCstr(CheckQty) + Environment.NewLine + "Allow only maximum " + clsCommon.myCstr(CheckQty) + " crate according vehicle crate capacity.")
                Return False
            End If
        End If
        Return True
        'funvalidatevehicle()
    End Function
    Private Function funvalidatevehicle() As Boolean
        Dim count As Decimal = 0
        Dim segno As String = ""
        Dim strvehiclenum As String = lblVehicleDesc.Text
        Dim sql As String = "select segment_code from TSPL_GL_SEGMENT_CODE where segment_code  = '" + Convert.ToString(txtVehicle.Value) + "' "
        If Not String.IsNullOrEmpty(connectSql.RunScalar(sql)) Then
            sql = "Select Number from TSPL_VEHICLE_MASTER where Vehicle_Id='" + txtVehicle.Value + "'"
            lblVehicleDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(sql))
            Return True
        Else
            Dim strmessage As String = "This vehicle code doesn't exist" + Environment.NewLine
            strmessage += "Do you want to continue "
            If common.clsCommon.MyMessageBoxShow(strmessage, Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                If clsCommon.myLen(lblVehicleDesc.Text) <= 0 Then
                    lblVehicleDesc.Focus()
                    Throw New Exception("Please Enter Vehicle No")
                End If
                txtVehicle.Value = clsCommon.incval(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Max(Segment_code) from TSPL_GL_SEGMENT_CODE where Segment_name = 'Vehicles'")))
                strvehiclenum = txtVehicle.Text
                sql = "select seg_no from tspl_gl_segment where seg_name='Vehicles'"
                segno = CStr(connectSql.RunScalar(sql))
                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                Try
                    connectSql.RunSpTransaction(trans, "sp_tspl_gl_segmentcode_insert", New SqlParameter("@segno", segno), New SqlParameter("@segmentname", "Vehicles"), New SqlParameter("@segmentcode", txtVehicle.Value), New SqlParameter("@desc", strvehiclenum), New SqlParameter("@acccode", "NULL"), New SqlParameter("@createdby", objCommonVar.CurrentUserCode), New SqlParameter("@createddate", connectSql.serverDate(trans)), New SqlParameter("@modifiedby", objCommonVar.CurrentUserCode), New SqlParameter("@modifieddate", connectSql.serverDate(trans)), New SqlParameter("@compcode", objCommonVar.CurrentCompanyCode))
                    connectSql.RunSpTransaction(trans, "SP_TSPL_VEHICLE_MASTER_INSERT", New SqlParameter("@Vehicle_Id", txtVehicle.Value), New SqlParameter("@Model", ""), New SqlParameter("@Number", strvehiclenum), New SqlParameter("@Description", strvehiclenum), New SqlParameter("@Type", "H"), New SqlParameter("@Start_Date", Date.Now.ToString("dd/MM/yyyy")), New SqlParameter("@End_Date", Date.Now.ToString("dd/MM/yyyy")), New SqlParameter("@Vehicle_Reg_No", ""), New SqlParameter("@Vehicle_Chesis_No", ""), New SqlParameter("@Capacity", "0"), New SqlParameter("@Insurance", ""), New SqlParameter("@Pollution_Check", ""), New SqlParameter("@Fitness", Date.Now.ToString("dd/MM/yyyy")), New SqlParameter("@Trans_Type", ""), New SqlParameter("@Road_Tax", ""), New SqlParameter("@Transport_Id", ""), New SqlParameter("@Created_By", objCommonVar.CurrentUserCode), New SqlParameter("@Created_Date", connectSql.serverDate(trans)), New SqlParameter("@Modified_By", objCommonVar.CurrentUserCode), New SqlParameter("@Modified_Date", connectSql.serverDate(trans)), New SqlParameter("@Comp_Code", objCommonVar.CurrentCompanyCode))
                    trans.Commit()
                Catch ex As Exception
                    txtVehicle.Value = ""
                    trans.Rollback()
                    Throw New Exception(ex.Message)
                End Try
                lblVehicleDesc.Text = txtVehicle.Text + "-Hired"
                txtVehicle.Text = txtVehicle.Value
                Return True
            Else
                txtVehicle.Value = String.Empty
                txtVehicle.Text = txtVehicle.Value
                Return False
            End If
        End If
    End Function
    Sub SaveData()
        Try
            If (AllowToSave()) Then
                Dim obj As New clsDairyGatePassEntry()
                obj.GPCode = txtCode.Value
                obj.GPDate = clsCommon.myCDate(txtDate.Value)
                obj.GatePassDate = clsCommon.myCDate(txtGatepassDate.Value)
                obj.Vehicle_Id = txtVehicle.Value
                If VehicleDesc Is Nothing AndAlso clsCommon.myLen(VehicleDesc) <= 0 Then
                    obj.Vehicle_Number = lblVehicleDesc.Text
                Else
                    obj.Vehicle_Number = VehicleDesc
                End If
                'obj.Item_Type = cmbitemtype.Text
                obj.Transporter = txtTransporter.Text
                obj.Salesman = txtSalesman.Text
                obj.Remarks = txtRemarks.Text
                obj.Comments = txtComments.Text
                obj.Ship_To_Location = txtShipToLocation.Text
                obj.Location_Code = txtLocCode.Value
                obj.Location_Desc = txtLocDesc.Text
                obj.Supply_Date = txtSupplyDate.Value
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    obj.IsIndividualCustomer = IIf(chkIndividualCustomer.Checked, 1, 0)
                    If clsCommon.myLen(txtDemandNo.Value) Then
                        obj.Demand_UniqueID = clsCommon.myCdbl(txtDemandNo.Value)
                    End If
                End If
                '===============Added by preeti Gupta Agianst Ticket no[]
                obj.Route_No = fndRouteNo.Value
                obj.TotalCrate = clsCommon.myCdbl(txtCrateQty.Text)
                obj.TotalCAN = clsCommon.myCdbl(txtCanQty.Text)
                obj.AgainstDocumentCode = clsCommon.GetMulcallString(txtmultiBooking.arrValueMember)
                obj.Opening_Km = txtOpKM.Value
                obj.Loading_Slip = txtLoadingSlip.Text
                obj.Driver_Name = txtDriverName.Text
                obj.Driver_ContactNo = txtDriverMobNo.Text
                obj.DistributorName = txtDistributorName.Text
                obj.Trip_No = clsCommon.myCdbl(txtTripNo.Text)
                If EnableProductSaleForJPR Then
                    If rbtn_Milk.IsChecked Then
                        obj.Item_Type = "M"
                        If CreateGatePassFromDemand = True Then
                            If rbtnEvening.IsChecked = True Then
                                obj.ShiftType = "Evening"
                            ElseIf rbtnMorning.IsChecked = True Then
                                obj.ShiftType = "Morning"
                            End If
                        Else
                            If rbtnEvening.IsChecked = True Then
                                obj.ShiftType = "Evening"
                            ElseIf rbtnMorning.IsChecked = True Then
                                obj.ShiftType = "Morning"
                            End If
                        End If
                    ElseIf rbtn_product.IsChecked Then
                        obj.Item_Type = "P"
                        obj.ShiftType = "Morning"
                    ElseIf rbtn_IceCream.IsChecked Then
                        obj.Item_Type = "I"
                        obj.ShiftType = "Morning"
                    End If
                Else
                    obj.Item_Type = "M"
                    If CreateGatePassFromDemand = True Then
                        If rbtnEvening.IsChecked = True Then
                            obj.ShiftType = "Evening"
                        ElseIf rbtnMorning.IsChecked = True Then
                            obj.ShiftType = "Morning"
                        End If
                    Else
                        If rbtnEvening.IsChecked = True Then
                            obj.ShiftType = "Evening"
                        ElseIf rbtnMorning.IsChecked = True Then
                            obj.ShiftType = "Morning"
                        End If
                    End If
                End If

                If chkAgainstTransfer.Checked = True Then
                    obj.IsTransfer = 1
                    obj.AgainstTransferNo = clsCommon.myCstr(FndTransferNo.Value)
                End If
                obj.Is_GHEE = chkGhee.Checked
                '=======================================================

                Dim totalCrate As Integer = 0
                Dim totalCan As Integer = 0
                obj.Arr = New List(Of clsDairyGPDetail)
                For Each grow As GridViewRowInfo In Gv1.Rows
                    If clsCommon.myCdbl(grow.Cells(colQty).Value) > 0 Then
                        Dim objTr As New clsDairyGPDetail()
                        objTr.PK_ID = clsCommon.myCDecimal(grow.Cells(colPKID).Value)
                        objTr.Item_Code = clsCommon.myCstr(grow.Cells(colItemCode).Value)
                        objTr.Unit_Code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                        objTr.Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                        objTr.HSN_Code = clsCommon.myCstr(grow.Cells(colHSNCode).Value)
                        objTr.Scheme_Item = clsCommon.myCstr(grow.Cells(colSchemeItem).Value)
                        objTr.Trip_No = clsCommon.myCdbl(txtTripNo.Text)
                        objTr.Cust_Code = clsCommon.myCstr(grow.Cells(ColCustCode).Value)
                        objTr.Customer_Name = clsCommon.myCstr(grow.Cells(ColCustName).Value)
                        objTr.Crate_Qty = clsCommon.myCstr(grow.Cells(colCrateIssue).Value)
                        obj.Arr.Add(objTr)
                    End If
                Next

                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Document", Me.Text)
                    Return
                End If
                If DifferentCrateTypeForFGItem Then
                    obj.ArrCrateType = New List(Of clsDairyGPCrateDetail)
                    For Each grow As GridViewRowInfo In gvCrateType.Rows
                        If clsCommon.myCdbl(grow.Cells(ColCTQty).Value) > 0 Then
                            Dim objTr As New clsDairyGPCrateDetail()
                            objTr.Item_Code = clsCommon.myCDecimal(grow.Cells(colCTICode).Value)
                            objTr.CRATE_TYPE_CODE = clsCommon.myCstr(grow.Cells(colCTCode).Value)
                            objTr.CRATE_QTY = clsCommon.myCdbl(grow.Cells(ColCTQty).Value)
                            obj.ArrCrateType.Add(objTr)
                        End If
                    Next
                End If
                Dim CheckVehicle As String = clsDBFuncationality.getSingleValue("select Vehicle_Number from TSPL_DAIRYSALE_GATEPASS_Master where Convert(Date,GPDate,103)=Convert(Date,'" + txtDate.Value + "',103) And Route_No='" + clsCommon.myCstr(fndRouteNo.Value) + "' And ShiftType='" + clsCommon.myCstr(obj.ShiftType) + "' and Vehicle_Number='" + clsCommon.myCstr(obj.Vehicle_Number) + "'")
                If CheckVehicle IsNot Nothing AndAlso clsCommon.myLen(CheckVehicle) > 0 Then
                    If clsCommon.MyMessageBoxShow(Me, "Vehicle No. is already used for another gatepass." + Environment.NewLine + "Do you want to proceed?", Me.Text, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                        If (obj.SaveData(obj, isNewEntry, "DS")) Then
                            common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                            LoadData(obj.GPCode, NavigatorType.Current)
                            arrShipmentFromMultiple = Nothing
                            btnGo.Enabled = False
                            RadGroupBox3.Enabled = False
                        End If
                    End If
                Else
                    If (obj.SaveData(obj, isNewEntry, "DS")) Then
                        common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                        LoadData(obj.GPCode, NavigatorType.Current)
                        arrShipmentFromMultiple = Nothing
                        btnGo.Enabled = False
                        RadGroupBox3.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub Addnew()
        txtLoadingSlip.Text = ""
        If AllowGatePassDemandTripWise Then
            txtTripNo.Visible = True
            lblTripNo.Visible = True
        Else
            txtTripNo.Visible = False
            lblTripNo.Visible = False
        End If
        isDepartmentRoute = False
        strVehicleNo = New List(Of String)
        txtTripNo.Text = ""
        txtTollAmount.Text = 0
        txtCode.Value = ""
        txtDate.Value = clsCommon.GETSERVERDATE
        txtSupplyDate.Value = txtDate.Value
        txtGatepassDate.Value = clsCommon.GETSERVERDATE()
        txtVehicle.Value = ""
        lblVehicleDesc.Text = ""
        If CreateGatePassFromDemand = False Then
            txtLocCode.Value = ""
            txtLocDesc.Text = ""
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            chkIndividualCustomer.Visible = True
            chkIndividualCustomer.Checked = False
            txtDemandNo.Value = ""
            lblDemandNo.Visible = True
            txtDemandNo.Visible = True
            txtDemandNo.Enabled = False
        Else
            chkIndividualCustomer.Visible = False
            chkIndividualCustomer.Checked = False
            lblDemandNo.Visible = False
            txtDemandNo.Visible = False
            txtDemandNo.Value = ""
        End If
        ControlFields(True)
        txtDistributorName.Text = ""
        txtTransporter.Text = ""
        txtSalesman.Text = ""
        txtRemarks.Text = ""
        txtComments.Text = ""
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnSave.Text = "Save"
        LoadBlankGrid()
        isCTQtyUpdate = False
        isNewEntry = True
        'cmbitemtype.Text = "Select"
        rbtn_Milk.IsChecked = True
        If EnableProductSaleForJPR Then
            rgbItemType.Visible = True
        Else
            rgbItemType.Visible = False
        End If

        isInsideLoadData = False
        fndRouteNo.Value = ""
        txtRouteName.Text = ""
        txtCanQty.Text = 0
        txtCrateQty.Text = 0
        txtmultiBooking.Enabled = False
        txtmultiBooking.arrValueMember = Nothing
        btnDelete.Enabled = True
        btnClKM.Enabled = False
        lblClosingDate.Text = ""
        UsLock1.Status = ERPTransactionStatus.Pending
        chkAgainstTransfer.Checked = False
        chkGhee.Checked = False
        FndTransferNo.Value = ""
        txtTransporter.Enabled = False
        txtOpKM.Value = 0
        txtClKM.Value = 0
        txtLocCode.Enabled = True
        TxtMultiDairyGPassReverse.arrValueMember = Nothing
        txtDriverName.Text = Nothing
        txtShipToLocation.Text = Nothing
        txtDriverMobNo.Text = Nothing
        If SetDefaultShiftTime.Length > 0 Then
            Dim CurrDateTime As DateTime = clsCommon.GETSERVERDATE
            Dim EndTime As DateTime = clsCommon.GetPrintDate(SetDefaultShiftTime, "dd/MMM/yyyy hh:mm tt")
            If CurrDateTime.TimeOfDay < EndTime.TimeOfDay Then
                txtSupplyDate.Value = clsCommon.GetPrintDate(CurrDateTime)
                rbtnEvening.IsChecked = True

            Else
                txtSupplyDate.Value = clsCommon.GetPrintDate(CurrDateTime.AddDays(1))
                rbtnMorning.IsChecked = True
            End If
        End If
        CheckTimeShift()
    End Sub
    Private Sub cmbitemtype_SelectedIndexChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs)
        'If isInsideLoadData = False Then
        '    txtVehicle.Value = ""
        '    lblVehicleDesc.Text = ""
        '    txtTransporter.Text = ""
        '    txtRemarks.Text = ""
        '    txtComments.Text = ""
        '    LoadBlankGrid()
        '    isNewEntry = True
        '    isInsideLoadData = False
        '    funFillGrid()
        'End If
    End Sub
    Private Sub txtDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'If isInsideLoadData = False Then
        '    If clsCommon.CompairString(cmbitemtype.Text, "Select") = CompairStringResult.Equal Then
        '        clsCommon.MyMessageBoxShow("Please select Item type")
        '    Else
        '        funFillGrid()
        '    End If
        'End If
    End Sub
    Private Sub Gv1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub txtLocCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtLocCode._MYValidating
        strQuery = "select distinct Bill_To_Location as Code,Location_Desc as Description from TSPL_SD_SALE_INVOICE_HEAD left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location=TSPL_LOCATION_MASTER.Location_Code"
        txtLocCode.Value = clsCommon.ShowSelectForm("LocationSegGP", strQuery, "Code", "screen_type='DS'", txtLocCode.Value, "Code", isButtonClicked)
        txtLocDesc.Text = clsDBFuncationality.getSingleValue("select  Location_Desc  from TSPL_LOCATION_MASTER where Location_Code='" & txtLocCode.Value & "'")
    End Sub
    Private Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
        If isInsideLoadData = False Then
            'If clsCommon.myLen(txtVehicle.Value) <= 0 AndAlso clsCommon.myLen(fndRouteNo.Value) <= 0 Then
            '    clsCommon.MyMessageBoxShow("Please Select  Vehicle code OR Route No OR Both.")
            '    Exit Sub
            'End If
            If CreateGatePassFromDemand = True Then
                If clsCommon.myLen(fndRouteNo.Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Please Select Route No.", Me.Text)
                    fndRouteNo.Focus()
                    Exit Sub
                End If
            End If
            LoadBlankGrid()
            isNewEntry = True
            isInsideLoadData = False
            If chkAgainstTransfer.Checked = True Then
                If clsCommon.myLen(FndTransferNo.Value) > 0 Then
                    funFillGrid_Transfer()
                Else
                    clsCommon.MyMessageBoxShow(Me, "Please select Transfer No", Me.Text)
                End If
            Else
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    If chkIndividualCustomer.Checked Then
                        If clsCommon.myLen(txtDemandNo.Value) > 0 Then
                            funFillGrid()
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Please Select Demand")
                        End If
                    Else
                        funFillGrid()
                    End If
                Else
                    funFillGrid()
                End If

            End If
        End If
    End Sub
    Private Sub frmDairyGatePass_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'Add Tool tip Task No- TEC/18/05/18-000237
        If e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnNew.Enabled Then
            btnNew.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso btnSave.Enabled AndAlso MyBase.isModifyFlag Then
            SaveData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso btnSave.Enabled AndAlso MyBase.isDeleteFlag Then
            btnDelete.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnPost.Enabled AndAlso MyBase.isPostFlag Then
            btnPost.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
            Me.Close()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then

                ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                                             "TSPL_DAIRYSALE_GATEPASS_MASTER " + Environment.NewLine +
                                             "TSPL_DAIRYSALE_GATEPASS_DETAIL  ")
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
        ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.F11 Then
            If btnMultiGPReverse.Visible Then
                multipleDelteVisible(False)
            Else
                Dim pwd As New FrmPWD(Nothing)
                pwd.strCode = clsFixedParameterCode.MultiDairyGatePassReversePWD
                pwd.strType = clsFixedParameterType.MultiDairyGatePassReversePWD
                pwd.ShowDialog()
                If pwd.isPasswordCorrect Then
                    multipleDelteVisible(True)
                End If
            End If
        ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.F10 Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIR
            frm.strCode = "GatePass Password"
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                EnableDispatch = True
            End If
        End If
    End Sub
    Sub multipleDelteVisible(ByVal val As Boolean)
        MyLabel34.Visible = val
        TxtMultiDairyGPassReverse.Visible = val
        btnMultiGPReverse.Visible = val
    End Sub
    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub txtVehicle__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtVehicle._MYValidating
        'done by priti BHA/25/07/18-000192 to remove transtype condition and added screen type condition.
        strQuery = ReturnVehicle()
        VehicleDesc = Nothing
        txtVehicle.Value = clsCommon.ShowSelectForm("Vehicle", strQuery, "Code", "", txtVehicle.Value, "Code", isButtonClicked)
        lblVehicleDesc.Text = clsDBFuncationality.getSingleValue("select Description from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'")
        If isCreateProvisionOfTransporterInDairyDispatch = True Then
            txtTransporter.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'"))
        End If
    End Sub
    Private Function ReturnVehicle() As String
        Dim Query As String = Nothing
        If CreateGatePassFromDemand = True Then
            Query = "select distinct Final.Vehicle_Code as Code,Description,tspl_route_master.Route_No as [Route No], tspl_route_master.Route_Desc as [Route Desc]  from ( " &
        "select distinct Vehicle_Code,TSPL_VEHICLE_MASTER.Description from TSPL_DEMAND_BOOKING_DETAIL left outer join " &
        "TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code=TSPL_VEHICLE_MASTER.Vehicle_Id  ) final left outer join tspl_route_master on tspl_route_master.vehicle_code = Final.Vehicle_Code "
        Else
            Query = "select distinct Final.Vehicle_Code as Code,Description,tspl_route_master.Route_No as [Route No], tspl_route_master.Route_Desc as [Route Desc]  from ( " &
        "select distinct Vehicle_Code,TSPL_VEHICLE_MASTER.Description from TSPL_SD_SHIPMENT_HEAD left outer join " &
        "TSPL_VEHICLE_MASTER on TSPL_SD_SHIPMENT_HEAD.Vehicle_Code=TSPL_VEHICLE_MASTER.Vehicle_Id where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' " &
        "union all " &
        "select distinct AlternateVehicle as Vehicle_Code,TSPL_VEHICLE_MASTER.Description from TSPL_SD_SHIPMENT_HEAD left outer join " &
        "TSPL_VEHICLE_MASTER on TSPL_SD_SHIPMENT_HEAD.AlternateVehicle=TSPL_VEHICLE_MASTER.Vehicle_Id where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' and AlternateVehicle <> '' " &
        "union all " &
        "select distinct ManualVehicle,'' as Description from TSPL_SD_SHIPMENT_HEAD where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' and ManualVehicle <> '' ) final left outer join tspl_route_master on tspl_route_master.vehicle_code = Final.Vehicle_Code"
        End If
        Return Query
    End Function
    Private Sub txtCode__MYNavigator(ByVal sender As Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtCode._MYNavigator
        Try
            Dim qry As String = "select count(*) from TSPL_DAIRYSALE_GATEPASS_MASTER where GPCode='" + txtCode.Value + "' "
            Dim count As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
            If count = 0 Then
                txtCode.MyReadOnly = False
            Else
                txtCode.MyReadOnly = True
            End If
            LoadData(txtCode.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtCode._MYValidating
        ' Dim qry As String = " SELECT  GPCode,convert(varchar(10),GPDate,103)  as GPDate,Vehicle_Id,Vehicle_Number FROM  TSPL_DAIRYSALE_GATEPASS_MASTER"
        ' Ticket No : ERO/23/05/19-000614 By prabhakar
        'Ticket No-ERO/05/08/19-000984 ,Sanjay, add pending / approved 
        'Ticket No-ERO/27/08/19-001004 ,Add Opening_Km,Closing_Km
        Dim qry As String = " SELECT  TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode,convert(varchar(10),TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate,103)  as GPDate,TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType,TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id,TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number,TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No,tspl_Route_Master.Route_Desc,TSPL_DAIRYSALE_GATEPASS_MASTER.Item_Type as [Item Type], case when TSPL_DAIRYSALE_GATEPASS_MASTER.Post='Y' then 'Approved' else 'Pending' end as Status,Opening_Km,Closing_Km ,isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.AgainstTransferNo,'') as [Against Transfer No], Case When TSPL_DAIRYSALE_GATEPASS_MASTER.Status='Y' Then 'Cancel' Else Null End As [GP Status] FROM  TSPL_DAIRYSALE_GATEPASS_MASTER " &
                            " left Outer join tspl_Route_Master on tspl_Route_Master.Route_No = TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No "
        LoadData(clsCommon.ShowSelectForm("GatepassEntry", qry, "GPCode", "", txtCode.Value, "GPCode", isButtonClicked, " TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate "), NavigatorType.Current)
        If clsCommon.myLen(txtCode.Value) > 0 Then
            txtCode.MyReadOnly = False
            btnGo.Enabled = False
            RadGroupBox3.Enabled = False
        Else
            txtCode.MyReadOnly = True
            btnGo.Enabled = True
            RadGroupBox3.Enabled = True
        End If
        'funLoadGrid(txtCode.Value)
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If clsCommon.CompairString(btnSave.Text, "Update") = CompairStringResult.Equal AndAlso OneTimeCheck = False Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIRC
            frm.strCode = clsFixedParameterCode.UpdatePassword
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                OneTimeCheck = True
            Else
                Exit Sub
            End If
        End If
        If clsCommon.CompairString(btnSave.Text, "Update") = CompairStringResult.Equal AndAlso OneTimeCheck Then
            SaveData()
        Else
            SaveData()
        End If
    End Sub
    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Addnew()
        cmbitemtype.Enabled = True
        txtDate.Enabled = True
        txtSupplyDate.Enabled = True
        btnGo.Enabled = True
        RadGroupBox3.Enabled = True
        VehicleDesc = Nothing
    End Sub
    Private Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Try
            If clsCommon.myLen(txtCode.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Document No not found to Post", Me.Text)
                Exit Sub
            End If
            Dim isPost As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select count (*) from TSPL_DAIRYSALE_GATEPASS_MASTER where GPCode = '" + txtCode.Value + "' and Post = 'Y'"))
            If isPost = True Then
                common.clsCommon.MyMessageBoxShow(Me, "Record Already posted.", Me.Text)
                Exit Sub
            End If
            If myMessages.postConfirm() Then
                If (clsDairyGatePassEntry.PostData(MyBase.Form_ID, txtCode.Value)) Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    UploadDoc(txtCode.Value)
                    Dim strQry As String = "select distinct Document_Code from TSPL_SD_SHIPMENT_DETAIL where PK_ID in(select PK_ID from TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL where GPCode='" + txtCode.Value + "')"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        For Each dr As DataRow In dt.Rows
                            Dim invoiceno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Sale_Invoice_No from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + clsCommon.myCstr(dr("Document_Code")) + "'"))
                            Dim docDate As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Date from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + clsCommon.myCstr(dr("Document_Code")) + "'"))
                            Dim CustCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Code from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + clsCommon.myCstr(dr("Document_Code")) + "'"))
                            UploadInvoice(invoiceno, docDate, CustCode)
                        Next
                    End If
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                    LoadData(txtCode.Value, NavigatorType.Current)

                    'btnSave.Enabled = False
                    'btnPost.Enabled = False
                End If
                'clsDBFuncationality.ExecuteNonQuery("Update TSPL_DAIRYSALE_GATEPASS_MASTER set post='Y' where gpcode='" & txtCode.Value & "'")
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Dim I As Integer = 0
        If (Gv1.Rows.Count > 0) AndAlso clsCommon.CompairString(btnSelect.Text, "Select All") = CompairStringResult.Equal Then
            For Each grow As GridViewRowInfo In Gv1.Rows
                grow.Cells(ColApply).Value = True
            Next
            btnSelect.Text = "Unselect All"
        ElseIf (Gv1.Rows.Count > 0) AndAlso clsCommon.CompairString(btnSelect.Text, "Unselect All") = CompairStringResult.Equal Then
            For Each grow As GridViewRowInfo In Gv1.Rows
                grow.Cells(ColApply).Value = False
            Next
            btnSelect.Text = "Select All"
        End If
    End Sub
    Private Sub Gv1_CurrentRowChanged(sender As Object, e As CurrentRowChangedEventArgs) Handles Gv1.CurrentRowChanged
    End Sub
    Private Sub gv1_CellDoubleClick(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles Gv1.CellDoubleClick
        Try
            Dim strItem = Gv1.Rows(e.RowIndex).Cells(1).Value
            strQuery = LoadDoubleClickQuery(strItem)
            Dim frmStock As New FrmStockDetail
            frmStock.LoadDispatchData(strQuery)
            frmStock.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Public Function CrateInOut() As String
        Return CrateInOut(txtDate.Value, fndRouteNo.Value, txtLocCode.Value)
    End Function

    Public Function CrateInOut(ByVal strDate As DateTime, ByVal strRoute As String, ByVal strLocation As String) As String
        Dim qry As String = Nothing
        Try
            If clsCommon.myLen(strRoute) > 0 AndAlso clsCommon.myLen(strLocation) Then
                '                Dim StrQry As String = "select  top 1 x.Cust_Code 
                'from(
                'select TSPL_DISTRIBUTOR_ROUTE.Code as Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,max(TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code) as cust_code
                'from TSPL_DISTRIBUTOR_ROUTE
                'left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
                'where  TSPL_DISTRIBUTOR_ROUTE.Status=1 and IS_Transpoter=0 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + fndRouteNo.Value + "'
                ' Group by TSPL_DISTRIBUTOR_ROUTE.Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks
                ') X"
                'qry = " WITH my_cte AS ( select ROW_NUMBER() over ( Partition by 1 order by Sale_Invoice_Date ) as SNO, * from ( select max(Customer_Name) Customer_Name, max(Comp_Name) Comp_Name, max(Location_Desc) Location_Desc, max(Location_Code) Location_Code, max(Vehicle_Id) Vehicle_Id, max(Vehicle_Number) Vehicle_Number, max(Route_No) Route_No, max(Route_Desc) Route_Desc, max(Customer_Code) Customer_Code, Sale_Invoice_Date, sum( Qty * case when RI =-1 THEN 1 else 0 end * case when ShiftType = 'M' then 1 else 0 end ) as Morning_Supply, sum( Qty * case when RI = 1 THEN 1 else 0 end * case when ShiftType = 'M' then 1 else 0 end ) as Morning_Return, sum( Qty * case when RI =-1 THEN 1 else 0 end * case when ShiftType = 'E' then 1 else 0 end ) as Evening_Supply, sum( Qty * case when RI = 1 THEN 1 else 0 end * case when ShiftType = 'E' then 1 else 0 end ) as Evening_Return from ( select TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Vehicle_Code AS Vehicle_Id, TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.VehicleNo AS Vehicle_Number, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code as Route_No, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.ShiftType, TSPL_route_master.Route_Desc, TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code, TSPL_CUSTOMER_MASTER.Customer_Name, CAST( TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Sale_Invoice_Date AS DATE ) AS Sale_Invoice_Date, TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.CrateQtyRecd as Qty, 1 as RI, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code, TSPL_LOCATION_MASTER.Location_Desc, TSPL_COMPANY_MASTER.Comp_Name From TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE left outer join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No = TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Document_No left outer join tspl_route_master on tspl_route_master.route_No = TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.route_Code left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code = TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Comp_Code left outer join tspl_customer_master on tspl_customer_master.Cust_Code = TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code union all select TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number, TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No, CASE WHEN TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType = 'Morning' THEN 'M' ELSE 'E' END AS ShiftType, tspl_route_master.Route_Desc, ( select cust_code from TSPL_CUSTOMER_MASTER where Route_No = '" + fndRouteNo.Value + "' and IsDistributor = 'Y' ) as Customer_Code, ( select Customer_Name from TSPL_CUSTOMER_MASTER where Route_No = '" + fndRouteNo.Value + "' and IsDistributor = 'Y' ) as Customer_Name, CAST( TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate AS DATE ) AS Sale_Invoice_Date, TSPL_DAIRYSALE_GATEPASS_MASTER.TotalCrate as Qty, -1 as RI, TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code, TSPL_LOCATION_MASTER.Location_Desc, TSPL_COMPANY_MASTER.Comp_Name from TSPL_DAIRYSALE_GATEPASS_MASTER left OUTER join tspl_route_master on tspl_route_master.route_no =  TSPL_DAIRYSALE_GATEPASS_MASTER.route_no LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code = TSPL_DAIRYSALE_GATEPASS_MASTER.Comp_Code where 2 = 2 and TSPL_DAIRYSALE_GATEPASS_MASTER.Status is null ) xx where 2 = 2 and Location_Code = '" + txtLocCode.Value + "' and Customer_Code In ('" + clsCommon.myCstr(clsDBFuncationality.getSingleValue(StrQry)) + "') group by Sale_Invoice_Date ) xxx ) select Comp_Name, Location_Code, Location_Desc, Vehicle_Id, Vehicle_Number, Route_No, Route_Desc, Customer_Code, Customer_Name, Sale_Invoice_Date, OP as Opening, (Morning_Supply + Evening_Supply) as [Issue], (Morning_Return + Evening_Return) as [Receive], ( OP +( (Morning_Supply + Evening_Supply)-(Morning_Return + Evening_Return) ) ) as Balance from ( select ( select isnull( sum( (Morning_Supply + Evening_Supply)-(Morning_Return + Evening_Return) ), 0 ) from my_cte as InnCTE where InnCTE.Sale_Invoice_Date < my_cte.Sale_Invoice_Date ) as OP, * from my_cte where Sale_Invoice_Date >= '" + clsCommon.GetPrintDate(txtDate.Value) + "' and Sale_Invoice_Date <= '" + clsCommon.GetPrintDate(txtDate.Value) + "' ) xx order by xx.Sale_Invoice_Date asc"
                qry = " With CTETemp as ( Select convert(varchar,Doc_Date,103) as Doc_Date,Vehicle_Code,Vehicle_Name,Customer_Code,Customer_Name,Customer_Phone, ZSM , ZSM_NAME,ZSM_Phone,ASM,ASM_Name,ASM_Phone,ASO , ASO_Name,ASO_Phone ,OpencrateQty, OpenJaaliQty, OpenBoxQty,OpenCanQty, CrateQtyRecd, JaaliQtyRecd, BoxQtyRecd ,CanQtyRecd,CrateOutQty, jaaliOutQty,boxOutQty  ,CanOutQty, CrateAdjQty , JaaliAdjQty  , BoxAdjQty ,CanAdjQty,SUM(CrateQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code) as CrateQtyClosing,  SUM(JaaliQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code) as JaaliQtyClosing,  SUM(BoxQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code) as BoxQtyClosing, SUM(CanQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code) as CanQtyClosing , Row_Number() OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code) as RowNo from(
 select  pp.Doc_Date  as Doc_Date,pp.Vehicle_Code,tspl_vehicle_master.Number as Vehicle_Name ,pp.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Phone1 as Customer_Phone ,   TSPL_CUSTOMER_MASTER.ZSM,TSPL_EMPLOYEE_MASTER_ZSM.Emp_Name as ZSM_NAME,case when len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ZSM.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ZSM.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO else '' end  as ZSM_Phone ,TSPL_CUSTOMER_MASTER.ASM, TSPL_EMPLOYEE_MASTER_ASM.Emp_Name as ASM_NAME, case when len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ASM.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASM.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO else '' end  as ASM_Phone,TSPL_CUSTOMER_MASTER.ASO,TSPL_EMPLOYEE_MASTER_ASO.Emp_Name as ASO_NAME,   case when len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ASO.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASO.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO else '' end  as ASO_Phone,pp.OpencrateQty as OpencrateQty,pp.OpenJaaliQty  as OpenJaaliQty,pp.OpenBoxQty  as OpenBoxQty,pp.OpenCanQty  as OpenCanQty,pp.CrateQtyRecd  as CrateQtyRecd,pp.JaaliQtyRecd  as JaaliQtyRecd,pp.BoxQtyRecd  as BoxQtyRecd ,pp.CanQtyRecd  as CanQtyRecd,pp.CrateOutQty  as CrateOutQty,pp.jaaliOutQty  as jaaliOutQty ,pp.boxOutQty  as boxOutQty,pp.CanOutQty  as CanOutQty ,pp.CrateQtyClosing as CrateQtyClosing, pp.JaaliQtyClosing as  JaaliQtyClosing, pp.BoxQtyClosing as BoxQtyClosing, pp.CanQtyClosing as CanQtyClosing,pp.CrateAdjQty , pp.JaaliAdjQty  , pp.BoxAdjQty , pp.CanAdjQty from ( 
 select  max(convert(date,Doc_Date,103))  as Doc_Date,max(xx.Vehicle_Code) as Vehicle_Code,xx.Customer_Code,sum(xx.OpencrateQty) as OpencrateQty,sum(xx.OpenJaaliQty ) as OpenJaaliQty,sum(xx.OpenBoxQty )  as OpenBoxQty,sum(xx.OpenCanQty )  as OpenCanQty,sum(xx.CrateQtyRecd) as CrateQtyRecd,sum(xx.JaaliQtyRecd) as JaaliQtyRecd,sum(xx.BoxQtyRecd) as BoxQtyRecd,sum(xx.CanQtyRecd) as CanQtyRecd,sum(xx.CrateOutQty ) as CrateOutQty,sum(xx.jaaliOutQty ) as jaaliOutQty ,sum(xx.boxOutQty ) as boxOutQty,sum(xx.CanOutQty ) as CanOutQty, sum(xx.CrateAdjQty ) as CrateAdjQty ,sum(xx.JaaliAdjQty )as JaaliAdjQty  ,sum(xx.BoxAdjQty )  as BoxAdjQty ,sum(xx.CanAdjQty )  as CanAdjQty,(sum(xx.OpencrateQty)+sum(xx.CrateOutQty )-sum(xx.CrateQtyRecd)-sum(xx.CrateAdjQty )) as CrateQtyClosing, (sum(xx.OpenJaaliQty)+sum(xx.jaaliOutQty)-sum(xx.JaaliQtyRecd)-sum(xx.JaaliAdjQty )) as JaaliQtyClosing, (sum(xx.OpenBoxQty)+sum(xx.boxOutQty )-sum(xx.BoxQtyRecd)-sum(xx.BoxAdjQty )) as BoxQtyClosing  , (sum(xx.OpenCanQty)+sum(xx.CanOutQty )-sum(xx.CanQtyRecd)-sum(xx.CanAdjQty )) as CanQtyClosing  from (select  max(Route_No) as Route_No,max(convert(date,'" & clsCommon.GetPrintDate(strDate) & "',103)) as Doc_Date,max(Opening.Vehicle_Code) as Vehicle_Code , Opening.Customer_Code ,sum(Opening.OpencrateQty*Type) as OpencrateQty,sum(Opening.OpenJaaliQty*Type ) as OpenJaaliQty,sum(Opening.OpenBoxQty *Type)  as OpenBoxQty,sum(Opening.OpenCanQty *Type)  as OpenCanQty,sum(Opening.CrateQtyRecd *Type) as CrateQtyRecd,sum(Opening.JaaliQtyRecd*Type ) as JaaliQtyRecd,sum(Opening.BoxQtyRecd*Type ) as BoxQtyRecd,sum(Opening.CanQtyRecd*Type ) as CanQtyRecd,sum(Opening.CrateOutQty*Type ) as CrateOutQty,sum(Opening.jaaliOutQty*Type ) as jaaliOutQty ,sum(Opening.boxOutQty*Type ) as boxOutQty,sum(Opening.CanOutQty*Type ) as CanOutQty,sum(Opening.CrateAdjQty*Type ) as CrateAdjQty,sum(Opening.JaaliAdjQty*Type ) as JaaliAdjQty,sum(Opening.BoxAdjQty *Type) as  BoxAdjQty ,sum(Opening.CanAdjQty *Type) as  canAdjQty from  ( select TSPL_SD_SHIPMENT_HEAD.Document_Date    as Document_Date,isnull(TSPL_SD_SHIPMENT_HEAD.Route_no,'') as Route_no,(case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end) as Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,1 as Type  ,'O' as Type1,TSPL_SD_SHIPMENT_HEAD.Crate as OpencrateQty,TSPL_SD_SHIPMENT_HEAD.jaali as OpenJaaliQty, TSPL_SD_SHIPMENT_HEAD.box  as OpenBoxQty , TSPL_SD_SHIPMENT_HEAD.ShippedCAN  as OpenCanQty  ,0 as CrateQtyRecd,0  as JaaliQtyRecd ,0 as BoxQtyRecd ,0 as CanQtyRecd ,0 as CrateOutQty,0 as jaaliOutQty,0 as boxOutQty ,0 as CanOutQty,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty  ,0 as  CanAdjQty  from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS'  AND TSPL_SD_SHIPMENT_HEAD.Status =1  union all  select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code,'') as Route_no,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , (TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty+TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment) as OpencrateQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty   as OpenjaaliQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty  as OpenboxQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty as OpenCanQty , 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1  union all select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code,'') as Route_no,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment,0)  as OpencrateQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment,0)    as OpenjaaliQty , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment,0)  as OpenboxQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment,0)  as OpenCanQty, 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1   )as Opening WHERE convert(date,Document_Date,103)<(convert(date,'" & clsCommon.GetPrintDate(strDate) & "',103)) group by Customer_Code 
 UNION All 
select Route_No,Document_Date,Vehicle_Code,Customer_Code,0 as OpencrateQty,0 as OpenjaaliQty ,0 as OpenboxQty,0 as OpenCanQty,Case When [Type]=1 Then CrateQtyRecd Else 0 End as CrateQtyRecd,Case When [Type]=1 Then JaaliQtyRecd Else 0 End as JaaliQtyRecd,Case When [Type]=1 Then BoxQtyRecd Else 0 End as BoxQtyRecd,Case When [Type]=1 Then CANQtyRec Else 0 End as CANQtyRec,
 Case When [Type]=-1 Then CrateOutQty Else 0 End as CrateOutQty,Case When [Type]=-1 Then jaaliOutQty Else 0 End as jaaliOutQty,Case When [Type]=-1 Then boxOutQty Else 0 End as boxOutQty,Case When [Type]=-1 Then CanOutQty Else 0 End as CanOutQty,CrateAdjQty*Type as CrateAdjQty,Case When [Type]=1 Then JaaliAdjQty Else 0 End as JaaliAdjQty,Case When [Type]=1 Then BoxAdjQty Else 0 End as BoxAdjQty,Case When [Type]=1 Then CANAdjustment Else 0 End as CANAdjustment 
 from ((select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code,'') as Route_no,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCanQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE
 left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No 
 where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1 ) union all 
 (select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code,'') as Route_no,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCANQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE
 left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No 
 where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1  union all 
 select TSPL_SD_SHIPMENT_HEAD.Document_Date   as Document_Date,isnull(TSPL_SD_SHIPMENT_HEAD.Route_no,'') as Route_no ,(case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end) as Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,-1 as Type  ,'O' as Type1,0 as OpencrateQty,0  as OpenBoxQty ,0 as OpenJaaliQty,0 as OpenCANQty, 0 as CrateQtyRecd, 0 JaaliQtyRecd ,  0  as BoxQtyRecd,0 CanQtyRecd  ,TSPL_SD_SHIPMENT_HEAD.Crate as CrateOutQty,jaali  as jaaliOutQty,TSPL_SD_SHIPMENT_HEAD.Box as boxOutQty,TSPL_SD_SHIPMENT_HEAD.ShippedCAN as CanOutQty ,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty,0 as  CanAdjQty    from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' AND TSPL_SD_SHIPMENT_HEAD.Status =1)  ) as Closing 
 WHERE convert(date,Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(strDate) & "',103) AND convert(date,Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(strDate) & "',103) 
   ) as xx inner join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = xx.Customer_Code   AND TSPL_CUSTOMER_MASTER.Status='N' where 2=2  and xx.Route_No in ('" & strRoute & "') 
 GROUP BY Customer_Code 
 ) as pp   left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =pp.Customer_Code 
 left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ZSM on TSPL_EMPLOYEE_MASTER_ZSM.EMP_CODE = TSPL_CUSTOMER_MASTER.ZSM 
 left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ASM on TSPL_EMPLOYEE_MASTER_ASM.EMP_CODE = TSPL_CUSTOMER_MASTER.ASM  
 left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ASO on TSPL_EMPLOYEE_MASTER_ASO.EMP_CODE = TSPL_CUSTOMER_MASTER.ASO  
 left join tspl_vehicle_master on tspl_vehicle_master.vehicle_id=pp.vehicle_code where 2=2   ) YYY ) Select OpencrateQty as Opening, CrateOutQty as Issue, CrateQtyRecd as Receive, OpencrateQty + CrateOutQty - CrateQtyRecd - CrateAdjQty as Balance  from (Select  case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySumOfVehicleWise, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySumOfVehicleWise, CTETemp.RowNo,CTETemp.Doc_Date ,CTETemp.Vehicle_Code,CTETemp.Vehicle_Name ,CTETemp.Customer_Code,CTETemp.Customer_Name, CTETemp.Customer_Phone, CTETemp.ZSM,CTETemp.ZSM_NAME,CTETemp.ZSM_Phone,CTETemp.ASM,CTETemp.ASM_NAME,CTETemp.ASM_Phone,CTETemp.ASO,CTETemp.ASO_NAME,CTETemp.ASO_Phone, CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) as OpencrateQty,  CTETemp.OpenJaaliQty+ISNULL(CT1.JaaliQtyClosing,0) as OpenJaaliQty,  CTETemp.OpenBoxQty+ISNULL(CT1.BoxQtyClosing,0) as OpenBoxQty, CTETemp.OpenCanQty+ISNULL(CT1.CanQtyClosing,0) as OpenCanQty,  CTETemp.CrateQtyRecd, CTETemp.JaaliQtyRecd, CTETemp.BoxQtyRecd,CTETemp.CanQtyRecd, CTETemp.CrateOutQty, CTETemp.jaaliOutQty, CTETemp.boxOutQty, CTETemp.CanOutQty, CTETemp.CrateAdjQty , CTETemp.JaaliAdjQty  , CTETemp.BoxAdjQty, CTETemp.CanAdjQty  from CTETemp LEFt OUTER JOIN CTETemp CT1 ON  CT1.Customer_Code=CTETemp.Customer_Code And  CT1.Vehicle_Code = CTETemp.Vehicle_Code    And (CTETemp.RowNo-CT1.RowNo)=1 ) ZZZ ORDER BY  Customer_Code,convert(date,Doc_Date,103),Vehicle_Code"
            Else
                Throw New Exception(" Please Select Route No/ Location ")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return qry
    End Function
    Public Function CountUOM(ByVal strCode As String) As String
        Dim qry As String = Nothing
        Try
            If clsCommon.myLen(strCode) > 0 Then
                qry = "select TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code,sum(TSPL_DAIRYSALE_GATEPASS_DETAIL.Qty)as Qty from TSPL_DAIRYSALE_GATEPASS_MASTER
left join TSPL_DAIRYSALE_GATEPASS_DETAIL on TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode
where TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode='" + strCode + "'
group by TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code"
            Else
                Throw New Exception("Dcoument not Found!")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return qry
    End Function

    Private Function GetProductPrintQry(ByVal strCode As String) As String
        Return GetProductPrintQry(strCode, False)
    End Function

    Private Function GetProductPrintQry(ByVal strCode As String, ByVal isCancel As Boolean) As String
        Dim tbl_DAIRYSALE_GATEPASS_MASTER As String = Nothing
        Dim tbl_DAIRYSALE_GATEPASS_DETAIL As String = Nothing
        If isCancel Then
            tbl_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_DETAIL "
        Else
            tbl_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_DETAIL "
        End If
        Dim Qry As String = "select '" & clsCommon.myCstr(IIf(isCancel, "Y", "N")) & "' As isCancelled,TSPL_COMPANY_MASTER.Comp_Code,TSPL_COMPANY_MASTER.Insurance_No, TSPL_COMPANY_MASTER.Insurance_Comp_Name, TSPL_COMPANY_MASTER.comp_name,TSPL_COMPANY_MASTER.ISO_No,TSPL_COMPANY_MASTER.add1 + case when len(TSPL_COMPANY_MASTER.add2)> 0 then ', ' + TSPL_COMPANY_MASTER.add2 else '' end + case when LEN( isnull(TSPL_COMPANY_MASTER.Add3, '') )> 0 then ', ' + isnull(TSPL_COMPANY_MASTER.Add3, '') else ' ' end as Comp_Address, tspl_location_master.add1 + case when len(tspl_location_master.add2)> 0 then ', ' + tspl_location_master.add2 else '' end + case when LEN( isnull(tspl_location_master.Add3, '') )> 0 then ', ' + isnull(tspl_location_master.Add3, '') else ' ' end as Loc_add, tspl_company_master.GSTReg_No,TSPL_COMPANY_MASTER.Logo_Img, FORMAT( TSPL_DAIRYSALE_GATEPASS_MASTER.Supply_Date, 'dd/MM/yyyy' ) as Supply_Date,TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType,TSPL_DAIRYSALE_GATEPASS_MASTER.DistributorName As 'Distributor', isnull( TSPL_DAIRYSALE_GATEPASS_MASTER.AgainstTransferNo, '' ) as AgainstTransferNo,TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No,TSPL_ROUTE_MASTER.Zone_Code,TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode, convert( varchar, TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, 103 ) as GPDate,TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id AS vehicle_id, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number as VehicleDesc, TSPL_DAIRYSALE_GATEPASS_MASTER.location_code, tspl_location_master.Location_desc, TSPL_DAIRYSALE_GATEPASS_MASTER.transporter, TSPL_DAIRYSALE_GATEPASS_MASTER.remarks, TSPL_DAIRYSALE_GATEPASS_MASTER.comments, TSPL_DAIRYSALE_GATEPASS_MASTER.post,tspl_item_master.sku_seq,TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_Name,TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_ContactNo,TSPL_DAIRYSALE_GATEPASS_MASTER.Loading_Slip,TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code,tspl_item_master.Short_Description,BulkUnit.UOM_Code as BulkUOM,CurrentUnit.UOM_Code as CurrentUOM,TSPL_DAIRYSALE_GATEPASS_DETAIL.Qty,floor( (TSPL_DAIRYSALE_GATEPASS_DETAIL.Qty * CurrentUnit.Conversion_Factor) / nullif(BulkUnit.Conversion_Factor, 0) ) as BulkQty,LooseUnit.UOM_Code as LooseUOM,(((TSPL_DAIRYSALE_GATEPASS_DETAIL.Qty * CurrentUnit.Conversion_Factor)- floor((TSPL_DAIRYSALE_GATEPASS_DETAIL.Qty * CurrentUnit.Conversion_Factor) / nullif(BulkUnit.Conversion_Factor, 0))* BulkUnit.Conversion_Factor)/ nullif(LooseUnit.Conversion_Factor, 0)) as LooseQty
from " & tbl_DAIRYSALE_GATEPASS_MASTER & "   
left join " & tbl_DAIRYSALE_GATEPASS_DETAIL & " on TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode=TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode  
left join tspl_vehicle_master on tspl_vehicle_master.Vehicle_id=TSPL_DAIRYSALE_GATEPASS_MASTER.vehicle_id  
left join tspl_location_master on tspl_location_master.location_code=TSPL_DAIRYSALE_GATEPASS_MASTER.location_code  
left join tspl_company_master on tspl_company_master.comp_code=TSPL_DAIRYSALE_GATEPASS_MASTER.comp_code  
left join tspl_item_master on tspl_item_master.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code  
left join  TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id  
left join  tspl_route_master on tspl_route_master.Route_No=TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No  
left join tspl_item_uom_detail CurrentUnit on CurrentUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and 	CurrentUnit.uom_code=	TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code
left join tspl_item_uom_detail BulkUnit on BulkUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and BulkUnit.Bulk_UOM='1'
left join tspl_item_uom_detail  LooseUnit on LooseUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and LooseUnit.Loose_UOM='1'
where TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode='" & strCode & "' order by Sku_Seq "
        Return Qry
    End Function
    '----------------------------------BKNGATEPASSPRINTADDED--------------------------
    Private Function GetAttachQryBKN(ByVal StrCode As String, ByVal isDepartmentRoute As Boolean) As String
        Return GetAttachQryBKN(StrCode, isDepartmentRoute, False)
    End Function
    Private Function GetAttachQryBKN(ByVal StrCode As String, ByVal isDepartmentRoute As Boolean, ByVal isCancel As Boolean) As String
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_MASTER As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_DETAIL As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_HEAD As String = Nothing
        Dim tbl_TSPL_BOOKING_MATSER As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL As String = Nothing
        Dim Customer_Name As String = ""
        Dim add1 As String = ""
        Dim add2 As String = ""
        Dim add3 As String = ""
        Dim PIN_Code As String = ""
        Dim Phone1 As String = ""
        Dim Phone2 As String = ""
        Dim GSTNO As String = ""

        If isCancel Then
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD_Cancel_Data As TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER_Cancel_Data As TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
        Else
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = "  TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
        End If
        Dim strqry As String = "select Customer_Name,add1,add2,add3,PIN_Code,Phone1,Phone2,GSTNO from TSPL_CUSTOMER_MASTER where Cust_Code in(
select Customer_Code from " + tbl_TSPL_SD_SHIPMENT_HEAD + " where Document_Code in(
select distinct DOCUMENT_CODE from " + tbl_TSPL_SD_SHIPMENT_DETAIL + " where PK_ID in( select PK_ID from " + tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL + " where GPCode='" + StrCode + "'))) and IsDistributor='Y'"
        Dim dts As DataTable = clsDBFuncationality.GetDataTable(strqry)
        If dts IsNot Nothing AndAlso dts.Rows.Count > 0 Then
            Customer_Name = clsCommon.myCstr(dts.Rows(0)("Customer_Name"))
            add1 = clsCommon.myCstr(dts.Rows(0)("add1"))
            add2 = clsCommon.myCstr(dts.Rows(0)("add2"))
            add3 = clsCommon.myCstr(dts.Rows(0)("add3"))
            PIN_Code = clsCommon.myCstr(dts.Rows(0)("PIN_Code"))
            Phone1 = clsCommon.myCstr(dts.Rows(0)("Phone1"))
            Phone2 = clsCommon.myCstr(dts.Rows(0)("Phone2"))
            GSTNO = clsCommon.myCstr(dts.Rows(0)("GSTNO"))
        End If

        Dim Qry As String = ""
        Qry = " Select '" & clsCommon.myCstr(IIf(isCancel, "Y", "N")) & "' As isCancelled,max(xxx.CopyType)CopyType,STRING_AGG(TabBatch.BatchNo,CHAR(10)) as BatchNo,--TabBatch.BatchNo,
  STRING_AGG(CAST(Qty AS INT), CHAR(10)) as Batch_Qty,max(Comp_Phone2)Comp_Phone2,max(Comp_Add3)Comp_Add3,max(Comp_Add2)Comp_Add2,max(ConversionInLtr)ConversionInLtr,max(ConversionInCrate)ConversionInCrate, max(ConversionInPouch)ConversionInPouch,max(GP.GPDate)Document_Date,max(Is_Taxable)Is_Taxable,max(Shift)Shift,max(xxx.DOCUMENT_CODE)DOCUMENT_CODE,max(Line_No)Line_No, max(Row_Type)Row_Type,xxx.Item_Code,Sum(Qty)Qty,Sum(Balance_Qty)Balance_Qty,max(Order_Code)Order_Code,max(Unit_code)Unit_code,max(Location)Location, max(Item_Cost)Item_Cost,max(TAX1)TAX1,max(TAX1_Base_Amt)TAX1_Base_Amt,max(TAX1_Rate)TAX1_Rate,sum(TAX1_Amt)TAX1_Amt,max(TAX2)TAX2,max(TAX2_Base_Amt)TAX2_Base_Amt,max(TAX2_Rate)TAX2_Rate,sum(TAX2_Amt)TAX2_Amt,max(TAX3)TAX3,max(TAX3_Base_Amt)TAX3_Base_Amt,max(TAX3_Rate)TAX3_Rate,sum(TAX3_Amt)TAX3_Amt,max(TAX4)TAX4,max(TAX4_Base_Amt)TAX4_Base_Amt,max(TAX4_Rate)TAX4_Rate,sum(TAX4_Amt)TAX4_Amt,max(TAX5)TAX5,max(TAX5_Base_Amt)TAX5_Base_Amt,max(TAX5_Rate)TAX5_Rate,sum(TAX5_Amt)TAX5_Amt,max(TAX6)TAX6,max(TAX6_Base_Amt)TAX6_Base_Amt,max(TAX6_Rate)TAX6_Rate,sum(TAX6_Amt)TAX6_Amt,max(TAX7)TAX7,max(TAX7_Base_Amt)TAX7_Base_Amt,max(TAX7_Rate)TAX7_Rate,sum(TAX7_Amt)TAX7_Amt,max(TAX8)TAX8,max(TAX8_Base_Amt)TAX8_Base_Amt,max(TAX8_Rate)TAX8_Rate,sum(TAX8_Amt)TAX8_Amt,max(TAX9)TAX9,max(TAX9_Base_Amt)TAX9_Base_Amt,max(TAX9_Rate)TAX9_Rate,sum(TAX9_Amt)TAX9_Amt,max(TAX10)TAX10,max(TAX10_Base_Amt)TAX10_Base_Amt,max(TAX10_Rate)TAX10_Rate,sum(TAX10_Amt)TAX10_Amt,sum(Amount)Amount,max(Disc_Per)Disc_Per,sum(Disc_Amt)Disc_Amt,sum(Amt_Less_Discount)Amt_Less_Discount,sum(Total_Tax_Amt)Total_Tax_Amt,sum(Item_Net_Amt)Item_Net_Amt,max(Status)Status,max(MRP)MRP,max(Batch_No)Batch_No, max(MFG_Date)MFG_Date,max(Expiry_Date)Expiry_Date,sum(Free_Qty)Free_Qty,max(Specification)Specification,max(Remarks)Remarks,max(Assessable)Assessable,
 sum(AssessableAmt)AssessableAmt,max(Is_Mannual_Amt)Is_Mannual_Amt,max(Bar_Code)Bar_Code,max(Price_code)Price_code,max(Price_Date)Price_Date,max(Abatement_Per)Abatement_Per,sum(Abatement_Amt)Abatement_Amt,max(Scheme_Code)Scheme_Code,max(Scheme_Applicable)Scheme_Applicable,max(Scheme_Item)Scheme_Item,max(FOC_Item)FOC_Item,max(Item_Tax)Item_Tax,max(Total_MRP_Amt)Total_MRP_Amt,sum(Total_Basic_Amt)Total_Basic_Amt,sum(Total_Disc_Amt)Total_Disc_Amt,sum(Cust_DiscountQty)Cust_DiscountQty,max(Cust_Discount)Cust_Discount,sum(Total_Cust_Discount)Total_Cust_Discount,max(Price_Amount1)Price_Amount1,max(Price_Amount2)Price_Amount2,max(Price_Amount3)Price_Amount3,max(Price_Amount4)Price_Amount4,max(Price_Amount5)Price_Amount5,max(Price_Amount6)Price_Amount6,max(Price_Amount7)Price_Amount7,max(Price_Amount8)Price_Amount8,max(Price_Amount9)Price_Amount9,max(Price_Amount10)Price_Amount10,max(ActualRate)ActualRate,max(Item_Weight)Item_Weight,max(Conv_Factor)Conv_Factor,max(TotalItem_Weight)TotalItem_Weight,max(Markup_On)Markup_On,max(Markup_Percent)Markup_Percent,max(Landing_Cost)Landing_Cost,sum(HeadDiscAmt)HeadDiscAmt,max(CustDiscPer)CustDiscPer,max(CasdDiscScheme_Code)CasdDiscScheme_Code,max(Purchase_Cost)Purchase_Cost,max(OrgRate)OrgRate,max(PrincipleCode)PrincipleCode,max(PrincipleDesc)PrincipleDesc,max(vendor_code)vendor_code,
 max(vendor_desc)vendor_desc,max(Bin_No)Bin_No,max(Weight_UOM)Weight_UOM,max(HeadDiscPer)HeadDiscPer,sum(HeadDiscPerAmt)HeadDiscPerAmt,sum(DeliverQty)DeliverQty,max(Delivery_Code)Delivery_Code,sum(Crate)Crate,max(Commission_Rate)Commission_Rate,max(Commission_Party)Commission_Party,sum(Commission_Amt)Commission_Amt,
  sum(Amt_Less_Commission)Amt_Less_Commission, max(OrgUnit_code)OrgUnit_code, max(Delivery_Code_PS)Delivery_Code_PS,max(Item_Group)Item_Group,
 max(BOOK_QTY_UOM)BOOK_QTY_UOM,max(BOOK_Rate)BOOK_Rate,max(BOOK_RATE_UOM)BOOK_RATE_UOM,max(TAX_PAID)TAX_PAID,max(Alternate_UOM)Alternate_UOM,max(RATE_UOM)RATE_UOM,max(Scheme_Type)Scheme_Type,max(Scheme_Item_Code)Scheme_Item_Code,sum(Scheme_Qty)Scheme_Qty,max(Scheme_Item_UOM)Scheme_Item_UOM,max(Total_Item_WeightMetric)Total_Item_WeightMetric,max(Cash_Scheme_Code)Cash_Scheme_Code,max(Cash_Scheme_Type)Cash_Scheme_Type,max(Cash_Scheme_Pers)Cash_Scheme_Pers,sum(Cash_Scheme_Amount)Cash_Scheme_Amount,max(OrgRateUnit_code)OrgRateUnit_code,max(Rate_UnitQty)Rate_UnitQty,max(Alter_UnitQty)Alter_UnitQty,max(Sampling)Sampling,max(GatePass_No)GatePass_No,max(Disc_Scheme_Code)Disc_Scheme_Code,max(Disc_Scheme_Type)Disc_Scheme_Type,max(Disc_Scheme_Pers)Disc_Scheme_Pers,sum(Disc_Scheme_Amount)Disc_Scheme_Amount,max(AlternateRate)AlternateRate,max(ItemwiseTaxCode)ItemwiseTaxCode,max(Structure_Code)Structure_Code,sum(CAN)CAN,sum(ManualCan)ManualCan,sum(ItemLeakageAmount)ItemLeakageAmount,max(VS_CashSchemeCode)VS_CashSchemeCode,Sum(VS_Cash_Amt)VS_Cash_Amt,sum(VS_ltrInCrate)VS_ltrInCrate,max(Sub_Location_code)Sub_Location_code,max(Distributor_Commission_PKID)Distributor_Commission_PKID,max(Distributor_Commission_Rate)Distributor_Commission_Rate,max(Distributor_Commission_RateWithTax)Distributor_Commission_RateWithTax,sum(Distributor_Commission_Amt)Distributor_Commission_Amt,max(PK_ID)PK_ID,max(Security_Rate)Security_Rate,sum(Security_Amt)Security_Amt,max(Transporter_Commission_Rate)Transporter_Commission_Rate,Sum(Transporter_Commission_Amt)Transporter_Commission_Amt,max(Transporter)Transporter,max(Against_Booking_PK_ID)Against_Booking_PK_ID,max(Booth_Security_Rate)Booth_Security_Rate,sum(Booth_Security_Amt)Booth_Security_Amt,max(Scheme_Main_Item)Scheme_Main_Item,max(Disc_Per_Unit)Disc_Per_Unit,sum(Disc_Unit_Amt)Disc_Unit_Amt,max(REF_PK_ID)REF_PK_ID,max(REF_TPT_PK_ID)REF_TPT_PK_ID,max(Trip_No)Trip_No,max(Against_Cust_Ord_PK_ID)Against_Cust_Ord_PK_ID,max(Billing_Unit_code)Billing_Unit_code, sum(Billing_Qty)Billing_Qty,sum(Amount_with_Tax)Amount_with_Tax,sum(Booking_Qty)Booking_Qty,max(Access_Officer)Access_Officer,max(Item_Desc)Item_Desc,max(HSN_Code)HSN_Code, Sum(QtyInLtr)QtyInLtr,sum(QtyInKG)QtyInKG,sum(QtyInCrate)QtyInCrate,sum(QtyInPouch)QtyInPouch,max(FAT_Per)FAT_Per,max(SNF_Per)SNF_Per,max(Acidity)Acidity,max(Temperature)Temperature,max(MBRT_Hours)MBRT_Hours,max(Route_Desc)Route_Desc,max(Vehicle_Id)Vehicle_Id,max(Vehicle_Number)Vehicle_Number,
"
        If clsCommon.myLen(Customer_Name) > 0 Then
            Qry += " '" & Customer_Name & "' as Customer_Name,'" & add1 & "' as Cust_Add1,'" & add2 & "' as Cust_Add2,'" & add3 & "' as Cust_Add3,'" & PIN_Code & "' as Cust_PINCode,'" & Phone1 & "' as Cust_Phone1,'" & Phone2 & "' as Cust_Phone2,'" & GSTNO & "' as GSTNO, "
        Else
            Qry += " max(Customer_Name)Customer_Name,max(Cust_Add1)Cust_Add1,max(Cust_Add2)Cust_Add2,max(Cust_Add3)Cust_Add3,max(Cust_PINCode)Cust_PINCode,max(Cust_Phone1)Cust_Phone1,max(Cust_Phone2)Cust_Phone2,max(GSTNO)GSTNO, "
        End If

        Qry += "  max(Comp_Name)Comp_Name	
 ,max(Comp_Add1)Comp_Add1,max(Comp_City)Comp_City,max(Comp_State)Comp_State,max(Comp_GSTReg_No)Comp_GSTReg_No,max(Comp_PanNo)Comp_PanNo,max(Comp_Email)Comp_Email,
 max(Comp_Pincode)Comp_Pincode,max(Comp_Phone1)Comp_Phone1,max(State_Code)State_Code,max(STATE_NAME)STATE_NAME,max(Route_No)Route_No,max(GP.supply_date)supply_date,
 max(sublocation)sublocation,max(Comp_Code1)Comp_Code1,max(Is_Ambient)Is_Ambient,max(Is_FreshItem)Is_FreshItem,max(COL1)COL1,max(COL2)COL2,max(CopyType1)CopyType1,
 '" & txtCode.Value & "' as GPCode,'" & txtShipToLocation.Text & "' as Ship_To_Location,'" & txtDriverName.Text & "' as Driver_Name ,max(orderby1)orderby1,max(orderby2)orderby2,max(orderby3)orderby3,Max(GatePass_Code)GatePass_Code
  from ("
        Qry += " select 1 As CopyType, '' as Comp_Phone2,'' as Comp_Add3,'' as Comp_Add2,  InLtr.Conversion_factor As [ConversionInLtr],InCrate.Conversion_factor As [ConversionInCrate],InPouch.Conversion_factor As [ConversionInPouch],TSPL_SD_SHIPMENT_HEAD.Document_Date,
                        TSPL_SD_SHIPMENT_HEAD.DO_Item_Type as Is_Taxable,Case When TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' OR TSPL_SD_SHIPMENT_HEAD.Shift_Type='M' Then '[M]' Else '[E]' End As Shift,
                        TSPL_SD_SHIPMENT_DETAIL.*,TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt as Amount_with_Tax ,TSPL_SD_SHIPMENT_DETAIL.Qty as Booking_Qty,TSPL_COMPANY_MASTER.Access_Officer,
                        TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.HSN_Code,
                        case when coalesce(InLtr.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InLtr.Conversion_factor,1)) end as QtyInLtr, 
case when coalesce(InKG.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InKG.Conversion_factor,1)) end as QtyInKG,
                        case when coalesce(InCrate.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InCrate.Conversion_factor,1)) end as QtyInCrate,
                        case when coalesce(InPouch.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InPouch.Conversion_factor,1)) end as QtyInPouch,
                        TSPL_SD_SHIPMENT_HEAD.FAT_Per,TSPL_SD_SHIPMENT_HEAD.SNF_Per,TSPL_SD_SHIPMENT_HEAD.Acidity,TSPL_SD_SHIPMENT_HEAD.Temperature,TSPL_SD_SHIPMENT_HEAD.MBRT_Hours, 
                        TSPL_Route_Master.Route_Desc,TSPL_VEHICLE_MASTER.Vehicle_Id,case when Len(TSPL_SD_SHIPMENT_HEAD.VehicleNo) >0 then TSPL_SD_SHIPMENT_HEAD.VehicleNo else TSPL_VEHICLE_MASTER.Number  end As Vehicle_Number,TSPL_CUSTOMER_MASTER.IsDistributor,
                        TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 As Cust_Add1,TSPL_CUSTOMER_MASTER.Add2 As Cust_Add2,TSPL_CUSTOMER_MASTER.Add3 As Cust_Add3,TSPL_CUSTOMER_MASTER.PIN_Code As Cust_PINCode,
                        TSPL_CUSTOMER_MASTER.Phone1 As Cust_Phone1,TSPL_CUSTOMER_MASTER.Phone2 As Cust_Phone2,TSPL_CUSTOMER_MASTER.GSTNO,
                        TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img As Comp_Logo1,TSPL_COMPANY_MASTER.Logo_Img2 As Comp_Logo2,TSPL_COMPANY_MASTER.Add1 As Comp_Add1,

TSPL_COMPANY_MASTER.City_Code As Comp_City,TSPL_COMPANY_MASTER.State As Comp_State,TSPL_COMPANY_MASTER.GSTReg_No As Comp_GSTReg_No,TSPL_COMPANY_MASTER.Pan_No As Comp_PanNo,
                        TSPL_COMPANY_MASTER.Email As Comp_Email,TSPL_COMPANY_MASTER.Pincode As Comp_Pincode,TSPL_COMPANY_MASTER.Phone1 As Comp_Phone1,
                        TSPL_STATE_MASTER.GST_STATE_Code As State_Code,TSPL_STATE_MASTER.STATE_NAME,TSPL_SD_SHIPMENT_HEAD.Route_No,supply_date,TSPL_SD_SHIPMENT_HEAD.sub_location_code as sublocation, TSPL_COMPANY_MASTER.Comp_Code1,TSPL_ITEM_MASTER.Is_Ambient,TSPL_ITEM_MASTER.Is_FreshItem ,TSPL_ITEM_MASTER.IsTaxable ,
						CASE WHEN (CASE WHEN Is_FreshItem = 1 THEN 1 ELSE 0 END) = 1
     AND (CASE WHEN IsTaxable = 0 THEN 1 ELSE 0 END) = 1 THEN 1 ELSE 2 END as orderby1,CASE WHEN (CASE WHEN Is_Ambient = 1 THEN 1 ELSE 0 END) = 1
     AND (CASE WHEN IsTaxable = 0 THEN 1 ELSE 0 END) = 1 THEN 1 ELSE 2 END as orderby2,CASE WHEN ((CASE WHEN Is_FreshItem = 1 OR Is_Ambient = 1 THEN 1 ELSE 2 END) = 1
        ) AND ((CASE WHEN IsTaxable = 1 THEN 1 ELSE 0 END) = 1 ) THEN 1 ELSE 0 END AS orderby3,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode As GatePass_Code
                        from  " & tbl_TSPL_SD_SHIPMENT_HEAD & "
                        Left Outer Join " & tbl_TSPL_SD_SHIPMENT_DETAIL & " On TSPL_SD_SHIPMENT_DETAIL.Document_Code=TSPL_SD_SHIPMENT_HEAD.Document_Code
                        Left Outer Join " & tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL & " On TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID=TSPL_SD_SHIPMENT_DETAIL.PK_ID
                        Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_SD_SHIPMENT_HEAD.Bill_To_Location 
                        Left Outer Join TSPL_Route_Master On TSPL_Route_Master.Route_No=TSPL_SD_SHIPMENT_HEAD.Route_No 
                        Left Outer Join TSPL_VEHICLE_MASTER On TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.Vehicle_Code 
                        left outer join TSPL_CITY_MASTER    On  TSPL_CITY_MASTER.City_Code =TSPL_LOCATION_MASTER.City_Code  
                        left outer join TSPL_STATE_MASTER   On TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.state  
                        Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code
                        left outer join TSPL_ITEM_UOM_DETAIL  on TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_ITEM_MASTER.Item_Code  and   TSPL_ITEM_UOM_DETAIL.UOM_Code =TSPL_SD_SHIPMENT_DETAIL.Unit_code
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.LTR_TYPE='Y') as InLtr on InLtr.Item_code=TSPL_ITEM_MASTER.Item_Code  
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL   where TSPL_ITEM_UOM_DETAIL.UOM_Code='KG') as InKG on InKG.Item_code=TSPL_ITEM_MASTER.Item_Code
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL 	  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Crate_type='Y') as InCrate on InCrate.Item_code=TSPL_ITEM_MASTER.Item_Code  
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Packet_Type='Y' ) as InPouch on InPouch.Item_code=TSPL_ITEM_MASTER.Item_Code  
						
--						left outer join (select Document_Code, STRING_AGG(Batch_No,CHAR(10)) as BatchNo , STRING_AGG(CAST(Qty AS INT), CHAR(10)) as Batch_Qty
--,Item_Code,UOM
--                from(
--SELECT TSPL_BATCH_ITEM.Document_Code, Batch_No, Qty, Parent_Line_No,Item_Code,UOM FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='FS-SH'
--)x group by Document_Code,Parent_Line_No,Item_Code,UOM         
--)TabBatch
--On TabBatch.Document_Code= TSPL_SD_SHIPMENT_HEAD.Document_Code And  TabBatch.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  and TabBatch.UOM=TSPL_SD_SHIPMENT_DETAIL.Unit_code  
                       
                        Left outer join TSPL_COMPANY_MASTER on  TSPL_COMPANY_MASTER.Comp_Code1 = 'BKN'
                    where TSPL_SD_SHIPMENT_HEAD.Document_Code in(Select DISTINCT(TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE) from " & tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL & " 
left outer join " & tbl_TSPL_SD_SHIPMENT_DETAIL & " ON  TSPL_SD_SHIPMENT_DETAIL.PK_ID=TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID where GPCode='" & StrCode & "'))xxx
                        Left OUTER JOIN (
Select 1 As COL1, 1 As COL2,  'ORIGINAL COPY' as CopyType1  "
        If Not clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            Qry += "UNION Select 1 as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 
UNION Select 1 as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 
UNION Select 1 as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1 "
        End If
        Qry += ") YYY ON YYY.COL1=XXX.CopyType
left outer join (select Document_Code, STRING_AGG(Batch_No,CHAR(10)) as BatchNo , STRING_AGG(CAST(Qty AS INT), CHAR(10)) as Batch_Qty
,Item_Code,UOM
                from(
SELECT TSPL_BATCH_ITEM.Document_Code, Batch_No, Qty, Parent_Line_No,Item_Code,UOM FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='FS-SH'
)x group by Document_Code,Parent_Line_No,Item_Code,UOM         
)TabBatch
On TabBatch.Document_Code= XXX.Document_Code And TabBatch.Item_Code=XXX.Item_Code  and TabBatch.UOM=XXX.Unit_code
left join(select Gpdate,Supply_Date from " & tbl_TSPL_DAIRYSALE_GATEPASS_MASTER & " where GPCode='" & StrCode & "') GP on GP.Supply_Date=xxx.Supply_Date
group by XXX.Item_Code 
ORDER BY max(YYY.COL2 ),max(orderby1) ,max(orderby2) ,max(orderby3),max(Item_Desc)    "
        Return Qry
    End Function

    '============================Changes by preeti gupta [09/01/2017],[BHA/02/08/18-000212]
    Private Function GetAttachQry(ByVal StrCode As String, ByVal isDepartmentRoute As Boolean)
        Return GetAttachQry(StrCode, isDepartmentRoute, False)
    End Function

    Private Function GetAttachQry(ByVal StrCode As String, ByVal isDepartmentRoute As Boolean, ByVal isCancel As Boolean) As String
        ''richa remove ceiling from crate qty 15 Nov,2019
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_MASTER As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_DETAIL As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_HEAD As String = Nothing
        Dim tbl_TSPL_BOOKING_MATSER As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_BOOKING_DETAIL As String = Nothing

        If isCancel Then
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD_Cancel_Data As TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER_Cancel_Data As TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
            tbl_TSPL_SD_SHIPMENT_BOOKING_DETAIL = " TSPL_SD_SHIPMENT_BOOKING_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_BOOKING_DETAIL"
        Else
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = "  TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
            tbl_TSPL_SD_SHIPMENT_BOOKING_DETAIL = "  TSPL_SD_SHIPMENT_BOOKING_DETAIL"
        End If


        Dim Qry As String = ""
        If isDepartmentRoute Then
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Qry = "  Select
 	  CAST(case when (Round(Qtycrate,0)*ConverLTR) > MilkQuantityltr and MilkQuantityltr > 0  then 
  ( MilkQuantityltr - (Round(Qtycrate,0)*ConverLTR))  when (Round(Qtycrate,0)*ConverLTR) < MilkQuantityltr and MilkQuantityltr > 0 then (MilkQuantityltr-(Round(Qtycrate,0)*ConverLTR)  ) else 0 end AS DECIMAL(18,2)
) as LooseLtr,
     Cast(case when (Round(Qtycrate,0)*ConverKG) > MilkQuantityKG and MilkQuantityKG > 0   then 
  ( MilkQuantityKG - (Round(Qtycrate,0)*ConverKG))   when (Round(Qtycrate,0)*ConverKG) < MilkQuantityKG and MilkQuantityKG > 0 then  (MilkQuantityKG-(Round(Qtycrate,0)*ConverKG) ) else 0  end as Decimal(18,2)) as Loosekg,
		
		* from ( Select Cast((Conversion_FactorCrt/CFinLTR) as decimal(18,6)  ) as ConverLTR,Cast((Conversion_FactorCrt/CFinKG) as decimal(18,6)  ) as ConverKG,CASE WHEN CFinLTR > 0 AND Conversion_FactorCrt > 0 THEN
        CAST(MilkQuantityltr * CFinLTR AS DECIMAL(18,0)) else 0 end as LTRQty,
		case when CFinKG > 0 and Conversion_FactorCrt > 0 then
		CAST(MilkQuantityKG * CFinKG as decimal(18,0)) else 0 end as KGQty,* from ( "
            End If
            Qry += " Select  ISNULL(CAST(qty / Conver_Factr AS INT), 0) AS Bulk_uom_Qty,'" & clsCommon.myCstr(IIf(isCancel, "Y", "N")) & "' As isCancelled,'" & objCommonVar.CurrentUserCode & "' as UserName,  Case When CFinPouch > 0 and Final.StokingUOM='LTR' Then ( ( ((Final.Box_Crate_Qty * Final.Conversion_Factor)/CFinPouch) + Final.Pouch_Qty )/ CFinLTR ) Else (Case When CFinPouch > 0  Then ( ( Final.Box_Crate_Qty * Final.Conversion_Factor + Final.Pouch_Qty )/ CFinPouch ) Else 0 End) End AS 'NoOfPouch',
Case When CFinLTR > 0 and Final.StokingUOM='LTR' or Final.StokingUOM = 'Pouch' Then ( final.LTR_Qty ) Else (Case When CFinLTR > 0 Then ( ( ( Final.Box_Crate_Qty * Final.Conversion_Factor + Final.Pouch_Qty ) )/ CFinLTR ) ELSE 0 end) End AS 'MilkQuantity',
Case When CFinLTR > 0 Then ( final.LTR_Qty ) Else 0 End AS 'MilkQuantityltr',
CAST( ( Final.Box_Crate_Qty * Final.Conversion_Factor ) / CFinKG AS DECIMAL(10, 2) ) AS 'MilkQuantityKG',
Case When Final.Column_Crate > 0 Then Cast( ( Cast((Final.qty*Final.CFinLTR/Conversion_FactorCrt) as int) / Final.Column_Crate ) AS int ) Else 0 End AS 'CrateLine',
Case When Column_Crate > 0 Then ( Cast((Final.qty*Final.CFinLTR/Conversion_FactorCrt) as int)-( Column_Crate * (Case When Final.Column_Crate > 0 Then Cast( ( Cast((Final.qty*Final.CFinLTR/Conversion_FactorCrt) as int) / Final.Column_Crate ) AS int ) Else 0 End)) ) Else 0 End AS 'LooseCrate',
(((Final.qty*Final.CFinLTR)-(Conversion_FactorCrt* Cast((Final.qty*Final.CFinLTR/Conversion_FactorCrt) as int)))) as LoosePouch,CAST( CASE WHEN qty > 0 AND Conversion_FactorCrt IS NOT NULL THEN FLOOR((CAST(Pouch_Qty AS decimal(18,2)) * CFinPouch) / Conversion_FactorCrt) ELSE 0 END AS decimal(18,2)) AS CrateQtydd,
CASE WHEN Unit_Code='POUCH' then qty*CFinPouch / Conversion_FactorCrt WHEN Unit_Code='LTR' then qty*CFinLTR / Conversion_FactorCrt WHEN Unit_Code='KG' then qty*CFinKG / Conversion_FactorCrt WHEN Unit_Code='CRATE' then qty*Conversion_FactorCrt / Conversion_FactorCrt WHEN Unit_Code='BOX' then qty*CFinBOX / Conversion_FactorCrt ELSE 0 END AS QtyCrate ,
Final.*,
tbl_Brand.Brand, tbl_Brand.BRANDDESC, TSPL_COMPANY_MASTER.Logo_Img,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " 	  FINAL.TAX1,FINAL.TAX1_Amt,FINAL.TAX2,FINAL.TAX2_Amt,FINAL.TAX3,FINAL.TAX3_Amt,FINAL.TAX3,FINAL.TAX4_Amt
						   ,FINAL.TAX5,FINAL.TAX5_Amt,FINAL.TAX6,FINAL.TAX7,FINAL.TAX7_Amt,FINAL.TAX8,Final.TAX8_Amt, "
            End If
            Qry += "    TSPL_COMPANY_MASTER.Logo_Img2,
             case when isnull(Final.CFinLTR,0)=0 then Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinKG) else Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinLTR) end as item_Cost"
        Else
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Qry = "  Select
 	  CAST(case when (Round(Qtycrate,0)*ConverLTR) > MilkQuantityltr and MilkQuantityltr > 0  then 
  ( MilkQuantityltr - (Round(Qtycrate,0)*ConverLTR))  when (Round(Qtycrate,0)*ConverLTR) < MilkQuantityltr and MilkQuantityltr > 0 then (MilkQuantityltr-(Round(Qtycrate,0)*ConverLTR)  ) else 0 end AS DECIMAL(18,2)
) as LooseLtr,
     Cast(case when (Round(Qtycrate,0)*ConverKG) > MilkQuantityKG and MilkQuantityKG > 0   then 
  ( MilkQuantityKG - (Round(Qtycrate,0)*ConverKG))   when (Round(Qtycrate,0)*ConverKG) < MilkQuantityKG and MilkQuantityKG > 0 then  (MilkQuantityKG-(Round(Qtycrate,0)*ConverKG) ) else 0  end as Decimal(18,2)) as Loosekg,
		
		* from ( Select Cast((Conversion_FactorCrt/CFinLTR) as decimal(18,6)  ) as ConverLTR,Cast((Conversion_FactorCrt/CFinKG) as decimal(18,6)  ) as ConverKG,CASE WHEN CFinLTR > 0 AND Conversion_FactorCrt > 0 THEN
        CAST(MilkQuantityltr * CFinLTR AS DECIMAL(18,0)) else 0 end as LTRQty,
		case when CFinKG > 0 and Conversion_FactorCrt > 0 then
		CAST(MilkQuantityKG * CFinKG as decimal(18,0)) else 0 end as KGQty,* from ( "
            End If

            Qry += " Select ISNULL(CAST(qty / Conver_Factr AS INT), 0) AS Bulk_uom_Qty,'" & clsCommon.myCstr(IIf(isCancel, "Y", "N")) & "' As isCancelled,'" + objCommonVar.CurrentUserCode + "' as UserName, "
            Qry += " Case When CFinPouch > 0 and Final.StokingUOM='LTR' Then ( ( ((Final.Crate_Qty * Final.Conversion_Factor)/CFinPouch) + Final.Pouch_Qty )/ CFinLTR ) Else (Case When CFinPouch > 0  Then ( ( Final.Crate_Qty * Final.Conversion_Factor + Final.Pouch_Qty )" + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal, "+(Final.LTR_Qty*Final.CFinLTR/CFinPouch)", "") + " )/ CFinPouch Else 0 End) End AS 'NoOfPouch',"
            Qry += " Case When CFinLTR > 0 and Final.StokingUOM='LTR' Then ( ( ( ((Final.Crate_Qty * Final.Conversion_Factor)/CFinPouch) + Final.Pouch_Qty ) )* CFinPouch ) Else (Case When CFinLTR > 0 Then ( ( ( Final.Crate_Qty * Final.Conversion_Factor + Final.Pouch_Qty )" + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal, "+(Final.LTR_Qty*Final.CFinLTR/CFinPouch)", "") + " )/ CFinLTR ) ELSE 0 end) End AS 'MilkQuantity', "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                Qry += "NoCrateIssue ,Demand_UniqueID, "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Qry += " Case When CFinLTR > 0 or Final.StokingUOM='LTR' then Final.LTR_Qty  else (case when CFinLtr>0 then ( ( ( Final.Crate_Qty * Final.Conversion_Factor )+( Final.Pouch_Qty * Final.CFinPouch ) +(Final.LTR_Qty*Final.CFinLTR/CFinPouch) )/ CFinLTR ) else 0 end)  End AS 'MilkQuantityltr',"
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " Case When CFinLTR > 0 Then ( ( ( Final.Crate_Qty * Final.Conversion_Factor )/Final.PrintUOMConv +( Final.Pouch_Qty * Final.CFinPouch )/ Final.PrintUOMConv  ) ) Else 0 End AS 'MilkQuantityltr',"
            Else
                Qry += " Case When CFinLTR > 0 Then ( ( ( Final.Crate_Qty * Final.Conversion_Factor )+( Final.Pouch_Qty * Final.CFinPouch ) " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal, "+(Final.LTR_Qty*Final.CFinLTR/CFinPouch)", "") + " )/ CFinLTR ) Else 0 End AS 'MilkQuantityltr',"
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += "  CAST( ( Final.Box_Crate_Qty * Final.Conversion_Factor ) / Final.PrintUOMConv AS DECIMAL(10, 3) ) AS 'MilkQuantityKG',"
            Else
                Qry += "  CAST( ( Final.Box_Crate_Qty * Final.Conversion_Factor ) / CFinKG AS DECIMAL(10, 3) ) AS 'MilkQuantityKG',"
            End If


            Qry += " Case When Final.Column_Crate > 0 Then Cast( ( Final.Crate_Qty / Final.Column_Crate ) AS int ) Else 0 End AS 'CrateLine', Case When Column_Crate > 0 Then ( crate_qty-( Column_Crate * (Case When Final.Column_Crate > 0 Then Cast( ( Final.Crate_Qty / Final.Column_Crate ) AS int ) Else 0 End)) ) Else 0 End AS 'LooseCrate'," + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, "Pouch_Qty as LoosePouch, ", "(Pouch_Qty -  ( CAST( CASE WHEN qty > 0 AND Conversion_FactorCrt IS NOT NULL THEN FLOOR((CAST(Pouch_Qty AS decimal(18,2)) * CFinPouch) / Conversion_FactorCrt) ELSE 0 END AS decimal(18,2))) * Conversion_FactorCrt/CFinPouch) As LoosePouch,") + " 
                           CAST( CASE WHEN qty > 0 AND Conversion_FactorCrt IS NOT NULL THEN FLOOR((CAST(Pouch_Qty AS decimal(18,2)) * CFinPouch) / Conversion_FactorCrt) ELSE 0 END AS decimal(18,2)) AS CrateQtydd,
                           CASE WHEN Unit_Code='POUCH' then qty*CFinPouch / Conversion_FactorCrt WHEN Unit_Code='LTR' then qty*CFinLTR / Conversion_FactorCrt WHEN Unit_Code='KG' then qty*CFinKG / Conversion_FactorCrt WHEN Unit_Code='CRATE' then qty*Conversion_FactorCrt / Conversion_FactorCrt WHEN Unit_Code='BOX' then qty*CFinBOX / Conversion_FactorCrt ELSE 0 END AS QtyCrate , Final.*, tbl_Brand.Brand, tbl_Brand.BRANDDESC, TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2, case when isnull(Final.CFinLTR,0)=0 then Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinKG) else Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinLTR) end as item_Cost"
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += " ,ROW_NUMBER() OVER (ORDER BY [item_desc]) AS SerialNumber "
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            Qry += "  ,(qty * Conversion_Factor/CFinPrintUOM ) as PrintUOM,case when UPPER(Unit_code)='CRATE' then qty else 0 end as TotalCrate,case when UPPER(Unit_code)='BOX' then qty else 0 end as TotalBOX,    case when UPPER(Unit_code)='CFC' then qty else 0 end as TotalCFC "
        End If
        Qry += "  FROM 
                   ( Select  max(Supply_Date)as Supply_Date,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            Qry += " max(Conversion_FactorCrt) as Conversion_FactorCrt , "
        Else
            Qry += " MAX(PrintUOMConv)PrintUOMConv,max(Conversion_FactorCrt) as Conversion_FactorCrt, "
        End If
        Qry += " Max(Distributor) Distributor,MAX(DistAddress)DistAddress " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ", max(Unit_Code) as Unit_code,max(CFinPrintUOM) as CFinPrintUOM ", "") + ", Max(CFinPouch) CFinPouch, Max(CFinLTR) CFinLTR, Max(CFinBOX) CFinBOX, Max(CFinKG) CFinKG, Max(Conversion_Factor) Conversion_Factor, Max(AgainstTransferNo) AgainstTransferNo, Max(Comp_Code) Comp_Code, Sum( Qty * case when Unit_Code = 'Crate' then 1 else 0 end ) Crate_Qty, Sum( Qty * case when Unit_Code = 'Pouch' then 1 else 0 end ) Pouch_Qty,Sum(Qty * case when Unit_Code = 'LTR' then 1 else 0 end) LTR_Qty, Sum(Box_Crate_Qty) Box_Crate_Qty, Max(Insurance_No) Insurance_No, Max(Insurance_Comp_Name) Insurance_Comp_Name, Max(comp_name) comp_name,Max(ISO_No) ISO_No"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal Then
            Qry += " ,Max(unit_code) unit_code "
        End If
        Qry += ",CASE WHEN max(unit_code) = 'crate' OR max(unit_code) = 'pouch' THEN 'Crate/Pouch' ELSE max(unit_code) END AS unit_code_result,
                 Sum(qty) qty, Max(Comp_Address) Comp_Address, Max(Loc_add) Loc_add, Max(Route_No) Route_No, Sum(Totalcrate) Totalcrate, Sum(TotalCan) TotalCan, Max(Route_Desc) Route_Desc, Max(GPCode) GPCode, Max(GPDate) GPDate, Max(GPTime) GPTime, Max(vehicle_id) vehicle_id, Max(VehicleDesc) VehicleDesc, Max(location_code) location_code, Max(Location_desc) Location_desc, Max(transporter) transporter, Max(remarks) remarks, Max(comments) comments, Max(post) post, Item_code, Max(item_desc) item_desc, Max(short_description) short_description, Max(sku_seq) sku_seq, Max(TranporterNameFromMaster) TranporterNameFromMaster, Max(HSN_Code) HSN_Code, Max(Salesman) Salesman, Max(Column_Crate) as  Column_Crate, Max(Area_Code) Area_Code, Max(Zone_Code) Zone_Code, Max(ShiftType) ShiftType, Max(GSTReg_No) GSTReg_No, Max(Loading_Slip) Loading_Slip, Max(DispatchDate) DispatchDate, Max(GatePass_Date) GatePass_Date, Max(Amount) Amount, Max(AmountWithoutTax) as AmountWithoutTax, Max(Margin) Margin, Max(SecurityAmt) as SecurityAmt,max(Total_TCS_Amt)Total_TCS_Amt, max(Dist_Commission_Ratewithtax) Dist_Commission_Ratewithtax,max(zonecode)zonecode ,max(RoundOffAmt) as RoundOffAmt " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ", max(ItemCost) as UnitRate", "") + ", Max(Driver_Name) Driver_Name, Max(Driver_ContactNo) Driver_ContactNo,max(UOM_Code) as StokingUOM "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " ,max(ActualRate)ActualRate "
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            Qry += " ,MAX(Demand_UniqueID)Demand_UniqueID, max(NoCrateIssue)NoCrateIssue "
        End If

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += "  ,'' as Demand_No,MAX(TAX1)TAX1
				 ,MAX(TAX1_Amt)TAX1_Amt
,				
MAX(TAX2)TAX2,MAX(TAX2_Amt)TAX2_Amt,
MAX(TAX3)TAX3,MAX(TAX3_Amt)TAX3_Amt,
MAX(TAX4)TAX4,MAX(TAX4_Amt)TAX4_Amt,
MAX(TAX5)TAX5,MAX(TAX5_Amt)TAX5_Amt,
MAX(TAX6)TAX6,MAX(TAX6_Amt)TAX6_Amt,
MAX(TAX7)TAX7,MAX(TAX7_Amt)TAX7_Amt,
MAX(TAX8)TAX8,MAX(TAX8_Amt)TAX8_Amt "
        End If
        Qry += " ,max(Conver_Factr)Conver_Factr,max(Bulk_UOM_Code)Bulk_UOM_Code,MAX(Sale_Invoice_No)Sale_Invoice_No,MAX(Booth)Booth,MAX(Sub_Route_Code)Sub_Route_Code,MAX(Sub_Route_Name)Sub_Route_Name,max(Vehicle_Type)Vehicle_Type  FROM
                   ( select   Bulk_UOM.UOM_Code as Bulk_UOM_Code,Bulk_UOM.Conversion_Factor AS Conver_Factr,FORMAT( TSPL_DAIRYSALE_GATEPASS_MASTER.Supply_Date, 'dd/MM/yyyy' ) as Supply_Date,tspl_item_uom_detail.Conversion_Factor As PrintUOMConv,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            Qry += "NoCrateIssue, Demand_UniqueID, "
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            Qry += " ItemConversionCrate.Conversion_Factor as Conversion_FactorCrt , "
        Else
            Qry += " ItemConversionCrate.Conversion_Factor as Conversion_FactorCrt , "
        End If
        Qry += " TSPL_DAIRYSALE_GATEPASS_MASTER.DistributorName As 'Distributor',DistAddress, ItemConversionInPouch.Conversion_Factor As 'CFinPouch', ItemConversionInLTR.Conversion_Factor AS 'CFinLTR', ItemConversionInKG.Conversion_Factor AS 'CFinKG', ItemConversionInBox.Conversion_Factor AS 'CFinBOX',"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            Qry += " ItemConversionInPrintUOM.Conversion_Factor as CFinPrintUOM,"
        End If
        Qry += "CurrentUnit.Conversion_Factor,StockUnit.UOM_Code, isnull( TSPL_DAIRYSALE_GATEPASS_MASTER.AgainstTransferNo, '' ) as AgainstTransferNo, TSPL_COMPANY_MASTER.Comp_Code, CASE WHEN TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Box' THEN convert( decimal(18, 2), ( TSPL_DAIRYSALE_GATEPASS_DETAIL.qty / CrateUnit.conversion_factor )* StockUnit.conversion_factor * CurrentUnit.conversion_factor ) ELSE 0 END as Crate_Qty, CASE WHEN TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Crate' and TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Pouch'" & IIf(isDepartmentRoute, "", " and TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'LTR' ") & "  THEN convert( decimal(18, 3), ( TSPL_DAIRYSALE_GATEPASS_DETAIL.qty / CurrentUnit.conversion_factor )* StockUnit.conversion_factor * CurrentUnit.conversion_factor ) else 0 end as Box_Crate_Qty, TSPL_COMPANY_MASTER.Insurance_No, TSPL_COMPANY_MASTER.Insurance_Comp_Name, TSPL_COMPANY_MASTER.comp_name,TSPL_COMPANY_MASTER.ISO_No, TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code, TSPL_DAIRYSALE_GATEPASS_DETAIL.qty, TSPL_COMPANY_MASTER.add1 + case when len(TSPL_COMPANY_MASTER.add2)> 0 then ', ' + TSPL_COMPANY_MASTER.add2 else '' end + case when LEN( isnull(TSPL_COMPANY_MASTER.Add3, '') )> 0 then ', ' + isnull(TSPL_COMPANY_MASTER.Add3, '') else ' ' end as Comp_Address, tspl_location_master.add1 + case when len(tspl_location_master.add2)> 0 then ', ' + tspl_location_master.add2 else '' end + case when LEN( isnull(tspl_location_master.Add3, '') )> 0 then ', ' + isnull(tspl_location_master.Add3, '') else ' ' end as Loc_add, TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No, TSPL_DAIRYSALE_GATEPASS_MASTER.Totalcrate, TSPL_DAIRYSALE_GATEPASS_MASTER.TotalCan, tspl_route_master.Route_Desc, TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode, convert( varchar, TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, 103 ) as GPDate, FORMAT( TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, 'hh:mm tt' ) as GPTime, TSPL_DAIRYSALE_GATEPASS_MASTER.GatePass_Date, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id AS vehicle_id, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number as VehicleDesc, TSPL_DAIRYSALE_GATEPASS_MASTER.location_code, tspl_location_master.Location_desc, TSPL_DAIRYSALE_GATEPASS_MASTER.transporter, TSPL_DAIRYSALE_GATEPASS_MASTER.remarks, TSPL_DAIRYSALE_GATEPASS_MASTER.comments, TSPL_DAIRYSALE_GATEPASS_MASTER.post, case when isnull(TSPL_DAIRYSALE_GATEPASS_DETAIL.Scheme_Item,'') = 'Y' then 
TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code + '-Scheme' else TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code end as Item_code, case when isnull(TSPL_DAIRYSALE_GATEPASS_DETAIL.Scheme_Item,'') = 'Y' then tspl_item_master.item_desc + '-Scheme' else tspl_item_master.item_desc end as item_desc, tspl_item_master.short_description, tspl_item_master.sku_seq, TSPL_TRANSPORT_MASTER.Transporter_Name as TranporterNameFromMaster, TSPL_DAIRYSALE_GATEPASS_DETAIL.HSN_Code, TSPL_DAIRYSALE_GATEPASS_MASTER.Salesman, tspl_vehicle_master.Column_Crate, TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No AS Area_Code, TSPL_CUSTOMER_MASTER.Zone_Code, TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType, tspl_company_master.GSTReg_No, TSPL_DAIRYSALE_GATEPASS_MASTER.Loading_Slip, ( Select Max(Document_Date) from TSPL_SD_SHIPMENT_HEAD where GPCode='" + StrCode + "') AS 'DispatchDate',case when TSPL_DAIRYSALE_GATEPASS_DETAIL.Scheme_Item = 'Y' then 0 else  xyz.Amount end as Amount, xyz.AmountWithoutTax,xyz.Margin, xyz.SecurityAmt, xyz.Dist_Commission_Ratewithtax,xyz.RoundOffAmt,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += " 
XYZ.TAX1,XYZ.TAX1_Amt,
XYZ.TAX2,XYZ.TAX2_Amt,
XYZ.TAX3,XYZ.TAX3_Amt,
XYZ.TAX4,XYZ.TAX4_Amt,
XYZ.TAX5,XYZ.TAX5_Amt,
XYZ.TAX6,XYZ.TAX6_Amt,
XYZ.TAX7,XYZ.TAX7_Amt,
XYZ.TAX8,XYZ.TAX8_Amt, "
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " xyz.ActualRate, "
        End If
        Qry += "case when TSPL_DAIRYSALE_GATEPASS_DETAIL.Scheme_Item = 'Y' then 0 else xyz.Total_TCS_Amt end as Total_TCS_Amt,xyz.Zone_Code as Zonecode" + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ", xyz.ItemCost", "") + ",TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_Name,TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_ContactNo,xyz.Sale_Invoice_No,xyz.Booth,xyz.Sub_Route_Code,xyz.Sub_Route_Name,xyz.Vehicle_Type    from " & tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL & " " &
                   " left outer join " & tbl_TSPL_DAIRYSALE_GATEPASS_MASTER & "  on TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode " &
                   " left outer join tspl_vehicle_master on tspl_vehicle_master.Vehicle_id=TSPL_DAIRYSALE_GATEPASS_MASTER.vehicle_id " &
                   " left outer join tspl_location_master on tspl_location_master.location_code=TSPL_DAIRYSALE_GATEPASS_MASTER.location_code " &
                   " left outer join tspl_company_master on tspl_company_master.comp_code=TSPL_DAIRYSALE_GATEPASS_MASTER.comp_code " &
                   " left outer join tspl_item_master on tspl_item_master.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code " &
                   " left outer join  TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id " &
                   " left outer join  tspl_route_master on tspl_route_master.Route_No=TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No " &
             " left join tspl_item_uom_detail StockUnit on StockUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code	and StockUnit.stocking_unit='Y'  " &
                   " left join tspl_item_uom_detail CurrentUnit on CurrentUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and 	CurrentUnit.uom_code=	TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code " &
                   " left join tspl_item_uom_detail CrateUnit on CrateUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code  and 	CrateUnit.uom_code=	'Crate' " &
                   " left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and tspl_item_uom_detail.Print_UOM=1" &
                   " left join tspl_item_uom_detail Bulk_UOM  on Bulk_UOM.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and Bulk_UOM.Bulk_UOM=1 " &
                   " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='Pouch') as ItemConversionInPouch on ItemConversionInPouch.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code
                    left join ( select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'Crate' ) as ItemConversionCrate on ItemConversionCrate.Item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code 
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='LTR') as ItemConversionInLTR on ItemConversionInLTR.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code 
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='Box') as ItemConversionInBox on ItemConversionInBox.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='KG') as ItemConversionInKG on ItemConversionInKG.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where Print_UOM = 1 ) as ItemConversionInPrintUOM on ItemConversionInPrintUOM.Item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code "
        End If

        '--" left join tspl_item_uom_detail gpUnit on gpUnit.item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and CrateUnit.uom_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            'Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=tspl_vehicle_master.Transport_Id  left join tspl_item_uom_detail gpUnit on gpUnit.item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and CrateUnit.uom_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code "
            Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=tspl_vehicle_master.Transport_Id   "
        Else
            Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=tspl_vehicle_master.Transport_Id  "
        End If
        Qry += " left outer join ("
        '             ((select Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty,0))>0 Then  Sum(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt/TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GP_Qty) Else 0 End As Amount,  Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty, 0))> 0 Then Sum(TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount) Else 0 End As AmountWithoutTax, Sum(IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0)/TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GP_Qty) AS Margin,sum( IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0))as SecurityAmt,TSPL_SD_SHIPMENT_DETAIL.Item_Code,max(TSPL_SD_SHIPMENT_DETAIL.Unit_code) as Unit_code,max(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate) as Dist_Commission_Ratewithtax,max(TSPL_SD_SHIPMENT_HEAD.RoundOffAmount) as RoundOffAmt " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ", max(TSPL_SD_SHIPMENT_DETAIL.Item_Cost) as ItemCost", "") + " from TSPL_SD_SHIPMENT_DETAIL   
        '               Left Outer Join TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
        'Left Outer Join TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL On TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID=TSPL_SD_SHIPMENT_DETAIL.PK_ID
        '             WHERE  TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode = '" + StrCode + "'
        '             Group By  TSPL_SD_SHIPMENT_DETAIL.Item_Code)
        '             Union All

        Qry += "(select Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty,0))>0 Then  Sum(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt) Else 0 End As Amount,  Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty, 0))> 0 Then Sum(TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount) Else 0 End As AmountWithoutTax, 
                     Sum(IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0)) AS Margin,sum( IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0))as SecurityAmt,
                     TSPL_SD_SHIPMENT_DETAIL.Item_Code ,max(TSPL_SD_SHIPMENT_DETAIL.Unit_code) as Unit_code ,max(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate) as Dist_Commission_Ratewithtax 
       ,case  WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX1) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX2) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX3) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX4) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX5) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX6) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt)
        WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX7) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt)
		WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX8) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt)
		WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX9) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt)
		WHEN max(TSPL_SD_SHIPMENT_DETAIL.TAX10) = 'TCS' THEN sum(TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt)
        ELSE 0 END AS Total_TCS_Amt 
         ,max(TSPL_SD_SHIPMENT_HEAD.RoundOffAmount) as RoundOffAmt,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += "  max(TSPL_SD_SHIPMENT_DETAIL.ActualRate)ActualRate  ,"
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += "  MAX(TSPL_SD_SHIPMENT_DETAIL.TAX1)TAX1,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt)TAX1_Amt,
		 MAX(TSPL_SD_SHIPMENT_DETAIL.TAX2_AMT)TAX2_AMT, MAX(TSPL_SD_SHIPMENT_DETAIL.TAX3_AMT)TAX3_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX4_AMT)TAX4_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX5_AMT)TAX5_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX6_AMT)TAX6_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX7_AMT)TAX7_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX8_AMT)TAX8_AMT,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX2)TAX2,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX3)TAX3,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX4)TAX4,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX5)TAX5,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX6)TAX6,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX7)TAX7,MAX(TSPL_SD_SHIPMENT_DETAIL.TAX8)TAX8, "
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            Qry += "max(NoCrateIssue)NoCrateIssue,"
        End If
        Qry += " max(TSPL_CUSTOMER_MASTER.Zone_Code)Zone_Code " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ", max(TSPL_SD_SHIPMENT_DETAIL.Item_Cost) as ItemCost", "") + ",Max(Concat(TSPL_CUSTOMER_MASTER.Add1,TSPL_CUSTOMER_MASTER.Add2)) As 'DistAddress',Max(TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No) As Sale_Invoice_No,MAX(TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booth_Code) As Booth,MAX(TSPL_CUSTOMER_MASTER_SUB_Route.Sub_Route_Code)Sub_Route_Code,Max(TSPL_SUB_ROUTE_MASTER.Name) As Sub_Route_Name,max(TSPL_SD_SHIPMENT_HEAD.Vehicle_Type) Vehicle_Type  from " & tbl_TSPL_SD_SHIPMENT_DETAIL & "   
 
                     Left Outer Join " & tbl_TSPL_SD_SHIPMENT_HEAD & " ON TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            Qry += " left outer join " & tbl_TSPL_BOOKING_MATSER & " on TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "

        End If
        Qry += " Left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
Left Outer Join " & tbl_TSPL_SD_SHIPMENT_BOOKING_DETAIL & " On TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.Document_Code And TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code= TSPL_SD_SHIPMENT_DETAIL.Item_Code
Left Outer Join TSPL_CUSTOMER_MASTER As TSPL_CUSTOMER_MASTER_SUB_Route On TSPL_CUSTOMER_MASTER_SUB_Route.Cust_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booth_Code
Left Outer Join TSPL_SUB_ROUTE_MASTER On TSPL_SUB_ROUTE_MASTER.Code=TSPL_CUSTOMER_MASTER_SUB_Route.Sub_Route_Code
                     WHERE  --TSPL_SD_SHIPMENT_HEAD.GPCode = '" + StrCode + "'
                      TSPL_SD_SHIPMENT_DETAIL.PK_ID in(select PK_ID from " & tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL & " where GPCode = '" + StrCode + "')
                     Group By  TSPL_SD_SHIPMENT_DETAIL.Item_Code "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += ",TSPL_SD_SHIPMENT_DETAIL.Unit_code "
            Qry += "))xyz ON xyz.Item_Code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code And xyz.Unit_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code "
        Else
            Qry += "))xyz ON xyz.Item_Code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code "
        End If

        Qry += " where 2=2 " &
                    "  and  TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode = '" + StrCode + "' " &
        " )As Main Group by Item_code "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += ",Unit_Code"
        End If
        Qry += " )AS Final left outer join  ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND], " &
        " max([DESCRP]) as [DESCRP],max([PACK]) as [PACK], max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA], " &
                    " max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW], max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC]) " &
                    " as [BRANDDESC],max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC], max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC]," &
                    " max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC], max([SCRAPDESC]) as [SCRAPDESC] " &
                    " from ( select * from (   select TSPL_ITEM_MASTER.Item_Code" + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code ", "") + ",TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   ," &
                    " TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER    left outer join TSPL_ITEM_MASTER_CATEGORY on " &
                    " TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code  left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and " &
                    " TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=2 )xx  Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE], " &
                    " [CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC], [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC]," &
                    " [CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, "", "") + " )  as tbl_Brand on tbl_Brand.Item_Code=Final.Item_Code   left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =Final.Comp_Code "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
            Qry += " )XX )YX order by YX.sku_seq "
        Else
            Qry += " order by Final.sku_seq "
        End If
        '" order by Final.sku_seq"
        '' Group by UnitCode removed by Vinod (as per Instructed by Prabhat Sirohi) date:- 22-Aug-2024
        '" )As Main Group by Item_code" + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code", "") + " )As Final left outer join  ( Select Item_Code,max([CATEGORY RM]) As [CATEGORY RM],max([BRAND]) As [BRAND],max([SUB BRAND]) As [SUB BRAND], " &
        '
        Return Qry
    End Function

    Public Function getUDPattachqry(ByVal GPCode As String, ByVal arrLocation As ArrayList, ByVal arrRoute As ArrayList, ByVal arrVehicle As ArrayList, ByVal ShiftType As String) As String
        Return getUDPattachqry(GPCode, arrLocation, arrRoute, arrVehicle, ShiftType, False)
    End Function

    Public Function getUDPattachqry(ByVal GPCode As String, ByVal arrLocation As ArrayList, ByVal arrRoute As ArrayList, ByVal arrVehicle As ArrayList, ByVal ShiftType As String, ByVal isCancel As String) As String
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_MASTER As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_DETAIL As String = Nothing
        Dim tbl_TSPL_SD_SHIPMENT_HEAD As String = Nothing
        Dim tbl_TSPL_BOOKING_MATSER As String = Nothing
        Dim tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL As String = Nothing
        If isCancel Then
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD_Cancel_Data As TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER_Cancel_Data As TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL_Cancel_Data As TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
        Else
            tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL = " TSPL_DAIRYSALE_GATEPASS_Detail "
            tbl_TSPL_DAIRYSALE_GATEPASS_MASTER = " TSPL_DAIRYSALE_GATEPASS_MASTER "
            tbl_TSPL_SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL "
            tbl_TSPL_SD_SHIPMENT_HEAD = "  TSPL_SD_SHIPMENT_HEAD "
            tbl_TSPL_BOOKING_MATSER = " TSPL_BOOKING_MATSER "
            tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL = " TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL"
        End If

        Dim qry As String = ""
        Dim shift As String = ""
        Dim whrcls As String = ""

        If arrLocation IsNot Nothing AndAlso arrLocation.Count > 0 Then
            whrcls += " And TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code In (" & clsCommon.GetMulcallString(arrLocation) & ") "
        End If
        If arrRoute IsNot Nothing AndAlso arrRoute.Count > 0 Then
            whrcls += " And TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No In (" & clsCommon.GetMulcallString(arrRoute) & ") "
        End If
        If arrVehicle IsNot Nothing AndAlso arrVehicle.Count > 0 Then
            whrcls += " And TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id In (" & clsCommon.GetMulcallString(arrVehicle) & ") "
        End If

        If clsCommon.CompairString(clsCommon.myCstr(ShiftType), "M") = CompairStringResult.Equal Then
            whrcls += " And TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType= 'Morning' "
            shift += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type= 'AM' "
        ElseIf clsCommon.CompairString(clsCommon.myCstr(ShiftType), "E") = CompairStringResult.Equal Then
            whrcls += " and TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType= 'Evening' "
            shift += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type= 'PM' "
        End If

        qry = " select '" & clsCommon.myCstr(IIf(isCancel, "Y", "N")) & "' As isCancelled,Type, Comp_Code,case when Is_FreshItem = 1 then isnull(MilkQuantityltr,0) else isnull(MilkQuantityKG,0) end as Qty , isnull(MilkQuantityKG,0) + isnull(MilkQuantityltr,0) as Total_Qty,*,Case When Pouch_Qty>0 Then (Pouch_Qty- (Cast(LooseCrateQtydd As int) * Conversion_FactorCrt)/CFinPouch) Else 0 End  AS LoosePouchQty from ( Select '" & objCommonVar.CurrentUserCode & "' as UserName, Case When CFinPouch > 0 Then ( ( Final.Crate_Qty * Final.Conversion_Factor )/ CFinPouch ) Else 0 End AS 'NoOfPouch', Case When CFinLTR > 0 Then ( ( ( Final.Crate_Qty * Final.Conversion_Factor + Final.Pouch_Qty ) )/ CFinLTR ) Else 0 End AS 'MilkQuantity', Case When CFinLTR > 0 Then ( ( ( Final.Crate_Qty * Final.Conversion_Factor )+( Final.Pouch_Qty * Final.CFinPouch ) )/ CFinLTR ) Else 0 End AS 'MilkQuantityltr', CAST( ( Final.Box_Crate_Qty * Final.Conversion_Factor ) / CFinKG AS DECIMAL(10, 2) ) AS 'MilkQuantityKG', Case When Final.Column_Crate > 0 Then Cast( ( Final.Crate_Qty / Final.Column_Crate ) AS int ) Else 0 End AS 'CrateLine', Case When Column_Crate > 0 Then cast( ( cast(qty as int)% Column_Crate ) as int ) Else 0 End AS 'LooseCrate', Pouch_Qty AS 'LoosePouch', CAST( CASE WHEN qty > 0 THEN CAST(qty AS decimal(18,2)) / Conversion_FactorCrt END AS decimal(18,2) ) AS CrateQtydd, CAST( CASE WHEN Pouch_Qty > 0 THEN Pouch_Qty*CFinPouch / Conversion_FactorCrt Else 0 END AS decimal(18,2)) LooseCrateQtydd, CASE WHEN Unit_Code='POUCH' then qty*CFinPouch / Conversion_FactorCrt WHEN Unit_Code='LTR' then qty*CFinLTR / Conversion_FactorCrt WHEN Unit_Code='KG' then qty*CFinKG / Conversion_FactorCrt WHEN Unit_Code='CRATE' then qty*Conversion_FactorCrt / Conversion_FactorCrt WHEN Unit_Code='BOX' then qty*CFinBOX / Conversion_FactorCrt ELSE 0 END AS QtyCrate , Final.*, tbl_Brand.Brand, tbl_Brand.BRANDDESC, TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2, case when isnull(Final.CFinLTR,0)=0 then Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinKG) else Final.Amount/((Final.qty*Final.Conversion_Factor)/Final.CFinLTR) end as item_Cost FROM 
                   ( Select  max(type)type,  max(Supply_Date)as Supply_Date, "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            qry += " max(Conversion_FactorCrt) as Conversion_FactorCrt , "
        Else
            qry += " max(Conversion_FactorCrt) as Conversion_FactorCrt, "
        End If
        qry += " Max(Distributor) Distributor" & IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code", "") & " , Max(CFinPouch) CFinPouch, Max(CFinLTR) CFinLTR, Max(CFinBOX) CFinBOX, Max(CFinKG) CFinKG, Max(Conversion_Factor) Conversion_Factor,
                 Max(AgainstTransferNo) AgainstTransferNo, Max(Comp_Code) Comp_Code, Sum( Qty * case when Unit_Code = 'Crate' then 1 else 0 end ) Crate_Qty, Sum( Qty * case when Unit_Code = 'Pouch' then 1 else 0 end ) Pouch_Qty, Sum(Box_Crate_Qty) Box_Crate_Qty, Max(Insurance_No) Insurance_No, Max(Insurance_Comp_Name) Insurance_Comp_Name, Max(comp_name) comp_name"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal Then
            qry += " ,Max(unit_code) unit_code "
        End If

        qry += "   ,CASE WHEN max(unit_code) = 'crate' OR max(unit_code) = 'pouch' THEN 'Crate/Pouch' ELSE max(unit_code) END AS unit_code_result,
                   Sum(qty) qty, Max(Comp_Address) Comp_Address, Max(Loc_add) Loc_add, Max(Route_No) Route_No, Sum(Totalcrate) Totalcrate, Sum(TotalCan) TotalCan, Max(Route_Desc) Route_Desc, (GPCode) GPCode, Max(GPDate) GPDate, Max(GPTime) GPTime, Max(vehicle_id) vehicle_id, Max(VehicleDesc) VehicleDesc, Max(location_code) location_code, Max(Location_desc) Location_desc, Max(transporter) transporter, Max(remarks) remarks, Max(comments) comments, Max(post) post, Item_code, Max(item_desc) item_desc, Max(short_description) short_description, Max(sku_seq) sku_seq, Max(TranporterNameFromMaster) TranporterNameFromMaster, Max(HSN_Code) HSN_Code, Max(Salesman) Salesman, Sum(Column_Crate) Column_Crate, Max(Area_Code) Area_Code, Max(Zone_Code) Zone_Code, Max(ShiftType) ShiftType, Max(GSTReg_No) GSTReg_No, Max(Loading_Slip) Loading_Slip, Max(DispatchDate) DispatchDate, Max(GatePass_Date) GatePass_Date, Sum(Amount) Amount, sum(AmountWithoutTax) as AmountWithoutTax, Sum(Margin) Margin, sum(SecurityAmt) as SecurityAmt, max(Dist_Commission_Ratewithtax) Dist_Commission_Ratewithtax ,max(RoundOffAmt) as RoundOffAmt , max(ItemCost) as UnitRate, Max(Driver_Name) Driver_Name, Max(Driver_ContactNo) Driver_ContactNo,max(Is_FreshItem)Is_FreshItem,max(NoCrateIssue)NoCrateIssue,MAX(Demand_UniqueID)Demand_UniqueID FROM
                   ( select   tspl_route_master.type, FORMAT( TSPL_DAIRYSALE_GATEPASS_MASTER.Supply_Date, 'dd/MM/yyyy' ) as Supply_Date,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            qry += " ItemConversionCrate.Conversion_Factor as Conversion_FactorCrt ,  "
        Else
            qry += " ItemConversionCrate.Conversion_Factor as Conversion_FactorCrt , "
        End If

        qry += "     TSPL_DAIRYSALE_GATEPASS_MASTER.DistributorName As 'Distributor', ItemConversionInPouch.Conversion_Factor As 'CFinPouch', ItemConversionInLTR.Conversion_Factor AS 'CFinLTR', ItemConversionInKG.Conversion_Factor AS 'CFinKG', ItemConversionInBox.Conversion_Factor AS 'CFinBOX', CurrentUnit.Conversion_Factor, isnull( TSPL_DAIRYSALE_GATEPASS_MASTER.AgainstTransferNo, '' ) as AgainstTransferNo, TSPL_COMPANY_MASTER.Comp_Code, CASE WHEN TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Box' THEN convert( decimal(18, 2), ( TSPL_DAIRYSALE_GATEPASS_DETAIL.qty / CrateUnit.conversion_factor )* StockUnit.conversion_factor * CurrentUnit.conversion_factor ) ELSE 0 END as Crate_Qty, CASE WHEN TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Crate' and TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code <> 'Pouch' THEN convert( decimal(18, 2), ( TSPL_DAIRYSALE_GATEPASS_DETAIL.qty / CurrentUnit.conversion_factor )* StockUnit.conversion_factor * CurrentUnit.conversion_factor ) else 0 end as Box_Crate_Qty, TSPL_COMPANY_MASTER.Insurance_No, TSPL_COMPANY_MASTER.Insurance_Comp_Name, TSPL_COMPANY_MASTER.comp_name, TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code, TSPL_DAIRYSALE_GATEPASS_DETAIL.qty, TSPL_COMPANY_MASTER.add1 + case when len(TSPL_COMPANY_MASTER.add2)> 0 then ', ' + TSPL_COMPANY_MASTER.add2 else '' end + case when LEN( isnull(TSPL_COMPANY_MASTER.Add3, '') )> 0 then ', ' + isnull(TSPL_COMPANY_MASTER.Add3, '') else ' ' end as Comp_Address, tspl_location_master.add1 + case when len(tspl_location_master.add2)> 0 then ', ' + tspl_location_master.add2 else '' end + case when LEN( isnull(tspl_location_master.Add3, '') )> 0 then ', ' + isnull(tspl_location_master.Add3, '') else ' ' end as Loc_add, TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No, TSPL_DAIRYSALE_GATEPASS_MASTER.Totalcrate, TSPL_DAIRYSALE_GATEPASS_MASTER.TotalCan, tspl_route_master.Route_Desc, TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode, convert( varchar, TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, 103 ) as GPDate, FORMAT( TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, 'hh:mm tt' ) as GPTime, TSPL_DAIRYSALE_GATEPASS_MASTER.GatePass_Date, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id AS vehicle_id, TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number as VehicleDesc, TSPL_DAIRYSALE_GATEPASS_MASTER.location_code, tspl_location_master.Location_desc, TSPL_DAIRYSALE_GATEPASS_MASTER.transporter, TSPL_DAIRYSALE_GATEPASS_MASTER.remarks, TSPL_DAIRYSALE_GATEPASS_MASTER.comments, TSPL_DAIRYSALE_GATEPASS_MASTER.post, TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code, tspl_item_master.item_desc, tspl_item_master.short_description, tspl_item_master.sku_seq, TSPL_TRANSPORT_MASTER.Transporter_Name as TranporterNameFromMaster, TSPL_DAIRYSALE_GATEPASS_DETAIL.HSN_Code, TSPL_DAIRYSALE_GATEPASS_MASTER.Salesman, tspl_vehicle_master.Column_Crate, TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No AS Area_Code, TSPL_CUSTOMER_MASTER.Zone_Code, TSPL_DAIRYSALE_GATEPASS_MASTER.ShiftType, tspl_company_master.GSTReg_No, TSPL_DAIRYSALE_GATEPASS_MASTER.Loading_Slip,
				   ( Select Max(Document_Date) from TSPL_SD_SHIPMENT_HEAD where GPCode in (" & GPCode & ")"
        qry += " ) AS 'DispatchDate'
		           ,xyz.Amount, xyz.AmountWithoutTax,xyz.Margin, xyz.SecurityAmt, xyz.Dist_Commission_Ratewithtax,xyz.RoundOffAmt, xyz.ItemCost,TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_Name,TSPL_DAIRYSALE_GATEPASS_MASTER.Driver_ContactNo,tspl_item_master.Is_FreshItem,NoCrateIssue,Demand_UniqueID 
                   from " & tbl_TSPL_DAIRYSALE_GATEPASS_DETAIL & "
				   left outer join " & tbl_TSPL_DAIRYSALE_GATEPASS_MASTER & " on TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode 
				   left outer join tspl_vehicle_master on tspl_vehicle_master.Vehicle_id=TSPL_DAIRYSALE_GATEPASS_MASTER.vehicle_id 
				   left outer join tspl_location_master on tspl_location_master.location_code=TSPL_DAIRYSALE_GATEPASS_MASTER.location_code 
				   left outer join tspl_company_master on tspl_company_master.comp_code=TSPL_DAIRYSALE_GATEPASS_MASTER.comp_code  
				   left outer join tspl_item_master on tspl_item_master.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_code  
				   left outer join  TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id 
				   left outer join  tspl_route_master on tspl_route_master.Route_No=TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No 
				   left join tspl_item_uom_detail StockUnit on StockUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code	and StockUnit.stocking_unit='Y'   
				   left join tspl_item_uom_detail CurrentUnit on CurrentUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and 	CurrentUnit.uom_code=	TSPL_DAIRYSALE_GATEPASS_DETAIL.unit_code 
				   left join tspl_item_uom_detail CrateUnit on CrateUnit.item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code  and 	CrateUnit.uom_code=	'Crate' 
				   left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='Pouch') as ItemConversionInPouch on ItemConversionInPouch.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code
                    left join ( select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'Crate' ) as ItemConversionCrate on ItemConversionCrate.Item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code 
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='LTR') as ItemConversionInLTR on ItemConversionInLTR.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code 
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='Box') as ItemConversionInBox on ItemConversionInBox.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code
                     left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='KG') as ItemConversionInKG on ItemConversionInKG.Item_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=tspl_vehicle_master.Transport_Id  left join tspl_item_uom_detail gpUnit on gpUnit.item_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.item_code and CrateUnit.uom_code = TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code "
        Else
            qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=tspl_vehicle_master.Transport_Id  "
        End If

        qry += " left outer join ((select Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty,0))>0 Then  Sum(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt/TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GP_Qty) Else 0 End As Amount,  Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty, 0))> 0 Then Sum(TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount) Else 0 End As AmountWithoutTax, Sum(IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0)/TSPL_SD_SHIPMENT_DETAIL.Qty*TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GP_Qty) AS Margin,sum( IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0))as SecurityAmt,TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Unit_code,max(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate) as Dist_Commission_Ratewithtax,max(TSPL_SD_SHIPMENT_HEAD.RoundOffAmount) as RoundOffAmt , max(TSPL_SD_SHIPMENT_DETAIL.Item_Cost) as ItemCost,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode,max(TSPL_BOOKING_MATSER.NoCrateIssue)NoCrateIssue,Max(TSPL_SD_SHIPMENT_HEAD.Is_Taxable)Is_Taxable from " & tbl_TSPL_SD_SHIPMENT_DETAIL & "   
                       Left Outer Join " & tbl_TSPL_SD_SHIPMENT_HEAD & " ON TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
                       left outer join " & tbl_TSPL_BOOKING_MATSER & " on TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
					   Left Outer Join " & tbl_TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL & " On TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID=TSPL_SD_SHIPMENT_DETAIL.PK_ID
                     WHERE  TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode  in(" & GPCode & ") "
        qry += " Group By  TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Unit_code,TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode)
                     Union All
                     (select Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty,0))>0 Then  Sum(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt) Else 0 End As Amount,  Case When Sum(Isnull(TSPL_SD_SHIPMENT_DETAIL.Qty, 0))> 0 Then Sum(TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount) Else 0 End As AmountWithoutTax, 
                     Sum(IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0)) AS Margin,sum( IsNull(TSPL_SD_SHIPMENT_DETAIL.Security_Amt,0))as SecurityAmt,TSPL_SD_SHIPMENT_DETAIL.Item_Code ,TSPL_SD_SHIPMENT_DETAIL.Unit_code ,max(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate) as Dist_Commission_Ratewithtax,max(TSPL_SD_SHIPMENT_HEAD.RoundOffAmount) as RoundOffAmt , max(TSPL_SD_SHIPMENT_DETAIL.Item_Cost) as ItemCost,TSPL_SD_SHIPMENT_HEAD.GPCode,max(TSPL_BOOKING_MATSER.NoCrateIssue)NoCrateIssue,Max(TSPL_SD_SHIPMENT_HEAD.Is_Taxable)Is_Taxable from " & tbl_TSPL_SD_SHIPMENT_DETAIL & "
                     Left Outer Join " & tbl_TSPL_SD_SHIPMENT_HEAD & " ON TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
                     left outer join " & tbl_TSPL_BOOKING_MATSER & " on TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
                     WHERE  TSPL_SD_SHIPMENT_HEAD.GPCode  in(" & GPCode & ") "
        qry += " Group By  TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Unit_code, TSPL_SD_SHIPMENT_HEAD.GPCode)
                     )xyz ON xyz.GPCode=TSPL_DAIRYSALE_GATEPASS_DETAIL.GPCode And xyz.Unit_code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Unit_Code and xyz.Item_Code=TSPL_DAIRYSALE_GATEPASS_DETAIL.Item_Code
					 where 2=2  and TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode in (" & GPCode & ") "
        qry += " )As Main Group by Item_code,GPCode" & IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code", "") & ")AS Final 
					 left outer join  ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND], 
                    max([DESCRP]) as [DESCRP],max([PACK]) as [PACK], max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA], 
                    max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW], max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC])  as [BRANDDESC],
                   max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC], max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC],
                   max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC], max([SCRAPDESC]) as [SCRAPDESC] 
                   from ( select * from (   select TSPL_ITEM_MASTER.Item_Code" & IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code", "") & ",TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   ,
                   TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER   
                   left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code 
                   left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and  TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  
                   where 2=2 )xx  Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE],  [CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt 
                   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC], [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC], [CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx 
                   group by Item_Code " & IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal, ",Unit_Code", "") & "  )  as tbl_Brand on tbl_Brand.Item_Code=Final.Item_Code 
                     left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =Final.Comp_Code )xx  order by xx.sku_seq "

        Return qry
    End Function

    Public Function funPrint2(ByVal Code As String, ByVal isfilePath As Boolean) As String
        Return funPrint2(Code, isfilePath, False)
    End Function

    Public Function funPrint2(ByVal Code As String, ByVal isfilePath As Boolean, ByVal isCancel As Boolean) As String
        Dim filePath As String = Nothing
        Try
            If CreateGatePassFromDemand Then
                Dim frm As New frmDemandBooking()
                frm.PrintGatePass("DG", Code, IIf(rbtnMorning.IsChecked, "Morning", "Evening"), Nothing, Nothing)
            Else
                Dim dt2 As DataTable = Nothing
                Dim subrptqry As String = ""
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    atchqry = getUDPattachqry("'" & Code & "'", Nothing, Nothing, Nothing, "", isCancel)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    atchqry = GetAttachQryBKN(Code, isDepartmentRoute, isCancel)
                Else
                    atchqry = GetAttachQry(Code, isDepartmentRoute, isCancel)
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                    subrptqry = CrateInOut()
                    dt2 = clsDBFuncationality.GetDataTable(subrptqry)
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    subrptqry = CountUOM(Code)
                    dt2 = clsDBFuncationality.GetDataTable(subrptqry)
                End If

                If dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntriesGNG", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal Then
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairyGatePassJSL", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                        'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntry", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                        ' pdfpath = frmCRV.funsubreportWithdt(True, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptDairySaleGatePassEntryNew", "Dairy Sale Gate Pass", clsCommon.myCDate(dt.Rows(0)("GPDate")), "crptDairySaleGatePassEntryNew.rpt", "", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, dt2, "crptDairySaleGatePassEntriesTNK", "Dairy Sale Gate Pass", "subrptCrateInOut.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntriesUDP", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                        'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntriesJPR", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                        filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, dt2, "crptDairySaleGatePassEntriesJPR", "Dairy Sale Gate Pass", "subrptCountUOM.rpt")

                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntriesJDP", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntriesALW", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                        filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWiseAJM1", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("GPDate")), "rptCompanyAddress.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
                        filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptDairySaleGatePassEntriesKTA", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("GPDate")), "rptCompanyAddress.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        filePath = filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptDairySaleGatepassBKN", "Gate Pass", "", "rptCompanyAddress.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        filePath = filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptDairySaleGatePassEntriesNAG", "Gate Pass", "", "rptCompanyAddress.rpt")

                    Else
                        filePath = frmCRV.funreport(MyBase.Form_ID, isfilePath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntries", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                        'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntries", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    End If
                    'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntry", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNew", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntries", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    frmCRV = Nothing
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return filePath
    End Function
    Public Sub funPrint(ByVal Code As String, ByVal isPDFPath As Boolean)
        Try
            If clsCommon.myLen(Code) <= 0 Then
                Throw New Exception("No data found to Print")
            End If
            Dim strPath As String = GatepassWithFilePath(Code, isPDFPath)
            If clsCommon.myLen(strPath) > 0 AndAlso clsCommon.myLen(objCommonVar.CurrentUserEmailID) > 0 Then
                clsCommonFunctionality.SendInstantEmail(objCommonVar.CurrentUserEmailID, "Email by XpertERP " & Me.Text, "Please find the attached " & Me.Text & " no [" & Code & "]", False, strPath)
                clsCommon.MyMessageBoxShow(Me, "Mail Send", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function GatepassWithFilePath(ByVal Code As String, ByVal isPDFPath As Boolean)
        Return GatepassWithFilePath(Code, txtDate.Value, IIf(rbtnMorning.IsChecked, "Morning", "Evening"), Nothing, Nothing, Nothing, Nothing, isPDFPath)
    End Function

    Function GatepassWithFilePath(ByVal Code As String, ByVal DocDate As DateTime, ByVal strShift As String, ByVal isFresh As Boolean, ByVal isAmbient As Boolean, ByVal strRoute As String, ByVal strLocation As String, ByVal isPDFPath As Boolean) As String
        Return GatepassWithFilePath(Code, txtDate.Value, IIf(rbtnMorning.IsChecked, "Morning", "Evening"), isFresh, isAmbient, strRoute, strLocation, isPDFPath, False)
    End Function

    Function GatepassWithFilePath(ByVal Code As String, ByVal DocDate As DateTime, ByVal strShift As String, ByVal isFresh As Boolean, ByVal isAmbient As Boolean, ByVal strRoute As String, ByVal strLocation As String, ByVal isPDFPath As Boolean, ByVal isCancel As Boolean) As String
        Dim strPath As String = Nothing
        If CreateGatePassFromDemand Then
            Dim frm As New frmDemandBooking()
            strPath = frm.PrintGatePass("DG", Code, strShift, isFresh, isAmbient, isPDFPath)
        Else
            If clsCommon.myLen(strRoute) > 0 Then
                ApplyDepartmentRoute = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyDepartmentRoute, clsFixedParameterCode.ApplyDepartmentRoute, Nothing)) = 1, True, False)
                If ApplyDepartmentRoute Then
                    isDepartmentRoute = clsCommon.myCBool(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Department_Route from TSPL_ROUTE_MASTER where Route_No='" & strRoute & "'")))
                End If
            End If
            If rbtn_Milk.IsChecked OrElse isFresh Then
                atchqry = GetAttachQry(Code, isDepartmentRoute, isCancel)
            ElseIf rbtn_product.IsChecked OrElse rbtn_IceCream.IsChecked OrElse isAmbient Then
                atchqry = GetProductPrintQry(Code, isCancel)
            End If

            Dim subrptqry As String = Nothing
            Dim dt2 As DataTable = Nothing
            If clsCommon.myLen(strRoute) > 0 AndAlso clsCommon.myLen(strLocation) > 0 Then
                subrptqry = CrateInOut(DocDate, strRoute, strLocation)
            Else
                subrptqry = CrateInOut()
            End If
            dt2 = clsDBFuncationality.GetDataTable(subrptqry)

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
            If dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNewALW", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                    strPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, dt2, "crptDairySaleGatePassEntryNewTNK", "Dairy Sale Gate Pass", "subrptCrateInOut.rpt")
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNewUDP", "Dairy GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    If (isDepartmentRoute AndAlso rbtn_Milk.IsChecked) OrElse isFresh Then
                        strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNewJPR-DRP", "Dairy GatePass Entry DRP", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    ElseIf rbtn_product.IsChecked OrElse rbtn_IceCream.IsChecked OrElse isAmbient Then
                        strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassProductJPR", "Dairy Product GatePass ", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    Else
                        strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNewJPR", "Dairy GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                    End If
                Else
                    strPath = frmCRV.funreport(MyBase.Form_ID, isPDFPath, CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNew", "Dairy GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))
                End If
                frmCRV = Nothing
            End If
        End If
        Return strPath
    End Function

    '  Public Function GetQueryGatePass() As String
    '      Dim Qry As String = "select TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1,TSPL_COMPANY_MASTER.Add2 as Comp_Add2,TSPL_COMPANY_MASTER.Add3 as Comp_Add3,TSPL_COMPANY_MASTER.Pincode as Comp_Pin
    '      ,TSPL_LOCATION_MASTER.Add1,TSPL_LOCATION_MASTER.Add2,TSPL_LOCATION_MASTER.ADD3,TSPL_LOCATION_MASTER.Pin_Code,TSPL_LOCATION_MASTER.Location_Desc
    '      ,Main_Final.Distributor,'" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "' shiftType,Main_Final.City_Name,Main_Final.Demand_No,Main_Final.Demand_Date,Main_Final.Route_No,Main_Final.Route_Desc
    '      ,Main_Final.Item_alies_name,Main_Final.Crate_Qty,Main_Final.Pouch_Qty,Main_Final.Loose_Qty,TotalLtr_ItemWise
    '      from (select max(TSPL_VENDOR_MASTER.vendor_name) as Distributor,
    '      max(TSPL_DEMAND_BOOKING_MASTER.shiftType) as shiftType,
    '      max(TSPL_city_MASTER.City_Name) as City_Name,
    '      max(TSPL_DEMAND_BOOKING_MASTER.Comp_Code) as Comp_Code,
    '      max(TSPL_DEMAND_BOOKING_MASTER.location_code) as location_code
    '      ,TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,max(convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)) as Demand_Date ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,max(isnull(TSPL_ROUTE_MASTER.Route_Desc,'')) as Route_Desc 
    '      ,max(TSPL_ITEM_MASTER.alies_name) as Item_alies_name
    '      ,sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Crate' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Crate_Qty
    ',sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Pouch' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Pouch_Qty
    ',sum(case when (TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Crate' and TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Pouch') then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Loose_Qty
    ',sum(TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise) AS TotalLtr_ItemWise
    '      from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
    '      on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
    '       left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
    '      left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
    '       left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
    '      left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
    '      left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
    '      left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VEHICLE_MASTER.Transport_Id
    '      where convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & " ' And  
    '      TSPL_DEMAND_BOOKING_MASTER.Location_Code='" & txtLocCode.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code='" + txtVehicle.Value + "'
    '      and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'
    '       group by TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Route_No
    '      ) as Main_Final
    '      LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=Main_Final.Comp_Code
    '       LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =Main_Final.location_code"
    '      Return Qry
    '  End Function
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        funPrint(txtCode.Value, False)
    End Sub
    Public Sub UploadDoc(ByVal Code As String)
        Dim pdfpath As String = ""
        Try
            If settFileUpload Then
                atchqry = GetAttachQry(Code, isDepartmentRoute)
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
                If dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptDairySaleGatePassEntryNewALW", "Dairy Sale GatePass Entry", clsCommon.myCDate(dt.Rows(0)("GPDate")))

                    Else
                        pdfpath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptDairySaleGatePassEntryNew", "Dairy Sale Gate Pass", clsCommon.myCDate(dt.Rows(0)("GPDate")), "crptDairySaleGatePassEntryNew.rpt", "", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        'pdfpath = frmCRV.funsubreportWithdt(True, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptDairySaleGatePassEntryNew", "Dairy Sale GatePass Entry", "", "", Nothing, "", Nothing, "", Nothing)
                    End If
                    frmCRV = Nothing
                End If
                If clsCommon.myLen(pdfpath) > 0 Then
                    Dim FileNo As Integer = clsAttachDocument.UploadWithHttpRequest(pdfpath, Path.GetFileName(pdfpath), clsUserMgtCode.frmDairyGatePass, Code)
                    If FileNo > 0 Then
                        Dim qry As String = " UPDATE TSPL_DAIRYSALE_GATEPASS_MASTER set FILE_INFO=" + clsCommon.myCstr(FileNo) + " where GPCode='" + Code + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry)
                    End If
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Public Sub UploadInvoice(ByVal SaleInvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String)
        Dim pdfpath As String = ""
        Try
            If settFileUpload Then
                Dim Qry As String = Nothing
                Dim frmCRV As New frmCrystalReportViewer()
                Dim objMultPrintInvoice As New FrmPrintFreshInvoice
                If clsCommon.myLen(SaleInvoiceNo) <= 0 Then
                    myMessages.blankValue(Me, "Invoice not found to Print", Me.Text)
                Else
                    Dim dtDocdate As Date?
                    dtDocdate = Nothing
                    Dim StrSql = "Select Document_Code,Document_Date,Customer_Code,Bill_To_Location,is_taxable,Tax_Group from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" + SaleInvoiceNo + "'"
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrSql)
                    If dt1.Rows.Count > 0 Then
                        dtDocdate = clsCommon.myCDate(dt1.Rows(0)("Document_Date"))
                    End If
                    Dim InvoiceNo As String = "'" + SaleInvoiceNo + "'"
                    Qry = objMultPrintInvoice.PrintInvoiceForAll(InvoiceNo, docDate, CustCode)
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                    If dt.Rows.Count > 0 Then
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                            pdfpath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceTNK", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        Else
                            pdfpath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        End If
                        frmCRV = Nothing
                    End If
                    If clsCommon.myLen(pdfpath) > 0 Then
                        Dim DocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Sale_Invoice_No='" + SaleInvoiceNo + "'"))
                        Dim FileNo As Integer = clsAttachDocument.UploadWithHttpRequest(pdfpath, Path.GetFileName(pdfpath), clsUserMgtCode.frmSaleDispatchDairy, DocNo)
                        If FileNo > 0 Then
                            StrSql = " UPDATE TSPL_SD_SHIPMENT_HEAD set FILE_INFO=" + clsCommon.myCstr(FileNo) + " where Document_Code='" + DocNo + "'"
                            clsDBFuncationality.ExecuteNonQuery(StrSql)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub btnPrint2_Click(sender As Object, e As EventArgs) Handles btnPrint2.Click
        '=====update by preeti gupta Against ticket no[ERO/05/09/19-001019,ERO/05/09/19-001020,TEC/20/05/19-000509]
        If clsCommon.myLen(txtCode.Value) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No data found to Print", Me.Text)
        Else
            funPrint2(txtCode.Value, False)
        End If
    End Sub
    Private Sub fndRouteNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndRouteNo._MYValidating
        If CreateGatePassFromDemand = True Then
            strQuery = " Select distinct Route_No As Code,route_desc from (" &
            " Select distinct TSPL_DEMAND_BOOKING_MASTER.Route_No,tspl_route_master.route_desc from TSPL_DEMAND_BOOKING_MASTER " &
            " left outer join tspl_route_master On tspl_route_master.route_no=TSPL_DEMAND_BOOKING_MASTER.route_no " &
            " where TSPL_DEMAND_BOOKING_MASTER.Route_No<>''" &
            " ) final"
        Else
            If chkAgainstTransfer.Checked = True Then
                strQuery = " select distinct Route_No as Code,route_desc from (select Route_No ,route_desc from tspl_route_master where Status='A' ) final"
            Else
                strQuery = " Select distinct Route_No As Code,route_desc from (" &
    " Select distinct TSPL_SD_SHIPMENT_HEAD.Route_No,tspl_route_master.route_desc from TSPL_SD_SHIPMENT_HEAD "
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    strQuery += "Left outer join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                End If
                strQuery += " left outer join tspl_route_master On tspl_route_master.route_no=TSPL_SD_SHIPMENT_HEAD.route_no " &
    " where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' and TSPL_SD_SHIPMENT_HEAD.Route_No<>'' "
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    If chkGhee.Checked Then
                        strQuery += " and TSPL_BOOKING_MATSER.Is_GHEE = 1 "
                    Else
                        strQuery += " and isnull(TSPL_BOOKING_MATSER.Is_GHEE,0) = 0 "
                    End If
                End If
                strQuery += " ) final"
            End If
        End If
        fndRouteNo.Value = clsCommon.ShowSelectForm("DSRoute", strQuery, "code", "", fndRouteNo.Value, "", isButtonClicked)
        If clsCommon.myLen(fndRouteNo.Value) > 0 Then
            txtRouteName.Text = clsDBFuncationality.getSingleValue("Select Route_Desc from TSPL_ROUTE_MASTER where Route_No='" & fndRouteNo.Value & "'")
            If CreateGatePassFromDemand = True Then
                Dim StrGP_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GPCode from TSPL_DAIRYSALE_GATEPASS_MASTER where convert(date, TSPL_DAIRYSALE_GATEPASS_MASTER.gpdate ,103)=convert(date,'" & txtDate.Value & "',103) and Route_No='" & fndRouteNo.Value & "'"))
                If clsCommon.myLen(StrGP_No) > 0 Then
                    clsCommon.MyMessageBoxShow(Me, "GatePass Already Created.", Me.Text)
                    LoadData(StrGP_No, NavigatorType.Current)
                Else
                    setRouteVehicleDetail()
                End If
            End If
        Else
            fndRouteNo.Value = ""
            txtRouteName.Text = ""
        End If
        strQuery = ReturnVehicle()
        strQuery += " where tspl_route_master.Route_No='" + clsCommon.myCstr(fndRouteNo.Value) + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQuery)
        If dt.Rows IsNot Nothing AndAlso dt.Rows.Count = 1 Then
            txtVehicle.Value = clsCommon.myCstr(dt.Rows(0)("Code"))
            lblVehicleDesc.Text = clsCommon.myCstr(dt.Rows(0)("Description"))
        End If
        If EnableLocation Then
            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                dt = Nothing
                strQuery = "select Location_Code from TSPL_ROUTE_MASTER where Route_No ='" + clsCommon.myCstr(fndRouteNo.Value) + "'"
                dt = clsDBFuncationality.GetDataTable(strQuery)
                If dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    txtLocCode.Value = clsCommon.myCstr(dt.Rows(0)("Location_Code"))
                    txtLocDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocCode.Value + "'"))
                End If
            End If
        End If
    End Sub
    Sub setRouteVehicleDetail()
        Try
            Dim qry As String = ""
            qry = "select TSPL_ROUTE_MASTER.City_Code,Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_ROUTE_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon,tspl_transport_master.Transporter_Name,TSPL_VEHICLE_MASTER.Transport_Id from TSPL_CUSTOMER_MASTER left outer join " &
                "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
                "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id left outer join tspl_transport_master on tspl_transport_master.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id "
            qry += " where TSPL_ROUTE_MASTER.Route_No ='" & fndRouteNo.Value & "'"
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                txtVehicle.Value = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
                lblVehicleDesc.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
                fndRouteNo.Value = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
                txtRouteName.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
                txtTransporter.Text = clsCommon.myCstr(dt1.Rows(0)("Transport_Id"))
                If clsCommon.myLen(clsCommon.myCstr(txtVehicle.Value)) <= 0 Then
                    Throw New Exception("Please Map Vehicle with Route " & txtRouteName.Text & "")
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub txtmultiBooking__My_Click(sender As Object, e As EventArgs) Handles txtmultiBooking._My_Click
        Dim qry As String = LoadQuery("", EnableDispatch)
        Dim strAllDoc As String = " select distinct PPPP.[Document No] as code, convert (varchar, PPPP .[Document Date], 103) as [Document Date] from  ( " & qry & "    ) As PPPP   "
        txtmultiBooking.arrValueMember = clsCommon.ShowMultipleSelectForm("TransType@@@@MulSel", strAllDoc, "Code", "code", txtmultiBooking.arrValueMember, txtmultiBooking.arrDispalyMember)
        funFillGrid2()
        TotalQty()
    End Sub
    Private Sub TotalQty()
        Dim TotalCrate As Integer = 0
        Dim TotalCan As Integer = 0
        For i As Integer = 0 To Gv1.Rows.Count - 1
            If clsCommon.CompairString(clsCommon.myCstr(Gv1.Rows(i).Cells(colUnit).Value), "Crate") = CompairStringResult.Equal Then
                TotalCrate = TotalCrate + clsCommon.myCdbl(Gv1.Rows(i).Cells(colQty).Value)
                'ElseIf clsCommon.CompairString(clsCommon.myCstr(Gv1.Rows(i).Cells(colUnit).Value), "Can") = CompairStringResult.Equal Then
                '    TotalCan = TotalCan + clsCommon.myCdbl(Gv1.Rows(i).Cells(colQty).Value)
            End If
        Next
        If clsCommon.myCdbl(TotalCrate) > 0 Then
            If Not AllowManualCrateForDispatch Then
                txtCrateQty.Text = TotalCrate
            End If
            'Else
            '    txtCrateQty.Text = 0
        End If
    End Sub
    Private Sub funFillGrid2()
        Try
            LoadBlankGrid()
            Dim qry As String = LoadQuery("", EnableDispatch)
            strQuery = "select [Item Code],max([Item Desc]) as [Item Desc],Unit,sum(qty) as Quantity,max(HSN_Code) as HSN_Code  from ( select xxx.* from(  " & qry & " )xxx  where xxx.[Document No] in (" + clsCommon.GetMulcallString(txtmultiBooking.arrValueMember) + ")  )  final group by [Item Code],Unit "
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(strQuery)
            Dim intLineNo As Integer = 0
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                For Each dr As DataRow In dt.Rows
                    Gv1.Rows.AddNew()
                    intLineNo += 1
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colLineNo).Value = intLineNo
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemCode).Value = clsCommon.myCstr(dr("Item Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemDesc).Value = clsCommon.myCstr(dr("Item Desc"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCstr(dr("Quantity"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colHSNCode).Value = clsCommon.myCstr(dr("HSN_Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colSchemeItem).Value = clsCommon.myCstr(dr("Scheme_Item"))
                    txtDate.Enabled = False
                    'txtSupplyDate.Enabled = False
                Next
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "GatePass Entry", MessageBoxButtons.OK, Me.Text)
        End Try
    End Sub
    Private Sub btnClKM_Click(sender As Object, e As EventArgs) Handles btnClKM.Click
        Try
            If chkAgainstTransfer.Checked = True Then
                Dim strTransferIn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No  from TSPL_TRANSFER_ORDER_HEAD where TransferOutNo ='" & clsCommon.myCstr(FndTransferNo.Value) & "' And TSPL_TRANSFER_ORDER_HEAD.status = 1"))
                If clsCommon.myLen(strTransferIn) > 0 Then
                    Dim strTransferReturnAgainstIn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No  from TSPL_TRANSFER_RETURN where Transfer_No ='" & strTransferIn & "'"))
                    If clsCommon.myLen(strTransferReturnAgainstIn) > 0 Then
                        Throw New Exception("Please create Transfer In Against Transfer Out No " & FndTransferNo.Value & " ")
                    End If
                Else
                    Throw New Exception("Please create Transfer In Against Transfer Out No " & FndTransferNo.Value & "")
                End If
            End If
            Dim dclDistanceInRoute As String = Nothing
            Dim dclPriceKM As String = Nothing
            Dim dclTollAmt As String = Nothing
            Dim qry As String = " select TSPL_ROUTE_MASTER.Distance,TSPL_VEHICLE_MASTER.Price_KM,  isnull (TSPL_ROUTE_MASTER.Toll_amount,0) as Toll_Amount from TSPL_DAIRYSALE_GATEPASS_MASTER left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DAIRYSALE_GATEPASS_MASTER.Route_No " &
                                " left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id = TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id " &
                                " where GPCode = '" + txtCode.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                dclDistanceInRoute = clsCommon.myCdbl(dt.Rows(0)("Distance"))
                dclPriceKM = clsCommon.myCdbl(dt.Rows(0)("Price_KM"))
                dclTollAmt = clsCommon.myCdbl(dt.Rows(0)("Toll_Amount"))
            End If
            If IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TollTaxMaster, clsFixedParameterCode.TollTaxMaster, Nothing)) = 0, False, True) = False Then
                dclTollAmt = txtTollAmount.Text
            End If
            clsDairyGatePassEntry.UpdateClosingKMAndCreateProvision(txtCode.Value, txtClKM.Value, dclDistanceInRoute, dclPriceKM, dclTollAmt, MyBase.Form_ID)
            clsCommon.MyMessageBoxShow(Me, "Closing KM Update successfully", Me.Text)
            LoadData(txtCode.Value, NavigatorType.Current)
        Catch ex As Exception
            If clsCommon.myLen(ex.Message) > 0 Then
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End If
        End Try
    End Sub
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If (myMessages.deleteConfirm()) Then
                If (clsDairyGatePassEntry.DeleteData(txtCode.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    btnNew.PerformClick()
                End If
            End If
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub
    Private Sub btnReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverse.Click
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                If clsDairyGatePassEntry.ReverseAndUnpost(txtCode.Value) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed", Me.Text)
                    LoadData(txtCode.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub chkAgainstTransfer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkAgainstTransfer.ToggleStateChanged
        Try
            If chkAgainstTransfer.Checked = True Then
                FndTransferNo.Visible = True
                MyLabel9.Visible = True
            Else
                FndTransferNo.Visible = False
                MyLabel9.Visible = False
            End If
            FndTransferNo.Value = ""
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub FndTransferNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles FndTransferNo._MYValidating
        Dim sQuery As String = String.Empty
        Dim whrcls As String = ""
        If chkAgainstTransfer.Checked = True Then
            sQuery = "select TSPL_TRANSFER_ORDER_HEAD.Document_NO as DocNo,CONVERT(varchar, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as DocDate,From_Location as " _
                       & " [From Location Code],frm.Location_Desc as [From Location],fto.Location_Code as [To Location Code],fto.Location_Desc as [To Location],DOC_Total_Amt as [Amount]  " _
                       & " from TSPL_TRANSFER_ORDER_HEAD left join tspl_Location_Master frm on frm.location_Code=TSPL_TRANSFER_ORDER_HEAD.From_Location left join " _
                       & " tspl_Location_Master fto on fto.GIT_Location=TSPL_TRANSFER_ORDER_HEAD.to_Location left join TSPL_TRANSFER_RETURN on transfer_No=TSPL_TRANSFER_ORDER_HEAD.Document_NO"
            whrcls = "  TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' and TSPL_TRANSFER_ORDER_HEAD.Status =1 and  TSPL_TRANSFER_ORDER_HEAD.Document_No not  in (select coalesce(TransferOutNo,'') from TSPL_TRANSFER_ORDER_HEAD where Document_No not in (select Transfer_No from TSPL_TRANSFER_RETURN) union all select AgainstTransferNo from TSPL_DAIRYSALE_GATEPASS_MASTER  where isnull(AgainstTransferNo ,'')<>'' ) and TSPL_TRANSFER_RETURN.transfer_No is Null and TSPL_TRANSFER_ORDER_HEAD.internalTransfer=0  and convert(date, TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)<=convert(date,'" & txtDate.Value & "',103)"
            '    sQuery = " Select TSPL_TRANSFER_ORDER_HEAD.Document_NO as DocNo,CONVERT(varchar, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as DocDate,From_Location as  [From Location Code],frm.Location_Desc as [From Location],fto.Location_Code as [To Location Code],fto.Location_Desc as [To Location],DOC_Total_Amt as [Amount]  from TSPL_TRANSFER_ORDER_HEAD " & Environment.NewLine & _
            '" left join tspl_Location_Master frm on frm.location_Code=TSPL_TRANSFER_ORDER_HEAD.From_Location left join  tspl_Location_Master fto on fto.GIT_Location=TSPL_TRANSFER_ORDER_HEAD.to_Location "
            '    whrcls = " Document_No in (Select TransferOutNo from TSPL_TRANSFER_ORDER_HEAD where isnull (TransferOutNo ,'') in (select TSPL_TRANSFER_ORDER_HEAD.Document_NO as DocNo from TSPL_TRANSFER_ORDER_HEAD  left join TSPL_TRANSFER_RETURN on transfer_No=TSPL_TRANSFER_ORDER_HEAD.Document_NO where " & Environment.NewLine & _
            '" TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' and TSPL_TRANSFER_ORDER_HEAD.Status =1 and  TSPL_TRANSFER_ORDER_HEAD.Document_No not  in (select coalesce(TransferOutNo,'') from TSPL_TRANSFER_ORDER_HEAD where Document_No not in (select Transfer_No from TSPL_TRANSFER_RETURN) and TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ='O' " & Environment.NewLine & _
            '" union all select AgainstTransferNo from TSPL_DAIRYSALE_GATEPASS_MASTER  where isnull(AgainstTransferNo ,'')<>'') and TSPL_TRANSFER_RETURN.transfer_No is Null " & Environment.NewLine & _
            '" --and TSPL_TRANSFER_ORDER_HEAD.Document_No in ('TTO-001/20-21/000294','TNT-001/20-21/00023') " & Environment.NewLine & _
            '" )) "
            FndTransferNo.Value = clsCommon.ShowSelectForm("SRNRetSRNFndOI", sQuery, "DocNo", whrcls, FndTransferNo.Value, "TSPL_TRANSFER_ORDER_HEAD.Document_Date", isButtonClicked, "TSPL_TRANSFER_ORDER_HEAD.Document_Date")
            If clsCommon.myLen(FndTransferNo.Value) > 0 Then
                txtLocCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select FROM_LOCATION from TSPL_TRANSFER_ORDER_HEAD WHERE DOCUMENT_NO='" & FndTransferNo.Value & "'"))
                txtLocDesc.Text = clsDBFuncationality.getSingleValue("select  Location_Desc  from TSPL_LOCATION_MASTER where Location_Code='" & txtLocCode.Value & "'")
                txtLocCode.Enabled = False
            Else
                txtLocCode.Value = ""
                txtLocDesc.Text = ""
                txtLocCode.Enabled = True
            End If
        Else
            clsCommon.MyMessageBoxShow(Me, "Please check Against Transfer checkbox first", Me.Text)
        End If
    End Sub
    Private Sub funFillGrid_Transfer()
        Try
            LoadBlankGrid()
            Dim qry As String = "select TSPL_TRANSFER_ORDER_DETAIL.Item_Code as [Item Code],TSPL_TRANSFER_ORDER_DETAIL.Item_Desc as [Item Desc],TSPL_TRANSFER_ORDER_DETAIL.Unit_code as Unit ,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,TSPL_ITEM_MASTER.HSN_Code,TSPL_TRANSFER_ORDER_HEAD.From_Location ,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code   from TSPL_TRANSFER_ORDER_HEAD " & Environment.NewLine &
" left outer join TSPL_TRANSFER_ORDER_DETAIL on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No " & Environment.NewLine &
" left outer join TSPL_ITEM_MASTER on TSPL_TRANSFER_ORDER_DETAIL.Item_Code =TSPL_ITEM_MASTER.Item_Code " & Environment.NewLine &
             " where 1=1 and TSPL_TRANSFER_ORDER_HEAD.Document_No='" + FndTransferNo.Value + "' "
            strQuery = "select [Item Code],max([Item Desc]) as [Item Desc],Unit,sum(qty) as Quantity,max(HSN_Code) as HSN_Code,max(From_Location) as LOC ,max(Vehicle_Code) as Vehicle_Code  from ( " & qry & " ) final group by [Item Code],Unit "
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(strQuery)
            Dim intLineNo As Integer = 0
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtLocCode.Value = clsCommon.myCstr(dt.Rows(0)("LOC"))
                txtLocDesc.Text = clsLocation.GetName(clsCommon.myCstr(dt.Rows(0)("LOC")), Nothing)
                If clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))) > 0 Then
                    txtVehicle.Value = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
                End If
                lblVehicleDesc.Text = clsDBFuncationality.getSingleValue("select Description from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'")
                If isCreateProvisionOfTransporterInDairyDispatch = True Then
                    txtTransporter.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicle.Value & "'"))
                End If
                txtLocCode.Enabled = False
                For Each dr As DataRow In dt.Rows
                    Gv1.Rows.AddNew()
                    intLineNo += 1
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colLineNo).Value = intLineNo
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemCode).Value = clsCommon.myCstr(dr("Item Code"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colItemDesc).Value = clsCommon.myCstr(dr("Item Desc"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCstr(dr("Quantity"))
                    Gv1.Rows(Gv1.Rows.Count - 1).Cells(colHSNCode).Value = clsCommon.myCstr(dr("HSN_Code"))
                    txtDate.Enabled = False
                    ' txtSupplyDate.Enabled = False
                Next
                Dim TotalCrate As Integer = 0
                Dim TotalCan As Integer = 0
                For i As Integer = 0 To Gv1.Rows.Count - 1
                    If clsCommon.CompairString(clsCommon.myCstr(Gv1.Rows(i).Cells(colUnit).Value), "Crate") = CompairStringResult.Equal Then
                        TotalCrate = TotalCrate + clsCommon.myCdbl(Gv1.Rows(i).Cells(colQty).Value)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(Gv1.Rows(i).Cells(colUnit).Value), "Can") = CompairStringResult.Equal Then
                        TotalCan = TotalCan + clsCommon.myCdbl(Gv1.Rows(i).Cells(colQty).Value)
                    End If
                Next
                If clsCommon.myCdbl(TotalCrate) > 0 Then
                    txtCrateQty.Text = TotalCrate
                Else
                    txtCrateQty.Text = 0
                End If
                If clsCommon.myCdbl(TotalCan) > 0 Then
                    txtCanQty.Text = TotalCan
                Else
                    txtCanQty.Text = 0
                End If
                ' **************************************************************************************************
                'txtmultiBooking.Enabled = True
                ''Dim strAllDoc As String = " select STUFF((SELECT ',' + Document_Code from (select distinct PPPP.Document_Code from  ( " & qry & "    ) As PPPP  ) Final FOR XML PATH('')), 1, 1, '') "
                'Dim strAllDoc As String = " select distinct PPPP.[Document No] from  ( " & qry & "    ) As PPPP   "
                'Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(strAllDoc)
                'Dim list As New ArrayList
                'If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                '    For Each dr As DataRow In dt1.Rows
                '        list.Add(dr("Document No"))
                '    Next
                'End If
                'txtmultiBooking.arrValueMember = list
                ' **************************************************************************************************
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "GatePass Entry", MessageBoxButtons.OK, Me.Text)
        End Try
    End Sub
    Private Sub TxtMultiDairyGPassReverse__My_Click(sender As Object, e As EventArgs) Handles TxtMultiDairyGPassReverse._My_Click
        Dim qry As String = "select GPCode,GPDate,Location_Code,Vehicle_Id,Route_No,TSPL_PROVISION_ENTRY_KNOCKOFF.Provision_No,TSPL_PROVISION_ENTRY_KNOCKOFF.AP_Invoice_No  from TSPL_DAIRYSALE_GATEPASS_MASTER left outer join TSPL_PROVISION_ENTRY_KNOCKOFF on TSPL_PROVISION_ENTRY_KNOCKOFF .Invoice_No =TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode where Post='Y'"
        TxtMultiDairyGPassReverse.arrValueMember = clsCommon.ShowMultipleSelectForm("Dairy@GPassRev", qry, "GPCode", "GPCode", TxtMultiDairyGPassReverse.arrValueMember, TxtMultiDairyGPassReverse.arrDispalyMember)
    End Sub
    Private Sub BtnMultiGPReverse_Click(sender As Object, e As EventArgs) Handles btnMultiGPReverse.Click
        Try
            If TxtMultiDairyGPassReverse.arrValueMember Is Nothing OrElse TxtMultiDairyGPassReverse.arrValueMember.Count <= 0 Then
                Throw New Exception("Please select At least One Document To unpost")
            End If
            Dim StrAllException As String = ""
            For ii As Integer = 0 To TxtMultiDairyGPassReverse.arrValueMember.Count - 1
                Try
                    clsDairyGatePassEntry.ReverseAndUnpost(TxtMultiDairyGPassReverse.arrValueMember(ii))
                Catch ex As Exception
                    StrAllException += "Error In Document no:" + TxtMultiDairyGPassReverse.arrValueMember(ii) + Environment.NewLine + ex.Message
                End Try
            Next
            If clsCommon.myLen(StrAllException) > 0 Then
                clsCommon.MyMessageBoxShow(Me, StrAllException, Me.Text)
            Else
                TxtMultiDairyGPassReverse.arrValueMember = Nothing
                clsCommon.MyMessageBoxShow(Me, "Successfully reverse and unposted", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub lblVehicleDesc_Leave(sender As Object, e As EventArgs) Handles lblVehicleDesc.Leave
        If clsCommon.myLen(lblVehicleDesc.Text) > 0 Then
            VehicleDesc = lblVehicleDesc.Text
        End If
    End Sub
    Private Sub btnGPCancel_Click(sender As Object, e As EventArgs) Handles btnGPCancel.Click
        Dim frm1 As New FrmPWD(Nothing)
        frm1.strType = clsFixedParameterType.Transactionupdate
        frm1.strCode = clsFixedParameterCode.GatePassCancel
        frm1.ShowDialog()
        If frm1.isPasswordCorrect Then
            CancelData()
            OneTimeCheck = True
        End If
    End Sub
    Sub CancelData()
        Try
            If clsCommon.myLen(txtCode.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Document No not found to Post", Me.Text)
                Exit Sub
            End If
            Dim isCancel As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select count (*) from TSPL_DAIRYSALE_GATEPASS_MASTER where GPCode = '" + txtCode.Value + "' and Status = 'Y'"))
            If isCancel = True Then
                common.clsCommon.MyMessageBoxShow(Me, "Record Already canceled.", Me.Text)
                Exit Sub
            End If
            If myMessages.cancelConfirm() Then
                If (clsDairyGatePassEntry.CancelData(MyBase.Form_ID, txtCode.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully canceled", Me.Text)
                    LoadData(txtCode.Value, NavigatorType.Current)
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                End If
                'clsDBFuncationality.ExecuteNonQuery("Update TSPL_DAIRYSALE_GATEPASS_MASTER set post='Y' where gpcode='" & txtCode.Value & "'")
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub Gv1_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles Gv1.CellValueChanged
        Try
            If Not isInsideLoadData Then
                Dim Qry As String = Nothing
                Dim dt As DataTable = Nothing
                isCTQtyUpdate = False
                'If clsCommon.CompairString(clsCommon.myCstr(Gv1.CurrentRow.Cells(colUnit).Value), "Crate") = CompairStringResult.Equal Then
                '    Qry = "Select TSPL_SD_SHIPMENT_DETAIL.PK_ID,TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Unit_Code,IsNull(GPUsed.GP_Qty,0)GP_Qty,TSPL_SD_SHIPMENT_DETAIL.Qty As TotalQty,
                '            Case When IsNull(GPUsed.GP_Qty,0)>0 Then (TSPL_SD_SHIPMENT_DETAIL.Qty-GPUsed.GP_Qty) Else TSPL_SD_SHIPMENT_DETAIL.Qty End As BalanceQty
                '            from TSPL_SD_SHIPMENT_DETAIL Left Join(select  PK_ID,Max(GPCode)GPCode,Max(Item_Code)Item_Code,Max(Unit_Code)Unit_Code,Sum(IsNull(GP_Qty,0))GP_Qty 
                '            from TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL Group By PK_ID) AS GPUsed On GPUsed.PK_ID=TSPL_SD_SHIPMENT_DETAIL.PK_ID
                '            Where TSPL_SD_SHIPMENT_DETAIL.PK_ID='" + clsCommon.myCstr(Gv1.CurrentRow.Cells(colPKID).Value) + "'"
                '    dt = clsDBFuncationality.GetDataTable(Qry)
                '    If clsCommon.myCdbl(Gv1.CurrentRow.Cells(colQty).Value) > 0 Then
                '        If dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                '            If clsCommon.myCDecimal(Gv1.CurrentRow.Cells(colQty).Value) > clsCommon.myCDecimal(dt.Rows(0)("BalanceQty")) AndAlso clsCommon.myCDecimal(dt.Rows(0)("BalanceQty")) > 0 Then
                '                clsCommon.MyMessageBoxShow(Me, "Balance Qty: " + clsCommon.myCstr(dt.Rows(0)("BalanceQty")) + Environment.NewLine + "Allow Only Maximum " + clsCommon.myCstr(dt.Rows(0)("BalanceQty")) + " Qty.", Me.Text)
                '                Gv1.CurrentRow.Cells(colQty).Value = clsCommon.myCDecimal(dt.Rows(0)("BalanceQty"))
                '            End If
                '        End If
                '        'ElseIf clsCommon.myCdbl(Gv1.CurrentRow.Cells(colQty).Value) <= 0 AndAlso dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 AndAlso clsCommon.myCDecimal(dt.Rows(0)("BalanceQty")) > 0 Then
                '        '    Gv1.CurrentRow.Cells(colQty).Value = clsCommon.myCDecimal(dt.Rows(0)("BalanceQty"))
                '    End If
                'End If
                CrateCount_CellValue()
                isInsideLoadData = False
            End If
        Catch ex As Exception
        Finally
            TotalQty()
        End Try
    End Sub
    Private Sub lblVehicleDesc_TextChanged(sender As Object, e As EventArgs) Handles lblVehicleDesc.TextChanged
        If clsCommon.myCstr(VehicleDesc) Is Nothing AndAlso clsCommon.myLen(VehicleDesc) <= 0 Then
            VehicleDesc = lblVehicleDesc.Text
        End If
    End Sub

    'Private Sub btnBoothSlip_Click(sender As Object, e As EventArgs) Handles btnBoothSlip.Click
    '    Try
    '        If clsCommon.myLen(txtCode.Value) > 0 Then


    '            Dim qry As String = " select 

    'XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name, max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq, max(XXFinal.Short_Description) +' '+max(XXFinal.Unit_code) as Short_Description,

    'sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount,'" + objCommonVar.CurrentUser + "' as UserName,
    'max(XXFinal.CompanyName) as CompanyName, '" + txtDistributorName.Text + "' as TranspoterName,'" + txtDriverName.Text + "' as DriverName,'" + txtDriverMobNo.Text + "' as DMobNo,'" + clsCommon.myCstr(lblVehicleDesc.Text) + "' as Vehicle_No,'" + fndRouteNo.Value + "' as Route_No,'" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' as DocumentDate

    'from 
    '(
    'select TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code as TR_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
    'TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,
    'TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount
    ',TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
    ',TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
    '   TSPL_DEMAND_BOOKING_MASTER.Route_No,
    '   TSPL_ROUTE_MASTER.Route_Desc,
    '    TSPL_COMPANY_MASTER.Comp_Name
    '  as CompanyName,
    '  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
    '  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No
    'from  TSPL_SD_SHIPMENT_BOOKING_DETAIL
    'left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
    'left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
    'left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
    'left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
    '  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
    '  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 

    '  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
    '  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
    '  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
    '  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrentCompanyCode + "'

    'where TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' and Route_No='" + clsCommon.myCstr(fndRouteNo.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' and status=1 )
    ')XXFinal
    'group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code "
    '            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '                Dim frmCRV As New frmCrystalReportViewer()
    '                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptBoothSlip", "Booth Demand Sheet")
    '            Else
    '                Throw New Exception(" Data not Found!")
    '            End If
    '        Else
    '            Throw New Exception("Document not Found!")
    '        End If



    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub

    Public Function GetBoothData() As String
        Dim qry As String = " Select *,(xy.ItemAmount) as ItemNetAmount from (select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name,
max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq, max(XXFinal.Short_Description) +' '+max(XXFinal.Unit_code) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemAmount,'" + objCommonVar.CurrentUser + "' as UserName,max(XXFinal.CompanyName) as CompanyName,
'" + txtDistributorName.Text + "' as TranspoterName,'" + txtDriverName.Text + "' as DriverName,'" + txtDriverMobNo.Text + "' as DMobNo,
'" + clsCommon.myCstr(lblVehicleDesc.Text) + "' as Vehicle_No,'" + fndRouteNo.Value + "' as Route_No,
'" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' as DocumentDate,sum(Total_TCS_Amt)Total_TCS_Amt
from (select case  WHEN (TSPL_SD_SHIPMENT_HEAD.TAX1) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX2) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX2_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX3) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX3_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX4) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX4_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX5) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX5_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX6) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX6_Amt)
        WHEN (TSPL_SD_SHIPMENT_HEAD.TAX7) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX7_Amt)
		WHEN (TSPL_SD_SHIPMENT_HEAD.TAX8) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX8_Amt)
		WHEN (TSPL_SD_SHIPMENT_HEAD.TAX9) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX9_Amt)
		WHEN (TSPL_SD_SHIPMENT_HEAD.TAX10) = 'TCS' THEN (TSPL_SD_SHIPMENT_HEAD.TAX10_Amt)
        ELSE 0 END AS Total_TCS_Amt ,
TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code as TR_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No from  TSPL_SD_SHIPMENT_BOOKING_DETAIL
left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
--left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrentCompanyCode + "'

where TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' and Route_No='" + clsCommon.myCstr(fndRouteNo.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' and status=1 )
)XXFinal
group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code)xy "

        Return qry
    End Function



    Private Sub RadMenuItem3_Click(sender As Object, e As EventArgs) Handles RadMenuItem3.Click
        Try
            If clsCommon.myLen(txtCode.Value) > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    PrintBoothSlipData()
                Else
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBoothData())
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        Dim frmCRV As New frmCrystalReportViewer()
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptBoothSlip", "Booth Demand Sheet")
                    Else
                        Throw New Exception("Data not Found!")
                    End If
                End If
            Else
                Throw New Exception("Document not Found!")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function PrintBoothSlipData() As Boolean
        Try
            Dim whrcls As String = ""

            Dim qry As String = "( SELECT TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_VEHICLE_MASTER.DriverName,TSPL_DEMAND_BOOKING_DETAIL.ShiftType, TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Logo_Img,Access_officer,Comp_Code1,Is_FreshItem,Is_Ambient ,TSPL_ITEM_MASTER.IsTaxable,TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number as Description,tspl_vehicle_master.Vehicle_Id,TSPL_COMPANY_MASTER.Comp_Name ,tspl_transport_master.Transporter_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.City_Code,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.State,TSPL_COMPANY_MASTER.Phone1   ,(TSPL_ITEM_MASTER.Alies_Name)Short_Description, "
            qry += "TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_DEMAND_BOOKING_MASTER.Document_Date, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,
TSPL_ITEM_MASTER.Short_Description + 'Amt' AS Item_Description,"
            qry += " TSPL_DEMAND_BOOKING_DETAIL.Unit_code, Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Crate' Then TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty Else 0 end CRATE,TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,TSPL_ITEM_MASTER.Sku_Seq,
		    		Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Pouch' Then TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty Else 0 End Pouch,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount"
            qry += " FROM TSPL_SD_SHIPMENT_BOOKING_DETAIL "
            qry += "left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE 
left join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE = TSPL_SD_SHIPMENT_HEAD.Document_Code and TSPL_SD_SHIPMENT_DETAIL.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code and TSPL_SD_SHIPMENT_DETAIL.Unit_code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code
left join TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL on TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.PK_ID = TSPL_SD_SHIPMENT_DETAIL.PK_ID
left join TSPL_DAIRYSALE_GATEPASS_MASTER on TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode = TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = 'UDP' where 2 = 2 "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                qry += " And IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N'  and TSPL_CUSTOMER_MASTER.Status='N' and TSPL_ITEM_MASTER.Active=1 "
            End If
            qry += " and TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtSupplyDate.Value)), "dd/MMM/yyyy") + "'  and status=1 "
            If rbtnMorning.IsChecked Then
                qry += " and Shift_Type = 'AM' "
            ElseIf rbtnEvening.IsChecked Then
                qry += " and Shift_Type = 'PM' "
            End If
            If clsCommon.myLen(fndRouteNo.Value) > 0 Then
                qry += " And Route_No = '" & fndRouteNo.Value & "' "
            End If
            qry += "))  "

            Dim dtPrint As DataTable = clsDBFuncationality.GetDataTable(qry + " order by Sku_Seq")
            If dtPrint Is Nothing OrElse dtPrint.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow("No Data Found to Print")
                Return False
                Exit Function
            ElseIf dtPrint.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()

                Dim dtItems As DataTable = clsDBFuncationality.GetDataTable("select Item_Code,max(Sku_Seq) as Sku_Seq,max(Short_Description) as Short_Description,(Unit_code)Unit_code from (" + qry + ") x group by Item_Code,Unit_code order by Sku_Seq,Unit_code")
                Dim FinalQuery As String = " With CTERawData as ( " + qry + "  )" + Environment.NewLine + Environment.NewLine
                For ii As Integer = 1 To dtItems.Rows.Count Step 11
                    If ii > 1 Then
                        FinalQuery += Environment.NewLine + " Union all " + Environment.NewLine
                    End If
                    FinalQuery += " select " + clsCommon.myCstr(ii) + " as Grp , ROW_NUMBER() over (order by (Cust_Code)) As SNo,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtSupplyDate.Value), "dd/MM/yyyy") + "' AS Date, max(Access_officer) as Access_officer,max(Comp_Code1) as Comp_Code1,max(Description) as Description,max(Vehicle_Id) as Vehicle_Id,max(Comp_Name) as Comp_Name,max(Transporter_Name) as Transporter_Name,max(Add1) as Add1,max(City_Code) as City_Code,max(Pincode) as Pincode,max(State) as State,max(Phone1) as Phone1 ,"

                    FinalQuery += "(Cust_Code)Cust_Code,max(Customer_Name)Customer_Name,max(DriverName)DriverName,'' as DMobNo,max(ShiftType)ShiftType,'" & objCommonVar.CurrentUser & "' as UserName, max(Route_No)Route_No,max(Route_Desc) as Route_Desc,sum(ItemNetAmount)ItemNetAmount,max(Document_Date) as Document_Date"
                    For jj As Integer = 1 To 11
                        Dim strJJ As String = clsCommon.myCstr(jj)
                        Dim strICODE As String = ""
                        Dim strIShortDesc As String = ""
                        Dim strIUnitCode As String
                        If (ii + jj - 1) > dtItems.Rows.Count Then
                            strICODE = ""
                            strIShortDesc = ""
                            strIUnitCode = ""
                        Else
                            strICODE = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Item_Code"))
                            strIShortDesc = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Short_Description"))
                            strIUnitCode = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Unit_code"))
                        End If
                        FinalQuery += " ,'" + strICODE + "' as Item_" + strJJ + " ,'" + strIShortDesc + "' as Item_Short_Description_" + strJJ + " ,'" + strIUnitCode + "' as Item_Unit_code_" + strJJ + "
,CEILING(sum(case when Item_Code='" + strICODE + "' and Unit_code = '" + strIUnitCode + "' then Qty else null end )) as ItemQtyCrate_" + strJJ + ""
                    Next
                    If ii > 1 Then
                        FinalQuery += " ,null as Amount"
                    Else
                        FinalQuery += ",sum(ItemNetAmount)Amount"
                    End If
                    FinalQuery += " from (
select xx.*,Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor as QtyStock,TabDefaultUOM.Conversion_Factor ConvFacNo,TabCrateUOM.Conversion_Factor as ConvFacCrate	from CTERawData xx
left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=xx.Item_Code and  TSPL_ITEM_UOM_DETAIL.UOM_Code=xx.Unit_code
left outer join TSPL_ITEM_UOM_DETAIL as TabDefaultUOM on TabDefaultUOM .Item_Code=xx.Item_Code and  TabDefaultUOM .Default_UOM=1
left outer join TSPL_ITEM_UOM_DETAIL as TabCrateUOM on TabCrateUOM.Item_Code=xx.Item_Code and  TabCrateUOM.UOM_Code='Crate' 
) x group by Cust_Code"


                Next
                dtPrint = clsDBFuncationality.GetDataTable(FinalQuery)
                If dtPrint.Rows.Count > 0 Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dtPrint, "rptGatePassBoothSlip", "Gate Pass Booth Slip")
                End If

                frmCRV = Nothing
            End If
            Return False
            Exit Function
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Export(EnumExportTo.Excel)
    End Sub
    Public Sub Export(ByVal exporter As EnumExportTo, Optional ByVal FromDate As DateTime = Nothing, Optional ByVal ToDate As DateTime = Nothing, Optional ByVal strRoute As String = Nothing, Optional ByVal strShift As String = Nothing)
        Try
            Dim ShiftType As String = ""
            Dim qry As String = ""
            Dim itemqry As String = ""
            Dim Freshitem As String = ""
            Dim ProductItem As String = ""

            Freshitem = "Select max(TSPL_ITEM_MASTER.Short_Description)Fresh_Item,max(TSPL_ITEM_MASTER.Item_Desc)Item_Desc,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Print_Sequence)Print_Sequence from TSPL_SD_SHIPMENT_DETAIL
LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code WHERE "
            If clsCommon.myLen(strShift) > 0 Then
                Freshitem += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                    Freshitem += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                End If
                If clsCommon.myLen(strRoute) > 0 Then
                    Freshitem += " And Route_No In (" + strRoute + ") "
                End If
            Else
                Freshitem += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                Freshitem += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                Freshitem += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
            End If
            Freshitem += " and TSPL_SD_SHIPMENT_HEAD.Status=1  "
            Freshitem += " AND ((TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 0 ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 1 and Is_CrateType = 1))
 group by TSPL_ITEM_MASTER.Item_Code ORDER BY Sku_Seq "

            ProductItem = "Select max(TSPL_ITEM_MASTER.Short_Description)Product_item,max(TSPL_ITEM_MASTER.Item_Desc)Item_Desc,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Print_Sequence)Print_Sequence from TSPL_SD_SHIPMENT_DETAIL
LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code WHERE "
            If clsCommon.myLen(strShift) > 0 Then
                ProductItem += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                    ProductItem += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                End If
                If clsCommon.myLen(strRoute) > 0 Then
                    ProductItem += " And Route_No In (" + strRoute + ") "
                End If
            Else
                ProductItem += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                ProductItem += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                ProductItem += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
            End If
            ProductItem += " and TSPL_SD_SHIPMENT_HEAD.Status=1 AND TSPL_ITEM_MASTER.Is_Ambient = 1  group by TSPL_ITEM_MASTER.Item_Code ORDER BY Sku_Seq "



            Dim BaseItemQry As String = "Select TSPL_ITEM_MASTER.Item_Code,max(TSPL_ITEM_MASTER.Short_Description)Short_Description,max(TSPL_ITEM_MASTER.Item_Desc)Item_Description,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Alies_Name2)Alies_Name2,MAX(TSPL_ITEM_MASTER.Alies_Name3)Alies_Name3,Convert(varchar,MAX(TSPL_ITEM_MASTER.Print_Sequence))Print_Sequence,Sum(TSPL_SD_SHIPMENT_DETAIL.Qty)Qty from TSPL_SD_SHIPMENT_DETAIL
LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code Where "
            If clsCommon.myLen(strShift) > 0 Then
                BaseItemQry += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                    BaseItemQry += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                End If
                If clsCommon.myLen(strRoute) > 0 Then
                    BaseItemQry += " And Route_No In (" + strRoute + ") "
                End If
            Else
                BaseItemQry += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                BaseItemQry += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                BaseItemQry += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
            End If
            BaseItemQry += " and TSPL_SD_SHIPMENT_HEAD.Status=1 group by TSPL_ITEM_MASTER.Item_Code " ' ORDER BY Sku_Seq"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                BaseItemQry += " union
 Select TSPL_ITEM_MASTER.Item_Code,max(TSPL_ITEM_MASTER.Short_Description)Short_Description,max(TSPL_ITEM_MASTER.Item_Desc)Item_Description,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Alies_Name2)Alies_Name2,MAX(TSPL_ITEM_MASTER.Alies_Name3)Alies_Name3,
Convert(varchar,MAX(TSPL_ITEM_MASTER.Print_Sequence))Print_Sequence,0 As Qty from TSPL_ITEM_MASTER
LEFT OUTER JOIN TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
WHERE TSPL_ITEM_MASTER.Print_Sequence is not null and TSPL_ITEM_MASTER.Active=1
 group by TSPL_ITEM_MASTER.Item_Code "
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                itemqry = "Select Max(Short_Description)Short_Description,Max(Item_Description)Item_Description,Max(Sku_Seq)Sku_Seq,Max(Alies_Name2)Alies_Name2,Max(Alies_Name3)Alies_Name3,Max(Print_Sequence)Print_Sequence,Sum(Qty)Qty from (" + BaseItemQry + ") xyz Group By Item_Code Order By Sku_Seq"
            Else
                itemqry = "Select * from (" + BaseItemQry + ") xyz Order By Sku_Seq"
            End If


            Dim itemName2 As String = Nothing
            Dim itemName1 As String = Nothing
            Dim itemNames1 As String = Nothing
            Dim itemNames2 As String = Nothing
            Dim itemNamesQty As String = Nothing
            Dim itemNamesAmt As String = Nothing
            Dim FinalItemNamesQty As String = Nothing
            Dim FinalItemNamesAmt As String = Nothing
            Dim ProductIemName As String = Nothing
            Dim FreshItemName As String = Nothing
            Dim FreshItemsName As String = Nothing
            Dim ProductIemsName As String = Nothing
            Dim ProductItemsAmt As String = Nothing
            Dim ItemSubGroup As String = Nothing
            Dim ItemSubGroupAvg As String = Nothing
            Dim ItemsSubGroup As String = Nothing
            Dim itemNamesFresh As String = Nothing
            Dim itemNamesProduct As String = Nothing
            Dim FreshItemNameMax As String = Nothing
            Dim ProductItemNameMax As String = Nothing
            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(itemqry)
            Dim dtFresh As DataTable = clsDBFuncationality.GetDataTable(Freshitem)
            Dim dtProduct As DataTable = clsDBFuncationality.GetDataTable(ProductItem)

            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    'itemName1 += "Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]" + ","
                    itemName2 += "Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "1]" + ","
                    FinalItemNamesQty += "SUM(XXFINAL.[" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]" + ","
                    FinalItemNamesAmt += "SUM(XXFINAL.[" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "1]) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "1]" + ","

                    If i = 0 Then
                        itemNamesQty += "ISNULL([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)"
                        itemNamesAmt += "ISNULL([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) + "],0)"
                        itemNames1 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "] "
                        itemNames2 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) + "] "
                        If clsCommon.myCDecimal(dtitemName.Rows(i)("Qty")) > 0 Then
                            itemName1 += "Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]"
                        Else
                            itemName1 += "0 As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]"
                        End If
                    Else
                        itemNamesQty += "+" + "ISNULL([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)"
                        itemNamesAmt += "+" + "ISNULL([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) + "],0)"
                        itemNames1 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "] "
                        itemNames2 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) + "] "
                        If clsCommon.myCDecimal(dtitemName.Rows(i)("Qty")) > 0 Then
                            itemName1 += ", Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]"
                        Else
                            itemName1 += ", 0 As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]"
                        End If
                    End If
                Next
            End If
            If dtFresh.Rows.Count > 0 Then
                For i As Integer = 0 To dtFresh.Rows.Count - 1
                    FreshItemName += " Sum(IsNull([" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "],0)) As [" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "]" + ","
                    FreshItemNameMax += "max(IsNull([" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "],0)) As [" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "]" + ","
                    If i = 0 Then
                        itemNamesFresh += "ISNULL([" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "],0)"
                        FreshItemsName += "[" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "] "
                    Else
                        itemNamesFresh += "+" + "ISNULL([" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "],0)"
                        FreshItemsName += ", [" + clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) + "] "
                    End If
                Next
            End If
            If dtProduct.Rows.Count > 0 Then
                For i As Integer = 0 To dtProduct.Rows.Count - 1
                    ProductIemName += "Sum(IsNull([" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "],0)) As [" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "]" + ","
                    ProductItemNameMax += " max(IsNull([" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "],0)) As [" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "]" + ","
                    If i = 0 Then
                        itemNamesProduct += "ISNULL([" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "],0)"
                        ProductIemsName += "[" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "] "
                    Else
                        itemNamesProduct += "+" + "ISNULL([" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "],0)"

                        ProductIemsName += ", [" + clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) + "] "
                    End If
                Next
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                qry = "Select max(Customer_Name)OUTLET,max(Display_Seq)as Display_Seq, " & itemName1 & ",sum(ItemNetAmount) as Amount from (select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name,max(XXFinal.Display_Seq) as Display_Seq, max(XXFinal.Short_Description) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount,sum(LTR_QTY)LTR_QTY,sum(KG_QTY)KG_QTY,max(Fresh_Item)Fresh_Item,max(Product_Item)Product_Item

from (select TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code as TR_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name,isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0) as Display_Seq,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,
  Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY ,
  Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY1,
  Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[LTR]) as LTR_QTY,
case when TSPL_ITEM_MASTER.Is_Ambient = 1  then TSPL_ITEM_MASTER.Short_Description  end AS Product_Item,
				case when (TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 0 ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 1 and Is_CrateType = 1) then  TSPL_ITEM_MASTER.Short_Description end as Fresh_Item
    from  TSPL_SD_SHIPMENT_BOOKING_DETAIL
left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrComp_Code1 + "'
  left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code=I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N'  And TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where "
                If clsCommon.myLen(strShift) > 0 Then
                    qry += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                    If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                        qry += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                    End If
                    If clsCommon.myLen(strRoute) > 0 Then
                        qry += " And Route_No In (" + strRoute + ") "
                    End If
                Else
                    qry += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                    qry += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                    qry += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
                End If

                If AllowGatePassDemandTripWise Then
                    qry += " and status=1 and TSPL_SD_SHIPMENT_BOOKING_DETAIL.Trip_No = '" + txtTripNo.Text + "') "
                Else
                    qry += " and status=1 ) "
                End If

                qry += "  Union
select Distinct '' as TR_Code,TSPL_CUSTOMER_MASTER.Cust_Code,
IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name,isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0) as Display_Seq,'' As Item_Code,'' As Short_Description,0 As Sku_Seq,0 As Qty,0 As ItemNetAmount,'' As Unit_code ,'' As ShiftType,TSPL_ROUTE_MASTER.Route_No, TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,0 as KG_QTY , 0 as KG_QTY1, 0 as LTR_QTY,'' AS Product_Item,'' as Fresh_Item
from TSPL_CUSTOMER_MASTER
Left Outer Join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
Left Join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id
LEFT Outer Join (Select Document_No,Document_Date,Route_No,Max(ShiftType)ShiftType from TSPL_DEMAND_BOOKING_MASTER 
--Where CONVERT(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='01-Apr-2024'
Group By Document_No,Document_Date,Route_No )TSPL_DEMAND_BOOKING_MASTER On TSPL_DEMAND_BOOKING_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
Left Outer Join TSPL_DEMAND_BOOKING_DETAIL On TSPL_DEMAND_BOOKING_DETAIL.Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No And TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
Left Join TSPL_SD_SHIPMENT_BOOKING_DETAIL ON TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code=TSPL_DEMAND_BOOKING_DETAIL.TR_Code  
Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = 'UDP'
----left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code=I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N' and TSPL_CUSTOMER_MASTER.Status='N' "
                If clsCommon.myLen(strRoute) > 0 Then
                    qry += " And TSPL_ROUTE_MASTER.Route_No in (" + strRoute + ") "
                Else
                    qry += " And TSPL_ROUTE_MASTER.Route_No='" + clsCommon.myCstr(fndRouteNo.Value) + "' "
                End If
                qry += " )XXFinal group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX "

                If dtFresh.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(LTR_QTY)  For Fresh_Item In (" & FreshItemsName & ") ) As pivot_fresh "
                End If
                If dtProduct.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(KG_QTY)   For Product_Item In (" & ProductIemsName & ") ) As  pivot_Product "
                End If
                qry += " group by Cust_Code "

                qry += " Union all 
                       Select 'TOTAL QNTY' as OUTLET,100000 as Display_Seq ," & itemName1 & " ,sum(Amount) as Amount
from (Select 1 AS Sno,Cust_Code,max(Customer_Name)Customer_Name,max(Display_Seq)as Display_Seq, " & itemName1 & " ,sum(ItemNetAmount) as Amount from (select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name,max(XXFinal.Display_Seq) as Display_Seq, max(XXFinal.Short_Description) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount,sum(LTR_QTY)LTR_QTY,sum(KG_QTY)KG_QTY,max(Fresh_Item)Fresh_Item,max(Product_Item)Product_Item

from (select TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code as TR_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name, isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0)as Display_Seq, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,
Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY ,
  Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY1,
  Convert(Decimal(18,2),(isnull(TSPL_SD_SHIPMENT_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[LTR]) as LTR_QTY,
case when TSPL_ITEM_MASTER.Is_Ambient = 1  then TSPL_ITEM_MASTER.Short_Description  end AS Product_Item,
				case when (TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 0 ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 1 and Is_CrateType = 1) then  TSPL_ITEM_MASTER.Short_Description end as Fresh_Item
    from  TSPL_SD_SHIPMENT_BOOKING_DETAIL
left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrComp_Code1 + "'
  left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code = I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N'  And TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where "
                If clsCommon.myLen(strShift) > 0 Then
                    qry += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                    If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                        qry += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                    End If
                    If clsCommon.myLen(strRoute) > 0 Then
                        qry += " And Route_No In (" + strRoute + ") "
                    End If
                Else
                    qry += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                    qry += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                    qry += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
                End If
                'qry += "and status=1 ) )XXFinal group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX "
                If AllowGatePassDemandTripWise Then
                    qry += " and status=1 and TSPL_SD_SHIPMENT_BOOKING_DETAIL.Trip_No = '" + txtTripNo.Text + "' ) )XXFinal group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX  "
                Else
                    qry += " and status=1 ) )XXFinal group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX "
                End If

                If dtFresh.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(LTR_QTY)  For Fresh_Item In (" & FreshItemsName & ") ) As pivot_fresh "
                End If
                If dtProduct.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(KG_QTY)   For Product_Item In (" & ProductIemsName & ") ) As  pivot_Product "
                End If

                qry += " group by Cust_Code )XX group by SNo "

                qry = "select Convert(Varchar,ROW_NUMBER() Over (Order By (Select 1))) As [SR.],* from (" + qry + ") XXXFinal order by Display_Seq "
            Else
                qry = "select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name, max(XXFinal.Short_Description) +' '+max(XXFinal.Unit_code) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount
from (select TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code as TR_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No from  TSPL_SD_SHIPMENT_BOOKING_DETAIL
left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.TR_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" + objCommonVar.CurrComp_Code1 + "'

where TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE in (select document_Code from TSPL_SD_SHIPMENT_HEAD where"
                If clsCommon.myLen(strShift) > 0 Then
                    qry += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(FromDate, "dd-MMM-yyyy") + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(ToDate, "dd-MMM-yyyy") + "' "
                    If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then
                        qry += " and Shift_Type = '" + IIf(strShift = "Morning", "AM", "PM") + "' "
                    End If
                    If clsCommon.myLen(strRoute) > 0 Then
                        qry += " And Route_No In (" + strRoute + ") "
                    End If
                Else
                    qry += " convert(date,Supply_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd-MMM-yyyy") + "' "
                    qry += " and Shift_Type = '" + IIf(rbtnMorning.IsChecked, "AM", "PM") + "' "
                    qry += " And Route_No = '" + clsCommon.myCstr(fndRouteNo.Value) + "' "
                End If
                qry += " )XXFinal group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code "
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            'If clsCommon.CompairString(strShift, "Both") <> CompairStringResult.Equal Then

            'End If
            If clsCommon.myLen(strShift) > 0 Then
                ShiftType = strShift
            Else
                If rbtnMorning.IsChecked Then
                    ShiftType = "Morning"
                Else
                    ShiftType = "Evening"
                End If
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Dim dr As DataRow = dt.NewRow
                If clsCommon.myLen(strShift) > 0 Then
                    dr("OUTLET") = ShiftType
                Else
                    dr("OUTLET") = clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MM/yyyy") + ", " + ShiftType
                End If
                For ii As Integer = 0 To dtitemName.Rows.Count - 1
                    'dr(clsCommon.myCstr(dtitemName.Rows(ii)("Short_Description"))) = clsCommon.myCstr(dtitemName.Rows(ii)("Print_Sequence"))
                    'Dim dr As DataRow = dt.NewRow
                    Dim colName As String = clsCommon.myCstr(dtitemName.Rows(ii)("Short_Description"))
                    Dim value As Decimal = clsCommon.myCDecimal(dtitemName.Rows(ii)("Print_Sequence"))

                    ' Check if value is numeric before assigning
                    If IsNumeric(value) AndAlso value > 0 Then
                        dr(colName) = clsCommon.myCDecimal(value)
                    Else
                        dr(colName) = DBNull.Value ' Or handle accordingly
                    End If
                Next
                dt.Rows.InsertAt(dr, 0)
                dt.AcceptChanges()
            End If


            MyRadGridView1.DataSource = Nothing
            MyRadGridView1.Rows.Clear()
            MyRadGridView1.Columns.Clear()
            MyRadGridView1.GroupDescriptors.Clear()
            MyRadGridView1.MasterView.Refresh()
            MyRadGridView1.GroupDescriptors.Clear()
            MyRadGridView1.EnableFiltering = True
            MyRadGridView1.MasterTemplate.SummaryRowsBottom.Clear()
            If dt.Rows.Count > 0 Then

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    ' Create a new DataTable to store converted data
                    Dim dtConverted As New DataTable()
                    ' Convert all columns to String type
                    For Each col As DataColumn In dt.Columns
                        If clsCommon.CompairString(clsCommon.myCstr(col.ColumnName), "Display_Seq") = CompairStringResult.Equal Then
                            Continue For
                        End If
                        dtConverted.Columns.Add(col.ColumnName, GetType(String))
                    Next
                    ' Copy data with replacements
                    For Each row As DataRow In dt.Rows
                        Dim newRow As DataRow = dtConverted.NewRow()

                        For Each col As DataColumn In dt.Columns
                            If clsCommon.CompairString(clsCommon.myCstr(col.ColumnName), "Display_Seq") = CompairStringResult.Equal Then
                                Continue For
                            End If
                            Dim cellValue As Object = row(col)
                            ' If numeric, check for 0 and replace
                            If IsNumeric(cellValue) Then
                                If Convert.ToDouble(cellValue) = 0 Then
                                    newRow(col.ColumnName) = "-" ' Replace 0 with "-"
                                Else
                                    newRow(col.ColumnName) = cellValue.ToString() ' Convert to string
                                End If
                            Else
                                newRow(col.ColumnName) = cellValue.ToString() ' Convert non-numeric to string
                            End If
                        Next
                        dtConverted.Rows.Add(newRow)
                    Next
                    ' Bind the converted DataTable to RadGridView
                    MyRadGridView1.DataSource = dtConverted
                Else
                    MyRadGridView1.DataSource = dt
                End If
                MyRadGridView1.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.None
                MyRadGridView1.MasterTemplate.Refresh()


                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    ApplyFormattingManually()
                    For i As Integer = 0 To dtitemName.Rows.Count - 1
                        MyRadGridView1.Columns("" + clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) + "").FormatString = "{0:n2}"
                        If clsCommon.myLen(clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name2"))) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name3"))) > 0 Then
                            MyRadGridView1.Columns("" + clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) + "").HeaderText = clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name2")) + Environment.NewLine + clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name3"))
                        Else
                            MyRadGridView1.Columns("" + clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) + "").HeaderText = clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description"))
                        End If
                        MyRadGridView1.Columns("" + clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) + "").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    Next
                    MyRadGridView1.Columns("Amount").FormatString = "{0:n2}"
                End If
                'View()
                ' SetGridFormation()
                'ReStoreGridLayout()
                MyRadGridView1.MasterTemplate.AutoExpandGroups = True
                'RadPageView1.SelectedPage = RadPageViewPage2
                MyRadGridView1.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            Dim arrHeader As List(Of String) = New List(Of String)()
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                arrHeader.Add("Supply Chart")
                If clsCommon.myLen(strShift) > 0 Then
                    arrHeader.Add("Date: " + IIf(clsCommon.myLen(strShift) > 0, clsCommon.GetPrintDate(FromDate, "dd/MM/yyyy") + " to " + clsCommon.GetPrintDate(ToDate, "dd/MM/yyyy") + "   " + strShift, clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MM/yyyy") + "   " + ShiftType) + "    " + "Route :" + IIf(clsCommon.myLen(strRoute) > 0, strRoute, clsCommon.myCstr(txtRouteName.Text)) + "    ")
                Else
                    arrHeader.Add("Transpoter : " + clsCommon.myCstr(txtDistributorName.Text) + "" + "     Date: " + IIf(clsCommon.myLen(strShift) > 0, clsCommon.GetPrintDate(FromDate, "dd/MM/yyyy") + " to " + clsCommon.GetPrintDate(ToDate, "dd/MM/yyyy") + "   " + strShift, clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MM/yyyy") + "   " + ShiftType) + "    " + "Route :" + IIf(clsCommon.myLen(strRoute) > 0, strRoute, clsCommon.myCstr(txtRouteName.Text)) + "    ")
                End If

                'arrHeader.Add(("Vehicle No: " + clsCommon.myCstr(lblVehicleDesc.Text) + "  "))
                'arrHeader.Add(("Shift Type: " + ShiftType + "  "))
                'arrHeader.Add(("Driver No: " + clsCommon.myCstr(txtDriverMobNo.Text) + "  "))
            Else
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptBookingQtyAmtReport & "'"))
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add(("Date: " + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MM/yyyy") + "  "))
                arrHeader.Add(("Route No: " + clsCommon.myCstr(fndRouteNo.Value) + "  "))
                arrHeader.Add(("Route Name: " + clsCommon.myCstr(txtRouteName.Text) + "  "))
                arrHeader.Add(("Vehicle No: " + clsCommon.myCstr(lblVehicleDesc.Text) + "  "))
                arrHeader.Add(("Shift Type: " + ShiftType + "  "))

                arrHeader.Add(("Transpoter Name: " + clsCommon.myCstr(txtDistributorName.Text) + "  "))
                arrHeader.Add(("Vehicle_No: " + clsCommon.myCstr(lblVehicleDesc.Text) + "  "))
                arrHeader.Add(("Driver: " + clsCommon.myCstr(txtDriverName.Text) + "  "))
                arrHeader.Add(("Driver No: " + clsCommon.myCstr(txtDriverMobNo.Text) + "  "))

            End If

            'If txtLocation.arrDispalyMember IsNot Nothing AndAlso txtLocation.arrDispalyMember.Count > 0 Then
            '    arrHeader.Add("Location : " + clsCommon.GetMulcallStringWithComma(txtLocation.arrDispalyMember))
            'End If
            'If txtItem.arrDispalyMember IsNot Nothing AndAlso txtItem.arrDispalyMember.Count > 0 Then
            '    arrHeader.Add("Item : " + clsCommon.GetMulcallStringWithComma(txtItem.arrDispalyMember))
            'End If
            'If txtTransaction.arrDispalyMember IsNot Nothing AndAlso txtTransaction.arrDispalyMember.Count > 0 Then
            '    arrHeader.Add("Transaction : " + clsCommon.GetMulcallStringWithComma(txtTransaction.arrDispalyMember))
            'End If
            'If exporter = EnumExportTo.Excel Then
            '    transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
            '    'clsCommon.MyExportToExcelGrid("Demand Booking Report", Gv1, arrHeader, Me.Text)
            '    transportSql.exportdata(Gv1, "GATEPASS DETAIL", Me.Text, False, arrHeader, False, False, False)
            'Else
            '    transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
            '    'clsCommon.MyExportToPDF("Demand Booking Report", Gv1, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
            '    clsCommon.MyExportToPDF(Me.Text, Gv1, arrHeader, Me.Text, True)
            'End If
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBoothData())
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    transportSql.exportdataBoothSlipGNG(Nothing, MyRadGridView1, "", "Dairy GatePass Booth Slip", 0, MyRadGridView1.Rows.Count, False, arrHeader, False, False, False, False, False, Nothing, True, True)
                Else
                    clsCommon.MyExportToExcelGrid("Dairy GatePass Booth Slip", MyRadGridView1, arrHeader, "Dairy GatePass Booth Slip")
                End If


                ' transportSql.ExporttoExcel(dt, Me)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Public Sub ApplyFormattingManually()
        For Each row As GridViewRowInfo In MyRadGridView1.Rows
            For Each cell As GridViewCellInfo In row.Cells
                If cell IsNot Nothing AndAlso IsNumeric(cell.Value) Then
                    Dim value As Double = clsCommon.myCdbl(cell.Value)
                    'Dim cellElement As GridCellElement = MyRadGridView1.TableElement.GetCellElement(cell.RowInfo, cell.ColumnInfo)
                    If value > 0 Then
                        ' Apply formatting directly
                        If value = Math.Floor(value) Then
                            cell.Value = value.ToString("0") ' No decimals
                        ElseIf value * 10 = Math.Floor(value * 10) Then
                            cell.Value = value.ToString("0.0") ' One decimal place
                        Else
                            cell.Value = value.ToString("0.00") ' Two decimal places
                        End If
                    End If
                End If
            Next
        Next
        ' Add CellFormatting event to make the first row bold
        'AddHandler MyRadGridView1.CellFormatting, AddressOf MyRadGridView1_CellFormatting
        '' Force UI refresh
        'MyRadGridView1.TableElement.Update(False)
        'MyRadGridView1.MasterTemplate.Refresh()
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtCode.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtCode.Value, "GPCode", "TSPL_DAIRYSALE_GATEPASS_MASTER", "TSPL_DAIRYSALE_GATEPASS_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub RadMenuItem4_Click(sender As Object, e As EventArgs) Handles RadMenuItem4.Click
        funPrint(txtCode.Value, True)
    End Sub

    Private Sub chkIndividualCustomer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkIndividualCustomer.ToggleStateChanged
        If chkIndividualCustomer.Checked Then
            txtDemandNo.Enabled = True
        Else
            txtDemandNo.Enabled = False
        End If
    End Sub

    Private Sub txtDemandNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDemandNo._MYValidating
        Try
            Dim strQry As String = "select distinct TSPL_DEMAND_BOOKING_MASTER.Demand_UniqueID as Code,TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_DEMAND_BOOKING_MASTER.Location_Code,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code as [Booth Code],TSPL_CUSTOMER_MASTER.Customer_Name from TSPL_DEMAND_BOOKING_MASTER left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code "
            Dim whrcls As String = "  convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='" + clsCommon.GetPrintDate(txtSupplyDate.Value, "dd/MMM/yyyy") + "' and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=1 and TSPL_DEMAND_BOOKING_MASTER.Posted=1 and TSPL_DEMAND_BOOKING_MASTER.Demand_UniqueID not in(select Demand_UniqueID from TSPL_DAIRYSALE_GATEPASS_MASTER where Demand_UniqueID is not null )"
            txtDemandNo.Value = clsCommon.ShowSelectForm("DemandSearch", strQry, "Code", whrcls, txtDemandNo.Value, "Code", isButtonClicked)
            fndRouteNo.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_DEMAND_BOOKING_MASTER where Demand_UniqueID='" + txtDemandNo.Value + "'"))
            txtLocCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_DEMAND_BOOKING_MASTER where Demand_UniqueID='" + txtDemandNo.Value + "'"))
            Dim shiftType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ShiftType from TSPL_DEMAND_BOOKING_MASTER where Demand_UniqueID='" + txtDemandNo.Value + "'"))
            If clsCommon.CompairString(shiftType, "Morning") = CompairStringResult.Equal Then
                rbtnMorning.IsChecked = True
            ElseIf clsCommon.CompairString(shiftType, "Evening") = CompairStringResult.Equal Then
                rbtnEvening.IsChecked = True
            End If
            setRouteVehicleDetail()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub ControlFields(ByVal flag As Boolean)
        txtLocCode.Enabled = flag
        txtVehicle.Enabled = flag
        fndRouteNo.Enabled = flag
        txtDemandNo.Enabled = flag
        chkIndividualCustomer.Enabled = flag
        txtSupplyDate.Enabled = flag
        RadGroupBox3.Enabled = flag
        rgbItemType.Enabled = flag
        txtTripNo.Enabled = flag
    End Sub

    Private Sub rbtn_product_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtn_product.ToggleStateChanged
        If rbtn_product.IsChecked OrElse rbtn_IceCream.IsChecked Then
            rbtnMorning.IsChecked = True
            RadGroupBox3.Enabled = False
            RadGroupBox3.Visible = False
        End If
    End Sub

    Private Sub rbtn_IceCream_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtn_IceCream.ToggleStateChanged
        If rbtn_product.IsChecked OrElse rbtn_IceCream.IsChecked Then
            rbtnMorning.IsChecked = True
            RadGroupBox3.Enabled = False
            RadGroupBox3.Visible = False

        End If
    End Sub

    Private Sub rbtn_Milk_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtn_Milk.ToggleStateChanged
        If rbtn_Milk.IsChecked Then
            RadGroupBox3.Enabled = True
            RadGroupBox3.Visible = True

        End If
    End Sub
    Private Sub CrateCount_CellValue()
        Try
            If DifferentCrateTypeForFGItem Then
                Dim lstDRobj As New List(Of clsDRDetail)
                For ii As Integer = 0 To Gv1.Rows.Count - 1
                    Dim DRobj As clsDRDetail = New clsDRDetail()
                    DRobj.Item_Code = clsCommon.myCstr(Gv1.Rows(ii).Cells(colItemCode).Value)
                    DRobj.Unit_Code = clsCommon.myCstr(Gv1.Rows(ii).Cells(colUnit).Value)
                    DRobj.Crate = clsCommon.myCdbl(Gv1.Rows(ii).Cells(colCrateIssue).Value)
                    DRobj.Qty = clsCommon.myCdbl(Gv1.Rows(ii).Cells(colQty).Value)

                    lstDRobj.Add(DRobj)
                Next
                LoadgvCrateType()
                Dim lstCTstr As List(Of String) = New List(Of String)
                Dim ctintRow As Integer = 0
                Dim groupbyItem = From i In lstDRobj
                                  Group By i.Item_Code, i.Unit_Code Into Group
                                  Select New With {
                    Key .Item = Item_Code,
                    Key .Unit = Unit_Code,
                    Key .TotalQty = Group.Sum(Function(x) x.Crate)
                }
                For Each result In groupbyItem
                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(result.Item) & "'"))
                    If ItemCrateType = 1 Then
                        Dim strCtCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select CrateType_Item  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(result.Item) & "'"))
                        If Not lstCTstr.Contains(strCtCode) Then
                            ctintRow += 1
                            gvCrateType.Rows.AddNew()
                            gvCrateType.Rows(ctintRow - 1).Cells(colCTICode).Value = clsCommon.myCstr(result.Item)
                            gvCrateType.Rows(ctintRow - 1).Cells(colCTCode).Value = strCtCode
                            gvCrateType.Rows(ctintRow - 1).Cells(ColCTName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Item_Desc from TSPL_ITEM_MASTER where Item_Code='" & strCtCode & "'"))
                            gvCrateType.Rows(ctintRow - 1).Cells(ColCTQty).Value = clsCommon.myCdbl(result.TotalQty)
                            lstCTstr.Add(strCtCode)
                        Else
                            For intinnerRow As Integer = 0 To gvCrateType.Rows.Count - 1
                                If clsCommon.CompairString(strCtCode, clsCommon.myCstr(gvCrateType.Rows(intinnerRow).Cells(colCTCode).Value)) = CompairStringResult.Equal Then
                                    gvCrateType.Rows(intinnerRow).Cells(ColCTQty).Value += clsCommon.myCdbl(result.TotalQty)
                                End If
                            Next
                        End If
                    End If
                Next
                UpdateCTQty()
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Sub txtDate_Validating(sender As Object, e As ComponentModel.CancelEventArgs) Handles txtDate.Validating
        Try
            CheckTimeShift()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class
Public Class clsDRDetail
    Public Item_Code As String
    Public Unit_Code As String
    Public Qty As Decimal
    Public Crate As Decimal

End Class
