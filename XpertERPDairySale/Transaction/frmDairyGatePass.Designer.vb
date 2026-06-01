<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDairyGatePass
    Inherits FrmMainTranScreen

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim RadListDataItem1 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem2 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem3 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim TableViewDefinition2 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition3 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.rpvpGatePass = New Telerik.WinControls.UI.RadPageViewPage()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rdBoth = New Telerik.WinControls.UI.RadRadioButton()
        Me.rdbNonTaxable = New Telerik.WinControls.UI.RadRadioButton()
        Me.rdbTaxable = New Telerik.WinControls.UI.RadRadioButton()
        Me.txtShipToLocation = New common.Controls.MyTextBox()
        Me.lblShipToLocation = New common.Controls.MyLabel()
        Me.lblDemandNo = New common.Controls.MyLabel()
        Me.txtDemandNo = New common.UserControls.txtFinder()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.MyLabel8 = New common.Controls.MyLabel()
        Me.lblOpKM = New common.Controls.MyLabel()
        Me.txtOpKM = New common.MyNumBox()
        Me.txtClKM = New common.MyNumBox()
        Me.lblClKM = New common.Controls.MyLabel()
        Me.chkIndividualCustomer = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkGhee = New Telerik.WinControls.UI.RadCheckBox()
        Me.rgbItemType = New Telerik.WinControls.UI.RadGroupBox()
        Me.rbtn_IceCream = New Telerik.WinControls.UI.RadRadioButton()
        Me.rbtn_product = New Telerik.WinControls.UI.RadRadioButton()
        Me.rbtn_Milk = New Telerik.WinControls.UI.RadRadioButton()
        Me.lblClosingDate = New common.Controls.MyLabel()
        Me.txtTripNo = New common.Controls.MyTextBox()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.lblTripNo = New common.Controls.MyLabel()
        Me.txtSupplyDate = New common.Controls.MyDateTimePicker()
        Me.lblSupplyDate = New common.Controls.MyLabel()
        Me.btnClKM = New Telerik.WinControls.UI.RadButton()
        Me.txtDistributorName = New common.Controls.MyTextBox()
        Me.lblDistributorName = New common.Controls.MyLabel()
        Me.txtDriverMobNo = New common.Controls.MyTextBox()
        Me.MyLabel14 = New common.Controls.MyLabel()
        Me.txtDriverName = New common.Controls.MyTextBox()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.txtGatepassDate = New common.Controls.MyDateTimePicker()
        Me.MyLabel12 = New common.Controls.MyLabel()
        Me.txtLoadingSlip = New common.Controls.MyTextBox()
        Me.MyLabel11 = New common.Controls.MyLabel()
        Me.RadGroupBox3 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rbtnEvening = New Telerik.WinControls.UI.RadRadioButton()
        Me.rbtnMorning = New Telerik.WinControls.UI.RadRadioButton()
        Me.btnMultiGPReverse = New Telerik.WinControls.UI.RadButton()
        Me.TxtMultiDairyGPassReverse = New common.UserControls.txtMultiSelectFinder()
        Me.MyLabel34 = New common.Controls.MyLabel()
        Me.lblTollAmount = New common.Controls.MyLabel()
        Me.txtTollAmount = New common.MyNumBox()
        Me.MyLabel9 = New common.Controls.MyLabel()
        Me.FndTransferNo = New common.UserControls.txtFinder()
        Me.chkAgainstTransfer = New Telerik.WinControls.UI.RadCheckBox()
        Me.txtSalesman = New common.Controls.MyTextBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.UsLock1 = New common.usLock()
        Me.MyLabel6 = New common.Controls.MyLabel()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.txtmultiBooking = New common.UserControls.txtMultiSelectFinder()
        Me.lblTotalCan = New common.Controls.MyLabel()
        Me.txtCrateQty = New common.Controls.MyTextBox()
        Me.txtRouteName = New common.Controls.MyTextBox()
        Me.lblSalesman = New common.Controls.MyLabel()
        Me.txtCanQty = New common.Controls.MyTextBox()
        Me.lblRoute = New common.Controls.MyLabel()
        Me.fndRouteNo = New common.UserControls.txtFinder()
        Me.btnGo = New Telerik.WinControls.UI.RadButton()
        Me.txtLocDesc = New common.Controls.MyTextBox()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.txtLocCode = New common.UserControls.txtFinder()
        Me.txtComments = New common.Controls.MyTextBox()
        Me.txtRemarks = New common.Controls.MyTextBox()
        Me.MyLabel3 = New common.Controls.MyLabel()
        Me.txtTransporter = New common.Controls.MyTextBox()
        Me.lblVehicleDesc = New common.Controls.MyTextBox()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.lblpaymentpostdate = New common.Controls.MyLabel()
        Me.cmbitemtype = New common.Controls.MyComboBox()
        Me.lblfullempty = New common.Controls.MyLabel()
        Me.lblpaymentno = New common.Controls.MyLabel()
        Me.btnNew = New Telerik.WinControls.UI.RadButton()
        Me.txtVehicle = New common.UserControls.txtFinder()
        Me.txtCode = New common.UserControls.txtNavigator()
        Me.Gv1 = New common.UserControls.MyRadGridView()
        Me.MyRadGridView1 = New common.UserControls.MyRadGridView()
        Me.rpvpCrateType = New Telerik.WinControls.UI.RadPageViewPage()
        Me.gvCrateType = New common.UserControls.MyRadGridView()
        Me.btnHistory = New Telerik.WinControls.UI.RadButton()
        Me.btnBoothSlip = New Telerik.WinControls.UI.RadSplitButton()
        Me.RadMenuItem3 = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnExcel = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnPrint2 = New Telerik.WinControls.UI.RadButton()
        Me.btnGPCancel = New Telerik.WinControls.UI.RadButton()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.btnReverse = New Telerik.WinControls.UI.RadButton()
        Me.btnDelete = New Telerik.WinControls.UI.RadButton()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.RadSplitButton1 = New Telerik.WinControls.UI.RadSplitButton()
        Me.RadMenuItem1 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem2 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem4 = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnPost = New Telerik.WinControls.UI.RadButton()
        Me.btnSelect = New Telerik.WinControls.UI.RadButton()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.RadMenu1 = New Telerik.WinControls.UI.RadMenu()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.rpvpGatePass.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.rdBoth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbNonTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtShipToLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShipToLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDemandNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblOpKM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOpKM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtClKM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblClKM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkIndividualCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkGhee, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rgbItemType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.rgbItemType.SuspendLayout()
        CType(Me.rbtn_IceCream, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtn_product, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtn_Milk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblClosingDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTripNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTripNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSupplyDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSupplyDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClKM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDistributorName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDistributorName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDriverMobNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDriverName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtGatepassDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLoadingSlip, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox3.SuspendLayout()
        CType(Me.rbtnEvening, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnMorning, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnMultiGPReverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel34, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTollAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTollAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAgainstTransfer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSalesman, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotalCan, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCrateQty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRouteName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSalesman, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCanQty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRoute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLocDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtComments, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRemarks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTransporter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVehicleDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblpaymentpostdate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbitemtype, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblfullempty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblpaymentno, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnNew, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Gv1.SuspendLayout()
        CType(Me.MyRadGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyRadGridView1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.rpvpCrateType.SuspendLayout()
        CType(Me.gvCrateType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvCrateType.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnBoothSlip, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGPCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 20)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadPageView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnHistory)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnBoothSlip)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnGPCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.MyLabel10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReverse)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadSplitButton1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSelect)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(1171, 436)
        Me.SplitContainer1.SplitterDistance = 400
        Me.SplitContainer1.TabIndex = 3
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.rpvpGatePass)
        Me.RadPageView1.Controls.Add(Me.rpvpCrateType)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.rpvpGatePass
        Me.RadPageView1.Size = New System.Drawing.Size(1171, 400)
        Me.RadPageView1.TabIndex = 2
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).ItemAlignment = Telerik.WinControls.UI.StripViewItemAlignment.Near
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripAlignment = Telerik.WinControls.UI.StripViewAlignment.Top
        '
        'rpvpGatePass
        '
        Me.rpvpGatePass.Controls.Add(Me.SplitContainer2)
        Me.rpvpGatePass.ItemSize = New System.Drawing.SizeF(66.0!, 26.0!)
        Me.rpvpGatePass.Location = New System.Drawing.Point(10, 35)
        Me.rpvpGatePass.Name = "rpvpGatePass"
        Me.rpvpGatePass.Size = New System.Drawing.Size(1150, 354)
        Me.rpvpGatePass.Text = "GatePass"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Gv1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1150, 354)
        Me.SplitContainer2.SplitterDistance = 202
        Me.SplitContainer2.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadGroupBox1)
        Me.Panel1.Controls.Add(Me.txtShipToLocation)
        Me.Panel1.Controls.Add(Me.lblShipToLocation)
        Me.Panel1.Controls.Add(Me.lblDemandNo)
        Me.Panel1.Controls.Add(Me.txtDemandNo)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.chkIndividualCustomer)
        Me.Panel1.Controls.Add(Me.chkGhee)
        Me.Panel1.Controls.Add(Me.rgbItemType)
        Me.Panel1.Controls.Add(Me.lblClosingDate)
        Me.Panel1.Controls.Add(Me.txtTripNo)
        Me.Panel1.Controls.Add(Me.lblTripNo)
        Me.Panel1.Controls.Add(Me.txtSupplyDate)
        Me.Panel1.Controls.Add(Me.btnClKM)
        Me.Panel1.Controls.Add(Me.lblSupplyDate)
        Me.Panel1.Controls.Add(Me.txtDistributorName)
        Me.Panel1.Controls.Add(Me.lblDistributorName)
        Me.Panel1.Controls.Add(Me.txtDriverMobNo)
        Me.Panel1.Controls.Add(Me.MyLabel14)
        Me.Panel1.Controls.Add(Me.txtDriverName)
        Me.Panel1.Controls.Add(Me.MyLabel13)
        Me.Panel1.Controls.Add(Me.txtGatepassDate)
        Me.Panel1.Controls.Add(Me.MyLabel12)
        Me.Panel1.Controls.Add(Me.txtLoadingSlip)
        Me.Panel1.Controls.Add(Me.MyLabel11)
        Me.Panel1.Controls.Add(Me.RadGroupBox3)
        Me.Panel1.Controls.Add(Me.btnMultiGPReverse)
        Me.Panel1.Controls.Add(Me.TxtMultiDairyGPassReverse)
        Me.Panel1.Controls.Add(Me.MyLabel34)
        Me.Panel1.Controls.Add(Me.lblTollAmount)
        Me.Panel1.Controls.Add(Me.txtTollAmount)
        Me.Panel1.Controls.Add(Me.MyLabel9)
        Me.Panel1.Controls.Add(Me.FndTransferNo)
        Me.Panel1.Controls.Add(Me.chkAgainstTransfer)
        Me.Panel1.Controls.Add(Me.txtSalesman)
        Me.Panel1.Controls.Add(Me.MyLabel7)
        Me.Panel1.Controls.Add(Me.UsLock1)
        Me.Panel1.Controls.Add(Me.MyLabel6)
        Me.Panel1.Controls.Add(Me.MyLabel5)
        Me.Panel1.Controls.Add(Me.txtmultiBooking)
        Me.Panel1.Controls.Add(Me.lblTotalCan)
        Me.Panel1.Controls.Add(Me.txtCrateQty)
        Me.Panel1.Controls.Add(Me.txtRouteName)
        Me.Panel1.Controls.Add(Me.txtCanQty)
        Me.Panel1.Controls.Add(Me.lblRoute)
        Me.Panel1.Controls.Add(Me.fndRouteNo)
        Me.Panel1.Controls.Add(Me.btnGo)
        Me.Panel1.Controls.Add(Me.txtLocDesc)
        Me.Panel1.Controls.Add(Me.MyLabel4)
        Me.Panel1.Controls.Add(Me.txtLocCode)
        Me.Panel1.Controls.Add(Me.txtComments)
        Me.Panel1.Controls.Add(Me.txtRemarks)
        Me.Panel1.Controls.Add(Me.txtTransporter)
        Me.Panel1.Controls.Add(Me.lblVehicleDesc)
        Me.Panel1.Controls.Add(Me.MyLabel2)
        Me.Panel1.Controls.Add(Me.MyLabel3)
        Me.Panel1.Controls.Add(Me.MyLabel1)
        Me.Panel1.Controls.Add(Me.txtDate)
        Me.Panel1.Controls.Add(Me.cmbitemtype)
        Me.Panel1.Controls.Add(Me.lblSalesman)
        Me.Panel1.Controls.Add(Me.lblfullempty)
        Me.Panel1.Controls.Add(Me.lblpaymentno)
        Me.Panel1.Controls.Add(Me.btnNew)
        Me.Panel1.Controls.Add(Me.txtVehicle)
        Me.Panel1.Controls.Add(Me.txtCode)
        Me.Panel1.Controls.Add(Me.lblpaymentpostdate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1150, 202)
        Me.Panel1.TabIndex = 0
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Controls.Add(Me.rdBoth)
        Me.RadGroupBox1.Controls.Add(Me.rdbNonTaxable)
        Me.RadGroupBox1.Controls.Add(Me.rdbTaxable)
        Me.RadGroupBox1.HeaderText = ""
        Me.RadGroupBox1.Location = New System.Drawing.Point(810, 112)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Size = New System.Drawing.Size(71, 60)
        Me.RadGroupBox1.TabIndex = 1463
        Me.RadGroupBox1.Visible = False
        '
        'rdBoth
        '
        Me.rdBoth.Location = New System.Drawing.Point(4, 41)
        Me.rdBoth.Name = "rdBoth"
        Me.rdBoth.Size = New System.Drawing.Size(44, 18)
        Me.rdBoth.TabIndex = 2
        Me.rdBoth.TabStop = False
        Me.rdBoth.Text = "Both"
        '
        'rdbNonTaxable
        '
        Me.rdbNonTaxable.Location = New System.Drawing.Point(4, 24)
        Me.rdbNonTaxable.Name = "rdbNonTaxable"
        Me.rdbNonTaxable.Size = New System.Drawing.Size(35, 18)
        Me.rdbNonTaxable.TabIndex = 1
        Me.rdbNonTaxable.TabStop = False
        Me.rdbNonTaxable.Text = "NT"
        '
        'rdbTaxable
        '
        Me.rdbTaxable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rdbTaxable.Location = New System.Drawing.Point(4, 4)
        Me.rdbTaxable.Name = "rdbTaxable"
        Me.rdbTaxable.Size = New System.Drawing.Size(26, 18)
        Me.rdbTaxable.TabIndex = 0
        Me.rdbTaxable.TabStop = False
        Me.rdbTaxable.Text = "T"
        Me.rdbTaxable.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'txtShipToLocation
        '
        Me.txtShipToLocation.CalculationExpression = Nothing
        Me.txtShipToLocation.FieldCode = Nothing
        Me.txtShipToLocation.FieldDesc = Nothing
        Me.txtShipToLocation.FieldMaxLength = 0
        Me.txtShipToLocation.FieldName = Nothing
        Me.txtShipToLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipToLocation.isCalculatedField = False
        Me.txtShipToLocation.IsSourceFromTable = False
        Me.txtShipToLocation.IsSourceFromValueList = False
        Me.txtShipToLocation.IsUnique = False
        Me.txtShipToLocation.Location = New System.Drawing.Point(431, 112)
        Me.txtShipToLocation.MaxLength = 200
        Me.txtShipToLocation.MendatroryField = False
        Me.txtShipToLocation.MyLinkLable1 = Me.lblShipToLocation
        Me.txtShipToLocation.MyLinkLable2 = Nothing
        Me.txtShipToLocation.Name = "txtShipToLocation"
        Me.txtShipToLocation.ReferenceFieldDesc = Nothing
        Me.txtShipToLocation.ReferenceFieldName = Nothing
        Me.txtShipToLocation.ReferenceTableName = Nothing
        Me.txtShipToLocation.Size = New System.Drawing.Size(253, 18)
        Me.txtShipToLocation.TabIndex = 1461
        '
        'lblShipToLocation
        '
        Me.lblShipToLocation.FieldName = Nothing
        Me.lblShipToLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShipToLocation.Location = New System.Drawing.Point(337, 113)
        Me.lblShipToLocation.Name = "lblShipToLocation"
        Me.lblShipToLocation.Size = New System.Drawing.Size(91, 16)
        Me.lblShipToLocation.TabIndex = 1462
        Me.lblShipToLocation.Text = "Ship To Location"
        '
        'lblDemandNo
        '
        Me.lblDemandNo.FieldName = Nothing
        Me.lblDemandNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDemandNo.Location = New System.Drawing.Point(509, 53)
        Me.lblDemandNo.Name = "lblDemandNo"
        Me.lblDemandNo.Size = New System.Drawing.Size(67, 16)
        Me.lblDemandNo.TabIndex = 1460
        Me.lblDemandNo.Text = "Demand No"
        '
        'txtDemandNo
        '
        Me.txtDemandNo.CalculationExpression = Nothing
        Me.txtDemandNo.FieldCode = Nothing
        Me.txtDemandNo.FieldDesc = Nothing
        Me.txtDemandNo.FieldMaxLength = 0
        Me.txtDemandNo.FieldName = Nothing
        Me.txtDemandNo.isCalculatedField = False
        Me.txtDemandNo.IsSourceFromTable = False
        Me.txtDemandNo.IsSourceFromValueList = False
        Me.txtDemandNo.IsUnique = False
        Me.txtDemandNo.Location = New System.Drawing.Point(576, 50)
        Me.txtDemandNo.MendatroryField = True
        Me.txtDemandNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDemandNo.MyLinkLable1 = Me.lblDemandNo
        Me.txtDemandNo.MyLinkLable2 = Nothing
        Me.txtDemandNo.MyReadOnly = False
        Me.txtDemandNo.MyShowMasterFormButton = False
        Me.txtDemandNo.Name = "txtDemandNo"
        Me.txtDemandNo.ReferenceFieldDesc = Nothing
        Me.txtDemandNo.ReferenceFieldName = Nothing
        Me.txtDemandNo.ReferenceTableName = Nothing
        Me.txtDemandNo.Size = New System.Drawing.Size(132, 20)
        Me.txtDemandNo.TabIndex = 1459
        Me.txtDemandNo.Value = ""
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.MyLabel8)
        Me.Panel2.Controls.Add(Me.lblOpKM)
        Me.Panel2.Controls.Add(Me.txtOpKM)
        Me.Panel2.Controls.Add(Me.txtClKM)
        Me.Panel2.Controls.Add(Me.lblClKM)
        Me.Panel2.Location = New System.Drawing.Point(899, 27)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(278, 24)
        Me.Panel2.TabIndex = 81
        Me.Panel2.Visible = False
        '
        'MyLabel8
        '
        Me.MyLabel8.FieldName = Nothing
        Me.MyLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(475, 3)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(30, 16)
        Me.MyLabel8.TabIndex = 81
        Me.MyLabel8.Text = "Date"
        '
        'lblOpKM
        '
        Me.lblOpKM.FieldName = Nothing
        Me.lblOpKM.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpKM.Location = New System.Drawing.Point(3, 4)
        Me.lblOpKM.Name = "lblOpKM"
        Me.lblOpKM.Size = New System.Drawing.Size(69, 16)
        Me.lblOpKM.TabIndex = 77
        Me.lblOpKM.Text = "Opening KM"
        '
        'txtOpKM
        '
        Me.txtOpKM.BackColor = System.Drawing.Color.White
        Me.txtOpKM.CalculationExpression = Nothing
        Me.txtOpKM.DecimalPlaces = 0
        Me.txtOpKM.FieldCode = Nothing
        Me.txtOpKM.FieldDesc = Nothing
        Me.txtOpKM.FieldMaxLength = 5
        Me.txtOpKM.FieldName = Nothing
        Me.txtOpKM.isCalculatedField = False
        Me.txtOpKM.IsSourceFromTable = False
        Me.txtOpKM.IsSourceFromValueList = False
        Me.txtOpKM.IsUnique = False
        Me.txtOpKM.Location = New System.Drawing.Point(78, 2)
        Me.txtOpKM.MendatroryField = False
        Me.txtOpKM.MyLinkLable1 = Me.lblOpKM
        Me.txtOpKM.MyLinkLable2 = Nothing
        Me.txtOpKM.Name = "txtOpKM"
        Me.txtOpKM.ReferenceFieldDesc = Nothing
        Me.txtOpKM.ReferenceFieldName = Nothing
        Me.txtOpKM.ReferenceTableName = Nothing
        Me.txtOpKM.Size = New System.Drawing.Size(55, 20)
        Me.txtOpKM.TabIndex = 76
        Me.txtOpKM.Text = "0"
        Me.txtOpKM.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOpKM.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'txtClKM
        '
        Me.txtClKM.BackColor = System.Drawing.Color.White
        Me.txtClKM.CalculationExpression = Nothing
        Me.txtClKM.DecimalPlaces = 0
        Me.txtClKM.FieldCode = Nothing
        Me.txtClKM.FieldDesc = Nothing
        Me.txtClKM.FieldMaxLength = 5
        Me.txtClKM.FieldName = Nothing
        Me.txtClKM.isCalculatedField = False
        Me.txtClKM.IsSourceFromTable = False
        Me.txtClKM.IsSourceFromValueList = False
        Me.txtClKM.IsUnique = False
        Me.txtClKM.Location = New System.Drawing.Point(206, 3)
        Me.txtClKM.MendatroryField = False
        Me.txtClKM.MyLinkLable1 = Me.lblClKM
        Me.txtClKM.MyLinkLable2 = Nothing
        Me.txtClKM.Name = "txtClKM"
        Me.txtClKM.ReferenceFieldDesc = Nothing
        Me.txtClKM.ReferenceFieldName = Nothing
        Me.txtClKM.ReferenceTableName = Nothing
        Me.txtClKM.Size = New System.Drawing.Size(61, 20)
        Me.txtClKM.TabIndex = 78
        Me.txtClKM.Text = "0"
        Me.txtClKM.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtClKM.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'lblClKM
        '
        Me.lblClKM.FieldName = Nothing
        Me.lblClKM.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClKM.Location = New System.Drawing.Point(136, 4)
        Me.lblClKM.Name = "lblClKM"
        Me.lblClKM.Size = New System.Drawing.Size(64, 16)
        Me.lblClKM.TabIndex = 79
        Me.lblClKM.Text = "Closing KM"
        '
        'chkIndividualCustomer
        '
        Me.chkIndividualCustomer.Location = New System.Drawing.Point(507, 28)
        Me.chkIndividualCustomer.Name = "chkIndividualCustomer"
        Me.chkIndividualCustomer.Size = New System.Drawing.Size(120, 18)
        Me.chkIndividualCustomer.TabIndex = 1458
        Me.chkIndividualCustomer.Text = "Individual Customer"
        '
        'chkGhee
        '
        Me.chkGhee.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGhee.Location = New System.Drawing.Point(717, 111)
        Me.chkGhee.Name = "chkGhee"
        Me.chkGhee.Size = New System.Drawing.Size(74, 16)
        Me.chkGhee.TabIndex = 1047
        Me.chkGhee.Text = "Only Ghee"
        Me.chkGhee.Visible = False
        '
        'rgbItemType
        '
        Me.rgbItemType.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.rgbItemType.Controls.Add(Me.rbtn_IceCream)
        Me.rgbItemType.Controls.Add(Me.rbtn_product)
        Me.rgbItemType.Controls.Add(Me.rbtn_Milk)
        Me.rgbItemType.HeaderText = ""
        Me.rgbItemType.Location = New System.Drawing.Point(883, 111)
        Me.rgbItemType.Name = "rgbItemType"
        Me.rgbItemType.Size = New System.Drawing.Size(97, 60)
        Me.rgbItemType.TabIndex = 1046
        '
        'rbtn_IceCream
        '
        Me.rbtn_IceCream.Location = New System.Drawing.Point(4, 41)
        Me.rbtn_IceCream.Name = "rbtn_IceCream"
        Me.rbtn_IceCream.Size = New System.Drawing.Size(67, 18)
        Me.rbtn_IceCream.TabIndex = 2
        Me.rbtn_IceCream.TabStop = False
        Me.rbtn_IceCream.Text = "IceCream"
        '
        'rbtn_product
        '
        Me.rbtn_product.Location = New System.Drawing.Point(4, 24)
        Me.rbtn_product.Name = "rbtn_product"
        Me.rbtn_product.Size = New System.Drawing.Size(59, 18)
        Me.rbtn_product.TabIndex = 1
        Me.rbtn_product.TabStop = False
        Me.rbtn_product.Text = "Product"
        '
        'rbtn_Milk
        '
        Me.rbtn_Milk.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtn_Milk.Location = New System.Drawing.Point(4, 4)
        Me.rbtn_Milk.Name = "rbtn_Milk"
        Me.rbtn_Milk.Size = New System.Drawing.Size(41, 18)
        Me.rbtn_Milk.TabIndex = 0
        Me.rbtn_Milk.TabStop = False
        Me.rbtn_Milk.Text = "Milk"
        Me.rbtn_Milk.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'lblClosingDate
        '
        Me.lblClosingDate.AutoSize = False
        Me.lblClosingDate.BorderVisible = True
        Me.lblClosingDate.FieldName = Nothing
        Me.lblClosingDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosingDate.Location = New System.Drawing.Point(985, 89)
        Me.lblClosingDate.Name = "lblClosingDate"
        Me.lblClosingDate.Size = New System.Drawing.Size(107, 18)
        Me.lblClosingDate.TabIndex = 138
        Me.lblClosingDate.TextWrap = False
        '
        'txtTripNo
        '
        Me.txtTripNo.CalculationExpression = Nothing
        Me.txtTripNo.FieldCode = Nothing
        Me.txtTripNo.FieldDesc = Nothing
        Me.txtTripNo.FieldMaxLength = 0
        Me.txtTripNo.FieldName = Nothing
        Me.txtTripNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTripNo.isCalculatedField = False
        Me.txtTripNo.IsSourceFromTable = False
        Me.txtTripNo.IsSourceFromValueList = False
        Me.txtTripNo.IsUnique = False
        Me.txtTripNo.Location = New System.Drawing.Point(790, 8)
        Me.txtTripNo.MaxLength = 200
        Me.txtTripNo.MendatroryField = False
        Me.txtTripNo.MyLinkLable1 = Me.MyLabel2
        Me.txtTripNo.MyLinkLable2 = Nothing
        Me.txtTripNo.Name = "txtTripNo"
        Me.txtTripNo.ReferenceFieldDesc = Nothing
        Me.txtTripNo.ReferenceFieldName = Nothing
        Me.txtTripNo.ReferenceTableName = Nothing
        Me.txtTripNo.Size = New System.Drawing.Size(64, 18)
        Me.txtTripNo.TabIndex = 1044
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(4, 112)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(61, 16)
        Me.MyLabel2.TabIndex = 23
        Me.MyLabel2.Text = "Comments"
        '
        'lblTripNo
        '
        Me.lblTripNo.FieldName = Nothing
        Me.lblTripNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTripNo.Location = New System.Drawing.Point(716, 9)
        Me.lblTripNo.Name = "lblTripNo"
        Me.lblTripNo.Size = New System.Drawing.Size(43, 16)
        Me.lblTripNo.TabIndex = 1045
        Me.lblTripNo.Text = "Trip No"
        '
        'txtSupplyDate
        '
        Me.txtSupplyDate.CalculationExpression = Nothing
        Me.txtSupplyDate.CustomFormat = "dd/MM/yyyy"
        Me.txtSupplyDate.FieldCode = Nothing
        Me.txtSupplyDate.FieldDesc = Nothing
        Me.txtSupplyDate.FieldMaxLength = 0
        Me.txtSupplyDate.FieldName = Nothing
        Me.txtSupplyDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtSupplyDate.isCalculatedField = False
        Me.txtSupplyDate.IsSourceFromTable = False
        Me.txtSupplyDate.IsSourceFromValueList = False
        Me.txtSupplyDate.IsUnique = False
        Me.txtSupplyDate.Location = New System.Drawing.Point(576, 6)
        Me.txtSupplyDate.MendatroryField = False
        Me.txtSupplyDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtSupplyDate.MyLinkLable1 = Me.lblSupplyDate
        Me.txtSupplyDate.MyLinkLable2 = Nothing
        Me.txtSupplyDate.Name = "txtSupplyDate"
        Me.txtSupplyDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtSupplyDate.ReferenceFieldDesc = Nothing
        Me.txtSupplyDate.ReferenceFieldName = Nothing
        Me.txtSupplyDate.ReferenceTableName = Nothing
        Me.txtSupplyDate.Size = New System.Drawing.Size(93, 20)
        Me.txtSupplyDate.TabIndex = 1042
        Me.txtSupplyDate.TabStop = False
        Me.txtSupplyDate.Text = "10/06/2011"
        Me.txtSupplyDate.Value = New Date(2011, 6, 10, 11, 51, 56, 953)
        '
        'lblSupplyDate
        '
        Me.lblSupplyDate.FieldName = Nothing
        Me.lblSupplyDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSupplyDate.Location = New System.Drawing.Point(505, 8)
        Me.lblSupplyDate.Name = "lblSupplyDate"
        Me.lblSupplyDate.Size = New System.Drawing.Size(68, 16)
        Me.lblSupplyDate.TabIndex = 1043
        Me.lblSupplyDate.Text = "Supply Date"
        '
        'btnClKM
        '
        Me.btnClKM.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClKM.Location = New System.Drawing.Point(884, 88)
        Me.btnClKM.Name = "btnClKM"
        Me.btnClKM.Size = New System.Drawing.Size(91, 20)
        Me.btnClKM.TabIndex = 80
        Me.btnClKM.Text = "Set Closing KM"
        '
        'txtDistributorName
        '
        Me.txtDistributorName.CalculationExpression = Nothing
        Me.txtDistributorName.FieldCode = Nothing
        Me.txtDistributorName.FieldDesc = Nothing
        Me.txtDistributorName.FieldMaxLength = 0
        Me.txtDistributorName.FieldName = Nothing
        Me.txtDistributorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDistributorName.isCalculatedField = False
        Me.txtDistributorName.IsSourceFromTable = False
        Me.txtDistributorName.IsSourceFromValueList = False
        Me.txtDistributorName.IsUnique = False
        Me.txtDistributorName.Location = New System.Drawing.Point(161, 153)
        Me.txtDistributorName.MaxLength = 200
        Me.txtDistributorName.MendatroryField = False
        Me.txtDistributorName.MyLinkLable1 = Me.lblDistributorName
        Me.txtDistributorName.MyLinkLable2 = Nothing
        Me.txtDistributorName.Name = "txtDistributorName"
        Me.txtDistributorName.ReadOnly = True
        Me.txtDistributorName.ReferenceFieldDesc = Nothing
        Me.txtDistributorName.ReferenceFieldName = Nothing
        Me.txtDistributorName.ReferenceTableName = Nothing
        Me.txtDistributorName.Size = New System.Drawing.Size(210, 18)
        Me.txtDistributorName.TabIndex = 1040
        '
        'lblDistributorName
        '
        Me.lblDistributorName.FieldName = Nothing
        Me.lblDistributorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistributorName.Location = New System.Drawing.Point(6, 153)
        Me.lblDistributorName.Name = "lblDistributorName"
        Me.lblDistributorName.Size = New System.Drawing.Size(149, 16)
        Me.lblDistributorName.TabIndex = 1041
        Me.lblDistributorName.Text = "Distributor/Transpoter Name"
        '
        'txtDriverMobNo
        '
        Me.txtDriverMobNo.CalculationExpression = Nothing
        Me.txtDriverMobNo.FieldCode = Nothing
        Me.txtDriverMobNo.FieldDesc = Nothing
        Me.txtDriverMobNo.FieldMaxLength = 0
        Me.txtDriverMobNo.FieldName = Nothing
        Me.txtDriverMobNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDriverMobNo.isCalculatedField = False
        Me.txtDriverMobNo.IsSourceFromTable = False
        Me.txtDriverMobNo.IsSourceFromValueList = False
        Me.txtDriverMobNo.IsUnique = False
        Me.txtDriverMobNo.Location = New System.Drawing.Point(422, 176)
        Me.txtDriverMobNo.MaxLength = 15
        Me.txtDriverMobNo.MendatroryField = False
        Me.txtDriverMobNo.MyLinkLable1 = Me.MyLabel14
        Me.txtDriverMobNo.MyLinkLable2 = Nothing
        Me.txtDriverMobNo.Name = "txtDriverMobNo"
        Me.txtDriverMobNo.ReferenceFieldDesc = Nothing
        Me.txtDriverMobNo.ReferenceFieldName = Nothing
        Me.txtDriverMobNo.ReferenceTableName = Nothing
        Me.txtDriverMobNo.Size = New System.Drawing.Size(253, 18)
        Me.txtDriverMobNo.TabIndex = 1038
        '
        'MyLabel14
        '
        Me.MyLabel14.FieldName = Nothing
        Me.MyLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel14.Location = New System.Drawing.Point(342, 177)
        Me.MyLabel14.Name = "MyLabel14"
        Me.MyLabel14.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel14.TabIndex = 1039
        Me.MyLabel14.Text = "Driver Mob No."
        '
        'txtDriverName
        '
        Me.txtDriverName.CalculationExpression = Nothing
        Me.txtDriverName.FieldCode = Nothing
        Me.txtDriverName.FieldDesc = Nothing
        Me.txtDriverName.FieldMaxLength = 0
        Me.txtDriverName.FieldName = Nothing
        Me.txtDriverName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDriverName.isCalculatedField = False
        Me.txtDriverName.IsSourceFromTable = False
        Me.txtDriverName.IsSourceFromValueList = False
        Me.txtDriverName.IsUnique = False
        Me.txtDriverName.Location = New System.Drawing.Point(83, 176)
        Me.txtDriverName.MaxLength = 200
        Me.txtDriverName.MendatroryField = False
        Me.txtDriverName.MyLinkLable1 = Me.MyLabel13
        Me.txtDriverName.MyLinkLable2 = Nothing
        Me.txtDriverName.Name = "txtDriverName"
        Me.txtDriverName.ReferenceFieldDesc = Nothing
        Me.txtDriverName.ReferenceFieldName = Nothing
        Me.txtDriverName.ReferenceTableName = Nothing
        Me.txtDriverName.Size = New System.Drawing.Size(253, 18)
        Me.txtDriverName.TabIndex = 1036
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(4, 177)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(70, 16)
        Me.MyLabel13.TabIndex = 1037
        Me.MyLabel13.Text = "Driver Name"
        '
        'txtGatepassDate
        '
        Me.txtGatepassDate.CalculationExpression = Nothing
        Me.txtGatepassDate.CustomFormat = "dd/MM/yyyy  hh:mm tt"
        Me.txtGatepassDate.FieldCode = Nothing
        Me.txtGatepassDate.FieldDesc = Nothing
        Me.txtGatepassDate.FieldMaxLength = 0
        Me.txtGatepassDate.FieldName = Nothing
        Me.txtGatepassDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtGatepassDate.isCalculatedField = False
        Me.txtGatepassDate.IsSourceFromTable = False
        Me.txtGatepassDate.IsSourceFromValueList = False
        Me.txtGatepassDate.IsUnique = False
        Me.txtGatepassDate.Location = New System.Drawing.Point(853, 174)
        Me.txtGatepassDate.MendatroryField = False
        Me.txtGatepassDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtGatepassDate.MyLinkLable1 = Me.MyLabel12
        Me.txtGatepassDate.MyLinkLable2 = Nothing
        Me.txtGatepassDate.Name = "txtGatepassDate"
        Me.txtGatepassDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtGatepassDate.ReferenceFieldDesc = Nothing
        Me.txtGatepassDate.ReferenceFieldName = Nothing
        Me.txtGatepassDate.ReferenceTableName = Nothing
        Me.txtGatepassDate.Size = New System.Drawing.Size(131, 20)
        Me.txtGatepassDate.TabIndex = 1034
        Me.txtGatepassDate.TabStop = False
        Me.txtGatepassDate.Text = "10/06/2011  11:51 AM"
        Me.txtGatepassDate.Value = New Date(2011, 6, 10, 11, 51, 56, 953)
        Me.txtGatepassDate.Visible = False
        '
        'MyLabel12
        '
        Me.MyLabel12.FieldName = Nothing
        Me.MyLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel12.Location = New System.Drawing.Point(772, 176)
        Me.MyLabel12.Name = "MyLabel12"
        Me.MyLabel12.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel12.TabIndex = 1035
        Me.MyLabel12.Text = "Gatepass Date"
        Me.MyLabel12.Visible = False
        '
        'txtLoadingSlip
        '
        Me.txtLoadingSlip.CalculationExpression = Nothing
        Me.txtLoadingSlip.FieldCode = Nothing
        Me.txtLoadingSlip.FieldDesc = Nothing
        Me.txtLoadingSlip.FieldMaxLength = 0
        Me.txtLoadingSlip.FieldName = Nothing
        Me.txtLoadingSlip.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLoadingSlip.isCalculatedField = False
        Me.txtLoadingSlip.IsSourceFromTable = False
        Me.txtLoadingSlip.IsSourceFromValueList = False
        Me.txtLoadingSlip.IsUnique = False
        Me.txtLoadingSlip.Location = New System.Drawing.Point(790, 28)
        Me.txtLoadingSlip.MaxLength = 200
        Me.txtLoadingSlip.MendatroryField = False
        Me.txtLoadingSlip.MyLinkLable1 = Me.MyLabel2
        Me.txtLoadingSlip.MyLinkLable2 = Nothing
        Me.txtLoadingSlip.Name = "txtLoadingSlip"
        Me.txtLoadingSlip.ReferenceFieldDesc = Nothing
        Me.txtLoadingSlip.ReferenceFieldName = Nothing
        Me.txtLoadingSlip.ReferenceTableName = Nothing
        Me.txtLoadingSlip.Size = New System.Drawing.Size(108, 18)
        Me.txtLoadingSlip.TabIndex = 8
        '
        'MyLabel11
        '
        Me.MyLabel11.FieldName = Nothing
        Me.MyLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel11.Location = New System.Drawing.Point(716, 29)
        Me.MyLabel11.Name = "MyLabel11"
        Me.MyLabel11.Size = New System.Drawing.Size(69, 16)
        Me.MyLabel11.TabIndex = 25
        Me.MyLabel11.Text = "Loading Slip"
        '
        'RadGroupBox3
        '
        Me.RadGroupBox3.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox3.Controls.Add(Me.rbtnEvening)
        Me.RadGroupBox3.Controls.Add(Me.rbtnMorning)
        Me.RadGroupBox3.HeaderText = ""
        Me.RadGroupBox3.Location = New System.Drawing.Point(985, 112)
        Me.RadGroupBox3.Name = "RadGroupBox3"
        Me.RadGroupBox3.Size = New System.Drawing.Size(97, 60)
        Me.RadGroupBox3.TabIndex = 1033
        '
        'rbtnEvening
        '
        Me.rbtnEvening.Location = New System.Drawing.Point(4, 24)
        Me.rbtnEvening.Name = "rbtnEvening"
        Me.rbtnEvening.Size = New System.Drawing.Size(59, 18)
        Me.rbtnEvening.TabIndex = 1
        Me.rbtnEvening.TabStop = False
        Me.rbtnEvening.Text = "Evening"
        '
        'rbtnMorning
        '
        Me.rbtnMorning.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtnMorning.Location = New System.Drawing.Point(4, 4)
        Me.rbtnMorning.Name = "rbtnMorning"
        Me.rbtnMorning.Size = New System.Drawing.Size(63, 18)
        Me.rbtnMorning.TabIndex = 0
        Me.rbtnMorning.TabStop = False
        Me.rbtnMorning.Text = "Morning"
        Me.rbtnMorning.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'btnMultiGPReverse
        '
        Me.btnMultiGPReverse.Location = New System.Drawing.Point(716, 152)
        Me.btnMultiGPReverse.Name = "btnMultiGPReverse"
        Me.btnMultiGPReverse.Size = New System.Drawing.Size(88, 19)
        Me.btnMultiGPReverse.TabIndex = 1032
        Me.btnMultiGPReverse.Text = ">>"
        '
        'TxtMultiDairyGPassReverse
        '
        Me.TxtMultiDairyGPassReverse.arrDispalyMember = Nothing
        Me.TxtMultiDairyGPassReverse.arrValueMember = Nothing
        Me.TxtMultiDairyGPassReverse.Location = New System.Drawing.Point(466, 152)
        Me.TxtMultiDairyGPassReverse.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMultiDairyGPassReverse.MyLinkLable1 = Nothing
        Me.TxtMultiDairyGPassReverse.MyLinkLable2 = Nothing
        Me.TxtMultiDairyGPassReverse.MyNullText = "All"
        Me.TxtMultiDairyGPassReverse.Name = "TxtMultiDairyGPassReverse"
        Me.TxtMultiDairyGPassReverse.Size = New System.Drawing.Size(243, 19)
        Me.TxtMultiDairyGPassReverse.TabIndex = 1031
        '
        'MyLabel34
        '
        Me.MyLabel34.FieldName = Nothing
        Me.MyLabel34.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel34.Location = New System.Drawing.Point(377, 151)
        Me.MyLabel34.Name = "MyLabel34"
        Me.MyLabel34.Size = New System.Drawing.Size(89, 18)
        Me.MyLabel34.TabIndex = 1030
        Me.MyLabel34.Text = "Multiple Reverse"
        '
        'lblTollAmount
        '
        Me.lblTollAmount.FieldName = Nothing
        Me.lblTollAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTollAmount.Location = New System.Drawing.Point(715, 89)
        Me.lblTollAmount.Name = "lblTollAmount"
        Me.lblTollAmount.Size = New System.Drawing.Size(67, 16)
        Me.lblTollAmount.TabIndex = 89
        Me.lblTollAmount.Text = "Toll Amount"
        '
        'txtTollAmount
        '
        Me.txtTollAmount.BackColor = System.Drawing.Color.White
        Me.txtTollAmount.CalculationExpression = Nothing
        Me.txtTollAmount.DecimalPlaces = 0
        Me.txtTollAmount.FieldCode = Nothing
        Me.txtTollAmount.FieldDesc = Nothing
        Me.txtTollAmount.FieldMaxLength = 5
        Me.txtTollAmount.FieldName = Nothing
        Me.txtTollAmount.isCalculatedField = False
        Me.txtTollAmount.IsSourceFromTable = False
        Me.txtTollAmount.IsSourceFromValueList = False
        Me.txtTollAmount.IsUnique = False
        Me.txtTollAmount.Location = New System.Drawing.Point(790, 87)
        Me.txtTollAmount.MendatroryField = False
        Me.txtTollAmount.MyLinkLable1 = Me.lblTollAmount
        Me.txtTollAmount.MyLinkLable2 = Nothing
        Me.txtTollAmount.Name = "txtTollAmount"
        Me.txtTollAmount.ReferenceFieldDesc = Nothing
        Me.txtTollAmount.ReferenceFieldName = Nothing
        Me.txtTollAmount.ReferenceTableName = Nothing
        Me.txtTollAmount.Size = New System.Drawing.Size(88, 20)
        Me.txtTollAmount.TabIndex = 88
        Me.txtTollAmount.Text = "0"
        Me.txtTollAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTollAmount.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'MyLabel9
        '
        Me.MyLabel9.FieldName = Nothing
        Me.MyLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(714, 69)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(66, 16)
        Me.MyLabel9.TabIndex = 87
        Me.MyLabel9.Text = "Transfer No"
        Me.MyLabel9.Visible = False
        '
        'FndTransferNo
        '
        Me.FndTransferNo.CalculationExpression = Nothing
        Me.FndTransferNo.FieldCode = Nothing
        Me.FndTransferNo.FieldDesc = Nothing
        Me.FndTransferNo.FieldMaxLength = 0
        Me.FndTransferNo.FieldName = Nothing
        Me.FndTransferNo.isCalculatedField = False
        Me.FndTransferNo.IsSourceFromTable = False
        Me.FndTransferNo.IsSourceFromValueList = False
        Me.FndTransferNo.IsUnique = False
        Me.FndTransferNo.Location = New System.Drawing.Point(790, 67)
        Me.FndTransferNo.MendatroryField = False
        Me.FndTransferNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FndTransferNo.MyLinkLable1 = Me.MyLabel9
        Me.FndTransferNo.MyLinkLable2 = Nothing
        Me.FndTransferNo.MyReadOnly = False
        Me.FndTransferNo.MyShowMasterFormButton = False
        Me.FndTransferNo.Name = "FndTransferNo"
        Me.FndTransferNo.ReferenceFieldDesc = Nothing
        Me.FndTransferNo.ReferenceFieldName = Nothing
        Me.FndTransferNo.ReferenceTableName = Nothing
        Me.FndTransferNo.Size = New System.Drawing.Size(233, 20)
        Me.FndTransferNo.TabIndex = 86
        Me.FndTransferNo.Value = ""
        Me.FndTransferNo.Visible = False
        '
        'chkAgainstTransfer
        '
        Me.chkAgainstTransfer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAgainstTransfer.Location = New System.Drawing.Point(718, 49)
        Me.chkAgainstTransfer.Name = "chkAgainstTransfer"
        Me.chkAgainstTransfer.Size = New System.Drawing.Size(103, 16)
        Me.chkAgainstTransfer.TabIndex = 85
        Me.chkAgainstTransfer.Text = "Against Transfer"
        Me.chkAgainstTransfer.Visible = False
        '
        'txtSalesman
        '
        Me.txtSalesman.CalculationExpression = Nothing
        Me.txtSalesman.FieldCode = Nothing
        Me.txtSalesman.FieldDesc = Nothing
        Me.txtSalesman.FieldMaxLength = 0
        Me.txtSalesman.FieldName = Nothing
        Me.txtSalesman.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSalesman.isCalculatedField = False
        Me.txtSalesman.IsSourceFromTable = False
        Me.txtSalesman.IsSourceFromValueList = False
        Me.txtSalesman.IsUnique = False
        Me.txtSalesman.Location = New System.Drawing.Point(576, 91)
        Me.txtSalesman.MaxLength = 200
        Me.txtSalesman.MendatroryField = False
        Me.txtSalesman.MyLinkLable1 = Me.MyLabel1
        Me.txtSalesman.MyLinkLable2 = Nothing
        Me.txtSalesman.Name = "txtSalesman"
        Me.txtSalesman.ReferenceFieldDesc = Nothing
        Me.txtSalesman.ReferenceFieldName = Nothing
        Me.txtSalesman.ReferenceTableName = Nothing
        Me.txtSalesman.Size = New System.Drawing.Size(137, 18)
        Me.txtSalesman.TabIndex = 84
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(507, 73)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(65, 16)
        Me.MyLabel1.TabIndex = 19
        Me.MyLabel1.Text = "Transporter"
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(507, 92)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(57, 16)
        Me.MyLabel7.TabIndex = 83
        Me.MyLabel7.Text = "Salesman"
        '
        'UsLock1
        '
        Me.UsLock1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.UsLock1.Location = New System.Drawing.Point(860, 6)
        Me.UsLock1.MyFont = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsLock1.Name = "UsLock1"
        Me.UsLock1.Size = New System.Drawing.Size(91, 20)
        Me.UsLock1.Status = common.ERPTransactionStatus.Pending
        Me.UsLock1.TabIndex = 82
        '
        'MyLabel6
        '
        Me.MyLabel6.FieldName = Nothing
        Me.MyLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(507, 132)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(62, 16)
        Me.MyLabel6.TabIndex = 21
        Me.MyLabel6.Text = "Total Crate"
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(4, 131)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(54, 16)
        Me.MyLabel5.TabIndex = 22
        Me.MyLabel5.Text = "Shipment"
        '
        'txtmultiBooking
        '
        Me.txtmultiBooking.arrDispalyMember = Nothing
        Me.txtmultiBooking.arrValueMember = Nothing
        Me.txtmultiBooking.Location = New System.Drawing.Point(83, 130)
        Me.txtmultiBooking.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmultiBooking.MyLinkLable1 = Nothing
        Me.txtmultiBooking.MyLinkLable2 = Nothing
        Me.txtmultiBooking.MyNullText = ""
        Me.txtmultiBooking.Name = "txtmultiBooking"
        Me.txtmultiBooking.Size = New System.Drawing.Size(253, 19)
        Me.txtmultiBooking.TabIndex = 8
        '
        'lblTotalCan
        '
        Me.lblTotalCan.FieldName = Nothing
        Me.lblTotalCan.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCan.Location = New System.Drawing.Point(340, 132)
        Me.lblTotalCan.Name = "lblTotalCan"
        Me.lblTotalCan.Size = New System.Drawing.Size(55, 16)
        Me.lblTotalCan.TabIndex = 20
        Me.lblTotalCan.Text = "Total Can"
        '
        'txtCrateQty
        '
        Me.txtCrateQty.CalculationExpression = Nothing
        Me.txtCrateQty.FieldCode = Nothing
        Me.txtCrateQty.FieldDesc = Nothing
        Me.txtCrateQty.FieldMaxLength = 0
        Me.txtCrateQty.FieldName = Nothing
        Me.txtCrateQty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCrateQty.isCalculatedField = False
        Me.txtCrateQty.IsSourceFromTable = False
        Me.txtCrateQty.IsSourceFromValueList = False
        Me.txtCrateQty.IsUnique = False
        Me.txtCrateQty.Location = New System.Drawing.Point(576, 131)
        Me.txtCrateQty.MaxLength = 200
        Me.txtCrateQty.MendatroryField = False
        Me.txtCrateQty.MyLinkLable1 = Me.MyLabel1
        Me.txtCrateQty.MyLinkLable2 = Nothing
        Me.txtCrateQty.Name = "txtCrateQty"
        Me.txtCrateQty.ReadOnly = True
        Me.txtCrateQty.ReferenceFieldDesc = Nothing
        Me.txtCrateQty.ReferenceFieldName = Nothing
        Me.txtCrateQty.ReferenceTableName = Nothing
        Me.txtCrateQty.Size = New System.Drawing.Size(137, 18)
        Me.txtCrateQty.TabIndex = 10
        Me.txtCrateQty.WordWrap = False
        '
        'txtRouteName
        '
        Me.txtRouteName.CalculationExpression = Nothing
        Me.txtRouteName.FieldCode = Nothing
        Me.txtRouteName.FieldDesc = Nothing
        Me.txtRouteName.FieldMaxLength = 0
        Me.txtRouteName.FieldName = Nothing
        Me.txtRouteName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRouteName.isCalculatedField = False
        Me.txtRouteName.IsSourceFromTable = False
        Me.txtRouteName.IsSourceFromValueList = False
        Me.txtRouteName.IsUnique = False
        Me.txtRouteName.Location = New System.Drawing.Point(337, 72)
        Me.txtRouteName.MaxLength = 200
        Me.txtRouteName.MendatroryField = False
        Me.txtRouteName.MyLinkLable1 = Me.lblSalesman
        Me.txtRouteName.MyLinkLable2 = Nothing
        Me.txtRouteName.Name = "txtRouteName"
        Me.txtRouteName.ReferenceFieldDesc = Nothing
        Me.txtRouteName.ReferenceFieldName = Nothing
        Me.txtRouteName.ReferenceTableName = Nothing
        Me.txtRouteName.Size = New System.Drawing.Size(166, 18)
        Me.txtRouteName.TabIndex = 18
        '
        'lblSalesman
        '
        Me.lblSalesman.FieldName = Nothing
        Me.lblSalesman.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSalesman.Location = New System.Drawing.Point(4, 52)
        Me.lblSalesman.Name = "lblSalesman"
        Me.lblSalesman.Size = New System.Drawing.Size(61, 16)
        Me.lblSalesman.TabIndex = 26
        Me.lblSalesman.Text = "Vehicle No"
        '
        'txtCanQty
        '
        Me.txtCanQty.CalculationExpression = Nothing
        Me.txtCanQty.FieldCode = Nothing
        Me.txtCanQty.FieldDesc = Nothing
        Me.txtCanQty.FieldMaxLength = 0
        Me.txtCanQty.FieldName = Nothing
        Me.txtCanQty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCanQty.isCalculatedField = False
        Me.txtCanQty.IsSourceFromTable = False
        Me.txtCanQty.IsSourceFromValueList = False
        Me.txtCanQty.IsUnique = False
        Me.txtCanQty.Location = New System.Drawing.Point(415, 131)
        Me.txtCanQty.MaxLength = 200
        Me.txtCanQty.MendatroryField = False
        Me.txtCanQty.MyLinkLable1 = Me.MyLabel1
        Me.txtCanQty.MyLinkLable2 = Nothing
        Me.txtCanQty.Name = "txtCanQty"
        Me.txtCanQty.ReadOnly = True
        Me.txtCanQty.ReferenceFieldDesc = Nothing
        Me.txtCanQty.ReferenceFieldName = Nothing
        Me.txtCanQty.ReferenceTableName = Nothing
        Me.txtCanQty.Size = New System.Drawing.Size(88, 18)
        Me.txtCanQty.TabIndex = 9
        Me.txtCanQty.WordWrap = False
        '
        'lblRoute
        '
        Me.lblRoute.FieldName = Nothing
        Me.lblRoute.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoute.Location = New System.Drawing.Point(4, 73)
        Me.lblRoute.Name = "lblRoute"
        Me.lblRoute.Size = New System.Drawing.Size(54, 16)
        Me.lblRoute.TabIndex = 25
        Me.lblRoute.Text = "Route No"
        '
        'fndRouteNo
        '
        Me.fndRouteNo.CalculationExpression = Nothing
        Me.fndRouteNo.FieldCode = Nothing
        Me.fndRouteNo.FieldDesc = Nothing
        Me.fndRouteNo.FieldMaxLength = 0
        Me.fndRouteNo.FieldName = Nothing
        Me.fndRouteNo.isCalculatedField = False
        Me.fndRouteNo.IsSourceFromTable = False
        Me.fndRouteNo.IsSourceFromValueList = False
        Me.fndRouteNo.IsUnique = False
        Me.fndRouteNo.Location = New System.Drawing.Point(83, 71)
        Me.fndRouteNo.MendatroryField = False
        Me.fndRouteNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fndRouteNo.MyLinkLable1 = Me.lblRoute
        Me.fndRouteNo.MyLinkLable2 = Nothing
        Me.fndRouteNo.MyReadOnly = False
        Me.fndRouteNo.MyShowMasterFormButton = False
        Me.fndRouteNo.Name = "fndRouteNo"
        Me.fndRouteNo.ReferenceFieldDesc = Nothing
        Me.fndRouteNo.ReferenceFieldName = Nothing
        Me.fndRouteNo.ReferenceTableName = Nothing
        Me.fndRouteNo.Size = New System.Drawing.Size(253, 20)
        Me.fndRouteNo.TabIndex = 4
        Me.fndRouteNo.Value = ""
        '
        'btnGo
        '
        Me.btnGo.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnGo.Location = New System.Drawing.Point(715, 131)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(91, 18)
        Me.btnGo.TabIndex = 11
        Me.btnGo.Text = ">>"
        '
        'txtLocDesc
        '
        Me.txtLocDesc.CalculationExpression = Nothing
        Me.txtLocDesc.FieldCode = Nothing
        Me.txtLocDesc.FieldDesc = Nothing
        Me.txtLocDesc.FieldMaxLength = 0
        Me.txtLocDesc.FieldName = Nothing
        Me.txtLocDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocDesc.isCalculatedField = False
        Me.txtLocDesc.IsSourceFromTable = False
        Me.txtLocDesc.IsSourceFromValueList = False
        Me.txtLocDesc.IsUnique = False
        Me.txtLocDesc.Location = New System.Drawing.Point(337, 30)
        Me.txtLocDesc.MaxLength = 200
        Me.txtLocDesc.MendatroryField = False
        Me.txtLocDesc.MyLinkLable1 = Me.MyLabel4
        Me.txtLocDesc.MyLinkLable2 = Nothing
        Me.txtLocDesc.Name = "txtLocDesc"
        Me.txtLocDesc.ReferenceFieldDesc = Nothing
        Me.txtLocDesc.ReferenceFieldName = Nothing
        Me.txtLocDesc.ReferenceTableName = Nothing
        Me.txtLocDesc.Size = New System.Drawing.Size(161, 18)
        Me.txtLocDesc.TabIndex = 16
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(4, 31)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(49, 16)
        Me.MyLabel4.TabIndex = 27
        Me.MyLabel4.Text = "Location"
        '
        'txtLocCode
        '
        Me.txtLocCode.CalculationExpression = Nothing
        Me.txtLocCode.FieldCode = Nothing
        Me.txtLocCode.FieldDesc = Nothing
        Me.txtLocCode.FieldMaxLength = 0
        Me.txtLocCode.FieldName = Nothing
        Me.txtLocCode.isCalculatedField = False
        Me.txtLocCode.IsSourceFromTable = False
        Me.txtLocCode.IsSourceFromValueList = False
        Me.txtLocCode.IsUnique = False
        Me.txtLocCode.Location = New System.Drawing.Point(83, 29)
        Me.txtLocCode.MendatroryField = True
        Me.txtLocCode.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocCode.MyLinkLable1 = Me.MyLabel4
        Me.txtLocCode.MyLinkLable2 = Nothing
        Me.txtLocCode.MyReadOnly = False
        Me.txtLocCode.MyShowMasterFormButton = False
        Me.txtLocCode.Name = "txtLocCode"
        Me.txtLocCode.ReferenceFieldDesc = Nothing
        Me.txtLocCode.ReferenceFieldName = Nothing
        Me.txtLocCode.ReferenceTableName = Nothing
        Me.txtLocCode.Size = New System.Drawing.Size(253, 20)
        Me.txtLocCode.TabIndex = 2
        Me.txtLocCode.Value = ""
        '
        'txtComments
        '
        Me.txtComments.CalculationExpression = Nothing
        Me.txtComments.FieldCode = Nothing
        Me.txtComments.FieldDesc = Nothing
        Me.txtComments.FieldMaxLength = 0
        Me.txtComments.FieldName = Nothing
        Me.txtComments.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComments.isCalculatedField = False
        Me.txtComments.IsSourceFromTable = False
        Me.txtComments.IsSourceFromValueList = False
        Me.txtComments.IsUnique = False
        Me.txtComments.Location = New System.Drawing.Point(83, 111)
        Me.txtComments.MaxLength = 200
        Me.txtComments.MendatroryField = False
        Me.txtComments.MyLinkLable1 = Me.MyLabel2
        Me.txtComments.MyLinkLable2 = Nothing
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ReferenceFieldDesc = Nothing
        Me.txtComments.ReferenceFieldName = Nothing
        Me.txtComments.ReferenceTableName = Nothing
        Me.txtComments.Size = New System.Drawing.Size(253, 18)
        Me.txtComments.TabIndex = 7
        '
        'txtRemarks
        '
        Me.txtRemarks.CalculationExpression = Nothing
        Me.txtRemarks.FieldCode = Nothing
        Me.txtRemarks.FieldDesc = Nothing
        Me.txtRemarks.FieldMaxLength = 0
        Me.txtRemarks.FieldName = Nothing
        Me.txtRemarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemarks.isCalculatedField = False
        Me.txtRemarks.IsSourceFromTable = False
        Me.txtRemarks.IsSourceFromValueList = False
        Me.txtRemarks.IsUnique = False
        Me.txtRemarks.Location = New System.Drawing.Point(83, 92)
        Me.txtRemarks.MaxLength = 200
        Me.txtRemarks.MendatroryField = False
        Me.txtRemarks.MyLinkLable1 = Me.MyLabel3
        Me.txtRemarks.MyLinkLable2 = Nothing
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ReferenceFieldDesc = Nothing
        Me.txtRemarks.ReferenceFieldName = Nothing
        Me.txtRemarks.ReferenceTableName = Nothing
        Me.txtRemarks.Size = New System.Drawing.Size(420, 18)
        Me.txtRemarks.TabIndex = 6
        '
        'MyLabel3
        '
        Me.MyLabel3.FieldName = Nothing
        Me.MyLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel3.Location = New System.Drawing.Point(4, 93)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(51, 16)
        Me.MyLabel3.TabIndex = 24
        Me.MyLabel3.Text = "Remarks"
        '
        'txtTransporter
        '
        Me.txtTransporter.CalculationExpression = Nothing
        Me.txtTransporter.FieldCode = Nothing
        Me.txtTransporter.FieldDesc = Nothing
        Me.txtTransporter.FieldMaxLength = 0
        Me.txtTransporter.FieldName = Nothing
        Me.txtTransporter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTransporter.isCalculatedField = False
        Me.txtTransporter.IsSourceFromTable = False
        Me.txtTransporter.IsSourceFromValueList = False
        Me.txtTransporter.IsUnique = False
        Me.txtTransporter.Location = New System.Drawing.Point(576, 72)
        Me.txtTransporter.MaxLength = 200
        Me.txtTransporter.MendatroryField = False
        Me.txtTransporter.MyLinkLable1 = Me.MyLabel1
        Me.txtTransporter.MyLinkLable2 = Nothing
        Me.txtTransporter.Name = "txtTransporter"
        Me.txtTransporter.ReferenceFieldDesc = Nothing
        Me.txtTransporter.ReferenceFieldName = Nothing
        Me.txtTransporter.ReferenceTableName = Nothing
        Me.txtTransporter.Size = New System.Drawing.Size(137, 18)
        Me.txtTransporter.TabIndex = 5
        '
        'lblVehicleDesc
        '
        Me.lblVehicleDesc.CalculationExpression = Nothing
        Me.lblVehicleDesc.FieldCode = Nothing
        Me.lblVehicleDesc.FieldDesc = Nothing
        Me.lblVehicleDesc.FieldMaxLength = 0
        Me.lblVehicleDesc.FieldName = Nothing
        Me.lblVehicleDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVehicleDesc.isCalculatedField = False
        Me.lblVehicleDesc.IsSourceFromTable = False
        Me.lblVehicleDesc.IsSourceFromValueList = False
        Me.lblVehicleDesc.IsUnique = False
        Me.lblVehicleDesc.Location = New System.Drawing.Point(337, 51)
        Me.lblVehicleDesc.MaxLength = 200
        Me.lblVehicleDesc.MendatroryField = False
        Me.lblVehicleDesc.MyLinkLable1 = Me.lblSalesman
        Me.lblVehicleDesc.MyLinkLable2 = Nothing
        Me.lblVehicleDesc.Name = "lblVehicleDesc"
        Me.lblVehicleDesc.ReferenceFieldDesc = Nothing
        Me.lblVehicleDesc.ReferenceFieldName = Nothing
        Me.lblVehicleDesc.ReferenceTableName = Nothing
        Me.lblVehicleDesc.Size = New System.Drawing.Size(161, 18)
        Me.lblVehicleDesc.TabIndex = 17
        '
        'txtDate
        '
        Me.txtDate.CalculationExpression = Nothing
        Me.txtDate.CustomFormat = "dd/MM/yyyy  hh:mm tt"
        Me.txtDate.FieldCode = Nothing
        Me.txtDate.FieldDesc = Nothing
        Me.txtDate.FieldMaxLength = 0
        Me.txtDate.FieldName = Nothing
        Me.txtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDate.isCalculatedField = False
        Me.txtDate.IsSourceFromTable = False
        Me.txtDate.IsSourceFromValueList = False
        Me.txtDate.IsUnique = False
        Me.txtDate.Location = New System.Drawing.Point(367, 7)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Me.lblpaymentpostdate
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(131, 20)
        Me.txtDate.TabIndex = 0
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "10/06/2011  11:51 AM"
        Me.txtDate.Value = New Date(2011, 6, 10, 11, 51, 56, 953)
        '
        'lblpaymentpostdate
        '
        Me.lblpaymentpostdate.FieldName = Nothing
        Me.lblpaymentpostdate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpaymentpostdate.Location = New System.Drawing.Point(336, 9)
        Me.lblpaymentpostdate.Name = "lblpaymentpostdate"
        Me.lblpaymentpostdate.Size = New System.Drawing.Size(30, 16)
        Me.lblpaymentpostdate.TabIndex = 14
        Me.lblpaymentpostdate.Text = "Date"
        '
        'cmbitemtype
        '
        Me.cmbitemtype.AutoCompleteDisplayMember = Nothing
        Me.cmbitemtype.AutoCompleteValueMember = Nothing
        Me.cmbitemtype.CalculationExpression = Nothing
        Me.cmbitemtype.DropDownAnimationEnabled = True
        Me.cmbitemtype.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbitemtype.FieldCode = Nothing
        Me.cmbitemtype.FieldDesc = Nothing
        Me.cmbitemtype.FieldMaxLength = 0
        Me.cmbitemtype.FieldName = Nothing
        Me.cmbitemtype.isCalculatedField = False
        Me.cmbitemtype.IsSourceFromTable = False
        Me.cmbitemtype.IsSourceFromValueList = False
        Me.cmbitemtype.IsUnique = False
        RadListDataItem1.Text = "Select"
        RadListDataItem2.Text = "Full"
        RadListDataItem3.Text = "Empty"
        Me.cmbitemtype.Items.Add(RadListDataItem1)
        Me.cmbitemtype.Items.Add(RadListDataItem2)
        Me.cmbitemtype.Items.Add(RadListDataItem3)
        Me.cmbitemtype.Location = New System.Drawing.Point(896, 49)
        Me.cmbitemtype.MendatroryField = False
        Me.cmbitemtype.MyLinkLable1 = Me.lblfullempty
        Me.cmbitemtype.MyLinkLable2 = Nothing
        Me.cmbitemtype.Name = "cmbitemtype"
        Me.cmbitemtype.ReferenceFieldDesc = Nothing
        Me.cmbitemtype.ReferenceFieldName = Nothing
        Me.cmbitemtype.ReferenceTableName = Nothing
        Me.cmbitemtype.Size = New System.Drawing.Size(137, 20)
        Me.cmbitemtype.TabIndex = 1
        Me.cmbitemtype.Visible = False
        '
        'lblfullempty
        '
        Me.lblfullempty.FieldName = Nothing
        Me.lblfullempty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfullempty.Location = New System.Drawing.Point(827, 51)
        Me.lblfullempty.Name = "lblfullempty"
        Me.lblfullempty.Size = New System.Drawing.Size(57, 16)
        Me.lblfullempty.TabIndex = 15
        Me.lblfullempty.Text = "Item Type"
        Me.lblfullempty.Visible = False
        '
        'lblpaymentno
        '
        Me.lblpaymentno.FieldName = Nothing
        Me.lblpaymentno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpaymentno.Location = New System.Drawing.Point(4, 9)
        Me.lblpaymentno.Name = "lblpaymentno"
        Me.lblpaymentno.Size = New System.Drawing.Size(77, 16)
        Me.lblpaymentno.TabIndex = 28
        Me.lblpaymentno.Text = "Gate Pass No"
        '
        'btnNew
        '
        Me.btnNew.BackgroundImage = Global.XpertERPDairySale.My.Resources.Resources._new
        Me.btnNew.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNew.Image = Global.XpertERPDairySale.My.Resources.Resources._new
        Me.btnNew.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnNew.Location = New System.Drawing.Point(317, 7)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(19, 21)
        Me.btnNew.TabIndex = 13
        '
        'txtVehicle
        '
        Me.txtVehicle.CalculationExpression = Nothing
        Me.txtVehicle.FieldCode = Nothing
        Me.txtVehicle.FieldDesc = Nothing
        Me.txtVehicle.FieldMaxLength = 0
        Me.txtVehicle.FieldName = Nothing
        Me.txtVehicle.isCalculatedField = False
        Me.txtVehicle.IsSourceFromTable = False
        Me.txtVehicle.IsSourceFromValueList = False
        Me.txtVehicle.IsUnique = False
        Me.txtVehicle.Location = New System.Drawing.Point(83, 50)
        Me.txtVehicle.MendatroryField = True
        Me.txtVehicle.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVehicle.MyLinkLable1 = Me.lblSalesman
        Me.txtVehicle.MyLinkLable2 = Nothing
        Me.txtVehicle.MyReadOnly = False
        Me.txtVehicle.MyShowMasterFormButton = False
        Me.txtVehicle.Name = "txtVehicle"
        Me.txtVehicle.ReferenceFieldDesc = Nothing
        Me.txtVehicle.ReferenceFieldName = Nothing
        Me.txtVehicle.ReferenceTableName = Nothing
        Me.txtVehicle.Size = New System.Drawing.Size(253, 20)
        Me.txtVehicle.TabIndex = 3
        Me.txtVehicle.Value = ""
        '
        'txtCode
        '
        Me.txtCode.FieldName = Nothing
        Me.txtCode.Location = New System.Drawing.Point(83, 7)
        Me.txtCode.MendatroryField = True
        Me.txtCode.MyCharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.txtCode.MyFont = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.txtCode.MyLinkLable1 = Me.lblpaymentno
        Me.txtCode.MyLinkLable2 = Nothing
        Me.txtCode.MyMaxLength = 30
        Me.txtCode.MyReadOnly = False
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(234, 21)
        Me.txtCode.TabIndex = 12
        Me.txtCode.Value = ""
        '
        'Gv1
        '
        Me.Gv1.Controls.Add(Me.MyRadGridView1)
        Me.Gv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Gv1.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.Gv1.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.Gv1.MasterTemplate.ShowHeaderCellButtons = True
        Me.Gv1.MasterTemplate.ViewDefinition = TableViewDefinition2
        Me.Gv1.MyExportAPI = False
        Me.Gv1.MyExportFilePath = ""
        Me.Gv1.MyStopExport = False
        Me.Gv1.Name = "Gv1"
        Me.Gv1.ShowHeaderCellButtons = True
        Me.Gv1.Size = New System.Drawing.Size(1150, 148)
        Me.Gv1.TabIndex = 1
        Me.Gv1.VarID = ""
        '
        'MyRadGridView1
        '
        Me.MyRadGridView1.Location = New System.Drawing.Point(8, 8)
        '
        '
        '
        Me.MyRadGridView1.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.MyRadGridView1.MasterTemplate.ShowHeaderCellButtons = True
        Me.MyRadGridView1.MasterTemplate.ViewDefinition = TableViewDefinition1
        Me.MyRadGridView1.MyExportAPI = False
        Me.MyRadGridView1.MyExportFilePath = ""
        Me.MyRadGridView1.MyStopExport = False
        Me.MyRadGridView1.Name = "MyRadGridView1"
        Me.MyRadGridView1.ShowHeaderCellButtons = True
        Me.MyRadGridView1.Size = New System.Drawing.Size(289, 197)
        Me.MyRadGridView1.TabIndex = 2
        Me.MyRadGridView1.VarID = ""
        Me.MyRadGridView1.Visible = False
        '
        'rpvpCrateType
        '
        Me.rpvpCrateType.Controls.Add(Me.gvCrateType)
        Me.rpvpCrateType.ItemSize = New System.Drawing.SizeF(104.0!, 26.0!)
        Me.rpvpCrateType.Location = New System.Drawing.Point(10, 35)
        Me.rpvpCrateType.Name = "rpvpCrateType"
        Me.rpvpCrateType.Size = New System.Drawing.Size(1150, 354)
        Me.rpvpCrateType.Text = "Crate Type Detail"
        '
        'gvCrateType
        '
        Me.gvCrateType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gvCrateType.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gvCrateType.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gvCrateType.MasterTemplate.ShowHeaderCellButtons = True
        Me.gvCrateType.MasterTemplate.ViewDefinition = TableViewDefinition3
        Me.gvCrateType.MyExportAPI = False
        Me.gvCrateType.MyExportFilePath = ""
        Me.gvCrateType.MyStopExport = False
        Me.gvCrateType.Name = "gvCrateType"
        Me.gvCrateType.ShowHeaderCellButtons = True
        Me.gvCrateType.Size = New System.Drawing.Size(1150, 354)
        Me.gvCrateType.TabIndex = 3
        Me.gvCrateType.VarID = ""
        '
        'btnHistory
        '
        Me.btnHistory.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHistory.Location = New System.Drawing.Point(746, 3)
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.Size = New System.Drawing.Size(68, 24)
        Me.btnHistory.TabIndex = 1036
        Me.btnHistory.Text = "History"
        '
        'btnBoothSlip
        '
        Me.btnBoothSlip.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnBoothSlip.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem3, Me.btnExcel})
        Me.btnBoothSlip.Location = New System.Drawing.Point(646, 3)
        Me.btnBoothSlip.Name = "btnBoothSlip"
        Me.btnBoothSlip.Size = New System.Drawing.Size(80, 24)
        Me.btnBoothSlip.TabIndex = 1035
        Me.btnBoothSlip.Text = "Booth Slip"
        '
        'RadMenuItem3
        '
        Me.RadMenuItem3.Name = "RadMenuItem3"
        Me.RadMenuItem3.Text = "Booth Slip"
        '
        'btnExcel
        '
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Text = "Excel"
        '
        'btnPrint2
        '
        Me.btnPrint2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrint2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint2.Location = New System.Drawing.Point(286, 3)
        Me.btnPrint2.Name = "btnPrint2"
        Me.btnPrint2.Size = New System.Drawing.Size(68, 24)
        Me.btnPrint2.TabIndex = 2
        Me.btnPrint2.Text = "Print2"
        '
        'btnGPCancel
        '
        Me.btnGPCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGPCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGPCancel.Location = New System.Drawing.Point(504, 3)
        Me.btnGPCancel.Name = "btnGPCancel"
        Me.btnGPCancel.Size = New System.Drawing.Size(69, 24)
        Me.btnGPCancel.TabIndex = 1034
        Me.btnGPCancel.Text = "Cancel"
        '
        'MyLabel10
        '
        Me.MyLabel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.ForeColor = System.Drawing.SystemColors.Highlight
        Me.MyLabel10.Location = New System.Drawing.Point(849, 9)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(251, 16)
        Me.MyLabel10.TabIndex = 1033
        Me.MyLabel10.Text = "Press Ctrl+Alt+Shift+F11 for Multiple Reverse"
        '
        'btnReverse
        '
        Me.btnReverse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReverse.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReverse.Location = New System.Drawing.Point(575, 3)
        Me.btnReverse.Name = "btnReverse"
        Me.btnReverse.Size = New System.Drawing.Size(69, 24)
        Me.btnReverse.TabIndex = 7
        Me.btnReverse.Text = "Reverse"
        Me.btnReverse.Visible = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(216, 3)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(68, 24)
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.Visible = False
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(146, 3)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(68, 24)
        Me.btnPrint.TabIndex = 1
        Me.btnPrint.Text = "Print"
        '
        'RadSplitButton1
        '
        Me.RadSplitButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadSplitButton1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem1, Me.RadMenuItem2, Me.RadMenuItem4})
        Me.RadSplitButton1.Location = New System.Drawing.Point(356, 3)
        Me.RadSplitButton1.Name = "RadSplitButton1"
        Me.RadSplitButton1.Size = New System.Drawing.Size(80, 24)
        Me.RadSplitButton1.TabIndex = 3
        Me.RadSplitButton1.Text = "Print"
        Me.RadSplitButton1.Visible = False
        '
        'RadMenuItem1
        '
        Me.RadMenuItem1.Name = "RadMenuItem1"
        Me.RadMenuItem1.Text = "PrePrinted"
        '
        'RadMenuItem2
        '
        Me.RadMenuItem2.Name = "RadMenuItem2"
        Me.RadMenuItem2.Text = "Print"
        '
        'RadMenuItem4
        '
        Me.RadMenuItem4.Name = "RadMenuItem4"
        Me.RadMenuItem4.Text = "Send Email"
        '
        'btnPost
        '
        Me.btnPost.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPost.Location = New System.Drawing.Point(76, 3)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(68, 24)
        Me.btnPost.TabIndex = 4
        Me.btnPost.Text = "Post"
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(438, 3)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(64, 24)
        Me.btnSelect.TabIndex = 2
        Me.btnSelect.Text = "Select All"
        Me.btnSelect.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(6, 3)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(68, 24)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(1101, 3)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(68, 24)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        '
        'RadMenu1
        '
        Me.RadMenu1.Location = New System.Drawing.Point(0, 0)
        Me.RadMenu1.Name = "RadMenu1"
        Me.RadMenu1.Size = New System.Drawing.Size(1171, 20)
        Me.RadMenu1.TabIndex = 0
        '
        'frmDairyGatePass
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1171, 456)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.RadMenu1)
        Me.Name = "frmDairyGatePass"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "GatePass Entry"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.rpvpGatePass.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.rdBoth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbNonTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtShipToLocation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShipToLocation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDemandNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblOpKM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOpKM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtClKM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblClKM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkIndividualCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkGhee, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rgbItemType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.rgbItemType.ResumeLayout(False)
        Me.rgbItemType.PerformLayout()
        CType(Me.rbtn_IceCream, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtn_product, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtn_Milk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblClosingDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTripNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTripNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSupplyDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSupplyDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClKM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDistributorName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDistributorName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDriverMobNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDriverName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtGatepassDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLoadingSlip, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox3.ResumeLayout(False)
        Me.RadGroupBox3.PerformLayout()
        CType(Me.rbtnEvening, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnMorning, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnMultiGPReverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel34, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTollAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTollAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAgainstTransfer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSalesman, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotalCan, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCrateQty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRouteName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSalesman, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCanQty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRoute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLocDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtComments, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRemarks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTransporter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVehicleDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblpaymentpostdate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbitemtype, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblfullempty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblpaymentno, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnNew, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Gv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Gv1.ResumeLayout(False)
        Me.Gv1.PerformLayout()
        CType(Me.MyRadGridView1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyRadGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.rpvpCrateType.ResumeLayout(False)
        CType(Me.gvCrateType.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvCrateType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnBoothSlip, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGPCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSelect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Gv1 As common.UserControls.MyRadGridView
    Friend WithEvents lblSalesman As common.Controls.MyLabel
    Friend WithEvents lblpaymentno As common.Controls.MyLabel
    Friend WithEvents btnSave As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnClose As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnNew As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtVehicle As common.UserControls.txtFinder
    Friend WithEvents lblpaymentpostdate As common.Controls.MyLabel
    Friend WithEvents txtCode As common.UserControls.txtNavigator
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents RadMenu1 As Telerik.WinControls.UI.RadMenu
    Friend WithEvents btnPost As Telerik.WinControls.UI.RadButton
    Friend WithEvents cmbitemtype As common.Controls.MyComboBox
    Friend WithEvents lblfullempty As common.Controls.MyLabel
    Friend WithEvents RadSplitButton1 As Telerik.WinControls.UI.RadSplitButton
    Friend WithEvents RadMenuItem1 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem2 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents MyLabel3 As common.Controls.MyLabel
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents txtComments As common.Controls.MyTextBox
    Friend WithEvents txtRemarks As common.Controls.MyTextBox
    Friend WithEvents txtTransporter As common.Controls.MyTextBox
    Friend WithEvents lblVehicleDesc As common.Controls.MyTextBox
    Friend WithEvents btnSelect As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtLocDesc As common.Controls.MyTextBox
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents txtLocCode As common.UserControls.txtFinder
    Friend WithEvents btnGo As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnPrint As Telerik.WinControls.UI.RadButton
    Friend WithEvents lblRoute As common.Controls.MyLabel
    Friend WithEvents fndRouteNo As common.UserControls.txtFinder
    Friend WithEvents txtRouteName As common.Controls.MyTextBox
    Friend WithEvents txtCrateQty As common.Controls.MyTextBox
    Friend WithEvents txtCanQty As common.Controls.MyTextBox
    Friend WithEvents txtmultiBooking As common.UserControls.txtMultiSelectFinder
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents MyLabel6 As common.Controls.MyLabel
    Friend WithEvents lblTotalCan As common.Controls.MyLabel
    Friend WithEvents txtClKM As common.MyNumBox
    Friend WithEvents lblClKM As common.Controls.MyLabel
    Friend WithEvents txtOpKM As common.MyNumBox
    Friend WithEvents lblOpKM As common.Controls.MyLabel
    Friend WithEvents btnClKM As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnDelete As Telerik.WinControls.UI.RadButton
    Friend WithEvents UsLock1 As common.usLock
    Friend WithEvents btnReverse As Telerik.WinControls.UI.RadButton
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents txtSalesman As common.Controls.MyTextBox
    Friend WithEvents MyLabel8 As common.Controls.MyLabel
    Friend WithEvents lblClosingDate As common.Controls.MyLabel
    Friend WithEvents chkAgainstTransfer As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents MyLabel9 As common.Controls.MyLabel
    Friend WithEvents FndTransferNo As common.UserControls.txtFinder
    Friend WithEvents lblTollAmount As common.Controls.MyLabel
    Friend WithEvents txtTollAmount As common.MyNumBox
    Friend WithEvents MyLabel34 As common.Controls.MyLabel
    Friend WithEvents TxtMultiDairyGPassReverse As common.UserControls.txtMultiSelectFinder
    Friend WithEvents btnMultiGPReverse As RadButton
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents RadGroupBox3 As RadGroupBox
    Friend WithEvents rbtnEvening As RadRadioButton
    Friend WithEvents rbtnMorning As RadRadioButton
    Friend WithEvents MyLabel11 As common.Controls.MyLabel
    Friend WithEvents txtLoadingSlip As common.Controls.MyTextBox
    Friend WithEvents txtGatepassDate As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel12 As common.Controls.MyLabel
    Friend WithEvents btnGPCancel As RadButton
    Friend WithEvents txtDriverName As common.Controls.MyTextBox
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents txtDriverMobNo As common.Controls.MyTextBox
    Friend WithEvents MyLabel14 As common.Controls.MyLabel
    Friend WithEvents btnPrint2 As RadButton
    Friend WithEvents txtDistributorName As common.Controls.MyTextBox
    Friend WithEvents lblDistributorName As common.Controls.MyLabel
    Friend WithEvents txtSupplyDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblSupplyDate As common.Controls.MyLabel
    Friend WithEvents txtTripNo As common.Controls.MyTextBox
    Friend WithEvents lblTripNo As common.Controls.MyLabel
    Friend WithEvents btnBoothSlip As RadSplitButton
    Friend WithEvents RadMenuItem3 As RadMenuItem
    Friend WithEvents btnExcel As RadMenuItem
    Friend WithEvents MyRadGridView1 As common.UserControls.MyRadGridView
    Friend WithEvents rgbItemType As RadGroupBox
    Friend WithEvents rbtn_IceCream As RadRadioButton
    Friend WithEvents rbtn_product As RadRadioButton
    Friend WithEvents rbtn_Milk As RadRadioButton
    Friend WithEvents chkGhee As RadCheckBox
    Friend WithEvents btnHistory As RadButton
    Friend WithEvents RadMenuItem4 As RadMenuItem
    Friend WithEvents lblDemandNo As common.Controls.MyLabel
    Friend WithEvents txtDemandNo As common.UserControls.txtFinder
    Friend WithEvents chkIndividualCustomer As RadCheckBox
    Friend WithEvents RadPageView1 As RadPageView
    Friend WithEvents rpvpGatePass As RadPageViewPage
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents rpvpCrateType As RadPageViewPage
    Friend WithEvents gvCrateType As common.UserControls.MyRadGridView
    Friend WithEvents txtShipToLocation As common.Controls.MyTextBox
    Friend WithEvents lblShipToLocation As common.Controls.MyLabel
    Friend WithEvents RadGroupBox1 As RadGroupBox
    Friend WithEvents rdBoth As RadRadioButton
    Friend WithEvents rdbNonTaxable As RadRadioButton
    Friend WithEvents rdbTaxable As RadRadioButton
End Class

