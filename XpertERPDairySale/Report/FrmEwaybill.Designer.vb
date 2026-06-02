<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmEwaybill
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
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtVehicleNo = New common.Controls.MyTextBox()
        Me.cmbEwaybillType = New common.Controls.MyComboBox()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.lblewbno = New common.Controls.MyLabel()
        Me.lblDate = New common.Controls.MyLabel()
        Me.lblTransID = New common.Controls.MyLabel()
        Me.txtewbno = New common.UserControls.txtFinder()
        Me.txtTransID = New common.UserControls.txtFinder()
        Me.txtRemaningDistance = New common.Controls.MyTextBox()
        Me.lblReasoncode = New common.Controls.MyLabel()
        Me.lblRemDistrance = New common.Controls.MyLabel()
        Me.cmbReasonForUpdateVehicle = New System.Windows.Forms.ComboBox()
        Me.cmbConsignmentStatus = New System.Windows.Forms.ComboBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.lblConsignmentstatus = New common.Controls.MyLabel()
        Me.cmbExtendValidityReason = New System.Windows.Forms.ComboBox()
        Me.RadLabel15 = New common.Controls.MyLabel()
        Me.lblVehicleNo = New common.Controls.MyLabel()
        Me.txtLocation = New common.UserControls.txtFinder()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.btnGo = New Telerik.WinControls.UI.RadButton()
        Me.btnclose = New Telerik.WinControls.UI.RadButton()
        Me.btnreset = New Telerik.WinControls.UI.RadButton()
        Me.txtclsewb = New common.Controls.MyDateTimePicker()
        Me.lblclsewb = New common.Controls.MyLabel()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.txtVehicleNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbEwaybillType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblewbno, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTransID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRemaningDistance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReasoncode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRemDistrance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblConsignmentstatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVehicleNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnclose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnreset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtclsewb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblclsewb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnGo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnclose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnreset)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 413
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtclsewb)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblclsewb)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtVehicleNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmbEwaybillType)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblewbno)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblTransID)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtewbno)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtTransID)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtRemaningDistance)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblReasoncode)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblRemDistrance)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmbReasonForUpdateVehicle)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmbConsignmentStatus)
        Me.SplitContainer2.Panel1.Controls.Add(Me.MyLabel1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblConsignmentstatus)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmbExtendValidityReason)
        Me.SplitContainer2.Panel1.Controls.Add(Me.RadLabel15)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblVehicleNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtLocation)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gv1)
        Me.SplitContainer2.Size = New System.Drawing.Size(800, 413)
        Me.SplitContainer2.SplitterDistance = 162
        Me.SplitContainer2.TabIndex = 1495
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "E-Way Bill Service Type"
        '
        'txtVehicleNo
        '
        Me.txtVehicleNo.CalculationExpression = Nothing
        Me.txtVehicleNo.FieldCode = Nothing
        Me.txtVehicleNo.FieldDesc = Nothing
        Me.txtVehicleNo.FieldMaxLength = 0
        Me.txtVehicleNo.FieldName = Nothing
        Me.txtVehicleNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVehicleNo.isCalculatedField = False
        Me.txtVehicleNo.IsSourceFromTable = False
        Me.txtVehicleNo.IsSourceFromValueList = False
        Me.txtVehicleNo.IsUnique = False
        Me.txtVehicleNo.Location = New System.Drawing.Point(124, 104)
        Me.txtVehicleNo.MaxLength = 200
        Me.txtVehicleNo.MendatroryField = False
        Me.txtVehicleNo.MyLinkLable1 = Nothing
        Me.txtVehicleNo.MyLinkLable2 = Nothing
        Me.txtVehicleNo.Name = "txtVehicleNo"
        Me.txtVehicleNo.ReferenceFieldDesc = Nothing
        Me.txtVehicleNo.ReferenceFieldName = Nothing
        Me.txtVehicleNo.ReferenceTableName = Nothing
        Me.txtVehicleNo.Size = New System.Drawing.Size(125, 18)
        Me.txtVehicleNo.TabIndex = 1494
        '
        'cmbEwaybillType
        '
        Me.cmbEwaybillType.AutoCompleteDisplayMember = Nothing
        Me.cmbEwaybillType.AutoCompleteValueMember = Nothing
        Me.cmbEwaybillType.CalculationExpression = Nothing
        Me.cmbEwaybillType.DropDownAnimationEnabled = True
        Me.cmbEwaybillType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbEwaybillType.FieldCode = Nothing
        Me.cmbEwaybillType.FieldDesc = Nothing
        Me.cmbEwaybillType.FieldMaxLength = 0
        Me.cmbEwaybillType.FieldName = Nothing
        Me.cmbEwaybillType.isCalculatedField = False
        Me.cmbEwaybillType.IsSourceFromTable = False
        Me.cmbEwaybillType.IsSourceFromValueList = False
        Me.cmbEwaybillType.IsUnique = False
        Me.cmbEwaybillType.Location = New System.Drawing.Point(140, 5)
        Me.cmbEwaybillType.MendatroryField = True
        Me.cmbEwaybillType.MyLinkLable1 = Nothing
        Me.cmbEwaybillType.MyLinkLable2 = Nothing
        Me.cmbEwaybillType.Name = "cmbEwaybillType"
        Me.cmbEwaybillType.ReferenceFieldDesc = Nothing
        Me.cmbEwaybillType.ReferenceFieldName = Nothing
        Me.cmbEwaybillType.ReferenceTableName = Nothing
        Me.cmbEwaybillType.Size = New System.Drawing.Size(236, 20)
        Me.cmbEwaybillType.TabIndex = 1474
        '
        'txtDate
        '
        Me.txtDate.CalculationExpression = Nothing
        Me.txtDate.CustomFormat = "dd/MM/yyyy"
        Me.txtDate.FieldCode = Nothing
        Me.txtDate.FieldDesc = Nothing
        Me.txtDate.FieldMaxLength = 0
        Me.txtDate.FieldName = Nothing
        Me.txtDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDate.isCalculatedField = False
        Me.txtDate.IsSourceFromTable = False
        Me.txtDate.IsSourceFromValueList = False
        Me.txtDate.IsUnique = False
        Me.txtDate.Location = New System.Drawing.Point(125, 126)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Nothing
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(124, 18)
        Me.txtDate.TabIndex = 1493
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011"
        Me.txtDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'lblewbno
        '
        Me.lblewbno.FieldName = Nothing
        Me.lblewbno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblewbno.Location = New System.Drawing.Point(25, 58)
        Me.lblewbno.Name = "lblewbno"
        Me.lblewbno.Size = New System.Drawing.Size(73, 16)
        Me.lblewbno.TabIndex = 1475
        Me.lblewbno.Text = "Eway Bill No."
        '
        'lblDate
        '
        Me.lblDate.FieldName = Nothing
        Me.lblDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.Location = New System.Drawing.Point(28, 129)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(79, 16)
        Me.lblDate.TabIndex = 1492
        Me.lblDate.Text = "Eway Bill Date"
        '
        'lblTransID
        '
        Me.lblTransID.FieldName = Nothing
        Me.lblTransID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTransID.Location = New System.Drawing.Point(25, 80)
        Me.lblTransID.Name = "lblTransID"
        Me.lblTransID.Size = New System.Drawing.Size(75, 16)
        Me.lblTransID.TabIndex = 1477
        Me.lblTransID.Text = "Transpoter ID"
        '
        'txtewbno
        '
        Me.txtewbno.CalculationExpression = Nothing
        Me.txtewbno.FieldCode = Nothing
        Me.txtewbno.FieldDesc = Nothing
        Me.txtewbno.FieldMaxLength = 0
        Me.txtewbno.FieldName = Nothing
        Me.txtewbno.isCalculatedField = False
        Me.txtewbno.IsSourceFromTable = False
        Me.txtewbno.IsSourceFromValueList = False
        Me.txtewbno.IsUnique = False
        Me.txtewbno.Location = New System.Drawing.Point(124, 57)
        Me.txtewbno.MendatroryField = False
        Me.txtewbno.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtewbno.MyLinkLable1 = Nothing
        Me.txtewbno.MyLinkLable2 = Nothing
        Me.txtewbno.MyReadOnly = False
        Me.txtewbno.MyShowMasterFormButton = False
        Me.txtewbno.Name = "txtewbno"
        Me.txtewbno.ReferenceFieldDesc = Nothing
        Me.txtewbno.ReferenceFieldName = Nothing
        Me.txtewbno.ReferenceTableName = Nothing
        Me.txtewbno.Size = New System.Drawing.Size(125, 19)
        Me.txtewbno.TabIndex = 1491
        Me.txtewbno.Value = ""
        '
        'txtTransID
        '
        Me.txtTransID.CalculationExpression = Nothing
        Me.txtTransID.FieldCode = Nothing
        Me.txtTransID.FieldDesc = Nothing
        Me.txtTransID.FieldMaxLength = 0
        Me.txtTransID.FieldName = Nothing
        Me.txtTransID.isCalculatedField = False
        Me.txtTransID.IsSourceFromTable = False
        Me.txtTransID.IsSourceFromValueList = False
        Me.txtTransID.IsUnique = False
        Me.txtTransID.Location = New System.Drawing.Point(124, 81)
        Me.txtTransID.MendatroryField = False
        Me.txtTransID.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTransID.MyLinkLable1 = Nothing
        Me.txtTransID.MyLinkLable2 = Nothing
        Me.txtTransID.MyReadOnly = False
        Me.txtTransID.MyShowMasterFormButton = False
        Me.txtTransID.Name = "txtTransID"
        Me.txtTransID.ReferenceFieldDesc = Nothing
        Me.txtTransID.ReferenceFieldName = Nothing
        Me.txtTransID.ReferenceTableName = Nothing
        Me.txtTransID.Size = New System.Drawing.Size(125, 19)
        Me.txtTransID.TabIndex = 1478
        Me.txtTransID.Value = ""
        '
        'txtRemaningDistance
        '
        Me.txtRemaningDistance.CalculationExpression = Nothing
        Me.txtRemaningDistance.FieldCode = Nothing
        Me.txtRemaningDistance.FieldDesc = Nothing
        Me.txtRemaningDistance.FieldMaxLength = 0
        Me.txtRemaningDistance.FieldName = Nothing
        Me.txtRemaningDistance.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemaningDistance.isCalculatedField = False
        Me.txtRemaningDistance.IsSourceFromTable = False
        Me.txtRemaningDistance.IsSourceFromValueList = False
        Me.txtRemaningDistance.IsUnique = False
        Me.txtRemaningDistance.Location = New System.Drawing.Point(384, 128)
        Me.txtRemaningDistance.MaxLength = 200
        Me.txtRemaningDistance.MendatroryField = False
        Me.txtRemaningDistance.MyLinkLable1 = Nothing
        Me.txtRemaningDistance.MyLinkLable2 = Nothing
        Me.txtRemaningDistance.Name = "txtRemaningDistance"
        Me.txtRemaningDistance.ReferenceFieldDesc = Nothing
        Me.txtRemaningDistance.ReferenceFieldName = Nothing
        Me.txtRemaningDistance.ReferenceTableName = Nothing
        Me.txtRemaningDistance.Size = New System.Drawing.Size(125, 18)
        Me.txtRemaningDistance.TabIndex = 1490
        '
        'lblReasoncode
        '
        Me.lblReasoncode.FieldName = Nothing
        Me.lblReasoncode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReasoncode.Location = New System.Drawing.Point(255, 59)
        Me.lblReasoncode.Name = "lblReasoncode"
        Me.lblReasoncode.Size = New System.Drawing.Size(75, 16)
        Me.lblReasoncode.TabIndex = 1479
        Me.lblReasoncode.Text = "Reason Code"
        '
        'lblRemDistrance
        '
        Me.lblRemDistrance.FieldName = Nothing
        Me.lblRemDistrance.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRemDistrance.Location = New System.Drawing.Point(257, 128)
        Me.lblRemDistrance.Name = "lblRemDistrance"
        Me.lblRemDistrance.Size = New System.Drawing.Size(105, 16)
        Me.lblRemDistrance.TabIndex = 1489
        Me.lblRemDistrance.Text = "Remaning Distance"
        '
        'cmbReasonForUpdateVehicle
        '
        Me.cmbReasonForUpdateVehicle.FormattingEnabled = True
        Me.cmbReasonForUpdateVehicle.Location = New System.Drawing.Point(384, 55)
        Me.cmbReasonForUpdateVehicle.Name = "cmbReasonForUpdateVehicle"
        Me.cmbReasonForUpdateVehicle.Size = New System.Drawing.Size(232, 21)
        Me.cmbReasonForUpdateVehicle.TabIndex = 1480
        '
        'cmbConsignmentStatus
        '
        Me.cmbConsignmentStatus.FormattingEnabled = True
        Me.cmbConsignmentStatus.Location = New System.Drawing.Point(384, 103)
        Me.cmbConsignmentStatus.Name = "cmbConsignmentStatus"
        Me.cmbConsignmentStatus.Size = New System.Drawing.Size(232, 21)
        Me.cmbConsignmentStatus.TabIndex = 1488
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(255, 84)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(123, 16)
        Me.MyLabel1.TabIndex = 1481
        Me.MyLabel1.Text = "Extend Validity Reason"
        '
        'lblConsignmentstatus
        '
        Me.lblConsignmentstatus.FieldName = Nothing
        Me.lblConsignmentstatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConsignmentstatus.Location = New System.Drawing.Point(255, 106)
        Me.lblConsignmentstatus.Name = "lblConsignmentstatus"
        Me.lblConsignmentstatus.Size = New System.Drawing.Size(108, 16)
        Me.lblConsignmentstatus.TabIndex = 1487
        Me.lblConsignmentstatus.Text = "Consignment Status"
        '
        'cmbExtendValidityReason
        '
        Me.cmbExtendValidityReason.FormattingEnabled = True
        Me.cmbExtendValidityReason.Location = New System.Drawing.Point(384, 79)
        Me.cmbExtendValidityReason.Name = "cmbExtendValidityReason"
        Me.cmbExtendValidityReason.Size = New System.Drawing.Size(232, 21)
        Me.cmbExtendValidityReason.TabIndex = 1482
        '
        'RadLabel15
        '
        Me.RadLabel15.FieldName = Nothing
        Me.RadLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel15.Location = New System.Drawing.Point(28, 36)
        Me.RadLabel15.Name = "RadLabel15"
        Me.RadLabel15.Size = New System.Drawing.Size(49, 16)
        Me.RadLabel15.TabIndex = 1486
        Me.RadLabel15.Text = "Location"
        '
        'lblVehicleNo
        '
        Me.lblVehicleNo.FieldName = Nothing
        Me.lblVehicleNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVehicleNo.Location = New System.Drawing.Point(25, 105)
        Me.lblVehicleNo.Name = "lblVehicleNo"
        Me.lblVehicleNo.Size = New System.Drawing.Size(61, 16)
        Me.lblVehicleNo.TabIndex = 1483
        Me.lblVehicleNo.Text = "Vehicle No"
        '
        'txtLocation
        '
        Me.txtLocation.CalculationExpression = Nothing
        Me.txtLocation.FieldCode = Nothing
        Me.txtLocation.FieldDesc = Nothing
        Me.txtLocation.FieldMaxLength = 0
        Me.txtLocation.FieldName = Nothing
        Me.txtLocation.isCalculatedField = False
        Me.txtLocation.IsSourceFromTable = False
        Me.txtLocation.IsSourceFromValueList = False
        Me.txtLocation.IsUnique = False
        Me.txtLocation.Location = New System.Drawing.Point(125, 32)
        Me.txtLocation.MendatroryField = True
        Me.txtLocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocation.MyLinkLable1 = Me.RadLabel15
        Me.txtLocation.MyLinkLable2 = Nothing
        Me.txtLocation.MyReadOnly = False
        Me.txtLocation.MyShowMasterFormButton = False
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.ReferenceFieldDesc = Nothing
        Me.txtLocation.ReferenceFieldName = Nothing
        Me.txtLocation.ReferenceTableName = Nothing
        Me.txtLocation.Size = New System.Drawing.Size(125, 20)
        Me.txtLocation.TabIndex = 1485
        Me.txtLocation.Value = ""
        '
        'gv1
        '
        Me.gv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gv1.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gv1.MasterTemplate.AllowAddNewRow = False
        Me.gv1.MasterTemplate.EnableFiltering = True
        Me.gv1.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gv1.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv1.MasterTemplate.ViewDefinition = TableViewDefinition1
        Me.gv1.MyExportAPI = False
        Me.gv1.MyExportFilePath = ""
        Me.gv1.MyStopExport = False
        Me.gv1.Name = "gv1"
        Me.gv1.ShowHeaderCellButtons = True
        Me.gv1.Size = New System.Drawing.Size(800, 247)
        Me.gv1.TabIndex = 2
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'btnGo
        '
        Me.btnGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGo.Location = New System.Drawing.Point(93, 7)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(70, 19)
        Me.btnGo.TabIndex = 89
        Me.btnGo.Text = "Go >>"
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.Location = New System.Drawing.Point(714, 7)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(70, 19)
        Me.btnclose.TabIndex = 88
        Me.btnclose.Text = "Close"
        '
        'btnreset
        '
        Me.btnreset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnreset.Location = New System.Drawing.Point(17, 7)
        Me.btnreset.Name = "btnreset"
        Me.btnreset.Size = New System.Drawing.Size(70, 19)
        Me.btnreset.TabIndex = 87
        Me.btnreset.Text = "Reset"
        '
        'txtclsewb
        '
        Me.txtclsewb.CalculationExpression = Nothing
        Me.txtclsewb.CustomFormat = "dd/MM/yyyy"
        Me.txtclsewb.FieldCode = Nothing
        Me.txtclsewb.FieldDesc = Nothing
        Me.txtclsewb.FieldMaxLength = 0
        Me.txtclsewb.FieldName = Nothing
        Me.txtclsewb.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtclsewb.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtclsewb.isCalculatedField = False
        Me.txtclsewb.IsSourceFromTable = False
        Me.txtclsewb.IsSourceFromValueList = False
        Me.txtclsewb.IsUnique = False
        Me.txtclsewb.Location = New System.Drawing.Point(384, 33)
        Me.txtclsewb.MendatroryField = False
        Me.txtclsewb.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtclsewb.MyLinkLable1 = Nothing
        Me.txtclsewb.MyLinkLable2 = Nothing
        Me.txtclsewb.Name = "txtclsewb"
        Me.txtclsewb.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtclsewb.ReferenceFieldDesc = Nothing
        Me.txtclsewb.ReferenceFieldName = Nothing
        Me.txtclsewb.ReferenceTableName = Nothing
        Me.txtclsewb.Size = New System.Drawing.Size(124, 18)
        Me.txtclsewb.TabIndex = 1496
        Me.txtclsewb.TabStop = False
        Me.txtclsewb.Text = "13/06/2011"
        Me.txtclsewb.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'lblclsewb
        '
        Me.lblclsewb.FieldName = Nothing
        Me.lblclsewb.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblclsewb.Location = New System.Drawing.Point(257, 36)
        Me.lblclsewb.Name = "lblclsewb"
        Me.lblclsewb.Size = New System.Drawing.Size(111, 16)
        Me.lblclsewb.TabIndex = 1495
        Me.lblclsewb.Text = "Eway Bill Close Date"
        '
        'FrmEwaybill
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "FrmEwaybill"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "Eway Bill API"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.txtVehicleNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbEwaybillType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblewbno, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTransID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRemaningDistance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReasoncode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRemDistrance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblConsignmentstatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVehicleNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnclose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnreset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtclsewb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblclsewb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label1 As Label
    Friend WithEvents btnGo As RadButton
    Friend WithEvents btnclose As RadButton
    Friend WithEvents btnreset As RadButton
    Friend WithEvents cmbEwaybillType As common.Controls.MyComboBox
    Friend WithEvents lblewbno As common.Controls.MyLabel
    Friend WithEvents lblTransID As common.Controls.MyLabel
    Friend WithEvents txtTransID As common.UserControls.txtFinder
    Friend WithEvents lblReasoncode As common.Controls.MyLabel
    Friend WithEvents cmbReasonForUpdateVehicle As ComboBox
    Friend WithEvents cmbExtendValidityReason As ComboBox
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents lblVehicleNo As common.Controls.MyLabel
    Friend WithEvents RadLabel15 As common.Controls.MyLabel
    Friend WithEvents txtLocation As common.UserControls.txtFinder
    Friend WithEvents cmbConsignmentStatus As ComboBox
    Friend WithEvents lblConsignmentstatus As common.Controls.MyLabel
    Friend WithEvents txtRemaningDistance As common.Controls.MyTextBox
    Friend WithEvents lblRemDistrance As common.Controls.MyLabel
    Friend WithEvents txtewbno As common.UserControls.txtFinder
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblDate As common.Controls.MyLabel
    Friend WithEvents txtVehicleNo As common.Controls.MyTextBox
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents txtclsewb As common.Controls.MyDateTimePicker
    Friend WithEvents lblclsewb As common.Controls.MyLabel
End Class
