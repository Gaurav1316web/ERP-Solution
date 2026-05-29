<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQuickDemand
    Inherits FrmMainTranScreen

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.cmbShift = New common.Controls.MyComboBox()
        Me.txtAddress = New common.Controls.MyTextBox()
        Me.lblAddress = New common.Controls.MyLabel()
        Me.txtDistributor = New common.Controls.MyTextBox()
        Me.lblDistributor = New common.Controls.MyLabel()
        Me.txtBoothName = New common.Controls.MyTextBox()
        Me.lblBoothName = New common.Controls.MyLabel()
        Me.txtShift = New common.Controls.MyTextBox()
        Me.lblShift = New common.Controls.MyLabel()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.lblDate = New common.Controls.MyLabel()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.btnclose = New Telerik.WinControls.UI.RadButton()
        Me.btnExport = New Telerik.WinControls.UI.RadSplitButton()
        Me.rmiImport = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiExcel = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.btnReset = New Telerik.WinControls.UI.RadButton()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.cmbShift, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDistributor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDistributor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBoothName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblBoothName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtShift, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShift, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnclose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnExport, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReset)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnclose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnExport)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 414
        Me.SplitContainer1.TabIndex = 1
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmbShift)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtAddress)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblAddress)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDistributor)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblDistributor)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtBoothName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblBoothName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtShift)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblShift)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblDate)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gv1)
        Me.SplitContainer2.Size = New System.Drawing.Size(800, 414)
        Me.SplitContainer2.SplitterDistance = 66
        Me.SplitContainer2.TabIndex = 0
        '
        'cmbShift
        '
        Me.cmbShift.AutoCompleteDisplayMember = Nothing
        Me.cmbShift.AutoCompleteValueMember = Nothing
        Me.cmbShift.CalculationExpression = Nothing
        Me.cmbShift.DropDownAnimationEnabled = True
        Me.cmbShift.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbShift.FieldCode = Nothing
        Me.cmbShift.FieldDesc = Nothing
        Me.cmbShift.FieldMaxLength = 0
        Me.cmbShift.FieldName = Nothing
        Me.cmbShift.isCalculatedField = False
        Me.cmbShift.IsSourceFromTable = False
        Me.cmbShift.IsSourceFromValueList = False
        Me.cmbShift.IsUnique = False
        Me.cmbShift.Location = New System.Drawing.Point(297, 10)
        Me.cmbShift.MendatroryField = True
        Me.cmbShift.MyLinkLable1 = Nothing
        Me.cmbShift.MyLinkLable2 = Nothing
        Me.cmbShift.Name = "cmbShift"
        Me.cmbShift.ReferenceFieldDesc = Nothing
        Me.cmbShift.ReferenceFieldName = Nothing
        Me.cmbShift.ReferenceTableName = Nothing
        Me.cmbShift.Size = New System.Drawing.Size(156, 20)
        Me.cmbShift.TabIndex = 1474
        '
        'txtAddress
        '
        Me.txtAddress.CalculationExpression = Nothing
        Me.txtAddress.FieldCode = Nothing
        Me.txtAddress.FieldDesc = Nothing
        Me.txtAddress.FieldMaxLength = 0
        Me.txtAddress.FieldName = Nothing
        Me.txtAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress.isCalculatedField = False
        Me.txtAddress.IsSourceFromTable = False
        Me.txtAddress.IsSourceFromValueList = False
        Me.txtAddress.IsUnique = False
        Me.txtAddress.Location = New System.Drawing.Point(513, 41)
        Me.txtAddress.MaxLength = 50
        Me.txtAddress.MendatroryField = False
        Me.txtAddress.Modified = True
        Me.txtAddress.MyLinkLable1 = Nothing
        Me.txtAddress.MyLinkLable2 = Nothing
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.ReadOnly = True
        Me.txtAddress.ReferenceFieldDesc = Nothing
        Me.txtAddress.ReferenceFieldName = Nothing
        Me.txtAddress.ReferenceTableName = Nothing
        Me.txtAddress.Size = New System.Drawing.Size(214, 18)
        Me.txtAddress.TabIndex = 1457
        '
        'lblAddress
        '
        Me.lblAddress.FieldName = Nothing
        Me.lblAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress.Location = New System.Drawing.Point(460, 42)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(48, 16)
        Me.lblAddress.TabIndex = 1456
        Me.lblAddress.Text = "Address"
        '
        'txtDistributor
        '
        Me.txtDistributor.CalculationExpression = Nothing
        Me.txtDistributor.FieldCode = Nothing
        Me.txtDistributor.FieldDesc = Nothing
        Me.txtDistributor.FieldMaxLength = 0
        Me.txtDistributor.FieldName = Nothing
        Me.txtDistributor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDistributor.isCalculatedField = False
        Me.txtDistributor.IsSourceFromTable = False
        Me.txtDistributor.IsSourceFromValueList = False
        Me.txtDistributor.IsUnique = False
        Me.txtDistributor.Location = New System.Drawing.Point(297, 40)
        Me.txtDistributor.MaxLength = 50
        Me.txtDistributor.MendatroryField = False
        Me.txtDistributor.Modified = True
        Me.txtDistributor.MyLinkLable1 = Nothing
        Me.txtDistributor.MyLinkLable2 = Nothing
        Me.txtDistributor.Name = "txtDistributor"
        Me.txtDistributor.ReadOnly = True
        Me.txtDistributor.ReferenceFieldDesc = Nothing
        Me.txtDistributor.ReferenceFieldName = Nothing
        Me.txtDistributor.ReferenceTableName = Nothing
        Me.txtDistributor.Size = New System.Drawing.Size(156, 18)
        Me.txtDistributor.TabIndex = 1455
        '
        'lblDistributor
        '
        Me.lblDistributor.FieldName = Nothing
        Me.lblDistributor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistributor.Location = New System.Drawing.Point(234, 43)
        Me.lblDistributor.Name = "lblDistributor"
        Me.lblDistributor.Size = New System.Drawing.Size(58, 16)
        Me.lblDistributor.TabIndex = 1454
        Me.lblDistributor.Text = "Distributor"
        '
        'txtBoothName
        '
        Me.txtBoothName.CalculationExpression = Nothing
        Me.txtBoothName.FieldCode = Nothing
        Me.txtBoothName.FieldDesc = Nothing
        Me.txtBoothName.FieldMaxLength = 0
        Me.txtBoothName.FieldName = Nothing
        Me.txtBoothName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBoothName.isCalculatedField = False
        Me.txtBoothName.IsSourceFromTable = False
        Me.txtBoothName.IsSourceFromValueList = False
        Me.txtBoothName.IsUnique = False
        Me.txtBoothName.Location = New System.Drawing.Point(92, 41)
        Me.txtBoothName.MaxLength = 50
        Me.txtBoothName.MendatroryField = False
        Me.txtBoothName.Modified = True
        Me.txtBoothName.MyLinkLable1 = Nothing
        Me.txtBoothName.MyLinkLable2 = Nothing
        Me.txtBoothName.Name = "txtBoothName"
        Me.txtBoothName.ReadOnly = True
        Me.txtBoothName.ReferenceFieldDesc = Nothing
        Me.txtBoothName.ReferenceFieldName = Nothing
        Me.txtBoothName.ReferenceTableName = Nothing
        Me.txtBoothName.Size = New System.Drawing.Size(134, 18)
        Me.txtBoothName.TabIndex = 1453
        '
        'lblBoothName
        '
        Me.lblBoothName.FieldName = Nothing
        Me.lblBoothName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoothName.Location = New System.Drawing.Point(15, 42)
        Me.lblBoothName.Name = "lblBoothName"
        Me.lblBoothName.Size = New System.Drawing.Size(69, 16)
        Me.lblBoothName.TabIndex = 1452
        Me.lblBoothName.Text = "Booth Name"
        '
        'txtShift
        '
        Me.txtShift.CalculationExpression = Nothing
        Me.txtShift.FieldCode = Nothing
        Me.txtShift.FieldDesc = Nothing
        Me.txtShift.FieldMaxLength = 0
        Me.txtShift.FieldName = Nothing
        Me.txtShift.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShift.isCalculatedField = False
        Me.txtShift.IsSourceFromTable = False
        Me.txtShift.IsSourceFromValueList = False
        Me.txtShift.IsUnique = False
        Me.txtShift.Location = New System.Drawing.Point(470, 9)
        Me.txtShift.MaxLength = 50
        Me.txtShift.MendatroryField = False
        Me.txtShift.Modified = True
        Me.txtShift.MyLinkLable1 = Nothing
        Me.txtShift.MyLinkLable2 = Nothing
        Me.txtShift.Name = "txtShift"
        Me.txtShift.ReadOnly = True
        Me.txtShift.ReferenceFieldDesc = Nothing
        Me.txtShift.ReferenceFieldName = Nothing
        Me.txtShift.ReferenceTableName = Nothing
        Me.txtShift.Size = New System.Drawing.Size(90, 18)
        Me.txtShift.TabIndex = 1451
        Me.txtShift.Visible = False
        '
        'lblShift
        '
        Me.lblShift.FieldName = Nothing
        Me.lblShift.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShift.Location = New System.Drawing.Point(234, 12)
        Me.lblShift.Name = "lblShift"
        Me.lblShift.Size = New System.Drawing.Size(29, 16)
        Me.lblShift.TabIndex = 1450
        Me.lblShift.Text = "Shift"
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
        Me.txtDate.Location = New System.Drawing.Point(91, 12)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Nothing
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(133, 18)
        Me.txtDate.TabIndex = 58
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011"
        Me.txtDate.Value = New Date(2011, 6, 13, 0, 0, 0, 0)
        '
        'lblDate
        '
        Me.lblDate.FieldName = Nothing
        Me.lblDate.Location = New System.Drawing.Point(19, 12)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(30, 18)
        Me.lblDate.TabIndex = 57
        Me.lblDate.Text = "Date"
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
        Me.gv1.ShowGroupPanel = False
        Me.gv1.ShowHeaderCellButtons = True
        Me.gv1.Size = New System.Drawing.Size(800, 344)
        Me.gv1.TabIndex = 2
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.Location = New System.Drawing.Point(718, 6)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(70, 19)
        Me.btnclose.TabIndex = 88
        Me.btnclose.Text = "Close"
        '
        'btnExport
        '
        Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExport.Items.AddRange(New Telerik.WinControls.RadItem() {Me.rmiImport, Me.rmiExcel})
        Me.btnExport.Location = New System.Drawing.Point(99, 6)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(123, 19)
        Me.btnExport.TabIndex = 87
        Me.btnExport.Text = "Import/Export"
        '
        'rmiImport
        '
        Me.rmiImport.Name = "rmiImport"
        Me.rmiImport.Text = "Import"
        Me.rmiImport.UseCompatibleTextRendering = False
        '
        'rmiExcel
        '
        Me.rmiExcel.Name = "rmiExcel"
        Me.rmiExcel.Text = "Export"
        Me.rmiExcel.UseCompatibleTextRendering = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(23, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(70, 19)
        Me.btnSave.TabIndex = 86
        Me.btnSave.Text = "Save"
        '
        'btnReset
        '
        Me.btnReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReset.Location = New System.Drawing.Point(228, 6)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(70, 19)
        Me.btnReset.TabIndex = 89
        Me.btnReset.Text = "Reset"
        '
        'frmQuickDemand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frmQuickDemand"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "Quick Demand"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.cmbShift, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAddress, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddress, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDistributor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDistributor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBoothName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblBoothName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtShift, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShift, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnclose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnExport, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents cmbShift As common.Controls.MyComboBox
    Friend WithEvents txtAddress As common.Controls.MyTextBox
    Friend WithEvents lblAddress As common.Controls.MyLabel
    Friend WithEvents txtDistributor As common.Controls.MyTextBox
    Friend WithEvents lblDistributor As common.Controls.MyLabel
    Friend WithEvents txtBoothName As common.Controls.MyTextBox
    Friend WithEvents lblBoothName As common.Controls.MyLabel
    Friend WithEvents txtShift As common.Controls.MyTextBox
    Friend WithEvents lblShift As common.Controls.MyLabel
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblDate As common.Controls.MyLabel
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents btnclose As RadButton
    Friend WithEvents btnExport As RadSplitButton
    Friend WithEvents rmiImport As RadMenuItem
    Friend WithEvents rmiExcel As RadMenuItem
    Friend WithEvents btnSave As RadButton
    Friend WithEvents btnReset As RadButton
End Class
