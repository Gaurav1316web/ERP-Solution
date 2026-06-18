<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmperdayDetailRpt
    'Inherits System.Windows.Forms.Form
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rdbDocDate = New common.Controls.MyRadioButton()
        Me.rdbSuppltDate = New common.Controls.MyRadioButton()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.MyRadioButton3 = New common.Controls.MyRadioButton()
        Me.rdbMilk = New common.Controls.MyRadioButton()
        Me.rdbProduct = New common.Controls.MyRadioButton()
        Me.RadGroupBox5 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rdbInvBoth = New common.Controls.MyRadioButton()
        Me.rbtnNonTaxable = New common.Controls.MyRadioButton()
        Me.rbtnTaxable = New common.Controls.MyRadioButton()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.txtMultiCustomer = New common.UserControls.txtMultiSelectFinder()
        Me.RadGroupBox4 = New Telerik.WinControls.UI.RadGroupBox()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.ToDate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.fromdate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.btnReset = New Telerik.WinControls.UI.RadButton()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.rdbDocDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbSuppltDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.MyRadioButton3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbMilk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox5.SuspendLayout()
        CType(Me.rdbInvBoth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnNonTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox4.SuspendLayout()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ToDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fromdate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadGroupBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadGroupBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadGroupBox5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.MyLabel13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtMultiCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadGroupBox4)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReset)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 409
        Me.SplitContainer1.TabIndex = 0
        '
        'RadGroupBox2
        '
        Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox2.Controls.Add(Me.rdbDocDate)
        Me.RadGroupBox2.Controls.Add(Me.rdbSuppltDate)
        Me.RadGroupBox2.HeaderText = ""
        Me.RadGroupBox2.Location = New System.Drawing.Point(513, 35)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(205, 40)
        Me.RadGroupBox2.TabIndex = 1511
        '
        'rdbDocDate
        '
        Me.rdbDocDate.Location = New System.Drawing.Point(6, 11)
        Me.rdbDocDate.MyLinkLable1 = Nothing
        Me.rdbDocDate.MyLinkLable2 = Nothing
        Me.rdbDocDate.Name = "rdbDocDate"
        Me.rdbDocDate.Size = New System.Drawing.Size(99, 18)
        Me.rdbDocDate.TabIndex = 396
        Me.rdbDocDate.TabStop = False
        Me.rdbDocDate.Text = "Document Date"
        '
        'rdbSuppltDate
        '
        Me.rdbSuppltDate.Location = New System.Drawing.Point(107, 11)
        Me.rdbSuppltDate.MyLinkLable1 = Nothing
        Me.rdbSuppltDate.MyLinkLable2 = Nothing
        Me.rdbSuppltDate.Name = "rdbSuppltDate"
        Me.rdbSuppltDate.Size = New System.Drawing.Size(81, 18)
        Me.rdbSuppltDate.TabIndex = 391
        Me.rdbSuppltDate.TabStop = False
        Me.rdbSuppltDate.Text = "Supply Date"
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Controls.Add(Me.MyRadioButton3)
        Me.RadGroupBox1.Controls.Add(Me.rdbMilk)
        Me.RadGroupBox1.Controls.Add(Me.rdbProduct)
        Me.RadGroupBox1.HeaderText = "Milk Type"
        Me.RadGroupBox1.Location = New System.Drawing.Point(285, 82)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox1.Size = New System.Drawing.Size(176, 40)
        Me.RadGroupBox1.TabIndex = 1510
        Me.RadGroupBox1.Text = "Milk Type"
        '
        'MyRadioButton3
        '
        Me.MyRadioButton3.Location = New System.Drawing.Point(125, 11)
        Me.MyRadioButton3.MyLinkLable1 = Nothing
        Me.MyRadioButton3.MyLinkLable2 = Nothing
        Me.MyRadioButton3.Name = "MyRadioButton3"
        Me.MyRadioButton3.Size = New System.Drawing.Size(44, 18)
        Me.MyRadioButton3.TabIndex = 397
        Me.MyRadioButton3.TabStop = False
        Me.MyRadioButton3.Text = "Both"
        '
        'rdbMilk
        '
        Me.rdbMilk.Location = New System.Drawing.Point(6, 11)
        Me.rdbMilk.MyLinkLable1 = Nothing
        Me.rdbMilk.MyLinkLable2 = Nothing
        Me.rdbMilk.Name = "rdbMilk"
        Me.rdbMilk.Size = New System.Drawing.Size(41, 18)
        Me.rdbMilk.TabIndex = 396
        Me.rdbMilk.TabStop = False
        Me.rdbMilk.Text = "Milk"
        '
        'rdbProduct
        '
        Me.rdbProduct.Location = New System.Drawing.Point(61, 11)
        Me.rdbProduct.MyLinkLable1 = Nothing
        Me.rdbProduct.MyLinkLable2 = Nothing
        Me.rdbProduct.Name = "rdbProduct"
        Me.rdbProduct.Size = New System.Drawing.Size(59, 18)
        Me.rdbProduct.TabIndex = 391
        Me.rdbProduct.TabStop = False
        Me.rdbProduct.Text = "Product"
        '
        'RadGroupBox5
        '
        Me.RadGroupBox5.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox5.Controls.Add(Me.rdbInvBoth)
        Me.RadGroupBox5.Controls.Add(Me.rbtnNonTaxable)
        Me.RadGroupBox5.Controls.Add(Me.rbtnTaxable)
        Me.RadGroupBox5.HeaderText = "Invoice Type"
        Me.RadGroupBox5.Location = New System.Drawing.Point(285, 33)
        Me.RadGroupBox5.Name = "RadGroupBox5"
        Me.RadGroupBox5.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox5.Size = New System.Drawing.Size(222, 40)
        Me.RadGroupBox5.TabIndex = 1509
        Me.RadGroupBox5.Text = "Invoice Type"
        '
        'rdbInvBoth
        '
        Me.rdbInvBoth.Location = New System.Drawing.Point(158, 11)
        Me.rdbInvBoth.MyLinkLable1 = Nothing
        Me.rdbInvBoth.MyLinkLable2 = Nothing
        Me.rdbInvBoth.Name = "rdbInvBoth"
        Me.rdbInvBoth.Size = New System.Drawing.Size(44, 18)
        Me.rdbInvBoth.TabIndex = 398
        Me.rdbInvBoth.TabStop = False
        Me.rdbInvBoth.Text = "Both"
        '
        'rbtnNonTaxable
        '
        Me.rbtnNonTaxable.Location = New System.Drawing.Point(6, 11)
        Me.rbtnNonTaxable.MyLinkLable1 = Nothing
        Me.rbtnNonTaxable.MyLinkLable2 = Nothing
        Me.rbtnNonTaxable.Name = "rbtnNonTaxable"
        Me.rbtnNonTaxable.Size = New System.Drawing.Size(84, 18)
        Me.rbtnNonTaxable.TabIndex = 396
        Me.rbtnNonTaxable.TabStop = False
        Me.rbtnNonTaxable.Text = "Non-Taxable"
        '
        'rbtnTaxable
        '
        Me.rbtnTaxable.Location = New System.Drawing.Point(94, 11)
        Me.rbtnTaxable.MyLinkLable1 = Nothing
        Me.rbtnTaxable.MyLinkLable2 = Nothing
        Me.rbtnTaxable.Name = "rbtnTaxable"
        Me.rbtnTaxable.Size = New System.Drawing.Size(58, 18)
        Me.rbtnTaxable.TabIndex = 391
        Me.rbtnTaxable.TabStop = False
        Me.rbtnTaxable.Text = "Taxable"
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(12, 81)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(55, 18)
        Me.MyLabel13.TabIndex = 336
        Me.MyLabel13.Text = "Customer"
        '
        'txtMultiCustomer
        '
        Me.txtMultiCustomer.arrDispalyMember = Nothing
        Me.txtMultiCustomer.arrValueMember = Nothing
        Me.txtMultiCustomer.Location = New System.Drawing.Point(73, 82)
        Me.txtMultiCustomer.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMultiCustomer.MyLinkLable1 = Me.MyLabel13
        Me.txtMultiCustomer.MyLinkLable2 = Nothing
        Me.txtMultiCustomer.MyNullText = "All"
        Me.txtMultiCustomer.Name = "txtMultiCustomer"
        Me.txtMultiCustomer.Size = New System.Drawing.Size(194, 19)
        Me.txtMultiCustomer.TabIndex = 335
        '
        'RadGroupBox4
        '
        Me.RadGroupBox4.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox4.Controls.Add(Me.MyLabel5)
        Me.RadGroupBox4.Controls.Add(Me.MyLabel7)
        Me.RadGroupBox4.Controls.Add(Me.ToDate)
        Me.RadGroupBox4.Controls.Add(Me.fromdate)
        Me.RadGroupBox4.HeaderText = "Date Range"
        Me.RadGroupBox4.Location = New System.Drawing.Point(12, 33)
        Me.RadGroupBox4.Name = "RadGroupBox4"
        Me.RadGroupBox4.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox4.Size = New System.Drawing.Size(255, 42)
        Me.RadGroupBox4.TabIndex = 55
        Me.RadGroupBox4.Text = "Date Range"
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Location = New System.Drawing.Point(135, 16)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(19, 18)
        Me.MyLabel5.TabIndex = 3
        Me.MyLabel5.Text = "To"
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Location = New System.Drawing.Point(5, 16)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(32, 18)
        Me.MyLabel7.TabIndex = 2
        Me.MyLabel7.Text = "From"
        '
        'ToDate
        '
        Me.ToDate.CustomFormat = "dd/MM/yyyy"
        Me.ToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.ToDate.Location = New System.Drawing.Point(162, 15)
        Me.ToDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.ToDate.Name = "ToDate"
        Me.ToDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.ToDate.Size = New System.Drawing.Size(83, 20)
        Me.ToDate.TabIndex = 1
        Me.ToDate.TabStop = False
        Me.ToDate.Text = "24/10/2011"
        Me.ToDate.Value = New Date(2011, 10, 24, 2, 29, 0, 265)
        '
        'fromdate
        '
        Me.fromdate.CustomFormat = "dd/MM/yyyy"
        Me.fromdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.fromdate.Location = New System.Drawing.Point(44, 15)
        Me.fromdate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.fromdate.Name = "fromdate"
        Me.fromdate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.fromdate.Size = New System.Drawing.Size(85, 20)
        Me.fromdate.TabIndex = 0
        Me.fromdate.TabStop = False
        Me.fromdate.Text = "24/10/2011"
        Me.fromdate.Value = New Date(2011, 10, 24, 2, 29, 0, 265)
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(717, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(71, 22)
        Me.btnClose.TabIndex = 156
        Me.btnClose.Text = "Close"
        '
        'btnReset
        '
        Me.btnReset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReset.Location = New System.Drawing.Point(89, 8)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(71, 22)
        Me.btnReset.TabIndex = 155
        Me.btnReset.Text = "Reset"
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(12, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(71, 22)
        Me.btnPrint.TabIndex = 154
        Me.btnPrint.Text = "Print"
        '
        'FrmperdayDetailRpt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "FrmperdayDetailRpt"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "FrmperdayDetailRpt"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        Me.RadGroupBox2.PerformLayout()
        CType(Me.rdbDocDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbSuppltDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.MyRadioButton3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbMilk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbProduct, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox5.ResumeLayout(False)
        Me.RadGroupBox5.PerformLayout()
        CType(Me.rdbInvBoth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnNonTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox4.ResumeLayout(False)
        Me.RadGroupBox4.PerformLayout()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ToDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fromdate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents RadGroupBox4 As RadGroupBox
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents ToDate As RadDateTimePicker
    Friend WithEvents fromdate As RadDateTimePicker
    Friend WithEvents btnPrint As RadButton
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents txtMultiCustomer As common.UserControls.txtMultiSelectFinder
    Friend WithEvents RadGroupBox5 As RadGroupBox
    Friend WithEvents rbtnNonTaxable As common.Controls.MyRadioButton
    Friend WithEvents rbtnTaxable As common.Controls.MyRadioButton
    Friend WithEvents RadGroupBox1 As RadGroupBox
    Friend WithEvents MyRadioButton3 As common.Controls.MyRadioButton
    Friend WithEvents rdbMilk As common.Controls.MyRadioButton
    Friend WithEvents rdbProduct As common.Controls.MyRadioButton
    Friend WithEvents rdbInvBoth As common.Controls.MyRadioButton
    Friend WithEvents RadGroupBox2 As RadGroupBox
    Friend WithEvents rdbDocDate As common.Controls.MyRadioButton
    Friend WithEvents rdbSuppltDate As common.Controls.MyRadioButton
    Friend WithEvents btnClose As RadButton
    Friend WithEvents btnReset As RadButton
End Class
