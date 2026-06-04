<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BMC_Transporter_Bill
    Inherits FrmMainTranScreen
    'Inherits System.Windows.Forms.Form

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
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.RadLabel6 = New common.Controls.MyLabel()
        Me.RadLabel3 = New common.Controls.MyLabel()
        Me.txtRemarks = New common.Controls.MyTextBox()
        Me.txtComment = New common.Controls.MyTextBox()
        Me.chkPrivate = New System.Windows.Forms.CheckBox()
        Me.UsLock1 = New common.usLock()
        Me.RadGroupBox3 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtFatShortageNMG = New common.Controls.MyTextBox()
        Me.MyLabel28 = New common.Controls.MyLabel()
        Me.TXTSNFShortageNMG = New common.Controls.MyTextBox()
        Me.MyLabel29 = New common.Controls.MyLabel()
        Me.TxtFatRateNMG = New common.Controls.MyTextBox()
        Me.MyLabel30 = New common.Controls.MyLabel()
        Me.TxtSnfRateNMG = New common.Controls.MyTextBox()
        Me.MyLabel35 = New common.Controls.MyLabel()
        Me.txtFatShortage = New common.Controls.MyTextBox()
        Me.btnGo = New Telerik.WinControls.UI.RadButton()
        Me.lblFatShortage = New common.Controls.MyLabel()
        Me.TxtSnfShortage = New common.Controls.MyTextBox()
        Me.TxtTankerprorata = New common.Controls.MyTextBox()
        Me.TxtIceCharge = New common.Controls.MyTextBox()
        Me.MyLabel11 = New common.Controls.MyLabel()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.TxtKMRate = New common.Controls.MyTextBox()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.MyLabel8 = New common.Controls.MyLabel()
        Me.TxtBarelCap = New common.Controls.MyTextBox()
        Me.TxtTollTax = New common.Controls.MyTextBox()
        Me.TxtFatRate = New common.Controls.MyTextBox()
        Me.TxtDieselMinus = New common.Controls.MyTextBox()
        Me.MyLabel12 = New common.Controls.MyLabel()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.MyLabel3 = New common.Controls.MyLabel()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.TxtSnfRate = New common.Controls.MyTextBox()
        Me.txtDieselplus = New common.Controls.MyTextBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.TxtRateprorata = New common.Controls.MyTextBox()
        Me.TxtTDR = New common.Controls.MyTextBox()
        Me.MyLabel9 = New common.Controls.MyLabel()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtToDate = New common.Controls.MyDateTimePicker()
        Me.txtFromDate = New common.Controls.MyDateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkDCS = New Telerik.WinControls.UI.RadCheckBox()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.txtTransporter = New common.Controls.MyTextBox()
        Me.lblTransporter = New common.Controls.MyLabel()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.RadLabel4 = New common.Controls.MyLabel()
        Me.lblTankerDesc = New common.Controls.MyLabel()
        Me.lblAbandonmentNo = New common.Controls.MyLabel()
        Me.lblTankerNo = New common.Controls.MyLabel()
        Me.txtTankerNo = New common.UserControls.txtFinder()
        Me.RadLabel1 = New common.Controls.MyLabel()
        Me.txtDocNo = New common.UserControls.txtNavigator()
        Me.RadPageViewPage2 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.NACCOEPF_PER = New common.Controls.MyLabel()
        Me.NEMPEPF_PER = New common.Controls.MyLabel()
        Me.NCOESI_PER = New common.Controls.MyLabel()
        Me.NEMPESI_PER = New common.Controls.MyLabel()
        Me.NTotalPF = New common.Controls.MyLabel()
        Me.NTotalESI = New common.Controls.MyLabel()
        Me.NESIEmployer = New common.Controls.MyLabel()
        Me.NESIEmployee = New common.Controls.MyLabel()
        Me.NEDLICharge = New common.Controls.MyLabel()
        Me.NAdmincharge = New common.Controls.MyLabel()
        Me.NemployerPF = New common.Controls.MyLabel()
        Me.NempPF = New common.Controls.MyLabel()
        Me.NLabourCharge = New common.Controls.MyLabel()
        Me.NDailyWage = New common.Controls.MyLabel()
        Me.MyLabel70 = New common.Controls.MyLabel()
        Me.MyLabel69 = New common.Controls.MyLabel()
        Me.MyLabel68 = New common.Controls.MyLabel()
        Me.MyLabel67 = New common.Controls.MyLabel()
        Me.MyLabel66 = New common.Controls.MyLabel()
        Me.MyLabel65 = New common.Controls.MyLabel()
        Me.MyLabel61 = New common.Controls.MyLabel()
        Me.MyLabel60 = New common.Controls.MyLabel()
        Me.MyLabel59 = New common.Controls.MyLabel()
        Me.MyLabel58 = New common.Controls.MyLabel()
        Me.MyLabel57 = New common.Controls.MyLabel()
        Me.MyLabel56 = New common.Controls.MyLabel()
        Me.MyLabel55 = New common.Controls.MyLabel()
        Me.MyLabel54 = New common.Controls.MyLabel()
        Me.MyLabel53 = New common.Controls.MyLabel()
        Me.MyLabel52 = New common.Controls.MyLabel()
        Me.MyLabel51 = New common.Controls.MyLabel()
        Me.MyLabel50 = New common.Controls.MyLabel()
        Me.MyLabel49 = New common.Controls.MyLabel()
        Me.MyLabel48 = New common.Controls.MyLabel()
        Me.MyLabel37 = New common.Controls.MyLabel()
        Me.MyLabel36 = New common.Controls.MyLabel()
        Me.RadPageViewPage4 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.TotalESIAmt = New common.Controls.MyLabel()
        Me.TotalEPFAmt = New common.Controls.MyLabel()
        Me.MyLabel64 = New common.Controls.MyLabel()
        Me.MyLabel38 = New common.Controls.MyLabel()
        Me.TxtTotalDeduction = New common.Controls.MyTextBox()
        Me.TxtTotalAddition = New common.Controls.MyTextBox()
        Me.MyLabel39 = New common.Controls.MyLabel()
        Me.MyLabel63 = New common.Controls.MyLabel()
        Me.MyLabel47 = New common.Controls.MyLabel()
        Me.MyLabel46 = New common.Controls.MyLabel()
        Me.MyLabel45 = New common.Controls.MyLabel()
        Me.MyLabel44 = New common.Controls.MyLabel()
        Me.MyLabel43 = New common.Controls.MyLabel()
        Me.MyLabel42 = New common.Controls.MyLabel()
        Me.MyLabel41 = New common.Controls.MyLabel()
        Me.MyLabel40 = New common.Controls.MyLabel()
        Me.TxtTotalIceCharge = New common.Controls.MyLabel()
        Me.TxtTotalAmount = New common.Controls.MyLabel()
        Me.TxtTotalFatSnfShortage = New common.Controls.MyLabel()
        Me.TxtGrossAmount = New common.Controls.MyLabel()
        Me.TxtTotalTollTax = New common.Controls.MyLabel()
        Me.MyLabel34 = New common.Controls.MyLabel()
        Me.MyLabel33 = New common.Controls.MyLabel()
        Me.MyLabel32 = New common.Controls.MyLabel()
        Me.MyLabel31 = New common.Controls.MyLabel()
        Me.TxtBMCTotal = New common.Controls.MyLabel()
        Me.TxtBMCDiesel = New common.Controls.MyLabel()
        Me.TxtBMCProrataamt = New common.Controls.MyLabel()
        Me.MyLabel27 = New common.Controls.MyLabel()
        Me.MyLabel26 = New common.Controls.MyLabel()
        Me.MyLabel62 = New common.Controls.MyLabel()
        Me.MyLabel25 = New common.Controls.MyLabel()
        Me.MyLabel24 = New common.Controls.MyLabel()
        Me.MyLabel23 = New common.Controls.MyLabel()
        Me.MyLabel22 = New common.Controls.MyLabel()
        Me.MyLabel21 = New common.Controls.MyLabel()
        Me.MyLabel20 = New common.Controls.MyLabel()
        Me.MyLabel19 = New common.Controls.MyLabel()
        Me.MyLabel18 = New common.Controls.MyLabel()
        Me.MyLabel17 = New common.Controls.MyLabel()
        Me.MyLabel16 = New common.Controls.MyLabel()
        Me.MyLabel15 = New common.Controls.MyLabel()
        Me.MyLabel14 = New common.Controls.MyLabel()
        Me.lblInvoiceDiscAmt = New common.Controls.MyLabel()
        Me.MyLabel6 = New common.Controls.MyLabel()
        Me.RadLabel32 = New common.Controls.MyLabel()
        Me.lblAddCharges1 = New common.Controls.MyLabel()
        Me.RadLabel9 = New common.Controls.MyLabel()
        Me.RadLabel27 = New common.Controls.MyLabel()
        Me.lblTotRAmt = New common.Controls.MyLabel()
        Me.RadLabel25 = New common.Controls.MyLabel()
        Me.lblTaxAmt = New common.Controls.MyLabel()
        Me.lblAmtAfterDiscount = New common.Controls.MyLabel()
        Me.lblDiscountAmt = New common.Controls.MyLabel()
        Me.lblAmtWithDiscount = New common.Controls.MyLabel()
        Me.RadLabel22 = New common.Controls.MyLabel()
        Me.RadLabel19 = New common.Controls.MyLabel()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.RadSplitButton1 = New Telerik.WinControls.UI.RadSplitButton()
        Me.btnImport = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnExport = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.btnReverse = New Telerik.WinControls.UI.RadButton()
        Me.btnLayout = New Telerik.WinControls.UI.RadSplitButton()
        Me.rmiImport = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiExport = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnDelete = New Telerik.WinControls.UI.RadButton()
        Me.btnPost = New Telerik.WinControls.UI.RadButton()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.btnAddNew = New Telerik.WinControls.UI.RadButton()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRemarks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtComment, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox3.SuspendLayout()
        CType(Me.txtFatShortageNMG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel28, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TXTSNFShortageNMG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtFatRateNMG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel30, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtSnfRateNMG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel35, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFatShortage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFatShortage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtSnfShortage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTankerprorata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtIceCharge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtKMRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtBarelCap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTollTax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtFatRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtDieselMinus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtSnfRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDieselplus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtRateprorata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTDR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDCS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTransporter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTransporter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTankerDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAbandonmentNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTankerNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage2.SuspendLayout()
        CType(Me.NACCOEPF_PER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NEMPEPF_PER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NCOESI_PER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NEMPESI_PER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NTotalPF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NTotalESI, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NESIEmployer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NESIEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NEDLICharge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NAdmincharge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NemployerPF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NempPF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NLabourCharge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NDailyWage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel70, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel69, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel68, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel67, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel66, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel65, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel61, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel60, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel59, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel58, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel57, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel56, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel55, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel54, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel53, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel52, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel51, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel50, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel49, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel48, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel37, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel36, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage4.SuspendLayout()
        CType(Me.TotalESIAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TotalEPFAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel64, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel38, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalDeduction, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalAddition, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel39, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel63, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel47, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel46, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel45, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel44, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel43, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel42, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel41, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel40, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalIceCharge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalFatSnfShortage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtGrossAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtTotalTollTax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel34, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel33, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtBMCTotal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtBMCDiesel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtBMCProrataamt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel62, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel24, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel20, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblInvoiceDiscAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotRAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAmtAfterDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDiscountAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAmtWithDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnLayout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadPageView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadSplitButton1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReverse)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnLayout)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 405
        Me.SplitContainer1.TabIndex = 0
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage2)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage4)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage1
        Me.RadPageView1.Size = New System.Drawing.Size(800, 405)
        Me.RadPageView1.TabIndex = 1
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).ItemAlignment = Telerik.WinControls.UI.StripViewItemAlignment.Near
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripAlignment = Telerik.WinControls.UI.StripViewAlignment.Top
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel6)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel3)
        Me.RadPageViewPage1.Controls.Add(Me.txtRemarks)
        Me.RadPageViewPage1.Controls.Add(Me.txtComment)
        Me.RadPageViewPage1.Controls.Add(Me.chkPrivate)
        Me.RadPageViewPage1.Controls.Add(Me.UsLock1)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox3)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox1)
        Me.RadPageViewPage1.Controls.Add(Me.chkDCS)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox2)
        Me.RadPageViewPage1.Controls.Add(Me.txtTransporter)
        Me.RadPageViewPage1.Controls.Add(Me.lblTransporter)
        Me.RadPageViewPage1.Controls.Add(Me.txtDate)
        Me.RadPageViewPage1.Controls.Add(Me.lblTankerDesc)
        Me.RadPageViewPage1.Controls.Add(Me.lblAbandonmentNo)
        Me.RadPageViewPage1.Controls.Add(Me.lblTankerNo)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel4)
        Me.RadPageViewPage1.Controls.Add(Me.txtTankerNo)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.txtDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.btnAddNew)
        Me.RadPageViewPage1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(89.0!, 22.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 31)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(779, 363)
        Me.RadPageViewPage1.Text = "Transporter Bill"
        '
        'RadLabel6
        '
        Me.RadLabel6.FieldName = Nothing
        Me.RadLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel6.Location = New System.Drawing.Point(433, 69)
        Me.RadLabel6.Name = "RadLabel6"
        Me.RadLabel6.Size = New System.Drawing.Size(51, 16)
        Me.RadLabel6.TabIndex = 1558
        Me.RadLabel6.Text = "Remarks"
        '
        'RadLabel3
        '
        Me.RadLabel3.FieldName = Nothing
        Me.RadLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel3.Location = New System.Drawing.Point(3, 71)
        Me.RadLabel3.Name = "RadLabel3"
        Me.RadLabel3.Size = New System.Drawing.Size(55, 16)
        Me.RadLabel3.TabIndex = 1559
        Me.RadLabel3.Text = "Comment"
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
        Me.txtRemarks.Location = New System.Drawing.Point(504, 68)
        Me.txtRemarks.MaxLength = 200
        Me.txtRemarks.MendatroryField = False
        Me.txtRemarks.MyLinkLable1 = Me.RadLabel6
        Me.txtRemarks.MyLinkLable2 = Nothing
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ReferenceFieldDesc = Nothing
        Me.txtRemarks.ReferenceFieldName = Nothing
        Me.txtRemarks.ReferenceTableName = Nothing
        Me.txtRemarks.Size = New System.Drawing.Size(262, 18)
        Me.txtRemarks.TabIndex = 1557
        '
        'txtComment
        '
        Me.txtComment.CalculationExpression = Nothing
        Me.txtComment.FieldCode = Nothing
        Me.txtComment.FieldDesc = Nothing
        Me.txtComment.FieldMaxLength = 0
        Me.txtComment.FieldName = Nothing
        Me.txtComment.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComment.isCalculatedField = False
        Me.txtComment.IsSourceFromTable = False
        Me.txtComment.IsSourceFromValueList = False
        Me.txtComment.IsUnique = False
        Me.txtComment.Location = New System.Drawing.Point(78, 68)
        Me.txtComment.MaxLength = 200
        Me.txtComment.MendatroryField = False
        Me.txtComment.MyLinkLable1 = Me.RadLabel3
        Me.txtComment.MyLinkLable2 = Nothing
        Me.txtComment.Name = "txtComment"
        Me.txtComment.ReferenceFieldDesc = Nothing
        Me.txtComment.ReferenceFieldName = Nothing
        Me.txtComment.ReferenceTableName = Nothing
        Me.txtComment.Size = New System.Drawing.Size(349, 18)
        Me.txtComment.TabIndex = 1556
        '
        'chkPrivate
        '
        Me.chkPrivate.AutoSize = True
        Me.chkPrivate.Location = New System.Drawing.Point(704, 3)
        Me.chkPrivate.Name = "chkPrivate"
        Me.chkPrivate.Size = New System.Drawing.Size(60, 18)
        Me.chkPrivate.TabIndex = 1507
        Me.chkPrivate.Text = "Manual"
        Me.chkPrivate.UseVisualStyleBackColor = True
        '
        'UsLock1
        '
        Me.UsLock1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.UsLock1.Location = New System.Drawing.Point(579, 1)
        Me.UsLock1.Margin = New System.Windows.Forms.Padding(4)
        Me.UsLock1.MyFont = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsLock1.Name = "UsLock1"
        Me.UsLock1.Size = New System.Drawing.Size(118, 18)
        Me.UsLock1.Status = common.ERPTransactionStatus.Pending
        Me.UsLock1.TabIndex = 1555
        '
        'RadGroupBox3
        '
        Me.RadGroupBox3.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox3.Controls.Add(Me.txtFatShortageNMG)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel28)
        Me.RadGroupBox3.Controls.Add(Me.TXTSNFShortageNMG)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel29)
        Me.RadGroupBox3.Controls.Add(Me.TxtFatRateNMG)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel30)
        Me.RadGroupBox3.Controls.Add(Me.TxtSnfRateNMG)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel35)
        Me.RadGroupBox3.Controls.Add(Me.txtFatShortage)
        Me.RadGroupBox3.Controls.Add(Me.btnGo)
        Me.RadGroupBox3.Controls.Add(Me.lblFatShortage)
        Me.RadGroupBox3.Controls.Add(Me.TxtSnfShortage)
        Me.RadGroupBox3.Controls.Add(Me.TxtTankerprorata)
        Me.RadGroupBox3.Controls.Add(Me.TxtIceCharge)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel11)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel13)
        Me.RadGroupBox3.Controls.Add(Me.TxtKMRate)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel2)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel8)
        Me.RadGroupBox3.Controls.Add(Me.TxtBarelCap)
        Me.RadGroupBox3.Controls.Add(Me.TxtTollTax)
        Me.RadGroupBox3.Controls.Add(Me.TxtFatRate)
        Me.RadGroupBox3.Controls.Add(Me.TxtDieselMinus)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel12)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel7)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel3)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel5)
        Me.RadGroupBox3.Controls.Add(Me.TxtSnfRate)
        Me.RadGroupBox3.Controls.Add(Me.txtDieselplus)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel1)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel4)
        Me.RadGroupBox3.Controls.Add(Me.TxtRateprorata)
        Me.RadGroupBox3.Controls.Add(Me.TxtTDR)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel9)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel10)
        Me.RadGroupBox3.HeaderText = ""
        Me.RadGroupBox3.Location = New System.Drawing.Point(3, 91)
        Me.RadGroupBox3.Name = "RadGroupBox3"
        Me.RadGroupBox3.Size = New System.Drawing.Size(763, 126)
        Me.RadGroupBox3.TabIndex = 1554
        '
        'txtFatShortageNMG
        '
        Me.txtFatShortageNMG.CalculationExpression = Nothing
        Me.txtFatShortageNMG.FieldCode = Nothing
        Me.txtFatShortageNMG.FieldDesc = Nothing
        Me.txtFatShortageNMG.FieldMaxLength = 0
        Me.txtFatShortageNMG.FieldName = Nothing
        Me.txtFatShortageNMG.isCalculatedField = False
        Me.txtFatShortageNMG.IsSourceFromTable = False
        Me.txtFatShortageNMG.IsSourceFromValueList = False
        Me.txtFatShortageNMG.IsUnique = False
        Me.txtFatShortageNMG.Location = New System.Drawing.Point(701, 4)
        Me.txtFatShortageNMG.MendatroryField = False
        Me.txtFatShortageNMG.MyLinkLable1 = Nothing
        Me.txtFatShortageNMG.MyLinkLable2 = Nothing
        Me.txtFatShortageNMG.Name = "txtFatShortageNMG"
        Me.txtFatShortageNMG.ReferenceFieldDesc = Nothing
        Me.txtFatShortageNMG.ReferenceFieldName = Nothing
        Me.txtFatShortageNMG.ReferenceTableName = Nothing
        Me.txtFatShortageNMG.Size = New System.Drawing.Size(60, 20)
        Me.txtFatShortageNMG.TabIndex = 1558
        Me.txtFatShortageNMG.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel28
        '
        Me.MyLabel28.FieldName = Nothing
        Me.MyLabel28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel28.Location = New System.Drawing.Point(591, 7)
        Me.MyLabel28.Name = "MyLabel28"
        Me.MyLabel28.Size = New System.Drawing.Size(106, 16)
        Me.MyLabel28.TabIndex = 1557
        Me.MyLabel28.Text = "FAT Shortage NMG"
        '
        'TXTSNFShortageNMG
        '
        Me.TXTSNFShortageNMG.CalculationExpression = Nothing
        Me.TXTSNFShortageNMG.FieldCode = Nothing
        Me.TXTSNFShortageNMG.FieldDesc = Nothing
        Me.TXTSNFShortageNMG.FieldMaxLength = 0
        Me.TXTSNFShortageNMG.FieldName = Nothing
        Me.TXTSNFShortageNMG.isCalculatedField = False
        Me.TXTSNFShortageNMG.IsSourceFromTable = False
        Me.TXTSNFShortageNMG.IsSourceFromValueList = False
        Me.TXTSNFShortageNMG.IsUnique = False
        Me.TXTSNFShortageNMG.Location = New System.Drawing.Point(702, 30)
        Me.TXTSNFShortageNMG.MendatroryField = False
        Me.TXTSNFShortageNMG.MyLinkLable1 = Nothing
        Me.TXTSNFShortageNMG.MyLinkLable2 = Nothing
        Me.TXTSNFShortageNMG.Name = "TXTSNFShortageNMG"
        Me.TXTSNFShortageNMG.ReferenceFieldDesc = Nothing
        Me.TXTSNFShortageNMG.ReferenceFieldName = Nothing
        Me.TXTSNFShortageNMG.ReferenceTableName = Nothing
        Me.TXTSNFShortageNMG.Size = New System.Drawing.Size(59, 20)
        Me.TXTSNFShortageNMG.TabIndex = 1559
        Me.TXTSNFShortageNMG.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel29
        '
        Me.MyLabel29.FieldName = Nothing
        Me.MyLabel29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel29.Location = New System.Drawing.Point(591, 30)
        Me.MyLabel29.Name = "MyLabel29"
        Me.MyLabel29.Size = New System.Drawing.Size(108, 16)
        Me.MyLabel29.TabIndex = 1554
        Me.MyLabel29.Text = "SNF Shortage NMG"
        '
        'TxtFatRateNMG
        '
        Me.TxtFatRateNMG.CalculationExpression = Nothing
        Me.TxtFatRateNMG.FieldCode = Nothing
        Me.TxtFatRateNMG.FieldDesc = Nothing
        Me.TxtFatRateNMG.FieldMaxLength = 0
        Me.TxtFatRateNMG.FieldName = Nothing
        Me.TxtFatRateNMG.isCalculatedField = False
        Me.TxtFatRateNMG.IsSourceFromTable = False
        Me.TxtFatRateNMG.IsSourceFromValueList = False
        Me.TxtFatRateNMG.IsUnique = False
        Me.TxtFatRateNMG.Location = New System.Drawing.Point(703, 56)
        Me.TxtFatRateNMG.MendatroryField = False
        Me.TxtFatRateNMG.MyLinkLable1 = Nothing
        Me.TxtFatRateNMG.MyLinkLable2 = Nothing
        Me.TxtFatRateNMG.Name = "TxtFatRateNMG"
        Me.TxtFatRateNMG.ReferenceFieldDesc = Nothing
        Me.TxtFatRateNMG.ReferenceFieldName = Nothing
        Me.TxtFatRateNMG.ReferenceTableName = Nothing
        Me.TxtFatRateNMG.Size = New System.Drawing.Size(58, 20)
        Me.TxtFatRateNMG.TabIndex = 1560
        Me.TxtFatRateNMG.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel30
        '
        Me.MyLabel30.FieldName = Nothing
        Me.MyLabel30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel30.Location = New System.Drawing.Point(615, 56)
        Me.MyLabel30.Name = "MyLabel30"
        Me.MyLabel30.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel30.TabIndex = 1555
        Me.MyLabel30.Text = "FAT Rate NMG"
        '
        'TxtSnfRateNMG
        '
        Me.TxtSnfRateNMG.CalculationExpression = Nothing
        Me.TxtSnfRateNMG.FieldCode = Nothing
        Me.TxtSnfRateNMG.FieldDesc = Nothing
        Me.TxtSnfRateNMG.FieldMaxLength = 0
        Me.TxtSnfRateNMG.FieldName = Nothing
        Me.TxtSnfRateNMG.isCalculatedField = False
        Me.TxtSnfRateNMG.IsSourceFromTable = False
        Me.TxtSnfRateNMG.IsSourceFromValueList = False
        Me.TxtSnfRateNMG.IsUnique = False
        Me.TxtSnfRateNMG.Location = New System.Drawing.Point(704, 82)
        Me.TxtSnfRateNMG.MendatroryField = False
        Me.TxtSnfRateNMG.MyLinkLable1 = Nothing
        Me.TxtSnfRateNMG.MyLinkLable2 = Nothing
        Me.TxtSnfRateNMG.Name = "TxtSnfRateNMG"
        Me.TxtSnfRateNMG.ReferenceFieldDesc = Nothing
        Me.TxtSnfRateNMG.ReferenceFieldName = Nothing
        Me.TxtSnfRateNMG.ReferenceTableName = Nothing
        Me.TxtSnfRateNMG.Size = New System.Drawing.Size(57, 20)
        Me.TxtSnfRateNMG.TabIndex = 1561
        Me.TxtSnfRateNMG.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel35
        '
        Me.MyLabel35.FieldName = Nothing
        Me.MyLabel35.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel35.Location = New System.Drawing.Point(617, 82)
        Me.MyLabel35.Name = "MyLabel35"
        Me.MyLabel35.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel35.TabIndex = 1556
        Me.MyLabel35.Text = "SNF RateNMG"
        '
        'txtFatShortage
        '
        Me.txtFatShortage.CalculationExpression = Nothing
        Me.txtFatShortage.FieldCode = Nothing
        Me.txtFatShortage.FieldDesc = Nothing
        Me.txtFatShortage.FieldMaxLength = 0
        Me.txtFatShortage.FieldName = Nothing
        Me.txtFatShortage.isCalculatedField = False
        Me.txtFatShortage.IsSourceFromTable = False
        Me.txtFatShortage.IsSourceFromValueList = False
        Me.txtFatShortage.IsUnique = False
        Me.txtFatShortage.Location = New System.Drawing.Point(510, 5)
        Me.txtFatShortage.MendatroryField = False
        Me.txtFatShortage.MyLinkLable1 = Nothing
        Me.txtFatShortage.MyLinkLable2 = Nothing
        Me.txtFatShortage.Name = "txtFatShortage"
        Me.txtFatShortage.ReadOnly = True
        Me.txtFatShortage.ReferenceFieldDesc = Nothing
        Me.txtFatShortage.ReferenceFieldName = Nothing
        Me.txtFatShortage.ReferenceTableName = Nothing
        Me.txtFatShortage.Size = New System.Drawing.Size(82, 20)
        Me.txtFatShortage.TabIndex = 1534
        Me.txtFatShortage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnGo
        '
        Me.btnGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGo.Location = New System.Drawing.Point(352, 92)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(57, 17)
        Me.btnGo.TabIndex = 1553
        Me.btnGo.Text = ">>"
        '
        'lblFatShortage
        '
        Me.lblFatShortage.FieldName = Nothing
        Me.lblFatShortage.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFatShortage.Location = New System.Drawing.Point(403, 8)
        Me.lblFatShortage.Name = "lblFatShortage"
        Me.lblFatShortage.Size = New System.Drawing.Size(106, 16)
        Me.lblFatShortage.TabIndex = 1533
        Me.lblFatShortage.Text = "FAT Shortage RMG"
        '
        'TxtSnfShortage
        '
        Me.TxtSnfShortage.CalculationExpression = Nothing
        Me.TxtSnfShortage.FieldCode = Nothing
        Me.TxtSnfShortage.FieldDesc = Nothing
        Me.TxtSnfShortage.FieldMaxLength = 0
        Me.TxtSnfShortage.FieldName = Nothing
        Me.TxtSnfShortage.isCalculatedField = False
        Me.TxtSnfShortage.IsSourceFromTable = False
        Me.TxtSnfShortage.IsSourceFromValueList = False
        Me.TxtSnfShortage.IsUnique = False
        Me.TxtSnfShortage.Location = New System.Drawing.Point(510, 31)
        Me.TxtSnfShortage.MendatroryField = False
        Me.TxtSnfShortage.MyLinkLable1 = Nothing
        Me.TxtSnfShortage.MyLinkLable2 = Nothing
        Me.TxtSnfShortage.Name = "TxtSnfShortage"
        Me.TxtSnfShortage.ReadOnly = True
        Me.TxtSnfShortage.ReferenceFieldDesc = Nothing
        Me.TxtSnfShortage.ReferenceFieldName = Nothing
        Me.TxtSnfShortage.ReferenceTableName = Nothing
        Me.TxtSnfShortage.Size = New System.Drawing.Size(82, 20)
        Me.TxtSnfShortage.TabIndex = 1535
        Me.TxtSnfShortage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTankerprorata
        '
        Me.TxtTankerprorata.CalculationExpression = Nothing
        Me.TxtTankerprorata.FieldCode = Nothing
        Me.TxtTankerprorata.FieldDesc = Nothing
        Me.TxtTankerprorata.FieldMaxLength = 0
        Me.TxtTankerprorata.FieldName = Nothing
        Me.TxtTankerprorata.isCalculatedField = False
        Me.TxtTankerprorata.IsSourceFromTable = False
        Me.TxtTankerprorata.IsSourceFromValueList = False
        Me.TxtTankerprorata.IsUnique = False
        Me.TxtTankerprorata.Location = New System.Drawing.Point(87, 92)
        Me.TxtTankerprorata.MendatroryField = False
        Me.TxtTankerprorata.MyLinkLable1 = Nothing
        Me.TxtTankerprorata.MyLinkLable2 = Nothing
        Me.TxtTankerprorata.Name = "TxtTankerprorata"
        Me.TxtTankerprorata.ReferenceFieldDesc = Nothing
        Me.TxtTankerprorata.ReferenceFieldName = Nothing
        Me.TxtTankerprorata.ReferenceTableName = Nothing
        Me.TxtTankerprorata.Size = New System.Drawing.Size(82, 20)
        Me.TxtTankerprorata.TabIndex = 1547
        Me.TxtTankerprorata.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtIceCharge
        '
        Me.TxtIceCharge.CalculationExpression = Nothing
        Me.TxtIceCharge.FieldCode = Nothing
        Me.TxtIceCharge.FieldDesc = Nothing
        Me.TxtIceCharge.FieldMaxLength = 0
        Me.TxtIceCharge.FieldName = Nothing
        Me.TxtIceCharge.isCalculatedField = False
        Me.TxtIceCharge.IsSourceFromTable = False
        Me.TxtIceCharge.IsSourceFromValueList = False
        Me.TxtIceCharge.IsUnique = False
        Me.TxtIceCharge.Location = New System.Drawing.Point(264, 28)
        Me.TxtIceCharge.MendatroryField = False
        Me.TxtIceCharge.MyLinkLable1 = Nothing
        Me.TxtIceCharge.MyLinkLable2 = Nothing
        Me.TxtIceCharge.Name = "TxtIceCharge"
        Me.TxtIceCharge.ReferenceFieldDesc = Nothing
        Me.TxtIceCharge.ReferenceFieldName = Nothing
        Me.TxtIceCharge.ReferenceTableName = Nothing
        Me.TxtIceCharge.Size = New System.Drawing.Size(82, 20)
        Me.TxtIceCharge.TabIndex = 1551
        Me.TxtIceCharge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel11
        '
        Me.MyLabel11.FieldName = Nothing
        Me.MyLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel11.Location = New System.Drawing.Point(9, 96)
        Me.MyLabel11.Name = "MyLabel11"
        Me.MyLabel11.Size = New System.Drawing.Size(61, 16)
        Me.MyLabel11.TabIndex = 1546
        Me.MyLabel11.Text = "PRORATA"
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(178, 30)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(62, 16)
        Me.MyLabel13.TabIndex = 1550
        Me.MyLabel13.Text = "Ice Charge"
        '
        'TxtKMRate
        '
        Me.TxtKMRate.CalculationExpression = Nothing
        Me.TxtKMRate.FieldCode = Nothing
        Me.TxtKMRate.FieldDesc = Nothing
        Me.TxtKMRate.FieldMaxLength = 0
        Me.TxtKMRate.FieldName = Nothing
        Me.TxtKMRate.isCalculatedField = False
        Me.TxtKMRate.IsSourceFromTable = False
        Me.TxtKMRate.IsSourceFromValueList = False
        Me.TxtKMRate.IsUnique = False
        Me.TxtKMRate.Location = New System.Drawing.Point(87, 70)
        Me.TxtKMRate.MendatroryField = False
        Me.TxtKMRate.MyLinkLable1 = Nothing
        Me.TxtKMRate.MyLinkLable2 = Nothing
        Me.TxtKMRate.Name = "TxtKMRate"
        Me.TxtKMRate.ReferenceFieldDesc = Nothing
        Me.TxtKMRate.ReferenceFieldName = Nothing
        Me.TxtKMRate.ReferenceTableName = Nothing
        Me.TxtKMRate.Size = New System.Drawing.Size(82, 20)
        Me.TxtKMRate.TabIndex = 1541
        Me.TxtKMRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(400, 31)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(108, 16)
        Me.MyLabel2.TabIndex = 1532
        Me.MyLabel2.Text = "SNF Shortage RMG"
        '
        'MyLabel8
        '
        Me.MyLabel8.FieldName = Nothing
        Me.MyLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(9, 74)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(50, 16)
        Me.MyLabel8.TabIndex = 1540
        Me.MyLabel8.Text = "KM Rate"
        '
        'TxtBarelCap
        '
        Me.TxtBarelCap.CalculationExpression = Nothing
        Me.TxtBarelCap.FieldCode = Nothing
        Me.TxtBarelCap.FieldDesc = Nothing
        Me.TxtBarelCap.FieldMaxLength = 0
        Me.TxtBarelCap.FieldName = Nothing
        Me.TxtBarelCap.isCalculatedField = False
        Me.TxtBarelCap.IsSourceFromTable = False
        Me.TxtBarelCap.IsSourceFromValueList = False
        Me.TxtBarelCap.IsUnique = False
        Me.TxtBarelCap.Location = New System.Drawing.Point(264, 54)
        Me.TxtBarelCap.MendatroryField = False
        Me.TxtBarelCap.MyLinkLable1 = Nothing
        Me.TxtBarelCap.MyLinkLable2 = Nothing
        Me.TxtBarelCap.Name = "TxtBarelCap"
        Me.TxtBarelCap.ReferenceFieldDesc = Nothing
        Me.TxtBarelCap.ReferenceFieldName = Nothing
        Me.TxtBarelCap.ReferenceTableName = Nothing
        Me.TxtBarelCap.Size = New System.Drawing.Size(82, 20)
        Me.TxtBarelCap.TabIndex = 1549
        Me.TxtBarelCap.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTollTax
        '
        Me.TxtTollTax.CalculationExpression = Nothing
        Me.TxtTollTax.FieldCode = Nothing
        Me.TxtTollTax.FieldDesc = Nothing
        Me.TxtTollTax.FieldMaxLength = 0
        Me.TxtTollTax.FieldName = Nothing
        Me.TxtTollTax.isCalculatedField = False
        Me.TxtTollTax.IsSourceFromTable = False
        Me.TxtTollTax.IsSourceFromValueList = False
        Me.TxtTollTax.IsUnique = False
        Me.TxtTollTax.Location = New System.Drawing.Point(87, 49)
        Me.TxtTollTax.MendatroryField = False
        Me.TxtTollTax.MyLinkLable1 = Nothing
        Me.TxtTollTax.MyLinkLable2 = Nothing
        Me.TxtTollTax.Name = "TxtTollTax"
        Me.TxtTollTax.ReferenceFieldDesc = Nothing
        Me.TxtTollTax.ReferenceFieldName = Nothing
        Me.TxtTollTax.ReferenceTableName = Nothing
        Me.TxtTollTax.Size = New System.Drawing.Size(82, 20)
        Me.TxtTollTax.TabIndex = 1536
        Me.TxtTollTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtFatRate
        '
        Me.TxtFatRate.CalculationExpression = Nothing
        Me.TxtFatRate.FieldCode = Nothing
        Me.TxtFatRate.FieldDesc = Nothing
        Me.TxtFatRate.FieldMaxLength = 0
        Me.TxtFatRate.FieldName = Nothing
        Me.TxtFatRate.isCalculatedField = False
        Me.TxtFatRate.IsSourceFromTable = False
        Me.TxtFatRate.IsSourceFromValueList = False
        Me.TxtFatRate.IsUnique = False
        Me.TxtFatRate.Location = New System.Drawing.Point(510, 57)
        Me.TxtFatRate.MendatroryField = False
        Me.TxtFatRate.MyLinkLable1 = Nothing
        Me.TxtFatRate.MyLinkLable2 = Nothing
        Me.TxtFatRate.Name = "TxtFatRate"
        Me.TxtFatRate.ReadOnly = True
        Me.TxtFatRate.ReferenceFieldDesc = Nothing
        Me.TxtFatRate.ReferenceFieldName = Nothing
        Me.TxtFatRate.ReferenceTableName = Nothing
        Me.TxtFatRate.Size = New System.Drawing.Size(82, 20)
        Me.TxtFatRate.TabIndex = 1535
        Me.TxtFatRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDieselMinus
        '
        Me.TxtDieselMinus.CalculationExpression = Nothing
        Me.TxtDieselMinus.FieldCode = Nothing
        Me.TxtDieselMinus.FieldDesc = Nothing
        Me.TxtDieselMinus.FieldMaxLength = 0
        Me.TxtDieselMinus.FieldName = Nothing
        Me.TxtDieselMinus.isCalculatedField = False
        Me.TxtDieselMinus.IsSourceFromTable = False
        Me.TxtDieselMinus.IsSourceFromValueList = False
        Me.TxtDieselMinus.IsUnique = False
        Me.TxtDieselMinus.Location = New System.Drawing.Point(87, 27)
        Me.TxtDieselMinus.MendatroryField = False
        Me.TxtDieselMinus.MyLinkLable1 = Nothing
        Me.TxtDieselMinus.MyLinkLable2 = Nothing
        Me.TxtDieselMinus.Name = "TxtDieselMinus"
        Me.TxtDieselMinus.ReferenceFieldDesc = Nothing
        Me.TxtDieselMinus.ReferenceFieldName = Nothing
        Me.TxtDieselMinus.ReferenceTableName = Nothing
        Me.TxtDieselMinus.Size = New System.Drawing.Size(82, 20)
        Me.TxtDieselMinus.TabIndex = 1539
        Me.TxtDieselMinus.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel12
        '
        Me.MyLabel12.FieldName = Nothing
        Me.MyLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel12.Location = New System.Drawing.Point(178, 53)
        Me.MyLabel12.Name = "MyLabel12"
        Me.MyLabel12.Size = New System.Drawing.Size(80, 16)
        Me.MyLabel12.TabIndex = 1548
        Me.MyLabel12.Text = "Barel Capacity"
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(9, 52)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(46, 16)
        Me.MyLabel7.TabIndex = 1535
        Me.MyLabel7.Text = "Toll Tax"
        '
        'MyLabel3
        '
        Me.MyLabel3.FieldName = Nothing
        Me.MyLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel3.Location = New System.Drawing.Point(424, 57)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel3.TabIndex = 1532
        Me.MyLabel3.Text = "FAT Rate RMG"
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(9, 30)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(72, 16)
        Me.MyLabel5.TabIndex = 1538
        Me.MyLabel5.Text = "Diesel Rate -"
        '
        'TxtSnfRate
        '
        Me.TxtSnfRate.CalculationExpression = Nothing
        Me.TxtSnfRate.FieldCode = Nothing
        Me.TxtSnfRate.FieldDesc = Nothing
        Me.TxtSnfRate.FieldMaxLength = 0
        Me.TxtSnfRate.FieldName = Nothing
        Me.TxtSnfRate.isCalculatedField = False
        Me.TxtSnfRate.IsSourceFromTable = False
        Me.TxtSnfRate.IsSourceFromValueList = False
        Me.TxtSnfRate.IsUnique = False
        Me.TxtSnfRate.Location = New System.Drawing.Point(510, 83)
        Me.TxtSnfRate.MendatroryField = False
        Me.TxtSnfRate.MyLinkLable1 = Nothing
        Me.TxtSnfRate.MyLinkLable2 = Nothing
        Me.TxtSnfRate.Name = "TxtSnfRate"
        Me.TxtSnfRate.ReadOnly = True
        Me.TxtSnfRate.ReferenceFieldDesc = Nothing
        Me.TxtSnfRate.ReferenceFieldName = Nothing
        Me.TxtSnfRate.ReferenceTableName = Nothing
        Me.TxtSnfRate.Size = New System.Drawing.Size(82, 20)
        Me.TxtSnfRate.TabIndex = 1535
        Me.TxtSnfRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDieselplus
        '
        Me.txtDieselplus.CalculationExpression = Nothing
        Me.txtDieselplus.FieldCode = Nothing
        Me.txtDieselplus.FieldDesc = Nothing
        Me.txtDieselplus.FieldMaxLength = 0
        Me.txtDieselplus.FieldName = Nothing
        Me.txtDieselplus.isCalculatedField = False
        Me.txtDieselplus.IsSourceFromTable = False
        Me.txtDieselplus.IsSourceFromValueList = False
        Me.txtDieselplus.IsUnique = False
        Me.txtDieselplus.Location = New System.Drawing.Point(87, 4)
        Me.txtDieselplus.MendatroryField = False
        Me.txtDieselplus.MyLinkLable1 = Nothing
        Me.txtDieselplus.MyLinkLable2 = Nothing
        Me.txtDieselplus.Name = "txtDieselplus"
        Me.txtDieselplus.ReferenceFieldDesc = Nothing
        Me.txtDieselplus.ReferenceFieldName = Nothing
        Me.txtDieselplus.ReferenceTableName = Nothing
        Me.txtDieselplus.Size = New System.Drawing.Size(82, 20)
        Me.txtDieselplus.TabIndex = 1537
        Me.txtDieselplus.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(9, 8)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(74, 16)
        Me.MyLabel1.TabIndex = 1536
        Me.MyLabel1.Text = "Diesel Rate +"
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(423, 83)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(85, 16)
        Me.MyLabel4.TabIndex = 1532
        Me.MyLabel4.Text = "SNF Rate RMG"
        '
        'TxtRateprorata
        '
        Me.TxtRateprorata.CalculationExpression = Nothing
        Me.TxtRateprorata.FieldCode = Nothing
        Me.TxtRateprorata.FieldDesc = Nothing
        Me.TxtRateprorata.FieldMaxLength = 0
        Me.TxtRateprorata.FieldName = Nothing
        Me.TxtRateprorata.isCalculatedField = False
        Me.TxtRateprorata.IsSourceFromTable = False
        Me.TxtRateprorata.IsSourceFromValueList = False
        Me.TxtRateprorata.IsUnique = False
        Me.TxtRateprorata.Location = New System.Drawing.Point(628, 101)
        Me.TxtRateprorata.MendatroryField = False
        Me.TxtRateprorata.MyLinkLable1 = Nothing
        Me.TxtRateprorata.MyLinkLable2 = Nothing
        Me.TxtRateprorata.Name = "TxtRateprorata"
        Me.TxtRateprorata.ReferenceFieldDesc = Nothing
        Me.TxtRateprorata.ReferenceFieldName = Nothing
        Me.TxtRateprorata.ReferenceTableName = Nothing
        Me.TxtRateprorata.Size = New System.Drawing.Size(82, 20)
        Me.TxtRateprorata.TabIndex = 1543
        Me.TxtRateprorata.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRateprorata.Visible = False
        '
        'TxtTDR
        '
        Me.TxtTDR.CalculationExpression = Nothing
        Me.TxtTDR.FieldCode = Nothing
        Me.TxtTDR.FieldDesc = Nothing
        Me.TxtTDR.FieldMaxLength = 0
        Me.TxtTDR.FieldName = Nothing
        Me.TxtTDR.isCalculatedField = False
        Me.TxtTDR.IsSourceFromTable = False
        Me.TxtTDR.IsSourceFromValueList = False
        Me.TxtTDR.IsUnique = False
        Me.TxtTDR.Location = New System.Drawing.Point(264, 5)
        Me.TxtTDR.MendatroryField = False
        Me.TxtTDR.MyLinkLable1 = Nothing
        Me.TxtTDR.MyLinkLable2 = Nothing
        Me.TxtTDR.Name = "TxtTDR"
        Me.TxtTDR.ReferenceFieldDesc = Nothing
        Me.TxtTDR.ReferenceFieldName = Nothing
        Me.TxtTDR.ReferenceTableName = Nothing
        Me.TxtTDR.Size = New System.Drawing.Size(82, 20)
        Me.TxtTDR.TabIndex = 1545
        Me.TxtTDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel9
        '
        Me.MyLabel9.FieldName = Nothing
        Me.MyLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(542, 105)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(61, 16)
        Me.MyLabel9.TabIndex = 1542
        Me.MyLabel9.Text = "PRORATA"
        Me.MyLabel9.Visible = False
        '
        'MyLabel10
        '
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(175, 8)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(29, 16)
        Me.MyLabel10.TabIndex = 1544
        Me.MyLabel10.Text = "TDR"
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Controls.Add(Me.txtToDate)
        Me.RadGroupBox1.Controls.Add(Me.txtFromDate)
        Me.RadGroupBox1.Controls.Add(Me.Label2)
        Me.RadGroupBox1.Controls.Add(Me.Label1)
        Me.RadGroupBox1.HeaderText = ""
        Me.RadGroupBox1.Location = New System.Drawing.Point(433, 30)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Size = New System.Drawing.Size(333, 35)
        Me.RadGroupBox1.TabIndex = 1552
        '
        'txtToDate
        '
        Me.txtToDate.CalculationExpression = Nothing
        Me.txtToDate.CustomFormat = "dd-MM-yyyy"
        Me.txtToDate.FieldCode = Nothing
        Me.txtToDate.FieldDesc = Nothing
        Me.txtToDate.FieldMaxLength = 0
        Me.txtToDate.FieldName = Nothing
        Me.txtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtToDate.isCalculatedField = False
        Me.txtToDate.IsSourceFromTable = False
        Me.txtToDate.IsSourceFromValueList = False
        Me.txtToDate.IsUnique = False
        Me.txtToDate.Location = New System.Drawing.Point(215, 9)
        Me.txtToDate.MendatroryField = False
        Me.txtToDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.MyLinkLable1 = Nothing
        Me.txtToDate.MyLinkLable2 = Nothing
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.ReferenceFieldDesc = Nothing
        Me.txtToDate.ReferenceFieldName = Nothing
        Me.txtToDate.ReferenceTableName = Nothing
        Me.txtToDate.Size = New System.Drawing.Size(82, 20)
        Me.txtToDate.TabIndex = 3
        Me.txtToDate.TabStop = False
        Me.txtToDate.Text = "13-07-2023"
        Me.txtToDate.Value = New Date(2023, 7, 13, 0, 0, 0, 0)
        '
        'txtFromDate
        '
        Me.txtFromDate.CalculationExpression = Nothing
        Me.txtFromDate.CustomFormat = "dd-MM-yyyy"
        Me.txtFromDate.FieldCode = Nothing
        Me.txtFromDate.FieldDesc = Nothing
        Me.txtFromDate.FieldMaxLength = 0
        Me.txtFromDate.FieldName = Nothing
        Me.txtFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtFromDate.isCalculatedField = False
        Me.txtFromDate.IsSourceFromTable = False
        Me.txtFromDate.IsSourceFromValueList = False
        Me.txtFromDate.IsUnique = False
        Me.txtFromDate.Location = New System.Drawing.Point(70, 9)
        Me.txtFromDate.MendatroryField = False
        Me.txtFromDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDate.MyLinkLable1 = Nothing
        Me.txtFromDate.MyLinkLable2 = Nothing
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDate.ReferenceFieldDesc = Nothing
        Me.txtFromDate.ReferenceFieldName = Nothing
        Me.txtFromDate.ReferenceTableName = Nothing
        Me.txtFromDate.Size = New System.Drawing.Size(82, 20)
        Me.txtFromDate.TabIndex = 2
        Me.txtFromDate.TabStop = False
        Me.txtFromDate.Text = "13-07-2023"
        Me.txtFromDate.Value = New Date(2023, 7, 13, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(164, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From Date"
        '
        'chkDCS
        '
        Me.chkDCS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDCS.Location = New System.Drawing.Point(802, 1)
        Me.chkDCS.Name = "chkDCS"
        Me.chkDCS.Size = New System.Drawing.Size(44, 16)
        Me.chkDCS.TabIndex = 1484
        Me.chkDCS.Text = "DCS"
        Me.chkDCS.Visible = False
        '
        'RadGroupBox2
        '
        Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadGroupBox2.Controls.Add(Me.gv1)
        Me.RadGroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadGroupBox2.HeaderText = "BMC Details"
        Me.RadGroupBox2.Location = New System.Drawing.Point(2, 220)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(764, 140)
        Me.RadGroupBox2.TabIndex = 28
        Me.RadGroupBox2.Text = "BMC Details"
        '
        'gv1
        '
        Me.gv1.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gv1.Cursor = System.Windows.Forms.Cursors.Default
        Me.gv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gv1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gv1.ForeColor = System.Drawing.Color.Black
        Me.gv1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gv1.Location = New System.Drawing.Point(10, 20)
        '
        '
        '
        Me.gv1.MasterTemplate.AllowDeleteRow = False
        Me.gv1.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gv1.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv1.MasterTemplate.ViewDefinition = TableViewDefinition1
        Me.gv1.MyExportAPI = False
        Me.gv1.MyExportFilePath = ""
        Me.gv1.MyStopExport = False
        Me.gv1.Name = "gv1"
        Me.gv1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gv1.ShowGroupPanel = False
        Me.gv1.ShowHeaderCellButtons = True
        Me.gv1.Size = New System.Drawing.Size(744, 110)
        Me.gv1.TabIndex = 17
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'txtTransporter
        '
        Me.txtTransporter.CalculationExpression = Nothing
        Me.txtTransporter.FieldCode = Nothing
        Me.txtTransporter.FieldDesc = Nothing
        Me.txtTransporter.FieldMaxLength = 0
        Me.txtTransporter.FieldName = Nothing
        Me.txtTransporter.isCalculatedField = False
        Me.txtTransporter.IsSourceFromTable = False
        Me.txtTransporter.IsSourceFromValueList = False
        Me.txtTransporter.IsUnique = False
        Me.txtTransporter.Location = New System.Drawing.Point(78, 46)
        Me.txtTransporter.MendatroryField = False
        Me.txtTransporter.MyLinkLable1 = Nothing
        Me.txtTransporter.MyLinkLable2 = Nothing
        Me.txtTransporter.Name = "txtTransporter"
        Me.txtTransporter.ReferenceFieldDesc = Nothing
        Me.txtTransporter.ReferenceFieldName = Nothing
        Me.txtTransporter.ReferenceTableName = Nothing
        Me.txtTransporter.Size = New System.Drawing.Size(348, 20)
        Me.txtTransporter.TabIndex = 1532
        '
        'lblTransporter
        '
        Me.lblTransporter.FieldName = Nothing
        Me.lblTransporter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTransporter.Location = New System.Drawing.Point(0, 48)
        Me.lblTransporter.Name = "lblTransporter"
        Me.lblTransporter.Size = New System.Drawing.Size(65, 16)
        Me.lblTransporter.TabIndex = 1531
        Me.lblTransporter.Text = "Transporter"
        '
        'txtDate
        '
        Me.txtDate.CalculationExpression = Nothing
        Me.txtDate.CustomFormat = "dd/MM/yyyy hh:mm tt"
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
        Me.txtDate.Location = New System.Drawing.Point(433, 3)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Me.RadLabel4
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(127, 18)
        Me.txtDate.TabIndex = 1524
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011 11:29 AM"
        Me.txtDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'RadLabel4
        '
        Me.RadLabel4.FieldName = Nothing
        Me.RadLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel4.Location = New System.Drawing.Point(397, 4)
        Me.RadLabel4.Name = "RadLabel4"
        Me.RadLabel4.Size = New System.Drawing.Size(30, 16)
        Me.RadLabel4.TabIndex = 1
        Me.RadLabel4.Text = "Date"
        '
        'lblTankerDesc
        '
        Me.lblTankerDesc.AutoSize = False
        Me.lblTankerDesc.BorderVisible = True
        Me.lblTankerDesc.FieldName = Nothing
        Me.lblTankerDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankerDesc.Location = New System.Drawing.Point(224, 25)
        Me.lblTankerDesc.Name = "lblTankerDesc"
        Me.lblTankerDesc.Size = New System.Drawing.Size(202, 19)
        Me.lblTankerDesc.TabIndex = 144
        Me.lblTankerDesc.TextWrap = False
        '
        'lblAbandonmentNo
        '
        Me.lblAbandonmentNo.FieldName = Nothing
        Me.lblAbandonmentNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAbandonmentNo.Location = New System.Drawing.Point(764, 11)
        Me.lblAbandonmentNo.Name = "lblAbandonmentNo"
        Me.lblAbandonmentNo.Size = New System.Drawing.Size(2, 2)
        Me.lblAbandonmentNo.TabIndex = 27
        '
        'lblTankerNo
        '
        Me.lblTankerNo.FieldName = Nothing
        Me.lblTankerNo.Location = New System.Drawing.Point(-1, 25)
        Me.lblTankerNo.Name = "lblTankerNo"
        Me.lblTankerNo.Size = New System.Drawing.Size(60, 18)
        Me.lblTankerNo.TabIndex = 119
        Me.lblTankerNo.Text = "Tanker No."
        '
        'txtTankerNo
        '
        Me.txtTankerNo.CalculationExpression = Nothing
        Me.txtTankerNo.FieldCode = Nothing
        Me.txtTankerNo.FieldDesc = Nothing
        Me.txtTankerNo.FieldMaxLength = 0
        Me.txtTankerNo.FieldName = Nothing
        Me.txtTankerNo.isCalculatedField = False
        Me.txtTankerNo.IsSourceFromTable = False
        Me.txtTankerNo.IsSourceFromValueList = False
        Me.txtTankerNo.IsUnique = False
        Me.txtTankerNo.Location = New System.Drawing.Point(78, 25)
        Me.txtTankerNo.MendatroryField = False
        Me.txtTankerNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTankerNo.MyLinkLable1 = Me.lblTankerNo
        Me.txtTankerNo.MyLinkLable2 = Nothing
        Me.txtTankerNo.MyReadOnly = False
        Me.txtTankerNo.MyShowMasterFormButton = False
        Me.txtTankerNo.Name = "txtTankerNo"
        Me.txtTankerNo.ReferenceFieldDesc = Nothing
        Me.txtTankerNo.ReferenceFieldName = Nothing
        Me.txtTankerNo.ReferenceTableName = Nothing
        Me.txtTankerNo.Size = New System.Drawing.Size(143, 19)
        Me.txtTankerNo.TabIndex = 22
        Me.txtTankerNo.Value = ""
        '
        'RadLabel1
        '
        Me.RadLabel1.FieldName = Nothing
        Me.RadLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel1.Location = New System.Drawing.Point(-1, 3)
        Me.RadLabel1.Name = "RadLabel1"
        Me.RadLabel1.Size = New System.Drawing.Size(75, 16)
        Me.RadLabel1.TabIndex = 22
        Me.RadLabel1.Text = "Document No"
        '
        'txtDocNo
        '
        Me.txtDocNo.FieldName = Nothing
        Me.txtDocNo.Location = New System.Drawing.Point(78, 1)
        Me.txtDocNo.MendatroryField = False
        Me.txtDocNo.MyCharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.txtDocNo.MyFont = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.txtDocNo.MyLinkLable1 = Me.RadLabel1
        Me.txtDocNo.MyLinkLable2 = Nothing
        Me.txtDocNo.MyMaxLength = 32767
        Me.txtDocNo.MyReadOnly = False
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.Size = New System.Drawing.Size(298, 22)
        Me.txtDocNo.TabIndex = 0
        Me.txtDocNo.TabStop = False
        Me.txtDocNo.Value = ""
        '
        'RadPageViewPage2
        '
        Me.RadPageViewPage2.Controls.Add(Me.NACCOEPF_PER)
        Me.RadPageViewPage2.Controls.Add(Me.NEMPEPF_PER)
        Me.RadPageViewPage2.Controls.Add(Me.NCOESI_PER)
        Me.RadPageViewPage2.Controls.Add(Me.NEMPESI_PER)
        Me.RadPageViewPage2.Controls.Add(Me.NTotalPF)
        Me.RadPageViewPage2.Controls.Add(Me.NTotalESI)
        Me.RadPageViewPage2.Controls.Add(Me.NESIEmployer)
        Me.RadPageViewPage2.Controls.Add(Me.NESIEmployee)
        Me.RadPageViewPage2.Controls.Add(Me.NEDLICharge)
        Me.RadPageViewPage2.Controls.Add(Me.NAdmincharge)
        Me.RadPageViewPage2.Controls.Add(Me.NemployerPF)
        Me.RadPageViewPage2.Controls.Add(Me.NempPF)
        Me.RadPageViewPage2.Controls.Add(Me.NLabourCharge)
        Me.RadPageViewPage2.Controls.Add(Me.NDailyWage)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel70)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel69)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel68)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel67)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel66)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel65)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel61)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel60)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel59)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel58)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel57)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel56)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel55)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel54)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel53)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel52)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel51)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel50)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel49)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel48)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel37)
        Me.RadPageViewPage2.Controls.Add(Me.MyLabel36)
        Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(102.0!, 22.0!)
        Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 31)
        Me.RadPageViewPage2.Name = "RadPageViewPage2"
        Me.RadPageViewPage2.Size = New System.Drawing.Size(779, 363)
        Me.RadPageViewPage2.Text = "PF/ESI Deduction"
        '
        'NACCOEPF_PER
        '
        Me.NACCOEPF_PER.AutoSize = False
        Me.NACCOEPF_PER.BorderVisible = True
        Me.NACCOEPF_PER.FieldName = Nothing
        Me.NACCOEPF_PER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NACCOEPF_PER.Location = New System.Drawing.Point(217, 119)
        Me.NACCOEPF_PER.Name = "NACCOEPF_PER"
        Me.NACCOEPF_PER.Size = New System.Drawing.Size(77, 18)
        Me.NACCOEPF_PER.TabIndex = 178
        Me.NACCOEPF_PER.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NEMPEPF_PER
        '
        Me.NEMPEPF_PER.AutoSize = False
        Me.NEMPEPF_PER.BorderVisible = True
        Me.NEMPEPF_PER.FieldName = Nothing
        Me.NEMPEPF_PER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NEMPEPF_PER.Location = New System.Drawing.Point(217, 75)
        Me.NEMPEPF_PER.Name = "NEMPEPF_PER"
        Me.NEMPEPF_PER.Size = New System.Drawing.Size(77, 18)
        Me.NEMPEPF_PER.TabIndex = 177
        Me.NEMPEPF_PER.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NCOESI_PER
        '
        Me.NCOESI_PER.AutoSize = False
        Me.NCOESI_PER.BorderVisible = True
        Me.NCOESI_PER.FieldName = Nothing
        Me.NCOESI_PER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NCOESI_PER.Location = New System.Drawing.Point(217, 215)
        Me.NCOESI_PER.Name = "NCOESI_PER"
        Me.NCOESI_PER.Size = New System.Drawing.Size(77, 18)
        Me.NCOESI_PER.TabIndex = 176
        Me.NCOESI_PER.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NEMPESI_PER
        '
        Me.NEMPESI_PER.AutoSize = False
        Me.NEMPESI_PER.BorderVisible = True
        Me.NEMPESI_PER.FieldName = Nothing
        Me.NEMPESI_PER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NEMPESI_PER.Location = New System.Drawing.Point(217, 189)
        Me.NEMPESI_PER.Name = "NEMPESI_PER"
        Me.NEMPESI_PER.Size = New System.Drawing.Size(77, 18)
        Me.NEMPESI_PER.TabIndex = 175
        Me.NEMPESI_PER.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NTotalPF
        '
        Me.NTotalPF.AutoSize = False
        Me.NTotalPF.BorderVisible = True
        Me.NTotalPF.FieldName = Nothing
        Me.NTotalPF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NTotalPF.Location = New System.Drawing.Point(134, 166)
        Me.NTotalPF.Name = "NTotalPF"
        Me.NTotalPF.Size = New System.Drawing.Size(77, 18)
        Me.NTotalPF.TabIndex = 174
        Me.NTotalPF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NTotalESI
        '
        Me.NTotalESI.AutoSize = False
        Me.NTotalESI.BorderVisible = True
        Me.NTotalESI.FieldName = Nothing
        Me.NTotalESI.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NTotalESI.Location = New System.Drawing.Point(134, 239)
        Me.NTotalESI.Name = "NTotalESI"
        Me.NTotalESI.Size = New System.Drawing.Size(77, 18)
        Me.NTotalESI.TabIndex = 173
        Me.NTotalESI.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NESIEmployer
        '
        Me.NESIEmployer.AutoSize = False
        Me.NESIEmployer.BorderVisible = True
        Me.NESIEmployer.FieldName = Nothing
        Me.NESIEmployer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NESIEmployer.Location = New System.Drawing.Point(134, 215)
        Me.NESIEmployer.Name = "NESIEmployer"
        Me.NESIEmployer.Size = New System.Drawing.Size(77, 18)
        Me.NESIEmployer.TabIndex = 169
        Me.NESIEmployer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NESIEmployee
        '
        Me.NESIEmployee.AutoSize = False
        Me.NESIEmployee.BorderVisible = True
        Me.NESIEmployee.FieldName = Nothing
        Me.NESIEmployee.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NESIEmployee.Location = New System.Drawing.Point(134, 189)
        Me.NESIEmployee.Name = "NESIEmployee"
        Me.NESIEmployee.Size = New System.Drawing.Size(77, 18)
        Me.NESIEmployee.TabIndex = 172
        Me.NESIEmployee.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NEDLICharge
        '
        Me.NEDLICharge.AutoSize = False
        Me.NEDLICharge.BorderVisible = True
        Me.NEDLICharge.FieldName = Nothing
        Me.NEDLICharge.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NEDLICharge.Location = New System.Drawing.Point(134, 142)
        Me.NEDLICharge.Name = "NEDLICharge"
        Me.NEDLICharge.Size = New System.Drawing.Size(77, 18)
        Me.NEDLICharge.TabIndex = 171
        Me.NEDLICharge.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NAdmincharge
        '
        Me.NAdmincharge.AutoSize = False
        Me.NAdmincharge.BorderVisible = True
        Me.NAdmincharge.FieldName = Nothing
        Me.NAdmincharge.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NAdmincharge.Location = New System.Drawing.Point(134, 119)
        Me.NAdmincharge.Name = "NAdmincharge"
        Me.NAdmincharge.Size = New System.Drawing.Size(77, 18)
        Me.NAdmincharge.TabIndex = 170
        Me.NAdmincharge.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NemployerPF
        '
        Me.NemployerPF.AutoSize = False
        Me.NemployerPF.BorderVisible = True
        Me.NemployerPF.FieldName = Nothing
        Me.NemployerPF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NemployerPF.Location = New System.Drawing.Point(134, 97)
        Me.NemployerPF.Name = "NemployerPF"
        Me.NemployerPF.Size = New System.Drawing.Size(77, 18)
        Me.NemployerPF.TabIndex = 169
        Me.NemployerPF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NempPF
        '
        Me.NempPF.AutoSize = False
        Me.NempPF.BorderVisible = True
        Me.NempPF.FieldName = Nothing
        Me.NempPF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NempPF.Location = New System.Drawing.Point(134, 75)
        Me.NempPF.Name = "NempPF"
        Me.NempPF.Size = New System.Drawing.Size(77, 18)
        Me.NempPF.TabIndex = 168
        Me.NempPF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NLabourCharge
        '
        Me.NLabourCharge.AutoSize = False
        Me.NLabourCharge.BorderVisible = True
        Me.NLabourCharge.FieldName = Nothing
        Me.NLabourCharge.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NLabourCharge.Location = New System.Drawing.Point(134, 54)
        Me.NLabourCharge.Name = "NLabourCharge"
        Me.NLabourCharge.Size = New System.Drawing.Size(77, 18)
        Me.NLabourCharge.TabIndex = 167
        Me.NLabourCharge.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'NDailyWage
        '
        Me.NDailyWage.AutoSize = False
        Me.NDailyWage.BorderVisible = True
        Me.NDailyWage.FieldName = Nothing
        Me.NDailyWage.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NDailyWage.Location = New System.Drawing.Point(134, 32)
        Me.NDailyWage.Name = "NDailyWage"
        Me.NDailyWage.Size = New System.Drawing.Size(77, 18)
        Me.NDailyWage.TabIndex = 166
        Me.NDailyWage.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel70
        '
        Me.MyLabel70.FieldName = Nothing
        Me.MyLabel70.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel70.Location = New System.Drawing.Point(13, 101)
        Me.MyLabel70.Name = "MyLabel70"
        Me.MyLabel70.Size = New System.Drawing.Size(102, 16)
        Me.MyLabel70.TabIndex = 139
        Me.MyLabel70.Text = "Employer PF(CPF)"
        '
        'MyLabel69
        '
        Me.MyLabel69.FieldName = Nothing
        Me.MyLabel69.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel69.Location = New System.Drawing.Point(441, 152)
        Me.MyLabel69.Name = "MyLabel69"
        Me.MyLabel69.Size = New System.Drawing.Size(91, 16)
        Me.MyLabel69.TabIndex = 136
        Me.MyLabel69.Text = "ACCOEPF_PER"
        Me.MyLabel69.Visible = False
        '
        'MyLabel68
        '
        Me.MyLabel68.FieldName = Nothing
        Me.MyLabel68.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel68.Location = New System.Drawing.Point(441, 128)
        Me.MyLabel68.Name = "MyLabel68"
        Me.MyLabel68.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel68.TabIndex = 134
        Me.MyLabel68.Text = "EMPEPF_PER"
        Me.MyLabel68.Visible = False
        '
        'MyLabel67
        '
        Me.MyLabel67.FieldName = Nothing
        Me.MyLabel67.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel67.Location = New System.Drawing.Point(538, 34)
        Me.MyLabel67.Name = "MyLabel67"
        Me.MyLabel67.Size = New System.Drawing.Size(79, 16)
        Me.MyLabel67.TabIndex = 132
        Me.MyLabel67.Text = "EMPESI_PER"
        Me.MyLabel67.Visible = False
        '
        'MyLabel66
        '
        Me.MyLabel66.FieldName = Nothing
        Me.MyLabel66.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel66.Location = New System.Drawing.Point(538, 18)
        Me.MyLabel66.Name = "MyLabel66"
        Me.MyLabel66.Size = New System.Drawing.Size(71, 16)
        Me.MyLabel66.TabIndex = 130
        Me.MyLabel66.Text = "COESI_PER"
        Me.MyLabel66.Visible = False
        '
        'MyLabel65
        '
        Me.MyLabel65.FieldName = Nothing
        Me.MyLabel65.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel65.Location = New System.Drawing.Point(14, 32)
        Me.MyLabel65.Name = "MyLabel65"
        Me.MyLabel65.Size = New System.Drawing.Size(64, 16)
        Me.MyLabel65.TabIndex = 128
        Me.MyLabel65.Text = "Daily Wage"
        '
        'MyLabel61
        '
        Me.MyLabel61.FieldName = Nothing
        Me.MyLabel61.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel61.Location = New System.Drawing.Point(321, 215)
        Me.MyLabel61.Name = "MyLabel61"
        Me.MyLabel61.Size = New System.Drawing.Size(160, 16)
        Me.MyLabel61.TabIndex = 127
        Me.MyLabel61.Text = "ESI Employee + ESI Employer"
        Me.MyLabel61.Visible = False
        '
        'MyLabel60
        '
        Me.MyLabel60.FieldName = Nothing
        Me.MyLabel60.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel60.Location = New System.Drawing.Point(321, 192)
        Me.MyLabel60.Name = "MyLabel60"
        Me.MyLabel60.Size = New System.Drawing.Size(86, 16)
        Me.MyLabel60.TabIndex = 126
        Me.MyLabel60.Text = "Labour × 3.25%"
        Me.MyLabel60.Visible = False
        '
        'MyLabel59
        '
        Me.MyLabel59.FieldName = Nothing
        Me.MyLabel59.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel59.Location = New System.Drawing.Point(321, 169)
        Me.MyLabel59.Name = "MyLabel59"
        Me.MyLabel59.Size = New System.Drawing.Size(86, 16)
        Me.MyLabel59.TabIndex = 125
        Me.MyLabel59.Text = "Labour × 0.75%"
        Me.MyLabel59.Visible = False
        '
        'MyLabel58
        '
        Me.MyLabel58.FieldName = Nothing
        Me.MyLabel58.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel58.Location = New System.Drawing.Point(529, 144)
        Me.MyLabel58.Name = "MyLabel58"
        Me.MyLabel58.Size = New System.Drawing.Size(236, 16)
        Me.MyLabel58.TabIndex = 124
        Me.MyLabel58.Text = "Employee PF + Employer PF + Admin + EDLI"
        Me.MyLabel58.Visible = False
        '
        'MyLabel57
        '
        Me.MyLabel57.FieldName = Nothing
        Me.MyLabel57.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel57.Location = New System.Drawing.Point(529, 121)
        Me.MyLabel57.Name = "MyLabel57"
        Me.MyLabel57.Size = New System.Drawing.Size(80, 16)
        Me.MyLabel57.TabIndex = 123
        Me.MyLabel57.Text = "Labour × 0.5%"
        Me.MyLabel57.Visible = False
        '
        'MyLabel56
        '
        Me.MyLabel56.FieldName = Nothing
        Me.MyLabel56.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel56.Location = New System.Drawing.Point(529, 98)
        Me.MyLabel56.Name = "MyLabel56"
        Me.MyLabel56.Size = New System.Drawing.Size(80, 16)
        Me.MyLabel56.TabIndex = 122
        Me.MyLabel56.Text = "Labour × 0.5%"
        Me.MyLabel56.Visible = False
        '
        'MyLabel55
        '
        Me.MyLabel55.FieldName = Nothing
        Me.MyLabel55.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel55.Location = New System.Drawing.Point(529, 75)
        Me.MyLabel55.Name = "MyLabel55"
        Me.MyLabel55.Size = New System.Drawing.Size(165, 16)
        Me.MyLabel55.TabIndex = 107
        Me.MyLabel55.Text = "Labour × 12% (or applicable %)"
        Me.MyLabel55.Visible = False
        '
        'MyLabel54
        '
        Me.MyLabel54.FieldName = Nothing
        Me.MyLabel54.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel54.Location = New System.Drawing.Point(529, 52)
        Me.MyLabel54.Name = "MyLabel54"
        Me.MyLabel54.Size = New System.Drawing.Size(136, 16)
        Me.MyLabel54.TabIndex = 121
        Me.MyLabel54.Text = "No. of Days × Daily Wage"
        Me.MyLabel54.Visible = False
        '
        'MyLabel53
        '
        Me.MyLabel53.FieldName = Nothing
        Me.MyLabel53.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel53.Location = New System.Drawing.Point(14, 238)
        Me.MyLabel53.Name = "MyLabel53"
        Me.MyLabel53.Size = New System.Drawing.Size(53, 16)
        Me.MyLabel53.TabIndex = 120
        Me.MyLabel53.Text = "Total ESI"
        '
        'MyLabel52
        '
        Me.MyLabel52.FieldName = Nothing
        Me.MyLabel52.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel52.Location = New System.Drawing.Point(13, 215)
        Me.MyLabel52.Name = "MyLabel52"
        Me.MyLabel52.Size = New System.Drawing.Size(75, 16)
        Me.MyLabel52.TabIndex = 118
        Me.MyLabel52.Text = "ESI Employer"
        '
        'MyLabel51
        '
        Me.MyLabel51.FieldName = Nothing
        Me.MyLabel51.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel51.Location = New System.Drawing.Point(13, 192)
        Me.MyLabel51.Name = "MyLabel51"
        Me.MyLabel51.Size = New System.Drawing.Size(78, 16)
        Me.MyLabel51.TabIndex = 116
        Me.MyLabel51.Text = "ESI Employee"
        '
        'MyLabel50
        '
        Me.MyLabel50.FieldName = Nothing
        Me.MyLabel50.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel50.Location = New System.Drawing.Point(13, 169)
        Me.MyLabel50.Name = "MyLabel50"
        Me.MyLabel50.Size = New System.Drawing.Size(49, 16)
        Me.MyLabel50.TabIndex = 114
        Me.MyLabel50.Text = "Total PF"
        '
        'MyLabel49
        '
        Me.MyLabel49.FieldName = Nothing
        Me.MyLabel49.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel49.Location = New System.Drawing.Point(13, 146)
        Me.MyLabel49.Name = "MyLabel49"
        Me.MyLabel49.Size = New System.Drawing.Size(77, 16)
        Me.MyLabel49.TabIndex = 112
        Me.MyLabel49.Text = "EDLI Charges"
        '
        'MyLabel48
        '
        Me.MyLabel48.FieldName = Nothing
        Me.MyLabel48.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel48.Location = New System.Drawing.Point(13, 123)
        Me.MyLabel48.Name = "MyLabel48"
        Me.MyLabel48.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel48.TabIndex = 110
        Me.MyLabel48.Text = "Admin Charges"
        '
        'MyLabel37
        '
        Me.MyLabel37.FieldName = Nothing
        Me.MyLabel37.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel37.Location = New System.Drawing.Point(13, 77)
        Me.MyLabel37.Name = "MyLabel37"
        Me.MyLabel37.Size = New System.Drawing.Size(104, 16)
        Me.MyLabel37.TabIndex = 108
        Me.MyLabel37.Text = "Employee PF(CPF)"
        '
        'MyLabel36
        '
        Me.MyLabel36.FieldName = Nothing
        Me.MyLabel36.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel36.Location = New System.Drawing.Point(13, 54)
        Me.MyLabel36.Name = "MyLabel36"
        Me.MyLabel36.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel36.TabIndex = 106
        Me.MyLabel36.Text = "Labour Amount"
        '
        'RadPageViewPage4
        '
        Me.RadPageViewPage4.Controls.Add(Me.TotalESIAmt)
        Me.RadPageViewPage4.Controls.Add(Me.TotalEPFAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel64)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel38)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalDeduction)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalAddition)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel39)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel63)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel47)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel46)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel45)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel44)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel43)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel42)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel41)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel40)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalIceCharge)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalAmount)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalFatSnfShortage)
        Me.RadPageViewPage4.Controls.Add(Me.TxtGrossAmount)
        Me.RadPageViewPage4.Controls.Add(Me.TxtTotalTollTax)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel34)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel33)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel32)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel31)
        Me.RadPageViewPage4.Controls.Add(Me.TxtBMCTotal)
        Me.RadPageViewPage4.Controls.Add(Me.TxtBMCDiesel)
        Me.RadPageViewPage4.Controls.Add(Me.TxtBMCProrataamt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel27)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel26)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel62)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel25)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel24)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel23)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel22)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel21)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel20)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel19)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel18)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel17)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel16)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel15)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel14)
        Me.RadPageViewPage4.Controls.Add(Me.lblInvoiceDiscAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel6)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel32)
        Me.RadPageViewPage4.Controls.Add(Me.lblAddCharges1)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel9)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel27)
        Me.RadPageViewPage4.Controls.Add(Me.lblTotRAmt)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel25)
        Me.RadPageViewPage4.Controls.Add(Me.lblTaxAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblAmtAfterDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.lblDiscountAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblAmtWithDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel22)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel19)
        Me.RadPageViewPage4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage4.ItemSize = New System.Drawing.SizeF(37.0!, 22.0!)
        Me.RadPageViewPage4.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage4.Name = "RadPageViewPage4"
        Me.RadPageViewPage4.Size = New System.Drawing.Size(779, 359)
        Me.RadPageViewPage4.Text = "Total"
        '
        'TotalESIAmt
        '
        Me.TotalESIAmt.AutoSize = False
        Me.TotalESIAmt.BorderVisible = True
        Me.TotalESIAmt.FieldName = Nothing
        Me.TotalESIAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TotalESIAmt.Location = New System.Drawing.Point(326, 315)
        Me.TotalESIAmt.Name = "TotalESIAmt"
        Me.TotalESIAmt.Size = New System.Drawing.Size(77, 18)
        Me.TotalESIAmt.TabIndex = 1546
        Me.TotalESIAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TotalEPFAmt
        '
        Me.TotalEPFAmt.AutoSize = False
        Me.TotalEPFAmt.BorderVisible = True
        Me.TotalEPFAmt.FieldName = Nothing
        Me.TotalEPFAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TotalEPFAmt.Location = New System.Drawing.Point(326, 293)
        Me.TotalEPFAmt.Name = "TotalEPFAmt"
        Me.TotalEPFAmt.Size = New System.Drawing.Size(77, 18)
        Me.TotalEPFAmt.TabIndex = 1545
        Me.TotalEPFAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel64
        '
        Me.MyLabel64.FieldName = Nothing
        Me.MyLabel64.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel64.Location = New System.Drawing.Point(35, 315)
        Me.MyLabel64.Name = "MyLabel64"
        Me.MyLabel64.Size = New System.Drawing.Size(64, 16)
        Me.MyLabel64.TabIndex = 1544
        Me.MyLabel64.Text = "TOTAL ESI"
        '
        'MyLabel38
        '
        Me.MyLabel38.FieldName = Nothing
        Me.MyLabel38.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel38.Location = New System.Drawing.Point(35, 293)
        Me.MyLabel38.Name = "MyLabel38"
        Me.MyLabel38.Size = New System.Drawing.Size(68, 16)
        Me.MyLabel38.TabIndex = 1543
        Me.MyLabel38.Text = "TOTAL EPF"
        '
        'TxtTotalDeduction
        '
        Me.TxtTotalDeduction.CalculationExpression = Nothing
        Me.TxtTotalDeduction.FieldCode = Nothing
        Me.TxtTotalDeduction.FieldDesc = Nothing
        Me.TxtTotalDeduction.FieldMaxLength = 0
        Me.TxtTotalDeduction.FieldName = Nothing
        Me.TxtTotalDeduction.isCalculatedField = False
        Me.TxtTotalDeduction.IsSourceFromTable = False
        Me.TxtTotalDeduction.IsSourceFromValueList = False
        Me.TxtTotalDeduction.IsUnique = False
        Me.TxtTotalDeduction.Location = New System.Drawing.Point(326, 269)
        Me.TxtTotalDeduction.MendatroryField = False
        Me.TxtTotalDeduction.MyLinkLable1 = Nothing
        Me.TxtTotalDeduction.MyLinkLable2 = Nothing
        Me.TxtTotalDeduction.Name = "TxtTotalDeduction"
        Me.TxtTotalDeduction.ReferenceFieldDesc = Nothing
        Me.TxtTotalDeduction.ReferenceFieldName = Nothing
        Me.TxtTotalDeduction.ReferenceTableName = Nothing
        Me.TxtTotalDeduction.Size = New System.Drawing.Size(77, 20)
        Me.TxtTotalDeduction.TabIndex = 1541
        Me.TxtTotalDeduction.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotalAddition
        '
        Me.TxtTotalAddition.CalculationExpression = Nothing
        Me.TxtTotalAddition.FieldCode = Nothing
        Me.TxtTotalAddition.FieldDesc = Nothing
        Me.TxtTotalAddition.FieldMaxLength = 0
        Me.TxtTotalAddition.FieldName = Nothing
        Me.TxtTotalAddition.isCalculatedField = False
        Me.TxtTotalAddition.IsSourceFromTable = False
        Me.TxtTotalAddition.IsSourceFromValueList = False
        Me.TxtTotalAddition.IsUnique = False
        Me.TxtTotalAddition.Location = New System.Drawing.Point(326, 247)
        Me.TxtTotalAddition.MendatroryField = False
        Me.TxtTotalAddition.MyLinkLable1 = Nothing
        Me.TxtTotalAddition.MyLinkLable2 = Nothing
        Me.TxtTotalAddition.Name = "TxtTotalAddition"
        Me.TxtTotalAddition.ReferenceFieldDesc = Nothing
        Me.TxtTotalAddition.ReferenceFieldName = Nothing
        Me.TxtTotalAddition.ReferenceTableName = Nothing
        Me.TxtTotalAddition.Size = New System.Drawing.Size(77, 20)
        Me.TxtTotalAddition.TabIndex = 1540
        Me.TxtTotalAddition.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'MyLabel39
        '
        Me.MyLabel39.FieldName = Nothing
        Me.MyLabel39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel39.Location = New System.Drawing.Point(35, 271)
        Me.MyLabel39.Name = "MyLabel39"
        Me.MyLabel39.Size = New System.Drawing.Size(113, 16)
        Me.MyLabel39.TabIndex = 164
        Me.MyLabel39.Text = "TOTAL DEDUCTION"
        '
        'MyLabel63
        '
        Me.MyLabel63.FieldName = Nothing
        Me.MyLabel63.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel63.Location = New System.Drawing.Point(35, 249)
        Me.MyLabel63.Name = "MyLabel63"
        Me.MyLabel63.Size = New System.Drawing.Size(100, 16)
        Me.MyLabel63.TabIndex = 183
        Me.MyLabel63.Text = "TOTAL ADDITION"
        '
        'MyLabel47
        '
        Me.MyLabel47.AutoSize = False
        Me.MyLabel47.BorderVisible = True
        Me.MyLabel47.FieldName = Nothing
        Me.MyLabel47.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel47.Location = New System.Drawing.Point(581, 50)
        Me.MyLabel47.Name = "MyLabel47"
        Me.MyLabel47.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel47.TabIndex = 181
        Me.MyLabel47.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel47.Visible = False
        '
        'MyLabel46
        '
        Me.MyLabel46.AutoSize = False
        Me.MyLabel46.BorderVisible = True
        Me.MyLabel46.FieldName = Nothing
        Me.MyLabel46.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel46.Location = New System.Drawing.Point(543, 26)
        Me.MyLabel46.Name = "MyLabel46"
        Me.MyLabel46.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel46.TabIndex = 180
        Me.MyLabel46.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel46.Visible = False
        '
        'MyLabel45
        '
        Me.MyLabel45.AutoSize = False
        Me.MyLabel45.BorderVisible = True
        Me.MyLabel45.FieldName = Nothing
        Me.MyLabel45.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel45.Location = New System.Drawing.Point(326, 126)
        Me.MyLabel45.Name = "MyLabel45"
        Me.MyLabel45.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel45.TabIndex = 179
        Me.MyLabel45.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel45.Visible = False
        '
        'MyLabel44
        '
        Me.MyLabel44.AutoSize = False
        Me.MyLabel44.BorderVisible = True
        Me.MyLabel44.FieldName = Nothing
        Me.MyLabel44.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel44.Location = New System.Drawing.Point(243, 126)
        Me.MyLabel44.Name = "MyLabel44"
        Me.MyLabel44.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel44.TabIndex = 178
        Me.MyLabel44.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel44.Visible = False
        '
        'MyLabel43
        '
        Me.MyLabel43.AutoSize = False
        Me.MyLabel43.BorderVisible = True
        Me.MyLabel43.FieldName = Nothing
        Me.MyLabel43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel43.Location = New System.Drawing.Point(326, 102)
        Me.MyLabel43.Name = "MyLabel43"
        Me.MyLabel43.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel43.TabIndex = 177
        Me.MyLabel43.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel43.Visible = False
        '
        'MyLabel42
        '
        Me.MyLabel42.AutoSize = False
        Me.MyLabel42.BorderVisible = True
        Me.MyLabel42.FieldName = Nothing
        Me.MyLabel42.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel42.Location = New System.Drawing.Point(243, 101)
        Me.MyLabel42.Name = "MyLabel42"
        Me.MyLabel42.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel42.TabIndex = 176
        Me.MyLabel42.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel42.Visible = False
        '
        'MyLabel41
        '
        Me.MyLabel41.AutoSize = False
        Me.MyLabel41.BorderVisible = True
        Me.MyLabel41.FieldName = Nothing
        Me.MyLabel41.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel41.Location = New System.Drawing.Point(326, 76)
        Me.MyLabel41.Name = "MyLabel41"
        Me.MyLabel41.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel41.TabIndex = 175
        Me.MyLabel41.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel41.Visible = False
        '
        'MyLabel40
        '
        Me.MyLabel40.AutoSize = False
        Me.MyLabel40.BorderVisible = True
        Me.MyLabel40.FieldName = Nothing
        Me.MyLabel40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel40.Location = New System.Drawing.Point(243, 76)
        Me.MyLabel40.Name = "MyLabel40"
        Me.MyLabel40.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel40.TabIndex = 174
        Me.MyLabel40.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel40.Visible = False
        '
        'TxtTotalIceCharge
        '
        Me.TxtTotalIceCharge.AutoSize = False
        Me.TxtTotalIceCharge.BorderVisible = True
        Me.TxtTotalIceCharge.FieldName = Nothing
        Me.TxtTotalIceCharge.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTotalIceCharge.Location = New System.Drawing.Point(326, 186)
        Me.TxtTotalIceCharge.Name = "TxtTotalIceCharge"
        Me.TxtTotalIceCharge.Size = New System.Drawing.Size(77, 18)
        Me.TxtTotalIceCharge.TabIndex = 173
        Me.TxtTotalIceCharge.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtTotalAmount
        '
        Me.TxtTotalAmount.AutoSize = False
        Me.TxtTotalAmount.BorderVisible = True
        Me.TxtTotalAmount.FieldName = Nothing
        Me.TxtTotalAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTotalAmount.Location = New System.Drawing.Point(326, 206)
        Me.TxtTotalAmount.Name = "TxtTotalAmount"
        Me.TxtTotalAmount.Size = New System.Drawing.Size(77, 18)
        Me.TxtTotalAmount.TabIndex = 172
        Me.TxtTotalAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtTotalFatSnfShortage
        '
        Me.TxtTotalFatSnfShortage.AutoSize = False
        Me.TxtTotalFatSnfShortage.BorderVisible = True
        Me.TxtTotalFatSnfShortage.FieldName = Nothing
        Me.TxtTotalFatSnfShortage.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTotalFatSnfShortage.Location = New System.Drawing.Point(326, 226)
        Me.TxtTotalFatSnfShortage.Name = "TxtTotalFatSnfShortage"
        Me.TxtTotalFatSnfShortage.Size = New System.Drawing.Size(77, 18)
        Me.TxtTotalFatSnfShortage.TabIndex = 165
        Me.TxtTotalFatSnfShortage.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtGrossAmount
        '
        Me.TxtGrossAmount.AutoSize = False
        Me.TxtGrossAmount.BorderVisible = True
        Me.TxtGrossAmount.FieldName = Nothing
        Me.TxtGrossAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtGrossAmount.Location = New System.Drawing.Point(326, 336)
        Me.TxtGrossAmount.Name = "TxtGrossAmount"
        Me.TxtGrossAmount.Size = New System.Drawing.Size(77, 18)
        Me.TxtGrossAmount.TabIndex = 165
        Me.TxtGrossAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtTotalTollTax
        '
        Me.TxtTotalTollTax.AutoSize = False
        Me.TxtTotalTollTax.BorderVisible = True
        Me.TxtTotalTollTax.FieldName = Nothing
        Me.TxtTotalTollTax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTotalTollTax.Location = New System.Drawing.Point(326, 166)
        Me.TxtTotalTollTax.Name = "TxtTotalTollTax"
        Me.TxtTotalTollTax.Size = New System.Drawing.Size(77, 18)
        Me.TxtTotalTollTax.TabIndex = 171
        Me.TxtTotalTollTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel34
        '
        Me.MyLabel34.AutoSize = False
        Me.MyLabel34.BorderVisible = True
        Me.MyLabel34.FieldName = Nothing
        Me.MyLabel34.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel34.Location = New System.Drawing.Point(498, 50)
        Me.MyLabel34.Name = "MyLabel34"
        Me.MyLabel34.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel34.TabIndex = 170
        Me.MyLabel34.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel34.Visible = False
        '
        'MyLabel33
        '
        Me.MyLabel33.AutoSize = False
        Me.MyLabel33.BorderVisible = True
        Me.MyLabel33.FieldName = Nothing
        Me.MyLabel33.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel33.Location = New System.Drawing.Point(149, 124)
        Me.MyLabel33.Name = "MyLabel33"
        Me.MyLabel33.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel33.TabIndex = 169
        Me.MyLabel33.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel33.Visible = False
        '
        'MyLabel32
        '
        Me.MyLabel32.AutoSize = False
        Me.MyLabel32.BorderVisible = True
        Me.MyLabel32.FieldName = Nothing
        Me.MyLabel32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel32.Location = New System.Drawing.Point(149, 100)
        Me.MyLabel32.Name = "MyLabel32"
        Me.MyLabel32.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel32.TabIndex = 168
        Me.MyLabel32.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel32.Visible = False
        '
        'MyLabel31
        '
        Me.MyLabel31.AutoSize = False
        Me.MyLabel31.BorderVisible = True
        Me.MyLabel31.FieldName = Nothing
        Me.MyLabel31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel31.Location = New System.Drawing.Point(149, 76)
        Me.MyLabel31.Name = "MyLabel31"
        Me.MyLabel31.Size = New System.Drawing.Size(77, 18)
        Me.MyLabel31.TabIndex = 167
        Me.MyLabel31.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.MyLabel31.Visible = False
        '
        'TxtBMCTotal
        '
        Me.TxtBMCTotal.AutoSize = False
        Me.TxtBMCTotal.BorderVisible = True
        Me.TxtBMCTotal.FieldName = Nothing
        Me.TxtBMCTotal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBMCTotal.Location = New System.Drawing.Point(326, 50)
        Me.TxtBMCTotal.Name = "TxtBMCTotal"
        Me.TxtBMCTotal.Size = New System.Drawing.Size(77, 18)
        Me.TxtBMCTotal.TabIndex = 166
        Me.TxtBMCTotal.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtBMCDiesel
        '
        Me.TxtBMCDiesel.AutoSize = False
        Me.TxtBMCDiesel.BorderVisible = True
        Me.TxtBMCDiesel.FieldName = Nothing
        Me.TxtBMCDiesel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBMCDiesel.Location = New System.Drawing.Point(243, 50)
        Me.TxtBMCDiesel.Name = "TxtBMCDiesel"
        Me.TxtBMCDiesel.Size = New System.Drawing.Size(77, 18)
        Me.TxtBMCDiesel.TabIndex = 165
        Me.TxtBMCDiesel.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtBMCProrataamt
        '
        Me.TxtBMCProrataamt.AutoSize = False
        Me.TxtBMCProrataamt.BorderVisible = True
        Me.TxtBMCProrataamt.FieldName = Nothing
        Me.TxtBMCProrataamt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBMCProrataamt.Location = New System.Drawing.Point(149, 50)
        Me.TxtBMCProrataamt.Name = "TxtBMCProrataamt"
        Me.TxtBMCProrataamt.Size = New System.Drawing.Size(77, 18)
        Me.TxtBMCProrataamt.TabIndex = 164
        Me.TxtBMCProrataamt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel27
        '
        Me.MyLabel27.FieldName = Nothing
        Me.MyLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel27.Location = New System.Drawing.Point(34, 337)
        Me.MyLabel27.Name = "MyLabel27"
        Me.MyLabel27.Size = New System.Drawing.Size(99, 16)
        Me.MyLabel27.TabIndex = 163
        Me.MyLabel27.Text = "GROSS AMOUNT"
        '
        'MyLabel26
        '
        Me.MyLabel26.FieldName = Nothing
        Me.MyLabel26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel26.Location = New System.Drawing.Point(35, 227)
        Me.MyLabel26.Name = "MyLabel26"
        Me.MyLabel26.Size = New System.Drawing.Size(120, 16)
        Me.MyLabel26.TabIndex = 162
        Me.MyLabel26.Text = "FAT SNF SHORTAGE"
        '
        'MyLabel62
        '
        Me.MyLabel62.FieldName = Nothing
        Me.MyLabel62.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel62.Location = New System.Drawing.Point(35, 207)
        Me.MyLabel62.Name = "MyLabel62"
        Me.MyLabel62.Size = New System.Drawing.Size(95, 16)
        Me.MyLabel62.TabIndex = 161
        Me.MyLabel62.Text = "TOTAL AMOUNT"
        '
        'MyLabel25
        '
        Me.MyLabel25.FieldName = Nothing
        Me.MyLabel25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel25.Location = New System.Drawing.Point(436, 72)
        Me.MyLabel25.Name = "MyLabel25"
        Me.MyLabel25.Size = New System.Drawing.Size(95, 16)
        Me.MyLabel25.TabIndex = 161
        Me.MyLabel25.Text = "TOTAL AMOUNT"
        Me.MyLabel25.Visible = False
        '
        'MyLabel24
        '
        Me.MyLabel24.FieldName = Nothing
        Me.MyLabel24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel24.Location = New System.Drawing.Point(35, 187)
        Me.MyLabel24.Name = "MyLabel24"
        Me.MyLabel24.Size = New System.Drawing.Size(77, 16)
        Me.MyLabel24.TabIndex = 160
        Me.MyLabel24.Text = "ICE CHARGE"
        '
        'MyLabel23
        '
        Me.MyLabel23.FieldName = Nothing
        Me.MyLabel23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel23.Location = New System.Drawing.Point(35, 167)
        Me.MyLabel23.Name = "MyLabel23"
        Me.MyLabel23.Size = New System.Drawing.Size(60, 16)
        Me.MyLabel23.TabIndex = 159
        Me.MyLabel23.Text = "TOLL TAX"
        '
        'MyLabel22
        '
        Me.MyLabel22.FieldName = Nothing
        Me.MyLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel22.Location = New System.Drawing.Point(462, 28)
        Me.MyLabel22.Name = "MyLabel22"
        Me.MyLabel22.Size = New System.Drawing.Size(56, 16)
        Me.MyLabel22.TabIndex = 158
        Me.MyLabel22.Text = "TOTAL=>"
        Me.MyLabel22.Visible = False
        '
        'MyLabel21
        '
        Me.MyLabel21.FieldName = Nothing
        Me.MyLabel21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel21.Location = New System.Drawing.Point(35, 126)
        Me.MyLabel21.Name = "MyLabel21"
        Me.MyLabel21.Size = New System.Drawing.Size(36, 16)
        Me.MyLabel21.TabIndex = 157
        Me.MyLabel21.Text = "NMG:"
        Me.MyLabel21.Visible = False
        '
        'MyLabel20
        '
        Me.MyLabel20.FieldName = Nothing
        Me.MyLabel20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel20.Location = New System.Drawing.Point(34, 102)
        Me.MyLabel20.Name = "MyLabel20"
        Me.MyLabel20.Size = New System.Drawing.Size(35, 16)
        Me.MyLabel20.TabIndex = 156
        Me.MyLabel20.Text = "MCC:"
        Me.MyLabel20.Visible = False
        '
        'MyLabel19
        '
        Me.MyLabel19.FieldName = Nothing
        Me.MyLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel19.Location = New System.Drawing.Point(34, 80)
        Me.MyLabel19.Name = "MyLabel19"
        Me.MyLabel19.Size = New System.Drawing.Size(36, 16)
        Me.MyLabel19.TabIndex = 156
        Me.MyLabel19.Text = "RMG:"
        Me.MyLabel19.Visible = False
        '
        'MyLabel18
        '
        Me.MyLabel18.FieldName = Nothing
        Me.MyLabel18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel18.Location = New System.Drawing.Point(35, 52)
        Me.MyLabel18.Name = "MyLabel18"
        Me.MyLabel18.Size = New System.Drawing.Size(34, 16)
        Me.MyLabel18.TabIndex = 155
        Me.MyLabel18.Text = "BMC:"
        '
        'MyLabel17
        '
        Me.MyLabel17.FieldName = Nothing
        Me.MyLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel17.Location = New System.Drawing.Point(341, 20)
        Me.MyLabel17.Name = "MyLabel17"
        Me.MyLabel17.Size = New System.Drawing.Size(43, 16)
        Me.MyLabel17.TabIndex = 154
        Me.MyLabel17.Text = "TOTAL"
        '
        'MyLabel16
        '
        Me.MyLabel16.FieldName = Nothing
        Me.MyLabel16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel16.Location = New System.Drawing.Point(246, 20)
        Me.MyLabel16.Name = "MyLabel16"
        Me.MyLabel16.Size = New System.Drawing.Size(70, 16)
        Me.MyLabel16.TabIndex = 153
        Me.MyLabel16.Text = "DIESEL(RD)"
        '
        'MyLabel15
        '
        Me.MyLabel15.FieldName = Nothing
        Me.MyLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel15.Location = New System.Drawing.Point(149, 20)
        Me.MyLabel15.Name = "MyLabel15"
        Me.MyLabel15.Size = New System.Drawing.Size(88, 16)
        Me.MyLabel15.TabIndex = 152
        Me.MyLabel15.Text = "PRORATA AMT"
        '
        'MyLabel14
        '
        Me.MyLabel14.FieldName = Nothing
        Me.MyLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel14.Location = New System.Drawing.Point(31, 20)
        Me.MyLabel14.Name = "MyLabel14"
        Me.MyLabel14.Size = New System.Drawing.Size(112, 16)
        Me.MyLabel14.TabIndex = 151
        Me.MyLabel14.Text = "CATEGORY HEADS"
        '
        'lblInvoiceDiscAmt
        '
        Me.lblInvoiceDiscAmt.AutoSize = False
        Me.lblInvoiceDiscAmt.BorderVisible = True
        Me.lblInvoiceDiscAmt.FieldName = Nothing
        Me.lblInvoiceDiscAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceDiscAmt.Location = New System.Drawing.Point(628, 155)
        Me.lblInvoiceDiscAmt.Name = "lblInvoiceDiscAmt"
        Me.lblInvoiceDiscAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblInvoiceDiscAmt.TabIndex = 7
        Me.lblInvoiceDiscAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblInvoiceDiscAmt.Visible = False
        '
        'MyLabel6
        '
        Me.MyLabel6.FieldName = Nothing
        Me.MyLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(483, 155)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(138, 16)
        Me.MyLabel6.TabIndex = 150
        Me.MyLabel6.Text = "- Invoice Discount Amount"
        Me.MyLabel6.Visible = False
        '
        'RadLabel32
        '
        Me.RadLabel32.FieldName = Nothing
        Me.RadLabel32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel32.Location = New System.Drawing.Point(485, 268)
        Me.RadLabel32.Name = "RadLabel32"
        Me.RadLabel32.Size = New System.Drawing.Size(140, 16)
        Me.RadLabel32.TabIndex = 135
        Me.RadLabel32.Text = "+ Total Additional Charges"
        Me.RadLabel32.Visible = False
        '
        'lblAddCharges1
        '
        Me.lblAddCharges1.AutoSize = False
        Me.lblAddCharges1.BorderVisible = True
        Me.lblAddCharges1.FieldName = Nothing
        Me.lblAddCharges1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges1.Location = New System.Drawing.Point(628, 268)
        Me.lblAddCharges1.Name = "lblAddCharges1"
        Me.lblAddCharges1.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges1.TabIndex = 11
        Me.lblAddCharges1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblAddCharges1.Visible = False
        '
        'RadLabel9
        '
        Me.RadLabel9.FieldName = Nothing
        Me.RadLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel9.Location = New System.Drawing.Point(501, 211)
        Me.RadLabel9.Name = "RadLabel9"
        Me.RadLabel9.Size = New System.Drawing.Size(120, 16)
        Me.RadLabel9.TabIndex = 120
        Me.RadLabel9.Text = "Amount After Discount"
        Me.RadLabel9.Visible = False
        '
        'RadLabel27
        '
        Me.RadLabel27.FieldName = Nothing
        Me.RadLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel27.Location = New System.Drawing.Point(522, 295)
        Me.RadLabel27.Name = "RadLabel27"
        Me.RadLabel27.Size = New System.Drawing.Size(100, 16)
        Me.RadLabel27.TabIndex = 123
        Me.RadLabel27.Text = "Document Amount"
        Me.RadLabel27.Visible = False
        '
        'lblTotRAmt
        '
        Me.lblTotRAmt.AutoSize = False
        Me.lblTotRAmt.BorderVisible = True
        Me.lblTotRAmt.FieldName = Nothing
        Me.lblTotRAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt.Location = New System.Drawing.Point(628, 293)
        Me.lblTotRAmt.Name = "lblTotRAmt"
        Me.lblTotRAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotRAmt.TabIndex = 12
        Me.lblTotRAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotRAmt.Visible = False
        '
        'RadLabel25
        '
        Me.RadLabel25.FieldName = Nothing
        Me.RadLabel25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel25.Location = New System.Drawing.Point(545, 241)
        Me.RadLabel25.Name = "RadLabel25"
        Me.RadLabel25.Size = New System.Drawing.Size(77, 16)
        Me.RadLabel25.TabIndex = 122
        Me.RadLabel25.Text = "+ Tax Amount"
        Me.RadLabel25.Visible = False
        '
        'lblTaxAmt
        '
        Me.lblTaxAmt.AutoSize = False
        Me.lblTaxAmt.BorderVisible = True
        Me.lblTaxAmt.FieldName = Nothing
        Me.lblTaxAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxAmt.Location = New System.Drawing.Point(628, 241)
        Me.lblTaxAmt.Name = "lblTaxAmt"
        Me.lblTaxAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTaxAmt.TabIndex = 10
        Me.lblTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTaxAmt.Visible = False
        '
        'lblAmtAfterDiscount
        '
        Me.lblAmtAfterDiscount.AutoSize = False
        Me.lblAmtAfterDiscount.BorderVisible = True
        Me.lblAmtAfterDiscount.FieldName = Nothing
        Me.lblAmtAfterDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtAfterDiscount.Location = New System.Drawing.Point(628, 211)
        Me.lblAmtAfterDiscount.Name = "lblAmtAfterDiscount"
        Me.lblAmtAfterDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtAfterDiscount.TabIndex = 9
        Me.lblAmtAfterDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblAmtAfterDiscount.Visible = False
        '
        'lblDiscountAmt
        '
        Me.lblDiscountAmt.AutoSize = False
        Me.lblDiscountAmt.BorderVisible = True
        Me.lblDiscountAmt.FieldName = Nothing
        Me.lblDiscountAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiscountAmt.Location = New System.Drawing.Point(628, 181)
        Me.lblDiscountAmt.Name = "lblDiscountAmt"
        Me.lblDiscountAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblDiscountAmt.TabIndex = 8
        Me.lblDiscountAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblDiscountAmt.Visible = False
        '
        'lblAmtWithDiscount
        '
        Me.lblAmtWithDiscount.AutoSize = False
        Me.lblAmtWithDiscount.BorderVisible = True
        Me.lblAmtWithDiscount.FieldName = Nothing
        Me.lblAmtWithDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtWithDiscount.Location = New System.Drawing.Point(640, 94)
        Me.lblAmtWithDiscount.Name = "lblAmtWithDiscount"
        Me.lblAmtWithDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtWithDiscount.TabIndex = 1
        Me.lblAmtWithDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.lblAmtWithDiscount.Visible = False
        '
        'RadLabel22
        '
        Me.RadLabel22.FieldName = Nothing
        Me.RadLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel22.Location = New System.Drawing.Point(521, 181)
        Me.RadLabel22.Name = "RadLabel22"
        Me.RadLabel22.Size = New System.Drawing.Size(99, 16)
        Me.RadLabel22.TabIndex = 121
        Me.RadLabel22.Text = "- Discount Amount"
        Me.RadLabel22.Visible = False
        '
        'RadLabel19
        '
        Me.RadLabel19.FieldName = Nothing
        Me.RadLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel19.Location = New System.Drawing.Point(436, 94)
        Me.RadLabel19.Name = "RadLabel19"
        Me.RadLabel19.Size = New System.Drawing.Size(186, 16)
        Me.RadLabel19.TabIndex = 119
        Me.RadLabel19.Text = "Document Amount without Discount"
        Me.RadLabel19.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(731, 9)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(57, 22)
        Me.btnClose.TabIndex = 164
        Me.btnClose.Text = "Close"
        '
        'RadSplitButton1
        '
        Me.RadSplitButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadSplitButton1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.btnImport, Me.btnExport})
        Me.RadSplitButton1.Location = New System.Drawing.Point(472, 11)
        Me.RadSplitButton1.Name = "RadSplitButton1"
        Me.RadSplitButton1.Size = New System.Drawing.Size(122, 22)
        Me.RadSplitButton1.TabIndex = 163
        Me.RadSplitButton1.Text = "Import/Export"
        '
        'btnImport
        '
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Text = "Import"
        Me.btnImport.UseCompatibleTextRendering = False
        '
        'btnExport
        '
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Text = "Export"
        Me.btnExport.UseCompatibleTextRendering = False
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(191, 11)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(57, 22)
        Me.btnPrint.TabIndex = 162
        Me.btnPrint.Text = "Print"
        '
        'btnReverse
        '
        Me.btnReverse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReverse.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReverse.Location = New System.Drawing.Point(347, 11)
        Me.btnReverse.Name = "btnReverse"
        Me.btnReverse.Size = New System.Drawing.Size(122, 22)
        Me.btnReverse.TabIndex = 161
        Me.btnReverse.Text = "Revese and Unpost"
        Me.btnReverse.Visible = False
        '
        'btnLayout
        '
        Me.btnLayout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLayout.Items.AddRange(New Telerik.WinControls.RadItem() {Me.rmiImport, Me.rmiExport})
        Me.btnLayout.Location = New System.Drawing.Point(251, 11)
        Me.btnLayout.Name = "btnLayout"
        Me.btnLayout.Size = New System.Drawing.Size(93, 22)
        Me.btnLayout.TabIndex = 160
        Me.btnLayout.Text = "Layout"
        '
        'rmiImport
        '
        Me.rmiImport.Name = "rmiImport"
        Me.rmiImport.Text = "Save Layout"
        Me.rmiImport.UseCompatibleTextRendering = False
        '
        'rmiExport
        '
        Me.rmiExport.Name = "rmiExport"
        Me.rmiExport.Text = "Delete Layout"
        Me.rmiExport.UseCompatibleTextRendering = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(71, 11)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(57, 22)
        Me.btnDelete.TabIndex = 22
        Me.btnDelete.Text = "Delete"
        '
        'btnPost
        '
        Me.btnPost.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPost.Location = New System.Drawing.Point(131, 11)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(57, 22)
        Me.btnPost.TabIndex = 3
        Me.btnPost.Text = "Post"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(11, 11)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(57, 22)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Image = Global.XpertERPBulkProcurement.My.Resources.Resources._new
        Me.btnAddNew.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnAddNew.Location = New System.Drawing.Point(376, 1)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(20, 22)
        Me.btnAddNew.TabIndex = 1
        '
        'BMC_Transporter_Bill
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "BMC_Transporter_Bill"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "BMC_Transporter_Bill"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.RadPageViewPage1.ResumeLayout(False)
        Me.RadPageViewPage1.PerformLayout()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRemarks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtComment, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox3.ResumeLayout(False)
        Me.RadGroupBox3.PerformLayout()
        CType(Me.txtFatShortageNMG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel28, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TXTSNFShortageNMG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtFatRateNMG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel30, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtSnfRateNMG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel35, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFatShortage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFatShortage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtSnfShortage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTankerprorata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtIceCharge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtKMRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtBarelCap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTollTax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtFatRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtDieselMinus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtSnfRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDieselplus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtRateprorata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTDR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDCS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTransporter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTransporter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTankerDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAbandonmentNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTankerNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage2.ResumeLayout(False)
        Me.RadPageViewPage2.PerformLayout()
        CType(Me.NACCOEPF_PER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NEMPEPF_PER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NCOESI_PER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NEMPESI_PER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NTotalPF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NTotalESI, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NESIEmployer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NESIEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NEDLICharge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NAdmincharge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NemployerPF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NempPF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NLabourCharge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NDailyWage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel70, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel69, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel68, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel67, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel66, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel65, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel61, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel60, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel59, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel58, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel57, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel56, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel55, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel54, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel53, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel52, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel51, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel50, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel49, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel48, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel37, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel36, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage4.ResumeLayout(False)
        Me.RadPageViewPage4.PerformLayout()
        CType(Me.TotalESIAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TotalEPFAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel64, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel38, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalDeduction, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalAddition, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel39, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel63, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel47, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel46, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel45, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel44, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel43, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel42, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel41, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel40, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalIceCharge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalFatSnfShortage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtGrossAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtTotalTollTax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel34, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel33, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtBMCTotal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtBMCDiesel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtBMCProrataamt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel62, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel24, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel20, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblInvoiceDiscAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotRAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAmtAfterDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDiscountAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAmtWithDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnLayout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents RadPageView1 As RadPageView
    Friend WithEvents RadPageViewPage1 As RadPageViewPage
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents RadLabel4 As common.Controls.MyLabel
    Friend WithEvents chkDCS As RadCheckBox
    Friend WithEvents lblTankerDesc As common.Controls.MyLabel
    Friend WithEvents lblAbandonmentNo As common.Controls.MyLabel
    Friend WithEvents lblTankerNo As common.Controls.MyLabel
    Friend WithEvents RadGroupBox2 As RadGroupBox
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents txtTankerNo As common.UserControls.txtFinder
    Friend WithEvents RadLabel1 As common.Controls.MyLabel
    Friend WithEvents txtDocNo As common.UserControls.txtNavigator
    Friend WithEvents btnAddNew As RadButton
    Friend WithEvents RadPageViewPage4 As RadPageViewPage
    Friend WithEvents lblInvoiceDiscAmt As common.Controls.MyLabel
    Friend WithEvents MyLabel6 As common.Controls.MyLabel
    Friend WithEvents RadLabel32 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges1 As common.Controls.MyLabel
    Friend WithEvents RadLabel9 As common.Controls.MyLabel
    Friend WithEvents RadLabel27 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt As common.Controls.MyLabel
    Friend WithEvents RadLabel25 As common.Controls.MyLabel
    Friend WithEvents lblTaxAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtAfterDiscount As common.Controls.MyLabel
    Friend WithEvents lblDiscountAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtWithDiscount As common.Controls.MyLabel
    Friend WithEvents RadLabel22 As common.Controls.MyLabel
    Friend WithEvents RadLabel19 As common.Controls.MyLabel
    Friend WithEvents txtTransporter As common.Controls.MyTextBox
    Friend WithEvents lblTransporter As common.Controls.MyLabel
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents MyLabel3 As common.Controls.MyLabel
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents lblFatShortage As common.Controls.MyLabel
    Friend WithEvents TxtKMRate As common.Controls.MyTextBox
    Friend WithEvents MyLabel8 As common.Controls.MyLabel
    Friend WithEvents TxtTollTax As common.Controls.MyTextBox
    Friend WithEvents TxtDieselMinus As common.Controls.MyTextBox
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents txtDieselplus As common.Controls.MyTextBox
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents TxtSnfRate As common.Controls.MyTextBox
    Friend WithEvents TxtFatRate As common.Controls.MyTextBox
    Friend WithEvents TxtSnfShortage As common.Controls.MyTextBox
    Friend WithEvents txtFatShortage As common.Controls.MyTextBox
    Friend WithEvents TxtTDR As common.Controls.MyTextBox
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents TxtTankerprorata As common.Controls.MyTextBox
    Friend WithEvents MyLabel11 As common.Controls.MyLabel
    Friend WithEvents TxtBarelCap As common.Controls.MyTextBox
    Friend WithEvents MyLabel12 As common.Controls.MyLabel
    Friend WithEvents TxtIceCharge As common.Controls.MyTextBox
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents MyLabel14 As common.Controls.MyLabel
    Friend WithEvents MyLabel17 As common.Controls.MyLabel
    Friend WithEvents MyLabel16 As common.Controls.MyLabel
    Friend WithEvents MyLabel15 As common.Controls.MyLabel
    Friend WithEvents MyLabel21 As common.Controls.MyLabel
    Friend WithEvents MyLabel20 As common.Controls.MyLabel
    Friend WithEvents MyLabel19 As common.Controls.MyLabel
    Friend WithEvents MyLabel18 As common.Controls.MyLabel
    Friend WithEvents MyLabel23 As common.Controls.MyLabel
    Friend WithEvents MyLabel22 As common.Controls.MyLabel
    Friend WithEvents MyLabel24 As common.Controls.MyLabel
    Friend WithEvents MyLabel25 As common.Controls.MyLabel
    Friend WithEvents MyLabel26 As common.Controls.MyLabel
    Friend WithEvents MyLabel27 As common.Controls.MyLabel
    Friend WithEvents MyLabel47 As common.Controls.MyLabel
    Friend WithEvents MyLabel46 As common.Controls.MyLabel
    Friend WithEvents MyLabel45 As common.Controls.MyLabel
    Friend WithEvents MyLabel44 As common.Controls.MyLabel
    Friend WithEvents MyLabel43 As common.Controls.MyLabel
    Friend WithEvents MyLabel42 As common.Controls.MyLabel
    Friend WithEvents MyLabel41 As common.Controls.MyLabel
    Friend WithEvents MyLabel40 As common.Controls.MyLabel
    Friend WithEvents TxtTotalIceCharge As common.Controls.MyLabel
    Friend WithEvents TxtTotalAmount As common.Controls.MyLabel
    Friend WithEvents TxtTotalFatSnfShortage As common.Controls.MyLabel
    Friend WithEvents TxtGrossAmount As common.Controls.MyLabel
    Friend WithEvents TxtTotalTollTax As common.Controls.MyLabel
    Friend WithEvents MyLabel34 As common.Controls.MyLabel
    Friend WithEvents MyLabel33 As common.Controls.MyLabel
    Friend WithEvents MyLabel32 As common.Controls.MyLabel
    Friend WithEvents MyLabel31 As common.Controls.MyLabel
    Friend WithEvents TxtBMCTotal As common.Controls.MyLabel
    Friend WithEvents TxtBMCDiesel As common.Controls.MyLabel
    Friend WithEvents TxtBMCProrataamt As common.Controls.MyLabel
    Friend WithEvents RadGroupBox1 As RadGroupBox
    Friend WithEvents txtToDate As common.Controls.MyDateTimePicker
    Friend WithEvents txtFromDate As common.Controls.MyDateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnPost As RadButton
    Friend WithEvents btnSave As RadButton
    Friend WithEvents btnGo As RadButton
    Friend WithEvents btnDelete As RadButton
    Friend WithEvents btnLayout As RadSplitButton
    Friend WithEvents rmiImport As RadMenuItem
    Friend WithEvents rmiExport As RadMenuItem
    Friend WithEvents RadGroupBox3 As RadGroupBox
    Friend WithEvents UsLock1 As common.usLock
    Friend WithEvents btnReverse As RadButton
    Friend WithEvents btnPrint As RadButton
    Friend WithEvents RadSplitButton1 As RadSplitButton
    Friend WithEvents btnImport As RadMenuItem
    Friend WithEvents btnExport As RadMenuItem
    Friend WithEvents chkPrivate As CheckBox
    Friend WithEvents RadLabel6 As common.Controls.MyLabel
    Friend WithEvents RadLabel3 As common.Controls.MyLabel
    Friend WithEvents txtRemarks As common.Controls.MyTextBox
    Friend WithEvents txtComment As common.Controls.MyTextBox
    Friend WithEvents btnClose As RadButton
    Friend WithEvents txtFatShortageNMG As common.Controls.MyTextBox
    Friend WithEvents MyLabel28 As common.Controls.MyLabel
    Friend WithEvents TXTSNFShortageNMG As common.Controls.MyTextBox
    Friend WithEvents MyLabel29 As common.Controls.MyLabel
    Friend WithEvents TxtFatRateNMG As common.Controls.MyTextBox
    Friend WithEvents MyLabel30 As common.Controls.MyLabel
    Friend WithEvents TxtSnfRateNMG As common.Controls.MyTextBox
    Friend WithEvents MyLabel35 As common.Controls.MyLabel
    Friend WithEvents TxtRateprorata As common.Controls.MyTextBox
    Friend WithEvents MyLabel9 As common.Controls.MyLabel
    Friend WithEvents MyLabel39 As common.Controls.MyLabel
    Friend WithEvents TxtTotalAddition As common.Controls.MyTextBox
    Friend WithEvents TxtTotalDeduction As common.Controls.MyTextBox
    Friend WithEvents RadPageViewPage2 As RadPageViewPage
    Friend WithEvents MyLabel61 As common.Controls.MyLabel
    Friend WithEvents MyLabel60 As common.Controls.MyLabel
    Friend WithEvents MyLabel59 As common.Controls.MyLabel
    Friend WithEvents MyLabel58 As common.Controls.MyLabel
    Friend WithEvents MyLabel57 As common.Controls.MyLabel
    Friend WithEvents MyLabel56 As common.Controls.MyLabel
    Friend WithEvents MyLabel55 As common.Controls.MyLabel
    Friend WithEvents MyLabel54 As common.Controls.MyLabel
    Friend WithEvents MyLabel53 As common.Controls.MyLabel
    Friend WithEvents MyLabel52 As common.Controls.MyLabel
    Friend WithEvents MyLabel51 As common.Controls.MyLabel
    Friend WithEvents MyLabel50 As common.Controls.MyLabel
    Friend WithEvents MyLabel49 As common.Controls.MyLabel
    Friend WithEvents MyLabel48 As common.Controls.MyLabel
    Friend WithEvents MyLabel37 As common.Controls.MyLabel
    Friend WithEvents MyLabel36 As common.Controls.MyLabel
    Friend WithEvents MyLabel64 As common.Controls.MyLabel
    Friend WithEvents MyLabel38 As common.Controls.MyLabel
    Friend WithEvents MyLabel63 As common.Controls.MyLabel
    Friend WithEvents MyLabel62 As common.Controls.MyLabel
    Friend WithEvents MyLabel65 As common.Controls.MyLabel
    Friend WithEvents MyLabel69 As common.Controls.MyLabel
    Friend WithEvents MyLabel68 As common.Controls.MyLabel
    Friend WithEvents MyLabel67 As common.Controls.MyLabel
    Friend WithEvents MyLabel66 As common.Controls.MyLabel
    Friend WithEvents MyLabel70 As common.Controls.MyLabel
    Friend WithEvents NDailyWage As common.Controls.MyLabel
    Friend WithEvents NACCOEPF_PER As common.Controls.MyLabel
    Friend WithEvents NEMPEPF_PER As common.Controls.MyLabel
    Friend WithEvents NCOESI_PER As common.Controls.MyLabel
    Friend WithEvents NEMPESI_PER As common.Controls.MyLabel
    Friend WithEvents NTotalPF As common.Controls.MyLabel
    Friend WithEvents NTotalESI As common.Controls.MyLabel
    Friend WithEvents NESIEmployer As common.Controls.MyLabel
    Friend WithEvents NESIEmployee As common.Controls.MyLabel
    Friend WithEvents NEDLICharge As common.Controls.MyLabel
    Friend WithEvents NAdmincharge As common.Controls.MyLabel
    Friend WithEvents NemployerPF As common.Controls.MyLabel
    Friend WithEvents NempPF As common.Controls.MyLabel
    Friend WithEvents NLabourCharge As common.Controls.MyLabel
    Friend WithEvents TotalESIAmt As common.Controls.MyLabel
    Friend WithEvents TotalEPFAmt As common.Controls.MyLabel
End Class
