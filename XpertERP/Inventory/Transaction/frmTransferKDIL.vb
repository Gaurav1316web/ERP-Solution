'--Created By -- [Pankaj Kumar Chaudhary], Date-  [31/12/2014]
''UPDATION BY RICHA AGARWAL AGAINST TICKET NO. BM00000005586,BM00000006863
'' working on Print format and Print data function agasist ticket no. UDL/15/05/18-000162
Imports common
Imports System
Imports XpertERPEngine
Public Class FrmTransferKDIL
    Inherits FrmMainTranScreen
    Dim Qry As String
    Dim dt As DataTable
#Region "Variables"
    'done by stuti on 05/10/2016 against ticket no BM00000009942
    Dim blnUpdateLoadInwithLoadOut As Boolean
    Public JobWorkTransfer As Boolean = False
    Dim GSTOn As Integer = 0
    Public EnableInternalTransfer As Boolean = False
    Dim strTaxType As String = Nothing
    Dim IsMandiTax As Boolean = False
    Dim GstStatus As Boolean = False
    Dim RunBatchFifowise As Integer = 0
    Dim RunBatchFifowisewithmodifyfunctionality As Boolean = False
    Dim ApplyFEFO As Boolean = False
    Public AllowThreeFormatByDefault As Boolean = False
    Public AllowOneFormatByDefault As Boolean = False 'Added by preeti gupta Against tciket no[MIL/06/05/19-000079,MIL/02/05/19-000077]
    Dim StrCrateTransferFromBooking As String = ""
    Dim ProvisionAllow As Boolean = False
    Dim Weight_MT_Unit As String = ""
    Dim Gross_Weight_Unit As String = ""
    Private isCellValueChangedTaxOpen As Boolean = False
    Public Const RowTypeItem As String = "Item"
    Public Const RowTypeMisc As String = "Misc"
    Private isCellValueChangedOpen As Boolean = False
    Private IsFormLoad As Boolean = False
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Const colAltUOM As String = "AltUOM"
    Const colINetMTWt As String = "INetMTWt"
    Const colINetWt As String = "INetWt"
    Const colIUnitWt As String = "IUnitWt"
    Const colLineNo As String = "COLLNO"
    Const colRowType As String = "COLTYPE"
    Const colRGPNo As String = "COLRGPNo"
    Const colComplete As String = "COMPLETE"
    Const colICode As String = "COLICODE"
    Const colgatePassTransferNo As String = "COLGPTNo"
    Const colIName As String = "COLINAME"
    Const colIHSN As String = "colIHSN"
    Const colOutQty As String = "COLOutQty"
    Const colInQty As String = "COLInQTy"
    Const colBreakQty As String = "COLBreakQty"
    Const colLeakQty As String = "COLLeakQty"
    Const colShortQty As String = "COLShortQty"
    Const colUnit As String = "COLUNIT"
    Const colRate As String = "COLRATE"
    Const colAmt As String = "COLAMT"
    Const colAbatementPer As String = "ABATEMENT%"
    Const colAbatementAmount As String = "ABATEMENTAMT"
    Const colDisPer As String = "COLDISPER"
    Const colDisAmt As String = "COLDISAMT"
    Const colAmtAfterDis As String = "COLAMTAFTERDIS"
    Public strTransferno As String = ""
    Const colBinNo As String = "colBinNo"
    Const colTax1 As String = "COLTAX1"
    Const colTaxBaseAmt1 As String = "COLTAXBASEAMT1"
    Const colTaxRate1 As String = "COLTAXRATE1"
    Const colTaxAmt1 As String = "COLTAXAMT1"
    Const colIsTaxable1 As String = "ISTAXABLE1"
    Const colIsSurTax1 As String = "ISSURTAX1"
    Const colSurTaxCode1 As String = "SURTAXCODE1"
    Const colIsExcisable1 As String = "ISEXCISABLE1"
    Const colTax2 As String = "COLTAX2"
    Const colTaxBaseAmt2 As String = "COLTAXBASEAMT2"
    Const colTaxRate2 As String = "COLTAXRATE2"
    Const colTaxAmt2 As String = "COLTAXAMT2"
    Const colIsTaxable2 As String = "ISTAXABLE2"
    Const colIsSurTax2 As String = "ISSURTAX2"
    Const colSurTaxCode2 As String = "SURTAXCODE2"
    Const colIsExcisable2 As String = "ISEXCISABLE2"
    Const colTax3 As String = "COLTAX3"
    Const colTaxBaseAmt3 As String = "COLTAXBASEAMT3"
    Const colTaxRate3 As String = "COLTAXRATE3"
    Const colTaxAmt3 As String = "COLTAXAMT3"
    Const colIsTaxable3 As String = "ISTAXABLE3"
    Const colIsSurTax3 As String = "ISSURTAX3"
    Const colSurTaxCode3 As String = "SURTAXCODE3"
    Const colIsExcisable3 As String = "ISEXCISABLE3"
    Const colTax4 As String = "COLTAX4"
    Const colTaxBaseAmt4 As String = "COLTAXBASEAMT4"
    Const colTaxRate4 As String = "COLTAXRATE4"
    Const colTaxAmt4 As String = "COLTAXAMT4"
    Const colIsTaxable4 As String = "ISTAXABLE4"
    Const colIsSurTax4 As String = "ISSURTAX4"
    Const colSurTaxCode4 As String = "SURTAXCODE4"
    Const colIsExcisable4 As String = "ISEXCISABLE4"
    Const colTax5 As String = "COLTAX5"
    Const colTaxBaseAmt5 As String = "COLTAXBASEAMT5"
    Const colTaxRate5 As String = "COLTAXRATE5"
    Const colTaxAmt5 As String = "COLTAXAMT5"
    Const colIsTaxable5 As String = "ISTAXABLE5"
    Const colIsSurTax5 As String = "ISSURTAX5"
    Const colSurTaxCode5 As String = "SURTAXCODE5"
    Const colIsExcisable5 As String = "ISEXCISABLE5"
    Const colTax6 As String = "COLTAX6"
    Const colTaxBaseAmt6 As String = "COLTAXBASEAMT6"
    Const colTaxRate6 As String = "COLTAXRATE6"
    Const colTaxAmt6 As String = "COLTAXAMT6"
    Const colIsTaxable6 As String = "ISTAXABLE6"
    Const colIsSurTax6 As String = "ISSURTAX6"
    Const colSurTaxCode6 As String = "SURTAXCODE6"
    Const colIsExcisable6 As String = "ISEXCISABLE6"
    Const colTax7 As String = "COLTAX7"
    Const colTaxBaseAmt7 As String = "COLTAXBASEAMT7"
    Const colTaxRate7 As String = "COLTAXRATE7"
    Const colTaxAmt7 As String = "COLTAXAMT7"
    Const colIsTaxable7 As String = "ISTAXABLE7"
    Const colIsSurTax7 As String = "ISSURTAX7"
    Const colSurTaxCode7 As String = "SURTAXCODE7"
    Const colIsExcisable7 As String = "ISEXCISABLE7"
    Const colTax8 As String = "COLTAX8"
    Const colTaxBaseAmt8 As String = "COLTAXBASEAMT8"
    Const colTaxRate8 As String = "COLTAXRATE8"
    Const colTaxAmt8 As String = "COLTAXAMT8"
    Const colIsTaxable8 As String = "ISTAXABLE8"
    Const colIsSurTax8 As String = "ISSURTAX8"
    Const colSurTaxCode8 As String = "SURTAXCODE8"
    Const colIsExcisable8 As String = "ISEXCISABLE8"
    Const colTax9 As String = "COLTAX9"
    Const colTaxBaseAmt9 As String = "COLTAXBASEAMT9"
    Const colTaxRate9 As String = "COLTAXRATE9"
    Const colTaxAmt9 As String = "COLTAXAMT9"
    Const colIsTaxable9 As String = "ISTAXABLE9"
    Const colIsSurTax9 As String = "ISSURTAX9"
    Const colSurTaxCode9 As String = "SURTAXCODE9"
    Const colIsExcisable9 As String = "ISEXCISABLE9"
    Const colTax10 As String = "COLTAX10"
    Const colTaxBaseAmt10 As String = "COLTAXBASEAMT10"
    Const colTaxRate10 As String = "COLTAXRATE10"
    Const colTaxAmt10 As String = "COLTAXAMT10"
    Const colIsTaxable10 As String = "ISTAXABLE10"
    Const colIsSurTax10 As String = "ISSURTAX10"
    Const colSurTaxCode10 As String = "SURTAXCODE10"
    Const colIsExcisable10 As String = "ISEXCISABLE10"
    Const colIsBatchItem As String = "colIsBatchItem"
    Const colFOCItem As String = "colFOCItem"
    Const colTotTaxAmt As String = "TAXAMT"
    Const colAmtAfterTax As String = "AMTAFTERTAX"
    Const colTransferOutNo As String = "ColTransferOutNo"
    Const colSpecification As String = "SPECIFICATION"
    Const colRemarks As String = "REMARKS"
    Const colItemUsedINGRN As String = "USEDINGRN"
    Const colisMRPMandatory As String = "colisMRPMandatory"
    Const colMRP As String = "MRP"
    'Const colAssessableRate As String = "ASSESSABLERATE"
    Const colAssessableAmount As String = "ASSESSABLEAMT"
    Const colIsSerialseItem As String = "COLISSERIALSEITEM"
    Const colIsPickAutoSrNo As String = "colIsPickAutoSrNo"
    ''Const colLocationCode As String = "LOCATIONCODE"
    ''Const colLocationName As String = "LOCATIONNAME"
    Const colTTaxAutCode As String = "TAXAUTCODE"
    Const colTTaxAutName As String = "TAXAUTNAME"
    'Const colTIsTaxable As String = "ISTAXABLE"
    Const colTTaxRate As String = "TAXRATE"
    'Const colTIsSurTax As String = "ISSURTAX"
    'Const colTSurTaxCode As String = "SURTAXCODE"
    Const colTBaseAmt As String = "TAXBASEAMT"
    Const colTTaxAmt As String = "TAXAMT"
    Const colTTaxAssessableAmt As String = "COLTTAXASSESSABLEAMT"
    Const colACCode As String = "COLACCODE"
    Const colACName As String = "COLACNAME"
    Const colACAmount As String = "COLACAMOUNT"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim i As Integer
    Dim repoBalQty As GridViewDecimalColumn
    Dim repoComplete As GridViewTextBoxColumn
    Dim IsItemRateeditable As Boolean = False
    Dim IsRateeditableItemwise As Boolean = False
    Dim Item_TaxType As Integer = 0
    Dim arrLoc As String = Nothing
    Private isALlowVehicleGateOutValidation As Boolean = False ' Add By Prabhakar
    Dim ExciseSecondaryCode As Boolean = False
    Dim GrossWtfromItemMaster As Boolean = False
    Dim ValidateTaxGroup As Boolean = False
    Dim PickRateFromPRICEChrtFORUMang As Boolean = False
    Dim settStockTranferFromTransferPriceAndInvJVWithAvgCost As Boolean = False
    Dim PreviousTaxGroupCode As String = Nothing
    Dim CheckRepair As Boolean = False
    Dim settPickProductCostFromItemUOMDetail As Boolean = False
    Dim ItemStructureMandatoryOnWeightConversion As Boolean = False
    Dim InDocMandatoryOnInternalTransfer As Boolean = False
    Dim CalculateTaxRatefromItemwsieTaxOnSale As Integer = 0
    Const colItemwiseTaxCode As String = "colItemwiseTaxCode"
    Dim EInvoiceType As String = ""
    Dim AllowOnlyOneIssueAgainstStoreRequisition As Boolean = False
    Dim ShowPriceCode As Boolean = False
#End Region
    Private Sub LOCATIONRIGTHS()
        Try
            Dim obj As New clsMCCCodes()
            obj = clsMCCCodes.GetData()
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 0 Then
                Dim check As Integer = 0
                check = clsDBFuncationality.getSingleValue("select count(*) from tspl_location_master where type='Physical' and location_code='" + obj.Default_LocCode + "'")
                If check > 0 Then
                    txtFromLocation.Value = obj.Default_LocCode
                    lblFromLoc.Text = obj.Default_LocName
                End If
            End If
            If obj.arrLocCodes IsNot Nothing AndAlso clsCommon.myLen(obj.arrLocCodes) > 0 Then
                arrLoc = obj.arrLocCodes
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.Transfer)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPrintNew.Visible = MyBase.isPrintFlag
        btnCancel.Visible = MyBase.isCancel_Flag
        btnReverseAndUnpost.Visible = False
        'If MyBase.isReverse Then
        '    btnReverseAndUnpost.Enabled = True
        'Else
        '    btnReverseAndUnpost.Enabled = False
        'End If
        If MyBase.isExport = True Then
            radExportTransferOut.Enabled = True
            radExportTransferIn.Enabled = True
            radImportTransferOut.Enabled = True
            radImportTransferIn.Enabled = True
        Else
            radExportTransferOut.Enabled = False
            radExportTransferIn.Enabled = False
            radImportTransferOut.Enabled = False
            radImportTransferIn.Enabled = False
        End If
        btn_CancelDel.Visible = MyBase.isCancel_Flag_After_Posting
    End Sub
    Private Sub FrmTransferKDIL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetUserMgmtNew()
        settPickProductCostFromItemUOMDetail = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickProductCostFromItemUOMDetail, clsFixedParameterCode.PickProductCostFromItemUOMDetail, Nothing)) = 1)
        RunBatchFifowise = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, Nothing))
        RunBatchFifowisewithmodifyfunctionality = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, Nothing)) = 1, True, False)
        AllowThreeFormatByDefault = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowThreeFormatByDefaultForPrint, clsFixedParameterCode.AllowThreeFormatByDefaultForPrint, Nothing)) = 1, True, False)
        AllowOneFormatByDefault = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowOneFormatByDefaultForPrint, clsFixedParameterCode.AllowOneFormatByDefaultForPrint, Nothing)) = 1, True, False)
        GrossWtfromItemMaster = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GrossWtFromItemMasterONCSATransfer, clsFixedParameterCode.GrossWtFromItemMasterONCSATransfer, Nothing)) = 1, True, False)
        Weight_MT_Unit = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select description from TSPL_FIXED_PARAMETER where code='" + clsFixedParameterCode.VehicleCapacityUnit + "' and type='" + clsFixedParameterType.VehicleCapacityUnit + "'"))
        Gross_Weight_Unit = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select description from TSPL_FIXED_PARAMETER where code='" + clsFixedParameterCode.GrossWeightUnit + "' and type='" + clsFixedParameterType.GrossWeightUnit + "'"))
        ProvisionAllow = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PROVISIONENTRYONSTOCKTRANSFER, clsFixedParameterCode.PROVISIONENTRYONSTOCKTRANSFER, Nothing)) = 1, True, False))
        ValidateTaxGroup = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ValidateTaxGroupForTransaction, clsFixedParameterCode.ValidateTaxGroupForTransaction, Nothing)) = 1, True, False))
        PickRateFromPRICEChrtFORUMang = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PickRateFromPRICEChrtMasterFORUMang, clsFixedParameterCode.PickRateFromPRICEChrtMasterFORUMang, Nothing)) = 1, True, False))
        settStockTranferFromTransferPriceAndInvJVWithAvgCost = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.StockTranferFromTransferPriceAndInvJVWithAvgCost, clsFixedParameterCode.StockTranferFromTransferPriceAndInvJVWithAvgCost, Nothing)) = 1)
        '' done by Panch Raj: if setting=0 then rate is no editable for all items, if setting=1 the editable for all items , if setting=2 then rate will be editable itemwise
        ApplyFEFO = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ApplyFEFO, clsFixedParameterCode.ApplyFEFO, Nothing)) = 1, True, False))
        ItemStructureMandatoryOnWeightConversion = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ItemStructureMandatoryOnWeightConversion, clsFixedParameterCode.ItemStructureMandatoryOnWeightConversion, Nothing)) = 1, True, False))
        InDocMandatoryOnInternalTransfer = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.InDocMandatoryOnInternalTransfer, clsFixedParameterCode.InDocMandatoryOnInternalTransfer, Nothing)) = 1, True, False))
        '==(Sanjeet)17/11/2016====
        ExciseSecondaryCode = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.GenerateSecondryCode, clsFixedParameterCode.GenerateSecondryCode, Nothing)) = 1, True, False))
        MyLabel24.Visible = ExciseSecondaryCode
        txtSecondary_Doc_Code.Visible = ExciseSecondaryCode
        '=====
        CalculateTaxRatefromItemwsieTaxOnSale = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, Nothing))
        AllowOnlyOneIssueAgainstStoreRequisition = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowOnlyOneIssueAgainstStoreRequisition, clsFixedParameterCode.AllowOnlyOneIssueAgainstStoreRequisition, Nothing)) = 1, True, False)
        Dim Value As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.IsItemRateEditableOnTransfer & "'"))
        ShowPriceCode = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DisplayPriceCodeForBKN, clsFixedParameterCode.DisplayPriceCodeForBKN, Nothing)) = 1, True, False)
        If ShowPriceCode Then
            fndPriceCode.Visible = True
            lblpricecode.Visible = True
        Else
            fndPriceCode.Visible = False
            lblpricecode.Visible = False
        End If
        If Value = 0 Then
            IsItemRateeditable = False
            IsRateeditableItemwise = False
        ElseIf Value = 1 Then
            IsItemRateeditable = True
            IsRateeditableItemwise = False
        ElseIf Value = 2 Then
            IsItemRateeditable = True
            IsRateeditableItemwise = True
        End If
        EnableInternalTransfer = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.EnableInternalTransfer & "'")) = 0, False, True)
        If clsCommon.myCBool(EnableInternalTransfer) = True Then
            chkInternalTransfer.Visible = True
            lblSRNO.Visible = True
            fndSRNO.Visible = True
            chkProductionRequest.Visible = True
        Else
            lblSRNO.Visible = False
            fndSRNO.Visible = False
        End If
        LoadModeOfTrasport()
        JobWorkTransfer = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.JobWorkTransfer, clsFixedParameterCode.JobWorkTransfer, Nothing)) = "1", True, False)

        If JobWorkTransfer Then
            chkJobWorkTransfer.Visible = True

        End If
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D Delete Trasnaction")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+N Adding New Trasnaction")
        RadPageView1.SelectedPage = RadPageViewPage1
        LoadBlankGrid()
        LoadBlankGridTax()
        IsFormLoad = True
        LoadType()
        LoadItemType()
        LoadTransferType()
        cboItemType.SelectedIndex = 2
        'ToolStrip1.LayoutStyle = ToolStripLayoutStyle.Table
        SetLength()
        IsFormLoad = False
        ''For Custom Fields
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If
        If objCommonVar.IsDemoERP Then
            UcAttachment1.Form_ID = MyBase.Form_ID
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed
        End If
        ''End of For Attachment
        AddNew()
        If clsCommon.myLen(strTransferno) > 0 Then
            txtDocNo.Value = strTransferno
            LoadData(strTransferno, NavigatorType.Current)
        End If
        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
        End If
        isALlowVehicleGateOutValidation = IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowVehicleGateOutValidationTransfer, clsFixedParameterCode.AllowVehicleGateOutValidationTransfer, Nothing)) = "1", True, False)
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
            rbtnPrintExcisable.Visible = True
        Else
            rbtnPrintExcisable.Visible = False
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Or clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            btnSTAMilkPrint.Visible = True
            btnSTAProductPrint.Visible = True
        Else
            btnSTAMilkPrint.Visible = False
            btnSTAProductPrint.Visible = False
        End If
        radExportTransferIn.Visibility = ElementVisibility.Collapsed
        radImportTransferIn.Visibility = ElementVisibility.Collapsed
    End Sub
    Sub SetLength()
        txtDocNo.MyMaxLength = 30
        txtRemarks.MaxLength = 200
        txtComment.MaxLength = 200
        cboModeOfTransport.MaxLength = 12
        cboTransferType.MaxLength = 1
        cboTransferType.MaxLength = 1
    End Sub
    Sub LoadTransferType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr = dt.NewRow()
        dr("Code") = "O"
        dr("Name") = "Transfer Out"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "I"
        dr("Name") = "Transfer In"
        dt.Rows.Add(dr)
        '============added by preeti gupta against ticket no[UDL/24/07/18-000207]
        dr = dt.NewRow()
        dr("Code") = "T"
        dr("Name") = "Return"
        dt.Rows.Add(dr)
        '==========================================
        dr = dt.NewRow()
        dr("Code") = "R"
        dr("Name") = "Rejected"
        dt.Rows.Add(dr)
        cboTransferType.DataSource = dt
        cboTransferType.ValueMember = "Code"
        cboTransferType.DisplayMember = "Name"
        cboTransferType.SelectedValue = "O"
    End Sub
    Sub LoadType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = ""
        dr("Name") = "Select"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Excise"
        dr("Name") = "Excise"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Depot"
        dr("Name") = "Depot"
        dt.Rows.Add(dr)
        cboType.DataSource = dt
        cboType.ValueMember = "Code"
        cboType.DisplayMember = "Name"
        cboType.SelectedValue = "Depot"
    End Sub
    Sub LoadItemType()
        cboItemType.DataSource = clsItemMaster.GetItemType()
        cboItemType.ValueMember = "Code"
        cboItemType.DisplayMember = "Name"
        cboItemType.SelectedIndex = 2
    End Sub
    Sub LoadModeOfTrasport()
        cboModeOfTransport.Items.Add("By Road")
        cboModeOfTransport.Items.Add("By Air")
        cboModeOfTransport.Items.Add("By Sea")
        cboModeOfTransport.Items.Add("By Hand")
    End Sub
    Sub BlankAllControls()
        'chkProductionRequest.Enabled = False
        'chkProductionRequest.Checked = False
        chkEwaybill.Checked = False
        chkEwaybill.Enabled = True
        txtFreightDistance.Value = 0
        EInvoiceType = ""
        chkForRepair.Enabled = True
        CheckRepair = False
        PreviousTaxGroupCode = ""
        txtglvoucher.Text = ""
        '============Added by preeti gupta =====
        chkIsManditax.Checked = False
        txtElecttefNo.Text = ""
        txtEWayBillNo.Text = ""
        txtEWayBillDate.Checked = False
        txtEWayBillDate.Value = clsCommon.GETSERVERDATE
        '=======================================
        txtEWayBillNo.ReadOnly = False
        txtEWayBillDate.ReadOnly = False
        txtvehicle_mannual_no.Text = ""
        txtSecondary_Doc_Code.Text = Nothing
        txtVehicle_Capacity.Text = Nothing
        txtvehicle_Charge.Tag = Nothing
        txttotal_Wt.Text = Nothing
        txtGross_Wt.Text = Nothing
        txtvehicle_Charge.Text = Nothing
        txtGR_No.Text = ""
        dtpGR.Value = clsCommon.GETSERVERDATE(Nothing)
        txtWayBill_No.Text = ""
        txtTransporter_Code.Value = Nothing
        txtTransporter_desc.Text = Nothing
        cboType.SelectedValue = "Depot"
        txtDocNo.Value = ""
        chkOnHold.Checked = False
        chkForm38.Checked = False
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtOutDate.Value = txtDate.Value
        ttxway_bill_date.Value = txtDate.Value
        txtFromLocation.Value = ""
        lblFromLoc.Text = ""
        txtToLoc.Value = ""
        lblToLoc.Text = ""
        txtRemarks.Text = ""
        txtComment.Text = ""
        txtTaxGroup.Value = ""
        lblTaxGrpName.Text = ""
        txtTermCode.Value = ""
        txtTermRemark.Text = ""
        lblTermName.Text = ""
        txtDueDate.Value = txtDate.Value
        txtRefNo.Text = ""
        lblAmtWithoutTax.Text = ""
        lblDiscountAmt.Text = ""
        lblAmtAfterDiscount.Text = ""
        lblTaxAmt.Text = ""
        lblTotRAmt.Text = ""
        lblTotRAmt1.Text = ""
        cboModeOfTransport.Text = "By Road"
        cboTransferType.SelectedValue = "O"
        'chkAgainst_Form.Enabled = False
        UsLock1.Status = ERPTransactionStatus.Pending
        txtDeliveryDate.Value = txtDate.Value
        txtVehicleCode.Value = ""
        lblVehicleNo.Text = ""
        cboTransferType.Enabled = True
        cboTransferType.SelectedIndex = 0
        txtTransferOutNo.Value = ""
        txtFromLocation.Enabled = True
        txtToLoc.Enabled = True
        lblAbandonmentNo.Text = ""
        btnAmendment.Enabled = False
        lblAddCharges1.Text = ""
        lblAddCharges1.Text = ""
        rbtnTaxCalAutomatic.IsChecked = True
        chkExciseOnQty.Checked = False
        chkExciseOnQty.Enabled = True
        txtKm.Text = ""
        txtRMDANo.Value = ""
        txtDeliveryDuration.Text = ""
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GUNTUR") = CompairStringResult.Equal Then
            lblDeliveryDate.Visible = False
            txtDeliveryDate.Visible = False
        Else
            Panel1.Visible = False
        End If
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.BlankAllControls()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        chkOwnVehicle.Checked = False
        LOCATIONRIGTHS()
        ''richa 03/03/2015
        chkAgainst_Form.Enabled = True
        chkAgainst_Form.Checked = False
        ttxway_bill_date.Enabled = False
        chkForRepair.Checked = False
        '------
        GSTOn = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, Nothing))
        If GSTOn = 1 Then
            If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                cboType.Enabled = False
                cboType.SelectedValue = "Depot"
            Else
                cboType.SelectedValue = ""
                cboType.Enabled = True
            End If
        End If
    End Sub
    Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLineNo)
        repoComplete = New GridViewTextBoxColumn()
        repoComplete.FormatString = ""
        repoComplete.HeaderText = "Complete"
        repoComplete.Width = 70
        repoComplete.Name = colComplete
        repoComplete.ReadOnly = True
        repoComplete.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoComplete)
        'Dim repoRowType As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        'repoRowType.FormatString = ""
        'repoRowType.HeaderText = "Row Type"
        'repoRowType.Name = colRowType
        'repoRowType.Width = 50
        'repoRowType.ReadOnly = False
        'repoRowType.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        'repoRowType.DataSource = GetItemType()
        'repoRowType.ValueMember = "Code"
        'repoRowType.DisplayMember = "Code"
        'gv1.MasterTemplate.Columns.Add(repoRowType)
        ''---------richa 20 july,2016
        Dim repoGPTranferCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoGPTranferCode.FormatString = ""
        repoGPTranferCode.HeaderText = "Gate Pass No"
        repoGPTranferCode.Name = colgatePassTransferNo
        repoGPTranferCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoGPTranferCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoGPTranferCode.Width = 100
        StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
        If StrCrateTransferFromBooking = "1" Then
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                repoGPTranferCode.IsVisible = True
            Else
                repoGPTranferCode.IsVisible = False
            End If
        Else
            repoGPTranferCode.IsVisible = False
        End If
        gv1.MasterTemplate.Columns.Add(repoGPTranferCode)
        ''----------------
        Dim repoICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICode.FormatString = ""
        repoICode.HeaderText = "Item Code"
        repoICode.Name = colICode
        repoICode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoICode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICode.Width = 100
        gv1.MasterTemplate.Columns.Add(repoICode)
        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "Item Description"
        repoIName.Name = colIName
        repoIName.Width = 150
        repoIName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIName)
        Dim repoIHSN As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIHSN.FormatString = ""
        repoIHSN.HeaderText = "HSN Code"
        repoIHSN.Name = colIHSN
        repoIHSN.Width = 150
        repoIHSN.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIHSN)
        ''Dim repoLocationCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ''repoLocationCode.FormatString = ""
        ''repoLocationCode.HeaderText = "Location Code"
        ''repoLocationCode.Name = colLocationCode
        ''repoLocationCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        ''repoLocationCode.TextImageRelation = TextImageRelation.TextBeforeImage
        ''repoLocationCode.Width = 100
        ''gv1.MasterTemplate.Columns.Add(repoLocationCode)
        ''Dim repoLocationName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ''repoLocationName.FormatString = ""
        ''repoLocationName.HeaderText = "Location"
        ''repoLocationName.Name = colLocationName
        ''repoLocationName.ReadOnly = True
        ''repoLocationName.Width = 150
        ''gv1.MasterTemplate.Columns.Add(repoLocationName)
        Dim repoUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnit.FormatString = ""
        repoUnit.HeaderText = "UOM"
        repoUnit.Name = colUnit
        repoUnit.Width = 80
        repoUnit.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoUnit)
        repoUnit = New GridViewTextBoxColumn()
        repoUnit.FormatString = ""
        repoUnit.HeaderText = "Alt. UOM"
        repoUnit.Name = colAltUOM
        repoUnit.Width = 80
        repoUnit.ReadOnly = False
        repoUnit.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoUnit.TextImageRelation = TextImageRelation.TextBeforeImage
        gv1.MasterTemplate.Columns.Add(repoUnit)
        Dim repoRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoRate = New GridViewDecimalColumn()
        repoRate.FormatString = ""
        repoRate.HeaderText = "Unit Cost"
        repoRate.Name = colRate
        repoRate.Width = 80
        repoRate.Minimum = 0
        repoRate.FormatString = "{0:n4}"
        repoRate.DecimalPlaces = 4
        If IsItemRateeditable = True Then
            repoRate.ReadOnly = False
        Else
            repoRate.ReadOnly = True
        End If
        repoRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoRate)
        Dim repoOutQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoOutQty = New GridViewDecimalColumn()
        repoOutQty.FormatString = ""
        repoOutQty.HeaderText = "Out Qty"
        repoOutQty.Name = colOutQty
        repoOutQty.Width = 80
        repoOutQty.Minimum = 0
        repoOutQty.ShowUpDownButtons = False
        repoOutQty.Step = 0
        repoOutQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoOutQty)
        Dim repoInQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoInQty = New GridViewDecimalColumn()
        repoInQty.FormatString = ""
        repoInQty.HeaderText = "In Qty"
        repoInQty.Name = colInQty
        repoInQty.Width = 80
        repoInQty.Minimum = 0
        repoInQty.ShowUpDownButtons = False
        repoInQty.Step = 0
        repoInQty.IsVisible = False
        repoInQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoInQty)
        '==============================================================
        Dim repounitwt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repounitwt.FormatString = ""
        repounitwt.HeaderText = "Unit Weight"
        repounitwt.Name = colIUnitWt
        repounitwt.Width = 80
        repounitwt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repounitwt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repounitwt)
        Dim reponetwt As GridViewDecimalColumn = New GridViewDecimalColumn()
        reponetwt.FormatString = ""
        reponetwt.HeaderText = "Net Weight"
        reponetwt.Name = colINetWt
        reponetwt.Width = 80
        reponetwt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        reponetwt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(reponetwt)
        Dim repoMTwt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMTwt.FormatString = ""
        repoMTwt.HeaderText = "Net MT Weight"
        repoMTwt.Name = colINetMTWt
        repoMTwt.Width = 80
        repoMTwt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMTwt.ReadOnly = True
        repoMTwt.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoMTwt)
        '==========================================================
        Dim repoBinNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBinNo.FormatString = ""
        repoBinNo.HeaderText = "Bin No"
        repoBinNo.Name = colBinNo
        repoBinNo.ReadOnly = True
        repoBinNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoBinNo)
        Dim repoBreakQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoBreakQty = New GridViewDecimalColumn()
        repoBreakQty.FormatString = ""
        repoBreakQty.HeaderText = "Breakage"
        repoBreakQty.Name = colBreakQty
        repoBreakQty.Width = 80
        repoBreakQty.Minimum = 0
        repoBreakQty.ShowUpDownButtons = False
        repoBreakQty.Step = 0
        repoBreakQty.IsVisible = False
        repoBreakQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoBreakQty)
        Dim repoLeakQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLeakQty = New GridViewDecimalColumn()
        repoLeakQty.FormatString = ""
        repoLeakQty.HeaderText = "Leak"
        repoLeakQty.Name = colLeakQty
        repoLeakQty.Width = 80
        repoLeakQty.Minimum = 0
        repoLeakQty.ShowUpDownButtons = False
        repoLeakQty.Step = 0
        repoLeakQty.IsVisible = False
        repoLeakQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLeakQty)
        Dim repoShortQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoShortQty = New GridViewDecimalColumn()
        repoShortQty.FormatString = ""
        repoShortQty.HeaderText = "Shortage"
        repoShortQty.Name = colShortQty
        repoShortQty.Width = 80
        repoShortQty.Minimum = 0
        repoShortQty.ShowUpDownButtons = False
        repoShortQty.Step = 0
        repoShortQty.IsVisible = False
        repoShortQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoShortQty)
        Dim repoAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt = New GridViewDecimalColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Extended Cost"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        repoAmt.Minimum = 0
        repoAmt.ReadOnly = True
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmt)
        Dim repoAbatementRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAbatementRate = New GridViewDecimalColumn()
        repoAbatementRate.FormatString = ""
        repoAbatementRate.HeaderText = "Abatement %"
        repoAbatementRate.Name = colAbatementPer
        repoAbatementRate.Width = 80
        repoAbatementRate.Minimum = 0
        repoAbatementRate.ReadOnly = True
        repoAbatementRate.IsVisible = True
        repoAbatementRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAbatementRate)
        Dim repoAbatementamount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAbatementamount = New GridViewDecimalColumn()
        repoAbatementamount.FormatString = ""
        repoAbatementamount.HeaderText = "Abatement Amount"
        repoAbatementamount.Name = colAbatementAmount
        repoAbatementamount.Width = 80
        repoAbatementamount.Minimum = 0
        repoAbatementamount.ReadOnly = True
        repoAbatementamount.IsVisible = True
        repoAbatementamount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAbatementamount)
        Dim repoDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisPer = New GridViewDecimalColumn()
        repoDisPer.FormatString = ""
        repoDisPer.HeaderText = "Discount %"
        repoDisPer.Minimum = 0
        repoDisPer.Maximum = 100
        repoDisPer.Name = colDisPer
        repoDisPer.Width = 80
        repoDisPer.IsVisible = False
        repoDisPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDisPer)
        Dim repoDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisAmt = New GridViewDecimalColumn()
        repoDisAmt.FormatString = ""
        repoDisAmt.HeaderText = "Discount Amount"
        repoDisAmt.WrapText = True
        repoDisAmt.Name = colDisAmt
        repoDisAmt.Width = 80
        repoDisAmt.IsVisible = False
        repoDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisAmt.VisibleInColumnChooser = False
        repoDisAmt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoDisAmt)
        Dim repoAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterDis = New GridViewDecimalColumn()
        repoAmtAfterDis.FormatString = ""
        repoAmtAfterDis.HeaderText = "Amount After Discount"
        repoAmtAfterDis.Name = colAmtAfterDis
        repoAmtAfterDis.WrapText = True
        repoAmtAfterDis.Width = 80
        repoAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterDis.VisibleInColumnChooser = False
        repoAmtAfterDis.ReadOnly = True
        repoAmtAfterDis.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAmtAfterDis)
        Dim repoTax1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax1.FormatString = ""
        repoTax1.HeaderText = "Tax 1"
        repoTax1.Name = colTax1
        repoTax1.ReadOnly = True
        repoTax1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax1)
        Dim repoTaxBaseAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt1.FormatString = ""
        repoTaxBaseAmt1.HeaderText = "Tax Base Amount 1"
        repoTaxBaseAmt1.Name = colTaxBaseAmt1
        repoTaxBaseAmt1.ReadOnly = True
        repoTaxBaseAmt1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt1)
        Dim repoTaxRate1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate1 = New GridViewDecimalColumn()
        repoTaxRate1.FormatString = ""
        repoTaxRate1.HeaderText = "Tax Rate 1"
        repoTaxRate1.Name = colTaxRate1
        repoTaxRate1.IsVisible = False
        repoTaxRate1.ReadOnly = True
        repoTaxRate1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate1)
        Dim repoTaxAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt1 = New GridViewDecimalColumn()
        repoTaxAmt1.FormatString = ""
        repoTaxAmt1.HeaderText = "Tax Amt 1"
        repoTaxAmt1.Name = colTaxAmt1
        repoTaxAmt1.IsVisible = False
        repoTaxAmt1.ReadOnly = True
        repoTaxAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt1)
        Dim repoIsSurTax1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax1.HeaderText = "Is Surtax 1"
        repoIsSurTax1.Name = colIsSurTax1
        repoIsSurTax1.ReadOnly = True
        repoIsSurTax1.IsVisible = False
        repoIsSurTax1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax1)
        Dim repoSurTaxCode1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode1.FormatString = ""
        repoSurTaxCode1.HeaderText = "Surtax 1"
        repoSurTaxCode1.Name = colSurTaxCode1
        repoSurTaxCode1.ReadOnly = True
        repoSurTaxCode1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode1)
        Dim repoIsTaxable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable1.HeaderText = "Is Taxable 1"
        repoIsTaxable1.Name = colIsTaxable1
        repoIsTaxable1.ReadOnly = True
        repoIsTaxable1.IsVisible = False
        repoIsTaxable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable1)
        Dim repoTax2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax2.FormatString = ""
        repoTax2.HeaderText = "Tax 2"
        repoTax2.Name = colTax2
        repoTax2.ReadOnly = True
        repoTax2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax2)
        Dim repoTaxBaseAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt2.FormatString = ""
        repoTaxBaseAmt2.HeaderText = "Tax Base Amount 2"
        repoTaxBaseAmt2.Name = colTaxBaseAmt2
        repoTaxBaseAmt2.ReadOnly = True
        repoTaxBaseAmt2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt2)
        Dim repoTaxRate2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate2 = New GridViewDecimalColumn()
        repoTaxRate2.FormatString = ""
        repoTaxRate2.HeaderText = "Tax Rate 2"
        repoTaxRate2.Name = colTaxRate2
        repoTaxRate2.IsVisible = False
        repoTaxRate2.ReadOnly = True
        repoTaxRate2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate2)
        Dim repoTaxAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt2 = New GridViewDecimalColumn()
        repoTaxAmt2.FormatString = ""
        repoTaxAmt2.HeaderText = "Tax Amt 2"
        repoTaxAmt2.Name = colTaxAmt2
        repoTaxAmt2.IsVisible = False
        repoTaxAmt2.ReadOnly = True
        repoTaxAmt2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt2)
        Dim repoIsSurTax2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax2.HeaderText = "Is Surtax 2"
        repoIsSurTax2.Name = colIsSurTax2
        repoIsSurTax2.ReadOnly = True
        repoIsSurTax2.IsVisible = False
        repoIsSurTax2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax2)
        Dim repoSurTaxCode2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode2.FormatString = ""
        repoSurTaxCode2.HeaderText = "Surtax 2"
        repoSurTaxCode2.Name = colSurTaxCode2
        repoSurTaxCode2.ReadOnly = True
        repoSurTaxCode2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode2)
        Dim repoIsTaxable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable2.HeaderText = "Is Taxable 2"
        repoIsTaxable2.Name = colIsTaxable2
        repoIsTaxable2.ReadOnly = True
        repoIsTaxable2.IsVisible = False
        repoIsTaxable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable2)
        Dim repoTax3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax3.FormatString = ""
        repoTax3.HeaderText = "Tax 3"
        repoTax3.Name = colTax3
        repoTax3.ReadOnly = True
        repoTax3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax3)
        Dim repoTaxBaseAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt3.FormatString = ""
        repoTaxBaseAmt3.HeaderText = "Tax Base Amount 3"
        repoTaxBaseAmt3.Name = colTaxBaseAmt3
        repoTaxBaseAmt3.ReadOnly = True
        repoTaxBaseAmt3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt3)
        Dim repoTaxRate3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate3 = New GridViewDecimalColumn()
        repoTaxRate3.FormatString = ""
        repoTaxRate3.HeaderText = "Tax Rate 3"
        repoTaxRate3.Name = colTaxRate3
        repoTaxRate3.IsVisible = False
        repoTaxRate3.ReadOnly = True
        repoTaxRate3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate3)
        Dim repoTaxAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt3 = New GridViewDecimalColumn()
        repoTaxAmt3.FormatString = ""
        repoTaxAmt3.HeaderText = "Tax Amt 3"
        repoTaxAmt3.Name = colTaxAmt3
        repoTaxAmt3.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt3)
        Dim repoIsSurTax3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax3.HeaderText = "Is Surtax 3"
        repoIsSurTax3.Name = colIsSurTax3
        repoIsSurTax3.ReadOnly = True
        repoIsSurTax3.IsVisible = False
        repoIsSurTax3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax3)
        Dim repoSurTaxCode3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode3.FormatString = ""
        repoSurTaxCode3.HeaderText = "Surtax 3"
        repoSurTaxCode3.Name = colSurTaxCode3
        repoSurTaxCode3.ReadOnly = True
        repoSurTaxCode3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode3)
        Dim repoIsTaxable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable3.HeaderText = "Is Taxable 3"
        repoIsTaxable3.Name = colIsTaxable3
        repoIsTaxable3.ReadOnly = True
        repoIsTaxable3.IsVisible = False
        repoIsTaxable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable3)
        Dim repoTax4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax4.FormatString = ""
        repoTax4.HeaderText = "Tax 4"
        repoTax4.Name = colTax4
        repoTax4.ReadOnly = True
        repoTax4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax4)
        Dim repoTaxBaseAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt4.FormatString = ""
        repoTaxBaseAmt4.HeaderText = "Tax Base Amount 4"
        repoTaxBaseAmt4.Name = colTaxBaseAmt4
        repoTaxBaseAmt4.ReadOnly = True
        repoTaxBaseAmt4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt4)
        Dim repoTaxRate4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate4 = New GridViewDecimalColumn()
        repoTaxRate4.FormatString = ""
        repoTaxRate4.HeaderText = "Tax Rate 4"
        repoTaxRate4.Name = colTaxRate4
        repoTaxRate4.IsVisible = False
        repoTaxRate4.ReadOnly = True
        repoTaxRate4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate4)
        Dim repoTaxAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt4 = New GridViewDecimalColumn()
        repoTaxAmt4.FormatString = ""
        repoTaxAmt4.HeaderText = "Tax Amt 4"
        repoTaxAmt4.Name = colTaxAmt4
        repoTaxAmt4.IsVisible = False
        repoTaxAmt4.ReadOnly = True
        repoTaxAmt4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt4)
        Dim repoIsSurTax4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax4.HeaderText = "Is Surtax 4"
        repoIsSurTax4.Name = colIsSurTax4
        repoIsSurTax4.ReadOnly = True
        repoIsSurTax4.IsVisible = False
        repoIsSurTax4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax4)
        Dim repoSurTaxCode4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode4.FormatString = ""
        repoSurTaxCode4.HeaderText = "Surtax 4"
        repoSurTaxCode4.Name = colSurTaxCode4
        repoSurTaxCode4.ReadOnly = True
        repoSurTaxCode4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode4)
        Dim repoIsTaxable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable4.HeaderText = "Is Taxable 4"
        repoIsTaxable4.Name = colIsTaxable4
        repoIsTaxable4.ReadOnly = True
        repoIsTaxable4.IsVisible = False
        repoIsTaxable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable4)
        Dim repoTax5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax5.FormatString = ""
        repoTax5.HeaderText = "Tax 5"
        repoTax5.Name = colTax5
        repoTax5.ReadOnly = True
        repoTax5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax5)
        Dim repoTaxBaseAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt5.FormatString = ""
        repoTaxBaseAmt5.HeaderText = "Tax Base Amount 5"
        repoTaxBaseAmt5.Name = colTaxBaseAmt5
        repoTaxBaseAmt5.ReadOnly = True
        repoTaxBaseAmt5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt5)
        Dim repoTaxRate5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate5 = New GridViewDecimalColumn()
        repoTaxRate5.FormatString = ""
        repoTaxRate5.HeaderText = "Tax Rate 5"
        repoTaxRate5.Name = colTaxRate5
        repoTaxRate5.IsVisible = False
        repoTaxRate5.ReadOnly = True
        repoTaxRate5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate5)
        Dim repoTaxAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt5 = New GridViewDecimalColumn()
        repoTaxAmt5.FormatString = ""
        repoTaxAmt5.HeaderText = "Tax Amt 5"
        repoTaxAmt5.Name = colTaxAmt5
        repoTaxAmt5.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt5)
        Dim repoIsSurTax5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax5.HeaderText = "Is Surtax 5"
        repoIsSurTax5.Name = colIsSurTax5
        repoIsSurTax5.ReadOnly = True
        repoIsSurTax5.IsVisible = False
        repoIsSurTax5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax5)
        Dim repoSurTaxCode5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode5.FormatString = ""
        repoSurTaxCode5.HeaderText = "Surtax 5"
        repoSurTaxCode5.Name = colSurTaxCode5
        repoSurTaxCode5.ReadOnly = True
        repoSurTaxCode5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode5)
        Dim repoIsTaxable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable5.HeaderText = "Is Taxable 5"
        repoIsTaxable5.Name = colIsTaxable5
        repoIsTaxable5.ReadOnly = True
        repoIsTaxable5.IsVisible = False
        repoIsTaxable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable5)
        Dim repoTax6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax6.FormatString = ""
        repoTax6.HeaderText = "Tax 6"
        repoTax6.Name = colTax6
        repoTax6.ReadOnly = True
        repoTax6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax6)
        Dim repoTaxBaseAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt6.FormatString = ""
        repoTaxBaseAmt6.HeaderText = "Tax Base Amount 6"
        repoTaxBaseAmt6.Name = colTaxBaseAmt6
        repoTaxBaseAmt6.ReadOnly = True
        repoTaxBaseAmt6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt6)
        Dim repoTaxRate6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate6 = New GridViewDecimalColumn()
        repoTaxRate6.FormatString = ""
        repoTaxRate6.HeaderText = "Tax Rate 6"
        repoTaxRate6.Name = colTaxRate6
        repoTaxRate6.IsVisible = False
        repoTaxRate6.ReadOnly = True
        repoTaxRate6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate6)
        Dim repoTaxAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt6 = New GridViewDecimalColumn()
        repoTaxAmt6.FormatString = ""
        repoTaxAmt6.HeaderText = "Tax Amt 6"
        repoTaxAmt6.Name = colTaxAmt6
        repoTaxAmt6.IsVisible = False
        repoTaxAmt6.ReadOnly = True
        repoTaxAmt6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt6)
        Dim repoIsSurTax6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax6.HeaderText = "Is Surtax 6"
        repoIsSurTax6.Name = colIsSurTax6
        repoIsSurTax6.ReadOnly = True
        repoIsSurTax6.IsVisible = False
        repoIsSurTax6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax6)
        Dim repoSurTaxCode6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode6.FormatString = ""
        repoSurTaxCode6.HeaderText = "Surtax 6"
        repoSurTaxCode6.Name = colSurTaxCode6
        repoSurTaxCode6.ReadOnly = True
        repoSurTaxCode6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode6)
        Dim repoIsTaxable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable6.HeaderText = "Is Taxable 6"
        repoIsTaxable6.Name = colIsTaxable6
        repoIsTaxable6.ReadOnly = True
        repoIsTaxable6.IsVisible = False
        repoIsTaxable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable6)
        Dim repoTax7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax7.FormatString = ""
        repoTax7.HeaderText = "Tax 7"
        repoTax7.Name = colTax7
        repoTax7.ReadOnly = True
        repoTax7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax7)
        Dim repoTaxBaseAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt7.FormatString = ""
        repoTaxBaseAmt7.HeaderText = "Tax Base Amount 7"
        repoTaxBaseAmt7.Name = colTaxBaseAmt7
        repoTaxBaseAmt7.ReadOnly = True
        repoTaxBaseAmt7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt7)
        Dim repoTaxRate7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate7 = New GridViewDecimalColumn()
        repoTaxRate7.FormatString = ""
        repoTaxRate7.HeaderText = "Tax Rate 7"
        repoTaxRate7.Name = colTaxRate7
        repoTaxRate7.IsVisible = False
        repoTaxRate7.ReadOnly = True
        repoTaxRate7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate7)
        Dim repoTaxAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt7 = New GridViewDecimalColumn()
        repoTaxAmt7.FormatString = ""
        repoTaxAmt7.HeaderText = "Tax Amt 7"
        repoTaxAmt7.Name = colTaxAmt7
        repoTaxAmt7.IsVisible = False
        repoTaxAmt7.ReadOnly = True
        repoTaxAmt7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt7)
        Dim repoIsSurTax7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax7.HeaderText = "Is Surtax 7"
        repoIsSurTax7.Name = colIsSurTax7
        repoIsSurTax7.ReadOnly = True
        repoIsSurTax7.IsVisible = False
        repoIsSurTax7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax7)
        Dim repoSurTaxCode7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode7.FormatString = ""
        repoSurTaxCode7.HeaderText = "Surtax 7"
        repoSurTaxCode7.Name = colSurTaxCode7
        repoSurTaxCode7.ReadOnly = True
        repoSurTaxCode7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode7)
        Dim repoIsTaxable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable7.HeaderText = "Is Taxable 7"
        repoIsTaxable7.Name = colIsTaxable7
        repoIsTaxable7.ReadOnly = True
        repoIsTaxable7.IsVisible = False
        repoIsTaxable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable7)
        Dim repoTax8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax8.FormatString = ""
        repoTax8.HeaderText = "Tax 8"
        repoTax8.Name = colTax8
        repoTax8.ReadOnly = True
        repoTax8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax8)
        Dim repoTaxBaseAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt8.FormatString = ""
        repoTaxBaseAmt8.HeaderText = "Tax Base Amount 8"
        repoTaxBaseAmt8.Name = colTaxBaseAmt8
        repoTaxBaseAmt8.ReadOnly = True
        repoTaxBaseAmt8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt8)
        Dim repoTaxRate8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate8 = New GridViewDecimalColumn()
        repoTaxRate8.FormatString = ""
        repoTaxRate8.HeaderText = "Tax Rate 8"
        repoTaxRate8.Name = colTaxRate8
        repoTaxRate8.IsVisible = False
        repoTaxRate8.ReadOnly = True
        repoTaxRate8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate8)
        Dim repoTaxAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt8 = New GridViewDecimalColumn()
        repoTaxAmt8.FormatString = ""
        repoTaxAmt8.HeaderText = "Tax Amt 8"
        repoTaxAmt8.Name = colTaxAmt8
        repoTaxAmt8.IsVisible = False
        repoTaxAmt8.ReadOnly = True
        repoTaxAmt8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt8)
        Dim repoIsSurTax8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax8.HeaderText = "Is Surtax 8"
        repoIsSurTax8.Name = colIsSurTax8
        repoIsSurTax8.ReadOnly = True
        repoIsSurTax8.IsVisible = False
        repoIsSurTax8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax8)
        Dim repoSurTaxCode8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode8.FormatString = ""
        repoSurTaxCode8.HeaderText = "Surtax 8"
        repoSurTaxCode8.Name = colSurTaxCode8
        repoSurTaxCode8.ReadOnly = True
        repoSurTaxCode8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode8)
        Dim repoIsTaxable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable8.HeaderText = "Is Taxable 8"
        repoIsTaxable8.Name = colIsTaxable8
        repoIsTaxable8.ReadOnly = True
        repoIsTaxable8.IsVisible = False
        repoIsTaxable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable8)
        Dim repoTax9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax9.FormatString = ""
        repoTax9.HeaderText = "Tax 9"
        repoTax9.Name = colTax9
        repoTax9.ReadOnly = True
        repoTax9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax9)
        Dim repoTaxBaseAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt9.FormatString = ""
        repoTaxBaseAmt9.HeaderText = "Tax Base Amount 9"
        repoTaxBaseAmt9.Name = colTaxBaseAmt9
        repoTaxBaseAmt9.ReadOnly = True
        repoTaxBaseAmt9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt9)
        Dim repoTaxRate9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate9 = New GridViewDecimalColumn()
        repoTaxRate9.FormatString = ""
        repoTaxRate9.HeaderText = "Tax Rate 9"
        repoTaxRate9.Name = colTaxRate9
        repoTaxRate9.IsVisible = False
        repoTaxRate9.ReadOnly = True
        repoTaxRate9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate9)
        Dim repoTaxAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt9 = New GridViewDecimalColumn()
        repoTaxAmt9.FormatString = ""
        repoTaxAmt9.HeaderText = "Tax Amt 9"
        repoTaxAmt9.Name = colTaxAmt9
        repoTaxAmt9.IsVisible = False
        repoTaxAmt9.ReadOnly = True
        repoTaxAmt9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt9)
        Dim repoIsSurTax9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax9.HeaderText = "Is Surtax 9"
        repoIsSurTax9.Name = colIsSurTax9
        repoIsSurTax9.ReadOnly = True
        repoIsSurTax9.IsVisible = False
        repoIsSurTax9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax9)
        Dim repoSurTaxCode9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode9.FormatString = ""
        repoSurTaxCode9.HeaderText = "Surtax 9"
        repoSurTaxCode9.Name = colSurTaxCode9
        repoSurTaxCode9.ReadOnly = True
        repoSurTaxCode9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode9)
        Dim repoIsTaxable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable9.HeaderText = "Is Taxable 9"
        repoIsTaxable9.Name = colIsTaxable9
        repoIsTaxable9.ReadOnly = True
        repoIsTaxable9.IsVisible = False
        repoIsTaxable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable9)
        Dim repoTax10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax10.FormatString = ""
        repoTax10.HeaderText = "Tax 10"
        repoTax10.Name = colTax10
        repoTax10.ReadOnly = True
        repoTax10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax10)
        Dim repoTaxBaseAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt10.FormatString = ""
        repoTaxBaseAmt10.HeaderText = "Tax Base Amount 10"
        repoTaxBaseAmt10.Name = colTaxBaseAmt10
        repoTaxBaseAmt10.ReadOnly = True
        repoTaxBaseAmt10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt10)
        Dim repoTaxRate10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate10 = New GridViewDecimalColumn()
        repoTaxRate10.FormatString = ""
        repoTaxRate10.HeaderText = "Tax Rate 10"
        repoTaxRate10.Name = colTaxRate10
        repoTaxRate10.IsVisible = False
        repoTaxRate10.ReadOnly = True
        repoTaxRate10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate10)
        Dim repoTaxAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt10 = New GridViewDecimalColumn()
        repoTaxAmt10.FormatString = ""
        repoTaxAmt10.HeaderText = "Tax Amt 10"
        repoTaxAmt10.Name = colTaxAmt10
        repoTaxAmt10.IsVisible = False
        repoTaxAmt10.ReadOnly = True
        repoTaxAmt10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt10)
        Dim repoIsSurTax10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax10.HeaderText = "Is Surtax 10"
        repoIsSurTax10.Name = colIsSurTax10
        repoIsSurTax10.ReadOnly = True
        repoIsSurTax10.IsVisible = False
        repoIsSurTax10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax10)
        Dim repoSurTaxCode10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode10.FormatString = ""
        repoSurTaxCode10.HeaderText = "Surtax 10"
        repoSurTaxCode10.Name = colSurTaxCode10
        repoSurTaxCode10.ReadOnly = True
        repoSurTaxCode10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode10)
        Dim repoIsTaxable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable10.HeaderText = "Is Taxable 10"
        repoIsTaxable10.Name = colIsTaxable10
        repoIsTaxable10.ReadOnly = True
        repoIsTaxable10.IsVisible = False
        repoIsTaxable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable10)
        Dim repoItemTaxCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoItemTaxCode = New GridViewTextBoxColumn()
        repoItemTaxCode.FormatString = ""
        repoItemTaxCode.HeaderText = "Item Tax Code"
        repoItemTaxCode.Name = colItemwiseTaxCode
        repoItemTaxCode.Width = 100
        repoItemTaxCode.IsVisible = False
        repoItemTaxCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoItemTaxCode)
        Dim repoTotTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotTaxAmt = New GridViewDecimalColumn()
        repoTotTaxAmt.FormatString = ""
        repoTotTaxAmt.HeaderText = "Tax Amount"
        repoTotTaxAmt.Name = colTotTaxAmt
        repoTotTaxAmt.Width = 80
        repoTotTaxAmt.ReadOnly = True
        repoTotTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTotTaxAmt)
        Dim repoAmtAfterTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterTax = New GridViewDecimalColumn()
        repoAmtAfterTax.FormatString = ""
        repoAmtAfterTax.HeaderText = "Included Tax Amount"
        repoAmtAfterTax.Name = colAmtAfterTax
        repoAmtAfterTax.WrapText = True
        repoAmtAfterTax.Width = 80
        repoAmtAfterTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterTax.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterTax)
        Dim repoTransOutNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTransOutNo.FormatString = ""
        repoTransOutNo.HeaderText = "TransferOut No"
        repoTransOutNo.Name = colTransferOutNo
        repoTransOutNo.ReadOnly = True
        repoTransOutNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoTransOutNo)
        Dim repoIsMRPMandatory As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsMRPMandatory.HeaderText = "Is MRP Mandatory"
        repoIsMRPMandatory.Name = colisMRPMandatory
        repoIsMRPMandatory.IsVisible = False
        repoIsMRPMandatory.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        repoIsMRPMandatory.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIsMRPMandatory)
        Dim repoMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMRP.WrapText = True
        repoMRP.ReadOnly = False
        repoMRP.FormatString = ""
        repoMRP.HeaderText = "MRP"
        repoMRP.Name = colMRP
        repoMRP.Width = 80
        repoMRP.Minimum = 0
        repoMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoMRP)
        ''Dim repoAssessable As GridViewDecimalColumn = New GridViewDecimalColumn()
        ''repoAssessable.WrapText = True
        ''repoAssessable.ReadOnly = True
        ''repoAssessable.FormatString = ""
        ''repoAssessable.HeaderText = "Assessable"
        ''repoAssessable.Name = colAssessableRate
        ''repoAssessable.Width = 80
        ''repoAssessable.Minimum = 0
        ''repoAssessable.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        ''gv1.MasterTemplate.Columns.Add(repoAssessable)
        Dim repoSpecification As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSpecification = New GridViewTextBoxColumn()
        repoSpecification.FormatString = ""
        repoSpecification.HeaderText = "Specification"
        repoSpecification.Name = colSpecification
        repoSpecification.Width = 100
        gv1.MasterTemplate.Columns.Add(repoSpecification)
        Dim repoRemarks As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRemarks = New GridViewTextBoxColumn()
        repoRemarks.FormatString = ""
        repoRemarks.HeaderText = "Remarks"
        repoRemarks.Name = colRemarks
        repoRemarks.Width = 100
        gv1.MasterTemplate.Columns.Add(repoRemarks)
        'Dim repoIsUsedInGRN As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        'repoIsUsedInGRN.HeaderText = "Is Item Used IN GRN "
        'repoIsUsedInGRN.Name = colItemUsedINGRN
        'repoIsUsedInGRN.IsVisible = False
        'repoIsUsedInGRN.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        'repoIsUsedInGRN.ReadOnly = True
        'gv1.MasterTemplate.Columns.Add(repoIsUsedInGRN)
        Dim repoIsExcisable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable1.HeaderText = "Is Excisable 1"
        repoIsExcisable1.Name = colIsExcisable1
        repoIsExcisable1.ReadOnly = True
        repoIsExcisable1.IsVisible = False
        repoIsExcisable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable1)
        Dim repoIsExcisable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable2.HeaderText = "Is Excisable 2"
        repoIsExcisable2.Name = colIsExcisable2
        repoIsExcisable2.ReadOnly = True
        repoIsExcisable2.IsVisible = False
        repoIsExcisable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable2)
        Dim repoIsExcisable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable3.HeaderText = "Is Excisable 3"
        repoIsExcisable3.Name = colIsExcisable3
        repoIsExcisable3.ReadOnly = True
        repoIsExcisable3.IsVisible = False
        repoIsExcisable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable3)
        Dim repoIsExcisable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable4.HeaderText = "Is Excisable 4"
        repoIsExcisable4.Name = colIsExcisable4
        repoIsExcisable4.ReadOnly = True
        repoIsExcisable4.IsVisible = False
        repoIsExcisable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable4)
        Dim repoIsExcisable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable5.HeaderText = "Is Excisable 5"
        repoIsExcisable5.Name = colIsExcisable5
        repoIsExcisable5.ReadOnly = True
        repoIsExcisable5.IsVisible = False
        repoIsExcisable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable5)
        Dim repoIsExcisable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable6.HeaderText = "Is Excisable 6"
        repoIsExcisable6.Name = colIsExcisable6
        repoIsExcisable6.ReadOnly = True
        repoIsExcisable6.IsVisible = False
        repoIsExcisable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable6)
        Dim repoIsExcisable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable7.HeaderText = "Is Excisable 7"
        repoIsExcisable7.Name = colIsExcisable7
        repoIsExcisable7.ReadOnly = True
        repoIsExcisable7.IsVisible = False
        repoIsExcisable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable7)
        Dim repoIsExcisable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable8.HeaderText = "Is Excisable 8"
        repoIsExcisable8.Name = colIsExcisable8
        repoIsExcisable8.ReadOnly = True
        repoIsExcisable8.IsVisible = False
        repoIsExcisable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable8)
        Dim repoIsExcisable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable9.HeaderText = "Is Excisable 9"
        repoIsExcisable9.Name = colIsExcisable9
        repoIsExcisable9.ReadOnly = True
        repoIsExcisable9.IsVisible = False
        repoIsExcisable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable9)
        Dim repoIsExcisable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable10.HeaderText = "Is Excisable 10"
        repoIsExcisable10.Name = colIsExcisable10
        repoIsExcisable10.ReadOnly = True
        repoIsExcisable10.IsVisible = False
        repoIsExcisable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable10)
        Dim repoIsSerItem As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSerItem.HeaderText = "Is Serialize Item"
        repoIsSerItem.Name = colIsSerialseItem
        repoIsSerItem.ReadOnly = True
        repoIsSerItem.IsVisible = False
        repoIsSerItem.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSerItem)
        Dim repoIsPickAutoSerNo As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsPickAutoSerNo.HeaderText = "Is Pick Auto Serial"
        repoIsPickAutoSerNo.Name = colIsPickAutoSrNo
        repoIsPickAutoSerNo.ReadOnly = True
        repoIsPickAutoSerNo.IsVisible = False
        repoIsPickAutoSerNo.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsPickAutoSerNo)
        'Dim repoRGPNO As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        'repoRGPNO.FormatString = ""
        'repoRGPNO.HeaderText = "RGP No"
        'repoRGPNO.Name = colRGPNo
        'repoRGPNO.Width = 150
        'repoRGPNO.ReadOnly = True
        'repoRGPNO.IsVisible = False
        'gv1.MasterTemplate.Columns.Add(repoRGPNO)
        Dim repoIsBatchItem As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsBatchItem.HeaderText = "Is Batch Item"
        repoIsBatchItem.Name = colIsBatchItem
        repoIsBatchItem.ReadOnly = True
        repoIsBatchItem.IsVisible = False
        repoIsBatchItem.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsBatchItem)
        Dim repoFOCItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFOCItem.FormatString = ""
        repoFOCItem.HeaderText = "FOC Item"
        repoFOCItem.Name = colFOCItem
        repoFOCItem.ReadOnly = True
        repoFOCItem.IsVisible = False
        repoFOCItem.Width = 100
        gv1.MasterTemplate.Columns.Add(repoFOCItem)
        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)
        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
    End Sub
    Sub LoadBlankGridTax()
        gv2.Rows.Clear()
        gv2.Columns.Clear()
        Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthCode.FormatString = ""
        repoTaxAuthCode.HeaderText = "Tax Authority Code"
        repoTaxAuthCode.Name = colTTaxAutCode
        repoTaxAuthCode.Width = 150
        repoTaxAuthCode.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthCode)
        Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthName.FormatString = ""
        repoTaxAuthName.HeaderText = "Tax Authority"
        repoTaxAuthName.Name = colTTaxAutName
        repoTaxAuthName.Width = 200
        repoTaxAuthName.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthName)
        Dim repoTaxAssessableAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAssessableAmt.FormatString = ""
        repoTaxAssessableAmt.HeaderText = "Assessable Amount"
        repoTaxAssessableAmt.Name = colTTaxAssessableAmt
        repoTaxAssessableAmt.Width = 100
        repoTaxAssessableAmt.ReadOnly = True
        repoTaxAssessableAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxAssessableAmt)
        Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt.FormatString = ""
        repoTaxBaseAmt.HeaderText = "Base Amount"
        repoTaxBaseAmt.Name = colTBaseAmt
        repoTaxBaseAmt.Width = 100
        repoTaxBaseAmt.ReadOnly = True
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)
        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = ""
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colTTaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = False
        repoTaxRate.IsVisible = True
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxRate)
        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = ""
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colTTaxAmt
        repoTaxAmt.Width = 100
        repoTaxAmt.ReadOnly = False
        repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxAmt)
        gv2.AllowAddNewRow = False
        gv2.ShowGroupPanel = False
        gv2.AllowColumnReorder = False
        gv2.AllowRowReorder = False
        gv2.EnableSorting = False
        gv2.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv2.MasterTemplate.ShowRowHeaderColumn = False
    End Sub
    Private Sub CalMT_WT(ByVal IntRow As Integer)
        If clsCommon.myLen(Gross_Weight_Unit) > 0 Then
            Dim Wt_Value As Decimal = 0
            Dim Weight_UOM As String = ""
            Wt_Value = clsCommon.myCdbl(clsItemMaster.GetItemWeightValue(clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value), Nothing))
            Weight_UOM = clsCommon.myCstr(clsItemMaster.GetItemWeightUnit(clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value), Nothing))
            Wt_Value = Wt_Value * clsCommon.myCdbl(clsItemMaster.GetConvertionFactor(clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRow).Cells(colUnit).Value), Nothing))
            gv1.Rows(IntRow).Cells(colIUnitWt).Value = Wt_Value
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "R") = CompairStringResult.Equal Then
                gv1.Rows(IntRow).Cells(colINetWt).Value = Wt_Value * clsCommon.myCdbl(gv1.Rows(IntRow).Cells(colOutQty).Value)
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                gv1.Rows(IntRow).Cells(colINetWt).Value = Wt_Value * clsCommon.myCdbl(gv1.Rows(IntRow).Cells(colInQty).Value)
            End If
            ''richa agarwal TEC/28/03/19-000462 add item structure on setting based
            Dim Qry As String = "select top 1 CF from (select (case when (Container_UOM='" & Weight_UOM & "' and Contained_UOM='" & Weight_MT_Unit & "') then round(Contained_Qty/Container_Qty,4) else case when (Container_UOM='" & Weight_MT_Unit & "' and Contained_UOM='" & Weight_UOM & "') then round(Container_Qty/Contained_Qty,4) end end) as CF,product_type from TSPL_WEIGHT_CONVERSION where product_type in ('ALL','" + clsItemMaster.GetItemProductType(clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value), Nothing) + "') " & IIf(ItemStructureMandatoryOnWeightConversion = True, "and isnull(Structure_Code,'') =(select Structure_Code  from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value) & "')", "") & " )aa where isnull(cast(CF as varchar),'')<>'' order by Product_Type desc"
            Dim mt_uom_cnvrsn As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
            If clsCommon.CompairString(Weight_UOM, Weight_MT_Unit) = CompairStringResult.Equal Then
                mt_uom_cnvrsn = 1
            End If
            Qry = "select top 1 CF from (select (case when (Container_UOM='" & Weight_UOM & "' and Contained_UOM='" & Gross_Weight_Unit & "') then round(Contained_Qty/Container_Qty,4) else case when (Container_UOM='" & Gross_Weight_Unit & "' and Contained_UOM='" & Weight_UOM & "') then round(Container_Qty/Contained_Qty,4) end end) as CF,product_type from TSPL_WEIGHT_CONVERSION where product_type in ('ALL','" + clsItemMaster.GetItemProductType(clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value), Nothing) + "')  " & IIf(ItemStructureMandatoryOnWeightConversion = True, "and isnull(Structure_Code,'') =(select Structure_Code  from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(gv1.Rows(IntRow).Cells(colICode).Value) & "')", "") & " )aa where isnull(cast(CF as varchar),'')<>'' order by Product_Type desc"
            Dim gross_uom_cnvrsn As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
            If clsCommon.CompairString(Weight_UOM, Gross_Weight_Unit) = CompairStringResult.Equal Then
                gross_uom_cnvrsn = 1
            End If
            If gross_uom_cnvrsn > 0 Then
                gv1.Rows(IntRow).Cells(colINetMTWt).Value = clsCommon.myCdbl(gv1.Rows(IntRow).Cells(colINetWt).Value) * gross_uom_cnvrsn
            Else
                gv1.Rows(IntRow).Cells(colINetMTWt).Value = 0
                If clsCommon.CompairString(Gross_Weight_Unit, Weight_UOM) = CompairStringResult.Equal Then
                    gv1.Rows(IntRow).Cells(colINetMTWt).Value = clsCommon.myCdbl(gv1.Rows(IntRow).Cells(colINetWt).Value)
                End If
            End If
            If ProvisionAllow = False Then
                gv1.Rows(IntRow).Cells(colINetMTWt).Value = clsCommon.myCdbl(gv1.Rows(IntRow).Cells(colINetWt).Value)
            End If
        End If
    End Sub
    Private Sub gv1_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.FieldName.StartsWith("_CFLD_") Then
                        clsCustomFieldGrid.getFinderForCustomFieldGrid(gv1, e.Column.Name.ToString, MyBase.Form_ID)
                    End If
                    If e.Column Is gv1.Columns(colTotTaxAmt) Then
                        UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
                        UpdateAllTotals()
                    End If
                    If e.Column Is gv1.Columns(colUnit) OrElse e.Column Is gv1.Columns(colTotTaxAmt) OrElse e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colOutQty) OrElse e.Column Is gv1.Columns(colOutQty) OrElse e.Column Is gv1.Columns(colInQty) OrElse e.Column Is gv1.Columns(colBreakQty) OrElse e.Column Is gv1.Columns(colLeakQty) OrElse e.Column Is gv1.Columns(colShortQty) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colSpecification) OrElse e.Column Is gv1.Columns(colRemarks) OrElse e.Column Is gv1.Columns(colDisPer) OrElse e.Column Is gv1.Columns(colAmt) OrElse (e.Column Is gv1.Columns(colMRP)) Then
                        If e.Column Is gv1.Columns(colOutQty) OrElse e.Column Is gv1.Columns(colInQty) OrElse e.Column Is gv1.Columns(colBreakQty) OrElse e.Column Is gv1.Columns(colLeakQty) OrElse e.Column Is gv1.Columns(colShortQty) OrElse e.Column Is gv1.Columns(colRate) OrElse (e.Column Is gv1.Columns(colAmt)) OrElse e.Column Is gv1.Columns(colMRP) Then
                            If (e.Column Is gv1.Columns(colOutQty) OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colInQty).Value) OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colLeakQty).Value) OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colBreakQty).Value) OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colShortQty).Value) > 0) Then
                                Dim dblOutQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colOutQty).Value)
                                Dim dblInQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colInQty).Value)
                                Dim dblLeakQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colLeakQty).Value)
                                Dim dblBreakQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colBreakQty).Value)
                                Dim dblShortQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colShortQty).Value)
                                Dim dblTransferInQty As Double = 0
                                dblTransferInQty = dblInQty + dblLeakQty + dblShortQty + dblBreakQty
                                If dblTransferInQty > dblOutQty Then
                                    gv1.CurrentRow.Cells(e.Column.Name).Value = 0
                                    Throw New Exception("In Quantity Can't be more than Out Quantity." + Environment.NewLine + "Out Quantity : " + clsCommon.myCstr(dblOutQty) + ". In Quantity : " + clsCommon.myCstr(dblTransferInQty))
                                End If
                            End If
                            If (e.Column Is gv1.Columns(colOutQty) OrElse e.Column Is gv1.Columns(colInQty)) Then
                                OpenSerialItem()
                                If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                                    If gv1.CurrentRow.Cells(colIsBatchItem).Value = True Then
                                        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                                            OpenBatchItem()
                                        End If

                                    End If
                                End If
                                'OpenBatchItem()
                            End If
                            '---------------Calculate Abatement Amount--------------------------
                            gv1.CurrentRow.Cells(colAbatementAmount).Value = ((clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) * clsCommon.myCdbl(gv1.CurrentRow.Cells(colAbatementPer).Value)) / 100) * clsCommon.myCdbl(gv1.CurrentRow.Cells(colOutQty).Value)
                            UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
                            If rbtnTaxCalManual.IsChecked Then
                                For ii As Integer = 0 To gv1.Rows.Count - 1
                                    UpdateCurrentRow(ii)
                                Next
                            End If
                            '============calc. unit wt and net wt.===========
                            CalMT_WT(gv1.CurrentRow.Index)
                            '==============================================
                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colICode) Then
                            If Not clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                                OpenICodeList(False)
                                Dim strItem As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                                Dim strUnit As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                                Dim strDate As String = txtDate.Value
                                Dim dblUnitCost As Double = 0
                                ''-----------------
                                ''richa agarwal 15 Nov,2016
                                If PickRateFromPRICEChrtFORUMang Then
                                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select (case when Is_Scheme_Item=1 then 'F' else item_type end) as item_type from TSPL_ITEM_MASTER  where Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' ")), "F") = CompairStringResult.Equal Then
                                        Dim qry As String = String.Empty
                                        If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) > 0 Then
                                            qry = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from ( " &
                                                   " Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                                                   " Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date, " &
                                                   " Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                                                   " TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                                                   " TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code where isnull(TSPL_ITEM_PRICE_MASTER.Type  ,'')='T'  and Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  " &
                                                   "  and UOM='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(txtFromLocation.Value) & "'  " &
                                                   " ) XXXE WHERE RowNo=1 "
                                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                                            If dt.Rows.Count > 0 Then
                                                gv1.CurrentRow.Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                                                gv1.CurrentRow.Cells(colRate).ReadOnly = True
                                            Else
                                                Throw New Exception("Please create Price chart for Location " & txtFromLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & ". ")
                                                gv1.CurrentCell.Focus()
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        setAvgRate(gv1.CurrentRow.Index)
                                    End If
                                ElseIf settStockTranferFromTransferPriceAndInvJVWithAvgCost Then
                                    If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) > 0 Then
                                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select (case when Is_Scheme_Item=1 then 'F' else item_type end) as item_type from TSPL_ITEM_MASTER  where Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' ")), "F") = CompairStringResult.Equal Then
                                            Qry = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from ( " &
                                               " Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                                               " Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date, " &
                                               " Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                                               " TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                                               " TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code where isnull(TSPL_ITEM_PRICE_MASTER.Type  ,'')='T'  and Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  " &
                                               "  and UOM='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(txtFromLocation.Value) & "'  " &
                                               " ) XXXE WHERE RowNo=1 "
                                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                                            If dt.Rows.Count > 0 Then
                                                gv1.CurrentRow.Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                                                gv1.CurrentRow.Cells(colRate).ReadOnly = True
                                            Else
                                                Throw New Exception("Please Create Transfer Price chart for item [" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "] and Location [" & txtFromLocation.Value & "].")
                                                gv1.CurrentCell.Focus()
                                                Exit Sub
                                            End If
                                        Else
                                            setAvgRate(gv1.CurrentRow.Index)
                                        End If
                                    End If
                                Else
                                    setAvgRate(gv1.CurrentRow.Index)
                                End If
                                ''-----------------end  richa 
                            End If
                            '============calc. unit wt and net wt.===========
                            CalMT_WT(gv1.CurrentRow.Index)
                            '==============================================
                        ElseIf e.Column Is gv1.Columns(colUnit) Then
                            If Not clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                                OpenUOMList(False)
                                Dim strItem As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                                Dim strUnit As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                                Dim strDate As String = txtDate.Value
                                Dim dblUnitCost As Double = 0
                                If PickRateFromPRICEChrtFORUMang = False Then
                                    Dim dblCostMethod As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Costing_Method as Costing_Method from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Item_Code='" & strItem & "' "))
                                    If dblCostMethod <> 0 Then
                                        dblUnitCost = clsInventoryMovement.GetCost(dblCostMethod, strItem, txtFromLocation.Value, 1, strDate, strDate, False, Nothing, "TSPL_INVENTORY_MOVEMENT", strUnit)
                                        gv1.CurrentRow.Cells(colRate).Value = dblUnitCost
                                    Else
                                        gv1.CurrentRow.Cells(colRate).Value = 0
                                    End If
                                End If
                                '============calc. unit wt and net wt.===========
                                CalMT_WT(gv1.CurrentRow.Index)
                                '==============================================
                            End If
                        End If
                        ''setGridFocus()
                    End If
                    isCellValueChangedOpen = False
                    If e.Column Is gv1.Columns(colAltUOM) Then
                        isCellValueChangedOpen = True
                        OpenAltUOMList(False)
                        isCellValueChangedOpen = False
                    End If
                End If
                If clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) > 0 OrElse clsCommon.myLen(fndSRNO.Value) > 0 Then
                    If AllowOnlyOneIssueAgainstStoreRequisition = False Then
                        gv1.CurrentRow.Cells(colICode).ReadOnly = True
                    End If
                Else
                    gv1.CurrentRow.Cells(colICode).ReadOnly = False
                End If
                If PickRateFromPRICEChrtFORUMang = False AndAlso settStockTranferFromTransferPriceAndInvJVWithAvgCost = False Then
                    If Not clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                        If (e.Column Is gv1.Columns(colOutQty) OrElse e.Column Is gv1.Columns(colUnit) OrElse e.Column Is gv1.Columns(colICode)) Then
                            GetConvRateUOM(gv1.CurrentRow.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            isCellValueChangedTaxOpen = False
            isCellValueChangedOpen = False
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub setAvgRate(ByVal rowID As Integer)
        Dim strDate As String = txtDate.Value
        If PickRateFromPRICEChrtFORUMang = False Then
            Dim dblCostMethod As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Costing_Method as Costing_Method from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Item_Code='" & clsCommon.myCstr(gv1.Rows(rowID).Cells(colICode).Value) & "' "))
            If dblCostMethod <> 0 Then
                gv1.Rows(rowID).Cells(colRate).Value = clsInventoryMovement.GetCost(dblCostMethod, clsCommon.myCstr(gv1.Rows(rowID).Cells(colICode).Value), txtFromLocation.Value, 1, strDate, strDate, False, Nothing)
            Else
                gv1.Rows(rowID).Cells(colRate).Value = 0
            End If
        End If
    End Sub
    Sub GetConvRateUOM(ByVal IntRowNo As Integer)
        Try
            Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colOutQty).Value)
            Dim dblRate As Double = 0
            Dim dblCostMethod As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Costing_Method as Costing_Method from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Item_Code='" & strItem & "' "))
            If dblCostMethod <> 0 Then
                dblRate = clsInventoryMovement.GetCost(dblCostMethod, strItem, txtFromLocation.Value, 1, txtDate.Value, txtDate.Value, False, Nothing)
            End If
            '  Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblConvRate As Double = 0
            If clsCommon.myLen(strOrgUnit) > 0 Then
                Dim dblOrgConvF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
                Dim dblStockingUnitConvF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
                dblConvRate = Math.Round((dblRate * dblOrgConvF), 2)
            End If
            gv1.Rows(IntRowNo).Cells(colRate).Value = dblConvRate
        Catch ex As Exception
        End Try
    End Sub
    Sub OpenUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        If clsCommon.myLen(strICode) > 0 Then
            Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
            Dim whrCls As String = "TSPL_ITEM_UOM_DETAIL.Item_Code='" + strICode + "'"
            gv1.CurrentRow.Cells(colUnit).Value = clsCommon.ShowSelectForm("TRNSKLUOMFND", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), "Code", isButtonClick)
        End If
    End Sub
    Sub OpenAltUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        If clsCommon.myLen(strICode) > 0 Then
            Dim qry As String = "select TSPL_ITEM_UOM_DETAIL.UOM_Code as Code,tspl_unit_master.unit_desc as [Description] from TSPL_ITEM_UOM_DETAIL"
            qry += " left outer join tspl_unit_master on tspl_unit_master.unit_code=TSPL_ITEM_UOM_DETAIL.uom_code "
            Dim whrCls As String = "TSPL_ITEM_UOM_DETAIL.Item_Code='" + strICode + "' and TSPL_ITEM_UOM_DETAIL.uom_code <> '" + clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) + "'"
            gv1.CurrentRow.Cells(colAltUOM).Value = clsCommon.ShowSelectForm("TRNSKLUOMFND1", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colAltUOM).Value), "Code", isButtonClick)
        End If
    End Sub
    Sub OpenICodeList(ByVal isButtonClick As Boolean)
        Dim obj As clsItemMaster = clsItemMaster.FinderForItem(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), "", False, isButtonClick, "", "")
        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Item_Code) > 0 Then
            gv1.CurrentRow.Cells(colICode).Value = obj.Item_Code
            gv1.CurrentRow.Cells(colIName).Value = obj.Item_Desc
            gv1.CurrentRow.Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
            gv1.CurrentRow.Cells(colUnit).Value = obj.Unit_Code
            gv1.CurrentRow.Cells(colisMRPMandatory).Value = obj.Is_MRP
            gv1.CurrentRow.Cells(colIsSerialseItem).Value = obj.Is_Serial_Item
            gv1.CurrentRow.Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
            gv1.CurrentRow.Cells(colIsPickAutoSrNo).Value = obj.Is_Pick_Auto_SrNo
            gv1.CurrentRow.Cells(colBinNo).Value = obj.Rack_No
            If obj.Tax_Exempted = 2 Then
                gv1.CurrentRow.Cells(colAbatementPer).Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select TOP(1) Abatement_Percent from TSPL_ABATEMENT_MASTER"))
            End If
        Else
            SetBlankOfItemColumns()
        End If
        ''End If
        SetitemWiseTaxSetting(True, True)
        setBalance()
    End Sub
    Private Sub SetBlankOfItemColumns()
        gv1.CurrentRow.Cells(colICode).Value = ""
        gv1.CurrentRow.Cells(colIName).Value = ""
        gv1.CurrentRow.Cells(colIHSN).Value = ""
        gv1.CurrentRow.Cells(colUnit).Value = ""
        gv1.CurrentRow.Cells(colMRP).Value = 0
        gv1.CurrentRow.Cells(colAbatementPer).Value = 0
        gv1.CurrentRow.Cells(colAbatementAmount).Value = 0
        gv1.CurrentRow.Cells(colisMRPMandatory).Value = False
        ''gv1.CurrentRow.Cells(colAssessableRate).Value = 0
    End Sub
    Private Sub gv1_UserAddedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserAddedRow
        For i As Integer = 0 To gv1.Rows.Count - 1
            gv1.Rows(0).Cells(0).Value = 1
            If i <> 0 Then
                gv1.Rows(i).Cells(colLineNo).Value = i + 1
            End If
        Next
    End Sub
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        GstStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        Dim arrTaxableAuth As New List(Of String)
        Dim dblOutQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colOutQty).Value)
        Dim dblInQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colInQty).Value)
        Dim dblLeakQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colLeakQty).Value)
        Dim dblBreakQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colBreakQty).Value)
        Dim dblShortQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colShortQty).Value)
        Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
        Dim dblAmt As Double = 0
        If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "R") = CompairStringResult.Equal Then
            dblAmt = dblOutQty * dblRate
            gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2)
        Else
            dblAmt = (dblInQty + dblLeakQty + dblBreakQty + dblShortQty) * dblRate
            gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2)
        End If
        ''richa agarwal
        Dim strItemCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strItemtype As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select case when CSA_TYPE='BULK-DESI GHEE' then 1 when CSA_TYPE='CPD-DESI GHEE' then 1 else 0 end   from TSPL_ITEM_MASTER where Item_Code ='" & strItemCode & "' "))
        strTaxType = clsLocationWiseTax.TaxType(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value, Nothing)
        ''------------------------------------
        If clsCommon.myLen(txtTaxGroup.Value) >= 0 Then
            'Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
            Dim dblAmtAfterDis As Double = dblAmt
            'If clsCommon.CompairString(cboItemType.SelectedValue, "O") = CompairStringResult.Equal Then '' CODE COMMENTED BY PANCH RAJ DISSCUSSED WITH RANJANA MAM
            If (clsCommon.CompairString(cboType.SelectedValue, "Excise") = CompairStringResult.Equal And GstStatus = False) Or (GstStatus AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal) Or (GstStatus AndAlso clsCommon.CompairString(strTaxType, "I") = CompairStringResult.Equal) Or (GstStatus AndAlso IsMandiTax = True) Or (clsCommon.CompairString(cboType.SelectedValue, "Depot") = CompairStringResult.Equal AndAlso clsCommon.CompairString(strItemtype, "1") = CompairStringResult.Equal) Or clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAbatementPer).Value) > 0 Then
                For ii As Integer = 1 To 10
                    Dim Strii As String = clsCommon.myCstr(ii)
                    If rbtnTaxCalAutomatic.IsChecked Then
                        Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                        If clsCommon.myLen(strTaxCode) > 0 Then
                            Dim dblTaxRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                            Dim IsSurTax As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value)
                            Dim strSurTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                            Dim IsTaxable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                            Dim dblBaseAmt As Double = 0
                            Dim dblTaxAmt As Double = 0
                            If IsSurTax Then
                                Dim dblSurTaxAmt As Double = GetCurrentRowSurTaxAmt(IntRowNo, ii, strSurTaxCode)
                                dblBaseAmt = dblSurTaxAmt
                            Else
                                Dim dblOtherTaxAmt As Double = 0
                                dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                                If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
                                    dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                                Else
                                    If clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAbatementPer).Value) = 0 Then
                                        dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                                    Else
                                        dblBaseAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAbatementAmount).Value) + dblOtherTaxAmt
                                    End If
                                End If
                            End If
                            'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Math.Round(dblBaseAmt + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAbatementAmount).Value), 2)
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Math.Round(dblBaseAmt, 2)
                            dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                            If ii = 1 Then
                                'If chkExciseOnQty.Checked Then
                                '    gv1.Rows(IntRowNo).Cells(colAssessableAmount).Value = Math.Round(dblOutQty, 2)
                                '    dblTaxAmt = (dblOutQty * dblTaxRate) / 100
                                'Else
                                '    gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblBaseAmt, 2)
                                'End If
                            End If
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, IIf(objCommonVar.IsRoundOffTaxToZeroDecimal, 0, 2))
                            If IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()) Then
                                arrTaxableAuth.Add(strTaxCode.ToUpper())
                            End If
                        Else
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxRate" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value = Nothing
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value = Nothing
                        End If
                    ElseIf rbtnTaxCalManual.IsChecked Then
                        If gv2.Rows.Count >= ii Then
                            Dim dblTaxAmt As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxAmt).Value)
                            Dim dblTaxRate As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxRate).Value)
                            Dim dblCurrRowAmt As Double = clsCommon.myCdbl(gv1.Rows(clsCommon.myCdbl(IntRowNo)).Cells(colAmt).Value)
                            Dim dblTotAmt As Double = 0
                            For jj As Integer = 0 To gv1.Rows.Count - 1
                                dblTotAmt += clsCommon.myCdbl(gv1.Rows(jj).Cells(colAmt).Value)
                            Next
                            Dim dblCurrCalTax As Double = 0
                            If dblTotAmt <> 0 Then
                                dblCurrCalTax = Math.Round(clsCommon.myCdbl(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 2, MidpointRounding.ToEven)
                            End If
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = dblCurrCalTax
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value = dblTaxRate
                        End If
                    End If
                Next
                Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmt(IntRowNo)
                Dim dblAmtAfterTax As Double = dblAmtAfterDis + dblTotTaxAmt
                gv1.Rows(IntRowNo).Cells(colDisAmt).Value = 0
                gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = 0
                gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = dblTotTaxAmt
                gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = dblAmtAfterTax
            Else
                gv1.Rows(IntRowNo).Cells(colDisAmt).Value = 0
                gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = dblAmt
                gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = 0
                gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = dblAmt
            End If
        Else
            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = 0
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = 0
            gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = 0
            gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = 0
        End If
    End Sub
    Private Function GetCurrentRowTotalTaxAmt(ByVal IntRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value)
            Else
                dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value)
            End If
        Next
        Return dblTotTax
    End Function
    Private Function GetCurrentRowSurTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Else
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            End If
        Next
        Return 0
    End Function
    Private Function GetCurrentRowOtherTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If IntRowNo < 0 Then
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                End If
            Next
        Next
        Return dblRet
    End Function
    Private Sub BlankTaxDetails(ByVal intRowNo As Integer)
        BlankTaxDetails(intRowNo, True)
    End Sub
    Private Sub BlankTaxDetails(ByVal intRowNo As Integer, ByVal isBlankRate As Boolean)
        For ii As Integer = 1 To 10
            Dim strII As String = clsCommon.myCstr(ii)
            If intRowNo < 0 Then
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXBASEAMT" + strII)).Value = Nothing
                If isBlankRate Then
                    gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                    gv1.CurrentRow.Cells(colItemwiseTaxCode).Value = Nothing
                End If
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            Else
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + strII)).Value = Nothing
                If isBlankRate Then
                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                    gv1.Rows(intRowNo).Cells(colItemwiseTaxCode).Value = Nothing
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            End If
        Next
    End Sub
    Private Sub UpdateAllTotals()
        Dim dblTotAmt As Double = 0
        Dim dblTotDisAmt As Double = 0
        Dim dblAmtAfterDis As Double = 0
        Dim dblTaxAssessableAmt As Double = 0
        Dim dblMTWt As Double = 0
        Dim dblTaxBaseAmt1 As Double = 0
        Dim dblTaxBaseAmt2 As Double = 0
        Dim dblTaxBaseAmt3 As Double = 0
        Dim dblTaxBaseAmt4 As Double = 0
        Dim dblTaxBaseAmt5 As Double = 0
        Dim dblTaxBaseAmt6 As Double = 0
        Dim dblTaxBaseAmt7 As Double = 0
        Dim dblTaxBaseAmt8 As Double = 0
        Dim dblTaxBaseAmt9 As Double = 0
        Dim dblTaxBaseAmt10 As Double = 0
        Dim dblTaxAmt1 As Double = 0
        Dim dblTaxAmt2 As Double = 0
        Dim dblTaxAmt3 As Double = 0
        Dim dblTaxAmt4 As Double = 0
        Dim dblTaxAmt5 As Double = 0
        Dim dblTaxAmt6 As Double = 0
        Dim dblTaxAmt7 As Double = 0
        Dim dblTaxAmt8 As Double = 0
        Dim dblTaxAmt9 As Double = 0
        Dim dblTaxAmt10 As Double = 0
        Dim dblTaxTotAmt As Double = 0
        Dim dblNetAmt As Double = 0
        For ii As Integer = 0 To gv1.Rows.Count - 1
            If (clsCommon.myLen(gv1.Rows(ii).Cells(colICode).Value) > 0) Then
                dblTotAmt = dblTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmt).Value)
                dblTotDisAmt = dblTotDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colDisAmt).Value)
                dblAmtAfterDis = dblAmtAfterDis + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterDis).Value)
                dblTaxAmt1 = dblTaxAmt1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt1).Value)
                dblTaxAmt2 = dblTaxAmt2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt2).Value)
                dblTaxAmt3 = dblTaxAmt3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt3).Value)
                dblTaxAmt4 = dblTaxAmt4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt4).Value)
                dblTaxAmt5 = dblTaxAmt5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt5).Value)
                dblTaxAmt6 = dblTaxAmt6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt6).Value)
                dblTaxAmt7 = dblTaxAmt7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt7).Value)
                dblTaxAmt8 = dblTaxAmt8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt8).Value)
                dblTaxAmt9 = dblTaxAmt9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt9).Value)
                dblTaxAmt10 = dblTaxAmt10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt10).Value)
                'dblTaxAssessableAmt = dblTaxAssessableAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAssessableAmount).Value)
                dblTaxBaseAmt1 = dblTaxBaseAmt1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt1).Value)
                dblTaxBaseAmt2 = dblTaxBaseAmt2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt2).Value)
                dblTaxBaseAmt3 = dblTaxBaseAmt3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt3).Value)
                dblTaxBaseAmt4 = dblTaxBaseAmt4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt4).Value)
                dblTaxBaseAmt5 = dblTaxBaseAmt5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt5).Value)
                dblTaxBaseAmt6 = dblTaxBaseAmt6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt6).Value)
                dblTaxBaseAmt7 = dblTaxBaseAmt7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt7).Value)
                dblTaxBaseAmt8 = dblTaxBaseAmt8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt8).Value)
                dblTaxBaseAmt9 = dblTaxBaseAmt9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt9).Value)
                dblTaxBaseAmt10 = dblTaxBaseAmt10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt10).Value)
                dblTaxTotAmt = dblTaxTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTotTaxAmt).Value)
                dblNetAmt = dblNetAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterTax).Value)
            End If
        Next
        If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
            If rbtnTaxCalAutomatic.IsChecked Then
                For ii As Integer = 1 To gv2.Rows.Count
                    Select Case (ii)
                        Case 1
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt1, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt1, 2)
                            If chkExciseOnQty.Checked Then
                                If dblTaxAssessableAmt <> 0 Then
                                    gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt1 * 100) / dblTaxAssessableAmt, 2)
                                Else
                                    gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                                End If
                            ElseIf dblTaxBaseAmt1 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt1 * 100) / dblTaxBaseAmt1, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = Math.Round(dblTaxAssessableAmt, 2)
                        Case 2
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt2, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt2, 2)
                            If dblTaxBaseAmt2 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt2 * 100) / dblTaxBaseAmt2, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 3
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt3, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt3, 2)
                            If dblTaxBaseAmt3 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt3 * 100) / dblTaxBaseAmt3, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 4
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt4, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt4, 2)
                            If dblTaxBaseAmt4 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt4 * 100) / dblTaxBaseAmt4, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 5
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt5, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt5, 2)
                            If dblTaxBaseAmt5 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt5 * 100) / dblTaxBaseAmt5, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 6
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt6, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt6, 2)
                            If dblTaxBaseAmt6 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt6 * 100) / dblTaxBaseAmt6, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 7
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt7, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt7, 2)
                            If dblTaxBaseAmt7 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt7 * 100) / dblTaxBaseAmt7, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 8
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt8, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt8, 2)
                            If dblTaxBaseAmt8 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt8 * 100) / dblTaxBaseAmt8, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 9
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt9, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt9, 2)
                            If dblTaxBaseAmt9 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt9 * 100) / dblTaxBaseAmt9, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                        Case 10
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt10, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt10, 2)
                            If dblTaxBaseAmt10 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt10 * 100) / dblTaxBaseAmt10, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAssessableAmt).Value = gv2.Rows(ii - 1).Cells(colTBaseAmt).Value
                    End Select
                Next
            End If
        Else
            dblNetAmt = clsCommon.myFormat(dblTotAmt)
        End If
        lblAmtWithoutTax.Text = clsCommon.myFormat(dblTotAmt)
        lblDiscountAmt.Text = clsCommon.myFormat(dblTotDisAmt)
        lblAmtAfterDiscount.Text = clsCommon.myFormat(dblAmtAfterDis)
        lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)
        dblMTWt = 0
        For Each grow As GridViewRowInfo In gv1.Rows
            dblMTWt = dblMTWt + clsCommon.myCdbl(grow.Cells(colINetMTWt).Value)
        Next
        dblNetAmt = dblNetAmt
        txttotal_Wt.Text = dblMTWt
        lblTotRAmt.Text = clsCommon.myFormat(dblNetAmt)
        ''================for gopal ji===============
        If clsCommon.CompairString(StrCrateTransferFromBooking, "1") = CompairStringResult.Equal Then
            Dim cratewt As Decimal = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ItemCrateWtinKg, clsFixedParameterCode.ItemCrateWtinKg, Nothing))
            Dim jaaliwt As Decimal = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ItemJaaliWtinKg, clsFixedParameterCode.ItemJaaliWtinKg, Nothing))
            Dim boxwt As Decimal = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ItemBoxWtinKg, clsFixedParameterCode.ItemBoxWtinKg, Nothing))
            If clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                cratewt = cratewt * clsCommon.myCdbl(txtCrateIn.Text)
                jaaliwt = jaaliwt * clsCommon.myCdbl(txtJaaliIn.Text)
                boxwt = boxwt * clsCommon.myCdbl(txtBoxIn.Text)
            ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                cratewt = cratewt * clsCommon.myCdbl(txtcrateout.Text)
                jaaliwt = jaaliwt * clsCommon.myCdbl(txtjaaliout.Text)
                boxwt = boxwt * clsCommon.myCdbl(txtBoxOut.Text)
            End If
            If clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal Then
                txtGross_Wt.Text = clsCommon.myCdbl(txttotal_Wt.Text) + clsCommon.myCdbl(cratewt) + clsCommon.myCdbl(jaaliwt) + clsCommon.myCdbl(boxwt)
            Else
                txtGross_Wt.Text = clsCommon.myCdbl(txttotal_Wt.Text)
            End If
        End If
        ''====================end here=======================
        If GrossWtfromItemMaster Then
            txtGross_Wt.Text = "0"
            txttotal_Wt.Text = "0"
            Dim unit_gross_wt As Double = 0
            Dim unit_Net_wt As Double = 0
            For Each grow As GridViewRowInfo In gv1.Rows
                Dim itemCode As String = clsCommon.myCstr(grow.Cells(colICode).Value)
                Dim Qty As Decimal = clsCommon.myCdbl(grow.Cells(colOutQty).Value)
                If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                    Qty = clsCommon.myCdbl(grow.Cells(colInQty).Value)
                End If
                Dim Unit As String = clsCommon.myCstr(grow.Cells(colUnit).Value)
                Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable("select gross_weight,Net_Weight from tspl_item_uom_detail where item_code='" + itemCode + "' and uom_code='" + Unit + "'")
                If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    If clsCommon.myCdbl(dtTemp.Rows(0)("gross_weight")) <= 0 Then
                        Throw New Exception("Please set gross weight for item:" + itemCode + " and UOM:" + Unit)
                    End If
                    If clsCommon.myCdbl(dtTemp.Rows(0)("Net_Weight")) <= 0 Then
                        Throw New Exception("Please set net weight for item:" + itemCode + " and UOM:" + Unit)
                    End If
                    unit_gross_wt += clsCommon.myCdbl(dtTemp.Rows(0)("gross_weight")) * Qty
                    unit_Net_wt += clsCommon.myCdbl(dtTemp.Rows(0)("Net_Weight")) * Qty
                End If
            Next
            txtGross_Wt.Text = clsCommon.myCstr(Math.Round(unit_gross_wt, 3))
            txttotal_Wt.Text = clsCommon.myCstr(Math.Round(unit_Net_wt, 3))
        End If
        Dim toloc As String = ""
        toloc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GIT_Location from TSPL_LOCATION_MASTER where Location_Code='" + txtToLoc.Value + "'", Nothing))
        'txtvehicle_Charge.Text = clsTransferDCC.GetProvisionCharge(txtFromLocation.Value, txtToLoc.Value, clsCommon.myCdbl(txtGross_Wt.Text), clsCommon.myCdbl(txtVehicle_Capacity.Text), txtTransporter_Code.Value)
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
        Else
            FillVehicleCharges()
        End If
        lblTotRAmt1.Text = lblTotRAmt.Text
    End Sub
    Private Sub FillVehicleCharges()
        Try
            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
            Else
                Dim dt As New DataTable()
                dt = clsTransferDCC.GetProvisionCharge(txtFromLocation.Value, txtToLoc.Value, clsCommon.myCdbl(txtGross_Wt.Text), clsCommon.myCdbl(txtVehicle_Capacity.Text), txtTransporter_Code.Value)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    txtvehicle_Charge.Text = clsCommon.myCdbl(dt.Rows(0)("FreightCharge"))
                    txtvehicle_Charge.Tag = dt
                Else
                    txtvehicle_Charge.Text = "0"
                    txtvehicle_Charge.Tag = Nothing
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function GetBaseOtherTaxableAmount(ByVal intEndCol As Integer) As Double
        ''Dim dblRetVal As Double = 0
        ''For ii As Integer = 0 To intEndCol - 1
        ''    If clsCommon.myCBool(gv2.Rows(ii).Cells(colTIsTaxable).Value) Then
        ''        dblRetVal = dblRetVal + clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
        ''    End If
        ''Next
        ''Return dblRetVal
    End Function
    Private Function GetBaseTaxAmount(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 0 To intEndCol
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAutCode).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
            End If
        Next
        Return 0
    End Function
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub
    Sub AddNew()
        isCellValueChangedOpen = False
        isInsideLoadData = False
        cboType.Enabled = True
        BlankAllControls()
        pnlLoadIn.Visible = False
        LoadBlankGrid()
        LoadBlankGridTax()
        txtTaxGroup.Enabled = True
        isNewEntry = True
        btnSave.Text = "Save"
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        btnPrintNew.Enabled = True
        ' BtnPrintChallan.Enabled = True
        If chkInternalTransfer.Checked OrElse chkJobWorkTransfer.Checked Then
            BtnPrintChallan.Enabled = False
            ' rbtnPrintExcisable.Enabled = False
        Else
            BtnPrintChallan.Enabled = True
            ' rbtnPrintExcisable.Enabled = True
        End If
        chkJobWork.Checked = False
        txtDate.Focus()
        gv1.Rows.AddNew()
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
            chkAgainst_Form.Checked = True
        Else
            chkAgainst_Form.Checked = False
            chkAgainst_Form.Enabled = False
        End If
        TxtTransportorMName.MendatroryField = True
        TxtTransportorMName.Visible = False
        fndTransferIndentNo.Value = ""
        ''richa agarwal 19 july,2016
        StrCrateTransferFromBooking = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing))
        If StrCrateTransferFromBooking = "1" Then
            GatePassPanel.Visible = True
            lblpricecode.Visible = True
            fndPriceCode.Visible = True
        ElseIf ShowPriceCode Then
            lblpricecode.Visible = True
            fndPriceCode.Visible = True
        Else
            GatePassPanel.Visible = False
            lblpricecode.Visible = False
            fndPriceCode.Visible = False
        End If
        FndGatePassNo.Value = ""
        fndPriceCode.Value = ""
        LoadGPItemType()
        txtToLoc.Enabled = True
        ''-------------
        Dim ShowPrintChallan As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.ShowPrintChallanInDairyDispatch, clsFixedParameterCode.ShowPrintChallanInDairyDispatch, Nothing) = "1", True, False))
        If ShowPrintChallan = True Then
            BtnPrintChallan.Visible = False
        Else
            BtnPrintChallan.Visible = True
        End If
        Dim ShowCrateJaaliBox As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, Nothing) = "1", True, False))
        If ShowCrateJaaliBox = True Then
            txtCrateIn.Visible = False
            txtJaaliIn.Visible = False
            txtBoxIn.Visible = False
            txtcrateout.Visible = True
            txtjaaliout.Visible = True
            txtBoxOut.Visible = True
            lblCrateIn.Visible = False
            lblJaaliIn.Visible = False
            lblBoxIn.Visible = False
            lblCrateOut.Visible = True
            lblJaaliOut.Visible = True
            lblBoxout.Visible = True
            txtCrateIn.Text = 0
            txtJaaliIn.Text = 0
            txtBoxIn.Text = 0
            txtcrateout.Text = 0
            txtjaaliout.Text = 0
            txtBoxOut.Text = 0
        Else
            txtCrateIn.Visible = False
            txtJaaliIn.Visible = False
            txtBoxIn.Visible = False
            txtcrateout.Visible = False
            txtjaaliout.Visible = False
            txtBoxOut.Visible = False
            lblCrateIn.Visible = False
            lblJaaliIn.Visible = False
            lblBoxIn.Visible = False
            lblCrateOut.Visible = False
            lblJaaliOut.Visible = False
            lblBoxout.Visible = False
        End If
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.SetDefaultValues()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        chkTaxable.Checked = False
        If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
            'chkTaxable.Checked = False
            'cboType.Enabled = False
            'chkAgainst_Form.Enabled = False
        Else
            'chkTaxable.Checked = True
            'cboType.Enabled = True
            'chkAgainst_Form.Enabled = True
        End If
        chkInternalTransfer.Checked = False
        chkInternalTransfer.Enabled = True
        chkJobWorkTransfer.Enabled = True
        chkJobWorkTransfer.Checked = False
        chkProductionRequest.Checked = False
        chkProductionRequest.Enabled = False
        fndSRNO.Value = ""
        fndSRNO.Enabled = True
        txtTransferOutNo.Enabled = True
        txtLoadingAdviceNo.Text = ""

        If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
            fndPriceCode.Visible = True
            lblpricecode.Visible = True
            Dim qry As String = " Select DISTINCT TSPL_PRICE_COMPONENT_MAPPING.Price_Code from TSPL_PRICE_COMPONENT_MAPPING   where TSPL_PRICE_COMPONENT_MAPPING.Default_For_Tarnsfer='1'  "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                fndPriceCode.Value = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
            End If
        End If
    End Sub
    Function AllowToSave() As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            If chkInternalTransfer.Checked Then
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, txtDocNo.Value)
            End If
            If ApplyFEFO = True Then
                If clsCommon.myLen(txtLoadingAdviceNo.Text) <= 0 Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Enter Loading Advice No ", Me.Text)
                    txtLoadingAdviceNo.Focus()
                    Return False
                End If
            End If
            Dim Qry As String
            If clsCommon.CompairString(cboTransferType.Text, "Transfer In") = CompairStringResult.Equal Then
                Qry = "select Document_Date  from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + txtTransferOutNo.Value + "'"
                If clsCommon.myCDate(txtDate.Value, "dd/MMM/yyyy") < clsCommon.myCDate(clsDBFuncationality.getSingleValue(Qry), "dd/MMM/yyyy") Then
                    Throw New Exception("Transfer In Date can't be less than transferout Date")
                End If
            End If
            If ShowPriceCode Then
                If clsCommon.myLen(fndPriceCode.Value) <= 0 Then
                    Throw New Exception("Please select Price Code")
                End If
            End If
            '=====================Added by preeti Gupta==================
            If AllowFutureDateTransaction(txtDate.Value, Nothing) = False Then
                txtDate.Focus()
                Return False
            End If
            '=================================================================
            ''richa agarwal 19 july,2016
            StrCrateTransferFromBooking = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing))
            If StrCrateTransferFromBooking = "1" Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                    If clsCommon.myLen(FndGatePassNo.Value) <= 0 Then
                        Throw New Exception("Please select Gate Pass No")
                    End If
                    If clsCommon.myLen(fndPriceCode.Value) <= 0 Then
                        Throw New Exception("Please select Price Code")
                    End If
                    If clsCommon.CompairString(cmbGPItemType.SelectedValue, "S") = CompairStringResult.Equal Then
                        Throw New Exception("Please select Item Type")
                    End If
                    FillLocationInfo(False)
                    If clsCommon.myLen(lblToLoc.Tag) <= 0 Then
                        Throw New Exception("Please create GIT Location of " & txtToLoc.Value & "")
                    End If
                    If clsCommon.CompairString(cmbGPItemType.SelectedValue, "E") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Excisable  from TSPL_LOCATION_MASTER  where  Location_Code ='" & clsCommon.myCstr(txtFromLocation.Value) & "'")), "T") <> CompairStringResult.Equal Then
                        Throw New Exception("Please select From Location of Excisable Type")
                    End If
                End If
            End If
            If clsCommon.CompairString(cboTransferType.Text, "Transfer In") = CompairStringResult.Equal Then
                Qry = "select isnull(Is_Status_IN ,'') from TSPL_TRANSFER_ORDER_HEAD WHERE Transfer_Type='O' and Document_No In (select Document_No from TSPL_TRANSFER_ORDER_HEAD WHERE Transfer_Type='I' AND Document_No<>'" + txtDocNo.Value + "' AND TransferOutNo='" + txtTransferOutNo.Value + "' ) AND TransferOutNo='" + txtTransferOutNo.Value + "' "
                Qry = clsDBFuncationality.getSingleValue(Qry)
                If clsCommon.CompairString(Qry, "Y") = CompairStringResult.Equal Then
                    Throw New Exception("Transfer-Out ''" + txtTransferOutNo.Value + "'' has been used with other Transfer-In ")
                End If
            End If
            If isALlowVehicleGateOutValidation = True Then
                If clsCommon.myLen(txtvehicle_mannual_no.Text) > 0 Then
                    Dim qry1 As String = String.Empty
                    qry1 = " SELECT Stuff((SELECT N', ' + TSPL_TRANSFER_ORDER_HEAD.Document_No FROM TSPL_TRANSFER_ORDER_HEAD left join TSPL_Transfer_Gate_Out on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_Transfer_Gate_Out.Transfer_No  where TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No='" & txtvehicle_mannual_no.Text & "' and TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' and TSPL_Transfer_Gate_Out.Transfer_No is null FOR XML PATH(''),TYPE).value('text()[1]','nvarchar(max)'),1,2,N'') "
                    Dim result As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry1))
                    If clsCommon.myLen(result) > 0 Then
                        common.clsCommon.MyMessageBoxShow("Vehicle No ('" & txtvehicle_mannual_no.Text & "') used in other Shipment No. You can create new Shipment with Vehicle No :('" & txtvehicle_mannual_no.Text & "')  After  Gate Out following Shipment No : '" & result & "'")
                        Return False
                    End If
                End If
            End If
            If chkAgainst_Form.Checked AndAlso chkTaxable.Checked = True Then
                If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                    Throw New Exception("Please select tax group.")
                End If
            End If
            If isNewEntry = False Then
                If clsCommon.myLen(txtDocNo.Value) > 0 Then
                    Dim intExemptedType As Integer = clsDBFuncationality.getSingleValue("select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & PreviousTaxGroupCode & "'", Nothing)
                    Dim ExemptedTypeNew As Integer = clsDBFuncationality.getSingleValue("select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & txtTaxGroup.Value & "'", Nothing)
                    If chkForRepair.Checked Then
                        If chkTaxable.Checked AndAlso CheckRepair = 0 Then
                            If chkForRepair.Checked Then
                                Throw New Exception("This Document can not be saved As Repair.")
                            End If
                        End If
                        If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                            If ExemptedTypeNew = 0 Then
                                Throw New Exception("Please Tag Only EXEMPTED Type of tax group.")
                            End If
                        End If
                    Else
                        If chkTaxable.Checked AndAlso intExemptedType = 1 Then
                            If intExemptedType <> ExemptedTypeNew Then
                                Throw New Exception("Please Tag Only EXEMPTED Type of tax group.")
                            End If
                        ElseIf chkTaxable.Checked AndAlso intExemptedType = 0 Then
                            If intExemptedType <> ExemptedTypeNew Then
                                Throw New Exception("Please Tag Same Type of tax group as Previous.")
                            End If
                        End If
                    End If
                End If
            End If
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                If clsCommon.myLen(clsCommon.myCstr(txtTransferOutNo.Value)) = 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Transfer No is mandatory.")
                    txtTransferOutNo.Focus()
                    Return False
                End If
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "R") = CompairStringResult.Equal Then
                If clsCommon.myLen(clsCommon.myCstr(txtRMDANo.Value)) = 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Rejected No is mandatory.")
                    txtRMDANo.Focus()
                    Return False
                End If
            End If
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso clsCommon.myLen(clsCommon.myCstr(cboType.SelectedValue)) <= 0 Then
                cboType.Focus()
                Throw New Exception("Please select Type")
            ElseIf clsCommon.myLen(cboModeOfTransport.Text) <= 0 Then
                cboModeOfTransport.Focus()
                Throw New Exception("Mode Of Transport Can't be Blank")
            ElseIf txtFromLocation.Value = String.Empty Then
                txtFromLocation.Focus()
                Throw New Exception("From Location can't be Blank")
            ElseIf txtToLoc.Value = String.Empty Then
                txtToLoc.Focus()
                Throw New Exception("To Location can't be Blank")
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal And txtTransferOutNo.Value = String.Empty Then
                txtTransferOutNo.Focus()
                Throw New Exception("LoadOut No. can't be Blank")
            End If
            Qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GIT_Location from TSPL_LOCATION_MASTER where Location_Code='" + txtFromLocation.Value + "'"))
            If clsCommon.CompairString(Qry, txtToLoc.Value) = CompairStringResult.Equal Then
                txtToLoc.Focus()
                Throw New Exception("To location can't be the GIT location of from locaion")
            End If
            If chkInternalTransfer.Checked = True Then
            ElseIf chkJobWorkTransfer.Checked Then
            ElseIf chkOwnVehicle.Checked = True AndAlso clsCommon.myLen(clsCommon.myCstr(TxtTransportorMName.Text)) <= 0 Then
                TxtTransportorMName.Focus()
                Throw New Exception("Please fill transpoter name")
            End If
            Dim ShowCrateJaaliBox As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, Nothing) = "1", True, False))
            If ShowCrateJaaliBox = True Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(txtCrateIn.Text) > clsCommon.myCdbl(txtcrateout.Text) Then
                        txtcrateout.Focus()
                        Throw New Exception("Crate In is not greater then crate Out")
                    End If
                    If clsCommon.myCdbl(txtBoxIn.Text) > clsCommon.myCdbl(txtBoxOut.Text) Then
                        txtBoxOut.Focus()
                        Throw New Exception("Box In is not greater then Box Out")
                    End If
                    If clsCommon.myCdbl(txtJaaliIn.Text) > clsCommon.myCdbl(txtjaaliout.Text) Then
                        txtjaaliout.Focus()
                        Throw New Exception("Jaali In is not greater then Jaali Out")
                    End If
                End If
            End If
            Dim arrProjNo As New List(Of String)
            Dim arrReqNo As New List(Of String)
            Dim arrICode As New List(Of String)()
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strTransferOutNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colTransferOutNo).Value)
                Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                Dim strIName As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                Dim strIHSNCode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIHSN).Value)
                Dim dblOutQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colOutQty).Value)
                Dim dblInQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colInQty).Value)
                Dim dblBreakQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colBreakQty).Value)
                Dim dblLeakQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colLeakQty).Value)
                Dim dblShortQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colShortQty).Value)
                Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colMRP).Value)
                If clsCommon.myLen(strICode) > 0 Then
                    '===============Added by preeti gupta Against ticket no[ERO/15/02/19-000490]=========
                    If clsCommon.myCdbl(dblOutQty) <= 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "Out Qty can't be zero for [" + strIName + "] Item. At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    '=====================================================================================
                    If clsCommon.myLen(strIHSNCode) <= 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "HSN Code can't be blank for [" + strIName + "] Item. At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    If clsCommon.myCBool(gv1.Rows(ii).Cells(colisMRPMandatory).Value) AndAlso dblMRP <= 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "Please enter MRP for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    If IsItemRateeditable = True AndAlso clsCommon.CompairString(cboTransferType.Text, "Transfer Out") = CompairStringResult.Equal Then
                        Dim strIsschemeItem = clsDBFuncationality.getSingleValue("select top 1 TSPL_GATEPASS_TRANSFER_HEAD.Booking_Code from TSPL_GATEPASS_TRANSFER_DETAIL left outer join " &
                                  "TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.document_no=TSPL_GATEPASS_TRANSFER_DETAIL.document_no left outer join " &
                                  "TSPL_BOOKING_DETAIL on TSPL_GATEPASS_TRANSFER_HEAD.Booking_Code=TSPL_BOOKING_DETAIL.Document_No where  " &
                                  "TSPL_GATEPASS_TRANSFER_HEAD.Document_No='" & clsCommon.myCstr(gv1.Rows(i).Cells(colgatePassTransferNo).Value) & "' and " &
                                  "TSPL_GATEPASS_TRANSFER_DETAIL.item_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value) & "' and Scheme_Item='Y'")
                        StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                        If clsCommon.myLen(strIsschemeItem) = 0 Then
                            If StrCrateTransferFromBooking = "1" Then
                                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colRate).Value) = 0 Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Please enter Item Rate for " + strICode + ". At Line No" + clsCommon.myCstr(ii + 1))
                                    Return False
                                End If
                            Else
                                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colRate).Value) = 0 Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Please enter Item Rate for " + strICode + ". At Line No" + clsCommon.myCstr(ii + 1))
                                    Return False
                                End If
                            End If
                        End If
                    End If
                    'Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strICode, txtFromLocation.Value, txtDocNo.Value, clsCommon.GETSERVERDATE(), Nothing, strUOM, dblMRP)
                    Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strICode, txtFromLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, strUOM, dblMRP)
                    If clsCommon.myLen(strTransferOutNo) > 0 Then
                        Dim dblTransferInQty As Double = dblInQty + dblLeakQty + dblBreakQty + dblShortQty
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then 'preeti gupta[UDL/24/07/18-000207]
                            If dblTransferInQty > dblOutQty Then
                                Throw New Exception("In Qty Should not be greater then Out Qty." + Environment.NewLine + "Out Quantity : " + clsCommon.myCstr(dblOutQty) + ". In Quantity : " + clsCommon.myCstr(dblTransferInQty))
                            End If
                        Else
                            If dblTransferInQty < dblOutQty Then
                                Throw New Exception("In Qty Should be equal to Out Qty." + Environment.NewLine + "Out Quantity : " + clsCommon.myCstr(dblOutQty) + ". In Quantity : " + clsCommon.myCstr(dblTransferInQty))
                            End If
                            If dblTransferInQty > dblOutQty Then
                                Throw New Exception("In Qty Should be equal to Out Qty." + Environment.NewLine + "Out Quantity : " + clsCommon.myCstr(dblOutQty) + ". In Quantity : " + clsCommon.myCstr(dblTransferInQty))
                            End If
                        End If
                        If dblInQty > 0 AndAlso clsCommon.myCBool(gv1.Rows(ii).Cells(colIsBatchItem).Value) AndAlso clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
                            Dim arrBatchNo As List(Of clsBatchInventory) = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                            If arrBatchNo Is Nothing Then
                                Throw New Exception("Please provide Batch no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                            Else
                                Dim tQty As Decimal = 0
                                For Each objBatch As clsBatchInventory In arrBatchNo
                                    tQty += objBatch.Qty
                                Next
                                If tQty <> dblInQty Then
                                    Throw New Exception("Item : " + strICode + " Entered Qty " + clsCommon.myCstr(dblInQty) + Environment.NewLine + "And Batchwise Qty " + clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                                End If
                            End If
                            If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                                Dim strqry As String = "SELECT (MAX(Out_Qty)-sUM(In_Qty)) AS Ret_Qty  FROM TSPL_TRANSFER_ORDER_DETAIL " &
                                " LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No WHERE  Transfer_Type ='T' " &
                                " and TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo='" & txtTransferOutNo.Value & "' and TSPL_TRANSFER_ORDER_DETAIL.Item_Code ='" & strICode & "'  and TSPL_TRANSFER_ORDER_DETAIL.Document_No <>'" + txtDocNo.Value + "' " &
                                " GROUP BY TSPL_TRANSFER_ORDER_DETAIL.transferoutno,Item_Code"
                                ''richa UDL/27/12/18-000246 26 Dec,2018
                                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strqry, Nothing)
                                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                    ' Dim dblremainingRetQty As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry, Nothing))
                                    Dim dblremainingRetQty As Double = clsCommon.myCdbl(dt.Rows(0)("Ret_Qty"))
                                    If dblInQty > dblremainingRetQty Then
                                        Throw New Exception("not allowed ,because all quantity return for this Item : " + strICode)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If dblOutQty > dblBalQty Then
                            common.clsCommon.MyMessageBoxShow(Me, "Item - " + strICode + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblOutQty) + " and Balance Quantity - " + clsCommon.myCstr(dblBalQty), Me.Text)
                            Return False
                        End If
                        If clsCommon.myLen(txtRMDANo.Value) > 0 Then
                            Dim strPINo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MAX( TSPL_PI_DETAIL.PI_No) as PI_No from TSPL_SRN_DETAIL  left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No left outer join TSPL_PI_DETAIL on TSPL_PI_DETAIL.SRN_Id=TSPL_SRN_DETAIL.SRN_No where TSPL_SRN_HEAD.RMDA_No='" + txtRMDANo.Value + "'"))
                            dblBalQty = clsPurchaseInvoiceDetail.GetBalancePIQty(strPINo, strICode, txtDocNo.Value, strUOM, dblMRP, 0, True)
                            If dblOutQty > dblBalQty Then
                                common.clsCommon.MyMessageBoxShow(Me, "Item - " + strICode + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblOutQty) + " and RMDA Quantity - " + clsCommon.myCstr(dblBalQty), Me.Text)
                                Return False
                            End If
                        End If
                        If dblOutQty > 0 AndAlso clsCommon.myCBool(gv1.Rows(ii).Cells(colIsBatchItem).Value) AndAlso clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
                            If (RunBatchFifowisewithmodifyfunctionality = True AndAlso clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) <= 0) Then
                                OpenBatchItem()
                            ElseIf RunBatchFifowise = 1 AndAlso RunBatchFifowisewithmodifyfunctionality = False Then
                                OpenBatchItem()
                            End If
                            'If RunBatchFifowise = 1 Then
                            '    gv1.CurrentRow = gv1.Rows(ii)
                            '    OpenBatchItem()
                            'End If
                            Dim arrBatchNo As List(Of clsBatchInventory) = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                            If arrBatchNo Is Nothing Then
                                Throw New Exception("Please provide Batch no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                            Else
                                Dim tQty As Decimal = 0
                                For Each objBatch As clsBatchInventory In arrBatchNo
                                    tQty += objBatch.Qty
                                Next
                                If tQty <> dblOutQty Then
                                    Throw New Exception("Item : " + strICode + " Entered Qty " + clsCommon.myCstr(dblOutQty) + Environment.NewLine + "And Batchwise Qty " + clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                                End If
                            End If
                        End If
                    End If
                End If
                If clsCommon.myLen(strICode) > 0 Then
                    If Not arrICode.Contains(strICode) Then
                        arrICode.Add(strICode)
                    End If
                End If
                If clsCommon.myLen(strICode) > 0 Then
                    For jj As Integer = 0 To gv1.Rows.Count - 1
                        If jj = ii Then
                            Continue For
                        End If
                        Dim strInnerICode As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                        Dim dblInnerMRP As Double = clsCommon.myCdbl(gv1.Rows(jj).Cells(colMRP).Value)
                        Dim strInnerReqNo As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colTransferOutNo).Value)
                        Dim strInnerUOM As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                        If dblMRP = dblInnerMRP AndAlso clsCommon.CompairString(strICode, strInnerICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOM, strInnerUOM) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strTransferOutNo, strInnerReqNo) = CompairStringResult.Equal Then
                            Dim Msg As String = "Same Item Exist at Row No " + clsCommon.myCstr(ii + 1) + " And " + clsCommon.myCstr(jj + 1)
                            Msg = Msg + Environment.NewLine + "Item: " + strICode + "(" + strIName + ")"
                            Msg = Msg + Environment.NewLine + "UOM: " + strUOM
                            If dblMRP > 0 Then
                                Msg = Msg + Environment.NewLine + "MRP: " + clsCommon.myCstr(dblMRP)
                            End If
                            common.clsCommon.MyMessageBoxShow(Me, Msg, Me.Text)
                            Return False
                        End If
                    Next
                End If
                If clsCommon.myCBool(gv1.Rows(ii).Cells(colIsSerialseItem).Value) Then
                    Dim arrSerailNo As List(Of clsSerializeInvenotry) = TryCast(gv1.Rows(ii).Tag, List(Of clsSerializeInvenotry))
                    Dim dblQty As Double = 0
                    If clsCommon.myLen(strTransferOutNo) > 0 Then
                        dblQty = dblInQty
                    Else
                        dblQty = dblOutQty
                    End If
                    If arrSerailNo Is Nothing OrElse dblQty <> arrSerailNo.Count Then
                        Throw New Exception("Please provide serial no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                    End If
                End If
                ''richa agarwal 17 Nov,2016
                If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                    If PickRateFromPRICEChrtFORUMang Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select item_type from TSPL_ITEM_MASTER  where Item_Code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "' ")), "F") = CompairStringResult.Equal Then
                            Dim qry1 As String = String.Empty
                            If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)) > 0 Then
                                qry1 = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from ( " &
                                       " Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                                       " Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date, " &
                                       " Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                                       " TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                                       " TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code where isnull(TSPL_ITEM_PRICE_MASTER.Type  ,'')='T'  and Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  " &
                                       "  and UOM='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(txtFromLocation.Value) & "'  " &
                                       " ) XXXE WHERE RowNo=1 "
                                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry1)
                                If dt.Rows.Count > 0 Then
                                    gv1.Rows(ii).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                                    gv1.Rows(ii).Cells(colRate).ReadOnly = True
                                Else
                                    Throw New Exception("Please create Price chart for Location " & txtFromLocation.Value & "  for item " & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & ". ")
                                End If
                            End If
                        Else
                            If settPickProductCostFromItemUOMDetail Then
                                GetConvRateUOM(ii)
                            End If
                        End If
                    ElseIf settStockTranferFromTransferPriceAndInvJVWithAvgCost Then
                        Dim qry1 As String = String.Empty
                        If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)) > 0 Then
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select item_type from TSPL_ITEM_MASTER  where Item_Code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "' ")), "F") = CompairStringResult.Equal Then
                                qry1 = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from ( " &
                                   " Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                                   " Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date, " &
                                   " Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                                   " TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                                   " TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code where isnull(TSPL_ITEM_PRICE_MASTER.Type  ,'')='T'  and Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  " &
                                   "  and UOM='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(txtFromLocation.Value) & "'  " &
                                   " ) XXXE WHERE RowNo=1 "
                                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry1)
                                If dt.Rows.Count > 0 Then
                                    gv1.Rows(ii).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                                    gv1.Rows(ii).Cells(colRate).ReadOnly = True
                                Else
                                    Throw New Exception("Please Create Transfer Price chart for item [" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "] and Location [" & txtFromLocation.Value & "].")
                                End If
                            Else
                                GetConvRateUOM(ii)
                            End If
                        End If
                    Else
                        If settPickProductCostFromItemUOMDetail Then
                            GetConvRateUOM(ii)
                        End If
                    End If
                End If
            Next
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal And clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select IsParlour from tspl_location_master where Location_code='" & clsCommon.myCstr(txtToLoc.Value) & "'")), "Y") = CompairStringResult.Equal Then
                If clsCommon.myLen(fndTransferIndentNo.Value) <= 0 Then
                    fndTransferIndentNo.Focus()
                    Throw New Exception("Please select Transfer Indent No")
                End If
            End If
            For ii As Integer = 0 To gv1.Rows.Count - 1
                UpdateCurrentRow(ii)
            Next
            UpdateAllTotals()
            UcCustomFields1.AllowToSave()
            UcAttachment1.AllowToSave()
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                chkTaxable.Checked = clsLocationWiseTax.IsTaxable(txtFromLocation.Value, txtToLoc.Value, txtDate.Value, Nothing)
                chkIsManditax.Checked = clsLocationWiseTax.IsMandiTax(txtTaxGroup.Value, Nothing)
            End If
            If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                    If chkForRepair.Checked = True Then
                        chkTaxable.Checked = False
                    Else
                        If chkTaxable.Checked Then
                            clsLocationWiseTax.IsValidTaxGroupForTransfer(txtTaxGroup.Value, txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value, False, Nothing)
                        End If
                    End If
                End If
            Else
                Dim intx As Integer = clsItemMaster.isItemOfSameExcisable(arrICode)
                If Not (intx = arrICode.Count OrElse intx = 0) Then
                    Throw New Exception("All item should be of Excisable or NonExcisable")
                End If
                If intx > 0 Then
                    If clsCommon.CompairString(cboTransferType.Text, "Transfer In") <> CompairStringResult.Equal Then
                        If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                            Throw New Exception("Please select tax group.")
                        Else
                            If clsLocation.isLocatinExcisable(txtFromLocation.Value) = True Then
                                For Each grow As GridViewRowInfo In gv2.Rows
                                    If Not clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Excisable from TSPL_TAX_MASTER WHERE Tax_Code='" + grow.Cells(colTTaxAutCode).Value + "'")), "Y") = CompairStringResult.Equal Then
                                        Throw New Exception("Atleast One tax should be excisable.")
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        Item_TaxType = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select TOP(1) Is_Tax_Exempted from TSPL_ITEM_MASTER where Item_Code in (" + clsCommon.GetMulcallString(arrICode) + ")"))
                    End If
                Else
                    Item_TaxType = 0
                End If
            End If
            If ValidateTaxGroup Then
                If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                    Throw New Exception("Please select tax group.")
                End If
            End If
            txtFreightDistance.Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select MAX(Distance) AS Distance from TSPL_LOCATION_DISTANCE_MAPPING where TransType='T' and ((Location_Code='" & txtFromLocation.Value & "' and Customer_Code='" & txtToLoc.Value & "') OR (Customer_Code='" & txtToLoc.Value & "' and Location_Code='" & txtFromLocation.Value & "'))", Nothing))
            'Dim ELocationType As String = clsDBFuncationality.getSingleValue("select CASE WHEN TSPL_LOCATION_MASTER.Registered=1 then 'BB' else 'BC' end AS Type from TSPL_LOCATION_MASTER where TSPL_LOCATION_MASTER.Location_Code='" + txtToLoc.Value + "'", Nothing)
            If chkJobWork.Checked = False AndAlso (clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal) AndAlso objCommonVar.GenerateEWayBillWithEInvoice = True AndAlso chkTaxable.Checked = True AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True Then
                'If clsCommon.myCdbl(txtFreightDistance.Value) <= 0 Then
                '    Throw New Exception("Please define Freight Distance in EWay Bill Distance Master.")
                'End If
                If chkOwnVehicle.Checked = False Then
                    If clsCommon.myLen(txtTransporter_desc.Text) <= 0 Then
                        Throw New Exception("Pls Select Transporter")
                        txtTransporter_Code.Focus()
                        Return False
                    End If
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select GSTRegistered from tspl_vendor_master where vendor_code='" & txtTransporter_Code.Value & "'", Nothing)) = 0 Then
                        Throw New Exception("Transporter must be registered.")
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SavingData(False)
    End Sub
    Private Function SavingData(ByVal ChekBtnPost As Boolean) As Boolean
        If (SaveData(False, ChekBtnPost)) Then
            If ChekBtnPost = False Then
                common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
            End If
            Return True
        Else
            Return False
        End If
    End Function
    '' Anubhooti 09-Sep-2014 BM00000003735
    Private Function FinYrCheck(ByVal Save As Boolean, ByVal Post As Boolean) As Boolean
        If (Save = True And Post = False Or Save = False And Post = True) Then
            Dim FormName As String = ""
            If chkInternalTransfer.Checked Then
                FormName = "ITransfer"
            ElseIf chkJobWorkTransfer.Checked Then
                FormName = "JTransfer"
            Else
                FormName = "Transfer"
            End If
            If FrmMainTranScreen.ValidateTransactionAccToFinYear(FormName, txtDate.Value) = False Then
                Return False
            End If
        End If
        Return True
        ''
    End Function
    Private Function SaveData(ByVal isDoAbandomentNo As Boolean, Optional ByVal ChekBtnPost As Boolean = False) As Boolean
        Try
            If FinYrCheck(True, False) = False Then
                Return False
            End If
            clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value))
            If (AllowToSave()) Then
                Dim obj As New clsTransferDCC()
                obj.GLVoucher_No = txtglvoucher.Text
                obj.Document_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Delivery_date = txtOutDate.Value
                obj.Delivery_Duration = clsCommon.myCstr(txtDeliveryDuration.Text)
                obj.Secondary_Code = txtSecondary_Doc_Code.Text
                obj.Loading_Advice_No = clsCommon.myCstr(txtLoadingAdviceNo.Text)
                obj.Vehicle_Mannual_No = clsCommon.myCstr(txtvehicle_mannual_no.Text)
                obj.Vehicle_Charge = clsCommon.myCdbl(txtvehicle_Charge.Text)
                obj.IsEwayBill = IIf(chkEwaybill.Checked, 1, 0)
                ''======================================================================
                If txtvehicle_Charge.Tag IsNot Nothing AndAlso TryCast(txtvehicle_Charge.Tag, DataTable) IsNot Nothing AndAlso TryCast(txtvehicle_Charge.Tag, DataTable).Rows.Count > 0 Then
                    Dim dt As DataTable = TryCast(txtvehicle_Charge.Tag, DataTable)
                    obj.Freight_Type = clsCommon.myCstr(dt.Rows(0)("FreightType"))
                    obj.FixedCharge = clsCommon.myCdbl(dt.Rows(0)("FixedCharge"))
                    obj.EmptyCharge = clsCommon.myCdbl(dt.Rows(0)("EmptyCharge"))
                End If
                ''======================================================================
                obj.Vehicle_Capacity = clsCommon.myCdbl(txtVehicle_Capacity.Text)
                obj.Total_Item_Wt = clsCommon.myCdbl(txttotal_Wt.Text)
                obj.Gross_Item_Wt = clsCommon.myCdbl(txtGross_Wt.Text)
                obj.TransferIndent_No = clsCommon.myCstr(fndTransferIndentNo.Value)
                obj.Form38 = chkForm38.Checked
                obj.Ref_No = txtRefNo.Text
                obj.Total_Tax_Amt = clsCommon.myCdbl(lblTaxAmt.Text)
                obj.Remarks = txtRemarks.Text
                obj.From_Location = txtFromLocation.Value
                If clsCommon.myCBool(chkInternalTransfer.Checked) = True OrElse chkJobWorkTransfer.Checked Then
                    obj.To_Location = txtToLoc.Value
                Else
                    obj.To_Location = clsCommon.myCstr(lblToLoc.Tag)
                End If
                obj.Comments = txtComment.Text
                obj.On_Hold = chkOnHold.Checked
                obj.Mode_Of_Transport = cboModeOfTransport.Text
                obj.Electronic_Ref_No = txtElecttefNo.Text
                obj.Waybill_No = clsCommon.myCstr(txtEWayBillNo.Text)
                ''--- add by parteek 17-08-2017
                obj.For_Repair = IIf(chkForRepair.Checked, 1, 0)
                ''--end
                obj.InternalTransfer = IIf(chkInternalTransfer.Checked, 1, 0)
                obj.JobWorkTransfer = IIf(chkJobWorkTransfer.Checked, 1, 0)
                obj.ProdRequestTransfer = IIf(chkProductionRequest.Checked, 1, 0)
                obj.IsJobWorkType = IIf(chkJobWork.Checked, 1, 0)
                obj.Requisition_Id = clsCommon.myCstr(fndSRNO.Value)
                ''richa 03/03/2015
                If txtEWayBillDate.Enabled Then
                    obj.Waybill_Date = clsCommon.myCDate(txtEWayBillDate.Text)
                End If
                ''---------------------
                obj.GR_No = clsCommon.myCstr(txtGR_No.Text)
                If clsCommon.myLen(obj.GR_No) > 0 Then
                    obj.GR_Date = clsCommon.myCDate(dtpGR.Value)
                Else
                    obj.GR_Date = Nothing
                End If
                If chkOwnVehicle.Checked = True Then
                    obj.Transport_Id = ""
                    obj.Transporter_Name_Manual = clsCommon.myCstr(TxtTransportorMName.Text)
                Else
                    obj.Transporter_Name_Manual = ""
                    obj.Transport_Id = clsCommon.myCstr(txtTransporter_Code.Value)
                End If
                obj.Freight_Distance = clsCommon.myCdbl(txtFreightDistance.Value)
                obj.Tax_Group = txtTaxGroup.Value
                obj.Transfer_Type = clsCommon.myCstr(cboTransferType.SelectedValue)
                obj.RMDA_Code = txtRMDANo.Value
                ''richa agarwal 20 july
                obj.GP_Item_Type = clsCommon.myCstr(cmbGPItemType.SelectedValue)
                obj.Price_Code = clsCommon.myCstr(fndPriceCode.Value)
                ''-------------
                '==========added by preeti Gupta===================
                obj.Crate_IN = clsCommon.myCdbl(txtcrateout.Text)
                obj.jaali_IN = clsCommon.myCdbl(txtJaaliIn.Text)
                obj.Box_IN = clsCommon.myCdbl(txtBoxIn.Text)
                obj.Crate_Out = clsCommon.myCdbl(txtcrateout.Text)
                obj.jaali_Out = clsCommon.myCdbl(txtjaaliout.Text)
                obj.Box_Out = clsCommon.myCdbl(txtBoxOut.Text)
                '=======================================================
                obj.Is_Taxable = IIf(chkTaxable.Checked, 1, 0)
                obj.Is_MandiTax = IIf(chkIsManditax.Checked, 1, 0)
                If (gv2.Rows.Count > 0) Then
                    obj.TAX1 = clsCommon.myCstr(gv2.Rows(0).Cells(colTTaxAutCode).Value)
                    obj.TAX1_Rate = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxRate).Value)
                    obj.TAX1_Base_Amt = clsCommon.myCdbl(gv2.Rows(0).Cells(colTBaseAmt).Value)
                    obj.TAX1_Amt = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 1) Then
                    obj.TAX2 = clsCommon.myCstr(gv2.Rows(1).Cells(colTTaxAutCode).Value)
                    obj.TAX2_Rate = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxRate).Value)
                    obj.TAX2_Base_Amt = clsCommon.myCdbl(gv2.Rows(1).Cells(colTBaseAmt).Value)
                    obj.TAX2_Amt = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 2) Then
                    obj.TAX3 = clsCommon.myCstr(gv2.Rows(2).Cells(colTTaxAutCode).Value)
                    obj.TAX3_Rate = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxRate).Value)
                    obj.TAX3_Base_Amt = clsCommon.myCdbl(gv2.Rows(2).Cells(colTBaseAmt).Value)
                    obj.TAX3_Amt = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 3) Then
                    obj.TAX4 = clsCommon.myCstr(gv2.Rows(3).Cells(colTTaxAutCode).Value)
                    obj.TAX4_Rate = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxRate).Value)
                    obj.TAX4_Base_Amt = clsCommon.myCdbl(gv2.Rows(3).Cells(colTBaseAmt).Value)
                    obj.TAX4_Amt = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 4) Then
                    obj.TAX5 = clsCommon.myCstr(gv2.Rows(4).Cells(colTTaxAutCode).Value)
                    obj.TAX5_Rate = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxRate).Value)
                    obj.TAX5_Base_Amt = clsCommon.myCdbl(gv2.Rows(4).Cells(colTBaseAmt).Value)
                    obj.TAX5_Amt = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 5) Then
                    obj.TAX6 = clsCommon.myCstr(gv2.Rows(5).Cells(colTTaxAutCode).Value)
                    obj.TAX6_Rate = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxRate).Value)
                    obj.TAX6_Base_Amt = clsCommon.myCdbl(gv2.Rows(5).Cells(colTBaseAmt).Value)
                    obj.TAX6_Amt = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 6) Then
                    obj.TAX7 = clsCommon.myCstr(gv2.Rows(6).Cells(colTTaxAutCode).Value)
                    obj.TAX7_Rate = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxRate).Value)
                    obj.TAX7_Base_Amt = clsCommon.myCdbl(gv2.Rows(6).Cells(colTBaseAmt).Value)
                    obj.TAX7_Amt = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 7) Then
                    obj.TAX8 = clsCommon.myCstr(gv2.Rows(7).Cells(colTTaxAutCode).Value)
                    obj.TAX8_Rate = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxRate).Value)
                    obj.TAX8_Base_Amt = clsCommon.myCdbl(gv2.Rows(7).Cells(colTBaseAmt).Value)
                    obj.TAX8_Amt = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 8) Then
                    obj.TAX9 = clsCommon.myCstr(gv2.Rows(8).Cells(colTTaxAutCode).Value)
                    obj.TAX9_Rate = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxRate).Value)
                    obj.TAX9_Base_Amt = clsCommon.myCdbl(gv2.Rows(8).Cells(colTBaseAmt).Value)
                    obj.TAX9_Amt = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 9) Then
                    obj.TAX10 = clsCommon.myCstr(gv2.Rows(9).Cells(colTTaxAutCode).Value)
                    obj.TAX10_Rate = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxRate).Value)
                    obj.TAX10_Base_Amt = clsCommon.myCdbl(gv2.Rows(9).Cells(colTBaseAmt).Value)
                    obj.TAX10_Amt = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxAmt).Value)
                End If
                obj.Terms_Code = txtTermCode.Value
                obj.Terms_Remark = txtTermRemark.Text
                obj.Due_Date = txtDueDate.Value
                obj.Discount_Base = clsCommon.myCdbl(lblAmtWithoutTax.Text)
                obj.Total_Amt_Less_Tax = clsCommon.myCdbl(lblAmtWithoutTax.Text)
                obj.Discount_Amt = clsCommon.myCdbl(lblDiscountAmt.Text)
                obj.Amount_Less_Discount = clsCommon.myCdbl(lblAmtAfterDiscount.Text)
                obj.DOC_Total_Amt = clsCommon.myCdbl(lblTotRAmt.Text)
                obj.TransferOutNo = txtTransferOutNo.Value
                obj.Vehicle_Code = txtVehicleCode.Value
                obj.Vehicle_No = lblVehicleNo.Text
                obj.Km_Reading = txtKm.Text
                obj.Is_AgainstFormF = IIf(chkAgainst_Form.Checked, 1, 0)
                obj.Type = cboType.SelectedValue
                obj.Item_Tax_Type = Item_TaxType
                If rbtnTaxCalAutomatic.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                ElseIf rbtnTaxCalManual.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                End If
                '=======================added By preeti Gupta=================
                obj.Crate_IN = clsCommon.myCdbl(txtCrateIn.Text)
                obj.Box_IN = clsCommon.myCdbl(txtBoxIn.Text)
                obj.jaali_IN = clsCommon.myCdbl(txtJaaliIn.Text)
                obj.Crate_Out = clsCommon.myCdbl(txtcrateout.Text)
                obj.Box_Out = clsCommon.myCdbl(txtBoxOut.Text)
                obj.jaali_Out = clsCommon.myCdbl(txtjaaliout.Text)
                '================================================================
                obj.Arr = New List(Of clsTransferDCCDetail)
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim objTr As New clsTransferDCCDetail()
                    objTr.Line_No = clsCommon.myCdbl(grow.Cells(colLineNo).Value)
                    'objTr.Row_Type = clsCommon.myCstr(grow.Cells(colRowType).Value)
                    objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
                    ''richa agarwal 20/07/2016
                    objTr.GatePassNo = clsCommon.myCstr(grow.Cells(colgatePassTransferNo).Value)
                    ''----------------------
                    objTr.Item_Desc = clsCommon.myCstr(grow.Cells(colIName).Value)
                    objTr.Out_Qty = clsCommon.myCdbl(grow.Cells(colOutQty).Value)
                    objTr.In_Qty = clsCommon.myCdbl(grow.Cells(colInQty).Value)
                    objTr.Breakage = clsCommon.myCdbl(grow.Cells(colBreakQty).Value)
                    objTr.Leak = clsCommon.myCdbl(grow.Cells(colLeakQty).Value)
                    objTr.Shortage = clsCommon.myCdbl(grow.Cells(colShortQty).Value)
                    objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                    objTr.Alt_Unit_Code = clsCommon.myCstr(grow.Cells(colAltUOM).Value)
                    If clsCommon.myLen(objTr.Alt_Unit_Code) <= 0 Then
                        'objTr.Alt_Unit_Code = objTr.Unit_code
                    End If
                    objTr.TransferOutNo = clsCommon.myCstr(grow.Cells(colTransferOutNo).Value)
                    'objTr.Location = clsCommon.myCstr(grow.Cells(colloc).Value)
                    objTr.Item_Cost = clsCommon.myCdbl(grow.Cells(colRate).Value)
                    objTr.Amount = clsCommon.myCdbl(grow.Cells(colAmt).Value)
                    objTr.Disc_Per = clsCommon.myCdbl(grow.Cells(colDisPer).Value)
                    objTr.Disc_Amt = clsCommon.myCdbl(grow.Cells(colDisAmt).Value)
                    objTr.Abatement_Per = clsCommon.myCdbl(grow.Cells(colAbatementPer).Value)
                    objTr.Abatement_Amt = clsCommon.myCdbl(grow.Cells(colAbatementAmount).Value)
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(grow.Cells(colAmtAfterDis).Value)
                    objTr.TAX1 = clsCommon.myCstr(grow.Cells(colTax1).Value)
                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt1).Value)
                    objTr.TAX1_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate1).Value)
                    objTr.TAX1_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt1).Value)
                    objTr.TAX2 = clsCommon.myCstr(grow.Cells(colTax2).Value)
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt2).Value)
                    objTr.TAX2_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate2).Value)
                    objTr.TAX2_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt2).Value)
                    objTr.TAX3 = clsCommon.myCstr(grow.Cells(colTax3).Value)
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt3).Value)
                    objTr.TAX3_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate3).Value)
                    objTr.TAX3_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt3).Value)
                    objTr.TAX4 = clsCommon.myCstr(grow.Cells(colTax4).Value)
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt4).Value)
                    objTr.TAX4_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate4).Value)
                    objTr.TAX4_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt4).Value)
                    objTr.TAX5 = clsCommon.myCstr(grow.Cells(colTax5).Value)
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt5).Value)
                    objTr.TAX5_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate5).Value)
                    objTr.TAX5_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt5).Value)
                    objTr.TAX6 = clsCommon.myCstr(grow.Cells(colTax6).Value)
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt6).Value)
                    objTr.TAX6_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate6).Value)
                    objTr.TAX6_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt6).Value)
                    objTr.TAX7 = clsCommon.myCstr(grow.Cells(colTax7).Value)
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt7).Value)
                    objTr.TAX7_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate7).Value)
                    objTr.TAX7_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt7).Value)
                    objTr.TAX8 = clsCommon.myCstr(grow.Cells(colTax8).Value)
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt8).Value)
                    objTr.TAX8_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate8).Value)
                    objTr.TAX8_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt8).Value)
                    objTr.TAX9 = clsCommon.myCstr(grow.Cells(colTax9).Value)
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt9).Value)
                    objTr.TAX9_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate9).Value)
                    objTr.TAX9_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt9).Value)
                    objTr.TAX10 = clsCommon.myCstr(grow.Cells(colTax10).Value)
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt10).Value)
                    objTr.TAX10_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate10).Value)
                    objTr.TAX10_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt10).Value)
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(grow.Cells(colTotTaxAmt).Value)
                    objTr.Item_Net_Amt = clsCommon.myCdbl(grow.Cells(colAmtAfterTax).Value)
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(grow.Cells(colItemwiseTaxCode).Value)
                    objTr.Specification = clsCommon.myCstr(grow.Cells(colSpecification).Value)
                    objTr.Remarks = clsCommon.myCstr(grow.Cells(colRemarks).Value)
                    objTr.Location = txtToLoc.Value 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)
                    objTr.MRP = clsCommon.myCdbl(grow.Cells(colMRP).Value)
                    ''objTr.Assessable = clsCommon.myCdbl(grow.Cells(colAssessableRate).Value)
                    objTr.Item_Unit_Wt = clsCommon.myCdbl(grow.Cells(colIUnitWt).Value)
                    objTr.Item_Net_Wt = clsCommon.myCdbl(grow.Cells(colINetWt).Value)
                    objTr.Item_Net_MT_Wt = clsCommon.myCdbl(grow.Cells(colINetMTWt).Value)
                    objTr.Bin_No = clsCommon.myCstr(grow.Cells(colBinNo).Value)
                    objTr.arrSrItem = TryCast(grow.Tag, List(Of clsSerializeInvenotry))
                    objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                        obj.Arr.Add(objTr)
                    End If
                Next
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Item", Me.Text)
                    Return False
                End If
                ''For Custom Fields
                obj.Form_ID = MyBase.Form_ID
                obj.arrCustomFields = New List(Of clsCustomFieldValues)
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.GetData(obj.arrCustomFields)
                End If
                If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
                    clsCustomFieldGrid.GetData(obj.arrCustomFields, gv1, MyBase.ArrDetailFields, colICode)
                End If
                ''End of For Custom Fields
                '' CurrencConversion
                If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
                    obj.CURRENCY_CODE = Me.txtCurrencyCode.Value
                    obj.ConvRate = clsCommon.myCdbl(Me.txtConversionRate.Text)
                    If clsCommon.myLen(txtApplicableFrom.Text) > 0 Then
                        obj.ApplicableFrom = Me.txtApplicableFrom.Text
                    Else
                        obj.ApplicableFrom = Nothing
                    End If
                Else
                    obj.CURRENCY_CODE = Nothing
                    obj.ConvRate = 1
                    obj.ApplicableFrom = Nothing
                End If
                '' end CurrencyConversion
                Dim isSaved As Boolean = obj.SaveData(obj, isNewEntry, isDoAbandomentNo)
                UcAttachment1.SaveData(obj.Document_No)
                'If chkInternalTransfer.Checked Then   'Asked By Shruti mam
                If ChekBtnPost = False Then
                    If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(1) as cc from TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL where  TRANS_Code='" + MyBase.Form_ID + "'  and Document_Code='" + obj.Document_No + "'")) <= 0 Then
                        Dim xNewDesc As String = "Description : " + obj.Description
                        clsApply_Approval.CheckApprovalRequired(MyBase.Form_ID, obj.Document_No, txtDate.Value, xNewDesc, txtRemarks.Text, clsCommon.myCdbl(lblTotRAmt.Text), 0, "")
                    End If
                End If
                'End If
                LoadData(obj.Document_No, NavigatorType.Current)
                Return isSaved
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Function
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            btnSave.Enabled = True
            btnPost.Enabled = True
            btnDelete.Enabled = True
            isInsideLoadData = True
            IsFormLoad = True
            isNewEntry = False
            btnSave.Text = "Update"
            BlankAllControls()
            LoadBlankGrid()
            LoadBlankGridTax()
            cboTransferType.Enabled = False
            txtFromLocation.Enabled = False
            txtToLoc.Enabled = False
            Dim obj As New clsTransferDCC()
            obj = clsTransferDCC.GetData(strCode, NavTyep)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                'txtTaxGroup.Enabled = False
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                    repoComplete.IsVisible = True
                    'repoBalQty.IsVisible = True
                    btn_CancelDel.Enabled = True
                    chkEwaybill.Enabled = False
                ElseIf obj.Status = ERPTransactionStatus.Pending Then
                    btn_CancelDel.Enabled = False
                    chkEwaybill.Enabled = True
                End If
                UsLock1.Status = obj.Status
                txtDocNo.Value = obj.Document_No
                txtDate.Value = obj.Document_Date
                If clsCommon.myLen(obj.Delivery_date) > 0 Then
                    txtDeliveryDate.Value = obj.Delivery_date
                    txtOutDate.Value = obj.Delivery_date
                End If
                txtglvoucher.Text = obj.GLVoucher_No
                txtDeliveryDuration.Text = obj.Delivery_Duration
                chkForm38.Checked = obj.Form38
                txtRefNo.Text = obj.Ref_No
                cboTransferType.SelectedValue = obj.Transfer_Type
                chkAgainst_Form.Checked = obj.Is_AgainstFormF
                chkOnHold.Checked = obj.On_Hold
                txtSecondary_Doc_Code.Text = obj.Secondary_Code
                txtvehicle_mannual_no.Text = obj.Vehicle_Mannual_No
                txtVehicle_Capacity.Text = obj.Vehicle_Capacity
                txtGross_Wt.Text = obj.Gross_Item_Wt
                txtvehicle_Charge.Text = obj.Vehicle_Charge
                txttotal_Wt.Text = obj.Total_Item_Wt
                chkEwaybill.Checked = IIf(obj.IsEwayBill = 1, True, False)
                '===================added by preeti===
                txtEWayBillNo.Text = obj.EWayBillNo
                txtElecttefNo.Text = obj.Electronic_Ref_No
                If obj.EWayBillDate IsNot Nothing Then
                    txtEWayBillDate.Value = obj.EWayBillDate
                    txtEWayBillDate.Checked = True
                End If
                fndTransferIndentNo.Value = obj.TransferIndent_No
                txtLoadingAdviceNo.Text = obj.Loading_Advice_No
                chkTaxable.Checked = IIf(obj.Is_Taxable = 1, True, False)
                chkIsManditax.Checked = IIf(obj.Is_MandiTax = 1, True, False)
                chkInternalTransfer.Checked = IIf(obj.InternalTransfer = 1, True, False)
                chkJobWorkTransfer.Checked = IIf(obj.JobWorkTransfer = 1, True, False)
                chkProductionRequest.Checked = IIf(obj.ProdRequestTransfer = 1, True, False)
                If chkInternalTransfer.Checked Then
                    BtnPrintChallan.Enabled = False
                    rbtnPrintExcisable.Enabled = False
                Else
                    BtnPrintChallan.Enabled = True
                    rbtnPrintExcisable.Enabled = True
                End If
                fndSRNO.Value = obj.Requisition_Id
                fndSRNO.Enabled = False
                '=======================================
                chkJobWork.Checked = IIf(obj.IsJobWorkType = 1, True, False)
                ''======================================================================
                If clsCommon.myCdbl(obj.Vehicle_Charge) > 0 Then
                    Dim dt As New DataTable
                    dt.Columns.Add("FixedCharge", GetType(Decimal))
                    dt.Columns.Add("EmptyCharge", GetType(Decimal))
                    dt.Columns.Add("FreightCharge", GetType(Decimal))
                    dt.Columns.Add("FreightType", GetType(String))
                    Dim dr As DataRow = dt.NewRow()
                    dr("FreightType") = obj.Freight_Type
                    dr("FixedCharge") = obj.FixedCharge
                    dr("FreightCharge") = obj.Vehicle_Charge
                    dr("EmptyCharge") = obj.EmptyCharge
                    txtvehicle_Charge.Tag = dt
                Else
                    txtvehicle_Charge.Tag = Nothing
                End If
                ''======================================================================
                txtTaxGroup.Value = obj.Tax_Group
                PreviousTaxGroupCode = obj.Tax_Group
                txtComment.Text = obj.Comments
                If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then 'Added by preeti Gupta[In Case Of Return]
                    pnlLoadIn.Visible = True
                    txtCrateIn.Visible = True
                    txtJaaliIn.Visible = True
                    txtBoxIn.Visible = True
                    lblCrateIn.Visible = True
                    lblJaaliIn.Visible = True
                    lblBoxIn.Visible = True
                    txtTaxGroup.Enabled = False
                    txtToLoc.Value = obj.To_Location
                    txtTransferOutNo.Value = obj.TransferOutNo
                    Dim qry = "Select EWayBillNo,Electronic_Ref_No,EWayBillDate from TSPL_Transfer_ORDER_Head where Document_No='" & txtTransferOutNo.Value & "'"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                        If dt.Rows(0)("EWayBillDate") IsNot DBNull.Value Then
                            txtEWayBillDate.Value = clsCommon.myCDate(dt.Rows(0)("EWayBillDate"))
                            txtEWayBillDate.Checked = True
                        End If
                        txtEWayBillNo.Text = clsCommon.myCstr(dt.Rows(0)("EWayBillNo"))
                        txtElecttefNo.Text = clsCommon.myCstr(dt.Rows(0)("Electronic_Ref_No"))
                    End If
                    txtTransferOutNo.Enabled = False
                Else
                    If clsCommon.myCBool(chkInternalTransfer.Checked) = True Then
                        txtToLoc.Value = obj.To_Location
                        chkInternalTransfer.Enabled = False
                        chkProductionRequest.Enabled = False
                    Else
                        txtToLoc.Value = obj.To_Location_Main
                    End If
                    txtTaxGroup.Enabled = True
                End If
                '====Enabled False of Repair Check Box ==============
                chkForRepair.Checked = IIf(obj.For_Repair = 1, True, False)
                CheckRepair = IIf(obj.For_Repair = 1, True, False)
                If chkTaxable.Checked = False AndAlso chkForRepair.Checked = True Then
                    chkForRepair.Enabled = False
                ElseIf chkTaxable.Checked = True AndAlso chkForRepair.Checked = False Then
                    chkForRepair.Enabled = False
                Else
                    chkForRepair.Enabled = True
                End If
                '=====================
                FillLocationInfo(isInsideLoadData)
                txtFromLocation.Value = obj.From_Location
                txtRemarks.Text = obj.Remarks
                chkOwnVehicle.Checked = False
                If clsCommon.myLen(obj.Transport_Id) <= 0 Then
                    chkOwnVehicle.Checked = True
                    TxtTransportorMName.Text = clsCommon.myCstr(obj.Transporter_Name_Manual)
                    txtTransporter_Code.Visible = False
                Else
                    TxtTransportorMName.Text = ""
                End If
                cboModeOfTransport.Text = obj.Mode_Of_Transport
                txtTransporter_Code.Value = obj.Transport_Id
                txtTransporter_desc.Text = clsTransferDCC.GetTransporterName(txtTransporter_Code.Value)
                txtFreightDistance.Value = obj.Freight_Distance
                txtGR_No.Text = obj.GR_No
                dtpGR.Value = obj.GR_Date
                txtWayBill_No.Text = obj.Waybill_No
                ''richa 03/03/2015
                If clsCommon.myLen(obj.Waybill_Date) > 0 Then
                    ttxway_bill_date.Value = obj.Waybill_Date
                End If
                ''----------------------
                ''richa agarwal 20 july
                cmbGPItemType.SelectedValue = obj.GP_Item_Type
                fndPriceCode.Value = obj.Price_Code
                ''-------------
                Dim objTaxGrpMaster As New clsTaxGroupMaster()
                objTaxGrpMaster = clsTransferDCC.GetTaxGroupData(obj.Tax_Group, "T")
                If (objTaxGrpMaster IsNot Nothing) Then
                    lblTaxGrpName.Text = objTaxGrpMaster.Tax_Group_Desc
                End If
                txtRMDANo.Value = obj.RMDA_Code
                cboTransferType.SelectedValue = obj.Transfer_Type
                txtVehicleCode.Value = obj.Vehicle_Code
                lblVehicleNo.Text = obj.Vehicle_No
                txtTermCode.Value = obj.Terms_Code
                txtTermRemark.Text = obj.Terms_Remark
                'lblTermName.Text = obj.Terms_Description
                If clsCommon.myLen(obj.Due_Date) > 0 Then
                    txtDueDate.Value = obj.Due_Date
                End If
                lblAmtWithoutTax.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amt)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax_Amt)
                lblTotRAmt.Text = clsCommon.myFormat(obj.DOC_Total_Amt)
                lblTotRAmt1.Text = lblTotRAmt.Text
                lblFromLoc.Text = obj.From_LocationName
                lblTaxGrpName.Text = obj.TaxGroupName
                lblTermName.Text = obj.TermsName
                cboType.SelectedValue = obj.Type
                '=======================added By preeti Gupta=================
                txtCrateIn.Text = clsCommon.myCdbl(obj.Crate_IN)
                txtBoxIn.Text = clsCommon.myCdbl(obj.Box_IN)
                txtJaaliIn.Text = clsCommon.myCdbl(obj.jaali_IN)
                txtcrateout.Text = clsCommon.myCdbl(obj.Crate_Out)
                txtBoxOut.Text = clsCommon.myCdbl(obj.Box_Out)
                txtjaaliout.Text = clsCommon.myCdbl(obj.jaali_Out)
                '================================================================
                EInvoiceType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_TRANSFER_ORDER_HEAD", "Document_No", txtDocNo.Value, Nothing)
                If chkJobWork.Checked = False AndAlso (clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal) AndAlso chkTaxable.Checked = True AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                    btnReverseAndUnpost.Enabled = False
                    If objCommonVar.GenerateEWayBillWithEInvoice = True Then
                        txtEWayBillNo.ReadOnly = True
                        txtEWayBillDate.ReadOnly = True
                    End If
                    If obj.Status = ERPTransactionStatus.Approved Then
                        btn_CancelDel.Enabled = True
                    ElseIf obj.Status = ERPTransactionStatus.Pending Then
                        btn_CancelDel.Enabled = False
                    End If
                End If
                If clsCommon.myLen(txtFromLocation.Value) > 0 AndAlso clsCommon.myLen(txtToLoc.Value) > 0 Then
                    strTaxType = clsLocationWiseTax.TaxType(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value, Nothing)
                End If
                If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                    IsMandiTax = clsLocationWiseTax.IsMandiTax(txtTaxGroup.Value, Nothing)
                End If
                If (clsCommon.myLen(obj.TAX1) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX1
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX1_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX1_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX1_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX1) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX2) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX2
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX2_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX2_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX2_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX2_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX2) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX3) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX3
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX3_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX3_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX3_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX3_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX3) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX4) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX4
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX4_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX4_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX4_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX4_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX4) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX5) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX5
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX5_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX5_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX5_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX5_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX5) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX6) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX6
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX6_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX6_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX6_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX6_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX6) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX7) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX7
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX7_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX7_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX7_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX7_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX7) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX8) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX8
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX8_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX8_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX8_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX8_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX8) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX9) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX9
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX9_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX9_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX9_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX9_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX9) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX10) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX10
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX10_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX10_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX10_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAssessableAmt).Value = obj.TAX10_Base_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX10) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic Then
                    rbtnTaxCalAutomatic.IsChecked = True
                ElseIf obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual Then
                    rbtnTaxCalManual.IsChecked = True
                End If
                txtKm.Text = obj.Km_Reading
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsTransferDCCDetail In obj.Arr
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Tag = objTr.arrSrItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = objTr.arrBatchItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = objTr.Line_No
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colComplete).Value = IIf(objTr.Status = 0, "No", "Yes")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colgatePassTransferNo).Value = objTr.GatePassNo
                        FndGatePassNo.Value = objTr.GatePassNo
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsPickAutoSrNo).Value = clsItemMaster.IsPickAutoSerializeItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = objTr.Out_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = objTr.In_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBreakQty).Value = objTr.Breakage
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLeakQty).Value = objTr.Leak
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colShortQty).Value = objTr.Shortage
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = objTr.Alt_Unit_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = objTr.TransferOutNo
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = True
                        Else
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = False
                        End If
                        '==================Added by preeti Gupta Against ticket no[UDL/24/07/18-000207]===========
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).ReadOnly = True
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemwiseTaxCode).Value = objTr.ItemwiseTaxCode
                        '========================================================================
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = objTr.Location
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = objTr.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = objTr.Disc_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisAmt).Value = objTr.Disc_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = objTr.Abatement_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementAmount).Value = objTr.Abatement_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterDis).Value = objTr.Amt_Less_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax1).Value = objTr.TAX1
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt1).Value = objTr.TAX1_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = objTr.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt1).Value = objTr.TAX1_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax2).Value = objTr.TAX2
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt2).Value = objTr.TAX2_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = objTr.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt2).Value = objTr.TAX2_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax3).Value = objTr.TAX3
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt3).Value = objTr.TAX3_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = objTr.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt3).Value = objTr.TAX3_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax4).Value = objTr.TAX4
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt4).Value = objTr.TAX4_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = objTr.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt4).Value = objTr.TAX4_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax5).Value = objTr.TAX5
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt5).Value = objTr.TAX5_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = objTr.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt5).Value = objTr.TAX5_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax6).Value = objTr.TAX6
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt6).Value = objTr.TAX6_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = objTr.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt6).Value = objTr.TAX6_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax7).Value = objTr.TAX7
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt7).Value = objTr.TAX7_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = objTr.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt7).Value = objTr.TAX7_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax8).Value = objTr.TAX8
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt8).Value = objTr.TAX8_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = objTr.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt8).Value = objTr.TAX8_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax9).Value = objTr.TAX9
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt9).Value = objTr.TAX9_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = objTr.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt9).Value = objTr.TAX9_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax10).Value = objTr.TAX10
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt10).Value = objTr.TAX10_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = objTr.TAX10_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt10).Value = objTr.TAX10_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Total_Tax_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterTax).Value = objTr.Item_Net_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSpecification).Value = objTr.Specification
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRemarks).Value = objTr.Remarks
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = objTr.MRP
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = objTr.Item_Unit_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = objTr.Item_Net_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = objTr.Item_Net_MT_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = objTr.Bin_No
                        Dim StrCrateTransferFromBooking As String = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                        'If StrCrateTransferFromBooking = "1" Then
                        '    gv1.Rows(gv1.Rows.Count - 1).Cells(colFOCItem).Value = objTr.FOCItem
                        'End If
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableRate).Value = objTr.Assessable
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableAmount).Value = objTr.AssessableAmt
                        'If clsCommon.myLen(objTr.TransferOutNo) > 0 Then
                        '    gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = clsRequistionDetail.GetBalanceRequitionQty(objTr.TransferOutNo, objTr.Item_Code, obj.Document_No, "")
                        'End If
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                            gv1.Columns(colInQty).IsVisible = False
                            gv1.Columns(colLeakQty).IsVisible = False
                            gv1.Columns(colBreakQty).IsVisible = False
                            gv1.Columns(colShortQty).IsVisible = False
                        Else
                            gv1.Columns(colInQty).IsVisible = True
                            gv1.Columns(colLeakQty).IsVisible = True
                            gv1.Columns(colBreakQty).IsVisible = True
                            gv1.Columns(colShortQty).IsVisible = True
                        End If
                    Next
                    If obj.Status = ERPTransactionStatus.Approved Then
                        btnAmendment.Enabled = True
                    End If
                    If obj.Status = ERPTransactionStatus.Pending Then
                        If clsCommon.myLen(fndTransferIndentNo.Value) <= 0 Then
                            gv1.Rows.AddNew()
                        End If
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                    End If
                    ''RICHA UDL/14/05/18-000161
                    If obj.Status = ERPTransactionStatus.Cancel Then
                        btnSave.Enabled = False
                        btnDelete.Enabled = False
                        btnPost.Enabled = False
                    End If
                    'gv1.Rows.AddNew()
                End If
                SetitemWiseTaxOnlySetting()
                ''RefreshReqNo()
                ''For Custom Fields
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.LoadData(obj.Document_No)
                End If
                clsCustomFieldGrid.FillDataInGrid(obj.Document_No, MyBase.Form_ID, gv1)
                ''End of For Custom Fields
                '' MULTICURRENCY
                Me.txtCurrencyCode.Value = obj.CURRENCY_CODE
                Me.txtConversionRate.Text = obj.ConvRate
                Me.txtApplicableFrom.Text = obj.ApplicableFrom.ToString
                '' end  MULTICURRENCY
                UcAttachment1.LoadData(obj.Document_No)
                If clsCommon.CompairString(clsCommon.myCstr(cboTransferType.SelectedValue), "I") = CompairStringResult.Equal Then
                    btnPrintNew.Enabled = False
                    BtnPrintChallan.Enabled = False
                Else
                    btnPrintNew.Enabled = True
                    BtnPrintChallan.Enabled = True
                End If
                'If chkInternalTransfer.Checked Then
                'If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                '    If Not clsApply_Approval.Visibility_PostButtonForApproval(MyBase.Form_ID, txtDocNo.Value, clsCommon.myCdbl(lblTotRAmt.Text), 0, "") Then
                '        btnPost.Visible = False
                '        If UsLock1.Status = ERPTransactionStatus.Pending Then
                '            UsLock1.Status = clsApply_Approval.ApprovalCondCheck_Doc(MyBase.Form_ID, txtDocNo.Value, Nothing)
                '        End If
                '    End If
                'End If
            Else
                AddNew()
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
            IsFormLoad = False
        End Try
    End Sub
    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CloseForm()
    End Sub
    Sub CloseForm()
        Me.Close()
    End Sub
    Private Sub gv2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv2.DoubleClick
        Try
            If rbtnTaxCalAutomatic.IsChecked Then
                '' done by priti on 22/11/2017
                If Not clsCommon.CompairString(cboTransferType.Text, "Transfer In") = CompairStringResult.Equal Then
                    If CalculateTaxRatefromItemwsieTaxOnSale = 0 Then
                        'Dim qry As String = "select Tax_Rate_Code as [Rate Code],Tax_Rate_Desc as [Rate Description],Tax_Rate as [Rate] from TSPL_TAX_RATES inner join TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_RATES.Tax_Code "
                        Dim dblNewRate As Double = clsLocationWiseTax.FinderForTaxRateForTransfer(txtFromLocation.Value, txtTaxGroup.Value, clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value), txtToLoc.Value, "T") 'clsCommon.myCdbl(clsCommon.ShowSelectForm("FndTaxfnd", qry, "Rate", "TSPL_TAX_RATES.Tax_Code='" + clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value) + "' and Tax_Type='T' and TSPL_TAX_MASTER.type='E'", "", "", True))
                        Dim intRowNo As Integer = gv2.CurrentRow.Index
                        If gv1.RowCount > 0 AndAlso intRowNo >= 0 Then
                            Dim strII As String = clsCommon.myCstr(intRowNo + 1)
                            For ii As Integer = 0 To gv1.Rows.Count - 1
                                gv1.Rows(ii).Cells("COLTAXRATE" + strII).Value = dblNewRate
                            Next
                        End If
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            UpdateCurrentRow(ii)
                        Next
                        UpdateAllTotals()
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        PostData()
    End Sub
    Private Function PostAllowToSave() As Boolean
        Try
            If Not ProvisionAllow Then
                Return True
            End If
            If clsCommon.CompairString(cboModeOfTransport.Text, "By Hand") <> CompairStringResult.Equal Then
                If clsCommon.myLen(txtTransporter_Code.Value) <= 0 AndAlso chkOwnVehicle.Checked = False Then
                    RadPageView1.SelectedPage = RadPageViewPage1
                    txtTransporter_Code.Focus()
                    txtTransporter_Code.Select()
                    clsCommon.MyMessageBoxShow(Me, "Select transporter detail.", Me.Text)
                    Exit Function
                End If
                If clsCommon.myLen(txtTransporter_Code.Value) > 0 AndAlso clsCommon.myLen(txtvehicle_mannual_no.Text) <= 0 Then
                    RadPageView1.SelectedPage = RadPageViewPage1
                    txtvehicle_mannual_no.Focus()
                    txtvehicle_mannual_no.Select()
                    clsCommon.MyMessageBoxShow(Me, "Fill vehicle no. for provision booking.", Me.Text)
                    Exit Function
                End If
                If clsCommon.myLen(txtTransporter_Code.Value) > 0 AndAlso clsCommon.myCdbl(txtVehicle_Capacity.Text) <= 0 Then
                    RadPageView1.SelectedPage = RadPageViewPage1
                    txtVehicle_Capacity.Focus()
                    txtVehicle_Capacity.Select()
                    clsCommon.MyMessageBoxShow(Me, "Fill vehicle capacity for provision booking.", Me.Text)
                    Exit Function
                End If
            End If
            If clsCommon.myLen(txtTransporter_Code.Value) > 0 AndAlso clsCommon.myCdbl(txtGross_Wt.Text) <= 0 Then
                RadPageView1.SelectedPage = RadPageViewPage1
                txtGross_Wt.Focus()
                txtGross_Wt.Select()
                clsCommon.MyMessageBoxShow(Me, "Fill Gross weight for provision booking.", Me.Text)
                Exit Function
            End If
            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Function
    Sub PostData()
        Try
            If (myMessages.postConfirm()) Then
                If clsCommon.myCBool(chkInternalTransfer.Checked) = False Then
                    If Not PostAllowToSave() Then
                        Exit Sub
                    End If
                End If
                If SavingData(True) Then
                    If ApplyFEFO = True AndAlso chkInternalTransfer.Checked = True AndAlso clsCommon.CompairString(cboTransferType.Text, "Transfer In") = CompairStringResult.Equal Then
                        Qry = "select code from TSPL_ATTACHMENTS where TransactionId ='" + clsCommon.myCstr(txtDocNo.Value) + "' and FormId ='STO-TRANSFER'"
                        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(Qry)
                        If dt1 IsNot Nothing AndAlso dt1.Rows.Count <= 0 Then
                            Throw New Exception("Plz attach the attachment,it is mandatory")
                        End If
                    End If
                    If (clsTransferDCC.postTransfer(txtDocNo.Value, ProvisionAllow, txtglvoucher.Text)) Then
                        common.clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                        LoadData(txtDocNo.Value, NavigatorType.Current)
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        DeleteData()
    End Sub
    Sub DeleteData()
        Try
            Dim Reason As String = ""
            If (myMessages.deleteConfirm()) Then
                If clsCancelLog.CheckForReasonOnDelete() Then
                    If chkInternalTransfer.Checked Then
                        clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, txtDocNo.Value)
                    End If
                    '' REASON FOR DELETE 
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Delete"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If
                End If
                If (clsTransferDCC.DeleteData(txtDocNo.Value)) Then
                    saveCancelLog(Reason, "Delete", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub
    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.txtDocNo.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
    End Function
    Private Sub gv1_CellEditorInitialized(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellEditorInitialized
        If TypeOf Me.gv1.CurrentColumn Is GridViewMultiComboBoxColumn Then
            Dim editor As RadMultiColumnComboBoxElement = DirectCast(Me.gv1.ActiveEditor, RadMultiColumnComboBoxElement)
            editor.AutoSizeDropDownToBestFit = True
            editor.EditorControl.MasterTemplate.BestFitColumns()
            editor.DropDownStyle = RadDropDownStyle.DropDown
            editor.AutoFilter = True
            If editor.EditorControl.MasterTemplate.FilterDescriptors.Count = 0 Then
                Dim autoFilter As FilterDescriptor = New FilterDescriptor("Name", FilterOperator.StartsWith, "")
                autoFilter.IsFilterEditor = True
                editor.EditorControl.FilterDescriptors.Add(autoFilter)
            End If
        End If
    End Sub
    Private Sub txtDocNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim qst As String = "select count(*) from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + txtDocNo.Value + "'"
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
    Private Sub txtDocNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtDocNo._MYValidating
        Dim qry As String = "select Document_No as [TransferNO],convert (varchar(10), Document_Date,103) as Date,DOC_Total_Amt as Amount,case when Status='0' then 'Pending' else 'Approved' end as [Status], From_Location+' - '+FromLocation.Location_Desc as [FromLocation], To_Location+' - '+ToLocation.Location_Desc as [ToLocation], Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then 'LOAD OUT' When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' Then 'LOAD IN' When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='T' Then 'Return' Else 'REJECTED' End as [Transfer Type], TSPL_TRANSFER_ORDER_HEAD.TransferOutNo,TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_TRANSFER_ORDER_HEAD.Price_Code as [Price Code] from TSPL_TRANSFER_ORDER_HEAD LEFT OUTER JOIN TSPL_LOCATION_MASTER FromLocation ON FromLocation.Location_Code=TSPL_TRANSFER_ORDER_HEAD.From_Location LEFT OUTER JOIN TSPL_LOCATION_MASTER ToLocation on ToLocation.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location"
        Dim whrClas As String = ""
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrClas = " (case when transfer_type='O' then from_location else to_location end) in (" + arrLoc + ")"
        End If
        LoadData(clsCommon.ShowSelectForm("POOrde123d", qry, "TransferNO", whrClas, txtDocNo.Value, "TSPL_TRANSFER_ORDER_HEAD.Document_Date desc", isButtonClicked, "TSPL_TRANSFER_ORDER_HEAD.Document_Date"), NavigatorType.Current)
    End Sub
    Private Sub FrmAPInvoiceEntry_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing AndAlso clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) <= 0 Then
            isCellValueChangedOpen = True
            If gv1.CurrentColumn Is gv1.Columns(colICode) Then
                'gv1.CurrentColumn = gv1.Columns(colIName)
                OpenICodeList(True)
                'gv1.CurrentColumn = gv1.Columns(colICode)
            End If
            ''setGridFocus()
            isCellValueChangedOpen = False
        ElseIf e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing AndAlso gv1.CurrentColumn Is gv1.Columns(colAltUOM) Then
            isCellValueChangedOpen = True
            OpenAltUOMList(True)
            isCellValueChangedOpen = False
        ElseIf e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag Then
            SavingData(False)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso MyBase.isPostFlag Then
            PostData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
            CloseForm()
        ElseIf e.Control AndAlso e.KeyCode = Keys.F7 AndAlso e.Alt = False Then
            SelectRequistionItems()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                                                      "TSPL_Transfer_ORDER_HEAD " + Environment.NewLine +
                                                      "TSPL_TRANSFER_ORDER_DETAIL " + Environment.NewLine +
                                                      "TSPL_BATCH_ITEM ( If Item is batch type) " + Environment.NewLine +
                                                      "TSPL_SERIAL_ITEM ( If Item is Serial type)" + Environment.NewLine +
                                                      "TSPL_PROVISION_ENTRY (If Provision Setting is On , Transfer type In and Transport Id entered, After Posting )  " + Environment.NewLine +
                                                      "TSPL_JOURNAL_MASTER (Journal Voucher Entry - After Posting )  " + Environment.NewLine +
                                                      "TSPL_JOURNAL_DETAILS  (After Posting) " + Environment.NewLine +
                                                      "TSPL_INVENTORY_MOVEMENT  - After Posting ")
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIRC
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverseAndUnpost.Visible = True
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                'MessageBox.Show("You are not authorized to perform this action.", "Unauthorized Access", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F8 Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIRC
            frm.strCode = clsFixedParameterCode.SIReversAndCreate
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                btnReverseAndUnpost.Visible = True
                blnUpdateLoadInwithLoadOut = True
            End If
        ElseIf e.KeyCode = Keys.F5 Then
            '======update by preeti gupta 16/10/2018
            If ApplyFEFO = True And chkInternalTransfer.Checked = True Then
                OpenBatchItemIfFIFIOSettingON()
            ElseIf RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                OpenBatchItem()
            Else
                OpenBatchItemIfFIFIOSettingON()
            End If
        End If
    End Sub
    Private Sub txtTermCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTermCode._MYValidating
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER"
        txtTermCode.Value = clsCommon.ShowSelectForm("PoermCodefnd", qry, "Code", "", txtTermCode.Value, "Code", isButtonClicked)
        SetTermDetails()
    End Sub
    Sub SetTermDetails()
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER where Terms_Code='" + txtTermCode.Value + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Description"))
            txtDueDate.Value = txtDate.Value.AddDays(clsCommon.myCdbl(dt.Rows(0)("No Of Days")))
        Else
            lblTermName.Text = ""
        End If
    End Sub
    Private Sub txtTaxGroup__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTaxGroup._MYValidating
        Try
            If clsCommon.myCBool(chkInternalTransfer.Checked) = True Then
                txtTaxGroup.Value = clsDBFuncationality.getSingleValue("select Tax_Group_Code from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='EXEMPTED' and tax_group_type='S'", Nothing)
                txtTaxGroup.Enabled = False
            Else
                GstStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
                If clsCommon.myLen(txtFromLocation.Value) > 0 AndAlso clsCommon.myLen(txtToLoc.Value) > 0 Then
                    strTaxType = clsLocationWiseTax.TaxType(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value, Nothing)
                    If GstStatus = True AndAlso clsCommon.CompairString(strTaxType, "L") = CompairStringResult.Equal Then
                        txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroupForTransfer(txtFromLocation.Value, txtToLoc.Value, "T", txtTaxGroup.Value, isButtonClicked)
                    Else
                        txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroupForTransfer(txtFromLocation.Value, txtToLoc.Value, "T", txtTaxGroup.Value, isButtonClicked)
                    End If
                Else
                    txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroupForTransfer(txtFromLocation.Value, txtToLoc.Value, "T", txtTaxGroup.Value, isButtonClicked)
                End If
                If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                    IsMandiTax = clsLocationWiseTax.IsMandiTax(txtTaxGroup.Value, Nothing)
                End If
                txtTaxGroup.Enabled = True
            End If
            SetTaxDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub SetTaxDetails()
        LoadBlankGridTax()
        Dim dt As New DataTable
        If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
            If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                Dim FromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & txtTransferOutNo.Value & "'"))
                dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, FromLoc)
            Else
                dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, txtFromLocation.Value)
            End If
        Else
            dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, txtFromLocation.Value)
        End If
        ''Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,Tax_Code,Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,Taxable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + txtTaxGroup.Value + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"
        'Dim dt As DataTable = clsTransferDCC.GetTaxDetailsTransfer(txtTaxGroup.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If (dt.Rows.Count > 10) Then
                MessageBox.Show(Me, "Can't Handle More than 10 Tax Types in a Group")
                Return
            End If
            lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            For Each dr As DataRow In dt.Rows
                gv2.Rows.AddNew()
                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dr("Tax_Code"))
                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = clsCommon.myCstr(dr("Tax_Code_Desc"))
                If rbtnTaxCalAutomatic.IsChecked Then
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
                End If
            Next
            If CalculateTaxRatefromItemwsieTaxOnSale = 1 Then
                SetitemWiseTaxSetting(True, False)
            End If
        Else
            lblTaxGrpName.Text = ""
            For ii As Integer = 0 To gv1.Rows.Count - 1
                BlankTaxDetails(ii)
            Next
        End If
        For ii As Integer = 0 To gv1.Rows.Count - 1
            UpdateCurrentRow(ii)
        Next
        UpdateAllTotals()
    End Sub
    Sub SetitemWiseTaxSetting(ByVal isChangeRate As Boolean, ByVal isForCurrentRow As Boolean)
        ''Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,Tax_Code,Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,Taxable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + txtTaxGroup.Value + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"
        'Dim dt As DataTable = clsTransferDCC.GetTaxDetailsTransfer(txtTaxGroup.Value)
        Dim dt As New DataTable
        If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
            If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                Dim FromLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select From_Location from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & txtTransferOutNo.Value & "'"))
                dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, FromLoc)
            Else
                dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, txtFromLocation.Value)
            End If
        Else
            dt = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, txtFromLocation.Value)
        End If
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If isForCurrentRow Then
                If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
                    BlankTaxDetails(gv1.CurrentRow.Index, isChangeRate)
                Else
                    BlankTaxDetails(gv1.CurrentRow.Index)
                End If
                If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                        If (CalculateTaxRatefromItemwsieTaxOnSale = 0) Then
                            If isChangeRate Then
                                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                            End If
                        Else
                            If isChangeRate Then
                                Dim objTM As clsItemWiseTaxAuthority
                                objTM = clsItemWiseTaxAuthority.GetAutoItemwiseTaxRate(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(txtTaxGroup.Value), clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value), txtDate.Value, "T")
                                If objTM IsNot Nothing Then
                                    gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = objTM.TAX_Rate
                                    gv1.CurrentRow.Cells(colItemwiseTaxCode).Value = objTM.HCODE
                                End If
                            End If
                        End If
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            Else
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
                        BlankTaxDetails(intRowNo, isChangeRate)
                    Else
                        BlankTaxDetails(intRowNo)
                    End If
                    If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode).Value) > 0 Then
                        Dim ii As Integer = 1
                        For Each dr As DataRow In dt.Rows
                            Dim strII As String = clsCommon.myCstr(ii)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                            If (CalculateTaxRatefromItemwsieTaxOnSale = 0) Then
                                If isChangeRate Then
                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                                End If
                            Else
                                If isChangeRate Then
                                    Dim objTM As clsItemWiseTaxAuthority
                                    objTM = clsItemWiseTaxAuthority.GetAutoItemwiseTaxRate(clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value), clsCommon.myCstr(txtTaxGroup.Value), clsCommon.myCstr(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value), txtDate.Value, "T")
                                    If objTM IsNot Nothing Then
                                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = objTM.TAX_Rate
                                        gv1.Rows(intRowNo).Cells(colItemwiseTaxCode).Value = objTM.HCODE
                                    End If
                                End If
                            End If
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                            ii = ii + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub
    Sub SetitemWiseTaxOnlySetting()
        ' Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetails(txtTaxGroup.Value)
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsByLocationForTransfer(txtTaxGroup.Value, "T", txtToLoc.Value, txtFromLocation.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode)) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            Next
        End If
    End Sub
    'Added by preeti Gupta Against ticket no[]==============
    Private Sub SetTax()
        If clsCommon.myLen(clsCommon.myCstr(txtToLoc.Value)) <= 0 Then
            txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroupforTransfer(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value)
            lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfPurchaseType(txtTaxGroup.Value, Nothing)
            SetTaxDetails()
        Else
            txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroupforTransfer(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value)
            lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfPurchaseType(txtTaxGroup.Value, Nothing)
            SetTaxDetails()
        End If
    End Sub
    '========================================================
    Private Sub txtBillToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtFromLocation._MYValidating
        Dim qry As String = ""
        txtToLoc.Value = ""
        'Dim qry As String = "Select  LM.Location_Code as Code,LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable'  from TSPL_LOCATION_MASTER as LM "
        'Dim whrclas As String = " LM.Location_Type ='Physical'"
        'txtFromLocation.Value = clsCommon.myCstr(clsCommon.ShowSelectForm("From Location", qry, "Code", whrclas, txtFromLocation.Value, "Code", isButtonClicked))
        'lblFromLoc.Text = clsCommon.myCstr(connectSql.RunScalar("SELECT Location_Desc  FROM TSPL_LOCATION_MASTER WHERE Location_Code = '" + Convert.ToString(txtFromLocation.Value) + "'"))
        'If clsTransferDCC.IsTaxable(txtFromLocation.Value, txtToLoc.Value, txtDate.Value, Nothing) Then
        '    chkTaxable.Checked = True
        'Else
        '    chkTaxable.Checked = False
        'End If
        If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso clsCommon.myLen(clsCommon.myCstr(cboType.SelectedValue)) <= 0 Then
            txtFromLocation.Value = ""
            clsCommon.MyMessageBoxShow(Me, "Please select Type", Me.Text)
            cboType.Focus()
            Exit Sub
        End If
        '' added functionality for UDL Internal Transfer
        If clsCommon.myCBool(EnableInternalTransfer) = True AndAlso (chkInternalTransfer.Checked) = True Then
            qry = "Select  LM.Location_Code as Code,LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable'  from TSPL_LOCATION_MASTER as LM "
            Dim whrclas As String = " LM.Location_Type ='Physical' and (LM.GIT_Type='N' or isnull(LM.GIT_Type,'')='')"
            txtFromLocation.Value = clsCommon.myCstr(clsCommon.ShowSelectForm("From Location", qry, "Code", whrclas, txtFromLocation.Value, "Code", isButtonClicked))
            lblFromLoc.Text = clsCommon.myCstr(connectSql.RunScalar("SELECT Location_Desc  FROM TSPL_LOCATION_MASTER WHERE Location_Code = '" + Convert.ToString(txtFromLocation.Value) + "'"))
        Else
            qry = "Select  LM.Location_Code as Code,LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable'  from TSPL_LOCATION_MASTER as LM "
            Dim whrclas As String = "Rejected_Type='N' and LM.Location_Type ='Physical'"
            If clsCommon.myLen(arrLoc) > 0 Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                    whrclas += " and LM.Location_Code in (" + arrLoc + ")"
                End If
            End If
            If clsCommon.CompairString(clsCommon.myCstr(cboType.SelectedValue), "Excise") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(cmbGPItemType.SelectedValue), "E") = CompairStringResult.Equal Then
                If clsERPFuncationality.GetGSTStatus(txtDate.Value) = False Then
                    whrclas += " and LM.Excisable ='T'"
                End If
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboType.SelectedValue), "Depot") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboType.SelectedValue), "Route") = CompairStringResult.Equal Then
                If clsERPFuncationality.GetGSTStatus(txtDate.Value) = False Then
                    whrclas += " and LM.Excisable <>'T'"
                End If
            End If
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                whrclas += " and (LM.GIT_Type='N' or isnull(LM.GIT_Type,'')='')"
            End If
            txtFromLocation.Value = clsCommon.myCstr(clsCommon.ShowSelectForm("From Location", qry, "Code", whrclas, txtFromLocation.Value, "Code", isButtonClicked))
            lblFromLoc.Text = clsCommon.myCstr(connectSql.RunScalar("SELECT Location_Desc  FROM TSPL_LOCATION_MASTER WHERE Location_Code = '" + Convert.ToString(txtFromLocation.Value) + "'"))
            'comment by balwinder on 31/01/2015.asked by amit sir becuase user will pick the to location.
            'If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
            '    txtToLoc.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select git_location from tspl_location_master where location_code='" + txtFromLocation.Value + "'"))
            '    lblToLoc.Text = clsLocation.GetName(txtToLoc.Value, Nothing)
            'Else
            '    txtToLoc.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select location_code from tspl_location_master where git_location='" + txtFromLocation.Value + "'"))
            '    lblToLoc.Text = clsLocation.GetName(txtToLoc.Value, Nothing)
            'End If
        End If
        If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso clsCommon.myLen(txtFromLocation.Value) > 0 Then
            cboType.Enabled = False
        End If
        '' Anubhooti 08-June-2015 (If FromLoc state and ToLoc state is diff then check f form checkbox should be enable)
        Dim ToLocState As String = String.Empty
        Dim FromLocState As String = String.Empty
        If clsCommon.myLen(txtToLoc.Value) > 0 Then
            ToLocState = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(STATE,'') AS StateName FROM TSPL_LOCATION_MASTER WHERE Location_Code='" + txtToLoc.Value + "'"))
        End If
        If clsCommon.myLen(txtFromLocation.Value) > 0 Then
            FromLocState = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(STATE,'') AS StateName FROM TSPL_LOCATION_MASTER WHERE Location_Code='" + txtFromLocation.Value + "'"))
        End If
        If clsCommon.myLen(ToLocState) > 0 AndAlso clsCommon.myLen(FromLocState) > 0 Then
            If clsCommon.CompairString(ToLocState, FromLocState) <> CompairStringResult.Equal Then
                chkAgainst_Form.Enabled = True
                chkAgainst_Form.Checked = True
            Else
                chkAgainst_Form.Enabled = False
                chkAgainst_Form.Checked = False
            End If
        Else
            chkAgainst_Form.Enabled = False
            chkAgainst_Form.Checked = False
        End If
        'Dim qry As String = "select Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER "
        'Dim WhrCls As String = " Location_Type='Physical'  "
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
        'End If
        'txtFromLocation.Value = clsCommon.ShowSelectForm("BILLTOLOCPO", qry, "Code", WhrCls, txtFromLocation.Value, "Code", isButtonClicked)
        'lblFromLoc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtFromLocation.Value + "'"))
        SetTax()
    End Sub
    Private Sub txtShipToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtToLoc._MYValidating
        Dim qry As String = ""
        If (clsCommon.myCBool(EnableInternalTransfer) = True AndAlso (chkInternalTransfer.Checked) = True) OrElse (clsCommon.myCBool(JobWorkTransfer) = True AndAlso (chkJobWorkTransfer.Checked) = True) Then
            If clsCommon.myLen(txtFromLocation.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select From Location", Me.Text)
                Exit Sub
            End If

            Dim locSub As String = clsDBFuncationality.getSingleValue("select is_sub_location from TSPL_LOCATION_MASTER where Location_Code='" & txtFromLocation.Value & "'")
            If clsCommon.CompairString(locSub, "Y") = CompairStringResult.Equal Then
                qry = "Select LM.Location_Code as Code,lm.Main_Location_Code as [Main Location Code],LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable',GIT_Location as GITLocation from TSPL_LOCATION_MASTER as LM "
                'Asked by Shruti mam, show all location which main loaction is equal to from location(Sub Location -To- Sub Loaction transfer)
                Dim StrMainLoc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Main_Location_Code,'') AS MainLoc FROM TSPL_LOCATION_MASTER WHERE Location_Code='" + txtFromLocation.Value + "'"))
                Dim whrclas As String = "2=2 and (Location_Code='" & StrMainLoc & "' or Main_Location_Code='" & StrMainLoc & "')"
                txtToLoc.Value = clsCommon.ShowSelectForm("To Location", qry, "Code", whrclas, txtToLoc.Value, "Code", isButtonClicked)
            Else
                qry = "Select  LM.Location_Code as Code,LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable',GIT_Location as GITLocation from TSPL_LOCATION_MASTER as LM "
                Dim whrclas As String = "2=2"
                whrclas += " and LM.Location_Type ='Physical'  and (GIT_Type='N' or   GIT_Type ='') and is_sub_location='Y' and Main_Location_Code='" & txtFromLocation.Value & "' "
                txtToLoc.Value = clsCommon.ShowSelectForm("To Location", qry, "Code", whrclas, txtToLoc.Value, "Code", isButtonClicked)
            End If
        Else
            qry = "Select  LM.Location_Code as Code,LM.Location_Desc as Description,LM.Loc_Short_Name as [Short Name],Location_type as 'Location Type',(case LM.Excisable when 'T'then 'Excisable'else 'Non-Excisable'end) as 'Excisable',GIT_Location as GITLocation from TSPL_LOCATION_MASTER as LM "
            Dim whrclas As String = "2=2"
            If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                whrclas += " and LM.Location_Type ='Physical'  and (GIT_Type='N' or   GIT_Type ='') and LM.location_code in (" + arrLoc + ") "
                txtToLoc.Value = clsCommon.ShowSelectForm("To Location", qry, "Code", whrclas, txtToLoc.Value, "Code", isButtonClicked)
            ElseIf clsLocation.isLocatinExcisable(txtFromLocation.Value) AndAlso clsERPFuncationality.GetGSTStatus(txtDate.Value) = False Then
                whrclas += " and LM.Location_Type ='Physical' and len(GIT_location)>0 "
                txtToLoc.Value = clsCommon.ShowSelectForm("To Location", qry, "Code", whrclas, txtToLoc.Value, "GITLocation", isButtonClicked)
            Else
                whrclas += " and LM.Location_Type ='Physical' and len(GIT_location)>0 "
                txtToLoc.Value = clsCommon.ShowSelectForm("To Location", qry, "Code", whrclas, txtToLoc.Value, "GITLocation", isButtonClicked)
            End If
        End If
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrclas += " and LM.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
        'End If
        FillLocationInfo(False)
        '' Anubhooti 08-June-2015 (If FromLoc state and ToLoc state is diff then check f form checkbox should be enable)
        Dim ToLocState As String = String.Empty
        Dim FromLocState As String = String.Empty
        If clsCommon.myLen(txtToLoc.Value) > 0 Then
            ToLocState = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(STATE,'') AS StateName FROM TSPL_LOCATION_MASTER WHERE Location_Code='" + txtToLoc.Value + "'"))
        End If
        If clsCommon.myLen(txtFromLocation.Value) > 0 Then
            FromLocState = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(STATE,'') AS StateName FROM TSPL_LOCATION_MASTER WHERE Location_Code='" + txtFromLocation.Value + "'"))
        End If
        If clsCommon.myLen(ToLocState) > 0 AndAlso clsCommon.myLen(FromLocState) > 0 Then
            If clsCommon.CompairString(ToLocState, FromLocState) <> CompairStringResult.Equal Then
                chkAgainst_Form.Enabled = True
                chkAgainst_Form.Checked = True
            Else
                chkAgainst_Form.Enabled = False
                chkAgainst_Form.Checked = False
            End If
        Else
            chkAgainst_Form.Enabled = False
            chkAgainst_Form.Checked = False
        End If
        SetTax()
    End Sub
    Private Sub FillLocationInfo(ByVal IsLoadData As Boolean)
        Try
            dt = clsDBFuncationality.GetDataTable("Select Location_Code, Location_Desc, GIT_location from TSPL_LOCATION_MASTER WHERE Location_Code ='" & txtToLoc.Value & "'")
            If dt.Rows.Count > 0 Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                    lblToLoc.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
                    lblToLoc.Tag = clsCommon.myCstr(dt.Rows(0)("Location_Code"))
                ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                    lblToLoc.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
                    lblToLoc.Tag = clsCommon.myCstr(dt.Rows(0)("Location_Code"))
                ElseIf clsLocation.isLocatinExcisable(txtFromLocation.Value) Then
                    lblToLoc.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
                    lblToLoc.Tag = clsCommon.myCstr(dt.Rows(0)("GIT_location"))
                Else
                    lblToLoc.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
                    lblToLoc.Tag = clsCommon.myCstr(dt.Rows(0)("GIT_location"))
                End If
            Else
                lblToLoc.Text = ""
                lblToLoc.Tag = ""
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub SelectRequistionItems()
        isInsideLoadData = True
        Dim frm As New frmPendingTransfer()
        frm.arrLoc = arrLoc
        frm.strCurrCode = txtDocNo.Value
        Dim TransferType As String = clsCommon.myCstr(cboTransferType.SelectedValue)
        'If clsCommon.CompairString(TransferType, "R") = CompairStringResult.Equal Then
        '    frm.strRetuenType = "R"
        'Else
        '    frm.strRetuenType = ""
        'End If
        frm.Text = "Pending Transfer"
        frm.InternalTransfer = IIf(chkInternalTransfer.Checked, 1, 0)
        frm.JobWorkType = IIf(chkJobWork.Checked, 1, 0)
        frm.ShowDialog()
        txtFromLocation.Enabled = True
        txtToLoc.Enabled = True
        txtTaxGroup.Enabled = False
        Dim strDate As String = txtDate.Value
        If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
            '=============Added by preeti Gupta against Ticket No[UDL/24/07/18-000207]
            cboTransferType.Enabled = False
            '''''''''================================================
            strDate = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Date from TSPL_TRANSFER_ORDER_HEAD  where Document_No='" & frm.ArrReturn(0).TransferOutNo & "'"))
        End If
        If clsERPFuncationality.GetGSTStatus(strDate) = False Then
            If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
                Dim objReq As clsTransferDCC = clsTransferDCC.GetData(frm.ArrReturn(0).TransferOutNo, NavigatorType.Current)
                If objReq IsNot Nothing AndAlso clsCommon.myLen(objReq.Document_No) > 0 Then
                    txtFromLocation.Enabled = False
                    txtToLoc.Enabled = False
                    If (clsCommon.myLen(txtFromLocation.Value) <= 0) Then
                        txtFromLocation.Value = objReq.To_Location
                        lblFromLoc.Text = objReq.To_LocationName
                    End If
                    'txtToLoc.Value = objReq.From_Location
                    'lblToLoc.Text = objReq.From_LocationName
                    If clsCommon.myCBool(objReq.InternalTransfer) = True Then
                        txtToLoc.Value = objReq.InternalTransfer
                    Else
                        txtToLoc.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select location_code from tspl_location_master where git_location='" + txtFromLocation.Value + "'"))
                    End If
                    lblToLoc.Text = clsLocation.GetName(txtToLoc.Value, Nothing)
                    txtVehicle_Capacity.Text = objReq.Vehicle_Capacity
                    txtvehicle_Charge.Text = objReq.Vehicle_Charge
                    txtvehicle_mannual_no.Text = objReq.Vehicle_Mannual_No
                    txttotal_Wt.Text = objReq.Total_Item_Wt
                    txtGross_Wt.Text = objReq.Gross_Item_Wt
                    'richa agarwal 2016
                    chkOwnVehicle.Checked = False
                    If clsCommon.myLen(objReq.Transport_Id) <= 0 Then
                        chkOwnVehicle.Checked = True
                        TxtTransportorMName.Text = clsCommon.myCstr(objReq.Transporter_Name_Manual)
                        txtTransporter_Code.Visible = False
                    Else
                        TxtTransportorMName.Text = ""
                    End If
                    chkInternalTransfer.Enabled = False
                    ''---------------------------------------
                    '====Enabled False of Repair Check Box ==============
                    chkForRepair.Checked = IIf(objReq.For_Repair = 1, True, False)
                    chkForRepair.Enabled = False
                    '==============
                    '' code done by Panch raj against email sent by ranjana mam(Anand rohela)
                    txtTransporter_Code.Value = objReq.Transport_Id
                    txtTransporter_desc.Text = clsTransferDCC.GetTransporterName(txtTransporter_Code.Value)
                    txtGR_No.Text = objReq.GR_No
                    dtpGR.Value = objReq.GR_Date
                    txtWayBill_No.Text = objReq.Waybill_No
                    ttxway_bill_date.Value = objReq.Waybill_Date
                    txtOutDate.Value = objReq.Document_Date
                    If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                        txtRemarks.Text = objReq.Remarks
                    End If
                    If (clsCommon.myLen(txtRefNo.Text) <= 0) Then
                        txtRefNo.Text = objReq.Ref_No
                    End If
                    If (clsCommon.myLen(txtVehicleCode.Value) <= 0) Then
                        txtVehicleCode.Value = objReq.Vehicle_Code
                        lblVehicleNo.Text = objReq.Vehicle_No
                    End If
                    If (clsCommon.myLen(cboTransferType.SelectedValue) <= 0) Then
                        cboTransferType.SelectedValue = objReq.Transfer_Type
                    End If
                    If (clsCommon.myLen(cboModeOfTransport.Text) <= 0) Then
                        cboModeOfTransport.Text = objReq.Mode_Of_Transport
                    End If
                    If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                        txtRemarks.Text = objReq.Remarks
                    End If
                    If (clsCommon.myLen(txtKm.Text) <= 0) Then
                        txtKm.Text = objReq.Km_Reading
                    End If
                    '=================Added By preeti Gupta==================
                    txtcrateout.Text = clsCommon.myCdbl(objReq.Crate_Out)
                    txtjaaliout.Text = clsCommon.myCdbl(objReq.jaali_Out)
                    txtBoxOut.Text = clsCommon.myCdbl(objReq.Box_Out)
                    '========================================================
                End If
                If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
                    gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
                End If
                gv1.DataSource = Nothing
                gv1.Rows.Clear()
                For Each obj As clsTransferDCCDetail In frm.ArrReturn
                    StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                    If StrCrateTransferFromBooking = "1" Then
                        gv1.Rows.AddNew()
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = frmGRN.RowTypeItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.TransferOutNo
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colFOCItem).Value = obj.FOCItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                        '=================Added by preeti gupta Against Ticket No[BHA/24/07/18-000189]
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                        '====================================================================
                        '====Added By Preeti Gupta=========================
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                        '====================================================
                        gv1.Columns(colInQty).IsVisible = True
                        gv1.Columns(colLeakQty).IsVisible = True
                        gv1.Columns(colBreakQty).IsVisible = True
                        gv1.Columns(colShortQty).IsVisible = True
                        CalMT_WT(gv1.Rows.Count - 1) 'done by stuti against kdil bug on 03/01/2017
                        SetitemWiseTaxSetting(True, True)
                        UpdateCurrentRow(gv1.Rows.Count - 1)
                    Else
                        If Not IsItemExistInGrid(obj) Then
                            gv1.Rows.AddNew()
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = frmGRN.RowTypeItem
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.TransferOutNo
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                            '=================Added by preeti gupta Against Ticket No[BHA/24/07/18-000189]
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                            '====================================================================
                            '====Added By Preeti Gupta=========================
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                            '====================================================
                            gv1.Columns(colInQty).IsVisible = True
                            gv1.Columns(colLeakQty).IsVisible = True
                            gv1.Columns(colBreakQty).IsVisible = True
                            gv1.Columns(colShortQty).IsVisible = True
                            CalMT_WT(gv1.Rows.Count - 1) ' done by stuti against kdil bug on 03/01/2017
                            SetitemWiseTaxSetting(True, True)
                            UpdateCurrentRow(gv1.Rows.Count - 1)
                        End If
                    End If
                Next
            End If
        Else
            If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
                Dim objReq As clsTransferDCC = clsTransferDCC.GetData(frm.ArrReturn(0).TransferOutNo, NavigatorType.Current)
                If objReq IsNot Nothing AndAlso clsCommon.myLen(objReq.Document_No) > 0 Then
                    txtFromLocation.Enabled = False
                    txtToLoc.Enabled = False
                    If clsCommon.myCBool(objReq.InternalTransfer) = True Then
                        If (clsCommon.myLen(txtFromLocation.Value) <= 0) Then
                            txtFromLocation.Value = objReq.From_Location
                            lblFromLoc.Text = objReq.From_LocationName
                        End If
                    Else
                        If (clsCommon.myLen(txtFromLocation.Value) <= 0) Then
                            txtFromLocation.Value = objReq.To_Location
                            lblFromLoc.Text = objReq.To_LocationName
                        End If
                    End If

                    'txtToLoc.Value = objReq.From_Location
                    'lblToLoc.Text = objReq.From_LocationName
                    If clsCommon.myCBool(objReq.InternalTransfer) = True Then
                        txtToLoc.Value = objReq.To_Location
                        chkInternalTransfer.Enabled = False
                    Else
                        txtToLoc.Value = objReq.To_Location_Main
                        'txtToLoc.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select location_code from tspl_location_master where git_location='" + txtToLoc.Value + "'"))
                        chkInternalTransfer.Enabled = False
                    End If
                    lblToLoc.Text = clsLocation.GetName(txtToLoc.Value, Nothing)
                    txtLoadingAdviceNo.Text = objReq.Loading_Advice_No
                    txtVehicle_Capacity.Text = objReq.Vehicle_Capacity
                    txtvehicle_Charge.Text = objReq.Vehicle_Charge
                    txtvehicle_mannual_no.Text = objReq.Vehicle_Mannual_No
                    txttotal_Wt.Text = objReq.Total_Item_Wt
                    txtGross_Wt.Text = objReq.Gross_Item_Wt
                    'richa agarwal 2016
                    chkOwnVehicle.Checked = False
                    txtEWayBillNo.Text = objReq.EWayBillNo
                    txtElecttefNo.Text = objReq.Electronic_Ref_No
                    If objReq.EWayBillDate IsNot Nothing Then
                        txtEWayBillDate.Value = objReq.EWayBillDate
                        txtEWayBillDate.Checked = True
                    End If
                    chkTaxable.Checked = IIf(objReq.Is_Taxable = 1, True, False)
                    If clsCommon.myLen(objReq.Transport_Id) <= 0 Then
                        chkOwnVehicle.Checked = True
                        TxtTransportorMName.Text = clsCommon.myCstr(objReq.Transporter_Name_Manual)
                        txtTransporter_Code.Visible = False
                    Else
                        TxtTransportorMName.Text = ""
                    End If
                    ''---------------------------------------
                    '====Enabled False of Repair Check Box ==============
                    chkForRepair.Checked = IIf(objReq.For_Repair = 1, True, False)
                    chkForRepair.Enabled = False
                    '==============
                    '' code done by Panch raj against email sent by ranjana mam(Anand rohela)
                    txtTransporter_Code.Value = objReq.Transport_Id
                    txtTransporter_desc.Text = clsTransferDCC.GetTransporterName(txtTransporter_Code.Value)
                    txtGR_No.Text = objReq.GR_No
                    dtpGR.Value = objReq.GR_Date
                    txtWayBill_No.Text = objReq.Waybill_No
                    ttxway_bill_date.Value = objReq.Waybill_Date
                    txtOutDate.Value = objReq.Document_Date
                    If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                        txtRemarks.Text = objReq.Remarks
                    End If
                    If (clsCommon.myLen(txtRefNo.Text) <= 0) Then
                        txtRefNo.Text = objReq.Ref_No
                    End If
                    If (clsCommon.myLen(txtVehicleCode.Value) <= 0) Then
                        txtVehicleCode.Value = objReq.Vehicle_Code
                        lblVehicleNo.Text = objReq.Vehicle_No
                    End If
                    If (clsCommon.myLen(cboTransferType.SelectedValue) <= 0) Then
                        cboTransferType.SelectedValue = objReq.Transfer_Type
                    End If
                    If (clsCommon.myLen(cboModeOfTransport.Text) <= 0) Then
                        cboModeOfTransport.Text = objReq.Mode_Of_Transport
                    End If
                    If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                        txtRemarks.Text = objReq.Remarks
                    End If
                    If (clsCommon.myLen(txtKm.Text) <= 0) Then
                        txtKm.Text = objReq.Km_Reading
                    End If
                    '=================Added By preeti Gupta==================
                    txtcrateout.Text = clsCommon.myCdbl(objReq.Crate_Out)
                    txtjaaliout.Text = clsCommon.myCdbl(objReq.jaali_Out)
                    txtBoxOut.Text = clsCommon.myCdbl(objReq.Box_Out)
                    '========================================================
                    txtTransferOutNo.Value = objReq.Document_No
                    txtTaxGroup.Value = objReq.Tax_Group
                    SetTaxDetails()
                End If
                If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
                    gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
                End If
                gv1.DataSource = Nothing
                gv1.Rows.Clear()
                For Each obj As clsTransferDCCDetail In objReq.Arr
                    StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                    If StrCrateTransferFromBooking = "1" Then
                        gv1.Rows.AddNew()
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = frmGRN.RowTypeItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.Document_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = True
                        Else
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = False
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colFOCItem).Value = obj.FOCItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                        '=================Added by preeti gupta Against Ticket No[BHA/24/07/18-000189]
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                        '====================================================================
                        '====Added By Preeti Gupta=========================
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
                        '====================================================
                        gv1.Columns(colInQty).IsVisible = True
                        gv1.Columns(colLeakQty).IsVisible = True
                        gv1.Columns(colBreakQty).IsVisible = True
                        gv1.Columns(colShortQty).IsVisible = True
                        CalMT_WT(gv1.Rows.Count - 1) 'done by stuti against kdil bug on 03/01/2017
                    Else
                        If Not IsItemExistInGrid(obj) Then
                            gv1.Rows.AddNew()
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = frmGRN.RowTypeItem
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.Document_No
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                            If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = True
                            Else
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = False
                            End If
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                            '=================Added by preeti gupta Against Ticket No[BHA/24/07/18-000189]
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                            '====================================================================
                            '====Added By Preeti Gupta=========================
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
                            '====================================================
                            gv1.Columns(colInQty).IsVisible = True
                            gv1.Columns(colLeakQty).IsVisible = True
                            gv1.Columns(colBreakQty).IsVisible = True
                            gv1.Columns(colShortQty).IsVisible = True
                            CalMT_WT(gv1.Rows.Count - 1) ' done by stuti against kdil bug on 03/01/2017
                            'SetitemWiseTaxSetting(True, True)
                            'UpdateCurrentRow(gv1.Rows.Count - 1)
                        End If
                    End If
                Next
                SetitemWiseTaxSetting(False, False)
                'isValid_CashScheme()
                For ii As Integer = 0 To gv1.RowCount - 1
                    UpdateCurrentRow(ii)
                Next
            End If
        End If
        isInsideLoadData = False
        UpdateAllTotals()
        FillLocationInfo(False)
        RefreshReqNo()
    End Sub
    '=====================Added by preeti Gupta Against Ticket No[UDL/24/07/18-000207]===================================
    Sub SelectReturnItems()
        isInsideLoadData = True
        txtFromLocation.Enabled = True
        txtToLoc.Enabled = True
        txtTaxGroup.Enabled = False
        ''UDL/24/07/18-000207 richa 
        'Dim whrclas As String = "Transfer_Type ='I' and status=1 and transferoutno not in (SELECT transferoutno FROM TSPL_TRANSFER_ORDER_HEAD WHERE  Transfer_Type ='T' and Status =1)"
        'Dim whrclas As String = "Transfer_Type ='I' and status=1 and transferoutno not in (sELECT DISTINCT Z.TransferOutNo  FROM (SELECT TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo ,Item_Code ,SUM(Out_Qty) AS Out_Qty ,SUM(In_Qty) AS In_Qty  FROM TSPL_TRANSFER_ORDER_DETAIL " & _
        '" LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No WHERE  Transfer_Type ='T' GROUP BY TSPL_TRANSFER_ORDER_DETAIL.transferoutno,Item_Code HAVING (MAX(Out_Qty)-sUM(In_Qty))=0)Z)"
        'If (clsCommon.myLen(txtFromLocation.Value) > 0) Then
        '    whrclas += " and To_Location='" + txtFromLocation.Value + "'"
        'End If
        'Dim qry As String = "select transferoutno as Document_No,convert(varchar,Document_Date,103) as Document_Date from TSPL_TRANSFER_ORDER_HEAD "
        Dim QRY As String = " Select distinct z.Document_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as Document_Date from  (select code as Document_No,SUM(Qty* case when RI=1 then 1 else 0 end) as OutQty, SUM(Qty* case when RI=-1 then 1 else 0 end) as ReturnQty, SUM(Unapproved) as UnapprovedQty,SUM((Qty *RI)- Unapproved) as PendingQty,Max(Unit_Code) as UnitCode from " &
        " ( " &
        " Select TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo as Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,0 as Unapproved,1 as RI,TSPL_TRANSFER_ORDER_DETAIL.Unit_Code FROM TSPL_TRANSFER_ORDER_DETAIL  LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No  where TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ='I' and  TSPL_TRANSFER_ORDER_HEAD.Status =1 " & Environment.NewLine &
        " Union All " &
        " Select TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo  as Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_ORDER_DETAIL.In_Qty  as Qty,0 as Unapproved,-1 as RI,TSPL_TRANSFER_ORDER_DETAIL.Unit_Code FROM TSPL_TRANSFER_ORDER_DETAIL  LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No  where " &
        " TSPL_TRANSFER_ORDER_HEAD.Transfer_Type ='T' --and  TSPL_TRANSFER_ORDER_HEAD.Status =1 " & Environment.NewLine &
        " ) Final  " &
        " group by Code,ICode having (SUM(Qty *RI)-SUM(Unapproved)) >0  )z " &
        " left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =z.Document_No "
        txtTransferOutNo.Value = clsCommon.myCstr(clsCommon.ShowSelectForm("Document_No", QRY, "Document_No", "", txtTransferOutNo.Value, "Document_No", True))
        Dim objReq As clsTransferDCC = clsTransferDCC.GetData(txtTransferOutNo.Value, NavigatorType.Current)
        If objReq IsNot Nothing AndAlso clsCommon.myLen(objReq.Document_No) > 0 Then
            cboTransferType.Enabled = False
            txtFromLocation.Enabled = False
            txtToLoc.Enabled = False
            Dim qryLocation As String = "select From_Location ,(Select location_code from tspl_location_master where git_location='" + objReq.To_Location + "') as To_location from TSPL_TRANSFER_ORDER_HEAD where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + txtTransferOutNo.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qryLocation)
            If (clsCommon.myLen(txtFromLocation.Value) <= 0) Then
                txtFromLocation.Value = dt.Rows(0).Item("To_location")
                lblFromLoc.Text = clsLocation.GetName(txtFromLocation.Value, Nothing)
            End If
            If clsCommon.myCBool(objReq.InternalTransfer) = True Then
                txtToLoc.Value = dt.Rows(0).Item("From_Location") 'objReq.From_Location
                chkInternalTransfer.Enabled = False
            Else
                txtToLoc.Value = dt.Rows(0).Item("From_Location")
                chkInternalTransfer.Enabled = False
            End If
            lblToLoc.Text = clsLocation.GetName(txtToLoc.Value, Nothing)
            txtVehicle_Capacity.Text = objReq.Vehicle_Capacity
            txtvehicle_Charge.Text = objReq.Vehicle_Charge
            txtvehicle_mannual_no.Text = objReq.Vehicle_Mannual_No
            txttotal_Wt.Text = objReq.Total_Item_Wt
            txtGross_Wt.Text = objReq.Gross_Item_Wt
            chkOwnVehicle.Checked = False
            txtEWayBillNo.Text = objReq.EWayBillNo
            txtElecttefNo.Text = objReq.Electronic_Ref_No
            If objReq.EWayBillDate IsNot Nothing Then
                txtEWayBillDate.Value = objReq.EWayBillDate
                txtEWayBillDate.Checked = True
            End If
            chkTaxable.Checked = IIf(objReq.Is_Taxable = 1, True, False)
            If clsCommon.myLen(objReq.Transport_Id) <= 0 Then
                chkOwnVehicle.Checked = True
                TxtTransportorMName.Text = clsCommon.myCstr(objReq.Transporter_Name_Manual)
                txtTransporter_Code.Visible = False
            Else
                TxtTransportorMName.Text = ""
            End If
            chkForRepair.Checked = IIf(objReq.For_Repair = 1, True, False)
            chkForRepair.Enabled = False
            txtTransporter_Code.Value = objReq.Transport_Id
            txtTransporter_desc.Text = clsTransferDCC.GetTransporterName(txtTransporter_Code.Value)
            txtGR_No.Text = objReq.GR_No
            dtpGR.Value = objReq.GR_Date
            txtWayBill_No.Text = objReq.Waybill_No
            ttxway_bill_date.Value = objReq.Waybill_Date
            txtOutDate.Value = objReq.Document_Date
            If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                txtRemarks.Text = objReq.Remarks
            End If
            If (clsCommon.myLen(txtRefNo.Text) <= 0) Then
                txtRefNo.Text = objReq.Ref_No
            End If
            If (clsCommon.myLen(txtVehicleCode.Value) <= 0) Then
                txtVehicleCode.Value = objReq.Vehicle_Code
                lblVehicleNo.Text = objReq.Vehicle_No
            End If
            If (clsCommon.myLen(cboTransferType.SelectedValue) <= 0) Then
                cboTransferType.SelectedValue = objReq.Transfer_Type
            End If
            If (clsCommon.myLen(cboModeOfTransport.Text) <= 0) Then
                cboModeOfTransport.Text = objReq.Mode_Of_Transport
            End If
            If (clsCommon.myLen(txtRemarks.Text) <= 0) Then
                txtRemarks.Text = objReq.Remarks
            End If
            If (clsCommon.myLen(txtKm.Text) <= 0) Then
                txtKm.Text = objReq.Km_Reading
            End If
            txtcrateout.Text = clsCommon.myCdbl(objReq.Crate_Out)
            txtjaaliout.Text = clsCommon.myCdbl(objReq.jaali_Out)
            txtBoxOut.Text = clsCommon.myCdbl(objReq.Box_Out)
            txtTransferOutNo.Value = objReq.Document_No
            txtTaxGroup.Value = objReq.Tax_Group
            SetTaxDetails()
            If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
                gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
            End If
            gv1.DataSource = Nothing
            gv1.Rows.Clear()
            For Each obj As clsTransferDCCDetail In objReq.Arr
                StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                If StrCrateTransferFromBooking = "1" Then
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.Document_No
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                    If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = True
                    Else
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = False
                    End If
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
                    gv1.Columns(colInQty).IsVisible = True
                    gv1.Columns(colLeakQty).IsVisible = True
                    gv1.Columns(colBreakQty).IsVisible = True
                    gv1.Columns(colShortQty).IsVisible = True
                    CalMT_WT(gv1.Rows.Count - 1)
                Else
                    If Not IsItemExistInGrid(obj) Then
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = obj.Document_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = True
                        Else
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).ReadOnly = False
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
                        If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                            Dim strqry As String = "SELECT (MAX(Out_Qty)-sUM(In_Qty)) AS Ret_Qty  FROM TSPL_TRANSFER_ORDER_DETAIL " &
                            " LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No WHERE  Transfer_Type ='T' " &
                            " and TSPL_TRANSFER_ORDER_DETAIL.TransferOutNo='" & obj.Document_No & "' and TSPL_TRANSFER_ORDER_DETAIL.Item_Code ='" & obj.Item_Code & "'  and TSPL_TRANSFER_ORDER_DETAIL.Document_No <>'" + txtDocNo.Value + "' " &
                            " GROUP BY TSPL_TRANSFER_ORDER_DETAIL.transferoutno,Item_Code"
                            Dim dblremainingRetQty As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry, Nothing))
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = dblremainingRetQty
                        Else
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = obj.Out_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colInQty).Value = obj.Out_Qty
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
                        gv1.Columns(colInQty).IsVisible = True
                        gv1.Columns(colLeakQty).IsVisible = True
                        gv1.Columns(colBreakQty).IsVisible = True
                        gv1.Columns(colShortQty).IsVisible = True
                        CalMT_WT(gv1.Rows.Count - 1)
                    End If
                End If
            Next
        End If
        SetitemWiseTaxSetting(False, False)
        For ii As Integer = 0 To gv1.RowCount - 1
            UpdateCurrentRow(ii)
        Next
        isInsideLoadData = False
        UpdateAllTotals()
        FillLocationInfo(False)
        RefreshReqNo()
    End Sub
    '================================================================================================
    Sub SelectGSTTransferOutItems()
        isInsideLoadData = True
        Dim frm As New frmPendingTransfer()
        frm.arrLoc = arrLoc
        frm.strCurrCode = txtDocNo.Value
        frm.Text = "Pending Transfer"
        frm.ShowDialog()
        txtFromLocation.Enabled = True
        txtToLoc.Enabled = True
        isInsideLoadData = False
        UpdateAllTotals()
        FillLocationInfo(False)
        RefreshReqNo()
    End Sub
    Function IsItemExistInGrid(ByVal obj As clsTransferDCCDetail)
        For ii As Integer = 0 To gv1.RowCount - 1
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
            Dim strReqCode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colTransferOutNo).Value)
            If clsCommon.myLen(strReqCode) > 0 AndAlso clsCommon.CompairString(strReqCode, obj.TransferOutNo) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strICode, obj.Item_Code) = CompairStringResult.Equal Then
                common.clsCommon.MyMessageBoxShow("Requition No : " + obj.TransferOutNo + "  Item : " + obj.Item_Desc + Environment.NewLine + "Already exist at row no:" + clsCommon.myCstr(ii + 1))
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub gv1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv1.DoubleClick
        Try
            If CalculateTaxRatefromItemwsieTaxOnSale = 0 Then
                If gv1.CurrentColumn Is gv1.Columns(colTotTaxAmt) AndAlso rbtnTaxCalAutomatic.IsChecked AndAlso clsCommon.CompairString(cboTransferType.Text, "Transfer Out") = CompairStringResult.Equal Then
                    Dim frm As New FrmPOItemTaxDetails()
                    frm.strLineNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value)
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    frm.dblTotTax = clsCommon.myCdbl(gv1.CurrentRow.Cells(colTotTaxAmt).Value)
                    frm.dblAmtAfterDis = clsCommon.myCdbl(gv1.CurrentRow.Cells(colAmtAfterDis).Value)
                    frm.strTaxType = "T"
                    frm.strTransLocation = clsCommon.myCstr(txtFromLocation.Value)
                    frm.strVendorCustomerCode = clsCommon.myCstr(txtToLoc.Value)
                    frm.strTaxGroup = clsCommon.myCstr(txtTaxGroup.Value)
                    If clsCommon.myLen(frm.strItemCode) > 0 Then
                        frm.ArrIn = New List(Of clsTempItemTaxDetails)
                        For ii As Integer = 1 To 10
                            Dim strii As String = clsCommon.myCstr(ii)
                            Dim obj As New clsTempItemTaxDetails()
                            obj.AuthorityCode = clsCommon.myCstr(gv1.CurrentRow.Cells("COLTAX" + strii).Value)
                            If clsCommon.myLen(obj.AuthorityCode) > 0 Then
                                obj.AuthorityName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Code_Desc from TSPL_TAX_MASTER WHERE Tax_Code='" + obj.AuthorityCode + "'"))
                                obj.Rate = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value)
                                obj.BaseAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value)
                                obj.TaxAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value)
                                obj.isSurTax = clsCommon.myCBool(gv1.CurrentRow.Cells("ISSURTAX" + strii).Value)
                                obj.SurTax = clsCommon.myCstr(gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value)
                                obj.IsTaxable = clsCommon.myCBool(gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value)
                                frm.ArrIn.Add(obj)
                            End If
                        Next
                        frm.ShowDialog()
                        If frm.ArrOut IsNot Nothing AndAlso frm.ArrOut.Count > 0 Then
                            BlankTaxDetails(gv1.CurrentRow.Index)
                            For ii As Integer = 0 To frm.ArrOut.Count - 1
                                Dim strii As String = clsCommon.myCstr(ii + 1)
                                gv1.CurrentRow.Cells("COLTAX" + strii).Value = frm.ArrOut(ii).AuthorityCode
                                gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value = frm.ArrOut(ii).Rate
                                gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value = frm.ArrOut(ii).BaseAmt
                                gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value = frm.ArrOut(ii).TaxAmt
                                gv1.CurrentRow.Cells("ISSURTAX" + strii).Value = frm.ArrOut(ii).isSurTax
                                gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value = frm.ArrOut(ii).SurTax
                                gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value = frm.ArrOut(ii).IsTaxable
                            Next
                            gv1.CurrentRow.Cells(colTotTaxAmt).Value = frm.dblTotTax
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            UpdateAllTotals()
                        End If
                    End If
                ElseIf (gv1.CurrentColumn Is gv1.Columns(colComplete) AndAlso (UsLock1.Status = ERPTransactionStatus.Approved)) Then
                    Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    Dim intSNo As Integer = Convert.ToInt32((clsCommon.myCdbl(gv1.CurrentRow.Cells(colLineNo).Value)))
                    Dim strStatus As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colComplete).Value)
                    If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso clsCommon.myLen(strICode) > 0 AndAlso intSNo > 0 AndAlso clsCommon.CompairString(strStatus, "No") = CompairStringResult.Equal Then
                        If common.clsCommon.MyMessageBoxShow("Do you want to complete the item " + clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value), Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                            If clsTransferDCCDetail.CompletePO(txtDocNo.Value, strICode, intSNo) Then
                                common.clsCommon.MyMessageBoxShow(Me, "Successfully Completed", Me.Text)
                                LoadData(txtDocNo.Value, NavigatorType.Current)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rdbtnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            myMessages.blankValue(Me, "Transfer Order No not found to Print", Me.Text)
        Else
            funprint(i)
        End If
    End Sub
    Public Sub funprint(ByVal i As Integer)
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Transfer Order No not found to Print", Me.Text)
            End If
            Dim arr As New ArrayList()
            arr.Add(txtDocNo.Value)
            PrintData(txtDocNo.Value, i)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        UpdateAllTotals()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
        Next
        RefreshReqNo()
    End Sub
    Private Sub gv1_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub
    Private Sub txtDept__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtVehicleCode._MYValidating
        Try
            Dim qry As String = "Select distinct  vehicle_id ,Description from TSPL_VEHICLE_MASTER"
            If isButtonClicked Then
                txtVehicleCode.Value = clsCommon.ShowSelectForm("Vehicle No", qry, "vehicle_id", "", txtVehicleCode.Value, "vehicle_id", isButtonClicked)
            End If
            'txtVehicle.Text = txtVehicleCode.Value
            If clsCommon.myLen(txtVehicleCode.Value) > 0 Then
                lblVehicleNo.Text = connectSql.RunScalar("Select Description  from TSPL_VEHICLE_MASTER where Vehicle_Id = '" + Convert.ToString(txtVehicleCode.Value) + "'")
                txtvehicle_mannual_no.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select number from TSPL_VEHICLE_MASTER where Vehicle_Id = '" + Convert.ToString(txtVehicleCode.Value) + "'"))
            Else
                lblVehicleNo.Text = Nothing
                txtvehicle_mannual_no.Text = Nothing
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv1.CellFormatting
        Try
            ''If e.Column Is gv1.Columns(colICode) Then
            ''    If clsCommon.myLen(gv1.CurrentRow.Cells(colReqistionNo).Value) > 0 OrElse clsCommon.myCBool(gv1.CurrentRow.Cells(colItemUsedINGRN).Value) Then
            ''        gv1.CurrentRow.Cells(colICode).ReadOnly = True
            ''    Else
            ''        gv1.CurrentRow.Cells(colICode).ReadOnly = False
            ''    End If
            ''ElseIf e.Column Is gv1.Columns(colQty) Then
            ''    If clsCommon.myCBool(gv1.CurrentRow.Cells(colItemUsedINGRN).Value) Then
            ''        gv1.CurrentRow.Cells(colQty).ReadOnly = True
            ''    Else
            ''        gv1.CurrentRow.Cells(colQty).ReadOnly = False
            ''    End If
            ''End If
            If e.Column Is gv1.Columns(colICode) Then
                If clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) > 0 OrElse clsCommon.myLen(fndSRNO.Value) > 0 Then
                    If AllowOnlyOneIssueAgainstStoreRequisition = False Then
                        gv1.CurrentRow.Cells(colICode).ReadOnly = True
                    End If
                Else
                    gv1.CurrentRow.Cells(colICode).ReadOnly = False
                End If
            ElseIf e.Column Is gv1.Columns(colUnit) Then
                'If clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) > 0 OrElse clsCommon.myCBool(gv1.CurrentRow.Cells(colItemUsedINGRN).Value) Then
                '    gv1.CurrentRow.Cells(colUnit).ReadOnly = True
                'Else
                '    gv1.CurrentRow.Cells(colUnit).ReadOnly = False
                'End If
            ElseIf e.Column Is gv1.Columns(colAltUOM) Then
                If clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) > 0 Then
                    If AllowOnlyOneIssueAgainstStoreRequisition = False Then
                        gv1.CurrentRow.Cells(colAltUOM).ReadOnly = True
                    End If
                Else
                    gv1.CurrentRow.Cells(colAltUOM).ReadOnly = False
                End If
            ElseIf e.Column Is gv1.Columns(colRate) Then
                If IsRateeditableItemwise Then
                    Dim dblChangeRate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Is_Rate_Change_OnDairyDispatch from tspl_item_master where item_code='" & gv1.CurrentRow.Cells(colICode).Value & "'"))
                    If dblChangeRate = 1 Then
                        gv1.CurrentRow.Cells(colRate).ReadOnly = False
                    Else
                        gv1.CurrentRow.Cells(colRate).ReadOnly = True
                    End If
                End If
            End If
            'ElseIf e.Column Is gv1.Columns(colQty) Then
            '    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
            '        gv1.CurrentRow.Cells(colQty).ReadOnly = False
            '    Else
            '        gv1.CurrentRow.Cells(colQty).ReadOnly = True
            '    End If
            'ElseIf e.Column Is gv1.Columns(colRate) Then
            '    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
            '        gv1.CurrentRow.Cells(colRate).ReadOnly = False
            '    Else
            '        gv1.CurrentRow.Cells(colRate).ReadOnly = True
            '    End If
            'ElseIf e.Column Is gv1.Columns(colAmt) Then
            '    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
            '        gv1.CurrentRow.Cells(colAmt).ReadOnly = True
            '    Else
            '        gv1.CurrentRow.Cells(colAmt).ReadOnly = False
            '    End If
            'End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub
    Private Sub setGridFocus()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
        If intCurrRow = gv1.Rows.Count - 1 Then
            gv1.Rows.AddNew()
            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
            gv1.CurrentRow = gv1.Rows(intCurrRow)
        End If
        If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
            If gv1.CurrentColumn Is gv1.Columns(colICode) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colOutQty)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colOutQty) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colRate)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colRate) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colDisPer)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colDisPer) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colSpecification)
                ''ElseIf gv1.CurrentColumn Is gv1.Columns(colMRP) Then
                ''    gv1.CurrentRow = gv1.Rows(intCurrRow)
                ''    gv1.CurrentColumn = gv1.Columns(colAssessableAmt)
                ''ElseIf gv1.CurrentColumn Is gv1.Columns(colAssessableAmt) Then
                ''    gv1.CurrentRow = gv1.Rows(intCurrRow)
                ''    gv1.CurrentColumn = gv1.Columns(colSpecification)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colSpecification) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colRemarks)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colRemarks) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow + 1)
                gv1.CurrentColumn = gv1.Columns(colICode)
            End If
        End If
    End Sub
    Private Sub txtReqNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTransferOutNo._MYValidating
        If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
            SelectReturnItems()
        Else
            SelectRequistionItems()
        End If
        'If clsERPFuncationality.GetGSTStatus("select Document_Date from TSPL_TRANSFER_ORDER_HEAD  where Document_No='" & txtDocNo.Value & "'") = True Then
        '    SelectGSTTransferOutItems()
        'Else
        '    SelectRequistionItems()
        'End If
    End Sub
    Sub RefreshReqNo()
        txtTransferOutNo.Value = ""
        If gv1.Rows.Count > 0 Then
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strReqNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colTransferOutNo).Value)
                If clsCommon.myLen(strReqNo) > 0 Then
                    txtTransferOutNo.Value = strReqNo
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub PrintAmendment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            myMessages.blankValue(Me, "Transfer Order No not found to Print", Me.Text)
        Else
            FrmPurchaseOrderReport.PrintAbandoment(txtDocNo.Value)
            '' ''clsCommon.ProgressBarShow()
            '' ''For index As Integer = 1 To Integer.MaxValue - 1
            '' ''Next
            '' ''clsCommon.ProgressBarHide()
        End If
    End Sub
    Private Sub txtDocNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocNo.Load
    End Sub
    Sub setBalance()
        UcItemBalance1.ItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        UcItemBalance1.ItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
        UcItemBalance1.ItemMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
        UcItemBalance1.LocationCode = txtFromLocation.Value
        UcItemBalance1.LocationName = lblFromLoc.Text
        UcItemBalance1.UOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
        UcItemBalance1.TransNo = txtDocNo.Value
        UcItemBalance1.TransDate = txtDate.Value
        UcItemBalance1.ShowPOQty = True
        UcItemBalance1.RefreshData()
    End Sub
    Private Sub gv1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv1.Click
        If gv1.CurrentRow IsNot Nothing Then
            setBalance()
        End If
    End Sub
    Private Sub gv1_CurrentRowChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentRowChangedEventArgs) Handles gv1.CurrentRowChanged
        If gv1.CurrentRow IsNot Nothing AndAlso Not e.CurrentRow.Index < 0 Then
            setBalance()
        End If
    End Sub
    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ''Printing the amendment
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Transfer Order No not found to Print", Me.Text)
        Else
            FrmPurchaseOrderReport.PrintAbandoment(txtDocNo.Value)
        End If
    End Sub
    Private Function GetItemType() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = RowTypeItem
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = RowTypeMisc
        dt.Rows.Add(dr)
        Return dt
    End Function
    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If gv1.RowCount > 0 Then
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                If Not clsCommon.myLen(txtTransferOutNo.Value) > 0 AndAlso clsCommon.myLen(fndTransferIndentNo.Value) <= 0 Then
                    gv1.Rows.AddNew()
                    'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                    gv1.CurrentRow = gv1.Rows(intCurrRow)
                End If
            End If
        End If
    End Sub
    Private Sub cbotransferType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTransferType.SelectedValueChanged
        If Not IsFormLoad Then
            gv1.Columns(colIName).ReadOnly = True
            gv1.Columns(colUnit).ReadOnly = True
            gv1.Columns(colOutQty).ReadOnly = False
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                pnlLoadIn.Visible = False
                pnlRMDA.Visible = False
                ''richa 03/03/2015
                chkAgainst_Form.Enabled = True
                If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
                    chkAgainst_Form.Checked = True
                Else
                    chkAgainst_Form.Checked = False
                End If
                '------
                ''richa agarwal 19 july,2016
                StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                If StrCrateTransferFromBooking = "1" Then
                    GatePassPanel.Visible = True
                    lblpricecode.Visible = True
                    fndPriceCode.Visible = True
                    LoadBlankGrid()
                Else
                    FndGatePassNo.Value = ""
                    fndPriceCode.Value = ""
                    GatePassPanel.Visible = False
                    lblpricecode.Visible = True
                    fndPriceCode.Visible = True
                End If
                LoadGPItemType()
                ''-------------
                '===================added By preeti Gupta===================
                Dim ShowCrateJaaliBox As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, Nothing) = "1", True, False))
                If ShowCrateJaaliBox = True Then
                    txtcrateout.Enabled = True
                    txtjaaliout.Enabled = True
                    txtBoxOut.Enabled = True
                    txtCrateIn.Visible = False
                    txtJaaliIn.Visible = False
                    txtBoxIn.Visible = False
                    lblCrateIn.Visible = False
                    lblJaaliIn.Visible = False
                    lblBoxIn.Visible = False
                End If
                '===========================================================
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                pnlLoadIn.Visible = True
                pnlRMDA.Visible = False
                gv1.Columns(colOutQty).ReadOnly = False
                ''richa 03/03/2015
                chkAgainst_Form.Enabled = False
                chkAgainst_Form.Checked = False
                '------
                ''richa agarwal 19 july,2016
                StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                If StrCrateTransferFromBooking = "1" Then
                    FndGatePassNo.Value = ""
                    fndPriceCode.Value = ""
                    GatePassPanel.Visible = False
                    lblpricecode.Visible = True
                    fndPriceCode.Visible = True
                    LoadBlankGrid()
                End If
                LoadGPItemType()
                '===================added By preeti Gupta===================
                Dim ShowCrateJaaliBox As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.ShowCrateJaaliBoxIntransfer, clsFixedParameterCode.ShowCrateJaaliBoxIntransfer, Nothing) = "1", True, False))
                If ShowCrateJaaliBox = True Then
                    txtcrateout.Enabled = False
                    txtjaaliout.Enabled = False
                    txtBoxOut.Enabled = False
                    txtCrateIn.Visible = True
                    txtJaaliIn.Visible = True
                    txtBoxIn.Visible = True
                    lblCrateIn.Visible = True
                    lblJaaliIn.Visible = True
                    lblBoxIn.Visible = True
                End If
                '===========================================================
                ''-------------
            ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "R") = CompairStringResult.Equal Then
                pnlLoadIn.Visible = False
                pnlRMDA.Visible = True
                ''richa 03/03/2015
                chkAgainst_Form.Enabled = True
                chkAgainst_Form.Checked = False
                '------
                ''richa agarwal 19 july,2016
                StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                If StrCrateTransferFromBooking = "1" Then
                    FndGatePassNo.Value = ""
                    fndPriceCode.Value = ""
                    GatePassPanel.Visible = False
                    lblpricecode.Visible = True
                    fndPriceCode.Visible = True
                    LoadBlankGrid()
                End If
                LoadGPItemType()
                ''-------------
            End If
        End If
    End Sub
    Private Sub rbtnTaxCalAutomatic_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnTaxCalAutomatic.ToggleStateChanged, rbtnTaxCalManual.ToggleStateChanged
        If Not isInsideLoadData Then
            If rbtnTaxCalAutomatic.IsChecked Then
                SetTaxDetails()
                chkExciseOnQty.Enabled = True
            ElseIf rbtnTaxCalManual.IsChecked Then
                chkExciseOnQty.Checked = False
                chkExciseOnQty.Enabled = False
                For intRowNo As Integer = 0 To gv2.Rows.Count - 1
                    gv2.Rows(intRowNo).Cells(colTTaxRate).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTTaxAmt).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTBaseAmt).Value = Nothing
                Next
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    For ii As Integer = 1 To 10
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                    Next
                Next
            End If
        End If
    End Sub
    Private Sub gv2_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv2.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If (e.Column Is gv2.Columns(colTTaxAmt)) Then
                    gv2.CurrentRow.Cells(colTTaxAmt).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                ElseIf (e.Column Is gv2.Columns(colTTaxRate)) Then
                    gv2.CurrentRow.Cells(colTTaxRate).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                End If
            End If
            'Dim cell As GridDataCellElement = TryCast(e.CellElement, GridDataCellElement)
            'cell.GradientStyle = GradientStyles.Solid
            'cell.BackColor = Color.FromArgb(243, 181, 51)
            'End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub
    Private Sub gv2_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv2.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedTaxOpen Then
                    isCellValueChangedTaxOpen = True
                    'If (e.Column Is (gv2.Columns(colTTaxAmt)) AndAlso rbtnTaxCalManual.IsChecked) Then
                    If (rbtnTaxCalManual.IsChecked) Then
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            UpdateCurrentRow(ii)
                        Next
                        UpdateAllTotals()
                    End If
                    isCellValueChangedTaxOpen = False
                End If
            End If
        Catch ex As Exception
            isCellValueChangedTaxOpen = False
            isCellValueChangedOpen = False
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprint.Click
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Transfer Order No not found to Print", Me.Text)
        Else
            i = 1
            funprint(i)
            i = 0
        End If
    End Sub
    Private Sub btnprintChallan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrintChallan.Click
        PrintChallanReport(Me.Form_ID, Me.Text, txtDocNo.Value, Nothing)
    End Sub

    Sub PrintChallanReport(ByVal strFormID As String, ByVal strFormName As String, ByVal strCode As String, ByVal isCancelOrDelete As String)
        Try
            Dim strQuery As String
            Dim dtDocdate As Date?
            dtDocdate = Nothing
            dtDocdate = (clsCommon.GetPrintDate(clsDBFuncationality.getSingleValue("select TSPL_TRANSFER_ORDER_HEAD.document_date from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + strCode + "' "), "dd/MMM/yyyy"))
            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal OrElse clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsJobWorkType from TSPL_TRANSFER_ORDER_HEAD where Document_No='" + strCode + "'")) = 1 Then
                strQuery = PrintChallan(strFormID, strFormName, strCode, isCancelOrDelete)
                dt = clsDBFuncationality.GetDataTable(strQuery)
                If dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "RptTransferChallanJW_Local", "DELIVERY CHALLAN/STOCK SHIFTING CHALLAN", dtDocdate)
                    frmCRV = Nothing
                End If
            Else
                Dim dtBarCode As New DataTable
                dtBarCode.Columns.Add("BarCodeImage", GetType(Byte()))
                Dim bytes() As Byte
                Dim BitmapConverter As System.ComponentModel.TypeConverter = System.ComponentModel.TypeDescriptor.GetConverter(clsCommon.MyBarcodeImage(strCode, 1, False).[GetType]())
                bytes = DirectCast(BitmapConverter.ConvertTo(clsCommon.MyBarcodeImage(strCode, 1, False), GetType(Byte())), Byte())
                strQuery = clsTransferDCC.GetAttachQry(strCode, isCancelOrDelete)
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQuery)
                dt.Columns.Add("BarCodeImage", GetType(Byte()))
                For Each dr As DataRow In dt.Rows
                    dr("BarCodeImage") = bytes
                Next
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    Throw New Exception("No Data found to print")
                End If
                If dt.Rows.Count > 0 Then
                    SetItemWiseTax(dt, strCode)
                    Dim frmCRV As New frmCrystalReportViewer()
                    If objCommonVar.IsKDIL = True Then
                        'frmInventoryReportViewer.funreport(dt, "crptStockTransferChallanInvoiceInterState", "Transfer")
                        frmCRV.funsubreportWithdt(strFormID, CrystalReportFolder.InventoryReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptStockTransferChallanInvoiceInterState_Challan", "Challan", "rptCompanyAddress.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso dt.Rows(0)("Taxable").ToString() = "T" Then
                        frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "crptStockTransferChallanInvoice_ChallanNew_Product", "Challan", dtDocdate)
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso dt.Rows(0)("Taxable").ToString() = "NT" Then
                        frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "crptStockTransferChallanInvoice_ChallanNew_Milk", "Challan", dtDocdate)
                    Else
                        frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "crptStockTransferChallanInvoice_ChallanNew", "Challan", dtDocdate)
                    End If
                    frmCRV = Nothing
                End If
            End If
            'If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Transfer_Type")), "J") = CompairStringResult.Equal Then
            '    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GUNTUR") = CompairStringResult.Equal Then
            '        PurchaseOrderViewer.funreport(dt, "WO-G", "Work Order")
            '    ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "VIZAG") = CompairStringResult.Equal Then
            '        PurchaseOrderViewer.funreport(dt, "WO-V", "Work Order")
            '    Else
            '        Throw New Exception("Invalid company")
            '    End If
            'ElseIf clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Transfer_Type")), "L") = CompairStringResult.Equal Then
            '    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GUNTUR") = CompairStringResult.Equal Then
            '        SetItemWiseTax(dt, txtDocNo.Value)
            '        PurchaseOrderViewer.funreport(dt, "PO-G", "Purchase Order")
            '    ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "VIZAG") = CompairStringResult.Equal Then
            '        dt.Columns.Add("CreatedBy", Type.GetType("System.String"))
            '        For Each dr As DataRow In dt.Rows
            '            dr("Range") = "BHIMILI"
            '            dr("DivisionCommission") = "||| Vishakapatnam"
            '            dr("CreatedBy") = objCommonVar.CurrentUser
            '        Next
            '        SetItemWiseTax(dt, txtDocNo.Value)
            '        PurchaseOrderViewer.funreport(dt, "PO-V", "Purchase Order")
            '    Else
            '        Throw New Exception("Invalid company")
            '    End If
            'Else
            '    Throw New Exception("Not a valid Po Type")
            'End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnpreprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpreprint.Click
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Transfer Order No not found to Print", Me.Text)
        Else
            i = 2
            funprint(i)
            i = 0
        End If
    End Sub
    Private Sub txtRGPNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean)
        SelectRGPItems()
    End Sub
    Sub SelectRGPItems()
        'isInsideLoadData = True
        'Dim frm As New frmPendingRGP()
        'frm.strCurrCode = txtDocNo.Value
        'frm.ShowDialog()
        'If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
        '    If clsCommon.myLen(frm.ArrReturn(0).RGP_No) > 0 Then
        '        Dim objMRNHead As clsRGPHead = clsRGPHead.GetData(frm.ArrReturn(0).RGP_No, NavigatorType.Current)
        '        If objMRNHead IsNot Nothing AndAlso clsCommon.myLen(objMRNHead.RGP_No) > 0 Then
        '            'If clsCommon.myLen(txtCarrier.Text) <= 0 Then
        '            '    txtVehicleNo.Text = objMRNHead.VehicleNo
        '            'End If
        '            'If clsCommon.myLen(txtRemarks.Text) <= 0 Then
        '            '    txtRemarks.Text = objMRNHead.Remarks
        '            'End If
        '            If (clsCommon.myLen(txtFromLocation.Value) <= 0) Then
        '                txtFromLocation.Value = objMRNHead.Location
        '                lblFromLoc.Text = objMRNHead.LocationName
        '            End If
        '        End If
        '    End If
        '    If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
        '        gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
        '    End If
        '    For Each obj As clsRGPDetail In frm.ArrReturn
        '        If IsValidItemForRGP(obj) Then
        '            gv1.Rows.AddNew()
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
        '            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colRGPNo).Value = obj.RGP_No
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
        '            cboTransferType.SelectedIndex = GetItemType(obj.Item_Code)
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
        '            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
        '            ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = obj.RGP_Qty
        '            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = obj.RGP_Qty
        '        End If
        '    Next
        '    SetitemWiseTaxSetting(False, False)
        '    For ii As Integer = 0 To gv1.RowCount - 1
        '        UpdateCurrentRow(ii)
        '    Next
        'End If
        'isInsideLoadData = False
        'UpdateAllTotals()
        'cboTransferType.SelectedValue = "O"
    End Sub
    'Sub RefreshGRPNo()
    '    txtRGPNo.Value = ""
    '    If gv1.Rows.Count > 0 Then
    '        For ii As Integer = 0 To gv1.Rows.Count - 1
    '            Dim strRGPNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colRGPNo).Value)
    '            If clsCommon.myLen(strRGPNo) > 0 Then
    '                txtRGPNo.Value = clsCommon.myCstr(strRGPNo)
    '                Exit Sub
    '            End If
    '        Next
    '    End If
    'End Sub
    Public Function GetTaxGrp(ByVal strItmType As String) As String
        Dim qry As String = "select Tax_Group  from TSPL_VENDOR_MASTER where Vendor_Code ='" + strItmType + "'"
        strItmType = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
        Return strItmType
    End Function
    Function IsValidItemForRGP(ByVal obj As clsRGPDetail)
        For ii As Integer = 0 To gv1.RowCount - 1
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
            Dim strAgaintRGPCode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colRGPNo).Value)
            If clsCommon.myLen(strAgaintRGPCode) > 0 AndAlso clsCommon.CompairString(strAgaintRGPCode, obj.RGP_No) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strICode, obj.Item_Code) = CompairStringResult.Equal Then
                common.clsCommon.MyMessageBoxShow("RGP No : " + obj.RGP_No + "  Item : " + obj.Item_Desc + Environment.NewLine + "Already exist at row no:" + clsCommon.myCstr(ii + 1))
                Return False
            End If
        Next
        Return True
    End Function
    Public Function GetItemType(ByVal strItmType As String) As String
        Dim qry As String = "select distinct Item_Type  from TSPL_ITEM_MASTER where Item_Code ='" + strItmType + "'"
        strItmType = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
        If strItmType = "F" Then
            strItmType = 0
        Else
            strItmType = 1
        End If
        Return strItmType
    End Function

    Public Sub PrintData(ByVal StrCode As String, Optional ByVal IsExcise As Integer = Nothing)
        PrintData(Me.Form_ID, Me.Text, False, StrCode, Nothing)
    End Sub

    Public Sub PrintData(ByVal strFormID As String, ByVal strFormText As String, ByVal isCancelOrDelete As String, ByVal StrCode As String, Optional ByVal IsExcise As Integer = Nothing)
        Dim strCancelOrDelete As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
        Else
            strCancelOrDelete = Nothing
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
        End If

        Dim frmCRV As New frmCrystalReportViewer()
        Dim strQueryKDIL As String
        Dim dtKDIL As DataTable = Nothing
        Dim IsMandiTax As Double = 0
        IsMandiTax = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code='" & txtTaxGroup.Value & "' and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y')"))

        Dim dtDocdate As Date?
        dtDocdate = Nothing
        dtDocdate = clsCommon.GetPrintDate(clsDBFuncationality.getSingleValue("select TSPL_TRANSFER_ORDER_HEAD.document_date from " & tbl_TSPL_TRANSFER_ORDER_HEAD & " where Document_No='" + StrCode + "' "), "dd/MMM/yyyy")
        If clsCommon.myLen(StrCode) > 0 Then
            Dim PrintType As String = ""
            Dim strQuery As String
            Dim dt As DataTable
            'If IsTaxable = 1 Then
            '    If clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
            '        strReportType = "L"
            '    Else
            '        strReportType = "I"
            '    End If
            'Else
            '    strReportType = "NT"
            'End If
            PrintType = "Select Item_Tax_Type from " & tbl_TSPL_TRANSFER_ORDER_HEAD & " WHERE Document_No='" + StrCode + "'"
            PrintType = clsDBFuncationality.getSingleValue(PrintType)
            If (PrintType <> "2" OrElse clsERPFuncationality.GetGSTStatus(dtDocdate) = True) Then
                strQuery = "select TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end +  case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end +  case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end + Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ' ,Phone'+TSPL_COMPANY_MASTER.Phone1 end + Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then '  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End +  case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end  as Company_Address
from TSPL_COMPANY_MASTER  
LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code  
LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State "
                Dim strCompanyAddress As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strQuery))

                Dim dtBarCode As New DataTable
                dtBarCode.Columns.Add("BarCodeImage", GetType(Byte()))
                Dim bytes() As Byte
                Dim BitmapConverter As System.ComponentModel.TypeConverter = System.ComponentModel.TypeDescriptor.GetConverter(clsCommon.MyBarcodeImage(StrCode, 1, False).[GetType]())
                bytes = DirectCast(BitmapConverter.ConvertTo(clsCommon.MyBarcodeImage(StrCode, 1, False), GetType(Byte())), Byte())
                strQuery = "select '" & strCancelOrDelete & "' As isCancelled, TSPL_TRANSFER_ORDER_HEAD.Tax_Group,ISNULL(TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted,0) AS Is_Tax_Exempted ,TSPL_TRANSFER_ORDER_HEAD.Is_MandiTax, TSPL_TRANSFER_ORDER_HEAD.Electronic_Ref_No,TSPL_LOCATION_MASTER.GSTNO as GSTIN_No ,TSPL_STATE_MASTER.GST_STATE_Code as From_Gst_StateCode,StateMaster_ToLocation.GST_STATE_Code as To_Loc_GSTStateCode,TSPL_LOCATION_MASTER_1.GSTNO as To_Loc_GSTINNo, TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_ITEM_MASTER.HSN_Code ,TSPL_LOCATION_MASTER.GSTNO as frm_GSTINNo ,TSPL_STATE_MASTER.GST_STATE_Code as Frm_StateGST,StateMaster_ToLocation.GST_STATE_Code as To_StateGST,TSPL_LOCATION_MASTER_1.GSTNO as To_GSTINNo,TSPL_TRANSFER_ORDER_HEAD.EWayBillNo ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as EWayBillDate ,convert(varchar(10),TSPL_COMPANY_MASTER.insurance_valid_date,103) as insurance_valid_date, TSPL_COMPANY_MASTER.Pan_No as ToLoc_PANNO, 
TSPL_TRANSFER_ORDER_HEAD.Is_taxable,TSPL_TRANSFER_ORDER_HEAD.For_Repair,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer, TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1 as dTAX1, TSPL_TRANSFER_ORDER_DETAIL.TAX2 as dTAX2, TSPL_TRANSFER_ORDER_DETAIL.TAX3 as  dTAX3, TSPL_TRANSFER_ORDER_DETAIL.TAX4 as  dTAX4, TSPL_TRANSFER_ORDER_DETAIL.TAX5 as  dTAX5, TSPL_TRANSFER_ORDER_DETAIL.TAX6 as  dTAX6, TSPL_TRANSFER_ORDER_DETAIL.TAX7 as  dTAX7, TSPL_TRANSFER_ORDER_DETAIL.TAX8 as dTAX8, TSPL_TRANSFER_ORDER_DETAIL.TAX9 as dTAX9, TSPL_TRANSFER_ORDER_DETAIL.TAX10 as  dTAX10, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, 
TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate as dTAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate as dTAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate as dTAX3_Rate
,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate as dTAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate as dTAX5_Rate 
,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate as dTAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate as dTAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate as dTAX8_Rate
,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate as dTAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate as dTAX10_Rate
,tax1.Type as tax1Type,tax2.Type as tax2Type,tax3.Type as tax3Type,tax4.Type as tax4Type,tax5.Type as tax5Type,tax6.Type as tax6Type,tax7.Type as tax7Type,tax8.Type as tax8Type,tax9.Type as tax9Type,tax10.Type as tax10Type
,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.crate_In else TSPL_TRANSFER_ORDER_HEAD.crate_out end as crate,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.jaali_in else TSPL_TRANSFER_ORDER_HEAD.jaali_Out end as jaali,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.box_in else TSPL_TRANSFER_ORDER_HEAD.box_Out end as box,TSPL_TRANSFER_ORDER_DETAIL.Line_No, '1' as CopyType,0 as Alter_UnitQty ,tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt,TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate, From_GP_Location.add1 +case when len(From_GP_Location.add2)>0 then ', '+From_GP_Location.add2 else '' end +case when LEN(isnull(From_GP_Location.Add3,''))>0 then ', '+isnull(From_GP_Location.Add3,'') else ' ' end  + case when len(From_GP_Location_State.STATE_NAME   )>0 then ', '+ From_GP_Location_State.STATE_NAME  else ' ' end    as From_Location_Address_GP, TSPL_TRANSFER_ORDER_HEAD.Transfer_Type, From_GP_Location.Location_Code as From_Location_GP_Code,From_GP_Location.Location_Desc as From_Location_GP_Name,From_GP_Location.Add1 as From_GP_Add1,From_GP_Location.Add2 as From_GP_Add2,From_GP_Location.Add3 as From_GP_Add3,From_GP_Location.Add4 as From_GP_Add4,From_GP_Location_State.STATE_NAME as From_GP_State_Name, 
From_GP_Location.TIN_No as From_GP_TINNO, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Location_Code else TSPL_LOCATION_MASTER_1.Location_Code end as To_location_GP_Code,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Location_Desc else TSPL_LOCATION_MASTER_1.Location_Desc end as To_Location_GP_Name,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add1 else TSPL_LOCATION_MASTER_1.Add1 end as To_GP_Add1,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add2 else TSPL_LOCATION_MASTER_1.Add2 end as To_GP_Add2,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add3 else TSPL_LOCATION_MASTER_1.Add3 end as To_GP_Add3,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add4 else TSPL_LOCATION_MASTER_1.Add4 end as To_GP_Add4,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location_State.STATE_NAME else StateMaster_ToLocation.STATE_NAME end as TO_GP_State_Name,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.TIN_No else TSPL_LOCATION_MASTER_1.TIN_No end  as To_GP_TINNO ,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.TAN_No else TSPL_LOCATION_MASTER_1.TAN_No end as TO_GP_FAX,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then isnull(To_GP_Location.Pin_Code,0) else isnull(TSPL_LOCATION_MASTER_1.Pin_Code,0) end as To_GP_Loc_Pin,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.City_Code else TSPL_LOCATION_MASTER_1.City_Code end as To_GP_City_Code,  
case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then Case when len(ISNULL(To_GP_Location.Phone1,''))>0 and To_GP_Location.Phone1='(+__)__________' then '' else To_GP_Location.Phone1 end + Case When   ISNULL(To_GP_Location.Phone2,'')<>'(+__)__________' Then '  '+ To_GP_Location.Phone2 Else'' end else Case when len(ISNULL(TSPL_LOCATION_MASTER_1.Phone1,''))>0 and TSPL_LOCATION_MASTER_1.Phone1='(+__)__________' then '' else TSPL_LOCATION_MASTER_1.Phone1 end + Case When   ISNULL(TSPL_LOCATION_MASTER_1.Phone2,'')<>'(+__)__________' Then '  '+ TSPL_LOCATION_MASTER_1.Phone2 Else'' end end To_GP_Phn,  
TSPL_LOCATION_MASTER.TIN_No as Loc_Tin_No,TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end  + case when len(TSPL_STATE_MASTER.STATE_NAME   )>0 then ', '+ TSPL_STATE_MASTER.STATE_NAME  else ' ' end    as Location_Address_GP,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_STATE_MASTER.state_code as frm_State_code,tspl_location_master.HOAdd1 ,TSPL_LOCATION_MASTER .HOAdd2,TSPL_TRANSFER_ORDER_HEAD.GR_No,TSPL_TRANSFER_ORDER_HEAD.document_type ,case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date ,TSPL_TRANSFER_ORDER_HEAD.WayBill_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as WayBill_Date,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,tspl_location_master_For_Location.City_Code as Location_City_Name, coalesce(cast(convert(decimal(18,0),(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * tspl_item_uom_detail.conversion_factor)/alt_convrsn.conversion_factor) as varchar)+' '+TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,'') as Alt_Unit_Code,TSPL_TRANSPORT_MASTER.Transporter_Name,TSPL_TRANSFER_ORDER_HEAD.transport_id,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual ,(case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF=1 then 'Against F-Form Due' else '' end) as Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Document_No  as[STN_NO] , tspl_transfer_order_head.Document_Date as [Date_N_Time_issue], 
TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as Discount , TSPL_TRANSFER_ORDER_DETAIL .Document_No as ref_doc_no , 
TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_TRANSFER_ORDER_HEAD.To_Location ,TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , 
TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1, 
TSPL_LOCATION_MASTER.Location_Code as From_Location_Code,  TSPL_LOCATION_MASTER.Location_Desc as From_Location_Dec,TSPL_LOCATION_MASTER.Add1 as From_Location_Add1 , 
TSPL_LOCATION_MASTER.Add2 as From_Location_Add2 , TSPL_LOCATION_MASTER.Add3 as From_Location_Add3,TSPL_LOCATION_MASTER.Add4  as From_Location_Add4,TSPL_LOCATION_MASTER.City_Code as From_Location_City_Code ,  
TSPL_STATE_MASTER.STATE_NAME  as From_Location_State,TSPL_LOCATION_MASTER.Pin_Code as From_Location_Pin_Code ,TSPL_LOCATION_MASTER.Country as From_Location_Country,  TSPL_LOCATION_MASTER.Telphone   as From_Location_Telphone
,TSPL_LOCATION_MASTER.Email as From_Location_Email,TSPL_LOCATION_MASTER.TIN_No AS From_Location_tin_no, TSPL_LOCATION_MASTER.CST_No AS From_Location_cstNo,TSPL_ITEM_MASTER.Weight_UOM as UOM2
,TSPL_ITEM_MASTER.Weight_Value as Weight, TSPL_LOCATION_MASTER_1.Location_Code AS To_Location_Code,TSPL_LOCATION_MASTER.add1 as transpotor,TSPL_LOCATION_MASTER_1.Add1 AS To_Location_Add1,  
TSPL_LOCATION_MASTER_1.Add2 as To_Location_Add2 ,TSPL_LOCATION_MASTER_1.Add3 as To_Location_Add3,TSPL_LOCATION_MASTER_1.Add4 as To_Location_Add4, TSPL_LOCATION_MASTER_1.Location_Desc as To_Location_Desc  , 
TSPL_LOCATION_MASTER_1.City_Code as To_Location_City_Code , StateMaster_ToLocation.State_Name as To_Location_State, TSPL_LOCATION_MASTER_1.Pin_Code as To_Location_Pin_Code,  TSPL_LOCATION_MASTER_1.Country as To_Location_Country, 
TSPL_LOCATION_MASTER_1.Telphone as To_Location_Telphone, TSPL_LOCATION_MASTER_1.Email as To_Location_Email ,  TSPL_LOCATION_MASTER_1 .TIN_No as to_location_tin_no, TSPL_LOCATION_MASTER_1 .CST_No as to_location_cstno,TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX3,TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX6  
,TSPL_COMPANY_MASTER.Comp_Name AS CompName  
,'" + strCompanyAddress + "' as Company_Address
, TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount 
from " & tbl_TSPL_TRANSFER_ORDER_Detail & " 
left outer join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No 
left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
left outer join TSPL_LOCATION_MASTER AS TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.GIT_LOCATION =  TSPL_TRANSFER_ORDER_HEAD.To_Location  
INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code  
left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code  
left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id  
LEFT OUTER JOIN tspl_location_master  AS tspl_location_master_For_Location ON tspl_location_master_For_Location.GIT_Location =TSPL_TRANSFER_ORDER_HEAD.To_Location  
left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
LEFT OUTER JOIN TSPL_STATE_MASTER StateMaster_ToLocation ON StateMaster_ToLocation.State_Code=TSPL_LOCATION_MASTER_1.State 
left join TSPL_TRANSFER_ORDER_HEAD GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno 
left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location 
left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State  
left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  GP.To_Location  
left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State  
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10  
LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' 
where  TSPL_TRANSFER_ORDER_HEAD. Document_No = '" + StrCode + "'"
                'If Not AllowOneFormatByDefault Then ' added by preeti gupta Against ticket no[MIL/16/05/19-000086]
                '    strQueryKDIL = "Select * from (" & strQuery & ") XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 "
                '    If Not AllowThreeFormatByDefault Then
                '        strQueryKDIL += " UNION Select '1' as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1 "
                '    End If
                '    strQueryKDIL += " ) YYY ON YYY.COL1=XXX.CopyType  ORDER BY YYY.COL2,Line_No "
                '    dtKDIL = clsDBFuncationality.GetDataTable(strQueryKDIL)
                'Else
                '    strQueryKDIL = "Select * from (" & strQuery & ") XXX"
                '    dtKDIL = clsDBFuncationality.GetDataTable(strQueryKDIL)
                'End If
                'If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal OrElse (clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHBA") = CompairStringResult.Equal AndAlso chkForRepair.Checked = False) Then
                '    strQuery = PrintChallan(StrCode)
                '    dt = clsDBFuncationality.GetDataTable(strQuery)
                'Else
                '    dt = clsDBFuncationality.GetDataTable(strQuery)
                '    If clsCommon.myCBool(dt.Rows(0)("For_Repair")) = True Then
                '        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHBA") = CompairStringResult.Equal Then
                '            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_Local_repair", "Transfer Local", dtDocdate, "rptCompanyAddress.rpt")
                '        Else
                '            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_Local", "Transfer Local", dtDocdate, "rptCompanyAddress.rpt")
                '        End If
                '        Exit Sub
                '    End If
                'End If

                dt = clsDBFuncationality.GetDataTable(strQuery)
                dt.Columns.Add("BarCodeImage", GetType(Byte()))
                For Each dr As DataRow In dt.Rows
                    dr("BarCodeImage") = bytes
                Next
            Else
                If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHBA") = CompairStringResult.Equal Then
                    strQuery = PrintChallan(StrCode)
                    dt = clsDBFuncationality.GetDataTable(strQuery)
                Else
                    strQuery = "Select '" & strCancelOrDelete & "' As isCancelled,* from (" &
            " select From_GP_Location.add1 +case when len(From_GP_Location.add2)>0 then ', '+From_GP_Location.add2 else '' end +case when LEN(isnull(From_GP_Location.Add3,''))>0 then ', '+isnull(From_GP_Location.Add3,'') else ' ' end  + case when len(From_GP_Location_State.STATE_NAME   )>0 then ', '+ From_GP_Location_State.STATE_NAME  else ' ' end    as From_Location_Address_GP, TSPL_TRANSFER_ORDER_HEAD.Transfer_Type, From_GP_Location.Location_Code as From_Location_GP_Code,From_GP_Location.Location_Desc as From_Location_GP_Name,From_GP_Location.Add1 as From_GP_Add1,From_GP_Location.Add2 as From_GP_Add2,From_GP_Location.Add3 as From_GP_Add3,From_GP_Location.Add4 as From_GP_Add4,From_GP_Location_State.STATE_NAME as From_GP_State_Name, From_GP_Location.TIN_No as From_GP_TINNO,To_GP_Location.Location_Code as To_location_GP_Code,To_GP_Location.Location_Desc as To_Location_GP_Name,To_GP_Location.Add1 as To_GP_Add1,To_GP_Location.Add2 as To_GP_Add2,To_GP_Location.Add3 as To_GP_Add3,To_GP_Location.Add4 as To_GP_Add4,To_GP_Location_State.STATE_NAME as TO_GP_State_Name,To_GP_Location.TIN_No  as To_GP_TINNO,LMFrom.TIN_No as Loc_Tin_No,LMFrom.add1 +case when len(LMFrom.add2)>0 then ', '+LMFrom.add2 else '' end +case when LEN(isnull(LMFrom.Add3,''))>0 then ', '+isnull(LMFrom.Add3,'') else ' ' end  + case when len(SMFrom.STATE_NAME   )>0 then ', '+ SMFrom.STATE_NAME  else ' ' end    as Location_Address_GP, TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, LMTo.city_code as City_Name ,TSPL_TRANSFER_ORDER_HEAD.Remarks as Description ,TSPL_TRANSFER_ORDER_DETAIL.Line_No,SMFrom.STATE_NAME as Loc_State_Name, SMFrom.state_code as frm_State_code, LMFrom.HOAdd1, LMFrom.HOAdd2, tspl_company_master.cst_lst as Comp_CSt_LST," &
            " TSPL_TRANSFER_ORDER_HEAD.Transport_Id, TSPL_TRANSPORT_MASTER.Transporter_Name,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual, TSPL_TRANSFER_ORDER_HEAD.GR_No as GRNO," &
            " case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date," &
            " TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No, 0 as Alter_UnitQty, TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as HeadDisc_Amt, 0 as HeadDisc_PerAmt, TSPL_ITEM_MASTER.Cheapter_Heads," &
            " TSPL_ITEM_master.ITF_CODE as Chap_Desc, LMFrom.Registration_Number, case when TSPL_TRANSFER_ORDER_HEAD.Terms_Code='A' then 'Advance' else TSPL_TRANSFER_ORDER_HEAD.Terms_Code end as Payment_Terms," &
            " TSPL_TRANSFER_ORDER_HEAD.Modify_By, TSPL_COMPANY_MASTER.Pan_No as Comp_PANNO, " &
            " LMFrom.add1 +case when len(LMFrom.add2)>0 then ', '+LMFrom.add2 else '' end +case when LEN(isnull(LMFrom.Add3,''))>0 then ', '+isnull(LMFrom.Add3,'') else ' ' end + case when LEN(CMFrom.City_Name)>0 then ', '+CMFrom .City_Name else ' ' end + case when len(SMFrom.STATE_NAME  )>0 then ', '+ SMFrom.STATE_NAME else ' ' end  + case when len(LMFrom.Pin_Code   )>0 then ', Pin Code - '+ cast(LMFrom.Pin_Code  as varchar)  else ' ' end  + case when len(LMFrom.Tin_No     )>0 then ', Tin No - '+ cast(LMFrom.Tin_No as varchar)  else ' ' end  + Case when len(ISNULL(LMFrom.Phone1,''))>0 and LMFrom.Phone1='(+__)__________' then '' else ', Phone'+LMFrom.Phone1 end +  Case When   ISNULL(LMFrom.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End   + case when len(LMFrom.Email    )>0 then ', Email - '+ LMFrom.Email else '' end  as Location_Address," &
            " LMFrom.CST_No as Loc_CSTNo, LMFrom.Excisable as loc_Excisable,LMFrom.Range_Address as Loc_range_Add,LMFrom.Division_Address  as loc_Division_Address,LMFrom.Commissionerate  as Loc_Commissionerate," &
            " '' as Challan_No, '' as Challan_Date, '' as Removal_Date, TSPL_TRANSFER_ORDER_HEAD.WayBill_No,TSPL_TRANSFER_ORDER_HEAD.For_Repair," &
            " TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(CM_Company.City_Name)>0 then ', '+CM_Company.City_Name else ' ' end + case when len(SMCompany.STATE_NAME  )>0 then ', '+ SMCompany.STATE_NAME else ' ' end  + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.Tin_No)>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.Pan_No)>0 then ', PAN No - '+ cast(TSPL_COMPANY_MASTER.Pan_No as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end+ Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ',Phone'+TSPL_COMPANY_MASTER.Phone1 end +  Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ',  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End   + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end  as Comp_Address," &
            " LMFrom .Add1 as Loc_Add1, LMFrom.Add2 as Loc_ADd2, LMFrom.Add3 as Loc_Add3, LMFrom.Pin_Code as Loc_Pin_Code, LMFrom.TIN_No as Loc_TinNo, Case when ISNULL(LMFrom.Phone1,'')='(+__)__________' then '' else LMFrom.Phone1 end +  Case When   ISNULL(LMFrom.Phone2,'')<>'(+__)__________' Then ', '+ LMFrom.Phone2 Else'' End as  Loc_Phn," &
            " LMFrom.Email as Loc_Email, 'Excise Invoice' as Invoice_Type, (case when len(isnull(TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,''))>0 then (Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty Else TSPL_TRANSFER_ORDER_DETAIL.In_Qty End)/case when alt_convrsn.Conversion_Factor<=0 then 1 else alt_convrsn.Conversion_Factor end else null end) as Alternet_Qty," &
            " TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code as Alternate_UOM, 0 as Scheme_Qty, '' as Scheme_Item_UOM, TSPL_TRANSFER_ORDER_HEAD.Discount_Base, TSPL_TRANSFER_ORDER_HEAD.GR_No as Dis_Doc_No," &
            "  TSPL_COMPANY_MASTER .State as Comp_State, '' as Buyer_order_no, '' as Buyer_order_date, '' as Terms_of_delivery, TSPL_TRANSFER_ORDER_HEAD.Document_No as InvoiceNo," &
            " TSPL_TRANSFER_ORDER_HEAD.Document_Date as Date_Time_Invoice, convert(varchar ,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as InvoiceDate, '' as ShipmentNo, 0 as Alt_Qty," &
            " TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code as Alt_UOM, '' as ShipmentDate, '' as DeliveryOrderNo, TSPL_TRANSFER_ORDER_HEAD.Terms_Code as TermCondition, LMFrom.Location_Desc , TSPL_COMPANY_MASTER.Comp_Name as CompName, ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+  Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone, TSPL_COMPANY_MASTER.Fax as CompFax,TSPL_COMPANY_MASTER.Email as ComMail,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.Pan_No as ComPanNo,  TSPL_COMPANY_MASTER.CST_LST as CompCSTLST,TSPL_COMPANY_MASTER.Pincode as ComPINCode, TSPL_COMPANY_MASTER .Tin_No as ComTinNO,ISNULL(tspl_company_Master.ADD1,'') as Compaddress1,ISNULL(tspl_company_Master.ADD2,'')   as Compaddress2,ISNULL(tspl_company_Master.ADD3,'') as Compaddress3," &
            " '' as P_Add1, '' as P_Add2, '' as P_Add3, '' as P_PinNo, '' as P_CstNo, '' as P_TinNo, '' as P_Email, '' as P_Fax, '' as P_LstNo, '' as P_CustCode, '' as P_Cust_Name, '' as P_City_Name, '' as P_State_Name, '' as P_Cust_Phn," &
            " LMTo.Location_Code as Cust_Code, LMTo.Location_Desc as Customer_Name, LMTo.Add1 as Cust_Add1, LMTo.Add2 as Cust_add2, LMTo.Add3 as cust_add3, case when ISNULL(LMTo.Phone1,'')='(+__)__________' then '' else LMTo.Phone1 end +  Case When   ISNULL(LMTo.Phone2,'')<>'(+__)__________' Then ', '+ LMTo.Phone2 Else'' End as Cust_Phn,LMTo.Tin_No  as Cust_TinNo, LMTo.CST_No as Cust_CSTNo, '' Cust_LSTNo, LMTo.Email as Cust_Email, LMTo.PIN_Code as Cust_PinCode, CMTo.City_Name as Cust_City_Name, '' as Cust_Fax, SMTo.STATE_NAME as Cust_State_Name, TSPL_TRANSFER_ORDER_DETAIL.item_code, TSPL_ITEM_MASTER.Item_Desc as itemdesc, Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty Else TSPL_TRANSFER_ORDER_DETAIL.In_Qty End as qty, TSPL_TRANSFER_ORDER_DETAIL.mrp, TSPL_TRANSFER_ORDER_DETAIL.amount as amount, TSPL_TRANSFER_ORDER_DETAIL.unit_code as uom, '' as RATE_UOM ,TSPL_TRANSFER_ORDER_DETAIL.item_cost as itemcost, tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate,TSPL_TRANSFER_ORDER_DETAIL.Disc_Per, isnull(TSPL_TRANSFER_ORDER_HEAD.Discount_Amt,0) as Discount_Amt,isnull(TSPL_TRANSFER_ORDER_HEAD.Amount_Less_Discount,0) as Amount_Less_Discount,  isnull(TSPL_TRANSFER_ORDER_HEAD .Total_Tax_Amt,0) as Total_Tax_Amt, TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt as Total_Amt, '1' as CopyType," &
            " '' as Add_Charge_Name1, 0 as Add_Charge_Amt1, '' as Add_Charge_Name2, 0 as Add_Charge_Amt2, '' as Add_Charge_Name3, 0 as Add_Charge_Amt3, '' as Add_Charge_Name4, 0 as Add_Charge_Amt4, '' as Add_Charge_Name5, 0 as Add_Charge_Amt5, '' as Add_Charge_Name6, 0 as Add_Charge_Amt6, '' as Add_Charge_Name7, 0 as Add_Charge_Amt7, '' as Add_Charge_Name8, 0 as Add_Charge_Amt8, '' as Add_Charge_Name9, 0 as Add_Charge_Amt9, '' as Add_Charge_Name10, 0 as Add_Charge_Amt10"
                    '========Sanjeet(UDL)21/11/2016======
                    strQuery += ", TSPL_TRANSFER_ORDER_HEAD.Modify_Date,'' AS Time_of_Prepration,'' AS Time_of_Removal,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code, " &
                    "TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Tcan_No AS WebSite,TSPL_COMPANY_MASTER.Ecc_No," &
                    "TSPL_COMPANY_MASTER.tin_no,TSPL_COMPANY_MASTER.ServiceTax_Reg_No,AM.Abatement_Percent,TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt," &
                    "BI.BATCH_NO, TSPL_TRANSFER_ORDER_HEAD.Gross_Item_Wt,TSPL_TRANSFER_ORDER_HEAD.Total_Item_Wt," &
                    "(TSPL_COMPANY_MASTER.Add1+' '+TSPL_COMPANY_MASTER.Add2+' '+TSPL_COMPANY_MASTER.Add3+' '+TSPL_COMPANY_MASTER.State) as COMP_ADDRESS2," &
                    "TSPL_COMPANY_MASTER.CE_Commissionerate,'' as Tariff,TSPL_COMPANY_MASTER.Access_Officer as FSSAI,LMFROM.Location_Desc AS LOCFROM_Customer_Name,'' as Customer_EccNo,From_GP_Location.Loc_Short_Name,case when len(From_GP_Location.Pin_Code   )>0 then From_GP_Location.Pin_Code ELSE Null END AS From_Location_Pin,case when len(From_GP_Location.Phone1   )>0 then From_GP_Location.Phone1 when len(From_GP_Location.Phone2)>0 then From_GP_Location.Phone2 ELSE Null END AS From_Location_phone," &
            "TSPL_COMPANY_MASTER.Insurance_No,TSPL_COMPANY_MASTER.Insurance_Comp_Name,CONVERT(VARCHAR(15),TSPL_COMPANY_MASTER.Insurance_Valid_Date,103) AS Insurance_Valid_Date,TSPL_TRANSFER_ORDER_HEAD.Remarks," &
                    "TSPL_TRANSFER_ORDER_HEAD.Secondary_Code,(CASE WHEN IT.Is_Tax_Exempted=2 THEN TSPL_TRANSFER_ORDER_DETAIL.MRP ELSE 0.00 END)AS Item_MRP," &
                    "TAX1=" &
                        "REPLACE((CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX1!='' THEN(Tax1.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt)) ELSE '' END)+CHAR(13)+" &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX2!='' THEN(TAX2.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX2_Amt))ELSE '' END) +CHAR(13)+ " &
                    "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX3!=''  THEN(TAX3.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX3_Amt))ELSE '' END) +CHAR(13)+ " &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX4!='' THEN(TAX4.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX4_Amt))ELSE '' END) +CHAR(13)+" &
                    "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX5!='' THEN(TAX5.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX5_Amt))ELSE '' END) +CHAR(13)+" &
                    "	(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX6!='' THEN(TAX6.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX6_Amt))ELSE '' END) +CHAR(13)+" &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX7!='' THEN(TAX7.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX7_Amt))ELSE '' END) +CHAR(13)+" &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX8!='' THEN(TAX8.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX8_Amt))ELSE '' END) +CHAR(13)+" &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX9!='' THEN(TAX9.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX9_Amt))ELSE '' END) +CHAR(13)+" &
                        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX10!=''  THEN(TAX10.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX10_Amt))ELSE '' END),CHAR(13),'')"
                    '===============================
                    strQuery += " from " & tbl_TSPL_TRANSFER_ORDER_Detail & "" + Environment.NewLine &
            " join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE" + Environment.NewLine &
            " left outer join TSPL_LOCATION_MASTER LMFrom on LMFrom.Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location" + Environment.NewLine &
            " left outer join TSPL_LOCATION_MASTER LMTo on LMTo.GIT_Location =  TSPL_TRANSFER_ORDER_HEAD.To_Location" + Environment.NewLine &
            " Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_CITY_MASTER  AS CM_Company ON CM_Company.City_Code =TSPL_COMPANY_MASTER.City_Code" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_CITY_MASTER  AS CMFrom ON CMFrom.City_Code =LMFrom.City_Code" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_CITY_MASTER  AS CMTo ON CMTo.City_Code =LMTo.City_Code" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_STATE_MASTER AS SMCompany  ON SMCompany.STATE_CODE  =TSPL_COMPANY_MASTER.State" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_STATE_MASTER SMFrom on SMFrom.STATE_CODE=LMFrom.State" + Environment.NewLine &
            " LEFT OUTER JOIN TSPL_STATE_MASTER SMTo ON SMTo.State_Code=LMTo.State" + Environment.NewLine &
            " left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code" + Environment.NewLine &
            " left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id" + Environment.NewLine &
            " left outer join TSPL_CHAPTER_HEAD on TSPL_CHAPTER_HEAD.Chapter_Head_Code =TSPL_ITEM_MASTER.Cheapter_Heads" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1" + Environment.NewLine &
            " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2" + Environment.NewLine &
            " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9" + Environment.NewLine &
            " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10" + Environment.NewLine &
            " left join " & tbl_TSPL_TRANSFER_ORDER_HEAD & " GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State  left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  GP.To_Location  left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State "
                    '=============sanjeet=====
                    strQuery += "Left Join TSPL_BATCH_ITEM AS BI ON TSPL_TRANSFER_ORDER_DETAIL.Document_No=BI.Document_Code AND TSPL_TRANSFER_ORDER_DETAIL.Line_No=BI.Parent_Line_No " &
            "AND TSPL_TRANSFER_ORDER_HEAD.Transfer_Type=BI.In_Out_Type  Left JOIN TSPL_ITEM_MASTER IT ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code=IT.Item_Code " &
            "LEFT JOIN TSPL_ABATEMENT_MASTER AM ON AM.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code"
                    '=================
                    strQuery += " where 2=2 and  TSPL_TRANSFER_ORDER_HEAD. Document_No = '" & StrCode & "'" + Environment.NewLine &
            " ) XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 "
                    If Not AllowThreeFormatByDefault Then
                        strQuery += " UNION Select '1' as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1 "
                    End If
                    strQuery += " ) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2,xxx.Line_No "
                    dt = clsDBFuncationality.GetDataTable(strQuery)
                End If
            End If
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("No Data found to print")
            End If
            Dim currentCompanyCode As String = clsDBFuncationality.getSingleValue("select * from TSPL_COMPANY_MASTER ")
            If dt.Rows.Count > 0 Then
                SetItemWiseTax(dt, StrCode)
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "crptStockTransferChallanInvoice", "Transfer", dtDocdate)


                'Sanjay Ticket No- MIL/31/03/19-000057 Add GMD Company code
                '        If (objCommonVar.IsKDIL = True OrElse clsCommon.CompairString(currentCompanyCode, "GDFPL") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "GHO") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "MPD") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "BHAD") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "GMD") = CompairStringResult.Equal) Then
                '            If PrintType = "2" AndAlso clsERPFuncationality.GetGSTStatus(dtDocdate) = False Then
                '                Dim Qry2 As String = "select TSPL_TRANSFER_ORDER_HEAD.Document_No as InvoiceNo, TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt, TSPL_ITEM_MASTER.Item_Code," &
                '                " TSPL_ITEM_MASTER.Item_Desc + '( MRP : ' +   convert(varchar,TSPL_TRANSFER_ORDER_DETAIL.MRP) + '  Abatement : ' + convert(varchar,convert(int,100- TSPL_TRANSFER_ORDER_DETAIL.Abatement_Per)) + '%)' as Item_Desc, " &
                '                " TSPL_TRANSFER_ORDER_DETAIL.TAX1, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate," &
                '                " TSPL_TRANSFER_ORDER_DETAIL.TAX2, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate," &
                '                " TSPL_TRANSFER_ORDER_DETAIL.TAX3 ,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate," &
                '                " TSPL_ITEM_MASTER.Cheapter_Heads, tax1.Tax_Code_Desc as tax1name,tax2.Tax_Code_Desc as tax2name,tax3.Tax_Code_Desc as tax3name" &
                '" from TSPL_TRANSFER_ORDER_DETAIL" &
                '" left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER .Item_Code =TSPL_TRANSFER_ORDER_DETAIL.Item_Code" &
                '" left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL.tax1" &
                '" left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2" &
                '" left outer join tspl_tax_master as tax3 on tax3.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX3" &
                '" LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No" &
                '" where TSPL_TRANSFER_ORDER_HEAD.Document_No ='" & StrCode & "'" &
                '" order by TSPL_TRANSFER_ORDER_DETAIL .line_no "
                '                Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(Qry2)
                '                'KwalitySalesReportViewer.funsubreportWithdt(dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptProductSaleInvoice", "Retail Invoice", "rptCompanyAddress.rpt")
                '                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, dt2, "rptProductExciseTransfer", "Excise Transfer", dtDocdate, "rptSubReportExciseTransferSaleInvoice.rpt", "rptCompanyAddress.rpt", clsERPFuncationality.CompanyAddresShowinFooter())
                '                'KwalitySalesReportViewer.funsubreportWithdt(dt, dt2, "rptProductExciseTransferSaleInvoice", "Excise Transfer", "rptSubReportExciseTransferSaleInvoice.rpt")
                '            Else
                '                'frmInventoryReportViewer.funreport(dt, "crptStockTransferChallanInvoiceInterState", "Transfer")
                '                ''as per preeti gupta
                '                'Client - KDIL Ticket No- KDI/25/09/18-000433 Remove name from Transfer printout,RptTransfer_Interstate,RptTransfer_Interstate_withmanditax,RptTransfer_Local,RptTransfer_Local_Withmanditax  
                '                If (objCommonVar.IsKDIL = True OrElse clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GHO") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "MPD") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "BHAD") = CompairStringResult.Equal OrElse clsCommon.CompairString(currentCompanyCode, "GMD") = CompairStringResult.Equal) Then
                '                    If clsERPFuncationality.GetGSTStatus(dtDocdate) Then
                '                        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("From_Location_State")), clsCommon.myCstr(dt.Rows(0)("To_Location_State"))) = CompairStringResult.Equal Then
                '                            If clsCommon.myCdbl(dt.Rows(0)("Is_MandiTax")) = 1 Then
                '                                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_Local_WithMandiTax", "Transfer Local With MandiTax", dtDocdate, "rptCompanyAddress.rpt")
                '                            Else
                '                                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_Local", "Transfer Local", dtDocdate, "rptCompanyAddress.rpt")
                '                            End If
                '                        Else
                '                            If IsMandiTax > 0 Then
                '                                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_InterState_WithMandiTax", "Transfer Local", dtDocdate, "rptCompanyAddress.rpt")
                '                            ElseIf clsCommon.CompairString(clsCommon.myCdbl(dt.Rows(0)("Is_Tax_Exempted")), 1) = CompairStringResult.Equal Then
                '                                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_Local", "Transfer Local", dtDocdate, "rptCompanyAddress.rpt")
                '                            Else
                '                                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "RptTransfer_InterState", "Transfer InterState", dtDocdate, "rptCompanyAddress.rpt")
                '                            End If
                '                        End If
                '                    Else
                '                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "rptProductSaleInvoice1", "Transfer", dtDocdate, "rptCompanyAddress.rpt")
                '                    End If
                '                ElseIf clsCommon.CompairString(currentCompanyCode, "GDFPL") = CompairStringResult.Equal Then
                '                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dtKDIL, clsERPFuncationality.CompanyAddresShowinFooter(), "rptTransferDepot", "Transfer", dtDocdate, "rptCompanyAddress.rpt")
                '                Else
                '                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptStockTransferChallanInvoiceInterState", "Transfer", "rptCompanyAddress.rpt")
                '                End If
                '            End If
                '        ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHBA") = CompairStringResult.Equal Then
                '            If clsERPFuncationality.GetGSTStatus(dtDocdate) Then
                '                If clsCommon.myCBool(dt.Rows(0)("InternalTransfer")) = True Then
                '                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "RptTransfer_Internal", "Internal Transfer", dtDocdate)
                '                ElseIf clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("To_Gp_stateCode")), clsCommon.myCstr(dt.Rows(0)("frm_state_code"))) = CompairStringResult.Equal Then
                '                    If clsCommon.myCdbl(dt.Rows(0)("Is_MandiTax")) = 1 Then
                '                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "RptTransfer_LocalWithMandiTax", "Transfer Excise Invoice", dtDocdate)
                '                    Else
                '                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "RptTransfer_Local", "Transfer Excise Invoice", dtDocdate)
                '                    End If
                '                Else
                '                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "RptTransfer_InterState", "Transfer Excise Invoice", dtDocdate)
                '                End If
                '            Else
                '                If clsCommon.CompairString(MyBase.Form_ID, clsCommon.myCstr(IsExcise), "1") = CompairStringResult.Equal Then
                '                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "RptTransfer_ExcisableNormal", "Transfer Excise Invoice", dtDocdate)
                '                    'frmCrystalReportViewer.funreport(CrystalReportFolder.InventoryReport, dt, "rptTransferChallanInvoice", "Challan/Transfe Invoice")
                '                Else
                '                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "rptDepTransfer_Challan", "Challan/Transfe", dtDocdate)
                '                End If
                '            End If
                '        Else
                '            SetItemWiseTax(dt, StrCode)
                '            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "crptStockTransferChallanInvoice", "Transfer", dtDocdate)
                '        End If
            End If
        End If
        frmCRV = Nothing
    End Sub

    Private Function PrintChallan(ByVal strCode As String) As String
        Return PrintChallan(Me.Form_ID, Me.Text, strCode, Nothing)
    End Function

    Private Function PrintChallan(ByVal strFormID As String, ByVal strFormName As String, ByVal strCode As String, ByVal isCancelOrDelete As String) As String
        Dim strCancelOrDelete As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
        Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strCancelOrDelete = "Cancelled"
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
        Else
            strCancelOrDelete = Nothing
            tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
            tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
        End If
        Dim StrGheeItem As Double = 0
        If clsCommon.myLen(strCode) > 0 Then
            StrGheeItem = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*) from tspl_item_master where Item_Code in (select Item_Code from TSPL_TRANSFER_ORDER_DETAIL where Document_no='" & strCode & "') and Structure_Code in('GHEE','CGHEE')"))
        End If
        Dim DateOfEInvoiceImplementation As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DateOfEInvoiceImplementation, clsFixedParameterCode.DateOfEInvoiceImplementation, Nothing))
        Dim strQuery As String = Nothing
        strQuery = " select '" & strCancelOrDelete & "' As IsCancelled, cast(TSPL_TRANSFER_ORDER_HEAD.BarCode_Img as image) As BarCode_Img,isnull (TSPL_TRANSFER_ORDER_HEAD.IRN_No,'') as IRN_No,isnull (TSPL_TRANSFER_ORDER_HEAD.Ack_No,'') as Ack_No,case when len(isnull (TSPL_TRANSFER_ORDER_HEAD.Ack_No,'')) > 0 then convert (varchar, TSPL_TRANSFER_ORDER_HEAD.Ack_Date,103) else ''  end as Ack_Date, case when TSPL_TRANSFER_ORDER_HEAD.Is_Taxable=1 and isnull(TSPL_TRANSFER_ORDER_HEAD.EInvoice_Type,'')='BB' AND convert(date ,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)>=convert(date ,'" + clsCommon.myCstr(DateOfEInvoiceImplementation) + "',103) then 1 else 0 end as  IsEInvoiceApply," &
            " " & StrGheeItem & " AS CheckForGhee,(case when " & StrGheeItem & " >0 THEN 'This is Stock Transfer being Dispatched to above Destination Court order and other relevent documents copy enclosed.' else '' end) as Special_Instruction," &
            " LMFrom.GSTNO fromGSTNo,SMFrom.gst_state_code as FromStateCode,TSPL_COMPANY_MASTER.GSTINNo AS COMP_GSTIN_NO,SMCompany.GST_STATE_Code AS Comp_GST_StateCode,To_GP_Location.GSTNO as To_Gp_GSTIN_NO,To_GP_Location_State.GST_STATE_Code AS To_GP_GST_StateCode,TSPL_ITEM_MASTER.HSN_Code ,TSPL_TRANSFER_ORDER_HEAD.EWayBillNo," &
        " To_GP_Location_State.STATE_CODE as To_Gp_stateCode,SMCompany.STATE_CODE as Comp_StateCode,TSPL_TRANSFER_ORDER_HEAD.Status,TSPL_TRANSFER_ORDER_HEAD.Is_MandiTax,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer, TSPL_COMPANY_MASTER.Circle_No,To_GP_Location.Ecc_Number as To_Gp_EccNo,Case when len(ISNULL(tspl_company_master.Phone1,''))>0 and tspl_company_master.Phone1='(+__)__________' then '' else tspl_company_master.Phone1 end +','+ Case When   ISNULL(tspl_company_master.Phone2,'')<>'(+__)__________' Then '  '+ tspl_company_master.Phone2 Else'' end as comp_Phn" &
        " ,tspl_company_master.CINNo,convert(varchar,tspl_company_master.TinNo_Issue_Date,103) as TinNo_Issue_Date,convert(varchar,tspl_company_master.PanNo_Issue_Date,103) as PanNo_Issue_Date, case when TSPL_ITEM_MASTER.is_scheme_item=1 then 'Y' else 'N' end as is_scheme_item ,To_GP_Location.CST_No as to_GP_CST_no,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as EWayBillDate,TSPL_TRANSFER_ORDER_HEAD.Electronic_Ref_No,From_GP_Location.add1 +case when len(From_GP_Location.add2)>0 then ', '+From_GP_Location.add2 else '' end +case when LEN(isnull(From_GP_Location.Add3,''))>0 then ', '+isnull(From_GP_Location.Add3,'') else ' ' end  + case when len(From_GP_Location_State.STATE_NAME   )>0 then ', '+ From_GP_Location_State.STATE_NAME  else ' ' end    as From_Location_Address_GP, TSPL_TRANSFER_ORDER_HEAD.Transfer_Type,isnull(TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF,0) as Is_AgainstFormF, From_GP_Location.Location_Code as From_Location_GP_Code,From_GP_Location.Location_Desc as From_Location_GP_Name,From_GP_Location.Add1 as From_GP_Add1,From_GP_Location.Add2 as From_GP_Add2,From_GP_Location.Add3 as From_GP_Add3,From_GP_Location.Add4 as From_GP_Add4,From_GP_Location_State.STATE_NAME as From_GP_State_Name, From_GP_Location.TIN_No as From_GP_TINNO,To_GP_Location.Location_Code as To_location_GP_Code,To_GP_Location.Location_Desc as To_Location_GP_Name,To_GP_Location.Add1 as To_GP_Add1,To_GP_Location.Add2 as To_GP_Add2,To_GP_Location.Add3 as To_GP_Add3,To_GP_Location.Add4 as To_GP_Add4,To_GP_Location_State.STATE_NAME as TO_GP_State_Name,To_GP_Location.TIN_No  as To_GP_TINNO,LMFrom.TIN_No as Loc_Tin_No," &
        " To_GP_Location.add1 +case when len(To_GP_Location.add2)>0 then ', '+To_GP_Location.add2 else '' end +case when LEN(isnull(To_GP_Location.Add3,''))>0 then ', '+isnull(To_GP_Location.Add3,'') else ' ' end  + case when len(To_GP_Location_State.STATE_NAME   )>0 then ', '+ To_GP_Location_State.STATE_NAME  else ' ' end    " &
      " as Location_Address_GP,  From_GP_Location.Loc_Short_Name as From_GP_Short_Name,To_GP_Location.Loc_Short_Name as To_Gp_Loc_Short_Name,LMFrom.Loc_Short_Name as LmFrom_Short_Name, " &
        " TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, To_GP_Location.city_code as City_Name ,TO_LOCATION_CITY.City_Name AS Destination_City,TSPL_TRANSFER_ORDER_HEAD.Remarks as Description ,TSPL_TRANSFER_ORDER_DETAIL.Line_No,SMFrom.STATE_NAME as Loc_State_Name, SMFrom.state_code as frm_State_code, LMFrom.HOAdd1, LMFrom.HOAdd2, tspl_company_master.cst_lst as Comp_CSt_LST," &
" TSPL_TRANSFER_ORDER_HEAD.Transport_Id, TSPL_TRANSPORT_MASTER.Transporter_Name,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual, TSPL_TRANSFER_ORDER_HEAD.GR_No as GRNO," &
" case when len(TSPL_TRANSPORT_MASTER.Transporter_Name)>0 and TSPL_TRANSPORT_MASTER.Transporter_Name <>'' then TSPL_TRANSPORT_MASTER.Transporter_Name else TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual end as Transporter, " &
" case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date," &
" TSPL_TRANSFER_ORDER_HEAD.Vehicle_No, 0 as Alter_UnitQty, TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as HeadDisc_Amt, 0 as HeadDisc_PerAmt, TSPL_ITEM_MASTER.Cheapter_Heads," &
" TSPL_ITEM_master.ITF_CODE as Chap_Desc, LMFrom.Registration_Number, case when TSPL_TRANSFER_ORDER_HEAD.Terms_Code='A' then 'Advance' else TSPL_TRANSFER_ORDER_HEAD.Terms_Code end as Payment_Terms," &
" TSPL_TRANSFER_ORDER_HEAD.Modify_By, TSPL_COMPANY_MASTER.Pan_No as Comp_PANNO, " &
" LMFrom.add1 +case when len(LMFrom.add2)>0 then ', '+LMFrom.add2 else '' end +case when LEN(isnull(LMFrom.Add3,''))>0 then ', '+isnull(LMFrom.Add3,'') else ' ' end + case when LEN(CMFrom.City_Name)>0 then ', '+CMFrom .City_Name else ' ' end + case when len(SMFrom.STATE_NAME  )>0 then ', '+ SMFrom.STATE_NAME else ' ' end  + case when len(LMFrom.Pin_Code   )>0 then ', Pin Code - '+ cast(LMFrom.Pin_Code  as varchar)  else ' ' end  + case when len(LMFrom.Tin_No     )>0 then ', Tin No - '+ cast(LMFrom.Tin_No as varchar)  else ' ' end  + Case when len(ISNULL(LMFrom.Phone1,''))>0 and LMFrom.Phone1='(+__)__________' then '' else ', Phone'+LMFrom.Phone1 end +  Case When   ISNULL(LMFrom.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End   + case when len(LMFrom.Email    )>0 then ', Email - '+ LMFrom.Email else '' end  as Location_Address," &
" LMFrom.CST_No as Loc_CSTNo, LMFrom.Excisable as loc_Excisable,LMFrom.Range_Address as Loc_range_Add,LMFrom.Division_Address  as loc_Division_Address,LMFrom.Commissionerate  as Loc_Commissionerate," &
" '' as Challan_No, '' as Challan_Date, '' as Removal_Date, TSPL_TRANSFER_ORDER_HEAD.WayBill_No," &
" TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(CM_Company.City_Name)>0 then ', '+CM_Company.City_Name else ' ' end + case when len(SMCompany.STATE_NAME  )>0 then ', '+ SMCompany.STATE_NAME else ' ' end  + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.Tin_No)>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.Pan_No)>0 then ', PAN No - '+ cast(TSPL_COMPANY_MASTER.Pan_No as varchar)  else ' ' end  + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end+ Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ',Phone'+TSPL_COMPANY_MASTER.Phone1 end +  Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ',  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End   + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end  as Comp_Address," &
" LMFrom .Add1 as Loc_Add1, LMFrom.Add2 as Loc_ADd2, LMFrom.Add3 as Loc_Add3, LMFrom.Pin_Code as Loc_Pin_Code, LMFrom.TIN_No as Loc_TinNo, Case when ISNULL(LMFrom.Phone1,'')='(+__)__________' then '' else LMFrom.Phone1 end +  Case When   ISNULL(LMFrom.Phone2,'')<>'(+__)__________' Then ', '+ LMFrom.Phone2 Else'' End as  Loc_Phn," &
" LMFrom.Email as Loc_Email, TSPL_TRANSFER_ORDER_HEAD.TYPE  as Invoice_Type, (case when len(isnull(TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,''))>0 then (Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty Else TSPL_TRANSFER_ORDER_DETAIL.In_Qty End)/case when alt_convrsn.Conversion_Factor<=0 then 1 else alt_convrsn.Conversion_Factor end else null end) as Alternet_Qty," &
" TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code as Alternate_UOM, 0 as Scheme_Qty, '' as Scheme_Item_UOM, TSPL_TRANSFER_ORDER_HEAD.Discount_Base, TSPL_TRANSFER_ORDER_HEAD.GR_No as Dis_Doc_No," &
"  TSPL_COMPANY_MASTER .State as Comp_State, '' as Buyer_order_no, '' as Buyer_order_date, '' as Terms_of_delivery, TSPL_TRANSFER_ORDER_HEAD.Document_No as InvoiceNo," &
" TSPL_TRANSFER_ORDER_HEAD.Document_Date as Date_Time_Invoice, convert(varchar ,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as InvoiceDate, '' as ShipmentNo, 0 as Alt_Qty," &
" TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code as Alt_UOM, '' as ShipmentDate, '' as DeliveryOrderNo, TSPL_TRANSFER_ORDER_HEAD.Terms_Code as TermCondition, LMFrom.Location_Desc , TSPL_COMPANY_MASTER.Comp_Name as CompName, ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+  Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone, TSPL_COMPANY_MASTER.Fax as CompFax,TSPL_COMPANY_MASTER.Email as ComMail,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.Pan_No as ComPanNo,  TSPL_COMPANY_MASTER.CST_LST as CompCSTLST,TSPL_COMPANY_MASTER.Pincode as ComPINCode, TSPL_COMPANY_MASTER .Tin_No as ComTinNO,ISNULL(tspl_company_Master.ADD1,'') as Compaddress1,ISNULL(tspl_company_Master.ADD2,'')   as Compaddress2,ISNULL(tspl_company_Master.ADD3,'') as Compaddress3," &
" '' as P_Add1, '' as P_Add2, '' as P_Add3, '' as P_PinNo, '' as P_CstNo, '' as P_TinNo, '' as P_Email, '' as P_Fax, '' as P_LstNo, '' as P_CustCode, '' as P_Cust_Name, '' as P_City_Name, '' as P_State_Name, '' as P_Cust_Phn," &
" LMTo.Location_Code as Cust_Code, LMTo.Location_Desc as Customer_Name, LMTo.Add1 as Cust_Add1, LMTo.Add2 as Cust_add2, LMTo.Add3 as cust_add3, case when ISNULL(LMTo.Phone1,'')='(+__)__________' then '' else LMTo.Phone1 end +  Case When   ISNULL(LMTo.Phone2,'')<>'(+__)__________' Then ', '+ LMTo.Phone2 Else'' End as Cust_Phn,LMTo.Tin_No  as Cust_TinNo, LMTo.CST_No as Cust_CSTNo, '' Cust_LSTNo, LMTo.Email as Cust_Email, LMTo.PIN_Code as Cust_PinCode, CMTo.City_Name as Cust_City_Name, '' as Cust_Fax, SMTo.STATE_NAME as Cust_State_Name, TSPL_TRANSFER_ORDER_DETAIL.item_code, TSPL_ITEM_MASTER.Item_Desc as itemdesc, Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty Else TSPL_TRANSFER_ORDER_DETAIL.In_Qty End as qty, " &
" case when TSPL_ITEM_MASTER.Is_Batch_Item=1 then BI.QTY else Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty Else TSPL_TRANSFER_ORDER_DETAIL.In_Qty End end  AS Item_Qty,  case when TSPL_ITEM_MASTER.Is_Batch_Item=1 then COALESCE(BI.QTY,0)* COALESCE(TSPL_TRANSFER_ORDER_DETAIL.amount,0)/Case When TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='O' Then COALESCE(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty,0) Else COALESCE(TSPL_TRANSFER_ORDER_DETAIL.In_Qty,0) End else COALESCE(TSPL_TRANSFER_ORDER_DETAIL.amount,0) end  AS Item_Amount,case when TSPL_ITEM_MASTER.Is_Batch_Item=1 then BI.MRP else (CASE WHEN IT.Is_Tax_Exempted=2 THEN TSPL_TRANSFER_ORDER_DETAIL.MRP ELSE 0.00 END) end AS  ItemB_mrp,case when TSPL_ITEM_MASTER.Is_Batch_Item=1 THEN BI.UOM ELSE TSPL_TRANSFER_ORDER_DETAIL.unit_code END AS Item_B_uom,case when TSPL_ITEM_MASTER.Is_Batch_Item=1 THEN BI.BATCH_NO ELSE '' END AS Item_BatchNo,IT.Is_Tax_Exempted,TSPL_ITEM_MASTER.Is_Batch_Item, " &
" TSPL_TRANSFER_ORDER_DETAIL.mrp, TSPL_TRANSFER_ORDER_DETAIL.amount as amount, TSPL_TRANSFER_ORDER_DETAIL.unit_code as uom, '' as RATE_UOM ,TSPL_TRANSFER_ORDER_DETAIL.item_cost as itemcost, tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate," &
"ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt,0) AS DTax1_Amt,  ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt,0) AS DTax2_Amt," &
 " ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt,0) AS DTax3_Amt,  ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt,0) AS DTax4_Amt, " &
  " ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt,0) AS DTax5_Amt,  ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt,0) AS DTax6_Amt, " &
   " ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt,0) AS DTax7_Amt,  ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt,0) AS DTax8_Amt, " &
     " ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt,0) AS DTax9_Amt,  ISNULL(TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt,0) AS DTax10_Amt, " &
" TSPL_TRANSFER_ORDER_DETAIL.Disc_Per, isnull(TSPL_TRANSFER_ORDER_HEAD.Discount_Amt,0) as Discount_Amt,isnull(TSPL_TRANSFER_ORDER_HEAD.Amount_Less_Discount,0) as Amount_Less_Discount,  isnull(TSPL_TRANSFER_ORDER_HEAD .Total_Tax_Amt,0) as Total_Tax_Amt, TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt as Total_Amt, '1' as CopyType," &
" '' as Add_Charge_Name1, 0 as Add_Charge_Amt1, '' as Add_Charge_Name2, 0 as Add_Charge_Amt2, '' as Add_Charge_Name3, 0 as Add_Charge_Amt3, '' as Add_Charge_Name4, 0 as Add_Charge_Amt4, '' as Add_Charge_Name5, 0 as Add_Charge_Amt5, '' as Add_Charge_Name6, 0 as Add_Charge_Amt6, '' as Add_Charge_Name7, 0 as Add_Charge_Amt7, '' as Add_Charge_Name8, 0 as Add_Charge_Amt8, '' as Add_Charge_Name9, 0 as Add_Charge_Amt9, '' as Add_Charge_Name10, 0 as Add_Charge_Amt10 ," &
        " Dtax1.Type as Taxtype1, Dtax2.Type as Taxtype2, Dtax3.Type as Taxtype3, Dtax4.Type as Taxtype4, Dtax5.Type as Taxtype5," &
 " Dtax6.Type as Taxtype6, Dtax7.Type as Taxtype7, Dtax8.Type as Taxtype8, Dtax9.Type as Taxtype9, Dtax10.Type as Taxtype10, "
        strQuery += " TSPL_TRANSFER_ORDER_HEAD.Modify_Date,'' AS Time_of_Prepration,'' AS Time_of_Removal,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code, " &
        "TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Tcan_No AS WebSite,TSPL_COMPANY_MASTER.Ecc_No," &
        "TSPL_COMPANY_MASTER.tin_no,TSPL_COMPANY_MASTER.ServiceTax_Reg_No,AM.Abatement_Percent,TSPL_TRANSFER_ORDER_DETAIL.Abatement_Amt," &
        "BI.BATCH_NO, TSPL_TRANSFER_ORDER_HEAD.Gross_Item_Wt,TSPL_TRANSFER_ORDER_HEAD.Total_Item_Wt," &
        "(TSPL_COMPANY_MASTER.Add1+' '+TSPL_COMPANY_MASTER.Add2+' '+TSPL_COMPANY_MASTER.Add3+' '+TSPL_COMPANY_MASTER.State) as COMP_ADDRESS2," &
        "TSPL_COMPANY_MASTER.CE_Commissionerate,'' as Tariff,TSPL_COMPANY_MASTER.Access_Officer as FSSAI,LMFROM.Location_Desc AS LOCFROM_Customer_Name,'' as Customer_EccNo,From_GP_Location.Loc_Short_Name,case when len(From_GP_Location.Pin_Code   )>0 then From_GP_Location.Pin_Code ELSE Null END AS From_Location_Pin,case when len(From_GP_Location.Phone1   )>0 then From_GP_Location.Phone1 when len(From_GP_Location.Phone2)>0 then From_GP_Location.Phone2 ELSE Null END AS From_Location_phone," &
        "TSPL_COMPANY_MASTER.Insurance_No,TSPL_COMPANY_MASTER.Insurance_Comp_Name,CONVERT(VARCHAR(15),TSPL_COMPANY_MASTER.Insurance_Valid_Date,103) AS Insurance_Valid_Date,TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_TRANSFER_ORDER_HEAD.For_Repair," &
        "TSPL_TRANSFER_ORDER_HEAD.Secondary_Code,(CASE WHEN IT.Is_Tax_Exempted=2 THEN TSPL_TRANSFER_ORDER_DETAIL.MRP ELSE 0.00 END)AS Item_MRP," &
        "TAX1=" &
            "REPLACE((CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX1!='' THEN(Tax1.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX1_Amt)) ELSE '' END)+CHAR(13)+" &
            "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX2!='' THEN(TAX2.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX2_Amt))ELSE '' END) +CHAR(13)+ " &
        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX3!=''  THEN(TAX3.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX3_Amt))ELSE '' END) +CHAR(13)+ " &
            "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX4!='' THEN(TAX4.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX4_Amt))ELSE '' END) +CHAR(13)+" &
        "(CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX5!='' THEN(TAX5.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX5_Amt))ELSE '' END) +CHAR(13)+" &
        " (CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX6!='' THEN(TAX6.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX6_Amt))ELSE '' END) +CHAR(13)+" &
            " (CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX7!='' THEN(TAX7.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX7_Amt))ELSE '' END) +CHAR(13)+" &
            " (CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX8!='' THEN(TAX8.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX8_Amt))ELSE '' END) +CHAR(13)+" &
            " (CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX9!='' THEN(TAX9.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX9_Amt))ELSE '' END) +CHAR(13)+" &
            " (CASE WHEN TSPL_TRANSFER_ORDER_HEAD.TAX10!=''  THEN(TAX10.Tax_Code_Desc+' ( '+CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate)+' % ) '+ CONVERT(NVARCHAR(10),TSPL_TRANSFER_ORDER_HEAD.TAX10_Amt))ELSE '' END),CHAR(13),'') ,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No "
        strQuery += " from " & tbl_TSPL_TRANSFER_ORDER_Detail & "" + Environment.NewLine &
" join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE" + Environment.NewLine &
" left outer join TSPL_LOCATION_MASTER LMFrom on LMFrom.Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location" + Environment.NewLine &
" left outer join TSPL_LOCATION_MASTER LMTo on LMTo.GIT_Location =  TSPL_TRANSFER_ORDER_HEAD.To_Location" + Environment.NewLine &
" Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_CITY_MASTER  AS CM_Company ON CM_Company.City_Code =TSPL_COMPANY_MASTER.City_Code" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_CITY_MASTER  AS CMFrom ON CMFrom.City_Code =LMFrom.City_Code" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_CITY_MASTER  AS CMTo ON CMTo.City_Code =LMTo.City_Code" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_STATE_MASTER AS SMCompany  ON SMCompany.STATE_CODE  =TSPL_COMPANY_MASTER.State" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_STATE_MASTER SMFrom on SMFrom.STATE_CODE=LMFrom.State" + Environment.NewLine &
" LEFT OUTER JOIN TSPL_STATE_MASTER SMTo ON SMTo.State_Code=LMTo.State" + Environment.NewLine &
" left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code" + Environment.NewLine &
" left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id" + Environment.NewLine &
" left outer join TSPL_CHAPTER_HEAD on TSPL_CHAPTER_HEAD.Chapter_Head_Code =TSPL_ITEM_MASTER.Cheapter_Heads" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1" + Environment.NewLine &
" left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2" + Environment.NewLine &
" left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10" + Environment.NewLine &
" left outer join TSPL_TAX_MASTER as Dtax1 on Dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL.tax1 " &
" left outer join tspl_tax_master as Dtax2 on Dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2" &
 " left outer join tspl_tax_master as Dtax3 on Dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3 " &
 " left outer join TSPL_TAX_MASTER as Dtax4 on Dtax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4 " &
 " left outer join TSPL_TAX_MASTER as Dtax5 on Dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5 " &
 " left outer join TSPL_TAX_MASTER as Dtax6 on Dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6 " &
 " left outer join TSPL_TAX_MASTER as Dtax7 on Dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7 " &
  " left outer join TSPL_TAX_MASTER as Dtax8 on Dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8 " &
 " left outer join TSPL_TAX_MASTER as Dtax9 on Dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9 " &
 " left outer join TSPL_TAX_MASTER as Dtax10 on Dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10 "
        If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
            strQuery += " left join TSPL_TRANSFER_ORDER_HEAD_Cancel_Data  GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
        ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
            strQuery += " left join  TSPL_TRANSFER_ORDER_HEAD_Delete_Data GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
        Else
            strQuery += " left join TSPL_TRANSFER_ORDER_HEAD  GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
        End If
        strQuery += " left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location  "
        strQuery += " left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State "
        ''richa agarwal 13 Nov,2019 add case return type of transfer UDL/13/11/19-001010
        If clsCommon.myCBool(chkInternalTransfer.Checked) = True OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
            strQuery += " left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.Location_Code =  TSPL_TRANSFER_ORDER_HEAD.To_Location "
        Else
            strQuery += " left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  TSPL_TRANSFER_ORDER_HEAD.To_Location "
        End If
        strQuery += " left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State " &
" LEFT OUTER JOIN TSPL_CITY_MASTER  TO_LOCATION_CITY ON TO_LOCATION_CITY.City_Code= To_GP_Location.City_Code " &
" Left Join TSPL_BATCH_ITEM AS BI ON TSPL_TRANSFER_ORDER_DETAIL.Document_No=BI.Document_Code AND TSPL_TRANSFER_ORDER_DETAIL.Line_No=BI.Parent_Line_No "
        ''richa agarwal 13 Nov,2019 do this work for return type of transfer UDL/13/11/19-001010
        If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
            strQuery += " and BI.In_Out_Type =( case when TSPL_ITEM_MASTER.Is_Batch_Item=1 then 'I' else TSPL_TRANSFER_ORDER_HEAD.Transfer_Type end ) "
        Else
            strQuery += "AND TSPL_TRANSFER_ORDER_HEAD.Transfer_Type=BI.In_Out_Type "
        End If
        strQuery += " Left JOIN TSPL_ITEM_MASTER IT ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code=IT.Item_Code " &
        "LEFT JOIN TSPL_ABATEMENT_MASTER AM ON AM.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code " &
        " where 2=2 and  TSPL_TRANSFER_ORDER_HEAD. Document_No = '" & strCode & "'"
        Return strQuery
    End Function
    Private Sub btnPrintNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintNew.Click
        Try
            PrintData(txtDocNo.Value)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function SetItemWiseTax(ByVal dtAfterModify As DataTable, ByVal strShipFrm As String) As DataTable
        dtAfterModify.Columns.Add("TAX1_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt3", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt3", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt3", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt3", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt3", GetType(Double))
        Dim qry As String = "select Tax,Rate,SUM(Amt) as TaxAmt"
        qry += " from ("
        qry += " select TAX1 as Tax,TAX1_Rate as Rate,TAX1_Amt as Amt"
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "' "
        qry += " union all "
        qry += " select TAX2 as Tax,TAX2_Rate as Rate,TAX2_Amt as Amt "
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX3 as Tax,TAX3_Rate as Rate,TAX3_Amt as Amt "
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX4 as Tax,TAX4_Rate as Rate,TAX4_Amt as Amt "
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX5 as Tax,TAX5_Rate as Rate,TAX5_Amt as Amt "
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "'   "
        qry += " union all "
        qry += " select TAX6 as Tax,TAX6_Rate as Rate,TAX6_Amt as Amt "
        qry += " from TSPL_TRANSFER_ORDER_DETAIL where Document_No='" + strShipFrm + "'   "
        qry += " )xxx "
        qry += " group by Tax,Rate   having SUM(Amt)>0   "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                For ii As Integer = 1 To 5
                    Dim strCol As String = "TAX" + clsCommon.myCstr(ii) + ""
                    If clsCommon.CompairString(clsCommon.myCstr(dtAfterModify.Rows(0)(strCol)), clsCommon.myCstr(dr("Tax"))) = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt1")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate1") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt1") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt2")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate2") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt2") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt3")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate3") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt3") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        Else
                            Throw New Exception("Printing Support only 3 Diffent Rates")
                        End If
                    End If
                Next
            Next
        End If
        Return dtAfterModify
    End Function
    Private Sub chkExciseOnQty_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkExciseOnQty.ToggleStateChanged
        If Not isInsideLoadData Then
            For ii As Integer = 0 To gv1.Rows.Count - 1
                UpdateCurrentRow(ii)
            Next
            UpdateAllTotals()
        End If
    End Sub
    Private Sub txtCurrencyCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtCurrencyCode._MYValidating
        Dim qry As String = "select CURRENCY_CODE AS Code, CURRENCY_NAME as Name from TSPL_CURRENCY_MASTER"
        txtCurrencyCode.Value = clsCommon.ShowSelectForm("CURRENCY_MASTER", qry, "Code", "", txtCurrencyCode.Value, "CURRENCY_CODE", isButtonClicked)
        ShowCurrencyDetail()
    End Sub
    Sub ShowCurrencyDetail()
        'Dim strq As String
        Dim dt As DataTable
        If clsCommon.myLen(txtCurrencyCode.Value) > 0 Then
            dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.txtDate.Value, txtCurrencyCode.Value)
            If dt.Rows.Count = 0 Then
                If Me.txtCurrencyCode.Value = objCommonVar.BaseCurrencyCode Then
                    Me.txtConversionRate.Text = 1
                    Me.txtApplicableFrom.Text = ""
                Else
                    clsCommon.MyMessageBoxShow(Me, "Conversion rate not entered for currency '" & Me.txtCurrencyCode.Value & "'")
                    Exit Sub
                End If
            Else
                Me.txtConversionRate.Text = dt.Rows(0).Item("Rate")
                Me.txtApplicableFrom.Text = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd/MMM/yyyy")
            End If
        Else
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
        End If
    End Sub
    Private Sub cboType_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboType.Validating
        If clsCommon.CompairString(clsCommon.myCstr(cboType.SelectedValue), "Excise") = CompairStringResult.Equal Then
            Dim qry As String = "select Tax_Group_Code from TSPL_TAX_GROUP_MASTER where Is_Transfer=1 and Tax_Group_Type='S'"
            Dim strTaxGroupofTransferType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            If clsCommon.myLen(strTaxGroupofTransferType) > 0 Then
                txtTaxGroup.Value = strTaxGroupofTransferType
                If clsCommon.myLen(txtFromLocation.Value) > 0 AndAlso clsCommon.myLen(txtToLoc.Value) > 0 Then
                    strTaxType = clsLocationWiseTax.TaxType(txtFromLocation.Value, txtToLoc.Value, "T", txtDate.Value, Nothing)
                End If
                If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                    IsMandiTax = clsLocationWiseTax.IsMandiTax(txtTaxGroup.Value, Nothing)
                End If
                SetTaxDetails()
            End If
        Else
            gv2.DataSource = Nothing
            gv2.Rows.Clear()
        End If
    End Sub
    Private Sub RadPageView1_SelectedPageChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadPageView1.SelectedPageChanged
    End Sub
    Private Sub btnReverseAndUnpost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverseAndUnpost.Click
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                '' to check balance in case of transfer In 20 Oct,2020
                If clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                        Dim strIName As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                        Dim dblInQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colInQty).Value)
                        Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                        Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colMRP).Value)
                        If clsCommon.myLen(strICode) > 0 Then
                            Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strICode, txtToLoc.Value, txtDocNo.Value, txtDate.Value, Nothing, strUOM, dblMRP)
                            If dblBalQty < dblInQty Then
                                Throw New Exception("You can't reverse this document because quantity of Item - " + strICode + " goes -ve")
                            End If
                        End If
                    Next
                End If
                '---------------------
                If clsTransferDCC.ReverseAndUnpost(txtDocNo.Value, blnUpdateLoadInwithLoadOut) Then
                    txtglvoucher.Text = clsDBFuncationality.getSingleValue("select GLVoucher_No from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & txtDocNo.Value & "'")
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                    blnUpdateLoadInwithLoadOut = False
                    btnReverseAndUnpost.Visible = False
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtRMDANo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtRMDANo._MYValidating
        isInsideLoadData = True
        Dim qry As String = "select RMDA_No as Code,RMDA_Date as Date from TSPL_SRN_HEAD  "
        Dim whrclas As String = " LEN(ISNULL(TSPL_SRN_HEAD.RMDA_No,''))>0 and RMDA_No not in (select RMDA_Code from TSPL_TRANSFER_ORDER_HEAD where TSPL_TRANSFER_ORDER_HEAD.Document_No not in('" + txtDocNo.Value + "') ) and exists (select 1 from TSPL_PI_DETAIL where TSPL_PI_DETAIL.SRN_Id=TSPL_SRN_HEAD.SRN_No) "
        txtRMDANo.Value = clsCommon.ShowSelectForm("RMDAInTransfer", qry, "Code", whrclas, txtRMDANo.Value, "", isButtonClicked)
        If clsCommon.myLen(txtRMDANo.Value) > 0 Then
            LoadBlankGrid()
            Dim obj As clsSRNHead = clsSRNHead.GetData(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select SRN_No from TSPL_SRN_HEAD where RMDA_No='" + txtRMDANo.Value + "'")), NavigatorType.Current)
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.RMDA_No) > 0 Then
                txtFromLocation.Value = clsLocation.GetRejectedLocation(obj.Bill_To_Location, Nothing)
                lblFromLoc.Text = clsLocation.GetName(txtFromLocation.Value, Nothing)
                txtToLoc.Value = obj.Bill_To_Location
                lblToLoc.Text = obj.BillToLocationName
                txtRemarks.Text = obj.Remarks
                txtRefNo.Text = obj.Ref_No
            End If
            If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
                gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
            End If
            For Each objtr As clsSRNDetail In obj.Arr
                If objtr.Rejected_Qty > 0 Then
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTransferOutNo).Value = ""
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objtr.Item_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objtr.Item_Desc
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(objtr.Item_Code, Nothing)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objtr.Item_Cost
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objtr.Unit_code
                    ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                    ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = objtr.Rejected_Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = objtr.MRP
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = clsItemMaster.IsMRPItem(objtr.Item_Code)
                    'gv1.Columns(colInQty).IsVisible = True
                    'gv1.Columns(colLeakQty).IsVisible = True
                    'gv1.Columns(colBreakQty).IsVisible = True
                    'gv1.Columns(colShortQty).IsVisible = True
                    SetitemWiseTaxSetting(True, True)
                    UpdateCurrentRow(gv1.Rows.Count - 1)
                End If
            Next
        End If
        isInsideLoadData = False
        UpdateAllTotals()
    End Sub
    Private Sub txtTransporter_Code__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtTransporter_Code._MYValidating
        Qry = "select Transport_Id as Code,Transporter_Name as Description,City,State,Pincode,Phone from TSPL_TRANSPORT_MASTER"
        txtTransporter_Code.Value = clsCommon.ShowSelectForm("TRANSPORTER_Transfer_KDIL", Qry, "Code", "", txtTransporter_Code.Value, "Code", isButtonClicked)
        txtTransporter_desc.Text = clsTransferDCC.GetTransporterName(txtTransporter_Code.Value)
        Dim toloc As String = ""
        toloc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GIT_Location from TSPL_LOCATION_MASTER where Location_Code='" + txtToLoc.Value + "'", Nothing))
        'txtvehicle_Charge.Text = clsTransferDCC.GetProvisionCharge(txtFromLocation.Value, txtToLoc.Value, clsCommon.myCdbl(txtGross_Wt.Text), clsCommon.myCdbl(txtVehicle_Capacity.Text), txtTransporter_Code.Value)
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
            If clsCommon.myLen(txtVehicle_Capacity.Text) <= 0 Then
                Exit Sub
            Else
                Dim qry As String = clsDBFuncationality.getSingleValue("select freight from TSPL_ROUTE_FREIGHT_DETAILS where ToLocation_Code='" & txtToLoc.Value & "' and location_code='" & txtFromLocation.Value & "' and transport_id='" & txtTransporter_Code.Value & "' and capacityMT='" & txtVehicle_Capacity.Text & "'")
                If clsCommon.myLen(qry) > 0 Then
                    txtvehicle_Charge.Text = qry
                Else
                    txtvehicle_Charge.Text = 0
                End If
            End If
        Else
            FillVehicleCharges()
        End If
    End Sub
    Private Sub txtGross_Wt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGross_Wt.TextChanged
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
        Else
            If clsCommon.myCdbl(txtGross_Wt.Text) > 0 Then
                Dim toloc As String = ""
                toloc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GIT_Location from TSPL_LOCATION_MASTER where Location_Code='" + txtToLoc.Value + "'", Nothing))
                'txtvehicle_Charge.Text = clsTransferDCC.GetProvisionCharge(txtFromLocation.Value, txtToLoc.Value, clsCommon.myCdbl(txtGross_Wt.Text), clsCommon.myCdbl(txtVehicle_Capacity.Text), txtTransporter_Code.Value)
                FillVehicleCharges()
            Else
                txtvehicle_Charge.Text = 0
                txtvehicle_Charge.Tag = Nothing
            End If
        End If
    End Sub
    Private Sub chkOwnVehicle_CheckStateChanged(sender As Object, e As EventArgs) Handles chkOwnVehicle.CheckStateChanged
        If chkOwnVehicle.Checked = True Then
            TxtTransportorMName.Visible = True
            TxtTransportorMName.Location = New Point(569, 2)
            txtTransporter_Code.Visible = False
            txtTransporter_desc.Visible = False
            TxtTransportorMName.Text = "Own Vehicle"
        Else
            TxtTransportorMName.Visible = False
            TxtTransportorMName.Location = New Point(342, 45)
            txtTransporter_Code.Visible = True
            txtTransporter_desc.Visible = True
            TxtTransportorMName.Text = ""
        End If
    End Sub
    Private Sub chkOwnVehicle_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkOwnVehicle.ToggleStateChanged
        txtTransporter_Code.MendatroryField = Not chkOwnVehicle.Checked
        txtTransporter_Code.Enabled = Not chkOwnVehicle.Checked
        txtTransporter_Code.Value = ""
        txtTransporter_desc.Text = ""
        txtvehicle_Charge.Text = Nothing
    End Sub
    ''richa 03/03/2015
    Private Sub txtWayBill_No_TextChanged(sender As Object, e As EventArgs) Handles txtWayBill_No.TextChanged
        If clsCommon.myLen(txtWayBill_No.Text) > 0 Then
            ttxway_bill_date.Enabled = True
        Else
            ttxway_bill_date.Enabled = False
        End If
    End Sub
    Sub OpenSerialItem()
        If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsSerialseItem).Value) Then
            Dim Item_type As String = clsDBFuncationality.getSingleValue("select Item_Type from TSPL_ITEM_MASTER where Item_Code='" + gv1.CurrentRow.Cells(colICode).Value + "'")
            If clsCommon.myLen(gv1.CurrentRow.Cells(colTransferOutNo).Value) > 0 Then
                Dim frm As frmSerializeItemOut = New frmSerializeItemOut()
                frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                frm.strAgaintsDocNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colTransferOutNo).Value)
                'frm.strBinNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colBinNo).Value)
                frm.strLocationCode = txtFromLocation.Value
                frm.strCurrDocNo = txtDocNo.Value
                If chkInternalTransfer.Checked Then
                    frm.strCurrDocType = "ITransfer"
                ElseIf chkJobWorkTransfer.Checked Then
                    frm.strCurrDocType = "JTransfer"
                Else
                    frm.strCurrDocType = "Transfer"
                End If
                frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colInQty).Value)
                frm.strItemType = Item_type
                frm.arr = TryCast(gv1.CurrentRow.Tag, List(Of clsSerializeInvenotry))
                frm.ShowDialog()
                If Not frm.isCencelButtonClicked Then
                    gv1.CurrentRow.Tag = frm.arr
                End If
            Else
                Dim frm As frmSerializeItemOut = New frmSerializeItemOut()
                frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                frm.strBinNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colBinNo).Value)
                frm.strLocationCode = txtFromLocation.Value
                frm.strCurrDocNo = txtDocNo.Value
                If chkInternalTransfer.Checked Then
                    frm.strCurrDocType = "ITransfer"
                ElseIf chkJobWorkTransfer.Checked Then
                    frm.strCurrDocType = "JTransfer"
                Else
                    frm.strCurrDocType = "Transfer"
                End If
                frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colOutQty).Value)
                frm.strItemType = Item_type
                frm.arr = TryCast(gv1.CurrentRow.Tag, List(Of clsSerializeInvenotry))
                frm.ShowDialog()
                If Not frm.isCencelButtonClicked Then
                    gv1.CurrentRow.Tag = frm.arr
                End If
            End If
        End If
    End Sub
    '============created by preeti gupta===============
    Public Sub OpenBatchItemIfFIFIOSettingON()
        If clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
            Dim arr As List(Of clsBatchInventory) = Nothing
            Dim strBatchunion As String = ""
            If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
                arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
            End If
            If Not arr Is Nothing Then
                If arr.Count > 0 Then
                    For Each obj As clsBatchInventory In arr
                        strBatchunion += " Batch No - " & clsCommon.myCstr(obj.Batch_No) & "         Qty - " & clsCommon.myCstr(obj.Qty) + Environment.NewLine
                    Next
                    clsCommon.MyMessageBoxShow(Me, strBatchunion, Me.Text)
                End If
            End If
        End If
    End Sub
    Private Sub gv1_KeyDown(sender As Object, e As KeyEventArgs) Handles gv1.KeyDown
        If e.KeyCode = Keys.F4 Then
            OpenSerialItem()
            'ElseIf e.KeyCode = Keys.F5 Then
            '    '======update by preeti gupta 16/10/2018
            '    If ApplyFEFO = True And (chkInternalTransfer.Checked = True OrElse chkJobWorkTransfer.Checked) Then
            '        OpenBatchItemIfFIFIOSettingON()
            '    ElseIf RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
            '        OpenBatchItem()
            '    Else
            '        OpenBatchItemIfFIFIOSettingON()
            '    End If
        End If
    End Sub
    ''richa agarwal work related to gopal ji
    Private Sub FndGatePassNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles FndGatePassNo._MYValidating
        Try
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cmbGPItemType.SelectedValue, "S") <> CompairStringResult.Equal Then
                Dim qry As String = " Select distinct TSPL_GATEPASS_TRANSFER_DETAIL.Document_No as Code,convert(varchar,TSPL_GATEPASS_TRANSFER_HEAD.Document_Date,103) as Date ,TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code from TSPL_GATEPASS_TRANSFER_DETAIL " &
               " left outer Join TSPL_TRANSFER_ORDER_DETAIL   on TSPL_TRANSFER_ORDER_DETAIL.GatePassNo=TSPL_GATEPASS_TRANSFER_DETAIL.Document_No  and TSPL_GATEPASS_TRANSFER_DETAIL.Item_Code =isnull(TSPL_TRANSFER_ORDER_DETAIL.Item_Code,'') " &
               " left Outer Join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No =TSPL_GATEPASS_TRANSFER_DETAIL.Document_No  "
                If isButtonClicked Then
                    FndGatePassNo.Value = clsCommon.ShowSelectForm("GatePass_No", qry, "Code", " isnull(TSPL_TRANSFER_ORDER_DETAIL.GatePassNo,'')='' ", FndGatePassNo.Value, "", isButtonClicked)
                    LoadGatePassTransferDetail(FndGatePassNo.Value)
                End If
            Else
                Throw New Exception("Please Select Item Type")
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadGatePassTransferDetail(ByVal GPNo As String)
        Dim qry As String = Nothing
        isInsideLoadData = True
        LoadBlankGrid()
        'Dim qry As String = "Select TSPL_ITEM_MASTER.Is_Serial_Item,TSPL_ITEM_MASTER.Is_MRP,TSPL_VEHICLE_MASTER.number,TSPL_GATEPASS_TRANSFER_DETAIL.Document_No ,TSPL_GATEPASS_TRANSFER_DETAIL.Item_Code ,TSPL_ITEM_MASTER.Item_Desc ,TSPL_GATEPASS_TRANSFER_DETAIL.Unit_code ,TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code ,TSPL_VEHICLE_MASTER.Description as VehicleName,TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code,TSPL_LOCATION_MASTER.Location_Desc,TSPL_GATEPASS_TRANSFER_DETAIL.Qty,TSPL_LOCATION_MASTER.GIT_Location ,GITLOcation.Location_Desc  from TSPL_GATEPASS_TRANSFER_DETAIL " & _
        '" left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  " & _
        '" left outer join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No =TSPL_GATEPASS_TRANSFER_DETAIL.Document_No " & _
        '" left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code  " & _
        '" left outer join TSPL_LOCATION_MASTER as GITLOcation on GITLOcation.Location_Code = TSPL_LOCATION_MASTER.GIT_Location " & _
        '" left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER .Vehicle_Id =TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code " & _
        '" where TSPL_GATEPASS_TRANSFER_DETAIL.Document_No='" & clsCommon.myCstr(FndGatePassNo.Value) & "' AND TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  not in (Select Item_Code from TSPL_TRANSFER_ORDER_DETAIL where isnull(GatePassNo ,'')='" & clsCommon.myCstr(FndGatePassNo.Value) & "')"
        ' Dim qry As String = "Select TSPL_GATEPASS_TRANSFER_HEAD.no_of_crate,TSPL_GATEPASS_TRANSFER_HEAD.no_of_jaali,TSPL_GATEPASS_TRANSFER_HEAD.no_of_box,TSPL_ITEM_MASTER.Is_Serial_Item,TSPL_ITEM_MASTER.Is_MRP,TSPL_VEHICLE_MASTER.number,TSPL_GATEPASS_TRANSFER_DETAIL.Document_No ,TSPL_GATEPASS_TRANSFER_DETAIL.Item_Code ,TSPL_ITEM_MASTER.Item_Desc ,TSPL_GATEPASS_TRANSFER_DETAIL.Unit_code ,TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code ,TSPL_VEHICLE_MASTER.Description as VehicleName,TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code,TSPL_LOCATION_MASTER.Location_Desc,TSPL_GATEPASS_TRANSFER_DETAIL.Qty  from TSPL_GATEPASS_TRANSFER_DETAIL " & _
        '" left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  " & _
        '" left outer join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No =TSPL_GATEPASS_TRANSFER_DETAIL.Document_No " & _
        '" left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code  " & _
        '" left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER .Vehicle_Id =TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code " & _
        '" where TSPL_GATEPASS_TRANSFER_DETAIL.Document_No='" & clsCommon.myCstr(FndGatePassNo.Value) & "' AND TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  not in (Select Item_Code from TSPL_TRANSFER_ORDER_DETAIL where isnull(GatePassNo ,'')='" & clsCommon.myCstr(FndGatePassNo.Value) & "')"
        StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
        If StrCrateTransferFromBooking = "1" Then
            qry = "Select TSPL_GATEPASS_TRANSFER_HEAD.no_of_crate,TSPL_GATEPASS_TRANSFER_HEAD.no_of_jaali,TSPL_GATEPASS_TRANSFER_HEAD.no_of_box,TSPL_ITEM_MASTER.Is_Serial_Item,TSPL_ITEM_MASTER.Is_MRP,TSPL_VEHICLE_MASTER.number,TSPL_GATEPASS_TRANSFER_DETAIL.Document_No ,TSPL_GATEPASS_TRANSFER_DETAIL.Item_Code ,TSPL_ITEM_MASTER.Item_Desc ,TSPL_GATEPASS_TRANSFER_DETAIL.Unit_code ,TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code ,TSPL_VEHICLE_MASTER.Description as VehicleName,TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code,TSPL_LOCATION_MASTER.Location_Desc,TSPL_GATEPASS_TRANSFER_DETAIL.Qty  from TSPL_GATEPASS_TRANSFER_DETAIL " &
      " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  " &
      " left outer join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No =TSPL_GATEPASS_TRANSFER_DETAIL.Document_No " &
      " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code  " &
      " left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER .Vehicle_Id =TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code " &
      " where TSPL_GATEPASS_TRANSFER_DETAIL.Document_No='" & clsCommon.myCstr(FndGatePassNo.Value) & "' AND TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  not in (Select Item_Code from TSPL_TRANSFER_ORDER_DETAIL where isnull(GatePassNo ,'')='" & clsCommon.myCstr(FndGatePassNo.Value) & "')"
        Else
            qry = "Select TSPL_GATEPASS_TRANSFER_HEAD.no_of_crate,TSPL_GATEPASS_TRANSFER_HEAD.no_of_jaali,TSPL_GATEPASS_TRANSFER_HEAD.no_of_box,TSPL_ITEM_MASTER.Is_Serial_Item,TSPL_ITEM_MASTER.Is_MRP,TSPL_VEHICLE_MASTER.number,TSPL_GATEPASS_TRANSFER_DETAIL.Document_No ,TSPL_GATEPASS_TRANSFER_DETAIL.Item_Code ,TSPL_ITEM_MASTER.Item_Desc ,TSPL_GATEPASS_TRANSFER_DETAIL.Unit_code ,TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code ,TSPL_VEHICLE_MASTER.Description as VehicleName,TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code,TSPL_LOCATION_MASTER.Location_Desc,TSPL_GATEPASS_TRANSFER_DETAIL.Qty  from TSPL_GATEPASS_TRANSFER_DETAIL " &
      " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  " &
      " left outer join TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.Document_No =TSPL_GATEPASS_TRANSFER_DETAIL.Document_No " &
      " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_GATEPASS_TRANSFER_HEAD.Branch_Code  " &
      " left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER .Vehicle_Id =TSPL_GATEPASS_TRANSFER_HEAD.Vehicle_Code " &
      " where TSPL_GATEPASS_TRANSFER_DETAIL.Document_No='" & clsCommon.myCstr(FndGatePassNo.Value) & "' AND TSPL_GATEPASS_TRANSFER_DETAIL .Item_Code  not in (Select Item_Code from TSPL_TRANSFER_ORDER_DETAIL where isnull(GatePassNo ,'')='" & clsCommon.myCstr(FndGatePassNo.Value) & "')"
        End If
        'If clsCommon.CompairString(cmbGPItemType.SelectedValue, "T") = CompairStringResult.Equal Then
        '    qry += " and TSPL_ITEM_MASTER.IsTaxable =1 and TSPL_ITEM_MASTER.Is_Tax_Exempted <>2"
        'ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal Then
        '    qry += " and TSPL_ITEM_MASTER.IsTaxable =0 and TSPL_ITEM_MASTER.Is_Tax_Exempted <>2"
        'ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "E") = CompairStringResult.Equal Then
        '    qry += " and (TSPL_ITEM_MASTER.IsTaxable=0 or TSPL_ITEM_MASTER.IsTaxable=1) and TSPL_ITEM_MASTER.Is_Tax_Exempted =2"
        'End If
        If clsCommon.CompairString(cmbGPItemType.SelectedValue, "T") = CompairStringResult.Equal Then
            qry += " and TSPL_ITEM_MASTER.IsTaxable =1 "
        ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal Then
            qry += " and TSPL_ITEM_MASTER.IsTaxable =0 and TSPL_ITEM_MASTER.Product_Type in ('MP') "
        ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "NO") = CompairStringResult.Equal Then
            qry += " and TSPL_ITEM_MASTER.IsTaxable =0 and TSPL_ITEM_MASTER.Product_Type not in ('MP') "
        ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "E") = CompairStringResult.Equal Then
            qry += " and TSPL_ITEM_MASTER.Is_Tax_Exempted =2"
        End If
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            txtToLoc.Value = clsCommon.myCstr(dt.Rows(0)("Branch_Code"))
            lblToLoc.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
            txtToLoc.Enabled = False
            FillLocationInfo(False)
            txtVehicleCode.Value = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
            lblVehicleNo.Text = clsCommon.myCstr(dt.Rows(0)("VehicleName"))
            txtvehicle_mannual_no.Text = clsCommon.myCstr(dt.Rows(0)("number"))
            If clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                txtCrateIn.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_crate"))
                txtBoxIn.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_box"))
                txtJaaliIn.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_jaali"))
            ElseIf clsCommon.CompairString(cmbGPItemType.SelectedValue, "NT") = CompairStringResult.Equal AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal Then
                txtcrateout.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_crate"))
                txtBoxOut.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_box"))
                txtjaaliout.Text = clsCommon.myCdbl(dt.Rows(0)("no_of_jaali"))
            End If
            For i As Integer = 0 To dt.Rows.Count - 1
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                gv1.Rows(gv1.Rows.Count - 1).Cells(colgatePassTransferNo).Value = clsCommon.myCstr(dt.Rows(i)("Document_No"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt.Rows(i)("Item_Code"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dt.Rows(i)("Item_Desc"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dt.Rows(i)("Item_Code")), Nothing)
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt.Rows(i)("Unit_code"))
                ' gv1.Rows(gv1.Rows.Count - 1).Cells(colAltUOM).Value = obj.Alt_Unit_Code
                ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = clsCommon.myCdbl(dt.Rows(i)("Qty"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).ReadOnly = True
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colIUnitWt).Value = obj.Item_Unit_Wt
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colINetWt).Value = obj.Item_Net_Wt
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colINetMTWt).Value = obj.Item_Net_MT_Wt
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = dt.Rows(i)("Is_MRP")
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = dt.Rows(i)("Is_Serial_Item")
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(clsCommon.myCstr(dt.Rows(i)("Item_Code")))
                StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                'If StrCrateTransferFromBooking = "1" Then
                '    gv1.Rows(gv1.Rows.Count - 1).Cells(colFOCItem).Value = dt.Rows(i)("Foc_Item")
                'End If
                gv1.Columns(colInQty).IsVisible = True
                gv1.Columns(colLeakQty).IsVisible = True
                gv1.Columns(colBreakQty).IsVisible = True
                gv1.Columns(colShortQty).IsVisible = True
                SetitemWiseTaxSetting(True, True)
                UpdateCurrentRow(gv1.Rows.Count - 1)
            Next
            UpdateAllTotals()
        Else
            txtToLoc.Enabled = True
            clsCommon.MyMessageBoxShow(Me, "There is no Item of " & cmbGPItemType.Text & " type in selected Gate Pass No. Please select another Item type")
        End If
        isInsideLoadData = False
    End Sub
    Sub LoadGPItemType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr = dt.NewRow()
        dr("Code") = "S"
        dr("Name") = "Select"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "T"
        dr("Name") = "Taxable"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "NT"
        dr("Name") = "Non Taxable Milk"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "NO"
        dr("Name") = "Non Taxable Other"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "E"
        dr("Name") = "Excise"
        dt.Rows.Add(dr)
        cmbGPItemType.DataSource = dt
        cmbGPItemType.ValueMember = "Code"
        cmbGPItemType.DisplayMember = "Name"
        cmbGPItemType.SelectedValue = "S"
    End Sub
    Private Sub fndPriceCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndPriceCode._MYValidating
        Try
            'If clsCommon.myLen(FndGatePassNo.Value) > 0 AndAlso clsCommon.myLen(txtToLoc.Value) > 0 Then
            Dim qry As String = " Select DISTINCT TSPL_ITEM_PRICE_MASTER.Price_Code,TSPL_ITEM_PRICE_MASTER.Price_Code_Desc from TSPL_ITEM_PRICE_MASTER  "
            If isButtonClicked Then
                'fndPriceCode.Value = clsCommon.ShowSelectForm("GatePassPriceCode", qry, "Price_Code", " isnull(TSPL_ITEM_PRICE_MASTER.Type  ,'')='T' ", fndPriceCode.Value, " ", isButtonClicked)
                fndPriceCode.Value = clsCommon.ShowSelectForm("GatePassPriceCode", qry, "Price_Code", Nothing, fndPriceCode.Value, " ", isButtonClicked)
                ''LoadRateForGatePassTransfer()
            End If
            'Else
            '    If clsCommon.myLen(txtToLoc.Value) = 0 Then
            '        Throw New Exception("Please Select To Location")
            '    Else
            '        Throw New Exception("Please Select GP No.")
            '    End If
            'End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadRateForGatePassTransfer()
        If gv1.Rows.Count > 0 Then
            Dim qry As String = String.Empty
            Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Branch_Code from TSPL_GATEPASS_TRANSFER_HEAD  where Document_No ='" & clsCommon.myCstr(FndGatePassNo.Value) & "' "))
            For i As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value)) > 0 Then
                    qry = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from ( " &
                           " Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                           " Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date, " &
                           " Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                           " TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                           " TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code where  Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  and   " &
                           " TSPL_ITEM_PRICE_MASTER.Price_Code='" & clsCommon.myCstr(fndPriceCode.Value) & "' and UOM='" & clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(strLocation) & "'  " &
                           " ) XXXE WHERE RowNo=1 "
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt.Rows.Count > 0 Then
                        gv1.Rows(i).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                    Else
                        StrCrateTransferFromBooking = clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing)
                        Dim strIsschemeItem = clsDBFuncationality.getSingleValue("select top 1 TSPL_GATEPASS_TRANSFER_HEAD.Booking_Code from TSPL_GATEPASS_TRANSFER_DETAIL left outer join " &
                                   "TSPL_GATEPASS_TRANSFER_HEAD on TSPL_GATEPASS_TRANSFER_HEAD.document_no=TSPL_GATEPASS_TRANSFER_DETAIL.document_no left outer join " &
                                   "TSPL_BOOKING_DETAIL on TSPL_GATEPASS_TRANSFER_HEAD.Booking_Code=TSPL_BOOKING_DETAIL.Document_No where  " &
                                   "TSPL_GATEPASS_TRANSFER_HEAD.Document_No='" & clsCommon.myCstr(gv1.Rows(i).Cells(colgatePassTransferNo).Value) & "' and " &
                                   "TSPL_GATEPASS_TRANSFER_DETAIL.item_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value) & "' and Scheme_Item='Y'")
                        If clsCommon.myLen(strIsschemeItem) = 0 Then
                            If StrCrateTransferFromBooking = "1" Then
                                clsCommon.MyMessageBoxShow("Please create Price chart for Location " & txtToLoc.Value & "  for item " & gv1.Rows(i).Cells(colICode).Value & ". ", Me.Text)
                                gv1.CurrentCell.Focus()
                                fndPriceCode.Value = ""
                            Else
                                clsCommon.MyMessageBoxShow("Please create Price chart for Location " & txtToLoc.Value & "  for item " & gv1.Rows(i).Cells(colICode).Value & ". ", Me.Text)
                                gv1.CurrentCell.Focus()
                                fndPriceCode.Value = ""
                            End If
                        End If
                        Exit Sub
                    End If
                End If
            Next
        End If
    End Sub
    ''---------------------------
    Private Sub cmbGPItemType_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbGPItemType.SelectedValueChanged
        If Not IsFormLoad Then
            If clsCommon.myLen(FndGatePassNo.Value) > 0 AndAlso clsCommon.CompairString(cmbGPItemType.SelectedValue, "S") <> CompairStringResult.Equal Then
                LoadGatePassTransferDetail(FndGatePassNo.Value)
            End If
        End If
    End Sub
    Sub OpenBatchItem()
        Dim isNewDocumentorExistingdoc As Boolean = True
        If clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
            If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsBatchItem).Value) Then
                If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                    Dim frm As frmBatchItemOut = New frmBatchItemOut()
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    frm.strLocationCode = txtFromLocation.Value
                    frm.strCurrDocNo = txtDocNo.Value
                    If chkInternalTransfer.Checked Then
                        frm.strCurrDocType = "ITransfer"
                    ElseIf chkJobWorkTransfer.Checked Then
                        frm.strCurrDocType = "JTransfer"
                    Else
                        frm.strCurrDocType = "Transfer"
                    End If
                    frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
                    If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                        frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colInQty).Value)
                    Else
                        frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colOutQty).Value)
                    End If
                    frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    If RunBatchFifowisewithmodifyfunctionality Then
                        If clsCommon.myLen(txtDocNo.Value) > 0 Then
                            isNewDocumentorExistingdoc = False
                        Else
                            isNewDocumentorExistingdoc = True
                        End If
                        If isNewDocumentorExistingdoc Then
                            frm.OpenSerialList(0, "")
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        Else
                            frm.ShowDialog()
                            If Not frm.isCencelButtonClicked Then
                                gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                            End If
                        End If
                    ElseIf RunBatchFifowise = 0 Then
                        frm.ShowDialog()
                        If Not frm.isCencelButtonClicked Then
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        End If
                    Else
                        If ApplyFEFO = True AndAlso (chkInternalTransfer.Checked = True OrElse chkJobWorkTransfer.Checked) Then
                            frm.OpenSerialList(0, "", "", True)
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        Else
                            frm.OpenSerialList(0, "")
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        End If
                    End If
                    'frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    'frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    'frm.strLocationCode = txtFromLocation.Value
                    'frm.strCurrDocNo = txtDocNo.Value
                    'frm.strCurrDocType = If(chkInternalTransfer.Checked = True, "ITransfer", "Transfer")
                    'frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
                    'If clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal Then
                    '    frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colInQty).Value)
                    'Else
                    '    frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colOutQty).Value)
                    'End If
                    'frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    'If ApplyFEFO = True AndAlso chkInternalTransfer.Checked = True Then
                    '    frm.OpenSerialList(0, "", "", True)
                    '    gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                    'ElseIf RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                    '    frm.ShowDialog()
                    '    If Not frm.isCencelButtonClicked Then
                    '        gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                    '    End If
                    'Else
                    '    frm.OpenSerialList(0, "")
                    '    gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                    'End If
                ElseIf clsCommon.CompairString(cboTransferType.SelectedValue, "I") = CompairStringResult.Equal Then
                    Dim frm As frmBatchItemOut = New frmBatchItemOut()
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    frm.strLocationCode = txtToLoc.Value
                    frm.strAgaintsDocNo = txtTransferOutNo.Value
                    frm.strCurrDocNo = txtDocNo.Value
                    If chkInternalTransfer.Checked Then
                        frm.strCurrDocType = "ITransfer"
                    ElseIf chkJobWorkTransfer.Checked Then
                        frm.strCurrDocType = "JTransfer"
                    Else
                        frm.strCurrDocType = "Transfer"
                    End If
                    frm.strSplTransaction = "TransferIN"
                    frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
                    frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colInQty).Value)
                    frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    frm.ShowDialog()
                    If Not frm.isCencelButtonClicked Then
                        gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub btnUpdateE_Click(sender As Object, e As EventArgs) Handles btnUpdateE.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                Dim obj As New clsTransferDCC
                obj.Document_No = clsCommon.myCstr(txtDocNo.Value)
                obj.EWayBillNo = txtEWayBillNo.Text
                obj.Electronic_Ref_No = txtElecttefNo.Text
                If txtEWayBillDate.Checked Then
                    obj.EWayBillDate = txtEWayBillDate.Value
                End If
                If clsTransferDCC.UpdateAfterPosting(obj, txtDocNo.Value, Nothing) Then
                    clsCommon.MyMessageBoxShow(Me, "Information updated successfully.")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDate_Validating(sender As Object, e As ComponentModel.CancelEventArgs) Handles txtDate.Validating
        Try
            GSTOn = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, Nothing))
            If GSTOn = 1 AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal And btnSave.Text = "Save" Then
                txtFromLocation.Value = ""
                lblFromLoc.Text = ""
                If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                    cboType.SelectedValue = "Depot"
                    cboType.Enabled = False
                Else
                    cboType.SelectedValue = ""
                    cboType.Enabled = True
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDate_ValueChanging(sender As Object, e As ValueChangingEventArgs) Handles txtDate.ValueChanging
        'Try
        '    GSTOn = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GSTApplicable, clsFixedParameterCode.GSTApplicable, Nothing))
        '    If GSTOn = 1 Then
        '        txtFromLocation.Value = ""
        '        lblFromLoc.Text = ""
        '        If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
        '            cboType.SelectedValue = "Depot"
        '            cboType.Enabled = False
        '        Else
        '            cboType.SelectedValue = ""
        '            cboType.Enabled = True
        '        End If
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'End Try
    End Sub
    '========================Added by preeti gupta[For GST Print]========================================
    'Public Sub ChkPrintGSTApplicableDate()
    '    If clsERPFuncationality.GetGSTStatus() Then
    '    End If
    'End Sub
    Private Sub chkForRepair_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkForRepair.ToggleStateChanged
        Try
            If chkForRepair.Checked = True Then
                txtTaxGroup.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Group_Code from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='EXEMPTED' and Tax_Group_Type='T'"))
                lblTaxGrpName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Group_Desc from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='EXEMPTED' and Tax_Group_Type='T'"))
                LoadBlankGridTax()
                'SetTaxDetails()
                SetitemWiseTaxSetting(True, False)
                txtTaxGroup.Enabled = False
            Else
                txtTaxGroup.Enabled = True
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub chkInternalTransfer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkInternalTransfer.ToggleStateChanged
        Try
            txtFromLocation.Value = ""
            txtToLoc.Value = ""
            lblFromLoc.Text = ""
            lblToLoc.Text = ""
            If clsCommon.myLen(fndSRNO.Value) > 0 Then
                AddNew()
            End If
            If chkInternalTransfer.Checked = True Then
                chkProductionRequest.Enabled = True
            Else
                chkProductionRequest.Enabled = False
                chkProductionRequest.Checked = False
            End If

            If chkInternalTransfer.Checked Then
                If chkJobWorkTransfer.Checked Then
                    chkInternalTransfer.Checked = False
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ''RICHA UDL/14/05/18-000161
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        CancelData()
    End Sub
    Sub CancelData()
        Try
            If (myMessages.cancelConfirm()) Then
                If (clsTransferDCC.CancelData(MyBase.Form_ID, txtDocNo.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled")
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ' Ticket : TEC/29/10/18-000353 By Sanjay
    Private Sub btnShowInventory_Click(sender As Object, e As EventArgs) Handles btnShowInventory.Click
        clsOpenInventory.ShowInventoryDatails(txtDocNo.Value)
    End Sub
    'Ticket No-MIL/14/08/19-000120,Sanjay,Transfer Out- Item Import/Export
    Private Sub radExportTransferOut_Click(sender As Object, e As EventArgs) Handles radExportTransferOut.Click
        If clsCommon.CompairString(txtDocNo.Value, "") = CompairStringResult.Equal Then
            common.clsCommon.MyMessageBoxShow(Me, "Please select document.")
            Exit Sub
        End If
        Dim str As String
        Dim whrcls As String = ""
        str = "SELECT TSPL_TRANSFER_ORDER_DETAIL.Item_Code as [Item Code] ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as [Item Unit]  "
        str += ",TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as [Item Qty]"
        str += " FROM TSPL_TRANSFER_ORDER_DETAIL "
        whrcls = " and document_no='" + txtDocNo.Value + "'"
        transportSql.ExporttoExcel(str, whrcls, Me)
    End Sub
    Private Sub radImportTransferOut_Click(sender As Object, e As EventArgs) Handles radImportTransferOut.Click
        If clsCommon.CompairString(txtFromLocation.Value, "") = CompairStringResult.Equal Then
            common.clsCommon.MyMessageBoxShow(Me, "Please select location.")
            Exit Sub
        End If
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        If transportSql.importExcel(gv, "Item Code", "Item Unit", "Item Qty") Then
            Dim dt As New DataTable()
            dt = gv.DataSource()
            gv1.Rows.Clear()
            For Each row As DataRow In dt.Rows
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(row("Item Code").ToString().Trim())
                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(row("Item Unit").ToString().Trim())
                gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = clsCommon.myCdbl(row("Item Qty").ToString().Trim())
                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = clsCommon.myCdbl(gv1.Rows.Count)
            Next
            gv1.Rows.AddNew()
        End If
    End Sub
    Private Sub fndSRNO__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndSRNO._MYValidating
        Try
            If clsCommon.myLen(txtFromLocation.Value) <= 0 Then
                Throw New Exception("Please first select from location")
            End If
            'If chkInternalTransfer.Checked = True Then
            Dim qry As String = " Select distinct z.Document_No as Code, z.[Date],z.Description,z.[Location Code],z.Location_Desc as [Location Name],z.[Dept Code],z.[Dept Desc],z.Mode_Of_Transport as [Mode Of Transport],z.Remarks,z.Comments,z.[All/Transfer] from ( " &
        " select code as Document_No,SUM(Qty* RI) as OutQty,Max(Unit_Code) as UnitCode ,max(Requisition_Date) as [Date], max(Description) as [Description], max(Location) as [Location Code], max(Location_Desc) as Location_Desc, max(Dept) as [Dept Code],max(Dept_Desc) as [Dept Desc], max(Remarks) as Remarks,max(Mode_Of_Transport) as Mode_Of_Transport, max(Comments) as Comments,Case When Max(All_Transfer_Issue)=0 Then 'All' Else (Case When Max(All_Transfer_Issue)=1 Then 'Transfer' Else '' End) End As [All/Transfer] " &
        "  from " &
         " ( Select TSPL_REQUISITION_HEAD.Requisition_Id as Code,TSPL_REQUISITION_DETAIL.Item_Code as ICode,TSPL_REQUISITION_DETAIL.Requisition_Qty as Qty " &
         " ,1 as RI,TSPL_REQUISITION_DETAIL.Unit_Code, Convert (varchar,TSPL_REQUISITION_HEAD.Requisition_Date,103) as Requisition_Date,TSPL_REQUISITION_HEAD.Description,TSPL_REQUISITION_HEAD.Location,TSPL_LOCATION_MASTER.Location_Desc,TSPL_REQUISITION_HEAD.Dept,TSPL_REQUISITION_HEAD.Dept_Desc,TSPL_REQUISITION_HEAD.Remarks,TSPL_REQUISITION_HEAD.Mode_Of_Transport,TSPL_REQUISITION_HEAD.Comments,TSPL_REQUISITION_HEAD.All_Transfer_Issue  FROM TSPL_REQUISITION_DETAIL  LEFT OUTER JOIN TSPL_REQUISITION_HEAD  " &
         " ON TSPL_REQUISITION_DETAIL.Requisition_Id = TSPL_REQUISITION_HEAD.Requisition_Id Left Outer Join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = TSPL_REQUISITION_HEAD.Location  where   TSPL_REQUISITION_HEAD.Status =1  " &
         " Union All " &
         " Select TSPL_TRANSFER_ORDER_HEAD.Requisition_Id  as Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty " &
         " ,-1 as RI,TSPL_TRANSFER_ORDER_DETAIL.Unit_Code,Convert (varchar,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) as Requisition_Date , TSPL_TRANSFER_ORDER_HEAD.Description ,TSPL_TRANSFER_ORDER_HEAD.From_Location as Location,TSPL_LOCATION_MASTER.Location_Desc ,'' as Dept,'' as Dept_Desc,TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_TRANSFER_ORDER_HEAD.Mode_Of_Transport,TSPL_TRANSFER_ORDER_HEAD.Comments,0 As  All_Transfer_Issue   FROM TSPL_TRANSFER_ORDER_DETAIL  LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD  " &
         " ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No  Left Outer Join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = TSPL_TRANSFER_ORDER_HEAD.From_Location where Requisition_Id is not null " &
         " ) Final where Final.code not in (select isnull(TSPL_TRANSFER_ORDER_HEAD.Requisition_Id,'') as Requisition_Id from TSPL_TRANSFER_ORDER_HEAD ) " &
         " group by Code,ICode,Unit_Code having SUM(Qty *RI) >0   " &
         " )z  "
            Dim whrcls = " [All/Transfer] in ('All','Transfer')"
            'InternalTransfer=1 and
            If isButtonClicked Then
                fndSRNO.Value = clsCommon.ShowSelectForm("SRNO1@1", qry, "Code", whrcls, fndSRNO.Value, "Code", isButtonClicked)
                LoadStoreRequisitionDetail(fndSRNO.Value)
            End If
            'End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadStoreRequisitionDetail(ByVal SRNO As String)
        Dim qry As String = Nothing
        isInsideLoadData = True
        LoadBlankGrid()
        qry = " Select z.Location,z.Document_No,z.OutQty as Qty,z.ICode as Item_Code,TSPL_ITEM_MASTER.Item_Desc,z.unitcode as Unit_Code,TSPL_ITEM_MASTER.Is_Serial_Item,TSPL_ITEM_MASTER.Is_MRP, z.Item_Cost  from ( " &
             " select max(Location) as Location, code as Document_No,SUM(Qty* RI) as OutQty,ICode,Max(Unit_Code) as UnitCode,Max(Item_Cost) as Item_Cost " &
             " from( Select  TSPL_REQUISITION_HEAD.Location, TSPL_REQUISITION_HEAD.Requisition_Id as Code,TSPL_REQUISITION_DETAIL.Item_Code as ICode,TSPL_REQUISITION_DETAIL.Requisition_Qty as Qty " &
            " ,1 as RI,TSPL_REQUISITION_DETAIL.Unit_Code,TSPL_REQUISITION_DETAIL.Item_Cost FROM TSPL_REQUISITION_DETAIL  LEFT OUTER JOIN TSPL_REQUISITION_HEAD  " &
            " ON TSPL_REQUISITION_DETAIL.Requisition_Id = TSPL_REQUISITION_HEAD.Requisition_Id  where   TSPL_REQUISITION_HEAD.Status =1  " &
            " Union All  " &
            " Select  null as Location, TSPL_TRANSFER_ORDER_HEAD.Requisition_Id  as Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Code as ICode,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty " &
            " ,-1 as RI,TSPL_TRANSFER_ORDER_DETAIL.Unit_Code,TSPL_TRANSFER_ORDER_DETAIL.Item_Cost FROM TSPL_TRANSFER_ORDER_DETAIL  LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD  " &
            " ON TSPL_TRANSFER_ORDER_DETAIL.Document_No = TSPL_TRANSFER_ORDER_HEAD.Document_No  where Requisition_Id is not null " &
            " ) Final  " &
            " group by Code,ICode,Unit_Code having SUM(Qty *RI) >0  )z  " &
            " left outer join tspl_item_master on tspl_item_master.Item_Code =z.ICode " &
            " where z.Document_No='" + SRNO + "'"
        'InternalTransfer=1 and ,Item_Cost
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            txtToLoc.Value = clsCommon.myCstr(dt.Rows(0)("Location"))
            'lblToLoc.Text = clsLocation.GetName(clsCommon.myCstr(dt.Rows(0)("Location")), Nothing)
            If chkInternalTransfer.Checked = True Then
                txtToLoc.Enabled = True
            Else
                FillLocationInfo(False)
            End If
            lblToLoc.Text = clsLocation.GetName(clsCommon.myCstr(txtToLoc.Value), Nothing)
            For i As Integer = 0 To dt.Rows.Count - 1
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colgatePassTransferNo).Value = clsCommon.myCstr(dt.Rows(i)("Document_No"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt.Rows(i)("Item_Code"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dt.Rows(i)("Item_Desc"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dt.Rows(i)("Item_Code")), Nothing)
                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt.Rows(i)("Unit_code"))
                gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(i)("Item_Cost"))
                If clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value) <= 0 Then
                    Dim strDate As String = txtDate.Value
                    Dim dblCostMethod As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Costing_Method as Costing_Method from TSPL_ITEM_MASTER left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code where Item_Code='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "' "))
                    If dblCostMethod <> 0 Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsInventoryMovement.GetCost(dblCostMethod, clsCommon.myCstr(dt.Rows(i)("Item_Code")), txtFromLocation.Value, 1, strDate, strDate, False, Nothing)
                    End If
                End If
                gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = clsCommon.myCdbl(dt.Rows(i)("Qty"))
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).ReadOnly = True
                gv1.Rows(gv1.Rows.Count - 1).Cells(colisMRPMandatory).Value = dt.Rows(i)("Is_MRP")
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = dt.Rows(i)("Is_Serial_Item")
                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(clsCommon.myCstr(dt.Rows(i)("Item_Code")))
                SetitemWiseTaxSetting(True, True)
                UpdateCurrentRow(gv1.Rows.Count - 1)
            Next
            If AllowOnlyOneIssueAgainstStoreRequisition = True Then
                gv1.Rows.AddNew()
            End If
            UpdateAllTotals()
        Else
            txtToLoc.Enabled = True
            clsCommon.MyMessageBoxShow("There is no Item of " & cmbGPItemType.Text & " type in selected Gate Pass No. Please select another Item type")
        End If
        isInsideLoadData = False
    End Sub
    Private Sub chkJobWork_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkJobWork.ToggleStateChanged
        If chkJobWork.Checked = True Then
            chkInternalTransfer.Checked = True
            chkInternalTransfer.Enabled = False
        Else
            chkInternalTransfer.Checked = False
            chkInternalTransfer.Enabled = True
        End If
    End Sub
    Private Sub BtnHistory_Click(sender As Object, e As EventArgs) Handles BtnHistory.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Code not found to show history")
                txtDocNo.Focus()
            End If
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                clsERPFuncationalityOLD.ShowTransHistoryData(clsCommon.myCstr(txtDocNo.Value), "DOCUMENT_NO", "TSPL_TRANSFER_ORDER_HEAD", "TSPL_TRANSFER_ORDER_DETAIL")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub fndTransferIndentNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndTransferIndentNo._MYValidating
        Try
            If clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal And clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select IsParlour from tspl_location_master where Location_code='" & clsCommon.myCstr(txtToLoc.Value) & "'")), "Y") = CompairStringResult.Equal AndAlso chkProductionRequest.Checked = False Then
                Dim strQuery As String = "Select TSPL_TRANSFER_INDENT_MASTER.Document_No as IndentNo, Convert (varchar,TSPL_TRANSFER_INDENT_MASTER.Document_Date,103) as [InvoiceDate] ,  TSPL_TRANSFER_INDENT_MASTER.Location_code as [Location Code], TSPL_Location_Master.Location_Desc as [Location Name] from TSPL_TRANSFER_INDENT_MASTER " & Environment.NewLine &
" Left Outer Join TSPL_Location_Master On TSPL_Location_Master.Location_Code = TSPL_TRANSFER_INDENT_MASTER.location_Code "
                Dim whr As String = " TSPL_TRANSFER_INDENT_MASTER.Booking_By ='Parlor' and  TSPL_TRANSFER_INDENT_MASTER.location_Code ='" & clsCommon.myCstr(txtToLoc.Value) & "' and TSPL_TRANSFER_INDENT_MASTER.Document_No not in ( select TransferIndent_No from tspl_transfer_order_head where isnull(TransferIndent_No,'')<>'' and tspl_transfer_order_head.document_no <>'" & clsCommon.myCstr(txtDocNo.Value) & "' )  " & Environment.NewLine &
" and Convert(date,TSPL_TRANSFER_INDENT_MASTER.Document_Date,103) <= Convert(date,'" + txtDate.Value + "',103)"
                'strQuery = "Select * from ( " + strQuery + whr + strProductionStoreRequest + " ) XXX"
                fndTransferIndentNo.Value = clsCommon.ShowSelectForm("IndentNo@Transfer", strQuery, "IndentNo", whr, fndTransferIndentNo.Value, "IndentNo", isButtonClicked)
                If clsCommon.myLen(fndTransferIndentNo.Value) > 0 Then
                    fillTransferIndentNo(fndTransferIndentNo.Value)
                Else
                    LoadBlankGrid()
                End If
            ElseIf chkProductionRequest.Checked = True AndAlso clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal AndAlso (chkInternalTransfer.Checked = True OrElse chkJobWorkTransfer.Checked) Then
                Dim strQuery As String = " Select TSPL_PP_REQUISITION_HEAD.Requisition_Id  as IndentNo, convert (varchar, TSPL_PP_REQUISITION_HEAD.Requisition_Date,103) as [InvoiceDate] ,TSPL_PP_REQUISITION_HEAD.Location  as [Location Code] , TSPL_Location_Master.Location_Desc  as [Location Name] from TSPL_PP_REQUISITION_HEAD 
                                           Left Outer Join TSPL_Location_Master on (TSPL_Location_Master.Location_Code = TSPL_PP_REQUISITION_HEAD.Location or TSPL_Location_Master.Main_Location_Code = TSPL_PP_REQUISITION_HEAD.Location ) "
                Dim whr As String = " TSPL_PP_REQUISITION_HEAD.close_yn = 'N' and TSPL_PP_REQUISITION_HEAD.Status = 1 and  TSPL_Location_Master.Location_Code  ='" & clsCommon.myCstr(txtToLoc.Value) & "'  and TSPL_PP_REQUISITION_HEAD.Requisition_Id not in ( select TransferIndent_No from tspl_transfer_order_head where isnull(TransferIndent_No,'')<>'' ) "
                fndTransferIndentNo.Value = clsCommon.ShowSelectForm("IndentNo@Transfer@PSR", strQuery, "IndentNo", whr, fndTransferIndentNo.Value, "IndentNo", isButtonClicked)
                If clsCommon.myLen(fndTransferIndentNo.Value) > 0 Then
                    fillTransferIndentNo(fndTransferIndentNo.Value)
                Else
                    LoadBlankGrid()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Transfer", MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub fillTransferIndentNo(ByVal strTransferIndentNo As String)
        Try
            LoadBlankGrid()
            Dim isProductionStoreReqDoc As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue(" select count (*)  from TSPL_PP_REQUISITION_HEAD where Requisition_Id = '" + strTransferIndentNo + "' "))
            Dim strQuery As String = ""
            If isProductionStoreReqDoc = True AndAlso chkInternalTransfer.Checked = True AndAlso chkProductionRequest.Checked = True Then
                strQuery = " select TSPL_PP_REQUISITION_Detail.Item_Code , TSPL_ITEM_MASTER.Item_Desc,  TSPL_ITEM_MASTER.Short_Description, TSPL_ITEM_MASTER.HSN_Code , TSPL_PP_REQUISITION_Detail.Unit_Code , sum(TSPL_PP_REQUISITION_Detail.Requisition_Qty) as Qty from TSPL_PP_REQUISITION_Detail left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_PP_REQUISITION_Detail.Item_Code
                             where Requisition_Id = '" + strTransferIndentNo + "'
                             group by TSPL_PP_REQUISITION_Detail.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,TSPL_PP_REQUISITION_Detail.Unit_Code    "
            Else
                strQuery = " Select TSPL_TRANSFER_INDENT_DETAILS.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,TSPL_TRANSFER_INDENT_DETAILS.Unit_Code , sum(TSPL_TRANSFER_INDENT_DETAILS.Tranfer_Qty) as Qty from TSPL_TRANSFER_INDENT_DETAILS  Left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code= TSPL_TRANSFER_INDENT_DETAILS.Item_Code 
                             where TSPL_TRANSFER_INDENT_DETAILS.Document_No = '" + strTransferIndentNo + "' group by TSPL_TRANSFER_INDENT_DETAILS.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,TSPL_TRANSFER_INDENT_DETAILS.Unit_Code "
            End If
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(strQuery)
            Dim intLineNo As Integer = 0
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                For Each dr As DataRow In dt.Rows
                    gv1.Rows.AddNew()
                    intLineNo += 1
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = intLineNo
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("Item_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOutQty).Value = clsCommon.myCstr(dr("Qty"))
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub gv1_UserAddingRow(sender As Object, e As GridViewRowCancelEventArgs) Handles gv1.UserAddingRow
    End Sub
    Private Sub btn_CancelDel_Click(sender As Object, e As EventArgs) Handles btn_CancelDel.Click
        CancelDelData()
    End Sub
    Function CancelDelData() As Boolean
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Code is empty")
            End If
            If clsCommon.MyMessageBoxShow("Are you sure to Cancel the Record?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                Return False
            End If
            Dim strTransfer_In_Ret_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_No from TSPL_TRANSFER_ORDER_HEAD where TransferOutNo='" & txtDocNo.Value & "' "))
            If clsCommon.myLen(strTransfer_In_Ret_No) > 0 Then
                Dim StrTransType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Transfer_Type from TSPL_TRANSFER_ORDER_HEAD where Document_No='" & strTransfer_In_Ret_No & "' "))
                If clsCommon.CompairString(StrTransType, "I") = CompairStringResult.Equal Then
                    Throw New Exception("You cannot cancelled this document because its Transfer In (" + clsCommon.myCstr(strTransfer_In_Ret_No) + ") has been created.")
                ElseIf clsCommon.CompairString(StrTransType, "T") = CompairStringResult.Equal Then 'Used UDL,Transfer Type - Retrun 
                    Throw New Exception("You cannot cancelled this document because its Transfer Return (" + clsCommon.myCstr(strTransfer_In_Ret_No) + ") has been created.")
                End If
            End If
            Dim strTransferReturnNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_No from TSPL_TRANSFER_RETURN where Transfer_No='" & txtDocNo.Value & "' "))
            If clsCommon.myLen(strTransferReturnNo) > 0 Then
                Throw New Exception("You cannot cancelled this document because its Transfer Return (" + clsCommon.myCstr(strTransferReturnNo) + ") has been created.")
            End If
            If chkJobWork.Checked = False AndAlso (clsCommon.CompairString(cboTransferType.SelectedValue, "O") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboTransferType.SelectedValue, "T") = CompairStringResult.Equal) AndAlso chkTaxable.Checked = True AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                Dim EInvoiceCancelTimeValid As Int64 = 0
                EInvoiceCancelTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,Posting_Date,GETDATE()),0) as PostedHours from TSPL_TRANSFER_ORDER_HEAD where  Document_No = '" + txtDocNo.Value + "'"))
                If EInvoiceCancelTimeValid >= 24 Then
                    Throw New Exception("Transfer can not be cancelled.It has been more than 24 hours.")
                End If
            End If
            clsTransferDCC.CancelDelData(Me.Form_ID, txtDocNo.Value, NavigatorType.Current)
            clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled", Me.Text)
            AddNew()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return True
    End Function
    Private Sub gv1_CommandCellClick(sender As Object, e As EventArgs) Handles gv1.CommandCellClick
    End Sub
    Private Sub btnSTAPrint_Click(sender As Object, e As EventArgs) Handles btnSTAPrint.Click
        STAPrint(Me.Form_ID, Me.Text, txtDocNo.Value, Nothing)
    End Sub

    Sub STAPrint(ByVal strFormID As String, ByVal strFormName As String, ByVal strCode As String, ByVal isCancelOrDelete As String)
        Try
            Dim strCancelOrDelete As String = Nothing
            Dim tbl_TSPL_TRANSFER_ORDER_HEAD As String = Nothing
            Dim tbl_TSPL_TRANSFER_ORDER_Detail As String = Nothing
            If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
                strCancelOrDelete = "Cancelled"
                tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Cancel_Data As TSPL_TRANSFER_ORDER_HEAD "
                tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Cancel_Data As TSPL_TRANSFER_ORDER_Detail "
            ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
                strCancelOrDelete = "Cancelled"
                tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD_Delete_Data As TSPL_TRANSFER_ORDER_HEAD "
                tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail_Delete_Data As TSPL_TRANSFER_ORDER_Detail "
            Else
                strCancelOrDelete = Nothing
                tbl_TSPL_TRANSFER_ORDER_HEAD = " TSPL_TRANSFER_ORDER_HEAD "
                tbl_TSPL_TRANSFER_ORDER_Detail = " TSPL_TRANSFER_ORDER_Detail "
            End If
            If clsCommon.myLen(strCode) > 0 Then
                Dim Qry As String = Nothing
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    Qry = "Select '" & strCancelOrDelete & "' As IsCancelled,* from (select TSPL_TRANSFER_ORDER_HEAD.Tax_Group,ISNULL(TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted,0) AS Is_Tax_Exempted ,TSPL_TRANSFER_ORDER_HEAD.Is_MandiTax, TSPL_TRANSFER_ORDER_HEAD.Electronic_Ref_No,
                       TSPL_LOCATION_MASTER.GSTNO as GSTIN_No ,TSPL_STATE_MASTER.GST_STATE_Code as From_Gst_StateCode,StateMaster_ToLocation.GST_STATE_Code as To_Loc_GSTStateCode,TSPL_LOCATION_MASTER_1.GSTNO as To_Loc_GSTINNo, 
                       TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_ITEM_MASTER.HSN_Code ,TSPL_LOCATION_MASTER.GSTNO as frm_GSTINNo ,TSPL_STATE_MASTER.GST_STATE_Code as Frm_StateGST,StateMaster_ToLocation.GST_STATE_Code as To_StateGST,
                       TSPL_LOCATION_MASTER_1.GSTNO as To_GSTINNo,TSPL_TRANSFER_ORDER_HEAD.EWayBillNo ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as EWayBillDate ,convert(varchar(10),TSPL_COMPANY_MASTER.insurance_valid_date,103) as insurance_valid_date,
                       TSPL_COMPANY_MASTER.Pan_No as ToLoc_PANNO,  TSPL_TRANSFER_ORDER_HEAD.Is_taxable,TSPL_TRANSFER_ORDER_HEAD.For_Repair,TSPL_TRANSFER_ORDER_HEAD.InternalTransfer, TSPL_TRANSFER_ORDER_DETAIL.Disc_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1 as dTAX1, TSPL_TRANSFER_ORDER_DETAIL.TAX2 as dTAX2, TSPL_TRANSFER_ORDER_DETAIL.TAX3 as  dTAX3, TSPL_TRANSFER_ORDER_DETAIL.TAX4 as  dTAX4, TSPL_TRANSFER_ORDER_DETAIL.TAX5 as  dTAX5, TSPL_TRANSFER_ORDER_DETAIL.TAX6 as  dTAX6, TSPL_TRANSFER_ORDER_DETAIL.TAX7 as  dTAX7, TSPL_TRANSFER_ORDER_DETAIL.TAX8 as dTAX8, TSPL_TRANSFER_ORDER_DETAIL.TAX9 as dTAX9, TSPL_TRANSFER_ORDER_DETAIL.TAX10 as  dTAX10, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX4_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX5_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX6_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX7_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX8_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX9_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX10_Amt, TSPL_TRANSFER_ORDER_DETAIL.TAX1_Rate as dTAX1_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX2_Rate as dTAX2_Rate, TSPL_TRANSFER_ORDER_DETAIL.TAX3_Rate as dTAX3_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX4_Rate as dTAX4_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX5_Rate as dTAX5_Rate  ,TSPL_TRANSFER_ORDER_DETAIL.TAX6_Rate as dTAX6_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX7_Rate as dTAX7_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX8_Rate as dTAX8_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX9_Rate as dTAX9_Rate ,TSPL_TRANSFER_ORDER_DETAIL.TAX10_Rate as dTAX10_Rate,  dtax1.Type as tax1Type,dtax2.Type as tax2Type,dtax3.Type as tax3Type,dtax4.Type as tax4Type,dtax5.Type as tax5Type,dtax6.Type as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type, case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.crate_In else TSPL_TRANSFER_ORDER_HEAD.crate_out end as crate,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.jaali_in else TSPL_TRANSFER_ORDER_HEAD.jaali_Out end as jaali,case when TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'I' then TSPL_TRANSFER_ORDER_HEAD.box_in else TSPL_TRANSFER_ORDER_HEAD.box_Out end as box,TSPL_TRANSFER_ORDER_DETAIL.Line_No, '1' as CopyType,0 as Alter_UnitQty ,tax1.Tax_Code_Desc as tax1name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax1_amt,0) as txt1amt,   tax2.Tax_Code_Desc as tax2name,isnull (TSPL_TRANSFER_ORDER_HEAD.tax2_amt,0) as txt2amt,   tax3.Tax_Code_Desc as tax3name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax3_amt,0) as txt3amt,   tax4.Tax_Code_Desc as tax4name,   isnull (TSPL_TRANSFER_ORDER_HEAD.tax4_amt,0) as txt4amt,   tax5.Tax_Code_Desc as tax5name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax5_amt,0) as txt5amt,   tax6.Tax_Code_Desc as tax6name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax6_amt,0) as txt6amt,   tax7.Tax_Code_Desc as tax7name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax7_amt,0) as txt7amt,   tax8.Tax_Code_Desc as tax8name, isnull (TSPL_TRANSFER_ORDER_HEAD.tax8_amt,0) as txt8amt, tax9.Tax_Code_Desc as tax9name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax9_amt,0) as txt9amt,   tax10.Tax_Code_Desc as tax10name,  isnull (TSPL_TRANSFER_ORDER_HEAD.tax10_amt,0) as txt10amt,TSPL_TRANSFER_ORDER_HEAD.TAX1_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX2_Rate ,TSPL_TRANSFER_ORDER_HEAD.TAX3_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX4_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX5_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX6_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX7_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX8_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX9_Rate,TSPL_TRANSFER_ORDER_HEAD.TAX10_Rate, From_GP_Location.add1 +case when len(From_GP_Location.add2)>0 then ', '+From_GP_Location.add2 else '' end +case when LEN(isnull(From_GP_Location.Add3,''))>0 then ', '+isnull(From_GP_Location.Add3,'') else ' ' end  + case when len(From_GP_Location_State.STATE_NAME   )>0 then ', '+ From_GP_Location_State.STATE_NAME  else ' ' end    as From_Location_Address_GP, TSPL_TRANSFER_ORDER_HEAD.Transfer_Type, From_GP_Location.Location_Code as From_Location_GP_Code,From_GP_Location.Location_Desc as From_Location_GP_Name,From_GP_Location.Add1 as From_GP_Add1,From_GP_Location.Add2 as From_GP_Add2,From_GP_Location.Add3 as From_GP_Add3,From_GP_Location.Add4 as From_GP_Add4,From_GP_Location_State.STATE_NAME as From_GP_State_Name, From_GP_Location.TIN_No as From_GP_TINNO, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Location_Code else TSPL_LOCATION_MASTER_1.Location_Code end as To_location_GP_Code, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Location_Desc else TSPL_LOCATION_MASTER_1.Location_Desc end as To_Location_GP_Name, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add1 else TSPL_LOCATION_MASTER_1.Add1 end as To_GP_Add1, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add2 else TSPL_LOCATION_MASTER_1.Add2 end as To_GP_Add2, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add3 else TSPL_LOCATION_MASTER_1.Add3 end as To_GP_Add3, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.Add4 else TSPL_LOCATION_MASTER_1.Add4 end as To_GP_Add4, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location_State.STATE_NAME else StateMaster_ToLocation.STATE_NAME end as TO_GP_State_Name, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.TIN_No else TSPL_LOCATION_MASTER_1.TIN_No end  as To_GP_TINNO , case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.TAN_No else TSPL_LOCATION_MASTER_1.TAN_No end as TO_GP_FAX, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then isnull(To_GP_Location.Pin_Code,0) else isnull(TSPL_LOCATION_MASTER_1.Pin_Code,0) end as To_GP_Loc_Pin, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then To_GP_Location.City_Code else TSPL_LOCATION_MASTER_1.City_Code end as To_GP_City_Code, case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' then Case when len(ISNULL(To_GP_Location.Phone1,''))>0 and To_GP_Location.Phone1='(+__)__________' then '' else To_GP_Location.Phone1 end + Case When   ISNULL(To_GP_Location.Phone2,'')<>'(+__)__________' Then '  '+ To_GP_Location.Phone2 Else'' end else Case when len(ISNULL(TSPL_LOCATION_MASTER_1.Phone1,''))>0 and TSPL_LOCATION_MASTER_1.Phone1='(+__)__________' then '' else TSPL_LOCATION_MASTER_1.Phone1 end + Case When   ISNULL(TSPL_LOCATION_MASTER_1.Phone2,'')<>'(+__)__________' Then '  '+ TSPL_LOCATION_MASTER_1.Phone2 Else'' end end To_GP_Phn, TSPL_LOCATION_MASTER.TIN_No as Loc_Tin_No,TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end  + case when len(TSPL_STATE_MASTER.STATE_NAME   )>0 then ', '+ TSPL_STATE_MASTER.STATE_NAME  else ' ' end    as Location_Address_GP,TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range, TSPL_TRANSFER_ORDER_HEAD.Remarks,TSPL_STATE_MASTER.state_code as frm_State_code,tspl_location_master.HOAdd1 ,TSPL_LOCATION_MASTER .HOAdd2,TSPL_TRANSFER_ORDER_HEAD.GR_No,TSPL_TRANSFER_ORDER_HEAD.document_type ,case when coalesce(TSPL_TRANSFER_ORDER_HEAD.GR_No,'')='' then null else convert(varchar,TSPL_TRANSFER_ORDER_HEAD.GR_Date,103) end as GR_Date ,TSPL_TRANSFER_ORDER_HEAD.WayBill_No ,convert(varchar,TSPL_TRANSFER_ORDER_HEAD.EWayBillDate,103) as WayBill_Date,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,tspl_location_master_For_Location.City_Code as Location_City_Name, coalesce(cast(convert(decimal(18,0),(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * tspl_item_uom_detail.conversion_factor)/alt_convrsn.conversion_factor) as varchar)+' '+TSPL_TRANSFER_ORDER_DETAIL.Alt_Unit_Code,'') as Alt_Unit_Code,TSPL_TRANSPORT_MASTER.Transporter_Name,TSPL_TRANSFER_ORDER_HEAD.transport_id,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual ,(case when TSPL_TRANSFER_ORDER_HEAD.Is_AgainstFormF=1 then 'Against F-Form Due' else '' end) as Is_AgainstFormF,TSPL_TRANSFER_ORDER_HEAD.Document_No  as[STN_NO] ,
                       tspl_transfer_order_head.Document_Date as [Date_N_Time_issue],  TSPL_TRANSFER_ORDER_HEAD.Discount_Amt  as Discount , TSPL_TRANSFER_ORDER_DETAIL .Document_No as ref_doc_no , TSPL_TRANSFER_ORDER_HEAD.From_Location, TSPL_LOCATION_MASTER.Location_Desc as From_LocationName,TSPL_LOCATION_MASTER_2.Location_Desc as To_LocationName,TSPL_TRANSFER_ORDER_HEAD.To_Location ,
                       TSPL_TRANSFER_ORDER_HEAD.Vehicle_No,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1, TSPL_LOCATION_MASTER.Location_Code as From_Location_Code,  TSPL_LOCATION_MASTER.Location_Desc as From_Location_Dec,TSPL_LOCATION_MASTER.Add1 as From_Location_Add1 , TSPL_LOCATION_MASTER.Add2 as From_Location_Add2 , TSPL_LOCATION_MASTER.Add3 as From_Location_Add3,TSPL_LOCATION_MASTER.Add4  as From_Location_Add4,TSPL_LOCATION_MASTER.City_Code as From_Location_City_Code ,  TSPL_STATE_MASTER.STATE_NAME  as From_Location_State,TSPL_LOCATION_MASTER.Pin_Code as From_Location_Pin_Code ,TSPL_LOCATION_MASTER.Country as From_Location_Country,  TSPL_LOCATION_MASTER.Telphone  as From_Location_Telphone,TSPL_LOCATION_MASTER.Email as From_Location_Email,TSPL_LOCATION_MASTER.TIN_No AS From_Location_tin_no, TSPL_LOCATION_MASTER.CST_No AS From_Location_cstNo,TSPL_ITEM_MASTER.Weight_UOM as UOM2, TSPL_ITEM_MASTER.Weight_Value as Weight,((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinLTR.Conversion_Factor ) As TotalQty, CFinLTR.UOM_Code As TotalQtyUOM,TSPL_LOCATION_MASTER_1.Location_Code AS To_Location_Code,TSPL_LOCATION_MASTER.add1 as transpotor,TSPL_LOCATION_MASTER_1.Add1 AS To_Location_Add1,  TSPL_LOCATION_MASTER_1.Add2 as To_Location_Add2 ,TSPL_LOCATION_MASTER_1.Add3 as To_Location_Add3,TSPL_LOCATION_MASTER_1.Add4 as To_Location_Add4, TSPL_LOCATION_MASTER_1.Location_Desc as To_Location_Desc  , TSPL_LOCATION_MASTER_1.City_Code as To_Location_City_Code , StateMaster_ToLocation.State_Name as To_Location_State, TSPL_LOCATION_MASTER_1.Pin_Code as To_Location_Pin_Code,  TSPL_LOCATION_MASTER_1.Country as To_Location_Country, TSPL_LOCATION_MASTER_1.Telphone as To_Location_Telphone, TSPL_LOCATION_MASTER_1.Email as To_Location_Email ,  TSPL_LOCATION_MASTER_1 .TIN_No as to_location_tin_no, TSPL_LOCATION_MASTER_1 .CST_No as to_location_cstno,TSPL_TRANSFER_ORDER_HEAD.TAX1,TSPL_TRANSFER_ORDER_HEAD.TAX2,TSPL_TRANSFER_ORDER_HEAD.TAX3,TSPL_TRANSFER_ORDER_HEAD.TAX4,TSPL_TRANSFER_ORDER_HEAD.TAX5,TSPL_TRANSFER_ORDER_HEAD.TAX6 ,TSPL_COMPANY_MASTER.Comp_Name AS CompName ,
                        TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end + Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ' ,Phone'+TSPL_COMPANY_MASTER.Phone1 end + Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then '  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end as Company_Address,'PAN No - '+TSPL_COMPANY_MASTER.Pan_No +',GSTIN - '+TSPL_COMPANY_MASTER.GSTReg_No As CompPanGST,
                        TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount ,  tspl_company_master.Pan_No,
                        tspl_company_master.State,TSPL_TRANSFER_ORDER_HEAD.Requisition_Id,TSPL_COMPANY_MASTER.GSTReg_No ,TSPL_COMPANY_MASTER.add1,TSPL_COMPANY_MASTER.add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Phone2
                        from " & tbl_TSPL_TRANSFER_ORDER_Detail & " 
                        join " & tbl_TSPL_TRANSFER_ORDER_HEAD & "  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No  
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
                        left outer join TSPL_LOCATION_MASTER AS TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.GIT_LOCATION =  TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                        INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='LTR') As CFinLTR On CFinLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code 
                        LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code 
                        LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State  
                        left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
                        left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code  
                        left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id  
                        LEFT OUTER JOIN tspl_location_master  AS tspl_location_master_For_Location ON tspl_location_master_For_Location.GIT_Location =TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
                        LEFT OUTER JOIN TSPL_STATE_MASTER StateMaster_ToLocation ON StateMaster_ToLocation.State_Code=TSPL_LOCATION_MASTER_1.State "
                    If clsCommon.CompairString(isCancelOrDelete, "Cancel") = CompairStringResult.Equal Then
                        Qry &= " left join TSPL_TRANSFER_ORDER_HEAD_Cancel_Data GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
                    ElseIf clsCommon.CompairString(isCancelOrDelete, "Delete") = CompairStringResult.Equal Then
                        Qry &= " left join TSPL_TRANSFER_ORDER_HEAD_Delete_Data GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
                    Else
                        Qry &= " left join TSPL_TRANSFER_ORDER_HEAD GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno "
                    End If
                    Qry &= " left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location 
                        left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State  
                        left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  GP.To_Location  
                        left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State  
                        left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 
                        left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    
                        left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   
                        left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   
                        left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   
                        left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    
                        left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    
                        left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  
                        left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    
                        left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10    
                        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL .tax1 
                        left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2    
                        left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3   
                        left outer join TSPL_TAX_MASTER as dtax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4   
                        left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5   
                        left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6    
                        left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7    
                        left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8  
                        left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9    
                        left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10  
                        LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' 
                        where 2=2 and TSPL_TRANSFER_ORDER_HEAD. Document_No = '" + clsCommon.myCstr(strCode) + "')xxx"
                Else
                    Qry = " select '" & strCancelOrDelete & "' As IsCancelled,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1+TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3 as Comp_Address,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.GSTReg_No,TSPL_TRANSFER_ORDER_HEAD.Vehicle_Mannual_No,TSPL_TRANSFER_ORDER_HEAD.Transporter_Name_Manual,TSPL_TRANSFER_ORDER_HEAD.TransferOutNo,TSPL_TRANSFER_ORDER_HEAD.Delivery_date,frmLocation.Location_Desc as From_Location,ToLocation.Location_Desc as To_Location,TSPL_TRANSFER_ORDER_HEAD.Document_No,TSPL_TRANSFER_ORDER_HEAD.Document_Date,TrasOut.Requisition_Id,TSPL_REQUISITION_HEAD.Requisition_Date,TSPL_TRANSFER_ORDER_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,TSPL_TRANSFER_ORDER_DETAIL.Unit_code,"
                    Qry += " case when TSPL_TRANSFER_ORDER_HEAD.Transfer_Type='I' Then case when isnull(TSPL_BATCH_ITEM.Batch_No,'')='' then TSPL_TRANSFER_ORDER_DETAIL.In_Qty else TSPL_BATCH_ITEM.Qty end else case when isnull(TSPL_BATCH_ITEM.Batch_No,'')='' then TSPL_TRANSFER_ORDER_DETAIL.Out_Qty else TSPL_BATCH_ITEM.Qty end End As Qty, "
                    Qry += " TSPL_BATCH_ITEM.Batch_No,TSPL_TRANSFER_ORDER_HEAD.Transfer_Type,
                    CASE
					WHEN TSPL_TRANSFER_ORDER_HEAD.Status = '1' THEN 'Approved'
					WHEN TSPL_TRANSFER_ORDER_HEAD.Status = '0' THEN 'UN Approved'
					ELSE ''
				END AS Status
from " & tbl_TSPL_TRANSFER_ORDER_HEAD & "
left join " & tbl_TSPL_TRANSFER_ORDER_Detail & " on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
left join TSPL_ITEM_MASTER on TSPL_TRANSFER_ORDER_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
left join TSPL_VEHICLE_MASTER on TSPL_TRANSFER_ORDER_HEAD.Vehicle_Code=TSPL_VEHICLE_MASTER.Vehicle_Id
left join TSPL_LOCATION_MASTER as frmLocation on TSPL_TRANSFER_ORDER_HEAD.From_Location=frmLocation.Location_Code
left join TSPL_LOCATION_MASTER as ToLocation on TSPL_TRANSFER_ORDER_HEAD.To_Location=ToLocation.Location_Code
left join TSPL_BATCH_ITEM on TSPL_TRANSFER_ORDER_DETAIL.Item_Code=TSPL_BATCH_ITEM.Item_Code and TSPL_TRANSFER_ORDER_DETAIL.Document_No=TSPL_BATCH_ITEM.Document_Code and TSPL_BATCH_ITEM.In_Out_Type='I'
left join TSPL_VENDOR_MASTER on TSPL_VEHICLE_MASTER.Transport_Id=TSPL_VENDOR_MASTER.Vendor_Code
left join TSPL_COMPANY_MASTER on TSPL_TRANSFER_ORDER_HEAD.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code
left join TSPL_TRANSFER_ORDER_HEAD as TrasOut on TSPL_TRANSFER_ORDER_HEAD.TransferOutNo=TSPL_TRANSFER_ORDER_HEAD.Document_No
left join TSPL_REQUISITION_HEAD on TSPL_TRANSFER_ORDER_HEAD.Requisition_Id=TSPL_REQUISITION_HEAD.Requisition_Id
where TSPL_TRANSFER_ORDER_HEAD.Document_No='" + clsCommon.myCstr(strCode) + "'"
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "StockTransferSTA", "Stock Transfer Advice", Nothing)
                    Else
                        frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "StockTransferSTA_Other", "Stock Transfer Advice", Nothing)
                    End If
                    'frmCRV.Close()
                Else
                    clsCommon.MyMessageBoxShow("Data not found to print.", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow("Select document.", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnSTAMilkPrint_Click(sender As Object, e As EventArgs) Handles btnSTAMilkPrint.Click
        STAMilkPrint(Me.Form_ID, Me.Text, txtDocNo.Value, Nothing)
    End Sub


    Sub STAMilkPrint(ByVal strFormID As String, ByVal strFormName As String, ByVal strCode As String, ByVal isCancelOrDelete As String)
        Try
            If clsCommon.myLen(strCode) > 0 Then
                Dim Qry As String = Nothing
                Qry = clsTransferDCC.GetSTAMlkPrint(strCode, isCancelOrDelete)
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(strFormID, CrystalReportFolder.InventoryReport, dt, "StockTransferSTAMilk", "Stock Transfer Advice", Nothing)
                    'frmCRV.Close()
                Else
                    clsCommon.MyMessageBoxShow("Data not found to print.", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow("Select document.", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnSTAProductPrint_Click(sender As Object, e As EventArgs) Handles btnSTAProductPrint.Click
        STAProductPrint(Me.Form_ID, Me.Text, txtDocNo.Value, Nothing)
    End Sub

    Sub STAProductPrint(ByVal strFormID As String, ByVal strFormName As String, ByVal strCode As String, ByVal IsCancelOrDelete As String)
        Try
            If clsCommon.myLen(strCode) > 0 Then
                Dim Qry As String = Nothing
                Qry = clsTransferDCC.GetSTAProductQry(strCode, IsCancelOrDelete)
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt.Rows IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "StockTransferSTAProductBKN", "Stock Transfer Advice", Nothing)
                    Else
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.InventoryReport, dt, "StockTransferSTAProduct", "Stock Transfer Advice", Nothing)
                    End If
                    'frmCRV.Close()
                Else
                    clsCommon.MyMessageBoxShow("Data not found to print.", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow("Select document.", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkJobWorkTransfer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkJobWorkTransfer.ToggleStateChanged
        Try
            txtFromLocation.Value = ""
            txtToLoc.Value = ""
            lblFromLoc.Text = ""
            lblToLoc.Text = ""
            If clsCommon.myLen(fndSRNO.Value) > 0 Then
                AddNew()
            End If
            If chkJobWorkTransfer.Checked Then
                If chkInternalTransfer.Checked Then
                    chkJobWorkTransfer.Checked = False
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
    End Sub

End Class