Imports System.IO
Imports common
Public Class rptDailyQtyReport
    Inherits FrmMainTranScreen
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim btnReferesh As Boolean = False
    Dim FatSnfRoundOff As Integer
    Dim corrFactor As Decimal = 0
    Dim AreaWiseBilling As Boolean = False
    Dim isPickCLRInsteadOfSNF As Boolean = False
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New UnauthorizedAccessException("Permission Denied")
        End If
    End Sub
    Sub Reset()
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        Dim viewBlank As New TableViewDefinition()
        Gv1.ViewDefinition = viewBlank
        RadPageView1.SelectedPage = RadPageViewPage1
        btnGo.Enabled = True
        'LoadShiftName()
        txtMCC.arrValueMember = Nothing
        TxtTankerNo.Value = ""
        fndArea.Value = ""
        ControlEnableDisable(True)
        LoadShif()
    End Sub
    Sub ControlEnableDisable(ByVal isEnable As Boolean)
        txtMCC.Enabled = isEnable
        fromDate.Enabled = isEnable
        ToDate.Enabled = isEnable
        dtpToDate.Enabled = isEnable
        cboItemType.Enabled = isEnable
        RadGroupBox1.Enabled = isEnable
        txtMCC_Code.Enabled = isEnable
        fndArea.Enabled = isEnable
        lblArea.Enabled = isEnable
        lblToleranceFAT.Enabled = isEnable
        lblToleranceSNF.Enabled = isEnable
        txtToleranceFat.Enabled = isEnable
        txtToleranceSNF.Enabled = isEnable
        rbtnBMCDock.Enabled = isEnable
        rbtnBMCDock.Checked = False
        rbtnBMCDCSPrint.Enabled = isEnable
        rbtnBMCDCSPrint.Checked = False
        rbtnBMCRoutePrint.Enabled = isEnable
        rbtnBMCRoutePrint.Checked = False
        rbtnDockSummary.Checked = False
        rbtnDockSummary.Enabled = isEnable
        rbtnSummary.Checked = False
        rbtnSummary.Enabled = isEnable
        txtMCC_Code.Value = Nothing
        ddlShift.Visible = False
        MyLabel6.Visible = False
        MyLabel7.Enabled = isEnable
        txtTolerenceSNF.Enabled = isEnable
        btnGainLoss.Enabled = isEnable
        rbtnDockDetail.Checked = False
        rbtnDockDetail.Enabled = isEnable
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        FatSnfRoundOff = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DailyQtyReport, clsFixedParameterCode.FATKGSNFKGRoundOff, Nothing))
        corrFactor = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.defaultCorrectionFactor, clsFixedParameterCode.MilkSetting, Nothing))
        GetReportID()
        If FatSnfRoundOff = 0 Then
            Try
                If rbtnBMCDock.Checked Then
                    BMCDock()
                    Exit Sub
                ElseIf rbtnDockSummary.Checked Then
                    DockSummary()
                    Exit Sub
                ElseIf rbtnSummary.Checked Then
                    Summary()
                    Exit Sub
                ElseIf rbtndockDetail.checked Then
                    DockDetail()
                    Exit Sub
                End If
                If rbtnTranpoterGainLoss.Checked Then
                        Dim checkRate As String

                        checkRate = "(SELECT top 1  TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and 
                                 CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, '" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "', 103))"
                        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(checkRate) 'GetDataTable(qry)
                        If dt1 Is Nothing OrElse dt1.Rows.Count <= 0 Then
                            Throw New InvalidOperationException("FAT/SNF RATE NOT FOUND!")
                        End If
                    End If
                    If rdbSummary.Checked Then
                    PageSetupReport_ID = MyBase.Form_ID & "_S"
                End If
                PageSetupReport_ID = MyBase.Form_ID
                TemplateGridview = Gv1
                Dim BaseQry As String = ""
                Dim qry As String = ""
                'Dim DockDateQry As String = ""
                Dim dt As New DataTable
                'Dim whre As String = ""
                'qry = " select TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id,TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No , TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo ,TSPL_MILK_COLLECTION_MCC.Route_Code  ,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME ,TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No
                '       ,TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty , ROUND(TSPL_MILK_COLLECTION_MCC_DETAIL.FAT,1,0) as MCC_FAT,round (TSPL_MILK_COLLECTION_MCC_DETAIL.SNF,1,0) as MCC_SNF,round (TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG,0,0) as MCC_FATKG , round (TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG,0,0) as MCC_SNFKG,
                '       isnull(XXXDCS.qty,0) as DCS_Qty,isnull(XXXDCS.FAT,0) as DCS_FAT , isnull(XXXDCS.SNF,0) as DCS_SNF, isnull(XXXDCS.FATKG,0) as DCS_FATKG, isnull(XXXDCS.SNFKG,0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty,0) as Diff_Qty, ROUND(TSPL_MILK_COLLECTION_MCC_DETAIL.FAT,1,0) - isnull(XXXDCS.FAT,0)  as Diff_FAT, round (TSPL_MILK_COLLECTION_MCC_DETAIL.SNF,1,0) - isnull(XXXDCS.SNF,0)  as Diff_SNF, round (TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG,0,0) - isnull(XXXDCS.FATKG,0)  as Diff_FATKG, round (TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG,0,0) - isnull(XXXDCS.SNFKG,0)  as Diff_SNFKG
                '       from TSPL_MILK_COLLECTION_MCC_DETAIL
                '       left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No=TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No
                '       left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code
                '       left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_MCC.Route_Code
                '       left outer join (  
                '       select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty ) as qty  , round( sum (TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG), 0,0) as FATKG ,  round( sum (TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG), 0,0) as SNFKG ,
                '       ROUND(((sum (TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG) / sum (TSPL_MILK_COLLECTION_DCS_DETAIL.qty )) * 100),1,0) as FAT , ROUND(((sum (TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG) / sum (TSPL_MILK_COLLECTION_DCS_DETAIL.qty )) * 100),1,0) as SNF 
                '       from TSPL_MILK_COLLECTION_DCS_DETAIL 
                '       left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                '       left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                '       where convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) = convert (date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and TSPL_MILK_COLLECTION_DCS.Status = 1
                '       group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                '       where convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) = convert (date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and TSPL_MILK_COLLECTION_MCC.Status = 1"
                Dim strMilkCollectionMCCStatus As String = ""
                Dim strMilkcollectionDCSStatus As String = ""
                If Not rbtnDock.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnDockShiftWise.Checked Then
                    If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                        strMilkCollectionMCCStatus = " and TSPL_MILK_COLLECTION_MCC.Status = 1 "
                        strMilkcollectionDCSStatus = " and TSPL_MILK_COLLECTION_DCS.Status = 1 "
                    ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                        strMilkCollectionMCCStatus = " and TSPL_MILK_COLLECTION_MCC.Status = 0 "
                        strMilkcollectionDCSStatus = " and TSPL_MILK_COLLECTION_DCS.Status = 0 "
                    End If
                End If
                'qry = "  select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg , max(Entered_SNFKg) as Entered_SNFKg , max(PK_Id) as PK_Id, Document_No ,  max(Document_Date) as Document_Date ,max(MCC_Code) as MCC_Code , max(MCC_NAME) as MCC_NAME, max(UploaderNo) as  UploaderNo,max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME ,max(Tanker_No) as Tanker_No ,max(Vehicle_No) as Vehicle_No
                '            , sum(MCC_Qty) as MCC_Qty , ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_FAT ,ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_SNF ,sum(MCC_FATKG) as MCC_FATKG , sum(MCC_SNFKG) as MCC_SNFKG,
                '            max(DCS_Qty) as DCS_Qty, max(DCS_FAT) as DCS_FAT , max(DCS_SNF) as DCS_SNF,  max(DCS_FATKG) as DCS_FATKG ,max(DCS_SNFKG) as DCS_SNFKG ,
                'max(DCS_Qty) - sum(MCC_Qty)  as Diff_Qty,
                'max(DCS_FAT) - ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_FAT 
                ',max(DCS_SNF) - ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_SNF
                ', max(DCS_FATKG) - sum(MCC_FATKG)  as  Diff_FATKG, 
                '    max(DCS_SNFKG) - sum(MCC_SNFKG) as Diff_SNFKG  
                '            from   (select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg , TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id,TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No ,convert (varchar,TSPL_MILK_COLLECTION_MCC.Document_Date,103) as  Document_Date ,TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo ,TSPL_MILK_COLLECTION_MCC.Route_Code  ,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME ,TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No
                '            ,TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty , TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT,TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF,TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG , TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG,
                '            isnull(XXXDCS.qty,0) as DCS_Qty,isnull(XXXDCS.FAT,0) as DCS_FAT , isnull(XXXDCS.SNF,0) as DCS_SNF, isnull(XXXDCS.FATKG,0) as DCS_FATKG, isnull(XXXDCS.SNFKG,0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty,0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT,0)  as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF,0)  as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG,0)  as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG,0)  as Diff_SNFKG
                '            from TSPL_MILK_COLLECTION_MCC_DETAIL
                '            left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No=TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No
                '            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code
                '            left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_MCC.Route_Code
                '            left outer join (  
                '            select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty ) as qty  ,  sum (TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG) as FATKG ,   sum (TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG) as SNFKG ,
                '            ((sum (TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG) / sum (TSPL_MILK_COLLECTION_DCS_DETAIL.qty )) * 100) as FAT , ((sum (TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG) / sum (TSPL_MILK_COLLECTION_DCS_DETAIL.qty )) * 100) as SNF 
                'from TSPL_MILK_COLLECTION_DCS_DETAIL 
                '            left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                '            left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                '            where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") + "',103) " + strMilkcollectionDCSStatus + "
                '            group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                '            where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") + "',103)) xyz group by Document_No  " + strMilkCollectionMCCStatus + " "
                'qry = "select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id, Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME, max(Tanker_No) as Tanker_No, max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 1, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 1, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, sum(DCS_Qty) as DCS_Qty, (sum(DCS_FATKG)/sum(DCS_Qty)*100) as DCS_FAT, (sum(DCS_SNFKG)/sum(DCS_Qty)*100) as DCS_SNF, sum(DCS_FATKG) as DCS_FATKG, sum(DCS_SNFKG) as DCS_SNFKG, sum(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (sum(DCS_FATKG)/sum(DCS_Qty)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 1, 0 ) as Diff_FAT, (sum(DCS_SNFKG)/sum(DCS_Qty)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 1, 0 ) as Diff_SNF, sum(DCS_FATKG) - sum(MCC_FATKG) as Diff_FATKG, sum(DCS_SNFKG) - sum(MCC_SNFKG) as Diff_SNFKG from ( select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo, TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_MILK_COLLECTION_MCC.Tanker_No, TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG from TSPL_MILK_COLLECTION_MCC_DETAIL left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) / sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ) ) * 100 ) as SNF
                '                       from TSPL_MILK_COLLECTION_DCS_DETAIL 
                '                        left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                '            left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                '            where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") + "',103) " + strMilkcollectionDCSStatus + "
                '            group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                '            where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") + "',103)) xyz group by Document_No  " + strMilkCollectionMCCStatus + " "
                If rbtnBmcSummary.Checked Then
                    qry = "select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id,  max(Document_No)Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME, max(Tanker_No) as Tanker_No,max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, max(DCS_Qty) as DCS_Qty, isnull( (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_FAT, isnull( (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_SNF, isnull(max(DCS_FATKG),0) as DCS_FATKG,isnull( max(DCS_SNFKG),0) as DCS_SNFKG, max(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_FAT, (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_SNF, sum(MCC_FATKG) -max(DCS_FATKG) as Diff_FATKG, sum(MCC_SNFKG)-max(DCS_SNFKG) as Diff_SNFKG,max(DCSCode)DCSCode,MAX(Alias_Name)Alias_Name from ( select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo,
                         (CASE WHEN TSPL_VLC_MASTER_HEAD.isOwnBMC = 1 THEN TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader ELSE '' END) AS DCSCode,
                       TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG ,Case When IsNull(TSPL_VENDOR_MASTER.Alies_Name,'')<>'' Then TSPL_VENDOR_MASTER.Alies_Name Else TSPL_VENDOR_MASTER.Vendor_Name End As Alias_Name from TSPL_MILK_COLLECTION_MCC_DETAIL
left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No 
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code 
LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.MCC=TSPL_MCC_MASTER.MCC_Code
Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code
left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) /nullif( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / nullif(sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as SNF
                        from TSPL_MILK_COLLECTION_DCS_DETAIL 
                        left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                        left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                        where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) " & strMilkcollectionDCSStatus & "
                        group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                        where  TSPL_VLC_MASTER_HEAD.isOwnBMC=1 and convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)" & strMilkCollectionMCCStatus & ") xyz group by UploaderNo  "
                Else
                    qry = "select max(Reject_Type)Reject_Type,max(Reject_Recovery_Per)Reject_Recovery_Per,max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id, Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME, max(Tanker_No) as Tanker_No,max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, max(DCS_Qty) as DCS_Qty, isnull( (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_FAT, isnull( (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_SNF, isnull(max(DCS_FATKG),0) as DCS_FATKG,isnull( max(DCS_SNFKG),0) as DCS_SNFKG, max(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_FAT, (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_SNF, sum(MCC_FATKG) -max(DCS_FATKG) as Diff_FATKG, sum(MCC_SNFKG)-max(DCS_SNFKG) as Diff_SNFKG from ( select TSPL_MILK_COLLECTION_MCC.Reject_Type,TSPL_MILK_COLLECTION_MCC.Reject_Recovery_Per,TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo, TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG from TSPL_MILK_COLLECTION_MCC_DETAIL left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) /nullif( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / nullif(sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as SNF
                        from TSPL_MILK_COLLECTION_DCS_DETAIL 
                        left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                        left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                        where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) " & strMilkcollectionDCSStatus & "
                        group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                        where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) " & strMilkCollectionMCCStatus & ") xyz group by UploaderNo, Document_No   "
                End If
                If rdbSummary.Checked Then
                    qry = " select XXXFinal.Document_No, max(XXXFinal.Document_Date) as Document_Date ,max( XXXFinal.Route_Code) as Route_Code, max(XXXFinal.ROUTE_NAME) as ROUTE_NAME ,
                            max(XXXFinal.Tanker_No ) as Tanker_No,case when max(XXXFinal.Reject_Type) is null then 'GOOD' ELSE max(XXXFinal.Reject_Type) END AS Reject_Types,Case when max(XXXFinal.Reject_Recovery_Per) is null then 0 else  max(XXXFinal.Reject_Recovery_Per) end as Reject_Recovery_Pers, max(XXXFinal.Vehicle_No) as Vehicle_No, max(Entered_Qty) as Entered_Qty,((max (Entered_FATKg) / nullif (max (Entered_Qty),0)) * 100) as Entered_FAT, ((max (Entered_SNFKg) / nullif (max (Entered_Qty),0)) * 100) as Entered_SNF , max(Entered_FATKg) as Entered_FATKg,max(Entered_SNFKg) as Entered_SNFKg , 
                            sum( XXXFinal.MCC_Qty) as  MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_FAT, ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_SNF,sum (MCC_FATKG) as MCC_FATKG , sum (MCC_SNFKG) as MCC_SNFKG , sum (DCS_Qty) as DCS_Qty , ROUND(((sum (DCS_FATKG) / nullif( sum (DCS_Qty),0)) * 100),1,0) as DCS_FAT,ROUND(((sum (DCS_SNFKG) / nullif (sum (DCS_Qty),0)) * 100),1,0) as DCS_SNF , sum(DCS_FATKG) as DCS_FATKG , sum(DCS_SNFKG) as DCS_SNFKG, sum (Diff_Qty) as Diff_Qty ,
                            max(DCS_FAT) - ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_FAT, max(DCS_SNF) - ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_SNF,sum(Diff_FATKG) as Diff_FATKG, sum(Diff_SNFKG) as Diff_SNFKG,  max(Entered_Qty) - sum( XXXFinal.MCC_Qty)  as DiffEnteredVsMCC_Qty ,
                            ((max (Entered_FATKg) / nullif (max (Entered_Qty),0)) * 100) - ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0)  as DiffEnteredVsMCC_FAT ,
                            ((max (Entered_SNFKg) / nullif (max (Entered_Qty),0)) * 100) - ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as DiffEnteredVsMCC_SNF ,
                            max(Entered_FATKg) -  sum (MCC_FATKG)  as DiffEnteredVsMCC_FATKG , max(Entered_SNFKg) - sum (MCC_SNFKG) as DiffEnteredVsMCC_SNFKG from ( " & qry & " ) XXXFinal group by XXXFinal.Document_No order by convert (datetime, max(XXXFinal.Document_Date),103) asc "
                ElseIf rdbDetails.Checked Then
                    '                qry = "select XXXFinal.PK_Id as PK_Id, XXXFinal.Document_No as Document_No, XXXFinal.Document_Date As Document_Date, XXXFinal.MCC_Code As MCC_Code, XXXFinal.MCC_NAME As MCC_NAME, XXXFinal.UploaderNo As UploaderNo, XXXFinal.Route_Code As Route_Code, XXXFinal.ROUTE_NAME As ROUTE_NAME, XXXFinal.Tanker_No As Tanker_No, XXXFinal.Vehicle_No As Vehicle_No, Entered_Qty As Entered_Qty, ((Entered_FATKg / nullif(Entered_Qty, 0)) * 100) As Entered_FAT, ((Entered_SNFKg / nullif(Entered_Qty, 0)) * 100) As Entered_SNF, Entered_FATKg As Entered_FATKg, Entered_SNFKg As Entered_SNFKg, XXXFinal.MCC_Qty As MCC_Qty, ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As MCC_FAT, ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As MCC_SNF, MCC_FATKG As MCC_FATKG, MCC_SNFKG As MCC_SNFKG, DCS_Qty As DCS_Qty, ROUND(((DCS_FATKG / nullif(DCS_Qty, 0)) * 100), 1, 0) As DCS_FAT, ROUND(((DCS_SNFKG / nullif(DCS_Qty, 0)) * 100), 1, 0) As DCS_SNF, DCS_FATKG As DCS_FATKG, DCS_SNFKG As DCS_SNFKG, Diff_Qty As Diff_Qty, DCS_FAT - ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_FAT, DCS_SNF - ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_SNF, Diff_FATKG As Diff_FATKG, Diff_SNFKG As Diff_SNFKG, Entered_Qty - XXXFinal.MCC_Qty As DiffEnteredVsMCC_Qty, ((Entered_FATKg / nullif(Entered_Qty, 0)) * 100) - ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As DiffEnteredVsMCC_FAT, ((Entered_SNFKg / nullif(Entered_Qty, 0)) * 100) - ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As DiffEnteredVsMCC_SNF, Entered_FATKg - MCC_FATKG As DiffEnteredVsMCC_FATKG, Entered_SNFKg - MCC_SNFKG as DiffEnteredVsMCC_SNFKG from( " + qry + " ) XXXFinal  
                    'left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No=XXXFinal.Document_No"
                    qry = "select XXXFinal.PK_Id as PK_Id, XXXFinal.Document_No as Document_No, XXXFinal.Document_Date As Document_Date, XXXFinal.MCC_Code As MCC_Code, XXXFinal.MCC_NAME As MCC_NAME, XXXFinal.UploaderNo As UploaderNo, XXXFinal.Route_Code As Route_Code, XXXFinal.ROUTE_NAME As ROUTE_NAME, XXXFinal.Tanker_No As Tanker_No, XXXFinal.Vehicle_No As Vehicle_No, XXXFinal.MCC_Qty As MCC_Qty, Convert(Decimal(18,2),(MCC_FATKG / nullif(MCC_Qty, 0)) * 100) As MCC_FAT, convert(decimal(18,2),(MCC_SNFKG / nullif(MCC_Qty, 0)) * 100) As MCC_SNF, MCC_FATKG As MCC_FATKG, MCC_SNFKG As MCC_SNFKG, DCS_Qty As DCS_Qty,isnull( ROUND(((DCS_FATKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_FAT, isnull(ROUND(((DCS_SNFKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_SNF, isnull(DCS_FATKG,0) As DCS_FATKG, isnull(DCS_SNFKG,0) As DCS_SNFKG, Diff_Qty As Diff_Qty, DCS_FAT - ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_FAT, DCS_SNF - ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_SNF, Diff_FATKG As Diff_FATKG, Diff_SNFKG As Diff_SNFKG from( " & qry & " ) XXXFinal  
                           left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No=XXXFinal.Document_No"
                ElseIf rbtnBmcSummary.Checked Then
                    qry = "select XXXFinal.PK_Id as PK_Id, XXXFinal.Document_No as Document_No, XXXFinal.Document_Date As Document_Date, XXXFinal.MCC_Code As MCC_Code, XXXFinal.MCC_NAME As MCC_NAME, XXXFinal.UploaderNo As UploaderNo,XXXFinal.DCSCode as DCSCode,XXXFinal.Alias_Name, XXXFinal.Route_Code As Route_Code, XXXFinal.ROUTE_NAME As ROUTE_NAME, XXXFinal.Tanker_No As Tanker_No, XXXFinal.Vehicle_No As Vehicle_No, XXXFinal.MCC_Qty As MCC_Qty, Convert(Decimal(18,2),(MCC_FATKG / nullif(MCC_Qty, 0)) * 100) As MCC_FAT, convert(decimal(18,2),(MCC_SNFKG / nullif(MCC_Qty, 0)) * 100) As MCC_SNF, MCC_FATKG As MCC_FATKG, MCC_SNFKG As MCC_SNFKG, DCS_Qty As DCS_Qty,isnull( ROUND(((DCS_FATKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_FAT, isnull(ROUND(((DCS_SNFKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_SNF, isnull(DCS_FATKG,0) As DCS_FATKG, isnull(DCS_SNFKG,0) As DCS_SNFKG, Diff_Qty As Diff_Qty, DCS_FAT - ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_FAT, DCS_SNF - ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_SNF, Diff_FATKG As Diff_FATKG, Diff_SNFKG As Diff_SNFKG from( " & qry & " ) XXXFinal  
                            left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No=XXXFinal.Document_No"
                    'qry = qry + " order by TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No asc "
                ElseIf rbtnDock.Checked OrElse rbtnDockDateWise.Checked OrElse rbtnDockShiftWise.Checked OrElse rbtnAliasName.Checked Then
                    qry = ""
                    Dim strFilterQry As String = ""
                    Dim strQryStatus1 As String = ""
                    Dim strQryStatus2 As String = ""
                    Dim Area As String = ""
                    Dim AreaName As String = Nothing
                    AreaName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))
                    If txtMCC_Code.Value.Length > 0 Then
                        strFilterQry = " and TSPL_MCC_MASTER.MCC_Code in ('" & txtMCC_Code.Value & "') "
                    End If
                    If clsCommon.myLen(fndArea.Value) > 0 Then
                        Area += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                    End If
                    If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                        strQryStatus1 = "where TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=1"
                        strQryStatus2 = "where TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Status=1"
                    ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                        strQryStatus1 = "where TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=0"
                        strQryStatus2 = "where TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Status=0"
                    End If
                    qry = "Select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SNo,CONVERT(varchar(10), tmp.Shift_Date, 103) As SDate ,tmp.Shift As Shift,No_Of_Cans As Qty,Milk_Weight As Milk_Wtd,FAT as FAT,"
                    If isPickCLRInsteadOfSNF Then
                        qry += " tmp.SNF/4+0.2*tmp.FAT+(select Description from TSPL_Fixed_parameter where type='defaultCorrectionFactor' and code='MilkSetting') as SNF,convert(decimal(18,2),((Milk_Weight*(tmp.SNF/4+0.2*tmp.FAT+(select Description from TSPL_Fixed_parameter where type='defaultCorrectionFactor' and code='MilkSetting')))/100)) As SNF_KG,"
                    Else
                        qry += " SNF as SNF, convert(decimal(18,2),((Milk_Weight*SNF)/100)) As SNF_KG, "
                    End If
                    qry += "  convert(decimal(18,2),((Milk_Weight*FAT)/100)) As Fat_KG,'' As Rate,'' As Amt,tmp.VLC_Code, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As DCS_Uploader_Code,TSPL_VLC_MASTER_HEAD.VLC_Code As DCS_Code,TSPL_VLC_MASTER_HEAD.VLC_Name  As DCS_Name,Case When IsNull(TSPL_VENDOR_MASTER.Alies_Name,'')<>'' Then TSPL_VENDOR_MASTER.Alies_Name Else TSPL_VENDOR_MASTER.Vendor_Name End As Alias_Name
                        ,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader  As BMC_Uploader_Code,TSPL_MCC_MASTER.MCC_Code As BMC_Code,TSPL_MCC_MASTER.MCC_NAME As BMC_Name ,'" & AreaName & "' as Area, TSPL_VLC_MASTER_HEAD.Route_Code
                        from(
                        Select Shift_Date,Shift,No_Of_Cans,Milk_Weight,FAT,SNF,VLC_Code  from TSPL_MILK_SHIFT_UPLOADER_DETAIL
                        Left Outer Join TSPL_MILK_SHIFT_UPLOADER_HEAD ON TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No=TSPL_MILK_SHIFT_UPLOADER_DETAIL.Document_No
	                    " & strQryStatus1 & "
                        Union All
                        Select Shift_Date,Shift,No_Of_Cans,Milk_Weight,FAT,SNF,VLC_Code from TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL
	                    Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD ON TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
	                    " & strQryStatus2 & ")tmp
                        Left Outer Join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=tmp.VLC_Code
                        Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
                        where convert (date,Shift_Date,103) >= convert (date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,Shift_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)" & strFilterQry & Area
                    If rbtnAliasName.Checked Then
                        qry = "Select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SNo,Alias_Name As [Alias Name] ,SUM(Qty)Qty ,SUM(Milk_Wtd)[Milk Weight] ,MAX(FAT)FAT,MAX(SNF)SNF,MAX(SNF_KG)[SNF KG],MAX(Fat_KG)[FAT KG] from(" & qry & ")xyz Group By Alias_Name Order By Alias_Name"
                    End If
                ElseIf (rbtnTranpoterGainLoss.Checked) OrElse (rbtnTranspoterGainlossSummary.Checked) Then
                    If txtToleranceFat.Value < 0 OrElse txtToleranceFat.Value >= 100 Then
                        clsCommon.MyMessageBoxShow(Me, "Invalid Input !, Please enter valid Tolerance Fat%.", Me.Text)
                        txtToleranceFat.Focus()
#Disable Warning
                        Exit Sub
#Enable Warning
                    End If
                    If txtToleranceSNF.Value < 0 OrElse txtToleranceSNF.Value >= 100 Then
                        clsCommon.MyMessageBoxShow(Me, "Invalid Input !, Please enter valid Tolerance Fat%.", Me.Text)
                        txtToleranceSNF.Focus()
#Disable Warning
                        Exit Sub
#Enable Warning
                    End If
                    Dim BaseQuery As String = " select GetAllGainLossRate.*,
                         case when GetAllGainLossRate.DiffMCCVsEntered_FATKG >= 0 then (
                         case when (GetAllGainLossRate.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_FATKG -" & clsCommon.myCstr(txtToleranceFat.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate) else 0 end)
                         else 0 end as FAT_AMT,case when GetAllGainLossRate.DiffMCCVsEntered_SNFKG > 0 then (case when (GetAllGainLossRate.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_SNFKG -" & clsCommon.myCstr(txtToleranceSNF.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate) else 0 end) else 0 end as SNF_AMT,
                         TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date from ( select GetRateCode.*,( SELECT top 1 TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, GetRateCode.Document_Date, 103) order by TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date desc ) as GainLossCode from ( SELECT max(XXXFinal.Reject_Type)Reject_Type,max(XXXFinal.Reject_Recovery_Per)Reject_Recovery_Per,XXXFinal.Document_No, MAX(XXXFinal.Document_Date) AS Document_Date, MAX(XXXFinal.Route_Code) AS Route_Code, MAX(XXXFinal.ROUTE_NAME) AS ROUTE_NAME, MAX(XXXFinal.Tanker_No) AS Tanker_No, MAX(XXXFinal.Vehicle_No) AS Vehicle_No, MAX(Entered_Qty) AS Entered_Qty, ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_FAT, ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_SNF, MAX(Entered_FATKg) AS Entered_FATKg, MAX(Entered_SNFKg) AS Entered_SNFKg, SUM(XXXFinal.MCC_Qty) AS MCC_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_SNF, SUM(MCC_FATKG) AS MCC_FATKG, SUM(MCC_SNFKG) AS MCC_SNFKG, SUM(DCS_Qty) AS DCS_Qty, ROUND(((SUM(DCS_FATKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_FAT, ROUND(((SUM(DCS_SNFKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_SNF, SUM(DCS_FATKG) AS DCS_FATKG, SUM(DCS_SNFKG) AS DCS_SNFKG, SUM(Diff_Qty) AS Diff_Qty, MAX(DCS_FAT) - ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_FAT, MAX(DCS_SNF) - ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_SNF, SUM(Diff_FATKG) AS Diff_FATKG, SUM(Diff_SNFKG) AS Diff_SNFKG, SUM(XXXFinal.MCC_Qty)- MAX(Entered_Qty) AS DiffMCCVsEntered_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_SNF, SUM(MCC_FATKG)-MAX(Entered_FATKg) AS DiffMCCVsEntered_FATKG, SUM(MCC_SNFKG)-MAX(Entered_SNFKg) AS DiffMCCVsEntered_SNFKG from ( " & qry & " ) XXXFinal group by XXXFinal.Document_No )GetRateCode )GetAllGainLossRate left outer join TSPL_OWN_BMC_GAIN_LOSS_RATE on TSPL_OWN_BMC_GAIN_LOSS_RATE.Code=GetAllGainLossRate.GainLossCode  "
                    If (rbtnTranpoterGainLoss.Checked) Then
                        qry = "select  XXGetAllRecords.Document_No,XXGetAllRecords.Document_Date,XXGetAllRecords.Route_Code,XXGetAllRecords.ROUTE_NAME,XXGetAllRecords.Tanker_No,XXGetAllRecords.Vehicle_No,
                              XXGetAllRecords.Reject_Type as Reject_Types,XXGetAllRecords.Reject_Recovery_Per as Reject_Recovery_Pers,
                        XXGetAllRecords.Entered_Qty,XXGetAllRecords.Entered_FATKg,XXGetAllRecords.Entered_SNFKg,
                        XXGetAllRecords.MCC_Qty,XXGetAllRecords.MCC_FATKG,XXGetAllRecords.MCC_SNFKG,
                        XXGetAllRecords.DiffMCCVsEntered_Qty,
                        XXGetAllRecords.DiffMCCVsEntered_FATKG," & clsCommon.myCstr(txtToleranceFat.Value) & " as FAT_Tolerence,case when (isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ",0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ",0) end as FATKG_Recovered,Round(XXGetAllRecords.FAT_AMT,2,0) as FAT_AMT,
                        XXGetAllRecords.DiffMCCVsEntered_SNFKG," & clsCommon.myCstr(txtToleranceSNF.Value) & " as SNF_Tolerence,case when (isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ",0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ",0) end as SNFKG_Recovered,Round(XXGetAllRecords.SNF_AMT,2,0) as SNF_AMT,Round((XXGetAllRecords.FAT_AMT + XXGetAllRecords.SNF_AMT),2,0) as AMOUNT,XXGetAllRecords.GainLossCode as GainLoss_Code,
                           XXGetAllRecords.Loss_FAT_Rate as Loss_FAT_Rate,
                           XXGetAllRecords.Loss_SNF_Rate as Loss_SNF_Rate,
                           XXGetAllRecords.Start_Date as Start_Date
                         from(" & BaseQuery & ") XXGetAllRecords"
                    ElseIf (rbtnTranspoterGainlossSummary.Checked) Then
                        'If txtToleranceFat.Value > 0 AndAlso txtTolerenceSNF.Value > 0 Then
                        '    Dim tolerence As Decimal = txtToleranceFat.Value
                        'End If
                        qry = "select max(PPP.Document_No)Document_No,
                        max(PPP.Document_Date)Document_Date,
                        (PPP.Route_Code)Route_Code,
                        (PPP.ROUTE_NAME)ROUTE_NAME,
                        max(PPP.Tanker_No)Tanker_No,
                        max(PPP.Vehicle_No)Vehicle_No,
                        max(ppp.Reject_type)Reject_types,
						max(ppp.Reject_Recovery_Per)Reject_Recovery_Pers,
                        sum(PPP.Entered_Qty)Entered_Qty,
						sum(PPP.Entered_FATKg)Entered_FATKg,
						sum(PPP.Entered_SNFKg)Entered_SNFKg,
                        sum(PPP.MCC_Qty)MCC_Qty,
						sum(PPP.MCC_FATKG)MCC_FATKG,
						sum(PPP.MCC_SNFKG)MCC_SNFKG,
                        sum(PPP.DiffMCCVsEntered_Qty)DiffMCCVsEntered_Qty,
						sum(ppp.DiffMCCVsEntered_FATKG)DiffMCCVsEntered_FATKG,
                          sum(ppp.FAT_Tolerence) as FAT_Tolerence, case when sum(ppp.FATKG_Recovered)<=0 then 0 else sum(ppp.FATKG_Recovered) end   as FATKG_Recovered
						  , case when sum(ppp.FATKG_Recovered)<=0 then 0 else  Sum(FAT_AMT) end as FAT_AMT,
                        sum(PPP.DiffMCCVsEntered_SNFKG)DiffMCCVsEntered_SNFKG,sum(SNF_Tolerence) as SNF_Tolerence,
						 case when Sum(ppp.SNFKG_Recovered)<=0 then 0 else Sum(ppp.SNFKG_Recovered) end as SNFKG_Recovered
						,Case when Sum(ppp.SNF_AMT)<=0 then 0 else Sum(ppp.SNF_AMT) end  as SNF_AMT
						, CASE 
            WHEN SUM(FAT_AMT) >= 0 AND SUM(SNF_AMT) >= 0 THEN 
                ISNULL(SUM(FAT_AMT), 0) + ISNULL(SUM(SNF_AMT), 0)
            WHEN SUM(FAT_AMT) >= 0 AND SUM(SNF_AMT) <= 0 THEN 
                SUM(FAT_AMT)
            WHEN SUM(FAT_AMT) <= 0 AND SUM(SNF_AMT) >= 0 THEN 
                SUM(SNF_AMT)
            ELSE 
                0
        END AS AMOUNT,max(PPP.GainLoss_Code) as GainLoss_Code,
                           sum(PPP.Loss_FAT_Rate) as Loss_FAT_Rate,
                           sum(PPP.Loss_SNF_Rate) as Loss_SNF_Rate,
                           max(PPP.Start_Date) as Start_Date
						   from (select  max(XXGetAllRecords.Document_No)Document_No,
                        max(XXGetAllRecords.Document_Date)Document_Date,
                        XXGetAllRecords.Route_Code,
                        XXGetAllRecords.ROUTE_NAME,
                        max(XXGetAllRecords.Tanker_No)Tanker_No,
                        max(XXGetAllRecords.Vehicle_No)Vehicle_No,
                        max(XXGetAllRecords.Reject_Type)Reject_type,
						max(XXGetAllRecords.Reject_Recovery_Per)Reject_Recovery_Per,
                        sum(XXGetAllRecords.Entered_Qty)Entered_Qty,
						sum(XXGetAllRecords.Entered_FATKg)Entered_FATKg,
						sum(XXGetAllRecords.Entered_SNFKg)Entered_SNFKg,
                        sum(XXGetAllRecords.MCC_Qty)MCC_Qty,
						sum(XXGetAllRecords.MCC_FATKG)MCC_FATKG,
						sum(XXGetAllRecords.MCC_SNFKG)MCC_SNFKG,
                        sum(XXGetAllRecords.DiffMCCVsEntered_Qty)DiffMCCVsEntered_Qty,
                          sum(isnull(XXGetAllRecords.Entered_FATKg,0)-isnull(XXGetAllRecords.MCC_FATKG,0))DiffMCCVsEntered_FATKG,sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100 as FAT_Tolerence,case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100  end as FATKG_Recovered,max(XXGetAllRecords.Loss_FAT_Rate)*case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100 end  as FAT_AMT,
                        sum(XXGetAllRecords.DiffMCCVsEntered_SNFKG)DiffMCCVsEntered_SNFKG,sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100 as SNF_Tolerence,case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100 end as SNFKG_Recovered,max(XXGetAllRecords.Loss_SNF_Rate)*case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100  end as SNF_AMT,max(XXGetAllRecords.Loss_FAT_Rate)*case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* '" & clsCommon.myCstr(txtTolerenceSNF.Value) & "'/100 end + max(XXGetAllRecords.Loss_SNF_Rate)*case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* '" & clsCommon.myCstr(txtTolerenceSNF.Value) & "'/100  end as AMOUNT,max(XXGetAllRecords.GainLossCode) as GainLoss_Code,
                           sum(XXGetAllRecords.Loss_FAT_Rate) as Loss_FAT_Rate,
                           sum(XXGetAllRecords.Loss_SNF_Rate) as Loss_SNF_Rate,
                           max(XXGetAllRecords.Start_Date) as Start_Date
                         from(" & BaseQuery & ") XXGetAllRecords group by Route_Code,ROUTE_NAME )PPP group  by  Route_Code,ROUTE_NAME order by Route_Code  "
                    End If
                    If clsCommon.myLen(TxtTankerNo.Value) > 0 Then
                        qry = qry & " where tanker_no ='" & TxtTankerNo.Value & "'"
                    End If
                    If rbtnTranpoterGainLoss.Checked Then
                        qry += " order by convert(date,XXGetAllRecords.Document_Date,103) asc "
                    End If
                ElseIf rdbTankerWise.Checked Then
                    qry = "SELECT  Max(XXXFinal.tanker_no) AS Tanker_No, Sum (MCC_Qty) AS MCC_Qty,  Sum (MCC_FATKG) AS MCC_FATKG, Sum (MCC_SNFKG) AS MCC_SNFKG, 
                          sum(entered_qty) AS Entered_Qty, Sum (entered_fatkg)  AS Entered_FATKg, Sum (entered_snfkg) AS Entered_SNFKg, Sum(entered_qty) - Sum(MCC_Qty) AS DiffEnteredVsMCC_Qty, Sum (entered_fatkg) - Sum (MCC_FATKG)  AS
                          DiffEnteredVsMCC_FAT, Sum (entered_snfkg)- Sum (MCC_SNFKG) AS DiffEnteredVsMCC_SNF FROM   (SELECT Max(entered_qty) AS Entered_Qty, Max(entered_fatkg) AS Entered_FATKg, Max(entered_snfkg) AS Entered_SNFKg, Max(document_date) AS Document_Date, document_no, Max(tanker_no)
                          AS Tanker_No, sum(MCC_Qty) AS MCC_Qty, Isnull(Sum(MCC_FATKG), 0) AS MCC_FATKG, Isnull(Sum(MCC_SNFKG), 0) AS MCC_SNFKG
                          FROM   (SELECT tspl_milk_collection_mcc.entered_qty, tspl_milk_collection_mcc.entered_fatkg, tspl_milk_collection_mcc.entered_snfkg,
                          tspl_milk_collection_mcc_detail.document_no, tspl_milk_collection_mcc.tanker_no, CONVERT (VARCHAR, tspl_milk_collection_mcc.document_date,103)
                          AS Document_Date, Isnull(qty, 0) AS MCC_Qty, Isnull(fat, 0) AS MCC_FAT, ISNULL(SNF, 0) AS MCC_SNF, ISNULL(FATKG, 0) AS MCC_FATKG, 
                          ISNULL(SNFKG, 0) AS MCC_SNFKG  FROM   tspl_milk_collection_mcc_detail
                          LEFT OUTER JOIN tspl_milk_collection_mcc ON tspl_milk_collection_mcc.document_no = tspl_milk_collection_mcc_detail.document_no 
                          LEFT OUTER JOIN tspl_mcc_master ON tspl_mcc_master.mcc_code = tspl_milk_collection_mcc_detail.mcc_code 
                          LEFT OUTER JOIN tspl_bulk_route_master ON tspl_bulk_route_master.route_no = tspl_milk_collection_mcc.route_code
                          WHERE  CONVERT(DATE, tspl_milk_collection_mcc.document_date, 103) >= CONVERT(DATE, '" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) AND CONVERT (DATE, tspl_milk_collection_mcc.document_date, 103)
                         <= CONVERT (DATE, '" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) ) xyz GROUP  BY  document_no) XXXFinal GROUP  BY  Tanker_No
                         ORDER  BY Tanker_No asc "
                ElseIf rbtnRouteWise.Checked Then
                    BaseQry = "( SELECT ROUTE_NO,max(ROUTE_NAME)ROUTE_NAME,  sum(entered_qty) AS Entered_Qty, Isnull(sum(entered_fatkg),0) AS Entered_FATKg, Isnull(sum(entered_snfkg),0) AS Entered_SNFKg, (document_date) AS Document_Date,	 
                          max(MCC_Qty) AS MCC_Qty, Isnull(max(MCC_FATKG), 0) AS MCC_FATKG, Isnull(max(MCC_SNFKG), 0) AS MCC_SNFKG, MAX(Reject_Type + '(' + CAST(Reject_Recovery_Per AS VARCHAR) + '%)') AS RejectInfo FROM   (
						  SELECT Trip_No, tspl_bulk_route_master.ROUTE_NO,tspl_bulk_route_master.ROUTE_NAME, ISNULL(qty,0) as entered_qty , FATKG as entered_fatkg, SNFKG as entered_snfkg,
                          tspl_milk_collection_mcc_detail.document_no, tspl_milk_collection_mcc.document_date
                          AS Document_Date, Isnull(tspl_milk_collection_mcc.Entered_Qty, 0) AS MCC_Qty, Isnull(fat, 0) AS MCC_FAT, ISNULL(SNF, 0) AS MCC_SNF, ISNULL(tspl_milk_collection_mcc.Entered_FATKg, 0) AS MCC_FATKG, 
                          ISNULL(tspl_milk_collection_mcc.Entered_SNFKg, 0) AS MCC_SNFKG,tspl_milk_collection_mcc.Reject_Type,tspl_milk_collection_mcc.Reject_Recovery_Per  FROM   tspl_milk_collection_mcc_detail
                          LEFT OUTER JOIN tspl_milk_collection_mcc ON tspl_milk_collection_mcc.document_no = tspl_milk_collection_mcc_detail.document_no 
                          LEFT OUTER JOIN tspl_mcc_master ON tspl_mcc_master.mcc_code = tspl_milk_collection_mcc_detail.mcc_code 
                          LEFT OUTER JOIN tspl_bulk_route_master ON tspl_bulk_route_master.route_no = tspl_milk_collection_mcc.route_code
                           WHERE  CONVERT(DATE, tspl_milk_collection_mcc.document_date, 103) >= CONVERT(DATE, '" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) AND CONVERT (DATE, tspl_milk_collection_mcc.document_date, 103)
                         <= CONVERT (DATE, '" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) ) xyz GROUP  BY ROUTE_NO, Document_Date,Trip_No)"
                    qry = "select * from ( SELECT 1 AS 'S.No.',  CONVERT (VARCHAR, Document_Date,103) AS Date , ROUTE_NO, max(ROUTE_NAME)ROUTE_NAME,MAX(RejectInfo) AS RejectInfo,  Sum (MCC_Qty) AS MCC_Qty, convert(decimal(18,2),(sum(MCC_FATKG)*100)/sum(MCC_Qty)) MCC_FAT_Per, convert(decimal(18,2),(sum(MCC_SNFKG)*100)/sum(MCC_Qty)) MCC_SNF_Per,
                          sum (MCC_FATKG) AS MCC_FATKG, sum (MCC_SNFKG) AS MCC_SNFKG, sum(entered_qty) AS Entered_Qty, convert(decimal(18,2),(sum(Entered_FATKg)*100)/sum(Entered_Qty)) Entered_FAT_Per,convert(decimal(18,2),(sum(Entered_SNFKg)*100)/sum(Entered_Qty)) Entered_SNF_Per, Sum (entered_fatkg)  AS Entered_FATKg,
                          Sum (entered_snfkg) AS Entered_SNFKg, Sum(MCC_Qty) - Sum(Entered_Qty) AS DiffEnteredVsMCC_Qty, Sum (MCC_FATKG) - Sum (Entered_FATKg)  AS DiffEnteredVsMCC_FAT, Sum (MCC_SNFKG)- Sum (Entered_SNFKg) AS DiffEnteredVsMCC_SNF FROM
                        " & BaseQry & " XXXFinal GROUP  BY  ROUTE_NO , Document_Date"
                    qry += "" & Environment.NewLine & " Union all " & Environment.NewLine & ""
                    qry += " SELECT 2 AS 'S.No.', 'Total' AS Date , ROUTE_NO, max(ROUTE_NAME)ROUTE_NAME,'' as RejectInfo,  Sum (MCC_Qty) AS MCC_Qty, convert(decimal(18,2),(sum(MCC_FATKG)*100)/sum(MCC_Qty)) MCC_FAT_Per, convert(decimal(18,2),(sum(MCC_SNFKG)*100)/sum(MCC_Qty)) MCC_SNF_Per,
                          sum (MCC_FATKG) AS MCC_FATKG, sum (MCC_SNFKG) AS MCC_SNFKG, sum(entered_qty) AS Entered_Qty, convert(decimal(18,2),(sum(Entered_FATKg)*100)/sum(Entered_Qty)) Entered_FAT_Per,convert(decimal(18,2),(sum(Entered_SNFKg)*100)/sum(Entered_Qty)) Entered_SNF_Per, Sum (entered_fatkg)  AS Entered_FATKg,
                          Sum (entered_snfkg) AS Entered_SNFKg, Sum(MCC_Qty) - Sum(Entered_Qty) AS DiffEnteredVsMCC_Qty, Sum (MCC_FATKG) - Sum (Entered_FATKg)  AS DiffEnteredVsMCC_FAT, Sum (MCC_SNFKG)- Sum (Entered_SNFKg) AS DiffEnteredVsMCC_SNF FROM " & BaseQry & " XXXFinal GROUP  BY  ROUTE_NO "
                    qry += "" & Environment.NewLine & " Union all " & Environment.NewLine & ""
                    qry += " SELECT 3 AS 'S.No.', '' AS Date , 'Grand Total' as ROUTE_NO, '' as ROUTE_NAME,'' as RejectInfo,  Sum (MCC_Qty) AS MCC_Qty, convert(decimal(18,2),(sum(MCC_FATKG)*100)/sum(MCC_Qty)) MCC_FAT_Per, convert(decimal(18,2),(sum(MCC_SNFKG)*100)/sum(MCC_Qty)) MCC_SNF_Per,
                          sum (MCC_FATKG) AS MCC_FATKG, sum (MCC_SNFKG) AS MCC_SNFKG, sum(entered_qty) AS Entered_Qty, convert(decimal(18,2),(sum(Entered_FATKg)*100)/sum(Entered_Qty)) Entered_FAT_Per,convert(decimal(18,2),(sum(Entered_SNFKg)*100)/sum(Entered_Qty)) Entered_SNF_Per, Sum (entered_fatkg)  AS Entered_FATKg,
                          Sum (entered_snfkg) AS Entered_SNFKg, Sum(MCC_Qty) - Sum(Entered_Qty) AS DiffEnteredVsMCC_Qty, Sum (MCC_FATKG) - Sum (Entered_FATKg)  AS DiffEnteredVsMCC_FAT, Sum (MCC_SNFKG)- Sum (Entered_SNFKg) AS DiffEnteredVsMCC_SNF FROM " & BaseQry & " XXXFinal  )xx "
                    qry += "ORDER  BY ROUTE_NO ,[S.No.], Date "
                ElseIf rdbCollectionWise.Checked Then
                    qry = " select TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.PK_Id,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_No,FORMAT(TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date, 'dd/MM/yyyy') AS Document_Date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code,
                            TSPL_MCC_MASTER.MCC_NAME,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code,
                            TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Tanker_No,Vehicle_No,Entered_Qty,Entered_FATKg,Entered_SNFKg,
                            CONVERT(DECIMAL(10, 2),( CLR/4 + (0.2*TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.FAT) + " & clsCommon.myCstr(corrFactor) & "),2) AS SNFPer,
                            TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.CLR ,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.FAT,FORMAT(Collection_Date, 'dd/MM/yyyy') AS Collection_Date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Name,
                            TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,Qty,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.FAT AS DCS_FAT,SNF,FATKG,SNFKG,CONVERT(DECIMAL(10, 2),(( (SNFKG) * 100 / NULLIF( (Qty),0))),2) as DCS_SNFPER
                            from TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL
                            left outer join TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS on TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_No= TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.Document_No
                            LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.VLC_Code
                            left outer join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code
                            left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code 
                            WHERE convert(date, TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) 
                            and convert (date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) order by PK_Id "
                ElseIf rdbMultiple.Checked Then
                    qry = "  SELECT Document_No,max(FORMAT(final.Document_Date, 'dd/MM/yyyy'))Document_Date,max(final.MCC_Code)MCC_Code,max(final.MCC_NAME)MCC_NAME,max(final.Mcc_Code_VLC_Uploader)Mcc_Code_VLC_Uploader,
                             max(final.Route_Code)Route_Code,max(final.ROUTE_NAME)ROUTE_NAME,max(final.Tanker_No)Tanker_No,max(final.Vehicle_No)Vehicle_No,max(final.Entered_Qty)Entered_Qty,
                             max(final.Entered_FATKg)Entered_FATKg,max(final.Entered_SNFKg)Entered_SNFKg,max(final.CLR)CLR,max(final.FATDAYS)FAT,max(final.VLC_Code)VLC_Code,max(final.VLC_Name)VLC_Name,max(final.VLC_Code_VLC_Uploader)VLC_Code_VLC_Uploader,sum(final.Qty)Qty,sum(final.FATKG)FATKG,sum(final.SNFKG)SNFKG, 
                             CONVERT(DECIMAL(10, 2),ROUND(((sum (FATKG) / NULLIF(sum (Qty),0)) * 100),2)) as DCS_FAT,CONVERT(DECIMAL(10, 2),ROUND(((sum (SNFKG) / nullif (sum (Qty),0)) * 100),2)) as DCS_SNF,
                             MAX(Entered_Qty) - SUM(Qty) AS Diff_Qty,MAX(Entered_FATKg) - SUM(FATKG) AS Diff_FAT,MAX(Entered_SNFKg) - SUM(SNFKG) AS Diff_SNF
                             FROM 
                            (select TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_No,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code,
                            TSPL_MCC_MASTER.MCC_NAME,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.CLR,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.FAT as FATDAYS,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,
                            TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Tanker_No,Vehicle_No,Entered_Qty,Entered_FATKg,Entered_SNFKg,Collection_Date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.VLC_Code,
                            TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,Qty,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.FAT,SNF,FATKG,SNFKG from TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL
                            left outer join TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS on TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_No= TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.Document_No
                            LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS_DETAIL.VLC_Code
                            left outer join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code
                            left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code 
                            WHERE convert(date, TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) 
                            and convert (date,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)) FINAL  group by Document_No"
                ElseIf rbtnBMCTankerCollection.Checked Then
                    qry = "Select Convert(varchar(10),xxx.Document_Date,103) As [BMC Date],xxx.ROUTE_NO As [Route No],(xxx.ROUTE_NAME) As [Route Name],(xxx.MCC_NAME) As [MCC Name],
                            (xxx.VLC_Code_VLC_Uploader)[VLC Uploader Code],(xxx.Alias_Name)[Alias Name],Convert(int,xxx.Qty)Qty,Convert(Decimal(18,1),xxx.FAT_Per) As [FAT(%)],Convert(Decimal(18,1),xxx.SNF_Per) As [SNF(%)],Convert(Decimal(18,3),xxx.Entered_FATKg) As [FAT KG],Convert(Decimal(18,3),xxx.Entered_SNFKg) As [SNF KG],xxx.TankerDocument_Date As [Tanker Date],xxx.TankerTanker_No AS[Tanker No],
                            Convert(int,xxx.TankerQty) AS [Tanker Qty],Convert(Decimal(18,1),xxx.TankerFAT_Per)[Tanker Fat(%)],Convert(decimal(18,1),xxx.TankerSNF_Per)[Tanker SNF(%)],Convert(Decimal(18,3),xxx.TankerEntered_FATKg) As [Tanker FAT KG],Convert(Decimal(18,3),xxx.TankerEntered_SNFKg) As [Tanker SNF KG],Convert(int,(xxx.TankerQty-xxx.Qty))[Diff. Qty],
                            Convert(Decimal(18,1),(xxx.TankerFAT_Per-xxx.FAT_Per))[Diff. FAT(%)],Convert(Decimal(18,1),(xxx.TankerSNF_Per-xxx.SNF_Per))[Diff. SNF(%)],
                            Convert(Decimal(18,3),(xxx.TankerEntered_FATKg-xxx.Entered_FATKg))[Diff. FAT KG],Convert(decimal(18,3),(xxx.TankerEntered_SNFKg-xxx.Entered_SNFKg))[Diff. SNF KG] 
							from 	
							(Select BMCData.*,TankerData.Document_Date As TankerDocument_Date,TankerData.ROUTE_NO As TankerROUTE_NO,TankerData.Tanker_No As TankerTanker_No,IsNull(TankerData.Qty,0)TankerQty,Isnull(TankerData.FAT_Per,0)TankerFAT_Per,IsNull(TankerData.SNF_Per,0)TankerSNF_Per,IsNull(TankerData.Entered_FATKg,0)TankerEntered_FATKg,IsNull(TankerData.Entered_SNFKg,0)TankerEntered_SNFKg from
							(select  (tspl_gate_entry_details.Date_And_Time)Document_Date,TSPL_BULK_ROUTE_MASTER.ROUTE_NO,Max(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME)ROUTE_NAME,
                             Max(tspl_gate_entry_details.Tanker_No)Tanker_No,Sum(Qty_In_Kg)Qty,Avg(FAT_Per)FAT_Per,Avg(SNF_Per)SNF_Per, sum((tspl_gate_entry_details.fat_per*tspl_gate_entry_details.Qty_In_Kg)/100) As Entered_FATKg,
                             sum((tspl_gate_entry_details.snf_Per*tspl_gate_entry_details.Qty_In_Kg)/100) As Entered_SNFKg  From tspl_gate_entry_details
                            Left Outer Join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=tspl_gate_entry_details.ROUTE_NO
							Where Convert(Date,Date_And_Time,103)>= convert (date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) And Convert(Date,Date_And_Time,103)<= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)
                            Group By TSPL_BULK_ROUTE_MASTER.ROUTE_NO,tspl_gate_entry_details.Date_And_Time )TankerData
							left outer join
							(Select max(TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date)Document_Date,TSPL_BULK_ROUTE_MASTER.ROUTE_NO,Max(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME)ROUTE_NAME,Max(TSPL_MCC_MASTER.MCC_NAME)MCC_NAME,Max(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader)VLC_Code_VLC_Uploader,Max(Case When IsNull(TSPL_VENDOR_MASTER.Alies_Name,'')<>'' Then TSPL_VENDOR_MASTER.Alies_Name Else TSPL_VENDOR_MASTER.Vendor_Name End) As Alias_Name,Sum(Entered_Qty)Qty,(sum(Entered_FATKg)*100)/sum(Entered_Qty) FAT_Per,
							(sum(Entered_SNFKg)*100)/sum(Entered_Qty) SNF_Per,Sum(Entered_FATKg)Entered_FATKg,Sum(Entered_SNFKg)Entered_SNFKg from TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS 
                            Left Outer Join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code 
                            Left Outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code= TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code
                            Left Outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader=TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader
                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code
							Where Convert(Date,Document_Date,103)>= convert (date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) And Convert(Date,Document_Date,103)<= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)
                            Group By TSPL_BULK_ROUTE_MASTER.ROUTE_NO,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date) As BMCData
							 On TankerData.Document_Date=BMCData.Document_Date And 
							 TankerData.ROUTE_NO=BMCData.ROUTE_NO)xxx"
                    'qry = "Select Convert(varchar(10),xxx.Document_Date,103) As [BMC Date],xxx.ROUTE_NO As [Route No],(xxx.ROUTE_NAME) As [Route Name],(xxx.MCC_NAME) As [MCC Name],
                    '        (xxx.VLC_Code_VLC_Uploader)[VLC Uploader Code],Convert(int,xxx.Qty)Qty,Convert(Decimal(18,1),xxx.FAT_Per) As [FAT(%)],Convert(Decimal(18,1),xxx.SNF_Per) As [SNF(%)],Convert(Decimal(18,3),xxx.Entered_FATKg) As [FAT KG],Convert(Decimal(18,3),xxx.Entered_SNFKg) As [SNF KG],xxx.TankerDocument_Date As [Tanker Date],xxx.TankerTanker_No AS[Tanker No],
                    '        Convert(int,xxx.TankerQty) AS [Tanker Qty],Convert(Decimal(18,1),xxx.TankerFAT_Per)[Tanker Fat(%)],Convert(decimal(18,1),xxx.TankerSNF_Per)[Tanker SNF(%)],Convert(Decimal(18,3),xxx.TankerEntered_FATKg) As [Tanker FAT KG],Convert(Decimal(18,3),xxx.TankerEntered_SNFKg) As [Tanker SNF KG],Convert(int,(xxx.TankerQty-xxx.Qty))[Diff. Qty],
                    '        Convert(Decimal(18,1),(xxx.TankerFAT_Per-xxx.FAT_Per))[Diff. FAT(%)],Convert(Decimal(18,1),(xxx.TankerSNF_Per-xxx.SNF_Per))[Diff. SNF(%)],
                    '        Convert(Decimal(18,3),(xxx.TankerEntered_FATKg-xxx.Entered_FATKg))[Diff. FAT KG],Convert(decimal(18,3),(xxx.TankerEntered_SNFKg-xxx.Entered_SNFKg))[Diff. SNF KG] from 				(Select BMCData.*,TankerData.Document_Date As TankerDocument_Date,TankerData.ROUTE_NO As TankerROUTE_NO,TankerData.Tanker_No As TankerTanker_No,IsNull(TankerData.Qty,0)TankerQty,Isnull(TankerData.FAT_Per,0)TankerFAT_Per,IsNull(TankerData.SNF_Per,0)TankerSNF_Per,IsNull(TankerData.Entered_FATKg,0)TankerEntered_FATKg,IsNull(TankerData.Entered_SNFKg,0)TankerEntered_SNFKg from (Select (TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date)Document_Date,TSPL_BULK_ROUTE_MASTER.ROUTE_NO,Max(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME)ROUTE_NAME,Max(TSPL_MCC_MASTER.MCC_NAME)MCC_NAME,Max(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader)VLC_Code_VLC_Uploader,Sum(Entered_Qty)Qty,Max((Entered_FATKg*100)/Entered_Qty)FAT_Per,Max((Entered_SNFKg*100)/Entered_Qty)SNF_Per,Sum(Entered_FATKg)Entered_FATKg,Sum(Entered_SNFKg)Entered_SNFKg from TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS 
                    '        Left Outer Join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Route_Code 
                    '        Left Outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code= TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.MCC_Code
                    '        Left Outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader=TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader 
                    '        Group By TSPL_BULK_ROUTE_MASTER.ROUTE_NO,TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date) As BMCData
                    '        Left Outer Join(select  (tspl_gate_entry_details.Date_And_Time)Document_Date,TSPL_BULK_ROUTE_MASTER.ROUTE_NO,Max(TSPL_BULK_ROUTE_MASTER.ROUTE_NAME)ROUTE_NAME,
                    '         Max(tspl_gate_entry_details.Tanker_No)Tanker_No,Sum(Qty_In_Kg)Qty,Avg(FAT_Per)FAT_Per,Avg(SNF_Per)SNF_Per, Max(((fat_per*Qty_In_Kg)/100)) As Entered_FATKg,
                    '         Max(((snf_Per*Qty_In_Kg)/100)) As Entered_SNFKg From tspl_gate_entry_details
                    '        Left Outer Join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=tspl_gate_entry_details.ROUTE_NO
                    '        Group By TSPL_BULK_ROUTE_MASTER.ROUTE_NO,tspl_gate_entry_details.Date_And_Time )TankerData On TankerData.Document_Date=BMCData.Document_Date And TankerData.ROUTE_NO=BMCData.ROUTE_NO)xxx
                    'Where Convert(Date,xxx.Document_Date,103)>= convert (date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") + "',103) And Convert(Date,xxx.Document_Date,103)<= convert (date,'" + clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") + "',103)"
                End If
                If rbtnDockDateWise.Checked Then
                    qry = "Select ROW_NUMBER() OVER(ORDER BY (Select 1)) As SNo, SDate , sum(Qty)Qty , sum(Milk_Wtd)Milk_Wtd ,convert(Decimal(18,2),(sum(Fat_KG * 100)/SUM(Milk_Wtd))) As FAT,convert(Decimal(18,2),(sum(SNF_KG * 100)/SUM(Milk_Wtd))) As SNF, sum(Fat_KG)Fat_KG , sum(SNF_KG)SNF_KG , max(Rate)Rate , max(Amt)Amt , max(VLC_Code)VLC_Code , max(DCS_Uploader_Code)DCS_Uploader_Code , max(DCS_Code)DCS_Code , max(DCS_Name)DCS_Name,Max(Alias_Name)Alias_Name , max(BMC_Uploader_Code)BMC_Uploader_Code , max(BMC_Code)BMC_Code , max(BMC_Name)BMC_Name , max(Area)Area , max(Route_Code)Route_Code from (
                       " & qry & " )xx group by SDate order by SDate"
                ElseIf rbtnDockShiftWise.Checked Then
                    qry = "select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SNo, SDate,Shift , sum(Qty)Qty , sum(Milk_Wtd)Milk_Wtd ,convert(decimal(18,2),(sum(Fat_KG * 100)/SUM(Milk_Wtd))) as FAT,convert(decimal(18,2),(sum(SNF_KG * 100)/SUM(Milk_Wtd))) AS SNF, sum(Fat_KG)Fat_KG , sum(SNF_KG)SNF_KG , max(Rate)Rate , max(Amt)Amt , max(VLC_Code)VLC_Code , max(DCS_Uploader_Code)DCS_Uploader_Code , max(DCS_Code)DCS_Code , max(DCS_Name)DCS_Name,MAX(Alias_Name)Alias_Name , max(BMC_Uploader_Code)BMC_Uploader_Code , max(BMC_Code)BMC_Code , max(BMC_Name)BMC_Name , max(Area)Area , max(Route_Code)Route_Code from (
                       " & qry & " )xx group by SDate , Shift order by SDate, shift desc"
                End If
                dt = clsDBFuncationality.GetDataTable(qry)
                Gv1.DataSource = Nothing
                Gv1.Rows.Clear()
                Gv1.Columns.Clear()
                Dim viewBlank As New TableViewDefinition()
                Gv1.ViewDefinition = viewBlank
                Gv1.GroupDescriptors.Clear()
                Gv1.MasterTemplate.SummaryRowsBottom.Clear()
                Gv1.MasterView.Refresh()
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Gv1.DataSource = dt
                    If rbtnDock.Checked OrElse rbtnDockDateWise.Checked OrElse rbtnDockShiftWise.Checked Then
                        SetGridFormationOFGV1Dock()
                    ElseIf rdbMultiple.Checked Then
                        SetGridFormationOFGV1Multiple()
                    ElseIf rdbCollectionWise.Checked Then
                        SetGridFormationOFGV1Collection()
                    Else
                        SetGridFormationOFGV1()
                    End If
                    For ii As Integer = 0 To Gv1.Columns.Count - 1
                        Gv1.Columns(ii).ReadOnly = True
                    Next
                    RadPageView1.SelectedPage = RadPageViewPage2
                    Gv1.BestFitColumns()
                    Gv1.EnableFiltering = True
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    If rdbSummary.Checked Then
                        '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
                        ' DiffEnteredVsMCC_Qty, DiffEnteredVsMCC_FAT,DiffEnteredVsMCC_SNF,DiffEnteredVsMCC_FATKG,DiffEnteredVsMCC_SNFKG
                        Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredQty)
                        Dim EnteredFatKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredFatKg)
                        Dim EnteredSnfKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredSnfKg)
                    End If
                    If Not rbtnBMCTankerCollection.Checked AndAlso Not rbtnRouteWise.Checked AndAlso Not rbtnAliasName.Checked Then
                        Dim MCCQty As New GridViewSummaryItem("MCC_Qty", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(MCCQty)
                        Dim MCCFatKg As New GridViewSummaryItem("MCC_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(MCCFatKg)
                        Dim MCCSnfKg As New GridViewSummaryItem("MCC_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(MCCSnfKg)
                        Dim DCSQty As New GridViewSummaryItem("DCS_Qty", "{0:F0}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DCSQty)
                        Dim DCSFatKg As New GridViewSummaryItem("DCS_FATKG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DCSFatKg)
                        Dim DCSSnfKg As New GridViewSummaryItem("DCS_SNFKG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DCSSnfKg)
                        Dim DiffQty As New GridViewSummaryItem("Diff_Qty", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DiffQty)
                        Dim DiffFatKg As New GridViewSummaryItem("Diff_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DiffFatKg)
                        Dim DiffSnfKg As New GridViewSummaryItem("Diff_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(DiffSnfKg)
                    End If
                    If Not rbtnDock.Checked AndAlso Not rbtnDockShiftWise.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnAliasName.Checked Then
                        If rdbSummary.Checked Then
                            '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
                            ' DiffEnteredVsMCC_Qty, DiffEnteredVsMCC_FAT,DiffEnteredVsMCC_SNF,DiffEnteredVsMCC_FATKG,DiffEnteredVsMCC_SNFKG
                            Dim DiffEnteredVsMCCQty As New GridViewSummaryItem("DiffEnteredVsMCC_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCQty)
                            Dim DiffEnteredVsMCCFatKg As New GridViewSummaryItem("DiffEnteredVsMCC_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCFatKg)
                            Dim DiffEnteredVsMCCSnfKg As New GridViewSummaryItem("DiffEnteredVsMCC_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCSnfKg)
                        End If
                        If rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                            Dim Entered_Qty As New GridViewSummaryItem("Entered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_Qty)
                            Dim Entered_FATKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_FATKg)
                            Dim Entered_SNFKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_SNFKg)
                            Dim DiffMCCVsEnteredQty As New GridViewSummaryItem("DiffMCCVsEntered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffMCCVsEnteredQty)
                            Dim DiffEnteredVsMCCFatKg As New GridViewSummaryItem("DiffMCCVsEntered_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCFatKg)
                            Dim DiffMCCVsEnteredSnfKg As New GridViewSummaryItem("DiffMCCVsEntered_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffMCCVsEnteredSnfKg)
                            Dim FATAmt As New GridViewSummaryItem("FAT_AMT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATAmt)
                            Dim SNFAmt As New GridViewSummaryItem("SNF_AMT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFAmt)
                            Dim FATKG_Recovered As New GridViewSummaryItem("FATKG_Recovered", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATKG_Recovered)
                            Dim SNFKG_Recovered As New GridViewSummaryItem("SNFKG_Recovered", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFKG_Recovered)
                            Dim Amount As New GridViewSummaryItem("AMOUNT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Amount)
                        End If
                        If rdbTankerWise.Checked Then
                            'Dim Tanker_No As New GridViewSummaryItem("Tanker_No", "{0:F2}", "Total")
                            'summaryRowItem.Add(Tanker_No)
                            Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredQty)
                            Dim EnteredFatKg As New GridViewSummaryItem("Entered_FATKg", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredFatKg)
                            Dim EnteredSnfKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredSnfKg)
                            Dim DiffDCSVsEnteredQty As New GridViewSummaryItem("DiffEnteredVsMCC_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffDCSVsEnteredQty)
                            Dim DiffEnteredVsDCSFatKg As New GridViewSummaryItem("DiffEnteredVsMCC_FAT", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsDCSFatKg)
                            Dim DiffEnteredVsDCSSnfKg As New GridViewSummaryItem("DiffEnteredVsMCC_SNF", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsDCSSnfKg)
                        End If
                        'If rbtnRouteWise.Checked Then
                        '    Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F0}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(EnteredQty)
                        '    Dim EnteredFatKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(EnteredFatKg)
                        '    Dim EnteredSnfKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(EnteredSnfKg)
                        '    Dim DiffDCSVsEnteredQty As New GridViewSummaryItem("DiffEnteredVsMCC_Qty", "{0:F0}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(DiffDCSVsEnteredQty)
                        '    Dim DiffEnteredVsDCSFatKg As New GridViewSummaryItem("DiffEnteredVsMCC_FAT", "{0:F2}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(DiffEnteredVsDCSFatKg)
                        '    Dim DiffEnteredVsDCSSnfKg As New GridViewSummaryItem("DiffEnteredVsMCC_SNF", "{0:F2}", GridAggregateFunction.Sum)
                        '    summaryRowItem.Add(DiffEnteredVsDCSSnfKg)
                        'End If
                        If rdbMultiple.Checked Then
                            Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredQty)
                            Dim Entered_FATKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_FATKg)
                            Dim Entered_SNFKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_SNFKg)
                            Dim Qty As New GridViewSummaryItem("Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Qty)
                            Dim FATKG As New GridViewSummaryItem("FATKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATKG)
                            Dim SNFKG As New GridViewSummaryItem("SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFKG)
                            Dim DiffQty1 As New GridViewSummaryItem("Diff_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffQty1)
                            Dim DiffFat1 As New GridViewSummaryItem("Diff_FAT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffFat1)
                            Dim DiffSnf1 As New GridViewSummaryItem("Diff_SNF", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffSnf1)
                        End If
                        If rdbCollectionWise.Checked Then
                            'Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            'summaryRowItem.Add(EnteredQty)
                            'Dim Entered_FATKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                            'summaryRowItem.Add(Entered_FATKg)
                            'Dim Entered_SNFKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                            'summaryRowItem.Add(Entered_SNFKg)
                            Dim Qty As New GridViewSummaryItem("Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Qty)
                            Dim FATKG As New GridViewSummaryItem("FATKG", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATKG)
                            Dim SNFKG As New GridViewSummaryItem("SNFKG", "{0:F3}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFKG)
                        End If
                        If Not rbtnRouteWise.Checked Then
                            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                        End If
                    End If
                    If rbtnBMCTankerCollection.Checked Then
                        Dim Qty As New GridViewSummaryItem("Qty", "{0:F0}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Qty)
                        Dim BMC_FAT As New GridViewSummaryItem("FAT(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(BMC_FAT)
                        Dim BMC_SNF As New GridViewSummaryItem("SNF(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(BMC_SNF)
                        Dim BMC_FATKG As New GridViewSummaryItem("FAT KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(BMC_FATKG)
                        Dim BMC_SNFKG As New GridViewSummaryItem("SNF KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(BMC_SNFKG)
                        Dim Tanker_Qty As New GridViewSummaryItem("Tanker Qty", "{0:F0}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Tanker_Qty)
                        Dim Tanker_FAT As New GridViewSummaryItem("Tanker FAT(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(Tanker_FAT)
                        Dim Tanker_SNF As New GridViewSummaryItem("Tanker SNF(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(Tanker_SNF)
                        Dim Tanker_FATKG As New GridViewSummaryItem("Tanker FAT KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Tanker_FATKG)
                        Dim Tanker_SNFKG As New GridViewSummaryItem("Tanker SNF KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Tanker_SNFKG)
                        Dim Diff_Qty As New GridViewSummaryItem("Diff. Qty", "{0:F0}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Diff_Qty)
                        Dim Diff_FAT As New GridViewSummaryItem("Diff. FAT(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(Diff_FAT)
                        Dim Diff_SNF As New GridViewSummaryItem("Diff. SNF(%)", "{0:F1}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(Diff_SNF)
                        Dim Diff_FATKG As New GridViewSummaryItem("Diff. FAT KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Diff_FATKG)
                        Dim Diff_SNFKG As New GridViewSummaryItem("Diff. SNF KG", "{0:F3}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Diff_SNFKG)
                    End If
                    If rbtnAliasName.Checked Then
                        Dim Qty As New GridViewSummaryItem("Qty", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Qty)
                        Dim MilkWt As New GridViewSummaryItem("Milk Weight", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(MilkWt)
                        Dim FAT As New GridViewSummaryItem("FAT", "{0:F2}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(FAT)
                        Dim SNF As New GridViewSummaryItem("SNF", "{0:F2}", GridAggregateFunction.Avg)
                        summaryRowItem.Add(SNF)
                        Dim FATKG As New GridViewSummaryItem("FAT KG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(FATKG)
                        Dim SNFKG As New GridViewSummaryItem("SNF KG", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(SNFKG)
                        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                    End If
                    ControlEnableDisable(False)
                    '=========================
                    'Gv1.GroupDescriptors.Add(New GridGroupByExpression("[Document_Date] as [Document Date] format ""{0}: {1}"" Group By [Document_Date]"))
                    'Gv1.AutoExpandGroups = True
                    '    Gv1.ShowGroupPanel = False
                    '    Gv1.ShowRowHeaderColumn = False
                    '    Gv1.AllowAddNewRow = False
                    '    Gv1.AllowDeleteRow = False
                    '    Gv1.EnableFiltering = True
                    '    Gv1.ShowFilteringRow = True
                    '==========================
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
#Disable Warning
                    Exit Sub
#Enable Warning
                End If
                ReStoreGridLayout()
                If Not rbtnDock.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnDockShiftWise.Checked AndAlso Not rbtnAliasName.Checked Then
                    View()
                End If
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try
        Else
            Try
                If rbtnTranpoterGainLoss.Checked Then
                    Dim checkRate As String
                    Dim dt1 As New DataTable
                    checkRate = "(SELECT top 1  TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and 
 CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, '" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "', 103))"
                    dt1 = clsDBFuncationality.GetDataTable(checkRate) 'GetDataTable(qry)
                    If dt1 Is Nothing OrElse dt1.Rows.Count <= 0 Then
                        Throw New InvalidOperationException("FAT/SNF RATE NOT FOUND!")
                    End If
                End If
                If rdbSummary.Checked Then
                    PageSetupReport_ID = MyBase.Form_ID & "_S"
                End If
                PageSetupReport_ID = MyBase.Form_ID
                TemplateGridview = Gv1
                'Dim BaseQry As String = ""
                Dim qry As String = ""
                Dim dt As New DataTable
                'Dim whre As String = ""
                Dim strMilkCollectionMCCStatus As String = ""
                Dim strMilkcollectionDCSStatus As String = ""
                If Not rbtnDock.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnDockShiftWise.Checked Then
                    If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                        strMilkCollectionMCCStatus = " and TSPL_MILK_COLLECTION_MCC.Status = 1 "
                        strMilkcollectionDCSStatus = " and TSPL_MILK_COLLECTION_DCS.Status = 1 "
                    ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                        strMilkCollectionMCCStatus = " and TSPL_MILK_COLLECTION_MCC.Status = 0 "
                        strMilkcollectionDCSStatus = " and TSPL_MILK_COLLECTION_DCS.Status = 0 "
                    End If
                End If
                qry = "select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id, Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME, max(Tanker_No) as Tanker_No,max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, max(DCS_Qty) as DCS_Qty, isnull( (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_FAT, isnull( (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_SNF, isnull(max(DCS_FATKG),0) as DCS_FATKG,isnull( max(DCS_SNFKG),0) as DCS_SNFKG, max(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_FAT, (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_SNF, sum(MCC_FATKG) -max(DCS_FATKG) as Diff_FATKG, sum(MCC_SNFKG)-max(DCS_SNFKG) as Diff_SNFKG from ( select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo, TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG from TSPL_MILK_COLLECTION_MCC_DETAIL left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) /nullif( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / nullif(sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as SNF
                         from TSPL_MILK_COLLECTION_DCS_DETAIL 
                                    left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                        left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                        where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) " & strMilkcollectionDCSStatus & "
                        group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                        where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)) xyz group by UploaderNo, Document_No  " & strMilkCollectionMCCStatus & " "
                If rdbSummary.Checked Then
                    qry = " select XXXFinal.Document_No, max(XXXFinal.Document_Date) as Document_Date ,max( XXXFinal.Route_Code) as Route_Code, max(XXXFinal.ROUTE_NAME) as ROUTE_NAME ,
                            max(XXXFinal.Tanker_No ) as Tanker_No, max(XXXFinal.Vehicle_No) as Vehicle_No, max(Entered_Qty) as Entered_Qty,((max (Entered_FATKg) / nullif (max (Entered_Qty),0)) * 100) as Entered_FAT, ((max (Entered_SNFKg) / nullif (max (Entered_Qty),0)) * 100) as Entered_SNF , CAST(ROUND(max(Entered_FATKg), 2) AS DECIMAL(10, 2)) as Entered_FATKg,CAST(ROUND( max(Entered_SNFKg), 2) AS DECIMAL(10, 2)) as Entered_SNFKg , 
                            sum( XXXFinal.MCC_Qty) as  MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_FAT, ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as MCC_SNF,CAST(ROUND(sum (MCC_FATKG), 2) AS DECIMAL(10, 2)) as MCC_FATKG ,CAST(ROUND( sum (MCC_SNFKG), 2) AS DECIMAL(10, 2)) as MCC_SNFKG , sum (DCS_Qty) as DCS_Qty , ROUND(((sum (DCS_FATKG) / nullif( sum (DCS_Qty),0)) * 100),1,0) as DCS_FAT,ROUND(((sum (DCS_SNFKG) / nullif (sum (DCS_Qty),0)) * 100),1,0) as DCS_SNF , CAST(ROUND(sum(DCS_FATKG), 2) AS DECIMAL(10, 2))  as DCS_FATKG ,CAST(ROUND(sum(DCS_SNFKG), 2) AS DECIMAL(10, 2))  as DCS_SNFKG, sum (Diff_Qty) as Diff_Qty ,
                            max(DCS_FAT) - ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_FAT, max(DCS_SNF) - ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as Diff_SNF,CAST(ROUND(sum(Diff_FATKG), 2) AS DECIMAL(10, 2)) as Diff_FATKG, CAST(ROUND( sum(Diff_SNFKG), 2) AS DECIMAL(10, 2)) as Diff_SNFKG,  max(Entered_Qty) - sum( XXXFinal.MCC_Qty)  as DiffEnteredVsMCC_Qty ,
                            ((max (Entered_FATKg) / nullif (max (Entered_Qty),0)) * 100) - ROUND(((sum (MCC_FATKG) / nullif (sum (MCC_Qty),0)) * 100),1,0)  as DiffEnteredVsMCC_FAT ,
                            ((max (Entered_SNFKg) / nullif (max (Entered_Qty),0)) * 100) - ROUND(((sum (MCC_SNFKG) / nullif (sum (MCC_Qty),0)) * 100),1,0) as DiffEnteredVsMCC_SNF ,
                            CAST(ROUND( max(Entered_FATKg) -  sum (MCC_FATKG), 2) AS DECIMAL(10, 2))  as DiffEnteredVsMCC_FATKG , CAST(ROUND(  max(Entered_SNFKg) - sum (MCC_SNFKG), 2) AS DECIMAL(10, 2)) as DiffEnteredVsMCC_SNFKG  from ( " & qry & " ) XXXFinal group by XXXFinal.Document_No order by convert (datetime, max(XXXFinal.Document_Date),103) asc "
                ElseIf rdbDetails.Checked Then
                    qry = "select XXXFinal.PK_Id as PK_Id, XXXFinal.Document_No as Document_No, XXXFinal.Document_Date As Document_Date, XXXFinal.MCC_Code As MCC_Code, XXXFinal.MCC_NAME As MCC_NAME, XXXFinal.UploaderNo As UploaderNo, XXXFinal.Route_Code As Route_Code, XXXFinal.ROUTE_NAME As ROUTE_NAME, XXXFinal.Tanker_No As Tanker_No, XXXFinal.Vehicle_No As Vehicle_No, XXXFinal.MCC_Qty As MCC_Qty, Convert(Decimal(18,2),(MCC_FATKG / nullif(MCC_Qty, 0)) * 100) As MCC_FAT, convert(decimal(18,2),(MCC_SNFKG / nullif(MCC_Qty, 0)) * 100) As MCC_SNF, CAST(ROUND(MCC_FATKG, 2) AS DECIMAL(10, 2))  As MCC_FATKG, CAST(ROUND(MCC_SNFKG, 2) AS DECIMAL(10, 2)) As MCC_SNFKG, DCS_Qty As DCS_Qty,isnull( ROUND(((DCS_FATKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_FAT, isnull(ROUND(((DCS_SNFKG / nullif(DCS_Qty, 0)) * 100), 1, 0),0) As DCS_SNF, CAST(ROUND(isnull(DCS_FATKG,0), 2) AS DECIMAL(10, 2)) As DCS_FATKG,CAST(ROUND(isnull(DCS_SNFKG,0), 2) AS DECIMAL(10, 2))  As DCS_SNFKG, Diff_Qty As Diff_Qty, DCS_FAT - ROUND(((MCC_FATKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_FAT, DCS_SNF - ROUND(((MCC_SNFKG / nullif(MCC_Qty, 0)) * 100), 1, 0) As Diff_SNF, CAST(ROUND(Diff_FATKG, 2) AS DECIMAL(10, 2)) As Diff_FATKG, CAST(ROUND(Diff_SNFKG, 2) AS DECIMAL(10, 2)) As Diff_SNFKG from( " & qry & " ) XXXFinal  
left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No=XXXFinal.Document_No"
                    'qry = qry + " order by TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No asc "
                ElseIf rbtnDock.Checked OrElse rbtnDockDateWise.Checked OrElse rbtnDockShiftWise.Checked Then
                    qry = ""
                    Dim strFilterQry As String = ""
                    Dim strQryStatus1 As String = ""
                    Dim strQryStatus2 As String = ""
                    Dim Area As String = ""
                    If txtMCC_Code.Value.Length > 0 Then
                        strFilterQry = " and TSPL_MCC_MASTER.MCC_Code in ('" & txtMCC_Code.Value & "') "
                    End If
                    Dim AreaName As String = Nothing
                    AreaName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))
                    If clsCommon.myLen(fndArea.Value) > 0 Then
                        Area += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                    End If
                    If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                        strQryStatus1 = "where TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=1"
                        strQryStatus2 = "where TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Status=1"
                    ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                        strQryStatus1 = "where TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=0"
                        strQryStatus2 = "where TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Status=0"
                    End If
                    qry = "Select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SNo,CONVERT(varchar(10), tmp.Shift_Date, 103) As SDate ,tmp.Shift As Shift,No_Of_Cans As Qty,Milk_Weight As Milk_Wtd,FAT as FAT,SNF as SNF,convert(decimal(18,2),((Milk_Weight*FAT)/100)) As Fat_KG,convert(decimal(18,2),((Milk_Weight*SNF)/100)) As SNF_KG,'' As Rate,'' As Amt,tmp.VLC_Code, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As DCS_Uploader_Code,TSPL_VLC_MASTER_HEAD.VLC_Code As DCS_Code,TSPL_VLC_MASTER_HEAD.VLC_Name  As DCS_Name
                        ,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader  As BMC_Uploader_Code,TSPL_MCC_MASTER.MCC_Code As BMC_Code,TSPL_MCC_MASTER.MCC_NAME As BMC_Name,'" & AreaName & "' as Area,TSPL_VLC_MASTER_HEAD.Route_Code
                        from(
                        Select Shift_Date,Shift,No_Of_Cans,Milk_Weight,FAT,SNF,VLC_Code  from TSPL_MILK_SHIFT_UPLOADER_DETAIL
                        Left Outer Join TSPL_MILK_SHIFT_UPLOADER_HEAD ON TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No=TSPL_MILK_SHIFT_UPLOADER_DETAIL.Document_No
	                    " & strQryStatus1 & "
                        Union All
                        Select Shift_Date,Shift,No_Of_Cans,Milk_Weight,FAT,SNF,VLC_Code from TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL
	                    Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD ON TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
	                    " & strQryStatus2 & ")tmp
                        Left Outer Join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=tmp.VLC_Code
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
                        where convert (date,Shift_Date,103) >= convert (date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,Shift_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103)" & strFilterQry
                ElseIf (rbtnTranpoterGainLoss.Checked) Then
                    If txtToleranceFat.Value < 0 OrElse txtToleranceFat.Value >= 100 Then
                        clsCommon.MyMessageBoxShow(Me, "Invalid Input !, Please enter valid Tolerance Fat%.", Me.Text)
                        txtToleranceFat.Focus()
#Disable Warning
                        Exit Sub
#Enable Warning
                    End If
                    If txtToleranceSNF.Value < 0 OrElse txtToleranceSNF.Value >= 100 Then
                        clsCommon.MyMessageBoxShow(Me, "Invalid Input !, Please enter valid Tolerance Fat%.", Me.Text)
                        txtToleranceSNF.Focus()
#Disable Warning
                        Exit Sub
#Enable Warning
                    End If
                    qry = "select  XXGetAllRecords.Document_No,XXGetAllRecords.Document_Date,XXGetAllRecords.Route_Code,XXGetAllRecords.ROUTE_NAME,XXGetAllRecords.Tanker_No,XXGetAllRecords.Vehicle_No,
XXGetAllRecords.Entered_Qty,CAST(ROUND( XXGetAllRecords.Entered_FATKg, 2) AS DECIMAL(10, 2)) as Entered_FATKg ,CAST(ROUND( XXGetAllRecords.Entered_SNFKg, 2) AS DECIMAL(10, 2))as Entered_SNFKg,
XXGetAllRecords.MCC_Qty,CAST(ROUND( XXGetAllRecords.MCC_FATKG, 2) AS DECIMAL(10, 2)) as MCC_FATKG,CAST(ROUND( XXGetAllRecords.MCC_SNFKG, 2) AS DECIMAL(10, 2)) as MCC_SNFKG,
XXGetAllRecords.DiffMCCVsEntered_Qty,
CAST(ROUND( XXGetAllRecords.DiffMCCVsEntered_FATKG, 2) AS DECIMAL(10, 2))
as DiffMCCVsEntered_FATKG," & clsCommon.myCstr(txtToleranceFat.Value) & " as FAT_Tolerence,CAST(ROUND( case when (isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ",0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ",0)  end , 2) AS DECIMAL(10, 2)) as FATKG_Recovered,Round(XXGetAllRecords.FAT_AMT,2,0) as FAT_AMT,
CAST(ROUND( XXGetAllRecords.DiffMCCVsEntered_SNFKG, 2) AS DECIMAL(10, 2))as DiffMCCVsEntered_SNFKG," & clsCommon.myCstr(txtToleranceSNF.Value) & " as SNF_Tolerence,CAST(ROUND( case when (isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ",0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ",0) end, 2) AS DECIMAL(10, 2))  as SNFKG_Recovered,Round(XXGetAllRecords.SNF_AMT,2,0) as SNF_AMT,Round((XXGetAllRecords.FAT_AMT + XXGetAllRecords.SNF_AMT),2,0) as AMOUNT,XXGetAllRecords.GainLossCode as GainLoss_Code,
   XXGetAllRecords.Loss_FAT_Rate as Loss_FAT_Rate,
   XXGetAllRecords.Loss_SNF_Rate as Loss_SNF_Rate,
   XXGetAllRecords.Start_Date as Start_Date
 from( select GetAllGainLossRate.*,
              case when GetAllGainLossRate.DiffMCCVsEntered_FATKG >= 0 then (
        case when (GetAllGainLossRate.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_FATKG -" & clsCommon.myCstr(txtToleranceFat.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate) else 0 end)
        else 0 end as FAT_AMT,
        case when GetAllGainLossRate.DiffMCCVsEntered_SNFKG > 0 then (case when (GetAllGainLossRate.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtToleranceSNF.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_SNFKG -" & clsCommon.myCstr(txtToleranceSNF.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate) else 0 end)
      else 0 end as SNF_AMT,
      TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date from ( select GetRateCode.*,( SELECT top 1 TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, GetRateCode.Document_Date, 103) order by TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date desc ) as GainLossCode from ( SELECT XXXFinal.Document_No, MAX(XXXFinal.Document_Date) AS Document_Date, MAX(XXXFinal.Route_Code) AS Route_Code, MAX(XXXFinal.ROUTE_NAME) AS ROUTE_NAME, MAX(XXXFinal.Tanker_No) AS Tanker_No, MAX(XXXFinal.Vehicle_No) AS Vehicle_No, MAX(Entered_Qty) AS Entered_Qty, ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_FAT, ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_SNF, MAX(Entered_FATKg) AS Entered_FATKg, MAX(Entered_SNFKg) AS Entered_SNFKg, SUM(XXXFinal.MCC_Qty) AS MCC_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_SNF, SUM(MCC_FATKG) AS MCC_FATKG, SUM(MCC_SNFKG) AS MCC_SNFKG, SUM(DCS_Qty) AS DCS_Qty, ROUND(((SUM(DCS_FATKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_FAT, ROUND(((SUM(DCS_SNFKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_SNF, SUM(DCS_FATKG) AS DCS_FATKG, SUM(DCS_SNFKG) AS DCS_SNFKG, SUM(Diff_Qty) AS Diff_Qty, MAX(DCS_FAT) - ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_FAT, MAX(DCS_SNF) - ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_SNF, SUM(Diff_FATKG) AS Diff_FATKG, SUM(Diff_SNFKG) AS Diff_SNFKG, SUM(XXXFinal.MCC_Qty)- MAX(Entered_Qty) AS DiffMCCVsEntered_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_SNF, SUM(MCC_FATKG)-MAX(Entered_FATKg) AS DiffMCCVsEntered_FATKG, SUM(MCC_SNFKG)-MAX(Entered_SNFKg) AS DiffMCCVsEntered_SNFKG from ( " & qry & " ) XXXFinal group by XXXFinal.Document_No )GetRateCode )GetAllGainLossRate left outer join TSPL_OWN_BMC_GAIN_LOSS_RATE on TSPL_OWN_BMC_GAIN_LOSS_RATE.Code=GetAllGainLossRate.GainLossCode ) XXGetAllRecords"
                    If clsCommon.myLen(TxtTankerNo.Value) > 0 Then
                        qry = qry & " where tanker_no ='" & TxtTankerNo.Value & "'"
                    End If
                ElseIf rdbTankerWise.Checked Then
                    qry = "SELECT  Max(XXXFinal.tanker_no) AS Tanker_No, Sum (MCC_Qty) AS MCC_Qty, CAST(ROUND( Sum (MCC_FATKG), 2) AS DECIMAL(10, 2))  AS MCC_FATKG,  CAST(ROUND( Sum (MCC_SNFKG), 2) AS DECIMAL(10, 2))  AS MCC_SNFKG, 
              sum(entered_qty) AS Entered_Qty, CAST(ROUND( Sum (entered_fatkg), 2) AS DECIMAL(10, 2))   AS Entered_FATKg, CAST(ROUND( Sum (entered_snfkg), 2) AS DECIMAL(10, 2))  AS Entered_SNFKg, Sum(entered_qty) - Sum(MCC_Qty) AS DiffEnteredVsMCC_Qty, Sum (entered_fatkg) - Sum (MCC_FATKG)  AS
              DiffEnteredVsMCC_FAT, Sum (entered_snfkg)- Sum (MCC_SNFKG) AS DiffEnteredVsMCC_SNF FROM   (SELECT Max(entered_qty) AS Entered_Qty, Max(entered_fatkg) AS Entered_FATKg, Max(entered_snfkg) AS Entered_SNFKg, Max(document_date) AS Document_Date, document_no, Max(tanker_no)
              AS Tanker_No, sum(MCC_Qty) AS MCC_Qty, Isnull(Sum(MCC_FATKG), 0) AS MCC_FATKG, Isnull(Sum(MCC_SNFKG), 0) AS MCC_SNFKG
              FROM   (SELECT tspl_milk_collection_mcc.entered_qty, tspl_milk_collection_mcc.entered_fatkg, tspl_milk_collection_mcc.entered_snfkg,
              tspl_milk_collection_mcc_detail.document_no, tspl_milk_collection_mcc.tanker_no, CONVERT (VARCHAR, tspl_milk_collection_mcc.document_date,103)
              AS Document_Date, Isnull(qty, 0) AS MCC_Qty, Isnull(fat, 0) AS MCC_FAT, ISNULL(SNF, 0) AS MCC_SNF, ISNULL(FATKG, 0) AS MCC_FATKG, 
              ISNULL(SNFKG, 0) AS MCC_SNFKG  FROM   tspl_milk_collection_mcc_detail
              LEFT OUTER JOIN tspl_milk_collection_mcc ON tspl_milk_collection_mcc.document_no = tspl_milk_collection_mcc_detail.document_no 
              LEFT OUTER JOIN tspl_mcc_master ON tspl_mcc_master.mcc_code = tspl_milk_collection_mcc_detail.mcc_code 
              LEFT OUTER JOIN tspl_bulk_route_master ON tspl_bulk_route_master.route_no = tspl_milk_collection_mcc.route_code
              WHERE  CONVERT(DATE, tspl_milk_collection_mcc.document_date, 103) >= CONVERT(DATE, '" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) AND CONVERT (DATE, tspl_milk_collection_mcc.document_date, 103)
             <= CONVERT (DATE, '" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) ) xyz GROUP  BY  document_no) XXXFinal GROUP  BY  Tanker_No
             ORDER  BY Tanker_No asc "
                End If
                If rbtnDockDateWise.Checked Then
                    qry = "Select ROW_NUMBER() OVER(ORDER BY (Select 1)) As SNo, SDate , sum(Qty)Qty , sum(Milk_Wtd)Milk_Wtd ,convert(Decimal(18,2),(sum(Fat_KG * 100)/SUM(Milk_Wtd))) As FAT,convert(Decimal(18,2),(sum(SNF_KG * 100)/SUM(Milk_Wtd))) As SNF, sum(Fat_KG)Fat_KG , sum(SNF_KG)SNF_KG , max(Rate)Rate , max(Amt)Amt , max(VLC_Code)VLC_Code , max(DCS_Uploader_Code)DCS_Uploader_Code , max(DCS_Code)DCS_Code , max(DCS_Name)DCS_Name , max(BMC_Uploader_Code)BMC_Uploader_Code , max(BMC_Code)BMC_Code , max(BMC_Name)BMC_Name , max(Area)Area , max(Route_Code)Route_Code from (
                       " & qry & " )xx group by SDate order by SDate"
                ElseIf rbtnDockShiftWise.Checked Then
                    qry = "select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS SNo, SDate,Shift , sum(Qty)Qty , sum(Milk_Wtd)Milk_Wtd ,convert(decimal(18,2),(sum(Fat_KG * 100)/SUM(Milk_Wtd))) as FAT,convert(decimal(18,2),(sum(SNF_KG * 100)/SUM(Milk_Wtd))) AS SNF, sum(Fat_KG)Fat_KG , sum(SNF_KG)SNF_KG , max(Rate)Rate , max(Amt)Amt , max(VLC_Code)VLC_Code , max(DCS_Uploader_Code)DCS_Uploader_Code , max(DCS_Code)DCS_Code , max(DCS_Name)DCS_Name , max(BMC_Uploader_Code)BMC_Uploader_Code , max(BMC_Code)BMC_Code , max(BMC_Name)BMC_Name , max(Area)Area , max(Route_Code)Route_Code from (
                       " & qry & " )xx group by SDate , Shift order by SDate, shift desc"
                End If
                dt = clsDBFuncationality.GetDataTable(qry)
                Gv1.DataSource = Nothing
                Gv1.Rows.Clear()
                Gv1.Columns.Clear()
                Dim viewBlank As New TableViewDefinition()
                Gv1.ViewDefinition = viewBlank
                Gv1.GroupDescriptors.Clear()
                Gv1.MasterTemplate.SummaryRowsBottom.Clear()
                Gv1.MasterView.Refresh()
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Gv1.DataSource = dt
                    If rbtnDock.Checked OrElse rbtnDockDateWise.Checked OrElse rbtnDockShiftWise.Checked Then
                        SetGridFormationOFGV1Dock()
                    Else
                        SetGridFormationOFGV1()
                    End If
                    For ii As Integer = 0 To Gv1.Columns.Count - 1
                        Gv1.Columns(ii).ReadOnly = True
                    Next
                    RadPageView1.SelectedPage = RadPageViewPage2
                    Gv1.BestFitColumns()
                    Gv1.EnableFiltering = True
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    If rdbSummary.Checked Then
                        Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredQty)
                        Dim EnteredFatKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredFatKg)
                        Dim EnteredSnfKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(EnteredSnfKg)
                    End If
                    Dim MCCQty As New GridViewSummaryItem("MCC_Qty", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(MCCQty)
                    Dim MCCFatKg As New GridViewSummaryItem("MCC_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(MCCFatKg)
                    Dim MCCSnfKg As New GridViewSummaryItem("MCC_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(MCCSnfKg)
                    Dim DCSQty As New GridViewSummaryItem("DCS_Qty", "{0:F0}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DCSQty)
                    Dim DCSFatKg As New GridViewSummaryItem("DCS_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DCSFatKg)
                    Dim DCSSnfKg As New GridViewSummaryItem("DCS_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DCSSnfKg)
                    Dim DiffQty As New GridViewSummaryItem("Diff_Qty", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DiffQty)
                    Dim DiffFatKg As New GridViewSummaryItem("Diff_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DiffFatKg)
                    Dim DiffSnfKg As New GridViewSummaryItem("Diff_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(DiffSnfKg)
                    If Not rbtnDock.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnDockShiftWise.Checked Then
                        If rdbSummary.Checked Then
                            '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
                            ' DiffEnteredVsMCC_Qty, DiffEnteredVsMCC_FAT,DiffEnteredVsMCC_SNF,DiffEnteredVsMCC_FATKG,DiffEnteredVsMCC_SNFKG
                            Dim DiffEnteredVsMCCQty As New GridViewSummaryItem("DiffEnteredVsMCC_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCQty)
                            Dim DiffEnteredVsMCCFatKg As New GridViewSummaryItem("DiffEnteredVsMCC_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCFatKg)
                            Dim DiffEnteredVsMCCSnfKg As New GridViewSummaryItem("DiffEnteredVsMCC_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCSnfKg)
                        End If
                        If rbtnTranpoterGainLoss.Checked Then
                            Dim Entered_Qty As New GridViewSummaryItem("Entered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_Qty)
                            Dim Entered_FATKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_FATKg)
                            Dim Entered_SNFKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Entered_SNFKg)
                            Dim DiffMCCVsEnteredQty As New GridViewSummaryItem("DiffMCCVsEntered_Qty", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffMCCVsEnteredQty)
                            Dim DiffEnteredVsMCCFatKg As New GridViewSummaryItem("DiffMCCVsEntered_FATKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsMCCFatKg)
                            Dim DiffMCCVsEnteredSnfKg As New GridViewSummaryItem("DiffMCCVsEntered_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffMCCVsEnteredSnfKg)
                            Dim FATAmt As New GridViewSummaryItem("FAT_AMT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATAmt)
                            Dim SNFAmt As New GridViewSummaryItem("SNF_AMT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFAmt)
                            Dim FATKG_Recovered As New GridViewSummaryItem("FATKG_Recovered", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(FATKG_Recovered)
                            Dim SNFKG_Recovered As New GridViewSummaryItem("SNFKG_Recovered", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(SNFKG_Recovered)
                            Dim Amount As New GridViewSummaryItem("AMOUNT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(Amount)
                        End If
                        If rdbTankerWise.Checked Then
                            'Dim Tanker_No As New GridViewSummaryItem("Tanker_No", "{0:F2}", "Total")
                            'summaryRowItem.Add(Tanker_No)
                            Dim EnteredQty As New GridViewSummaryItem("Entered_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredQty)
                            Dim EnteredFatKg As New GridViewSummaryItem("Entered_FATKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredFatKg)
                            Dim EnteredSnfKg As New GridViewSummaryItem("Entered_SNFKg", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(EnteredSnfKg)
                            Dim DiffDCSVsEnteredQty As New GridViewSummaryItem("DiffEnteredVsMCC_Qty", "{0:F0}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffDCSVsEnteredQty)
                            Dim DiffEnteredVsDCSFatKg As New GridViewSummaryItem("DiffEnteredVsMCC_FAT", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsDCSFatKg)
                            Dim DiffEnteredVsDCSSnfKg As New GridViewSummaryItem("DiffEnteredVsMCC_SNF", "{0:F2}", GridAggregateFunction.Sum)
                            summaryRowItem.Add(DiffEnteredVsDCSSnfKg)
                        End If
                        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                    End If
                    ControlEnableDisable(False)
                    '=========================
                    'Gv1.GroupDescriptors.Add(New GridGroupByExpression("[Document_Date] as [Document Date] format ""{0}: {1}"" Group By [Document_Date]"))
                    'Gv1.AutoExpandGroups = True
                    '    Gv1.ShowGroupPanel = False
                    '    Gv1.ShowRowHeaderColumn = False
                    '    Gv1.AllowAddNewRow = False
                    '    Gv1.AllowDeleteRow = False
                    '    Gv1.EnableFiltering = True
                    '    Gv1.ShowFilteringRow = True
                    '==========================
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
#Disable Warning
                    Exit Sub
#Enable Warning
                End If
                ReStoreGridLayout()
                If Not rbtnDock.Checked AndAlso Not rbtnDockDateWise.Checked AndAlso Not rbtnDockShiftWise.Checked Then
                    View()
                End If
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try
        End If
    End Sub

    Sub GetReportID()
        Dim VarID As String = ""
        If rdbSummary.Checked Then
            VarID += "_S"
        ElseIf rdbDetails.Checked Then
            VarID += "_D"
        ElseIf rbtnDock.Checked Then
            VarID += "_D"
        ElseIf rbtnTranpoterGainLoss.Checked Then
            VarID += "_TGL"
        ElseIf rbtnBmcSummary.Checked Then
            VarID += "_BS"
        ElseIf rdbTankerWise.Checked Then
            VarID += "_TW"
        ElseIf rdbMultiple.Checked Then
            VarID += "_M"
        ElseIf rdbCollectionWise.Checked Then
            VarID += "_CW"
        ElseIf rbtnBMCTankerCollection.Checked Then
            VarID += "_BTC"
        ElseIf rbtnRouteWise.Checked Then
            VarID += "_RW"
        ElseIf rbtnDockDateWise.Checked Then
            VarID += "_DDW"
        ElseIf rbtnDockShiftWise.Checked Then
            VarID += "_DSW"
        ElseIf rbtnTranspoterGainlossSummary.Checked Then
            VarID += "_TGLS"
        End If

        If rbtnBMCDock.Checked Then
            VarID += "_BD"
        ElseIf rbtnDockSummary.Checked Then
            VarID += "_DS"
        ElseIf rbtnSummary.Checked Then
            VarID += "_S"
        ElseIf rbtnBMCDCSPrint.Checked Then
            VarID += "_BMP"
        ElseIf rbtnBMCRoutePrint.Checked Then
            VarID += "_BRP"
        End If
        Gv1.VarID = VarID
    End Sub

    Sub SetGridFormationOFGV1()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            'PK_Id,Document_No,MCC_Code,MCC_NAME,
            'UploaderNo,Route_Code,ROUTE_NAME,Tanker_No,Vehicle_No,MCC_Qty,MCC_FAT,MCC_FATKG,MCC_SNF,MCC_SNFKG,
            'DCS_Qty,DCS_FAT,DCS_SNF,DCS_FATKG,DCS_SNF_KG
            If rdbDetails.Checked Then
                Gv1.Columns("PK_Id").HeaderText = "PK Id"
                Gv1.Columns("PK_Id").IsVisible = False
            End If
            If rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                Gv1.Columns("Document_No").IsVisible = False
                Gv1.Columns("Vehicle_No").IsVisible = False
                Gv1.Columns("GainLoss_Code").HeaderText = "Gain Loss Code"
                Gv1.Columns("GainLoss_Code").IsVisible = False
                Gv1.Columns("Loss_FAT_Rate").HeaderText = "Loss FAT Rate"
                Gv1.Columns("Loss_FAT_Rate").IsVisible = False
                Gv1.Columns("Loss_SNF_Rate").HeaderText = "Loss SNF Rate"
                Gv1.Columns("Loss_SNF_Rate").IsVisible = False
                Gv1.Columns("Start_Date").HeaderText = "Start Date"
                Gv1.Columns("Start_Date").IsVisible = False
                Gv1.Columns("Reject_Types").HeaderText = "Reject Type"
                Gv1.Columns("Reject_Types").IsVisible = False
                Gv1.Columns("Reject_Recovery_Pers").HeaderText = "Reject Recovery"
                Gv1.Columns("Reject_Recovery_Pers").IsVisible = False
            End If
            If Not rdbTankerWise.Checked AndAlso Not rbtnBMCTankerCollection.Checked AndAlso Not rbtnRouteWise.Checked AndAlso Not rbtnAliasName.Checked Then
                Gv1.Columns("Document_No").HeaderText = "Document No"
                Gv1.Columns("Route_Code").HeaderText = "Route Code"
                Gv1.Columns("ROUTE_NAME").HeaderText = "Route Name"
                Gv1.Columns("Document_Date").HeaderText = "Document Date"
                Gv1.Columns("Vehicle_No").HeaderText = "Vehicle No"
            End If
            If Not rbtnBMCTankerCollection.Checked AndAlso Not rbtnRouteWise.Checked AndAlso Not rbtnAliasName.Checked Then
                Gv1.Columns("Tanker_No").HeaderText = "Tanker No"
                ' Gv1.Columns("Reject_Types").HeaderText = "Reject Type"
                ' Gv1.Columns("Reject_Recovery_Pers").HeaderText = "Reject Recovery"
            End If
            If rdbDetails.Checked Then
                Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
                Gv1.Columns("MCC_NAME").HeaderText = "MCC Name"
                Gv1.Columns("UploaderNo").HeaderText = "UploaderNo"
            End If
            If rbtnBmcSummary.Checked Then
                Gv1.Columns("Tanker_No").HeaderText = "Vehicle No"
                Gv1.Columns("Tanker_No").IsVisible = False
                Gv1.Columns("Vehicle_No").HeaderText = "Vehicle No"
                Gv1.Columns("Vehicle_No").IsVisible = False
                Gv1.Columns("Route_Code").HeaderText = "Route Code"
                Gv1.Columns("Route_Code").IsVisible = False
                Gv1.Columns("ROUTE_NAME").HeaderText = "Route Name"
                Gv1.Columns("ROUTE_NAME").IsVisible = False
                Gv1.Columns("Document_No").HeaderText = "Document No"
                Gv1.Columns("Document_No").IsVisible = False
                Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
                Gv1.Columns("MCC_Code").IsVisible = True
                Gv1.Columns("MCC_NAME").HeaderText = "MCC Name"
                Gv1.Columns("MCC_NAME").IsVisible = True
                Gv1.Columns("UploaderNo").HeaderText = "UploaderNo"
                Gv1.Columns("UploaderNo").IsVisible = True
                Gv1.Columns("DCSCode").HeaderText = "DCS Code"
                Gv1.Columns("DCSCode").IsVisible = True
                Gv1.Columns("Alias_Name").HeaderText = "Alias Name"
                Gv1.Columns("Alias_Name").IsVisible = False
                Gv1.Columns("Alias_Name").VisibleInColumnChooser = True
            End If
            If rdbSummary.Checked Then
                '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
                Gv1.Columns("Entered_Qty").HeaderText = "Qty"
                Gv1.Columns("Entered_Qty").FormatString = "{0:n3}"
                Gv1.Columns("Entered_FAT").HeaderText = "FAT %"
                Gv1.Columns("Entered_FAT").FormatString = "{0:n2}"
                Gv1.Columns("Entered_SNF").HeaderText = "SNF %"
                Gv1.Columns("Entered_SNF").FormatString = "{0:n2}"
                Gv1.Columns("Entered_FATKg").HeaderText = "FAT KG"
                Gv1.Columns("Entered_SNFKg").HeaderText = "SNF KG"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n3}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n2}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n2}"
                End If
            ElseIf rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                Gv1.Columns("Entered_Qty").HeaderText = "Qty"
                Gv1.Columns("Entered_Qty").FormatString = "{0:n3}"
                Gv1.Columns("Entered_FATKg").HeaderText = "FAT KG"
                Gv1.Columns("Entered_SNFKg").HeaderText = "SNF KG"
                Gv1.Columns("MCC_FATKG").HeaderText = "FAT KG"
                Gv1.Columns("MCC_SNFKG").HeaderText = "SNF KG"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n3}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n3}"
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n2}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n2}"
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n2}"
                End If
                Gv1.Columns("MCC_Qty").HeaderText = "Qty"
                Gv1.Columns("MCC_Qty").FormatString = "{0:n3}"
                'Gv1.Columns("MCC_FAT").HeaderText = "FAT %"
                'Gv1.Columns("MCC_FAT").FormatString = "{0:n2}"
                'Gv1.Columns("MCC_SNF").HeaderText = "SNF %"
                'Gv1.Columns("MCC_SNF").FormatString = "{0:n2}"
            ElseIf (rbtnDock.Checked) Then
                Gv1.Columns("Entered_Qty").HeaderText = "Qty"
                'Gv1.Columns("Entered_FAT").HeaderText = "FAT %"
                'Gv1.Columns("Entered_SNF").HeaderText = "SNF %"
                Gv1.Columns("Entered_FATKg").HeaderText = "FAT KG"
                Gv1.Columns("Entered_SNFKg").HeaderText = "SNF KG" 'IsVisible = False
                Gv1.Columns("Entered_Qty").IsVisible = False
                'Gv1.Columns("Entered_FAT").IsVisible = False
                'Gv1.Columns("Entered_SNF").IsVisible = False
                Gv1.Columns("Entered_FATKg").IsVisible = False
                Gv1.Columns("Entered_SNFKg").IsVisible = False
            ElseIf rdbTankerWise.Checked Then
                Gv1.Columns("Tanker_No").HeaderText = "Tankerno"
                Gv1.Columns("MCC_Qty").HeaderText = "Quantity"
                Gv1.Columns("MCC_Qty").FormatString = "{0:n0}"
                Gv1.Columns("MCC_FATKG").HeaderText = "KGFAT"
                Gv1.Columns("MCC_SNFKG").HeaderText = "KGSNF"
                Gv1.Columns("Entered_Qty").HeaderText = "Quantity"
                Gv1.Columns("Entered_Qty").FormatString = "{0:n0}"
                Gv1.Columns("Entered_FATKg").HeaderText = "KGFAT"
                Gv1.Columns("Entered_SNFKg").HeaderText = "KGSNF"
                Gv1.Columns("DiffEnteredVsMCC_Qty").HeaderText = "Quantity"
                Gv1.Columns("DiffEnteredVsMCC_Qty").FormatString = "{0:n0}"
                Gv1.Columns("DiffEnteredVsMCC_FAT").HeaderText = "KGFAT"
                Gv1.Columns("DiffEnteredVsMCC_SNF").HeaderText = "KGSNF"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n3}"
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n3}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n3}"
                    Gv1.Columns("DiffEnteredVsMCC_FAT").FormatString = "{0:n3}"
                    Gv1.Columns("DiffEnteredVsMCC_SNF").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n2}"
                    Gv1.Columns("Entered_FATKg").FormatString = "{0:n2}"
                    Gv1.Columns("Entered_SNFKg").FormatString = "{0:n2}"
                    Gv1.Columns("DiffEnteredVsMCC_FAT").FormatString = "{0:n2}"
                    Gv1.Columns("DiffEnteredVsMCC_SNF").FormatString = "{0:n2}"
                End If
            End If
            If Not rbtnTranpoterGainLoss.Checked AndAlso Not rbtnTranspoterGainlossSummary.Checked AndAlso Not rdbTankerWise.Checked AndAlso Not rbtnBMCTankerCollection.Checked AndAlso Not rbtnRouteWise.Checked AndAlso Not rbtnAliasName.Checked Then
                Gv1.Columns("MCC_Qty").HeaderText = "Qty"
                Gv1.Columns("MCC_Qty").FormatString = "{0:n3}"
                Gv1.Columns("MCC_FAT").HeaderText = "FAT %"
                Gv1.Columns("MCC_FAT").FormatString = "{0:n2}"
                Gv1.Columns("MCC_SNF").HeaderText = "SNF %"
                Gv1.Columns("MCC_SNF").FormatString = "{0:n2}"
                Gv1.Columns("MCC_FATKG").HeaderText = "FAT KG"
                Gv1.Columns("MCC_SNFKG").HeaderText = "SNF KG"
                Gv1.Columns("DCS_FATKG").HeaderText = "FAT KG"
                Gv1.Columns("DCS_SNFKG").HeaderText = "SNF KG"
                Gv1.Columns("Diff_FATKG").HeaderText = "FAT KG"
                Gv1.Columns("Diff_SNFKG").HeaderText = "SNF KG"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n3}"
                    Gv1.Columns("DCS_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("DCS_SNFKG").FormatString = "{0:n3}"
                    Gv1.Columns("Diff_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("Diff_SNFKG").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("MCC_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("MCC_SNFKG").FormatString = "{0:n2}"
                    Gv1.Columns("DCS_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("DCS_SNFKG").FormatString = "{0:n2}"
                    Gv1.Columns("Diff_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("Diff_SNFKG").FormatString = "{0:n2}"
                End If
                Gv1.Columns("DCS_Qty").HeaderText = "Qty"
                Gv1.Columns("DCS_Qty").FormatString = "{0:n3}"
                Gv1.Columns("DCS_FAT").HeaderText = "FAT %"
                Gv1.Columns("DCS_FAT").FormatString = "{0:n2}"
                Gv1.Columns("DCS_SNF").HeaderText = "SNF %"
                Gv1.Columns("DCS_SNF").FormatString = "{0:n2}"
                'Diff
                Gv1.Columns("Diff_Qty").HeaderText = "Qty"
                Gv1.Columns("Diff_Qty").FormatString = "{0:n3}"
                Gv1.Columns("Diff_FAT").HeaderText = "FAT %"
                Gv1.Columns("Diff_FAT").FormatString = "{0:n2}"
                Gv1.Columns("Diff_SNF").HeaderText = "SNF %"
                Gv1.Columns("Diff_SNF").FormatString = "{0:n2}"
            End If
            If rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                Gv1.Columns("DiffMCCVsEntered_Qty").HeaderText = "Qty"
                Gv1.Columns("DiffMCCVsEntered_Qty").FormatString = "{0:n3}"
                Gv1.Columns("DiffMCCVsEntered_FATKG").HeaderText = "FAT Difference"
                Gv1.Columns("FATKG_Recovered").HeaderText = "FAT KG to be Recovered"
                Gv1.Columns("DiffMCCVsEntered_SNFKG").HeaderText = "SNF Difference"
                Gv1.Columns("SNFKG_Recovered").HeaderText = "SNF KG to be Recovered"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("DiffMCCVsEntered_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("FATKG_Recovered").FormatString = "{0:n3}"
                    Gv1.Columns("DiffMCCVsEntered_SNFKG").FormatString = "{0:n3}"
                    Gv1.Columns("SNFKG_Recovered").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("DiffMCCVsEntered_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("FATKG_Recovered").FormatString = "{0:n2}"
                    Gv1.Columns("DiffMCCVsEntered_SNFKG").FormatString = "{0:n2}"
                    Gv1.Columns("SNFKG_Recovered").FormatString = "{0:n2}"
                End If
                Gv1.Columns("FAT_Tolerence").HeaderText = "Allow FAT Tolerence"
                Gv1.Columns("FAT_Tolerence").FormatString = "{0:n3}"
                Gv1.Columns("FAT_AMT").HeaderText = "FAT Amount Recovered"
                Gv1.Columns("FAT_AMT").FormatString = "{0:n2}"
                Gv1.Columns("SNF_Tolerence").HeaderText = "Allow SNF Tolerence"
                Gv1.Columns("SNF_Tolerence").FormatString = "{0:n3}"
                Gv1.Columns("SNF_AMT").HeaderText = "SNF Amount Recovered"
                Gv1.Columns("SNF_AMT").FormatString = "{0:n2}"
                Gv1.Columns("AMOUNT").HeaderText = "AMOUNT"
                Gv1.Columns("AMOUNT").FormatString = "{0:n2}"
            End If
            If rdbSummary.Checked Then
                Gv1.Columns("DiffEnteredVsMCC_Qty").HeaderText = "Qty"
                Gv1.Columns("DiffEnteredVsMCC_Qty").FormatString = "{0:n3}"
                Gv1.Columns("DiffEnteredVsMCC_FAT").HeaderText = "FAT %"
                Gv1.Columns("DiffEnteredVsMCC_FAT").FormatString = "{0:n2}"
                Gv1.Columns("DiffEnteredVsMCC_SNF").HeaderText = "SNF %"
                Gv1.Columns("DiffEnteredVsMCC_SNF").FormatString = "{0:n2}"
                Gv1.Columns("DiffEnteredVsMCC_FATKG").HeaderText = "FAT KG"
                Gv1.Columns("DiffEnteredVsMCC_SNFKG").HeaderText = "SNF KG"
                If FatSnfRoundOff = 0 Then
                    Gv1.Columns("DiffEnteredVsMCC_FATKG").FormatString = "{0:n3}"
                    Gv1.Columns("DiffEnteredVsMCC_SNFKG").FormatString = "{0:n3}"
                Else
                    Gv1.Columns("DiffEnteredVsMCC_FATKG").FormatString = "{0:n2}"
                    Gv1.Columns("DiffEnteredVsMCC_SNFKG").FormatString = "{0:n2}"
                End If
            End If
            '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
            ' DiffEnteredVsMCC_Qty, DiffEnteredVsMCC_FAT,DiffEnteredVsMCC_SNF,DiffEnteredVsMCC_FATKG,DiffEnteredVsMCC_SNFKG
            If rbtnRouteWise.Checked Then
                Gv1.Columns("ROUTE_NO").HeaderText = "Route Code"
                Gv1.Columns("S.No.").IsVisible = False
                Gv1.Columns("ROUTE_NAME").HeaderText = "Route Name"
                Gv1.Columns("MCC_Qty").HeaderText = "Quantity"
                Gv1.Columns("MCC_Qty").FormatString = "{0:n0}"
                Gv1.Columns("MCC_FAT_Per").HeaderText = "%FAT"
                Gv1.Columns("MCC_SNF_Per").HeaderText = "%SNF"
                Gv1.Columns("MCC_FATKG").HeaderText = "KGFAT"
                Gv1.Columns("MCC_FATKG").FormatString = "{0:n2}"
                Gv1.Columns("MCC_SNFKG").HeaderText = "KGSNF"
                Gv1.Columns("MCC_SNFKG").FormatString = "{0:n2}"
                Gv1.Columns("Entered_Qty").HeaderText = "Quantity"
                Gv1.Columns("Entered_Qty").FormatString = "{0:n0}"
                Gv1.Columns("Entered_FAT_Per").HeaderText = "%FAT"
                Gv1.Columns("Entered_SNF_Per").HeaderText = "%SNF"
                Gv1.Columns("Entered_FATKg").HeaderText = "KGFAT"
                Gv1.Columns("Entered_SNFKg").HeaderText = "KGSNF"
                Gv1.Columns("Entered_FATKg").FormatString = "{0:n2}"
                Gv1.Columns("Entered_SNFKg").FormatString = "{0:n2}"
                Gv1.Columns("DiffEnteredVsMCC_Qty").HeaderText = "Quantity"
                Gv1.Columns("DiffEnteredVsMCC_Qty").FormatString = "{0:n0}"
                Gv1.Columns("DiffEnteredVsMCC_FAT").HeaderText = "KGFAT"
                Gv1.Columns("DiffEnteredVsMCC_SNF").HeaderText = "KGSNF"
                Gv1.Columns("DiffEnteredVsMCC_FAT").FormatString = "{0:n2}"
                Gv1.Columns("DiffEnteredVsMCC_SNF").FormatString = "{0:n2}"
            End If
            Gv1.Columns(ii).BestFit()
        Next
    End Sub
    Sub SetGridFormationOFGV1Multiple()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        'Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            Gv1.Columns("Document_No").HeaderText = "Document No."
            Gv1.Columns("Document_Date").HeaderText = "Document Date"
            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_NAME").HeaderText = "MCC NAME"
            Gv1.Columns("Mcc_Code_VLC_Uploader").HeaderText = "Uploader No."
            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("ROUTE_NAME").HeaderText = "Route Description"
            Gv1.Columns("Tanker_No").HeaderText = "Tanker No."
            Gv1.Columns("Vehicle_No").HeaderText = "Vehicle No."
            Gv1.Columns("Entered_Qty").HeaderText = "Entered Qty"
            Gv1.Columns("Entered_Qty").FormatString = "{0:n2}"
            Gv1.Columns("Entered_FATKg").HeaderText = "FATKG"
            Gv1.Columns("Entered_FATKg").FormatString = "{0:n2}"
            Gv1.Columns("Entered_SNFKg").HeaderText = "SNFKG"
            Gv1.Columns("Entered_SNFKg").FormatString = "{0:n2}"
            Gv1.Columns("CLR").HeaderText = "CLR"
            Gv1.Columns("CLR").FormatString = "{0:n0}"
            Gv1.Columns("FAT").HeaderText = "FAT"
            Gv1.Columns("FAT").FormatString = "{0:n2}"
            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Uploader No."
            Gv1.Columns("Qty").HeaderText = "Qty"
            Gv1.Columns("Qty").FormatString = "{0:n2}"
            Gv1.Columns("FATKG").HeaderText = "FATKG"
            Gv1.Columns("FATKG").FormatString = "{0:n2}"
            Gv1.Columns("SNFKG").HeaderText = "SNFKG"
            Gv1.Columns("SNFKG").FormatString = "{0:n2}"
            Gv1.Columns("DCS_FAT").HeaderText = "FAT%"
            Gv1.Columns("DCS_FAT").FormatString = "{0:n2}"
            Gv1.Columns("DCS_SNF").HeaderText = "SNF%"
            Gv1.Columns("DCS_SNF").FormatString = "{0:n2}"
            Gv1.Columns("Diff_Qty").HeaderText = "Qty"
            Gv1.Columns("Diff_Qty").FormatString = "{0:n2}"
            Gv1.Columns("Diff_FAT").HeaderText = "FATKG"
            Gv1.Columns("Diff_FAT").FormatString = "{0:n2}"
            Gv1.Columns("Diff_SNF").HeaderText = "SNFKG"
            Gv1.Columns("Diff_SNF").FormatString = "{0:n2}"
        Next
    End Sub
    Sub SetGridFormationOFGV1Collection()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        'Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            Gv1.Columns("Document_No").HeaderText = "Document No."
            Gv1.Columns("Document_Date").HeaderText = "Document Date"
            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_NAME").HeaderText = "MCC NAME"
            Gv1.Columns("Mcc_Code_VLC_Uploader").HeaderText = "Uploader No."
            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("ROUTE_NAME").HeaderText = "Route Description"
            Gv1.Columns("Tanker_No").HeaderText = "Tanker No."
            Gv1.Columns("Vehicle_No").HeaderText = "Vehicle No."
            Gv1.Columns("Entered_Qty").HeaderText = "Entered Qty"
            Gv1.Columns("Entered_Qty").FormatString = "{0:n2}"
            Gv1.Columns("Entered_FATKg").HeaderText = "FATKG"
            Gv1.Columns("Entered_FATKg").FormatString = "{0:n3}"
            Gv1.Columns("Entered_SNFKg").HeaderText = "SNFKG"
            Gv1.Columns("Entered_SNFKg").FormatString = "{0:n3}"
            Gv1.Columns("CLR").HeaderText = "CLR"
            Gv1.Columns("CLR").FormatString = "{0:n3}"
            Gv1.Columns("DCS_FAT").HeaderText = "FAT"
            Gv1.Columns("DCS_FAT").FormatString = "{0:n3}"
            Gv1.Columns("DCS_SNFPER").HeaderText = "SNF%"
            Gv1.Columns("DCS_SNFPER").FormatString = "{0:n2}"
            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Uploader No."
            Gv1.Columns("Collection_Date").HeaderText = "Collection Date"
            Gv1.Columns("Qty").HeaderText = "Qty"
            Gv1.Columns("Qty").FormatString = "{0:n2}"
            Gv1.Columns("FAT").HeaderText = "FAT"
            Gv1.Columns("FAT").FormatString = "{0:n2}"
            Gv1.Columns("SNF").HeaderText = "CLR"
            Gv1.Columns("SNF").FormatString = "{0:n2}"
            Gv1.Columns("FATKG").HeaderText = "FATKG"
            Gv1.Columns("FATKG").FormatString = "{0:n3}"
            Gv1.Columns("SNFKG").HeaderText = "SNFKG"
            Gv1.Columns("SNFKG").FormatString = "{0:n3}"
            Gv1.Columns("SNFPER").HeaderText = "SNF%"
            Gv1.Columns("SNFPER").FormatString = "{0:n2}"
            'Gv1.Columns("DCS_FAT").HeaderText = "DCS_FAT"
            'Gv1.Columns("DCS_FAT").FormatString = "{0:n2}"
            'Gv1.Columns("DCS_SNF").HeaderText = "DCS_SNF"
            'Gv1.Columns("DCS_SNF").FormatString = "{0:n2}"
        Next
    End Sub
    Sub SetGridFormationOFGV1Dock()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            Gv1.Columns("SNo").HeaderText = "S.No."
            Gv1.Columns("SDate").HeaderText = "Date"
            If Not rbtnDockDateWise.Checked Then
                Gv1.Columns("Shift").HeaderText = "Shift"
            End If
            Gv1.Columns("Milk_Wtd").HeaderText = "QTY"
            Gv1.Columns("BMC_Uploader_Code").HeaderText = "BMC Uploader Code"
            Gv1.Columns("BMC_Code").HeaderText = "BMC Code"
            Gv1.Columns("BMC_Code").IsVisible = False
            Gv1.Columns("BMC_Name").HeaderText = "BMC Name"
            Gv1.Columns("BMC_Name").IsVisible = False
            If AreaWiseBilling Then
                Gv1.Columns("Area").HeaderText = "Area"
            End If
            Gv1.Columns("Route_Code").HeaderText = "Route No"
            Gv1.Columns("Route_Code").IsVisible = True
            Gv1.Columns("DCS_Uploader_Code").HeaderText = "DCS Uploader Code"
            Gv1.Columns("DCS_Code").HeaderText = "DCS Code"
            Gv1.Columns("DCS_Code").IsVisible = False
            Gv1.Columns("DCS_Name").HeaderText = "DCS Name"
            Gv1.Columns("DCS_Name").IsVisible = False
            Gv1.Columns("Qty").HeaderText = "Can"
            Gv1.Columns("Qty").FormatString = "{0:n3}"
            Gv1.Columns("FAT").HeaderText = "FAT %"
            Gv1.Columns("FAT").FormatString = "{0:n2}"
            Gv1.Columns("SNF").HeaderText = "SNF %"
            Gv1.Columns("SNF").FormatString = "{0:n2}"
            Gv1.Columns("FAT_KG").HeaderText = "FAT KG"
            Gv1.Columns("SNF_KG").HeaderText = "SNF KG"
            If FatSnfRoundOff = 0 Then
                Gv1.Columns("FAT_KG").FormatString = "{0:n3}"
                Gv1.Columns("SNF_KG").FormatString = "{0:n3}"
            Else
                Gv1.Columns("FAT_KG").FormatString = "{0:n2}"
                Gv1.Columns("SNF_KG").FormatString = "{0:n2}"
            End If
            Gv1.Columns("RATE").HeaderText = "RATE"
            Gv1.Columns("RATE").FormatString = "{0:n3}"
            Gv1.Columns("RATE").IsVisible = False
            Gv1.Columns("Amt").HeaderText = "Amt"
            Gv1.Columns("Amt").FormatString = "{0:n3}"
            Gv1.Columns("Amt").IsVisible = False
            Gv1.Columns("VLC_CODE").HeaderText = "VLC CODE"
            Gv1.Columns("Alias_Name").HeaderText = "Alias Name"
            Gv1.Columns("Alias_Name").IsVisible = False
            Gv1.Columns("Alias_Name").VisibleInColumnChooser = True

            Gv1.Columns(ii).BestFit()
        Next
        Dim EnteredQty As New GridViewSummaryItem("Milk_Wtd", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty)
        Dim EnteredFatKg As New GridViewSummaryItem("Fat_KG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg)
        Dim EnteredSnfKg As New GridViewSummaryItem("SNF_KG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg)
        Dim summaryItem1 As New GridViewSummaryItem()
        summaryItem1.FormatString = "{0:F2}"
        summaryItem1.Name = "FAT"
        summaryItem1.AggregateExpression = "sum(Fat_KG)*100/sum(Milk_Wtd)"
        summaryRowItem.Add(summaryItem1)
        Dim summaryItem2 As New GridViewSummaryItem()
        summaryItem2.FormatString = "{0:F2}"
        summaryItem2.Name = "SNF"
        summaryItem2.AggregateExpression = "sum(SNF_KG)*100/sum(Milk_Wtd)"
        summaryRowItem.Add(summaryItem2)
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Sub View()
        If Gv1.Rows.Count > 0 Then
            Dim view As New ColumnGroupsViewDefinition()
            view.ColumnGroups.Add(New GridViewColumnGroup(" "))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            If rdbDetails.Checked Then
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_NAME").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("UploaderNo").Name)
            End If
            If rbtnBmcSummary.Checked Then
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_NAME").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("UploaderNo").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("DCSCode").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Alias_Name").Name)
            End If
            If Not rdbTankerWise.Checked AndAlso Not rbtnBMCTankerCollection.Checked AndAlso Not rbtnRouteWise.Checked AndAlso Not rbtnAliasName.Checked Then
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Document_No").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Document_Date").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Route_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("ROUTE_NAME").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker_No").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Vehicle_No").Name)
                'view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Reject_Types").Name)
                'view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Reject_Recovery_Pers").Name)
            End If
            If rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Reject_Types").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Reject_Recovery_Pers").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("GainLoss_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Loss_FAT_Rate").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Loss_SNF_Rate").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Start_Date").Name)
            End If
            '  Entered_Qty , Entered_FAT,Entered_SNF,Entered_FATKg,Entered_SNFKg
            ' DiffEnteredVsMCC_Qty, DiffEnteredVsMCC_FAT,DiffEnteredVsMCC_SNF,DiffEnteredVsMCC_FATKG,DiffEnteredVsMCC_SNFKG
            If rdbSummary.Checked Then
                view.ColumnGroups.Add(New GridViewColumnGroup("Entered Data"))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FAT").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNF").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Qty").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FAT").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNF").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FATKG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("DCS Data"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FAT").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNF").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FATKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data (BMC-DCS)"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_Qty").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FAT").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNF").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FATKG").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data(Entered - BMC)"))
                view.ColumnGroups(5).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_Qty").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_FAT").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_SNF").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_FATKG").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_SNFKG").Name)
            ElseIf rbtnTranpoterGainLoss.Checked OrElse rbtnTranspoterGainlossSummary.Checked Then
                If rbtnTranspoterGainlossSummary.Checked Then
                    view.ColumnGroups.Add(New GridViewColumnGroup("Receipt at Dock"))
                Else
                    view.ColumnGroups.Add(New GridViewColumnGroup("Entered Data"))
                End If
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FAT").Name)
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNF").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                'If rbtnTranspoterGainlossSummary.Checked = True Then
                '    view.ColumnGroups.Add(New GridViewColumnGroup("Milk Dispatch (BMC)"))
                'Else
                '    view.ColumnGroups.Add(New GridViewColumnGroup("Difference QTY"))
                'End If
                If rbtnTranspoterGainlossSummary.Checked Then
                    view.ColumnGroups.Add(New GridViewColumnGroup("Milk Dispatch (BMC)"))
                Else
                    view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                End If
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Qty").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FAT").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNF").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FATKG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data(BMC - Entered)"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DiffMCCVsEntered_Qty").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("FAT"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffMCCVsEntered_FATKG").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("FAT_Tolerence").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("FATKG_Recovered").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("FAT_AMT").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("SNF"))
                view.ColumnGroups(5).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("DiffMCCVsEntered_SNFKG").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("SNF_Tolerence").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("SNFKG_Recovered").Name)
                view.ColumnGroups(5).Rows(0).ColumnNames.Add(Gv1.Columns("SNF_AMT").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Total Recovery Amount"))
                view.ColumnGroups(6).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(6).Rows(0).ColumnNames.Add(Gv1.Columns("AMOUNT").Name)
            ElseIf rdbMultiple.Checked Then
                'view.ColumnGroups.Add(New GridViewColumnGroup(""))
                'view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Document_No").Name)
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Document_Date").Name)
                'view.ColumnGroups.Add(New GridViewColumnGroup(""))
                'view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_NAME").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Mcc_Code_VLC_Uploader").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Route_Code").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("ROUTE_NAME").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker_No").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Vehicle_No").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("FAT").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("CLR").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Code").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Name").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Code_VLC_Uploader").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("DCS Data"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("FATKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("SNFKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FAT").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNF").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data (BMC-DCS)"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_Qty").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FAT").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNF").Name)
                'view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FATKG").Name)
                'view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNFKG").Name)
            ElseIf rdbCollectionWise.Checked Then
                'view.ColumnGroups.Add(New GridViewColumnGroup(""))
                'view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Document_No").Name)
                'view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Document_Date").Name)
                'view.ColumnGroups.Add(New GridViewColumnGroup(""))
                'view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Code").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_NAME").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(Gv1.Columns("Mcc_Code_VLC_Uploader").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Route_Code").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("ROUTE_NAME").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker_No").Name)
                'view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Vehicle_No").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("CLR").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("FAT").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("SNFPer").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Code").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Name").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("VLC_Code_VLC_Uploader").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Collection_Date").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("DCS Data"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FAT").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("SNF").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("FATKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("SNFKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNFPER").Name)
                'view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FAT").Name)
                'view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNF").Name)
            ElseIf rdbTankerWise.Checked Then
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker_No").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Details Received from Tanker at BMC"))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Qty").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FATKG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Details Received at Dairy Dock"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_Qty").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_FAT").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_SNF").Name)
            ElseIf rbtnRouteWise.Checked Then
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("S.No.").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Date").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("ROUTE_NO").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("ROUTE_NAME").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("RejectInfo").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("As per Tanker Receipt as Dairy Dock"))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Qty").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FAT_Per").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNF_Per").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FATKG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("As per Milk Despatch"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FAT_Per").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNF_Per").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_FATKg").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Entered_SNFKg").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_Qty").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_FAT").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("DiffEnteredVsMCC_SNF").Name)
            ElseIf rbtnBMCTankerCollection.Checked Then
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("BMC Date").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Route No").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Route Name").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC Name").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("VLC Uploader Code").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("Alias Name").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("TANKER Data"))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker Qty").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker FAT(%)").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker SNF(%)").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker FAT KG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("Tanker SNF KG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("FAT(%)").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("SNF(%)").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("FAT KG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("SNF KG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data (TANKER-BMC)"))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff. Qty").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff. FAT(%)").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff. SNF(%)").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff. FAT KG").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(Gv1.Columns("Diff. SNF KG").Name)
            Else
                view.ColumnGroups.Add(New GridViewColumnGroup("BMC Data"))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_Qty").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FAT").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNF").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_FATKG").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(Gv1.Columns("MCC_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("DCS Data"))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_Qty").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FAT").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNF").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_FATKG").Name)
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(Gv1.Columns("DCS_SNFKG").Name)
                view.ColumnGroups.Add(New GridViewColumnGroup("Difference Data (BMC-DCS)"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_Qty").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FAT").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNF").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_FATKG").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(Gv1.Columns("Diff_SNFKG").Name)
            End If
            Gv1.ViewDefinition = view
        End If
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Reset()
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(PageSetupReport_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If obj IsNot Nothing AndAlso obj.GridColumns >= Gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To Gv1.Columns.Count - 1 Step ii + 1
                        Gv1.Columns(ii).IsVisible = False
                        Gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    Gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            common.clsCommon.MyMessageBoxShow(Me, err.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        Try
            If Gv1 IsNot Nothing AndAlso Gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptDailyQtyReport & "'"))
                arrHeader.Add("Date : " & clsCommon.myCstr(fromDate.Text) & "  To " & clsCommon.myCstr(dtpToDate.Text))
                If rdbSummary.Checked Then
                    arrHeader.Add("Report Type : " & "Summary")
                End If
                If rdbDetails.Checked Then
                    arrHeader.Add("Report Type : " & "Details")
                End If
                If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                    arrHeader.Add("Document Status : " & "Posted")
                ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                    arrHeader.Add("Document Status : " & "Unposted")
                Else
                    arrHeader.Add("Document Status : " & "All")
                End If
                transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
                'If rbtnDock.Checked Then
                '    transportSql.QuickExportToExcel(Gv1, "", Me.Text, , arrHeader)
                'Else
                transportSql.QuickExportToExcel(Gv1, "", Me.Text, , arrHeader)
                'transportSql.exportdata(Gv1, "", Me.Text, , arrHeader, False, False, True)
                'End If
                'clsCommon.MyExportToExcelGrid(Me.Text, Gv1, arrHeader, Me.Text, True)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmiPDF_Click(sender As Object, e As EventArgs) Handles rmiPDF.Click
        Try
            If Gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                If rdbTankerWise.Checked Then
                    arrHeader.Add("                                   " & objCommonVar.CurrentCompanyName)
                    arrHeader.Add("                                       Tanker KGFAT-KGSNF-Quantity of period " & clsCommon.myCstr(fromDate.Text) & "-" & clsCommon.myCstr(dtpToDate.Text))
                Else
                    ' arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                    'arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptDailyQtyReport & "'"))
                    arrHeader.Add("Date : " & clsCommon.myCstr(fromDate.Text) & "  To " & clsCommon.myCstr(dtpToDate.Text))
                End If
                If rdbSummary.Checked Then
                    arrHeader.Add("Report Type : " & "Summary")
                End If
                If rdbDetails.Checked Then
                    arrHeader.Add("Report Type : " & "Details")
                End If
                If clsCommon.CompairString(cboItemType.SelectedValue, "Posted") = CompairStringResult.Equal Then
                    arrHeader.Add("Document Status : " & "Posted")
                ElseIf clsCommon.CompairString(cboItemType.SelectedValue, "Unposted") = CompairStringResult.Equal Then
                    arrHeader.Add("Document Status : " & "Unposted")
                ElseIf Not rdbTankerWise.Checked Then
                    arrHeader.Add("Document Status : " & "All")
                End If
                transportSql.applyExportTemplate(Gv1, PageSetupReport_ID)
                If rdbTankerWise.Checked Then
                    Dim style As New GridPrintStyle()
                    style.PrintGrouping = True
                    style.HeaderCellBackColor = Color.White
                    style.GroupRowBackColor = Color.White
                    style.SummaryCellBackColor = Color.White
                    style.PrintSummaries = True
                    Gv1.PrintStyle = style
                    Dim doc As New clsMyPrintDocument()
                    doc.Margins.Top = 50
                    doc.Margins.Bottom = 50
                    doc.Margins.Left = 50
                    doc.Margins.Right = 50
                    doc.HeaderHeight = 90
                    doc.Landscape = True
                    doc.AssociatedObject = Gv1
                    doc.DocumentName = objCommonVar.CurrentCompanyName
                    doc.MiddleHeader = objCommonVar.CurrentCompanyName & Environment.NewLine & "    Tanker KGFAT-KGSNF-Quantity of period " & clsCommon.myCstr(fromDate.Text) & "-" & clsCommon.myCstr(dtpToDate.Text)
                    doc.HeaderFont = New Font("Segoe UI", 10, FontStyle.Bold)
                    doc.AssociatedObject = Gv1
                    doc.RightFooter = "Page [Page #] Of [Total Pages]"
                    Dim dialog As New RadPrintPreviewDialog
                    dialog.Document = doc
                    dialog.ToolMenu.Visible = True
                    dialog.Show()
                    doc.Print()
                    'doc = Nothing
                Else
                    'Gv1.Columns("Milk_Wtd").Width = 60
                    clsCommon.MyExportToPDF("Daily Quantity Report", Gv1, arrHeader, "", PageSetupReport_ID, objCommonVar.CurrentUserCode)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = PageSetupReport_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            Gv1.SaveLayout(obj.GridLayout)
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            obj.GridColumns = Gv1.ColumnCount
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub
    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
    End Sub
    Private Sub rptMccCollectionDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnPrint.Visible = False
        ToDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        fromDate.Value = clsCommon.GETSERVERDATE() 'ToDate.Value.AddMonths(-1)
        AreaWiseBilling = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = 1)
        isPickCLRInsteadOfSNF = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, Nothing)) = 1, True, False)
        'fndArea.Visible = AreaWiseBilling
        'lblArea.Visible = AreaWiseBilling

        If AreaWiseBilling Then
            rbtnBMCDCSPrint.Visible = True
            rbtnBMCRoutePrint.Visible = True
        Else
            rbtnBMCDCSPrint.Visible = False
            rbtnBMCRoutePrint.Visible = False
        End If

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            rbtnDockDetail.Visible = True
        End If
        LoadShiftName()
        TranspoterBoxhandler(False)
        Reset()
    End Sub
    Sub LoadShiftName()
        cboItemType.DataSource = LoadShift()
        cboItemType.ValueMember = "Code"
        cboItemType.DisplayMember = "Name"
        cboItemType.SelectedIndex = 0
    End Sub
    Public Function LoadShift() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "All"
        dr("Name") = "ALL"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Posted"
        dr("Name") = "Posted"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Unposted"
        dr("Name") = "Unposted"
        dt.Rows.Add(dr)
        Return dt
    End Function
    Private Sub txtMCC__My_Click(sender As Object, e As EventArgs) Handles txtMCC._My_Click
        Dim qry As String = " select TSPL_MCC_MASTER.MCC_Code as Code , TSPL_MCC_MASTER.MCC_NAME as Name from TSPL_MCC_MASTER "
        txtMCC.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelMCC@MCCCollectionRPT", qry, "Code", "Code", txtMCC.arrValueMember, txtMCC.arrDispalyMember)
    End Sub
    Private Sub rbtnDock_Click(sender As Object, e As EventArgs) Handles rbtnDock.Click
        If rbtnDock.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        btnGainLoss.Visible = False
        MyLabel4.Visible = True
        txtMCC_Code.Visible = True
        MyLabel5.Visible = False
        TxtTankerNo.Visible = False
        If AreaWiseBilling Then
            fndArea.Visible = True
            lblArea.Visible = True
        End If
        TranspoterBoxhandler(False)
    End Sub
    Private Sub txtMCC_Code__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtMCC_Code._MYValidating
        Try
            Dim qry As String = "select MCC_Code as Code,MCC_NAME as Name from TSPL_MCC_MASTER"
            txtMCC_Code.Value = clsCommon.ShowSelectForm("vbaMccm", qry, "Code", "", txtMCC_Code.Value, "Code", isButtonClicked)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text, MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub rdbDetails_Click(sender As Object, e As EventArgs) Handles rdbDetails.Click
        If rdbDetails.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        btnGainLoss.Visible = False
        MyLabel4.Visible = False
        txtMCC_Code.Visible = False
        MyLabel5.Visible = False
        TxtTankerNo.Visible = False
        TranspoterBoxhandler(False)
    End Sub
    Private Sub rdbSummary_Click(sender As Object, e As EventArgs) Handles rdbSummary.Click
        If rdbSummary.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        btnGainLoss.Visible = False
        MyLabel4.Visible = False
        txtMCC_Code.Visible = False
        MyLabel5.Visible = False
        TxtTankerNo.Visible = False
        TranspoterBoxhandler(False)
    End Sub
    Private Sub rbtnTranpoterGainLoss_Click(sender As Object, e As EventArgs) Handles rbtnTranpoterGainLoss.Click
        If rbtnTranpoterGainLoss.Checked Then
            btnPrint.Visible = True
        Else
            btnPrint.Visible = False
        End If
        btnGainLoss.Visible = False
        MyLabel4.Visible = False
        lblToleranceFAT.Visible = False
        txtMCC_Code.Visible = False
        MyLabel5.Visible = True
        TxtTankerNo.Visible = True
        fndArea.Visible = False
        lblArea.Visible = False
        TranspoterBoxhandler(True)
    End Sub
    Sub TranspoterBoxhandler(ByVal flag As Boolean)
        lblToleranceFAT.Visible = flag
        MyLabel7.Visible = flag
        lblToleranceFAT.Visible = flag
        lblToleranceSNF.Visible = flag
        txtToleranceFat.Visible = flag
        txtToleranceSNF.Visible = flag
        txtTolerenceSNF.Visible = flag
        'btnGainLoss.Visible = flag
    End Sub
    Private Sub TxtTankerNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles TxtTankerNo._MYValidating
        Try
            Dim qry As String = " select Tanker_No from TSPL_TANKER_MASTER where Tanker_No like '%" & TxtTankerNo.Value & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 AndAlso dt.Rows.Count = 1 Then
                'If dt.Rows.Count = 1 Then
                TxtTankerNo.Value = clsCommon.myCstr(dt.Rows(0)("Tanker_No"))
                'End If
            End If
            TxtTankerNo.Value = clsfrmTankerMaster.GetFinder("", TxtTankerNo.Value, isButtonClicked)
            'txtVehicleNo.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TANKER_NAME from TSPL_TANKER_MASTER where Tanker_No='" & TxtTankerNo.Value & "'"))
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnTranpoterGainLoss_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnTranpoterGainLoss.CheckedChanged
        If rbtnTranpoterGainLoss.Checked Then
            MyLabel5.Visible = True
            TxtTankerNo.Visible = True
            txtToleranceFat.Visible = True
            TxtTankerNo.Value = ""
            txtTolerenceSNF.Value = 0
            txtTolerenceSNF.Visible = True
            MyLabel7.Visible = True
        Else
            MyLabel5.Visible = False
            TxtTankerNo.Visible = False
            TxtTankerNo.Value = ""
            fndArea.Visible = False
            lblArea.Visible = False
            txtTolerenceSNF.Value = 0
            txtTolerenceSNF.Visible = False
            MyLabel7.Visible = False
            txtToleranceFat.Visible = False
        End If
    End Sub

    Sub DockDetail()
        Try
            Dim qry As String = ""
            Dim dt As New DataTable
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))

            qry = " Select ROW_NUMBER() OVER (ORDER BY (shift_date)) AS SrNo,Convert(Varchar,(xxx.Shift_Date),103)Shift_Date,Shift,max(Bulk_Route_Code)Bulk_Route_Code,(SNo)SNo,(Uploader_Code)Uploader_Code,QBD,SUM(No_Of_Cans)Can,sum(Milk_Weight)Milk_Weight,
                    case when max(Manual_Weight)=1 then 'Manual' else 'Auto' end as Manual_Auto_Weight,SUM(FAT)FAT,SUM(SNF)SNF,case when max(Manual_Sample)=1 then 'Manual' else 'Auto' end as QC_Manual_Auto,max(QAT)QAT

                     from (SELECT TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.No_Of_Cans,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift_Date,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Bulk_Route_Code,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Dock_Collection_Milk_Type,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as FATKG
                    ,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as SNFKG,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader_Code],TSPL_MCC_MASTER.MCC_NAME As MCC,
                    (case when TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No is not null then isnull (TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,'SWEET') 
                     end) as QBD,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.QAT,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Manual_Sample,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Manual_Weight,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNo

                        FROM TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL 
                        Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD On TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code  
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.MCC_Code
                        Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'
                        where Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
            If AreaWiseBilling Then
                If clsCommon.myLen(fndArea.Value) > 0 Then
                    qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                End If
            Else
                If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                    qry += " and TSPL_MCC_MASTER.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                End If
            End If
            'If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
            '    qry += " and TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(txtMCC_Code.Value) + "'"
            'End If
            If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='M'"
            ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='E'"
            End If
            qry += " ) xxx Group By xxx.Shift_Date,xxx.Shift, xxx.Uploader_Code,xxx.QBD,xxx.SNo order by xxx.Shift_Date"

            dt = clsDBFuncationality.GetDataTable(qry)
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                Gv1.DataSource = dt
                RadPageView1.SelectedPage = RadPageViewPage2
                SetGridFormatDockDetail()
                ControlEnableDisable(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub Summary()
        Try
            Dim qry As String = ""
            Dim dt As New DataTable
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))

            qry = "Select Convert(Varchar,(xxx.Shift_Date),103)Shift_Date,Max(xxx.VLC_Name)VLC_Name,Max(xxx.MCC)MCC,'" & Area & "' as Area,
                        xxx.Shift,Max(xxx.Bulk_Route_Code)Bulk_Route_Code,Case When Max(xxx.Dock_Collection_Milk_Type)='M' Then 'Mix' Else '' End As Dock_Collection_Milk_Type,Max(xxx.Reject_Type)Reject_Type,
                        Sum(xxx.No_Of_Cans)No_Of_Cans,
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FAT],
                        Sum(xxx.FATKG * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNF],
                        Sum(xxx.SNFKG  * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNFKG],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FAT],
                        Sum(xxx.FATKG * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNF],
                        Sum(xxx.SNFKG * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNFKG],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_FAT],
                        Sum(xxx.SNFKG  * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNF],
                        Sum(xxx.SNFKG * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNFKG]
                         from
                        (SELECT TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.No_Of_Cans,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift_Date,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Bulk_Route_Code,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Dock_Collection_Milk_Type,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as FATKG
,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as SNFKG,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader_Code],TSPL_MCC_MASTER.MCC_NAME As MCC,TSPL_COMPANY_MASTER.Comp_Name 
                        FROM TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL 
                        Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD On TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code  
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.MCC_Code
                        Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'
                        where Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
            If AreaWiseBilling Then
                If clsCommon.myLen(fndArea.Value) > 0 Then
                    qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                End If
            Else
                If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                    qry += " and TSPL_MCC_MASTER.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                End If
            End If
            'If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
            '    qry += " and TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(txtMCC_Code.Value) + "'"
            'End If
            If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='M'"
            ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='E'"
            End If
            qry += " ) xxx Group By xxx.Shift_Date,xxx.Shift ORDER BY xxx.Shift_Date"

            dt = clsDBFuncationality.GetDataTable(qry)
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                Gv1.DataSource = dt
                RadPageView1.SelectedPage = RadPageViewPage2
                SetGridFormatSummary()
                ControlEnableDisable(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub DockSummary()
        Try
            Dim FinalQuery As String = ""
            Dim dt As New DataTable
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))

            FinalQuery = " SELECT ROW_NUMBER() OVER (ORDER BY CONVERT(INT, A.Uploader_Code)) AS SrNo,A.*
FROM
(Select xxx.SNO,Convert(Varchar,(xxx.Shift_Date),103)Shift_Date,Convert(int,xxx.Uploader_Code)Uploader_Code,Max(xxx.VLC_Name)VLC_Name,Max(xxx.MCC)MCC,'" & Area & "' as Area,
                        xxx.Shift,Max(xxx.Bulk_Route_Code)Bulk_Route_Code,Case When Max(xxx.Dock_Collection_Milk_Type)='M' Then 'Mix' Else '' End As Dock_Collection_Milk_Type,Max(xxx.Reject_Type)Reject_Type,
                        Sum(xxx.No_Of_Cans)No_Of_Cans,                        
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) As [Sweet_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_Weight],
                        Max(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNF],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNF],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as[CURD_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNF]  
                         from
                        (SELECT TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.*,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader_Code],TSPL_MCC_MASTER.MCC_NAME As MCC,TSPL_COMPANY_MASTER.Comp_Name 
                        FROM TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL 
                        Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD On TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code  
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.MCC_Code    
                        Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'
                        where Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  
                        And Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
            If AreaWiseBilling Then
                If clsCommon.myLen(fndArea.Value) > 0 Then
                    FinalQuery += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                End If
            Else
                If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                    FinalQuery += " and TSPL_MCC_MASTER.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                End If
            End If
            'If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
            '    qry += " and TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(txtMCC_Code.Value) + "'"
            'End If
            If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                FinalQuery += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='M'"
            ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                FinalQuery += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='E'"
            End If
            FinalQuery += " ) xxx Group By xxx.Uploader_Code,xxx.Shift_Date,xxx.Shift,xxx.SNO ) A ORDER BY  Convert(int,A.Uploader_Code)"


            dt = clsDBFuncationality.GetDataTable(FinalQuery)
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                Gv1.DataSource = dt
                RadPageView1.SelectedPage = RadPageViewPage2
                SetGridFormatDockSummary()
                ControlEnableDisable(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub BMCDock()
        Try
            Dim qry As String = ""
            Dim dt As New DataTable
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))


            qry = "Select Max(Convert(Varchar,xxx.Shift_Date,103))Shift_Date,Case When xxx.Shift='M' Then 'Morning' Else 'Evening' End As Shift,"
            If AreaWiseBilling Then
                qry += " Max(xxx.Location_Desc)MCC_NAME,"
            Else
                qry += " Max(xxx.MCC_NAME)MCC_NAME,"
            End If
            qry += "Max(xxx.VLC)VLC,'" & Area & "' as Area,Max(xxx.[VLC Name])[VLC Name],Max(xxx.Route)Route,(Convert(int,xxx.[Route Code]))[Route Code],Sum(xxx.[No of Cans])[No of Cans],Sum(xxx.[Good Qty])[Good Qty],Sum(xxx.[Good FAT %])[Good FAT %],Sum(xxx.[Good FATKg])[Good FATKg],Sum(xxx.[Good SNF %])[Good SNF %],Sum(xxx.[Good SNFKG])[Good SNFKG],Sum(xxx.[SOUR Qty])[SOUR Qty],Sum(xxx.[SOUR FAT %])[SOUR FAT %],Sum(xxx.[SOUR FATKg])[SOUR FATKg],Sum(xxx.[SOUR SNF %])[SOUR SNF %],Sum(xxx.[SOUR SNFKG])[SOUR SNFKG],Sum(xxx.[CURD Qty])[CURD Qty],Sum([CURD FAT %])[CURD FAT %],Sum(xxx.[CURD FATKg])[CURD FATKg],Sum(xxx.[CURD SNF %])[CURD SNF %],Sum(xxx.[CURD SNFKG])[CURD SNFKG]
                        from (select TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No,"
            If AreaWiseBilling Then
                qry += " xxxSetLocation.Location_Desc,"
            End If
            qry += "TSPL_MILK_SHIFT_UPLOADER_DETAIL.SNo,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC, TSPL_MILK_SHIFT_UPLOADER_DETAIL.VLC_Code as [VLC Code],TSPL_VLC_MASTER_HEAD.VLC_Name as [VLC Name],TSPL_MILK_SHIFT_UPLOADER_DETAIL.No_Of_Cans as [No of Cans],TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO as [Route Code],TSPL_BULK_ROUTE_MASTER.ROUTE_NAME as [Route],TSPL_COMPANY_MASTER.Comp_Name,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date
                        ,case When isnull(Reject_Type,'')='' then Milk_Weight else 0 end as [Good Qty]
                        ,case When isnull(Reject_Type,'')='' then FAT else 0 end as [Good FAT %]
                        ,case When isnull(Reject_Type,'')='' then cast(Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [Good FATKg]
                        ,case When isnull(Reject_Type,'')='' then SNF else 0 end as [Good SNF %]
                        ,case When isnull(Reject_Type,'')='' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [Good SNFKG],case When isnull(Reject_Type,'')='SOUR' then Milk_Weight else 0 end as [SOUR Qty]
                        ,case When isnull(Reject_Type,'')='SOUR' then FAT else 0 end as [SOUR FAT %]
                        ,case When isnull(Reject_Type,'')='SOUR' then cast (Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [SOUR FATKg]
                        ,case When isnull(Reject_Type,'')='SOUR' then SNF else 0 end as [SOUR SNF %]
                        ,case When isnull(Reject_Type,'')='SOUR' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [SOUR SNFKG],case When isnull(Reject_Type,'')='CURD' then Milk_Weight else 0 end as [CURD Qty]
                        ,case When isnull(Reject_Type,'')='CURD' then FAT else 0 end as [CURD FAT %]
                        ,case When isnull(Reject_Type,'')='CURD' then cast (Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [CURD FATKg]
                        ,case When isnull(Reject_Type,'')='CURD' then SNF else 0 end as [CURD SNF %]
                        ,case When isnull(Reject_Type,'')='CURD' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [CURD SNFKG] ,TSPL_MILK_SHIFT_UPLOADER_DETAIL.PageNo 
                        from TSPL_MILK_SHIFT_UPLOADER_DETAIL 
                        left outer Join TSPL_MILK_SHIFT_UPLOADER_HEAD On TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No=TSPL_MILK_SHIFT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_SHIFT_UPLOADER_DETAIL.VLC_Code
                        left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC "
            If AreaWiseBilling Then
                qry += " Left Outer Join( select TSPL_PAYMENT_PROCESS_HEAD.Doc_No,tspl_location_master.Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code)  xxxSetLocation On xxxSetLocation.Location_Code=TSPL_MCC_MASTER.area_Location_code "
            End If
            qry += "  Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'                        
                        where Convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date,103)<=convert (date,'" & dtpToDate.Value & "',103) "
            If AreaWiseBilling AndAlso clsCommon.myLen(fndArea.Value) > 0 Then
                'If clsCommon.myLen(fndArea.Value) > 0 Then
                qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                'End If
            Else
                If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                    qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                End If
            End If
            If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift='M'"
            ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift='E'"
            End If
            qry += ") xxx Group By xxx.[Route Code],xxx.Shift_Date,xxx.shift order by Convert(int,xxx.[Route Code]),xxx.shift desc"

            dt = clsDBFuncationality.GetDataTable(qry)
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                Gv1.DataSource = dt
                RadPageView1.SelectedPage = RadPageViewPage2
                SetGridFormatBMCDock()
                ControlEnableDisable(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Sub SetGridFormatSummary()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            'Gv1.Columns("SrNo").HeaderText = "S.No."
            'Gv1.Columns("SNO").IsVisible = False
            Gv1.Columns("Shift_Date").HeaderText = "Date"
            Gv1.Columns("VLC_Name").HeaderText = "DCS Name"
            Gv1.Columns("MCC").HeaderText = "MCC Name"
            Gv1.Columns("Area").HeaderText = "Area"
            Gv1.Columns("Shift").HeaderText = "Shift"
            Gv1.Columns("Bulk_Route_Code").HeaderText = "Route Code"
            Gv1.Columns("Dock_Collection_Milk_Type").HeaderText = "Dock Collection Milk Type"
            Gv1.Columns("Reject_Type").HeaderText = "Reject Type"
            Gv1.Columns("No_Of_Cans").HeaderText = "No Of Cans"
            Gv1.Columns("Sweet_No_Of_Cans").HeaderText = "Sweet No Of Cans"
            Gv1.Columns("Sweet_Milk_Weight").HeaderText = "Sweet Milk Weight"
            Gv1.Columns("Sweet_Milk_FAT").HeaderText = "Sweet Milk FAT"
            Gv1.Columns("Sweet_Milk_FATKG").HeaderText = "Sweet Milk FATKG"
            Gv1.Columns("Sweet_Milk_SNF").HeaderText = "Sweet Milk SNF"
            Gv1.Columns("Sweet_Milk_SNFKG").HeaderText = "Sweet Milk SNFKG"
            Gv1.Columns("SOUR_No_Of_Cans").HeaderText = "SOUR No Of Cans"
            Gv1.Columns("SOUR_Milk_Weight").HeaderText = "SOUR Milk Weight"
            Gv1.Columns("SOUR_Milk_FAT").HeaderText = "SOUR Milk FAT"
            Gv1.Columns("SOUR_Milk_FATKG").HeaderText = "SOUR Milk FATKG"
            Gv1.Columns("SOUR_Milk_SNF").HeaderText = "SOUR Milk SNF"
            Gv1.Columns("SOUR_Milk_SNFKG").HeaderText = "SOUR Milk SNFKG"
            Gv1.Columns("CURD_No_Of_Cans").HeaderText = "CURD No Of Cans"
            Gv1.Columns("CURD_Milk_Weight").HeaderText = "CURD Milk Weight"
            Gv1.Columns("CURD_Milk_FAT").HeaderText = "CURD Milk FAT"
            Gv1.Columns("CURD_Milk_FATKG").HeaderText = "CURD Milk FATKG"
            Gv1.Columns("CURD_Milk_SNF").HeaderText = "CURD Milk SNF"
            Gv1.Columns("CURD_Milk_SNFKG").HeaderText = "CURD Milk SNFKG"

            Gv1.Columns(ii).BestFit()
        Next

        Dim EnteredQty As New GridViewSummaryItem("No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty)
        Dim EnteredFatKg As New GridViewSummaryItem("Sweet_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg)
        Dim EnteredSnfKg As New GridViewSummaryItem("Sweet_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg)
        Dim EnteredSnfKg1 As New GridViewSummaryItem("Sweet_Milk_FATKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg1)
        Dim EnteredSnfKg2 As New GridViewSummaryItem("Sweet_Milk_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg2)
        Dim EnteredQty1 As New GridViewSummaryItem("SOUR_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty1)
        Dim EnteredFatKg1 As New GridViewSummaryItem("SOUR_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg1)
        Dim EnteredFatKg12 As New GridViewSummaryItem("SOUR_Milk_FATKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg12)
        Dim EnteredFatKg13 As New GridViewSummaryItem("SOUR_Milk_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg13)
        Dim EnteredSnfKg11 As New GridViewSummaryItem("CURD_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg11)
        Dim EnteredQty11 As New GridViewSummaryItem("CURD_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty11)
        Dim EnteredQty12 As New GridViewSummaryItem("CURD_Milk_FATKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty12)
        Dim EnteredQty113 As New GridViewSummaryItem("CURD_Milk_SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty113)

        Gv1.EnableFiltering = True
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Sub SetGridFormatBMCDock()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            'Gv1.Columns("SrNo").HeaderText = "S.No."
            'Gv1.Columns("SNO").IsVisible = False
            Gv1.Columns("Shift_Date").HeaderText = "Date"
            Gv1.Columns("Shift").HeaderText = "Shift"
            Gv1.Columns("MCC_NAME").HeaderText = "MCC Name"
            Gv1.Columns("Area").HeaderText = "Area"
            Gv1.Columns("VLC Name").HeaderText = "DCS Name"
            Gv1.Columns("Route").HeaderText = "Route No"
            Gv1.Columns("Route Code").HeaderText = "Route Code"
            'Gv1.Columns("MCC").HeaderText = "MCC"
            'Gv1.Columns("Dock_Collection_Milk_Type").HeaderText = "Dock Collection Milk Type"
            'Gv1.Columns("Reject_Type").HeaderText = "Reject Type"
            Gv1.Columns("No Of Cans").HeaderText = "No Of Cans"
            Gv1.Columns("Good Qty").HeaderText = "Good Qty"
            Gv1.Columns("Good FAT %").HeaderText = "Good FAT %"
            Gv1.Columns("Good FATKg").HeaderText = "Good FATKg"
            Gv1.Columns("Good SNF %").HeaderText = "Good SNF %"
            Gv1.Columns("Good SNFKG").HeaderText = "Good SNFKG"
            Gv1.Columns("SOUR Qty").HeaderText = "SOUR Qty"
            Gv1.Columns("SOUR FAT %").HeaderText = "SOUR FAT %"
            Gv1.Columns("SOUR FATKg").HeaderText = "SOUR FATKg"
            Gv1.Columns("SOUR SNF %").HeaderText = "SOUR SNF %"
            Gv1.Columns("SOUR SNFKG").HeaderText = "SOUR SNFKG"
            Gv1.Columns("CURD Qty").HeaderText = "CURD Qty"
            Gv1.Columns("CURD FAT %").HeaderText = "CURD FAT %"
            Gv1.Columns("CURD FATKg").HeaderText = "CURD FATKg"
            Gv1.Columns("CURD SNF %").HeaderText = "CURD SNF %"
            Gv1.Columns("CURD SNFKG").HeaderText = "CURD SNFKG"

            Gv1.Columns(ii).BestFit()
        Next

        Dim EnteredQty As New GridViewSummaryItem("No Of Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty)
        Dim EnteredFatKg As New GridViewSummaryItem("Good Qty", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg)
        Dim EnteredSnfKg As New GridViewSummaryItem("Good FATKg", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg)
        Dim EnteredQty1 As New GridViewSummaryItem("Good SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty1)
        Dim EnteredFatKg1 As New GridViewSummaryItem("SOUR Qty", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg1)
        Dim EnteredSnfKg1 As New GridViewSummaryItem("SOUR FATKg", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg1)
        Dim EnteredQty11 As New GridViewSummaryItem("SOUR SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty11)
        Dim EnteredQty13 As New GridViewSummaryItem("CURD Qty", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty13)
        Dim EnteredQty14 As New GridViewSummaryItem("CURD FATKg", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty14)
        Dim EnteredQty15 As New GridViewSummaryItem("CURD SNFKG", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty15)

        Gv1.EnableFiltering = True
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Sub SetGridFormatDockSummary()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            Gv1.Columns("SrNo").HeaderText = "S.No."
            Gv1.Columns("SNO").IsVisible = False
            Gv1.Columns("Shift_Date").HeaderText = "Date"
            Gv1.Columns("Uploader_Code").HeaderText = "DCS Code"
            Gv1.Columns("VLC_Name").HeaderText = "DCS Name"
            Gv1.Columns("MCC").HeaderText = "MCC"
            Gv1.Columns("Area").HeaderText = "Area"
            Gv1.Columns("Shift").HeaderText = "Shift"
            Gv1.Columns("Bulk_Route_Code").HeaderText = "Route No"
            Gv1.Columns("Dock_Collection_Milk_Type").HeaderText = "Dock Collection Milk Type"
            Gv1.Columns("Reject_Type").HeaderText = "Reject Type"
            Gv1.Columns("No_Of_Cans").HeaderText = "No Of Cans"
            Gv1.Columns("Sweet_No_Of_Cans").HeaderText = "Sweet No Of Cans"
            Gv1.Columns("Sweet_Milk_Weight").HeaderText = "Sweet Milk Weight"
            Gv1.Columns("Sweet_Milk_FAT").HeaderText = "Sweet Milk FAT"
            Gv1.Columns("Sweet_Milk_SNF").HeaderText = "Sweet Milk SNF"
            Gv1.Columns("SOUR_No_Of_Cans").HeaderText = "SOUR No Of Cans"
            Gv1.Columns("SOUR_Milk_Weight").HeaderText = "SOUR Milk Weight"
            Gv1.Columns("SOUR_Milk_FAT").HeaderText = "SOUR_Milk_FAT"
            Gv1.Columns("SOUR_Milk_SNF").HeaderText = "SOUR Milk SNF"
            Gv1.Columns("CURD_No_Of_Cans").HeaderText = "CURD No Of Cans"
            Gv1.Columns("CURD_Milk_Weight").HeaderText = "CURD Milk Weight"
            Gv1.Columns("CURD_Milk_FAT").HeaderText = "CURD Milk FAT"
            Gv1.Columns("CURD_Milk_SNF").HeaderText = "CURD Milk SNF"


            Gv1.Columns(ii).BestFit()
        Next
        Dim EnteredQty As New GridViewSummaryItem("No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty)
        Dim EnteredFatKg As New GridViewSummaryItem("Sweet_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg)
        Dim EnteredSnfKg As New GridViewSummaryItem("Sweet_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg)
        Dim EnteredQty1 As New GridViewSummaryItem("SOUR_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty1)
        Dim EnteredFatKg1 As New GridViewSummaryItem("SOUR_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg1)
        Dim EnteredSnfKg1 As New GridViewSummaryItem("CURD_No_Of_Cans", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg1)
        Dim EnteredQty11 As New GridViewSummaryItem("CURD_Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty11)

        Gv1.EnableFiltering = True
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Sub SetGridFormatDockDetail()
        Gv1.TableElement.TableHeaderHeight = 40
        Gv1.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).IsVisible = True
            Gv1.Columns(ii).Width = 40
            Gv1.Columns("SrNo").HeaderText = "S.No."
            Gv1.Columns("Shift_Date").HeaderText = "Date"
            Gv1.Columns("Shift").HeaderText = "Shift"
            Gv1.Columns("Bulk_Route_Code").HeaderText = "Route No"
            Gv1.Columns("SNO").HeaderText = "Sample No"
            Gv1.Columns("Uploader_Code").HeaderText = "DCS Code"
            Gv1.Columns("QBD").HeaderText = "Milk Type"
            Gv1.Columns("Can").HeaderText = "Can"
            Gv1.Columns("Milk_Weight").HeaderText = "Qty"
            Gv1.Columns("Manual_Auto_Weight").HeaderText = "Weight AutoManual"
            Gv1.Columns("FAT").HeaderText = "FAT"
            Gv1.Columns("SNF").HeaderText = "SNF"
            Gv1.Columns("QC_Manual_Auto").HeaderText = "QC ManualAuto"
            Gv1.Columns("QAT").HeaderText = "QAP"

            Gv1.Columns(ii).BestFit()
        Next
        Dim EnteredQty As New GridViewSummaryItem("Can", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredQty)
        Dim EnteredFatKg As New GridViewSummaryItem("QAT", "{0:F0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredFatKg)
        Dim EnteredSnfKg As New GridViewSummaryItem("Milk_Weight", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(EnteredSnfKg)

        Gv1.EnableFiltering = True
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim qry As String = Nothing
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))
            If rbtnDockSummary.Checked Then
                qry = "Select '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [FromDate],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [ToDate],'" & objCommonVar.CurrentUser & "' as User_Name,xxx.SNO,Convert(Date,(xxx.Shift_Date),103)Shift_Date,Convert(int,xxx.Uploader_Code)Uploader_Code,Max(xxx.VLC_Name)VLC_Name,Max(xxx.MCC)MCC,'" & Area & "' as Area,
                        xxx.Shift,Max(xxx.Bulk_Route_Code)Bulk_Route_Code,Case When Max(xxx.Dock_Collection_Milk_Type)='M' Then 'Mix' Else '' End As Dock_Collection_Milk_Type,Max(xxx.Reject_Type)Reject_Type,
                        Sum(xxx.No_Of_Cans)No_Of_Cans,                        
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) As [Sweet_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_Weight],
                        Max(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNF],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNF],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as[CURD_Milk_FAT],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNF],  
                        Max(xxx.Comp_Name)Comp_Name from
                        (SELECT TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.*,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader_Code],TSPL_MCC_MASTER.MCC_NAME As MCC,TSPL_COMPANY_MASTER.Comp_Name 
                        FROM TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL 
                        Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD On TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code  
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.MCC_Code    
                        Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'
                        where Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
                If AreaWiseBilling Then
                    If clsCommon.myLen(fndArea.Value) > 0 Then
                        qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                    End If
                Else
                    If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                        qry += " and TSPL_MCC_MASTER.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                    End If
                End If
                'If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                '    qry += " and TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(txtMCC_Code.Value) + "'"
                'End If
                If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='M'"
                ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='E'"
                End If
                qry += " ) xxx Group By xxx.Uploader_Code,xxx.Shift_Date,xxx.Shift,xxx.SNO ORDER BY  Convert(int,xxx.Uploader_Code)"
            ElseIf rbtnBMCDCSPrint.Checked Then
                qry = "Select    XXX.SNo AS [SNO], (xxx.Comp_Name) as [Comp Name],xxx.Comp_Code as[comp code],xxx.Add1 as [Add1],xxx.Location_Desc as [MCC NAME]
,case When (isnull(Milk_Type,''))='Good' then (isnull(qty,0)) else 0 end as [Good Qty]
,case When (isnull(Milk_Type,''))='Good' then (FAT) else 0 end as [Good FAT %]
,case When (isnull(Milk_Type,''))='Good' then (cast(qty*FAT/100 as decimal(18,3))) else 0 end as [Good FATKg]
,case When (isnull(Milk_Type,''))='Good' then (SNF) else 0 end as [Good SNF %]
,case When (isnull(Milk_Type,''))='Good' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [Good SNFKG],
case When (isnull(Milk_Type,''))='SOUR' then (qty) else 0 end as [SOUR Qty]
,case When (isnull(Milk_Type,''))='SOUR' then (FAT) else 0 end as [SOUR FAT %]
,case When (isnull(Milk_Type,''))='SOUR' then (cast (qty*FAT/100 as decimal(18,3))) else 0 end as [SOUR FATKg]
,case When (isnull(Milk_Type,''))='SOUR' then (SNF) else 0 end as [SOUR SNF %]
,case When (isnull(Milk_Type,''))='SOUR' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [SOUR SNFKG],
case When (isnull(Milk_Type,''))='CURD' then (qty) else 0 end as [CURD Qty]
,case When (isnull(Milk_Type,''))='CURD' then (FAT) else 0 end as [CURD FAT %]
,case When (isnull(Milk_Type,''))='CURD' then (cast (qty*FAT/100 as decimal(18,3))) else 0 end as [CURD FATKg]
,case When (isnull(Milk_Type,''))='CURD' then (SNF) else 0 end as [CURD SNF %]
,case When (isnull(Milk_Type,''))='CURD' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [CURD SNFKG] ,
Case when Shift='M' Then Qty Else 0 End As MorningQty,
       Case when Shift='E' Then Qty Else 0 End As EveningQty,
	   Case when Shift='M' Then (Snf) Else 0 End As MorningSNF,
	   Case when Shift='E' Then Snf Else 0 End As EveningSNF,
	   Case when Shift='M' Then FAT Else 0 End As MorningFat,
	   Case when Shift='E' Then FAT Else 0 End As EveningFat,
	   Case when Shift='M' Then FATKG Else 0 End As MorningFatKG,
	   Case when Shift='E' Then FATKG Else 0 End As EveningFatKG,
	     Case when Shift='M' Then SNFKG Else 0 End As MorningSNFKG,
	   Case when Shift='E' Then SNFKG Else 0 End As EveningSNFKG,
	   XXX.Milk_Type AS [Milk Type],
	   XXX.VLC_Name AS [DCS NAME],
	   xxx.Route_no as [Route Code],
	   XXX.VLC_NAME AS [VLC NAME],
	   XXX.VLC_Code_VLC_Uploader AS [VLC],
  XXX.SHIFT AS [Shift] 
	,[Shift Date],xxx.Logo_Img,XXX.Logo_Img2,XXX.User_Name,xxx.ROUTE_NAME
		,XXX.Tanker_No As [Tanker No],xxx.Schedule_Time_Morning as [sch Time]
from 
(Select '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [FromDate],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [ToDate],TSPL_MILK_COLLECTION_DCS_DETAIL.Milk_Type,TSPL_MILK_COLLECTION_DCS_DETAIL.Shift, TSPL_MILK_COLLECTION_DCS_DETAIL.Qty,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Comp_Code,TSPL_COMPANY_MASTER.Add1,tspl_location_master.Location_Desc,
TSPL_MILK_COLLECTION_DCS_DETAIL.fat,TSPL_MILK_COLLECTION_DCS_DETAIL.snf,TSPL_MILK_COLLECTION_DCS_DETAIL.SNo, '" & objCommonVar.CurrentUser & "' as User_Name,Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) As [Shift Date],TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,
TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG,TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG,
TSPL_MILK_COLLECTION_DCS_DETAIL.VLC_Code,
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_BULK_ROUTE_MASTER.Route_no,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME
,TSPL_BULK_ROUTE_MASTER.Tanker_No,TSPL_BULK_ROUTE_MASTER.Schedule_Time_Morning
FROM TSPL_MILK_COLLECTION_DCS_DETAIL 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_COLLECTION_DCS_DETAIL.VLC_Code 
left outer join tspl_mcc_master on tspl_mcc_master.MCC_Code =TSPL_VLC_MASTER_HEAD.mcc
left outer join tspl_location_master on tspl_location_master.Location_Code =tspl_mcc_master.Area_Location_Code
left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No=TSPL_VLC_MASTER_HEAD.Route_Code
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code
LEFT OUTER JOIN TSPL_MILK_COLLECTION_DCS ON TSPL_MILK_COLLECTION_DCS.Document_No=TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                         left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_VLC_MASTER_HEAD.comp_code
				 where Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
                If AreaWiseBilling Then
                    qry += "and tspl_mcc_master.Area_Location_Code= '" & fndArea.Value & "' "
                End If
                If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_COLLECTION_DCS_DETAIL.Shift='M' "
                ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_COLLECTION_DCS_DETAIL.Shift='E' "
                End If
                qry += " )xxx order by sno"
            ElseIf rbtnBMCRoutePrint.Checked Then
                qry = "Select '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [FromDate],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [ToDate],TSPL_COMPANY_MASTER.Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,
tspl_location_master.Location_Desc ,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_BULK_ROUTE_MASTER.ROUTE_NO as [Route Code],TSPL_BULK_ROUTE_MASTER.Route_Name as [Route Name],
TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MILK_COLLECTION_DCS_DETAIL.Shift As [Shift],Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) As [Shift Date],TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,'" & objCommonVar.CurrentUser & "' as User_Name,
case When (isnull(Milk_Type,''))='Good' then (isnull(qty,0)) else 0 end as [Good Qty]
,case When (isnull(Milk_Type,''))='Good' then (FAT) else 0 end as [Good FAT %]
,case When (isnull(Milk_Type,''))='Good' then (cast(qty*FAT/100 as decimal(18,3))) else 0 end as [Good FATKg]
,case When (isnull(Milk_Type,''))='Good' then (SNF) else 0 end as [Good SNF %]
,case When (isnull(Milk_Type,''))='Good' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [Good SNFKG],
case When (isnull(Milk_Type,''))='SOUR' then (qty) else 0 end as [SOUR Qty]
,case When (isnull(Milk_Type,''))='SOUR' then (FAT) else 0 end as [SOUR FAT %]
,case When (isnull(Milk_Type,''))='SOUR' then (cast (qty*FAT/100 as decimal(18,3))) else 0 end as [SOUR FATKg]
,case When (isnull(Milk_Type,''))='SOUR' then (SNF) else 0 end as [SOUR SNF %]
,case When (isnull(Milk_Type,''))='SOUR' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [SOUR SNFKG],
case When (isnull(Milk_Type,''))='CURD' then (qty) else 0 end as [CURD Qty]
,case When (isnull(Milk_Type,''))='CURD' then (FAT) else 0 end as [CURD FAT %]
,case When (isnull(Milk_Type,''))='CURD' then (cast (qty*FAT/100 as decimal(18,3))) else 0 end as [CURD FATKg]
,case When (isnull(Milk_Type,''))='CURD' then (SNF) else 0 end as [CURD SNF %]
,case When (isnull(Milk_Type,''))='CURD' then (cast (qty*SNF/100 as decimal(18,3))) else 0 end as [CURD 
SNFKG] 
FROM TSPL_MILK_COLLECTION_DCS_DETAIL 
 left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_COLLECTION_DCS_DETAIL.VLC_Code 
 left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No=TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
 left outer join tspl_mcc_master on tspl_mcc_master.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
 left outer join tspl_location_master on tspl_location_master.Location_Code=tspl_mcc_master.Area_Location_Code
left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No=TSPL_VLC_MASTER_HEAD.Route_Code
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code
                          left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_VLC_MASTER_HEAD.comp_code
				 where Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103) "
                If AreaWiseBilling Then
                    qry += " and tspl_mcc_master.Area_Location_Code= '" & fndArea.Value & "' "
                End If
                If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_COLLECTION_DCS_DETAIL.Shift='M' "
                ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_COLLECTION_DCS_DETAIL.Shift='E' "
                End If
            ElseIf rbtnSummary.Checked Then
                qry = "Select '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [FromDate],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [ToDate],'" & objCommonVar.CurrentUser & "' as User_Name,Convert(Date,(xxx.Shift_Date),103)Shift_Date,Max(xxx.VLC_Name)VLC_Name,Max(xxx.MCC)MCC,'" & Area & "' as Area,
                        xxx.Shift,Max(xxx.Bulk_Route_Code)Bulk_Route_Code,Case When Max(xxx.Dock_Collection_Milk_Type)='M' Then 'Mix' Else '' End As Dock_Collection_Milk_Type,Max(xxx.Reject_Type)Reject_Type,
                        Sum(xxx.No_Of_Cans)No_Of_Cans,
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FAT],
                        Sum(xxx.FATKG * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNF],
                        Sum(xxx.SNFKG  * case when Isnull(xxx.Reject_Type,'')='' then 1 else 0 end  ) as [Sweet_Milk_SNFKG],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FAT],
                        Sum(xxx.FATKG * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNF],
                        Sum(xxx.SNFKG * case when Isnull(xxx.Reject_Type,'')='SOUR' then 1 else 0 end  ) as [SOUR_Milk_SNFKG],
                        Sum(IsNull(xxx.No_Of_Cans,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_No_Of_Cans],
                        Sum(IsNull(xxx.Milk_Weight,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_Weight],
                        Sum(IsNull(xxx.FAT,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_FAT],
                        Sum(xxx.SNFKG  * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_FATKG],
                        Sum(IsNull(xxx.SNF,0) * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNF],
                        Sum(xxx.SNFKG * case when Isnull(xxx.Reject_Type,'')='CURD' then 1 else 0 end  ) as [CURD_Milk_SNFKG],
                        Max(xxx.Comp_Name)Comp_Name from
                        (SELECT TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.No_Of_Cans,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift_Date,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Bulk_Route_Code,TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Dock_Collection_Milk_Type,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.FAT/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as FATKG
,case when isnull(TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF,0) =0 then 0 else cast(cast(cast(((TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Milk_Weight*TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.SNF/100)*1000) as int) as decimal(18,0))/1000 as decimal(18,3)) end as SNFKG,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader_Code],TSPL_MCC_MASTER.MCC_NAME As MCC,TSPL_COMPANY_MASTER.Comp_Name 
                        FROM TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL 
                        Left Outer Join TSPL_MILK_PROCUREMENT_UPLOADER_HEAD On TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.VLC_Code  
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.MCC_Code
                        Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'
                        where Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_HEAD.Document_Date,103)<=convert (date,'" & dtpToDate.Value & "',103)  "
                If AreaWiseBilling Then
                    If clsCommon.myLen(fndArea.Value) > 0 Then
                        qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                    End If
                Else
                    If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                        qry += " and TSPL_MCC_MASTER.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                    End If
                End If
                'If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                '    qry += " and TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(txtMCC_Code.Value) + "'"
                'End If
                If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='M'"
                ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Shift='E'"
                End If
                qry += " ) xxx Group By xxx.Shift_Date,xxx.Shift ORDER BY xxx.Shift_Date"
            ElseIf rbtnBMCDock.Checked Then
                qry = "Select '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [From Date],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [To Date],'" & objCommonVar.CurrentUser & "' as User_Name,Max(Convert(Date,xxx.Shift_Date,103))Shift_Date,Case When xxx.Shift='M' Then 'Morning' Else 'Evening' End As Shift,"
                If AreaWiseBilling Then
                    qry += " Max(xxx.Location_Desc)MCC_NAME,"
                Else
                    qry += " Max(xxx.MCC_NAME)MCC_NAME,"
                End If
                qry += "Max(xxx.VLC)VLC,'" & Area & "' as Area,Max(xxx.[VLC Name])[VLC Name],Max(xxx.Route)Route,(Convert(int,xxx.[Route Code]))[Route Code],Sum(xxx.[No of Cans])[No of Cans],Sum(xxx.[Good Qty])[Good Qty],Sum(xxx.[Good FAT %])[Good FAT %],Sum(xxx.[Good FATKg])[Good FATKg],Sum(xxx.[Good SNF %])[Good SNF %],Sum(xxx.[Good SNFKG])[Good SNFKG],Sum(xxx.[SOUR Qty])[SOUR Qty],Sum(xxx.[SOUR FAT %])[SOUR FAT %],Sum(xxx.[SOUR FATKg])[SOUR FATKg],Sum(xxx.[SOUR SNF %])[SOUR SNF %],Sum(xxx.[SOUR SNFKG])[SOUR SNFKG],Sum(xxx.[CURD Qty])[CURD Qty],Sum([CURD FAT %])[CURD FAT %],Sum(xxx.[CURD FATKg])[CURD FATKg],Sum(xxx.[CURD SNF %])[CURD SNF %],Sum(xxx.[CURD SNFKG])[CURD SNFKG],Max(xxx.Comp_Name)Comp_Name
                        from (select TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No,"
                If AreaWiseBilling Then
                    qry += " xxxSetLocation.Location_Desc,"
                End If
                qry += "TSPL_MILK_SHIFT_UPLOADER_DETAIL.SNo,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC, TSPL_MILK_SHIFT_UPLOADER_DETAIL.VLC_Code as [VLC Code],TSPL_VLC_MASTER_HEAD.VLC_Name as [VLC Name],TSPL_MILK_SHIFT_UPLOADER_DETAIL.No_Of_Cans as [No of Cans],TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO as [Route Code],TSPL_BULK_ROUTE_MASTER.ROUTE_NAME as [Route],TSPL_COMPANY_MASTER.Comp_Name,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date
                        ,case When isnull(Reject_Type,'')='' then Milk_Weight else 0 end as [Good Qty]
                        ,case When isnull(Reject_Type,'')='' then FAT else 0 end as [Good FAT %]
                        ,case When isnull(Reject_Type,'')='' then cast(Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [Good FATKg]
                        ,case When isnull(Reject_Type,'')='' then SNF else 0 end as [Good SNF %]
                        ,case When isnull(Reject_Type,'')='' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [Good SNFKG],case When isnull(Reject_Type,'')='SOUR' then Milk_Weight else 0 end as [SOUR Qty]
                        ,case When isnull(Reject_Type,'')='SOUR' then FAT else 0 end as [SOUR FAT %]
                        ,case When isnull(Reject_Type,'')='SOUR' then cast (Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [SOUR FATKg]
                        ,case When isnull(Reject_Type,'')='SOUR' then SNF else 0 end as [SOUR SNF %]
                        ,case When isnull(Reject_Type,'')='SOUR' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [SOUR SNFKG],case When isnull(Reject_Type,'')='CURD' then Milk_Weight else 0 end as [CURD Qty]
                        ,case When isnull(Reject_Type,'')='CURD' then FAT else 0 end as [CURD FAT %]
                        ,case When isnull(Reject_Type,'')='CURD' then cast (Milk_Weight*FAT/100 as decimal(18,3)) else 0 end as [CURD FATKg]
                        ,case When isnull(Reject_Type,'')='CURD' then SNF else 0 end as [CURD SNF %]
                        ,case When isnull(Reject_Type,'')='CURD' then cast (Milk_Weight*SNF/100 as decimal(18,3)) else 0 end as [CURD SNFKG] ,TSPL_MILK_SHIFT_UPLOADER_DETAIL.PageNo 
                        from TSPL_MILK_SHIFT_UPLOADER_DETAIL 
                        left outer Join TSPL_MILK_SHIFT_UPLOADER_HEAD On TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No=TSPL_MILK_SHIFT_UPLOADER_DETAIL.Document_No
                        left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_SHIFT_UPLOADER_DETAIL.VLC_Code
                        left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SHIFT_UPLOADER_DETAIL.BULK_ROUTE_NO
                        Left Outer Join TSPL_MCC_MASTER ON TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC "
                If AreaWiseBilling Then
                    qry += " Left Outer Join( select TSPL_PAYMENT_PROCESS_HEAD.Doc_No,tspl_location_master.Location_Desc,tspl_location_master.Location_Code   From TSPL_PAYMENT_PROCESS_HEAD left  join tspl_location_master on tspl_location_master.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code)  xxxSetLocation On xxxSetLocation.Location_Code=TSPL_MCC_MASTER.area_Location_code "
                End If
                qry += "  Left outer join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & clsCommon.myCstr(objCommonVar.CurrComp_Code1) & "'                        
                        where Convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date,103)>=convert(date,'" & fromDate.Value & "',103)  And Convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date,103)<=convert (date,'" & dtpToDate.Value & "',103) "
                If AreaWiseBilling AndAlso clsCommon.myLen(fndArea.Value) > 0 Then
                    'If clsCommon.myLen(fndArea.Value) > 0 Then
                    qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                    'End If
                Else
                    If txtMCC_Code.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC_Code.Value) > 0 Then
                        qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.MCC_Code='" & clsCommon.myCstr(txtMCC_Code.Value) & "'"
                    End If
                End If
                If ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Morning") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift='M'"
                ElseIf ddlShift.SelectedValue IsNot Nothing AndAlso clsCommon.CompairString(ddlShift.SelectedValue, "Evening") = CompairStringResult.Equal Then
                    qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift='E'"
                End If
                qry += ") xxx Group By xxx.[Route Code],xxx.Shift_Date,xxx.shift order by Convert(int,xxx.[Route Code]),xxx.shift desc"
            Else
                qry = " select '" & fromDate.Value & "' As [From Date],'" & dtpToDate.Value & "' As [To Date],'" & objCommonVar.CurrentUser & "' as User_Name, XXGetAllRecords.Document_No,XXGetAllRecords.Document_Date,XXGetAllRecords.Route_Code,XXGetAllRecords.ROUTE_NAME,XXGetAllRecords.city_code,XXGetAllRecords.Tanker_No,XXGetAllRecords.Comp_Name,XXGetAllRecords.Vehicle_No,
                                XXGetAllRecords.Entered_Qty,XXGetAllRecords.Entered_FATKg,XXGetAllRecords.Entered_SNFKg,
                                XXGetAllRecords.MCC_Qty,XXGetAllRecords.MCC_FATKG,XXGetAllRecords.MCC_SNFKG,
                                XXGetAllRecords.DiffMCCVsEntered_Qty,
                                XXGetAllRecords.DiffMCCVsEntered_FATKG,0 as FAT_Tolerence,case when (isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0) end as FATKG_Recovered,Round(XXGetAllRecords.FAT_AMT,2,0) as FAT_AMT,
                                XXGetAllRecords.DiffMCCVsEntered_SNFKG,0 as SNF_Tolerence,case when (isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0) end as SNFKG_Recovered,Round(XXGetAllRecords.SNF_AMT,2,0) as SNF_AMT,Round((XXGetAllRecords.FAT_AMT + XXGetAllRecords.SNF_AMT),2,0) as AMOUNT,XXGetAllRecords.GainLossCode as GainLoss_Code,
                                   XXGetAllRecords.Loss_FAT_Rate as Loss_FAT_Rate,
                                   XXGetAllRecords.Loss_SNF_Rate as Loss_SNF_Rate,
                                   XXGetAllRecords.Start_Date as Start_Date
                                 from( select GetAllGainLossRate.*,
                                              case when GetAllGainLossRate.DiffMCCVsEntered_FATKG >= 0 then (
                                        case when (GetAllGainLossRate.DiffMCCVsEntered_FATKG - 0) > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_FATKG -0) *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate) else 0 end)
                                        else 0 end as FAT_AMT,
                                        case when GetAllGainLossRate.DiffMCCVsEntered_SNFKG > 0 then (case when (GetAllGainLossRate.DiffMCCVsEntered_SNFKG - 0) > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_SNFKG -0) *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate) else 0 end)
                                      else 0 end as SNF_AMT,
                                      TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date from ( select GetRateCode.*,( SELECT top 1 TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, GetRateCode.Document_Date, 103) order by TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date desc ) as GainLossCode from ( SELECT XXXFinal.Document_No, MAX(XXXFinal.Document_Date) AS Document_Date, MAX(XXXFinal.Route_Code) AS Route_Code, MAX(XXXFinal.ROUTE_NAME) AS ROUTE_NAME,MAX(XXXFinal.City_code) AS City_code, MAX(XXXFinal.Tanker_No) AS Tanker_No,MAX(XXXFinal.Comp_Name) AS Comp_Name, MAX(XXXFinal.Vehicle_No) AS Vehicle_No, MAX(Entered_Qty) AS Entered_Qty, ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_FAT, ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_SNF, MAX(Entered_FATKg) AS Entered_FATKg, MAX(Entered_SNFKg) AS Entered_SNFKg, SUM(XXXFinal.MCC_Qty) AS MCC_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_SNF, SUM(MCC_FATKG) AS MCC_FATKG, SUM(MCC_SNFKG) AS MCC_SNFKG, SUM(DCS_Qty) AS DCS_Qty, ROUND(((SUM(DCS_FATKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_FAT, ROUND(((SUM(DCS_SNFKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_SNF, SUM(DCS_FATKG) AS DCS_FATKG, SUM(DCS_SNFKG) AS DCS_SNFKG, SUM(Diff_Qty) AS Diff_Qty, MAX(DCS_FAT) - ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_FAT, MAX(DCS_SNF) - ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_SNF, SUM(Diff_FATKG) AS Diff_FATKG, SUM(Diff_SNFKG) AS Diff_SNFKG, SUM(XXXFinal.MCC_Qty)- MAX(Entered_Qty) AS DiffMCCVsEntered_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_SNF, SUM(MCC_FATKG)-MAX(Entered_FATKg) AS DiffMCCVsEntered_FATKG, SUM(MCC_SNFKG)-MAX(Entered_SNFKg) AS DiffMCCVsEntered_SNFKG from ( select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id, Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME,max(City_code) as City_code, max(Tanker_No) as Tanker_No,max(Comp_Name) as Comp_Name, max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, max(DCS_Qty) as DCS_Qty, isnull( (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_FAT, isnull( (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_SNF, isnull(max(DCS_FATKG),0) as DCS_FATKG,isnull( max(DCS_SNFKG),0) as DCS_SNFKG, max(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_FAT, (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_SNF, sum(MCC_FATKG) -max(DCS_FATKG) as Diff_FATKG, sum(MCC_SNFKG)-max(DCS_SNFKG) as Diff_SNFKG from ( select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo, TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME,TSPL_COMPANY_MASTER.City_code, TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_COMPANY_MASTER.Comp_Name,TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG from TSPL_MILK_COLLECTION_MCC_DETAIL left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_BULK_ROUTE_MASTER.Comp_Code left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) /nullif( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / nullif(sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as SNF
                                                         from TSPL_MILK_COLLECTION_DCS_DETAIL 
                                        left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                            left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                            where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" & fromDate.Value & "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" & dtpToDate.Value & "',103) 
                            group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                            where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" & fromDate.Value & "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" & dtpToDate.Value & "',103)) xyz group by UploaderNo, Document_No    ) XXXFinal group by XXXFinal.Document_No )GetRateCode )GetAllGainLossRate left outer join TSPL_OWN_BMC_GAIN_LOSS_RATE on TSPL_OWN_BMC_GAIN_LOSS_RATE.Code=GetAllGainLossRate.GainLossCode ) XXGetAllRecords
    						 "
                If clsCommon.myLen(TxtTankerNo.Value) > 0 Then
                    qry += " where Tanker_No='" & TxtTankerNo.Value & "'"
                End If
                If rbtnTranpoterGainLoss.Checked Then
                    qry += " order by convert(date,XXGetAllRecords.Document_Date,103) asc "
                End If
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If rbtnDockSummary.Checked Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptDockSummaryReport", "DCS Summary Report", clsCommon.myCDate(fromDate.Value))
                ElseIf rbtnSummary.Checked Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptDSummary", "Dock Summary", clsCommon.myCDate(fromDate.Value))
                ElseIf rbtnBMCDock.Checked Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptBMCDockReport", "BMC Dock Report", clsCommon.myCDate(fromDate.Value))
                ElseIf rbtnBMCDCSPrint.Checked Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptBMCDCSPrint", "BMC DCS Report", clsCommon.myCDate(fromDate.Value))
                ElseIf rbtnBMCRoutePrint.Checked Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptBMCRoutePrint", "BMC Route Report", clsCommon.myCDate(fromDate.Value))
                Else
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptdailyqtyreport", "Daily Quantity Report", clsCommon.myCDate(fromDate.Value))
                End If
                frmCRV = Nothing
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found.", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnBMCDock_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnBMCDock.CheckedChanged
        Try
            If rbtnBMCDock.Checked Then
                btnPrint.Visible = True
            Else
                btnPrint.Visible = False
            End If
            If rbtnBMCDock.Checked Then
                MyLabel4.Visible = True
                txtMCC_Code.Visible = True
                MyLabel6.Visible = True
                ddlShift.Visible = True
                If AreaWiseBilling Then
                    fndArea.Visible = True
                    lblArea.Visible = True
                End If
            Else
                MyLabel4.Visible = False
                txtMCC_Code.Visible = False
                ddlShift.Visible = False
                MyLabel6.Visible = False
                fndArea.Visible = False
                lblArea.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnDockSummary_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnDockSummary.CheckedChanged
        Try
            If rbtnDockSummary.Checked Then
                btnPrint.Visible = True
            Else
                btnPrint.Visible = False
            End If
            If rbtnDockSummary.Checked Then
                MyLabel4.Visible = True
                txtMCC_Code.Visible = True
                ddlShift.Visible = True
                MyLabel6.Visible = True
                If AreaWiseBilling Then
                    fndArea.Visible = True
                    lblArea.Visible = True
                End If
            Else
                MyLabel4.Visible = False
                txtMCC_Code.Visible = False
                ddlShift.Visible = False
                MyLabel6.Visible = False
                fndArea.Visible = False
                lblArea.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadShif()
        ddlShift.DataSource = Shift()
        ddlShift.ValueMember = "Code"
        ddlShift.DisplayMember = "Name"
        ddlShift.SelectedIndex = 0
    End Sub
    Public Function Shift() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "Both"
        dr("Name") = "Both"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Morning"
        dr("Name") = "Morning"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "Evening"
        dr("Name") = "Evening"
        dt.Rows.Add(dr)
        Return dt
    End Function
    Private Sub rbtnSummary_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnSummary.CheckedChanged
        Try
            If rbtnSummary.Checked Then
                btnPrint.Visible = True
            Else
                btnPrint.Visible = False
            End If
            If rbtnSummary.Checked Then
                MyLabel4.Visible = True
                txtMCC_Code.Visible = True
                ddlShift.Visible = True
                MyLabel6.Visible = True
                If AreaWiseBilling Then
                    fndArea.Visible = True
                    lblArea.Visible = True
                End If
            Else
                MyLabel4.Visible = False
                txtMCC_Code.Visible = False
                ddlShift.Visible = False
                MyLabel6.Visible = False
                fndArea.Visible = False
                lblArea.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub fndArea__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndArea._MYValidating
        Try
            Dim sQuery As String = " Select TSPL_LOCATION_MASTER.Location_Code as Code ,  TSPL_LOCATION_MASTER.Location_Desc, Type from TSPL_LOCATION_MASTER "
            fndArea.Value = clsCommon.ShowSelectForm("Location@Plant@Master", sQuery, "Code", "TSPL_LOCATION_MASTER.Type <> 'PLANT' OR TSPL_LOCATION_MASTER.Location_Category <> 'Mcc'", fndArea.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.ToString, Me.Text)
        End Try
    End Sub
    Private Sub rbtnBMCDCSPrint_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnBMCDCSPrint.CheckedChanged
        Try
            If rbtnBMCDCSPrint.Checked Then

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    MyLabel4.Visible = False
                    txtMCC_Code.Visible = False
                End If
                MyLabel4.Visible = False
                cboItemType.Visible = False
                MyLabel39.Visible = False
                'txtMCC_Code.Visible = True
                MyLabel6.Visible = True
                ddlShift.Visible = True
                If AreaWiseBilling Then
                    fndArea.Visible = True
                    lblArea.Visible = True
                End If
            Else
                MyLabel4.Visible = False
                txtMCC_Code.Visible = False
                ddlShift.Visible = False
                MyLabel6.Visible = False
                fndArea.Visible = False
                lblArea.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnBMCRoutePrint_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnBMCRoutePrint.CheckedChanged
        Try
            If rbtnBMCRoutePrint.Checked Then
                MyLabel4.Visible = False
                'txtMCC_Code.Visible = True
                MyLabel6.Visible = True
                ddlShift.Visible = True
                If AreaWiseBilling Then
                    fndArea.Visible = True
                    lblArea.Visible = True
                End If
            Else
                MyLabel4.Visible = False
                txtMCC_Code.Visible = False
                ddlShift.Visible = False
                MyLabel6.Visible = False
                fndArea.Visible = False
                lblArea.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnRouteWise_Click(sender As Object, e As EventArgs) Handles rbtnRouteWise.Click
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
            btnRoutePrint.Visible = True
        End If
    End Sub
    Private Sub btnRoutePrint_Click(sender As Object, e As EventArgs) Handles btnRoutePrint.Click
        Try
            Dim qry As String = Nothing
            Dim Area As String = Nothing
            Area = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MCC_NAME from TSPL_MCC_MASTER WHERE Area_Location_Code = '" & fndArea.Value & "' "))
            qry = "  SELECT '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MM/yyyy") & "' As [FromDate],'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd/MM/yyyy") & "' As [ToDate],'" & objCommonVar.CurrentUser & "' as User_Name,
                          max(Comp_Name)Comp_Name,max(add1)add1,max(Add2)Add2,max(Add3)Add3,CONVERT (VARCHAR, Document_Date,103) AS Date , ROUTE_NO, max(ROUTE_NAME)ROUTE_NAME,  Sum (MCC_Qty) AS MCC_Qty, convert(decimal(18,2),(sum(MCC_FATKG)*100)/sum(MCC_Qty)) MCC_FAT_Per, convert(decimal(18,2),(sum(MCC_SNFKG)*100)/sum(MCC_Qty)) MCC_SNF_Per,
                          sum (MCC_FATKG) AS MCC_FATKG, sum (MCC_SNFKG) AS MCC_SNFKG, sum(entered_qty) AS Entered_Qty, convert(decimal(18,2),(sum(Entered_FATKg)*100)/sum(Entered_Qty)) Entered_FAT_Per,convert(decimal(18,2),(sum(Entered_SNFKg)*100)/sum(Entered_Qty)) Entered_SNF_Per, Sum (entered_fatkg)  AS Entered_FATKg,
                          Sum (entered_snfkg) AS Entered_SNFKg, Sum(MCC_Qty) - Sum(Entered_Qty) AS DiffEnteredVsMCC_Qty, Sum (MCC_FATKG) - Sum (Entered_FATKg)  AS DiffEnteredVsMCC_FAT, Sum (MCC_SNFKG)- Sum (Entered_SNFKg) AS DiffEnteredVsMCC_SNF,max(RejectInfo)RejectInfo 
						  FROM
                          ( SELECT max(Comp_Name)Comp_Name,max(Add1)Add1,max(Add2)Add2,max(Add3)Add3,ROUTE_NO,max(ROUTE_NAME)ROUTE_NAME,  sum(entered_qty) AS Entered_Qty, Isnull(sum(entered_fatkg),0) AS Entered_FATKg, Isnull(sum(entered_snfkg),0) AS Entered_SNFKg, (document_date) AS Document_Date,	 
                          max(MCC_Qty) AS MCC_Qty, Isnull(max(MCC_FATKG), 0) AS MCC_FATKG, Isnull(max(MCC_SNFKG), 0) AS MCC_SNFKG,MAX(Reject_Type + '(' + CAST(Reject_Recovery_Per AS VARCHAR) + '%)') AS RejectInfo FROM 
						  (SELECT Tspl_Company_Master.Comp_Name,Tspl_Company_Master.Add1,Tspl_Company_Master.Add2,Tspl_Company_Master.Add3,Trip_No, tspl_bulk_route_master.ROUTE_NO,tspl_bulk_route_master.ROUTE_NAME, ISNULL(qty,0) as entered_qty , FATKG as entered_fatkg, SNFKG as entered_snfkg,
                          tspl_milk_collection_mcc_detail.document_no, tspl_milk_collection_mcc.document_date
                          AS Document_Date, Isnull(tspl_milk_collection_mcc.Entered_Qty, 0) AS MCC_Qty, Isnull(fat, 0) AS MCC_FAT, ISNULL(SNF, 0) AS MCC_SNF, ISNULL(tspl_milk_collection_mcc.Entered_FATKg, 0) AS MCC_FATKG, 
                          ISNULL(tspl_milk_collection_mcc.Entered_SNFKg, 0) AS MCC_SNFKG ,tspl_milk_collection_mcc.Reject_Type,tspl_milk_collection_mcc.Reject_Recovery_Per FROM   tspl_milk_collection_mcc_detail
                          LEFT OUTER JOIN tspl_milk_collection_mcc ON tspl_milk_collection_mcc.document_no = tspl_milk_collection_mcc_detail.document_no 
                          LEFT OUTER JOIN tspl_mcc_master ON tspl_mcc_master.mcc_code = tspl_milk_collection_mcc_detail.mcc_code 
                          LEFT OUTER JOIN tspl_bulk_route_master ON tspl_bulk_route_master.route_no = tspl_milk_collection_mcc.route_code
						  left outer join Tspl_Company_Master on Tspl_Company_Master.Comp_Code = tspl_mcc_master.Comp_Code
                           WHERE  CONVERT(DATE, tspl_milk_collection_mcc.document_date, 103) >= CONVERT(DATE, '" & fromDate.Value & "',103) AND CONVERT (DATE, tspl_milk_collection_mcc.document_date, 103)
                         <= CONVERT (DATE, '" & dtpToDate.Value & "',103) "
            If AreaWiseBilling AndAlso clsCommon.myLen(fndArea.Value) > 0 Then
                'If clsCommon.myLen(fndArea.Value) > 0 Then
                qry += " And TSPL_MCC_MASTER.Area_Location_Code = '" & fndArea.Value & "' "
                'End If
            End If
            qry += " ) xyz GROUP  BY ROUTE_NO, Document_Date,Trip_No) XXXFinal GROUP  BY  xxxfinal.ROUTE_NO , xxxfinal.Document_Date
						   order by xxxfinal.ROUTE_NO "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptRouteWisereport", "Route Wise", clsCommon.myCDate(fromDate.Value))
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnBmcSummary_Click(sender As Object, e As EventArgs) Handles rbtnBmcSummary.Click
        If rbtnBmcSummary.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        lblToleranceFAT.Visible = False
        btnGainLoss.Visible = False
    End Sub
    Private Sub rdbTankerWise_Click(sender As Object, e As EventArgs) Handles rdbTankerWise.Click
        If rdbTankerWise.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        lblToleranceFAT.Visible = False
        btnGainLoss.Visible = False
    End Sub
    Private Sub rdbMultiple_Click(sender As Object, e As EventArgs) Handles rdbMultiple.Click
        If rdbMultiple.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        btnGainLoss.Visible = False
    End Sub
    Private Sub rdbCollectionWise_Click(sender As Object, e As EventArgs) Handles rdbCollectionWise.Click
        If rdbCollectionWise.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        lblToleranceFAT.Visible = False
        btnGainLoss.Visible = False
    End Sub
    Private Sub rbtnBMCTankerCollection_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnBMCTankerCollection.CheckedChanged
        If rbtnBMCTankerCollection.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
        lblToleranceFAT.Visible = False
        btnGainLoss.Visible = False
    End Sub
    Private Sub rbtnRouteWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnRouteWise.CheckedChanged
        If rbtnRouteWise.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
    End Sub
    Private Sub rbtnDockDateWise_Click(sender As Object, e As EventArgs) Handles rbtnDockDateWise.Click
        If rbtnDockDateWise.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
    End Sub
    Private Sub rbtnDockShiftWise_Click(sender As Object, e As EventArgs) Handles rbtnDockShiftWise.Click
        If rbtnDockShiftWise.Checked Then
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If
    End Sub
    Private Sub rbtnTranspoterGainlossSummary_Click(sender As Object, e As EventArgs) Handles rbtnTranspoterGainlossSummary.Click
        If rbtnTranspoterGainlossSummary.Checked Then
            btnPrint.Visible = True
        Else
            btnPrint.Visible = False
        End If
        MyLabel4.Visible = False
        lblToleranceFAT.Visible = False
        txtMCC_Code.Visible = False
        MyLabel5.Visible = True
        TxtTankerNo.Visible = True
        fndArea.Visible = False
        lblArea.Visible = False
        TranspoterBoxhandler(True)
    End Sub
    Private Sub rbtnTranspoterGainlossSummary_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnTranspoterGainlossSummary.CheckedChanged
        If rbtnTranspoterGainlossSummary.Checked Then
            MyLabel5.Visible = True
            TxtTankerNo.Visible = True
            txtToleranceFat.Visible = True
            TxtTankerNo.Value = ""
            'txtTolerenceSNF.Value = 0
            txtTolerenceSNF.Visible = True
            MyLabel7.Visible = True
            btnGainLoss.Visible = True
        Else
            MyLabel5.Visible = False
            TxtTankerNo.Visible = False
            TxtTankerNo.Value = ""
            fndArea.Visible = False
            lblArea.Visible = False
            txtTolerenceSNF.Value = 0
            txtToleranceFat.Value = 0
            txtTolerenceSNF.Visible = False
            MyLabel7.Visible = False
            txtToleranceFat.Visible = False
            btnGainLoss.Visible = False
        End If
    End Sub

    Private Sub btnGainLoss_Click(sender As Object, e As EventArgs) Handles btnGainLoss.Click
        Try
            Dim Qry As String = ""
            Dim dt As DataTable = Nothing
            If rbtnTranspoterGainlossSummary.Checked Then
                Qry = " select max(PPP.Document_No)Document_No,
                        max(PPP.Document_Date)Document_Date,
                        (PPP.Route_Code)Route_Code,
                        (PPP.ROUTE_NAME)ROUTE_NAME,
                        max(PPP.Tanker_No)Tanker_No,
                        max(PPP.Vehicle_No)Vehicle_No,
                        sum(PPP.Entered_Qty)Entered_Qty,
						sum(PPP.Entered_FATKg)Entered_FATKg,
						sum(PPP.Entered_SNFKg)Entered_SNFKg,
                        sum(PPP.MCC_Qty)MCC_Qty,
						sum(PPP.MCC_FATKG)MCC_FATKG,
						sum(PPP.MCC_SNFKG)MCC_SNFKG,
                        sum(PPP.DiffMCCVsEntered_Qty)DiffMCCVsEntered_Qty,
						sum(ppp.DiffMCCVsEntered_FATKG)DiffMCCVsEntered_FATKG,
                          sum(ppp.FAT_Tolerence) as FAT_Tolerence, case when sum(ppp.FATKG_Recovered)<=0 then 0 else sum(ppp.FATKG_Recovered) end   as FATKG_Recovered
						  , case when sum(ppp.FATKG_Recovered)<=0 then 0 else  Sum(FAT_AMT) end as FAT_AMT,
                        sum(PPP.DiffMCCVsEntered_SNFKG)DiffMCCVsEntered_SNFKG,sum(SNF_Tolerence) as SNF_Tolerence,
						  case when Sum(ppp.SNFKG_Recovered)<= 0 then 0 else Sum(ppp.SNFKG_Recovered) end as SNFKG_Recovered, 
                          Case when Sum(ppp.SNF_AMT)<= 0 then 0 else Sum(ppp.SNF_AMT) end as SNF_AMT, 
                          CASE WHEN SUM(FAT_AMT) >= 0 
                          AND SUM(SNF_AMT) >= 0 THEN ISNULL(
                            SUM(FAT_AMT), 
                            0
                          ) + ISNULL(
                            SUM(SNF_AMT), 
                            0
                          ) WHEN SUM(FAT_AMT) >= 0 
                          AND SUM(SNF_AMT) <= 0 THEN SUM(FAT_AMT) WHEN SUM(FAT_AMT) <= 0 
                          AND SUM(SNF_AMT) >= 0 THEN SUM(SNF_AMT) ELSE 0 END AS AMOUNT
                        ,max(PPP.GainLoss_Code) as GainLoss_Code,
                           sum(PPP.Loss_FAT_Rate) as Loss_FAT_Rate,
                           sum(PPP.Loss_SNF_Rate) as Loss_SNF_Rate,
                           max(PPP.Start_Date) as Start_Date,max(Comp_Name)Comp_Name,max(Add1)Add1,max(Add2)Add2
						   from (select  max(XXGetAllRecords.Document_No)Document_No,
                        max(XXGetAllRecords.Document_Date)Document_Date,
                        XXGetAllRecords.Route_Code,
                        XXGetAllRecords.ROUTE_NAME,
                        max(XXGetAllRecords.Tanker_No)Tanker_No,
                        max(XXGetAllRecords.Vehicle_No)Vehicle_No,
                        sum(XXGetAllRecords.Entered_Qty)Entered_Qty,
						sum(XXGetAllRecords.Entered_FATKg)Entered_FATKg,
						sum(XXGetAllRecords.Entered_SNFKg)Entered_SNFKg,
                        sum(XXGetAllRecords.MCC_Qty)MCC_Qty,
						sum(XXGetAllRecords.MCC_FATKG)MCC_FATKG,
						sum(XXGetAllRecords.MCC_SNFKG)MCC_SNFKG,
                        sum(XXGetAllRecords.DiffMCCVsEntered_Qty)DiffMCCVsEntered_Qty,
                          sum(isnull(XXGetAllRecords.Entered_FATKg,0)-isnull(XXGetAllRecords.MCC_FATKG,0))DiffMCCVsEntered_FATKG,sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100 as FAT_Tolerence,case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100  end as FATKG_Recovered,max(XXGetAllRecords.Loss_FAT_Rate)*case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtToleranceFat.Value) & "/100 end  as FAT_AMT,
                        sum(XXGetAllRecords.DiffMCCVsEntered_SNFKG)DiffMCCVsEntered_SNFKG,sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100 as SNF_Tolerence,case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100 end as SNFKG_Recovered,max(XXGetAllRecords.Loss_SNF_Rate)*case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100  end as SNF_AMT,max(XXGetAllRecords.Loss_FAT_Rate)*case when sum((isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0)))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_FATKG - 0,0))-sum(XXGetAllRecords.MCC_FATKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100 end + max(XXGetAllRecords.Loss_SNF_Rate)*case when sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))<0 then 0 else sum(isnull(XXGetAllRecords.DiffMCCVsEntered_SNFKG - 0,0))-sum(XXGetAllRecords.MCC_SNFKG)* " & clsCommon.myCstr(txtTolerenceSNF.Value) & "/100  end as AMOUNT,max(XXGetAllRecords.GainLossCode) as GainLoss_Code,
                           sum(XXGetAllRecords.Loss_FAT_Rate) as Loss_FAT_Rate,
                           sum(XXGetAllRecords.Loss_SNF_Rate) as Loss_SNF_Rate,
                           max(XXGetAllRecords.Start_Date) as Start_Date,max(Comp_Name)Comp_Name,max(Add1)Add1,max(Add2)Add2
                         from
						 ( select GetAllGainLossRate.*,
                         case when GetAllGainLossRate.DiffMCCVsEntered_FATKG >= 0 then (
                         case when (GetAllGainLossRate.DiffMCCVsEntered_FATKG - " & clsCommon.myCstr(txtToleranceFat.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_FATKG -" & clsCommon.myCstr(txtToleranceFat.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate) else 0 end)
                         else 0 end as FAT_AMT,case when GetAllGainLossRate.DiffMCCVsEntered_SNFKG > 0 then (case when (GetAllGainLossRate.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtTolerenceSNF.Value) & ") > 0 then ((GetAllGainLossRate.DiffMCCVsEntered_SNFKG - " & clsCommon.myCstr(txtTolerenceSNF.Value) & ") *TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate) else 0 end) else 0 end as SNF_AMT,
                         TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Gain_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_FAT_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Loss_SNF_Rate, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date from 
						 ( select GetRateCode.*,( SELECT top 1 TSPL_OWN_BMC_GAIN_LOSS_RATE.Code FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE TSPL_OWN_BMC_GAIN_LOSS_RATE.Posted=1 and TSPL_OWN_BMC_GAIN_LOSS_RATE.Inactive=0 and CONVERT(date, TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date, 103) <= CONVERT(date, GetRateCode.Document_Date, 103) order by TSPL_OWN_BMC_GAIN_LOSS_RATE.Start_Date desc ) as GainLossCode from 
						 ( SELECT XXXFinal.Document_No, MAX(XXXFinal.Document_Date) AS Document_Date, MAX(XXXFinal.Route_Code) AS Route_Code, MAX(XXXFinal.ROUTE_NAME) AS ROUTE_NAME, MAX(XXXFinal.Tanker_No) AS Tanker_No, MAX(XXXFinal.Vehicle_No) AS Vehicle_No, MAX(Entered_Qty) AS Entered_Qty, ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_FAT, ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS Entered_SNF, MAX(Entered_FATKg) AS Entered_FATKg, MAX(Entered_SNFKg) AS Entered_SNFKg, SUM(XXXFinal.MCC_Qty) AS MCC_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS MCC_SNF, SUM(MCC_FATKG) AS MCC_FATKG, SUM(MCC_SNFKG) AS MCC_SNFKG, SUM(DCS_Qty) AS DCS_Qty, ROUND(((SUM(DCS_FATKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_FAT, ROUND(((SUM(DCS_SNFKG) / NULLIF(SUM(DCS_Qty), 0)) * 100), 3, 0) AS DCS_SNF, SUM(DCS_FATKG) AS DCS_FATKG, SUM(DCS_SNFKG) AS DCS_SNFKG, SUM(Diff_Qty) AS Diff_Qty, MAX(DCS_FAT) - ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_FAT, MAX(DCS_SNF) - ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0) AS Diff_SNF, SUM(Diff_FATKG) AS Diff_FATKG, SUM(Diff_SNFKG) AS Diff_SNFKG, SUM(XXXFinal.MCC_Qty)- MAX(Entered_Qty) AS DiffMCCVsEntered_Qty, ROUND(((SUM(MCC_FATKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_FATKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_FAT, ROUND(((SUM(MCC_SNFKG) / NULLIF(SUM(MCC_Qty), 0)) * 100), 3, 0)- ((MAX(Entered_SNFKg) / NULLIF(MAX(Entered_Qty), 0)) * 100) AS DiffMCCVsEntered_SNF, SUM(MCC_FATKG)-MAX(Entered_FATKg) AS DiffMCCVsEntered_FATKG, SUM(MCC_SNFKG)-MAX(Entered_SNFKg) AS DiffMCCVsEntered_SNFKG,max(Comp_Name)Comp_Name,max(Add1)Add1,max(Add2)Add2 from ( select max(Entered_Qty) as Entered_Qty, max(Entered_FATKg) as Entered_FATKg, max(Entered_SNFKg) as Entered_SNFKg, max(PK_Id) as PK_Id, Document_No, max(Document_Date) as Document_Date, max(MCC_Code) as MCC_Code, max(MCC_NAME) as MCC_NAME, max(UploaderNo) as UploaderNo, max(Route_Code) as Route_Code, max(ROUTE_NAME) as ROUTE_NAME, max(Tanker_No) as Tanker_No,max(Vehicle_No) as Vehicle_No, sum(MCC_Qty) as MCC_Qty, ROUND(((sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_FAT, ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as MCC_SNF, sum(MCC_FATKG) as MCC_FATKG, sum(MCC_SNFKG) as MCC_SNFKG, max(DCS_Qty) as DCS_Qty, isnull( (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_FAT, isnull( (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100),0) as DCS_SNF, isnull(max(DCS_FATKG),0) as DCS_FATKG,isnull( max(DCS_SNFKG),0) as DCS_SNFKG, max(DCS_Qty) - sum(MCC_Qty) as Diff_Qty, (max(DCS_FATKG)/ nullif(max(DCS_Qty),0)*100)- ROUND( ( ( sum (MCC_FATKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_FAT, (max(DCS_SNFKG)/ nullif(max(DCS_Qty),0)*100)-ROUND( ( ( sum (MCC_SNFKG) / nullif ( sum (MCC_Qty), 0 ) ) * 100 ), 3, 0 ) as Diff_SNF, sum(MCC_FATKG) -max(DCS_FATKG) as Diff_FATKG, sum(MCC_SNFKG)-max(DCS_SNFKG) as Diff_SNFKG,max(Comp_Name)Comp_Name,max(Add1)Add1,max(Add2)Add2 from 
						 ( select TSPL_MILK_COLLECTION_MCC.Entered_Qty, TSPL_MILK_COLLECTION_MCC.Entered_FATKg, TSPL_MILK_COLLECTION_MCC.Entered_SNFKg, TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id, TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No, convert ( varchar, TSPL_MILK_COLLECTION_MCC.Document_Date, 103 ) as Document_Date, TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code, TSPL_MCC_MASTER.MCC_NAME, TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as UploaderNo, TSPL_MILK_COLLECTION_MCC.Route_Code, TSPL_BULK_ROUTE_MASTER.ROUTE_NAME, TSPL_MILK_COLLECTION_MCC.Tanker_No,TSPL_MILK_COLLECTION_MCC.Vehicle_No, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty as MCC_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT as MCC_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF as MCC_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG as MCC_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG as MCC_SNFKG, isnull(XXXDCS.qty, 0) as DCS_Qty, isnull(XXXDCS.FAT, 0) as DCS_FAT, isnull(XXXDCS.SNF, 0) as DCS_SNF, isnull(XXXDCS.FATKG, 0) as DCS_FATKG, isnull(XXXDCS.SNFKG, 0) as DCS_SNFKG, TSPL_MILK_COLLECTION_MCC_DETAIL.Qty - isnull(XXXDCS.qty, 0) as Diff_Qty, TSPL_MILK_COLLECTION_MCC_DETAIL.FAT - isnull(XXXDCS.FAT, 0) as Diff_FAT, TSPL_MILK_COLLECTION_MCC_DETAIL.SNF - isnull(XXXDCS.SNF, 0) as Diff_SNF, TSPL_MILK_COLLECTION_MCC_DETAIL.FATKG - isnull(XXXDCS.FATKG, 0) as Diff_FATKG, TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKG - isnull(XXXDCS.SNFKG, 0) as Diff_SNFKG ,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2
						 from TSPL_MILK_COLLECTION_MCC_DETAIL left outer join TSPL_MILK_COLLECTION_MCC on TSPL_MILK_COLLECTION_MCC.Document_No = TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code = TSPL_MILK_COLLECTION_MCC_DETAIL.MCC_Code left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code 
						 left join TSPL_COMPANY_MASTER on 2=2 
						 left outer join ( select TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail, sum (qty) as qty, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) as FATKG, sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) as SNFKG, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.FATKG ) /nullif( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as FAT, ( ( sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.SNFKG ) / nullif(sum ( TSPL_MILK_COLLECTION_DCS_DETAIL.qty ),0) ) * 100 ) as SNF
                        from TSPL_MILK_COLLECTION_DCS_DETAIL 
                        left outer join TSPL_MILK_COLLECTION_DCS_MCC_DETAIL on TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No 
                        left outer join TSPL_MILK_COLLECTION_DCS on TSPL_MILK_COLLECTION_DCS.Document_No = TSPL_MILK_COLLECTION_DCS_DETAIL.Document_No
                        where convert(date, TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) 
                        group by TSPL_MILK_COLLECTION_DCS_MCC_DETAIL.Against_Milk_Collection_MCC_Detail ) XXXDCS on XXXDCS.Against_Milk_Collection_MCC_Detail = TSPL_MILK_COLLECTION_MCC_DETAIL.PK_Id
                        where convert(date, TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate.Value, "dd-MMM-yyyy") & "',103) and convert (date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert (date,'" & clsCommon.GetPrintDate(dtpToDate.Value, "dd-MMM-yyyy") & "',103) ) xyz group by UploaderNo, Document_No    ) XXXFinal group by XXXFinal.Document_No )GetRateCode )GetAllGainLossRate left outer join TSPL_OWN_BMC_GAIN_LOSS_RATE on TSPL_OWN_BMC_GAIN_LOSS_RATE.Code=GetAllGainLossRate.GainLossCode  ) XXGetAllRecords group by Route_Code,ROUTE_NAME )PPP group  by  Route_Code,ROUTE_NAME order by Route_Code "

                dt = clsDBFuncationality.GetDataTable(Qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "TransportGainLossSummary", "Gain/Loss Report", Nothing)
                    frmCRV = Nothing
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Please click on Transpoter Gain/Loss Summary", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnDockDetail_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnDockDetail.CheckedChanged
        If rbtnDockDetail.Checked Then
            MyLabel4.Visible = True
            txtMCC_Code.Visible = True
        Else
            MyLabel4.Visible = False
            txtMCC_Code.Visible = False
        End If
    End Sub
End Class
