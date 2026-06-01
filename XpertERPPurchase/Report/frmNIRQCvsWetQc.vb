
Imports System.Data.SqlClient
Imports common
Public Class frmNIRQCvsWetQc
    Inherits FrmMainTranScreen

    Private Sub frmNIRQCvsWetQc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        'If clsCommon.myLen(objCommonVar.arrCurrUserLocations) > 0 Then
        '    Dim Union As ArrayList = Nothing
        '    Dim qry As String = " Select DataBase_Name from TSPL_USER_MASTER where User_Code = '" + objCommonVar.CurrentUserCode + "'"
        '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '        Union = New ArrayList()
        '        For Each drZone As DataRow In dt.Rows
        '            Union.Add(clsCommon.myCstr(drZone("DataBase_Name")))
        '        Next
        '    End If
        '    txtLocation.arrValueMember = Union
        'End If
    End Sub
    Sub Reset()
        Try
            funreset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Sub funreset()
        gv1.DataSource = Nothing
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        RadPageView1.SelectedPage = RadPageViewPage1
        RadGroupBox1.Enabled = True
        RadPageView1.SelectedPage = RadPageViewPage1
        txtLocation.arrValueMember = Nothing
        TxtRAL.arrValueMember = Nothing
        txtItem.arrValueMember = Nothing
        TxtVendor.arrValueMember = Nothing
    End Sub
    Private Sub txtLocation__My_Click(sender As Object, e As EventArgs)
        Try
            Dim qry As String = "select Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER where Location_Type='Physical' "
            'qry += " where 2=2 and Seg_No = '7' AND GIT='N' "
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                qry += " and  TSPL_LOCATION_MASTER.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "Code", "Name", txtLocation.arrValueMember, txtLocation.arrDispalyMember)

            'Dim dt As DataTable
            'Dim qry As String = ""
            'qry = "select Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER where Location_Type='Physical' "
            ''qry += " where 2=2 and Seg_No = '7' AND GIT='N' "
            'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            '    qry += " and  tspl_location_master.location_code in (" + objCommonVar.strCurrUserLocations + ")"
            'End If

            'txtLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "Code", "Name", txtLocation.arrValueMember, txtLocation.arrDispalyMember)

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub TxtRAL__My_Click_1(sender As Object, e As EventArgs)
        Try
            Dim qry As String = "select  tspl_tender_header.DocumentCode as RALNO ,max(tspl_tender_header.DocumentDate) as DocumentDate  from tspl_tender_header
                            left outer join TSPL_TENDER_DETAIL on TSPL_TENDER_DETAIL.DocumentCode=tspl_tender_header.DocumentCode "
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                qry += " where TSPL_TENDER_DETAIL.Location In (" + clsCommon.GetMulcallString(txtLocation.arrValueMember) + ")"
            End If
            qry += " group by tspl_tender_header.DocumentCode "
            TxtRAL.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "RALNO", "", TxtRAL.arrValueMember, TxtRAL.arrDispalyMember)

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtVendor__My_Click_1(sender As Object, e As EventArgs)
        Dim qry As String = "select Vendor_Code as Code,Vendor_Name as Name from TSPL_VENDOR_MASTER  WHERE  Status='N'  order by Vendor_Code"
        TxtVendor.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelVUP", qry, "Code", "Name", TxtVendor.arrValueMember, TxtVendor.arrDispalyMember)
    End Sub

    Private Sub txtItem__My_Click(sender As Object, e As EventArgs)
        Try
            Dim qry As String = " Select Item_Code as Code,Item_Desc as Name,Short_Description,Structure_Code from TSPL_ITEM_MASTER "
            txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("@Item", qry, "Code", "Name", txtItem.arrValueMember, txtItem.arrDispalyMember)

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub Griddata(ByVal print As Boolean)
        Try

            Dim fromdate As String = clsCommon.myCDate(txtFromDate.Value, "dd/MM/yyyy")
            Dim Todate As String = clsCommon.myCDate(txtToDate.Value, "dd/MM/yyyy")
            Dim QRY As String = "SELECT  '" + objCommonVar.CurrentUserCode + "' as UserName ,  ROW_NUMBER() OVER (ORDER BY xxx.Document_No) AS SrNo,  max(Item_Desc)Item_Desc,FORMAT(CONVERT(date, MAX(Document_Date), 103), 'dd/MM/yy') AS Document_Date ,max(XXX.document_code)document_code,(xxx.Document_No)Document_No,max(xxx.RefTendorNo)RefTendorNo,max(xxx.SRN_NO)SRN_NO,FORMAT(CONVERT(date, MAX(xxx.SRN_DATE), 103), 'dd/MM/yy') AS SRN_DATE,
max(xxx.MRN_NO)MRN_NO,FORMAT(CONVERT(date, MAX(xxx.mrn_date), 103), 'dd/MM/yy') AS mrn_date,
max(xxx.Against_GRN)Against_GRN,FORMAT(CONVERT(date, MAX(xxx.GRN_DATE), 103), 'dd/MM/yy') AS GRN_DATE,
max(xxx.Vendor_Code)Vendor_Code,max(xxx.Vendor_Name)Vendor_Name,max(xxx.Bill_To_Location)Bill_To_Location,max(VehicleNo)VehicleNo,max(xxx.ITEM_CODE)ITEM_CODE ,
             max(xxx.Moistures)Moistures,max(xxx.Silica_DM)Silica_DM,max(xxx.Fat_DM)Fat_DM,max(xxx.Protein_DM)Protein_DM,max(xxx.Fiber_DM)Fiber_DM
             ,max(xxx.Moisture) Moisture,MAX(XXX.Silica)Silica,MAX(XXX.Fat) Fat,max(xxx.Protein) Protein,max(xxx.Fiber)Fiber,MAX(Comp_Name)Comp_Name,MAX(Add1)Add1,MAX(Add2)Add2,MAX(City_Code)LocDesc
             FROM ( SELECT * FROM ( SELECT TSPL_GRN_HEAD.VehicleNo,TSPL_GRN_DETAIL.Item_Desc ,TSPL_QC_CHECK_SRN_DETAIL.document_code,TSPL_NIR_QC.Document_No,convert(Varchar,TSPL_NIR_QC.Document_Date,103)Document_Date ,
             TSPL_SRN_HEAD.SRN_NO,convert(Varchar,TSPL_SRN_HEAD.SRN_DATE,103) AS SRN_DATE ,TSPL_MRN_HEAD.MRN_NO , convert(Varchar,TSPL_MRN_HEAD.mrn_date,103) AS mrn_date,
             TSPL_MRN_HEAD.Against_GRN as Against_GRN,convert(date,tspl_grn_head.GRN_DATE,103) as GRN_DATE,TSPL_MRN_HEAD.Vendor_Code as Vendor_Code,
             TSPL_MRN_HEAD.Vendor_Name as Vendor_Name,TSPL_MRN_HEAD.Bill_To_Location as Bill_To_Location,TSPL_MRN_detail.ITEM_CODE,
             TSPL_NIR_QC_FOSS.Moisture AS Moistures, TSPL_NIR_QC_FOSS.Silica_DM as Silica_DM,TSPL_NIR_QC_FOSS.Fat_DM AS Fat_DM,
             TSPL_NIR_QC_FOSS.Protein_DM AS Protein_DM,TSPL_NIR_QC_FOSS.Fiber_DM AS Fiber_DM, TSPL_QC_LOG_SHEET_MASTER.NIRQC_Para_type,
             TSPL_QC_CHECK_SRN_DETAIL.InputData,QC_Param_Code,TSPL_PURCHASE_ORDER_HEAD.RefTendorNo,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.ADD2,TSPL_LOCATION_MASTER.City_Code
    FROM TSPL_NIR_QC
    LEFT JOIN TSPL_NIR_QC_FOSS  ON TSPL_NIR_QC_FOSS.PK_Id = TSPL_NIR_QC.Against_Foss_PK_ID
    LEFT JOIN TSPL_MRN_HEAD  ON TSPL_MRN_HEAD.mrn_no = TSPL_NIR_QC.MRN_No
	    LEFT JOIN TSPL_MRN_detail  ON TSPL_MRN_detail.mrn_no = TSPL_MRN_HEAD.MRN_No

    LEFT JOIN TSPL_QC_CHECK_SRN_DETAIL  ON TSPL_QC_CHECK_SRN_DETAIL.MRN_NO = TSPL_MRN_HEAD.MRN_NO
    LEFT JOIN tspl_grn_head  ON tspl_grn_head.Grn_no = TSPL_MRN_HEAD.Against_GRN
	LEFT JOIN TSPL_GRN_DETAIL ON TSPL_GRN_DETAIL.Grn_no = tspl_grn_head.GRN_No
    LEFT JOIN TSPL_SRN_HEAD  ON TSPL_SRN_HEAD.Against_MRN = TSPL_MRN_HEAD.MRN_No
    LEFT JOIN TSPL_QC_LOG_SHEET_MASTER  ON TSPL_QC_LOG_SHEET_MASTER.code = TSPL_QC_CHECK_SRN_DETAIL.QC_Param_Code 
left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_MRN_HEAD.Against_PO
left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_MRN_HEAD.Bill_To_Location
LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
WHERE 2=2 
"
            QRY += "  and convert(Date,TSPL_NIR_QC.Document_Date,103) >='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And convert(Date,TSPL_NIR_QC.Document_Date,103)<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "

            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                QRY += " and TSPL_MRN_HEAD.Bill_To_Location in(" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ")" & Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                QRY += " and TSPL_QC_CHECK_SRN_DETAIL.ITEM_CODE in(" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ")" & Environment.NewLine
            End If
            If TxtVendor.arrValueMember IsNot Nothing AndAlso TxtVendor.arrValueMember.Count > 0 Then
                QRY += " and TSPL_MRN_HEAD.Vendor_Code in(" & clsCommon.GetMulcallString(TxtVendor.arrValueMember) & ")" & Environment.NewLine
            End If
            If TxtRAL.arrValueMember IsNot Nothing AndAlso TxtRAL.arrValueMember.Count > 0 Then
                QRY += " and TSPL_PURCHASE_ORDER_HEAD.RefTendorNo in(" & clsCommon.GetMulcallString(TxtRAL.arrValueMember) & ")" & Environment.NewLine
            End If
            If chkAutoNIR.Checked Then
                QRY += "   and Against_Foss_PK_ID is not null "
            End If
            QRY += " ) AS SourceTable
            PIVOT (MAX(InputData) FOR NIRQC_Para_type IN ( [Moisture],[Silica],[Fat],[Protein],[Fiber] )) AS PivotTable )XXX GROUP BY  xxx.Document_No order by GRN_DATE"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(QRY)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv1.DataSource = dt
                gv1.BestFitColumns()
                gv1.EnableFiltering = True
                gv1.ShowFilteringRow = True
                gv1.ReadOnly = True
                RadGroupBox1.Enabled = False
                EnableDisableCntrl(False)
                SetGridFormationOFGV1Collection()
                View()
                RadPageView1.SelectedPage = RadPageViewPage2
                Dim frmCRV As New frmCrystalReportViewer()
                If print = True Then
                    If chkAutoNIR.Checked Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptNirQCvsWetQCAutoNIR", "NIRQC vs WetQC")
                    Else
                        'Dim frmCRV As New frmCrystalReportViewer()
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptNirQCvsWetQC", "NIRQC vs WetQC")
                    End If
                    'Dim frmCRV As New frmCrystalReportViewer()
                    'frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptNirQCvsWetQC", "NIRQC vs WetQC")
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, "No data found to display", Me.Text)
        End Try
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Griddata(False)
    End Sub
    Sub EnableDisableCntrl(ByVal val As Boolean)
        RadGroupBox1.Enabled = False
    End Sub
    Sub SetGridFormationOFGV1Collection()
        gv1.TableElement.TableHeaderHeight = 40
        gv1.ShowRowHeaderColumn = False
        gv1.AllowAddNewRow = False
        gv1.AllowDeleteRow = False
        gv1.EnableFiltering = True
        gv1.ShowFilteringRow = True
        gv1.AutoExpandGroups = False

        gv1.MasterTemplate.ShowRowHeaderColumn = False
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).BestFit()
        Next
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).IsVisible = True
            gv1.Columns("UserName").IsVisible = False
            gv1.Columns("LocDesc").IsVisible = False
            gv1.Columns("Document_No").HeaderText = "NIRQC NO"
            gv1.Columns("Document_Date").HeaderText = "NIRQC Date"
            gv1.Columns("document_code").HeaderText = "Wet QC NO"
            gv1.Columns("SRN_NO").HeaderText = "SRN NO"
            gv1.Columns("SRN_DATE").HeaderText = "SRN DATE"
            gv1.Columns("MRN_NO").HeaderText = "MRN NO"
            gv1.Columns("mrn_date").HeaderText = "MRN Date"
            gv1.Columns("Against_GRN").HeaderText = "GRN"
            gv1.Columns("GRN_DATE").HeaderText = "GRN DATE"
            gv1.Columns("Vendor_Code").HeaderText = "Vendor Code"
            gv1.Columns("Vendor_Name").HeaderText = "Vendor Name"
            gv1.Columns("Bill_To_Location").HeaderText = "Location"
            gv1.Columns("ITEM_CODE").HeaderText = "Item Code"
            gv1.Columns("Item_Desc").HeaderText = "Item Desc"
            ' gv1.Columns("ITEM_CODE").HeaderText = "ITEM CODE"
            gv1.Columns("Moistures").HeaderText = "Moisture "
            gv1.Columns("Silica_DM").HeaderText = "Silica"
            gv1.Columns("Fat_DM").HeaderText = "Fat"
            gv1.Columns("Protein_DM").HeaderText = "Protein"
            gv1.Columns("Fiber_DM").HeaderText = "Fiber"
            gv1.Columns("RefTendorNo").HeaderText = "RAL "
        Next
        gv1.AutoSizeRows = True
        gv1.BestFitColumns()
        gv1.MasterTemplate.AutoExpandGroups = True
    End Sub
    Sub View()
        If gv1.Rows.Count > 0 Then
            Dim view As New ColumnGroupsViewDefinition()
            view.ColumnGroups.Add(New GridViewColumnGroup(" "))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Document_No").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Document_Date").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("document_code").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("SRN_NO").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("SRN_DATE").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("MRN_NO").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("mrn_date").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Against_GRN").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("GRN_DATE").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Vendor_Code").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Vendor_Name").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Bill_To_Location").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("ITEM_CODE").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Item_Desc").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("RefTendorNo").Name)
            view.ColumnGroups.Add(New GridViewColumnGroup("FOSS NIRQC "))
            view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Moistures").Name)
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Silica_DM").Name)
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Fat_DM").Name)
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Protein_DM").Name)
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Fiber_DM").Name)
            view.ColumnGroups.Add(New GridViewColumnGroup("WET QC "))
            view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Moisture").Name)
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Silica").Name)
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Fat").Name)
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Protein").Name)
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Fiber").Name)
            'End If
            gv1.ViewDefinition = view
        End If
    End Sub
    Sub ReportGrid1()
        gv1.DataSource = Nothing
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        gv1.GroupDescriptors.Clear()
        gv1.MasterTemplate.SummaryRowsBottom.Clear()
        gv1.MasterView.Refresh()
        gv1.ReadOnly = True
    End Sub

    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        Try
            ExportToExcel(EnumExportTo.Excel, gv1)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ExportToExcel(ByVal exporter As EnumExportTo, ByVal gv As RadGridView)
        Try
            If gv IsNot Nothing AndAlso gv.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                Dim strtemp As String = "Date Range : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") & " To " & clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")
                arrHeader.Add(strtemp)
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                'If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                '    arrHeader.Add(" Customer : " + clsCommon.GetMulcallStringWithComma(txtMultiCustomer.arrDispalyMember))
                'End If
                If exporter = EnumExportTo.Excel Then
                    clsCommon.MyExportToExcelGrid("NIRQC VS WetQc Report", gv, arrHeader, Me.Text, True)
                Else
                    clsCommon.MyExportToPDF("NIRQC VS WetQc Report", gv, arrHeader, "NIRQC VS WetQc Report", PageSetupReport_ID, objCommonVar.CurrentUserCode)
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiPDF_Click(sender As Object, e As EventArgs)
        Try
            ExportToExcel(EnumExportTo.PDF, gv1)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        Try
            funreset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnprint_Click(sender As Object, e As EventArgs) Handles btnprint.Click
        Griddata(True)
    End Sub

    Private Sub txtItem__My_Click_1(sender As Object, e As EventArgs) Handles txtItem._My_Click
        Try
            Dim qry As String = " Select Item_Code as Code,Item_Desc as Name,Short_Description,Structure_Code from TSPL_ITEM_MASTER "
            txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("@Item", qry, "Code", "Name", txtItem.arrValueMember, txtItem.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtLocation__My_Click_1(sender As Object, e As EventArgs) Handles txtLocation._My_Click
        Try
            Dim qry As String = "select Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER where Location_Type='Physical' "
            'qry += " where 2=2 and Seg_No = '7' AND GIT='N' "
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                qry += " and  TSPL_LOCATION_MASTER.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "Code", "Name", txtLocation.arrValueMember, txtLocation.arrDispalyMember)

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub TxtRAL__My_Click(sender As Object, e As EventArgs) Handles TxtRAL._My_Click
        Try
            Dim qry As String = "select  tspl_tender_header.DocumentCode as RALNO ,max(tspl_tender_header.DocumentDate) as DocumentDate  from tspl_tender_header
                            left outer join TSPL_TENDER_DETAIL on TSPL_TENDER_DETAIL.DocumentCode=tspl_tender_header.DocumentCode "
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                qry += " where TSPL_TENDER_DETAIL.Location In (" + clsCommon.GetMulcallString(txtLocation.arrValueMember) + ")"
            End If
            qry += " group by tspl_tender_header.DocumentCode "
            TxtRAL.arrValueMember = clsCommon.ShowMultipleSelectForm("LocMulSel", qry, "RALNO", "", TxtRAL.arrValueMember, TxtRAL.arrDispalyMember)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub TxtVendor__My_Click(sender As Object, e As EventArgs) Handles TxtVendor._My_Click
        Dim qry As String = "select Vendor_Code as Code,Vendor_Name as Name from TSPL_VENDOR_MASTER  WHERE  Status='N'  order by Vendor_Code"
        TxtVendor.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelVUP", qry, "Code", "Name", TxtVendor.arrValueMember, TxtVendor.arrDispalyMember)
    End Sub
End Class