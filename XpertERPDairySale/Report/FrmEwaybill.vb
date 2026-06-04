Imports System.Reflection
Imports common
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Public Class FrmEwaybill
    Inherits FrmMainTranScreen
#Region "Variables"
    Const colewbNo As String = "colewbNo"
    Const colewbDate As String = "colewbDate"
    Const colstatus As String = "colstatus"
    Const colgenGstin As String = "colgenGstin"
    Const coldocNo As String = "coldocNo"
    Const coldocDate As String = "coldocDate"
    Const coldelPinCode As String = "coldelPinCode"
    Const coldelStateCode As String = "coldelStateCode"
    Const coldelPlace As String = "coldelPlace"
    Const colvalidUpto As String = "colvalidUpto"
    Const colextendedTimes As String = "colextendedTimes"
    Const colrejectStatus As String = "colrejectStatus"

#End Region
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
    End Sub
    Private Sub FrmEwaybill_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            SetUserMgmtNew()
            CreateTable()
            LoadEwaybillType()
            LoadvechicleReasonType()
            LoadExtendValidity()
            LoadConsignmentType()
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        gv1.DataSource = Nothing
        gv1.Rows.AddNew()
        Dim repoewbNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoewbNo.FormatString = ""
        repoewbNo.HeaderText = "EWB No"
        repoewbNo.Name = colewbNo
        repoewbNo.Width = 130
        repoewbNo.ReadOnly = True
        repoewbNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoewbNo)

        Dim repoewbDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoewbDate.FormatString = ""
        repoewbDate.HeaderText = "Ewb Date"
        repoewbDate.Name = colewbDate
        repoewbDate.Width = 150
        repoewbDate.ReadOnly = True
        repoewbDate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoewbDate)
        Dim repostaus As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repostaus.FormatString = ""
        repostaus.HeaderText = "Status"
        repostaus.Name = colstatus
        repostaus.Width = 50
        repostaus.ReadOnly = True
        repostaus.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repostaus)
        Dim repoGenGST As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoGenGST.FormatString = ""
        repoGenGST.HeaderText = "Gen GSTIN"
        repoGenGST.Name = colgenGstin
        repoGenGST.Width = 120
        repoGenGST.ReadOnly = True
        repoGenGST.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoGenGST)
        Dim repoDocno As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDocno.FormatString = ""
        repoDocno.HeaderText = "Dco No"
        repoDocno.Name = coldocNo
        repoDocno.Width = 120
        repoDocno.ReadOnly = True
        repoDocno.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDocno)
        Dim repoDocDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDocDate.FormatString = ""
        repoDocDate.HeaderText = "Doc Date"
        repoDocDate.Name = coldocDate
        repoDocDate.Width = 150
        repoDocDate.ReadOnly = True
        repoDocDate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDocDate)
        Dim repoPCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPCode.FormatString = ""
        repoPCode.HeaderText = "Pin Code"
        repoPCode.Name = coldelPinCode
        repoPCode.Width = 50
        repoPCode.ReadOnly = True
        repoPCode.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoPCode)
        Dim repoScode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoScode.FormatString = ""
        repoScode.HeaderText = "State Code"
        repoScode.Name = coldelStateCode
        repoScode.Width = 50
        repoScode.ReadOnly = True
        repoScode.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoScode)
        Dim repoPlace As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPlace.FormatString = ""
        repoPlace.HeaderText = "Place"
        repoPlace.Name = coldelPlace
        repoPlace.Width = 100
        repoPlace.ReadOnly = True
        repoPlace.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoPlace)
        Dim repoVupto As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoVupto.FormatString = ""
        repoVupto.HeaderText = "Valid Upto"
        repoVupto.Name = colvalidUpto
        repoVupto.Width = 150
        repoVupto.ReadOnly = True
        repoVupto.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoVupto)
        Dim repoextTime As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoextTime.FormatString = ""
        repoextTime.HeaderText = "Extended Times"
        repoextTime.Name = colextendedTimes
        repoextTime.Width = 50
        repoextTime.ReadOnly = True
        repoextTime.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoextTime)
        Dim repoRejstatus As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRejstatus.FormatString = ""
        repoRejstatus.HeaderText = "Reject Status"
        repoRejstatus.Name = colrejectStatus
        repoRejstatus.Width = 50
        repoRejstatus.ReadOnly = True
        repoRejstatus.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoRejstatus)



        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.EnableFiltering = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40

    End Sub
    Sub LoadEwaybillType()
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Name", GetType(String))
            Dim dr As DataRow = Nothing
            dr = dt.NewRow()
            dr("Code") = ""
            dr("Name") = "-- Select Eway Bill Types --"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Eway Bill Details"
            dr("Name") = "Eway Bill Details"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Update Transpoter"
            dr("Name") = "Update Transpoter"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Update PART-B"
            dr("Name") = "Update PART-B/Vehicle Number"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Cancel Eway Bill"
            dr("Name") = "Cancel Eway Bill"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Eway Bills By Date"
            dr("Name") = "Eway Bills By Date"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Extend Validity of Eway Bill"
            dr("Name") = "Extend Validity of Eway Bill"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Close Eway Bill"
            dr("Name") = "Close Eway Bill"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "Error List"
            dr("Name") = "Error List"
            dt.Rows.Add(dr)
            cmbEwaybillType.ValueMember = "Code"
            cmbEwaybillType.DisplayMember = "Name"
            cmbEwaybillType.DataSource = dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub LoadvechicleReasonType()
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Name", GetType(String))
            Dim dr As DataRow = Nothing
            dr = dt.NewRow()
            dr("Code") = ""
            dr("Name") = "-- Select Reason Type --"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "1"
            dr("Name") = "Due to Breakdown"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "2"
            dr("Name") = "Due to Transshipment"
            dt.Rows.Add(dr)
            'dr = dt.NewRow()
            'dr("Code") = "4"
            'dr("Name") = "First Time "
            'dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "3"
            dr("Name") = "Others"
            dt.Rows.Add(dr)
            cmbReasonForUpdateVehicle.ValueMember = "Code"
            cmbReasonForUpdateVehicle.DisplayMember = "Name"
            cmbReasonForUpdateVehicle.DataSource = dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub LoadExtendValidity()
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Name", GetType(String))
            Dim dr As DataRow = Nothing
            dr = dt.NewRow()
            dr("Code") = ""
            dr("Name") = "-- Select Extend Validity Reason --"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "1"
            dr("Name") = "Natural Calamity"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "2"
            dr("Name") = "Law and Order Situation"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "4"
            dr("Name") = "Transshipment"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "5"
            dr("Name") = "Accident"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "99"
            dr("Name") = "Others"
            dt.Rows.Add(dr)
            cmbExtendValidityReason.ValueMember = "Code"
            cmbExtendValidityReason.DisplayMember = "Name"
            cmbExtendValidityReason.DataSource = dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub LoadConsignmentType()
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Name", GetType(String))
            Dim dr As DataRow = Nothing
            dr = dt.NewRow()
            dr("Code") = ""
            dr("Name") = "-- Select Consignment Status --"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "M"
            dr("Name") = "inMovement"
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Code") = "T"
            dr("Name") = "inTransit"
            dt.Rows.Add(dr)
            cmbConsignmentStatus.ValueMember = "Code"
            cmbConsignmentStatus.DisplayMember = "Name"
            cmbConsignmentStatus.DataSource = dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub btnreset_Click(sender As Object, e As EventArgs) Handles btnreset.Click
        Try
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub Reset()
        Try
            cmbEwaybillType.SelectedValue = ""
            txtewbno.Enabled = False
            txtewbno.Value = ""
            txtTransID.Value = ""
            txtVehicleNo.Text = ""
            cmbExtendValidityReason.SelectedValue = ""
            cmbReasonForUpdateVehicle.SelectedValue = ""
            cmbConsignmentStatus.SelectedValue = ""
            txtRemaningDistance.Text = ""
            txtDate.Value = clsCommon.GETSERVERDATE()
            txtDate.Enabled = False
            LoadBlankGrid()
            gv1.Visible = False

            txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            Dim objResult As Object = Nothing
            If clsCommon.CompairString(cmbEwaybillType.SelectedValue, "") = CompairStringResult.Equal Then
                Throw New Exception("Please Select Eway Bill Type")
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Eway Bill Details") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 Then
                    Try
                        Dim isewbno As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(EWBNO) from TSPL_EWAY_BILL_REPORT_DETAIL where EWBNO='" & txtewbno.Value & "' "))
                        If isewbno = 0 Then
                            GetEwayBill(txtewbno.Value, True)
                        Else
                            LoadEWBDetail(txtewbno.Value)
                        End If
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Update Transpoter") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 AndAlso clsCommon.myLen(txtTransID.Value) > 0 Then
                    'Dim BillToLocaiton As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Bill_To_Location from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.value) & "'", Nothing))
                    Dim TransID As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select GSTFinalNo,*  from TSPL_VENDOR_MASTER where Vendor_Code='" & txtTransID.Value & "' and Vendor_Group_Code='TPT'", Nothing))
                    Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'", Nothing))
                    If clsCommon.myLen(TransID) > 0 Then
                        objResult = ClsEInvoiceOFAPIs.EWayBill_Update_Transporter(objCommonVar.CurrentCompanyCode, txtewbno.Value, TransID, txtLocation.Value, Nothing)
                        If objResult Is Nothing Then
                            Throw New Exception("Update Transpoter failed!")
                        Else
                            Dim ewayBillNo As String = objResult.SelectToken("data.ewayBillNo").ToString
                            Dim transporterId As String = objResult.SelectToken("data.transporterId").ToString
                            Dim transUpdateDate As String = objResult.SelectToken("data.transUpdateDate").ToString
                            Dim strTransCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vendor_Code from TSPL_VENDOR_MASTER where GSTFinalNo='" & transporterId & "'"))
                            Dim strUpdateQry As String = "update tspl_sd_sale_invoice_head set Transport_Code='" & clsCommon.myCstr(strTransCode) & "',transUpdateDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(transUpdateDate), "dd/MMM/yyyy hh:mm tt") & "' where EWayBillNo='" & clsCommon.myCstr(ewayBillNo) & "'"
                            clsDBFuncationality.ExecuteNonQuery(strUpdateQry)
                            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Nothing)
                            clsCommon.MyMessageBoxShow(Me, "Update Transpoter Successfully." & Environment.NewLine & "transporterId :" & strTransCode & Environment.NewLine & "Update Date :" & transUpdateDate, Me.Text)
                        End If
                    Else
                        Throw New Exception("Gstno not map with Transport [" & txtTransID.Value & "]")
                    End If
                Else
                    Throw New Exception("Please Enter ewb no/select Transpoter")
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Update PART-B") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 AndAlso clsCommon.CompairString(cmbReasonForUpdateVehicle.SelectedValue, "") <> CompairStringResult.Equal AndAlso clsCommon.myLen(txtVehicleNo.Text) > 0 Then
                    Dim isewbno As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(EWBNO) from TSPL_EWAY_BILL_REPORT_DETAIL where EWBNO='" & txtewbno.Value & "' "))
                    'Dim vehicleNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Number from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicleNo.Text & "' "))
                    Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'", Nothing))
                    If isewbno = 0 Then
                        GetEwayBill(txtewbno.Value, False)
                    Else
                        Dim obj As New clsEwayBillReportHead()
                        obj = clsEwayBillReportHead.GetData(txtewbno.Value)
                        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.ewbno) > 0 Then
                            Dim ObjVehicle As New clsEwayBillUpdatePartB_Vehicle()
                            ObjVehicle.ewbNo = obj.ewbno
                            ObjVehicle.vehicleNo = txtVehicleNo.Text
                            ObjVehicle.fromPlace = obj.fromPlace
                            ObjVehicle.fromState = obj.fromStateCode
                            ObjVehicle.reasonCode = clsCommon.myCstr(cmbReasonForUpdateVehicle.SelectedValue)
                            ObjVehicle.reasonRem = cmbReasonForUpdateVehicle.Text
                            If obj.VehiclListDetails.Count > 0 Then
                                ObjVehicle.transDocNo = obj.VehiclListDetails(0).transDocNo
                                ObjVehicle.transDocDate = clsCommon.myCstr(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(Nothing), "dd/MM/yyyy")) 'obj.VehiclListDetails(0).transDocDate
                                ObjVehicle.transMode = obj.VehiclListDetails(0).transMode
                            End If
                            objResult = ClsEInvoiceOFAPIs.EWayBill_Update_PartB(objCommonVar.CurrentCompanyCode, ObjVehicle, txtLocation.Value, Nothing)
                            If objResult Is Nothing Then
                                Throw New Exception("Update Vehicle no failed!")
                            Else
                                Dim josnReult As String = objResult.SelectToken("data").ToString
                                Dim vehUpdDate As String = objResult.SelectToken("data.vehUpdDate").ToString
                                Dim validUpto As String = objResult.SelectToken("data.validUpto").ToString
                                Dim strUpdateQry As String = "update tspl_sd_sale_invoice_head set VehicleNo='" & txtVehicleNo.Text & "', vehUpdDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(vehUpdDate), "dd/MMM/yyyy hh:mm tt") & "',EWayBillValidDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(validUpto), "dd/MMM/yyyy hh:mm tt") & "' where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'"
                                clsDBFuncationality.ExecuteNonQuery(strUpdateQry)
                                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Nothing)
                                clsCommon.MyMessageBoxShow(Me, "Update Prat-B/Vehicle number Successfully." & Environment.NewLine & "Vehicle Update Date :" & vehUpdDate & " " & Environment.NewLine & "validUpto :" & validUpto, Me.Text)
                            End If
                        End If
                    End If
                Else
                    Throw New Exception("Please Enter ewb no/Select Update Vehicle Reason.")
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Cancel Eway Bill") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 Then
                    Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'", Nothing))
                    If clsCommon.myLen(txtLocation.Value) > 0 Then
                        objResult = ClsEInvoiceOFAPIs.CancelEWayBill(objCommonVar.CurrentCompanyCode, txtewbno.Value, "Order Cancelled", txtLocation.Value, Nothing)
                        If objResult Is Nothing Then
                            Throw New Exception("e-way bill cancellation failed!")
                        Else
                            Dim CancelDate As String = objResult.SelectToken("data.cancelDate").ToString
                            Dim strUpdateQry As String = "update tspl_sd_sale_invoice_head set Ewb_cancelDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(CancelDate), "dd/MMM/yyyy hh:mm tt") & "' where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'"
                            clsDBFuncationality.ExecuteNonQuery(strUpdateQry)
                            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Nothing)
                            clsCommon.MyMessageBoxShow(Me, "Cancelled Successfully.", Me.Text)
                        End If
                    End If
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Eway Bills By Date") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtLocation.Value) > 0 Then
                    objResult = ClsEInvoiceOFAPIs.GetEwayBillByDate(objCommonVar.CurrentCompanyCode, txtDate.Value, txtLocation.Value, Nothing)
                    If objResult Is Nothing Then
                        Throw New Exception("Get Ewb Bills failed!")
                    Else
                        Dim ewayBillJson As String = objResult.SelectToken("data").ToString
                        'Dim ewayBill As clsEwayBillsByDate = JsonConvert.DeserializeObject(Of clsEwayBillsByDate)(ewayBillNo)
                        Dim ewayBill As List(Of clsEwayBillsByDateDetail) =
    JsonConvert.DeserializeObject(Of List(Of clsEwayBillsByDateDetail))(ewayBillJson)
                        If ewayBill IsNot Nothing And ewayBill.Count > 0 Then
                            LoadBlankGrid()
                            Dim dblrow As Integer = 0
                            For Each items As clsEwayBillsByDateDetail In ewayBill
                                gv1.Rows(dblrow).Cells(colewbNo).Value = items.ewbNo
                                gv1.Rows(dblrow).Cells(colewbDate).Value = items.ewbDate
                                gv1.Rows(dblrow).Cells(colstatus).Value = items.status
                                gv1.Rows(dblrow).Cells(colgenGstin).Value = items.genGstin
                                gv1.Rows(dblrow).Cells(coldocNo).Value = items.docNo
                                gv1.Rows(dblrow).Cells(coldocDate).Value = items.docDate
                                gv1.Rows(dblrow).Cells(coldelPinCode).Value = items.delPinCode
                                gv1.Rows(dblrow).Cells(coldelStateCode).Value = items.delStateCode
                                gv1.Rows(dblrow).Cells(coldelPlace).Value = items.delPlace
                                gv1.Rows(dblrow).Cells(colvalidUpto).Value = items.validUpto
                                gv1.Rows(dblrow).Cells(colextendedTimes).Value = items.extendedTimes
                                gv1.Rows(dblrow).Cells(colrejectStatus).Value = items.rejectStatus
                                gv1.Rows.AddNew()
                                dblrow += 1

                            Next
                        End If
                        'clsCommon.MyMessageBoxShow(Me, ewayBillNo, Me.Text)
                    End If
                Else
                    Throw New Exception("Please Select Locaiton!")
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Extend Validity of Eway Bill") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 AndAlso clsCommon.CompairString(cmbExtendValidityReason.SelectedValue, "") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(cmbConsignmentStatus.SelectedValue, "") <> CompairStringResult.Equal AndAlso clsCommon.myCdbl(txtRemaningDistance.Text) > 0 Then
                    Dim isewbno As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(EWBNO) from TSPL_EWAY_BILL_REPORT_DETAIL where EWBNO='" & txtewbno.Value & "' "))
                    Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'", Nothing))
                    If isewbno = 0 Then
                        GetEwayBill(txtewbno.Value, False)
                    Else
                        Dim obj As New clsEwayBillReportHead()
                        obj = clsEwayBillReportHead.GetData(txtewbno.Value)
                        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.ewbno) > 0 Then
                            Dim objExtend As New clsextendvalidity()
                            Dim EwbExtendTimeValid As Int64 = 0
                            EwbExtendTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EWayBillValidDate,GETDATE()),0) as validTill from tspl_sd_sale_invoice_head where  document_code = '" & strDocNo & "'"))
                            If EwbExtendTimeValid >= 8 OrElse EwbExtendTimeValid < -8 Then
                                Throw New Exception("The validity of EWB can be extended between 8 hours before expiry time and 8 hours after expiry time.")
                            End If
                            objExtend.ewbNo = clsCommon.myCdbl(obj.ewbno)
                            objExtend.vehicleNo = clsCommon.myCstr(obj.VehiclListDetails(0).vehicleNo)
                            objExtend.fromPlace = clsCommon.myCstr(obj.VehiclListDetails(0).fromPlace)
                            objExtend.fromState = clsCommon.myCdbl(obj.VehiclListDetails(0).fromState)
                            objExtend.remainingDistance = clsCommon.myCdbl(txtRemaningDistance.Text)
                            objExtend.transDocNo = clsCommon.myCstr(obj.VehiclListDetails(0).transDocNo)
                            objExtend.transDocDate = clsCommon.myCstr(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(Nothing), "dd/MM/yyyy"))
                            objExtend.transMode = clsCommon.myCstr(obj.VehiclListDetails(0).transMode)
                            If clsCommon.myCdbl(obj.VehiclListDetails(0).transMode) >= 1 AndAlso clsCommon.myCdbl(obj.VehiclListDetails(0).transMode) <= 4 Then
                                objExtend.transitType = ""
                                objExtend.addressLine1 = ""
                                objExtend.addressLine2 = ""
                                objExtend.addressLine3 = ""
                                If clsCommon.CompairString(cmbConsignmentStatus.SelectedValue, "M") <> CompairStringResult.Equal Then
                                    Throw New Exception("Consignmgnet status should be inMovement")
                                End If
                            ElseIf clsCommon.myCdbl(obj.VehiclListDetails(0).transMode) = 5 Then
                                objExtend.transitType = clsCommon.myCstr(obj.transactionType)
                                objExtend.addressLine1 = clsCommon.myCstr(obj.toAddr1)
                                objExtend.addressLine2 = clsCommon.myCstr(obj.toAddr2)
                                objExtend.addressLine3 = clsCommon.myCstr(obj.toAddr2)
                                If clsCommon.CompairString(cmbConsignmentStatus.SelectedValue, "T") <> CompairStringResult.Equal Then
                                    Throw New Exception("Consignmgnet status should be inTransit")
                                End If
                            End If
                            objExtend.extnRsnCode = clsCommon.myCdbl(cmbExtendValidityReason.SelectedValue)
                            objExtend.extnRemarks = clsCommon.myCstr(cmbExtendValidityReason.Text)
                            objExtend.fromPincode = clsCommon.myCdbl(obj.fromPincode)
                            objExtend.consignmentStatus = clsCommon.myCstr(cmbConsignmentStatus.SelectedValue)
                            objResult = ClsEInvoiceOFAPIs.EWayBill_ExtendValidity(objCommonVar.CurrentCompanyCode, objExtend, txtLocation.Value, Nothing)
                            If objResult Is Nothing Then
                                Throw New Exception("Extend Validity Failed!")
                            Else
                                Dim ewayBillNo As String = objResult.SelectToken("data.ewayBillNo").ToString
                                Dim updatedDate As String = objResult.SelectToken("data.updatedDate").ToString
                                Dim validUpto As String = objResult.SelectToken("data.validUpto").ToString
                                Dim strUpdateQry As String = "update tspl_sd_sale_invoice_head set ExtendValidityUpdate='" & clsCommon.GetPrintDate(clsCommon.myCDate(updatedDate), "dd/MMM/yyyy hh:mm tt") & "',EWayBillValidDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(validUpto), "dd/MMM/yyyy hh:mm tt") & "' where EWayBillNo='" & clsCommon.myCstr(ewayBillNo) & "'"
                                clsDBFuncationality.ExecuteNonQuery(strUpdateQry)
                                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Nothing)
                                clsCommon.MyMessageBoxShow(Me, "Extend Validity Successfully." & Environment.NewLine & objResult.SelectToken("data").ToString, Me.Text)
                            End If
                        End If
                    End If
                Else
                    Throw New Exception("Please Enter ewb no/Select Update Vehicle Reason/Select Consignmnet Status/enter Remaining Distance")
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Close Eway Bill") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtewbno.Value) > 0 AndAlso clsCommon.CompairString(cmbExtendValidityReason.SelectedValue, "") <> CompairStringResult.Equal Then
                    Dim isewbno As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(EWBNO) from TSPL_EWAY_BILL_REPORT_DETAIL where EWBNO='" & txtewbno.Value & "' "))
                    Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_sale_invoice_head where EWayBillNo='" & clsCommon.myCstr(txtewbno.Value) & "'", Nothing))
                    If isewbno = 0 Then
                        GetEwayBill(txtewbno.Value, False)
                    Else
                        Dim obj As New clsEwayBillReportHead()
                        obj = clsEwayBillReportHead.GetData(txtewbno.Value)
                        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.ewbno) > 0 Then
                            Dim objExtend As New clsEwayBillClose()
                            Dim EwbExtendTimeValid As Int64 = 0
                            EwbExtendTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EWayBillValidDate,GETDATE()),0) as validTill from tspl_sd_sale_invoice_head where  document_code = '" & strDocNo & "'"))
                            If EwbExtendTimeValid >= 8 OrElse EwbExtendTimeValid < -8 Then
                                Throw New Exception("The validity of EWB can be extended between 8 hours before expiry time and 8 hours after expiry time.")
                            End If
                            objExtend.ewbNo = clsCommon.myCdbl(obj.ewbno)
                            objExtend.closureDate = clsCommon.myCstr(clsCommon.GetPrintDate(txtclsewb.Value, "dd/MM/yyyy"))
                            objExtend.remarks = clsCommon.myCstr("Close the order")

                            objResult = ClsEInvoiceOFAPIs.EWayBill_Close(objCommonVar.CurrentCompanyCode, objExtend, txtLocation.Value, Nothing)
                            If objResult Is Nothing Then
                                Throw New Exception("Close EWB Failed!")
                            Else
                                Dim ewayBillNo As String = objResult.SelectToken("data.ewayBillNo").ToString
                                Dim CloseDate As String = objResult.SelectToken("data.ewbClosureDate").ToString
                                Dim strUpdateQry As String = "update tspl_sd_sale_invoice_head set EWayBillCloseDate='" & clsCommon.GetPrintDate(clsCommon.myCDate(CloseDate), "dd/MMM/yyyy hh:mm tt") & "' where EWayBillNo='" & clsCommon.myCstr(ewayBillNo) & "'"
                                clsDBFuncationality.ExecuteNonQuery(strUpdateQry)
                                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Nothing)
                                clsCommon.MyMessageBoxShow(Me, "Close EWB Successfully." & Environment.NewLine & objResult.SelectToken("data").ToString, Me.Text)
                            End If
                        End If
                    End If
                Else
                    Throw New Exception("Please select ewb no.")
                End If
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Error List") = CompairStringResult.Equal Then
                Throw New Exception("This API endpoint is not currently available.")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub GetEwayBill(ByVal strewbno As String, ByVal isGetewbno As Boolean)
        Try
            Dim objResult As Object = ClsEInvoiceOFAPIs.GetEwayBillDetail(objCommonVar.CurrentCompanyCode, txtewbno.Value, txtLocation.Value, Nothing)
            If objResult IsNot Nothing Then
                Dim Ewb_Json As String = objResult.SelectToken("data").ToString
                clsEwayBillReportHead.SaveData(Ewb_Json)
                If isGetewbno Then
                    LoadEWBDetail(strewbno)
                End If
            Else
                Throw New Exception("Something went worng!")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub LoadEWBDetail(ByVal strewbno As String)
        Try
            'If clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Eway Bill Details") = CompairStringResult.Equal Then
            '    Dim obj As New clsEwayBillReportHead()
            '    obj = clsEwayBillReportHead.GetData(strewbno)
            '    If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.ewbno) > 0) Then
            '    End If
            'ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Update Transpoter") = CompairStringResult.Equal Then
            'End If
            Dim QrCodeLen As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select DATALENGTH(EWayBill_QR_Code) as EWayBill_QR_Code from TSPL_EWAY_BILL_REPORT_DETAIL where ewbNo='" & strewbno & "'"))
            If QrCodeLen = 0 Then
                Dim CompGSTNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select userGstin from TSPL_EWAY_BILL_REPORT_DETAIL where ewbNo='" & strewbno & "'"))
                Dim EwbDt As String = clsCommon.GetPrintDate(clsDBFuncationality.getSingleValue("select ewayBillDate from TSPL_EWAY_BILL_REPORT_DETAIL where ewbNo='" & strewbno & "'"), "dd/MMM/yyyy hh:mm tt")
                Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(strewbno + "/" + CompGSTNo + "/" + EwbDt)
                clsDBFuncationality.UpdateImage("EWayBill_QR_Code", TempByte, "TSPL_EWAY_BILL_REPORT_DETAIL", "TSPL_EWAY_BILL_REPORT_DETAIL.ewbno='" & strewbno & "'")
            End If
            Dim Qry As String = "select cast(TSPL_EWAY_BILL_REPORT_DETAIL.EWayBill_QR_Code as image) as EWayBill_QR_Code, TSPL_SD_SALE_INVOICE_HEAD.IRN_No,
TSPL_EWAY_BILL_REPORT_DETAIL.ewbNo,TSPL_EWAY_BILL_REPORT_DETAIL.ewayBillDate,TSPL_EWAY_BILL_REPORT_DETAIL.genMode,TSPL_EWAY_BILL_REPORT_DETAIL.userGstin,TSPL_EWAY_BILL_REPORT_DETAIL.supplyType,TSPL_EWAY_BILL_REPORT_DETAIL.subSupplyType,
TSPL_EWAY_BILL_REPORT_DETAIL.docType,TSPL_EWAY_BILL_REPORT_DETAIL.docNo,TSPL_EWAY_BILL_REPORT_DETAIL.docDate,TSPL_EWAY_BILL_REPORT_DETAIL.fromGstin,TSPL_EWAY_BILL_REPORT_DETAIL.fromTrdName,TSPL_EWAY_BILL_REPORT_DETAIL.fromAddr1,TSPL_EWAY_BILL_REPORT_DETAIL.fromAddr2,TSPL_EWAY_BILL_REPORT_DETAIL.fromPlace,TSPL_EWAY_BILL_REPORT_DETAIL.fromPincode,TSPL_EWAY_BILL_REPORT_DETAIL.fromStateCode,TSPL_EWAY_BILL_REPORT_DETAIL.toGstin,TSPL_EWAY_BILL_REPORT_DETAIL.toTrdName,TSPL_EWAY_BILL_REPORT_DETAIL.toAddr1,TSPL_EWAY_BILL_REPORT_DETAIL.toAddr2,TSPL_EWAY_BILL_REPORT_DETAIL.toPlace,TSPL_EWAY_BILL_REPORT_DETAIL.toPincode,TSPL_EWAY_BILL_REPORT_DETAIL.toStateCode,totalValue,TSPL_EWAY_BILL_REPORT_DETAIL.totInvValue,TSPL_EWAY_BILL_REPORT_DETAIL.cgstValue,TSPL_EWAY_BILL_REPORT_DETAIL.sgstValue,TSPL_EWAY_BILL_REPORT_DETAIL.igstValue,TSPL_EWAY_BILL_REPORT_DETAIL.cessValue,TSPL_EWAY_BILL_REPORT_DETAIL.transporterId,TSPL_EWAY_BILL_REPORT_DETAIL.transporterName,TSPL_EWAY_BILL_REPORT_DETAIL.status,TSPL_EWAY_BILL_REPORT_DETAIL.actualDist,TSPL_EWAY_BILL_REPORT_DETAIL.noValidDays,TSPL_EWAY_BILL_REPORT_DETAIL.validUpto,TSPL_EWAY_BILL_REPORT_DETAIL.extendedTimes,TSPL_EWAY_BILL_REPORT_DETAIL.rejectStatus,TSPL_EWAY_BILL_REPORT_DETAIL.vehicleType,TSPL_EWAY_BILL_REPORT_DETAIL.actFromStateCode,TSPL_EWAY_BILL_REPORT_DETAIL.actToStateCode,TSPL_EWAY_BILL_REPORT_DETAIL.transactionType,TSPL_EWAY_BILL_REPORT_DETAIL.otherValue,TSPL_EWAY_BILL_REPORT_DETAIL.cessNonAdvolValue,
TSPL_EWAY_BILL_REPORT_Item_DETAIL.itemNo,TSPL_EWAY_BILL_REPORT_Item_DETAIL.productId,TSPL_EWAY_BILL_REPORT_Item_DETAIL.productDesc,TSPL_EWAY_BILL_REPORT_Item_DETAIL.hsnCode,TSPL_EWAY_BILL_REPORT_Item_DETAIL.quantity,TSPL_EWAY_BILL_REPORT_Item_DETAIL.qtyUnit,TSPL_EWAY_BILL_REPORT_Item_DETAIL.cgstRate,TSPL_EWAY_BILL_REPORT_Item_DETAIL.sgstRate,TSPL_EWAY_BILL_REPORT_Item_DETAIL.cessRate,TSPL_EWAY_BILL_REPORT_Item_DETAIL.cessNonAdvol,TSPL_EWAY_BILL_REPORT_Item_DETAIL.taxableAmount,
TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.updMode,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.vehicleNo,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.fromPlace as vfromPlace,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.fromState as vfromState,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.userGSTINTransin,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.enteredDate,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.transMode,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.transDocNo,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.transDocDate,TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.groupNo
from TSPL_EWAY_BILL_REPORT_DETAIL
left join TSPL_EWAY_BILL_REPORT_Item_DETAIL on TSPL_EWAY_BILL_REPORT_Item_DETAIL.ewbNo=TSPL_EWAY_BILL_REPORT_DETAIL.ewbNo
left join TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL on TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL.ewbNo=TSPL_EWAY_BILL_REPORT_DETAIL.ewbNo
left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo=TSPL_EWAY_BILL_REPORT_DETAIL.ewbNo
where TSPL_EWAY_BILL_REPORT_DETAIL.ewbNo='" & strewbno & "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptewbprint", "E-WayBill", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                frmCRV = Nothing
            Else
                Throw New Exception("No Data Found ")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub cmbEwaybillType_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs) Handles cmbEwaybillType.SelectedIndexChanged
        Try
            If clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Eway Bill Details") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = False
                txtclsewb.Enabled = False
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Update Transpoter") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = True
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = False
                txtclsewb.Enabled = False
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Update PART-B") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = True
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = True
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = False
                txtclsewb.Enabled = False
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Cancel Eway Bill") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = False
                txtclsewb.Enabled = False
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Eway Bills By Date") = CompairStringResult.Equal Then
                txtewbno.Enabled = False
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = True
                txtclsewb.Enabled = False
                gv1.Visible = True
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Extend Validity of Eway Bill") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = True
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = True
                cmbConsignmentStatus.Enabled = True
                txtDate.Enabled = False
                txtclsewb.Enabled = False
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Close Eway Bill") = CompairStringResult.Equal Then
                txtewbno.Enabled = True
                txtTransID.Enabled = False
                cmbReasonForUpdateVehicle.Enabled = False
                cmbExtendValidityReason.Enabled = False
                txtVehicleNo.Enabled = False
                txtRemaningDistance.Enabled = False
                cmbConsignmentStatus.Enabled = False
                txtDate.Enabled = False
                txtclsewb.Enabled = True
                gv1.Visible = False
            ElseIf clsCommon.CompairString(cmbEwaybillType.SelectedValue, "Error List") = CompairStringResult.Equal Then
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub CreateTable()
        Try
            Dim coll As Dictionary(Of String, String)
            coll = New Dictionary(Of String, String)()
            coll.Add("ewbNo", "VARCHAR(30) NOT NULL Primary key")
            coll.Add("ewayBillDate", "VARCHAR(30) NOT NULL")
            coll.Add("genMode", "Varchar(30) null")
            coll.Add("userGstin", "varchar(50) NULL")
            coll.Add("supplyType", "varchar(30) NULL")
            coll.Add("subSupplyType", "varchar(30) NULL")
            coll.Add("docType", "varchar(30) NULL")
            coll.Add("docNo", "varchar(30) NULL")
            coll.Add("docDate", "varchar(30) NULL")
            coll.Add("fromGstin", "varchar(30) NULL")
            coll.Add("fromTrdName", "varchar(100) NULL")
            coll.Add("fromAddr1", "varchar(100) NULL")
            coll.Add("fromAddr2", "varchar(100) NULL")
            coll.Add("fromPlace", "varchar(30) NULL")
            coll.Add("fromPincode", "varchar(30) NULL")
            coll.Add("fromStateCode", "varchar(30) NULL")
            coll.Add("toGstin", "varchar(30) NULL")
            coll.Add("toTrdName", "varchar(100) NULL")
            coll.Add("toAddr1", "varchar(100) NULL")
            coll.Add("toAddr2", "varchar(100) NULL")
            coll.Add("toPlace", "varchar(30) NULL")
            coll.Add("toPincode", "varchar(30) NULL")
            coll.Add("toStateCode", "varchar(30) NULL")
            coll.Add("totalValue", "varchar(30) NULL")
            coll.Add("totInvValue", "varchar(30) NULL")
            coll.Add("cgstValue", "varchar(30) NULL")
            coll.Add("sgstValue", "varchar(30) NULL")
            coll.Add("igstValue", "varchar(30) NULL")
            coll.Add("cessValue", "varchar(30) NULL")
            coll.Add("transporterId", "varchar(30) NULL")
            coll.Add("transporterName", "varchar(30) NULL")
            coll.Add("status", "varchar(30) NULL")
            coll.Add("actualDist", "varchar(30) NULL")
            coll.Add("noValidDays", "varchar(30) NULL")
            coll.Add("validUpto", "varchar(30) NULL")
            coll.Add("extendedTimes", "varchar(30) NULL")
            coll.Add("rejectStatus", "varchar(30) NULL")
            coll.Add("vehicleType", "varchar(30) NULL")
            coll.Add("actFromStateCode", "varchar(30) NULL")
            coll.Add("actToStateCode", "varchar(30) NULL")
            coll.Add("transactionType", "varchar(30) NULL")
            coll.Add("otherValue", "varchar(30) NULL")
            coll.Add("cessNonAdvolValue", "varchar(30) NULL")
            coll.Add("EWayBill_QR_Code", "image null")
            clsCommonFunctionality.CreateOrAlterTable("TSPL_EWAY_BILL_REPORT_DETAIL", coll)
            coll = New Dictionary(Of String, String)()
            coll.Add("ewbNo", "Varchar(30) not null References TSPL_EWAY_BILL_REPORT_DETAIL(ewbNo)")
            coll.Add("itemNo", "Varchar(30)")
            coll.Add("productId", "Varchar(30) ")
            coll.Add("productName", "Varchar(30) ")
            coll.Add("productDesc", "Varchar(30) ")
            coll.Add("hsnCode", "Varchar(30) ")
            coll.Add("quantity", "Varchar(30) ")
            coll.Add("qtyUnit", "Varchar(30) ")
            coll.Add("cgstRate", "Varchar(30) ")
            coll.Add("sgstRate", "Varchar(30) ")
            coll.Add("igstRate", "Varchar(30) ")
            coll.Add("cessRate", "Varchar(30) ")
            coll.Add("cessNonAdvol", "Varchar(30) ")
            coll.Add("taxableAmount", "Varchar(30) ")
            clsCommonFunctionality.CreateOrAlterTable("TSPL_EWAY_BILL_REPORT_ITEM_DETAIL", coll)
            coll = New Dictionary(Of String, String)()
            coll.Add("ewbNo", "Varchar(30) not null References TSPL_EWAY_BILL_REPORT_DETAIL(ewbNo)")
            coll.Add("updMode", "Varchar(30)")
            coll.Add("vehicleNo", "Varchar(30) ")
            coll.Add("fromPlace", "Varchar(30) ")
            coll.Add("fromState", "Varchar(30) ")
            coll.Add("tripshtNo", "Varchar(30) ")
            coll.Add("userGSTINTransin", "Varchar(30) ")
            coll.Add("enteredDate", "Varchar(30) ")
            coll.Add("transMode", "Varchar(30) ")
            coll.Add("transDocNo", "Varchar(30) ")
            coll.Add("transDocDate", "Varchar(30) ")
            coll.Add("groupNo", "Varchar(30) ")
            clsCommonFunctionality.CreateOrAlterTable("TSPL_EWAY_BILL_REPORT_VEHICLE_DETAIL", coll)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtTransID__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTransID._MYValidating
        Try
            Dim qry As String = "select Transport_Id,Transporter_Name from TSPL_TRANSPORT_MASTER"
            txtTransID.Value = clsCommon.ShowSelectForm("DSTransport No", qry, "Transport_Id", "", txtTransID.Value, "Transport_Id", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocation._MYValidating
        Try
            Dim qry As String = "select Location_Code as Code,Location_Desc as Name,Loc_Short_Name as [Short Name] from TSPL_LOCATION_MASTER "
            '' Dim WhrCls As String = "  Location_Type='Physical' and CSA_Type='N' and Is_Section='N' and Is_Sub_Location='N'  "
            Dim WhrCls As String = " 2=2 "
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtLocation.Value = clsCommon.ShowSelectForm("DS-ewbLocFndr", qry, "Code", WhrCls, txtLocation.Value, "Code", isButtonClicked)
        Catch ex As Exception
        End Try
    End Sub
    'Private Sub txtVehicleNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVehicleNo._MYValidating
    '    Try
    '        Dim qry As String = "Select distinct  vehicle_id ,Number,Description from TSPL_VEHICLE_MASTER"
    '         txtVehicleNo.Text = clsCommon.ShowSelectForm("Vehicle No", qry, "vehicle_id", "",  txtVehicleNo.Text, "vehicle_id", isButtonClicked)
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Private Sub txtewbno__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtewbno._MYValidating
        Try
            Dim strqry As String = "select EWayBillNo as ewbno,Document_Code as Invoice_no,Document_Date as Invoice_Date from TSPL_SD_SALE_INVOICE_HEAD"
            Dim Whrcls As String = " EWayBillNo is not null and convert(date,EWayBillValidDate,103)>='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE()) & "' "
            txtewbno.Value = clsCommon.ShowSelectForm("ewbno", strqry, "ewbno", Whrcls, txtewbno.Value, "Document_Date desc", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class
