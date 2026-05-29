Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports common.UserControls
Imports Microsoft.Office.Interop
Imports Telerik.WinControls.UI.Export

Public Module transportSql

    Dim ds As New DataSet()

    Public Sub FillGridView(ByVal sql As String, ByVal gv As RadGridView)
        Dim bs As New BindingSource()
        ds = connectSql.RunSQLReturnDS(sql)
        bs.DataSource = ds.Tables(0)
        gv.DataSource = bs
    End Sub
    Public Sub FillGridViewWithDt(ByVal Dt As DataTable, ByVal gv As RadGridView)
        Dim bs As New BindingSource()
        bs.DataSource = Dt
        gv.DataSource = Dt
    End Sub
    Public Sub FillComboBox(ByVal sql As String, ByVal ddl As RadDropDownList, ByVal dispMember As String, ByVal valueMember As String)
        ds = connectSql.RunSQLReturnDS(sql)
        ddl.DataSource = ds.Tables(0)
        ddl.DisplayMember = dispMember
        ddl.ValueMember = valueMember
        ddl.Text = "Select"
    End Sub
    Public Sub FillMultiColumnComboBox(ByVal sql As String, ByVal ddl As RadMultiColumnComboBox, ByVal dispMember As String, ByVal valueMember As String)
        ds = connectSql.RunSQLReturnDS(sql)
        ddl.DataSource = ds.Tables(0)
        ddl.DisplayMember = dispMember
        ddl.ValueMember = valueMember
        ddl.Text = "Select"
    End Sub


    Public Function OpenExporttoExcel(ByVal sql As String, ByVal frm As RadForm) As Boolean
        Dim path As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        If Not System.IO.Directory.Exists(path) Then
            System.IO.Directory.CreateDirectory(path)
        End If
        path += "\" + frm.Text + ".xls"
        If Not System.IO.File.Exists(path) Then
            System.IO.File.Delete(path)
        End If
        'Dim sfd As SaveFileDialog = New SaveFileDialog()
        'sfd.FileName = frm.Text
        'sfd.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        'path = forlderName
        'Else
        '    Return False
        'End If
        If Not path.Equals(String.Empty) Then
            Dim gv As New UserControls.MyRadGridView
            Try
                ''''' Dim exporter As New RadGridViewExcelExporter()
                gv.Name = "gTax"
                frm.Controls.Add(gv)
                FillGridView(sql, gv)
                If gv.Rows.Count = 0 Then
                    common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                    Return False
                End If
                Dim i As Integer = 0
                For i = 0 To gv.ColumnCount - 1
                    Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                    If TypeOf grow.Cells(i).Value Is DateTime Then
                        Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                        datecol.ExcelExportType = DisplayFormatType.ShortDate
                    End If
                Next i
                Dim exporter As New ExportToExcelML(gv)
                exporter.ExportVisualSettings = True
                AddHandler exporter.ExcelCellFormatting, AddressOf exporter_ExcelCellFormattingForOpen

                exporter.ExportHierarchy = True
                exporter.SheetMaxRows = ExcelMaxRows._65536
                exporter.SheetName = frm.Text
                exporter.RunExport(path)

                frm.Controls.Remove(gv)
                Dim xlApp As Excel.Application


                xlApp = New Excel.Application
                Process.Start(path)

            Catch ex As Exception
                frm.Controls.Remove(gv)
                common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
                Return False
            End Try
        End If
        Return True
    End Function

    Public Function ExporttoExcel(ByVal sql As String, ByVal frm As RadForm) As Boolean
        Return ExporttoExcel(sql, "", "", frm)
    End Function

    Public Function ExporttoExcel(ByVal sql As String, ByVal whrClaus As String, ByVal frm As RadForm) As Boolean
        Return ExporttoExcel(sql, whrClaus, "", frm)
    End Function

    Public Function ExporttoExcel(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm) As Boolean
        Return ExporttoExcel(sql, whrClaus, OrderByClaus, frm, Nothing)
    End Function

    Public Function ExporttoExcel(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm, ByVal manadatoryField As List(Of String)) As Boolean
        Return ExporttoExcel(sql, whrClaus, OrderByClaus, frm, manadatoryField, Nothing, "")
    End Function

    Public Function ExporttoExcel(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm, ByVal manadatoryField As List(Of String), ByVal SupermanadatoryField As List(Of String), ByVal formid As String) As Boolean 'ByVal ParamArray manadatoryField As String()
        Try
            ''************* Filter Block Start
            '============Add By Rohit on June 17,2014 to show column Filter========
            'Dim Goinside As Boolean = True
            Dim frmFilterCol As New frmFilterColumnsToExport()
            frmFilterCol.qry = sql
            frmFilterCol.whrCls = " Where 2=2 " + whrClaus
            frmFilterCol.formid = formid
            frmFilterCol.ListIEColumnsMandatory = SupermanadatoryField

            frmFilterCol.ShowDialog()
            If frmFilterCol.isCancel Then
                '   Goinside = False
                GoTo a
                'Return False
            End If
            sql = frmFilterCol.qry
            'Goinside = True
            '========================================================================
a:          Dim frmFilter As New frmFilterToExport()
            frmFilter.qry = sql

            frmFilter.ShowDialog()
            If frmFilter.isCancel Then
                Return False
            End If
            sql = frmFilter.qry
            If Not frmFilter.qry.ToUpper().Contains("ORDER BY") AndAlso clsCommon.myLen(OrderByClaus) > 0 Then
                sql = sql & " Order by " + OrderByClaus
            End If
            ''************* Filter Block End
            Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads" + "\"
            'Dim forlderName As String = "D" + ":\ERPTempFolder" + "\Downloads"
            Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            If IsExists = False Then
                System.IO.Directory.CreateDirectory(forlderName)
            End If
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim filePath As String = forlderName
            sfd.FileName = frm.Text
            sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = forlderName + "\" + sfd.FileName
            'Else
            '    Return False
            'End If



            If Not filePath.Equals(String.Empty) Then
                Dim gv As New UserControls.MyRadGridView
                Try
                    gv.Name = "gTax"
                    frm.Controls.Add(gv)
                    FillGridView(sql, gv)
                    gv.Visible = False
                    If gv.Rows.Count = 0 And frmFilter.IsBlankSheet = False Then
                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                        Return False
                    End If

                    Dim i As Integer = 0
                    '===============Add If Condition For Row Count by Rohit on june 05,2014 If Sheet is Blank then This Loop should not execute===
                    If gv.Rows.Count > 0 Then
                        For i = 0 To gv.ColumnCount - 1
                            Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                            If TypeOf grow.Cells(i).Value Is DateTime Then
                                Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                                datecol.ExcelExportType = DisplayFormatType.ShortDate
                            End If
                        Next i
                    End If
                    '==================================================================

                    Dim ext As String = Path.GetExtension(filePath)
                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        exportdataInCSV(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), IIf(frmFilter.IsBlankSheet, True, False)) 'frm.Text)
                    Else
                        'sanjay
                        If SupermanadatoryField IsNot Nothing AndAlso SupermanadatoryField.Count > 0 AndAlso clsCommon.myLen(formid) > 0 Then
                            transportSql.applyExpImpTemplate(gv, formid)
                        End If
                        exportdata(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), IIf(frmFilter.IsBlankSheet, True, False), Nothing, False, False, False, False, manadatoryField) 'frm.Text)
                        'sanjay
                    End If


                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        GoTo xxx
                    End If



xxx:

                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
                        System.Diagnostics.Process.Start(filePath)
                    End If



                Catch ex As Exception

                    frm.Controls.Remove(gv)

                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
        End Try
        Return True
    End Function

    Public Function ExporttoExcel(ByVal dt As DataTable, ByVal frm As RadForm) As Boolean
        Try
            Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
            Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            If IsExists = False Then
                System.IO.Directory.CreateDirectory(forlderName)
            End If
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim filePath As String
            sfd.FileName = frm.Text
            sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = forlderName + "\" + sfd.FileName
            'Else
            '    Return False
            'End If

            If Not filePath.Equals(String.Empty) Then
                Dim gv As New UserControls.MyRadGridView
                Try
                    gv.Name = "gTax1"
                    frm.Controls.Add(gv)
                    FillGridViewWithDt(dt, gv)
                    gv.Visible = False
                    If gv.Rows.Count = 0 Then
                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                        Return False
                    End If
                    Dim i As Integer = 0
                    If gv.Rows.Count > 0 Then
                        For i = 0 To gv.ColumnCount - 1
                            Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                            If TypeOf grow.Cells(i).Value Is DateTime Then
                                Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                                datecol.ExcelExportType = DisplayFormatType.ShortDate
                            End If
                        Next i
                    End If

                    Dim ext As String = Path.GetExtension(filePath)
                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        exportdataInCSV(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), False)
                    Else
                        exportdata(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), False, Nothing, False, False, False, False)
                    End If
                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
                        System.Diagnostics.Process.Start(filePath)
                    End If
                Catch ex As Exception
                    frm.Controls.Remove(gv)
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
        End Try
        Return True
    End Function
    Function isManadatory(ByVal colName As String, ByVal fieldNames As List(Of String)) As Boolean 'ByVal ParamArray fieldNames As String()
        For Each field As String In fieldNames
            If clsCommon.CompairString(field, colName) = CompairStringResult.Equal Then
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    ''''''''''''''''''''' function to export data with custom field
    Public Function ExporttoExcelWithCustomField(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm, ByVal formid As String) As Boolean

        ''************* Filter Block Start
        Dim frmFilter As New frmFilterToExport()
        frmFilter.qry = sql
        frmFilter.whrCls = " Where 2=2 " + whrClaus
        If clsCommon.myLen(OrderByClaus) > 0 Then
            frmFilter.orderByClause = " Order by " + OrderByClaus
        End If
        frmFilter.ShowDialog()
        If frmFilter.isCancel Then
            Return False
        End If
        sql = frmFilter.qry
        ''************* Filter Block End
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
        If IsExists = False Then
            System.IO.Directory.CreateDirectory(forlderName)
        End If
        Dim sfd As SaveFileDialog = New SaveFileDialog()
        Dim path As String
        sfd.FileName = frm.Text
        sfd.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        path = forlderName
        'Else
        '    Return False
        'End If
        If Not path.Equals(String.Empty) Then
            Dim gv As New UserControls.MyRadGridView
            Try
                ''''' Dim exporter As New RadGridViewExcelExporter()
                gv.Name = "gTax"
                frm.Controls.Add(gv)
                FillGridView(sql, gv)
                If gv.Rows.Count = 0 Then
                    common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                    Return False
                End If

                Dim i As Integer = 0
                For i = 0 To gv.ColumnCount - 1
                    Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                    If TypeOf grow.Cells(i).Value Is DateTime Then
                        Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                        datecol.ExcelExportType = DisplayFormatType.ShortDate
                    End If
                Next i

                clsCommon.ProgressBarShow()
                exportdatawithcustomfield(gv, path, frm.Text, formid)


                frm.Controls.Remove(gv)
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
                Dim excel As New Microsoft.Office.Interop.Excel.Application
                excel.Workbooks.Open(path)
                excel.Visible = True


            Catch ex As Exception
                frm.Controls.Remove(gv)
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
                Return False
            End Try
        End If
        Return True
    End Function
    Public Function ImportCustomFields(grow As GridViewRowInfo, FormId As String, UniqueColName As String, trans As SqlTransaction) As Boolean
        Dim isSaved As Boolean = True
        Try
            'Dim ChkCustomField As String = clsDBFuncationality.getSingleValue("select TSPL_CUSTOM_FIELD_HEAD.name,TSPL_CUSTOM_FIELD_HEAD.code from TSPL_CUSTOM_FIELD_MAPPING left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code =TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code where Program_Code ='" + clsUserMgtCode.frmMPMaster + "' ", trans)
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select TSPL_CUSTOM_FIELD_HEAD.name,TSPL_CUSTOM_FIELD_HEAD.code from TSPL_CUSTOM_FIELD_MAPPING left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code =TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code where Program_Code ='" + FormId + "' ", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                For j As Integer = 0 To dt.Rows.Count - 1
                    Try
                        Dim value As String = grow.Cells(dt.Rows(j)("name")).value
                        If clsCommon.myLen(value) > 0 Then
                            Dim chk As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_CUSTOM_FIELD_VALUES where Program_Code='" & FormId & "' and Transaction_Code='" & clsCommon.myCstr(grow.Cells(UniqueColName).Value) & "' and Custom_Field_Code ='" & dt.Rows(j)("Code") & "'", trans))
                            Dim qry As String = ""
                            If chk > 0 Then
                                ''Update query
                                qry = "update TSPL_CUSTOM_FIELD_VALUES set value='" & value & "' where Program_Code='" & FormId & "' and Transaction_Code='" & clsCommon.myCstr(grow.Cells(UniqueColName).Value) & "' and Custom_Field_Code ='" & dt.Rows(j)("Code") & "'"
                                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)

                            Else
                                ''Insert Query
                                qry = "insert into TSPL_CUSTOM_FIELD_VALUES  values('" & FormId & "','" & clsCommon.myCstr(grow.Cells(UniqueColName).Value) & "','" & dt.Rows(j)("Code") & "','" & value & "',0)"
                                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
                            End If
                        End If
                    Catch ex1 As Exception
                    End Try
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return isSaved
    End Function
    '-------------Overloaded By Pankaj Jha for Custom Filed Query Optimization on 30.05.2014
    Public Function ExporttoExcelWithCustomField(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm, ByVal formid As String, ByVal transCodeColumnName As String) As Boolean

        '-Getting Custom Fields Name
        Dim CustFldNames As String = clsDBFuncationality.getSingleValue("select isnull( left(fields,len(fields)-1),'') as Fields  from (select     ( select '['+Name  +'],'   from (select TSPL_CUSTOM_FIELD_HEAD.Name,TSPL_CUSTOM_FIELD_HEAD.Code  from TSPL_CUSTOM_FIELD_MAPPING  left outer join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code = TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code  where TSPL_CUSTOM_FIELD_MAPPING.Program_Code='" & formid & "') xx  FOR XML PATH ('')) Fields ) yy")
        If clsCommon.myLen(CustFldNames) > 0 Then
            sql = " select *  from ( select xxxx.*,  TSPL_CUSTOM_FIELD_HEAD.Name as Field_heading,TSPL_CUSTOM_FIELD_VALUES.Value as Field_Value from (  " & sql & " ) xxxx left outer join  TSPL_CUSTOM_FIELD_VALUES on TSPL_CUSTOM_FIELD_VALUES.Transaction_Code= " & transCodeColumnName & " left outer join  TSPL_CUSTOM_FIELD_MAPPING on TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code =  TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code and TSPL_CUSTOM_FIELD_MAPPING.Program_Code='" & formid & "' left outer join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code  )  xx Pivot (MAX(field_value) for Field_Heading in (" & CustFldNames & " )) ZZZ "
        End If
        ''************* Filter Block Start
        Dim frmFilter As New frmFilterToExport()
        frmFilter.qry = sql
        frmFilter.whrCls = " Where 2=2 " + whrClaus
        If clsCommon.myLen(OrderByClaus) > 0 Then
            frmFilter.orderByClause = " Order by " + OrderByClaus
        End If
        frmFilter.ShowDialog()
        If frmFilter.isCancel Then
            Return False
        End If
        sql = frmFilter.qry
        ''************* Filter Block End
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
        If IsExists = False Then
            System.IO.Directory.CreateDirectory(forlderName)
        End If
        Dim sfd As SaveFileDialog = New SaveFileDialog()
        Dim path As String
        sfd.FileName = frm.Text
        sfd.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        path = forlderName
        'Else
        '    Return False
        'End If
        If Not path.Equals(String.Empty) Then
            Dim gv As New UserControls.MyRadGridView
            Try
                ''''' Dim exporter As New RadGridViewExcelExporter()
                gv.Name = "gTax"
                frm.Controls.Add(gv)
                FillGridView(sql, gv)
                If gv.Rows.Count = 0 Then
                    common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                    Return False
                End If

                Dim i As Integer = 0
                For i = 0 To gv.ColumnCount - 1
                    Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                    If TypeOf grow.Cells(i).Value Is DateTime Then
                        Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                        datecol.ExcelExportType = DisplayFormatType.ShortDate
                    End If
                Next i


                clsCommon.ProgressBarShow()

                exportdata(gv, path, frm.Text, False, Nothing, False, False, False, False)


                frm.Controls.Remove(gv)
                clsCommon.ProgressBarHide()


            Catch ex As Exception
                frm.Controls.Remove(gv)
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
                Return False
            End Try
        End If
        Return True
    End Function

    Public Function exportdata(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing) As Integer
        Return exportdata(gv, flname, sname, 0, gv.Rows.Count, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, False, UseFilePath, manadatoryField)
    End Function

    Public Function exportdataBOQ(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing) As Integer
        Return exportdataBOQ(gv, flname, sname, 0, gv.Rows.Count, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, False, UseFilePath, manadatoryField)
    End Function
    'Ticket No-TEC/13/02/19-000423 sanjay add parameter UseFilePath
    'Ticket No-TEC/10/04/19-000466 , sanjay ,setting-ExportToDefineLocation

    Public Function exportdata(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing) As Integer
        Dim MaxRowExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsInExcelExport, clsFixedParameterCode.MaxRowsInExcelExport, Nothing))
        Return exportdata(MaxRowExport, gv, flname, sname, fromRow, ToRow, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
    End Function

    Public Function exportdataBOQ(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing) As Integer
        Dim MaxRowExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsInExcelExport, clsFixedParameterCode.MaxRowsInExcelExport, Nothing))
        Return exportdataBOQ(MaxRowExport, gv, flname, sname, fromRow, ToRow, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
    End Function

    Public Function exportdataBOQ(ByVal MaxRowExport As Integer, ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing) As Integer
        Dim FileCount As Integer = 1

        'sanjay
        If sname.Contains("/") Then
            sname = sname.Replace("/", " ")
        End If
        If sname.Contains("\") Then
            sname = sname.Replace("\", " ")
        End If
        'If UseFilePath = False And MultipleFiles = False Then
        '    'If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ExportToDefineLocation, clsFixedParameterCode.ExportToDefineLocation, Nothing)) = 1 Then
        '    '    UseFilePath = True
        '    '    Dim sfd As SaveFileDialog = New SaveFileDialog()
        '    '    sfd.FileName = sname
        '    '    'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 *.xlsx|(*.xlsx);|CSV Files (*.csv) |*.csv"
        '    '    sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
        '    '    If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '    '        flname = sfd.FileName
        '    '        sname = flname.Substring(flname.LastIndexOf("\") + 1, flname.Length - flname.LastIndexOf("\") - 1)
        '    '    Else
        '    '        Return False
        '    '    End If
        '    'End If
        'End If
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
        If IsExists = False Then
            System.IO.Directory.CreateDirectory(forlderName)
        End If

        If UseFilePath = False And MultipleFiles = False Then
            Dim currTime As DateTime = DateTime.Now
            sname = sname & clsCommon.GetPrintDate(currTime, "yyyyMMddhhmmss")
            flname = forlderName & "\" & sname & ".xls"
            sname = sname & ".xls"
        End If
        'sanjay

        Try
            '==========================Add isblanksheet Variable in Function to Save Blank Excel Sheet===================
            If ((gv.Columns.Count = 0) Or (gv.ChildRows.Count = 0)) And isblanksheet = False Then
                Return FileCount
            End If
            If fromRow <= 0 Then
                fromRow = 1
            End If
            If ToRow <= 0 Then
                ToRow = gv.Rows.Count
            End If
            '========================================================================================
            '' check rows limit on work sheet

            If MaxRowExport <= 0 Then
                MaxRowExport = 1048576
            End If
            If gv.Rows.Count >= MaxRowExport And MultipleFiles = False Then
                MultipleFiles = True
                Dim FilePathRoot As String = System.IO.Path.GetDirectoryName(flname)
                Dim fileName As String = System.IO.Path.GetFileName(flname)
                Dim fileExtn As String = System.IO.Path.GetExtension(flname)
                fileName = fileName.Replace(fileExtn, "")
                flname = FilePathRoot & "\" & fileName
                If System.IO.Directory.Exists(flname) = False Then
                    System.IO.Directory.CreateDirectory(flname)
                End If

                'Dim tblArr() As DataTable
                Dim tableCount = Math.Ceiling(gv.Rows.Count / MaxRowExport)
                Dim Divisor = gv.Rows.Count / tableCount
                '' for data boun through data reader
                Dim fromCount As Integer = 1
                Dim ToCount As Integer = gv.Rows.Count
                FileCount = tableCount
                For intLoop As Integer = 0 To tableCount - 1
                    Dim file As String
                    If intLoop = tableCount - 1 Then
                        ToCount = gv.Rows.Count
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdataBOQ(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = fromCount + (MaxRowExport)
                        common.clsCommon.MyMessageBoxShow("Data Exported in directory -" & System.IO.Path.GetDirectoryName(flname) & "\" & System.IO.Path.GetFileName(flname) & " in " & FileCount & " files.")
                    Else
                        ToCount = (fromCount + 0) + (MaxRowExport - 1)
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdataBOQ(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = (fromCount) + (MaxRowExport)
                    End If
                Next
                Return FileCount
                'Throw New Exception("Max row limit for worksheet in excel is 1048576")
            End If
            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()



            clsCommon.ProgressBarPercentShow()

            'add column to that table        
            For i As Integer = 0 To gv.ColumnCount - 1
                'dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
                dset.Tables(0).Columns.Add("Column" & (i + 1))
                dset.Tables(0).Columns("Column" & (i + 1)).Caption = IIf(ExportWithoutHeader, "", gv.Columns(i).HeaderText)
            Next

            'add rows to the table
            'dset.Tables(0).Rows(0).Delete()

            Dim dr1 As DataRow



            Dim excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet


            excel.Visible = True ' Set this to True if you want to see the Excel application
            wBook = excel.Workbooks.Open(flname)
            wSheet = wBook.Sheets("BoQ1")


            Dim GridCurrentRowIndex As Int64 = fromRow - 2
            Dim GridLastSavedRowIndex As Int64 = fromRow - 2

            'wBook = excel.Workbooks.Add()


            'sanjay for summary row
            If MaxRowExport <> ToRow Then
                If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                    ToRow = ToRow + 1
                End If
            End If
            'sanjay for summary row

            Dim rawData((ToRow - fromRow + 1), gv.Columns.Count - 1) As Object
            wSheet = wBook.ActiveSheet()

            If clsCommon.myLen(sname) > 31 Then
                sname = sname.Substring(0, 31)
            End If

            'Dim dp As Object
            ' dp = wBook.BuiltinDocumentProperties
            'dp("Tags").Value = clsUserMgtCode.frmComplaintMaster

            wSheet.Name = sname

            Dim flag As Boolean = False
            Dim colnum As Integer = -1
            Dim PrevCol As Integer = -1
            Dim isResteRawData As Boolean = True
            Dim ColNums(0 To gv.Columns.Count - 1) As Integer
            Dim MaxRowsToExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickExport, clsFixedParameterCode.MaxRowsForQuickExport, Nothing))
            Dim jk As Integer = 0
            For i As Integer = 0 To gv.Columns.Count - 1
                jk += 1
                ''richa agarwal 12jan BM00000008685
                If FormatCellofExcel = False Then
                    If TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                    End If
                Else
                    ''richa agarwal 08-jan-2015 format for excel sheet cell format @ as text,General as for General, 0 as numnber,0.00 as decimal against ticket no BM00000008659,BM00000008854
                    If clsCommon.CompairString(gv.Columns(i).Name, "ColAcNo") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "ColPAYEEACNO") = CompairStringResult.Equal Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                    ElseIf clsCommon.CompairString(gv.Columns(i).Name, "colSCACNO") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "colBeniACNO") = CompairStringResult.Equal Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                    ElseIf TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                        ' wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "General"
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                    ElseIf TypeOf gv.Columns(i) Is GridViewDateTimeColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "dd/mmm/yyyy"
                    End If
                    ''-------------------
                End If
            Next

            Dim colIndex As Integer = 1
            Dim rowIndex As Integer = 13

            If Not IsNothing(arrHeader) Then
                For Each Str As String In arrHeader
                    excel.Cells(rowIndex, colIndex) = Str
                    rowIndex += 1
                Next
            End If


            ''richa agarwal 09-feb-2016 to show double header in excel BM00000008811
            If doubleheadershowninExcel = True Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                Dim k As Integer = 0
                If view.ColumnGroups.Count > 0 Then
                    Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                        Dim l As Integer = 0
                        For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
                            'If view.ColumnGroups(i).Rows(0).ColumnNames(m).IsVisible = False Then ''TELERIK2015->2022
                            '    l = l + 1
                            'End If

                            'If view.ColumnGroups(i).IsVisible = False Then ''TELERIK2015->2022
                            If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
                                l = l + 1
                            End If

                        Next

                        k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                        chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
                        chartRange.Merge()
                        chartRange.VerticalAlignment = 2
                        chartRange.HorizontalAlignment = 3


                        'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
                        j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                    Next

                    rowIndex += 1
                End If

            End If
            ''------------------------------------------
            While isResteRawData
                If (ToRow - fromRow + 1) <= MaxRowsToExport Then

                    ReDim rawData((ToRow - fromRow), gv.Columns.Count - 1)
                    GridCurrentRowIndex = fromRow - 1
                    GridLastSavedRowIndex = ToRow - 1 ''gv.ChildRows.Count - 1
                    isResteRawData = False
                Else
                    GridCurrentRowIndex = GridLastSavedRowIndex + 1
                    If (GridLastSavedRowIndex + 1) + MaxRowsToExport < (ToRow) Then
                        ReDim rawData(MaxRowsToExport - 1, gv.Columns.Count - 1)
                        GridLastSavedRowIndex = (GridLastSavedRowIndex + (MaxRowsToExport))
                        isResteRawData = True
                    Else
                        ReDim rawData((ToRow - fromRow) - (GridLastSavedRowIndex - fromRow + 1 - 1), gv.Columns.Count - 1)
                        GridLastSavedRowIndex = ToRow - 1
                        isResteRawData = False
                    End If
                    'GridLastSavedRowIndex = (GridLastSavedRowIndex + MaxRowsToExport)
                End If
                Dim RowDataRIndix As Integer = 0
                For i As Integer = GridCurrentRowIndex To GridLastSavedRowIndex
                    'gv.RowCount, gv.Columns.Count - 1
                    dr1 = dset.Tables(0).NewRow
                    clsCommon.ProgressBarPercentUpdate(((i - fromRow + 1 + 1) * 100) / (ToRow - fromRow + 1), " Exporting Record  " & (i - fromRow + 1 + 1) & "  Out of " & (ToRow - fromRow + 1))
                    Try
                        For j As Integer = 0 To gv.Columns.Count - 1
                            dr1(j) = clsCommon.myCstr(gv.ChildRows(i).Cells(j).Value)
                            rawData(RowDataRIndix, j) = dr1(j).ToString()
                        Next
                    Catch ex As Exception
                        RowDataRIndix = RowDataRIndix + 1
                        Exit For
                    End Try
                    dset.Tables(0).Rows.Add(dr1)
                    RowDataRIndix = RowDataRIndix + 1
                Next

                'sanjay
                Try
                    If isResteRawData = False AndAlso MaxRowExport <> ToRow Then
                        If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                            RowDataRIndix = RowDataRIndix - 1
                            dr1 = dset.Tables(0).NewRow
                            For ii As Integer = 0 To gv.MasterTemplate.SummaryRowsBottom(0).Count - 1
                                Dim colName As String = gv.MasterTemplate.SummaryRowsBottom(0)(ii).Name
                                If gv.Columns.Contains(colName) AndAlso gv.Columns(colName).IsVisible Then
                                    Dim summaryItem As GridViewSummaryItem = gv.SummaryRowsBottom(0)(ii)

                                    Dim summary As Object = Nothing

                                    Try
                                        summary = summaryItem.Evaluate(gv.MasterTemplate)
                                    Catch ex As Exception
                                    End Try
                                    dr1(gv.Columns(colName).Index) = IIf(((summary = Nothing) Or (Single.NaN.Equals(summary))), DBNull.Value, clsCommon.myCstr(summary))
                                    rawData(RowDataRIndix, gv.Columns(colName).Index) = clsCommon.myCstr(dr1(gv.Columns(colName).Index).ToString())
                                End If
                            Next

                            dset.Tables(0).Rows.Add(dr1)
                        End If
                    End If
                Catch ex As Exception
                End Try
                'sanjay



                Dim dt As System.Data.DataTable = dset.Tables(0)


                colIndex = 0



                Dim LastColumn As String = ColumnIndexToColumnLetter(dt.Columns.Count)
                Dim Lastrow As Integer = 13 + dt.Rows.Count + 1 ''change

                Dim row As Integer = 0
                Dim col As Integer = 0
                If Not isblanksheet Then
                    wSheet.Range("A" & ((GridCurrentRowIndex - fromRow + 1) + rowIndex + 1), LastColumn & ((GridLastSavedRowIndex - fromRow + 1) + rowIndex + 1)).Value2 = rawData
                End If
                rawData = Nothing
                dt = Nothing
                ''new code by Panch Raj
                dset.Tables(0).Rows.Clear()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                If (ToRow - fromRow + 1) - 1 > (GridLastSavedRowIndex - fromRow + 1) Then
                    isResteRawData = True
                Else
                    isResteRawData = False
                End If
            End While

            'Hide Group columns
            If doubleheadershowninExcel = True Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                Dim k As Integer = 0
                If view.ColumnGroups.Count > 0 Then
                    'Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        'excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                        Dim l As Integer = 0
                        For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
                            'If view.ColumnGroups(i).Rows(0).Columns(m).IsVisible = False Then''TELERIK2015->2022
                            'l = l + 1
                            'End If
                            'If view.ColumnGroups(i).IsVisible = False Then
                            If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
                                l = l + 1
                            End If
                        Next

                        k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                        'chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
                        'chartRange.Merge()
                        'chartRange.VerticalAlignment = 2
                        'chartRange.HorizontalAlignment = 3
                        If view.ColumnGroups(i).IsVisible = False Then
                            'chartRange.EntireColumn.Hidden = True
                            wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k)).EntireColumn.Hidden = True
                        End If

                        'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
                        j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                    Next

                    rowIndex += 1
                End If

            End If

            clsCommon.ProgressBarPercentUpdate(100, "Wait Adjusting Columns Autofit ")

            If doubleheadershowninExcel = False Then
                wSheet.Columns.AutoFit()
            End If

            'Manadatory Field Coloring
            If manadatoryField IsNot Nothing AndAlso manadatoryField.Count > 0 Then

                For c As Integer = 0 To wSheet.Columns.Count - 1
                    If clsCommon.myLen(wSheet.Cells(1, c + 1).value) <= 0 Then
                        Exit For
                    End If

                    If isManadatory(wSheet.Cells(1, c + 1).value, manadatoryField) Then
                        wSheet.Cells(1, c + 1).interior.color = RGB(Color.LightGoldenrodYellow.R, Color.LightGoldenrodYellow.G, Color.LightGoldenrodYellow.B)
                    End If
                Next

                For i As Integer = 0 To gv.Columns.Count - 1
                    If Not gv.Columns(i).IsVisible Then
                        wSheet.Columns(i + 1).EntireColumn.Hidden = True
                    End If
                Next


            End If


            Dim strFileName As String = flname
            Dim blnFileOpen As Boolean = False

            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try
        clsCommon.ProgressBarPercentHide()
        Return FileCount
    End Function

    Public Function exportdata(ByVal MaxRowExport As Integer, ByVal gv As MyRadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing, Optional ByVal AllCellsInString As Boolean = False) As Integer
        Return exportdata(False, MaxRowExport, gv, flname, sname, fromRow, ToRow, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField, AllCellsInString)
    End Function

    Public Function exportdata(ByVal isArrHeaderMerged As Boolean, ByVal MaxRowExport As Integer, ByVal gv As MyRadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing, Optional ByVal AllCellsInString As Boolean = False, Optional isFontSizeIncrease As Boolean = False) As Integer
        Dim FileCount As Integer = 1
        If AllCellsInString Then
            MaxRowExport = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsInExcelExport, clsFixedParameterCode.MaxRowsInExcelExport, Nothing))
        End If
        'sanjay
        If sname.Contains("/") Then
            sname = sname.Replace("/", " ")
        End If
        If sname.Contains("\") Then
            sname = sname.Replace("\", " ")
        End If
        'If UseFilePath = False And MultipleFiles = False Then
        '    If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ExportToDefineLocation, clsFixedParameterCode.ExportToDefineLocation, Nothing)) = 1 Then
        '        UseFilePath = True
        '        Dim sfd As SaveFileDialog = New SaveFileDialog()
        '        sfd.FileName = sname
        '        'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 *.xlsx|(*.xlsx);|CSV Files (*.csv) |*.csv"
        '        sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
        '        If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '            flname = sfd.FileName
        '            sname = flname.Substring(flname.LastIndexOf("\") + 1, flname.Length - flname.LastIndexOf("\") - 1)
        '        Else
        '            Return False
        '        End If
        '    End If
        'End If
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
        If IsExists = False Then
            System.IO.Directory.CreateDirectory(forlderName)
        End If

        If UseFilePath = False And MultipleFiles = False Then
            Dim currTime As DateTime = DateTime.Now
            sname = sname & clsCommon.GetPrintDate(currTime, "yyyyMMddhhmmss")
            flname = forlderName & "\" & sname & ".xls"
            sname = sname & ".xls"
        Else
            flname = flname & ".xls"
        End If
        'sanjay

        Try
            '==========================Add isblanksheet Variable in Function to Save Blank Excel Sheet===================
            If ((gv.Columns.Count = 0) Or (gv.ChildRows.Count = 0)) And isblanksheet = False Then
                Return FileCount
            End If
            If fromRow <= 0 Then
                fromRow = 1
            End If
            If ToRow <= 0 Then
                ToRow = gv.Rows.Count
            End If
            '========================================================================================
            '' check rows limit on work sheet

            If MaxRowExport <= 0 Then
                MaxRowExport = 1048576
            End If
            If gv.Rows.Count >= MaxRowExport And MultipleFiles = False Then
                MultipleFiles = True
                Dim FilePathRoot As String = System.IO.Path.GetDirectoryName(flname)
                Dim fileName As String = System.IO.Path.GetFileName(flname)
                Dim fileExtn As String = System.IO.Path.GetExtension(flname)
                fileName = fileName.Replace(fileExtn, "")
                flname = FilePathRoot & "\" & fileName
                If System.IO.Directory.Exists(flname) = False Then
                    System.IO.Directory.CreateDirectory(flname)
                End If

                'Dim tblArr() As DataTable
                Dim tableCount = Math.Ceiling(gv.Rows.Count / MaxRowExport)
                Dim Divisor = gv.Rows.Count / tableCount
                '' for data boun through data reader
                Dim fromCount As Integer = 1
                Dim ToCount As Integer = gv.Rows.Count
                FileCount = tableCount
                For intLoop As Integer = 0 To tableCount - 1
                    Dim file As String
                    If intLoop = tableCount - 1 Then
                        ToCount = gv.Rows.Count
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdata(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = fromCount + (MaxRowExport)
                        common.clsCommon.MyMessageBoxShow("Data Exported in directory -" & System.IO.Path.GetDirectoryName(flname) & "\" & System.IO.Path.GetFileName(flname) & " in " & FileCount & " files.")
                    Else
                        ToCount = (fromCount + 0) + (MaxRowExport - 1)
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdata(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = (fromCount) + (MaxRowExport)
                    End If
                Next
                Return FileCount
                'Throw New Exception("Max row limit for worksheet in excel is 1048576")
            End If
            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()

            'Do not remove column for master Export/Import
            If Not (manadatoryField IsNot Nothing AndAlso manadatoryField.Count > 0) Then
                Dim IndexList As List(Of String) = New List(Of String)

                For i As Integer = 0 To gv.Columns.Count - 1
                    If Not gv.Columns(i).IsVisible Then
                        IndexList.Add(gv.Columns(i).Name)
                    End If
                Next

                For i As Integer = 0 To IndexList.Count - 1
                    gv.Columns.Remove(IndexList.Item(i).ToString())
                Next
            End If

            clsCommon.ProgressBarPercentShow()
            'add column to that table            
            For i As Integer = 0 To gv.ColumnCount - 1
                'dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
                dset.Tables(0).Columns.Add("Column" & (i + 1))
                dset.Tables(0).Columns("Column" & (i + 1)).Caption = IIf(ExportWithoutHeader, "", gv.Columns(i).HeaderText)
            Next
            'add rows to the table
            'dset.Tables(0).Rows(0).Delete()

            Dim dr1 As DataRow


            'For row = 0 To dt.Rows.Count - 1
            '    For col = 0 To dt.Columns.Count - 1
            '        rawData(row, col) = dt.Rows(row).ItemArray(col)
            '    Next
            '    clsCommon.ProgressBarPercentUpdate((row * 100) / dt.Rows.Count, " Exporting Record  " & row & "  Out of " & dt.Rows.Count)
            'Next
            Dim excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            Dim GridCurrentRowIndex As Int64 = fromRow - 2
            Dim GridLastSavedRowIndex As Int64 = fromRow - 2

            wBook = excel.Workbooks.Add()

            'sanjay for summary row
            If MaxRowExport <> ToRow Then
                If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                    ToRow = ToRow + 1
                End If
            End If
            'sanjay for summary row

            Dim rawData((ToRow - fromRow + 1), gv.Columns.Count - 1) As Object
            wSheet = wBook.ActiveSheet()

            If clsCommon.myLen(sname) > 31 Then
                sname = sname.Substring(0, 31)
            End If

            'Dim dp As Object
            ' dp = wBook.BuiltinDocumentProperties
            'dp("Tags").Value = clsUserMgtCode.frmComplaintMaster

            wSheet.Name = sname

            Dim flag As Boolean = False
            Dim colnum As Integer = -1
            Dim PrevCol As Integer = -1
            Dim isResteRawData As Boolean = True
            Dim ColNums(0 To gv.Columns.Count - 1) As Integer
            Dim MaxRowsToExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickExport, clsFixedParameterCode.MaxRowsForQuickExport, Nothing))
            Dim jk As Integer = 0

            If isFontSizeIncrease Then
                For i As Integer = arrHeader.Count To gv.Rows.Count + arrHeader.Count + 2
                    'If i = arrHeader.Count + 2 Then
                    '    wSheet.Rows(i).font.bold = True
                    '    wSheet.Rows(i).font.size = 14
                    'Else
                    wSheet.Rows(i).font.bold = True
                    wSheet.Rows(i).font.size = 14
                    'End If
                Next
            End If

            For i As Integer = 0 To gv.Columns.Count - 1
                jk += 1
                ''richa agarwal 12jan BM00000008685
                If AllCellsInString Then
                    If TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                        If gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4152 ' Right Align
                        ElseIf gv.Columns(i).FormatString = "" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4108 ' Center Align
                        Else
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4131 ' Left Align
                        End If
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4152 ' Right Align
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4108 ' Center Align
                    End If
                    wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Borders.LineStyle = True
                Else
                    If FormatCellofExcel = False Then
                        If TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                        End If
                    Else
                        ''richa agarwal 08-jan-2015 format for excel sheet cell format @ as text,General as for General, 0 as numnber,0.00 as decimal against ticket no BM00000008659,BM00000008854
                        If clsCommon.CompairString(gv.Columns(i).Name, "ColAcNo") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "ColPAYEEACNO") = CompairStringResult.Equal Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf clsCommon.CompairString(gv.Columns(i).Name, "colSCACNO") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "colBeniACNO") = CompairStringResult.Equal Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "General"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDateTimeColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "dd/mmm/yyyy"
                        End If
                        ''-------------------
                    End If
                End If
            Next

            Dim colIndex As Integer = 1
            Dim rowIndex As Integer = 1
            If Not IsNothing(arrHeader) Then
                If isArrHeaderMerged Then
                    For Each Str As String In arrHeader
                        Dim startCell As String = "A" & rowIndex.ToString()
                        Dim endCell As String = ColumnIndexToColumnLetter(gv.Columns.Count) & rowIndex.ToString()
                        Dim mergeRange As String = startCell & ":" & endCell
                        excel.Cells(rowIndex, (gv.Columns.Count \ 2)) = Str
                        With excel.Range(mergeRange)
                            .Merge()
                            .HorizontalAlignment = -4108   ' Center Align
                            .Font.Bold = True
                            .Font.Size = 18
                        End With
                        rowIndex += 1
                    Next
                Else
                    For Each Str As String In arrHeader
                        excel.Cells(rowIndex, colIndex) = Str
                        rowIndex += 1
                    Next
                End If
            End If

            'If Not IsNothing(arrHeader) Then
            '    For Each Str As String In arrHeader
            '        excel.Cells(rowIndex, colIndex) = Str
            '        rowIndex += 1
            '    Next
            'End If

            ''richa agarwal 09-feb-2016 to show double header in excel BM00000008811
            'If doubleheadershowninExcel = True Then
            '    Dim view As New ColumnGroupsViewDefinition()
            '    view = gv.ViewDefinition
            '    Dim j1 As Integer = 1
            '    Dim k As Integer = 0
            '    If view.ColumnGroups.Count > 0 Then
            '        Dim chartRange As Excel.Range
            '        For i As Integer = 0 To view.ColumnGroups.Count - 1
            '            excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
            '            Dim l As Integer = 0
            '            For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
            '                'If view.ColumnGroups(i).Rows(0).ColumnNames(m).IsVisible = False Then ''TELERIK2015->2022
            '                '    l = l + 1
            '                'End If

            '                'If view.ColumnGroups(i).IsVisible = False Then ''TELERIK2015->2022
            '                If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
            '                    l = l + 1
            '                End If

            '            Next

            '            k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
            '            chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
            '            'chartRange.Merge()
            '            chartRange.MergeCells = False
            '            chartRange.VerticalAlignment = 2
            '            chartRange.HorizontalAlignment = 3


            '            'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
            '            j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
            '        Next

            '        rowIndex += 1
            '    End If

            'End If

            If doubleheadershowninExcel Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                If view.ColumnGroups.Count > 0 Then
                    Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        Dim visibleColCount As Integer = 0
                        'Count only visible columns
                        For m As Integer = 0 To view.ColumnGroups(i).Rows(0).ColumnNames.Count - 1
                            Dim colName As String = view.ColumnGroups(i).Rows(0).ColumnNames(m)
                            If view.ViewTemplate.Columns.Contains(colName) AndAlso view.ViewTemplate.Columns(colName).IsVisible Then
                                visibleColCount += 1
                            End If
                        Next
                        'Skip if no visible columns
                        If visibleColCount > 0 Then
                            excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                            Dim endCol As Integer = j1 + visibleColCount - 1
                            chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, endCol))
                            'chartRange.Merge()
                            With chartRange
                                .Merge()
                                .WrapText = True
                            End With
                            'Move next start column
                            j1 = endCol + 1
                        End If
                    Next
                    rowIndex += 1
                End If
            End If

            ''------------------------------------------
            While isResteRawData
                If (ToRow - fromRow + 1) <= MaxRowsToExport Then

                    ReDim rawData((ToRow - fromRow), gv.Columns.Count - 1)
                    GridCurrentRowIndex = fromRow - 1
                    GridLastSavedRowIndex = ToRow - 1 ''gv.ChildRows.Count - 1
                    isResteRawData = False
                Else
                    GridCurrentRowIndex = GridLastSavedRowIndex + 1
                    If (GridLastSavedRowIndex + 1) + MaxRowsToExport < (ToRow) Then
                        ReDim rawData(MaxRowsToExport - 1, gv.Columns.Count - 1)
                        GridLastSavedRowIndex = (GridLastSavedRowIndex + (MaxRowsToExport))
                        isResteRawData = True
                    Else
                        ReDim rawData((ToRow - fromRow) - (GridLastSavedRowIndex - fromRow + 1 - 1), gv.Columns.Count - 1)
                        GridLastSavedRowIndex = ToRow - 1
                        isResteRawData = False
                    End If
                    'GridLastSavedRowIndex = (GridLastSavedRowIndex + MaxRowsToExport)
                End If
                Dim RowDataRIndix As Integer = 0
                For i As Integer = GridCurrentRowIndex To GridLastSavedRowIndex
                    'gv.RowCount, gv.Columns.Count - 1
                    dr1 = dset.Tables(0).NewRow
                    clsCommon.ProgressBarPercentUpdate(((i - fromRow + 1 + 1) * 100) / (ToRow - fromRow + 1), " Exporting Record  " & (i - fromRow + 1 + 1) & "  Out of " & (ToRow - fromRow + 1))
                    Try
                        For j As Integer = 0 To gv.Columns.Count - 1
                            dr1(j) = clsCommon.myCstr(gv.ChildRows(i).Cells(j).Value)
                            rawData(RowDataRIndix, j) = dr1(j).ToString()
                        Next
                    Catch ex As Exception
                        RowDataRIndix = RowDataRIndix + 1
                        Exit For
                    End Try
                    dset.Tables(0).Rows.Add(dr1)
                    RowDataRIndix = RowDataRIndix + 1
                Next

                'sanjay
                Try
                    If isResteRawData = False AndAlso MaxRowExport <> ToRow Then
                        If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                            RowDataRIndix = RowDataRIndix - 1
                            dr1 = dset.Tables(0).NewRow
                            For ii As Integer = 0 To gv.MasterTemplate.SummaryRowsBottom(0).Count - 1
                                Dim colName As String = gv.MasterTemplate.SummaryRowsBottom(0)(ii).Name
                                If gv.Columns.Contains(colName) AndAlso gv.Columns(colName).IsVisible Then
                                    Dim summaryItem As GridViewSummaryItem = gv.SummaryRowsBottom(0)(ii)

                                    Dim summary As Object = Nothing

                                    Try
                                        summary = summaryItem.Evaluate(gv.MasterTemplate)
                                    Catch ex As Exception
                                    End Try
                                    dr1(gv.Columns(colName).Index) = IIf(((summary = Nothing) Or (Single.NaN.Equals(summary))), DBNull.Value, clsCommon.myCstr(summary))
                                    rawData(RowDataRIndix, gv.Columns(colName).Index) = clsCommon.myCstr(dr1(gv.Columns(colName).Index).ToString())
                                End If
                            Next

                            dset.Tables(0).Rows.Add(dr1)
                        End If
                    End If
                Catch ex As Exception
                End Try
                'sanjay

                Try
                    CType(wBook.Sheets("Sheet2"), Excel.Worksheet).Delete()
                    CType(wBook.Sheets("Sheet3"), Excel.Worksheet).Delete()
                Catch ex As Exception
                End Try

                Dim dt As System.Data.DataTable = dset.Tables(0)
                Dim dc As System.Data.DataColumn
                ' Dim dr As System.Data.DataRow
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex, colIndex) = dc.Caption
                Next

                Dim LastColumn As String = ColumnIndexToColumnLetter(dt.Columns.Count)
                Dim Lastrow As Integer = dt.Rows.Count + 1 ''change

                Dim row As Integer = 0
                Dim col As Integer = 0
                If Not isblanksheet Then
                    wSheet.Range("A" & ((GridCurrentRowIndex - fromRow + 1) + rowIndex + 1), LastColumn & ((GridLastSavedRowIndex - fromRow + 1) + rowIndex + 1)).Value2 = rawData
                End If
                rawData = Nothing
                dt = Nothing
                ''new code by Panch Raj
                dset.Tables(0).Rows.Clear()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                If (ToRow - fromRow + 1) - 1 > (GridLastSavedRowIndex - fromRow + 1) Then
                    isResteRawData = True
                Else
                    isResteRawData = False
                End If
            End While

            'Hide Group columns
            If doubleheadershowninExcel = True Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                Dim k As Integer = 0
                If view.ColumnGroups.Count > 0 Then
                    'Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        'excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                        Dim l As Integer = 0
                        For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
                            'If view.ColumnGroups(i).Rows(0).Columns(m).IsVisible = False Then''TELERIK2015->2022
                            'l = l + 1
                            'End If
                            'If view.ColumnGroups(i).IsVisible = False Then
                            If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
                                l = l + 1
                            End If
                        Next

                        k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                        'chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
                        'chartRange.Merge()
                        'chartRange.VerticalAlignment = 2
                        'chartRange.HorizontalAlignment = 3
                        If view.ColumnGroups(i).IsVisible = False Then
                            'chartRange.EntireColumn.Hidden = True
                            wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k)).EntireColumn.Hidden = True
                        End If

                        'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
                        j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                    Next

                    rowIndex += 1
                End If

            End If

            clsCommon.ProgressBarPercentUpdate(100, "Wait Adjusting Columns Autofit ")

            If doubleheadershowninExcel = False Then
                wSheet.Columns.AutoFit()
            End If

            If AllCellsInString Then
                wSheet.Rows.AutoFit()
                wSheet.Columns.AutoFit()
            End If

            'Manadatory Field Coloring
            If manadatoryField IsNot Nothing AndAlso manadatoryField.Count > 0 Then
                For c As Integer = 0 To wSheet.Columns.Count - 1
                    If clsCommon.myLen(wSheet.Cells(1, c + 1).value) <= 0 Then
                        Exit For
                    End If

                    If isManadatory(wSheet.Cells(1, c + 1).value, manadatoryField) Then
                        wSheet.Cells(1, c + 1).interior.color = RGB(Color.LightGoldenrodYellow.R, Color.LightGoldenrodYellow.G, Color.LightGoldenrodYellow.B)
                    End If
                Next

                For i As Integer = 0 To gv.Columns.Count - 1
                    If Not gv.Columns(i).IsVisible Then
                        wSheet.Columns(i + 1).EntireColumn.Hidden = True
                    End If
                Next
                'oWB.SaveAs(filePath)
                'oWB.Close()
                'oApp.Quit()

            End If


            Dim strFileName As String = flname
            Dim blnFileOpen As Boolean = False
            clsCommon.ProgressBarPercentUpdate(100, "Wait Saving Excel file in expected Format ")
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()

            Catch ex As Exception
                blnFileOpen = False
            End Try


            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If


            clsCommon.ProgressBarPercentUpdate(100, "Wait Opening Excel file ")
            wBook.SaveAs(strFileName)
            wBook.Close(True)

            clsCommon.ProgressBarPercentHide()
            wBook = Nothing
            wSheet = Nothing
            excel.Quit()
            excel = Nothing
            rawData = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()

            'sanjay
            If MultipleFiles = False Then
                If FileCount > 1 Then
                    common.clsCommon.MyMessageBoxShow(gv, "Data Exported in directory -" & System.IO.Path.GetDirectoryName(flname) & "\" & System.IO.Path.GetFileName(flname) & " in " & FileCount & " files.")
                Else
                    common.clsCommon.MyMessageBoxShow(gv, "Exported Successfully.")
                    If clsCommon.myLen(gv.MyExportFilePath) <= 0 Then
                        Process.Start(flname)
                    Else
                        gv.MyExportFilePath = flname
                    End If
                End If
            End If
            'sanjay

        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try

        Return FileCount
    End Function

    Public Function exportdataBoothSlipGNG(ByVal MaxRowExport As Integer, ByVal gv As MyRadGridView, ByVal flname As String, ByVal sname As String, ByVal fromRow As Integer, ToRow As Integer, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ExportWithoutHeader As Boolean = False, Optional FormatCellofExcel As Boolean = False, Optional doubleheadershowninExcel As Boolean = False, Optional ByVal MultipleFiles As Boolean = False, Optional ByVal UseFilePath As Boolean = False, Optional ByVal manadatoryField As List(Of String) = Nothing, Optional ByVal AllCellsInString As Boolean = False, Optional ByVal BoothSlipFormat As Boolean = False) As Integer
        Dim FileCount As Integer = 1
        If AllCellsInString Then
            MaxRowExport = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsInExcelExport, clsFixedParameterCode.MaxRowsInExcelExport, Nothing))
        End If
        'sanjay
        If sname.Contains("/") Then
            sname = sname.Replace("/", " ")
        End If
        If sname.Contains("\") Then
            sname = sname.Replace("\", " ")
        End If
        'If UseFilePath = False And MultipleFiles = False Then
        '    If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ExportToDefineLocation, clsFixedParameterCode.ExportToDefineLocation, Nothing)) = 1 Then
        '        UseFilePath = True
        '        Dim sfd As SaveFileDialog = New SaveFileDialog()
        '        sfd.FileName = sname
        '        'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 *.xlsx|(*.xlsx);|CSV Files (*.csv) |*.csv"
        '        sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
        '        If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '            flname = sfd.FileName
        '            sname = flname.Substring(flname.LastIndexOf("\") + 1, flname.Length - flname.LastIndexOf("\") - 1)
        '        Else
        '            Return False
        '        End If
        '    End If
        'End If
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
        If IsExists = False Then
            System.IO.Directory.CreateDirectory(forlderName)
        End If

        If UseFilePath = False And MultipleFiles = False Then
            Dim currTime As DateTime = DateTime.Now
            sname = sname & clsCommon.GetPrintDate(currTime, "yyyyMMddhhmmss")
            flname = forlderName & "\" & sname & ".xls"
            sname = sname & ".xls"
        End If
        'sanjay

        Try
            '==========================Add isblanksheet Variable in Function to Save Blank Excel Sheet===================
            If ((gv.Columns.Count = 0) Or (gv.ChildRows.Count = 0)) And isblanksheet = False Then
                Return FileCount
            End If
            If fromRow <= 0 Then
                fromRow = 1
            End If
            If ToRow <= 0 Then
                ToRow = gv.Rows.Count
            End If
            '========================================================================================
            '' check rows limit on work sheet

            If MaxRowExport <= 0 Then
                MaxRowExport = 1048576
            End If
            If gv.Rows.Count >= MaxRowExport And MultipleFiles = False Then
                MultipleFiles = True
                Dim FilePathRoot As String = System.IO.Path.GetDirectoryName(flname)
                Dim fileName As String = System.IO.Path.GetFileName(flname)
                Dim fileExtn As String = System.IO.Path.GetExtension(flname)
                fileName = fileName.Replace(fileExtn, "")
                flname = FilePathRoot & "\" & fileName
                If System.IO.Directory.Exists(flname) = False Then
                    System.IO.Directory.CreateDirectory(flname)
                End If

                'Dim tblArr() As DataTable
                Dim tableCount = Math.Ceiling(gv.Rows.Count / MaxRowExport)
                Dim Divisor = gv.Rows.Count / tableCount
                '' for data boun through data reader
                Dim fromCount As Integer = 1
                Dim ToCount As Integer = gv.Rows.Count
                FileCount = tableCount
                For intLoop As Integer = 0 To tableCount - 1
                    Dim file As String
                    If intLoop = tableCount - 1 Then
                        ToCount = gv.Rows.Count
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdata(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = fromCount + (MaxRowExport)
                        common.clsCommon.MyMessageBoxShow("Data Exported in directory -" & System.IO.Path.GetDirectoryName(flname) & "\" & System.IO.Path.GetFileName(flname) & " in " & FileCount & " files.")
                    Else
                        ToCount = (fromCount + 0) + (MaxRowExport - 1)
                        file = flname & "\" & fileName & intLoop + 1 & fileExtn
                        sname = fileName & intLoop + 1 & fileExtn
                        'IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        exportdata(gv, file, sname, fromCount, ToCount, isblanksheet, arrHeader, ExportWithoutHeader, FormatCellofExcel, doubleheadershowninExcel, MultipleFiles, UseFilePath, manadatoryField)
                        fromCount = (fromCount) + (MaxRowExport)
                    End If
                Next
                Return FileCount
                'Throw New Exception("Max row limit for worksheet in excel is 1048576")
            End If
            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()

            'Do not remove column for master Export/Import
            If Not (manadatoryField IsNot Nothing AndAlso manadatoryField.Count > 0) Then
                Dim IndexList As List(Of String) = New List(Of String)

                For i As Integer = 0 To gv.Columns.Count - 1
                    If Not gv.Columns(i).IsVisible Then
                        IndexList.Add(gv.Columns(i).Name)
                    End If
                Next

                For i As Integer = 0 To IndexList.Count - 1
                    gv.Columns.Remove(IndexList.Item(i).ToString())
                Next
            End If

            clsCommon.ProgressBarPercentShow()
            'add column to that table            
            For i As Integer = 0 To gv.ColumnCount - 1
                'dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
                dset.Tables(0).Columns.Add("Column" & (i + 1))
                dset.Tables(0).Columns("Column" & (i + 1)).Caption = IIf(ExportWithoutHeader, "", gv.Columns(i).HeaderText)
            Next
            'add rows to the table
            'dset.Tables(0).Rows(0).Delete()

            Dim dr1 As DataRow


            'For row = 0 To dt.Rows.Count - 1
            '    For col = 0 To dt.Columns.Count - 1
            '        rawData(row, col) = dt.Rows(row).ItemArray(col)
            '    Next
            '    clsCommon.ProgressBarPercentUpdate((row * 100) / dt.Rows.Count, " Exporting Record  " & row & "  Out of " & dt.Rows.Count)
            'Next
            Dim excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            Dim GridCurrentRowIndex As Int64 = fromRow - 2
            Dim GridLastSavedRowIndex As Int64 = fromRow - 2

            wBook = excel.Workbooks.Add()

            'sanjay for summary row
            If MaxRowExport <> ToRow Then
                If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                    ToRow = ToRow + 1
                End If
            End If
            'sanjay for summary row

            Dim rawData((ToRow - fromRow + 1), gv.Columns.Count - 1) As Object
            wSheet = wBook.ActiveSheet()

            If clsCommon.myLen(sname) > 31 Then
                sname = sname.Substring(0, 31)
            End If

            'Dim dp As Object
            ' dp = wBook.BuiltinDocumentProperties
            'dp("Tags").Value = clsUserMgtCode.frmComplaintMaster

            wSheet.Name = sname

            Dim flag As Boolean = False
            Dim colnum As Integer = -1
            Dim PrevCol As Integer = -1
            Dim isResteRawData As Boolean = True
            Dim ColNums(0 To gv.Columns.Count - 1) As Integer
            Dim MaxRowsToExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.QuickExport, clsFixedParameterCode.MaxRowsForQuickExport, Nothing))
            Dim jk As Integer = 0

            If BoothSlipFormat Then
                For i As Integer = arrHeader.Count + 2 To gv.Rows.Count + arrHeader.Count + 1
                    If i = arrHeader.Count + 2 Then
                        wSheet.Rows(i).font.bold = True
                        wSheet.Rows(i).font.size = 18
                    Else
                        wSheet.Rows(i).font.size = 18
                    End If
                Next
            End If


            For i As Integer = 0 To gv.Columns.Count - 1
                jk += 1
                ''richa agarwal 12jan BM00000008685
                If AllCellsInString Then
                    If TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                        If gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "General"
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4152 ' Right Align
                        Else
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "General"
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4131 ' Left Align
                        End If
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4152 ' Right Align
                    ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).HorizontalAlignment = -4108 ' Center Align
                    End If
                    If BoothSlipFormat Then
                        wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Borders.LineStyle = True
                    End If
                Else
                    If FormatCellofExcel = False Then
                        If TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                        End If
                    Else
                        ''richa agarwal 08-jan-2015 format for excel sheet cell format @ as text,General as for General, 0 as numnber,0.00 as decimal against ticket no BM00000008659,BM00000008854
                        If clsCommon.CompairString(gv.Columns(i).Name, "ColAcNo") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "ColPAYEEACNO") = CompairStringResult.Equal Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf clsCommon.CompairString(gv.Columns(i).Name, "colSCACNO") = CompairStringResult.Equal Or clsCommon.CompairString(gv.Columns(i).Name, "colBeniACNO") = CompairStringResult.Equal Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "@"
                        ElseIf TypeOf gv.Columns(i) Is GridViewTextBoxColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "General"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn AndAlso gv.Columns(i).FormatString = "{0:n2}" Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0.00"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDecimalColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "0"
                        ElseIf TypeOf gv.Columns(i) Is GridViewDateTimeColumn Then
                            wSheet.Range(ColumnIndexToColumnLetter(jk) & ":" & ColumnIndexToColumnLetter(jk)).Cells.NumberFormat = "dd/mmm/yyyy"
                        End If
                        ''-------------------
                    End If
                End If
            Next

            Dim colIndex As Integer = 1
            Dim rowIndex As Integer = 1

            If Not IsNothing(arrHeader) Then
                For Each Str As String In arrHeader
                    If BoothSlipFormat Then
                        Dim startCell As String = "A" & rowIndex.ToString()
                        Dim endCell As String = ColumnIndexToColumnLetter(gv.Columns.Count) & rowIndex.ToString()
                        Dim mergeRange As String = startCell & ":" & endCell
                        If rowIndex = 1 Then
                            excel.Cells(rowIndex, (gv.Columns.Count / 2)) = Str
                            excel.Rows(rowIndex).Font.Size = 26
                            excel.Rows(rowIndex).Font.Bold = True
                            excel.Range(mergeRange).Merge()
                            excel.Range(mergeRange).HorizontalAlignment = -4108 ' Center Align
                        ElseIf rowIndex = 2 Then
                            excel.Cells(rowIndex, (gv.Columns.Count / 2)) = Str
                            excel.Rows(rowIndex).Font.Size = 18
                            excel.Rows(rowIndex).Font.Bold = True
                            excel.Range(mergeRange).Merge()
                            excel.Range(mergeRange).HorizontalAlignment = -4108 ' Center Align
                        Else
                            excel.Cells(rowIndex, colIndex) = Str
                        End If
                    Else
                        excel.Cells(rowIndex, colIndex) = Str
                    End If
                    rowIndex += 1
                Next
            End If

            ''richa agarwal 09-feb-2016 to show double header in excel BM00000008811
            If doubleheadershowninExcel = True Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                Dim k As Integer = 0
                If view.ColumnGroups.Count > 0 Then
                    Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                        Dim l As Integer = 0
                        For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
                            'If view.ColumnGroups(i).Rows(0).ColumnNames(m).IsVisible = False Then ''TELERIK2015->2022
                            '    l = l + 1
                            'End If

                            'If view.ColumnGroups(i).IsVisible = False Then ''TELERIK2015->2022
                            If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
                                l = l + 1
                            End If

                        Next

                        k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                        chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
                        chartRange.Merge()
                        chartRange.VerticalAlignment = 2
                        chartRange.HorizontalAlignment = 3


                        'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
                        j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                    Next

                    rowIndex += 1
                End If

            End If
            ''------------------------------------------
            While isResteRawData
                If (ToRow - fromRow + 1) <= MaxRowsToExport Then

                    ReDim rawData((ToRow - fromRow), gv.Columns.Count - 1)
                    GridCurrentRowIndex = fromRow - 1
                    GridLastSavedRowIndex = ToRow - 1 ''gv.ChildRows.Count - 1
                    isResteRawData = False
                Else
                    GridCurrentRowIndex = GridLastSavedRowIndex + 1
                    If (GridLastSavedRowIndex + 1) + MaxRowsToExport < (ToRow) Then
                        ReDim rawData(MaxRowsToExport - 1, gv.Columns.Count - 1)
                        GridLastSavedRowIndex = (GridLastSavedRowIndex + (MaxRowsToExport))
                        isResteRawData = True
                    Else
                        ReDim rawData((ToRow - fromRow) - (GridLastSavedRowIndex - fromRow + 1 - 1), gv.Columns.Count - 1)
                        GridLastSavedRowIndex = ToRow - 1
                        isResteRawData = False
                    End If
                    'GridLastSavedRowIndex = (GridLastSavedRowIndex + MaxRowsToExport)
                End If
                Dim RowDataRIndix As Integer = 0
                For i As Integer = GridCurrentRowIndex To GridLastSavedRowIndex
                    'gv.RowCount, gv.Columns.Count - 1
                    dr1 = dset.Tables(0).NewRow
                    clsCommon.ProgressBarPercentUpdate(((i - fromRow + 1 + 1) * 100) / (ToRow - fromRow + 1), " Exporting Record  " & (i - fromRow + 1 + 1) & "  Out of " & (ToRow - fromRow + 1))
                    Try
                        For j As Integer = 0 To gv.Columns.Count - 1
                            dr1(j) = clsCommon.myCstr(gv.ChildRows(i).Cells(j).Value)
                            rawData(RowDataRIndix, j) = dr1(j).ToString()
                        Next
                    Catch ex As Exception
                        RowDataRIndix = RowDataRIndix + 1
                        Exit For
                    End Try
                    dset.Tables(0).Rows.Add(dr1)
                    RowDataRIndix = RowDataRIndix + 1
                Next

                'sanjay
                Try
                    If isResteRawData = False AndAlso MaxRowExport <> ToRow Then
                        If (gv.MasterTemplate.SummaryRowsBottom.Count > 0) Then
                            RowDataRIndix = RowDataRIndix - 1
                            dr1 = dset.Tables(0).NewRow
                            For ii As Integer = 0 To gv.MasterTemplate.SummaryRowsBottom(0).Count - 1
                                Dim colName As String = gv.MasterTemplate.SummaryRowsBottom(0)(ii).Name
                                If gv.Columns.Contains(colName) AndAlso gv.Columns(colName).IsVisible Then
                                    Dim summaryItem As GridViewSummaryItem = gv.SummaryRowsBottom(0)(ii)

                                    Dim summary As Object = Nothing

                                    Try
                                        summary = summaryItem.Evaluate(gv.MasterTemplate)
                                    Catch ex As Exception
                                    End Try
                                    dr1(gv.Columns(colName).Index) = IIf(((summary = Nothing) Or (Single.NaN.Equals(summary))), DBNull.Value, clsCommon.myCstr(summary))
                                    rawData(RowDataRIndix, gv.Columns(colName).Index) = clsCommon.myCstr(dr1(gv.Columns(colName).Index).ToString())
                                End If
                            Next

                            dset.Tables(0).Rows.Add(dr1)
                        End If
                    End If
                Catch ex As Exception
                End Try
                'sanjay

                Try
                    CType(wBook.Sheets("Sheet2"), Excel.Worksheet).Delete()
                    CType(wBook.Sheets("Sheet3"), Excel.Worksheet).Delete()
                Catch ex As Exception
                End Try

                Dim dt As System.Data.DataTable = dset.Tables(0)
                Dim dc As System.Data.DataColumn
                ' Dim dr As System.Data.DataRow
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex, colIndex) = dc.Caption
                Next

                Dim LastColumn As String = ColumnIndexToColumnLetter(dt.Columns.Count)
                Dim Lastrow As Integer = dt.Rows.Count + 1 ''change

                Dim row As Integer = 0
                Dim col As Integer = 0
                If Not isblanksheet Then
                    wSheet.Range("A" & ((GridCurrentRowIndex - fromRow + 1) + rowIndex + 1), LastColumn & ((GridLastSavedRowIndex - fromRow + 1) + rowIndex + 1)).Value2 = rawData
                End If
                rawData = Nothing
                dt = Nothing
                ''new code by Panch Raj
                dset.Tables(0).Rows.Clear()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                If (ToRow - fromRow + 1) - 1 > (GridLastSavedRowIndex - fromRow + 1) Then
                    isResteRawData = True
                Else
                    isResteRawData = False
                End If
            End While

            'Hide Group columns
            If doubleheadershowninExcel = True Then
                Dim view As New ColumnGroupsViewDefinition()
                view = gv.ViewDefinition
                Dim j1 As Integer = 1
                Dim k As Integer = 0
                If view.ColumnGroups.Count > 0 Then
                    'Dim chartRange As Excel.Range
                    For i As Integer = 0 To view.ColumnGroups.Count - 1
                        'excel.Cells(rowIndex, j1) = clsCommon.myCstr(view.ColumnGroups(i).Text)
                        Dim l As Integer = 0
                        For m As Integer = 0 To clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - 1
                            'If view.ColumnGroups(i).Rows(0).Columns(m).IsVisible = False Then''TELERIK2015->2022
                            'l = l + 1
                            'End If
                            'If view.ColumnGroups(i).IsVisible = False Then
                            If view.ViewTemplate.Columns.Contains(view.ColumnGroups(i).Rows(0).ColumnNames(m)) = False Then
                                l = l + 1
                            End If
                        Next

                        k = k + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                        'chartRange = wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k))
                        'chartRange.Merge()
                        'chartRange.VerticalAlignment = 2
                        'chartRange.HorizontalAlignment = 3
                        If view.ColumnGroups(i).IsVisible = False Then
                            'chartRange.EntireColumn.Hidden = True
                            wSheet.Range(wSheet.Cells(rowIndex, j1), wSheet.Cells(rowIndex, k)).EntireColumn.Hidden = True
                        End If

                        'j1 = clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).Columns.Count) + 1 - l
                        j1 = j1 + clsCommon.myCdbl(view.ColumnGroups(i).Rows(0).ColumnNames.Count) - l
                    Next

                    rowIndex += 1
                End If

            End If

            clsCommon.ProgressBarPercentUpdate(100, "Wait Adjusting Columns Autofit ")

            If doubleheadershowninExcel = False Then
                wSheet.Columns.AutoFit()
            End If

            If AllCellsInString Then
                wSheet.Columns.AutoFit()
            End If





            'Manadatory Field Coloring
            If manadatoryField IsNot Nothing AndAlso manadatoryField.Count > 0 Then
                For c As Integer = 0 To wSheet.Columns.Count - 1
                    If clsCommon.myLen(wSheet.Cells(1, c + 1).value) <= 0 Then
                        Exit For
                    End If

                    If isManadatory(wSheet.Cells(1, c + 1).value, manadatoryField) Then
                        wSheet.Cells(1, c + 1).interior.color = RGB(Color.LightGoldenrodYellow.R, Color.LightGoldenrodYellow.G, Color.LightGoldenrodYellow.B)
                    End If
                Next

                For i As Integer = 0 To gv.Columns.Count - 1
                    If Not gv.Columns(i).IsVisible Then
                        wSheet.Columns(i + 1).EntireColumn.Hidden = True
                    End If
                Next
                'oWB.SaveAs(filePath)
                'oWB.Close()
                'oApp.Quit()

            End If


            Dim strFileName As String = flname
            Dim blnFileOpen As Boolean = False
            clsCommon.ProgressBarPercentUpdate(100, "Wait Saving Excel file in expected Format ")
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()

            Catch ex As Exception
                blnFileOpen = False
            End Try


            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            clsCommon.ProgressBarPercentUpdate(100, "Wait Opening Excel file ")
            wBook.SaveAs(strFileName)
            wBook.Close(True)

            clsCommon.ProgressBarPercentHide()
            wBook = Nothing
            wSheet = Nothing
            excel.Quit()
            excel = Nothing
            rawData = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()

            'sanjay
            If MultipleFiles = False Then
                If FileCount > 1 Then
                    common.clsCommon.MyMessageBoxShow(gv, "Data Exported in directory -" & System.IO.Path.GetDirectoryName(flname) & "\" & System.IO.Path.GetFileName(flname) & " in " & FileCount & " files.")
                Else
                    common.clsCommon.MyMessageBoxShow(gv, "Exported Successfully.")
                    If clsCommon.myLen(gv.MyExportFilePath) <= 0 Then
                        Process.Start(flname)
                    Else
                        gv.MyExportFilePath = flname
                    End If
                End If
            End If
            'sanjay

        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try

        Return FileCount
    End Function


    Public Function QuickExportToExcel(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, Optional ByVal isblanksheet As Boolean = False, Optional ByVal arrHeader As List(Of String) = Nothing, Optional ByVal UseFilePath As Boolean = False) As Integer
        Try
            Return exportdata(gv, flname, sname, isblanksheet, arrHeader, False, False, False, UseFilePath)
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try
    End Function


    Private Function ColumnIndexToColumnLetter(ByVal colIndex As Integer) As String
        Dim div As Integer = colIndex
        Dim colLetter As String = [String].Empty
        Dim [mod] As Integer = 0
        While div > 0
            [mod] = (div - 1) Mod 26
            colLetter = (Convert.ToChar(65 + [mod])).ToString & colLetter
            div = CInt((div - [mod]) / 26)
        End While
        Return colLetter
    End Function
    Public Sub exportdataInCSV(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, Optional ByVal isblanksheet As Boolean = False)
        '==========================Add isblanksheet Variable in Function to Save Blank Excel Sheet===================
        If ((gv.Columns.Count = 0) Or (gv.Rows.Count = 0)) And isblanksheet = False Then
            Exit Sub
        End If
        '========================================================================================
        'Creating dataset to export
        Dim dset As New DataSet
        'add table to dataset
        dset.Tables.Add()
        'add column to that table
        For i As Integer = 0 To gv.ColumnCount - 1
            dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr1 As DataRow
        For i As Integer = 0 To gv.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To gv.Columns.Count - 1
                dr1(j) = gv.Rows(i).Cells(j).Value.ToString
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim strData As String = String.Empty
        For Each dc In dt.Columns
            colIndex = colIndex + 1
            strData = strData & dc.ColumnName
            If colIndex <> (dt.Columns.Count) Then
                strData = strData & ","
            Else
                strData = strData & Environment.NewLine
            End If
        Next
        clsCommon.ProgressBarPercentShow()
        For Each dr In dt.Rows
            rowIndex = rowIndex + 1

            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                'excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName).ToString
                strData = strData & dr(dc.ColumnName).ToString
                If colIndex <> (dt.Columns.Count) Then
                    strData = strData & ","
                Else
                    strData = strData & Environment.NewLine
                End If
            Next
            clsCommon.ProgressBarPercentUpdate(rowIndex * 100 / dt.Rows.Count, "Exporting " + clsCommon.myCstr(rowIndex) + "/" + clsCommon.myCstr(dt.Rows.Count))
        Next
        clsCommon.ProgressBarPercentHide()
        File.WriteAllText(flname, strData)
    End Sub
    ''''''Function to export data with custom filed
    Public Sub exportdatawithcustomfield(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, ByVal formid As String)
        If ((gv.Columns.Count = 0) Or (gv.Rows.Count = 0)) Then
            Exit Sub
        End If

        'Creating dataset to export
        Dim dset As New DataSet
        'add table to dataset
        dset.Tables.Add()
        'add column to that table
        For i As Integer = 0 To gv.ColumnCount - 1
            dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
        Next
        Dim drr As DataTable
        Dim qry1 As String = "select TSPL_CUSTOM_FIELD_HEAD.Name,TSPL_CUSTOM_FIELD_HEAD.Code,TSPL_CUSTOM_FIELD_MAPPING.Program_Code  from TSPL_CUSTOM_FIELD_HEAD left outer join TSPL_CUSTOM_FIELD_MAPPING on TSPL_CUSTOM_FIELD_MAPPING.Custom_Field_Code =TSPL_CUSTOM_FIELD_HEAD.Code WHERE PROGRAM_CODE='" & formid & "'"
        drr = clsDBFuncationality.GetDataTable(qry1)
        Dim j As Integer = gv.Columns.Count
        'Dim i As Integer = 0
        For Each row As DataRow In drr.Rows
            dset.Tables(0).Columns.Add(row(1).ToString)
            dset.Tables(0).Columns(j).ColumnName = row(1).ToString
            dset.Tables(0).Columns(j).Caption = row(0).ToString
            j = j + 1
        Next


        'add rows to the table
        Dim dr1 As DataRow
        For i As Integer = 0 To gv.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j = 0 To gv.Columns.Count - 1
                dr1(j) = gv.Rows(i).Cells(j).Value.ToString
            Next
            For j = gv.Columns.Count To dset.Tables(0).Columns.Count - 1

                Dim q1 As String = "select value from tspl_custom_field_values where program_code='" & formid & "' and transaction_code='" & dr1(0).ToString & "' and custom_field_code='" & dset.Tables(0).Columns(j).ColumnName & "'"

                dr1(j) = clsDBFuncationality.getSingleValue(q1)
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()

        wSheet = wBook.ActiveSheet()
        wSheet.Cells.NumberFormat = "@"
        wSheet.Name = sname
        CType(wBook.Sheets("Sheet2"), Excel.Worksheet).Delete()
        CType(wBook.Sheets("Sheet3"), Excel.Worksheet).Delete()
        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            If colIndex >= gv.Columns.Count Then
                excel.Cells(1, colIndex) = dc.Caption
            Else
                excel.Cells(1, colIndex) = dc.ColumnName
            End If

        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName).ToString

            Next
        Next

        wSheet.Columns.AutoFit()
        Dim strFileName As String = flname
        Dim blnFileOpen As Boolean = False
        Try
            Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
            fileTemp.Close()
        Catch ex As Exception
            blnFileOpen = False
        End Try

        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)

        End If


        wBook.SaveAs(strFileName)

        excel.DisplayAlerts = False
        wBook.Close()

    End Sub

    Private Sub exporter_ExcelCellFormatting(ByVal sender As Object, ByVal e As ExcelML.ExcelCellFormattingEventArgs)
        If e.GridRowInfoType Is GetType(GridViewTableHeaderRowInfo) Then
            e.ExcelStyleElement.FontStyle.Bold = True
            e.ExcelStyleElement.FontStyle.Size = 10
        ElseIf e.GridRowInfoType Is GetType(GridViewDataRowInfo) Then
            e.ExcelStyleElement.FontStyle.Size = 9
        End If
    End Sub
    Private Sub exporter_ExcelCellFormattingForOpen(ByVal sender As Object, ByVal e As ExcelML.ExcelCellFormattingEventArgs)
        If e.GridRowInfoType Is GetType(GridViewTableHeaderRowInfo) Then
            e.ExcelStyleElement.FontStyle.Bold = True
            e.ExcelStyleElement.FontStyle.Size = 10
            e.ExcelStyleElement.InteriorStyle.Color = Color.LightBlue
            e.ExcelStyleElement.FontStyle.Color = Color.AliceBlue
        ElseIf e.GridRowInfoType Is GetType(GridViewDataRowInfo) Then
            e.ExcelStyleElement.FontStyle.Size = 9
        End If
    End Sub

    Dim count As Integer = 0

    Public Function importExcelForItemMaster(ByVal strSheetName As String, ByVal gv As RadGridView, ByVal gv1 As RadGridView, ByVal gv2 As RadGridView, ByVal filePath As String, ByVal ParamArray fieldNames As String()) As Boolean
        count += 1
        Try
            Dim Extension As String = Path.GetExtension(filePath)
            Dim conStr As String = ""
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;"
            conStr = [String].Format(conStr, filePath)
            Dim connExcel As New OleDbConnection(conStr)
            Dim cmdExcel As New OleDbCommand()
            Dim oda As New OleDbDataAdapter()
            Dim ds As New DataTable()
            cmdExcel.Connection = connExcel

            'Get the name of First Sheet  
            connExcel.Open()
            Dim dtExcelSchema As DataTable
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim SheetName As String
            Dim isFound As Boolean = False
            For Each dr As DataRow In dtExcelSchema.Rows
                SheetName = clsCommon.myCstr(dr("TABLE_NAME"))

                If ((clsCommon.CompairString(SheetName, strSheetName) = CompairStringResult.Equal) AndAlso (clsCommon.CompairString(SheetName, "ItemMaster$") = CompairStringResult.Equal)) Then
                    isFound = True
                    connExcel.Close()
                    connExcel.Open()
                    cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
                    oda.SelectCommand = cmdExcel
                    oda.Fill(ds)
                    connExcel.Close()
                    gv.DataSource = ds.DefaultView
                    gv.AllowColumnReorder = True
                    Dim fieldCount As Integer = fieldNames.Length
                    Dim strfields As String = ""
                    For Each field As String In fieldNames
                        strfields = strfields + field + ","
                    Next
                    If fieldCount <= gv.ColumnCount Then
                        Dim i As Integer = 0
                        Dim arr As ArrayList = New ArrayList()
                        For Each GC As GridViewColumn In gv.Columns
                            arr.Add(GC.HeaderText)
                        Next
                        For Each field As String In fieldNames
                            If arr.Contains(field) Then
                                For Each GC As GridViewColumn In gv.Columns
                                    If GC.HeaderText = field Then
                                        gv.Columns.Move(GC.Index, i)
                                        Exit For
                                    End If
                                Next
                            Else
                                Throw New Exception("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
                            End If
                            i = i + 1
                        Next
                    Else
                        Throw New Exception("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
                    End If
                    connExcel.Close()
                ElseIf ((clsCommon.CompairString(SheetName, strSheetName) = CompairStringResult.Equal) AndAlso (clsCommon.CompairString(SheetName, "ItemDetails$") = CompairStringResult.Equal)) Then
                    isFound = True
                    connExcel.Close()
                    'Read Data from First Sheet  
                    connExcel.Open()
                    cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
                    oda.SelectCommand = cmdExcel
                    ds.Dispose()

                    oda.Fill(ds)
                    connExcel.Close()
                    gv1.DataSource = ds.DefaultView
                    gv1.AllowColumnReorder = True

                    Dim fieldCount As Integer = fieldNames.Length
                    Dim strfields As String = ""
                    For Each field As String In fieldNames
                        strfields = strfields + field + ","
                    Next
                    If fieldCount <= gv1.ColumnCount Then
                        Dim i As Integer = 0
                        Dim arr As ArrayList = New ArrayList()
                        For Each GC As GridViewColumn In gv1.Columns
                            arr.Add(GC.HeaderText)
                        Next
                        For Each field As String In fieldNames
                            If arr.Contains(field) Then
                                For Each GC As GridViewColumn In gv1.Columns
                                    If GC.HeaderText = field Then
                                        gv1.Columns.Move(GC.Index, i)
                                        Exit For
                                    End If
                                Next
                            Else
                                Throw New Exception("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
                            End If
                            i = i + 1
                        Next
                    Else
                        Throw New Exception("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
                    End If
                    connExcel.Close()
                ElseIf ((clsCommon.CompairString(SheetName, strSheetName) = CompairStringResult.Equal) AndAlso (clsCommon.CompairString(SheetName, "ItemUOMDetails$") = CompairStringResult.Equal)) Then
                    isFound = True
                    connExcel.Close()
                    connExcel.Open()
                    cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
                    oda.SelectCommand = cmdExcel
                    oda.Fill(ds)
                    connExcel.Close()
                    gv2.DataSource = ds.DefaultView
                    gv2.AllowColumnReorder = True
                    Dim fieldCount As Integer = fieldNames.Length
                    Dim strfields As String = ""
                    For Each field As String In fieldNames
                        strfields = strfields + field + ","
                    Next
                    If fieldCount <= gv2.ColumnCount Then
                        Dim i As Integer = 0
                        Dim arr As ArrayList = New ArrayList()
                        For Each GC As GridViewColumn In gv2.Columns
                            arr.Add(GC.HeaderText)
                        Next
                        For Each field As String In fieldNames
                            If arr.Contains(field) Then
                                For Each GC As GridViewColumn In gv2.Columns
                                    If GC.HeaderText = field Then
                                        gv2.Columns.Move(GC.Index, i)
                                        Exit For
                                    End If
                                Next
                            Else
                                Throw New Exception("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
                            End If
                            i = i + 1
                        Next
                    Else
                        Throw New Exception("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
                    End If
                End If
            Next
            If isFound Then
                Return True
            Else
                Throw New Exception(strSheetName + "Excel Sheet not found- ")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Function getColumns(ByVal fileName As String) As String()
        Try
            Dim fileReader As New StreamReader(fileName)
            Dim line As String = fileReader.ReadLine
            fileReader.Close()
            Dim Columns() As String = line.Split(",")
            Return Columns
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function ReturnData(ByVal fileName As String) As DataTable
        Try
            Dim dt As New DataTable
            For Each columnName As String In getColumns(fileName)
                dt.Columns.Add(columnName)
            Next
            Dim fileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(fileName)
            ' If ColumnNames Then
            fileReader.SetDelimiters(",")
            'fileReader.ReadLine()
            fileReader.ReadFields()
            'End If
            Dim line() As String
            While Not fileReader.EndOfData
                'line = line.Replace(Chr(34), "")
                line = fileReader.ReadFields()
                dt.Rows.Add(line)
                'line = fileReader.ReadLine
            End While
            fileReader.Close()
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function
    Public Function importExcelFromCSV(ByVal gv As RadGridView, ByVal filePath As String, ByVal ParamArray fieldNames As String()) As Boolean
        Dim ds As DataTable = ReturnData(filePath)
        gv.DataSource = ds
        gv.AllowColumnReorder = True
        Dim fieldCount As Integer = fieldNames.Length
        Dim strfields As String = ""
        For Each field As String In fieldNames
            strfields = strfields + field + ","
        Next

        If fieldCount <= gv.ColumnCount Then
            Dim i As Integer = 0
            Dim arr As ArrayList = New ArrayList()
            For Each GC As GridViewColumn In gv.Columns
                arr.Add(GC.HeaderText)
            Next
            For Each field As String In fieldNames
                If arr.Contains(field.Trim()) Then

                    For Each GC As GridViewColumn In gv.Columns
                        If GC.HeaderText = field Then
                            gv.Columns.Move(GC.Index, i)
                            Exit For
                        End If
                    Next
                Else
                    common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
                    Return False
                End If
                i = i + 1
            Next
        Else
            common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
            Return False
        End If
        Return True
    End Function
    'Public Function importExcel(ByVal gv As RadGridView, ByVal ParamArray fieldNames As String()) As Boolean
    '    Try

    '        Dim ofd As OpenFileDialog = New OpenFileDialog()
    '        Dim filePath As String
    '        ofd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
    '        If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
    '            filePath = ofd.FileName
    '        Else
    '            Return False
    '        End If
    '        Dim Extension As String = Path.GetExtension(filePath)
    '        Dim conStr As String = ""


    '        'Dim oApp As Excel.Application
    '        'Dim oWB As Excel.Workbook
    '        'oApp = New Excel.Application
    '        'oWB = oApp.Workbooks.Open(filePath)
    '        'MessageBox.Show(oWB.FileFormat.ToString)
    '        Dim rvalue As Boolean = False
    '        If clsCommon.CompairString(Extension, ".csv") = CompairStringResult.Equal Then
    '            rvalue = importExcelFromCSV(gv, filePath, fieldNames)
    '            Return rvalue
    '        End If

    '        Select Case Extension
    '            Case ".xls"
    '                '        'Excel 97-03 
    '                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
    '                Exit Select
    '            Case ".xlsx"
    '                '        'Excel 07  
    '                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    '                Exit Select
    '        End Select
    '        'conStr = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"";"
    '        conStr = [String].Format(conStr, filePath)

    '        Dim connExcel As New OleDbConnection(conStr)
    '        Dim cmdExcel As New OleDbCommand()
    '        Dim oda As New OleDbDataAdapter()
    '        Dim ds As New DataTable()
    '        cmdExcel.Connection = connExcel

    '        'Get the name of First Sheet  
    '        connExcel.Open()
    '        Dim dtExcelSchema As DataTable
    '        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '        Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
    '        connExcel.Close()

    '        'Read Data from First Sheet  
    '        connExcel.Open()
    '        cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
    '        oda.SelectCommand = cmdExcel
    '        oda.Fill(ds)
    '        connExcel.Close()
    '        gv.DataSource = ds.DefaultView
    '        gv.AllowColumnReorder = True
    '        Dim fieldCount As Integer = fieldNames.Length
    '        Dim strfields As String = ""
    '        For Each field As String In fieldNames
    '            strfields = strfields + field + ","
    '        Next

    '        If fieldCount <= gv.ColumnCount Then
    '            Dim i As Integer = 0
    '            Dim arr As ArrayList = New ArrayList()
    '            For Each GC As GridViewColumn In gv.Columns
    '                arr.Add(GC.HeaderText)
    '            Next
    '            For Each field As String In fieldNames
    '                If arr.Contains(field.Trim()) Then

    '                    For Each GC As GridViewColumn In gv.Columns
    '                        If GC.HeaderText = field Then
    '                            gv.Columns.Move(GC.Index, i)
    '                            Exit For
    '                        End If
    '                    Next
    '                Else
    '                    common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
    '                    Return False
    '                End If
    '                i = i + 1
    '            Next
    '        Else
    '            common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
    '            Return False
    '        End If
    '       Return True

    '    Catch ex As Exception
    '        'common.clsCommon.MyMessageBoxShow("No data transfered.", "Import Error", MessageBoxButtons.OK)
    '        clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Function
    Public Function importExcelPivot(ByVal gv As RadGridView, ByVal strPivotCount As Integer, ByVal strPivot As String, ByVal ParamArray fieldNames As String()) As Boolean
        Try
            If Not LoadDocument(gv, "", fieldNames) Then
                Return False
            End If

            'Dim fieldCount As Integer = fieldNames.Length
            'fieldCount = fieldCount + strPivotCount
            'Dim strfields As String = ""
            'For Each field As String In fieldNames
            '    strfields = strfields + field + ","
            'Next
            'strfields = strfields.Substring(0, strfields.Length - 1)
            'strfields = strfields + "," + strPivot

            'If fieldCount <= gv.ColumnCount Then
            '    Dim i As Integer = 0
            '    Dim arr As ArrayList = New ArrayList()
            '    For Each GC As GridViewColumn In gv.Columns
            '        arr.Add(GC.HeaderText)
            '    Next
            '    For Each field As String In fieldNames
            '        If arr.Contains(field.Trim()) Then

            '            For Each GC As GridViewColumn In gv.Columns
            '                If GC.HeaderText = field Then
            '                    gv.Columns.Move(GC.Index, i)
            '                    Exit For
            '                End If
            '            Next
            '        Else
            '            common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
            '            Return False
            '        End If
            '        i = i + 1
            '    Next
            'Else
            '    common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
            '    Return False
            'End If


        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow("No data transfered.", "Import Error", MessageBoxButtons.OK)
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return True
    End Function
    Public Function importExcelWithoutReadColumnName(ByVal gv As RadGridView, ByVal ParamArray fieldNames As String()) As Boolean
        Try

            If Not LoadDocument(gv, "", "") Then
                Return False
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return True
    End Function


    Public Function importExcel(ByVal gv As RadGridView, ByVal ListField As List(Of String)) As Boolean
        Dim TemplatefieldNamesArrString() As String = ListField.ToArray()
        Dim BoolResponse As Boolean = importExcel("", "", gv, TemplatefieldNamesArrString)
        Return BoolResponse
    End Function

    Public Function importExcel(ByVal gv As RadGridView, ByVal ParamArray fieldNames As String()) As Boolean
        Return importExcel("", "", gv, fieldNames)
    End Function
    Public Function importExcel(ByRef FileName As String, ByRef SafeFileName As String, ByVal gv As RadGridView, ByVal ParamArray fieldNames As String()) As Boolean
        Return importExcel(False, FileName, SafeFileName, gv, fieldNames)
    End Function
    Public Function importExcel(ByVal DBFOnly As Boolean, ByRef FileName As String, ByRef SafeFileName As String, ByVal gv As RadGridView, ByVal ParamArray fieldNames As String()) As Boolean
        Try
            If Not LoadDocument(DBFOnly, gv, "", FileName, SafeFileName, fieldNames) Then
                Return False
            End If
            Dim fieldCount As Integer = fieldNames.Length
            Dim strfields As String = ""
            For Each field As String In fieldNames
                strfields = strfields + field + ","
            Next

            If fieldCount <= gv.ColumnCount Then
                Dim i As Integer = 0
                Dim arr As ArrayList = New ArrayList()
                For Each GC As GridViewColumn In gv.Columns
                    arr.Add(GC.HeaderText.Trim().ToUpper())
                Next
                For Each field As String In fieldNames
                    If arr.Contains(field.Trim().ToUpper()) Then
                        For Each GC As GridViewColumn In gv.Columns
                            If GC.HeaderText = field Then
                                gv.Columns.Move(GC.Index, i)
                                Exit For
                            End If
                        Next
                    Else
                        common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format.It should have the columns named - " + strfields)
                        Return False
                    End If
                    i = i + 1
                Next
            Else
                common.clsCommon.MyMessageBoxShow("Excel Sheet is not in expected format. It should have the columns named - " + strfields)
                Return False
            End If


        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow("No data transfered.", "Import Error", MessageBoxButtons.OK)
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return True
    End Function


    '====================Made By Monika For getting runtime excel sheet Columns Name===============02/06/2014===
    Public Function GetExcelColumnsName(ByVal gv1 As RadGridView) As String
        Return GetExcelColumnsName(gv1, "", "")
    End Function
    Public Function GetExcelColumnsName(ByVal gv1 As RadGridView, ByRef FileName As String, ByRef SafeFileName As String) As String
        Dim columnsname As String = ""
        Try
            Dim ofd As OpenFileDialog = New OpenFileDialog()
            Dim filePath As String
            ofd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                filePath = ofd.FileName
            Else
                Return columnsname
            End If
            Dim Extension As String = Path.GetExtension(filePath)
            Dim conStr As String = ""

            FileName = ofd.FileName
            SafeFileName = ofd.SafeFileName

            If clsCommon.CompairString(Extension, ".CSV") = CompairStringResult.Equal Then
                gv1.DataSource = ReturnData(filePath)
            Else
                Select Case Extension
                    Case ".xls"
                        '        'Excel 97-03 
                        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
                        Exit Select
                    Case ".xlsx"
                        '        'Excel 07  
                        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=Excel 12.0 Xml;HDR=Yes;IMEX=1"";"
                        Exit Select
                End Select


                conStr = [String].Format(conStr, filePath)
                Dim connExcel As New OleDbConnection(conStr)
                Dim cmdExcel As New OleDbCommand()
                Dim oda As New OleDbDataAdapter()
                Dim ds As New DataTable()
                cmdExcel.Connection = connExcel
                'Get the name of First Sheet  
                connExcel.Open()
                Dim dtExcelSchema As DataTable
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
                connExcel.Close()
                'Read Data from First Sheet  
                connExcel.Open()
                cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
                oda.SelectCommand = cmdExcel
                oda.Fill(ds)
                connExcel.Close()
                gv1.DataSource = ds.DefaultView
                gv1.AllowColumnReorder = True
            End If
            columnsname = ""
            If gv1.Rows.Count > 0 Then
                For Each GC As GridViewColumn In gv1.Columns
                    columnsname = columnsname + "," + Chr(34) + clsCommon.myCstr(GC.HeaderText) + Chr(34)
                Next
            End If

            Try
                If columnsname.Substring(0, 1) = "," Then
                    columnsname = columnsname.Substring(1, columnsname.Length - 1).Replace("#", ".")
                End If
            Catch exx As Exception
            End Try
            Return columnsname
        Catch ex As Exception
            Return columnsname
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Sub generateReport(ByVal reportName As String, ByVal caption As String, Optional ByVal StrClause As String = vbNullString, Optional ByVal StrClause1 As String = vbNullString, Optional ByVal StrClause2 As String = vbNullString, Optional ByVal StrClause3 As String = vbNullString, Optional ByVal StrClause4 As String = vbNullString, Optional ByVal StrClause5 As String = vbNullString, Optional ByVal StrClause6 As String = vbNullString, Optional ByVal StrClause7 As String = vbNullString, Optional ByVal StrClause8 As String = vbNullString, Optional ByVal strReportTitle As String = vbNullString)
        ''To be Uncomment
        'Dim frm As New FrmReportViewer()
        'frm.Text = caption
        'frm.proShowReport(reportName, StrClause, StrClause1, StrClause2, StrClause3, StrClause4, StrClause5, StrClause6, StrClause7, StrClause8, strReportTitle)
        'frm.ShowDialog()
    End Sub







    Public Function ExporttoExcelNew(ByVal sql As String, ByVal frm As RadForm, Optional ByVal pivotCols As String = "", Optional ByVal whrClaus As String = "", Optional ByVal OrderByClaus As String = "") As Boolean
        Try
            ''************* Filter Block Start
            '============Add By Rohit on June 17,2014 to show column Filter========
            Dim frmFilterCol As New frmFilterColumnsToExport()
            frmFilterCol.qry = sql
            frmFilterCol.pivotCols = pivotCols
            frmFilterCol.whrCls = " Where 2=2 " + whrClaus
            If clsCommon.myLen(OrderByClaus) > 0 Then
                frmFilterCol.orderByClause = " Order by " + OrderByClaus
            End If
            frmFilterCol.ShowDialog()
            If frmFilterCol.isCancel Then
                GoTo a
                'Return False
            End If
            sql = frmFilterCol.qry
            '========================================================================
a:          Dim frmFilter As New frmFilterToExport()
            frmFilter.qry = sql
            frmFilter.whrCls = " Where 2=2 " + whrClaus
            If clsCommon.myLen(OrderByClaus) > 0 Then
                frmFilter.orderByClause = " Order by " + OrderByClaus
            End If
            frmFilter.ShowDialog()
            If frmFilter.isCancel Then
                Return False
            End If
            sql = frmFilter.qry
            ''************* Filter Block End
            Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads" + "\"
            'Dim forlderName As String = "D" + ":\ERPTempFolder" + "\Downloads" + "\"
            Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            If IsExists = False Then
                System.IO.Directory.CreateDirectory(forlderName)
            End If
            'Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
            'Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            'If IsExists = False Then
            '    System.IO.Directory.CreateDirectory(forlderName)
            'End If
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim path As String
            sfd.FileName = frm.Text
            sfd.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
            ' If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = forlderName + frm.Text
            'Else
            '    Return False
            'End If
            If InStr(path, ".xlsx") <> -1 Then
                path = Replace(path, ".xlsx", ".xls")
            End If
            If Not path.Equals(String.Empty) Then
                Dim gv As New UserControls.MyRadGridView
                Try
                    'Dim exporter As New RadGridViewExcelExporter()
                    gv.Name = "gTax"
                    frm.Controls.Add(gv)
                    FillGridView(sql, gv)
                    'gv.MasterGridViewTemplate.AllowAddNewRow = False
                    'gv.MasterGridViewTemplate.AutoGenerateColumns = True
                    'gv.DataSource = Nothing
                    'gv.DataSource = clsDBFuncationality.GetDataTable(sql)
                    If gv.Rows.Count = 0 And frmFilter.IsBlankSheet = False Then
                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                        Return False
                    End If

                    Dim i As Integer = 0
                    '===============Add If Condition For Row Count by Rohit on june 05,2014 If Sheet is Blank then This Loop should not execute===
                    If gv.Rows.Count > 0 Then
                        For i = 0 To gv.ColumnCount - 1
                            Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                            If TypeOf grow.Cells(i).Value Is DateTime Then
                                Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                                datecol.ExcelExportType = DisplayFormatType.ShortDate
                            End If
                        Next i
                    End If
                    '==================================================================
                    'AddHandler exporter.Progress, AddressOf PB.exporter_Progress
                    'ShowPb(gv.Rows.Count)

                    'exporter.Export(gv, path, frm.Text)
                    'HidePb()
                    clsCommon.ProgressBarShow()
                    '================ADD IIF (Blank Sheet) Condition to Save Blank Excel Sheet=========
                    'exportdata(gv, path, path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1), IIf(frmFilter.chkBlankSheet.Checked, True, False)) 'frm.Text)
                    '============================================================
                    Dim exporter1 As New ExportToExcelML(gv)

                    ' AddHandler exporter1.ExcelCellFormatting, AddressOf exporter_ExcelCellFormatting
                    'exporter1.ExportHierarchy = True
                    'exporter1.ExportVisualSettings = True
                    'exporter1.SheetMaxRows = ExcelMaxRows._1048576
                    exporter1.SheetName = frm.Text
                    exporter1.RunExport(path)
                    frm.Controls.Remove(gv)
                    Dim oApp As Excel.Application
                    Dim oWB As Excel.Workbook
                    oApp = New Excel.Application
                    oWB = oApp.Workbooks.Open(path)
                    oApp.DisplayAlerts = False
                    oWB.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal)
                    oWB.Close()
                    oApp.Quit()

                    'My.Computer.FileSystem.RenameFile(Microsoft.VisualBasic.Left(path, Len(path) - 1), path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1))
                    clsCommon.ProgressBarHide()
                    common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)

                    ' Ensure the path ends with ".xls"
                    If Not path.ToLower().EndsWith(".xls") Then
                        path &= ".xls"
                    End If

                    ' Now start the file if it exists
                    If System.IO.File.Exists(path) Then
                        System.Diagnostics.Process.Start(path)
                    Else
                        MessageBox.Show("File not found: " & path)
                    End If

                    'System.Diagnostics.Process.Start(path)
                    'Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
                    'excel.Workbooks.Open(path)
                    'excel.Visible = True


                Catch ex As Exception
                    frm.Controls.Remove(gv)
                    clsCommon.ProgressBarHide()
                    'HidePb()
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
        End Try
        Return True
    End Function
    '======shivani

    Public Function ExporttoExcelForPivot(ByVal sql As String, ByVal frm As RadForm, Optional ByVal pivotCols As String = "", Optional ByVal whrClaus As String = "", Optional ByVal OrderByClaus As String = "") As Boolean
        Try

a:          Dim frmFilter As New frmFilterToExport()
            frmFilter.qry = sql
            frmFilter.whrCls = " Where 2=2 " + whrClaus
            If clsCommon.myLen(OrderByClaus) > 0 Then
                frmFilter.orderByClause = " Order by " + OrderByClaus
            End If
            frmFilter.ShowDialog()
            If frmFilter.isCancel Then
                Return False
            End If
            sql = frmFilter.qry
            ''************* Filter Block End
            Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
            Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            If IsExists = False Then
                System.IO.Directory.CreateDirectory(forlderName)
            End If
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim path As String
            sfd.FileName = frm.Text
            sfd.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = forlderName
            'Else
            '    Return False
            'End If
            If InStr(path, ".xlsx") <> -1 Then
                path = Replace(path, ".xlsx", ".xls")
            End If
            If Not path.Equals(String.Empty) Then
                Dim gv As New UserControls.MyRadGridView
                Try
                    'Dim exporter As New RadGridViewExcelExporter()
                    gv.Name = "gTax"
                    frm.Controls.Add(gv)
                    FillGridView(sql, gv)
                    'gv.MasterGridViewTemplate.AllowAddNewRow = False
                    'gv.MasterGridViewTemplate.AutoGenerateColumns = True
                    'gv.DataSource = Nothing
                    'gv.DataSource = clsDBFuncationality.GetDataTable(sql)
                    If gv.Rows.Count = 0 And frmFilter.IsBlankSheet = False Then
                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                        Return False
                    End If

                    Dim i As Integer = 0
                    '===============Add If Condition For Row Count by Rohit on june 05,2014 If Sheet is Blank then This Loop should not execute===
                    If gv.Rows.Count > 0 Then
                        For i = 0 To gv.ColumnCount - 1
                            Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                            If TypeOf grow.Cells(i).Value Is DateTime Then
                                Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                                datecol.ExcelExportType = DisplayFormatType.ShortDate
                            End If
                        Next i
                    End If
                    '==================================================================
                    'AddHandler exporter.Progress, AddressOf PB.exporter_Progress
                    'ShowPb(gv.Rows.Count)

                    'exporter.Export(gv, path, frm.Text)
                    'HidePb()
                    clsCommon.ProgressBarShow()
                    '================ADD IIF (Blank Sheet) Condition to Save Blank Excel Sheet=========
                    'exportdata(gv, path, path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1), IIf(frmFilter.chkBlankSheet.Checked, True, False)) 'frm.Text)
                    '============================================================
                    Dim exporter1 As New ExportToExcelML(gv)

                    ' AddHandler exporter1.ExcelCellFormatting, AddressOf exporter_ExcelCellFormatting
                    'exporter1.ExportHierarchy = True
                    'exporter1.ExportVisualSettings = True
                    'exporter1.SheetMaxRows = ExcelMaxRows._1048576
                    exporter1.SheetName = frm.Text
                    exporter1.RunExport(path)
                    frm.Controls.Remove(gv)
                    Dim oApp As Excel.Application
                    Dim oWB As Excel.Workbook
                    oApp = New Excel.Application
                    oWB = oApp.Workbooks.Open(path)
                    oApp.DisplayAlerts = False
                    oWB.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal)
                    oWB.Close()
                    oApp.Quit()

                    'My.Computer.FileSystem.RenameFile(Microsoft.VisualBasic.Left(path, Len(path) - 1), path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1))
                    clsCommon.ProgressBarHide()
                    common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
                    System.Diagnostics.Process.Start(path)
                    'Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
                    'excel.Workbooks.Open(path)
                    'excel.Visible = True


                Catch ex As Exception
                    frm.Controls.Remove(gv)
                    clsCommon.ProgressBarHide()
                    'HidePb()
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
        End Try
        Return True
    End Function
    '================

    Public Function ExporttoExcelWithoutFilter(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm, Optional Display_Firstrow As Boolean = False) As Boolean
        Try
            Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
            Dim IsExists As Boolean = System.IO.Directory.Exists(forlderName)
            If IsExists = False Then
                System.IO.Directory.CreateDirectory(forlderName)
            End If
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim filePath As String
            sfd.FileName = frm.Text
            sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 *.xlsx|(*.xlsx);|CSV Files (*.csv) |*.csv"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = forlderName + "\" + sfd.FileName

            'Else
            '    Return False
            'End If

            If Not filePath.Equals(String.Empty) Then
                Dim gv As New UserControls.MyRadGridView
                Try
                    'Dim exporter As New RadGridViewExcelExporter()
                    gv.Name = "gTax"
                    frm.Controls.Add(gv)
                    FillGridView(sql, gv)
                    If gv.Rows.Count = 0 Then
                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
                        Return False
                    End If

                    Dim i As Integer = 0
                    '===============Add If Condition For Row Count by Rohit on june 05,2014 If Sheet is Blank then This Loop should not execute===
                    If gv.Rows.Count > 0 Then
                        For i = 0 To gv.ColumnCount - 1
                            Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
                            If TypeOf grow.Cells(i).Value Is DateTime Then
                                Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
                                datecol.ExcelExportType = DisplayFormatType.ShortDate
                            End If
                        Next i
                    End If
                    '==================================================================
                    'AddHandler exporter.Progress, AddressOf PB.exporter_Progress
                    'ShowPb(gv.Rows.Count)

                    'exporter.Export(gv, path, frm.Text)
                    'HidePb()
                    '            clsCommon.ProgressBarShow()
                    '================ADD IIF (Blank Sheet) Condition to Save Blank Excel Sheet=========
                    Dim ext As String = Path.GetExtension(filePath)
                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        exportdataInCSV(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), False) 'frm.Text)
                    Else
                        'sanjay
                        exportdata(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), False, , Display_Firstrow, False, False, True) 'frm.Text)
                        'sanjay
                    End If


                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        GoTo xxx
                    End If
                    Dim oApp As Excel.Application
                    Dim oWB As Excel.Workbook
                    oApp = New Excel.Application
                    oWB = oApp.Workbooks.Open(filePath)
                    oApp.DisplayAlerts = False

                    oWB.SaveAs(filePath)
                    oWB.Close()
                    oApp.Quit()
xxx:
                    'My.Computer.FileSystem.RenameFile(Microsoft.VisualBasic.Left(path, Len(path) - 1), path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1))

                    If clsCommon.CompairString(ext, ".csv") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
                        System.Diagnostics.Process.Start(filePath)
                    End If
                    Return True
                Catch ex As Exception
                    frm.Controls.Remove(gv)

                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
        End Try
        Return True
    End Function
    Public Function LoadDocument(ByVal gv As RadGridView, sheetName As String, ByVal ParamArray fieldNames As String()) As Boolean
        Return LoadDocument(gv, sheetName, "", "", fieldNames)
    End Function
    Public Function LoadDocument(ByVal gv As RadGridView, sheetName As String, ByRef FileName As String, ByRef SafeFileName As String, ByVal ParamArray fieldNames As String()) As Boolean
        Return LoadDocument(False, gv, sheetName, FileName, SafeFileName, fieldNames)
    End Function
    Public Function LoadDocument(ByVal DBFOnly As Boolean, ByVal gv As RadGridView, sheetName As String, ByRef FileName As String, ByRef SafeFileName As String, ByVal ParamArray fieldNames As String()) As Boolean
        Dim workbook As Excel.Workbook = Nothing
        Dim rvalue As Boolean = False
        Dim ofd As OpenFileDialog = New OpenFileDialog()
        Dim filePath As String
        If clsCommon.myLen(sheetName) <= 0 Then
            sheetName = "Sheet1"
        End If
        FileName = ""
        SafeFileName = ""
        If DBFOnly Then
            ofd.Filter = "DBF Files (*.DBF) |*.DBF"
        Else
            ofd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
        End If
        If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = ofd.FileName
            FileName = ofd.FileName
            SafeFileName = ofd.SafeFileName
        Else
            Return False
        End If
        Dim Extension As String = Path.GetExtension(filePath)
        Dim conStr As String = ""
        Dim selectedFormat As String = Extension
        Dim dt As DataTable = New DataTable()
        Dim colCount As Integer = 0
        clsCommon.ProgressBarPercentShow()
        Try
            Dim oApp As Excel.Application
            oApp = New Excel.Application
            oApp.Visible = False
            oApp.DisplayAlerts = False
            workbook = oApp.Workbooks.Open(filePath)
            Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = workbook.Worksheets(1)

            Dim r As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim array(,) As Object = r.Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault)
            Dim bound0 As Integer = array.GetUpperBound(0)
            Dim bound1 As Integer = array.GetUpperBound(1)
            For i As Integer = 1 To bound1
                clsCommon.ProgressBarPercentUpdate(((i) * 100 / bound1), "Getting Field List " & (i) & " Of Total " & bound1 & " From Excel Sheet")
                dt.Columns.Add(clsCommon.myCstr(array(1, i)).Trim(), "".GetType())
                dt.Columns(clsCommon.myCstr(array(1, i)).Trim()).Caption = array(1, i)
            Next

            For j As Integer = 2 To bound0
                clsCommon.ProgressBarPercentUpdate(((j) * 100 / bound0), "Getting Record List " & (j) & " Of Total " & bound0 & " From Excel Sheet")
                Dim dr As DataRow = dt.NewRow()
                For x As Integer = 1 To bound1
                    dr(clsCommon.myCstr(array(1, x)).Trim()) = array(j, x)
                Next
                dt.Rows.Add(dr)
            Next
            oApp.Workbooks.Close()
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                rvalue = False
            Else
                gv.DataSource = dt
                rvalue = True
            End If
            clsCommon.ProgressBarPercentHide()
        Catch ex As IOException
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try
        Return rvalue
    End Function
    Public Function BulkExport(ByVal ReportName As String, ByVal _Query As String, ByVal OrderByClause As String, ByVal File_Type As String, Optional ByVal CTE_Separater As String = "") As String
        'Ticket No-TEC/01/07/19-000576 ,sanjay ,Create and copy file from remote server
        Dim FinalQuery As String = String.Empty
        Dim OrderBy As String = String.Empty
        Dim Server As String = objCommonVar.Database_Server
        Dim subPath As String = ""
        Dim ReportPath As String = ""
        Dim NetworkSubPath As String = ""
        Dim forlderName As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"

        If clsCommon.CompairString(Server.Substring(0, IIf(Server.Contains("\"), Server.IndexOf("\"), Server.Length)), Environment.MachineName, False) <> CompairStringResult.Equal Then
            subPath = "\\" + Server.Substring(0, IIf(Server.Contains("\"), Server.IndexOf("\"), Server.Length)) + "\ERPTempFolder"
            ReportPath = "\\" + Server.Substring(0, IIf(Server.Contains("\"), Server.IndexOf("\"), Server.Length)) + "\ERPTempFolder"
        Else
            subPath = forlderName
            ReportPath = forlderName
        End If
        NetworkSubPath = forlderName
        'Dim subPath As String = "E:\\DataBase"
        'Dim ReportPath As String = "E:\\DataBase"
        'Dim NetworkSubPath As String = "DataBase"
        ''E:\DataBase
        Dim ReportPathLog As String = String.Empty
        ReportName = ReportName.Replace(" ", "_")
        ReportName = ReportName.Replace(",", "")
        ReportName = ReportName.Replace(".", "")
        If clsCommon.myLen(File_Type) <= 0 Then
            File_Type = "csv"
        ElseIf Not (clsCommon.CompairString(File_Type, "csv") = CompairStringResult.Equal OrElse clsCommon.CompairString(File_Type, "xls") = CompairStringResult.Equal) Then
            Throw New Exception("File format must be xls or csv")
        End If
        Try
            '' find order by clause 
            'clsCommon.ProgressBarShow()
            'Dim arr() As String
            If clsCommon.myLen(OrderByClause) > 0 Then
                If Not OrderByClause.ToLower.Contains("order by") Then
                    Throw New Exception("order by clause must start from order by")
                End If
                If _Query.ToLower.Contains(OrderByClause.ToLower) Then
                    _Query = _Query.Replace(OrderByClause, "")
                    'If arr.Length > 0 Then
                    '    _Query = arr(0)
                    'End If
                End If
                If _Query.ToLower.Contains("order by".ToLower) Then
                    Dim subQery As String = _Query
                    Dim indx As Integer = 0
                    Dim FinalLength As Integer = 0

                    'Dim Indx As Integer = _Query.ToLower.LastIndexOf("order by")

                    ''check if partition by cntains then first resolve this
                    If subQery.ToLower.Contains("partition by") Then
                        indx = subQery.ToLower.LastIndexOf("partition by")
                        If indx > 0 Then
                            FinalLength += indx
                            If clsCommon.myLen(subQery) > indx Then
                                subQery = subQery.Substring(indx, clsCommon.myLen(subQery) - (indx + 1))
                            Else
                                subQery = subQery.Substring(indx, clsCommon.myLen(subQery) - indx)
                            End If
                        End If
                        indx = 0

                        ''=======get first index of order by
                        indx = subQery.ToLower.IndexOf("order by")
                        If indx > 0 Then
                            FinalLength += indx
                            If clsCommon.myLen(subQery) > indx Then
                                subQery = subQery.Substring(indx, clsCommon.myLen(subQery) - (indx + 1))
                            Else
                                subQery = subQery.Substring(indx, clsCommon.myLen(subQery) - indx)
                            End If
                        End If
                        indx = 0

                        ''============now check for last index of order by, if exist
                        indx = subQery.ToLower.LastIndexOf("order by")

                        If indx > 0 Then
                            FinalLength += indx
                            _Query = _Query.ToLower.Substring(0, FinalLength)
                        End If
                    Else
                        indx = subQery.ToLower.LastIndexOf("order by")

                        If indx >= 0 Then
                            _Query = _Query.ToLower.Substring(0, indx)
                        End If
                    End If
                End If
            Else
                Dim indx As Integer = _Query.ToLower.LastIndexOf("order by")
                If indx >= 0 Then
                    _Query = _Query.ToLower.Substring(0, indx)
                End If
                'Dim strTemp As String = ""
                'For intChar As Integer = _Query.Length - 1 To 0 Step -1
                '    If Not strTemp.Contains("from") Then
                '        strTemp = strTemp & _Query.Chars(intChar)
                '    Else
                '        If Not strTemp.Trim.Contains("order by") Then
                '            Throw New Exception("Query must not contain order by ")
                '        End If
                '    End If
                'Next
            End If


            Dim qry As String = "SELECT CONVERT(INT, ISNULL(value, value_in_use)) AS config_value FROM sys.configurations WHERE  name = 'xp_cmdshell' "
            Dim Check As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
            If Check = 0 Then
                qry = " EXEC sp_configure 'show advanced options', 1;" &
                      " RECONFIGURE " &
                      " EXEC sp_configure 'xp_cmdshell', 1 ;" &
                      " RECONFIGURE "
                clsDBFuncationality.ExecuteNonQuery(qry)
            End If
            Dim Col As String = "select "
            Dim ColNew As String = "select "
            Dim dt As New DataTable
            Dim colNo As Integer = 0
            If clsCommon.myLen(CTE_Separater) > 0 Then
                qry = " begin tran" &
                  " SET FMTONLY ON; " &
                  " " & _Query & " " &
                  " SET FMTONLY OFF " &
                  " commit tran"
            Else
                qry = " begin tran" &
                      " SET FMTONLY ON " &
                      " " & _Query & " " &
                      " SET FMTONLY OFF " &
                      " commit tran"
            End If

            dt = clsDBFuncationality.GetDataTable(qry)
            clsCommon.ProgressBarPercentUpdate(30, "Report Header Constructed")


            For Each dc As DataColumn In dt.Columns
                If colNo = 0 Then
                    Col = Col & "" & "'" & dc.ColumnName & "'as [" & dc.ColumnName & "]"
                    If clsCommon.CompairString(File_Type, "csv") = CompairStringResult.Equal Then
                        ColNew = ColNew & " replace(cast([" & dc.ColumnName & "] as varchar(max)),',',' ') "
                    Else
                        ColNew = ColNew & " coalesce(cast([" & dc.ColumnName & "] as varchar(max)),'') "
                    End If

                Else
                    Col = Col & "," & "'" & dc.ColumnName & "'as [" & dc.ColumnName & "]"
                    If clsCommon.CompairString(File_Type, "csv") = CompairStringResult.Equal Then
                        ColNew = ColNew & "," & " replace(cast([" & dc.ColumnName & "]  as varchar(max)),',',' ') "
                    Else
                        ColNew = ColNew & "," & " coalesce(cast([" & dc.ColumnName & "]  as varchar(max)),'') "
                    End If

                End If
                colNo = colNo + 1
            Next

            '' get path of the exported file

            '' your code goes here
            Dim IsExists As Boolean = System.IO.Directory.Exists(subPath)
            If IsExists Then
                System.IO.Directory.CreateDirectory(subPath)
            End If
            Dim Networkpath As String = ""
            Dim currTime As DateTime = DateTime.Now
            ReportPath = ReportPath & "\" & ReportName & clsCommon.GetPrintDate(currTime, "yyyyMMddhhmmss")
            Networkpath = NetworkSubPath & "\" & ReportName & clsCommon.GetPrintDate(currTime, "yyyyMMddhhmmss")
            subPath = ReportPath
            ReportPath = subPath & "." & File_Type
            ReportPathLog = subPath & ".txt"
            If System.IO.File.Exists(ReportPath) = True Then
                System.IO.File.Delete(ReportPath)
            End If
            '' export data
            If clsCommon.myLen(CTE_Separater) > 0 Then
                Dim InnerQry As String = ""
                Dim cteQry As String = ""
                Dim CteIndex As Integer = _Query.ToLower.LastIndexOf(CTE_Separater.ToLower)
                If CteIndex >= 0 Then
                    InnerQry = _Query.Substring(CteIndex, (_Query.Length - CteIndex))
                    cteQry = _Query.Substring(0, CteIndex)
                Else
                    Throw New Exception("CTE separater passed does not exists in Query.")
                End If
                FinalQuery = cteQry & "" & Environment.NewLine & Col & " union all " & ColNew & " from (" & InnerQry & ") Final"
            Else
                FinalQuery = Col & "union all " & ColNew & " from (" & _Query & ") Final"
            End If

            qry = String.Empty
            Dim Rpt_view_Name As String = "view_temp_" & ReportName
            qry = String.Empty
            qry = "SELECT count(*) FROM sys.views where name='" & Rpt_view_Name & "'"
            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) <= 0 Then
                qry = "create view " & Rpt_view_Name & " as " & FinalQuery
            Else
                qry = "alter view " & Rpt_view_Name & " as " & FinalQuery
            End If
            clsDBFuncationality.ExecuteNonQuery(qry)
            clsCommon.ProgressBarPercentUpdate(40, "Report View Created")
            qry = String.Empty
            'Dim Server As String = objCommonVar.Database_Server ''clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT SERVERPROPERTY('MachineName')"))            
            'Dim StrUserNameAndPassword As String = clsCommon.DecryptString(objCommonVar.Database_Server_Password)
            qry = "select * from " & objCommonVar.CurrDatabase & ".dbo." & Rpt_view_Name & ""
            If clsCommon.CompairString(File_Type, "xls") = CompairStringResult.Equal Then
                FinalQuery = "exec master..xp_cmdshell'bcp """ & qry & """ queryout " & ReportPath & " -o " & ReportPathLog & " -c -C RAW, -T -S " & Server & "'"
            Else
                FinalQuery = "exec master..xp_cmdshell'bcp """ & qry & """ queryout " & ReportPath & " -o " & ReportPathLog & " -c -t, -T -S " & Server & "'"
            End If

            clsDBFuncationality.ExecuteNonQuery(FinalQuery)
            '' check for operation is performed on database server or local network system

            If clsCommon.CompairString(Server.Substring(0, IIf(Server.Contains("\"), Server.IndexOf("\"), Server.Length)), Environment.MachineName, False) <> CompairStringResult.Equal Then
                'Dim SrcPath As String = "\\" & Server & "\" & Networkpath & "." & File_Type
                'System.IO.File.Copy(SrcPath, ReportPath, True)
                System.Threading.Thread.Sleep(10000)
                Dim DestinationReportPath = Networkpath & "." & File_Type
                System.IO.File.Copy(ReportPath, DestinationReportPath, True)
                ReportPath = DestinationReportPath
            End If
            clsCommon.ProgressBarPercentUpdate(95, "Data transfered to file.")
            '' open excel file
            Process.Start(ReportPath)
            clsCommon.ProgressBarPercentUpdate(100, "File Opened.")
            ''richa BHA/05/06/19-000900 comment clsCommon.ProgressBarHide()
        Catch ex As Exception
            'clsCommon.ProgressBarHide()
            Throw New Exception(ex.Message)
        Finally
            'clsCommon.ProgressBarHide()
        End Try
        Return ReportPath
    End Function


    Public Function GetExcelData(ByVal filePath As String, sheetName As String) As DataTable
        Dim dt As DataTable = New DataTable()
        Try
            If clsCommon.myLen(sheetName) <= 0 Then
                sheetName = "Sheet1"
            End If
            Dim oApp As Excel.Application
            oApp = New Excel.Application
            oApp.Visible = False
            oApp.DisplayAlerts = False
            Dim workbook As Excel.Workbook = oApp.Workbooks.Open(filePath)
            Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = workbook.Worksheets(1)
            Dim r As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim array(,) As Object = r.Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault)
            Dim bound0 As Integer = array.GetUpperBound(0)
            Dim bound1 As Integer = array.GetUpperBound(1)
            For i As Integer = 1 To bound1
                dt.Columns.Add(array(1, i), "".GetType())
                dt.Columns(array(1, i)).Caption = array(1, i)
            Next
            For j As Integer = 2 To bound0
                Dim dr As DataRow = dt.NewRow()
                For x As Integer = 1 To bound1
                    dr(array(1, x)) = array(j, x)
                Next
                dt.Rows.Add(dr)
            Next
            oApp.Workbooks.Close()
        Catch ex As IOException
            Throw New Exception(ex.Message)
        End Try
        Return dt
    End Function
    Public Function ExportCSV(ByVal sender As RadGridView, Optional ByVal FromRowNo As Integer = 0, Optional ByVal ToRowNo As Integer = 0, Optional ByVal AddHeader As Boolean = False) As String()
        Dim ItemArray As New List(Of String)
        Dim OpenInExcel As Boolean = True

        If FromRowNo <= 0 Then
            FromRowNo = 0
        End If
        If ToRowNo <= 0 Then
            ToRowNo = sender.Rows.Count
        End If
        If AddHeader Then
            ItemArray.Add(
                String.Join(",",
                (
                    From T In sender.Columns.Cast(Of GridViewColumn)()
                    Select T.Name
                ).ToArray)
            )
        End If
        ''And Not DirectCast(row, GridViewRowInfo).IsCurrent
        If OpenInExcel = True Then
            ItemArray.AddRange(
            (
                From row In sender.Rows
                Where (row.Index >= FromRowNo And row.Index <= ToRowNo)
                Let RowItem = String.Join(",",
                Array.ConvertAll(
                DirectCast(row, GridViewRowInfo).Cells.Cast(Of GridViewCellInfo).ToArray,
                Function(c As GridViewCellInfo)
                    Return If(c.Value Is Nothing, "", If(c.Value.ToString.StartsWith("0"), "=""" & c.Value.ToString & """", c.Value.ToString).Replace(",", " "))
                End Function))
                Select RowItem
            ).ToList
        )
            ''And Not DirectCast(row, GridViewRowInfo).IsCurrent
        Else
            ItemArray.AddRange(
            (
                From row In sender.Rows
                Where (row.Index >= FromRowNo And row.Index <= ToRowNo)
                Let RowItem = String.Join(",",
                Array.ConvertAll(
                DirectCast(row, GridViewRowInfo).Cells.Cast(Of GridViewCellInfo).ToArray,
                Function(c As GridViewCellInfo)
                    Return If(c.Value Is Nothing, "", c.Value.ToString.Replace(",", " "))
                End Function))
                Select RowItem
            ).ToList
        )
        End If
        Return ItemArray.ToArray

    End Function
    Public Function ExportCSVMultipleFile(ByVal gv As RadGridView, ByVal FilePath As String, Optional ByVal AddHeader As Boolean = False) As Integer
        clsCommon.ProgressBarShow()
        Dim FileCount As Integer = 1
        Try
            '' get max row for csv export from fix parameter
            If gv.Rows.Count <= 0 Then
                Throw New Exception("No record in grid to export")
            End If
            Dim IndexList As List(Of String) = New List(Of String)
            For i As Integer = 0 To gv.Columns.Count - 1
                If Not gv.Columns(i).IsVisible Then
                    IndexList.Add(gv.Columns(i).Name)
                End If
            Next
            For i As Integer = 0 To IndexList.Count - 1
                gv.Columns.Remove(IndexList.Item(i).ToString())
            Next
            Dim MaxRowExport As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsInCSVExport, clsFixedParameterCode.MaxRowsInCSVExport, Nothing))
            If MaxRowExport <= 0 Then
                MaxRowExport = 200000
            End If

            If gv.Rows.Count < MaxRowExport Then
                IO.File.WriteAllLines(FilePath, transportSql.ExportCSV(gv, 0, gv.Rows.Count, AddHeader))
                FileCount = 1
            Else
                Dim FilePathRoot As String = System.IO.Path.GetDirectoryName(FilePath)
                Dim fileName As String = System.IO.Path.GetFileName(FilePath)
                Dim fileExtn As String = System.IO.Path.GetExtension(FilePath)
                fileName = fileName.Replace(fileExtn, "")
                FilePath = FilePathRoot & "\" & fileName
                If System.IO.Directory.Exists(FilePath) = False Then
                    System.IO.Directory.CreateDirectory(FilePath)
                End If

                'Dim tblArr() As DataTable
                Dim tableCount = Math.Ceiling(gv.Rows.Count / MaxRowExport)
                Dim Divisor = gv.Rows.Count / tableCount
                '' for data boun through data reader
                Dim fromCount As Integer = 1
                Dim ToCount As Integer = gv.Rows.Count
                FileCount = tableCount
                For intLoop As Integer = 0 To tableCount - 1
                    Dim file As String
                    If intLoop = tableCount - 1 Then
                        ToCount = gv.Rows.Count
                        file = FilePath & "\" & fileName & intLoop + 1 & fileExtn
                        IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        fromCount = fromCount + (MaxRowExport)
                    Else
                        ToCount = (fromCount + 0) + (MaxRowExport - 1)
                        file = FilePath & "\" & fileName & intLoop + 1 & fileExtn
                        IO.File.WriteAllLines(file, transportSql.ExportCSV(gv, (fromCount - 1), (ToCount - 1), AddHeader))
                        fromCount = (fromCount) + (MaxRowExport)
                    End If
                Next



            End If
        Catch ex As Exception
            clsCommon.ProgressBarHide()
            Throw New Exception(ex.Message)
        End Try
        clsCommon.ProgressBarHide()
        Return FileCount

    End Function



    Public Function applyExpImpTemplate(ByVal gv As RadGridView, ByVal program_Code As String) As Boolean
        Dim qry As String = ""
        Dim whrCls As String = ""
        Dim Export_code As String = ""
        Try
            qry = "select Export_Code as Code,Template_Name as Name from TSPL_TEMPLATE_EXPIMP_HEAD "
            whrCls = " Program_Code='" & program_Code & "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry & " where " & whrCls)
            If dt.Rows.Count > 0 Then
                '' show finder in all cases because user can ignore template to export all columns
                Export_code = clsTemplateExpImp.GetFinder(" TSPL_TEMPLATE_EXPIMP_HEAD.Program_Code='" + program_Code + "'", "", True)
            End If

            If clsCommon.myLen(Export_code) > 0 Then
                Dim obj As New clsTemplateExpImp
                obj = clsTemplateExpImp.GetData(Export_code, program_Code, Nothing)
                If Not obj Is Nothing AndAlso obj.Arr.Count > 0 Then
                    '' invisible all columns of grid
                    For Each gcol As GridViewColumn In gv.Columns
                        gcol.IsVisible = False
                    Next
                    For Each objtr As clsTemplateExpImpDetail In obj.Arr
                        If gv.Columns.IndexOf(objtr.Column_Name) >= 0 Then
                            gv.Columns(objtr.Column_Name).IsVisible = True
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function



    Public Function applyExportTemplate(ByVal gv As RadGridView, ByVal program_Code As String, Optional ByVal Report_type As String = "") As Boolean
        '' added by Panch Raj on 05-05-2018 against tickt: KDI/02/05/18-000288
        '' Work: if user created an  export column template then it must be exported with columns defined in template.
        Dim qry As String = ""
        Dim whrCls As String = ""
        Dim Export_code As String = ""
        Try
            '' query change by Panch Raj against Ticket No: KDI/11/05/18-000311
            qry = "select Export_Code as Code,Template_Name as Name from TSPL_EXPORT_TEMPLATE_HEAD "
            whrCls = " Program_Code='" & program_Code & "' and Created_By='" & objCommonVar.CurrentUserCode & "'  " & If(clsCommon.myLen(Report_type) > 0, " and Report_Type='" & Report_type & "'", "") & ""
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry & " where " & whrCls)
            If dt.Rows.Count > 0 Then
                '    Export_code = clsExportTemplate.GetFinder(" TSPL_EXPORT_TEMPLATE_HEAD.Program_Code='" + program_Code + "' " & If(clsCommon.myLen(Report_type) > 0, " and TSPL_EXPORT_TEMPLATE_HEAD.Report_Type='" & Report_type & "' and TSPL_EXPORT_TEMPLATE_HEAD.Created_By='" & objCommonVar.CurrentUserCode & "' ", " and TSPL_EXPORT_TEMPLATE_HEAD.Created_By='" & objCommonVar.CurrentUserCode & "'") & "", "", False)
                'ElseIf dt.Rows.Count = 1 Then
                '    Export_code = dt.Rows(0).Item("Code")
                '' show finder in all cases because user can ignore template to export all columns
                Export_code = clsExportTemplate.GetFinder(" TSPL_EXPORT_TEMPLATE_HEAD.Program_Code='" + program_Code + "' " & If(clsCommon.myLen(Report_type) > 0, " and TSPL_EXPORT_TEMPLATE_HEAD.Report_Type='" & Report_type & "' and TSPL_EXPORT_TEMPLATE_HEAD.Created_By='" & objCommonVar.CurrentUserCode & "' ", " and TSPL_EXPORT_TEMPLATE_HEAD.Created_By='" & objCommonVar.CurrentUserCode & "'") & "", "", True)
            End If

            If clsCommon.myLen(Export_code) > 0 Then
                Dim obj As New clsExportTemplate
                obj = clsExportTemplate.GetData(Export_code, program_Code, Report_type, Nothing)
                If Not obj Is Nothing AndAlso obj.Arr.Count > 0 Then
                    '' invisible all columns of grid
                    For Each gcol As GridViewColumn In gv.Columns
                        gcol.IsVisible = False
                    Next
                    For Each objtr As clsExportTemplateDetail In obj.Arr
                        If gv.Columns.IndexOf(objtr.Column_Name) >= 0 Then
                            gv.Columns(objtr.Column_Name).IsVisible = True
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    'sanjay
    Public Function applyTemplate(ByVal gv As RadGridView, ByVal ReportId As String) As Boolean
        Try
            Dim qry As String = ""
            Dim whrCls As String = ""
            Dim TemplateName As String = ""

            qry = "select TemplateName as Name from TSPL_MANAGE_TEMPLATE "
            whrCls = " ReportId='" & ReportId & "' and UserID='" & objCommonVar.CurrentUserCode & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry & " where " & whrCls)
            If dt.Rows.Count > 0 Then

                TemplateName = clsManageTemplate.GetFinder(" TSPL_MANAGE_TEMPLATE.ReportId='" + ReportId + "' and TSPL_MANAGE_TEMPLATE.UserId='" & objCommonVar.CurrentUserCode & "' ", "", True)
            End If


            If clsCommon.myLen(TemplateName) > 0 And clsCommon.myLen(ReportId) > 0 Then
                Dim obj As New clsManageTemplate()
                obj = clsManageTemplate.GetData(ReportId, TemplateName)
                Dim strGridLayout As String = obj.GridLayout
                Dim byteArray As Byte() = System.Text.Encoding.ASCII.GetBytes(strGridLayout)
                Dim GridLayout As Stream = Nothing
                GridLayout = New MemoryStream(byteArray)
                'gv.LoadLayout(obj.GridLayout)
                'obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)

                If Not obj Is Nothing AndAlso obj.GridColumns >= gv.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gv.Columns.Count - 1 Step ii + 1
                        gv.Columns(ii).IsVisible = False
                        gv.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv.LoadLayout(GridLayout)
                    GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If

            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
        Return True
    End Function
End Module



