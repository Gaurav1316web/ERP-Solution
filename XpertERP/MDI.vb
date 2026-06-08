'Imports common
Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports CgtFpAccessCSD200Dotnet
Imports common
Imports XpertErpBullManagement
Imports XpertERPCustomerComplaint
Imports XpertERPElectrical
Imports XpertERPEngineeringAndPlantManagement
Imports XpertERPFixedAssets
Imports XpertERPHRandPayroll
Imports XpertERPJobWorkOutward
Imports XpertERPMIS
Imports XpertERPParavetServices
Imports XpertERPProcessProduction
Imports XpertERPProjectManagement
Imports XpertERPPurchase
Imports XpertERPReco
Imports XpertERPRiceProduction
Imports XpertERPService
Imports XpertERPSheed
Imports XpertERPTDS

Public Class MDI
#Region "Varaibles"
    Dim DeleteTempData As Boolean = False
    Dim DeleteTempDataFromTime As String
    'Dim DeleteTempDataToTime As DateTime
    Dim ImportExportDrive As String

    Public Shared blnShowAllMenu As Boolean = False
    Private isUtilityAdded As Boolean = False
    Private ArrImageList As New Dictionary(Of String, Integer)
    Private ArrBold As New List(Of String)
    Public arrExcluded As New List(Of String)
    Public settDCS As Boolean
    'Public Shared IsMailSend As String = "NO"
    Public Shared IsLoc_Third_Party As String = "NO"
    Public Shared IsLoaction_NLevel As String = "NO"
    Public Shared EnableScreenSelection As Boolean = False
    Private SettSameUserCanNotloginmultipletimes As Boolean = False
    Public PasswordRules As Boolean = False
    Public Shared IsVendor_NLevel As String = "NO"
    Public Shared IsCustomer_NLevel As String = "NO"
    Dim OldThemeName As String = ""
    Public frm
    Public Shared isAutoClosing As Boolean = False
    '    Public SystemIdleTimer1 As New SystemIdleTimer
    Public isIdle As Integer = 0
    Public IdleTimeinSeconds As Integer = 0
    Dim Qry As String = ""
    Dim dt As DataTable
    Dim ArrItem As List(Of clsItemMaster)
    Dim IsDBRestored As Boolean = False
    Public isApplicationRun As Boolean = False
    Public isLoadAppIntegrator As Boolean = False
    Public isLoadBulkPurchaseUploader As Boolean = False
    Public isLoadBankUpdateUploader As Boolean = False
    Public IsLoadMccBugReports As Boolean = False
    Public IsOriginalName As Boolean = False
    '' For Multithreading
    Dim th As Threading.Thread = Nothing
    Dim th1 As Threading.Thread = Nothing
    ''
    Dim OLDshortDate As String = ""
    Dim SettingHighSecurityOnWeighingIntegratedScreen As Boolean = False
    Dim SingleUserParticularDairyBookingEdit As Boolean = False
    Dim ShowNotificationInMDI As Integer = 0
    Dim ShowNotificationWithoutSMSAPP As Boolean = False
    Dim SetNotificationRefreshTimeInMinutes As Integer = 0
#End Region

#Region "RadButtons"
    Dim arralert As New Dictionary(Of String, RadDesktopAlert)()
    Dim radbuttonelement As New RadButtonElement("Snooze")
    Dim radbuttonDontShow As New RadButtonElement("Don't Show Again")
    Dim radbuttonelementA As RadButtonElement = New RadButtonElement("Snooze")
    Dim radbuttonDontShowA As RadButtonElement = New RadButtonElement("Don't Show Again")
#End Region


    Private Sub MDI_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            clsDBFuncationality._LastActiveTime = DateTime.Now()
            GetType(SplitContainer).GetMethod("SetStyle", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).Invoke(SplitContainer1.Panel1, New Object() {ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True})
            GetType(SplitContainer).GetMethod("SetStyle", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).Invoke(SplitContainer1.Panel2, New Object() {ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True})
            GetType(SplitContainer).GetMethod("SetStyle", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).Invoke(SplitContainer3.Panel1, New Object() {ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True})
            GetType(SplitContainer).GetMethod("SetStyle", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).Invoke(SplitContainer3.Panel2, New Object() {ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True})
        Catch ex As Exception
        End Try
    End Sub

    Private Sub MDI_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim strQ As String = "update tspl_user_master set IP_Address=NULL,Login_Status=0 where user_code='" + objCommonVar.CurrentUserCode + "'"
        clsDBFuncationality.ExecuteNonQuery(strQ)

        If Not IsDBRestored Then
            strQ = "Update TSPL_UserLogin_Info set Logout_DateTime=' " + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") + "'  where Login_Code ='" + objCommonVar.CurrentLoginID + "'"
            connectSql.RunSql(strQ)
        End If
        If clsCommon.CompairString(OLDshortDate, "MM/dd/yyyy") = CompairStringResult.Equal Then
            Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", OLDshortDate)
        End If
        GC.Collect()
    End Sub

    Public Function IsProcessRunning(name As String) As Boolean

        'here we're going to get a list of all running processes on  

        'the computer  

        For Each clsProcess As Process In Process.GetProcesses()

            If clsProcess.ProcessName.StartsWith(name) Then

                'process found so it's running so return true  

                Return True

            End If

        Next

        'process not found, return false  

        Return False

    End Function

    Private Sub MDI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If CheckConfigFile() Then
            RadButton18.Text = "Dashboard"
            LoadTheme()
            LoadWelcomeScreen()
            Dim coll As New Dictionary(Of String, String)()
            coll.Add("Program_Code", "VARCHAR(12) NOT NULL ")
            coll.Add("Created_By", "varchar(12) null")
            coll.Add("Created_Date", "Datetime not null")
            clsCommonFunctionality.CreateOrAlterTable("TSPL_PROGRAM_MASTER_COUNTER", coll)

            If clsCommon.CompairString(clsFixedParameter.GetData(clsFixedParameterType.AutoBackUp, clsFixedParameterCode.AutoBackUp, Nothing), "0") = CompairStringResult.Equal Then
                Timer2.Enabled = False
            End If
            ReminderTimer.Enabled = False
            blnShowAllMenu = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ShowAllMenu, clsFixedParameterCode.ShowAllMenu, Nothing)) = "1", True, False))
            IsLoaction_NLevel = clsCommon.myCstr(IIf(clsFixedParameter.GetData(clsFixedParameterType.NLevelAtLocation, clsFixedParameterCode.NLevelAtLocation, Nothing) = "1", "YES", "NO"))
            IsVendor_NLevel = clsCommon.myCstr(IIf(clsFixedParameter.GetData(clsFixedParameterType.NLevelAtVendor, clsFixedParameterCode.NLevelAtVendor, Nothing) = "1", "YES", "NO"))
            IsCustomer_NLevel = clsCommon.myCstr(IIf(clsFixedParameter.GetData(clsFixedParameterType.NLevelAtCustomer, clsFixedParameterCode.NLevelAtCustomer, Nothing) = "1", "YES", "NO"))
            '' Changes by Parteek for Screen Level Rights Ticket No : TEC/16/03/18-000101 on 01/05/2018
            EnableScreenSelection = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EnableScreenSelection, clsFixedParameterCode.EnableScreenSelection, Nothing)) = "1", True, False))
            ''End
            'TakeBackup()
            'If Not clsCommon.CompairString(clsFixedParameter.GetData(clsFixedParameterType.UseControlMForHelp, clsFixedParameterCode.UseControlMForHelp, Nothing), "0") = CompairStringResult.Equal Then
            '    Label4.Text = "Press Ctrl+M For Help"
            'End If
            DeleteTempData = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DeleteTempData, clsFixedParameterCode.DeleteTempData, Nothing)) = 1)

            SettSameUserCanNotloginmultipletimes = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SameuserCanNotloginmultipletimes, clsFixedParameterCode.SameuserCanNotloginmultipletimes, Nothing)) = 1)
            ShowNotificationInMDI = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowNotificationInMDI, clsFixedParameterCode.ShowNotificationInMDI, Nothing))
            ShowNotificationWithoutSMSAPP = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ShowNotificationWithoutSMSAPP, clsFixedParameterCode.ShowNotificationWithoutSMSAPP, Nothing)) = "1", True, False))
            SetNotificationRefreshTimeInMinutes = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SetNotificationRefreshTimeInMinutes, clsFixedParameterCode.SetNotificationRefreshTimeInMinutes, Nothing))
            If clsCommon.myCDate(clsCommon.GETSERVERDATE()) > clsCommon.myCDate("31/01/2019") Then
                Panel2.Visible = False
            Else
                Label3.Location = New Point(Panel2.ClientSize.Width,
                   Panel2.ClientSize.Height / 2 - (Label3.Height / 2))
                Timer4.Start()
            End If

            Timer4.Start()


            'If ShowNotificationWithoutSMSAPP = True Then
            '    ShowForm(clsUserMgtCode.frmPromptMsgNotification, "Notification", True, "", True)
            '    Timer5.Interval = SetNotificationRefreshTimeInMinutes * 60 * 1000
            '    Timer5.Start()
            'ElseIf ShowNotificationInMDI > 0 Then
            '    Timer5.Interval = ShowNotificationInMDI * 60 * 1000
            '    Timer5.Start()
            'Else
            '    Timer5.Enabled = False
            'End If

            If ShowNotificationWithoutSMSAPP = False Then
                If ShowNotificationInMDI > 0 Then
                    Timer5.Interval = ShowNotificationInMDI * 60 * 1000
                    Timer5.Start()
                Else
                    Timer5.Enabled = False
                End If
            End If
        Else
            IsDBRestored = True
            isAutoClosing = True
            Me.Close()
        End If
        ' DeleteTempData = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DeleteTempData, clsFixedParameterCode.DeleteTempData, Nothing)) = 1)

    End Sub
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        Try
            If objCommonVar.SystemClockStatus Then
                lblServerDate.Text = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(Nothing, "dd/MM/yyyy hh:mm tt", False), "dd/MM/yyyy hh:mm tt")
            End If
        Catch ex As Exception
        End Try
    End Sub


    Sub LoadClientImage()
        Try
            Dim img As Byte() = DirectCast(clsDBFuncationality.getSingleValue("select top 1 Logo_Img  from tspl_company_master "), Byte())
            Dim ms As MemoryStream = New MemoryStream(img)
            PicClient.Image = Image.FromStream(ms)
        Catch ex As Exception

        End Try

    End Sub

    Function CheckConfigFile() As Boolean
        If Not File.Exists("config.Txp") Then
            Dim frm As New FrmServerConfig
            frm.ShowDialog()
            If Not File.Exists("config.Txp") Then
                Return False
            End If
        End If
        Return True
    End Function

    Sub LoadWelcomeScreen()
        SplitPanel2.Collapsed = False
        SplitPanel3.Collapsed = True
        SplitPanel4.Collapsed = True
        SplitPanel1.Collapsed = True
        Dim myAssembly As Assembly = Assembly.GetExecutingAssembly()
        Dim myAssemblyName As AssemblyName = myAssembly.GetName()
        lblVersion.Text = clsCommon.myCstr(myAssemblyName.Version).Trim()
        Dim aDescAttr As AssemblyDescriptionAttribute = AssemblyDescriptionAttribute.GetCustomAttribute(myAssembly, GetType(AssemblyDescriptionAttribute)) ' clsCommon.GetPrintDate(File.GetCreationTime(Application.StartupPath + "\ERP.exe"), "dd-MMM-yyyy")
        'lblCreatedDate.Text = aDescAttr.Description.ToString
        lblVersion.Text += " [" + aDescAttr.Description.ToString.Trim() + "]"
        SetConnectionWithCommonDLL("")
        objCommonVar.CurrentCompanyCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select comp_Code from tspl_company_master"))
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDP") = CompairStringResult.Equal Then
            SplitContainer2.Panel1Collapsed = True
            SplitContainer2.Panel2Collapsed = False
            PictureBox8.Visible = False
            PictureBox7.Visible = True
        ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "RCDFCF") = CompairStringResult.Equal Then
            SplitContainer2.Panel1Collapsed = True
            SplitContainer2.Panel2Collapsed = False
            PictureBox8.Visible = True
            PictureBox7.Visible = False
        Else
            SplitContainer2.Panel1Collapsed = False
            SplitContainer2.Panel2Collapsed = True
        End If


        LoadClientImage()
        If Not CallCreateTableFunction() Then
            Exit Sub
        End If

        '' Added By Pankaj Jha As suggested by balwinder sir, to bring login screen directly after startup based on setting
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.LoadLoginScreen, clsFixedParameterCode.LoadLoginScreenDirectlyAfterStartup, Nothing)) = 0 Then
            RadCarousel1.AutoLoopPauseCondition = Telerik.WinControls.UI.AutoLoopPauseConditions.None
            Dim carouselEllipsePath1 As New Telerik.WinControls.UI.CarouselEllipsePath()
            carouselEllipsePath1.Center = New Telerik.WinControls.UI.Point3D(49.358974358974358, 46.315789473684212, -20)
            carouselEllipsePath1.FinalAngle = 60
            carouselEllipsePath1.InitialAngle = 60
            carouselEllipsePath1.U = New Telerik.WinControls.UI.Point3D(37.93530426465815, -18.191666666666663, 0)
            carouselEllipsePath1.V = New Telerik.WinControls.UI.Point3D(-11.489983091663683, -15.391666666666662, -20)
            carouselEllipsePath1.ZScale = 60
            RadCarousel1.CarouselPath = carouselEllipsePath1
            RadCarousel1.Dock = System.Windows.Forms.DockStyle.Fill
            RadCarousel1.EnableAutoLoop = True
            RadCarousel1.AutoLoopDirection = AutoLoopDirections.Forward
            RadCarousel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
            RadCarousel1.ForeColor = System.Drawing.Color.Black
            RadCarousel1.ImageScalingSize = New System.Drawing.Size(0, 0)
            'RadCarousel1.Items.AddRange(New Telerik.WinControls.RadItem() {radButtonElement1, radButtonElement2, radButtonElement3, radButtonElement4, radButtonElement5, radButtonElement6, radButtonElement7, radButtonElement8, radButtonElement9})
            RadCarousel1.Location = New System.Drawing.Point(0, 132)
            RadCarousel1.MinFadeOpacity = 0.5
            'RadCarousel1.Name = "RadCarousel1"
            RadCarousel1.NavigationButtonsOffset = New System.Drawing.Size(15, 15)
            Dim imagePrimitive = New ImagePrimitive()

            imagePrimitive.Image = Global.ERP.My.Resources.Resources.BackImageXpertERP

            imagePrimitive.Alignment = ContentAlignment.TopRight
            imagePrimitive.StretchHorizontally = False
            imagePrimitive.StretchVertically = False
            imagePrimitive.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed
            RadCarousel1.CarouselElement.Children.Insert(1, imagePrimitive)
            MyLabel2.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            llblLogin.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Else
            LoadLoginScreen()
        End If
        ''Check Licence
        CheckLicence()
        ''End of Check Licence

    End Sub

    Private Sub llblLogin_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblLogin.LinkClicked
        LoadLoginScreen()
    End Sub

    Sub LoadLoginScreen()
        SplitPanel1.Collapsed = True
        SplitPanel3.Collapsed = True
        SplitPanel4.Collapsed = True
        SplitPanel2.Collapsed = False
        LoadCompany()
        LoadDataBase()
        ddllocationfill()
        Dim myCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture
        OLDshortDate = myCulture.DateTimeFormat.ShortDatePattern
        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy")
        txtUserName.Focus()

        Dim qry As String = "Select User_Code,Password from TSPL_USER_MASTER where SSO='" + System.Environment.UserName + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            txtUserName.Text = clsCommon.myCstr(dt.Rows(0)("User_Code"))
            txtPassword.Text = clsCommon.DecryptString(clsCommon.myCstr(dt.Rows(0)("Password")))
            CheckAndLogin()
        End If
    End Sub

    Private Const LOCALE_USER_DEFAULT = &H400
    Private Const LOCALE_SSHORTDATE = &H1F ' short date format string
    Private Const LOCALE_SLONGDATE = &H20 ' long date format string
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


    Private Sub SetConnectionWithCommonDLL(ByVal strDatabase As String)
        Try
            Dim line As String
            Dim RemaningStr As String = ""
            Dim objReader As New System.IO.StreamReader("config.Txp")
            Do While objReader.Peek() <> -1
                line = objReader.ReadLine()
                '-------------change the existing dbname with dbname comes from company master-BM00000003569------------
                Dim indexofhash As Integer = line.IndexOf("#")
                objCommonVar.Database_Server = line.Substring(0, indexofhash).Trim
                RemaningStr = line.Substring(indexofhash + 1, line.Length - indexofhash - 1).Trim
                Dim reststr As String = line.Substring(indexofhash + 1, line.Length - indexofhash - 1)
                indexofhash = reststr.IndexOf("#")
                RemaningStr = RemaningStr.Substring(indexofhash + 1, RemaningStr.Length - indexofhash - 1).Trim
                reststr = reststr.Substring(0, indexofhash)

                If clsCommon.myLen(objCommonVar.CurrDatabase) > 0 Then
                    line = line.Replace(reststr, "  " + objCommonVar.CurrDatabase + " ")
                Else
                    objCommonVar.CurrDatabase = reststr
                End If

                indexofhash = RemaningStr.IndexOf("#")
                objCommonVar.Database_Server_UserName = RemaningStr.Substring(0, indexofhash).Trim
                objCommonVar.Database_Server_Password = RemaningStr.Substring(indexofhash + 1, RemaningStr.Length - indexofhash - 1).Trim
                '--------------------------------------------------------------------------------------------
                clsDBFuncationality.SetConnectionEncryptFormat(line)
                objCommonVar.ConnString = clsDBFuncationality.connectionString
            Loop
            ''stuti regarding memory leakage
            objReader.Close()
            objReader.Dispose()
            connectSql.strConn = clsDBFuncationality.connectionString
            lblServerDate.Text = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MM/yyyy hh:mm tt")
            lblDataBase.Text = objCommonVar.CurrDatabase.Trim() + "[" + clsCommon.myCstr(clsDBFuncationality.getSingleValue("select @@SPID")) + "]"
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function CallCreateTableFunction() As Boolean
        Dim serverDate As Date = clsCommon.GetDateWithStartTime(clsCommon.GETSERVERDATE())
        Dim FILE_NAME As String = Application.StartupPath + "\Table.Txp"
        Dim myAssembly As Assembly = Assembly.GetExecutingAssembly()
        Dim myAssemblyName As AssemblyName = myAssembly.GetName()
        Dim CurrEXEVersion As String = clsCommon.myCstr(myAssemblyName.Version).Trim()
        Dim dbEXEVersion As String = ""
        Try
            dbEXEVersion = clsDBFuncationality.getSingleValue("select Last_Version from TSPL_Version_Info")
        Catch ex As Exception
            dbEXEVersion = ""
        End Try

        Try
            Dim strTempVersion As String = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertCommon.dll").FileVersion
            If Not clsCommon.CompairString(strTempVersion, "2.1.6.97") = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertCommon ")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPBlankTableScript.dll").FileVersion
            If Not clsCommon.CompairString(strTempVersion, "1.0.0.1") = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPBlankTableScript ")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPEngine.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPEngine ")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPHRandPayroll.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPHRandPayroll ")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPSalesAndDistribution.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPSalesAndDistribution ")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPAdminServices.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPAdminServices")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPCommanServices.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPCommanServices")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPReco.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPReco")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPJobWorkOutward.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPJobWorkOutward")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPFixedAssets.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPFixedAssets")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPTDS.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPTDS")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPService.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPService")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPRiceProduction.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPRiceProduction")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPFarmerPayment.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPFarmerPayment")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPDairySale.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPDairySale")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPParavetServices.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPParavetServices")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPProjectManagement.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPProjectManagement")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPElectrical.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPElectrical")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPBulkProcurement.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPBulkProcurement")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPGeneralLedger.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPGeneralLedger")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPEngineeringAndPlantManagement.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPEngineeringAndPlantManagement")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPCustomerComplaint.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPCustomerComplaint")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPProcessProduction.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPProcessProduction")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPPurchase.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPPurchase")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPEngineFine.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPEngineFine")
            End If
            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertERPMIS.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertERPMIS")
            End If

            strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertErpBullManagement.dll").FileVersion
            If Not clsCommon.CompairString(CurrEXEVersion, strTempVersion) = CompairStringResult.Equal Then
                Throw New Exception("Wrong DLL Version" + Environment.NewLine + "XpertErpBullManagement")
            End If
            '--Check Apps Version
            If File.Exists(Application.StartupPath + "\XpertSMSApp.exe") Then
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertSMSApp.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "1.0.1.8") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertSMSApp")
                End If
            End If
            If File.Exists(Application.StartupPath + "\XpertBookingSchedularApp.exe") Then ''Not Upgraded Teleik 2022
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertBookingSchedularApp.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "1.0.1.2") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertBookingSchedularApp")
                End If
            End If

            If File.Exists(Application.StartupPath + "\XpertDispatchSchedularApp.exe") Then ''Not Upgraded Teleik 2022
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertDispatchSchedularApp.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "1.0.0.1") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertDispatchSchedularApp")
                End If
            End If

            If File.Exists(Application.StartupPath + "\XpertDataSync.exe") Then
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertDataSync.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "2.0.0.20") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertDataSync")
                End If
            End If
            If File.Exists(Application.StartupPath + "\XpertBioMetricSync.exe") Then
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertBioMetricSync.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "1.0.0.9") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertBioMetricSync")
                End If
            End If
            If File.Exists(Application.StartupPath + "\XpertAlertApp.exe") Then ''Not Upgraded Teleik 2022
                strTempVersion = FileVersionInfo.GetVersionInfo(Application.StartupPath + "\XpertAlertApp.exe").FileVersion
                If Not clsCommon.CompairString(strTempVersion, "1.0.0.3") = CompairStringResult.Equal Then
                    Throw New Exception("Wrong Exe Version" + Environment.NewLine + "XpertAlertApp")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            isAutoClosing = True
            Me.Close()
            Return False
        End Try



        Dim dtTE As DataTable
        If System.IO.File.Exists(FILE_NAME) OrElse clsCommon.myLen(dbEXEVersion) <= 0 OrElse clsCommon.CompairString(CurrEXEVersion, dbEXEVersion) = CompairStringResult.Greater Then
            Try
                Dim qryFun As String = " select * from Information_schema.Routines where SPECIFIC_SCHEMA='dbo' AND SPECIFIC_NAME = 'fnColList' AND Routine_Type='FUNCTION' "
                Dim dtt As DataTable = clsDBFuncationality.GetDataTable(qryFun)

                If dtt Is Nothing OrElse dtt.Rows.Count <= 0 Then
                    qryFun = " create function fnColList(@in_vcTbl_name varchar(8000)) "
                    qryFun = qryFun & " returns varchar(8000) "
                    qryFun = qryFun & " as "
                    qryFun = qryFun & " begin  "
                    qryFun = qryFun & " declare @colList2BuildAuditTable  varchar(max) "
                    qryFun = qryFun & " SELECT @colList2BuildAuditTable = coalesce(@colList2BuildAuditTable+ ',', '')+ ''+ B.NAME +''   "
                    qryFun = qryFun & " FROM SYSOBJECTS A JOIN SYSCOLUMNS B ON A.ID = B.ID  "
                    qryFun = qryFun & " WHERE A.ID = OBJECT_ID(@in_vcTbl_name)  "
                    qryFun = qryFun & " ORDER BY B.COLORDER "
                    qryFun = qryFun & " return @colList2BuildAuditTable  "
                    qryFun = qryFun & " End "
                    clsDBFuncationality.ExecuteNonQuery(qryFun)
                End If
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(ex.Message)
            End Try

            Dim RunTables As String = Application.StartupPath + "\RunTables.Txp"
            If Not System.IO.File.Exists(RunTables) Then
                XpertERPEngine.clsCreateAllTables.CreateAllTable()
                XpertERPEngine.clsAllSQLView.CreateAllSQLView()
                XpertERPEngine.clsAllSQLFunction.CreateAllSQLFunction()
                XpertERPEngine.clsAllStoreProcedure.CreateAllStoreProcedure()
                XpertERPEngine.clsAllSQLTrigger.CreateAllTrigger()
            Else
                If clsCommon.MyMessageBoxShow("Do you want to run tables ", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    XpertERPEngine.clsCreateAllTables.CreateAllTable()
                    XpertERPEngine.clsAllSQLView.CreateAllSQLView()
                    XpertERPEngine.clsAllSQLFunction.CreateAllSQLFunction()
                    XpertERPEngine.clsAllStoreProcedure.CreateAllStoreProcedure()
                    XpertERPEngine.clsAllSQLTrigger.CreateAllTrigger()
                End If
            End If


            dtTE = clsDBFuncationality.GetDataTable("select top 1 Comp_Code from TSPL_COMPANY_MASTER")
            Dim isFirstTime As Boolean = False
            If dtTE Is Nothing OrElse dtTE.Rows.Count <= 0 Then
                isFirstTime = True
                Dim frm As New FrmCompany()
                frm.ShowDialog()
                If Not frm.isOperationSucced Then
                    isAutoClosing = True
                    Me.Close()
                    Exit Function
                End If
            Else
                objCommonVar.CurrentCompanyCode = clsCommon.myCstr(dtTE.Rows(0)("Comp_Code"))
            End If
            If Not System.IO.File.Exists(FILE_NAME) Then
                Qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select max(Version_No) as Version_No from TSPL_Exe_Deployment"))
                If CurrEXEVersion > Qry Then
                    clsFixedParameter.UpdateData(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, clsCommon.EncryptString(clsCommon.GetPrintDate(serverDate.AddMonths(4), "dd/MMM/yyyy"), objCommonVar.CurrentCompanyCode), Nothing)
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_FIXED_PARAMETER set Specification=1 where Code='" + clsFixedParameterCode.BigValidity + "' and Type ='" + clsFixedParameterType.BigValidity + "'")
                End If
            End If

            Dim Exe As String = clsDBFuncationality.getSingleValue("select Version_No from TSPL_Exe_Deployment where Version_No= '" + CurrEXEVersion + "' ")
            If CurrEXEVersion <> Exe Then
                clsDBFuncationality.ExecuteNonQuery("insert into TSPL_Exe_Deployment select '" + CurrEXEVersion + "',' " + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") + "'")
            End If


            'FrmUtility.CreateIndex()
            ProgramCodeNew.SetProgramCode()
            clsFixedParameter.FixedParameterValues()
            clsBIFilterDeveloper.CreateDeveloperFilter()
            clsBIReportDeveloperReports.CreateDeveloperReports()
            ' ''If Not isFirstTime Then
            ' ''    XpertERPEngine.clsPostCreateTable.Post_AlterOrUpdateAllTables(dbEXEVersion)
            ' ''End If


            XpertERPEngine.clsCancelTableClass.CancelTableValues()
            XpertERPEngine.clsCancelTableClass.CancelValidationValues()
            XpertERPEngine.clsCancelTableClass.CancelConditionTableValues()

            ''To Run Customize Function
            Dim intCustoizeDLLCount As Integer = 0
            If System.IO.File.Exists(Application.StartupPath + "\XpertErpMPD.dll") Then
                intCustoizeDLLCount += 1
            End If
            If System.IO.File.Exists(Application.StartupPath + "\XpertErpViney.dll") Then
                intCustoizeDLLCount += 1
            End If
            If System.IO.File.Exists(Application.StartupPath + "\XpertErpJakson.dll") Then
                intCustoizeDLLCount += 1
            End If
            If System.IO.File.Exists(Application.StartupPath + "\XpertErpPatanjali.dll") Then
                intCustoizeDLLCount += 1
            End If
            If intCustoizeDLLCount > 1 Then
                Throw New Exception("More Than one Customize DLL exists here.Please remove excess customize dll.")
            End If

            'If System.IO.File.Exists(Application.StartupPath + "\XpertErpMPD.dll") Then
            '    frmAppIntegrator.CallStartupFunction("XpertErpMPD.dll")
            'ElseIf System.IO.File.Exists(Application.StartupPath + "\XpertErpViney.dll") Then
            '    frmAppIntegrator.CallStartupFunction("XpertErpViney.dll")
            'ElseIf System.IO.File.Exists(Application.StartupPath + "\XpertErpJakson.dll") Then
            '    frmAppIntegrator.CallStartupFunction("XpertErpJakson.dll")
            'ElseIf System.IO.File.Exists(Application.StartupPath + "\XpertErpPatanjali.dll") Then
            '    frmAppIntegrator.CallStartupFunction("XpertErpPatanjali.dll")
            'End If
            ''To Run Customize Function
            clsDBFuncationality.ExecuteNonQuery("delete from TSPL_Version_Info")
            clsDBFuncationality.ExecuteNonQuery("insert into TSPL_Version_Info(Last_Version) Values('" + CurrEXEVersion + "')")
        End If



        dtTE = clsDBFuncationality.GetDataTable("select top 1 Comp_Code from TSPL_COMPANY_MASTER")
        If dtTE Is Nothing OrElse dtTE.Rows.Count <= 0 Then

            Dim frm As New FrmCompany()
            frm.ShowDialog()
            If Not frm.isOperationSucced Then
                isAutoClosing = True
                Me.Close()
            End If
        Else
            objCommonVar.CurrentCompanyCode = clsCommon.myCstr(dtTE.Rows(0)("Comp_Code"))
        End If

        Dim strFixVersion As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Fix_Version from TSPL_Version_Fix"))
        If clsCommon.myLen(strFixVersion) > 0 Then
            If Not clsCommon.CompairString(CurrEXEVersion, strFixVersion) = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow(Me, "Fixed Application version is  :" + strFixVersion + " and your  Current Version :" + CurrEXEVersion)
                Application.Exit()
            End If
        Else
            dbEXEVersion = clsDBFuncationality.getSingleValue("select Last_Version from TSPL_Version_Info")
            If Not clsCommon.CompairString(CurrEXEVersion, dbEXEVersion) = CompairStringResult.Equal Then
                IsDBRestored = True
                clsCommon.MyMessageBoxShow(Me, "Application version is not updated." + Environment.NewLine + "Update Version :" + dbEXEVersion + " Current Version :" + CurrEXEVersion)
                For Each P As Process In Process.GetProcessesByName("XpertAlertApp")
                    P.Kill()
                Next
                Try
                    System.Diagnostics.Process.Start("XpertCopyEXE.exe")
                Catch ex As Exception
                End Try
                Application.Exit()
            End If
        End If

        ''richa 9 June ,2020 to update batch file specification through ClearCache.txt file
        ExecuteQueryForBatchLicenceThroughNotepadFile()

        ''7 Jan 2021 delete einvoice credentials from table in database iis local and credentials of live
        DeleteEinvoiceCredentialsfromLocalDatabase()

        Try
            If clsCommon.myCdbl(clsFixedParameter.GetSpecification(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, Nothing)) <> 1 Then
                ''richa done on 9 June,2020
                'Throw New Exception("XXX")
            End If
            Qry = clsFixedParameter.GetData(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, Nothing)
            If clsCommon.myLen(Qry) <= 0 Then
                ''richa done on 9 June,2020
                'Throw New Exception("XXX")
            End If
            Qry = clsCommon.DecryptString(Qry, objCommonVar.CurrentCompanyCode)
            Dim validdate As Date = clsCommon.myCDate(Qry)
            If serverDate > validdate Then
                clsDBFuncationality.ExecuteNonQuery("Update TSPL_FIXED_PARAMETER set Specification=0 where Code='" + clsFixedParameterCode.BigValidity + "' and Type ='" + clsFixedParameterType.BigValidity + "'")
                Qry = "select max(SNo) as TotalMax,max(SNo * case when Is_Current=1 then 1 else 0 end) as CurrentMax from TSPL_Exception"
                dt = clsDBFuncationality.GetDataTable(Qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Qry = "update TSPL_Exception set Is_Current=0"
                    clsDBFuncationality.ExecuteNonQuery(Qry)
                    If clsCommon.myCdbl(dt.Rows(0)("TotalMax")) = clsCommon.myCdbl(dt.Rows(0)("CurrentMax")) Then
                        Qry = "update TSPL_Exception set Is_Current=1 where SNo=1"
                        clsDBFuncationality.ExecuteNonQuery(Qry)
                    Else
                        Qry = "update TSPL_Exception set Is_Current=1 where SNo=" + clsCommon.myCstr(clsCommon.myCdbl(dt.Rows(0)("CurrentMax")) + 1) + ""
                        clsDBFuncationality.ExecuteNonQuery(Qry)
                    End If
                End If

                ''Write update query.
                ''richa done on 9 June,2020
                'Throw New Exception("XXX")
            End If
        Catch ex As Exception
            If clsCommon.CompairString(ex.Message, "XXX") = CompairStringResult.Equal Then
                Qry = "select Exception_Msg from TSPL_Exception where Is_Current=1"
                dt = clsDBFuncationality.GetDataTable(Qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If clsCommon.myLen(dt.Rows(0)("Exception_Msg")) > 0 Then
                        clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Exception_Msg")), Me.Text)
                    End If
                End If
            End If
            isAutoClosing = True
            Me.Close()
            Return False
        End Try
        Return True
    End Function
    Public Sub DeleteEinvoiceCredentialsfromLocalDatabase()
        Try
            If File.Exists("Einvoicetest.txp") Then
                Dim qry As String = " Select * from(select * from TSPL_EINVOICEHEADER_INFO where username ='mastergst' and password='Malli#123' " &
    " union all " &
    " Select * From TSPL_EINVOICEHEADER_INFO Where UserName ='CLEARTAX' and url='https://einvoicing.internal.cleartax.co/v2/eInvoice/generate' " &
    " union all " &
    " Select * From TSPL_EINVOICEHEADER_INFO Where UserName ='CLEARTAX' and url='https://einvoicing.internal.cleartax.co/v2/eInvoice/cancel') z"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Else
                    clsDBFuncationality.ExecuteNonQuery("delete from TSPL_EINVOICEHEADER_INFO")
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub


    Public Sub ExecuteQueryForBatchLicenceThroughNotepadFile()
        Try
            If clsCommon.myCdbl(clsFixedParameter.GetSpecification(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, Nothing)) <> 1 Then
                If File.Exists("ClearCache.txt") Then
                    '' write code to read data from file
                    Dim objStreamReader As StreamReader
                    Dim CompanyCode As String
                    Dim strDate As String
                    Dim FormCounter As String = 0
                    'Pass the file path and the file name to the StreamReader constructor.
                    objStreamReader = New StreamReader(Application.StartupPath + "\ClearCache.txt")

                    'Read the first line of text.
                    strDate = objStreamReader.ReadLine


                    'Read the next line.
                    CompanyCode = objStreamReader.ReadLine

                    FormCounter = objStreamReader.ReadLine

                    Dim strDBCompanyCode As String = clsDBFuncationality.getSingleValue("select top 1 Comp_Code from TSPL_COMPANY_MASTER")
                    If clsCommon.DecryptString(CompanyCode) <> strDBCompanyCode Then
                        Throw New Exception("file Is Not in correct format")
                    End If

                    clsFixedParameter.UpdateData(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, clsCommon.EncryptString(clsCommon.DecryptString(strDate), clsCommon.DecryptString(CompanyCode)), Nothing)
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_FIXED_PARAMETER set Specification=1 where Code='" + clsFixedParameterCode.BigValidity + "' and Type ='" + clsFixedParameterType.BigValidity + "'")
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_PROGRAM_MASTER set FormCounter =0,ExceptionNo=NULL")
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_FIXED_PARAMETER set Description='" & clsCommon.EncryptString(clsCommon.DecryptString(FormCounter)) & "' where Code='" + clsFixedParameterCode.BatchFileCounter + "' and Type ='" + clsFixedParameterType.BatchFileCounter + "'")
                End If


            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub CheckLicence()
        Dim isCloseEXE As Boolean = False
        Try
            ''Check Expiry Date
            Dim dtCurrentDate As Date = clsCommon.GetDateWithStartTime(clsCommon.GETSERVERDATE())
            Dim strSpec As String = clsCommon.DecryptString(clsFixedParameter.GetSpecification(clsFixedParameterType.LicenceExpiryDate, clsFixedParameterCode.LicenceExpiryDate, Nothing), objCommonVar.CurrentCompanyCode + "B")
            Dim strVal As String = clsCommon.DecryptString(clsFixedParameter.GetData(clsFixedParameterType.LicenceExpiryDate, clsFixedParameterCode.LicenceExpiryDate, Nothing), objCommonVar.CurrentCompanyCode + "A")
            If clsCommon.CompairString(strSpec, "-1") = CompairStringResult.Equal Then
            ElseIf clsCommon.CompairString(strSpec, "1") = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow("Application Has Been Expired,For Renewal or More Details," + Environment.NewLine + objCommonVar.LicenceMessageContactPersion, "Attention")
                isCloseEXE = True
            ElseIf clsCommon.CompairString(strSpec, "0") = CompairStringResult.Equal Then
                Try
                    Dim dt As Date = clsCommon.myCDate(strVal)
                    Dim remDays As Integer = DateDiff(DateInterval.Day, dtCurrentDate, dt)
                    If remDays < 0 Then
                        Throw New Exception("Application Has Been Expired,For Renewal or More Details," + Environment.NewLine + objCommonVar.LicenceMessageContactPersion)
                    ElseIf remDays = 0 Then
                        clsCommon.MyMessageBoxShow("Last Day of your ERP Application.For Renewal or More Details " + Environment.NewLine + objCommonVar.LicenceMessageContactPersion + Environment.NewLine + ".Please purchase the licence", "Attention")
                    ElseIf remDays <= 10 Then
                        clsCommon.MyMessageBoxShow("Application will be Expired after " + clsCommon.myCstr(remDays) + " Days" + Environment.NewLine + objCommonVar.LicenceMessageContactPersion + Environment.NewLine + ".Please purchase the licence", "Attention")
                    End If
                Catch ex As Exception
                    clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
                    Qry = "update TSPL_FIXED_PARAMETER set Specification='" + clsCommon.EncryptString("1", objCommonVar.CurrentCompanyCode + "B") + "' where Type='" + clsFixedParameterType.LicenceExpiryDate + "' and Code='" + clsFixedParameterCode.LicenceExpiryDate + "'"
                    clsDBFuncationality.ExecuteNonQuery(Qry)
                    isCloseEXE = True
                End Try
            End If
            ''End of Check Expiry Date
        Catch exx As Exception
            clsCommon.MyMessageBoxShow(exx.Message, Me.Text)
            isCloseEXE = True
        End Try
        If isCloseEXE Then
            If clsCommon.MyMessageBoxShow("Do you want to enter product key", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                Dim frm As New FrmLicenceActivate()
                frm.ShowDialog()
            End If
            isAutoClosing = True
            Me.Close()
            Exit Sub
        End If
    End Sub



    Public Sub LoadCompany()
        Try
            Dim qry As String = "select Comp_Code as Code,Comp_Name as Name from TSPL_COMPANY_MASTER"
            cboCompany.DataSource = clsDBFuncationality.GetDataTable(qry)
            cboCompany.ValueMember = "Code"
            cboCompany.DisplayMember = "Name"
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub

    Private Sub LoadDataBase()
        Try
            Dim Qry As String = "select DataBase_Name as DB, Comp_Name as Name from TSPL_COMPANY_MASTER"
            cmbDB.DataSource = clsDBFuncationality.GetDataTable(Qry)
            cmbDB.DisplayMember = "Name"
            cmbDB.ValueMember = "DB"
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ddllocationfill()
        Try
            Dim strquery As String = "select segment_code,description from TSPL_GL_SEGMENT_CODE where Seg_No='7'"
            transportSql.FillComboBox(strquery, ddllocation, "description", "segment_code")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message)
        End Try
    End Sub

    Public Shared Function RefeshUserApplicableLocationsAndGLAccount() As Boolean
        If clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
            objCommonVar.arrCurrUserLocations = Nothing
            objCommonVar.strCurrUserLocations = ""
            Dim qry As String = "SELECT SEGMENT_CODE FROM TSPL_GL_SEGMENT_CODE WHERE TSPL_GL_SEGMENT_CODE.Seg_No='7'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                objCommonVar.strCurrUserLocationsSegment = ""
                For Each dr As DataRow In dt.Rows
                    If clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                        objCommonVar.strCurrUserLocationsSegment += ","
                    End If
                    objCommonVar.strCurrUserLocationsSegment += "'" + clsCommon.myCstr(dr("Segment_Code")) + "'"
                Next
            End If
        Else
            ''richa agarwal 24 Feb,2020 to select location alaong with its mapped user
            ' Dim qry As String = "select Segment_Code from TSPL_GL_SEGMENT_PERMISSION where User_Code='" + objCommonVar.CurrentUserCode + "' and GL_Segment='7'"
            Dim qry As String = "select distinct Segment_Code from TSPL_GL_SEGMENT_PERMISSION where User_Code in (select '" + objCommonVar.CurrentUserCode + "' union  select TSPL_USER_MAPPING_DETAIL.Mapped_UserCode AS User_Code  from TSPL_USER_MAPPING_DETAIL Where TSPL_USER_MAPPING_DETAIL.User_Code='" + objCommonVar.CurrentUserCode + "' ) and GL_Segment='7'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            objCommonVar.strCurrUserLocationsSegment = "''"
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                objCommonVar.strCurrUserLocationsSegment = ""
                For Each dr As DataRow In dt.Rows
                    If clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                        objCommonVar.strCurrUserLocationsSegment += ","
                    End If
                    objCommonVar.strCurrUserLocationsSegment += "'" + clsCommon.myCstr(dr("Segment_Code")) + "'"
                Next
            End If

            qry = "select Location_Code from TSPL_LOCATION_MASTER where Loc_Segment_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
            dt = clsDBFuncationality.GetDataTable(qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                objCommonVar.arrCurrUserLocations = New List(Of String)
                For Each dr As DataRow In dt.Rows
                    objCommonVar.arrCurrUserLocations.Add(clsCommon.myCstr(dr("Location_Code")))
                Next
                objCommonVar.strCurrUserLocations = clsCommon.GetMulcallString(objCommonVar.arrCurrUserLocations)
            End If

            qry = "select Account_Code from TSPL_GL_ACCOUNTS where Account_Seg_Code7 in (select segment_code from TSPL_GL_SEGMENT_PERMISSION where User_Code='" + objCommonVar.CurrentUserCode + "' and GL_Segment='7') "
            qry += " union "
            qry += " select Account_Code from TSPL_GL_ACCOUNT_PERMISSION where User_Code='" + objCommonVar.CurrentUserCode + "'"
            dt = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                objCommonVar.strCurrUserGLAccount = qry
            End If
        End If
        Return True
    End Function

    Public Function RefeshUserApplicableZonesAndCustomer() As Boolean
        Dim strZone As String = ""
        'If clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
        '    objCommonVar.strCurrUserCustomers = ""
        '    objCommonVar.strCurrUserZones = ""
        '    Dim qry As String = "SELECT ZONE_CODE FROM TSPL_ZONE_MASTER"
        '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '        objCommonVar.strCurrUserZones = ""
        '        For Each dr As DataRow In dt.Rows
        '            If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
        '                objCommonVar.strCurrUserZones += ","
        '            End If
        '            objCommonVar.strCurrUserZones += "'" + clsCommon.myCstr(dr("ZONE_CODE")) + "'"
        '        Next
        '    End If
        'Else
        objCommonVar.strCurrUserCustomers = ""
        objCommonVar.strCurrUserZones = ""
        Dim qry As String = "select TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description from TSPL_ZONE_MASTER 
                                LEFT JOIN TSPL_USER_CUSTOMER_ZONE ON TSPL_USER_CUSTOMER_ZONE.Zone_Code=TSPL_ZONE_MASTER.Zone_Code
                                LEFT JOIN TSPL_USER_MASTER ON TSPL_USER_MASTER.User_Code=TSPL_USER_CUSTOMER_ZONE.User_Code
                                WHERE TSPL_USER_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            objCommonVar.strCurrUserZones = ""
            For Each dr As DataRow In dt.Rows
                If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
                    objCommonVar.strCurrUserZones += ","
                    strZone += ","
                End If
                objCommonVar.strCurrUserZones += "'" + clsCommon.myCstr(dr("ZONE_CODE")) + "'"
                strZone += clsCommon.myCstr(dr("Description"))
            Next
        End If
        'Customers
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            Dim qry1 As String = "DECLARE @colsScheme AS NVARCHAR(MAX),@query  AS NVARCHAR(MAX) SELECT   STUFF((SELECT distinct ',' +'''' + ( Customer_Code ) +''''  as Alies_Name
                                      FROM tspl_customer_master left outer join TSPL_CUSTOMER_LOCATION_MAPPING on TSPL_CUSTOMER_LOCATION_MAPPING.Customer_Code= TSPL_CUSTOMER_MASTER.Cust_Code 
                                      WHERE TSPL_CUSTOMER_LOCATION_MAPPING.Location_Code in (" + objCommonVar.strCurrUserLocations + ") FOR XML PATH(''), TYPE ).value('.', 'NVARCHAR(MAX)') ,1,1,'') "
            objCommonVar.strCurrUserCustomers = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry1))
        End If
        If clsCommon.myLen(strZone) > 0 Then
            lblUser.Text += "[" + strZone + "]"
        End If
        'End If
    End Function

    Public Sub CheckLicenceNoOfUsers()
        Dim isCloseEXE As Boolean = False
        Try
            Dim strVal As String = ""
            ''Check No of connection
            If Not isCloseEXE Then
                strVal = clsCommon.DecryptString(clsFixedParameter.GetData(clsFixedParameterType.LicenceNoOfExeConnection, clsFixedParameterCode.LicenceNoOfExeConnection, Nothing), objCommonVar.CurrentCompanyCode + "C")
                If clsCommon.CompairString(strVal, "-1") = CompairStringResult.Equal Then
                Else
                    Dim qry As String = clsLoginInfo.funGetActiveUserQuery(True)
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If SettSameUserCanNotloginmultipletimes Then
                        For ii As Integer = 0 To dt.Rows.Count - 1
                            If clsCommon.CompairString(objCommonVar.CurrentUserCode, clsCommon.myCstr(dt.Rows(ii)("ERP User Code"))) = CompairStringResult.Equal Then
                                If clsCommon.CompairString("Reserved Licence", clsCommon.myCstr(dt.Rows(ii)("Login Code"))) = CompairStringResult.Equal Then
                                    dt.Rows.RemoveAt(ii)
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    If dt.Rows.Count + 1 > clsCommon.myCdbl(strVal) Then
                        clsCommon.MyMessageBoxShow("Please ask your administrator to purchase licence for more users" + Environment.NewLine + objCommonVar.LicenceMessageContactPersion, Me.Text)
                        If clsCommon.MyMessageBoxShow("Do you want to open Active User report", "Active Users", MessageBoxButtons.YesNo, WinControls.RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                            ShowForm(clsUserMgtCode.rptActiveUsers, "Active User", False, "$#$#", False, False)
                        End If
                        isCloseEXE = True
                    End If
                End If
            End If
            ''End of Check No of connection
        Catch exx As Exception
            If exx.Message.Contains("Cannot resolve the collation conflict between") Then
                Try
                    Dim strTempBreak As String() = exx.Message.Split("""")
                    Qry = "ALTER TABLE TSPL_USERLOGIN_INFO ALTER COLUMN Machine_Name varchar(50) COLLATE " + strTempBreak(3)
                    clsDBFuncationality.ExecuteNonQuery(Qry)
                    Qry = "ALTER TABLE TSPL_USERLOGIN_INFO ALTER COLUMN Window_User_Name varchar(50) COLLATE " + strTempBreak(3)
                    clsDBFuncationality.ExecuteNonQuery(Qry)
                    CheckLicenceNoOfUsers()
                Catch ex123 As Exception
                    clsCommon.MyMessageBoxShow(exx.Message + Environment.NewLine + ex123.Message, Me.Text)
                    isCloseEXE = True
                End Try

            Else
                clsCommon.MyMessageBoxShow(exx.Message, Me.Text)
                isCloseEXE = True
            End If

        End Try
        If isCloseEXE Then
            isAutoClosing = True
            Me.Close()
            Exit Sub
        End If
    End Sub
    Sub DeleteOldFiles(folderPath As String, referenceTime As DateTime)
        Try
            ' Get all files in the folder
            Dim files = Directory.GetFiles(folderPath)

            For Each filePath As String In files
                Dim creationTime As DateTime = File.GetCreationTime(filePath)

                ' Check if the file is older than 24 hours compared to the reference time
                If (referenceTime - creationTime).TotalHours >= 24 Then
                    Try
                        File.Delete(filePath)
                        'Console.WriteLine("Deleted: " & filePath)
                    Catch ex As Exception
                        'Console.WriteLine("Failed to delete: " & filePath & " - " & ex.Message)
                    End Try
                End If
            Next

        Catch ex As Exception
            ' Console.WriteLine("Error accessing folder: " & ex.Message)
        End Try
    End Sub
    ''Sub DeleteOldExportFiles(folderPath As String)
    ''    Try
    ''        ' Get all files in the folder
    ''        Dim files = Directory.GetFiles(folderPath)

    ''        For Each filePath As String In files
    ''            ' Check if the file is older than 24 hours compared to the reference time
    ''            File.Delete(filePath)
    ''            'Console.WriteLine("Deleted: " & filePath)
    ''        Next

    ''    Catch ex As Exception
    ''        ' Console.WriteLine("Error accessing folder: " & ex.Message)
    ''    End Try
    ''End Sub

    Sub CheckAndLogin()
        Dim isLoginError As Boolean = False
        If clsCommon.myLen(txtUserName.Text) <= 0 Then
            clsCommon.MyMessageBoxShow("Please Enter " + txtUserName.MyLinkLable1.Text, Me.Text)
            txtUserName.Focus()
            Exit Sub
        End If

        If clsCommon.myLen(txtPassword.Text) <= 0 Then
            clsCommon.MyMessageBoxShow("Please Enter " + txtPassword.MyLinkLable1.Text, Me.Text)
            txtPassword.Focus()
            Exit Sub
        End If
        Dim DeleteTempDataFromTime As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DeleteTempDataFromTime, clsFixedParameterCode.DeleteTempDataFromTime, Nothing))
        Dim timeParts() As String
        If clsCommon.myLen(DeleteTempDataFromTime) > 0 Then
            timeParts = DeleteTempDataFromTime.Split("-"c)
        End If
        If timeParts IsNot Nothing AndAlso timeParts.Count > 1 Then
            Dim fromTime As DateTime = clsCommon.myCDate(timeParts(0).Trim())
            Dim toTime As DateTime = clsCommon.myCDate(timeParts(1).Trim())
            Dim serverTime As DateTime = clsCommon.myCDate(clsCommon.GETSERVERDATE())
            If serverTime >= fromTime AndAlso serverTime <= toTime Then
                Dim folder As String = "C:\ERPTempFolder"
                'Dim folder As String = "D:\backup"
                Dim userTime As DateTime = DateTime.Now ' 
                DeleteOldFiles(folder, userTime)
            End If
        End If
        PasswordRules = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PasswordRules, clsFixedParameterCode.PasswordRules, Nothing)) = "1", True, False))

        Dim qry As String = "select TSPL_USER_MASTER.password,TSPL_USER_MASTER.User_Code,TSPL_USER_MASTER.User_Name,TSPL_USER_MASTER.Level, ApprovalLevel,ExpiryDate,TSPL_USER_MASTER.IP_Address,TSPL_USER_MASTER.Login_Status,TSPL_USER_MASTER.Modify_Date,TSPL_USER_MASTER.InActive,TSPL_USER_MASTER.HR_Admin,TSPL_USER_MASTER.E_Mail,TSPL_USER_MASTER.Default_Location from TSPL_USER_MASTER where TSPL_USER_MASTER.User_Code='" + txtUserName.Text + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

        ' ImportExportDrive = objCommonVar.ImportExportDrive = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ImportExportDrive, clsFixedParameterCode.ImportExportDrive, Nothing))

        'Dim exportdata As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Export"
        'Dim importdata As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Import"
        'DeleteOldExportFiles(exportdata)


        Dim strIpAdd As String = ""
        Dim strLoginStatus As Boolean = False

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            strIpAdd = clsCommon.myCstr(dt.Rows(0)("IP_Address"))
            strLoginStatus = clsCommon.myCBool(dt.Rows(0)("Login_Status"))

            Dim ExpiryDate As String = clsCommon.myCstr(dt.Rows(0)("ExpiryDate"))
            If clsCommon.myLen(ExpiryDate) > 0 AndAlso clsCommon.myCDate(ExpiryDate) < clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MM/yyyy") Then
                clsCommon.MyMessageBoxShow(Me, "Can't Access in demo version. " + Environment.NewLine + " For any queries/details, contact tecxpert@tecxpert.in. ", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                Exit Sub
            End If
            Dim Pwd As String = clsCommon.myCstr(dt.Rows(0)("password"))


            If clsCommon.CompairString(txtUserName.Text, "admin") = CompairStringResult.Equal Then
                If Not clsCommon.HighSecurityVerifyNumber(clsCommon.myCDecimal(lblOTPCode.Text)) = clsCommon.myCDecimal(txtPassword.Text) Then
                    clsCommon.MyMessageBoxShow(Me, "Wrong OTP.Try again...", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                    Exit Sub
                End If
                lblOTPCode.Visible = False
            Else
                If Not clsCommon.CompairString(Pwd, clsCommon.EncryptString(txtPassword.Text)) = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_MASTER.DBO.TSPL_FIXED_PARAMETER where code='Master PWD'")), txtPassword.Text, True) = CompairStringResult.Equal Then
                        clsCommon.MyMessageBoxShow(Me, "Correct Password is: " & clsCommon.DecryptString(Pwd), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                    Else
                        clsCommon.MyMessageBoxShow(Me, "Please enter Correct User ID and Password ", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                    End If
                    Exit Sub
                End If
            End If

            objCommonVar.CurrentUserCode = clsCommon.myCstr(dt.Rows(0)("User_Code"))
            If SettSameUserCanNotloginmultipletimes Then
                qry = clsLoginInfo.funGetActiveUserQuery(True, clsCommon.myCstr(dt.Rows(0)("User_Code")))
                Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    clsCommon.MyMessageBoxShow(Me, "User [" + dt.Rows(0)("User_Code") + "] is already logged in From IP [" + clsCommon.myCstr(dtTemp.Rows(0)("IP Address")) + "]", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                    Exit Sub
                End If
            End If

            If clsCommon.CompairString("Y", clsCommon.myCstr(dt.Rows(0)("InActive"))) = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow(Me, "You are not active user.", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
                Exit Sub
            End If

            CheckLicenceNoOfUsers()
            If isAutoClosing Then
                Exit Sub
            End If

            clsCommon.LoginId = objCommonVar.CurrentUserCode
            objCommonVar.CurrentUserEmailID = clsCommon.myCstr(dt.Rows(0)("E_Mail"))
            objCommonVar.CurrentUser = clsCommon.myCstr(dt.Rows(0)("User_Name"))
            objCommonVar.CurrUserLevel = clsCommon.myCdbl(dt.Rows(0)("ApprovalLevel"))
            objCommonVar.IsLoginUserHRAdmin = IIf(clsCommon.myCdbl(dt.Rows(0)("HR_Admin")) = 1, True, False)
            objCommonVar.strDefaultUserLocation = clsCommon.myCstr(dt.Rows(0)("Default_Location"))
            qry = "select Comp_Code,Comp_Name,DataBase_Name,BaseCurrencyCode, Case When ApplyMultiCurrency=1 Then 'True' Else 'False' End as ApplyMultiCurrency,Comp_Code1 from TSPL_COMPANY_MASTER where Comp_Code='" + clsCommon.myCstr(cboCompany.SelectedValue) + "'"
            dt = clsDBFuncationality.GetDataTable(qry)
            objCommonVar.CurrentCompanyCode = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            objCommonVar.CurrentCompanyName = clsCommon.myCstr(dt.Rows(0)("Comp_Name"))
            'objCommonVar.CurrDatabase = clsCommon.myCstr(dt.Rows(0)("DataBase_Name"))
            objCommonVar.BaseCurrencyCode = clsCommon.myCstr(dt.Rows(0)("BaseCurrencyCode"))
            objCommonVar.IsMultiCurrencyCompany = clsCommon.myCstr(dt.Rows(0)("ApplyMultiCurrency"))
            objCommonVar.CurrComp_Code1 = clsCommon.myCstr(dt.Rows(0)("Comp_Code1"))
            objCommonVar.CurrLocationCode = clsCommon.myCstr(ddllocation.SelectedValue)
            'objCommonVar.CurrLocationCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
            objCommonVar.CurrLocationName = clsCommon.myCstr(ddllocation.Text)
            '' work done on setting based agaist ticket no. KDI/04/05/18-000296
            If PasswordRules = True Then
                Dim UserNameDays As String = clsDBFuncationality.getSingleValue("select DATEDIFF(dd, convert(datetime,Modify_Date,103), convert(datetime,getdate(),103)) AS DaysDiff from TSPL_USER_MASTER where DATEDIFF(dd, convert(datetime,Modify_Date,103), convert(datetime,getdate(),103))>30 and User_Code='" + txtUserName.Text + "'")
                If clsCommon.myLen(UserNameDays) > 0 Then
                    Dim frm1 As RadForm = New FrmChangePassword()
                    frm1.StartPosition = FormStartPosition.CenterScreen
                    frm1.MaximizeBox = False
                    frm1.MinimizeBox = False
                    frm1.ControlBox = False
                    frm1.Location = Me.PictureBox1.Location
                    frm1.Height = Me.PictureBox1.Height + 150
                    frm1.Width = Me.PictureBox1.Width + 100
                    frm1.ShowDialog()

                End If

            End If
            '' End work




            objCommonVar.RefreshCommonVar()

            'SetConnectionWithCommonDLL(objCommonVar.CurrDatabase)

            RefeshUserApplicableLocationsAndGLAccount()

            CreateAutoIndentAccordingReorderLevel()

            common.clsUserInfo.CurrentUserCode = objCommonVar.CurrentUserCode
            qry = "select 1 from sys.databases where name = '" + objCommonVar.CurrDatabase + "'"
            dt = clsDBFuncationality.GetDataTable(qry)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Company :" + cboCompany.Text + " is not mapped with any Database ")
            Else
                'CallCreateTableFunction()
                Dim obj As clsLoginInfo = New clsLoginInfo()
                obj.SaveData()
                'clsDBFuncationality.ExecuteNonQuery("update TSPL_CUSTOMER_MASTER set Customer_Class=Cust_Type_Code")
                LoadWorkingScreen()
            End If
            RefeshUserApplicableZonesAndCustomer()
            If clsCommon.CompairString(Pwd, clsCommon.EncryptString(txtUserName.Text)) = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow("You are using system generated password." + Environment.NewLine + "Please reset you password", Me.Text)
                Dim frm As New FrmChangePassword()
                frm.ShowDialog()
            End If

            'done by stuti on 06/12/2016 for notification and approval work
            Try
                If System.IO.File.Exists(Application.StartupPath & "\XpertAlertApp.exe") AndAlso Not System.IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Programs + "\startup\XpertAlertApp.Ink") Then
                    Dim szName As String = "XpertAlertApp"
                    Dim szTagetFile As String = Application.StartupPath & "\XpertAlertApp.exe"
                    Dim szWorkingDir As String = Application.StartupPath
                    Dim szCmdLine As String = "" '(Optional) If your application have cmd line
                    Dim szComment As String = "" '(Optional) Description of shortcut.
                    Dim szIcon As String = ""
                    Dim nIndex As Integer = 0 'Default index is 0
                    Dim WinStyle As ProcessWindowStyle = ProcessWindowStyle.Normal
                    Dim StartupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
                    Dim szTarget As String = StartupFolder
                    Dim szCreateDirForStartMenu As String = ""
                    'AddLnkShortcut(szName, szTagetFile, szWorkingDir, szCmdLine, szComment, szIcon, 0, WinStyle, szTarget, szCreateDirForStartMenu)
                End If
            Catch ex As Exception
            End Try
            If System.IO.File.Exists(Application.StartupPath & "\XpertAlertApp.exe") AndAlso Not IsProcessRunning("XpertAlertApp") AndAlso Not IsProcessRunning("RadForm1") Then
                Try
                    Process.Start(Application.StartupPath & "\XpertAlertApp.exe")
                Catch ex As Exception
                End Try
            End If
            '------------end here-------------------


        Else
            isLoginError = True
            clsCommon.MyMessageBoxShow(Me, "User Name or Password is not Correct.Please provide the correct login information.")
        End If
        Dim AllowAutoLockTransaction As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowAutoLockTransaction, clsFixedParameterCode.AllowAutoLockTransaction, Nothing))
        If AllowAutoLockTransaction = 1 Then
            AUTOLOCKTRANSACTION()
        End If

        qry = "select IsThirdPartyLocationOnERP from TSPL_INV_PARAMETERS"
        IsLoc_Third_Party = clsDBFuncationality.getSingleValue(qry)

        If IsLoc_Third_Party = "1" Then
            IsLoc_Third_Party = "YES"
        Else
            IsLoc_Third_Party = "NO"
        End If
        '-------------------------------------------------------------------------------

        clsScreenNotificationSchedule.ShowLoginNotifications(objCommonVar.CurrentUserCode)
        Timer1.Start()

        qry = "Select User_Code from TSPL_LOCATION_SETTING  where User_Code='" + objCommonVar.CurrentUserCode + "' "
        Dim usercode = clsDBFuncationality.getSingleValue(qry)
        If clsCommon.myLen(usercode) > 0 Then
            Dim frmLoc As New frmSelectLocation
            frmLoc.ShowDialog()
        End If

        '-------------------04/07/2014----------BM00000003039
        ReminderTimer.Interval = 100000
        ReminderTimer.Enabled = True
        RadDesktopAlert1.ButtonItems.Add(radbuttonelement)
        RadDesktopAlert1.ButtonItems.Add(radbuttonDontShow)
        AddHandler radbuttonelement.Click, AddressOf radbuttonelement_Click
        AddHandler radbuttonDontShow.Click, AddressOf DontShowAgain_Click

        ' GetPendingSaleOrder()
        'GetPendingSaleBooking()
        If clsCommon.myLen(objCommonVar.CurrentUserCode) > 0 Then
            GetMccFssaiPopUp()
            'Dim objsms As New FrmMccSMSSetting
            'objsms.SendMail("", clsCommon.GETSERVERDATE().AddDays(-1), "", clsUserMgtCode.frmMilkShiftEndMCC, "")
        End If
        '---------------end here
        '------For Application Idle State Checking
        Dim FILE_NAME As String = Application.StartupPath + "\Table.Txp"
        If Not System.IO.File.Exists(FILE_NAME) Then
            isIdle = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.Idle, clsFixedParameterCode.isIdleTimerOn, Nothing))
            IdleTimeinSeconds = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.Idle, clsFixedParameterCode.idleTime, Nothing)) * 60
            If isIdle = 1 Then
                If IdleTimeinSeconds > 0 Then
                    Timer3.Enabled = True
                End If
            End If
        End If

        '===================call approval alert screen
        Dim OpenWorkFlowInERP As Boolean = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.WorkApprovalFlowInERP, clsFixedParameterCode.WorkApprovalFlowInERP, Nothing) = "1", True, False))
        If OpenWorkFlowInERP Then
            Dim strqry As String = "select isnull(Read_Flag,0) from TSPL_GROUP_PROGRAM_MAPPING where Program_Code='APP-SUM-SCR' and Group_Code in (select Group_Code  from TSPL_USER_GROUP_MAPPING where User_Code='" + objCommonVar.CurrentUserCode + "')"
            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry)) > 0 OrElse clsCommon.CompairString("admin", objCommonVar.CurrentUserCode) = CompairStringResult.Equal Then
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllLevelApprovalIsMandatory, clsFixedParameterCode.AllLevelApprovalIsMandatory, Nothing)) > 0 Then
                    ''if all level approval mandatory then current user seen alert screen only when before user approved ore rejected the docu.
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum( case when tspl_approval_level_transaction_detail.No_Of_Level=1 then 1 when tspl_approval_level_transaction_detail.No_Of_Level<AppLvl.No_Of_Level and isnull(tspl_approval_level_transaction_detail.status,'')<>'' then 1 else 0 end) from tspl_approval_level_transaction_detail " &
                                                                           " left outer join (select no_of_level,User_Code,trans_code,document_code,status from tspl_approval_level_transaction_detail where user_code='" + objCommonVar.CurrentUserCode + "')AppLvl on AppLvl.TRANS_Code=tspl_approval_level_transaction_detail.TRANS_Code and AppLvl.Document_Code=tspl_approval_level_transaction_detail.Document_Code ")) > 0 Then
                        Dim frm As New FrmApprovalAlertSumm()
                        frm.Text = "Approval Alert"
                        'frm.MdiParent = Me
                        frm.Show()
                    End If
                Else
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(trans_code) from tspl_approval_level_transaction_detail where user_code='" + objCommonVar.CurrentUserCode + "' ")) > 0 Then
                        Dim frm As New FrmApprovalAlertSumm()
                        frm.Text = "Approval Alert"
                        'frm.MdiParent = Me
                        frm.Show()
                    End If
                End If

            End If
        End If

        '-- KDIL's Work < Load pending Documents from whole ERP > < Preeti Gupta > < Kunal > <Setting Based Work >
        Dim OnOff As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPromptPendingDocs, clsFixedParameterType.AllowPromptPendingDocs, Nothing))
        If OnOff = 1 Then
            Dim promptPendingDocument As Integer = 0
            promptPendingDocument = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PromptMsgForPendingDocIntervel, clsFixedParameterType.PromptMsgForPendingDocIntervel, Nothing))
            Dim dt1 As DataTable = Nothing
            Dim ChkDate As Date? = Nothing
            Dim ChkOpenFromDate As Date? = Nothing
            Dim ChkOpenFromDateNew As Date? = Nothing
            Dim ChkOpenToDate As Date? = Nothing
            Dim Strqry As String = "select  Open_From_Date,Open_To_Date,DATEADD (day," & promptPendingDocument & ", Open_To_Date ) as ChkDate,DATEADD (day,1,Open_To_Date )  as ChkOpenFromDateNew  from tspl_user_master where User_Code ='" + objCommonVar.CurrentUserCode + "' "
            dt1 = clsDBFuncationality.GetDataTable(Strqry)
            Dim lastSubmitDate As Date = clsDBFuncationality.getSingleValue("select coalesce(max(created_date),'') created_date from TSPL_PROMPT_MSG_PENDING_HEAD where 1=1 and Created_By = '" + objCommonVar.CurrentUserCode + "'")
            If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                Dim frmPrmpt As New FrmPromptMsgRelatedToPendency()
                If clsCommon.myCDate(lastSubmitDate, "dd/MMM/yyyy") < clsCommon.myCDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy") Then
                    frmPrmpt.Text = "Pending Documents List"
                    frmPrmpt.ChkOpenFromDate = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE().AddDays(-promptPendingDocument), "yyyy-MM-dd")
                    frmPrmpt.ChkOpenToDate = clsCommon.GETSERVERDATE()
                    frmPrmpt.ShowDialog()
                End If
            End If
        End If
        SettingHighSecurityOnWeighingIntegratedScreen = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AddHighSecurityOnWeighingIntegratedScreen, clsFixedParameterCode.AddHighSecurityOnWeighingIntegratedScreen, Nothing)) = 1, True, False)
        SingleUserParticularDairyBookingEdit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SingleUserParticularDairyBookingEdit, clsFixedParameterCode.SingleUserParticularDairyBookingEdit, Nothing)) = 1, True, False)

        If ShowNotificationWithoutSMSAPP = True Then
            Timer5.Interval = SetNotificationRefreshTimeInMinutes * 60 * 1000
            Timer5.Start()
            Dim PendingNotification As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(*) as Notification from TSPL_NOTIFICATION_DETAIL LEFT OUTER JOIN TSPL_NOTIFICATION_HEAD ON TSPL_NOTIFICATION_HEAD.Code = TSPL_NOTIFICATION_DETAIL.Code LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  =  TSPL_NOTIFICATION_DETAIL.User_Name   LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  = TSPL_USER_MASTER.EmployeeCode where TSPL_NOTIFICATION_DETAIL.Sender_Replay=0 and TSPL_USER_MASTER.user_code='" + objCommonVar.CurrentUserCode + "'"))
            If PendingNotification > 0 Then
                Dim frmPromptMsgNotification As New frmPromptMsgNotification()
                frmPromptMsgNotification.Location = New Point(Screen.PrimaryScreen.WorkingArea.Width - frmPromptMsgNotification.Width, Screen.PrimaryScreen.WorkingArea.Height - frmPromptMsgNotification.Height)
                frmPromptMsgNotification.Show()
            End If
        End If
        '==========================================
        'Dim strZoneCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Zone_Code from TSPL_USER_MASTER where User_Code = '" + objCommonVar.CurrentUserCode + "'"))
        'If clsCommon.myLen(strZoneCode) > 0 AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "TSDDCF") = CompairStringResult.Equal Then
        '    If clsCommon.MyMessageBoxShow("Currently  you are joined " + strZoneCode + " Zone.Do you want to Continue with " + strZoneCode + " Zone?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
        '        clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.userMaster, txtUserName.Text)
        '    End If
        'End If
        'frmDefaultUserZone1
        Dim isUserDefaulZoneValid As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select count ( distinct TSPL_USER_CUSTOMER_ZONE.Zone_code ) as aa  from TSPL_USER_CUSTOMER_ZONE   left outer join TSPL_USER_CUSTOMER_CATEGORY on  TSPL_USER_CUSTOMER_CATEGORY.User_Code = TSPL_USER_CUSTOMER_ZONE.User_Code  where TSPL_USER_CUSTOMER_ZONE.USER_Code = '" + objCommonVar.CurrentUserCode + "'  and TSPL_USER_CUSTOMER_CATEGORY.Customer_Category in ('Vendor','Institution CR','Institution SO','Distributor')"))
        If isUserDefaulZoneValid = True AndAlso isLoginError = False Then
            Dim frmDefaultUserZone1 As New frmDefaultUserZone1()
            frmDefaultUserZone1.Location = New Point(Screen.PrimaryScreen.WorkingArea.Width - frmDefaultUserZone1.Width, Screen.PrimaryScreen.WorkingArea.Height - frmDefaultUserZone1.Height)
            frmDefaultUserZone1.ShowDialog()
        End If
        Dim folder1 As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
        Dim folder2 As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Temps"
        Dim folder3 As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Uplpads"
        Dim userTime1 As DateTime = DateTime.Now ' 
        DeleteOldDownloadsFiles(folder1, userTime1)
        DeleteOldTempsFiles(folder2, userTime1)
        DeleteOldUploadsFiles(folder3, userTime1)
        '==========================================
    End Sub
    Sub DeleteOldDownloadsFiles(folderPath As String, referenceTime As DateTime)
        Try
            Dim files = Directory.GetFiles(folderPath)

            For Each filePath As String In files
                Dim creationTime As DateTime = File.GetCreationTime(filePath)

                If (referenceTime - creationTime).TotalHours >= 24 Then
                    Try
                        File.Delete(filePath)
                    Catch ex As Exception
                    End Try
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Sub DeleteOldTempsFiles(folderPath As String, referenceTime As DateTime)
        Try
            Dim files = Directory.GetFiles(folderPath)
            For Each filePath As String In files
                Dim creationTime As DateTime = File.GetCreationTime(filePath)
                If (referenceTime - creationTime).TotalHours >= 24 Then
                    Try
                        File.Delete(filePath)
                    Catch ex As Exception
                    End Try
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Sub DeleteOldUploadsFiles(folderPath As String, referenceTime As DateTime)
        Try
            Dim files = Directory.GetFiles(folderPath)
            For Each filePath As String In files
                Dim creationTime As DateTime = File.GetCreationTime(filePath)
                If (referenceTime - creationTime).TotalHours >= 24 Then
                    Try
                        File.Delete(filePath)
                    Catch ex As Exception
                    End Try
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Function LastDayOfMonth(aDate As DateTime) As Date
        Return New DateTime(aDate.Year, aDate.Month, DateTime.DaysInMonth(aDate.Year, aDate.Month))
    End Function
    Private Function LastDayOfPreviousMonth(aDate As DateTime) As Date
        Return New DateTime(aDate.Year, aDate.Month - 1, DateTime.DaysInMonth(aDate.Year, aDate.Month - 1))
    End Function
    Sub ShowUnPostedDocument()
        Try
            Dim qry As String = ""
            Dim currentDate As Date = clsCommon.GETSERVERDATE()
            Dim datLastDay As Date = LastDayOfMonth(currentDate)
            Dim AllowAutoLockTransaction As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowAutoLockTransaction, clsFixedParameterCode.AllowAutoLockTransaction, Nothing))
            Dim PromptTimeToPostTransactions As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PromptTimeToPostTransactions, clsFixedParameterCode.PromptTimeToPostTransactions, Nothing))

            If AllowAutoLockTransaction = 1 AndAlso PromptTimeToPostTransactions > 0 Then
                Dim PromptDateToPostTransactions As Date = datLastDay.AddDays(-PromptTimeToPostTransactions)
                If clsCommon.myCDate(currentDate) >= clsCommon.myCDate(PromptDateToPostTransactions) Then

                    qry = "select * from ( " &
                               "select 'Common Module' as Module,'Bank Reverse' as TransactionName,isnull( ( select ' '+tspl_bank_reverse.Document_No+' ,  '    from tspl_bank_reverse where isnull(Post,'N')='N' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'Bank Transfer' as TransactionName, isnull(( select ' '+TSPL_BANK_TRANSFER.Transfer_No +' ,  '    from TSPL_BANK_TRANSFER where isnull(Post,'N')='N' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'Bank Reco' as TransactionName, isnull(( select ' '+tspl_BankReco_Head.Reconciliation_Id +' ,  '    from tspl_BankReco_Head where isnull(Post,'N')='N' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'CForm Header' as TransactionName, isnull(( select ' '+TSPL_CForm_HEADER.Document_No +' ,  '    from TSPL_CForm_HEADER where isnull(Posted,'N')='N' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'Bank Gauantee' as TransactionName, isnull(( select ' '+tspl_bank_guarantee_master.DocNo +' ,  '    from tspl_bank_guarantee_master where isnull(status,'N')='N' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'Bank Opening Reco' as TransactionName, isnull(( select ' '+TSPL_BANK_OPENING_RECO.Code +' ,  '    from TSPL_BANK_OPENING_RECO where isnull(status,'0')='0' and Created_By=''  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Common Module' as Module,'Revaluation Entry' as TransactionName, isnull(( select ' '+TSPL_REVALUATION_HEAD.Document_No +' ,  '    from TSPL_REVALUATION_HEAD where isnull(status,'0')='0' and Created_By=''  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Receivables' as Module,'Receipt Entry' as TransactionName, isnull(( select ' '+TSPL_RECEIPT_HEADER.Receipt_No +' ,  '    from TSPL_RECEIPT_HEADER where isnull(Posted,'N')='N' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Receivables' as Module,'Receipt Adjustment' as TransactionName, isnull(( select ' '+TSPL_Receipt_Adjustment_Header.Adjustment_No +' ,  '    from TSPL_Receipt_Adjustment_Header where isnull(is_post,'N')='N' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Receivables' as Module,'AR Invoice Entry' as TransactionName, isnull(( select ' '+TSPL_Customer_Invoice_Head.Document_No +' ,  '    from TSPL_Customer_Invoice_Head where isnull(Status,0)='0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payable' as Module,'Payment Entry' as TransactionName, isnull(( select ' '+TSPL_PAYMENT_HEADER.Document_No +' ,  '    from TSPL_PAYMENT_HEADER where isnull(Posted,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payable' as Module,'AP invoice Entry' as TransactionName, isnull(( select ' '+TSPL_VENDOR_INVOICE_HEAD.Document_No +' ,  '    from TSPL_VENDOR_INVOICE_HEAD where isnull(Posting_Date,'')='' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payable' as Module,'Payment Adjusment Entry' as TransactionName, isnull(( select ' '+TSPL_Payment_Adjustment_Header.Adjustment_No +' ,  '    from TSPL_Payment_Adjustment_Header where isnull(is_post,'N')='N' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                                "select 'Payable' as Module,'Supplier Registration' as TransactionName, isnull(( select ' '+TSPL_SUPPLIER_REGISTRATION.Registration_No +' ,  '    from TSPL_SUPPLIER_REGISTRATION where isnull(Posted,'0')='0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                                " union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Purchase Indent' as TransactionName, isnull(( select ' '+TSPL_REQUISITION_HEAD.Requisition_Id +' ,  '    from TSPL_REQUISITION_HEAD where isnull(Status,'0')='0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'RFQ Detail' as TransactionName, isnull(( select ' '+TSPL_RFQ_HEAD.RFQ_NO +' ,  '    from TSPL_RFQ_HEAD where isnull(Is_Post,'0')='0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Vendor Quotation' as TransactionName, isnull(( select ' '+TSPL_VENDOR_QUOTATION_HEAD.Quotation_No +' ,  '    from TSPL_VENDOR_QUOTATION_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Purchase Order' as TransactionName, isnull(( select ' '+TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No +' ,  '    from TSPL_PURCHASE_ORDER_HEAD where isnull(Status,0)=0 and Created_By='" & objCommonVar.CurrentUserCode & "'  and MT_Is_Merchant_Trade=0 for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Gate Receipt Note' as TransactionName, isnull(( select ' '+TSPL_GRN_HEAD.GRNo +' ,  '    from TSPL_GRN_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'PO Weighment' as TransactionName, isnull(( select ' '+TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code +' ,  '    from TSPL_PO_WEIGHTMENT_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               " and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'MRN Head' as TransactionName, isnull(( select ' '+TSPL_MRN_HEAD.MRN_No +' ,  '    from TSPL_MRN_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'SRN Head' as TransactionName, isnull(( select ' '+TSPL_SRN_HEAD.SRN_No +' ,  '    from TSPL_SRN_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'   for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Purchase Invoice' as TransactionName, isnull(( select ' '+TSPL_PI_HEAD.PI_No +' ,  '    from TSPL_PI_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Purchase Return' as TransactionName, isnull(( select ' '+TSPL_PR_HEAD.PR_No +' ,  '    from TSPL_PR_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'NRGP Request' as TransactionName, isnull(( select ' '+TSPL_NRGP_REQUEST_HEAD.BOOKING_NO +' ,  '    from TSPL_NRGP_REQUEST_HEAD where isnull(Posted,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'RGP/NRGP' as TransactionName, isnull(( select ' '+TSPL_RGP_head.rgp_no +' ,  '    from TSPL_RGP_head where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Purchase' as Module,'Issue/Return/Transfer' as TransactionName, isnull(( select ' '+TSPL_IssueReturn_HEAD.Doc_No +' ,  '    from TSPL_IssueReturn_HEAD where isnull(Status,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'TDS Deduction' as Module,'TDS Payment' as TransactionName, isnull(( select ' '+TSPL_TDS_PAYMENT_HEADER.Document_No +' ,  '    from TSPL_TDS_PAYMENT_HEADER where isnull(posted,0)=0 " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               " union all " + Environment.NewLine &
                               "select 'General Ledger' as Module,'TDS Payment' as TransactionName, isnull(( select ' '+TSPL_JOURNAL_MASTER.Voucher_No +' ,  '    from TSPL_JOURNAL_MASTER where isnull(Authorized,'N') = 'N' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'General Ledger' as Module,'VCGL Entry' as TransactionName, isnull(( select ' '+TSPL_VCGL_Head.Document_No +' ,  '    from TSPL_VCGL_Head where isnull(Status,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Material Management' as Module,'Transfer' as TransactionName, isnull(( select ' '+TSPL_TRANSFER_ORDER_HEAD.Document_No +' ,  '    from TSPL_TRANSFER_ORDER_HEAD where isnull(Status,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Material Management' as Module,'Store/Production/Empty Adjustment' as TransactionName, isnull(( select ' '+TSPL_adjustment_header.adjustment_no +' ,  '    from TSPL_adjustment_header where isnull(posted,'N') = 'N' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fixed Assets' as Module,'Acquisition Entry' as TransactionName, isnull(( select ' '+TSPL_ACQUISITION_HEAD.Acquisition_Code +' ,  '    from TSPL_ACQUISITION_HEAD where isnull(Status,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fixed Assets' as Module,'Disposal Entry' as TransactionName, isnull(( select ' '+TSPL_ASSET_SCRAP_HEAD.Document_No +' ,  '    from TSPL_ASSET_SCRAP_HEAD where isnull(Status,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fixed Assets' as Module,'Asset Work Expense' as TransactionName, isnull(( select ' '+TSPL_ASSET_WORK_HEAD.Document_Code +' ,  '    from TSPL_ASSET_WORK_HEAD where isnull(Status,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Allowance Details' as TransactionName, isnull(( select ' '+TSPL_ALLOWANCE.ALLOWANCE_CODE +' ,  '    from TSPL_ALLOWANCE where isnull(Posted,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Deduction Detail' as TransactionName, isnull(( select ' '+TSPL_DEDUCTION.DEDUCTION_CODE +' ,  '    from TSPL_DEDUCTION where isnull(Posted,'0') = '0' " + Environment.NewLine &
                               "and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Loan Application' as TransactionName, isnull(( select ' '+TSPL_LOAN_APPLICATION.LOAN_CODE +' ,  '    from TSPL_LOAN_APPLICATION where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Generate Bonus' as TransactionName, isnull(( select ' '+TSPL_EMPLOYEE_BONUS.EMP_BONUS_CODE +' ,  '    from TSPL_EMPLOYEE_BONUS where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Loan Generation' as TransactionName, isnull(( select ' '+TSPL_LOAN_GENERATION.LOAN_GENERATION_CODE +' ,  '    from TSPL_LOAN_GENERATION where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Daily Attendance' as TransactionName, isnull(( select ' '+TSPL_DAILY_ATTENDANCE.DLA_CODE +' ,  '    from TSPL_DAILY_ATTENDANCE where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Monthly Attendance' as TransactionName, isnull(( select ' '+TSPL_MONTHLY_ATTENDANCE.MTA_CODE +' ,  '    from TSPL_MONTHLY_ATTENDANCE where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Salary Grneration' as TransactionName, isnull(( select ' '+TSPL_GENERATE_SALARY.SALARY_GENERATION_CODE +' ,  '    from TSPL_GENERATE_SALARY where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Payroll' as Module,'Employee Increment' as TransactionName, isnull(( select ' '+TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_CODE +' ,  '    from TSPL_EMPLOYEE_INCREMENT_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Gate Entry In' as TransactionName, isnull(( select ' '+TSPL_MILK_GATE_ENTRY_IN.Entry_Code +' ,  '    from TSPL_MILK_GATE_ENTRY_IN where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Gate Entry Weighment' as TransactionName, isnull(( select ' '+TSPL_MILK_GATE_ENTRY_WEIGHTMENT.Weighment_Code +' ,  '    from TSPL_MILK_GATE_ENTRY_WEIGHTMENT where isnull(GW_Status,'0') = '0'and GW_Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Receipt' as TransactionName, isnull(( select ' '+TSPL_MILK_RECEIPT_HEAD.DOC_CODE +' ,  '    from TSPL_MILK_RECEIPT_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Sample' as TransactionName, isnull(( select ' '+TSPL_MILK_SAMPLE_HEAD.DOC_CODE +' ,  '    from TSPL_MILK_SAMPLE_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk SRN' as TransactionName, isnull(( select ' '+TSPL_MILK_SRN_HEAD.DOC_CODE +' ,  '    from TSPL_MILK_SRN_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Truck Sheet' as TransactionName, isnull(( select ' '+tspl_milk_truck_sheet_Head.DOC_CODE +' ,  '    from tspl_milk_truck_sheet_Head where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Shift End' as TransactionName, isnull(( select ' '+TSPL_MILK_Shift_End_HEAD.DOC_CODE +' ,  '    from TSPL_MILK_Shift_End_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Tanker Dispatch' as TransactionName, isnull(( select ' '+TSPL_MCC_Dispatch_Challan.Chalan_NO +' ,  '    from TSPL_MCC_Dispatch_Challan where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Tanker Location Charge' as TransactionName, isnull(( select ' '+tspl_MCC_dispatch_transfer.Doc_No +' ,  '    from tspl_MCC_dispatch_transfer where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Purchase Invoice Head' as TransactionName, isnull(( select ' '+TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE +' ,  '    from TSPL_MILK_PURCHASE_INVOICE_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'VSP Asset Issue' as TransactionName, isnull(( select ' '+TSPL_VSPAsset_HEAD.Doc_No +' ,  '    from TSPL_VSPAsset_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'MCC Material Sale' as TransactionName, isnull(( select ' '+TSPL_SD_SHIPMENT_HEAD.Document_Code +' ,  '    from TSPL_SD_SHIPMENT_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'MCC Material Sale Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'VSP Item Issue' as TransactionName, isnull(( select ' '+TSPL_VSPItem_HEAD.Doc_No +' ,  '    from TSPL_VSPItem_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Payment Process' as TransactionName, isnull(( select ' '+TSPL_PAYMENT_PROCESS_HEAD.Doc_No +' ,  '    from TSPL_PAYMENT_PROCESS_HEAD where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'Milk Recurring Payable Invoice' as TransactionName, isnull(( select ' '+TSPL_Recurring_Payable_INVOICE_Head.Document_No +' ,  '    from TSPL_Recurring_Payable_INVOICE_Head where isnull(Posting_Date,'') = '' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Procurement MCC' as Module,'MCC tanker Dispatch Return' as TransactionName, isnull(( select ' '+TSPL_MCC_Tanker_Dispatch_Return_head.Return_NO +' ,  '    from TSPL_MCC_Tanker_Dispatch_Return_head where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Gate Entry' as TransactionName, isnull(( select ' '+Tspl_Gate_Entry_Details.Gate_Entry_No +' ,  '    from Tspl_Gate_Entry_Details where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Weighment' as TransactionName, isnull(( select ' '+TSPL_Weighment_Detail.Weighment_No +' ,  '    from TSPL_Weighment_Detail where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Quality Check' as TransactionName, isnull(( select ' '+TSPL_QUALITY_CHECK.QC_No +' ,  '    from TSPL_QUALITY_CHECK where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Unloading' as TransactionName, isnull(( select ' '+TSPL_MILK_UNLOADING.Unloading_No +' ,  '    from TSPL_MILK_UNLOADING where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Unloading' as TransactionName, isnull(( select ' '+TSPL_MILK_UNLOADING.Unloading_No +' ,  '    from TSPL_MILK_UNLOADING where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Cleaning' as TransactionName, isnull(( select ' '+TSPL_Cleaning.Doc_No +' ,  '    from TSPL_Cleaning where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Bulk Milk SRN' as TransactionName, isnull(( select ' '+TSPL_Bulk_MILK_SRN.SRN_NO +' ,  '    from TSPL_Bulk_MILK_SRN where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Bulk Milk Purchase Invoice' as TransactionName, isnull(( select ' '+tspl_Bulk_milk_purchase_Invoice_head.DOC_NO +' ,  '    from tspl_Bulk_milk_purchase_Invoice_head where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Milk Transfer In' as TransactionName, isnull(( select ' '+TSPL_MILK_TRANSFER_IN.Receipt_Challan_No +' ,  '    from TSPL_MILK_TRANSFER_IN where isnull(isPosted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Bulk Procurement' as Module,'Provision Entry' as TransactionName, isnull(( select ' '+TSPL_PROVISION_ENTRY.Doc_No +' ,  '    from TSPL_PROVISION_ENTRY where isnull(isPosted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Gate Entry' as TransactionName, isnull(( select ' '+TSPL_GATEENTRY_SALE.Document_No +' ,  '    from TSPL_GATEENTRY_SALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Weighment' as TransactionName, isnull(( select ' '+TSPL_WEIGHMENT_DETAIL_BULKSALE.Weighment_No +' ,  '    from TSPL_WEIGHMENT_DETAIL_BULKSALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'LoadIn Tanker Detais' as TransactionName, isnull(( select ' '+TSPL_LOADING_TANKER_DETAIL_BULKSALE.LoadingTanker_No +' ,  '    from TSPL_LOADING_TANKER_DETAIL_BULKSALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Fat/SNF Check  /  QC Details' as TransactionName, isnull(( select ' '+TSPL_QUALITY_CHECK_BULKSALE.QC_No +' ,  '    from TSPL_QUALITY_CHECK_BULKSALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Bulk Dispatch' as TransactionName, isnull(( select ' '+TSPL_Dispatch_BulkSale.Document_No +' ,  '    from TSPL_Dispatch_BulkSale where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Bulk Invoice' as TransactionName, isnull(( select ' '+TSPL_INVOICE_MASTER_BULKSALE.Document_No +' ,  '    from TSPL_INVOICE_MASTER_BULKSALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Bulk Dispatch Trade' as TransactionName, isnull(( select ' '+TSPL_Dispatch_BulkSale_Trade.Document_No +' ,  '    from TSPL_Dispatch_BulkSale_Trade where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Bulk Sale' as Module,'Bulk Sale Return' as TransactionName, isnull(( select ' '+TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No +' ,  '    from TSPL_SALE_RETURN_MASTER_BULKSALE where isnull(Posted,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'CSA Sale' as Module,'CSA Delivery Order' as TransactionName, isnull(( select ' '+TSPL_CSA_DO_HEAD.Doc_No +' ,  '    from TSPL_CSA_DO_HEAD where isnull(Is_Post,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'CSA Sale' as Module,'CSA Transfer' as TransactionName, isnull(( select ' '+TSPL_CSA_transfer_head.DOC_CODE +' ,  '    from TSPL_CSA_transfer_head where isnull(Status,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'CSA Sale' as Module,'Sale Patti' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='CSA' for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'CSA Sale' as Module,'CSA Transfer Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='CSA'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'CSA Sale' as Module,'CSA Sale Patti Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0' and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='CPR'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh  Booking' as TransactionName, isnull(( select ' '+TSPL_BOOKING_MATSER.Document_No +' ,  '    from TSPL_BOOKING_MATSER where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh  Delivery Order' as TransactionName, isnull(( select ' '+TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No +' ,  '    from TSPL_DELIVERY_NOTE_MASTER_FRESHSALE where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh  Dispatch' as TransactionName, isnull(( select ' '+TSPL_SD_SHIPMENT_HEAD.Document_Code +' ,  '    from TSPL_SD_SHIPMENT_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SHIPMENT_HEAD.Trans_Type='FS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh Sale Invoice' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='FS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh Sale Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='FS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Fresh Sale' as Module,'Fresh Crate Received' as TransactionName, isnull(( select ' '+TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No +' ,  '    from TSPL_CRATE_RECEIVED_HEAD_FRESHSALE where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Product  Booking' as TransactionName, isnull(( select ' '+TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code +' ,  '    from TSPL_BOOKING_MASTER_PRODUCTSALE where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Product  Delivery Order' as TransactionName, isnull(( select ' '+TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code +' ,  '    from TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Product  Dispatch' as TransactionName, isnull(( select ' '+TSPL_SD_SHIPMENT_HEAD.Document_Code +' ,  '    from TSPL_SD_SHIPMENT_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SHIPMENT_HEAD.Trans_Type='PS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Product Sale Invoice' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='PS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Product Sale Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='PS'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Product Sale' as Module,'Sale Order' as TransactionName, isnull(( select ' '+TSPL_SD_SALES_ORDER_HEAD.Document_Code +' ,  '    from TSPL_SD_SALES_ORDER_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALES_ORDER_HEAD.Trans_Type='PS'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Sale Quotaion' as TransactionName, isnull(( select ' '+TSPL_SD_QUOTATION_HEAD.Document_Code +' ,  '    from TSPL_SD_QUOTATION_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_QUOTATION_HEAD.Trans_Type='EXP'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Export Sale Order' as TransactionName, isnull(( select ' '+TSPL_SD_SALES_ORDER_HEAD.Document_Code +' ,  '    from TSPL_SD_SALES_ORDER_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALES_ORDER_HEAD.Trans_Type='EXP'  and TSPL_SD_SALES_ORDER_HEAD.salesorder_type='EX'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Export  Performa Invoice' as TransactionName, isnull(( select ' '+TSPL_EX_PI_HEAD.Document_Code +' ,  '    from TSPL_EX_PI_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_EX_PI_HEAD.document_type='EX'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Export Commercial Invoice' as TransactionName, isnull(( select ' '+TSPL_EX_COMMERCIAL_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_EX_COMMERCIAL_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and document_type='EX'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Export Sale Invoice' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type='EX'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Export Sale' as Module,'Export Sale Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_RETURN_HEAD.Document_Code='EX'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Trade' as Module,'Merchant Purchase Order' as TransactionName, isnull(( select ' '+TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No +' ,  '    from TSPL_PURCHASE_ORDER_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and  TSPL_PURCHASE_ORDER_HEAD.MT_Is_Merchant_Trade=1  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'Merchant Sale Order' as TransactionName, isnull(( select ' '+TSPL_SD_SALES_ORDER_HEAD.Document_Code +' ,  '    from TSPL_SD_SALES_ORDER_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALES_ORDER_HEAD.Trans_Type='EXP'  and TSPL_SD_SALES_ORDER_HEAD.salesorder_type='MT'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'Merchant  Performa Invoice' as TransactionName, isnull(( select ' '+TSPL_EX_PI_HEAD.Document_Code +' ,  '    from TSPL_EX_PI_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_EX_PI_HEAD.document_type='MT'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'LC Request' as TransactionName, isnull(( select ' '+TSPL_LC_REQUEST_MT.LCRequestNo +' ,  '    from TSPL_LC_REQUEST_MT where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'LC Creation' as TransactionName, isnull(( select ' '+TSPL_LC_CREATION_MT.LCCreationNo +' ,  '    from TSPL_LC_CREATION_MT where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'Document Acceptance' as TransactionName, isnull(( select ' '+TSPL_DOCUMENT_ACCEPTANCE_MT.DocumentAcceptanceNo +' ,  '    from TSPL_DOCUMENT_ACCEPTANCE_MT where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'   for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'Merchant Sale Invoice' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_INVOICE_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_INVOICE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_INVOICE_HEAD.Document_Type='MT'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Merchant Sale' as Module,'Merchant Sale Return' as TransactionName, isnull(( select ' '+TSPL_SD_SALE_RETURN_HEAD.Document_Code +' ,  '    from TSPL_SD_SALE_RETURN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "' and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='EXP' and TSPL_SD_SALE_RETURN_HEAD.Document_Code='MT'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Planning' as TransactionName, isnull(( select ' '+TSPL_PP_PRODUCTION_PLAN_HEAD.Plan_Code +' ,  '    from TSPL_PP_PRODUCTION_PLAN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Planning' as TransactionName, isnull(( select ' '+TSPL_PP_PRODUCTION_PLAN_HEAD.Plan_Code +' ,  '    from TSPL_PP_PRODUCTION_PLAN_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Batch Order' as TransactionName, isnull(( select ' '+TSPL_PP_BATCH_ORDER_HEAD.Plan_Code +' ,  '    from TSPL_PP_BATCH_ORDER_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Issue Entry' as TransactionName, isnull(( select ' '+TSPL_PP_ISSUE_HEAD.Issue_Code +' ,  '    from TSPL_PP_ISSUE_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Standardization' as TransactionName, isnull(( select ' '+TSPL_PP_STANDARDIZATION_HEAD.Standardization_Code +' ,  '    from TSPL_PP_STANDARDIZATION_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Stage Process' as TransactionName, isnull(( select ' '+TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_CODE +' ,  '    from TSPL_PP_STAGE_PROCESS_HEAD where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Production Entry' as TransactionName, isnull(( select ' '+TSPL_PP_PRODUCTION_ENTRY.PROD_ENTRY_CODE +' ,  '    from TSPL_PP_PRODUCTION_ENTRY where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No] " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'Assemblies' as TransactionName, isnull(( select ' '+TSPL_PROD_ASSEMBLIES.CODE +' ,  '    from TSPL_PROD_ASSEMBLIES where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Production' as Module,'WRECKAGE ENTRY' as TransactionName, isnull(( select ' '+TSPL_WRECKAGE_ENTRY.WRECKAGE_ENTRY_CODE +' ,  '    from TSPL_WRECKAGE_ENTRY where isnull(Posted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Milk RGP' as TransactionName, isnull(( select ' '+TSPL_Milk_RGP_HEAD.RGP_No +' ,  '    from TSPL_Milk_RGP_HEAD where isnull(Status,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Milk Gate Entry' as TransactionName, isnull(( select ' '+TSPL_MILK_GATE_ENTRY_DETAILS.Gate_Entry_No +' ,  '    from TSPL_MILK_GATE_ENTRY_DETAILS where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Weighment Detail' as TransactionName, isnull(( select ' '+tspl_Milk_weighment_detail.Weighment_No +' ,  '    from tspl_Milk_weighment_detail where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Quality Check' as TransactionName, isnull(( select ' '+tspl_Milk_quality_check.QC_No +' ,  '    from tspl_Milk_quality_check where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Unloading' as TransactionName, isnull(( select ' '+TSPL_JOB_MILK_UNLOADING.Unloading_No +' ,  '    from TSPL_JOB_MILK_UNLOADING where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  " + Environment.NewLine &
                               "union all " + Environment.NewLine &
                               "select 'Milk Job Work' as Module,'Milk SRN' as TransactionName, isnull(( select ' '+tspl_Job_milk_srn.SRN_NO +' ,  '    from tspl_Job_milk_srn where isnull(isPosted,'0') = '0'and Created_By='" & objCommonVar.CurrentUserCode & "'  for xml path('')  ),'') as [Document No]  )  aa  where [Document No] <> ''"
                    Dim msg As String = ""
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        Dim strModule As String = ""
                        Dim strTrans As String = ""
                        Dim strDocument As String = ""
                        For Each dr As DataRow In dt.Rows
                            strModule = clsCommon.myCstr(dr("Module"))
                            strTrans = clsCommon.myCstr(dr("TransactionName"))
                            strDocument = clsCommon.myCstr(dr("Document No"))
                            msg += "Module - " + strModule + "  Transaction - " + strTrans + "  Document - " + clsCommon.myCstr(strDocument) + Environment.NewLine
                        Next
                        clsCommon.MyMessageBoxShow(msg)
                    End If
                End If
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message)
        End Try
    End Sub
    '' TO Auto Lock All Transaction location and location segment wise
    Function AUTOLOCKTRANSACTION() As Boolean
        Dim qry As String = ""
        Dim currentDate As Date = clsCommon.GETSERVERDATE()
        Dim dblLockdays As Double = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DaysToStartAutoLock, clsFixedParameterCode.DaysToStartAutoLock, Nothing))
        Dim datLastDay As Date = currentDate.AddDays(-dblLockdays)
        'Dim datLastDay As Date = LastDayOfPreviousMonth(currentDate)
        Dim intCount As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_LOCK_LOCATION where End_Date='" & clsCommon.GetPrintDate(datLastDay, "dd/MMM/yyyy") & "'"))

        'If intCount <= 0 Then
        '    Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()

        '    Try
        '        Dim ArrLoc As New ArrayList

        '        '' Location Segement wise
        '        clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_LOCK_LOCATION_SEGMENT ", trans)
        '        clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_LOCK_LOCATION_SEGMENT_USER ", trans)
        '        qry = " Select Segment_code as Code, Description from TSPL_GL_SEGMENT_CODE where Seg_No=7 "
        '        Dim dtLoc As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        '        Dim FrmLockTransaction As New FrmLockTransaction1
        '        qry = " Select * from (" + FrmLockTransaction.LockTransactionNameLocationSegwise() + ") xxx order by Module, [Transaction]"
        '        Dim dtTrans As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        '        If (dtLoc IsNot Nothing AndAlso dtLoc.Rows.Count > 0) Then
        '            For Each dr As DataRow In dtLoc.Rows
        '                Dim arr As New List(Of clsLockTransactionLocationSegmentwise)
        '                Dim obj As New clsLockTransactionLocationSegmentwise
        '                For Each dr1 As DataRow In dtTrans.Rows
        '                    obj = New clsLockTransactionLocationSegmentwise
        '                    obj.Location_Segment_Code = clsCommon.myCstr(dr("Code"))
        '                    obj.Module_Name = clsCommon.myCstr(dr1("Module"))
        '                    obj.Trans_Name = clsCommon.myCstr(dr1("Transaction"))
        '                    obj.Is_Locked = 1
        '                    obj.Start_Date = clsCommon.GetPrintDate("01-01-2001", "dd/MMM/yyyy")
        '                    obj.End_Date = datLastDay
        '                    arr.Add(obj)
        '                Next
        '                clsLockTransactionLocationSegmentwise.SaveData(obj.Location_Segment_Code, "", arr, trans)
        '            Next
        '        End If

        '        '' Location wise
        '        clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_LOCK_LOCATION ", trans)
        '        clsDBFuncationality.ExecuteNonQuery("Delete from TSPL_LOCK_LOCATION_USER ", trans)
        '        qry = " Select Location_Code as Code from TSPL_LOCATION_MASTER Where Location_Type='Physical' "
        '        Dim dtLocSeg As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        '        qry = " Select * from (" + FrmLockTransaction.LockTransactionNameLocationwise() + ") xxx order by Module, [Transaction]"
        '        dtTrans = clsDBFuncationality.GetDataTable(qry, trans)

        '        If (dtLocSeg IsNot Nothing AndAlso dtLocSeg.Rows.Count > 0) Then

        '            For Each dr As DataRow In dtLocSeg.Rows
        '                Dim obj As New clsLockTransactionLocationwise
        '                Dim arr1 As New List(Of clsLockTransactionLocationwise)
        '                For Each dr1 As DataRow In dtTrans.Rows
        '                    obj = New clsLockTransactionLocationwise
        '                    obj.Location_Code = clsCommon.myCstr(dr("Code"))
        '                    obj.Module_Name = clsCommon.myCstr(dr1("Module"))
        '                    obj.Trans_Name = clsCommon.myCstr(dr1("Transaction"))
        '                    obj.Is_Locked = 1
        '                    obj.Start_Date = clsCommon.GetPrintDate("01-01-2001", "dd/MMM/yyyy")
        '                    obj.End_Date = datLastDay
        '                    arr1.Add(obj)
        '                Next
        '                clsLockTransactionLocationwise.SaveData(obj.Location_Code, "", arr1, trans)
        '            Next
        '        End If
        '        trans.Commit()
        '        clsCommon.MyMessageBoxShow("Transaction Locked Successfully", Me.Text)
        '    Catch ex As Exception
        '        trans.Rollback()
        '        clsCommon.MyMessageBoxShow(Me, ex.Message)
        '    End Try
        'End If

    End Function
    Public Shared Sub CreateAutoIndentAccordingReorderLevel()
        Try
            Dim EnableMsgPopupforReorderLevel As Boolean = False
            Dim qry As String = Nothing
            Dim dt As New DataTable()
            Dim strlocation = Nothing
            EnableMsgPopupforReorderLevel = clsCommon.myCBool(clsDBFuncationality.getSingleValue("Select TSPL_PURCHASE_SETTINGS.ENABLE_POPUP_REORDERLEVEL from TSPL_PURCHASE_SETTINGS", Nothing))
            If EnableMsgPopupforReorderLevel Then
                strlocation = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_USER_MASTER.Default_Location from TSPL_USER_MASTER where TSPL_USER_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "'", Nothing))
                qry = "select z.ItemCode,z.ItemType,z.ItemDesc ,z.Qty,z.Unit  from (select coalesce(xx.ItemCode,TSPL_ITEM_REORDER_LEVEL_NEW.item_Code) as ItemCode,(((TSPL_ITEM_REORDER_LEVEL_NEW.Reorder_Qty)*uom1.Conversion_Factor)/TSPL_ITEM_UOM_DETAIL.Conversion_Factor-isnull(xxx.Qty,0)) as Qty,TSPL_ITEM_MASTER.Item_Type as ItemType," &
                        " TSPL_ITEM_MASTER.Item_Desc as ItemDesc ,TSPL_ITEM_UOM_DETAIL.UOM_Code as Unit from TSPL_ITEM_REORDER_LEVEL_NEW left outer join (select TSPL_INVENTORY_MOVEMENT.Item_Code as ItemCode," &
                        " (SUM(TSPL_INVENTORY_MOVEMENT.Stock_Qty * case when TSPL_INVENTORY_MOVEMENT.InOut ='I' then 1 else -1 end)) as Balance from TSPL_INVENTORY_MOVEMENT where TSPL_INVENTORY_MOVEMENT.Location_Code='" + strlocation + "'  group by TSPL_INVENTORY_MOVEMENT.Item_Code)xx on xx.ItemCode=TSPL_ITEM_REORDER_LEVEL_NEW.Item_Code" &
                        " left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_ITEM_REORDER_LEVEL_NEW.Item_Code" &
                        " left outer join (select TSPL_REQUISITION_DETAIL.Item_Code,SUM(TSPL_REQUISITION_DETAIL.Requisition_Qty) AS Qty from TSPL_REQUISITION_DETAIL " &
                        " left outer join TSPL_REQUISITION_HEAD ON TSPL_REQUISITION_HEAD.Requisition_Id=TSPL_REQUISITION_DETAIL.Requisition_Id " &
                        " WHERE TSPL_REQUISITION_HEAD.Status=0  group by TSPL_REQUISITION_DETAIL.Item_Code)xxx on xxx.Item_Code = TSPL_ITEM_REORDER_LEVEL_NEW.Item_Code" &
                        " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_REORDER_LEVEL_NEW.Item_Code left outer join TSPL_ITEM_UOM_DETAIL uom1 on uom1.Item_Code = TSPL_ITEM_REORDER_LEVEL_NEW.Item_Code and uom1.UOM_Code=(case when isnull(TSPL_ITEM_REORDER_LEVEL_NEW.UOM_Code,'')='' then uom1.UOM_Code else TSPL_ITEM_REORDER_LEVEL_NEW.Uom_Code end)" &
                        " where TSPL_ITEM_REORDER_LEVEL_NEW.Location_Code='" + strlocation + "' and TSPL_ITEM_REORDER_LEVEL_NEW.Reorder_Level>coalesce(xx.Balance,0)" &
                        " and TSPL_ITEM_UOM_DETAIL.Stocking_Unit='Y' and TSPL_ITEM_REORDER_LEVEL_NEW.Apply='Y')z where z.Qty>0"

                dt = clsDBFuncationality.GetDataTable(qry, Nothing)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal Then
                        clsCommon.MyMessageBoxShow("Some items reached their re-order level.", "Warning", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    '' done by Panch Raj on 29-11-2017
                    If clsCommon.CompairString(objCommonVar.CurrentUserCode, "admin") <> CompairStringResult.Equal Then
                        Dim qryCheck As String = " select count(*) as rec from TSPL_GROUP_PROGRAM_MAPPING " &
                                        " inner join TSPL_USER_GROUP_MAPPING on TSPL_GROUP_PROGRAM_MAPPING.Group_Code=TSPL_USER_GROUP_MAPPING.Group_Code " &
                                        " where TSPL_GROUP_PROGRAM_MAPPING.Program_Code='ITM-REOD-M' and TSPL_USER_GROUP_MAPPING.User_Code='" & objCommonVar.CurrentUserCode & "' and TSPL_GROUP_PROGRAM_MAPPING.Modify_Flag=1"
                        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qryCheck)) <= 0 Then
                            qryCheck = Nothing
                            Exit Sub
                        End If
                    End If


                    If clsCommon.MyMessageBoxShow("Some items reached their re-order level. Do you want to create auto indent ?", "Question", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then

                        Dim frm As New frmItemReorder()
                        frm.strLocation = strlocation
                        frm.ShowDialog()
                        'Dim obj As New clsRequistionHead()
                        'Dim drrow As DataRow() = Nothing
                        'For i As Integer = 0 To 6
                        '    If i = 0 Then
                        '        drrow = dt.Select("ItemType = 'F'")
                        '        obj.Item_Type = "F"
                        '    ElseIf i = 1 Then
                        '        drrow = dt.Select("ItemType = 'S'")
                        '        obj.Item_Type = "S"
                        '    ElseIf i = 2 Then
                        '        drrow = dt.Select("ItemType = 'R'")
                        '        obj.Item_Type = "R"
                        '    ElseIf i = 3 Then
                        '        drrow = dt.Select("ItemType = 'A'")
                        '        obj.Item_Type = "A"
                        '    ElseIf i = 4 Then
                        '        drrow = dt.Select("ItemType = 'T'")
                        '        obj.Item_Type = "T"
                        '    ElseIf i = 5 Then
                        '        drrow = dt.Select("ItemType = 'N'")
                        '        obj.Item_Type = "N"
                        '    ElseIf i = 6 Then
                        '        drrow = dt.Select("ItemType = 'O'")
                        '        obj.Item_Type = "O"
                        '    End If
                        '    If drrow IsNot Nothing AndAlso drrow.Count > 0 Then
                        '        obj.Requisition_Id = ""
                        '        obj.Requisition_Date = clsCommon.GETSERVERDATE(Nothing)
                        '        obj.On_Hold = 0
                        '        obj.Location = strlocation
                        '        obj.RQ_Detail_Total_Amt = 0
                        '        obj.Total_RQ_Amt = 0
                        '        obj.Mode_Of_Transport = "By Road"
                        '        obj.Is_Internal = "N"
                        '        'obj.Item_Type = clsCommon.myCstr(cboItemType.SelectedValue)
                        '        'obj.Dept = txtDept.Value
                        '        'obj.Dept_Desc = lblDept.Text
                        '        obj.Requisition_Type = "L"
                        '        obj.Category = "Regular"
                        '        obj.close_yn = "N"
                        '        obj.Approvel_Level_Required = 2

                        '        obj.ArrTr = New List(Of clsRequistionDetail)
                        '        For Each drow As DataRow In drrow
                        '            Dim objTr As New clsRequistionDetail()
                        '            objTr.Item_Code = clsCommon.myCstr(drow("ItemCode"))
                        '            objTr.Item_Desc = clsCommon.myCstr(drow("ItemDesc"))
                        '            objTr.Requisition_Qty = clsCommon.myCdbl(drow("Qty"))
                        '            objTr.Balance_Qty = clsCommon.myCdbl(drow("Qty"))
                        '            objTr.Item_Cost = 0
                        '            objTr.Item_Net_Amt = 0
                        '            objTr.Location = strlocation
                        '            objTr.Unit_Code = clsCommon.myCstr(drow("Unit"))
                        '            objTr.Status = "N"
                        '            If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                        '                obj.ArrTr.Add(objTr)
                        '            End If
                        '        Next
                        '        If obj.ArrTr IsNot Nothing AndAlso obj.ArrTr.Count > 0 Then
                        '            obj.SaveData(obj, "True")
                        '        End If
                        '    End If
                        'Next
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message.ToString())
        End Try
    End Sub

    Private Sub SystemIdleTimer1_OnEnterIdleState(ByVal sender As System.Object, ByVal e As IdleEventArgs) Handles SystemIdleTimer1.OnEnterIdleState
        isAutoClosing = True
        clsERPFuncationality.closeForm(Me)
    End Sub

    Private Sub SystemIdleTimer1_OnExitIdleState(ByVal sender As System.Object, ByVal e As IdleEventArgs) Handles SystemIdleTimer1.OnExitIdleState
        GC.Collect()
        GC.WaitForPendingFinalizers()
        isAutoClosing = False
    End Sub

    Public Sub GetMccFssaiPopUp()
        Dim obj As New FrmMCCMaster
        obj.GetMccFssai()
    End Sub

    Public Sub GetPendingSaleOrder()
        Try
            Dim qry As String = "select isnull((Select distinct '['+TSPL_SD_SALES_ORDER_HEAD.Document_Code+']  ' from TSPL_SD_SALES_ORDER_HEAD left outer join TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE on TSPL_SD_SALES_ORDER_HEAD.Document_Code=TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Against_Sales_Order where  TSPL_SD_SALES_ORDER_HEAD.Status=1 and TSPL_SD_SALES_ORDER_HEAD.Delivery_date > '" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MMM/yyyy") & "' and TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code not in (select isnull(Delivery_Code_PS,'')   from TSPL_SD_SHIPMENT_HEAD ) for xml path('')),'')  as DocNo "
            Dim strDocNo As String = clsDBFuncationality.getSingleValue(qry)
            If clsCommon.myLen(strDocNo) > 0 Then
                clsCommon.MyMessageBoxShow("Delivery Date will expired for these Sales order " & strDocNo & " ")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub GetPendingSaleBooking()
        Try
            Dim qry As String = "select isnull((Select distinct '['+TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code+']  ' from TSPL_BOOKING_MASTER_PRODUCTSALE left outer join TSPL_SD_SALES_ORDER_HEAD on TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code=TSPL_SD_SALES_ORDER_HEAD.Against_Booking_No where  TSPL_BOOKING_MASTER_PRODUCTSALE.Status=1 and TSPL_BOOKING_MASTER_PRODUCTSALE.BookValidity_date > '" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MMM/yyyy") & "' and ( TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code not in (select isnull(Against_Booking_No,'')   from TSPL_SD_SALES_ORDER_HEAD ) or TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code not in (select isnull(Against_Booking_No,'')   from TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE ) ) for xml path('')),'')  as DocNo "
            Dim strDocNo As String = clsDBFuncationality.getSingleValue(qry)
            If clsCommon.myLen(strDocNo) > 0 Then
                clsCommon.MyMessageBoxShow("Booking Date will expired for these Booking " & strDocNo & " ")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub btnLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogIn.Click
        CheckAndLogin()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Sub LoadWorkingScreen()
        SplitPanel1.Collapsed = True
        SplitPanel2.Collapsed = True
        SplitPanel4.Collapsed = True
        SplitPanel3.Collapsed = False
        LoadImageList()
        LoadMenu()
        'ToolWindow2.ToolCaptionButtons = ToolStripCaptionButtons.AutoHide




        Dim splitPanelElement = TryCast(RadDock1.RootElement.Children(0), SplitPanelElement)
        Dim imagePrimitive = New ImagePrimitive()
        '' Work done by Parteek on 08/01/2019 for BackGround Images
        Dim msBG As MemoryStream = Nothing
        Dim imgBG As Byte()
        Try
            imgBG = DirectCast(clsDBFuncationality.getSingleValue("select top 1 BackgroundImage  from tspl_company_master "), Byte())
            msBG = New MemoryStream(imgBG)
        Catch ex As Exception

        End Try

        If clsCommon.myLen(msBG) > 0 Then
            imagePrimitive.Image = Image.FromStream(msBG)
        Else
            imagePrimitive.Image = Global.ERP.My.Resources.Resources.BackImageXpertERP
        End If
        ' End
        '  imagePrimitive.Image = Global.ERP.My.Resources.Resources.BackImageXpertERP 'BackImageDemo
        'If Not objCommonVar.IsDemoERP Then
        '    imagePrimitive.Image = Global.ERP.My.Resources.Resources.BackImageXpertERPFMCGN
        'End If


        imagePrimitive.Alignment = ContentAlignment.TopRight
        imagePrimitive.StretchHorizontally = True
        imagePrimitive.StretchVertically = True


        imagePrimitive.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed
        RadDock1.MainDocumentContainer.SplitPanelElement.Children.Add(imagePrimitive)

        Try
            imagePrimitive = New ImagePrimitive()
            Dim img As Byte() = DirectCast(clsDBFuncationality.getSingleValue("select top 1 Logo_Img  from tspl_company_master "), Byte())
            Dim ms As MemoryStream = New MemoryStream(img)
            imagePrimitive.Image = Image.FromStream(ms)
            imagePrimitive.MaxSize = New Size(80, 50)
            imagePrimitive.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed
            imagePrimitive.Alignment = ContentAlignment.TopLeft
            'imagePrimitive.Size = New Point(50, 20)


            RadDock1.MainDocumentContainer.SplitPanelElement.Children.Add(imagePrimitive)
        Catch ex As Exception

        End Try


        lblUserCode.Text = objCommonVar.CurrentUserCode
        lblUser.Text = objCommonVar.CurrentUser
        lblCompanyCode.Text = objCommonVar.CurrentCompanyCode
        lblCompany.Text = Replace(objCommonVar.CurrentCompanyName.Trim(), "&", "&&")
        Me.Text = Me.Text + "-" + objCommonVar.CurrentCompanyName.Trim()
        strUserCode = objCommonVar.CurrentUserCode
        strCompany = objCommonVar.CurrentCompanyCode
        'lblDataBase.Text = objCommonVar.CurrDatabase.Trim()
        lblLocation.Text = objCommonVar.CurrLocationName.Trim()
        lblLocation.Text = clsLocation.GetName(clsGateEntry.getUsersDefaultLocation(), Nothing)



        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
            Me.WindowState = FormWindowState.Normal
        End If
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        Me.Height = Screen.PrimaryScreen.WorkingArea.Height
        Me.Location = New Point(0, 0)



        'RadDock1.BackColor = Color.Transparent
        'RadDock1.BackgroundImage = ImageList2.Images.Item(0)
        'RadDock1.BackgroundImageLayout = ImageLayout.Center

        If lblCompany.Text = "" Then
            RTV2.Visible = False
        Else
            RTV2.Visible = True
        End If

        LoadMenuInCombo()
        '' alert reminder for bday/anniversary 
        Dim UserCode As String = objCommonVar.CurrentUserCode
        If clsFixedParameter.GetData(clsFixedParameterType.AllowToDispalyAlertForBDayAnniversary, clsFixedParameterType.AllowToDispalyAlertForBDayAnniversary, Nothing) = "1" Then
            If clsEmployeeMaster.CheckUserForHRDepartment(UserCode) Then
                Dim msg As String = clsEmployeeMaster.GetBdayAnniversaryMSG()
                If clsCommon.myLen(msg) > 0 Then
                    clsERPFuncationalityOLD.ShowAlert(msg, "B'Day/Anniversary Reminder")
                End If
            End If
        End If
        '' email reminder for bday/anniversary 
        If clsFixedParameter.GetData(clsFixedParameterType.AllowToSendEmailForBDayAnniversary, clsFixedParameterType.AllowToSendEmailForBDayAnniversary, Nothing) = "1" Then
            If clsEmployeeMaster.CheckUserForHRDepartment(UserCode) Then
                clsEmployeeMaster.SendBdayAnniversaryEmail(clsUserMgtCode.frmEmployee_Master)
            End If
        End If


    End Sub

    Public Sub LoadMenu()
        Try
            arrExcluded.Clear()
            If Not isUtilityAdded Then
                arrExcluded.Add(clsUserMgtCode.ModuleUtility)
            End If
            arrExcluded.Add(clsUserMgtCode.frmJEReverse)
            'arrExcluded.Add(clsUserMgtCode.frmStanderdProductionEntry)
            If objCommonVar.IsDemoERP Then
                ''Menu item Only for  FMCG  
                arrExcluded.Add(clsUserMgtCode.ModuleSales)
                arrExcluded.Add(clsUserMgtCode.Indent)
                'arrExcluded.Add(clsUserMgtCode.Transfer)
                arrExcluded.Add(clsUserMgtCode.CreateTransfer)
                arrExcluded.Add(clsUserMgtCode.FrmItemMcMapping)
                arrExcluded.Add(clsUserMgtCode.mbtnEmptyTrans)
                arrExcluded.Add(clsUserMgtCode.frmMorningReport)
                arrExcluded.Add(clsUserMgtCode.StockStatement)
                arrExcluded.Add(clsUserMgtCode.FrmMCCMilkTransPortorInvoice)
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SHowBulkMilkWeighment, clsFixedParameterCode.SHowBulkMilkWeighment, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.frmWeighment)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.isIntimationRequired, clsFixedParameterCode.isIntimationRequired, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.frmIntimation)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowGateEntryAgainstPO, clsFixedParameterCode.AllowGateEntryAgainstPO, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.frmPOBulkProc)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowFreshPriceChartOnProductSale, clsFixedParameterCode.AllowFreshPriceChartOnProductSale, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.frmdispatchAdviceProductSale)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowGateReturn, clsFixedParameterCode.AllowGateReturn, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.frmGateEntryReturnTransfer)
                    arrExcluded.Add(clsUserMgtCode.frmGateEntryReturnPS)
                    arrExcluded.Add(clsUserMgtCode.frmGateEntryReturnCS)
                End If
                arrExcluded.Add(clsUserMgtCode.packType)
                'arrExcluded.Add(clsUserMgtCode.PriceMaster)
                arrExcluded.Add(clsUserMgtCode.SchemeMaster)
                arrExcluded.Add(clsUserMgtCode.mbtnBreakageHead1)
                arrExcluded.Add(clsUserMgtCode.ItemExciseMapping)
                arrExcluded.Add(clsUserMgtCode.ItemBasicPrice)

                arrExcluded.Add(clsUserMgtCode.ItemPrice)
                arrExcluded.Add(clsUserMgtCode.StockRecoReport)
                arrExcluded.Add(clsUserMgtCode.FrmStockDispatchReport)
                arrExcluded.Add(clsUserMgtCode.FrmAdjustmentStatusReport1)
                arrExcluded.Add(clsUserMgtCode.BreakageReportSummary)
                arrExcluded.Add(clsUserMgtCode.mbtnBreakageReport)
                arrExcluded.Add(clsUserMgtCode.RoutewiseBreakageReport)
                arrExcluded.Add(clsUserMgtCode.SchemeReport)
                arrExcluded.Add(clsUserMgtCode.StockReportForFinishedGoods)
                arrExcluded.Add(clsUserMgtCode.FrmAdjustmentReport)
                arrExcluded.Add(clsUserMgtCode.rptVehicleWiseLoadout)
                arrExcluded.Add(clsUserMgtCode.FrmPendingIndentTransferReport)
                arrExcluded.Add(clsUserMgtCode.FrmExpiredItemDetails)
                arrExcluded.Add(clsUserMgtCode.itemMaster)
                arrExcluded.Add(clsUserMgtCode.ShiptoLocation)
                'KUNAL > KDIL > REMOVED THE EMPTY LINK AS PER RANJANA MADAM DISCUSSION. > DATE 11-NOV-2016
                arrExcluded.Add(clsUserMgtCode.RptPaymentRelization)
                If clsCommon.CompairString(objCommonVar.CurrentIndustryType, "D") = CompairStringResult.Equal Then
                    'arrExcluded.Add(clsUserMgtCode.ACCSETMFG)
                    'arrExcluded.Add(clsUserMgtCode.COSTMAINTAIN)
                    'arrExcluded.Add(clsUserMgtCode.SETT)
                    'arrExcluded.Add(clsUserMgtCode.EXPENSE)
                    'arrExcluded.Add(clsUserMgtCode.PRO)
                    'arrExcluded.Add(clsUserMgtCode.ITEMCATEGORY)
                    'arrExcluded.Add(clsUserMgtCode.frmBOMImport)
                    'arrExcluded.Add(clsUserMgtCode.ALTER)
                    'arrExcluded.Add(clsUserMgtCode.frmBillOfMaterial)

                    'arrExcluded.Add(clsUserMgtCode.frmProductionPlanning)
                    'arrExcluded.Add(clsUserMgtCode.frmBatchOrder)
                    'arrExcluded.Add(clsUserMgtCode.frmProductionRequisition)
                    'arrExcluded.Add(clsUserMgtCode.frmStoreIssue)
                    'arrExcluded.Add(clsUserMgtCode.frmProductionReturn)
                    'arrExcluded.Add(clsUserMgtCode.frmProductionReceipt)

                    'arrExcluded.Add(clsUserMgtCode.LALT)
                    'arrExcluded.Add(clsUserMgtCode.LACCt)
                    'arrExcluded.Add(clsUserMgtCode.frmListOfBOM)
                    'arrExcluded.Add(clsUserMgtCode.LOIC)
                    'arrExcluded.Add(clsUserMgtCode.PRODREPORT)
                    'arrExcluded.Add(clsUserMgtCode.frmListofRequisition)
                    'arrExcluded.Add(clsUserMgtCode.Resource)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowGRN, clsFixedParameterCode.ShowGRN, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.mbtnGRN)
                End If
                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowMRN, clsFixedParameterCode.ShowMRN, Nothing)) = 0 Then
                    arrExcluded.Add(clsUserMgtCode.mbtnMRN)
                End If
                arrExcluded.Add(clsUserMgtCode.frmMilkCollectionArea)
                arrExcluded.Add(clsUserMgtCode.frmMilkVehicleTypeMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkTransportRateMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkComponentMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkComponentRateList)
                arrExcluded.Add(clsUserMgtCode.frmMilkAdvanceMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkRateTypeMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkShiftMaster)
                arrExcluded.Add(clsUserMgtCode.frmSeasonMaster)
                arrExcluded.Add(clsUserMgtCode.frmUOMMaster)

                arrExcluded.Add(clsUserMgtCode.frmMilkSuppliers)
                arrExcluded.Add(clsUserMgtCode.frmMCCRouteMapping)
                arrExcluded.Add(clsUserMgtCode.frmMCCSuperwiserMapping)
                arrExcluded.Add(clsUserMgtCode.frmMCCSupplierMapping)
                arrExcluded.Add(clsUserMgtCode.frmMilkCollection)
                arrExcluded.Add(clsUserMgtCode.frmMilkQualityCheck)
                arrExcluded.Add(clsUserMgtCode.frmMilkRateProcessingScheme)
                arrExcluded.Add(clsUserMgtCode.frmVehicleMovement)
                arrExcluded.Add(clsUserMgtCode.frmMilkBillGeneration)

                If IsLoaction_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmLocationCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmLocationCategoryStructure)
                End If
                If IsCustomer_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmCustomerCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmCustomerCategoryStructure)
                End If
                If IsVendor_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmVendorCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmVendorCategoryStructure)
                End If
                arrExcluded.Add(clsUserMgtCode.AssetSegment)
                'arrExcluded.Add(clsUserMgtCode.frmSecondaryCustomer)

            Else

                ''Menu item Only for Xpert ERP
                arrExcluded.Add(clsUserMgtCode.frmProductionReceiptDemo)
                arrExcluded.Add(clsUserMgtCode.frmProductionItemSerialReplace)
                arrExcluded.Add(clsUserMgtCode.frmProductionSerializedReport)

                arrExcluded.Add(clsUserMgtCode.DVAT30)
                arrExcluded.Add(clsUserMgtCode.DVAT31)
                'arrExcluded.Add(clsUserMgtCode.frmBalanceSheetPerforma)
                'arrExcluded.Add(clsUserMgtCode.rptBalanceSheet)
                arrExcluded.Add(clsUserMgtCode.ModuleSalesNew)
                'arrExcluded.Add(clsUserMgtCode.frmBarCodeGenerator)
                'arrExcluded.Add(clsUserMgtCode.frmRequisitionApproval)
                arrExcluded.Add(clsUserMgtCode.RequisitSubTypeMaster)
                arrExcluded.Add(clsUserMgtCode.mbtnPendingApprovalOfReq)
                arrExcluded.Add(clsUserMgtCode.RFQ)
                arrExcluded.Add(clsUserMgtCode.VendorQuotation)
                arrExcluded.Add(clsUserMgtCode.VendorComparison)
                arrExcluded.Add(clsUserMgtCode.VendorComparisonApproval)
                arrExcluded.Add(clsUserMgtCode.ModuleBI)
                arrExcluded.Add(clsUserMgtCode.FrmCFormEntry)
                'arrExcluded.Add(clsUserMgtCode.frmMapLedgerAccToTally)
                'arrExcluded.Add(clsUserMgtCode.frmPostAllGLToTally)
                arrExcluded.Add(clsUserMgtCode.frmCFormReport)
                arrExcluded.Add(clsUserMgtCode.frmPurchaseOrderList)
                arrExcluded.Add(clsUserMgtCode.FrmUserApproval)
                arrExcluded.Add(clsUserMgtCode.FrmBudgetMaintenance)
                arrExcluded.Add(clsUserMgtCode.ModuleProjectManagement)
                arrExcluded.Add(clsUserMgtCode.FrmExpenseType)
                arrExcluded.Add(clsUserMgtCode.FrmProjectStatus)
                arrExcluded.Add(clsUserMgtCode.FrmPJCExpense)
                arrExcluded.Add(clsUserMgtCode.stockRecoNew)
                arrExcluded.Add(clsUserMgtCode.FrmProcessMaster1)
                arrExcluded.Add(clsUserMgtCode.frmLabourWorkingSheet)
                arrExcluded.Add(clsUserMgtCode.frmOperaterEfficiencyReport)
                'arrExcluded.Add(clsUserMgtCode.ACCSETMFG)
                arrExcluded.Add(clsUserMgtCode.frmDemoProductionPlanning)
                arrExcluded.Add(clsUserMgtCode.COSTMAINTAIN)
                'arrExcluded.Add(clsUserMgtCode.SETT)
                arrExcluded.Add(clsUserMgtCode.EXPENSE)
                arrExcluded.Add(clsUserMgtCode.TOOLTYPE)
                arrExcluded.Add(clsUserMgtCode.frmWorkCenterMaster)
                arrExcluded.Add(clsUserMgtCode.frmResourceMaster)
                arrExcluded.Add(clsUserMgtCode.TOOL)
                arrExcluded.Add(clsUserMgtCode.frmOperationMaster)
                arrExcluded.Add(clsUserMgtCode.frmBOMImport)
                arrExcluded.Add(clsUserMgtCode.ALTER)
                arrExcluded.Add(clsUserMgtCode.FrmProcessMaster1)
                'arrExcluded.Add(clsUserMgtCode.frmBillOfMaterialCosting)
                arrExcluded.Add(clsUserMgtCode.frmManufacturingOrder)
                arrExcluded.Add(clsUserMgtCode.AssetSegment)
                arrExcluded.Add(clsUserMgtCode.FrmApprovalSetting)
                arrExcluded.Add(clsUserMgtCode.frmBarCodeGenerator1)
                arrExcluded.Add(clsUserMgtCode.WarrantyMaster)
                arrExcluded.Add(clsUserMgtCode.FrmItemSerialTrackingReport)
                arrExcluded.Add(clsUserMgtCode.ChangeItemSerialNumber)
                arrExcluded.Add(clsUserMgtCode.frmSchemeMasterNew) ' New Scheme Master @ Material Mgmt>>Master
                'arrExcluded.Add(clsUserMgtCode.frmWeightConversion) ' New Scheme Master @ Material Mgmt>>Master

                '' SETUP REPORTS PRODUCTION DEMO
                arrExcluded.Add(clsUserMgtCode.LALT)
                arrExcluded.Add(clsUserMgtCode.LACCt)
                arrExcluded.Add(clsUserMgtCode.LOIC)
                arrExcluded.Add(clsUserMgtCode.LOPER)
                arrExcluded.Add(clsUserMgtCode.Resource)
                arrExcluded.Add(clsUserMgtCode.LToolT)
                arrExcluded.Add(clsUserMgtCode.LTOOL)
                arrExcluded.Add(clsUserMgtCode.LWC)

                arrExcluded.Add(clsUserMgtCode.frmMilkCollectionArea)
                arrExcluded.Add(clsUserMgtCode.frmMilkVehicleTypeMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkTransportRateMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkComponentMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkComponentRateList)
                arrExcluded.Add(clsUserMgtCode.frmMilkAdvanceMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkRateTypeMaster)
                arrExcluded.Add(clsUserMgtCode.frmMilkShiftMaster)
                arrExcluded.Add(clsUserMgtCode.frmSeasonMaster)
                arrExcluded.Add(clsUserMgtCode.frmUOMMaster)

                arrExcluded.Add(clsUserMgtCode.frmMilkSuppliers)
                arrExcluded.Add(clsUserMgtCode.frmMCCRouteMapping)
                arrExcluded.Add(clsUserMgtCode.frmMCCSuperwiserMapping)
                arrExcluded.Add(clsUserMgtCode.frmMCCSupplierMapping)
                arrExcluded.Add(clsUserMgtCode.frmMilkCollection)
                arrExcluded.Add(clsUserMgtCode.frmMilkQualityCheck)
                arrExcluded.Add(clsUserMgtCode.frmMilkRateProcessingScheme)
                arrExcluded.Add(clsUserMgtCode.frmVehicleMovement)
                arrExcluded.Add(clsUserMgtCode.frmMilkBillGeneration)
                If IsLoaction_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmLocationCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmLocationCategoryStructure)
                End If
                If IsCustomer_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmCustomerCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmCustomerCategoryStructure)
                End If
                If IsVendor_NLevel = "NO" Then
                    arrExcluded.Add(clsUserMgtCode.frmVendorCategoryLevel)
                    arrExcluded.Add(clsUserMgtCode.frmVendorCategoryStructure)
                End If
                arrExcluded.Add(clsUserMgtCode.rptSparePartStatus)
            End If

            If clsCommon.CompairString(objCommonVar.CurrentIndustryType, "R") <> CompairStringResult.Equal Then
                arrExcluded.Add(clsUserMgtCode.frmRiceBOM)
                arrExcluded.Add(clsUserMgtCode.frmRiceMixingEntry)
                arrExcluded.Add(clsUserMgtCode.frmRiceProcessingEntry)
            End If

            If clsCommon.CompairString(clsFixedParameter.GetData("MilkProc", "EnableMilkProc", Nothing), "1") = CompairStringResult.Equal Then
                arrExcluded.Add(clsUserMgtCode.ModuleMilkProcurement)
            End If

            ' Hiding Redundant Copy of Price Chart Master , Done By Pankaj Jha As suggested by Ranjana MAM
            arrExcluded.Add(clsUserMgtCode.frmPriceChartMaster_Bulk)
            arrExcluded.Add(clsUserMgtCode.rptDailyProgressReport)
            arrExcluded.Add(clsUserMgtCode.rptShiftCodeWise)
            'arrExcluded.Add(clsUserMgtCode.RptBulkMilkRegister)
            arrExcluded.Add(clsUserMgtCode.rptSectionWiseStockReport)
            arrExcluded.Add(clsUserMgtCode.frmAssetRequisition)

            '' payroll reports 
            arrExcluded.Add(clsUserMgtCode.frmAditionalEarning_DeductionReport)
            arrExcluded.Add(clsUserMgtCode.frmAttendedDaysReport)
            arrExcluded.Add(clsUserMgtCode.frmDeductionDetailsReport)
            arrExcluded.Add(clsUserMgtCode.frmOT_Reports)
            arrExcluded.Add(clsUserMgtCode.frmSalaryComponentDetails)
            arrExcluded.Add(clsUserMgtCode.frmSalaryIncrementReport)
            arrExcluded.Add(clsUserMgtCode.frmSalarySheet_Reports)
            arrExcluded.Add(clsUserMgtCode.frmSalaryVoucher_Reports)
            arrExcluded.Add(clsUserMgtCode.frmVarianceReport)
            arrExcluded.Add(clsUserMgtCode.FrmSalarySlipRpt)
            arrExcluded.Add(clsUserMgtCode.FrmAMAcquisitionCode)

            '===shivani against ticket[BM00000009243,BM00000009240] 
            arrExcluded.Add(clsUserMgtCode.RptMultiplePaymentAdvice1)
            arrExcluded.Add(clsUserMgtCode.RptVehicleWiseReport)
            '==============
            'arrExcluded.Add(clsUserMgtCode.FrmEmployeePFRpt)
            ' arrExcluded.Add(clsUserMgtCode.RptBOILetterReport)
            'arrExcluded.Add(clsUserMgtCode.frmSalaryCertificate)
            'arrExcluded.Add(clsUserMgtCode.frmNewSalCertificate)
            arrExcluded.Add(clsUserMgtCode.frmEmployeeIncrement)
            'arrExcluded.Add(clsUserMgtCode.rptMCCMilkRegisterDripSaver)
            arrExcluded.Add(clsUserMgtCode.rptVSPOrVLCVarationRpt)
            '=======================
            arrExcluded.Add(clsUserMgtCode.frmLeaveAllotment)
            arrExcluded.Add(clsUserMgtCode.frmLeaveOpeningBalance)
            ''===against [BM00000008017]
            arrExcluded.Add(clsUserMgtCode.frmVisi_Install_Pullout_Report)
            arrExcluded.Add(clsUserMgtCode.frmDistributor_VS_SecondaryCustomer_Sale)
            arrExcluded.Add(clsUserMgtCode.frmSecondaryCustomer)
            arrExcluded.Add(clsUserMgtCode.frmSecondaryCustomerSale)
            arrExcluded.Add(clsUserMgtCode.frmVisi_Install_Pullout)

            arrExcluded.Add(clsUserMgtCode.mbtnItemMovement)
            arrExcluded.Add(clsUserMgtCode.DVAT30)
            arrExcluded.Add(clsUserMgtCode.DVAT31)
            arrExcluded.Add(clsUserMgtCode.CrptRG1Detail1)
            arrExcluded.Add(clsUserMgtCode.frmExciseChapterWise)
            arrExcluded.Add(clsUserMgtCode.FrmPurchasebookReport1)
            arrExcluded.Add(clsUserMgtCode.frmEmp_Id)
            arrExcluded.Add(clsUserMgtCode.frmLabelPrinting)
            arrExcluded.Add(clsUserMgtCode.frmPF_Covering_Letter)
            arrExcluded.Add(clsUserMgtCode.frmPF_Covering_Letter)
            arrExcluded.Add(clsUserMgtCode.frmBankStatement_Reports)
            arrExcluded.Add(clsUserMgtCode.frmSalaryCertificate)
            arrExcluded.Add(clsUserMgtCode.frmESICRpt)
            arrExcluded.Add(clsUserMgtCode.frmNewSalCertificate)
            arrExcluded.Add(clsUserMgtCode.RptBOILetterReport)
            arrExcluded.Add(clsUserMgtCode.rptMilkPaymentRegister)
            arrExcluded.Add(clsUserMgtCode.rptCollectionCenterChart)
            arrExcluded.Add(clsUserMgtCode.rptCollectionLevelChart)
            arrExcluded.Add(clsUserMgtCode.CustomerDetails)
            arrExcluded.Add(clsUserMgtCode.mbtnCustomerEmptyTrial)
            '' TDS EXCLUSION -Master
            arrExcluded.Add(clsUserMgtCode.FinancialYear)
            arrExcluded.Add(clsUserMgtCode.BranchDetails)
            arrExcluded.Add(clsUserMgtCode.ResponsiblePerson)
            arrExcluded.Add(clsUserMgtCode.StateCode)

            '' TDS EXCLUSION -Transaction
            arrExcluded.Add(clsUserMgtCode.mbtnCreateRemittance)
            arrExcluded.Add(clsUserMgtCode.remittanceentry)

            '' TDS EXCLUSION -Reports
            arrExcluded.Add(clsUserMgtCode.TDSForm26Q)
            arrExcluded.Add(clsUserMgtCode.Form16AReport)
            arrExcluded.Add(clsUserMgtCode.TDSSectionSummaryReport)
            ''Purchase Report
            arrExcluded.Add(clsUserMgtCode.Parti_VS_Rejected)
            arrExcluded.Add(clsUserMgtCode.frmPendingSaleInvoiceforChilpPO)
            If Not clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
                Dim dtNEWSC As DataTable = clsDBFuncationality.GetDataTable("select TSPL_USER_MASTER.Default_Location,TSPL_Mcc_MASTER.MCC_NAME,TSPL_MCC_MASTER.is_Reuired_Gate_Entry from TSPL_USER_MASTER  inner join TSPL_Mcc_MASTER on TSPL_Mcc_MASTER.mcc_code=Default_Location where TSPL_USER_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "' and isnull( is_Reuired_Gate_Entry,0)=1")
                If dtNEWSC Is Nothing OrElse dtNEWSC.Rows.Count <= 0 Then
                    arrExcluded.Add(clsUserMgtCode.MilkGateEntryIn)
                    arrExcluded.Add(clsUserMgtCode.MilkGateEntryWeightment)
                    arrExcluded.Add(clsUserMgtCode.MilkGateEntryOut)
                    'arrExcluded.Add(clsUserMgtCode.MilkReject) 'Ticket No : ERO/03/05/19-000586
                Else
                    dtNEWSC = Nothing
                End If
            End If


            If Not isLoadAppIntegrator Then
                arrExcluded.Add(clsUserMgtCode.frmAppIntegrator)
            End If
            If Not IsLoadMccBugReports Then
                arrExcluded.Add("SRNWtSample")
                arrExcluded.Add("SAMWTSRNRPT")
                arrExcluded.Add("RecWtSmpRpt")
                arrExcluded.Add("RcpWtDifRpt")
            End If

            'If Not isLoadBulkPurchaseUploader Then
            '    arrExcluded.Add(clsUserMgtCode.frmBulkPurchaseUploader)
            'End If
            If Not isLoadBankUpdateUploader Then
                arrExcluded.Add(clsUserMgtCode.FrmBankUpdateUploader)
            End If

            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.OpenPOforRejectShortageQty, clsFixedParameterCode.OpenPOforRejectShortageQty, Nothing)) = 0 Then
                arrExcluded.Add(clsUserMgtCode.RptPendingPO)
            End If

            'End Of code 
            '-----------------------------If not Process Production then Excluded menu's---------------------------------
            'Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
            'If Not OpenProcessProductionBOm Then
            '    arrExcluded.Add(clsUserMgtCode.frmProcessProductionIssueEntry)
            'End If
            '----------------------------------------------------------------------------------------------------------------

            '======================exclude schedule form is po scheduling off---
            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowPOScheduling, clsFixedParameterCode.AllowPOScheduling, Nothing)), "0") = CompairStringResult.Equal Then
                arrExcluded.Add(clsUserMgtCode.frmPurchaseSchedule)
            End If
            arrExcluded.Add(clsUserMgtCode.frmProductionEntryWithoutBatch)

            '=======================================================

            Dim strGrpWhrClas As String = ""
            Dim strReadPermission As String = ""
            If blnShowAllMenu = False Then
                strReadPermission = "TSPL_GROUP_PROGRAM_MAPPING.Read_Flag=1 and "
            End If
            If Not clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
                strGrpWhrClas += " and exists(select 1 from TSPL_GROUP_PROGRAM_MAPPING where " & strReadPermission & " TSPL_GROUP_PROGRAM_MAPPING.Program_Code=TSPL_PROGRAM_MASTER.Program_Code and TSPL_GROUP_PROGRAM_MAPPING.Group_Code in (select Group_Code  from TSPL_USER_GROUP_MAPPING where User_Code='" + objCommonVar.CurrentUserCode + "')) " + Environment.NewLine
            End If
            '===========Updated by rohit on may 27,2014. form will display according to module permission ===========
            Dim sQuery As String = "select Module_Name from TSPL_MODULE_PERMISSION"
            Dim dtmodule As DataTable = clsDBFuncationality.GetDataTable(sQuery)
            For Each rowModule As DataRow In dtmodule.Rows()
                If arrExcluded.Contains(rowModule("Module_Name")) Then
                    arrExcluded.Remove(rowModule("Module_Name"))
                End If
            Next

            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.INDUSTRYTYPE, clsFixedParameterCode.INDUSTRYTYPE, Nothing)), "A") <> CompairStringResult.Equal Then
                'arrExcluded.Add(clsUserMgtCode.frmPartNoMaster)
                arrExcluded.Add(clsUserMgtCode.ModuleServiceAndWarranty)
            End If

            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowQualityModuleInERP, clsFixedParameterCode.AllowQualityModuleInERP, Nothing)), "1") <> CompairStringResult.Equal Then
                arrExcluded.Add(clsUserMgtCode.ModuleQualityControl)
            End If
            arrExcluded.Add(clsUserMgtCode.frmPaySlip_Reports)
            arrExcluded.Add(clsUserMgtCode.rptFromNO21)
            arrExcluded.Add(clsUserMgtCode.rptFARReport)
            ' arrExcluded.Add(clsUserMgtCode.FrmItemTypeMaster)
            arrExcluded.Add(clsUserMgtCode.rptVLCwiseTPTimeTable)
            '=======================================================Ticket No : TEC/20/06/19-000559,TEC/20/06/19-000558====================================================
            arrExcluded.Add(clsUserMgtCode.frmFatSnfStockReport)
            arrExcluded.Add(clsUserMgtCode.RptDeliveryOrderReport1)
            arrExcluded.Add(clsUserMgtCode.rptDemandForsingleBranch)
            arrExcluded.Add(clsUserMgtCode.rptCrateJalliReportForTransfer)
            arrExcluded.Add(clsUserMgtCode.frmDistributerLedgerReport)
            arrExcluded.Add(clsUserMgtCode.rptDairySaleRegisterReport)
            arrExcluded.Add(clsUserMgtCode.RptEffectiveRateReport1)
            arrExcluded.Add(clsUserMgtCode.rptPlantCustomerDemand)
            arrExcluded.Add(clsUserMgtCode.RptDairyBookingDistributorReport)
            arrExcluded.Add(clsUserMgtCode.RptFlavouredMilk)
            arrExcluded.Add(clsUserMgtCode.RptMonthWiseSaleAnalysis)
            arrExcluded.Add(clsUserMgtCode.rptSalesHierarchyReport)
            arrExcluded.Add(clsUserMgtCode.rptDispatchChallanReportFresh)
            arrExcluded.Add(clsUserMgtCode.RptFreshBookingStatus)
            arrExcluded.Add(clsUserMgtCode.RptSubsidyCredit)
            arrExcluded.Add(clsUserMgtCode.MISStockLedgerReport)
            arrExcluded.Add(clsUserMgtCode.MeterialstockReco)
            arrExcluded.Add("DisChallan")
            arrExcluded.Add("ZnWisSale")
            'Ticket No : 
            'arrExcluded.Add(clsUserMgtCode.rptUnpostedPO)
            arrExcluded.Add(clsUserMgtCode.rptCDA)
            'arrExcluded.Add(clsUserMgtCode.mbtnGRNReport)
            'arrExcluded.Add(clsUserMgtCode.FrmIssueOrReturnItemWiseSummary)
            'arrExcluded.Add(clsUserMgtCode.frmSrnReport)
            arrExcluded.Add(clsUserMgtCode.rptSparePartStatus)
            arrExcluded.Add(clsUserMgtCode.FormMaster)
            arrExcluded.Add(clsUserMgtCode.frmBulkPostingNew)
            arrExcluded.Add(clsUserMgtCode.FrmFormSerialNoMaster)
            arrExcluded.Add(clsUserMgtCode.FrmCFormEntry)
            arrExcluded.Add(clsUserMgtCode.frmOpeningBalance)
            arrExcluded.Add(clsUserMgtCode.frmCFormReport)
            arrExcluded.Add(clsUserMgtCode.frmPromptMsgRelatedtopending)
            arrExcluded.Add(clsUserMgtCode.FrmPOSGRoupMaster)
            arrExcluded.Add(clsUserMgtCode.FrmConsumerDetailsForm)
            arrExcluded.Add(clsUserMgtCode.vendorSubgroup)
            arrExcluded.Add(clsUserMgtCode.FrmSupplierReg)
            arrExcluded.Add(clsUserMgtCode.FrmApprovedSuppliers)
            'arrExcluded.Add(clsUserMgtCode.VendorRegistration)
            'arrExcluded.Add(clsUserMgtCode.frmMapLedgerAccToTally)
            'arrExcluded.Add(clsUserMgtCode.frmPostAllGLToTally)
            ' arrExcluded.Add(clsUserMgtCode.chapterhead)
            arrExcluded.Add(clsUserMgtCode.frmStandardscheme)
            arrExcluded.Add(clsUserMgtCode.frmStandardRateItem)
            'arrExcluded.Add(clsUserMgtCode.frmBarCodeGenerator)
            arrExcluded.Add(clsUserMgtCode.frmPriceGroupMapping)
            arrExcluded.Add(clsUserMgtCode.frmWarehouseBreakage)
            arrExcluded.Add(clsUserMgtCode.frmBarCodeGenerator1)
            arrExcluded.Add(clsUserMgtCode.RequisitSubTypeMaster)
            arrExcluded.Add(clsUserMgtCode.FrmFormIssueReceiptEntry)
            arrExcluded.Add(clsUserMgtCode.frmVendorWiseReturnableGoodBalance)
            arrExcluded.Add(clsUserMgtCode.Vendor_Rating_Rejection)
            arrExcluded.Add(clsUserMgtCode.rptMaterialSendforJobWork)
            arrExcluded.Add(clsUserMgtCode.rptMaterialReceivedAfterJobWork)
            arrExcluded.Add(clsUserMgtCode.rptBalanceStockForJobWork)
            'arrExcluded.Add(clsUserMgtCode.frmPerformaInvoiceDairy)
            arrExcluded.Add(clsUserMgtCode.rptDemandForsingleBranch)
            arrExcluded.Add(clsUserMgtCode.rptCrateJalliReportForTransfer)
            arrExcluded.Add(clsUserMgtCode.rptDairySaleRegisterReport)
            arrExcluded.Add(clsUserMgtCode.RptEffectiveRateReport1)
            arrExcluded.Add(clsUserMgtCode.rptPlantCustomerDemand)
            arrExcluded.Add(clsUserMgtCode.frmItemChargeCategoryMaster)
            arrExcluded.Add(clsUserMgtCode.FreightChargesMaster)
            arrExcluded.Add(clsUserMgtCode.frmBookingDairyMultipleDistributor)
            arrExcluded.Add(clsUserMgtCode.frmPOSBookingDairyMultipleDistributor)
            arrExcluded.Add(clsUserMgtCode.FrmMCCTransactionApproval)
            arrExcluded.Add(clsUserMgtCode.frmMilkVehicleMaster)
            arrExcluded.Add(clsUserMgtCode.FrmBulkCreditLimitApproval)
            arrExcluded.Add(clsUserMgtCode.frmMilkCollectionLevels)
            arrExcluded.Add(clsUserMgtCode.FrmMCCMilkTransPortorInvoice)
            arrExcluded.Add(clsUserMgtCode.MccMilkTransferPrice)
            arrExcluded.Add("RcpChgRutRpt")
            arrExcluded.Add("MCCChngDet")
            ''  arrExcluded.Add(clsUserMgtCode.frmCostCenterTypeMaster)
            If clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ApplyFinancialCostCenter, clsFixedParameterCode.ApplyFinancialCostCenter, Nothing)) = "1", True, False)) = True Then
                arrExcluded.Add(clsUserMgtCode.CostFACenter)
            End If

            Dim qry As String = ""


            qry = "select distinct tt.* from (select Customise_SNo AS [SERNO],Program_Code,Name,Parent_Code from (" + Environment.NewLine '"select Program_Code,Name,Parent_Code from (" + Environment.NewLine
            qry += " select Program_Code,"
            If IsOriginalName = True Then
                qry += "Program_Name as Name, " + Environment.NewLine
            Else
                qry += "Case when LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as Name," + Environment.NewLine
            End If
            qry += " Parent_Code,Customise_SNo from TSPL_PROGRAM_MASTER  where 2=2 and  Parent_Code is null " + Environment.NewLine
            qry += " union " + Environment.NewLine
            qry += " select Program_Code,"
            If IsOriginalName = True Then
                qry += "Program_Name AS NAME," + Environment.NewLine
            Else
                qry += "case When LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as Name," + Environment.NewLine
            End If
            qry += " Parent_Code,Customise_SNo from TSPL_PROGRAM_MASTER  where 2=2 and  Type In ('SM') and Program_Code in (select distinct Parent_Code from TSPL_PROGRAM_MASTER where 2=2 " + strGrpWhrClas + " )" + Environment.NewLine
            qry += " union " + Environment.NewLine
            qry += " select Program_Code,"
            If IsOriginalName = True Then
                qry += "  Program_Name as NAME," + Environment.NewLine
            Else
                qry += "case when LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as Name," + Environment.NewLine
            End If
            qry += " Parent_Code,Customise_SNo from TSPL_PROGRAM_MASTER  where 2=2 and  APP_No=0 and Type In ('M') and Program_Code in (select distinct Parent_Code from TSPL_PROGRAM_MASTER where Program_Code in (select distinct Parent_Code from TSPL_PROGRAM_MASTER where 2=2 " + strGrpWhrClas + "))" + Environment.NewLine
            qry += " union " + Environment.NewLine
            qry += " select Program_Code,"
            If IsOriginalName = True Then
                qry += " Program_Name as NAME," + Environment.NewLine
            Else
                qry += "case when LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as Name," + Environment.NewLine
            End If
            qry += "Parent_Code,Customise_SNo from TSPL_PROGRAM_MASTER  where Program_Code='" + clsUserMgtCode.ModuleFavourite + "' " + Environment.NewLine
            qry += " union " + Environment.NewLine

            qry += " select Program_Code,"
            If IsOriginalName = True Then
                qry += " Program_Name as NAME, " + Environment.NewLine
            Else
                qry += "case when LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as Name,"
            End If
            qry += " Parent_Code,Customise_SNo from TSPL_PROGRAM_MASTER  where 2=2 and APP_No=0 " + strGrpWhrClas + Environment.NewLine

            qry += " union all " + Environment.NewLine
            qry += " select TSPL_FAVOURITE_MENU.Program_Code,"
            If IsOriginalName = True Then
                qry += "tspl_program_master.program_name AS NAME, " + Environment.NewLine
            Else
                qry += "case when LEN(isnull(TSPL_PROGRAM_MASTER.Re_Name,''))>0 then TSPL_PROGRAM_MASTER.Re_Name else TSPL_PROGRAM_MASTER.Program_Name end as Name, " + Environment.NewLine
            End If
            qry += "'" + clsUserMgtCode.ModuleFavourite + "' as Parent_Code,TSPL_FAVOURITE_MENU.SNo from TSPL_FAVOURITE_MENU left outer join  TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code= TSPL_FAVOURITE_MENU.Program_Code  where 2=2 and TSPL_FAVOURITE_MENU.User_Code='" + objCommonVar.CurrentUserCode + "' " + strGrpWhrClas + Environment.NewLine
            qry += " )xxx where 2=2 "
            qry += " and Program_Code not in (" + clsCommon.GetMulcallString(arrExcluded) + ") and Program_Code not in ( select Screen_Name from TSPL_MODULE_SCREEN_PERMISSION where  TSPL_MODULE_SCREEN_PERMISSION.P_Code= xxx.Parent_Code )"
            qry += ") tt inner join (select Module_Name,Program_Code as [prg_Code] from tspl_Program_Master tpm inner join tspl_Module_Permission tmm on " _
            & " tpm.Parent_Code=tmm.Module_Name  and tpm.APP_No=0 union select 'MFavourite','MFavourite' " & IIf(isUtilityAdded, " union select Program_Code as [Module_Name],Program_Code as [prg_Code] from tspl_Program_Master where Parent_Code ='Mutility'", "") & ") " _
            & " tpm on (tpm.module_Name=Parent_Code or tpm.prg_Code=Parent_Code or tpm.module_Name =Program_Code  or Parent_Code is NULL  or Parent_Code ='ExpertERP') " _
            & " and Program_Code not in (select distinct Program_Code as [prg_Code] from tspl_Program_Master tpm Left join tspl_Module_Permission tmm on tpm.Program_Code=tmm.Module_Name where Type='M' and module_Name is null " & IIf(isUtilityAdded, " and Program_Code <>'Mutility'", "") & ") order by SERNO" '" order by SNo"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)


            RTV2.DataSource = Nothing
            RTV2.TreeViewElement.AutoSizeItems = True
            RTV2.ShowLines = True
            RTV2.ShowRootLines = True
            RTV2.TreeViewElement.ViewElement.Margin = New Padding(4)
            RTV2.ShowExpandCollapse = True
            RTV2.TreeIndent = 15
            RTV2.FullRowSelect = False
            RTV2.ShowLines = True
            RTV2.LineStyle = TreeLineStyle.Dot
            RTV2.LineColor = Color.FromArgb(110, 153, 210)
            RTV2.ExpandAnimation = ExpandAnimation.Opacity
            RTV2.AllowEdit = False
            RTV2.ShowRootLines = False
            RTV2.ValueMember = "Program_Code"
            RTV2.DisplayMember = "Name"
            RTV2.ChildMember = "Program_Code"
            RTV2.ParentMember = "Parent_Code"
            RTV2.DataSource = dt

            LoadMenuInCombo()
            ' Set Image
            For i As Integer = 0 To RTV2.Nodes.Count - 1
                SetImage(RTV2.Nodes(i))
            Next
            RTV2.Nodes.Add("")
            RTV2.Nodes.Add("")
            RTV2.Nodes.Add("")
            RTV2.AllowEdit = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try

        RTV2.CollapseAll()
        If RTV2.Nodes.Count > 0 Then
            RTV2.Nodes(0).Expand()
        End If

    End Sub

    Protected Sub SetImage(ByVal subRoot As RadTreeNode)
        ' check for null (this can be removed since within th
        If (subRoot Is Nothing) Then
            Exit Sub
        End If
        If ArrImageList.ContainsKey(clsCommon.myCstr(subRoot.Value)) Then
            subRoot.Image = ImageList1.Images.Item(ArrImageList(clsCommon.myCstr(subRoot.Value)))
        End If
        ' add all it's children
        For i As Integer = 0 To subRoot.Nodes.Count - 1
            SetImage(subRoot.Nodes(i))
        Next
    End Sub

    Public Sub LoadImageList()
        ArrImageList.Clear()
        Dim qry As String = "select Program_Code,Image_Number from TSPL_PROGRAM_MASTER"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        For Each dr As DataRow In dt.Rows
            ArrImageList.Add(clsCommon.myCstr(dr("Program_Code")), clsCommon.myCdbl(dr("Image_Number")))
        Next

        qry = "select Program_Code  from TSPL_PROGRAM_MASTER where Parent_Code is null or Type in ('M')"
        dt = clsDBFuncationality.GetDataTable(qry)
        For Each dr As DataRow In dt.Rows
            ArrBold.Add(clsCommon.myCstr(dr("Program_Code")))
        Next
    End Sub

    Sub LoadMenuInCombo()
        'GC.Collect()
        Try
            If clsCommon.myLen(clsDBFuncationality.connectionString) > 0 Then
                Dim strGrpWhrClas As String = ""
                Dim strReadPermission As String = ""
                If blnShowAllMenu = False Then
                    strReadPermission = "TSPL_GROUP_PROGRAM_MAPPING.Read_Flag=1 and "
                End If
                If Not clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
                    strGrpWhrClas += " and exists(select 1 from TSPL_GROUP_PROGRAM_MAPPING where " & strReadPermission & " TSPL_GROUP_PROGRAM_MAPPING.Program_Code=TSPL_PROGRAM_MASTER.Program_Code and TSPL_GROUP_PROGRAM_MAPPING.Group_Code in (select Group_Code  from TSPL_USER_GROUP_MAPPING where User_Code='" + objCommonVar.CurrentUserCode + "')) " + Environment.NewLine
                End If
                Dim qry As String = "select Program_Code,"
                If IsOriginalName = True Then
                    qry += "Program_Name as PROGRAM_NAME"
                Else
                    qry += "case when LEN(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as PROGRAM_NAME"
                End If
                qry += " From TSPL_PROGRAM_MASTER  inner Join (Select Module_Name, Program_Code As [prg_Code] from tspl_Program_Master tpm inner join tspl_Module_Permission tmm On tpm.Parent_Code= tmm.Module_Name" _
                & " union Select 'MFavourite','MFavourite' " & IIf(isUtilityAdded, "Union select Program_Code as [Module_Name],Program_Code as [prg_Code] from tspl_Program_Master where Parent_Code ='Mutility'", "") & ") tmm on tspl_Program_Master.Parent_Code=tmm.prg_Code where 2=2 and  TSPL_PROGRAM_MASTER.Program_Code not in (" + Environment.NewLine
                qry += " select Program_Code from TSPL_PROGRAM_MASTER where Parent_Code in (select Program_Code from TSPL_PROGRAM_MASTER where Parent_Code in (select Program_Code from TSPL_PROGRAM_MASTER as innerProgramMaster where innerProgramMaster.Program_Code in (" + clsCommon.GetMulcallString(arrExcluded) + ") and Type='M') and Type='SM')"
                qry += " union "
                qry += " select Program_Code from TSPL_PROGRAM_MASTER where Program_Code in (" + clsCommon.GetMulcallString(arrExcluded) + ") and type=''"

                qry += " union select Screen_Name as Program_Code from TSPL_MODULE_SCREEN_PERMISSION  )  " + strGrpWhrClas + " and Type Not in ('M','SM')  and Parent_Code is not null   order by PROGRAM_NAME "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                Dim dr As DataRow = dt.NewRow()
                dr("Program_Code") = Nothing
                dr("PROGRAM_NAME") = Nothing
                dt.Rows.InsertAt(dr, 0)
                cboMenu.DataSource = dt
                cboMenu.ValueMember = "Program_Code"
                cboMenu.DisplayMember = "PROGRAM_NAME"
                cboMenu.SelectedIndex = 0
                cboMenu.DropDownListElement.AutoCompleteSuggest.SuggestMode = SuggestMode.Contains
            End If
        Catch ex As Exception

        End Try
        cboMenu.NullText = "Quick Menu"

    End Sub

    Private Sub MDI_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.F12 Then
            If lblDataBase.Visibility = ElementVisibility.Visible Then
                lblDataBase.Visibility = ElementVisibility.Hidden
                RadLabelElement1.Visibility = ElementVisibility.Hidden
                lblCompanyCode.Visibility = ElementVisibility.Hidden
            Else
                lblDataBase.Text = objCommonVar.CurrDatabase.Trim() + "[" + clsCommon.myCstr(clsDBFuncationality.getSingleValue("select @@SPID")) + "]"
                lblDataBase.Visibility = ElementVisibility.Visible
                RadLabelElement1.Visibility = ElementVisibility.Visible
                lblCompanyCode.Visibility = ElementVisibility.Visible
                lblCompanyCode.Text = "   "
            End If

        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.U Then
            isUtilityAdded = Not isUtilityAdded
            LoadMenu()
        ElseIf e.Control AndAlso e.KeyCode = Keys.U Then
            Dim frm As New FrmItemConverion()
            frm.Show()
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.C AndAlso clsCommon.CompairString(objCommonVar.CurrentUserCode, "pankajjha") = CompairStringResult.Equal Then
            Dim s As String = InputBox("Enter Query")
            clsCommon.MyMessageBoxShow(MSSQLTOORACLE.Covert(s))
        ElseIf e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.M Then
            Dim dbpwd As String = clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='USERPWD' and TYPE='PWD'")
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = "USERPWD"
            pwd.strType = "PWD"
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                '' Changes by Parteek for Screen Level Rights Ticket No : TEC/16/03/18-000101 on 01/05/2018
                'If EnableScreenSelection = True Then
                '    Dim frmmodule As New frmModule
                '    frmmodule.ShowDialog()
                '    LoadMenu()
                'Else
                Dim frmmodule As New frmModuleScreen
                frmmodule.ShowDialog()
                LoadMenu()
                ' End If
                'ENd
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.I Then
            Dim dbpwd As String = clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='USERPWD' and TYPE='PWD'")
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = "USERPWD"
            pwd.strType = "PWD"
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                isLoadAppIntegrator = True
                LoadMenu()
            Else
                isLoadAppIntegrator = False
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.P Then
            Dim dbpwd As String = clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='USERPWD' and TYPE='PWD'")
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.UploaderPassword
            pwd.strType = clsFixedParameterType.UploaderPassword
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                'isLoadBulkPurchaseUploader = True
                LoadMenu()
            Else
                isLoadBulkPurchaseUploader = False
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.R Then
            Dim dbpwd As String = clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='USERPWD' and TYPE='PWD'")
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.BankUploaderPassword
            pwd.strType = clsFixedParameterType.BankUploaderPassword
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                isLoadBankUpdateUploader = True
                LoadMenu()
            Else
                isLoadBankUpdateUploader = False
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.W Then
            Dim dbpwd As String = clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='USERPWD' and TYPE='PWD'")
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = "USERPWD"
            pwd.strType = "PWD"
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                IsLoadMccBugReports = True
                LoadMenu()
            Else
                IsLoadMccBugReports = False
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.Y Then
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.DocumentSequence
            pwd.strType = clsFixedParameterType.DocumentSequence
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                Dim frm As New frmDocumentSequence()
                frm.MdiParent = Me
                frm.Show()
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.Z Then

            Dim frmPWD As New FrmPWD(Nothing)
            frmPWD.strType = clsFixedParameterType.SettlementBankOnlyPWD
            frmPWD.strCode = clsFixedParameterCode.SettlementBankOnlyPWD
            frmPWD.ShowDialog()
            If frmPWD.isPasswordCorrect Then
                Dim frm As New frmUnCleardDoc()
                frm.MdiParent = Me
                frm.Show()
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.B Then
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.MilkProcurementUploader
            pwd.strType = clsFixedParameterType.MilkProcurementUploader
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                Dim frm As New frmMilkProcurementUploader()
                frm.MdiParent = Me
                frm.Show()
            End If
        ElseIf e.Control AndAlso e.Alt = False AndAlso e.Shift = False AndAlso e.KeyCode = Keys.B Then
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.MilkProcurementUploader
            pwd.strType = clsFixedParameterType.MilkProcurementUploader
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                Dim frm As New frmMilkShiftUploader()
                frm.MdiParent = Me
                frm.Show()
            End If
        ElseIf e.Control AndAlso e.Alt AndAlso e.Shift AndAlso e.KeyCode = Keys.S Then
            Dim pwd As New FrmPWD(Nothing)
            pwd.strCode = clsFixedParameterCode.TankerDispatchBulkUploader
            pwd.strType = clsFixedParameterType.TankerDispatchBulkUploader
            pwd.ShowDialog()
            If pwd.isPasswordCorrect Then
                Dim frm As New frmBulkProcurementUploader()
                frm.MdiParent = Me
                frm.Show()
            End If
        End If
    End Sub

    ' Recursive function to find matching node
    Public Function FindNode(nodes As RadTreeNodeCollection) As RadTreeNode
        Dim targetText As String = clsCommon.myCstr(cboMenu.SelectedText)
        Dim targetFormID As String = clsCommon.myCstr(cboMenu.SelectedValue)
        For Each node As RadTreeNode In nodes
            If node.Text = targetText AndAlso node.Value IsNot Nothing AndAlso node.Value.ToString() = targetFormID Then
                Return node
            End If
            Dim childMatch As RadTreeNode = FindNode(node.Nodes)
            If childMatch IsNot Nothing Then Return childMatch
        Next
        Return Nothing
    End Function

    Private Sub cboMenu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMenu.KeyDown
        'Try
        '    If clsCommon.myLen(clsCommon.myCstr(cboMenu.SelectedValue)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(cboMenu.Text)) > 0 Then
        '        If e.KeyCode = Keys.Enter Then
        '            ShowForm(clsCommon.myCstr(cboMenu.SelectedValue), clsCommon.myCstr(cboMenu.SelectedText), True)
        '            RTV2.CollapseAll()
        '            RTV2.Nodes(0).Expand()
        '            Try
        '                RTV2.SelectedNode = RTV2.Nodes(0)
        '                RTV2.SelectedNode = RTV2.Find(cboMenu.SelectedText)
        '                RTV2.SelectedNode.Expand()
        '            Catch ex As Exception
        '            End Try
        '        End If
        '    Else
        '        cboMenu.SelectedIndex = 0
        '    End If

        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        'End Try
        Try
            If clsCommon.myLen(clsCommon.myCstr(cboMenu.SelectedValue)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(cboMenu.Text)) > 0 Then
                If e.KeyCode = Keys.Enter Then
                    ShowForm(clsCommon.myCstr(cboMenu.SelectedValue), clsCommon.myCstr(cboMenu.SelectedText), True)
                    RTV2.CollapseAll()
                    RTV2.Nodes(0).Expand()
                    Try
                        Dim foundNode As RadTreeNode = Nothing
                        foundNode = FindNode(RTV2.Nodes)
                        If foundNode IsNot Nothing Then
                            RTV2.SelectedNode = foundNode
                            foundNode.Expand()
                            Dim parent As RadTreeNode = foundNode.Parent
                            While parent IsNot Nothing
                                parent.Expand()
                                parent = parent.Parent
                            End While
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Else
                cboMenu.SelectedIndex = 0
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RTV2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RTV2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If RTV2.SelectedNode IsNot Nothing Then
                Dim strCode As String = clsCommon.myCstr(RTV2.SelectedNode.Value)
                If clsCommon.myLen(strCode) > 0 Then
                    ShowForm(strCode, clsCommon.myCstr(RTV2.SelectedNode.Text), True)
                End If
            End If
        End If
    End Sub
    Public Function CountCounter(ByVal strProgramCode As String)
        clsDBFuncationality.ExecuteNonQuery("INSERT TSPL_PROGRAM_COUNTER(program_code ,Created_By,Created_Date) values ('" & strProgramCode & "','" + objCommonVar.CurrentUserCode + "',getdate() )")

        clsDBFuncationality.ExecuteNonQuery(" Select count(*) FROM TSPL_PROGRAM_COUNTER WHERE Created_By='" + objCommonVar.CurrentUserCode + "'")


    End Function
    Public Function setCountertoblockforOpenForm(ByVal strProgramCode As String)
        Try
            clsDBFuncationality.ExecuteNonQuery("Update TSPL_PROGRAM_MASTER set Form_Open_Counter=Form_Open_Counter+1  where program_code ='" & strProgramCode & "' ")
            clsDBFuncationality.ExecuteNonQuery("INSERT TSPL_PROGRAM_MASTER_COUNTER(program_code ,Created_By,Created_Date) values ('" & strProgramCode & "','" + objCommonVar.CurrentUserCode + "',getdate() )")

            If clsCommon.myCdbl(clsFixedParameter.GetSpecification(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, Nothing)) <> 1 Then
                Qry = clsFixedParameter.GetData(clsFixedParameterType.BigValidity, clsFixedParameterCode.BigValidity, Nothing)
                Dim BatchFileCounter As Integer = clsCommon.DecryptString(clsFixedParameter.GetData(clsFixedParameterType.BatchFileCounter, clsFixedParameterCode.BatchFileCounter, Nothing))
                If clsCommon.myLen(Qry) >= 0 Then
                    Dim strCounter As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select FormCounter from TSPL_PROGRAM_MASTER where program_code ='" & strProgramCode & "'"))
                    If strCounter >= BatchFileCounter Then
                        Qry = "select Exception_Msg,SNo from TSPL_Exception where SNo =(Select ExceptionNo from TSPL_PROGRAM_MASTER where program_code ='" & strProgramCode & "') "
                        dt = clsDBFuncationality.GetDataTable(Qry)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            If clsCommon.myLen(dt.Rows(0)("Exception_Msg")) > 0 Then
                                clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Exception_Msg")))
                                Return False
                            End If
                        End If
                    Else
                        If strCounter = 0 Then
                            Qry = "select max(SNo) as TotalMax,max(SNo * case when Is_Current=1 then 1 else 0 end) as CurrentMax from TSPL_Exception"
                            dt = clsDBFuncationality.GetDataTable(Qry)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                Qry = "update TSPL_Exception set Is_Current=0"
                                clsDBFuncationality.ExecuteNonQuery(Qry)
                                If clsCommon.myCdbl(dt.Rows(0)("TotalMax")) = clsCommon.myCdbl(dt.Rows(0)("CurrentMax")) Then
                                    Qry = "update TSPL_Exception set Is_Current=1 where SNo=1"
                                    clsDBFuncationality.ExecuteNonQuery(Qry)
                                Else
                                    Qry = "update TSPL_Exception set Is_Current=1 where SNo=" + clsCommon.myCstr(clsCommon.myCdbl(dt.Rows(0)("CurrentMax")) + 1) + ""
                                    clsDBFuncationality.ExecuteNonQuery(Qry)
                                End If
                            End If

                            Qry = "select Exception_Msg,SNo from TSPL_Exception where Is_Current=1"
                            dt = clsDBFuncationality.GetDataTable(Qry)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                If clsCommon.myLen(dt.Rows(0)("Exception_Msg")) > 0 Then
                                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_PROGRAM_MASTER set FormCounter =(select max(FormCounter)+1 from TSPL_PROGRAM_MASTER where program_code ='" & strProgramCode & "'),ExceptionNo='" & clsCommon.myCstr(dt.Rows(0)("SNo")) & "' where program_code ='" & strProgramCode & "' ")
                                    ''clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Exception_Msg")))

                                End If
                            End If
                        Else
                            clsDBFuncationality.ExecuteNonQuery("Update TSPL_PROGRAM_MASTER set FormCounter =(select max(FormCounter)+1 from TSPL_PROGRAM_MASTER where program_code ='" & strProgramCode & "') where program_code ='" & strProgramCode & "' ")
                            'Qry = "select Exception_Msg,SNo from TSPL_Exception where SNo =(Select ExceptionNo from TSPL_PROGRAM_MASTER where program_code ='" & strProgramCode & "') "
                            'dt = clsDBFuncationality.GetDataTable(Qry)
                            'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            '    If clsCommon.myLen(dt.Rows(0)("Exception_Msg")) > 0 Then
                            '        clsCommon.MyMessageBoxShow(clsCommon.myCstr(dt.Rows(0)("Exception_Msg")))
                            '    End If
                            'End If


                        End If
                        Return True

                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
        Return True
    End Function


    Public Sub ShowForm(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal isOpenInMDI As Boolean)
        ShowForm(strProgramCode, strProgramName, isOpenInMDI, "")
    End Sub
    Public Sub ShowForm(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal isOpenInMDI As Boolean, ByVal strDocNo As String, Optional ByVal IFTrueShowFormElseShowDialog As Boolean = True, Optional ByVal IsAllowModificationByApprovalUser As Boolean = False)
        GC.Collect()
        Dim strProgramCodeToOpen As String = strProgramCode
        If Not strProgramCode Is Nothing Then
            If setCountertoblockforOpenForm(strProgramCode) = True Then
                CountCounter(strProgramCode)
                If IsOriginalName = True Then
                    strProgramName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Program_Name as Program_Name from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "'"))
                Else
                    strProgramName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when LEN(ISNULL(Re_Name,''))>0 then Re_Name else Program_Name end as Program_Name from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "'"))
                End If
                Dim qry As String = " select * from tspl_Program_master where Program_code='" & strProgramCode & "'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim IsRunFromOtherAsm As Integer = clsCommon.myCdbl(dt.Rows(0)("IsLoadFromOtherAssembly"))
                    If IsRunFromOtherAsm = 1 Then
                        Dim FormName As String = clsCommon.myCstr(dt.Rows(0)("FormName"))
                        Dim AsmName As String = clsCommon.myCstr(Application.StartupPath & "\" & dt.Rows(0)("OtherAssemblyFilePathAndName"))
                        Dim AsmToLoad As Assembly = Nothing
                        Dim obj As Object = Nothing
                        AsmToLoad = Assembly.LoadFile(AsmName)
                        Dim classType As Type = AsmToLoad.[GetType](FormName)
                        'obj = AsmToLoad.CreateInstance("ERP." & FormName, True)
                        ' Dim M As Assembly.Module = AsmToLoad.FrmMainTranScreen
                        obj = AsmToLoad.CreateInstance(FormName, True)
                        Dim frm As RadForm = TryCast(obj, RadForm)
                        If isOpenInMDI Then
                            frm.MdiParent = Me
                            frm.Text = strProgramName
                            frm.Show()
                        Else
                            If clsCommon.myLen(strDocNo) > 0 Then
                                frm.Tag = strDocNo
                            End If
                            frm.WindowState = FormWindowState.Maximized
                            frm.Text = strProgramName
                            frm.ShowDialog()
                        End If
                        Exit Sub
                    End If

                    If clsCommon.myLen(dt.Rows(0)("Program_Code_Original")) > 0 Then
                        strProgramCodeToOpen = clsCommon.myCstr(dt.Rows(0)("Program_Code_Original"))
                    End If
                End If
                Select Case strProgramCodeToOpen

                    Case clsUserMgtCode.FrmCompanyMaster
                        frm = New FrmCompanyMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CostCenter
                        frm = New FrmCostCenter(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CostFACenter
                        frm = New FrmFACostCenter(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ReverseEntry
                        frm = New FrmReverseEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.cityMaster
                        frm = New frmCityMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.taxAuthority
                        frm = New frmTaxAuthority(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.taxRate
                        frm = New FrmTaxRates(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.taxGroup
                        frm = New FrmTaxGroups(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.paymentTerms
                        frm = New frmPaymentTerms(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.paymentCodes
                        frm = New FrmPaymentCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLocationDistanceMapping
                        frm = New frmLocationDistanceMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.employeeMaster) = CompairStringResult.Equal  Then
                    '    frm=New frmEmployeeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.designationMaster
                    '    frm=New frmDesignationMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.rptDBTSummaryMonthlyWise
                        frm = New rptDBTSummaryMonthlyWise()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BankGroupMaster
                        frm = New FrmBankGroupMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.bankMaster
                        frm = New frmBankMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.bankBranchMaster
                        frm = New FrmBankBrachMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNotificationScreen
                        frm = New FrmNotificationScreen
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.ChangePwd
                        frm = New FrmChangePassword()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAbateMentMaster
                        frm = New FrmAbateMentMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.PrefixGeneration
                        Dim frm1 As New FrmPrefixGenerationNew()
                        formShow(frm1, strProgramCode, strProgramName, True, strDocNo)
                    Case clsUserMgtCode.FormMaster
                        frm = New FrmFormMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFormSerialNoMaster
                        frm = New FrmFormSerialNoMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAdditionalCharges
                        frm = New FrmAdditionalCharges(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case strProgramCode = "FrmChangePassword"
                        frm = New FrmChangePassword()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmCurrencyConversion
                    '    frm = New frmCurrencyConversion
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.CustomFieldMaster
                    '    frm = New FrmCustomFieldMaster()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.CustomFieldMapping
                    '    frm = New frmCustomFieldMapping()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmModuleCurrencyMapping
                    '    frm = New frmModuleCurrencyMapping()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CommonServicesSetting
                        frm = New frmCommonServicesSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRegionMaster
                        frm = New FrmRegionMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCountryMaster1
                        frm = New frmCountryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStateMaster1
                        frm = New frmStateMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DistrictMaster
                        frm = New frmDistrictMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AreaMaster
                        frm = New frmAreaMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPromptMsgRelatedtopending
                        frm = New FrmPromptMsgRelatedToPendency
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSaleReco
                        frm = New rptSaleRecoNew
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Re-added by stuti on my computer ---
                    Case clsUserMgtCode.rptPurReco
                        frm = New rptPurchaseReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVendorReco
                        frm = New rptVendReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptVSPCustomerReco
                        frm = New rptVSPCustomerReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptSRNReco
                        frm = New rptSRNReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTankerDispatchGainLossReco
                        frm = New rptTankerDispatchGainLossReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptCustomerReco
                        frm = New rptCustomerReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptInvReco
                        frm = New RptInventoryRecoReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptTransporterProvisionReport
                        frm = New frmRptTransporterProvision()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBatchItemReport1
                        frm = New rptBatchItemReport1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptSaleAccountSetList
                        frm = New rptSaleAccountSetList()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptItemPurchaseSet
                        frm = New RptItemPurchaseAccount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptItemSalePurchaseSet
                        frm = New RptItemSalePurchaseAccount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.FrmItemCostUOM
                        frm = New RptItemCostUOM()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    'Case clsUserMgtCode.FrmBranchAccountMapping
                    '    frm = New FrmBranchAccountMapping
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMCCDiscountMaster
                        frm = New FrmDiscountMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "CaLSCRN"
                        System.Diagnostics.Process.Start("calc.exe")

                    Case clsUserMgtCode.RptBranchAccountMapping
                        frm = New RptBranchAccountMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmLockTransactionReport
                        frm = New FrmLockTransactionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptBoothWiseBillReport
                        frm = New rptBoothWiseBillReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '------------------ Common services Transactions---------------------------------------
                    Case clsUserMgtCode.bankTransfer
                        frm = New FrmBankTransfer(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.reverseTransaction
                        frm = New frmReverseTransaction(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankBook
                        frm = New FrmBankBook(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReconcilationRpt
                        frm = New frmReconcilationReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankBookRecoReport
                        frm = New frmBankBookRecoReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankBookLocationDetail
                        frm = New FrmBankBookLocationDetail(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmBankBookChart
                    '    frm = New FrmBankBookChart()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmBankBookClosing
                    '    frm = New FrmBankBookClosing()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmCustomerAgingSummary
                    '    frm = New FrmBICustomerAgeing()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmVendorAgingSummary
                    '    frm = New FrmBIVendorAgeing()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmFormCollection
                        frm = New FrmFormCollection()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLoadReport
                        frm = New FrmLoadReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLoadOutInvoiceRecoReport
                        frm = New FrmLoadOutInvoiceRecoReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCFormEntry
                        frm = New FrmCFormEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBankGuaranteeMaster
                        frm = New FrmBankGuaranteeMaster1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BankOpeningReco
                        frm = New frmBankOpeningReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '-------------------Common Services Report -----------------------
                    Case clsUserMgtCode.DVAT30
                        frm = New FrmDVAT30()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DVAT31
                        frm = New FrmDVAT31()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDetailsOfForm2B
                        frm = New FrmDetailsOfForm2B()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCashVoucher
                        frm = New FrmCashVoucher
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.fundprovision
                        frm = New rptFundProvision
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TaxTracking
                        frm = New FrmTaxTracking
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FormIssue
                        frm = New FrmFormIssueDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBankReconcilliation
                        frm = New RptBankReconcilliation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RevaluationEntry
                    '    frm = New frmRevaluationEntry
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------Receivables---------------------------------------
                    'Case clsUserMgtCode.ShiptoLocation
                    '    frm=New frmShipToLocation(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmShipToLocationDetails
                        frm = New FrmShipToLocationDetails(lblUser.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmReceivablePaymentTerms
                        frm = New FrmReceivablePaymentTerms
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmReceivableSettings
                        frm = New FrmReceivableSettings
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerType
                        frm = New frmCustomerType(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomeCategory
                        frm = New frmCustomerCategory(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.mbtnCustomerInfo
                    '    frm=New FrmCustomerInfo(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.CustomerMaster
                        frm = New frmCustomer(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.SecondaryCustomerMaster
                        frm = New FrmSecondaryCustomerMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmCompetitorMaster
                    '    frm = New frmCompetitorMaster
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerRouteShiftMaster
                        frm = New frmCustomerRouteShiftMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPOSGRoupMaster
                        frm = New FrmPOSGRoupMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.CustomerGroup
                        frm = New frmCustomerGroup(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerAccountSet
                        frm = New frmCustomerAccountSet(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '-----------------Receivables Transactions----------

                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.ReceiptEntry
                    '    'Xtra.UpdateSaleInvoiceBalanceAmt()
                    '    frm=New FrmReceiptNew(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.ReceiptEntry
                        frm = New FrmReceipttNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FinanceAdjustment
                    '    frm = New frmAdj()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.FrmCustomersSetOff
                        frm = New FrmCustomerSetOff()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCustomersOutstanding
                        frm = New FrmCustomerOutstanding()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCustomersSetOff
                        frm = New RptCustomerSetOff()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ReceiptAdjustmentEntry
                        frm = New frmAdj()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnARInvoiceEntry
                        frm = New FrmARInvoiceEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQuickBook
                        frm = New FrmQuickEntry1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMakePayment
                        frm = New FrmMakePayment(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBankReco
                        frm = New FrmBankReco(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCustomerInquiry
                        'Xtra.UpdateSaleInvoiceBalanceAmt()
                        frm = New FrmCustomerInquiry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '-----------------Receivables Reports----------
                    Case clsUserMgtCode.CustomerGroupReport
                        frm = New FrmCustomerGroupReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerDetails
                        frm = New FrmCustomerMasterReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.SaleRegister
                    ''    frm=New FrmRptSales()
                    ''         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.mbtnCustomerLedger
                        frm = New FrmRptCustomerLedgerDemo(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.CustomerLedgerVsAgeing
                        frm = New frmCustomerLedgerVsAgeing(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnCustomerLedgerZoneAreaWise
                        frm = New FrmRptCustomerLedgerDemoZoneAreaWise(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRoute_CustomerOutstanding
                        frm = New FrmRoute_CustomerOutStanding()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerAgeing
                        'If objCommonVar.IsDemoERP Then
                        '    'frm=New FrmCustomerAgingDEMO()
                        '    '     formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                        '    frm = New rptCustomerAgeingDrillDown(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                        'Else
                        '    frm = New FrmCustomerAgeing()
                        '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                        'End If
                        frm = New rptCustomerAgeingDrillDown(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmCustomerAdvanceKnockOff
                    'frm = New rptCustomerAdvanceKnockOff(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                    'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomersListReport
                        frm = New frmCustomerListRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RouteListReport11
                        frm = New frmRouteListReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.mbtnCustomerEmptyTrial  ''Added By Pankaj
                        frm = New FrmCustomerEmptyTrial2()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSecurityLevel
                        frm = New FrmSecurityLevel
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVendorSecurity
                        frm = New FrmVendorSecurity
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCarteJaliRpt
                        frm = New FrmCrateJaliReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIBPReport
                        frm = New frmIBPReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCrateRegister
                        frm = New rptCrateRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCrateJalliReportForTransfer
                        frm = New RptCrateJalliBoxTransferDS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmExciseChapterWise  ''Added By Manoj

                        If objCommonVar.IsDemoERP Then
                            frm = New frmER1Demo()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmExciseChapterWise()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If


                    'Case clsUserMgtCode.FrmSecurityDeposit1  ''Added By Manoj
                    '    frm = New FrmSecurityDeposit1()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.FrmBankReverse  ''Added By Abhishek  as on 27 Nov 2012
                        frm = New FrmBankReverse()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmCustomerOutstanding
                    '    frm = New FrmCustomerOutstanding()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.RptQualityStatus
                        frm = New RptQualityStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '------------------ GL Masters---------------------------------------

                    Case clsUserMgtCode.glOptions
                        frm = New frmgloption(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.segmentCode
                        frm = New Frmsegmentcode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.accountStructure
                        frm = New frmGLStructure(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.accountGroup
                        frm = New frmAccountGroup(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.glAccount
                        frm = New frmGLAccount(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.createAccounts
                        frm = New frmCreateAccountNew(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.sourceCode
                        frm = New FrmSourceCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.glsecurity
                        frm = New Frmglsecurity(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBalanceSheetPerforma
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.BalanceSheetPerformaWithFormula, clsFixedParameterCode.BalanceSheetPerformaWithFormula, Nothing)) > 0 Then
                            frm = New frmBalanceSheetPerformaFormula()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New FrmBalanceSheetPerforma1()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If

                    'Case clsUserMgtCode.CashFlowPerforma
                    '    frm = New frmCashFlowPerforma()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmGL_account_excluded
                    '    frm = New FrmGL_account_excluded()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmGLControlAccountMapping
                        frm = New frmGLControlAccountMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frm_Account_Mapping
                        Dim fmr As New Frm_Account_Mapping()
                        formShow(fmr, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.frmMapLedgerAccToTally
                        frm = New frmMapLedgerAccToTally()
                        If objCommonVar.IsSendToTally Then
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmPostAllGLToTally
                        frm = New frmPostAllGLToTally()
                        If objCommonVar.IsSendToTally Then
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.FiscalYear
                        frm = New frmFinancialYearMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CostCentreFinancial
                        frm = New FrmCostCentreFinancial()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AccountMainGroup
                        frm = New FrmAccountMainGroup()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AccountSubGroup
                        frm = New FrmAccountSubGroup()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AccountGLMainAccount
                        frm = New frmAccountMainGLAccount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ GL Transactions---------------------------------------

                    Case clsUserMgtCode.journalEntry
                        frm = New frmJournalEntry(lblUserCode.Text, objCommonVar.CurrentCompanyCode, strDocNo, clsUserMgtCode.journalEntry)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ReversejournalEntry
                        frm = New frmJournalEntry(lblUserCode.Text, objCommonVar.CurrentCompanyCode, strDocNo, clsUserMgtCode.ReversejournalEntry)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnVCGLEntry
                        frm = New frmVCGLEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJEReverse
                        frm = New FrmJEReverse()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSettingDetails
                        frm = New frmSettingDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '------------------ GL Transactions Report---------------------------------------
                    Case clsUserMgtCode.FrmCostCenterAnalysisRpt
                        frm = New FrmCostCenterAnalysisRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGLTransReport
                        frm = New GLTransReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptChartOfAccount
                        frm = New RptChartOfAccount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmJrnlVoucher
                        frm = New JrnlVoucherReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf strFormName = "Trial Balance Report" Then
                    '    frm=New frmTrialBalanceReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.JECheckSystem
                        frm = New rptJECheck()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTrialBalance
                        frm = New frmRptTrialBalanceNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTrialBalanceCV
                        frm = New frmRptTrialBalanceVC()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnJournalBook
                        frm = New frmRptDayWiseJournalBook()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBalanceSheet
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.BalanceSheetPerformaWithFormula, clsFixedParameterCode.BalanceSheetPerformaWithFormula, Nothing)) > 0 Then
                            frm = New frmRptBalanceSheetFormula()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmRptBalanceSheet()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.CashFlow
                        frm = New frmRptCashFlow()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankBookDayWise
                        frm = New FrmBankBookDayWise()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Added By abhishek as on 26/11/2012
                    Case clsUserMgtCode.frmUnpostedJV
                        frm = New FrmUnpostedJV()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Code End 
                    '------------------ Administrative Services---------------------------------------
                    Case clsUserMgtCode.EmployeeMaster
                        frm = New frmEmployeeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ' KUNAL > TICKET : BM00000009879 > 30 - SEP - 2016 
                    Case clsUserMgtCode.DesignationMaster
                        frm = New frmDesignationMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode, clsUserMgtCode.DesignationMaster)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DesignationMasterHierarchy
                        frm = New frmDesignationHierarchyMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.userMaster
                        frm = New FrmUserMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.TimeTable
                    '    frm = New frmTimeTable
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Security_Matr
                        frm = New RptSecurityMatrix
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDocumentVersionReport
                        frm = New FrmDocumentVersionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSystemAdministratorTool
                        frm = New FrmSystemAdministratorTool
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.userGroupMaster
                        frm = New FrmUserGroupMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.userGroupMapping
                        frm = New FrmUserGroupMapping(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.groupProgramMapping
                        frm = New GroupProgramMapping(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmScheduling
                        frm = New FrmScheduling()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmApprovalLevelScreen
                        frm = New frmApprovalScreen()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmApprovalAlertSumm
                        frm = New FrmApprovalAlertSumm()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNotification
                        frm = New frmNotification()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmOfficeOrderTemplate
                        frm = New FrmOfficeOrderTemplate()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLocationSetting
                        frm = New frmLocationLogin()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmSynchronization
                    '    frm = New frmSynchronization
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FisaclYearEndProcess
                        frm = New FrmFiscalYearEndProcess
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAppIntegrator
                    '    frm = New frmAppIntegrator
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ Purchase Masters---------------------------------------
                    Case clsUserMgtCode.vendormaster
                        frm = New frmVendorMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.capexmaster
                        frm = New FrmCapexMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.capexbudget
                        frm = New FrmCapexBudget()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCapxRevHis
                        frm = New RptCapexBudgetRevHis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorRegistration
                        frm = New FrmVendorReg()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.EmployeeBandMaster
                        frm = New FrmEmployeeBandMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.TankerMasterSale
                        frm = New frmTankerMasterSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.vendoraccountset
                        frm = New frmvendoraccountset(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Farmeraccountset
                        frm = New frmvendoraccountset(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.vendorgroup
                        frm = New frmVendorGroup(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.vendortype
                        frm = New frmVendorType(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmRequisitionApproval
                    '    frm = New FrmRequisitionApproval()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RequisitSubTypeMaster
                        frm = New FrmRequisitSubTypeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHirerachyLevelMaster
                        frm = New FrmHirerachyLevelMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Anubhooti 02-Sep-2014 
                    Case clsUserMgtCode.FrmPayableSettings
                        frm = New FrmPayableSettings
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Anubhooti 05-Sep-2014 BM00000003755
                    Case clsUserMgtCode.FrmPaymentUploader
                        frm = New FrmPaymentUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorBankMaster
                        frm = New FrmVendorBankMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.vendorSubgroup
                        frm = New frmVendorsubGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '-----------------Purchase Transactions--------------------------------------
                    Case clsUserMgtCode.PaymentEntryNew
                        frm = New FrmPaymentNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVendorSetOff
                        frm = New FrmVendorSetOff()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmEmployeeSetOff
                        frm = New FrmEmployeeSetOff()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVSPCSASetOff
                        frm = New FrmVSPCSASetOff()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMultipleRTGS
                        frm = New RptMultipleRTGS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMultiplePaymentAdvice1
                        frm = New RptMultiplePaymentAdvice1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PaymentAdjustmentEntry
                        frm = New frmPaymentAdjEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnAPInvoiceEntry
                        frm = New FrmAPInvoiceEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVendorService
                        frm = New FrmVendorService()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmVendorInquiry
                        frm = New FrmVendorInquiry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPurchaseHistory
                        frm = New FrmPurchaseHistory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTenderTrackingReport
                        frm = New FrmTenderTrackingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmERPStatusTrackingReport
                        frm = New FrmERPStatusTrackingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptrlPenaltyRegister
                        frm = New rptrlPenaltyRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptPerformanceReport
                        frm = New rptPerformanceReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.VehicleUnloadingReport
                        frm = New VehicleUnloadingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptRMUnloading
                        frm = New rptRMUnloading()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.rptTransactionWiseStock
                        frm = New rptTransactionWiseStock()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.Rptralwisematerialreceivedandpayment
                        frm = New Rptralwisematerialreceivedandpayment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RmSecurityDeduction
                        frm = New RmSecurityDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.WeightmentAutoAndManualReport
                        frm = New Weightment_Auto_and_Manual_Report()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmHSNMaster
                        frm = New frmHSNMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmOverheadCostMaster
                        frm = New frmOverheadCostMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmOverheadCostGroup
                        frm = New FrmOverheadCostGroup()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmItemCostMapping
                        frm = New FrmItemCostMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWeightUomMaster
                        frm = New frmWeightUomMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmSupplierReg
                        frm = New FrmSupplierReg()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmApprovedSuppliers
                        frm = New FrmApprovedSuppliers()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '----------------------------Purchase Report --------------------------------
                    Case clsUserMgtCode.frmScrapSaleInvoice
                        frm = New FrmScrapSaleInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPaymentEntry
                        frm = New FrmPaymentEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnAPInvoiceReport
                        frm = New frmRptAPInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptAPInvoiceDetailsReport
                        frm = New FrmRptAPInvoiceDetailsReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptHierarchyWiseReport
                        frm = New RptHierarchyWiseReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAdvancePaymentRegister
                        frm = New FrmAdvancePaymentRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptAPReport
                        frm = New rptAPReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptGSTRReport
                        frm = New rptGSTR()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVendorAccountSetReport
                        frm = New rptVendorAccountSetReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptARReport
                        frm = New RptARReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptShippmentClearing
                        frm = New FrmShipmentReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPayableClearing
                        frm = New rptpayableReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptCustomerAccountSet
                        frm = New RptCustomerAccount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.VendorLedgerReport
                        frm = New frmRptVendorLedger(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorLedgerVsAgeing
                        frm = New frmRptVendorLedgerVsAgeing(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorCustomerLedgerReport
                        frm = New frmRptVendorCustomerLedger(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAgingPayble
                        frm = New rptAPAgeingDrillDown(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRptVendorList
                        frm = New FrmRptVendorList()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptVendorAgeingDetails
                        frm = New FrmRptVendorTransList()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptVendorTransList
                        frm = New FrmRptVendorTransHistory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSrnReport
                        frm = New FrmSrnReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAgingDrillDown
                    '    frm=New rptAPAgeingDrillDown(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    '------------------ (Material Management)Inventory Masters---------------------------------------
                    Case clsUserMgtCode.PricePlan
                        frm = New frmPriceMasterPlan()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PriceMaster
                        frm = New FrmPriceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PriceComponentMasters
                        frm = New FrmPriceComponantMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PriceComponentMapping
                        frm = New FrmPriceComponantMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SchemeMaster
                        frm = New FrmSchmeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.inventorySetting
                        frm = New frmInventorySetting(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.chapterhead
                        frm = New frmChapterHead(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.itemStructure
                        frm = New frmItemStructure(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.itemGroups
                        frm = New frmItemGroup(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.itemMaster
                    '    frm = New frmItemMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.locationMaster
                        frm = New frmLocationMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerLocationMapping
                        frm = New FrmCustomerLocationMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.itemPurchaseAccount
                        frm = New frmPurcahseAccountSetCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.itemSaleAccount
                        frm = New frmSaleAccountSetCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "itemPriceList"

                    Case clsUserMgtCode.GSTunitMeasure
                        frm = New frmGSTunitMeasure()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.unitMaster
                        frm = New frmUnitOfCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "conversionMaster"

                    Case "itemPriceMaster"

                    Case clsUserMgtCode.packType
                        frm = New Frmpacktype(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "frmExcisableLocationDetails"
                        frm = New FrmExcisableLocationDetails(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmItemMasterRMOther
                        frm = New FrmItemMasterRMOther()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TypeOfItemMaster
                        frm = New TypeOfItemMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBatchManufacturingMaster
                        frm = New frmBatchManufacturingMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmPartNoMaster
                        frm = New FrmPartNoMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.InvetorySourceCode
                        frm = New frmInventorySourceCode()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '===============Greivance Type ======================
                    Case clsUserMgtCode.frmGrievanceTypeMaster
                        frm = New frmGrievanceTypeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    'Case clsUserMgtCode.ItemRackBinMapping
                    '    frm = New frmItemRackBinMapping()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.LoanEntry
                    '    frm = New frmLoanEntry()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.LoanInstallment
                    '    frm = New frmLoanInstallmentEntry()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRequestAproval
                        frm = New frmRequestAproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '===============Greivance Logging ======================
                    Case clsUserMgtCode.frmGrievanceLogging
                        frm = New frmGrievanceLogging()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    '===============Greivance Allocation ======================
                    Case clsUserMgtCode.frmGrievanceAllocation
                        frm = New FrmGrievanceAllocation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    '===============Greivance Resolution ======================
                    Case clsUserMgtCode.frmGrievanceResolution
                        frm = New FrmGrievanceResolution()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================


                    '========================================================================================================================================
                    '==============================Employee Equipment Tracking==========================================================================
                    '===============Asset Category Master ======================
                    Case clsUserMgtCode.frmAssetCategoryMaster
                        frm = New FrmAssetTypeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    '===============Asset Sub Category Master ======================
                    Case clsUserMgtCode.frmAssetSubCategoryMaster
                        frm = New FrmAssetSubCategoryMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    '===============Asset Master======================
                    Case clsUserMgtCode.frmAssetMaster
                        frm = New FrmAssetDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==================================================
                    '===============Asset Issue Return======================
                    Case clsUserMgtCode.frmAssetIssueReturn
                        frm = New frmAssetsIssueReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '====================================================================================================================================
                    '=========================================================================================================================================

                    '===============Exit Management ======================

                    Case clsUserMgtCode.FrmHRSettings
                        frm = New FrmHRSettings()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmResignationLetter
                        frm = New FrmResignationLetter()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '======================HR Reports
                    Case clsUserMgtCode.RptRegisterOfDeduction
                        frm = New RptRegisterOfDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTerminationLetter
                        frm = New FrmHREXTerminationLetter()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmResignationAcceptanceOrRejection
                        frm = New FrmResignationAcceptanceORRejection()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHREMInterviewQuestion
                        frm = New FrmHREMInterviewQuestion()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHREMExitInterview
                        frm = New FrmExitInterview
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRPerformanceRatingRpt
                        frm = New rptPerformanceRating()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '===================End Exit Management===============================
                    Case clsUserMgtCode.FrmItemListRpt
                        frm = New FrmItemListRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVendorListRPT
                        frm = New FrmVendorListRPT()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ItemLocationDetails
                    '    frm = New frmItemLocationDetails(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.ItemReorderLevel
                        frm = New frmItemReorderLevel1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnItemCategory
                        frm = New FrmItemCategory1(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnItemSubCategory
                        frm = New FrmItemSubCategory(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemExciseMapping
                        frm = New FrmItemExciseMapping(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemBasicPrice
                        frm = New FrmItemBasicPrice(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStandardscheme
                        frm = New frmStandardscheme(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStandardRateItem
                        frm = New frmStandardRateItem(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmItemCategoryLevel
                        frm = New frmItemCategoryLevel(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "ITEM")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorCategoryLevel
                        frm = New frmItemCategoryLevel(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "VENDOR")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerCategoryLevel
                        frm = New frmItemCategoryLevel(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "CUSTOMER")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLocationCategoryLevel
                        frm = New frmItemCategoryLevel(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "LOCATION")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmItemCategoryStructure
                        frm = New frmItemCategoryStructure(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "ITEM")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorCategoryStructure
                        frm = New frmItemCategoryStructure(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "VENDOR")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerCategoryStructure
                        frm = New frmItemCategoryStructure(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "CUSTOMER")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLocationCategoryStructure
                        frm = New frmItemCategoryStructure(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, "LOCATION")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''richa 21/08/2014
                    Case clsUserMgtCode.FrmCatalogMaster
                        frm = New FrmCatalogMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBarCodeGenerator
                        frm = New FrmBarCodeGenerator()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBarCodeGenerator1
                        frm = New FrmBarCodeGenerator1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.WarrantyMaster
                        frm = New frmWarrentyMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSchemeMasterNew
                        ' Ticket No : KDI/22/08/18-000419 By Prabhakar : if Industry Type is D then show "Scheme master dairy screen" other then show Scheme Master in Material Module  when  Scheme Master  open.
                        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.INDUSTRYTYPE, clsFixedParameterCode.INDUSTRYTYPE, Nothing)), "D") = CompairStringResult.Equal Then
                            frm = New FrmSchemeMasterDairy()
                        Else
                            frm = New FrmSchemeMasterNew()
                        End If
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWeightConversion
                        frm = New FrmWeightCoversion()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInventoryAgeingReport
                        'frm = New frmStockAgeingReport
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        frm = New frmInventoryAgeingReportNew
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPriceGroupMapping
                        frm = New FrmPriceGroupMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTragetMaster
                        frm = New FrmTragetMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPrintProductInvoiceStatement
                        frm = New FrmPrintProductInvoiceStatement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptProductSaleRegister1
                        'frm = New RptProductSaleRegister1()
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.DispatchChecklist
                        frm = New FrmDispatchChecklist()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptProductBookingStatus
                        frm = New RptProductBookingStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptProductDispatchStatus
                        frm = New RptProductDispatchStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptProductDOStatus
                        frm = New RptProductDOStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptProductSaleOrderStatus
                        frm = New RptProductSaleOrderStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)




                    '------------------ Inventory Transactions---------------------------------------
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.adjust
                    '    frm=New FrmAdjustments1(lblUserCode.Text, objCommonVar.CurrentCompanyCode, "Adjustment Entry")
                    '    frm.Text = "Adjustment Entry"
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    Case clsUserMgtCode.mbtnEmptyTrans
                        frm = New frmAdjustmentEmpty()
                        'frm.Text = "Empty Transactions"
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnProductionEntry
                        frm = New frmAdjustmentProduction()
                        'frm.Text = "Production Entry"
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnStoreAdjustment
                        frm = New frmAdjustmentStore()
                        frm.Text = "Store Adjustment"
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStartBatchEntry
                        frm = New frmStartBatchEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmJobWorkInventory
                        frm = New frmJobWorkInventory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRawMilkConsumtion
                        frm = New frmRawMilkConsumption()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Transfer
                        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrentIndustryType, "D") = CompairStringResult.Equal Then
                            frm = New FrmTransferKDIL()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmTransferDCC()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.FrmTransferGateOut
                        frm = New FrmTransferGateOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmGeneralWeighment
                        frm = New frmGeneralWeighment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmRMProcessLoss
                        frm = New frmRMProcessLoss()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmAdjProductionEntry
                        frm = New frmAdjProductionEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAdjProductionEntryQC
                    '    frm = New frmAdjProductionEntryQCC()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAdjProductionStoreEntry
                    '    frm = New frmAdjProductionStoreEntry()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.TransferReturn
                        frm = New frmTransferKDILReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptFarmerSaleRegister
                        frm = New rptFarmerSaleRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.CreateTransfer
                        frm = New FrmTransfer3rdDoc()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Indent
                        frm = New frmIndent(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "transferEntry"

                    Case "adjustmentEntry"
                    Case clsUserMgtCode.FrmItemMcMapping
                        frm = New FrmItemMcMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWarehouseBreakage
                        frm = New FrmWarehouseBreakage()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmExpiryDateEntry
                        frm = New FrmExpiryDateEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPhysicalStock
                        frm = New FrmPhysicalStock()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPhysicalStockMultipleLocation
                        frm = New frmPhysicalStockMultipleLocation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ChangeItemSerialNumber
                        frm = New frmChangeSerialNumber()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemStockConversion
                        frm = New frmItemToItemStockConverion
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGSTDocumentList
                        frm = New frmGSTDocumentList
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ Inventory Reports---------------------------------------

                    'Case "ItemLocationReport"
                    'frm=New FrmLocationsReport()
                    '     formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.mbtnItemMovement
                        frm = New frmRptInventoryMovement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "mtbnTransfer"
                        frm = New frmRptTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ItemLocationDetailsReport
                    '    frm = New RptItemLocationDetailsNewVersion()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    Case clsUserMgtCode.KeyValue
                        frm = New FrmKeyvalueReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmLeakageBreakage
                        frm = New FrmLeakageBreakage()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.GatePass_Vs_actual
                        frm = New RptGatePassVSActual()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptInvoiceAgainstInward
                        frm = New RptInvoiceAgainstInward()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.ItemPrice
                        frm = New frmItemPrice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.StockRecoReport
                        frm = New FrmShippingStockreport1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmStockDispatchReport
                        frm = New FrmStockDispatchReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''Added By PANKAJ on 02/Counter/2011
                    Case clsUserMgtCode.mbtnStockAdjustmentReport
                        frm = New FrmStockAdjustmentReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAdjustmentStatusReport1
                        frm = New FrmAdjustmentStatusReport1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnBreakageReport
                        frm = New FrmBreakageReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BreakageReportSummary
                        frm = New FrmBreakageReportSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RoutewiseBreakageReport
                        frm = New FrmRoutewiseBreakageSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf clsCommon.CompairString(strFormName, "StockReport"
                    '    frm=New FrmStockReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.ReportTransfer
                        frm = New ReportTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "SamplingReportSummary"
                        frm = New FrmSamplingReportSummary1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SchemeReport
                        frm = New FrmSchemeReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.StockReportForFinishedGoods
                        frm = New FrmStockReportFinishedGoods()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAdjustmentReport
                        frm = New FrmAdjustmentReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVehicleWiseLoadout
                        frm = New frmVehicleWiseTransfe()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingIndentTransferReport
                        frm = New FrmPendingIndentTransferReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.stockRecoNew
                        frm = New FrmStockReco(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.UnionStockReport
                        frm = New rptUnionStock()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FATSNFGainLoss
                        frm = New frmFATSNFGainLoss(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductReceivingReport
                        frm = New rptProductReceivingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCustomerWiseStockReco
                        frm = New rptCustomerWiseStockReco()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemStockReport
                        frm = New ItemStockReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemRateStatusReport
                        frm = New ItemRateStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDailySTAFGSStatement
                        frm = New rptDailySTAFGSStatement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        'Case clsUserMgtCode.rptStockReport
                    '    frm = New rptStockReport()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.stockRecoNewJR
                        frm = New FrmStockReco(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJWParameterMaster
                        frm = New frmJWParameterMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJWFormulaMaster
                        frm = New frmJWFormulaMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJWVendorFormula
                        frm = New frmVendorFormulaMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWIItemPriceMaster
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.JWIRateofFGasPerRM, clsFixedParameterCode.JWIRateofFGasPerRM, Nothing)) = 1 Then
                            frm = New frmJWIItemPriceMaster()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        End If


                    Case clsUserMgtCode.frmJobWorkBillig
                        frm = New frmJobWorkBillig()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSRNJobWorkEstimate
                        frm = New FrmSRNJobWorkEstimate()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptJobworkMilkReceipt
                        frm = New rptJobworkMilkReceipt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptJobworkBilling
                        frm = New rptJobworkBilling()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptJobworkProductionReport
                        frm = New rptJobworkProductionReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptJWTankerStatusReport
                        frm = New rptJWTankerStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptJWDispatchReport
                        frm = New rptJWDispatchReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==============================================
                    Case clsUserMgtCode.frmSlotMaster
                        frm = New frmSlotMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDGMaster
                        frm = New frmDGMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailyElectricalEntry
                        frm = New frmDailyElectricalEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDailyElectricalEntryReport
                        frm = New rptDailyElectricalEntryReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '==============================================

                    'Eng. And Plant Management
                    Case clsUserMgtCode.frmSectionMasterEng
                        frm = New frmSectionMasterEng
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmConsumptionTypeMaster
                        frm = New frmConsumptionTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSectionConsumptionMapping
                        frm = New frmSectionConsumptionMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmParameterMasterEng
                        frm = New frmParameterMasterEng
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSectionParameterMapping
                        frm = New frmSectionParameterMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLogSheetEng
                        frm = New frmLogSheetEng
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.WorkRequisitionEng
                        frm = New frmPurchaseRequistion(strProgramCode)
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWorkEstimationEng
                        frm = New frmWorkEstimationEng()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.WorkOrderEng
                        frm = New frmPurchaseOrder(strProgramCode)
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmWorkOrderStatusEng
                        frm = New frmWorkOrderStatusEng()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.stockRecoBatch
                        frm = New FrmStockRecoBatch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmItemSerialTrackingReport
                        frm = New FrmItemSerialTrackingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTransferRegister
                        frm = New FrmTransferRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmFatSnfStockReport
                        frm = New FrmFatSNFStock
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDatewiseQtyFatSnfStockReport
                        frm = New FrmDatewiseMilkStock
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMilkLossGain
                        frm = New frmMCCMilkLossGain
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MeterialstockReco
                        frm = New FrmMeterialStockReco(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.RptUnpostingTransItemQty
                        frm = New RptUnpostingTransItemQty
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    ''''End
                    '------------------ Sales And Distribution Masters---------------------------------------


                    Case clsUserMgtCode.Sampling_Master
                        frm = New Sampling_Master()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.cash_Register_Details4
                        frm = New Cash_Register_Details4()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQuickSettlement
                        frm = New FrmQuickSettlement(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSettlementMaster
                        frm = New FrmSettlementMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.groupMasterRoute
                        frm = New frmRouteGroupMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.routeMaster
                        frm = New frmRouteMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.transportMaster
                        frm = New frmTransportMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode, clsUserMgtCode.transportMaster)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.transportMasterVendor
                        frm = New frmTransportMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode, clsUserMgtCode.transportMasterVendor)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.vhicleMaster
                        frm = New frmVehicleMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.channelCategory
                        frm = New frmchannelCategory(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.channelMaster
                        frm = New frmChannelMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.visiMaster
                        frm = New frmvisimaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.customerType
                    '    frm=New frmCustomerType(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    If funSetUserAccess("CUST-TYPE", frm) = False Then Exit Sub
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.priceMaster
                    '    frm=New FrmPriceMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.priceComponentMaster) = CompairStringResult.Equal  Then
                    '    frm=New FrmPriceComponantMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.priceComponentMapping
                    '    frm=New FrmPriceComponantMapping(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.customerCategory
                    '    frm=New frmCustomerCategory(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    If funSetUserAccess("CUST-CAT-M", frm) = False Then Exit Sub
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.customerMaster
                    '    frm=New frmCustomer(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.customerGroup
                    '    frm=New frmCustomerGroup(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    If funSetUserAccess("CUST-GRP-M", frm) = False Then Exit Sub
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.customerAccountSet
                    '    frm=New frmCustomerAccountSet(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    If funSetUserAccess("CUST-ACT-ST", frm) = False Then Exit Sub
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    Case clsUserMgtCode.frmCommissionMaster
                        frm = New FrmCommissionMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case "customer/OutletMaster"
                    'ElseIf clsCommon.CompairString(strFormName, clsUserMgtCode.PriceMaster) = CompairStringResult.Equal strFormName = "schemeMaster" Then
                    '    frm=New FrmSchmeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'Case clsUserMgtCode.ShiptoLocation

                    Case clsUserMgtCode.FrmAbateMentMaster
                        frm = New FrmAbateMentMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.transportType
                        frm = New FrmTransportType(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.routetypemaster
                        frm = New FrmRouteTypeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    ' Paravet Services

                    Case clsUserMgtCode.frmCattleType
                        frm = New FrmCattleTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBredType
                        frm = New FrmBredTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCattleColor
                        frm = New FrmCattleColorMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCattleMaster
                        frm = New FrmCattleMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmServiceGroup
                        frm = New FrmServiceGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.FrmServiceName
                    '    frm = New FrmServiceName
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.FrmServiceMaster
                        frm = New FrmServiceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmNDDBMaster
                        frm = New FrmNDDBMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBullMaster
                        frm = New FrmBullMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmFarmerServiceOrderWithRate
                        frm = New FrmFarmerServiceOrderWithRate
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSMSDetails
                        frm = New frmFarmerServiceOrder
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptActiveUsers
                        frm = New rptActiveUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmSalesManHierarchy
                    '    frm = New FrmSalesManHierarchy(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.mbtnBreakageHead1
                        frm = New FrmBreakagehead()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnCashDiscountReport
                        frm = New FrmCashDiscountReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerVendorMapping
                        frm = New FrmCustomerVendorMapping()
                        frm.formtype = clsUserMgtCode.CustomerVendorMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCheckPrinting
                        frm = New frmPrintCheckMultiple
                        'frm.formtype = clsUserMgtCode.CustomerVendorMappingVendor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNEFTUploader
                        frm = New FrmNEFTUploader(strProgramCode)

                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNEFTUploaderFarmer
                        frm = New FrmNEFTUploader(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmQuickPaymentBySingleCheque
                        frm = New frmQuickPaymentBySingleCheque
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.CustomerVendorMappingVendor
                        frm = New FrmCustomerVendorMapping()
                        frm.formtype = clsUserMgtCode.CustomerVendorMappingVendor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerEmployeeMapping
                        frm = New frmCustomerEmployeeMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.TDMTARGET
                        frm = New TDMwiseTarget()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRouteShifting
                        frm = New FrmRoute_Shifting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDiscountMaster
                        frm = New FrmDiscountMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDCSforSale
                        frm = New frmDCSforSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmFrieghtRateMaster
                        frm = New frmFrieghtRateMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDiscountCategoryMaster
                        frm = New FrmDiscountCategoryMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnTargetMaster
                        frm = New FrmTargetMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnTmplateCreation
                        frm = New FrmTemplateCreation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmViewPunchingInvoice
                        frm = New FrmViewPunchingInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDayWiseLoadOutEntered
                        frm = New FrmDayWiseLoadOutEntered()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerTargetFixing
                        frm = New frmCustomerTargetFixing()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmClaimMaster
                        frm = New frmClaimMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNShipmentImportExport
                        frm = New frmShipmentImportExport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRemarMaster
                        frm = New frmRemarkMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '' Zone Master
                    Case clsUserMgtCode.FrmZoneMaster
                        frm = New FrmZoneMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''
                    '------------------ Sales And Distribution Transactions---------------------------------------


                    ''----------------- Sales And Distribution Transaction NEW--------------------------------------
                    Case clsUserMgtCode.customerItemDetails
                        frm = New FrmCustomerItemDetails()
                        frm.isFromApprovalForm = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBookingEntry
                        frm = New FrmBookingEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmShortCloseDO
                        frm = New FrmShortCloseDO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmShortCloseDOPS
                        frm = New FrmShortCloseDOPS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmShortCloseDOCS
                        frm = New FrmShortCloseDOCS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDispatchFreshSale
                        frm = New frmDispatchNoteFreshSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDispatchMultipleFreshSale
                        frm = New frmDispatchMultipleFreshSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.GatePassTransfer
                        If clsCommon.CompairString(clsFixedParameter.GetData(clsFixedParameterType.CreateTransferFromBooking, clsFixedParameterCode.CreateTransferFromBooking, Nothing), "1") = CompairStringResult.Equal Then
                            frm = New FrmGatePassTransfer
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.TransferCrateReceived
                        frm = New frmCreateReceivedCustomerDairySale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCreditLimitApproval
                        frm = New FrmCreditLimitApprovalMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFreshCreditLimitApproval
                        frm = New FrmCreditLimitApprovalMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesLevelHierarchy
                        frm = New FrmSalesLevelhierarchy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesHierarchy
                        frm = New FrmsalesHierarchy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesHierarchyMapping
                        frm = New FrmSalesHierarchyMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBulkCreditLimitApproval
                        frm = New FrmCreditLimitApprovalMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInvoiceFreshSale
                        frm = New frmInvoiceFreshSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmInvoiceCrateLinerDetail
                        frm = New FrmInvoiceCrateLinerDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCreateReceived
                        frm = New frmCreateReceived
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalesOrderBS
                        frm = New FrmSalesOrderBS_Pavitra()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmGateEntrySale
                        frm = New FrmGateEntrySale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmWeighmentEntry
                        frm = New FrmWeighmentEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmMCCGateEntry
                    '    frm = New frmMCCGateEntry()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmMCCWeighment
                    '    frm = New frmMCCTankerWeighment()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLoadingTanker
                        frm = New FrmLoadingTanker()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmQualityCheckBulkSale
                        frm = New FrmQualityCheckBulkSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmTranReverse
                    '    frm = New frmTransactionReverse()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDispatchBulkSale
                        frm = New FrmDispatchBulkSale()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkSaleAcknowledgement
                        frm = New frmBulkSaleAcknowledgement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmInvoiceBulkSale
                        frm = New FrmInvoiceBulkSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCreateAutoInvoiceBS
                        frm = New FrmCreateAutoInvoiceBS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBulkSaleReturn
                        frm = New FrmBulkSaleReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCanSaleUploader
                        frm = New FrmCanSaleUploader()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCanSale
                        frm = New FrmCanSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCrateCanreceiptReport
                        frm = New frmCrateCanReceivingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCanReceived
                        frm = New frmCanReceived()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBulkDispatchReturnSale
                        frm = New FrmBulkDispatchReturnSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTankerOut
                        frm = New FrmTankerOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmDispatchBulkSaleTrade
                    '    frm = New FrmDispatchBulkSaleTrade()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmDispatchBulkSaleTradeReturn
                    '    frm = New FrmDispatchBulkSaleTradereturn()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFixedDeposit
                        frm = New FrmFixedDeposit()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProformaInvoiceMT
                        frm = New frmEXPorformaInvoice(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPurchaseOrderMT
                        frm = New frmPurchaseOrder(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSavingsMaster
                        frm = New frmSavingsMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSectionAllowanceMaster
                        frm = New frmSectionAllowanceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeSavingsMapping
                        frm = New frmEmployeeSavingsMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmBulkCloser
                    '    frm = New FrmBulkCloser()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesOrderMT
                        frm = New frmEXSalesOrder(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesOrderR
                        frm = New frmEXSalesOrderReturn(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCommercialInvoiceMT
                        frm = New frmEXCommercialInvoice(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSRNMT
                        frm = New frmSRN(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesInvoiceMT
                        frm = New frmEXSalesInvoice(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesReturnMT
                        frm = New frmEXSalesReturn(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLCRequest
                        frm = New FrmLCRequest()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLCCreation
                        frm = New FrmLCCreation()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDocumentAcceptance
                        frm = New FrmDocumentAcceptance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''shivani
                    Case clsUserMgtCode.RptCashAgainstDocs
                        frm = New RptCashAgainstDocs()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMTPIInHead
                        frm = New RptMTPIInHead()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptEXProductWiseDetail
                        frm = New RptMTProductWiseDetailReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptdateOfArrivalofCons
                        frm = New RptDateOfArrivalOfCon()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesCancelledTransation
                        frm = New frmCancelledTransactions_Exportsale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionCancelledTransation
                        frm = New frmCancelledTransactions_Production()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionCancelledTransation
                        frm = New frmCancelledTransactions_Production()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CancelledTransactionReportDS
                        frm = New frmCancelledTransactions_DairySale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBulkSalePriceChart
                        frm = New FrmBulkSalePriceChart()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPrintBulkInvoiceStatement
                        frm = New FrmPrintBulkInvoiceStatement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBulkSaleRegister
                        'frm = New RptBulkSaleRegister()
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)



                        '' changed by panch raj on 08-05-18 against ticket No: KDI/04/05/18-000295
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSubsidyCreditNote
                        frm = New frmSubsidyNote
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptSubsidyCredit
                        frm = New rptSubsidyCreditReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptSchemeDetail
                        frm = New RptSchemeReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptLeakageDetail
                        frm = New rptLeakageDetailsReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEnquiryMaster
                        frm = New FrmEnquiryMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCHAChargeMaster
                        frm = New FrmCHAChargeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmExIncentiveMaster
                        frm = New FrmExIncentiveMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNotifiedPartyMaster
                        frm = New FrmNotifiedPartyMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesQuotation
                        frm = New FrmEXSalesQuotation(clsUserMgtCode.frmEXSalesQuotation)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesOrder
                        frm = New frmEXSalesOrder(clsUserMgtCode.frmEXSalesOrder)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXPorformaInvoice
                        frm = New frmEXPorformaInvoice(clsUserMgtCode.frmEXPorformaInvoice)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXCommercialInvoice
                        frm = New frmEXCommercialInvoice(clsUserMgtCode.frmEXCommercialInvoice)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesInvoice
                        frm = New frmEXSalesInvoice(clsUserMgtCode.frmEXSalesInvoice)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEXSalesReturn
                        frm = New frmEXSalesReturn(clsUserMgtCode.frmEXSalesReturn)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptExportSaleRegister
                        'frm = New RptExportSaleRegister()
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '' changed by panch raj on 08-05-18 against ticket No: KDI/04/05/18-000295
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Ticket No-TEC/26/08/19-000996,Sanjay
                    Case clsUserMgtCode.frmPriceMasterPS
                        frm = New FrmCSAPriceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCSAPriceMaster
                        frm = New FrmCSAPriceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSACommissionItemWise
                        frm = New FrmCSACommissionItemWise()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSAAccountSet
                        frm = New FrmCSAAccountSet()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntry_JWO
                        frm = New frmJWOGateEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWO_Item_Acceptance
                        Dim settJobWorkOutwardComsumeItemAccordingToBOM As Boolean = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.JobWorkOutwardComsumeItemAccordingToBOM, clsFixedParameterCode.JobWorkOutwardComsumeItemAccordingToBOM, Nothing)) = 1)
                        If settJobWorkOutwardComsumeItemAccordingToBOM Then
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        Else
                            frm = New frmJWOGateEntry()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmWeighment_JWO
                        frm = New FrmMilkWeighment_JWO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQC_JWO
                        frm = New FrmMilkQualityCheck_JWO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnloading_JWO
                        frm = New FrmMilkUnloading_JWO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWO_SRN
                        frm = New FrmJWOSRN()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWO_SRN_Return
                        frm = New FrmJWOSRNReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJobWorkConsumption
                        frm = New frmJobWorkConsumption()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCSABooking
                        frm = New frmCSABooking(clsUserMgtCode.frmCSABooking)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSARequest
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowCSARequestScreen, clsFixedParameterCode.ShowCSARequestScreen, Nothing)) = 1 Then
                            frm = New frmCSABooking(clsUserMgtCode.frmCSARequest)
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("You are not authorize to access CSA Request.")
                        End If
                    Case clsUserMgtCode.frmCSADeliveryOrder
                        frm = New FrmCSADeliveryOrder
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSATransfer
                        frm = New frmCSATransfer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSASaleInvoice
                        frm = New FrmCSASaleInvoice
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSASalePattiReturn
                        frm = New FrmCSASalePattiReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSATransferReturn
                        frm = New FrmCSATransferReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSATransferReport
                        frm = New FrmCSATransferReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCSADOReport
                        frm = New FrmCSADOReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCSASaleRegister
                        'frm = New RptCSASaleRegister
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        '' changed by panch raj on 08-05-18 against ticket No: KDI/04/05/18-000295
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCSACustomerLedger
                        frm = New frmRptCSACustomerLedger(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPartyWiseSale
                        frm = New RptPartyWiseSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptGPDetail
                        frm = New RptGPDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptCSAmonthlywisereport
                        frm = New Frm_MW_SaleAnalysiReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''Test
                    Case clsUserMgtCode.FrmBulkSaleSettings
                        frm = New FrmBulkSaleSettings()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMultCustBookingDisp
                        frm = New FrmMultCustBookingDispatch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeliveryNoteFreshSale
                        frm = New frmDeliveryNoteFreshSale()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBookingProductSale
                        frm = New frmBookingProductSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSchemeMasterDairy
                        frm = New FrmSchemeMasterDairy()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRouteMaster
                        frm = New frmRouteMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmDistanceMappingMaster
                    '    frm = New FrmDistanceMappingMaster()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmSalesOrderProductSale
                        frm = New frmDeliveryOrderProductSale()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeliveryPrderProductSale
                        frm = New frmSaleOrderProductSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmShipmentProductSale
                        frm = New frmShipmentProductSale()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntryReturnPS
                        frm = New frmGateEntryReturnPS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntryReturnCS
                        frm = New frmGateEntryReturnCSA()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntryReturnTransfer
                        frm = New frmGateEntryReturnTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCSATransferGateOut
                        frm = New FrmCSATransferGateOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmProductDispatchGateOut
                        frm = New FrmProductDispatchGateOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSendBilltoDistributor
                        frm = New frmSendBilltoDistributor()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmdispatchAdviceProductSale
                        frm = New frmDispatchAdviceProductSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleInvoiceProductSale
                        frm = New frmSaleInvoiceProductSale(strProgramCode)
                        'frm = New frmSaleInvoiceProductSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleReturnProductSale
                        frm = New frmSaleReturnProductSale()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptsaleRegisterReport
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSalesHierarchyReport
                        frm = New rptSalesHierarchyReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.RptVehicleWiseReport
                    '    frm = New RptVehicleWiseReport
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.FrmGatePassFS
                        frm = New FrmGatePassFS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmGatePassPS
                        frm = New FrmGatePassPS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleReturnFreshSale
                        frm = New frmSaleReturnFreshSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRouteFreightDetails
                        frm = New FrmRouteFreightDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''Case clsUserMgtCode.FrmDispatchFreshSale
                    ''    frm=New FrmDispatchFreshSale()
                    ''         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmPendingQuotationApproval
                        frm = New FrmPendingQuotationApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmApprovalSetting
                        frm = New FrmApprovalSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalesmanTarget
                        frm = New FrmSalesmanTarget()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.customerItemDetails
                        frm = New FrmCustomerItemDetails()
                        frm.isFromApprovalForm = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerItemDetailApproval
                        frm = New FrmCustomerItemDetails()
                        frm.isFromApprovalForm = True
                        frm.Text = "Item Price List Approval"
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.customerItemMapping
                        frm = New frmCustomerItemMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmItemPriceListLevel3
                        frm = New FrmItemPriceListLevel3()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frm_User_Customer_Rate_Settings
                        frm = New Frm_User_Customer_Rate_Settings()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.routeMaster
                        frm = New frmRouteMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.routetypemaster
                        frm = New FrmRouteTypeMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleQuotation
                        frm = New frmSNSalesQuotation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmSNSalesOrder
                        frm = New frmSNSalesOrder()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNShipment
                        frm = New frmSNShipment()
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNSaleInvoice
                        frm = New frmSNSaleInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNServiceInvoice
                        frm = New frmSNServiceInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNSaleReturn
                        frm = New frmSNSaleReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAutoSTN
                        frm = New FrmAutoSTN()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmProspect
                    '    frm = New frmProspectDetail()
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    ''richa
                    Case clsUserMgtCode.FrmSaleSetting
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingFresh
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingBulk
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingMerchant
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMerchantPaymentTerms
                        frm = New FrmMerchantPaymentTerms()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMerchantPaymentTermsGroup
                        frm = New FrmMerchantPaymentTermsGroup()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMTReportContextFormat
                        frm = New FrmMTReportContextFormat()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingExport
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingCSA
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleSettingProduct
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''=================
                    Case clsUserMgtCode.frmSNPOS
                        frm = New frmSNPOS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSNReceiptChallan
                        frm = New frmReceiptChallan()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSalesmanTarget
                        frm = New frmRptSalesmanTarge()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleOrderDetail
                        frm = New frmSaleOrderDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleOrderSummary
                        frm = New FrmSaleOrderSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleInvoiceSummary
                        frm = New FrmSaleInvoiceSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleInvoiceDetail
                        frm = New FrmSaleInvoiceDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmShipmentDetail
                        frm = New FrmShipmentDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmShipmentSummary
                        frm = New FrmShipmentSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleRegisterDemo
                        ' frm = New FrmSaleRegisterDemo()
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSaleRegisterForAdv
                        ' frm = New FrmSaleRegisterDemo()
                        frm = New RptSaleRegisterReportForAdv(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSalesCancelledRegister2
                        frm = New RptSalesCancelledRegister2
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISSaleRegisterWithCSASalePatti
                        'frm = New RptSaleRegisterReportWithCSASalePatti(strProgramCode)
                        '' changed by panch raj on 02-05-18 against ticket No: UDL/27/04/18-000143
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptSKUWiseSale
                    '    frm = New rptSKUWiseSale(strProgramCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.MISSaleRegisterWithCSASalePattiProductLocationWise
                        frm = New RptSaleRegisterReportWithCSASalePattiProductLocationWise(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISSaleRegisterWithCSASalePattiProductPackWise
                        frm = New RptSaleRegisterReportWithCSASalePattiProductPackWise(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPendingSaleInvoiceforChilpPO
                        frm = New frmPendingSaleInvoiceforChilpPO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ORDNEW
                        frm = New FrmOrdertracking()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.mbtnItemMovement
                    '    frm = New frmRptInventoryMovement()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmReceiptChallanReport
                        frm = New frmReceiptChallanReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmProspectDetailReport
                    '    frm = New frmProspectDetailReport()
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case "quotation"
                    'Case clsUserMgtCode.rptDetailedCardReport
                    '    frm = New rptDetailedCardReport()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.saleOrders
                        frm = New frmSaleOrder()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCheckSlipEntry
                        frm = New FrmCheckSlipEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.LoadOut
                    '    frm = New frmShipmentInvoice()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "saleInvoice"
                    ''frm=New FrmSaleInvoice(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    ''     formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.saleReturn
                        frm = New frmSalesReturnNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SaleReturnInterCompany
                    'frm = New frmSaleReturnInter()
                    'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.FrmCompleteTransfer
                        frm = New FrmCompleteTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LoadOutStatus
                        frm = New FrmCompleteLoadout(lblUserCode.Text, lblCompany.Text)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ScrapSale
                        frm = New frmScrapSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.JobWorkDispatch
                    '    frm = New frmJobWorkDispatch()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JobWorkDispatchProduction
                        frm = New frmJobWorkDispatch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmScrapSaleGateOut
                        frm = New FrmScrapSaleGateOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case "ScrapInvoice"
                        frm = New frmScrapInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSettlementEntry
                        frm = New FrmSettlementEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTransferIncompleteRemarks1
                        frm = New FrmTransferIncompleteRemarks1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "FrmReverseEntry"
                        Dim frmP As New FrmPWD(Nothing)
                        frmP.strCode = "TempProvisional"
                        frmP.strType = "TempProvisional"
                        frmP.ShowDialog()
                        If frmP.isPasswordCorrect Then
                            Try
                                frm = New FrmReverseEntry()
                                formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                            Catch ex As Exception
                                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                            End Try
                        End If

                    Case clsUserMgtCode.ScrapSaleRetrun
                        frm = New frmScrapSaleReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ChangeInvoiceSalesman
                        frm = New FrmChangeInvoiceSalesman()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.frmtransfer
                    '    frm=New frmTransfer(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.FrmReceipt
                    '    frm=New FrmReceipt(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmSaleHistory
                        frm = New FrmSaleHistory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmExpiryDate
                        frm = New FrmExpiryDate()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTransactionApproval
                        frm = New FrmTransactionApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFreshTransactionApproval
                        frm = New FrmTransactionApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBulkTransactionApproval
                        frm = New FrmTransactionApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMCCFarmerMapping
                        frm = New FrmMCCFarmerMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VLCMappingForMPAmount
                        frm = New frmVLCMappingForMPAmount()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VSP_Commsission
                        frm = New frmVSPCommission()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VSP_Deduction
                        frm = New frmVSPDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VSP_DayWiseIncetive
                        frm = New frmVSPDayWiseIncentive()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VSP_Mapping
                        frm = New frmVSPMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FarmerProMaster
                        frm = New frmFarmerProMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CappingMaster
                        frm = New frmCappingMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DCSFinancialHead
                        frm = New frmDCSFinancialHead()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DCSFinancialEntry
                        frm = New frmDCSFinancialEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DCSSupervisorTagging
                        frm = New frmDCSSupervisorTagging()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MPIncetiveSlab
                        frm = New frmMPIncetiveSlab()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBMCCategory
                        frm = New FrmBMCCategory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ChillingChargesSlab
                        frm = New frmChillingChargesSlab()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmOwnBmcExpanse
                        frm = New FrmOwnBmcExpanse()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.OwnBMCGainLossRate
                        frm = New frmOwnBMCGainLossRate()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VLCMappingForMP_PP
                        frm = New frmVLCMappingForMP_PaymentProcess()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FreightChargesMaster
                        frm = New frmFreightChargesMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMCCTransactionApproval
                        frm = New FrmTransactionApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPrintFreshInvoice
                        frm = New FrmPrintFreshInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ' Case clsUserMgtCode.RptFreshSaleRegister1
                    'frm = New RptFreshSaleRegister1()
                    'formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptFreshSaleRegister1
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCrateAccounting
                        frm = New rptCrateAccounting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCrateAccountingReport
                        frm = New RptCrateAccountingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.RptFreshBookingStatus
                        frm = New RptFreshBookingStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)



                    Case clsUserMgtCode.rptZoneWiseFreshSaleReport
                        frm = New RptZoneWiseFreshSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDispatchChallanReportFresh
                        frm = New RptDispatchChallanReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMatrixFreshSalesReport
                        frm = New RptMatrixFreshSalesReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMatrixFreshSalesReportSaleDairy
                        frm = New RptMatrixFreshSalesReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''----Parteek ---''
                    Case clsUserMgtCode.rptPriceChartFreshSalesReport
                        frm = New frmRptPriceChartMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''---End---''
                    Case clsUserMgtCode.rptSaleReturnGateEntryReport
                        frm = New RptSaleReturnGateEntryReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCrateLinerReport
                        frm = New rptCrateLinerReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptVehicleCapacityFreshSaleReport
                        frm = New RptVehicleCapacityFreshSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBulkMultipleDispatch
                        frm = New RptBulkMultipleDispatch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ Sales And Distribution Report---------------------------------------

                    Case clsUserMgtCode.vehicle_Details_Report1
                        frm = New Vehicle_Details_Report()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptTransfer_IncompleteReport
                        frm = New RptTransfer_IncompleteReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.reportQuickSettlement
                    '    frm = New FrmReportForQuickSettlement(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.visiDetail1
                        frm = New VisiDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PendingSaleOrderReport
                        frm = New PendingSaleOrderReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCDiscountReport
                        frm = New FrmMCDiscReport(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.crptLoadOut
                    '    frm = New FrmLoadOutRpt(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.receiptFillreport
                        frm = New FrmRECEIPTSAGAINSTSALES_FILLED_()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.receiptWOTreport
                        frm = New FrmInvoiceswithoutreceipt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.TransporterRpt
                    '    frm=New frmTransportMasterRpt(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.VehicleMasterRpt
                    '    frm=New FrmVehicleMasterRpt(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.receiptreport
                        frm = New Frmreceiptvoucher2(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.rptCustomerGroupDetails
                    '    frm=New FrmCustomerGroupReport()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'ElseIf  clsCommon.CompairString(strFormName, clsUserMgtCode.frmCustomerDetails
                    '    frm=New FrmCustomerMasterReport()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)

                    Case clsUserMgtCode.nrptSales
                        frm = New FrmRptSales()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnRptSalesManSalesReport
                        frm = New frmRptSalesManReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Settlement
                        frm = New FrmSettlementReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'ElseIf clsCommon.CompairString(strFormName, "ProvSales"
                    '    frm=New FrmProvionalSalesReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.CustomerRouteHistoryReport
                        frm = New FrmRptCustomerRouteHistory(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemDiscountReport
                        frm = New FrmDiscountReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "ItemCommissionReport"
                        frm = New FrmItemCommissionReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LoadOutStatusreport1
                        frm = New FrmLoadOutStatusreport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LoadOutReport1
                        frm = New LoadOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "EmptyInwardRegisterSummary"
                        frm = New FrmEmptyInwardRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnNetSaleReport
                        frm = New FrmNetSaleReport1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptNetSaleDetailReport
                    '    frm=New frmRptNetSaleDetailReport()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.FrmDistrbutorSaleTarget
                        frm = New FrmDistrbutorSaleTarget()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDayReportDirectSale
                        frm = New FrmDayReportDirectSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.ExciseSummary1
                    '    'If objCommonVar.IsDemoERP Then
                    '    '    frm = New FrmExciseSummary_DEMO()
                    '    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    'ElseIf clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "Guntur"
                    '    '    frm = New FrmExciseSummaryNew()
                    '    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    'Else
                    '    '    frm = New FrmExciseSummaryReport()
                    '    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    'End If
                    '    frm = New FrmExciseSummaryNew()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.OtherPartySale
                        frm = New FrmOtherPartySale1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTransitBreakageReport1
                        frm = New FrmTransitBreakageReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCreditSaleReport
                        frm = New frmRptCreditSales()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.NoSaleReport
                        frm = New frmRptNoSales()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPenetration
                        frm = New frmRptPenetration()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FilloutwardRegisterReport1
                        frm = New FrmFilledOutwardRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SaleReport
                        frm = New FrmTDMSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PrimarySales
                        frm = New FrmPrimarySalesReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SecondarySales
                        frm = New FrmSecondarySales(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.EmptyInwardSaleRegister1
                        frm = New frmInwardRegister(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ProvSaleDetail
                        frm = New FrmProvisionalSalesRoutewise(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmGatePassENtry1
                        frm = New FrmGatePassENtry1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SaleAccountBreakDetail
                        frm = New FrmSaleAccountBreakOrCashDisc(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SaleAccountBreakage
                        frm = New FrmSaleAccountBreakageReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "DailyStockAccount"
                        frm = New FrmDailyStockAccountRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "OverAllDisc"
                        frm = New FrmOverallDiscountReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CrptRG1Detail1
                        If objCommonVar.IsDemoERP Then
                            frm = New frmRG1Demo()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmRG1()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmCFormReport
                        If objCommonVar.IsDemoERP Then
                            frm = New FrmCFormReport()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmOpeningBalance
                        frm = New FrmOpenningBalance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptInventoryMovement
                        frm = New RptInventoryMovement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptItemWiseTaxMasterReport
                        frm = New rptItemWiseTaxMasterReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDasboard
                        frm = New FrmDasboard()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDasboardCombine
                        frm = New frmDB()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDocumentCancelledReport
                        frm = New frmDocumentCancelledReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.DashboardMilkProcurement
                        frm = New DashboardMilkProcurement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.JanaadharStatusReport
                        frm = New JanaadharStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptDataEntryTracingReport
                    '    frm = New rptDataEntryTracingReport()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.RptServiceTaxDetail
                    '    frm = New RptServiceTaxDetail()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBankWiseChequeIssue
                        frm = New RptBankWiseChequeIssue()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ProvisionalSaleReport
                    '    frm = New FrmProvionalSalesReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ItemCommissionSummary
                        frm = New FrmItemCommissionSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DistributedDiscountReport
                        frm = New FrmDistribuorDiscount(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleDiscount1
                        frm = New FrmSaleDiscount1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CEAllocationReport
                        frm = New FrmCEAllocationRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SalesCollection
                        frm = New FrmSalesCollection(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DailySettlement
                        frm = New frmDailySettlement(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMismatchSettlement
                        frm = New FrmMismatchSettlement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSettlementSheetReconcilationeport
                        frm = New FrmSettlementSheetReconcilationReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.OutletEmpty1
                        frm = New FrmOutletEmptyReport1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VehiclewiseSale1
                        frm = New FrmVehiclewiseSale(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Channelwisecustomer1
                        frm = New FrmChannelwiseCustomer1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnCustomerRanking
                        frm = New FrmCustomerRankingReport1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnVisiVPO1
                        frm = New FrmVisiVPOReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnMismatchReport
                        frm = New FrmMismatchRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnRouteSale
                        frm = New RouteSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerTargetReport
                        frm = New frmCustomerTargetReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmTDMReport
                        frm = New frmTDMReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SalesmanSalesOrderReport
                        frm = New FrmSalemanSaleOrder()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleAnalysisReport
                        frm = New frmSaleAnalysisReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDealerManagementReport
                        frm = New FrmDealerManagementReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSalesAnalysis
                        frm = New frmSalesAnalysisReport(strUserCode, strCompany)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRptFormOfGuarntee
                        frm = New frmRptFormOfGuarntee
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCustomerWiseMonthlySalesAnalysis
                        frm = New frmCustomerWiseMonthlySalesAnalysis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPaySlipReport
                        frm = New frmCheckDepositPaySlip
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalarySlipRpt
                        frm = New FrmSalarySlipRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalarySummaryRpt
                        frm = New FrmSalarySummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmEmployeePFRpt
                        frm = New FrmEmployeePF
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmESICRpt
                        frm = New FrmESICRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptForm34
                        frm = New RptForm34
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBonusStatement
                        frm = New RptBonusStatement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptBOILetterReport
                        frm = New RptBOILetterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptForm22
                        frm = New RptForm22
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.EmployeeWiseReport
                        frm = New EmployeeWiseReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptESIHalfYearly
                        frm = New RptESIHalfYearly
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFForm3A
                        frm = New RptPFForm3A
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFForm5
                        frm = New RptPFForm5
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFForm10
                        frm = New RptPFForm10
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RPtPFForm11_Revised_
                        frm = New RPtPFForm11_Revised_
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFForm12A_revised_
                        frm = New RptPFForm12A_revised_
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFChallanStatement
                        frm = New RptPFChallanStatement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptKDILSalarySlip
                        frm = New RptKDILSalarySlip
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPerformaForContributiondetail
                        frm = New RptPayrollPerformaforcontribution
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmPaymentMode
                        frm = New FrmPaymentMode
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptBankTransferDetail
                        frm = New RptBankTransferDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmDepartmentwiseSalarySheetRpt
                        frm = New RptDepartmentWiseSalarySheet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDetailOfWelfareFundAmount
                        frm = New RptDetailOfWelfareFundAmount
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.RptESICStatement
                        frm = New RptESICStatement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPFStatement
                        frm = New RptPFStatement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPFForm6
                        frm = New RptPFForm6
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptESICForm6
                        frm = New RptESICForm6
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptESICChallan
                        frm = New RptESICChallan
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptESICDeclarationForm
                        frm = New RptESICDeclarationForm
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptActurialValuation
                        frm = New RptActurialValuation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptAttendanceReport
                        frm = New RptAttendaceReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBiometricAttendanceReport
                        frm = New rptBiometricAttendanceReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptEmployeeStatusReport
                        frm = New rptEmployeeStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPFForm2
                        frm = New RptPFForm2
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDVAT32
                        frm = New frmRptDVAT32
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerBillWiseDetail
                        frm = New frmCustomerBillWiseDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleSummaryAgainstPO
                        frm = New frmSaleSummaryAgainstPO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmYearMonthWiseSaleComparison
                        frm = New frmYearMonthWiseSaleComparison()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCompanyMonthWiseSaleComparison
                        frm = New frmCompanyMonthWiseSaleComparison()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDistributerPerformance
                        frm = New rptDistributerPerformance(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCustomerBillWiseDuesSummary
                        frm = New frmCustomerBillWiseDuesSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmVendorGroupWiseSaleReport
                        frm = New frmVendorGroupWiseSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBatchReceiptSTD
                        frm = New FrmBatchReceipt(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBatchReceiptPepsi
                        frm = New FrmBatchReceipt(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFilledOutWard
                        frm = New FrmFilledOutWard()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CashDiscount
                        frm = New rptCashDiscountReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CashDiscountReport
                        frm = New frmCashDiscountNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TransferRegister
                        frm = New TransferRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.EmptyReportDetail
                        frm = New EmptyReportDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPendingSettlement
                        frm = New RptPendingSettlement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDetailOfForm2A
                        frm = New FrmDetailOfForm2A()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTargetReport1
                        frm = New FrmTargetReport1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDailySettlementActualAndProvisionalReport
                        frm = New FrmDailySettlementActualAndProvisionalReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSettlement_CashMemoStatus
                        frm = New FrmSettlement_CashMemoStatus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReverseSettlementDetail
                        frm = New FrmReverseSettlement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMismatchCashMemo
                        frm = New FrmMismatchCashMemo()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCanceledSaleInvoice
                        frm = New FrmCanceledSaleInvoice1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleVolumeTracker
                        frm = New FrmSaleVolumeTracker()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmEmptyTransactionReport
                        frm = New FrmRptEmptyTransaction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleOrderSummary
                        frm = New FrmSaleOrderSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDiscountAnalysis
                        frm = New FrmDiscountAnalysis()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '''''' Added BY Abhishek a s on 3 Nov 2012---
                    Case clsUserMgtCode.frmpendingLoadin
                        frm = New FrmPendingLoadIn_Transfer_Type
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptSalesReturn
                        frm = New FrmRptSalesReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmQuickSettlementHead
                        frm = New FrmQuickSettlementHead
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmVendorsOutstandings
                    '    frm = New FrmVendorsOutstandings
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.LO_vs_Vechile
                        frm = New FrmloadoutVSvechileCapacity2
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptClaimMaster
                        frm = New rptClaimMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVehiclePendingStatusRpt
                        frm = New FrmVehiclePendingStatusRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '--------------------TDS Master-----------------------------

                    Case clsUserMgtCode.NatureOfDeduction
                        frm = New frmNatureOfDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PartyDetails
                        frm = New frmPartyDetails(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TDSSection
                        frm = New frmTDSSection(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BranchDetails
                        frm = New frmBranchDetails(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FinancialYear
                        frm = New frmFinancialYear(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.StateCode
                        frm = New frmStateCode(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ResponsiblePerson
                        frm = New frmResponsiblePerson(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnCreateRemittance
                        frm = New FrmCreateRemittance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.remittanceentry
                        frm = New Frmremittanceentry(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnAPInvoiceEntryTDS
                        frm = New FrmAPInvoiceEntryTDS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcurementDeduction
                        frm = New FrmAPInvoiceEntry(clsUserMgtCode.frmProcurementDeduction)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMultipleProcDeduction
                        frm = New FrmMultipleProcDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '-------------------------TDS Report -----------------------------------

                    Case clsUserMgtCode.frmrptTDSLedger
                        frm = New FrmrptTDSLedger()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TDSForm26Q
                        frm = New form26Q27Q()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TDSSectionSummaryReport
                        frm = New FrmTDSsectionSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Form16AReport
                        frm = New FrmForm16A()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ Purchase Order Master---------------------------------------
                    Case clsUserMgtCode.frmPurchaseSetting
                        frm = New frmPurchaseSettings
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPurchaseRegisterReport
                        frm = New RptPurchaseRegisterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorItemDetails
                        frm = New frmVendorItemDetails(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ Purchase Order---------------------------------------
                    Case clsUserMgtCode.mbtnPurchaseRequistion
                        frm = New frmPurchaseRequistion(strProgramCode)
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RMDemanApproval
                        frm = New frmRMDemandApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmStoreRequistion
                        frm = New frmStoreRequistion()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.mbtnPendingApprovalOfReq
                        frm = New FrmPendingReqForApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RequisitSubTypeMaster
                        frm = New FrmRequisitSubTypeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RFQ
                        frm = New FrmRFQ()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorQuotation
                        frm = New frmVendorQuotation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorComparison
                        frm = New FrmVendorComparison1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorComparisonApproval
                        frm = New frmVendorComparisonApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnPurchaseOrder
                        frm = New frmPurchaseOrder(strProgramCode)
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.SetPOSchedule
                    '    frm = New frmSetPOSchedule()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.mbtnPurchaseOrder
                    '    frm = New frmPurchaseOrder()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmPurchaseSchedule
                        frm = New FrmPurchaseSchedule()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnMRN
                        frm = New frmMRN()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkSaleAcknowledgementUploader
                        frm = New frmBulkSaleAcknowledgementUploader()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BulkSaleFreightMaster
                        frm = New frmBulkSaleFreightMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkSaleFreightCalculation
                        frm = New frmBulkSaleFreightCalculation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.NIRQC
                        frm = New frmNIRQC()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnSRN
                        frm = New frmSRN(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PurchaseGateOut
                        frm = New PurchaseGateOut()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SRNReturn
                        frm = New frmSRNReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnPurchaseInvoice
                        frm = New frmPurchaseInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.mbtnPurchaseReturn
                        frm = New frmPurchaseReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPDAccountSummariesReport
                        frm = New frmPDAccountSummariesReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'sanjay BHA/09/05/18-000014 
                    Case clsUserMgtCode.frmMaterialQuotation
                        frm = New frmMaterialQuotation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMaterialQuotationOrder
                        frm = New frmMaterialQuotationOrder()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMaterialQuotationComparison
                        frm = New frmMaterialQuotationComparison()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    'KUNAL > CLIENT : UDIL > TICKET : BM00000010226 
                    Case clsUserMgtCode.mbtnNRGP
                        frm = New frmNRGPBooking(clsUserMgtCode.mbtnNRGP)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.mbtnGatePass
                        frm = New frmRGP()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnIssueReturn
                        frm = New frmIssueReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRoutewiseSaleReport
                        frm = New FrmRoutewiseSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmKPIReport
                        frm = New FrmKPIReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProvSaleExcel
                        frm = New FrmProvSaleExcel(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    ''richa
                    'Case clsUserMgtCode.frmItemQuantityInformation
                    '    frm=New FrmItemQuantityInformation()
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '------------------------Purchase Order Report -------------------------------

                    Case clsUserMgtCode.RM_Consumption_Detail
                        frm = New RM_Consumption_Detail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMaterialSendforJobWork
                        frm = New RptMaterialSendForJobWork()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMaterialReceivedAfterJobWork
                        frm = New RptMaterialReceivedAfterJobWork()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBalanceStockForJobWork
                        frm = New RptBalanceStockForJobWork()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptRGPWiseJobWork
                        frm = New RptRGPWiseJobWork()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmIndentReport
                        frm = New FrmIndentReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmExpiredItemDetails
                        frm = New FrmExpiredItemDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmConsumptionReport1
                        frm = New FrmConsumptionReport1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DebitAdviseReport
                        frm = New FrmDebitAdviseReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Vendor_Rating_Rejection
                        frm = New Vendor_Rating_Rejection()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Store_Receipt_Note
                        frm = New Store_Receipt_Note()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnDailyRcptNoteSummary
                        frm = New FrmDailyReceipNoteSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Parti_VS_Rejected
                        frm = New Parti_VS_Rejected()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPurchaseOrderReport
                        frm = New FrmPurchaseOrderReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPo_action
                        frm = New frmPo_action()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingRequisitionQty
                        frm = New FrmPendingRequisitionQty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "FrmPendingPO_Qty"
                        frm = New FrmPendingPO_Qty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingGrn_Qty
                        frm = New FrmPendingGrn_Qty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingMrn_Qty
                        frm = New FrmPendingMrn_Qty
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingSrn_Qty
                        frm = New FrmPendingSrn_Qty
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case "FrmPendingInvoice_Qty"
                        frm = New FrmPendingInvoice_Qty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MRDAReport
                        frm = New FrmMRDAReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmIssueOrReturnItemWiseSummary
                        frm = New FrmIssueOrReturnItemWiseSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPurchasebookReport1
                        frm = New FrmPurchasebookReport1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStockAnalysis
                        frm = New FrmStockAnalysis1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMorningReport
                        frm = New FrmMorningReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AddCharge
                        frm = New FrmAdditionalCharge1(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSrnReport
                        frm = New FrmSrnReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PJVReport
                        frm = New FrmPJVReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.StockStatement
                        frm = New FrmStockStatementReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPurchaseOrderRegister
                        frm = New FrmPurchaseOrderRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmItemWiseDispatchLedger3
                        'frm = New FrmItemWiseDispatchLedger3()
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '' changed by panch raj on 08-05-18 against ticket No: KDI/04/05/18-000295
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnRGP_NRGP_Rpt
                        frm = New frmRGP_NRGP_Rpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DetailofWtdPriceofRawMaterial
                        frm = New FrmWTDRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorWiseReturnableGoodBalance
                        frm = New FrmVendorWiseReturnableGoodBalance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnStoreLedger
                        frm = New FrmStoresLedgerNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnStoresLedger
                        frm = New FrmStoresLedger()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRGP_Register_NRGP
                        frm = New FrmRGP_Register_NRGP()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'KUNAL > TICKET : BM00000010298 > CLIENT : UDL > DATE : 28-NOV-2016
                    Case clsUserMgtCode.FrmRpt_OutStnd_Items_RGP
                        frm = New FrmRpt_OutStnd_Items_RGP()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmFreightCosting
                        frm = New FrmFreightCosting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptPurchaseReturnBook
                        frm = New FrmRptPurchaseReturnBook()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPurchaseOrderList
                        frm = New frmPurchaseOrderList()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'added by stuti on 01/03/2017 for kdil
                    Case clsUserMgtCode.frmPurchaseOrderAmd
                        frm = New frmPOAmendmentReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.SRNReturnListCancellation
                    '    frm = New frmSRNReturnListForCancellation()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCapexRegister
                        frm = New rptCapexPurchaseRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptDayWisePurchasePriceReport
                        frm = New RptDayWisePurchasePriceReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptrlPenaltyRegister
                    '    frm = New rptrlPenaltyRegister()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPurchasePlanReport
                        frm = New RptPurchasePlanReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptPurchaseMaterialRegister
                        frm = New RptPurchaseMaterialRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.mbtnGRN
                    '    frm = New frmGRN
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnGRN
                        frm = New frmGRN
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VisualRandomQC
                        frm = New frmGRN
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    'Pankaj------------------------GRN Report ------------------------------------------

                    Case clsUserMgtCode.POWeighment
                        frm = New frmPOWeighment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.POUnloading
                        frm = New frmPOUnloading
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnGRNReport
                        frm = New FrmGRNReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptRejectionReport
                        frm = New rptRejectionReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CancelPurchaseInvoiceReport
                        frm = New CancelPurchaseInvoiceReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRALQCReport
                        frm = New frmRALQCReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPIPaymentStatusReport
                        frm = New frmPIPaymentStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptUnpostedPO
                        frm = New frmRptUnpostedPO()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'DATE : 14-FEB-2017 , CLIENT : UDL -- MONTHLY CONSUMPTION REPORT  =====
                    Case clsUserMgtCode.frmMonthlyConsumptionReport
                        frm = New frmMonthlyConsumptionReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    '--------Added By--Pankaj Kumar----------------Fixed Assets----------------------------------------------
                    '================================Master======================================
                    Case clsUserMgtCode.fixedsetting
                        frm = New FrmFixedSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AssetSegment
                        frm = New FrmAssetSegment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Template
                        frm = New FrmTemplateMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DepAccSets
                        frm = New FrmDepAccountSet()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Categories
                        frm = New FrmCategories()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDepreciationField
                        frm = New FrmDepreciationField()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetGroups
                        frm = New FrmGroups()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetBookMaster
                        frm = New frmAssetBookMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDepreciationMethod
                        frm = New frmDepreciationMethod()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DepPeriod
                        frm = New FrmDepreciationPeriods()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAMAcquisitionCode
                        frm = New FrmAMAcquisitionCode()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FAMergeAcquisitionEntry
                        frm = New frmFAMergeAsset()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetAccountChange
                        frm = New frmAssetAccountChange()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetDispatchRetailer
                        frm = New frmAssetDispatchRetailer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSecondaryCustomer
                        frm = New FrmSecondaryCustomer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmConsumerDetailsForm
                        frm = New frmConsumerMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==============================Transaction===================================
                    Case clsUserMgtCode.FAAcquisitionEntry
                        frm = New frmAcquisionEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FAAssetDepreciation
                        frm = New FrmAssetDepreciation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FADisposalEntry
                        frm = New frmAssetScrapSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVisi_Install_Pullout
                        frm = New FrmVisi_Install_Pullout()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAsset_Issue_Return
                        frm = New FrmAsset_Issue_Return()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetRequisition
                        frm = New frmAssetRequisition(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetStoreRequistion
                        frm = New frmAssetStoreRequistion()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSecondaryCustomerSale
                        frm = New FrmSecondaryCustomerSale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIssueItemsToAsset
                        frm = New frmItemIssueToAssembledAsset
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FAAssetWork
                        frm = New frmAssetWork
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '===============================Report=======================================
                    Case clsUserMgtCode.FrmAssetRegister
                        frm = New FrmAssetRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAssetDetail
                        frm = New FrmAssetDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDisposalDetail
                        frm = New FrmDisposalDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVisi_Install_Pullout_Report
                        frm = New FrmVisi_Install_Pullout_Report()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAsset_Issue_Return_Report
                        frm = New FrmAsset_Issue_Return_Report()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDistributor_VS_SecondaryCustomer_Sale
                        frm = New FrmDistributor_VS_SecondaryCustomer_Sale()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptFARReport
                        frm = New RptFAFARReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCWIPReport
                        frm = New rptCwipReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptALCRReport
                        frm = New rptALCRReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCapexConsumptionRpt
                        frm = New rptCapexConsumptionReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '-------------------------------------Fixed Assets(Ends Here)--------------------------------------------


                    'Dipti----------------------Utility---------------------------------------------------------
                    'Case clsUserMgtCode.mbtnCreateReceiptAgainstSale
                    '    frm = New FrmCreateReceiptAgainstSales()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.mbtnCreateReceiptAgainstInvoice
                    '    frm = New FrmCreateReceiptAgainstInvoice()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnTakeBackup
                        frm = New FrmBackup()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnRestoreDB
                        Dim strMsg As String = "Your current tasks will be closed." + Environment.NewLine
                        strMsg += "Do you want to continue?"
                        If clsCommon.MyMessageBoxShow(strMsg, Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes Then
                            RadDock1.RemoveAllDocumentWindows()
                            SplitPanel3.Collapsed = True
                            SplitPanel2.Collapsed = True
                            SplitPanel3.Collapsed = True
                            SplitPanel4.Collapsed = False
                        End If
                    Case "calc"
                        System.Diagnostics.Process.Start("calc.exe")
                    Case clsUserMgtCode.mbtnPendingApproval1
                        frm = New FrmPendingAproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDisplaySquenece
                        frm = New frmDisplaySquenece()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBlockMaster
                        frm = New frmBlockMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmRevenueVillageMaster
                        frm = New frmRevenueVillageMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGrampanchayatMaster
                        frm = New frmGrampanchayatMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPanchayatSamitiMaster
                        frm = New frmPanchayatSamitiMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVidhanSabhaMaster
                        frm = New frmVidhanSabhaMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRequestMaster
                        frm = New frmRequestMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkPostingNew
                        frm = New FrmBulkPostingNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.lockTransaction
                        frm = New FrmLockTransaction1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUserPerformanceDetail
                        frm = New FrmUserPerformanceDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStockReport
                        frm = New FrmStockReport(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmReconciliationSetting
                        frm = New FrmReconciliationSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.FrmUtilityForm
                        frm = New frmPWDHighSecrity(Nothing)
                        frm.ShowDialog()
                        If frm.isPasswordCorrect Then
                            Dim frmNew As New FrmUtility()
                            formShow(frmNew, strProgramCode, strProgramName, True, "")
                        End If
                    '==============================PAYROLL===================================
                    '===============================Setup=======================================
                    Case clsUserMgtCode.frmSalaryAccountSetting
                        frm = New frmSalaryAccountSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDepartmentMaster
                        frm = New frmDepartmentMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSubDepartmentMaster
                        frm = New frmSubDepartmentMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmEmployeeTransfer
                        frm = New FrmEmployeeTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEPFund
                        frm = New frmEPF()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSkillMaster
                        frm = New frmSkillMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLanguageMaster
                        frm = New frmLanguageMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCourseMaster
                        frm = New frmCourseMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmShiftMaster
                        frm = New frmShiftMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDocumentMaster
                        frm = New frmDocumentMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPayPeriodMaster
                        frm = New frmPayPeriodMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCountryMaster
                        frm = New frmCountryMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReligionMaster
                        frm = New frmReligionMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmCurrencyMaster
                    '    frm = New frmCurrencyMaster()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCastCategoryMaster
                        frm = New frmCastCategoryMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGradeMaster
                        frm = New frmGradeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStateMaster
                        frm = New frmStateMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDevisionMaster
                        frm = New frmDevisionMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOccupationMaster
                        frm = New frmOccupationMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPFRulesMaster
                        frm = New frmPFRulesMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmESIRulesMaster
                        frm = New frmESIRulesMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOTMaster
                        frm = New frmOTMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAttendanceMaster
                        frm = New frmAttendanceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLICPolicyMaster
                        frm = New FrmLICPolicyMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeOTEntry
                        frm = New frmEmployeeOTEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBranchMaster
                        frm = New frmBranchMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBonusMaster
                        frm = New frmBonusMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOTSlab
                        frm = New frmOTSlab(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPTSlab
                        frm = New frmPTSlab(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRRSlab
                        frm = New frmHRRSlab(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmEmployeeDeductionMaster
                        frm = New frmEmployeeDeductionMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSavingsMaster
                        frm = New frmSavingsMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSectionAllowanceMaster
                        frm = New frmSectionAllowanceMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeSavingsMapping
                        frm = New frmEmployeeSavingsMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmConveyanceRateMaster
                        frm = New frmConveyanceRateMaster(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmODMaster
                        frm = New frmODMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ' KUNAL > TICKET : BM00000009879 > 30 - SEP - 2016 
                    Case clsUserMgtCode.frmPayrollDesignationMaster
                        frm = New frmDesignationMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode, clsUserMgtCode.frmPayrollDesignationMaster)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==============================Transaction===================================

                    Case clsUserMgtCode.frmOTSheet
                        frm = New frmOTSheet()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveMaster
                        frm = New frmLeaveMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveSetting
                        If clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
                            frm = New frmLeaveSetting()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        End If
                    Case clsUserMgtCode.frmGeneralHolidays
                        frm = New frmGeneralHolidays()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveOpeningBalance
                        frm = New frmLeaveOpeningBalance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveStartingDateSetting
                        frm = New frmLeaveStartingDateSetting()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveAllotment
                        frm = New frmLeaveAllotment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPayHeadDefinitions
                        frm = New frmPayHeadDefinitions()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryStructure
                        frm = New frmSalaryStructure()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMapPayHeadsToSalaStructure
                        frm = New frmMapPayHeadsToSalaStructure()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWeeklyHolidays
                        frm = New frmWeeklyHolidays()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveApplication
                        frm = New frmLeaveApplication()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMonthlyAttendance
                        frm = New frmMonthlyAttendance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveAdjustment
                        frm = New frmLeaveAdjustment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGenerateBonus
                        frm = New frmGenerateBonus()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailyAttendance
                        frm = New frmDailyAttendance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHourlyAttendance
                        frm = New frmHourlyAttendance()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAdjustmentVoucher
                        frm = New frmAdjustmentVoucher()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TOOLTYPE
                        frm = New FrmToolType()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReimbursementDetails
                        frm = New frmReimbursementDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAllowanceDetails
                        frm = New frmAllowanceDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeductionDetails
                        frm = New frmDeductionDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmApplyLoan
                        frm = New frmApplyLoan()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmArrear
                        frm = New frmArrear()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDA_Arrear
                        frm = New frmDA_Arrear()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveEncashment
                        frm = New frmLeaveEncashment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLICEmployeeTagging
                        frm = New frmLICEmployeeTagging()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeOTCalculation
                        frm = New frmEmployeeOTCalculation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmEmployee_Master
                        frm = New frmEmployee_Master()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmpSalary
                        frm = New frmEmployee_Salary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeStatus
                        frm = New frmEmployee_Status
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLoanAdjustment
                        frm = New frmLoanAdjustment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLoanGeneration
                        frm = New frmLoanGeneration
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmEmployeeIncrement
                    '    frm = New frmEmployeeIncrement
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmSalaryGeneration
                        frm = New frmSalaryGeneration
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptEmployeeAdvanceLedger
                        frm = New rptEmployeeAdvanceLedger
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptLoanStatement
                        frm = New rptLoanstatement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeGratuity
                        frm = New FrmEmployeeGratuity
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLTAClaim
                        frm = New frmLTAClaim
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMediclaimEntry
                        frm = New FrmMediclaimEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmFullAndFinalSettlement
                        Dim frm As New frmEmpFullAndFinalSettlement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeShiftChange
                        Dim frm As New frmEmployeeShiftChange
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmODSheet
                        Dim frm As New frmODSheet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmConveyanceClaim
                        Dim frm As New frmConveyanceClaim
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPayrollSetting
                        Dim frm As New frmPayrollSetting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmEmpIncrement
                        Dim frm As New FrmEmpIncrement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSentSalarySlip
                        Dim frm As New FrmSentSalarySlip
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAllotmentOfLeaves
                        Dim frm As New FrmAllotmentOfLeaves
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==============================Report===================================
                    Case clsUserMgtCode.frmSalaryGenerationRegister
                        frm = New frmSalaryGenerationRegister(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryGenerationRegisterArrear
                        frm = New frmSalaryGenerationRegister(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAllownceRegister
                        frm = New frmAllownceRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeductionRegister
                        frm = New frmDeductionRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReimbursementRegister
                        frm = New frmReimbursementRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAdjustmentRegister
                        frm = New frmAdjustmentRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAttendanceRegister
                        frm = New frmAttendanceRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeaveRegisterReport
                        frm = New frmLeaveRegisterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeRegister
                        frm = New frmEmployeeRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmActiveInactiveEmpRegister
                        frm = New frmActiveInactiveEmpRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptEmployeeDeductionMasterReport
                        frm = New rptEmployeeDeductionMasterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTentativeReport
                        frm = New RptTentativeSalary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptEmployeeOTReport
                        frm = New rptEmployeeOTReport
                    '   formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmPF_ESI_Reports
                        frm = New frmPF_ESI_Reports
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptEmployeeBday6
                        frm = New RptEmployeeBday6
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '-------------------------------------Monthly Report -----------------------------
                    Case clsUserMgtCode.frmPaySlip_Reports
                        frm = New frmPaySlip_Reports
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalarySheet_Reports
                        frm = New frmSalaryGenerationRegister(strProgramCode)
                        frm.Text = "Salary Sheet "
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryAbstractReport
                        frm = New frmSalaryAbstractReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAnnualCensusReport
                        frm = New frmAnnualCensusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAttendedDaysReport
                        frm = New frmAttendedDaysReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryVoucher_Reports
                        frm = New frmSalaryVoucher_Reports
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankStatement_Reports
                        frm = New frmBankStatement_Reports
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOT_Reports
                        frm = New frmOT_Reports
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVarianceReport
                        frm = New frmVarianceReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryComponentDetails
                        frm = New frmSalaryComponentDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAditionalEarning_DeductionReport
                        frm = New frmAditionalEarning_DeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryCertificate
                    'frm = New frmSalaryCertificate
                    '     formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.frmPF_Covering_Letter
                        frm = New frmPF_Covering_Letter
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryIncrementReport
                        frm = New frmSalaryIncrementReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankSummary_Report
                        frm = New frmBankSummary_Report
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeductionDetailsReport
                        frm = New frmDeductionDetailsReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVoucherPaymentsRegister
                        frm = New frmVoucherPaymentsRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmp_Id
                        frm = New frmEmp_Id
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLabelPrinting
                        frm = New frmLabelPrinting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmForm_T
                        frm = New frmForm_T
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    ''''''''''''''''''''''''''''''''''''''''''''Production Added by shipra------------------'''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''Master------------------'''''''''''''
                    Case clsUserMgtCode.ACCSETMFGSTD
                        frm = New FrmAccountSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ACCSETMFGDairy
                        frm = New FrmAccountSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ACCSETMFGPepsi
                        frm = New FrmAccountSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ITEMCATEGORY
                        frm = New FrmItemProductionCategory
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TOOL
                        frm = New FrmToolMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SETTSTD
                        frm = New FrmSettings(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SETTPep
                        frm = New FrmSettings(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.EXPENSE
                        frm = New FrmExpenseHead
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWorkCenterMaster
                        frm = New frmWorkCenterMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmResourceMaster
                        frm = New frmResourceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ALTER
                        frm = New FrmAlternateItem
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBillOfMaterialPepsi
                        frm = New frmBillOfMaterial
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
                    'If OpenProcessProductionBOm Then
                    '    frm = New frmBOM
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'Else
                    '    frm = New frmBiclsUserMgtCode.frmPPLogSheetMasterllOfMaterial
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    'End If
                    Case clsUserMgtCode.frmBillOfMaterialDairy
                        frm = New frmBOM
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProfitCenter
                        frm = New frmProfitCenter
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcessProductionLogSheet
                        'Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
                        'If OpenProcessProductionBOm Then
                        frm = New FrmProcessProductionLogSheet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'End If
                    Case clsUserMgtCode.frmPPLogSheetMaster
                        frm = New frmPPLogSheetMaster(clsUserMgtCode.frmPPLogSheetMaster)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPPLogSheetMaster_QC
                        frm = New frmPPLogSheetMaster(clsUserMgtCode.frmPPLogSheetMaster_QC)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorItemQCMapping
                        frm = New FrmVendorItemQCMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQualityCheckForSRN
                        frm = New FrmQualityCheckForSRN(clsUserMgtCode.frmQualityCheckForSRN, "Incoming")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog) 'frmQualityCheckApprovalForSRN
                    Case clsUserMgtCode.frmOutgoingQC
                        frm = New frmOutgoingQCEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQualityCheckApprovalForSRN
                        frm = New FrmQualityCheckApprovalForSRN(clsUserMgtCode.frmQualityCheckApprovalForSRN, "Incoming")
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPendingQCReport
                        frm = New rptPendingQCReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.QualitySummaryReport
                        frm = New QualitySummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptQCAnalysisReport
                        frm = New rptQCAnalysisReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptNIRQC
                        frm = New rptNIRQC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTruckWiseRejectionReport
                        frm = New frmTruckWiseRejectionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNIRQCvsWetQc
                        frm = New frmNIRQCvsWetQc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmOperationMaster
                        frm = New frmOperationMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBOMImport
                        frm = New frmBOMImport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.COSTMAINTAIN
                        frm = New FrmCostMaintainance
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PRO
                        frm = New FrmProductionLines
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProductionLines
                        frm = New FrmProductionLines
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBreakDownMaster
                        frm = New frmBreakDownMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBreakDownEntry
                        frm = New frmBreakDownEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProcessMaster1
                        frm = New FrmProcessMaster1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLineMaster
                        frm = New FrmLineMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSectionMaster
                        frm = New FrmSectionMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmStageMaster
                        frm = New frmStageMasters
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSectionStageMapping
                        frm = New FrmSectionStageMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''''''''''''''''''''''''''''''''''''''''''End of Master''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    ''''''''''''''''''''''''''''''''''''''''''''Transaction'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Case clsUserMgtCode.frmBillOfMaterialCosting
                        'Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
                        'If OpenProcessProductionBOm Then
                        '    frm = New frmBOM
                        '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                        'Else
                        frm = New frmBillOfMaterialCosting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmStanderdProductionEntry
                        frm = New frmStanderdProductionEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBagReceipt
                        frm = New FrmBagReceipt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'End If
                    'Case clsUserMgtCode.frmProductionPlanning
                    '    Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
                    '    If OpenProcessProductionBOm Then
                    '        frm = New FrmProcessProductionPlanning
                    '             formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    Else
                    '        frm = New frmProductionPlanning
                    '             formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    End If
                    Case clsUserMgtCode.frmProductionPlanningSTD
                        frm = New frmProductionPlanning(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMRPAutoMobile
                        frm = New frmMRPAutoMobile()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionPlanningPepsi
                        frm = New FrmProcessProductionPlanning
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionPlanningDairy
                        frm = New FrmProcessProductionPlanning
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcessProductionIssueEntry
                        frm = New FrmProcessProductionIssueEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSiloMilkTransfer
                        frm = New frmSiloMilkTransfer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmSiloMilkTransfer_JOBWORK
                    '    frm = New frmSiloMilkTransfer_JobWork
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSiloMilkTransferUploader
                        frm = New frmSiloMilkTransferUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMRPForProduction
                        frm = New frmMRPForProduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcessProductionStandardization
                        Dim ActivateProductionWithoutBatch As Decimal = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ActivateProductionWithoutBatch, clsFixedParameterCode.ActivateProductionWithoutBatch, Nothing))
                        If ActivateProductionWithoutBatch > 0 Then
                            frm = New frmRCDFStandardization
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmProcessProductionStandardization
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                        ActivateProductionWithoutBatch = Nothing
                    Case clsUserMgtCode.ProcessProductionStandardizationFinalQC
                        frm = New frmProcessProductionStandardizationFinalQC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcessProductionStageProcess
                        frm = New frmProcessProductionStageProcess
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '  Case clsUserMgtCode.frmProcessProductionStageProcess
                    '     frm = New RptDairyProductionWreckageReport
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmProductionEntry
                        Dim ActivateProductionWithoutBatch As Decimal = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ActivateProductionWithoutBatch, clsFixedParameterCode.ActivateProductionWithoutBatch, Nothing))
                        If ActivateProductionWithoutBatch > 0 Then
                            frm = New frmProductionEntryWithoutBatch
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmProductionEntry
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                        ActivateProductionWithoutBatch = Nothing
                    'Case clsUserMgtCode.frmProductionEntryWithoutBatch
                    '    frm = New frmProductionEntryWithoutBatch
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionEntryFinalQC
                        frm = New frmProductionEntryFinalQC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DariyProductionUploader
                        frm = New frmDairyProductionUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ProductionShiftMgmt
                        frm = New frmProductionShiftMgmt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmWreckageBooking
                        frm = New frmWreckage
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProcessProdReturn
                        frm = New frmProcessProductionReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssembDis
                        frm = New frmAssembDis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmMRP
                    '    frm = New frmMRP
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemoProductionPlanning
                        'frm=New frmDemoProductionPlanning
                        frm = New frmProductionPlanningDemo
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionRequisition
                        frm = New frmProductionRequisition
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRiceBOM
                        frm = New frmRiceBOM
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRiceMixingEntry
                        frm = New FrmRiceMixingEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRiceProcessingEntry
                        frm = New FrmRiceProcessingEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionItemSerialReplace
                        frm = New frmProductionItemSerialReplace
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionSerializedReport
                        frm = New frmProductionSerializedReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmProductionReceiptDemo
                        frm = New FrmProductionReceiptDemo
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmManufacturingOrder
                        frm = New frmManufacturingOrder
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmBatchOrder
                    '    Dim OpenProcessProductionBOm As Boolean = clsDBFuncationality.getSingleValue("select IsBOMFromProcessProduction from TSPL_INV_PARAMETERS")
                    '    If OpenProcessProductionBOm Then
                    '        frm = New FrmProcessBatchOrder
                    '             formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    Else
                    '        frm = New frmBatchOrder
                    '             formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '    End If
                    '==================preeti gupta==============
                    '=============Production Report==============
                    Case clsUserMgtCode.rptproductionEntryReport
                        frm = New RptproductionEntryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptItemConsumptionReport
                        frm = New RptItemConsumptionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductionIssueStatusReport
                        frm = New RptProductionIssueStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptWTPReport
                        frm = New RptWIPReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDairyProductionWreckageReport
                        frm = New RptWreckageReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProductionAndSaleReport
                        frm = New FrmProductionAndSaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptJobWorkProduction
                    '    frm = New rptJobWorkProduction
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptAvailableQtyForProduction
                        frm = New rptAvailableQtyForProduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBatchStatus
                        frm = New RptBatchStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptStandardQCReport
                        frm = New RptStandardQCReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptStandardActualConsumption
                        frm = New rptStandardActualConsumption
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptProductionQCReport
                        frm = New RptProductionQCReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptProductionWiseStockReco
                        frm = New rptProductionWiseStockReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptIssueWIPConsumptionReport
                        frm = New rptIssueWIPConsumptionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProcessGainLossReport
                        frm = New rptProcessGainLossReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSectionWiseStockReport
                        frm = New RptSectionWiseStockReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBatchOrderSTD
                        frm = New frmBatchOrder(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBatchOrderPepsi
                        frm = New frmBatchOrder(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBatchOrderDairy
                        frm = New FrmProcessBatchOrder
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionStoreRequest
                        frm = New frmProductionStoreRequest
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmStoreIssueSTD
                        frm = New frmStoreIssue(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmStoreIssuePepsi
                        frm = New frmStoreIssue(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmProductionReturnSTD
                        frm = New frmProductionReturn(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionReturnPep
                        frm = New frmProductionReturn(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionReceiptSTD
                        frm = New frmProductionReceipt(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLabourWorkingSheet
                        frm = New FrmLabourWorkingSheet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''''''''''''''''''''''''''''''''''''''''''End of Transaction''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''Reports''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Case clsUserMgtCode.Resource
                        frm = New RptListOf_Resource
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LTOOL
                        frm = New RptListOfToolType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LACCt
                        frm = New RptListOfAccountSet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LALT
                        frm = New RptListOfAlternateItem
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LWC
                        frm = New RptListOfWorkCenter
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LOIC
                        frm = New rptListOfItemCost
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LOPER
                        frm = New RptListOfOperations
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PRODREPORT
                        frm = New FrmListOfProductionLines
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LToolT
                        frm = New RptListOfTools
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmListOfBOM
                        frm = New frmListOfBOM
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIssueReturnItemWiseReportSTD
                        frm = New frmIssueReturnItemWiseReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIssueReturnItemWiseReportPepsi
                        frm = New frmIssueReturnItemWiseReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDatewiseProduction
                        frm = New frmDatewiseProduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionPlanReportSTD
                        frm = New frmProductionPlanReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionPlanReportPepsi
                        frm = New frmProductionPlanReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLineProductivity
                        frm = New frmLineProductivity
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBatchOrderReportSTD
                        frm = New frmBatchOrderReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBatchOrderReportPepsi
                        frm = New frmBatchOrderReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmGraphicalBatchOrder
                    '    frm = New frmGraphicalBatchOrder
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmGraphicalCategorywiseProduction
                    '    frm = New frmGraphicalCategorywiseProduction
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmListofRequisition
                        frm = New frmListofRequisition
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOperaterEfficiencyReport
                        frm = New FrmOperatorEfficiencyReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    ''BI Forms
                    'Case clsUserMgtCode.BICreateReport
                    '    frm = New frmCreateBIReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BICreateFilter
                    '    frm = New frmCreateBIFilter
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BICreateDashBoard
                    '    frm = New FrmCreateDashBoard
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BIMonthWiseSale
                    '    frm = New FrmBIMonthWiseSale
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BITopCustomer
                    '    frm = New frmBITopCustomer
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BITopItemCategory
                    '    frm = New frmBITopItemCategory
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionVarianceSTD
                        frm = New frmProductionVariance(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionVariancePepsi
                        frm = New frmProductionVariance(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BIMonthWisePurchase
                    '    frm = New FrmBIMonthWisePurchase
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BITopVendor
                    '    frm = New frmBITopVendor
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BITopItemCategoryPurchase
                    '    frm = New frmBITopItemCategoryPurchase
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BIMonthWiseAssset
                    '    frm = New frmBIMonthWiseAsset
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BITopExpence
                    '    frm = New frmBITopExpence
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.BIDashBoadr
                    '    frm = New frmBIDashBoard
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''''''''''''''''''''''''''''''''''''''''''End of Reports''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    ''''''''''''''''''''''''''''''''''''''''''''Project Management''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''Setup (Project Management)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Case clsUserMgtCode.frmPJCSettings
                        frm = New frmPJCSettings
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCostTypes
                        frm = New frmCostType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPJCAccountSets
                        frm = New frmPJCAccountSetting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmJobMaster
                        frm = New frmJobMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTaskMaster
                        frm = New frmTaskMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPJCEmployeeMaster
                        frm = New frmPJCEmployeeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmUserApproval
                        frm = New FrmUserApproval
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBudgetMaintenance
                        frm = New FrmBudgetMaintenance
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ProjectMaster
                        frm = New FrmProjectMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmExpenseType
                        frm = New FrmExpenseType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' transaction
                    Case clsUserMgtCode.frmTimeSheet
                        frm = New frmTimesheet
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUserLog
                        frm = New FrmUserLog
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProjectStatus
                        frm = New FrmProjectStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssemblies
                        frm = New frmAssemblies
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPJCExpense
                        frm = New FrmPJCExpense
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '-Reports-------
                    Case clsUserMgtCode.frmProjectListReport
                        frm = New FrmProjectListReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProjectDetails
                        frm = New FrmProjectDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProjectProfitReport
                        frm = New FrmProjectProfitReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSubRouteMaster
                        frm = New frmSubRouteMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProfitAndLossPerforma
                        frm = New frmProfitAndLossPerforma
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProfitAndLoss
                        frm = New frmRptProfitAndLoss
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''''''''''''''''''''''''''''''''''''''''''''end Project Management''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Service''''''''''''''''''''''''''''''''''''
                    Case clsUserMgtCode.frmComplaintGroupMaster
                        frm = New frmComplaintGroupMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPrimaryReasonMaster
                        frm = New FrmPrimaryReasonMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmComplaintMaster
                        frm = New frmComplaintMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPendingReasonMaster
                        frm = New frmPendingReasonMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmItemChargeCategoryMaster
                        frm = New frmItemChargeCategoryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmItemChargeFranchiseMappingMaster
                        frm = New FrmItemChargeFranchiseMappingMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetServiceMaster
                        frm = New FrmAssetServiceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------ TRANSACTION ---------------------------------'
                    Case clsUserMgtCode.frmAssetAgreement
                        frm = New FrmAssetAgreement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetInstallPullOut
                        frm = New frmAssetInstallPullOut
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmComplaintDetailEntry
                        frm = New FrmComplaintDetailEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQuickComplaintDetailEntry
                        frm = New FrmQuickComplaintDetailEntry(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetDistatch
                        frm = New frmAssetDispatch
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCartMaintenanceEntry
                        frm = New FrmCartMaintenanceEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingComplaintDetail
                        frm = New FrmPendingComplaintDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '--------------------REPORT-------------------------------------'
                    Case clsUserMgtCode.frmFranchiseChargesReport
                        frm = New FrmFranchiseChargesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomersListReport
                        frm = New FrmCustomersListReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPullOutRedeployReport
                        frm = New frmPullOutRedeployReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorBillDetails
                        frm = New frmVendorBillDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDaywisePendingComplaint
                        frm = New FrmDaywisePendingComplaint
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmClaimReport
                        frm = New FrmClaimReportNew
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAssetDetailReport
                        frm = New FrmAssetDetailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmspareStockReport2
                    '    frm=New FrmspareStockReport2
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    '--------------------MILK PROCUREMENT MASTER-------------------------------------'
                    Case clsUserMgtCode.frmMilkCollectionLevels
                        frm = New frmMilkCollectionLevelsMain
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkAdvanceMaster
                        frm = New frmJWPriceCodeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkVehicleTypeMaster
                        frm = New VehicleTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkComponentMaster
                        frm = New MilkComponentMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkRateTypeMaster
                        frm = New MilkRateType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkShiftMaster
                        frm = New ShiftMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSeasonMaster
                        frm = New SeasonMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkRouteMaster
                        frm = New FrmMilkRouteMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkVehicleMaster
                        frm = New frmVehicleMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VehicleMasterForDairySale
                        frm = New frmVehicleMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDistributorCommission
                        frm = New frmDistributorCommission
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBoothCommissionMaster
                        frm = New FrmBoothCommissionMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmItemCapacityLimt
                        frm = New FrmItemCapacityLimit
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBoothRouteMapping
                        frm = New FrmBoothRouteMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDayBook
                        frm = New frmDayBook
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAccountSalesReport
                        frm = New frmAccountSalesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VehicleMasterForProductSale
                        frm = New frmVehicleMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkCollectionArea
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMilkTransportRateMaster
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMilkComponentRateList
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmVillageMaster
                        frm = New FrmVillageMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVSPMaster
                        frm = New frmVSPMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmVSP_VLCMaster
                        frm = New frmVSP_VLCMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmPrimaryTransporterMaster
                        frm = New FrmPrimaryTransporterMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTankerTransporterMaster
                        frm = New frmTankerTransporterMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaster
                        frm = New FrmMCCMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTankerMaster
                        frm = New FrmTankerMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVLCMaster
                        frm = New FrmVLCMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVLCMasterTarget
                        frm = New FrmVlcTargetMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeliveryTermsMaster
                        frm = New frmDeliveryTermsMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPriceChartUploader
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, Nothing)) > 0 AndAlso
                            clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.PickPriceFromFATAndSNF, Nothing)) <= 0 Then
                            frm = New frmPriceChartUploaderCLR
                        Else
                            frm = New FrmPriceChartUploader
                        End If
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MPMilkPrice
                        frm = New frmMPMilkPrice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnpostBmcDcs
                        frm = New frmUnpostBmcDcs
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDCSTransportationCharges
                        frm = New frmDCSTransportationCharges
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmManualIoTFarmerCollection
                        frm = New frmManualIoTFarmerCollection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.GazeReading
                        frm = New frmGazeReading
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmHeadLoadMaster
                        frm = New frmHeadLoadMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDeductionTypeMaster
                        frm = New frmDeductionTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmPriceChartMaster
                        frm = New FrmPriceChartMaster(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkPricePlanning
                        Dim intPricePlan As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.OpenPriceChartPlanningScreenOnTotalSolid, clsFixedParameterCode.OpenPriceChartPlanningScreenOnTotalSolid, Nothing))
                        If intPricePlan = 1 Then
                            frm = New frmPriceChartPlanMasterGHO
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 2 Then
                            frm = New frmPriceChartPlanMasterGK
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 3 Then
                            frm = New frmPriceChartPlanMasterTSDDCF
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 4 Then
                            frm = New frmPriceChartPlanMasterBhole
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 5 Then
                            frm = New frmPriceChartPlanMasterUCDF
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 6 Then
                            frm = New frmPriceChartPlanMasterRCDF
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 7 Then
                            frm = New frmPriceChartPlanMasterJPR
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        ElseIf intPricePlan = 8 Then
                            frm = New frmPriceChartPlanMasterAJM
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        Else
                            frm = New frmPriceChartPlanMaster
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmPriceChartMaster_Bulk
                        frm = New FrmPriceChartMaster(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFormIssueReceiptEntry
                        frm = New FrmFormIssueReceiptEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmReceiptInvoiceMapping
                    '    frm = New FrmReceiptInvoiceMapping
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCostCentreGroupStores
                        frm = New FrmCostCentreGroupStores
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVLCUploader
                        frm = New FrmVLCUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMPMaster
                        frm = New FrmMPMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VerifyMPIFSC
                        frm = New frmVerifyMPIFSC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmParameterMaster
                        frm = New FrmParameterMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmParameterRangeMaster
                        frm = New FrmParameterRangeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVLCRouteShiftMaster
                        frm = New FrmVLCRouteShiftMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPrimaryTransporterVehicalMaster
                        frm = New FrmPrimaryTransporterVehicalMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOpenMCCShift
                        '==Add New Variable in Open Mcc Shift .it is Running Two Screens on Basis on It By: Rohit Gupta=========s
                        frm = New FrmOpenMCCShift(clsUserMgtCode.frmOpenMCCShift)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOpenMCCShiftManual
                        frm = New FrmOpenMCCShift(clsUserMgtCode.frmOpenMCCShiftManual)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmParameterRangeMasterForQC
                        frm = New frmParameterRangeMasterForQC(clsUserMgtCode.frmParameterRangeMasterForQC)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQualityModuleParameterRangeMaster
                        frm = New frmParameterRangeMasterForQC(clsUserMgtCode.frmQualityModuleParameterRangeMaster)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPortSettings
                        frm = New FrmPortSettings
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPriceChartBulkProc
                        frm = New frmPriceChartBulkProc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ParameterValueMaster
                        frm = New FrmParameterValueMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LostDefectSealNo
                        frm = New FrmLostDefectSealNo
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.LocationDistanceMapping
                        frm = New FrmLocationDistanceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MccMilkTransferPrice
                        frm = New FrmMccMilkTransferPrice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorPriceChartMapping
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowBulkPriceChartMultiplepriceToMultipleVendor, clsFixedParameterCode.AllowBulkPriceChartMultiplepriceToMultipleVendor, Nothing)) > 0 Then
                            frm = New frmVendorPriceChartMappingUDL()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmVendorPriceChartMapping()
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmdeductionGroup
                        frm = New FrmDeductionGroup1()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeductionMaster
                        frm = New FrmDeductionMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DeductionMapping
                        frm = New frmDeductionMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CanMaster
                        frm = New frmCanMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DockMaster
                        frm = New frmDockMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmTankerDispatchPriceMaster
                        frm = New FrmTankerDispatchPrice_Master()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGroupOfDeduction
                        frm = New FrmGroupOfDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkRejectType
                        frm = New frmMilkRejectType()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DCSAdditionDeduction
                        frm = New frmDCSAdditionDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmChildRouteFreight
                        frm = New frmchildRouteFreight()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    '--------------------MILK PROCUREMENT TRANSACTION-------------------------------------'
                    Case clsUserMgtCode.frmMilkCollectionCenters
                        frm = New FrmMilkCollectionCenters
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmMilkSuppliers
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMCCRouteMapping
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMCCSuperwiserMapping
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMCCSupplierMapping
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMilkCollection
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMilkQualityCheck
                        clsCommon.MyMessageBoxShow("Under Development")

                    Case clsUserMgtCode.frmMilkRateProcessingScheme
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmVehicleMovement
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.frmMilkBillGeneration
                        clsCommon.MyMessageBoxShow("Under Development")
                    Case clsUserMgtCode.MilkGateEntryIn
                        frm = New frmMilkGateEntryIn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AMCUMaster
                        frm = New AMCUMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.WEIGHINGMaster
                        frm = New WEIGHINGMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        'Case clsUserMgtCode.MilkRetesting
                        ' frm = New frmMilkRetesting
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntryTransaction
                        frm = New frmGateEntryTransaction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkGateEntryWeightment
                        frm = New frmMilkGateEntryWeighment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkGateEntryOut
                        frm = New frmMilkGateEntryOut
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkReceipt
                        frm = New frmMilkReceiptMCC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCancelAfterPosting
                        frm = New FrmCancelAfterPosting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMCCMilkTransPortorInvoice
                        frm = New FrmRecurringPayableInvoice 'frmMccMilkTransportorInvoice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateEntry
                        frm = New FrmGateEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCDispatch
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, Nothing)) = 1 Then
                            frm = New frmMccDispatchChamber
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New FrmMccDispatch
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmAcknowledgementEntry
                        frm = New frmAcknowledgementEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkProcurementUploader
                        frm = New frmMilkProcurementUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkShiftUploader
                        Dim Type As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ProcurmentShiftUploaderNo, clsFixedParameterCode.ProcurmentShiftUploaderNo, Nothing))
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        If Type = 2 Then
                            objCommonVar.IsAutoTabOrdering = False
                            frm = New frmMilkShiftUploaderRaj
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                            objCommonVar.IsAutoTabOrdering = x
                        ElseIf Type = 3 Then
                            objCommonVar.IsAutoTabOrdering = False
                            frm = New frmMilkShiftUploaderUCDF
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                            objCommonVar.IsAutoTabOrdering = x
                        Else
                            frm = New frmMilkShiftUploader
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.MilkProcurementFarmerCorrection
                        frm = New frmCorrectionFarmer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkProcurementCorrection
                        frm = New frmCorrection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkRetesting
                        frm = New frmCorrection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.PrimaryTransportProvisionCorrection
                        frm = New frmPrimaryTransporterProvisionCorrection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOutputEntry
                        frm = New frmOutputEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmMCCTankerDispatchReturn
                    '    frm = New FrmMccTankerDispatchReturn
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmMCCTankerGateOut
                    '    frm = New FrmMCCTankerGateOut
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmMCCTankerGateOutSecurity
                    '    If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MCCBulkProcumentSecurityGateOut, clsFixedParameterCode.MCCBulkProcumentSecurityGateOut, Nothing)) = 1 Then
                    '        frm = New FrmMCCTankerGateOutSecurity
                    '        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '    Else
                    '        clsCommon.MyMessageBoxShow("This feature is not for you")
                    '    End If

                    Case clsUserMgtCode.frmWeighment
                        frm = New FrmWeighment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.SecondarySettingForQC
                    '    frm = New FrmSecondarySettingForQC
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TDSPAYMENT
                        frm = New FrmTDSPayment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQualityCheck
                        frm = New FrmQualityCheck
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkSample
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MilkProcuremntPickCLRInsteadOfSNF, clsFixedParameterCode.MilkProcuremntPickCLRInsteadOfSNF, Nothing)) > 0 Then
                            frm = New frmMilkSampleMCCOddEvenCLR
                        ElseIf clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.Open4AnalyzerForm, clsFixedParameterCode.Open4AnalyzerForm, Nothing)) = 1 Then
                            frm = New frmMilkSampleMCC124
                        ElseIf clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.OpenODDEvenForm, clsFixedParameterCode.OpenODDEvenForm, Nothing)) = 1 Then
                            frm = New frmMilkSampleMCCOddEven
                        Else
                            frm = New frmMilkSampleMCCOddEven ''frmMilkSampleMCC By Balwinder on 12/04/2022 Auto/Manual Work fine on oddeven form
                        End If
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCSMSSettiing
                        frm = New FrmMccSMSSetting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDCSSaleEntry
                        frm = New frmDCSSaleEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMobileDCSDemand
                        frm = New frmMobileDCSDemand
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCancelDCSSale
                        frm = New frmCancelDCSSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterial
                        frm = New frmMCCMaterialSale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterialSaleReturn
                        frm = New frmMccMaterialSaleReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterialSalePriceChart
                        frm = New FrmMCCMaterialSalePriceChart
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkSRN
                        frm = New frmMilkSRNMCC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MCCSetting
                        frm = New frmMCCProcurementSetting(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVSPIncentiveTagging
                        frm = New FrmVSPIncentiveTagging
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmMilkPurchaseInvoice
                        frm = New frmMilkPurchaseInvoiceMCC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkShiftEndMCC
                        frm = New frmMilkShiftClosingMCC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.frmMilkTransferIn
                        frm = New FrmMilkTransferIn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkTransferInReturn
                        frm = New frmMilkTransferInReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQCSeparation
                        frm = New FrmQCSeparation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTankerProvision
                        frm = New frmTankerProvision
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkCollectionGenerate
                        frm = New frmMilkCollectionGenerate
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkCollectionMCCMultipleDays
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        frm = New frmMilkCollectionMCCMultipleDays
                        objCommonVar.IsAutoTabOrdering = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.MilkCollectionDCSMultipleDays
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        frm = New frmMilkCollectionDCSMultipleDays
                        objCommonVar.IsAutoTabOrdering = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.MilkCollectionDCSMultipleDaysMerge
                        settDCS = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyMergeForDCSMultipleDays, clsFixedParameterCode.ApplyMergeForDCSMultipleDays, Nothing)) = 1)
                        If settDCS Then
                            settDCS = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowDCSDetMerge, clsFixedParameterCode.ShowDCSDetMerge, Nothing))
                            If settDCS = True Then
                                Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                                frm = New frmDCSMilkCollectionMergeSetting
                                objCommonVar.IsAutoTabOrdering = False
                                formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                                objCommonVar.IsAutoTabOrdering = x
                            Else
                                Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                                frm = New frmMilkCollectionDCSMultipleDaysMerge
                                objCommonVar.IsAutoTabOrdering = False
                                formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                                objCommonVar.IsAutoTabOrdering = x
                            End If
                        Else
                            clsCommon.MyMessageBoxShow(Me, "This feature is not for you")
                        End If

                    'Case clsUserMgtCode.MilkCollectionDCSMultipleDaysMerge
                    '    Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                    '    frm = New frmMilkCollectionDCSMultipleDaysMerge
                    '    objCommonVar.IsAutoTabOrdering = False
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '    objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.MilkCollectionMCC
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        frm = New frmMilkCollectionMCC
                        objCommonVar.IsAutoTabOrdering = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.MilkCollectionMCCGateEntry
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        frm = New frmMilkCollectionMCC
                        objCommonVar.IsAutoTabOrdering = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.MilkCollectionMCCSample
                        frm = New frmMilkCollectionMCCQC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.MilkCollectionDCS
                        Dim x As Boolean = objCommonVar.IsAutoTabOrdering
                        frm = New frmMilkCollectionDCS
                        objCommonVar.IsAutoTabOrdering = False
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        objCommonVar.IsAutoTabOrdering = x
                    Case clsUserMgtCode.frmBulkMilkSRN
                        frm = New FrmBulkMilkSRN
                        frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDBTNEFTUnionReport
                        frm = New frmDBTNEFTUnionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DBTPDAccountReport
                        If objCommonVar.RCDFCFP Then
                            frm = New frmDBTPDAccountReport
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        End If

                    Case clsUserMgtCode.frmDCSSavingLedger
                        frm = New frmDCSSavingLedger
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnionWiseDockEntryReport
                        frm = New frmUnionWiseDockEntryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmInActiveDCSReport
                        frm = New FrmInActiveDCSReport11
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDBTStatusAndLastDPTStatus
                        frm = New frmDBTStatusAndLastDPTStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDBTApprovalStatus
                        frm = New frmDBTApprovalStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDBTNEFTPaymentDetailReport
                        frm = New rptDBTNEFTPaymentDetailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBulkMilkSRNReturn
                        frm = New FrmBulkMilkSRNReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnloading
                        frm = New FrmUnloading
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkMilkPurchaseInvoice
                        frm = New FrmMilkPurchaseInvoice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BulkProcurementUploader
                        frm = New BulkProcurementUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple
                        frm = New frmBulkMilkPurchaseInvoiceMultiple
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMilkPurchaseReturn
                        frm = New FrmMilkPurchaseReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCleaning
                        frm = New FrmCleaning
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGateOut
                        frm = New FrmGateOut
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBulkPurchaseUploader
                        frm = New frmBulkPurchaseUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBankUpdateUploader
                        frm = New FrmBankUpdateUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVlcdataUploadar
                        frm = New FrmVlcDataUploadar
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmVLCDataUploaderManual
                        frm = New FrmVLCDataUploaderManual
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MPIncentiveEntry
                        frm = New frmMPIncentiveEntrty
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DCSMPIncentiveReco
                        frm = New frmMPDCSIncentiveReco
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DBTMonthlyFarmerMilk
                        frm = New frmDBTMonthlyFarmerMilk
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DBTCappingCheck
                        If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DBTMilkQtyCapping, clsFixedParameterCode.DBTMilkQtyCapping, Nothing)) > 0 Then
                            frm = New frmDBTCaping
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        End If
                    Case clsUserMgtCode.DBTCappingIncrease
                        If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DBTMilkQtyCapping, clsFixedParameterCode.DBTMilkQtyCapping, Nothing)) > 0 Then
                            frm = New frmDBTCaping
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            clsCommon.MyMessageBoxShow("This feature is not for you")
                        End If

                    Case clsUserMgtCode.DBTNEFTUploader
                        frm = New frmDBTNEFTUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCreateBMCDCSbyMobile
                        frm = New FrmCreateBMCDCSbyMobile
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DBTNEFTReject
                        frm = New frmDBTNEFTReject
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DBTPayment
                        frm = New FrmDBTPayment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SendDBTToJanaadhar
                        frm = New FrmSendDBTToJanaadhar
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.MPIncentiveEntryReport
                        frm = New frmMPIncetiveEntryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FailedBDF
                    '    frm = New frmFailBDF
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VLCProgressReport
                        frm = New frmVLCProgressReportReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FATSNFDiffReport
                        frm = New frmFATSNFDiffReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMCCShiftReportRouteWise
                        frm = New FrmMCCShiftReportRouteWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmDispatchTransfer
                    '    frm = New FrmDispatchTransfer
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '--------------------MILK PROCUREMENT Report-------------------------------------'
                    Case clsUserMgtCode.rptCollectionLevelChart
                        frm = New rptCollectionLevelChart
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCollectionCenterChart
                        frm = New rptCollectionCenterChart
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptvillageslip
                        frm = New RptVillageSlip
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkBillMCC
                        frm = New RptMilkBillMCC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkBillRouteWise
                        frm = New RptMilkBillRouteWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMCCMilkBillSummary
                        frm = New RptMccMilkBillSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptTankerSummaryReport
                        frm = New RptTankerSummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptShiftCodeWise
                    '    frm = New RptShiftReportCodeWise
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.rptShifReportZeroAmtSample
                        frm = New RptShiftReportZeroAmtSample
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkPaymentRegister
                        frm = New RptPaymentRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptLowProcurement
                        frm = New RptLowProcurement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMccSaleRegister
                        'frm = New RptMCCsaleRegister
                        'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '' changed by panch raj on 08-05-18 against ticket No: KDI/04/05/18-000295
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptBulkMilkRegister
                        frm = New RptBulkMilkRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptTotalMilkProcurement
                        frm = New rpttotalMilkProcurement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptMDConversion
                    '    frm = New RptMDConversionAtUDL
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMilkRouteVehicleReport
                        frm = New RptMilkRouteVehicleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCDA
                        frm = New RptCDA
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptDailyProgressReport
                    '    frm = New RptDailyProgressReport
                    '         formShow(frm,strProgramCode, strProgramName, isOpenInMDI,strDocNo)
                    Case clsUserMgtCode.rptMonthlyVLCProcurement
                        frm = New RptMonthlyVLCProcurement1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptSecondaryQuality
                    '    frm = New RptSecondaryQualityReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDailyDifferentReport
                        frm = New RptDailyDifferentRow_vb
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVSPAssetIssue
                        frm = New frmVSPAssetIssue
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmVSPItemIssue
                    '    frm = New frmVSPItemIssue
                    '    frm.AllowModifcationByApprovalUser = IsAllowModificationByApprovalUser
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMillPurchaseBill
                        frm = New RptMilkPurchaseBill
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptGainSheetPeriod
                    '    frm = New RptGainSheetPeriod
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptWeighment
                        frm = New RptWeightment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptTankerVariation
                    '    frm = New RptTankerVariationReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptDailyGainDay
                        frm = New RptDailyGainDay
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPendingMilkSRN
                        frm = New RptPendingMilkSRN
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPendingProvisionReport
                        frm = New FrmPendingProvisionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptPTReport
                    '    frm = New FrmPTReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MCCMilkRegister
                        frm = New FrmMCCMilkRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptLiabilityReport
                        frm = New rptLiabilityReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkBillProcurementSummary
                        frm = New rptMilkBillProcurementSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmVendorBankAdvice
                        frm = New frmVendorBankAdvice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.rptMilkAnalysis
                        frm = New rptMilkAnalysis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSocietyLedgerReport
                        frm = New rptSocietyLedgerReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkCostReport
                        frm = New rptMilkCostReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDcscollectionHistory
                        frm = New FrmDcscollectionHistory
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHeadLoadHistory
                        frm = New FrmHeadLoadHistory
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmOneDayStop
                        frm = New FrmOneDayStop
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmLocalMilkSaleRegister
                        frm = New FrmLocalMilkSaleRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDCSFinancial
                        frm = New rptDCSFinancial
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPaymentProcessReportBMCSocietyWise
                        frm = New rptPaymentProcessReportBMCSocietyWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMCCBillGenrationStatus
                        frm = New rptMCCBillGenrationStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkProcurementVisualReport
                        frm = New VisualTopProcurement ''VisualReportProcurement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptPaymentProcessRouteReport
                        frm = New rptPaymentProcessRouteReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.BmcLabReport
                        frm = New BmcLabReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.CanSummaryReport
                        frm = New Can_Summary_Report
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.YearlyDBTSummaryReport
                        frm = New YearlyDBTSummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmDayWiseMilkCollectionVsDispatch
                        frm = New FrmDayWiseMilkCollectionVsDispatch
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptPaymentProcessReport
                        frm = New rptPaymentProcessReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptVSPMilkNotsold
                        frm = New rptVSPMilkNotSold
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMCCDataEntrySummaryReport
                        frm = New rptMCCDataEntrySummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptTankerAllocationReport
                        frm = New rptTankerAllocationReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMilkCollectionShiftWiseReport
                        frm = New rptMilkCollectionShiftWiseReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMccDeductionReport
                        frm = New rptMccDeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptShedIOAbstract
                    '    frm = New rptShedIOAbstract
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptWholeMilkAccount
                    '    frm = New rptWholeMilkAccount
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptPTPBill
                    '    frm = New rptPTPBill
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptVLCTransportExpense
                    '    frm = New rptVLCTransportExpense
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptTankerDispatchvsAckn
                        frm = New RptTankerDispatchvsAckn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptDairyBookingDistributorReport
                        frm = New RptDairyBookingDistributorReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBoothDemandReport
                        frm = New FrmBoothDemandReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmBoothCountReport
                        frm = New FrmBoothCountReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDairyTruckSheetReport
                        frm = New rptDairyTruckSheetReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDayWiseBoothDemand
                        frm = New rptDayWiseBoothDemand()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSalesComparisonReport
                        frm = New rptSalesComparisonReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSaleRegisterUOMReport
                        frm = New rptSaleRegisterUOMReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCommissionTPTReport
                        frm = New RptcommissionTPTReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGatepassDetailReport
                        frm = New frmGatepassDetailReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSaleRegisterDetail
                        frm = New rptSaleRegisterDetail()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBookingWiseRegister
                        frm = New rptBookingWiseRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDemandNill
                        frm = New RptBoothNilDemandl()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMonthlyBillSummaryReport
                        frm = New rptMonthlyBillSummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemandHistory
                        frm = New frmDemandHistory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptItemWiseBillReport
                        frm = New rptItemWiseBillReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDistributorRouteMarginStatusReport
                        frm = New rptDistributorRouteMarginStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPartWiseItemReport
                        frm = New rptPartWiseItemReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmEwaybill
                        frm = New FrmEwaybill()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPrintMultipleGatepass
                        frm = New FrmPrintMultipleGatepass()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBoothTruckSheet
                        frm = New rptBoothTruckSheet()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDemandReport
                        frm = New rptDemandReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptZoneWiseMSVisual
                        frm = New rptZoneWiseMilkSaleVisual()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VisualReportSale
                        frm = New VisualReportSaleNew()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDairySaleRegisterReport
                        frm = New rptDairySaleRegisterReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.InvoicePriceWiseReport
                        frm = New InvoicePriceWiseReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptInvoiceDetailReport
                        frm = New rptInvoiceDetailReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDailyReceiptReport
                        frm = New rptDailyreceiptReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTransporterWiseReport
                        frm = New rptTransporterWiseReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDailyLeakageReplacementReport
                        frm = New rptDailyLeakageReplacementReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptUnpostedDocumentReport
                        frm = New rptUnpostedDocumentReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmZoneWiseSKUReport
                        frm = New FrmZoneWiseSKUReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDairySaleSchemeReport
                        frm = New FrmDairySaleSchemeReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBankDetailsWithVendorMargin
                        frm = New rptBankDetailsWithVendorMargin()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptAbsentBooth
                        frm = New rptAbsentBooth()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCustomerWiseSalesReport
                        frm = New rptCustomerWiseSalesReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBookingReport
                        frm = New rptBookingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBookingQtyAmtReport
                        frm = New rptBookingQtyAmtReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSalesLedgerReport
                        frm = New rptSalesLedgerReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSalesStock
                        frm = New rptSalesStock()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptProvisionChart
                        frm = New rptProvisionChart()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSalesVehicleReport
                        frm = New rptSalesVehicleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptRouteWiseSaleRegister
                        frm = New RptRouteWiseSaleRegister()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.SaleEinvoiceReport
                        frm = New SaleEinvoiceReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.rptSaleInvoiceStatusReport
                        frm = New rptSaleInvoiceStatusReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptProductCreditSaleReport
                        frm = New rptProductCreditSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptVSPIncentiveRegister
                        frm = New rptVSPIncentiveRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.MccSummaryReport
                        frm = New FrmMCCSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkStockLedgerSummary
                        frm = New RptMilkStockLegderSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptMilkWeigmentRegister
                        frm = New RptMilkWeigmentRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMemberPaymentSlip
                        frm = New Rptmemberpaymentslip3
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMPWiseMilkCollection
                        frm = New RptMPWiseMilkCollection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMPWiseMilkCollectionATPoolingPoint
                        frm = New RptMPWiseMilkCollectionAtPoolingPoint3
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptVillageDiffReport
                        frm = New RptVillageDiffReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptVillageDiffReportParas
                        frm = New RptVillageDifferenceREportParas
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMCCVLCVariationReportPourersNo
                        frm = New rptMCCVLCVariationReportPourersNo
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)



                    Case clsUserMgtCode.rptPrimaryTransporter
                        frm = New RptPrimaryTransporter
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptMCCMilkRegisterDripSaver
                    '    frm = New RptMCCMilkRegisterDripSaver
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMCCVLCVarationReport
                        frm = New FrmMCCVLCVarationReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVSPOrVLCVarationRpt
                        frm = New RptVSPOrVLCVarationReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptCollectionAnalysis
                        frm = New RptCollectionAnalysis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMPIDReport
                        frm = New RptMPIDReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmProvisionEntry
                        frm = New FrmProvisionEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.IncentiveEntry
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MilkIncetiveByMilkSRN, clsFixedParameterCode.MilkIncetiveByMilkSRN, Nothing)) > 0 Then
                            frm = New frmIncetiveEntryBySRN
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        Else
                            frm = New frmIncetiveEntry
                            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        End If
                    Case clsUserMgtCode.frmTransferToSaving
                        frm = New frmTransferToSaving()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMilkProcurementSocietySummary
                        frm = New FrmMilkProcurementSocietySummary()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMakeSavingPayment
                        frm = New frmMakeSavingPayment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPaymentProcess
                        frm = New FrmPaymentProcess
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BMC_Transporter_Bill
                        frm = New BMC_Transporter_Bill
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTransporterPaymentProcess
                        frm = New frmTransporterPaymentProcess
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCattleFeedSalePurchaseUploader
                        frm = New FrmCattleFeedSalePurchaseUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBankAdvise
                        frm = New frmBankAdvise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTDSReport
                        frm = New frmTDSReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSendBillToDCS
                        frm = New frmSendBillToDCS
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSendSMSToDCS
                        frm = New frmSendSMSToDCS
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReviseMilkBill
                        frm = New frmReviseMilkBill
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTankerStatusReport
                        frm = New rptTankerStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTruckSheetReport
                        frm = New rptTruckSheetReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptUnionBookingReport
                        frm = New rptUnionBookingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDailyQtyReport
                        frm = New rptDailyQtyReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTankerUtilization
                        frm = New rptTankerapprovalutilization
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmMonthWiseBeneficieriesReport
                        frm = New frmMonthWiseBeneficieriesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCGazeCapacityUtilizationReport
                        frm = New rptBMCGazeCapacityUtilizationReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCCapacityUtilizationReport
                        frm = New rptBMCCapacityUtilizationReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCPerodicalReport
                        frm = New rptBMCPerodicalReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptPerodicalDispatchReport
                        frm = New rptPerodicalDispatchReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptHeadLoadReport
                        frm = New rptHeadLoadReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDBTCappingReport
                        frm = New rptDBTCappingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDcsSecurity
                        frm = New rptDCSSecurity
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDCSCredit
                        frm = New rptDCSCredit
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDayWiseCol
                        frm = New rptDayWiseCollection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptMarginAmt
                        frm = New rptMarginAmount
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptCattleFeedSaleReport
                        frm = New rptCattleFeedSaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTransferToSavingReport
                        frm = New TransferToSavingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmTTSavingReport
                        frm = New frmTTSavingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBmcTankerProfitLossReport
                        frm = New rptBmcTankerProfitLossReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDcsPaymentReport
                        frm = New RptDcsPaymentReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptYearlyMonthlyDcsCollectionReport
                        frm = New rptYearlyMonthlyDcsCollectionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.DcsMonthWiseMilkCollectionReport
                        frm = New DcsMonthWiseMilkCollectionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.ZoneWiseReport
                        frm = New ZoneWiseReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptMilkProcurementReport
                        frm = New rptMilkProcurementReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDCSTruckSheetRegister
                        frm = New rptDCSTruckSheetRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptAvgSaleDetailReport
                        frm = New rptAvgSaleDetailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptdairydispatch
                        frm = New rptdailydispatch
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptRouteSaleRegister
                        frm = New RptRouteSaleRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RptSalesSummaryReport
                        frm = New RptSalesSummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptGatePassReport
                        frm = New GatePassReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptNewSalesReport
                        frm = New rptNewSalesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDemandReturnReport
                        frm = New rptDemandReturnReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.RouteWiseDemandReport
                        frm = New RouteWiseDemandReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDailyStatementReport
                        frm = New rptDailyStatementReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptSchemeSaleReport
                        frm = New rptSchemeSaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptCreditCustomerReport
                        frm = New rptCreditCustomerReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptD1D2Report
                        frm = New rptD1D2Report
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptZoneWiseReport1
                        frm = New rptZoneWiseReport1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptMonthWiseDMRReport
                        frm = New rptMonthWiseDMRReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptPaymentCycleWiseReport
                        frm = New rptPaymentCycleWiseReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTempTruckSheetCollectionReport
                        frm = New rptTempTruckSheetCollectionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptMobileAppMilkCollection
                        frm = New rptMobileAppMilkCollection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCMobileHistory
                        frm = New rptBMCMobileHistory
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.YearlyBillReport
                        frm = New YearlyBillReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.FrmLastDCSCollectionItem
                        frm = New FrmLastDCSCollectionItem
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptGheeAndCattleFeedDeductionStatementReport
                        frm = New rptGheeAndCattleFeedDeductionStatementReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmDBTRecoVsIncentiveReport
                        frm = New frmDBTRecoVsIncentiveReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmAutoAdditionDeductionReport
                        frm = New frmAutoAdditionDeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTruckSheetDailySummaryReport
                        frm = New rptTruckSheetDailySummaryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCTankerTestingReport
                        frm = New rptBMCTankerTesting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.FrmTotalDeductionReport
                        frm = New FrmTotalDeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBMCDailyQtyReport
                        frm = New rptBMCDailyQtyReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptDBTMonthWiseFarmerDetail
                        frm = New rptDBTMonthWiseFarmerDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptTemporaryPaymentDeductionSummary
                        frm = New rptTemporaryPaymentDeductionSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.CompareVoucherReport
                        frm = New frmCompareVoucherReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptBankSavingAdvice
                        frm = New rptBankSavingAdvice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmDCSWiseHoldPayments
                        frm = New frmDCSWiseHoldPayments
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptAutoMultipleAdditionDeduction
                        frm = New rptAutoMultipleAdditionDeduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptBmcCollection
                        frm = New rptBmcCollection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptDBTMilkPayment
                        frm = New rptDBTMilkPayment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptMilkPaymentSummary
                        frm = New rptMilkPaymentSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.rptMilkUnion
                        frm = New rptmilkunion
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptMilkUnion2
                        frm = New rptmilkunion2
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDbtStatusRCDF
                        frm = New rptDbtStatusRCDF
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptMachineSurveyRegister
                        frm = New rptMachineSurveyRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDBTStatusReport
                        frm = New rptDBTStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptUnionWiseMilkTankerCollectionDetail
                        frm = New rptUnionWiseMilkTankerCollectionDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.frmDCSDeductionReport
                        frm = New frmDCSDeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptTankerMilkQcVsERPEntry
                        frm = New rptTankerMilkQcVsERPEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    Case clsUserMgtCode.frmUnionWiseAppUserReport
                        frm = New frmUnionWiseAppUserReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.FrmCmSambalYojanaSubsidy
                        frm = New FrmCmSambalYojanaSunsidy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptDBTBankResponse
                        frm = New rptDBTBankResponse
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptUnionMilkPurchase
                        frm = New rptUnionMilkRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.latestPaymentProcess
                        frm = New LatestPaymentProcess
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.rptExeVersionReport
                        frm = New rptExeVersionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductionReport
                        frm = New rptProductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductionStoreReport
                        frm = New rptProductionStoreReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.frmPaymentProcessFarmer
                        frm = New frmPaymentProcessFarmer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptVSPItemIssue
                    '    frm = New RptVSPItemIssue
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptVSPAssetIssue1
                        frm = New RptVSPAssetIssue1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptPriceRateDifferenceReport
                    '    frm = New RptPriceRateDifferenceReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptMCCMilkStatus
                        frm = New RptMCCMilkStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptDispatchOfMilkTransfer
                        frm = New RptDispatchofmilkTransfer2
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.RptVLCTragetMasterReport
                    '    frm = New RptVLVCTragetMasterReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptMCCVLCTragetMonthWiseReport
                    '    frm = New RptMCCVLCTragetMonthWiseReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.RptBulkMilkMultiplePurchaseInvoice
                        frm = New RptBulkMilkMultiplePurchaseInvoice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptRoutewiseTPTimeTable
                        frm = New RptRoutewiseTPTimeTable
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '-------------------Milk Procurement module end------------------------------------'
                    'Case clsUserMgtCode.SublocationMaster
                    '    frm = New frmSubLocationMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ItemLocationMapping
                    '    frm = New frmTransferLocationMapping
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '------------------Sale Purchase Security------
                    Case clsUserMgtCode.FrmBankPermission
                        frm = New FrmBankPermission
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerPermission
                        frm = New FrmCustomerPermission
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmVendorPermission
                        frm = New FrmVendorPermission
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerPermissionReport
                        frm = New frmCustomerPermissionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Anubhooti 05-Aug-2014
                    '--------------------Human Resource-------------------------------------'
                    '--------------------SetUp-------------------------------------'
                    Case clsUserMgtCode.JobTitle
                        frm = New FrmJobTilte
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRParameterMaster
                        frm = New FrmHRParameterMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.EmployeeTypeMaster
                        frm = New frmHREmployeeTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQualificationMaster
                        frm = New FrmQualificationMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProfileMaster
                        frm = New FrmProfileMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRoundMaster
                        frm = New FrmRoundMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmchkList
                        frm = New FrmCheckListMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRelationMaster
                        frm = New FrmRelationMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSourceTypeMaster
                        frm = New FrmSourceTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSourceTypeDetail
                        frm = New FrmSourceTypeDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.HRBudgeting
                        frm = New FrmHRBudgeting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RequesitionEntry
                        frm = New FrmRequesitionEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RequesitionApproval
                        frm = New FrmRequesitionApprovel
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RequesitionClose
                        frm = New FrmCloseRequestion
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.HRIndustryType
                        frm = New frmHRIndustryType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.HRVerticalMaster
                        frm = New FrmHRVerticalMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''
                    ' Transaction
                    Case clsUserMgtCode.frmApplicantEntry
                        frm = New FrmApplicantEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmShortlist
                        frm = New FrmShortlist
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInterviewSchedule
                        frm = New FrmInterviewSchedule
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInterviewFeedback
                        frm = New FrmInterviewFeedback
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.OfferChkList
                        frm = New FrmOfferCheckList
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JOININGCHECKLIST
                        frm = New FrmJoiningChecklist
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmReferenceCheck
                        frm = New FrmReferenceCheck
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSalaryFitment
                        frm = New FrmSalaryFitment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOfferLetterHR
                        frm = New FrmOfferLetterHR
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmAppointmentLetterHR
                        frm = New frmAppointmentLetterHR
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHireEmployee
                        frm = New FrmHireEmployee
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHrTraineeFeedBack
                        frm = New FrmTraineeFeedback
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHrTrainerFeedBack
                        frm = New FrmHrTrainerFeedback
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDamageCaused
                        frm = New FrmDamageCaused
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ''
                    Case clsUserMgtCode.AgencyMaster
                        frm = New FrmAgencyMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDamageMaster
                        frm = New FrmDamageMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.TrainingMaster
                        frm = New frmTrainingMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TrainingResourceMaster
                        frm = New frmTrainingResourceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TrainingInstituteMaster
                        frm = New frmInstituteMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TrainingRequestMaster
                        frm = New frmRequestForTrainingMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TRAINER_MASTER
                        frm = New frmTrainerMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Schedule_Training
                        frm = New frmScheduleForTraining
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Training_Attendance
                        frm = New frmTrainingAttendance
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TrainingMaster
                        frm = New frmTrainingMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Performance Evaluation
                    Case clsUserMgtCode.frmHRPerformanceCategoryMaster
                        frm = New frmHRPerformanceCategoryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRPerformanceMaster
                        frm = New FrmHRPerformanceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRPerformanceGroup
                        frm = New FrmHRPerformanceGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRPerformanceGroupMapping
                        frm = New FrmPerformanceGroupMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRPerformanceRating
                        frm = New FrmPerformanceRating
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' Reimbursement
                    Case clsUserMgtCode.frmHRReimbursementTypeMaster
                        frm = New frmHRReimbursementTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelPurposeMaster
                        frm = New FrmHRTravelPurposeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelCategoryMaster
                        frm = New FrmHRTravelCategoryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelBookedForMaster
                        frm = New FrmHRTravelBookedForMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelModeTypeMaster
                        frm = New frmHRTravelModeTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelCityMaster
                        frm = New FrmHRTravelCityMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelClassTypeMaster
                        frm = New FrmHRTravelClassTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRHotelRatingMaster
                        frm = New FrmHRHotelRatingMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelRoomTypeMaster
                        frm = New FrmHRTravelRoomTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelCarTypeMaster
                        frm = New FrmHRTravelCarTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRRaiseTravelRequisition
                        frm = New FrmHRRaiseTravelRequisition
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelReqApproval
                        frm = New frmHRTravelReqApproval
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelReimbursementExpense
                        frm = New FrmHRTravelReimbursementExpense
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRTravelExpenseApproval
                        frm = New FrmHRTravelExpenseApproval
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmHRApprovarCreationMaster
                        frm = New FrmApproverCreationMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' ---------------------------- End HR --------------------------------------
                    ''''''''''''''''''''''''''''''''''''''''''''Visual Process Flow''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''Case clsUserMgtCode.FrmVPFSettings
                    ''    frm = New FrmVPFSettings
                    ''    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmVPFActiveReport
                    '    frm = New FrmVPFActiveScreens
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmGLCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmGLCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmCommonCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmCommonCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmAPCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmAPCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmARCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmARCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmBulkSaleCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmBulkSaleCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmFreshSaleCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmFreshSaleCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmProductSaleCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmProductSaleCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmCSASaleCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmCSASaleCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmMMCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmMMCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmMCCProcurementCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmMCCProcurementCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmBulkProcurementCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmBulkProcurementCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmPurchaseCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmPurchaseCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmDProductionCycle
                    '    frm = New FrmModuleCycle(clsUserMgtCode.FrmDProductionCycle)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '''''''''''''''''''''''''''''''''''''''''''' End Visual Process Flow ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  

                    '' Anubhooti 28-Aug-2015
                    '''''''''''''''''''''''''''''''''''''''''''' Service And Warranty ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '--------------------SetUp-------------------------------------'
                    Case clsUserMgtCode.FrmFaultCategory
                        frm = New FrmFaultCategory
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmFaultMaster
                        frm = New FrmFaultMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmServiceChargeMaster
                        frm = New FrmServiceChargeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProblemType
                        frm = New FrmProblemType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCallType
                        frm = New FrmCallType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmActivityType
                        frm = New FrmActivityType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSolutionKnowledgeBase
                        frm = New FrmSolutionKnowledgeBase
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmServiceCall
                        frm = New FrmServiceCall
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmServiceEnquiry
                        frm = New FrmServiceEnquiry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmServiceAllocation
                        frm = New FrmServiceAllocation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmServiceVisitDetails
                        frm = New FrmServiceVisitDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '''''''''''''''''''''''''''''''''''''''''''' End Service And Warranty ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
                    'Case clsUserMgtCode.FrmItemConversion
                    '    frm = New FrmItemConversion
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkReasonMaster
                        frm = New frmMilkReasonMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.GenratePaymentCycle
                        frm = New frmGeneratePaymentCycle
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPaymentCycleMaster
                        frm = New frmPaymentCycleMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ApproveFailedSample
                    '    frm = New frmApproveFailedSample()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.MilkMPPayment
                    '    frm = New FrmMilkVSPPayment(clsUserMgtCode.MilkMPPayment)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIncentiveMaster
                        frm = New frmIncentiveMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.MilkTruckSheet
                    '    frm = New FrmMilkTruckSheet
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.ApplyCapping
                    '    frm = New FrmApplyCapping(clsUserMgtCode.MilkVSPPayment)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkVSPPayment
                        frm = New FrmMilkVSPPayment(clsUserMgtCode.MilkVSPPayment)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MPBillGeneration
                        frm = New FrmMilkVSPPayment(clsUserMgtCode.MPBillGeneration)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkVSPIssuePayment
                        frm = New FrmMilkVSPPayment(clsUserMgtCode.MilkVSPIssuePayment)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkRecurringScheduler
                        frm = New FrmMilkRecurringScheduler
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '========================================MIS Reports===========================================
                    Case clsUserMgtCode.MISDebtorReport
                        frm = New FrmRptCustomerLedgerDemo(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISStockReco
                        frm = New FrmStockReco(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISSaleRegister
                        frm = New RptSaleRegisterReport(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTCSLedger
                        frm = New rptTCSLedger(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RMStockConsumption
                        frm = New RMStockConsumption
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.MSIProductionSaleReport
                        frm = New MSIProductionSaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SundryDebtorsReport
                        frm = New frmSundryDebtorsReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AccountStatementReport
                        frm = New rptAccountStatementReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptDCSSaleRegister
                    '    frm = New rptDCSSaleRegister()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISCreditorReport
                        frm = New frmRptVendorLedger(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptCustomerTransList
                        frm = New FrmRptCustomerTransList()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRptCustomerTransHistory
                        frm = New FrmRptCustomerTransHistory()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.MISStockLedgerReport
                    '    frm = New FrmStockReco(strProgramCode)
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)
                    Case clsUserMgtCode.MISStockLedgerReport
                        frm = New FrmStockReco(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmItemReloadReport
                        frm = New FrmItemReloadReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmMilkJobWork
                        frm = New frmMilkRGP
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkJobWorkTransfer
                        frm = New frmMilkJobWorkTransfer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkJobWorkTransferReturn
                        frm = New frmMilkJobWorkTransferReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkJobWorkTransferOther
                        frm = New frmJWOTransferOther
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkJobWorkTransferOtherReturn
                        frm = New frmJWOTransferOtherReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMilkGateEntry
                        frm = New FrmMilkGateEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMilkWeighment
                        frm = New FrmMilkWeighment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmJobMilkQualityCheck
                        frm = New FrmJobMilkQualityCheck
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmJobMilkSRN
                        frm = New FrmJobMilkSRN
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptJobWorkStatus
                        frm = New RptJobWorkStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMilkUnloading
                        frm = New FrmMilkUnloading
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptJobWorkRegister
                        frm = New RptJobWorkRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptJobWorOutwardPurchasekRegister
                        frm = New RptJobWorktPurchaseRegisterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmJobworkItemInStatusReport
                        frm = New frmJobworkItemInStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmJobworkTransfer
                        frm = New frmJobworkTransfer()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmJobworkChargesReport
                        frm = New FrmJobworkChargesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmJobworkSRNReceiptReport
                        frm = New frmJobworkSRNReceiptReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    'Case clsUserMgtCode.FrmMilkCleaning
                    '    frm = New FrmMilkCleaning
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    '========================================MIS Reports============================Ends Here======
                    '=================================TDS Module ====================Added by Preeti Gupta========================
                    Case clsUserMgtCode.frmIncomeTaxSlab
                        'frm = New frmIncomeTaxSlab
                        frm = New frmIncomeTaxTDSSlabMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.IncomeTaxTDSCalculation
                        frm = New frmIncomeTaxTDSCalculation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmEmployeeTDSRpt
                        frm = New frmEmployeeTDSRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmITSection
                        frm = New FrmITSection
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInvestmentType
                        frm = New FrmInvestmentType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHouseRentDeclaration
                        frm = New FrmHouseRentDeclaration
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmHRAExemptionRule
                        frm = New FrmHRAExemptionRule
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmInvestmentDeclaration
                        frm = New FrmInvestmentDeclaration
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MCCProvisonReport
                        frm = New frmRptMCCProvision
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MCCProvisonReport
                        frm = New frmRptMCCProvision
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmContractTanker
                        frm = New frmContractTanker
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSupplierMaster
                        frm = New frmSupplierMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDivertedContractor
                        frm = New frmDivertedContractor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkTypeMast
                        frm = New frmMilkTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmIntimation
                        frm = New frmIntimation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPOBulkProc
                        frm = New frmPoBulkProc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMilkGradeMaster
                        frm = New frmMilkGradeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBulkRoutMaster
                        frm = New FrmBulkRoutMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TankerCleaningItem
                        frm = New frmTankerCleaningItem
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmShiftMasterBulk
                        frm = New frmShiftMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '========================================END TDS Module======================================================

                    '==========================================================================================================
                    ' ADDED BY KUNAL TICKET : BM00000009674
                    '==========================================================================================================
                    Case clsUserMgtCode.FrmSaleSettingFreshDS
                        frm = New FrmSaleSetting(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'sanjay 26/july/2018 Ticket No- ERO/27/07/18-000387
                    Case clsUserMgtCode.FrmZoneMasterDS
                        frm = New FrmZoneMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'sanjay 26/july/2018 

                    Case clsUserMgtCode.frmRouteMasterDS
                        frm = New frmRouteMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmFreshTransactionApprovalDS
                        frm = New FrmTransactionApproval()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmSchemeMasterDairyDS
                        frm = New FrmSchemeMasterDairy()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerDeduction
                        frm = New frmCustomerDeduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmRouteFreightDetailsDS
                        frm = New FrmRouteFreightDetails()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLocationItemMapping
                        frm = New RptLocationItemMappingDS()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDistributorRouteTagging
                        frm = New frmDistributorRouteTagging()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailyDemand
                        frm = New frmDailyDemand()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUpdateAppLocation
                        frm = New frmUpdateAppLocation()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDBTUnionPayment
                        frm = New frmDBTUnionPayment()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCattleFeedProduction
                        frm = New rptCattleFeedProduction()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnionLastGstInvoice
                        frm = New frmUnionLastGstInvoice()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailyDemandReport
                        frm = New frmDailyDemandReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptHSNWiseSaleReport
                        frm = New rptHSNWiseSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptMilkSupplySaleReport
                        frm = New rptMilkSupplySaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmperdayDetailRpt
                        frm = New FrmperdayDetailRpt()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBookingVsDispatchReport
                        frm = New rptBookingVsDispatchReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCustItemWiseSaleReport
                        frm = New rptCustItemWiseSaleReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptNewTenderPartyListReport
                        frm = New rptNewTenderPartyListReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptItemAndShiftWiseSaleSummaryReport
                        frm = New rptItemAndShiftWiseSaleSummaryReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.SaleIncentiveMaster
                        frm = New frmSaleIncentiveMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.CardSale
                    '    frm = New CardSale()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmCustCategoryWiseDefaultItemUomMaster
                    '    frm = New frmCustCategoryWiseDefaultItemUomMaster()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmTranspoterDeductionMaster
                    '    frm = New frmTranspoterDeductionMaster2()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmItemSublocationMapping
                    '    frm = New frmItemSublocationMapping()
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptPlantCustomerDemand
                        frm = New RptPlantCustomerDemandReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDemandForsingleBranch
                        frm = New RptDemandForSingleBranch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmPrintDistributerInvoiceStatement
                        frm = New FrmPrintDistributerInvoiceStatement()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDistributerLedgerReport
                        frm = New frmDistributerLedgerReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPendingBooking
                        frm = New FrmPendingBookingReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerZone
                        frm = New FrmCustomerZoneReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmQuickDemandReport
                        frm = New FrmQuickDemandReport()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==========================================================================================================
                    ' ADDED BY KUNAL TICKET : BM00000009674 ENDED HERE
                    '==========================================================================================================

                    Case clsUserMgtCode.frmbookingdairy
                        frm = New frmBookingDairyMultipleCustomer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmbookingdairyFreshSale
                    '    frm = New frmDairyBookingCustomer_FreshSale
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDemandBooking
                        frm = New frmDemandBooking
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProductDemandBooking
                        frm = New FrmProductDemandBooking
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmProductDispatch
                        frm = New FrmProductDispatch
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemandAdjustment
                        frm = New frmDemandAdjustment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemandUploader
                        frm = New frmDemandUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemandApproval
                        frm = New frmDemandApproval
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDemand_Sheet
                        frm = New frmDemand_Sheet()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmQuickDemand
                        frm = New frmQuickDemand()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductQuickDemand
                        frm = New FrmProductQuickDemandBooking()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmBoothCommission
                        frm = New FrmBoothCommission
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeakedSaleReturn
                        frm = New frmLeakedSaleReturn()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDCSDEmandBooking
                        frm = New frmDCSDemandBooking()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBoothDispatch
                        frm = New frmBoothDispatch()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCashIndentBookingMobApp
                        frm = New frmCashIndentBookingMobApp
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.POSDairySale
                        frm = New frmSNPOS
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmTranspoterDeduction
                        frm = New frmTranspoterDeduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmNotepadFileMatching
                        frm = New frmNotepadFileMatching
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDairyBookingUploader
                        frm = New frmDairyBookingUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGenerateBookingFreshSale
                        frm = New FrmGenerateFreshBooking
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerTender
                        frm = New frmCustomerTender
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalesOrder
                        frm = New FrmSalesOrder
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalesOrderDispatch
                        frm = New FrmSalesOrderDispatch
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSalesOrderGatePass
                        frm = New FrmSalesOrderGatePass
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmAPSSaleReport
                        frm = New FrmAPSSaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAdvanceForCD
                    '    frm = New frmCustomerAdvanceAgainstBooking
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmAcknowledgeMentOfGRN
                    '    frm = New FrmAcknowledgeOfGRN
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeliveryOrderDairy
                        frm = New frmDeliveryNoteDairySale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleDispatchDairy
                        frm = New frmShipmentDairy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmGatePassDairy
                        frm = New frmGatePassDairySale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleInvoicedairy
                        'sanjay
                        'frm = New frmSaleInvoiceDairy
                        frm = New frmSaleInvoiceProductSale(strProgramCode)
                        'frm = New frmSaleInvoiceProductSale()
                        'sanjay
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSaleReturndairy
                        frm = New frmSaleReturnDairy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDispatchReturn
                        frm = New FrmDispatchReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCrateReceviedDairySale
                        frm = New frmCreateReceivedDairySale
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDairyBookingCustomer
                        frm = New frmDairyBookingCustomer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPerformaInvoiceDairy
                        frm = New frmPerformaInvoiceDairy
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDairyFreshDispatchMultiple
                        frm = New frmDairyFreshDispatchMultiple
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMultipleInvoice
                        frm = New frmMultipleInvoice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCommissionTPTCalculation
                        frm = New frmCommissionTPTCalculation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerPenalty
                        frm = New frmCustomerPenalty
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRouteWiseSaleTarget
                        frm = New frmRouteWiseSaleTarget
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRouteWiseSaleTargetReport
                        frm = New frmRouteWiseSaleTargetReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCommissionTPTCalculation
                        frm = New frmCommissionTPTCalculation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmItemProjection
                    '    frm = New FrmItemProjection
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.CustomerIncentiveEntry
                        frm = New frmCustomerIncetiveEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '=========
                    Case clsUserMgtCode.frmChangePasswordScreen
                        frm = New frmChangePasswordScreen()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptEffectiveRateReport1
                        frm = New RptEffectiveRateReport1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCustomerEffective_ItemRate
                        frm = New RptCustomerEffective_ItemRate
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptDeliveryOrderReport1
                        frm = New RptDeliveryOrderReport1
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBookingDairyMultipleDistributor
                        frm = New frmBookingDairyMultipleDistributor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmPOSBookingDairyMultipleDistributor
                        frm = New frmPOSBookingDairyMultipleDistributor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)



                    '=======================================================================================================================
                    ' = SILAGE PRODUCTION FORMS ==
                    '=======================================================================================================================
                    '== Added by kunal for Silage Production Forms ========================================================================

                    '    ''comment by balwinder due to solution auto close
                    'Case clsUserMgtCode.frmSilageAreaMaster
                    '    frm = New frmSilageAreaMaster
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    'Case clsUserMgtCode.frmSilageCriteriaMaster
                    '    frm = New frmSilageCriteriaMaster
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    'Case clsUserMgtCode.frmSilageProductionApplication
                    '    frm = New frmSilageProductionApplication
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    'Case clsUserMgtCode.frmSilageEnterPrenur
                    '    frm = New frmSilageEnterPrenur
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    'Case clsUserMgtCode.frmSilageFormerProduction
                    '    frm = New frmSilageFarmerSelection
                    '    formShow(frm,strProgramCode, strProgramName, isOpenInMDI, strDocNo)

                    '    ''end of comment by balwinder due to solution auto close


                    Case clsUserMgtCode.frmSilageTankerTransporterMaster
                        frm = New frmTankerTransporterMaster(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageTankerMaster
                        frm = New FrmTankerMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageParameterMaster
                        frm = New FrmParameterMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSialgeParameterRangeMaster
                        frm = New FrmParameterRangeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSialgeParameterRangeMasterForQC
                        frm = New frmParameterRangeMasterForQC(clsUserMgtCode.frmParameterRangeMasterForQC)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSialgeParameterValueMaster
                        frm = New FrmParameterValueMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilagePriceChartBulkProc
                        frm = New frmPriceChartBulkProc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageVendorPriceChartMapping
                        frm = New frmPriceChartBulkProc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSialgeSupplierMaster
                        frm = New frmSupplierMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageDivertedContractor
                        frm = New frmDivertedContractor
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageGradeMaster
                        frm = New frmGradeMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '' transaction 
                    Case clsUserMgtCode.frmSilageGateEntry
                        frm = New FrmGateEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageWeighment
                        frm = New FrmWeighment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageQualityCheck
                        frm = New FrmQualityCheck
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageUnloading
                        frm = New FrmUnloading
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSialgeCleaning
                        frm = New FrmCleaning
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageGateOut
                        frm = New FrmGateOut
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmSilageBulkSRN
                        frm = New FrmBulkMilkSRN
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSilageBulkSRNReturn
                        frm = New FrmBulkMilkSRNReturn
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MilkReject
                        frm = New frmMilkRejectEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.rptTankerDispatchWidthDeduction
                    '    frm = New frmRptTankerDispatchWithDeduction
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.MCCDispatchReturn
                    '    frm = New frmTankerDispatchReturn
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ' ================== END OF SILAGE PRODUCTION FORMS ===========================================================================
                    '======Sanjeet(21/11/2016)============
                    'Case clsUserMgtCode.FrmTruckSheetRouteWiseRpt
                    '    frm = New FrmTruckSheetRouteWiseRpt
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '23/11/2016
                    ''Case clsUserMgtCode.FrmMccWeightDifferenceRpt
                    ''    frm = New FrmMCCWeightDiifferenceRpt
                    ''    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '09/01/2017
                    'Case clsUserMgtCode.RptAClassMilkRate
                    '    frm = New RptAClassMilkRate
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '11/01/2017
                    Case clsUserMgtCode.RptPendingPO
                        frm = New RptPendingPO
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '12/01/2017
                    Case clsUserMgtCode.RptIssueReturnHirerachyWise
                        frm = New RptIssueReturnHirerachyWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMccSaleAdjustment
                        frm = New RptMccSaleAdjustment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '==========
                    Case clsUserMgtCode.rptVLCwiseTPTimeTable
                        frm = New RptVLCwiseTPTimeTable
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '=========Parteek added form 10-09-2017
                    Case clsUserMgtCode.frmProductBooking
                        frm = New frmProductBooking()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMccBulkmilkRegister
                        frm = New RptMccBulkMilkRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptDailyStanderdMilkQtyMCCWise
                        frm = New RptDailyStanderdMilkQtyMCCWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMCCRouteTimeTable
                        frm = New rptMCCRouteTimeTable
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptTankerInTransit
                        frm = New RptTankerInTransit
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptRouteWiseTrendBar
                    '    frm = New RptRouteWiseTrendBar
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMccCollectionDetails
                        frm = New rptMccCollectionDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMccMasterDetail
                        frm = New rptMccMasterDetail
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptListofCowDCS
                        frm = New rptListofCowDCS
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)




                    Case clsUserMgtCode.rptMccProcurementUploader
                        frm = New rptMccProcurementUploaderReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMCCWiseAbstractReport
                        frm = New rptMCCWiseAbstractReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCollectionDataChangeReport
                        frm = New rptCollectionDataChangeReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.rptFarmerPaymentApprovalReport
                    '    frm = New rptFarmerPaymentApprovalReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptMCCPaymentSummary
                        frm = New rptMCCPaymentSummary
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptWeighmentRegister
                        frm = New RptWeighmentRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RptVLCVehicleWeigmentRegister
                        frm = New RptVLCVehicleWeigmentRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMilkReceiptImproperWeight
                        frm = New FrmRptMilkReceiptImproperWeight
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptImproperMilkSample
                    '    frm = New FrmRpImproperMilkSample
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDailyWiseMilkCost
                        frm = New RptDailyWiseMilkCost
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.RptDailyLandedCost
                    '    frm = New FrmVLCDailyLandedCost
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSaleVsReceipReport
                        frm = New FrmSaleVsReceipReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCustomerIncentiveEntry
                        frm = New rptCustomerIncentiveEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmCostCentreConsumptionRpt
                        frm = New FrmCostCentreConsumptionRpt
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPOAgainstDocument
                        frm = New FrmPurchaseReportDocument
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBillChecklist
                        frm = New frmBillChecklist
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCancelledTransactions_Purchase
                        frm = New frmCancelledTransactions_Purchase
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDairyGatePass
                        frm = New frmDairyGatePass
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMccScrapGatePass
                        frm = New frmMccGatePass
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMccMatSaleUploader
                        frm = New frmMCCMaterialSaleUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmUnitMaster
                        frm = New FrmUnitMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCostCenterTypeMaster
                        frm = New FrmCostCetreTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMaterialSalePriceChart
                        frm = New FrmMaterialSalePriceChart
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptPendingDocumentList
                        frm = New RptPendingDocumentList
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptApprovalReport
                        frm = New RptApprovalReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptUserScreenRightsReport
                        frm = New rptUserScreenRightsReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptScreenSettingReport
                        frm = New rptScreenSettingReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Ticket No  TEC/06/09/19-001003 ,Sanjay
                    Case clsUserMgtCode.rptDeleteHistoryReport
                        frm = New rptDeleteHistoryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RCDFDashboard
                        frm = New RCDFDashboard
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RCDFUnionDashboard
                        frm = New RCDFUnionDashboard
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.Complainfeedback
                        frm = New RptComplainfeedback
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DashboardMilkUnion
                        frm = New DashboardMilkUnion
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.DairySaleDashboard
                        frm = New frmDairyDashboard ''DairySaleDashboard
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDBTDashboard
                        frm = New rptDBTDashboard
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.SalesMarketingDashboard
                        frm = New SalesMarketingDashboard
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MostUserScreen
                        frm = New MostUserScreen()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.rptSMSDetailsReport
                        frm = New rptSMSDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '===========Sanjeet(21/02/2017)======

                    ''Added by Parteek on 23/08/2017
                    Case clsUserMgtCode.frmJobWorkoutwordMaster
                        frm = New frmJWPriceCodeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmVendorItemChargeMaster
                        frm = New frmVendorItemChargeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    'BSP
                    Case clsUserMgtCode.JWOParameterMaster
                        frm = New frmJWParameterMaster()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWOParameterFormula
                        frm = New frmJWOFormulaMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWOFormulaVendorMapping
                        frm = New frmJWOVendorFormulaMapping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.JWO_Estimation
                        frm = New frmJWOEstimate
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)



                    ''======End
                    '===========Panch Raj(27/02/2017)======
                    Case clsUserMgtCode.frmFarmerLedgerReport
                        frm = New frmFarmerLedgerReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptVspFarmerPaymentDetail
                        frm = New frmVSPFarmerPaymentDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptFarmerNotMappedWithVSP
                        frm = New rptFarmerNotMappedWithVSP
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptFarmerAbstractReport
                        frm = New rptFarmerAbstractReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCFarmerMappingFP
                        frm = New FrmMCCFarmerMapping()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterialFarmer
                        frm = New frmMCCMaterialSaleFarmer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterialFarmerUploader
                        frm = New frmMCCMaterialSaleFarmerUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMCCMaterialSaleReturnFarmer
                        frm = New frmMccMaterialSaleReturnFarmer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmFarmerPaymentAdjustment
                        frm = New frmFarmerPaymentAdjEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmDailySaleReport
                        frm = New FrmDailySaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmMonthlySaleReport
                        frm = New FrmMonthlySaleReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmCustomerGroupOutstanding
                        frm = New FrmCustomerGroupOutstanding
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmStockAgeingAnalysisReport
                        frm = New FrmStockAgeingAnalysisReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.mbtnPurchaseJobWork
                        frm = New frmPurchaseJobwork(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MonthlyProgressReport
                        frm = New frmMonthlyProgressReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.DCSDeductionBalanceReport
                    '    frm = New DCSDeductionBalanceReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSuspanceDCSReport
                        frm = New rptSuspanceDCSReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTankerGainLossReport
                        frm = New rptTankerGainLossReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.FrmCategoryAnalysisReport
                        frm = New FrmCategoryAnalysisReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPromptMsgPendindDoc
                        frm = New RptPrmoptMsgRelatedToPendencyDoc
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISMassBalanceReport
                        frm = New rptMassBalanceReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptSecondaryTransporterReport
                        frm = New RptSecondaryTransporterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RpttankerReport
                        frm = New RpttankerReportForErode
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptFlavouredMilk
                        frm = New RptFlavouredMilk
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmItemTypeMaster
                        frm = New FrmItemTypeMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptMonthWiseSaleAnalysis
                        frm = New RptMonthWiseSaleAnalysis
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmjobWorkDebitNote
                    '    frm = New frmjobWorkDebitNote
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptJobWorkDebitNoteReport
                        frm = New RptJobWorkDebitNoteReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSAC
                        frm = New frmSAC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmSACWiseTax
                        frm = New FrmSacWiseTaxMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.FrmRackBinMaster
                        frm = New frmRackBinMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmDeptHeadCustomerMapping
                    '    frm = New frmDeptHeadCustomerMapping
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    '===========end======
                    Case clsUserMgtCode.FrmItemWiseTax
                        frm = New FrmItemWiseTax
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmConfigureSynchronization
                    '    frm = New frmConfigureSynchronization
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.FrmSendSMSMultipleUser
                    '    frm = New FrmSendSMSMultipleUser
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.UtilityImportExport
                    '    frm = New frmImplementImportExport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSendSMSEmailSetting
                        frm = New FrmSendSMSEmailSetting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmRCDFRateControl
                        frm = New frmRCDFRateControl
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmOfficeOrder
                        frm = New frmOfficeOrder
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmLockMPCollectionPC
                        frm = New frmLockMPCollectionPC
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    'Case clsUserMgtCode.frmItemTaxRate
                    'frm = New frmItemTaxRate
                    'formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmFarmerPaymentEntry
                        frm = New frmFarmerPaymentEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmFarmerMilkPurchaseInvoice
                        frm = New frmFarmerMilkPurchaseInvoice
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptCustomerAdvanceReg
                        frm = New frmAdvanceRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    '====================Added by preeti Gupta[14/03/2018]=========================
                    Case clsUserMgtCode.frmsaleReturnGateEntryFS
                        frm = New FrmsaleReturnGateEntry(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmsaleReturnGateEntryPS
                        frm = New FrmsaleReturnGateEntry(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmsaleReturnGateEntryMISSAle
                        frm = New FrmsaleReturnGateEntry(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmTender
                        frm = New frmTender
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.TenderShortPenalty
                        frm = New frmTenderShortPenalty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RALNOC
                        frm = New frmRALNOC()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmShortSupplyPenalty
                        frm = New frmShortSupplyPenalty()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCorrectionforWrongEntry
                        frm = New frmCorrectionforWrongEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDeletionForEntry
                        frm = New frmDeletionForEntry()
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.frmsaleReturnGateEntryMCCSAle
                    '    frm = New FrmsaleReturnGateEntry(strProgramCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    ' Ticket No : KDI/26/04/18-000277  By prabhakar  ( Tester Remarks )
                    'Case clsUserMgtCode.frmsaleReturnGateEntryBulkSAle
                    '    frm = New FrmsaleReturnGateEntry(strProgramCode)
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmsaleReturnGateEntryExportSAle
                        frm = New FrmsaleReturnGateEntry(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmsaleReturnGateEntryCSATransfer
                        frm = New FrmsaleReturnGateEntry(strProgramCode)
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailReceivable
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailPayables
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.VendorPaymentDetails
                        frm = New RptVendorPaymentDetails
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptPaymentPayable
                        frm = New rptPaymentPayable
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.AuditTrailDairySale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailSystemAdmin
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailCommonServices
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailGeneralLedger
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptCostCenterReport
                        frm = New rptCostCenterReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailSaleAndDistribution
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptDCSSaleRegister
                        frm = New rptDCSSaleRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.RoutePartyWisePriceList
                        frm = New RoutePartyWisePriceList
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.saleconsignee
                        frm = New SaleConsignee
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSalesReport
                        frm = New rptSalesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptDealerSalesReport
                        frm = New rptDealerSalesReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptPlantWiseAvailableStockQtyReport
                        frm = New rptPlantWiseAvailableStockQtyReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptSaleReportCustomerWise
                        frm = New rptSaleReportCustomerWise
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailMaterialManagement
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailPurchase
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailTaxDeductedAtSource
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailFixedAssets
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.RptAssetDispatchRetailer
                        frm = New rptAssetDispatchRetailer
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)


                    Case clsUserMgtCode.rptMultipleDeductionReport
                        frm = New rptMultipleDeductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailHRAndPayroll
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailStandardProduction
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptSPItemConsumptionReport
                        frm = New rptSPItemConsumptionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ProductionReport
                        frm = New ProductionReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    'Case clsUserMgtCode.QualitySummaryReport
                    '    frm = New QualitySummaryReport
                    '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.AuditTrailDairyProduction
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductionStatusReport
                        frm = New rptProductionStatusReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptProductionStandardization
                        frm = New rptProductionStandardization
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.ProductionEntryReport
                        frm = New ProductionEntryReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptTankerProfitLossReport
                        frm = New rptTankerProfitLossReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.rptBmcTankerInOutReport
                        frm = New RptBmcTankerInOutReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionUtilityCost
                        frm = New frmProductionUtilityCost
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailMilkProcurementMCC
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailMilkProcurementBulk
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailHumanResource
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailBulkSale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.BulkSaleAcknowledgementUploaderReport
                        frm = New rptBulkSaleAcknowledgementUploaderReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailCSASale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailFreshSale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailProductSale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailExportSale
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailMerchantTrade
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailMilkJobWork
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailTDSPayroll
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailFarmerPayment
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailJobWorkOutward
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailJobWorkInward
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.AuditTrailElectrical
                        frm = New rptAuditTrailReport
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerComplaintMaster
                        frm = New frmCustomerComplaintMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmCustomerComplaint
                        frm = New frmCustomerComplain
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmLeakageReplacementUploader
                        frm = New frmLeakageReplacementUploader
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '==================================================================
                    Case clsUserMgtCode.ShareMaster
                        frm = New ShareMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmMultipleShareAllotment
                        frm = New frmMultipleShareAllotment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmShareAllotment
                        frm = New frmShareAllotment
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.rptShareRegister
                        frm = New rptShareRegister
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '---------------MIS------------------
                    Case clsUserMgtCode.MISitemGroups
                        frm = New frmMISItemGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.MISitemMaster
                        frm = New frmMISItemMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailyMilkProducts
                        frm = New frmDailyMilkProducts
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmProductionTransactionType
                        frm = New frmProductionTransactionType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmDailySMPProduction
                        frm = New frmDailySMPProduction
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        '------------Bull Management-----------
                    Case clsUserMgtCode.frmBullBreedMaster
                        frm = New frmBullBreedMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullTestParameter
                        frm = New frmBullTestParameter
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullParameterGroup
                        frm = New frmBullParameterGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullShedMaster
                        frm = New frmBullShedMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullShedParameterMaster
                        frm = New frmBullShedParameterMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullShedParameterGroup
                        frm = New frmBullShedParameterGroup
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullMasters
                        frm = New frmBullMasters
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullStatus
                        frm = New frmBullStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSubStatus
                        frm = New frmBullSubStatus
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSourcePainting
                        frm = New frmBullSourcePainting
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullRating
                        frm = New frmBullRating
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSourceName
                        frm = New frmBullSourceName
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullPurchaseNo
                        frm = New frmBullPurchaseNo
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullInsuranceMaster
                        frm = New frmBullInsuranceMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullInsuranceType
                        frm = New frmBullInsuranceType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmDiseaseMaster
                        frm = New frmDiseaseMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBullMovementType
                        frm = New frmBullMovementType
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmBullVaccinationEntry
                        frm = New frmBullVaccinationEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSpeciesMaster
                        frm = New frmBullSpeciesMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullCategoryMaster
                        frm = New frmBullCategoryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSubCategoryMaster
                        frm = New frmBullSubCategoryMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullSSCentreMaster
                        frm = New frmBullSSCentreMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullPenIDMaster
                        frm = New frmBullPenIDMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullMovement
                        frm = New frmBullMovement
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullCurlingEntry
                        frm = New frmBullCurlingEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullInsurance
                        frm = New frmBullInsurance
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullTestParameterEntry
                        frm = New frmBullTestParameterEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmBullCMUGrouping
                        frm = New frmBullCMUGrouping
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmInsuranceTagAllocation
                        frm = New frmInsuranceTagAllocation
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                    Case clsUserMgtCode.frmCMUChecklistEntry
                        frm = New frmCMUChecklistEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)




                        '------------- SHEED  -----------
                    Case clsUserMgtCode.frmSeedGrowerMaster
                        frm = New frmSeedGrowerMaster
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                    Case clsUserMgtCode.frmSeedSelectionEntry
                        frm = New frmSeedSelectionEntry
                        formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        '-------------------- MIS Master---------------

                        'Case clsUserMgtCode.MISitemGroups
                        '    frm = New frmMISI(lblUserCode.Text, objCommonVar.CurrentCompanyCode)
                        '    formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)

                        'Case Else
                        '    Dim dtt As DataTable = clsDBFuncationality.GetDataTable("select 'BI-RPT' as Code from TSPL_CREATE_BI_REPORT where Code='" + strProgramCode + "' union select 'BI-DBR' as Code from TSPL_CREATE_DASHBOARD where code='" + strProgramCode + "' ")
                        '    If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                        '        If clsCommon.CompairString(clsCommon.myCstr(dtt.Rows(0)("Code")), "BI-RPT") = CompairStringResult.Equal Then
                        '            frm = New FrmBIReport
                        '            frm.obj = clsCreateBIReport.GetData(strProgramCode, True, NavigatorType.Current)
                        '            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtt.Rows(0)("Code")), "BI-DBR") = CompairStringResult.Equal Then
                        '            frm = New frmDashboard
                        '            frm.objDB = clsCreateDashboard.GetData(strProgramCode, NavigatorType.Current)
                        '            formShow(frm, strProgramCode, strProgramName, isOpenInMDI, strDocNo, IFTrueShowFormElseShowDialog)
                        '        End If


                        'End If
                End Select



            End If
        End If

        frm = Nothing
    End Sub

    Private Sub RadContextMenu2_DropDownOpening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RadContextMenu2.DropDownOpening
        Try
            RadContextMenu2.Items.Clear()
            If clsCommon.CompairString(clsCommon.myCstr(RTV2.SelectedNode.Value), clsUserMgtCode.ExpertERP) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(RTV2.SelectedNode.Value), clsUserMgtCode.ModuleFavourite) = CompairStringResult.Equal Then
                e.Cancel = True
            End If

            Dim mbtnChangeCaption As New RadMenuItem("Change Caption")
            AddHandler mbtnChangeCaption.Click, AddressOf ChangeCaption
            RadContextMenu2.Items.Add(mbtnChangeCaption)
            If Not clsCommon.CompairString(clsCommon.myCstr(RTV2.SelectedNode.Value), clsUserMgtCode.ExpertERP) = CompairStringResult.Equal Then
                If clsCommon.CompairString(clsCommon.myCstr(RTV2.SelectedNode.Parent.Value), clsUserMgtCode.ModuleFavourite) = CompairStringResult.Equal Then
                    Dim mbtnRemoveFromFavourite As New RadMenuItem("Remove from Favourite")
                    AddHandler mbtnRemoveFromFavourite.Click, AddressOf RemoveFromFavourite
                    RadContextMenu2.Items.Add(mbtnRemoveFromFavourite)
                Else
                    Dim mbtnAddToFavourite As New RadMenuItem("Add To Favourite")
                    AddHandler mbtnAddToFavourite.Click, AddressOf AddToFavourite
                    RadContextMenu2.Items.Add(mbtnAddToFavourite)
                End If
            End If
        Catch ex As Exception
            RadMessageBox.Show(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ChangeCaption()
        Dim frm As New frmChangeCaption()
        frm.strProgramCode = clsCommon.myCstr(RTV2.SelectedNode.Value)
        frm.ShowDialog()
        If Not frm.isCancel_Flag Then
            LoadMenu()
        End If
    End Sub

    Private Sub AddToFavourite()
        Try
            clsFavouriteMenu.SaveData(RTV2.SelectedNode.Value)
            LoadMenu()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RemoveFromFavourite()
        Try
            clsFavouriteMenu.DeleteData(RTV2.SelectedNode.Value)
            LoadMenu()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Public Sub formShow(ByVal frm As FrmMainTranScreen, ByVal strProgramCode As String, ByVal strProgramName As String, ByVal isOpenInMDI As Boolean, ByVal strDocNo As String, Optional ByVal IFTrueShowFormElseShowDialog As Boolean = True)
        Try
            If SettingHighSecurityOnWeighingIntegratedScreen Then
                'If clsCommon.CompairString(strProgramCode, clsUserMgtCode.frmMCCWeighment) = CompairStringResult.Equal OrElse clsCommon.CompairString(strProgramCode, clsUserMgtCode.MilkGateEntryWeightment) = CompairStringResult.Equal OrElse clsCommon.CompairString(strProgramCode, clsUserMgtCode.frmMilkReceipt) = CompairStringResult.Equal OrElse clsCommon.CompairString(strProgramCode, clsUserMgtCode.frmWeighment) = CompairStringResult.Equal OrElse clsCommon.CompairString(strProgramCode, clsUserMgtCode.POWeighment) = CompairStringResult.Equal OrElse clsCommon.CompairString(strProgramCode, clsUserMgtCode.FrmWeighmentEntry) = CompairStringResult.Equal Then
                '    If Me.RadDock1.DockWindows.DocumentWindows.Where(Function(w) w.Text = strProgramName).Count() > 0 Then
                '        Me.RadDock1.ActivateWindow(Me.RadDock1.DockWindows.DocumentWindows.Where(Function(w) w.Text = strProgramName).First())
                '        Return
                '    End If
                'End If
            End If

            'sanjay GKD/20/06/18-000150 
            If SingleUserParticularDairyBookingEdit Then
                If clsCommon.CompairString(strProgramCode, clsUserMgtCode.frmbookingdairy) = CompairStringResult.Equal Then
                    If Me.RadDock1.DockWindows.DocumentWindows.Where(Function(w) w.Text = strProgramName).Count() > 0 Then
                        Me.RadDock1.ActivateWindow(Me.RadDock1.DockWindows.DocumentWindows.Where(Function(w) w.Text = strProgramName).First())
                        Return
                    End If
                End If
            End If
            'sanjay GKD/20/06/18-000150 

            frm.Tag = strDocNo
            frm.Text = strProgramName
            frm.SetUserMgmt(strProgramCode)

            If IFTrueShowFormElseShowDialog Then
                If isOpenInMDI Then
                    frm.MdiParent = Me
                Else
                    frm.WindowState = FormWindowState.Maximized
                End If
                frm.Focus()
                frm.Show()
                If isApplicationRun Then
                    isApplicationRun = False
                    'frm.WindowState = FormWindowState.Maximized
                    Application.Run(frm)
                Else
                    'frm.WindowState = FormWindowState.Maximized
                    frm.Show()
                End If
            Else
                frm.ShowDialog(Me)
                frm.TopMost = True


            End If
        Catch ex As Exception
            If Not ex.Message.Contains("Object reference not set to an instance of an object.") Then ''becuase when need to close the form this message come.
                clsCommon.MyMessageBoxShow(Me, ex.Message)
                frm.Close()
            End If
        End Try
    End Sub

    Private Sub btnEditCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditCaption.Click
        RTV2.CollapseAll()
        RTV2.Nodes(0).Expand()
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton2.Click
        RTV2.ExpandAll()
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton1.Click
        'Dim frm As New FrmCarousal(Me)
        'frm.MdiParent = Me
        'frm.Show()
        'frm.Focus()
    End Sub

    Private Sub txtUserName_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyDown, txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            CheckAndLogin()
        End If
    End Sub

    Private Sub RTV2_NodeFormatting(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.TreeNodeFormattingEventArgs) Handles RTV2.NodeFormatting
        If ArrBold.Contains(clsCommon.myCstr(e.Node.Value)) Then
            e.NodeElement.ContentElement.Text = "<html><b>" & e.Node.Text
        End If
    End Sub

    Private Sub RadMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem1.Click
        'Ticket No  ERO/26/08/19-001002 ,sanjay
        'Dim objLogin As UserLoginInfo = New UserLoginInfo
        'objLogin.Show()
        ShowForm(clsUserMgtCode.rptActiveUsers, "Active Users", True, "", True)
    End Sub





    Private Sub RadDock1_DockStateChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.Docking.DockWindowEventArgs) Handles RadDock1.DockStateChanged
        ' Set Image
        For i As Integer = 0 To RTV2.Nodes.Count - 1
            SetImage(RTV2.Nodes(i))


        Next
    End Sub

    Private Sub MDI_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            Dim strQ As String = "update tspl_user_master set IP_Address=NULL,Login_Status=0 where user_code='" + objCommonVar.CurrentUserCode + "'"
            clsDBFuncationality.ExecuteNonQuery(strQ)

            Dim frmCollection As New FormCollection()
            frmCollection = Application.OpenForms()
            If frmCollection.Item("frmMccDispatch").IsHandleCreated Then
                If FrmMccDispatch.isPortOpened Then
                    e.Cancel = True
                    clsCommon.MyMessageBoxShow("Can not Close Application. Please Close MCC Dispatch Form because it uses an Opened Serial Port")
                    Exit Sub
                End If
            End If

            If frmCollection.Item("frmQualityCheck").IsHandleCreated Then
                If FrmQualityCheck.isPortOpened Then
                    e.Cancel = True
                    clsCommon.MyMessageBoxShow("Can not Close Application. Please Close Quality Check Form because it uses an Opened Serial Port")
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try

        'sanjay
        Try
            Dim frmCollection1 As New FormCollection()
            frmCollection1 = Application.OpenForms()
            If frmCollection1.Item("frmBookingDairyMultipleCustomer").IsHandleCreated Then
                If frmBookingDairyMultipleCustomer.LockUnlock = 1 Then
                    e.Cancel = True
                    clsCommon.MyMessageBoxShow("Can not Close Application. Please Close Dairy Booking first.")
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try
        'sanjay

        If Not IsDBRestored Then
            If Not isAutoClosing Then
                If clsCommon.MyMessageBoxShow("Do you want to close/log off the Xpert ERP", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.No Then
                    e.Cancel = True
                    'Else
                    '    'GC.Collect()
                Else

                End If
            End If
        End If

        If th IsNot Nothing Then
            Try
                th.Abort()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub SplitPanel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SplitPanel2.Click

    End Sub

    Private Sub RadMenuItem4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadMenuItem4.Click
        System.Diagnostics.Process.Start(Application.StartupPath & "\HTMLHELPERP\usermanual.chm")
    End Sub

    Private Function GetMasterDBConnectionStr(ByVal strDBName As String) As String
        Try
            Dim strConn As String = clsDBFuncationality.connectionString ''clsCommon.myCstr(Configuration.ConfigurationSettings.AppSettings("connectionString"))
            strConn = clsCommon.ReplaceString(strConn, strDBName, "Master")
            Return strConn
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub btnBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestore.Click
        Try
            If clsCommon.myLen(txtDBSource.Text) <= 0 Then
                bDestination.Focus()
                Throw New Exception("Please select source file.")
            End If

            Dim strMsg As String = "You are going to restore" + Environment.NewLine
            strMsg += " DataBase : - '" + cmbDB.SelectedValue + "'" + Environment.NewLine
            strMsg += " Company : - '" + cmbDB.SelectedText + "'" + Environment.NewLine
            strMsg += " Are you sure?"
            If clsCommon.MyMessageBoxShow(strMsg, Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes Then
                If RestoreDataBase("" + cmbDB.SelectedValue + "") Then
                    clsCommon.ProgressBarHide()
                    clsCommon.MyMessageBoxShow(Me, "DataBase Restored Sucessfully.")
                    IsDBRestored = True
                    Application.Restart()
                End If
            End If
        Catch ex As Exception
            clsCommon.ProgressBarHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function RestoreDataBase(ByVal strDBName As String) As Boolean
        Dim conn As SqlConnection = Nothing
        Dim cmd As SqlCommand
        Try
            clsDBFuncationality.ExecuteNonQuery("Update TSPL_UserLogin_Info set Logout_DateTime=' " + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") + "'  where Login_Code ='" + objCommonVar.CurrentLoginID + "'")
            clsDBFuncationality.ExecuteNonQuery("ALTER DATABASE " + cmbDB.SelectedValue + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE")
            Dim ConStr As String = GetMasterDBConnectionStr(strDBName)
            conn = clsDBFuncationality.GetConnnection()
            If conn.State = ConnectionState.Open Then
                conn.Close()
                conn.Dispose()
            End If
            conn = New SqlConnection(ConStr)
            cmd = New SqlCommand("sp_StopDBProcess", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = conn
            conn.Open()
            cmd.Parameters.Add("@strDBName", SqlDbType.VarChar).Value = strDBName
            cmd.ExecuteNonQuery()
            Dim qry As String = "Restore database " + strDBName + " from Disk = '" + txtDBSource.Text + "'"
            cmd = New SqlCommand(qry, conn)
            cmd.CommandTimeout = 3600
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            clsCommon.ProgressBarHide()
            cmd = New SqlCommand("ALTER DATABASE " + cmbDB.SelectedValue + "  SET MULTI_USER", conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            conn.Dispose()
            clsDBFuncationality.SetConnection(objCommonVar.ConnString)
        End Try
        Return True
    End Function

    Private Sub bDestination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bDestination.Click
        Try
            OpenFileDialog1.InitialDirectory = "C:\"
            OpenFileDialog1.Title = "Open a DataBase File"
            OpenFileDialog1.Filter = "DataBase Files|*.bak"
            If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                txtDBSource.Text = OpenFileDialog1.FileName
            Else
                txtDBSource.Text = ""
            End If
        Catch ex As Exception
            RadMessageBox.Show("Error: " + ex.Message)
        End Try
    End Sub

    Private Sub RadButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton13.Click
        LoadLoginScreen()
    End Sub

    Private Sub RTV2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RTV2.DoubleClick


    End Sub

    Private Sub RTV2_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.RadTreeViewEventArgs) Handles RTV2.NodeMouseDoubleClick
        Try
            Dim strCode As String = clsCommon.myCstr(RTV2.SelectedNode.Value)
            If clsCommon.myLen(strCode) > 0 Then
                ShowForm(strCode, RTV2.SelectedNode.Text, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MDI_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus

    End Sub

    Private Sub mnuRefreshMem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefreshMem.Click
        GC.Collect()
        GC.WaitForPendingFinalizers()
        clsCommon.MyMessageBoxShow("Memory Refreshed ")
    End Sub

    Private Sub RadMenuItem3_Disposing(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

#Region "Reminder Code" '-----------By Monika--------04/07/2014----------BM00000003039
    'Private Sub ReminderTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReminderTimer.Tick
    '    Try
    '        Dim ccccc As Integer = 0
    '        Dim cn As SqlConnection = clsDBFuncationality.GetConnnection()
    '        Dim Qry As String = "select count(*) from sys.dm_tran_database_transactions where isnull(database_transaction_status,0)>0"
    '        Dim cmd As SqlCommand = New SqlCommand(Qry, cn)
    '        cn.Close()
    '        cn.Open()
    '        Try
    '            ccccc = CInt(cmd.ExecuteScalar())
    '        Catch exx As Exception
    '            ccccc = 0
    '        End Try
    '        cn.Close()

    '        If ccccc > 0 Then
    '            Exit Sub
    '        End If

    '        Qry = "select count(*) from information_schema.TABLES where table_name='TSPL_DISPLAY_NOTIFICATIONS'"
    '        ccccc = clsDBFuncationality.getSingleValue(Qry)
    '        If ccccc <= 0 Then
    '            Return
    '        End If

    '        Dim xtime As String = ""
    '        xtime = clsCommon.myCstr(clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm:ss tt"))
    '        Qry = "select 'Trans_Id : '+doc_id+' Notification : '+message+'(Detail :'+item_name+ ')' as values1 from TSPL_DISPLAY_NOTIFICATIONS where user_code='" + objCommonVar.CurrentUserCode + "' and status<>'1' and isnull(snooze_time,'')<='" + xtime + "'"
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
    '        Dim str As String = ""

    '        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '            For Each dr As DataRow In dt.Rows()
    '                str = clsCommon.myCstr(dr("values1"))
    '                If clsCommon.myLen(str) > 0 Then
    '                    RadDesktopAlert1.AutoClose = False
    '                    RadDesktopAlert1.ShowOptionsButton = False
    '                    RadDesktopAlert1.ShowCloseButton = False


    '                    radbuttonelement.Tag = str
    '                    radbuttonDontShow.Tag = str


    '                    RadDesktopAlert1.FixedSize = New Size(529, 100)
    '                    RadDesktopAlert1.CaptionText = "Notification :"
    '                    RadDesktopAlert1.PopupAnimation = True
    '                    RadDesktopAlert1.ContentText = str
    '                    RadDesktopAlert1.Show()

    '                    arralert.Add(str, RadDesktopAlert1)
    '                    ReminderTimer.Enabled = False
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub

    Private Sub radbuttonelement_Click(ByVal sender As Object, ByVal e As EventArgs)
        snoozeOrDontShowAgain(sender, True)
        ReminderTimer.Enabled = True
    End Sub

    Private Sub DontShowAgain_Click(ByVal sender As Object, ByVal e As EventArgs)
        snoozeOrDontShowAgain(sender, False)
        ReminderTimer.Enabled = True
    End Sub

    Sub snoozeOrDontShowAgain(ByVal sender As Object, ByVal issnoozed As Boolean)
        Dim radButtonElement As RadButtonElement = TryCast(sender, RadButtonElement)
        Dim strCode As String = clsCommon.myCstr(radButtonElement.Tag)
        If clsCommon.myLen(strCode) > 0 Then
            If arralert.ContainsKey(strCode) Then
                arralert(strCode).Hide()
                arralert.Remove(strCode)
                If issnoozed Then
                    clsfrmNotificationScreen.Snooze(strCode)
                Else
                    clsfrmNotificationScreen.DontShowAgain(strCode)
                End If
            End If
        End If
    End Sub
#End Region

    Private Sub RadMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem6.Click, RadMenuItem7.Click, RadMenuItem8.Click, RadMenuItem9.Click, RadMenuItem10.Click, RadMenuItem11.Click, RadMenuItem12.Click, RadMenuItem13.Click, RadMenuItem14.Click, RadMenuItem15.Click, RadMenuItem16.Click, RadMenuItem17.Click, RadMenuItem18.Click, RadMenuItem19.Click, RadMenuItem20.Click, RadMenuItem21.Click, RadMenuItem22.Click, RadMenuItem25.Click, RadMenuItem26.Click, RadMenuItem27.Click, RadMenuItem28.Click, RadMenuItem29.Click, RadMenuItem30.Click, RadMenuItem31.Click, RadMenuItem32.Click, RadMenuItem33.Click, RadMenuItem34.Click, RadMenuItem35.Click, RadMenuItem36.Click
        Try
            Dim OldThemeName As String = ""
            Dim clickedCtrl As Telerik.WinControls.UI.RadMenuItem = DirectCast(sender, Telerik.WinControls.UI.RadMenuItem)
            ThemeResolutionService.ApplicationThemeName = clickedCtrl.Text
            OldThemeName = ""
            Dim FILE_NAME As String = Application.StartupPath + "\Theme.Txp"
            If System.IO.File.Exists(FILE_NAME) Then
                '==============read theme name from existing file============
                Dim objreader As New System.IO.StringReader(FILE_NAME)
                If objreader IsNot Nothing AndAlso clsCommon.myLen(objreader) > 0 Then
                    OldThemeName = clsCommon.myCstr(objreader.ReadToEnd())

                End If
                '==================================
                System.IO.File.Delete(FILE_NAME)
            End If
            File.Create(FILE_NAME).Dispose()
            Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
            objWriter.Write(clickedCtrl.Text)
            Me.OldThemeName = clickedCtrl.Text
            objWriter.Close()
            'Me.OldThemeName = FILE_NAME
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadTheme()
        Try
            Dim line As String
            Dim objReader As New System.IO.StreamReader("Theme.Txp")
            Do While objReader.Peek() <> -1
                line = objReader.ReadLine()
                ThemeResolutionService.ApplicationThemeName = line
                OldThemeName = line
            Loop
            ''stuti regarding memory leakage
            objReader.Close()
            objReader.Dispose()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadMenuItem5_DropDownOpening()
        Try

            If OldThemeName IsNot Nothing AndAlso OldThemeName.Length > 0 Then
                RadMenuItem6.IsChecked = False
                RadMenuItem7.IsChecked = False
                RadMenuItem8.IsChecked = False
                RadMenuItem9.IsChecked = False
                RadMenuItem10.IsChecked = False
                RadMenuItem11.IsChecked = False
                RadMenuItem12.IsChecked = False
                RadMenuItem13.IsChecked = False
                RadMenuItem14.IsChecked = False
                RadMenuItem15.IsChecked = False
                RadMenuItem16.IsChecked = False
                RadMenuItem17.IsChecked = False
                RadMenuItem18.IsChecked = False
                RadMenuItem19.IsChecked = False
                RadMenuItem20.IsChecked = False
                RadMenuItem21.IsChecked = False
                RadMenuItem22.IsChecked = False

                RadMenuItem25.IsChecked = False
                RadMenuItem26.IsChecked = False
                RadMenuItem27.IsChecked = False
                RadMenuItem28.IsChecked = False
                RadMenuItem29.IsChecked = False
                RadMenuItem30.IsChecked = False
                RadMenuItem31.IsChecked = False
                RadMenuItem32.IsChecked = False
                RadMenuItem33.IsChecked = False
                RadMenuItem34.IsChecked = False
                RadMenuItem35.IsChecked = False
                RadMenuItem36.IsChecked = False

                If clsCommon.CompairString(RadMenuItem6.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem6.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem7.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem7.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem8.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem8.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem9.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem9.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem10.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem10.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem11.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem11.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem12.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem12.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem13.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem13.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem14.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem14.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem15.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem15.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem16.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem16.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem17.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem17.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem18.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem18.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem19.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem19.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem20.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem20.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem21.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem21.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem22.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem22.IsChecked = True

                ElseIf clsCommon.CompairString(RadMenuItem25.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem25.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem26.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem26.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem27.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem27.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem28.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem28.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem29.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem29.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem30.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem30.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem31.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem31.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem32.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem32.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem33.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem33.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem34.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem34.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem35.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem35.IsChecked = True
                ElseIf clsCommon.CompairString(RadMenuItem36.Text, OldThemeName) = CompairStringResult.Equal Then
                    RadMenuItem36.IsChecked = True
                End If
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub RadMenuItem5_DropDownOpening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RadMenuItem5.DropDownOpening
        RadMenuItem5_DropDownOpening()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Try
            Dim IdleSec As Long = DateDiff(DateInterval.Second, clsDBFuncationality._LastActiveTime, DateTime.Now)
            '    Me.Text = IdleSec
            If IdleSec > 0 Then
                If IdleSec > IdleTimeinSeconds Then
                    isAutoClosing = True
                    Me.Close() ''ERO/05/07/19-000675 by balwinder on 17/07/2019
                    'Application.Restart()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub MDI_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        Try
            clsDBFuncationality._LastActiveTime = DateTime.Now()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MDI_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Try
            clsDBFuncationality._LastActiveTime = DateTime.Now()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadDock1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RadDock1.MouseMove

    End Sub

    Private Sub lblChangePWD_Click(sender As Object, e As EventArgs) Handles btnChangePassword.Click
        Try
            PasswordRules = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PasswordRules, clsFixedParameterCode.PasswordRules, Nothing)) = "1", True, False))
            If clsCommon.myLen(txtUserName.Text) <= 0 Then
                Throw New Exception("Please Enter User Name")
            End If
            Dim isUserfound As Integer = 0
            isUserfound = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from tspl_user_master where user_code='" & txtUserName.Text & "'"))
            If isUserfound = 0 Then
                Throw New Exception("Invalid User Name")
            End If
            objCommonVar.CurrentUserCode = txtUserName.Text
            objCommonVar.CurrentUser = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select user_name from tspl_user_master where user_code='" & txtUserName.Text & "'"))

            Dim frm1 As RadForm = New FrmChangePassword()
            frm1.StartPosition = FormStartPosition.CenterScreen
            frm1.MaximizeBox = False
            frm1.MinimizeBox = False
            frm1.ControlBox = False
            frm1.Location = Me.PictureBox1.Location
            If PasswordRules = True Then
                frm1.Height = Me.PictureBox1.Height + 150
                frm1.Width = Me.PictureBox1.Width + 100
            Else
                frm1.Height = Me.PictureBox1.Height + 50
                frm1.Width = Me.PictureBox1.Width + 50
            End If


            frm1.ShowDialog()


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
    End Sub

    Private Sub Receipt_Click(sender As Object, e As EventArgs) Handles Receipt.Click

    End Sub

    Private Sub RadMenuItem23_Click(sender As Object, e As EventArgs) Handles RadMenuItem23.Click
        Dim frm As New FrmLicenceActivate()
        frm.ShowDialog()
    End Sub

    Private Sub OpenFormFromOtherDLL()
        Try
            th = New Threading.Thread(AddressOf OpenFormFromOtherDLLMain)
            th.Start()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub OpenFormFromOtherDLLMain()
        If clsCommon.myLen(objCommonVar.ScreenToOpen) > 0 Then
            th1 = New Threading.Thread(AddressOf ShowFormFinal)
            th1.Start()
        End If

        Dim i As Int64 = 0
        While i <= 400000000
            i = i + 1
        End While
        'Try
        '    th1.Abort()
        'Catch ex As Exception
        'End Try

        OpenFormFromOtherDLLMain()
    End Sub
    Sub ShowFormFinal()
        Dim localScreenToOpen As String = objCommonVar.ScreenToOpen
        Dim localDocToOpen As String = objCommonVar.ScreenToOpenDocNo
        objCommonVar.ScreenToOpenDocNo = ""
        objCommonVar.ScreenToOpen = ""
        __txtDocNo.Text = ""
        __txtScreenID.Text = ""
        If objCommonVar.ScreenToOpen_Text <> "" Then
            Dim frm As New frmBalanceQty
            frm.strUOM = objCommonVar.ScreenToOpenUOM
            frm.IsMRPMandatory = objCommonVar.ScreenToOpenIsMRPMandatory
            frm._METEXT = "Order Quantity"
            'If _IsMRPMandatory Then
            frm.qry = objCommonVar.ScreenToOpenQry
            'End If
            frm.Show()
        Else
            If clsCommon.myLen(localScreenToOpen) > 0 Then
                isApplicationRun = False
                ShowForm(localScreenToOpen, "", False, IIf(clsCommon.myLen(localDocToOpen) > 0, localDocToOpen, ""))
            End If
        End If

    End Sub

    Private Sub __txtDocNo_TextChanged(sender As Object, e As EventArgs) Handles __txtDocNo.TextChanged
        If clsCommon.myLen(__txtDocNo.Text) > 0 AndAlso clsCommon.myLen(__txtScreenID.Text) > 0 Then
            objCommonVar.ScreenToOpen = __txtScreenID.Text
            objCommonVar.ScreenToOpenDocNo = __txtDocNo.Text
            ShowFormFinal()
        End If
    End Sub

    Private Sub __txtScreenID_TextChanged(sender As Object, e As EventArgs) Handles __txtScreenID.TextChanged
        If clsCommon.myLen(__txtDocNo.Text) > 0 AndAlso clsCommon.myLen(__txtScreenID.Text) > 0 Then
            objCommonVar.ScreenToOpen = __txtScreenID.Text
            objCommonVar.ScreenToOpenDocNo = __txtDocNo.Text
            ShowFormFinal()
        End If
    End Sub

    Private csd200Obj As CgtFpAccessCSD200Dotnet.MMMCogentCSD200APIs
    Private Sub RadButton17_Click(sender As Object, e As EventArgs) Handles RadButton17.Click
        Try
            FingerPrintScanner()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub FingerPrintScanner()
        Try
            csd200Obj = New MMMCogentCSD200APIs()
            csd200Obj.initializeScanner()
            Dim num As Integer = -1
            Dim captureBytes As Byte() = Nothing
            Dim width As Integer = 0
            Dim height As Integer = 0
            Dim isoTemplateBytes As Byte() = Nothing
            Dim nfiq As Integer = 0
            pBoxFingerPrint1.Image = Nothing
            pBoxFingerPrint1.Refresh()
            pBoxFingerPrint2.Image = Nothing
            pBoxFingerPrint2.Refresh()

            num = csd200Obj.captureFP(&H7530, captureBytes, width, height, nfiq, isoTemplateBytes)
            Dim isBiomatrikFound As Boolean = False
            If (num = CSD200APICodes.SUCCESS) AndAlso (captureBytes IsNot Nothing) Then
                pBoxFingerPrint1.Image = frmFreeImage.CreateGreyscaleBitmap(captureBytes, width, height)
                Dim ms As New MemoryStream()
                pBoxFingerPrint1.Image.Save(ms, ImageFormat.Bmp)
                Dim data As Byte() = ms.GetBuffer()
                Dim isoTemplateBytesForMatch1 As Byte() = Nothing
                ExtractTemplate(data, width, height, isoTemplateBytesForMatch1)



                Dim qry As String = "select User_Code,Password,biometric_image from tspl_user_master where biometric_image is not null"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        ms = New MemoryStream()

                        data = DirectCast(dr("biometric_image"), Byte())
                        ms = New MemoryStream(data)
                        pBoxFingerPrint2.Image = Image.FromStream(ms)

                        ms = New MemoryStream()
                        pBoxFingerPrint2.Image.Save(ms, ImageFormat.Bmp)
                        Dim dataNew As Byte() = ms.GetBuffer()
                        Dim isoTemplateBytesForMatch2 As Byte() = Nothing

                        ExtractTemplate(dataNew, width, height, isoTemplateBytesForMatch2)

                        'clsCommon.MyMessageBoxShow("Next " + clsCommon.myCstr(dr("User_Code")))
                        If csd200Obj.matchTemplates(isoTemplateBytesForMatch1, isoTemplateBytesForMatch2) Then
                            txtUserName.Text = clsCommon.myCstr(dr("User_Code"))
                            txtPassword.Text = clsCommon.DecryptString(clsCommon.myCstr(dr("Password")))
                            CheckAndLogin()
                            isBiomatrikFound = True
                        End If

                    Next
                End If
                If Not isBiomatrikFound Then
                    clsCommon.MyMessageBoxShow("Invalid Fingerprint...")
                End If
            ElseIf num = CSD200APICodes.ERROR_TIMEOUT Then
                clsCommon.MyMessageBoxShow("Fingerprint Capture Timeout")
            Else
                clsCommon.MyMessageBoxShow("Fingerprint Capture Failed. ErrorCode: " + num)
            End If
        Catch exception As Exception
            MessageBox.Show(exception.Message)
        End Try
    End Sub


    Private Function ExtractTemplate(bImage As Byte(), width As Integer, height As Integer, ByRef bTemplateData As Byte()) As Integer
        Dim bExtract_Init As Boolean = True
        Dim [error] As Integer = -1
        Dim inParameter As Integer = 0
        Dim outMessage As New System.Text.StringBuilder(&H100)
        If Not bExtract_Init Then
            [error] = BioSdk710Wrapper.InitExtract("", "", inParameter, outMessage)
            If [error] < 0 Then
                Throw New BioSDK710Exception([error])
            End If
            bExtract_Init = True
        End If
        Dim outTemplateSize As Integer = 0
        outMessage = New System.Text.StringBuilder(&H100)
        Dim outTemplateData As Byte() = New Byte(2047) {}
        Dim destinationArray As Byte() = Nothing
        Dim num4 As Integer = BioSdk710Wrapper.ExtractTemplate(bImage, height, width, &HC5, &HC5, 0, outTemplateData, outTemplateSize, outMessage)
        If num4 < 0 Then
            Return num4
        End If
        If outTemplateSize > 0 Then
            destinationArray = New Byte(outTemplateSize - 1) {}
            Array.Copy(outTemplateData, destinationArray, outTemplateSize)
        End If
        bTemplateData = destinationArray
        Return outTemplateSize
    End Function





    Dim counter As Integer = 0
    Private Sub RadLabelElement1_DoubleClick(sender As Object, e As EventArgs) Handles RadLabelElement1.DoubleClick
        counter += 1
        If counter Mod 2 = 0 Then
            ShowForm(clsUserMgtCode.FrmUtilityForm, "", True)
        End If
    End Sub

    Private Sub RadMenuItem24_Click(sender As Object, e As EventArgs) Handles RadMenuItem24.Click
        Dim frm As New frmShortCutInfo()
        frm.ShowDialog()
    End Sub

    Private Sub RadSplitButton1_Click(sender As Object, e As EventArgs) Handles RadSplitButton1.Click
        If objCommonVar.ControlMForHelp = True Then
            RadMenuItem4.Text = "Help (Ctrl+M)"
        Else
            RadMenuItem4.Text = "Help (F1)"
        End If
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        Try
            If ShowNotificationWithoutSMSAPP = True Then
                Dim TempNotifiactionOpen As Boolean = False
                For Each frm As Form In Application.OpenForms
                    If frm.Name.ToString().Contains("frmPromptMsgNotification") Then
                        TempNotifiactionOpen = True
                        Exit For
                    End If
                Next
                If TempNotifiactionOpen = False Then
                    Dim PendingNotification As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(*) as Notification from TSPL_NOTIFICATION_DETAIL LEFT OUTER JOIN TSPL_NOTIFICATION_HEAD ON TSPL_NOTIFICATION_HEAD.Code = TSPL_NOTIFICATION_DETAIL.Code LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  =  TSPL_NOTIFICATION_DETAIL.User_Name   LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  = TSPL_USER_MASTER.EmployeeCode where TSPL_NOTIFICATION_DETAIL.Sender_Replay=0 and TSPL_USER_MASTER.user_code='" + objCommonVar.CurrentUserCode + "'"))
                    If PendingNotification > 0 Then
                        Dim frmPromptMsgNotification As New frmPromptMsgNotification()
                        frmPromptMsgNotification.Location = New Point(Screen.PrimaryScreen.WorkingArea.Width - frmPromptMsgNotification.Width, Screen.PrimaryScreen.WorkingArea.Height - frmPromptMsgNotification.Height)
                        frmPromptMsgNotification.Show()
                    End If
                End If
            ElseIf ShowNotificationInMDI > 0 Then
                SendNotificationSnooze()
                SendNotification()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SendNotificationSnooze()
        Try
            'Dim IPAddress As String = ""
            'IPAddress = CStr((Dns.GetHostByName(Dns.GetHostName()).AddressList.GetValue(0)).ToString())
            Dim qry As String = "SELECT * FROM TSPL_SNOOZE_NOTIFICATION LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_SNOOZE_NOTIFICATION.UserCode  = TSPL_USER_MASTER.User_Code where TSPL_SNOOZE_NOTIFICATION.NextReminderTime<=convert(datetime,'" & clsCommon.GetPrintDate(DateTime.Now, "dd/MM/yyyy hh:mm:ss") & "',103) and TSPL_USER_MASTER.user_code='" + objCommonVar.CurrentUserCode + "'" 'and TSPL_USER_MASTER.IP_For_Alert='" & IPAddress & "'
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim totalNo As Decimal = dt.Rows.Count
                Dim intPer As Int32 = 0
                Dim ii As Int32 = 0

                For Each dr As DataRow In dt.Rows
                    ii += 1
                    intPer = Convert.ToInt16((ii / totalNo) * 100)
                    'lblMsg.Text = "Sending Notification " & clsCommon.myCstr(ii) & "/" + clsCommon.myCstr(totalNo)
                    'notifyIcon1.Text = lblMsg.Text & " [ " + clsCommon.myCstr(intPer) & "% ]"
                    'radProgressBar1.Value1 = intPer
                    Dim client As System.Net.WebClient = New System.Net.WebClient()
                    Dim Caption As String, NotiText As String = ""
                    Caption = clsCommon.myCstr(dr("NotifyCaption"))
                    NotiText = clsCommon.myCstr(dr("NotifyText"))
                    Dim RadDesktopAlertA As RadDesktopAlert = New RadDesktopAlert()
                    RadDesktopAlertA.AutoClose = False
                    RadDesktopAlertA.ShowOptionsButton = False
                    RadDesktopAlertA.ShowCloseButton = True
                    RadDesktopAlertA.FixedSize = New Size(400, 175)
                    RadDesktopAlertA.CaptionText = Caption
                    RadDesktopAlertA.PopupAnimation = True
                    'RadDesktopAlert1.ContentImage = XpertAlertApp.Properties.Resources.alert
                    RadDesktopAlertA.ThemeName = clsCommon.myCstr(dr("UserCode"))
                    RadDesktopAlertA.ContentText = NotiText
                    radbuttonelementA = New RadButtonElement("Snooze")
                    radbuttonelementA.Tag = RadDesktopAlertA
                    radbuttonelementA.ButtonFillElement.GradientStyle = Telerik.WinControls.GradientStyles.Solid
                    radbuttonelementA.BorderElement.BoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders
                    radbuttonDontShowA = New RadButtonElement("Don't Show Again")
                    radbuttonDontShowA.Tag = RadDesktopAlertA
                    radbuttonDontShowA.ButtonFillElement.GradientStyle = Telerik.WinControls.GradientStyles.Solid
                    radbuttonDontShowA.BorderElement.BoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders
                    'radbuttonelementA.Click += radbuttonelementA_Click()
                    'radbuttonDontShow.Click += radbuttonDontShow_Click
                    'RadDesktopAlert1.ButtonItems.Add(radbuttonelement)
                    'RadDesktopAlert1.ButtonItems.Add(radbuttonDontShow)
                    RadDesktopAlertA.ButtonItems.Add(radbuttonelementA)
                    RadDesktopAlertA.ButtonItems.Add(radbuttonDontShowA)
                    AddHandler radbuttonelementA.Click, AddressOf radbuttonelementA_Click
                    AddHandler radbuttonDontShowA.Click, AddressOf radbuttonDontShowA_Click
                    RadDesktopAlertA.Show()
                    qry = "Update TSPL_SNOOZE_NOTIFICATION set LastReminderTime=convert(datetime,'" & clsCommon.GetPrintDate(DateTime.Now, "dd/MM/yyyy hh:mm:ss") & "',103),NextReminderTime=convert(datetime,'" + clsCommon.GetPrintDate(DateTime.Now.AddMinutes(10), "dd/MM/yyyy hh:mm:ss") & "',103) where NotifyText='" & NotiText & "' and NotifyCaption='" & Caption & "' and UserCode='" + clsCommon.myCstr(dr("UserCode")) & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry)
                    'lblMsg.Refresh()
                    'radProgressBar1.Refresh()
                    'Me.Refresh()
                    Application.DoEvents()

                    'If Not isAppStart Then
                    '    lblMsg.Text = ""
                    '    radProgressBar1.Value1 = 0
                    '    Exit For
                    'End If
                Next
            End If

        Catch ex As Exception
            'SetErrorRowInGrid("", "", ex.Message)
        End Try
    End Sub

    Private Sub SendNotification()
        Try
            'Dim IPAddress As String = ""
            'IPAddress = CStr((Dns.GetHostByName(Dns.GetHostName()).AddressList.GetValue(0)).ToString())
            Dim qry As String = "SELECT TSPL_NOTIFICATION_HEAD.Code,TSPL_NOTIFICATION_HEAD.Notification_Caption,TSPL_NOTIFICATION_HEAD.Notification_Text,TSPL_NOTIFICATION_DETAIL.User_Name, TSPL_USER_MASTER.IP_For_Alert" & " FROM TSPL_NOTIFICATION_DETAIL LEFT OUTER JOIN TSPL_NOTIFICATION_HEAD ON TSPL_NOTIFICATION_HEAD.Code = TSPL_NOTIFICATION_DETAIL.Code LEFT OUTER JOIN TSPL_EMPLOYEE_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  = TSPL_NOTIFICATION_DETAIL.User_Name " & " LEFT OUTER JOIN TSPL_USER_MASTER ON TSPL_EMPLOYEE_MASTER.EMP_CODE  = TSPL_USER_MASTER.EmployeeCode where TSPL_NOTIFICATION_DETAIL.Sender_Replay=0 and TSPL_USER_MASTER.user_code='" + objCommonVar.CurrentUserCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim totalNo As Decimal = dt.Rows.Count
                Dim intPer As Int32 = 0
                Dim ii As Int32 = 0

                For Each dr As DataRow In dt.Rows
                    ii += 1
                    intPer = Convert.ToInt16((ii / totalNo) * 100)
                    'lblMsg.Text = "Sending Notification " & clsCommon.myCstr(ii) & "/" + clsCommon.myCstr(totalNo)
                    'notifyIcon1.Text = lblMsg.Text & " [ " + clsCommon.myCstr(intPer) & "% ]"
                    'radProgressBar1.Value1 = intPer
                    Dim client As System.Net.WebClient = New System.Net.WebClient()
                    Dim Caption As String, NotiText As String = ""
                    Caption = clsCommon.myCstr(dr("Notification_Caption"))
                    NotiText = clsCommon.myCstr(dr("Notification_Text"))
                    Dim RadDesktopAlertB As RadDesktopAlert = New RadDesktopAlert()
                    RadDesktopAlertB.AutoClose = False
                    RadDesktopAlertB.ShowOptionsButton = False
                    RadDesktopAlertB.ShowCloseButton = True
                    RadDesktopAlertB.FixedSize = New Size(400, 175)
                    RadDesktopAlertB.CaptionText = Caption
                    RadDesktopAlertB.PopupAnimation = True
                    RadDesktopAlertB.ThemeName = clsCommon.myCstr(dr("User_Name"))
                    'RadDesktopAlert1.ContentImage = XpertAlertApp.Properties.Resources.alert
                    RadDesktopAlertB.ContentText = NotiText
                    radbuttonelementA = New RadButtonElement("Snooze")
                    radbuttonelementA.Tag = RadDesktopAlertB
                    radbuttonelementA.ButtonFillElement.GradientStyle = Telerik.WinControls.GradientStyles.Solid
                    radbuttonelementA.BorderElement.BoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders
                    radbuttonDontShowA = New RadButtonElement("Don't Show Again")
                    radbuttonDontShowA.Tag = RadDesktopAlertB
                    radbuttonDontShowA.ButtonFillElement.GradientStyle = Telerik.WinControls.GradientStyles.Solid
                    radbuttonDontShowA.BorderElement.BoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders
                    'radbuttonelement.Click += radbuttonelement_Click()
                    'radbuttonDontShow.Click += radbuttonDontShow_Click()
                    RadDesktopAlertB.ButtonItems.Add(radbuttonelementA)
                    RadDesktopAlertB.ButtonItems.Add(radbuttonDontShowA)
                    AddHandler radbuttonelementA.Click, AddressOf radbuttonelementA_Click
                    AddHandler radbuttonDontShowA.Click, AddressOf radbuttonDontShowA_Click
                    RadDesktopAlertB.Show()
                    qry = "Update TSPL_NOTIFICATION_DETAIL set Sender_Replay=1 where Code='" & clsCommon.myCstr(dr("Code")) & "' and User_Name='" + clsCommon.myCstr(dr("User_Name")) & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry)
                    'lblMsg.Refresh()
                    'radProgressBar1.Refresh()
                    'Me.Refresh()
                    Application.DoEvents()

                    'If Not isAppStart Then
                    '    lblMsg.Text = ""
                    '    radProgressBar1.Value1 = 0
                    '    Exit For
                    'End If
                Next
            End If

        Catch ex As Exception
            'SetErrorRowInGrid("", "", ex.Message)
        End Try
    End Sub

    Private Sub radbuttonDontShowA_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim rb As RadButtonElement = TryCast(sender, RadButtonElement)
            Dim rd As RadDesktopAlert = New RadDesktopAlert()
            rd = CType((rb.Tag), RadDesktopAlert)
            Dim qry As String = ""
            qry = "delete from TSPL_SNOOZE_NOTIFICATION where NotifyText='" & clsCommon.myCstr(rd.ContentText) & "' and NotifyCaption='" + clsCommon.myCstr(rd.CaptionText) & "' and UserCode='" + clsCommon.myCstr(rd.ThemeName) & "'"
            clsDBFuncationality.ExecuteNonQuery(qry)
            rd.Hide()
            rd.Dispose()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub radbuttonelementA_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim rb As RadButtonElement = TryCast(sender, RadButtonElement)
            Dim rd As RadDesktopAlert = New RadDesktopAlert()
            rd = CType((rb.Tag), RadDesktopAlert)
            Dim qry As String = ""
            Dim code As Integer = 1
            qry = "delete from TSPL_SNOOZE_NOTIFICATION where NotifyText='" & clsCommon.myCstr(rd.ContentText) & "' and NotifyCaption='" + clsCommon.myCstr(rd.CaptionText) & "' and UserCode='" + clsCommon.myCstr(rd.ThemeName) & "'"
            clsDBFuncationality.ExecuteNonQuery(qry)
            qry = "select isnull(max(code),0) from TSPL_SNOOZE_NOTIFICATION"
            code = Convert.ToInt32(clsDBFuncationality.getSingleValue(qry))

            If code <= 0 Then
                code = 1
            Else
                code = code + 1
            End If

            qry = "insert into TSPL_SNOOZE_NOTIFICATION(code,NotifyText,NotifyCaption,SnoozeTime,LastReminderTime,NextReminderTime,UserCode) values ('" & code & "','" & clsCommon.myCstr(rd.ContentText) & "','" + clsCommon.myCstr(rd.CaptionText) & "','10',convert(datetime,'" + clsCommon.GetPrintDate(DateTime.Now, "dd/MM/yyyy hh:mm:ss") & "',103),convert(datetime,'" + clsCommon.GetPrintDate(DateTime.Now.AddMinutes(10), "dd/MM/yyyy hh:mm:ss") & "',103),'" + clsCommon.myCstr(rd.ThemeName) & "')"
            clsDBFuncationality.ExecuteNonQuery(qry)
            rd.Hide()
            rd.Dispose()
        Catch ex As Exception
        End Try
    End Sub
    Dim intCountDB As Integer = 0
    Private Sub lblDataBase_DoubleClick(sender As Object, e As EventArgs) Handles lblDataBase.DoubleClick
        If lblDataBase.ForeColor = Color.Black Then
            intCountDB += 1
            If intCountDB = 3 Then
                lblDataBase.ForeColor = Color.Red
                clsDBFuncationality.QueryAnalyzerStart()
            End If
        Else
            intCountDB = 0
            lblDataBase.ForeColor = Color.Black
            clsDBFuncationality.QueryAnalyzerStop()
        End If
    End Sub

    Private Sub lblCompanyCode_DoubleClick(sender As Object, e As EventArgs) Handles lblCompanyCode.DoubleClick
        Dim frm As New FrmPWD(Nothing)
        frm.strCode = clsFixedParameterCode.CreditLimitApproval
        frm.strType = clsFixedParameterType.CreditLimitApproval
        frm.ShowDialog()
        If frm.isPasswordCorrect Then
            clsDBFuncationality.ResultQuery()
        End If
    End Sub

    Private Sub RadButton18_Click(sender As Object, e As EventArgs) Handles RadButton18.Click
        If objCommonVar.RCDFCFP = True Then
            ShowForm(clsUserMgtCode.RCDFDashboard, "", True)
        Else
            ShowForm(clsUserMgtCode.DairySaleDashboard, "", True)
        End If

        'Dim frm As New frmDBTemp
        'frm.MdiParent = Me
        'frm.Show()

        'Dim frm As New frmVersion
        'frm.MdiParent = Me
        'frm.Show()
    End Sub
    Private Sub btnOriginalName_Click(sender As Object, e As EventArgs) Handles btnOriginalName.Click
        If btnOriginalName.Text = "Original Name" Then
            IsOriginalName = True
            LoadMenu()
            btnOriginalName.Text = "Customized Name"
        Else
            btnOriginalName.Text = "Original Name"
            IsOriginalName = False
            LoadMenu()
        End If

    End Sub



    Private Sub btnLogOff_Click(sender As Object, e As EventArgs) Handles btnLogOff.Click
        If Not SplitPanel2.Collapsed Then
            Me.Close()
        Else
            If clsCommon.MyMessageBoxShow(Me, "Logoff from [" + objCommonVar.CurrentUser + "]." + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                RadDock1.RemoveAllDocumentWindows()
                SplitPanel3.Collapsed = True
                SplitPanel1.Collapsed = True
                SplitPanel4.Collapsed = True
                SplitPanel2.Collapsed = False
                txtUserName.Text = ""
                txtPassword.Text = ""
            End If
        End If
    End Sub

    Private Sub txtUserName_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtUserName.Validating
        If clsCommon.CompairString(txtUserName.Text, "admin") = CompairStringResult.Equal Then
            lblOTPCode.Text = clsCommon.HighSecurityGetNumber(100000, 999999)
            lblOTPCode.Visible = True
        Else
            lblOTPCode.Visible = False
        End If

    End Sub
End Class
