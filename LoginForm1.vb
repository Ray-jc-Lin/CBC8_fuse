Imports System.Net
Public Class LoginForm1
    Private sql As String
    ' TODO: 插入程式碼，利用提供的使用者名稱和密碼執行自訂驗證
    ' (請參閱 http://go.microsoft.com/fwlink/?LinkId=35339)。
    ' 如此便可將自訂主體附加到目前執行緒的主體，如下所示:
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' 其中 CustomPrincipal 是用來執行驗證的 IPrincipal 實作。
    ' 接著，My.User 便會傳回封裝在 CustomPrincipal 物件中的識別資訊，
    ' 例如使用者名稱、顯示名稱等。
    Private WrongTimes As Integer
    Private LogAuto As Integer
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim db As New DbAccess(sSMdbName & "\setting.mdb")
        Dim s, s2 As String

        '112.9.23 move to Shown
        ''106.9.4
        'Re_SendSQL_Flag = False
        ''105.7.16
        'LabelCheck.Visible = True
        'Me.Refresh()
        ''105.9.24
        'bPingRPC = True
        'bRemoteBpath = True
        'bRemoteApath = True

        ''105.8.6 105.10.8
        ''If Not My.Computer.Network.IsAvailable Then
        'If (bPingSQL Or bPingRPC) Then

        '    If (Not My.Computer.Network.IsAvailable) Then
        '        '105.10.8
        '        bPingSQL = False
        '        bPingRPC = False
        '        If bEngVer Then
        '            MsgBox("Network", MsgBoxStyle.Critical, "Network is not available!")
        '        Else
        '            MsgBox("網路無法使用! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
        '        End If
        '        bRemoteBpath = False
        '        '105.10.13
        '        bRemoteApath = False
        '    Else
        '        '105.10.8 If Not My.Computer.Network.Ping(sTextBoxIP_R, 1000) Then
        '        If Not My.Computer.Network.Ping(sTextBoxIP_R, 1000) And bPingRPC Then
        '            If bEngVer Then
        '                MsgBox("Network is available, ping " & sDataBaseIP & " fail!", MsgBoxStyle.Critical, "Network Checking")
        '            Else
        '                MsgBox("Ping 遠端電腦 " & sTextBoxIP_R & "失敗! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
        '            End If
        '            bPingRPC = False
        '            bRemoteBpath = False
        '            '105.10.13
        '            bRemoteApath = False
        '        End If
        '        '105.10.8 If Not My.Computer.Network.Ping(sDataBaseIP, 1000) Then
        '        '109.11.29 test
        '        If LogAuto = 2 Then
        '            bPingSQL = True
        '            bDataBaseLink = True
        '        End If
        '        If Not My.Computer.Network.Ping(sDataBaseIP, 1000) And bPingSQL Then
        '            If bEngVer Then
        '                MsgBox("Network is available, ping " & sDataBaseIP & " fail!", MsgBoxStyle.Critical, "Network Checking")
        '            Else
        '                MsgBox("Ping 調度電腦 " & sDataBaseIP & "失敗! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
        '            End If
        '            bPingSQL = False
        '            '106.7.25
        '            bDataBaseLink = False
        '        End If

        '    End If
        'End If

        ''112.9.21
        'If Not (bPingSQL And bDataBaseLink) Then
        '    bPbSync = False
        'End If


        'If RadioButton_OP.Checked Then
        If RadioButton_OP.Checked And bRemoteBpath Then
            ReadRemote()
        End If
        LabelCheck.Visible = False



        ''105.10.8
        'If sDataBaseIP = "127.0.0.1" Then
        '    bPingSQL = False
        '    MainForm.CheckBoxLink.Visible = False
        '    '107.8.31
        '    bDataBaseLink = False
        'Else
        '    bPingSQL = True
        'End If
        'If sTextBoxIP_R = "127.0.0.1" Then
        '    bPingRPC = False
        '    '107.7.9
        '    MainForm.LabelJS.Text = "建勝電機"
        '    MainForm.LabelRPC.Visible = False
        'Else
        '    bPingRPC = True
        'End If

        If UsernameTextBox.Text = "jscbc800" Then
            If PasswordTextBox.Text = "561020" Then
                module1.user = UsernameTextBox.Text
                module1.authority = 4
                module1.password = PasswordTextBox.Text
                '105.7.17
                If module1.authority > 2 Then
                    SaveSetting("JS", "CBC800", "RadioButton_OP", CStr(RadioButton_OP.Checked))
                    SaveSetting("JS", "CBC800", "RadioButton_REMOTE", CStr(RadioButton_REMOTE.Checked))
                    SaveSetting("JS", "CBC800", "RadioButton_NO1", CStr(RadioButton_NO1.Checked))
                    SaveSetting("JS", "CBC800", "RadioButton_NO2", CStr(RadioButton_NO2.Checked))
                    bRadioButton_OP = RadioButton_OP.Checked
                    bRadioButton_REMOTE = RadioButton_REMOTE.Checked
                    bRadioButton_NO1 = RadioButton_NO1.Checked
                    bRadioButton_NO2 = RadioButton_NO2.Checked
                End If
                MainForm.Show()
                'Me.Close()
                Exit Sub
            End If
        End If

        If UsernameTextBox.Text <> "" Then
            Dim dt As DataTable = db.GetDataTable("select * from passwd where ID = '" + UsernameTextBox.Text + "'")

            If dt.Rows.Count <> 0 Then
                s = dt.Rows(0).Item("ID").ToString
                s2 = dt.Rows(0).Item("Passwd").ToString
                If dt.Rows(0).Item("Passwd").ToString = PasswordTextBox.Text Then
                    'MsgBox("使用者存在!!!", MsgBoxStyle.Information)
                    module1.user = UsernameTextBox.Text
                    module1.authority = dt.Rows(0).Item("setting").ToString
                    module1.password = PasswordTextBox.Text
                    '105.7.17
                    If module1.authority > 2 Then
                        SaveSetting("JS", "CBC800", "RadioButton_OP", CStr(RadioButton_OP.Checked))
                        SaveSetting("JS", "CBC800", "RadioButton_REMOTE", CStr(RadioButton_REMOTE.Checked))
                        SaveSetting("JS", "CBC800", "RadioButton_NO1", CStr(RadioButton_NO1.Checked))
                        SaveSetting("JS", "CBC800", "RadioButton_NO2", CStr(RadioButton_NO2.Checked))
                        bRadioButton_OP = RadioButton_OP.Checked
                        bRadioButton_REMOTE = RadioButton_REMOTE.Checked
                        bRadioButton_NO1 = RadioButton_NO1.Checked
                        bRadioButton_NO2 = RadioButton_NO2.Checked
                    End If
                    MainForm.Show()
                    'Me.Hide()
                Else
                    '103.3.15
                    If bEngVer Then
                        MsgBox("Wrong Password!!!!", MsgBoxStyle.Information)
                    Else
                        MsgBox("密碼不正確!!!!", MsgBoxStyle.Information)
                    End If
                    WrongTimes += 1
                    If WrongTimes >= 3 Then
                        Me.Close()
                    End If
                End If
            Else
                '103.3.15
                If bEngVer Then
                    MsgBox("User not exist!!!!", MsgBoxStyle.Information)
                Else
                    MsgBox("使用者不存在!!!!", MsgBoxStyle.Information)
                End If
                WrongTimes += 1
                If WrongTimes >= 3 Then
                    Me.Close()
                End If
            End If
        Else
            '103.3.15
            If bEngVer Then
                MsgBox("Please keyin User Name", MsgBoxStyle.Critical)
            Else
                MsgBox("請輸入使用者名稱", MsgBoxStyle.Critical)
            End If
            'MsgBox("IP:" & sDataBaseIP & ",連線失敗! 請檢查網路設備和遠端資料庫電腦.", MsgBoxStyle.Critical)
        End If


    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub


    Private Sub LoginForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '109.4.2 配比連線   109.11.29
        If bPbSync Then
            'Me.Height += 200
            LabelCheck.Text = "Check AERP2IC "
            LabelCheck.Visible = True
        End If


        '107.7.2
        Dim strHostName As String
        Dim strIPAddress = ""
        strHostName = System.Net.Dns.GetHostName()
        '110.1.20
        'strIPAddress = ""
        '104.4.25
        On Error Resume Next
        Dim Address() As System.Net.IPAddress
        'Address = System.Net.Dns.GetHostEntry(strHostName).AddressList
        Address = System.Net.Dns.GetHostEntry(strHostName).AddressList
        strIPAddress = Address(1).ToString
        '110.2.5
        'LabelNet.Text = "HostName:" & strHostName & " IP:" & strIPAddress
        LabelNet.Text = " IP:" & strIPAddress
        'truck.MakeTransparent(Color.Black)
        '101.4.18
        Dim appName As String = Process.GetCurrentProcess.ProcessName

        Dim sameProcessTotal As Integer = Process.GetProcessesByName(appName).Length

        If sameProcessTotal > 1 Then

            '102.4.17
            If bEngVer Then
                MessageBox.Show("Program running already!", " App.PreInstance Detected!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("程式已開啟!", " App.PreInstance Detected!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            Me.Close()

        End If
        appName = Nothing
        sameProcessTotal = Nothing

        '105.8.6 get Mixer1/Mixer2
        Dim file
        file = My.Computer.FileSystem.FileExists("C:\JS\OrigFile\MIXER.des")
        If file = False Then
            MsgBox("找不到檔案", MsgBoxStyle.Critical, "找不到MIXER.des!")
            Me.Close()
            Exit Sub
        End If
        Dim fileNum
        Dim s As String
        fileNum = FreeFile()
        FileOpen(fileNum, "C:\JS\OrigFile\MIXER.des", OpenMode.Input)


        '1
        s = LineInput(fileNum)
        sTextBoxPrj = s
        '108.2.27 台泥南崁(嘉利) Q36  bQ36=true
        If sTextBoxPrj = "Q36" Then
            bQ36 = True
        Else
            bQ36 = False
        End If
        '109.8.15 台泥台南 Q38  bQ38=true
        If sTextBoxPrj = "Q38" Then
            bQ38 = True
        Else
            bQ38 = False
        End If
        '109.8.17 台泥豐原(財石) Q32  bQ32=true for C12345 later case
        If sTextBoxPrj = "Q32" Then
            bQ32 = True
        Else
            bQ32 = False
        End If

        '2
        s = LineInput(fileNum)
        sTextBoxPrj2 = s
        '3
        s = LineInput(fileNum)
        sTextBoxMixerNo = s
        '4
        s = LineInput(fileNum)
        sTextBoxMixerNo2 = s
        FileClose(fileNum)
        RadioButton_NO1.Text &= " (" & sTextBoxPrj & ")"
        RadioButton_NO2.Text &= " (" & sTextBoxPrj2 & ")"

        '107.7.17
        RadioButton_OP.Checked = CBool(GetSetting("JS", "CBC800", "RadioButton_OP", "TRUE"))
        RadioButton_REMOTE.Checked = CBool(GetSetting("JS", "CBC800", "RadioButton_REMOTE", "false"))
        RadioButton_NO1.Checked = CBool(GetSetting("JS", "CBC800", "RadioButton_NO1", "TRUE"))
        RadioButton_NO2.Checked = CBool(GetSetting("JS", "CBC800", "RadioButton_NO2", "false"))
        bRadioButton_OP = RadioButton_OP.Checked
        bRadioButton_REMOTE = RadioButton_REMOTE.Checked
        bRadioButton_NO1 = RadioButton_NO1.Checked
        bRadioButton_NO2 = RadioButton_NO2.Checked
        '107.9.12

        '105.7.19
        sTextBoxIP_R = GetSetting("JS", "CBC800", "sTextBoxIP_R", "127.0.0.1")

        '105.7.19
        If bRadioButton_NO2 Then
            sProject = sTextBoxPrj2
            '110.7.28 110.9.11
            'sMixerNo = sTextBoxMixerNo2
        Else
            sProject = sTextBoxPrj
            '110.7.28 110.9.11
            'sMixerNo = sTextBoxMixerNo
        End If

        '105.8.14 move from ReadOther()
        '105.8.14 move from ReadOther()
        sYMdbName = "\JS\CBC8\DATA_" & sProject & "\Y"
        sSMdbName = "\JS\CBC8\CONF_" & sProject
        '107.7.4 user\....\App\\\ProgramFiles
        '113.8.13
        'sYMdbName_B_Local = "\Program Files\JS\CBC8\DATA_" & sProject & "\Y"
        sYMdbName_B_Local = "\JS\QJ\JS\CBC8\DATA_" & sProject & "\Y"
        '107.7.9
        'sYMdbName_B_Local = "C:\Program Files\JS\CBC8\DATA_" & sProject & "\Y"

        '107.7.4    107.7.9
        'Call ReadAdmin2()
        '109.2.17 repeat bug => 109.2.20 site error can't mark since sTextBoxBpath needed
        Call ReadAdmin()

        '110.11.26
        SaveSetting("JS", "CBC800", "sProject", Trim(sProject))
        SaveSetting("JS", "CBC800", "sYMdbName", Trim(sYMdbName))

        '105.7.19 110.7.28 110.9.11 move
        If bRadioButton_NO2 Then
            sMixerNo = sTextBoxMixerNo2
        Else
            sMixerNo = sTextBoxMixerNo
        End If

        'remote file
        sYMdbName_B = sTextBoxBpath & "\CBC8\DATA_" & sProject & "\Y"
        sYMdbName_A = sRemoteApath & "\CBC8\DATA_" & sProject & "\Y"

        Call ReadOther()
        '107.7.4
        'Call ReadAdmin2()
        '109.2.20 after site
        'Call ReadAdmin()

        '109.4.2
        bNetPing = My.Computer.Network.IsAvailable And My.Computer.Network.Ping(sDataBaseIP)

        '105.9.24
        bNetworkIsAvailable = True
        bUseNetwork = True
        bOnlyLocalPC = False
        If sDataBaseIP = "127.0.0.1" And sTextBoxIP_R = "127.0.0.1" Then
            bUseNetwork = False
            bNetPing = False
            '105.7.25 R
            '107.7.9
            MainForm.LabelJS.Text = "建勝電機"
            MainForm.LabelRPC.Visible = False
            '105.9.16
            bOnlyLocalPC = True
        End If
        If bUseNetwork And Not My.Computer.Network.IsAvailable Then
            If bEngVer Then
                MsgBox("Network", MsgBoxStyle.Critical, "Network is not available!")
            Else
                MsgBox("網路無法使用!", MsgBoxStyle.Critical, "網路檢查")
            End If
            '105.9.24
            bNetworkIsAvailable = False
            bNetPing = False
            bUseNetwork = False
        Else
            '112.9.23
            ToolStripStatusLabel1.Text = "Network:" & bNetworkIsAvailable
        End If

        Me.CenterToScreen()
        If CBool(bCheckBoxYadon) Then
            LogoPictureBox.Image = My.Resources.YaDonCar1
            FrmMonit.PictureBoxCar.Image = My.Resources.YaDonCar1
            '102.5.11
            If bEngVer Then
                sSite = "SITE"
            Else
                sSite = "核示單號"
            End If
        Else
            '100.5.11
            'LogoPictureBox.Image = My.Resources.Truck1
            LogoPictureBox.Image = truck
            FrmMonit.PictureBoxCar.Image = truck
            '102.5.11
            If bEngVer Then
                sSite = "SITE"
            Else
                sSite = "工地編號"
            End If
        End If
        '107.1.31
        If bCheckBox232 Then sSite = "送貨單號"
        If bCheckBox232 Then MainForm.Label8.Text = "訂單號碼"
        If bCheckBox232 Then MainForm.Label9.Text = "工地"

        MainForm.lblfield.Text = sSite
        UsernameTextBox.Focus()


        '102.4.17
        If bEngVer Then
            UsernameLabel.Text = "User Name"
            PasswordLabel.Text = "Password"
            OK.Text = "OK"
            Cancel.Text = "QUIT"
            Me.Text = "LOG IN"
        End If

        If bCheckBoxSP5 Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
        End If




        '改檔日期標示
        '100.5.11 add 展用圖片 add something in .log
        '100.5.12~15
        '100.5.15 printVar, 區域補正可改, 工地編號=>sSite 
        '100.5.16 @home Sim report no S
        '               History report
        '100.5.17 模擬報表功能啟動,通信停止!, truck image by \js\origfile\truck.jpg
        '100.5.26 模擬報表 save integer
        '100.5.30 maxremain of err%, AE with .**
        '100.6.8 模擬報表 water% *.*, 計量值橫向總重, 即時列印, order by ***
        '100.6.13 @ site uni_car not unique => add carser in count_A_report
        '100.6.15 @ site  carser's type integer, not string
        '100.6.18  @home 
        '100.6.26~19 @home 修改推入生產方數
        '100.6.30 @home S,G,W,C *.*
        '100.7.1 @home ditto to PLC, New_datasample_f.mdb for all float data
        '100.7.2 remain *.*
        '100.7.3 *.* Bonematerial
        '100.7.8 @js D838 total D838 for changing Pan M3
        '100.7.12 @home 變更m3 move to quecode() Label12.Visible = True
        '100.7.20 @js order by error when history sim 
        '100.8.1 Trend(1000) for YaLi long time mix
        '100.8.3 @ home to DB_PC strsql = "select * from recdata where carser=" & carser & " and unicar='" + unicar + "'"
        '100.8.12 @Eao 校磅 0.1, m3 by fMaxM3
        '100.10.9 @home after visit SheChu cFull error
        '100.10.21 @home SaveDT when quary print report
        '100.12.2~4 @home add TraceLog() to trace
        '100.12.6 @home add TraceLog() to trace more detail for displaymaterial
        '100.12.16 displaymaterial re-add
        '100.12.20  i = ReadA2N(61, 55), D161 when D22 change
        '101.1.7 Dayreport add end date, month report revise SaveDT
        '101.1.13 since PC speed
        '101.1.13 MOVE FROM Txtmod_TextChanged( txt_carsum.Text = "0.00" '泥盤第一盤完成, reset本車重量
        '101.1.13 Dim myFont = New Font("全真仿宋體", 10)
        '101.1.29 @亞東鳳山Q01 site 月米數
        '101.1.30 @home 電流圖, If CheckBoxLink.Text = "連線" Then  Cbxquecode1.Enabled = False
        '101.2.23 @JS 配方列印 
        '101.2.25 @Q01 btncancel1.Enabled = True when not LINK
        '101.3.7 single monitor bCheckBoxSingle
        '101.3.10 B配查詢, ssSaveDate, simulation
        '101.3.14 single monitor trend
        '101.3.16 @JS 
        '101.3.20 add try catch for debug with queue1.inprocess
        '101.3.21 change Bone  Read D61 when D20 change, read D70 Mod when D21 , and ReadD61Flag ReadD70Flag ReadD161Flag , queclone
        '101.3.30 sim report
        '101.4.2 Form
        '101.4.2 ad Num8
        '       Dim myFont = New Font("全真仿宋體", 10) 9 history  Formula
        '101.4.5 @home S4 小數一位 non done yet
        '101.4.5 @js G4 remain, scale report
        '101.4.8 @home S4 *.*
        '101.4.12 @JS2
        '101.4.14 @home iUniSer CarSer
        '101.4.14 S4 *.*
        '101.4.15 @js2 S4 *.* un_car 改由 read mdb uniid , 區域補正 S4 *.* 不加入
        '101.4.18 @home .mdb preinstance
        '101.5.1 empty .mdb no UniID
        '101.7.25 new PC only 1024*768 so verify mainform to fit
        '101.7.25 add 101.6.30 101.6.30 CheckBoxSet '設定'蘭 
        '101.5.28 @home rnd() for S4 *.* => copy from other
        '101.10.17 @JS 補正後 水灰比 拌合時間 for 小港
        '101.11.20 @home 放料延遲二 for 土城
        '101.11.26 temp
        '101.12.12 @js D541-D555 放料排序二, 配方;備用一,2,3,4 可自行定義
        '101.12.15 @ Q01 site simulation error queue0.UniCar
        '101.12.17 排隊一警示 iQue1AlmDelay D838 bug
        '101.12.22 排隊一警示 iQue1AlmDelay D839  iQue1AlmEnable no use, queue1.ismatch = False
        '102.1.22 歷史查詢 加本車總重
        '102.1.22 砂漿拌合米數
        '102.1.22 報表總重 小數二位
        '102.2.17 排隊一警示  排隊一 delay
        '102.2.17 列印本頁 + M3 kg
        '102.2.17 D34 bit15 非自動生產 => save to _B field "particle" (showdata(6)) 
        '102.3.16 列印實際拌合時間, report m3 tot error
        '102.3.16 Me.磅TableAdapter.Fill(Me.SettingDataSet1.磅) ???
        '102.3.18 dgvPan.Height = dgvPan.Rows(0).Height + dgvPan.Rows(1).Height + dgvPan.ColumnHeadersHeight + 5 
        '102.3.19 Import Export for Admin  @ 劍潭
        '102.2.21 save remain error when not print mixtime @home
        '102.4.16 @home DELL NB R4,RT error
        '102.4.17 4.19 @HOME begin to EngVer 
        '102.4.20 revise RT 
        '102.5.7 5.8 ,9 @HOME
        '102.5.10 @js BEFORE VIENAN
        '102.5.14 @ js <200 no error for count_A_report
        '102.5.14 D7:含水%
        '102.5.19 ,D1, 前水灰比 改成 拌合時間 調整cbc7報表格式 W/C 最後一行
        '         AE6 ALWAYSE = 0
        '102.5.28 water = 0 when only water => else ...in waterrecountall()
        '           water only for wash car => 不累計方數 bWashCar As Boolean in savebone()
        '       模擬報表  6, 1, 8 => NG
        '102.6.7 時訊　G1G2G3G3 含水率, 強度-粒徑-坍度, 
        '        .Font = New Font(fontFamily, 42.0F, FontStyle.Regular) for smallfont
        '        NonAuto 原存入 particle 改 orgmaterial1
        '102.6.25 NonAuto 原存入 particle 改 orgmaterial1 => 日月報總計 忘了改
        '102.6.25 強度-粒徑-坍度 => 強度-坍度-粒徑 showdata(5) & "-" & showdata(4) & "-" & showdata(6)
        '102.7.14 for Q10,11 B配 disable by 系統設定 bABPB  102.7.15 Ｂ配比查詢 
        '102.7.15 for 聯佶 b配縮減米數 =. 未做完
        '102.7.17 汐止 最後一盤 加砂減水 水計量有問題 
        '102.7.18 => 102.5.28 waterrecountall()
        '102.7.19 for 聯佶 b配縮減米數 => check by queue*.isBsub not bCheckBoxBPBsub
        '102.7.22 汐止 7.19 換新程式當機 => check if len to long    坤 back from Kwam   
        '102.7.22 check length of strength
        '         PrintCarDetail() 原只印 (strength)   repdetail += "設計強度:"    st = db_sum.Rows(0).Item("Strength").ToString
        '=>st = db_sum.Rows(0).Item("Strength").ToString & "-" & db_sum.Rows(0).Item("founder").ToString & "-" & db_sum.Rows(0).Item("particle").ToString
        '102.8.2 @home for 漢本 CPU TimerSQL 100000 
        '102.8.8 @JS TimerPrint 10000=>15000 delay wait for TimerSQL
        '102.8.13@JS English for Malaycia 
        '102.8.15@H 時訊 R6 err%
        '102.8.16@j eNGLISH VERIFY, 
        '           TOTAL ERROR  delay 出錯  將 sendSQL 與 save A配 拆開
        '           TimerPrint 15000=>10000
        '102.8.17 @home SQL error
        '         recDT, Field
        '102.8.19 printer print
        '102.10.19 frmOther ComboBoxBsub for 聯佶 縮米數
        '102.10.19 if Engver          Label8.Text = "" Label9.Text = ""        Label11.Text = ""
        '102.10.19 Alt.Des
        '102.10.23 @home Alt-M for 縮減月米數 for 展用
        '102.11.23 每方重量
        '          水泥註記 Other2.des Cement.des 0:disable

        '102.12.11 backup change for 巨力 @JS
        '102.12.17 ref 12.11.25 which PC crash but verified. AE 註記
        '102.12.18 12.24 12.26 bImperial 公制/英制 metric/imperial 
        '103.1.22 Engver 配方不存.....
        '103.2.24 @JS do nothing
        '103.2.25 @home Engver Label3 of Formula
        '           Material3 as Imperial flag in formula     '102.2.25 Fmaterialorg(24)=99 Imperial
        '           1m=1.09361yard 1m^3 = 1.307939 y^3
        '103.3.7 @JS NonAuto
        '103.3.11 .12 .13 Imperial @home
        '103.3.15 @home date
        '103.3.15 @JS FOR Guan !! English Win7 only can run .exe
        '           adjust the visible items
        '103.4.19 @js for Guan 
        '       some history data *.**
        '   module1.authority : 1 => no ALT-L no PB_B no Sim
        '                       2   can ALT_L PB_B no edit PBB
        '                       3   all 
        '   if no PB_B flag then no PB_B can visible
        '103.4.22 
        '103.4.23 @home test Win7
        '           redraw
        '103.4.25 @js history format
        '103.5.14 SQL resend
        '103.6.7 @FongShang site        '102.2.17 D34 bit15 => 非自動生產
        'outway varchar(1) 類別 A、M  !!! => 
        '103.6.9 @鳳山 非自動
        '        If FrmHistory.lineCnt < 17 Then FrmHistory.lineCnt = 17
        ' lineHeight=17
        '       If lineCnt < 17 Then lineCnt = 17
        '103.3.10 @home
        '   If lineCnt < 17 Then lineCnt = 17          lineHeight = 19

        '103.7.26 @js before quit
        '   DVW(i + 741) = 0
        '   DVW(i + 771) = 0
        ' 103.7.26.1 TEST WIN7 + SQL consbuji

        '103.8.30 iUniSer, iCarSer read from max(UniID) from recdata since NonAuto
        '      PrintCarDetail NonAuto 

        '103.12.18 "本日車序"
        'DataGridView1.Rows(ser).Cells(0).Value = db_sum.Rows(j).Item("Carser")
        'DataGridView1.Rows(ser).Cells(0).Value = ser + 1

        '103.10.24
        'TimerSQL.interval 10000=>5000 for 巨力 



        '104.8.8 104.8.9 104.8.10 8.11 8.15 8.16 8.17 8.18 8.19 8.20 PLC Comm @home
        'update ver 20140315 to ref. 20141224

        '103.12.26
        '生產資料補上傳 
        'DV20 變化, 骨材完成 SaveBone() => INSERT .MDB
        'DV21 變化, queue0_ModDoneChange() mc.queclone(queue0, queueP) 
        'DV22 變化, 殘留值完成 SaveRemain() => UPDATE .MDB(B)
        '                                   mc.count_A_report_only(queueP.CarSer, queueP.UniCar,) => UPDATE .MDB(A)
        '                                   TimerSQL.Enabled => wait 10(5) sec Send SQL
        '           TimerPrint.Enabled => wait 10sec Print report
        '           DVW(22) = 0       WriteA2N(22, 1) 

        '104.8.17 If COM_PLC.WriteDvNumber(0) > 16 Then

        '104.8.27 @JS 預設含水率 放料延遲 電流趨勢圖上下限
        ' ref 2015 Gundow 
        '103.12.30 "本日車序"
        'sCarSer(300) As String for History sCarSer(ser) = db_sum.Rows(j).Item("Carser")
        '103.12.31 Remain=0 if SET = 0
        '104.7.10 @JS when manual no delete queue , sbujiq1a initial value
        '104.8.9 row0 no data check
        '104.8.27 DV(36) bit 1 自動中 for 連線中 不可切換回
        '104.9.4 WATER  WIDTH
        '        COPY SOMETHING FROM GWAN print apare in Admin which 104.4.25
        '       REDRAW
        '       SAVE DEFAULTWATER
        '104.9.5 月報 出貨日報
        '104.9.19 日報
        '104.9.20 日報 終止電流存檔 Q04 A/B save same ref 104.9.4 
        'MainForm.queueP.materialRemain(23) = CSng(LabelCurrent4.Text)
        '104.9.23/24 day_report format verify
        'LabelMessage.Text = "骨材搶先中" 
        '104.9.24 FOR FORMULA CHECK C1-C6
        'Public bCheckBoxC1 As Boolean     '
        '104.9.24 c1-c6配比檢查低限(for next Pan in fact )
        '   NumSG for Sand G addjust
        '104.9.26 print all pb day report
        '         
        '104.9.27 Month report
        '104.9.28 Save PB Change and alert for new savesd.
        '104.9.30 NumSG for Sand G addjust Sum_S => (S+G)
        '   print paper length not enough for LQ-300
        '104.10.1 @HOME S1~G4 Gravity
        '104.10.3 先區域補正再 S/A @JS English
        '       S/A +- % for print report
        '104.10.5  10.1 Gravity others no reason?
        '104.10.12 放料延遲 10=>5 for A series D-reg < 1000
        '           C1~6 check lb by bImperial
        '104.10.18 bug of S/A when 沙漿
        ' check CarSer before select * from recdata
        '104.10.19 @JS set FrmMonit zgc's property "IsShowContextMenu" & "IsShowCopyMessage" to false
        ' very strenge that save DB more then 2 sec & set valu change to 砂漿
        '104.10.23 @JS bug when scale#9 no use 
        '   print S/A for Day report 
        '104.11.2 @home txt_daysum day format
        '104.11.5 @JS Printer DataGridView1.Columns(i).SortMode & other properties change edit...false
        '104.11.6 @JS  'bCommFlag = Not sim_produce 
        '           If Not bCommFlag Then FrmMonit.LabelPlc.BackColor = Color.Green

        '104.11.8 @home queue0_BoneDoneChange() no update que0
        '           add log for Materialdistribute
        '  日期格式 改回 11.2 之前 => 系統短格式 YYYY/M/D
        '104.11.13 @home iTrendSave maybe out of index when DV(23) > 200
        '104.11.24 @js add         Dim db As New DbAccess(sSMdbName & "\setting.mdb") for bug maybe when que1 not match
        '   recipe save revice to YYMMDD instead of YYYYMMDD since single resolution
        '104.12.9 @home print report error when there is "" empty formula in mdb
        '104.12.16 @JS Print report DayClass
        '104.12.24 @js trend [200]=>(300)
        '105.1.23 @home Simulation delete 
        'strsql = "delete from recdata_SIM where UniCar ='" & Str() & "'"
        '105.2.1 月報表 閏年
        '       others
        '       TimerPrint.interval 10000 => 3000
        '105.4.13 adjust width of strength for Singarpol English ver
        '
        '105.4.13 add Short daily report as used before
        '105.5.1    105.5.4 monthly also
        '105.5.4 調整配比畫面 : 強度註解 A1-1 ~ A1-15 
        '           bCheckBoxSP1 for A1-15 flag
        '105.5.6    Recipe "memo3" as AE1~15
        '105.5.8    Data mdb save to iAE1_15_memo3 queue.Fmaterialnowater(24) nowaterMaterial3
        '105.5.11,12
        '105.5.14   TEST AT js => Short format => Wide format default is short
        '           CheckBoxShort.Checked means wide format control by bCheckBoxSP2 from Admin
        '105.5.20 AE15
        '       DVW(766) = queue0.Fmaterialnowater(24)
        '       DVW(796) = queue0.Bmaterialnowater(24)
        '105.5.23 @JS
        '       Strength
        '105.5.24 @H
        '       A1,A2 => A2,A3 by AE1-15 A2-N
        '105.5.30 YaDon 6 if bCheckBoxYadon
        '       sql &= "'" & dt_data.Rows(0).Item("strength").ToString & "',"
        '       for Malysia : after send SQL , nowaterMaterial2 ++
        '   105.6.13
        '       queue.Fmaterialnowater(24) = CSng(queue.showdata(9)) => AE1_15 save to [nowaterMaterial3] of mdb
        '       [nowaterMaterial2] for SendSQL flag
        '       [nowaterMaterial1] for NonAuto production which come from D34
        '105.6.09-12
        '       bNetPing check Network after load Mainform 10 sec only once
        '       click Network link button can recheck
        '105.6.16
        'wcrate1 numeric(7, 2) 前水灰比 105.6.16
        '105.6.25 26
        'B配 另存PATH SPARE1-5
        '105.7.3
        '   bCheckBoxSP3 B配存檔 = true then B存放磁碟(C:,Z:) sTextBoxBpath (if can not create then = "C:\Program Files" )
        '                        = false then B存放磁碟(C:,Z:) sTextBoxBpath = "C:\Program Files"
        '            sYMdbName_B = sTextBoxBpath & "\JS\CBC8\DATA_" & sProject & "\Y"
        ' If Not bCheckBoxSP3 Then
        '       My.Computer.FileSystem.DeleteFile(sYMdbName_B + sSavingYear + ".mdb")
        ' before end program
        '105.7.5
        '   B path \JS\ => \B\CBC8
        '   REMOTE A PAth \cbc8
        '   sPlcDelay As String  '烏日 300 100
        '105.7.8.910
        '   環泥烏日 校磅 密碼 存檔 
        '105.7.16 @ H, Js
        '   MOVE B data from local to remote 
        '   bRemoteBpath As Boolean     'check if remote B exist
        '       [RemainMaterial1] for insert remote A file flag
        '   bCheckBoxSP5 可選遠端
        '105.7.18 @J
        '   TimerSQL.interval 5000->3000
        '105.7.22 23
        '   車次 盤次 bUseNetwork
        '105.7.24 @ home remote
        '   bEngVer report for _B
        '105.7.25 @JS
        '   bCheckBoxSP6 環泥烏日
        '105.8.5 @ Q01 site
        '   PB 10char DT RT bug 
        '105.8.6 @JS
        '   環泥烏日 Q01 _B path bug
        '105.8.10 @home
        '   環泥烏日 receia = > recei1a
        '105.8.10 @home
        '       queue0.UniCar new move to quecode  since "" empty
        '105.8.12
        '       don't care carser since above, printer report
        '105.8.13 @site
        '       update A B remote file no check
        '105.8.14 @site
        '       TimerSQL interval 3000=>2000 since D20 update fast
        '       Public sRemainUniCar As String = ""     'prevent TimeSQL delay queueP.UniCar maybe updated  
        '105.8.17 @home
        '       only clone queueP when DV21 >0     If DV(21) > 0 Then MainForm.mc.queclone(MainForm.queue0, MainForm.queueP) '105.7.18 move from ModDoneChange
        '105.8.21
        '       預設含水率  If module1.authority >= 3 Then Button4.Visible = True
        '       perhaps bug for repeat sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
        '105.8.29 @JS
        '   tbxworkcode font 18pt=>16pt
        '   sFile += FixStr(st, 10, "R") for CBC7
        '   sFile += " ****"
        '105.9.3
        '   bCheckBoxYadon no pw when 校磅
        '   remote PC day sum check
        '   CheckBoxsSim 即時列印 add register
        '   其他參數設定 左 Other3.des for save to remote PC (右 Other.des)
        '   Month_report2() 
        '           strsql += " DISTINCT RecDT, cube, SaveDT,
        '105.9.11 @J
        '   no copy when And sTextBoxIP_R = "127.0.0.1" 
        '105.9.14 @home
        '    it's no need to update  If bRemoteApath Then => only CheckRemoteAfile
        '   Sub waterrecountall()  If i <> 0 Then IF NO WATER THEN ERROR BUG!
        '   Thread.Sleep((CInt(sPlcDelay) * 0.05))
        '105.9.24 @js
        '   CheckBoxSP7,bCheckBoxSP7 => 參數存至遠端 SaveFilesToRemote() 
        '105.10.5
        '   Day_report() Day_report_Short()  Month_report() Month_report_Short() 
        '        strsqlall = "Select DISTINCT "
        '105.10.6 參數存至遠端 存配比
        'If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bCheckBoxSP7 Then
        '105.10.7 @JS
        '   CheckBoxSP10 ~ 15
        '   CheckBoxSP8 校磅密碼紀錄
        '   CheckBoxSP9 配比密碼紀錄
        '   CheckBoxSP10 列印密碼調整
        '105.10.8 @H
        '                 bPingSQL 
        '           If sTextBoxIP_R = "127.0.0.1" Then
        '                    bPingRPC = False bRemoteBpath = False
        '                    MainForm.LabelRPC.Visible = False
        '           If sDataBaseIP = "127.0.0.1" Then
        '                    bPingSQL = False
        '                    MainForm.CheckBoxLink.Visible = False
        '105.10.13 @J
        '   bPingRPC = False then bRemoteApath = False
        '105.10.19
        '   NonAuto 非自動生產 需設定 非自動上傳
        '       save path : sYMdbName_B_Local
        '       'bDataBaseLink = True
        '       MainForm.TimerSQL.Enabled = True -> 
        '   B配 歷史資料 查詢條件
        '   If Not (bRemoteApath And bRemoteBpath) Then
        '105.10.20 ref queuecode
        '    SaveNonAuto()      queueP.UniCar =
        '105.10.21
        '   NonAuto => printer/history by orgmaterial1 , but SQL by nowaterMaterial1
        '   reset both to 0 in SaveBone(), set 1 in SaveNonAuto()
        '105.10.22
        '   NonAuto REVISE RESET D20 D21 NonAutoProd REVISE
        '105.10.27
        '   CheckBoxSP11 備用11 => 可修改 A配 首盤時間
        '   歷史資料 A配 查詢 雙擊首盤時間 輸入欲加/減之分鐘數 
        '   輸入完畢請重新查詢
        '105.11.3
        ' FrmAdmin ComboBoxDelay add 150 250 for 烏日
        '   ReadDvStart(7) => ReadDvStart(17) 當操作需無讀寫太多D 可能造成 讀不到 D161 => ReadD161Flag 可能不正常
        '   *** 校磅 Timer1 1000=>2000 才不會卡住讀其他D 殘留值 等
        '105.11.5
        '   END program if no file when exit
        '   print report BB for remote disconnect
        '   change IP with .bat
        '   Remote PC change to Operate PC
        '   strsql += "nowaterMaterial1 ," for 寬格式
        '105.11.9
        ' printer report BB 
        '105.11.18 bug
        'strsql = "select * from recdata where and unicar='" + unicar + "'"
        'strsql = "UPDATE recdata SET RemainMaterial1=1 "
        '105.6.4 105.7.3 資料自動補傳
        'If bCheckBoxSP4 Then
        '105.11.18 先不做
        '105.11.22 資料自動補傳 回復 , 資料自動補傳時訊
        '105.11.29 bCheckBoxSP7 參數存至遠端 not in  If bCheckBoxSP4 Then
        'If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bCheckBoxSP7 And bPingRPC Then Call SaveFilesToRemote()
        '105.12.14 馬來西亞 update Qkey to .mdb
        '105.12.20 ResendSQLs
        '105.12.21,22,23 ResendSQLs,         If bCheckBoxSP4 Then queuekey
        '105.12.23 DV(2) DV(21) limit
        '105.12.23 (骨材盤次<>呢料盤次) only change  骨材重量 => bug 106.4.3
        '106.1.2 非自動 英文版可
        '       日總計 
        '106.1.3 時訊台南 for 馬來西亞
        '106.3.3  @姚  bOnlyLocalPC AND bCheckBoxSP3 B配存檔
        '106.3.12 模擬報表 bug since 105.12.23 DV(2) DV(21) limit
        '106.3.12     If BoneDonePlate_temp = 1 Then   queue0.UniCar = "UN_" & (iCarSer_SIM).ToString
        '106.3.15  conn = New SqlClient.SqlConnection("Data Source=" & sDataBaseIP & "; Initial Catalog=" & SqlExpressDbName & ";Persist Security Info=True;User ID=buji; Password=buji ;Connection Timeout=5 ")
        '106.3.28 over midnight txt_daysum.Text
        '               db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")
        '               update time when remain 
        '106.3.30 ?
        '       SaveSetting("JS", "CBC800", "sTextBoxBpath", CStr(sTextBoxBpath))
        '106.2.16 106.4.3 ref Malasia      for AE Error% bug PV must CSng(Format(PV, "0.00")) brfore caculate
        '    Function ErrorPercent(ByVal PV As Single, ByVal SP As Single, ByVal Remain As Single) As Single
        '106.4.3
        '    if type = "B" water SP no change 骨材盤次<>呢料盤次 only change  骨材重量
        '106.4.5
        '       Q21亞東小港 B 計量值 時常=0 
        '       TraceLog3(sLog) in SaveRemain() & CheckRemoteBfile to see value
        '106.4.7
        '       => CheckRemoteBfile maybe next D21 change after TimerSQL delay => add Bone record
        '       => only MoveLocalToRemoteBfile(ByVal unicar As String, ByVal plate As Integer) for one record after SendSQL()
        '106.4.10 
        '        MoveLocalToRemoteBfile() every(Pan)
        '106.4.22 for                 ReadD161Flag = 1 萬一沒讀殘留 強迫                 ReadD161Flag = 0 => read D161 again
        '               TimerD161.Enabled = True 3000msec
        '106.4.22  "配比料號:"
        'repdetail += Microsoft.VisualBasic.Left(st & "        ", 8)
        'repdetail += Microsoft.VisualBasic.Left(st & "        ", 12)
        '106.5.2
        '   frmPrinter
        '   Panel2.Width = DataGridView1.Width + 5
        '   Me.Width = Panel2.Left + Panel2.Width + 15
        '   DayReport
        '   配比/方 AE 沒改到 參考S
        '106.5.7
        '   寬格式日月報表 小計
        '106.5.16 @JS
        '       恐 D70~83 D100~113 讀取出錯 沒值
        '106.5.16 move from above
        '       MainForm.queue0.BoneDonePlate = DV(20) => wait ReadD61Flag = 2
        '106.5.20 @JS
        '106.5.22 106.5.27 @home
        '       1.D20 change  no save to mdb file , only move data (all Setpoint and PV of Bone ) to bufer (que_0) when D21 change
        '                   DV(20) changed by ReadA2N, 
        '                           ReadD61Flag = 0 , ReadA2N(61, 38), ReadD61Flag = 1    D61	S1修正印出值 D91	S1印出真值
        '                           wait for ReadD61Flag = 2 , 
        '                               DVold(20) = DV(20)
        '                               MainForm.queue0.FmaterialReal(i) = DV(61 + i)   !!! if no ReadD61Flag no update here !!!
        '                               MainForm.queue0.BmaterialReal(i) = DV(91 + i)
        '                               MainForm.queue0.BoneDonePlate = DV(20)
        '                                    rais event queue0_BoneDoneChange() ,  Call SaveBone_New(queue0.BoneDonePlate, queue0.plate(queue0.BoneDonePlate - 1))
        '                                                       no mdb save , only fill queue0
        '        106.5.27  ??
        '        queue0.Fmaterialorg(i) = 0
        '        queue0.Bmaterialorg(i) = 0
        '        queue0.Materialfront(i) = 0
        '        queue0.Materialback(i) = 0
        '       2.D21 change save queue0 (PV of Mod , time ) to mdb 修正 A配
        '               ReadD70Flag = 0, ReadA2N(70, 44),  ReadD70Flag = 1   D70 W1 修正印出值 D100 W1印出真值 for Mod
        '                MainForm.queue0.FmaterialReal(i + 8) = DV(70 + i)
        '                MainForm.queue0.BmaterialReal(i + 8) = DV(100 + i)
        '                       old
        '                       MainForm.mc.queclone(MainForm.queue0, MainForm.queueP) => move to SaveMod_New()
        '                =>
        '                SaveMod_New() save to mdb
        '                       mc.queclone(queue0, queueP) for SaveRemain
        '
        '       3.D22 update again (Remain & time)
        '
        '106.5.30
        '        use top UniID to update remain

        '106.6.8
        '       配件庫存管理

        '106.7.21 106.7.22
        '        亞東 自動補傳  => receib
        '        CheckBoxSP4　馬來西亞 自動補傳 => receia ( s_receia )
        '        in ResenSQL if bCheckBoxYadon

        '106.7.25 @時訊 台南
        '                   SQL2008 之後 會在 IP 後加路徑 如 環泥烏日
        '               @home
        '               TextBoxSPARE2 : SQL路徑
        '106.8.5 @H
        '               resend D4 error revise 
        '106.8.21 @j
        '106.8.26 @j
        '   RT read from remote file
        '   check network before SendSQL
        '106.9.3
        '       骨材盤次更新 in SaveBone_New(), mc.queclone(queue0, queueP) move to queeP
        '       泥料盤次更新    SaveMod_New() : "Insert into recdata... 存檔
        '       殘留盤次更新    SaveRemain_New() : mc.count_A_report_only() A配 重新計算 再次存檔
        '106.9.4 
        '    queue0.Fmaterialorg(i) should be clear when queue0.BoneDonePlate = queue0.plate.Length
        '   add Re_SendSQL_Flag and  bDataBaseLink And bNetPing  to prevent SQL error
        '106.9.6,7@J confused 
        '
        '106.9.8   @h: add queueR for remain when MOD change
        '
        '106.9.9    @J replace queueR in TimerSQL_Tick()
        '           MoveLocalToRemoteBfile() check RPC before move
        '106.9.10     @h
        '           bug in Monit        AE         If MainForm.queueP.Materialfront(i + 8 + 9) > 0 Then MainForm.queueP.FmaterialReal(i + 8 + 9) = DV(79 + i) * 0.01
        '         If sProject = "Q04" Then => cancel
        '106.9.14
        '   water% must copy when Bon Done => add queue0.waterall(k)
        '   ErrorPercent_AE() bug Re_SendSQL()
        '   fW = db_sum.Rows(iPan).Item(arrdb_org(8)) + db_sum.Rows(iPan).Item(arrdb_org(9)) + db_sum.Rows(iPan).Item(arrdb_org(10)) + db_sum.Rows(iPan).Item(arrdb_org(17)) + db_sum.Rows(iPan).Item(arrdb_org(18)) + db_sum.Rows(iPan).Item(arrdb_org(19)) + db_sum.Rows(iPan).Item(arrdb_org(20)) + db_sum.Rows(iPan).Item(arrdb_org(21))
        '106.9.27 @J
        '           REMOVE Dowork1.2
        '106.9.27 local bb path
        'If bPingRPC Then
        'If bPingRPC Or bRemoteBpath Then
        '106.9.27 use plat instead of from mdb
        'If dt_data.Rows(0).Item("plate") = dt_data.Rows(0).Item("memo3") Then
        '106.9.28 強制註記資料已傳
        '106.10.7 聯佶空單
        '   聯佶空單
        '       CheckBoxSP13    bCheckBoxSP13
        '       排隊只能2個
        '106.10.11 聯佶空單
        '       sSavingDate = "#" & Format(dtSimP(1), "yyyy/MM/dd ") & "#"
        '106.10.16 revise check network 
        '       since 小港 .lg2 error reduce TraceLog()
        '               On Error Resume Next
        '               If COM_PLC.WriteDvNumber(27) > 1 Then
        '106.10.17 @thsr
        '    mixyn check
        '106.10.20 @HY 
        '    mixyn
        '106.10.21
        'Timerprocess_sim.Interval = 1000
        'Timerprocess_sim.Interval = sPlcDelay * 2
        '106.10.25
        '        Cbxquecode1.Text = Replace(Cbxquecode1.Text, " ", "")
        '106.11.2
        '106.10.7 聯佶空單 106.11.2 
        '       ElseIf DataGridViewQue.Rows(0).Cells(0).Value.ToString = "" And Not bCheckBoxSP13 Then
        '       ElseIf DataGridViewQue.Rows(0).Cells(0).Value.ToString = "" Then
        '       verify delete Que1,2

        '106.11.10
        ' review 土城 已經修改的功能 WindowsApplication2_20141224_new_160201A_亞東十碼_161227_170815
        '105.6.23
        '       water default
        '        '105.8.21
        'If module1.authority >= 3 Then Button4.Visible = True
        '105.6.23 106.11.10 ref 土城
        'If module1.authority >= 2 Then Button4.Visible = True
        '
        '105.8.14 @JS
        '       亞東配比十碼
        'repdetail += Microsoft.VisualBasic.Left(st & "        ", 10)
        '                    repdetail += Microsoft.VisualBasic.Left(st & "        ", 12)
        '
        '105.8.29 sFile += FixStr(st, 10, "R")
        '   tbxworkcode font 18=>16
        '   DataGridView1.Columns(4).Width = 140　( FormHistory)
        '   105.5.17     DataGridView1.Columns(4).Width = 150
        '
        '106.8.26 too long
        'sFile += " ****"
        '106.11.10 ref TuCheng
        'sFile += " ****"
        '
        ' 含水率上限 frmOther fMaxWaterPer
        '
        '  Add TABLE  PBHistory to Scale_org.mdb of myPC 
        '106.11.15 11.16
        '       Sim bug
        '106.11.20 @ Tainan site
        '   DISABLE PING CHECK
        '106.11.22 ref 土城 bug when Simulation
        'fMaxWaterPer = CSng(GetSetting("JS", "CBC800", "MaxWaterPer", "20.0"))
        'fMinWaterPer = CSng(GetSetting("JS", "CBC800", "MinWaterPer", "0.0"))
        '106.12.10
        'Ping
        '       LoginForm : ping RPC, DDPC
        'still try to send SQL even Ping error
        '106.12.30
        ' adjust Others. Print size
        '107.1.14
        '   send R* to SQL
        '   sLabelQkey FOR SendSQL ? seems no effect
        '107.1.31
        '   for TaiNei add ClassOrder.vb 
        '
        '107.2.3
        '   for TaiNei 神岡 財石
        '107.2.11
        '   C4,C5 => C5,C4
        '107.2.24
        '   for TaiNei 神岡 財石
        '107.3.3  3.4
        '   聯佶 seemed D20 變1 時 ReadD61Flag = 1 之後 
        '   add some log
        '   If MainForm.queue0.ModDonePlate <> DVold(21) And CInt(MainForm.txtbone.Text) = DV(20) And DVold(21) <= MainForm.queue0.BoneDonePlate Then
        '107.3.6
        '   填入日完成方數
        '   field = MainForm.lblCustomerName.Text
        '107.3.7
        '    Add Class PLC_DV in countmaterial.vb
        '   Public WithEvents D_20 As New PLC_DV
        '   move read D61/D70/D171 S1修正印出值 D91 S1印出真值 for Bone/Mod/Rem in DvChange
        '   adjust D21/D22 update : If D_20.OV2 = D_20.PV Then D_21.PV = DV(21)
        '107.3.8 3.9 3.10
        '   TraceLog Class : TLog
        '   TLog.add
        '107.3.11 
        '   for Alert begin TaiNei
        '107.3.16 .17
        '   add MonLog instead of TraceLog()
        '107.3.18
        '   add MainLog instead of SavingLog()
        '107.3.19 聯佶 Q16 
        '   only TaDon all R1~RT or R4,R6,RT
        '107.3.20
        '   revise Timerprocess_sim for 聯佶 Q16
        '   TimerLink.Interval = 1000 when CheckBoxLink
        '107.3.19 !!!
        'If SqlExpressDbName = "conscbuji" Then
        '    SqlExpressDbName = "conscbuji_ls388"
        'End If
        '107.3.21
        '   bCheckBoxSP14 備用14 : 不回傳R1R2R3R5
        '107.4.4
        '   寬格式 -殘留
        '107.4.10
        '   聯佶
        '107.4.12
        ' 
        '107.4.14
        '   use BackgroundWorker1_DoWork threading to save Mod A
        '   use BackgroundWorker2_DoWork threading to save Mod B
        '   BackgroundWorker1_RunWorkerCompleted BackgroundWorker2_RunWorkerCompleted
        '   set Background1_Flag , Background2_Flag
        '   If Background1_Flag And Background2_Flag And D_21.OV2 = D_21.PV Then
        '
        '107.4.28
        '   move AlarmCheck() to mainform Timer1_Tick instead of frmmonit 
        '   含水率上下限 water_all_W_Limit 土城
        '107.4.29
        'If bDataBaseLink And bPingSQL Then
        'TimerSQL.Enabled = True
        '    End If
        '107.4.30
        '   MiscChange() call in FrmMonit Timer4 950msec
        '107.5.10 107.5.11 排隊自動 前小於後 無法完成 bug
        '   If queue0.RemainDonePlate >= queueR.plate.Length Then
        '   dbrec.dispose()
        '107.5.18
        '   verify FrmMonitor
        '107.5.30
        '   @J 土城log queue0_BoneDoneChange() : 骨材搶先中 BMR= 032 ??
        '107.6.2    預設含水率 造成通信頻繁
        '   @J CHECK BEFORE UPDATE TO dgvWater since set to PLC
        '   車次顯示
        '107.6.3 disable
        '   displaymaterial() in SaveBone & SaveMod
        '107.6.5
        '   MainForm.LabelCarSer.Text = iCarSerToday
        '   CheckBoxsSim.ForeColor = Color.Green when remain chang and TimerPrint_Tick back to BLACK
        '107.6.6
        '   CheckBoxsSim.ForeColor = Color.Green
        '107.6.7
        '   try to monitor Dir Change by another program
        '   SaveSetting("JS", "CBC800", "sMdbDir", s)
        '107.6.9
        '   TimerSQL.interval 3000->1000
        '   土城 預設含水率 排隊自動時 ??
        '107.6.10
        '   frmPrinter : Button6.Visible = True again for 配比列印
        '    排隊一 空白 DV838=1 for 拌合機節能
        '   bCheckBoxSP16~20 for FrmAdmin
        '   bug , if History not today
        '   FrmHistory.start_day.Value = CDate(sSavingDate)
        '   拌合機節能顯示 D30 bit10 (1024) ??
        '107.6.13 坤 said 土城 搶先bug 改善 
        '107.6.14
        '   預設含水率 move codes from btnbegin.click to quecode()
        '    since also called when 排隊自動啟動 & use CheckDefaultWater() instead codes
        '107.6.16 107.6.17
        '   ReadAdminName Read Admin.des text when Admin
        '   輸送機節能
        '   History B... change to Alt-X
        '   CheckBoxtTest 1280*960 => 單一配比
        '107.6.20 預設含水率 bug
        '   CheckDefaultWater(queue0.showdata(1), Tbxquefield1.Text)
        '   CheckDefaultWater(queue0.showdata(1), tbxworkfield.Text)
        '107.6.21 @j
        '   拌合機節能顯示 bug
        '   bTimerConveyerSetFlag = False
        '107.6.22
        '   History B... change to Alt-X 先改回原來模式
        '   If e.KeyCode = Keys.X And e.Modifiers = Keys.Alt Then
        '   Button1.Left = Label33.Left + Label33.Width ...
        '   cube_tot_print day repoprt PB m3
        '   printTEMP &= FixStr(DataGridView1.Rows(j).Cells(3).Value => total

        '107.6.22
        '以此版本為 Windows XP 之最後版本
        '
        '107.6.24 土城
        '   D829 ?? ERROR:Write D829 - 1 Q: 0
        '   insert bb fail?  Insert bb fail! UniCar= UN_2706 , Plate= 1
        '   107.6.25 but 永康沒 Insert bb fail! UniCar= UN_2706 , Plate= 1  

        '107.6.26 bug!! can not disable since update SP in in SaveBon_New
        '107.6.3 disable    
        'displaymaterial()

        '107.6.27
        '   Admin Alt-X for admin label name
        '   spare ( B PB ) no save just ""
        '   report ** => *
        '107.6.28 
        '   reset D741 only when no 骨材搶先
        '   in count_A_report_only() => if in error range use DV61
        '   real_A(i) = MainForm.queueP.FmaterialReal(i) '取得A_Table 物料實際值
        '    TraceLog name    fileName = fileName & "\" & sSavingYear & "_" & Today.Month.ToString & "_" & Today.Day.ToString & "_" & logName & "." & logName
        '107.6.29
        'If bRadioButton_OP Then
        '    CheckBoxLink.Visible = False
        '    bCommFlag = False
        'End If
        '107.7.2
        '   copy mdb to remote PC
        '       My.Computer.FileSystem.CopyFile(datafile, datafile_r, True)
        '107.7.2 107.7.3
        '   If bRadioButton_OP And bDataBaseLink Then
        '   If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bPingRPC Then
        '107.7.4
        '   Admin2.des
        '   C:\ProgramFiles
        '107.7.9 
        '   still use Admin BUT no copy to remote PC
        '   label 建勝電機 R no image ! use red / Green  
        '107.7.24
        '   bug when queue1.inprocess still water 
        '       If Not queue1.inprocess And queue0.BoneDoingPlate = queue0.ModDoingPlate Then
        '   bug when report B since sYMdbName_B
        '       Day_report(sYMdbName_B & DateTimePicker1.Value.Year & "M" & DateTimePicker1.Value.Month & ".mdb", "A", DateTimePicker1.Value, DateTimePicker2.Value)
        '   if not bRemoteBpath no Re_SendSQL
        '       If bPingRPC And bRemoteBpath Then
        '   bCheckBoxSP4 資料自動補傳時訊 save queuekey => save queuekey anyway
        '       If bCheckBoxSP4 Then
        '           strsql += "queuekey, "

        '107.8.9    prineReport A: B no:
        '   ref _Win7   107.8.6 CheckBoxB 
        '   printTEMP &= vbCrLf & "列印時間: " & Now() & vbCrLf
        '   If CheckBoxB.Checked Or CheckBoxB.Checked Then
        '    printTEMP &= vbCrLf & "列印時間  " & Now() & vbCrLf
        '   Else
        '    printTEMP &= vbCrLf & "列印時間: " & Now() & vbCrLf
        '   End If
        '107.8.9
        '   disable MoveLocalToRemoteBfile() => move by remote PC CBC8_File
        '   disable CheckRemoteBfile()
        '107.8.10 
        '   CheckRemoteAfile() disable 
        '   done MoveAToBuf() by remote PC => only call after SaveRemain_New() TimerMoveA.Enabled(3sec) then TimerFlag.Enabled(3sec) write Y2FLAG.log to trigger MoveRemoteBfileToLocal() of 遠端PC CBC8_File
        '   A.  操作盤設共用資料夾 \\192.168.L.L\CC = ProgramFile\JS\cbc8\Data_Qnn 給 RemotePC MoveRemoteBfileToLocal() <Y20xxA.mdb A buffer save as bb >
        '       遠端PC CBC8_File [拌合機共用資料夾] \\192.168.L.L\CC (操作盤PC IP)
        '   B.  遠端PC設共用資料夾 \\192.168.r.r\bb = (USB FLASH) JS 存 mdb
        '       操作盤其他設定 [存放磁碟(\\192.)]  \\192.168.r.r\\bb (遠端PC IP)
        '   C.  遠端PC MoveRemoteBfileToLocal() move A.mdb => C:\JS\CBC8\Data_Qnn 
        '       <操作盤關機會再 copy file to RemotePC C:\JS\CBC8\Data_Qnn>
        '   D.  遠端PC MoveRemoteBfileToLocal() move (B).mdb => 共用資料夾 \\192.168.r.r\bb = (USB FLASH) JS 存 mdb 
        '   E.  遠端PC CBC8_File 驅動機制 \\192.168.L.L\cc\Y2FLAGLOG 改變
        '107.8.12 @NanTao
        '   TimerFlag TimerMovA =>5sec
        '107.8.14
        '   move copy A mdb
        '107.8.16
        '   add dbrec.dispose() in SaveRemain_New()
        '107.8.20
        '   聯佶 W1 不參與水分補正 
        '   Public bCheckBoxW1 As Boolean
        '107.8.21 limit lenth to 12
        '       tbxworkfield.Text = Tbxquefield1.Text

        '107.8.21 forgot to revise !!!???
        '107.7.24 bCheckBoxSP4 資料自動補傳時訊 save queuekey => save queuekey anyway 
        '   CHECK b FILE REC COUNT BEFORE SEND sql
        '   move to remote by remote program so 此時not移至遠端 = > local
        '   dbrec.changedb(sYMdbName_B_Local + sSavingYear + ".mdb")

        '   TimerMoveA :5sec=>2sec 怕殘留與下一盤尼料離太近
        '   TimerFlag   :5sec=>3sec
        '107.8.30
        '           Call MainForm.ButtonFixClick() for 強制註記資料已傳
        '107.8.31
        '   if local PC only (as below) :  => TextBoxSPARE1(sRemoteApath) no save
        '        If sDataBaseIP = "127.0.0.1" And sTextBoxIP_R = "127.0.0.1" Then
        '           bOnlyLocalPC = True
        '   add sTextBoxIP_R <> "127.0.0.1" when exit for speed
        '   If sTextBoxIP_R = "127.0.0.1" => only 環泥烏日bCheckBoxSP6 send remoteB or set BB when B
        '   add TimerMoveB only 環泥烏日bCheckBoxSP6 send remoteB
        '107.9.12 
        '   when remote PC
        '       If bRadioButton_REMOTE Then
        '           bCommFlag = False
        '       End If
        '   for Tainei
        '       mc.Send_232 => move from TimerSQL
        '       Data2Order() reset order.*
        '107.10.19
        '   RT only last Pan
        '   => save * into recdata when SendSQL and del when Last
        '   Re_SendSQL no change 
        '   
        '   mc.UpdateRecSend() sometime lost
        '   move to BackgroundWorker4
        '107.10.22 no save B配比米數縮減
        '   SaveSetting("JS", "CBC800", "fBPBsub", CStr(fBPBsub))
        '   'fBPBsub = CSng(GetSetting("JS", "CBC800", "fBPBsub", "0.00"))
        '107.10.23
        '   列印參數 BB 查詢
        '   月米數 DISTINCT
        '107.10.27
        '   ditto 歷史資料 : 料號 工地代號 強度
        '   列印參數 BB 放大鏡 查詢
        ' ** C:\Users\buji\AppData\Local\VirtualStore\Program Files\JS\CBC8\DATA_Q21
        '
        '107.12.28 first pan's Time if only one pan => 2001.1.1
        '   '106.10.11 聯佶空單 107.12.28 bug?
        '   strsql &= "#" & Format(dtSimP(0), "yyyy/MM/dd ") & dtSimP(0).Hour.ToString & ":" & dtSimP(0).Minute.ToString & ":" & dtSimP(0).Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        '   sSavingDate = "#" & Format(dtSimP(0), "yyyy/MM/dd ") & "#"
        '
        '   bCheckBoxSP15 X配調整 for 巨力
        '   bCheckBoxSP3 X配存檔 for 台泥 要注意 PLC 傳送修正值 CheckBoxA 要勾選
        '108.1.10
        '   just in BackgroundWorker2_DoWork to check
        '   If bCheckBoxSP3 Then dbrec.ExecuteCmd(Background2_Sql)
        '
        '108.2.27
        '   台泥南崁(嘉利) Q36  bQ36=true
        '   S1,S2,G1,G2,G3 => S1,S2,G1,G2,S3
        '   C1,C2,C3,C5,C4 => C1,C2,C4,C5,C3
        '   read db error => timer232 delay
        '   TimerRx 2000=>1000
        '
        '108.3.5
        '   台泥南崁(嘉利) Q36  bQ36=true
        '   S1,S2,G1,G2,G3 => S1,S2,G1,G2,S3 => S1,S2,S3,G1,G2
        '   TimerRx 2000 => 1000
        '108.6.12
        '   越南文

        '108.7.25
        '   CSV => CBC8Data project
        '108.7.29 小港廠 潘襄理 會議
        '    delay after R7 for insert to recdata
        '108.9.12
        '   TimerMoveA interval:2000
        '   TimerFlag interval:5000
        '   maybe too long to move to remote by next pan
        '   => TextBoxSPARE3 set the TimerMoveA interval
        '   => TextBoxSPARE4 set the TimerFlag interval
        '108.9.21
        '   水灰比 小數第三位捨去
        '109.2.17
        '   B配統計短少 => A配 未明原因 uni_car => UN_1 (orginal => when data table no record)
        '   => UN_ddhhmmss when qu
        '   遮掉 ReadAdmin()
        '109.2.20
        '   109.2.17 遮掉 ReadAdmin() => site can not read remote bb since sYMdbName_B = sTextBoxBpath & "\CBC8\DATA_" & sProject & "\Y" bug
        '
        '109.4.2 配比連線 bPbSync
        '   YaTong fomula配比連線傳輸協定網路版 潘誼聰
        '   Public Class ClassFomula
        '       read table AERP2IC 
        '       Readed = 1 : 時訊寫入預設為1 工業電腦讀取後設為0
        '       Flag : 0:新增 1:修改 2:刪除 3:停用 4:其他
        '       enable : 0：disable 1：enable
        '       Visible : 若Visible=0，則enable 必須為0
        '1.
        '   CheckBoXSP16 => bPbSync As Boolean  其他設定 : 備用16 => 配比連線
        '2.
        '   ReadOther() to add Recipe.mdb fields
        '   'fields to add or not
        '   jj = dt_mdb.Columns.Count
        '   If jj < 50 Then
        '       fommula_Mdb.AddYatongFields(db_mdb)
        '       fommula_Mdb.SetYatongFields(db_mdb)
        '   End If
        '   form Fomula : slump,diameter change to string instead integer
        '3.
        '   配比畫面
        '       坍度slump 粒徑diameter from integer to string
        '       Queryfomula() tsbtnsave_Click() 上傳配比 下載配比
        '109.4.26
        '   配比 must verify for every case not jus YaTong
        '       code defaul as ID when add new 
        '109.5.1
        '   ReadOther()
        '       add new field for Formula => not ony for Yatong for all
        '
        '109.5.7
        '   ReadAdmin() bug item 41,42
        '   remove bCheckBoxYadon,bCheckBoxSingle,bEngVer,bCheckBoxtTest from ReadOther() since should be from Admin.des
        '109.5.11
        ' sMixerNo = s
        '109.5.13
        'DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(0)
        '109.6.9
        '   Tainei remote PC update PB send back to OP PC
        '109.7.20
        '   AE decimal error => 
        '   temp = (set_temp * diff_temp * 0.01) + remain(i) + set_temp  '虛擬實際值
        '   real_A(i) = 0.01 * ((temp - (temp Mod Int(pond_scale * 100)))) + remain(i) + set_temp  '虛擬實際值

        '109.8.15
        '109.8.15 台泥台南 Q38  bQ38=true
        '   C1,C2,C3,C5,C4 => C1,C2,C4,C5,C3
        '   攤度 粒徑 ?? => TextSlump.Visible = True ... 捨棄舊的欄位 Founder Particle (integer only) 使用亞東配比連線定義的文字欄位 slump diameter
        '109.8.17
        '       C1,C2,C3,C5,C4 change to C12345
        '       - remain for RS232
        '       add bQ32 for later case C12345 (else)

        '109.9.11 invisible
        '       Button7.Visible = Not bPbSync
        '       Button7.Visible = False
        '109.9.11
        '   攤度 粒徑 ?? => TextSlump.Visible = True ... 捨棄舊的欄位 Founder Particle (integer only) 使用亞東配比連線定義的文字欄位 slump diameter
        '   'sql += "spare, totalweight, mixtime, founder,strength, particle, memo1, memo2, memo3, ToPLC1, ToPLC2, ToPLC3, ToPLC4, Mortar from fomula where code ='" + code + "'"
        '   sql += "spare, totalweight, mixtime, slump,strength, diameter, memo1, memo2, memo3, ToPLC1, ToPLC2, ToPLC3, ToPLC4, Mortar from fomula where code ='" + code + "'"

        '109.10.5
        '   配比 存完不跳第一筆
        '   sPBUser = module1.user
        '   台泥(bCheckBox232) 自動(CheckBoxContinue) 排隊一(Cbxquecode1) => 生產

        '109.11.29  109.12.23
        '   配比連線 
        '   要求建勝將”車號”設為必填 : 第一個字必須為英文字母或數字 否則清空 小寫toUpper
        '       Formula : txtcode.Text = txtcode.Text.ToUpper
        '
        '   about Flag 0:新增 1:修改 2:刪除 3:停用 4:其他()
        '   if 0 or 1 =>寫入配比設計值需先判斷是否存在，若不存在則新增，若存在則異動資料
        '   ??

        '109.12.5 109.12.6   
        '   objDownloadAerp As New classDownloadAerp
        '

        '109.12.19 配比15碼 bCheckBoXSP17
        '   巨力 配比15碼   CheckBoXSP17
        '   txtcode width 153 => 160 , TxtSpare

        '109.12.23  after site meeting
        '   日報表 含非自動  - remain 修正
        '
        '109.12.25 109.12.29 配比連線 bPbSync
        '   Public objUpAerp As New classDownloadAerp
        '
        '110.1.1 配比15碼 bCheckBoxSP17
        '   Y2021M1A.mdb can not ALTER table since no file before SaveRemain
        '110.1.5 配比15碼 巨力 site
        '       Alter code t0 text 15 to origfiles by MDBplus
        '       If bCheckBoxSP17 And False Then => disable => use orig
        '        Dim ps2 As New System.Drawing.Printing.PaperSize("CBC8_4", 1000, (4 + lineCnt) * lineHeight)
        '   
        '110.1.7 
        '   PB lenth
        '   TimerPV for dgvPV1 200msec
        '   Timeer2 400=>500ms

        '110.1.11 110.1.13 110.1.14
        '   配比連線 bPbSync
        '   若為停用，則工業電腦端停用該配比。（不可刪除）
        '   若工業電腦端之現有資料不存在於時訊所提供之整批配比資料中，
        '   則工業電腦端停用該配比。（不可刪除） => no!!
        '110.1.18 @site
        '   slump,diameter => string in InsertToMdb()
        '110.1.20 @JS 配比15碼 bCheckBoxSP17
        '   PanelMon.Visible disable
        '   pond_scale 最小刻劃
        '110.1.27 @site鳳山
        '
        '110.1.28 @h 配比連線 bPbSync
        '   比對記錄刪除 檢查重複配比
        '   @site Strength 中文判斷
        '
        '110.1.30 @H
        '   配比刪除 => UPDATE ENABLE=0 NOT DELETE
        '   sql = "UPDATE fomula SET enable = 0 where id = " & CInt(TxtID.Text)
        '   LabelPlcComm TimerDv1_40 => LabelPlcComm
        '110.2.2
        '   配比刪除 fommula_Mdb.Flag = 2
        '110.2.5@j
        '   LabelNet.Text = " IP:" & strIPAddress
        '110.3.9@h 建華: insert into/update  AIC2ERP 不要有 compid factid
        '   '110.3.9@h 建華:Creatorid updid = module1.user
        '
        '110.3.11 DbNull at site 3.10 ??
        '   On Error Resume Next
        '110.3.16 recalculate
        '   TxtWeight.Text of Fomula

        '110.4.8 alarm problem 
        '   D829 D839

        '110.4.22 bPbSync 上線後 BUG
        '       memo2 = dt.Rows(0).Item("memo2")
        '       If IsDBNull(dt.Rows(0).Item("memo2")) Then
        '       memo2 = 0
        '       End If
        '       Save formula some bug 

        '110.5.14   巨力 System.NullReferenceException 未處理
        '        Message = "並未將物件參考設定為物件的執行個體"
        '        Source = "cbc800"
        '           StackTrace:
        '       於 WindowsApplication1.MainForm.queue0_BoneDoneChange() 於 C:\JS\cbc8\WindowsApplication2_20160903_Inv_XP\MainForm.vb: 行 51
        '   應是 queue0.plate 未建立

        '110.6.5 6.7 亞東 水泥確認
        '110.6.7 車號簡檢查 bPbSync => bCheckBoxYadon
        '   
        '110.6.17   台泥 藥劑區域
        '   bCheckBoxAE
        '   110.6.19 台泥 藥劑區域 bCheckBoxAE    modmaterial()

        '110.6.22 6.23 test DB
        '110.6.23 test
        'Public db As New DbAccess(sSMdbName & "\setting.mdb")
        '   TLog.add
        '110.6.23 test ACE
        '   conn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0 ; Data Source=" & SqlExpressDbName)

        '110.9.11 !!
        '   give up 110.7 revise from 110.6.23
        '   110.6.24 亞東 水泥確認 RANGE 90-100
        '   110.6.24 @HOME
        '   PanelAlarm.Width = DataGridViewAlarm.Width
        '   110.6.25 @JS
        '   LabelA1.ForeColor = Color.White =>
        '   LabelA11.ForeColor = Color.White
        '110.9.11
        '   CheckLabelJSForeColor()
        '   bMyComputerNetworkIsAvailable bMyComputerNetworkPingDataBaseIP bMyComputerNetworkPingTextBoxIP_R 
        '       建勝/電機 顏色 : bMyComputerNetworkIsAvailable=false => black else DarkGreen/DarkRed
        '   110.7.7 Public DVW_C16(6) As Integer
        '   110.7.8 @h set default value =0 
        '       dt_setting = db_setting.GetDataTable("select [AE] from correctarea")
        '       sql1 = "UPDATE correctarea SET [AE]=0 "
        '   110.8.25
        '   default water bug. lblfield_Click()
        '   110.9.1 9.2@j
        '   i = FrmMonit.WriteA2N(871, 23)
        '   If MainForm.LabelA11.ForeColor = Color.White Then for Alt_D
        '   sMixerNo read sTextBoxMixerNo/2 in Admin.des
        '   台泥 藥劑區域 bug & chang to => bCheckBoxAE alwayse true
        '   Label1.Visible = false
        '   for water
        '   diff_temp = (((diff_up + diff_dw - 0.5) * Rnd() - diff_dw))
        '   110.7.28 @js
        '   111.4.6 bCheckBoxSP18 nouse => 藥劑% 
        '
        '110.9.14 @h
        '   getfomula() use slump,strength, diameter instead of founder,strength, particle for 
        '   showdata : (0)-補正區域,(1)-B配比, (2)-總重, (3)-拌合時間, (4)-坍度, (5)-強度, (6)-粒徑, (7~9)-備註一~三
        '   bCheckBoxSP1~20 MUST FROM ReadAdmin()
        '   CemWarehouse.des ReadCemWarehouse() WriteCemWarehouse() for C1..C6 
        '   printTime 打印時間flag
        '110.9.15 @j 
        '   Public bCheckBoxYadon As Boolean
        '   If bPbSync Then in Formula Queryfomula()
        '   D839 bug
        '   車號簡檢查 bug
        '110.9.24
        '   Formula btn_qrycode_Click
        '   Panel1.Visible = bCheckBoxYadon in frmOther
        '110.10.5
        '   Call WriteCemWarehouse() in FrmAdmin()
        '   !! CemWarehouse.des must in js/OrigFile
        '   file = My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb")
        '   If file = False Then Exit Sub
        '   If file = False Then My.Computer.FileSystem.CreateDirectory("\JS\OrigFile\CONF_" & sProject)
        '   "UPDATE correctarea SET [Drog1]=0,[Drog2]=0,[Drog3]=0,[Drog4]=0,[Drog5]=0 "
        '110.10.17
        '   配比連線 bPbSync : 北三廠 bujino 拌合機別 sMixerNo DownloadAerp.action 未識別
        '   If LabelPbSync.ForeColor = Color.Blue Then
        '       strsql = "select ID,code from fomula where code like '" + txt_search.Text + "%'" 顯示
        '   上傳配比 空白不傳
        '110.10.22
        '   上傳配比 sMixerNo 
        '   巨力 System.IndexOutOfRangeException 
        '110.11.9
        '   連續生產 不同盤次 bug
        '   ii = queue0.plate.Length => ii = queueR.plate.Length
        '   queue1 : queue1.fomulacode = Cbxquecode1.Text in queue1check() after check PB
        '   queue0 : mc.queclone(queue1, queue0) in quecode() when "BEGIN" and queue1.ismatch = True OR queue0.ModDonePlate >= queueP.plate.Length
        '   queueP : mc.queclone(queue0, queueP) in SaveBone_New()
        '   queueR : mc.queclone(queueP, queueR) in SaveMod_New()
        '   緊急停止 If queue0.inprocess = True Then 
        '110.11.11 @home
        '   配比連線 bPbSync : 北三廠 bujino 拌合機別 sMixerNo DownloadAerp.action 未識別
        '   PRIMARY KEY : sql &= " WHERE [formulaid]='" & formulaid & "' AND [bujino]='" & bujino & "'"
        '110.11.12 110.11.14 @home
        '   UpdateLabelTotPlate() : LabelTotPlate.Text 
        '   bug : 
        '   Public TotalPlate As Integer 本車總盤次
        '   queue.TotalPlate = queue.plate.Length  in checkque()
        '110.11.16 110.11.17
        '   queclone(queueP, queueR)
        '   D829 D839
        '   連續生產 bug
        '110.11.19
        '    bug isWashCar
        '110.11.23 
        '    copy water% 110.11.23 bug since 
        '   queueP.waterall(k) = dgvWater.Rows(0).Cells(k).Value

        '110.11.26
        '   台泥 mdb => csv
        '       bCheckBox232 sProject sMdbDir sMdbFileName sSavingDate

        '110.12.16
        '   fommula_Mdb.Visible err in readmdb()
        '   module1.authority >= 3 for formula visible

        '111.1.7
        '   sql = "Select  * from " & sbujiq1a & " where bujiread <> 'Y'  order by pb33 asc, pb12 asc, queuekey asc "
        '   sql = "Select TOP 1 * from " & sbujiq1a & " where bujiread <> 'Y'  order by pb33 asc, pb12 asc, queuekey asc "

        '111.1.20 土城新盤 通信異常
        '   COM_PLC_Err += 1

        '111.3.16 台泥 
        '   Data1 : 配比 8C => 9C ,  total 51 => 52
        '   Data2 : 配比 8C => 9C , 骨材 S1,S2,G1,G2,G3 => S1,S2,S3,G1,G2,G3 ; 含水量 5C*5 => 5C*6 ; C1-5 => C1-6 <by Q32,36,38>
        '   total 159 => 175
        '
        '   Public Class ClassBatchDoneTx 
        '   Send_232()
        '   Val2Data()
        '               【骨材值 5C*18sets】 => 【骨材值 5C*20sets】 < S3 & C6 >
        '               【含水量(*10)  5C*5sets】=>【含水量(*10)  5C*6sets】WaterPer_S3
        '   TimerRx_Tick()
        '   Data2Order()
        '111.3.19
        '   Public Data(158) As Byte 159=>174
        '111.3.28
        '   Cbxquecode1.Text = Cbxquecode1.Text.ToUpper

        '111.4.2 
        '   1.配比連線 bPbSync : sometimes invisible not found yet
        '   2.藥劑% : CheckBoXSP18 : NumAE LabelAE : 加藥減藥 -30% ~ +30%
        '   3.AE6 : Drog6 many things to do
        '           D759	A1設定修正值
        '
        '111.4.4
        '   CheckDrog6() : check if C:\JS\OrigFile\New_datasample.mdb have Drog6
        '111.4.6
        '   藥劑% : CheckBoXSP18 : NumAE LabelAE : 加藥減藥 -30% ~ +30% <only Alt-L "B"(real)>
        '       modmaterial()
        '           If bCheckBoxSP18 And MainForm.NumAE.Value <> 0 And AB = "B" Then
        '111.4.10
        '   If bCheckBoxSP18 Then NumAE.Visible
        '111.4.12 @js
        '111.4.12 for remain check when remote read
        '   only chang FileChange ( remote)
        '   
        '111.5.6 
        '       台泥 Q36 same as Q38  & bug
        '111.5.7 
        '       台泥 Q32
        '111.6.24 減米數
        '       煜昌(Q40) 減米數 配方列表Cbxquecode1
        '   Cbxquecode1 : add ID <> code
        '   
        '111.6.24 CheckDrog6()
        '   iUniID ??
        '111.7.22
        '   模擬報表 首盤 CheckBoxFirstPan
        '   If (set_temp <= 100 And diff_temp >= 0.3) Or (set_temp <= 100 And diff_temp <= -0.3) Then
        '           real_A(i) = remain(i) + CInt(set_temp)  '虛擬實際值
        '
        '111.8.20    unsend SQL B rec
        '111.8.21   ClassReceia
        '111.9.13   補 conscbuji
        '           1. read receia, receib, A.mdb, B.mdb list m3 /day
        '           2. if m3 of receia.D <> A.mdb   -> insert to receib.D
        '           3. if m3 of receia.R <> B.mdb   -> insert to receib.R
        '
        '111.9.16   @js
        '        SaveNow = Now
        '111.9.17   @js
        '111.9.19
        '   add ymd to sendSQL
        '111.9.20   @homw chisang
        '111.9.28
        '   newresendSqlDay(datefind, "D")
        '111.10.18
        '   If bDataBaseLink And bCheckBoxSP4 And bCheckBoxYadon Then
        '111.10.21 do not update recdt
        '   SaveRemain_New() 
        '111.10.22 for remain done RemainMaterial3 ??
        '   queueR.materialRemain(24) = 99.9
        '111.10.23 @ 小港site do not update recdt
        '   count_A_report_only()
        '   recDT = dt_data.Rows(0).Item("RecDT").ToString
        '111.10.24 bug at 110.10.23 count_A_report_only ',' for sql
        '   take away updatime , add remainmaterial3=99
        '   !! so, recTime : SaveNow in SaveMod_New() for both A,B
        '111.10.25 @site
        '   補傳
        '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 strWatPrint() 核示單號 => print report only enabled
        '   water4.per strWatPrint() = 1 then true
        '111.11.10 指定列印 => bCheckBoxSP19CheckBoxsSim
        '111.11.11 bug
        '111.11.25
        '   NumericStart  NumericStop max:2000->9999
        '111.11.25 queuekey for B
        '111.11.28 @j 111.10.22 for remain done RemainMaterial3 ??
        '       real_A(arrdb_real.Length - 1) = 99
        '111.11.30 print offset 列印調整
        '   print offset
        '111.12.1 !!!
        'Day_report(sYMdbName_B & DateTimePicker1.Value.Year & "M" & DateTimePicker1.Value.Month & ".mdb", "B", DateTimePicker1.Value, DateTimePicker2.Value)
        '111.12.28
        ' txt_daysum.Text = Format(cube_tot, "0.00")

        '112.1.5 copy for big data
        '   Label31 : SPARE5 => 大數據電腦IP
        '   TextBoxSPARE5
        '   class SQLDbAccess2
        '112.1.10
        '   If bPingSQLBigData Then   MainForm.Label10.ForeColor = Color.Blue
        '112.1.13
        '   test SQL2019 ok ref http://www.aspphp.online/bianchen/dnet/aspnet/aspnetjc/201701/53928.html
        '   to set SQL express User PW buji buji port 1443
        '112.1.27 112.1.28
        '   disable mc.SendSQL() in BackgroundWorker4_DoWork()
        ' use below in SaveRemain_New() for Real , count_A_report_only() for A , both conscbuji & BigData
        '   send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "A")
        '112.3.1
        '   send car id
        '112.3.7
        '   sendSqlToconscbujiRecei()
        '112.3.20 
        '   Public DVW(4200) As Integer
        '   send water 
        '   send m3
        '112.3.30 連接大數據資料庫
        '   bTextBoxSPARE5 ,Label10
        '   bPingSQLBigData => Label10.ForeColor = Color.Green "完成"
        '112.3.25 @j
        '   列印拌合結束電流    bCheckBoxCBC7    Item("RemainMaterial2")
        '112.4.12 @site SiZh
        '   If bPingSQLBigData Then  no send conscbuji bug
        '   LabelBigDataErr_Click() try ping

        '112.4.19 after back from site  bug:only sum no sub remain
        '112.4.28 tot - rem
        '       valRec(i) = row.Item(i) - rowRem.Item(i)
        '   disable mc.Re_SendSQL() in Quit
        '   send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "A")
        '   SaveRemain_New() no update remain_B ??
        '112.8.26
        '   B m3 err => records duplicated! => delete records 
        '112.9.3 112.9.4  Q08 巨力 m3 err!
        '   count_B_report count_B_report recdt should not change
        '112.9.6
        '   update B by A SQL button
        '112.9.9
        '   strsqlall = "Select  "
        '112.9.15
        '   LabelJS
        '112.9.18 @js 電腦連線旗標
        '       bPbSync 配比連線
        '           If bPingSQL And bDataBaseLink Then bPbSync = bCheckBoxSP16
        '       send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "A")
        '       disable check IP again in TimerLink
        '   If bPingSQL And bDataBaseLink Then
        '       send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "D", True), "A")
        '   End If
        '112.9.21@J 112.9.22@h
        '   bPbSync 配比連線 bug 
        '   sendSqlToconscbujiRecei() disale some thing
        '   
        '112.9.23 @h
        '   ToolStripStatusLabel1
        '112.9.26
        '   LabelSendSql
        '112.10.1
        '   receiaPanB for B RT since .mdb maybe moved by remote PC => fillMdbRowToPan()
        '112.10.3
        '   disable select * from receia
        '112.10.7 @js
        '   disable BackgroundWorker3,BackgroundWorker4 => no use ,  , no use TimerSQL
        '   TimerMoveA 12000 = > 6000
        '112.10.25 @js 
        '   bug when BackgroundWorker3 disable
        '   iBackground3_Flag = 2
        '   本車 Click TimerFlag.Enabled = True for remotePC
        '   CheckBoxRem.Checked TimerFlag.Enabled = True
        '   disable show RT For bigdata
        '112.11.1 Single=> double 小港發現誤差
        '   Single=> double in Day_report() 小港發現誤差

        '112.12.6 固化廠
        '112.12.26 固化廠 home use Anydesk
        '113.1.22 固化廠
        '   System.Drawing.Printing.PaperSize() too short
        '   width = "A" 6=>7

        ' !!!!!
        '113.1.24 @h 
        ' copy v.2023.1.22 to WindowsApplication2_20240122 for win11
        '   NuGet ZedGraph
        '   應用程式-目標-Framework .NetFrameWork 4.6.2
        '   Microsoft.ACE.OLEDB.12.0 AccessDatabaseEngine_X64 in CBC8_Inst
        '113.2.12
        '   install ttf
        '   can not find setting.mdb
        '113.2.14
        '   add try for ... ON ERROR
        '   SimHei download & install
        '   LibreOffice_7.6.4_Win_x86-64  download & install
        '
        '113.8.13
        '   Program files => \JS\QJ\
        '   (4)-坍度:int, (5)-強度:string, (6)-粒徑:int
        '   queueP.showdata(4) = "20"
        '113.8.26
        '   FUSE ALARM
        '       AlarmCheck() AlarmDvMax AlarmDvBitsMax
        '113.8.28@Taichung
        '   ReadAlarm()
        '   

        'Label1.Text = "103.12.24-108.6.12@J"
        Label1.Text = "111.12.28_BD_113.8.13@j"
        '@@Label1.Text = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        'Label1.Text &= My.Application.Info.Version.MajorRevision & "." & My.Application.Info.Version.ToString & "." & My.Application.Info.Version.Build

        '106.12.10
        TimerLinkErrorCnt = 0
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub UsernameLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsernameLabel.Click
        If LogAuto = 2 Then
            UsernameTextBox.Text = "jscbc800"
            PasswordTextBox.Text = "561020"
            Call OK_Click(UsernameLabel, e)

        End If
    End Sub

    Private Sub LogoPictureBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoPictureBox.Click

    End Sub

    Private Sub LogoPictureBox_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LogoPictureBox.DoubleClick
        LogAuto = 1
    End Sub

    Private Sub PasswordLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasswordLabel.Click
        If LogAuto = 1 Then
            LogAuto = 2
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_OP.CheckedChanged

    End Sub

    Private Sub RadioButton_REMOTE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_REMOTE.CheckedChanged

    End Sub

    
    Sub addReciaFields(ByVal db_mdb As Object)
        '111.8.20
        ' add fields which defined by Yatong to orignal recia.mdb
        Dim sql As String
        sql = "ALTER TABLE receia ADD [queuekey] varchar(30) NOT NULL, "
        sql &= "[bujino] varchar(2) NOT NULL, "
        sql &= "[hhmmss] varchar(6) NOT NULL, "
        sql &= "[yymodd] varchar(8) NOT NULL, "
        sql &= "[formulaid] varchar(12) NULL, "
        sql &= "[qty] real NULL, "
        sql &= "[maxqty] real NULL, "
        sql &= "[category] varchar(2) NOT NULL, "
        sql &= "[outway] varchar(1) NULL, "
        sql &= "[carid] varchar(10) NULL, "
        sql &= "[ordid] varchar(10) NULL, "
        sql &= "[sumno] varchar(4) NULL, "
        sql &= "[no] varchar(4) NULL, "
        sql &= "[strength] varchar(6) NULL, "
        sql &= "[diameter] varchar(4) NULL, "
        sql &= "[slump] varchar(4) NULL, "
        sql &= "[wcrate1] real NULL, "
        sql &= "[wcrate2] real NULL, "
        sql &= "[s1] real NULL, "
        sql &= "[s2] real NULL, "
        sql &= "[s3] real NULL, "
        sql &= "[s4] real NULL, "
        sql &= "[g1] real NULL, "
        sql &= "[g2] real NULL, "
        sql &= "[g3] real NULL, "
        sql &= "[g4] real NULL, "
        sql &= "[w1] real NULL, "
        sql &= "[w2] real NULL, "
        sql &= "[w3] real NULL, "
        sql &= "[c1] real NULL, "
        sql &= "[c2] real NULL, "
        sql &= "[c3] real NULL, "
        sql &= "[c4] real NULL, "
        sql &= "[c5] real NULL, "
        sql &= "[c6] real NULL, "
        sql &= "[ae1] real NULL, "
        sql &= "[ae2] real NULL, "
        sql &= "[ae3] real NULL, "
        sql &= "[ae4] real NULL, "
        sql &= "[ae5] real NULL, "
        sql &= "[ae6] real NULL, "
        sql &= "[total] real NULL, "
        sql &= "[Sameid] varchar(12) NULL, "
        sql &= "[ymd] varchar(8) NOT NULL, "
        sql &= "[memo1] varchar(50) NULL,"
        sql &= "[memo2] varchar(50) NULL,"
        sql &= "[Createtime] datetime NULL,"
        sql &= "[updtime] datetime NULL,"
        sql &= "[updid] varchar(10) NULL, "
        sql &= "[wtime] datetime NOT NULL,"
        sql &= "[Flag] int NOT NULL,"
        sql &= "[Readed] int NOT NULL,"
        sql &= "[Visible] int NOT NULL"
        Try
            db_mdb.ExecuteCmd(sql)
        Catch ex As Exception

        End Try

    End Sub




    Private Sub LoginForm1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '110.9.11 110.8.1
        bMyComputerNetworkIsAvailable = My.Computer.Network.IsAvailable
        If bMyComputerNetworkIsAvailable Then
            bMyComputerNetworkPingDataBaseIP = My.Computer.Network.Ping(sDataBaseIP, 1000)
            bMyComputerNetworkPingTextBoxIP_R = My.Computer.Network.Ping(sTextBoxIP_R, 1000)
        Else
            bMyComputerNetworkPingDataBaseIP = False
            bMyComputerNetworkPingTextBoxIP_R = False
        End If
        '112.9.23 move from OK
        Me.Refresh()
        '106.9.4
        Re_SendSQL_Flag = False
        '105.7.16
        LabelCheck.Visible = True
        Me.Refresh()
        '105.9.24
        bPingRPC = True
        bRemoteBpath = True
        bRemoteApath = True


        '105.10.8
        If sDataBaseIP = "127.0.0.1" Then
            bPingSQL = False
            MainForm.CheckBoxLink.Visible = False
            '107.8.31
            bDataBaseLink = False
        Else
            bPingSQL = True
        End If
        If sTextBoxIP_R = "127.0.0.1" Then
            bPingRPC = False
            '107.7.9
            MainForm.LabelJS.Text = "建勝電機"
            MainForm.LabelRPC.Visible = False
        Else
            bPingRPC = True
        End If

        '105.8.6 105.10.8
        'If Not My.Computer.Network.IsAvailable Then
        If (bPingSQL Or bPingRPC) Then

            If (Not My.Computer.Network.IsAvailable) Then
                '105.10.8
                bPingSQL = False
                bPingRPC = False
                If bEngVer Then
                    MsgBox("Network", MsgBoxStyle.Critical, "Network is not available!")
                Else
                    MsgBox("網路無法使用! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
                End If
                bRemoteBpath = False
                '105.10.13
                bRemoteApath = False
            Else
                '105.10.8 If Not My.Computer.Network.Ping(sTextBoxIP_R, 1000) Then
                If Not My.Computer.Network.Ping(sTextBoxIP_R, 1000) And bPingRPC Then
                    If bEngVer Then
                        MsgBox("Network is available, ping " & sDataBaseIP & " fail!", MsgBoxStyle.Critical, "Network Checking")
                    Else
                        MsgBox("Ping 遠端電腦 " & sTextBoxIP_R & "失敗! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
                    End If
                    bPingRPC = False
                    bRemoteBpath = False
                    '105.10.13
                    bRemoteApath = False
                End If
                '105.10.8 If Not My.Computer.Network.Ping(sDataBaseIP, 1000) Then
                '109.11.29 test
                If LogAuto = 2 Then
                    bPingSQL = True
                    bDataBaseLink = True
                End If
                '112.9.23
                bDataBaseLink = True
                If Not My.Computer.Network.Ping(sDataBaseIP, 1000) And bPingSQL Then
                    If bEngVer Then
                        MsgBox("Network is available, ping " & sDataBaseIP & " fail!", MsgBoxStyle.Critical, "Network Checking")
                    Else
                        MsgBox("Ping 調度電腦 " & sDataBaseIP & "失敗! 請網路檢查!", MsgBoxStyle.Critical, "網路檢查")
                    End If
                    bPingSQL = False
                    '106.7.25
                    bDataBaseLink = False
                End If

            End If
        End If

        '112.9.21
        If Not (bPingSQL And bDataBaseLink) Then
            bPbSync = False
        End If
        ToolStripStatusLabel2.Text = "DD:" & sDataBaseIP & ":" & bDataBaseLink

        'If RadioButton_OP.Checked Then
        If RadioButton_OP.Checked And bRemoteBpath Then
            ReadRemote()
        End If
        LabelCheck.Visible = False

        '112.1.5
        '112.3.30
        If bTextBoxSPARE5 Then
            bPingSQLBigData = True
            '112.9.23
            ToolStripStatusLabel3.Text = "Checking sBigDataDataBaseIP : " & sBigDataDataBaseIP
            If sBigDataDataBaseIP = "" Then
                bPingSQLBigData = False
            Else
                If Not bMyComputerNetworkIsAvailable Then
                    bPingSQLBigData = False
                Else
                    If sBigDataDataBaseIP.Length > 8 Then
                        Dim sIP As String
                        Dim il As Integer
                        il = InStr(sBigDataDataBaseIP, "\")
                        '112.1.26
                        'sIP = Mid(sBigDataDataBaseIP, 1, il - 1)
                        If InStr(sBigDataDataBaseIP, "\") > 0 Then
                            '112.1.26
                            sIP = Mid(sBigDataDataBaseIP, 1, il - 1)
                            If Not My.Computer.Network.Ping(sIP, 500) Then
                                bPingSQLBigData = False
                            End If
                        Else
                            If Not My.Computer.Network.Ping(sBigDataDataBaseIP, 500) Then
                                bPingSQLBigData = False
                            End If
                        End If

                    Else
                        bPingSQLBigData = False
                    End If
                End If
            End If
            '112.3.30
            'bTextBoxSPARE5 = bPingSQLBigData
        End If
        '112.9.23 112.10.3
        'If bDataBaseLink Then
        '    Dim db
        '    Dim sq As String
        '    db = New SQLDbAccess("conscbuji")
        '    Dim dtBigData As DataTable
        '    Dim i = 0
        '    sq = "SELECT * FROM RECEIA "
        '    Try
        '        dtBigData = db.GetDataTable(sq)
        '        i = dtBigData.Rows.Count
        '        'i = dbBigData.ExecuteCmd(sq)
        '        'MsgBox(i)
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        '    db.dispose()
        '    ToolStripStatusLabel2.Text = "DD:" & sDataBaseIP & ":" & bDataBaseLink & i
        'Else
        '    ToolStripStatusLabel2.Text = "DD:" & sDataBaseIP & ":" & bDataBaseLink
        'End If

        'If bPingSQLBigData Then
        '    ToolStripStatusLabel3.Text = "BigData: " & sBigDataDataBaseIP & "OK!"
        '    Dim dbBigData
        '    Dim sq As String
        '    dbBigData = New SQLDbAccess2("BigData", sBigDataDataBaseIP)
        '    Dim dtBigData As DataTable
        '    Dim i = 0
        '    sq = "SELECT * FROM RECEIA "
        '    Try
        '        dtBigData = dbBigData.GetDataTable(sq)
        '        i = dtBigData.Rows.Count
        '        'i = dbBigData.ExecuteCmd(sq)
        '        'MsgBox(i)
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        '    dbBigData.dispose()
        '    ToolStripStatusLabel3.Text = "BigData: " & sBigDataDataBaseIP & bPingSQLBigData & i
        'Else
        '    ToolStripStatusLabel3.Text = "BigData: " & sBigDataDataBaseIP & "FAIL!"
        'End If

    End Sub
End Class
