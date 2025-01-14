Public Class MainForm
    '112.10.1
    Dim receiaPanB As New classReceiaPan()

    '111.1.20
    Public COM_PLC_Err As Integer
    '110.9.11 110.6.24
    Public labelAlarmTime(10) As Label
    Public labelAlarmDesc(10) As Label

    Public WithEvents queue0 As New product
    '99.08.02  for wait remain 106.9.8 => infact clone when D20 Bone change , so maybe overwrite again befor D22 change
    Public WithEvents queueP As New product
    '106.9.8  for wait remain clone when D21 change
    Public WithEvents queueR As New product
    Public exist(25), sim_addshow_weight As Boolean  'sim_addshow_weight:模擬報表是否加入每盤重量
    'Private db As New DbAccess("\JS\CBC8\files\setting.mdb")
    Private db As New DbAccess("\JS\CBC8\Conf_" & sProject & "\setting.mdb")
    '100.4.4
    Private db_recipe As New DbAccess("\JS\CBC8\Conf_" & sProject & "\Recipe.mdb")

    '103.6.7
    'Private mc As New countmaterial
    Public mc As New countmaterial

    Private plateindex, reptype As Integer  'reptype 報表形式
    Private emgstop, begin As Boolean
    'Private boneplate, modplate, remain, watertemp As Integer '計量中盤次
    Private remain, watertemp As Integer '計量中盤次
    'Public BoneDonePlate, ModDonePlate As Integer '計量完成盤次 come from D20,21
    'Private dbrec As New DbAccess("\JS\CBC8\data\Y2010.mdb")
    Private dbrec As New DbAccess("\JS\CBC8\Data_" & sProject & "\Y2010.mdb")
    Private waterall(8), Fbone(32), Bbone(32), FMod(51), BMod(51) As Single
    Private unicar, flag_last, flag_unicar As String 'batch生產為一序號
    Private strsql As String '
    Private sqlconn As OleDb.OleDbConnection
    Private file_A, file_B, file_print As String '每日記錄文字檔檔名
    Private PLCconst As New PLC_Const
    Public VGAP As Integer
    Private i As Integer
    Private combow_pos As Integer 'for combow
    Private FormLoded As Boolean
    Private bSaveRemainDone As Boolean
    Private bPlcErrOld As Boolean
    Public cub As Single
    '107.6.22 use by thread must public
    ''99.12.09
    'Public sbujiq1a As String

    '102.8.2
    Private RemainDonePlate_temp_SQL As Integer
    Private remain_B_SQL(25) As Single
    Public order232 As New ClassOrder



    Private Sub queue0_BoneDoneChange() Handles queue0.BoneDoneChange
        Dim i, j As Integer
        '110.5.14   巨力 System.NullReferenceException 未處理
        '107.3.9 test 110.5.14
        Dim ii
        '110.11.16
        If queue0.BoneDonePlate >= 1 Then
            mc.queclone(queue0, queueP)
            Call UpdateLabelTotPlate()
        End If
        '106.9.14 copy water% 110.11.23 bug since 
        Dim k
        For k = 0 To 7
            queueP.waterall(k) = dgvWater.Rows(0).Cells(k).Value
            queueP.water_all_W_Limit(k) = water_all_W_Limit(k)
        Next

        Try
            ii = queue0.plate.Length
            '110.10.22
            If queue0.BoneDonePlate > ii Then
                MsgBox("queue0_BoneDoneChange() : 程序有誤 queue0.BoneDonePlate=" & queue0.BoneDonePlate & " queue0.plate.Length= " & ii, MsgBoxStyle.Critical)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("queue0_BoneDoneChange() : 程序有誤 " & ex.Message & "queue0.BoneDonePlate=" & queue0.BoneDonePlate, MsgBoxStyle.Critical)
            Exit Sub
        End Try

        '107.6.2
        Label_PLC_D20.add = "B1"

        Try
            '107.6.2
            MainLog.add = "queue0_BoneDoneChange() : queue0.BoneDonePlate = " & queue0.BoneDonePlate & " , queue0.plate.Length =" & queue0.plate.Length
            '107.5.30
            'If queue0.BoneDonePlate <= queue0.plate.Length Then
            If queue0.BoneDonePlate <= queue0.plate.Length And queue0.plate.Length > 0 Then
                '107.6.2
                Label_PLC_D20.add = "B2"
                If queue0.inprocess = True Then
                    'BoneDoingPlate:計量中的盤次 BoneDonePlate;完成的盤數
                    If queue0.BoneDoingPlate = queue0.BoneDonePlate Then
                        '110.11.16
                        'If queue0.BoneDonePlate > 0 And queue0.BoneDonePlate <= queue0.plate.Length Then
                        If queue0.BoneDonePlate > 0 And queue0.BoneDonePlate <= queueP.TotalPlate Then
                            '完成盤數不等於零, 紀錄骨材至資料庫 =? 106.5.27 no save to mdb
                            '106.5.27 no save to mdb, only fill to queue0
                            'Call SaveBone(queue0.BoneDonePlate, queue0.plate(queue0.BoneDonePlate - 1))
                            Call SaveBone_New(queue0.BoneDonePlate, queue0.plate(queue0.BoneDonePlate - 1))
                        End If
                    End If

                    '101.3.21 move from txtBon
                    If queue0.BoneDonePlate = queue0.plate.Length Then

                        queue1check()

                        '107.6.2
                        Label_PLC_D20.add = "B4"
                        If CheckBoxContinue.Checked = True And queue1.ismatch = True Then
                            '107.6.2
                            Label_PLC_D20.add = "B5"
                            queue1.inprocess = True
                            'send 骨材 of queue1.fmaterial & queue1.bmaterial
                            '將send to PLC骨材資料暫存到Fbone
                            For i = 0 To 7
                                Fbone(i) = queue1.Fmaterialorg(i)
                                Fbone(i + 8) = queue1.Materialfront(i)
                                Fbone(i + 16) = queue1.Fmaterialnowater(i)
                                Fbone(i + 24) = waterall(i)
                                Bbone(i) = queue1.Bmaterialorg(i)
                                Bbone(i + 8) = queue1.Materialback(i)
                                Bbone(i + 16) = queue1.Bmaterialnowater(i)
                                Bbone(i + 24) = waterall(i)
                            Next
                            '骨材搶先
                            ' reset PanCount
                            txtbone.BackColor = Color.Yellow
                            dgvPan.Rows(0).Cells(1).Style.BackColor = Color.Yellow
                            dgvPan.Rows(0).Cells(1).Value = Format(Val(queue1.plate(0)), "0.00")
                            LabelMessage.Text = "骨材搶先中"

                            '100.6.18
                            '104.9.24 only disable here
                            '107.5.30 move from below
                            Cbxquecode1.Enabled = False
                            numquetri1.Enabled = False
                            Tbxquecar1.Enabled = False
                            Tbxquefield1.Enabled = False
                            btnque1.Enabled = False
                            btncancel1.Enabled = False
                            '107.5.30
                            Cbxquecode1.Refresh()

                            MainLog.add = "queue0_BoneDoneChange() : 骨材搶先中 BMR= " & queue0.BoneDonePlate & queue0.ModDonePlate & queue0.RemainDonePlate & " plt.Len:" & queue0.plate.Length
                            LabelMessage.Visible = True
                            DVW(835) = 1
                            i = FrmMonit.WriteA2N(835, 1)

                            'Cbxquecode2.Enabled = False
                            'numquetri2.Enabled = False
                            'Tbxquecar2.Enabled = False

                            '107.6.14
                            MainLog.add = "CheckDefaultWater() : 骨材搶先"
                            '104.8.24
                            CheckDefaultWater(queue1.showdata(1), Tbxquefield1.Text)
                        Else
                            '107.6.28 only when no 骨材搶先
                            'reset D741
                            For i = 0 To 7
                                DVW(i + 741) = 0
                                DVW(i + 771) = 0

                                '106.5.27 106.9.4 should be clear
                                queue0.Fmaterialorg(i) = 0
                                queue0.Bmaterialorg(i) = 0
                                queue0.Materialfront(i) = 0
                                queue0.Materialback(i) = 0
                            Next
                            i = FrmMonit.WriteA2N(741, 8)
                            i = FrmMonit.WriteA2N(771, 8)
                        End If

                        If queue0.BoneDonePlate = queue0.plate.Length Then
                            txtbone.BackColor = Color.Green
                        Else
                            txtbone.BackColor = Color.LightGreen
                        End If
                        MonLog.add = "displaymaterial() : BoneDoneChange"
                        displaymaterial()
                    ElseIf queue0.BoneDonePlate < queue0.plate.Length Then
                        '未至最後一盤
                        j = -1
                        '補水計算
                        For i = 0 To 7
                            waterall(i) = dgvWater.Rows(0).Cells(i).Value
                            If exist(i) Then
                                j += 1
                            End If
                        Next

                    End If

                End If
            End If
            queue0.DisplayPlates()
        Catch ex As Exception
            MsgBox("queue0_BoneDoneChange() : " & ex.Message, MsgBoxStyle.Critical)
            '107.6.2
            Label_D20.Text &= "B6"
        End Try
    End Sub

    Private Sub queue0_ModDoneChange() Handles queue0.ModDoneChange

        '110.5.14   巨力 System.NullReferenceException 未處理
        Dim ii As Integer
        '107.3.9 test 110.5.14
        Try
            ii = queue0.plate.Length
            '110.10.22
            If queue0.ModDonePlate > ii Then
                MsgBox("queue0_ModDoneChange() : 程序有誤 queue0.ModDonePlate=" & queue0.ModDonePlate & " queue0.plate.Length= " & ii, MsgBoxStyle.Critical)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("queue0_ModDoneChange() : 程序有誤 " & ex.Message & "queue0.ModDonePlate=" & queue0.ModDonePlate, MsgBoxStyle.Critical)
            Exit Sub
        End Try


        '107.4.10
        TLog.add = " ModDoneChange()"
        '107.6.2
        Label_PLC_D21.add = "B1"
        If queue0.ModDonePlate <= queue0.plate.Length Then
            If queue0.inprocess = True Then
                If queue0.ModDonePlate > 0 Then
                    strTrendPb = queue0.showdata(1)
                    strTrendM3 = (queue0.plate(queue0.ModDonePlate - 1) * 100).ToString
                    '112.3.20 
                    '113.1.22 only Bigdata
                    'DVW(4009) = queue0.plate(queue0.ModDonePlate - 1) * 100
                    'i = FrmMonit.WriteA2N(4009, 1)
                    If bTextBoxSPARE5 Then
                        DVW(4009) = queue0.plate(queue0.ModDonePlate - 1) * 100
                        i = FrmMonit.WriteA2N(4009, 1)
                    End If

                End If

                If queue0.ModDonePlate = queue0.plate.Length Then
                    Txtmod.BackColor = Color.DeepPink
                Else
                    Txtmod.BackColor = Color.LightPink
                End If

                'modDoingPlate:計量中的盤次 modDonePlate;完成的盤數
                If queue0.ModDoingPlate = queue0.ModDonePlate Then
                    '110.11.16
                    'If queue0.ModDonePlate > 0 And queue0.ModDonePlate <= queue0.plate.Length Then '完成盤數不等於零, 紀錄骨材至資料庫
                    If queue0.ModDonePlate > 0 And queue0.ModDonePlate <= queueP.TotalPlate Then '完成盤數不等於零, 紀錄骨材至資料庫
                        '107.6.2
                        Label_PLC_D21.add = "B2"
                        '106.8.26
                        LabelQkey.Text = LabelQkey0.Text
                        '106.5.27  save queue0 to mdb
                        Call SaveMod_New(queue0.ModDonePlate, queue0.plate(queue0.ModDonePlate - 1))

                        If queue0.ModDonePlate < queue0.plate.Length Then
                            '107.6.2
                            Label_PLC_D21.add = "B3"
                            queue0.ModDoingPlate = queue0.ModDonePlate + 1
                            mc.queclone(queue0, queue)
                            queue = mc.Materialdistribute(queue, queue0.ModDoingPlate, cbxcorrect.Text, waterall, "M")
                            mc.queclone(queue, queue0)
                            MonLog.add = "displaymaterial() : ModDoneChange,< :" & queue0.ModDonePlate
                            displaymaterial()
                        End If
                        '99.08.02
                        If queue0.ModDonePlate >= queueP.plate.Length Then
                            '107.6.2
                            Label_PLC_D21.add = "B4"
                            lblwork.BackColor = Color.Lavender
                            queue0.LastPlate = True
                            '110.11.12 110.11.14 @home 110.11.16
                            queueP.LastPlate = True
                            'queueR.TotalPlate = queueP.TotalPlate

                            queue0.inprocess = False
                            '101.12.17
                            '110.9.24 生產完畢開始計時
                            'dtBatDon = Now
                            dtBatDon = Now
                            '106.10.7 聯佶空單
                            lblwork.ForeColor = Color.Blue
                            LabelPb12_0.Visible = False
                            '100.6.29
                            Label12.Visible = False
                            '99.08.07 MOVE FROM TEXTCHANGE
                            '107.9.3 If queue0.ModDonePlate = queue0.plate.Length And queue0.ModDonePlate > 0 Then
                            If queue0.ModDonePlate = queueP.plate.Length And queue0.ModDonePlate > 0 Then
                                '107.6.2
                                Label_PLC_D21.add = "B5"
                                'reset D741 D741	S1設定修正值
                                For i = 8 To 25
                                    DVW(i + 741) = 0
                                    DVW(i + 771) = 0
                                    '99.08.01
                                    queue0.Fmaterialorg(i) = 0
                                    queue0.Bmaterialorg(i) = 0
                                    queue0.Materialfront(i) = 0
                                    queue0.Materialback(i) = 0
                                Next
                                i = FrmMonit.WriteA2N(750, 17)
                                i = FrmMonit.WriteA2N(780, 17)
                                MonLog.add = "displaymaterial() : ModDoneChange(),= :" & queue0.ModDonePlate
                                displaymaterial()
                            End If
                            queue1check()
                            '判斷是否有排隊及自動模式
                            If queue1.ismatch = True And CheckBoxContinue.Checked = True Then
                                '107.6.2
                                Label_PLC_D21.add = "B6"
                                queue0.inprocess = False
                                '107.3.20 聯佶空單
                                If bCheckBoxSP13 And btnque1.ForeColor = Color.Green Then
                                    Dim i
                                    i = MsgBox("確認自動模擬回傳?", MsgBoxStyle.YesNo, "確認自動模擬回傳")
                                    If i = vbYes Then
                                        MixynIng = True
                                        '107.3.20
                                        TLog.add = " 自動模擬回傳"
                                        DV(20) = 0
                                        DV(21) = 0
                                        '107.3.3 ??? if last PAN was not done?
                                        DV(22) = 0
                                        dtp_simtime.Text = LabelPb12_0.Text
                                        '106.10.11
                                        For i = 0 To 24
                                            DVW(i + 91) = 0
                                        Next
                                        i = FrmMonit.WriteA2N(91, 25)
                                    Else
                                        '106.10.11 聯佶空單
                                        clearproduce()
                                        Exit Sub
                                    End If
                                    Dim generator As New Random
                                    Dim randomValue As Integer
                                    Dim s As String
                                    dtSimP(0) = dtp_simtime.Value
                                    If queue0.plate.Length > 1 Then
                                        For i = 1 To queue0.plate.Length - 1
                                            randomValue = generator.Next(1, 10)
                                            s = CInt(tbxmixtime.Text) + Num_quetime_set.Value + randomValue
                                            dtSimP(i) = dtSimP(i - 1).AddSeconds(CInt(s))
                                        Next
                                    End If
                                    Timerprocess_sim.Enabled = True
                                End If
                                '107.3.18
                                'Call SavingLog("ModDone= " & queue0.BoneDonePlate & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate)
                                MainLog.add = ("ModDoneChange() 排隊及自動: queueing queue0.BoneDonePlate=" & queue0.BoneDonePlate & " , Mod=" & queue0.ModDonePlate & " , Rem=" & queue0.RemainDonePlate)
                                '107.6.14
                                MainLog.add = "CheckDefaultWater() : 排隊及自動 quecode()"
                                '110.11.16
                                mc.queclone(queue0, queueP)
                                Call UpdateLabelTotPlate()
                                '106.9.14 copy water% 110.11.23 bug since 
                                Dim k
                                For k = 0 To 7
                                    queueP.waterall(k) = dgvWater.Rows(0).Cells(k).Value
                                    queueP.water_all_W_Limit(k) = water_all_W_Limit(k)
                                Next
                                quecode()  '進入下一batch生產
                            End If
                        End If
                    End If
                End If
            End If
        End If
        '107.6.2
        Label_PLC_D21.add = "B7"
        queue0.DisplayPlates()

        '104.8.25
        If strTrendPb = Nothing Then strTrendPb = queue0.showdata(1)
        If strTrendM3 = Nothing Then strTrendM3 = (queue0.plate(queue0.ModDonePlate) * 100).ToString
        FrmMonit.ReadTrendFile(strTrendPb & "_" & strTrendM3)

    End Sub
    Private Sub queue0_RemiianDoneChange() Handles queue0.RemainDoneChange
        'Dim ct, h, i As Integer
        Dim i As Integer
        '100.5.2
        Dim ti
        ti = Microsoft.VisualBasic.DateAndTime.Timer

        '110.10.22   巨力 System.IndexOutOfRangeException 
        '110.11.12
        'Dim ii As Integer
        'Try
        '    '110.11.9
        '    'ii = queue0.plate.Length
        '    'If queue0.RemainDonePlate > ii Then
        '    '    MsgBox("queue0_RemiianDoneChange() : 程序有誤 queue0.RemainDonePlate=" & queue0.RemainDonePlate & " queue0.plate.Length= " & ii, MsgBoxStyle.Critical)
        '    '    Exit Sub
        '    'End If
        '    ii = queueR.plate.Length
        '    If queueR.RemainDonePlate > ii Then
        '        MsgBox("queue0_RemiianDoneChange() : 程序有誤 queue0.RemainDonePlate=" & queue0.RemainDonePlate & " queueR.RemainDonePlate=" & queueR.RemainDonePlate & " queue0.plate.Length= " & queue0.plate.Length & " queueR.plate.Length= " & ii, MsgBoxStyle.Critical)
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    '110.11.9
        '    'MsgBox("queue0_RemiianDoneChange() : 程序有誤 " & ex.Message & "queue0.RemainDonePlate=" & queue0.RemainDonePlate, MsgBoxStyle.Critical)
        '    MsgBox("queue0_RemiianDoneChange() : 程序有誤 queue0.RemainDonePlate=" & queue0.RemainDonePlate & " queueR.RemainDonePlate=" & queueR.RemainDonePlate & " queue0.plate.Length= " & queue0.plate.Length & " queueR.plate.Length= " & ii, MsgBoxStyle.Critical)
        '    Exit Sub
        'End Try

        '107.6.2
        Label_PLC_D22.add = "B0"
        '106.9.7 107.3.18
        'Call SavingLog("queue0_RemiianDoneChange() queue0.RemainDonePlate:" & queue0.RemainDonePlate & " queueR.plate.Length:" & queueR.plate.Length & " queue0.fomulacode:" & queue0.fomulacode & " queueP.fomulacode:" & queueP.fomulacode)
        '110.11.12
        'MainLog.add = "queue0_RemiianDoneChange() : queue0.RemPlt=" & queue0.RemainDonePlate & " queueR.plt.Len=" & queueR.plate.Length & " queue0.fomu=" & queue0.fomulacode & " queueP.fomu=" & queueP.fomulacode
        '99.08.02
        'If queue0.RemainDonePlate >= 1 And queue0.RemainDonePlate <= queue0.plate.Length Then
        '106.9.8 queueR update by D21 MOD changed
        'If queue0.RemainDonePlate >= 1 And queue0.RemainDonePlate <= queueP.plate.Length Then
        '110.11.14 
        'If queue0.RemainDonePlate >= 1 And queue0.RemainDonePlate <= queueR.plate.Length Then
        If queue0.RemainDonePlate >= 1 Then
            Call SaveRemain_New(queue0.RemainDonePlate)
        End If
        '99.08.02
        'If queue0.RemainDonePlate >= queue0.plate.Length Then
        '107.5.11 排隊自動 前小於後 無法完成 bug
        'If queue0.RemainDonePlate >= queueP.plate.Length Then
        '110.11.14
        'If queue0.RemainDonePlate >= queueR.plate.Length Then
        '110.11.16 ??
        If queue0.RemainDonePlate >= queueR.TotalPlate Then
            '107.6.5
            CheckBoxsSim.ForeColor = Color.Green
            CheckBoxsSim.Refresh()
            '100.7.20 move from below
            TimerPrint.Enabled = True

            '100.5.2
            LabelMsg.Text &= "..報表 : " & (Microsoft.VisualBasic.DateAndTime.Timer - ti)
            'lblwork.BackColor = Color.Lavender
            'queue0.LastPlate = True
            'queue0.inprocess = False
            DVW(22) = 0
            i = FrmMonit.WriteA2N(22, 1)
            '110.11.16 110.11.19 bug
            'queueR.TotalPlate = 0
            'Call UpdateLabelTotPlate()

        End If
        queue0.DisplayPlates()
    End Sub

    Public Sub SaveNonAuto()
        '102.2.17 save C1~4 when D34 bit15 NonAutoProd
        '骨材完成盤數增加 存檔
        'save for all SP and water , So, SaveMod() only  not save SP again! PV will read & save untill D22 2010.0519 
        Dim i As Integer
        Dim d(10) As Integer
        '106.1.3
        Dim sTemp = ""
        sTemp = tbxworkcode.Text
        '105.10.22
        clearproduce()
        Label12.Visible = False
        Cbxquecode1.Enabled = True
        numquetri1.Enabled = True
        Tbxquecar1.Enabled = True
        Tbxquefield1.Enabled = True
        btnque1.Enabled = True
        btncancel1.Enabled = True

        '紀錄配方與骨材資料到database()_recdt
        '102.6.7 原存入 particle 改 orgmaterial1
        'strsql = "Insert into recdata (recdt, unicar,CarSer,particle, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        '103.3.7
        'strsql = "Insert into recdata (recdt, unicar,CarSer,orgmaterial1, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        '103.6.7
        'strsql = "Insert into recdata (recdt, unicar,CarSer,orgmaterial1, fomula, TotalCube, car, field, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        '105.6.13 save nowaterMaterial1=1 for NonAuto
        'strsql = "Insert into recdata (recdt, unicar ,CarSer ,plate ,orgmaterial1, fomula, TotalCube, car, field, strength, particle, founder,  concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        strsql = "Insert into recdata (recdt, unicar ,CarSer ,plate ,orgmaterial1, fomula, TotalCube, car, field, strength, particle, founder,  concrete1, concrete2, concrete3, concrete4, concrete5, concrete6, nowaterMaterial1,"
        '106.1.2
        'strsql += "SaveDT) values ("
        strsql += "SaveDT,queuekey) values ("
        strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'"
        iCarSer += 1
        '107.6.5
        'LabelCarSer.Text = iCarSer

        '106.1.2 106.1.3
        'LabelQkey.Text = Format(Now, "yyyyMMdd") & Format(Now, "HHmmss")
        LabelQkey.Text = ""

        'queueP.CarSer = LabelCarSer.Text
        '103.6.10
        '105.10.20 UniCar revise
        'iUniSer += 1
        'queueP.CarSer = (iUniSer).ToString
        'queueP.UniCar = "UN_" & (iUniSer).ToString

        '105.10.20 ref queuecode
        Dim db As New DbAccess(sYMdbName + sSavingYear + ".mdb")
        Dim dt As DataTable
        If sim_produce Then
            dt = db.GetDataTable("select  top 1 UniID from recdata_SIM order by UniID desc ")
        Else
            dt = db.GetDataTable("select  top 1 UniID from recdata order by UniID desc ")
        End If
        If dt.Rows.Count < 1 Then
            iUniSer = 1
        Else
            iUniSer = CInt(dt.Rows(0).Item("UniID").ToString) + 1
        End If
        queue0.UniCar = "UN_" & (iUniSer).ToString
        queue0.CarSer = iUniSer
        LabelUniCar.Text = queue0.UniCar
        queueP.UniCar = queue0.UniCar
        queueP.CarSer = queue0.CarSer
        '106.1.2 for call SendSQL
        sRemainUniCar = queueP.UniCar

        RemainDonePlate_temp_SQL = 1
        strsql &= queueP.UniCar & "'," & iUniSer & ", 1 ," & 1
        'strsql &= iCarSer & "'," & iCarSer & "," & iCarSer

        '103.3.7
        Dim f
        f = Val(TbxworkTri.Text)
        '103.6.7
        If queue0.showdata(1) = "" Then queue0.showdata(1) = "1"
        '106.1.3
        'If tbxworkcode.Text = "" Then tbxworkcode.Text = "123"
        If sTemp = "" Then
            sTemp = " "
        End If
        If Tbxworkcar.Text = "" Then Tbxworkcar.Text = "."
        If tbxworkfield.Text = "" Then tbxworkfield.Text = "."
        If queue0.showdata(1) = "" Then queue0.showdata(1) = "1"
        '106.1.3
        'strsql &= ",'" & tbxworkcode.Text & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "' , '0' , '0' , '0' "
        strsql &= ",'" & sTemp & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "' , '0' , '0' , '0' "
        'strsql &= ",'" & queue0.showdata(1) & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "'"
        'mc.SendSQL(queueP.CarSer, queueP.UniCar, RemainDonePlate_temp_SQL, sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb", remain_B_SQL, "recdata_SIM")
        SaveSetting("JS", "CBC800", "iCarSer", CStr(iCarSer))

        For i = 0 To 5
            strsql &= "," & DV(73 + i).ToString
        Next
        '105.6.13 NonAuto to [nowaterMaterial1] 
        strsql &= ", 1 "

        '106.1.2
        'strsql &= "," & sSavingDate & ")"
        '106.1.3
        'strsql &= "," & sSavingDate & "," & Format(Now, "yyyyMMdd") & Format(Now, "HHmmss") & ")"
        strsql &= "," & sSavingDate & ",' ' )"

        If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb") Then
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
        End If
        If Not sim_produce Then
            Try
                dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
                dbrec.ExecuteCmd(strsql)
            Catch ex As Exception
                MsgBox("A資料庫新增資料失敗")
            End Try
        End If

        ''紀錄配方與骨材資料到database()_recdt_B
        '103.3.7
        'strsql = "Insert into recdata_B (recdt, unicar,CarSer,particle, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"

        '103.6.7
        'strsql = "Insert into recdata_B (recdt, unicar, CarSer,orgmaterial1, fomula, TotalCube, car, field, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        strsql = "Insert into recdata_B (recdt, unicar ,CarSer ,plate ,orgmaterial1, fomula, TotalCube, car, field, strength, particle, founder,  concrete1, concrete2, concrete3, concrete4, concrete5, concrete6,"
        strsql += "SaveDT) values ("
        strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'"
        '103.6.7 
        'iCarSer += 1
        SaveSetting("JS", "CBC800", "iCarSer", CStr(iCarSer))
        '107.6.5
        'LabelCarSer.Text = iCarSer
        strsql &= queueP.UniCar & "'," & iUniSer & ", 1 ," & 1
        'strsql &= iCarSer & "'," & iCarSer & "," & iCarSer

        '103.3.7
        'strsql &= ",'" & queue0.showdata(1) & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "'"
        '106.1.3
        'strsql &= ",'" & tbxworkcode.Text & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "' , '0' , '0' , '0' "
        strsql &= ",'" & sTemp & "'," & f & ",'" & Tbxworkcar.Text & "','" & tbxworkfield.Text & "' , '0' , '0' , '0' "


        For i = 0 To 5
            '103.3.7
            'strsql &= "," & DV(73 + i).ToString
            strsql &= "," & DV(103 + i).ToString
        Next
        strsql &= "," & sSavingDate & ")"


        '105.7.18 check if mdb exist 105.10.19
        'If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear_ + ".mdb") Then
        '    Return False
        'End If
        '105.10.19
        'If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb") Then
        '    My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
        '    dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
        'End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear + ".mdb") Then
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B_Local + sSavingYear + ".mdb")
        End If
        If Not sim_produce Then
            Try
                '105.10.19
                dbrec.changedb(sYMdbName_B_Local + sSavingYear + ".mdb")
                dbrec.ExecuteCmd(strsql)
            Catch ex As Exception
                '110.11.16
                'MsgBox("B資料庫新增資料失敗")
                MsgBox("B資料庫新增資料失敗:" & ex.Message)
            End Try
            '105.10.19
            'Dim b
            'b = CheckRemoteBfile(False)
        End If

        '103.6.9
        Call FrmHistory.PrintCarDetail(queueP.UniCar, queueP.CarSer, "recdata")
        '106.1.2    107.3.17
        'Call FrmMonit.TraceLog("SaveNonAuto()" & Now.ToString & " queueP.UniCar:" & queueP.UniCar)
        MonLog.add = "SaveNonAuto()" & Now.ToString & " queueP.UniCar:" & queueP.UniCar
        '107.2.11 'prevent TimeSQL delay queueP.UniCar maybe updated 
        sRemainUniCar = queueP.UniCar
        bNonAuto = True

    End Sub

    Public Sub SaveBone_New(ByVal BoneDonePlate_temp As Integer, ByVal cube As Single)
        '106.5.22
        '骨材完成盤數增加 不存檔 ! 
        '   only fill D61~68 修正印出值 , D91~98 印出真值 read already
        '=> 泥料完成盤數增加 才表示 本盤完成 才有紀錄 
        '
        'old:骨材完成盤數增加 存檔
        'save for all SP and water , So, SaveMod() only  not save SP again! PV will read & save untill D22 2010.0519 
        Dim i As Integer
        Dim d(10) As Integer

        '107.6.2
        Label_PLC_D20.add = "C0"

        If dgvPan.RowCount = 0 Then
            dgvPan.RowCount = 3
        End If

        '102.7.22 check length of strength
        'showdata : (0)-補正區域,(1)-B配比, (2)-總重, (3)-拌合時間, (4)-坍度, (5)-強度, (6)-粒徑, (7~9)-備註一~三
        '109.12.19 配比15碼 bCheckBoXSP17
        'If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 12 Then
        '    tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 12)
        'End If
        If bCheckBoxSP17 Then
            If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 15 Then
                tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 15)
            End If
        Else
            If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 12 Then
                tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 12)
            End If
        End If
        '(4)-坍度 FOUNDER integer
        If Microsoft.VisualBasic.Len(queue0.showdata(4)) > 4 Then
            queue0.showdata(4) = Microsoft.VisualBasic.Left(queue0.showdata(4), 4)
        End If
        '105.5.20
        'If Microsoft.VisualBasic.Len(queue0.showdata(5)) > 6 Then
        '    queue0.showdata(5) = Microsoft.VisualBasic.Left(queue0.showdata(5), 6)
        'End If
        '5)-強度 STRENTH string
        If Microsoft.VisualBasic.Len(queue0.showdata(5)) > 12 Then
            queue0.showdata(5) = Microsoft.VisualBasic.Left(queue0.showdata(5), 12)
        End If
        '(6)-粒徑 particle integer
        If Microsoft.VisualBasic.Len(queue0.showdata(6)) > 4 Then
            queue0.showdata(6) = Microsoft.VisualBasic.Left(queue0.showdata(6), 4)
        End If

        '102.5.28 洗車不記入資料庫
        bWashCar = True
        For i = 0 To 7
            If queue0.Fmaterialorg(i) > 0 Then
                bWashCar = False
            End If
            If queue0.Bmaterialorg(i) > 0 Then
                bWashCar = False
            End If
        Next
        queue0.isWashCar = bWashCar

        '106.9.9
        If module1.authority = 4 Then
            Label1SQL.Text = "BonC:"
            For i = 0 To 7
                Label1SQL.Text &= queue0.FmaterialReal(i) & ","
                Label1SQL.Text &= queue0.BmaterialReal(i) & "_"
            Next
            Label1SQL.Visible = True
            Label1SQL.Refresh()
        End If

        Dim k As Integer
        '106.9.14 copy water% 
        For k = 0 To 7
            queue0.waterall(k) = dgvWater.Rows(0).Cells(k).Value
            '107.4.28
            queue0.water_all_W_Limit(k) = water_all_W_Limit(k)
        Next
        '106.5.27 5.30  骨材完成 move to queueP no save to mdb 
        '106.9.3 should be copy queueP first 105.7.18 done when DV21 > 0
        '110.11.16
        'mc.queclone(queue0, queueP)
        ''110.11.12 @home 110.11.16
        ''queueP.TotalPlate = queue0.TotalPlate
        'Call UpdateLabelTotPlate()
        ''110.11.16 
        'If queue0.BoneDonePlate >= 1 Then
        '    mc.queclone(queue0, queueP)
        '    Call UpdateLabelTotPlate()
        'End If
        mc.queclone(queue0, queue)
        For k = 0 To 7
            waterall(k) = dgvWater.Rows(0).Cells(k).Value
            '106.9.14
            waterall(k) = dgvWater.Rows(0).Cells(k).Value
        Next

        If BoneDonePlate_temp < queue0.plate.Length Then
            queue0.BoneDoingPlate = BoneDonePlate_temp + 1
            mc.quecloneBone(queue0, queue)
            '106.9.3 106.9.4 10.3.17 
            'Call FrmMonit.TraceLog("SaveBone_New queue0(S1) F:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " R:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " , BoneDonePlate_temp:" & BoneDonePlate_temp)
            queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
            mc.queclone(queue, queue0)
            '107.6.28
            MonLog.add = "displaymaterial() : SaveBone_New"
            '107.6.26 bug can not disable since update
            '107.6.3 disable    
            displaymaterial()
        End If
        '107.3.9
        '106.9.6
        'Call SavingLog("SaveBone_New Bon= " & BoneDonePlate_temp & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate & " , DV= " & DV(20) & DV(21) & DV(22))
        TLog.add = "    SaveBone(),BM=" & BoneDonePlate_temp & queue0.ModDonePlate & ",DV20=" & D_20.PV & D_21.PV & D_22.PV & ",0/P:" & queue0.UniCar & "/" & queueP.UniCar & ",(S1) F:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " R:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " , BoneDonePlate_temp:" & BoneDonePlate_temp
        MonLog.add = "    SaveBone(),BM=" & BoneDonePlate_temp & queue0.ModDonePlate & ",DV20=" & D_20.PV & D_21.PV & D_22.PV & ",0/P:" & queue0.UniCar & "/" & queueP.UniCar & ",(S1) F:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " R:" & queue0.Fmaterialnowater(0) & "," & queue0.Fmaterialnowater(1) & " , BoneDonePlate_temp:" & BoneDonePlate_temp
    End Sub


    Public Sub SaveMod_New(ByVal ModDonePlate_temp As Integer, ByVal cube As Single)
        '106.5.27 ref. SaveBone old
        '泥料完成盤數增加 存檔
        ' 取代 骨材完成盤數增加 存檔 
        '106.9.6 BoneDonePlate_temp => ModDonePlate_temp

        '111.9.16   @js
        Dim SaveNow As Date
        SaveNow = Now

        '107.6.2
        Label_PLC_D21.add = "C0"
        '107.2.3
        BatchDone.StartTime = SaveNow.ToString("HH:mm:ss")


        Dim i As Integer
        Dim d(10) As Integer

        '107.3.14
        'Call FrmMonit.TraceLog("SaveMod_New()" & "ModDonePlate_temp:" & ModDonePlate_temp)

        If dgvPan.RowCount = 0 Then
            dgvPan.RowCount = 3
        End If

        BoneDoneChanging = False
        txtbone.ForeColor = Color.Black
        Dim ss As String

        'S1~G4
        For i = 0 To 7
            ss = Format(Val(queueP.Fmaterialorg(i)), "0")
            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                Fbone(i) = CSng(Format(Val(queueP.Fmaterialorg(i)), "0.0"))
                Fbone(i + 8) = CSng(Format(Val(queueP.Materialfront(i)), "0.0"))
                Fbone(i + 16) = CSng(Format(Val(queueP.Fmaterialnowater(i)), "0.0"))
                '106.9.14
                'Fbone(i + 24) = CSng(Format(Val(waterall(i)), "0.0"))
                '107.4.28
                'Fbone(i + 24) = CSng(Format(Val(queueP.waterall(i)), "0.0"))
                Fbone(i + 24) = CSng(Format(Val(queueP.water_all_W_Limit(i)), "0.0"))
                Bbone(i) = CSng(Format(Val(queueP.Bmaterialorg(i)), "0.0"))
                Bbone(i + 8) = CSng(Format(Val(queueP.Materialback(i)), "0.0"))
                Bbone(i + 16) = CSng(Format(Val(queueP.Bmaterialnowater(i)), "0.0"))
            Else
                Fbone(i) = CSng(Format(Val(queueP.Fmaterialorg(i)), "0"))
                Fbone(i + 8) = CSng(Format(Val(queueP.Materialfront(i)), "0"))
                Fbone(i + 16) = CSng(Format(Val(queueP.Fmaterialnowater(i)), "0"))
                '106.9.14
                'Fbone(i + 24) = CSng(Format(Val(waterall(i)), "0.0"))
                '107.4.28
                'Fbone(i + 24) = CSng(Format(Val(queueP.waterall(i)), "0.0"))
                Fbone(i + 24) = CSng(Format(Val(queueP.water_all_W_Limit(i)), "0.0"))
                Bbone(i) = CSng(Format(Val(queueP.Bmaterialorg(i)), "0"))
                Bbone(i + 8) = CSng(Format(Val(queueP.Materialback(i)), "0"))
                Bbone(i + 16) = CSng(Format(Val(queueP.Bmaterialnowater(i)), "0"))
            End If
            '101.4.5
            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                Fbone(i) = CSng(Format(Val(queueP.Fmaterialorg(i)), "0.0"))
                Fbone(i + 8) = CSng(Format(Val(queueP.Materialfront(i)), "0.0"))
                Fbone(i + 16) = CSng(Format(Val(queueP.Fmaterialnowater(i)), "0.0"))
                '106.9.14
                'Fbone(i + 24) = CSng(Format(Val(waterall(i)), "0.0"))
                '107.4.28
                'Fbone(i + 24) = CSng(Format(Val(queueP.waterall(i)), "0.0"))
                Fbone(i + 24) = CSng(Format(Val(queueP.water_all_W_Limit(i)), "0.0"))
                Bbone(i) = CSng(Format(Val(queueP.Bmaterialorg(i)), "0.0"))
                Bbone(i + 8) = CSng(Format(Val(queueP.Materialback(i)), "0.0"))
                Bbone(i + 16) = CSng(Format(Val(queueP.Bmaterialnowater(i)), "0.0"))
            End If
            '100.6.8 Bbone(i + 24) = CSng(Format(Val(waterall(i)), "0"))
            '106.9.14
            'Bbone(i + 24) = CSng(Format(Val(waterall(i)), "0.0"))
            Bbone(i + 24) = CSng(Format(Val(queueP.waterall(i)), "0.0"))
        Next
        ' 尼材亦須紀錄暫存
        'W1,2,3,C1,2,3,4,5,6
        For i = 0 To 8
            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                FMod(i) = CSng(Format(Val(queueP.Fmaterialorg(i + 8)), "0.0"))
                FMod(i + 17) = CSng(Format(Val(queueP.Materialfront(i + 8)), "0.0"))
                FMod(i + 34) = CSng(Format(Val(queueP.Fmaterialnowater(i + 8)), "0.0"))
                BMod(i) = CSng(Format(Val(queueP.Bmaterialorg(i + 8)), "0.0"))
                BMod(i + 17) = CSng(Format(Val(queueP.Materialback(i + 8)), "0.0"))
                BMod(i + 34) = CSng(Format(Val(queueP.Bmaterialnowater(i + 8)), "0.0"))
            Else
                FMod(i) = CSng(Format(Val(queueP.Fmaterialorg(i + 8)), "0"))
                FMod(i + 17) = CSng(Format(Val(queueP.Materialfront(i + 8)), "0"))
                FMod(i + 34) = CSng(Format(Val(queueP.Fmaterialnowater(i + 8)), "0"))
                BMod(i) = CSng(Format(Val(queueP.Bmaterialorg(i + 8)), "0"))
                BMod(i + 17) = CSng(Format(Val(queueP.Materialback(i + 8)), "0"))
                BMod(i + 34) = CSng(Format(Val(queueP.Bmaterialnowater(i + 8)), "0"))
            End If
            '101.4.5
            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                FMod(i) = CSng(Format(Val(queueP.Fmaterialorg(i + 8)), "0.0"))
                FMod(i + 17) = CSng(Format(Val(queueP.Materialfront(i + 8)), "0.0"))
                FMod(i + 34) = CSng(Format(Val(queueP.Fmaterialnowater(i + 8)), "0.0"))
                BMod(i) = CSng(Format(Val(queueP.Bmaterialorg(i + 8)), "0.0"))
                BMod(i + 17) = CSng(Format(Val(queueP.Materialback(i + 8)), "0.0"))
                BMod(i + 34) = CSng(Format(Val(queueP.Bmaterialnowater(i + 8)), "0.0"))
            End If
        Next
        'A1~
        For i = 9 To 16
            FMod(i) = queueP.Fmaterialorg(i + 8)
            FMod(i + 17) = queueP.Materialfront(i + 8)
            FMod(i + 34) = queueP.Fmaterialnowater(i + 8)
            BMod(i) = queueP.Bmaterialorg(i + 8)
            BMod(i + 17) = queueP.Materialback(i + 8)
            BMod(i + 34) = queueP.Bmaterialnowater(i + 8)
        Next

        '紀錄配方(S1~A5)與骨材資料(S1~G4)到database()_recdt

        '泥料變化 存檔 , so queueP.CarSer, queueP.UniCar update here
        If ModDonePlate_temp = 1 Then
            '107.6.2
            Label_PLC_D21.add = "C1"
            If sim_produce Then
                iCarSer_SIM += 1
                SaveSetting("JS", "CBC800", "iCarSer_SIM", CStr(iCarSer_SIM))
                '107.6.3
                'LabelCarSer.Text = iCarSer_SIM
                '106.3.12
                '106.11.15 SIM bug
                'queueP.UniCar = "UN_" & (iCarSer_SIM).ToString
                queue0.UniCar = "UN_" & (iCarSer_SIM).ToString
                queueP.UniCar = "UN_" & (iCarSer_SIM).ToString
            Else
                iCarSer += 1
                SaveSetting("JS", "CBC800", "iCarSer", CStr(iCarSer))
                '107.6.3
                'LabelCarSer.Text = iCarSer
                iCarSerToday += 1
                LabelCarSer.Text = iCarSerToday
            End If

        End If

        '102.7.22 check length of strength
        'showdata : (0)-補正區域,(1)-B配比, (2)-總重, (3)-拌合時間, (4)-坍度, (5)-強度, (6)-粒徑, (7~9)-備註一~三
        '                                                         varchar(4)  varchar(6) varchar(4) 
        '109.12.19 配比15碼 bCheckBoXSP17
        'If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 12 Then
        '    tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 12)
        'End If
        If bCheckBoxSP17 Then
            If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 15 Then
                tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 15)
            End If
        Else
            If Microsoft.VisualBasic.Len(tbxworkcode.Text) > 12 Then
                tbxworkcode.Text = Microsoft.VisualBasic.Left(tbxworkcode.Text, 12)
            End If
        End If
        '113.8.13
        Dim temp As String = ""
        If Microsoft.VisualBasic.Len(queueP.showdata(4)) > 4 Then
            queueP.showdata(4) = Microsoft.VisualBasic.Left(queueP.showdata(4), 4)
            temp = queueP.showdata(4)
            Try
                If temp = "" Then
                    queueP.showdata(4) = "20"
                End If
            Catch ex As Exception

            End Try
        End If
        If Microsoft.VisualBasic.Len(queueP.showdata(5)) > 12 Then
            queueP.showdata(5) = Microsoft.VisualBasic.Left(queueP.showdata(5), 12)
        End If
        If Microsoft.VisualBasic.Len(queueP.showdata(6)) > 4 Then
            queueP.showdata(6) = Microsoft.VisualBasic.Left(queueP.showdata(6), 4)
            temp = queueP.showdata(6)
            Try
                If temp = "" Then
                    queueP.showdata(6) = "20"
                End If
            Catch ex As Exception

            End Try
        End If

        If sim_produce Then
            strsql = "Insert into recdata_SIM (recdt, unicar, plate, fomula, totalcube, cube, car, field, mixtime, correctID, "
        Else
            strsql = "Insert into recdata (recdt, unicar, plate, fomula, totalcube, cube, car, field, mixtime, correctID, "
        End If
        strsql += "totalweight, founder, strength, particle, spare, CarSer, "
        strsql += "orgsand1, orgsand2, orgsand3, orgsand4, orgstone1, orgstone2, orgstone3, orgstone4, "
        strsql += "setsand1, setsand2, setsand3, setsand4, setstone1, setstone2, setstone3, setstone4,"
        strsql += "nowatersand1, nowatersand2, nowatersand3, nowatersand4, nowaterstone1, nowaterstone2, nowaterstone3, nowaterstone4,"
        strsql += "Watercompsand1, Watercompsand2, Watercompsand3, Watercompsand4, Watercompstone1, Watercompstone2, Watercompstone3, Watercompstone4, "
        strsql += "orgwater1, orgwater2, orgwater3, orgconcrete1, orgconcrete2, orgconcrete3, orgconcrete4, orgconcrete5, orgconcrete6, "
        strsql += "orgdrog1,  orgdrog2, orgdrog3, orgdrog4, orgdrog5, orgmaterial1, orgmaterial2, orgmaterial3, "
        strsql += "setwater1, setwater2, setwater3, setconcrete1, setconcrete2, setconcrete3, setconcrete4, setconcrete5, setconcrete6, "
        strsql += "setdrog1,  setdrog2, setdrog3, setdrog4, setdrog5, setmaterial1, setmaterial2, setmaterial3,  "
        strsql += "nowaterwater1, nowaterwater2, nowaterwater3, nowaterconcrete1, nowaterconcrete2, nowaterconcrete3, nowaterconcrete4, nowaterconcrete5, nowaterconcrete6, "
        strsql += "nowaterdrog1,  nowaterdrog2, nowaterdrog3, nowaterdrog4, nowaterdrog5, nowatermaterial1, nowatermaterial2, nowatermaterial3,  "
        strsql += "sand1, sand2, sand3, sand4, stone1, stone2, stone3, stone4, "
        '106.5.27
        strsql += "water1, water2, water3, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6, "
        strsql += "drog1,  drog2, drog3, drog4, drog5, material1, material2, material3,  "
        '107.7.24 bCheckBoxSP4 資料自動補傳時訊 save queuekey => save queuekey anyway
        'If bCheckBoxSP4 Then
        '    strsql += "queuekey, "
        'End If
        strsql += "queuekey, "

        strsql += "memo1, memo2, memo3, SaveDT) values ("
        '106.10.7 聯佶空單
        'If sim_produce Then
        '    strsql &= "#" & Format(dtSim, "yyyy/MM/dd ") & dtSim.Hour.ToString & ":" & dtSim.Minute.ToString & ":" & dtSim.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        'Else
        '    strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        'End If
        If sim_produce Then
            strsql &= "#" & Format(dtSim, "yyyy/MM/dd ") & dtSim.Hour.ToString & ":" & dtSim.Minute.ToString & ":" & dtSim.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        Else
            '106.10.20
            'If queueP.MixYn = "Y" Or queueP.MixYn = "y" Then
            If queueP.MixYn = "N" Or queueP.MixYn = "n" Then
                '107.12.28 first pan
                'strsql &= "#" & Format(dtSimP(ModDonePlate_temp), "yyyy/MM/dd ") & dtSimP(ModDonePlate_temp).Hour.ToString & ":" & dtSimP(ModDonePlate_temp).Minute.ToString & ":" & dtSimP(ModDonePlate_temp).Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
                ''106.10.11 聯佶空單 107.12.28 bug?
                'sSavingDate = "#" & Format(dtSimP(1), "yyyy/MM/dd ") & "#"
                strsql &= "#" & Format(dtSimP(0), "yyyy/MM/dd ") & dtSimP(0).Hour.ToString & ":" & dtSimP(0).Minute.ToString & ":" & dtSimP(0).Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
                '106.10.11 聯佶空單 107.12.28 bug?
                sSavingDate = "#" & Format(dtSimP(0), "yyyy/MM/dd ") & "#"
            Else
                '111.9.16
                'strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
                strsql &= "#" & Format(SaveNow, "yyyy/MM/dd ") & SaveNow.Hour.ToString & ":" & SaveNow.Minute.ToString & ":" & SaveNow.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & tbxworkcode.Text & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
            End If
        End If
        '107.6.27  spare(B) no save just ""
        'strsql &= Tbxworkcar.Text & "','" & tbxworkfield.Text & "'," & queueP.showdata(3) & ",'" & queueP.showdata(0) & "'," & queueP.showdata(2) & "," & queueP.showdata(4) & ",'" & queueP.showdata(5) & "'," & queueP.showdata(6) & ",'" & queueP.showdata(1) & "',"
        strsql &= Tbxworkcar.Text & "','" & tbxworkfield.Text & "'," & queueP.showdata(3) & ",'" & queueP.showdata(0) & "'," & queueP.showdata(2) & "," & queueP.showdata(4) & ",'" & queueP.showdata(5) & "'," & queueP.showdata(6) & ",'" & "" & "',"


        If sim_produce Then
            strsql &= iCarSer_SIM & ","
            sSavingDate = "#" & Format(dtp_simtime.Value, "yyyy/MM/dd ") & "#"
        Else
            strsql &= iCarSer & ","
        End If
        For i = 0 To 31
            strsql &= Fbone(i).ToString + ","
        Next
        FMod(14) = 0
        FMod(48) = 0
        ' 111.11.25 test bug
        For i = 0 To 50
            strsql &= FMod(i).ToString + ","
        Next


        '106.9.9 9.10
        If module1.authority = 4 Then
            Label1SQL.Text = "ModCh:"
            For i = 8 To 14
                'Label1SQL.Text &= queue0.BmaterialReal(i) & " D" & i + 92 & ":" & DV(i + 92) & ","
                Label1SQL.Text &= queueP.FmaterialReal(i) & ","
                Label1SQL.Text &= queueP.BmaterialReal(i) & "_"
            Next
            Label1SQL.Visible = True
            Label1SQL.Refresh()
        End If

        ' 111.11.25 test bug
        For i = 0 To 7
            strsql &= queueP.FmaterialReal(i).ToString & ","
        Next

        For i = 8 To 24
            strsql &= queueP.FmaterialReal(i).ToString & ","
        Next

        '107.7.24 bCheckBoxSP4 資料自動補傳時訊 save queuekey => save queuekey anyway 
        '107.8.21 forgot to revise !!!???
        'If bCheckBoxSP4 Then
        '    strsql &= "'" & LabelQkey.Text & "',"
        'End If
        strsql &= "'" & LabelQkey.Text & "',"

        'use memo3 as total plate
        queueP.showdata(9) = queueP.plate.Length
        strsql &= "'" & queueP.showdata(7) + "','" & queueP.showdata(8) & "','" & queueP.showdata(9) & "'," & sSavingDate & ")"
        Dim ti As Double
        ti = Microsoft.VisualBasic.DateAndTime.Timer
        '107.6.7
        'LabelMsg.Text = "資料存檔..."
        Dim s As String
        s = "資料存檔..."

        '100.6.11 sim
        If sim_produce Then
            If Not My.Computer.FileSystem.FileExists(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb") Then
                My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
            End If
            dbrec.changedb(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
            '107.4.14
            'MainLog.add = "SaveBone DB= " & sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb"
            TLog.add = "SaveMod DB= " & sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb"
        Else
            If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb") Then
                My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
            End If
            dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
            '107.6.7
            s = Microsoft.VisualBasic.Left(sYMdbName, Len(sYMdbName) - 2)
            SaveSetting("JS", "CBC800", "sMdbDir", s)
            s = sYMdbName + sSavingYear + ".mdb"
            SaveSetting("JS", "CBC800", "sMdbFileName", s)
            '107.4.14
            TLog.add = "SaveMod DB= " & sYMdbName + sSavingYear & ".mdb"
        End If

        '102.5.28 洗車不記入資料庫 110.11.19 no save bug
        Dim bWashCar As Boolean
        bWashCar = True
        For i = 0 To 7
            If queueP.Fmaterialorg(i) > 0 Then
                bWashCar = False
            End If
            If queueP.Bmaterialorg(i) > 0 Then
                bWashCar = False
            End If
        Next
        '112.12.26
        '112.12.6 固化廠
        Dim existFlag_0 As Boolean = True 'all no used =true
        For i = 0 To 7
            If exist(i) Then existFlag_0 = False ' check
        Next
        If existFlag_0 Then
            bWashCar = False
        End If

        queueP.isWashCar = bWashCar

        If Not queueP.isWashCar Then
            '107.6.2
            Label_PLC_D21.add = "C2"
            Try
                '107.4.14
                Background1_Sql = strsql
                sBackground3_UniCar = queueP.UniCar
                iBackground3_ModDone = ModDonePlate_temp
                fBackground3_cube = cube
                Me.BackgroundWorker1.RunWorkerAsync()
                'dbrec.ExecuteCmd(strsql)
            Catch ex As Exception
                MsgBox("A資料庫新增資料失敗")
            End Try
        End If
        '105.7.24 106.5.30 107.4.14 !!!
        'Dim db_sum
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where unicar = '" + queueP.UniCar + "'")
        'txt_carsum.Text = db_sum.Rows(0).Item(0).ToString
        'Dim m3 As Single
        'm3 = Val(txt_carsum.Text)
        'txt_carsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
        'Dim sum
        ''107.3.6
        ''填入日完成方數
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') =  '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "/" + CDate(sSavingDate).Day.ToString + "'")
        'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")

        ''填入月完成方數
        'db_sum = dbrec.GetDataTable("select DISTINCT  RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m') = '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "'")
        'sum = 0
        'For i = 0 To db_sum.Rows.Count - 1
        '    sum += Format(Val(db_sum.Rows(i).Item(1).ToString), "0.00")
        'Next
        'txt_monthsum.Text = Format(sum, "0.00")
        'lblMachTot.Text = Format(Val(lblMachTot.Text) + CSng(queueP.plate(ModDonePlate_temp - 1)), "0.00")
        'SaveSetting("JS", "CBC800", "lblMachTot", lblMachTot.Text)



        '紀錄配方與骨材資料到database()_recdt_B
        strsql = "Insert into recdata_B (recdt, unicar, plate, fomula, totalcube, cube, car, field, mixtime, correctID, "
        strsql &= "totalweight, founder, strength, particle, spare, CarSer, "
        'strsql &= "totalweight, founder, strength, particle, spare, "
        strsql &= "orgsand1, orgsand2, orgsand3, orgsand4, orgstone1, orgstone2, orgstone3, orgstone4, "
        strsql &= "setsand1, setsand2, setsand3, setsand4, setstone1, setstone2, setstone3, setstone4,"
        strsql &= "nowatersand1, nowatersand2, nowatersand3, nowatersand4, nowaterstone1, nowaterstone2, nowaterstone3, nowaterstone4,"
        strsql &= "Watercompsand1, Watercompsand2, Watercompsand3, Watercompsand4, Watercompstone1, Watercompstone2, Watercompstone3, Watercompstone4, "
        strsql &= "orgwater1, orgwater2, orgwater3, orgconcrete1, orgconcrete2, orgconcrete3, orgconcrete4, orgconcrete5, orgconcrete6, "
        strsql &= "orgdrog1,  orgdrog2, orgdrog3, orgdrog4, orgdrog5, orgmaterial1, orgmaterial2, orgmaterial3, "
        strsql &= "setwater1, setwater2, setwater3, setconcrete1, setconcrete2, setconcrete3, setconcrete4, setconcrete5, setconcrete6, "
        strsql &= "setdrog1,  setdrog2, setdrog3, setdrog4, setdrog5, setmaterial1, setmaterial2, setmaterial3,  "
        strsql &= "nowaterwater1, nowaterwater2, nowaterwater3, nowaterconcrete1, nowaterconcrete2, nowaterconcrete3, nowaterconcrete4, nowaterconcrete5, nowaterconcrete6, "
        strsql &= "nowaterdrog1,  nowaterdrog2, nowaterdrog3, nowaterdrog4, nowaterdrog5, nowatermaterial1, nowatermaterial2, nowatermaterial3,  "
        strsql &= "sand1, sand2, sand3, sand4, stone1, stone2, stone3, stone4, "
        '106.5.27
        strsql += "water1, water2, water3, concrete1, concrete2, concrete3, concrete4, concrete5, concrete6, "
        strsql += "drog1,  drog2, drog3, drog4, drog5, material1, material2, material3,  "
        '111.11.25 queuekey for B
        strsql += "queuekey, "
        strsql += "memo1, memo2, memo3, SaveDT) values ("
        If queueP.isBsub And bCheckBoxBPBsuben Then
            cube = cube - fBPBsub
            If cube < 0 Then cube = 0
        End If
        '106.10.7 聯佶空單
        'strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & queueP.showdata(1) & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        '106.10.20
        'If queueP.MixYn = "Y" Or queueP.MixYn = "y" Then
        If queueP.MixYn = "N" Or queueP.MixYn = "n" Then
            strsql &= "#" & Format(dtSimP(ModDonePlate_temp), "yyyy/MM/dd ") & dtSimP(ModDonePlate_temp).Hour.ToString & ":" & dtSimP(ModDonePlate_temp).Minute.ToString & ":" & dtSimP(ModDonePlate_temp).Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & queueP.showdata(1) & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        Else
            '111.9.13 ?? Now?
            'strsql &= "#" & Format(Now, "yyyy/MM/dd ") & Now.Hour.ToString & ":" & Now.Minute.ToString & ":" & Now.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & queueP.showdata(1) & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
            strsql &= "#" & Format(SaveNow, "yyyy/MM/dd ") & SaveNow.Hour.ToString & ":" & SaveNow.Minute.ToString & ":" & SaveNow.Second.ToString & "#,'" & queueP.UniCar & "', " & ModDonePlate_temp & ",'" & queueP.showdata(1) & "'," & TbxworkTri.Text & "," & cube.ToString & ",'"
        End If
        '107.6.27 spare B no save
        'strsql &= Tbxworkcar.Text & "','" & tbxworkfield.Text & "'," & queueP.showdata(3) & ",'" & queueP.showdata(0) & "'," & queueP.showdata(2) & "," & queueP.showdata(4) & ",'" & queueP.showdata(5) & "'," & queueP.showdata(6) & ",'" & queueP.showdata(1) & "',"
        strsql &= Tbxworkcar.Text & "','" & tbxworkfield.Text & "'," & queueP.showdata(3) & ",'" & queueP.showdata(0) & "'," & queueP.showdata(2) & "," & queueP.showdata(4) & ",'" & queueP.showdata(5) & "'," & queueP.showdata(6) & ",'" & "" & "',"
        strsql &= iCarSer & ","
        For i = 0 To 31
            strsql &= Bbone(i).ToString & ","
        Next
        BMod(14) = 0
        BMod(48) = 0
        For i = 0 To 50
            strsql &= BMod(i).ToString & ","
        Next
        For i = 0 To 7
            strsql &= queueP.BmaterialReal(i).ToString & ","
        Next
        For i = 8 To 24
            strsql &= queueP.BmaterialReal(i).ToString & ","
        Next

        '111.11.25 queuekey for B
        strsql &= "'" & LabelQkey.Text & "',"

        strsql &= "'" & queueP.showdata(7) + "','" & queueP.showdata(8) & "','" & queueP.showdata(9) & "'," & sSavingDate & ")"

        '105.7.3 save B file local
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear + ".mdb") Then
            '105.7.3
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
        End If
        'dbrec.changedb(sYMdbName_B + sSavingYear + ".mdb")
        '105.7.14 save B file to local disk first
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear + ".mdb") Then
            '105.7.3
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B_Local + sSavingYear + ".mdb")
        End If
        dbrec.changedb(sYMdbName_B_Local + sSavingYear + ".mdb")

        '107.12.28 bCheckBoxSP3 X配存檔
        'If Not queueP.isWashCar Then
        'If Not queueP.isWashCar And bCheckBoxSP3 Then 108.1.10 => just no ExecuteCmd 
        If Not queueP.isWashCar Then
            If Not sim_produce Then
                Try
                    '107.4.14
                    Background2_Sql = strsql
                    Me.BackgroundWorker2.RunWorkerAsync()
                    'dbrec.ExecuteCmd(strsql)
                Catch ex As Exception
                    '110.11.16 
                    'MsgBox("B資料庫新增資料失敗") 
                    MsgBox("B資料庫新增資料失敗SaveMod_New:" & ex.Message)
                End Try
            End If
        End If

        '106.9.8  泥料完成 move to queueR for remain when DV21 change
        '110.11.12 110.11.14 @home 110.11.16
        'If queue0.ModDonePlate >= queueP.plate.Length Then
        '    queueP.LastPlate = True
        '    queueR.TotalPlate = queueP.TotalPlate
        'End If
        mc.queclone(queueP, queueR)
        Call UpdateLabelTotPlate()

        mc.queclone(queue0, queue)
        Dim k As Integer
        For k = 0 To 7
            waterall(k) = dgvWater.Rows(0).Cells(k).Value
        Next

        If ModDonePlate_temp < queue0.plate.Length Then
            '107.6.2
            Label_PLC_D21.add = "C3"
            '107.6.2 !!! ???
            queue0.BoneDoingPlate = ModDonePlate_temp + 1
            mc.quecloneBone(queue0, queue)
            queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
            mc.queclone(queue, queue0)

            '100.12.6   107.3.17 107.4.14
            'Call FrmMonit.TraceLog("displaymaterial() : SaveMod_New")
            '107.6.3 disable
            'displaymaterial()
            'TLog.add = "displaymaterial() : SaveMod_New"
            'MonLog.add = "displaymaterial() : SaveMod_New"
        End If

        '100.4.23
        '106.9.8 LabelMsg.Text &= "..完成 : " & (Microsoft.VisualBasic.DateAndTime.Timer - ti)
        s &= "..完成 : " & (Microsoft.VisualBasic.DateAndTime.Timer - ti)
        '100.5.11Call SavingLog("SaveBone Bon= " & ModDonePlate_temp & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate)
        '107.3.9
        'Call SavingLog("SaveMod_New Bon= " & queue0.BoneDonePlate & " , Mod= " & ModDonePlate_temp & " , Rem= " & queue0.RemainDonePlate & " , DV= " & DV(20) & DV(21) & DV(22) & s)
        TLog.add = "    SaveMod():BM=" & queue0.BoneDonePlate & ModDonePlate_temp & ",DV20=" & D_20.PV & D_21.PV & D_22.PV & ",0/P:" & queue0.UniCar & "/" & queueP.UniCar
    End Sub

   
    Public Sub SaveRemain_New(ByVal RemainDonePlate_temp As Integer)
        '106.5.27 copy from SaveRemain()

        Dim remain_B(25) As Single
        'Dim db_sum As DataTable
        Dim sLog As String

        sLog = ""
        '107.6.2
        Label_PLC_D22.add = "C0"

        '106.9.8 queueP replaced by queueR in this block when parssing
        '107.3.18
        'Call SavingLog("SaveRemain_New()" & "RemainDonePlate_temp:" & RemainDonePlate_temp & " bWashCar:" & bWashCar)

        '*****************讀PLC D61-85, D91-115資料記錄到database  only D22 2010.0517
        '109.11.29
        'Label1.Top = Panel2.Top - 20
        'Label1.Text = ""
        Dim ti As Double
        ' AE
        For i = 0 To 4
            queueR.FmaterialReal(i + 8 + 9) = DV(79 + i) * 0.01
        Next

        Txtmod.ForeColor = Color.Black
        ti = Microsoft.VisualBasic.DateAndTime.Timer

        For i = 0 To 23
            If i <= 7 Then
                If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Then
                    queueR.materialRemain(i) = CStr(DV(161 + i) * 0.1) 'D161
                    remain_B(i) = DV(161 + i) * 0.1 'D161
                Else
                    queueR.materialRemain(i) = CStr(DV(161 + i)) 'D161
                    remain_B(i) = DV(161 + i) 'D161
                End If
                '101.4.5
                If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                    queueR.materialRemain(i) = CStr(DV(161 + i) * 0.1) 'D161
                    remain_B(i) = DV(161 + i) * 0.1 'D161
                End If
            ElseIf i < 17 Then
                If (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                    queueR.materialRemain(i) = CStr(DV(161 + i + 1) * 0.1) 'D161
                    remain_B(i) = DV(161 + i + 1) * 0.1 'D161
                Else
                    queueR.materialRemain(i) = CStr(DV(161 + i + 1)) 'D161
                    remain_B(i) = DV(161 + i + 1) 'D161
                End If
            Else
                queueR.materialRemain(i) = CStr(DV(161 + i + 1)) * 0.01 'D161
                remain_B(i) = DV(161 + i + 1) * 0.01 'D161
            End If
            '106.9.9

        Next
        txtremain.ForeColor = Color.Black


        Dim arrdbmodreal() As String = {"sand1", "sand2", "sand3", "sand4", "stone1", "stone2", "stone3", "stone4", _
        "water1", "water2", "water3", "concrete1", "concrete2", "concrete3", "concrete4", "concrete5", _
        "concrete6", "drog1", "drog2", "drog3", "drog4", "drog5", "material1", "material2", "material3"}

        Dim arrdbremain() As String = {"remainsand1", "remainsand2", "remainsand3", "remainsand4", "remainstone1", "remainstone2", _
        "remainstone3", "remainstone4", "remainwater1", "remainwater2", "remainwater3", "remainconcrete1", "remainconcrete2", _
        "remainconcrete3", "remainconcrete4", "remainconcrete5", "remainconcrete6", "remaindrog1", "remaindrog2", "remaindrog3", _
        "remaindrog4", "remaindrog5", "remainmaterial1", "remainmaterial2", "remainmaterial3"}


        If RemainDonePlate_temp <= 0 Then Exit Sub

        '106.5.16 恐 D70~83 D100~113 讀取出錯 沒值
        Dim err As Boolean
        'D70 W1~AE
        err = True
        For i = 0 To 14
            If queueR.FmaterialReal(i + 8) > 0 Then err = False
        Next
        If err Then
            For i = 0 To 14
                queueR.FmaterialReal(i + 8) = queueR.Materialfront(i + 8)
            Next
        End If
        'D100
        err = True
        For i = 0 To 14
            If queueR.BmaterialReal(i + 8) > 0 Then err = False
        Next
        If err Then
            For i = 0 To 14
                queueR.BmaterialReal(i + 8) = queueR.Materialback(i + 8)
            Next
        End If

        '110.11.16 maybe bug
        Dim strsql As String

        strsql = " Update recdata_B set "

        For i = 0 To 23
            '106.9.9
            'If MatUse(i) = 0 Then
            '    queueR.materialRemain(i) = 0
            'End If
            If MatUse(i) = 0 Or queueR.Materialback(i) <= 0 Then
                queueR.materialRemain(i) = 0
                '112.4.28
                remain_B(i) = 0.0
            End If

            '104.9.24 materialRemain(24) for Current diff. 105.5.27 seems no use materialRemain(24)
            queueR.materialRemain(23) = EndCurrent

            '112.3.25 列印拌合結束電流 from LabelCurrent3 when D22 change

        Next

        Dim ftest
        ftest = queueR.materialRemain(23)

        str_cur_weight(1) = 0  ''每盤完成後實際重量 A (0):report, (1):real
        For i = 0 To 22
            If sProject = "Q04" Then
                str_cur_weight(1) += queueR.FmaterialReal(i) - queueR.materialRemain(i)
            Else
                str_cur_weight(1) += queueR.BmaterialReal(i) - queueR.materialRemain(i)
            End If
        Next

        '112.4.28
        For i = 0 To arrdbremain.Length - 3
            strsql += arrdbremain(i) + "=" + remain_B(i).ToString & ","
        Next


        '112.3.25 列印拌合結束電流 from LabelCurrent3 when D22 change
        queueR.materialRemain(23) = EndCurrent
        strsql &= arrdbremain(23) & "=" + queueR.materialRemain(23) & ","

        '111.10.22 for remain done RemainMaterial3 ??
        queueR.materialRemain(24) = 99
        strsql &= arrdbremain(24) & "=" + queueR.materialRemain(24)


        '102.3.16 列印實際拌合時間
        If bCheckBoxMix Then
            strsql += "," & " mixtime = " & DV(23)
        End If

        strsql &= " where unicar = '" & queueR.UniCar & "' and plate =" & RemainDonePlate_temp
        remain_B(24) = CType(queueR.materialRemain(24), Single)
        '105.8.13 'prevent TimeSQL delay queueP.UniCar maybe updated 
        sRemainUniCar = queueR.UniCar

        '107.12.28 bCheckBoxSP3 X配存檔
        'If Not bWashCar Then
        If Not bWashCar And bCheckBoxSP3 Then
            If Not sim_produce Then
                Try
                    '105.7.3 7.14
                    'dbrec.changedb(sYMdbName_B + sSavingYear + ".mdb")
                    dbrec.changedb(sYMdbName_B_Local + sSavingYear + ".mdb")
                    dbrec.ExecuteCmd(strsql)
                    '107.12.28 bCheckBoxSP15 B配也修正 for 巨力
                    If bCheckBoxSP15 Then
                        mc.count_B_report(queueR.CarSer, queueR.UniCar, RemainDonePlate_temp, sYMdbName_B_Local & sSavingYear & ".mdb", remain_B, "Recdata_B")
                    End If
                Catch ex As Exception
                    '110.11.16 ex.Msg:位置0沒有資料列
                    'MsgBox("B資料庫更新remain資料失敗")
                    MsgBox("B資料庫更新remain資料失敗:" & ex.Message & " SQL:" & strsql)
                    TLog.add = "B資料庫更新remain資料失敗: : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & strsql
                End Try
            End If
        End If

        '112.1.5    112.1.10
        '   use "queueR.CarSer, queueR.UniCar, RemainDonePlate_temp," 
        '   to get row to ReceiaHeader ReceiaPan
        '112.4.12
        'If bPingSQLBigData Then
        Dim strsqlall As String
        'strsql = "select * from recdata where unicar='" + unicar + "' and plate =" + plate.ToString & " order by RecDT DESC "
        strsqlall = "Select DISTINCT * from recdata_B WHERE unicar='" & queueR.UniCar & "' and plate =" & RemainDonePlate_temp.ToString & " order by RecDT DESC "
        Dim db_sum As New DataTable
        db_sum = dbrec.GetDataTable(strsqlall)

        '112.10.1 reset receiaPanB then add in fillMdbRowToPan()
        Dim kk
        If RemainDonePlate_temp = 1 Then
            For kk = 0 To receiaPanB.RT.valRec.Length - 1
                receiaPanB.RT.valRec(kk) = 0
            Next
            receiaPanB.RT.qty = 0
        End If

        Dim receiaPan As New classReceiaPan()
        Dim jj As Integer
        '112.1.28 112.3.25
        Dim send
        For jj = 0 To db_sum.Rows.Count - 1
            '112.10.1 receiaPan = fillMdbRowToPan(db_sum.Rows(jj), "A")
            receiaPan = fillMdbRowToPan(db_sum.Rows(jj), "B")
            If receiaPan.Header.sumno = receiaPan.Header.no Then
                '112.10.1 maybe mdb moved by remote => bug
                'strsqlall = "Select SUM(sand1),SUM(sand2),SUM(sand3),SUM(sand4),SUM(Stone1),SUM(Stone2),SUM(Stone3),SUM(Stone4),SUM(Water1),SUM(Water2),SUM(Water3)"
                'strsqlall &= ",SUM(Concrete1),SUM(Concrete2),SUM(Concrete3),SUM(Concrete4),SUM(Concrete5),SUM(Concrete6)"
                'strsqlall &= ",SUM(Drog1),SUM(Drog2),SUM(Drog3),SUM(Drog4),SUM(Drog5),SUM(cube) from recdata_B WHERE [UniCar] ='" & receiaPan.Header.UniCar & "'"
                'Dim db_Dt As New DataTable
                'db_Dt = dbrec.GetDataTable(strsqlall)

                ''112.4.19 bug:only sum no remain
                'strsqlall = "Select SUM(remainsand1),SUM(remainsand2),SUM(remainsand3),SUM(remainsand4),SUM(remainStone1),SUM(remainStone2),SUM(remainStone3),SUM(remainStone4),SUM(remainWater1),SUM(remainWater2),SUM(remainWater3)"
                'strsqlall &= ",SUM(remainConcrete1),SUM(remainConcrete2),SUM(remainConcrete3),SUM(remainConcrete4),SUM(remainConcrete5),SUM(remainConcrete6)"
                'strsqlall &= ",SUM(remainDrog1),SUM(remainDrog2),SUM(remainDrog3),SUM(remainDrog4),SUM(remainDrog5),SUM(cube) from recdata_B WHERE [UniCar] ='" & receiaPan.Header.UniCar & "'"
                'Dim db_Dt_rem As New DataTable
                'db_Dt_rem = dbrec.GetDataTable(strsqlall)
                'If db_Dt.Rows.Count > 0 Then
                '    '112.4.19
                '    'receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0))
                '    receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0), db_Dt_rem.Rows(0))
                '    receiaPan.DT.qty = db_Dt.Rows(0).Item(22)
                'End If

                '112.3.25
                '112.9.18 
                'send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "A")
                If bPingSQL And bDataBaseLink Then
                    send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "A")
                End If
                If bPingSQLBigData Then
                    send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", True), "B")
                End If
            Else
                '112.1.28
                '112.3.25
                '112.9.18 
                'send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", False), "A")
                If bPingSQL And bDataBaseLink Then
                    send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", False), "A")
                End If
                If bPingSQLBigData Then
                    send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "R", False), "B")
                End If
            End If
        Next
        'End If 112.4.12

        '107.8.16
        dbrec.dispose()

        '107.8.10 
        RemainDonePlate_MoveA = RemainDonePlate_temp
        UniCar_MoveA = queueR.UniCar
        TimerMoveA.Enabled = True

        'recheck the err% of real value after receive remain
        '計算A報表remain 資料 將資料重新填回資料庫
        If Not bWashCar Then
            '102.8.16 delay 出錯  將 SQL 與 save A配 拆開   
            If sim_produce Then
                mc.count_A_report_only(queueR.CarSer, queueR.UniCar, RemainDonePlate_temp, sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb", remain_B, "recdata_SIM")
            Else
                mc.count_A_report_only(queueR.CarSer, queueR.UniCar, RemainDonePlate_temp, sYMdbName + sSavingYear + ".mdb", remain_B, "recdata")
            End If

            RemainDonePlate_temp_SQL = RemainDonePlate_temp
            For i = 0 To 24
                remain_B_SQL(i) = remain_B(i)
            Next
            '103.6.7
            bOutway = "A"

            '112.10.7 no use TimerSQL
            'If (bDataBaseLink And bPingSQL) And bCommFlag Then
            '    TimerSQL.Enabled = True
            'End If


            '107.1.31   107.9.21 Tainei move from TimerSQL bCommFlag
            '109.6.12
            'If bCheckBox232 Then
            If bCheckBox232 And bCommFlag Then
                '108.2.27 move to Timer232
                'mc.Send_232(queueR.CarSer, sRemainUniCar, RemainDonePlate_temp_SQL, sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
                'TimerBatDone.Enabled = True
                'BatchDone.DoneStep = 1
                'If BatchDone.PanCt = 1 Then
                '    BatchDone.PanBegin = 1
                'Else
                '    BatchDone.PanBegin = 3
                'End If
                '108.2.27 move to Timer232
                Timer232.Enabled = True
                Exit Sub
            End If

            '107.8.31 
            '   add TimerMoveB only 環泥烏日bCheckBoxSP6 send remoteB
            If bCheckBoxSP6 And sTextBoxIP_R <> "127.0.0.1" Then
                TimerMoveB.Enabled = True
            End If

            '107.7.4
            'If ((bRemoteBpath)) Or (bOnlyLocalPC And bCheckBoxSP3) Then
            '    LabelMessage2.Visible = True
            '    LabelMessage2.Text = "資料搜尋中(TimerSQL_B)請稍候(Wait...) "
            '    Me.Refresh()
            '    Label4.Image = My.Resources._0013
            '    Label4.Refresh()
            '    b = MoveLocalToRemoteBfile(sRemainUniCar, RemainDonePlate_temp_SQL)
            '    Label4.Image = My.Resources._000
            '    LabelMessage2.Visible = False
            '    bLastPlateBySql = False
            'End If

        End If

        '100.4.23
        LabelMsg.Text &= "," & sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb" & "..殘留直存檔完成 : " & (Microsoft.VisualBasic.DateAndTime.Timer - ti)
        ti = Microsoft.VisualBasic.DateAndTime.Timer

        '106.5.30 by SaveMod_New
        ''填入每車完成方數
        ''99.08.07
        ''db_sum = dbrec.GetDataTable("select sum(cube) from recdata where unicar = '" + LabelUniCar.Text + "'")
        ''105.7.24
        'dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where unicar = '" + queueP.UniCar + "'")
        'txt_carsum.Text = db_sum.Rows(0).Item(0).ToString
        'Dim m3 As Single
        'm3 = Val(txt_carsum.Text)
        'txt_carsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")

        ''填入每日完成方數()
        'Dim s As String
        's = ssSavingDate
        'db_sum.Clear()
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")

        'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
        'SaveSetting("JS", "CBC800", "txt_daysum", txt_daysum.Text)

        '填入每月完成方數
        'db_sum.Clear()
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m') = '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
        ''txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")

        ''102.10.23
        'If lblMachTot.Visible = True Then
        '    txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
        'Else
        '    Month_report2(sYMdbName & Today.Year & "M" & Today.Month & ".mdb", "A", Today)
        'End If
        'SaveSetting("JS", "CBC800", "txt_monthsum", txt_monthsum.Text)

        '本機累計
        '99.12.09
        'lblMachTot.Text = Format(Val(lblMachTot.Text) + Val(queueP.plate(RemainDonePlate_temp - 1)), "0.00")
        'lblMachTot.Text = Format(Val(lblMachTot.Text) + CSng(queueP.plate(RemainDonePlate_temp - 1)), "0.00")
        'SaveSetting("JS", "CBC800", "lblMachTot", lblMachTot.Text)
        '110.11.16
        'cub = CSng(queueR.plate(RemainDonePlate_temp - 1))
        Try
            cub = CSng(queueR.plate(RemainDonePlate_temp - 1))
        Catch ex As Exception
            cub = 1
        End Try
        If cub = 0 Then cub = 1

        If chk_A_B.Checked Then
            If sim_addshow_weight Then
                'LabelRealTot1M3.Text = Convert.ToSingle(str_cur_weight(0)) + Convert.ToSingle(str_cur_weight(2))
            Else
                '110.11.16
                'cub = CSng(queueR.plate(RemainDonePlate_temp - 1))
                'If cub = 0 Then cub = 1
                LabelRealTot1M3.Text = Format(str_cur_weight(1) / cub, "0")
            End If
        Else
            '110.11.16
            'cub = CSng(queueR.plate(RemainDonePlate_temp - 1))
            'If cub = 0 Then cub = 1
            LabelRealTot1M3.Text = Format(str_cur_weight(0) / cub, "0")
        End If

        '***********開始產生報表
        reptype = 1

        '106.9.8
        'Call SavingLog("SaveRema_New() Bon= " & queueR.BoneDonePlate & " , Mod= " & queueR.ModDonePlate & " , Rem= " & RemainDonePlate_temp & " , DV= " & DV(20) & DV(21) & DV(22))
        ''106.4.5
        'sLog = queueR.UniCar & " plate =" & RemainDonePlate_temp & " : "
        'For i = 0 To 24
        '    sLog &= queueR.BmaterialReal(i) & ","
        'Next
        'Call TraceLog3(sLog)
        sLog = queueR.UniCar & " plate =" & RemainDonePlate_temp & " : "
        For i = 0 To 24
            sLog &= queueR.BmaterialReal(i) & ","
        Next
        '107.3.14
        'Call SavingLog("SaveRema():BM=" & queueR.BoneDonePlate & queueR.ModDonePlate & " , Rem= " & RemainDonePlate_temp & ",DV=" & DV(20) & DV(21) & DV(22) & "," & sLog)
        '107.3.9
        'Call SavingLog("SaveRema_New() Bon= " & queueR.BoneDonePlate & " , Mod= " & queueR.ModDonePlate & " , Rem= " & RemainDonePlate_temp & " , DV= " & DV(20) & DV(21) & DV(22) & " , " & sLog)
        TLog.add = "    SaveRema(),BMR=" & queueR.BoneDonePlate & queueR.ModDonePlate & RemainDonePlate_temp & ",DV20=" & D_20.PV & D_21.PV & D_22.PV & ",0/P:" & queue0.UniCar & "/" & queueP.UniCar
    End Sub

    

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.W And e.Modifiers = Keys.Alt Then
            DataGridView1.Focus()
            DataGridView1.Rows(0).Cells(0).Selected = True
        End If
        If e.KeyCode = Keys.E And e.Modifiers = Keys.Alt Then
            ErrFlag += 1
            If ErrFlag > 2 Then ErrFlag = 0
        End If
        If e.KeyCode = Keys.D And e.Modifiers = Keys.Alt Then
            '110.6.25
            'If LabelA1.ForeColor = Color.Cyan Then
            If LabelA11.ForeColor = Color.Cyan Then
                DVW(836) = 1
                '110.6.25 @JS
                'LabelA1.ForeColor = Color.White
                LabelA11.ForeColor = Color.White
            Else
                DVW(836) = 0
                '110.6.25 @JS
                'LabelA1.ForeColor = Color.Cyan
                LabelA11.ForeColor = Color.Cyan
            End If
            i = FrmMonit.WriteA2N(836, 1)
        End If
        If e.KeyCode = Keys.R And e.Modifiers = Keys.Alt Then
            '99.08.07
            If LabelA2.ForeColor = Color.Yellow Then
                DVW(837) = 1
                LabelA2.ForeColor = Color.White
            Else
                DVW(837) = 0
                LabelA2.ForeColor = Color.Yellow
            End If
            i = FrmMonit.WriteA2N(837, 1)
        End If
        '103.4.19
        'If e.KeyCode = Keys.L And e.Modifiers = Keys.Alt Then
        If e.KeyCode = Keys.L And e.Modifiers = Keys.Alt And module1.authority > 1 Then
            'tbxspare.Visible = chk_A_B.Checked
            'lblspare.Visible = chk_A_B.Checked
            'cub = CSng(queue0.plate(queue0.RemainDonePlate - 1))
            If cub = 0 Then cub = 1
            If chk_A_B.Checked Then
                If sim_addshow_weight Then
                    'LabelRealTot1M3.Text = Convert.ToSingle(str_cur_weight(0)) + Convert.ToSingle(str_cur_weight(2))
                Else
                    LabelRealTot1M3.Text = Format(str_cur_weight(0) / cub, "0")
                End If
                lblcode.ForeColor = Color.Black
                DVW(830) = 0
                i = FrmMonit.WriteA2N(830, 1)
                '99.08.02
                '100.5.15
                'cbxcorrect.Text = queue0.showdata(0)
                LabelArea.Visible = False
                '100.5.15
                cbxcorrect.Visible = False
                chk_A_B.Checked = False
                '104.9.26
                NumSG.Visible = False
                LabelSA.Visible = False
                '111.4.2
                NumAE.Visible = False
                LabelAE.Visible = False
            Else
                LabelRealTot1M3.Text = Format(str_cur_weight(1) / cub, "0")
                lblcode.ForeColor = Color.Violet
                DVW(830) = 1
                Application.DoEvents()
                i = FrmMonit.WriteA2N(830, 1)
                '99.08.02
                cbxcorrect.Text = queue0.corr_B
                LabelArea.Text = queue0.corr_B
                LabelArea.Visible = True
                '100.5.15
                cbxcorrect.Visible = True
                chk_A_B.Checked = True
                '104.9.26
                NumSG.Visible = True
                LabelSA.Visible = True
                '111.4.10
                '111.4.2
                'NumAE.Visible = True
                'LabelAE.Visible = True
                If bCheckBoxSP18 Then
                    NumAE.Visible = True
                    LabelAE.Visible = True
                End If
            End If
            MonLog.add = "displaymaterial() : Alt-L"
            displaymaterial()
        End If
        '103.3.15
        'If e.KeyCode = Keys.S And e.Modifiers = Keys.Alt Then
        '103.4.19 If e.KeyCode = Keys.S And e.Modifiers = Keys.Alt And Not bEngVer Then
        '103.4.22
        'if e.KeyCode = Keys.S And e.Modifiers = Keys.Alt And Not bEngVer And module1.authority > 1 Then
        If e.KeyCode = Keys.S And e.Modifiers = Keys.Alt And Not bImperial And module1.authority > 1 Then
            If queue0.inprocess = True Then
                PanelSim.Visible = Not PanelSim.Visible
                If PanelSim.Visible Then
                    simstatus(True)
                Else
                    simstatus(False)
                End If
            End If
        End If
        If e.KeyCode = Keys.Q And e.Modifiers = Keys.Alt Then
            Cbxquecode1.Focus()
            Cbxquecode1.SelectAll()
        End If
        If e.KeyCode = Keys.W And e.Modifiers = Keys.Alt Then
            NumW1.Focus()
        End If
        If e.KeyCode = Keys.Z And e.Modifiers = Keys.Alt Then
            Call btnbegin_Click(btnbegin, e)
        End If
        '102.12.18 102.2.25 disable
        'If e.KeyCode = Keys.X And e.Modifiers = Keys.Alt Then
        '    If bEngVer Then
        '        bImperial = Not bImperial
        '        If bImperial Then
        '            fImperial = KG_TO_LB
        '        Else
        '            fImperial = 1.0
        '        End If
        '        Call UpdateImperial()
        '    End If
        'End If
        '102.10.23 縮減月米數
        If bCheckBoxMsuben Then
            If e.KeyCode = Keys.M And e.Modifiers = Keys.Alt Then
                '填入每日完成方數()
                Dim db_sum As DataTable
                Dim s As String
                Dim day_t, mon_t As Single
                s = Today.ToShortDateString
                db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")
                'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
                day_t = Val(db_sum.Rows(0).Item(0).ToString)
                db_sum.Clear()
                db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m') = '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
                mon_t = Val(db_sum.Rows(0).Item(0).ToString)
                If lblMachTot.Visible = True Then
                    lblMachTot.Visible = False
                    Label3.Visible = False
                    ButtonＭ.Visible = True
                    '填入每月完成方數
                    'txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString) * fMsub + day_t * (1 - fMsub), "0.00")
                    'txt_monthsum.Text = Format((mon_t * fMsub + day_t * (1 - fMsub)), "0.00")
                    Month_report2(sYMdbName & Today.Year & "M" & Today.Month & ".mdb", "A", Today)
                Else
                    lblMachTot.Visible = True
                    Label3.Visible = True
                    ButtonＭ.Visible = False
                    '填入每月完成方數
                    txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
                End If
            End If
        End If

    End Sub

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dtconfig As DataTable
        Dim i = 0
        Dim j = 0
        Dim j1 = 0
        Dim j2, j3 As Integer
        Dim st As String
        Dim water_item As Integer

        '109.11.29
        Label1.Top = Panel2.Top - 20
        Label1.Text = "調度電腦IP: " & sDataBaseIP
        '110.9.11
        'Label1.Visible = True
        ' 配比連線   109.11.29
        If bPbSync Then
            Dim sorceFile As String
            Dim destFile As String
            sorceFile = "\JS\CBC8\Conf_" & sProject & "\Recipe.mdb"
            destFile = "\JS\CBC8\log\Recipe_" & Now.ToString("yyyyMMdd") & ".mdb"
            My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)
            objDownloadAerp.action(db_recipe)
            MsgBox("下載配比... 新增:" & objDownloadAerp.FlagInsert & " 筆, 更新:" & objDownloadAerp.FlagUpdate & " 筆, 刪除:" & objDownloadAerp.FlagDelet & " 筆.", MsgBoxStyle.Information)
        End If

        'Dim w As Integer
        'Dim g As Integer

        '105.7.16
        'bRemoteBpath = True
        'bRemoteApath = True

        If module1.authority > 2 Then LabelUser.ForeColor = Color.Green
        If module1.authority = 4 Then
            Panel3.Visible = True
        End If


        '107.6.29
        If bRadioButton_REMOTE Then
            CheckBoxLink.Visible = False
        End If

        '107.5.23 
        'VGAP = 3
        VGAP = (Me.Height - Label2.Height - Panel4.Height - Panel6.Height - dgvWater.Height - DataGridView2.Height - dgvScale.Height - dgvPan.Height - Panel5.Height) / 8

        '107.3.16
        TLog.logName = "tlg"
        MonLog.logName = "lg2"
        MainLog.logName = "lgm"
        DbLog.logName = "lgd"
        '109.4.2 配比連線 bPbSync
        UploadPbLog.logName = "lup"
        '111.10.21 rec連線 
        RecSqlLog.logName = "rec"

        '105.8.10 環泥烏日
        If bCheckBoxSP6 Then
            s_receia = "recei1a"
        Else
            s_receia = "receia"
        End If

        '103.6.7
        bOutway = "A"

        '103.3.13
        If bImperial Then
            sUNIT = "lb"
        End If

        '102.4.17 
        If bEngVer Then
            'CheckBoxLink.Text = "MAN"
            btnbegin.Text = "GO"
            lblcode.Text = cRECIPE
            lblTri.Text = cM3
            lblcar.Text = cCARNO
            lblfield.Text = "Site No."
            Label8.Text = "Project Name"
            Label9.Text = cDESTINATION
            Label11.Text = cCUSTOMER

            lblmemo1.Text = "Mix.T."
            lblmemo3.Text = "Disch.T."
            lblstrength.Text = "Strength"
            Label4.Text = "Total Volume"
            '102.8.13 
            'Label13.Font = New Font("Arial Narrow", 12)
            Label13.Font = New Font("Arial Narrow", 10)
            Label13.Text = "Truck"
            lbl_daysum.Font = Label13.Font
            lbl_daysum.Text = "Day"
            lbl_monthsum.Font = Label13.Font
            lbl_monthsum.Text = "Mon."
            Label3.Font = Label13.Font
            Label3.Text = "Tot."
            LabelUser.Text = "Username:"
            PasswordLabel.Text = "Input Password"
            Label10.Font = New Font("Arial Narrow", 11)
            Label10.Text = "Done"
            lblwork.Font = New Font("Arial Narrow", 14)
            lblwork.Text = "Batching"
            btnque1.Text = "Que1"
            btnque2.Text = "Que2"
            LabelWater.Text = "Mois.%"
            LabelA1.Font = New Font("Arial Narrow", 11)
            LabelA2.Font = LabelA1.Font
            LabelA3.Font = LabelA1.Font
            LabelA1.Text = "Hint"
            LabelA2.Text = "Warn"
            LabelA3.Text = "Alarm"
            Label11m3.Font = New Font("Arial Narrow", 10)
            '103.3.13
            'Label11m3.Text = "Set(Kg/m³)"
            Label11m3.Text = "Set(" & sUNIT & "/cube)"
            Label5.Font = New Font("Arial Narrow", 10)
            '103.3.13
            'Label5.Text = "Net(Kg/m³)"
            Label5.Text = "Net(" & sUNIT & "/cube)"
            LabelWeight.Text = "Weight"
            Btnfomula.Text = cRECIPE
            btnDynamic.Text = "Dynamic"
            Button2.Text = "History"
            Button1.Text = "Exit"
            btnparam.Text = "Parameter"
            'CheckBoxContinue.Text = "Manual"
            'Btnfomula.Text = "Recipe"
            'Btnfomula.Text = "Recipe"
            'Btnfomula.Text = "Recipe"
            DataGridViewAlarm.Columns(0).HeaderText = "Time"
            DataGridViewAlarm.Columns(1).HeaderText = "Desc."
            dgvPan.Columns(0).HeaderCell.Style.Font = New Font("Arial Narrow", 11)
            dgvPan.Columns(0).HeaderText = "Batch"
            dgvPan.Columns(1).HeaderText = "1"
            dgvPan.Columns(2).HeaderText = "2"
            dgvPan.Columns(3).HeaderText = "3"
            dgvPan.Columns(4).HeaderText = "4"
            dgvPan.Columns(5).HeaderText = "5"
            dgvPan.Columns(6).HeaderText = "6"
            dgvPan.Columns(7).HeaderText = "7"
            dgvPan.Columns(8).HeaderText = "8"
            dgvPan.Columns(9).HeaderText = "9"
            dgvPan.Columns(10).HeaderText = "10"

            CheckBoxsSim.Text = "Print"
            Label12.Text = "Veri."

            DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial Narrow", 12)

            DataGridViewAlarm.DefaultCellStyle.Font = New Font("Arial Narrow", 12)
            DataGridViewAlarm.ColumnHeadersDefaultCellStyle.Font = New Font("Arial Narrow", 12)

            '102.10.19
            Label2.Font = New Font("Arial", 20)
            Label8.Text = ""
            Label9.Text = ""
            Label11.Text = ""

            '103.4.22
            btn_simstart.Text = "Start"
            btn_simstop.Text = "Cancel"
            lbl_simtime.Text = "Time to Start"

        End If
        '105.5.6
        'tbxstrengt.Width = 350
        tbxstrengt.Font = New Font("Arial", 16)
        tbxstrengt.Width = 420
        tbxstrengt.Left = lblstrength.Left + lblstrength.Width
        LabelArea.Left = tbxstrengt.Left + tbxstrengt.Width
        cbxcorrect.Left = LabelArea.Left + LabelArea.Width
        LabelQkey.Left = cbxcorrect.Left + cbxcorrect.Width + 5
        LabelQkey0.Left = cbxcorrect.Left + cbxcorrect.Width + 5
        LabelQkey.Width = 120
        LabelQkey0.Width = 120
        '105.5.14
        LabelQkey1.Left = LabelQkey.Left + LabelQkey.Width + 5
        LabelQkey2.Left = LabelQkey.Left + LabelQkey.Width + 5
        LabelQkey1.Width = 120
        LabelQkey2.Width = 120

        '103.4.19 4.22
        LabelUser.Text &= module1.user
        LabelUser.Text &= "-" & module1.authority


        '100.8.12
        numquetri1.Items.Clear()
        If fMaxM3 <= 1 Then
            numquetri1.Items.Add("0.1")
            numquetri1.Items.Add("0.2")
            numquetri1.Items.Add("0.25")
            numquetri1.Items.Add("0.3")
            numquetri1.Items.Add("0.4")
            numquetri1.Items.Add("0.5")
            numquetri1.Items.Add("0.6")
            numquetri1.Items.Add("0.7")
            numquetri1.Items.Add("0.8")
            numquetri1.Items.Add("0.9")
            numquetri1.Items.Add("1.0")
            numquetri1.Items.Add("1.1")
            numquetri1.Items.Add("1.2")
            numquetri1.Items.Add("1.3")
            numquetri1.Items.Add("1.4")
            numquetri1.Items.Add("1.5")
        Else
            numquetri1.Items.Add("1.0")
            numquetri1.Items.Add("2.0")
            numquetri1.Items.Add("3.0")
            numquetri1.Items.Add("3.5")
            numquetri1.Items.Add("4.0")
            numquetri1.Items.Add("4.5")
            numquetri1.Items.Add("5.0")
            numquetri1.Items.Add("5.5")
            numquetri1.Items.Add("6.0")
            numquetri1.Items.Add("6.5")
            numquetri1.Items.Add("7.0")
            numquetri1.Items.Add("7.5")
            numquetri1.Items.Add("8.0")
            numquetri1.Items.Add("8.5")
            numquetri1.Items.Add("9.0")
            numquetri1.Items.Add("9.5")
            numquetri1.Items.Add("10.0")
            numquetri1.Items.Add("10.5")
            numquetri1.Items.Add("11.0")
            numquetri1.Items.Add("11.5")
            numquetri1.Items.Add("12.0")
        End If
        numquetri2.Items.Clear()
        If fMaxM3 <= 1 Then
            numquetri2.Items.Add("0.1")
            numquetri2.Items.Add("0.2")
            numquetri2.Items.Add("0.25")
            numquetri2.Items.Add("0.3")
            numquetri2.Items.Add("0.4")
            numquetri2.Items.Add("0.5")
            numquetri2.Items.Add("0.6")
            numquetri2.Items.Add("0.7")
            numquetri2.Items.Add("0.8")
            numquetri2.Items.Add("0.9")
            numquetri2.Items.Add("1.0")
            numquetri2.Items.Add("1.1")
            numquetri2.Items.Add("1.2")
            numquetri2.Items.Add("1.3")
            numquetri2.Items.Add("1.4")
            numquetri2.Items.Add("1.5")
        Else
            numquetri2.Items.Add("1.0")
            numquetri2.Items.Add("2.0")
            numquetri2.Items.Add("3.0")
            numquetri2.Items.Add("3.5")
            numquetri2.Items.Add("4.0")
            numquetri2.Items.Add("4.5")
            numquetri2.Items.Add("5.0")
            numquetri2.Items.Add("5.5")
            numquetri2.Items.Add("6.0")
            numquetri2.Items.Add("6.5")
            numquetri2.Items.Add("7.0")
            numquetri2.Items.Add("7.5")
            numquetri2.Items.Add("8.0")
            numquetri2.Items.Add("8.5")
            numquetri2.Items.Add("9.0")
            numquetri2.Items.Add("9.5")
            numquetri2.Items.Add("10.0")
            numquetri2.Items.Add("10.5")
            numquetri2.Items.Add("11.0")
            numquetri2.Items.Add("11.5")
            numquetri2.Items.Add("12.0")
        End If


        colorSand1 = Color.FromArgb(255, 192, 192)
        colorStone1 = Color.FromArgb(192, 192, 255)
        colorWater1 = Color.FromArgb(192, 255, 255)
        colorCement1 = Color.FromArgb(255, 255, 164)
        colorDrog1 = Color.FromArgb(164, 255, 164)

        colorSand2 = Color.FromArgb(255, 164, 164)
        colorStone2 = Color.FromArgb(164, 164, 255)
        colorWater2 = Color.FromArgb(164, 255, 255)
        colorCement2 = Color.FromArgb(255, 255, 100)
        colorDrog2 = Color.FromArgb(100, 255, 100)

        '101.7.25Me.Width = 1280 => no use since Max
        '107.5.21
        'Me.Width = 1024
        'Me.Height = 768

        '113.7.5
        'LabelTime.Left = 1050
        LabelTime.Left = Me.Width - LabelTime.Width - 20
        LabelTime.Top = 8
        LabelCust.Left = LabelTime.Left + LabelTime.Width - 5
        'LabelQkey.Left = LabelTime.Left + LabelTime.Width - 5


        cbxcorrect.Text = "0"
        tbxspare.Visible = chk_A_B.Checked
        lblspare.Visible = chk_A_B.Checked
        'dtconfig = db.GetDataTable("select material, title from config where type = 'Pond' order by Idx")
        dtconfig = db.GetDataTable("select material, title,setting from config where type = 'Pond' order by Idx")

        j1 = 0
        j2 = 0
        j3 = 0
        For i = 0 To 14
            If dtconfig.Rows(i).Item(0).ToString = "使用" Then
                j1 += 1
                SiloUse(i) = True
                DgvLeft(i) = j1
                If i < 8 Then
                    j2 += 1
                Else
                    j3 += 1
                End If
            Else
                SiloUse(i) = False
                DgvLeft(i) = 0
            End If
        Next
        SiloCounts = j1

        j1 = 0
        For i = 0 To 7
            If dtconfig.Rows(i).Item(0).ToString = "使用" Then
                j1 += 1
                SiloUse1(i) = True
                ZgcLeft1(i) = j1
            Else
                SiloUse1(i) = False
                ZgcLeft1(i) = 0
            End If
            SiloMatName1(i) = dtconfig.Rows(i).Item(1)
        Next
        SiloCount1 = j1
        If SiloCount1 = 0 Then SiloCount1 = 1

        j1 = 0
        For i = 8 To 14
            If dtconfig.Rows(i).Item(0).ToString = "使用" Then
                j1 += 1
                SiloUse2(i - 8) = True
                ZgcLeft2(i - 8) = j1
            Else
                SiloUse2(i - 8) = False
                ZgcLeft2(i - 8) = 0
            End If
            SiloMatName2(i - 8) = dtconfig.Rows(i).Item(1)
        Next
        SiloCount2 = j1
        If SiloCount2 = 0 Then SiloCount2 = 1
        '0420
        SiloCounts = 8

        'dtconfig = db.GetDataTable("select title, setting, scale from config where type = 'Material' order by Idx")
        'dtconfig = db.GetDataTable("select title, setting, scale,sr from config where type = 'Material' order by Idx")
        dtconfig = db.GetDataTable("select title, setting, scale,sr,Material from config where type = 'Material' order by Idx")

        DataGridView2.ColumnCount = 30
        DataGridView2.Rows.Add(4)
        DataGridView2.ColumnHeadersVisible = True
        DataGridView2.RowHeadersVisible = False
        '102.4.17
        Dim myFont = New Font("全真顏體", 14)
        DataGridView2.Rows(0).Cells(0).Style.Font = DataGridView2.ColumnHeadersDefaultCellStyle.Font
        DataGridView2.Rows(1).Cells(0).Style.Font = DataGridView2.ColumnHeadersDefaultCellStyle.Font
        'DataGridView2.Rows(0).Cells(0).Style.Font = New Font("標楷體", 14)
        'DataGridView2.Rows(1).Cells(0).Style.Font = New Font("標楷體", 14)
        If bEngVer Then
            DataGridView2.Rows(0).Cells(0).Style.Font = New Font("全真顏體", 13)
            DataGridView2.Rows(0).Cells(0).Value = "ORIGINAL"
            DataGridView2.Rows(1).Cells(0).Style.Font = New Font("全真顏體", 13)
            DataGridView2.Rows(1).Cells(0).Value = "SETTING"
            DataGridView2.Rows(2).Cells(0).Value = "STATUS"
        Else
            DataGridView2.Rows(0).Cells(0).Value = "單方"
            DataGridView2.Rows(1).Cells(0).Value = "設定"
            DataGridView2.Rows(2).Cells(0).Value = "狀態"
        End If
        DataGridView2.Rows(0).Height = 30
        DataGridView2.Rows(1).Height = 30
        DataGridView2.Rows(2).Height = 25
        DataGridView2.Rows(3).Height = 0
        'For i = 0 To dtconfig.Rows.Count - 1
        For i = 0 To dtconfig.Rows.Count - 1
            DataGridView2.Rows(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Rows(2).Cells(i).Style.Font = myFont
        Next
        DataGridView2.Rows(2).Visible = True
        DataGridView2.Rows(3).Visible = False
        DataGridView2.Width = Me.Width - 6

        dgvWater.ColumnCount = 9
        dgvWater.Rows.Add()

        MatCounts = 0
        For i = 0 To dtconfig.Rows.Count - 1
            '材料名稱
            MatName(i) = dtconfig.Rows(i).Item(0).ToString
            '105.5.23
            MatName_O(i) = dtconfig.Rows(i).Item(0).ToString
            '材料種類
            MatType(i) = dtconfig.Rows(i).Item(4).ToString
            '99.07.28
            If MatType(i) = "S" Or MatType(i) = "G" Then
                SiloMatType1(CInt(dtconfig.Rows(i).Item(1).ToString)) = dtconfig.Rows(i).Item(4)
            End If
            If dtconfig.Rows(i).Item(1).ToString <> "0" Then
                MatUse(i) = dtconfig.Rows(i).Item(1)
                MatCounts += 1
                'DataGridView2.Rows(0).Cells(CType(dtconfig.Rows(i).Item(1).ToString, Integer) - 1).Value += dtconfig.Rows(i).Item(0).ToString + " "
                st = dtconfig.Rows(i).Item(3)
                j = dtconfig.Rows(i).Item(1) - 1
                '101.4.8 S4 *.* 
                If i = 3 And bCheckBoxS4 Then
                    S4SiloNo = j
                End If
                SiloMatSer(j, SiloMatCount(j)) = dtconfig.Rows(i).Item(3)
                MatSiloSR(i) = dtconfig.Rows(i).Item(3)
                SiloMatCount(j) = SiloMatCount(j) + 1
                MatSiloNo(i) = dtconfig.Rows(i).Item(1)
                MatSiloOrder(i) = SiloMatCount(j)
                'MatSiloNo(i) = dtconfig.Rows(i).Item(1)
                DataGridView2.Columns(i + 1).visible = True
                'DataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightPink
                '101.4.5 ?? SiloMatType(j) = dtconfig.Rows(i).Item(4)
                SiloMatType(j) = dtconfig.Rows(i).Item(4)
                If dtconfig.Rows(i).Item(4) = "S" Then
                    'DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = Color.FromArgb(150, 0, 0)
                    DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = colorSand1
                    DataGridView2.Rows(1).Cells(i + 1).Style.BackColor = colorSand1
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = colorSand2
                    DataGridView2.Rows(3).Cells(i + 1).Style.BackColor = colorSand2
                    dgvWater.Columns(i).visible = True
                ElseIf dtconfig.Rows(i).Item(4) = "G" Then
                    DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = colorStone1
                    DataGridView2.Rows(1).Cells(i + 1).Style.BackColor = colorStone1
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = colorStone2
                    DataGridView2.Rows(3).Cells(i + 1).Style.BackColor = colorStone2
                    dgvWater.Columns(i).visible = True
                ElseIf dtconfig.Rows(i).Item(4) = "W" Then
                    DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = colorWater1
                    DataGridView2.Rows(1).Cells(i + 1).Style.BackColor = colorWater1
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = colorWater2
                    DataGridView2.Rows(3).Cells(i + 1).Style.BackColor = colorWater2
                ElseIf dtconfig.Rows(i).Item(4) = "C" Then
                    DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = colorCement1
                    DataGridView2.Rows(1).Cells(i + 1).Style.BackColor = colorCement1
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = colorCement2
                    DataGridView2.Rows(3).Cells(i + 1).Style.BackColor = colorCement2
                ElseIf dtconfig.Rows(i).Item(4) = "A" Then
                    DataGridView2.Rows(0).Cells(i + 1).Style.BackColor = colorDrog1
                    DataGridView2.Rows(1).Cells(i + 1).Style.BackColor = colorDrog1
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = colorDrog2
                    DataGridView2.Rows(3).Cells(i + 1).Style.BackColor = colorDrog2
                End If
                'DataGridView2.Rows(0).Cells(i).Value() = MatName(i)
            Else
                DataGridView2.Columns(i + 1).visible = False
            End If
        Next
        'Public MatSiloNoFirst(28) As Boolean     '材料磅號 99.12.04 is first Mater of this silo for remain of simulatio
        Dim flag(15) As Integer
        For j = 1 To 14
            flag(j) = 0
        Next
        For i = 0 To dtconfig.Rows.Count - 1
            MatSiloNoFirst(i) = True
            For j = 1 To 14
                If MatSiloNo(i) = j Then
                    flag(j) += 1
                    If flag(j) > 1 Then MatSiloNoFirst(i) = False
                End If
            Next
        Next
        'If MatCounts = 0 Then MatCounts += 1
        For i = 0 To 14
            If SiloMatType(i) = "A" Then
                'SiloFullScale(i) = CInt(SiloFullScale(i) * 0.01)
            End If
        Next
        MatCounts += 1
        j = (Me.Width - 10) \ MatCounts
        DataGridView2.Columns(0).Width = j
        If bEngVer Then
            DataGridView2.ColumnHeadersDefaultCellStyle.Font = txtremain.Font
            DataGridView2.Columns(0).HeaderText = sUNIT
        Else
            '108.6.12
            'DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 14)
            DataGridView2.Columns(0).HeaderText = sUNIT
        End If

        For i = 0 To 25
            DataGridView2.Columns(i + 1).width = j
            '105.5.7 A1:17 no action
            DataGridView2.Columns(i + 1).headertext = MatName(i)
        Next

        j = 0
        For i = 0 To 7
            If dtconfig.Rows(i).Item(1).ToString <> "0" Then
                dgvWater.Columns(i).width = DataGridView2.Columns(i + 1).Width
                j += dgvWater.Columns(i).width
            Else
                dgvWater.Columns(i).visible = False
            End If
        Next
        dgvWater.Width = j

        renewcmxcode()

        j = 0
        For i = 0 To 7
            If dtconfig.Rows(i).Item(1).ToString = "0" Then
                exist(i) = False
            Else
                exist(i) = True
                j += 1
            End If
        Next
        j1 = 0
        For i = 8 To 24
            If dtconfig.Rows(i).Item(1).ToString = "0" Then
                exist(i) = False
            Else
                exist(i) = True
                j1 += 1
            End If
        Next

        If j < j1 Then j = j1
        j1 = (DataGridView1.Height - 30) \ (j + 1)
        For i = 0 To j - 1
            Dim row() As String = {"0.0", "", "", "", "", "", ""}
            DataGridView1.Rows.Add(row)
            DataGridView1.Rows(i).Height = j1
        Next
        DataGridView1.ColumnHeadersHeight = DataGridView1.Rows(0).Height
        j = 0
        water_item = 0
        For i = 0 To 7
            If exist(i) Then
                DataGridView1.Rows(j).Cells(1).Value = dtconfig.Rows(i).Item(0).ToString
                DataGridView1.Rows(j).Cells(0).Style.Format = "D01"
                DataGridView1.Rows(j).Cells(0).Style.ForeColor = Color.Blue
                DataGridView1.Rows(j).Cells(0).ReadOnly = False
                DataGridView1.Rows(j).Cells(2).Style.Format = "D0" + dtconfig.Rows(i).Item("scale").ToString
                DataGridView1.Rows(j).Cells(3).Style.Format = "D" + dtconfig.Rows(i).Item("scale").ToString
                If dtconfig.Rows(i).Item(4).ToString = "S" Then
                    DataGridView1.Rows(j).Cells(0).Style.BackColor = colorSand1
                    DataGridView1.Rows(j).Cells(1).Style.BackColor = colorSand1
                    DataGridView1.Rows(j).Cells(2).Style.BackColor = colorSand1
                    DataGridView1.Rows(j).Cells(3).Style.BackColor = colorSand1
                    water_item += 1
                ElseIf dtconfig.Rows(i).Item(4).ToString = "G" Then
                    DataGridView1.Rows(j).Cells(0).Style.BackColor = colorStone1
                    DataGridView1.Rows(j).Cells(1).Style.BackColor = colorStone1
                    DataGridView1.Rows(j).Cells(2).Style.BackColor = colorStone1
                    DataGridView1.Rows(j).Cells(3).Style.BackColor = colorStone1
                    water_item += 1
                Else
                    DataGridView1.Rows(j1).Cells(4).Style.BackColor = Color.GhostWhite
                    DataGridView1.Rows(j1).Cells(5).Style.BackColor = Color.GhostWhite
                    DataGridView1.Rows(j1).Cells(6).Style.BackColor = Color.GhostWhite
                End If
                j += 1
            End If
        Next
        j1 = 0
        For i = 8 To 24
            If exist(i) Then
                If j1 > j - 1 Then
                    DataGridView1.Rows(j1).Cells(0).ReadOnly = True
                    DataGridView1.Rows(j1).Cells(0).Value = ""
                End If
                DataGridView1.Rows(j1).Cells(4).Value = dtconfig.Rows(i).Item(0).ToString
                DataGridView1.Rows(j1).Cells(5).Style.Format = "F" & dtconfig.Rows(i).Item("scale").ToString
                DataGridView1.Rows(j1).Cells(6).Style.Format = "F" & dtconfig.Rows(i).Item("scale").ToString
                If dtconfig.Rows(i).Item(4).ToString = "W" Then
                    DataGridView1.Rows(j1).Cells(4).Style.BackColor = colorWater1
                    DataGridView1.Rows(j1).Cells(5).Style.BackColor = colorWater1
                    DataGridView1.Rows(j1).Cells(6).Style.BackColor = colorWater1
                ElseIf dtconfig.Rows(i).Item(4).ToString = "C" Then
                    DataGridView1.Rows(j1).Cells(4).Style.BackColor = colorCement1
                    DataGridView1.Rows(j1).Cells(5).Style.BackColor = colorCement1
                    DataGridView1.Rows(j1).Cells(6).Style.BackColor = colorCement1
                ElseIf dtconfig.Rows(i).Item(4).ToString = "A" Then
                    DataGridView1.Rows(j1).Cells(4).Style.BackColor = colorDrog1
                    DataGridView1.Rows(j1).Cells(5).Style.BackColor = colorDrog1
                    DataGridView1.Rows(j1).Cells(6).Style.BackColor = colorDrog1
                Else
                    DataGridView1.Rows(j1).Cells(4).Style.BackColor = Color.GhostWhite
                    DataGridView1.Rows(j1).Cells(5).Style.BackColor = Color.GhostWhite
                    DataGridView1.Rows(j1).Cells(6).Style.BackColor = Color.GhostWhite
                End If
                j1 += 1
            End If
        Next

        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(5).Visible = False

        DataGridView1.Columns(0).Width = 75 '%
        DataGridView1.Columns(1).Width = 55 'name
        DataGridView1.Columns(2).Width = 80 '1m3
        DataGridView1.Columns(3).Width = DataGridView1.Columns(2).Width '總重
        DataGridView1.Columns(4).Width = 55 'name
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = DataGridView1.Columns(5).Width
        j = 0
        For i = 0 To 2
            j += DataGridView1.Columns(i).Width
        Next
        For i = 4 To 5
            j += DataGridView1.Columns(i).Width
        Next
        DataGridView1.Width = j + 3
        DataGridView1.Left = Me.ClientSize.Width - DataGridView1.Width - 10
        DataGridView1.Columns(0).Resizable = DataGridViewTriState.False

        j = 0
        For i = 0 To DataGridView1.RowCount - 1
            j += DataGridView1.Rows(i).Height
        Next
        j += DataGridView1.ColumnHeadersHeight + 2
        DataGridView1.Height = j
        '107.5.23
        'Panel6.Top = Panel4.Top + Panel4.Height + VGAP
        'LabelWater.Top = Panel6.Top + Panel6.Height + VGAP
        'dgvWater.Top = LabelWater.Top
        'DataGridView2.Top = dgvWater.Top + dgvWater.Height + VGAP
        'dgvScale.Top = LabelWeight.Top
        'LabelWeight.Top = DataGridView2.Top + DataGridView2.Height + VGAP
        'LabelErr.Top = LabelWeight.Top + LabelWeight.Height + 2

        dgvWater.ColumnHeadersVisible = False
        dgvWater.Rows(0).Height = 30
        dgvWater.Height = dgvWater.Rows(0).Height + 3
        LabelWater.Height = dgvWater.Height
        LabelWater.Left = DataGridView2.Left
        LabelWater.Width = DataGridView2.Columns(0).Width
        LabelWater.Height = dgvWater.Height
        dgvWater.Left = DataGridView2.Left + DataGridView2.Columns(0).Width
        dgvWater.Rows(0).Cells(8).Selected = True

        DataGridView2.Left = 3

        DataGridView2.Height = 1 + DataGridView2.Rows(0).Height + DataGridView2.Rows(1).Height + DataGridView2.Rows(2).Height + DataGridView2.Rows(3).Height + DataGridView2.ColumnHeadersHeight
        dgvScale.ColumnCount = 16

        dgvScale.Rows.Add(2)
        dgvScale.ColumnHeadersVisible = False
        dgvScale.Rows(0).Height = 26
        dgvScale.Rows(1).Height = 26
        dgvScale.Height = dgvScale.Rows(0).Height * 2 + 2
        dgvScale.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgvScale.Rows(0).Cells(15).Selected = True
        LabelWeight.Height = dgvScale.Rows(0).Height
        LabelErr.Height = LabelWeight.Height
        LabelWeight.Left = DataGridView2.Left
        LabelErr.Left = LabelWeight.Left
        LabelWeight.Width = DataGridView2.Columns(0).Width
        LabelErr.Width = LabelWeight.Width
        dgvScale.Left = DataGridView2.Left + DataGridView2.Columns(0).Width
        j = 0
        j1 = DataGridView2.Columns(0).Width
        j2 = 0
        For i = 0 To dgvScale.ColumnCount - 1
            If SiloUse(i) Then
                j = SiloMatCount(i) * j1
                dgvScale.Columns(i).width = j
                j2 += j
            Else
                dgvScale.Columns(i).visible = False
            End If
        Next
        dgvScale.Width = j2 + 3

        DataGridView1.Top = dgvScale.Top + dgvScale.Height + VGAP

        Dim file
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\water.per")
        If file = False Then
            For i = 1 To 15
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\water.per", "0.0" & vbCrLf, True, System.Text.Encoding.Default)
            Next
        End If
        Dim s As String = "0.0"
        i = DataGridView1.RowCount
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\water.per", OpenMode.Input)
        If Not EOF(fileNum) Then
            For j = 0 To 7
                's = LineInput(fileNum)
                'Input(fileNum, MainForm.DataGridView1.Rows(j).Cells(0).Value)
                s = LineInput(fileNum)
                'DataGridView1.Rows(j).Cells(0).Value = s
                dgvWater.Rows(0).Cells(j).Value = Trim(s)
                '104.8.10
                fWatOrig(j) = Val(s)
            Next
            FileClose(fileNum)
        End If
        FormMainLocate()

        flag_last = "0"
        flag_unicar = "0"
        dbrec.changedb(sYMdbName + sSavingYear + ".mdb")

        '判斷年度子資料夾是否存在
        '100.4.18
        'If Not My.Computer.FileSystem.DirectoryExists(sYMdbName + sSavingYear + "\M" + Today.Month.ToString + "_A") Then
        '    My.Computer.FileSystem.CreateDirectory(sYMdbName + sSavingYear + "\M" + Today.Month.ToString + "_A")
        'End If
        'If Not My.Computer.FileSystem.DirectoryExists(sYMdbName + sSavingYear + "\M" + Today.Month.ToString + "_B") Then
        '    My.Computer.FileSystem.CreateDirectory(sYMdbName + sSavingYear + "\M" + Today.Month.ToString + "_B")
        'End If

        '給定每日文字檔檔名
        'file_A = symdbname + sSavingYear + "\M" + Today.Month.ToString + "_A\" + sSavingYear + Today.Month.ToString + Today.Day.ToString + ".txt"
        'file_B = symdbname + sSavingYear + "\M" + Today.Month.ToString + "_B\" + sSavingYear + Today.Month.ToString + Today.Day.ToString + ".txt"
        'file_print = "\JS\CBC8\DATA\printfile.txt"


        Call ReadAlarm()
        '102.11.23
        Call ReadCement()

        '105.5.12
        Call ReadAE15()
        '111.11.30 print offset 列印調整
        Call ReadOffset()



        'txt_daysum.Text = (GetSetting("JS", "CBC800", "txt_daysum", "0.00"))
        'txt_monthsum.Text = (GetSetting("JS", "CBC800", "txt_monthsum", "0.00"))
        lblMachTot.Text = (GetSetting("JS", "CBC800", "lblMachTot", "1234.00"))
        '107.4.14
        MachTot = CSng(lblMachTot.Text)
        'SaveSetting("JS", "CBC800", "lblMachTot", lblMachTot.Text)
        Num_quetime_set.Text = (GetSetting("JS", "CBC800", "Num_quetime_set", "40"))

        '100.2.24
        DataGridViewQue.RowCount = 4
        DataGridViewQue.Left = Panel1.Left + Panel1.Width + 5
        DataGridViewQue.Height = Panel1.Height
        DataGridViewQue.Width = Panel5.Width - Panel1.Width - 10
        DataGridViewQue.Rows(0).Cells(0).Value = ""
        DataGridViewQue.Rows(0).Cells(1).Value = "0"
        DataGridViewQue.Rows(0).Cells(2).Value = ""
        DataGridViewQue.Rows(0).Cells(3).Value = ""
        DataGridViewQue.Rows(0).Cells(4).Value = ""
        DataGridViewQue.Rows(0).Cells(5).Value = ""
        DataGridViewQue.Rows(0).Cells(6).Value = ""
        DataGridViewQue.Rows(0).Cells(7).Value = ""
        '106.11.2
        DataGridViewQue.Rows(0).Cells(8).Value = ""
        DataGridViewQue.Rows(0).Cells(9).Value = ""

        DataGridViewQue.Rows(1).Cells(0).Value = ""
        DataGridViewQue.Rows(1).Cells(1).Value = "0"
        DataGridViewQue.Rows(1).Cells(2).Value = ""
        DataGridViewQue.Rows(1).Cells(3).Value = ""
        DataGridViewQue.Rows(1).Cells(4).Value = ""
        DataGridViewQue.Rows(1).Cells(5).Value = ""
        DataGridViewQue.Rows(1).Cells(6).Value = ""
        DataGridViewQue.Rows(1).Cells(7).Value = ""
        '106.11.2
        DataGridViewQue.Rows(1).Cells(8).Value = ""
        DataGridViewQue.Rows(1).Cells(9).Value = ""

        DataGridViewQue.Rows(2).Cells(0).Value = ""
        DataGridViewQue.Rows(2).Cells(1).Value = "0"
        DataGridViewQue.Rows(2).Cells(2).Value = ""
        DataGridViewQue.Rows(2).Cells(3).Value = ""
        DataGridViewQue.Rows(2).Cells(4).Value = ""
        DataGridViewQue.Rows(2).Cells(5).Value = ""
        DataGridViewQue.Rows(2).Cells(6).Value = ""
        DataGridViewQue.Rows(2).Cells(7).Value = ""
        '106.11.2
        DataGridViewQue.Rows(2).Cells(8).Value = ""
        DataGridViewQue.Rows(2).Cells(9).Value = ""

        DataGridViewQue.Rows(3).Cells(0).Value = ""
        DataGridViewQue.Rows(3).Cells(1).Value = "0"
        DataGridViewQue.Rows(3).Cells(2).Value = ""
        DataGridViewQue.Rows(3).Cells(3).Value = ""
        DataGridViewQue.Rows(3).Cells(4).Value = ""
        DataGridViewQue.Rows(3).Cells(5).Value = ""
        DataGridViewQue.Rows(3).Cells(6).Value = ""
        DataGridViewQue.Rows(3).Cells(7).Value = ""
        '106.11.2
        DataGridViewQue.Rows(3).Cells(8).Value = ""
        DataGridViewQue.Rows(3).Cells(9).Value = ""
        'DataGridViewQue.Rows(3).Cells(0).Value = "635520"
        'DataGridViewQue.Rows(3).Cells(1).Value = "12.34"
        'DataGridViewQue.Rows(3).Cells(2).Value = "9876-YQ"
        'DataGridViewQue.Rows(3).Cells(3).Value = "OR95110001"
        'DataGridViewQue.Rows(3).Cells(4).Value = "工程名稱國民隊"
        'DataGridViewQue.Rows(3).Cells(5).Value = "施工部位草地"
        'DataGridViewQue.Rows(3).Cells(6).Value = "客戶名稱王建民"
        'DataGridViewQue.Rows(3).Cells(7).Value = "01|F1|SH05011432|0004|N"
        DataGridViewQue.Rows(3).Cells(7).Selected = True
        DataGridViewQue.Enabled = False

        'Panel5.Height = DataGridViewAlarm.Height + 10
        For i = 0 To DataGridViewAlarm.ColumnCount - 1
            DataGridViewAlarm.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        FormLoded = True

        '填入每日完成方數()
        Dim db_sum As DataTable
        '104.11.2 11.8
        s = Today.ToShortDateString
        '105.1.2
        'db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")

        '111.12.28
        'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') =  '" + Today.Year.ToString + "/" + Today.Month.ToString + "/" + Today.Day.ToString + "'")
        'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
        Dim cube_tot As Single = 0
        '113.2.14
        Try
            db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m/d') =  '" + Today.Year.ToString + "/" + Today.Month.ToString + "/" + Today.Day.ToString + "'")
            For j = 0 To db_sum.Rows.Count - 1
                cube_tot += db_sum.Rows(j).Item("cube")
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        txt_daysum.Text = Format(cube_tot, "0.00")

        db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m') =  '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
        cube_tot = 0
        For j = 0 To db_sum.Rows.Count - 1
            cube_tot += db_sum.Rows(j).Item("cube")
        Next
        txt_monthsum.Text = Format(cube_tot, "0.00")

        '106.1.2 111.12.28
        'Dim sum, d
        ''填入每月完成方數
        'db_sum.Clear()
        'db_sum = dbrec.GetDataTable("select DISTINCT  RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m') = '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
        'sum = 0
        'For i = 0 To db_sum.Rows.Count - 1
        '    d = db_sum.Rows(i).Item(0)
        '    sum += Format(Val(db_sum.Rows(i).Item(1).ToString), "0.00")
        'Next
        ''txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
        'txt_monthsum.Text = Format(sum, "0.00")

        If LoginForm1.Visible Then LoginForm1.Visible = False
        FrmMonit.Show()
        FrmMonit.SetupDGV1()
        FrmMonit.Setup_dgvSilo1(FrmMonit.dgvSilo1, 1)
        FrmMonit.PopulatedgvSilo1(FrmMonit.dgvSilo1, 1)
        i = FrmMonit.dgvSilo1.Top
        FrmMonit.Setup_dgvSilo1(FrmMonit.dgvSilo2, 2)
        FrmMonit.PopulatedgvSilo1(FrmMonit.dgvSilo2, 2)
        j = FrmMonit.dgvSilo2.Top
        If SiloUse1(0) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV1, 1)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc1, 1)
        Else
            FrmMonit.DGV1.Visible = False
            FrmMonit.Zgc1.Visible = False
        End If
        If SiloUse1(1) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV2, 2)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc2, 2)
        Else
            FrmMonit.DGV2.Visible = False
            FrmMonit.Zgc2.Visible = False
        End If
        If SiloUse1(2) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV3, 3)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc3, 3)
        Else
            FrmMonit.DGV3.Visible = False
            FrmMonit.Zgc3.Visible = False
        End If
        If SiloUse1(3) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV4, 4)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc4, 4)
        Else
            FrmMonit.DGV4.Visible = False
            FrmMonit.Zgc4.Visible = False
        End If
        If SiloUse1(4) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV5, 5)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc5, 5)
        Else
            FrmMonit.DGV5.Visible = False
            FrmMonit.Zgc5.Visible = False
        End If
        If SiloUse1(5) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV6, 6)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc6, 6)
        Else
            FrmMonit.DGV6.Visible = False
            FrmMonit.Zgc6.Visible = False
        End If
        If SiloUse1(6) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV7, 7)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc7, 7)
        Else
            FrmMonit.DGV7.Visible = False
            FrmMonit.Zgc7.Visible = False
        End If
        If SiloUse1(7) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV8, 8)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc8, 8)
        Else
            FrmMonit.DGV8.Visible = False
            FrmMonit.Zgc8.Visible = False
        End If
        If SiloUse2(0) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV9, 9)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc9, 9)
        Else
            FrmMonit.DGV9.Visible = False
            FrmMonit.Zgc9.Visible = False
        End If
        If SiloUse2(1) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV10, 10)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc10, 10)
        Else
            FrmMonit.DGV10.Visible = False
            FrmMonit.Zgc10.Visible = False
        End If
        If SiloUse2(2) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV11, 11)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc11, 11)
        Else
            FrmMonit.DGV11.Visible = False
            FrmMonit.Zgc11.Visible = False
        End If
        If SiloUse2(3) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV12, 12)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc12, 12)
        Else
            FrmMonit.DGV12.Visible = False
            FrmMonit.Zgc12.Visible = False
        End If
        If SiloUse2(4) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV13, 13)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc13, 13)
        Else
            FrmMonit.DGV13.Visible = False
            FrmMonit.Zgc13.Visible = False
        End If
        If SiloUse2(5) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV14, 14)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc14, 14)
        Else
            FrmMonit.DGV14.Visible = False
            FrmMonit.Zgc14.Visible = False
        End If
        If SiloUse2(6) Then
            FrmMonit.SetupDGV2(FrmMonit.DGV15, 15)
            FrmMonit.CreateZGC_ZBars(FrmMonit.Zgc15, 15)
        Else
            FrmMonit.DGV15.Visible = False
            FrmMonit.Zgc15.Visible = False
        End If
        i = FrmMonit.dgvSilo1.Top
        j = FrmMonit.dgvSilo2.Top
        Call FrmMonit.GetSiloFull()
        Call FrmMonit.createchart(FrmMonit.zgc)
        i = FrmMonit.dgvSilo1.Top
        j = FrmMonit.dgvSilo2.Top
        '101.4.2
        'FrmMonit.dgvSilo1.Height = FrmMonit.Height / 3
        'FrmMonit.dgvSilo2.Height = FrmMonit.Height / 3
        FrmMonit.Panel2.Height = FrmMonit.Height - (FrmMonit.dgvSilo1.Height + FrmMonit.dgvSilo2.Height) - 20
        FrmMonit.Panel1.Height = FrmMonit.Panel2.Height - 30
        'FrmMonit.Show()
        'FrmMonit.dgvPV1.Top = 1
        'FrmMonit.dgvSilo1.Top = 1
        'FrmMonit.dgvSilo2.Top = 251 + 5
        If bWaitTank Then
            FrmMonit.PanelTank.Visible = True
        Else
            FrmMonit.PanelTank.Visible = False
        End If
        i = FrmMonit.dgvSilo1.Top
        j = FrmMonit.dgvSilo2.Top
        '99.08.02 inital to A 
        DVW(830) = 0
        i = FrmMonit.WriteA2N(830, 1)
        '101.12.17 inital to 0 101.12.22 D839
        DVW(839) = 0
        i = FrmMonit.WriteA2N(839, 1)

        '101.12.17
        '110.9.24 生產完畢開始計時
        dtBatDon = Now


        '102.5.14 for print report
        '112.3.25 列印拌合結束電流
        'If bCheckBoxCBC7 Then
        '    If MatCounts >= 17 Then
        '        repFont = New Font("細明體", 11)
        '    Else
        '        repFont = New Font("細明體", 12)
        '    End If
        'Else
        '    If MatCounts >= 17 Then
        '        repFont = New Font("細明體", 9)
        '    Else
        '        repFont = New Font("細明體", 10)
        '    End If
        'End If
        If MatCounts >= 17 Then
            repFont = New Font("細明體", 9)
        Else
            repFont = New Font("細明體", 10)
        End If

        If MatCounts >= 17 Then
            repFont = New Font("細明體", 9)
        Else
            repFont = New Font("細明體", 10)
        End If

        '104.8.25
        strTrendPb = "12345678"
        strTrendM3 = "0.50"

        '104.7.10 104.8.27 ref. 105
        sbujiq1a = "bujiq" & sMixerNo & "a"

        '105.5.6
        Tbxquefield1.Left = lblfield.Left
        Tbxquefield2.Left = lblfield.Left


        '105.7.17
        Dim b
        If (bRadioButton_OP And bUseNetwork) Then '105.7.22
            '106.12.10 no TimerPing
            'TimerPing.Enabled = True
            TimerPing.Enabled = False
            '105.9.24 If bRemoteBpath Then
            'If bPingRPC And bRemoteBpath Then
            If (bPingRPC Or sTextBoxIP_R = "127.0.0.1") And bRemoteBpath Then
                ''105.11.5
                'LabelMessage2.Visible = True
                'LabelMessage2.Text = "資料搜尋中(Load_B)請稍後(Wait...) "
                'Me.Refresh()
                'Label4.Image = My.Resources._0013
                'Label4.Refresh()
                '105.7.19 105.8.13
                'b = CheckRemoteBfile(True)  'Last month
                'bLastMonthFile = False
                'b = CheckRemoteBfile(False)
                'Label4.Image = My.Resources._000
                'LabelMessage2.Visible = False

                '107.8.31 only 環泥烏日bCheckBoxSP6
                If bCheckBoxSP6 Then
                    '105.11.5
                    LabelMessage2.Visible = True
                    LabelMessage2.Text = "資料搜尋中(Load_B)請稍後(Wait...) "
                    Me.Refresh()
                    Label4.Image = My.Resources._0013
                    Label4.Refresh()
                    b = CheckRemoteBfile(True)  'Last month
                    bLastMonthFile = False
                    b = CheckRemoteBfile(False)
                    Label4.Image = My.Resources._000
                    LabelMessage2.Visible = False
                End If
            End If
            '105.9.24 If bRemoteApath Then
            '107.8.10 disable done by remote PC
            'If bPingRPC And bRemoteApath Then
            '    '105.11.5
            '    LabelMessage2.Visible = True
            '    LabelMessage2.Text = "資料搜尋中(Load_A)請稍候(Wait...) "
            '    Me.Refresh()
            '    Label4.Image = My.Resources._0013
            '    Label4.Refresh()
            '    '107.8.10 disable done by remote PC
            '    'b = CheckRemoteAfile(True)  'Last month
            '    '107.8.10 disable done by remote PC
            '    'b = CheckRemoteAfile(False)
            '    '105.11.5
            '    Label4.Image = My.Resources._000
            '    LabelMessage2.Visible = False
            'End If
        End If

        '105.9.16
        If bOnlyLocalPC Then
            file = My.Computer.FileSystem.DirectoryExists(sTextBoxBpath & "\CBC8\DATA_" & sProject)
            If file = False Then
                Try
                    My.Computer.FileSystem.CreateDirectory(sTextBoxBpath & "\CBC8\DATA_" & sProject)
                Catch ex As Exception
                    MsgBox(sTextBoxBpath & " not found! Please check LocalPC dir .")
                    bRemoteBpath = False
                End Try
            End If
            If Not My.Computer.FileSystem.FileExists(sYMdbName_B + sSavingYear + ".mdb") Then
                Try
                    My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B + sSavingYear + ".mdb")
                Catch ex As Exception
                    MsgBox("Copy " & sYMdbName_B + sSavingYear + ".mdb" & " fail!", MsgBoxStyle.Critical, "Load")
                End Try
            End If
            'bRemoteBpath = True
            '105.11.5   107.8.31 disabled => save in ProgramFiles... BB
            'LabelMessage2.Visible = True
            'LabelMessage2.Text = "資料搜尋中(Load_mdb)請稍候(Wait...) "
            'Me.Refresh()
            'Label4.Image = My.Resources._0013
            'Label4.Refresh()
            'b = CheckRemoteBfile(True)  'Last month
            'bLastMonthFile = False
            'b = CheckRemoteBfile(False)
            ''105.11.5
            'Label4.Image = My.Resources._000
            'LabelMessage2.Visible = False
        End If

        Label2.Text = sCompany
        Me.Refresh()
        '110.9.11
        ''107.7.9
        ''If bRemoteApath And bRemoteBpath Then LabelRPC.Image = My.Resources.connect Else LabelRPC.Image = My.Resources.disconnect
        'If bRemoteApath And bRemoteBpath Then
        '    LabelRPC.ForeColor = Color.Green
        'Else
        '    LabelRPC.ForeColor = Color.Purple
        'End If
        ''109.2.25
        'LabelRPC.Refresh()
        '106.7.22
        If bCheckBoxSP4 And bCheckBoxYadon Then
            Button4.Visible = True
        End If

        '107.1.31
        If bCheckBox232 Then TimerRx.Enabled = True

        '110.6.5 亞東 水泥確認
        If bCheckBoxYadon Then
            Dim str As String
            str = "水泥槽:"
            If bCheckBoxC1 Then str &= " C1"
            If bCheckBoxC2 Then str &= " C2"
            If bCheckBoxC3 Then str &= " C3"
            If bCheckBoxC4 Then str &= " C4"
            If bCheckBoxC5 Then str &= " C5"
            If bCheckBoxC6 Then str &= " C6"

            LabelCement.Text = str
            LabelCement.Top = Panel6.Top + Panel6.Height + 5
            '110.6.7
            LabelCement.Left = Panel6.Width * 0.3
            LabelCement.Visible = True
            TimerShowCement.Enabled = True
        End If
        '110.9.11
        '110.6.25
        TimerAlarm.Enabled = True
        '110.8.1
        CheckLabelJSForeColor()
        '111.7.22
        '   模擬報表 首盤 CheckBoxFirstPan
        CheckBoxFirstPan.Checked = bCheckBoxFirstPan

        '112.3.30    
        If bTextBoxSPARE5 Then
            If bPingSQLBigData Then
                Label10.ForeColor = Color.Green
                LabelBigDataErr.Visible = False
            Else
                Label10.ForeColor = Color.Black
                LabelBigDataErr.Visible = True
            End If
        End If


    End Sub

    Public Sub renewcmxcode()
        Dim i = 0
        Dim a As String = Cbxquecode1.Text
        Dim b As String = Cbxquecode2.Text
        Dim s As String
        'Dim dt1 As DataTable = db.GetDataTable("select code from fomula where code is not null")
        '100.4.4
        'Dim dt1 As DataTable = db.GetDataTable("select code from fomula where code is not null order by code")
        'Dim dt2 As DataTable = db.GetDataTable("select code from fomula where code is not null")
        db_recipe.changedb("\JS\CBC8\Conf_" & sProject & "\Recipe.mdb")
        '110.1.11
        'Dim dt1 As DataTable = db_recipe.GetDataTable("select code from fomula where code is not null order by code")
        'Dim dt2 As DataTable = db_recipe.GetDataTable("select code from fomula where code is not null")
        Dim dt1 As DataTable
        'Dim dt2 As DataTable
        If bPbSync Then
            dt1 = db_recipe.GetDataTable("select code,enable from fomula where code is not null AND enable='1' order by code")
            'dt2 = db_recipe.GetDataTable("select code,enable from fomula where code is not null AND enable='1' order by code")
        Else
            '111.6.24 配方列表
            'dt1 = db_recipe.GetDataTable("select code from fomula where code is not null order by code")
            dt1 = db_recipe.GetDataTable("select code, ID from fomula where code is not null order by code")
        End If
        Cbxquecode1.Items.Clear()
        Cbxquecode2.Items.Clear()
        For i = 0 To dt1.Rows.Count - 1
            s = dt1.Rows(i).Item(0)
            '111.6.24
            'If s <> "" Then
            '    Cbxquecode1.Items.Add(dt1.Rows(i).Item(0))
            '    Cbxquecode2.Items.Add(dt1.Rows(i).Item(0))
            'End If
            Dim sId
            sId = dt1.Rows(i).Item(1).ToString
            If s <> "" And s <> sId Then
                Cbxquecode1.Items.Add(dt1.Rows(i).Item(0))
                Cbxquecode2.Items.Add(dt1.Rows(i).Item(0))
            End If
        Next
        'Cbxquecode1.DataSource = dt1
        'Cbxquecode1.DisplayMember = "code"
        'Cbxquecode2.DataSource = dt2
        'Cbxquecode2.DisplayMember = "code"
        Cbxquecode1.Text = a
        Cbxquecode2.Text = b
        'txtcode.DataSource = dt
        'txtcode.DisplayMember = "code"


    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        watertemp = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        'If Val(DataGridView1.CurrentCell.Value) > 20 Then
        '    DataGridView1.CurrentCell.Value = 20
        'End If
        'DataGridView1.CurrentCell.Value = Format(Val(DataGridView1.CurrentCell.Value), "0.0")
        'If queue0.inprocess = True Then
        '    If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
        '        Dim x As Single = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        '        Dim index As Integer = e.RowIndex
        '        mc.queclone(queue0, queue)
        '        queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
        '        'waterrecount(x, index, queue)
        '        mc.queclone(queue, queue0)
        '        displaymaterial()
        '        'send material

        '        If queue1.inprocess = True Then
        '            mc.queclone(queue1, queue)
        '            'queue = waterrecount(x, index, queue)
        '            queue = mc.Materialdistribute(queue, 1, cbxcorrect.Text, waterall, "B")
        '            mc.queclone(queue, queue1)
        '            displaymaterial()
        '            'send material
        '        End If
        '    Else
        '        DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = watertemp
        '    End If
        'End If
        'DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = False

        'Dim i, j As Integer
        'i = DataGridView1.RowCount
        'Dim fileNum
        'fileNum = FreeFile()
        'FileOpen(fileNum, "\JS\cbc8\files\water.per", OpenMode.Output)
        'For j = 0 To i - 1
        '    PrintLine(fileNum, DataGridView1.Rows(j).Cells(0).Value)
        'Next
        'FileClose(fileNum)

    End Sub


    Public Sub displaymaterial()
        '103.3.11 改回
        '102.12.18 改英制 bImperial 之前 的 displaymaterial()
        Dim i, j As Integer
        Dim sumorg = 0, sumfront = 0
        Dim temp
        Dim DVW_A(25) As Integer 'PB_A SP
        Dim DVW_B(25) As Integer 'PB_B SP
        '101.3.20
        temp = 0

        '將所有材料計算結果顯示到gridview上
        '100.12.4
        ' 100.12.06 Call FrmMonit.TraceLog("displaymaterial()")
        j = 0
        If queue0.Materialfront IsNot Nothing Then
            For i = 0 To 15
                SiloSP(i) = 0
            Next
            'S1~G7
            For i = 0 To 7
                If exist(i) Then
                    '101.3.20
                    Try
                        If queue1.inprocess = True Then
                            '100.7.1
                            'DVW_A(i) = queue1.Materialfront(i)
                            'DVW_B(i) = queue1.Materialback(i)
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                DVW_A(i) = CInt(Val(queue1.Materialfront(i)) * 10)
                                DVW_B(i) = CInt(Val(queue1.Materialback(i)) * 10)
                            Else
                                DVW_A(i) = queue1.Materialfront(i)
                                DVW_B(i) = queue1.Materialback(i)
                            End If
                            '101.4.5
                            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                DVW_A(i) = CInt(Val(queue1.Materialfront(i)) * 10)
                                DVW_B(i) = CInt(Val(queue1.Materialback(i)) * 10)
                            End If
                            '101.3.20
                            'Dim test1, test2 As Integer
                            'test2 = 0
                            'test1 = temp / test2
                        Else
                            'DVW_A(i) = queue0.Materialfront(i)
                            'DVW_B(i) = queue0.Materialback(i)
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                DVW_A(i) = CInt(Val(queue0.Materialfront(i)) * 10)
                                DVW_B(i) = CInt(Val(queue0.Materialback(i)) * 10)
                            Else
                                DVW_A(i) = queue0.Materialfront(i)
                                DVW_B(i) = queue0.Materialback(i)
                            End If
                            '101.4.5
                            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                DVW_A(i) = CInt(Val(queue0.Materialfront(i)) * 10)
                                DVW_B(i) = CInt(Val(queue0.Materialback(i)) * 10)
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "displaymaterial-1")
                    End Try

                    If chk_A_B.Checked Then
                        Try

                            If queue1.inprocess = True Then
                                SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + queue1.Materialback(i)
                                '100.6.30
                                'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0")
                                'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0")
                                If MatType(i) = "S" And bCheckBoxS Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "G" And bCheckBoxG Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "W" And bCheckBoxW Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "C" And bCheckBoxC Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0.0")
                                Else
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0")
                                End If
                                '101.4.5
                                If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialback(i)), "0.0")
                                End If
                                DataGridView2.Rows(0).Cells(i + 1).Style.ForeColor = Color.Yellow
                                DataGridView2.Rows(1).Cells(i + 1).Style.ForeColor = Color.Yellow
                                temp = queue1.Materialback(i)
                            Else
                                SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + queue0.Materialback(i)
                                '100.6.30
                                'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0")
                                'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0")
                                If MatType(i) = "S" And bCheckBoxS Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "G" And bCheckBoxG Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "W" And bCheckBoxW Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                ElseIf MatType(i) = "C" And bCheckBoxC Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                Else
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0")
                                End If
                                '101.4.5
                                If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                End If
                                DataGridView2.Rows(0).Cells(i + 1).Style.ForeColor = Color.Black
                                DataGridView2.Rows(1).Cells(i + 1).Style.ForeColor = Color.Black
                                temp = queue0.Materialback(i)
                            End If
                            'sumorg += DataGridView1.Rows(j).Cells(2).Value
                            'sumfront += DataGridView1.Rows(j).Cells(3).Value
                            '99.12.09
                            sumorg += CSng(DataGridView2.Rows(0).Cells(i + 1).Value)
                            sumfront += DataGridView2.Rows(1).Cells(i + 1).Value
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical, "displaymaterial-2")
                        End Try
                        If MatSiloNo(i) > 0 Then
                            '100.6.30
                            'FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0")
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0.0")
                            Else
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0")
                            End If
                            '101.4.5
                            If (MatType(3) = "S" And bCheckBoxS4) And i = 3 Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0.0")
                            End If
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = MatSiloSR(i)
                        Else
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = 0
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = "."
                        End If
                    Else
                        '101.3.20
                        Try

                            If queue1.inprocess = True Then
                                SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + Val(queue1.Materialfront(i))
                                '100.6.30
                                If MatType(i) = "S" And bCheckBoxS Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "G" And bCheckBoxG Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "W" And bCheckBoxW Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "C" And bCheckBoxC Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0.0")
                                Else
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0")
                                End If
                                '101.4.5
                                If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0.0")
                                End If '100.6.30
                                'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue1.Fmaterialorg(i)), "0")
                                'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue1.Materialfront(i)), "0")
                                DataGridView2.Rows(0).Cells(i + 1).Style.ForeColor = Color.Yellow
                                DataGridView2.Rows(1).Cells(i + 1).Style.ForeColor = Color.Yellow
                                temp = queue1.Materialfront(i)
                            Else
                                SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + Val(queue0.Materialfront(i))
                                '100.6.30
                                If MatType(i) = "S" And bCheckBoxS Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "G" And bCheckBoxG Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "W" And bCheckBoxW Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                ElseIf MatType(i) = "C" And bCheckBoxC Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                Else
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0")
                                End If
                                '101.4.5
                                If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                    DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                    DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                End If
                                '100.6.30
                                'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0")
                                'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0")
                                DataGridView2.Rows(0).Cells(i + 1).Style.ForeColor = Color.Black
                                DataGridView2.Rows(1).Cells(i + 1).Style.ForeColor = Color.Black
                                temp = queue0.Materialfront(i)
                            End If
                            'sumorg += DataGridView1.Rows(j).Cells(2).Value
                            'sumfront += DataGridView1.Rows(j).Cells(3).Value
                            sumorg += CSng(DataGridView2.Rows(0).Cells(i + 1).Value)
                            sumfront += DataGridView2.Rows(1).Cells(i + 1).Value
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Critical, "displaymaterial-3")
                        End Try
                        If MatSiloNo(i) > 0 Then
                            '100.6.30
                            'FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0")
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0.0")
                            Else
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0")
                            End If
                            '101.4.5
                            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(temp), "0.0")
                            End If
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = MatSiloSR(i)
                        Else
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = 0
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = "."
                        End If
                    End If
                    If queue1.ismatch = True Then
                        'DataGridView1.Rows(j).Cells(2).Value = Format(Val(queue1.Bmaterialorg(i)), "0")
                    End If
                    j = j + 1
                End If
            Next
            j = 0
            For i = 8 To 21
                If exist(i) Then
                    If i < 17 Then
                        '100.7.1
                        'DVW_A(i) = CInt(queue0.Materialfront(i))
                        'DVW_B(i) = CInt(queue0.Materialback(i))
                        If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                            DVW_A(i) = CInt(Val(queue0.Materialfront(i)) * 10)
                            DVW_B(i) = CInt(Val(queue0.Materialback(i)) * 10)
                        Else
                            DVW_A(i) = CInt(queue0.Materialfront(i))
                            DVW_B(i) = CInt(queue0.Materialback(i))
                        End If
                    Else
                        DVW_A(i) = CInt(queue0.Materialfront(i) * 100)
                        DVW_B(i) = CInt(queue0.Materialback(i) * 100)
                    End If
                    If chk_A_B.Checked Then
                        If i < 17 Then
                            '100.6.30
                            If MatType(i) = "S" And bCheckBoxS Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            ElseIf MatType(i) = "G" And bCheckBoxG Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            ElseIf MatType(i) = "W" And bCheckBoxW Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            ElseIf MatType(i) = "C" And bCheckBoxC Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            Else
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0")
                            End If
                            '101.4.5
                            If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            End If
                            '100.6.30
                            'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0")
                            'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0")
                            temp = Format(queue0.Materialback(i), "0")
                        Else
                            'DataGridView1.Rows(j).Cells(5).Value = Format(Val(queue0.Bmaterialorg(i)), "0.00")
                            'DataGridView1.Rows(j).Cells(6).Value = Format(Val(queue0.Materialback(i)), "0.00")
                            DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Bmaterialorg(i)), "0.00")
                            DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialback(i)), "0.00")
                            temp = Format(Val(queue0.Materialback(i)), "0.00")
                        End If
                        'sumorg += DataGridView1.Rows(j).Cells(5).Value
                        'sumfront += DataGridView1.Rows(j).Cells(6).Value
                        sumorg += CSng(DataGridView2.Rows(0).Cells(i + 1).Value)
                        sumfront += DataGridView2.Rows(1).Cells(i + 1).Value
                        SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + queue0.Materialback(i)
                        If MatSiloNo(i) > 8 Then
                            If i < 17 Then
                                '100.6.30
                                'FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0")
                                If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                Else
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0")
                                End If
                                '101.4.5
                                If (MatType(i) = "S" And bCheckBoxS And i = 3) Then
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0.0")
                                End If
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 - 3).Value = MatSiloSR(i)
                            Else
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0.00")
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 - 3).Value = MatSiloSR(i)
                            End If
                        ElseIf MatSiloNo(i) <= 8 And MatSiloNo(i) > 0 Then
                            '100.7.1
                            'FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0")
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0")
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            Else
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0")
                            End If
                            '101.4.5
                            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialback(i)), "0.0")
                            End If
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = MatSiloSR(i)
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(queue0.Materialback(i), "0")
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 - 3).Value = MatSiloSR(i)
                        Else
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = 0
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 - 3).Value = "."
                        End If
                    Else
                        If i < 17 Then
                            '100.6.30
                            If MatType(i) = "S" And bCheckBoxS Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            ElseIf MatType(i) = "G" And bCheckBoxG Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            ElseIf MatType(i) = "W" And bCheckBoxW Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            ElseIf MatType(i) = "C" And bCheckBoxC Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            Else
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0")
                            End If
                            '101.4.5
                            If MatType(i) = "S" And bCheckBoxS4 And i = 3 Then
                                DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.0")
                                DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            End If
                            '100.6.30
                            'DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0")
                            'DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0")
                        Else
                            'DataGridView1.Rows(j).Cells(5).Value = Format(Val(queue0.Fmaterialorg(i)), "0.00")
                            'DataGridView1.Rows(j).Cells(6).Value = Format(Val(queue0.Materialfront(i)), "0.00")
                            DataGridView2.Rows(0).Cells(i + 1).Value = Format(Val(queue0.Fmaterialorg(i)), "0.00")
                            DataGridView2.Rows(1).Cells(i + 1).Value = Format(Val(queue0.Materialfront(i)), "0.00")
                        End If
                        'sumorg += DataGridView1.Rows(j).Cells(5).Value
                        'sumfront += DataGridView1.Rows(j).Cells(6).Value
                        sumorg += CSng(DataGridView2.Rows(0).Cells(i + 1).Value)
                        sumfront += DataGridView2.Rows(1).Cells(i + 1).Value
                        SiloSP(MatSiloNo(i)) = SiloSP(MatSiloNo(i)) + Val(queue0.Materialfront(i))
                        If MatSiloNo(i) > 8 Then
                            If i < 17 Then
                                '100.6.30
                                'FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0")
                                If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                Else
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0")
                                End If
                                '101.4.5
                                If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                    FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                                End If
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 - 3).Value = MatSiloSR(i)
                            Else
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0.00")
                                FrmMonit.dgvSilo2.Rows(6 - MatSiloOrder(i)).Cells((MatSiloNo(i) - 8) * 3 - 3).Value = MatSiloSR(i)
                            End If
                        ElseIf MatSiloNo(i) <= 8 And MatSiloNo(i) > 0 Then
                            '100.6.30
                            'FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0")
                            If (MatType(i) = "S" And bCheckBoxS) Or (MatType(i) = "G" And bCheckBoxG) Or (MatType(i) = "W" And bCheckBoxW) Or (MatType(i) = "C" And bCheckBoxC) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            Else
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0")
                            End If
                            '101.4.5
                            If (MatType(i) = "S" And bCheckBoxS4 And i = 3) Then
                                FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = Format(Val(queue0.Materialfront(i)), "0.0")
                            End If
                            FrmMonit.dgvSilo1.Rows(6 - MatSiloOrder(i)).Cells(MatSiloNo(i) * 3 - 3).Value = MatSiloSR(i)
                        Else
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 + 1 - 3).Value = 0
                            'FrmMonit.dgvSilo1.Rows(MatSiloOrder(i) + 2).Cells(MatSiloNo(i) * 3 - 3).Value = "."
                        End If
                    End If
                    j += 1
                End If
            Next
            '102.5.11 MOVE TO quecode()
            If bUPDATEM3_1 = True Then
                LabelFormula1M3.Text = Format(sumorg, "0")
                bUPDATEM3_1 = False
            End If
            If fCubeLast <= 0 Then
                fCubeLast = 1.0
            End If
            If chk_A_B.Checked Then
                sumfront = fSumFrontA
            Else
                sumfront = fSumFrontB
            End If
            LabelFormula.Text = Format(sumfront, "0")
            'LabelRealTot1M3.Text = Format(sumorg, "0")
            'LabelRealTot.Text = Format(sumfront, "0")
        End If
        'Ray
        'For i = 0 To 7
        '    DataGridView2.Rows(0).Cells(i + 1).Value = DataGridView1.Rows(i).Cells(2).Value
        '    DataGridView2.Rows(1).Cells(i + 1).Value = DataGridView1.Rows(i).Cells(3).Value
        'Next
        'For i = 8 To 21
        '    DataGridView2.Rows(0).Cells(i + 1).Value = DataGridView1.Rows(i - 8).Cells(5).Value
        '    DataGridView2.Rows(1).Cells(i + 1).Value = DataGridView1.Rows(i - 8).Cells(6).Value
        'Next

        If FrmMonit.dgvSilo1.ColumnCount > 5 Then
            For i = 0 To 7
                '100.6.30
                'FrmMonit.dgvSilo1.Rows(1).Cells(i * 3 + 1).Value = Format(SiloSP(i + 1), "0")
                If (SiloMatType(i) = "S" And bCheckBoxS) Or (SiloMatType(i) = "G" And bCheckBoxG) Or (SiloMatType(i) = "W" And bCheckBoxW) Or (SiloMatType(i) = "C" And bCheckBoxC) Then
                    FrmMonit.dgvSilo1.Rows(1).Cells(i * 3 + 1).Value = Format(SiloSP(i + 1), "0.0")
                Else
                    FrmMonit.dgvSilo1.Rows(1).Cells(i * 3 + 1).Value = Format(SiloSP(i + 1), "0")
                End If
                '101.4.8
                If (SiloMatType(i) = "S" And bCheckBoxS4 And i = S4SiloNo) Then
                    FrmMonit.dgvSilo1.Rows(1).Cells(i * 3 + 1).Value = Format(SiloSP(i + 1), "0.0")
                End If
            Next
            For i = 8 To 14
                If SiloMatType(i) = "A" Then
                    FrmMonit.dgvSilo2.Rows(1).Cells((i - 8) * 3 + 1).Value = Format(SiloSP(i + 1), "0.00")
                Else
                    '100.6.30
                    'FrmMonit.dgvSilo2.Rows(1).Cells((i - 8) * 3 + 1).Value = Format(SiloSP(i + 1), "0")
                    If (SiloMatType(i) = "S" And bCheckBoxS) Or (SiloMatType(i) = "G" And bCheckBoxG) Or (SiloMatType(i) = "W" And bCheckBoxW) Or (SiloMatType(i) = "C" And bCheckBoxC) Then
                        FrmMonit.dgvSilo2.Rows(1).Cells((i - 8) * 3 + 1).Value = Format(SiloSP(i + 1), "0.0")
                    Else
                        FrmMonit.dgvSilo2.Rows(1).Cells((i - 8) * 3 + 1).Value = Format(SiloSP(i + 1), "0")
                    End If
                    '101.4.14
                    If (SiloMatType(i) = "S" And bCheckBoxS4 And i = S4SiloNo) Then
                        FrmMonit.dgvSilo2.Rows(1).Cells((i - 8) * 3 + 1).Value = Format(SiloSP(i + 1), "0.0")
                    End If
                End If
            Next
        End If
        For i = 0 To 7
            DVW(i + 741) = DVW_A(i)
            DVW(i + 771) = DVW_B(i)
        Next
        'water D750
        For i = 7 To 23
            DVW(i + 741 + 1) = DVW_A(i)
            DVW(i + 771 + 1) = DVW_B(i)
        Next

        '110.7.10 110.9.11
        'If queue0.ModDoingPlate >= 1 Then
        If Not (queue0.plate Is Nothing) Then
            Dim q0pl, q0mp
            q0pl = queue0.plate.Length
            q0mp = queue0.ModDoingPlate
            If queue0.ModDoingPlate >= 1 And queue0.plate.Length >= queue0.ModDoingPlate Then
                For i = 0 To 5
                    '110.7.10
                    'DVW_C16(i) = CInt((CType(queue0.Bmaterialorg(i + 11), Single) * queue0.plate_Bsub(queue0.ModDoingPlate - 1)))
                    DVW_C16(i) = CInt((CType(queue0.Bmaterialorg(i + 11), Single) * queue0.plate(queue0.ModDoingPlate - 1)))
                Next
            End If
        End If
        '110.7.7    C1-6
        DVW(749) = DVW_C16(0)
        DVW(764) = DVW_C16(1)
        DVW(765) = DVW_C16(2)
        DVW(779) = DVW_C16(3)
        DVW(794) = DVW_C16(4)
        DVW(795) = DVW_C16(5)

        '105.5.20 AE15
        DVW(766) = queue0.Fmaterialnowater(24)
        DVW(796) = queue0.Bmaterialnowater(24)

        For i = 0 To 3
            DVW(i + 767) = queue0.ToPLC_A(i)
            DVW(i + 797) = queue0.ToPLC_B(i)
        Next
        i = FrmMonit.WriteA2N(741, 30)
        i = FrmMonit.WriteA2N(771, 30)
        '102.7.19 test for W SP error
        LabelW.Text = LabelW2.Text
        LabelW2.Text = LabelW3.Text
        LabelW3.Text = LabelW4.Text
        LabelW4.Text = "W1: " & DVW(780)

        '102.7.15
        If queue0.plate Is Nothing Then Exit Sub
        If Not chk_A_B.Checked Then
            For i = 0 To queue0.plate.Length - 1
                dgvPan.Rows(0).Cells(i + 1).Value = Format(Val(queue0.plate(i)), "0.00")
                dgvPan.Rows(1).Cells(i + 1).Value = Format(Val(queue0.plate(i)), "0.00")
            Next
        Else
            For i = 0 To queue0.plate.Length - 1
                dgvPan.Rows(0).Cells(i + 1).Value = Format(Val(queue0.plate_Bsub(i)), "0.00")
                dgvPan.Rows(1).Cells(i + 1).Value = Format(Val(queue0.plate_Bsub(i)), "0.00")
            Next
        End If

        '105.5.20
        Dim memo3, k As Integer
        Dim A1Flag As Boolean
        A1Flag = False
        memo3 = queue0.Fmaterialnowater(24)
        If bCheckBoxSP1 Then
            If Not chk_A_B.Checked Then
                k = queue0.Fmaterialnowater(24)
            Else
                k = queue0.Bmaterialnowater(24)
            End If
            For i = 0 To 15
                If (k Mod 2) = 1 Then
                    If A1Flag Then
                        '105.5.24
                        'MatName(18) = strAE15Name(i + 1)
                        'DataGridView2.Columns(19).HeaderText = strAE15Name(i + 1)
                        MatName(19) = strAE15Name(i + 1)
                        DataGridView2.Columns(20).HeaderText = strAE15Name(i + 1)
                    Else
                        'MatName(17) = strAE15Name(i + 1)
                        'DataGridView2.Columns(18).HeaderText = strAE15Name(i + 1)
                        MatName(18) = strAE15Name(i + 1)
                        DataGridView2.Columns(19).HeaderText = strAE15Name(i + 1)
                        A1Flag = True
                    End If
                End If
                k = k \ 2
            Next
        End If


    End Sub
    Private Sub cbxcorrect_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxcorrect.SelectionChangeCommitted
        '99.08.02 disable only by B area
        '100.5.15 ref dgcwater
        If queue0.inprocess = True Then
            queue0.corr_B = cbxcorrect.Text
            LabelArea.Text = cbxcorrect.Text
            Dim k As Integer
            For k = 0 To 7
                waterall(k) = dgvWater.Rows(0).Cells(k).Value
            Next
            mc.queclone(queue0, queue)
            queue.BoneDoingPlate = queue0.BoneDoingPlate
            queue.ModDoingPlate = queue0.ModDoingPlate
            If queue0.BoneDoingPlate = queue0.ModDoingPlate Then
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
            Else
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
            End If

            mc.queclone(queue, queue0)
            MonLog.add = "displaymaterial() : cbxcorrect,queue0"
            displaymaterial()

            '101.3.20
            Try

                If queue1.inprocess = True Then
                    mc.queclone(queue1, queue)
                    queue = mc.Materialdistribute(queue, queue1.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
                    mc.queclone(queue, queue1)
                    MonLog.add = "displaymaterial() : cbxcorrect,queue1"
                    displaymaterial()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "correct-1")
            End Try
        End If
    End Sub

    Private Sub btnbegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbegin.Click
        '110.6.5 亞東 水泥確認
        LabelCement.Visible = False
        TimerShowCement.Enabled = False

        ''105.5.6
        'Tbxquefield2.Text = Tbxquefield1.Left
        Tbxquefield1.Left = lblfield.Left
        Tbxquefield2.Left = lblfield.Left
        LabelTime.Top = 8

        '99.08.07
        queue1check()
        If queue0.inprocess = False And queue1.ismatch = True Then
            '進入下一batch生產
            i = queue1.Fmaterialnowater(24)
            '107.6.14
            MainLog.add = "CheckDefaultWater() : btnbegin quecode()"
            quecode()

            '107.5.10
            'mc.queclone(queue0, queue)
            'i = queue0.Fmaterialnowater(24)
            'queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
            'mc.queclone(queue, queue0)
            ''100.12.6   107.3.17
            ''Call FrmMonit.TraceLog("displaymaterial() : btnbegin")
            'MainLog.add = "displaymaterial() : btnbegin pb= " & queue0.fomulacode & " , m3= " & queue0.tritotal.ToString
            'displaymaterial()

            '110.9.24 生產完畢開始計時
            'dtBatDon = Now

            btnbegin.BackColor = Color.Cyan

            '106.10.7 聯佶空單
            If bCheckBoxSP13 And lblwork.ForeColor = Color.Green Then
                Dim i
                i = MsgBox("確認自動模擬回傳?", MsgBoxStyle.YesNo, "確認自動模擬回傳")
                If i = vbYes Then
                    MixynIng = True
                    '107.3.20
                    TLog.add = " 自動模擬回傳"
                    DV(20) = 0
                    DV(21) = 0
                    '107.3.3 ??? if last PAN was not done?
                    DV(22) = 0
                    dtp_simtime.Text = LabelPb12_0.Text
                    '106.10.11
                    For i = 0 To 24
                        DVW(i + 91) = 0
                    Next
                    i = FrmMonit.WriteA2N(91, 25)
                Else
                    '106.10.11 聯佶空單
                    clearproduce()
                    Exit Sub
                End If
                Dim generator As New Random
                Dim randomValue As Integer
                Dim s As String
                dtSimP(0) = dtp_simtime.Value
                If queue0.plate.Length > 1 Then
                    For i = 1 To queue0.plate.Length - 1
                        randomValue = generator.Next(1, 10)
                        s = CInt(tbxmixtime.Text) + Num_quetime_set.Value + randomValue
                        dtSimP(i) = dtSimP(i - 1).AddSeconds(CInt(s))
                    Next
                End If
                Timerprocess_sim.Enabled = True
            End If
        End If

        If LabelArea.Visible = True Then LabelArea.Text = queue0.corr_B

        '107.6.14 
        '   move below to quecode() since also should do when ModDoneChange by auto
        '   => and use CheckDefaultWater() instead of these code
        '104.8.8
        'CreateTrendFile(queue0.showdata(1) & "_" & TbxworkTri.Text)
        'FrmMonit.ReadTrendFile(queue0.showdata(1) & "_" & TbxworkTri.Text)
        '104.8.9 
        'check if PB,Field in strWatPb(15),strWatField(15) , 
        'if YES 
        '   NumW1~NumW8 change to fWat() the matched record
        'bDefaultWatEn = False
        'Dim j
        'For i = 0 To 14
        '    If strWatPb(i) = queue0.showdata(1) Then
        '        '107.6.11
        '        'If strWatField(i) = tbxworkfield.Text Then
        '        If strWatField(i) = tbxworkfield.Text Or (tbxworkfield.Text = "" And strWatField(i) = ".") Then
        '            If Not bDefaultWatEn_Last Then
        '                For j = 0 To 7
        '                    fWatOrig(j) = dgvWater.Rows(0).Cells(j).Value
        '                Next
        '            End If
        '            '107.6.2 CHECK BEFORE UPDATE TO dgvWater since set to PLC
        '            '  check each NumW , if not = fWat change
        '            If NumW1.Value <> fWat(0, i) Then
        '                NumW1.Value = fWat(0, i)
        '                dgvWater.Rows(0).Cells(0).Value = NumW1.Value
        '                NumW1.BackColor = Color.SkyBlue
        '            End If
        '            '107.6.9 fix bug
        '            If Numw2.Value <> fWat(1, i) Then
        '                Numw2.Value = fWat(1, i)
        '                dgvWater.Rows(0).Cells(1).Value = Numw2.Value
        '                Numw2.BackColor = Color.SkyBlue
        '            End If
        '            If Numw3.Value <> fWat(2, i) Then
        '                Numw3.Value = fWat(2, i)
        '                dgvWater.Rows(0).Cells(2).Value = Numw3.Value
        '                Numw3.BackColor = Color.SkyBlue
        '            End If
        '            If Numw4.Value <> fWat(3, i) Then
        '                Numw4.Value = fWat(3, i)
        '                dgvWater.Rows(0).Cells(3).Value = Numw4.Value
        '                Numw4.BackColor = Color.SkyBlue
        '            End If
        '            If Numw5.Value <> fWat(4, i) Then
        '                Numw5.Value = fWat(4, i)
        '                dgvWater.Rows(0).Cells(4).Value = Numw5.Value
        '                Numw5.BackColor = Color.SkyBlue
        '            End If
        '            If Numw6.Value <> fWat(5, i) Then
        '                Numw6.Value = fWat(5, i)
        '                dgvWater.Rows(0).Cells(5).Value = Numw6.Value
        '                Numw6.BackColor = Color.SkyBlue
        '            End If
        '            If Numw7.Value <> fWat(6, i) Then
        '                Numw7.Value = fWat(6, i)
        '                dgvWater.Rows(0).Cells(6).Value = Numw7.Value
        '                Numw7.BackColor = Color.SkyBlue
        '            End If
        '            If Numw8.Value <> fWat(7, i) Then
        '                Numw8.Value = fWat(7, i)
        '                dgvWater.Rows(0).Cells(7).Value = Numw8.Value
        '                Numw8.BackColor = Color.SkyBlue
        '            End If
        '            bDefaultWatEn = True
        '            iDefaultWatInd = i
        '            '107.6.11
        '            Exit For
        '        End If
        '    End If
        'Next

        'If bDefaultWatEn_Last And Not bDefaultWatEn Then
        '    '107.6.11
        '    If NumW1.Value <> fWatOrig(0) Then
        '        NumW1.Value = fWatOrig(0)
        '        dgvWater.Rows(0).Cells(0).Value = NumW1.Value
        '        NumW1.BackColor = Color.SkyBlue
        '    End If
        '    If Numw2.Value <> fWatOrig(1) Then
        '        Numw2.Value = fWatOrig(1)
        '        dgvWater.Rows(0).Cells(1).Value = Numw2.Value
        '        Numw2.BackColor = Color.SkyBlue
        '    End If
        '    If Numw3.Value <> fWatOrig(2) Then
        '        Numw3.Value = fWatOrig(2)
        '        dgvWater.Rows(0).Cells(2).Value = Numw3.Value
        '        Numw3.BackColor = Color.SkyBlue
        '    End If
        '    If Numw4.Value <> fWatOrig(3) Then
        '        Numw4.Value = fWatOrig(3)
        '        dgvWater.Rows(0).Cells(3).Value = Numw4.Value
        '        Numw4.BackColor = Color.SkyBlue
        '    End If
        '    If Numw5.Value <> fWatOrig(4) Then
        '        Numw5.Value = fWatOrig(4)
        '        dgvWater.Rows(0).Cells(4).Value = Numw5.Value
        '        Numw5.BackColor = Color.SkyBlue
        '    End If
        '    If Numw6.Value <> fWatOrig(5) Then
        '        Numw6.Value = fWatOrig(5)
        '        dgvWater.Rows(0).Cells(5).Value = Numw6.Value
        '        Numw6.BackColor = Color.SkyBlue
        '    End If
        '    If Numw7.Value <> fWatOrig(6) Then
        '        Numw7.Value = fWatOrig(6)
        '        dgvWater.Rows(0).Cells(6).Value = Numw7.Value
        '        Numw7.BackColor = Color.SkyBlue
        '    End If
        '    If Numw8.Value <> fWatOrig(7) Then
        '        Numw8.Value = fWatOrig(7)
        '        dgvWater.Rows(0).Cells(7).Value = Numw8.Value
        '        Numw8.BackColor = Color.SkyBlue
        '    End If
        'End If
        'bDefaultWatEn_Last = bDefaultWatEn
        'If bDefaultWatEn Then
        '    lblfield.ForeColor = Color.Red
        'Else
        '    lblfield.ForeColor = Color.Blue
        'End If

        '104.9.28
        If strChangedPb <> "" Then
            If bEngVer Then
                MsgBox("Formula (" & strChangedPb & ") changed,please confirm!", MsgBoxStyle.Critical)
            Else
                MsgBox("此配比(" & strChangedPb & ")內容已更改,請檢查是否正確!", MsgBoxStyle.Critical)
            End If
            strChangedPb = ""
        End If


    End Sub

    Private Sub txtbone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtbone.TextChanged

        TxtMessage.Text = TxtMessage.Text & vbCrLf & Now() & "骨材盤次:" & txtbone.Text
        If CType(txtbone.Text, Integer) <= 0 Or queue0.inprocess = False Then Exit Sub
        '繪製盤數圖形
        'If 骨盤盤數 = 生產總盤數, 生產模式為自動 且排隊中符合生產規定 ==> 搶先送骨材

    End Sub

    Private Sub Txtmod_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txtmod.TextChanged

        TxtMessage.Text = TxtMessage.Text & vbCrLf & Now() & "泥材盤次:" & Txtmod.Text
        If CType(Txtmod.Text, Integer) <= 0 Then Exit Sub
        If queue0.ModDonePlate <= 0 Then Exit Sub

    End Sub

    Private Sub clearproduce()
        '將盤數清空
        Dim i As Integer
        For i = 0 To 8
            dgvPan.Rows(0).Cells(i + 1).Value = ""
            dgvPan.Rows(0).Cells(i + 1).Style.BackColor = Color.LightYellow
            dgvPan.Rows(1).Cells(i + 1).Value = ""
            dgvPan.Rows(1).Cells(i + 1).Style.BackColor = Color.LightYellow
        Next
        '106.1.3

        tbxworkcode.Text = ""
        tbxworkcode.Text = ""
        TbxworkTri.Text = ""
        Tbxworkcar.Text = ""
        tbxworkfield.Text = ""
        lblCustomerNo.Text = ""
        lblCustomerName.Text = ""
        LabelQkey0.Text = " "
        LabelQkey.Text = " "
        LabelCust.Text = " "
        '106.10.7
        lblwork.ForeColor = Color.Blue
        LabelPb12_0.Visible = False

        DVW(20) = 0
        DVW(21) = 0
        DVW(22) = 0
        i = FrmMonit.WriteA2N(20, 3)
        '107.3.18
        'Call SavingLog("clearproduce reset D20,21,22 BON= " & queue0.BoneDonePlate & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate)
        MainLog.add = "clearproduce() : reset D20,21,22 BON= " & queue0.BoneDonePlate & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate

        If Not bCommFlag Then
            DV(20) = 0
            DV(21) = 0
            DV(22) = 0
        End If

        queue0.clear()
        showotherdata(queue0.showdata)
        MonLog.add = "displaymaterial() : clearproduce"
        displaymaterial()

    End Sub

    Private Sub btnboneadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnboneadd.Click
        If queue0.inprocess = False Then Exit Sub
        Dim i
        '107.3.6
        DVW(20) = DV(20) + 1
        '107.6.5    6.14
        'DVW(21) = DV(21) + 1
        'i = FrmMonit.WriteA2N(20, 2)
        i = FrmMonit.WriteA2N(20, 1)
        'DVW(22) = DV(22) + 1

        '107.3.20
        'DVW(20) = DV(20) + 1
        'DVW(21) = DVW(20)
        ''107.3.17
        'DVW(22) = DVW(20)
        'Dim i
        ''107.3.17
        ''i = FrmMonit.WriteA2N(20, 2)
        'i = FrmMonit.WriteA2N(20, 3)
    End Sub

    Private Sub btnmodadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnmodadd.Click
        If queue0.inprocess = False Then Exit Sub
        Dim i
        '107.3.6 107.4.12 107.5.11
        '110.6.23 test
        DVW(21) = DV(21) + 1
        i = FrmMonit.WriteA2N(21, 1)
        'DVW(21) = DV(21) + 1
        'DVW(22) = DVW(21)
        'i = FrmMonit.WriteA2N(21, 2)
    End Sub

    Public Sub Numquetri1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numquetri1.Leave
        '109.11.29 
        'If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方

        If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
            '110.6.7 車號簡檢查
            'If Not (Tbxquecar1.Text <> "" And bPbSync) Then
            '110.9.15
            'If Not (Tbxquecar1.Text <> "" And bCheckBoxYadon) Then
            '    queue1check()
            'End If
            'If (Not (Tbxquecar1.Text <> "" And bCheckBoxYadon)) Or (Not bCheckBoxYadon) Then
            '    queue1check()
            'End If
            queue1check()
        End If
    End Sub

    Private Sub Cbxquecode1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbxquecode1.Leave
        '109.11.29 
        'If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
        If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
            '110.6.7 車號簡檢查
            'If Not (Tbxquecar1.Text <> "" And bPbSync) Then
            '110.9.15
            'If Not (Tbxquecar1.Text <> "" And bCheckBoxYadon) Then
            '    queue1check()
            'End If
            'If Not (Tbxquecar1.Text <> "" And bCheckBoxYadon) Or Not bCheckBoxYadon Then
            '    queue1check()
            'End If
            queue1check()
        End If
    End Sub

    Public Sub queue1check()

        '109.11.29 109.12.23
        'If bPbSync Then
        '110.6.7 車號簡檢查
        'If bPbSync And Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0 Then
        If bCheckBoxYadon And Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0 Then
            If Tbxquecar1.Text.Length = 0 Then
                MsgBox("請輸入車號!")
                Tbxquecar1.Focus()
            Else
                If Not Char.IsLetterOrDigit(Tbxquecar1.Text.Substring(0, 1)) Then
                    MsgBox("請輸入車號!")
                    Tbxquecar1.Text = ""
                    Tbxquecar1.Focus()
                End If
            End If
        End If
        'If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
        '110.6.7 車號簡檢查
        'If ((Not (Tbxquecar1.Text = "" And bPbSync)) And Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
        If ((Not (Tbxquecar1.Text = "" And bCheckBoxYadon)) And Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
            '111.3.28
            Cbxquecode1.Text = Cbxquecode1.Text.ToUpper
            queue1.fomulacode = Cbxquecode1.Text
            queue1.tritotal = Convert.ToSingle(numquetri1.Text)

            '106.10.7 聯佶空單
            If bCheckBoxSP13 Then
                If btnque1.ForeColor = Color.Green Then
                    '106.10.20
                    'queue1.MixYn = "Y"
                    queue1.MixYn = "N"
                Else
                    '106.10.20
                    'queue1.MixYn = "N"
                    queue1.MixYn = "Y"
                End If
            End If

            '計算waterall
            Dim j = -1
            Dim i As Integer
            For i = 0 To 7
                waterall(i) = dgvWater.Rows(0).Cells(i).Value
                If exist(i) Then
                    j += 1
                    'waterall(i) = DataGridView1.Rows(j).Cells(0).Value
                    'waterall(i) = dgvWater.Rows(0).Cells(i).Value
                End If
            Next

            mc.queclone(queue1, queue)
            queue = mc.checkque(queue, waterall)


            mc.queclone(queue, queue1)
            If queue1.ismatch = False Then
                Cbxquecode1.BackColor = Color.Red
                numquetri1.BackColor = Color.Red
                'lblmsgque1.Text = "配方錯誤, 請檢查配方"
                '103.1.22
                'MsgBox("配方錯誤, 請檢查配方")
                If bEngVer Then
                    MsgBox("Wrong formula, Please check!")
                Else
                    MsgBox("配方錯誤, 請檢查配方")
                End If

                '104.11.24
                queue1.clear()
                mc.queclone(queue0, queue)

            Else
                Cbxquecode1.BackColor = Color.Aquamarine
                numquetri1.BackColor = Color.Aquamarine
                'lblmsgque1.Text = "可以生產"
                '110.11.12 @home
                '110.11.16
                'queue1.TotalPlate = queue.TotalPlate
                Call UpdateLabelTotPlate()
            End If
        Else
            queue1.clear()
            Cbxquecode1.BackColor = Color.White
            numquetri1.BackColor = Color.White
            '101.12.22
            queue1.ismatch = False
        End If
    End Sub



    Private Sub Timerprocess_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timerprocess.Tick
        Timerprocess.Enabled = False

    End Sub

    Public Sub quecode()
        '進入下一batch生產
        Dim i = 0
        Dim j = 0
        '100.12.4
        '107.3.17
        'Call FrmMonit.TraceLog("quecode()")
        TLog.add = "quecode()"
        If queue1.ismatch = True Then
            '排隊一 check OK
            ' !!!clearproduce()
            mc.queclone(queue1, queue0)
            '110.11.12 @home 110.11.16
            'queue0.TotalPlate = queue1.TotalPlate
            Call UpdateLabelTotPlate()

            tbxworkcode.Text = queue0.fomulacode
            TbxworkTri.Text = queue0.tritotal
            '107.8.21 limit lenth to 12
            'Tbxworkcar.Text = Tbxquecar1.Text
            '112.9.18 If Tbxquefield1.Text.Length > 12 Then
            If Tbxquecar1.Text.Length > 12 Then
                Tbxworkcar.Text = Microsoft.VisualBasic.Left(Tbxquecar1.Text, 12)
            Else
                Tbxworkcar.Text = Tbxquecar1.Text
            End If
            '107.8.21 limit lenth to 12
            'tbxworkfield.Text = Tbxquefield1.Text
            If Tbxquefield1.Text.Length > 12 Then
                tbxworkfield.Text = Microsoft.VisualBasic.Left(Tbxquefield1.Text, 12)
            Else
                tbxworkfield.Text = Tbxquefield1.Text
            End If
            '100.5.15
            'cbxcorrect.Text = queue0.showdata(0)
            cbxcorrect.Text = queue0.corr_B
            tbxmixtime.Text = queue0.showdata(3)
            FrmMonit.tbxmixtime.Text = tbxmixtime.Text
            lblwork.BackColor = Color.LightPink
            LabelQkey0.Text = LabelQkey1.Text
            LabelCust.Text = LabelCust1.Text

            Cbxquecode1.Text = Cbxquecode2.Text
            numquetri1.Text = numquetri2.Text
            Tbxquecar1.Text = Tbxquecar2.Text
            Tbxquefield1.Text = Tbxquefield2.Text
            LabelQkey1.Text = LabelQkey2.Text
            LabelCust1.Text = LabelCust2.Text

            lblCustomerNo.Text = tbxCusNo1.Text
            tbxCusNo1.Text = tbxCusNo2.Text
            lblCustomerName.Text = lblCusName1.Text
            lblCusName1.Text = lblCusName2.Text

            '106.10.7 聯佶空單
            If bCheckBoxSP13 Then
                btnque1.ForeColor = btnque2.ForeColor
                LabelPb12_0.Text = LabelPb12_1.Text
                LabelPb12_1.Text = LabelPb12_2.Text
                If btnque1.ForeColor = Color.Green Then
                    LabelPb12_1.Visible = True
                Else
                    LabelPb12_1.Visible = False
                End If
                '106.10.20
                'If queue0.MixYn = "Y" Or queue0.MixYn = "y" Then
                If queue0.MixYn = "N" Or queue0.MixYn = "n" Then
                    lblwork.ForeColor = Color.Green
                    LabelPb12_0.Visible = True
                Else
                    lblwork.ForeColor = Color.Blue
                    LabelPb12_0.Visible = False
                End If
            End If

            '99.09.28
            LabelCust1.Text = LabelCust2.Text
            queue0.inprocess = True
            queue1.inprocess = False

            Cbxquecode1.BackColor = Color.White
            numquetri1.BackColor = Color.White
            If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
                '111.3.28
                Cbxquecode1.Text = Cbxquecode1.Text.ToUpper
                queue1.fomulacode = Cbxquecode1.Text
                queue1.tritotal = Convert.ToSingle(numquetri1.Text)
                queue1check()
            Else
                queue1.clear()
            End If

            '100.2.24
            '100.3.20 change back
            Cbxquecode2.Text = DataGridViewQue.Rows(0).Cells(0).Value.ToString
            numquetri2.Text = DataGridViewQue.Rows(0).Cells(1).Value.ToString
            Tbxquecar2.Text = DataGridViewQue.Rows(0).Cells(2).Value.ToString
            Tbxquefield2.Text = DataGridViewQue.Rows(0).Cells(3).Value.ToString
            tbxCusNo2.Text = DataGridViewQue.Rows(0).Cells(4).Value.ToString
            lblCusName2.Text = DataGridViewQue.Rows(0).Cells(5).Value.ToString
            LabelCust2.Text = DataGridViewQue.Rows(0).Cells(6).Value.ToString
            LabelQkey2.Text = DataGridViewQue.Rows(0).Cells(7).Value.ToString

            '107.5.30 move from below
            '100.6.18
            Cbxquecode1.Enabled = True
            numquetri1.Enabled = True
            Tbxquecar1.Enabled = True
            Tbxquefield1.Enabled = True
            btnque1.Enabled = True
            btncancel1.Enabled = True
            '101.2.25
            btnque2.Enabled = True
            btncancel2.Enabled = True
            '107.5.30
            Cbxquecode1.Refresh()
            numquetri1.Refresh()

            '106.11.2a
            '106.10.7 聯佶空單
            'btnque2.ForeColor = Color.Black
            'LabelPb12_2.Visible = False
            '106.11.2 聯佶空單
            If bCheckBoxSP13 Then
                Dim MixYn
                MixYn = DataGridViewQue.Rows(0).Cells(8).Value.ToString
                LabelPb12_2.Text = DataGridViewQue.Rows(0).Cells(9).Value.ToString
                If MixYn = "N" Or MixYn = "n" Then
                    btnque2.ForeColor = Color.Green
                    LabelPb12_2.Visible = True
                Else
                    btnque2.ForeColor = Color.Black
                    LabelPb12_2.Visible = False
                End If
            End If

            DataGridViewQue.Rows(0).Cells(0).Value = DataGridViewQue.Rows(1).Cells(0).Value
            DataGridViewQue.Rows(0).Cells(1).Value = DataGridViewQue.Rows(1).Cells(1).Value
            DataGridViewQue.Rows(0).Cells(2).Value = DataGridViewQue.Rows(1).Cells(2).Value
            DataGridViewQue.Rows(0).Cells(3).Value = DataGridViewQue.Rows(1).Cells(3).Value
            DataGridViewQue.Rows(0).Cells(4).Value = DataGridViewQue.Rows(1).Cells(4).Value
            DataGridViewQue.Rows(0).Cells(5).Value = DataGridViewQue.Rows(1).Cells(5).Value
            DataGridViewQue.Rows(0).Cells(6).Value = DataGridViewQue.Rows(1).Cells(6).Value
            DataGridViewQue.Rows(0).Cells(7).Value = DataGridViewQue.Rows(1).Cells(7).Value
            '106.11.2 聯佶空單
            DataGridViewQue.Rows(0).Cells(8).Value = DataGridViewQue.Rows(1).Cells(8).Value
            DataGridViewQue.Rows(0).Cells(9).Value = DataGridViewQue.Rows(1).Cells(9).Value

            DataGridViewQue.Rows(1).Cells(0).Value = DataGridViewQue.Rows(2).Cells(0).Value
            DataGridViewQue.Rows(1).Cells(1).Value = DataGridViewQue.Rows(2).Cells(1).Value
            DataGridViewQue.Rows(1).Cells(2).Value = DataGridViewQue.Rows(2).Cells(2).Value
            DataGridViewQue.Rows(1).Cells(3).Value = DataGridViewQue.Rows(2).Cells(3).Value
            DataGridViewQue.Rows(1).Cells(4).Value = DataGridViewQue.Rows(2).Cells(4).Value
            DataGridViewQue.Rows(1).Cells(5).Value = DataGridViewQue.Rows(2).Cells(5).Value
            DataGridViewQue.Rows(1).Cells(6).Value = DataGridViewQue.Rows(2).Cells(6).Value
            DataGridViewQue.Rows(1).Cells(7).Value = DataGridViewQue.Rows(2).Cells(7).Value
            '106.11.2 聯佶空單
            DataGridViewQue.Rows(1).Cells(8).Value = DataGridViewQue.Rows(2).Cells(8).Value
            DataGridViewQue.Rows(1).Cells(9).Value = DataGridViewQue.Rows(2).Cells(9).Value

            DataGridViewQue.Rows(2).Cells(0).Value = DataGridViewQue.Rows(3).Cells(0).Value
            DataGridViewQue.Rows(2).Cells(1).Value = DataGridViewQue.Rows(3).Cells(1).Value
            DataGridViewQue.Rows(2).Cells(2).Value = DataGridViewQue.Rows(3).Cells(2).Value
            DataGridViewQue.Rows(2).Cells(3).Value = DataGridViewQue.Rows(3).Cells(3).Value
            DataGridViewQue.Rows(2).Cells(4).Value = DataGridViewQue.Rows(3).Cells(4).Value
            DataGridViewQue.Rows(2).Cells(5).Value = DataGridViewQue.Rows(3).Cells(5).Value
            DataGridViewQue.Rows(2).Cells(6).Value = DataGridViewQue.Rows(3).Cells(6).Value
            DataGridViewQue.Rows(2).Cells(7).Value = DataGridViewQue.Rows(3).Cells(7).Value
            '106.11.2 聯佶空單
            DataGridViewQue.Rows(2).Cells(8).Value = DataGridViewQue.Rows(3).Cells(8).Value
            DataGridViewQue.Rows(2).Cells(9).Value = DataGridViewQue.Rows(3).Cells(9).Value

            DataGridViewQue.Rows(3).Cells(0).Value = ""
            DataGridViewQue.Rows(3).Cells(1).Value = "0"
            DataGridViewQue.Rows(3).Cells(2).Value = ""
            DataGridViewQue.Rows(3).Cells(3).Value = ""
            DataGridViewQue.Rows(3).Cells(4).Value = ""
            DataGridViewQue.Rows(3).Cells(5).Value = ""
            DataGridViewQue.Rows(3).Cells(6).Value = ""
            DataGridViewQue.Rows(3).Cells(7).Value = ""
            '106.11.2 聯佶空單
            DataGridViewQue.Rows(3).Cells(8).Value = ""
            DataGridViewQue.Rows(3).Cells(9).Value = ""

            '107.4.29
            'TimerLink.Enabled = True
            If bPingSQL Then
                TimerLink.Enabled = True
            End If

            'show盤數資料
            For i = 0 To 9
                dgvPan.Rows(0).Cells(i + 1).value = ""
                dgvPan.Rows(1).Cells(i + 1).value = ""
            Next
            For i = 0 To queue0.plate.Length - 1
                dgvPan.Rows(0).Cells(i + 1).value = Format(Val(queue0.plate(i)), "0.00")
                dgvPan.Rows(1).Cells(i + 1).value = Format(Val(queue0.plate(i)), "0.00")
            Next
            '107.3.18 merge with below
            'LabelMessage.Visible = False
            'DVW(835) = 0
            'i = FrmMonit.WriteA2N(835, 1)
            ''100.6.18
            'Cbxquecode1.Enabled = True
            'numquetri1.Enabled = True
            'Tbxquecar1.Enabled = True
            'Tbxquefield1.Enabled = True
            ''100.6.29
            'btnque1.Enabled = True
            'btncancel1.Enabled = True
            ''104.9.24
            'Cbxquecode1.Enabled = True
            'numquetri1.Enabled = True
            'Tbxquecar1.Enabled = True
            ''101.2.25
            'btnque2.Enabled = True
            'btncancel2.Enabled = True

            ''107.5.30
            'Me.Refresh()

            '第一盤計量
            'dgvPan.Rows(0).Cells(1).Style.BackColor = Color.Green
            'dgvPan.Rows(1).Cells(1).Style.BackColor = Color.DeepPink
            '107.3.18
            'Call SavingLog("quecode reset D20,21 BON= " & queue0.BoneDonePlate & " , Mod= " & queue0.ModDonePlate & " , Rem= " & queue0.RemainDonePlate)
            MonLog.add = "quecode() : write D832..835, reset D20,21, queue0.BMR=" & queue0.BoneDonePlate & queue0.ModDonePlate & queue0.RemainDonePlate
            'D832 拌合時間設定
            'D833 拌合次數設定
            'D834 排料時間設定
            'D835 搶先中
            DVW(832) = CInt(tbxmixtime.Text)
            DVW(833) = queue0.plate.Length
            j = DVW(833)
            DVW(834) = CInt(Num_quetime_set.Text)
            '107.3.18
            'i = FrmMonit.WriteA2N(832, 3)
            LabelMessage.Visible = False
            DVW(835) = 0
            i = FrmMonit.WriteA2N(832, 4)

            '設定生產起始值  
            'remain = 0
            '110.11.12 @home
            'LabelTotPlate.Text = queue0.plate.Length
            'Call UpdateLabelTotPlate()

            DVW(20) = 0
            DVW(21) = 0
            'queue0.BoneDonePlate = 0
            'queue0.ModDonePlate = 0
            i = FrmMonit.WriteA2N(20, 2)

            'Timerprocess.Enabled = True
            '第一盤計量中
            '105.5.8
            i = queue1.Fmaterialnowater(24)
            j = queue0.Fmaterialnowater(24)

            mc.queclone(queue0, queue)
            queue = mc.Materialdistribute(queue, 1, cbxcorrect.Text, waterall, "A")
            mc.queclone(queue, queue0)
            '105.5.8
            i = queue1.Fmaterialnowater(24)
            j = queue0.Fmaterialnowater(24)

            queue0.BoneDoingPlate = 1
            queue0.ModDoingPlate = 1
            queue0.DisplayPlates()
            showotherdata(queue0.showdata)
            txt_carsum.Text = "0.00"


            '判斷年度檔案是否存在
            '100.6.11 sim
            If sim_produce Then
                If Not My.Computer.FileSystem.FileExists(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb") Then
                    '105.7.3
                    'My.Computer.FileSystem.CopyFile("\JS\CBC8\files\datasample.mdb", sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
                    My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
                    dbrec.changedb(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
                End If
            Else
                If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb") Then
                    '105.7.3
                    'My.Computer.FileSystem.CopyFile("\JS\CBC8\files\datasample.mdb", sYMdbName + sSavingYear + ".mdb")
                    My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
                    dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
                End If
            End If


            '105.8.10 move from SaveBone => too late if SaveBone
            Dim db As New DbAccess(sYMdbName + sSavingYear + ".mdb")
            '109.4.2 Dim dt As DataTable

            ' 109.2.17
            '   B配統計短少 => A配 未明原因 uni_car => UN_1 (orginal => when data table no record)
            'If sim_produce Then '106.3.12  此時
            '    dt = db.GetDataTable("select  top 1 UniID from recdata_SIM order by UniID desc ")
            'Else
            '    dt = db.GetDataTable("select  top 1 UniID from recdata order by UniID desc ")
            'End If
            'If dt.Rows.Count < 1 Then
            '    iUniSer = 1
            'Else
            '    iUniSer = CInt(dt.Rows(0).Item("UniID").ToString) + 1
            'End If
            iUniSer = Now.Day * 1000000 + Now.Hour * 10000 + Now.Minute * 100 + Now.Second

            queue0.UniCar = "UN_" & (iUniSer).ToString
            queue0.CarSer = iUniSer
            LabelUniCar.Text = queue0.UniCar

            MonLog.add = "displaymaterial() : quecode " & queue0.UniCar
            displaymaterial()

            db.dispose()

            '102.5.11
            bUPDATEM3_1 = True
        End If

        '100.7.12
        Label12.Visible = True

        '103.3.15
        If bEngVer Then
            Label12.Visible = False
        End If

        '102.7.19
        If queue0.isBsub Then
            lblcar.ForeColor = Color.Blue
        Else
            lblcar.ForeColor = Color.Black
        End If
        '103.2.25
        If bEngVer Then
            '103.3.11
            'If queue0.Fmaterialorg(24) = 99 Then
            '    bImperial = True
            'Else
            '    bImperial = False
            'End If
            'If bImperial Then
            '    fImperial = KG_TO_LB
            'Else
            '    fImperial = 1.0
            'End If
            '103.3.11
            'Call UpdateImperial()
        End If

        '105.5.8
        i = queue1.Fmaterialnowater(24)
        j = queue0.Fmaterialnowater(24)

        '107.6.14
        '   預設含水率 move from btnbegin.click since also called when 排隊自動啟動 use CheckDefaultWater()
        '107.6.20 bug
        'CheckDefaultWater(queue0.showdata(1), Tbxquefield1.Text)
        CheckDefaultWater(queue0.showdata(1), tbxworkfield.Text)

    End Sub

    Private Sub showotherdata(ByVal showdata() As String)
        '100.5.15
        'cbxcorrect.Text = showdata(0)
        tbxspare.Text = showdata(1)
        tbxmixtime.Text = showdata(3)
        FrmMonit.tbxmixtime.Text = tbxmixtime.Text
        tbxfounder.Text = showdata(4)
        '102.6.7
        'tbxstrengt.Text = showdata(5)
        '102.6.25 強度-粒徑-坍度 => 強度-坍度-粒徑
        'tbxstrengt.Text = showdata(5) & "-" & showdata(6) & "-" & showdata(4)

        '109.9.11
        'tbxstrengt.Text = showdata(5) & "-" & showdata(4) & "-" & showdata(6)
        '110.9.14 強度-坍度-粒徑
        'showdata : (0)-補正區域,(1)-B配比, (2)-總重, (3)-拌合時間, (4)-坍度, (5)-強度, (6)-粒徑, (7~9)-備註一~三
        tbxstrengt.Text = showdata(5) & "-" & showdata(4) & "-" & showdata(6)

        tbxparticle.Text = showdata(6)
        'tbxmemo1.Text = showdata(7)
        'tbxmemo2.Text = showdata(8)
        'tbxmemo3.Text = showdata(9)
    End Sub

    Private Sub btnstop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnstop.Click
        '102.4.17
        If bEngVer Then
            If MsgBox("EMERGENCY STOP!!!", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                '110.11.9
                'If queue0.inprocess = True Then
                If True Then
                    clearproduce()
                    emgstop = True
                    '100.6.29
                    Label12.Visible = False
                Else : emgstop = False
                End If
            End If
        Else
            If MsgBox("確定要緊急停止!!!", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                '110.11.9
                'If queue0.inprocess = True Then
                If True Then
                    clearproduce()
                    emgstop = True
                    '100.6.29
                    Label12.Visible = False
                    '104.9.24
                    Cbxquecode1.Enabled = True
                    numquetri1.Enabled = True
                    Tbxquecar1.Enabled = True
                    Tbxquefield1.Enabled = True
                    btnque1.Enabled = True
                    btncancel1.Enabled = True
                Else : emgstop = False
                End If
            End If
        End If
    End Sub

    Private Sub Cbxquecode2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbxquecode2.Leave
        If Cbxquecode1.Text = "" Then
            Cbxquecode1.Text = Cbxquecode2.Text
            Cbxquecode2.Text = ""
            queue1check()
        End If
    End Sub

    Private Sub Numquetri2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numquetri2.Leave
        If Convert.ToSingle(numquetri1.Text) = 0 Then
            numquetri1.Text = numquetri2.Text
            numquetri2.Text = "0.0"
            queue1check()
        End If
    End Sub

    Private Sub btncancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel1.Click
        If LabelMessage.Visible = False Then
            numquetri1.Text = "0.0"
            Cbxquecode1.Text = ""
            Tbxquecar1.Text = ""
            Tbxquefield1.Text = ""
            tbxCusNo1.Text = ""
            lblCusName1.Text = ""
            '99.09.28
            LabelCust1.Text = ""
            LabelQkey1.Text = ""

            '100.2.3
            'If Convert.ToSingle(numquetri2.Text) <> 0 Then numquetri1.Text = numquetri2.Text
            'If Cbxquecode2.Text <> "" Then Cbxquecode1.Text = Cbxquecode2.Text
            'If Tbxquecar2.Text <> "" Then Tbxquecar1.Text = Tbxquecar2.Text
            'If Tbxquefield2.Text <> "" Then Tbxquefield1.Text = Tbxquefield2.Text
            If Convert.ToSingle(numquetri2.Text) <> 0 Then numquetri1.Text = numquetri2.Text
            Cbxquecode1.Text = Cbxquecode2.Text
            Tbxquecar1.Text = Tbxquecar2.Text
            Tbxquefield1.Text = Tbxquefield2.Text
            tbxCusNo1.Text = tbxCusNo2.Text
            lblCusName1.Text = lblCusName2.Text
            LabelCust1.Text = LabelCust2.Text
            LabelQkey1.Text = LabelQkey2.Text
            '106.10.7
            btnque1.ForeColor = btnque2.ForeColor
            'If btnque1.ForeColor = Color.Black Then
            '    LabelPb12_1.Visible = False
            'End If
            '106.11.2
            LabelPb12_1.Visible = LabelPb12_2.Visible
            btnque2.ForeColor = Color.Black
            LabelPb12_2.Visible = False

            '106.10.7 106.11.2
            'btnque1.ForeColor = Color.Black
            'LabelPb12_1.Visible = False


            QueShift()
            queue1check()
        Else
            '103.1.22
            'MsgBox("骨材搶先中 不可刪除排隊一!", MsgBoxStyle.Critical, "骨材搶先中")
            If bEngVer Then
                MsgBox("Delete Inhibited", MsgBoxStyle.Critical, "Warning")
            Else
                MsgBox("骨材搶先中 不可刪除排隊一!", MsgBoxStyle.Critical, "骨材搶先中")
            End If
        End If
    End Sub

    Private Sub btncancel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel2.Click
        '106.11.2
        btnque2.ForeColor = Color.Black
        LabelPb12_2.Visible = False
        QueShift()
    End Sub

    Private Sub btnque1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnque1.Click
        If tbxworkcode.Text <> "" And TbxworkTri.Text <> "" Then
            '100.4.14If Cbxquecode1.Text = "" And numquetri1.Text = "0.0" Then
            If Cbxquecode1.Text = "" And CInt(numquetri1.Text) = 0 Then
                Cbxquecode1.Text = tbxworkcode.Text
                numquetri1.Text = TbxworkTri.Text
                Tbxquecar1.Text = Tbxworkcar.Text
                Tbxquefield1.Text = tbxworkfield.Text
                queue1check()
            End If
        End If
    End Sub

    Private Sub btnque2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnque2.Click
        If Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0 Then
            If Cbxquecode2.Text = "" And Convert.ToSingle(numquetri2.Text) = 0 Then
                Cbxquecode2.Text = Cbxquecode1.Text
                numquetri2.Text = numquetri1.Text
                Tbxquecar2.Text = Tbxquecar1.Text
                Tbxquefield2.Text = Tbxquefield1.Text
            End If
        End If
    End Sub

    Private Sub Btnfomula_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnfomula.Click
        '105.7.25 環泥烏日
        'If bCheckBoxSP6 Then
        '105.10.7 BY 配比密碼紀錄 bCheckBoxSP9
        Dim db As New DbAccess(sSMdbName & "\setting.mdb")
        If bCheckBoxSP9 Then
            Dim s, s2
            sPBUser = InputBox("請輸入使用者名稱", "配方")
            If sPBUser <> "" Then
                Dim dt As DataTable = db.GetDataTable("select * from passwd where ID = '" + sPBUser + "'")
                '105.10.21
                If dt.Rows.Count <> 0 Then
                    'If dt.Rows.Count <> 0 Or sPBUser = "js" Then
                    s = dt.Rows(0).Item("ID").ToString
                    s2 = InputBox("請輸入使用者密碼", "配方")
                    If dt.Rows(0).Item("Passwd").ToString = s2 Then
                        'MsgBox("使用者存在!!!", MsgBoxStyle.Information)
                        s = dt.Rows(0).Item("setting").ToString
                        If s > 1 Then
                            showform(Fomula)
                        End If
                    Else
                        If bEngVer Then
                            MsgBox("Wrong Password!!!!", MsgBoxStyle.Information)
                        Else
                            MsgBox("密碼不正確!!!!", MsgBoxStyle.Information)
                        End If
                    End If
                Else
                    If bEngVer Then
                        MsgBox("User not exist!!!!", MsgBoxStyle.Information)
                    Else
                        MsgBox("使用者不存在!!!!", MsgBoxStyle.Information)
                    End If
                End If
            Else
                If bEngVer Then
                    MsgBox("Please keyin User Name", MsgBoxStyle.Critical)
                Else
                    MsgBox("請輸入使用者名稱", MsgBoxStyle.Critical)
                End If
            End If
        Else
            PasswordTextBox.Text = ""
            Panel7.Visible = True
            PasswordTextBox.Focus()
        End If
        'showform(Fomula)
    End Sub

    Private Sub showform(ByVal frm As Form)
        'frmpasswd.ShowDialog()
        'frm.TopMost = True
        'RAY
        'module1.authority = 3
        If module1.authority >= 0 Then
            module1.callfrm = 1
            frm.ShowDialog()
        End If
    End Sub

    Private Sub btnparam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnparam.Click
        showform(Frmshowsetting)
    End Sub


    Private Sub btn_simproduce_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simproduce.Click
        simstatus(True)
        'ErrorProvider1.SetError(btn_simproduce, "error提供者")
    End Sub

    Private Sub btn_simstop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simstop.Click
        simstatus(False)
    End Sub

    Private Sub btn_simstart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simstart.Click

        '判斷年度檔案是否存在
        '100.5.15
        If Not My.Computer.FileSystem.FileExists(sYMdbName + dtp_simtime.Value.Year.ToString & "M" & dtp_simtime.Value.Month.ToString + ".mdb") Then
            My.Computer.FileSystem.CopyFile("\JS\Origfile\new_datasample_f.mdb", sYMdbName + dtp_simtime.Value.Year.ToString & dtp_simtime.Value.Month.ToString + ".mdb")
        End If

        Dim db As New DbAccess(sYMdbName + dtp_simtime.Value.Year.ToString & "M" & dtp_simtime.Value.Month.ToString + ".mdb")
        Dim dt As DataTable = db.GetDataTable("select max(Carser) from recdata_SIM ")

        '100.7.12 Dim i As Integer
        Dim i
        '100.7.12 If dt.Rows.Count <> 0 Then
        On Error Resume Next
        If dt.Rows.Count >= 1 Then
            i = dt.Rows(0).Item(0)
            '101.3.3
            'If i < 31999 Then
            '    iCarSer_SIM = dtp_simtime.Value.Month * 10000
            'Else
            '    iCarSer_SIM = i
            'End If
            iCarSer_SIM = i
        Else
            iCarSer_SIM = 1
        End If

        bCommFlag = False
        '100.5.17
        LabelMessage2.Text = "模擬報表功能啟動,通信停止!"
        LabelMessage2.Visible = True
        '107.3.20
        TLog.add = " 模擬報表啟動"
        sim_TimerCt = 0
        '100.6.8 dtSimP(11)
        Dim generator As New Random
        Dim randomValue As Integer
        Dim s As String
        '111.7.22
        '   模擬報表 首盤 CheckBoxFirstPan
        'dtSimP(queue0.plate.Length - 1) = dtp_simtime.Value
        If bCheckBoxFirstPan Then
            dtSimP(0) = dtp_simtime.Value
        Else
            dtSimP(queue0.plate.Length - 1) = dtp_simtime.Value
        End If
        If queue0.plate.Length > 1 Then
            For i = 0 To queue0.plate.Length - 2
                randomValue = generator.Next(1, 10)
                s = CInt(tbxmixtime.Text) + Num_quetime_set.Value + randomValue
                '111.7.22
                '   模擬報表 首盤 CheckBoxFirstPan
                'dtSimP(queue0.plate.Length - i - 2) = dtSimP(queue0.plate.Length - i - 1).AddSeconds(CInt(s) * -1)
                If bCheckBoxFirstPan Then
                    dtSimP(i + 1) = dtSimP(i).AddSeconds(CInt(s))
                Else
                    dtSimP(queue0.plate.Length - i - 2) = dtSimP(queue0.plate.Length - i - 1).AddSeconds(CInt(s) * -1)
                End If
            Next

        End If

        '101.3.3
        DV(20) = 0
        DV(21) = 0
        DV(22) = 0
        '102.5.29 107.3.20 !!!
        'ReadD61Flag = 2
        'ReadD70Flag = 2

        '101.3.30 Timerprocess_sim.Interval = 500
        DVold(20) = 0
        DVold(21) = 0
        DVold(22) = 0
        '106.10.21
        'Timerprocess_sim.Interval = 1000

        '107.3.20 revised Timerprocess_sim process, default 750ms
        'Timerprocess_sim.Interval = sPlcDelay * 2

        'While DV(22) < queue0.plate.Length
        Timerprocess_sim.Enabled = True
        While DVold(22) < queue0.plate.Length
            randomValue = generator.Next(1, 10)
            '100.6.8
            'txt_mixtime_real.Text = CInt(tbxmixtime.Text) + Num_quetime_set.Value + randomValue
            'dtSim = dtp_simtime.Value.AddSeconds(CInt(txt_mixtime_real.Text))
            '101.3.30 dtSim = dtSimP(DV(22))
            dtSim = dtSimP(DVold(22))
            LabelSimDt.Text = dtSim.ToString
            dtp_simtime.Value = dtSim
            'While Timerprocess_sim.Enabled = True
            Application.DoEvents()
            'End While
            'DV(20) += 1
            'DV(21) += 1
            'DV(22) += 1
            '99.12.04 randomValue = generator.Next(20, 100)
            '100.5.15
        End While

    End Sub

    Private Sub btnremain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnremain.Click
        '99.08.02
        'If queue0.inprocess = False Then Exit Sub
        DVW(22) = DV(22) + 1
        Dim i
        i = FrmMonit.WriteA2N(22, 1)

        'txtremain.Text += 1
        'dgvPan.Rows(0).Cells(1).Style.ForeColor = Color.DeepPink
        'dgvPan.Rows(0).Cells(0).Style.ForeColor = Color.DeepPink
    End Sub

    Private Sub chk_1_all_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_1_all.CheckedChanged
        Dim chkview As Boolean
        If chk_1_all.Checked Then
            chk_1_all.Text = "單方"
            DataGridView1.Columns(2).HeaderText = "1.00"
            DataGridView1.Columns(5).HeaderText = "1.00"
            chkview = True
        Else
            chk_1_all.Text = "總重"
            If queue0.inprocess Then
                DataGridView1.Columns(3).HeaderText = Format(Val(queue0.plate(Val(txtbone.Text))), "N02")
                DataGridView1.Columns(6).HeaderText = Format(Val(queue0.plate(Val(Txtmod.Text))), "N02")
            Else
                DataGridView1.Columns(3).HeaderText = "1.00"
                DataGridView1.Columns(6).HeaderText = "1.00"
            End If
            chkview = False
        End If
        DataGridView1.Columns(2).Visible = chkview
        DataGridView1.Columns(3).Visible = Not chkview
        DataGridView1.Columns(5).Visible = chkview
        DataGridView1.Columns(6).Visible = Not chkview

    End Sub

    Private Sub chk_A_B_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_A_B.CheckedChanged
        'tbxspare.Visible = chk_A_B.Checked
        'lblspare.Visible = chk_A_B.Checked
        If chk_A_B.Checked Then
            If sim_addshow_weight Then
                LabelRealTot1M3.Text = Convert.ToSingle(str_cur_weight(0)) + Convert.ToSingle(str_cur_weight(2))
            Else
                'LabelRealTot1M3.Text = Format(str_cur_weight(0), "0")
            End If
        Else
            'LabelRealTot1M3.Text = Format(str_cur_weight(1), "0")
        End If

    End Sub


    Private Sub lbl_monthsum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_monthsum.Click
        'Dim i As Integer
        'Dim j As Integer
        'Dim k As Integer
        'DVx(30) = DVx(30) + 1
        'DVx(31) = DVx(31) + 31

        'j = 14
        'For i = 0 To 9
        '    k = j Mod 2
        '    If MatSiloNo(i) = 1 Then
        '        FrmMonit.SetDGV2Color(FrmMonit.DGV1, MatSiloOrder(i), k)
        '    ElseIf MatSiloNo(i) = 2 Then
        '        FrmMonit.SetDGV2Color(FrmMonit.DGV2, MatSiloOrder(i), k)
        '    ElseIf MatSiloNo(i) = 3 Then
        '        FrmMonit.SetDGV2Color(FrmMonit.DGV3, MatSiloOrder(i), k)
        '    ElseIf MatSiloNo(i) = 4 Then
        '        FrmMonit.SetDGV2Color(FrmMonit.DGV4, MatSiloOrder(i), k)
        '    End If
        '    j = j \ 2
        'Next
        'Call FrmMonit.Check_DGV_Color()
        Call FrmMonit.CheckDgvSiloMatColor()

    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '102.4.17
        '105.12.17
        'If bEngVer Then
        '106.8.21
        Dim sendFlag As Boolean
        sendFlag = False

        If False Then
            If MsgBox("結束程式? Quit program? (請勿任意關閉程式 Please confirm.)", MsgBoxStyle.YesNo, "QUIT PROGRAM") = MsgBoxResult.Yes Then End
        Else
            '103.7.26
            'If MsgBox("結束程式?(請勿任意關閉程式)", MsgBoxStyle.YesNo, "關閉程式") = MsgBoxResult.Yes Then End
            If MsgBox("結束程式? Quit program? (請勿任意關閉程式 Please confirm.)", MsgBoxStyle.YesNo, "關閉程式(QUIT)") = MsgBoxResult.Yes Then
                Dim i
                For i = 0 To 29
                    DVW(i + 741) = 0
                    DVW(i + 771) = 0
                Next
                i = FrmMonit.WriteA2N(741, 30)
                i = FrmMonit.WriteA2N(771, 30)
                '107.6.11
                DVW(829) = 0
                i = FrmMonit.WriteA2N(829, 1)
                LabelMessage2.Visible = True
                '105.6.9
                LabelMessage2.Width = 500
                LabelMessage2.Height = 500

                LabelMessage2.Text = "請稍候(Wait)..."
                LabelMessage2.Refresh()
                'For i = 0 To 3000
                '    LabelMessage2.Text = "請稍候...." & i
                'Next
                '104.8.10
                Dim j
                Dim fileNum
                Dim s As String
                fileNum = FreeFile()
                FileOpen(fileNum, sSMdbName & "\water.per", OpenMode.Output)
                For j = 0 To 14
                    '104.8.10
                    's = dgvWater.Rows(0).Cells(j).Value
                    'PrintLine(fileNum, Trim(s))
                    s = fWatOrig(j).ToString
                    PrintLine(fileNum, Format(Val(s), "0.0"))
                Next
                FileClose(fileNum)

                '104.8.18
                fileNum = FreeFile()
                FileOpen(fileNum, sSMdbName & "\water2.per", OpenMode.Output)
                For i = 0 To 14
                    For j = 0 To 7
                        s = fWat(j, i).ToString
                        PrintLine(fileNum, Format(Val(s), "0.0"))
                    Next
                Next
                FileClose(fileNum)
                '107.3.8
                TLog.saveLog()
                MonLog.saveLog()
                MainLog.saveLog()
                DbLog.saveLog()
                '109.4.2
                UploadPbLog.saveLog()
                '111.10.21
                RecSqlLog.saveLog()

                '
                '105.11.29 bCheckBoxSP7 參數存至遠端
                If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bCheckBoxSP7 And bPingRPC Then Call SaveFilesToRemote()

                '105.6.4 105.7.3 資料自動補傳時訊 105.11.29
                'If bCheckBoxSP4 Then
                '105.12.23
                'If bCheckBoxSP4 Then

                '106.7.25 try IP before exit
                bNetPing = False
                '107.8.31 
                'If My.Computer.Network.IsAvailable Then
                If My.Computer.Network.IsAvailable And sDataBaseIP <> "127.0.0.1" Then
                    If My.Computer.Network.Ping(sDataBaseIP, 1000) Then
                        bNetPing = True
                        '106.9.9
                        bDataBaseLink = True
                    End If
                End If

                '107.8.14 move from below
                '107.7.2 107.7.3
                'If bRadioButton_OP And bDataBaseLink Then
                If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bPingRPC Then
                    Dim sSavingYear_ As String
                    sSavingYear_ = sSavingYear
                    Dim datafile As String
                    Dim datafile_r As String
                    datafile = sYMdbName + sSavingYear_ + ".mdb"
                    datafile_r = sYMdbName_A + sSavingYear_ + ".mdb"
                    LabelMessage2.Text = "請稍候(Wait)..A" & (j)
                    Me.Refresh()
                    My.Computer.FileSystem.CopyFile(datafile, datafile_r, True)
                End If
                '107.8.31
                'If sDataBaseIP = "127.0.0.1" Then
                '    End
                'End If
                ''If bCheckBoxSP4 And bNetPing Then
                '111.9.28 disable old 
                'If bCheckBoxSP4 And bNetPing And sDataBaseIP <> "127.0.0.1" Then
                If False And bCheckBoxSP4 And bNetPing And sDataBaseIP <> "127.0.0.1" Then
                    Dim yy, mm
                    Dim dbfile
                    Dim db_sum As DataTable
                    Dim CarSer, UniCar, plate, Resend
                    '106.1.3
                    Dim ssSavingYear
                    '105.9.24 bCheckBoxSP7 bPingRPC 105.11.5 105.11.29
                    ' If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bCheckBoxSP7 And bPingRPC Then Call SaveFilesToRemote()
                    mm = Now.Month
                    yy = Now.Year
                    dbfile = sYMdbName & yy & "M" & mm & ".mdb"

                    '106.7.22 亞東 自動補傳 FOR UpdateRecSend()
                    Dim sTemp As String
                    sTemp = s_receia
                    If bCheckBoxYadon Then
                        s_receia = "receib"
                    End If

                    '105.11.5
                    Dim file
                    file = My.Computer.FileSystem.FileExists(dbfile)
                    If file = False Then
                        End
                    End If
                    dbrec = New DbAccess(dbfile)
                    '105.12.20
                    strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
                    'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata "
                    db_sum = dbrec.GetDataTable(strsql)

                    dbrec.ExecuteCmd(strsql)

                    For j = 0 To db_sum.Rows.Count - 1
                        UniCar = db_sum.Rows(j).Item("UniCar")
                        CarSer = db_sum.Rows(j).Item("Carser")
                        plate = db_sum.Rows(j).Item("plate")
                        Resend = db_sum.Rows(j).Item("nowaterMaterial2")
                        '106.9.4If bNetPing And Resend = 0 And bRadioButton_OP Then
                        If bNetPing And Resend = 0 And bRadioButton_OP And bDataBaseLink Then
                            LabelMessage2.Text = "請稍候(Wait)...共" & (j + 1) & "/" & db_sum.Rows.Count & "盤次, "
                            LabelMessage2.Refresh()
                            ssSavingYear = yy & "M" & mm
                            '106.7.22 亞東 自動補傳 : Transed : 工控寫入(A：自動補傳、H：手動補傳)
                            sTransed = "A"
                            Re_SendSQL_Flag = True
                            '107.4.15
                            TLog.add = "Exit.mc.Re_SendSQL(" & CarSer & "," & UniCar & "," & plate & ") " & sYMdbName & yy & "M" & mm & ".mdb"

                            '112.4.28
                            'mc.Re_SendSQL(CarSer, UniCar, plate, sYMdbName & yy & "M" & mm & ".mdb", ssSavingYear)
                            'Dim strsqlall As String
                            'strsqlall = "Select DISTINCT * from recdata WHERE unicar='" & UniCar & "' and plate =" & plate.ToString & " order by RecDT DESC "
                            'Dim db_sumSQL As New DataTable
                            'db_sumSQL = dbrec.GetDataTable(strsqlall)
                            'Dim receiaPan As New classReceiaPan()
                            'Dim jj As Integer
                            'For jj = 0 To db_sumSQL.Rows.Count - 1
                            '    receiaPan = fillMdbRowToPan(db_sumSQL.Rows(jj), "A")
                            '    If receiaPan.Header.sumno = receiaPan.Header.no Then
                            '        strsqlall = "Select SUM(sand1),SUM(sand2),SUM(sand3),SUM(sand4),SUM(Stone1),SUM(Stone2),SUM(Stone3),SUM(Stone4),SUM(Water1),SUM(Water2),SUM(Water3)"
                            '        strsqlall &= ",SUM(Concrete1),SUM(Concrete2),SUM(Concrete3),SUM(Concrete4),SUM(Concrete5),SUM(Concrete6)"
                            '        strsqlall &= ",SUM(Drog1),SUM(Drog2),SUM(Drog3),SUM(Drog4),SUM(Drog5),SUM(cube) from recdata WHERE [UniCar] ='" & UniCar & "'"
                            '        Dim db_Dt As New DataTable
                            '        db_Dt = dbrec.GetDataTable(strsqlall)

                            '        '112.4.19 bug:only sum no remain
                            '        strsqlall = "Select SUM(remainsand1),SUM(remainsand2),SUM(remainsand3),SUM(remainsand4),SUM(remainStone1),SUM(remainStone2),SUM(remainStone3),SUM(remainStone4),SUM(remainWater1),SUM(remainWater2),SUM(remainWater3)"
                            '        strsqlall &= ",SUM(remainConcrete1),SUM(remainConcrete2),SUM(remainConcrete3),SUM(remainConcrete4),SUM(remainConcrete5),SUM(remainConcrete6)"
                            '        strsqlall &= ",SUM(remainDrog1),SUM(remainDrog2),SUM(remainDrog3),SUM(remainDrog4),SUM(remainDrog5),SUM(cube) from recdata WHERE [UniCar] ='" & UniCar & "'"
                            '        Dim db_Dt_rem As New DataTable
                            '        db_Dt_rem = dbrec.GetDataTable(strsqlall)

                            '        If db_Dt.Rows.Count > 0 Then
                            '            '112.4.19
                            '            'receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0))
                            '            receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0), db_Dt_rem.Rows(0))
                            '            receiaPan.DT.qty = db_Dt.Rows(0).Item(22)
                            '        End If
                            '        Dim send
                            '        send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", True), "A")
                            '        If bPingSQLBigData Then
                            '            send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", True), "B")
                            '        End If
                            '    Else
                            '        '112.4.12
                            '        'Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "D", False), "B")
                            '        ''112.1.28
                            '        'send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "D", False), "A")
                            '        Dim send
                            '        send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "D", False), "A")
                            '        If bPingSQLBigData Then
                            '            send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receia", "D", False), "B")
                            '        End If
                            '    End If
                            'Next

                            s = SQLyymodd
                            s = SQLhhmmss
                            'If bNetPing Then mc.UpdateRecSend(CarSer, UniCar, plate, dbfile, "recdata")
                            '105.12.21
                            'mc.UpdateRecSend(CarSer, UniCar, plate, dbfile, "recdata")
                            mc.UpdateRecSend(UniCar, plate, dbfile)
                            sendFlag = True
                        Else '106.9.4 if not link
                            Exit For
                        End If
                    Next

                    If Now.Month = 1 Then
                        mm = 12
                    Else
                        mm = Now.Month - 1
                    End If
                    If mm = 12 Then
                        yy = Now.Year - 1
                    Else
                        yy = Now.Year
                    End If
                    dbfile = sYMdbName & yy & "M" & mm & ".mdb"

                    '105.11.5
                    file = My.Computer.FileSystem.FileExists(dbfile)
                    '107.4.15
                    'If file = False Then
                    '    End
                    'End If
                    If file = True Then
                        dbrec = New DbAccess(dbfile)
                        'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
                        '105.12.20
                        'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata "
                        strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0 "
                        db_sum = dbrec.GetDataTable(strsql)
                        dbrec.ExecuteCmd(strsql)
                        For j = 0 To db_sum.Rows.Count - 1
                            UniCar = db_sum.Rows(j).Item("UniCar")
                            CarSer = db_sum.Rows(j).Item("Carser")
                            plate = db_sum.Rows(j).Item("plate")
                            Resend = db_sum.Rows(j).Item("nowaterMaterial2")
                            '106.9.4If bNetPing And Resend = 0 And bRadioButton_OP Then
                            If bNetPing And Resend = 0 And bRadioButton_OP And bDataBaseLink Then
                                '105.12.20
                                LabelMessage2.Text = "請稍候(Wait).....共_" & (j + 1) & "/" & db_sum.Rows.Count & "盤次, "
                                LabelMessage2.Refresh()
                                '106.1.3.
                                ssSavingYear = yy & "M" & mm
                                '106.7.22 亞東 自動補傳 : Transed : 工控寫入(A：自動補傳、H：手動補傳)
                                sTransed = "A"
                                Re_SendSQL_Flag = True
                                '107.4.15
                                TLog.add = "Exit.mc.Re_SendSQL(" & CarSer & "," & UniCar & "," & plate & ") " & sYMdbName & yy & "M" & mm & ".mdb"
                                mc.Re_SendSQL(CarSer, UniCar, plate, sYMdbName & yy & "M" & mm & ".mdb", ssSavingYear)
                                s = SQLyymodd
                                s = SQLhhmmss
                                mc.UpdateRecSend(UniCar, plate, dbfile)
                                sendFlag = True
                            Else '106.9.4 if not link
                                Exit For
                            End If
                        Next
                    End If
                End If

                If sendFlag = True Then
                    For j = 0 To 5
                        LabelMessage2.Text = "請稍候(Wait)..." & (j)
                        Me.Refresh()
                    Next
                End If

                '105.7.3 kill B
                If Not bCheckBoxSP3 Then
                    'My.Computer.FileSystem.DeleteFile(sYMdbName_B + sSavingYear + ".mdb")
                    'My.Computer.FileSystem.DeleteFile(sYMdbName_B_Local + sSavingYear + ".mdb")
                End If

                '111.9.23 new resendSQL
                'If bCheckBoxSP4 And bNetPing And bCheckBoxYadon Then
                '111.10.18
                'If bDataBaseLink And bCheckBoxSP4 And bCheckBoxYadon Then
                If bDataBaseLink And bCheckBoxSP4 And bCheckBoxYadon Then
                    datefind = CDate(sSavingDate).Year & Format(CDate(sSavingDate).Month, "00") & Format(CDate(sSavingDate).Day, "00")
                    'datefind = "20220802"

                    Dim da As Date
                    Dim dd As Integer
                    LabelMessage2.Refresh()
                    For dd = -10 To 0
                        da = DateAdd(DateInterval.Day, dd, CDate(sSavingDate))
                        datefind = da.Year & Format(da.Month, "00") & Format(da.Day, "00")
                        LabelMessage2.Text = "稍候." & datefind

                        i = resendSqlDay(datefind, "D")
                        LabelMessage2.Text = " 補傳...... " & datefind & " " & (i)
                        LabelMessage2.Refresh()
                        LabelMessage2.Text = "稍候__" & datefind
                        i = resendSqlDay(datefind, "R")
                        '112.9.23
                        If i < 0 Then
                            End
                        End If

                        LabelMessage2.Text = " 補傳___... " & datefind & " " & (i)
                        LabelMessage2.Refresh()
                    Next

                    For j = 0 To 5
                        LabelMessage2.Text = "請稍候(Wait)..." & (j)
                        Me.Refresh()
                    Next
                End If


                End

            End If

        End If

    End Sub
    Public Sub SaveFilesToRemote()
        '105.9.3 copy files in conf_Qxx to remote PC
        Dim file
        Dim sorceFile, destFile As String

        'sorceFile = "\JS\CBC8\Conf_" & sProject & "\Alarm.des"
        'destFile = "\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject & "\Alarm.des"
        file = My.Computer.FileSystem.DirectoryExists("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject)
        If file = False Then
            Try
                My.Computer.FileSystem.CreateDirectory("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject)
            Catch ex As Exception
                MsgBox("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject & "Directory can not create!")
                Exit Sub
            End Try
        End If

        ' Dim FileNames As String in My.Computer.FileSystem.GetFiles("\JS\CBC8\Conf_" & sProject)
        'Dim foundFile
        'For Each foundFile As String In My.Computer.FileSystem.GetFiles _
        '(My.Computer.FileSystem.SpecialDirectories.MyDocuments)
        '107.7.4 no copy Admin2.des
        '107.7.9 no Admin2 action => back to use Admin only and NO COPY to REMOTE!
        Dim s
        For Each foundFile As String In My.Computer.FileSystem.GetFiles("\JS\CBC8\Conf_" & sProject, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            '106.9.4
            'sorceFile = foundFile
            'destFile = "\\" & sTextBoxIP_R & "\" & Microsoft.VisualBasic.Right(foundFile, Len(foundFile) - 2)
            'My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)
            '107.7.4
            s = Microsoft.VisualBasic.Right(foundFile, 9)
            Try
                sorceFile = foundFile
                destFile = "\\" & sTextBoxIP_R & "\" & Microsoft.VisualBasic.Right(foundFile, Len(foundFile) - 2)
                '107.7.4
                'My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)
                '107.7.9
                'If s <> "Admin2.des" Then
                If s <> "Admin.des" Then
                    My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)
                Else
                    s = "test"
                End If
            Catch ex As Exception
                MsgBox("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject & "Directory can not create!")
                Exit Sub
            End Try
        Next

        'sorceFile = "\JS\CBC8\Conf_" & sProject & "\Admin.des"
        'destFile = "\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject & "\Admin.des"
        'file = My.Computer.FileSystem.DirectoryExists("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject)
        'My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)

        'sorceFile = "\JS\CBC8\Conf_" & sProject & "\AE15.des"
        'destFile = "\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject & "\AE15.des"
        'file = My.Computer.FileSystem.DirectoryExists("\\" & sTextBoxIP_R & "\JS\CBC8\Conf_" & sProject)
        'My.Computer.FileSystem.CopyFile(sorceFile, destFile, True)

    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '107.6.10 排隊一 空白 DV829=1 for 拌合機節能
        '107.6.24
        Static Cbxquecode1_O1 As String = "***"
        Static Cbxquecode1_O2 As String = "*****"

        ''107.6.24
        Cbxquecode1_O2 = Cbxquecode1_O1
        Cbxquecode1_O1 = Cbxquecode1.Text


        '107.4.28 
        '110.9.11 110.6.24 move to TimerAlarm
        'Call AlarmCheck()
        '107.4.29
        Call Dgv2Change()

        'LabelTime.Text = Year(Now) - 1911 & "/" & Month(Now) & "/" & Microsoft.VisualBasic.DateAndTime.Day(Now) & " " & Format(Now, "HH:mm:ss")
        LabelTime.Text = Year(Now) & "/" & Month(Now) & "/" & Microsoft.VisualBasic.DateAndTime.Day(Now) & " " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
        '101.12.17 Label7.Text = "ComCt: " & FrmMonit.COM_PLC.ErrorCount & " - " & FrmMonit.COM_PLC.PollCount
        Dim dt As Double
        '102.2.17
        Dim dt2 As Double

        '105.9.14 TEST plc comm 105.9.16
        'If LabelMsg.Visible = True Then
        '    DVW(1) = DV(1) + 1
        '    DVW(2) = DV(2) + 2
        '    Dim i
        '    i = FrmMonit.WriteA2N(1, 2)
        'End If


        '110.9.15
        'If queue1.ismatch = True Then
        '110.9.24
        'If (queue1.ismatch = True Or (Cbxquecode1.Text <> "" And numquetri1.Text <> "")) Then
        If (queue1.ismatch = True) Then

        Else
            '排隊一有料開始計時
            dtBatDonQ1 = Now
        End If
        dt2 = (Now.Ticks - dtBatDonQ1.Ticks) * 0.0000001
        If dt2 > 60000 Then dt2 = 0
        '排隊一有料開始計時
        dtBatDonIntQ1 = CInt(dt2)

        '110.9.24
        If queue0.inprocess Then
            '生產完畢開始計時
            dtBatDon = Now
        End If
        dt = (Now.Ticks - dtBatDon.Ticks) * 0.0000001
        If dt > 60000 Then dt = 0
        '生產完畢開始計時
        dtBatDonInt = CInt(dt)
        '105.11.3
        'Label7.Text = "      ComCt: " & FrmMonit.COM_PLC.ErrorCount & " - " & FrmMonit.COM_PLC.PollCount & " - " & (dtBatDonInt)
        Label7.Text = "      ComCt: " & FrmMonit.COM_PLC.ErrorCount & " - " & FrmMonit.COM_PLC.PollCount & " - " & (ReadDvPointer)
        '104.11.6
        If Not bCommFlag Then FrmMonit.LabelPlc.BackColor = Color.LightGray
        '101.12.22
        'If iQue1AlmDelay < dtBatDonInt Then
        '102.2.17 If iQue1AlmDelay < dtBatDonInt And queue1.ismatch = True Then
        '106.10.16
        Dim D839 As Integer
        '110.9.15
        If Not queue0.inprocess Then
            D839_No_Cnt = 0
            If (iQue1AlmDelay < dtBatDonInt And iQue1AlmDelay < dtBatDonIntQ1) Then
                '110.9.24
                'If (queue1.ismatch = True Or (Cbxquecode1.Text <> "" And numquetri1.Text <> "")) Then
                If (queue1.ismatch = True) Then
                    D839_Cnt += 1
                    '110.11.16
                    'If (D839_Cnt Mod 15) = 1 Then
                    If btnque1.ForeColor <> Color.Purple Then
                        btnque1.ForeColor = Color.Purple
                        DVW(839) = 1
                        i = FrmMonit.WriteA2N(839, 1)
                    End If
                    '110.9.24
                    If DVW(839) = 1 Then
                        If btnbegin.BackColor = Color.Yellow Then
                            btnbegin.BackColor = Color.Cyan
                        Else
                            btnbegin.BackColor = Color.Yellow
                        End If
                    End If
                End If
            End If
        Else
            D839_No_Cnt += 1
            D839_Cnt = 0
            '110.10.22 110.11.16
            'If (D839_No_Cnt Mod 30) = 1 Then
            If btnque1.ForeColor = Color.Purple Then
                btnque1.ForeColor = Color.Black
                btnbegin.BackColor = Color.Cyan
                DVW(839) = 0
                i = FrmMonit.WriteA2N(839, 1)
                '110.9.15
            End If
        End If
        D839 = DVW(839)


        '101.5.4 test 
        'If Second(Now) Mod 10 = 1 Then
        '    NumW1.Value += 1
        'ElseIf Second(Now) Mod 10 = 2 Then
        '    Numw2.Value += 1
        'ElseIf Second(Now) Mod 10 = 3 Then
        '    Numw3.Value += 1
        'ElseIf Second(Now) Mod 10 = 4 Then
        '    Numw4.Value += 1
        'ElseIf Second(Now) Mod 10 = 5 Then
        '    Numw5.Value += 1
        'End If

        '105.9.3 107.4.14
        'If bRadioButton_REMOTE Then
        If True Then
            If Now.Second Mod 10 = 0 Then
                Dim db_sum As DataTable
                '110.10.5
                Dim file
                file = My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb")
                If file = False Then
                    Exit Sub
                End If

                dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
                'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where unicar = '" + queueP.UniCar + "'")
                'txt_carsum.Text = db_sum.Rows(0).Item(0).ToString
                'Dim m3 As Single
                'm3 = Val(txt_carsum.Text)
                'txt_carsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
                '填入每日完成方數()
                Dim s As String
                ssSavingDate = CDate(sSavingDate).ToShortDateString
                s = ssSavingDate

                '105.11.9 107.4.14
                'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")
                'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")

                'db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m/d') = '" + s + "'")
                'Dim sum, d
                'sum = 0
                'For i = 0 To db_sum.Rows.Count - 1
                '    d = db_sum.Rows(i).Item(0)
                '    sum += Format(Val(db_sum.Rows(i).Item(1).ToString), "0.00")
                'Next
                'txt_daysum.Text = Format(sum, "0.00")

                SaveSetting("JS", "CBC800", "txt_daysum", txt_daysum.Text)
                '105.11.9 107.4.14 111.12.28
                'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m') = '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "'")
                'txt_monthsum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")

                'db_sum = dbrec.GetDataTable("select DISTINCT  RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m') = '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
                'sum = 0
                'For i = 0 To db_sum.Rows.Count - 1
                '    sum += Format(Val(db_sum.Rows(i).Item(1).ToString), "0.00")
                'Next
                'txt_monthsum.Text = Format(sum, "0.00")

                '107.3.6 111.12.28
                'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') =  '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "/" + CDate(sSavingDate).Day.ToString + "'")
                'txt_daysum.Text = Format(Val(db_sum.Rows(0).Item(0).ToString), "0.00")
                db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m/d') =  '" + Today.Year.ToString + "/" + Today.Month.ToString + "/" + Today.Day.ToString + "'")
                Dim cube_tot As Single = 0
                Dim j
                For j = 0 To db_sum.Rows.Count - 1
                    cube_tot += db_sum.Rows(j).Item("cube")
                Next
                txt_daysum.Text = Format(cube_tot, "0.00")

                db_sum = dbrec.GetDataTable("select DISTINCT RecDT,cube , savedt from recdata where format(savedt, 'yyyy/m') =  '" + Today.Year.ToString + "/" + Today.Month.ToString + "'")
                cube_tot = 0
                For j = 0 To db_sum.Rows.Count - 1
                    cube_tot += db_sum.Rows(j).Item("cube")
                Next
                txt_monthsum.Text = Format(cube_tot, "0.00")

            End If

        End If

    End Sub

    Private Sub lblmod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = DataGridView1.CurrentCell.RowIndex
        DataGridView1.Rows(i).Cells(0).Selected = True
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        chk_1_all.Checked = Not chk_1_all.Checked
        Call chk_1_all_CheckedChanged(DataGridView1, e)
    End Sub

    Private Sub CheckBoxContinue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxContinue.CheckedChanged
        If CheckBoxContinue.Checked = True Then
            '102.4.17
            'If bEngVer Then
            '    CheckBoxContinue.Text = "AUTO"
            'Else
            '    CheckBoxContinue.Text = "連續生產"
            'End If
            'CheckBoxContinue.BackColor = Color.LightPink
            CheckBoxContinue.BackgroundImage = My.Resources.autorun
        Else
            '102.4.17
            'If bEngVer Then
            '    CheckBoxContinue.Text = "MAN"
            'Else
            '    CheckBoxContinue.Text = "手動啟動"
            'End If
            'CheckBoxContinue.Text = "手動啟動"
            CheckBoxContinue.BackColor = Color.LightGray

            CheckBoxContinue.BackgroundImage = My.Resources.hand1

        End If

    End Sub





    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Dim x, y, w
        w = FrmMonit.dgvSilo1.Columns(0).Width + FrmMonit.dgvSilo1.Columns(1).Width + FrmMonit.dgvSilo1.Columns(2).Width
        x = FrmMonit.dgvSilo1.Columns(0).Width + FrmMonit.dgvSilo1.Columns(1).Width
        y = FrmMonit.dgvSilo1.Rows(0).Height + FrmMonit.dgvSilo1.Rows(1).Height + FrmMonit.dgvSilo1.Rows(2).Height + FrmMonit.dgvSilo1.Rows(3).Height + FrmMonit.dgvSilo1.Rows(4).Height + FrmMonit.dgvSilo1.Rows(5).Height
        'FrmMonit.DrawLines(x + w * 0, y, FrmMonit.dgvSilo1, Color.Yellow)
        'FrmMonit.DrawLines(x + w * 1, y, FrmMonit.dgvSilo1, Color.YellowGreen)
        'FrmMonit.DrawLines(x + w * 2, y, FrmMonit.dgvSilo1, Color.GreenYellow)
        'FrmMonit.DrawLines(x + w * 3, y, FrmMonit.dgvSilo1, Color.DarkGreen)
        'FrmMonit.DrawLines(x + w * 4, y, FrmMonit.dgvSilo1, Color.Cyan)
        'FrmMonit.DrawSiloOut1(1, FrmMonit.dgvSilo1, Color.YellowGreen)
    End Sub

    Private Sub FormMainLocate()
        Dim w As Integer
        Dim g As Integer
        Dim i, j As Integer

        '107.5.23 
        VGAP = (Me.Height - Label2.Height - Panel4.Height - Panel6.Height - dgvWater.Height - DataGridView2.Height - dgvScale.Height - dgvPan.Height - Panel5.Height - 10) / 5
        Panel4.Top = Label2.Top + Label2.Height + VGAP
        Panel6.Top = Panel4.Top + Panel4.Height + VGAP
        LabelWater.Top = Panel6.Top + Panel6.Height + VGAP
        dgvWater.Top = LabelWater.Top
        Numw8.Top = LabelWater.Top
        NumSG.Top = LabelWater.Top
        CheckBoxsSim.Top = LabelWater.Top
        PanelWeight.Top = LabelWater.Top
        DataGridView2.Top = dgvWater.Top + dgvWater.Height + 3
        LabelWeight.Top = DataGridView2.Top + DataGridView2.Height + 2
        LabelErr.Top = LabelWeight.Top + LabelWeight.Height + 2
        dgvScale.Top = LabelWeight.Top
        dgvPan.Top = dgvScale.Height + dgvScale.Top + VGAP
        LabelWater.Height = dgvWater.Height + 2
        LabelWeight.Height = dgvScale.Rows(0).Height + 2
        LabelWeight.BackColor = Color.DarkGray

        dgvPan.Left = Panel4.Left
        '101.7.25 dgvPan.Width = Me.Width * 0.5
        'dgvPan.Width = Me.Width * 0.6
        dgvPan.Width = Me.Width * 0.6
        For i = 0 To dgvPan.ColumnCount - 1
            dgvPan.Columns(i).Width = dgvPan.Width / 11
        Next
        dgvPan.RowCount = 2
        dgvPan.Rows(0).Height = 30
        dgvPan.Rows(1).Height = 30
        dgvPan.Height = dgvPan.Rows(0).Height + dgvPan.Rows(1).Height + dgvPan.ColumnHeadersHeight + 5

        Dim columnHeaderStyle As New DataGridViewCellStyle()

        'columnHeaderStyle.BackColor = Color.Beige
        'columnHeaderStyle.Font = New Font("標楷體", 16)
        '101.7.25
        'columnHeaderStyle.Font = New Font("全真顏體", 16)
        columnHeaderStyle.Font = New Font("全真顏體", 14)
        columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvPan.ColumnHeadersDefaultCellStyle = columnHeaderStyle
        'dgvPan.ColumnCount = 11
        'dgvPan.Columns(10).HeaderText = "十"

        'dgvPan.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvPan.Columns(10).DefaultCellStyle.Font = New Font("全真顏體", 16)
        '101.7.25
        'dgvPan.Rows(0).Cells(0).Style.Font = New Font("全真顏體", 16)
        'dgvPan.Rows(1).Cells(0).Style.Font = New Font("全真顏體", 16)
        '102.4.17
        If bEngVer Then
            columnHeaderStyle.Font = New Font("Arial Narrow", 16)
            dgvPan.Rows(0).Cells(0).Style.Font = New Font("Arial Narrow", 16)
            dgvPan.Rows(1).Cells(0).Style.Font = New Font("Arial Narrow", 16)
        Else
            columnHeaderStyle.Font = New Font("全真顏體", 14)
            dgvPan.Rows(0).Cells(0).Style.Font = New Font("全真顏體", 14)
            dgvPan.Rows(1).Cells(0).Style.Font = New Font("全真顏體", 14)
        End If

        dgvPan.Rows(1).Cells(9).Selected = True
        dgvPan.Rows(1).Cells(9).Style.BackColor = Color.LightYellow
        dgvPan.Rows(0).Cells(0).Style.ForeColor = Color.Green
        dgvPan.Rows(1).Cells(0).Style.ForeColor = Color.DeepPink
        dgvPan.Rows(0).Cells(0).Style.BackColor = Color.LightGreen
        dgvPan.Rows(1).Cells(0).Style.BackColor = Color.LightPink
        '102.4.17 102.8.13
        If bEngVer Then
            dgvPan.Rows(0).Cells(0).Value = "AGG"
            dgvPan.Rows(1).Cells(0).Value = "MUD"
        Else
            dgvPan.Rows(0).Cells(0).Value = "骨材"
            dgvPan.Rows(1).Cells(0).Value = "泥料"
        End If
        For j = 1 To 8
            dgvPan.Rows(0).Cells(j).Style.Font = New Font("Arial Narrow", 16)
            dgvPan.Rows(1).Cells(j).Style.Font = New Font("Arial Narrow", 16)
        Next
        '101.7.25 Label10.Left = dgvPan.Left + dgvPan.Width + 5
        Label10.Left = dgvPan.Left + dgvPan.Width + 1
        Label10.Top = dgvPan.Top
        Label10.Height = dgvPan.ColumnHeadersHeight

        '101.3.7
        Label12.Top = Label10.Top
        '101.7.25 Label12.Left = Label10.Left + Label10.Width + 5
        Label12.Left = Label10.Left + Label10.Width + 1
        Label12.Height = Label10.Height
        CheckBoxsSim.Top = PanelWeight.Top
        CheckBoxsSim.Left = PanelWeight.Left - CheckBoxsSim.Width - 5

        txtbone.Left = Label10.Left
        txtbone.Top = Label10.Top + Label10.Height + 2
        txtbone.Height = dgvPan.Rows(0).Height

        Txtmod.Left = Label10.Left
        Txtmod.Top = txtbone.Top + txtbone.Height + 2
        Txtmod.Height = dgvPan.Rows(1).Height
        txtremain.Left = Txtmod.Left + Txtmod.Width + 2
        txtremain.Top = Txtmod.Top
        LabelTotPlate.Left = txtremain.Left
        LabelTotPlate.Top = txtbone.Top
        chk_A_B.Top = Label10.Top
        chk_A_B.Left = Label10.Left + Label10.Width + 2

        LabelWeight.Left = DataGridView2.Left

        'PanelWeight.Left = dgvWater.Left + dgvWater.Width + 5
        PanelWeight.Left = Panel4.Left + Panel4.Width - PanelWeight.Width
        PanelWeight.Top = LabelWater.Top

        '101.3.7
        CheckBoxsSim.Top = PanelWeight.Top
        CheckBoxsSim.Left = PanelWeight.Left - CheckBoxsSim.Width - 5

        LabelAltD.Top = Label10.Top
        LabelAltR.Top = Label10.Top
        LabelAltD.Left = Label10.Left + Label10.Width * 1.05
        LabelAltR.Left = Label10.Left + Label10.Width * 1.45
        '101.7.25
        'LabelTotPlate.Top = txtbone.Top
        'LabelD20.Top = txtbone.Top
        'LabelTotPlate.Left = txtbone.Left + txtbone.Width * 1.05
        'LabelD20.Left = txtbone.Left + txtbone.Width * 1.45

        LabelTotPlate.Top = txtbone.Top
        LabelD20.Top = txtbone.Top + 15
        LabelTotPlate.Left = txtbone.Left + txtbone.Width + 2
        LabelD20.Left = txtbone.Left + txtbone.Width + 2


        Panel4.Width = Me.Width - 10
        Panel4.Left = 5
        '101.7.25 g = 3
        g = 2
        w = Panel4.Width / 12

        '101.7.25
        CheckBoxLink.Left = g
        '102.4.20
        'CheckBoxLink.Width = w * 0.9 - g
        CheckBoxContinue.Left = CheckBoxLink.Left
        CheckBoxContinue.Width = CheckBoxLink.Width
        CheckBoxContinue.Height = CheckBoxLink.Height
        btnstop.Left = CheckBoxContinue.Left + CheckBoxContinue.Width + g
        btncancel1.Left = CheckBoxContinue.Left + CheckBoxContinue.Width + g
        btncancel2.Left = CheckBoxContinue.Left + CheckBoxContinue.Width + g

        btnbegin.Left = w * 0.9
        btnbegin.Width = w * 1 - g
        lblwork.Left = btnbegin.Left
        lblwork.Width = btnbegin.Width
        btnque1.Left = btnbegin.Left
        btnque1.Width = btnbegin.Width
        btnque2.Left = btnbegin.Left
        btnque2.Width = btnbegin.Width
        If bCheckBoxSP1 Then
            btnbegin.Left = w * 0.9
            btnbegin.Width = w * 1 - g - 25
            lblwork.Left = btnbegin.Left
            lblwork.Width = btnbegin.Width
            btnque1.Left = btnbegin.Left
            btnque1.Width = btnbegin.Width
            btnque2.Left = btnbegin.Left
            btnque2.Width = btnbegin.Width
        End If

        lblcode.Left = w * 1.9
        '101.7.25
        'lblcode.Width = w * 1.3 - g
        '109.12.19 配比15碼 
        lblcode.Top = btnbegin.Top
        lblTri.Top = btnbegin.Top
        lblcar.Top = btnbegin.Top
        lblfield.Top = btnbegin.Top
        Label8.Top = btnbegin.Top
        Label9.Top = btnbegin.Top
        Label11.Top = btnbegin.Top

        Cbxquecode1.Top = lblwork.Top
        TbxworkTri.Top = lblwork.Top
        Tbxworkcar.Top = lblwork.Top
        tbxworkfield.Top = lblwork.Top
        lblCustomerNo.Top = lblwork.Top
        lblCustomerName.Top = lblwork.Top
        LabelCust.Top = lblwork.Top

        Cbxquecode1.Top = btnque1.Top
        numquetri1.Top = btnque1.Top
        Tbxquecar1.Top = btnque1.Top
        Tbxquefield1.Top = btnque1.Top
        tbxCusNo1.Top = btnque1.Top
        lblCusName1.Top = btnque1.Top
        LabelCust1.Top = btnque1.Top

        Cbxquecode2.Top = btnque2.Top
        numquetri2.Top = btnque2.Top
        Tbxquecar2.Top = btnque2.Top
        Tbxquefield2.Top = btnque2.Top
        tbxCusNo2.Top = btnque2.Top
        lblCusName2.Top = btnque2.Top
        LabelCust2.Top = btnque2.Top

        '109.12.19 配比15碼 
        If bCheckBoxSP17 Then
            lblcode.Width = w * 2.0 - g
            lblTri.Width = w * 0.7 - g
            lblcar.Width = w * 1.0 - g
            lblfield.Width = 1.3 * w - g
            Label8.Width = 1.5 * w - g
            Label9.Width = 1.4 * w - g
            Label11.Width = 1.4 * w - g
        Else
            lblcode.Width = w * 1.5 - g
            lblTri.Width = w * 0.9 - g
            lblcar.Width = w * 1.1 - g
            lblfield.Width = 1.5 * w - g
            Label8.Width = 1.6 * w - g
            Label9.Width = 1.5 * w - g
        End If
        tbxworkcode.Left = lblcode.Left
        tbxworkcode.Width = lblcode.Width
        Cbxquecode1.Left = lblcode.Left
        Cbxquecode1.Width = lblcode.Width
        Cbxquecode2.Left = lblcode.Left
        Cbxquecode2.Width = lblcode.Width
        '105.5.6
        If bCheckBoxSP1 Then
            lblcode.Left = w * 1.9 - 25
            '109.12.19 配比15碼 
            'lblcode.Width = w * 1.5 - g + 35
            tbxworkcode.Left = lblcode.Left
            tbxworkcode.Width = lblcode.Width
            Cbxquecode1.Left = lblcode.Left
            Cbxquecode1.Width = lblcode.Width
            Cbxquecode2.Left = lblcode.Left
            Cbxquecode2.Width = lblcode.Width
        End If

        '109.12.19 配比15碼 
        'lblTri.Left = w * 3.4
        'lblTri.Width = w * 0.9 - g
        lblTri.Left = lblcode.Left + lblcode.Width + g
        TbxworkTri.Left = lblTri.Left
        TbxworkTri.Width = lblTri.Width
        numquetri1.Left = lblTri.Left
        numquetri1.Width = lblTri.Width
        numquetri2.Left = lblTri.Left
        numquetri2.Width = lblTri.Width

        '105.5.6
        If bCheckBoxSP1 Then
            lblTri.Left = w * 3.4 + 10
            lblTri.Width = w * 0.9 - g - 10
            TbxworkTri.Left = lblTri.Left
            TbxworkTri.Width = lblTri.Width
            numquetri1.Left = lblTri.Left
            numquetri1.Width = lblTri.Width
            numquetri2.Left = lblTri.Left
            numquetri2.Width = lblTri.Width
        End If

        '109.12.19 配比15碼 
        'lblcar.Left = w * 4.3
        'lblcar.Width = w * 1.1 - g
        lblcar.Left = lblTri.Left + lblTri.Width + g
        Tbxworkcar.Left = lblcar.Left
        Tbxworkcar.Width = lblcar.Width
        Tbxquecar1.Left = lblcar.Left
        Tbxquecar1.Width = lblcar.Width
        Tbxquecar2.Left = lblcar.Left
        Tbxquecar2.Width = lblcar.Width

        '109.12.19 配比15碼 
        'lblfield.Left = w * 5.4
        'lblfield.Width = 1.5 * w - g
        lblfield.Left = lblcar.Left + lblcar.Width + g
        tbxworkfield.Left = lblfield.Left
        tbxworkfield.Width = lblfield.Width
        Tbxquefield1.Left = lblfield.Left
        '
        Tbxquefield1.Width = lblfield.Width
        Tbxquefield2.Left = lblfield.Left
        Tbxquefield2.Width = lblfield.Width

        '工程名稱
        '109.12.19 配比15碼 
        'Label8.Left = w * 6.9
        'Label8.Width = 1.6 * w - g
        Label8.Left = lblfield.Left + lblfield.Width + g
        lblCustomerNo.Left = Label8.Left
        lblCustomerNo.Width = Label8.Width
        tbxCusNo1.Left = Label8.Left
        tbxCusNo1.Width = Label8.Width
        tbxCusNo2.Left = Label8.Left
        tbxCusNo2.Width = Label8.Width

        '施工部位
        '109.12.19 配比15碼 
        'Label9.Left = w * 8.5
        'Label9.Width = 1.5 * w - g
        Label9.Left = Label8.Left + Label8.Width + g
        lblCustomerName.Left = Label9.Left
        lblCustomerName.Width = Label9.Width
        lblCusName1.Left = Label9.Left
        lblCusName1.Width = Label9.Width
        lblCusName2.Left = Label9.Left
        lblCusName2.Width = Label9.Width
        '106.10.7
        LabelPb12_0.Left = lblCustomerName.Left
        LabelPb12_1.Left = lblCusName1.Left
        LabelPb12_2.Left = lblCusName2.Left

        '客戶名稱
        '109.12.19 配比15碼 
        'Label11.Left = w * 10
        Label11.Left = Label9.Left + Label9.Width + g
        LabelCust.Left = Label11.Left
        LabelCust1.Left = Label11.Left
        LabelCust2.Left = Label11.Left
        'LabelQkey.Left = Label11.Left
        'LabelQkey1.Left = Label11.Left
        'LabelQkey2.Left = Label11.Left

        '101.3.7

        i = 0
        DataGridViewAlarm.RowCount = 10

        '110.4.8
        DataGridViewAlarm.DefaultCellStyle.SelectionForeColor = Color.Yellow
        DataGridViewAlarm.DefaultCellStyle.SelectionBackColor = Color.Yellow
        'For j = 0 To 5
        For j = 0 To 9
            'DataGridViewAlarm.Rows(j).Height = 28
            '101.7.25
            DataGridViewAlarm.Rows(j).Cells(0).Style.Font = New Font("Arial Narrow", 13)
            '102.5.11
            DataGridViewAlarm.Rows(j).Cells(1).Style.Font = New Font("全真顏體", 14)
            If bEngVer Then
                '104.5.2 104.8.27
                'DataGridViewAlarm.Rows(j).Cells(1).Style.Font = New Font("Arial Narrow", 12)
                DataGridViewAlarm.Rows(j).Cells(0).Style.Font = New Font("Arial Narrow", 12)
                DataGridViewAlarm.Rows(j).Cells(1).Style.Font = New Font("Arial Narrow", 14, FontStyle.Bold)
            End If
            DataGridViewAlarm.Rows(j).Height = 26
            i += DataGridViewAlarm.Rows(j).Height
        Next
        '101.7.25 
        'DataGridViewAlarm.Top = dgvScale.Top + dgvScale.Height + 5
        'DataGridViewAlarm.Height = i + DataGridViewAlarm.ColumnHeadersHeight + 5
        'DataGridViewAlarm.Left = Me.Width - DataGridViewAlarm.Width - 5

        DataGridViewAlarm.ColumnHeadersDefaultCellStyle.Font = New Font("全真顏體", 14)
        '102.5.11
        If bEngVer Then
            DataGridViewAlarm.ColumnHeadersDefaultCellStyle.Font = New Font("Arial Narrow", 12)
            '104.5.2 104.8.27
            DataGridViewAlarm.Columns(0).DefaultCellStyle.Font = New Font("Arial Narrow", 12, FontStyle.Italic)
            DataGridViewAlarm.Columns(1).DefaultCellStyle.Font = New Font("Arial", 14, FontStyle.Bold)
        End If
        'DataGridViewAlarm.Columns(0).DefaultCellStyle.Font = New Font("全真顏體", 12)
        'DataGridViewAlarm.Columns(1).DefaultCellStyle.Font = New Font("全真顏體", 12)
        DataGridViewAlarm.Top = dgvScale.Top + dgvScale.Height + 2
        DataGridViewAlarm.Height = i + DataGridViewAlarm.ColumnHeadersHeight + 2

        DataGridViewAlarm.Top = dgvScale.Top + dgvScale.Height + 2
        DataGridViewAlarm.Height = i + DataGridViewAlarm.ColumnHeadersHeight + 2

        '104.5.2 104.8.27
        DataGridViewAlarm.Columns(0).Width = 90
        DataGridViewAlarm.Columns(1).Width = 260

        DataGridViewAlarm.Width = DataGridViewAlarm.Columns(0).Width + DataGridViewAlarm.Columns(1).Width + 5
        DataGridViewAlarm.Left = Me.Width - DataGridViewAlarm.Width - 2
        '101.7.25 
        LabelA1.Top = DataGridViewAlarm.Top + 5
        LabelA2.Top = DataGridViewAlarm.Top + 5
        LabelA3.Top = DataGridViewAlarm.Top + 5
        LabelA3.Left = DataGridViewAlarm.Left + DataGridViewAlarm.Width - LabelA3.Width - 5
        LabelA2.Left = LabelA3.Left + -LabelA2.Width - 5
        LabelA1.Left = LabelA2.Left + -LabelA1.Width - 5

        '109.5.7 no use 
        'If CBool(bCheckBoxSingle) Then
        '    Panel5.Top = DataGridViewAlarm.Top + DataGridViewAlarm.Height + VGAP
        '    Panel5.Width = DataGridViewAlarm.Width
        '    Panel5.Left = DataGridViewAlarm.Left
        '    txt_carsum.Width = 90
        '    txt_daysum.Width = 90
        '    Label3.Left = Panel5.Width * 0.4
        '    lbl_monthsum.Left = Label3.Left
        '    PanelSim.Left = Panel1.Left
        '    PanelSim.Top = Panel1.Top
        '    txt_monthsum.Left = Label3.Left + Label3.Width
        '    lblMachTot.Left = txt_monthsum.Left
        '    '101.3.14
        '    '107.5.21
        '    'PanelT.Top = dgvPan.Top + dgvPan.Height + VGAP
        '    'Panel10.Top = VGAP
        '    'PanelMixer.Top = Panel10.Height + VGAP * 2
        '    'PanelT.Height = PanelMixer.Height + Panel10.Height + VGAP * 3
        '    'PanelT.Visible = True
        'Else
        '    Panel5.Top = dgvPan.Top + dgvPan.Height + VGAP
        '    Panel5.Width = Me.Width - DataGridView1.Width - 20
        '    Panel5.Left = Panel4.Left
        '    '107.5.21
        '    Panel5.Height = Me.Height - Panel5.Top - 10
        'End If
        Panel5.Top = dgvPan.Top + dgvPan.Height + VGAP
        Panel5.Width = Me.Width - DataGridView1.Width - 20
        Panel5.Left = Panel4.Left
        Panel5.Height = Me.Height - Panel5.Top - 10

        Panel6.Width = Panel4.Width
        Panel6.Left = Panel4.Left

        NumW1.Left = dgvWater.Left
        NumW1.Height = dgvWater.Height
        NumW1.Top = dgvWater.Top
        NumW1.Width = dgvWater.Columns(0).Width
        Dim s As String
        s = dgvWater.Rows(0).Cells(0).Value
        NumW1.Value = CSng(dgvWater.Rows(0).Cells(0).Value)
        NumW1.Visible = dgvWater.Columns(0).Visible
        NumW1.BackColor = Color.SkyBlue

        j = NumW1.Left
        i = dgvWater.Columns(0).Width
        If NumW1.Visible = True Then j += i
        Numw2.Left = j
        Numw2.Height = dgvWater.Height
        Numw2.Top = dgvWater.Top
        Numw2.Width = dgvWater.Columns(0).Width
        Numw2.Value = CSng(dgvWater.Rows(0).Cells(1).Value)
        Numw2.Visible = dgvWater.Columns(1).Visible
        Numw2.BackColor = Color.SkyBlue

        If Numw2.Visible = True Then j += i
        Numw3.Left = j
        Numw3.Height = dgvWater.Height
        Numw3.Top = dgvWater.Top
        Numw3.Width = dgvWater.Columns(0).Width
        Numw3.Value = CSng(dgvWater.Rows(0).Cells(2).Value)
        Numw3.Visible = dgvWater.Columns(2).Visible
        Numw3.BackColor = Color.SkyBlue

        If Numw3.Visible = True Then j += i
        Numw4.Left = j
        Numw4.Height = dgvWater.Height
        Numw4.Top = dgvWater.Top
        Numw4.Width = dgvWater.Columns(0).Width
        Numw4.Value = CSng(dgvWater.Rows(0).Cells(3).Value)
        Numw4.Visible = dgvWater.Columns(3).Visible
        Numw4.BackColor = Color.SkyBlue

        If Numw4.Visible = True Then j += i
        Numw5.Left = j
        Numw5.Height = dgvWater.Height
        Numw5.Top = dgvWater.Top
        Numw5.Width = dgvWater.Columns(0).Width
        Numw5.Value = CSng(dgvWater.Rows(0).Cells(4).Value)
        Numw5.Visible = dgvWater.Columns(4).Visible
        Numw5.BackColor = Color.SkyBlue

        If Numw5.Visible = True Then j += i
        Numw6.Left = j
        Numw6.Height = dgvWater.Height
        Numw6.Top = dgvWater.Top
        Numw6.Width = dgvWater.Columns(0).Width
        Numw6.Value = CSng(dgvWater.Rows(0).Cells(5).Value)
        Numw6.Visible = dgvWater.Columns(5).Visible
        Numw6.BackColor = Color.SkyBlue

        If Numw6.Visible = True Then j += i
        Numw7.Left = j
        Numw7.Height = dgvWater.Height
        Numw7.Top = dgvWater.Top
        Numw7.Width = dgvWater.Columns(0).Width
        Numw7.Value = CSng(dgvWater.Rows(0).Cells(6).Value)
        Numw7.Visible = dgvWater.Columns(6).Visible
        Numw7.BackColor = Color.SkyBlue

        '101.4.2 max 8 bone
        If Numw7.Visible = True Then j += i
        Numw8.Left = j
        Numw8.Height = dgvWater.Height
        Numw8.Top = dgvWater.Top
        Numw8.Width = dgvWater.Columns(0).Width
        Numw8.Value = CSng(dgvWater.Rows(0).Cells(7).Value)
        Numw8.Visible = dgvWater.Columns(7).Visible
        Numw8.BackColor = Color.SkyBlue

        '104.9.24 NumSG
        If Numw8.Visible = True Then j += i
        NumSG.Left = j + i
        NumSG.Height = dgvWater.Height
        NumSG.Top = dgvWater.Top
        NumSG.Width = dgvWater.Columns(0).Width
        NumSG.BackColor = Color.SkyBlue
        '104.9.26
        '111.4.2 LabelAE NumAE
        'LabelSA.Left = NumSG.Left + NumSG.Width
        'LabelSA.Top = NumSG.Top
        LabelSA.Left = NumSG.Left
        LabelSA.Top = NumSG.Top - LabelSA.Height

        '111.4.2 LabelAE NumAE
        NumAE.Left = NumSG.Left + dgvWater.Columns(0).Width
        NumAE.Height = NumSG.Height
        NumAE.Top = NumSG.Top
        NumAE.Width = NumSG.Width
        NumAE.BackColor = NumSG.BackColor
        LabelAE.Left = NumAE.Left
        LabelAE.Top = NumAE.Top - LabelAE.Height

        LabelB.Left = lblcode.Left + lblcode.Width * 0.5

        Dim dt As DataTable = db.GetDataTable("select 全滿刻度 from config where type= 'Pond' order by Idx")
        For i = 0 To 10
            DVW(401 + i) = CInt(dt.Rows(i).Item(0))
            SiloFullScale(i) = dt.Rows(i).Item(0)
        Next
        For i = 11 To 14
            DVW(401 + i) = CInt(dt.Rows(i).Item(0) * 100)
            SiloFullScale(i) = dt.Rows(i).Item(0)
        Next
        i = FrmMonit.WriteA2N(401, 15)

        '101.7.25
        'LabelJS.Left = Panel2.Width + 30
        'LabelJS.Top = Panel2.Top + 10
        'Label7.Left = LabelJS.Left
        'Label7.Top = LabelJS.Top + 20
        '103.6.7
        'LabelJS.Left = DataGridViewAlarm.Left
        'LabelJS.Top = DataGridViewAlarm.Top + DataGridViewAlarm.Height + 1
        '104.5.2  104.8.27
        'LabelJS.Left = DataGridViewAlarm.Left - 200
        'LabelJS.Top = DataGridViewAlarm.Top + DataGridViewAlarm.Height * 0.8
        '105.7.25
        'LabelRPC.Left = DataGridViewAlarm.Left + DataGridViewAlarm.Width - LabelRPC.Width
        'LabelRPC.Top = LabelJS.Top

        '103.4.23
        'Panel2.Top = Me.Height - Panel2.Height - 5
        'Panel2.Top = Me.Height - Panel2.Height - 20
        'Panel2.Left = Panel5.Left
        '107.5.21
        Panel2.Top = Panel5.Height - Panel2.Height - 5
        Panel2.Left = Panel1.Left
        Label1SQL.Left = Panel2.Left
        Label1SQL.Top = Panel2.Top - 1 - Label1SQL.Height
        '112.9.26
        LabelSendSql.Left = Panel1.Left + Panel1.Width + 1
        LabelSendSql.Top = LabelSendSql.Top
        '112.10.25
        LabelSendSqlBig.Left = LabelSendSql.Left
        LabelSendSqlBig.Top = LabelSendSql.Top + 12

        '106.9.3 107.5.21
        'LabelJS.Left = DataGridViewAlarm.Left
        'LabelJS.Top = DataGridViewAlarm.Top + DataGridViewAlarm.Height
        'Label7.Left = LabelJS.Left + LabelJS.Width
        'Label7.Top = LabelJS.Top
        Label7.Top = Panel2.Top - Label7.Height
        Label7.Left = LabelJS.Left - Label7.Width - 20
        LabelJS.Left = Panel2.Width + 5
        LabelJS.Top = Panel5.Height - LabelJS.Height - 10
        '107.7.9
        'LabelRPC.Left = LabelJS.Left + LabelJS.Width + 10
        LabelRPC.Left = LabelJS.Left + LabelJS.Width - 10
        LabelRPC.Top = LabelJS.Top
        '110.3.30
        LabelBigDataErr.Top = LabelRPC.Top
        LabelBigDataErr.Left = LabelRPC.Left + LabelRPC.Width

        'DataGridViewAlarm.Height = Me.Height - DataGridViewAlarm.Top - 10
        DataGridViewAlarm.Top = Me.Height - DataGridViewAlarm.Height - 10
        LabelA1.Top = DataGridViewAlarm.Top + 5
        LabelA2.Top = DataGridViewAlarm.Top + 5
        LabelA3.Top = DataGridViewAlarm.Top + 5
        LabelA3.Left = DataGridViewAlarm.Left + DataGridViewAlarm.Width - LabelA3.Width - 5
        LabelA2.Left = LabelA3.Left + -LabelA2.Width - 5
        LabelA1.Left = LabelA2.Left + -LabelA1.Width - 5

        '107.6.2 車次顯示
        Dim SN As String
        Dim sql As String
        Dim k
        SN = ""
        iCarSerToday = 0
        db.changedb(sYMdbName + sSavingYear + ".mdb")
        sql = "select SaveDT, UniCar from recdata where  format(SaveDT, 'yyyy/M/d') = '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "/" + CDate(sSavingDate).Day.ToString + "'"
        '113.2.14
        Try
            dt = db.GetDataTable(sql)
            For k = 0 To dt.Rows.Count - 1
                If SN = dt.Rows(k).Item("UniCar") Then
                Else
                    iCarSerToday += 1
                    SN = dt.Rows(k).Item("UniCar")
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        LabelCarSer.Text = iCarSerToday

        '107.6.2 add PanelMon
        PanelMon.Left = Panel1.Left + Panel1.Width + 5
        PanelMon.Top = Panel1.Top
        PanelMon.Width = Panel5.Width - Panel1.Width - 5
        Label_D20.Text = "D20:"
        Label_D21.Text = "D21:"
        Label_D22.Text = "D22:"
        'LabelMsg.Left = Panel2.Left
        'LabelMsg.Top = Panel2.Top - 1 - Label1SQL.Height - LabelMsg.Height

        '107.6.10   Label1SaveEnergy
        Label1SaveEnergy.Left = lblfield.Left
        Label1SaveEnergy.Top = tbxworkfield.Top + tbxworkfield.Height + 1
        'Label1SaveEnergy.Visible = True
        LabelUniCar.Top = Label11.Top + 2
        LabelCarSer.Top = LabelUniCar.Top + LabelUniCar.Height + 1
        LabelUniCar.Left = Me.Width - LabelUniCar.Width - 50
        LabelCarSer.Left = LabelUniCar.Left
        '107.6.16
        'Label1SaveEnergy2.Visible = True
        Label1SaveEnergy2.Left = Label1SaveEnergy.Left
        Label1SaveEnergy2.Top = Label1SaveEnergy.Top + Label1SaveEnergy.Height + 1


        '110.9.11 110.6.24
        PanelAlarm.Width = DataGridViewAlarm.Width
        PanelAlarm.Height = DataGridViewAlarm.Height
        PanelAlarm.Top = DataGridViewAlarm.Top
        PanelAlarm.Left = DataGridViewAlarm.Left
        DataGridViewAlarm.Visible = False
        For i = 0 To 9
            With LabelA20
                labelAlarmTime(i) = New Label
                labelAlarmTime(i).Font = .Font
                labelAlarmTime(i).ForeColor = .ForeColor
                labelAlarmTime(i).BackColor = .BackColor
                labelAlarmTime(i).Visible = True
                PanelAlarm.Controls.Add(labelAlarmTime(i))
                labelAlarmTime(i).Height = 24
                labelAlarmTime(i).Width = 80
                labelAlarmTime(i).Left = 3
                labelAlarmTime(i).Top = 30 + 25 * i
                labelAlarmTime(i).Text = ""
            End With
            With LabelA21
                labelAlarmDesc(i) = New Label
                labelAlarmDesc(i).Font = .Font
                labelAlarmDesc(i).ForeColor = .ForeColor
                labelAlarmDesc(i).BackColor = .BackColor
                labelAlarmDesc(i).Visible = True
                PanelAlarm.Controls.Add(labelAlarmDesc(i))
                labelAlarmDesc(i).Left = labelAlarmTime(i).Left + labelAlarmTime(i).Width + 1
                labelAlarmDesc(i).Top = labelAlarmTime(i).Top
                labelAlarmDesc(i).Height = labelAlarmTime(i).Height
                labelAlarmDesc(i).Width = PanelAlarm.Width - labelAlarmTime(i).Width - 8
                labelAlarmDesc(i).Text = ""
            End With
        Next
        LabelA20.Visible = False
        LabelA21.Visible = False


    End Sub

    Private Sub btnDynamic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDynamic.Click
        frmDynamic.ShowDialog()

    End Sub



    Private Sub lblwork_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblwork.Click

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        '102.5.11 ORIG CBC800 EPORT
        'Dim lineHeight As Single
        'Dim myFont = New Font("全真仿宋體", 9)
        'If MatCounts >= 17 Then
        '    myFont = New Font("全真仿宋體", 9)
        'Else
        '    myFont = New Font("全真仿宋體", 10)
        'End If
        'lineHeight = myFont.getheight(e.Graphics)
        'e.Graphics.DrawString(FrmHistory.repdetail, myFont, Brushes.Black, 1, 1)

        e.Graphics.DrawString(FrmHistory.repdetail, repFont, Brushes.Black, 1, 1)

    End Sub


    Private Sub tbxmixtime_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxmixtime.TextChanged
        If tbxmixtime.Text = "" Then tbxmixtime.Text = 50

    End Sub

    Private Sub tbxmixtime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbxmixtime.Validated
        If tbxmixtime.Text = "" Then tbxmixtime.Text = 50
        DVW(832) = CInt(tbxmixtime.Text)
        i = FrmMonit.WriteA2N(832, 1)
        FrmMonit.tbxmixtime.Text = tbxmixtime.Text
    End Sub

    Private Sub Num_quetime_set_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Num_quetime_set.Validated
        DVW(834) = CInt(Num_quetime_set.Text)
        i = FrmMonit.WriteA2N(834, 1)

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        showform(FrmHistory)
    End Sub






    Private Sub dgvWater_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWater.CellValueChanged
        If FormLoded = False Then Exit Sub
        If Val(dgvWater.CurrentCell.Value) > 20 Then
            dgvWater.CurrentCell.Value = 20
        End If
        dgvWater.CurrentCell.Value = Format(Val(dgvWater.CurrentCell.Value), "0.0")
        If queue0.inprocess = True Then
            If IsNumeric(dgvWater.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                Dim x As Single = dgvWater.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim index As Integer = e.ColumnIndex
                'mc.queclone(queue0, queue)
                'waterrecount(x, index, queue)
                Dim k As Integer
                For k = 0 To 7
                    waterall(k) = dgvWater.Rows(0).Cells(k).Value
                Next
                mc.queclone(queue0, queue)
                '99.12.28
                queue.BoneDoingPlate = queue0.BoneDoingPlate
                queue.ModDoingPlate = queue0.ModDoingPlate
                '107.7.24 bug when queue1.inprocess still water              
                'If queue0.BoneDoingPlate = queue0.ModDoingPlate Then
                If Not queue1.inprocess And queue0.BoneDoingPlate = queue0.ModDoingPlate Then
                    queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
                Else
                    queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
                End If

                mc.queclone(queue, queue0)
                MonLog.add = "displaymaterial() : dgvWater"
                displaymaterial()
                'send material

                '101.3.20
                Try

                    If queue1.inprocess = True Then
                        mc.queclone(queue1, queue)
                        queue = mc.Materialdistribute(queue, queue1.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
                        mc.queclone(queue, queue1)
                        MonLog.add = "displaymaterial() : dgvWater,queue1"
                        displaymaterial()
                        'send material
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "dgvwater-1")
                End Try
            Else
                dgvWater.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = watertemp
            End If
        End If
        dgvWater.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = False

        Dim i, j As Integer
        i = dgvWater.Columns.Count
        Dim fileNum
        Dim s As String
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\water.per", OpenMode.Output)
        For j = 0 To 7
            s = dgvWater.Rows(0).Cells(j).Value
            PrintLine(fileNum, Trim(s))
        Next
        FileClose(fileNum)

        '104.8.18
        If bDefaultWatEn Then
            For j = 0 To 7
                fWat(j, iDefaultWatInd) = dgvWater.Rows(0).Cells(j).Value
            Next
        End If

        '112.3.20
        Dim iCarId(7) As Integer
        For i = 0 To 7
            iCarId(i) = CSng(dgvWater.Rows(0).Cells(i).Value) * 10
            DVW(4011 + i) = iCarId(i)
        Next
        i = FrmMonit.WriteA2N(4011, 8)

    End Sub




    Private Sub lblcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblcode.Click
        If queue0.inprocess = True Then
            LabelB.Text = queue0.showdata(1)
            LabelB.Visible = True
        End If
    End Sub

    Private Sub lblcode_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblcode.MouseLeave
        LabelB.Visible = False
    End Sub

    Private Sub Num_quetime_set_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Num_quetime_set.ValueChanged
        DVW(834) = CInt(Num_quetime_set.Value)
        i = FrmMonit.WriteA2N(834, 1)
        SaveSetting("JS", "CBC800", "Num_quetime_set", Num_quetime_set.Value)

    End Sub

    Private Sub NumW1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumW1.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(0).Value = NumW1.Value
        NumW1.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(0) = NumW1.Value
    End Sub

    Private Sub NumW1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumW1.ValueChanged
        NumW1.BackColor = Color.Yellow
    End Sub

    Private Sub Numw2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw2.Leave

        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(1).Value = Numw2.Value
        Numw2.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(1) = Numw2.Value
    End Sub

    Private Sub Numw2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw2.ValueChanged
        Numw2.BackColor = Color.Yellow
    End Sub

    Private Sub Numw3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw3.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(2).Value = Numw3.Value
        Numw3.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(2) = Numw3.Value
    End Sub

    Private Sub Numw3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw3.ValueChanged
        Numw3.BackColor = Color.Yellow
    End Sub

    Private Sub Numw4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw4.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(3).Value = Numw4.Value
        Numw4.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(3) = Numw4.Value
    End Sub

    Private Sub Numw4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw4.ValueChanged
        Numw4.BackColor = Color.Yellow
    End Sub

    Private Sub Numw5_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw5.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(4).Value = Numw5.Value
        Numw5.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(4) = Numw5.Value
    End Sub

    Private Sub Numw5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw5.ValueChanged
        Numw5.BackColor = Color.Yellow
    End Sub

    Private Sub Numw6_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw6.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(5).Value = Numw6.Value
        Numw6.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(5) = Numw6.Value
    End Sub

    Private Sub Numw6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw6.ValueChanged
        Numw6.BackColor = Color.Yellow
    End Sub

    Private Sub Numw7_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw7.Leave
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(6).Value = Numw7.Value
        Numw7.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(6) = Numw7.Value
    End Sub

    Private Sub Numw7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw7.ValueChanged
        Numw7.BackColor = Color.Yellow
    End Sub

    Private Sub Numw8_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Numw8.Leave
        '101.4.2
        If FormLoded = False Then Exit Sub
        dgvWater.Rows(0).Cells(7).Value = Numw8.Value
        Numw8.BackColor = Color.SkyBlue
        '104.8.10
        If Not bDefaultWatEn Then fWatOrig(7) = Numw8.Value
    End Sub
    Private Sub Numw8_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Numw8.ValueChanged
        '101.4.2
        Numw8.BackColor = Color.Yellow
    End Sub


    Private Sub CheckBoxLink_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLink.CheckedChanged
        '106.9.3
        '106.12.10
        'Label1SQL.Text = Now
        'Label1SQL.ForeColor = Color.Purple
        'Label1SQL.Refresh()
        'For i = 1 To 100000
        '    Application.DoEvents()
        'Next
        'Label1SQL.Text = Now
        'Label1SQL.ForeColor = Color.Purple
        'Label1SQL.Refresh()

        If CheckBoxLink.Checked = True Then
            '106.12.10
            Label1SQL.Text = ".."
            Label1SQL.ForeColor = Color.Purple
            Label1SQL.Refresh()
            '105.6.12 
            ' 105.10.8 TimerPing.Enabled = True
            If My.Computer.Network.IsAvailable = False Then
                '103.3.15
                If bEngVer Then
                    MsgBox("Network Fail! Please check Network.", MsgBoxStyle.Critical)
                Else
                    MsgBox("網路無法使用!請檢查網路設備.", MsgBoxStyle.Critical)
                End If
                bDataBaseLink = False
                CheckBoxLink.Checked = False
                '106.9.3
                CheckBoxLink.BackgroundImage = My.Resources.keyboard2
                Exit Sub
            End If
            If My.Computer.Network.Ping(sDataBaseIP) = False Then
                '103.3.15
                If bEngVer Then
                    MsgBox("IP:" & sDataBaseIP & ",connect Fail! Please check network and database computer.", MsgBoxStyle.Critical)
                Else
                    MsgBox("調度電腦IP:" & sDataBaseIP & ",連線失敗! 請檢查網路設備和調度電腦.", MsgBoxStyle.Critical)
                End If
                bDataBaseLink = False
                CheckBoxLink.Checked = False
                '106.9.3
                CheckBoxLink.BackgroundImage = My.Resources.keyboard2
                Exit Sub
            Else
                '106.12.10 no TimerPing
                'TimerPing.Enabled = True
                TimerPing.Enabled = False

                bPingSQL = True

                bDataBaseLink = True
                '102.4.20
                'CheckBoxLink.Text = "連線"
                CheckBoxLink.BackgroundImage = My.Resources.network2
                CheckBoxLink.BackColor = Color.Gold
                '107.3.20
                TimerLink.Interval = 1000
                TimerLink.Enabled = True
                '104.9.24 by LabelMessage.Text = "骨材搶先中" 
                'Cbxquecode1.Enabled = False
                'numquetri1.Enabled = False
                'Tbxquecar1.Enabled = False
                'Cbxquecode2.Enabled = False
                'numquetri2.Enabled = False
                'Tbxquecar2.Enabled = False
                'btnque1.Enabled = False
                'btncancel1.Enabled = False
                'btnque2.Enabled = False
                'btncancel2.Enabled = False

            End If
            '102.10.19
            If bEngVer Then
                Label8.Text = "Project Name"
                Label9.Text = cDESTINATION
                Label11.Text = cCUSTOMER
            End If
        Else
            '104.8.27 DV(36) bit 1 自動中 for 連線中 不可切換回
            If bAuto = True Then
                If bEngVer Then
                    MsgBox("Batching! No Ching. ", MsgBoxStyle.Critical)
                Else
                    MsgBox("生產程序自動中!不可切換.", MsgBoxStyle.Critical)
                End If
                Exit Sub
            End If
            bDataBaseLink = False
            '102.4.20
            'CheckBoxLink.Text = "輸入"
            CheckBoxLink.BackgroundImage = My.Resources.keyboard2

            CheckBoxLink.BackColor = Color.LightGray
            Cbxquecode1.Enabled = True
            numquetri1.Enabled = True
            Tbxquecar1.Enabled = True
            Cbxquecode2.Enabled = True
            numquetri2.Enabled = True
            Tbxquecar2.Enabled = True
            '100.6.29
            btnque1.Enabled = True
            btncancel1.Enabled = True
            '101.2.25
            btnque2.Enabled = True
            btncancel2.Enabled = True
            '102.10.19
            If bEngVer Then
                Label8.Text = ""
                Label9.Text = ""
                Label11.Text = ""
            End If
        End If

    End Sub

    Private Sub TimerLink_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerLink.Tick
        '106.8.26
        '106.8.26 try IP before Link 106.10.16 revise check network
        'If Not My.Computer.Network.IsAvailable Or Not My.Computer.Network.Ping(sDataBaseIP, 1000) Then
        '    bDataBaseLink = False
        '    CheckBoxLink.Checked = False
        '    CheckBoxLink.BackgroundImage = My.Resources.keyboard2
        '    'CheckBoxLink.BackColor = Color.LightGray
        'End If

        '106.11.26
        TimerLink.Enabled = False
        '107.3.20
        TimerLink.Interval = 8000

        '112.9.18 disable check IP again
        'If Not My.Computer.Network.IsAvailable Then
        '    bDataBaseLink = False
        '    CheckBoxLink.Checked = False
        '    CheckBoxLink.BackgroundImage = My.Resources.keyboard2
        '    TimerLink.Enabled = False
        '    bNetPing = False
        '    Exit Sub
        'Else
        '    bNetPing = True
        '    '106.11.20 disable check  106.12.10 just no messace
        '    If Not My.Computer.Network.Ping(sDataBaseIP, 1000) Then
        '        '106.12.10
        '        'MsgBox("Ping調度電腦IP:" & sDataBaseIP & "失敗! !", MsgBoxStyle.Critical, "網路檢查")
        '        TimerLinkErrorCnt = TimerLinkErrorCnt + 1
        '        Label1SQL.Text = "Ping調度電腦IP:" & sDataBaseIP & " 失敗!! < " & TimerLinkErrorCnt & " > ... " & Now
        '        Label1SQL.ForeColor = Color.Red
        '        Label1SQL.Refresh()
        '        '107.4.11
        '        DbLog.add = (Label1SQL.Text)
        '        If TimerLinkErrorCnt >= 3 Then
        '            bDataBaseLink = False
        '            bNetPing = False
        '            CheckBoxLink.Checked = False
        '            CheckBoxLink.BackgroundImage = My.Resources.keyboard2
        '            TimerLink.Enabled = False
        '            MsgBox("Ping調度電腦IP:" & sDataBaseIP & "失敗! !", MsgBoxStyle.Critical, "網路檢查")
        '            '107.4.11
        '            DbLog.add = ("Ping調度電腦IP:" & sDataBaseIP & "失敗! !")
        '        Else
        '            TimerLink.Enabled = True
        '            bDataBaseLink = True
        '            bNetPing = True
        '        End If
        '        Exit Sub
        '    Else
        '        TimerLinkErrorCnt = 0
        '    End If
        'End If
        If bDataBaseLink Then
            'Dim db As New SQLDbAccess("10.18.92.1", "conscbuji")
            '100.2.28 Dim db As New SQLDbAccess("conscbuji")
            Dim db
            '106.12.10
            'Dim dt As Data.DataTable
            Dim dt As New DataTable
            Dim sql As String = ""
            Dim SqlExpressDataSource As String
            'SqlExpressDataSource = sDataBaseIP & "\SQLEXPRESS"
            SqlExpressDataSource = sDataBaseIP
            db = New SQLDbAccess("conscbuji")

            sbujiq1a = "bujiq" & sMixerNo & "a"
            '111.1.7
            'sql = "Select  * from " & sbujiq1a & " where bujiread <> 'Y'  order by pb33 asc, pb12 asc, queuekey asc "
            sql = "Select TOP 1 * from " & sbujiq1a & " where bujiread <> 'Y'  order by pb33 asc, pb12 asc, queuekey asc "

            '100.2.28
            If bDataBaseLink = False Then Exit Sub
            '106.7.25
            'Dim dc As New SqlClient.SqlCommand(sql)
            'dt = db.GetDataTable(sql)
            Try
                '106.11.26
                'db.ExecuteCmd(sql)
                dt = db.GetDataTable(sql)
            Catch ex As Exception
                MsgBox("Read bujiq1a 'Y' Fail!")
                '106.7.25
                bDataBaseLink = False
                '106.9.3
                TimerLink.Enabled = False
                '106.11.20
                'LabelJS.Image = My.Resources.disconnect
                '107.7.9 no image = > color
                'LabelJS.Image = My.Resources.disconnect
                CheckBoxLink.BackgroundImage = My.Resources.keyboard2
                Exit Sub
            End Try

            '107.8.31
            'If sDataBaseIP = "127.0.0.1" Then
            '    TimerLink.Enabled = False
            '    Exit Sub
            'Else
            '    TimerLink.Enabled = True
            'End If
            TimerLink.Enabled = True

            '107.3.18
            '105.8.21 for debug repeat
            'Call SavingLog(sql & " dt.Rows.Count:" & dt.Rows.Count)
            MainLog.add = "TimerLink() : " & sql & " dt.Rows.Count:" & dt.Rows.Count

            '106.7.25
            If dt.Rows.Count = 0 Then Exit Sub

            '106.10.11
            If dt.Rows.Count >= 1 Then
                Dim yy, mm
                Dim MixYn
                '106.10.17 聯佶空單
                'MixYn = dt.Rows(0).Item("mixyn").ToString
                If bCheckBoxSP13 Then
                    MixYn = dt.Rows(0).Item("mixyn").ToString
                Else
                    '106.10.20
                    'MixYn = "N"
                    MixYn = "Y"
                End If
                yy = Year(CDate(dt.Rows(0).Item("pb12").ToString))
                mm = Month(CDate(dt.Rows(0).Item("pb12").ToString))
                '106.10.20
                'If Not (yy = Now.Year And mm = Now.Month) And (MixYn = "Y" Or MixYn = "y") Then
                If Not (yy = Now.Year And mm = Now.Month) And (MixYn = "N" Or MixYn = "n") Then
                    sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
                    Try
                        db.ExecuteCmd(sql)
                    Catch ex As Exception
                        MsgBox("更新 bujiq1a to 'Y' 失敗!")
                    End Try
                    MsgBox("自動模擬回傳" & vbCrLf & "配比 : " & (dt.Rows(0).Item("formulaid").ToString) & vbCrLf & _
                    "方數 : " & (dt.Rows(0).Item("Pbqty").ToString) & vbCrLf & _
                    "車號 : " & (dt.Rows(0).Item("Carid").ToString) & vbCrLf & _
                    "客戶 : " & (dt.Rows(0).Item("Custname").ToString) & vbCrLf & _
                    "日期: " & vbCrLf & CDate(dt.Rows(0).Item("pb12").ToString) & vbCrLf & " 必須是本月份 !" & vbCrLf & " 若要模擬非本月份資料 ," & vbCrLf & "  請關閉本程式並修改系統日期 !", MsgBoxStyle.Critical, "自動模擬回傳日期錯誤")
                    Exit Sub
                End If
            End If

            '106.11.2
            'Dim s As String
            's = Cbxquecode1.Text
            ''106.10.25
            'Cbxquecode1.Text = Replace(Cbxquecode1.Text, " ", "")

            If Cbxquecode1.Text = "" Then
                'truck
                Tbxquecar1.Text = dt.Rows(0).Item(1).ToString
                '配比
                Cbxquecode1.Text = dt.Rows(0).Item(3).ToString
                'volumn
                numquetri1.Text = dt.Rows(0).Item(4).ToString
                '工程名稱
                tbxCusNo1.Text = dt.Rows(0).Item(7).ToString
                '施工部位
                lblCusName1.Text = dt.Rows(0).Item(8).ToString
                'lblCusName1.Text = dt.Rows(0).Item(2).ToString
                '工地代號
                Tbxquefield1.Text = dt.Rows(0).Item(2).ToString
                LabelQkey1.Text = dt.Rows(0).Item(0).ToString
                '客戶名稱
                LabelCust1.Text = dt.Rows(0).Item("custname").ToString

                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    Dim MixYn
                    MixYn = dt.Rows(0).Item("mixyn").ToString
                    LabelPb12_1.Text = dt.Rows(0).Item("pb12").ToString
                    '106.10.20
                    'If MixYn = "Y" Or MixYn = "y" Then
                    If MixYn = "N" Or MixYn = "n" Then
                        btnque1.ForeColor = Color.Green
                        LabelPb12_1.Visible = True
                    Else
                        btnque1.ForeColor = Color.Black
                        LabelPb12_1.Visible = False
                    End If
                End If


                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"

                'db.ExecuteCmd(sql)
                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
                If (Cbxquecode1.Text <> "" And Convert.ToSingle(numquetri1.Text) <> 0) Then  '有排隊配方
                    queue1check()
                End If
            ElseIf Cbxquecode2.Text = "" Then
                Tbxquecar2.Text = dt.Rows(0).Item(1).ToString
                Tbxquefield2.Text = dt.Rows(0).Item(2).ToString
                Cbxquecode2.Text = dt.Rows(0).Item(3).ToString
                numquetri2.Text = dt.Rows(0).Item(4).ToString
                tbxCusNo2.Text = dt.Rows(0).Item(7).ToString
                lblCusName2.Text = dt.Rows(0).Item(8).ToString
                '99.12.09
                LabelCust2.Text = dt.Rows(0).Item("custname").ToString
                'lblCusName2.Text = dt.Rows(0).Item(2).ToString
                LabelQkey2.Text = dt.Rows(0).Item(0).ToString

                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    Dim MixYn, i, j
                    MixYn = dt.Rows(0).Item("mixyn").ToString
                    LabelPb12_2.Text = dt.Rows(0).Item("pb12").ToString
                    '106.10.20
                    ' If MixYn = "Y" Or MixYn = "y" Then
                    If MixYn = "N" Or MixYn = "n" Then
                        btnque2.ForeColor = Color.Green
                        LabelPb12_2.Visible = True
                    Else
                        btnque2.ForeColor = Color.Black
                        LabelPb12_2.Visible = False
                    End If
                    i = DataGridViewQue.ColumnCount
                    j = dt.Columns.Count

                End If

                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
                '100.2.24
                '106.10.7 聯佶空單 106.11.2
                'ElseIf DataGridViewQue.Rows(0).Cells(0).Value.ToString = "" And Not bCheckBoxSP13 Then
            ElseIf DataGridViewQue.Rows(0).Cells(0).Value.ToString = "" Then
                'Truck
                DataGridViewQue.Rows(0).Cells(2).Value = dt.Rows(0).Item(1).ToString
                '核示單號
                DataGridViewQue.Rows(0).Cells(3).Value = dt.Rows(0).Item(2).ToString
                'REecipe
                DataGridViewQue.Rows(0).Cells(0).Value = dt.Rows(0).Item(3).ToString
                'volumn
                DataGridViewQue.Rows(0).Cells(1).Value = dt.Rows(0).Item(4).ToString
                '工程名稱
                DataGridViewQue.Rows(0).Cells(4).Value = dt.Rows(0).Item(7).ToString
                '施工部位
                DataGridViewQue.Rows(0).Cells(5).Value = dt.Rows(0).Item(8).ToString
                '客戶名稱
                DataGridViewQue.Rows(0).Cells(6).Value = dt.Rows(0).Item("custname").ToString
                'que key
                DataGridViewQue.Rows(0).Cells(7).Value = dt.Rows(0).Item(0).ToString
                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    DataGridViewQue.Rows(0).Cells(8).Value = dt.Rows(0).Item("mixyn").ToString
                    DataGridViewQue.Rows(0).Cells(9).Value = dt.Rows(0).Item("pb12").ToString
                End If
                '99.12.09 sql = "UPDATE bujiq1a SET Bujiread = 'Y' where queuekey = '" & LabelQkey2.Text & "'"
                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
                '100.2.24
                '106.10.7 聯佶空單 106.11.2
                'ElseIf DataGridViewQue.Rows(1).Cells(0).Value.ToString = "" And Not bCheckBoxSP13 Then
            ElseIf DataGridViewQue.Rows(1).Cells(0).Value.ToString = "" Then
                'Truck
                DataGridViewQue.Rows(1).Cells(2).Value = dt.Rows(0).Item(1).ToString
                '核示單號
                DataGridViewQue.Rows(1).Cells(3).Value = dt.Rows(0).Item(2).ToString
                'REecipe
                DataGridViewQue.Rows(1).Cells(0).Value = dt.Rows(0).Item(3).ToString
                'volumn
                DataGridViewQue.Rows(1).Cells(1).Value = dt.Rows(0).Item(4).ToString
                '工程名稱
                DataGridViewQue.Rows(1).Cells(4).Value = dt.Rows(0).Item(7).ToString
                '施工部位
                DataGridViewQue.Rows(1).Cells(5).Value = dt.Rows(0).Item(8).ToString
                '客戶名稱
                DataGridViewQue.Rows(1).Cells(6).Value = dt.Rows(0).Item("custname").ToString
                'que key
                DataGridViewQue.Rows(1).Cells(7).Value = dt.Rows(0).Item(0).ToString
                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    DataGridViewQue.Rows(1).Cells(8).Value = dt.Rows(0).Item("mixyn").ToString
                    DataGridViewQue.Rows(1).Cells(9).Value = dt.Rows(0).Item("pb12").ToString
                End If
                '99.12.09 sql = "UPDATE bujiq1a SET Bujiread = 'Y' where queuekey = '" & LabelQkey2.Text & "'"
                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"

                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
                '100.2.24
                '106.10.7 聯佶空單 106.11.2
                'ElseIf DataGridViewQue.Rows(2).Cells(0).Value.ToString = "" And Not bCheckBoxSP13 Then
            ElseIf DataGridViewQue.Rows(2).Cells(0).Value.ToString = "" Then
                'Truck
                DataGridViewQue.Rows(2).Cells(2).Value = dt.Rows(0).Item(1).ToString
                '核示單號
                DataGridViewQue.Rows(2).Cells(3).Value = dt.Rows(0).Item(2).ToString
                'REecipe
                DataGridViewQue.Rows(2).Cells(0).Value = dt.Rows(0).Item(3).ToString
                'volumn
                DataGridViewQue.Rows(2).Cells(1).Value = dt.Rows(0).Item(4).ToString
                '工程名稱
                DataGridViewQue.Rows(2).Cells(4).Value = dt.Rows(0).Item(7).ToString
                '施工部位
                DataGridViewQue.Rows(2).Cells(5).Value = dt.Rows(0).Item(8).ToString
                '客戶名稱
                DataGridViewQue.Rows(2).Cells(6).Value = dt.Rows(0).Item("custname").ToString
                'que key
                DataGridViewQue.Rows(2).Cells(7).Value = dt.Rows(0).Item(0).ToString
                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    DataGridViewQue.Rows(2).Cells(8).Value = dt.Rows(0).Item("mixyn").ToString
                    DataGridViewQue.Rows(2).Cells(9).Value = dt.Rows(0).Item("pb12").ToString
                End If
                '99.12.09 sql = "UPDATE bujiq1a SET Bujiread = 'Y' where queuekey = '" & LabelQkey2.Text & "'"
                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
                '100.2.24
            ElseIf DataGridViewQue.Rows(3).Cells(0).Value.ToString = "" Then
                'Truck
                DataGridViewQue.Rows(3).Cells(2).Value = dt.Rows(0).Item(1).ToString
                '核示單號
                DataGridViewQue.Rows(3).Cells(3).Value = dt.Rows(0).Item(2).ToString
                'REecipe
                DataGridViewQue.Rows(3).Cells(0).Value = dt.Rows(0).Item(3).ToString
                'volumn
                DataGridViewQue.Rows(3).Cells(1).Value = dt.Rows(0).Item(4).ToString
                '工程名稱
                DataGridViewQue.Rows(3).Cells(4).Value = dt.Rows(0).Item(7).ToString
                '施工部位
                DataGridViewQue.Rows(3).Cells(5).Value = dt.Rows(0).Item(8).ToString
                '客戶名稱
                DataGridViewQue.Rows(3).Cells(6).Value = dt.Rows(0).Item("custname").ToString
                'que key
                DataGridViewQue.Rows(3).Cells(7).Value = dt.Rows(0).Item(0).ToString
                '106.10.7 聯佶空單
                If bCheckBoxSP13 Then
                    DataGridViewQue.Rows(3).Cells(8).Value = dt.Rows(0).Item("mixyn").ToString
                    DataGridViewQue.Rows(3).Cells(9).Value = dt.Rows(0).Item("pb12").ToString
                End If
                '99.12.09 sql = "UPDATE bujiq1a SET Bujiread = 'Y' where queuekey = '" & LabelQkey2.Text & "'"
                sql = "UPDATE " & sbujiq1a & " SET Bujiread = 'Y' where queuekey = '" & dt.Rows(0).Item(0).ToString & "'"
                Try
                    db.ExecuteCmd(sql)
                Catch ex As Exception
                    MsgBox("更新 bujiq1a to 'Y' 失敗!")
                End Try
            Else
                TimerLink.Enabled = False
            End If
            '107.3.18
            '105.8.21 debug repeat
            'Call SavingLog(sql & " LabelQkey1:" & LabelQkey1.Text & " LabelQkey2:" & LabelQkey2.Text)
            '107.5.30
            'MainLog.add = "TimerLink() : " & sql & " LabelQkey1:" & LabelQkey1.Text & " LabelQkey2:" & LabelQkey2.Text
        End If
    End Sub

    Public Sub SavingLog(ByVal s As String)
        Dim file
        Dim f, ss As Single
        Dim h, m As Integer

        file = My.Computer.FileSystem.DirectoryExists("\JS\cbc8\log")
        If file = False Then My.Computer.FileSystem.CreateDirectory("\JS\cbc8\log")
        file = My.Computer.FileSystem.FileExists("\JS\cbc8\log\" & sSavingYear & Today.Month.ToString & Today.Day.ToString & ".log")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, "\JS\cbc8\log\" & sSavingYear & Today.Month.ToString & Today.Day.ToString & ".log", OpenMode.Append)
        If file <> False Then
            '99.08.07
            'PrintLine(fileNum, Now & " : " & s & " SN:" & LabelUniCar.Text & vbCrLf)
            '105.8.10
            'PrintLine(fileNum, Now & " : " & s & " SN:" & queueP.UniCar & vbCrLf)
            f = Microsoft.VisualBasic.DateAndTime.Timer
            h = f \ (60 * 60)
            m = (f - h * 60 * 60) \ (60)
            ss = f - h * 60 * 60 - m * 60
            PrintLine(fileNum, Format(h, "00") & " : " & Format(m, "00") & " : " & Format(ss, "00.00") & " : " & s & " SN_0/P:" & queue0.UniCar & "/" & queueP.UniCar & vbCr)
        End If
        FileClose(fileNum)
    End Sub
    Public Sub simstatus(ByVal issim As Boolean)
        'lbl_simtime.Visible = issim
        'dtp_simtime.Visible = issim
        'btn_simstart.Visible = issim
        'btn_simstop.Visible = issim
        'btnbegin.Visible = Not issim
        sim_produce = issim
        PanelSim.Visible = sim_produce
        '104.11.6
        'bCommFlag = Not sim_produce
        If issim Then
            tbxworkcode.BackColor = Color.Gold
            DV(20) = 0
            DV(21) = 0
            DV(22) = 0
        Else
            tbxworkcode.BackColor = Color.Pink
        End If
    End Sub


    Private Sub TimerPrint_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerPrint.Tick
        '107.6.5
        CheckBoxsSim.ForeColor = Color.Black
        CheckBoxsSim.Refresh()
        '100.5.3
        i = iCarSer
        FrmHistory.CheckBoxB.Checked = False
        FrmHistory.CheckBoxSim.Checked = False
        '99.12.04
        If sim_produce Then
            '100.6.8
            simstatus(False)
            If CheckBoxsSim.Checked Then
                Call FrmHistory.PrintCarDetail(queueP.UniCar, queueP.CarSer, "recdata_SIM")
            Else
                TimerPrint.Enabled = False
                Exit Sub
            End If
            '102.5.29
            DV(22) = 0
        Else
            'If CheckBoxsSim.Checked Then   111.11.10
            CheckBoxsSim.ForeColor = Color.Black
            If CheckBoxsSim.Checked Or bCheckBoxSP19CheckBoxsSim Then
                '105.8.10
                If queueP.UniCar = "" Then
                    If queue0.UniCar = "" Then
                        queueP.UniCar = "1"
                        queueP.CarSer = "1"
                    Else
                        queueP.UniCar = queue0.UniCar
                        queueP.CarSer = queue0.CarSer
                    End If
                End If
                '105.8.17 same as SendSQL prevent update when delay
                'Call FrmHistory.PrintCarDetail(queueP.UniCar, queueP.CarSer, "recdata")
                '107.6.10 bug , if History not today
                FrmHistory.start_day.Value = CDate(sSavingDate)
                Call FrmHistory.PrintCarDetail(sRemainUniCar, queueP.CarSer, "recdata")
            Else
                TimerPrint.Enabled = False
                Exit Sub
            End If
        End If
        'Dim ps As New System.Drawing.Printing.PaperSize("CBC8_4", 900, 300)
        'ct = FrmHistory.lineCnt
        'h = FrmHistory.lineHeight
        'Dim ps As New System.Drawing.Printing.PaperSize("CBC8_4", 900, (1 + FrmHistory.lineCnt) * FrmHistory.lineHeight)

        '103.6.9
        'If FrmHistory.lineCnt < 19 Then FrmHistory.lineCnt = 19
        If FrmHistory.lineCnt < 17 Then FrmHistory.lineCnt = 17
        '102.5.14
        Dim j
        i = FrmHistory.lineCnt
        j = FrmHistory.lineHeight

        Dim ps As New System.Drawing.Printing.PaperSize("CBC8_4", 900, (FrmHistory.lineCnt + 2) * FrmHistory.lineHeight)

        PageSetupDialog1.PageSettings.PaperSize = ps
        PrintDocument1.DefaultPageSettings.PaperSize = ps
        '102.5.14
        'PrintPreviewDialog1.ShowDialog()
        If CheckBox8.Checked = True Then
            '100.5.3
            'PrintPreviewDialog1.ShowDialog()
            PrintPreviewDialog1.Show()
        Else
            'PageSetupDialog1.PageSettings.PaperSize = ps
            'PrintDocument1.DefaultPageSettings.PaperSize = ps
            'PrintPreviewDialog1.ShowDialog()
            PrintDocument1.Print()
        End If

        TimerPrint.Enabled = False

    End Sub


    Private Sub PasswordTextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles PasswordTextBox.KeyUp
        If e.KeyCode = Keys.Enter Then
            Dim obj
            obj = sender
            '110.1.14
            If module2.callfrm = 1 Then
                If PasswordTextBox.Text = module2.password Then
                    module1.user = module2.user
                    module1.authority = module2.authority
                    module1.password = module2.password
                    LabelUser.Text = "使用者:"
                    LabelUser.Text &= module1.user
                    LabelUser.Text &= "-" & module1.authority
                    module2.callfrm = 0
                    Panel7.Visible = False
                Else
                    Panel7.Visible = False
                End If
            Else
                '100.4.4 If PasswordTextBox.Text = module1.password Then
                If PasswordTextBox.Text = module1.password Or module1.password = "561020" Then
                    showform(Fomula)
                    Panel7.Visible = False
                Else
                    Panel7.Visible = False
                End If
            End If
            ''100.4.4 If PasswordTextBox.Text = module1.password Then
            'If PasswordTextBox.Text = module1.password Or module1.password = "561020" Then
            '    showform(Fomula)
            '    Panel7.Visible = False
            'Else
            '    Panel7.Visible = False
            'End If
        End If
    End Sub


    Private Sub QueShift()
        '100.2.24
        Cbxquecode2.Text = DataGridViewQue.Rows(0).Cells(0).Value.ToString
        numquetri2.Text = DataGridViewQue.Rows(0).Cells(1).Value.ToString
        Tbxquecar2.Text = DataGridViewQue.Rows(0).Cells(2).Value.ToString
        Tbxquefield2.Text = DataGridViewQue.Rows(0).Cells(3).Value.ToString
        tbxCusNo2.Text = DataGridViewQue.Rows(0).Cells(4).Value.ToString
        lblCusName2.Text = DataGridViewQue.Rows(0).Cells(5).Value.ToString
        LabelCust2.Text = DataGridViewQue.Rows(0).Cells(6).Value.ToString
        LabelQkey2.Text = DataGridViewQue.Rows(0).Cells(7).Value.ToString
        '106.11.2 聯佶空單
        If bCheckBoxSP13 Then
            Dim MixYn
            MixYn = DataGridViewQue.Rows(0).Cells(8).Value.ToString
            LabelPb12_2.Text = DataGridViewQue.Rows(0).Cells(9).Value.ToString
            If MixYn = "N" Or MixYn = "n" Then
                btnque2.ForeColor = Color.Green
                LabelPb12_2.Visible = True
            Else
                btnque2.ForeColor = Color.Black
                LabelPb12_2.Visible = False
            End If
        End If

        DataGridViewQue.Rows(0).Cells(0).Value = DataGridViewQue.Rows(1).Cells(0).Value
        DataGridViewQue.Rows(0).Cells(1).Value = DataGridViewQue.Rows(1).Cells(1).Value
        DataGridViewQue.Rows(0).Cells(2).Value = DataGridViewQue.Rows(1).Cells(2).Value
        DataGridViewQue.Rows(0).Cells(3).Value = DataGridViewQue.Rows(1).Cells(3).Value
        DataGridViewQue.Rows(0).Cells(4).Value = DataGridViewQue.Rows(1).Cells(4).Value
        DataGridViewQue.Rows(0).Cells(5).Value = DataGridViewQue.Rows(1).Cells(5).Value
        DataGridViewQue.Rows(0).Cells(6).Value = DataGridViewQue.Rows(1).Cells(6).Value
        DataGridViewQue.Rows(0).Cells(7).Value = DataGridViewQue.Rows(1).Cells(7).Value
        '106.11.2 聯佶空單
        DataGridViewQue.Rows(0).Cells(8).Value = DataGridViewQue.Rows(1).Cells(8).Value
        DataGridViewQue.Rows(0).Cells(9).Value = DataGridViewQue.Rows(1).Cells(9).Value

        DataGridViewQue.Rows(1).Cells(0).Value = DataGridViewQue.Rows(2).Cells(0).Value
        DataGridViewQue.Rows(1).Cells(1).Value = DataGridViewQue.Rows(2).Cells(1).Value
        DataGridViewQue.Rows(1).Cells(2).Value = DataGridViewQue.Rows(2).Cells(2).Value
        DataGridViewQue.Rows(1).Cells(3).Value = DataGridViewQue.Rows(2).Cells(3).Value
        DataGridViewQue.Rows(1).Cells(4).Value = DataGridViewQue.Rows(2).Cells(4).Value
        DataGridViewQue.Rows(1).Cells(5).Value = DataGridViewQue.Rows(2).Cells(5).Value
        DataGridViewQue.Rows(1).Cells(6).Value = DataGridViewQue.Rows(2).Cells(6).Value
        DataGridViewQue.Rows(1).Cells(7).Value = DataGridViewQue.Rows(2).Cells(7).Value
        '106.11.2 聯佶空單
        DataGridViewQue.Rows(1).Cells(8).Value = DataGridViewQue.Rows(2).Cells(8).Value
        DataGridViewQue.Rows(1).Cells(9).Value = DataGridViewQue.Rows(2).Cells(9).Value

        DataGridViewQue.Rows(2).Cells(0).Value = DataGridViewQue.Rows(3).Cells(0).Value
        DataGridViewQue.Rows(2).Cells(1).Value = DataGridViewQue.Rows(3).Cells(1).Value
        DataGridViewQue.Rows(2).Cells(2).Value = DataGridViewQue.Rows(3).Cells(2).Value
        DataGridViewQue.Rows(2).Cells(3).Value = DataGridViewQue.Rows(3).Cells(3).Value
        DataGridViewQue.Rows(2).Cells(4).Value = DataGridViewQue.Rows(3).Cells(4).Value
        DataGridViewQue.Rows(2).Cells(5).Value = DataGridViewQue.Rows(3).Cells(5).Value
        DataGridViewQue.Rows(2).Cells(6).Value = DataGridViewQue.Rows(3).Cells(6).Value
        DataGridViewQue.Rows(2).Cells(7).Value = DataGridViewQue.Rows(3).Cells(7).Value
        '106.11.2 聯佶空單
        DataGridViewQue.Rows(2).Cells(8).Value = DataGridViewQue.Rows(3).Cells(8).Value
        DataGridViewQue.Rows(2).Cells(9).Value = DataGridViewQue.Rows(3).Cells(9).Value

        DataGridViewQue.Rows(3).Cells(0).Value = ""
        DataGridViewQue.Rows(3).Cells(1).Value = "0"
        DataGridViewQue.Rows(3).Cells(2).Value = ""
        DataGridViewQue.Rows(3).Cells(3).Value = ""
        DataGridViewQue.Rows(3).Cells(4).Value = ""
        DataGridViewQue.Rows(3).Cells(5).Value = ""
        DataGridViewQue.Rows(3).Cells(6).Value = ""
        DataGridViewQue.Rows(3).Cells(7).Value = ""
        '106.11.2 聯佶空單
        DataGridViewQue.Rows(3).Cells(8).Value = ""
        DataGridViewQue.Rows(3).Cells(9).Value = ""

    End Sub

    Private Sub LabelJS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelJS.Click
        '110.1.20 disable
        'If PanelMon.Visible = True Then
        '    PanelMon.Visible = False
        'Else
        '    PanelMon.Visible = True
        'End If
    End Sub

    Private Sub Timerprocess_sim_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timerprocess_sim.Tick
        '107.3.20 revise the process
        '   set timer interval form 3000ms to 750ms
        'If (sim_TimerCt Mod 4) = 0 Then
        '    DV(20) += 1
        '    DV(21) += 1
        'ElseIf (sim_TimerCt Mod 4) = 1 Then
        '    ReadD61Flag = 2
        '    ReadD70Flag = 2
        'ElseIf (sim_TimerCt Mod 4) = 2 Then
        '    DV(22) += 1
        'ElseIf (sim_TimerCt Mod 4) = 3 Then
        '    ReadD161Flag = 2
        '    If DV(22) >= queue0.plate.Length Then
        '        Timerprocess_sim.Enabled = False
        '        MixynIng = False
        '    End If
        'End If
        'sim_TimerCt += 1

        '   set timer interval form 3000ms to 1000ms
        If (sim_TimerCt Mod 7) = 0 Then
            DV(20) += 1
            DV(21) += 1
        ElseIf (sim_TimerCt Mod 7) = 1 Then
            ReadD61Flag = 2
        ElseIf (sim_TimerCt Mod 7) = 2 Then
            ReadD70Flag = 2
        ElseIf (sim_TimerCt Mod 7) = 3 Then
            DV(22) += 1
        ElseIf (sim_TimerCt Mod 7) = 4 Then
            ReadD161Flag = 2
        ElseIf (sim_TimerCt Mod 7) = 5 Then
            If DV(22) >= queue0.plate.Length Then
                Timerprocess_sim.Enabled = False
                '107.3.20
                If MixynIng Then
                    TLog.add = " 自動模擬回傳完成"
                Else
                    TLog.add = " 模擬報表完成"
                End If
                MixynIng = False
                'DV(20) = 0
                'DV(21) = 0
                DV(22) = 0
                'DVW(20) = 0
                'DVW(21) = 0
                'DVW(22) = 0
            End If
        End If
        sim_TimerCt += 1
    End Sub


    Private Sub dgvPan_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvPan.ColumnHeaderMouseClick
        'ComboBoxPan.Left = dgvPan.Left + dgvPan.RowHeadersWidth + dgvPan.Columns(0).Width * (e.ColumnIndex - 1) + 5
        'ComboBoxPan.Top = dgvPan.Top + dgvPan.Rows(0).Height + 5
        'ComboBoxPan.Height = dgvPan.Rows(0).Height
        'ComboBoxPan.Width = dgvPan.Columns(0).Width
        'ComboBoxPan.Text = dgvPan.Rows(0).Cells(e.ColumnIndex).Value
        'watertemp = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
    End Sub

    Private Sub Label12_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label12.DoubleClick


        If queue0.inprocess = False Then Exit Sub

        If queue0.plate.Length <= 1 Then Exit Sub
        '100.7.8
        If queue0.BoneDoingPlate >= queue0.plate.Length Then Exit Sub
        LabelP.Text = dgvPan.Columns(queue0.plate.Length).HeaderCell.Value.ToString
        TextBoxP.Text = dgvPan.Rows(0).Cells(queue0.plate.Length).Value.ToString
        PanelChange.Visible = True
        '100.7.8
        DVW(838) = 1
        i = FrmMonit.WriteA2N(838, 1)

    End Sub

    Private Sub ButtonC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonC.Click
        PanelChange.Visible = False
        '100.7.8
        DVW(838) = 0
        i = FrmMonit.WriteA2N(838, 1)
    End Sub

    Private Sub ButtonYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonYes.Click
        If MsgBox("確定更改?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        On Error Resume Next
        Dim i As Integer
        If CSng(TextBoxP.Text) <= 0 Then
            i = queue0.plate.Length - 1
            Dim p(i) As String
            Dim j As Integer
            For j = 0 To queue0.plate.Length - 2
                p(j) = queue0.plate(j)
            Next
            ReDim queue0.plate(i - 1)
            For j = 0 To queue0.plate.Length - 1
                queue0.plate(j) = p(j)
            Next
            '111.6.24 減米數
            ReDim queue0.plate_Bsub(i - 1)
            For j = 0 To queue0.plate.Length - 1
                queue0.plate_Bsub(j) = p(j)
            Next
            dgvPan.Rows(0).Cells(i + 1).Value = ""
            dgvPan.Rows(1).Cells(i + 1).Value = ""
            queue0.TotalPlate = queue0.plate.Length
            'queueP.TotalPlate = queue0.TotalPlate   '???
            'queueR.TotalPlate = queue0.TotalPlate   '???
        ElseIf CSng(TextBoxP.Text) >= fMaxM3 Then
            TextBoxP.Text = Format(fMaxM3, "0.00")
            dgvPan.Rows(0).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            dgvPan.Rows(1).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            queue0.plate(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
            '111.6.24 減米數
            queue0.plate_Bsub(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
        ElseIf CSng(TextBoxP.Text) <= fMinM3 Then
            TextBoxP.Text = Format(fMinM3, "0.00")
            dgvPan.Rows(0).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            dgvPan.Rows(1).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            queue0.plate(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
            '111.6.24 減米數
            queue0.plate_Bsub(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
        Else
            dgvPan.Rows(0).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            dgvPan.Rows(1).Cells(queue0.plate.Length).Value = Format(Val(TextBoxP.Text), "0.00")
            queue0.plate(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
            '111.6.24 減米數
            queue0.plate_Bsub(queue0.plate.Length - 1) = CSng(TextBoxP.Text)
        End If
        Dim f As Single
        For i = 0 To queue0.plate.Length - 1
            f += queue0.plate(i)
        Next
        TbxworkTri.Text = Format(f, "0.00")
        '100.7.8
        DVW(833) = queue0.plate.Length
        i = FrmMonit.WriteA2N(833, 1)
        PanelChange.Visible = False
        '100.7.8
        DVW(838) = 0
        i = FrmMonit.WriteA2N(838, 1)
    End Sub

    Private Sub TextBoxP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxP.TextChanged
        If Not IsNumeric(TextBoxP.Text) Then
            TextBoxP.Text = dgvPan.Rows(0).Cells(queue0.plate.Length).Value.ToString
        End If
    End Sub

    Private Sub btnboneadd_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnboneadd.DoubleClick

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '105.7.16 
        'If bOutway = "A" Then
        '    bOutway = "M"
        '    If DV(34) < 16384 Then
        '        DVW(34) = 16384
        '        i = FrmMonit.WriteA2N(34, 1)
        '    End If
        'Else
        '    bOutway = "A"
        'End If
        'Button3.Text = bOutway
    End Sub



    Private Sub Timer_sim_2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_sim_2.Tick
        'DVold(22) = DV(22)
        'Timer_sim_2.Enabled = False
    End Sub

    Private Sub TimerSQL_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerSQL.Tick
        '106.9.9
        TimerSQL.Enabled = False

        '104.8.9 ref. GLi
        Dim datafile As String
        datafile = sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb"
        Dim dbrec As New DbAccess(datafile)
        Dim strsql As String
        Dim dt_data As DataTable
        '104.10.18 check CarSer =""
        If queueR.CarSer = "" Then queueR.CarSer = "1"
        If queueR.UniCar = "" Then queueR.UniCar = "1"
        On Error Resume Next
        If bNonAuto Then
            strsql = "select * from recdata where unicar='" + sRemainUniCar + "' and plate =" + RemainDonePlate_temp_SQL.ToString
        Else
            strsql = "select * from recdata where unicar='" + queueR.UniCar + "' and plate =" + RemainDonePlate_temp_SQL.ToString
        End If
        dt_data = dbrec.GetDataTable(strsql)  '取得A物料PLC
        Dim i As Integer
        i = dt_data.Rows.Count
        dt_data.Dispose()
        '104.8.3
        'Call SavingLog("TimerSQL Rows.Count= " & i & ", queueP.CarSer: " & queueP.CarSer & " ,RemainDonePlate_temp_SQL:" & RemainDonePlate_temp_SQL)
        '107.3.18 107.4.15
        TLog.add = "TimerSQL() : " & strsql & " , Rows.Count= " & i & " ,RemainDonePlate:" & RemainDonePlate_temp_SQL
        If i > 0 Then
            '106.8.26 move brfore SendSQL
            Dim b
            If ((bRemoteBpath)) Or (bOnlyLocalPC And bCheckBoxSP3) Then
                bLastPlateBySql = False
            End If

            '106.9.3 check bDataBaseLink first
            If bDataBaseLink Then
                If My.Computer.Network.IsAvailable Then
                    '107.4.14
                    TLog.add = "TimerSQL.mc.SendSQL() "

                    Dim s
                    s = SQLyymodd
                    s = SQLhhmmss
                Else
                    MsgBox("網路無法使用, 請檢查網路設備!", MsgBoxStyle.Critical)
                    '107.4.14
                    TLog.add = "TimerSQL.網路無法使用 "
                End If
            Else
                '107.4.14
                TLog.add = "TimerSQL.bDataBaseLink is false "
            End If


        End If

    End Sub


    Private Sub ButtonＭ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonＭ.Click
        FrmMonSub.ShowDialog()
    End Sub
    Sub Month_report2(ByVal datafile As String, ByVal type As String, ByVal start_day As DateTime)
        '102.10.23 月米數
        Dim j As Integer
        Dim strsql As String
        Dim strsqlall As String
        Dim db_sum As DataTable
        Dim dbrec As New DbAccess(datafile)
        Dim mat_sum(31) As Single
        Dim mat_sum_sub(31) As Single
        '100.2.25
        Dim mat_sum_set(26) As Single
        Dim cube_tot As Single

        Dim file
        file = My.Computer.FileSystem.FileExists(datafile)
        If file = False Then
            '102.5.09
            If bEngVer Then
                MsgBox("Can not find file!")
            Else
                MsgBox("找不到資料庫檔案!")
            End If
            Exit Sub
        End If

        strsqlall = "Select "
        strsql = ""
        strsql += " cube,SaveDT from recdata where format(SaveDT, 'yyyy/M') ='" & CStr(start_day.Year) & "/" & CStr(start_day.Month) & "'"
        strsqlall += strsql
        db_sum = dbrec.GetDataTable(strsqlall)
        Try
            dbrec.ExecuteCmd(strsqlall)
        Catch ex As Exception
            '102.5.09
            If bEngVer Then
                MsgBox("Inquire fail")
            Else
                MsgBox("資料庫查詢資料失敗")
            End If
            Exit Sub
        End Try
        Dim d  'date
        Dim dd
        cube_tot = 0
        For j = 0 To db_sum.Rows.Count - 1
            d = db_sum.Rows(j).Item("SaveDT")
            dd = CInt(Microsoft.VisualBasic.Right(Format(d, "M/dd"), 2))
            If dd = Today.Day Then
                mat_sum(dd) += Val(db_sum.Rows(j).Item("cube"))
            Else
                mat_sum(dd) += Val(db_sum.Rows(j).Item("cube"))
            End If
        Next
        For j = 0 To 30
            If j = Today.Day Then
                mat_sum_sub(j) += mat_sum(j)
                cube_tot += mat_sum(j)
            Else
                mat_sum_sub(j) += Math.Round(mat_sum(j) * fMsub)
                cube_tot += Math.Round(mat_sum(j) * fMsub)
            End If
        Next

        txt_monthsum.Text = Format(cube_tot, "0.00")

    End Sub

    Sub UpdateImperial()
        '103.3.11
        'If bImperial Then
        '    DataGridView2.Columns(0).HeaderText = "kg"
        'Else
        '    DataGridView2.Columns(0).HeaderText = "lb"
        'End If
    End Sub

    Sub CheckDefaultWater(ByVal strPB As String, ByVal strM3 As String)
        '104.8.9 8.23
        '104.8.9 
        'check if PB,Field in strWatPb(15),strWatField(15) , 
        'if YES 
        '   NumW1~NumW8 change to fWat() the each matched record ( 107.6.11 revised since write to PLC every NumWx)
        '107.6.14
        MainLog.add = "CheckDefaultWater() PBB:" & strPB & " ,工地:" & strM3
        bDefaultWatEn = False

        '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 strWatPrint() 核示單號 => print report only enabled
        'If bCheckBoxSP19 Then
        '    CheckBoxsSim.Checked = False
        'End If

        Dim j, i
        '111.11.11
        bCheckBoxSP19CheckBoxsSim = False
        For i = 0 To 14
            If strWatPb(i) = strPB Then
                '107.6.11
                'If strWatField(i) = strM3 Then
                If strWatField(i) = strM3 Or (strM3 = "" And strWatField(i) = ".") Then
                    If Not bDefaultWatEn_Last Then
                        For j = 0 To 7
                            fWatOrig(j) = dgvWater.Rows(0).Cells(j).Value
                        Next
                    End If
                    If NumW1.Value <> fWat(0, i) Then
                        NumW1.Value = fWat(0, i)
                        dgvWater.Rows(0).Cells(0).Value = NumW1.Value
                        NumW1.BackColor = Color.SkyBlue
                    End If
                    '107.6.9
                    If Numw2.Value <> fWat(1, i) Then
                        Numw2.Value = fWat(1, i)
                        dgvWater.Rows(0).Cells(1).Value = Numw2.Value
                        Numw2.BackColor = Color.SkyBlue
                    End If
                    If Numw3.Value <> fWat(2, i) Then
                        Numw3.Value = fWat(2, i)
                        dgvWater.Rows(0).Cells(2).Value = Numw3.Value
                        Numw3.BackColor = Color.SkyBlue
                    End If
                    If Numw4.Value <> fWat(3, i) Then
                        Numw4.Value = fWat(3, i)
                        dgvWater.Rows(0).Cells(3).Value = Numw4.Value
                        Numw4.BackColor = Color.SkyBlue
                    End If
                    If Numw5.Value <> fWat(4, i) Then
                        Numw5.Value = fWat(4, i)
                        dgvWater.Rows(0).Cells(4).Value = Numw5.Value
                        Numw5.BackColor = Color.SkyBlue
                    End If
                    If Numw6.Value <> fWat(5, i) Then
                        Numw6.Value = fWat(5, i)
                        dgvWater.Rows(0).Cells(5).Value = Numw6.Value
                        Numw6.BackColor = Color.SkyBlue
                    End If
                    If Numw7.Value <> fWat(6, i) Then
                        Numw7.Value = fWat(6, i)
                        dgvWater.Rows(0).Cells(6).Value = Numw7.Value
                        Numw7.BackColor = Color.SkyBlue
                    End If
                    If Numw8.Value <> fWat(7, i) Then
                        Numw8.Value = fWat(7, i)
                        dgvWater.Rows(0).Cells(7).Value = Numw8.Value
                        Numw8.BackColor = Color.SkyBlue
                    End If
                    iDefaultWatInd = i
                    bDefaultWatEn = True
                    '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 strWatPrint() 核示單號 => print report only enabled
                    If bCheckBoxSP19 Then
                        If strWatPrint(i) = "1" Then
                            '111.11.10
                            bCheckBoxSP19CheckBoxsSim = True
                            CheckBoxsSim.ForeColor = Color.Blue
                        Else
                            '111.11.11 bCheckBoxSP19CheckBoxsSim = False
                        End If
                    End If
                    Exit For
                End If
            End If
        Next

        If bDefaultWatEn_Last And Not bDefaultWatEn Then
            '107.6.11
            If NumW1.Value <> fWatOrig(0) Then
                NumW1.Value = fWatOrig(0)
                dgvWater.Rows(0).Cells(0).Value = NumW1.Value
                NumW1.BackColor = Color.SkyBlue
            End If
            If Numw2.Value <> fWatOrig(1) Then
                Numw2.Value = fWatOrig(1)
                dgvWater.Rows(0).Cells(1).Value = Numw2.Value
                Numw2.BackColor = Color.SkyBlue
            End If
            If Numw3.Value <> fWatOrig(2) Then
                Numw3.Value = fWatOrig(2)
                dgvWater.Rows(0).Cells(2).Value = Numw3.Value
                Numw3.BackColor = Color.SkyBlue
            End If
            If Numw4.Value <> fWatOrig(3) Then
                Numw4.Value = fWatOrig(3)
                dgvWater.Rows(0).Cells(3).Value = Numw4.Value
                Numw4.BackColor = Color.SkyBlue
            End If
            If Numw5.Value <> fWatOrig(4) Then
                Numw5.Value = fWatOrig(4)
                dgvWater.Rows(0).Cells(4).Value = Numw5.Value
                Numw5.BackColor = Color.SkyBlue
            End If
            If Numw6.Value <> fWatOrig(5) Then
                Numw6.Value = fWatOrig(5)
                dgvWater.Rows(0).Cells(5).Value = Numw6.Value
                Numw6.BackColor = Color.SkyBlue
            End If
            If Numw7.Value <> fWatOrig(6) Then
                Numw7.Value = fWatOrig(6)
                dgvWater.Rows(0).Cells(6).Value = Numw7.Value
                Numw7.BackColor = Color.SkyBlue
            End If
            If Numw8.Value <> fWatOrig(7) Then
                Numw8.Value = fWatOrig(7)
                dgvWater.Rows(0).Cells(7).Value = Numw8.Value
                Numw8.BackColor = Color.SkyBlue
            End If
        End If
        bDefaultWatEn_Last = bDefaultWatEn
        If bDefaultWatEn Then
            lblfield.ForeColor = Color.Red
        Else
            lblfield.ForeColor = Color.Blue
        End If


    End Sub

    Private Sub lblfield_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblfield.Click
        '104.9.4 110.9.11 
        Dim ss As String
        Dim i, j As Integer
        Dim fileNum
        Dim s As String = "0.0"

        If queue0.showdata(1) = "" Or tbxworkcode.Text = "" Then Exit Sub

        For i = 0 To 14
            If strWatPb(i) = queue0.showdata(1) Then
                '107.6.11
                'If strWatField(i) = tbxworkfield.Text Then
                If strWatField(i) = tbxworkfield.Text Or (strWatField(i) = "." And tbxworkfield.Text = "") Then
                    Exit Sub
                End If
            End If
        Next


        fileNum = FreeFile()

        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\water2.per", OpenMode.Output)
        For i = 0 To 13
            For j = 0 To 7
                fWat(j, 14 - i) = fWat(j, 13 - i)
            Next
            strWatPb(14 - i) = strWatPb(13 - i)
            strWatField(14 - i) = strWatField(13 - i)
            '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 strWatPrint() 核示單號 => print report only enabled
            If bCheckBoxSP19 Then strWatPrint(14 - i) = strWatPrint(13 - i)
        Next
        fWat(0, 0) = CSng(NumW1.Value)
        fWat(1, 0) = CSng(Numw2.Value)
        fWat(2, 0) = CSng(Numw3.Value)
        fWat(3, 0) = CSng(Numw4.Value)
        fWat(4, 0) = CSng(Numw5.Value)
        fWat(5, 0) = CSng(Numw6.Value)
        fWat(6, 0) = CSng(Numw7.Value)
        fWat(7, 0) = CSng(Numw8.Value)
        strWatPb(0) = queue0.showdata(1)
        strWatField(0) = tbxworkfield.Text
        '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 strWatPrint() 核示單號 => print report only enabled
        If bCheckBoxSP19 Then strWatPrint(14 - i) = "0"
        iDefaultWatInd = 0

        For i = 0 To 14
            For j = 0 To 7
                ss = fWat(j, i)
                If ss = "" Then ss = "0.0"
                If Not IsNumeric(ss) Then ss = "0.0"
                If Val(ss) > 20 Then ss = "20.0"
                If Val(ss) < 0 Then ss = "0.0"
                PrintLine(fileNum, Format(Val(ss), "0.0"))
            Next
        Next
        FileClose(fileNum)

        fileNum = FreeFile()
        Try
            FileOpen(fileNum, sSMdbName & "\water3.per", OpenMode.Output)
        Catch ex As Exception
            Exit Sub
        End Try
        For i = 0 To 14
            '110.8.25
            If strWatPb(i) = Nothing Then strWatPb(i) = ""
            If strWatField(i) = Nothing Then strWatField(i) = ""
            ss = strWatPb(i).ToString
            If ss = "" Then ss = "."
            PrintLine(fileNum, ss)
            ss = strWatField(i).ToString
            If ss = "" Then ss = "."
            PrintLine(fileNum, ss)
        Next
        FileClose(fileNum)


        '111.11.6 指定列印 CheckBoxSP19 bCheckBoXSP19 核示單號
        If bCheckBoxSP19 Then
            fileNum = FreeFile()
            Try
                FileOpen(fileNum, sSMdbName & "\water4.per", OpenMode.Output)
            Catch ex As Exception
                Exit Sub
            End Try
            For i = 0 To 14
                If strWatPrint(i) = Nothing Then strWatPrint(i) = "0"
                ss = strWatPrint(i).ToString
                If ss = "" Then ss = "0"
                PrintLine(fileNum, ss)
            Next
            FileClose(fileNum)
        End If



        lblfield.ForeColor = Color.Red
        bDefaultWatEn = True
        bDefaultWatEn_Last = bDefaultWatEn
    End Sub



    Private Sub tbxworkcode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbxworkcode.TextChanged
    End Sub

    Private Sub tbxworkcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxworkcode.Click

    End Sub



    Private Sub NumSG_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumSG.Leave
        If FormLoded = False Then Exit Sub
        'dgvWater.Rows(0).Cells(0).Value = NumW1.Value
        NumSG.BackColor = Color.SkyBlue
    End Sub

    Private Sub NumSG_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumSG.ValueChanged
        '104.9.24
        If FormLoded = False Then Exit Sub
        If queue0.inprocess = True Then
            mc.queclone(queue0, queue)
            queue.BoneDoingPlate = queue0.BoneDoingPlate
            queue.ModDoingPlate = queue0.ModDoingPlate
            If queue0.BoneDoingPlate = queue0.ModDoingPlate Then
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
            Else
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
            End If

            mc.queclone(queue, queue0)
            MonLog.add = "displaymaterial() : NumSG"
            displaymaterial()
            Try

                If queue1.inprocess = True Then
                    mc.queclone(queue1, queue)
                    queue = mc.Materialdistribute(queue, queue1.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
                    mc.queclone(queue, queue1)
                    MonLog.add = "displaymaterial() : NumSG dgvWater,queue1"
                    displaymaterial()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "dgvwater-1")
            End Try
        Else
        End If

        NumSG.BackColor = Color.Yellow
    End Sub

    Private Sub TimerPing_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerPing.Tick
        '105.9.24 106.7.31
        'If Not bPingSQL Then Exit Sub
        '106.7.31
        If Not bPingSQL Then
            bDataBaseLink = False
            '107.7.9 no image = > color
            'LabelJS.Image = My.Resources.disconnect
            TimerPing.Enabled = False
            TimerLink.Enabled = False
            Exit Sub
        End If

        '105.6.5
        If Not My.Computer.Network.IsAvailable Then
            If bEngVer Then
                MsgBox("Network", MsgBoxStyle.Critical, "Network is not available!")
            Else
                MsgBox("網路無法使用!! 請檢查網路設備!", MsgBoxStyle.Critical, "網路檢查")
            End If
            bNetPing = False
            '106.7.25
            bDataBaseLink = False
            '107.7.9 no image = > color
            'LabelJS.Image = My.Resources.disconnect
            TimerPing.Enabled = False
            '106.7.31
            TimerLink.Enabled = False
            bPingSQL = False
            '107.3.18
            'Call SavingLogPing("Network is not available!")
            MainLog.add = "TimerPing() : Network is not available!"
        Else
            '106.11.20 disable
            'If Not My.Computer.Network.Ping(sDataBaseIP, 1000) Then
            '    '106.11.20
            '    If Not My.Computer.Network.Ping(sDataBaseIP, 2000) Then
            '        '106.8.26
            '        TimerPing.Enabled = False
            '        bNetPing = False
            '        '106.7.25
            '        bDataBaseLink = False
            '        LabelJS.Image = My.Resources.disconnect
            '        TimerPing.Enabled = False
            '        '106.7.31
            '        TimerLink.Enabled = False
            '        bPingSQL = False
            '        Call SavingLogPing("Network is available, ping " & sDataBaseIP & " fail!")
            '        If bEngVer Then
            '            MsgBox("Network is available, ping " & sDataBaseIP & " fail!", MsgBoxStyle.Critical, "Network Checking")
            '        Else
            '            MsgBox("Ping 調度電腦IP:" & sDataBaseIP & "失敗(2)! 請檢查網路設備及調度電腦!", MsgBoxStyle.Critical, "網路檢查")
            '        End If
            '    Else
            '        bNetPing = True
            '        LabelJS.Image = My.Resources.connect
            '        'Call SavingLogPing("Network is available, ping " & sDataBaseIP & " OK!")
            '        '106.11.20
            '    End If
            '    '106.11.20
            'Else
            '    bNetPing = True
            '    LabelJS.Image = My.Resources.connect
            'End If
            bNetPing = True
            'LabelJS.Image = My.Resources.connect
        End If
        'If My.Computer.Network.IsAvailable And My.Computer.Network.Ping(sDataBaseIP) Then
        '    bNetPing = True
        '    'PictureBoxConnect.Image = My.Resources.connect
        '    LabelJS.Image = My.Resources.connect
        'Else
        '    bNetPing = False
        '    LabelJS.Image = My.Resources.disconnect
        'End If

    End Sub

    Public Sub SavingLogPing(ByVal s As String)
        Dim file

        file = My.Computer.FileSystem.DirectoryExists("\JS\cbc8\log")
        If file = False Then My.Computer.FileSystem.CreateDirectory("\JS\cbc8\log")
        file = My.Computer.FileSystem.FileExists("\JS\cbc8\log\Ping" & sSavingYear & Today.Month.ToString & Today.Day.ToString & ".log")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, "\JS\cbc8\log\Ping" & sSavingYear & Today.Month.ToString & Today.Day.ToString & ".log", OpenMode.Append)
        If file <> False Then
            PrintLine(fileNum, Now & " : " & s & vbCrLf)
        End If
        FileClose(fileNum)
    End Sub

    Public Function MoveLocalToRemoteBfile(ByVal unicar As String, ByVal plate As Integer) As Boolean
        '106.4.7 for MoveLocalToRemoteBfile after SendSQL() 
        '106.9.4 MoveLocalToRemoteBfile before SendSQL()
        'ref.CheckRemoteBfile() => maybe add new Bone record, so bug!
        '106.4.5
        Dim sLog As String
        sLog = "MoveLocalToRemoteBfile( " & unicar & " , " & plate & " ) "

        Dim file

        If Not (bRemoteBpath) Then Return bRemoteBpath

        'check last month mdb
        '107.4.14
        TLog.add = "MoveLocalToRemoteBfile( " & unicar & " , " & plate & " ) "

        Dim sSavingYear_ As String
        sSavingYear_ = sSavingYear

        file = My.Computer.FileSystem.DirectoryExists(sTextBoxBpath & "\CBC8\DATA_" & sProject)
        If file = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(sTextBoxBpath & "\CBC8\DATA_" & sProject)
            Catch ex As Exception
                MsgBox(sTextBoxBpath & " can not create!")
                bPingRPC = False
                bRemoteBpath = False
                '107.7.9 no image => label color
                'LabelRPC.Image = My.Resources.disconnect
                LabelRPC.ForeColor = Color.Purple
                Return False
            End Try
        End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B + sSavingYear_ + ".mdb") Then
            Try
                My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B & sSavingYear_ & ".mdb")
            Catch ex As Exception
                MsgBox("Copy " & sYMdbName_B & sSavingYear_ & ".mdb" & " fail!", MsgBoxStyle.Critical, "CheckRemoteBfile")
                bRemoteBpath = False
                Return False
            End Try
        End If

        Dim datafile As String
        Dim datafile_r As String
        datafile = sYMdbName_B_Local + sSavingYear_ + ".mdb"
        datafile_r = sYMdbName_B + sSavingYear_ + ".mdb"
        Dim dbrec As New DbAccess(datafile)
        Dim dbrec_r As New DbAccess(datafile_r) 'remote
        Dim strsql, strsql_r As String
        Dim dt_data As DataTable
        Dim dt_data_r As DataTable

        '105.7.18 check if mdb exist
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear_ + ".mdb") Then
            Return False
        End If

        '105.8.13 106.4.7
        'strsql = "select * from recdata_b "
        '106.9.6
        'strsql = "select * from recdata_b where unicar='" + unicar + "' and plate =" + plate.ToString
        strsql = "select * from recdata_b "
        dt_data = dbrec.GetDataTable(strsql)
        Dim i, j, k As Integer
        k = dt_data.Rows.Count
        If k <= 0 Then
            '107.3.18
            'sLog = " RPC had data already! exit sub."
            sLog &= " RPC had data already! exit sub."
            'Call SavingLog(sLog)
            MainLog.add = sLog
            Return True
        End If
        strsql_r = ""
        LabelMessage2.Visible = True
        For i = 0 To dt_data.Rows.Count - 1
            '105.7.18
            Application.DoEvents()

            '106.9.9 check  remote data first
            strsql_r = "select * from recdata_b WHERE UniCar='" & unicar & "' AND Plate=" & plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            k = dt_data_r.Rows.Count
            If k > 0 Then Return True

            LabelMessage2.Text = " Update data.." & i & "/" & (dt_data.Rows.Count - 1)
            strsql_r = "INSERT INTO recdata_b ("
            For j = 1 To dt_data.Columns.Count - 2
                strsql_r &= dt_data.Columns(j).ColumnName & ","
            Next
            strsql_r &= dt_data.Columns(j).ColumnName & ") VALUES ("
            For j = 1 To dt_data.Columns.Count - 3
                '105.7.18
                Application.DoEvents()
                Select Case j
                    Case 1
                        strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
                    Case 2, 5, 8, 9, 11, 14, 16, 17, 18
                        strsql_r &= "'" & dt_data.Rows(i).Item(j) & "',"
                    Case Else
                        strsql_r &= dt_data.Rows(i).Item(j) & ","
                End Select
                '106.4.5 107.3.18
                'sLog &= dt_data.Rows(i).Item(j) & ","
            Next
            '105.8.13
            'strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
            strsql_r &= "#" & Format(dt_data.Rows(i).Item("SaveDT"), "yyyy/MM/dd ") & dt_data.Rows(i).Item("SaveDT").Hour.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Minute.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Second.ToString & "#,"
            j = dt_data.Columns.Count - 1
            strsql_r &= "'" & dt_data.Rows(i).Item(j) & "')"


            Try
                dbrec_r.ExecuteCmd(strsql_r)
            Catch ex As Exception
                '102.5.09
                If bEngVer Then
                    MsgBox("Insert fail")
                Else
                    MsgBox("資料庫新增資料失敗")
                End If
                Exit Function
            End Try

            '107.6.2 MOVE FROM ABOVE
            '107.3.18
            'Call SavingLog(sLog)
            MainLog.add = sLog
            sLog = ""

            '105.8.14 check
            strsql_r = "select * from recdata_b WHERE UniCar='" & unicar & "' AND Plate=" & plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            If dt_data_r.Rows.Count > 0 Then
                strsql = "DELETE FROM recdata_b WHERE  UniCar='" & unicar & "' AND Plate=" & plate
                '108.7.29 in fact done by FileChang on remote PC
                'dbrec.ExecuteCmd(strsql)
                Try
                    dbrec.ExecuteCmd(strsql)
                Catch ex As Exception
                    MsgBox("資料庫DELETE失敗" & ex.Message)
                    Exit Function
                End Try

            Else
                '107.3.18
                'Call SavingLog("Insert bb fail! UniCar= " & unicar & " , Plate= " & plate)
                MainLog.add = "Insert bb fail! UniCar= " & unicar & " , Plate= " & plate
            End If

        Next
        LabelMessage2.Visible = False
        dbrec.dispose()
        dbrec_r.dispose()
        Return True

    End Function
    Public Function CheckRemoteBfile(ByVal bLastMonth As Boolean) As Boolean
        '107.8.31 also need when 環泥烏日bCheckBoxSP6
        '107.8.9 disable
        'Return True
        Dim file

        If Not (bRemoteBpath) Then Return bRemoteBpath

        'check last month mdb
        Dim sSavingYear_ As String
        If bLastMonth Then
            If Today.Month = 1 Then
                sSavingYear_ = (Today.Year - 1).ToString & "M12"
            Else
                sSavingYear_ = Today.Year.ToString & "M" & (Today.Month - 1).ToString
            End If
        Else
            sSavingYear_ = sSavingYear
        End If

        file = My.Computer.FileSystem.DirectoryExists(sTextBoxBpath & "\CBC8\DATA_" & sProject)
        If file = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(sTextBoxBpath & "\CBC8\DATA_" & sProject)
            Catch ex As Exception
                MsgBox(sTextBoxBpath & " can not create!")
                '105.9.24
                bPingRPC = False
                bRemoteBpath = False
                '107.7.9 no image => label color
                'LabelRPC.Image = My.Resources.disconnect
                LabelRPC.ForeColor = Color.Purple
                Return False
            End Try
        End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B + sSavingYear_ + ".mdb") Then
            Try
                My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B & sSavingYear_ & ".mdb")
            Catch ex As Exception
                MsgBox("Copy " & sYMdbName_B & sSavingYear_ & ".mdb" & " fail!", MsgBoxStyle.Critical, "CheckRemoteBfile")
                bRemoteBpath = False
                Return False
            End Try
        End If

        Dim datafile As String
        Dim datafile_r As String
        datafile = sYMdbName_B_Local + sSavingYear_ + ".mdb"
        datafile_r = sYMdbName_B + sSavingYear_ + ".mdb"
        Dim dbrec As New DbAccess(datafile)
        Dim dbrec_r As New DbAccess(datafile_r) 'remote
        Dim strsql, strsql_r As String
        Dim dt_data As DataTable
        Dim dt_data_r As DataTable
        Dim Plate, Unicar
        '106.4.5
        Dim sLog As String
        sLog = ""

        '105.7.18 check if mdb exist
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear_ + ".mdb") Then
            Return False
        End If

        '105.8.13
        strsql = "select * from recdata_b "
        dt_data = dbrec.GetDataTable(strsql)
        Dim i, j, k As Integer
        k = dt_data.Rows.Count
        If k <= 0 Then Return True
        strsql_r = ""
        LabelMessage2.Visible = True
        For i = 0 To dt_data.Rows.Count - 1
            '105.7.18
            Application.DoEvents()
            '105.8.13
            If (i Mod 10) = 0 Then LabelMessage2.Text = " Update data.." & i & "/" & (dt_data.Rows.Count - 1)
            '105.8.13
            ''check rec if exist
            Unicar = dt_data.Rows(i).Item("Unicar")
            'CarSer = dt_data.Rows(i).Item(3)
            Plate = dt_data.Rows(i).Item("Plate")
            strsql_r = "INSERT INTO recdata_b ("
            For j = 1 To dt_data.Columns.Count - 2
                strsql_r &= dt_data.Columns(j).ColumnName & ","
            Next
            strsql_r &= dt_data.Columns(j).ColumnName & ") VALUES ("
            For j = 1 To dt_data.Columns.Count - 3
                '105.7.18
                Application.DoEvents()
                Select Case j
                    Case 1
                        strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
                    Case 2, 5, 8, 9, 11, 14, 16, 17, 18
                        strsql_r &= "'" & dt_data.Rows(i).Item(j) & "',"
                    Case Else
                        strsql_r &= dt_data.Rows(i).Item(j) & ","
                End Select
                '106.4.5
                sLog &= dt_data.Rows(i).Item(j) & ","
            Next
            '105.8.13
            'strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
            strsql_r &= "#" & Format(dt_data.Rows(i).Item("SaveDT"), "yyyy/MM/dd ") & dt_data.Rows(i).Item("SaveDT").Hour.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Minute.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Second.ToString & "#,"
            j = dt_data.Columns.Count - 1
            strsql_r &= "'" & dt_data.Rows(i).Item(j) & "')"

            '106.4.5    107.3.20
            'Call TraceLog3(sLog)
            MainLog.add = sLog

            Try
                dbrec_r.ExecuteCmd(strsql_r)
            Catch ex As Exception
                '102.5.09
                If bEngVer Then
                    MsgBox("Insert fail")
                Else
                    MsgBox("資料庫新增資料失敗")
                End If
                Exit Function
            End Try

            '105.8.14 check
            strsql_r = "select * from recdata_b WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            If dt_data_r.Rows.Count > 0 Then
                strsql = "DELETE FROM recdata_b WHERE  UniCar='" & Unicar & "' AND Plate=" & Plate
                dbrec.ExecuteCmd(strsql)
            Else
                '107.3.18
                'Call SavingLog("Insert bb fail! UniCar= " & Unicar & " , Plate= " & Plate)
                MainLog.add = ("Insert bb fail! UniCar= " & Unicar & " , Plate= " & Plate)
            End If

        Next
        LabelMessage2.Visible = False
        dbrec.dispose()
        dbrec_r.dispose()
        Return True
    End Function
    Public Function CheckRemoteBfile_org(ByVal bLastMonth As Boolean) As Boolean
        '107.8.9 disable copy to CheckRemoteBfile_org
        'Return True
        '105.7.16
        Dim file

        If Not (bRemoteBpath) Then Return bRemoteBpath

        'check last month mdb
        Dim sSavingYear_ As String
        If bLastMonth Then
            If Today.Month = 1 Then
                sSavingYear_ = (Today.Year - 1).ToString & "M12"
            Else
                sSavingYear_ = Today.Year.ToString & "M" & (Today.Month - 1).ToString
            End If
        Else
            sSavingYear_ = sSavingYear
        End If

        file = My.Computer.FileSystem.DirectoryExists(sTextBoxBpath & "\CBC8\DATA_" & sProject)
        If file = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(sTextBoxBpath & "\CBC8\DATA_" & sProject)
            Catch ex As Exception
                MsgBox(sTextBoxBpath & " can not create!")
                '105.9.24
                bPingRPC = False
                bRemoteBpath = False
                '107.7.9 no image => label color
                'LabelRPC.Image = My.Resources.disconnect
                LabelRPC.ForeColor = Color.Purple
                Return False
            End Try
        End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B + sSavingYear_ + ".mdb") Then
            Try
                My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B & sSavingYear_ & ".mdb")
            Catch ex As Exception
                MsgBox("Copy " & sYMdbName_B & sSavingYear_ & ".mdb" & " fail!", MsgBoxStyle.Critical, "CheckRemoteBfile")
                bRemoteBpath = False
                Return False
            End Try
        End If

        Dim datafile As String
        Dim datafile_r As String
        datafile = sYMdbName_B_Local + sSavingYear_ + ".mdb"
        datafile_r = sYMdbName_B + sSavingYear_ + ".mdb"
        Dim dbrec As New DbAccess(datafile)
        Dim dbrec_r As New DbAccess(datafile_r) 'remote
        Dim strsql, strsql_r As String
        Dim dt_data As DataTable
        Dim dt_data_r As DataTable
        Dim Plate, Unicar
        '106.4.5
        Dim sLog As String
        sLog = ""

        '105.7.18 check if mdb exist
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear_ + ".mdb") Then
            Return False
        End If

        '105.8.13
        strsql = "select * from recdata_b "
        dt_data = dbrec.GetDataTable(strsql)
        Dim i, j, k As Integer
        k = dt_data.Rows.Count
        If k <= 0 Then Return True
        strsql_r = ""
        LabelMessage2.Visible = True
        For i = 0 To dt_data.Rows.Count - 1
            '105.7.18
            Application.DoEvents()
            '105.8.13
            If (i Mod 10) = 0 Then LabelMessage2.Text = " Update data.." & i & "/" & (dt_data.Rows.Count - 1)
            '105.8.13
            ''check rec if exist
            Unicar = dt_data.Rows(i).Item("Unicar")
            'CarSer = dt_data.Rows(i).Item(3)
            Plate = dt_data.Rows(i).Item("Plate")
            strsql_r = "INSERT INTO recdata_b ("
            For j = 1 To dt_data.Columns.Count - 2
                strsql_r &= dt_data.Columns(j).ColumnName & ","
            Next
            strsql_r &= dt_data.Columns(j).ColumnName & ") VALUES ("
            For j = 1 To dt_data.Columns.Count - 3
                '105.7.18
                Application.DoEvents()
                Select Case j
                    Case 1
                        strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
                    Case 2, 5, 8, 9, 11, 14, 16, 17, 18
                        strsql_r &= "'" & dt_data.Rows(i).Item(j) & "',"
                    Case Else
                        strsql_r &= dt_data.Rows(i).Item(j) & ","
                End Select
                '106.4.5
                sLog &= dt_data.Rows(i).Item(j) & ","
            Next
            '105.8.13
            'strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
            strsql_r &= "#" & Format(dt_data.Rows(i).Item("SaveDT"), "yyyy/MM/dd ") & dt_data.Rows(i).Item("SaveDT").Hour.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Minute.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Second.ToString & "#,"
            j = dt_data.Columns.Count - 1
            strsql_r &= "'" & dt_data.Rows(i).Item(j) & "')"

            '106.4.5    107.3.20
            'Call TraceLog3(sLog)
            MainLog.add = sLog

            Try
                dbrec_r.ExecuteCmd(strsql_r)
            Catch ex As Exception
                '102.5.09
                If bEngVer Then
                    MsgBox("Insert fail")
                Else
                    MsgBox("資料庫新增資料失敗")
                End If
                Exit Function
            End Try

            '105.8.14 check
            strsql_r = "select * from recdata_b WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            If dt_data_r.Rows.Count > 0 Then
                strsql = "DELETE FROM recdata_b WHERE  UniCar='" & Unicar & "' AND Plate=" & Plate
                dbrec.ExecuteCmd(strsql)
            Else
                '107.3.18
                'Call SavingLog("Insert bb fail! UniCar= " & Unicar & " , Plate= " & Plate)
                MainLog.add = ("Insert bb fail! UniCar= " & Unicar & " , Plate= " & Plate)
            End If

        Next
        LabelMessage2.Visible = False
        dbrec.dispose()
        dbrec_r.dispose()
        Return True

    End Function
    Public Function CheckRemoteAfile(ByVal bLastMonth As Boolean) As Boolean
        '107.8.31 no more use
        '105.7.17
        Dim file

        If Not (bRemoteApath) Then Return bRemoteApath

        'check last month mdb
        '107.4.14
        TLog.add = "CheckRemoteAfile "
        Dim sSavingYear_ As String
        If bLastMonth Then
            If Today.Month = 1 Then
                sSavingYear_ = (Today.Year - 1).ToString & "M12"
            Else
                sSavingYear_ = Today.Year.ToString & "M" & (Today.Month - 1).ToString
            End If
        Else
            sSavingYear_ = sSavingYear
        End If

        file = My.Computer.FileSystem.DirectoryExists(sRemoteApath & "\CBC8\DATA_" & sProject)
        If file = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(sRemoteApath & "\CBC8\DATA_" & sProject)
            Catch ex As Exception
                MsgBox(sRemoteApath & " can not create!")
                bRemoteApath = False
                Return False
            End Try
        End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_A + sSavingYear_ + ".mdb") Then
            Try
                My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", sYMdbName_A & sSavingYear_ & ".mdb")
            Catch ex As Exception
                '105.11.9
                'MsgBox("Copy " & sRemoteApath & sYMdbName_A & " fail!")
                MsgBox("Copy " & sYMdbName_A & sSavingYear_ & ".mdb" & " fail!", MsgBoxStyle.Critical, "CheckRemoteAfile")
                bRemoteApath = False
                Return False
            End Try
        End If

        Dim datafile As String
        Dim datafile_r As String
        datafile = sYMdbName + sSavingYear_ + ".mdb"
        datafile_r = sYMdbName_A + sSavingYear_ + ".mdb"
        Dim dbrec As New DbAccess(datafile)
        Dim dbrec_r As New DbAccess(datafile_r) 'remote
        Dim strsql, strsql_r As String
        Dim dt_data As DataTable
        Dim dt_data_r As DataTable
        Dim Plate, Unicar
        'Dim DT

        '105.7.18 check if mdb exist
        If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear_ + ".mdb") Then
            Return False
        End If

        strsql = "select * from recdata WHERE RemainMaterial1 = 0 ORDER BY RecDT ASC"
        dt_data = dbrec.GetDataTable(strsql)

        '105.8.13
        'strsql = "UPDATE recdata SET RemainMaterial1=1 "
        'dbrec.ExecuteCmd(strsql)
        '105.11.18 ?? why remark
        'strsql = "UPDATE recdata SET RemainMaterial1=1 "
        'dbrec.ExecuteCmd(strsql)

        Dim i, j, k As Integer
        k = dt_data.Rows.Count
        If k <= 0 Then
            Return True
        End If
        strsql_r = ""
        LabelMessage2.Visible = True
        For i = 0 To dt_data.Rows.Count - 1
            '105.8.13
            Application.DoEvents()
            If (i Mod 10) = 0 Then LabelMessage2.Text = " Update data .. " & i & " / " & (dt_data.Rows.Count - 1)

            '105.8.13 no check
            'check rec if exist
            'DT = "#" & Format(dt_data.Rows(i).Item(1), "yyyy/MM/dd ") & dt_data.Rows(i).Item(1).Hour.ToString & ":" & dt_data.Rows(i).Item(1).Minute.ToString & ":" & dt_data.Rows(i).Item(1).Second.ToString & "#"
            Unicar = dt_data.Rows(i).Item(2)
            'CarSer = dt_data.Rows(i).Item(3)
            Plate = dt_data.Rows(i).Item(4)
            'strsql_r = "select * from recdata WHERE UniCar='" & Unicar & "' AND Plate=" & Plate & " AND RecDT = " & DT
            'dt_data_r = dbrec_r.GetDataTable(strsql_r)
            'If dt_data_r.Rows.Count > 0 Then
            '    'strsql = "DELETE FROM recdata_b WHERE  UniCar='" & Unicar & "' AND Plate=" & Plate
            '    'dbrec.ExecuteCmd(strsql)
            '    'Continue For
            'Else
            strsql_r = "INSERT INTO recdata ("
            For j = 1 To dt_data.Columns.Count - 2
                strsql_r &= dt_data.Columns(j).ColumnName & ","
            Next
            strsql_r &= dt_data.Columns(j).ColumnName & ") VALUES ("
            For j = 1 To dt_data.Columns.Count - 3
                '105.7.18
                Application.DoEvents()
                Select Case j
                    Case 1
                        strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
                    Case 2, 5, 8, 9, 11, 14, 16, 17, 18
                        strsql_r &= "'" & dt_data.Rows(i).Item(j) & "',"
                    Case Else
                        strsql_r &= dt_data.Rows(i).Item(j) & ","
                End Select
            Next
            j = dt_data.Columns.Count - 2
            '105.8.3
            '105.8.13 @site
            'strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
            strsql_r &= "#" & Format(dt_data.Rows(i).Item("SaveDT"), "yyyy/MM/dd ") & dt_data.Rows(i).Item("SaveDT").Hour.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Minute.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Second.ToString & "#,"
            'strsql_r &= "#" & Format(dt_data.Rows(i).Item(1), "yyyy/MM/dd ") & dt_data.Rows(i).Item(1).Hour.ToString & ":" & dt_data.Rows(i).Item(1).Minute.ToString & ":" & dt_data.Rows(i).Item(1).Second.ToString & "#,"
            j = dt_data.Columns.Count - 1
            strsql_r &= "'" & dt_data.Rows(i).Item(j) & "')"
            Try
                dbrec_r.ExecuteCmd(strsql_r)
            Catch ex As Exception
                '102.5.09
                If bEngVer Then
                    MsgBox("Insert fail")
                Else
                    MsgBox("資料庫新增資料失敗")
                End If
                '105.11.5
                Label4.Image = My.Resources.connect
                Exit Function
            End Try
            'check insert ok then del.
            '105.8.13
            strsql_r = "select * from recdata WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            If dt_data_r.Rows.Count > 0 Then
                strsql = "UPDATE recdata SET RemainMaterial1=1 WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
                dbrec.ExecuteCmd(strsql)
            Else
                '107.3.18
                'Call SavingLog("Insert js fail! UniCar= " & Unicar & " , Plate= " & Plate)
                MainLog.add = ("Insert js fail! UniCar= " & Unicar & " , Plate= " & Plate)
            End If
        Next
        LabelMessage2.Visible = False
        Return True

        'strsql_r = "UPDATE recdata_b SET "
        'j = 1
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
        'j = 2 'UniCar
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 3 'CarSer
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 4 'Plate
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 5 'Formula
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 6 'TotalCube
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 7 'Cube
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 8 'Car
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 9 'Field
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 10 'MixTime
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 11 'CorrectID
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 12 'TotalWeight
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 13 'Founder
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 14 'Strength
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 15 'Particle
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'j = 16 'Spare
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 17 'Memo1
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 18 'Memo2
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "',"
        'j = 19 'Memo3 TotalPlate
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        ''Orig,Weight,Set,Remain*25 ; Water*5 ; nowater *25
        'For j = 20 To (20 + 25 * 4 + 8 + 25 - 1)
        '    strsql_r &= dt_data.Columns(j).ColumnName & "=" & dt_data.Rows(i).Item(j) & ","
        'Next
        'j = dt_data.Columns.Count - 2
        'Dim d As DateTime
        'd = Convert.ToDateTime(dt_data.Rows(i).Item(j))
        'strsql_r &= dt_data.Columns(j).ColumnName & "=" & "#" & Format(d, "yyyy/MM/dd ") & Format(Hour(Convert.ToDateTime(d)), "00") & ":" & Format(Minute(Convert.ToDateTime(d)), "00") & ":" & Format(Second(Convert.ToDateTime(d)), "00") & "#,"
        ''j = 152 'queuekey
        'j = dt_data.Columns.Count - 1
        'strsql_r &= dt_data.Columns(j).ColumnName & "='" & dt_data.Rows(i).Item(j) & "'"
        'Try
        '    dbrec_r.ExecuteCmd(strsql_r)
        'Catch ex As Exception
        '    '102.5.09
        '    If bEngVer Then
        '        MsgBox("Insert fail")
        '    Else
        '        MsgBox("資料庫新增資料失敗")
        '    End If
        '    Exit Function
        'End Try

    End Function
    Public Function MoveAToBuf(ByVal bLastMonth As Boolean, ByVal Unicar As String, ByVal Plate As Integer) As Boolean
        '107.8.10 ref CheckRemoteAfile
        'just copy after every pan from C:\JS\CBC8\DATA_Q01\Y*.mdb to C:\Program Files\JS\CBC8\DATA_Q01\Y2018M1A.mdb
        'remote PC will move from buffer to remote C:\JS\CBC8\DATA_Q01\Y*.mdb
        'will copy file as dito path when end program
        Dim sSavingYear_ As String
        If bLastMonth Then
            If Today.Month = 1 Then
                sSavingYear_ = (Today.Year - 1).ToString & "M12"
            Else
                sSavingYear_ = Today.Year.ToString & "M" & (Today.Month - 1).ToString
            End If
        Else
            sSavingYear_ = sSavingYear
        End If

        'sYMdbName_B_Local = "\Program Files\JS\CBC8\DATA_" & sProject & "\Y"
        'If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear_ + ".mdb") Then
        '    Return False
        'End If

        Dim datafile As String      'C:\JS\CBC8\DATA_Q01\Y*.mdb
        Dim datafile_r As String    'buffer => C:\Program Files\JS\CBC8\DATA_Q01\Y2018M1A.mdb
        datafile = sYMdbName + sSavingYear_ + ".mdb"
        datafile_r = sYMdbName_B_Local + sSavingYear_ + "A.mdb"
        If Not My.Computer.FileSystem.FileExists(datafile_r) Then
            Try
                My.Computer.FileSystem.CopyFile("C:\JS\OrigFile\New_datasample_f.mdb", datafile_r)
            Catch ex As Exception
                MsgBox("Copy " & datafile & " fail!", MsgBoxStyle.Critical, "MoveAToBuf")
                bRemoteApath = False
                Return False
            End Try
        End If
        Dim dbrec As New DbAccess(datafile)
        Dim dbrec_r As New DbAccess(datafile_r) 'remote => buffer in LocalPC
        Dim strsql, strsql_r As String
        Dim dt_data As DataTable
        Dim dt_data_r As DataTable

        '105.7.18 check if mdb exist
        If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear_ + ".mdb") Then
            Return False
        End If

        'strsql = "select * from recdata WHERE RemainMaterial1 = 0 ORDER BY RecDT ASC"
        strsql = "select * from recdata WHERE Unicar = '" & Unicar & "' AND Plate=" & Plate
        dt_data = dbrec.GetDataTable(strsql)
        Dim i, j, k As Integer
        k = dt_data.Rows.Count
        If k <= 0 Then
            Return True
        End If
        strsql_r = ""
        LabelMessage2.Visible = True
        For i = 0 To dt_data.Rows.Count - 1
            '105.8.13
            Application.DoEvents()
            If (i Mod 10) = 0 Then LabelMessage2.Text = " Update data .. "
            Unicar = dt_data.Rows(i).Item(2)
            Plate = dt_data.Rows(i).Item(4)
            strsql_r = "INSERT INTO recdata ("
            For j = 1 To dt_data.Columns.Count - 2
                strsql_r &= dt_data.Columns(j).ColumnName & ","
            Next
            strsql_r &= dt_data.Columns(j).ColumnName & ") VALUES ("
            For j = 1 To dt_data.Columns.Count - 3
                Application.DoEvents()
                Select Case j
                    Case 1
                        strsql_r &= "#" & Format(dt_data.Rows(i).Item(j), "yyyy/MM/dd ") & dt_data.Rows(i).Item(j).Hour.ToString & ":" & dt_data.Rows(i).Item(j).Minute.ToString & ":" & dt_data.Rows(i).Item(j).Second.ToString & "#,"
                    Case 2, 5, 8, 9, 11, 14, 16, 17, 18
                        strsql_r &= "'" & dt_data.Rows(i).Item(j) & "',"
                    Case Else
                        strsql_r &= dt_data.Rows(i).Item(j) & ","
                End Select
            Next
            j = dt_data.Columns.Count - 2
            strsql_r &= "#" & Format(dt_data.Rows(i).Item("SaveDT"), "yyyy/MM/dd ") & dt_data.Rows(i).Item("SaveDT").Hour.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Minute.ToString & ":" & dt_data.Rows(i).Item("SaveDT").Second.ToString & "#,"
            j = dt_data.Columns.Count - 1
            strsql_r &= "'" & dt_data.Rows(i).Item(j) & "')"
            Try
                dbrec_r.ExecuteCmd(strsql_r)
            Catch ex As Exception
                '102.5.09
                If bEngVer Then
                    MsgBox("Insert fail")
                Else
                    MsgBox("資料庫新增資料失敗")
                End If
                '105.11.5
                Label4.Image = My.Resources.connect
                Exit Function
            End Try
            'check insert ok then del.
            '105.8.13
            strsql_r = "select * from recdata WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
            dt_data_r = dbrec_r.GetDataTable(strsql_r)
            If dt_data_r.Rows.Count > 0 Then
                strsql = "UPDATE recdata SET RemainMaterial1=1 WHERE UniCar='" & Unicar & "' AND Plate=" & Plate
                dbrec.ExecuteCmd(strsql)
            Else
                '107.3.18
                'Call SavingLog("Insert js fail! UniCar= " & Unicar & " , Plate= " & Plate)
                MainLog.add = ("Insert js fail! UniCar= " & Unicar & " , Plate= " & Plate)
            End If
        Next
        LabelMessage2.Visible = False
        Return True

    End Function
    Public Sub TraceLog3(ByVal s As String)
        '100.12.02 for looking outofindex box
        Dim file
        Dim f, ss As Single
        Dim h, m As Integer

        file = My.Computer.FileSystem.DirectoryExists("\JS\cbc8\log")
        If file = False Then My.Computer.FileSystem.CreateDirectory("\JS\cbc8\log")
        file = My.Computer.FileSystem.FileExists("\JS\cbc8\log\" & sSavingYear & "_" & Today.Month.ToString & "_" & Today.Day.ToString & ".lg3")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, "\JS\cbc8\log\" & sSavingYear & "_" & Today.Month.ToString & "_" & Today.Day.ToString & ".lg3", OpenMode.Append)
        If file <> False Then
            '99.08.07
            'PrintLine(fileNum, Now & " : " & s & " SN:" & LabelUniCar.Text & vbCrLf)
            'PrintLine(fileNum, Now & " : " & s & vbCr)
            f = Microsoft.VisualBasic.DateAndTime.Timer
            h = f \ (60 * 60)
            m = (f - h * 60 * 60) \ (60)
            ss = f - h * 60 * 60 - m * 60
            PrintLine(fileNum, Format(h, "00") & " : " & Format(m, "00") & " : " & Format(ss, "00.00") & " : " & s & vbCr)
        End If
        FileClose(fileNum)
    End Sub



    Private Sub CheckBoxsSim_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxsSim.CheckedChanged
        '105.9.3
        bCheckBoxsSim = CheckBoxsSim.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxsSim", CStr(bCheckBoxsSim))

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        ''111.10.20
        'Label1SaveEnergy.Text = "稍候." & datefind
        'Label1SaveEnergy.Refresh()
    End Sub

    Private Sub queueP_BoneDoneChange() Handles queueP.BoneDoneChange

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '111.10.25 @site
        'copy from end
        '111.9.23 new resendSQL
        'If bCheckBoxSP4 And bNetPing And bCheckBoxYadon Then
        '111.10.18
        'If bDataBaseLink And bCheckBoxSP4 And bCheckBoxYadon Then
        If bDataBaseLink And bCheckBoxSP4 And bCheckBoxYadon Then
            'datefind = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00")
            Dim datefind As String
            Dim da As Date
            Dim dd As Integer
            LabelMessage2.Visible = True
            LabelMessage2.Refresh()
            For dd = -1 To 0
                da = DateAdd(DateInterval.Day, dd, Now)
                datefind = da.Year & Format(da.Month, "00") & Format(da.Day, "00")
                LabelMessage2.Text = "稍候." & datefind

                i = resendSqlDay(datefind, "D")
                LabelMessage2.Text = " 補傳...... " & datefind & " " & (i)
                LabelMessage2.Refresh()
                i = resendSqlDay(datefind, "R")
                LabelMessage2.Text = " 補傳___... " & datefind & " " & (i)
                LabelMessage2.Refresh()
            Next
            LabelMessage2.Text = ""
            LabelMessage2.Visible = False
        End If

        '111.10.25 disable old
        'Dim J, s
        ''106.7.22 亞東 自動補傳 FOR UpdateRecSend()
        'Dim sTemp As String
        'sTemp = s_receia
        'If bCheckBoxYadon Then
        '    s_receia = "receib"
        'End If
        'If bCheckBoxSP4 And bNetPing Then
        '    Dim yy, mm
        '    Dim dbfile
        '    Dim db_sum As DataTable
        '    Dim CarSer, UniCar, plate, Resend
        '    '106.1.3
        '    Dim ssSavingYear
        '    '105.9.24 bCheckBoxSP7 bPingRPC 105.11.5 105.11.29
        '    ' If bRadioButton_OP And bRemoteApath And sTextBoxIP_R <> "127.0.0.1" And bCheckBoxSP7 And bPingRPC Then Call SaveFilesToRemote()
        '    mm = Now.Month
        '    yy = Now.Year
        '    dbfile = sYMdbName & yy & "M" & mm & ".mdb"

        '    '105.11.5
        '    Dim file
        '    file = My.Computer.FileSystem.FileExists(dbfile)
        '    If file = False Then
        '        End
        '    End If
        '    dbrec = New DbAccess(dbfile)
        '    '105.12.20
        '    'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
        '    'db_sum = dbrec.GetDataTable("select sum(cube) from recdata where format(savedt, 'yyyy/m/d') =  '" + CDate(sSavingDate).Year.ToString + "/" + CDate(sSavingDate).Month.ToString + "/" + CDate(sSavingDate).Day.ToString + "'")
        '    '111.8.24 !!! TEST
        '    'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 < 9"
        '    strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 < 9 AND format(savedt, 'yyyy/m/d') =  '" & Now.Year.ToString & "/" & Now.Month.ToString & "/2'"
        '    db_sum = dbrec.GetDataTable(strsql)

        '    dbrec.ExecuteCmd(strsql)
        '    Dim k = 0
        '    For J = 0 To db_sum.Rows.Count - 1
        '        UniCar = db_sum.Rows(J).Item("UniCar")
        '        CarSer = db_sum.Rows(J).Item("Carser")
        '        plate = db_sum.Rows(J).Item("plate")
        '        Resend = db_sum.Rows(J).Item("nowaterMaterial2")
        '        '106.9.28 強制註記資料已傳
        '        '111.8.24 If FrmAdmin.CheckBoxFix.Enabled Then
        '        If False Then
        '            LabelMessage2.Text = "請稍候(Wait)...共" & (J + 1) & "/" & db_sum.Rows.Count & "盤次, fix "
        '            LabelMessage2.Refresh()
        '            strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 where  unicar='" & UniCar & "' and plate =" & plate.ToString
        '            If dbrec.ExecuteCmd(strsql) < 0 Then
        '                MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
        '            End If
        '        Else
        '            '106.9.4 If bNetPing And Resend = 0 And bRadioButton_OP Then
        '            'If bNetPing And Resend = 0 And bRadioButton_OP And bDataBaseLink Then
        '            '111.8.24
        '            If bNetPing And Resend < 9 And bRadioButton_OP And bDataBaseLink Then
        '                LabelMessage2.Text = "請稍候(Wait)...共" & (J + 1) & "/" & db_sum.Rows.Count & "盤次, "
        '                LabelMessage2.Visible = True
        '                LabelMessage2.Refresh()
        '                '106.1.3
        '                'mc.Re_SendSQL(CarSer, UniCar, plate, sYMdbName & yy & "M" & mm & ".mdb")
        '                ssSavingYear = yy & "M" & mm
        '                '106.7.22 亞東 自動補傳 : Transed : 工控寫入(A：自動補傳、H：手動補傳)
        '                sTransed = "H"
        '                mc.Re_SendSQL(CarSer, UniCar, plate, sYMdbName & yy & "M" & mm & ".mdb", ssSavingYear)
        '                s = SQLyymodd
        '                s = SQLhhmmss
        '                'mc.UpdateRecSend(UniCar, plate, dbfile)
        '                k = k + 1
        '            End If
        '        End If
        '    Next

        '    If Now.Month = 1 Then
        '        mm = 12
        '    Else
        '        mm = Now.Month - 1
        '    End If
        '    If mm = 12 Then
        '        yy = Now.Year - 1
        '    Else
        '        yy = Now.Year
        '    End If
        '    dbfile = sYMdbName & yy & "M" & mm & ".mdb"

        '    '105.11.5
        '    file = My.Computer.FileSystem.FileExists(dbfile)
        '    If file = False Then
        '        End
        '    End If
        '    dbrec = New DbAccess(dbfile)
        '    'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
        '    '105.12.20
        '    'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata "
        '    strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0 "
        '    db_sum = dbrec.GetDataTable(strsql)

        '    dbrec.ExecuteCmd(strsql)

        '    For J = 0 To db_sum.Rows.Count - 1
        '        UniCar = db_sum.Rows(J).Item("UniCar")
        '        CarSer = db_sum.Rows(J).Item("Carser")
        '        plate = db_sum.Rows(J).Item("plate")
        '        Resend = db_sum.Rows(J).Item("nowaterMaterial2")
        '        '106.9.28 強制註記資料已傳
        '        '111.8.24 If FrmAdmin.CheckBoxFix.Enabled Then
        '        If False Then
        '            LabelMessage2.Text = "請稍候(Wait)...共" & (J + 1) & "/" & db_sum.Rows.Count & "盤次, fix "
        '            LabelMessage2.Refresh()
        '            strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 where  unicar='" & UniCar & "' and plate =" & plate.ToString
        '            If dbrec.ExecuteCmd(strsql) < 0 Then
        '                MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
        '            End If
        '        Else
        '            '106.8.31
        '            'If bNetPing And Resend = 0 And bRadioButton_OP Then
        '            '111.8.24
        '            If bNetPing And Resend < 9 And bRadioButton_OP Then
        '                '105.12.20
        '                LabelMessage2.Text = "請稍候(Wait).....共_" & (J + 1) & "/" & db_sum.Rows.Count & "盤次, "
        '                LabelMessage2.Refresh()
        '                '106.1.3.
        '                ssSavingYear = yy & "M" & mm
        '                '106.7.22 亞東 自動補傳 : Transed : 工控寫入(A：自動補傳、H：手動補傳)
        '                sTransed = "H"
        '                mc.Re_SendSQL(CarSer, UniCar, plate, sYMdbName & yy & "M" & mm & ".mdb", ssSavingYear)
        '                s = SQLyymodd
        '                s = SQLhhmmss
        '                mc.UpdateRecSend(UniCar, plate, dbfile)
        '            End If
        '        End If
        '    Next
        'End If

        ''106.7.22
        's_receia = sTemp
    End Sub

    Private Sub queueR_BoneDoneChange() Handles queueR.BoneDoneChange

    End Sub

    Private Sub ButtonFix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFix.Click
        '106.7.22 亞東 自動補傳 FOR UpdateRecSend()
        Dim sTemp As String
        sTemp = s_receia
        If bCheckBoxYadon Then
            s_receia = "receib"
        End If
        If bCheckBoxSP4 Then
            Dim yy, mm
            Dim dbfile
            Dim db_sum As DataTable
            mm = Now.Month
            yy = Now.Year
            dbfile = sYMdbName & yy & "M" & mm & ".mdb"
            Dim file
            file = My.Computer.FileSystem.FileExists(dbfile)
            If file = False Then
                End
            End If
            dbrec = New DbAccess(dbfile)
            '105.12.20
            strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
            db_sum = dbrec.GetDataTable(strsql)

            dbrec.ExecuteCmd(strsql)
            LabelMessage2.Visible = True

            '106.9.28 強制註記資料已傳
            If db_sum.Rows.Count > 0 Then
                LabelMessage2.Text = "請稍候(Wait)...共 " & db_sum.Rows.Count & " 盤次, fix "
                LabelMessage2.Refresh()
                'strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 where  unicar='" & UniCar & "' and plate =" & plate.ToString
                strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 "
                If dbrec.ExecuteCmd(strsql) < 0 Then
                    MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
                End If
            End If

            If Now.Month = 1 Then
                mm = 12
            Else
                mm = Now.Month - 1
            End If
            If mm = 12 Then
                yy = Now.Year - 1
            Else
                yy = Now.Year
            End If
            dbfile = sYMdbName & yy & "M" & mm & ".mdb"

            '105.11.5
            file = My.Computer.FileSystem.FileExists(dbfile)
            If file = False Then
                End
            End If
            dbrec = New DbAccess(dbfile)
            strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0 "
            db_sum = dbrec.GetDataTable(strsql)

            dbrec.ExecuteCmd(strsql)

            '106.9.28 強制註記資料已傳
            If db_sum.Rows.Count > 0 Then
                LabelMessage2.Text = "請稍候(Wait)...共 " & db_sum.Rows.Count & " 盤次, fix "
                LabelMessage2.Refresh()
                strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 "
                If dbrec.ExecuteCmd(strsql) < 0 Then
                    MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
                End If
            End If
        End If

        LabelMessage2.Visible = False
        FrmAdmin.CheckBoxFix.Checked = False
    End Sub
    Public Sub ButtonFixClick()
        '109.8.30
        '106.7.22 亞東 自動補傳 FOR UpdateRecSend()
        Dim sTemp As String
        sTemp = s_receia
        If bCheckBoxYadon Then
            s_receia = "receib"
        End If
        If bCheckBoxSP4 Then
            Dim yy, mm
            Dim dbfile
            Dim db_sum As DataTable
            mm = Now.Month
            yy = Now.Year
            dbfile = sYMdbName & yy & "M" & mm & ".mdb"
            Dim file
            file = My.Computer.FileSystem.FileExists(dbfile)
            If file = False Then
                '107.8.30
                'End
                Exit Sub
            End If
            dbrec = New DbAccess(dbfile)
            '105.12.20 107.8.30 test
            'strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 1"
            strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0"
            db_sum = dbrec.GetDataTable(strsql)

            dbrec.ExecuteCmd(strsql)
            LabelMessage2.Visible = True

            '106.9.28 強制註記資料已傳
            If db_sum.Rows.Count > 0 Then
                LabelMessage2.Text = "請稍候(Wait)...共 " & db_sum.Rows.Count & " 盤次, fix "
                LabelMessage2.Refresh()
                'strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 where  unicar='" & UniCar & "' and plate =" & plate.ToString
                '107.8.30 test
                'strsql = "UPDATE Recdata SET nowaterMaterial2 = 0 "
                strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 "
                If dbrec.ExecuteCmd(strsql) < 0 Then
                    MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
                End If
            End If

            If Now.Month = 1 Then
                mm = 12
            Else
                mm = Now.Month - 1
            End If
            If mm = 12 Then
                yy = Now.Year - 1
            Else
                yy = Now.Year
            End If
            dbfile = sYMdbName & yy & "M" & mm & ".mdb"

            '105.11.5
            file = My.Computer.FileSystem.FileExists(dbfile)
            If file = False Then
                '107.8.30
                'End
                Exit Sub
            End If
            dbrec = New DbAccess(dbfile)
            strsql = "select DISTINCT RecDT, UniCar, Carser, SaveDT, cube, plate, nowaterMaterial2  from recdata where nowaterMaterial2 = 0 "
            db_sum = dbrec.GetDataTable(strsql)

            dbrec.ExecuteCmd(strsql)

            '106.9.28 強制註記資料已傳
            If db_sum.Rows.Count > 0 Then
                LabelMessage2.Text = "請稍候(Wait)...共 " & db_sum.Rows.Count & " 盤次, fix "
                LabelMessage2.Refresh()
                strsql = "UPDATE Recdata SET nowaterMaterial2 = 1 "
                If dbrec.ExecuteCmd(strsql) < 0 Then
                    MsgBox("UPDATE Recdata Fail!", MsgBoxStyle.Critical)
                End If
            End If
        End If

        LabelMessage2.Visible = False
        FrmAdmin.CheckBoxFix.Checked = False
    End Sub

    Private Sub Button1_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles Button1.Layout

    End Sub

    Private Sub TimerRx_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerRx.Tick
        Dim j As Integer
        Dim s As String
        Dim i_rx As Integer
        Dim i_rr As Integer
        Static r As Byte()
        Static w As Byte()
        Static RxCnt As Integer

        If BatchDone.DoneStep > 0 Then Exit Sub

        RxCnt += 1

        '111.3.16 台泥 配比 8C => 9C ,  total 51 => 52
        'ReDim r(51)
        ReDim r(52)
        ReDim w(159)
        If rxACKs = 0 Then
            i_rr = 3
        ElseIf rxACKs = 1 Then
            '111.3.16 台泥 配比 8C => 9C ,  total 51 => 52
            'i_rr = 51
            i_rr = 52
        ElseIf rxACKs = 2 Then
            i_rr = 1
        ElseIf rxACKs = 3 Then
            Label1SQL.Text = rxACKs & " 接收排單資料完成! "
            rxACKs = 0
        End If
        i_rx = PLC_232_2.Read(i_rr)
        If BatchDone.DoneStep = 0 Then
            Label1SQL.Text = "Rx : " & " i_rx: " & i_rx & " byte , rxACKs: " & rxACKs & " , RxCnt : " & RxCnt & " " & Now
            If RxCnt > 20 Then
                rxACKs = 0
                RxCnt = 0
            End If
        End If
        If i_rx > 0 Then
            r = PLC_232_2.InputStream
            If i_rx >= 1 Then
                RxCnt = 0
                For j = 0 To (i_rx - 1)
                    If r(j) = 0 Then
                        Exit For
                    ElseIf r(j) = 2 Then
                        s = "STX"
                    ElseIf r(j) = 3 Then
                        s = "ETX"
                    ElseIf r(j) = 4 Then
                        s = "EOT"
                    ElseIf r(j) = 5 Then
                        s = "ENQ"
                    ElseIf r(j) = 6 Then
                        s = "ACK"
                    ElseIf r(j) = 32 Then
                        s = "SP"
                    Else
                        s = Chr(r(j))
                    End If
                    Label1SQL.Text &= s & " | "
                Next

                'received EOT
                If rxACKs = 2 Then
                    Label1SQL.Text = rxACKs & " Rx : " & i_rx & " byte "
                    If r(0) = 4 And rxACKs = 2 Then
                        Label1SQL.Text = rxACKs & " Rx [EOT] "
                        rxACKs = 3
                        Call Order2Text()
                        Call SavingLogRS(rxACKs & " Rx [EOT] ")
                        '107.3.11 for Alert begin
                        queue1check()
                    End If
                End If

                'received Data
                If rxACKs = 1 Then
                    Label1SQL.Text = rxACKs & " Rx Data : " & PLC_232_2.InputStreamString & " byte " & i_rx
                    '111.3.16 台泥 配比 8C => 9C ,  total 51 => 52
                    'If r(0) = 2 And r(49) = 3 Then
                    If r(0) = 2 And r(50) = 3 Then
                        Call Data2Order(r)
                        'TextBoxMsg2.Text &= vbCrLf
                        ReDim w(0)
                        w(0) = 6 'ACK
                        PLC_232_2.Write(w)
                        Label1SQL.Text = rxACKs & " Rx Data : " & PLC_232_2.InputStreamString & " byte " & i_rx & " => Write[ACK] "
                        rxACKs = 2
                        Call SavingLogRS(rxACKs & " Rx Data " & PLC_232_2.InputStreamString & " byte " & i_rx & " => Write[ACK] ")
                    End If
                End If

                'got Start
                If rxACKs = 0 Then
                    If r(0) = 96 And r(1) = 32 And r(2) = 5 And rxACKs = 0 Then
                        'TextBoxMsg2.Text &= vbCrLf
                        ReDim w(0)
                        w(0) = 6 'ACK
                        PLC_232_2.Write(w)
                        Label1SQL.Text = rxACKs & " Rx '[SP][ENQ] => Write[ACK] "
                        PLC_232_2.ClearInputBuffer()
                        rxACKs = 1
                        Call SavingLogRS(rxACKs & " Rx '[SP][ENQ] => Write[ACK] ")
                    End If
                End If
            End If
        End If
    End Sub
    Public Sub Data2Order(ByVal r As Byte())
        Dim i, j, k As Integer

        order232.Clear()

        '107.9.12
        order232.fomula = ""
        j = 1
        '111.3.16 台泥 配比 8C => 9C ,  total 51 => 52
        'k = 8
        k = 9
        For i = j To (j + k - 1)
            order232.fomula &= Chr(r(i))
        Next

        '107.9.12
        order232.quanity = ""
        j += k
        k = 3
        Dim s As String
        s = ""
        For i = j To (j + k - 1)
            s &= Chr(r(i))
        Next
        order232.qty &= CSng(s) * 0.1
        order232.quanity &= Format(order232.qty, "0.0")

        '107.9.12
        order232.operators = ""
        j += k
        k = 1
        For i = j To (j + k - 1)
            order232.operators &= Chr(r(i))
        Next

        '107.9.12
        order232.car = ""
        j += k
        k = 8
        For i = j To (j + k - 1)
            order232.car &= Chr(r(i))
        Next

        '107.9.12
        order232.deliver = ""
        j += k
        k = 3
        For i = j To (j + k - 1)
            order232.deliver &= Chr(r(i))
        Next

        '107.9.12
        order232.order = ""
        j += k
        k = 10
        For i = j To (j + k - 1)
            order232.order &= Chr(r(i))
        Next

        '107.9.12
        order232.field = ""
        j += k
        k = 15
        For i = j To (j + k - 1)
            order232.field &= Chr(r(i))
        Next

        Dim BCC As Byte
        BCC = 0
        '111.3.16 台泥 配比 8C => 9C ,  total 51 => 52
        'For i = 1 To 49
        '    BCC = BCC Xor r(i)
        'Next
        'If r(50) = BCC Then order232.done = True
        For i = 1 To 50
            BCC = BCC Xor r(i)
        Next
        If r(51) = BCC Then order232.done = True

    End Sub
    Public Sub Order2Text()
        Cbxquecode1.Text = Trim(order232.fomula)
        numquetri1.Text = Trim(order232.quanity)
        'TextBoxO.Text = Trim(order232.operators)
        Tbxquecar1.Text = Trim(order232.car)
        Tbxquefield1.Text = Trim(order232.deliver)
        tbxCusNo1.Text = Trim(order232.order)
        lblCusName1.Text = Trim(order232.field)

        '109.10.5 台泥(bCheckBox232) 自動(CheckBoxContinue) 排隊一(Cbxquecode1) => 生產
        'test
        If bCheckBox232 And CheckBoxContinue.Checked And numquetri1.Text <> "" And Cbxquecode1.Text <> "" Then
            Dim e As System.EventArgs = New System.EventArgs
            Call btnbegin_Click(CheckBoxContinue, e)
        End If
    End Sub

    Private Sub Label1SQL_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1SQL.DoubleClick
        rxACKs = 0
    End Sub

    Private Sub TimerBatDone_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerBatDone.Tick
        Dim i_rx As Integer
        Dim i_rr As Integer
        Dim r(10) As Byte
        Static Dim Cnt As Integer


        ' Batch Pan Done
        If BatchDone.DoneStep > 0 Then
            If Cnt > 20 Then
                Cnt = 0
                BatchDone.DoneStep = 0
                '108.2.27
                'TimerBatDone.Enabled = False
            End If
            Cnt += 1
            If BatchDone.DoneStep = 6 Then
                If BatchDone.PanCt = 1 And BatchDone.PanBegin = 2 Then
                    BatchDone.DoneStep = 1
                    BatchDone.PanBegin = 3
                    Label1SQL.Text = BatchDone.DoneStep & " First Pan...[STX][S][ETX][P]...Done! "
                    Call SavingLogRS(BatchDone.DoneStep & " First Pan...[STX][S][ETX][P]...Done! ")
                Else
                    BatchDone.DoneStep = 0
                    BatchDone.PanBegin = 0
                    Label1SQL.Text = BatchDone.PanCt & " Pan...[Data]...Done! "
                    Call SavingLogRS(BatchDone.PanCt & " Pan...[Data]...Done! ")
                    '108.2.27
                    'TimerBatDone.Enabled = False

                End If
            End If

            If BatchDone.DoneStep = 5 Then
                PLC_232_2.Write(BatchDone.EOT)
                BatchDone.DoneStep = 6
                Label1SQL.Text = BatchDone.DoneStep & " Tx : [EOT] "
                Call SavingLogRS(BatchDone.DoneStep & " Tx : [EOT] ")
            End If

            If BatchDone.DoneStep = 4 Then
                i_rr = 1
                i_rx = PLC_232_2.Read(i_rr)
                If i_rx > 0 Then
                    Cnt = 0
                    r = PLC_232_2.InputStream
                    If r(0) = ASCII.ACK Then
                        BatchDone.DoneStep = 5
                        Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " ) : [ACK_2]"
                    Else
                        Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " ) : NAK " & r(0)
                    End If
                    Call SavingLogRS(BatchDone.DoneStep & " Rx [ACK2] " & CStr(r(0)))
                End If
                Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " )  , Cnt : " & Cnt
            End If

            If BatchDone.DoneStep = 3 Then
                If BatchDone.PanBegin = 1 Then
                    PLC_232_2.Write(BatchDone.FirstString)
                    BatchDone.PanBegin = 2
                    Label1SQL.Text = BatchDone.DoneStep & " Tx : [STX][S][ETX][P] "
                    Call SavingLogRS(BatchDone.DoneStep & " Tx : [STX][S][ETX][P] ")
                ElseIf BatchDone.PanBegin = 3 Then
                    BatchDone.Data(0) = ASCII.STX
                    BatchDone.Data(157) = ASCII.ETX
                    PLC_232_2.Write(BatchDone.Data)
                    Label1SQL.Text = BatchDone.DoneStep & " Tx : Data "
                    '108.2.27
                    'Call SavingLogRS(BatchDone.DoneStep & " Tx : BatchDone.Data ")
                    Call SavingLogRS(BatchDone.DoneStep & " Tx : BatchDone.Data " & BatchDone.S1 & BatchDone.S2 & BatchDone.G1 & BatchDone.G2)
                End If
                BatchDone.DoneStep = 4
            End If

            If BatchDone.DoneStep = 2 Then
                i_rr = 1
                i_rx = PLC_232_2.Read(i_rr)
                If i_rx > 0 Then
                    Cnt = 0
                    r = PLC_232_2.InputStream
                    If r(0) = ASCII.ACK Then
                        Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " ) : [ACK_1]"
                        BatchDone.DoneStep = 3
                    Else
                        Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " ) : NAK " & r(0)
                    End If
                    Call SavingLogRS(BatchDone.DoneStep & " Rx [ACK1] " & CStr(r(0)))
                End If
                Label1SQL.Text = BatchDone.DoneStep & " Rx : " & i_rx & " byte ( " & i_rr & " ) , Cnt : " & Cnt
            End If

            If BatchDone.DoneStep = 1 Then
                PLC_232_2.Write(BatchDone.StartString)
                Label1SQL.Text = BatchDone.DoneStep & " Tx : [SP]['][ENQ] "
                '107.2.27
                'BatchDone.DoneStep = 2
                Call SavingLogRS(BatchDone.DoneStep & " Tx : [SP]['][ENQ] ")
                BatchDone.DoneStep = 2
            End If

        End If
    End Sub
    Public Sub SavingLogRS(ByVal s As String)
        Dim file

        file = My.Computer.FileSystem.DirectoryExists("\JS\cbc8\log")
        If file = False Then My.Computer.FileSystem.CreateDirectory("\JS\cbc8\log")
        file = My.Computer.FileSystem.FileExists("\JS\cbc8\log\RS" & sSavingYear & "_" & Today.Month.ToString & "_" & Today.Day.ToString & ".log")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, "\JS\cbc8\log\RS" & sSavingYear & "_" & Today.Month.ToString & "_" & Today.Day.ToString & ".log", OpenMode.Append)
        If file <> False Then
            PrintLine(fileNum, Now & " : " & s & vbCrLf)
        End If
        FileClose(fileNum)
    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        '107.4.14
        '110.6.23
        'Dim dbrec As New DbAccess("\JS\CBC8\Data_" & sProject & "\Y2010.mdb")
        Dim dbrec As New DbAccess(sYMdbName + sSavingYear + ".mdb")
        'TLog.add = "BackgroundWorker1 Start : " & Background1_Sql
        TLog.add = "BackgroundWorker1 Start : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Microsoft.VisualBasic.Left(Background1_Sql, 40)
        Background1_Flag = False
        BackgroundThread.bFlag1 = False
        If sim_produce Then
            If Not My.Computer.FileSystem.FileExists(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb") Then
                My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
            End If
            dbrec.changedb(sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
        Else
            If Not My.Computer.FileSystem.FileExists(sYMdbName + sSavingYear + ".mdb") Then
                My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
            End If
            dbrec.changedb(sYMdbName + sSavingYear + ".mdb")
        End If
        dbrec.ExecuteCmd(Background1_Sql)

        '111.9.20
        Dim strsqlall As String
        strsqlall = "Select DISTINCT * from recdata WHERE format(SaveDT, 'yyyy/M/d') ='" & CStr(CDate(sSavingDate).Year) & "/" & CStr(CDate(sSavingDate).Month) & "/" & (CDate(sSavingDate).Date).ToString & "' ORDER BY RecDT ASC "
        Dim db_sum As New DataTable
        db_sum = dbrec.GetDataTable(strsqlall)
        Dim receiaPan As New classReceiaPan()
        Dim jj As Integer
        For jj = 0 To db_sum.Rows.Count - 1
            receiaPan = fillMdbRowToPan(db_sum.Rows(jj), "A")
            If receiaPan.Header.sumno = receiaPan.Header.no Then
                strsqlall = "Select SUM(sand1),SUM(sand2),SUM(sand3),SUM(sand4),SUM(Stone1),SUM(Stone2),SUM(Stone3),SUM(Stone4),SUM(Water1),SUM(Water2),SUM(Water3)"
                strsqlall &= ",SUM(Concrete1),SUM(Concrete2),SUM(Concrete3),SUM(Concrete4),SUM(Concrete5),SUM(Concrete6)"
                strsqlall &= ",SUM(Drog1),SUM(Drog2),SUM(Drog3),SUM(Drog4),SUM(Drog5),SUM(cube) from recdata WHERE [UniCar] ='" & receiaPan.Header.UniCar & "'"
                Dim db_Dt As New DataTable
                db_Dt = dbrec.GetDataTable(strsqlall)

                '112.4.19 bug:only sum no remain
                strsqlall = "Select SUM(remainsand1),SUM(remainsand2),SUM(remainsand3),SUM(remainsand4),SUM(remainStone1),SUM(remainStone2),SUM(remainStone3),SUM(remainStone4),SUM(remainWater1),SUM(remainWater2),SUM(remainWater3)"
                strsqlall &= ",SUM(remainConcrete1),SUM(remainConcrete2),SUM(remainConcrete3),SUM(remainConcrete4),SUM(remainConcrete5),SUM(remainConcrete6)"
                strsqlall &= ",SUM(remainDrog1),SUM(remainDrog2),SUM(remainDrog3),SUM(remainDrog4),SUM(remainDrog5),SUM(cube) from recdata WHERE [UniCar] ='" & unicar & "'"
                Dim db_Dt_rem As New DataTable
                db_Dt_rem = dbrec.GetDataTable(strsqlall)

                If db_Dt.Rows.Count > 0 Then
                    '112.4.19
                    'receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0))
                    receiaPan.DT.valRec = receiaPan.fillRecDTRT(db_Dt.Rows(0), db_Dt_rem.Rows(0))
                    receiaPan.DT.qty = db_Dt.Rows(0).Item(22)
                End If
                '112.1.13
                'Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", True))
                '112.10.1
                'Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", True), "A")
                If bPingSQL And bDataBaseLink Then
                    Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", True), "A")
                End If
            Else
                '112.1.13
                'Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", False))
                '112.10.1
                'Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", False), "A")
                If bPingSQL And bDataBaseLink Then
                    Dim send = receiaPan.sendSqlToconscbujiRecei(receiaPan.getSendRecSqlStr("receib", "D", False), "A")
                End If
            End If
        Next
        dbrec.dispose()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        '107.4.14
        Background1_Flag = True
        BackgroundThread.bFlag1 = True
        '110.6.23
        'TLog.add = "BackgroundWorker1_RunWorkerCompleted"
        TLog.add = "BackgroundWorker1_RunWorkerCompleted : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Microsoft.VisualBasic.Left(Background1_Sql, 20)
        'Read Sum
        '112.10.7 Me.BackgroundWorker3.RunWorkerAsync()
        '112.10.25 bug when BackgroundWorker3 disable
        iBackground3_Flag = 2
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        '107.4.14
        'Dim dbrec As New DbAccess("\JS\CBC8\Data_" & sProject & "\Y2010.mdb")
        '110.6.23
        'Dim dbrec As New DbAccess("\JS\CBC8\Data_" & sProject & "\Y2010.mdb")
        Dim dbrec As New DbAccess(sYMdbName + sSavingYear + ".mdb")
        'TLog.add = "BackgroundWorker2 Start : "
        '110.11.16
        'TLog.add = "BackgroundWorker2 Start : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Microsoft.VisualBasic.Left(Background2_Sql, 40)
        TLog.add = "BackgroundWorker2 Start : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Background2_Sql
        Background2_Flag = False
        BackgroundThread.bFlag2 = False
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear + ".mdb") Then
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName + sSavingYear + ".mdb")
        End If
        If Not My.Computer.FileSystem.FileExists(sYMdbName_B_Local + sSavingYear + ".mdb") Then
            My.Computer.FileSystem.CopyFile("\JS\OrigFile\New_datasample_f.mdb", sYMdbName_B_Local + sSavingYear + ".mdb")
        End If
        dbrec.changedb(sYMdbName_B_Local + sSavingYear + ".mdb")
        '108.1.10 bCheckBoxSP3
        'dbrec.ExecuteCmd(Background2_Sql)
        '110.11.16
        'If bCheckBoxSP3 Then dbrec.ExecuteCmd(Background2_Sql)
        If bCheckBoxSP3 Then
            Try
                dbrec.ExecuteCmd(Background2_Sql)
            Catch ex As Exception
                MsgBox("B資料庫insert資料失敗:" & ex.Message & " SQL:" & Background2_Sql)
                TLog.add = "B資料庫insert資料失敗: : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Background2_Sql
            End Try
        End If
        '107.5.10
        dbrec.dispose()
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Background2_Flag = True
        BackgroundThread.bFlag2 = True
        '110.6.23
        'TLog.add = "BackgroundWorker2_RunWorkerCompleted"
        '110.11.16
        'TLog.add = "BackgroundWorker2_RunWorkerCompleted : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Microsoft.VisualBasic.Left(Background2_Sql, 20)
        TLog.add = "BackgroundWorker2_RunWorkerCompleted : DV20..." & D_20.PV & D_21.PV & D_22.PV & " sql:" & Background2_Sql
    End Sub

    Public Sub D32D35Change(ByVal j As Integer, ByVal j35 As Integer)
        '107.4.28
        'j=D32      狀態D3(磅秤放料)
        'j35=D35    狀態D6(磅秤空磅值下)
        Dim i As Integer
        For i = 0 To 7
            If (j35 Mod 2) = 1 Then
                dgvScale.Rows(0).Cells(i).Style.BackColor = Color.Black
            ElseIf (j Mod 2) = 1 Then
                dgvScale.Rows(0).Cells(i).Style.BackColor = Color.DarkViolet
            Else
                dgvScale.Rows(0).Cells(i).Style.BackColor = Color.Gray
            End If
            j = j \ 2
            j35 = j35 \ 2
        Next
        For i = 0 To 6
            If (j35 Mod 2) = 1 Then
                dgvScale.Rows(0).Cells(i + 8).Style.BackColor = Color.Black
            ElseIf (j Mod 2) = 1 Then
                dgvScale.Rows(0).Cells(i + 8).Style.BackColor = Color.DarkViolet
            Else
                dgvScale.Rows(0).Cells(i + 8).Style.BackColor = Color.Gray
            End If
            j = j \ 2
            j35 = j35 \ 2
        Next
    End Sub
    Public Sub Dgv2Change()
        '107.4.29
        Dim i
        DataGridView2.RowCount = 3
        DataGridView2.ColumnCount = 30
        For i = 0 To 7
            If MatSiloOrder(i) > 0 Then
                If SiloMatOver1(i) Then
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Red
                    DataGridView2.Rows(2).Cells(i + 1).Value = "超量"
                    If bEngVer Then
                        DataGridView2.Rows(2).Cells(i + 1).Value = cOVER
                    End If
                ElseIf SiloMatUnder1(i) Then
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Fuchsia
                    DataGridView2.Rows(2).Cells(i + 1).Value = "不足"
                    If bEngVer Then
                        DataGridView2.Rows(2).Cells(i + 1).Value = cUNDER
                    End If
                ElseIf SiloMatStatus1(i) Then
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Green
                    DataGridView2.Rows(2).Cells(i + 1).Value = "計量中"
                    If bEngVer Then
                        DataGridView2.Rows(2).Cells(i + 1).Value = cWEIGHTING
                    End If
                ElseIf SiloMatDone1(i) Then
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Yellow
                    DataGridView2.Rows(2).Cells(i + 1).Value = "完成"
                    If bEngVer Then
                        DataGridView2.Rows(2).Cells(i + 1).Value = cDONE
                    End If
                Else
                    DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.White
                    DataGridView2.Rows(2).Cells(i + 1).Value = ""
                End If
            End If
        Next
        For i = 8 To 21
            If MatSiloOrder(i) > 0 Then
                If MatSiloNo(i) > 8 Then
                    If SiloMatOver2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Red
                        DataGridView2.Rows(2).Cells(i + 1).Value = "超量"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cOVER
                        End If
                    ElseIf SiloMatUnder2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Fuchsia
                        DataGridView2.Rows(2).Cells(i + 1).Value = "不足"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cUNDER
                        End If
                    ElseIf SiloMatStatus2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Green
                        DataGridView2.Rows(2).Cells(i + 1).Value = "計量中"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cWEIGHTING
                        End If
                    ElseIf SiloMatDone2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Yellow
                        DataGridView2.Rows(2).Cells(i + 1).Value = "完成"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cDONE
                        End If
                    Else
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.White
                        DataGridView2.Rows(2).Cells(i + 1).Value = ""
                    End If
                Else
                    If SiloMatOver2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Red
                        DataGridView2.Rows(2).Cells(i + 1).Value = "超量"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cOVER
                        End If
                    ElseIf SiloMatUnder2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Fuchsia
                        DataGridView2.Rows(2).Cells(i + 1).Value = "不足"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cUNDER
                        End If
                    ElseIf SiloMatStatus2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Green
                        DataGridView2.Rows(2).Cells(i + 1).Value = "計量中"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cWEIGHTING
                        End If
                    ElseIf SiloMatDone2(i - 8) Then
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.Yellow
                        DataGridView2.Rows(2).Cells(i + 1).Value = "完成"
                        If bEngVer Then
                            DataGridView2.Rows(2).Cells(i + 1).Value = cDONE
                        End If
                    Else
                        DataGridView2.Rows(2).Cells(i + 1).Style.BackColor = Color.White
                        DataGridView2.Rows(2).Cells(i + 1).Value = ""
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub TimerFlag_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TimerFlag.Tick
        '107.8.9 for FileSystemWatcher
        'My.Computer.FileSystem.WriteAllText(sYMdbName_B_Local & "2FLAG.log", queueR.UniCar & vbCrLf, True, System.Text.Encoding.Default)
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sYMdbName_B_Local & "2FLAG.log", OpenMode.Output)
        PrintLine(fileNum, queueR.UniCar)
        FileClose(fileNum)
        TimerFlag.Enabled = False
    End Sub

    Private Sub TimerMoveA_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMoveA.Tick
        Dim a
        a = MoveAToBuf(False, UniCar_MoveA, RemainDonePlate_MoveA)
        TimerFlag.Enabled = True
        TimerMoveA.Enabled = False
    End Sub

    Private Sub LabelRPC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelRPC.Click

    End Sub

    Private Sub TimerMoveB_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TimerMoveB.Tick
        '107.8.31
        '   add TimerMoveB only 環泥烏日bCheckBoxSP6 send remoteB

        If bCheckBoxSP6 And sTextBoxIP_R <> "127.0.0.1" Then
            LabelMessage2.Visible = True
            LabelMessage2.Text = "資料搜尋中(TimerSQL_B)請稍候(Wait...) "
            Me.Refresh()
            Label4.Image = My.Resources._0013
            Label4.Refresh()
            Dim b
            b = MoveLocalToRemoteBfile(sRemainUniCar, RemainDonePlate_temp_SQL)
            Label4.Image = My.Resources._000
            LabelMessage2.Visible = False
        End If
        TimerMoveB.Enabled = False

    End Sub


    Private Sub Timer232_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer232.Tick
        '108.2.27 read db delay
        BatchDone.DoneStep = 1
        mc.Send_232(queueR.CarSer, sRemainUniCar, RemainDonePlate_temp_SQL, sYMdbName & dtp_simtime.Value.Year & "M" & dtp_simtime.Value.Month & ".mdb")
        If BatchDone.PanCt = 1 Then
            BatchDone.PanBegin = 1
        Else
            BatchDone.PanBegin = 3
        End If
        Timer232.Enabled = False
        TimerBatDone.Enabled = True
    End Sub


    Private Sub numquetri1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles numquetri1.TextChanged
        ''109.10.5 台泥(bCheckBox232) 自動(CheckBoxContinue) 排隊一(Cbxquecode1) => 生產
        ''test
        'If bCheckBox232 And CheckBoxContinue.Checked And numquetri1.Text <> "" And Cbxquecode1.Text <> "" Then
        '    'Dim e As System.EventArgs = New System.EventArgs
        '    Call btnbegin_Click(CheckBoxContinue, e)
        'End If
    End Sub


    Private Sub Tbxquecar1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tbxquecar1.Leave
        '109.11.29 
        '110.6.7 車號簡檢查 bPbSync => bCheckBoxYadon
        'If bPbSync Then
        If bCheckBoxYadon Then
            If Tbxquecar1.Text.Length > 0 Then
                If Not Char.IsLetterOrDigit(Tbxquecar1.Text.Substring(0, 1)) Then
                    Tbxquecar1.Text = ""
                Else
                End If
            End If
            '110.9.15 check again when leave
            queue1check()
        End If
        Tbxquecar1.Text = Tbxquecar1.Text.ToUpper
    End Sub

    Private Sub Tbxquecar2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tbxquecar2.Leave
        '109.11.29 
        Tbxquecar2.Text = Tbxquecar2.Text.ToUpper
    End Sub

    Private Sub Cbxquecode1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbxquecode1.SelectedIndexChanged

    End Sub

    Private Sub LabelUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelUser.Click

    End Sub

    Private Sub LabelUser_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LabelUser.DoubleClick
        '110.1.14
        Dim r
        r = MsgBox("切換使用者?", MsgBoxStyle.OkCancel, "切換使用者")
        '112.4.19 ??If r = DialogResult.OK Then
        If r = Windows.Forms.DialogResult.OK Then
            module2.user = ""
            Dim Message, Title, sDefault
            Message = "請輸入新使用者"    ' Set prompt.
            Title = "切換使用者"    ' Set title.
            sDefault = ""    ' Set default.
            ' Display message, title, and default value.
            module2.user = InputBox(Message, Title, sDefault)
            Dim db As New DbAccess(sSMdbName & "\setting.mdb")
            Dim dt As DataTable = db.GetDataTable("select * from passwd where ID = '" + module2.user + "'")
            Dim s
            'sPBUser = InputBox("請輸入使用者名稱", "配方")
            If module2.user <> "" Then
                module2.callfrm = 0
                If dt.Rows.Count <> 0 Then
                    s = dt.Rows(0).Item("ID").ToString
                    module2.password = dt.Rows(0).Item("Passwd").ToString
                    module2.authority = dt.Rows(0).Item("setting").ToString
                    module2.callfrm = 1 '
                    PasswordTextBox.Text = ""
                    Panel7.Visible = True
                    PasswordTextBox.Focus()
                Else
                    If bEngVer Then
                        MsgBox("User not exist!!!!", MsgBoxStyle.Information)
                    Else
                        MsgBox("使用者不存在!!!!", MsgBoxStyle.Information)
                    End If
                End If
            Else
                If bEngVer Then
                    MsgBox("Please keyin User Name", MsgBoxStyle.Critical)
                Else
                    MsgBox("請輸入使用者名稱", MsgBoxStyle.Critical)
                End If
            End If
        Else
            'PasswordTextBox.Text = ""
            'Panel7.Visible = True
            'PasswordTextBox.Focus()
        End If


    End Sub


    Private Sub TimerShowCement_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerShowCement.Tick
        LabelCement.Visible = False
        TimerShowCement.Enabled = False
    End Sub

    Private Sub numquetri1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numquetri1.SelectedIndexChanged

    End Sub

    Private Sub TimerAlarm_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerAlarm.Tick
        '110.6.24
        Call AlarmCheck()

    End Sub
    Public Sub AlarmCheck_old_20240826()

        '110.9.11
        '107.4.28 move from FrmMonit Timer2, use 
        AlarmDV(0) = DV(16)
        AlarmDV(1) = DV(17)
        AlarmDV(2) = DV(18)
        AlarmDV(3) = DV(19)
        AlarmDV(4) = DV(29)
        AlarmDV(5) = DV(37)
        AlarmDV(6) = DV(40)

        Dim i, j, k, l, m As Integer
        j = 0
        For i = 0 To 6
            If AlarmDVOld(i) <> AlarmDV(i) Then j = 1
        Next
        '110.6.24 DO NOT CARE 0
        '110.4.8
        'For i = 0 To 6
        '    If AlarmDV(i) <> 0 Then j = 1
        'Next
        'If FrmMonit.COM_PLC.Err <> bPlcErrOld Then j = 1
        'bPlcErrOld = FrmMonit.COM_PLC.Err


        '111.1.20
        'If FrmMonit.COM_PLC.Err <> bPlcErrOld Then
        '    j = 1
        'End If
        If FrmMonit.COM_PLC.Err Then
            COM_PLC_Err += 1
        Else
            COM_PLC_Err = 0
        End If
        '110.1.20
        If COM_PLC_Err >= 2 Then
            labelAlarmTime(9).Text = Format(Now, "HH:mm:ss")
            labelAlarmDesc(9).Text = "PLC通信異常 " & COM_PLC_Err
            labelAlarmTime(9).ForeColor = Color.Red
            labelAlarmDesc(9).ForeColor = Color.Red
        Else
            labelAlarmTime(9).Text = ""
            labelAlarmDesc(9).Text = ""
        End If

        If j = 0 Then
            '110.4.8
            Exit Sub
        End If

        For l = 0 To 6
            j = AlarmDV(l)
            k = AlarmDVOld(l)
            For i = 0 To 14
                If j Mod 2 = 1 Then
                    AlarmValue(i + 15 * l) = True
                    If k Mod 2 = 0 Then AlarmTime(i + 15 * l) = Now
                Else
                    AlarmValue(i + 15 * l) = False
                End If
                j \= 2
                k \= 2
            Next
        Next

        j = 0

        For i = 0 To 104
            If AlarmValue(i) = True And j < 10 Then
                j += 1
                '107.4.28
                m = DataGridViewAlarm.RowCount
                DataGridViewAlarm.RowCount = 10 '???
                DataGridViewAlarm.Rows(j - 1).Cells(0).Value = Format(AlarmTime(i), "HH:mm:ss")
                DataGridViewAlarm.Rows(j - 1).Cells(1).Value = AlarmDesc(i)
                '110.6.24
                labelAlarmTime(j - 1).Text = Format(AlarmTime(i), "HH:mm:ss")
                labelAlarmDesc(j - 1).Text = AlarmDesc(i)
                Select Case AlarmType(i)
                    Case 1
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Green
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Green
                        labelAlarmTime(j - 1).ForeColor = Color.Green
                        labelAlarmDesc(j - 1).ForeColor = Color.Green
                    Case 2
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Cyan
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Cyan
                        labelAlarmTime(j - 1).ForeColor = Color.Cyan
                        labelAlarmDesc(j - 1).ForeColor = Color.Cyan
                    Case 3
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Yellow
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Yellow
                        labelAlarmTime(j - 1).ForeColor = Color.Yellow
                        labelAlarmDesc(j - 1).ForeColor = Color.Yellow
                    Case Else
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Fuchsia
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Fuchsia
                        labelAlarmTime(j - 1).ForeColor = Color.Fuchsia
                        labelAlarmDesc(j - 1).ForeColor = Color.Fuchsia
                End Select
            Else
                'DataGridViewAlarm.Rows(j).Cells(1).Value = AlarmDesc(i)
            End If
        Next

        '111.1.20 move to up
        'If FrmMonit.COM_PLC.Err = True And j < 10 Then
        'If COM_PLC_Err > 2 And j < 10 Then
        '    j += 1
        '    '111.1.20
        '    'DataGridViewAlarm.Rows(j - 1).Cells(0).Value = Format(Now, "HH:mm:ss")
        '    'DataGridViewAlarm.Rows(j - 1).Cells(1).Value = "PLC通信異常"
        '    'DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Fuchsia
        '    'DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Fuchsia
        '    '110.6.24
        '    '111.1.20 labelAlarmTime(j - 1).Text = Format(AlarmTime(i), "HH:mm:ss")
        '    labelAlarmTime(j - 1).Text = Format(Now, "HH:mm:ss")
        '    labelAlarmDesc(j - 1).Text = "PLC通信異常 " & COM_PLC_Err
        'End If

        If j < 10 Then
            '111.1.20 labelAlarmDesc(9).Text for PLC
            'For i = j To 9
            For i = j To 8
                DataGridViewAlarm.Rows(i).Cells(0).Value = ""
                DataGridViewAlarm.Rows(i).Cells(1).Value = ""
                '110.4.8 110.6.24
                'DataGridViewAlarm.Refresh()
                '110.6.24
                labelAlarmTime(i).Text = ""
                labelAlarmDesc(i).Text = ""
            Next
        End If
        '110.4.8
        'DataGridViewAlarm.Sort(DataGridViewAlarm.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        DataGridViewAlarm.Rows(9).Cells(1).Selected = True
        For i = 0 To 6
            AlarmDVOld(i) = AlarmDV(i)
        Next
        '110.4.8 110.6.24
        DataGridViewAlarm.Refresh()
    End Sub

    Public Sub AlarmCheck()
        '113.8.26 
        '   add fuse alarm
        '   D341-D356   保險絲燒毀bit(每個D 16bit) 由D33-12決定掃瞄
        '
        ' Module_Public
        ' Alarm D16~19(運轉狀況1~4), D29(運轉狀況5:閘門失誤),D37(運轉狀況6:骨材入倉),D40(運轉狀況7:泥料入倉) 15*7=105
        '   
        'Public AlarmName(105) As String 'Alarm description D16-00 ~ D19-15
        'Public AlarmDesc(105) As String 'Alarm description D16-00 ~ D19-15
        'Public AlarmType(105) As Integer  'Alarm type 1:Green, 2:Cyan, 3:Yellow, 4:lightRed
        'Public AlarmValue(105) As Boolean  'Alarm type D16-00 ~ D19-15 true=alarm  false=none
        'Public AlarmTime(105) As DateTime  'Alarm DateTime
        'Public AlarmDV(7) As Integer    'Alarm DV D16~19 29 37 40 AlarmDV(7) for PC:COM..
        'Public AlarmDVOld(7) As Integer    'Alarm DV D16~19
        Dim i, j, k, l, m As Integer
        AlarmDV(0) = DV(16)
        AlarmDV(1) = DV(17)
        AlarmDV(2) = DV(18)
        AlarmDV(3) = DV(19)
        AlarmDV(4) = DV(29)
        AlarmDV(5) = DV(37)
        AlarmDV(6) = DV(40)
        '113.8.26 fuse alarm
        For i = 7 To AlarmDvMax
            AlarmDV(i) = DV(i + 334)  'AlarmDV(7)=DV(341)
        Next


        j = 0
        '113.8.26
        'For i = 0 To 6
        '    If AlarmDVOld(i) <> AlarmDV(i) Then j = 1
        'Next
        For i = 0 To AlarmDvMax
            If AlarmDVOld(i) <> AlarmDV(i) Then j = 1
        Next

        '110.6.24 DO NOT CARE 0
        '110.4.8
        'For i = 0 To 6
        '    If AlarmDV(i) <> 0 Then j = 1
        'Next
        'If FrmMonit.COM_PLC.Err <> bPlcErrOld Then j = 1
        'bPlcErrOld = FrmMonit.COM_PLC.Err


        '111.1.20
        'If FrmMonit.COM_PLC.Err <> bPlcErrOld Then
        '    j = 1
        'End If
        If FrmMonit.COM_PLC.Err Then
            COM_PLC_Err += 1
        Else
            COM_PLC_Err = 0
        End If
        '110.1.20
        If COM_PLC_Err >= 2 Then
            labelAlarmTime(9).Text = Format(Now, "HH:mm:ss")
            labelAlarmDesc(9).Text = "PLC通信異常 " & COM_PLC_Err
            labelAlarmTime(9).ForeColor = Color.Red
            labelAlarmDesc(9).ForeColor = Color.Red
        Else
            labelAlarmTime(9).Text = ""
            labelAlarmDesc(9).Text = ""
        End If

        If j = 0 Then
            '110.4.8
            Exit Sub
        End If

        '113.8.26
        'For l = 0 To 6
        '    j = AlarmDV(l)
        '    k = AlarmDVOld(l)
        '    For i = 0 To 14
        '        If j Mod 2 = 1 Then
        '            AlarmValue(i + 15 * l) = True
        '            If k Mod 2 = 0 Then AlarmTime(i + 15 * l) = Now
        '        Else
        '            AlarmValue(i + 15 * l) = False
        '        End If
        '        j \= 2
        '        k \= 2
        '    Next
        'Next
        For l = 0 To AlarmDvMax
            j = AlarmDV(l)
            k = AlarmDVOld(l)
            For i = 0 To 15
                If j Mod 2 = 1 Then
                    Try
                        AlarmValue(i + 16 * l) = True
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                    If k Mod 2 = 0 Then AlarmTime(i + 16 * l) = Now
                Else
                    AlarmValue(i + 16 * l) = False
                End If
                j \= 2
                k \= 2
            Next
        Next

        j = 0

        '113.8.26
        'For i = 0 To 104
        '    If AlarmValue(i) = True And j < 10 Then
        '        j += 1
        '        '107.4.28
        '        m = DataGridViewAlarm.RowCount
        '        DataGridViewAlarm.RowCount = 10 '???
        '        DataGridViewAlarm.Rows(j - 1).Cells(0).Value = Format(AlarmTime(i), "HH:mm:ss")
        '        DataGridViewAlarm.Rows(j - 1).Cells(1).Value = AlarmDesc(i)
        '        '110.6.24
        '        labelAlarmTime(j - 1).Text = Format(AlarmTime(i), "HH:mm:ss")
        '        labelAlarmDesc(j - 1).Text = AlarmDesc(i)
        '        Select Case AlarmType(i)
        '            Case 1
        '                DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Green
        '                DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Green
        '                labelAlarmTime(j - 1).ForeColor = Color.Green
        '                labelAlarmDesc(j - 1).ForeColor = Color.Green
        '            Case 2
        '                DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Cyan
        '                DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Cyan
        '                labelAlarmTime(j - 1).ForeColor = Color.Cyan
        '                labelAlarmDesc(j - 1).ForeColor = Color.Cyan
        '            Case 3
        '                DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Yellow
        '                DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Yellow
        '                labelAlarmTime(j - 1).ForeColor = Color.Yellow
        '                labelAlarmDesc(j - 1).ForeColor = Color.Yellow
        '            Case Else
        '                DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Fuchsia
        '                DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Fuchsia
        '                labelAlarmTime(j - 1).ForeColor = Color.Fuchsia
        '                labelAlarmDesc(j - 1).ForeColor = Color.Fuchsia
        '        End Select
        '    Else
        '        'DataGridViewAlarm.Rows(j).Cells(1).Value = AlarmDesc(i)
        '    End If
        'Next

        For i = 0 To AlarmDvBitsMax
            If AlarmValue(i) = True And j < 10 Then
                j += 1
                m = DataGridViewAlarm.RowCount
                DataGridViewAlarm.RowCount = 10 '???
                DataGridViewAlarm.Rows(j - 1).Cells(0).Value = Format(AlarmTime(i), "HH:mm:ss")
                DataGridViewAlarm.Rows(j - 1).Cells(1).Value = AlarmDesc(i)
                labelAlarmTime(j - 1).Text = Format(AlarmTime(i), "HH:mm:ss")
                labelAlarmDesc(j - 1).Text = AlarmDesc(i)
                Select Case AlarmType(i)
                    Case 1
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Green
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Green
                        labelAlarmTime(j - 1).ForeColor = Color.Green
                        labelAlarmDesc(j - 1).ForeColor = Color.Green
                    Case 2
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Cyan
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Cyan
                        labelAlarmTime(j - 1).ForeColor = Color.Cyan
                        labelAlarmDesc(j - 1).ForeColor = Color.Cyan
                    Case 3
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Yellow
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Yellow
                        labelAlarmTime(j - 1).ForeColor = Color.Yellow
                        labelAlarmDesc(j - 1).ForeColor = Color.Yellow
                    Case Else
                        DataGridViewAlarm.Rows(j - 1).Cells(0).Style.ForeColor = Color.Fuchsia
                        DataGridViewAlarm.Rows(j - 1).Cells(1).Style.ForeColor = Color.Fuchsia
                        labelAlarmTime(j - 1).ForeColor = Color.Fuchsia
                        labelAlarmDesc(j - 1).ForeColor = Color.Fuchsia
                End Select
            Else
                'DataGridViewAlarm.Rows(j).Cells(1).Value = AlarmDesc(i)
            End If
        Next

        If j < 10 Then
            '111.1.20 labelAlarmDesc(9).Text for PLC
            'For i = j To 9
            For i = j To 8
                DataGridViewAlarm.Rows(i).Cells(0).Value = ""
                DataGridViewAlarm.Rows(i).Cells(1).Value = ""
                labelAlarmTime(i).Text = ""
                labelAlarmDesc(i).Text = ""
            Next
        End If
        DataGridViewAlarm.Rows(9).Cells(1).Selected = True
        '113.8.26
        'For i = 0 To 6
        '    AlarmDVOld(i) = AlarmDV(i)
        'Next
        For i = 0 To AlarmDvMax
            AlarmDVOld(i) = AlarmDV(i)
        Next
        DataGridViewAlarm.Refresh()
    End Sub
    Private Sub CheckLabelJSForeColor()
        '110.8.1
        '   CheckLabelForeColor of LabelRPC / LabelRPC 
        bMyComputerNetworkIsAvailable = My.Computer.Network.IsAvailable
        If bMyComputerNetworkIsAvailable Then
            bMyComputerNetworkPingDataBaseIP = My.Computer.Network.Ping(sDataBaseIP, 1000)
            bMyComputerNetworkPingTextBoxIP_R = My.Computer.Network.Ping(sTextBoxIP_R, 1000)
        Else
            bMyComputerNetworkPingDataBaseIP = False
            bMyComputerNetworkPingTextBoxIP_R = False
        End If
        If bMyComputerNetworkIsAvailable Then
            'Database Server of 時訊
            If bMyComputerNetworkPingDataBaseIP Then
                LabelJS.ForeColor = Color.DarkGreen
            Else
                LabelJS.ForeColor = Color.DarkRed
            End If
            'RemotePC of JS
            If bMyComputerNetworkPingTextBoxIP_R Then
                LabelRPC.ForeColor = Color.DarkGreen
            Else
                LabelRPC.ForeColor = Color.DarkRed
            End If
        Else
            LabelJS.ForeColor = Color.Black
            LabelRPC.ForeColor = Color.Black
        End If

    End Sub

    Private Sub Cbxquecode1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cbxquecode1.TextChanged
        '110.7.10 110.9.11
        If Cbxquecode1.Text <> "" Then
            Label11.ForeColor = Color.Blue
            DVW(829) = 1
            i = FrmMonit.WriteA2N(829, 1)
        Else
            Label11.ForeColor = Color.Black
            DVW(829) = 0
            i = FrmMonit.WriteA2N(829, 1)
        End If
    End Sub
    Private Sub UpdateLabelTotPlate()
        '110.11.14
        LabelTotPlate.Text = queue1.TotalPlate & "-" & queue0.TotalPlate & "-" & queueP.TotalPlate & "-" & queueR.TotalPlate
    End Sub

    Private Sub NumAE_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumAE.Leave
        '111.4.2
        If FormLoded = False Then Exit Sub
        NumAE.BackColor = Color.SkyBlue

    End Sub

    Private Sub NumAE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumAE.ValueChanged
        '111.4.2 藥劑% ref NumSG 
        If FormLoded = False Then Exit Sub
        If queue0.inprocess = True Then
            mc.queclone(queue0, queue)
            queue.BoneDoingPlate = queue0.BoneDoingPlate
            queue.ModDoingPlate = queue0.ModDoingPlate
            If queue0.BoneDoingPlate = queue0.ModDoingPlate Then
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "A")
            Else
                queue = mc.Materialdistribute(queue, queue0.BoneDoingPlate, cbxcorrect.Text, waterall, "B")
            End If

            mc.queclone(queue, queue0)
            MonLog.add = "displaymaterial() : NumAE"
            displaymaterial()
        Else
        End If

        NumAE.BackColor = Color.Yellow
    End Sub

    Private Sub PanelSim_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PanelSim.GotFocus
    End Sub


    Private Sub CheckBoxFirstPan_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBoxFirstPan.Validated
        '111.7.22
        '   模擬報表 首盤 CheckBoxFirstPan
        bCheckBoxFirstPan = Not bCheckBoxFirstPan
        CheckBoxFirstPan.Checked = bCheckBoxFirstPan
        SaveSetting("JS", "CBC800", "bCheckBoxFirstPan", CStr(bCheckBoxFirstPan))

    End Sub

    Public Function fillMdbRowToPan(ByVal row As Object, ByVal pre As String) As classReceiaPan
        '111.9.20
        'row : dataRow from mdb Recdata
        Dim reciaReadHeader As New ClassReceiaHeader(row)
        Dim receiaPan As New classReceiaPan()
        receiaPan.Header = reciaReadHeader
        receiaPan.D2.valRec = receiaPan.fillRecCat(row, "org")
        receiaPan.D3.valRec = receiaPan.fillRecCat(row, "set")
        receiaPan.D4.valRec = receiaPan.fillRecCat(row, "")
        receiaPan.D5.valRec = receiaPan.fillRecCat(row, "remain")
        receiaPan.D6.valRec = receiaPan.fillRecErrPer(receiaPan.D3.valRec, receiaPan.D4.valRec, receiaPan.D5.valRec)
        receiaPan.D7.valRec = receiaPan.fillRecWaterComp(row, "Watercomp")

        '112.9.29 112.10.1
        If pre = "B" Then
            Dim kk
            Dim sum As Single = 0
            For kk = 0 To receiaPanB.RT.valRec.Length - 1 's1(0)..ae6(22) tot(23)
                receiaPanB.RT.valRec(kk) += receiaPan.D4.valRec(kk) - receiaPan.D5.valRec(kk)    'real-remain
                sum += receiaPan.DT.valRec(kk)
                '112.10.1
                receiaPan.DT.valRec(kk) = receiaPanB.RT.valRec(kk)    'real-remain
            Next
            kk = receiaPan.DT.valRec.Length - 1
            receiaPanB.RT.valRec(kk) += sum
            receiaPan.DT.valRec(kk) = receiaPanB.RT.valRec(kk)    'real-remain
            receiaPanB.RT.qty += receiaPan.Header.qty
            receiaPan.DT.qty = receiaPanB.RT.qty
        End If

        fillMdbRowToPan = receiaPan

    End Function


    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label1SQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1SQL.Click
        '112.1.5 test
        If Not bPingSQLBigData Then
            Exit Sub
        End If
        'Dim dbBigData As New SQLDbAccess2("conscbuji", sBigDataDataBaseIP)
        Dim dbBigData As New SQLDbAccess2("BigData", sBigDataDataBaseIP)
        Dim dtBigData As DataTable
        Dim stsql As String = "select * from receia where qty > 2 "
        If bPingSQLBigData Then
            dtBigData = dbBigData.GetDataTable(stsql)
            Label1SQL.Text &= ".. dtBigData.Rows.Count=" & dtBigData.Rows.Count
        End If

    End Sub

    Private Sub Tbxworkcar_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tbxworkcar.TextChanged
        '112.3.1
        'If bPingSQLBigData And Tbxworkcar.Text <> "" Then
        If Tbxworkcar.Text <> "" Then
            Dim str, s, s2 As String
            str = Microsoft.VisualBasic.Left(Tbxworkcar.Text.Trim, 8)
            Dim i, len As Integer
            len = str.Length
            Dim iCarId(3) As Integer
            For i = 0 To len - 1
                s = Microsoft.VisualBasic.Mid(str, i + 1, 1)
                s2 = Microsoft.VisualBasic.Mid(str, i + 2, 1)
                Dim codeInt As Integer
                Dim codeInt2 As Integer
                ' The following line of code sets codeInt to 65.
                'codeInt = Asc("A")
                If (i Mod 2) = 0 Then
                    If s2 = "" Then
                        codeInt = Asc(s)
                        codeInt2 = 0
                    Else
                        codeInt = Asc(s)
                        codeInt2 = Asc(s2)
                    End If
                    iCarId(i \ 2) = codeInt * 256 + codeInt2
                End If
            Next
            For i = 0 To 3
                DVW(4001 + i) = iCarId(i)
            Next
            i = FrmMonit.WriteA2N(4001, 4)
        End If
    End Sub


 

    Private Sub LabelBigDataErr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelBigDataErr.Click
        '112.4.12
        Label10.ForeColor = Color.Yellow
        If My.Computer.Network.Ping(sBigDataDataBaseIP, 500) Then
            bPingSQLBigData = True
            Label10.ForeColor = Color.Green
            LabelBigDataErr.Visible = False
        Else
            Label10.ForeColor = Color.Black
        End If

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        '110.10.25
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sYMdbName_B_Local & "2FLAG.log", OpenMode.Output)
        PrintLine(fileNum, 9)
        FileClose(fileNum)
    End Sub

    Private Sub MainForm_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

    End Sub
End Class