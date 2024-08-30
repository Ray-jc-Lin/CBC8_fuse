Public Class FrmAdmin
    Private db As New DbAccess(sSMdbName & "\setting.mdb")
    Dim senddButton As Integer
    Dim temp
    Dim Keys_X As Boolean = False

    Private Sub FrmAdmin_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Keys_X = False

    End Sub

    Private Sub FrmAdmin_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Keys_X = False
    End Sub

    Private Sub FrmAdmin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '107.6.27
        If e.KeyCode = Keys.X And e.Modifiers = Keys.Alt Then
            If Keys_X Then
                CheckBoxABPB.Text = "X配比查詢"
                Label26.Text = "X存放資料夾(\\192.)"
                Label27.Text = "遠端A存放資料夾"
                CheckBoxBPBsubEn.Text = "轉料剩料處理"
                CheckBoxSP3.Text = "X配存檔"
                CheckBoxSP13.Text = "車輛回傳測試"
                CheckBoxSP14.Text = "不回傳R1R2R3R5"
                '107.12.28
                CheckBoXSP15.Text = "X配調整"
            Else
                CheckBoxABPB.Text = "配比查詢"
                Label26.Text = "存放資料夾(\\192.)"
                Label27.Text = "遠端存放資料夾"
                CheckBoxBPBsubEn.Text = "轉料剩料處理"
                CheckBoxSP3.Text = "存檔"
                CheckBoxSP13.Text = "車輛回傳測試"
                CheckBoxSP14.Text = "不回傳"
                '107.12.28
                CheckBoXSP15.Text = "調整"
            End If
            Keys_X = Not Keys_X
        End If

    End Sub

    Private Sub FrmAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '110.1.20   110.2.5
        'LogLabel1.Text = LoginForm1.LabelNet.Text & vbCrLf & LoginForm1.Label1.Text
        LogLabel1.Text = LoginForm1.LabelNet.Text & vbCrLf & LoginForm1.Label1.Text & vbCrLf & Now()
        '107.6.27
        Keys_X = False
        CheckBoxABPB.Text = "配比查詢"
        Label26.Text = "存放磁碟(\\192.)"
        Label27.Text = "遠端存放磁碟"
        CheckBoxBPBsubEn.Text = "米數"
        CheckBoxSP3.Text = "存檔"
        CheckBoxSP13.Text = "備用空"
        CheckBoxSP14.Text = "不回傳"
        Me.CenterToScreen()
        Frmshowsetting.Hide()
        Dim dt As DataTable = db.GetDataTable("select title, material, updiff, 全滿刻度 from config where type= 'Pond' order by Idx")
        Dim i = 0
        Dim str() As String = {"磅一", "磅二", "磅三", "磅四", "磅五", "磅六", "磅七", "磅八", _
                               "磅九", "磅十", "磅十一", "磅十二", "磅十三", "磅十四", "磅十五"}
        'Dim str() As String = {"磅秤一", "磅秤二", "磅秤三", "磅秤四", "磅秤五", "磅秤六", "磅秤七", "磅秤八", _
        '                       "磅秤九", "磅秤十", "磅秤十一", "磅秤十二", "磅秤十三", "磅秤十四", "磅秤十五"}
        DataGridView1.Rows.Add(8)
        For i = 0 To 7
            DataGridView1.Rows(i).Cells(0).Value = str(i)
            DataGridView1.Rows(i).Cells(1).Value = dt.Rows(i).Item(0).ToString
            DataGridView1.Rows(i).Cells(2).Value = dt.Rows(i).Item(1).ToString
            DataGridView1.Rows(i).Cells(3).Value = dt.Rows(i).Item(2).ToString
            DataGridView1.Rows(i).Cells(4).Value = dt.Rows(i).Item(3).ToString
            If i <> 7 Then
                DataGridView1.Rows(i).Cells(5).Value = str(i + 8)
                DataGridView1.Rows(i).Cells(6).Value = dt.Rows(i + 8).Item(0).ToString
                DataGridView1.Rows(i).Cells(7).Value = dt.Rows(i + 8).Item(1).ToString
                DataGridView1.Rows(i).Cells(8).Value = dt.Rows(i + 8).Item(2).ToString
                DataGridView1.Rows(i).Cells(9).Value = dt.Rows(i + 8).Item(3).ToString
            Else
                DataGridView1.Rows(i).Cells(6).ReadOnly = True
                DataGridView1.Rows(i).Cells(6).Style.BackColor = Color.Moccasin
                DataGridView1.Rows(i).Cells(7).ReadOnly = True
                DataGridView1.Rows(i).Cells(7).Style.BackColor = Color.Moccasin
                DataGridView1.Rows(i).Cells(5).ReadOnly = True
                DataGridView1.Rows(i).Cells(5).Style.BackColor = Color.Moccasin
            End If
        Next
        DataGridView2.Rows.Add(14)
        i = DataGridView2.RowCount
        Me.CenterToScreen()
        'MainForm.Hide()
        'Frmshowsetting.Hide()

        DgvPLC.RowCount = 20
        Me.Width = TabControl1.Width + 6
        TabControl1.Left = 2
        TabControl1.Top = 40
        TabControl1.Height = Me.Height - 80

        'COMBOBOXCOM
        '100.4.1
        For i = 0 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For i = 0 To DataGridView2.ColumnCount - 1
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For i = 0 To DataGridView3.ColumnCount - 1
            DataGridView3.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For i = 0 To DataGridView4.ColumnCount - 1
            DataGridView4.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        '107.6.16 107.6.27
        'Call ReadAdminName()
        '
        Me.KeyPreview = True

    End Sub

    Private Sub TabPage2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Enter
        Dim dt As DataTable
        Dim i = 0
        Dim j = 0
        Dim str() As String = {"砂一", "砂二", "砂三", "砂四", "石一", "石二", "石三", "石四", "水一", "水二", "水三", _
                               "水泥一", "水泥二", "水泥三", "水泥四", "水泥五", "水泥六", "藥一", "藥二", "藥三", "藥四", "藥五", _
                               "材料一", "材料二", "材料三", "預備一", "預備二", "預備三"}
        Dim s As String
        Dim s2 As String
        Dim s3 As String
        dt = db.GetDataTable("select material from config where type = 'Pond' order by Idx")
        'materialApond.Items.Clear()
        'materialApond.Items.Add("0")
        'materialBpond.Items.Clear()
        'materialBpond.Items.Add("0")
        For i = 0 To dt.Rows.Count - 1
            'If dt.Rows(i).Item(0).ToString = "使用" Then materialApond.Items.Add((i + 1).ToString)
            'If dt.Rows(i).Item(0).ToString = "使用" Then materialBpond.Items.Add((i + 1).ToString)
        Next

        'dtconfig = db.GetDataTable("select title, setting, scale,sr,Material from config where type = 'Material' order by Idx")
        dt = db.GetDataTable("select title, setting from config where type= 'Material' order by Idx")
        'dt = db.GetDataTable("select title, setting, memo1 from config where type= 'Material' order by Idx")
        i = DataGridView2.RowCount
        s2 = ""
        For i = 0 To 10
            DataGridView2.Rows(i).Cells(0).Value = str(i)
            DataGridView2.Rows(i).Cells(1).Value = dt.Rows(i).Item(0).ToString
            s2 += i & ":" & dt.Rows(i).Item(1).ToString & ","
            s3 = dt.Rows(i).Item(1).ToString
            If s3 = "" Then s3 = "99"
            Try
                'DataGridView2.Rows(i).Cells(2).Value = dt.Rows(i).Item(1).ToString
                DataGridView2.Rows(i).Cells(2).Value = s3
            Catch ex As Exception
                'DataGridView2.Rows(i).Cells(2).Value = "0"
            End Try
            s = DataGridView2.Rows(i).Cells(2).Value
        Next
        'TextBox1.Text = s2

        For i = 11 To 24
            DataGridView2.Rows(i - 11).Cells(3).Value = str(i)
            DataGridView2.Rows(i - 11).Cells(4).Value = dt.Rows(i).Item(0).ToString
            s2 += i & ":" & dt.Rows(i).Item(1).ToString & ","
            Try
                DataGridView2.Rows(i - 11).Cells(5).Value = dt.Rows(i).Item(1).ToString
            Catch ex As Exception
                ' DataGridView2.Rows(i - 11).Cells(5).Value = "0"
            End Try
        Next
        TextBox1.Text += s2
        'For i = 0 To 13
        '    DataGridView2.Rows(i).Cells(0).Value = str(i)
        '    DataGridView2.Rows(i).Cells(3).Value = str(i + 14)
        '    DataGridView2.Rows(i).Cells(1).Value = dt.Rows(i).Item(0).ToString
        '    DataGridView2.Rows(i).Cells(4).Value = dt.Rows(i + 14).Item(0).ToString
        '    Try
        '        DataGridView2.Rows(i).Cells(2).Value = dt.Rows(i).Item(1).ToString
        '        DataGridView2.Rows(i).Cells(5).Value = dt.Rows(i + 14).Item(1).ToString
        '    Catch ex As Exception
        '        DataGridView2.Rows(i).Cells(2).Value = "0"
        '        DataGridView2.Rows(i).Cells(5).Value = "0"
        '    End Try
        'Next

        '100.6.29
        CheckBoxS.Checked = bCheckBoxS
        CheckBoxG.Checked = bCheckBoxG
        CheckBoxW.Checked = bCheckBoxW
        CheckBoxC.Checked = bCheckBoxC
        '101.4.5
        CheckBoxS4.Checked = bCheckBoxS4
        '102.10.19
        ReadAltDes()
        '104.9.24
        CheckBoxC1.Checked = bCheckBoxC1
        CheckBoxC2.Checked = bCheckBoxC2
        CheckBoxC3.Checked = bCheckBoxC3
        CheckBoxC4.Checked = bCheckBoxC4
        CheckBoxC5.Checked = bCheckBoxC5
        CheckBoxC6.Checked = bCheckBoxC6
        '107.8.20
        '   聯佶 W1 不參與水分補正 
        CheckBoxW1.Checked = bCheckBoxW1
        CheckBoxW2.Checked = bCheckBoxW2
        CheckBoxW3.Checked = bCheckBoxW3
        '110.6.17   台泥 藥劑區域
        CheckBoxAE.Checked = bCheckBoxAE
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim str() As String = {"Pond1", "Pond2", "Pond3", "Pond4", "Pond5", "Pond6", "Pond7", "Pond8", _
                                "Pond9", "Pond10", "Pond11", "Pond12", "Pond13", "Pond14", "Pond15"}
        Dim sql As String = ""
        Dim i = e.RowIndex
        'If e.ColumnIndex > 3 Then i = i + 8
        If e.ColumnIndex > 4 Then i = i + 8
        Dim col As String = ""
        If e.ColumnIndex = 1 Or e.ColumnIndex = 6 Then col = "title"
        If e.ColumnIndex = 2 Or e.ColumnIndex = 7 Then col = "material"
        If e.ColumnIndex = 3 Or e.ColumnIndex = 8 Then col = "updiff"
        If e.ColumnIndex = 4 Or e.ColumnIndex = 9 Then col = "全滿刻度"
        sql = "update config set " + col + " = '" + DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value + "' where ID ='" + str(i) + "'"
        db.ExecuteCmd(sql)
    End Sub

    Private Sub DataGridView2_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        Dim str() As String = {"Sand1", "Sand2", "Sand3", "Sand4", "Stone1", "Stone2", "Stone3", "Stone4", _
                                "Water1", "Water2", "Water3", "Concrete1", "Concrete2", "Concrete3", "Concrete4", "Concrete5", "Concrete6", _
                                "Drog1", "Drog2", "Drog3", "Drog4", "Drog5", "Material1", "Material2", "Material3", "memo1", "memo2", "memo3"}
        Dim sql As String = ""
        Dim i = e.RowIndex
        'If e.ColumnIndex > 4 Then i = i + 13
        If e.ColumnIndex > 3 Then i = i + 11
        Dim col As String = ""
        If e.ColumnIndex = 1 Or e.ColumnIndex = 4 Then col = "title"
        If e.ColumnIndex = 2 Or e.ColumnIndex = 5 Then col = "setting"
        'If e.ColumnIndex = 1 Or e.ColumnIndex = 5 Then col = "title"
        'If e.ColumnIndex = 2 Or e.ColumnIndex = 6 Then col = "setting"
        'If e.ColumnIndex = 3 Or e.ColumnIndex = 7 Then col = "memo1"
        sql = "update config set " + col + " = '" + DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value + "' where ID ='" + str(i) + "'"
        db.ExecuteCmd(sql)
    End Sub

    Private Sub DataGridView2_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

    End Sub


    Private Sub FrmAdmin_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Keys_X = False
        MainForm.Show()
        Frmshowsetting.Show()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Keys_X = False
        Me.Close()
    End Sub

 
    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter
        '112.3.25 列印拌合結束電流 from LabelCurrent3 when D22 change
        CheckBoxCBC7.Text = "列印拌合機電流"
        '103.4.25 104.9.4
        CheckBoxRemain.Checked = bCheckBoxRemain
        CheckBoxSP1.Checked = bCheckBoxSP1
        CheckBoxSP2.Checked = bCheckBoxSP2
        CheckBoxSP3.Checked = bCheckBoxSP3
        '103.6.7
        CheckBoxNonAuto.Checked = bCheckBoxNonAuto
        '103.3.11
        CheckBoxImperial.Checked = bCheckBoxImperial
        '102.12.11
        CheckBoxBack.Checked = bCheckBoxBack
        '102.11.23
        CheckBoxCement.Checked = bCheckBoxCement
        '102.7.15
        CheckBoxBPBsubEn.Checked = bCheckBoxBPBsuben
        '102.5.11
        '112.3.25 bCheckBoxCBC7 改成列印拌合結束電流
        CheckBoxCBC7.Checked = bCheckBoxCBC7
        CheckBoxWaitTank.Checked = bWaitTank
        CheckBoxABPB.Checked = bABPB
        '100.3.26
        CheckBoxYadon.Checked = bCheckBoxYadon
        '101.3.7
        '102.4.17 CheckBoxsingle.Checked = bCheckBoxSingle
        CheckBoxsingle.Checked = bEngVer
        '101.4.2
        CheckBoxtTest.Checked = bCheckBoxtTest
        TextBoxMaxM3.Text = fCarMaxM3
        TextBoxZeroMax.Text = fScaleZeroMax
        TextBoxSpanMax.Text = fScaleSpanMax
        TextBoxSpanMin.Text = fScaleSpanMin
        TextBoxCompany.Text = sCompany
        '101.12.12
        TextBoxSP1.Text = sTextBoxSP1
        TextBoxSP2.Text = sTextBoxSP2
        TextBoxSP3.Text = sTextBoxSP3
        TextBoxSP4.Text = sTextBoxSP4
        '100.3.1
        lblMachTot.Text = MainForm.lblMachTot.Text
        '100.3.10
        lblUniSer.Text = iUniSer.ToString
        '102.1.22
        CheckBoxtTot2dec.Checked = bCheckBoxtTot2dec
        '102.3.16
        CheckBoxMix.Checked = bCheckBoxMix

        '105.6.26
        TextBoxBpath.Text = sTextBoxBpath
        '105.7.5
        'TextBoxSPARE1.Text = sTextBoxSPARE1
        TextBoxSPARE1.Text = sRemoteApath
        TextBoxSPARE2.Text = sTextBoxSPARE2
        '108.9.12
        'TextBoxSPARE3.Text = sTextBoxSPARE3
        'TextBoxSPARE4.Text = sTextBoxSPARE4
        TextBoxSPARE3.Text = MainForm.TimerMoveA.Interval
        TextBoxSPARE4.Text = MainForm.TimerFlag.Interval

        TextBoxSPARE5.Text = sTextBoxSPARE5
        CheckBoxSP4.Checked = bCheckBoxSP4
        CheckBoxSP5.Checked = bCheckBoxSP5
        CheckBoxSP6.Checked = bCheckBoxSP6
        CheckBoxSP7.Checked = bCheckBoxSP7
        CheckBoxSP8.Checked = bCheckBoxSP8
        CheckBoxSP9.Checked = bCheckBoxSP9
        '105.10.7
        CheckBoxSP10.Checked = bCheckBoxSP10
        CheckBoxSP11.Checked = bCheckBoxSP11
        CheckBoxSP12.Checked = bCheckBoxSP12
        CheckBoxSP13.Checked = bCheckBoxSP13
        CheckBoxSP14.Checked = bCheckBoxSP14
        CheckBoXSP15.Checked = bCheckBoxSP15
        '107.6.10
        CheckBoXSP16.Checked = bCheckBoxSP16
        CheckBoXSP17.Checked = bCheckBoxSP17
        CheckBoXSP18.Checked = bCheckBoxSP18
        CheckBoXSP19.Checked = bCheckBoxSP19
        CheckBoXSP20.Checked = bCheckBoxSP20

        '99.12.03
        'TextBoxMixerNo.Text = 1
        TextBoxIP.Text = sDataBaseIP
        DataGridView4.RowCount = 30

        '105.7.21 REF bRadioButton_NO2
        'TextBoxPrj.Text = sProject
        'TextBoxMixerNo.Text = sMixerNo
        TextBoxPrj.Text = sTextBoxPrj
        TextBoxPrj2.Text = sTextBoxPrj2
        TextBoxMixerNo.Text = sTextBoxMixerNo
        TextBoxMixerNo2.Text = sTextBoxMixerNo2
        TextBoxIP_R.Text = sTextBoxIP_R
        If bRadioButton_NO2 Then
            TextBoxPrj2.BackColor = Color.Pink
            TextBoxPrj.BackColor = Color.White
            TextBoxMixerNo2.BackColor = Color.Pink
            TextBoxMixerNo.BackColor = Color.White
        Else
            TextBoxPrj.BackColor = Color.Pink
            TextBoxPrj2.BackColor = Color.White
            TextBoxMixerNo.BackColor = Color.Pink
            TextBoxMixerNo2.BackColor = Color.White
        End If

        Dim i As Integer
        For i = 0 To DataGridView4.ColumnCount - 1
            DataGridView4.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        Call ReadOtherDes()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '106.6.10
        If Button2.BackColor = Color.Blue Then
            '107.6.10
            Call WriteAdminName()
            Button2.BackColor = Color.Pink
            Exit Sub
        End If


        '105.9.3
        sTextBoxIP_R = TextBoxIP_R.Text


        If IsNumeric(TextBoxMaxM3.Text) Then
            fCarMaxM3 = CSng(TextBoxMaxM3.Text)
            SaveSetting("JS", "CBC800", "fCarMaxM3", CStr(fCarMaxM3))
        Else
            MsgBox("[最大拌合方數]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
            TextBoxMaxM3.Text = fCarMaxM3
        End If

        If IsNumeric(TextBoxZeroMax.Text) Then
            fScaleZeroMax = CSng(TextBoxZeroMax.Text)
            SaveSetting("JS", "CBC800", "fScaleZeroMax", CStr(fScaleZeroMax))
        Else
            MsgBox("[校磅歸零最大值]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
            TextBoxZeroMax.Text = fScaleZeroMax
        End If

        If IsNumeric(TextBoxSpanMax.Text) Then
            fScaleSpanMax = CSng(TextBoxSpanMax.Text)
            SaveSetting("JS", "CBC800", "fScaleSpanMax", CStr(fScaleSpanMax))
        Else
            MsgBox("[校磅校準最大值]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
            TextBoxSpanMax.Text = fScaleSpanMax
        End If

        If IsNumeric(TextBoxSpanMin.Text) Then
            fScaleSpanMin = CSng(TextBoxSpanMin.Text)
            SaveSetting("JS", "CBC800", "fScaleSpanMin", CStr(fScaleSpanMin))
        Else
            MsgBox("[校磅校準最小值]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
            TextBoxSpanMin.Text = fScaleSpanMin
        End If

        sCompany = TextBoxCompany.Text
        SaveSetting("JS", "CBC800", "sCompany", CStr(sCompany))
        MainForm.Label2.Text = sCompany


        sDataBaseIP = TextBoxIP.Text
        SaveSetting("JS", "CBC800", "sDataBaseIP", CStr(sDataBaseIP))

        '105.7.21 REF bRadioButton_NO2 
        'sProject = Trim(TextBoxPrj.Text)
        'SaveSetting("JS", "CBC800", "sProject", Trim(sProject))
        'sMixerNo = TextBoxMixerNo.Text
        'SaveSetting("JS", "CBC800", "sMixerNo", CStr(sMixerNo))

        '100.3.1
        '104.10.19
        'SaveSetting("JS", "CBC800", "lblMachTot", lblMachTot.Text)
        If IsNumeric(lblMachTot.Text) Then
            SaveSetting("JS", "CBC800", "lblMachTot", lblMachTot.Text)
        Else
            MsgBox("[本機]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
        End If
        '104.10.19
        'SaveSetting("JS", "CBC800", "iUniSer", lblUniSer.Text)
        If IsNumeric(lblUniSer.Text) Then
            SaveSetting("JS", "CBC800", "iUniSer", lblUniSer.Text)
        Else
            MsgBox("[UniCar]輸入錯誤!", MsgBoxStyle.Critical, "輸入錯誤")
        End If

        '105.8.14
        'sTextBoxBpath = TextBoxBpath.Text
        sTextBoxSPARE1 = TextBoxSPARE1.Text
        '106.3.30
        'SaveSetting("JS", "CBC800", "sTextBoxBpath", CStr(sTextBoxBpath))
        SaveSetting("JS", "CBC800", "sTextBoxSPARE1", CStr(sTextBoxSPARE1))
        sTextBoxSPARE2 = TextBoxSPARE2.Text
        SaveSetting("JS", "CBC800", "sTextBoxSPARE2", CStr(sTextBoxSPARE2))
        sTextBoxSPARE3 = TextBoxSPARE3.Text
        SaveSetting("JS", "CBC800", "sTextBoxSPARE3", CStr(sTextBoxSPARE3))
        sTextBoxSPARE4 = TextBoxSPARE4.Text
        SaveSetting("JS", "CBC800", "sTextBoxSPARE4", CStr(sTextBoxSPARE4))
        sTextBoxSPARE5 = TextBoxSPARE5.Text
        SaveSetting("JS", "CBC800", "sTextBoxSPARE5", CStr(sTextBoxSPARE5))

        '102.5.11
        '112.3.25 bCheckBoxCBC7 改成列印拌合結束電流
        bCheckBoxCBC7 = CheckBoxCBC7.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxCBC7", CStr(bCheckBoxCBC7))
        '102.7.15
        bCheckBoxBPBsuben = CheckBoxBPBsubEn.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxBPBsuben", CStr(bCheckBoxBPBsuben))
        '102.11.23
        bCheckBoxCement = CheckBoxCement.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxCement", CStr(bCheckBoxCement))
        '102.12.11
        bCheckBoxBack = CheckBoxBack.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxBack", CStr(bCheckBoxBack))
        '103.3.11
        bCheckBoxImperial = CheckBoxImperial.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxImperial", CStr(bCheckBoxImperial))
        '103.6.7
        bCheckBoxNonAuto = CheckBoxNonAuto.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxNonAuto", CStr(bCheckBoxNonAuto))
        '103.4.25
        bCheckBoxRemain = CheckBoxRemain.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxRemain", CStr(bCheckBoxRemain))
        bCheckBoxSP1 = CheckBoxSP1.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP1", CStr(bCheckBoxSP1))
        bCheckBoxSP2 = CheckBoxSP2.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP2", CStr(bCheckBoxSP2))
        bCheckBoxSP3 = CheckBoxSP3.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP3", CStr(bCheckBoxSP3))
        bCheckBoxSP1 = CheckBoxSP1.Checked
        '10５.6.26
        bCheckBoxSP4 = CheckBoxSP4.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP4", CStr(bCheckBoxSP4))
        bCheckBoxSP5 = CheckBoxSP5.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP5", CStr(bCheckBoxSP5))
        bCheckBoxSP6 = CheckBoxSP6.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP6", CStr(bCheckBoxSP6))
        bCheckBoxSP7 = CheckBoxSP7.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP7", CStr(bCheckBoxSP7))
        bCheckBoxSP8 = CheckBoxSP8.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP8", CStr(bCheckBoxSP8))
        bCheckBoxSP9 = CheckBoxSP9.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP9", CStr(bCheckBoxSP9))
        '105.10.7
        bCheckBoxSP10 = CheckBoxSP10.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP10", CStr(bCheckBoxSP10))
        bCheckBoxSP11 = CheckBoxSP11.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP11", CStr(bCheckBoxSP11))
        bCheckBoxSP12 = CheckBoxSP12.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP12", CStr(bCheckBoxSP12))
        bCheckBoxSP13 = CheckBoxSP13.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP13", CStr(bCheckBoxSP13))
        bCheckBoxSP14 = CheckBoxSP14.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP14", CStr(bCheckBoxSP14))
        bCheckBoxSP15 = CheckBoXSP15.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP15", CStr(bCheckBoxSP15))
        '107.6.10
        bCheckBoxSP16 = CheckBoXSP16.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP16", CStr(bCheckBoxSP16))
        bCheckBoxSP17 = CheckBoXSP17.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP17", CStr(bCheckBoxSP17))
        bCheckBoxSP18 = CheckBoXSP18.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP18", CStr(bCheckBoxSP18))
        bCheckBoxSP19 = CheckBoXSP19.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP19", CStr(bCheckBoxSP19))
        bCheckBoxSP20 = CheckBoXSP20.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxSP20", CStr(bCheckBoxSP20))
        '103.3.13
        bImperial = bCheckBoxImperial

        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Other.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Other.des", OpenMode.Output)
        If file <> False Then
            For i = 1 To 30
                'If IsNumeric(DataGridView4.Rows(i - 1).Cells(2).Value) Then
                '    j = CInt(DataGridView4.Rows(i - 1).Cells(2).Value)
                'Else
                '    MsgBox("輸入錯誤!請輸入整數.", MsgBoxStyle.Critical, "輸入錯誤")
                '    DataGridView4.Rows(i - 1).Cells(2).Value = 1
                '    j = CInt(DataGridView4.Rows(i - 1).Cells(2).Value)
                'End If
                PrintLine(fileNum, DataGridView4.Rows(i - 1).Cells(0).Value.ToString)
                PrintLine(fileNum, DataGridView4.Rows(i - 1).Cells(1).Value.ToString)
                PrintLine(fileNum, DataGridView4.Rows(i - 1).Cells(2).Value.ToString)
                PrintLine(fileNum, DataGridView4.Rows(i - 1).Cells(3).Value.ToString)
                PrintLine(fileNum, DataGridView4.Rows(i - 1).Cells(4).Value.ToString)
            Next
        End If
        FileClose(fileNum)

        '102.5.14 for print report
        '112.3.25 bCheckBoxCBC7 改成列印拌合結束電流
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

        '107.7.34
        '107.7.9 still use Admin
        'Call WriteAdmin2()
        '105.7.21
        Call WriteAdmin()

        fileNum = FreeFile()
        FileOpen(fileNum, "C:\JS\OrigFile\MIXER.des", OpenMode.Output)
        i = 0
        '1
        PrintLine(fileNum, TextBoxPrj.Text)
        '2
        PrintLine(fileNum, TextBoxPrj2.Text)
        '3
        PrintLine(fileNum, TextBoxMixerNo.Text)
        '4
        PrintLine(fileNum, TextBoxMixerNo2.Text)
        FileClose(fileNum)

        '112.2.6
        sBigDataDataBaseIP = TextBoxSPARE5.Text
    End Sub



    Private Sub TabPage4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage4.Enter
        'Dim dtconfig As DataTable
        'dtconfig = db.GetDataTable("select PLC_Address, Alarm_Type, Alarm_Desc from Alarm_Setting order by PLC_Address")
        'DataGridView3.AutoGenerateColumns = False
        'DataGridView3.DataSource = dtconfig
        'DataGridView3.Columns(0).DataPropertyName = "PLC_Address"
        'DataGridView3.Columns(1).DataPropertyName = "Alarm_Type"
        'DataGridView3.Columns(2).DataPropertyName = "Alarm_Desc"
        'DataGridView3.Refresh()
        'DataGridView3.Columns(0).Selected = False
        '113.8.28
        'DataGridView3.RowCount = 105
        DataGridView3.RowCount = AlarmDvMax * 16
        Dim j As Integer
        For j = 0 To (AlarmDvMax * 16 - 1)
            DataGridView3.Rows(j).Cells(0).Value = AlarmName(j)
            DataGridView3.Rows(j).Cells(1).Value = AlarmType(j)
            DataGridView3.Rows(j).Cells(2).Value = AlarmDesc(j)
        Next

    End Sub




    Private Sub TabPage5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage5.Click
        CheckBox1.Checked = bCommFlag
        ComboBoxCOM.Text = sPlcCom
        '105.7.5
        ComboBoxDelay.Text = sPlcDelay
        '107.1.31
        CheckBox232.Checked = bCheckBox232
        ComboBox232.Text = sPlcCom2
        '107.9.12
        CheckBoxA.Checked = bCheckBoxA
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim i
        i = FrmMonit.ReadA2N(CInt(TextBoxD.Text), 40)
        LabelD.Text = TextBoxD.Text
        Timer1.Enabled = True
        For i = 0 To 19
            'DgvPLC.Rows(i). = i
            DgvPLC.Rows(i).Cells(0).Value = CInt(TextBoxD.Text) + i
            DgvPLC.Rows(i).Cells(1).Value = DV(CInt(TextBoxD.Text) + i)
            DgvPLC.Rows(i).Cells(2).Value = CInt(TextBoxD.Text) + i + 20
            DgvPLC.Rows(i).Cells(3).Value = DV(CInt(TextBoxD.Text) + i + 20)
        Next
        senddButton = 4
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim i
        DVW(CInt(TextBoxD.Text)) = CInt(TextBoxV.Text)
        i = FrmMonit.WriteA2N(CInt(TextBoxD.Text), 1)

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim i
        'i = FrmMonit.ReadA2N(CInt(LabelD.Text), 40)
        Timer1.Enabled = True
        For i = 0 To 19
            If senddButton = 5 Then
                DgvPLC.Rows(i).Cells(1).Value = DVW(CInt(LabelD.Text) + i)
                DgvPLC.Rows(i).Cells(3).Value = DVW(CInt(LabelD.Text) + i + 20)
            Else
                DgvPLC.Rows(i).Cells(1).Value = DV(CInt(LabelD.Text) + i)
                DgvPLC.Rows(i).Cells(3).Value = DV(CInt(LabelD.Text) + i + 20)
            End If
        Next
        Label9.Text = "ComCt: " & FrmMonit.COM_PLC.ErrorCount & " - " & FrmMonit.COM_PLC.PollCount
        Label10.Text = "R_Q: " & FrmMonit.COM_PLC.ReadDvStart(0) & " - " & FrmMonit.COM_PLC.ReadDvStart(1) & " - " & FrmMonit.COM_PLC.ReadDvStart(2) & " - " & FrmMonit.COM_PLC.ReadDvStart(3) _
                     & " - " & FrmMonit.COM_PLC.ReadDvStart(4) & " - " & FrmMonit.COM_PLC.ReadDvStart(5) & " - " & FrmMonit.COM_PLC.ReadDvStart(6) & " - " & FrmMonit.COM_PLC.ReadDvStart(7)
        LabelRA2.Text = "RA2:" & FrmMonit.COM_PLC.ReadDvStart(0) & "-" & FrmMonit.COM_PLC.ReadDvStart(1) & "-" & FrmMonit.COM_PLC.ReadDvStart(2) & "-" & FrmMonit.COM_PLC.ReadDvStart(3) & "-" & FrmMonit.COM_PLC.ReadDvStart(4) & "-"
    End Sub

    Private Sub TabPage5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Enter
        CheckBox1.Checked = bCommFlag
        ComboBoxCOM.Text = sPlcCom
        '105.7.5
        ComboBoxDelay.Text = sPlcDelay
        Label9.Text = "Com:" & FrmMonit.COM_PLC.ErrorCount & "-" & FrmMonit.COM_PLC.PollCount
        '107.1.31
        CheckBox232.Checked = bCheckBox232
        ComboBox232.Text = sPlcCom2
        '107.9.12
        CheckBoxA.Checked = bCheckBoxA

    End Sub

    Private Sub TabPage5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.GotFocus
        CheckBox1.Checked = bCommFlag
        ComboBoxCOM.Text = sPlcCom
        '105.7.5
        ComboBoxDelay.Text = sPlcDelay
    End Sub

    Private Sub TabPage5_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Leave
        Timer1.Enabled = False

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim i
        ' 104.8.17 
        'For i = 0 To 24
        '    DVW(CInt(TextBoxD.Text) + i) = CInt(TextBoxV.Text) + i
        'Next
        'i = FrmMonit.WriteA2N(CInt(TextBoxD.Text), 25)
        '104.8.20
        If TextBoxD.Text = "23" Then
            iSimTime = 0
            Timer2.Enabled = True
        Else
            For i = 0 To 31
                DVW(CInt(TextBoxD.Text) + i) = CInt(TextBoxV.Text) + i
            Next
            i = FrmMonit.WriteA2N(CInt(TextBoxD.Text), 32)
        End If
        senddButton = 5
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        If Button2.BackColor = Color.Blue Then
            Button2.BackColor = Color.Pink
            ButtonReadName.Visible = False
        Else
            Button2.BackColor = Color.Blue
            ButtonReadName.Visible = True
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            bCommFlag = True
            Label1Par.Text = FrmMonit.SerialPort1.BaudRate & "," & FrmMonit.SerialPort1.Parity & "," & FrmMonit.SerialPort1.DataBits & "," & FrmMonit.SerialPort1.StopBits
        Else
            bCommFlag = False
        End If
        SaveSetting("JS", "CBC800", "CommFlag", CStr(bCommFlag))

    End Sub

    Private Sub ComboBoxCOM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxCOM.SelectedIndexChanged
        sPlcCom = ComboBoxCOM.Text
        SaveSetting("JS", "CBC800", "PlcCom", sPlcCom)

    End Sub

    Private Sub CheckBoxYadon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxYadon.CheckedChanged
        '102.3.19
        'bCheckBoxYadon = CheckBoxYadon.Checked
        'SaveSetting("JS", "CBC800", "bCheckBoxYadon", CStr(bCheckBoxYadon))
    End Sub

    Private Sub DataGridView3_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellEndEdit
        'Dim index As Integer = e.RowIndex
        'Dim strsql As String
        'Dim s1, s2, s3 As String

        's1 = DataGridView3.Rows(index).Cells(1).Value.ToString
        's2 = DataGridView3.Rows(index).Cells(2).Value.ToString
        's3 = DataGridView3.Rows(index).Cells(0).Value.ToString
        'strsql = "Update Alarm_Setting set Alarm_Type =" & s1 & ", Alarm_Desc = '" & s2 & "' Where PLC_Address = '" & s3 & "'"
        'db.ExecuteCmd(strsql)


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim file
        Dim i As Integer

        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Alarm.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Alarm.des", OpenMode.Output)
        If file <> False Then
            For i = 1 To DataGridView3.RowCount
                PrintLine(fileNum, DataGridView3.Rows(i - 1).Cells(0).Value.ToString)
                PrintLine(fileNum, DataGridView3.Rows(i - 1).Cells(1).Value.ToString)
                PrintLine(fileNum, DataGridView3.Rows(i - 1).Cells(2).Value.ToString)
                AlarmType(i - 1) = CInt(DataGridView3.Rows(i - 1).Cells(1).Value.ToString)
                AlarmDesc(i - 1) = DataGridView3.Rows(i - 1).Cells(2).Value.ToString
            Next
        End If
        FileClose(fileNum)
    End Sub
    Public Sub ReadOtherDes()
        Dim i, j As Integer
        Dim Dreg As Integer
        'Dim dtconfig As DataTable
        'dtconfig = db.GetDataTable("select PLC_Address, Alarm_Type, Alarm_Desc from Alarm_Setting order by PLC_Address")
        'For i = 0 To 63
        '    AlarmType(i) = dtconfig.Rows(i).Item(1).ToString
        '    AlarmType(i) = dtconfig.Rows(i).Item(1).ToString
        '    AlarmDesc(i) = dtconfig.Rows(i).Item(2).ToString
        'Next

        Dim file
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Other.des")
        If file = False Then
            For i = 711 To 740
                Dreg = i
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\Other.des", "D" & Dreg & vbCrLf, True, System.Text.Encoding.Default)
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\Other.des", "Des" & Dreg & vbCrLf, True, System.Text.Encoding.Default)
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\Other.des", "10" & vbCrLf, True, System.Text.Encoding.Default)
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\Other.des", "Unit" & Dreg & vbCrLf, True, System.Text.Encoding.Default)
                My.Computer.FileSystem.WriteAllText(sSMdbName & "\Other.des", "Note" & Dreg & vbCrLf, True, System.Text.Encoding.Default)
            Next
        End If
        Dim s As String = "0.0"
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Other.des", OpenMode.Input)
        If Not EOF(fileNum) Then
            For j = 0 To 29
                DataGridView4.Rows(j).Cells(0).Value = LineInput(fileNum)
                DataGridView4.Rows(j).Cells(1).Value = LineInput(fileNum)
                DataGridView4.Rows(j).Cells(2).Value = LineInput(fileNum)
                DataGridView4.Rows(j).Cells(3).Value = LineInput(fileNum)
                DataGridView4.Rows(j).Cells(4).Value = LineInput(fileNum)
            Next
            FileClose(fileNum)
        End If

    End Sub
     Private Sub TabPage4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage4.Click

    End Sub

    Private Sub DataGridView4_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellEndEdit
        If e.ColumnIndex = 2 Then
            If Not IsNumeric(DataGridView4.CurrentCell.Value) Then
                MsgBox("輸入錯誤!請輸入整數.", MsgBoxStyle.Critical, "輸入錯誤")
                DataGridView4.CurrentCell.Value = temp
            End If
        End If
    End Sub

    Private Sub DataGridView4_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellEnter
        temp = DataGridView4.CurrentCell.Value
    End Sub



    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        My.Computer.FileSystem.WriteAllText(sSMdbName & "\note.txt", TextBox2.Text, False, System.Text.Encoding.Default)

    End Sub

 

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox2.Text = My.Computer.FileSystem.ReadAllText(sSMdbName & "\note.txt", System.Text.Encoding.Default)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim i
        For i = 1 To 20
            'DgvPLC.Rows(i). = i
            DgvPLC.Rows(i - 1).Cells(0).Value = i
            DgvPLC.Rows(i - 1).Cells(1).Value = FrmMonit.TrendSec(i) & ":" & FrmMonit.Trend(i)
            DgvPLC.Rows(i - 1).Cells(2).Value = i + 20
            DgvPLC.Rows(i - 1).Cells(3).Value = FrmMonit.TrendSec(i + 20) & ":" & FrmMonit.Trend(i + 20)
        Next

    End Sub


    Private Sub CheckBoxS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '100.6.29
        bCheckBoxS = CheckBoxS.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxS", CStr(bCheckBoxS))

    End Sub
    Private Sub CheckBoxS_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxS.CheckedChanged
        '100.8.12
        bCheckBoxS = CheckBoxS.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxS", CStr(bCheckBoxS))

    End Sub
    Private Sub CheckBoxG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '100.6.29
        bCheckBoxG = CheckBoxG.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxG", CStr(bCheckBoxG))

    End Sub
    Private Sub CheckBoxG_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxG.CheckedChanged
        '100.8.12
        bCheckBoxG = CheckBoxG.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxG", CStr(bCheckBoxG))


    End Sub
    Private Sub CheckBoxW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '100.6.29
        bCheckBoxW = CheckBoxW.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxW", CStr(bCheckBoxW))
    End Sub

    Private Sub CheckBoxW_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxW.CheckedChanged
        '100.8.12
        bCheckBoxW = CheckBoxW.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxW", CStr(bCheckBoxW))

    End Sub
    Private Sub CheckBoxC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '100.6.29
        bCheckBoxC = CheckBoxC.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC", CStr(bCheckBoxC))

    End Sub
    Private Sub CheckBoxC_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC.CheckedChanged
        '100.6.29
        bCheckBoxC = CheckBoxC.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC", CStr(bCheckBoxC))

    End Sub
    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click
        '100.8.12
        bCheckBoxC = CheckBoxC.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC", CStr(bCheckBoxC))
        ReadAltDes()

    End Sub

    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub



    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub TabPage1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage1.Enter
        CheckBoxS.Checked = bCheckBoxS
        CheckBoxG.Checked = bCheckBoxG
        CheckBoxW.Checked = bCheckBoxW
        CheckBoxC.Checked = bCheckBoxC
        '101.4.5
        CheckBoxS4.Checked = bCheckBoxS4
    End Sub



    Private Sub CheckBoxsingle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxsingle.CheckedChanged
        '102.3.19 move to save button
        'bCheckBoxSingle = CheckBoxsingle.Checked
        'SaveSetting("JS", "CBC800", "bCheckBoxSingle", CStr(bCheckBoxSingle))

    End Sub

    Private Sub CheckBoxtTest_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxtTest.CheckedChanged
        '102.3.19 move to save button
        'bCheckBoxtTest = CheckBoxtTest.Checked
        'SaveSetting("JS", "CBC800", "bCheckBoxtTest", CStr(bCheckBoxtTest))
    End Sub

    Private Sub CheckBoxS4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxS4.CheckedChanged
        '100.6.29
        bCheckBoxS4 = CheckBoxS4.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxS4", CStr(bCheckBoxS4))
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBoxZeroMax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxZeroMax.TextChanged

    End Sub

    Private Sub TextBoxSpanMax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxSpanMax.TextChanged

    End Sub

    Private Sub TextBoxSpanMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxSpanMin.TextChanged

    End Sub

    Private Sub ButtonOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOut.Click
        '102.3.20
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Other2.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Other2.des", OpenMode.Output)
        i = 0
        If file <> False Then
            '1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTest.Text)
            PrintLine(fileNum, CheckBoxtTest.Checked.ToString)

            '2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxYadon.Text)
            PrintLine(fileNum, CheckBoxYadon.Checked.ToString)

            '3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxWaitTank.Text)
            PrintLine(fileNum, CheckBoxWaitTank.Checked.ToString)

            '4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxABPB.Text)
            PrintLine(fileNum, CheckBoxABPB.Checked.ToString)

            '5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxMix.Text)
            PrintLine(fileNum, CheckBoxMix.Checked.ToString)

            '6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label4.Text)
            PrintLine(fileNum, TextBoxMaxM3.Text)

            '7
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label5.Text)
            PrintLine(fileNum, TextBoxCompany.Text)

            '8
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj.Text)

            '9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxsingle.Text)
            PrintLine(fileNum, CheckBoxsingle.Checked.ToString)

            '10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label20.Text)
            PrintLine(fileNum, lblUniSer.Text)

            '11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label19.Text)
            PrintLine(fileNum, lblMachTot.Text)

            '12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTot2dec.Text)
            PrintLine(fileNum, CheckBoxtTot2dec.Checked.ToString)

            '13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label6.Text)
            PrintLine(fileNum, TextBoxZeroMax.Text)

            '14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label7.Text)
            PrintLine(fileNum, TextBoxSpanMax.Text)

            '15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label8.Text)
            PrintLine(fileNum, TextBoxSpanMin.Text)

            '16
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo.Text)

            '17
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label12.Text)
            PrintLine(fileNum, TextBoxIP.Text)

            '18
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label21.Text)
            PrintLine(fileNum, TextBoxSP1.Text)

            '19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label22.Text)
            PrintLine(fileNum, TextBoxSP2.Text)

            '20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label23.Text)
            PrintLine(fileNum, TextBoxSP3.Text)

            '21
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label24.Text)
            PrintLine(fileNum, TextBoxSP4.Text)

            '22
            '102.5.11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCBC7.Text)
            PrintLine(fileNum, CheckBoxCBC7.Checked.ToString)

            '23
            '102.7.15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBPBsubEn.Text)
            PrintLine(fileNum, CheckBoxBPBsubEn.Checked.ToString)

            '24
            '102.7.15 列印水泥註記
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCement.Text)
            PrintLine(fileNum, CheckBoxCement.Checked.ToString)

            '25
            '102.12.11 廠別切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBack.Text)
            PrintLine(fileNum, CheckBoxBack.Checked.ToString)

            '26
            '103.6.9 英制切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxImperial.Text)
            PrintLine(fileNum, CheckBoxImperial.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxImperial.Checked.ToString)

            '27
            '103.6.9 非自動上傳
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxNonAuto.Text)
            PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)

            '28  104.9.4 COPY FROM GWAN
            '104.4.25 REMAIN VISIBLE FOR PRINT
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxRemain.Text)
            PrintLine(fileNum, CheckBoxRemain.Checked.ToString)

            '29
            '104.4.25 SPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP1.Text)
            PrintLine(fileNum, CheckBoxSP1.Checked.ToString)

            '30
            '104.4.25 SPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP2.Text)
            PrintLine(fileNum, CheckBoxSP2.Checked.ToString)

            '31
            '104.4.25 SPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP3.Text)
            PrintLine(fileNum, CheckBoxSP3.Checked.ToString)

            '32
            '106.6.26 SPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP4.Text)
            PrintLine(fileNum, CheckBoxSP4.Checked.ToString)

            '33
            '106.6.26 SPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP5.Text)
            PrintLine(fileNum, CheckBoxSP5.Checked.ToString)

            '34
            '106.6.26 SPARE6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP6.Text)
            PrintLine(fileNum, CheckBoxSP6.Checked.ToString)

            '35
            '106.6.26 SPARE7
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP7.Text)
            PrintLine(fileNum, CheckBoxSP7.Checked.ToString)

            '36
            '106.6.26 SPARE8
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP8.Text)
            PrintLine(fileNum, CheckBoxSP8.Checked.ToString)

            '37
            '106.6.26 SPARE9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP9.Text)
            PrintLine(fileNum, CheckBoxSP9.Checked.ToString)

            '38
            '106.6.26 TextBoxBpath
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label26.Text)
            PrintLine(fileNum, TextBoxBpath.Text)

            '39
            '106.6.26 TextBoxSPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label27.Text)
            PrintLine(fileNum, TextBoxSPARE1.Text)

            '40
            '106.6.26 TextBoxSPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label28.Text)
            PrintLine(fileNum, TextBoxSPARE2.Text)

            '41
            '105.6.26 TextBoxSPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label29.Text)
            PrintLine(fileNum, TextBoxSPARE3.Text)

            '42
            '105.6.26 TextBoxSPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label30.Text)
            PrintLine(fileNum, TextBoxSPARE4.Text)

            '43
            '105.6.26 TextBoxSPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label31.Text)
            PrintLine(fileNum, TextBoxSPARE5.Text)

            '44
            '105.7.19 遠端電腦IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, LabelIP_R.Text)
            PrintLine(fileNum, TextBoxIP_R.Text)

            '45
            '105.7.19 專案編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj2.Text)

            '46
            '105.7.19 拌合機編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo2.Text)

        End If
        FileClose(fileNum)

    End Sub

    Private Sub ButtonIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonIn.Click
        '102.3.20
        'Dim i, j As Integer

        Dim file
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Other2.des")
        If file = False Then
            MsgBox("找不到檔案", MsgBoxStyle.Critical, "找不到Other2.des!")
            Exit Sub
        End If
        Dim fileNum
        Dim s As String
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Other2.des", OpenMode.Input)

        '104.4.25
        On Error Resume Next

        '1
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxtTest.Checked = CBool(s)

        '2
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxYadon.Checked = CBool(s)

        '3
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxWaitTank.Checked = CBool(s)

        '4
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxABPB.Checked = CBool(s)

        '5
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxMix.Checked = CBool(s)

        '6
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxMaxM3.Text = s

        '7
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxCompany.Text = s

        '8
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxPrj.Text = s

        '9
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxsingle.Checked = CBool(s)

        '10
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        lblUniSer.Text = s

        '11
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        lblMachTot.Text = s

        '12
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxtTot2dec.Checked = CBool(s)

        '13
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxZeroMax.Text = s

        '14
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSpanMax.Text = s

        '15
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSpanMin.Text = s

        '16
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxMixerNo.Text = s

        '17
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxIP.Text = s

        '18
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSP1.Text = s

        '19
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSP2.Text = s

        '20
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSP3.Text = s

        '21
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSP4.Text = s

        '22  102.5.11
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxCBC7.Checked = CBool(s)

        '23  102.7.15
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxBPBsubEn.Checked = CBool(s)

        '24  102.11.23
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxCement.Checked = CBool(s)

        '25  102.12.11 廠別切換
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxBack.Checked = CBool(s)

        '26  103.6.9 英制切換
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxImperial.Checked = CBool(s)

        '27  103.6.9 非自動上傳
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxNonAuto.Checked = CBool(s)

        '28  104.4.25 REMAIN
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxRemain.Checked = CBool(s)

        '29  104.4.25 SPARE1
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP1.Checked = CBool(s)

        '30  104.4.25 SPARE2
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP2.Checked = CBool(s)

        '31  104.4.25 SPARE3
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP3.Checked = CBool(s)

        '32  105.6.26 SPARE4
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP4.Checked = CBool(s)

        '33  105.6.26 SPARE5
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP5.Checked = CBool(s)

        '34  105.6.26 SPARE6
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP6.Checked = CBool(s)

        '35  105.6.26 SPARE7
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP7.Checked = CBool(s)

        '36  105.6.26 SPARE8
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP8.Checked = CBool(s)

        '37  105.6.26 SPARE9
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP9.Checked = CBool(s)

        '38  105.6.26 BPATH
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxBpath.Text = s

        '39  105.6.26 SPARE1
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSPARE1.Text = s

        '40  105.6.26 SPARE2
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSPARE2.Text = s

        '41  105.6.26 SPARE3
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSPARE3.Text = s

        '42  105.6.26 SPARE4
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSPARE4.Text = s

        '43  105.6.26 SPARE5
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxSPARE5.Text = s

        '44  105.7.19 遠端電腦IP
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxIP_R.Text = s

        '45  105.7.19 TextBoxPrj2
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxPrj2.Text = s

        '46  105.7.19 TextBoxPrj2
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        TextBoxMixerNo2.Text = s

        FileClose(fileNum)
    End Sub

    Public Sub ReadAltDes()
        '102.10.19
        Dim file
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Alt.des")
        If file = False Then
            My.Computer.FileSystem.WriteAllText(sSMdbName & "\Alt.des", "Alt" & vbCrLf, True, System.Text.Encoding.Default)
        End If
        Dim s As String = "0.0"
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\alt.des", OpenMode.Input)
        TextBox1.Text = ""
        While Not EOF(fileNum)
            TextBox3.Text &= LineInput(fileNum) & vbCrLf
        End While
        FileClose(fileNum)

        '102.11.23
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Cement.des")
        If file = False Then
            My.Computer.FileSystem.WriteAllText(sSMdbName & "\Cement.des", "C1" & vbCrLf, True, System.Text.Encoding.Default)
        End If
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Cement.des", OpenMode.Input)
        TextBox1.Text = ""
        While Not EOF(fileNum)
            TextBox4.Text &= LineInput(fileNum) & vbCrLf
        End While
        FileClose(fileNum)
    End Sub

    Private Sub CheckBoxCBC7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCBC7.CheckedChanged

    End Sub

    Private Sub CheckBoxABPB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxABPB.CheckedChanged

    End Sub


    Private Sub CheckBoxBack_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxBack.CheckedChanged

    End Sub

    Private Sub CheckBoxCement_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCement.CheckedChanged

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim i
        Dim generator As New Random
        iSimTime += 1
        DVW(23) = iSimTime
        If iSimTime < 50 Then
            DVW(25) = generator.Next(1, 15) + DVW(25) - 4
        Else
            DVW(25) = DVW(25) - generator.Next(1, 10) + 4
        End If
        If DVW(25) <= 0 Then DVW(25) = 0
        i = FrmMonit.WriteA2N(23, 3)
        If iSimTime > 150 Or DV(23) < 1 Then
            DVW(23) = 0
            DVW(25) = 0
            Timer2.Enabled = False
        End If
    End Sub

    Private Sub CheckBoxC1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC1.CheckedChanged
        '104.9.24
        bCheckBoxC1 = CheckBoxC1.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC1", CStr(bCheckBoxC1))
        '110.10.5
        Call WriteCemWarehouse()

    End Sub

    Private Sub CheckBoxC2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC2.CheckedChanged
        '104.9.24
        bCheckBoxC2 = CheckBoxC2.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC2", CStr(bCheckBoxC2))
        '110.10.5
        Call WriteCemWarehouse()
    End Sub

    Private Sub CheckBoxC3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC3.CheckedChanged
        '104.9.24
        bCheckBoxC3 = CheckBoxC3.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC3", CStr(bCheckBoxC3))
        '110.10.5
        Call WriteCemWarehouse()

    End Sub

    Private Sub CheckBoxC4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC4.CheckedChanged
        '104.9.24
        bCheckBoxC4 = CheckBoxC4.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC4", CStr(bCheckBoxC4))
        '110.10.5
        Call WriteCemWarehouse()

    End Sub

    Private Sub CheckBoxC5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC5.CheckedChanged
        '104.9.24
        bCheckBoxC5 = CheckBoxC5.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC5", CStr(bCheckBoxC5))
        '110.10.5
        Call WriteCemWarehouse()

    End Sub

    Private Sub CheckBoxC6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxC6.CheckedChanged
        '104.9.24
        bCheckBoxC6 = CheckBoxC6.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxC6", CStr(bCheckBoxC6))
        '110.10.5
        Call WriteCemWarehouse()

    End Sub

    Private Sub ComboBoxDelay_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxDelay.SelectedIndexChanged
        '105.7.5
        sPlcDelay = ComboBoxDelay.Text
        SaveSetting("JS", "CBC800", "PlcDelay", sPlcDelay)
    End Sub

    Public Sub WriteAdmin2()
        '107.7.9 still use Admin
        '107.7.4
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Admin2.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Admin2.des", OpenMode.Output)
        i = 0
        If file <> False Then
            '1  遠端電腦IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, sTextBoxIP_R)

            '2  存放磁碟
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, sTextBoxBpath)
            PrintLine(fileNum, TextBoxBpath.Text)

            '3  遠端存放磁碟
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, sTextBoxSPARE1)
            PrintLine(fileNum, TextBoxSPARE1.Text)
        End If
        FileClose(fileNum)
    End Sub

    Public Sub WriteAdmin()
        '107.6.27
        '   just save value , no Name
        '105.7.21
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Admin.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Admin.des", OpenMode.Output)
        i = 0
        If file <> False Then
            '1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxtTest.Text)
            PrintLine(fileNum, CheckBoxtTest.Checked.ToString)

            '2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxYadon.Text)
            PrintLine(fileNum, CheckBoxYadon.Checked.ToString)

            '3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxWaitTank.Text)
            PrintLine(fileNum, CheckBoxWaitTank.Checked.ToString)

            '4  Ｂ配比查詢
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxABPB.Text)
            PrintLine(fileNum, CheckBoxABPB.Checked.ToString)

            '5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxMix.Text)
            PrintLine(fileNum, CheckBoxMix.Checked.ToString)

            '6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label4.Text)
            PrintLine(fileNum, TextBoxMaxM3.Text)

            '7  公司名稱
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label5.Text)
            PrintLine(fileNum, TextBoxCompany.Text)

            '8  專案編號
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj.Text)

            '9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxsingle.Text)
            PrintLine(fileNum, CheckBoxsingle.Checked.ToString)

            '10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label20.Text)
            PrintLine(fileNum, lblUniSer.Text)

            '11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label19.Text)
            PrintLine(fileNum, lblMachTot.Text)

            '12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxtTot2dec.Text)
            PrintLine(fileNum, CheckBoxtTot2dec.Checked.ToString)

            '13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label6.Text)
            PrintLine(fileNum, TextBoxZeroMax.Text)

            '14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label7.Text)
            PrintLine(fileNum, TextBoxSpanMax.Text)

            '15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label8.Text)
            PrintLine(fileNum, TextBoxSpanMin.Text)

            '16
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo.Text)

            '17
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label12.Text)
            PrintLine(fileNum, TextBoxIP.Text)

            '18
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label21.Text)
            PrintLine(fileNum, TextBoxSP1.Text)

            '19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label22.Text)
            PrintLine(fileNum, TextBoxSP2.Text)

            '20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label23.Text)
            PrintLine(fileNum, TextBoxSP3.Text)

            '21
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label24.Text)
            PrintLine(fileNum, TextBoxSP4.Text)

            '22
            '102.5.11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxCBC7.Text)
            PrintLine(fileNum, CheckBoxCBC7.Checked.ToString)

            '23
            '102.7.15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxBPBsubEn.Text)
            PrintLine(fileNum, CheckBoxBPBsubEn.Checked.ToString)

            '24
            '102.7.15 列印水泥註記
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxCement.Text)
            PrintLine(fileNum, CheckBoxCement.Checked.ToString)

            '25
            '102.12.11 廠別切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBack.Text)
            PrintLine(fileNum, CheckBoxBack.Checked.ToString)

            '26
            '103.6.9 英制切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxImperial.Text)
            PrintLine(fileNum, CheckBoxImperial.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxImperial.Checked.ToString)

            '27
            '103.6.9 非自動上傳
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxNonAuto.Text)
            PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)

            '28  104.9.4 COPY FROM GWAN
            '104.4.25 REMAIN VISIBLE FOR PRINT
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxRemain.Text)
            PrintLine(fileNum, CheckBoxRemain.Checked.ToString)

            '29 A1-15
            '104.4.25 SPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP1.Text)
            PrintLine(fileNum, CheckBoxSP1.Checked.ToString)

            '30 日月報表寬格式
            '104.4.25 SPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP2.Text)
            PrintLine(fileNum, CheckBoxSP2.Checked.ToString)

            '31 B配存檔
            '104.4.25 SPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP3.Text)
            PrintLine(fileNum, CheckBoxSP3.Checked.ToString)

            '32 資料自動補傳時訊
            '106.6.26 SPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP4.Text)
            PrintLine(fileNum, CheckBoxSP4.Checked.ToString)

            '33 可選遠端
            '106.6.26 SPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP5.Text)
            PrintLine(fileNum, CheckBoxSP5.Checked.ToString)

            '34 環泥烏日
            '106.6.26 SPARE6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP6.Text)
            PrintLine(fileNum, CheckBoxSP6.Checked.ToString)

            '35 參數存至遠端
            '106.6.26 SPARE7
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP7.Text)
            PrintLine(fileNum, CheckBoxSP7.Checked.ToString)

            '36 校磅密碼紀錄
            '106.6.26 SPARE8
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP8.Text)
            PrintLine(fileNum, CheckBoxSP8.Checked.ToString)

            '37 配比密碼紀錄
            '106.6.26 SPARE9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP9.Text)
            PrintLine(fileNum, CheckBoxSP9.Checked.ToString)

            '38 B存放磁碟(\\192.)
            '106.6.26 TextBoxBpath
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label26.Text)
            PrintLine(fileNum, TextBoxBpath.Text)

            '39 遠端A存放磁碟
            '106.6.26 TextBoxSPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label27.Text)
            PrintLine(fileNum, TextBoxSPARE1.Text)

            '40 SQL路徑
            '106.6.26 TextBoxSPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label28.Text)
            PrintLine(fileNum, TextBoxSPARE2.Text)

            '41 SPARE3
            '105.6.26 TextBoxSPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label29.Text)
            PrintLine(fileNum, TextBoxSPARE3.Text)

            '42 SPARE4
            '105.6.26 TextBoxSPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label30.Text)
            PrintLine(fileNum, TextBoxSPARE4.Text)

            '43 SPARE5
            '105.6.26 TextBoxSPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label31.Text)
            PrintLine(fileNum, TextBoxSPARE5.Text)

            '44 遠端電腦IP
            '105.7.19 遠端電腦IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, LabelIP_R.Text)
            PrintLine(fileNum, TextBoxIP_R.Text)

            '45 專案編號
            '105.7.19 專案編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj2.Text)

            '46 拌合機編號
            '105.7.19 拌合機編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo2.Text)

            '47 列印密碼調整
            '106.10.7 SPARE10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP10.Text)
            PrintLine(fileNum, CheckBoxSP10.Checked.ToString)

            '48 可修改首盤時間
            '106.10.7 SPARE11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP11.Text)
            PrintLine(fileNum, CheckBoxSP11.Checked.ToString)

            '49 配件庫存管理
            '106.10.7 SPARE12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP12.Text)
            PrintLine(fileNum, CheckBoxSP12.Checked.ToString)

            '50 聯佶空單
            '106.10.7 SPARE13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP13.Text)
            PrintLine(fileNum, CheckBoxSP13.Checked.ToString)

            '51 不回傳R1R2R3R5
            '106.10.7 SPARE14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoxSP14.Text)
            PrintLine(fileNum, CheckBoxSP14.Checked.ToString)

            '52 備用15
            '106.10.7 SPARE15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP15.Text)
            PrintLine(fileNum, CheckBoXSP15.Checked.ToString)

            '53 備用16
            '107.6.10 SPARE16   109.4.2 bPbSync 配比連線
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP16.Text)
            PrintLine(fileNum, CheckBoXSP16.Checked.ToString)

            '54 備用17
            '107.6.10 SPARE17 109.12.19 配比15碼
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP17.Text)
            PrintLine(fileNum, CheckBoXSP17.Checked.ToString)

            '55 備用18
            '107.6.10 SPARE18   
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP18.Text)
            PrintLine(fileNum, CheckBoXSP18.Checked.ToString)

            '56 備用19
            '107.6.10 SPARE19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP19.Text)
            PrintLine(fileNum, CheckBoXSP19.Checked.ToString)

            '57 備用20
            '107.6.10 SPARE20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            'PrintLine(fileNum, CheckBoXSP20.Text)
            PrintLine(fileNum, CheckBoXSP20.Checked.ToString)

        End If
        FileClose(fileNum)
    End Sub
    Public Sub WriteAdmin_Old()
        '105.7.21
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Admin.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\Admin.des", OpenMode.Output)
        i = 0
        If file <> False Then
            '1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTest.Text)
            PrintLine(fileNum, CheckBoxtTest.Checked.ToString)

            '2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxYadon.Text)
            PrintLine(fileNum, CheckBoxYadon.Checked.ToString)

            '3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxWaitTank.Text)
            PrintLine(fileNum, CheckBoxWaitTank.Checked.ToString)

            '4  Ｂ配比查詢
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxABPB.Text)
            PrintLine(fileNum, CheckBoxABPB.Checked.ToString)

            '5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxMix.Text)
            PrintLine(fileNum, CheckBoxMix.Checked.ToString)

            '6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label4.Text)
            PrintLine(fileNum, TextBoxMaxM3.Text)

            '7  公司名稱
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label5.Text)
            PrintLine(fileNum, TextBoxCompany.Text)

            '8  專案編號
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj.Text)

            '9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxsingle.Text)
            PrintLine(fileNum, CheckBoxsingle.Checked.ToString)

            '10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label20.Text)
            PrintLine(fileNum, lblUniSer.Text)

            '11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label19.Text)
            PrintLine(fileNum, lblMachTot.Text)

            '12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTot2dec.Text)
            PrintLine(fileNum, CheckBoxtTot2dec.Checked.ToString)

            '13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label6.Text)
            PrintLine(fileNum, TextBoxZeroMax.Text)

            '14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label7.Text)
            PrintLine(fileNum, TextBoxSpanMax.Text)

            '15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label8.Text)
            PrintLine(fileNum, TextBoxSpanMin.Text)

            '16
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo.Text)

            '17
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label12.Text)
            PrintLine(fileNum, TextBoxIP.Text)

            '18
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label21.Text)
            PrintLine(fileNum, TextBoxSP1.Text)

            '19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label22.Text)
            PrintLine(fileNum, TextBoxSP2.Text)

            '20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label23.Text)
            PrintLine(fileNum, TextBoxSP3.Text)

            '21
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label24.Text)
            PrintLine(fileNum, TextBoxSP4.Text)

            '22
            '102.5.11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCBC7.Text)
            PrintLine(fileNum, CheckBoxCBC7.Checked.ToString)

            '23
            '102.7.15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBPBsubEn.Text)
            PrintLine(fileNum, CheckBoxBPBsubEn.Checked.ToString)

            '24
            '102.7.15 列印水泥註記
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCement.Text)
            PrintLine(fileNum, CheckBoxCement.Checked.ToString)

            '25
            '102.12.11 廠別切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBack.Text)
            PrintLine(fileNum, CheckBoxBack.Checked.ToString)

            '26
            '103.6.9 英制切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxImperial.Text)
            PrintLine(fileNum, CheckBoxImperial.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxImperial.Checked.ToString)

            '27
            '103.6.9 非自動上傳
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxNonAuto.Text)
            PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)

            '28  104.9.4 COPY FROM GWAN
            '104.4.25 REMAIN VISIBLE FOR PRINT
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxRemain.Text)
            PrintLine(fileNum, CheckBoxRemain.Checked.ToString)

            '29 A1-15
            '104.4.25 SPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP1.Text)
            PrintLine(fileNum, CheckBoxSP1.Checked.ToString)

            '30 日月報表寬格式
            '104.4.25 SPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP2.Text)
            PrintLine(fileNum, CheckBoxSP2.Checked.ToString)

            '31 B配存檔
            '104.4.25 SPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP3.Text)
            PrintLine(fileNum, CheckBoxSP3.Checked.ToString)

            '32 資料自動補傳時訊
            '106.6.26 SPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP4.Text)
            PrintLine(fileNum, CheckBoxSP4.Checked.ToString)

            '33 可選遠端
            '106.6.26 SPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP5.Text)
            PrintLine(fileNum, CheckBoxSP5.Checked.ToString)

            '34 環泥烏日
            '106.6.26 SPARE6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP6.Text)
            PrintLine(fileNum, CheckBoxSP6.Checked.ToString)

            '35 參數存至遠端
            '106.6.26 SPARE7
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP7.Text)
            PrintLine(fileNum, CheckBoxSP7.Checked.ToString)

            '36 校磅密碼紀錄
            '106.6.26 SPARE8
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP8.Text)
            PrintLine(fileNum, CheckBoxSP8.Checked.ToString)

            '37 配比密碼紀錄
            '106.6.26 SPARE9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP9.Text)
            PrintLine(fileNum, CheckBoxSP9.Checked.ToString)

            '38 B存放磁碟(\\192.)
            '106.6.26 TextBoxBpath
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label26.Text)
            PrintLine(fileNum, TextBoxBpath.Text)

            '39 遠端A存放磁碟
            '106.6.26 TextBoxSPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label27.Text)
            PrintLine(fileNum, TextBoxSPARE1.Text)

            '40 SQL路徑
            '106.6.26 TextBoxSPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label28.Text)
            PrintLine(fileNum, TextBoxSPARE2.Text)

            '41 SPARE3
            '105.6.26 TextBoxSPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label29.Text)
            PrintLine(fileNum, TextBoxSPARE3.Text)

            '42 SPARE4
            '105.6.26 TextBoxSPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label30.Text)
            PrintLine(fileNum, TextBoxSPARE4.Text)

            '43 SPARE5
            '105.6.26 TextBoxSPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label31.Text)
            PrintLine(fileNum, TextBoxSPARE5.Text)

            '44 遠端電腦IP
            '105.7.19 遠端電腦IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, LabelIP_R.Text)
            PrintLine(fileNum, TextBoxIP_R.Text)

            '45 專案編號
            '105.7.19 專案編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj2.Text)

            '46 拌合機編號
            '105.7.19 拌合機編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo2.Text)

            '47 列印密碼調整
            '106.10.7 SPARE10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP10.Text)
            PrintLine(fileNum, CheckBoxSP10.Checked.ToString)

            '48 可修改首盤時間
            '106.10.7 SPARE11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP11.Text)
            PrintLine(fileNum, CheckBoxSP11.Checked.ToString)

            '49 配件庫存管理
            '106.10.7 SPARE12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP12.Text)
            PrintLine(fileNum, CheckBoxSP12.Checked.ToString)

            '50 聯佶空單
            '106.10.7 SPARE13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP13.Text)
            PrintLine(fileNum, CheckBoxSP13.Checked.ToString)

            '51 不回傳R1R2R3R5
            '106.10.7 SPARE14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP14.Text)
            PrintLine(fileNum, CheckBoxSP14.Checked.ToString)

            '52 備用15
            '106.10.7 SPARE15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP15.Text)
            PrintLine(fileNum, CheckBoXSP15.Checked.ToString)

            '53 備用16
            '107.6.10 SPARE16
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP16.Text)
            PrintLine(fileNum, CheckBoXSP16.Checked.ToString)

            '54 備用17
            '107.6.10 SPARE17
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP17.Text)
            PrintLine(fileNum, CheckBoXSP17.Checked.ToString)

            '55 備用18
            '107.6.10 SPARE18
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP18.Text)
            PrintLine(fileNum, CheckBoXSP18.Checked.ToString)

            '56 備用19
            '107.6.10 SPARE19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP19.Text)
            PrintLine(fileNum, CheckBoXSP19.Checked.ToString)

            '57 備用20
            '107.6.10 SPARE20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP20.Text)
            PrintLine(fileNum, CheckBoXSP20.Checked.ToString)

        End If
        FileClose(fileNum)
    End Sub
    Public Sub ReadAdminName()
        '107.6.10 ref. ReadAdmin
        Dim file
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        'file = My.Computer.FileSystem.FileExists(sSMdbName & "\AdminName.des")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\Admin.des")
        If file = False Then
            '107.6.16
            'My.Computer.FileSystem.CopyFile("\JS\OrigFile\AdminName.des", sSMdbName & "\Admin.des")
            Exit Sub
            MsgBox("Admin.des 找不到Admin.des!!", MsgBoxStyle.Critical, "找不到Admin.des!")
        End If
        Dim fileNum
        Dim s As String
        fileNum = FreeFile()
        '107.6.27
        FileOpen(fileNum, sSMdbName & "\AdminName.des", OpenMode.Input)
        'FileOpen(fileNum, sSMdbName & "\Admin.des", OpenMode.Input)

        'On Error Resume Next
        On Error GoTo Errorhandler

        '1  1280*960
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxtTest.Text = s
        s = LineInput(fileNum)

        '2  亞東
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxYadon.Text = s
        s = LineInput(fileNum)

        '3  待拌槽
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxWaitTank.Text = s
        s = LineInput(fileNum)

        '4  Ｂ配比查詢
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxABPB.Text = s
        s = LineInput(fileNum)
        'bABPB = CBool(s)

        '5  列印實際拌合時間
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxMix.Text = s
        s = LineInput(fileNum)

        '6  最大拌合方數
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label4.Text = s
        s = LineInput(fileNum)

        '7  公司名稱
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label5.Text = s
        s = LineInput(fileNum)

        '8  專案編號
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label18.Text = s
        s = LineInput(fileNum)

        '9  英文版
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxsingle.Text = s
        s = LineInput(fileNum)

        '10 UniCar(X)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label20.Text = s
        s = LineInput(fileNum)
        ' no use lblUniSer.Text = s

        '11 本機(X)
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label19.Text = s
        s = LineInput(fileNum)
        ' no use lblMachTot.Text = s

        '12 總重小數兩位
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxtTot2dec.Text = s
        s = LineInput(fileNum)

        '13 校磅歸零最大值
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label6.Text = s
        s = LineInput(fileNum)

        '14 校磅校準最大值
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label7.Text = s
        s = LineInput(fileNum)

        '15 校磅校準最小值
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label8.Text = s
        s = LineInput(fileNum)

        '16 拌合機編號
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label11.Text = s
        s = LineInput(fileNum)

        '17 調度資料伺服器IP
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label12.Text = s
        s = LineInput(fileNum)

        '18 配比備用一
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label21.Text = s
        s = LineInput(fileNum)

        '19 配比備用二
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label22.Text = s
        s = LineInput(fileNum)

        '20 配比備用三
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label23.Text = s
        s = LineInput(fileNum)

        '21 配比備用四
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label24.Text = s
        s = LineInput(fileNum)

        '22  102.5.11   CBC7報表格式
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxCBC7.Text = s
        s = LineInput(fileNum)

        '23  102.7.15   B配米數縮減
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxBPBsubEn.Text = s
        s = LineInput(fileNum)

        '24  102.11.23  列印水泥註記
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxCement.Text = s
        s = LineInput(fileNum)

        '25  102.12.11 廠別切換
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxBack.Text = s
        s = LineInput(fileNum)

        '26  103.6.9 英制切換
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxImperial.Text = s
        s = LineInput(fileNum)

        '27  103.6.9 非自動上傳
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxNonAuto.Text = s
        s = LineInput(fileNum)

        '28  104.4.25 殘留值
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxRemain.Text = s
        s = LineInput(fileNum)

        '29  104.4.25 SPARE1    A1-15
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP1.Text = s
        s = LineInput(fileNum)

        '30  104.4.25 SPARE2    日月報表寬格式
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP2.Text = s
        s = LineInput(fileNum)

        '31  104.4.25 SPARE3    B配存檔
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP3.Text = s
        s = LineInput(fileNum)

        '32  105.6.26 SPARE4    資料自動補傳
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP4.Text = s
        s = LineInput(fileNum)

        '33  105.6.26 SPARE5    可選遠端
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP5.Text = s
        s = LineInput(fileNum)

        '34  105.6.26 SPARE6    環泥烏日
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP6.Text = s
        s = LineInput(fileNum)

        '35  105.6.26 SPARE7
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label20.Text = s
        s = LineInput(fileNum)

        '36  105.6.26 SPARE8
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label20.Text = s
        s = LineInput(fileNum)

        '37  105.6.26 SPARE9
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        Label20.Text = s
        s = LineInput(fileNum)

        '38  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP1.Text = s
        s = LineInput(fileNum)

        '39  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP2.Text = s
        s = LineInput(fileNum)

        '40  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP3.Text = s
        s = LineInput(fileNum)

        '41  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP4.Text = s
        s = LineInput(fileNum)

        '42  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP5.Text = s
        s = LineInput(fileNum)

        '43  
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP6.Text = s
        s = LineInput(fileNum)

        '44 
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP7.Text = s
        s = LineInput(fileNum)

        '45 
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP8.Text = s
        s = LineInput(fileNum)

        '46   CheckBoxSP9
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP9.Text = s
        s = LineInput(fileNum)

        '47  105.7.10 SPARE10
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP10.Text = s
        s = LineInput(fileNum)

        '48  105.7.10 SPARE11
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP11.Text = s
        s = LineInput(fileNum)

        '49  105.7.10 SPARE12
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP12.Text = s
        s = LineInput(fileNum)

        '50  105.7.10 SPARE13
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP13.Text = s
        s = LineInput(fileNum)

        '51  105.7.10 SPARE14
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoxSP14.Text = s
        s = LineInput(fileNum)

        '52  105.7.10 SPARE15
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP15.Text = s
        s = LineInput(fileNum)

        '53  107.6.10 SPARE16
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP16.Text = s
        s = LineInput(fileNum)

        '54  107.6.10 SPARE17
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP17.Text = s
        s = LineInput(fileNum)

        '55  107.6.10 SPARE18
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP18.Text = s
        s = LineInput(fileNum)

        '56  107.6.10 SPARE19
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP19.Text = s
        s = LineInput(fileNum)

        '57  107.6.10 SPARE20
        s = LineInput(fileNum)
        s = LineInput(fileNum)
        CheckBoXSP20.Text = s
        s = LineInput(fileNum)
        'bCheckBoxSP20 = CBool(s)

        FileClose(fileNum)

        Exit Sub
Errorhandler:
        MsgBox("AdminName.des 內容有誤!", MsgBoxStyle.Critical, "找不到AdminName.des!")
        '105.10.8
        FileClose(fileNum)
    End Sub
    Public Sub WriteAdminName()
        '106.6.10 copy from WriteAdmin 
        '   將螢幕上標示文字 可能
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\AdminName.des")
        Dim fileNum
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\AdminName.des", OpenMode.Output)
        i = 0
        If True Then
            '1  1280*960
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTest.Text)
            PrintLine(fileNum, CheckBoxtTest.Checked.ToString)

            '2  亞東
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxYadon.Text)
            PrintLine(fileNum, CheckBoxYadon.Checked.ToString)

            '3  待拌槽
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxWaitTank.Text)
            PrintLine(fileNum, CheckBoxWaitTank.Checked.ToString)

            '4  Ｂ配比查詢
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxABPB.Text)
            PrintLine(fileNum, CheckBoxABPB.Checked.ToString)

            '5  列印實際拌合時間
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxMix.Text)
            PrintLine(fileNum, CheckBoxMix.Checked.ToString)

            '6  最大拌合方數
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label4.Text)
            PrintLine(fileNum, TextBoxMaxM3.Text)

            '7  公司名稱
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label5.Text)
            PrintLine(fileNum, TextBoxCompany.Text)

            '8  專案編號
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj.Text)

            '9  英文版
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxsingle.Text)
            PrintLine(fileNum, CheckBoxsingle.Checked.ToString)

            '10 配比密碼紀錄???
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label20.Text)
            PrintLine(fileNum, lblUniSer.Text)

            '11 本機(X)
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label19.Text)
            PrintLine(fileNum, lblMachTot.Text)

            '12 總重(X)
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxtTot2dec.Text)
            PrintLine(fileNum, CheckBoxtTot2dec.Checked.ToString)

            '13 校磅歸零最大值
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label6.Text)
            PrintLine(fileNum, TextBoxZeroMax.Text)

            '14 校磅校準最大值
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label7.Text)
            PrintLine(fileNum, TextBoxSpanMax.Text)

            '15 校磅校準最小值
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label8.Text)
            PrintLine(fileNum, TextBoxSpanMin.Text)

            '16 拌合機編號
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo.Text)

            '17 調度資料伺服器IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label12.Text)
            PrintLine(fileNum, TextBoxIP.Text)

            '18 配比備用一
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label21.Text)
            PrintLine(fileNum, TextBoxSP1.Text)

            '19 配比備用二
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label22.Text)
            PrintLine(fileNum, TextBoxSP2.Text)

            '20 配比備用三
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label23.Text)
            PrintLine(fileNum, TextBoxSP3.Text)

            '21 配比備用四
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label24.Text)
            PrintLine(fileNum, TextBoxSP4.Text)

            '22 CBC7報表格式
            '102.5.11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCBC7.Text)
            PrintLine(fileNum, CheckBoxCBC7.Checked.ToString)

            '23 B配米數縮減
            '102.7.15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBPBsubEn.Text)
            PrintLine(fileNum, CheckBoxBPBsubEn.Checked.ToString)

            '24
            '102.7.15 列印水泥註記
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxCement.Text)
            PrintLine(fileNum, CheckBoxCement.Checked.ToString)

            '25
            '102.12.11 廠別切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxBack.Text)
            PrintLine(fileNum, CheckBoxBack.Checked.ToString)

            '26
            '103.6.9 英制切換
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxImperial.Text)
            PrintLine(fileNum, CheckBoxImperial.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxImperial.Checked.ToString)

            '27
            '103.6.9 非自動上傳
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxNonAuto.Text)
            PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)
            '106.6.26 BUG PrintLine(fileNum, CheckBoxNonAuto.Checked.ToString)

            '28  104.9.4 COPY FROM GWAN 殘留值
            '104.4.25 REMAIN VISIBLE FOR PRINT
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxRemain.Text)
            PrintLine(fileNum, CheckBoxRemain.Checked.ToString)

            '29 B存放磁碟(\\192.)
            '104.4.25 SPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP1.Text)
            PrintLine(fileNum, CheckBoxSP1.Checked.ToString)

            '30
            '104.4.25 SPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP2.Text)
            PrintLine(fileNum, CheckBoxSP2.Checked.ToString)

            '31
            '104.4.25 SPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP3.Text)
            PrintLine(fileNum, CheckBoxSP3.Checked.ToString)

            '32
            '106.6.26 SPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP4.Text)
            PrintLine(fileNum, CheckBoxSP4.Checked.ToString)

            '33
            '106.6.26 SPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP5.Text)
            PrintLine(fileNum, CheckBoxSP5.Checked.ToString)

            '34
            '106.6.26 SPARE6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP6.Text)
            PrintLine(fileNum, CheckBoxSP6.Checked.ToString)

            '35
            '106.6.26 SPARE7
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP7.Text)
            PrintLine(fileNum, CheckBoxSP7.Checked.ToString)

            '36
            '106.6.26 SPARE8
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP8.Text)
            PrintLine(fileNum, CheckBoxSP8.Checked.ToString)

            '37
            '106.6.26 SPARE9
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP9.Text)
            PrintLine(fileNum, CheckBoxSP9.Checked.ToString)

            '38
            '106.6.26 TextBoxBpath
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label26.Text)
            PrintLine(fileNum, TextBoxBpath.Text)

            '39
            '106.6.26 TextBoxSPARE1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label27.Text)
            PrintLine(fileNum, TextBoxSPARE1.Text)

            '40
            '106.6.26 TextBoxSPARE2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label28.Text)
            PrintLine(fileNum, TextBoxSPARE2.Text)

            '41
            '105.6.26 TextBoxSPARE3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label29.Text)
            PrintLine(fileNum, TextBoxSPARE3.Text)

            '42
            '105.6.26 TextBoxSPARE4
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label30.Text)
            PrintLine(fileNum, TextBoxSPARE4.Text)

            '43
            '105.6.26 TextBoxSPARE5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label31.Text)
            PrintLine(fileNum, TextBoxSPARE5.Text)

            '44
            '105.7.19 遠端電腦IP
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, LabelIP_R.Text)
            PrintLine(fileNum, TextBoxIP_R.Text)

            '45
            '105.7.19 專案編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label18.Text)
            PrintLine(fileNum, TextBoxPrj2.Text)

            '46
            '105.7.19 拌合機編號2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, Label11.Text)
            PrintLine(fileNum, TextBoxMixerNo2.Text)

            '47
            '106.10.7 SPARE10
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP10.Text)
            PrintLine(fileNum, CheckBoxSP10.Checked.ToString)

            '48
            '106.10.7 SPARE11
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP11.Text)
            PrintLine(fileNum, CheckBoxSP11.Checked.ToString)

            '49
            '106.10.7 SPARE12
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP12.Text)
            PrintLine(fileNum, CheckBoxSP12.Checked.ToString)

            '50
            '106.10.7 SPARE13
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP13.Text)
            PrintLine(fileNum, CheckBoxSP13.Checked.ToString)

            '51
            '106.10.7 SPARE14
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoxSP14.Text)
            PrintLine(fileNum, CheckBoxSP14.Checked.ToString)

            '52
            '106.10.7 SPARE15
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP15.Text)
            PrintLine(fileNum, CheckBoXSP15.Checked.ToString)

            '53
            '107.6.10 SPARE16
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP16.Text)
            PrintLine(fileNum, CheckBoXSP16.Checked.ToString)

            '54
            '107.6.10 SPARE17
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP17.Text)
            PrintLine(fileNum, CheckBoXSP17.Checked.ToString)

            '55
            '107.6.10 SPARE18
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP18.Text)
            PrintLine(fileNum, CheckBoXSP18.Checked.ToString)

            '56
            '107.6.10 SPARE19
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP19.Text)
            PrintLine(fileNum, CheckBoXSP19.Checked.ToString)

            '57
            '107.6.10 SPARE20
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, CheckBoXSP20.Text)
            PrintLine(fileNum, CheckBoXSP20.Checked.ToString)

        End If
        FileClose(fileNum)
    End Sub
    Private Sub CheckBoxSP4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxSP4.CheckedChanged

    End Sub

    Private Sub CheckBoxFix_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxFix.CheckedChanged
        MainForm.ButtonFix.Visible = CheckBoxFix.Checked
        '107.8.30
        Call MainForm.ButtonFixClick()
    End Sub

    Private Sub CheckBoxBPBsubEn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxBPBsubEn.CheckedChanged

    End Sub

    Private Sub ComboBox232_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox232.SelectedIndexChanged
        sPlcCom2 = ComboBox232.Text
        SaveSetting("JS", "CBC800", "PlcCom2", sPlcCom2)
        PLC_232_2.Port = CInt(Microsoft.VisualBasic.Strings.Right(sPlcCom2, 1))

    End Sub


    Private Sub CheckBox232_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox232.CheckedChanged
        '107.1.31
        If CheckBox232.Checked = True Then
            bCheckBox232 = True
            'PLC_232_2.Open()
            MainForm.TimerRx.Enabled = True
        Else
            bCheckBox232 = False
            'PLC_232_2.Close()
            MainForm.TimerRx.Enabled = False
        End If
        SaveSetting("JS", "CBC800", "CheckBox232", CStr(bCheckBox232))
    End Sub

    Private Sub CheckBoxA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxA.CheckedChanged
        '107.2.3
        If CheckBoxA.Checked = True Then
            bCheckBoxA = True
        Else
            bCheckBoxA = False
        End If
        SaveSetting("JS", "CBC800", "CheckBoxA", CStr(bCheckBoxA))
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTLog.Click
        TextBox2.Text = ""
        Dim i
        If TLog.log(TLog.pointer) <> "" Then
            For i = TLog.pointer To 49
                TextBox2.Text &= TLog.log(i) & vbCrLf
            Next
        End If
        For i = 0 To TLog.pointer - 1
            TextBox2.Text &= TLog.log(i) & vbCrLf
        Next
        ButtonTLog.Text = "TLog " & TLog.pointer - 1
    End Sub


    Private Sub ButtonReadName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReadName.Click
        '107.6.10
        Call ReadAdminName()
        ButtonReadName.Visible = False
    End Sub

    Private Sub CheckBoxW1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxW1.CheckedChanged
        '107.8.20
        '   聯佶 W1 不參與水分補正 
        bCheckBoxW1 = CheckBoxW1.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxW1", CStr(bCheckBoxW1))
    End Sub

    Private Sub CheckBoxW2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxW2.CheckedChanged
        '107.8.20
        '   聯佶 W1 不參與水分補正 
        bCheckBoxW2 = CheckBoxW2.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxW2", CStr(bCheckBoxW2))

    End Sub

    Private Sub CheckBoxW3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxW3.CheckedChanged
        '107.8.20
        '   聯佶 W1 不參與水分補正 
        bCheckBoxW3 = CheckBoxW3.Checked
        SaveSetting("JS", "CBC800", "bCheckBoxW3", CStr(bCheckBoxW3))

    End Sub



    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub CheckBoxAE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxAE.CheckedChanged

    End Sub
    Private Sub WriteCemWarehouse()
        '110.9.14 ref Write Admin
        Dim file
        Dim i As Integer
        file = My.Computer.FileSystem.DirectoryExists(sSMdbName & "")
        If file = False Then My.Computer.FileSystem.CreateDirectory(sSMdbName & "")
        file = My.Computer.FileSystem.FileExists(sSMdbName & "\CemWarehouse.des")
        Dim fileNum, s
        fileNum = FreeFile()
        FileOpen(fileNum, sSMdbName & "\CemWarehouse.des", OpenMode.Output)
        i = 0
        If file <> False Then
            '1
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC1.Checked.ToString
            PrintLine(fileNum, s)

            '2
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC2.Checked.ToString
            PrintLine(fileNum, s)

            '3
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC3.Checked.ToString
            PrintLine(fileNum, s)

            '4  
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC4.Checked.ToString
            PrintLine(fileNum, s)

            '5
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC5.Checked.ToString
            PrintLine(fileNum, s)

            '6
            i += 1
            PrintLine(fileNum, i.ToString)
            PrintLine(fileNum, i.ToString)
            s = CheckBoxC6.Checked.ToString
            PrintLine(fileNum, s)

        End If
        FileClose(fileNum)
    End Sub

    Private Sub CheckBoXSP19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoXSP19.CheckedChanged

    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        Dim i
        DV(CInt(TextBoxD.Text)) = CInt(TextBoxV.Text)
    End Sub
End Class