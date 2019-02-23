Imports System.Text

Public Class ユニット日誌
    'ログインユーザーの印影ファイルパス
    Private userSealFilePath As String

    '右クリック色付け可否
    Private canPaintFontColor As Boolean = False

    '内容文字数制限用
    Private Const LIMIT_LENGTH_BYTE As Integer = 90

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    'フォントカラー
    Private fontColorTable As New Dictionary(Of Integer, Color) From {{0, Color.Black}, {1, Color.Blue}, {2, Color.Red}}

    '編集不可部分のセルスタイル
    Private readOnlyCellStyle As DataGridViewCellStyle

    '印刷条件フォーム
    Private printForm As 印刷条件

    '便観察フォーム
    Private benForm As 便観察

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(sealFileName As String, className As String)
        InitializeComponent()
        Me.WindowState = FormWindowState.Maximized

        '印影ファイルパス
        userSealFilePath = TopForm.sealBoxDirPath & "\" & sealFileName & ".wmf"

        '右クリック色付け可否
        Dim classNum As String = className.Substring(0, 1)
        If classNum = "5" OrElse classNum = "8" Then
            canPaintFontColor = True
        Else
            paintColorLabel.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' Loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ユニット日誌_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'ユニットリスト初期値
        unitListBox.Items.AddRange(unitArray)

        'テキストボックス初期設定
        initTextBox()

        'セルスタイル作成
        createCellStyles()

        'dgv初期設定
        initDgvUnitDiary()

        '現在日付を初期値に
        YmdBox.setADStr(Today.ToString("yyyy/MM/dd"))
    End Sub

    ''' <summary>
    ''' セルスタイル作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub createCellStyles()
        readOnlyCellStyle = New DataGridViewCellStyle()
        readOnlyCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        readOnlyCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        readOnlyCellStyle.SelectionForeColor = Color.Black
        readOnlyCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    ''' <summary>
    ''' 対象日付、ユニットの日誌データ表示
    ''' </summary>
    ''' <param name="unitName">ユニット名</param>
    ''' <param name="ymd">日付(yyyy/MM/dd)</param>
    ''' <remarks></remarks>
    Private Sub displayUnitDiary(unitName As String, ymd As String)
        'データ表示部分クリア
        inputClear()

        '表示処理
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Ymd, Unit, Gyo, Nyu1, Nyu2, Nyu3, Gai1, Gai2, Gai3, Kyo1, Kyo2, Kyo3, Sign6, Sign7, Nam, NClr, Text, TClr from UNis where Ymd='" & ymd & "' And Unit='" & unitName & "' order by Gyo"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        While Not rs.EOF
            Dim gyo As Integer = Util.checkDBNullValue(rs.Fields("Gyo").Value)
            If gyo = 0 Then
                'テキスト
                Nyu1Box.Text = If(Util.checkDBNullValue(rs.Fields("Nyu1").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Nyu1").Value)) '入院者数 男
                Nyu2Box.Text = If(Util.checkDBNullValue(rs.Fields("Nyu2").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Nyu2").Value)) '入院者数 女
                Nyu3Box.Text = If(Util.checkDBNullValue(rs.Fields("Nyu3").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Nyu3").Value)) '入院者数 計
                Gai1Box.Text = If(Util.checkDBNullValue(rs.Fields("Gai1").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Gai1").Value)) '外泊者数 男
                Gai2Box.Text = If(Util.checkDBNullValue(rs.Fields("Gai2").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Gai2").Value)) '外泊者数 女
                Gai3Box.Text = If(Util.checkDBNullValue(rs.Fields("Gai3").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Gai3").Value)) '外泊者数 計
                Kyo1Box.Text = If(Util.checkDBNullValue(rs.Fields("Kyo1").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Kyo1").Value)) '入居者数 男
                Kyo2Box.Text = If(Util.checkDBNullValue(rs.Fields("Kyo2").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Kyo2").Value)) '入居者数 女
                Kyo3Box.Text = If(Util.checkDBNullValue(rs.Fields("Kyo3").Value) = 0, "", Util.checkDBNullValue(rs.Fields("Kyo3").Value)) '入居者数 計

                '印影画像
                Dim dayWorkSealPath As String = TopForm.sealBoxDirPath & "\" & Util.checkDBNullValue(rs.Fields("Sign6").Value) & ".wmf"
                Dim nightWorkSealPath As String = TopForm.sealBoxDirPath & "\" & Util.checkDBNullValue(rs.Fields("Sign7").Value) & ".wmf"
                If System.IO.File.Exists(dayWorkSealPath) Then
                    dayWorkPic.ImageLocation = dayWorkSealPath
                End If
                If System.IO.File.Exists(nightWorkSealPath) Then
                    nightWorkPic.ImageLocation = nightWorkSealPath
                End If
            Else
                '入居者名列
                If gyo <= 30 Then
                    dgvUnitDiary("Nam", gyo - 1).Value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                    dgvUnitDiary("Nam", gyo - 1).Style.ForeColor = fontColorTable(CInt(Util.checkDBNullValue(rs.Fields("NClr").Value)))
                    dgvUnitDiary("Nam", gyo - 1).Style.SelectionForeColor = fontColorTable(CInt(Util.checkDBNullValue(rs.Fields("NClr").Value)))
                End If

                '経過内容列
                dgvUnitDiary("Text", gyo - 1).Value = Util.checkDBNullValue(rs.Fields("Text").Value)
                dgvUnitDiary("Text", gyo - 1).Style.ForeColor = fontColorTable(CInt(Util.checkDBNullValue(rs.Fields("TClr").Value)))
                dgvUnitDiary("Text", gyo - 1).Style.SelectionForeColor = fontColorTable(CInt(Util.checkDBNullValue(rs.Fields("TClr").Value)))
            End If
            rs.MoveNext()
        End While

        rs.Close()
        cnn.Close()
    End Sub

    ''' <summary>
    ''' 入力内容クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub inputClear()
        Nyu1Box.Text = ""
        Nyu2Box.Text = ""
        Nyu3Box.Text = ""
        Gai1Box.Text = ""
        Gai2Box.Text = ""
        Gai3Box.Text = ""
        Kyo1Box.Text = ""
        Kyo2Box.Text = ""
        Kyo3Box.Text = ""
        dayWorkPic.ImageLocation = ""
        nightWorkPic.ImageLocation = ""
        For i As Integer = 1 To 34
            If i <> 18 Then
                If i = 31 Then
                    dgvUnitDiary("Text", i).Value = ""
                    dgvUnitDiary("Text", i).Style.ForeColor = Color.Black
                    dgvUnitDiary("Text", i).Style.SelectionForeColor = Color.Black
                Else
                    dgvUnitDiary("Nam", i).Value = ""
                    dgvUnitDiary("Nam", i).Style.ForeColor = Color.Black
                    dgvUnitDiary("Nam", i).Style.SelectionForeColor = Color.Black
                    dgvUnitDiary("Text", i).Value = ""
                    dgvUnitDiary("Text", i).Style.ForeColor = Color.Black
                    dgvUnitDiary("Text", i).Style.SelectionForeColor = Color.Black
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' dgv初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvUnitDiary()
        Util.EnableDoubleBuffering(dgvUnitDiary)

        With dgvUnitDiary
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .BorderStyle = BorderStyle.None
            .MultiSelect = False
            .RowHeadersVisible = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersVisible = False
            .RowTemplate.Height = 15
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .ScrollBars = ScrollBars.None
            .ImeMode = Windows.Forms.ImeMode.Hiragana
            If canPaintFontColor Then
                .ContextMenuStrip = Me.colorContextMenu
            End If
        End With

        '列追加、空の行追加
        dgvUnitDiary.dt.Columns.Add("Nam", Type.GetType("System.String"))
        dgvUnitDiary.dt.Columns.Add("Text", Type.GetType("System.String"))
        Dim row As DataRow = dgvUnitDiary.dt.NewRow()
        row(0) = "入居者名"
        row(1) = "日勤　介護支援経過内容"
        dgvUnitDiary.dt.Rows.Add(row)
        For i = 0 To 16
            row = dgvUnitDiary.dt.NewRow()
            row(0) = ""
            row(1) = ""
            dgvUnitDiary.dt.Rows.Add(row)
        Next
        row = dgvUnitDiary.dt.NewRow()
        row(0) = "入居者名"
        row(1) = "夜勤　介護支援経過内容"
        dgvUnitDiary.dt.Rows.Add(row)
        For i = 0 To 15
            row = dgvUnitDiary.dt.NewRow()
            row(0) = ""
            row(1) = ""
            dgvUnitDiary.dt.Rows.Add(row)
        Next
        dgvUnitDiary.dt.Rows(31)(0) = "特記事項"

        '表示
        dgvUnitDiary.DataSource = dgvUnitDiary.dt

        '幅設定等
        With dgvUnitDiary
            With .Columns("Nam")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Width = 100
            End With
            With .Columns("Text")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Width = 456
            End With
        End With

        'ヘッダー部分の設定
        dgvUnitDiary("Nam", 0).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 0).ReadOnly = True
        dgvUnitDiary("Text", 0).Style = readOnlyCellStyle
        dgvUnitDiary("Text", 0).ReadOnly = True
        dgvUnitDiary("Nam", 18).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 18).ReadOnly = True
        dgvUnitDiary("Text", 18).Style = readOnlyCellStyle
        dgvUnitDiary("Text", 18).ReadOnly = True
        '使用しないセル
        dgvUnitDiary("Nam", 31).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 31).ReadOnly = True
        dgvUnitDiary("Nam", 32).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 32).ReadOnly = True
        dgvUnitDiary("Nam", 33).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 33).ReadOnly = True
        dgvUnitDiary("Nam", 34).Style = readOnlyCellStyle
        dgvUnitDiary("Nam", 34).ReadOnly = True

        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgvUnitDiary.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    ''' <summary>
    ''' テキストボックス初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initTextBox()
        '入院者数　男
        Nyu1Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Nyu1Box.TextAlign = HorizontalAlignment.Center

        '入院者数　女
        Nyu2Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Nyu2Box.TextAlign = HorizontalAlignment.Center

        '入院者数　計
        Nyu3Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Nyu3Box.TextAlign = HorizontalAlignment.Center

        '外泊者数　男
        Gai1Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Gai1Box.TextAlign = HorizontalAlignment.Center

        '外泊者数　女
        Gai2Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Gai2Box.TextAlign = HorizontalAlignment.Center

        '外泊者数　計
        Gai3Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Gai3Box.TextAlign = HorizontalAlignment.Center

        '入居者数　男
        Kyo1Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Kyo1Box.TextAlign = HorizontalAlignment.Center

        '入居者数　女
        Kyo2Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Kyo2Box.TextAlign = HorizontalAlignment.Center

        '入居者数　計
        Kyo3Box.ImeMode = Windows.Forms.ImeMode.Alpha
        Kyo3Box.TextAlign = HorizontalAlignment.Center
    End Sub

    ''' <summary>
    ''' 入居者リスト取得
    ''' </summary>
    ''' <param name="unitName">ユニット名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getResidentList(unitName As String) As List(Of String)
        Dim resultList As New List(Of String) '取得結果リスト

        '存在しないユニット名の場合は空リストを返す
        Dim existFlg As Boolean = False
        For Each s As String In unitArray
            If s = unitName Then
                existFlg = True
                Exit For
            End If
        Next
        If existFlg = False Then
            Return resultList
        End If

        '対象のユニットの入居者リスト作成
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Nam from UsrM where Dsp=1 And Unt='" & unitName & "' order by Kana"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        While Not rs.EOF
            resultList.Add(Util.checkDBNullValue(rs.Fields("Nam").Value))
            rs.MoveNext()
        End While
        Return resultList
    End Function

    ''' <summary>
    ''' ユニット値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub unitListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles unitListBox.SelectedValueChanged
        'ユニット名
        Dim unitName As String = CType(sender, ListBox).SelectedItem.ToString

        '対象のユニットの入居者リストをセット
        residentListBox.Items.Clear()
        residentListBox.Items.AddRange(getResidentList(unitName).ToArray())

        '日誌データ表示
        Dim ymdStr As String = YmdBox.getADStr()
        displayUnitDiary(unitName, ymdStr)
    End Sub

    ''' <summary>
    ''' 入居者名リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub residentListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles residentListBox.SelectedValueChanged
        '選択氏名
        Dim selectedName As String = residentListBox.SelectedItem

        '選択した氏名をdgvのセルへ反映
        If Not IsNothing(dgvUnitDiary.CurrentCell) AndAlso dgvUnitDiary.CurrentCell.ReadOnly = False AndAlso selectedName <> "" Then
            dgvUnitDiary.CurrentCell.Value = selectedName
        End If
    End Sub

    ''' <summary>
    ''' テキストボックス部分のkeyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub textBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Nyu1Box.KeyDown, Nyu2Box.KeyDown, Nyu3Box.KeyDown, Gai1Box.KeyDown, Gai2Box.KeyDown, Gai3Box.KeyDown, Kyo1Box.KeyDown, Kyo2Box.KeyDown, Kyo3Box.KeyDown
        Dim tb As TextBox = CType(sender, TextBox)
        Dim tbName As String = tb.Name
        Dim tbType As String = tbName.Substring(0, 3)
        Dim tbNum As String = tbName.Substring(3, 1)
        If e.KeyCode = Keys.Enter AndAlso tbName <> "Kyo3Box" Then
            Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
        ElseIf e.KeyCode = Keys.Up AndAlso tbType <> "Nyu" Then
            tbType = If(tbType = "Gai", "Nyu", "Gai")
            Dim targetName As String = tbType & tbNum & "Box"
            Controls(targetName).Focus()
        ElseIf e.KeyCode = Keys.Down AndAlso tbType <> "Kyo" Then
            tbType = If(tbType = "Nyu", "Gai", "Kyo")
            Dim targetName = tbType & tbNum & "Box"
            Controls(targetName).Focus()
        End If
    End Sub

    ''' <summary>
    ''' 日付ボックス変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub YmdBox_YmdTextChange(sender As Object, e As System.EventArgs) Handles YmdBox.YmdTextChange
        Dim unitName As String = unitListBox.SelectedItem
        Dim ymdStr As String = YmdBox.getADStr()
        displayUnitDiary(unitName, ymdStr)
    End Sub

    ''' <summary>
    ''' 日勤ラジオボタン変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rbtnDayWork_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnDayWork.CheckedChanged
        If rbtnDayWork.Checked AndAlso System.IO.File.Exists(userSealFilePath) Then
            'ログイン者の印影画像セット
            dayWorkPic.ImageLocation = userSealFilePath
        End If
    End Sub

    ''' <summary>
    ''' 夜勤ラジオボタン変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rbtnNightWork_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnNightWork.CheckedChanged
        If rbtnNightWork.Checked AndAlso System.IO.File.Exists(userSealFilePath) Then
            'ログイン者の印影画像セット
            nightWorkPic.ImageLocation = userSealFilePath
        End If
    End Sub

    ''' <summary>
    ''' 日勤印影画像ダブルクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dayWorkPic_DoubleClick(sender As Object, e As System.EventArgs) Handles dayWorkPic.DoubleClick
        '画像を空白に
        dayWorkPic.ImageLocation = ""
    End Sub

    ''' <summary>
    ''' 夜勤印影画像ダブルクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub nightWorkPic_DoubleClick(sender As Object, e As System.EventArgs) Handles nightWorkPic.DoubleClick
        '画像を空白に
        nightWorkPic.ImageLocation = ""
    End Sub

    ''' <summary>
    ''' 右クリックメニューで黒選択時のイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub paintBlack_Click(sender As System.Object, e As System.EventArgs) Handles paintBlack.Click
        If Not IsNothing(dgvUnitDiary.CurrentCell) AndAlso dgvUnitDiary.CurrentCell.ReadOnly = False Then
            dgvUnitDiary.CurrentCell.Style.ForeColor = Color.Black
            dgvUnitDiary.CurrentCell.Style.SelectionForeColor = Color.Black
        End If
    End Sub

    ''' <summary>
    ''' 右クリックメニューで青選択時のイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub paintBlue_Click(sender As System.Object, e As System.EventArgs) Handles paintBlue.Click
        If Not IsNothing(dgvUnitDiary.CurrentCell) AndAlso dgvUnitDiary.CurrentCell.ReadOnly = False Then
            dgvUnitDiary.CurrentCell.Style.ForeColor = Color.Blue
            dgvUnitDiary.CurrentCell.Style.SelectionForeColor = Color.Blue
        End If
    End Sub

    ''' <summary>
    ''' 右クリックメニューで赤選択時のイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub paintRed_Click(sender As System.Object, e As System.EventArgs) Handles paintRed.Click
        If Not IsNothing(dgvUnitDiary.CurrentCell) AndAlso dgvUnitDiary.CurrentCell.ReadOnly = False Then
            dgvUnitDiary.CurrentCell.Style.ForeColor = Color.Red
            dgvUnitDiary.CurrentCell.Style.SelectionForeColor = Color.Red
        End If
    End Sub

    ''' <summary>
    ''' 登録用数値に変換
    ''' </summary>
    ''' <param name="inputStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function convNum(inputStr As String) As Integer
        If inputStr = "" Then
            Return 0
        Else
            If IsNumeric(inputStr) Then
                Return CInt(inputStr)
            Else
                Return -1
            End If
        End If
    End Function

    ''' <summary>
    ''' 印刷ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        If IsNothing(printForm) OrElse printForm.IsDisposed Then
            printForm = New 印刷条件()
            printForm.Owner = Me
            printForm.ShowDialog()
            printForm.Dispose()
        End If
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '入力内容取得
        Dim ymd As String = YmdBox.getADStr() '日付
        Dim unit As String = If(IsNothing(unitListBox.SelectedItem), "", unitListBox.SelectedItem.ToString()) 'ユニット
        Dim sign6 As String = System.IO.Path.GetFileNameWithoutExtension(dayWorkPic.ImageLocation) '日勤印影
        Dim sign7 As String = System.IO.Path.GetFileNameWithoutExtension(nightWorkPic.ImageLocation) '夜勤印影
        Dim nyu1 As Integer = convNum(Nyu1Box.Text) '入院者数　男
        Dim nyu2 As Integer = convNum(Nyu2Box.Text) '入院者数　女
        Dim nyu3 As Integer = convNum(Nyu3Box.Text) '入院者数　計
        Dim gai1 As Integer = convNum(Gai1Box.Text) '外泊者数　男
        Dim gai2 As Integer = convNum(Gai2Box.Text) '外泊者数　女
        Dim gai3 As Integer = convNum(Gai3Box.Text) '外泊者数　計
        Dim kyo1 As Integer = convNum(Kyo1Box.Text) '入居者数　男
        Dim kyo2 As Integer = convNum(Kyo2Box.Text) '入居者数　女
        Dim kyo3 As Integer = convNum(Kyo3Box.Text) '入居者数　計

        '入力チェック
        If unit = "" Then
            MsgBox("ユニットを選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If
        If rbtnDayWork.Checked = False AndAlso rbtnNightWork.Checked = False Then
            MsgBox("日勤/夜勤が選択されていません。", MsgBoxStyle.Exclamation)
            Return
        End If
        If nyu1 = -1 OrElse nyu2 = -1 OrElse nyu3 = -1 OrElse gai1 = -1 OrElse gai2 = -1 OrElse gai3 = -1 OrElse kyo1 = -1 OrElse kyo2 = -1 OrElse kyo3 = -1 Then
            MsgBox("入院者数、外泊者数、入居者数は数値で入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '既存データ削除
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        cmd.CommandText = "delete from UNis where Ymd='" & ymd & "' and Unit='" & unit & "'"
        cmd.Execute()

        '登録
        Dim rs As New ADODB.Recordset()
        rs.Open("UNis", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        'Gyo = 0 のみ
        rs.AddNew()
        rs.Fields("Div").Value = TopForm.DIV 'Div
        rs.Fields("Ymd").Value = ymd '日付
        rs.Fields("Unit").Value = unit 'ユニット
        rs.Fields("Gyo").Value = 0 'Gyo
        rs.Fields("Nyu1").Value = nyu1 '入院者数　男
        rs.Fields("Nyu2").Value = nyu2 '入院者数　女
        rs.Fields("Nyu3").Value = nyu3 '入院者数　計
        rs.Fields("Gai1").Value = gai1 '外泊者数　男
        rs.Fields("Gai2").Value = gai2 '外泊者数　女
        rs.Fields("Gai3").Value = gai3 '外泊者数　計
        rs.Fields("Kyo1").Value = kyo1 '入院者数　男
        rs.Fields("Kyo2").Value = kyo2 '入院者数　女
        rs.Fields("Kyo3").Value = kyo3 '入院者数　計
        rs.Fields("Sign6").Value = sign6 '日勤印影ファイル名
        rs.Fields("Sign7").Value = sign7 '夜勤印影ファイル名
        rs.Fields("Nam").Value = ""
        rs.Fields("NClr").Value = 0
        rs.Fields("Text").Value = ""
        rs.Fields("TClr").Value = 0
        'Gyo = 0 以外
        For i As Integer = 0 To 34
            If i = 0 OrElse i = 18 Then
                Continue For
            End If

            Dim inputNam As String = Util.checkDBNullValue(dgvUnitDiary("Nam", i).Value) '名前
            Dim inputText As String = Util.checkDBNullValue(dgvUnitDiary("Text", i).Value) '内容

            If inputNam = "" AndAlso inputText = "" Then
                Continue For
            End If

            Dim gyo As Integer = i + 1 'Gyo
            Dim nClr As Color = dgvUnitDiary("Nam", i).Style.ForeColor
            Dim nClrNum As Integer = If(nClr = Color.Black, 0, If(nClr = Color.Blue, 1, 2)) 'NClr
            Dim tClr As Color = dgvUnitDiary("Text", i).Style.ForeColor
            Dim tClrNum As Integer = If(tClr = Color.Black, 0, If(tClr = Color.Blue, 1, 2)) 'TClr

            rs.AddNew()
            rs.Fields("Div").Value = TopForm.DIV 'Div
            rs.Fields("Ymd").Value = ymd '日付
            rs.Fields("Unit").Value = unit 'ユニット
            rs.Fields("Gyo").Value = gyo 'Gyo
            rs.Fields("Nyu1").Value = 0
            rs.Fields("Nyu2").Value = 0
            rs.Fields("Nyu3").Value = 0
            rs.Fields("Gai1").Value = 0
            rs.Fields("Gai2").Value = 0
            rs.Fields("Gai3").Value = 0
            rs.Fields("Kyo1").Value = 0
            rs.Fields("Kyo2").Value = 0
            rs.Fields("Kyo3").Value = 0
            rs.Fields("Sign6").Value = ""
            rs.Fields("Sign7").Value = ""
            rs.Fields("Nam").Value = inputNam '名前
            rs.Fields("NClr").Value = nClrNum '名前色番号
            rs.Fields("Text").Value = inputText '内容
            rs.Fields("TClr").Value = tClrNum '内容色番号
        Next
        rs.Update()
        rs.Close()
        cnn.Close()

        MsgBox("登録しました。", MsgBoxStyle.Information)

    End Sub

    ''' <summary>
    ''' 便観察ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnObserve_Click(sender As System.Object, e As System.EventArgs) Handles btnObserve.Click
        If IsNothing(benForm) OrElse benForm.IsDisposed Then
            Dim unit As String = If(IsNothing(unitListBox.SelectedItem), "", unitListBox.SelectedItem.ToString())
            Dim ad As String = YmdBox.getADStr()
            Dim wareki As String = YmdBox.getWarekiStr()
            Dim dayStr As String = YmdBox.getDay()
            If unit = "" Then
                MsgBox("ユニットを選択して下さい。", MsgBoxStyle.Exclamation)
                Return
            End If
            benForm = New 便観察(unit, ad, wareki, dayStr, Me.Location.X, Me.Location.Y)
            benForm.StartPosition = FormStartPosition.Manual
            benForm.Owner = Me
            benForm.ShowDialog()
            benForm.Dispose()
        End If
    End Sub

    Private Sub dgvUnitDiary_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvUnitDiary.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)
        Dim columnName As String = dgvUnitDiary.Columns(dgvUnitDiary.CurrentCell.ColumnIndex).Name

        If columnName = "Text" Then
            'イベントハンドラを削除、追加
            RemoveHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
            AddHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
        End If
    End Sub

    Private Sub dgvTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim text As String = CType(sender, DataGridViewTextBoxEditingControl).Text
        Dim lengthByte As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(text)

        If lengthByte >= LIMIT_LENGTH_BYTE Then '設定されているバイト数以上の時
            If e.KeyChar = ChrW(Keys.Back) Then
                'Backspaceは入力可能
                e.Handled = False
            Else
                '入力できなくする
                e.Handled = True
            End If
        End If
    End Sub
End Class