Public Class ユニット日誌
    'ログインユーザーの印影ファイルパス
    Private userSealFilePath As String

    '右クリック色付け可否
    Private canPaintFontColor As Boolean = False

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    'フォントカラー
    Private fontColorTable As New Dictionary(Of Integer, Color) From {{0, Color.Black}, {1, Color.Blue}, {2, Color.Red}}

    '編集不可部分のセルスタイル
    Private readOnlyCellStyle As DataGridViewCellStyle

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
    ''' 対象ユニット、日付の日誌データ表示
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
        dayWorkPic.ImageLocation = Nothing
        nightWorkPic.ImageLocation = Nothing
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
            .RowTemplate.Height = 17
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
                .Width = 486
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
        dayWorkPic.ImageLocation = Nothing
    End Sub

    ''' <summary>
    ''' 夜勤印影画像ダブルクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub nightWorkPic_DoubleClick(sender As Object, e As System.EventArgs) Handles nightWorkPic.DoubleClick
        '画像を空白に
        nightWorkPic.ImageLocation = Nothing
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
End Class