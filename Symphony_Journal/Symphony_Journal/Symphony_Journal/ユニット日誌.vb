Public Class ユニット日誌
    'ログインユーザーの印影ファイルパス
    Private userSealFilePath As String

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    '編集不可部分のセルスタイル
    Private readOnlyCellStyle As DataGridViewCellStyle

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(sealFileName As String)
        InitializeComponent()

        Me.WindowState = FormWindowState.Maximized
        userSealFilePath = TopForm.sealBoxDirPath & "\" & sealFileName & ".wmf"
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

        '現在日付を初期値に
        YmdBox.setADStr(Today.ToString("yyyy/MM/dd"))

        'テキストボックス初期設定
        initTextBox()

        'セルスタイル作成
        createCellStyles()

        'dgv初期設定
        initDgvUnitDiary()
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
    End Sub

    ''' <summary>
    ''' 入居者名リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub residentListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles residentListBox.SelectedValueChanged
        '選択した氏名をdgvのセルへ反映
        '
        '
        '
    End Sub

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
End Class