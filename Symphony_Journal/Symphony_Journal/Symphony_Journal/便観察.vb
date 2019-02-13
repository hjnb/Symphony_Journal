Imports System.Text

Public Class 便観察

    'ユニット名
    Private unitName As String

    '日付(西暦)
    Private adStr As String

    '日付(和暦)
    Private warekiStr As String

    '曜日
    Private dayOfWeek As String

    'cellEnterフラグ
    Private canEnter As Boolean = False

    '文字数制限用
    Private Const LIMIT_LENGTH_BYTE As Integer = 90

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="unitName">ユニット名</param>
    ''' <param name="warekiStr">和暦</param>
    ''' <remarks></remarks>
    Public Sub New(unitName As String, adStr As String, warekiStr As String, dayOfWeek As String, parentX As Integer, parentY As Integer)
        InitializeComponent()

        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        '位置設定
        Me.Location = New Point(parentX + 295, parentY + 40)

        Me.unitName = unitName
        Me.adStr = adStr
        Me.warekiStr = warekiStr
        Me.dayOfWeek = dayOfWeek
    End Sub

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 便観察_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'ラベル設定
        unitLabel.Text = unitName & unitLabel.Text
        dateLabel.Text = warekiStr & "(" & dayOfWeek & ")"

        'dgv初期設定
        initDgvBen()

        'データ表示
        displayDgvBen()
    End Sub

    ''' <summary>
    ''' データグリッドビュー初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvBen()
        Util.EnableDoubleBuffering(dgvBen)

        With dgvBen
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .BorderStyle = BorderStyle.None
            .MultiSelect = False
            .RowHeadersVisible = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersHeight = 20
            .RowTemplate.Height = 17
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .ScrollBars = ScrollBars.None
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '空行追加等
        Dim dt As New DataTable()
        dt.Columns.Add("Text", Type.GetType("System.String"))
        For i = 0 To 4
            Dim row As DataRow = dt.NewRow()
            row(0) = ""
            dt.Rows.Add(row)
        Next
        dgvBen.DataSource = dt

        '幅設定等
        With dgvBen
            With .Columns("Text")
                .HeaderText = "便観察"
                .Width = 432
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
        End With

        'cellEnterフラグ更新
        canEnter = True
    End Sub

    ''' <summary>
    ''' 便観察データ表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvBen()
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Gyo, Text from Ben where Ymd='" & adStr & "' And Unit='" & unitName & "' order by Gyo"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        While Not rs.EOF
            Dim gyo As Integer = rs.Fields("Gyo").Value
            dgvBen("Text", gyo - 1).Value = Util.checkDBNullValue(rs.Fields("Text").Value)
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()
    End Sub

    ''' <summary>
    ''' セルエンターイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvBen_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBen.CellEnter
        If canEnter Then
            dgvBen.BeginEdit(False)
        End If
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '入力テキスト
        Dim text1 As String = Util.checkDBNullValue(dgvBen("Text", 0).Value) '1行目
        Dim text2 As String = Util.checkDBNullValue(dgvBen("Text", 1).Value) '2行目
        Dim text3 As String = Util.checkDBNullValue(dgvBen("Text", 2).Value) '3行目
        Dim text4 As String = Util.checkDBNullValue(dgvBen("Text", 3).Value) '4行目
        Dim text5 As String = Util.checkDBNullValue(dgvBen("Text", 4).Value) '5行目

        '未入力の場合
        If text1 = "" AndAlso text2 = "" AndAlso text3 = "" AndAlso text4 = "" AndAlso text5 = "" Then
            'メッセージ表示
            benLabel.Visible = True
            Return
        Else
            'メッセージ非表示
            benLabel.Visible = False
        End If

        '既存データ削除
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        cmd.CommandText = "delete from Ben where Ymd='" & adStr & "' and Unit='" & unitName & "'"
        cmd.Execute()

        '登録
        Dim rs As New ADODB.Recordset()
        rs.Open("Ben", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        For i As Integer = 4 To 0 Step -1
            Dim inputText As String = Util.checkDBNullValue(dgvBen("Text", i).Value)
            If inputText = "" Then
                Continue For
            End If
            Dim gyo As Integer = i + 1
            rs.AddNew()
            rs.Fields("Div").Value = TopForm.DIV
            rs.Fields("Ymd").Value = adStr
            rs.Fields("Unit").Value = unitName
            rs.Fields("Gyo").Value = gyo
            rs.Fields("Text").Value = inputText
        Next
        rs.Update()
        rs.Close()
        cnn.Close()

        'クリア
        For i As Integer = 0 To 4
            dgvBen("Text", i).Value = ""
        Next
    End Sub

    Private Sub dgvBen_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvBen.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)

        'イベントハンドラを削除、追加
        RemoveHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
        AddHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
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