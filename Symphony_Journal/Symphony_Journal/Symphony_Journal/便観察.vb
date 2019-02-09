Public Class 便観察

    'ユニット名
    Private unitName As String

    '日付(西暦)
    Private adStr As String

    '日付(和暦)
    Private warekiStr As String

    'cellEnterフラグ
    Private canEnter As Boolean = False

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="unitName">ユニット名</param>
    ''' <param name="warekiStr">和暦</param>
    ''' <remarks></remarks>
    Public Sub New(unitName As String, adStr As String, warekiStr As String)
        InitializeComponent()
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        '位置設定
        '
        '

        Me.unitName = unitName
        Me.adStr = adStr
        Me.warekiStr = warekiStr
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
        dateLabel.Text = warekiStr

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

    Private Sub displayDgvBen()

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

        '登録


    End Sub
End Class