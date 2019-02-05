Public Class 印刷条件

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    'テキストボックスのマウスダウンイベント制御用
    Private mdFlag As Boolean = False

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        InitializeComponent()

        Me.StartPosition = FormStartPosition.CenterScreen
        Me.KeyPreview = True
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub

    ''' <summary>
    ''' keyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印刷条件_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印刷条件_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '日付ボックスの初期値を現在日付にセット
        Dim todayStr As String = Today.ToString("yyyy/MM/dd")
        startYmdBox.setADStr(todayStr)
        endYmdBox.setADStr(todayStr)

        '印影ファイル名、印刷ラジオボタンの設定読み込み
        facilityManagerTextBox.Text = Util.getIniString("System", "Sign1", TopForm.iniFilePath)
        consulteeTextBox.Text = Util.getIniString("System", "Sign2", TopForm.iniFilePath)
        specialistTextBox.Text = Util.getIniString("System", "Sign3", TopForm.iniFilePath)
        consensual1TextBox.Text = Util.getIniString("System", "Sign4", TopForm.iniFilePath)
        consensual2TextBox.Text = Util.getIniString("System", "Sign5", TopForm.iniFilePath)
        Dim printState As String = Util.getIniString("System", "Printer", TopForm.iniFilePath)
        If printState = "Y" Then
            rbtnPrint.Checked = True
        Else
            rbtnPreview.Checked = True
        End If

        'dgv初期設定
        initDgvUnit()

    End Sub

    ''' <summary>
    ''' dgv初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvUnit()
        Util.EnableDoubleBuffering(dgvUnit)

        With dgvUnit
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
            .ReadOnly = True
            .RowTemplate.Height = 32
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .ScrollBars = ScrollBars.None
            .ImeMode = Windows.Forms.ImeMode.Disable
        End With

        '行追加
        Dim dt As New DataTable()
        dt.Columns.Add("All", Type.GetType("System.String"))
        dt.Columns.Add("Hosi", Type.GetType("System.String"))
        dt.Columns.Add("Mori", Type.GetType("System.String"))
        dt.Columns.Add("Sora", Type.GetType("System.String"))
        dt.Columns.Add("Tuki", Type.GetType("System.String"))
        dt.Columns.Add("Hana", Type.GetType("System.String"))
        dt.Columns.Add("Umi", Type.GetType("System.String"))
        Dim row As DataRow = dt.NewRow()
        row(0) = ""
        For i As Integer = 1 To 6
            row(i) = unitArray(i - 1)
        Next
        dt.Rows.Add(row)
        dgvUnit.DataSource = dt

        '幅設定
        With dgvUnit
            For i As Integer = 0 To 6
                With .Columns(i)
                    .Width = 29
                End With
            Next
        End With

    End Sub

    ''' <summary>
    ''' （下の）合議テキストボックスkeyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub consensual2TextBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles consensual2TextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnExcute.Focus()
        End If
    End Sub

    ''' <summary>
    ''' テキストボックスenterイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox_Enter(sender As Object, e As System.EventArgs) Handles facilityManagerTextBox.Enter, consulteeTextBox.Enter, specialistTextBox.Enter, consensual1TextBox.Enter, consensual2TextBox.Enter
        Dim tb As TextBox = CType(sender, TextBox)
        tb.SelectAll()
        mdFlag = True
    End Sub

    ''' <summary>
    ''' テキストボックスマウスダウンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles facilityManagerTextBox.MouseDown, consulteeTextBox.MouseDown, specialistTextBox.MouseDown, consensual1TextBox.MouseDown, consensual2TextBox.MouseDown
        If mdFlag = True Then
            Dim tb As TextBox = CType(sender, TextBox)
            tb.SelectAll()
            mdFlag = False
        End If
    End Sub

    ''' <summary>
    ''' プレビューラジオボタン値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rbtnPreview_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnPreview.CheckedChanged
        If rbtnPreview.Checked = True Then
            Util.putIniString("System", "Printer", "N", TopForm.iniFilePath)
        End If
    End Sub

    ''' <summary>
    ''' 印刷ラジオボタン値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rbtnPrint_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtnPrint.CheckedChanged
        If rbtnPrint.Checked = True Then
            Util.putIniString("System", "Printer", "Y", TopForm.iniFilePath)
        End If
    End Sub

    ''' <summary>
    ''' セルマウスクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvUnit_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUnit.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim unitName As String = Util.checkDBNullValue(dgvUnit.CurrentCell.Value)
            unitLabel.Text = unitName
        End If
    End Sub

    Private Sub dgvUnit_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUnit.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            'Dim Pen As New Pen(Me.dgvUnit.GridColor)
            'With e.CellBounds
            '    .Offset(-1, -1)
            '    e.Graphics.DrawLine(Pen, .Left, .Top, .Left, .Bottom)
            'End With

            Dim pParts = e.PaintParts And (Not DataGridViewPaintParts.Background)
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' 実行ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExcute_Click(sender As System.Object, e As System.EventArgs) Handles btnExcute.Click

    End Sub
End Class