Imports System.Data.OleDb
Imports System.Text

Public Class 申し送り

    '入力文字数制限用
    Private Const LIMIT_LENGTH_BYTE As Integer = 80

    '行選択フラグ
    Private selectedRowFlg As Boolean = False

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 申し送り_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If Not System.IO.File.Exists(TopForm.dbWorkFilePath) Then
            MsgBox(TopForm.dbWorkFilePath & "が存在しません。iniファイルのDB2Dirに正しいパスを設定して下さい。")
            Me.Close()
            Exit Sub
        End If

        Me.WindowState = FormWindowState.Maximized

        '現在日付セット
        YmdBox.setADStr(Today.ToString("yyyy/MM/dd"))

        '現在時刻セット
        Dim hh As String = DateTime.Now.ToString("HH")
        Dim mm As String = DateTime.Now.ToString("mm")
        HmBox.setTime(hh, mm)

        '記入者リストボックス初期設定
        initWriterList()

        'dgv初期設定
        initDgvInput() '上の
        initDgvRead() '下の

        'データ表示
        displayDgvRead()
    End Sub

    ''' <summary>
    ''' 記入者リストボックス初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initWriterList()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String

        'workの勤務表から当月の勤務者（パート除く）の名前取得
        Dim ym As String = Today.ToString("yyyy/MM")
        cnn.Open(TopForm.DB_Work)
        sql = "select Nam from KinD where Ym='" & ym & "' And Rdr<>'' order by Seq2, Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        While Not rs.EOF
            writerListBox.Items.Add(Util.checkDBNullValue(rs.Fields("Nam").Value))
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()

        'EtcMから名前取得
        cnn.Open(TopForm.DB_Journal)
        sql = "select Nam from EtcM order by Num"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        While Not rs.EOF
            writerListBox.Items.Add(Util.checkDBNullValue(rs.Fields("Nam").Value))
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()
    End Sub

    ''' <summary>
    ''' 記入者リストボックス値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub writerListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles writerListBox.SelectedValueChanged
        Dim nam As String = writerListBox.SelectedItem
        If nam <> "" Then
            writerLabel.Text = nam
        End If
    End Sub

    ''' <summary>
    ''' dgv(上)初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvInput()
        Util.EnableDoubleBuffering(dgvInput)

        With dgvInput
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .BorderStyle = BorderStyle.None
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .RowHeadersVisible = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersVisible = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowTemplate.Height = 15.5
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '空行追加等
        Dim dt As New DataTable()
        dt.Columns.Add("Text", Type.GetType("System.String"))
        For i = 0 To 29
            Dim row As DataRow = dt.NewRow()
            row(0) = ""
            dt.Rows.Add(row)
        Next
        dgvInput.DataSource = dt

        '幅設定等
        With dgvInput
            With .Columns("Text")
                .Width = 478
            End With
        End With
    End Sub

    ''' <summary>
    ''' dgv(下)初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvRead()
        Util.EnableDoubleBuffering(dgvRead)

        With dgvRead
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .BorderStyle = BorderStyle.None
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .RowHeadersVisible = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersHeight = 19
            .RowTemplate.Height = 15
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.ForeColor = Color.Black
            .DefaultCellStyle.SelectionBackColor = Color.Black
            .DefaultCellStyle.SelectionForeColor = Color.White
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
    End Sub

    ''' <summary>
    ''' データ表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvRead()
        dgvRead.Columns.Clear()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Ymd, Hm, Gyo, Text, Tanto From Rprt where Div=" & TopForm.DIV & " order by Ymd Desc, Hm, Gyo"
        cnn.Open(TopForm.DB_Journal)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "Rprt")
        cnn.Close()

        '西暦を和暦に変換
        For Each row As DataRow In ds.Tables("Rprt").Rows
            row("Ymd") = Util.convADStrToWarekiStr(row("Ymd"))
        Next

        '表示
        dgvRead.DataSource = ds.Tables("Rprt")

        '列設定等
        With dgvRead

            .Columns("Gyo").Visible = False

            With .Columns("Ymd")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .SortMode = DataGridViewColumnSortMode.NotSortable
                .HeaderText = "年月日"
                .Width = 75
            End With

            With .Columns("Hm")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .SortMode = DataGridViewColumnSortMode.NotSortable
                .HeaderText = "時間"
                .Width = 45
            End With

            With .Columns("Text")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .SortMode = DataGridViewColumnSortMode.NotSortable
                .HeaderText = "内 容"
                .Width = 477
            End With

            With .Columns("Tanto")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .SortMode = DataGridViewColumnSortMode.NotSortable
                .HeaderText = "記載者"
                .Width = 95
            End With
        End With

        selectedRowFlg = False
    End Sub

    ''' <summary>
    ''' dgv(下)CellFormattingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvRead_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvRead.CellFormatting
        If dgvRead.Columns(e.ColumnIndex).Name = "Ymd" Then
            '年月日のグループ化
            If e.RowIndex > 0 AndAlso dgvRead(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        ElseIf dgvRead.Columns(e.ColumnIndex).Name = "Hm" Then
            '時間のグループ化
            If e.RowIndex > 0 AndAlso dgvRead("Ymd", e.RowIndex - 1).Value = dgvRead("Ymd", e.RowIndex).Value AndAlso dgvRead(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' dgv(上)CellPaintingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvInput_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvInput.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            Dim pParts As DataGridViewPaintParts = e.PaintParts And Not DataGridViewPaintParts.Background
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' dgv(上)EditingControlShowingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvInput_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvInput.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)

        'イベントハンドラを削除、追加
        RemoveHandler editTextBox.KeyPress, AddressOf dgvInputTextBox_KeyPress
        AddHandler editTextBox.KeyPress, AddressOf dgvInputTextBox_KeyPress
    End Sub

    ''' <summary>
    ''' dgv(上)KeyPressイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvInputTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
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

    ''' <summary>
    ''' dgv(下)セルマウスクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvRead_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvRead.CellMouseClick
        If e.RowIndex >= 0 Then
            selectedRowFlg = True
            Dim hm As String = Util.checkDBNullValue(dgvRead("Hm", e.RowIndex).FormattedValue)
            Dim readRowIndex As Integer = e.RowIndex
            While hm = ""
                readRowIndex -= 1
                hm = Util.checkDBNullValue(dgvRead("Hm", readRowIndex).FormattedValue)
            End While
            Dim wareki As String = Util.checkDBNullValue(dgvRead("Ymd", readRowIndex).Value)
            Dim writer As String = Util.checkDBNullValue(dgvRead("Tanto", readRowIndex).Value)

            '内容クリア
            For Each row As DataGridViewRow In dgvInput.Rows
                row.Cells("Text").Value = ""
            Next

            '値セット
            Dim inputRowIndex As Integer
            While (wareki = Util.checkDBNullValue(dgvRead("Ymd", readRowIndex).Value)) AndAlso (hm = Util.checkDBNullValue(dgvRead("Hm", readRowIndex).Value))
                dgvInput("Text", inputRowIndex).Value = Util.checkDBNullValue(dgvRead("Text", readRowIndex).Value)
                inputRowIndex += 1
                readRowIndex += 1
            End While
            YmdBox.setWarekiStr(wareki)
            HmBox.setTime(hm.Substring(0, 2), hm.Substring(3, 2))
            writerLabel.Text = writer

            '
            dgvInput.FirstDisplayedScrollingRowIndex = 0
            dgvInput(0, 0).Selected = True
            dgvInput.Focus()
        End If
    End Sub

    ''' <summary>
    ''' 入力内容クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub clearInput()
        'dgv内容クリア
        For i As Integer = 0 To dgvInput.Rows.Count - 1
            dgvInput("Text", i).Value = ""
        Next
        dgvInput.FirstDisplayedScrollingRowIndex = 0
        dgvInput(0, 0).Selected = True
        dgvInput.Focus()
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim writer As String = writerLabel.Text '記載者
        If writer = "" Then
            MsgBox("記載者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '入力チェック
        Dim emptyFlg As Boolean = True
        Dim lastInputRowIndex As Integer '最終入力行
        For i As Integer = dgvInput.Rows.Count - 1 To 0 Step -1
            If Util.checkDBNullValue(dgvInput("Text", i).Value) <> "" Then
                lastInputRowIndex = i
                emptyFlg = False
                Exit For
            End If
        Next
        If emptyFlg Then
            MsgBox("登録内容が空です。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim ymd As String = YmdBox.getADStr() '日付(yyyy/MM/dd)
        Dim hm As String = HmBox.getTime() '時間(HH:mm)

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM Rprt WHERE Div=" & TopForm.DIV & " and Ymd='" & ymd & "' and Hm='" & hm & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount = 0 Then
            rs.Close()
            '新規登録
            rs.Open("Rprt", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            For i As Integer = 0 To lastInputRowIndex
                rs.AddNew()
                rs.Fields("Div").Value = TopForm.DIV
                rs.Fields("Ymd").Value = ymd
                rs.Fields("Hm").Value = hm
                rs.Fields("Gyo").Value = i + 1
                rs.Fields("Text").Value = Util.checkDBNullValue(dgvInput("Text", i).Value)
                If i = 0 Then
                    rs.Fields("Tanto").Value = writer
                End If
            Next
            rs.Update()
            rs.Close()
            cnn.Close()

            'クリア
            clearInput()

            '再表示
            displayDgvRead()
        Else
            rs.Close()
            '変更登録
            Dim result As DialogResult = MessageBox.Show("変更登録してよろしいですか？", "登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = Windows.Forms.DialogResult.Yes Then
                '既存データ削除
                Dim cmd As New ADODB.Command()
                cmd.ActiveConnection = cnn
                cmd.CommandText = "delete from Rprt where Div=" & TopForm.DIV & " and Ymd='" & ymd & "' and Hm='" & hm & "'"
                cmd.Execute()

                '登録
                rs.Open("Rprt", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
                For i As Integer = 0 To lastInputRowIndex
                    rs.AddNew()
                    rs.Fields("Div").Value = TopForm.DIV
                    rs.Fields("Ymd").Value = ymd
                    rs.Fields("Hm").Value = hm
                    rs.Fields("Gyo").Value = i + 1
                    rs.Fields("Text").Value = Util.checkDBNullValue(dgvInput("Text", i).Value)
                    If i = 0 Then
                        rs.Fields("Tanto").Value = writer
                    End If
                Next
                rs.Update()
                rs.Close()
                cnn.Close()

                'クリア
                clearInput()

                '再表示
                displayDgvRead()
            Else
                cnn.Close()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 行削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        If selectedRowFlg = False Then
            MsgBox("削除データが選択されていません。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim ymd As String = YmdBox.getADStr() '日付
            Dim hm As String = HmBox.getTime() '時間
            Dim cnn As New ADODB.Connection
            cnn.Open(TopForm.DB_Journal)
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = cnn
            cmd.CommandText = "delete from Rprt where Div=" & TopForm.DIV & " and Ymd='" & ymd & "' and Hm='" & hm & "'"
            cmd.Execute()
            cnn.Close()

            'クリア
            clearInput()

            '再表示
            displayDgvRead()
        End If
    End Sub
End Class