Imports System.Data.OleDb

Public Class 申し送り

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 申し送り_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Nam from EtcM order by Num"
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
        dgvRead.DataSource = ds.Tables("Rprt")
        cnn.Close()

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
    End Sub

    Private Sub dgvRead_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvRead.CellFormatting
        If dgvRead.Columns(e.ColumnIndex).Name = "Ymd" Then
            '年月日のグループ化
            If e.RowIndex > 0 AndAlso dgvRead(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        ElseIf dgvRead.Columns(e.ColumnIndex).Name = "Hm" Then
            '曜日の表示設定,グループ化
            If e.RowIndex > 0 AndAlso dgvRead(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If
    End Sub
End Class