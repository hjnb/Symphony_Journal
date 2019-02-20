Public Class SS生活の様子

    'ショートステイのユニット名
    Private Const SS_UNIT_NAME As String = "海"

    'データグリッドビュー用データテーブル
    Private dt As DataTable

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SS生活の様子_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.KeyPreview = True

        '入居者リストボックス初期設定
        initResidentListBox()

        'dgv初期設定
        initDgvShtM()
    End Sub

    Private Sub SS生活の様子_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If Not dgvShtM.Focused Then
            If e.KeyCode = Keys.Enter Then
                If e.Control = False Then
                    Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 入居者リストボックス初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initResidentListBox()
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Nam from UsrM where Dsp=1 And Unt='" & SS_UNIT_NAME & "' order by Kana"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        While Not rs.EOF
            residentListBox.Items.Add(Util.checkDBNullValue(rs.Fields("Nam").Value))
            rs.MoveNext()
        End While
        cn.Close()
    End Sub

    ''' <summary>
    ''' データグリッドビュー初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvShtM()
        Util.EnableDoubleBuffering(dgvShtM)

        With dgvShtM
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
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '列追加、空の行追加
        dt = New DataTable()
        dt.Columns.Add("Title", Type.GetType("System.String"))
        dt.Columns.Add("Text", Type.GetType("System.String"))
        Dim titleDic As New Dictionary(Of Integer, String) From {{0, "食　　　事"}, {5, "排　　　泄"}, {10, "入　　　浴"}, {15, "夜間の入眠"}, {20, "全　　　般"}, {35, "(本行より次頁)"}}
        For i As Integer = 0 To 82
            Dim row As DataRow = dt.NewRow()
            If i = 0 OrElse i = 5 OrElse i = 10 OrElse i = 15 OrElse i = 20 OrElse i = 35 Then
                row(0) = titleDic(i)
            Else
                row(0) = ""
            End If
            row(1) = ""
            dt.Rows.Add(row)
        Next

        '表示
        dgvShtM.DataSource = dt

        '幅設定等
        With dgvShtM
            With .Columns("Title")
                .DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .ReadOnly = True
                .Width = 80
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
            With .Columns("Text")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Font = New Font("ＭＳ ゴシック", 9)
                .Width = 462
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
        End With

    End Sub

    ''' <summary>
    ''' データ表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvShtM(residentName As String, firstDate As String, endDate As String)
        '入力クリア
        clearInput()

        'データ取得、表示
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Gyo, [First], [End], Bath, Ben, Date, Tanto, Text from ShtM where Nam='" & residentName & "' And [First]='" & firstDate & "' And [End]='" & endDate & "' order by Gyo"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        Dim count As Integer = 0
        While Not rs.EOF
            If count = 0 Then
                firstYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("First").Value))
                endYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("End").Value))
                count = 1
            End If
            Dim gyo As Integer = rs.Fields("Gyo").Value
            If gyo = 1 Then
                bathYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("Bath").Value))
                benYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("Ben").Value))
                dateYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("date").Value))
                tantoBox.Text = Util.checkDBNullValue(rs.Fields("Tanto").Value)
            End If
            dgvShtM("Text", gyo - 1).Value = Util.checkDBNullValue(rs.Fields("Text").Value)
            rs.MoveNext()
        End While
        rs.Close()
        cn.Close()
    End Sub

    ''' <summary>
    ''' 入力内容クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub clearInput()
        firstYmdBox.clearText() '利用期間from
        endYmdBox.clearText() '利用期間to
        bathYmdBox.clearText() '最終入浴
        benYmdBox.clearText() '最終排便
        dateYmdBox.clearText() '記載日
        tantoBox.Text = "" '記載者

        'dgv内容クリア
        For i As Integer = 0 To dgvShtM.Rows.Count - 1
            dgvShtM("Text", i).Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 履歴リスト取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getHistoryList(residentName As String) As List(Of String)
        Dim resultList As New List(Of String) '取得結果リスト
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select distinct [First], [End] from ShtM where Nam='" & residentName & "' order by First Desc"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        While Not rs.EOF
            Dim listItem As String = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("First").Value)) & "～" & Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("End").Value))
            resultList.Add(listItem)
            rs.MoveNext()
        End While
        cn.Close()
        Return resultList
    End Function

    ''' <summary>
    ''' 入居者リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub residentListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles residentListBox.SelectedValueChanged
        Dim nam As String = residentListBox.Text
        If nam <> "" AndAlso nam <> namLabel.Text Then
            namLabel.Text = nam
            historyListBox.Items.Clear()
            historyListBox.Items.AddRange(getHistoryList(nam).ToArray())
        End If
    End Sub

    ''' <summary>
    ''' 履歴リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub historyListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles historyListBox.SelectedValueChanged
        Dim selectedText As String = historyListBox.Text
        If selectedText <> "" Then
            Dim firstDate As String = Util.convWarekiStrToADStr(selectedText.Split("～")(0))
            Dim endDate As String = Util.convWarekiStrToADStr(selectedText.Split("～")(1))
            displayDgvShtM(namLabel.Text, firstDate, endDate)
        End If
    End Sub

    ''' <summary>
    ''' テキストクリアボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnTextClear_Click(sender As System.Object, e As System.EventArgs) Handles btnTextClear.Click
        Dim result As DialogResult = MessageBox.Show("テキストをクリアしますか？", "クリア", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            clearInput()
        End If
    End Sub

    ''' <summary>
    ''' 行挿入ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowInsert_Click(sender As System.Object, e As System.EventArgs) Handles btnRowInsert.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvShtM.CurrentRow), -1, dgvShtM.CurrentRow.Index)
        If selectedRowIndex <> -1 Then
            Dim loopStartCount As Integer = 0
            If 0 <= selectedRowIndex AndAlso selectedRowIndex <= 4 Then
                loopStartCount = 3
            ElseIf 5 <= selectedRowIndex AndAlso selectedRowIndex <= 9 Then
                loopStartCount = 8
            ElseIf 10 <= selectedRowIndex AndAlso selectedRowIndex <= 14 Then
                loopStartCount = 13
            ElseIf 15 <= selectedRowIndex AndAlso selectedRowIndex <= 19 Then
                loopStartCount = 18
            ElseIf 20 <= selectedRowIndex AndAlso selectedRowIndex <= 82 Then
                loopStartCount = 81
            End If
            For i As Integer = loopStartCount To selectedRowIndex Step -1
                dt.Rows(i + 1).Item("Text") = dt.Rows(i).Item("Text")
            Next
            dt.Rows(selectedRowIndex).Item("Text") = ""
        End If
    End Sub

    ''' <summary>
    ''' 行削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvShtM.CurrentRow), -1, dgvShtM.CurrentRow.Index)
        If selectedRowIndex <> -1 Then
            Dim loopEndCount As Integer = 0
            If 0 <= selectedRowIndex AndAlso selectedRowIndex <= 4 Then
                loopEndCount = 4
            ElseIf 5 <= selectedRowIndex AndAlso selectedRowIndex <= 9 Then
                loopEndCount = 9
            ElseIf 10 <= selectedRowIndex AndAlso selectedRowIndex <= 14 Then
                loopEndCount = 14
            ElseIf 15 <= selectedRowIndex AndAlso selectedRowIndex <= 19 Then
                loopEndCount = 19
            ElseIf 20 <= selectedRowIndex AndAlso selectedRowIndex <= 82 Then
                loopEndCount = 82
            End If
            For i As Integer = selectedRowIndex To loopEndCount - 1
                dt.Rows(i).Item("Text") = dt.Rows(i + 1).Item("Text")
            Next
            dt.Rows(loopEndCount).Item("Text") = ""
        End If
    End Sub
End Class