Public Class 施設介護支援経過表

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 施設介護支援経過表_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False

        'dgv初期設定
        initDgvSkei()

        'ユニットをセット
        unitListBox.Items.AddRange(unitArray)

        '記載者リスト初期設定
        initWriterList()

        '現在日付をセット
        YmdBox.setADStr(Today.ToString("yyyy/MM/dd"))
    End Sub

    Private Sub initDgvSkei()
        Util.EnableDoubleBuffering(dgvSkei)

        With dgvSkei
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
            .ColumnHeadersHeight = 18
            .RowTemplate.Height = 13
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
        For i = 0 To 46
            Dim row As DataRow = dt.NewRow()
            row(0) = ""
            dt.Rows.Add(row)
        Next
        dgvSkei.DataSource = dt

        '幅設定等
        With dgvSkei
            With .Columns("Text")
                .Width = 610
                .HeaderText = "内　　容"
            End With
        End With
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
    ''' 対象ユニットの入居者リストを取得
    ''' </summary>
    ''' <param name="unitName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getResidentList(unitName As String) As List(Of String)
        Dim resultList As New List(Of String) '取得結果リスト
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
    ''' 対象の入居者の経過履歴リスト取得
    ''' </summary>
    ''' <param name="residentName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getHistoryList(residentName As String) As List(Of String)
        Dim resultList As New List(Of String) '取得結果リスト
        '対象の入居者の経過履歴リスト作成
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select distinct Ymd from Skei where Div=" & TopForm.DIV & " And Nam='" & residentName & "' order by Ymd Desc"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        While Not rs.EOF
            Dim wareki As String = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ymd").Value))
            resultList.Add(wareki)
            rs.MoveNext()
        End While
        Return resultList
    End Function

    ''' <summary>
    ''' ユニットリスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub unitListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles unitListBox.SelectedValueChanged
        Dim selectedUnitName As String = unitListBox.Text
        residentListBox.Items.Clear()
        residentListBox.Items.AddRange(getResidentList(selectedUnitName).ToArray())
        historyListBox.Items.Clear()
    End Sub

    ''' <summary>
    ''' 入居者リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub residentListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles residentListBox.SelectedValueChanged
        Dim nam As String = residentListBox.Text
        If nam <> "" Then
            historyListBox.Items.Clear()
            historyListBox.Items.AddRange(getHistoryList(nam).ToArray())
        End If
    End Sub
End Class