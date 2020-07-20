Imports System.Text

Public Class 施設介護支援経過表

    '文字数制限用
    Private Const LIMIT_LENGTH_BYTE As Integer = 100

    '日付ボックス変更制御用
    Private ymdTextChangeFlg As Boolean = True

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海"}

    '施設サービス計画作成者
    Private tantoArray() As String = {"齋藤　久美子", "中野　美加", "河合　哲也"}

    'データテーブル
    Private dtSkei As DataTable

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 施設介護支援経過表_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If Not System.IO.File.Exists(TopForm.dbWorkFilePath) Then
            MsgBox(TopForm.dbWorkFilePath & "が存在しません。iniファイルのDB2Dirに正しいパスを設定して下さい。")
            Me.Close()
            Exit Sub
        End If

        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False

        'dgv初期設定
        initDgvSkei()

        'ユニットをセット
        unitListBox.Items.AddRange(unitArray)

        '施設サービス計画作成者
        tantoBox.Items.AddRange(tantoArray)

        '記載者リスト初期設定
        initWriterList()

        '現在日付をセット
        YmdBox.setADStr(Today.ToString("yyyy/MM/dd"))

        '日付ボックス上下ボタン長押し不可
        YmdBox.canHoldDownButton = False
    End Sub

    ''' <summary>
    ''' データグリッドビュー初期設定
    ''' </summary>
    ''' <remarks></remarks>
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
            .EditMode = DataGridViewEditMode.EditOnEnter
        End With

        '空行追加等
        dtSkei = New DataTable()
        dtSkei.Columns.Add("Text", Type.GetType("System.String"))
        For i = 0 To 46
            Dim row As DataRow = dtSkei.NewRow()
            row(0) = ""
            dtSkei.Rows.Add(row)
        Next
        dgvSkei.DataSource = dtSkei

        '幅設定等
        With dgvSkei
            With .Columns("Text")
                .Width = 610
                .HeaderText = "内　　容"
                .SortMode = DataGridViewColumnSortMode.NotSortable
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
    ''' 経過内容データ表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvSkei(residentName As String, ymd As String)
        clearInput()
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Tanto, Text, Kisai from Skei where Div=" & TopForm.DIV & " And Ymd='" & ymd & "' And Nam='" & residentName & "' order by Gyo"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

        '表示処理
        Dim rowIndex As Integer = 0
        While Not rs.EOF
            If rowIndex = 0 Then
                writerBox.Text = Util.checkDBNullValue(rs.Fields("Kisai").Value)
                tantoBox.Text = Util.checkDBNullValue(rs.Fields("Tanto").Value)
            End If
            dgvSkei("Text", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Text").Value)
            rowIndex += 1
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
        'dgvクリア
        For i As Integer = 0 To dgvSkei.Rows.Count - 1
            dgvSkei("Text", i).Value = ""
        Next

        '記載者クリア
        writerBox.Text = ""
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
            writerBox.Text = nam
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
        rs.Close()
        cn.Close()
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
            Dim ymd As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
            resultList.Add(ymd)
            rs.MoveNext()
        End While
        rs.Close()
        cn.Close()
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
        namLabel.Text = ""
        writerBox.Text = ""
        'dgvクリア
        For i As Integer = 0 To dgvSkei.Rows.Count - 1
            dgvSkei("Text", i).Value = ""
        Next
    End Sub

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
            'dgvクリア
            For i As Integer = 0 To dgvSkei.Rows.Count - 1
                dgvSkei("Text", i).Value = ""
            Next
        End If
    End Sub

    ''' <summary>
    ''' 経過履歴リスト値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub historyListBox_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles historyListBox.SelectedValueChanged
        Dim residentName As String = residentListBox.Text
        Dim ymd As String = historyListBox.Text
        If residentName <> "" AndAlso ymd <> "" Then
            YmdBox.setADStr(ymd)
        End If
    End Sub

    ''' <summary>
    ''' 日付ボックス値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub YmdBox_YmdTextChange(sender As Object, e As System.EventArgs) Handles YmdBox.YmdTextChange
        Dim residentName As String = namLabel.Text
        Dim ymd As String = YmdBox.getADStr()
        If residentName <> "" AndAlso ymd <> "" Then
            displayDgvSkei(residentName, ymd)
        End If
    End Sub

    ''' <summary>
    ''' EditingControlShowingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvSkei_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvSkei.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)

        'イベントハンドラを削除、追加
        RemoveHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
        AddHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
    End Sub

    ''' <summary>
    ''' dgvのKeyPressイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' 行挿入ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowInsert_Click(sender As System.Object, e As System.EventArgs) Handles btnRowInsert.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvSkei.CurrentCell), -1, dgvSkei.CurrentRow.Index)
        If selectedRowIndex <> -1 Then
            Dim row As DataRow = dtSkei.NewRow()
            dtSkei.Rows.InsertAt(row, selectedRowIndex) '追加
            dtSkei.Rows.RemoveAt(47) '一番下の行削除
        End If
    End Sub

    ''' <summary>
    ''' 行削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvSkei.CurrentCell), -1, dgvSkei.CurrentRow.Index)
        If selectedRowIndex <> -1 Then
            dtSkei.Rows.RemoveAt(selectedRowIndex) '選択行削除
            Dim row As DataRow = dtSkei.NewRow()
            dtSkei.Rows.Add(row) '追加
        End If
    End Sub

    ''' <summary>
    ''' クリアボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        For i As Integer = 0 To dgvSkei.Rows.Count - 1
            dgvSkei("Text", i).Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim tanto As String = tantoBox.Text 'サービス計画作成者
        If tanto = "" Then
            MsgBox("作成者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim nam As String = namLabel.Text '入居者名
        If nam = "" Then
            MsgBox("入居者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '内容が空の場合は何もしない
        Dim emptyFlg As Boolean = True
        For i As Integer = 0 To dgvSkei.Rows.Count - 1
            If dgvSkei("Text", i).Value <> "" Then
                emptyFlg = False
                Exit For
            End If
        Next
        If emptyFlg Then
            Return
        End If

        '既存データ削除
        Dim ymd As String = YmdBox.getADStr() '日付
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        cmd.CommandText = "delete from Skei where Div=" & TopForm.DIV & " and Ymd='" & ymd & "' and Nam='" & nam & "'"
        cmd.Execute()

        '登録
        Dim rs As New ADODB.Recordset
        rs.Open("Skei", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        For i As Integer = 0 To dgvSkei.Rows.Count - 1
            Dim text As String = Util.checkDBNullValue(dgvSkei("Text", i).Value)
            If text <> "" Then
                rs.AddNew()
                rs.Fields("Div").Value = TopForm.DIV
                rs.Fields("Nam").Value = nam
                rs.Fields("Ymd").Value = ymd
                rs.Fields("Tanto").Value = tanto
                rs.Fields("Kisai").Value = writerBox.Text
                rs.Fields("Gyo").Value = i + 1
                rs.Fields("Text").Value = text
            End If
        Next
        rs.Update()
        rs.Close()
        cnn.Close()

        '再表示
        displayDgvSkei(nam, ymd)
        historyListBox.Items.Clear()
        historyListBox.Items.AddRange(getHistoryList(nam).ToArray())

    End Sub

    ''' <summary>
    ''' 削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim nam As String = namLabel.Text '入居者名
        If nam = "" Then
            MsgBox("入居者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim result As DialogResult = MessageBox.Show("当日付の記録を抹消しますか？", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim ymd As String = YmdBox.getADStr() '日付
            Dim cnn As New ADODB.Connection
            cnn.Open(TopForm.DB_Journal)
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = cnn
            cmd.CommandText = "delete from Skei where Div=" & TopForm.DIV & " and Ymd='" & ymd & "' and Nam='" & nam & "'"
            cmd.Execute()
            cnn.Close()

            '再表示
            displayDgvSkei(nam, ymd)
            historyListBox.Items.Clear()
            historyListBox.Items.AddRange(getHistoryList(nam).ToArray())
        End If
    End Sub

    ''' <summary>
    ''' 印刷ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Dim nam As String = namLabel.Text '入居者名
        If nam = "" Then
            MsgBox("入居者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '経過履歴が無い場合は終了
        If historyListBox.Items.Count <= 0 Then
            MsgBox("該当データなし。", MsgBoxStyle.Exclamation)
            Return
        End If

        '管理者パスワードフォーム表示
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 3)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        '経過履歴の最古と最新の日付取得
        Dim oldDate As String = historyListBox.Items(historyListBox.Items.Count - 1)
        Dim newDate As String = historyListBox.Items(0)

        '印刷範囲フォーム表示
        Dim printForm As 印刷範囲 = New 印刷範囲(nam, oldDate, newDate)
        printForm.ShowDialog()
    End Sub
End Class