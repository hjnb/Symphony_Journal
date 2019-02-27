Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

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
            .BorderStyle = BorderStyle.FixedSingle
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
                .Width = 469
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
        End With

        dgvShtM(1, 0).Selected = True

    End Sub

    ''' <summary>
    ''' データ表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvShtM(residentName As String, firstDate As String, endDate As String)
        '入力クリア
        clearInput()

        'スクロール等初期位置へ
        dgvShtM.FirstDisplayedScrollingRowIndex = 0
        dgvShtM(1, 0).Selected = True

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
            clearInput()
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
    ''' 和暦を変換
    ''' </summary>
    ''' <param name="warekiStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function formatDateStr(warekiStr As String) As String
        If warekiStr = "" Then
            Return ""
        End If

        Dim kanji As String = Util.getKanji(warekiStr)
        Dim eraNum As String = CInt(warekiStr.Substring(1, 2))
        Dim monthNum As String = CInt(warekiStr.Substring(4, 2))
        Dim dateNum As String = CInt(warekiStr.Substring(7, 2))
        Return kanji & " " & eraNum & " 年 " & monthNum & " 月 " & dateNum & " 日 "
    End Function

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

    ''' <summary>
    ''' 新規ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        clearInput()
        historyListBox.SelectedItem = Nothing
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '入力チェック
        Dim residentName As String = namLabel.Text '入居者名
        Dim firstDate As String = firstYmdBox.getADStr() '利用開始日
        If residentName = "" Then
            MsgBox("利用者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If
        If firstDate = "" Then
            MsgBox("利用開始日を入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '履歴リストが選択されている場合、選択されている日付の既存データ削除
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        If Not IsNothing(historyListBox.SelectedItem) Then
            Dim firstD As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(0))
            Dim endD As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(1))
            cmd.CommandText = "delete from ShtM where Nam='" & residentName & "' and First='" & firstD & "' and End='" & endD & "'"
            cmd.Execute()
        End If

        Dim endDate As String = endYmdBox.getADStr() '利用終了日
        Dim bathDate As String = bathYmdBox.getADStr() '最終入浴日
        Dim benDate As String = benYmdBox.getADStr() '最終排便日
        Dim writeDate As String = dateYmdBox.getADStr() '記載日
        Dim tanto As String = tantoBox.Text '記載者

        '既存データ削除
        cmd.CommandText = "delete from ShtM where Nam='" & residentName & "' and First='" & firstDate & "' and End='" & endDate & "'"
        cmd.Execute()

        '入力の最終行インデックスを取得
        Dim lastInputRowIndex As Integer
        For i As Integer = dgvShtM.Rows.Count - 1 To 0 Step -1
            If dgvShtM("Text", i).Value <> "" Then
                lastInputRowIndex = i
                Exit For
            End If
        Next
        If lastInputRowIndex <= 34 Then
            '35行目まで登録するため
            lastInputRowIndex = 34
        End If

        '登録
        Dim rs As New ADODB.Recordset
        rs.Open("ShtM", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        For i As Integer = 0 To lastInputRowIndex
            rs.AddNew()
            If i = 0 Then
                rs.Fields("Bath").Value = bathDate
                rs.Fields("Ben").Value = benDate
                rs.Fields("Date").Value = writeDate
                rs.Fields("Tanto").Value = tanto
            End If
            rs.Fields("Nam").Value = residentName
            rs.Fields("Gyo").Value = i + 1
            rs.Fields("First").Value = firstDate
            rs.Fields("End").Value = endDate
            rs.Fields("Text").Value = Util.checkDBNullValue(dgvShtM("Text", i).Value)
        Next
        rs.Update()
        rs.Close()
        cnn.Close()

        '再表示
        Dim firstWareki As String = Util.convADStrToWarekiStr(firstDate)
        Dim endWareki As String = Util.convADStrToWarekiStr(endDate)
        historyListBox.Items.Clear()
        historyListBox.Items.AddRange(getHistoryList(residentName).ToArray())
        historyListBox.SelectedItem = firstWareki & "～" & endWareki

    End Sub

    ''' <summary>
    ''' 削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim result As DialogResult = MessageBox.Show("該当期間の記録を抹消しますか", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            If namLabel.Text <> "" AndAlso Not IsNothing(historyListBox.SelectedItem) Then
                Dim firstD As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(0))
                Dim endD As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(1))
                Dim cnn As New ADODB.Connection
                cnn.Open(TopForm.DB_Journal)
                Dim cmd As New ADODB.Command()

                cmd.ActiveConnection = cnn
                cmd.CommandText = "delete from ShtM where Nam='" & namLabel.Text & "' and First='" & firstD & "' and End='" & endD & "'"
                cmd.Execute()

                clearInput()
                historyListBox.Items.Clear()
                historyListBox.Items.AddRange(getHistoryList(namLabel.Text).ToArray())
            End If
        End If
    End Sub

    ''' <summary>
    ''' 印刷ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        '管理者パスワードフォーム表示
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 3)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        Dim residentName As String = namLabel.Text '入居者名
        '入居者名未選択の場合
        If residentName = "" Then
            MsgBox("利用者を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If
        '履歴リストが空の場合
        If historyListBox.Items.Count = 0 Then
            MsgBox("該当がありません。", MsgBoxStyle.Exclamation)
            Return
        End If
        '履歴リスト未選択の場合
        If IsNothing(historyListBox.SelectedItem) Then
            MsgBox("対象年月日を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        'データ取得
        Dim firstDate As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(0))
        Dim endDate As String = Util.convWarekiStrToADStr(historyListBox.Text.Split("～")(1))
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Gyo, [First], [End], Bath, Ben, Date, Tanto, Text from ShtM where Nam='" & residentName & "' And [First]='" & firstDate & "' And [End]='" & endDate & "' order by Gyo"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Dim recordCount As Integer = rs.RecordCount 'レコード数

        '書き込みデータ作成
        Dim bathDate As String = ""
        Dim benDate As String = ""
        Dim writeDate As String = ""
        Dim tanto As String = ""
        Dim dataArray1(34, 0) As String '1枚目用
        Dim dataArray2(47, 0) As String '2枚目用
        If recordCount <= 35 Then
            '1枚目データのみ作成
            While Not rs.EOF
                Dim gyo As Integer = rs.Fields("Gyo").Value
                If gyo = 1 Then
                    bathDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Bath").Value))
                    benDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ben").Value))
                    writeDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Date").Value))
                    tanto = Util.checkDBNullValue(rs.Fields("Tanto").Value)
                End If
                dataArray1(gyo - 1, 0) = Util.checkDBNullValue(rs.Fields("Text").Value)
                rs.MoveNext()
            End While
        Else
            '1枚目、2枚目データ作成
            While Not rs.EOF
                Dim gyo As Integer = rs.Fields("Gyo").Value
                '共通部分
                If gyo = 1 Then
                    bathDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Bath").Value))
                    benDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ben").Value))
                    writeDate = Util.convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Date").Value))
                    tanto = Util.checkDBNullValue(rs.Fields("Tanto").Value)
                End If

                If gyo <= 35 Then
                    '1枚目データ
                    dataArray1(gyo - 1, 0) = Util.checkDBNullValue(rs.Fields("Text").Value)
                ElseIf gyo >= 36 Then
                    '2枚目データ
                    dataArray2(gyo - 1 - 35, 0) = Util.checkDBNullValue(rs.Fields("Text").Value)
                End If
                rs.MoveNext()
            End While
        End If

        'エクセル準備
        Dim objExcel As Excel.Application = CreateObject("Excel.Application")
        Dim objWorkBooks As Excel.Workbooks = objExcel.Workbooks
        Dim objWorkBook As Excel.Workbook = objWorkBooks.Open(TopForm.excelFilePass)
        Dim oSheet As Excel.Worksheet

        objExcel.Calculation = Excel.XlCalculation.xlCalculationManual
        objExcel.ScreenUpdating = False

        If recordCount <= 35 Then
            oSheet = objWorkBook.Worksheets("ｼｮｰﾄｽﾃｲ5改")
            oSheet.range("C4").value = residentName '氏名
            oSheet.range("C6").value = formatDateStr(Util.convADStrToWarekiStr(firstDate)) & "～" & formatDateStr(Util.convADStrToWarekiStr(endDate)) '利用期間
            oSheet.range("C8").value = formatDateStr(bathDate) '最終入浴日
            oSheet.range("G8").value = formatDateStr(benDate) '最終排便日
            oSheet.range("C10").value = formatDateStr(writeDate) '記載日
            oSheet.range("G10").value = tanto '記載者
            oSheet.range("C13", "C47").value = dataArray1 '内容
        Else
            '1枚目
            oSheet = objWorkBook.Worksheets("ｼｮｰﾄｽﾃｲ5-21改")
            oSheet.range("C4").value = residentName '氏名
            oSheet.range("C6").value = formatDateStr(Util.convADStrToWarekiStr(firstDate)) & "～" & formatDateStr(Util.convADStrToWarekiStr(endDate)) '利用期間
            oSheet.range("C8").value = formatDateStr(bathDate) '最終入浴日
            oSheet.range("G8").value = formatDateStr(benDate) '最終排便日
            oSheet.range("C10").value = formatDateStr(writeDate) '記載日
            oSheet.range("G10").value = tanto '記載者
            oSheet.range("C13", "C47").value = dataArray1 '内容

            '2枚目
            oSheet = objWorkBook.Worksheets("ｼｮｰﾄｽﾃｲ5-22改")
            oSheet.range("C4").value = residentName '氏名
            oSheet.range("C6").value = formatDateStr(Util.convADStrToWarekiStr(firstDate)) & "～" & formatDateStr(Util.convADStrToWarekiStr(endDate)) '利用期間
            oSheet.range("C8", "C55").value = dataArray2 '内容
        End If

        objExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        objExcel.ScreenUpdating = True

        '変更保存確認ダイアログ非表示
        objExcel.DisplayAlerts = False

        '印刷
        If TopForm.rbtnPrintout.Checked = True Then
            If recordCount <= 35 Then
                objWorkBook.Worksheets({"ｼｮｰﾄｽﾃｲ5改"}).printOut()
            Else
                objWorkBook.Worksheets({"ｼｮｰﾄｽﾃｲ5-21改", "ｼｮｰﾄｽﾃｲ5-22改"}).printOut()
            End If
        ElseIf TopForm.rbtnPreview.Checked = True Then
            objExcel.Visible = True
            If recordCount <= 35 Then
                objWorkBook.Worksheets({"ｼｮｰﾄｽﾃｲ5改"}).PrintPreview(1)
            Else
                objWorkBook.Worksheets({"ｼｮｰﾄｽﾃｲ5-21改", "ｼｮｰﾄｽﾃｲ5-22改"}).PrintPreview(1)
            End If
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
End Class