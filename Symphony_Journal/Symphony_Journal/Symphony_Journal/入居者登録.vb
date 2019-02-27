Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class 入居者登録

    'ユニット
    Private unitArray() As String = {"星", "森", "空", "月", "花", "海", "虹", "光", "丘", "風", "雪"}

    '選択行のAutono保持用
    Private selectedAutono As Integer = -1

    ''' <summary>
    ''' 行ヘッダーのカレントセルを表す三角マークを非表示に設定する為のクラス。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class dgvRowHeaderCell

        'DataGridViewRowHeaderCell を継承
        Inherits DataGridViewRowHeaderCell

        'DataGridViewHeaderCell.Paint をオーバーライドして行ヘッダーを描画
        Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, _
           ByVal cellBounds As Rectangle, ByVal rowIndex As Integer, ByVal cellState As DataGridViewElementStates, _
           ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, _
           ByVal cellStyle As DataGridViewCellStyle, ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
           ByVal paintParts As DataGridViewPaintParts)
            '標準セルの描画からセル内容の背景だけ除いた物を描画(-5)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, _
                     formattedValue, errorText, cellStyle, advancedBorderStyle, _
                     Not DataGridViewPaintParts.ContentBackground)
        End Sub

    End Class

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        InitializeComponent()

        Me.StartPosition = FormStartPosition.CenterScreen
        Me.KeyPreview = True
        Me.MaximizeBox = False
    End Sub

    ''' <summary>
    ''' Loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 入居者登録_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '初期設定
        initDgvResident()
        initTextBox()

        'データグリッドビュー表示
        displayDgvResident()
    End Sub

    ''' <summary>
    ''' KeyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 入居者登録_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    ''' <summary>
    ''' テキストボックス等初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initTextBox()
        'ユニット
        unitBox.Items.AddRange(unitArray)
        unitBox.ImeMode = Windows.Forms.ImeMode.Alpha

        '入居者名
        namBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'ｶﾅ
        kanaBox.ImeMode = Windows.Forms.ImeMode.KatakanaHalf

        'ON,OFF
        rbtnDisplayOn.Checked = True
    End Sub

    ''' <summary>
    ''' データグリッドビュー初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvResident()
        Util.EnableDoubleBuffering(dgvResident)

        With dgvResident
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .ReadOnly = True
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 20
            .RowTemplate.Height = 16
            .RowHeadersWidth = 35
            .EnableHeadersVisualStyles = False
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
        End With
    End Sub

    ''' <summary>
    ''' データグリッドビュー表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub displayDgvResident()
        dgvResident.Columns.Clear()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Autono, Unt, Nam, Kana, Dsp from UsrM order by Unt, Kana"
        cnn.Open(TopForm.DB_Journal)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "UsrM")
        dgvResident.DataSource = ds.Tables("UsrM")
        cnn.Close()

        settingDgvResident()
        Me.ActiveControl = namBox
        selectedAutono = -1
    End Sub

    ''' <summary>
    ''' dgv表示設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub settingDgvResident()
        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgvResident.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        With dgvResident

            '非表示
            .Columns("Autono").Visible = False

            With .Columns("Unt")
                .HeaderText = "ﾕﾆｯﾄ"
                .Width = 45
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Nam")
                .HeaderText = "入居者名"
                .Width = 95
            End With

            With .Columns("Kana")
                .HeaderText = "ｶﾅ"
                .Width = 95

            End With

            With .Columns("Dsp")
                .HeaderText = "表示"
                .Width = 45
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
        End With
    End Sub

    ''' <summary>
    ''' 入力内容クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub inputClear()
        unitBox.Text = ""
        namBox.Text = ""
        kanaBox.Text = ""
        rbtnDisplayOn.Checked = True
    End Sub

    ''' <summary>
    ''' CellFormattingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvResident_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvResident.CellFormatting
        If dgvResident.Columns(e.ColumnIndex).Name = "Dsp" Then
            If Util.checkDBNullValue(e.Value) = "1" Then
                e.Value = "○"
            Else
                e.Value = ""
            End If
            e.FormattingApplied = True
        End If
    End Sub

    ''' <summary>
    ''' セルマウスクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvResident_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvResident.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim unit As String = Util.checkDBNullValue(dgvResident("Unt", e.RowIndex).Value) 'ユニット
            Dim nam As String = Util.checkDBNullValue(dgvResident("Nam", e.RowIndex).Value) '入居者名
            Dim kana As String = Util.checkDBNullValue(dgvResident("Kana", e.RowIndex).Value) 'ｶﾅ
            Dim dsp As Integer = Util.checkDBNullValue(dgvResident("Dsp", e.RowIndex).Value) '表示
            selectedAutono = Util.checkDBNullValue(dgvResident("Autono", e.RowIndex).Value) '表示

            '値設定
            unitBox.Text = unit
            namBox.Text = nam
            kanaBox.Text = kana
            If dsp = 1 Then
                rbtnDisplayOn.Checked = True
            Else
                rbtnDisplayOff.Checked = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' CellPaintingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvResident_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvResident.CellPainting
        '行ヘッダーかどうか調べる
        If e.ColumnIndex < 0 AndAlso e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                e.CellStyle.ForeColor, _
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' 列ヘッダーダブルクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvResident_ColumnHeaderMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvResident.ColumnHeaderMouseDoubleClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim targetColumn As DataGridViewColumn = dgv.Columns(e.ColumnIndex) '選択列
        dgv.Sort(targetColumn, System.ComponentModel.ListSortDirection.Ascending) '昇順でソート
    End Sub

    ''' <summary>
    ''' 新規ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnNewRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnNewRegist.Click
        Dim unit As String = unitBox.Text 'ユニット
        Dim nam As String = namBox.Text '入居者名
        Dim kana As String = kanaBox.Text 'カナ
        Dim dsp As Integer = If(rbtnDisplayOn.Checked, 1, 0) '表示(ON:1,OFF:0)

        'ユニットチェック
        If unit = "" Then
            MsgBox("ﾕﾆｯﾄを選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '入居者名チェック
        If nam = "" Then
            MsgBox("漢字氏名を入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        'ｶﾅチェック
        If kana = "" Then
            MsgBox("ｶﾅ氏名を入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim result As DialogResult = MessageBox.Show("新規登録してよろしいですか？", "登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim cn As New ADODB.Connection()
            cn.Open(TopForm.DB_Journal)
            Dim rs As New ADODB.Recordset
            rs.Open("UsrM", cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

            'レコード追加
            rs.AddNew()
            rs.Fields("Unt").Value = unit
            rs.Fields("Nam").Value = nam
            rs.Fields("Kana").Value = kana
            rs.Fields("Dsp").Value = dsp
            rs.Update()

            rs.Close()
            cn.Close()

            '再表示
            inputClear()
            displayDgvResident()
        End If
    End Sub

    ''' <summary>
    ''' 変更ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnChange_Click(sender As System.Object, e As System.EventArgs) Handles btnChange.Click
        'dgvの行をｸﾘｯｸで選択していない場合は新規登録処理へ
        If selectedAutono < 0 Then
            btnNewRegist.PerformClick()
            Return
        End If

        Dim unit As String = unitBox.Text 'ユニット
        Dim nam As String = namBox.Text '入居者名
        Dim kana As String = kanaBox.Text 'カナ
        Dim dsp As Integer = If(rbtnDisplayOn.Checked, 1, 0) '表示(ON:1,OFF:0)

        'ユニットチェック
        If unit = "" Then
            MsgBox("ﾕﾆｯﾄを選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '入居者名チェック
        If nam = "" Then
            MsgBox("漢字氏名を入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        'ｶﾅチェック
        If kana = "" Then
            MsgBox("ｶﾅ氏名を入力して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim result As DialogResult = MessageBox.Show("変更登録してよろしいですか？", "登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim cn As New ADODB.Connection()
            cn.Open(TopForm.DB_Journal)
            Dim sql As String = "select Autono, Unt, Nam, Kana, Dsp from UsrM where Autono=" & selectedAutono
            Dim rs As New ADODB.Recordset
            rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

            'レコード更新
            rs.Fields("Unt").Value = unit
            rs.Fields("Nam").Value = nam
            rs.Fields("Kana").Value = kana
            rs.Fields("Dsp").Value = dsp
            rs.Update()

            rs.Close()
            cn.Close()

            '再表示
            inputClear()
            displayDgvResident()
        End If
    End Sub

    ''' <summary>
    ''' 削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        'dgvの行をｸﾘｯｸで選択していない場合
        If selectedAutono < 0 Then
            MsgBox("選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim cn As New ADODB.Connection()
            cn.Open(TopForm.DB_Journal)
            Dim sql As String = "select Autono from UsrM where Autono=" & selectedAutono
            Dim rs As New ADODB.Recordset
            rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

            'レコード削除
            rs.Delete()
            rs.Update()

            rs.Close()
            cn.Close()

            '再表示
            inputClear()
            displayDgvResident()
        End If

    End Sub

    ''' <summary>
    ''' 印刷ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Dim objExcel As Excel.Application
        Dim objWorkBooks As Excel.Workbooks
        Dim objWorkBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheet = objWorkBook.Worksheets("入居者")

        objExcel.Calculation = Excel.XlCalculation.xlCalculationManual
        objExcel.ScreenUpdating = False

        '既存文字削除
        oSheet.range("E2").value = ""
        oSheet.range("B4").value = ""
        oSheet.range("C4").value = ""
        oSheet.range("D4").value = ""
        oSheet.range("E4").value = ""
        oSheet.range("F4").value = ""

        '日付
        Dim dateStr As String = Util.convADStrToWarekiStr(Today.ToString("yyyy/MM/dd")) '和暦表記日付
        Dim nowTime As String = DateTime.Now.ToString("HH:mm")
        oSheet.range("E2").value = dateStr & " " & nowTime

        Dim rCount As Integer = dgvResident.Rows.Count 'レコード数
        Dim loopCount As Integer = rCount \ 60

        '2枚目以降の枠準備
        If rCount > 60 Then
            For i As Integer = 0 To loopCount - 1
                Dim xlPasteRange As Excel.Range = oSheet.Range("A" & (65 + (64 * i))) 'ペースト先
                oSheet.rows("1:64").copy(xlPasteRange)
            Next
        End If

        'データ配列作成
        Dim rowIndex As Integer = 0
        Dim dataArrayList As New List(Of String(,))
        For i As Integer = 0 To loopCount
            Dim count As Integer = 0
            Dim no As Integer = 60 * i
            Dim dataArray(59, 4) As String
            While rowIndex <> rCount
                dataArray(count, 0) = (count + 1 + no).ToString() 'No.
                dataArray(count, 1) = Util.checkDBNullValue(dgvResident("Unt", rowIndex).Value) 'ユニット
                dataArray(count, 2) = Util.checkDBNullValue(dgvResident("Nam", rowIndex).Value) '入居者
                dataArray(count, 3) = Util.checkDBNullValue(dgvResident("Kana", rowIndex).Value) 'カナ
                dataArray(count, 4) = If(Util.checkDBNullValue(dgvResident("Dsp", rowIndex).Value) = 1, "○", "") '表示

                If count = 59 Then
                    dataArrayList.Add(dataArray)
                    rowIndex += 1
                    Continue For
                End If
                count += 1
                rowIndex += 1
            End While
            dataArrayList.Add(dataArray)
        Next

        'エクセルへデータ書き込み
        For i As Integer = 0 To dataArrayList.Count - 1
            oSheet.range("B" & ((64 * i) + 4), "F" & ((64 * i) + 63)).value = dataArrayList(i)
        Next

        objExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        objExcel.ScreenUpdating = True

        '変更保存確認ダイアログ非表示
        objExcel.DisplayAlerts = False

        '印刷
        If TopForm.rbtnPrintout.Checked = True Then
            oSheet.PrintOut()
        ElseIf TopForm.rbtnPreview.Checked = True Then
            objExcel.Visible = True
            oSheet.PrintPreview(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
End Class