Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

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
        Dim targetUnit As String = unitLabel.Text '対象のユニット名
        Dim targetContent As Integer = If(rbtnDiary.Checked, 0, 1) '印刷対象内容(0:日誌, 1:便観察)
        Dim fromYmd As String = startYmdBox.getADStr() 'from日付
        Dim toYmd As String = endYmdBox.getADStr() 'to日付

        '印刷対象データ
        Dim sql As String
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        If targetContent = 0 Then
            '日誌データ取得用
            If targetUnit = "" Then
                '全てのユニットが対象
                sql = "select Ymd, Unit, Gyo, Nyu1, Nyu2, Nyu3, Gai1, Gai2, Gai3, Kyo1, Kyo2, Kyo3, Sign6, Sign7, Nam, NClr, Text, TClr from UNis where ('" & fromYmd & "' <= Ymd And Ymd <= '" & toYmd & "') And Div=" & TopForm.DIV & " order by Unit, Ymd, Gyo"
            Else
                sql = "select Ymd, Unit, Gyo, Nyu1, Nyu2, Nyu3, Gai1, Gai2, Gai3, Kyo1, Kyo2, Kyo3, Sign6, Sign7, Nam, NClr, Text, TClr from UNis where ('" & fromYmd & "' <= Ymd And Ymd <= '" & toYmd & "') And Unit='" & targetUnit & "' order by Ymd, Gyo"
            End If
        Else
            '便観察データ取得用
            sql = "select Ymd, Unit, Gyo, Text from Ben where ('" & fromYmd & "' <= Ymd And Ymd <= '" & toYmd & "') And Div=" & TopForm.DIV & " order by Ymd, Unit, Gyo"
        End If
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        '印刷対象データがない場合
        If rs.RecordCount <= 0 Then
            rs.Close()
            cnn.Close()
            MsgBox("該当がありません。", MsgBoxStyle.Exclamation)
            Return
        End If

        'エクセル準備
        Dim objExcel As Object = CreateObject("Excel.Application")
        Dim objWorkBooks As Object = objExcel.Workbooks
        Dim objWorkBook As Object = objWorkBooks.Open(TopForm.excelFilePass)
        Dim oSheet As Object

        '書き込み
        If targetContent = 0 Then
            oSheet = objWorkBook.Worksheets("ユニット日誌")
            writeUnitDiarySheet(oSheet, rs)
        Else
            oSheet = objWorkBook.Worksheets("便観察")
            writeBenSheet(oSheet, rs)
        End If
        rs.Close()
        cnn.Close()

        '変更保存確認ダイアログ非表示
        objExcel.DisplayAlerts = False

        '印刷
        If TopForm.rbtnPrintout.Checked = True Then
            oSheet.printOut()
        ElseIf TopForm.rbtnPreview.Checked = True Then
            objExcel.Visible = True
            oSheet.PrintPreview(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    ''' <summary>
    ''' ユニット日誌印刷書き込み
    ''' </summary>
    ''' <param name="oSheet"></param>
    ''' <remarks></remarks>
    Private Sub writeUnitDiarySheet(oSheet As Object, rs As ADODB.Recordset)
        '既存文字削除
        oSheet.range("B3").value = ""
        oSheet.range("J4").value = ""
        oSheet.range("K4").value = ""
        oSheet.range("L4").value = ""
        oSheet.range("M4").value = ""
        oSheet.range("N4").value = ""
        oSheet.range("O4").value = ""
        oSheet.range("P4").value = ""
        oSheet.range("B6").value = ""
        oSheet.range("D8").value = ""
        oSheet.range("F8").value = ""
        oSheet.range("H8").value = ""
        oSheet.range("D9").value = ""
        oSheet.range("F9").value = ""
        oSheet.range("H9").value = ""
        oSheet.range("D10").value = ""
        oSheet.range("F10").value = ""
        oSheet.range("H10").value = ""
        oSheet.range("B14").value = ""
        oSheet.range("E14").value = ""

        '
        Dim ymdTemp As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
        Dim unitTemp As String = Util.checkDBNullValue(rs.Fields("Unit").Value)
        Dim pageCount As Integer = 1
        Dim dayData(16, 3) As String
        Dim nightData(11, 3) As String
        Dim spData(3, 0) As String
        While Not rs.EOF
            Dim ymd As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
            Dim unit As String = Util.checkDBNullValue(rs.Fields("Unit").Value)
            If (ymd <> ymdTemp) OrElse (unit <> unitTemp) Then
                '現在のページにデータをセット
                oSheet.range("B" & (14 + 50 * (pageCount - 1)), "E" & (30 + 50 * (pageCount - 1))).value = dayData
                oSheet.range("B" & (33 + 50 * (pageCount - 1)), "E" & (44 + 50 * (pageCount - 1))).value = nightData
                oSheet.range("E" & (46 + 50 * (pageCount - 1)), "E" & (49 + 50 * (pageCount - 1))).value = spData

                '次のページ準備
                Dim xlPasteRange As Excel.Range = oSheet.Range("A" & (1 + 50 * pageCount)) 'ペースト先
                oSheet.rows("1:50").copy(xlPasteRange)

                '配列内容クリア
                Array.Clear(dayData, 0, dayData.Length)
                Array.Clear(nightData, 0, nightData.Length)
                Array.Clear(spData, 0, spData.Length)

                '更新
                ymdTemp = ymd
                unitTemp = unit
                pageCount += 1
            End If

            Dim gyo As Integer = rs.Fields("Gyo").Value
            If gyo = 0 Then
                oSheet.range("B" & (3 + 50 * (pageCount - 1))).value = unit & "のいえ" 'ユニット名


                '日付途中
                '
                '
                oSheet.range("B" & (6 + 50 * (pageCount - 1))).value = ymd.Substring(0, 4) & "年" & ymd.Substring(5, 2) & "月" & ymd.Substring(8, 2) & "日" '日付


                '日勤印影
                '夜勤印影
                oSheet.range("D" & (8 + 50 * (pageCount - 1))).value = If(rs.Fields("Nyu1").Value = 0, "", rs.Fields("Nyu1").Value) '入院者数　男
                oSheet.range("F" & (8 + 50 * (pageCount - 1))).value = If(rs.Fields("Nyu2").Value = 0, "", rs.Fields("Nyu2").Value) '入院者数　女
                oSheet.range("H" & (8 + 50 * (pageCount - 1))).value = If(rs.Fields("Nyu3").Value = 0, "", rs.Fields("Nyu3").Value) '入院者数　計
                oSheet.range("D" & (9 + 50 * (pageCount - 1))).value = If(rs.Fields("Gai1").Value = 0, "", rs.Fields("Gai1").Value) '外泊者数　男
                oSheet.range("F" & (9 + 50 * (pageCount - 1))).value = If(rs.Fields("Gai2").Value = 0, "", rs.Fields("Gai2").Value) '外泊者数　女
                oSheet.range("H" & (9 + 50 * (pageCount - 1))).value = If(rs.Fields("Gai3").Value = 0, "", rs.Fields("Gai3").Value) '外泊者数　計
                oSheet.range("D" & (10 + 50 * (pageCount - 1))).value = If(rs.Fields("Kyo1").Value = 0, "", rs.Fields("Kyo1").Value) '入居者数　男
                oSheet.range("F" & (10 + 50 * (pageCount - 1))).value = If(rs.Fields("Kyo2").Value = 0, "", rs.Fields("Kyo2").Value) '入居者数　女
                oSheet.range("H" & (10 + 50 * (pageCount - 1))).value = If(rs.Fields("Kyo3").Value = 0, "", rs.Fields("Kyo3").Value) '入居者数　計
            ElseIf 2 <= gyo AndAlso gyo <= 19 Then
                '日勤日誌データ作成
                dayData(gyo - 2, 0) = Util.checkDBNullValue(rs.Fields("Nam").Value)
                dayData(gyo - 2, 3) = Util.checkDBNullValue(rs.Fields("Text").Value)
            ElseIf 20 <= gyo AndAlso gyo <= 31 Then
                '夜勤日誌データ作成
                nightData(gyo - 20, 0) = Util.checkDBNullValue(rs.Fields("Nam").Value)
                nightData(gyo - 20, 3) = Util.checkDBNullValue(rs.Fields("Text").Value)
            ElseIf 32 <= gyo Then
                '特記事項データ作成
                spData(gyo - 32, 0) = Util.checkDBNullValue(rs.Fields("Text").Value)
            End If
            rs.MoveNext()
        End While
        '現在のページにデータをセット
        oSheet.range("B" & (14 + 50 * (pageCount - 1)), "E" & (30 + 50 * (pageCount - 1))).value = dayData
        oSheet.range("B" & (33 + 50 * (pageCount - 1)), "E" & (44 + 50 * (pageCount - 1))).value = nightData
        oSheet.range("E" & (46 + 50 * (pageCount - 1)), "E" & (49 + 50 * (pageCount - 1))).value = spData

    End Sub

    ''' <summary>
    ''' 便観察印刷書き込み
    ''' </summary>
    ''' <param name="oSheet"></param>
    ''' <remarks></remarks>
    Private Sub writeBenSheet(oSheet As Object, rs As ADODB.Recordset)
        '既存文字削除
        oSheet.range("D2").value = ""
        oSheet.range("C4").value = ""

        'ユニット名、配列格納位置対応dic
        Dim unitIndexDic As New Dictionary(Of String, Integer) From {{"星", 0}, {"森", 5}, {"空", 10}, {"月", 15}, {"花", 20}, {"海", 25}}

        'データ作成、エクセルに書き込み
        Dim ymdTemp As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
        Dim dataArray(31, 1) As String
        Dim pageCount As Integer = 1
        dataArray(0, 1) = Util.convADStrToWarekiStr(ymdTemp)
        While Not rs.EOF
            Dim ymd As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
            If ymd <> ymdTemp Then
                '現在のページにデータをセット
                oSheet.range("C" & (2 + 47 * (pageCount - 1)), "D" & (33 + 47 * (pageCount - 1))).value = dataArray

                '次のページ準備
                Dim xlPasteRange As Excel.Range = oSheet.Range("A" & (1 + 47 * pageCount)) 'ペースト先
                oSheet.rows("1:47").copy(xlPasteRange)

                '配列内容クリア
                Array.Clear(dataArray, 0, dataArray.Length)

                '更新
                ymdTemp = ymd
                pageCount += 1
                dataArray(0, 1) = Util.convADStrToWarekiStr(ymdTemp)
            End If
            Dim unitName As String = Util.checkDBNullValue(rs.Fields("Unit").Value)
            Dim gyo As Integer = rs.Fields("Gyo").Value - 1
            dataArray(unitIndexDic(unitName) + gyo + 2, 0) = Util.checkDBNullValue(rs.Fields("Text").Value)
            rs.MoveNext()
        End While
        '現在のページにデータをセット
        oSheet.range("C" & (2 + 47 * (pageCount - 1)), "D" & (33 + 47 * (pageCount - 1))).value = dataArray

    End Sub
End Class