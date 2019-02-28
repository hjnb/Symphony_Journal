Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class 印鑑登録

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
        Me.FormBorderStyle = FormBorderStyle.FixedSingle        
    End Sub

    ''' <summary>
    ''' Loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印鑑登録_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'sealBoxフォルダ存在チェック
        If Not System.IO.Directory.Exists(TopForm.sealBoxDirPath) Then
            MsgBox(TopForm.sealBoxDirPath & "が存在しません。iniファイルにsealBoxの正しいパスを設定して下さい。")
            Me.Close()
            Exit Sub
        End If

        '初期設定
        initTextBox()
        initDgvSeal()

        'dgv表示
        displayDgvSeal()
    End Sub

    ''' <summary>
    ''' keyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印鑑登録_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        '職制
        classBox.Items.AddRange({"9合議", "8看護", "7記入者", "6介護支援専門員", "5生活相談員", "4介護部長", "3事務長", "2副施設長", "1施設長"})
        classBox.ImeMode = Windows.Forms.ImeMode.Off

        '職員名
        namBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'パスワード
        pwdBox.ImeMode = Windows.Forms.ImeMode.Disable

        '印影ファイル名
        sealNameBox.ImeMode = Windows.Forms.ImeMode.Disable
    End Sub

    ''' <summary>
    ''' データグリッドビュー初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvSeal()
        Util.EnableDoubleBuffering(dgvSeal)

        With dgvSeal
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
    Private Sub displayDgvSeal()
        dgvSeal.Columns.Clear()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Class, Nam, Pwd, File from SealM order by Class, Nam DESC"
        cnn.Open(TopForm.DB_Journal)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "SealM")
        dgvSeal.DataSource = ds.Tables("SealM")
        cnn.Close()

        settingDgvSeal()
    End Sub

    ''' <summary>
    ''' dgv表示設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub settingDgvSeal()
        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgvSeal.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        With dgvSeal
            With .Columns("Class")
                .HeaderText = "職制"
                .Width = 105
            End With

            With .Columns("Nam")
                .HeaderText = "職員名"
                .Width = 105
            End With

            With .Columns("Pwd")
                .HeaderText = "ﾊﾟｽﾜｰﾄﾞ"
                .Width = 105

            End With

            With .Columns("File")
                .HeaderText = "印影ﾌｧｲﾙ名"
                .Width = 105
            End With
        End With
    End Sub

    ''' <summary>
    ''' dgvセルマウスクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvSeal_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSeal.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim className As String = Util.checkDBNullValue(dgvSeal("Class", e.RowIndex).Value) '職制
            Dim nam As String = Util.checkDBNullValue(dgvSeal("Nam", e.RowIndex).Value) '職員名
            Dim pwd As String = Util.checkDBNullValue(dgvSeal("Pwd", e.RowIndex).Value) 'パスワード
            Dim fileNam As String = Util.checkDBNullValue(dgvSeal("File", e.RowIndex).Value) '印影ファイル名
            Dim sealFilePath As String = TopForm.sealBoxDirPath & "\" & fileNam & ".wmf" '印影ファイルパス

            '各ボックスへセット
            classBox.Text = className
            namBox.Text = nam
            pwdBox.Text = pwd
            sealNameBox.Text = fileNam
            If System.IO.File.Exists(sealFilePath) Then
                sealPicBox.ImageLocation = sealFilePath
            Else
                sealPicBox.ImageLocation = Nothing
            End If
        End If
    End Sub

    ''' <summary>
    ''' CellPaintingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvSeal_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvSeal.CellPainting
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
    Private Sub dgvSeal_ColumnHeaderMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSeal.ColumnHeaderMouseDoubleClick
        Dim targetColumn As DataGridViewColumn = dgvSeal.Columns(e.ColumnIndex) '選択列
        dgvSeal.Sort(targetColumn, System.ComponentModel.ListSortDirection.Ascending) '昇順でソート
    End Sub

    ''' <summary>
    ''' 入力内容クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub inputClear()
        classBox.Text = ""
        namBox.Text = ""
        pwdBox.Text = ""
        sealNameBox.Text = ""
        sealPicBox.ImageLocation = Nothing
    End Sub

    ''' <summary>
    ''' 職制が登録済みかチェック(1,2,3,4,5,6のみ)
    ''' </summary>
    ''' <param name="className">職制</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function existsClass(className As String) As Boolean
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT Class FROM SealM WHERE Class='" & className & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount = 0 Then
            rs.Close()
            cnn.Close()
            Return False
        Else
            rs.Close()
            cnn.Close()
            Return True
        End If
    End Function

    ''' <summary>
    ''' ﾊﾟｽﾜｰﾄﾞが登録済みかチェック
    ''' </summary>
    ''' <param name="inputPass">ﾊﾟｽﾜｰﾄﾞ</param>
    ''' <returns>登録済みの有無</returns>
    ''' <remarks></remarks>
    Private Function existsPassword(inputPass As String) As Boolean
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT Pwd FROM SealM WHERE Pwd='" & inputPass & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount = 0 Then
            rs.Close()
            cnn.Close()
            Return False
        Else
            rs.Close()
            cnn.Close()
            Return True
        End If
    End Function

    ''' <summary>
    ''' 印影ファイルがSealBoxに存在するかチェック
    ''' </summary>
    ''' <param name="sealFileName">印影ファイル名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function existsSealFile(sealFileName As String) As Boolean
        '印影ファイルパス
        Dim sealFilePath As String = TopForm.sealBoxDirPath & "\" & sealFileName & ".wmf"

        'ファイルが存在している場合はtrue
        If System.IO.File.Exists(sealFilePath) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim className As String = classBox.Text '職制
        Dim nam As String = namBox.Text '職員名
        Dim pwd As String = pwdBox.Text 'パスワード
        Dim fileName As String = sealNameBox.Text '印影ファイル名

        '職制チェック
        Dim classNum As String = If(className <> "", className.Substring(0, 1), "") '頭文字の数字
        If className = "" Then
            MsgBox("職制を選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        ElseIf (classNum = "1" OrElse classNum = "2" OrElse classNum = "3" OrElse classNum = "4" OrElse classNum = "5" OrElse classNum = "6") AndAlso existsClass(className) Then
            MsgBox("職制は既に登録されています。", MsgBoxStyle.Exclamation)
            Return
        End If

        '職員名チェック
        If nam = "" Then
            MsgBox("職員名を入力して下さい。", MsgBoxStyle.Exclamation)
            namBox.Focus()
            namBox.SelectionStart = 0
            Return
        End If

        'パスワードチェック
        If pwd = "" Then
            MsgBox("ﾊﾟｽﾜｰﾄﾞを入力して下さい。", MsgBoxStyle.Exclamation)
            pwdBox.Focus()
            pwdBox.SelectionStart = 0
            Return
        ElseIf existsPassword(pwd) Then
            MsgBox("このﾊﾟｽﾜｰﾄﾞは既に使われています", MsgBoxStyle.Exclamation)
            pwdBox.Focus()
            pwdBox.SelectionStart = 0
            Return
        End If

        '印影ファイル名チェック
        If fileName = "" Then
            MsgBox("印影ファイル名を入力して下さい。", MsgBoxStyle.Exclamation)
            sealNameBox.Focus()
            sealNameBox.SelectionStart = 0
            Return
        ElseIf Not existsSealFile(fileName) Then
            MsgBox("印影ファイルがSealBoxにありません。", MsgBoxStyle.Exclamation)
            sealNameBox.Focus()
            sealNameBox.SelectionStart = 0
            Return
        End If

        '登録処理
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset
        rs.Open("SealM", cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs.Fields("Class").Value = className
        rs.Fields("Nam").Value = nam
        rs.Fields("Pwd").Value = pwd
        rs.Fields("File").Value = fileName
        rs.Update()
        cn.Close()

        '再表示
        inputClear()
        displayDgvSeal()
    End Sub

    ''' <summary>
    ''' 削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        '削除対象者パスワード
        Dim targetPwd As String = pwdBox.Text

        '登録されているか確認
        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset()
        Dim sql As String = "select * from SealM where Pwd='" & targetPwd & "'"
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If rs.RecordCount <= 0 Then
            MsgBox("登録されていません。", MsgBoxStyle.Exclamation)
            cn.Close()
            Return
        End If

        '削除
        Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            rs.Delete()
            rs.Update()
            cn.Close()
        Else
            cn.Close()
            Return
        End If

        '再表示
        inputClear()
        displayDgvSeal()
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
        oSheet = objWorkBook.Worksheets("SealBox")

        objExcel.Calculation = Excel.XlCalculation.xlCalculationManual
        objExcel.ScreenUpdating = False

        '日付
        Dim ymd As String = Util.convADStrToWarekiStr(Today.ToString("yyyy/MM/dd"))
        oSheet.range("E2").value = ymd

        '既存文字削除
        oSheet.Range("B3").value = ""
        oSheet.Range("B4").value = ""

        Dim cn As New ADODB.Connection()
        cn.Open(TopForm.DB_Journal)
        Dim rs As New ADODB.Recordset()
        Dim sql As String = "select Class, Nam, Pwd, File from SealM order by Class"
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim rCount As Integer = rs.RecordCount 'レコード数
        If rCount > 180 Then
            '2枚目準備
            Dim xlPasteRange As Excel.Range = oSheet.Range("A34") 'ペースト先
            oSheet.rows("1:33").copy(xlPasteRange)
        End If

        '書き込み
        Dim columnStrArray() As String = {"B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"}
        Dim arrayLength As Integer = columnStrArray.Length
        Dim xlShapes As Excel.Shapes = DirectCast(oSheet.Shapes, Excel.Shapes)
        Dim rowCount As Integer = 0
        Dim excelRowIndex As Integer = 3
        While Not rs.EOF
            '書き込み列
            Dim columnStr As String = columnStrArray(rowCount Mod arrayLength)
            '印影ファイルが存在する場合のみ表示
            Dim filePath As String = TopForm.sealBoxDirPath & "\" & Util.checkDBNullValue(rs.Fields("File").Value) & ".wmf"
            If System.IO.File.Exists(filePath) Then
                Dim cell As Excel.Range = DirectCast(oSheet.Cells(excelRowIndex, columnStr), Excel.Range)
                xlShapes.AddPicture(filePath, False, True, cell.Left + 10, cell.Top + 6, 30, 30)
            End If
            '氏名
            oSheet.range(columnStr & (excelRowIndex + 1)).value = Util.checkDBNullValue(rs.Fields("Nam").Value)

            rowCount += 1
            If columnStr = "M" Then
                If rowCount = 180 Then
                    '2枚目へ書き込みのための処理
                    excelRowIndex = 36
                Else
                    excelRowIndex += 2
                End If
            End If
            rs.MoveNext()
        End While

        rs.Close()
        cn.Close()

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

    ''' <summary>
    ''' プリントボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrintSealM_Click(sender As System.Object, e As System.EventArgs) Handles btnPrintSealM.Click
        '行数
        Dim rowCount As Integer = dgvSeal.Rows.Count

        If rowCount = 0 Then
            Return
        End If

        'エクセル貼り付け用の配列作成
        Dim dataArray(rowCount - 1, 4) As String
        Dim count As Integer = 0
        For Each row As DataGridViewRow In dgvSeal.Rows
            dataArray(count, 0) = count + 1
            dataArray(count, 1) = Util.checkDBNullValue(row.Cells("Class").Value)
            dataArray(count, 2) = Util.checkDBNullValue(row.Cells("Nam").Value)
            dataArray(count, 3) = Util.checkDBNullValue(row.Cells("Pwd").Value)
            dataArray(count, 4) = Util.checkDBNullValue(row.Cells("File").Value)
            count += 1
        Next

        '現在日付、時刻
        Dim nowTime As String = Now.ToString("yyyy/MM/dd HH:MM:ss")

        'エクセル準備
        Dim objExcel As Object = CreateObject("Excel.Application")
        Dim objWorkBooks As Object = objExcel.Workbooks
        Dim objWorkBook As Object = objWorkBooks.Open(TopForm.excelFilePass)
        Dim oSheet As Object = objWorkBook.Worksheets("印影リスト")

        '日付時刻
        oSheet.range("B3").value = nowTime

        '人数分枠を用意
        Dim xlPasteRange As Excel.Range
        For i As Integer = 0 To rowCount - 2
            xlPasteRange = oSheet.Range("A" & (7 + i)) 'ペースト先
            oSheet.rows("6:6").copy(xlPasteRange)
        Next

        'データ貼り付け
        oSheet.range("B6", "F" & (5 + rowCount)).value = dataArray

        '2枚目が有る場合
        If rowCount >= 67 Then
            '空行insertし、そこにヘッダー部分をコピペ
            oSheet.rows("73:77").insert()
            xlPasteRange = oSheet.Range("A73") 'ペースト先
            oSheet.rows("1:5").copy(xlPasteRange)

            '改ページ
            oSheet.HpageBreaks.add(oSheet.Range("A73"))
        End If

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

    ''' <summary>
    ''' 印影ﾌｧｲﾙ名ボックスKeyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub sealNameBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles sealNameBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            'ファイルが存在する場合、画像を表示させる
            Dim fileName As String = sealNameBox.Text
            Dim sealFilePath As String = TopForm.sealBoxDirPath & "\" & fileName & ".wmf" '印影ファイルパス
            If System.IO.File.Exists(sealFilePath) Then
                sealPicBox.ImageLocation = sealFilePath
            Else
                sealPicBox.ImageLocation = Nothing
            End If
        End If
    End Sub
End Class