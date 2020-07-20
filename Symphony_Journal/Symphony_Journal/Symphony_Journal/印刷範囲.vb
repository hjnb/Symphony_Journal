Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class 印刷範囲

    '印刷From日付
    Private fromDate As String

    '印刷To日付
    Private toDate As String

    '入居者名
    Private residentName As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(residentName As String, fromDate As String, toDate As String)
        InitializeComponent()

        'フォーム設定
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        Me.residentName = residentName
        Me.fromDate = fromDate
        Me.toDate = toDate
    End Sub

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印刷範囲_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '初期値設定
        fromYmdBox.setADStr(fromDate)
        toYmdBox.setADStr(toDate)
    End Sub

    ''' <summary>
    ''' OKボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        Dim fromAD As String = fromYmdBox.getADStr() 'from日付
        Dim toAD As String = toYmdBox.getADStr() 'to日付

        'From,Toが正しいかチェック
        Dim fDt As New DateTime(CInt(fromAD.Substring(0, 4)), CInt(fromAD.Substring(5, 2)), CInt(fromAD.Substring(8, 2)))
        Dim tDt As New DateTime(CInt(toAD.Substring(0, 4)), CInt(toAD.Substring(5, 2)), CInt(toAD.Substring(8, 2)))
        Dim compareResult As Integer = fDt.CompareTo(tDt)
        If compareResult >= 1 Then
            MsgBox("FromがToより新しい日付になっています。", MsgBoxStyle.Exclamation)
            Return
        End If

        '印刷データ取得
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Journal)
        Dim sql As String = "select Ymd, Tanto, Text, Kisai from Skei where Div=" & TopForm.DIV & " And Nam='" & residentName & "' And ('" & fromAD & "' <= Ymd and Ymd <= '" & toAD & "') order by Ymd, Gyo"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        If rs.RecordCount <= 0 Then
            MsgBox("該当データがありません。", MsgBoxStyle.Exclamation)
            Return
        End If

        'エクセル準備
        Dim objExcel As Excel.Application = CreateObject("Excel.Application")
        Dim objWorkBooks As Excel.Workbooks = objExcel.Workbooks
        Dim objWorkBook As Excel.Workbook = objWorkBooks.Open(TopForm.excelFilePass)
        Dim oSheet As Excel.Worksheet = objWorkBook.Worksheets("支援経過改")

        objExcel.Calculation = Excel.XlCalculation.xlCalculationManual
        objExcel.ScreenUpdating = False

        '件数に応じてページ準備
        oSheet.range("E4").value = residentName
        Dim recordCount As Integer = rs.RecordCount
        For i As Integer = 0 To (recordCount \ 55) - 1
            Dim xlPasteRange As Excel.Range = oSheet.Range("A" & (1 + 62 * (i + 1))) 'ペースト先
            oSheet.Rows("1:62").copy(xlPasteRange)
            oSheet.HPageBreaks.Add(oSheet.Range("A" & (1 + 62 * (i + 1)))) '改ページ
        Next

        'データ書き込み
        Dim tanto As String = Util.checkDBNullValue(rs.Fields("Tanto").Value)
        Dim count As Integer = 0
        Dim pageCount As Integer = 0
        Dim ymdTemp As String = ""
        Dim dataArray(54, 10) As String
        While Not rs.EOF
            If count = 55 Then
                'データ貼り付け
                oSheet.range("B" & (7 + 62 * pageCount), "L" & (61 + 62 * pageCount)).value = dataArray
                oSheet.Range("K" & (4 + 62 * pageCount)).Value = tanto

                '配列内容クリア
                Array.Clear(dataArray, 0, dataArray.Length)

                '更新
                tanto = Util.checkDBNullValue(rs.Fields("Tanto").Value)
                pageCount += 1
                count = 0
            End If
            Dim ymd As String = Util.checkDBNullValue(rs.Fields("Ymd").Value)
            If ymd <> ymdTemp Then
                '日付
                dataArray(count, 0) = ymd
                ymdTemp = ymd
                '記載者
                dataArray(count, 10) = Util.checkDBNullValue(rs.Fields("Kisai").Value).Split("　")(0)
            End If
            '内容
            dataArray(count, 3) = Util.checkDBNullValue(rs.Fields("Text").Value)
            count += 1
            rs.MoveNext()
        End While
        'データ貼り付け
        oSheet.range("B" & (7 + 62 * pageCount), "L" & (61 + 62 * pageCount)).value = dataArray
        oSheet.Range("K" & (4 + 62 * pageCount)).Value = tanto

        objExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        objExcel.ScreenUpdating = True

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

        'フォーム閉じる
        Me.Close()
    End Sub

    ''' <summary>
    ''' キャンセルボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class