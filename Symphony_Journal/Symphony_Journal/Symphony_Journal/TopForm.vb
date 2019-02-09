Public Class TopForm
    'データベースのパス
    Public dbFilePath As String = My.Application.Info.DirectoryPath & "\Journal.mdb"
    Public DB_Journal As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'エクセルのパス
    Public excelFilePass As String = My.Application.Info.DirectoryPath & "\Journal.xls"

    '.iniファイルのパス
    Public iniFilePath As String = My.Application.Info.DirectoryPath & "\Journal1.ini"

    '画像パス
    Public imageFilePath As String = My.Application.Info.DirectoryPath & "\Div1.wmf"

    'SealBoxフォルダパス
    Public sealBoxDirPath As String = Util.getIniString("System", "SealBoxDir", iniFilePath)

    'フォーム
    Private unitDiary As ユニット日誌
    Private handover As 申し送り
    Private registSeal As 印鑑登録
    Private registResident As 入居者登録
    Private progressTable As 施設介護支援経過表
    Private ssLife As SS生活の様子

    'Div
    Public Const DIV As Integer = 1

    Private Sub TopForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'データベース、エクセル、構成ファイルの存在チェック
        If Not System.IO.File.Exists(dbFilePath) Then
            MsgBox("データベースファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(excelFilePass) Then
            MsgBox("エクセルファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(iniFilePath) Then
            MsgBox("構成ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(imageFilePath) Then
            MsgBox("画像ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        '画面サイズ等
        Me.WindowState = FormWindowState.Maximized
        Me.MinimizeBox = False
        Me.MaximizeBox = False

        '画像の配置処理
        topPicture.Image = System.Drawing.Image.FromFile(imageFilePath)

        '印刷ラジオボタンの初期設定
        initPrintState()
    End Sub

    Private Sub initPrintState()
        Dim state As String = Util.getIniString("System", "Printer", iniFilePath)
        If state = "Y" Then
            rbtnPrintout.Checked = True
        Else
            rbtnPreview.Checked = True
        End If
    End Sub

    Private Sub rbtnPreview_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPreview.CheckedChanged
        If rbtnPreview.Checked = True Then
            Util.putIniString("System", "Printer", "N", iniFilePath)
        End If
    End Sub

    Private Sub rbtnPrint_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPrintout.CheckedChanged
        If rbtnPrintout.Checked = True Then
            Util.putIniString("System", "Printer", "Y", iniFilePath)
        End If
    End Sub

    Private Sub topPicture_Click(sender As System.Object, e As System.EventArgs) Handles topPicture.Click
        Me.Close()
    End Sub

    Private Sub btnUnit_Click(sender As System.Object, e As System.EventArgs) Handles btnUnit.Click
        '印鑑パスワードフォーム表示
        Dim passForm As 印鑑パスワード = New 印鑑パスワード(DB_Journal)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        If IsNothing(unitDiary) OrElse unitDiary.IsDisposed Then
            unitDiary = New ユニット日誌(passForm.getSealFileName(), passForm.getClassName())
            unitDiary.Owner = Me
            unitDiary.Show()
        End If
    End Sub

    Private Sub btnHandover_Click(sender As System.Object, e As System.EventArgs) Handles btnHandover.Click
        If IsNothing(handover) OrElse handover.IsDisposed Then
            handover = New 申し送り()
            handover.Owner = Me
            handover.Show()
        End If
    End Sub

    Private Sub btnSeal_Click(sender As System.Object, e As System.EventArgs) Handles btnSeal.Click
        '管理者パスワードフォーム表示
        Dim passForm As Form = New passwordForm(iniFilePath, 3)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        If IsNothing(registSeal) OrElse registSeal.IsDisposed Then
            registSeal = New 印鑑登録()
            registSeal.Owner = Me
            registSeal.Show()
        End If
    End Sub

    Private Sub btnResident_Click(sender As System.Object, e As System.EventArgs) Handles btnResident.Click
        If IsNothing(registResident) OrElse registResident.IsDisposed Then
            registResident = New 入居者登録()
            registResident.Owner = Me
            registResident.Show()
        End If
    End Sub

    Private Sub btnProgress_Click(sender As System.Object, e As System.EventArgs) Handles btnProgress.Click
        If IsNothing(progressTable) OrElse progressTable.IsDisposed Then
            progressTable = New 施設介護支援経過表()
            progressTable.Owner = Me
            progressTable.Show()
        End If
    End Sub

    Private Sub btnSSLife_Click(sender As System.Object, e As System.EventArgs) Handles btnSSLife.Click
        If IsNothing(ssLife) OrElse ssLife.IsDisposed Then
            ssLife = New SS生活の様子()
            ssLife.Owner = Me
            ssLife.Show()
        End If
    End Sub
End Class
