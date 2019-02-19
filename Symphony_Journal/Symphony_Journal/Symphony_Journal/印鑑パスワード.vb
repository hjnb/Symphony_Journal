Imports System.Runtime.InteropServices

Public Class 印鑑パスワード

    'データベースファイルパス
    Private mdbFilePass As String

    '印鑑ファイル名
    Private sealFileName As String

    '職制
    Private className As String

    'テキストボックスのマウスダウンイベント制御用
    Private mdFlag As Boolean = False

    'パスワードチェックエラー文
    Private Const PASS_CHECK_ERROR As String = "印鑑パスワード未登録"

    '変更確認チェックエラー文
    Private Const PASS_CONFIRM_ERROR As String = "変更確認パスワードが合いません"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="mdbFilePass">データベースファイルパス</param>
    ''' <remarks></remarks>
    Public Sub New(mdbFilePass As String)

        InitializeComponent()

        'データベースファイルパス
        Me.mdbFilePass = mdbFilePass

        'フォーム設定
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

    End Sub

    ''' <summary>
    ''' keyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印鑑パスワード_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 印鑑パスワード_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'テキストボックス設定
        settingTextBox()
    End Sub

    ''' <summary>
    ''' テキストボックス設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub settingTextBox()
        passBox.ImeMode = Windows.Forms.ImeMode.Disable
        newPassBox.ImeMode = Windows.Forms.ImeMode.Disable
        confirmPassBox.ImeMode = Windows.Forms.ImeMode.Disable

        passBox.PasswordChar = "*"c
        newPassBox.PasswordChar = "*"c
        confirmPassBox.PasswordChar = "*"c
    End Sub

    ''' <summary>
    ''' パスワード照合
    ''' </summary>
    ''' <param name="inputPass">入力パスワード</param>
    ''' <returns>パスワード登録済みの有無</returns>
    ''' <remarks></remarks>
    Private Function checkPassword(inputPass As String) As Boolean
        Dim cnn As New ADODB.Connection
        cnn.Open(mdbFilePass)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT Class, File FROM SealM WHERE Pwd='" & inputPass & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            rs.Close()
            cnn.Close()
            Return False
        Else
            className = Util.checkDBNullValue(rs.Fields("Class").Value) '職制名を保持
            sealFileName = Util.checkDBNullValue(rs.Fields("File").Value) '印鑑ファイル名を保持
            rs.Close()
            cnn.Close()
            Return True
        End If
    End Function

    ''' <summary>
    ''' 登録パスワード変更
    ''' </summary>
    ''' <param name="oldPass">旧パスワード</param>
    ''' <param name="newPass">新パスワード</param>
    ''' <returns>変更の有無</returns>
    ''' <remarks></remarks>
    Private Function changePassword(oldPass As String, newPass As String) As Boolean
        Dim cnn As New ADODB.Connection
        cnn.Open(mdbFilePass)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT Class, Pwd, File FROM SealM WHERE Pwd='" & oldPass & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            rs.Close()
            cnn.Close()
            Return False
        Else
            '新パスワードが登録可能かチェック
            Dim rs2 As New ADODB.Recordset
            Dim sql2 = "SELECT Pwd, File FROM SealM WHERE Pwd='" & newPass & "'"
            rs2.Open(sql2, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            If rs2.RecordCount >= 1 Then
                rs2.Close()
                cnn.Close()
                MsgBox("このパスワードは既に使用されているため登録できません。")
                Return False
            Else
                rs.Fields("Pwd").Value = newPass
                className = Util.checkDBNullValue(rs.Fields("Class").Value) '職制名を保持
                sealFileName = Util.checkDBNullValue(rs.Fields("File").Value) '印鑑ファイル名を保持
                rs.Update()
                rs.Close()
                rs2.Close()
                cnn.Close()
                Return True
            End If
        End If
    End Function

    ''' <summary>
    ''' OKボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        Dim inputPass As String = passBox.Text
        Dim inputNewPass As String = newPassBox.Text
        Dim inputConfirmPass As String = confirmPassBox.Text

        If inputNewPass = "" AndAlso inputConfirmPass = "" Then
            '入力パスワード確認のみの場合
            If checkPassword(inputPass) = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                errorLabel.Text = PASS_CHECK_ERROR
                errorLabel.Visible = True
                passBox.Focus()
            End If
        Else
            '入力パスワード確認、パスワード変更ありの場合
            If checkPassword(inputPass) = True Then
                If inputNewPass = inputConfirmPass Then
                    'パスワード変更処理
                    If changePassword(inputPass, inputNewPass) = False Then
                        MsgBox("パスワードの変更に失敗しました。")
                        Return
                    End If

                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                Else
                    errorLabel.Text = PASS_CONFIRM_ERROR
                    errorLabel.Visible = True
                    confirmPassBox.Focus()
                End If
            Else
                errorLabel.Text = PASS_CHECK_ERROR
                errorLabel.Visible = True
                passBox.Focus()
            End If
        End If

    End Sub

    ''' <summary>
    ''' キャンセルボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Function getSealFileName() As String
        Return sealFileName
    End Function

    Public Function getClassName() As String
        Return className
    End Function

    Private Sub textBox_Enter(sender As Object, e As System.EventArgs) Handles passBox.Enter, newPassBox.Enter, confirmPassBox.Enter
        Dim tb As TextBox = CType(sender, TextBox)
        tb.SelectAll()
        mdFlag = True
    End Sub

    Private Sub textBox_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles passBox.MouseDown, newPassBox.MouseDown, confirmPassBox.MouseDown
        If mdFlag = True Then
            Dim tb As TextBox = CType(sender, TextBox)
            tb.SelectAll()
            mdFlag = False
        End If
    End Sub
End Class