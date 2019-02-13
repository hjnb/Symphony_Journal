<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 印刷条件
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.dgvUnit = New System.Windows.Forms.DataGridView()
        Me.rbtnDiary = New System.Windows.Forms.RadioButton()
        Me.rbtnObservation = New System.Windows.Forms.RadioButton()
        Me.startYmdBox = New ymdBox.ymdBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.endYmdBox = New ymdBox.ymdBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.sign1Box = New System.Windows.Forms.TextBox()
        Me.sign2Box = New System.Windows.Forms.TextBox()
        Me.sign3Box = New System.Windows.Forms.TextBox()
        Me.sign5Box = New System.Windows.Forms.TextBox()
        Me.sign4Box = New System.Windows.Forms.TextBox()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.rbtnPrint = New System.Windows.Forms.RadioButton()
        Me.btnExcute = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.unitLabel = New System.Windows.Forms.Label()
        Me.errorLabel = New System.Windows.Forms.Label()
        CType(Me.dgvUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvUnit
        '
        Me.dgvUnit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUnit.Location = New System.Drawing.Point(44, 18)
        Me.dgvUnit.Name = "dgvUnit"
        Me.dgvUnit.RowTemplate.Height = 21
        Me.dgvUnit.Size = New System.Drawing.Size(214, 34)
        Me.dgvUnit.TabIndex = 0
        '
        'rbtnDiary
        '
        Me.rbtnDiary.AutoSize = True
        Me.rbtnDiary.Checked = True
        Me.rbtnDiary.Location = New System.Drawing.Point(3, 8)
        Me.rbtnDiary.Name = "rbtnDiary"
        Me.rbtnDiary.Size = New System.Drawing.Size(47, 16)
        Me.rbtnDiary.TabIndex = 1
        Me.rbtnDiary.TabStop = True
        Me.rbtnDiary.Text = "日誌"
        Me.rbtnDiary.UseVisualStyleBackColor = True
        '
        'rbtnObservation
        '
        Me.rbtnObservation.AutoSize = True
        Me.rbtnObservation.Location = New System.Drawing.Point(80, 8)
        Me.rbtnObservation.Name = "rbtnObservation"
        Me.rbtnObservation.Size = New System.Drawing.Size(59, 16)
        Me.rbtnObservation.TabIndex = 2
        Me.rbtnObservation.Text = "便観察"
        Me.rbtnObservation.UseVisualStyleBackColor = True
        '
        'startYmdBox
        '
        Me.startYmdBox.boxType = 2
        Me.startYmdBox.DateText = ""
        Me.startYmdBox.EraLabelText = "H31"
        Me.startYmdBox.EraText = ""
        Me.startYmdBox.Location = New System.Drawing.Point(17, 122)
        Me.startYmdBox.MonthLabelText = "02"
        Me.startYmdBox.MonthText = ""
        Me.startYmdBox.Name = "startYmdBox"
        Me.startYmdBox.Size = New System.Drawing.Size(110, 34)
        Me.startYmdBox.TabIndex = 30
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(137, 133)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "～"
        '
        'endYmdBox
        '
        Me.endYmdBox.boxType = 2
        Me.endYmdBox.DateText = ""
        Me.endYmdBox.EraLabelText = "H31"
        Me.endYmdBox.EraText = ""
        Me.endYmdBox.Location = New System.Drawing.Point(161, 122)
        Me.endYmdBox.MonthLabelText = "02"
        Me.endYmdBox.MonthText = ""
        Me.endYmdBox.Name = "endYmdBox"
        Me.endYmdBox.Size = New System.Drawing.Size(110, 34)
        Me.endYmdBox.TabIndex = 31
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 182)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "施設長"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 204)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 12)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "相談員"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 226)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 12)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "専門員"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 271)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(29, 12)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "合議"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 249)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 12)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "合議"
        '
        'sign1Box
        '
        Me.sign1Box.Location = New System.Drawing.Point(70, 179)
        Me.sign1Box.Name = "sign1Box"
        Me.sign1Box.Size = New System.Drawing.Size(100, 19)
        Me.sign1Box.TabIndex = 11
        '
        'sign2Box
        '
        Me.sign2Box.Location = New System.Drawing.Point(70, 201)
        Me.sign2Box.Name = "sign2Box"
        Me.sign2Box.Size = New System.Drawing.Size(100, 19)
        Me.sign2Box.TabIndex = 12
        '
        'sign3Box
        '
        Me.sign3Box.Location = New System.Drawing.Point(70, 223)
        Me.sign3Box.Name = "sign3Box"
        Me.sign3Box.Size = New System.Drawing.Size(100, 19)
        Me.sign3Box.TabIndex = 13
        '
        'sign5Box
        '
        Me.sign5Box.Location = New System.Drawing.Point(70, 267)
        Me.sign5Box.Name = "sign5Box"
        Me.sign5Box.Size = New System.Drawing.Size(100, 19)
        Me.sign5Box.TabIndex = 15
        '
        'sign4Box
        '
        Me.sign4Box.Location = New System.Drawing.Point(70, 245)
        Me.sign4Box.Name = "sign4Box"
        Me.sign4Box.Size = New System.Drawing.Size(100, 19)
        Me.sign4Box.TabIndex = 14
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Location = New System.Drawing.Point(3, 10)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(63, 16)
        Me.rbtnPreview.TabIndex = 16
        Me.rbtnPreview.TabStop = True
        Me.rbtnPreview.Text = "ﾌﾟﾚﾋﾞｭｰ"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrint
        '
        Me.rbtnPrint.AutoSize = True
        Me.rbtnPrint.Location = New System.Drawing.Point(88, 10)
        Me.rbtnPrint.Name = "rbtnPrint"
        Me.rbtnPrint.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrint.TabIndex = 17
        Me.rbtnPrint.TabStop = True
        Me.rbtnPrint.Text = "印刷"
        Me.rbtnPrint.UseVisualStyleBackColor = True
        '
        'btnExcute
        '
        Me.btnExcute.Location = New System.Drawing.Point(196, 343)
        Me.btnExcute.Name = "btnExcute"
        Me.btnExcute.Size = New System.Drawing.Size(75, 34)
        Me.btnExcute.TabIndex = 32
        Me.btnExcute.Text = "実行"
        Me.btnExcute.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rbtnDiary)
        Me.Panel1.Controls.Add(Me.rbtnObservation)
        Me.Panel1.Location = New System.Drawing.Point(17, 73)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 31)
        Me.Panel1.TabIndex = 33
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.rbtnPreview)
        Me.Panel2.Controls.Add(Me.rbtnPrint)
        Me.Panel2.Location = New System.Drawing.Point(17, 301)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(200, 36)
        Me.Panel2.TabIndex = 34
        '
        'unitLabel
        '
        Me.unitLabel.AutoSize = True
        Me.unitLabel.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.unitLabel.ForeColor = System.Drawing.Color.Blue
        Me.unitLabel.Location = New System.Drawing.Point(16, 25)
        Me.unitLabel.Name = "unitLabel"
        Me.unitLabel.Size = New System.Drawing.Size(0, 16)
        Me.unitLabel.TabIndex = 35
        '
        'errorLabel
        '
        Me.errorLabel.AutoSize = True
        Me.errorLabel.ForeColor = System.Drawing.Color.Red
        Me.errorLabel.Location = New System.Drawing.Point(15, 354)
        Me.errorLabel.Name = "errorLabel"
        Me.errorLabel.Size = New System.Drawing.Size(99, 12)
        Me.errorLabel.TabIndex = 36
        Me.errorLabel.Text = "印影ファイル未登録"
        Me.errorLabel.Visible = False
        '
        '印刷条件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(285, 392)
        Me.Controls.Add(Me.errorLabel)
        Me.Controls.Add(Me.unitLabel)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnExcute)
        Me.Controls.Add(Me.sign5Box)
        Me.Controls.Add(Me.sign4Box)
        Me.Controls.Add(Me.sign3Box)
        Me.Controls.Add(Me.sign2Box)
        Me.Controls.Add(Me.sign1Box)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.endYmdBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.startYmdBox)
        Me.Controls.Add(Me.dgvUnit)
        Me.Name = "印刷条件"
        Me.Text = "印刷条件"
        CType(Me.dgvUnit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvUnit As System.Windows.Forms.DataGridView
    Friend WithEvents rbtnDiary As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnObservation As System.Windows.Forms.RadioButton
    Friend WithEvents startYmdBox As ymdBox.ymdBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents endYmdBox As ymdBox.ymdBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents sign1Box As System.Windows.Forms.TextBox
    Friend WithEvents sign2Box As System.Windows.Forms.TextBox
    Friend WithEvents sign3Box As System.Windows.Forms.TextBox
    Friend WithEvents sign5Box As System.Windows.Forms.TextBox
    Friend WithEvents sign4Box As System.Windows.Forms.TextBox
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrint As System.Windows.Forms.RadioButton
    Friend WithEvents btnExcute As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents unitLabel As System.Windows.Forms.Label
    Friend WithEvents errorLabel As System.Windows.Forms.Label
End Class
