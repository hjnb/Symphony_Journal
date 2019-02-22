<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 印鑑登録
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
        Me.classBox = New System.Windows.Forms.ComboBox()
        Me.namBox = New System.Windows.Forms.TextBox()
        Me.pwdBox = New System.Windows.Forms.TextBox()
        Me.sealNameBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnPrintSealM = New System.Windows.Forms.Button()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgvSeal = New System.Windows.Forms.DataGridView()
        Me.sealPicBox = New System.Windows.Forms.PictureBox()
        CType(Me.dgvSeal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sealPicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'classBox
        '
        Me.classBox.FormattingEnabled = True
        Me.classBox.Location = New System.Drawing.Point(94, 25)
        Me.classBox.Name = "classBox"
        Me.classBox.Size = New System.Drawing.Size(121, 20)
        Me.classBox.TabIndex = 0
        '
        'namBox
        '
        Me.namBox.Location = New System.Drawing.Point(94, 54)
        Me.namBox.Name = "namBox"
        Me.namBox.Size = New System.Drawing.Size(121, 19)
        Me.namBox.TabIndex = 1
        '
        'pwdBox
        '
        Me.pwdBox.Location = New System.Drawing.Point(94, 82)
        Me.pwdBox.Name = "pwdBox"
        Me.pwdBox.Size = New System.Drawing.Size(121, 19)
        Me.pwdBox.TabIndex = 2
        '
        'sealNameBox
        '
        Me.sealNameBox.Location = New System.Drawing.Point(94, 110)
        Me.sealNameBox.Name = "sealNameBox"
        Me.sealNameBox.Size = New System.Drawing.Size(121, 19)
        Me.sealNameBox.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(18, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "職制"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(18, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "職員名"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(18, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 12)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "ﾊﾟｽﾜｰﾄﾞ"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(18, 113)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 12)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "印影ﾌｧｲﾙ名"
        '
        'btnPrintSealM
        '
        Me.btnPrintSealM.Location = New System.Drawing.Point(367, 40)
        Me.btnPrintSealM.Name = "btnPrintSealM"
        Me.btnPrintSealM.Size = New System.Drawing.Size(58, 23)
        Me.btnPrintSealM.TabIndex = 8
        Me.btnPrintSealM.Text = "ﾌﾟﾘﾝﾄ"
        Me.btnPrintSealM.UseVisualStyleBackColor = True
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(279, 107)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(63, 27)
        Me.btnRegist.TabIndex = 4
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(341, 107)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(63, 27)
        Me.btnDelete.TabIndex = 10
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(403, 107)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(63, 27)
        Me.btnPrint.TabIndex = 11
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(63, 153)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(161, 12)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "ﾀﾞﾌﾞﾙｸﾘｯｸした項目名で並べます"
        '
        'dgvSeal
        '
        Me.dgvSeal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSeal.Location = New System.Drawing.Point(20, 173)
        Me.dgvSeal.Name = "dgvSeal"
        Me.dgvSeal.RowTemplate.Height = 21
        Me.dgvSeal.Size = New System.Drawing.Size(474, 422)
        Me.dgvSeal.TabIndex = 13
        '
        'sealPicBox
        '
        Me.sealPicBox.Location = New System.Drawing.Point(242, 25)
        Me.sealPicBox.Name = "sealPicBox"
        Me.sealPicBox.Size = New System.Drawing.Size(75, 72)
        Me.sealPicBox.TabIndex = 14
        Me.sealPicBox.TabStop = False
        '
        '印鑑登録
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(513, 609)
        Me.Controls.Add(Me.sealPicBox)
        Me.Controls.Add(Me.dgvSeal)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.btnPrintSealM)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.sealNameBox)
        Me.Controls.Add(Me.pwdBox)
        Me.Controls.Add(Me.namBox)
        Me.Controls.Add(Me.classBox)
        Me.Name = "印鑑登録"
        Me.Text = "印鑑登録"
        CType(Me.dgvSeal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sealPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents classBox As System.Windows.Forms.ComboBox
    Friend WithEvents namBox As System.Windows.Forms.TextBox
    Friend WithEvents pwdBox As System.Windows.Forms.TextBox
    Friend WithEvents sealNameBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnPrintSealM As System.Windows.Forms.Button
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgvSeal As System.Windows.Forms.DataGridView
    Friend WithEvents sealPicBox As System.Windows.Forms.PictureBox
End Class
