<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 入居者登録
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rbtnDisplayOn = New System.Windows.Forms.RadioButton()
        Me.rbtnDisplayOff = New System.Windows.Forms.RadioButton()
        Me.unitBox = New System.Windows.Forms.ComboBox()
        Me.namBox = New System.Windows.Forms.TextBox()
        Me.kanaBox = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgvResident = New System.Windows.Forms.DataGridView()
        Me.btnNewRegist = New System.Windows.Forms.Button()
        Me.btnChange = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        CType(Me.dgvResident, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ﾕﾆｯﾄ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "入居者名"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 12)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "ｶﾅ"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(24, 113)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "表示切替"
        '
        'rbtnDisplayOn
        '
        Me.rbtnDisplayOn.AutoSize = True
        Me.rbtnDisplayOn.Location = New System.Drawing.Point(94, 111)
        Me.rbtnDisplayOn.Name = "rbtnDisplayOn"
        Me.rbtnDisplayOn.Size = New System.Drawing.Size(42, 16)
        Me.rbtnDisplayOn.TabIndex = 10
        Me.rbtnDisplayOn.TabStop = True
        Me.rbtnDisplayOn.Text = "ＯＮ"
        Me.rbtnDisplayOn.UseVisualStyleBackColor = True
        '
        'rbtnDisplayOff
        '
        Me.rbtnDisplayOff.AutoSize = True
        Me.rbtnDisplayOff.Location = New System.Drawing.Point(148, 111)
        Me.rbtnDisplayOff.Name = "rbtnDisplayOff"
        Me.rbtnDisplayOff.Size = New System.Drawing.Size(49, 16)
        Me.rbtnDisplayOff.TabIndex = 11
        Me.rbtnDisplayOff.TabStop = True
        Me.rbtnDisplayOff.Text = "ＯＦＦ"
        Me.rbtnDisplayOff.UseVisualStyleBackColor = True
        '
        'unitBox
        '
        Me.unitBox.FormattingEnabled = True
        Me.unitBox.Location = New System.Drawing.Point(94, 27)
        Me.unitBox.Name = "unitBox"
        Me.unitBox.Size = New System.Drawing.Size(65, 20)
        Me.unitBox.TabIndex = 6
        '
        'namBox
        '
        Me.namBox.Location = New System.Drawing.Point(94, 54)
        Me.namBox.Name = "namBox"
        Me.namBox.Size = New System.Drawing.Size(112, 19)
        Me.namBox.TabIndex = 7
        '
        'kanaBox
        '
        Me.kanaBox.Location = New System.Drawing.Point(94, 81)
        Me.kanaBox.Name = "kanaBox"
        Me.kanaBox.Size = New System.Drawing.Size(112, 19)
        Me.kanaBox.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(48, 152)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(161, 12)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "ﾀﾞﾌﾞﾙｸﾘｯｸした項目名で並べます"
        '
        'dgvResident
        '
        Me.dgvResident.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvResident.Location = New System.Drawing.Point(21, 171)
        Me.dgvResident.Name = "dgvResident"
        Me.dgvResident.RowTemplate.Height = 21
        Me.dgvResident.Size = New System.Drawing.Size(334, 422)
        Me.dgvResident.TabIndex = 5
        '
        'btnNewRegist
        '
        Me.btnNewRegist.Location = New System.Drawing.Point(271, 26)
        Me.btnNewRegist.Name = "btnNewRegist"
        Me.btnNewRegist.Size = New System.Drawing.Size(57, 25)
        Me.btnNewRegist.TabIndex = 9
        Me.btnNewRegist.Text = "新規"
        Me.btnNewRegist.UseVisualStyleBackColor = True
        '
        'btnChange
        '
        Me.btnChange.Location = New System.Drawing.Point(271, 50)
        Me.btnChange.Name = "btnChange"
        Me.btnChange.Size = New System.Drawing.Size(57, 25)
        Me.btnChange.TabIndex = 12
        Me.btnChange.Text = "変更"
        Me.btnChange.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(271, 74)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(57, 25)
        Me.btnDelete.TabIndex = 13
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(271, 98)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(57, 25)
        Me.btnPrint.TabIndex = 14
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        '入居者登録
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(379, 610)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnChange)
        Me.Controls.Add(Me.btnNewRegist)
        Me.Controls.Add(Me.dgvResident)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.kanaBox)
        Me.Controls.Add(Me.namBox)
        Me.Controls.Add(Me.unitBox)
        Me.Controls.Add(Me.rbtnDisplayOff)
        Me.Controls.Add(Me.rbtnDisplayOn)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "入居者登録"
        Me.Text = "入居者登録"
        CType(Me.dgvResident, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents rbtnDisplayOn As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnDisplayOff As System.Windows.Forms.RadioButton
    Friend WithEvents unitBox As System.Windows.Forms.ComboBox
    Friend WithEvents namBox As System.Windows.Forms.TextBox
    Friend WithEvents kanaBox As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgvResident As System.Windows.Forms.DataGridView
    Friend WithEvents btnNewRegist As System.Windows.Forms.Button
    Friend WithEvents btnChange As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
