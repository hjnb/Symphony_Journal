<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 印刷範囲
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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.fromYmdBox = New ADBox2.ADBox2()
        Me.toYmdBox = New ADBox2.ADBox2()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "利用期間 From, To の設定"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(76, 120)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(58, 28)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(165, 120)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(58, 28)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "ｷｬﾝｾﾙ"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'fromYmdBox
        '
        Me.fromYmdBox.dateText = ""
        Me.fromYmdBox.Location = New System.Drawing.Point(13, 67)
        Me.fromYmdBox.Mode = 1
        Me.fromYmdBox.monthText = ""
        Me.fromYmdBox.Name = "fromYmdBox"
        Me.fromYmdBox.Size = New System.Drawing.Size(125, 32)
        Me.fromYmdBox.TabIndex = 1
        Me.fromYmdBox.textReadOnly = False
        Me.fromYmdBox.yearText = ""
        '
        'toYmdBox
        '
        Me.toYmdBox.dateText = ""
        Me.toYmdBox.Location = New System.Drawing.Point(165, 67)
        Me.toYmdBox.Mode = 1
        Me.toYmdBox.monthText = ""
        Me.toYmdBox.Name = "toYmdBox"
        Me.toYmdBox.Size = New System.Drawing.Size(125, 32)
        Me.toYmdBox.TabIndex = 2
        Me.toYmdBox.textReadOnly = False
        Me.toYmdBox.yearText = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(141, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(22, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "～"
        '
        '印刷範囲
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(306, 172)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.toYmdBox)
        Me.Controls.Add(Me.fromYmdBox)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.Label1)
        Me.Name = "印刷範囲"
        Me.Text = "印刷範囲"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents fromYmdBox As ADBox2.ADBox2
    Friend WithEvents toYmdBox As ADBox2.ADBox2
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
