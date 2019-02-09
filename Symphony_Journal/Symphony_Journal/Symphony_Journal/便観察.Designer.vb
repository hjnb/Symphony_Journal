<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 便観察
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
        Me.dgvBen = New System.Windows.Forms.DataGridView()
        Me.unitLabel = New System.Windows.Forms.Label()
        Me.dateLabel = New System.Windows.Forms.Label()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.benLabel = New System.Windows.Forms.Label()
        CType(Me.dgvBen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvBen
        '
        Me.dgvBen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBen.Location = New System.Drawing.Point(12, 12)
        Me.dgvBen.Name = "dgvBen"
        Me.dgvBen.RowTemplate.Height = 21
        Me.dgvBen.Size = New System.Drawing.Size(433, 109)
        Me.dgvBen.TabIndex = 0
        '
        'unitLabel
        '
        Me.unitLabel.AutoSize = True
        Me.unitLabel.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.unitLabel.Location = New System.Drawing.Point(12, 134)
        Me.unitLabel.Name = "unitLabel"
        Me.unitLabel.Size = New System.Drawing.Size(40, 14)
        Me.unitLabel.TabIndex = 1
        Me.unitLabel.Text = "のいえ"
        '
        'dateLabel
        '
        Me.dateLabel.AutoSize = True
        Me.dateLabel.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.dateLabel.Location = New System.Drawing.Point(68, 134)
        Me.dateLabel.Name = "dateLabel"
        Me.dateLabel.Size = New System.Drawing.Size(0, 14)
        Me.dateLabel.TabIndex = 2
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(377, 127)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(68, 28)
        Me.btnRegist.TabIndex = 3
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'benLabel
        '
        Me.benLabel.AutoSize = True
        Me.benLabel.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.benLabel.ForeColor = System.Drawing.Color.Red
        Me.benLabel.Location = New System.Drawing.Point(204, 134)
        Me.benLabel.Name = "benLabel"
        Me.benLabel.Size = New System.Drawing.Size(124, 14)
        Me.benLabel.TabIndex = 4
        Me.benLabel.Text = "便観察がありません。"
        Me.benLabel.Visible = False
        '
        '便観察
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(457, 162)
        Me.Controls.Add(Me.benLabel)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.dateLabel)
        Me.Controls.Add(Me.unitLabel)
        Me.Controls.Add(Me.dgvBen)
        Me.Name = "便観察"
        Me.Text = "便観察"
        CType(Me.dgvBen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvBen As System.Windows.Forms.DataGridView
    Friend WithEvents unitLabel As System.Windows.Forms.Label
    Friend WithEvents dateLabel As System.Windows.Forms.Label
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents benLabel As System.Windows.Forms.Label
End Class
