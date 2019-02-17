<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 申し送り
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
        Me.dgvInput = New System.Windows.Forms.DataGridView()
        Me.YmdBox = New ymdBox.ymdBox()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.writerLabel = New System.Windows.Forms.Label()
        Me.writerListBox = New System.Windows.Forms.ListBox()
        Me.dgvRead = New System.Windows.Forms.DataGridView()
        Me.HmBox = New hmBox.hmBox()
        CType(Me.dgvInput, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvRead, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvInput
        '
        Me.dgvInput.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvInput.Location = New System.Drawing.Point(218, 44)
        Me.dgvInput.Name = "dgvInput"
        Me.dgvInput.RowTemplate.Height = 21
        Me.dgvInput.Size = New System.Drawing.Size(496, 209)
        Me.dgvInput.TabIndex = 0
        '
        'YmdBox
        '
        Me.YmdBox.boxType = 8
        Me.YmdBox.DateText = ""
        Me.YmdBox.EraLabelText = "H31"
        Me.YmdBox.EraText = ""
        Me.YmdBox.Location = New System.Drawing.Point(28, 87)
        Me.YmdBox.MonthLabelText = "02"
        Me.YmdBox.MonthText = ""
        Me.YmdBox.Name = "YmdBox"
        Me.YmdBox.Size = New System.Drawing.Size(174, 46)
        Me.YmdBox.TabIndex = 1
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(730, 51)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(69, 33)
        Me.btnRegist.TabIndex = 2
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(730, 90)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(69, 33)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "行削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.writerLabel)
        Me.GroupBox1.Location = New System.Drawing.Point(832, 215)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(111, 52)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "記載者"
        '
        'writerLabel
        '
        Me.writerLabel.AutoSize = True
        Me.writerLabel.ForeColor = System.Drawing.Color.Blue
        Me.writerLabel.Location = New System.Drawing.Point(15, 23)
        Me.writerLabel.Name = "writerLabel"
        Me.writerLabel.Size = New System.Drawing.Size(0, 12)
        Me.writerLabel.TabIndex = 0
        '
        'writerListBox
        '
        Me.writerListBox.BackColor = System.Drawing.SystemColors.Control
        Me.writerListBox.FormattingEnabled = True
        Me.writerListBox.ItemHeight = 12
        Me.writerListBox.Location = New System.Drawing.Point(835, 276)
        Me.writerListBox.Name = "writerListBox"
        Me.writerListBox.Size = New System.Drawing.Size(105, 388)
        Me.writerListBox.TabIndex = 5
        '
        'dgvRead
        '
        Me.dgvRead.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRead.Location = New System.Drawing.Point(98, 276)
        Me.dgvRead.Name = "dgvRead"
        Me.dgvRead.RowTemplate.Height = 21
        Me.dgvRead.Size = New System.Drawing.Size(710, 410)
        Me.dgvRead.TabIndex = 6
        '
        'HmBox
        '
        Me.HmBox.hourText = "00"
        Me.HmBox.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.HmBox.Location = New System.Drawing.Point(95, 153)
        Me.HmBox.minuteText = "00"
        Me.HmBox.Name = "HmBox"
        Me.HmBox.Size = New System.Drawing.Size(111, 46)
        Me.HmBox.TabIndex = 7
        '
        '申し送り
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(973, 717)
        Me.Controls.Add(Me.HmBox)
        Me.Controls.Add(Me.dgvRead)
        Me.Controls.Add(Me.writerListBox)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.YmdBox)
        Me.Controls.Add(Me.dgvInput)
        Me.Name = "申し送り"
        Me.Text = "Report - 申し送り"
        CType(Me.dgvInput, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvRead, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvInput As System.Windows.Forms.DataGridView
    Friend WithEvents YmdBox As ymdBox.ymdBox
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents writerLabel As System.Windows.Forms.Label
    Friend WithEvents writerListBox As System.Windows.Forms.ListBox
    Friend WithEvents dgvRead As System.Windows.Forms.DataGridView
    Friend WithEvents HmBox As hmBox.hmBox
End Class
