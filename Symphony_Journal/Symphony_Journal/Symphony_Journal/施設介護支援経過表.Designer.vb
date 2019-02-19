<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 施設介護支援経過表
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.namLabel = New System.Windows.Forms.Label()
        Me.YmdBox = New ymdBox.ymdBox()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tantoBox = New System.Windows.Forms.ComboBox()
        Me.dgvSkei = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.residentListBox = New System.Windows.Forms.ListBox()
        Me.unitListBox = New System.Windows.Forms.ListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.historyListBox = New System.Windows.Forms.ListBox()
        Me.btnRowInsert = New System.Windows.Forms.Button()
        Me.btnRowDelete = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.writerLabel = New System.Windows.Forms.Label()
        Me.writerListBox = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.dgvSkei, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'namLabel
        '
        Me.namLabel.AutoSize = True
        Me.namLabel.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.namLabel.ForeColor = System.Drawing.Color.Blue
        Me.namLabel.Location = New System.Drawing.Point(39, 26)
        Me.namLabel.Name = "namLabel"
        Me.namLabel.Size = New System.Drawing.Size(0, 19)
        Me.namLabel.TabIndex = 0
        '
        'YmdBox
        '
        Me.YmdBox.boxType = 8
        Me.YmdBox.DateText = ""
        Me.YmdBox.EraLabelText = "H31"
        Me.YmdBox.EraText = ""
        Me.YmdBox.Location = New System.Drawing.Point(171, 12)
        Me.YmdBox.MonthLabelText = "02"
        Me.YmdBox.MonthText = ""
        Me.YmdBox.Name = "YmdBox"
        Me.YmdBox.Size = New System.Drawing.Size(174, 46)
        Me.YmdBox.TabIndex = 1
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(380, 15)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(65, 36)
        Me.btnRegist.TabIndex = 2
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(444, 15)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(65, 36)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(508, 15)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(65, 36)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(615, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 11)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "施設サービス計画作成者"
        '
        'tantoBox
        '
        Me.tantoBox.FormattingEnabled = True
        Me.tantoBox.Location = New System.Drawing.Point(615, 34)
        Me.tantoBox.Name = "tantoBox"
        Me.tantoBox.Size = New System.Drawing.Size(114, 20)
        Me.tantoBox.TabIndex = 6
        '
        'dgvSkei
        '
        Me.dgvSkei.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.0!)
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSkei.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvSkei.Location = New System.Drawing.Point(171, 65)
        Me.dgvSkei.Name = "dgvSkei"
        Me.dgvSkei.RowTemplate.Height = 21
        Me.dgvSkei.Size = New System.Drawing.Size(612, 633)
        Me.dgvSkei.TabIndex = 7
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.residentListBox)
        Me.GroupBox1.Controls.Add(Me.unitListBox)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 61)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(140, 300)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "入居者選択"
        '
        'residentListBox
        '
        Me.residentListBox.BackColor = System.Drawing.SystemColors.Control
        Me.residentListBox.FormattingEnabled = True
        Me.residentListBox.ItemHeight = 12
        Me.residentListBox.Location = New System.Drawing.Point(8, 91)
        Me.residentListBox.Name = "residentListBox"
        Me.residentListBox.Size = New System.Drawing.Size(123, 196)
        Me.residentListBox.TabIndex = 1
        '
        'unitListBox
        '
        Me.unitListBox.FormattingEnabled = True
        Me.unitListBox.ItemHeight = 12
        Me.unitListBox.Location = New System.Drawing.Point(17, 26)
        Me.unitListBox.Name = "unitListBox"
        Me.unitListBox.Size = New System.Drawing.Size(50, 52)
        Me.unitListBox.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.historyListBox)
        Me.GroupBox2.Location = New System.Drawing.Point(14, 370)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(140, 318)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "経過履歴"
        '
        'historyListBox
        '
        Me.historyListBox.BackColor = System.Drawing.SystemColors.Control
        Me.historyListBox.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.historyListBox.FormattingEnabled = True
        Me.historyListBox.ItemHeight = 15
        Me.historyListBox.Location = New System.Drawing.Point(8, 18)
        Me.historyListBox.Name = "historyListBox"
        Me.historyListBox.Size = New System.Drawing.Size(125, 289)
        Me.historyListBox.TabIndex = 0
        '
        'btnRowInsert
        '
        Me.btnRowInsert.Location = New System.Drawing.Point(816, 87)
        Me.btnRowInsert.Name = "btnRowInsert"
        Me.btnRowInsert.Size = New System.Drawing.Size(64, 29)
        Me.btnRowInsert.TabIndex = 10
        Me.btnRowInsert.Text = "行挿入"
        Me.btnRowInsert.UseVisualStyleBackColor = True
        '
        'btnRowDelete
        '
        Me.btnRowDelete.Location = New System.Drawing.Point(816, 115)
        Me.btnRowDelete.Name = "btnRowDelete"
        Me.btnRowDelete.Size = New System.Drawing.Size(64, 29)
        Me.btnRowDelete.TabIndex = 11
        Me.btnRowDelete.Text = "行削除"
        Me.btnRowDelete.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(816, 143)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(64, 29)
        Me.btnClear.TabIndex = 12
        Me.btnClear.Text = "クリア"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.writerLabel)
        Me.GroupBox3.Location = New System.Drawing.Point(818, 279)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(108, 44)
        Me.GroupBox3.TabIndex = 13
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "記載者"
        '
        'writerLabel
        '
        Me.writerLabel.AutoSize = True
        Me.writerLabel.ForeColor = System.Drawing.Color.Blue
        Me.writerLabel.Location = New System.Drawing.Point(17, 21)
        Me.writerLabel.Name = "writerLabel"
        Me.writerLabel.Size = New System.Drawing.Size(0, 12)
        Me.writerLabel.TabIndex = 0
        '
        'writerListBox
        '
        Me.writerListBox.BackColor = System.Drawing.SystemColors.Control
        Me.writerListBox.FormattingEnabled = True
        Me.writerListBox.ItemHeight = 12
        Me.writerListBox.Location = New System.Drawing.Point(816, 336)
        Me.writerListBox.Name = "writerListBox"
        Me.writerListBox.Size = New System.Drawing.Size(117, 352)
        Me.writerListBox.TabIndex = 14
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.Label2.Location = New System.Drawing.Point(787, 202)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(15, 10)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "10"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.Label3.Location = New System.Drawing.Point(786, 331)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(15, 10)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "20"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.Label4.Location = New System.Drawing.Point(786, 461)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(15, 10)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "30"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.Label5.Location = New System.Drawing.Point(786, 590)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(15, 10)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "40"
        '
        '施設介護支援経過表
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(963, 722)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.writerListBox)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnRowDelete)
        Me.Controls.Add(Me.btnRowInsert)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgvSkei)
        Me.Controls.Add(Me.tantoBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.YmdBox)
        Me.Controls.Add(Me.namLabel)
        Me.Name = "施設介護支援経過表"
        Me.Text = "施設介護支援経過表"
        CType(Me.dgvSkei, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents namLabel As System.Windows.Forms.Label
    Friend WithEvents YmdBox As ymdBox.ymdBox
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tantoBox As System.Windows.Forms.ComboBox
    Friend WithEvents dgvSkei As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents residentListBox As System.Windows.Forms.ListBox
    Friend WithEvents unitListBox As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents historyListBox As System.Windows.Forms.ListBox
    Friend WithEvents btnRowInsert As System.Windows.Forms.Button
    Friend WithEvents btnRowDelete As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents writerLabel As System.Windows.Forms.Label
    Friend WithEvents writerListBox As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
