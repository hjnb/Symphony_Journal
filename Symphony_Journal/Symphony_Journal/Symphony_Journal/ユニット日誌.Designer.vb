<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ユニット日誌
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
        Me.components = New System.ComponentModel.Container()
        Me.unitListBox = New System.Windows.Forms.ListBox()
        Me.residentListBox = New System.Windows.Forms.ListBox()
        Me.YmdBox = New ymdBox.ymdBox()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnObserve = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.paintColorLabel = New System.Windows.Forms.Label()
        Me.Nyu1Box = New System.Windows.Forms.TextBox()
        Me.Gai1Box = New System.Windows.Forms.TextBox()
        Me.Kyo1Box = New System.Windows.Forms.TextBox()
        Me.Kyo2Box = New System.Windows.Forms.TextBox()
        Me.Gai2Box = New System.Windows.Forms.TextBox()
        Me.Nyu2Box = New System.Windows.Forms.TextBox()
        Me.Kyo3Box = New System.Windows.Forms.TextBox()
        Me.Gai3Box = New System.Windows.Forms.TextBox()
        Me.Nyu3Box = New System.Windows.Forms.TextBox()
        Me.rbtnDayWork = New System.Windows.Forms.RadioButton()
        Me.rbtnNightWork = New System.Windows.Forms.RadioButton()
        Me.dayWorkPic = New System.Windows.Forms.PictureBox()
        Me.nightWorkPic = New System.Windows.Forms.PictureBox()
        Me.colorContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.paintBlack = New System.Windows.Forms.ToolStripMenuItem()
        Me.paintBlue = New System.Windows.Forms.ToolStripMenuItem()
        Me.paintRed = New System.Windows.Forms.ToolStripMenuItem()
        Me.dgvUnitDiary = New Symphony_Journal.ExDataGridView(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.dayWorkPic, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nightWorkPic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.colorContextMenu.SuspendLayout()
        CType(Me.dgvUnitDiary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'unitListBox
        '
        Me.unitListBox.BackColor = System.Drawing.SystemColors.Control
        Me.unitListBox.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.unitListBox.FormattingEnabled = True
        Me.unitListBox.ItemHeight = 27
        Me.unitListBox.Location = New System.Drawing.Point(24, 23)
        Me.unitListBox.Name = "unitListBox"
        Me.unitListBox.Size = New System.Drawing.Size(38, 166)
        Me.unitListBox.TabIndex = 0
        '
        'residentListBox
        '
        Me.residentListBox.BackColor = System.Drawing.SystemColors.Control
        Me.residentListBox.FormattingEnabled = True
        Me.residentListBox.ItemHeight = 12
        Me.residentListBox.Location = New System.Drawing.Point(24, 241)
        Me.residentListBox.Name = "residentListBox"
        Me.residentListBox.Size = New System.Drawing.Size(91, 208)
        Me.residentListBox.TabIndex = 1
        '
        'YmdBox
        '
        Me.YmdBox.boxType = 8
        Me.YmdBox.DateText = ""
        Me.YmdBox.EraLabelText = "H31"
        Me.YmdBox.EraText = ""
        Me.YmdBox.Location = New System.Drawing.Point(134, 8)
        Me.YmdBox.MonthLabelText = "02"
        Me.YmdBox.MonthText = ""
        Me.YmdBox.Name = "YmdBox"
        Me.YmdBox.Size = New System.Drawing.Size(174, 46)
        Me.YmdBox.TabIndex = 18
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(401, 14)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(65, 30)
        Me.btnRegist.TabIndex = 3
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnObserve
        '
        Me.btnObserve.Location = New System.Drawing.Point(465, 14)
        Me.btnObserve.Name = "btnObserve"
        Me.btnObserve.Size = New System.Drawing.Size(65, 30)
        Me.btnObserve.TabIndex = 4
        Me.btnObserve.Text = "便観察"
        Me.btnObserve.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(529, 14)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(65, 30)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(134, 64)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(70, 18)
        Me.Panel1.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "入院者数"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(134, 82)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(70, 18)
        Me.Panel2.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 2)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "外泊者数"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Location = New System.Drawing.Point(134, 100)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(70, 18)
        Me.Panel3.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 12)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "入居者数"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(214, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 12)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "男"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(214, 85)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 12)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "男"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(214, 103)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 12)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "男"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(297, 103)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "女"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(297, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(17, 12)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "女"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(297, 67)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(17, 12)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "女"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(380, 103)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(17, 12)
        Me.Label10.TabIndex = 17
        Me.Label10.Text = "計"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(380, 85)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(17, 12)
        Me.Label11.TabIndex = 16
        Me.Label11.Text = "計"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(380, 67)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(17, 12)
        Me.Label12.TabIndex = 15
        Me.Label12.Text = "計"
        '
        'paintColorLabel
        '
        Me.paintColorLabel.AutoSize = True
        Me.paintColorLabel.ForeColor = System.Drawing.Color.Blue
        Me.paintColorLabel.Location = New System.Drawing.Point(474, 81)
        Me.paintColorLabel.Name = "paintColorLabel"
        Me.paintColorLabel.Size = New System.Drawing.Size(71, 12)
        Me.paintColorLabel.TabIndex = 18
        Me.paintColorLabel.Text = "着色：右ｸﾘｯｸ"
        '
        'Nyu1Box
        '
        Me.Nyu1Box.Location = New System.Drawing.Point(237, 64)
        Me.Nyu1Box.Name = "Nyu1Box"
        Me.Nyu1Box.Size = New System.Drawing.Size(52, 19)
        Me.Nyu1Box.TabIndex = 19
        Me.Nyu1Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Gai1Box
        '
        Me.Gai1Box.Location = New System.Drawing.Point(237, 82)
        Me.Gai1Box.Name = "Gai1Box"
        Me.Gai1Box.Size = New System.Drawing.Size(52, 19)
        Me.Gai1Box.TabIndex = 22
        Me.Gai1Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Kyo1Box
        '
        Me.Kyo1Box.Location = New System.Drawing.Point(237, 100)
        Me.Kyo1Box.Name = "Kyo1Box"
        Me.Kyo1Box.Size = New System.Drawing.Size(52, 19)
        Me.Kyo1Box.TabIndex = 25
        Me.Kyo1Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Kyo2Box
        '
        Me.Kyo2Box.Location = New System.Drawing.Point(320, 100)
        Me.Kyo2Box.Name = "Kyo2Box"
        Me.Kyo2Box.Size = New System.Drawing.Size(52, 19)
        Me.Kyo2Box.TabIndex = 26
        Me.Kyo2Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Gai2Box
        '
        Me.Gai2Box.Location = New System.Drawing.Point(320, 82)
        Me.Gai2Box.Name = "Gai2Box"
        Me.Gai2Box.Size = New System.Drawing.Size(52, 19)
        Me.Gai2Box.TabIndex = 23
        Me.Gai2Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Nyu2Box
        '
        Me.Nyu2Box.Location = New System.Drawing.Point(320, 64)
        Me.Nyu2Box.Name = "Nyu2Box"
        Me.Nyu2Box.Size = New System.Drawing.Size(52, 19)
        Me.Nyu2Box.TabIndex = 20
        Me.Nyu2Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Kyo3Box
        '
        Me.Kyo3Box.Location = New System.Drawing.Point(403, 100)
        Me.Kyo3Box.Name = "Kyo3Box"
        Me.Kyo3Box.Size = New System.Drawing.Size(52, 19)
        Me.Kyo3Box.TabIndex = 27
        Me.Kyo3Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Gai3Box
        '
        Me.Gai3Box.Location = New System.Drawing.Point(403, 82)
        Me.Gai3Box.Name = "Gai3Box"
        Me.Gai3Box.Size = New System.Drawing.Size(52, 19)
        Me.Gai3Box.TabIndex = 24
        Me.Gai3Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Nyu3Box
        '
        Me.Nyu3Box.Location = New System.Drawing.Point(403, 64)
        Me.Nyu3Box.Name = "Nyu3Box"
        Me.Nyu3Box.Size = New System.Drawing.Size(52, 19)
        Me.Nyu3Box.TabIndex = 21
        Me.Nyu3Box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'rbtnDayWork
        '
        Me.rbtnDayWork.AutoSize = True
        Me.rbtnDayWork.Location = New System.Drawing.Point(596, 55)
        Me.rbtnDayWork.Name = "rbtnDayWork"
        Me.rbtnDayWork.Size = New System.Drawing.Size(47, 16)
        Me.rbtnDayWork.TabIndex = 28
        Me.rbtnDayWork.TabStop = True
        Me.rbtnDayWork.Text = "日勤"
        Me.rbtnDayWork.UseVisualStyleBackColor = True
        '
        'rbtnNightWork
        '
        Me.rbtnNightWork.AutoSize = True
        Me.rbtnNightWork.Location = New System.Drawing.Point(650, 55)
        Me.rbtnNightWork.Name = "rbtnNightWork"
        Me.rbtnNightWork.Size = New System.Drawing.Size(47, 16)
        Me.rbtnNightWork.TabIndex = 29
        Me.rbtnNightWork.TabStop = True
        Me.rbtnNightWork.Text = "夜勤"
        Me.rbtnNightWork.UseVisualStyleBackColor = True
        '
        'dayWorkPic
        '
        Me.dayWorkPic.BackColor = System.Drawing.SystemColors.Window
        Me.dayWorkPic.Location = New System.Drawing.Point(599, 73)
        Me.dayWorkPic.Name = "dayWorkPic"
        Me.dayWorkPic.Size = New System.Drawing.Size(39, 42)
        Me.dayWorkPic.TabIndex = 30
        Me.dayWorkPic.TabStop = False
        '
        'nightWorkPic
        '
        Me.nightWorkPic.BackColor = System.Drawing.SystemColors.Window
        Me.nightWorkPic.Location = New System.Drawing.Point(653, 73)
        Me.nightWorkPic.Name = "nightWorkPic"
        Me.nightWorkPic.Size = New System.Drawing.Size(39, 42)
        Me.nightWorkPic.TabIndex = 31
        Me.nightWorkPic.TabStop = False
        '
        'colorContextMenu
        '
        Me.colorContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.paintBlack, Me.paintBlue, Me.paintRed})
        Me.colorContextMenu.Name = "colorContextMenu"
        Me.colorContextMenu.Size = New System.Drawing.Size(147, 70)
        '
        'paintBlack
        '
        Me.paintBlack.Name = "paintBlack"
        Me.paintBlack.Size = New System.Drawing.Size(146, 22)
        Me.paintBlack.Text = "黒"
        '
        'paintBlue
        '
        Me.paintBlue.Name = "paintBlue"
        Me.paintBlue.Size = New System.Drawing.Size(146, 22)
        Me.paintBlue.Text = "青（相談員）"
        '
        'paintRed
        '
        Me.paintRed.Name = "paintRed"
        Me.paintRed.Size = New System.Drawing.Size(146, 22)
        Me.paintRed.Text = "赤（看護師）"
        '
        'dgvUnitDiary
        '
        Me.dgvUnitDiary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUnitDiary.Location = New System.Drawing.Point(134, 135)
        Me.dgvUnitDiary.Name = "dgvUnitDiary"
        Me.dgvUnitDiary.RowTemplate.Height = 21
        Me.dgvUnitDiary.Size = New System.Drawing.Size(558, 530)
        Me.dgvUnitDiary.TabIndex = 32
        '
        'ユニット日誌
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(775, 720)
        Me.Controls.Add(Me.dgvUnitDiary)
        Me.Controls.Add(Me.nightWorkPic)
        Me.Controls.Add(Me.dayWorkPic)
        Me.Controls.Add(Me.rbtnNightWork)
        Me.Controls.Add(Me.rbtnDayWork)
        Me.Controls.Add(Me.Kyo3Box)
        Me.Controls.Add(Me.Gai3Box)
        Me.Controls.Add(Me.Nyu3Box)
        Me.Controls.Add(Me.Kyo2Box)
        Me.Controls.Add(Me.Gai2Box)
        Me.Controls.Add(Me.Nyu2Box)
        Me.Controls.Add(Me.Kyo1Box)
        Me.Controls.Add(Me.Gai1Box)
        Me.Controls.Add(Me.Nyu1Box)
        Me.Controls.Add(Me.paintColorLabel)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnObserve)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.YmdBox)
        Me.Controls.Add(Me.residentListBox)
        Me.Controls.Add(Me.unitListBox)
        Me.Name = "ユニット日誌"
        Me.Text = "ユニット日誌"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.dayWorkPic, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nightWorkPic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.colorContextMenu.ResumeLayout(False)
        CType(Me.dgvUnitDiary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents unitListBox As System.Windows.Forms.ListBox
    Friend WithEvents residentListBox As System.Windows.Forms.ListBox
    Friend WithEvents YmdBox As ymdBox.ymdBox
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnObserve As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents paintColorLabel As System.Windows.Forms.Label
    Friend WithEvents Nyu1Box As System.Windows.Forms.TextBox
    Friend WithEvents Gai1Box As System.Windows.Forms.TextBox
    Friend WithEvents Kyo1Box As System.Windows.Forms.TextBox
    Friend WithEvents Kyo2Box As System.Windows.Forms.TextBox
    Friend WithEvents Gai2Box As System.Windows.Forms.TextBox
    Friend WithEvents Nyu2Box As System.Windows.Forms.TextBox
    Friend WithEvents Kyo3Box As System.Windows.Forms.TextBox
    Friend WithEvents Gai3Box As System.Windows.Forms.TextBox
    Friend WithEvents Nyu3Box As System.Windows.Forms.TextBox
    Friend WithEvents rbtnDayWork As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNightWork As System.Windows.Forms.RadioButton
    Friend WithEvents dayWorkPic As System.Windows.Forms.PictureBox
    Friend WithEvents nightWorkPic As System.Windows.Forms.PictureBox
    Friend WithEvents dgvUnitDiary As Symphony_Journal.ExDataGridView
    Friend WithEvents colorContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents paintBlack As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents paintBlue As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents paintRed As System.Windows.Forms.ToolStripMenuItem
End Class
