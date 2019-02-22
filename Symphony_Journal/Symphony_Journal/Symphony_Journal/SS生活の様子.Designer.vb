<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SS生活の様子
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
        Me.residentListBox = New System.Windows.Forms.ListBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.benYmdBox = New ymdBox.ymdBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.bathYmdBox = New ymdBox.ymdBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.endYmdBox = New ymdBox.ymdBox()
        Me.firstYmdBox = New ymdBox.ymdBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.dateYmdBox = New ymdBox.ymdBox()
        Me.tantoBox = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnTextClear = New System.Windows.Forms.Button()
        Me.historyListBox = New System.Windows.Forms.ListBox()
        Me.btnRowInsert = New System.Windows.Forms.Button()
        Me.btnRowDelete = New System.Windows.Forms.Button()
        Me.namLabel = New System.Windows.Forms.Label()
        Me.dgvShtM = New Symphony_Journal.SSDataGridView(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel6.SuspendLayout()
        CType(Me.dgvShtM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'residentListBox
        '
        Me.residentListBox.BackColor = System.Drawing.SystemColors.Control
        Me.residentListBox.FormattingEnabled = True
        Me.residentListBox.ItemHeight = 12
        Me.residentListBox.Location = New System.Drawing.Point(26, 66)
        Me.residentListBox.Name = "residentListBox"
        Me.residentListBox.Size = New System.Drawing.Size(120, 460)
        Me.residentListBox.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(168, 46)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(312, 72)
        Me.Panel1.TabIndex = 18
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.SystemColors.Window
        Me.Panel5.Controls.Add(Me.benYmdBox)
        Me.Panel5.Location = New System.Drawing.Point(90, 47)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(86, 23)
        Me.Panel5.TabIndex = 23
        '
        'benYmdBox
        '
        Me.benYmdBox.boxType = 9
        Me.benYmdBox.DateText = ""
        Me.benYmdBox.EraLabelText = "H31"
        Me.benYmdBox.EraText = ""
        Me.benYmdBox.Location = New System.Drawing.Point(-1, 2)
        Me.benYmdBox.MonthLabelText = "02"
        Me.benYmdBox.MonthText = ""
        Me.benYmdBox.Name = "benYmdBox"
        Me.benYmdBox.Size = New System.Drawing.Size(86, 20)
        Me.benYmdBox.TabIndex = 17
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.SystemColors.Window
        Me.Panel4.Controls.Add(Me.bathYmdBox)
        Me.Panel4.Location = New System.Drawing.Point(90, 23)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(86, 24)
        Me.Panel4.TabIndex = 22
        '
        'bathYmdBox
        '
        Me.bathYmdBox.boxType = 9
        Me.bathYmdBox.DateText = ""
        Me.bathYmdBox.EraLabelText = "H31"
        Me.bathYmdBox.EraText = ""
        Me.bathYmdBox.Location = New System.Drawing.Point(-1, 2)
        Me.bathYmdBox.MonthLabelText = "02"
        Me.bathYmdBox.MonthText = ""
        Me.bathYmdBox.Name = "bathYmdBox"
        Me.bathYmdBox.Size = New System.Drawing.Size(86, 20)
        Me.bathYmdBox.TabIndex = 16
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.Window
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.endYmdBox)
        Me.Panel3.Controls.Add(Me.firstYmdBox)
        Me.Panel3.Location = New System.Drawing.Point(90, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(221, 24)
        Me.Panel3.TabIndex = 21
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(94, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "～"
        '
        'endYmdBox
        '
        Me.endYmdBox.boxType = 9
        Me.endYmdBox.DateText = ""
        Me.endYmdBox.EraLabelText = "H31"
        Me.endYmdBox.EraText = ""
        Me.endYmdBox.Location = New System.Drawing.Point(118, 1)
        Me.endYmdBox.MonthLabelText = "02"
        Me.endYmdBox.MonthText = ""
        Me.endYmdBox.Name = "endYmdBox"
        Me.endYmdBox.Size = New System.Drawing.Size(86, 20)
        Me.endYmdBox.TabIndex = 15
        '
        'firstYmdBox
        '
        Me.firstYmdBox.boxType = 9
        Me.firstYmdBox.DateText = ""
        Me.firstYmdBox.EraLabelText = "H31"
        Me.firstYmdBox.EraText = ""
        Me.firstYmdBox.Location = New System.Drawing.Point(-1, 1)
        Me.firstYmdBox.MonthLabelText = "02"
        Me.firstYmdBox.MonthText = ""
        Me.firstYmdBox.Name = "firstYmdBox"
        Me.firstYmdBox.Size = New System.Drawing.Size(86, 20)
        Me.firstYmdBox.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "最終排便"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 12)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "最終入浴"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "利用期間"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Panel6)
        Me.Panel2.Controls.Add(Me.tantoBox)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Location = New System.Drawing.Point(485, 46)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(207, 45)
        Me.Panel2.TabIndex = 19
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.SystemColors.Window
        Me.Panel6.Controls.Add(Me.dateYmdBox)
        Me.Panel6.Location = New System.Drawing.Point(77, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(129, 24)
        Me.Panel6.TabIndex = 22
        '
        'dateYmdBox
        '
        Me.dateYmdBox.boxType = 9
        Me.dateYmdBox.DateText = ""
        Me.dateYmdBox.EraLabelText = "H31"
        Me.dateYmdBox.EraText = ""
        Me.dateYmdBox.Location = New System.Drawing.Point(-1, 1)
        Me.dateYmdBox.MonthLabelText = "02"
        Me.dateYmdBox.MonthText = ""
        Me.dateYmdBox.Name = "dateYmdBox"
        Me.dateYmdBox.Size = New System.Drawing.Size(86, 20)
        Me.dateYmdBox.TabIndex = 16
        '
        'tantoBox
        '
        Me.tantoBox.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.tantoBox.Location = New System.Drawing.Point(77, 24)
        Me.tantoBox.Name = "tantoBox"
        Me.tantoBox.Size = New System.Drawing.Size(129, 19)
        Me.tantoBox.TabIndex = 23
        Me.tantoBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 12)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "記載者"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 5)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 12)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "記載日"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(379, 5)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(66, 37)
        Me.btnNew.TabIndex = 3
        Me.btnNew.Text = "新規"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(444, 5)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(66, 37)
        Me.btnRegist.TabIndex = 4
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(509, 5)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(66, 37)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(574, 5)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(66, 37)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(159, 125)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 19)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "生活の様子"
        '
        'btnTextClear
        '
        Me.btnTextClear.Location = New System.Drawing.Point(617, 103)
        Me.btnTextClear.Name = "btnTextClear"
        Me.btnTextClear.Size = New System.Drawing.Size(75, 23)
        Me.btnTextClear.TabIndex = 9
        Me.btnTextClear.Text = "ﾃｷｽﾄｸﾘｱ"
        Me.btnTextClear.UseVisualStyleBackColor = True
        '
        'historyListBox
        '
        Me.historyListBox.BackColor = System.Drawing.SystemColors.Control
        Me.historyListBox.FormattingEnabled = True
        Me.historyListBox.ItemHeight = 12
        Me.historyListBox.Location = New System.Drawing.Point(764, 77)
        Me.historyListBox.Name = "historyListBox"
        Me.historyListBox.Size = New System.Drawing.Size(151, 460)
        Me.historyListBox.TabIndex = 10
        '
        'btnRowInsert
        '
        Me.btnRowInsert.Location = New System.Drawing.Point(764, 574)
        Me.btnRowInsert.Name = "btnRowInsert"
        Me.btnRowInsert.Size = New System.Drawing.Size(55, 29)
        Me.btnRowInsert.TabIndex = 11
        Me.btnRowInsert.Text = "行挿入"
        Me.btnRowInsert.UseVisualStyleBackColor = True
        '
        'btnRowDelete
        '
        Me.btnRowDelete.Location = New System.Drawing.Point(818, 574)
        Me.btnRowDelete.Name = "btnRowDelete"
        Me.btnRowDelete.Size = New System.Drawing.Size(55, 29)
        Me.btnRowDelete.TabIndex = 12
        Me.btnRowDelete.Text = "行削除"
        Me.btnRowDelete.UseVisualStyleBackColor = True
        '
        'namLabel
        '
        Me.namLabel.AutoSize = True
        Me.namLabel.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.namLabel.ForeColor = System.Drawing.Color.Blue
        Me.namLabel.Location = New System.Drawing.Point(186, 14)
        Me.namLabel.Name = "namLabel"
        Me.namLabel.Size = New System.Drawing.Size(0, 19)
        Me.namLabel.TabIndex = 13
        '
        'dgvShtM
        '
        Me.dgvShtM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShtM.Location = New System.Drawing.Point(168, 151)
        Me.dgvShtM.Name = "dgvShtM"
        Me.dgvShtM.RowTemplate.Height = 21
        Me.dgvShtM.Size = New System.Drawing.Size(569, 543)
        Me.dgvShtM.TabIndex = 20
        '
        'SS生活の様子
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(953, 706)
        Me.Controls.Add(Me.dgvShtM)
        Me.Controls.Add(Me.namLabel)
        Me.Controls.Add(Me.btnRowDelete)
        Me.Controls.Add(Me.btnRowInsert)
        Me.Controls.Add(Me.historyListBox)
        Me.Controls.Add(Me.btnTextClear)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.residentListBox)
        Me.Name = "SS生活の様子"
        Me.Text = "SS生活の様子_期間別_"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        CType(Me.dgvShtM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents residentListBox As System.Windows.Forms.ListBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnTextClear As System.Windows.Forms.Button
    Friend WithEvents historyListBox As System.Windows.Forms.ListBox
    Friend WithEvents btnRowInsert As System.Windows.Forms.Button
    Friend WithEvents btnRowDelete As System.Windows.Forms.Button
    Friend WithEvents namLabel As System.Windows.Forms.Label
    Friend WithEvents benYmdBox As ymdBox.ymdBox
    Friend WithEvents bathYmdBox As ymdBox.ymdBox
    Friend WithEvents endYmdBox As ymdBox.ymdBox
    Friend WithEvents firstYmdBox As ymdBox.ymdBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tantoBox As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dateYmdBox As ymdBox.ymdBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgvShtM As Symphony_Journal.SSDataGridView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
End Class
