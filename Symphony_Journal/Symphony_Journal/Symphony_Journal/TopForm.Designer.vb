<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TopForm
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
        Me.topPicture = New System.Windows.Forms.PictureBox()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.rbtnPrintout = New System.Windows.Forms.RadioButton()
        Me.btnUnit = New System.Windows.Forms.Button()
        Me.btnHandover = New System.Windows.Forms.Button()
        Me.btnSeal = New System.Windows.Forms.Button()
        Me.btnResident = New System.Windows.Forms.Button()
        Me.btnProgress = New System.Windows.Forms.Button()
        Me.btnSSLife = New System.Windows.Forms.Button()
        CType(Me.topPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'topPicture
        '
        Me.topPicture.Location = New System.Drawing.Point(686, 27)
        Me.topPicture.Name = "topPicture"
        Me.topPicture.Size = New System.Drawing.Size(149, 172)
        Me.topPicture.TabIndex = 0
        Me.topPicture.TabStop = False
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Location = New System.Drawing.Point(686, 223)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(67, 16)
        Me.rbtnPreview.TabIndex = 1
        Me.rbtnPreview.Text = "プレビュー"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrintout
        '
        Me.rbtnPrintout.AutoSize = True
        Me.rbtnPrintout.Location = New System.Drawing.Point(775, 223)
        Me.rbtnPrintout.Name = "rbtnPrintout"
        Me.rbtnPrintout.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrintout.TabIndex = 2
        Me.rbtnPrintout.Text = "印刷"
        Me.rbtnPrintout.UseVisualStyleBackColor = True
        '
        'btnUnit
        '
        Me.btnUnit.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnUnit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(96, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(96, Byte), Integer))
        Me.btnUnit.Location = New System.Drawing.Point(12, 27)
        Me.btnUnit.Name = "btnUnit"
        Me.btnUnit.Size = New System.Drawing.Size(320, 95)
        Me.btnUnit.TabIndex = 3
        Me.btnUnit.Text = "ユニット日誌"
        Me.btnUnit.UseVisualStyleBackColor = True
        '
        'btnHandover
        '
        Me.btnHandover.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnHandover.ForeColor = System.Drawing.Color.FromArgb(CType(CType(148, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.btnHandover.Location = New System.Drawing.Point(338, 27)
        Me.btnHandover.Name = "btnHandover"
        Me.btnHandover.Size = New System.Drawing.Size(157, 95)
        Me.btnHandover.TabIndex = 4
        Me.btnHandover.Text = "申し送り"
        Me.btnHandover.UseVisualStyleBackColor = True
        '
        'btnSeal
        '
        Me.btnSeal.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnSeal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(168, Byte), Integer), CType(CType(96, Byte), Integer))
        Me.btnSeal.Location = New System.Drawing.Point(12, 128)
        Me.btnSeal.Name = "btnSeal"
        Me.btnSeal.Size = New System.Drawing.Size(157, 95)
        Me.btnSeal.TabIndex = 5
        Me.btnSeal.Text = "印鑑登録"
        Me.btnSeal.UseVisualStyleBackColor = True
        '
        'btnResident
        '
        Me.btnResident.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnResident.ForeColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(168, Byte), Integer), CType(CType(96, Byte), Integer))
        Me.btnResident.Location = New System.Drawing.Point(175, 128)
        Me.btnResident.Name = "btnResident"
        Me.btnResident.Size = New System.Drawing.Size(157, 95)
        Me.btnResident.TabIndex = 6
        Me.btnResident.Text = "入居者登録"
        Me.btnResident.UseVisualStyleBackColor = True
        '
        'btnProgress
        '
        Me.btnProgress.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnProgress.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(200, Byte), Integer))
        Me.btnProgress.Location = New System.Drawing.Point(338, 330)
        Me.btnProgress.Name = "btnProgress"
        Me.btnProgress.Size = New System.Drawing.Size(320, 95)
        Me.btnProgress.TabIndex = 8
        Me.btnProgress.Text = "施設介護支援経過表"
        Me.btnProgress.UseVisualStyleBackColor = True
        '
        'btnSSLife
        '
        Me.btnSSLife.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnSSLife.ForeColor = System.Drawing.Color.Magenta
        Me.btnSSLife.Location = New System.Drawing.Point(12, 432)
        Me.btnSSLife.Name = "btnSSLife"
        Me.btnSSLife.Size = New System.Drawing.Size(320, 95)
        Me.btnSSLife.TabIndex = 9
        Me.btnSSLife.Text = "ＳＳ生活の様子（期間別）"
        Me.btnSSLife.UseVisualStyleBackColor = True
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1019, 691)
        Me.Controls.Add(Me.btnSSLife)
        Me.Controls.Add(Me.btnProgress)
        Me.Controls.Add(Me.btnResident)
        Me.Controls.Add(Me.btnSeal)
        Me.Controls.Add(Me.btnHandover)
        Me.Controls.Add(Me.btnUnit)
        Me.Controls.Add(Me.rbtnPrintout)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.topPicture)
        Me.Name = "TopForm"
        Me.Text = "介護日誌"
        CType(Me.topPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents topPicture As System.Windows.Forms.PictureBox
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrintout As System.Windows.Forms.RadioButton
    Friend WithEvents btnUnit As System.Windows.Forms.Button
    Friend WithEvents btnHandover As System.Windows.Forms.Button
    Friend WithEvents btnSeal As System.Windows.Forms.Button
    Friend WithEvents btnResident As System.Windows.Forms.Button
    Friend WithEvents btnProgress As System.Windows.Forms.Button
    Friend WithEvents btnSSLife As System.Windows.Forms.Button

End Class
