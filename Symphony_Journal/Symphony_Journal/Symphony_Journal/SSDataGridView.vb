Imports System.Text

Public Class SSDataGridView
    Inherits DataGridView

    Private Const LIMIT_LENGTH_BYTE As Integer = 82

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim tb As DataGridViewTextBoxEditingControl = CType(Me.EditingControl, DataGridViewTextBoxEditingControl)
        If Not IsNothing(tb) AndAlso ((e.KeyCode = Keys.Left AndAlso tb.SelectionStart = 0) OrElse (e.KeyCode = Keys.Right AndAlso tb.SelectionStart = tb.TextLength)) Then
            Return False
        Else
            Return MyBase.ProcessDataGridViewKey(e)
        End If
    End Function

    ''' <summary>
    ''' CellPaintingイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ExDataGridView_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles Me.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            Dim pParts As DataGridViewPaintParts
            If e.ColumnIndex = 0 Then
                pParts = e.PaintParts And (Not DataGridViewPaintParts.Background And Not DataGridViewPaintParts.Border)
            Else
                pParts = e.PaintParts And Not DataGridViewPaintParts.Background
            End If

            '区切りの縦黒線
            If e.ColumnIndex = 1 Then
                With e.CellBounds
                    .Offset(-2, 0)
                    e.Graphics.DrawLine(New Pen(Color.Black), .Left, .Top, .Left, .Bottom)
                End With
            End If

            '項目区切りの横黒線
            If e.ColumnIndex = 1 AndAlso (e.RowIndex = 4 OrElse e.RowIndex = 9 OrElse e.RowIndex = 14 OrElse e.RowIndex = 19) Then
                With e.CellBounds
                    .Offset(0, -2)
                    e.Graphics.DrawLine(New Pen(Color.Black), .Left - 80, .Bottom, .Right, .Bottom)
                End With
            End If

            'ページ区切りの横青線
            If e.ColumnIndex = 1 AndAlso e.RowIndex = 35 Then
                With e.CellBounds
                    .Offset(0, -15)
                    e.Graphics.DrawLine(New Pen(Color.Blue), .Left, .Bottom, .Right, .Bottom)
                End With
            End If
            
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    Private Sub dgv_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)

        'イベントハンドラを削除、追加
        RemoveHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
        AddHandler editTextBox.KeyPress, AddressOf dgvTextBox_KeyPress
    End Sub

    Private Sub dgvTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim text As String = CType(sender, DataGridViewTextBoxEditingControl).Text
        Dim lengthByte As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(text)

        If lengthByte >= LIMIT_LENGTH_BYTE Then '設定されているバイト数以上の時
            If e.KeyChar = ChrW(Keys.Back) Then
                'Backspaceは入力可能
                e.Handled = False
            Else
                '入力できなくする
                e.Handled = True
            End If
        End If
    End Sub

End Class
