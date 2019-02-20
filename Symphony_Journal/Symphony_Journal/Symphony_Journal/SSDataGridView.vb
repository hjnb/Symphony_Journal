﻿Public Class SSDataGridView
    Inherits DataGridView

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

            'If e.RowIndex = 0 OrElse e.RowIndex = 18 OrElse (e.RowIndex >= 31 AndAlso e.ColumnIndex = 0) Then
            '    pParts = e.PaintParts And (Not DataGridViewPaintParts.Background And Not DataGridViewPaintParts.Border)
            'Else
            '    pParts = e.PaintParts And Not DataGridViewPaintParts.Background
            'End If
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

End Class
