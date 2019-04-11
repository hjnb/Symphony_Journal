Public Class ExDataGridView
    Inherits DataGridView

    Public dt As New DataTable()

    Protected Overrides Function ProcessDialogKey(keyData As System.Windows.Forms.Keys) As Boolean
        If keyData = Keys.Enter Then
            Return Me.ProcessTabKey(keyData)
        Else
            Return MyBase.ProcessDialogKey(keyData)
        End If
    End Function

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            Return Me.ProcessTabKey(e.KeyCode)
        End If

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
            If e.RowIndex = 0 OrElse e.RowIndex = 18 OrElse (e.RowIndex >= 31 AndAlso e.ColumnIndex = 0) Then
                pParts = e.PaintParts And (Not DataGridViewPaintParts.Background And Not DataGridViewPaintParts.Border)
            Else
                pParts = e.PaintParts And Not DataGridViewPaintParts.Background
            End If
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    Public Overrides Function GetClipboardContent() As DataObject
        '元のDataObjectを取得する
        Dim oldData As DataObject = MyBase.GetClipboardContent()

        '新しいDataObjectを作成する
        Dim newData As New DataObject()

        'テキスト形式のデータをセットする（UnicodeTextもセットされる）
        newData.SetData(DataFormats.Text, oldData.GetData(DataFormats.Text))

        Return newData
    End Function

    Private Sub ExDataGridView_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'ctrl + x で選択セル内容切り取り
        If e.Control AndAlso e.KeyCode = Keys.X Then
            '選択セルの内容コピー
            Clipboard.SetDataObject(GetClipboardContent())

            '選択セルの内容削除
            For Each cell As DataGridViewCell In SelectedCells
                cell.Value = ""
            Next
        End If
    End Sub
End Class
