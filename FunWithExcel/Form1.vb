Imports System.Windows.Forms
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.StartPosition = FormStartPosition.CenterScreen
        'Me.StartPosition = FormStartPosition.CenterParent
        If UBound(tmpT) > -1 And UBound(tmpUW) > -1 Then
            Dim i As Integer
            For i = 0 To UBound(tmpT)
                DataGridView2.Rows.Add()
                DataGridView2.Rows(i).Cells(0).Value = tmpT(i)
            Next
            For i = 0 To UBound(tmpUW)
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i).Cells(0).Value = tmpUW(i, 0)
                DataGridView1.Rows(i).Cells(1).Value = tmpUW(i, 1)
                DataGridView1.Rows(i).Cells(2).Value = tmpUW(i, 2)
            Next
        End If
    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged
        e.Row.HeaderCell.Value = (e.Row.Index + 1).ToString
    End Sub

    Private Sub DataGridView2_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView2.RowStateChanged
        e.Row.HeaderCell.Value = (e.Row.Index + 1).ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim row As Integer = 2
        Dim rowcun, rowcun1, i As Integer
        rowcun = DataGridView1.RowCount - 1
        rowcun1 = DataGridView2.RowCount - 1
        'ReDim ArrayStr(1 To rowcun, 1 To 3)
        ReDim tmpNumber(rowcun1 - 1)
        For i = 0 To rowcun1 - 1
            tmpNumber(i) = DataGridView2.Rows(i).Cells(0).Value
        Next
        ReDim ArrayStr(rowcun - 1, 2)
        For i = 0 To rowcun - 1
            ArrayStr(i, 0) = DataGridView1.Rows(i).Cells(0).Value
            ArrayStr(i, 1) = DataGridView1.Rows(i).Cells(1).Value
            ArrayStr(i, 2) = DataGridView1.Rows(i).Cells(2).Value
        Next
        ReDim tmpStrText(rowcun - 1, 1)
        For i = 0 To rowcun - 1
            If i = 0 Then
                tmpStrText(i, 0) = 2
                tmpStrText(i, 1) = (ArrayStr(i, 1) - ArrayStr(i, 0) + 1) * ArrayStr(i, 2) * 2 + 1
            Else
                tmpStrText(i, 0) = tmpStrText(i - 1, 1) + 1
                tmpStrText(i, 1) = tmpStrText(i, 0) - 1 + (ArrayStr(i, 1) - ArrayStr(i, 0) + 1) * ArrayStr(i, 2) * 2
            End If
        Next
        Me.Close()
    End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs)
    '    CopyExcelToGrid()
    'End Sub
    'Private Sub CopyExcelToGrid()
    '    Dim i, j As Integer
    '    Dim pRow, pCol As Integer
    '    Dim selectedCellCount As Integer
    '    Dim startRow, startCol, endRow, endCol As Integer
    '    Dim pasteText, strline, strVal As String
    '    Dim strlines, vals As String()
    '    Dim pasteData(,) As String
    '    Dim flag As Boolean = False
    '    ' 当前单元格是否选择的判断
    '    If DataGridView1.CurrentCell Is Nothing Then
    '        Return
    '    End If
    '    Dim insertRowIndex As Integer = DataGridView1.CurrentCell.RowIndex
    '    ' 获取DataGridView选择区域，并计算要复制的行列开始、结束位置
    '    startRow = 9999
    '    startCol = 9999
    '    endRow = 0
    '    endCol = 0
    '    selectedCellCount = DataGridView1.GetCellCount(DataGridViewElementStates.Selected)
    '    For i = 0 To selectedCellCount - 1
    '        startRow = Math.Min(DataGridView1.SelectedCells(i).RowIndex, startRow)
    '        startCol = Math.Min(DataGridView1.SelectedCells(i).ColumnIndex, startCol)
    '        endRow = Math.Max(DataGridView1.SelectedCells(i).RowIndex, endRow)
    '        endCol = Math.Max(DataGridView1.SelectedCells(i).ColumnIndex, endCol)
    '    Next
    '    ' 获取剪切板的内容，并按行分割
    '    pasteText = Clipboard.GetText()
    '    If String.IsNullOrEmpty(pasteText) Then
    '        Return
    '    End If
    '    pasteText = pasteText.Replace(vbCrLf, vbLf)
    '    ReDim strlines(0)
    '    strlines = pasteText.Split(vbLf)
    '    pRow = strlines.Length        '行数
    '    pCol = 0
    '    For Each strline In strlines
    '        ReDim vals(0)
    '        vals = strline.Split(New Char() {vbTab, vbCr, vbNullChar, vbNullString}, 256, StringSplitOptions.RemoveEmptyEntries) ' 按Tab分割数据
    '        pCol = Math.Max(vals.Length, pCol) '列数
    '    Next
    '    ReDim pasteData(pRow, pCol)
    '    pasteText = Clipboard.GetText()
    '    pasteText = pasteText.Replace(vbCrLf, vbLf)
    '    ReDim strlines(0)
    '    strlines = pasteText.Split(vbLf)
    '    i = 1
    '    For Each strline In strlines
    '        j = 1
    '        ReDim vals(0)
    '        strline.TrimEnd(New Char() {vbLf})
    '        vals = strline.Split(New Char() {vbTab, vbCr, vbNullChar, vbNullString}, 256, StringSplitOptions.RemoveEmptyEntries)
    '        For Each strVal In vals
    '            pasteData(i, j) = strVal
    '            j = j + 1
    '        Next
    '        i = i + 1
    '    Next
    '    flag = False
    '    For j = 1 To pCol
    '        If pasteData(pRow, j) <> "" Then
    '            flag = True
    '            Exit For
    '        End If
    '    Next
    '    If flag = False Then
    '        pRow = Math.Max(pRow - 1, 0)
    '    End If
    '    For i = 1 To endRow - startRow + 1
    '        Dim row As DataGridViewRow = DataGridView1.Rows(i + startRow - 1)
    '        If i <= pRow Then
    '            For j = 1 To endCol - startCol + 1
    '                If j <= pCol Then
    '                    row.Cells(j + startCol - 1).Value = pasteData(i, j)
    '                Else
    '                    Exit For
    '                End If
    '            Next
    '        Else
    '            Exit For
    '        End If
    '    Next
    'End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        '该函数还存在问题，需完善
        If (e.Modifiers = Keys.Control) And (e.KeyCode = Keys.C) Then
            '   Clipboard.SetText(this.dvCustomer.GetClipboardContent().GetData(DataFormats.Text).ToString()); 
            DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
            '  Clipboard.SetDataObject(Me.DataGridView1.GetClipboardContent())
            Clipboard.SetText(Me.DataGridView1.GetClipboardContent().GetData(DataFormats.Text).ToString())
        End If
        If (e.Modifiers = Keys.Control) And (e.KeyCode = Keys.V) Then
            Try
                Dim str As String
                str = Clipboard.GetText()
                If str = "" Then Exit Sub
                Dim lines As String()
                lines = str.Split(Chr(13))
                Dim line As String
                Dim Cells As String()
                Dim cell As String
                Dim i As Int32 = Me.DataGridView1.CurrentCell.RowIndex - 1
                Dim j0 As Int32 = Me.DataGridView1.CurrentCell.ColumnIndex
                Dim j As Int32
                For Each line In lines
                    i = i + 1
                    'If i < Me.DataGridView1.RowCount - 1 Then
                    '    Me.DataGridView1.Rows.Add()
                    'End If
                    If line.Trim() = "" Then Continue For
                    j = j0
                    Cells = line.Split(Chr(Keys.Tab))
                    For Each cell In Cells
                        Me.DataGridView1.Rows(i).Cells(j).Value = cell
                        j = j + 1
                    Next
                Next
            Catch ex As Exception
            End Try
        End If
        If (e.KeyCode = Keys.Delete) Then
            Dim i As Int32
            For i = 0 To Me.DataGridView1.SelectedCells.Count - 1
                Me.DataGridView1.SelectedCells.Item(i).Value = ""
            Next
        End If
    End Sub

    Private Sub 删除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem.Click
        For Each r As DataGridViewRow In DataGridView2.SelectedRows
            If Not r.IsNewRow Then
                DataGridView2.Rows.Remove(r)
                'DataGridView2.Refresh()
            End If
        Next
    End Sub

    Private Sub 删除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem1.Click
        For Each r As DataGridViewRow In DataGridView1.SelectedRows
            If Not r.IsNewRow Then
                DataGridView1.Rows.Remove(r)
                'DataGridView2.Refresh()
            End If
        Next
    End Sub
End Class