Imports System.Drawing.Printing
Imports System.IO
Imports System.Text

Public Class frmAdjustTimetable

    Private mRow As Integer = 0
    Private newpage As Boolean = True


    Private Sub frmAdjustTimetable_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMain.Show()
    End Sub

    Private Sub frmAdjustTimetable_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dirInfo As DirectoryInfo = New DirectoryInfo(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Timetables")
        Dim files As FileInfo() = dirInfo.GetFiles()
        For Each file As FileInfo In files
            If file.Extension.Equals(".csv") Then
                tvAvailableTimetables.Nodes.Add(file.FullName, file.Name)
            End If
        Next

        FileSystemWatcher1.Path = dirInfo.FullName

    End Sub

    Private Sub LoadTimetable(file As String)

        If dgvTimetable.ColumnCount > 0 Then
            dgvTimetable.Columns.Clear()
        End If

        Dim reader As StreamReader = New StreamReader(file)
        Dim currentText As String = ""
        Dim data As String()
        Dim lineCount As Integer = 0
        'Grid1.Cols = 2
        'Grid1.Rows = 2

        While reader.EndOfStream = False

            currentText = reader.ReadLine()
            data = currentText.Split(",")

            If lineCount = 0 Then

                Dim colCount As Integer = 0
                For Each str As String In data

                    'Grid1.Column(colCount).CellType = FlexCell.CellTypeEnum.DefaultType
                    'Grid1.Cell(1, colCount).Text = str
                    'Grid1.Cols += 1

                    dgvTimetable.Columns.Add(str, str)
                    dgvTimetable.Columns(colCount).SortMode = DataGridViewColumnSortMode.NotSortable
                    colCount += 1
                Next

                lineCount += 1
            Else

                'Grid1.Rows += 1
                'Dim rowCount As Integer = 1
                'For col = 1 To Grid1.Cols - 2
                '    Grid1.Cell(Grid1.Rows - 1, col).Text = data(col - 1)
                'Next

                dgvTimetable.Rows.Add(data)
                lineCount += 1
            End If


        End While
    End Sub

    Private Sub tvAvailableTimetables_DoubleClick(sender As Object, e As EventArgs) Handles tvAvailableTimetables.DoubleClick
        LoadTimetable(tvAvailableTimetables.SelectedNode.Name)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Save(tvAvailableTimetables.SelectedNode.Name)

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Close()
        frmMain.Show()
    End Sub

    Private Sub btnSaveAs_Click(sender As Object, e As EventArgs) Handles btnSaveAs.Click

        Dim sfd As SaveFileDialog = New SaveFileDialog()
        With sfd
            .DefaultExt = ".csv"
            .AddExtension = True
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Timetables"
            .RestoreDirectory = True
            .Title = "Choose filename for the timetable"
            .Filter = "ALL FIles|*.*|CSV|*.csv"
            .FilterIndex = 2
        End With

        If sfd.ShowDialog = DialogResult.OK Then
            Save(sfd.FileName)
        End If
    End Sub

    Private Sub Save(filename As String)
        Dim dir As DirectoryInfo = Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Timetables")

        Dim Data As StringBuilder = New StringBuilder()
        Dim str As String = String.Empty

        For index = 0 To dgvTimetable.ColumnCount - 1
            If index = dgvTimetable.ColumnCount - 1 Then
                str &= dgvTimetable.Columns(index).Name
            Else
                str &= dgvTimetable.Columns(index).Name & ","
            End If
        Next
        Data.AppendLine(str)
        str = String.Empty

        For row = 0 To dgvTimetable.RowCount - 1

            If dgvTimetable.Rows(row).IsNewRow Then Exit For

            For col = 0 To dgvTimetable.ColumnCount - 1

                If col = dgvTimetable.ColumnCount - 1 Then

                    If dgvTimetable.Rows(row).Cells(col).Value = Nothing Then
                        dgvTimetable.Rows(row).Cells(col).Value = ""
                    End If
                    str &= dgvTimetable.Rows(row).Cells(col).Value.ToString()

                Else

                    If dgvTimetable.Rows(row).Cells(col).Value = Nothing Then
                        dgvTimetable.Rows(row).Cells(col).Value = ""
                    End If
                    str &= dgvTimetable.Rows(row).Cells(col).Value.ToString() & ","
                End If
            Next

            Data.AppendLine(str)
            str = String.Empty
        Next


        IO.File.WriteAllText(filename, Data.ToString(), Encoding.UTF8)
        MsgBox("Timetable successfully saved",, "Save Timetable")
    End Sub


    Private Sub FileSystemWatcher1_Changed(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Changed

        'MsgBox(e.Name & " has been modified")
    End Sub

    Private Sub FileSystemWatcher1_Deleted(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Deleted
        tvAvailableTimetables.Nodes.RemoveByKey(e.FullPath)
        MsgBox(e.Name & " has been deleted externally")
    End Sub

    Private Sub FileSystemWatcher1_Created(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Created
        tvAvailableTimetables.Nodes.Add(e.FullPath, e.Name)
        'MsgBox("New file " & e.Name & " created")
    End Sub

    Private Sub FileSystemWatcher1_Renamed(sender As Object, e As RenamedEventArgs) Handles FileSystemWatcher1.Renamed
        Dim index As Integer = tvAvailableTimetables.Nodes.IndexOfKey(e.OldFullPath)
        tvAvailableTimetables.Nodes(index).Name = e.FullPath
        tvAvailableTimetables.Nodes(index).Text = e.Name
        MsgBox("'" & e.OldName & "' has been renamed externally to '" & e.Name & "'")
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click

        'Grid1.PrintPreview("Print timetable", True, True, True, True, True, 1, 0, 0, 0, 0)
        'Grid1.ExportToPDF(tvAvailableTimetables.SelectedNode.Name)

        'Dim btnPreview As PrintPreviewDialog = New PrintPreviewDialog()
        'Dim printDoc As PrintDocument = New PrintDocument()
        'btnPreview.Document = printDoc

        'If btnPreview.ShowDialog = DialogResult.OK Then
        '    printDoc.Print()
        'End If


        Dim printObject As printDgv = New printDgv(dgvTimetable)
        'printObject.DefaultPageSettings.Landscape = True
        printObject.DocumentName = tvAvailableTimetables.SelectedNode.Text

        Dim pdb As PrintDialog = New PrintDialog()
        pdb.AllowPrintToFile = True
        pdb.AllowSomePages = True
        pdb.AllowCurrentPage = True
        pdb.AllowSelection = True
        pdb.Document = printObject

        'printObject.DefaultPageSettings.Margins = New Margins(0, 0, 0, 0)
        'printObject.DefaultPageSettings.PrinterSettings.PrintToFile = True

        'If printPreview.ShowDialog = DialogResult.OK Then
        If pdb.ShowDialog() = DialogResult.OK Then
            printObject.Print()
        End If
        'End If

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' sets it to show '...' for long text
        Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
        fmt.LineAlignment = StringAlignment.Center
        fmt.Trimming = StringTrimming.EllipsisCharacter
        Dim y As Int32 = e.MarginBounds.Top
        Dim rc As Rectangle
        Dim x As Int32
        Dim h As Int32 = 0
        Dim row As DataGridViewRow
        Dim rowsToPaint As Integer() = {0, 6, 12, 18, 24, 30}
        Dim groupRow As Boolean = False
        Dim groupRowHeight As Integer = 0

        ' print the header text for a new page
        '   use a grey bg just like the control
        If newpage Then
            row = dgvTimetable.Rows(mRow)
            x = e.MarginBounds.Left
            For Each cell As DataGridViewCell In row.Cells
                ' since we are printing the control's view,
                ' skip invisible columns
                If cell.Visible Then
                    rc = New Rectangle(x, y, 100, 22)

                    e.Graphics.FillRectangle(Brushes.LightGray, rc)
                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    ' reused in the data print - should be a function
                    'Select Case dgvTimetable.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
                    '    Case DataGridViewContentAlignment.BottomRight, DataGridViewContentAlignment.MiddleRight
                    '        fmt.Alignment = StringAlignment.Far
                    '        rc.Offset(-1, 0)
                    '    Case DataGridViewContentAlignment.BottomCenter, DataGridViewContentAlignment.MiddleCenter
                    '        fmt.Alignment = StringAlignment.Center
                    '    Case Else
                    '        fmt.Alignment = StringAlignment.Near
                    '        rc.Offset(2, 0)
                    'End Select
                    FormatCell(cell, fmt, rc)

                    e.Graphics.DrawString(dgvTimetable.Columns(cell.ColumnIndex).HeaderText, dgvTimetable.Font, Brushes.Black, rc, fmt)
                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If
            Next
            y += h

        End If
        newpage = False

        ' now print the data for each row
        Dim thisNDX As Integer
        For thisNDX = mRow To dgvTimetable.RowCount - 1
            ' no need to try to print the new row
            If dgvTimetable.Rows(thisNDX).IsNewRow Then Exit For

            row = dgvTimetable.Rows(thisNDX)
            x = e.MarginBounds.Left
            h = 0

            ' reset X for data
            ' x = e.MarginBounds.Left

            ' print the data
            For Each cell As DataGridViewCell In row.Cells

                If cell.Visible Then

                    If cell.Value.ToString().Contains("/" & vbNewLine) Then
                        rc = New Rectangle(x, y, 100, 52)
                        groupRow = True
                        groupRowHeight = 52
                    ElseIf cell.Value.ToString().Contains("/") Then
                        cell.Value = cell.Value.ToString.Replace("/", "/" & vbNewLine)
                        rc = New Rectangle(x, y, 100, 52)
                        groupRow = True
                        groupRowHeight = 52
                    Else
                        rc = New Rectangle(x, y, 100, 22)
                    End If


                    ' SAMPLE CODE: How To setup a RowPrePaint rule
                    If rowsToPaint.Contains(mRow) Then
                        Using br As New SolidBrush(Color.MistyRose)
                            e.Graphics.FillRectangle(br, rc)
                        End Using
                    End If

                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    'Select Case dgvTimetable.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
                    '    Case DataGridViewContentAlignment.BottomRight,
                    '         DataGridViewContentAlignment.MiddleRight
                    '        fmt.Alignment = StringAlignment.Far
                    '        rc.Offset(-1, 0)
                    '    Case DataGridViewContentAlignment.BottomCenter,
                    '        DataGridViewContentAlignment.MiddleCenter
                    '        fmt.Alignment = StringAlignment.Center
                    '    Case Else
                    '        fmt.Alignment = StringAlignment.Near
                    '        rc.Offset(2, 0)
                    'End Select
                    FormatCell(cell, fmt, rc)

                    e.Graphics.DrawString(cell.Value.ToString(), dgvTimetable.Font, Brushes.Black, rc, fmt)

                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If

            Next


            ' check for other cells in gropuRow that are not the same height
            If groupRow Then
                x = e.MarginBounds.Left
                h = 0

                For Each cell As DataGridViewCell In row.Cells

                    If cell.Visible Then

                        If cell.Size.Height < groupRowHeight Then
                            rc = New Rectangle(x, y, 100, groupRowHeight)
                            e.Graphics.FillRectangle(New SolidBrush(Color.White), rc)
                            e.Graphics.DrawRectangle(Pens.Black, rc)
                            FormatCell(cell, fmt, rc)
                            e.Graphics.DrawString(cell.Value.ToString(), dgvTimetable.Font, Brushes.Black, rc, fmt)
                        End If

                        x += rc.Width
                        h = Math.Max(h, rc.Height)
                    End If
                Next
            End If
            groupRow = False

            y += h
            ' next row to print
            mRow = thisNDX + 1

            If y + h > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                ' mRow -= 1   causes last row to rePrint on next page
                newpage = True
                Return
            End If
        Next
    End Sub

    Private Sub PrintDocument1_BeginPrint(sender As Object, e As Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        mRow = 0
        newpage = True
        PrintPreviewDialog1.PrintPreviewControl.StartPage = 0
        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1.0
    End Sub

    Private Sub FormatCell(cell As DataGridViewCell, ByRef fmt As StringFormat, ByRef rc As Rectangle)

        Select Case dgvTimetable.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
            Case DataGridViewContentAlignment.BottomRight, DataGridViewContentAlignment.MiddleRight
                fmt.Alignment = StringAlignment.Far
                rc.Offset(-1, 0)
            Case DataGridViewContentAlignment.BottomCenter, DataGridViewContentAlignment.MiddleCenter
                fmt.Alignment = StringAlignment.Center
            Case Else
                fmt.Alignment = StringAlignment.Near
                rc.Offset(2, 0)
        End Select

    End Sub
End Class