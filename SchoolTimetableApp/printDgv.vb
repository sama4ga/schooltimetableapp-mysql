Imports System.Drawing.Printing

Public Class printDgv
    Inherits PrintDocument

    Private mRow As Integer = 0
    Private newpage As Boolean = True
    Private dgv As DataGridView
    Private fntPrintFont As Font

    Public Sub New(dgv As DataGridView)
        Me.dgv = dgv
    End Sub

    Public Property DataGridView As DataGridView
        Get
            Return dgv
        End Get
        Set(ByVal value As DataGridView)
            Me.dgv = value
        End Set
    End Property

    Public Property Font As Font
        ' Allows the user to override the default font
        Get
            Return fntPrintFont
        End Get
        Set(ByVal Value As Font)
            fntPrintFont = Value
        End Set
    End Property



    Protected Overrides Sub OnPrintPage(e As PrintPageEventArgs)

        MyBase.OnPrintPage(e)

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
        Dim groupRowHeight As Integer = 52
        Dim width As Integer = 90
        Dim height As Integer = 22

        ' print the header text for a new page
        '   use a grey bg just like the control
        If newpage Then

            e.Graphics.DrawString("PRINTED FROM SCHOOLTTA, A PRODUCT OF SAMASERVICES", New Font("Times New Romans", 6, FontStyle.Italic), Brushes.Gray, New Point(y, e.MarginBounds.Left))
            e.Graphics.DrawString("Date Printed: " & DateTime.Now().ToLongDateString(), New Font("Times New Romans", 6, FontStyle.Italic), Brushes.Gray, New Point(e.MarginBounds.Right - 200, e.MarginBounds.Left))

            Dim mBrush As Brush = Brushes.Black

            Dim fnt As Font = New Font("Times New Romans", 10, FontStyle.Regular)
            Dim leading As Long = 10

            Dim startX As Short = 0
            Dim startY As Short = leading
            Dim offset As Short = 10

            Dim formatCenter As StringFormat = New StringFormat(StringFormatFlags.NoClip)
            formatCenter.Alignment = StringAlignment.Center
            Dim layoutSize As SizeF = New SizeF(e.PageBounds.Width - offset * 2, leading)
            Dim Layout As RectangleF = New RectangleF(New PointF(startX, startY + offset), layoutSize)


            fnt = New Font("Times New Romans", 10, FontStyle.Bold)
            leading = fnt.GetHeight() + 5
            offset += leading
            layoutSize = New SizeF(e.MarginBounds.Width, leading)
            Layout = New RectangleF(New PointF(startX, startY + offset), layoutSize)
            e.Graphics.DrawString(My.Settings.school_name, fnt, mBrush, Layout, formatCenter)

            fnt = New Font("Times New Romans", 9, FontStyle.Regular)
            leading = fnt.GetHeight() + 5
            offset += leading
            layoutSize = New SizeF(e.MarginBounds.Width, leading)
            Layout = New RectangleF(New PointF(startX, startY + offset), layoutSize)
            e.Graphics.DrawString(My.Settings.school_address, fnt, mBrush, Layout, formatCenter)

            fnt = New Font("Times New Romans", 8, FontStyle.Italic)
            leading = fnt.GetHeight() + 5
            offset += leading
            layoutSize = New SizeF(e.MarginBounds.Width, leading)
            Layout = New RectangleF(New PointF(startX, startY + offset), layoutSize)
            e.Graphics.DrawString("Motto: " & My.Settings.school_motto, fnt, mBrush, Layout, formatCenter)

            fnt = New Font("Times New Romans", 10, FontStyle.Italic)
            leading = fnt.GetHeight() + 15
            offset += leading
            layoutSize = New SizeF(e.MarginBounds.Width, leading)
            Layout = New RectangleF(New PointF(startX, startY + offset), layoutSize)
            e.Graphics.DrawString("School Timetable for " & My.Settings.term & " term " & My.Settings.session & " academic session", fnt, mBrush, Layout, formatCenter)

            y += 90
            row = dgv.Rows(mRow)
            x = e.MarginBounds.Left
            For Each cell As DataGridViewCell In row.Cells
                ' since we are printing the control's view,
                ' skip invisible columns
                If cell.Visible Then
                    rc = New Rectangle(x, y, width, height)

                    e.Graphics.FillRectangle(Brushes.LightGray, rc)
                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    FormatCell(cell, fmt, rc)

                    e.Graphics.DrawString(dgv.Columns(cell.ColumnIndex).HeaderText, dgv.Font, Brushes.Black, rc, fmt)
                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If
            Next
            y += h

        End If
        newpage = False

        ' now print the data for each row
        Dim thisNDX As Integer
        For thisNDX = mRow To dgv.RowCount - 1

            ' no need to try to print the new row
            If dgv.Rows(thisNDX).IsNewRow Then Exit For

            row = dgv.Rows(thisNDX)
            x = e.MarginBounds.Left
            h = 0


            ' check for class rows
            If rowsToPaint.Contains(mRow) Then
                x = e.MarginBounds.Left
                h = 0
                rc = New Rectangle(x, y, x + row.Cells.Count * width, height)

                Using br As New SolidBrush(Color.MistyRose)
                    e.Graphics.FillRectangle(br, rc)
                    e.Graphics.DrawRectangle(Pens.Black, rc)
                    fmt.Alignment = StringAlignment.Center
                    e.Graphics.DrawString(row.Cells(0).Value.ToString(), dgv.Font, Brushes.Black, rc, fmt)
                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End Using
                GoTo c
            End If



            ' print the data
            For Each cell As DataGridViewCell In row.Cells

                If cell.Visible Then

                    If cell.Value Is Nothing Then cell.Value = ""

                    If cell.Value.ToString().Contains("/" & vbNewLine) Then
                        rc = New Rectangle(x, y, width, groupRowHeight)
                        groupRow = True
                    ElseIf cell.Value.ToString().Contains("/") Then
                        cell.Value = cell.Value.ToString.Replace("/", "/" & vbNewLine)
                        rc = New Rectangle(x, y, width, groupRowHeight)            'cell.Size.Width, cell.Size.Height + 30
                        groupRow = True
                    Else
                        rc = New Rectangle(x, y, width, height)
                    End If


                    ' SAMPLE CODE: How To setup a RowPrePaint rule
                    'If rowsToPaint.Contains(mRow) Then
                    '    Using br As New SolidBrush(Color.MistyRose)
                    '        e.Graphics.FillRectangle(br, rc)
                    '    End Using
                    'End If

                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    FormatCell(cell, fmt, rc)

                    e.Graphics.DrawString(cell.Value.ToString(), dgv.Font, Brushes.Black, rc, fmt)

                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If

            Next




            '' check for class rows
            'If rowsToPaint.Contains(mRow) Then
            '    x = e.MarginBounds.Left
            '    h = 0
            '    rc = New Rectangle(x, y, x + row.Cells.Count * width, height)

            '    Using br As New SolidBrush(Color.MistyRose)
            '        e.Graphics.FillRectangle(br, rc)
            '        e.Graphics.DrawRectangle(Pens.Black, rc)
            '        fmt.Alignment = StringAlignment.Center
            '        e.Graphics.DrawString(row.Cells(0).Value.ToString(), dgv.Font, Brushes.Black, rc, fmt)
            '        x += rc.Width
            '        h = Math.Max(h, rc.Height)
            '    End Using

            'End If




            ' check for other cells in groupRow that are not the same height
            If groupRow Then
                x = e.MarginBounds.Left
                h = 0

                For Each cell As DataGridViewCell In row.Cells

                    If cell.Visible Then

                        If cell.Size.Height < groupRowHeight Then
                            rc = New Rectangle(x, y, width, groupRowHeight)
                            e.Graphics.FillRectangle(New SolidBrush(Color.White), rc)
                            e.Graphics.DrawRectangle(Pens.Black, rc)
                            FormatCell(cell, fmt, rc)
                            e.Graphics.DrawString(cell.Value.ToString(), dgv.Font, Brushes.Black, rc, fmt)
                        End If

                        x += rc.Width
                        h = Math.Max(h, rc.Height)
                    End If
                Next
            End If
            groupRow = False

c:          y += h

            ' next row to print
            mRow = thisNDX + 1

            If y + h > e.MarginBounds.Bottom Or mRow = 18 Then
                e.HasMorePages = True
                ' mRow -= 1   causes last row to rePrint on next page
                newpage = True
                Return
            End If
        Next
    End Sub


    Protected Overrides Sub OnBeginPrint(e As Printing.PrintEventArgs)

        MyBase.OnBeginPrint(e)

        mRow = 0
        newpage = True

        ' Sets the default font
        If fntPrintFont Is Nothing Then
            fntPrintFont = New Font("Times New Roman", 12)
        End If


        DefaultPageSettings.Landscape = True
        DefaultPageSettings.Margins = New Margins(20, 20, 50, 15)
        DefaultPageSettings.PrinterSettings.PrintToFile = True
    End Sub


    Protected Overrides Sub OnEndPrint(e As PrintEventArgs)
        MyBase.OnEndPrint(e)

    End Sub

    Protected Overrides Sub OnQueryPageSettings(e As QueryPageSettingsEventArgs)
        MyBase.OnQueryPageSettings(e)

    End Sub

    Private Sub FormatCell(cell As DataGridViewCell, ByRef fmt As StringFormat, ByRef rc As Rectangle)

        Select Case dgv.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
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
