Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text
Imports System.Drawing.Printing
Imports System.Linq

Public Class frmGenerateTimeTable

    Dim sql As MysqlCon = New MysqlCon()
    Dim reader As MySqlDataReader

    Dim days As List(Of String) = New List(Of String)
    Dim start_time As String = String.Empty
    Dim end_time As String = ""
    Dim classes As List(Of String) = New List(Of String)
    Public subjects As List(Of subject) = New List(Of subject)
    Dim teachers As List(Of teacher) = New List(Of teacher)
    Dim breaks As List(Of break) = New List(Of break)
    Dim DoublePeriodsHolder As List(Of memory) = New List(Of memory)
    Dim activeSubjects As activeSubjects = New activeSubjects
    Dim group_subject As List(Of String) = New List(Of String)
    Dim group_category As List(Of String) = New List(Of String)

    Private mRow As Integer = 0
    Private newpage As Boolean = True

    Dim timetableFilename As String = ""

    Private Sub frmGenerateTimeTable_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMain.Show()
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Label1.Show()
        dgvTimeTable.Show()
    End Sub

    Private Sub frmGenerateTimeTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ' get days
        days.AddRange(New String() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"})



        ' get classes
        reader = sql.readData("SELECT * FROM `class` ORDER BY `name`;")
        If Not sql.result Then
            Exit Sub
        End If
        If reader.HasRows Then
            While reader.Read
                classes.Add(reader.GetString("name"))
            End While
            reader.Close()
        End If



        'get teachers
        reader = sql.readData("SELECT `teacher_id`,`name`,ifnull(`Type`,'') AS 'Type' FROM `teachers` ORDER BY `name`;")
        If Not sql.result Then
            Exit Sub
        End If
        If reader.HasRows Then
            Dim teacher As teacher
            While sql.rd.Read
                teacher = New teacher(reader.GetString("name"), reader.GetInt16("teacher_id"), reader.GetString("Type"))
                teachers.Add(teacher)
            End While
            reader.Close()
        End If



        'get teachers time
        For Each teacher As teacher In teachers
            If Not teacher.Type.Equals("Full Time") Then
                'get teachers
                reader = sql.readData("SELECT * FROM `teacher_available_time` WHERE `teacher_id`='" & teacher.Id & "';")
                If Not sql.result Then
                    Exit Sub
                End If
                If reader.HasRows Then
                    Dim available_time As teachers_avialable_time
                    While sql.rd.Read
                        available_time = New teachers_avialable_time(reader.GetInt32("from"), reader.GetInt32("to"), reader.GetString("day"))
                        teacher.AddAvailableTime(available_time)
                    End While
                    reader.Close()
                End If
            End If
        Next



        ' get subjects and other associated data
        reader = sql.readData("SELECT 
                                s.`subject_id`,
                                s.`name` AS 'Subject',
                                sc.`category` AS 'Category',
                                sc.`no_times_per_week` /*AS 'No. of Times per Week'*/,
                                sc.`no_double_period` /*AS 'No. of Double Periods'*/,
                                sc.`time_per_period` /*AS 'Time per Period(mins)'*/,
                                sc.`total_time_per_week` /*AS 'Total Time per Week(mins)'*/,
                                c.`name` AS 'Class',
                                c.`subject_combination` /* AS 'Subject Combination' */,
                                ifnull(t.`name`,'') AS 'Teacher',
                                ifnull(t.`teacher_id`,-1) AS 'teacher_id'
                              FROM
                                `subjects` s
                                    JOIN
                                `subject_class` sc ON s.`subject_id` = sc.`subject_id`
                                    LEFT JOIN
                                `teacher_class` tc ON tc.`subject_id` = sc.`subject_id`
                                    AND tc.`class_id` = sc.`class_id`
                                    JOIN
                                `class` c ON sc.`class_id` = c.`class_id`
                                    LEFT JOIN
                                `teachers` t ON t.`teacher_id` = tc.`teacher_id`
                              ORDER BY sc.`no_double_period` desc;")
        If Not sql.result Then
            Exit Sub
        End If
        If reader.HasRows Then
            Dim subj As subject
            Dim category As String = ""
            Dim unknown_count As Integer = 1

            While reader.Read

                If reader.GetString("subject_combination").Equals("all") Then
                    category = "General"
                Else
                    category = reader.GetString("category")
                End If

                subj = New subject(reader.GetString("Subject"), reader.GetInt16("time_per_period"), reader.GetInt16("no_times_per_week"), reader.GetString("Class"), reader.GetInt16("no_double_period"), category)

                For index = 0 To teachers.Count - 1
                    If reader.GetInt16("teacher_id") <> -1 And teachers(index).Id.Equals(reader.GetInt16("teacher_id")) = True Then
                        subj.Teacher = teachers(index)
                    End If
                Next

                If IsNothing(subj.Teacher) Then

                    For index = 0 To subjects.Count - 1
                        If subjects(index).Name = subj.Name Then
                            subj.Teacher = subjects(index).Teacher
                            GoTo t
                        End If
                    Next
                    subj.Teacher = New teacher("unknown" & unknown_count, "Full Time")
                    unknown_count += 1
                End If

t:              subjects.Add(subj)
            End While
            reader.Close()
        End If



        ' get breaks
        reader = sql.readData("SELECT * FROM `misc` ORDER BY `start_time` desc;")
        If Not sql.result Then
            Exit Sub
        End If
        If reader.HasRows Then
            Dim br As break
            While reader.Read
                br = New break(reader.GetString("details"), reader.GetString("start_time"), reader.GetString("end_time"), reader.GetString("days"))
                breaks.Add(br)
            End While
            reader.Close()
        End If



        start_time = My.Settings.start_time
        end_time = My.Settings.end_time




        ' start computing the timetable

        dgvTimeTable.Columns.Add("Days/Time", "Days/Time")
        dgvTimeTable.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        ' add rows for 5 days a week per class and extra 5 to hold the classes
        dgvTimeTable.Rows.Add(days.Count * classes.Count + classes.Count)

        Dim i As Integer = 0
        For index = 0 To classes.Count - 1
            dgvTimeTable.Rows(i).Cells(0).Value = classes(index).ToString()
            'dgvTimeTable.Rows(i).Cells(0).Style = DataGridViewCellStyle.BackColor = Color.Gray
            dgvTimeTable.Rows(i + 1).Cells(0).Value = days(0).ToString()
            dgvTimeTable.Rows(i + 2).Cells(0).Value = days(1).ToString()
            dgvTimeTable.Rows(i + 3).Cells(0).Value = days(2).ToString()
            dgvTimeTable.Rows(i + 4).Cells(0).Value = days(3).ToString()
            dgvTimeTable.Rows(i + 5).Cells(0).Value = days(4).ToString()
            i += 6
        Next

        'Exit Sub

        Dim break_count As Integer = 0
        Dim row_count As Integer = 0
        Dim col_count As Integer = 0
        Dim subject_count As Integer = 0
        Dim subj_added As Boolean = False
        Dim trials As Integer = 0
        Dim no_groups As Integer = 0




#Region "compute"
        For day_count = 0 To days.Count - 1
            col_count = 0

a:          While start_time.Equals(end_time) = False
                row_count = day_count + 1

#Region "check for break"
                break_count = 0
                For Each br As break In breaks
                    If br.StartTime.Equals(start_time) = True And br.Day.Equals(days(day_count)) = True Then

                        'add to time table and increment time
                        If day_count < 1 Then
                            dgvTimetable.Columns.Add(br.StartTime & "-" & br.EndTime, br.StartTime & "-" & br.EndTime)
                            dgvTimetable.Columns(col_count + 1).SortMode = DataGridViewColumnSortMode.NotSortable
                            'dgvTimeTable.Columns(col_count + 1).Width = 100
                        End If
                        col_count += 1

                        For index = 0 To classes.Count - 1
                            dgvTimeTable.Rows(row_count).Cells(col_count).Value = br.Name
                            row_count += 6
                        Next

                        'increase start time
                        start_time = br.EndTime

                        'remove the added break period from list of break periods
                        breaks.RemoveAt(break_count)
                        'breaks.Remove(br)

                        GoTo a
                        Exit For

                    End If
                    break_count += 1
                Next

#End Region

                ' no break, check for avilable subjects

                If day_count < 1 Then
                    dgvTimeTable.Columns.Add(start_time & "-" & IncreaseTime(start_time, subjects(subject_count).TimePerPeriod), start_time & "-" & IncreaseTime(start_time, subjects(subject_count).TimePerPeriod))        '"Period" & col_count + 1
                    dgvTimeTable.Columns(col_count + 1).SortMode = DataGridViewColumnSortMode.NotSortable
                    'dgvTimetable.Columns(col_count + 1).Width = 100
                End If
                col_count += 1



                For Each std_class As String In classes
                    subj_added = False
                    trials = 0
                    subject_count = 0

#Region "check memory for double periods"
                    If DoublePeriodsHolder.Count > 0 Then
                        For index = 0 To DoublePeriodsHolder.Count - 1
                            If DoublePeriodsHolder(index).SubjectClass.Equals(std_class) = True Then
                                'add subject
                                dgvTimeTable.Rows(row_count).Cells(col_count).Value = DoublePeriodsHolder(index).Name
                                row_count += 6

                                'add teacher to activeteacher list
                                DoublePeriodsHolder(index).Teacher.Active = True

                                'remove subject from list
                                DoublePeriodsHolder.RemoveAt(index)

                                'change added to true
                                subj_added = True

                                GoTo b

                                Exit For
                            End If
                        Next

                    End If
#End Region


#Region "check for subjects"
c:                  If subjects(subject_count).SubjectClass.Equals(std_class) Then

#Region "check for double periods"

                        'check if subject is active
                        If subjects(subject_count).ActiveStatus = False Then

                            'check if teacher is free
                            If subjects(subject_count).Teacher.Active = False Then

                                'check if teacher workload is reached: maximum workload = 5
                                If subjects(subject_count).Teacher.WorkLoad <= 5 Then

                                    'check if teacher is a part-time teacher
                                    If Not subjects(subject_count).Teacher.Type.Equals("Full Time") Then

                                        'teacher is part-time, check for available period
                                        Dim current_time As Integer() = ConvertToTime(start_time)
                                        Dim current_day As String = days(day_count)
                                        Dim result = From available_time As teachers_avialable_time In subjects(subject_count).Teacher.AvailableTime
                                                     Where available_time.Day.Equals(current_day) AndAlso current_time(0) >= available_time.StartTime AndAlso current_time(0) <= available_time.EndTime
                                                     Select available_time.StartTime & "-" & available_time.EndTime

                                        If result.Count < 1 Then
                                            GoTo s
                                        End If                                          'available time

                                    End If                                              'part_time


                                    'check if double period is adequately spaced              subjects(subject_count).LastAdded = -1 
                                    If day_count = 0 Or subjects(subject_count).LastAdded = -1 Or day_count > subjects(subject_count).LastAdded + 1 Then


                                        'check for double period
                                        'Dim current_time As Integer() = ConvertToTime(start_time)
                                        If subjects(subject_count).NoOfDoublePeriod > 0 And ConvertToTime(start_time)(0) <= 10 And ConvertToTime(start_time)(1) <= 10 Then
                                            Dim mem As memory
                                            'add the subject
                                            dgvTimetable.Rows(row_count).Cells(col_count).Value = subjects(subject_count).Name
                                            row_count += 6

                                            'add to memory
                                            mem = New memory(subjects(subject_count))
                                            DoublePeriodsHolder.Add(mem)

                                            'add teacher to activeteacher list
                                            subjects(subject_count).Teacher.Active = True
                                            subjects(subject_count).Teacher.WorkLoad += 2

                                            'add to activesubject list
                                            subjects(subject_count).ActiveStatus = True
                                            subjects(subject_count).LastAdded = day_count
                                            activeSubjects.Add(subjects(subject_count))

                                            'reduce no of double periods
                                            subjects(subject_count).NoOfDoublePeriod -= 1

                                            'reduce number of times per week
                                            subjects(subject_count).NoOfTimesPerWeek -= 2
                                            If subjects(subject_count).NoOfTimesPerWeek <= 0 Then
                                                subjects.RemoveAt(subject_count)
                                            End If

                                            'change added to true
                                            subj_added = True

                                            GoTo b
#End Region




#Region "no double period"

                                        Else

                                            'process group addition
                                            If no_groups < 6 Then
                                                If subjects(subject_count).Category <> "General" And day_count > 0 Then
                                                    Do
                                                        If subjects(subject_count).Category <> "General" And group_category.Contains(subjects(subject_count).Category) = False And group_subject.Contains(subjects(subject_count).Name) = False And subjects(subject_count).SubjectClass = std_class Then
                                                            If subjects(subject_count).ActiveStatus = False Then
                                                                If subjects(subject_count).Teacher.Active = False Then
                                                                    If subjects(subject_count).Teacher.WorkLoad <= 5 Then

                                                                        'check if teacher is a part-time teacher
                                                                        If Not subjects(subject_count).Teacher.Type.Equals("Full Time") Then

                                                                            'teacher is part-time, check for available period
                                                                            Dim current_time As Integer() = ConvertToTime(start_time)
                                                                            Dim current_day As String = days(day_count)
                                                                            Dim result = From available_time As teachers_avialable_time In subjects(subject_count).Teacher.AvailableTime
                                                                                         Where available_time.Day.Equals(current_day) AndAlso current_time(0) >= available_time.StartTime AndAlso current_time(0) <= available_time.EndTime
                                                                                         Select available_time.StartTime & "-" & available_time.EndTime

                                                                            If result.Count > 0 Then
                                                                                group_subject.Add(subjects(subject_count).Name)
                                                                                group_category.Add(subjects(subject_count).Category)

                                                                                'add teacher to activeteacher list
                                                                                subjects(subject_count).Teacher.Active = True
                                                                                subjects(subject_count).Teacher.WorkLoad += 1

                                                                                'add to activesubject list
                                                                                subjects(subject_count).ActiveStatus = True
                                                                                subjects(subject_count).LastAdded = -1
                                                                                activeSubjects.Add(subjects(subject_count))

                                                                                'reduce number of times per week
                                                                                subjects(subject_count).NoOfTimesPerWeek -= 1
                                                                                If subjects(subject_count).NoOfTimesPerWeek <= 0 Then
                                                                                    subjects.RemoveAt(subject_count)
                                                                                End If

                                                                            End If                                          'available time

                                                                        Else

                                                                            group_subject.Add(subjects(subject_count).Name)
                                                                            group_category.Add(subjects(subject_count).Category)

                                                                            'add teacher to activeteacher list
                                                                            subjects(subject_count).Teacher.Active = True
                                                                            subjects(subject_count).Teacher.WorkLoad += 1

                                                                            'add to activesubject list
                                                                            subjects(subject_count).ActiveStatus = True
                                                                            subjects(subject_count).LastAdded = -1
                                                                            activeSubjects.Add(subjects(subject_count))

                                                                            'reduce number of times per week
                                                                            subjects(subject_count).NoOfTimesPerWeek -= 1
                                                                            If subjects(subject_count).NoOfTimesPerWeek <= 0 Then
                                                                                subjects.RemoveAt(subject_count)
                                                                            End If

                                                                        End If                                              'part_time



                                                                    End If

                                                                End If

                                                            End If

                                                        End If

                                                        IncreaseSubjectCount(subject_count, subjects)

                                                        trials += 1
                                                        If trials > subjects.Count * 2 Then
                                                            Exit Do
                                                        End If


                                                    Loop Until group_category.Count > 3

                                                    'add the subject
                                                    Dim subj As String = ""
                                                    For index = 0 To group_subject.Count - 1
                                                        If index = 0 Then
                                                            subj = group_subject(0)
                                                        Else
                                                            subj &= "/" & group_subject(index)
                                                        End If
                                                    Next
                                                    dgvTimetable.Rows(row_count).Cells(col_count).Value = subj
                                                    row_count += 6

                                                    group_subject.Clear()
                                                    group_category.Clear()

                                                    ' increase no of groups
                                                    no_groups += 1

                                                    'change added to true
                                                    subj_added = True

                                                    GoTo b

                                                End If
                                            End If                                                                   ' group addition



                                            ' group limit reached

                                            'add the subject
                                            dgvTimetable.Rows(row_count).Cells(col_count).Value = subjects(subject_count).Name
                                            row_count += 6


                                            'add teacher to activeteacher list
                                            subjects(subject_count).Teacher.Active = True
                                            subjects(subject_count).Teacher.WorkLoad += 1

                                            'add to activesubject list
                                            subjects(subject_count).ActiveStatus = True
                                            ' subjects(subject_count).LastAdded = day_count
                                            subjects(subject_count).LastAdded = -1
                                            activeSubjects.Add(subjects(subject_count))

                                            'reduce number of times per week
                                            subjects(subject_count).NoOfTimesPerWeek -= 1
                                            If subjects(subject_count).NoOfTimesPerWeek <= 0 Then
                                                subjects.RemoveAt(subject_count)
                                            End If

                                            'change added to true
                                            subj_added = True

                                            GoTo b


                                        End If

                                    End If                                              ' spacing

                                End If                                                  ' workload

#End Region
                            End If                                                      ' active teacher

                        End If                                                         ' active subject

                    End If                                                               ' class

s:                  trials += 1
                    If trials > subjects.Count * 2 Then
                        'add the subject
                        dgvTimetable.Rows(row_count).Cells(col_count).Value = ""
                        row_count += 6
                        GoTo b
                    End If

                    'increase subject_count
                    IncreaseSubjectCount(subject_count, subjects)
                    GoTo c


#End Region
b:              Next std_class

                'clear activeteacher list
                activeSubjects.ClearTeacher()
                'activeTeachers.Clear()

                'change start_time
                start_time = IncreaseTime(start_time, subjects(subject_count).TimePerPeriod)

            End While                                                                       ' start time

            'clear activesubject list
            activeSubjects.ClearTeacher()
            activeSubjects.ClearTeacherWorkLoad()
            activeSubjects.ClearList()
            activeSubjects.Clear()
            no_groups = 0

            start_time = My.Settings.start_time

        Next day_count


#End Region




        If subjects.Count > 0 Then
            dgvExtraSubjects.Columns.Add("Subject", "Subject")
            dgvExtraSubjects.Columns.Add("Class", "Class")
            For index = 0 To subjects.Count - 1
                dgvExtraSubjects.Rows.Add(subjects(index).Name, subjects(index).SubjectClass)
            Next
            GroupBox1.Visible = True
        End If




    End Sub

    ''' <summary>
    ''' Converts a time in string to DateTime method and returns the hours and minutes in 24hour format
    ''' </summary>
    ''' <param name="time">value of the time as a <see cref="String"/>string</param>
    ''' <return>Time as an array containing hours and minutes</return>
    Private Function ConvertToTime(time As String) As Integer()
        Dim part As String() = time.Split(New Char() {":"}, StringSplitOptions.None)
        Dim hours As Integer = CInt(part(0))
        Dim minutes As Integer = CInt(part(1))
        Dim result As Integer() = New Integer() {hours, minutes}
        Return result
    End Function

    ''' <summary>
    ''' Increase time by a specified number of minutes 
    ''' </summary>
    ''' <param name="time">The time to be increased</param>
    ''' <param name="increment">Number of minutes to add to the time</param>
    ''' <returns>New time as string in the format HH:mm</returns>
    Private Function IncreaseTime(time As String, increment As Integer) As String
        Dim part As String() = time.Split(New Char() {":"}, StringSplitOptions.None)
        Dim hours As Integer = CInt(part(0))
        Dim minutes As Integer = CInt(part(1))
        minutes += increment
        If minutes >= 60 Then
            hours += 1
            minutes = minutes - 60
        End If

        Dim result As String = ""
        If minutes.ToString().Length < 2 Then
            result = hours & ":" & minutes & "0"
        Else
            result = hours & ":" & minutes
        End If
        Return result
    End Function

    ''' <summary>
    ''' Used to increase the subject count
    ''' </summary>
    ''' <param name="subject_count">the current subject_count</param>
    Private Function IncreaseSubjectCount(subject_count As Integer) As Integer
        subject_count += 1
        If subject_count > subjects.Count - 1 Then
            subject_count = 0
        End If

        Return subject_count
    End Function


    ''' <summary>
    ''' Used to increase the subject count
    ''' </summary>
    ''' <param name="subject_count">the current subject_count</param>
    Private Sub IncreaseSubjectCount(ByRef subject_count As Integer, ByRef subj As List(Of subject))
        subject_count += 1
        If subject_count > subj.Count - 1 Then
            subject_count = 0
        End If
    End Sub

    ''' <summary>
    ''' Searches for a teacher in a list that holds the details of teachers using the teacher's name
    ''' </summary>
    ''' <param name="teacherName">Name of teacher to search for</param>
    ''' <param name="teachers">List holding details of teachers</param>
    ''' <returns>Teacher</returns>
    Private Function CheckTeacher(teacherName As String, teachers As List(Of teacher)) As teacher
        Dim newteacher As teacher = New teacher(teacherName, "Full Time")
        Dim index As Integer = teachers.IndexOf(newteacher)
        Return teachers(index)
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim dir As DirectoryInfo = Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Timetables")

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
            'Save(sfd.FileName)
            timetableFilename = sfd.FileName.Split("\").Last
            'timetableFilename = "timetable" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".csv"

            Dim Data As StringBuilder = New StringBuilder()
            Dim str As String = String.Empty

            For index = 0 To dgvTimetable.ColumnCount - 1
                If index = dgvTimetable.ColumnCount - 1 Then
                    str &= dgvTimetable.Columns(index).Name '& "'" & ";" "'" & 
                Else
                    str &= dgvTimetable.Columns(index).Name & "," '& "'"  "'" &
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
                        str &= dgvTimetable.Rows(row).Cells(col).Value.ToString() '& "'"  & ";"   "'" &

                    Else
                        If dgvTimetable.Rows(row).Cells(col).Value = Nothing Then
                            dgvTimetable.Rows(row).Cells(col).Value = ""
                        End If
                        str &= dgvTimetable.Rows(row).Cells(col).Value.ToString() & ","   '& "'"   "'" &
                    End If
                Next

                Data.AppendLine(str)
                str = String.Empty
            Next


            IO.File.WriteAllText(dir.FullName & "\" & timetableFilename, Data.ToString(), Encoding.UTF8)
            MsgBox("Timetable successfully saved",, "Save Timetable")
        End If



    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim printObject As printDgv = New printDgv(dgvTimetable)
        'Dim printObject As PrintDocument = PrintDocument1
        printObject.DocumentName = timetableFilename
        'printObject.DefaultPageSettings.Landscape = True
        'printObject.DefaultPageSettings.Margins = New Margins(20, 20, 50, 15)
        'printObject.DefaultPageSettings.PrinterSettings.PrintToFile = True

        Dim printPreview As PrintPreviewDialog = New PrintPreviewDialog()
        printPreview.Document = printObject

        Dim pdb As PrintDialog = New PrintDialog()
        pdb.AllowPrintToFile = True
        pdb.AllowSomePages = True
        pdb.AllowCurrentPage = True
        pdb.AllowSelection = True
        pdb.Document = printObject

        'If printPreview.ShowDialog = DialogResult.OK Then
        If pdb.ShowDialog() = DialogResult.OK Then
            printObject.Print()
        End If
        'End If

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        '' sets it to show '...' for long text
        'Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
        'fmt.LineAlignment = StringAlignment.Center
        'fmt.Trimming = StringTrimming.EllipsisCharacter
        'Dim y As Int32 = e.MarginBounds.Top
        'Dim rc As Rectangle
        'Dim x As Int32
        'Dim h As Int32 = 0
        'Dim row As DataGridViewRow
        'Dim rowsToPaint As Integer() = {0, 6, 12, 18, 24, 30}
        'Dim groupRow As Boolean = False
        'Dim groupRowHeight As Integer = 52
        'Dim width As Integer = 90
        'Dim height As Integer = 22

        '' print the header text for a new page
        ''   use a grey bg just like the control
        'If newpage Then
        '    row = dgvTimeTable.Rows(mRow)
        '    x = e.MarginBounds.Left
        '    For Each cell As DataGridViewCell In row.Cells
        '        ' since we are printing the control's view,
        '        ' skip invisible columns
        '        If cell.Visible Then
        '            rc = New Rectangle(x, y, width, height)

        '            e.Graphics.FillRectangle(Brushes.LightGray, rc)
        '            e.Graphics.DrawRectangle(Pens.Black, rc)

        '            FormatCell(cell, fmt, rc)

        '            e.Graphics.DrawString(dgvTimeTable.Columns(cell.ColumnIndex).HeaderText, dgvTimeTable.Font, Brushes.Black, rc, fmt)
        '            x += rc.Width
        '            h = Math.Max(h, rc.Height)
        '        End If
        '    Next
        '    y += h

        'End If
        'newpage = False

        '' now print the data for each row
        'Dim thisNDX As Integer
        'For thisNDX = mRow To dgvTimeTable.RowCount - 1
        '    ' no need to try to print the new row
        '    If dgvTimeTable.Rows(thisNDX).IsNewRow Then Exit For

        '    row = dgvTimeTable.Rows(thisNDX)
        '    x = e.MarginBounds.Left
        '    h = 0

        '    ' reset X for data
        '    ' x = e.MarginBounds.Left

        '    ' print the data
        '    For Each cell As DataGridViewCell In row.Cells

        '        If cell.Visible Then

        '            If cell.Value Is Nothing Then cell.Value = ""

        '            If cell.Value.ToString().Contains("/" & vbNewLine) Then
        '                rc = New Rectangle(x, y, width, groupRowHeight)
        '                groupRow = True
        '            ElseIf cell.Value.ToString().Contains("/") Then
        '                cell.Value = cell.Value.ToString.Replace("/", "/" & vbNewLine)
        '                rc = New Rectangle(x, y, width, groupRowHeight)            'cell.Size.Width, cell.Size.Height + 30
        '                groupRow = True
        '            Else
        '                rc = New Rectangle(x, y, width, height)
        '            End If


        '            ' SAMPLE CODE: How To setup a RowPrePaint rule
        '            'If rowsToPaint.Contains(mRow) Then
        '            '    Using br As New SolidBrush(Color.MistyRose)
        '            '        e.Graphics.FillRectangle(br, rc)
        '            '    End Using
        '            'End If

        '            e.Graphics.DrawRectangle(Pens.Black, rc)

        '            FormatCell(cell, fmt, rc)

        '            e.Graphics.DrawString(cell.Value.ToString(), dgvTimeTable.Font, Brushes.Black, rc, fmt)

        '            x += rc.Width
        '            h = Math.Max(h, rc.Height)
        '        End If

        '    Next




        '    ' check for class rows
        '    If rowsToPaint.Contains(mRow) Then
        '        x = e.MarginBounds.Left
        '        h = 0
        '        rc = New Rectangle(x, y, x + row.Cells.Count * width, height)

        '        Using br As New SolidBrush(Color.MistyRose)
        '            e.Graphics.FillRectangle(br, rc)
        '            e.Graphics.DrawRectangle(Pens.Black, rc)
        '            fmt.Alignment = StringAlignment.Center
        '            e.Graphics.DrawString(row.Cells(0).Value.ToString(), dgvTimetable.Font, Brushes.Black, rc, fmt)
        '            x += rc.Width
        '            h = Math.Max(h, rc.Height)
        '        End Using

        '    End If




        '    ' check for other cells in groupRow that are not the same height
        '    If groupRow Then
        '        x = e.MarginBounds.Left
        '        h = 0

        '        For Each cell As DataGridViewCell In row.Cells

        '            If cell.Visible Then

        '                If cell.Size.Height < groupRowHeight Then
        '                    rc = New Rectangle(x, y, width, groupRowHeight)
        '                    e.Graphics.FillRectangle(New SolidBrush(Color.White), rc)
        '                    e.Graphics.DrawRectangle(Pens.Black, rc)
        '                    FormatCell(cell, fmt, rc)
        '                    e.Graphics.DrawString(cell.Value.ToString(), dgvTimeTable.Font, Brushes.Black, rc, fmt)
        '                End If

        '                x += rc.Width
        '                h = Math.Max(h, rc.Height)
        '            End If
        '        Next
        '    End If
        '    groupRow = False

        '    y += h
        '    ' next row to print
        '    mRow = thisNDX + 1

        '    If y + h > e.MarginBounds.Bottom Then
        '        e.HasMorePages = True
        '        ' mRow -= 1   causes last row to rePrint on next page
        '        newpage = True
        '        Return
        '    End If
        'Next
    End Sub

    Private Sub PrintDocument1_BeginPrint(sender As Object, e As Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        'mRow = 0
        'newpage = True
        'PrintPreviewDialog1.PrintPreviewControl.StartPage = 0
        'PrintPreviewDialog1.PrintPreviewControl.Zoom = 1.0
    End Sub

    Private Sub FormatCell(cell As DataGridViewCell, ByRef fmt As StringFormat, ByRef rc As Rectangle)

        Select Case dgvTimeTable.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
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

    Private Sub btnManuallyAdjust_Click(sender As Object, e As EventArgs) Handles btnManuallyAdjust.Click
        If String.IsNullOrEmpty(timetableFilename) Then
            If MsgBox("Are you sure you want to proceed without saving the timetable" & vbCrLf & vbCrLf & "The generated timetabel will be lost if you proceed", MsgBoxStyle.Question & MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmAdjustTimetable.Show()
                Close()
            Else
                btnSave.PerformClick()
            End If
        End If
    End Sub

    Private Sub btnReduce_Click(sender As Object, e As EventArgs) Handles btnReduce.Click
        If String.IsNullOrEmpty(timetableFilename) Then
            If MsgBox("Are you sure you want to proceed without saving the timetable" & vbCrLf & vbCrLf & "The generated timetabel will be lost if you proceed", MsgBoxStyle.Question & MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmManageSubjects.Show()
                Close()
            Else
                btnSave.PerformClick()
            End If
        End If
    End Sub
End Class