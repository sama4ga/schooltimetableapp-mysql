Public Class frmManageTeachers

    Dim sql As MysqlCon = New MysqlCon()

    Private Sub frmManageTeachers_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMain.Show()
    End Sub

    Private Sub frmManageTeachers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql.LoadComboBox("SELECT * FROM `class`;", cmbClass, "class")
        If Not sql.result Then
            Exit Sub
        End If
        cmbClass.SelectedValue = cmbClass.Items(0).Row(0).ToString()
        refresh_teacher()
        refresh_data()

        ListView1.Columns.Add("From")
        ListView1.Columns.Add("To")
        ListView1.Columns.Add("Day", 100, HorizontalAlignment.Center)

        cmbType.SelectedIndex = 0
    End Sub

    Private Sub cmbClass_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbClass.SelectionChangeCommitted
        Dim class_id As Integer = CInt(cmbClass.Items(cmbClass.SelectedIndex).Row(0).ToString())

        sql.LoadComboBox("SELECT 
                            s.`subject_id`, s.`name`
                          FROM
                            `subjects` s
                                LEFT JOIN
                            `subject_class` sc ON s.`subject_id` = sc.`subject_id`
                                LEFT JOIN
                            `class` c ON sc.`class_id` = c.`class_id`
                          WHERE
                            c.`class_id`='" & class_id & "';", cmbSubjects, "subject")
        If Not sql.result Then
            Exit Sub
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If cmbName.SelectedIndex < 0 Then
            MsgBox("Select a teacher to proceed")
            cmbName.Focus()
        ElseIf cmbSubjects.SelectedIndex < 0 Then
            MsgBox("Select a subject to proceed")
            cmbSubjects.Focus()
        ElseIf cmbClass.SelectedIndex < 0 Then
            MsgBox("Select a class to proceed")
            cmbClass.Focus()
        Else

            ' check if this subject has already been added to this teacher
            sql.readData("SELECT * FROM `teacher_class` WHERE `teacher_id`='" & cmbName.SelectedValue.ToString() & "' AND
                        `subject_id`='" & cmbSubjects.SelectedValue.ToString() & "' AND
                        `class_id` = '" & cmbClass.SelectedValue.ToString() & "';")
            If Not sql.result Then
                Exit Sub
            End If
            If sql.rd.HasRows() Then
                MsgBox(cmbSubjects.SelectedItem.row(1).ToString() & " has already been added for " & cmbClass.SelectedItem.row(1).ToString())
                sql.CloseCon()
            Else

                sql.CloseCon()
                Dim ParamValues As List(Of String) = New List(Of String)
                ParamValues.Add(cmbName.SelectedValue.ToString())
                ParamValues.Add(cmbSubjects.SelectedValue.ToString())
                ParamValues.Add(cmbClass.SelectedValue.ToString())
                sql.InsertQuery("INSERT INTO `teacher_class`(`teacher_id`,`subject_id`,`class_id`) VALUES(@1,@2,@3);", ParamValues)
                If Not sql.result Then
                    Exit Sub
                End If
                refresh_data()

            End If

        End If
    End Sub

    Private Sub refresh_data()
        sql.LoadDatGridView("SELECT
	                            t.`teacher_id`,
                                t.`name` AS 'Name',
                                c.`name` AS 'Class',
                                s.`name` AS 'Subject'
                             FROM
                                `teachers` t
                                    RIGHT JOIN
                                `teacher_class` tc ON t.`teacher_id` = tc.`teacher_id`
                                    LEFT JOIN
                                `class` c ON c.`class_id` = tc.`class_id`
                                    LEFT JOIN
                                `subjects` s ON s.`subject_id` = tc.`subject_id`
                                ORDER BY c.`name`;", dgvTeachersList, "Teachers")
        If Not sql.result Then
            Exit Sub
        End If

        dgvTeachersList.Columns(0).Visible = False
        dgvTeachersList.Columns(1).Width = 200
        dgvTeachersList.Columns(2).ReadOnly = True
        dgvTeachersList.Columns(3).Width = 200
        dgvTeachersList.Columns(3).ReadOnly = True

    End Sub

    Private Sub btnAddTeacher_Click(sender As Object, e As EventArgs) Handles btnAddTeacher.Click

        Dim teacher_id As Long = 0

        sql.readData("SELECT * FROM `teachers` WHERE `name`='" & txtName.Text & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            MsgBox("Teacher already exists in  the database")
            sql.CloseCon()
        Else
            sql.CloseCon()
            Dim ParamValues As List(Of String) = New List(Of String)
            ParamValues.Add(txtName.Text)
            ParamValues.Add(cmbType.SelectedItem.ToString())
            teacher_id = Long.Parse(sql.InsertQuery("INSERT INTO `teachers`(`name`,`type`) VALUES(@1,@2);", ParamValues))
            If Not sql.result Then
                Exit Sub
            End If

            If teacher_id <> 0 And cmbType.SelectedIndex <> 0 Then
                ParamValues.Clear()
                ParamValues.Add(Trim(txtFrom.Text))
                ParamValues.Add(Trim(txtTo.Text))
                ParamValues.Add(cmbDay.SelectedItem.ToString())
                For Each item As ListViewItem In ListView1.Items
                    sql.InsertQuery("INSERT INTO `teacher_available_time`(`from`, `to`,`day`) VALUES (@1,@2,@3); ", ParamValues)
                    If Not sql.result Then
                        Exit Sub
                    End If
                Next
            End If
            MsgBox("Teacher successfully added",, "Manage Teachers")
            refresh_teacher()
        End If

    End Sub

    Private Sub refresh_teacher()
        sql.LoadComboBox("SELECT * FROM `teachers` ORDER BY `name`;", cmbName, "teacher")
        If Not sql.result Then
            Exit Sub
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim id As Integer = dgvTeachersList.CurrentRow.Cells(0).Value

        If MsgBox("Are you sure you want to delete teacher" & vbCrLf & "This will remove the teacher's name from the database and unassign all subjects assigned to the teacher", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Teachers") = MsgBoxResult.Yes Then
            sql.ExecuteQuery("DELETE FROM `teachers` WHERE `teacher_id`='" & id & "'")
            If Not sql.result Then
                Exit Sub
            End If
            sql.ExecuteQuery("DELETE FROM `teacher_class` WHERE `teacher_id`='" & id & "'")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Teacher successfully deleted")
            refresh_data()
            refresh_teacher()
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        refresh_data()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

    End Sub

    Private Sub dgvTeachersList_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTeachersList.CellEndEdit
        If e.RowIndex = 1 Then
            Dim teacher_name As String = dgvTeachersList.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
            Dim teacher_id As Integer = CInt(dgvTeachersList.Rows(e.RowIndex).Cells(0).Value.ToString())
            sql.ExecuteQuery("UPDATE `teachers` SET `name`='" & teacher_name & "' WHERE `teacher_id`='" & teacher_id & "';")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Update successful")
            refresh_data()
        End If
    End Sub

    Private Sub btnDeleteSubjectAllClassess_Click(sender As Object, e As EventArgs) Handles btnDeleteSubjectAllClassess.Click
        Dim id As Integer = dgvTeachersList.CurrentRow.Cells(0).Value
        Dim subject As String = dgvTeachersList.CurrentRow.Cells(3).Value.ToString()
        Dim teacher As String = dgvTeachersList.CurrentRow.Cells(2).Value.ToString()
        Dim teacher_class As String = dgvTeachersList.CurrentRow.Cells(2).Value.ToString()

        sql.readData("SELECT `subject_id` FROM `subjects` WHERE `name`='" & subject & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            sql.rd.Read()
            Dim subject_id As Integer = sql.rd.GetInt32("subject_id")
            sql.CloseCon()

            If MsgBox("Do you really want to unassign all subjects from " & teacher, MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Teachers") = MsgBoxResult.Yes Then
                sql.ExecuteQuery("DELETE FROM `teacher_class` WHERE `teacher_id`='" & id & "' AND `subject_id`='" & subject_id & "';")
                If Not sql.result Then
                    Exit Sub
                End If
                MsgBox(subject & " successfully removed from Teacher across all classes")
                refresh_data()
            End If

        End If

    End Sub

    Private Sub btnDeleteAllSubjectForClass_Click(sender As Object, e As EventArgs) Handles btnDeleteAllSubjectForClass.Click
        Dim id As Integer = dgvTeachersList.CurrentRow.Cells(0).Value
        Dim subject As String = dgvTeachersList.CurrentRow.Cells(3).Value.ToString()
        Dim teacher As String = dgvTeachersList.CurrentRow.Cells(1).Value.ToString()
        Dim teacher_class As String = dgvTeachersList.CurrentRow.Cells(2).Value.ToString()


        sql.readData("SELECT `class_id` FROM `class` WHERE `name`='" & teacher_class & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            sql.rd.Read()
            Dim class_id As Integer = sql.rd.GetInt32("class_id")
            sql.CloseCon()

            If MsgBox("Are you sure you want to unassign all subjects from " & teacher & " for " & teacher_class, MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Teachers") = MsgBoxResult.Yes Then
                sql.ExecuteQuery("DELETE FROM `teacher_class` WHERE `teacher_id`='" & id & "' AND `class_id`='" & class_id & "';")
                If Not sql.result Then
                    Exit Sub
                End If
                MsgBox("Subjects successfully removed from " & teacher & " for " & teacher_class)
                refresh_data()
            End If
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        txtName.Clear()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        cmbName.SelectedIndex = 0
        cmbClass.SelectedIndex = 0
        cmbSubjects.SelectedIndex = 0
    End Sub

    Private Sub btnTruncate_Click(sender As Object, e As EventArgs) Handles btnTruncate.Click

        If MsgBox("Are you sure you want to truncate this table?" & vbCrLf & "This will delete all records", MsgBoxStyle.Question & MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            sql.ExecuteQuery("TRUNCATE TABLE `teacher_class`;")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Truncate operation successful")
            refresh_data()

        End If
    End Sub

    Private Sub btnTruncateTeachers_Click(sender As Object, e As EventArgs) Handles btnTruncateTeachers.Click
        If MsgBox("Are you sure you want to truncate this table?" & vbCrLf & "This will delete all records of teachers", MsgBoxStyle.Question & MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            sql.ExecuteQuery("TRUNCATE TABLE `teachers`;")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Teacher List successfully truncated")
            If Not sql.result Then
                Exit Sub
            End If
            refresh_teacher()
        End If
    End Sub

    Private Sub dgvTeachersList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTeachersList.CellContentClick

    End Sub

    Private Sub btnDeleteSelection_Click(sender As Object, e As EventArgs) Handles btnDeleteSelection.Click
        Dim id As Integer = dgvTeachersList.CurrentRow.Cells(0).Value
        Dim subject As String = dgvTeachersList.CurrentRow.Cells(3).Value.ToString()
        Dim teacher As String = dgvTeachersList.CurrentRow.Cells(1).Value.ToString()
        Dim teacher_class As String = dgvTeachersList.CurrentRow.Cells(2).Value.ToString()
        Dim subject_id As Integer = 0
        Dim class_id As Integer = 0


        sql.readData("SELECT `class_id` FROM `class` WHERE `name`='" & teacher_class & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            sql.rd.Read()
            class_id = sql.rd.GetInt32("class_id")
            sql.CloseCon()
        End If

        sql.readData("SELECT `subject_id` FROM `subjects` WHERE `name`='" & subject & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            sql.rd.Read()
            subject_id = sql.rd.GetInt32("subject_id")
            sql.CloseCon()
        End If

        If class_id <> 0 And subject_id <> 0 Then
            If MsgBox("Are you sure you want to unassign " & subject & " from " & teacher & " for " & teacher_class, MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Teachers") = MsgBoxResult.Yes Then
                sql.ExecuteQuery("DELETE FROM `teacher_class` WHERE `teacher_id`='" & id & "' AND `class_id`='" & class_id & "' AND `subject_id`='" & subject_id & "';")
                If Not sql.result Then
                    Exit Sub
                End If
                MsgBox("Subject successfully removed ")
                refresh_data()
            End If
        End If
    End Sub

    Private Sub btnAddTime_Click(sender As Object, e As EventArgs) Handles btnAddTime.Click

        Dim arrLVItem As ListViewItem = New ListViewItem  '(New String() {txtFrom.Text, txtTo.Text, cmbDay.SelectedItem.ToString()})
        arrLVItem.SubItems(0).Text = txtFrom.Text
        arrLVItem.SubItems.Add(txtTo.Text)
        arrLVItem.SubItems.Add(cmbDay.SelectedItem.ToString())
        ListView1.Items.Add(arrLVItem)
    End Sub

    Private Sub btnDeleteTime_Click(sender As Object, e As EventArgs) Handles btnDeleteTime.Click
        For index = 0 To ListView1.SelectedItems.Count - 1
            ListView1.Items.RemoveAt(ListView1.SelectedIndices(0))
        Next
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        If cmbType.SelectedIndex <> 0 Then
            gbAvailableTeacherTime.Visible = True
        Else
            gbAvailableTeacherTime.Visible = False
        End If
    End Sub

    Private Sub btnUpdateTeacher_Click(sender As Object, e As EventArgs) Handles btnUpdateTeacher.Click
        frmUpdateTeacherRecord.teacher_id = Integer.Parse(cmbName.SelectedValue.ToString())
        frmUpdateTeacherRecord.Show()
        Hide()
    End Sub

    Private Sub btnDeleteTeacherRecord_Click(sender As Object, e As EventArgs) Handles btnDeleteTeacherRecord.Click
        If MsgBox("Are you sure you want to delete teacher" & vbCrLf & "This will remove the teacher's name from the database and unassign all subjects assigned to the teacher", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Teachers") = MsgBoxResult.Yes Then
            sql.ExecuteQuery("DELETE FROM `teachers` WHERE `teacher_id`='" & cmbName.SelectedValue.ToString() & "'")
            If Not sql.result Then
                Exit Sub
            End If
            sql.ExecuteQuery("DELETE FROM `teacher_class` WHERE `teacher_id`='" & cmbName.SelectedValue.ToString() & "'")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Teacher successfully deleted")
            refresh_teacher()
            refresh_data()
        End If
    End Sub
End Class