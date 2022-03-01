'Imports MySql.Data.MySqlClient

Public Class frmManageSubjects

    Dim sql As MysqlCon = New MysqlCon()
    Dim querylist As List(Of String) = New List(Of String)
    'Dim listClassType As List(Of String) = New List(Of String)


    Private Sub frmManageSubjects_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMain.Show()
    End Sub

    Private Sub txtTimesPerWeek_TextChanged(sender As Object, e As EventArgs) Handles txtTimesPerWeek.TextChanged
        If Not String.IsNullOrWhiteSpace(txtTimePerPeriod.Text) And Not String.IsNullOrWhiteSpace(txtTimesPerWeek.Text) Then
            txtTotalTimePerWeek.Text = txtTimesPerWeek.Text * txtTimePerPeriod.Text
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

        Dim subject_id As Long = 0
        Dim paramValues As List(Of String) = New List(Of String)

        'check if relevant details have been supplied
        If String.IsNullOrWhiteSpace(Trim(txtName.Text)) = True Then
            MsgBox("Subject name is a mandatory field.")
            txtName.Focus()
            Exit Sub
        ElseIf listClass.CheckedItems.Count < 1 Then
            MsgBox("Select atleast one class to proceed")
            listClass.Focus()
            Exit Sub
        ElseIf cmbCategory.SelectedIndex < 0 Then
            MsgBox("Select a category to proceed")
            cmbCategory.Focus()
            Exit Sub
        ElseIf Trim(txtTimePerPeriod.Text) < 1 Then
            MsgBox("Time per period cannot be less than 1 minute")
            txtTimePerPeriod.Focus()
            Exit Sub
        ElseIf Trim(txtTimesPerWeek.Text) < 1 Then
            MsgBox("Number of times per week cannot be less than 1")
            txtTimesPerWeek.Focus()
            Exit Sub
        ElseIf Trim(txtNoOfDoublePeriods.Text) < 0 Then
            MsgBox("Number of double periods cannot be less than 0")
            txtNoOfDoublePeriods.Focus()
            Exit Sub
        Else

            'check uif subject is already in subjects table
            sql.readData("SELECT * FROM `subjects` WHERE `name` = '" + Trim(txtName.Text) + "';")
            If Not sql.result Then
                Exit Sub
            End If
            If sql.rd.HasRows Then
                subject_id = sql.rd.GetInt32("subject_id")
                ' MsgBox("Subject already exists in the database")
                sql.CloseCon()
                'Exit Sub
            Else
                sql.CloseCon()

                Dim query As String = "INSERT INTO `subjects`(`name`) VALUES(@1);"
                paramValues.Clear()
                paramValues.Add(Trim(txtName.Text))
                subject_id = sql.InsertQuery(query, paramValues)


                Dim source As AutoCompleteStringCollection = New AutoCompleteStringCollection()
                sql.readData("SELECT * FROM `subjects`;")
                If Not sql.result Then
                    Exit Sub
                End If
                If sql.rd.HasRows() Then
                    While sql.rd.Read()
                        source.Add(sql.rd.GetString("name"))
                    End While
                    sql.CloseCon()
                End If

                With txtName
                    .AutoCompleteCustomSource = source
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .AutoCompleteSource = AutoCompleteSource.CustomSource
                End With

                'txtName.AutoCompleteSource = sql.CreateDatasource("SELECT * FROM `subjects`", "subjects")
                'txtName.DataBindings.Add("Text", sql.CreateDatasource("SELECT * FROM `subjects`;", "subjects"), "name")

            End If


            For Each index In listClass.CheckedIndices
                paramValues.Clear()
                paramValues.Add(subject_id)
                paramValues.Add(listClass.Items(index).Row(0).ToString())
                paramValues.Add(Trim(txtTimePerPeriod.Text))
                paramValues.Add(Trim(txtTimesPerWeek.Text))
                paramValues.Add(Trim(txtNoOfDoublePeriods.Text))
                paramValues.Add(Trim(txtTotalTimePerWeek.Text))
                If listClass.Items(index).Row(2).ToString() = "all" Then
                    paramValues.Add("General")
                Else
                    paramValues.Add(cmbCategory.SelectedItem.ToString())
                End If

                'MsgBox(listClass.Items(index).Row(0).ToString())
                sql.InsertQuery("INSERT INTO `subject_class`(`subject_id`,`class_id`,`time_per_period`,`no_times_per_week`,`no_double_period`,`total_time_per_week`,`category`) 
                                VALUES(@1,@2,@3,@4,@5,@6,@7);", paramValues)
                If Not sql.result Then
                    Exit Sub
                End If
            Next

            MsgBox("Subject successfully added")

            refresh_data()


            'Refresh()

        End If

    End Sub

    Private Sub txtNoOfDoublePeriods_TextChanged(sender As Object, e As EventArgs) Handles txtNoOfDoublePeriods.TextChanged
        If Not String.IsNullOrWhiteSpace(txtNoOfDoublePeriods.Text) And Not String.IsNullOrWhiteSpace(txtTimesPerWeek.Text) Then
            If CInt(txtNoOfDoublePeriods.Text) * 2 > CInt(txtTimesPerWeek.Text) Then
                MsgBox("Number of double periods exceeds number of times per week", MsgBoxStyle.Information, "Error")
            End If
        End If
    End Sub


    Private Sub frmManageSubjects_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql.LoadListBox("SELECT * FROM `class` ORDER BY `name`;", listClass, "class")
        If Not sql.result Then
            Exit Sub
        End If
        refresh_data()

        Dim source As AutoCompleteStringCollection = New AutoCompleteStringCollection()
        sql.readData("SELECT * FROM `subjects`;")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows() Then
            While sql.rd.Read()
                source.Add(sql.rd.GetString("name"))
            End While
            sql.CloseCon()
        End If

        With txtName
            .AutoCompleteCustomSource = source
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
        End With



        'txtName.AutoCompleteSource = sql.CreateDatasource("SELECT * FROM `subjects`", "subjects")
        'txtName.DataBindings.Add("Text", sql.CreateDatasource("SELECT * FROM `subjects`;", "subjects"), "name")


    End Sub

    Private Sub refresh_data()
        sql.LoadDatGridView("SELECT 
                            s.`subject_id`, s.`name` as 'Subject',c.`name` as 'Class',sc.`category` as 'Category',
                            sc.`no_times_per_week` as 'No of Times per Week',
                            sc.`no_double_period` as 'No of Double Periods',
                            sc.`time_per_period` as 'Time per Period(mins)',
                            sc.`total_time_per_week` as 'Total Time per Week(mins)'
                          FROM
                            `subjects` s
                                LEFT JOIN
                            `subject_class` sc ON s.`subject_id` = sc.`subject_id`
                                LEFT JOIN
                            `class` c ON sc.`class_id` = c.`class_id` ORDER BY c.`name`;", dgvSubjectList, "subject")

        If Not sql.result Then
            Exit Sub
        End If
        dgvSubjectList.Columns(0).Visible = False
        dgvSubjectList.Columns(1).Width = 200           'subject name
        dgvSubjectList.Columns(2).ReadOnly = True       'class
        'dgvSubjectList.Columns(3).ReadOnly = True       'category
        dgvSubjectList.Columns(6).ReadOnly = True       'time per period
        dgvSubjectList.Columns(7).ReadOnly = True       'total time per week

    End Sub

    Private Sub dtpSubjectList_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSubjectList.CellMouseEnter

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        refresh_data()
    End Sub

    Private Sub btnDeleteAll_Click(sender As Object, e As EventArgs) Handles btnDeleteAll.Click
        'delete for all classes
        Dim id As Integer = dgvSubjectList.CurrentRow.Cells(0).Value
        Dim subj As String = dgvSubjectList.CurrentRow.Cells(1).Value.ToString()
        If MsgBox("Do you really want to delete " & subj & " for all classes", MsgBoxStyle.YesNo & MsgBoxStyle.Question) = MsgBoxResult.Yes Then
            sql.ExecuteQuery("DELETE FROM `subject_class` WHERE `subject_id`='" & id & "'")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Subject successfully deleted")
            refresh_data()
        End If

    End Sub

    Private Sub dgvSubjectList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSubjectList.CellValueChanged
        'Dim headerText As String = dgvSubjectList.Columns(e.ColumnIndex).HeaderText
        Dim value As String = dgvSubjectList.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
        Dim subject_id As String = dgvSubjectList.Rows(e.RowIndex).Cells(0).Value.ToString()
        Dim sub_class As String = dgvSubjectList.CurrentRow.Cells(2).Value
        Dim class_id As Integer = 0
        'get the class id
        sql.readData("SELECT `class_id` FROM `class` WHERE `name`='" & sub_class & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows Then
            sql.rd.Read()
            class_id = sql.rd.GetInt32("class_id")
            sql.CloseCon()
        End If

        If e.ColumnIndex = 1 Then
            sql.ExecuteQuery("UPDATE `subjects` SET `name`='" & value & "' WHERE `subject_id`='" & subject_id & "' AND `class_id`='" & class_id & "';")
            If Not sql.result Then
                Exit Sub
            End If
        ElseIf e.ColumnIndex > 3

            Dim no_times_per_week As Integer = CInt(dgvSubjectList.Rows(e.RowIndex).Cells(4).Value.ToString())
            Dim no_of_double_periods As Integer = CInt(dgvSubjectList.Rows(e.RowIndex).Cells(5).Value.ToString())
            Dim time_per_period As Integer = CInt(dgvSubjectList.Rows(e.RowIndex).Cells(6).Value.ToString())
            Dim total_time_per_week As Integer = time_per_period * no_times_per_week

            'If e.ColumnIndex = 5 Then
            If no_of_double_periods * 2 > no_times_per_week Then
                MsgBox("Number of double periods exceeds number of times per week", MsgBoxStyle.Information, "Error")
                Exit Sub
            End If

            sql.ExecuteQuery("UPDATE `subject_class` SET `no_double_period`='" & no_of_double_periods & "',
                                  `no_times_per_week` = '" & no_times_per_week & "',
                                  `time_per_period` = '" & time_per_period & "',
                                  `total_time_per_week` = '" & total_time_per_week & "'
                                  WHERE `subject_id`='" & subject_id & "' AND `class_id`='" & class_id & "';")
            'End If
            If Not sql.result Then
                Exit Sub
            End If
        ElseIf e.ColumnIndex = 3
            sql.ExecuteQuery("UPDATE `subject_class` SET `category`='" & value & "' WHERE `subject_id`='" & subject_id & "' AND `class_id`='" & class_id & "';")
            If Not sql.result Then
                Exit Sub
            End If
        End If

        MsgBox("Update successful")
        'refresh_data()
    End Sub

    Private Sub dgvSubjectList_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSubjectList.CellEndEdit
        'Label10.Text = "Edit ended"
    End Sub


    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim id As Integer = dgvSubjectList.CurrentRow.Cells(0).Value
        Dim sub_class As String = dgvSubjectList.CurrentRow.Cells(2).Value
        Dim subj As String = dgvSubjectList.CurrentRow.Cells(1).Value

        'get the class id
        sql.readData("SELECT `class_id` FROM `class` WHERE `name`='" & sub_class & "';")
        If Not sql.result Then
            Exit Sub
        End If
        If sql.rd.HasRows Then
            Dim class_id = sql.rd.GetInt32("class_id")
            sql.CloseCon()

            If MsgBox("Do you really want to delete " & subj & " for " & sub_class, MsgBoxStyle.YesNo & MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                sql.ExecuteQuery("DELETE FROM `subject_class` WHERE `subject_id`='" & id & "' AND `class_id`='" & class_id & "';")
                If Not sql.result Then
                    Exit Sub
                End If
                MsgBox("Subject successfully deleted")
                refresh_data()
            End If
        End If



    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

    End Sub

    Private Sub btnTruncate_Click(sender As Object, e As EventArgs) Handles btnTruncate.Click
        If MsgBox("Are you sure you want to truncate this table" & vbCrLf & "This will delete all data in the table", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Manage Subjects") = MsgBoxResult.Yes Then
            sql.ExecuteQuery("TRUNCATE TABLE `subject_class`;")
            If Not sql.result Then
                Exit Sub
            End If
            sql.ExecuteQuery("TRUNCATE TABLE `subjects`;")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Subjects successfully deleted")
            refresh_data()
        End If
    End Sub

    Private Sub dgvSubjectList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSubjectList.CellContentClick

    End Sub
End Class