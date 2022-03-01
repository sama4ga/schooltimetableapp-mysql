Public Class frmUpdateTeacherRecord
    Public teacher_id As Integer = 0
    Dim sql As MysqlCon = New MysqlCon()

    Private Sub frmUpdateTeacherRecord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If teacher_id <> 0 Then
            sql.readData("SELECT t.`teacher_id`,t.`name`,ifnull(t.`Type`,'') as 'type',ifnull(tt.`from`,'') as 'from',ifnull(tt.`to`,'') as 'to',ifnull(tt.`day`,'') as 'day' FROM `teachers` t 
                          LEFT JOIN `teacher_available_time` tt ON t.`teacher_id`=tt.`teacher_id`
                          /* JOIN `teacher_class` tc ON t.`teacher_id`=tc.`teacher_id` */
                          WHERE t.`teacher_id`='" & teacher_id & "';")
            If Not sql.result Then
                Exit Sub
            End If
            If sql.rd.HasRows Then
                While sql.rd.Read()
                    txtName.Text = sql.rd.GetString("name")
                    cmbType.SelectedItem = sql.rd.GetString("type")
                    If Not sql.rd.GetString("type").Equals("Full Time") Then
                        ListView1.Items.Add(New ListViewItem(New String() {sql.rd.GetInt32("from").ToString(), sql.rd.GetInt32("to").ToString(), sql.rd.GetString("day")}))
                    End If
                End While

                If cmbType.SelectedIndex <> 0 Then
                    gbAvailableTeacherTime.Visible = True
                Else
                    gbAvailableTeacherTime.Visible = False
                End If
            End If
            sql.CloseCon()

            ListView1.Columns.Add("From")
            ListView1.Columns.Add("To")
            ListView1.Columns.Add("Day", 100, HorizontalAlignment.Center)

        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click


        sql.ExecuteQuery("UPDATE `teachers` SET `name`='" & Trim(txtName.Text) & "',`type`='" & cmbType.SelectedItem.ToString() & "' WHERE `teacher_id`='" & teacher_id & "';")
        If Not sql.result Then
            Exit Sub
        End If

        Dim ParamValues As List(Of String) = New List(Of String)
        If teacher_id <> 0 And cmbType.SelectedIndex <> 0 Then
            sql.ExecuteQuery("DELETE FROM `teacher_available_time` WHERE `teacher_id`='" & teacher_id & "'")
            If Not sql.result Then
                Exit Sub
            End If
            ParamValues.Clear()
            ParamValues.Add(teacher_id)
            ParamValues.Add(Trim(txtFrom.Text))
            ParamValues.Add(Trim(txtTo.Text))
            ParamValues.Add(cmbDay.SelectedItem.ToString())
            For Each item As ListViewItem In ListView1.Items
                sql.InsertQuery("INSERT INTO `teacher_available_time`(`teacher_id`,`from`,`to`,`day`) VALUES (@1,@2,@3,@4); ", ParamValues)
                If Not sql.result Then
                    Exit Sub
                End If
            Next
        End If
        MsgBox("Teacher record successfully updated",, "Update Teacher Record")
    End Sub

    Private Sub btnAddTime_Click(sender As Object, e As EventArgs) Handles btnAddTime.Click
        ListView1.Items.Add(New ListViewItem(New String() {Trim(txtFrom.Text), Trim(txtTo.Text), cmbDay.SelectedItem.ToString()}))

    End Sub

    Private Sub btnDeleteTime_Click(sender As Object, e As EventArgs) Handles btnDeleteTime.Click
        For index = 0 To ListView1.SelectedItems.Count - 1
            ListView1.Items.RemoveAt(ListView1.SelectedIndices(0))
        Next
        'ListView1.SelectedItems.Clear()
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        If cmbType.SelectedIndex <> 0 Then
            gbAvailableTeacherTime.Visible = True
        Else
            gbAvailableTeacherTime.Visible = False
        End If
    End Sub

    Private Sub frmUpdateTeacherRecord_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmManageTeachers.Show()
    End Sub
End Class