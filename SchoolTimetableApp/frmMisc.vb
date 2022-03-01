Public Class frmMisc

    Dim sql As MysqlCon = New MysqlCon()

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        'My.Settings.start_time = dtpLStartTime.Value.ToString("HH:mm")
        'My.Settings.end_time = dtpLEndTime.Value.ToString("HH:mm")

        My.Settings.start_time = txtLStartTime.Text
        My.Settings.end_time = txtLEndTime.Text
        My.Settings.Save()

        MsgBox("Update Successful",, "Update Lesson Start and End Time")

    End Sub

    Private Sub frmMisc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'dtpLStartTime.Value = My.Settings.start_time
        'dtpLEndTime.Value = My.Settings.end_time

        txtLStartTime.Text = My.Settings.start_time
        txtLEndTime.Text = My.Settings.end_time
        refresh_data()
    End Sub


    Private Sub refresh_data()
        sql.LoadDatGridView("SELECT `id`,`details` as 'Name',`start_time` as 'Start Time',`end_time` as 'End Time',`days` as 'Days' FROM `misc` ORDER BY `details`;", dgvBreaks, "Breaks")
        If Not sql.result Then
            Exit Sub
        End If
        dgvBreaks.Columns(0).Visible = False
        dgvBreaks.Columns(1).Width = 200
        dgvBreaks.Columns(4).Width = 100

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim ParamValues As List(Of String) = New List(Of String)
        ParamValues.Add(txtBName.Text)

        'ParamValues.Add(dtpBStartTime.Value.ToString("HH:mm"))
        'ParamValues.Add(dtpBEndTme.Value.ToString("HH:mm"))

        ParamValues.Add(txtBStartTime.Text)
        ParamValues.Add(txtBEndTIme.Text)

        For Each index In listDays.CheckedIndices
            ParamValues.Add(listDays.Items(index).ToString())
            sql.InsertQuery("INSERT INTO `misc`(`details`,`start_time`,`end_time`,`days`) VALUES(@1,@2,@3,@4);", ParamValues)
            If Not sql.result Then
                Exit Sub
            End If
            ParamValues.RemoveAt(ParamValues.IndexOf(ParamValues.Last))
        Next

        MsgBox("Breaks successfully added",, "Add breaks")

        refresh_data()
    End Sub

    Private Sub frmMisc_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMain.Show()
    End Sub

    Private Sub btnBDelete_Click(sender As Object, e As EventArgs) Handles btnBDelete.Click
        sql.ExecuteQuery("DELETE FROM `misc` WHERE `id`='" & dgvBreaks.CurrentRow.Cells(0).Value.ToString() & "';")
        'sql.ds.Tables("Breaks").Rows(dgvBreaks.CurrentCell.RowIndex).Delete()
        If Not sql.result Then
            Exit Sub
        End If
        refresh_data()
        MsgBox("Break successful deleted",, "Delete Breaks")
    End Sub

    Private Sub btnBUpdate_Click(sender As Object, e As EventArgs) Handles btnBUpdate.Click
        sql.ExecuteQuery("UPDATE `misc` SET `details`='" & dgvBreaks.CurrentRow.Cells(1).Value.ToString() & "',
                         `start_time`='" & dgvBreaks.CurrentRow.Cells(2).Value.ToString() & "',
                         `end_time`='" & dgvBreaks.CurrentRow.Cells(3).Value.ToString() & "',
                         `days`='" & dgvBreaks.CurrentRow.Cells(4).Value.ToString() & "' 
                         WHERE `id`='" & dgvBreaks.CurrentRow.Cells(0).Value.ToString() & "';")
        'sql.ds.Tables("Breaks").AcceptChanges()
        If Not sql.result Then
            Exit Sub
        End If
        refresh_data()
        MsgBox("Break successful updated",, "Update Breaks")
    End Sub

    Private Sub btnBAdd_Click(sender As Object, e As EventArgs)
        sql.ds.Tables("Breaks").AcceptChanges()
        MsgBox("Break successful added",, "Add Breaks")
    End Sub

    Private Sub btnBTruncate_Click(sender As Object, e As EventArgs) Handles btnBTruncate.Click
        sql.ExecuteQuery("TRUNCATE TABLE `misc`;")
        'sql.ds.Tables("Breaks").Clear()
        If Not sql.result Then
            Exit Sub
        End If
        refresh_data()
        MsgBox("Break successful truncated",, "Truncate Breaks")
    End Sub

    Private Sub dgvBreaks_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBreaks.CellValueChanged
        sql.ExecuteQuery("UPDATE `misc` SET `details`='" & dgvBreaks.Rows(e.RowIndex).Cells(1).Value.ToString() & "',
                         `start_time`='" & dgvBreaks.Rows(e.RowIndex).Cells(2).Value.ToString() & "',
                         `end_time`='" & dgvBreaks.Rows(e.RowIndex).Cells(3).Value.ToString() & "',
                         `days`='" & dgvBreaks.Rows(e.RowIndex).Cells(4).Value.ToString() & "' 
                         WHERE `id`='" & dgvBreaks.Rows(e.RowIndex).Cells(0).Value.ToString() & "';")
        'sql.ds.Tables("Breaks").AcceptChanges()
        If Not sql.result Then
            Exit Sub
        End If
        refresh_data()
        MsgBox("Break successful updated",, "Update Breaks")
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        refresh_data()
    End Sub

    Private Sub btnClassDetails_Click(sender As Object, e As EventArgs) Handles btnClassDetails.Click
        frmEditClasses.ShowDialog()
    End Sub

    Private Sub btnTermDetails_Click(sender As Object, e As EventArgs) Handles btnTermDetails.Click
        frmTermDetails.ShowDialog()
        Hide()
    End Sub
End Class