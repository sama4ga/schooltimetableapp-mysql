Public Class frmTermDetails
    Private Sub frmTermDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTerm.Text = My.Settings.term
        txtSession.Text = My.Settings.session
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Close()
        frmMisc.Show()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        My.Settings.session = txtSession.Text
        My.Settings.term = txtTerm.Text
        My.Settings.Save()
    End Sub

    Private Sub frmTermDetails_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmMisc.Show()
    End Sub
End Class