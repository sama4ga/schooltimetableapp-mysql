Imports System.Windows.Forms.ListView
Imports MySql.Data.MySqlClient

Public Class frmEditClasses

    Dim sql As MysqlCon = New MysqlCon()
    Dim reader As MySqlDataReader
    Dim class_id As Integer = 0

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        frmMain.Show()
        Close()
    End Sub

    Private Sub lvClasses_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvClasses.SelectedIndexChanged
        For Each item In lvClasses.SelectedIndices
            txtClass.Text = lvClasses.Items(item).SubItems(0).Text
            cmbType.SelectedItem = lvClasses.Items(item).SubItems(1).Text
            class_id = Integer.Parse(lvClasses.Items(item).Name)
            'MsgBox(class_id)
        Next

    End Sub

    Private Sub frmEditClasses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvClasses.Columns.Clear()
        lvClasses.Columns.Add("Class", 200)
        lvClasses.Columns.Add("Subject Combination", 100)

        reader = sql.readData("SELECT * FROM `class`;")
        If Not sql.result Then
            Exit Sub
        End If
        If reader.HasRows Then
            lvClasses.Items.Clear()
            Dim lvitem As ListViewItem
            While reader.Read()
                lvitem = New ListViewItem()
                lvitem.SubItems(0).Text = (reader.GetString("name"))
                lvitem.SubItems.Add(reader.GetString("subject_combination"))
                lvitem.Name = reader.GetInt32("class_id")
                lvClasses.Items.Add(lvitem)
            End While

        End If
        reader.Close()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If class_id <> 0 Then
            sql.ExecuteQuery("UPDATE `class` SET `name`='" & Trim(txtClass.Text) & "',`subject_combination`='" & cmbType.SelectedItem.ToString() & "' WHERE `class_id`='" & class_id & "';")
            If Not sql.result Then
                Exit Sub
            End If
            MsgBox("Update successful")


            lvClasses.Columns.Clear()
            lvClasses.Columns.Add("Class", 200)
            lvClasses.Columns.Add("Subject Combination", 100)

            reader = sql.readData("SELECT * FROM `class`;")
            If Not sql.result Then
                Exit Sub
            End If
            If reader.HasRows Then
                lvClasses.Items.Clear()
                Dim lvitem As ListViewItem
                While reader.Read()
                    lvitem = New ListViewItem()
                    lvitem.SubItems(0).Text = (reader.GetString("name"))
                    lvitem.SubItems.Add(reader.GetString("subject_combination"))
                    lvitem.Name = reader.GetInt32("class_id")
                    lvClasses.Items.Add(lvitem)
                End While

            End If
            reader.Close()

        End If
    End Sub

    Private Sub lvClasses_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lvClasses.ItemSelectionChanged
        'If e.IsSelected Then
        '    txtClass.Text = lvClasses.SelectedItems(e.ItemIndex).SubItems(0).Text
        '    cmbType.SelectedItem = e.Item.SubItems(1).Text
        '    class_id = Integer.Parse(lvClasses.SelectedItems(e.ItemIndex).Name)
        '    MsgBox(class_id)
        'End If

    End Sub
End Class