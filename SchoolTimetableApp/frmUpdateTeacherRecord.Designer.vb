<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateTeacherRecord
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdateTeacherRecord))
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.gbAvailableTeacherTime = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.btnDeleteTime = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.btnAddTime = New System.Windows.Forms.Button()
        Me.cmbDay = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.GroupBox3.SuspendLayout()
        Me.gbAvailableTeacherTime.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmbType)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.txtName)
        Me.GroupBox3.Controls.Add(Me.gbAvailableTeacherTime)
        Me.GroupBox3.Controls.Add(Me.btnUpdate)
        Me.GroupBox3.Location = New System.Drawing.Point(28, 19)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(651, 404)
        Me.GroupBox3.TabIndex = 24
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "New Teacher Details"
        '
        'cmbType
        '
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Items.AddRange(New Object() {"Full Time", "Part Time", "Furthering"})
        Me.cmbType.Location = New System.Drawing.Point(86, 75)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(183, 21)
        Me.cmbType.TabIndex = 30
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 78)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Type"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 13)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Name"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(84, 40)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(524, 20)
        Me.txtName.TabIndex = 3
        '
        'gbAvailableTeacherTime
        '
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label11)
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label10)
        Me.gbAvailableTeacherTime.Controls.Add(Me.btnDeleteTime)
        Me.gbAvailableTeacherTime.Controls.Add(Me.ListView1)
        Me.gbAvailableTeacherTime.Controls.Add(Me.btnAddTime)
        Me.gbAvailableTeacherTime.Controls.Add(Me.cmbDay)
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label9)
        Me.gbAvailableTeacherTime.Controls.Add(Me.txtTo)
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label7)
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label6)
        Me.gbAvailableTeacherTime.Controls.Add(Me.txtFrom)
        Me.gbAvailableTeacherTime.Controls.Add(Me.Label5)
        Me.gbAvailableTeacherTime.Location = New System.Drawing.Point(6, 122)
        Me.gbAvailableTeacherTime.Name = "gbAvailableTeacherTime"
        Me.gbAvailableTeacherTime.Size = New System.Drawing.Size(610, 225)
        Me.gbAvailableTeacherTime.TabIndex = 35
        Me.gbAvailableTeacherTime.TabStop = False
        Me.gbAvailableTeacherTime.Text = "Available Periods"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Label11.Location = New System.Drawing.Point(509, 99)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(95, 29)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Should not be greater than end time"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Label10.Location = New System.Drawing.Point(507, 54)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(95, 29)
        Me.Label10.TabIndex = 41
        Me.Label10.Text = "Should not be less than start time"
        '
        'btnDeleteTime
        '
        Me.btnDeleteTime.Location = New System.Drawing.Point(80, 195)
        Me.btnDeleteTime.Name = "btnDeleteTime"
        Me.btnDeleteTime.Size = New System.Drawing.Size(253, 30)
        Me.btnDeleteTime.TabIndex = 40
        Me.btnDeleteTime.Text = "Delete Selected Time"
        Me.btnDeleteTime.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(15, 30)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.ShowItemToolTips = True
        Me.ListView1.Size = New System.Drawing.Size(318, 161)
        Me.ListView1.TabIndex = 36
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'btnAddTime
        '
        Me.btnAddTime.Location = New System.Drawing.Point(412, 181)
        Me.btnAddTime.Name = "btnAddTime"
        Me.btnAddTime.Size = New System.Drawing.Size(89, 38)
        Me.btnAddTime.TabIndex = 31
        Me.btnAddTime.Text = "Add"
        Me.btnAddTime.UseVisualStyleBackColor = True
        '
        'cmbDay
        '
        Me.cmbDay.FormattingEnabled = True
        Me.cmbDay.Items.AddRange(New Object() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"})
        Me.cmbDay.Location = New System.Drawing.Point(412, 143)
        Me.cmbDay.Name = "cmbDay"
        Me.cmbDay.Size = New System.Drawing.Size(89, 21)
        Me.cmbDay.TabIndex = 36
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(359, 146)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(26, 13)
        Me.Label9.TabIndex = 39
        Me.Label9.Text = "Day"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(412, 99)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(89, 20)
        Me.txtTo.TabIndex = 32
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(368, 102)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(20, 13)
        Me.Label7.TabIndex = 38
        Me.Label7.Text = "To"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(346, 57)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "From"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(412, 54)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(89, 20)
        Me.txtFrom.TabIndex = 28
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(359, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Available Time"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(21, 369)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(133, 29)
        Me.btnUpdate.TabIndex = 23
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'frmUpdateTeacherRecord
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(705, 443)
        Me.Controls.Add(Me.GroupBox3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(200, 200)
        Me.Name = "frmUpdateTeacherRecord"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Update Teacher Record"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gbAvailableTeacherTime.ResumeLayout(False)
        Me.gbAvailableTeacherTime.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents cmbType As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtName As TextBox
    Friend WithEvents gbAvailableTeacherTime As GroupBox
    Friend WithEvents Label11 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents btnDeleteTime As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents btnAddTime As Button
    Friend WithEvents cmbDay As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtTo As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents txtFrom As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents btnUpdate As Button
End Class
