<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmManageTeachers
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmManageTeachers))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnDeleteSelection = New System.Windows.Forms.Button()
        Me.btnTruncate = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgvTeachersList = New System.Windows.Forms.DataGridView()
        Me.btnDeleteSubjectAllClassess = New System.Windows.Forms.Button()
        Me.btnDeleteAllSubjectForClass = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbName = New System.Windows.Forms.ComboBox()
        Me.cmbSubjects = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbClass = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnTruncateTeachers = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnAddTeacher = New System.Windows.Forms.Button()
        Me.gbAvailableTeacherTime = New System.Windows.Forms.GroupBox()
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
        Me.btnUpdateTeacher = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btnDeleteTeacherRecord = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvTeachersList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gbAvailableTeacherTime.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.btnDelete)
        Me.GroupBox2.Controls.Add(Me.btnDeleteSelection)
        Me.GroupBox2.Controls.Add(Me.btnTruncate)
        Me.GroupBox2.Controls.Add(Me.Panel1)
        Me.GroupBox2.Controls.Add(Me.btnDeleteSubjectAllClassess)
        Me.GroupBox2.Controls.Add(Me.btnDeleteAllSubjectForClass)
        Me.GroupBox2.Controls.Add(Me.btnRefresh)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.btnUpdate)
        Me.GroupBox2.Location = New System.Drawing.Point(640, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(727, 642)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "List of Teachers"
        '
        'btnDeleteSelection
        '
        Me.btnDeleteSelection.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteSelection.Location = New System.Drawing.Point(617, 75)
        Me.btnDeleteSelection.Name = "btnDeleteSelection"
        Me.btnDeleteSelection.Size = New System.Drawing.Size(110, 89)
        Me.btnDeleteSelection.TabIndex = 26
        Me.btnDeleteSelection.Text = "Unassign select subject"
        Me.btnDeleteSelection.UseVisualStyleBackColor = True
        '
        'btnTruncate
        '
        Me.btnTruncate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTruncate.Location = New System.Drawing.Point(617, 518)
        Me.btnTruncate.Name = "btnTruncate"
        Me.btnTruncate.Size = New System.Drawing.Size(110, 63)
        Me.btnTruncate.TabIndex = 24
        Me.btnTruncate.Text = "Truncate"
        Me.btnTruncate.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.dgvTeachersList)
        Me.Panel1.Location = New System.Drawing.Point(3, 28)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(608, 597)
        Me.Panel1.TabIndex = 9
        '
        'dgvTeachersList
        '
        Me.dgvTeachersList.AllowUserToAddRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.dgvTeachersList.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvTeachersList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvTeachersList.BackgroundColor = System.Drawing.Color.White
        Me.dgvTeachersList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTeachersList.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvTeachersList.Location = New System.Drawing.Point(5, 5)
        Me.dgvTeachersList.Name = "dgvTeachersList"
        Me.dgvTeachersList.RowHeadersVisible = False
        Me.dgvTeachersList.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.dgvTeachersList.RowTemplate.DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvTeachersList.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.dgvTeachersList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTeachersList.Size = New System.Drawing.Size(600, 589)
        Me.dgvTeachersList.TabIndex = 5
        '
        'btnDeleteSubjectAllClassess
        '
        Me.btnDeleteSubjectAllClassess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteSubjectAllClassess.Location = New System.Drawing.Point(617, 170)
        Me.btnDeleteSubjectAllClassess.Name = "btnDeleteSubjectAllClassess"
        Me.btnDeleteSubjectAllClassess.Size = New System.Drawing.Size(110, 163)
        Me.btnDeleteSubjectAllClassess.TabIndex = 25
        Me.btnDeleteSubjectAllClassess.Text = "Unassign selected subject for all classes"
        Me.btnDeleteSubjectAllClassess.UseVisualStyleBackColor = True
        '
        'btnDeleteAllSubjectForClass
        '
        Me.btnDeleteAllSubjectForClass.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteAllSubjectForClass.Location = New System.Drawing.Point(617, 339)
        Me.btnDeleteAllSubjectForClass.Name = "btnDeleteAllSubjectForClass"
        Me.btnDeleteAllSubjectForClass.Size = New System.Drawing.Size(110, 173)
        Me.btnDeleteAllSubjectForClass.TabIndex = 25
        Me.btnDeleteAllSubjectForClass.Text = "Unassign all subjects for selected class"
        Me.btnDeleteAllSubjectForClass.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(617, 587)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(110, 55)
        Me.btnRefresh.TabIndex = 3
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(19, 41)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(546, 604)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(82, 47)
        Me.btnUpdate.TabIndex = 0
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        Me.btnUpdate.Visible = False
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(298, 170)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(133, 43)
        Me.btnClear.TabIndex = 18
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnDeleteTeacherRecord)
        Me.GroupBox1.Controls.Add(Me.btnUpdateTeacher)
        Me.GroupBox1.Controls.Add(Me.cmbName)
        Me.GroupBox1.Controls.Add(Me.cmbSubjects)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btnClear)
        Me.GroupBox1.Controls.Add(Me.cmbClass)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.btnAdd)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 434)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(622, 220)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Assign Subject to Teacher"
        '
        'cmbName
        '
        Me.cmbName.FormattingEnabled = True
        Me.cmbName.Location = New System.Drawing.Point(157, 35)
        Me.cmbName.Name = "cmbName"
        Me.cmbName.Size = New System.Drawing.Size(275, 32)
        Me.cmbName.TabIndex = 22
        '
        'cmbSubjects
        '
        Me.cmbSubjects.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbSubjects.FormattingEnabled = True
        Me.cmbSubjects.Location = New System.Drawing.Point(156, 124)
        Me.cmbSubjects.Name = "cmbSubjects"
        Me.cmbSubjects.Size = New System.Drawing.Size(275, 32)
        Me.cmbSubjects.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(58, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 24)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Class"
        '
        'cmbClass
        '
        Me.cmbClass.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmbClass.FormattingEnabled = True
        Me.cmbClass.Location = New System.Drawing.Point(156, 80)
        Me.cmbClass.Name = "cmbClass"
        Me.cmbClass.Size = New System.Drawing.Size(275, 32)
        Me.cmbClass.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(58, 127)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 24)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Subjects"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(159, 170)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(133, 43)
        Me.btnAdd.TabIndex = 0
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(56, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Name"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(84, 40)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(524, 29)
        Me.txtName.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 24)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Name"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmbType)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.btnTruncateTeachers)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Controls.Add(Me.txtName)
        Me.GroupBox3.Controls.Add(Me.gbAvailableTeacherTime)
        Me.GroupBox3.Controls.Add(Me.btnAddTeacher)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(622, 404)
        Me.GroupBox3.TabIndex = 23
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "New Teacher Details"
        '
        'cmbType
        '
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Items.AddRange(New Object() {"Full Time", "Part Time", "Furthering"})
        Me.cmbType.Location = New System.Drawing.Point(86, 75)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(183, 32)
        Me.cmbType.TabIndex = 30
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 78)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 24)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Type"
        '
        'btnTruncateTeachers
        '
        Me.btnTruncateTeachers.Location = New System.Drawing.Point(315, 366)
        Me.btnTruncateTeachers.Name = "btnTruncateTeachers"
        Me.btnTruncateTeachers.Size = New System.Drawing.Size(276, 32)
        Me.btnTruncateTeachers.TabIndex = 25
        Me.btnTruncateTeachers.Text = "Truncate Teacher List"
        Me.btnTruncateTeachers.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(176, 366)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(133, 32)
        Me.Button1.TabIndex = 24
        Me.Button1.Text = "Clear"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnAddTeacher
        '
        Me.btnAddTeacher.Location = New System.Drawing.Point(21, 369)
        Me.btnAddTeacher.Name = "btnAddTeacher"
        Me.btnAddTeacher.Size = New System.Drawing.Size(133, 29)
        Me.btnAddTeacher.TabIndex = 23
        Me.btnAddTeacher.Text = "Add"
        Me.btnAddTeacher.UseVisualStyleBackColor = True
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
        Me.gbAvailableTeacherTime.Visible = False
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
        Me.cmbDay.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmbDay.FormattingEnabled = True
        Me.cmbDay.Items.AddRange(New Object() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"})
        Me.cmbDay.Location = New System.Drawing.Point(412, 143)
        Me.cmbDay.Name = "cmbDay"
        Me.cmbDay.Size = New System.Drawing.Size(89, 32)
        Me.cmbDay.TabIndex = 36
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(359, 146)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(42, 24)
        Me.Label9.TabIndex = 39
        Me.Label9.Text = "Day"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(412, 99)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(89, 29)
        Me.txtTo.TabIndex = 32
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(368, 102)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(33, 24)
        Me.Label7.TabIndex = 38
        Me.Label7.Text = "To"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(346, 57)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 24)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "From"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(412, 54)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(89, 29)
        Me.txtFrom.TabIndex = 28
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(359, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(134, 24)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Available Time"
        '
        'btnUpdateTeacher
        '
        Me.btnUpdateTeacher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateTeacher.Location = New System.Drawing.Point(438, 48)
        Me.btnUpdateTeacher.Name = "btnUpdateTeacher"
        Me.btnUpdateTeacher.Size = New System.Drawing.Size(153, 32)
        Me.btnUpdateTeacher.TabIndex = 23
        Me.btnUpdateTeacher.Text = "Update Record"
        Me.btnUpdateTeacher.UseVisualStyleBackColor = True
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
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Label11.Location = New System.Drawing.Point(507, 94)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(95, 40)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Should not be greater than end time"
        '
        'btnDeleteTeacherRecord
        '
        Me.btnDeleteTeacherRecord.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteTeacherRecord.Location = New System.Drawing.Point(438, 15)
        Me.btnDeleteTeacherRecord.Name = "btnDeleteTeacherRecord"
        Me.btnDeleteTeacherRecord.Size = New System.Drawing.Size(153, 32)
        Me.btnDeleteTeacherRecord.TabIndex = 24
        Me.btnDeleteTeacherRecord.Text = "Delete Record"
        Me.btnDeleteTeacherRecord.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(617, 13)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 58)
        Me.btnDelete.TabIndex = 27
        Me.btnDelete.Text = "Delete Teacher"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'frmManageTeachers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1370, 666)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.Name = "frmManageTeachers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "S"
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvTeachersList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gbAvailableTeacherTime.ResumeLayout(False)
        Me.gbAvailableTeacherTime.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnClear As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label8 As Label
    Friend WithEvents btnAdd As Button
    Friend WithEvents txtName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents dgvTeachersList As DataGridView
    Friend WithEvents cmbSubjects As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbClass As ComboBox
    Friend WithEvents Button3 As Button
    Friend WithEvents btnUpdate As Button
    Friend WithEvents cmbName As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Button1 As Button
    Friend WithEvents btnAddTeacher As Button
    Friend WithEvents btnRefresh As Button
    Friend WithEvents btnDeleteAllSubjectForClass As Button
    Friend WithEvents btnDeleteSubjectAllClassess As Button
    Friend WithEvents btnTruncate As Button
    Friend WithEvents btnTruncateTeachers As Button
    Friend WithEvents btnDeleteSelection As Button
    Friend WithEvents cmbType As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents gbAvailableTeacherTime As GroupBox
    Friend WithEvents btnAddTime As Button
    Friend WithEvents cmbDay As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtTo As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents txtFrom As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents ListView1 As ListView
    Friend WithEvents btnDeleteTime As Button
    Friend WithEvents btnUpdateTeacher As Button
    Friend WithEvents Label11 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnDeleteTeacherRecord As Button
End Class
