<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMisc
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMisc))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnBTruncate = New System.Windows.Forms.Button()
        Me.btnBDelete = New System.Windows.Forms.Button()
        Me.btnBUpdate = New System.Windows.Forms.Button()
        Me.dgvBreaks = New System.Windows.Forms.DataGridView()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtLEndTime = New System.Windows.Forms.TextBox()
        Me.txtLStartTime = New System.Windows.Forms.TextBox()
        Me.dtpLEndTime = New System.Windows.Forms.DateTimePicker()
        Me.dtpLStartTime = New System.Windows.Forms.DateTimePicker()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.listDays = New System.Windows.Forms.CheckedListBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtBEndTIme = New System.Windows.Forms.TextBox()
        Me.txtBStartTime = New System.Windows.Forms.TextBox()
        Me.dtpBEndTme = New System.Windows.Forms.DateTimePicker()
        Me.dtpBStartTime = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtBName = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnClassDetails = New System.Windows.Forms.Button()
        Me.btnTermDetails = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvBreaks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.dgvBreaks)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(573, 26)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(569, 523)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Breaks"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Panel1.Controls.Add(Me.btnRefresh)
        Me.Panel1.Controls.Add(Me.btnBTruncate)
        Me.Panel1.Controls.Add(Me.btnBDelete)
        Me.Panel1.Controls.Add(Me.btnBUpdate)
        Me.Panel1.Location = New System.Drawing.Point(41, 438)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(508, 70)
        Me.Panel1.TabIndex = 0
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(259, 13)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(91, 46)
        Me.btnRefresh.TabIndex = 4
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnBTruncate
        '
        Me.btnBTruncate.Location = New System.Drawing.Point(391, 13)
        Me.btnBTruncate.Name = "btnBTruncate"
        Me.btnBTruncate.Size = New System.Drawing.Size(100, 46)
        Me.btnBTruncate.TabIndex = 3
        Me.btnBTruncate.Text = "Truncate"
        Me.btnBTruncate.UseVisualStyleBackColor = True
        '
        'btnBDelete
        '
        Me.btnBDelete.Location = New System.Drawing.Point(146, 13)
        Me.btnBDelete.Name = "btnBDelete"
        Me.btnBDelete.Size = New System.Drawing.Size(91, 46)
        Me.btnBDelete.TabIndex = 1
        Me.btnBDelete.Text = "Delete"
        Me.btnBDelete.UseVisualStyleBackColor = True
        '
        'btnBUpdate
        '
        Me.btnBUpdate.Location = New System.Drawing.Point(26, 13)
        Me.btnBUpdate.Name = "btnBUpdate"
        Me.btnBUpdate.Size = New System.Drawing.Size(91, 46)
        Me.btnBUpdate.TabIndex = 0
        Me.btnBUpdate.Text = "Update"
        Me.btnBUpdate.UseVisualStyleBackColor = True
        '
        'dgvBreaks
        '
        Me.dgvBreaks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvBreaks.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvBreaks.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvBreaks.BackgroundColor = System.Drawing.Color.White
        Me.dgvBreaks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBreaks.Location = New System.Drawing.Point(31, 36)
        Me.dgvBreaks.Name = "dgvBreaks"
        Me.dgvBreaks.RowHeadersVisible = False
        Me.dgvBreaks.Size = New System.Drawing.Size(508, 387)
        Me.dgvBreaks.TabIndex = 6
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Ivory
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 26)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(523, 672)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Input Details"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtLEndTime)
        Me.GroupBox4.Controls.Add(Me.txtLStartTime)
        Me.GroupBox4.Controls.Add(Me.dtpLEndTime)
        Me.GroupBox4.Controls.Add(Me.dtpLStartTime)
        Me.GroupBox4.Controls.Add(Me.btnUpdate)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Location = New System.Drawing.Point(28, 36)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(451, 165)
        Me.GroupBox4.TabIndex = 25
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Start and End Time Details"
        '
        'txtLEndTime
        '
        Me.txtLEndTime.Location = New System.Drawing.Point(251, 74)
        Me.txtLEndTime.Name = "txtLEndTime"
        Me.txtLEndTime.Size = New System.Drawing.Size(79, 29)
        Me.txtLEndTime.TabIndex = 23
        Me.txtLEndTime.Text = "14:00"
        '
        'txtLStartTime
        '
        Me.txtLStartTime.Location = New System.Drawing.Point(251, 32)
        Me.txtLStartTime.Name = "txtLStartTime"
        Me.txtLStartTime.Size = New System.Drawing.Size(79, 29)
        Me.txtLStartTime.TabIndex = 24
        Me.txtLStartTime.Text = "08:30"
        '
        'dtpLEndTime
        '
        Me.dtpLEndTime.CustomFormat = "HH:mm"
        Me.dtpLEndTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpLEndTime.Location = New System.Drawing.Point(357, 73)
        Me.dtpLEndTime.Name = "dtpLEndTime"
        Me.dtpLEndTime.Size = New System.Drawing.Size(75, 29)
        Me.dtpLEndTime.TabIndex = 7
        Me.dtpLEndTime.Value = New Date(2019, 11, 19, 14, 0, 0, 0)
        Me.dtpLEndTime.Visible = False
        '
        'dtpLStartTime
        '
        Me.dtpLStartTime.CustomFormat = "HH:mm"
        Me.dtpLStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpLStartTime.Location = New System.Drawing.Point(357, 30)
        Me.dtpLStartTime.Name = "dtpLStartTime"
        Me.dtpLStartTime.Size = New System.Drawing.Size(75, 29)
        Me.dtpLStartTime.TabIndex = 6
        Me.dtpLStartTime.Value = New Date(2019, 11, 19, 8, 30, 0, 0)
        Me.dtpLStartTime.Visible = False
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(329, 108)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(103, 36)
        Me.btnUpdate.TabIndex = 0
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 77)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(159, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Lesson End Time"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(51, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 24)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Lesson Start Time"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.listDays)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.txtBEndTIme)
        Me.GroupBox3.Controls.Add(Me.txtBStartTime)
        Me.GroupBox3.Controls.Add(Me.dtpBEndTme)
        Me.GroupBox3.Controls.Add(Me.dtpBStartTime)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.btnAdd)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.txtBName)
        Me.GroupBox3.Location = New System.Drawing.Point(28, 224)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(381, 423)
        Me.GroupBox3.TabIndex = 8
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Break Details"
        '
        'listDays
        '
        Me.listDays.CheckOnClick = True
        Me.listDays.FormattingEnabled = True
        Me.listDays.Items.AddRange(New Object() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"})
        Me.listDays.Location = New System.Drawing.Point(158, 208)
        Me.listDays.Name = "listDays"
        Me.listDays.Size = New System.Drawing.Size(178, 124)
        Me.listDays.TabIndex = 25
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(42, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 24)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "Days"
        '
        'txtBEndTIme
        '
        Me.txtBEndTIme.Location = New System.Drawing.Point(158, 148)
        Me.txtBEndTIme.Name = "txtBEndTIme"
        Me.txtBEndTIme.Size = New System.Drawing.Size(79, 29)
        Me.txtBEndTIme.TabIndex = 22
        '
        'txtBStartTime
        '
        Me.txtBStartTime.Location = New System.Drawing.Point(158, 95)
        Me.txtBStartTime.Name = "txtBStartTime"
        Me.txtBStartTime.Size = New System.Drawing.Size(79, 29)
        Me.txtBStartTime.TabIndex = 21
        '
        'dtpBEndTme
        '
        Me.dtpBEndTme.CustomFormat = "HH:mm"
        Me.dtpBEndTme.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBEndTme.Location = New System.Drawing.Point(261, 148)
        Me.dtpBEndTme.Name = "dtpBEndTme"
        Me.dtpBEndTme.Size = New System.Drawing.Size(75, 29)
        Me.dtpBEndTme.TabIndex = 20
        Me.dtpBEndTme.Visible = False
        '
        'dtpBStartTime
        '
        Me.dtpBStartTime.CustomFormat = "HH:mm"
        Me.dtpBStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBStartTime.Location = New System.Drawing.Point(261, 95)
        Me.dtpBStartTime.Name = "dtpBStartTime"
        Me.dtpBStartTime.Size = New System.Drawing.Size(75, 29)
        Me.dtpBStartTime.TabIndex = 19
        Me.dtpBStartTime.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(42, 151)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 24)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "End Time"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(158, 354)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(178, 48)
        Me.btnAdd.TabIndex = 14
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(42, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 24)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(42, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 24)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Start Time"
        '
        'txtBName
        '
        Me.txtBName.Location = New System.Drawing.Point(158, 51)
        Me.txtBName.Name = "txtBName"
        Me.txtBName.Size = New System.Drawing.Size(178, 29)
        Me.txtBName.TabIndex = 17
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.btnTermDetails)
        Me.GroupBox5.Controls.Add(Me.btnClassDetails)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!)
        Me.GroupBox5.Location = New System.Drawing.Point(573, 563)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(567, 110)
        Me.GroupBox5.TabIndex = 1
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Other Details"
        '
        'btnClassDetails
        '
        Me.btnClassDetails.Location = New System.Drawing.Point(41, 27)
        Me.btnClassDetails.Name = "btnClassDetails"
        Me.btnClassDetails.Size = New System.Drawing.Size(223, 65)
        Me.btnClassDetails.TabIndex = 1
        Me.btnClassDetails.Text = "View/Edit Class Details"
        Me.btnClassDetails.UseVisualStyleBackColor = True
        '
        'btnTermDetails
        '
        Me.btnTermDetails.Location = New System.Drawing.Point(273, 27)
        Me.btnTermDetails.Name = "btnTermDetails"
        Me.btnTermDetails.Size = New System.Drawing.Size(274, 65)
        Me.btnTermDetails.TabIndex = 2
        Me.btnTermDetails.Text = "View/Edit Term and Session Details"
        Me.btnTermDetails.UseVisualStyleBackColor = True
        '
        'frmMisc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1154, 716)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMisc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manage Miscellaneous Details"
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvBreaks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents dgvBreaks As DataGridView
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnUpdate As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnBTruncate As Button
    Friend WithEvents btnBDelete As Button
    Friend WithEvents btnBUpdate As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents dtpBEndTme As DateTimePicker
    Friend WithEvents dtpBStartTime As DateTimePicker
    Friend WithEvents Label5 As Label
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtBName As TextBox
    Friend WithEvents txtLEndTime As TextBox
    Friend WithEvents txtLStartTime As TextBox
    Friend WithEvents txtBEndTIme As TextBox
    Friend WithEvents txtBStartTime As TextBox
    Friend WithEvents dtpLEndTime As DateTimePicker
    Friend WithEvents dtpLStartTime As DateTimePicker
    Friend WithEvents Label6 As Label
    Friend WithEvents listDays As CheckedListBox
    Friend WithEvents btnRefresh As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents btnClassDetails As Button
    Friend WithEvents btnTermDetails As Button
End Class
