<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenerateTimeTable
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenerateTimeTable))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.dgvTimetable = New System.Windows.Forms.DataGridView()
        Me.dgvExtraSubjects = New System.Windows.Forms.DataGridView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnManuallyAdjust = New System.Windows.Forms.Button()
        Me.btnReduce = New System.Windows.Forms.Button()
        CType(Me.dgvTimetable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvExtraSubjects, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(281, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Timetable"
        Me.Label1.Visible = False
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnGenerate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenerate.Location = New System.Drawing.Point(12, 574)
        Me.btnGenerate.Margin = New System.Windows.Forms.Padding(2)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(256, 39)
        Me.btnGenerate.TabIndex = 3
        Me.btnGenerate.Text = "Click here to &Generate Timetable"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(293, 574)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(183, 39)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save Timetable"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(502, 574)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(2)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(184, 39)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "Print Timetable"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'PrintDocument1
        '
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'dgvTimetable
        '
        Me.dgvTimetable.AllowUserToAddRows = False
        Me.dgvTimetable.AllowUserToDeleteRows = False
        Me.dgvTimetable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvTimetable.BackgroundColor = System.Drawing.Color.White
        Me.dgvTimetable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTimetable.Location = New System.Drawing.Point(12, 42)
        Me.dgvTimetable.Name = "dgvTimetable"
        Me.dgvTimetable.ReadOnly = True
        Me.dgvTimetable.Size = New System.Drawing.Size(674, 524)
        Me.dgvTimetable.TabIndex = 6
        '
        'dgvExtraSubjects
        '
        Me.dgvExtraSubjects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvExtraSubjects.BackgroundColor = System.Drawing.Color.White
        Me.dgvExtraSubjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExtraSubjects.Location = New System.Drawing.Point(6, 40)
        Me.dgvExtraSubjects.Name = "dgvExtraSubjects"
        Me.dgvExtraSubjects.Size = New System.Drawing.Size(361, 372)
        Me.dgvExtraSubjects.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(74, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(189, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "The following subjects were not added"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnReduce)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.btnManuallyAdjust)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dgvExtraSubjects)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(692, 35)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(379, 578)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(64, 451)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(157, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Manually adjusting the timetable"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(64, 471)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(269, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Reducing the number of times subjects should be taken"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(33, 429)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Try on of these options:"
        '
        'btnManuallyAdjust
        '
        Me.btnManuallyAdjust.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnManuallyAdjust.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManuallyAdjust.Location = New System.Drawing.Point(61, 507)
        Me.btnManuallyAdjust.Margin = New System.Windows.Forms.Padding(2)
        Me.btnManuallyAdjust.Name = "btnManuallyAdjust"
        Me.btnManuallyAdjust.Size = New System.Drawing.Size(272, 24)
        Me.btnManuallyAdjust.TabIndex = 10
        Me.btnManuallyAdjust.Text = "Manually Adjust Timetable"
        Me.btnManuallyAdjust.UseVisualStyleBackColor = True
        '
        'btnReduce
        '
        Me.btnReduce.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnReduce.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReduce.Location = New System.Drawing.Point(61, 539)
        Me.btnReduce.Margin = New System.Windows.Forms.Padding(2)
        Me.btnReduce.Name = "btnReduce"
        Me.btnReduce.Size = New System.Drawing.Size(272, 24)
        Me.btnReduce.TabIndex = 11
        Me.btnReduce.Text = "Reduce Number of Tiimes per Week"
        Me.btnReduce.UseVisualStyleBackColor = True
        '
        'frmGenerateTimeTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1079, 621)
        Me.Controls.Add(Me.dgvTimetable)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmGenerateTimeTable"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generate Timetable"
        CType(Me.dgvTimetable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvExtraSubjects, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents btnGenerate As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnPrint As Button
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As PrintPreviewDialog
    Friend WithEvents dgvTimetable As DataGridView
    Friend WithEvents dgvExtraSubjects As DataGridView
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents btnReduce As Button
    Friend WithEvents btnManuallyAdjust As Button
End Class
