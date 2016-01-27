<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckedListBox_Clients = New System.Windows.Forms.CheckedListBox()
        Me.cmbEnv = New System.Windows.Forms.ComboBox()
        Me.DateTimePickerFrom = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePickerTo = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar_Results = New System.Windows.Forms.ProgressBar()
        Me.lblAllDone = New System.Windows.Forms.Label()
        Me.lblClientName = New System.Windows.Forms.Label()
        Me.lblMemberStatus = New System.Windows.Forms.Label()
        Me.lblGroupStatus = New System.Windows.Forms.Label()
        Me.lblClientStatus = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 132)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Client(s):"
        '
        'CheckedListBox_Clients
        '
        Me.CheckedListBox_Clients.CheckOnClick = True
        Me.CheckedListBox_Clients.FormattingEnabled = True
        Me.CheckedListBox_Clients.Location = New System.Drawing.Point(95, 132)
        Me.CheckedListBox_Clients.Name = "CheckedListBox_Clients"
        Me.CheckedListBox_Clients.Size = New System.Drawing.Size(200, 199)
        Me.CheckedListBox_Clients.TabIndex = 0
        '
        'cmbEnv
        '
        Me.cmbEnv.FormattingEnabled = True
        Me.cmbEnv.Items.AddRange(New Object() {"DEV01", "DEV02", "PROD01", "PROD03"})
        Me.cmbEnv.Location = New System.Drawing.Point(95, 34)
        Me.cmbEnv.Name = "cmbEnv"
        Me.cmbEnv.Size = New System.Drawing.Size(121, 21)
        Me.cmbEnv.TabIndex = 2
        '
        'DateTimePickerFrom
        '
        Me.DateTimePickerFrom.Location = New System.Drawing.Point(95, 61)
        Me.DateTimePickerFrom.Name = "DateTimePickerFrom"
        Me.DateTimePickerFrom.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePickerFrom.TabIndex = 3
        Me.DateTimePickerFrom.Value = New Date(2015, 1, 1, 0, 0, 0, 0)
        '
        'DateTimePickerTo
        '
        Me.DateTimePickerTo.Location = New System.Drawing.Point(95, 87)
        Me.DateTimePickerTo.Name = "DateTimePickerTo"
        Me.DateTimePickerTo.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePickerTo.TabIndex = 4
        Me.DateTimePickerTo.Value = New Date(2015, 1, 1, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "From:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(23, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "To:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Environment:"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 20)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(543, 429)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox4)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(535, 403)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Start"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.cbAll)
        Me.GroupBox4.Controls.Add(Me.DateTimePickerTo)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.btnStart)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.DateTimePickerFrom)
        Me.GroupBox4.Controls.Add(Me.CheckedListBox_Clients)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.cmbEnv)
        Me.GroupBox4.Location = New System.Drawing.Point(84, 18)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(362, 379)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Run Jobs"
        '
        'cbAll
        '
        Me.cbAll.AutoSize = True
        Me.cbAll.Location = New System.Drawing.Point(74, 132)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(15, 14)
        Me.cbAll.TabIndex = 2
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(159, 350)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 1
        Me.btnStart.Text = "Run"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(535, 403)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Results"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ProgressBar_Results)
        Me.GroupBox2.Controls.Add(Me.lblAllDone)
        Me.GroupBox2.Controls.Add(Me.lblClientName)
        Me.GroupBox2.Controls.Add(Me.lblMemberStatus)
        Me.GroupBox2.Controls.Add(Me.lblGroupStatus)
        Me.GroupBox2.Controls.Add(Me.lblClientStatus)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 17)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(510, 106)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Status"
        '
        'ProgressBar_Results
        '
        Me.ProgressBar_Results.Location = New System.Drawing.Point(398, 48)
        Me.ProgressBar_Results.Name = "ProgressBar_Results"
        Me.ProgressBar_Results.Size = New System.Drawing.Size(100, 16)
        Me.ProgressBar_Results.TabIndex = 12
        '
        'lblAllDone
        '
        Me.lblAllDone.AutoSize = True
        Me.lblAllDone.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAllDone.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblAllDone.Location = New System.Drawing.Point(393, 72)
        Me.lblAllDone.Name = "lblAllDone"
        Me.lblAllDone.Size = New System.Drawing.Size(105, 25)
        Me.lblAllDone.TabIndex = 11
        Me.lblAllDone.Text = "All Done!"
        Me.lblAllDone.Visible = False
        '
        'lblClientName
        '
        Me.lblClientName.AutoSize = True
        Me.lblClientName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientName.Location = New System.Drawing.Point(83, 27)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.Size = New System.Drawing.Size(84, 16)
        Me.lblClientName.TabIndex = 10
        Me.lblClientName.Text = "[client name]"
        '
        'lblMemberStatus
        '
        Me.lblMemberStatus.AutoSize = True
        Me.lblMemberStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMemberStatus.Location = New System.Drawing.Point(109, 72)
        Me.lblMemberStatus.Name = "lblMemberStatus"
        Me.lblMemberStatus.Size = New System.Drawing.Size(104, 16)
        Me.lblMemberStatus.TabIndex = 7
        Me.lblMemberStatus.Text = "[member status]"
        '
        'lblGroupStatus
        '
        Me.lblGroupStatus.AutoSize = True
        Me.lblGroupStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroupStatus.Location = New System.Drawing.Point(109, 53)
        Me.lblGroupStatus.Name = "lblGroupStatus"
        Me.lblGroupStatus.Size = New System.Drawing.Size(89, 16)
        Me.lblGroupStatus.TabIndex = 6
        Me.lblGroupStatus.Text = "[group status]"
        '
        'lblClientStatus
        '
        Me.lblClientStatus.AutoSize = True
        Me.lblClientStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientStatus.Location = New System.Drawing.Point(413, 16)
        Me.lblClientStatus.Name = "lblClientStatus"
        Me.lblClientStatus.Size = New System.Drawing.Size(85, 16)
        Me.lblClientStatus.TabIndex = 5
        Me.lblClientStatus.Text = "[client status]"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Modern No. 20", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(37, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(66, 18)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Member"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Modern No. 20", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(38, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 18)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Group"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Modern No. 20", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(15, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 18)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Client"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtLog)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 129)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(516, 259)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Log"
        '
        'txtLog
        '
        Me.txtLog.Location = New System.Drawing.Point(6, 19)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(498, 234)
        Me.txtLog.TabIndex = 1
        '
        'BackgroundWorker1
        '
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(566, 461)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckedListBox_Clients As System.Windows.Forms.CheckedListBox
    Friend WithEvents cmbEnv As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents lblClientName As System.Windows.Forms.Label
    Friend WithEvents lblMemberStatus As System.Windows.Forms.Label
    Friend WithEvents lblGroupStatus As System.Windows.Forms.Label
    Friend WithEvents lblClientStatus As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents lblAllDone As System.Windows.Forms.Label
    Friend WithEvents ProgressBar_Results As System.Windows.Forms.ProgressBar

End Class
