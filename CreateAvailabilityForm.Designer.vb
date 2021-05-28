<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreateAvailabilityForm
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
        Me.CreateAvailabilityBtn = New System.Windows.Forms.Button()
        Me.EndTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.StartTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.StartTimeLabel = New System.Windows.Forms.Label()
        Me.EndTimeLabel = New System.Windows.Forms.Label()
        Me.OperatorComboBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ScheduledCheckBox = New System.Windows.Forms.CheckBox()
        Me.RepeatCheckBox = New System.Windows.Forms.CheckBox()
        Me.UntilDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.UntilLabel = New System.Windows.Forms.Label()
        Me.MondayCheckBox = New System.Windows.Forms.CheckBox()
        Me.DayCheckListBox = New System.Windows.Forms.CheckedListBox()
        Me.StartDatePicker = New System.Windows.Forms.DateTimePicker()
        Me.EndDatePicker = New System.Windows.Forms.DateTimePicker()
        Me.WeekUpDown = New System.Windows.Forms.NumericUpDown()
        Me.RepeatLabelOne = New System.Windows.Forms.Label()
        Me.RepeatLabelTwo = New System.Windows.Forms.Label()
        CType(Me.WeekUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CreateAvailabilityBtn
        '
        Me.CreateAvailabilityBtn.Location = New System.Drawing.Point(160, 264)
        Me.CreateAvailabilityBtn.Name = "CreateAvailabilityBtn"
        Me.CreateAvailabilityBtn.Size = New System.Drawing.Size(112, 24)
        Me.CreateAvailabilityBtn.TabIndex = 0
        Me.CreateAvailabilityBtn.Text = "Create Availability"
        Me.CreateAvailabilityBtn.UseVisualStyleBackColor = True
        '
        'EndTimePicker
        '
        Me.EndTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.EndTimePicker.Location = New System.Drawing.Point(304, 48)
        Me.EndTimePicker.Name = "EndTimePicker"
        Me.EndTimePicker.ShowUpDown = True
        Me.EndTimePicker.Size = New System.Drawing.Size(96, 23)
        Me.EndTimePicker.TabIndex = 1
        Me.EndTimePicker.Value = New Date(2021, 3, 26, 12, 0, 0, 0)
        '
        'StartTimePicker
        '
        Me.StartTimePicker.CustomFormat = ""
        Me.StartTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.StartTimePicker.Location = New System.Drawing.Point(304, 16)
        Me.StartTimePicker.Name = "StartTimePicker"
        Me.StartTimePicker.ShowUpDown = True
        Me.StartTimePicker.Size = New System.Drawing.Size(96, 23)
        Me.StartTimePicker.TabIndex = 3
        Me.StartTimePicker.Value = New Date(2021, 3, 26, 0, 0, 0, 0)
        '
        'StartTimeLabel
        '
        Me.StartTimeLabel.AutoSize = True
        Me.StartTimeLabel.Location = New System.Drawing.Point(24, 16)
        Me.StartTimeLabel.Name = "StartTimeLabel"
        Me.StartTimeLabel.Size = New System.Drawing.Size(61, 15)
        Me.StartTimeLabel.TabIndex = 5
        Me.StartTimeLabel.Text = "Start Time"
        '
        'EndTimeLabel
        '
        Me.EndTimeLabel.AutoSize = True
        Me.EndTimeLabel.Location = New System.Drawing.Point(24, 48)
        Me.EndTimeLabel.Name = "EndTimeLabel"
        Me.EndTimeLabel.Size = New System.Drawing.Size(57, 15)
        Me.EndTimeLabel.TabIndex = 6
        Me.EndTimeLabel.Text = "End Time"
        '
        'OperatorComboBox
        '
        Me.OperatorComboBox.FormattingEnabled = True
        Me.OperatorComboBox.Location = New System.Drawing.Point(224, 136)
        Me.OperatorComboBox.Name = "OperatorComboBox"
        Me.OperatorComboBox.Size = New System.Drawing.Size(208, 23)
        Me.OperatorComboBox.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(168, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 15)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Operator"
        '
        'ScheduledCheckBox
        '
        Me.ScheduledCheckBox.AutoSize = True
        Me.ScheduledCheckBox.Location = New System.Drawing.Point(168, 200)
        Me.ScheduledCheckBox.Name = "ScheduledCheckBox"
        Me.ScheduledCheckBox.Size = New System.Drawing.Size(81, 19)
        Me.ScheduledCheckBox.TabIndex = 9
        Me.ScheduledCheckBox.Text = "Scheduled"
        Me.ScheduledCheckBox.UseVisualStyleBackColor = True
        '
        'RepeatCheckBox
        '
        Me.RepeatCheckBox.AutoSize = True
        Me.RepeatCheckBox.Location = New System.Drawing.Point(56, 80)
        Me.RepeatCheckBox.Name = "RepeatCheckBox"
        Me.RepeatCheckBox.Size = New System.Drawing.Size(67, 19)
        Me.RepeatCheckBox.TabIndex = 11
        Me.RepeatCheckBox.Text = "Repeat?"
        Me.RepeatCheckBox.UseVisualStyleBackColor = True
        '
        'UntilDateTimePicker
        '
        Me.UntilDateTimePicker.Enabled = False
        Me.UntilDateTimePicker.Location = New System.Drawing.Point(224, 104)
        Me.UntilDateTimePicker.Name = "UntilDateTimePicker"
        Me.UntilDateTimePicker.Size = New System.Drawing.Size(208, 23)
        Me.UntilDateTimePicker.TabIndex = 12
        '
        'UntilLabel
        '
        Me.UntilLabel.AutoSize = True
        Me.UntilLabel.Location = New System.Drawing.Point(168, 104)
        Me.UntilLabel.Name = "UntilLabel"
        Me.UntilLabel.Size = New System.Drawing.Size(32, 15)
        Me.UntilLabel.TabIndex = 13
        Me.UntilLabel.Text = "Until"
        '
        'MondayCheckBox
        '
        Me.MondayCheckBox.AutoSize = True
        Me.MondayCheckBox.Location = New System.Drawing.Point(-56, 128)
        Me.MondayCheckBox.Name = "MondayCheckBox"
        Me.MondayCheckBox.Size = New System.Drawing.Size(37, 19)
        Me.MondayCheckBox.TabIndex = 14
        Me.MondayCheckBox.Text = "M"
        Me.MondayCheckBox.UseVisualStyleBackColor = True
        '
        'DayCheckListBox
        '
        Me.DayCheckListBox.CheckOnClick = True
        Me.DayCheckListBox.Enabled = False
        Me.DayCheckListBox.FormattingEnabled = True
        Me.DayCheckListBox.Items.AddRange(New Object() {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"})
        Me.DayCheckListBox.Location = New System.Drawing.Point(48, 104)
        Me.DayCheckListBox.Name = "DayCheckListBox"
        Me.DayCheckListBox.Size = New System.Drawing.Size(104, 130)
        Me.DayCheckListBox.TabIndex = 18
        '
        'StartDatePicker
        '
        Me.StartDatePicker.Location = New System.Drawing.Point(96, 16)
        Me.StartDatePicker.Name = "StartDatePicker"
        Me.StartDatePicker.Size = New System.Drawing.Size(200, 23)
        Me.StartDatePicker.TabIndex = 19
        '
        'EndDatePicker
        '
        Me.EndDatePicker.Location = New System.Drawing.Point(96, 48)
        Me.EndDatePicker.Name = "EndDatePicker"
        Me.EndDatePicker.Size = New System.Drawing.Size(200, 23)
        Me.EndDatePicker.TabIndex = 20
        '
        'WeekUpDown
        '
        Me.WeekUpDown.Location = New System.Drawing.Point(248, 168)
        Me.WeekUpDown.Name = "WeekUpDown"
        Me.WeekUpDown.Size = New System.Drawing.Size(32, 23)
        Me.WeekUpDown.TabIndex = 21
        Me.WeekUpDown.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'RepeatLabelOne
        '
        Me.RepeatLabelOne.AutoSize = True
        Me.RepeatLabelOne.Location = New System.Drawing.Point(168, 168)
        Me.RepeatLabelOne.Name = "RepeatLabelOne"
        Me.RepeatLabelOne.Size = New System.Drawing.Size(74, 15)
        Me.RepeatLabelOne.TabIndex = 22
        Me.RepeatLabelOne.Text = "Repeat Every"
        '
        'RepeatLabelTwo
        '
        Me.RepeatLabelTwo.AutoSize = True
        Me.RepeatLabelTwo.Location = New System.Drawing.Point(288, 168)
        Me.RepeatLabelTwo.Name = "RepeatLabelTwo"
        Me.RepeatLabelTwo.Size = New System.Drawing.Size(36, 15)
        Me.RepeatLabelTwo.TabIndex = 23
        Me.RepeatLabelTwo.Text = "Week"
        '
        'CreateAvailabilityForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(451, 304)
        Me.Controls.Add(Me.RepeatLabelTwo)
        Me.Controls.Add(Me.RepeatLabelOne)
        Me.Controls.Add(Me.WeekUpDown)
        Me.Controls.Add(Me.EndDatePicker)
        Me.Controls.Add(Me.StartDatePicker)
        Me.Controls.Add(Me.DayCheckListBox)
        Me.Controls.Add(Me.MondayCheckBox)
        Me.Controls.Add(Me.UntilLabel)
        Me.Controls.Add(Me.UntilDateTimePicker)
        Me.Controls.Add(Me.RepeatCheckBox)
        Me.Controls.Add(Me.ScheduledCheckBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.OperatorComboBox)
        Me.Controls.Add(Me.EndTimeLabel)
        Me.Controls.Add(Me.StartTimeLabel)
        Me.Controls.Add(Me.StartTimePicker)
        Me.Controls.Add(Me.EndTimePicker)
        Me.Controls.Add(Me.CreateAvailabilityBtn)
        Me.Name = "CreateAvailabilityForm"
        Me.Text = "Create Availability"
        CType(Me.WeekUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CreateAvailabilityBtn As Button
    Friend WithEvents EndTimePicker As DateTimePicker
    Friend WithEvents StartTimePicker As DateTimePicker
    Friend WithEvents StartTimeLabel As Label
    Friend WithEvents EndTimeLabel As Label
    Friend WithEvents OperatorComboBox As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ScheduledCheckBox As CheckBox
    Friend WithEvents RepeatCheckBox As CheckBox
    Friend WithEvents UntilDateTimePicker As DateTimePicker
    Friend WithEvents UntilLabel As Label
    Friend WithEvents MondayCheckBox As CheckBox
    Friend WithEvents DayCheckListBox As CheckedListBox
    Friend WithEvents StartDatePicker1 As DateTimePicker
    Friend WithEvents EndDatePicker As DateTimePicker
    Friend WithEvents StartDatePicker As DateTimePicker
    Friend WithEvents WeekUpDown As NumericUpDown
    Friend WithEvents RepeatLabelOne As Label
    Friend WithEvents RepeatLabelTwo As Label
End Class
