Imports System.Windows.Forms.Calendar

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class OperatorMainForm
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.OperatorDataGridView = New System.Windows.Forms.DataGridView()
        Me.SaveBtn = New System.Windows.Forms.Button()
        Me.OperatorTab = New System.Windows.Forms.TabControl()
        Me.OperatorTabPage = New System.Windows.Forms.TabPage()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ScheduleTab = New System.Windows.Forms.TabPage()
        Me.RefreshButton = New System.Windows.Forms.Button()
        Me.SaveAvalabilityChangesBtn = New System.Windows.Forms.Button()
        Me.ShowMissingBtn = New System.Windows.Forms.Button()
        Me.IsScheduledCheckBox = New System.Windows.Forms.CheckBox()
        Me.ShowAllCheckBox = New System.Windows.Forms.CheckBox()
        Me.AddAvailabilityBtn = New System.Windows.Forms.Button()
        Me.AvailabilityDataGridView = New System.Windows.Forms.DataGridView()
        Me.AvailabilityMonthCalendar = New System.Windows.Forms.MonthCalendar()
        Me.ConsoleTab = New System.Windows.Forms.TabPage()
        Me.MonthYearPicker = New System.Windows.Forms.DateTimePicker()
        Me.FilterButton = New System.Windows.Forms.Button()
        Me.ActionComboBox = New System.Windows.Forms.ComboBox()
        Me.ConsoleRichTextBox = New System.Windows.Forms.RichTextBox()
        Me.ActionButton = New System.Windows.Forms.Button()
        CType(Me.OperatorDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OperatorTab.SuspendLayout()
        Me.OperatorTabPage.SuspendLayout()
        Me.ScheduleTab.SuspendLayout()
        CType(Me.AvailabilityDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ConsoleTab.SuspendLayout()
        Me.SuspendLayout()
        '
        'OperatorDataGridView
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.OperatorDataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.OperatorDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.OperatorDataGridView.DefaultCellStyle = DataGridViewCellStyle2
        Me.OperatorDataGridView.Location = New System.Drawing.Point(32, 48)
        Me.OperatorDataGridView.Name = "OperatorDataGridView"
        Me.OperatorDataGridView.RowTemplate.Height = 25
        Me.OperatorDataGridView.Size = New System.Drawing.Size(520, 296)
        Me.OperatorDataGridView.TabIndex = 0
        '
        'SaveBtn
        '
        Me.SaveBtn.Location = New System.Drawing.Point(200, 352)
        Me.SaveBtn.Name = "SaveBtn"
        Me.SaveBtn.Size = New System.Drawing.Size(176, 23)
        Me.SaveBtn.TabIndex = 10
        Me.SaveBtn.Text = "Save Changes"
        Me.SaveBtn.UseVisualStyleBackColor = True
        '
        'OperatorTab
        '
        Me.OperatorTab.Controls.Add(Me.OperatorTabPage)
        Me.OperatorTab.Controls.Add(Me.ScheduleTab)
        Me.OperatorTab.Controls.Add(Me.ConsoleTab)
        Me.OperatorTab.Location = New System.Drawing.Point(0, 0)
        Me.OperatorTab.Name = "OperatorTab"
        Me.OperatorTab.SelectedIndex = 0
        Me.OperatorTab.Size = New System.Drawing.Size(800, 448)
        Me.OperatorTab.TabIndex = 11
        '
        'OperatorTabPage
        '
        Me.OperatorTabPage.Controls.Add(Me.Button1)
        Me.OperatorTabPage.Controls.Add(Me.SaveBtn)
        Me.OperatorTabPage.Controls.Add(Me.OperatorDataGridView)
        Me.OperatorTabPage.Location = New System.Drawing.Point(4, 24)
        Me.OperatorTabPage.Name = "OperatorTabPage"
        Me.OperatorTabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.OperatorTabPage.Size = New System.Drawing.Size(792, 420)
        Me.OperatorTabPage.TabIndex = 0
        Me.OperatorTabPage.Text = "Operators"
        Me.OperatorTabPage.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(640, 216)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ScheduleTab
        '
        Me.ScheduleTab.Controls.Add(Me.RefreshButton)
        Me.ScheduleTab.Controls.Add(Me.SaveAvalabilityChangesBtn)
        Me.ScheduleTab.Controls.Add(Me.ShowMissingBtn)
        Me.ScheduleTab.Controls.Add(Me.IsScheduledCheckBox)
        Me.ScheduleTab.Controls.Add(Me.ShowAllCheckBox)
        Me.ScheduleTab.Controls.Add(Me.AddAvailabilityBtn)
        Me.ScheduleTab.Controls.Add(Me.AvailabilityDataGridView)
        Me.ScheduleTab.Controls.Add(Me.AvailabilityMonthCalendar)
        Me.ScheduleTab.Location = New System.Drawing.Point(4, 24)
        Me.ScheduleTab.Name = "ScheduleTab"
        Me.ScheduleTab.Padding = New System.Windows.Forms.Padding(3)
        Me.ScheduleTab.Size = New System.Drawing.Size(792, 420)
        Me.ScheduleTab.TabIndex = 1
        Me.ScheduleTab.Text = "Schedule"
        Me.ScheduleTab.UseVisualStyleBackColor = True
        '
        'RefreshButton
        '
        Me.RefreshButton.Location = New System.Drawing.Point(704, 384)
        Me.RefreshButton.Name = "RefreshButton"
        Me.RefreshButton.Size = New System.Drawing.Size(75, 23)
        Me.RefreshButton.TabIndex = 7
        Me.RefreshButton.Text = "Refresh"
        Me.RefreshButton.UseVisualStyleBackColor = True
        '
        'SaveAvalabilityChangesBtn
        '
        Me.SaveAvalabilityChangesBtn.Location = New System.Drawing.Point(464, 384)
        Me.SaveAvalabilityChangesBtn.Name = "SaveAvalabilityChangesBtn"
        Me.SaveAvalabilityChangesBtn.Size = New System.Drawing.Size(112, 23)
        Me.SaveAvalabilityChangesBtn.TabIndex = 6
        Me.SaveAvalabilityChangesBtn.Text = "Save Changes"
        Me.SaveAvalabilityChangesBtn.UseVisualStyleBackColor = True
        '
        'ShowMissingBtn
        '
        Me.ShowMissingBtn.Location = New System.Drawing.Point(72, 280)
        Me.ShowMissingBtn.Name = "ShowMissingBtn"
        Me.ShowMissingBtn.Size = New System.Drawing.Size(112, 23)
        Me.ShowMissingBtn.TabIndex = 5
        Me.ShowMissingBtn.Text = "Show Missing"
        Me.ShowMissingBtn.UseVisualStyleBackColor = True
        '
        'IsScheduledCheckBox
        '
        Me.IsScheduledCheckBox.AutoSize = True
        Me.IsScheduledCheckBox.Location = New System.Drawing.Point(88, 232)
        Me.IsScheduledCheckBox.Name = "IsScheduledCheckBox"
        Me.IsScheduledCheckBox.Size = New System.Drawing.Size(113, 19)
        Me.IsScheduledCheckBox.TabIndex = 4
        Me.IsScheduledCheckBox.Text = "Show Scheduled"
        Me.IsScheduledCheckBox.UseVisualStyleBackColor = True
        '
        'ShowAllCheckBox
        '
        Me.ShowAllCheckBox.AutoSize = True
        Me.ShowAllCheckBox.Location = New System.Drawing.Point(88, 200)
        Me.ShowAllCheckBox.Name = "ShowAllCheckBox"
        Me.ShowAllCheckBox.Size = New System.Drawing.Size(78, 19)
        Me.ShowAllCheckBox.TabIndex = 3
        Me.ShowAllCheckBox.Text = "Show ALL"
        Me.ShowAllCheckBox.UseVisualStyleBackColor = True
        '
        'AddAvailabilityBtn
        '
        Me.AddAvailabilityBtn.Location = New System.Drawing.Point(88, 352)
        Me.AddAvailabilityBtn.Name = "AddAvailabilityBtn"
        Me.AddAvailabilityBtn.Size = New System.Drawing.Size(75, 23)
        Me.AddAvailabilityBtn.TabIndex = 2
        Me.AddAvailabilityBtn.Text = "Add"
        Me.AddAvailabilityBtn.UseVisualStyleBackColor = True
        '
        'AvailabilityDataGridView
        '
        Me.AvailabilityDataGridView.AllowUserToAddRows = False
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.AvailabilityDataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.AvailabilityDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.AvailabilityDataGridView.DefaultCellStyle = DataGridViewCellStyle4
        Me.AvailabilityDataGridView.Location = New System.Drawing.Point(256, 16)
        Me.AvailabilityDataGridView.Name = "AvailabilityDataGridView"
        Me.AvailabilityDataGridView.RowTemplate.Height = 25
        Me.AvailabilityDataGridView.Size = New System.Drawing.Size(520, 360)
        Me.AvailabilityDataGridView.TabIndex = 1
        '
        'AvailabilityMonthCalendar
        '
        Me.AvailabilityMonthCalendar.Location = New System.Drawing.Point(16, 16)
        Me.AvailabilityMonthCalendar.MaxSelectionCount = 31
        Me.AvailabilityMonthCalendar.Name = "AvailabilityMonthCalendar"
        Me.AvailabilityMonthCalendar.ShowTodayCircle = False
        Me.AvailabilityMonthCalendar.TabIndex = 0
        '
        'ConsoleTab
        '
        Me.ConsoleTab.Controls.Add(Me.MonthYearPicker)
        Me.ConsoleTab.Controls.Add(Me.FilterButton)
        Me.ConsoleTab.Controls.Add(Me.ActionComboBox)
        Me.ConsoleTab.Controls.Add(Me.ConsoleRichTextBox)
        Me.ConsoleTab.Controls.Add(Me.ActionButton)
        Me.ConsoleTab.Location = New System.Drawing.Point(4, 24)
        Me.ConsoleTab.Name = "ConsoleTab"
        Me.ConsoleTab.Padding = New System.Windows.Forms.Padding(3)
        Me.ConsoleTab.Size = New System.Drawing.Size(792, 420)
        Me.ConsoleTab.TabIndex = 2
        Me.ConsoleTab.Text = "Console"
        Me.ConsoleTab.UseVisualStyleBackColor = True
        '
        'MonthYearPicker
        '
        Me.MonthYearPicker.Location = New System.Drawing.Point(256, 16)
        Me.MonthYearPicker.Name = "MonthYearPicker"
        Me.MonthYearPicker.Size = New System.Drawing.Size(192, 23)
        Me.MonthYearPicker.TabIndex = 7
        '
        'FilterButton
        '
        Me.FilterButton.Location = New System.Drawing.Point(456, 16)
        Me.FilterButton.Name = "FilterButton"
        Me.FilterButton.Size = New System.Drawing.Size(75, 23)
        Me.FilterButton.TabIndex = 6
        Me.FilterButton.Text = "Filter"
        Me.FilterButton.UseVisualStyleBackColor = True
        '
        'ActionComboBox
        '
        Me.ActionComboBox.FormattingEnabled = True
        Me.ActionComboBox.Items.AddRange(New Object() {"Refresh email", "Send email to inner operators", "Send email to ALL operators", "Help"})
        Me.ActionComboBox.Location = New System.Drawing.Point(248, 384)
        Me.ActionComboBox.Name = "ActionComboBox"
        Me.ActionComboBox.Size = New System.Drawing.Size(200, 23)
        Me.ActionComboBox.TabIndex = 5
        '
        'ConsoleRichTextBox
        '
        Me.ConsoleRichTextBox.Location = New System.Drawing.Point(16, 48)
        Me.ConsoleRichTextBox.Name = "ConsoleRichTextBox"
        Me.ConsoleRichTextBox.ReadOnly = True
        Me.ConsoleRichTextBox.Size = New System.Drawing.Size(768, 328)
        Me.ConsoleRichTextBox.TabIndex = 4
        Me.ConsoleRichTextBox.Text = ""
        '
        'ActionButton
        '
        Me.ActionButton.Location = New System.Drawing.Point(456, 384)
        Me.ActionButton.Name = "ActionButton"
        Me.ActionButton.Size = New System.Drawing.Size(64, 23)
        Me.ActionButton.TabIndex = 3
        Me.ActionButton.Text = "Start"
        Me.ActionButton.UseVisualStyleBackColor = True
        '
        'OperatorMainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.OperatorTab)
        Me.Name = "OperatorMainForm"
        Me.Text = "Operators"
        CType(Me.OperatorDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OperatorTab.ResumeLayout(False)
        Me.OperatorTabPage.ResumeLayout(False)
        Me.ScheduleTab.ResumeLayout(False)
        Me.ScheduleTab.PerformLayout()
        CType(Me.AvailabilityDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ConsoleTab.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents OperatorDataGridView As DataGridView
    Friend WithEvents SaveBtn As Button
    Friend WithEvents OperatorTab As TabControl
    Friend WithEvents OperatorTabPage As TabPage
    Friend WithEvents ScheduleTab As TabPage
    Friend WithEvents AvailabilityMonthCalendar As MonthCalendar
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents AvailabilityDataGridView As DataGridView
    Friend WithEvents AddAvailabilityBtn As Button
    Friend WithEvents ShowAllCheckBox As CheckBox
    Friend WithEvents IsScheduledCheckBox As CheckBox
    Friend WithEvents SaveAvalabilityChangesBtn As Button
    Friend WithEvents ShowMissingBtn As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents ConsoleTab As TabPage
    Friend WithEvents ActionComboBox As ComboBox
    Friend WithEvents ConsoleRichTextBox As RichTextBox
    Friend WithEvents ActionButton As Button
    Friend WithEvents FilterButton As Button
    Friend WithEvents MonthYearPicker As DateTimePicker
    Friend WithEvents RefreshButton As Button
End Class
