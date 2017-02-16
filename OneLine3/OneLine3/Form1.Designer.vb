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
        Me.OFD1 = New System.Windows.Forms.OpenFileDialog()
        Me.btGetFileName = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataFileName = New System.Windows.Forms.TextBox()
        Me.btClose = New System.Windows.Forms.Button()
        Me.lsbxExcel = New System.Windows.Forms.ListBox()
        Me.lbListName1 = New System.Windows.Forms.Label()
        Me.btIn = New System.Windows.Forms.Button()
        Me.lsbxAutocad = New System.Windows.Forms.ListBox()
        Me.btOut = New System.Windows.Forms.Button()
        Me.btGenerate = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbLabel = New System.Windows.Forms.CheckBox()
        Me.cbRatedActivePower = New System.Windows.Forms.CheckBox()
        Me.cbPowerFactor = New System.Windows.Forms.CheckBox()
        Me.cbRatedCurrent = New System.Windows.Forms.CheckBox()
        Me.cbEstimatedLength = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cbCableVoltageDeviation = New System.Windows.Forms.CheckBox()
        Me.cbCable = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'OFD1
        '
        '
        'btGetFileName
        '
        Me.btGetFileName.Location = New System.Drawing.Point(559, 16)
        Me.btGetFileName.Name = "btGetFileName"
        Me.btGetFileName.Size = New System.Drawing.Size(75, 23)
        Me.btGetFileName.TabIndex = 0
        Me.btGetFileName.Text = "Открыть"
        Me.btGetFileName.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataFileName)
        Me.GroupBox1.Controls.Add(Me.btGetFileName)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(645, 50)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Путь к файлу с расчетами"
        '
        'DataFileName
        '
        Me.DataFileName.Location = New System.Drawing.Point(6, 19)
        Me.DataFileName.Name = "DataFileName"
        Me.DataFileName.Size = New System.Drawing.Size(547, 20)
        Me.DataFileName.TabIndex = 1
        '
        'btClose
        '
        Me.btClose.Location = New System.Drawing.Point(258, 391)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(75, 25)
        Me.btClose.TabIndex = 3
        Me.btClose.Text = "Закрыть"
        Me.btClose.UseVisualStyleBackColor = True
        '
        'lsbxExcel
        '
        Me.lsbxExcel.FormattingEnabled = True
        Me.lsbxExcel.Location = New System.Drawing.Point(33, 108)
        Me.lsbxExcel.Name = "lsbxExcel"
        Me.lsbxExcel.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lsbxExcel.Size = New System.Drawing.Size(129, 264)
        Me.lsbxExcel.TabIndex = 4
        '
        'lbListName1
        '
        Me.lbListName1.AutoSize = True
        Me.lbListName1.Location = New System.Drawing.Point(30, 82)
        Me.lbListName1.Name = "lbListName1"
        Me.lbListName1.Size = New System.Drawing.Size(79, 13)
        Me.lbListName1.TabIndex = 5
        Me.lbListName1.Text = "Список щитов"
        '
        'btIn
        '
        Me.btIn.Location = New System.Drawing.Point(168, 210)
        Me.btIn.Name = "btIn"
        Me.btIn.Size = New System.Drawing.Size(30, 29)
        Me.btIn.TabIndex = 6
        Me.btIn.Text = ">"
        Me.btIn.UseVisualStyleBackColor = True
        '
        'lsbxAutocad
        '
        Me.lsbxAutocad.FormattingEnabled = True
        Me.lsbxAutocad.Location = New System.Drawing.Point(204, 108)
        Me.lsbxAutocad.Name = "lsbxAutocad"
        Me.lsbxAutocad.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lsbxAutocad.Size = New System.Drawing.Size(129, 264)
        Me.lsbxAutocad.TabIndex = 7
        '
        'btOut
        '
        Me.btOut.Location = New System.Drawing.Point(168, 245)
        Me.btOut.Name = "btOut"
        Me.btOut.Size = New System.Drawing.Size(30, 29)
        Me.btOut.TabIndex = 8
        Me.btOut.Text = "<"
        Me.btOut.UseVisualStyleBackColor = True
        '
        'btGenerate
        '
        Me.btGenerate.Location = New System.Drawing.Point(177, 393)
        Me.btGenerate.Name = "btGenerate"
        Me.btGenerate.Size = New System.Drawing.Size(75, 23)
        Me.btGenerate.TabIndex = 9
        Me.btGenerate.Text = "Начертить"
        Me.btGenerate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(599, 511)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "NabiyevTR@gmail.com"
        '
        'cbLabel
        '
        Me.cbLabel.AutoSize = True
        Me.cbLabel.Checked = True
        Me.cbLabel.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbLabel.Location = New System.Drawing.Point(384, 144)
        Me.cbLabel.Name = "cbLabel"
        Me.cbLabel.Size = New System.Drawing.Size(99, 17)
        Me.cbLabel.TabIndex = 11
        Me.cbLabel.Text = "Номер группы"
        Me.cbLabel.UseVisualStyleBackColor = True
        '
        'cbRatedActivePower
        '
        Me.cbRatedActivePower.AutoSize = True
        Me.cbRatedActivePower.Checked = True
        Me.cbRatedActivePower.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbRatedActivePower.Location = New System.Drawing.Point(384, 167)
        Me.cbRatedActivePower.Name = "cbRatedActivePower"
        Me.cbRatedActivePower.Size = New System.Drawing.Size(128, 17)
        Me.cbRatedActivePower.TabIndex = 12
        Me.cbRatedActivePower.Text = "Расчетная нагрузка"
        Me.cbRatedActivePower.UseVisualStyleBackColor = True
        '
        'cbPowerFactor
        '
        Me.cbPowerFactor.AutoSize = True
        Me.cbPowerFactor.Checked = True
        Me.cbPowerFactor.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbPowerFactor.Location = New System.Drawing.Point(384, 190)
        Me.cbPowerFactor.Name = "cbPowerFactor"
        Me.cbPowerFactor.Size = New System.Drawing.Size(145, 17)
        Me.cbPowerFactor.TabIndex = 13
        Me.cbPowerFactor.Text = "Коэффицент мощности"
        Me.cbPowerFactor.UseVisualStyleBackColor = True
        '
        'cbRatedCurrent
        '
        Me.cbRatedCurrent.AutoSize = True
        Me.cbRatedCurrent.Checked = True
        Me.cbRatedCurrent.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbRatedCurrent.Location = New System.Drawing.Point(33, 104)
        Me.cbRatedCurrent.Name = "cbRatedCurrent"
        Me.cbRatedCurrent.Size = New System.Drawing.Size(101, 17)
        Me.cbRatedCurrent.TabIndex = 14
        Me.cbRatedCurrent.Text = "Расчетный ток"
        Me.cbRatedCurrent.UseVisualStyleBackColor = True
        '
        'cbEstimatedLength
        '
        Me.cbEstimatedLength.AutoSize = True
        Me.cbEstimatedLength.Checked = True
        Me.cbEstimatedLength.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbEstimatedLength.Location = New System.Drawing.Point(384, 237)
        Me.cbEstimatedLength.Name = "cbEstimatedLength"
        Me.cbEstimatedLength.Size = New System.Drawing.Size(101, 17)
        Me.cbEstimatedLength.TabIndex = 15
        Me.cbEstimatedLength.Text = "Длина участка"
        Me.cbEstimatedLength.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbCable)
        Me.GroupBox2.Controls.Add(Me.cbCableVoltageDeviation)
        Me.GroupBox2.Controls.Add(Me.cbRatedCurrent)
        Me.GroupBox2.Location = New System.Drawing.Point(351, 110)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(356, 164)
        Me.GroupBox2.TabIndex = 16
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Параметры  кабельной линии"
        '
        'cbCableVoltageDeviation
        '
        Me.cbCableVoltageDeviation.AutoSize = True
        Me.cbCableVoltageDeviation.Checked = True
        Me.cbCableVoltageDeviation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbCableVoltageDeviation.Location = New System.Drawing.Point(184, 34)
        Me.cbCableVoltageDeviation.Name = "cbCableVoltageDeviation"
        Me.cbCableVoltageDeviation.Size = New System.Drawing.Size(128, 17)
        Me.cbCableVoltageDeviation.TabIndex = 17
        Me.cbCableVoltageDeviation.Text = "Потеря напряжения"
        Me.cbCableVoltageDeviation.UseVisualStyleBackColor = True
        '
        'cbCable
        '
        Me.cbCable.AutoSize = True
        Me.cbCable.Checked = True
        Me.cbCable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbCable.Location = New System.Drawing.Point(184, 57)
        Me.cbCable.Name = "cbCable"
        Me.cbCable.Size = New System.Drawing.Size(151, 17)
        Me.cbCable.TabIndex = 18
        Me.cbCable.Text = "Марка и сечение кабеля"
        Me.cbCable.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(730, 533)
        Me.Controls.Add(Me.cbEstimatedLength)
        Me.Controls.Add(Me.cbPowerFactor)
        Me.Controls.Add(Me.cbRatedActivePower)
        Me.Controls.Add(Me.cbLabel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btGenerate)
        Me.Controls.Add(Me.btOut)
        Me.Controls.Add(Me.lsbxAutocad)
        Me.Controls.Add(Me.btIn)
        Me.Controls.Add(Me.lbListName1)
        Me.Controls.Add(Me.lsbxExcel)
        Me.Controls.Add(Me.btClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "Form1"
        Me.Text = "Компоновка однолинейной схемы"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OFD1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btGetFileName As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DataFileName As System.Windows.Forms.TextBox
    Friend WithEvents btClose As System.Windows.Forms.Button
    Friend WithEvents lsbxExcel As System.Windows.Forms.ListBox
    Friend WithEvents lbListName1 As System.Windows.Forms.Label
    Friend WithEvents btIn As System.Windows.Forms.Button
    Friend WithEvents lsbxAutocad As System.Windows.Forms.ListBox
    Friend WithEvents btOut As System.Windows.Forms.Button
    Friend WithEvents btGenerate As System.Windows.Forms.Button

    Private Sub btGetFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGetFileName.Click

    End Sub

    Private Sub btGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGenerate.Click

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbLabel As System.Windows.Forms.CheckBox
    Friend WithEvents cbRatedActivePower As System.Windows.Forms.CheckBox
    Friend WithEvents cbPowerFactor As System.Windows.Forms.CheckBox
    Friend WithEvents cbRatedCurrent As System.Windows.Forms.CheckBox
    Friend WithEvents cbEstimatedLength As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cbCable As System.Windows.Forms.CheckBox
    Friend WithEvents cbCableVoltageDeviation As System.Windows.Forms.CheckBox
End Class
