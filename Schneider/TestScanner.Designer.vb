<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestScanner
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
        Me.components = New System.ComponentModel.Container()
        Me.txtScan = New System.Windows.Forms.TextBox()
        Me.lblLabel = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.lblDetik = New System.Windows.Forms.Label()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.txtData = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtData1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtScan
        '
        Me.txtScan.Location = New System.Drawing.Point(73, 69)
        Me.txtScan.Name = "txtScan"
        Me.txtScan.Size = New System.Drawing.Size(161, 20)
        Me.txtScan.TabIndex = 0
        '
        'lblLabel
        '
        Me.lblLabel.AutoSize = True
        Me.lblLabel.Location = New System.Drawing.Point(70, 136)
        Me.lblLabel.Name = "lblLabel"
        Me.lblLabel.Size = New System.Drawing.Size(39, 13)
        Me.lblLabel.TabIndex = 1
        Me.lblLabel.Text = "Label1"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'lblDetik
        '
        Me.lblDetik.AutoSize = True
        Me.lblDetik.Location = New System.Drawing.Point(70, 207)
        Me.lblDetik.Name = "lblDetik"
        Me.lblDetik.Size = New System.Drawing.Size(13, 13)
        Me.lblDetik.TabIndex = 2
        Me.lblDetik.Text = "0"
        '
        'txtData
        '
        Me.txtData.Location = New System.Drawing.Point(73, 152)
        Me.txtData.Name = "txtData"
        Me.txtData.Size = New System.Drawing.Size(161, 20)
        Me.txtData.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(73, 263)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(93, 32)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Read"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "%MW11000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(181, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "BENCH 1"
        '
        'txtData1
        '
        Me.txtData1.Location = New System.Drawing.Point(73, 178)
        Me.txtData1.Name = "txtData1"
        Me.txtData1.Size = New System.Drawing.Size(161, 20)
        Me.txtData1.TabIndex = 7
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(172, 263)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(93, 32)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Clear DB"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TestScanner
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 323)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtData1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtData)
        Me.Controls.Add(Me.lblDetik)
        Me.Controls.Add(Me.lblLabel)
        Me.Controls.Add(Me.txtScan)
        Me.Name = "TestScanner"
        Me.Text = "TestScanner"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtScan As TextBox
    Friend WithEvents lblLabel As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents lblDetik As Label
    Friend WithEvents SerialPort1 As IO.Ports.SerialPort
    Friend WithEvents txtData As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtData1 As TextBox
    Friend WithEvents Button2 As Button
End Class
